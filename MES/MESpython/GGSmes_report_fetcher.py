"""
MES Report → Google Sheets (All Tabs)
ดึงข้อมูลทุก endpoint ของหน้า /#/report → push ขึ้น Google Sheets (แต่ละ endpoint เป็น 1 worksheet)

⚠️ ชุดทดสอบ — credentials ฝังในไฟล์
   ใช้งานจริงควรย้ายไป .env / env vars

Auth: Firebase signInWithPassword (ขอ token ใหม่ทุกครั้งที่รัน — ไม่ต้องใช้ refresh token)

Google Sheets Setup (ทำครั้งเดียว):
  1. สร้าง Service Account ใน Google Cloud → ดาวน์โหลด service_account.json
  2. เปิด Google Sheets API + Drive API
  3. สร้าง Google Sheet ชื่อ SHEET_NAME (ดูค่าใน CONFIG)
  4. แชร์ sheet ให้ email service account เป็น Editor
  5. วาง service_account.json ไว้โฟลเดอร์เดียวกับไฟล์นี้

Requirements:
  pip install requests pandas gspread oauth2client
"""

import json
import logging
import os
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import numpy as np
import pandas as pd
import requests

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# โหลด .env (pip install python-dotenv) — ถ้าไม่ติดตั้งก็ข้ามไป
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env")
except ImportError:
    pass


# ==========================================
# ⚙️ CONFIG
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

# 🔐 MES Credentials (โหลดจาก .env — ห้าม hardcode)
MES_EMAIL = os.getenv("MES_EMAIL", "")
MES_PASSWORD = os.getenv("MES_PASSWORD", "")
if not MES_EMAIL or not MES_PASSWORD:
    raise RuntimeError(
        "❌ MES_EMAIL/MES_PASSWORD ไม่ได้ตั้งค่า\n"
        "💡 สร้างไฟล์ .env ในโฟลเดอร์เดียวกัน (ดู .env.example เป็นตัวอย่าง)\n"
        "💡 และติดตั้ง: pip install python-dotenv"
    )

# Firebase API Key ของ MES (ค่าคงที่)
FIREBASE_API_KEY = "AIzaSyBXpxmaKmfjzNVtgO7cZtCXmPPwt6gv2xM"

# 🎯 Target
PROJECT_ID = 32
DATE_FROM = datetime(2026, 5, 1, 7, 0, 0)
DATE_TO = datetime(2026, 6, 31, 7, 0, 0)

# 🔗 MES Base URL
BASE_URL = "https://mes.systemtd.com"

# 📊 Google Sheets
# ใช้ SHEET_URL ก่อน (แน่ใจกว่า) — ถ้าเว้นว่างจะ fallback ไปเปิดด้วย SHEET_NAME
SHEET_URL = "https://docs.google.com/spreadsheets/d/1slFaR0CeXU8OjCBHglDrhOcJ3tdkJjqOycFJiLEFC7k/edit"
SHEET_NAME = "Energy_Data"  # แชร์ให้ tdmes-6@fuel-mes.iam.gserviceaccount.com แล้ว
SERVICE_ACCOUNT_JSON = str(Path(__file__).parent / "service_account.json")
GSHEET_BATCH_DELAY = 1.2  # หน่วงระหว่าง worksheet (วินาที) กัน rate limit (60 writes/min/project)
MES_DAY_DELAY = 0.25      # หน่วงระหว่างยิงรายวัน (วินาที) กัน rate limit ฝั่ง MES

# 💾 CSV backup (ออปชั่น)
SAVE_CSV_BACKUP = True
# path relative ของไฟล์เอง → ทำงานได้ทั้ง Windows local และ Linux runner (GitHub Actions)
OUTPUT_DIR = Path(__file__).parent / "output"


# ==========================================
# 🚨 Custom Exceptions
# ==========================================
class NetworkDownError(Exception):
    """ยิง MES ไม่ได้เลย — DNS/network ดับ → ควรหยุดทั้งสคริปต์"""
    pass


# ==========================================
# 🔐 Firebase Login
# ==========================================
def get_fresh_token(email: str, password: str) -> str:
    """Login Firebase → ID Token (อายุ ~1 ชม.) — ขอใหม่ทุกครั้งที่รัน"""
    url = (
        "https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword"
        f"?key={FIREBASE_API_KEY}"
    )
    r = requests.post(
        url,
        json={"email": email, "password": password, "returnSecureToken": True},
        timeout=30,
    )
    if r.status_code != 200:
        raise RuntimeError(f"Login failed [HTTP {r.status_code}]: {r.text}")
    data = r.json()
    if "idToken" not in data:
        raise RuntimeError(f"Login response ไม่มี idToken: {data}")
    logging.info(
        f"✅ Login: {data.get('email')} (token อายุ {data.get('expiresIn')}s)"
    )
    return data["idToken"]


# ==========================================
# 📦 MES Report Fetcher
# ==========================================
class MESReportFetcher:
    def __init__(self, id_token: str, project_id: int):
        self.token = id_token
        self.project_id = project_id
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Bearer {id_token}",
                "Accept": "application/json, text/plain, */*",
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                ),
                "Origin": BASE_URL,
                "Referer": f"{BASE_URL}/",
            }
        )

    # ----- Date encoders -----
    @staticmethod
    def encode_js_date(dt: datetime) -> str:
        """
        JS Date.toString() แบบ URL-encoded — hardcode weekday/month (กัน locale ไทยทำ %a/%b เพี้ยน)
        ตัวอย่าง: 'Fri May 01 2026 07:00:00 GMT+0700 (Indochina Time)'
        """
        days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        dow = days[dt.weekday()]
        mon = months[dt.month - 1]
        return quote(
            f"{dow} {mon} {dt.day:02d} {dt.year} "
            f"{dt.strftime('%H:%M:%S')} GMT+0700 (Indochina Time)"
        )

    @staticmethod
    def encode_iso_utc(dt: datetime, tz_offset_hours: int = 7) -> str:
        utc = dt - timedelta(hours=tz_offset_hours)
        return utc.strftime("%Y-%m-%dT%H:%M:%S.000Z")

    # ----- HTTP -----
    def get(self, path: str, params: Optional[Dict[str, Any]] = None) -> Optional[Any]:
        """
        GET พร้อม retry สำหรับ ConnectionError (DNS/network blip)
        ถ้า fail ครบทั้ง 3 ครั้ง → raise NetworkDownError เพื่อให้ main loop หยุดเลย
        ไม่ retry สำหรับ HTTP error (4xx/5xx) — return None ตามเดิม
        """
        url = f"{BASE_URL}{path}"
        max_retries = 3
        delays = [5, 10, 20]  # exponential backoff

        for attempt in range(max_retries):
            try:
                r = self.session.get(url, params=params, timeout=120)
                r.raise_for_status()
                return r.json()
            except requests.exceptions.ConnectionError as e:
                # DNS fail / connection refused / network down
                if attempt < max_retries - 1:
                    wait = delays[attempt]
                    logging.warning(
                        f"⚠️ Network error ({type(e).__name__}) — "
                        f"retry {attempt+1}/{max_retries} ใน {wait}s..."
                    )
                    time.sleep(wait)
                    continue
                # หมดสิทธิ์แล้ว → network ดับจริง → abort ทั้งสคริปต์
                raise NetworkDownError(
                    f"Network ใช้ไม่ได้ — fail {max_retries} ครั้งติดที่ {path}\n{e}"
                )
            except requests.exceptions.Timeout as e:
                if attempt < max_retries - 1:
                    logging.warning(f"⚠️ Timeout — retry {attempt+1}/{max_retries}...")
                    time.sleep(delays[attempt])
                    continue
                logging.error(f"❌ GET {path}: timeout หลัง retry {max_retries} ครั้ง")
                return None
            except requests.exceptions.RequestException as e:
                # HTTP 4xx/5xx — endpoint ผิด/permission → ไม่ retry
                logging.error(f"❌ GET {path}: {e}")
                return None
            except json.JSONDecodeError:
                logging.error(f"❌ GET {path}: JSON decode error")
                return None
        return None

    # ----- Endpoints -----
    def haulingProduction(self, s, e):       return self.get(f"/report/haulingProduction/{s}/{e}/{self.project_id}")
    def haulingProductionTimes(self, s, e):  return self.get(f"/report/haulingProductionTimes/{s}/{e}/{self.project_id}")
    def haulingProductionTrip(self, s, e):   return self.get(f"/report/haulingProductionTrip/{s}/{e}/{self.project_id}")
    def loadingProduction(self, s, e):       return self.get(f"/report/loadingProduction/{s}/{e}/{self.project_id}")
    def loadingProductionTimes(self, s, e):  return self.get(f"/report/loadingProductionTimes/{s}/{e}/{self.project_id}")
    def loadingTruckTrip(self, s, e):        return self.get(f"/report/loadingTruckTrip/{s}/{e}")
    def loadingExcavatorTrip(self, s, e):    return self.get(f"/report/loadingExcavatorTrip/{s}/{e}")
    def loadingsupportTrip(self, s, e):      return self.get(f"/report/loadingsupportTrip/{s}/{e}")
    def energyUsageHistory(self, s, e):      return self.get(f"/report/energyUsageHistory/{s}/{e}/{self.project_id}")
    def energyUsageProject(self, s, e):      return self.get(f"/report/energyUsageProject/{s}/{e}/{self.project_id}")
    def supportMachineActivity(self, s, e):  return self.get(f"/report/supportMachineActivity/{s}/{e}/{self.project_id}")
    def supportMachineMajor(self, s, e):     return self.get(f"/report/supportMachineMajor/{s}/{e}/{self.project_id}")
    def dailyTruckTrip(self, s, e):          return self.get(f"/report/dailyTruckTrip/{s}/{e}/{self.project_id}")

    # ----- Endpoints เสริม (signature ยังไม่ verify 100% — ลองยิงดู ถ้าได้ 0 แถวอาจต้องปรับ params) -----
    def listAvgMeter(self, s, e):                  return self.get(f"/report/listAvgMeter/{s}/{e}/{self.project_id}")
    def listcoalTransportByMachineId(self, s, e):  return self.get(f"/report/listcoalTransportByMachineId/{s}/{e}/{self.project_id}")
    def listActiveAdjustedStockReport(self, s, e): return self.get(f"/report/listActiveAdjustedStockReport/{s}/{e}/{self.project_id}")
    def listMaterialRequestReport(self, s, e):     return self.get(f"/report/listMaterialRequestReport/{s}/{e}/{self.project_id}")


# ==========================================
# 🛠️ Helpers
# ==========================================
def unwrap(data):
    """ดึง list ออกจาก response ที่อาจห่อใน {'data': [...]} ฯลฯ"""
    if isinstance(data, dict):
        for k in ("data", "result", "rows", "items"):
            if k in data and isinstance(data[k], list):
                return data[k]
        return [data]
    return data or []


def fetch_full_range(fn, date_from: datetime, date_to: datetime, label: str = "") -> list:
    """
    ยิง endpoint ด้วยทั้งช่วงวันที่ในครั้งเดียว (เลียนแบบหน้าเว็บกด LOAD DATA)
    encode_js_date() ที่สร้าง start/end ใหม่ทุกครั้ง = "select date" ก่อนดึงทุก endpoint
    """
    s = MESReportFetcher.encode_js_date(date_from)
    e = MESReportFetcher.encode_js_date(date_to)
    logging.info(
        f"     📅 {label}: select range {date_from:%Y-%m-%d %H:%M} ~ {date_to:%Y-%m-%d %H:%M}"
    )
    rows = unwrap(fn(s, e))
    logging.info(f"     ✔️ {label}: {len(rows)} แถว")
    return rows


def flatten_row(row: Any) -> Dict[str, Any]:
    """แปลง nested dict → flat dict (gspread ไม่รับ dict/list ใน cell)"""
    if not isinstance(row, dict):
        return {"value": str(row)}
    out = {}
    for k, v in row.items():
        if isinstance(v, (dict, list)):
            out[k] = json.dumps(v, ensure_ascii=False)
        else:
            out[k] = v
    return out


# ==========================================
# 📊 Google Sheets Writer
# ==========================================
def open_spreadsheet():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_JSON, scope)
    client = gspread.authorize(creds)

    if SHEET_URL:
        try:
            return client.open_by_url(SHEET_URL)
        except Exception as ex:
            raise RuntimeError(
                f"❌ เปิด SHEET_URL ไม่ได้: {ex}\n"
                f"💡 ตรวจว่า service account ถูกแชร์เป็น Editor ของ sheet นี้แล้ว"
            )

    try:
        return client.open(SHEET_NAME)
    except gspread.SpreadsheetNotFound:
        raise RuntimeError(
            f"❌ ไม่พบ Google Sheet '{SHEET_NAME}' — สร้าง sheet + แชร์ให้ "
            f"service account email เป็น Editor ก่อน"
        )


def push_to_worksheet(spreadsheet, ws_name: str, rows: List[Any]) -> int:
    """เคลียร์ worksheet เก่า → push ข้อมูลใหม่ ถ้าไม่มี ws จะสร้างใหม่"""
    if not rows:
        logging.warning(f"   ⚠️ {ws_name}: ไม่มีข้อมูล (ข้าม)")
        return 0

    flat = [flatten_row(r) for r in rows]
    df = pd.DataFrame(flat)

    # 🔧 ล้าง NaN/inf ก่อน stringify (gspread JSON encoder รับ NaN ไม่ได้)
    df = df.replace([np.inf, -np.inf], np.nan)
    df = df.where(pd.notnull(df), "")                      # NaN จริง → ""
    df = df.astype(str)
    df = df.replace({"nan": "", "None": "", "NaT": ""})    # เผื่อหลงเหลือเป็น string

    try:
        ws = spreadsheet.worksheet(ws_name)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(
            title=ws_name, rows=max(len(df) + 50, 1000), cols=max(len(df.columns) + 5, 26)
        )
        logging.info(f"   ✨ สร้าง worksheet '{ws_name}' ใหม่")

    payload = [df.columns.tolist()] + df.values.tolist()
    ws.update(payload, value_input_option="USER_ENTERED")
    logging.info(f"   💾 {ws_name}: {len(df)} แถว → Google Sheets ✅")
    return len(df)


def save_csv_backup(rows: List[Any], filepath: Path, label: str) -> int:
    if not rows:
        return 0
    flat = [flatten_row(r) for r in rows]
    df = pd.DataFrame(flat)
    filepath.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(filepath, index=False, encoding="utf-8-sig")
    logging.info(f"   📂 {label}: {len(df)} แถว → {filepath.name}")
    return len(df)


# ==========================================
# ⚡ ENERGY SUMMARY (liter/hour แบบ ก: Σliter ÷ Σ Δhour ต่อเครื่อง)
# ==========================================
def _to_num(series):
    """แปลงเป็นตัวเลข ล้าง comma/ค่าว่าง/#REF!"""
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0.0)


def _classify_machine(mid: str) -> str:
    """แยกประเภทจาก machineId prefix (ปรับ mapping ได้ตามจริง)"""
    m = str(mid).upper().strip()
    if m.startswith("E-"):
        return "รถขุด"
    if m.startswith("MN-"):
        return "รถบรรทุก"
    if m.startswith(("TN-", "PX-")):
        return "รถบริการ"
    if m.startswith(("B-", "HE-", "H-", "MW-", "WL-", "C-", "SL-", "LT-", "TS-")):
        return "เครื่องจักรสนับสนุน"
    return "อื่นๆ"


def build_energy_summary(rows_21: list) -> dict:
    """
    รับ rows ดิบจาก 21_EnergyUsageProject → คืน dict ของ DataFrame หลายมุม
    liter/hour (แบบ ก) = Σliter ÷ Σ Δhour ; Δhour = max(hourMeter) − min(hourMeter) ต่อเครื่อง
    """
    if not rows_21:
        return {}

    df = pd.DataFrame([flatten_row(r) for r in rows_21])
    for c in ("liter", "hourMeter", "mileage", "kwUnit", "price"):
        if c in df.columns:
            df[c] = _to_num(df[c])
    if "machineId" not in df.columns:
        df["machineId"] = ""
    if "costCenter" not in df.columns:
        df["costCenter"] = ""

    df["date"] = df.get("workingDate", "").astype(str).str.slice(0, 10)
    df["mtype"] = df["machineId"].map(_classify_machine)

    # ---------- (A) สรุปรายเครื่อง ----------
    def _per_machine(g):
        liters = g["liter"].sum()
        hmin = g["hourMeter"].replace(0, np.nan).min()
        hmax = g["hourMeter"].max()
        dhour = (hmax - hmin) if (pd.notna(hmin) and hmax > hmin) else 0.0
        return pd.Series({
            "type": g["mtype"].iloc[0],
            "costCenter": g["costCenter"].mode().iloc[0] if not g["costCenter"].mode().empty else "",
            "fillups": int(len(g)),
            "totalLiter": round(liters, 1),
            "hourStart": round(hmin, 1) if pd.notna(hmin) else 0,
            "hourEnd": round(hmax, 1) if pd.notna(hmax) else 0,
            "deltaHour": round(dhour, 1),
            "literPerHour": round(liters / dhour, 3) if dhour > 0 else 0.0,
        })

    by_machine = (
        df.groupby("machineId").apply(_per_machine).reset_index()
        .sort_values("totalLiter", ascending=False)
    )

    # ---------- (B) สรุปรายวัน ----------
    by_day = (
        df.groupby("date").agg(
            totalLiter=("liter", "sum"),
            fillups=("liter", "size"),
            machines=("machineId", "nunique"),
        ).reset_index().sort_values("date")
    )
    by_day["totalLiter"] = by_day["totalLiter"].round(1)

    # ---------- (C) สรุปรายศูนย์ต้นทุน ----------
    by_cc = (
        df.groupby("costCenter").agg(
            totalLiter=("liter", "sum"),
            fillups=("liter", "size"),
            machines=("machineId", "nunique"),
        ).reset_index().sort_values("totalLiter", ascending=False)
    )
    by_cc["totalLiter"] = by_cc["totalLiter"].round(1)

    # ---------- (D) สรุปรายประเภท ----------
    by_type = (
        df.groupby("mtype").agg(
            totalLiter=("liter", "sum"),
            machines=("machineId", "nunique"),
        ).reset_index().rename(columns={"mtype": "type"})
    )
    by_type["totalLiter"] = by_type["totalLiter"].round(1)

    # ---------- (E) KPI รวม ----------
    sum_delta_hour = float(by_machine["deltaHour"].sum())
    total_liter = float(df["liter"].sum())
    overall = pd.DataFrame([{
        "totalLiter": round(total_liter, 1),
        "totalFillups": int(len(df)),
        "activeMachines": int(df["machineId"].nunique()),
        "sumDeltaHour": round(sum_delta_hour, 1),
        "avgLiterPerHour": round(total_liter / sum_delta_hour, 3) if sum_delta_hour > 0 else 0.0,
        "avgLiterPerFillup": round(total_liter / len(df), 1) if len(df) else 0.0,
    }])

    return {
        "90_EnergyByMachine": by_machine,
        "91_EnergyByDay": by_day,
        "92_EnergyByCostCenter": by_cc,
        "93_EnergyByType": by_type,
        "94_EnergyKPI": overall,
    }


def push_df(spreadsheet, ws_name: str, df: pd.DataFrame) -> int:
    """push DataFrame ตรงๆ (logic เดียวกับ push_to_worksheet แต่รับ df แล้ว)"""
    if df is None or df.empty:
        logging.warning(f"   ⚠️ {ws_name}: ไม่มีข้อมูล (ข้าม)")
        return 0
    df = df.replace([np.inf, -np.inf], np.nan).where(pd.notnull(df), "").astype(str)
    df = df.replace({"nan": "", "None": "", "NaT": ""})
    try:
        ws = spreadsheet.worksheet(ws_name)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(
            title=ws_name, rows=max(len(df) + 50, 100), cols=max(len(df.columns) + 5, 12)
        )
        logging.info(f"   ✨ สร้าง worksheet '{ws_name}' ใหม่")
    ws.update([df.columns.tolist()] + df.values.tolist(), value_input_option="USER_ENTERED")
    logging.info(f"   💾 {ws_name}: {len(df)} แถว → Google Sheets ✅")
    return len(df)


# ==========================================
# 🚀 Main
# ==========================================
def main():
    logging.info("=" * 60)
    logging.info("🔋 MES Report → Google Sheets")
    logging.info(f"   Project: {PROJECT_ID} | {DATE_FROM:%Y-%m-%d %H:%M} → {DATE_TO:%Y-%m-%d %H:%M}")
    logging.info("=" * 60)

    # 1) Auth
    token = get_fresh_token(MES_EMAIL, MES_PASSWORD)
    fetcher = MESReportFetcher(token, PROJECT_ID)

    # 3) เปิด Spreadsheet
    target = SHEET_URL if SHEET_URL else f"name='{SHEET_NAME}'"
    logging.info(f"📊 เปิด Google Sheet ({target})...")
    spreadsheet = open_spreadsheet()
    logging.info(f"   ✅ '{spreadsheet.title}' → {spreadsheet.url}")

    # 4) เตรียม CSV backup folder
    if SAVE_CSV_BACKUP:
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        backup_dir = OUTPUT_DIR / f"gsheet_{stamp}"
        backup_dir.mkdir(parents=True, exist_ok=True)
        logging.info(f"   📁 CSV backup: {backup_dir}")

    # 5) ลิสต์ของทุก endpoint → (ws_name, csv_name, fetch_fn)
    tasks = [
        ("01_HaulingProduction",      "01_hauling_production.csv",       fetcher.haulingProduction),
        ("02_HaulingProductionTimes", "02_hauling_production_times.csv", fetcher.haulingProductionTimes),
        ("03_HaulingProductionTrip",  "03_hauling_production_trip.csv",  fetcher.haulingProductionTrip),
        ("10_LoadingProduction",      "10_loading_production.csv",       fetcher.loadingProduction),
        ("11_LoadingProductionTimes", "11_loading_production_times.csv", fetcher.loadingProductionTimes),
        ("12_LoadingTruckTrip",       "12_loading_truck_trip.csv",       fetcher.loadingTruckTrip),
        ("13_LoadingExcavatorTrip",   "13_loading_excavator_trip.csv",   fetcher.loadingExcavatorTrip),
        ("14_LoadingSupportTrip",     "14_loading_support_trip.csv",     fetcher.loadingsupportTrip),
        ("20_EnergyUsageHistory",     "20_energy_usage_history.csv",     fetcher.energyUsageHistory),
        ("21_EnergyUsageProject",     "21_energy_usage_project.csv",     fetcher.energyUsageProject),
        ("30_SupportActivity",        "30_support_activity.csv",         fetcher.supportMachineActivity),
        ("31_SupportMajor",           "31_support_major.csv",            fetcher.supportMachineMajor),
        ("40_DailyTruckTrip",         "40_daily_truck_trip.csv",         fetcher.dailyTruckTrip),
        # --- Endpoints เสริม (signature ยังไม่ verify 100%) ---
        ("50_AvgMeter",               "50_avg_meter.csv",                fetcher.listAvgMeter),
        ("51_CoalTransport",          "51_coal_transport.csv",           fetcher.listcoalTransportByMachineId),
        ("52_AdjustedStock",          "52_adjusted_stock.csv",           fetcher.listActiveAdjustedStockReport),
        ("53_MaterialRequest",        "53_material_request.csv",         fetcher.listMaterialRequestReport),
    ]

    # 6) Loop fetch + push (ยิงแบบรายวัน → รวมผล → push)
    summary = []
    network_down = False
    for ws_name, csv_name, fn in tasks:
        logging.info(f"▶ {ws_name}")
        try:
            data = fetch_full_range(fn, DATE_FROM, DATE_TO, ws_name)
            n = push_to_worksheet(spreadsheet, ws_name, data)
            if SAVE_CSV_BACKUP:
                save_csv_backup(data, backup_dir / csv_name, ws_name)
            summary.append((ws_name, n))
            time.sleep(GSHEET_BATCH_DELAY)
        except NetworkDownError as ex:
            logging.error(f"   🚨 {ws_name}: Network ดับ — หยุดทั้งสคริปต์: {ex}")
            summary.append((ws_name, -1))
            network_down = True
            break
        except Exception as ex:
            logging.error(f"   ❌ {ws_name}: {type(ex).__name__}: {ex!r}")
            summary.append((ws_name, -1))

    # 6.5) Energy Summary (90-94) — รวมยอด liter/hour ต่อเครื่อง/วัน/costCenter/type + KPI
    if network_down:
        logging.warning("⏭️ ข้าม Energy Summary เพราะ network ดับ")
    else:
        logging.info("▶ Energy Summary (90-94)")
        try:
            raw21 = fetch_full_range(fetcher.energyUsageProject, DATE_FROM, DATE_TO, "21_raw")
            summaries = build_energy_summary(raw21)
            for ws_name, df_sum in summaries.items():
                n = push_df(spreadsheet, ws_name, df_sum)
                if SAVE_CSV_BACKUP and not df_sum.empty:
                    df_sum.to_csv(backup_dir / f"{ws_name}.csv", index=False, encoding="utf-8-sig")
                summary.append((ws_name, n))
                time.sleep(GSHEET_BATCH_DELAY)
        except NetworkDownError as ex:
            logging.error(f"   🚨 EnergySummary: Network ดับ: {ex}")
        except Exception as ex:
            logging.error(f"   ❌ EnergySummary: {type(ex).__name__}: {ex!r}")

    # 7) สรุป
    logging.info("=" * 60)
    logging.info("📋 สรุป:")
    total = 0
    for name, n in summary:
        flag = "❌" if n < 0 else ("⚠️" if n == 0 else "✅")
        logging.info(f"   {flag} {name}: {n if n >= 0 else 'ERROR'} แถว")
        if n > 0:
            total += n
    logging.info(f"   📊 รวม: {total} แถว ใน {len([s for s in summary if s[1] > 0])} worksheet")
    logging.info(f"   🔗 {spreadsheet.url}")
    logging.info("=" * 60)
    logging.info("🎉 เสร็จสมบูรณ์")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.error(f"❌ {e}")
        raise
