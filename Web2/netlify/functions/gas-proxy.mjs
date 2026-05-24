const GAS_URL = "https://script.google.com/macros/s/AKfycbwp-ijTu2p3dg8K3E1hyctBlRdh9UV03EXEA_3N4zDZvkd6Ty5CKBLC3SJ3Cg5YMu9JJA/exec";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

export async function handler(event) {
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: CORS, body: "" };
  }

  try {
    const qs = event.rawQuery ? `?${event.rawQuery}` : "";
    const response = await fetch(GAS_URL + qs, {
      method: event.httpMethod || "GET",
      redirect: "follow",
    });
    const body = await response.text();
    return {
      statusCode: response.status,
      headers: {
        "Content-Type": response.headers.get("Content-Type") || "application/json",
        ...CORS,
      },
      body,
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: { "Content-Type": "application/json", ...CORS },
      body: JSON.stringify({ error: err.message }),
    };
  }
}
