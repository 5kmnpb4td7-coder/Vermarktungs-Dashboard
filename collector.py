#!/usr/bin/env python3
"""
Vermarktungs-Tracker: Mapraz Parc + Zentrum Worben
=====================================================

Laeuft in GitHub Actions (Zeitplan alle 2h) UND lokal testbar.

Was das Skript macht:
1. Ruft die aktuellen Vermietungsdaten von mapraz-parc.ch ab
   (Logik uebernommen aus dem bestehenden, bewaehrten mapraz_tracker.py).
2. Ruft die aktuellen Verkaufsdaten von der Zentrum-Worben-API ab
   (salezentrumworben.api.melon.sale).
3. Fuehrt beide Projekte in einer gemeinsamen, normalisierten Datenstruktur
   zusammen.
4. Schreibt:
   - data.json          (Rohdaten, maschinenlesbar, fuer Nachvollziehbarkeit)
   - historie.xlsx       (ein Snapshot-Tab pro Lauf + "Verlauf"-Tab)
   - index.html          (fertiges, eigenstaendiges Dashboard fuer GitHub Pages)

Wichtig: Wenn der Abruf einer Quelle fehlschlaegt (z.B. Blockade durch
Cloud-IP), bricht das Skript NICHT komplett ab, sondern meldet den Fehler
klar in der Konsole und im Dashboard, und behaelt die letzten bekannten
Daten fuer diese Quelle (aus data.json), damit nichts verloren geht.
"""

import json
import os
import random
import re
import sys
import time
from datetime import datetime, timezone

import requests
from bs4 import BeautifulSoup

try:
    import pandas as pd
except ImportError:
    pd = None  # Excel-Export wird dann uebersprungen, Rest funktioniert trotzdem

# ---------------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------------

MAPRAZ_URL = "https://www.mapraz-parc.ch/louez-votre-appartement-a-mapraz-parc-1413"
WORBEN_API_URL = "https://salezentrumworben.api.melon.sale/api/v2/objects/"
WORBEN_SITE = "https://salezentrumworben.melon.sale"

DATA_JSON = "data.json"
HISTORIE_XLSX = "historie.xlsx"
DASHBOARD_HTML = "index.html"

MAPRAZ_STATUS_MAP = {
    "disponible": "Verfuegbar",
    "loué":       "Vermietet",
    "loue":       "Vermietet",
    "attribué":   "Attribuiert",
    "attribue":   "Attribuiert",
    "réservé":    "Reserviert",
    "reserve":    "Reserviert",
}

WORBEN_STATUS_MAP = {
    "free":     "Frei",
    "reserved": "Reserviert",
    "sold":     "Verkauft",
}

BROWSER_HEADERS_BASE = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "de-CH,de;q=0.9,fr-CH;q=0.8,fr;q=0.7,en;q=0.6",
}


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


# ---------------------------------------------------------------------------
# Quelle 1: Mapraz Parc (Vermietung, Ecublens VD)
# ---------------------------------------------------------------------------

def parse_chf(text):
    try:
        clean = re.sub(r"[^0-9]", "", text)
        return int(clean) if clean else 0
    except Exception:
        return 0


def normalize_mapraz_status(raw):
    low = raw.strip().lower()
    for key, val in MAPRAZ_STATUS_MAP.items():
        if key in low:
            return val
    return raw.strip() or "Unbekannt"


def fetch_mapraz():
    log("Mapraz Parc: lade Webseite ...")
    session = requests.Session()
    headers = {
        **BROWSER_HEADERS_BASE,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate, br",
        "Referer": "https://www.mapraz-parc.ch/",
        "Connection": "keep-alive",
    }
    try:
        session.get("https://www.mapraz-parc.ch/", headers=headers, timeout=20)
        time.sleep(random.uniform(0.8, 1.5))
        r = session.get(MAPRAZ_URL, headers=headers, timeout=30)
        r.raise_for_status()
    except requests.RequestException as e:
        log(f"Mapraz Parc: FEHLER beim Laden ({e})")
        return None, str(e)

    soup = BeautifulSoup(r.text, "html.parser")
    wohnungen = []
    aktuell_geb = ""

    for element in soup.find_all(True):
        text = element.get_text(separator=" ", strip=True)

        if element.name in ("h2", "h3", "div", "p", "span", "strong", "b"):
            t = text.upper()
            if "IMMEUBLE A" in t and len(text) < 50:
                aktuell_geb = "A"
            elif "IMMEUBLE B" in t and len(text) < 50:
                aktuell_geb = "B"
            elif "IMMEUBLE C" in t and len(text) < 50:
                aktuell_geb = "C"

        if element.name == "tr":
            tds = element.find_all("td")
            if len(tds) < 9:
                continue
            lot_text = tds[0].get_text(strip=True)
            if not lot_text.startswith("Lot"):
                continue

            num = tds[1].get_text(strip=True)
            typ = tds[2].get_text(strip=True)
            flaeche = tds[3].get_text(strip=True)
            loyer_r = tds[6].get_text(strip=True)
            status_r = tds[8].get_text(strip=True)

            geb = aktuell_geb or (num[0] if num and num[0] in ("A", "B", "C") else "")

            wohnungen.append({
                "projekt": "Mapraz Parc",
                "typ_vermarktung": "Miete",
                "gebaeude": geb,
                "referenz": num,
                "zimmer": typ,
                "flaeche": flaeche,
                "preis_chf": parse_chf(loyer_r),
                "preis_einheit": "CHF/Monat (netto)",
                "status_roh": status_r,
                "status": normalize_mapraz_status(status_r),
            })

    log(f"Mapraz Parc: {len(wohnungen)} Wohnungen gefunden.")
    if not wohnungen:
        return None, "Keine Wohnungen im HTML gefunden (Seitenstruktur evtl. geaendert)"
    return wohnungen, None


# ---------------------------------------------------------------------------
# Quelle 2: Zentrum Worben (Verkauf, Haus B1)
# ---------------------------------------------------------------------------

def normalize_worben_status(raw_key):
    return WORBEN_STATUS_MAP.get(raw_key, raw_key or "Unbekannt")


def fetch_worben():
    log("Zentrum Worben: lade API ...")
    session = requests.Session()
    headers = {
        **BROWSER_HEADERS_BASE,
        "Accept": "application/json, text/plain, */*",
        "Referer": WORBEN_SITE + "/",
        "Origin": WORBEN_SITE,
    }
    try:
        r = session.get(WORBEN_API_URL, headers=headers, timeout=30)
        r.raise_for_status()
        data = r.json()
    except requests.RequestException as e:
        log(f"Zentrum Worben: FEHLER beim Laden ({e})")
        return None, str(e)
    except ValueError:
        log("Zentrum Worben: Antwort war kein gueltiges JSON (evtl. blockiert)")
        return None, "Antwort war kein gueltiges JSON"

    objekte = []
    for o in data:
        building = (o.get("building") or {}).get("building_title", "")
        objekte.append({
            "projekt": "Zentrum Worben",
            "typ_vermarktung": "Verkauf",
            "gebaeude": building,
            "referenz": o.get("reference_number", ""),
            "zimmer": o.get("rooms"),
            "flaeche": o.get("area"),
            "preis_chf": o.get("selling_price") or 0,
            "preis_einheit": "CHF Kaufpreis",
            "status_roh": o.get("object_state_text", ""),
            "status": normalize_worben_status(o.get("object_state", "")),
        })

    log(f"Zentrum Worben: {len(objekte)} Objekte gefunden.")
    if not objekte:
        return None, "Keine Objekte in der API-Antwort"
    return objekte, None


# ---------------------------------------------------------------------------
# Konsolidierung, Speicherung, Dashboard
# ---------------------------------------------------------------------------

def load_previous():
    if os.path.exists(DATA_JSON):
        with open(DATA_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def main():
    now = datetime.now(timezone.utc).astimezone()
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S %Z")

    previous = load_previous()

    mapraz_data, mapraz_err = fetch_mapraz()
    worben_data, worben_err = fetch_worben()

    # Fallback auf letzten bekannten Stand, falls ein Abruf fehlschlaegt
    used_fallback = []
    if mapraz_data is None and previous:
        mapraz_data = [o for o in previous.get("objekte", []) if o["projekt"] == "Mapraz Parc"]
        if mapraz_data:
            used_fallback.append("Mapraz Parc (alte Daten beibehalten)")
    if worben_data is None and previous:
        worben_data = [o for o in previous.get("objekte", []) if o["projekt"] == "Zentrum Worben"]
        if worben_data:
            used_fallback.append("Zentrum Worben (alte Daten beibehalten)")

    alle_objekte = (mapraz_data or []) + (worben_data or [])

    result = {
        "generiert_am": timestamp,
        "quellen": {
            "mapraz_parc": {
                "url": MAPRAZ_URL,
                "status": "ok" if mapraz_err is None else "fehler",
                "fehler": mapraz_err,
                "anzahl": len(mapraz_data or []),
            },
            "zentrum_worben": {
                "url": WORBEN_API_URL,
                "status": "ok" if worben_err is None else "fehler",
                "fehler": worben_err,
                "anzahl": len(worben_data or []),
            },
        },
        "verwendete_fallbacks": used_fallback,
        "objekte": alle_objekte,
    }

    with open(DATA_JSON, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    log(f"Geschrieben: {DATA_JSON} ({len(alle_objekte)} Objekte total)")

    if pd is not None and alle_objekte:
        write_excel_history(alle_objekte, now)

    write_dashboard(result)
    log(f"Geschrieben: {DASHBOARD_HTML}")

    # Exit-Code: Fehler, wenn BEIDE Quellen fehlgeschlagen sind und kein Fallback half
    if not alle_objekte:
        log("FEHLER: Keine Daten aus keiner Quelle verfuegbar.")
        sys.exit(1)


def write_excel_history(objekte, now):
    tab_name = now.strftime("%Y-%m-%d_%H%M")
    df = pd.DataFrame(objekte)
    df["Zeitstempel"] = now.strftime("%Y-%m-%d %H:%M:%S")

    if os.path.exists(HISTORIE_XLSX):
        with pd.ExcelWriter(HISTORIE_XLSX, engine="openpyxl", mode="a",
                             if_sheet_exists="replace") as w:
            df.to_excel(w, sheet_name=tab_name, index=False)
        alle = pd.read_excel(HISTORIE_XLSX, sheet_name=None)
        frames = [d for name, d in alle.items() if name != "Verlauf"]
        verlauf = pd.concat(frames, ignore_index=True) if frames else df
        with pd.ExcelWriter(HISTORIE_XLSX, engine="openpyxl", mode="a",
                             if_sheet_exists="replace") as w:
            verlauf.to_excel(w, sheet_name="Verlauf", index=False)
    else:
        with pd.ExcelWriter(HISTORIE_XLSX, engine="openpyxl") as w:
            df.to_excel(w, sheet_name=tab_name, index=False)
            df.to_excel(w, sheet_name="Verlauf", index=False)
    log(f"Geschrieben: {HISTORIE_XLSX} (Tab {tab_name})")


DASHBOARD_TEMPLATE = r"""<!DOCTYPE html>
<html lang="de-CH">
<head>
<meta charset="UTF-8">
<title>Vermarktungs-Dashboard – Mapraz Parc & Zentrum Worben</title>
<style>
  :root{
    --bg:#f4f5f7; --card:#ffffff; --ink:#1f2430; --muted:#6b7280;
    --frei:#16a34a; --reserviert:#f59e0b; --belegt:#2563eb;
    --border:#e5e7eb;
  }
  *{box-sizing:border-box;}
  body{margin:0;padding:24px;background:var(--bg);color:var(--ink);
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif;}
  .wrap{max-width:1180px;margin:0 auto;}
  header h1{margin:0 0 4px 0;font-size:26px;}
  header .sub{color:var(--muted);font-size:14px;}
  .source-note{margin-top:10px;font-size:12px;color:var(--muted);
    background:#fff;border:1px solid var(--border);border-radius:8px;padding:8px 12px;
    display:inline-block;}
  .source-note.warn{background:#fef3c7;color:#92400e;border-color:#fde68a;}
  .kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));
    gap:12px;margin:20px 0;}
  .kpi{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:16px;}
  .kpi .label{font-size:12px;color:var(--muted);margin-bottom:6px;}
  .kpi .value{font-size:24px;font-weight:700;}
  .kpi.frei .value{color:var(--frei);}
  .kpi.reserviert .value{color:var(--reserviert);}
  .kpi.belegt .value{color:var(--belegt);}
  .project-section{background:var(--card);border:1px solid var(--border);border-radius:12px;
    padding:18px;margin-bottom:20px;}
  .project-section h2{margin:0 0 4px 0;font-size:17px;}
  .project-section .meta{color:var(--muted);font-size:12px;margin-bottom:14px;}
  table{width:100%;border-collapse:collapse;font-size:13px;}
  th,td{text-align:left;padding:7px 9px;border-bottom:1px solid var(--border);}
  th{color:var(--muted);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.02em;}
  tr:last-child td{border-bottom:none;}
  .badge{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:600;}
  .badge.frei{background:#dcfce7;color:#166534;}
  .badge.reserviert{background:#fef3c7;color:#92400e;}
  .badge.belegt{background:#dbeafe;color:#1e40af;}
  footer{margin-top:20px;font-size:12px;color:var(--muted);text-align:center;}
  .grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;}
  @media (max-width:820px){.grid2{grid-template-columns:1fr;}}
  canvas{max-width:100%;}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>Vermarktungs-Dashboard</h1>
    <div class="sub">Mapraz Parc (Vermietung) &middot; Zentrum Worben (Verkauf)</div>
    <div id="sourceNotes"></div>
  </header>

  <div class="kpi-grid" id="kpiGrid"></div>

  <div class="grid2">
    <div class="project-section">
      <h2>Status übergreifend</h2>
      <canvas id="statusChart" width="500" height="240"></canvas>
    </div>
    <div class="project-section">
      <h2>Einheiten pro Projekt</h2>
      <canvas id="projectChart" width="500" height="240"></canvas>
    </div>
  </div>

  <div id="projectTables"></div>

  <footer id="footerNote"></footer>
</div>

<script>
const DATA = __DATA_JSON__;

function chf(n){ return new Intl.NumberFormat('de-CH').format(Math.round(n||0)) + " CHF"; }
function num(n, d){ return new Intl.NumberFormat('de-CH',{maximumFractionDigits:d??0}).format(n||0); }

const OCCUPIED_STATUSES = ["Vermietet", "Verkauft", "Attribuiert"];
const FREE_STATUSES = ["Verfuegbar", "Frei"];
const RESERVED_STATUSES = ["Reserviert"];

function classify(status){
  if (OCCUPIED_STATUSES.includes(status)) return "belegt";
  if (RESERVED_STATUSES.includes(status)) return "reserviert";
  if (FREE_STATUSES.includes(status)) return "frei";
  return "unbekannt";
}

const objekte = DATA.objekte || [];
const total = objekte.length;
const countFrei = objekte.filter(o=>classify(o.status)==="frei").length;
const countReserviert = objekte.filter(o=>classify(o.status)==="reserviert").length;
const countBelegt = objekte.filter(o=>classify(o.status)==="belegt").length;

// --- Quellenhinweise ---
const notesEl = document.getElementById("sourceNotes");
let notesHtml = `<div class="source-note">Generiert am ${DATA.generiert_am}. Quellen: mapraz-parc.ch, salezentrumworben.api.melon.sale — ausschliesslich diese, keine Annahmen.</div>`;
if (DATA.verwendete_fallbacks && DATA.verwendete_fallbacks.length){
  notesHtml += `<div class="source-note warn">Hinweis: ${DATA.verwendete_fallbacks.join(", ")} — letzter erfolgreicher Abruf wurde beibehalten, da der aktuelle Abruf fehlschlug.</div>`;
}
notesEl.innerHTML = notesHtml;

// --- KPIs ---
const kpis = [
  {label:"Einheiten total", value: total, cls:""},
  {label:"Frei / Verfügbar", value: countFrei, cls:"frei"},
  {label:"Reserviert", value: countReserviert, cls:"reserviert"},
  {label:"Vermietet / Verkauft", value: countBelegt, cls:"belegt"},
  {label:"Vermarktungsquote", value: total ? num((countBelegt+countReserviert)/total*100,0)+" %" : "–", cls:""},
];
const kpiGrid = document.getElementById("kpiGrid");
kpis.forEach(k=>{
  const d = document.createElement("div");
  d.className = "kpi " + k.cls;
  d.innerHTML = `<div class="label">${k.label}</div><div class="value">${k.value}</div>`;
  kpiGrid.appendChild(d);
});

// --- Canvas Hilfsfunktionen ---
function setupHiDPI(canvas){
  const ratio = window.devicePixelRatio || 1;
  const w = canvas.width, h = canvas.height;
  canvas.style.width = w+"px"; canvas.style.height = h+"px";
  canvas.width = w*ratio; canvas.height = h*ratio;
  const ctx = canvas.getContext("2d");
  ctx.scale(ratio, ratio);
  return {ctx, w, h};
}

function drawDonut(canvasId, items){
  const canvas = document.getElementById(canvasId);
  const {ctx, w, h} = setupHiDPI(canvas);
  ctx.clearRect(0,0,w,h);
  const cx=w/2, cy=h/2+5, rOuter=Math.min(w,h)/2-10, rInner=rOuter*0.6;
  const total = items.reduce((s,i)=>s+i.value,0) || 1;
  let start=-Math.PI/2;
  items.forEach(it=>{
    const ang=(it.value/total)*Math.PI*2;
    ctx.beginPath(); ctx.moveTo(cx,cy);
    ctx.arc(cx,cy,rOuter,start,start+ang); ctx.closePath();
    ctx.fillStyle=it.color; ctx.fill();
    start+=ang;
  });
  ctx.beginPath(); ctx.arc(cx,cy,rInner,0,Math.PI*2); ctx.fillStyle="#fff"; ctx.fill();
  ctx.fillStyle="#1f2430"; ctx.font="bold 20px -apple-system,sans-serif"; ctx.textAlign="center";
  ctx.fillText(items.reduce((s,i)=>s+i.value,0), cx, cy+7);
  ctx.font="11px -apple-system,sans-serif"; ctx.fillStyle="#6b7280";
  ctx.fillText("Einheiten", cx, cy+23);

  const legend = document.createElement("div");
  legend.style.cssText = "display:flex;gap:14px;font-size:12px;color:#6b7280;margin-top:10px;flex-wrap:wrap;justify-content:center;";
  legend.innerHTML = items.map(it=>`<span style="display:inline-flex;align-items:center;gap:5px;"><span style="width:9px;height:9px;border-radius:50%;background:${it.color};display:inline-block;"></span>${it.label} (${it.value})</span>`).join("");
  canvas.parentNode.appendChild(legend);
}

function drawBars(canvasId, items){
  const canvas = document.getElementById(canvasId);
  const {ctx, w, h} = setupHiDPI(canvas);
  ctx.clearRect(0,0,w,h);
  const pad = {top:20,right:20,bottom:40,left:40};
  const cw = w-pad.left-pad.right, ch = h-pad.top-pad.bottom;
  const maxVal = Math.max(...items.map(i=>i.value),1)*1.2;
  const gap=16, barW = cw/items.length - gap;
  ctx.strokeStyle="#e5e7eb"; ctx.fillStyle="#6b7280"; ctx.font="11px -apple-system,sans-serif"; ctx.textAlign="right";
  for(let s=0;s<=4;s++){
    const y = pad.top+ch-(ch*s/4);
    ctx.beginPath(); ctx.moveTo(pad.left,y); ctx.lineTo(w-pad.right,y); ctx.stroke();
    ctx.fillText(Math.round(maxVal*s/4), pad.left-8, y+4);
  }
  items.forEach((it,i)=>{
    const barH = (it.value/maxVal)*ch;
    const x = pad.left + i*(barW+gap) + gap/2;
    const y = pad.top+ch-barH;
    ctx.fillStyle = it.color;
    ctx.fillRect(x,y,barW,barH);
    ctx.fillStyle="#1f2430"; ctx.font="bold 12px -apple-system,sans-serif"; ctx.textAlign="center";
    ctx.fillText(it.value, x+barW/2, y-6);
    ctx.fillStyle="#6b7280"; ctx.font="11px -apple-system,sans-serif";
    ctx.fillText(it.label, x+barW/2, pad.top+ch+16);
  });
}

drawDonut("statusChart", [
  {label:"Frei", value:countFrei, color:"#16a34a"},
  {label:"Reserviert", value:countReserviert, color:"#f59e0b"},
  {label:"Vermietet/Verkauft", value:countBelegt, color:"#2563eb"},
]);

const byProject = {};
objekte.forEach(o=>{ byProject[o.projekt] = (byProject[o.projekt]||0)+1; });
drawBars("projectChart", Object.keys(byProject).map((p,i)=>({
  label:p, value:byProject[p], color: i===0 ? "#0ea5e9" : "#7c3aed"
})));

// --- Tabellen pro Projekt ---
const tablesEl = document.getElementById("projectTables");
const projekte = [...new Set(objekte.map(o=>o.projekt))];
projekte.forEach(proj=>{
  const rows = objekte.filter(o=>o.projekt===proj).sort((a,b)=>(a.referenz||"").localeCompare(b.referenz||""));
  const frei = rows.filter(o=>classify(o.status)==="frei").length;
  const res = rows.filter(o=>classify(o.status)==="reserviert").length;
  const belegt = rows.filter(o=>classify(o.status)==="belegt").length;
  const typ = rows[0]?.typ_vermarktung || "";
  const einheit = rows[0]?.preis_einheit || "";

  const section = document.createElement("div");
  section.className = "project-section";
  section.innerHTML = `
    <h2>${proj} <span style="font-weight:400;color:#6b7280;font-size:13px;">(${typ})</span></h2>
    <div class="meta">${rows.length} Einheiten &middot; Frei: ${frei} &middot; Reserviert: ${res} &middot; Vermietet/Verkauft: ${belegt}</div>
    <table>
      <thead><tr>
        <th>Gebäude</th><th>Referenz</th><th>Zimmer</th><th>Fläche</th><th>Status</th><th>${einheit}</th>
      </tr></thead>
      <tbody>
        ${rows.map(o=>{
          const cls = classify(o.status);
          return `<tr>
            <td>${o.gebaeude||"–"}</td>
            <td>${o.referenz||"–"}</td>
            <td>${o.zimmer??"–"}</td>
            <td>${o.flaeche??"–"}</td>
            <td><span class="badge ${cls}">${o.status}</span></td>
            <td>${chf(o.preis_chf)}</td>
          </tr>`;
        }).join("")}
      </tbody>
    </table>
  `;
  tablesEl.appendChild(section);
});

document.getElementById("footerNote").textContent =
  `Automatisch generiert · ${DATA.generiert_am} · ${total} Einheiten total`;
</script>
</body>
</html>
"""


def write_dashboard(result):
    html = DASHBOARD_TEMPLATE.replace("__DATA_JSON__", json.dumps(result, ensure_ascii=False))
    with open(DASHBOARD_HTML, "w", encoding="utf-8") as f:
        f.write(html)


if __name__ == "__main__":
    main()
