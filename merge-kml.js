const $ = (id) => document.getElementById(id);
const statusEl = $("status");

let mergedRows = [];
let kmlIndex = new Map();

function setStatus(msg) {
  statusEl.textContent = msg;
}

function normalizeKey(v) {
  return String(v ?? "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

async function readXlsxWorkbook(file) {
  const ab = await file.arrayBuffer();
  return XLSX.read(ab, { type: "array" });
}

function sheetToRows(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Aba "${sheetName}" não encontrada no arquivo.`);
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
}


function colIndex(letter) {
  let n = 0;
  for (const ch of letter.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

async function readKmlIndex(file) {
  const fname = file.name.toLowerCase();
  let kmlText = "";

  if (fname.endsWith(".kmz")) {
    const ab = await file.arrayBuffer();
    const u8 = new Uint8Array(ab);
    const unzipped = window.fflate.unzipSync(u8);

    let kmlEntry = unzipped["doc.kml"];
    if (!kmlEntry) {
      const key = Object.keys(unzipped).find(k => k.endsWith(".kml"));
      kmlEntry = unzipped[key];
    }

    kmlText = new TextDecoder().decode(kmlEntry);
  } else {
    kmlText = await file.text();
  }

  const doc = new DOMParser().parseFromString(kmlText, "text/xml");
  const placemarks = [...doc.getElementsByTagName("Placemark")];

  const idx = new Map();

  for (const pm of placemarks) {
    const name = pm.getElementsByTagName("name")[0]?.textContent ?? "";
    const coords = pm.getElementsByTagName("coordinates")[0]?.textContent ?? "";
    const first = coords.trim().split(/\s+/)[0] || "";
    const [lon, lat] = first.split(",").map(Number);

    const key = normalizeKey(name);
    if (!key || !lat || !lon) continue;

    if (!idx.has(key)) idx.set(key, { lat, lon });
  }

  return idx;
}

function buildFromInspecao(rows) {
  const iE = colIndex("E");
  const iH = colIndex("H");
  const iAP = colIndex("AP");

  return rows.slice(1).map(row => {
    const disp = row[iAP] ?? "";
    return {
      key: normalizeKey(disp),
      TIPO: "INSPECAO",
      DISPOSITIVO: disp,
      INSTALACAO_NOVA: row[iE] ?? "",
      NUMERO_OT: row[iH] ?? "",
      ALIMENTADOR: ""
    };
  }).filter(r => r.key);
}

function buildFromReiteradas(rows) {
  const iA = colIndex("A");
  const iC = colIndex("C");

  return rows.slice(1).map(row => {
    const disp = row[iA] ?? "";
    return {
      key: normalizeKey(disp),
      TIPO: "REITERADA",
      DISPOSITIVO: disp,
      INSTALACAO_NOVA: "",
      NUMERO_OT: "",
      ALIMENTADOR: row[iC] ?? ""
    };
  }).filter(r => r.key);
}

function mergeAndDiff(ins, rei) {
  const setIns = new Set(ins.map(x => x.key));
  const setRei = new Set(rei.map(x => x.key));
  const intersection = new Set([...setIns].filter(k => setRei.has(k)));

  return [
    ...rei.filter(x => !intersection.has(x.key)),
    ...ins.filter(x => !intersection.has(x.key))
  ];
}

function escapeXml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}


function buildKml(rows, idx) {
  const CATEGORY_PREFIXES = {
    "Canindé":     ["CND", "BVG", "INP", "MCA"],
    "Nova Russas": ["ARU", "SQT", "MTB", "IPU", "ARR", "NVR"],
    "Crateús":     ["CAT", "IDP"],
    "Quixadá":     ["BNB", "QXD", "QXB", "JTM"]
  };

  // prefixo -> categoria
  const prefixToCategory = new Map();
  for (const [cat, prefixes] of Object.entries(CATEGORY_PREFIXES)) {
    for (const p of prefixes) prefixToCategory.set(p.toUpperCase(), cat);
  }

  function extractPrefix3(value) {
    const s = String(value ?? "").trim().toUpperCase();
    const m = s.match(/[A-Z]{3}/); // primeira sequência de 3 letras
    return m ? m[0] : "";
  }

  function getCategory(row) {
    const p1 = extractPrefix3(row.DISPOSITIVO);
    const p2 = extractPrefix3(row.ALIMENTADOR);

    if (p1 && prefixToCategory.has(p1)) return prefixToCategory.get(p1);
    if (p2 && prefixToCategory.has(p2)) return prefixToCategory.get(p2);
    return "Outros";
  }

  // ===== Agrupamento: categoria -> tipo -> placemarks =====
  // Ex.: groups.get("Canindé").get("INSPECAO") => [xml, xml...]
  const groups = new Map();
  let missing = 0;

  const PUSH_PIN = "http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png";

  for (const r of rows) {
    const geo = idx.get(r.key);
    if (!geo) { missing++; continue; }

    const cat = getCategory(r);
    const tipo = r.TIPO === "INSPECAO" ? "INSPEÇÃO" : "REITERADA";

    if (!groups.has(cat)) groups.set(cat, new Map([["INSPEÇÃO", []], ["REITERADA", []]]));
    const sub = groups.get(cat);

    // Roxo INSPEÇÃO | Branco REITERADA
    const color = (tipo === "INSPEÇÃO") ? "ff800080" : "ffffffff";

    const dispositivo = (r.DISPOSITIVO ?? "").toString();
    const ot = (r.NUMERO_OT ?? "").toString();
    const alim = (r.ALIMENTADOR ?? "").toString();
    const inst = (r.INSTALACAO_NOVA ?? "").toString();

    const placemark = `
<Placemark>
  <name>${escapeXml(dispositivo)}</name>

  <Style>
    <IconStyle>
      <color>${color}</color>
      <scale>1.8</scale>
      <Icon><href>${PUSH_PIN}</href></Icon>
    </IconStyle>
  </Style>

  <description><![CDATA[
    <div style="font-family: Arial; font-size: 13px;">
      <b>CATEGORIA:</b> ${escapeXml(cat)}<br/>
      <b>TIPO:</b> ${escapeXml(tipo)}<br/>
      <b>DISPOSITIVO_PROTECAO / ELEMENTO:</b> ${escapeXml(dispositivo)}<br/>
      <b>NÚMERO OT:</b> ${escapeXml(ot || "-")}<br/>
      <b>ALIMENTADOR:</b> ${escapeXml(alim || "-")}<br/>
      <b>INSTALACAO_NOVA:</b> ${escapeXml(inst || "-")}<br/>
    </div>
  ]]></description>

  <ExtendedData>
    <Data name="CATEGORIA"><value>${escapeXml(cat)}</value></Data>
    <Data name="TIPO"><value>${escapeXml(tipo)}</value></Data>
    <Data name="DISPOSITIVO"><value>${escapeXml(dispositivo)}</value></Data>
    <Data name="NUMERO_OT"><value>${escapeXml(ot)}</value></Data>
    <Data name="ALIMENTADOR"><value>${escapeXml(alim)}</value></Data>
    <Data name="INSTALACAO_NOVA"><value>${escapeXml(inst)}</value></Data>
  </ExtendedData>

  <Point><coordinates>${geo.lon},${geo.lat},0</coordinates></Point>
</Placemark>
`;

    sub.get(tipo).push(placemark);
  }

  // Ordem fixa
  const orderedCats = ["Canindé", "Nova Russas", "Crateús", "Quixadá", "Outros"];

  function folderBlock(catName, tipoName, placemarks) {
    const colorDot = (tipoName === "INSPEÇÃO") ? "🟣" : "⚪";
    return `
<Folder>
  <name>${escapeXml(colorDot + " " + tipoName)}</name>
  ${placemarks.join("\n")}
</Folder>`;
  }

  const folders = orderedCats
    .filter(cat => groups.has(cat))
    .map(cat => {
      const sub = groups.get(cat);
      const insp = sub.get("INSPEÇÃO") || [];
      const rei = sub.get("REITERADA") || [];

      return `
<Folder>
  <name>${escapeXml(cat)}</name>
  ${folderBlock(cat, "INSPEÇÃO", insp)}
  ${folderBlock(cat, "REITERADA", rei)}
</Folder>`;
    })
    .join("\n");

  const kml = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>Resultado - Reiteradas x Inspeção</name>
    ${folders}
  </Document>
</kml>`;

  return { kml, missing };
}

function download(text, filename, type) {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

$("btnGerarPlanilha").addEventListener("click", async () => {
  const fIns = $("fileInspecao").files[0];
  const fRei = $("fileReiteradas").files[0];

  if (!fIns || !fRei) {
    setStatus("Envie as duas planilhas.");
    return;
  }

  setStatus("Processando planilhas...");

  const wbIns = await readXlsxWorkbook(fIns);
  const insRows = sheetToRows(wbIns, "PBM-CE - Inspecao"); // ✅ força a aba certa
  const ins = buildFromInspecao(insRows);
  
  const wbRei = await readXlsxWorkbook(fRei);
  const reiRows = sheetToRows(wbRei, wbRei.SheetNames[0]); // reiteradas = primeira aba (padrão)
  const rei = buildFromReiteradas(reiRows);
  

  mergedRows = mergeAndDiff(ins, rei);

  const ws = XLSX.utils.json_to_sheet(mergedRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RESULTADO");

  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  download(buf, "resultado.xlsx", "application/octet-stream");

  $("btnGerarKml").disabled = false;
  setStatus("Planilha gerada com sucesso.");
});

$("btnGerarKml").addEventListener("click", async () => {
  const fKml = $("fileKmlGeral").files[0];
  if (!fKml) {
    setStatus("Envie o KML/KMZ geral.");
    return;
  }

  setStatus("Gerando KML final...");

  const idx = await readKmlIndex(fKml);
  const { kml, missing } = buildKml(mergedRows, idx);

  download(kml, "resultado_google_earth.kml", "application/vnd.google-earth.kml+xml");

  setStatus(`KML gerado com sucesso.\nSem coordenadas encontradas: ${missing}`);
});
