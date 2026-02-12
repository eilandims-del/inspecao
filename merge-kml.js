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
  const CATEGORY_BY_ALIM = new Map([
    // CANINDÉ
    ["CND01C1","Canindé"], ["CND01C2","Canindé"], ["CND01C3","Canindé"], ["CND01C4","Canindé"], ["CND01C5","Canindé"], ["CND01C6","Canindé"],
    ["INP01N3","Canindé"], ["INP01N4","Canindé"], ["INP01N5","Canindé"],
    ["BVG01P1","Canindé"], ["BVG01P2","Canindé"], ["BVG01P3","Canindé"], ["BVG01P4","Canindé"],
    ["MCA01L1","Canindé"], ["MCA01L2","Canindé"], ["MCA01L3","Canindé"],
  
    // QUIXADÁ
    ["BNB01Y2","Quixadá"],
    ["JTM01N2","Quixadá"],
    ["QXD01P1","Quixadá"], ["QXD01P2","Quixadá"], ["QXD01P3","Quixadá"], ["QXD01P4","Quixadá"], ["QXD01P5","Quixadá"], ["QXD01P6","Quixadá"],
    ["QXB01N2","Quixadá"], ["QXB01N3","Quixadá"], ["QXB01N4","Quixadá"], ["QXB01N5","Quixadá"], ["QXB01N6","Quixadá"], ["QXB01N7","Quixadá"],
  
    // NOVA RUSSAS
    ["IPU01L2","Nova Russas"], ["IPU01L3","Nova Russas"], ["IPU01L4","Nova Russas"], ["IPU01L5","Nova Russas"],
    ["ARR01L1","Nova Russas"], ["ARR01L2","Nova Russas"], ["ARR01L3","Nova Russas"],
    ["SQT01F2","Nova Russas"], ["SQT01F3","Nova Russas"], ["SQT01F4","Nova Russas"],
    ["ARU01Y1","Nova Russas"], ["ARU01Y2","Nova Russas"], ["ARU01Y4","Nova Russas"], ["ARU01Y5","Nova Russas"], ["ARU01Y6","Nova Russas"], ["ARU01Y7","Nova Russas"], ["ARU01Y8","Nova Russas"],
    ["NVR01N1","Nova Russas"], ["NVR01N2","Nova Russas"], ["NVR01N3","Nova Russas"], ["NVR01N5","Nova Russas"],
    ["MTB01S2","Nova Russas"], ["MTB01S3","Nova Russas"], ["MTB01S4","Nova Russas"],
  
    // CRATEÚS
    ["IDP01I1","Crateús"], ["IDP01I2","Crateús"], ["IDP01I3","Crateús"], ["IDP01I4","Crateús"],
    ["CAT01C1","Crateús"], ["CAT01C2","Crateús"], ["CAT01C3","Crateús"], ["CAT01C4","Crateús"], ["CAT01C5","Crateús"], ["CAT01C6","Crateús"], ["CAT01C7","Crateús"],
  ]);
  

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
    const alim = String(row.ALIMENTADOR || "").trim().toUpperCase();
    if (alim && CATEGORY_BY_ALIM.has(alim)) return CATEGORY_BY_ALIM.get(alim);
  
    // fallback: tenta prefixo 3 letras do ALIMENTADOR ou do DISPOSITIVO
    const tryPrefix = (val) => {
      const s = String(val || "").toUpperCase();
      const m = s.match(/[A-Z]{3}/);
      return m ? m[0] : "";
    };
  
    const pA = tryPrefix(alim);
    const pD = tryPrefix(row.DISPOSITIVO);
  
    // fallback simples por prefixo (mantém seu comportamento antigo)
    const prefixMap = {
      "CND":"Canindé","INP":"Canindé","BVG":"Canindé","MCA":"Canindé",
      "BNB":"Quixadá","JTM":"Quixadá","QXD":"Quixadá","QXB":"Quixadá",
      "IPU":"Nova Russas","ARR":"Nova Russas","SQT":"Nova Russas","ARU":"Nova Russas","NVR":"Nova Russas","MTB":"Nova Russas",
      "IDP":"Crateús","CAT":"Crateús",
    };
  
    if (pA && prefixMap[pA]) return prefixMap[pA];
    if (pD && prefixMap[pD]) return prefixMap[pD];
  
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
      <name>${escapeXml(r.DISPOSITIVO)}</name>   <!-- 🔥 FORÇA usar DISPOSITIVO -->
    
      <Style>
        <IconStyle>
          <color>${color}</color>
          <scale>1.8</scale>
          <Icon>
            <href>http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png</href>
          </Icon>
        </IconStyle>
      </Style>
    
      <description><![CDATA[
        <div style="font-family: Arial; font-size: 13px;">
          <b>CATEGORIA:</b> ${escapeXml(cat)}<br/>
          <b>TIPO:</b> ${escapeXml(tipo)}<br/>
          <b>DISPOSITIVO_PROTECAO:</b> ${escapeXml(r.DISPOSITIVO)}<br/>
          <b>NÚMERO OT:</b> ${escapeXml(r.NUMERO_OT || "-")}<br/>
          <b>ALIMENTADOR:</b> ${escapeXml(r.ALIMENTADOR || "-")}<br/>
          <b>INSTALACAO_NOVA:</b> ${escapeXml(r.INSTALACAO_NOVA || "-")}<br/>
        </div>
      ]]></description>
    
      <Point>
        <coordinates>${geo.lon},${geo.lat},0</coordinates>
      </Point>
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
