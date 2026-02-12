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

async function readXlsxFile(file) {
  const ab = await file.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
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

function buildKml(rows, idx) {
  const placemarks = [];
  let missing = 0;

  for (const r of rows) {
    const geo = idx.get(r.key);
    if (!geo) { missing++; continue; }

    const color = r.TIPO === "INSPECAO" ? "ff800080" : "ffffffff"; 
    // Roxo = ff800080 | Branco = ffffffff

    placemarks.push(`
<Placemark>
  <name>${r.DISPOSITIVO}</name>
  <Style>
    <IconStyle>
      <color>${color}</color>
      <scale>1.6</scale>
      <Icon>
        <href>http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png</href>
      </Icon>
    </IconStyle>
  </Style>
  <description>
    <![CDATA[
    <b>TIPO:</b> ${r.TIPO}<br/>
    <b>OT:</b> ${r.NUMERO_OT}<br/>
    <b>ALIMENTADOR:</b> ${r.ALIMENTADOR}<br/>
    ]]>
  </description>
  <Point>
    <coordinates>${geo.lon},${geo.lat},0</coordinates>
  </Point>
</Placemark>
`);
  }

  return {
    kml: `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
${placemarks.join("\n")}
</Document>
</kml>`,
    missing
  };
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

  const ins = buildFromInspecao(await readXlsxFile(fIns));
  const rei = buildFromReiteradas(await readXlsxFile(fRei));

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
