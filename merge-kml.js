/* merge-kml.js
   - Lê 2 XLSX + 1 KML/KMZ
   - Remove interseção (itens presentes nas duas planilhas)
   - Gera XLSX final + KML final (pins azul/verde) para Google Earth
*/

const $ = (id) => document.getElementById(id);
const statusEl = $("status");

let mergedRows = [];          // resultado final (linhas exclusivas)
let kmlIndex = new Map();     // chave(normalizada) -> { lat, lon, rawName }

function setStatus(msg) {
  statusEl.textContent = `Status:\n${msg}`;
}

// Normalização agressiva: remove tudo que não seja letra/número
function normalizeKey(v) {
  return String(v ?? "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

// ========== XLSX helpers ==========
async function readXlsxFile(file) {
  const ab = await file.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  // header:1 => matriz [ [c0,c1,c2...] ... ]
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  return rows;
}

// converte letra de coluna para índice (A=0, B=1, ..., AP=41)
function colIndex(letter) {
  let n = 0;
  for (const ch of letter.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

// ========== KML/KMZ parsing ==========
async function readKmlIndex(file) {
  const fname = (file?.name || "").toLowerCase();
  let kmlText = "";

  if (fname.endsWith(".kmz")) {
    // abre KMZ (zip) e pega o .kml interno (geralmente doc.kml)
    const ab = await file.arrayBuffer();
    const u8 = new Uint8Array(ab);

    const unzipped = window.fflate.unzipSync(u8);

    let kmlEntry = unzipped["doc.kml"];
    if (!kmlEntry) {
      const kmlKey = Object.keys(unzipped).find(k => k.toLowerCase().endsWith(".kml"));
      if (!kmlKey) throw new Error("KMZ não contém arquivo .kml (ex.: doc.kml).");
      kmlEntry = unzipped[kmlKey];
    }

    kmlText = new TextDecoder("utf-8").decode(kmlEntry);
  } else {
    // KML direto
    kmlText = await file.text();
  }

  const doc = new DOMParser().parseFromString(kmlText, "text/xml");
  const placemarks = [...doc.getElementsByTagName("Placemark")];

  const idx = new Map();

  for (const pm of placemarks) {
    const pmName = pm.getElementsByTagName("name")[0]?.textContent ?? "";
    const coords = pm.getElementsByTagName("coordinates")[0]?.textContent ?? "";

    // coordinates: "lon,lat,alt" (pode ter várias, pegamos a primeira)
    const first = coords.trim().split(/\s+/)[0] || "";
    const [lon, lat] = first.split(",").map((x) => Number(String(x).trim()));

    const key = normalizeKey(pmName);
    if (!key || !Number.isFinite(lat) || !Number.isFinite(lon)) continue;

    // Se houver duplicado, mantém o primeiro (mais seguro)
    if (!idx.has(key)) idx.set(key, { lat, lon, rawName: pmName });
  }

  return idx;
}

// ========== Build lists ==========
function buildFromInspecao(inspecaoRows) {
  // E=4, H=7, AP=41
  const iE = colIndex("E");
  const iH = colIndex("H");
  const iAP = colIndex("AP");

  const out = [];
  for (let r = 1; r < inspecaoRows.length; r++) {
    const row = inspecaoRows[r] || [];
    const instalacao = row[iE] ?? "";
    const numeroOT = row[iH] ?? "";
    const disp = row[iAP] ?? "";

    const key = normalizeKey(disp);
    if (!key) continue;

    out.push({
      key,
      TIPO: "INSPECAO",
      DISPOSITIVO: String(disp),
      INSTALACAO_NOVA: String(instalacao),
      NUMERO_OT: String(numeroOT),
      ALIMENTADOR: ""
    });
  }
  return out;
}

function buildFromReiteradas(reitRows) {
  // A=0, C=2
  const iA = colIndex("A");
  const iC = colIndex("C");

  const out = [];
  for (let r = 1; r < reitRows.length; r++) {
    const row = reitRows[r] || [];
    const elemento = row[iA] ?? "";
    const alim = row[iC] ?? "";

    const key = normalizeKey(elemento);
    if (!key) continue;

    out.push({
      key,
      TIPO: "REITERADA",
      DISPOSITIVO: String(elemento),
      INSTALACAO_NOVA: "",
      NUMERO_OT: "",
      ALIMENTADOR: String(alim)
    });
  }
  return out;
}

// ========== Merge + Diff ==========
function mergeAndDiff(listInspecao, listReiteradas) {
  const setIns = new Set(listInspecao.map(x => x.key));
  const setRei = new Set(listReiteradas.map(x => x.key));

  // interseção => remover
  const intersection = new Set([...setIns].filter(k => setRei.has(k)));

  const onlyIns = listInspecao.filter(x => !intersection.has(x.key));
  const onlyRei = listReiteradas.filter(x => !intersection.has(x.key));

  const merged = [...onlyRei, ...onlyIns];

  for (const x of merged) {
    x.DIFERENCA = (x.TIPO === "REITERADA")
      ? "Está só em REITERADAS (não aparece na inspeção)"
      : "Está só em INSPEÇÃO (não aparece nas reiteradas)";
  }

  return { merged, removedCount: intersection.size };
}

// ========== XLSX output ==========
function downloadXlsx(rows, filename) {
  const ws = XLSX.utils.json_to_sheet(rows.map(r => ({
    TIPO: r.TIPO,
    DISPOSITIVO: r.DISPOSITIVO,
    ALIMENTADOR: r.ALIMENTADOR,
    INSTALACAO_NOVA: r.INSTALACAO_NOVA,
    NUMERO_OT: r.NUMERO_OT,
    DIFERENCA: r.DIFERENCA
  })));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RESULTADO");

  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ========== KML output ==========
function escapeXml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function buildKmlFromRows(rows, idx) {
  // Ícones padrão do Google (compatível com Google Earth)
  const ICON_BLUE = "http://maps.google.com/mapfiles/ms/icons/blue-dot.png";
  const ICON_GREEN = "http://maps.google.com/mapfiles/ms/icons/green-dot.png";

  const styles = `
    <Style id="pinBlue"><IconStyle><Icon><href>${ICON_BLUE}</href></Icon></IconStyle></Style>
    <Style id="pinGreen"><IconStyle><Icon><href>${ICON_GREEN}</href></Icon></IconStyle></Style>
  `;

  const placemarks = [];
  let missing = 0;

  for (const r of rows) {
    const geo = idx.get(r.key);
    if (!geo) { missing++; continue; }

    const styleUrl = (r.TIPO === "REITERADA") ? "#pinBlue" : "#pinGreen";

    const desc = `
      <![CDATA[
        <b>TIPO:</b> ${escapeXml(r.TIPO)}<br/>
        <b>DISPOSITIVO:</b> ${escapeXml(r.DISPOSITIVO)}<br/>
        <b>ALIMENTADOR:</b> ${escapeXml(r.ALIMENTADOR)}<br/>
        <b>INSTALACAO_NOVA:</b> ${escapeXml(r.INSTALACAO_NOVA)}<br/>
        <b>NUMERO_OT:</b> ${escapeXml(r.NUMERO_OT)}<br/>
      ]]>
    `;

    placemarks.push(`
      <Placemark>
        <name>${escapeXml(r.DISPOSITIVO)}</name>
        <styleUrl>${styleUrl}</styleUrl>
        <description>${desc}</description>
        <ExtendedData>
          <Data name="TIPO"><value>${escapeXml(r.TIPO)}</value></Data>
          <Data name="DISPOSITIVO"><value>${escapeXml(r.DISPOSITIVO)}</value></Data>
          <Data name="ALIMENTADOR"><value>${escapeXml(r.ALIMENTADOR)}</value></Data>
          <Data name="INSTALACAO_NOVA"><value>${escapeXml(r.INSTALACAO_NOVA)}</value></Data>
          <Data name="NUMERO_OT"><value>${escapeXml(r.NUMERO_OT)}</value></Data>
        </ExtendedData>
        <Point><coordinates>${geo.lon},${geo.lat},0</coordinates></Point>
      </Placemark>
    `);
  }

  const kml = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>Resultado - Reiteradas x Inspecao</name>
    ${styles}
    ${placemarks.join("\n")}
  </Document>
</kml>`;

  return { kml, missing };
}

function downloadText(text, filename, mime) {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ========== UI wiring ==========
$("btnGerarPlanilha").addEventListener("click", async () => {
  try {
    const fIns = $("fileInspecao").files?.[0];
    const fRei = $("fileReiteradas").files?.[0];

    if (!fIns || !fRei) {
      setStatus("Envie as 2 planilhas (inspeção e reiteradas).");
      return;
    }

    setStatus("Lendo planilhas...");
    const insRows = await readXlsxFile(fIns);
    const reiRows = await readXlsxFile(fRei);

    setStatus("Processando diferença (removendo interseção)...");
    const listIns = buildFromInspecao(insRows);
    const listRei = buildFromReiteradas(reiRows);

    const { merged, removedCount } = mergeAndDiff(listIns, listRei);
    mergedRows = merged;

    setStatus(
      `OK.\n` +
      `Inspeção lida: ${listIns.length}\n` +
      `Reiteradas lidas: ${listRei.length}\n` +
      `Removidos (presentes nas duas): ${removedCount}\n` +
      `Resultado final (exclusivos): ${mergedRows.length}\n\n` +
      `Baixando XLSX...`
    );

    downloadXlsx(mergedRows, "resultado_reiteradas_inspecao.xlsx");

    $("btnGerarKml").disabled = false;
  } catch (e) {
    console.error(e);
    setStatus(`Erro: ${e?.message || e}`);
  }
});

$("btnGerarKml").addEventListener("click", async () => {
  try {
    const fKml = $("fileKmlGeral").files?.[0];
    if (!fKml) {
      setStatus("Envie o KML/KMZ geral primeiro.");
      return;
    }
    if (!mergedRows.length) {
      setStatus("Gere a planilha primeiro (botão GERAR PLANILHA).");
      return;
    }

    setStatus("Lendo KML/KMZ geral e indexando coordenadas...");
    kmlIndex = await readKmlIndex(fKml);

    setStatus(`Gerando KML do resultado (${mergedRows.length} itens)...`);
    const { kml, missing } = buildKmlFromRows(mergedRows, kmlIndex);

    setStatus(
      `OK.\n` +
      `Itens no resultado: ${mergedRows.length}\n` +
      `Encontrados no KML geral: ${mergedRows.length - missing}\n` +
      `Sem coordenadas no KML geral: ${missing}\n\n` +
      `Baixando KML...`
    );

    downloadText(kml, "resultado_reiteradas_inspecao.kml", "application/vnd.google-earth.kml+xml");
  } catch (e) {
    console.error(e);
    setStatus(`Erro: ${e?.message || e}`);
  }
});
