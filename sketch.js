// sketch.js - versione aggiornata: centro Italia boxes stessa dimensione; slider/pannello centrato in alto
// Assicurati: data/manifest.json esiste e contiene la proprietà "files" con i nomi esatti dei .xlsx

let manifest = [];
let datasets = []; // array di oggetti aggregati per anno: { year, files, aggByRegion }
let activeIndex = 0;
const DATA_FOLDER = "data/"; // NON cambiare

// ordine regioni come richiesto
const TOP_ROW = ["Valle d'Aosta","Piemonte","Liguria","Lombardia","Trentino-Alto Adige","Veneto","Friuli-Venezia Giulia","Emilia-Romagna"];
const MID_ROW = ["Lazio","Marche","Toscana","Umbria"];
const BOTTOM_ROW = ["Abruzzo","Basilicata","Calabria","Campania","Molise","Puglia","Sardegna","Sicilia"];
const ALL_REGIONS = [...TOP_ROW, ...MID_ROW, ...BOTTOM_ROW];

let regionBoxes = [];
let loadingStatus = "init";
let tooltip = null;

function setup() {
  createCanvas(windowWidth, windowHeight);
  textFont('Helvetica, Arial, sans-serif');
  rectMode(CORNER);

  // Centra il pannello dei controlli in alto (evita sovrapposizioni)
  const controls = document.getElementById("controls");
  if (controls) {
    controls.style.left = "50%";
    controls.style.top = "12px";
    controls.style.transform = "translateX(-50%)";
    controls.style.width = "640px"; // larghezza comoda centrale
    controls.style.pointerEvents = "auto";
  }
  // Riduci larghezza legend se esiste per non sovrapporre
  const legend = document.getElementById("legend");
  if (legend) {
    legend.style.top = "76px";
  }

  // slider binding (assume slider esiste in index.html)
  const slider = document.getElementById("slider");
  if (slider) {
    slider.addEventListener("input", () => {
      activeIndex = parseInt(slider.value);
      updateYearLabel();
    });
  }

  // Carica manifest -> carica file -> aggrega
  startLoadingSequence();
}

async function startLoadingSequence() {
  setStatus("Caricamento manifest.json...");
  manifest = await fetchManifest();
  if (!manifest || manifest.length === 0) {
    setStatus("Errore: manifest.json mancante o vuoto nella cartella 'data'.");
    console.error("Manifest", manifest);
    return;
  }
  document.getElementById("fileCount").innerText = manifest.length;
  setStatus("Manifest caricato: " + manifest.length + " file. Avvio parsing XLSX...");
  await loadAllFilesAndAggregate();
  setStatus("Pronto. Usa lo slider per cambiare referendum.");
  populateSlider();
}

async function fetchManifest() {
  try {
    const resp = await fetch(DATA_FOLDER + "manifest.json");
    if (!resp.ok) {
      console.error("fetch manifest error", resp.status);
      return [];
    }
    const json = await resp.json();
    if (Array.isArray(json)) return json;
    if (json && Array.isArray(json.files)) return json.files;
    console.warn("manifest.json non nel formato atteso");
    return [];
  } catch (err) {
    console.error("Errore fetch manifest", err);
    return [];
  }
}

function setStatus(msg) {
  loadingStatus = msg;
  const el = document.getElementById("status");
  if (el) el.innerText = "Stato: " + msg;
}

async function loadAllFilesAndAggregate() {
  const loaded = [];
  for (let i = 0; i < manifest.length; i++) {
    const fname = manifest[i];
    setStatus(`Caricamento ${fname} (${i+1}/${manifest.length})`);
    try {
      const url = DATA_FOLDER + encodeURIComponent(fname);
      const resp = await fetch(url);
      if (!resp.ok) {
        console.warn("File non trovato o errore:", fname, resp.status, resp.statusText);
        continue;
      }
      const ab = await resp.arrayBuffer();
      const data = new Uint8Array(ab);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rawJson = XLSX.utils.sheet_to_json(sheet, { defval: null });
      const year = extractYearFromName(fname);
      loaded.push({ filename: fname, year: year, raw: rawJson });
    } catch (err) {
      console.error("Errore caricamento file", fname, err);
    }
  }

  if (loaded.length === 0) {
    setStatus("Nessun dataset valido caricato.");
    return;
  }

  // Raggruppa per anno (più file nello stesso anno -> verranno uniti)
  const byYear = {};
  loaded.forEach(item => {
    const ky = item.year || ("file_" + item.filename);
    if (!byYear[ky]) byYear[ky] = [];
    byYear[ky].push(item);
  });

  const combined = [];
  Object.keys(byYear).sort((a,b)=>{
    if (typeof a === 'number' && typeof b === 'number') return a-b;
    return (""+a).localeCompare(""+b);
  }).forEach(ky => {
    const items = byYear[ky];
    let allRows = [];
    items.forEach(it => { allRows = allRows.concat(it.raw); });
    const agg = aggregateRowsByRegion(allRows);
    combined.push({ year: ky, files: items.map(i=>i.filename), rowsCount: allRows.length, aggByRegion: agg });
  });

  datasets = combined;
}

function extractYearFromName(name) {
  const m = name.match(/(19|20)\d{2}/);
  if (m) return parseInt(m[0]);
  return null;
}

function canonicalizeColName(s) {
  if (!s && s !== 0) return "";
  return (""+s).toLowerCase().replace(/[\s\-_\.]/g, "");
}

function parseNumber(v) {
  if (v == null) return NaN;
  if (typeof v === 'number') return v;
  const s = (""+v).replace(/\./g,"").replace(/,/g,"").trim();
  const n = Number(s);
  return isNaN(n) ? NaN : n;
}

function aggregateRowsByRegion(rows) {
  const regionMap = {};
  rows.forEach(row => {
    const normalized = {};
    for (const k of Object.keys(row)) normalized[canonicalizeColName(k)] = row[k];

    const reg = normalized['regione'] || normalized['region'] || normalized['regionename'] || normalized['reg'] || normalized['regioni'];
    const elettori = parseNumber(normalized['elettori'] || normalized['elettoritot'] || normalized['elettoritotali'] || normalized['elettori_tot']);
    const votanti = parseNumber(normalized['votanti'] || normalized['numvotanti'] || normalized['votantitot']|| normalized['votanti_tot']);

    if (!reg) return;
    const rKey = String(reg).trim();
    if (!regionMap[rKey]) regionMap[rKey] = { elettori: 0, votanti: 0, provinces: 0 };

    if (!isNaN(elettori) && !isNaN(votanti) && elettori > 0) {
      regionMap[rKey].elettori += elettori;
      regionMap[rKey].votanti += votanti;
      regionMap[rKey].provinces += 1;
    } else {
      // skip rows senza numeri utili
    }
  });

  const result = {};
  for (const r of Object.keys(regionMap)) {
    const obj = regionMap[r];
    const aff = (obj.elettori > 0) ? (obj.votanti / obj.elettori) : null;
    result[r] = { elettori: obj.elettori, votanti: obj.votanti, affluenza: aff, provinces: obj.provinces };
  }
  return result;
}

function populateSlider() {
  const slider = document.getElementById("slider");
  if (!slider) return;
  slider.min = 0;
  slider.max = datasets.length - 1;
  slider.value = 0;
  activeIndex = 0;
  document.getElementById("fileCount").innerText = manifest.length || 0;
  updateYearLabel();
}

function updateYearLabel() {
  const lab = document.getElementById("yearLabel");
  if (!lab) return;
  const d = datasets[activeIndex];
  if (!d) { lab.innerText = "—"; return; }
  lab.innerText = "" + d.year + " (" + d.files.join(", ") + ")";
}

function draw() {
  background(250);
  fill(30);
  noStroke();
  textSize(14);
  textAlign(LEFT, TOP);
  

  // layout calcoli
  const sidePadding = 40;
  const usableW = width - sidePadding*2;
  const rowHeight = (height - 160) / 3; // lasciamo spazio per controls al top

  // calcolo boxW di riferimento (basato su 8 box) - questo assicura che il centro abbia stessa dimensione
  const boxGap = 12;
  const boxW_ref = (usableW - (boxGap * (8 - 1))) / 8; // dimensione base per 8 colonne

  // Top row: 8 boxes (usa boxW_ref)
  drawRowBoxes(TOP_ROW, sidePadding, 110, usableW, rowHeight*0.95, false, boxW_ref);

  // Middle row: 4 boxes - centrate, ma usa boxW_ref per avere la stessa dimensione
  drawRowBoxes(MID_ROW, sidePadding, 110 + rowHeight, usableW, rowHeight*0.95, true, boxW_ref);

  // Bottom row: 8 boxes
  drawRowBoxes(BOTTOM_ROW, sidePadding, 110 + rowHeight*2, usableW, rowHeight*0.95, false, boxW_ref);

  if (tooltip) drawTooltip(tooltip);

  fill(80);
  noStroke();
  textSize(12);
  textAlign(LEFT,BOTTOM);
  text(loadingStatus, 12, height - 8);
}

function drawRowBoxes(regions, x, y, w, h, center=false, forcedBoxW=null) {
  const count = regions.length;
  const boxGap = 12;
  // totalGap per il conteggio effettivo della riga
  const totalGap = boxGap * (count - 1);

  // se viene passato forcedBoxW (es: per mantenere la stessa dimensione centrale),
  // lo usiamo e calcoliamo startX per centrare il gruppo
  let boxW;
  if (forcedBoxW && typeof forcedBoxW === 'number') {
    boxW = forcedBoxW;
  } else {
    boxW = (w - totalGap) / count;
  }

  // limit boxW se è troppo grande per lo spazio (prevenzione)
  const maxBoxW = (w - totalGap) / Math.max(count,1);
  if (boxW > maxBoxW) boxW = maxBoxW;

  const boxH = Math.min(boxW, h * 0.95);
  let startX = x;
  if (center) {
    const totalWidth = count * boxW + totalGap;
    startX = x + (w - totalWidth)/2;
  }

  // pick dataset per activeIndex
  const d = datasets[activeIndex];
  // reset regionBoxes se è la prima riga disegnata dall'alto
  if (y <= 200) regionBoxes = [];

  for (let i=0;i<count;i++) {
    const rx = startX + i*(boxW + boxGap);
    const ry = y;
    const regionName = regions[i];
    fill(230);
    stroke(200);
    rect(rx, ry, boxW, boxH, 6);

    let aff = null;
    if (d && d.aggByRegion) {
      if (d.aggByRegion[regionName] && d.aggByRegion[regionName].affluenza != null) {
        aff = d.aggByRegion[regionName].affluenza;
      } else {
        const keys = Object.keys(d.aggByRegion);
        for (const k of keys) {
          if (k && k.toLowerCase() === regionName.toLowerCase()) { aff = d.aggByRegion[k].affluenza; break; }
        }
      }
    }

    if (aff != null && !isNaN(aff)) {
      const pct = constrain(aff, 0, 1);
      const fillH = boxH * pct;
      noStroke();
      fill(70,140,220);
      rect(rx, ry + boxH - fillH, boxW, fillH, 6);
      fill(20);
      textSize(12);
      textAlign(CENTER, BOTTOM);
      text(nf(pct*100,1,1) + "%", rx + boxW/2, ry + boxH - fillH - 6);
    } else {
      fill(200);
      textSize(11);
      textAlign(CENTER,CENTER);
      fill(120);
      text("N/D", rx + boxW/2, ry + boxH/2);
    }

    fill(30);
    textSize(12);
    textAlign(CENTER, TOP);
    text(regionName, rx + boxW/2, ry + boxH + 6);

    regionBoxes.push({ x: rx, y: ry, w: boxW, h: boxH, region: regionName, aff: aff, yearObj: d });
  }
}

function mouseMoved() {
  tooltip = null;
  for (const rb of regionBoxes) {
    if (!rb) continue;
    if (mouseX >= rb.x && mouseX <= rb.x + rb.w && mouseY >= rb.y && mouseY <= rb.y + rb.h) {
      tooltip = { x: mouseX, y: mouseY, region: rb.region, aff: rb.aff, year: rb.yearObj ? rb.yearObj.year : null, files: rb.yearObj ? rb.yearObj.files : [] };
      document.getElementById("hoverInfo").innerText = (rb.aff!=null ? nf(rb.aff*100,1,1)+"%" : "N/D") + " — " + rb.region;
      return;
    }
  }
  document.getElementById("hoverInfo").innerText = "–";
}

function drawTooltip(t) {
  if (!t) return;
  push();
  const tw = 300;
  const th = 70;
  let tx = t.x + 14;
  let ty = t.y + 14;
  if (tx + tw > width) tx = t.x - tw - 14;
  if (ty + th > height) ty = t.y - th - 14;
  fill(255);
  stroke(140);
  rect(tx, ty, tw, th, 8);
  noStroke();
  fill(30);
  textSize(13);
  textAlign(LEFT, TOP);
  text("Regione: " + t.region, tx + 10, ty + 8);
  text("Affluenza: " + (t.aff != null ? nf(t.aff*100,1,1) + "%" : "N/D"), tx + 10, ty + 26);
  text("Anno/files: " + (t.year ? (t.year + " — " + t.files.join("; ")) : "—"), tx + 10, ty + 44);
  pop();
}
