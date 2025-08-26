// --- Spiele laden und anzeigen ---
async function loadExcelData() {
  try {
    const url = 'https://docs.google.com/spreadsheets/d/138LmOzYFQ0pc1KX-opAFtkQdbYrHsXJC/export?format=xlsx&gid=1986844004';
    const response = await fetch(url);
    if (!response.ok) throw new Error("HTTP_ERROR");

    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array' });

    const firstSheet = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheet];
    let jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false });

    // Nach Datum+Zeit sortieren
    jsonData.sort((a, b) => {
      const dateA = parseDateTime(a["Datum"], a["Zeit"]);
      const dateB = parseDateTime(b["Datum"], b["Zeit"]);
      return (dateA?.getTime() || 0) - (dateB?.getTime() || 0);
    });

    // Filter: nur zukünftige Spiele (ab jetzt minus 2 Stunden) und maximal 3 Tage in der Zukunft
    const now = new Date();
    now.setHours(now.getHours() - 2); // 2 Stunden zurück

    const maxDate = new Date(now);
    maxDate.setDate(maxDate.getDate() + 3); // 3 Tage in der Zukunft

    jsonData = jsonData.filter(spiel => {
      const dt = parseDateTime(spiel["Datum"], spiel["Zeit"]);
      return dt && dt >= now && dt <= maxDate;
    });

    displaySpieleNeu(jsonData);

  } catch (error) {
    console.error("Fehler beim Laden der Excel-Datei:", error);
    const container = document.getElementById("spiele-container");
    if (error.message === "HTTP_ERROR") {
      container.innerHTML = `<p style="color:red;">Daten konnten nicht geladen werden.</p>`;
    } else {
      container.innerHTML = `<p style="color:red;">Keine Daten verfügbar.</p>`;
    }
  }
}

// Hilfsfunktion: Datum parsen (für Strings oder Excel-Datumszahlen)
function parseDate(datum) {
  if (datum instanceof Date && !isNaN(datum)) return datum;

  if (typeof datum === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const ms = datum * 24 * 60 * 60 * 1000;
    const date = new Date(excelEpoch.getTime() + ms);
    if (!isNaN(date)) return date;
  }

  if (typeof datum === "string") {
    if (/^\d+$/.test(datum.trim())) {
      const numDatum = parseInt(datum.trim(), 10);
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const ms = numDatum * 24 * 60 * 60 * 1000;
      const date = new Date(excelEpoch.getTime() + ms);
      if (!isNaN(date)) return date;
    }

    const parts = datum.split(".");
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      const year = parseInt(parts[2], 10);
      const dateObj = new Date(year, month, day);
      if (!isNaN(dateObj)) return dateObj;
    }

    const isoDate = new Date(datum);
    if (!isNaN(isoDate)) return isoDate;
  }

  return null;
}

// Hilfsfunktion: Datum + Zeit parsen
function parseDateTime(datum, zeit) {
  const dateObj = parseDate(datum);
  if (!dateObj) return null;

  if (zeit) {
    // Zeit parsen: sowohl HH:mm als auch HH.mm
    const match = zeit.match(/(\d{1,2})[:.](\d{2})/);
    if (match) {
      dateObj.setHours(parseInt(match[1], 10));
      dateObj.setMinutes(parseInt(match[2], 10));
      dateObj.setSeconds(0);
      dateObj.setMilliseconds(0);
    }
  }
  return dateObj;
}

// Tabelle für Spiele anzeigen
function displaySpieleNeu(data) {
  const container = document.getElementById("spiele-container");
  container.innerHTML = "";

  if (!data || data.length === 0) {
    container.textContent = "Keine Spieldaten verfügbar.";
    return;
  }

  const table = document.createElement("table");
  table.classList.add("spiele-tabelle");

  const headerNames = [
    "Team", "Tag", "Datum", "Zeit", "Paarung", "Kabine Heim", "Kabine Gast", "Schiedsrichter", "Platz"
  ];

  const thead = document.createElement("thead");
  const trHead = document.createElement("tr");
  headerNames.forEach(hdr => {
    const th = document.createElement("th");
    th.textContent = hdr;
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");

  data.forEach(spiel => {
    const tr = document.createElement("tr");

    const tdTeam = document.createElement("td");
    tdTeam.textContent = spiel["Team"] || "-";
    tr.appendChild(tdTeam);

    const tdTag = document.createElement("td");
    tdTag.textContent = spiel["Tag"] || "-";
    tr.appendChild(tdTag);

    const tdDatum = document.createElement("td");
    const datumObj = parseDate(spiel["Datum"]);
    tdDatum.textContent = datumObj ? datumObj.toLocaleDateString('de-CH') : "-";
    tr.appendChild(tdDatum);

    const tdZeit = document.createElement("td");
    tdZeit.textContent = spiel["Zeit"] || "-";
    tr.appendChild(tdZeit);

    const tdPaarung = document.createElement("td");
    tdPaarung.textContent = `${spiel["Paarung"] || ""} - ${spiel["__EMPTY_2"] || ""}`;
    tr.appendChild(tdPaarung);

    const tdKabineHeim = document.createElement("td");
    tdKabineHeim.textContent = spiel["Kabine"] || "-";
    tr.appendChild(tdKabineHeim);

    const tdKabineGast = document.createElement("td");
    tdKabineGast.textContent = spiel["Kabine_1"] || "-";
    tr.appendChild(tdKabineGast);

    const tdSchiri = document.createElement("td");
    tdSchiri.textContent = spiel["Spielleiter/ Schiedsrichter"] || "-";
    tr.appendChild(tdSchiri);

    const tdPlatz = document.createElement("td");
    tdPlatz.textContent = spiel["Platz"] || "-";
    tr.appendChild(tdPlatz);

    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  container.appendChild(table);
}

// --- Platzstatus laden und anzeigen ---
async function loadStatusData() {
  try {
    const url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTrWtthLNjr6G450ajFkDk9-0Y3i8NESN0T-KXk7Lh1VVF1C51Z8wDcy7fgPJQhfy1tMfBNBzcTNz1q/pub?gid=0&single=true&output=csv';
    const response = await fetch(url + `&t=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) throw new Error("HTTP_ERROR");

    // Wichtig: als Text einlesen (UTF-8)
    const text = await response.text();

    // CSV mit XLSX parsen (als String)
    const workbook = XLSX.read(text, { type: 'string' });

    const firstSheet = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheet];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false });

    displayStatus(jsonData);

  } catch (error) {
    console.error("Platzstatus konnte nicht geladen werden:", error);
    const container = document.getElementById("status-container");
    if (error.message === "HTTP_ERROR") {
      container.innerHTML = `<p style="color:red;">Daten konnten nicht geladen werden.</p>`;
    } else {
      container.innerHTML = `<p style="color:red;">Keine Daten verfügbar.</p>`;
    }
  }
}

function displayStatus(data) {
  const container = document.getElementById("status-container");
  container.innerHTML = "";

  if (!data || data.length === 0) {
    container.textContent = "Keine Platzstatus-Daten verfügbar.";
    return;
  }

  data.forEach(entry => {
    const box = document.createElement("div");
    box.classList.add("status-box");

    const platz = entry["Platz"] || "Unbekannter Platz";
    const status = (entry["Status"] || "").toLowerCase();

    if (status === "offen") {
      box.classList.add("offen");
    } else if (status === "nur 1. mannschaft") {
      box.classList.add("nur-eins");
    } else {
      box.classList.add("gesperrt");
    }

    box.innerHTML = `
      <div class="platz-name">${platz}</div>
      <div class="status-text">${
        status === "offen" ? "✅ Offen" :
        status === "nur 1. mannschaft" ? "⚠️ Nur 1. Mannschaft" :
        "🔒 Gesperrt"
      }</div>
    `;

    container.appendChild(box);
  });
}

// Letztes Update anzeigen
function showLastUpdated() {
  const footer = document.getElementById("last-updated");
  const now = new Date();
  footer.textContent = `Zuletzt aktualisiert: ${now.toLocaleString('de-CH')}`;
}

// Start bei Seitenaufruf
window.addEventListener("load", () => {
  loadExcelData();
  loadStatusData();
  showLastUpdated();
});

// Alle 5 Minuten komplette Seite neu laden (ideal im Kiosk-Modus)
setTimeout(() => {
  location.reload();
}, 5 * 60 * 1000); // nach 5 Minuten Seite neu laden
