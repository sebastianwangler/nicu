// Apps Script Web App für Les Marmottes. Läuft im Google-Konto
// lesmarmottesb@gmail.com (dort in script.google.com einfügen + als Web App
// deployen). Legt bei einer Anfrage automatisch das Sheet "Les Marmottes –
// Anfragen" an, falls es noch nicht existiert.
//
// Ablauf:
//   doPost   <- Website sendet neue Anfrage
//              -> ANGEFRAGT-Eintrag im öffentlichen Kalender
//              -> Zeile im Sheet
//              -> Mail an den Verwalter mit "Zusagen"/"Ablehnen"-Links
//   doGet    <- Verwalter klickt einen der beiden Links in der Mail
//     ablehnen  -> öffnet vorbereitete Antwort-Mail an den Gast
//     zusagen   -> Eintrag im privaten Kalender, ANGEFRAGT -> BELEGT im
//                  öffentlichen Kalender, öffnet vorbereitete Bestätigungs-Mail

const OEFFENTLICHER_KALENDER_ID =
  "4272fb587c47328f819830065cf06566108ca68f64dcf0a6a89f6fd6f399b933@group.calendar.google.com";
const VERWALTER_EMAIL = "lesmarmottesb@gmail.com";
const HAUS_NAME = "Les Marmottes";
const SHEET_DATEINAME = "Les Marmottes – Anfragen";
const BLATT_NAME = "Anfragen";
const SHEET_SPALTEN = [
  "id", "timestamp", "haus", "name", "email", "telefon", "von", "bis",
  "erwachsene", "kinder", "tiere", "tierart", "status", "eventIdOeffentlich"
];

function getOeffentlicherKalender() {
  return CalendarApp.getCalendarById(OEFFENTLICHER_KALENDER_ID);
}

// Der private Kalender "Buchungen"/Admin-Kalender = der normale Standard-
// kalender dieses Google-Kontos. Falls stattdessen ein separat angelegter
// Kalender genutzt werden soll: hier durch
// CalendarApp.getCalendarById("...") ersetzen.
function getPrivaterKalender() {
  return CalendarApp.getDefaultCalendar();
}

function getSheet() {
  const dateien = DriveApp.getFilesByName(SHEET_DATEINAME);
  let ss;
  if (dateien.hasNext()) {
    ss = SpreadsheetApp.open(dateien.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_DATEINAME);
    const blatt = ss.getSheets()[0];
    blatt.setName(BLATT_NAME);
    blatt.appendRow(SHEET_SPALTEN);
  }
  return ss.getSheetByName(BLATT_NAME) || ss.getSheets()[0];
}

function findeZeile(blatt, id) {
  const daten = blatt.getDataRange().getValues();
  for (let i = 1; i < daten.length; i++) {
    if (daten[i][0] === id) return i + 1; // Sheet-Zeilen sind 1-basiert
  }
  return -1;
}

// Deutsches Datumsformat für Betreffzeilen, passend zur Website.
function formatiereDatum(isoDatum) {
  const [j, m, t] = isoDatum.split("-");
  return `${t}.${m}.${j}`;
}

// Google: end.date bei ganztägigen Events ist exklusiv (Tag NACH dem
// letzten belegten Tag). Unser "bis" von der Website ist inklusive.
function alsGanztagesEnde(bisISO) {
  const datum = new Date(bisISO + "T00:00:00");
  datum.setDate(datum.getDate() + 1);
  return Utilities.formatDate(datum, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function doPost(e) {
  const daten = JSON.parse(e.postData.contents);
  const { name, email, telefon, von, bis, erwachsene, kinder, tiere, tierart } = daten;

  const oeffentlich = getOeffentlicherKalender();

  // Serverseitig nochmal prüfen, falls zwei Gäste gleichzeitig anfragen.
  const kollisionen = oeffentlich.getEvents(
    new Date(von + "T00:00:00"),
    new Date(alsGanztagesEnde(bis) + "T00:00:00")
  );
  if (kollisionen.length > 0) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, fehler: "belegt" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const id = Utilities.getUuid();
  const angefragtEvent = oeffentlich.createAllDayEvent(
    `${name} (ANGEFRAGT)`,
    new Date(von + "T00:00:00"),
    new Date(alsGanztagesEnde(bis) + "T00:00:00")
  );

  const blatt = getSheet();
  blatt.appendRow([
    id, new Date(), "haus1", name, email, telefon, von, bis,
    erwachsene, kinder, tiere, tierart || "", "angefragt", angefragtEvent.getId()
  ]);

  const webAppUrl = ScriptApp.getService().getUrl();
  const zusagenLink = `${webAppUrl}?action=zusagen&id=${encodeURIComponent(id)}`;
  const ablehnenLink = `${webAppUrl}?action=ablehnen&id=${encodeURIComponent(id)}`;

  const htmlBody = `
    <p>Neue Reservationsanfrage für ${HAUS_NAME}:</p>
    <ul>
      <li><strong>Name:</strong> ${name}</li>
      <li><strong>E-Mail:</strong> ${email}</li>
      <li><strong>Telefon:</strong> ${telefon}</li>
      <li><strong>Zeitspanne:</strong> ${formatiereDatum(von)} – ${formatiereDatum(bis)}</li>
      <li><strong>Erwachsene:</strong> ${erwachsene}, <strong>Kinder:</strong> ${kinder}, <strong>Tiere:</strong> ${tiere}${tierart ? " (" + tierart + ")" : ""}</li>
    </ul>
    <p>
      <a href="${zusagenLink}" style="display:inline-block;padding:10px 20px;background:#2e7d32;color:#fff;text-decoration:none;border-radius:4px;">Zusagen</a>
      &nbsp;&nbsp;
      <a href="${ablehnenLink}" style="display:inline-block;padding:10px 20px;background:#c62828;color:#fff;text-decoration:none;border-radius:4px;">Ablehnen</a>
    </p>
  `;

  MailApp.sendEmail({
    to: VERWALTER_EMAIL,
    subject: `Reservationsanfrage ${HAUS_NAME} – ${name}, ${formatiereDatum(von)} – ${formatiereDatum(bis)}`,
    htmlBody
  });

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function seitenAusgabe(html) {
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Rendert eine Seite mit Bestätigungstext + Mail-Link. Versucht zusätzlich
// automatisch weiterzuleiten (funktioniert meist, aber nicht garantiert je
// nach Browser/Mail-App) — der sichtbare Button ist deshalb immer da, falls
// die automatische Weiterleitung nicht greift.
function mailSeite(ueberschrift, mailtoLink) {
  return seitenAusgabe(`
    <div style="font-family:sans-serif;max-width:480px;margin:40px auto;text-align:center;">
      <h2>${ueberschrift}</h2>
      <p><a href="${mailtoLink}" style="display:inline-block;padding:12px 24px;background:#b5654a;color:#fff;text-decoration:none;border-radius:6px;font-size:16px;">Antwort-Mail öffnen</a></p>
      <p style="color:#888;font-size:13px;">Öffnet dein Standard-Mailprogramm (Mail, Outlook, Gmail-App) mit vorausgefülltem Empfänger und Betreff.</p>
    </div>
    <script>window.location.href = ${JSON.stringify(mailtoLink)};</script>
  `);
}

function doGet(e) {
  const action = e.parameter.action;
  const id = e.parameter.id;
  const blatt = getSheet();
  const zeile = findeZeile(blatt, id);

  if (zeile === -1) {
    return seitenAusgabe("<p>Anfrage nicht gefunden (evtl. schon bearbeitet).</p>");
  }

  const werte = blatt.getRange(zeile, 1, 1, SHEET_SPALTEN.length).getValues()[0];
  const [, , , name, email, telefon, von, bis, erwachsene, kinder, tiere, tierart, , eventIdOeffentlich] = werte;

  if (action === "ablehnen") {
    blatt.getRange(zeile, 13).setValue("abgelehnt"); // Spalte "status"
    const mailtoLink =
      `mailto:${encodeURIComponent(email)}` +
      `?subject=${encodeURIComponent("Ihre Anfrage " + HAUS_NAME)}` +
      `&body=${encodeURIComponent("Leider keine Kapazität.")}`;
    return mailSeite("Anfrage abgelehnt.", mailtoLink);
  }

  if (action === "zusagen") {
    // Aktion 1: Eintrag im privaten Kalender mit allen Details.
    const privat = getPrivaterKalender();
    const beschreibung =
      `E-Mail: ${email}\nTelefon: ${telefon}\n` +
      `Erwachsene: ${erwachsene}, Kinder: ${kinder}, Tiere: ${tiere}${tierart ? " (" + tierart + ")" : ""}`;
    privat.createAllDayEvent(
      name,
      new Date(von + "T00:00:00"),
      new Date(alsGanztagesEnde(bis) + "T00:00:00"),
      { description: beschreibung }
    );

    // Aktion 2: öffentlicher Kalender ANGEFRAGT -> BELEGT, Name bleibt.
    const oeffentlich = getOeffentlicherKalender();
    const event = oeffentlich.getEventById(eventIdOeffentlich);
    if (event) event.setTitle(`${name} (BELEGT)`);

    blatt.getRange(zeile, 13).setValue("bestätigt"); // Spalte "status"

    // Aktion 3: Bestätigungsmail vorbereiten.
    const betreff = `Bestätigung Reservationsanfrage ${HAUS_NAME} – ${name}, ${formatiereDatum(von)} – ${formatiereDatum(bis)}`;
    const mailtoLink =
      `mailto:${encodeURIComponent(email)}` +
      `?subject=${encodeURIComponent(betreff)}` +
      `&body=${encodeURIComponent("Vielen Dank, hier noch Infos.")}`;
    return mailSeite("Buchung bestätigt.", mailtoLink);
  }

  return seitenAusgabe("<p>Unbekannte Aktion.</p>");
}
