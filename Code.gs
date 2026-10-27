// Plik Code.gs
const CACHE_KEY_LOGINS = 'LOGIN_MAP_CACHE';
const PROP_KEY_SESSIONS = 'activeSessions';
const PROP_KEY_TOP3 = 'TOP_3_CACHE';
function generateAndSendNewPasswordToAdmin() {
  try {
    const newPassword = generateNewPassword(); 
    PropertiesService.getScriptProperties().setProperty('accessPassword', newPassword);
    logDebug(newPassword);
    // Wysyłanie hasła do Google Chat
    const chatApiUrl = 'https://chat.googleapis.com/v1/spaces/AAQA2x_A1_E/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=qra4pAPSR-3maIGLjtc4NbWfCRJ1tScI-HpIbae45Sk';
    logDebug("url");
    const message = {'text': `Nowe hasło do aplikacji: ${newPassword}`};
    logDebug(message);
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(message),
      'muteHttpExceptions': true
    
    };
    logDebug(options);
    const response = UrlFetchApp.fetch(chatApiUrl, options);
    logDebug(response);
    if (response.getResponseCode() === 200) {
      logDebug('Hasło wysłano pomyślnie na Google Chat.');
      return newPassword;
    } else {
      logDebug('Błąd wysyłania hasła na Google Chat.');
      return null;
    }
  } catch (e) {
    logDebug('Krytyczny błąd w generateAndSendNewPasswordToAdmin: ' + e.toString());
    return null;
  }
}
function logDebug(message) {
  const debug = PropertiesService.getScriptProperties().getProperty('DEBUG_ENABLED') === 'true';
  if (debug) {
    console.log(message); 
  }
}
function setDebugMode(status) {
  PropertiesService.getScriptProperties().setProperty('DEBUG_ENABLED', status);
  return "Tryb debugowania ustawiony na: " + status;
}
function autoCleanupSessions() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); 
    
    const props = PropertiesService.getScriptProperties();
    const sessionsData = props.getProperty("activeSessions");
    
    if (!sessionsData) return;
    
    let sessions = JSON.parse(sessionsData);
    const now = Date.now();
    const TWELVE_HOURS = 14 * 60 * 60 * 1000; // 12 godzin w milisekundach

    const initialCount = sessions.length;
    sessions = sessions.filter(s => {
      const startTime = s.startTimestamp || (s.start ? new Date(s.start).getTime() : now);
      return (now - startTime) < TWELVE_HOURS;
    });

    if (sessions.length < initialCount) {
      props.setProperty("activeSessions", JSON.stringify(sessions));
      logDebug(`Auto-cleanup: Usunięto ${initialCount - sessions.length} przedawnionych sesji.`);
    }
    
  } catch (e) {
    logDebug("Błąd podczas auto-cleanupu: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}
function verifyPassword(enteredPassword) {
  const storedPassword = PropertiesService.getScriptProperties().getProperty('accessPassword');
  
  if (storedPassword === enteredPassword) {
    logDebug('Weryfikacja hasła Pomyślna!');
    return 'OK';
  } else {
    logDebug('Weryfikacja hasła NIEUDANA!');
    return 'Błędne hasło.';
  }
}
function getTopOklejanieToday() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calosc");
  if (!sheet) {
    logDebug("Arkusz 'Calosc' nie został znaleziony.");
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) { 
    logDebug("Brak danych w arkuszu 'Calosc'.");
    return [];
  }

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayFormatted = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const userTotalsToday = {};
  const userTotalsMonth = {};

  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Sprawdzamy czy wiersz ma kolumnę M (indeks 12)
    if (row.length < 13 || !row[2]) continue;

    const recordDate = new Date(row[2]);
    const ilosc = parseFloat(row[5]) || 0;
    if (ilosc === 0) continue;

    const procedureWorkMode = row[9];
    
    // ZMIANA: Zamiast row[1] (login), pobieramy row[12] (Kolumna M - Identyfikator)
    const userIdentifier = row[12] || "Brak ID";

    const dateRecord = Utilities.formatDate(recordDate, tz, "yyyy-MM-dd");
    
    // Logika sumowania - używamy identyfikatora jako klucza
    if (dateRecord === todayFormatted && procedureWorkMode === "S_OKL") {
      userTotalsToday[userIdentifier] = (userTotalsToday[userIdentifier] || 0) + ilosc;
    }

    if (recordDate.getFullYear() === currentYear && recordDate.getMonth() === currentMonth && procedureWorkMode === "S_OKL") {
      userTotalsMonth[userIdentifier] = (userTotalsMonth[userIdentifier] || 0) + ilosc;
    }
  }

  // Budowanie rankingu Top 3
  let top3 = Object.entries(userTotalsToday)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([user, qty]) => ({ user, qty: Math.round(qty * 100) / 100 }));

  if (top3.length === 0) {
    logDebug("Brak danych dziennych. Używam danych miesięcznych.");
    top3 = Object.entries(userTotalsMonth)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([user, qty]) => ({ user, qty: Math.round(qty * 100) / 100 }));
  }

  logDebug("Top 3 Oklejanie (Anonimizowane): " + JSON.stringify(top3));
  return top3;
}
function generateNewPassword() {
  const length = 8;
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}
function sendNewPasswordToAdmin() {
  return generateAndSendNewPasswordToAdmin();
}
function doGet() {
  const newPassword = generateAndSendNewPasswordToAdmin();
  
  if (newPassword) {
    logDebug('Nowe hasło wygenerowane i wysłane pomyślnie. Zwracam stronę.');
  } else {
    logDebug('BŁĄD: Nie udało się wygenerować lub wysłać hasła.');
  }
  
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("FM Logistic – Rejestracja procesu")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getLoginMapCached() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY_LOGINS);
  if (cached) return JSON.parse(cached);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Login");
  if (!sheet) return {};
  
  // Batching: Pobieramy wszystko jednym strzałem
  const data = sheet.getRange("A1:B" + sheet.getLastRow()).getValues();
  const map = {};
  data.forEach(row => {
    if (row[0] && row[1]) map[row[0].toString().trim().toLowerCase()] = row[1].toString().trim();
  });

  cache.put(CACHE_KEY_LOGINS, JSON.stringify(map), 21600); // Cache na 6h
  return map;
}
function forceClearCache() {
  //const cache = CacheService.getScriptCache();
  cache.remove('CACHE_KEY_LOGINS');
  console.log("Pamięć podręczna (Cache) loginu została wyczyszczona.");
}
function updateTopStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Calosc");
  const targetSheet = ss.getSheetByName("Dashboard_Data") || ss.insertSheet("Dashboard_Data");

  // 1. Wykonaj ciężkie obliczenia (Twoja obecna logika z getTopOklejanieToday)
  const top3 = getTopOklejanieToday(); 

  // 2. Wyczyść stare wyniki i zapisz nowe
  targetSheet.clear();
  if (top3.length > 0) {
    const rows = top3.map(item => [item.user, item.qty]);
    targetSheet.getRange(1, 1, rows.length, 2).setValues(rows);
  }
  
  logDebug("Statystyki Top 3 zostały zaktualizowane: " + new Date());
}
function resolveLogin(login) {
  const map = getLoginMapCached();
  const cleanLogin = login.toString().trim().toLowerCase();
  return map[cleanLogin] || login;
}
function getUsersAndLoginMap() {
  return {
    users: getActiveSessions(),
    loginMap: getLoginMapCached()
  };
}
function getTopOklejanieFromCache() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard_Data");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  // Mapujemy prosto z gotowej tabelki
  return data.map(row => ({ user: row[0], qty: row[1] }));
}

function getActiveSessions() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const data = scriptProperties.getProperty("activeSessions");
  try {
    return data ? JSON.parse(data) : [];
  } catch (e) {
    logDebug("Błąd podczas parsowania aktywnych sesji: " + e.toString() + ". Zwracam pustą tablicę.");
    return []; 
  }
}
function setActiveSessions(sessions) {
  try {
    PropertiesService.getScriptProperties().setProperty("activeSessions", JSON.stringify(sessions));
  } catch (e) {
    logDebug("Błąd podczas zapisywania aktywnych sesji: " + e.toString());
    
  }
}
function checkIfUserIsActive(identifier) {
  logDebug("checkIfUserIsActive wywołana z identyfikatorem: " + identifier);
  const sessions = getActiveSessions();
  const isActive = sessions.some(s => 
    (s.loginMask && s.loginMask.toString().trim().toLowerCase() === identifier.toString().trim().toLowerCase()) ||
    (s.login && s.login.toString().trim().toLowerCase() === identifier.toString().trim().toLowerCase())
  );
  
  if (isActive) {
    logDebug("Użytkownik " + identifier + " JEST aktywny.");
    
    return { isActive: true, message: `Użytkownik ${identifier} ma już aktywną sesję.` };
  } else {
    logDebug("Użytkownik " + identifier + " NIE JEST aktywny.");
    return { isActive: false }; 
  }
}
function addToActive(session) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); // Czekaj max 5s na zwolnienie blokady
    
    const sessions = getActiveSessions();
    const alreadyActive = sessions.some(s => s.login === session.login);
    if (session.procedure === "Anomalia") {
      sendAnomalyNotification(session);
    }
  
    if (alreadyActive) {
      throw new Error(`Użytkownik ${session.login} jest już aktywny.`);
    }

    sessions.push(session);
    setActiveSessions(sessions);
  } finally {
    lock.releaseLock();
  }
}
function sendAnomalyNotification(data) {
  
  const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAQA2x_A1_E/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=HY8HgJGBqoKXwaMpfFGsf38ssp4IhShAG-bHHiJI-08"; 

  const maskedUser = data.loginMask || data.login;
  const timeFormatted = Utilities.formatDate(new Date(), "GMT+2", "HH:mm:ss");

  
  const payload = {
    "text": "⚠️ *ALERT: ZGŁOSZONO ANOMALIĘ!*\n\n" +
            "👤 *Pracownik:* " + maskedUser + "\n" +
            "⏰ *Czas rozpoczęcia:* " + timeFormatted + "\n" +
            "📌 *Proces:* Anomalia"
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    UrlFetchApp.fetch(WEBHOOK_URL, options);
  } catch (error) {
    Logger.log("Błąd wysyłania wiadomości na czat: " + error.toString());
  }
}
function removeFromActive(login, procedure) {
  const lock = LockService.getScriptLock();
  try {
    // Czekaj do 5 sekund na dostęp, jeśli inny proces aktualnie zapisuje dane
    lock.waitLock(5000); 

    let sessions = getActiveSessions(); // Pobiera aktualną listę
    const initialCount = sessions.length;
    
    // Filtrowanie sesji w pamięci RAM (bardzo szybkie)
    sessions = sessions.filter(s => 
      !(s.login && s.login.toString().trim().toLowerCase() === login.toString().trim().toLowerCase() && 
        s.procedure && s.procedure.toString().trim() === procedure.toString().trim())
    );
    
    if (sessions.length < initialCount) {
      setActiveSessions(sessions); // Zapisuje zaktualizowaną listę
      logDebug(`Usunięto sesję dla: ${login}`);
    }
  } catch (e) {
    logDebug(`Błąd w removeFromActive: ${e.toString()}`);
    throw new Error("Serwer jest zajęty, spróbuj zakończyć proces ponownie za chwilę.");
  } finally {
    lock.releaseLock(); // Zawsze zwalniaj blokadę dla innych użytkowników
  }
}
function clearAllActiveSessions() {
  PropertiesService.getScriptProperties().deleteProperty("activeSessions");
  logDebug("Wszystkie aktywne sesje zostały wyczyszczone.");
}
function fetchBigQueryData(sscc) {
  const PROJECT_ID = 'fml-prod-pl-bo';
  
  const ssccToSearch = sscc ? sscc.substring(2) : null; 
  if (!ssccToSearch) return null;

  // ZAPYTANIE: OrderNumber został usunięty
  const QUERY = `
    SELECT
        t1.ItemCode,
        t1.QtyAfter / t2.NbUnitPerParcel AS Kartonow
    FROM
        \`fml-prod-pl-bo.WAREHOUSING.IPGMove\` AS t1
    LEFT JOIN
        \`fml-prod-pl-bo.WAREHOUSING.Item\` AS t2
        ON t1.ItemCode = t2.ItemCode AND t2.ActivityCode = 'MTE'
    WHERE
        t1.ActivityCode = 'MTE' AND t1.QtyAfter <> 0
        AND (t1.HDNumber = '${ssccToSearch}' OR SUBSTR(t1.HDNumber, 3) = '${ssccToSearch}')
    LIMIT 1`; 

  try {
    const request = {
      query: QUERY,
      useLegacySql: false
    };

    const queryResults = BigQuery.Jobs.query(request, PROJECT_ID);

    if (queryResults.jobComplete && queryResults.rows && queryResults.rows.length > 0) {
      const row = queryResults.rows[0];
      
      // 💡 KOREKTA INDEKSOWANIA:
      const ItemCode = row.f[0].v; // Nadal indeks [0]
      const Kartonow = row.f[1].v; // ZMIENIONY indeks na [1]! (był [2])
      
      return {
        // K: REF/BO (teraz tylko ItemCode)
        ref_bo: ItemCode || "", // Dodaj || "" na wypadek NULL
        // L: PLAN_QTY/COP_QTY (Kartonow)
        qty: Kartonow ? Math.round(parseFloat(Kartonow)) : 0 
      };
    }
    
    logDebug("Brak danych z BigQuery dla SSCC: " + sscc);
    return null;

  } catch (e) {
    logDebug("Błąd zapytania BigQuery: " + e.toString());
    return null;
  }
}  
function saveProcessData(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calosc") || SpreadsheetApp.getActiveSpreadspreadsheet().insertSheet("Calosc");
  if (sheet.getLastRow() === 0) { 
    const headers = ["SSCC/ID_PROC", "PRACOWNIK", "DATA", "START", "STOP", "ILOSC", "CZAS_TRWANIA", "SUMA_ILOSCI", "ILOSC_LUDZI", "TRYB_PRACY", "REF/BO", "PLAN_QTY/COP_QTY", "LOGIN_RAW"];
    sheet.appendRow(headers);
    sheet.getRange("A1:M1").setFontWeight("bold").setHorizontalAlignment("center");
    SpreadsheetApp.flush();
  }

  // --- Przygotowanie danych (bez zmian) ---
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const dateStr = data.start ? Utilities.formatDate(new Date(data.start), tz, "yyyy-MM-dd") : Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const startStr = data.start ? Utilities.formatDate(new Date(data.start), tz, "HH:mm:ss") : "";
  const stopStr = data.stop ? Utilities.formatDate(new Date(data.stop), tz, "HH:mm:ss") : "";
  
  let ssccFinal;
  if (data.procedure === "Anomalia") {
    ssccFinal = data["Komentarz"] || "";
  } else if (data.procedure === "Copacking") {
    ssccFinal = "Copacking";
  } else if (data.procedure === "Delegacja") {
    ssccFinal = data["Klient"];
  } else {
    ssccFinal = data.sscc || data.procedure; 
  } 

  let duration = "00:00:00";
  if (data.start && data.stop) {
    const durationMS = new Date(data.stop).getTime() - new Date(data.start).getTime();
    if (durationMS >= 0) {
      const hours = Math.floor(durationMS / 3600000);
      const minutes = Math.floor((durationMS % 3600000) / 60000);
      const seconds = Math.floor((durationMS % 60000) / 1000);
      duration = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    }
  }
  
  const iloscDoZapisu = data.ilosc || data["Ilosc_pallet"] || (data.procedure === "Przerwa" || data.procedure === "Delegacja" || data.procedure === "Anomalia" ? "" : "1");

  // --- Nowa Logika BigQuery dla Oklejania ---
  let refBoValue = (data.procedure === "Copacking" ? (data["Referencja"] || "") : "");
  let planQtyValue = (data.procedure === "Copacking" ? (data["Ilosc_kartonow"] || "") : "");

  if (data.procedure === "Oklejanie") {
    // 💡 WYWOŁANIE BIGQUERY
    const bqData = fetchBigQueryData(ssccFinal);

    if (bqData) {
      refBoValue = bqData.ref_bo; // REF/BO (OrderNumber/ItemCode)
      planQtyValue = bqData.qty;  // PLAN_QTY/COP_QTY (Kartonow)
    } else {
      // Jeśli zapytanie się nie powiodło lub nie zwróciło danych,
      // zostawiamy puste lub dodajemy notatkę
      refBoValue = "BRAK DANYCH BQ"; 
      planQtyValue = 0;
    }
  } 

  // --- Zapis wiersza ---
  const rowData = [
    ssccFinal,                   // A
    data.loginMask || data.login || "", // B
    dateStr,                     // C
    startStr,                    // D
    stopStr,                     // E
    iloscDoZapisu,               // F
    duration,                    // G
    data["Ilosc_referencji"] || iloscDoZapisu,               // H
    data["Ilosc_ludzi"] || (data.procedure === "Przerwa" || data.procedure === "Delegacja" ? "" : "1"), // I
    getWorkMode(data.procedure), // J
    refBoValue,                  // K
    planQtyValue,                // L
    data.login || ""             // M
  ];
  sheet.appendRow(rowData);

  logDebug("Dane procesu zapisane: " + JSON.stringify(rowData));
}
function getNameByLogin(login) {
  const map = getLoginMapCached();
  return map[login.toString().trim().toLowerCase()] || login;
}
function getWorkMode(proc) {
  const map = {
    "Oklejanie": "S_OKL",
    "Foliowanie": "S_FOL",
    "Kontener/Bulkzal": "S_KON",
    "Przerwa": "S_PRW",
    "Anomalia": "S_ANO",
    "Copacking": "S_COP",
    "Delegacja": "S_DEL",
    "Instrukcje": "S_INS"
  };
  return map[proc] || proc || "";
}
