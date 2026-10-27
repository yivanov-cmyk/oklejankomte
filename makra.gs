const SHEET_NAME = "Login";
const ID_COLUMN = 1;
const LOGIN_COLUMN = 2;
const TARGET_ID_LENGTH = 10;
//const CACHE_KEY_LOGINS = 'LOGIN_MAP_CACHE'; 

function clearLoginCache() {
  const cache = CacheService.getScriptCache();
  cache.remove(CACHE_KEY_LOGINS);
  console.log("Cache wyczyszczony pomyślnie.");
}

function findRowById(idValue) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][ID_COLUMN - 1]).trim() === String(idValue).trim()) {
      return i + 1; 
    }
  }
  return -1;
}

function manageLoginPopup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    ui.alert(`BŁĄD: Arkusz o nazwie "${SHEET_NAME}" nie istnieje.`);
    return;
  }

  const idPrompt = ui.prompt('🔑 Zarządzanie Loginami', 'Wprowadź/Zbliż Identyfikator (RFID):', ui.ButtonSet.OK_CANCEL);
  if (idPrompt.getSelectedButton() !== ui.Button.OK) return;
  
  let id = idPrompt.getResponseText().trim();
  if (!id) {
    ui.alert('Wprowadzenie Identyfikatora jest obowiązkowe.');
    return;
  }

  if (id.length < TARGET_ID_LENGTH) {
    if (id.length === TARGET_ID_LENGTH - 1) {
       id = '0' + id;
    } else {
       ui.alert(`OSTRZEŻENIE: Identyfikator "${id}" za krótki. Operacja anulowana.`);
       return;
    }
  }

  const existingRow = findRowById(id);
  
  if (existingRow === -1) {
    // --- SCENARIUSZ A: DODAWANIE ---
    const loginPrompt = ui.prompt('🔑 Dodawanie Nowego Loginu', `ID "${id}" jest NOWE. Podaj login:`, ui.ButtonSet.OK_CANCEL);
    if (loginPrompt.getSelectedButton() !== ui.Button.OK) return;

    const login = loginPrompt.getResponseText().trim();
    if (!login) return;

    sheet.appendRow([id, login]);
    
    clearLoginCache(); 
    //ui.alert(`✅ DODANO I ODŚWIEŻONO CACHE!\nID: "${id}"\nLogin: "${login}"`);
    
  } else {
    // --- SCENARIUSZ B: EDYCJA LUB USUNIĘCIE ---
    const currentLogin = sheet.getRange(existingRow, LOGIN_COLUMN).getValue();
    const actionResponse = ui.alert('⚠ ID JUŻ ISTNIEJE!', `Wybrano: "${currentLogin}".\n\n[TAK] = MODYFIKUJ\n[NIE] = USUŃ\n[ANULUJ] = Przerwij`, ui.ButtonSet.YES_NO_CANCEL);

    if (actionResponse === ui.Button.CANCEL) return;
    
    if (actionResponse === ui.Button.NO) {
      const confirmDelete = ui.alert('❌ POTWIERDŹ', `Usunąć "${currentLogin}"?`, ui.ButtonSet.YES_NO);
      if (confirmDelete === ui.Button.YES) {
        sheet.deleteRow(existingRow);
        
        clearLoginCache(); 
        //ui.alert(`✅ USUNIĘTO I ODŚWIEŻONO CACHE!`);
      }
      
    } else if (actionResponse === ui.Button.YES) {
      const newLoginPrompt = ui.prompt('🔄 MODYFIKACJA', `Podaj nowy login dla ID "${id}":`, ui.ButtonSet.OK_CANCEL);
      if (newLoginPrompt.getSelectedButton() !== ui.Button.OK) return;
      
      const newLogin = newLoginPrompt.getResponseText().trim();
      if (!newLogin) return;
      
      sheet.getRange(existingRow, LOGIN_COLUMN).setValue(newLogin);
      
      clearLoginCache(); 
      //ui.alert(`✅ ZMIENIONO I ODŚWIEŻONO CACHE!\nNowy login: **${newLogin}**`);
    }
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('🔑 Zarządzanie')
      .addItem('Start: Dodaj/Modyfikuj/Usuń', 'manageLoginPopup')
      .addToUi();
}
