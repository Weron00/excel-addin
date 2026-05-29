Office.onReady((info) => {
    document.getElementById("btn-fetch").onclick = () => fetchRowData(null, false);
    document.getElementById("btn-cancel-data").onclick = resetUI;
    document.getElementById("btn-to-machine").onclick = showMachineSelection;
    document.getElementById("btn-start-timer").onclick = writeStartTime;
    document.getElementById("btn-stop").onclick = handleStop;
    document.getElementById("btn-save-partial").onclick = () => saveIncidents(false);
    document.getElementById("btn-save-full").onclick = () => saveIncidents(true);
    
    if (info.host === Office.HostType.Excel) {
        setStatus("Inicjalizacja...");
        Excel.run(async (context) => {
            try {
                // Podpinamy nasłuchiwanie zmiany zakładki!
                context.workbook.worksheets.onActivated.add(onWorksheetActivated);
                
                await initializeColumnMap(context);
                setStatus("Skanowanie listy niezakończonych...");
                await scanForUnfinished(context);
            } catch (e) {
                console.error(e);
                setStatus("Błąd: " + e.message);
            }
        });
    } else {
        setStatus("UWAGA: Uruchom ten link jako Dodatek Wewnątrz Excela!");
    }
});

let colMap = {};
let dataStartRowIndex = -1;

let currentRowIndex = -1;
let activeSheetName = ""; // Zmienna przechowująca zakładkę, na której rozpoczęto pracę
let timerInterval = null;
let autoSaveInterval = null;
let secondsElapsed = 0;
let isContinuing = false;
let currentIntervalStartCol = -1;
let currentIntervalEndCol = -1;

function getFormattedDate() {
    const d = new Date();
    const datePart = `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
    let hours = d.getHours();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12;
    const minutes = d.getMinutes().toString().padStart(2, '0');
    const seconds = d.getSeconds().toString().padStart(2, '0');
    return `${datePart} ${hours}:${minutes}:${seconds} ${ampm}`;
}

async function onWorksheetActivated(event) {
    if (timerInterval !== null) {
        // Jeśli leci czas, ignorujemy zmianę zakładki, żeby nie zepsuć interfejsu 
        // (zapisy i tak polecą do poprawnego arkusza dzięki zmiennej activeSheetName)
        return;
    }
    
    // Jeśli czasu nie ma, resetujemy widok i mapujemy nową zakładkę
    clearInterval(timerInterval);
    stopAutoSave();
    document.getElementById("data-card").classList.add("hidden");
    document.getElementById("machine-card").classList.add("hidden");
    document.getElementById("running-card").classList.add("hidden");
    document.getElementById("incidents-card").classList.add("hidden");
    document.getElementById("initial-card").classList.remove("hidden");
    
    setStatus("Zmieniono zakładkę. Remapowanie kolumn...");
    try {
        await Excel.run(async (context) => {
            await initializeColumnMap(context);
            await scanForUnfinished(context);
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd zmiany zakładki: " + error.message);
    }
}

async function initializeColumnMap(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    // Pobieranie pierwszych 10 wierszy i dużej liczby kolumn
    const range = sheet.getRange("A1:EU10"); 
    range.load("values");
    await context.sync();
    
    colMap = {};
    let itemRow = -1;
    let startDayRow = -1;

    for (let r = 0; r < 10; r++) {
        const row = range.values[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            const val = row[c] ? row[c].toString().trim() : "";
            if (!val) continue;

            const valUpper = val.toUpperCase();

            if (valUpper === "ITEM PRODUKTU") { colMap.item = c; itemRow = r; }
            else if (valUpper === "REWIZJA") { colMap.rev = c; }
            else if (valUpper === "NAZWA PRODUKTU") { colMap.product = c; }
            else if (valUpper === "NAZWA NESTINGU") { colMap.nesting = c; }
            else if (valUpper === "LICZBA KITÓW/WARSTWA") { colMap.expLayers = c; }
            else if (valUpper === "KITY DO ZROBIENIA") { colMap.kits = c; }
            else if (valUpper === "OPERATOR") { colMap.operator = c; }
            else if (valUpper === "ILOŚĆ PRACOWNIKÓW") { colMap.workers = c; }
            else if (valUpper === "RZECZYWISTE WARSTWY") { colMap.realLayers = c; }
            else if (valUpper.includes("START (DAY")) { colMap.startGlobal = c; startDayRow = r; }
            else if (valUpper.includes("END (DAY")) { colMap.endGlobal = c; }
            else if (valUpper === "NOTES") { colMap.notes = c; }
            else if (valUpper === "ZMIANA MATERIAŁU?") { colMap.chkMat = c; }
            else if (valUpper === "PRZERWA?") { colMap.chkBreak = c; }
            else if (valUpper === "AWARIA?") { colMap.chkBreakdown = c; }
            else if (valUpper === "MASZYNA") { colMap.machine = c; }
            else if (valUpper === "PRZEDZIAŁ 1 START") { colMap.int1S = c; }
            else if (valUpper === "PRZEDZIAŁ 1 END") { colMap.int1E = c; }
            else if (valUpper === "PRZEDZIAŁ 2 START") { colMap.int2S = c; }
            else if (valUpper === "PRZEDZIAŁ 2 END") { colMap.int2E = c; }
            else if (valUpper === "PRZEDZIAŁ 3 START") { colMap.int3S = c; }
            else if (valUpper === "PRZEDZIAŁ 3 END") { colMap.int3E = c; }
            else if (valUpper === "PRZEDZIAŁ KOMENTARZ") { colMap.intComment = c; }
        }
    }

    const missing = [];
    if (colMap.item === undefined) missing.push("ITEM PRODUKTU");
    if (colMap.rev === undefined) missing.push("REWIZJA");
    if (colMap.product === undefined) missing.push("NAZWA PRODUKTU");
    if (colMap.nesting === undefined) missing.push("NAZWA NESTINGU");
    if (colMap.expLayers === undefined) missing.push("LICZBA KITÓW/WARSTWA");
    if (colMap.kits === undefined) missing.push("KITY DO ZROBIENIA");
    if (colMap.operator === undefined) missing.push("OPERATOR");
    if (colMap.workers === undefined) missing.push("ILOŚĆ PRACOWNIKÓW");
    if (colMap.realLayers === undefined) missing.push("RZECZYWISTE WARSTWY");
    if (colMap.startGlobal === undefined) missing.push("Start (Day...");
    if (colMap.endGlobal === undefined) missing.push("End (Day...");
    if (colMap.notes === undefined) missing.push("NOTES");
    if (colMap.chkMat === undefined) missing.push("ZMIANA MATERIAŁU?");
    if (colMap.chkBreak === undefined) missing.push("PRZERWA?");
    if (colMap.chkBreakdown === undefined) missing.push("AWARIA?");
    if (colMap.machine === undefined) missing.push("MASZYNA");
    if (colMap.int1S === undefined) missing.push("PRZEDZIAŁ 1 START");
    if (colMap.int1E === undefined) missing.push("PRZEDZIAŁ 1 END");
    if (colMap.int2S === undefined) missing.push("PRZEDZIAŁ 2 START");
    if (colMap.int2E === undefined) missing.push("PRZEDZIAŁ 2 END");
    if (colMap.int3S === undefined) missing.push("PRZEDZIAŁ 3 START");
    if (colMap.int3E === undefined) missing.push("PRZEDZIAŁ 3 END");
    if (colMap.intComment === undefined) missing.push("PRZEDZIAŁ KOMENTARZ");

    if (missing.length > 0) {
        document.getElementById("unfinished-list").innerHTML = `<div style="color:red; font-size:12px;"><b>Błąd:</b> Brakuje kolumn:<br>${missing.join(", ")}</div>`;
        throw new Error("Brakuje kolumn w pierwszych 10 wierszach: " + missing.join(", "));
    }
    
    dataStartRowIndex = Math.max(itemRow, startDayRow) + 1;
}

async function scanForUnfinished(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    // Nakładka szuka produktów na konkretnym arkuszu, w którym aktualnie jesteś
    const range = sheet.getRangeByIndexes(dataStartRowIndex, 0, 2000, 150);
    range.load("values");
    await context.sync();
    
    const listContainer = document.getElementById("unfinished-list");
    listContainer.innerHTML = "";
    let foundAny = false;
    
    for (let i = 0; i < range.values.length; i++) {
        const row = range.values[i];
        if (!row) continue;
        
        if (!row[colMap.startGlobal] && !row[colMap.item]) continue;
        
        const valAA = row[colMap.startGlobal] ? row[colMap.startGlobal].toString().trim() : "";
        const valAB = row[colMap.endGlobal] ? row[colMap.endGlobal].toString().trim() : "";
        
        if (valAA !== "" && valAB === "") {
            foundAny = true;
            const itemValue = (colMap.item !== undefined && row[colMap.item]) ? row[colMap.item].toString() : "Brak Itemu";
            
            const btn = document.createElement("button");
            btn.className = "unfinished-item";
            btn.innerText = `Wiersz ${dataStartRowIndex + i + 1} | Item: ${itemValue}`;
            btn.onclick = () => fetchRowData(dataStartRowIndex + i, true);
            listContainer.appendChild(btn);
        }
    }
    
    if (!foundAny) {
        listContainer.innerHTML = `<div style="font-size:12px; color:#9ca3af;">Brak niezakończonych zadań na tej zakładce.</div>`;
    }
    
    document.getElementById("unfinished-container").style.display = "block";
    setStatus("Gotowe. Zaznacz wiersz lub wybierz z listy.");
}

async function fetchRowData(forcedRowIndex, isCont) {
    isContinuing = isCont;
    try {
        setStatus("Pobieranie danych...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");
            await context.sync();
            activeSheetName = sheet.name; // Zapisujemy nazwę arkusza operacyjnego
            
            let rowIdx = forcedRowIndex;
            if (rowIdx === null) {
                const activeCell = context.workbook.getActiveCell();
                activeCell.load("rowIndex");
                await context.sync();
                rowIdx = activeCell.rowIndex;
                
                if (rowIdx < dataStartRowIndex) {
                    setStatus("Wybrano wiersz nagłówkowy. Wybierz wiersz poniżej tytułów.");
                    return;
                }
                
                const checkAA = sheet.getCell(rowIdx, colMap.startGlobal).load("values");
                const checkAB = sheet.getCell(rowIdx, colMap.endGlobal).load("values");
                await context.sync();
                
                const valAA = checkAA.values[0][0] ? checkAA.values[0][0].toString().trim() : "";
                const valAB = checkAB.values[0][0] ? checkAB.values[0][0].toString().trim() : "";
                
                if (valAB !== "") {
                    setStatus("UWAGA: Ten produkt został już całkowicie zakończony!");
                    return;
                }
                if (valAA !== "") {
                    isContinuing = true;
                }
            }
            
            currentRowIndex = rowIdx;
            
            const rowRange = sheet.getRangeByIndexes(currentRowIndex, 0, 1, 150).load("values");
            await context.sync();
            const vals = rowRange.values[0];
            
            document.getElementById("val-item").innerText = (colMap.item !== undefined && vals[colMap.item]) ? vals[colMap.item] : "-";
            document.getElementById("val-rev").innerText = (colMap.rev !== undefined && vals[colMap.rev]) ? vals[colMap.rev] : "-";
            document.getElementById("val-product").innerText = (colMap.product !== undefined && vals[colMap.product]) ? vals[colMap.product] : "-";
            document.getElementById("val-nesting").innerText = (colMap.nesting !== undefined && vals[colMap.nesting]) ? vals[colMap.nesting] : "-";
            document.getElementById("val-warstwy").innerText = (colMap.expLayers !== undefined && vals[colMap.expLayers]) ? vals[colMap.expLayers] : "-";
            document.getElementById("val-kit").innerText = (colMap.kits !== undefined && vals[colMap.kits]) ? vals[colMap.kits] : "-";
            
            document.getElementById("in-real-layers").value = (colMap.expLayers !== undefined && vals[colMap.expLayers]) ? vals[colMap.expLayers] : "";
            document.getElementById("in-operator").value = "";
            document.getElementById("in-workers").value = "4";
            
            document.getElementById("chk-material").checked = false;
            document.getElementById("chk-break").checked = false;
            document.getElementById("chk-breakdown").checked = false;
            document.getElementById("in-other-incidents").value = "";
            
            if (isContinuing) {
                document.getElementById("btn-to-machine").innerText = "KONTYNUUJ PROCES";
                if (colMap.operator !== undefined && vals[colMap.operator]) document.getElementById("in-operator").value = vals[colMap.operator];
                if (colMap.workers !== undefined && vals[colMap.workers]) document.getElementById("in-workers").value = vals[colMap.workers];
                if (colMap.realLayers !== undefined && vals[colMap.realLayers]) document.getElementById("in-real-layers").value = vals[colMap.realLayers];
                
                if (colMap.chkMat !== undefined && vals[colMap.chkMat] === "TAK") document.getElementById("chk-material").checked = true;
                if (colMap.chkBreak !== undefined && vals[colMap.chkBreak] === "TAK") document.getElementById("chk-break").checked = true;
                if (colMap.chkBreakdown !== undefined && vals[colMap.chkBreakdown] === "TAK") document.getElementById("chk-breakdown").checked = true;
                if (colMap.notes !== undefined && vals[colMap.notes]) document.getElementById("in-other-incidents").value = vals[colMap.notes];
            } else {
                document.getElementById("btn-to-machine").innerText = "DALEJ";
            }
            
            document.getElementById("initial-card").classList.add("hidden");
            document.getElementById("data-card").classList.remove("hidden");
            
            setStatus(`Pobrano dane z wiersza ${currentRowIndex + 1}.`);
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd pobierania: " + error.message);
    }
}

async function showMachineSelection() {
    try {
        setStatus("Ładowanie listy maszyn...");
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name");
            const activeSheet = worksheets.getActiveWorksheet();
            activeSheet.load("name");
            await context.sync();
            
            const selMachine = document.getElementById("sel-machine");
            selMachine.innerHTML = "";
            
            worksheets.items.forEach((sheet) => {
                const opt = document.createElement("option");
                opt.value = sheet.name;
                opt.text = sheet.name;
                // Jeśli nazwa operacyjnego arkusza się zgadza to domyślnie wybrana
                if (sheet.name === activeSheetName) {
                    opt.selected = true;
                }
                selMachine.appendChild(opt);
            });
            
            document.getElementById("data-card").classList.add("hidden");
            document.getElementById("machine-card").classList.remove("hidden");
            setStatus("Wybierz maszynę i kliknij Start Czasu.");
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd wyboru maszyny: " + error.message);
    }
}

async function writeStartTime() {
    const operator = document.getElementById("in-operator").value;
    const workers = document.getElementById("in-workers").value;
    const realLayers = document.getElementById("in-real-layers").value;
    const machine = document.getElementById("sel-machine").value;
    
    try {
        setStatus("Rozpoczynanie...");
        await Excel.run(async (context) => {
            // Zapis do bezpiecznego arkusza operacyjnego (odporność na klikanie po zakładkach)
            const sheet = context.workbook.worksheets.getItem(activeSheetName);
            const dateStr = getFormattedDate();
            
            if (!isContinuing) {
                if (colMap.operator !== undefined) sheet.getCell(currentRowIndex, colMap.operator).values = [[operator]];
                if (colMap.workers !== undefined) sheet.getCell(currentRowIndex, colMap.workers).values = [[workers]];
                if (colMap.realLayers !== undefined) sheet.getCell(currentRowIndex, colMap.realLayers).values = [[realLayers]];
                if (colMap.startGlobal !== undefined) sheet.getCell(currentRowIndex, colMap.startGlobal).values = [[dateStr]];
            }
            
            if (colMap.machine !== undefined) sheet.getCell(currentRowIndex, colMap.machine).values = [[machine]];
            
            const colsToLoad = [];
            if (colMap.int1S !== undefined) colsToLoad.push(sheet.getCell(currentRowIndex, colMap.int1S).load("values"));
            if (colMap.int2S !== undefined) colsToLoad.push(sheet.getCell(currentRowIndex, colMap.int2S).load("values"));
            if (colMap.int3S !== undefined) colsToLoad.push(sheet.getCell(currentRowIndex, colMap.int3S).load("values"));
            
            await context.sync();
            
            const v1 = (colsToLoad.length > 0 && colsToLoad[0].values[0][0]) ? colsToLoad[0].values[0][0].toString().trim() : "";
            const v2 = (colsToLoad.length > 1 && colsToLoad[1].values[0][0]) ? colsToLoad[1].values[0][0].toString().trim() : "";
            const v3 = (colsToLoad.length > 2 && colsToLoad[2].values[0][0]) ? colsToLoad[2].values[0][0].toString().trim() : "";
            
            if (v1 === "" && colMap.int1S !== undefined) {
                currentIntervalStartCol = colMap.int1S; currentIntervalEndCol = colMap.int1E;
            } else if (v2 === "" && colMap.int2S !== undefined) {
                currentIntervalStartCol = colMap.int2S; currentIntervalEndCol = colMap.int2E;
            } else if (v3 === "" && colMap.int3S !== undefined) {
                currentIntervalStartCol = colMap.int3S; currentIntervalEndCol = colMap.int3E;
            } else {
                currentIntervalStartCol = colMap.int3S; currentIntervalEndCol = colMap.int3E;
                if (colMap.intComment !== undefined) {
                    sheet.getCell(currentRowIndex, colMap.intComment).values = [["Proces wznawiano więcej niż 3 razy"]];
                }
            }
            
            if (currentIntervalStartCol !== undefined && currentIntervalStartCol !== -1) {
                sheet.getCell(currentRowIndex, currentIntervalStartCol).values = [[dateStr]];
                await context.sync();
            }
            
            const iTxt = document.getElementById("val-item").innerText;
            const rTxt = document.getElementById("val-rev").innerText;
            const pTxt = document.getElementById("val-product").innerText;
            const nTxt = document.getElementById("val-nesting").innerText;
            document.getElementById("running-info").innerHTML = `Item: ${iTxt}, Rev: ${rTxt}<br>Prod: ${pTxt}<br>Nesting: ${nTxt}`;
            
            document.getElementById("machine-card").classList.add("hidden");
            document.getElementById("running-card").classList.remove("hidden");
            
            startTimer();
            startAutoSave();
            setStatus("W trakcie pracy...");
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd Start: " + error.message);
    }
}

function startTimer() {
    secondsElapsed = 0;
    updateTimerDisplay();
    timerInterval = setInterval(() => {
        secondsElapsed++;
        updateTimerDisplay();
    }, 1000);
}

function updateTimerDisplay() {
    const hrs = Math.floor(secondsElapsed / 3600).toString().padStart(2, '0');
    const mins = Math.floor((secondsElapsed % 3600) / 60).toString().padStart(2, '0');
    const secs = (secondsElapsed % 60).toString().padStart(2, '0');
    document.getElementById("timer").innerText = `${hrs}:${mins}:${secs}`;
}

function startAutoSave() {
    autoSaveInterval = setInterval(async () => {
        if (currentRowIndex !== -1 && currentIntervalEndCol !== -1 && currentIntervalEndCol !== undefined) {
            try {
                await Excel.run(async (ctx) => {
                    const sheet = ctx.workbook.worksheets.getItem(activeSheetName);
                    sheet.getCell(currentRowIndex, currentIntervalEndCol).values = [[getFormattedDate()]];
                    await ctx.sync();
                });
            } catch (e) {
                console.warn("Autozapis w tle:", e);
            }
        }
    }, 30000);
}

function stopAutoSave() {
    if (autoSaveInterval) {
        clearInterval(autoSaveInterval);
        autoSaveInterval = null;
    }
}

function handleStop() {
    clearInterval(timerInterval);
    stopAutoSave(); 
    
    document.getElementById("running-card").classList.add("hidden");
    document.getElementById("incidents-card").classList.remove("hidden");
    setStatus("Czas zatrzymany. Wybierz opcję zakończenia.");
}

async function saveIncidents(fullComplete) {
    const material = document.getElementById("chk-material").checked;
    const breakTime = document.getElementById("chk-break").checked;
    const breakdown = document.getElementById("chk-breakdown").checked;
    const incidentsText = document.getElementById("in-other-incidents").value;
    
    try {
        setStatus("Zapisywanie...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(activeSheetName);
            const dateStr = getFormattedDate();
            
            if (currentIntervalEndCol !== -1 && currentIntervalEndCol !== undefined) {
                sheet.getCell(currentRowIndex, currentIntervalEndCol).values = [[dateStr]];
            }
            
            if (fullComplete && colMap.endGlobal !== undefined) {
                sheet.getCell(currentRowIndex, colMap.endGlobal).values = [[dateStr]];
            }
            
            if (colMap.chkMat !== undefined) sheet.getCell(currentRowIndex, colMap.chkMat).values = [[material ? "TAK" : ""]];
            if (colMap.chkBreak !== undefined) sheet.getCell(currentRowIndex, colMap.chkBreak).values = [[breakTime ? "TAK" : ""]];
            if (colMap.chkBreakdown !== undefined) sheet.getCell(currentRowIndex, colMap.chkBreakdown).values = [[breakdown ? "TAK" : ""]];
            if (colMap.notes !== undefined) sheet.getCell(currentRowIndex, colMap.notes).values = [[incidentsText]];
            
            await context.sync();
            
            resetUI();
            setStatus(fullComplete ? "Zakończono produkt pomyślnie!" : "Przerwano produkt. Zapisano zmianę.");
            
            // Ponieważ użytkownik mógł zmienić zakładkę podczas trwania procesu,
            // dla bezpieczeństwa remapujemy kolumny pod obecną zakładkę przed skanem
            await initializeColumnMap(context);
            await scanForUnfinished(context);
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd zapisu: " + error.message);
    }
}

function resetUI() {
    clearInterval(timerInterval);
    stopAutoSave();
    document.getElementById("data-card").classList.add("hidden");
    document.getElementById("machine-card").classList.add("hidden");
    document.getElementById("running-card").classList.add("hidden");
    document.getElementById("incidents-card").classList.add("hidden");
    document.getElementById("initial-card").classList.remove("hidden");
}

function setStatus(message) {
    const statusEl = document.getElementById("status-message");
    if (statusEl) {
        statusEl.innerText = message;
    }
}
