Office.onReady((info) => {
    document.getElementById("btn-fetch").onclick = () => fetchRowData(null, false);
    document.getElementById("btn-cancel-data").onclick = resetUI;
    document.getElementById("btn-to-machine").onclick = showMachineSelection;
    document.getElementById("btn-start-timer").onclick = writeStartTime;
    document.getElementById("btn-stop").onclick = handleStop;
    document.getElementById("btn-save-partial").onclick = () => saveIncidents(false, false);
    document.getElementById("btn-save-full").onclick = () => saveIncidents(true, false);
    
    // Zmiana rolek w trakcie
    document.getElementById("btn-change-rolls").onclick = changeRollsDuringProcess;
    
    // Awarie
    document.getElementById("btn-awaria").onclick = toggleAwaria;
    
    // Pracownicy
    document.getElementById("btn-add-worker").onclick = () => adjustWorkers(1);
    document.getElementById("btn-sub-worker").onclick = () => adjustWorkers(-1);
    
    // Kalkulator na żywo
    document.getElementById("in-real-rolls").addEventListener("input", updateKitsCalc);
    
    if (info.host === Office.HostType.Excel) {
        setStatus("Inicjalizacja...");
        Excel.run(async (context) => {
            try {
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
let activeSheetName = "";

// Timer główny
let timerInterval = null;
let autoSaveInterval = null;
let secondsElapsed = 0;
let isContinuing = false;

// Przedziały
let currentIntervalIndex = -1; // 0 do 9
let currentIntervalStartCol = -1;

// Pracownicy i operatorzy
let currentWorkersCount = 4;
let startWorkersCount = 4;
let currentWorkerGlobalString = ""; // np. "4+1/3-1"
let currentOperatorGlobalString = ""; // np. "AB/CD"

// Awarie
let isAwariaActive = false;
let awariaTimerInterval = null;
let awariaSecondsElapsed = 0;
let totalAwariaSecondsGlobal = 0;

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

function secondsToHms(d) {
    d = Number(d);
    const h = Math.floor(d / 3600);
    const m = Math.floor(d % 3600 / 60);
    const s = Math.floor(d % 3600 % 60);
    return `${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}:${s.toString().padStart(2,'0')}`;
}

function hmsToSeconds(hms) {
    if (!hms || typeof hms !== 'string') return 0;
    const parts = hms.split(":");
    if (parts.length !== 3) return 0;
    return (+parts[0]) * 3600 + (+parts[1]) * 60 + (+parts[2]);
}

async function onWorksheetActivated(event) {
    if (timerInterval !== null) return;
    
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
    const range = sheet.getRange("A1:EU10").load("values"); 
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
            else if (valUpper === "TOLERANCJA") { colMap.tol = c; }
            else if (valUpper === "LICZBA KITÓW/WARSTWA") { colMap.expLayers = c; }
            else if (valUpper === "MAX ROLEK") { colMap.maxRolls = c; }
            else if (valUpper === "OPERATOR") { colMap.operator = c; }
            else if (valUpper === "ILOŚĆ PRACOWNIKÓW") { colMap.workers = c; }
            else if (valUpper.includes("START (DAY")) { colMap.startGlobal = c; startDayRow = r; }
            else if (valUpper.includes("END (DAY")) { colMap.endGlobal = c; }
            else if (valUpper === "NOTES") { colMap.notes = c; }
            else if (valUpper === "ZMIANA MATERIAŁU?") { colMap.chkMat = c; }
            else if (valUpper === "PRZERWA?") { colMap.chkBreak = c; }
            else if (valUpper === "AWARIA?") { colMap.chkBreakdown = c; }
            else if (valUpper === "MASZYNA") { colMap.machine = c; }
            else if (valUpper === "AWARIE") { colMap.awarie = c; }
            else if (valUpper === "START PRZEDZIAŁÓW") { colMap.intervalsStart = c; }
        }
    }

    const missing = [];
    if (colMap.item === undefined) missing.push("ITEM PRODUKTU");
    if (colMap.rev === undefined) missing.push("REWIZJA");
    if (colMap.product === undefined) missing.push("NAZWA PRODUKTU");
    if (colMap.nesting === undefined) missing.push("NAZWA NESTINGU");
    if (colMap.tol === undefined) missing.push("TOLERANCJA");
    if (colMap.expLayers === undefined) missing.push("LICZBA KITÓW/WARSTWA");
    if (colMap.maxRolls === undefined) missing.push("MAX ROLEK");
    if (colMap.operator === undefined) missing.push("OPERATOR");
    if (colMap.workers === undefined) missing.push("ILOŚĆ PRACOWNIKÓW");
    if (colMap.startGlobal === undefined) missing.push("Start (Day...");
    if (colMap.endGlobal === undefined) missing.push("End (Day...");
    if (colMap.awarie === undefined) missing.push("AWARIE");
    if (colMap.intervalsStart === undefined) missing.push("START PRZEDZIAŁÓW");

    if (missing.length > 0) {
        document.getElementById("unfinished-list").innerHTML = `<div style="color:red; font-size:12px;"><b>Błąd:</b> Brakuje kolumn:<br>${missing.join(", ")}</div>`;
        throw new Error("Brakuje kolumn w pierwszych 10 wierszach: " + missing.join(", "));
    }
    
    dataStartRowIndex = Math.max(itemRow, startDayRow) + 1;
}

async function scanForUnfinished(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRangeByIndexes(dataStartRowIndex, 0, 2000, 150).load("values");
    await context.sync();
    
    const listContainer = document.getElementById("unfinished-list");
    listContainer.innerHTML = "";
    let foundAny = false;
    
    for (let i = 0; i < range.values.length; i++) {
        const row = range.values[i];
        if (!row || (!row[colMap.startGlobal] && !row[colMap.item])) continue;
        
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

function updateKitsCalc() {
    const rolls = parseFloat(document.getElementById("in-real-rolls").value) || 0;
    const kitsPerLayer = parseFloat(document.getElementById("val-kpl").innerText) || 0;
    document.getElementById("lbl-calc-kits").innerText = Math.round(rolls * kitsPerLayer);
}

async function fetchRowData(forcedRowIndex, isCont) {
    isContinuing = isCont;
    try {
        setStatus("Pobieranie danych...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");
            await context.sync();
            activeSheetName = sheet.name; 
            
            let rowIdx = forcedRowIndex;
            if (rowIdx === null) {
                const activeCell = context.workbook.getActiveCell().load("rowIndex");
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
                if (valAA !== "") isContinuing = true;
            }
            currentRowIndex = rowIdx;
            
            // Do 250 kolumn, żeby złapać przedziały
            const rowRange = sheet.getRangeByIndexes(currentRowIndex, 0, 1, 250).load("values");
            await context.sync();
            const vals = rowRange.values[0];
            
            document.getElementById("val-item").innerText = vals[colMap.item] || "-";
            document.getElementById("val-rev").innerText = vals[colMap.rev] || "-";
            document.getElementById("val-product").innerText = vals[colMap.product] || "-";
            document.getElementById("val-nesting").innerText = vals[colMap.nesting] || "-";
            document.getElementById("val-tol").innerText = vals[colMap.tol] || "-";
            document.getElementById("val-warstwy").innerText = vals[colMap.maxRolls] || "-";
            document.getElementById("val-kpl").innerText = vals[colMap.expLayers] || "0";
            
            document.getElementById("in-real-rolls").value = vals[colMap.maxRolls] || "";
            document.getElementById("in-operator").value = "";
            document.getElementById("in-workers").value = "4";
            updateKitsCalc();
            
            currentWorkerGlobalString = vals[colMap.workers] ? vals[colMap.workers].toString() : "";
            currentOperatorGlobalString = vals[colMap.operator] ? vals[colMap.operator].toString() : "";
            totalAwariaSecondsGlobal = hmsToSeconds(vals[colMap.awarie] ? vals[colMap.awarie].toString() : "00:00:00");
            
            if (isContinuing) {
                document.getElementById("btn-to-machine").innerText = "KONTYNUUJ PROCES";
                // Jeśli kontynuujemy, to in-workers ustawiamy na domyślne 4 (bo zaczyna nową zmianę), operator też puste.
                document.getElementById("in-operator").value = "";
                document.getElementById("in-workers").value = "4";
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
            const worksheets = context.workbook.worksheets.load("items/name");
            const activeSheet = worksheets.getActiveWorksheet().load("name");
            await context.sync();
            
            const selMachine = document.getElementById("sel-machine");
            selMachine.innerHTML = "";
            worksheets.items.forEach((s) => {
                const opt = document.createElement("option");
                opt.value = s.name;
                opt.text = s.name;
                if (s.name === activeSheetName) opt.selected = true;
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
    const operator = document.getElementById("in-operator").value.trim() || "Brak";
    startWorkersCount = parseInt(document.getElementById("in-workers").value) || 4;
    currentWorkersCount = startWorkersCount;
    document.getElementById("val-current-workers").innerText = currentWorkersCount;
    
    const realRolls = document.getElementById("in-real-rolls").value;
    const machine = document.getElementById("sel-machine").value;
    
    try {
        setStatus("Rozpoczynanie...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(activeSheetName);
            const dateStr = getFormattedDate();
            
            // Global Strings Updates
            if (currentOperatorGlobalString === "") {
                currentOperatorGlobalString = operator;
            } else {
                currentOperatorGlobalString += "/" + operator;
            }
            
            if (currentWorkerGlobalString === "") {
                currentWorkerGlobalString = startWorkersCount.toString();
            } else {
                currentWorkerGlobalString += "/" + startWorkersCount.toString();
            }
            
            sheet.getCell(currentRowIndex, colMap.operator).values = [[currentOperatorGlobalString]];
            sheet.getCell(currentRowIndex, colMap.workers).values = [[currentWorkerGlobalString]];
            
            if (!isContinuing) {
                sheet.getCell(currentRowIndex, colMap.startGlobal).values = [[dateStr]];
            }
            sheet.getCell(currentRowIndex, colMap.machine).values = [[machine]];
            
            // Znajdowanie przedziału
            const intervalsRange = sheet.getRangeByIndexes(currentRowIndex, colMap.intervalsStart, 1, 60).load("values");
            await context.sync();
            
            currentIntervalIndex = 9; // domyślnie ostatni jeśli wszystkie zajęte
            for (let i = 0; i < 10; i++) {
                const startCellVal = intervalsRange.values[0][i * 6 + 4] ? intervalsRange.values[0][i * 6 + 4].toString().trim() : "";
                if (startCellVal === "") {
                    currentIntervalIndex = i;
                    break;
                }
            }
            
            currentIntervalStartCol = colMap.intervalsStart + (currentIntervalIndex * 6);
            
            // Zapis przedziału (Operator, Pracownicy Start, Pracownicy End (na razie = Start), Rolki, Start, Puste)
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 0).values = [[operator]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 1).values = [[startWorkersCount]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 2).values = [[currentWorkersCount]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 3).values = [[realRolls]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 4).values = [[dateStr]];
            await context.sync();
            
            const iTxt = document.getElementById("val-item").innerText;
            const rTxt = document.getElementById("val-rev").innerText;
            const pTxt = document.getElementById("val-product").innerText;
            document.getElementById("running-info").innerHTML = `Item: ${iTxt}, Rev: ${rTxt}<br>Prod: ${pTxt}<br>Rolki: ${realRolls}`;
            
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
    document.getElementById("timer").innerText = secondsToHms(secondsElapsed);
}

function startAutoSave() {
    autoSaveInterval = setInterval(async () => {
        if (currentRowIndex !== -1 && currentIntervalStartCol !== -1) {
            try {
                await Excel.run(async (ctx) => {
                    const sheet = ctx.workbook.worksheets.getItem(activeSheetName);
                    // col + 5 = Stop Time
                    sheet.getCell(currentRowIndex, currentIntervalStartCol + 5).values = [[getFormattedDate()]];
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

function adjustWorkers(amount) {
    currentWorkersCount += amount;
    if (currentWorkersCount < 0) currentWorkersCount = 0;
    document.getElementById("val-current-workers").innerText = currentWorkersCount;
    
    // Update global string
    const op = amount > 0 ? "+" : "";
    currentWorkerGlobalString += `${op}${amount}`;
    
    // Auto update excel values (End workers + Global string)
    Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(activeSheetName);
        sheet.getCell(currentRowIndex, colMap.workers).values = [[currentWorkerGlobalString]];
        sheet.getCell(currentRowIndex, currentIntervalStartCol + 2).values = [[currentWorkersCount]];
        await ctx.sync();
    }).catch(e => console.warn(e));
}

function toggleAwaria() {
    isAwariaActive = !isAwariaActive;
    const btn = document.getElementById("btn-awaria");
    const timerUI = document.getElementById("awaria-timer");
    
    if (isAwariaActive) {
        document.body.classList.add("awaria-active");
        btn.innerText = "ZAKOŃCZ STAN AWARII";
        timerUI.classList.remove("hidden");
        
        awariaSecondsElapsed = 0;
        timerUI.innerText = secondsToHms(0);
        awariaTimerInterval = setInterval(() => {
            awariaSecondsElapsed++;
            timerUI.innerText = secondsToHms(awariaSecondsElapsed);
        }, 1000);
    } else {
        document.body.classList.remove("awaria-active");
        btn.innerText = "STAN AWARII";
        timerUI.classList.add("hidden");
        clearInterval(awariaTimerInterval);
        
        totalAwariaSecondsGlobal += awariaSecondsElapsed;
        
        // Zapisz sumę awarii do excela
        Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getItem(activeSheetName);
            sheet.getCell(currentRowIndex, colMap.awarie).values = [[secondsToHms(totalAwariaSecondsGlobal)]];
            await ctx.sync();
        }).catch(e => console.warn(e));
    }
}

async function changeRollsDuringProcess() {
    const newRollsStr = prompt("Podaj nową liczbę rolek:");
    if (!newRollsStr) return;
    const newRolls = parseFloat(newRollsStr);
    if (isNaN(newRolls)) {
        alert("Błędna wartość."); return;
    }
    
    // 1. Zakończ obecny przedział (zapisz Stop Time)
    // 2. Wystartuj nowy przedział z nową liczbą rolek
    try {
        setStatus("Zmiana rolek - zapis...");
        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getItem(activeSheetName);
            const dateStr = getFormattedDate();
            
            // Koniec starego
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 5).values = [[dateStr]];
            
            // Ustal nowy index przedziału
            currentIntervalIndex++;
            if (currentIntervalIndex > 9) currentIntervalIndex = 9; // Overwrite last
            currentIntervalStartCol = colMap.intervalsStart + (currentIntervalIndex * 6);
            
            // Zapis nowego przedziału
            const operator = document.getElementById("in-operator").value.trim() || "Brak";
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 0).values = [[operator]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 1).values = [[currentWorkersCount]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 2).values = [[currentWorkersCount]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 3).values = [[newRolls]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 4).values = [[dateStr]];
            await ctx.sync();
            
            document.getElementById("in-real-rolls").value = newRolls;
            updateKitsCalc();
            
            const iTxt = document.getElementById("val-item").innerText;
            const rTxt = document.getElementById("val-rev").innerText;
            const pTxt = document.getElementById("val-product").innerText;
            document.getElementById("running-info").innerHTML = `Item: ${iTxt}, Rev: ${rTxt}<br>Prod: ${pTxt}<br>Rolki: ${newRolls}`;
            
            setStatus("Zmieniono rolki. Przedział rozdzielony.");
        });
    } catch (e) {
        console.error(e);
        setStatus("Błąd zmiany rolek: " + e.message);
    }
}

function handleStop() {
    if (isAwariaActive) {
        alert("Najpierw wyłącz Stan Awarii!");
        return;
    }
    clearInterval(timerInterval);
    stopAutoSave(); 
    
    document.getElementById("running-card").classList.add("hidden");
    document.getElementById("incidents-card").classList.remove("hidden");
    setStatus("Czas zatrzymany. Wybierz opcję zakończenia.");
}

async function saveIncidents(fullComplete, isRollChange) {
    const material = document.getElementById("chk-material").checked;
    const breakTime = document.getElementById("chk-break").checked;
    const breakdown = document.getElementById("chk-breakdown").checked;
    const incidentsText = document.getElementById("in-other-incidents").value;
    
    try {
        setStatus("Zapisywanie...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(activeSheetName);
            const dateStr = getFormattedDate();
            
            if (currentIntervalStartCol !== -1) {
                sheet.getCell(currentRowIndex, currentIntervalStartCol + 5).values = [[dateStr]]; // Stop Time
            }
            
            if (fullComplete) {
                sheet.getCell(currentRowIndex, colMap.endGlobal).values = [[dateStr]];
                
                // --- PODSUMOWANIE (3 kolumny na samym końcu przedziałów = intervalsStart + 60) ---
                const dataRange = sheet.getRangeByIndexes(currentRowIndex, colMap.intervalsStart, 1, 60).load("values");
                await context.sync();
                
                const ivals = dataRange.values[0];
                let totalTimeMs = 0;
                let totalKits = 0;
                let sumWorkerTime = 0;
                
                const kitsPerLayer = parseFloat(document.getElementById("val-kpl").innerText) || 0;
                
                for (let i = 0; i < 10; i++) {
                    const wStart = parseFloat(ivals[i*6 + 1]);
                    const wStop = parseFloat(ivals[i*6 + 2]);
                    const rolls = parseFloat(ivals[i*6 + 3]) || 0;
                    const startStr = ivals[i*6 + 4];
                    const stopStr = ivals[i*6 + 5];
                    
                    if (startStr && stopStr) {
                        const tStart = new Date(startStr).getTime();
                        const tStop = new Date(stopStr).getTime();
                        if (!isNaN(tStart) && !isNaN(tStop)) {
                            const durationMs = tStop - tStart;
                            if (durationMs > 0) {
                                totalTimeMs += durationMs;
                                totalKits += (rolls * kitsPerLayer);
                                
                                const avgWorkers = (isNaN(wStart) || isNaN(wStop)) ? 0 : ((wStart + wStop) / 2);
                                sumWorkerTime += (avgWorkers * durationMs);
                            }
                        }
                    }
                }
                
                const totalAwariaMs = totalAwariaSecondsGlobal * 1000;
                let netTimeMs = totalTimeMs - totalAwariaMs;
                if (netTimeMs < 0) netTimeMs = 0;
                
                let avgWorkersFinal = 0;
                if (totalTimeMs > 0) {
                    avgWorkersFinal = sumWorkerTime / totalTimeMs;
                }
                
                const netTimeHms = secondsToHms(netTimeMs / 1000);
                
                // Write summary to 60, 61, 62
                sheet.getCell(currentRowIndex, colMap.intervalsStart + 60).values = [[netTimeHms]]; // Czas Netto
                sheet.getCell(currentRowIndex, colMap.intervalsStart + 61).values = [[Math.round(totalKits)]]; // Suma Kitów
                sheet.getCell(currentRowIndex, colMap.intervalsStart + 62).values = [[avgWorkersFinal.toFixed(2)]]; // Średnia prac
            }
            
            if (colMap.chkMat !== undefined) sheet.getCell(currentRowIndex, colMap.chkMat).values = [[material ? "TAK" : ""]];
            if (colMap.chkBreak !== undefined) sheet.getCell(currentRowIndex, colMap.chkBreak).values = [[breakTime ? "TAK" : ""]];
            if (colMap.chkBreakdown !== undefined) sheet.getCell(currentRowIndex, colMap.chkBreakdown).values = [[breakdown ? "TAK" : ""]];
            if (colMap.notes !== undefined) sheet.getCell(currentRowIndex, colMap.notes).values = [[incidentsText]];
            
            await context.sync();
            
            resetUI();
            setStatus(fullComplete ? "Zakończono produkt pomyślnie!" : "Przerwano produkt. Zapisano zmianę.");
            await scanForUnfinished(context);
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd zapisu: " + error.message);
    }
}

function resetUI() {
    clearInterval(timerInterval);
    clearInterval(awariaTimerInterval);
    stopAutoSave();
    
    isAwariaActive = false;
    document.body.classList.remove("awaria-active");
    document.getElementById("awaria-timer").classList.add("hidden");
    document.getElementById("btn-awaria").innerText = "STAN AWARII";
    
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
