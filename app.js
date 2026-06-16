Office.onReady((info) => {
    document.getElementById("btn-fetch").onclick = () => fetchRowData(null, false);
    document.getElementById("btn-cancel-data").onclick = resetUI;
    document.getElementById("btn-to-machine").onclick = handleToMachineClick;
    document.getElementById("btn-start-timer").onclick = handleStartTimerClick;
    document.getElementById("btn-stop").onclick = handleStop;
    document.getElementById("btn-save-partial").onclick = () => saveIncidents(false);
    document.getElementById("btn-save-full").onclick = () => saveIncidents(true);
    
    // Nawigacja wierszami
    document.getElementById("btn-prev-row").onclick = () => { if (currentRowIndex > 1) fetchRowData(currentRowIndex - 1, isContinuing); };
    document.getElementById("btn-next-row").onclick = () => fetchRowData(currentRowIndex + 1, isContinuing);
    
    // Zmiana rolek w trakcie
    document.getElementById("btn-change-rolls").onclick = () => {
        document.getElementById("change-rolls-panel").classList.remove("hidden");
        document.getElementById("in-new-rolls").value = document.getElementById("in-real-rolls").value;
        updateModalKitsCalc();
    };
    document.getElementById("in-new-rolls").addEventListener("input", updateModalKitsCalc);
    document.getElementById("btn-confirm-rolls").onclick = confirmChangeRolls;
    
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
let currentColumnIndex = 0;
let activeSheetName = "";

// Timer główny
let timerInterval = null;
let autoSaveInterval = null;
let secondsElapsed = 0;
let isContinuing = false;

// Przedziały
let currentIntervalIndex = -1; // 0 do 9
let currentIntervalStartCol = -1;
let previousTotalGrossSeconds = 0;
let theoreticalSeconds = 0;

// Zamknięcia i kontynuacje
let lastIntervalUnexpected = false;
let lastIntervalIndex = -1;
let unexpectedIntervalDuration = 0;
let resumeUnexpected = false;

// Pracownicy i operatorzy
let currentWorkersCount = 4;
let startWorkersCount = 4;
let intervalWorkerDiff = 0; // śledzi zmiany od startu przedziału
let previousGlobalWorkerString = ""; // co było przed wejściem w ten przedział
let currentWorkerGlobalString = ""; // cały wynik (np. "4+3/3-1")
let currentOperatorGlobalString = "";

// Do pamiętania maszyny przy kontynuacji
let selectedMachineForContinuation = "";

// Awarie
let isAwariaActive = false;
let awariaTimerInterval = null;
let awariaSecondsElapsed = 0;
let totalAwariaSecondsGlobal = 0;

function safeStr(val) {
    if (val === undefined || val === null) return "";
    return "'" + val.toString();
}

function getFormattedDate() {
    const d = new Date();
    const day = d.getDate().toString().padStart(2, '0');
    const mo = (d.getMonth() + 1).toString().padStart(2, '0');
    const y = d.getFullYear();
    const h = d.getHours().toString().padStart(2, '0');
    const m = d.getMinutes().toString().padStart(2, '0');
    return `${day}-${mo}-${y} ${h}:${m}`;
}

function parseCustomDate(val) {
    if (val === undefined || val === null || val === "") return NaN;
    const str = val.toString().trim().replace(/^'/, "");
    
    // Obsługa numerycznych dat i czasu bezpośrednio z Excela
    if (!isNaN(parseFloat(str)) && !str.includes("-") && !str.includes(":")) {
        const num = parseFloat(str);
        return new Date(1899, 11, 30).getTime() + Math.round(num * 86400 * 1000);
    }
    
    const parts = str.split(" ");
    if (parts.length < 2) return NaN;
    const dateParts = parts[0].split("-");
    const timeParts = parts[1].split(":");
    if (dateParts.length !== 3 || timeParts.length < 2) return NaN;
    return new Date(+dateParts[2], (+dateParts[1]) - 1, +dateParts[0], +timeParts[0], +timeParts[1]).getTime();
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
            else if (valUpper === "CZAS TEORETYCZNY") { colMap.theoretical = c; }
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
        document.getElementById("btn-fetch").style.display = "none";
        document.getElementById("unfinished-list").innerHTML = `<div style="color:#dc2626; font-size:13px; padding:10px; border:1px solid #fca5a5; background:#fef2f2; border-radius:6px; line-height:1.4;"><b>BŁĄD STRUKTURY ARKUSZA:</b><br>Brakuje następujących kolumn w pierwszych 10 wierszach:<br><br><b>${missing.join(", ")}</b></div>`;
        document.getElementById("unfinished-container").style.display = "block";
        throw new Error("Brakuje kolumn w pierwszych 10 wierszach: " + missing.join(", "));
    } else {
        document.getElementById("btn-fetch").style.display = "inline-block";
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
    window.unfinishedMachines = {};
    
    for (let i = 0; i < range.values.length; i++) {
        const row = range.values[i];
        if (!row || (!row[colMap.startGlobal] && !row[colMap.item])) continue;
        
        const valAA = row[colMap.startGlobal] ? row[colMap.startGlobal].toString().trim() : "";
        const valAB = row[colMap.endGlobal] ? row[colMap.endGlobal].toString().trim() : "";
        
        if (valAA !== "" && valAB === "") {
            foundAny = true;
            const itemValue = (colMap.item !== undefined && row[colMap.item]) ? row[colMap.item].toString() : "Brak Itemu";
            const machineVal = (colMap.machine !== undefined && row[colMap.machine]) ? row[colMap.machine].toString().trim() : "";
            if (machineVal !== "") {
                window.unfinishedMachines[machineVal] = {
                    row: dataStartRowIndex + i,
                    item: itemValue
                };
            }
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
    const kitsPerLayer = parseFloat(document.getElementById("val-kpl").textContent) || 0;
    document.getElementById("lbl-calc-kits").innerText = Math.round(rolls * kitsPerLayer);
}

function updateModalKitsCalc() {
    const rolls = parseFloat(document.getElementById("in-new-rolls").value) || 0;
    const kitsPerLayer = parseFloat(document.getElementById("val-kpl").textContent) || 0;
    document.getElementById("lbl-new-calc-kits").innerText = Math.round(rolls * kitsPerLayer);
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
            const activeCell = context.workbook.getActiveCell().load(["rowIndex", "columnIndex"]);
            await context.sync();
            currentColumnIndex = activeCell.columnIndex;
            
            if (rowIdx === null) {
                rowIdx = activeCell.rowIndex;
            }
            if (rowIdx < dataStartRowIndex) {
                setStatus("Wybierz wiersz poniżej tytułów.");
                return;
            }
            currentRowIndex = rowIdx;
            
            // Pobieramy cały wiersz (do 250 kolumn), co zmniejsza liczbę requestów.
            const rowRange = sheet.getRangeByIndexes(currentRowIndex, 0, 1, 250).load("values");
            await context.sync();
            const vals = rowRange.values[0];
            
            // Weryfikacja zakończenia
            const valAB = vals[colMap.endGlobal] ? vals[colMap.endGlobal].toString().trim() : "";
            if (valAB !== "") {
                setStatus("UWAGA: Ten produkt został już całkowicie zakończony!");
                return;
            }
            const valAA = vals[colMap.startGlobal] ? vals[colMap.startGlobal].toString().trim() : "";
            isContinuing = (valAA !== "");
            
            if (isContinuing) {
                document.getElementById("btn-prev-row").style.display = "none";
                document.getElementById("btn-next-row").style.display = "none";
            } else {
                sheet.getCell(currentRowIndex, currentColumnIndex).select();
                document.getElementById("btn-prev-row").style.display = "block";
                document.getElementById("btn-next-row").style.display = "block";
            }
            
            document.getElementById("val-item").innerText = vals[colMap.item] || "-";
            document.getElementById("val-rev").innerText = vals[colMap.rev] || "-";
            document.getElementById("val-product").innerText = vals[colMap.product] || "-";
            document.getElementById("val-nesting").innerText = vals[colMap.nesting] || "-";
            document.getElementById("val-tol").innerText = vals[colMap.tol] || "-";
            document.getElementById("val-warstwy").innerText = vals[colMap.maxRolls] || "-";
            document.getElementById("val-kpl").innerText = vals[colMap.expLayers] || "0";
            
            let lastDeclaredRolls = vals[colMap.maxRolls] || "";
            document.getElementById("in-operator").value = "";
            document.getElementById("in-workers").value = "4";
            
            selectedMachineForContinuation = vals[colMap.machine] ? vals[colMap.machine].toString() : "";
            currentWorkerGlobalString = vals[colMap.workers] ? vals[colMap.workers].toString() : "";
            previousGlobalWorkerString = currentWorkerGlobalString; // Zachowaj historię
            currentOperatorGlobalString = vals[colMap.operator] ? vals[colMap.operator].toString() : "";
            totalAwariaSecondsGlobal = hmsToSeconds(vals[colMap.awarie] ? vals[colMap.awarie].toString() : "00:00:00");
            
            const existingNotes = vals[colMap.notes] ? vals[colMap.notes].toString() : "";
            document.getElementById("in-other-incidents").value = existingNotes;
            document.getElementById("in-running-notes").value = existingNotes;
            
            // Obliczamy wcześniejsze czasy z przedziałów i skanujemy pod kątem przerwanych
            previousTotalGrossSeconds = 0;
            lastIntervalUnexpected = false;
            lastIntervalIndex = -1;
            unexpectedIntervalDuration = 0;
            
            
            if (colMap.intervalsStart !== undefined) {
                const intervalsStartStatus = vals[colMap.intervalsStart] ? vals[colMap.intervalsStart].toString().trim() : "";
                for (let i = 0; i < 10; i++) {
                    const startIdx = colMap.intervalsStart + 1 + (i * 6) + 4;
                    const stopIdx = colMap.intervalsStart + 1 + (i * 6) + 5;
                    const rlsIdx = colMap.intervalsStart + 1 + (i * 6) + 3;
                    
                    const startStr = vals[startIdx] ? vals[startIdx].toString().replace(/^'/, "") : "";
                    const stopStr = vals[stopIdx] ? vals[stopIdx].toString().replace(/^'/, "") : "";
                    
                    if (startStr) {
                        lastIntervalIndex = i;
                        const rls = vals[rlsIdx];
                        if (rls) lastDeclaredRolls = rls;
                    }
                    
                    if (startStr && stopStr) {
                        const tStart = parseCustomDate(startStr);
                        const tStop = parseCustomDate(stopStr);
                        if (!isNaN(tStart) && !isNaN(tStop)) {
                            const durationMs = tStop - tStart;
                            if (durationMs > 0) {
                                previousTotalGrossSeconds += Math.floor(durationMs / 1000);
                            }
                        }
                    }
                }
                
                if (lastIntervalIndex >= 0) {
                    if (intervalsStartStatus !== "ZAMKNIĘTO") {
                        lastIntervalUnexpected = true;
                        // Odejmujemy czas tego przedziału od zsumowanego gross
                        const uStartStr = vals[colMap.intervalsStart + 1 + (lastIntervalIndex * 6) + 4] ? vals[colMap.intervalsStart + 1 + (lastIntervalIndex * 6) + 4].toString().replace(/^'/, "") : "";
                        const uStopStr = vals[colMap.intervalsStart + 1 + (lastIntervalIndex * 6) + 5] ? vals[colMap.intervalsStart + 1 + (lastIntervalIndex * 6) + 5].toString().replace(/^'/, "") : "";
                        if (uStartStr && uStopStr) {
                            const tuStart = parseCustomDate(uStartStr);
                            const tuStop = parseCustomDate(uStopStr);
                            if (!isNaN(tuStart) && !isNaN(tuStop) && (tuStop - tuStart > 0)) {
                                unexpectedIntervalDuration = Math.floor((tuStop - tuStart) / 1000);
                            }
                        }
                    } else {
                        lastIntervalUnexpected = false;
                    }
                }
            }
            
            document.getElementById("in-real-rolls").value = lastDeclaredRolls;
            updateKitsCalc();
            
            theoreticalSeconds = 0;
            if (colMap.theoretical !== undefined) {
                const theoVal = vals[colMap.theoretical] ? vals[colMap.theoretical].toString().trim() : "";
                if (theoVal.includes(":")) {
                    theoreticalSeconds = hmsToSeconds(theoVal);
                } else {
                    const num = parseFloat(theoVal);
                    if (!isNaN(num)) theoreticalSeconds = num * 24 * 3600;
                }
            }
            
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

function handleToMachineClick() {
    if (isContinuing && lastIntervalUnexpected) {
        document.getElementById("data-card").classList.add("hidden");
        document.getElementById("unexpected-card").classList.remove("hidden");
    } else {
        resumeUnexpected = false;
        showMachineSelection();
    }
}

document.getElementById("btn-unexp-finished").onclick = async () => {
    resumeUnexpected = false;
    try {
        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getItem(activeSheetName);
            sheet.getCell(currentRowIndex, colMap.intervalsStart).values = [["ZAMKNIĘTO"]];
            await ctx.sync();
        });
        document.getElementById("unexpected-card").classList.add("hidden");
        showMachineSelection();
    } catch (e) {
        setStatus("Błąd dopisywania ZAMKNIĘTO: " + e.message);
    }
};

document.getElementById("btn-unexp-resume").onclick = () => {
    resumeUnexpected = true;
    document.getElementById("unexpected-card").classList.add("hidden");
    writeStartTime();
};

document.getElementById("btn-unexp-cancel").onclick = () => {
    document.getElementById("unexpected-card").classList.add("hidden");
    document.getElementById("initial-card").classList.remove("hidden");
};

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
                
                if (isContinuing && selectedMachineForContinuation && s.name === selectedMachineForContinuation) {
                    opt.selected = true;
                } else if (!isContinuing && s.name === activeSheetName) {
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

function handleStartTimerClick() {
    const machine = document.getElementById("sel-machine").value;
    
    // Weryfikacja innej otwartej pracy na tej samej maszynie
    if (window.unfinishedMachines && window.unfinishedMachines[machine]) {
        const unfin = window.unfinishedMachines[machine];
        if (unfin.row !== currentRowIndex) {
            document.getElementById("machine-warning-text").innerText = `Na maszynie ${machine} znajduje się już niezakończony produkt (Wiersz ${unfin.row + 1}: Item ${unfin.item}).\n\nCzy na pewno chcesz rozpocząć nowe cięcie na tej maszynie?`;
            document.getElementById("machine-card").classList.add("hidden");
            document.getElementById("machine-warning-card").classList.remove("hidden");
            return;
        }
    }
    writeStartTime();
}

document.getElementById("btn-mach-warn-continue").onclick = () => {
    document.getElementById("machine-warning-card").classList.add("hidden");
    writeStartTime();
};

document.getElementById("btn-mach-warn-cancel").onclick = () => {
    resetUI();
};

async function writeStartTime() {
    const operator = document.getElementById("in-operator").value.trim() || "Brak";
    startWorkersCount = parseInt(document.getElementById("in-workers").value) || 4;
    currentWorkersCount = startWorkersCount;
    document.getElementById("val-current-workers").innerText = currentWorkersCount;
    
    const realRolls = document.getElementById("in-real-rolls").value;
    let machine = document.getElementById("sel-machine").value;
    
    if (!machine && resumeUnexpected) {
        machine = selectedMachineForContinuation;
    }
    
    try {
        setStatus("Rozpoczynanie...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(activeSheetName);
            const dateStr = getFormattedDate();
            
            // Global Strings Updates
            if (!resumeUnexpected) {
                if (currentOperatorGlobalString === "") {
                    currentOperatorGlobalString = operator;
                } else {
                    currentOperatorGlobalString += "/" + operator;
                }
                
                intervalWorkerDiff = 0; // reset różnicy dla nowego przedziału
                if (previousGlobalWorkerString === "") {
                    currentWorkerGlobalString = startWorkersCount.toString();
                } else {
                    currentWorkerGlobalString = previousGlobalWorkerString + "/" + startWorkersCount.toString();
                }
                
                sheet.getCell(currentRowIndex, colMap.operator).values = [[safeStr(currentOperatorGlobalString)]];
                sheet.getCell(currentRowIndex, colMap.workers).values = [[safeStr(currentWorkerGlobalString)]];
            }
            
            if (!isContinuing) {
                sheet.getCell(currentRowIndex, colMap.startGlobal).values = [[safeStr(dateStr)]];
            }
            sheet.getCell(currentRowIndex, colMap.machine).values = [[machine]];
            sheet.getCell(currentRowIndex, colMap.intervalsStart).values = [[""]]; // Wyczyszczenie ZAMKNIĘTO
            
            // Znajdowanie przedziału
            if (resumeUnexpected && lastIntervalIndex >= 0) {
                currentIntervalIndex = lastIntervalIndex;
                currentIntervalStartCol = colMap.intervalsStart + 1 + (currentIntervalIndex * 6);
                
                // Read Start Time to calculate secondsElapsed accurately
                const startCell = sheet.getCell(currentRowIndex, currentIntervalStartCol + 4).load("values");
                await context.sync();
                const sVal = startCell.values[0][0] ? startCell.values[0][0].toString().replace(/^'/, "") : getFormattedDate();
                const tStart = parseCustomDate(sVal);
                if (!isNaN(tStart)) {
                    secondsElapsed = Math.floor((new Date().getTime() - tStart) / 1000);
                    if (secondsElapsed < 0) secondsElapsed = 0;
                } else {
                    secondsElapsed = 0;
                }
                
                // Odejmujemy czas nieoczekiwanego zamknięcia z previousTotalGrossSeconds, żeby nie policzyć go podwójnie!
                previousTotalGrossSeconds -= unexpectedIntervalDuration;
                if (previousTotalGrossSeconds < 0) previousTotalGrossSeconds = 0;
                
            } else {
                const intervalsRange = sheet.getRangeByIndexes(currentRowIndex, colMap.intervalsStart + 1, 1, 60).load("values");
                await context.sync();
                
                currentIntervalIndex = 9; // domyślnie ostatni jeśli wszystkie zajęte
                for (let i = 0; i < 10; i++) {
                    const startCellVal = intervalsRange.values[0][i * 6 + 4] ? intervalsRange.values[0][i * 6 + 4].toString().trim() : "";
                    if (startCellVal === "") {
                        currentIntervalIndex = i;
                        break;
                    }
                }
                currentIntervalStartCol = colMap.intervalsStart + 1 + (currentIntervalIndex * 6);
                sheet.getCell(currentRowIndex, currentIntervalStartCol + 4).values = [[safeStr(dateStr)]];
                secondsElapsed = 0;
            }
            
            // Zapis przedziału (Operator, Pracownicy Start, Pracownicy End (na razie = Start), Rolki, Start, Puste)
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 0).values = [[operator]];
            if (!resumeUnexpected) {
                sheet.getCell(currentRowIndex, currentIntervalStartCol + 1).values = [[startWorkersCount]];
            }
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 2).values = [[currentWorkersCount]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 3).values = [[realRolls]];
            await context.sync();
            
            const iTxt = document.getElementById("val-item").innerText;
            const rTxt = document.getElementById("val-rev").innerText;
            const pTxt = document.getElementById("val-product").innerText;
            document.getElementById("running-info").innerHTML = `Item: ${iTxt}, Rev: ${rTxt}<br>Prod: ${pTxt}`;
            document.getElementById("running-rolls-display").innerText = realRolls;
            
            document.getElementById("machine-card").classList.add("hidden");
            document.getElementById("running-card").classList.remove("hidden");
            document.body.classList.add("timer-active"); // Tło zielone!
            document.getElementById("main-header").classList.add("hidden");
            
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
    if (!resumeUnexpected) secondsElapsed = 0;
    updateTimerDisplay();
    timerInterval = setInterval(() => {
        secondsElapsed++;
        updateTimerDisplay();
    }, 1000);
}

function updateTimerDisplay() {
    document.getElementById("timer").innerText = secondsToHms(secondsElapsed);
    
    let totalGrossSecondsForTarget = previousTotalGrossSeconds + secondsElapsed;
    
    const targetValue = document.getElementById("target-value");
    const targetDetails = document.getElementById("target-details");
    
    // Aktualizacja UI targetu tylko co 3 minuty (180s) lub na starcie
    if (secondsElapsed % 180 === 0 || secondsElapsed === 0 || targetValue.innerText === "START..." || targetValue.innerText === "Brak") {
        if (theoreticalSeconds > 0 && totalGrossSecondsForTarget > 0) {
            const targetPct = (theoreticalSeconds / totalGrossSecondsForTarget) * 100;
            targetValue.innerText = Math.round(targetPct) + " %";
            
            if (targetPct >= 100) {
                targetValue.style.color = "#16a34a"; // green
            } else if (targetPct >= 80) {
                targetValue.style.color = "#d97706"; // yellow/orange
            } else {
                targetValue.style.color = "#dc2626"; // red
            }
        } else if (theoreticalSeconds > 0 && totalGrossSecondsForTarget === 0) {
            targetValue.innerText = "START...";
            targetValue.style.color = "#6b7280";
        } else {
            targetValue.innerText = "Brak";
            targetValue.style.color = "#6b7280";
        }

        if (theoreticalSeconds > 0) {
            targetDetails.innerText = `Teor: ${secondsToHms(theoreticalSeconds)} | Brutto: ${secondsToHms(totalGrossSecondsForTarget)}`;
        } else {
            targetDetails.innerText = "Brak zdefiniowanego czasu teoretycznego.";
        }
    }
}

function startAutoSave() {
    autoSaveInterval = setInterval(async () => {
        if (currentRowIndex !== -1 && currentIntervalStartCol !== -1) {
            try {
                await Excel.run(async (ctx) => {
                    const sheet = ctx.workbook.worksheets.getItem(activeSheetName);
                    sheet.getCell(currentRowIndex, currentIntervalStartCol + 5).values = [[safeStr(getFormattedDate())]];
                    
                    if (colMap.notes !== undefined) {
                        const notesVal = document.getElementById("in-running-notes").value;
                        sheet.getCell(currentRowIndex, colMap.notes).values = [[notesVal]];
                    }
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
    
    intervalWorkerDiff += amount;
    
    // Budujemy string "4+3"
    let diffStr = "";
    if (intervalWorkerDiff > 0) {
        diffStr = "+" + intervalWorkerDiff;
    } else if (intervalWorkerDiff < 0) {
        diffStr = intervalWorkerDiff.toString(); // ma znak minus w sobie
    }
    
    if (previousGlobalWorkerString === "") {
        currentWorkerGlobalString = startWorkersCount.toString() + diffStr;
    } else {
        currentWorkerGlobalString = previousGlobalWorkerString + "/" + startWorkersCount.toString() + diffStr;
    }
    
    // Auto update excel values (End workers + Global string)
    Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(activeSheetName);
        sheet.getCell(currentRowIndex, colMap.workers).values = [[safeStr(currentWorkerGlobalString)]];
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
        document.body.classList.remove("timer-active"); // Usuwa zielony
        btn.innerText = "ZAKOŃCZ STAN AWARII";
        timerUI.classList.remove("hidden");
        
        awariaSecondsElapsed = 0;
        timerUI.innerText = secondsToHms(0);
        updateTimerDisplay(); // aktualizacja by zamrozić netto
        awariaTimerInterval = setInterval(() => {
            awariaSecondsElapsed++;
            timerUI.innerText = secondsToHms(awariaSecondsElapsed);
            updateTimerDisplay(); // aktualizacja by zamrozić netto w trakcie awarii
        }, 1000);
    } else {
        document.body.classList.remove("awaria-active");
        document.body.classList.add("timer-active"); // Przywraca zielony
        btn.innerText = "STAN AWARII";
        timerUI.classList.add("hidden");
        clearInterval(awariaTimerInterval);
        
        totalAwariaSecondsGlobal += awariaSecondsElapsed;
        
        // Zapisz sumę awarii do excela
        Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getItem(activeSheetName);
            sheet.getCell(currentRowIndex, colMap.awarie).values = [[safeStr(secondsToHms(totalAwariaSecondsGlobal))]];
            await ctx.sync();
        }).catch(e => console.warn(e));
    }
}

async function confirmChangeRolls() {
    const newRollsStr = document.getElementById("in-new-rolls").value;
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
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 5).values = [[safeStr(dateStr)]];
            
            // Ustal nowy index przedziału
            currentIntervalIndex++;
            if (currentIntervalIndex > 9) currentIntervalIndex = 9; // Overwrite last
            currentIntervalStartCol = colMap.intervalsStart + 1 + (currentIntervalIndex * 6);
            
            // Ponieważ zaczynamy nowy przedział, modyfikujemy TYLKO zapis startu przedziału. 
            // NIE modyfikujemy stringa globalnego pracowników, bo to jest "w locie".
            // Nie resetujemy intervalWorkerDiff. Zapis przedziału łapie po prostu obecny stan jako "Workers Start".
            
            // Zapis nowego przedziału
            const operator = document.getElementById("in-operator").value.trim() || "Brak";
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 0).values = [[operator]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 1).values = [[currentWorkersCount]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 2).values = [[currentWorkersCount]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 3).values = [[newRolls]];
            sheet.getCell(currentRowIndex, currentIntervalStartCol + 4).values = [[safeStr(dateStr)]];
            await ctx.sync();
            
            document.getElementById("in-real-rolls").value = newRolls;
            updateKitsCalc();
            
            const iTxt = document.getElementById("val-item").innerText;
            const rTxt = document.getElementById("val-rev").innerText;
            const pTxt = document.getElementById("val-product").innerText;
            document.getElementById("running-info").innerHTML = `Item: ${iTxt}, Rev: ${rTxt}<br>Prod: ${pTxt}`;
            document.getElementById("running-rolls-display").innerText = newRolls;
            
            setStatus("Zmieniono rolki. Przedział rozdzielony.");
            document.getElementById("change-rolls-panel").classList.add("hidden");
            document.getElementById("in-new-rolls").value = "";
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
    timerInterval = null;
    stopAutoSave(); 
    
    document.getElementById("in-other-incidents").value = document.getElementById("in-running-notes").value;
    
    document.getElementById("running-card").classList.add("hidden");
    document.getElementById("incidents-card").classList.add("hidden");
    document.getElementById("main-header").classList.remove("hidden");
    document.getElementById("incidents-card").classList.remove("hidden");
    document.body.classList.remove("timer-active"); // Usuń zielone tło
    setStatus("Czas zatrzymany. Wybierz opcję zakończenia.");
}

async function saveIncidents(fullComplete) {
    const material = document.getElementById("chk-material").checked;
    const breakTime = document.getElementById("chk-break").checked;
    const incidentsText = document.getElementById("in-other-incidents").value;
    
    try {
        setStatus("Zapisywanie...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(activeSheetName);
            const dateStr = getFormattedDate();
            
            if (currentIntervalStartCol !== -1) {
                sheet.getCell(currentRowIndex, currentIntervalStartCol + 5).values = [[safeStr(dateStr)]]; // Stop Time
            }
            
            // Oznaczenie poprawnego zamknięcia sesji
            sheet.getCell(currentRowIndex, colMap.intervalsStart).values = [["ZAMKNIĘTO"]];
            
            if (fullComplete) {
                sheet.getCell(currentRowIndex, colMap.endGlobal).values = [[safeStr(dateStr)]];
                
                // --- PODSUMOWANIE (3 kolumny na samym końcu przedziałów = intervalsStart + 60) ---
                const dataRange = sheet.getRangeByIndexes(currentRowIndex, colMap.intervalsStart + 1, 1, 60).load("values");
                await context.sync();
                
                const ivals = dataRange.values[0];
                let totalTimeMs = 0;
                let totalKits = 0;
                let sumWorkerTime = 0;
                
                const kitsPerLayer = parseFloat(document.getElementById("val-kpl").textContent) || 0;
                
                for (let i = 0; i < 10; i++) {
                    const wStart = parseFloat(ivals[i*6 + 1]);
                    const wStop = parseFloat(ivals[i*6 + 2]);
                    const rolls = parseFloat(ivals[i*6 + 3]) || 0;
                    const startStr = ivals[i*6 + 4] ? ivals[i*6 + 4].toString().replace(/^'/, "") : "";
                    const stopStr = ivals[i*6 + 5] ? ivals[i*6 + 5].toString().replace(/^'/, "") : "";
                    
                    if (startStr && rolls > 0) {
                        if (i === 0 || rolls !== parseFloat(ivals[(i-1)*6 + 3])) {
                            totalKits += (rolls * kitsPerLayer);
                        }
                    }

                    if (startStr && stopStr) {
                        const tStart = parseCustomDate(startStr);
                        const tStop = parseCustomDate(stopStr);
                        if (!isNaN(tStart) && !isNaN(tStop)) {
                            const durationMs = tStop - tStart;
                            if (durationMs > 0) {
                                totalTimeMs += durationMs;
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
                
                // Write summary to 3 columns BEFORE intervalsStart
                const summaryStartCol = colMap.intervalsStart - 3;
                if (summaryStartCol >= 0) {
                    sheet.getCell(dataStartRowIndex - 1, summaryStartCol).values = [["CZAS NETTO"]];
                    sheet.getCell(dataStartRowIndex - 1, summaryStartCol + 1).values = [["SUMA KITÓW"]];
                    sheet.getCell(dataStartRowIndex - 1, summaryStartCol + 2).values = [["ŚREDNIA PRACOWNIKÓW"]];
                    
                    // Wymuszamy format liczbowy/ogólny w Excelu dla wyników
                    const sumRange = sheet.getRangeByIndexes(currentRowIndex, summaryStartCol, 1, 3);
                    sumRange.numberFormat = [["@", "0", "0.00"]];
                    
                    sheet.getCell(currentRowIndex, summaryStartCol).values = [[safeStr(netTimeHms)]];
                    sheet.getCell(currentRowIndex, summaryStartCol + 1).values = [[Math.round(totalKits)]];
                    sheet.getCell(currentRowIndex, summaryStartCol + 2).values = [[Number(avgWorkersFinal.toFixed(2))]];
                }
                
                // Generowanie nagłówków i obramowań dla wykorzystanych przedziałów
                if (colMap.intervalsStart !== undefined && currentIntervalIndex >= 0) {
                    const usedColsCount = (currentIntervalIndex + 1) * 6;
                    for (let i = 0; i <= currentIntervalIndex; i++) {
                        const sCol = colMap.intervalsStart + 1 + (i * 6);
                        sheet.getCell(dataStartRowIndex - 1, sCol + 0).values = [[`Operator ${i+1}`]];
                        sheet.getCell(dataStartRowIndex - 1, sCol + 1).values = [[`Prac. Start ${i+1}`]];
                        sheet.getCell(dataStartRowIndex - 1, sCol + 2).values = [[`Prac. Koniec ${i+1}`]];
                        sheet.getCell(dataStartRowIndex - 1, sCol + 3).values = [[`Rolki ${i+1}`]];
                        sheet.getCell(dataStartRowIndex - 1, sCol + 4).values = [[`Start ${i+1}`]];
                        sheet.getCell(dataStartRowIndex - 1, sCol + 5).values = [[`Koniec ${i+1}`]];
                    }
                    
                    const intervalHeaderRange = sheet.getRangeByIndexes(dataStartRowIndex - 1, colMap.intervalsStart + 1, 1, usedColsCount);
                    intervalHeaderRange.format.borders.getItem('EdgeTop').style = 'Continuous';
                    intervalHeaderRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
                    intervalHeaderRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
                    intervalHeaderRange.format.borders.getItem('EdgeRight').style = 'Continuous';
                    intervalHeaderRange.format.borders.getItem('InsideVertical').style = 'Continuous';
                    intervalHeaderRange.format.borders.color = "#a3a3a3";
                    intervalHeaderRange.format.borders.weight = "Thin";
                    
                    const intervalDataRange = sheet.getRangeByIndexes(currentRowIndex, colMap.intervalsStart + 1, 1, usedColsCount);
                    intervalDataRange.format.borders.getItem('EdgeTop').style = 'Continuous';
                    intervalDataRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
                    intervalDataRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
                    intervalDataRange.format.borders.getItem('EdgeRight').style = 'Continuous';
                    intervalDataRange.format.borders.getItem('InsideVertical').style = 'Continuous';
                    intervalDataRange.format.borders.color = "#a3a3a3";
                    intervalDataRange.format.borders.weight = "Thin";
                }
            }
            
            if (colMap.chkMat !== undefined) sheet.getCell(currentRowIndex, colMap.chkMat).values = [[material ? "TAK" : ""]];
            if (colMap.chkBreak !== undefined) sheet.getCell(currentRowIndex, colMap.chkBreak).values = [[breakTime ? "TAK" : ""]];
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
    document.body.classList.remove("timer-active");
    document.getElementById("awaria-timer").classList.add("hidden");
    document.getElementById("btn-awaria").innerText = "STAN AWARII";
    document.getElementById("change-rolls-panel").classList.add("hidden");
    
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
