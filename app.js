Office.onReady((info) => {
    document.getElementById("btn-fetch").onclick = () => fetchRowData(null, false);
    document.getElementById("btn-cancel-data").onclick = resetUI;
    document.getElementById("btn-to-machine").onclick = showMachineSelection;
    document.getElementById("btn-start-timer").onclick = writeStartTime;
    document.getElementById("btn-stop").onclick = handleStop;
    document.getElementById("btn-save-partial").onclick = () => saveIncidents(false);
    document.getElementById("btn-save-full").onclick = () => saveIncidents(true);
    
    if (info.host === Office.HostType.Excel) {
        setStatus("Ładowanie listy niezakończonych...");
        scanForUnfinished();
    } else {
        setStatus("UWAGA: Uruchom ten link jako Dodatek Wewnątrz Excela!");
    }
});

let currentRowIndex = -1;
let timerInterval = null;
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

async function scanForUnfinished() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange("A1:AB1000"); // Skanujemy do wiersza 1000
            range.load("values");
            await context.sync();
            
            const listContainer = document.getElementById("unfinished-list");
            listContainer.innerHTML = "";
            let foundAny = false;
            
            for (let i = 1; i < range.values.length; i++) { // Pomijamy nagłówek (0)
                const row = range.values[i];
                if (!row) continue;
                
                const valAA = row[26] ? row[26].toString().trim() : "";
                const valAB = row[27] ? row[27].toString().trim() : "";
                
                if (valAA !== "" && valAB === "") {
                    // Niezakończony proces!
                    foundAny = true;
                    const itemValue = row[1] ? row[1].toString() : "Brak Itemu"; // B(1)
                    
                    const btn = document.createElement("button");
                    btn.className = "unfinished-item";
                    btn.innerText = `Wiersz ${i + 1} | Item: ${itemValue}`;
                    btn.onclick = () => fetchRowData(i, true);
                    listContainer.appendChild(btn);
                }
            }
            
            if (!foundAny) {
                listContainer.innerHTML = `<div style="font-size:12px; color:#9ca3af;">Brak niezakończonych zadań w tym arkuszu.</div>`;
            }
            
            document.getElementById("unfinished-container").style.display = "block";
            setStatus("Gotowe. Zaznacz wiersz lub wybierz z listy.");
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd skanowania: " + error.message);
    }
}

async function fetchRowData(forcedRowIndex, isCont) {
    isContinuing = isCont;
    try {
        setStatus("Pobieranie danych...");
        await Excel.run(async (context) => {
            let rowIdx = forcedRowIndex;
            if (rowIdx === null) {
                const activeCell = context.workbook.getActiveCell();
                activeCell.load("rowIndex");
                await context.sync();
                rowIdx = activeCell.rowIndex;
                
                // Sprawdź czy to nowy proces czy kontynuacja (żeby nie pozwolić na zepsucie danych)
                const checkAA = context.workbook.worksheets.getActiveWorksheet().getCell(rowIdx, 26).load("values");
                const checkAB = context.workbook.worksheets.getActiveWorksheet().getCell(rowIdx, 27).load("values");
                await context.sync();
                
                const valAA = checkAA.values[0][0] ? checkAA.values[0][0].toString().trim() : "";
                const valAB = checkAB.values[0][0] ? checkAB.values[0][0].toString().trim() : "";
                
                if (valAB !== "") {
                    setStatus("UWAGA: Ten produkt został już całkowicie zakończony!");
                    return;
                }
                if (valAA !== "") {
                    isContinuing = true; // Auto-wykrycie kontynuacji
                }
            }
            
            currentRowIndex = rowIdx;
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Pobieramy szerszy zakres dla danego wiersza, od A do CC (index 0 do 80)
            const rowRange = sheet.getRangeByIndexes(currentRowIndex, 0, 1, 81).load("values");
            await context.sync();
            const vals = rowRange.values[0];
            
            // B(1), C(2), D(3), E(4), M(12), O(14)
            document.getElementById("val-item").innerText = vals[1] || "-";
            document.getElementById("val-rev").innerText = vals[2] || "-";
            document.getElementById("val-product").innerText = vals[3] || "-";
            document.getElementById("val-nesting").innerText = vals[4] || "-";
            document.getElementById("val-warstwy").innerText = vals[12] || "-";
            document.getElementById("val-kit").innerText = vals[14] || "-";
            
            document.getElementById("in-real-layers").value = vals[12] || "";
            document.getElementById("in-operator").value = "";
            document.getElementById("in-workers").value = "4";
            
            // Wyczyść incydenty
            document.getElementById("chk-material").checked = false;
            document.getElementById("chk-break").checked = false;
            document.getElementById("chk-breakdown").checked = false;
            document.getElementById("in-other-incidents").value = "";
            
            if (isContinuing) {
                document.getElementById("btn-to-machine").innerText = "KONTYNUUJ PROCES";
                // Załaduj stare dane do podglądu (Y, Z, BP, incydenty)
                if (vals[24]) document.getElementById("in-operator").value = vals[24];
                if (vals[25]) document.getElementById("in-workers").value = vals[25];
                if (vals[67]) document.getElementById("in-real-layers").value = vals[67];
                
                if (vals[68] === "TAK") document.getElementById("chk-material").checked = true;
                if (vals[69] === "TAK") document.getElementById("chk-break").checked = true;
                if (vals[70] === "TAK") document.getElementById("chk-breakdown").checked = true;
                if (vals[36]) document.getElementById("in-other-incidents").value = vals[36];
            } else {
                document.getElementById("btn-to-machine").innerText = "DALEJ";
            }
            
            // Zmiana widoku
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
                if (sheet.name === activeSheet.name) {
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
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const dateStr = getFormattedDate();
            
            if (!isContinuing) {
                // Zapisz dane tylko na starcie globalnym
                sheet.getCell(currentRowIndex, 24).values = [[operator]];
                sheet.getCell(currentRowIndex, 25).values = [[workers]];
                sheet.getCell(currentRowIndex, 67).values = [[realLayers]];
                sheet.getCell(currentRowIndex, 26).values = [[dateStr]]; // Global Start (AA)
            }
            
            // Zawsze zapisz wybraną maszynę
            sheet.getCell(currentRowIndex, 71).values = [[machine]]; // BT
            
            // Szukanie wolnego przedziału
            // BW(74), BX(75) | BY(76), BZ(77) | CA(78), CB(79)
            const intervalsRange = sheet.getRangeByIndexes(currentRowIndex, 74, 1, 5).load("values");
            await context.sync();
            
            const iv = intervalsRange.values[0];
            const vBW = iv[0] ? iv[0].toString().trim() : "";
            const vBY = iv[2] ? iv[2].toString().trim() : "";
            const vCA = iv[4] ? iv[4].toString().trim() : "";
            
            if (vBW === "") {
                currentIntervalStartCol = 74; currentIntervalEndCol = 75;
            } else if (vBY === "") {
                currentIntervalStartCol = 76; currentIntervalEndCol = 77;
            } else if (vCA === "") {
                currentIntervalStartCol = 78; currentIntervalEndCol = 79;
            } else {
                // Nadpisuje trzeci i dodaje notatkę w CC (80)
                currentIntervalStartCol = 78; currentIntervalEndCol = 79;
                sheet.getCell(currentRowIndex, 80).values = [["Proces wznawiano więcej niż 3 razy"]];
            }
            
            // Wpisanie czasu Start do przedziału
            sheet.getCell(currentRowIndex, currentIntervalStartCol).values = [[dateStr]];
            await context.sync();
            
            // Ustaw tekst info
            const iTxt = document.getElementById("val-item").innerText;
            const rTxt = document.getElementById("val-rev").innerText;
            const pTxt = document.getElementById("val-product").innerText;
            const nTxt = document.getElementById("val-nesting").innerText;
            document.getElementById("running-info").innerHTML = `Item: ${iTxt}, Rev: ${rTxt}<br>Prod: ${pTxt}<br>Nesting: ${nTxt}`;
            
            // Przejście widoku do stopera
            document.getElementById("machine-card").classList.add("hidden");
            document.getElementById("running-card").classList.remove("hidden");
            
            startTimer();
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

function handleStop() {
    clearInterval(timerInterval);
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
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const dateStr = getFormattedDate();
            
            // Zapisz do aktualnego przedziału
            if (currentIntervalEndCol !== -1) {
                sheet.getCell(currentRowIndex, currentIntervalEndCol).values = [[dateStr]];
            }
            
            // Jeżeli kończymy całkowicie
            if (fullComplete) {
                sheet.getCell(currentRowIndex, 27).values = [[dateStr]]; // AB Global End
            }
            
            // BQ(68)=Materiał, BR(69)=Przerwa, BS(70)=Awaria, AK(36)=Inne
            // Jeśli puste, nie nadpisujemy "TAK", tylko sprawdzamy co jest
            if (material) sheet.getCell(currentRowIndex, 68).values = [["TAK"]];
            else sheet.getCell(currentRowIndex, 68).values = [[""]];
            
            if (breakTime) sheet.getCell(currentRowIndex, 69).values = [["TAK"]];
            else sheet.getCell(currentRowIndex, 69).values = [[""]];
            
            if (breakdown) sheet.getCell(currentRowIndex, 70).values = [["TAK"]];
            else sheet.getCell(currentRowIndex, 70).values = [[""]];
            
            sheet.getCell(currentRowIndex, 36).values = [[incidentsText]];
            
            await context.sync();
            
            resetUI();
            scanForUnfinished(); // Odśwież listę
            setStatus(fullComplete ? "Zakończono produkt pomyślnie!" : "Przerwano produkt. Zapisano zmianę.");
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd zapisu: " + error.message);
    }
}

function resetUI() {
    clearInterval(timerInterval);
    document.getElementById("data-card").classList.add("hidden");
    document.getElementById("machine-card").classList.add("hidden");
    document.getElementById("running-card").classList.add("hidden");
    document.getElementById("incidents-card").classList.add("hidden");
    document.getElementById("initial-card").classList.remove("hidden");
    setStatus("Gotowe. Zaznacz wiersz lub wybierz z listy.");
}

function setStatus(message) {
    const statusEl = document.getElementById("status-message");
    if (statusEl) {
        statusEl.innerText = message;
    }
}
