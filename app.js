Office.onReady((info) => {
    // Przypisanie zdarzeń niezależnie od hosta, aby w razie błędu pokazać komunikat
    document.getElementById("btn-fetch").onclick = fetchRowData;
    document.getElementById("btn-start").onclick = writeStartTime;
    document.getElementById("btn-stop").onclick = handleStop;
    document.getElementById("btn-save").onclick = saveIncidents;
    
    if (info.host === Office.HostType.Excel) {
        setStatus("Gotowe. Zaznacz wiersz w Excelu.");
    } else {
        setStatus("UWAGA: Uruchom ten link jako Dodatek Wewnątrz Excela!");
    }
});

let currentRowIndex = -1;
let timerInterval = null;
let secondsElapsed = 0;

function getFormattedDate() {
    const d = new Date();
    const datePart = `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
    let hours = d.getHours();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12;
    const minutes = d.getMinutes().toString().padStart(2, '0');
    const seconds = d.getSeconds().toString().padStart(2, '0');
    // Używamy spacji zgodnie z wymogiem np. 5/28/2026  2:27:08 AM
    return `${datePart} ${hours}:${minutes}:${seconds} ${ampm}`;
}

async function fetchRowData() {
    try {
        setStatus("Pobieranie danych...");
        await Excel.run(async (context) => {
            const activeCell = context.workbook.getActiveCell();
            activeCell.load("rowIndex");
            await context.sync();
            
            currentRowIndex = activeCell.rowIndex;
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Odczyt: B(1), C(2), D(3), E(4), M(12), O(14)
            const rangeB = sheet.getCell(currentRowIndex, 1).load("values");
            const rangeC = sheet.getCell(currentRowIndex, 2).load("values");
            const rangeD = sheet.getCell(currentRowIndex, 3).load("values");
            const rangeE = sheet.getCell(currentRowIndex, 4).load("values");
            const rangeM = sheet.getCell(currentRowIndex, 12).load("values");
            const rangeO = sheet.getCell(currentRowIndex, 14).load("values");
            
            await context.sync();
            
            document.getElementById("val-item").innerText = rangeB.values[0][0] || "-";
            document.getElementById("val-rev").innerText = rangeC.values[0][0] || "-";
            document.getElementById("val-product").innerText = rangeD.values[0][0] || "-";
            document.getElementById("val-nesting").innerText = rangeE.values[0][0] || "-";
            document.getElementById("val-warstwy").innerText = rangeM.values[0][0] || "-";
            document.getElementById("val-kit").innerText = rangeO.values[0][0] || "-";
            
            // Ustaw domyślną wartość Rzeczywistych Warstw z kolumny M
            document.getElementById("in-real-layers").value = rangeM.values[0][0] || "";
            
            // Pokaż sekcję z pobranymi danymi
            document.getElementById("data-section").classList.remove("hidden");
            
            setStatus(`Pobrano wiersz ${currentRowIndex + 1}.`);
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd: " + error.message);
    }
}

async function writeStartTime() {
    if (currentRowIndex === -1) return;
    
    const operator = document.getElementById("in-operator").value;
    const workers = document.getElementById("in-workers").value;
    const realLayers = document.getElementById("in-real-layers").value;
    
    try {
        setStatus("Rozpoczynanie...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Zapis: Y(24)=Operator, Z(25)=Pracownicy, BP(67)=Rzecz.Warstwy, AA(26)=Start Czas
            sheet.getCell(currentRowIndex, 24).values = [[operator]];
            sheet.getCell(currentRowIndex, 25).values = [[workers]];
            sheet.getCell(currentRowIndex, 67).values = [[realLayers]];
            sheet.getCell(currentRowIndex, 26).values = [[getFormattedDate()]];
            
            await context.sync();
            
            // Przejście widoku do stopera
            document.getElementById("btn-fetch").classList.add("hidden");
            document.getElementById("data-section").classList.add("hidden");
            document.getElementById("running-section").classList.remove("hidden");
            
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

async function handleStop() {
    clearInterval(timerInterval);
    try {
        setStatus("Zatrzymano. Zapis czasu...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            // Zapisz czas stopu do kolumny AB(27)
            sheet.getCell(currentRowIndex, 27).values = [[getFormattedDate()]];
            await context.sync();
            
            // Przejście widoku do incydentów
            document.getElementById("running-section").classList.add("hidden");
            document.getElementById("incidents-section").classList.remove("hidden");
            
            setStatus("Czas zapisany. Uzupełnij incydenty.");
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd Stop: " + error.message);
    }
}

async function saveIncidents() {
    const material = document.getElementById("chk-material").checked;
    const breakTime = document.getElementById("chk-break").checked;
    const breakdown = document.getElementById("chk-breakdown").checked;
    const incidentsText = document.getElementById("in-other-incidents").value;
    
    try {
        setStatus("Zapisywanie incydentów...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // BQ(68)=Materiał, BR(69)=Przerwa, BS(70)=Awaria, AK(36)=Inne
            if (material) sheet.getCell(currentRowIndex, 68).values = [["TAK"]];
            if (breakTime) sheet.getCell(currentRowIndex, 69).values = [["TAK"]];
            if (breakdown) sheet.getCell(currentRowIndex, 70).values = [["TAK"]];
            if (incidentsText.trim() !== "") sheet.getCell(currentRowIndex, 36).values = [[incidentsText]];
            
            await context.sync();
            
            // Reset interfejsu
            document.getElementById("incidents-section").classList.add("hidden");
            document.getElementById("btn-fetch").classList.remove("hidden");
            
            // Wyczyść checkboxy i tekst
            document.getElementById("chk-material").checked = false;
            document.getElementById("chk-break").checked = false;
            document.getElementById("chk-breakdown").checked = false;
            document.getElementById("in-other-incidents").value = "";
            document.getElementById("in-operator").value = "";
            
            setStatus("Zapisano pomyślnie. Zaznacz nowy wiersz.");
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd zapisu: " + error.message);
    }
}

function setStatus(message) {
    const statusEl = document.getElementById("status-message");
    if (statusEl) {
        statusEl.innerText = message;
    }
}
