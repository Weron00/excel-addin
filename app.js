// Zapewnienie inicjalizacji Office.js przed użyciem
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("btn-fetch").onclick = fetchRowData;
        document.getElementById("btn-start").onclick = writeStartTime;
        document.getElementById("btn-stop").onclick = writeStopTime;
        setStatus("Dodatek załadowany pomyślnie.");
    }
});

let currentRowIndex = -1;

async function fetchRowData() {
    try {
        setStatus("Pobieranie danych...");
        await Excel.run(async (context) => {
            const activeCell = context.workbook.getActiveCell();
            activeCell.load("rowIndex");
            await context.sync();
            
            currentRowIndex = activeCell.rowIndex;
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Kolumny: E (indeks 4), M (indeks 12), O (indeks 14)
            const rangeE = sheet.getCell(currentRowIndex, 4);
            const rangeM = sheet.getCell(currentRowIndex, 12);
            const rangeO = sheet.getCell(currentRowIndex, 14);
            
            rangeE.load("values");
            rangeM.load("values");
            rangeO.load("values");
            
            await context.sync();
            
            // Wyświetlenie danych w interfejsie
            document.getElementById("val-program").innerText = rangeE.values[0][0] || "-";
            document.getElementById("val-warstwy").innerText = rangeM.values[0][0] || "-";
            document.getElementById("val-kit").innerText = rangeO.values[0][0] || "-";
            
            // Odblokowanie przycisków Start / Stop
            document.getElementById("btn-start").disabled = false;
            document.getElementById("btn-stop").disabled = false;
            
            setStatus(`Pobrano dane z wiersza ${currentRowIndex + 1}.`);
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd pobierania: " + error.message);
    }
}

async function writeStartTime() {
    if (currentRowIndex === -1) {
        setStatus("Najpierw pobierz dane!");
        return;
    }
    try {
        setStatus("Zapisywanie Start...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            // Kolumna AA to indeks 26
            const rangeAA = sheet.getCell(currentRowIndex, 26);
            
            const now = new Date();
            const dateStr = now.toLocaleDateString() + " " + now.toLocaleTimeString();
            
            rangeAA.values = [[dateStr]];
            await context.sync();
            
            setStatus(`Start zapisano w wierszu ${currentRowIndex + 1} (AA).`);
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd zapisu Start: " + error.message);
    }
}

async function writeStopTime() {
    if (currentRowIndex === -1) {
        setStatus("Najpierw pobierz dane!");
        return;
    }
    try {
        setStatus("Zapisywanie Stop...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            // Kolumna AB to indeks 27
            const rangeAB = sheet.getCell(currentRowIndex, 27);
            
            const now = new Date();
            const dateStr = now.toLocaleDateString() + " " + now.toLocaleTimeString();
            
            rangeAB.values = [[dateStr]];
            await context.sync();
            
            setStatus(`Stop zapisano w wierszu ${currentRowIndex + 1} (AB).`);
        });
    } catch (error) {
        console.error(error);
        setStatus("Błąd zapisu Stop: " + error.message);
    }
}

function setStatus(message) {
    document.getElementById("status-message").innerText = message;
}
