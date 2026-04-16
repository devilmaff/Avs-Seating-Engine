// Global State
let allStudents1 = [];

function checkLogin() {
    const user = document.getElementById('loginUsername').value;
    const pass = document.getElementById('loginPassword').value;
    if (user === 'husain2006' && pass === 'husain2006') {
        document.getElementById('login-overlay').style.display = 'none';
        document.getElementById('main-app').style.display = 'flex';
    } else {
        document.getElementById('loginError').style.display = 'block';
    }
}
let allStudents2 = [];
let currentFile1Name = "";
let currentFile2Name = "";
let generatedTablesData = {};
let activeContinueTableId = null;

function customModalConfig(title, text, showInput, defaultValue, confirmText, showCancel) {
    return new Promise((resolve) => {
        let overlay = document.getElementById('custom-modal-overlay');
        document.getElementById('modal-title').textContent = title;

        let textLines = text.split('\n');
        let textHtml = textLines.map(line => `<div>${line}</div>`).join('');
        document.getElementById('modal-text').innerHTML = textHtml;

        let inputElement = document.getElementById('modal-input');
        if (showInput) {
            inputElement.style.display = 'block';
            inputElement.value = defaultValue || '';
            setTimeout(() => inputElement.focus(), 100);
        } else {
            inputElement.style.display = 'none';
        }

        let confirmBtn = document.getElementById('modal-confirm-btn');
        confirmBtn.textContent = confirmText || 'Confirm';

        let cancelBtn = document.getElementById('modal-cancel-btn');
        if (showCancel) {
            cancelBtn.style.display = 'inline-block';
        } else {
            cancelBtn.style.display = 'none';
        }

        overlay.style.display = 'flex';

        function cleanup() {
            overlay.style.display = 'none';
            confirmBtn.removeEventListener('click', onConfirm);
            cancelBtn.removeEventListener('click', onCancel);
            inputElement.removeEventListener('keydown', onKeydown);
        }

        function onConfirm() {
            cleanup();
            resolve(showInput ? inputElement.value : true);
        }

        function onCancel() {
            cleanup();
            resolve(null);
        }

        function onKeydown(e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                onConfirm();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                onCancel();
            }
        }

        confirmBtn.addEventListener('click', onConfirm);
        cancelBtn.addEventListener('click', onCancel);
        inputElement.addEventListener('keydown', onKeydown);
    });
}

function customPrompt(text, defaultValue, title = "Input Required") {
    return customModalConfig(title, text, true, defaultValue, "Confirm", true);
}

function customAlert(text, title = "Notification") {
    return customModalConfig(title, text, false, "", "OK", false);
}

function toggleFile2() {
    let fileCount = document.getElementById('fileCount').value;
    if (fileCount === "1") {
        document.getElementById('file2-container').style.display = 'none';
        document.getElementById('bench-container').style.display = 'block';
    } else if (fileCount === "3") {
        document.getElementById('file2-container').style.display = 'block';
        document.getElementById('bench-container').style.display = 'block';
    } else {
        document.getElementById('file2-container').style.display = 'block';
        document.getElementById('bench-container').style.display = 'none';
    }
}

// Handle File Parsing
async function readRegisterNumbers(fileId) {
    const input = document.getElementById(fileId);
    if (!input.files || input.files.length === 0) return [];

    const file = input.files[0];
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Find the column index for "Reg"
    let headerRowIdx = -1;
    let regColIdx = -1;

    for (let i = 0; i < Math.min(10, rows.length); i++) {
        for (let j = 0; j < rows[i].length; j++) {
            let cell = String(rows[i][j] || "").toLowerCase().replace(/[^a-z0-9]/g, '');
            if (cell.includes("regno") || cell.includes("registerno") || cell.includes("register")) {
                headerRowIdx = i;
                regColIdx = j;
                break;
            }
        }
        if (regColIdx !== -1) break;
    }

    let result = [];
    if (regColIdx !== -1) {
        for (let i = headerRowIdx + 1; i < rows.length; i++) {
            if (rows[i][regColIdx]) result.push(String(rows[i][regColIdx]).trim());
        }
    } else {
        // Fallback: flat list or just grab column 0 if it looks like numbers
        for (let i = 0; i < rows.length; i++) {
            let val = rows[i][0];
            if (!val) val = rows[i][1];
            if (val && !isNaN(parseInt(val.toString()[0]))) {
                result.push(String(val).trim());
            }
        }
    }
    return result;
}

async function generateSeatingUI() {
    let fileCountValue = document.getElementById('fileCount').value;
    let isSingleFile = fileCountValue === "1";
    let isSequential2Files = fileCountValue === "3";
    let personsPerBench = parseInt(document.getElementById('personsPerBench').value) || 2;

    let file1Input = document.getElementById('file1');
    let file2Input = document.getElementById('file2');

    let f1Name = file1Input.files.length > 0 ? file1Input.files[0].name : "";
    let f2Name = file2Input.files.length > 0 ? file2Input.files[0].name : "";

    // Only fetch if inputs actually changed
    if (f1Name !== currentFile1Name) {
        allStudents1 = await readRegisterNumbers('file1');
        currentFile1Name = f1Name;
    }
    if (!isSingleFile && f2Name !== currentFile2Name) {
        allStudents2 = await readRegisterNumbers('file2');
        currentFile2Name = f2Name;
    }

    if (allStudents1.length === 0 && (!isSingleFile && allStudents2.length === 0)) {
        await customAlert("Please upload valid Excel files containing student register numbers!", "Error");
        return;
    }

    let rowsPerBox = 6;
    let totalColumns = 9; // Grid has 9 columns visually (A=1,2,3; B=4,5,6; C=7,8,9)

    let targetCols = [];
    if (personsPerBench === 2) {
        targetCols = [0, 2, 3, 5, 6, 8];
    } else if (personsPerBench === 3) {
        targetCols = [0, 1, 2, 3, 4, 5, 6, 7, 8];
    }

    // Explicitly target all columns if working with mix classes
    if (fileCountValue === "2") {
        targetCols = [0, 1, 2, 3, 4, 5, 6, 7, 8];
    }
    let sections = Math.floor(targetCols.length / personsPerBench);
    let maxSeatsTotal = targetCols.length * rowsPerBox;

    document.getElementById('print-area').innerHTML = ''; // Clear previous outputs

    let currentFileForSequential = 1;
    let keepGenerating = true;
    let currentHallNo = document.getElementById('hallNo').value;

    while (keepGenerating && (allStudents1.length > 0 || allStudents2.length > 0)) {
        let columnsData = Array.from({ length: totalColumns }, () => []);

        if (isSingleFile) {
            // --- Single File Logic (Packed) ---
            let s1 = allStudents1.splice(0, maxSeatsTotal);

            for (let c = 0; c < targetCols.length; c++) {
                let colIndex = targetCols[c];
                for (let r = 0; r < rowsPerBox; r++) {
                    if (s1.length > 0) {
                        columnsData[colIndex].push(s1.shift());
                    } else {
                        columnsData[colIndex].push("");
                    }
                }
            }

            for (let i = 0; i < totalColumns; i++) {
                while (columnsData[i].length < rowsPerBox) {
                    columnsData[i].push("");
                }
            }

        } else if (isSequential2Files) {
            // --- NEW LOGIC: 2 Files (Same Class, Separate Benches, Serial No Wise) ---
            for (let i = 0; i < totalColumns; i++) {
                for (let r = 0; r < rowsPerBox; r++) {
                    columnsData[i].push("");
                }
            }

            for (let s = 0; s < sections; s++) {
                let sectionCols = targetCols.slice(s * personsPerBench, (s + 1) * personsPerBench);
                let rowHasGirl = new Array(rowsPerBox).fill(false);
                let rowHasBoy = new Array(rowsPerBox).fill(false);

                for (let c = 0; c < sectionCols.length; c++) {
                    let colIndex = sectionCols[c];

                    for (let r = 0; r < rowsPerBox; r++) {
                        let placed = false;

                        if (currentFileForSequential === 1) {
                            if (allStudents1.length > 0) {
                                if (!rowHasBoy[r]) {
                                    columnsData[colIndex][r] = allStudents1.shift();
                                    rowHasGirl[r] = true;
                                    placed = true;
                                }
                            } else {
                                currentFileForSequential = 2; // Switch to file 2 instantly in the same cell
                            }
                        }

                        if (!placed && currentFileForSequential === 2 && allStudents2.length > 0) {
                            if (!rowHasGirl[r]) {
                                columnsData[colIndex][r] = allStudents2.shift();
                                rowHasBoy[r] = true;
                                placed = true;
                            }
                        }
                    }
                }
            }

        } else {
            // --- Original 2-File Alternating Logic ---
            let s1 = allStudents1.splice(0, 30);
            let s2 = allStudents2.splice(0, 24);

            for (let c = 0; c < totalColumns; c++) {
                let isFile1 = (c % 2 === 0);
                let currentSource = isFile1 ? s1 : s2;
                for (let r = 0; r < rowsPerBox; r++) {
                    if (currentSource.length > 0) {
                        columnsData[c].push(currentSource.shift());
                    } else {
                        columnsData[c].push("");
                    }
                }
            }
        }

        let tableHtml = renderTable(columnsData, rowsPerBox, currentHallNo, targetCols, personsPerBench, fileCountValue);
        document.getElementById('print-area').insertAdjacentHTML('beforeend', tableHtml);

        let remaining = allStudents1.length + (isSingleFile ? 0 : allStudents2.length);
        if (remaining > 0) {
            let nextHall = await customPrompt(`Table generated for Hall: ${currentHallNo}.\n\nThere are ${remaining} students remaining. Please enter the NEXT Hall No to generate the seating below:`, "", "Next Hall Number");
            if (nextHall === null || nextHall.trim() === "") {
                keepGenerating = false;
                await customAlert(`Stopped generating. There are still ${remaining} students left in memory. Reload the files if you want to start over.`, "Generation Stopped");
            } else {
                currentHallNo = nextHall.trim();
            }
        } else {
            keepGenerating = false; // No students remaining, stop the loop
        }
    }
}

function loadDummyData() {
    allStudents1 = [];
    allStudents2 = [];
    for (let i = 1; i <= 30; i++) allStudents1.push("623524106" + String(i).padStart(3, '0'));
    for (let i = 1; i <= 24; i++) allStudents2.push("623523106" + String(i).padStart(3, '0'));

    document.getElementById('testName').value = "INTERNAL ASSESSMENT TEST-2";
    document.getElementById('hallNo').value = "II ECE";
    document.getElementById('session').value = "AN";
    document.getElementById('seatingDay').value = "WEDNESDAY";
    document.getElementById('rowsPerCol').value = "6";

    generateSeatingUI();
}

function renderTable(columnsData, rowsPerBox, hallNoOverride, targetCols, personsPerBench, fileCountValue) {
    let testName = document.getElementById('testName').value;
    let hallNo = hallNoOverride || document.getElementById('hallNo').value;
    let session = document.getElementById('session').value;
    let seatingDay = document.getElementById('seatingDay').value;
    let leftSig = document.getElementById('leftSignature') ? document.getElementById('leftSignature').value : 'EC';
    let centerSig = document.getElementById('centerSignature') ? document.getElementById('centerSignature').value : 'HOD';
    let rightSig = document.getElementById('rightSignature') ? document.getElementById('rightSignature').value : 'PRINCIPAL';

    let validCols = targetCols;
    if (!validCols || validCols.length === 0) {
        if (personsPerBench === 2 || String(personsPerBench) === "2") {
            validCols = [0, 2, 3, 5, 6, 8];
        } else if (fileCountValue === "2" || personsPerBench === 3 || String(personsPerBench) === "3") {
            validCols = [0, 1, 2, 3, 4, 5, 6, 7, 8];
        } else {
            validCols = [0, 1, 2, 3, 4, 5, 6, 7, 8];
        }
    }

    let emptyCount = 0;
    for (let c of validCols) {
        for (let r = 0; r < rowsPerBox; r++) {
            if (!columnsData[c] || !columnsData[c][r] || columnsData[c][r] === "") {
                emptyCount++;
            }
        }
    }

    let tableId = "export-table-" + Math.random().toString(36).substr(2, 9);

    generatedTablesData[tableId] = {
        columnsData: JSON.parse(JSON.stringify(columnsData)),
        rowsPerBox: rowsPerBox,
        hallNoOverride: hallNo,
        targetCols: validCols,
        personsPerBench: personsPerBench !== undefined ? personsPerBench : document.getElementById('personsPerBench').value,
        fileCountValue: fileCountValue || document.getElementById('fileCount').value
    };

    let html = `
    <div id="container-${tableId}">
    <div class="watermark-container" style="page-break-after: always; margin-bottom: 40px;">
    <table class="excel-table export-table-class" id="${tableId}">
        <colgroup>
            ${Array(9).fill().map(() => `<col class="col-sno"><col class="col-reg">`).join('')}
        </colgroup>
        <tbody>
            <tr>
                <td colspan="18" class="fontWeightBold textCenter no-border-vert" style="border-top: 2px solid #000; font-size: 14px;">AVS COLLEGE OF TECHNOLOGY</td>
            </tr>
            <tr>
                <td colspan="18" class="textCenter no-border-vert border-bottom-only">${testName.toUpperCase()}</td>
            </tr>
            <tr>
                <td colspan="18" class="textCenter no-border-vert border-bottom-only">Seating Arrangement</td>
            </tr>
            <tr>
                <td colspan="6" class="textLeft no-border-vert border-bottom-only">Hall No &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: ${hallNo}</td>
                <td colspan="6" class="textCenter no-border-vert border-bottom-only"></td>
                <td colspan="6" class="textRight no-border-vert border-bottom-only">Session &nbsp;${session}</td>
            </tr>
            <tr>
                <td colspan="18" class="textCenter fontWeightBold">Reg.No of Candidates</td>
            </tr>
            <tr>
                <td colspan="6" class="textCenter fontWeightBold" style="font-size: 18px;">A</td>
                <td colspan="6" class="textCenter fontWeightBold" style="font-size: 18px;">B</td>
                <td colspan="6" class="textCenter fontWeightBold" style="font-size: 18px;">C</td>
            </tr>
            <tr>
    `;

    for (let c = 0; c < 9; c++) {
        html += `<td class="textCenter" style="font-size: 10px;">S.<br>No</td><td class="textCenter">Reg.No</td>`;
    }
    html += `</tr>`;

    // Fill logic handling the actual items generated with 54 count checks
    for (let r = 0; r < rowsPerBox; r++) {
        html += `<tr>`;
        for (let c = 0; c < 9; c++) {
            let reg = columnsData[c] && columnsData[c][r] ? columnsData[c][r] : "";
            // If there's a problem with numbering over multiple tables, we remove static S.No (or recalculate it based on real counts)
            let sno = (c * rowsPerBox) + r + 1;
            html += `<td class="fontWeightBold textCenter">${reg ? sno : ""}</td><td class="fontWeightBold textCenter">${reg}</td>`;
        }
        html += `</tr>`;
    }

    // Empty row
    html += `<tr>${Array(18).fill().map(() => `<td></td>`).join('')}</tr>`;

    // Seating Date Box
    html += `<tr>`;
    html += `<td colspan="6" class="no-border"></td>`;
    html += `<td colspan="6" class="textCenter fontWeightBold" style="border: 2px solid #000; height: 40px; vertical-align: middle;">SEATING : ${seatingDay.toUpperCase()}</td>`;
    html += `<td colspan="6" class="no-border"></td>`;
    html += `</tr>`;

    // Two empty rows
    html += `<tr>${Array(18).fill().map(() => `<td class="no-border"></td>`).join('')}</tr>`;
    html += `<tr>${Array(18).fill().map(() => `<td class="no-border"></td>`).join('')}</tr>`;

    // Signatures
    html += `<tr>`;
    html += `<td colspan="6" class="textCenter fontWeightBold no-border" style="padding-top:20px;">${leftSig}</td>`;
    html += `<td colspan="6" class="textCenter fontWeightBold no-border" style="padding-top:20px;">${centerSig}</td>`;
    html += `<td colspan="6" class="textCenter fontWeightBold no-border" style="padding-top:20px;">${rightSig}</td>`;
    html += `</tr>`;

    html += `
        </tbody>
    </table>
    </div>
    <div style="text-align: center; margin-bottom: 20px;">
        <button type="button" class="btn btn-secondary" style="display: inline-flex;" onclick="exportToExcel('${tableId}', '${hallNo}')">⬇️ Download Excel for ${hallNo}</button>
    `;

    if (emptyCount > 0) {
        html += `<button type="button" class="btn btn-primary" style="display: inline-flex; margin-left: 10px;" onclick="triggerContinueNextFile('${tableId}')">➕ Continue Next File</button>`;
    }

    html += `
    </div>
    </div>
    `;

    return html;
}


function exportToExcel(tableId, hallNo) {
    let table = document.getElementById(tableId);
    if (!table) return;

    let wb = XLSX.utils.table_to_book(table, { sheet: hallNo, raw: true });
    let ws = wb.Sheets[hallNo];

    if (!ws['!ref']) return;
    let range = XLSX.utils.decode_range(ws['!ref']);

    // 1. Set Column Widths (S.No narrow, Reg.No wide)
    ws['!cols'] = [];
    for (let i = 0; i < 9; i++) {
        ws['!cols'].push({ wch: 6 });
        ws['!cols'].push({ wch: 16 });
    }

    // 2. Set Row Heights
    ws['!rows'] = [];
    for (let R = 0; R <= range.e.r; ++R) {
        ws['!rows'].push({ hpt: 22 });
    }
    ws['!rows'][0] = { hpt: 28 };
    ws['!rows'][range.e.r - 3] = { hpt: 35 };
    ws['!rows'][range.e.r] = { hpt: 30 };

    // 3. Apply styles for borders, alignment, fonts
    let borderThin = { style: "thin", color: { rgb: "000000" } };
    let borderMedium = { style: "medium", color: { rgb: "000000" } };
    let dataEndRow = range.e.r - 5;

    for (let R = 0; R <= range.e.r; ++R) {
        for (let C = 0; C <= range.e.c; ++C) {
            let cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
            if (!ws[cell_ref]) {
                ws[cell_ref] = { t: "s", v: "" };
            }
            let cell = ws[cell_ref];

            let cellStyle = {
                font: { name: "Arial", sz: 10 },
                alignment: { horizontal: "center", vertical: "center" }
            };

            if (R <= dataEndRow) {
                cellStyle.border = {
                    top: borderThin, bottom: borderThin, left: borderThin, right: borderThin
                };

                if (R === 0) cellStyle.border.top = borderMedium;
                if (R === dataEndRow) cellStyle.border.bottom = borderMedium;
                if (C === 0) cellStyle.border.left = borderMedium;
                if (C === range.e.c) cellStyle.border.right = borderMedium;

                if (R === 0) { cellStyle.font.bold = true; cellStyle.font.sz = 14; }
                if (R === 4 || R === 5 || R === 6) { cellStyle.font.bold = true; }

                if (R === 3) {
                    if (C < 6) { cellStyle.alignment.horizontal = "left"; }
                    if (C >= 12) { cellStyle.alignment.horizontal = "right"; }
                }

                if (R >= 7) {
                    cellStyle.font.bold = true;
                    if (C % 2 === 0) { cellStyle.font.sz = 9; }
                }
            }

            if (R === range.e.r - 3) {
                if (C >= 6 && C <= 11) {
                    cellStyle.border = {
                        top: borderMedium, bottom: borderMedium, left: borderMedium, right: borderMedium
                    };
                    cellStyle.font.bold = true;
                }
            }

            if (R === range.e.r) {
                cellStyle.font.bold = true;
            }

            cell.s = cellStyle;
        }
    }

    let safeHallNo = hallNo.replace(/[^a-z0-9]/gi, '_');
    XLSX.writeFile(wb, `Seating_Arrangement_${safeHallNo}.xlsx`);
}

function triggerContinueNextFile(tableId) {
    activeContinueTableId = tableId;
    document.getElementById('continueFile').click();
}

async function handleContinueFile(input) {
    if (!input.files || input.files.length === 0) return;
    if (!activeContinueTableId || !generatedTablesData[activeContinueTableId]) return;

    const file = input.files[0];
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    let headerRowIdx = -1;
    let regColIdx = -1;

    for (let i = 0; i < Math.min(10, rows.length); i++) {
        for (let j = 0; j < rows[i].length; j++) {
            let cell = String(rows[i][j] || "").toLowerCase().replace(/[^a-z0-9]/g, '');
            if (cell.includes("regno") || cell.includes("registerno") || cell.includes("register")) {
                headerRowIdx = i;
                regColIdx = j;
                break;
            }
        }
        if (regColIdx !== -1) break;
    }

    let newStudents = [];
    if (regColIdx !== -1) {
        for (let i = headerRowIdx + 1; i < rows.length; i++) {
            if (rows[i][regColIdx]) newStudents.push(String(rows[i][regColIdx]).trim());
        }
    } else {
        for (let i = 0; i < rows.length; i++) {
            let val = rows[i][0];
            if (!val) val = rows[i][1];
            if (val && !isNaN(parseInt(val.toString()[0]))) {
                newStudents.push(String(val).trim());
            }
        }
    }

    input.value = ""; // Reset input

    if (newStudents.length === 0) {
        await customAlert("No valid register numbers found in the uploaded file.", "Error");
        return;
    }

    let tableData = generatedTablesData[activeContinueTableId];
    let { columnsData, rowsPerBox, hallNoOverride, targetCols, personsPerBench, fileCountValue } = tableData;

    let newPersonsPerBench = await customPrompt(`You have ${newStudents.length} new students. Should they be seated 2 per bench or 3 per bench?\n\nEnter '2' or '3':`, personsPerBench ? String(personsPerBench) : document.getElementById('personsPerBench').value, "Seating Arrangement");

    if (newPersonsPerBench !== "2" && newPersonsPerBench !== "3") {
        await customAlert("Invalid input or cancelled. Continued generation aborted.", "Action Cancelled");
        return;
    }

    newPersonsPerBench = parseInt(newPersonsPerBench);
    let newTargetCols = [];
    if (newPersonsPerBench === 2) {
        newTargetCols = [0, 2, 3, 5, 6, 8];
    } else if (newPersonsPerBench === 3) {
        newTargetCols = [0, 1, 2, 3, 4, 5, 6, 7, 8];
    }

    tableData.targetCols = newTargetCols; // Update active table's targetCols state
    tableData.personsPerBench = newPersonsPerBench;

    // Fill remaining spots in the current table AFTER the last overall occupied cell
    let cellSequence = [];
    let lastOccupiedIndex = -1;

    for (let c = 0; c < newTargetCols.length; c++) {
        let colIndex = newTargetCols[c];
        for (let r = 0; r < rowsPerBox; r++) {
            cellSequence.push({ c: colIndex, r: r });
            if (columnsData[colIndex] && columnsData[colIndex][r] !== "") {
                lastOccupiedIndex = cellSequence.length - 1;
            }
        }
    }

    for (let i = lastOccupiedIndex + 1; i < cellSequence.length; i++) {
        if (newStudents.length > 0) {
            let pos = cellSequence[i];
            columnsData[pos.c][pos.r] = newStudents.shift();
        } else {
            break;
        }
    }

    // Re-render updated table
    let container = document.getElementById("container-" + activeContinueTableId);
    if (container) {
        container.outerHTML = renderTable(columnsData, rowsPerBox, hallNoOverride, newTargetCols, newPersonsPerBench, fileCountValue);
    }

    // Generate subsequent tables if newStudents still has remaining items
    let keepGenerating = true;
    let currentHallNo = hallNoOverride;
    let totalColumns = 9;

    while (keepGenerating && newStudents.length > 0) {
        let nextHall = await customPrompt(`Current table filled.\n\nThere are ${newStudents.length} students remaining from the new file. Please enter the NEXT Hall No to generate the seating below:`, "", "Next Hall Number");
        if (nextHall === null || nextHall.trim() === "") {
            keepGenerating = false;
            await customAlert(`Stopped generating. There are still ${newStudents.length} students left unprocessed.`, "Generation Stopped");
        } else {
            currentHallNo = nextHall.trim();

            let newColumns = Array.from({ length: totalColumns }, () => []);
            for (let i = 0; i < totalColumns; i++) {
                for (let r = 0; r < rowsPerBox; r++) newColumns[i].push("");
            }

            for (let c = 0; c < newTargetCols.length; c++) {
                let colIndex = newTargetCols[c];
                for (let r = 0; r < rowsPerBox; r++) {
                    if (newStudents.length > 0) {
                        newColumns[colIndex][r] = newStudents.shift();
                    }
                }
            }

            let newTableHtml = renderTable(newColumns, rowsPerBox, currentHallNo, newTargetCols, newPersonsPerBench, fileCountValue);
            document.getElementById('print-area').insertAdjacentHTML('beforeend', newTableHtml);
        }
    }
}
