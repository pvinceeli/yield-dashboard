/* ================================
   EXISTING SIDEBAR / TAB LOGIC
================================ */
let workbook = null;
let lineChart = null;
let worksheet = null;
let top10Defects = {};

window.addEventListener("DOMContentLoaded", () => {
    initEmptyLineChart();
});


function toggleSidebar() {
    document.getElementById("sidebar").classList.toggle("collapsed");
}

function switchTab(index) {
    const tabs = document.querySelectorAll(".tab");
    const panels = document.querySelectorAll(".tab-panel");

    tabs.forEach((tab, i) => {
        tab.classList.toggle("active", i === index);
        panels[i].classList.toggle("active", i === index);
    });
}

document.querySelector(".open-file-btn").addEventListener("click", async () => {
    try {
        const [fileHandle] = await window.showOpenFilePicker({
            types: [
                {
                    description: "Excel Files",
                    accept: {
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
                        "application/vnd.ms-excel": [".xls"]
                    }
                }
            ],
            multiple: false
        });

        const file = await fileHandle.getFile();
        const arrayBuffer = await file.arrayBuffer();

        // Parse workbook
        workbook = XLSX.read(arrayBuffer, { type: "array" });

        // Display filename
        document.getElementById("selected-file").textContent = file.name;

        // Populate sheet dropdown
        populateSheetDropdown(workbook);

        // Optionally, set default worksheet
        worksheet = workbook.Sheets[workbook.SheetNames[0]];

        // ✅ SAFE ZONE — business logic
        const processes = getProcesses(workbook);
        setProcessList(processes);

    } catch (err) {
        if (err.name !== "AbortError") {
            console.error("Error opening Excel file:", err);
            alert("Failed to open Excel file.");
        }
    }
});

function populateSheetDropdown(workbook) {
    const sheetSelect = document.getElementById("sheet-select");

    // Clear existing options
    sheetSelect.innerHTML = "";

    // Add placeholder
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Change Sheet";
    placeholder.disabled = true;
    placeholder.selected = true;
    sheetSelect.appendChild(placeholder);

    // Add sheets
    workbook.SheetNames.forEach(sheetName => {
        const option = document.createElement("option");
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelect.appendChild(option);
    });

    // Optional: change worksheet when user selects a sheet
    sheetSelect.addEventListener("change", (e) => {
        const selectedSheet = e.target.value;
        worksheet = workbook.Sheets[selectedSheet];
        console.log("Selected sheet:", selectedSheet);
        // You can also re-run business logic here if needed
        const processes = getProcesses(workbook);
        setProcessList(processes);
	clearTopDefects();
	updateLineChart([], [], "")
	top10Defects = {};
	disableTopDefectRadios();
    });
}

function disableTopDefectRadios() {
    const radios = document.querySelectorAll('input[name="topDefect"]');

    radios.forEach(radio => {
        radio.disabled = true;  // disable the radio button
        radio.checked = false;  // optional: uncheck it
    });
}

function getProcesses(workbook) {
    const sheet = worksheet;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const processYieldMap = {};

    for (let i = 0; i < rows.length; i++) {
        const colB = rows[i][0]; // "Process"
        const colC = rows[i][1]; // Process Name
	if (!colB) continue;
        if (colB.replace(/\s+/g, '').toLowerCase() === "process" && colC) {
            const startIndex = i;
            let endIndex = rows.length - 1;
		
            // Find yield %
            for (let j = i + 1; j < rows.length; j++) {
		if(!rows[j][0]) continue;
                if (rows[j][0].replace(/\s+/g, '').toLowerCase() === "yield%" ) {
                    processYieldMap[colC] = {
                        yield: Number((rows[j][63] * 100).toFixed(2)),
                        indices: [startIndex, null]
                    };
                }

                // Stop when next process is found
                if (rows[j][0].replace(/\s+/g, '').toLowerCase() === "process") {
                    endIndex = j - 1;
                    i = j - 1;
                    break;
                }
            }

            // Finalize indices
	    try{
            	processYieldMap[colC].indices[1] = endIndex;
	    }catch(error){
	    	console.error("Yield% not found on " + colC, error.message);
		alert("Yield% not found on " + colC);
	    }
        }
    }
    return processYieldMap;
}



function setProcessList(processes) {
    const container = document.getElementById("dynamic-container");
    container.innerHTML = "";

    Object.entries(processes).forEach(([processName, processData]) => {
        const row = document.createElement("div");
        row.className = "process-row";

        // Process name
        const processInput = document.createElement("input");
        processInput.className = "process-input";
        processInput.type = "text";
        processInput.value = processName;
        processInput.disabled = true;

        // Yield
        const yieldInput = document.createElement("input");
        yieldInput.className = "yield-input";
        yieldInput.type = "text";
        yieldInput.value = `${processData.yield.toFixed(2)}%`;
        yieldInput.disabled = true;

        // Action button
        const btn = document.createElement("button");
        btn.className = "process-btn";
        btn.textContent = "Open";
        btn.addEventListener("click", () => {
    		const data = getProcessCumulative({
        		processName,
        		yieldValue: processData.yield,
        		indices: processData.indices
    		});
	});


        row.appendChild(processInput);
        row.appendChild(yieldInput);
        row.appendChild(btn);

        container.appendChild(row);
    });
}

function normalizeRange(a, b) {
    return [Math.min(a, b), Math.max(a, b)];
}

function getProcessCumulative({ processName, yieldValue, indices }) {
    if (!workbook) {
        console.error("Workbook not loaded!");
        return {};
    }
    const sheet = worksheet;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const [startRow, endRow] = [Math.min(indices[0], indices[1]), Math.max(indices[0], indices[1])];

    const targets = new Set([
        "Total Input",
        "Total Output",
        "Total Quantity Rejected",
        "Target Yield "
    ]);

    const result = {
        process: processName,
        yield: yieldValue
    };

    // Pre-normalize the targets to avoid repeated string processing
const normalizedTargets = new Set([...targets].map(t => t.replace(/\s+/g, '').toLowerCase()));

for (let i = startRow; i <= endRow + 1 && i < rows.length; i++) {
    const label = rows[i][0];
    if (!label) continue;
    const normalizedLabel = label.replace(/\s+/g, '').toLowerCase();
    if (normalizedTargets.has(normalizedLabel)) {
        // Find the original target string to store in result
        const originalTarget = [...targets].find(t => t.replace(/\s+/g, '').toLowerCase() === normalizedLabel);
        result[originalTarget] = rows[i][63];
    }
}

    setProcessCumulative(processName, result);
    const topDefects = getTopDefect(indices);
    top10Defects = topDefects.overall;
    setTopDefects(topDefects);
    bindTopDefectRadios(topDefects);
    setDefectLineChart(topDefects, top10Defects[0].name)
}

function setProcessCumulative(processname, totals) {
    // Get the container
    const container = document.querySelector(".rows");
    if (!container) return;
    //Update Process Span
    const processSpan = document.querySelector(".process-span");
    processSpan.textContent = processname;

    // Update Target Yield
    const targetYieldSpan = container.querySelector(".target-yield");
    if (targetYieldSpan && totals["Target Yield "] != null) {
        targetYieldSpan.textContent = totals["Target Yield "].toFixed(2)*100 + "%";
    }

    // Update Actual Yield
    const actualYieldSpan = container.querySelector(".actual-yield");
    if (actualYieldSpan && totals.yield != null) {
        actualYieldSpan.textContent = totals.yield.toFixed(2) + "%";
    }

    // Update Total Input
    const totalInputSpan = container.querySelector(".total-input");
    if (totalInputSpan && totals["Total Input"] != null) {
        totalInputSpan.textContent = totals["Total Input"];
    }

    // Update Total Output
    const totalOutputSpan = container.querySelector(".total-output");
    if (totalOutputSpan && totals["Total Output"] != null) {
        totalOutputSpan.textContent = totals["Total Output"];
    }

    // Update Rejected Quantity
    const rejectQtySpan = container.querySelector(".reject-qty");
    if (rejectQtySpan && totals["Total Quantity Rejected"] != null) {
        rejectQtySpan.textContent = totals["Total Quantity Rejected"];
    }
}

function clearTopDefects() {
    const fillEmpty = (table, maxRows) => {
        if (!table) return;

        const rows = table.querySelectorAll("tr");
        for (let i = 1; i <= maxRows; i++) { // skip header
            const row = rows[i];
            if (!row) continue;

            // Set all cells in the row to '-'
            for (let cell of row.cells) {
                cell.textContent = "-";
            }
        }
    };

    // Weekly tables (ww1 → ww5)
    for (let week = 1; week <= 5; week++) {
        const table = document.querySelector(
            `.tab2-container .table-container:nth-child(${week + 1}) table`
        );
        fillEmpty(table, 5); // top 5 rows per week
    }

    // Overall table (top 10)
    const overallTable = document.querySelector(".overall-container");
    fillEmpty(overallTable, 10); // top 10 rows
}

function getTopDefect(indices) {
    if (!workbook) {
        console.error("Workbook not loaded!");
        return {};
    }

    const sheet = worksheet;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const startIndex = Math.min(indices[0], indices[1]);
    const endIndex = Math.max(indices[0], indices[1]);

    // Week columns: BO → BS (Week 1 → Week 5)
    const weekColumns = [65, 66, 67, 68, 69];  

    // Find the "MACHINING DEFECT" row within the process block
    let startRow = null;
    for (let i = startIndex; i <= endIndex; i++) {
        if (rows[i][0]?.toString().trim() === "MACHINING DEFECT") {
            startRow = i + 1;
            break;
        }
    }

    if (startRow === null) {
        console.warn("MACHINING DEFECT not found in process block");
        return {};
    }

    const weekDefects = {};
    const overallDefects = {};

    weekColumns.forEach((colIndex, i) => {
        const weekKey = `ww${i + 1}`;
        const defects = {};

        for (let r = startRow; r <= endIndex; r++) {
            const label = rows[r][0];       // Column A
            const value = rows[r][colIndex]; // Week-specific column
            if (label && typeof value === "number" && value > 0) {
                defects[label] = value;
                overallDefects[label] = (overallDefects[label] || 0) + value;
            }
        }
        // Sort and pick top 5 for this week
        const topDefects = Object.entries(defects)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 5)
            .map(([name, value]) => ({ name, value }));

        weekDefects[weekKey] = topDefects;
    });

    // Overall top 10 defects across all weeks
    const topOverall = Object.entries(overallDefects)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(([name, value]) => ({ name, value }));

    weekDefects.overall = topOverall;
    return weekDefects;
}


function setTopDefects(topDefects) {
    if (!topDefects) return;

    // Read total input once
    const totalInputSpan = document.querySelector(".total-input");
    const totalInputs = totalInputSpan
        ? Number(totalInputSpan.textContent.replace(/,/g, "")) || 0
        : 0;

    const fillTable = (table, data, maxRows) => {
        if (!table) return;
	
        const rows = table.querySelectorAll("tr");

        for (let i = 1; i <= maxRows; i++) {
            const row = rows[i];
            if (!row) continue;

            const defect = data[i - 1];

            if (defect) {
                row.cells[0].textContent = defect.name.toUpperCase();
                row.cells[1].textContent = defect.value;

                // NG%
                if (row.cells[2]) {
                    const ngPct = totalInputs > 0
                        ? ((defect.value / totalInputs) * 100).toFixed(2)
                        : "0.00";
                    row.cells[2].textContent = ngPct;
                }
            } else {
                row.cells[0].textContent = "-";
                row.cells[1].textContent = "-";
                if (row.cells[2]) row.cells[2].textContent = "-";
            }
        }
    };
    /* ---------- Weekly tables (ww1 → ww5) ---------- */
    for (let week = 1; week <= 5; week++) {
        const weekKey = `ww${week}`;
        const weekData = topDefects[weekKey] || [];

        const table = document.querySelector(
            `.tab2-container .table-container:nth-child(${week+1}) table`
        );
        fillTable(table, weekData, 5);
    }

    /* ---------- Overall table (top 10) ---------- */
    if (topDefects.overall) {
        const overallTable = document.querySelector(
            ".overall-container"
        );

        fillTable(overallTable, topDefects.overall, 10);
    }
}

function initEmptyLineChart() {
    const ctx = document.getElementById("lineChart").getContext("2d");

    lineChart = new Chart(ctx, {
        type: "line",
        data: {
            labels: [],           // 👈 empty X-axis
            datasets: [{
                label: "Defect Trend",
                data: [],          // 👈 empty Y values
                borderColor: "#1f77b4",
                backgroundColor: "rgba(31, 119, 180, 0.15)",
                borderWidth: 2,
                tension: 0.3,
                fill: true,
                pointRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: false,     // smoother when updating often
            plugins: {
                legend: {
                    display: true
                },
                tooltip: {
                    enabled: true
                }
            },
            scales: {
                x: {
                    title: {
                        display: true,
                        text: "Week"
                    }
                },
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: "Quantity"
                    }
                }
            }
        }
    });
}




function setDefectLineChart(weekDefects, defectName) {
    if (!weekDefects || !defectName) return;

    // Only weekly keys (exclude "overall")
    const weekKeys = Object.keys(weekDefects)
        .filter(k => k.startsWith("ww"))
        .sort((a, b) => Number(a.slice(2)) - Number(b.slice(2)));

    const labels = [];
    const values = [];
 
    weekKeys.forEach(weekKey => {
        labels.push(weekKey.toUpperCase());
        const weekData = weekDefects[weekKey] || [];
        const match = weekData.find(d => d.name === defectName);
        values.push(match ? match.value : 0);
    });

    updateLineChart(labels, values, defectName);
}



function bindTopDefectRadios(weekDefects) {
    if (!weekDefects || !Array.isArray(weekDefects.overall)) return;

    const radios = document.querySelectorAll('input[name="topDefect"]');

    // Remove old listeners by cloning
    radios.forEach(radio => {
        radio.replaceWith(radio.cloneNode(true));
    });

    const newRadios = document.querySelectorAll('input[name="topDefect"]');
    const topOverall = weekDefects.overall;

    newRadios.forEach((radio, index) => {
        const defect = topOverall[index];
        if (!defect) return;

        radio.addEventListener("change", () => {
            setDefectLineChart(weekDefects, defect.name);
        });
    });

    // Update radio labels/UI
    updateTopDefectRadios(topOverall);

    // Initial load → Top 1 defect
    if (topOverall.length > 0) {
        setDefectLineChart(weekDefects, topOverall[0].name);
    }
}



function updateLineChart(labels, values, defectName) {
    if (!lineChart) return;

    lineChart.data.labels = labels;
    lineChart.data.datasets[0].label = defectName.toUpperCase();
    lineChart.data.datasets[0].data = values;
    lineChart.update();
}

function updateTopDefectRadios(top10Defects) {
    const radios = document.querySelectorAll('input[name="topDefect"]');
    if (!Array.isArray(top10Defects)) return;

    radios.forEach((radio, index) => {
        const defect = top10Defects[index];

        if (!defect) {
            radio.disabled = true;
            radio.checked = false;
        } else {
            radio.disabled = false;
            radio.dataset.defectName = defect.name;
        }
    });

    // Auto-select the first enabled radio
    const firstEnabled = Array.from(radios).find(r => !r.disabled);
    if (firstEnabled) {
	firstEnabled.checked = true;
	setDefectLineChart(top10Defects, top10Defects[0].name);
	}
}









