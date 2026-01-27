/* ================================
   EXISTING SIDEBAR / TAB LOGIC
================================ */
let workbook = null;
let lineChart = null;
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

/* ================================
   FOLDER-BASED DROPDOWNS
================================ */

let baseDirHandle = null;

const selects = [
    document.getElementById("pg"),       // PRODUCT GROUP
    document.getElementById("product"),  // PRODUCT
    document.getElementById("model"),    // MODEL
    document.getElementById("year"),     // YEAR
    document.getElementById("month")     // MONTH
];

// Button: select base directory ONCE
document.getElementById("pickFolder").addEventListener("click", async () => {
    try {
        baseDirHandle = await window.showDirectoryPicker();
        await loadLevel(baseDirHandle, selects[0]);
        console.log("Base folder selected");
    } catch (err) {
        console.warn("Folder selection cancelled");
    }
});

// Load folder names into a dropdown
async function loadLevel(dirHandle, select) {
    select.innerHTML = `<option value="">Select</option>`;
    for await (const entry of dirHandle.values()) {
        if (entry.kind === "directory") {
            const opt = document.createElement("option");
            opt.value = entry.name;
            opt.textContent = entry.name;
            select.appendChild(opt);
        }
    }
}

// Cascading dropdown behavior
selects.forEach((select, index) => {
    select.addEventListener("change", async () => {
        // Clear downstream selects
        for (let i = index + 1; i < selects.length; i++) {
            selects[i].innerHTML = `<option value="">Select</option>`;
        }

        if (!baseDirHandle || !select.value) return;

        try {
            let currentHandle = baseDirHandle;
            for (let i = 0; i <= index; i++) {
                currentHandle = await currentHandle.getDirectoryHandle(selects[i].value);
            }

            if (selects[index + 1]) {
                await loadLevel(currentHandle, selects[index + 1]);
            }
        } catch (err) {
            console.error("Folder not found:", err);
        }
    });
});

/* ================================
   OPEN EXCEL & RENDER CHART
================================ */

document.querySelector(".open-file-btn").addEventListener("click", async () => {
    if (!baseDirHandle) {
        alert("Please select a base folder first");
        return;
    }

    const values = selects.map(s => s.value);
    if (values.includes("")) {
        alert("Please select all dropdowns");
        return;
    }

    // Declare dirHandle here
    let dirHandle = baseDirHandle;

    try {
        // Traverse dropdown-selected subdirectories
        for (let val of values) {
            dirHandle = await dirHandle.getDirectoryHandle(val);
        }

        // Open Excel file
        const fileHandle = await dirHandle.getFileHandle("data.xlsx");
        const file = await fileHandle.getFile();
        const arrayBuffer = await file.arrayBuffer();

        // Parse workbook
        workbook = XLSX.read(arrayBuffer, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        worksheet = workbook.Sheets[sheetName];

    } catch (err) {
        console.error("Error opening file:", err);
        alert("File or folder not found. Make sure 'data.xlsx' exists in the selected path.");
        return; // stop execution if parsing failed
    }

    // ✅ SAFE ZONE — business logic only
    const processes = getProcesses(workbook);
    console.log("The processes are:", processes);
    setProcessList(processes);
});

// Chart.js rendering function
function renderChart(data) {
    // Example assumes Excel has columns: "Label" and "Value"
    const labels = data.map(row => row.Label);
    const values = data.map(row => row.Value);

    const canvas = document.getElementById("myChart");
    if (!canvas) {
        console.warn("No canvas element found with id 'myChart'");
        return;
    }

    const ctx = canvas.getContext("2d");

    // Destroy previous chart if exists
    if (window.myChart) window.myChart.destroy();

    window.myChart = new Chart(ctx, {
        type: "bar",
        data: {
            labels,
            datasets: [{
                label: "Excel Data",
                data: values,
                backgroundColor: "rgba(75, 192, 192, 0.6)"
            }]
        },
        options: {
            responsive: true,
            plugins: { legend: { display: true } }
        }
    });
}

function getProcesses(workbook) {
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const processYieldMap = {};

    for (let i = 0; i < rows.length; i++) {
        const colB = rows[i][0]; // "Process"
        const colC = rows[i][1]; // Process Name

        if (colB === "Process" && colC) {
            const startIndex = i;
            let endIndex = rows.length - 1;

            // Find yield %
            for (let j = i + 1; j < rows.length; j++) {
                if (rows[j][0] === "Yield %") {
                    processYieldMap[colC] = {
                        yield: Number((rows[j][63] * 100).toFixed(2)),
                        indices: [startIndex, null]
                    };
                }

                // Stop when next process is found
                if (rows[j][0] === "Process") {
                    endIndex = j - 1;
                    i = j - 1;
                    break;
                }
            }

            // Finalize indices
            processYieldMap[colC].indices[1] = endIndex;
        }
    }
    console.log(processYieldMap);
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

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
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

    for (let i = startRow; i <= endRow+1 && i < rows.length; i++) {
        const label = rows[i][0]; // Column B
        if (targets.has(label)) {
            result[label] = rows[i][63]; // Column BM
        }
    }
    console.log(result);
    setProcessCumulative(processName, result);
    const topDefects = getTopDefect(indices);
    setTopDefects(topDefects);
    setTopDefectsSummary(topDefects);
    updateChartByDefect(topDefects, 0);
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

function getTopDefect(indices) {
    if (!workbook) {
        console.error("Workbook not loaded!");
        return {};
    }

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const endRow = Math.max(indices[0], indices[1]);
    console.log("endRow is: ", endRow);

    // Week columns: BO → BS (Week 1 → Week 5)
    const weekColumns = [61, 62, 63, 64, 65];  

    // Find the "MACHINING DEFECT" row within the process block
    let startRow = null;
    for (let i = 0; i < rows.length; i++) {
    if (i < indices[0] || i > indices[1]) continue; // skip outside process block
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

    weekColumns.forEach((colIndex, i) => {
        const weekKey = `ww${i + 1}`;
        const defects = {};
	console.log("startRow: ", startRow);
	console.log("endRow: ", endRow);
	console.log("for loop: ", startRow <= endRow );
        for (let r = startRow; r <= endRow; r++) {
            const label = rows[r][0];       // Column B
            const value = rows[r][colIndex]; // Week-specific column
	    console.log(value, typeof value === "number");
            if (label && typeof value === "number" && value > 0) {
                defects[label] = value;
            }
        }

        // Sort and pick top 5
        const topDefects = Object.entries(defects)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 5)
            .map(([name, value]) => ({ name, value }));

        weekDefects[weekKey] = topDefects;
    });

    return weekDefects;
}
function setTopDefects(topDefects) {
    // topDefects = { ww1: [{name, value}, ...], ww2: [...], ... ww5: [...] }

    for (let week = 1; week <= 5; week++) {
        const weekKey = `ww${week}`;
        const weekData = topDefects[weekKey] || [];

        // Find the corresponding table
        const tableContainer = document.querySelector(`.tab2-container .table-container:nth-child(${week}) table`);
        if (!tableContainer) continue;

        // Get all rows except header (first row)
        const rows = tableContainer.querySelectorAll("tr");
        
        for (let i = 1; i <= 5; i++) { // only fill top 5 rows
            const defect = weekData[i - 1];
            const row = rows[i];
            if (!row) continue;

            if (defect) {
                row.cells[0].textContent = defect.name; // Top Defects
                row.cells[1].textContent = defect.value; // QTY
                // row.cells[2] left as is for NG%
            } else {
                row.cells[0].textContent = "-";
                row.cells[1].textContent = "-";
            }
        }
    }
}
function setTopDefectsSummary(weekDefects) {
    const combinedDefects = {};

    // Step 1: Combine all weeks
    Object.values(weekDefects).forEach(week => {
        week.forEach(({ name, value }) => {
            if (!combinedDefects[name]) combinedDefects[name] = 0;
            combinedDefects[name] += value; // sum over weeks
        });
    });

    // Step 2: Sort by value descending and pick top 10
    const top10 = Object.entries(combinedDefects)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(([name, value]) => ({ name, value }));

    // Step 3: Populate the table
    const table = document.querySelector(".overall-container table"); // assuming it's the first table after the <h4>
    const rows = table.querySelectorAll("tr");

    // Clear previous values (keep header row)
    for (let i = 1; i < rows.length; i++) {
        const td = rows[i].querySelectorAll("td");
        td[0].textContent = "-";
        td[1].textContent = "-";
        td[2].textContent = "-";
    }

    // Fill top 10 defects
    top10.forEach((defect, index) => {
        if (index + 1 >= rows.length) return; // avoid overflow
        const td = rows[index + 1].querySelectorAll("td");
        td[0].textContent = defect.name;
        td[1].textContent = defect.value;
        td[2].textContent = "-"; // NG% placeholder
    });

    console.log("Top 10 monthly defects:", top10);
}
function updateChartByDefect(topDefects, defectIndex) {
    if (!lineChart) {
        console.error("Chart instance not found!");
        return;
    }

    const weeks = ["ww1", "ww2", "ww3", "ww4", "ww5"];

    // Get quantity values for the selected defect across all weeks
    const data = weeks.map(week => {
        const defects = topDefects[week] || [];
        return defects[defectIndex] ? defects[defectIndex].value : 0;
    });

    // Update chart
    lineChart.data.labels = ["Week 1", "Week 2", "Week 3", "Week 4", "Week 5"];
    lineChart.data.datasets[0].data = data;
    lineChart.data.datasets[0].label = `Top ${defectIndex + 1} Defect`;

    lineChart.update();
}





