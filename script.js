/* ==================================================
   SIDEBAR, TAB NAVIGATION, AND FILE STATE MANAGEMENT
   ================================================== */
let workbook = null;       // Holds the parsed Excel workbook
let lineChart = null;      // Reference to the line chart instance
let worksheet = null;      // Currently selected worksheet
let top10Defects = {};     // Cache for top 10 defect metrics
const loader = document.getElementById("loader");

// Initialize UI components once the DOM is fully loaded
window.addEventListener("DOMContentLoaded", () => {
    initEmptyLineChart();
});

function showLoader() {
    loader.classList.remove("hidden");
}

function hideLoader() {
    loader.classList.add("hidden");
}


// Toggles the sidebar between expanded and collapsed states
function toggleSidebar() {
    document.getElementById("sidebar").classList.toggle("collapsed");
}

// Switches active tab and corresponding content panel
function switchTab(index) {
    const tabs = document.querySelectorAll(".tab");
    const panels = document.querySelectorAll(".tab-panel");

    tabs.forEach((tab, i) => {
        tab.classList.toggle("active", i === index);
        panels[i].classList.toggle("active", i === index);
    });
}

// Handles Excel file selection and initial processing
document.querySelector(".open-file-btn").addEventListener("click", async () => {
    try {
        // Open native file picker restricted to Excel formats
	showLoader();
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

        // Read selected file into memory
        const file = await fileHandle.getFile();
        const arrayBuffer = await file.arrayBuffer();

        // Parse Excel file into a workbook object
        workbook = XLSX.read(arrayBuffer, { type: "array" });

        // Update UI with selected filename
        document.getElementById("selected-file").textContent = file.name;

        // Populate worksheet selection dropdown
        populateSheetDropdown(workbook);

        // Default to the first worksheet in the workbook
        worksheet = workbook.Sheets[workbook.SheetNames[0]];

        // Business logic entry point — extract and bind process data
        const processes = getProcesses(workbook);
        setProcessList(processes);

    } catch (err) {
        // Ignore user-cancelled file selection
        if (err.name !== "AbortError") {
            console.error("Error opening Excel file:", err);
            alert("Failed to open Excel file.");
        }
    }
    hideLoader();
});

function populateSheetDropdown(workbook) {
    const sheetSelect = document.getElementById("sheet-select");

    // Reset dropdown to remove any previously loaded sheet options
    sheetSelect.innerHTML = "";

    // Add a disabled placeholder option to prompt user selection
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Change Sheet";
    placeholder.disabled = true;
    placeholder.selected = true;
    sheetSelect.appendChild(placeholder);

    // Populate dropdown with all sheet names from the workbook
    workbook.SheetNames.forEach(sheetName => {
        const option = document.createElement("option");
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelect.appendChild(option);
    });

    // Handle worksheet change initiated by user selection
    sheetSelect.addEventListener("change", (e) => {
        const selectedSheet = e.target.value;

        // Update active worksheet reference
        worksheet = workbook.Sheets[selectedSheet];
        console.log("Selected sheet:", selectedSheet);

        // Re-run business logic to refresh process-related data
        const processes = getProcesses(workbook);
        setProcessList(processes);

        // Reset dependent UI and cached state
	clearTopDefects();
	updateLineChart([], [], "");
	top10Defects = {};
	disableTopDefectRadios();
    });
}

function disableTopDefectRadios() {
    const radios = document.querySelectorAll('input[name="topDefect"]');

    radios.forEach(radio => {
        radio.disabled = true;  // Prevent user interaction
        radio.checked = false;  // Ensure no stale selection remains
    });
}

function getProcesses(workbook) {
    const sheet = worksheet;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const processYieldMap = {}; // Stores yield and row indices per process

    for (let i = 0; i < rows.length; i++) {
        const colB = rows[i][0]; // Process identifier label
        const colC = rows[i][1]; // Process name
	if (!colB) continue;

        // Detect start of a process block
        if (colB.replace(/\s+/g, '').toLowerCase() === "process" && colC) {
            const startIndex = i;
            let endIndex = rows.length - 1;
		
            // Scan forward to locate yield percentage and process boundary
            for (let j = i + 1; j < rows.length; j++) {
		if(!rows[j][0]) continue;

                // Capture yield percentage row
                if (rows[j][0].replace(/\s+/g, '').toLowerCase() === "yield%" ) {
                    processYieldMap[colC] = {
                        yield: Number((rows[j][63] * 100).toFixed(2)),
                        indices: [startIndex, null]
                    };
                }

                // Detect start of next process to determine end boundary
                if (rows[j][0].replace(/\s+/g, '').toLowerCase() === "process") {
                    endIndex = j - 1;
                    i = j - 1; // Skip ahead to avoid reprocessing
                    break;
                }
            }

            // Finalize process row range
	    try{
            	processYieldMap[colC].indices[1] = endIndex;
	    }catch(error){
                // Guard against missing yield rows
	    	console.error("Yield% not found on " + colC, error.message);
		alert("Yield% not found on " + colC);
	    }
        }
    }

    return processYieldMap;
}




function setProcessList(processes) {
    const container = document.getElementById("dynamic-container");

    // Remove any previously rendered process entries
    container.innerHTML = "";

    Object.entries(processes).forEach(([processName, processData]) => {
        const row = document.createElement("div");
        row.className = "process-row";

        // Read-only field displaying the process name
        const processInput = document.createElement("input");
        processInput.className = "process-input";
        processInput.type = "text";
        processInput.value = processName;
        processInput.disabled = true;

        // Read-only field displaying the calculated yield percentage
        const yieldInput = document.createElement("input");
        yieldInput.className = "yield-input";
        yieldInput.type = "text";
        yieldInput.value = `${processData.yield.toFixed(2)}%`;
        yieldInput.disabled = true;

        // Action button to load detailed cumulative and defect data for the process
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

        // Assemble and render the process row
        row.appendChild(processInput);
        row.appendChild(yieldInput);
        row.appendChild(btn);

        container.appendChild(row);
    });
}

// Utility helper to normalize two numeric values into an ordered range
function normalizeRange(a, b) {
    return [Math.min(a, b), Math.max(a, b)];
}

function getProcessCumulative({ processName, yieldValue, indices }) {
    // Guard clause to ensure workbook data is available
    if (!workbook) {
        console.error("Workbook not loaded!");
        return {};
    }

    const sheet = worksheet;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Normalize process row boundaries to ensure correct iteration order
    const [startRow, endRow] = [Math.min(indices[0], indices[1]), Math.max(indices[0], indices[1])];

    // Define cumulative metrics expected within a process block
    const targets = new Set([
        "Total Input",
        "Total Output",
        "Total Quantity Rejected",
        "Target Yield "
    ]);

    // Base result object for downstream processing and UI updates
    const result = {
        process: processName,
        yield: yieldValue
    };

    // Pre-normalize target labels to minimize repeated string normalization
    const normalizedTargets = new Set(
        [...targets].map(t => t.replace(/\s+/g, '').toLowerCase())
    );

    // Scan rows within the process range to extract cumulative metrics
    for (let i = startRow; i <= endRow + 1 && i < rows.length; i++) {
        const label = rows[i][0];
        if (!label) continue;

        const normalizedLabel = label.replace(/\s+/g, '').toLowerCase();

        if (normalizedTargets.has(normalizedLabel)) {
            // Resolve the original target key to preserve expected result structure
            const originalTarget = [...targets].find(
                t => t.replace(/\s+/g, '').toLowerCase() === normalizedLabel
            );
            result[originalTarget] = rows[i][63];
        }
    }

    // Update cumulative process data in the UI
    setProcessCumulative(processName, result);

    // Retrieve and bind defect data for the selected process
    const topDefects = getTopDefect(indices);
    top10Defects = topDefects.overall;
    setTopDefects(topDefects);
    bindTopDefectRadios(topDefects);

    // Initialize defect trend chart using the highest-ranked defect
    setDefectLineChart(topDefects, top10Defects[0].name);
}

function setProcessCumulative(processname, totals) {
    // Locate the main container for cumulative process metrics
    const container = document.querySelector(".rows");
    if (!container) return;

    // Update the displayed process name
    const processSpan = document.querySelector(".process-span");
    processSpan.textContent = processname;

    // Update target yield value (expected yield percentage)
    const targetYieldSpan = container.querySelector(".target-yield");
    if (targetYieldSpan && totals["Target Yield "] != null) {
        targetYieldSpan.textContent = totals["Target Yield "].toFixed(2)*100 + "%";
    }

    // Update actual yield value calculated from process data
    const actualYieldSpan = container.querySelector(".actual-yield");
    if (actualYieldSpan && totals.yield != null) {
        actualYieldSpan.textContent = totals.yield.toFixed(2) + "%";
    }

    // Update total input quantity for the process
    const totalInputSpan = container.querySelector(".total-input");
    if (totalInputSpan && totals["Total Input"] != null) {
        totalInputSpan.textContent = totals["Total Input"];
    }

    // Update total output quantity for the process
    const totalOutputSpan = container.querySelector(".total-output");
    if (totalOutputSpan && totals["Total Output"] != null) {
        totalOutputSpan.textContent = totals["Total Output"];
    }

    // Update total rejected quantity for the process
    const rejectQtySpan = container.querySelector(".reject-qty");
    if (rejectQtySpan && totals["Total Quantity Rejected"] != null) {
        rejectQtySpan.textContent = totals["Total Quantity Rejected"];
    }
}

function clearTopDefects() {
    // Helper to reset table cells to a placeholder value
    const fillEmpty = (table, maxRows) => {
        if (!table) return;

        const rows = table.querySelectorAll("tr");
        for (let i = 1; i <= maxRows; i++) { // Skip header row
            const row = rows[i];
            if (!row) continue;

            // Reset all cells in the row to a default placeholder
            for (let cell of row.cells) {
                cell.textContent = "-";
            }
        }
    };

    // Reset weekly defect tables (WW1 → WW5)
    for (let week = 1; week <= 5; week++) {
        const table = document.querySelector(
            `.tab2-container .table-container:nth-child(${week + 1}) table`
        );
        fillEmpty(table, 5); // Clear top 5 defect rows per week
    }

    // Reset overall defect table (top 10 defects)
    const overallTable = document.querySelector(".overall-container");
    fillEmpty(overallTable, 10); // Clear top 10 rows
}


function getTopDefect(indices) {
    // Guard clause to ensure workbook data is available
    if (!workbook) {
        console.error("Workbook not loaded!");
        return {};
    }

    const sheet = worksheet;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Normalize process row boundaries
    const startIndex = Math.min(indices[0], indices[1]);
    const endIndex = Math.max(indices[0], indices[1]);

    // Column indices corresponding to weekly defect data (Week 1 → Week 5)
    const weekColumns = [65, 66, 67, 68, 69];

    // Locate the start of the machining defect section within the process block
    let startRow = null;
    for (let i = startIndex; i <= endIndex; i++) {
        if (rows[i][0]?.toString().trim() === "MACHINING DEFECT") {
            startRow = i + 1; // Data begins immediately after the header row
            break;
        }
    }

    // Exit early if no machining defect section is found
    if (startRow === null) {
        console.warn("MACHINING DEFECT not found in process block");
        return {};
    }

    const weekDefects = {};     // Stores sorted defect data per week
    const overallDefects = {};  // Accumulates defect totals across all weeks

    weekColumns.forEach((colIndex, i) => {
        const weekKey = `ww${i + 1}`;
        const defects = {};

        // Scan defect rows within the process range
        for (let r = startRow; r <= endIndex; r++) {
            const label = rows[r][0];
            const value = rows[r][colIndex];

            // Capture only valid defect entries with positive counts
            if (label && typeof value === "number" && value > 0) {
                defects[label] = value;
                overallDefects[label] = (overallDefects[label] || 0) + value;
            }
        }

        // Sort all weekly defects in descending order (no slicing applied)
        weekDefects[weekKey] = Object.entries(defects)
            .sort((a, b) => b[1] - a[1])
            .map(([name, value]) => ({ name, value }));
    });

    // Sort all accumulated defects across weeks (no slicing applied)
    weekDefects.overall = Object.entries(overallDefects)
        .sort((a, b) => b[1] - a[1])
        .map(([name, value]) => ({ name, value }));

    return weekDefects;
}

function setTopDefects(topDefects) {
    // Exit if no defect data is available
    if (!topDefects) return;

    // Retrieve total input once for NG% calculation
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
                // Populate defect name and quantity
                row.cells[0].textContent = defect.name.toUpperCase();
                row.cells[1].textContent = defect.value;

                // Calculate and display NG% relative to total input
                if (row.cells[2]) {
                    const ngPct = totalInputs > 0
                        ? ((defect.value / totalInputs) * 100).toFixed(2)
                        : "0.00";
                    row.cells[2].textContent = ngPct;
                }
            } else {
                // Reset unused rows to placeholder values
                row.cells[0].textContent = "-";
                row.cells[1].textContent = "-";
                if (row.cells[2]) row.cells[2].textContent = "-";
            }
        }
    };

    /* ---------- Weekly defect tables (ww1 → ww5) ---------- */
    for (let week = 1; week <= 5; week++) {
        const weekKey = `ww${week}`;
        const weekData = topDefects[weekKey] || [];

        const table = document.querySelector(
            `.tab2-container .table-container:nth-child(${week + 1}) table`
        );
        fillTable(table, weekData, 5);
    }

    /* ---------- Overall defect table (top 10 defects) ---------- */
    if (topDefects.overall) {
        const overallTable = document.querySelector(".overall-container");
        fillTable(overallTable, topDefects.overall, 10);
    }
}

function initEmptyLineChart() {
    // Initialize the line chart with no data to establish baseline configuration
    const ctx = document.getElementById("lineChart").getContext("2d");

    lineChart = new Chart(ctx, {
        type: "line",
        data: {
            labels: [],           // Initialize with an empty X-axis (weeks populated later)
            datasets: [{
                label: "Defect Trend",
                data: [],          // Initialize with no Y-axis values
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
            animation: false,     // Disable animation for smoother frequent updates
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
    // Exit early if required data is not available
    if (!weekDefects || !defectName) return;

    // Extract and sort weekly keys only (exclude aggregate "overall" data)
    const weekKeys = Object.keys(weekDefects)
        .filter(k => k.startsWith("ww"))
        .sort((a, b) => Number(a.slice(2)) - Number(b.slice(2)));

    const labels = [];
    const values = [];
 
    // Build chart labels and values for the selected defect across weeks
    weekKeys.forEach(weekKey => {
        labels.push(weekKey.toUpperCase());

        const weekData = weekDefects[weekKey] || [];
        const match = weekData.find(d => d.name === defectName);

        // Default to zero when the defect is not present for a given week
        values.push(match ? match.value : 0);
    });

    // Update the line chart with the new defect trend data
    updateLineChart(labels, values, defectName);
}



function bindTopDefectRadios(weekDefects) {
    // Ensure overall defect data is available before binding events
    if (!weekDefects || !Array.isArray(weekDefects.overall)) return;

    const radios = document.querySelectorAll('input[name="topDefect"]');

    // Remove any previously bound event listeners by cloning radio elements
    radios.forEach(radio => {
        radio.replaceWith(radio.cloneNode(true));
    });

    const newRadios = document.querySelectorAll('input[name="topDefect"]');
    const topOverall = weekDefects.overall;

    newRadios.forEach((radio, index) => {
        const defect = topOverall[index];
        if (!defect) return;

        // Bind change handler to update chart based on selected defect
        radio.addEventListener("change", () => {
            setDefectLineChart(weekDefects, defect.name);
        });
    });

    // Update radio labels and enabled/disabled state
    updateTopDefectRadios(topOverall);

    // Initialize chart with the highest-ranked defect by default
    if (topOverall.length > 0) {
        setDefectLineChart(weekDefects, topOverall[0].name);
    }
}

function updateLineChart(labels, values, defectName) {
    // Guard clause to ensure chart instance is initialized
    if (!lineChart) return;

    // Update chart axes, dataset label, and values
    lineChart.data.labels = labels;
    lineChart.data.datasets[0].label = defectName.toUpperCase();
    lineChart.data.datasets[0].data = values;

    // Trigger chart re-render
    lineChart.update();
}

function updateTopDefectRadios(top10Defects) {
    const radios = document.querySelectorAll('input[name="topDefect"]');

    // Exit if defect data is not in the expected format
    if (!Array.isArray(top10Defects)) return;

    radios.forEach((radio, index) => {
        const defect = top10Defects[index];

        if (!defect) {
            // Disable radios that do not map to an available defect
            radio.disabled = true;
            radio.checked = false;
        } else {
            // Enable radio and associate it with a defect name
            radio.disabled = false;
            radio.dataset.defectName = defect.name;
        }
    });

    // Auto-select the first enabled defect radio and initialize the chart
    const firstEnabled = Array.from(radios).find(r => !r.disabled);
    if (firstEnabled) {
	firstEnabled.checked = true;
	setDefectLineChart(top10Defects, top10Defects[0].name);
	}
}









