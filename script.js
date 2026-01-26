/* ================================
   EXISTING SIDEBAR / TAB LOGIC
================================ */
console.log("Script Loaded")
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

    // Get dropdown values
    const values = selects.map(s => s.value);
    if (values.includes("")) {
        alert("Please select all dropdowns");
        return;
    }

    let workbook = null;
    let worksheet = null;

    try {
        // Traverse directories based on dropdowns
        let dirHandle = baseDirHandle;
        for (let val of values) {
            dirHandle = await dirHandle.getDirectoryHandle(val);
        }

        // Open Excel file
        const fileHandle = await dirHandle.getFileHandle("data.xlsx");
        const file = await fileHandle.getFile();
        const arrayBuffer = await file.arrayBuffer();

        // Parse Excel (ONLY workbook + worksheet here)
        workbook = XLSX.read(arrayBuffer, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        worksheet = workbook.Sheets[sheetName];

    } catch (err) {
        console.error("Error opening file:", err);
        alert("File or folder not found. Make sure 'data.xlsx' exists in the selected path.");
        return; // ⛔ stop execution if parsing failed
    }

    // ✅ SAFE ZONE — business logic only
    const processes = getProcesses(workbook);
    console.log(`the processes are: ${processes}`);
    setProcessList(processes);

    // later…
    // const yieldData = getProcessYield(workbook, selectedProcess);
    // const topDefects = getTopDefects(workbook);
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

    // Read raw rows (array-of-arrays so we can use column indexes)
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const processYieldMap = {};

    for (let i = 0; i < rows.length; i++) {
        const colB = rows[i][0]; // Column B
        const colC = rows[i][1]; // Column C (Process Name)
        const colD = rows[i][2]; // Column D (Yield)
        console.log(colB, colB === "Process");
        if (colB === "Process" /*&& colC != null && colD != null*/) {
            
            processYieldMap[colC] = colD;
        }
    }

    return processYieldMap;
}

function setProcessList(processes) {
    const container = document.getElementById("dynamic-container");
    container.innerHTML = "";

    Object.entries(processes).forEach(([processName, yieldValue]) => {
        const row = document.createElement("div");
        row.className = "process-row";

        // Process name (uneditable)
        const processInput = document.createElement("input");
        processInput.className = "process-input";
        processInput.type = "text";
        processInput.value = processName;
        processInput.disabled = true;

        // Yield (uneditable)
        const yieldInput = document.createElement("input");
        yieldInput.className = "yield-input";
        yieldInput.type = "text";
        yieldInput.value = yieldValue;
        yieldInput.disabled = true;

        // Action button
        const btn = document.createElement("button");
        btn.className="process-btn"
        btn.textContent = "Open";
        btn.addEventListener("click", () => {
            console.log("Process selected:", processName);
            // future: open P-chart / details
        });

        row.appendChild(processInput);
        row.appendChild(yieldInput);
        row.appendChild(btn);

        container.appendChild(row);
    });
}
