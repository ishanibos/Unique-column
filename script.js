// --- GLOBAL STATE ---
let mainData = [];     // Source of truth
let activeData = [];   // Currently viewed/filtered data
let analysisData = []; // Secondary table for analysis results
let headers = [];

let currentView = 'main'; // 'main' or 'analysis'
let formulaMode = 'text'; // 'text', 'math', 'logic'

// --- INIT & UTILS ---
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('fileInput').addEventListener('change', handleFileLoad);
    initUI();
});

function handleFileLoad(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById('fileName').innerText = file.name;
    document.getElementById('fileInfoDisplay').classList.remove('hidden');

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const sheetData = XLSX.utils.sheet_to_json(ws, { defval: "" });

        // --- PRE-PROCESS TO STRINGS ONLY IF NEEDED, BUT KEEP TYPES FOR MATH ---
        // For this tool, we keep native types (number/string) for math ops to work easier
        mainData = sheetData;

        activeData = [...mainData];
        refreshHeaders();
        initUI();
        renderTable();
    };
    reader.readAsArrayBuffer(file);
}

function refreshHeaders() {
    if (activeData.length) headers = Object.keys(activeData[0]);
    updateDropdowns();
    updateCheckboxes();
}

function initUI() {
    document.getElementById('emptyState').classList.add('hidden');
    renderFormulaInputs(); // Setup initial formula UI
}

function updateDropdowns() {
    // Update all column selects
    const selects = ['filterCol', 'sortCol', 'groupByCol', 'calcCol'];
    selects.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        const prev = el.value;
        el.innerHTML = '';
        headers.forEach(h => {
            const opt = document.createElement('option');
            opt.value = h;
            opt.innerText = h;
            el.appendChild(opt);
        });
        if (prev && headers.includes(prev)) el.value = prev;
    });
}

function updateCheckboxes() {
    const container = document.getElementById('hierarchyList');
    if (!container) return;
    container.innerHTML = '';
    headers.forEach(h => {
        container.innerHTML += `
            <div class="flex items-center gap-2 mb-1">
                <input type="checkbox" value="${h}" class="h-chk text-blue-600 rounded">
                <span class="truncate">${h}</span>
            </div>`;
    });
}

function switchTab(tabName) {
    // UI Toggle
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active', 'text-blue-600'));
    // This is a bit hacky selector-wise for compactness
    const btns = document.querySelectorAll('button[onclick^="switchTab"]');
    // Simple way to handle button active state matching the tab name
    // Assuming order: Ops(0), Formulas(1), Analysis(2)
    if (tabName === 'ops' && btns[0]) btns[0].classList.add('active', 'text-blue-600');
    if (tabName === 'formulas' && btns[1]) btns[1].classList.add('active', 'text-blue-600');
    if (tabName === 'analysis' && btns[2]) btns[2].classList.add('active', 'text-blue-600');

    document.querySelectorAll('.tab-section').forEach(s => s.classList.add('hidden'));
    const target = document.getElementById(`tab-${tabName}`);
    if (target) target.classList.remove('hidden');
}

// --- RENDER ---
function renderTable(data = activeData) {
    const Thead = document.getElementById('tableHead');
    const Tbody = document.getElementById('tableBody');

    Thead.innerHTML = '';
    Tbody.innerHTML = '';

    if (!data || data.length === 0) return;

    // Headers
    const cols = Object.keys(data[0]);
    cols.forEach(c => {
        const th = document.createElement('th');
        th.className = "px-4 py-2 border-b border-r border-slate-200 last:border-r-0 whitespace-nowrap";
        th.innerText = c;
        Thead.appendChild(th);
    });

    // Body (Virtualize limit)
    const limit = 200;
    const display = data.slice(0, limit);

    document.getElementById('rowCount').innerText = `${data.length} rows`;
    document.getElementById('fileStats').innerText = `${cols.length} cols`;

    display.forEach((row, idx) => {
        const tr = document.createElement('tr');
        tr.className = idx % 2 ? "bg-slate-50/50 hover:bg-blue-50" : "bg-white hover:bg-blue-50";

        cols.forEach(c => {
            const td = document.createElement('td');
            td.className = "px-4 py-2 border-b border-slate-100 truncate max-w-[200px]";
            let val = row[c];
            // Handle Errors or Objects
            if (typeof val === 'object' && val !== null) val = JSON.stringify(val);
            td.innerText = val === undefined ? "" : val;
            td.title = String(val);
            tr.appendChild(td);
        });
        Tbody.appendChild(tr);
    });

    if (data.length > limit) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td colspan="${cols.length}" class="p-4 text-center text-slate-400 italic">... ${data.length - limit} more rows ...</td>`;
        Tbody.appendChild(tr);
    }
}

// --- OPERATIONS ---
function applyFilter() {
    const col = document.getElementById('filterCol').value;
    const op = document.getElementById('filterOp').value;
    const val = document.getElementById('filterVal').value.toLowerCase();

    activeData = activeData.filter(row => {
        let cell = String(row[col] || "").toLowerCase();
        if (op === 'contains') return cell.includes(val);
        if (op === 'equals') return cell === val;
        if (op === 'starts') return cell.startsWith(val);
        if (op === 'ends') return cell.endsWith(val);
        return true;
    });
    currentView = 'main';
    renderTable();
}

function resetData() {
    activeData = [...mainData];
    currentView = 'main';
    document.getElementById('viewLabel').innerText = "Main Data";
    refreshHeaders();
    renderTable();
}

function applySort() {
    const col = document.getElementById('sortCol').value;
    activeData.sort((a, b) => {
        const va = a[col], vb = b[col];
        // Numeric check
        const na = parseFloat(va), nb = parseFloat(vb);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return String(va).localeCompare(String(vb));
    });
    renderTable();
}

function applyDedup() {
    const seen = new Set();
    const prev = activeData.length;
    // Dedup on ALL columns
    activeData = activeData.filter(r => {
        const values = Object.values(r).map(String).join('||');
        if (seen.has(values)) return false;
        seen.add(values);
        return true;
    });
    alert(`Removed ${prev - activeData.length} duplicates.`);
    renderTable();
}

function applyHierarchy() {
    const inputs = Array.from(document.querySelectorAll('.h-chk:checked'));
    if (!inputs.length) return alert("Select columns for Hierarchy");
    const cols = inputs.map(i => i.value);

    const map = new Map();
    let id = 1;

    activeData = activeData.map(row => {
        const key = cols.map(c => row[c] || "").join('|');
        if (!map.has(key)) map.set(key, id++);
        return { ...row, "UniqueKey": key, "ID": map.get(key) };
    });

    refreshHeaders();
    renderTable();
}

// --- FORMULAS ---
function setFormulaMode(mode) {
    formulaMode = mode;
    document.querySelectorAll('.f-mode').forEach(b => {
        b.classList.remove('bg-indigo-100', 'text-indigo-700');
        b.classList.add('text-slate-400');
    });
    // Highlight active
    const map = { text: 'mode-text', math: 'mode-math', logic: 'mode-logic' };
    const el = document.getElementById(map[mode]);
    if (el) {
        el.classList.add('bg-indigo-100', 'text-indigo-700');
        el.classList.remove('text-slate-400');
    }

    renderFormulaInputs();
}

function renderFormulaInputs() {
    const container = document.getElementById('formulaInputs');
    if (!container) return;
    container.innerHTML = '';

    const colOptions = headers.map(h => `<option value="${h}">${h}</option>`).join('');

    if (formulaMode === 'text') {
        container.innerHTML = `
            <select id="f_func" class="w-full text-xs p-2 border rounded">
                <option value="CONCAT">CONCAT (Join Columns)</option>
                <option value="TRIM">TRIM (Remove Spaces)</option>
                <option value="UPPER">UPPER (Text Case)</option>
                <option value="LOWER">LOWER (Text Case)</option>
                <option value="LEN">LEN (Length)</option>
                <option value="LEFT">LEFT (First N Chars)</option>
                <option value="RIGHT">RIGHT (Last N Chars)</option>
            </select>
            <!-- Params vary by func, simplifying for demo -->
            <div id="f_params" class="space-y-2">
                <select id="f_col1" class="w-full text-xs p-2 border rounded">${colOptions}</select>
                <input type="text" id="f_arg1" placeholder="Optional Arg (e.g. Length)" class="w-full text-xs p-2 border rounded hidden">
                <select id="f_col2" class="w-full text-xs p-2 border rounded h-20 hidden" multiple title="Select multiple for CONCAT">${colOptions}</select>
            </div>
        `;
        // Dynamic show/hide based on selection
        document.getElementById('f_func').addEventListener('change', (e) => {
            const f = e.target.value;
            const arg1 = document.getElementById('f_arg1');
            const col1 = document.getElementById('f_col1');
            const col2 = document.getElementById('f_col2');

            if (f === 'CONCAT') {
                col1.classList.add('hidden');
                arg1.classList.add('hidden');
                col2.classList.remove('hidden');
            } else if (['LEFT', 'RIGHT'].includes(f)) {
                col1.classList.remove('hidden');
                arg1.classList.remove('hidden');
                col2.classList.add('hidden');
            } else {
                col1.classList.remove('hidden');
                arg1.classList.add('hidden');
                col2.classList.add('hidden');
            }
        });
        // Trigger change init
        document.getElementById('f_func').dispatchEvent(new Event('change'));

    } else if (formulaMode === 'math') {
        container.innerHTML = `
            <select id="f_func" class="w-full text-xs p-2 border rounded">
                <option value="ROUND">ROUND</option>
                <option value="CEIL">CEILING (Round Up)</option>
                <option value="FLOOR">FLOOR (Round Down)</option>
                <option value="ABS">ABS (Absolute)</option>
                <option value="ADD">ADD (+)</option>
                <option value="SUB">SUBTRACT (-)</option>
                <option value="MULT">MULTIPLY (*)</option>
                <option value="DIV">DIVIDE (/)</option>
            </select>
            <div class="flex gap-1">
                <select id="f_col1" class="w-1/2 text-xs p-2 border rounded">${colOptions}</select>
                <input type="number" id="f_val_scalar" placeholder="Value" class="w-1/2 text-xs p-2 border rounded">
            </div>
            <!-- Optional second column for arithmetic -->
            <div class="text-center text-[10px] text-slate-400 font-bold my-1">- OR -</div>
            <select id="f_col2" class="w-full text-xs p-2 border rounded text-slate-500">
                <option value="">(Select 2nd Column)</option>
                ${colOptions}
            </select>
        `;

    } else if (formulaMode === 'logic') {
        container.innerHTML = `
            <div class="p-2 bg-slate-50 border rounded text-xs space-y-2">
                <div class="font-bold text-slate-500">IF</div>
                <div class="flex gap-1">
                    <select id="l_col" class="w-1/2 p-1 border rounded">${colOptions}</select>
                    <select id="l_op" class="w-1/2 p-1 border rounded">
                        <option value="==">Equals</option>
                        <option value="!=">Not Equals</option>
                        <option value=">">Greater</option>
                        <option value="<">Less</option>
                        <option value="contains">Contains</option>
                    </select>
                </div>
                <input type="text" id="l_val" placeholder="Value to check" class="w-full p-1 border rounded">
                
                <div class="font-bold text-slate-500 mt-2">THEN</div>
                <input type="text" id="l_true" placeholder="Result if True" class="w-full p-1 border rounded">
                
                <div class="font-bold text-slate-500 mt-2">ELSE</div>
                <input type="text" id="l_false" placeholder="Result if False" class="w-full p-1 border rounded">
            </div>
        `;
    }
}

function applyFormula() {
    const newName = document.getElementById('newColName').value.trim();
    if (!newName) return alert("Enter a column name!");

    activeData = activeData.map(row => {
        let res = "";

        try {
            if (formulaMode === 'text') {
                const func = document.getElementById('f_func').value;
                const c1 = document.getElementById('f_col1').value;
                const val1 = String(row[c1] || "");

                if (func === 'TRIM') res = val1.trim();
                if (func === 'UPPER') res = val1.toUpperCase();
                if (func === 'LOWER') res = val1.toLowerCase();
                if (func === 'LEN') res = val1.length;
                if (func === 'LEFT') {
                    const len = parseInt(document.getElementById('f_arg1').value) || 1;
                    res = val1.substring(0, len);
                }
                if (func === 'RIGHT') {
                    const len = parseInt(document.getElementById('f_arg1').value) || 1;
                    res = val1.slice(-len);
                }
                if (func === 'CONCAT') {
                    const opts = document.getElementById('f_col2').selectedOptions;
                    res = Array.from(opts).map(o => row[o.value] || "").join("");
                }
            }

            else if (formulaMode === 'math') {
                const func = document.getElementById('f_func').value;
                const c1 = document.getElementById('f_col1').value;
                const c2 = document.getElementById('f_col2').value;
                const scalarRaw = document.getElementById('f_val_scalar').value;
                const scalar = scalarRaw ? parseFloat(scalarRaw) : 0;

                let n1 = parseFloat(row[c1]);
                if (isNaN(n1)) n1 = 0;

                let n2 = c2 ? parseFloat(row[c2]) : scalar;
                if (isNaN(n2)) n2 = 0;

                if (func === 'ROUND') res = Math.round(n1);
                if (func === 'CEIL') res = Math.ceil(n1);
                if (func === 'FLOOR') res = Math.floor(n1);
                if (func === 'ABS') res = Math.abs(n1);
                if (func === 'ADD') res = n1 + n2;
                if (func === 'SUB') res = n1 - n2;
                if (func === 'MULT') res = n1 * n2;
                if (func === 'DIV') res = (n2 !== 0) ? (n1 / n2) : 0;
            }

            else if (formulaMode === 'logic') {
                const col = document.getElementById('l_col').value;
                const op = document.getElementById('l_op').value;
                const checkVal = document.getElementById('l_val').value; // string comparison mostly
                const tVal = document.getElementById('l_true').value;
                const fVal = document.getElementById('l_false').value;

                const rowVal = String(row[col] || "");
                let match = false;

                if (op === '==') match = rowVal == checkVal;
                if (op === '!=') match = rowVal != checkVal;
                if (op === 'contains') match = rowVal.includes(checkVal);
                // Numeric logic
                if (op === '>' || op === '<') {
                    const nr = parseFloat(rowVal);
                    const nc = parseFloat(checkVal);
                    if (!isNaN(nr) && !isNaN(nc)) {
                        if (op === '>') match = nr > nc;
                        if (op === '<') match = nr < nc;
                    }
                }

                res = match ? tVal : fVal;
            }
        } catch (e) { res = "#ERROR"; console.error(e); }

        return { ...row, [newName]: res };
    });

    refreshHeaders();
    renderTable();
    alert(`Column '${newName}' created.`);
}

// --- ANALYSIS ---
function runAnalysis() {
    const groupCol = document.getElementById('groupByCol').value;
    const valCol = document.getElementById('calcCol').value;
    const func = document.getElementById('aggFunc').value;

    const groups = {};
    activeData.forEach(row => {
        const key = row[groupCol] || "(Blank)";
        if (!groups[key]) groups[key] = [];
        groups[key].push(row);
    });

    analysisData = Object.keys(groups).map(key => {
        const rows = groups[key];
        const cleanVals = rows.map(r => parseFloat(r[valCol])).filter(n => !isNaN(n));

        let result = 0;
        if (func === 'COUNT') result = rows.length;
        else if (func === 'SUM') result = cleanVals.reduce((a, b) => a + b, 0);
        else if (func === 'AVERAGE') result = cleanVals.length ? (cleanVals.reduce((a, b) => a + b, 0) / cleanVals.length) : 0;
        else if (func === 'MIN') result = Math.min(...cleanVals);
        else if (func === 'MAX') result = Math.max(...cleanVals);

        return { [groupCol]: key, [`${func}_${valCol}`]: result };
    });

    currentView = 'analysis';
    document.getElementById('viewLabel').innerText = `Analysis: ${func} by ${groupCol}`;
    renderTable(analysisData);
}

function runUnique() {
    // Extract unique of FIRST selected filter col or ask user?
    // Simple approach: prompt or use current filterCol dropdown which is convenient
    const col = document.getElementById('filterCol').value || headers[0];
    const set = new Set(activeData.map(r => r[col]));
    analysisData = Array.from(set).map(v => ({ "Unique Value": v }));

    currentView = 'analysis';
    document.getElementById('viewLabel').innerText = `Unique: ${col}`;
    renderTable(analysisData);
}

function runTranspose() {
    // Limited transpose (max 100 rows for sanity)
    const limit = Math.min(activeData.length, 50);
    const subset = activeData.slice(0, limit);

    // New headers = Field Name, Row 1, Row 2...
    const transData = headers.map(h => {
        const rowObj = { "Field": h };
        subset.forEach((r, i) => { rowObj[`Row_${i + 1}`] = r[h]; });
        return rowObj;
    });

    analysisData = transData;
    currentView = 'analysis';
    document.getElementById('viewLabel').innerText = `Transpose (First ${limit} rows)`;
    renderTable(analysisData);
}

// --- EXPORT ---
function exportData() {
    const dataToExport = currentView === 'main' ? activeData : analysisData;
    if (!dataToExport || !dataToExport.length) return alert("No data");

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Export");
    XLSX.writeFile(wb, "Export.xlsx");
}
