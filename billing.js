let franchiseData = {}; 
let currentProject = null;
let globalRate = 250;
let billingMonth = "APRIL @ 2026";
let issuingMonth = "MAY @ 2026";
let dueDateStr = "15TH MAY, 2026";
let currDateStr = "27/04/26";
let prevDateStr = "31/03/26";

// Handle Excel Upload
function handleFileUpload(event) {
    const files = event.target.files;
    if(files.length === 0) return;

    const tbody = document.getElementById('billing-tbody');
    tbody.innerHTML = '<tr><td colspan="6" class="text-center py-4"><i data-lucide="loader-2" class="w-6 h-6 animate-spin mx-auto text-brand-orange"></i><p class="mt-2 text-sm">Processing Excel files...</p></td></tr>';
    lucide.createIcons();

    let processedCount = 0;
    for(let i=0; i<files.length; i++) {
        const file = files[i];
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            processWorkbook(workbook, file.name);
            
            processedCount++;
            if(processedCount === files.length) {
                renderTree();
                if(Object.keys(franchiseData).length > 0) {
                    const firstFranchise = Object.keys(franchiseData)[0];
                    const firstProject = Object.keys(franchiseData[firstFranchise])[0];
                    selectProject(firstFranchise, firstProject);
                }
            }
        };
        reader.readAsArrayBuffer(file);
    }
}

function processWorkbook(workbook, filename) {
    let projectName = filename.split('[')[1]?.split('-')[0]?.trim() || "Unknown Project";
    let franchiseName = "Urban Gaz - Main Franchise";

    if(!franchiseData[franchiseName]) franchiseData[franchiseName] = {};
    if(!franchiseData[franchiseName][projectName]) franchiseData[franchiseName][projectName] = { flats: [], rate: globalRate, address: "Address not set" };

    let sheetName = workbook.SheetNames.find(s => s.includes('BILL') || s.includes('R, C'));
    if(!sheetName) sheetName = workbook.SheetNames[0];

    const worksheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, {header: 1});

    let currentRate = globalRate;
    let flats = [];

    for(let i=0; i<json.length; i++) {
        let row = json[i];
        if(!row || row.length === 0) continue;

        let rowStr = row.join(' ').toLowerCase();
        if(rowStr.includes('taka/')) {
            let match = rowStr.match(/@\s*([\d.]+)\s*taka/i);
            if(match && match[1]) currentRate = parseFloat(match[1]);
        }

        let sl = parseInt(row[0]);
        if(!isNaN(sl) && row[1] && typeof row[1] === 'string' && row[1].length < 10) {
            let flatNo = row[1];
            let numbers = row.slice(2).filter(c => typeof c === 'number');
            if(numbers.length >= 2) {
                let prevRead = numbers[0];
                let currRead = numbers[1];
                let cons = currRead - prevRead;
                if(cons < 0) cons = 0;
                let total = cons * currentRate * 1.06;

                flats.push({
                    flatNo: flatNo,
                    prevRead: prevRead.toFixed(2),
                    currRead: currRead.toFixed(2),
                    consumption: cons.toFixed(2),
                    total: total.toFixed(2),
                    ownerName: "Customer " + flatNo,
                    phone: "01700000000",
                    nid: "",
                    address: ""
                });
            }
        }
    }

    let custSheet = workbook.SheetNames.find(s => s.includes('CUSTOMER'));
    if(custSheet) {
        const cJson = XLSX.utils.sheet_to_json(workbook.Sheets[custSheet], {header: 1});
        cJson.forEach(row => {
            if(row && row[1] && row[2]) {
                let fno = String(row[1]).trim();
                let flatObj = flats.find(f => f.flatNo === fno);
                if(flatObj) {
                    if(row[2]) flatObj.ownerName = row[2];
                    if(row[3]) flatObj.nid = row[3];
                    if(row[4]) flatObj.phone = row[4];
                }
            }
        });
    }

    if(flats.length > 0) {
        franchiseData[franchiseName][projectName].flats = flats;
        franchiseData[franchiseName][projectName].rate = currentRate;
    }
}

function updateGlobalSettings() {
    globalRate = parseFloat(document.getElementById('setting-rate').value) || 250;
    billingMonth = document.getElementById('setting-bill-month').value || "APRIL @ 2026";
    issuingMonth = document.getElementById('setting-issue-month').value || "MAY @ 2026";
    dueDateStr = document.getElementById('setting-due-date').value || "15TH MAY, 2026";
    currDateStr = document.getElementById('setting-curr-date').value || "27/04/26";
    prevDateStr = document.getElementById('setting-prev-date').value || "31/03/26";
    
    if(currentProject) {
        currentProject.data.rate = globalRate;
        currentProject.data.flats.forEach(f => {
            f.total = (f.consumption * globalRate * 1.06).toFixed(2);
        });
        selectProject(currentProject.fName, currentProject.pName);
    }
}

function renderTree() {
    const treeDiv = document.getElementById('billing-tree');
    treeDiv.innerHTML = '<h3 class="font-bold text-slate-700 text-sm mb-2 border-b border-slate-200 pb-2">Franchise Hierarchy</h3>';

    for(let fName in franchiseData) {
        let fDiv = document.createElement('div');
        fDiv.className = 'mb-2';
        fDiv.innerHTML = `<div class="flex items-center justify-between font-bold text-slate-800 text-sm hover:text-brand-orange">
            <div class="flex items-center gap-2 cursor-pointer" onclick="editFranchise('${fName}')">
                <i data-lucide="building-2" class="w-4 h-4 text-brand-orange"></i> ${fName}
            </div>
        </div>`;
        
        let pList = document.createElement('div');
        pList.className = 'ml-4 pl-2 border-l border-slate-200 mt-1 flex flex-col gap-1';

        for(let pName in franchiseData[fName]) {
            let pDiv = document.createElement('div');
            pDiv.className = 'flex items-center justify-between text-sm text-slate-600 hover:bg-orange-50 py-1 px-2 rounded';
            pDiv.innerHTML = `<div class="flex items-center gap-2 cursor-pointer hover:text-brand-orange flex-1" onclick="selectProject('${fName}', '${pName}')">
                <i data-lucide="home" class="w-3 h-3"></i> ${pName} 
            </div>
            <div class="flex items-center gap-2">
                <span class="bg-slate-200 text-slate-700 text-[10px] px-1.5 rounded">${franchiseData[fName][pName].flats.length} flats</span>
                <i data-lucide="edit" class="w-3 h-3 cursor-pointer hover:text-brand-orange" onclick="editProject('${fName}', '${pName}')"></i>
            </div>`;
            pList.appendChild(pDiv);
        }

        fDiv.appendChild(pList);
        treeDiv.appendChild(fDiv);
    }
    lucide.createIcons();
}

function selectProject(fName, pName) {
    currentProject = { fName, pName, data: franchiseData[fName][pName] };
    
    document.getElementById('billing-table-container').classList.remove('hidden');
    document.getElementById('billing-table-title').textContent = pName;
    document.getElementById('billing-unit-price').textContent = `Rate: ${currentProject.data.rate} BDT/m3`;
    document.getElementById('setting-rate').value = currentProject.data.rate;

    const tbody = document.getElementById('billing-tbody');
    tbody.innerHTML = '';

    currentProject.data.flats.forEach((flat, index) => {
        let tr = document.createElement('tr');
        tr.className = "bg-white hover:bg-slate-50 transition-colors";
        tr.innerHTML = `
            <td class="px-4 py-3 font-medium text-slate-800">
                ${flat.flatNo}
                <div class="text-[10px] text-slate-500">${flat.ownerName}</div>
            </td>
            <td class="px-4 py-3"><input type="number" class="w-20 border rounded p-1 text-xs" value="${flat.prevRead}" onchange="updateFlatData('${fName}', '${pName}', '${flat.flatNo}', 'prevRead', this.value)"></td>
            <td class="px-4 py-3 font-bold"><input type="number" class="w-20 border rounded p-1 text-xs" value="${flat.currRead}" onchange="updateFlatData('${fName}', '${pName}', '${flat.flatNo}', 'currRead', this.value)"></td>
            <td class="px-4 py-3 text-blue-600 font-bold">${flat.consumption}</td>
            <td class="px-4 py-3 text-green-600 font-bold">৳ ${flat.total}</td>
            <td class="px-4 py-3">
                <div class="flex items-center justify-center gap-2">
                    <button onclick='editFlatModal(${JSON.stringify(flat)})' class="p-1.5 bg-slate-50 text-slate-600 hover:bg-slate-100 rounded-lg tooltip-wrapper" title="Edit Flat">
                        <i data-lucide="edit-2" class="w-4 h-4"></i>
                    </button>
                    <button onclick='generatePDF(${JSON.stringify(flat)})' class="p-1.5 bg-red-50 text-red-600 hover:bg-red-100 rounded-lg tooltip-wrapper" title="Generate PDF Bill">
                        <i data-lucide="file-text" class="w-4 h-4"></i>
                    </button>
                    <button onclick='sendSMS(${JSON.stringify(flat)})' class="p-1.5 bg-blue-50 text-blue-600 hover:bg-blue-100 rounded-lg tooltip-wrapper" title="Send SMS Bill">
                        <i data-lucide="message-square" class="w-4 h-4"></i>
                    </button>
                    <button onclick='collectBillModal(${JSON.stringify(flat)})' class="p-1.5 bg-green-50 text-green-600 hover:bg-green-100 rounded-lg tooltip-wrapper" title="Collect Payment">
                        <i data-lucide="credit-card" class="w-4 h-4"></i>
                    </button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    lucide.createIcons();
}

function updateFlatData(fName, pName, flatNo, field, val) {
    let flat = franchiseData[fName][pName].flats.find(f => f.flatNo === flatNo);
    if(flat) {
        flat[field] = parseFloat(val).toFixed(2);
        let cons = parseFloat(flat.currRead) - parseFloat(flat.prevRead);
        if(cons < 0) cons = 0;
        flat.consumption = cons.toFixed(2);
        flat.total = (cons * franchiseData[fName][pName].rate * 1.06).toFixed(2);
        selectProject(fName, pName);
    }
}

function generatePDF(flat) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    function drawHalf(yOffset, isCustomer) {
        // TOP LEFT DETAILS
        doc.setFont("helvetica", "normal");
        doc.setFontSize(10);
        doc.text("NAME : " + currentProject.pName, 15, yOffset + 20);
        doc.text("C/O : " + (currentProject.data.address || "Address not set"), 15, yOffset + 25);
        doc.text("FLAT NO " + flat.flatNo, 15, yOffset + 35);
        doc.text("FLAT OWNER : " + flat.ownerName, 60, yOffset + 35);
        doc.text("BILLING MONTH : " + billingMonth, 15, yOffset + 45);
        doc.text("BILL ISSUING MONTH : " + issuingMonth, 15, yOffset + 50);

        // TOP RIGHT HEADER
        doc.line(140, yOffset + 15, 140, yOffset + 35);
        doc.setFont("helvetica", "bold");
        doc.setFontSize(20);
        doc.text("GAS BILL", 145, yOffset + 23);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(12);
        doc.text(isCustomer ? "CUSTOMER COPY" : "OFFICE COPY", 145, yOffset + 30);

        // TABLE
        const Y_START = yOffset + 60;
        const R_HEIGHT = 7;
        const X0 = 15, X1 = 45, X2 = 75, X3 = 105, X4 = 160, X5 = 195;

        // Draw Outer Box and inner horizontal lines
        doc.rect(X0, Y_START, 180, 70); // 10 rows
        for(let i=1; i<10; i++) {
            doc.line(X0, Y_START + i*R_HEIGHT, X5, Y_START + i*R_HEIGHT);
        }
        
        // Draw Vertical Lines
        doc.line(X1, Y_START, X1, Y_START + 70);
        doc.line(X2, Y_START, X2, Y_START + 70);
        doc.line(X3, Y_START, X3, Y_START + 70);
        doc.line(X4, Y_START, X4, Y_START + 70);

        // CALCULATIONS
        let gasBill = (flat.consumption * currentProject.data.rate).toFixed(2);
        let mf = (flat.consumption * currentProject.data.rate * 0.06).toFixed(2);
        let surcharge = (flat.total * 0.05).toFixed(2);
        let totalDue = (parseFloat(flat.total) + parseFloat(surcharge)).toFixed(2);

        doc.setFontSize(8);

        // Row 1 (Index 0)
        doc.setFont("helvetica", "bold");
        doc.text("FLAT NO", X0 + 2, Y_START + 5);
        doc.text(flat.flatNo, X1 + 2, Y_START + 5);
        doc.text("FLAT OWNER", X2 + 2, Y_START + 5);
        doc.setFont("helvetica", "normal");
        doc.text(flat.ownerName, X3 + 2, Y_START + 5);

        // Row 2
        doc.setFont("helvetica", "bold");
        doc.text("BILLING MONTH", X0 + 2, Y_START + R_HEIGHT + 5);
        doc.setFont("helvetica", "normal");
        doc.text(billingMonth, X1 + 2, Y_START + R_HEIGHT + 5);
        doc.text("PARTICULARS", X3 + 2, Y_START + R_HEIGHT + 5);
        doc.text("TAKA", X4 + 2, Y_START + R_HEIGHT + 5);

        // Row 3
        doc.text("BILL ISSUING MONTH", X0 + 2, Y_START + 2*R_HEIGHT + 5);
        doc.text(issuingMonth, X1 + 2, Y_START + 2*R_HEIGHT + 5);
        doc.text("USED GAS BILL", X3 + 2, Y_START + 2*R_HEIGHT + 5);
        doc.text(gasBill, X4 + 2, Y_START + 2*R_HEIGHT + 5);

        // Row 4
        doc.text("DATE", X1 + 2, Y_START + 3*R_HEIGHT + 5);
        doc.text("METER READING", X2 + 2, Y_START + 3*R_HEIGHT + 5);
        doc.text("MANAGEMENT FEE [6%]", X3 + 2, Y_START + 3*R_HEIGHT + 5);
        doc.text(mf, X4 + 2, Y_START + 3*R_HEIGHT + 5);

        // Row 5
        doc.text("CURRENT READING", X0 + 2, Y_START + 4*R_HEIGHT + 5);
        doc.text(currDateStr, X1 + 2, Y_START + 4*R_HEIGHT + 5);
        doc.text(flat.currRead, X2 + 2, Y_START + 4*R_HEIGHT + 5);
        doc.setFont("helvetica", "bold");
        doc.text("TOTAL", X3 + 2, Y_START + 4*R_HEIGHT + 5);
        doc.text(flat.total, X4 + 2, Y_START + 4*R_HEIGHT + 5);
        doc.setFont("helvetica", "normal");

        // Row 6
        doc.text("PREVIOUS READING", X0 + 2, Y_START + 5*R_HEIGHT + 5);
        doc.text(prevDateStr, X1 + 2, Y_START + 5*R_HEIGHT + 5);
        doc.text(flat.prevRead, X2 + 2, Y_START + 5*R_HEIGHT + 5);
        doc.text("5% SURCHARGE", X3 + 2, Y_START + 5*R_HEIGHT + 5);
        doc.text(surcharge, X4 + 2, Y_START + 5*R_HEIGHT + 5);

        // Row 7
        doc.text("DIFFERENCE [M3]", X0 + 2, Y_START + 6*R_HEIGHT + 5);
        doc.text(flat.consumption, X2 + 2, Y_START + 6*R_HEIGHT + 5);
        doc.text("[AFTER DUE DATE]", X3 + 2, Y_START + 6*R_HEIGHT + 5);

        // Row 8
        doc.text("UNIT PRICE [BDT/M3]", X0 + 2, Y_START + 7*R_HEIGHT + 5);
        doc.text(currentProject.data.rate.toString(), X2 + 2, Y_START + 7*R_HEIGHT + 5);
        doc.setFont("helvetica", "bold");
        doc.text("TOTAL BILL", X3 + 2, Y_START + 7*R_HEIGHT + 5);
        doc.text(totalDue, X4 + 2, Y_START + 7*R_HEIGHT + 5);
        doc.setFont("helvetica", "normal");

        // Row 9
        doc.text("DUE DATE", X0 + 2, Y_START + 8*R_HEIGHT + 5);
        doc.text(dueDateStr, X1 + 2, Y_START + 8*R_HEIGHT + 5);
        doc.text("[AFTER DUE DATE]", X3 + 2, Y_START + 8*R_HEIGHT + 5);

        // Row 10
        doc.setFont("helvetica", "bold");
        doc.text("PLEASE PAY WITHIN DUE DATE FOR UNINTERRUPTED GAS SUPPLY", 105, Y_START + 9*R_HEIGHT + 5, {align: "center"});

        // SIGNATURES
        doc.setFont("helvetica", "normal");
        doc.text("RECEIVED BY", 15, yOffset + 145);
        doc.line(15, yOffset + 140, 50, yOffset + 140);

        doc.text("AUTHORISED SIGNATURE", 150, yOffset + 145);
        doc.line(150, yOffset + 140, 195, yOffset + 140);
    }

    drawHalf(0, true);
    doc.line(0, 148, 210, 148); // dashed line separator
    drawHalf(148, false);

    let fileName = `( ${currentProject.pName} - ${billingMonth.replace(' @ ', ' ')} ) - ${flat.flatNo}.pdf`;
    doc.save(fileName);
}

function sendSMS(flat) {
    const btn = event.currentTarget;
    btn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i>';
    
    let gasBill = (flat.consumption * currentProject.data.rate).toFixed(2);
    let mf = (flat.consumption * currentProject.data.rate * 0.06).toFixed(2);
    let due = "0.00"; 
    
    let bMonth = billingMonth.split(' @ ')[0] || "August";
    let bYear = billingMonth.split(' @ ')[1] ? billingMonth.split(' @ ')[1].slice(2) : "26";
    let projCode = currentProject.pName.substring(0,2).toUpperCase() + bYear;

    let smsBody = `Gas Bill : ${bMonth}, ${bYear}. 
ID : ${projCode}-${flat.flatNo.padStart(4, '0')}

Usage : ${flat.consumption} CM. 
Unit Price : ${currentProject.data.rate} Taka.

Gas : ${gasBill} Taka,
MF : ${mf} Taka &
Due : ${due} Taka 
Total Bill: ${flat.total} Taka. 

For Details : http://urbangaz.net/bill/${flat.flatNo}`;

    setTimeout(() => {
        alert("SMS Sent:\n\n" + smsBody);
        btn.innerHTML = '<i data-lucide="check" class="w-4 h-4 text-green-600"></i>';
        setTimeout(() => {
            btn.innerHTML = '<i data-lucide="message-square" class="w-4 h-4"></i>';
            lucide.createIcons();
        }, 2000);
    }, 1000);
}

// Payment Modal
function collectBillModal(flat) {
    let existingModal = document.getElementById('payment-modal');
    if(existingModal) existingModal.remove();

    const modalHtml = `
    <div id="payment-modal" class="fixed inset-0 z-[100] flex items-center justify-center bg-slate-900/40 backdrop-blur-sm modal-enter transition-opacity">
        <div class="bg-white w-full max-w-md rounded-2xl p-6 relative shadow-xl">
            <button onclick="document.getElementById('payment-modal').remove()" class="absolute top-4 right-4 text-slate-400 hover:text-slate-600"><i data-lucide="x" class="w-5 h-5"></i></button>
            <h3 class="font-bold text-xl text-slate-800 mb-1">Collect Payment</h3>
            <p class="text-sm text-slate-500 mb-6">Flat ${flat.flatNo} • ${flat.ownerName}</p>
            <div class="bg-orange-50 border border-orange-200 rounded-xl p-4 text-center mb-6">
                <div class="text-xs font-bold text-orange-600 uppercase tracking-widest mb-1">Total Due</div>
                <div class="text-3xl font-bold text-slate-800">৳ ${flat.total}</div>
            </div>
            <p class="text-sm font-bold text-slate-700 mb-3">Select Payment Method</p>
            <div class="grid grid-cols-3 gap-3 mb-6">
                <button class="border-2 border-pink-500 bg-pink-50 text-pink-700 font-bold py-2 rounded-xl flex flex-col items-center gap-1">bKash</button>
                <button class="border-2 border-orange-500 bg-orange-50 text-orange-700 font-bold py-2 rounded-xl flex flex-col items-center gap-1">Nagad</button>
                <button class="border-2 border-green-500 bg-green-50 text-green-700 font-bold py-2 rounded-xl flex flex-col items-center gap-1">Cash</button>
            </div>
            <button onclick="confirmPayment(this)" class="w-full bg-brand-orange text-white font-bold py-3 rounded-xl shadow-sm hover:bg-orange-600">Confirm Payment Received</button>
        </div>
    </div>
    `;
    document.body.insertAdjacentHTML('beforeend', modalHtml);
    lucide.createIcons();
}
function confirmPayment(btn) {
    btn.innerHTML = 'Payment Successful!';
    btn.classList.replace('bg-brand-orange', 'bg-green-500');
    setTimeout(() => document.getElementById('payment-modal').remove(), 1000);
}

// Entity Modals
function addFranchiseModal() { showEntityModal('Add Franchise', 'franchise', '', {}); }
function editFranchise(fName) { showEntityModal('Edit Franchise', 'franchise', fName, {name: fName}); }
function editProject(fName, pName) { showEntityModal('Edit Building/Project', 'project', pName, {fName: fName, pName: pName, address: franchiseData[fName][pName].address}); }
function editFlatModal(flat) { showEntityModal('Edit Flat Details', 'flat', flat.flatNo, flat); }

function showEntityModal(title, type, id, data) {
    let existing = document.getElementById('entity-modal');
    if(existing) existing.remove();

    let fields = '';
    if(type === 'franchise') {
        fields = `
            <input type="text" id="ent-name" class="w-full border p-2 rounded mb-2 text-sm" placeholder="Franchise Name" value="${data.name || ''}">
            <input type="text" id="ent-contact" class="w-full border p-2 rounded mb-2 text-sm" placeholder="Manager Contact">
            <input type="text" id="ent-addr" class="w-full border p-2 rounded mb-2 text-sm" placeholder="Address">
            <input type="text" id="ent-tin" class="w-full border p-2 rounded mb-2 text-sm" placeholder="TIN Number">
            <input type="text" id="ent-nid" class="w-full border p-2 rounded mb-2 text-sm" placeholder="Manager NID">
        `;
    } else if (type === 'project') {
        fields = `
            <input type="hidden" id="ent-fname" value="${data.fName}">
            <input type="text" id="ent-pname" class="w-full border p-2 rounded mb-2 text-sm" placeholder="Building Name" value="${data.pName || ''}">
            <input type="text" id="ent-nid" class="w-full border p-2 rounded mb-2 text-sm" placeholder="Manager NID">
            <input type="text" id="ent-phone" class="w-full border p-2 rounded mb-2 text-sm" placeholder="Manager Phone">
            <input type="text" id="ent-addr" class="w-full border p-2 rounded mb-2 text-sm" placeholder="Address" value="${data.address || ''}">
        `;
    } else if (type === 'flat') {
        fields = `
            <input type="hidden" id="ent-fno" value="${data.flatNo}">
            <input type="text" id="ent-tenant" class="w-full border p-2 rounded mb-2 text-sm" placeholder="Tenant Name" value="${data.ownerName || ''}">
            <input type="text" id="ent-contact" class="w-full border p-2 rounded mb-2 text-sm" placeholder="Contact" value="${data.phone || ''}">
            <input type="text" id="ent-nid" class="w-full border p-2 rounded mb-2 text-sm" placeholder="NID" value="${data.nid || ''}">
            <input type="text" id="ent-addr" class="w-full border p-2 rounded mb-2 text-sm" placeholder="Address" value="${data.address || ''}">
        `;
    }

    const modalHtml = `
    <div id="entity-modal" class="fixed inset-0 z-[100] flex items-center justify-center bg-slate-900/40 backdrop-blur-sm modal-enter">
        <div class="bg-white w-full max-w-sm rounded-2xl p-6 relative shadow-xl">
            <h3 class="font-bold text-xl text-slate-800 mb-4">${title}</h3>
            ${fields}
            <div class="flex gap-3 mt-4">
                <button onclick="document.getElementById('entity-modal').remove()" class="flex-1 bg-slate-100 py-2 rounded-xl text-sm font-bold">Cancel</button>
                <button onclick="saveEntity('${type}'); document.getElementById('entity-modal').remove()" class="flex-1 bg-brand-orange text-white py-2 rounded-xl text-sm font-bold">Save</button>
            </div>
        </div>
    </div>
    `;
    document.body.insertAdjacentHTML('beforeend', modalHtml);
}

function saveEntity(type) {
    if(type === 'franchise') {
        let name = document.getElementById('ent-name').value;
        if(name && !franchiseData[name]) {
            franchiseData[name] = {};
            renderTree();
        }
    } else if (type === 'project') {
        let fname = document.getElementById('ent-fname').value;
        let oldPname = document.getElementById('ent-pname').defaultValue;
        let newPname = document.getElementById('ent-pname').value;
        let addr = document.getElementById('ent-addr').value;
        if(franchiseData[fname] && franchiseData[fname][oldPname]) {
            franchiseData[fname][oldPname].address = addr;
            if(oldPname !== newPname && newPname) {
                franchiseData[fname][newPname] = franchiseData[fname][oldPname];
                delete franchiseData[fname][oldPname];
                if(currentProject && currentProject.pName === oldPname) currentProject.pName = newPname;
            }
            renderTree();
        }
    } else if (type === 'flat') {
        let fno = document.getElementById('ent-fno').value;
        if(currentProject) {
            let flat = currentProject.data.flats.find(f => f.flatNo === fno);
            if(flat) {
                flat.ownerName = document.getElementById('ent-tenant').value;
                flat.phone = document.getElementById('ent-contact').value;
                flat.nid = document.getElementById('ent-nid').value;
                flat.address = document.getElementById('ent-addr').value;
                selectProject(currentProject.fName, currentProject.pName);
            }
        }
    }
}
