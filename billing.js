let franchiseData = {}; 
let currentProject = null;
let globalRate = 250;
let billingMonth = "APRIL @ 2026";
let issuingMonth = "MAY @ 2026";
let dueDateStr = "15TH MAY, 2026";

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
    
    if(currentProject) {
        currentProject.data.rate = globalRate;
        // Recalculate totals
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
        doc.setFont("helvetica", "bold");
        doc.setFontSize(16);
        doc.text("GAS BILL", 105, yOffset + 15, { align: "center" });
        doc.setFontSize(10);
        doc.text(isCustomer ? "CUSTOMER COPY" : "OFFICE COPY", 105, yOffset + 22, { align: "center" });

        doc.setFontSize(10);
        doc.text("NAME : " + currentProject.pName, 15, yOffset + 35);
        doc.text("C/O : " + currentProject.data.address, 15, yOffset + 40);
        
        doc.text("FLAT NO: " + flat.flatNo, 15, yOffset + 50);
        doc.text("FLAT OWNER: " + flat.ownerName, 60, yOffset + 50);

        doc.text("BILLING MONTH: " + billingMonth, 15, yOffset + 60);
        doc.text("BILL ISSUING MONTH: " + issuingMonth, 15, yOffset + 65);

        // Borders
        doc.rect(15, yOffset + 70, 180, 50);
        doc.line(15, yOffset + 77, 195, yOffset + 77); // Header bottom
        doc.line(105, yOffset + 70, 105, yOffset + 120); // Center split

        // Left Table (Reading)
        doc.setFontSize(8);
        doc.text("DATE", 20, yOffset + 75);
        doc.text("METER READING", 60, yOffset + 75);
        
        doc.text("CURRENT READING", 20, yOffset + 85);
        doc.text("27/04/26", 60, yOffset + 85);
        doc.text(flat.currRead, 85, yOffset + 85);

        doc.text("PREVIOUS READING", 20, yOffset + 95);
        doc.text("31/03/26", 60, yOffset + 95);
        doc.text(flat.prevRead, 85, yOffset + 95);

        doc.text("DIFFERENCE [M3]", 20, yOffset + 105);
        doc.text(flat.consumption, 85, yOffset + 105);

        doc.text("UNIT PRICE [BDT/M3]", 20, yOffset + 115);
        doc.text(currentProject.data.rate.toString(), 85, yOffset + 115);

        // Right Table (Particulars)
        doc.text("PARTICULARS", 110, yOffset + 75);
        doc.text("TAKA", 175, yOffset + 75);

        let gasBill = (flat.consumption * currentProject.data.rate).toFixed(2);
        let mf = (gasBill * 0.06).toFixed(2);

        doc.text("USED GAS BILL", 110, yOffset + 85);
        doc.text(gasBill, 175, yOffset + 85);

        doc.text("MANAGEMENT FEE [6%]", 110, yOffset + 95);
        doc.text(mf, 175, yOffset + 95);

        doc.text("TOTAL", 110, yOffset + 105);
        doc.text(flat.total, 175, yOffset + 105);

        let surcharge = (flat.total * 0.05).toFixed(2);
        let totalDue = (parseFloat(flat.total) + parseFloat(surcharge)).toFixed(2);

        doc.text("5% SURCHARGE [AFTER DUE DATE]", 110, yOffset + 112);
        doc.text(surcharge, 175, yOffset + 112);

        doc.text("TOTAL BILL [AFTER DUE DATE]", 110, yOffset + 118);
        doc.text(totalDue, 175, yOffset + 118);

        // Footers
        doc.text("DUE DATE: " + dueDateStr, 15, yOffset + 130);
        doc.text("RECEIVED BY", 15, yOffset + 145);
        doc.line(15, yOffset + 140, 50, yOffset + 140);

        doc.text("AUTHORISED SIGNATURE", 150, yOffset + 145);
        doc.line(150, yOffset + 140, 195, yOffset + 140);

        doc.setFont("helvetica", "italic");
        doc.text("PLEASE PAY WITHIN DUE DATE FOR UNINTERRUPTED GAS SUPPLY", 105, yOffset + 155, { align: "center" });
    }

    drawHalf(0, true);
    doc.line(0, 160, 210, 160); // dashed line separator
    drawHalf(165, false);

    let fileName = `( ${currentProject.pName} - ${billingMonth.replace(' @ ', ' ')} ) - ${flat.flatNo}.pdf`;
    doc.save(fileName);
}

function sendSMS(flat) {
    const btn = event.currentTarget;
    btn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i>';
    
    let gasBill = (flat.consumption * currentProject.data.rate).toFixed(2);
    let mf = (gasBill * 0.06).toFixed(2);
    let due = "0.00"; 
    
    let bMonth = billingMonth.split(' @ ')[0] || "August";
    let bYear = billingMonth.split(' @ ')[1] ? billingMonth.split(' @ ')[1].slice(2) : "26";
    let projCode = currentProject.pName.substring(0,2).toUpperCase() + bYear;

    let smsBody = `Gas Bill : ${bMonth}, ${bYear}. \nID : ${projCode}-${flat.flatNo.padStart(4, '0')}\n\nUsage : ${flat.consumption} CM. \nUnit Price : ${currentProject.data.rate} Taka.\n\nGas : ${gasBill} Taka,\nMF : ${mf} Taka &\nDue : ${due} Taka \nTotal Bill: ${flat.total} Taka. \n\nFor Details : http://urbangaz.net/bill/${flat.flatNo}`;

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
