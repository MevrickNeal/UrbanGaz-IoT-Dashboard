let franchiseData = {}; 
let currentProject = null;

// Handle Excel Upload
function handleFileUpload(event) {
    const files = event.target.files;
    if(files.length === 0) return;

    // Simulate loading
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
                    // Auto select first project
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
    // Extract project name from filename (e.g., "[MAHFUZ RESIDENCE - ...]")
    let projectName = filename.split('[')[1]?.split('-')[0]?.trim() || "Unknown Project";
    let franchiseName = "Urban Gaz - Main Franchise";

    if(!franchiseData[franchiseName]) franchiseData[franchiseName] = {};
    if(!franchiseData[franchiseName][projectName]) franchiseData[franchiseName][projectName] = { flats: [], rate: 250 };

    // Try to find R, C & BILL sheet
    let sheetName = workbook.SheetNames.find(s => s.includes('BILL') || s.includes('R, C'));
    if(!sheetName) sheetName = workbook.SheetNames[0];

    const worksheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, {header: 1});

    // Parse logic based on our earlier exploration
    let currentRate = 250; // Default
    let flats = [];

    for(let i=0; i<json.length; i++) {
        let row = json[i];
        if(!row || row.length === 0) continue;

        // Try to find rate
        let rowStr = row.join(' ').toLowerCase();
        if(rowStr.includes('taka/')) {
            let match = rowStr.match(/@\s*([\d.]+)\s*taka/i);
            if(match && match[1]) currentRate = parseFloat(match[1]);
        }

        // Try to find flat data: SL, Flat, Readings
        let sl = parseInt(row[0]);
        if(!isNaN(sl) && row[1] && typeof row[1] === 'string' && row[1].length < 10) {
            let flatNo = row[1];
            // Find prev reading and current reading.
            let numbers = row.slice(2).filter(c => typeof c === 'number');
            if(numbers.length >= 2) {
                let prevRead = numbers[0];
                let currRead = numbers[1];
                let cons = currRead - prevRead;
                if(cons < 0) cons = 0; // fallback
                let total = cons * currentRate * 1.06; // + 6% MF

                flats.push({
                    flatNo: flatNo,
                    prevRead: prevRead.toFixed(2),
                    currRead: currRead.toFixed(2),
                    consumption: cons.toFixed(2),
                    total: total.toFixed(2),
                    ownerName: "Customer " + flatNo,
                    phone: "01700000000"
                });
            }
        }
    }

    // Try CUSTOMER DETAILS sheet to get real names
    let custSheet = workbook.SheetNames.find(s => s.includes('CUSTOMER'));
    if(custSheet) {
        const cJson = XLSX.utils.sheet_to_json(workbook.Sheets[custSheet], {header: 1});
        cJson.forEach(row => {
            if(row && row[1] && row[2]) {
                let fno = String(row[1]).trim();
                let flatObj = flats.find(f => f.flatNo === fno);
                if(flatObj) {
                    if(row[2]) flatObj.ownerName = row[2];
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

function renderTree() {
    const treeDiv = document.getElementById('billing-tree');
    treeDiv.innerHTML = '<h3 class="font-bold text-slate-700 text-sm mb-2 border-b border-slate-200 pb-2">Franchise Hierarchy</h3>';

    for(let fName in franchiseData) {
        let fDiv = document.createElement('div');
        fDiv.className = 'mb-2';
        fDiv.innerHTML = `<div class="flex items-center gap-2 font-bold text-slate-800 text-sm cursor-pointer hover:text-brand-orange">
            <i data-lucide="building-2" class="w-4 h-4 text-brand-orange"></i> ${fName}
        </div>`;
        
        let pList = document.createElement('div');
        pList.className = 'ml-4 pl-2 border-l border-slate-200 mt-1 flex flex-col gap-1';

        for(let pName in franchiseData[fName]) {
            let pDiv = document.createElement('div');
            pDiv.className = 'flex items-center gap-2 text-sm text-slate-600 cursor-pointer hover:text-brand-orange py-1 px-2 rounded hover:bg-orange-50';
            pDiv.innerHTML = `<i data-lucide="home" class="w-3 h-3"></i> ${pName} <span class="bg-slate-200 text-slate-700 text-[10px] px-1.5 rounded ml-auto">${franchiseData[fName][pName].flats.length} flats</span>`;
            pDiv.onclick = () => selectProject(fName, pName);
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
            <td class="px-4 py-3">${flat.prevRead}</td>
            <td class="px-4 py-3 font-bold">${flat.currRead}</td>
            <td class="px-4 py-3 text-blue-600 font-bold">${flat.consumption}</td>
            <td class="px-4 py-3 text-green-600 font-bold">৳ ${flat.total}</td>
            <td class="px-4 py-3">
                <div class="flex items-center justify-center gap-2">
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

function generatePDF(flat) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    // Urban Gaz Header
    doc.setFontSize(22);
    doc.setTextColor(241, 90, 36); // Brand Orange
    doc.text("URBAN GAZ", 105, 20, { align: "center" });
    
    doc.setFontSize(10);
    doc.setTextColor(100, 100, 100);
    doc.text("Smart LPG Control & Distribution", 105, 26, { align: "center" });

    // Bill Details
    doc.setFontSize(14);
    doc.setTextColor(50, 50, 50);
    doc.text("MONTHLY GAS BILL", 105, 40, { align: "center" });

    doc.setFontSize(11);
    doc.text(`Project: ${currentProject.pName}`, 20, 55);
    doc.text(`Flat Number: ${flat.flatNo}`, 20, 62);
    doc.text(`Customer Name: ${flat.ownerName}`, 20, 69);
    doc.text(`Contact: ${flat.phone}`, 20, 76);

    doc.text(`Date: ${new Date().toLocaleDateString()}`, 150, 55);

    // Table
    doc.autoTable({
        startY: 85,
        head: [['Description', 'Value']],
        body: [
            ['Previous Meter Reading (m3)', flat.prevRead],
            ['Current Meter Reading (m3)', flat.currRead],
            ['Gas Consumption (m3)', flat.consumption],
            ['Unit Price (BDT)', currentProject.data.rate],
            ['Maintenance Fee (6%)', (flat.consumption * currentProject.data.rate * 0.06).toFixed(2)],
        ],
        foot: [['TOTAL BILL', 'BDT ' + flat.total]],
        theme: 'striped',
        headStyles: { fillColor: [241, 90, 36] },
        footStyles: { fillColor: [50, 50, 50] },
    });

    // Save
    doc.save(`UrbanGaz_Bill_${currentProject.pName}_Flat${flat.flatNo}.pdf`);
}

function sendSMS(flat) {
    const btn = event.currentTarget;
    btn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i>';
    
    // Simulate SMS API call
    setTimeout(() => {
        alert(`SMS sent to ${flat.phone} successfully!\nMessage: Dear ${flat.ownerName}, your Urban Gaz bill for Flat ${flat.flatNo} is BDT ${flat.total}. Please pay soon.`);
        btn.innerHTML = '<i data-lucide="check" class="w-4 h-4 text-green-600"></i>';
        setTimeout(() => {
            btn.innerHTML = '<i data-lucide="message-square" class="w-4 h-4"></i>';
            lucide.createIcons();
        }, 2000);
    }, 1000);
}

// Payment Modal Logic
function collectBillModal(flat) {
    let existingModal = document.getElementById('payment-modal');
    if(existingModal) existingModal.remove();

    const modalHtml = `
    <div id="payment-modal" class="fixed inset-0 z-[100] flex items-center justify-center bg-slate-900/40 backdrop-blur-sm modal-enter transition-opacity">
        <div class="bg-white w-full max-w-md rounded-2xl p-6 relative shadow-xl transform transition-transform">
            <button onclick="document.getElementById('payment-modal').remove()" class="absolute top-4 right-4 text-slate-400 hover:text-slate-600"><i data-lucide="x" class="w-5 h-5"></i></button>
            
            <h3 class="font-bold text-xl text-slate-800 mb-1">Collect Payment</h3>
            <p class="text-sm text-slate-500 mb-6">Flat ${flat.flatNo} • ${flat.ownerName}</p>

            <div class="bg-orange-50 border border-orange-200 rounded-xl p-4 text-center mb-6">
                <div class="text-xs font-bold text-orange-600 uppercase tracking-widest mb-1">Total Due</div>
                <div class="text-3xl font-bold text-slate-800">৳ ${flat.total}</div>
            </div>

            <p class="text-sm font-bold text-slate-700 mb-3">Select Payment Method</p>
            <div class="grid grid-cols-3 gap-3 mb-6">
                <button class="border-2 border-pink-500 bg-pink-50 text-pink-700 font-bold py-2 rounded-xl flex flex-col items-center gap-1 hover:bg-pink-100 transition-colors">
                    <img src="https://freelogopng.com/images/all_img/1656234745bkash-app-logo-png.png" class="h-6 object-contain">
                    <span class="text-xs">bKash</span>
                </button>
                <button class="border-2 border-orange-500 bg-orange-50 text-orange-700 font-bold py-2 rounded-xl flex flex-col items-center gap-1 hover:bg-orange-100 transition-colors">
                    <img src="https://downloadr2.apkmirror.com/wp-content/uploads/2019/08/5d54a24329d95.png" class="h-6 object-contain rounded">
                    <span class="text-xs">Nagad</span>
                </button>
                <button class="border-2 border-green-500 bg-green-50 text-green-700 font-bold py-2 rounded-xl flex flex-col items-center gap-1 hover:bg-green-100 transition-colors">
                    <i data-lucide="banknote" class="w-6 h-6"></i>
                    <span class="text-xs">Cash</span>
                </button>
            </div>

            <button onclick="confirmPayment(this)" class="w-full bg-brand-orange text-white font-bold py-3 rounded-xl shadow-sm hover:bg-orange-600 transition-colors">Confirm Payment Received</button>
        </div>
    </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);
    lucide.createIcons();
}

function confirmPayment(btn) {
    btn.innerHTML = '<i data-lucide="loader-2" class="w-5 h-5 animate-spin mx-auto"></i>';
    setTimeout(() => {
        btn.innerHTML = 'Payment Successful!';
        btn.classList.replace('bg-brand-orange', 'bg-green-500');
        btn.classList.replace('hover:bg-orange-600', 'hover:bg-green-600');
        setTimeout(() => {
            document.getElementById('payment-modal').remove();
        }, 1000);
    }, 1000);
}

// Add Franchise Modal
function addFranchiseModal() {
    let existingModal = document.getElementById('franchise-modal');
    if(existingModal) existingModal.remove();

    const modalHtml = `
    <div id="franchise-modal" class="fixed inset-0 z-[100] flex items-center justify-center bg-slate-900/40 backdrop-blur-sm modal-enter">
        <div class="bg-white w-full max-w-sm rounded-2xl p-6 relative shadow-xl">
            <h3 class="font-bold text-xl text-slate-800 mb-4">Add Entity</h3>
            <div class="space-y-4">
                <div>
                    <label class="block text-xs font-bold text-slate-500 mb-1">Entity Type</label>
                    <select class="w-full border border-slate-200 rounded-lg p-2 text-sm outline-none focus:border-brand-orange">
                        <option>Franchise</option>
                        <option>Building/Project</option>
                        <option>Flat/User</option>
                    </select>
                </div>
                <div>
                    <label class="block text-xs font-bold text-slate-500 mb-1">Name / ID</label>
                    <input type="text" class="w-full border border-slate-200 rounded-lg p-2 text-sm outline-none focus:border-brand-orange" placeholder="Enter name...">
                </div>
            </div>
            <div class="flex gap-3 mt-6">
                <button onclick="document.getElementById('franchise-modal').remove()" class="flex-1 bg-slate-100 text-slate-700 font-bold py-2 rounded-xl">Cancel</button>
                <button onclick="document.getElementById('franchise-modal').remove(); alert('Entity Added!')" class="flex-1 bg-brand-orange text-white font-bold py-2 rounded-xl shadow-sm">Save</button>
            </div>
        </div>
    </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);
}
