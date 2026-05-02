// ==========================================
// 1. CORE UTILITIES & API FETCHER
// ==========================================

async function fetchAPI(endpoint, options = {}) {
    const response = await fetch(endpoint, options);
    if (!response.ok) {
        let errDetails = "";
        try {
            const errJson = await response.json();
            errDetails = errJson.detail || await response.text();
        } catch(e) {
            errDetails = await response.text();
        }
        throw new Error(errDetails);
    }
    return await response.json();
}

// ==========================================
// 2. SETTINGS TABLES LOGIC (ALIASES & CUSTOM)
// ==========================================

const standardFields = [
    "invoice_number", "date", "vendor_name", "vendor_address", 
    "bill_to", "remit_to", "origin", "destination", "currency", 
    "subtotal", "shipping_name", "shipping_handling", "total_amount",
    "material", "description", "line_origin", "line_destination", 
    "quantity", "uom", "unit_price", "line_total",
    "tax_name", "tax_amount", "fee_name", "fee_amount"
];

const defaultAliases = {
    "invoice_number": "Bill No, Inv #, Invoice No, Reference Number, Document Number",
    "date": "Invoice Date, Billing Date, Document Date, Issue Date",
    "vendor_name": "Supplier, Biller, Merchant, Company",
    "vendor_address": "Supplier Address, Remittance Address",
    "bill_to": "Customer, Sold To, Billed To, Buyer",
    "remit_to": "Pay To, Remittance, Make Checks Payable To",
    "origin": "Ship From, Origin Address, Sender Address",
    "destination": "Ship To, Delivery Address, Consignee",
    "subtotal": "Net Amount, Total Before Tax, Pre-tax Total",
    "shipping_handling": "Freight, Shipping, Handling, Delivery, Postage",
    "total_amount": "Grand Total, Amount Due, Total Payable, Balance Due",
    "material": "Item Code, Part Number, SKU, Product ID",
    "description": "Item Description, Product Name, Details",
    "quantity": "Qty, Ordered, Shipped, Count",
    "uom": "Unit, Measure",
    "unit_price": "Rate, Price, Cost, Unit Cost",
    "line_total": "Amount, Ext Price, Total, Net Price",
    "tax_name": "Tax Type, GST, VAT, QST, HST, TPS, TVQ",
    "tax_amount": "Tax Total, Tax Amount",
    "fee_name": "Charge Name, Fee Type, Surcharge",
    "fee_amount": "Fee Total, Charge Amount"
};

function updateDefaultAlias(selectElement) {
    const selectedField = selectElement.value;
    const inputElement = selectElement.parentElement.nextElementSibling.querySelector('input');
    
    if (inputElement && defaultAliases[selectedField]) {
        inputElement.value = defaultAliases[selectedField];
    } else if (inputElement) {
        inputElement.value = ""; 
    }
}

function addRow(tableId) {
    const tbody = document.querySelector(`#${tableId} tbody`);
    const tr = document.createElement('tr');
    
    let keyHtml = "";
    
    if (tableId === 'aliasTable') {
        let optionsHtml = standardFields.map(f => `<option value="${f}">${f}</option>`).join('');
        keyHtml = `<select style="width: 100%; box-sizing: border-box; padding: 5px;" onchange="updateDefaultAlias(this)">
                        <option value="" disabled selected>Select Standard Field...</option>
                        ${optionsHtml}
                   </select>`;
    } else {
        keyHtml = `<input type="text" style="width: 100%; box-sizing: border-box; padding: 5px;" placeholder="New Field Name...">`;
    }

    tr.innerHTML = `
        <td>${keyHtml}</td>
        <td><input type="text" style="width: 100%; box-sizing: border-box; padding: 5px;" placeholder="Aliases or Instructions..."></td>
        <td><button type="button" onclick="this.parentElement.parentElement.remove()" style="color: red; border: none; background: none; cursor: pointer;">X</button></td>
    `;
    tbody.appendChild(tr);
}

function getTableData(tableId) {
    const data = {};
    const rows = document.querySelectorAll(`#${tableId} tbody tr`);
    if (rows) {
        rows.forEach(row => {
            const keyElement = row.cells[0].querySelector('input, select');
            const valElement = row.cells[1].querySelector('input');
            
            const key = keyElement ? keyElement.value.trim() : '';
            const val = valElement ? valElement.value.trim() : '';
            
            if (key) data[key] = val;
        });
    }
    return data;
}

// ==========================================
// 3. DOCUMENT GRID & NATIVE PDF VIEWER
// ==========================================

const filesInput = document.getElementById('pdfFiles');
const pdfGrid = document.getElementById('pdfGrid');
const gridSearch = document.getElementById('gridSearch');
const pdfModal = document.getElementById('pdfModal');
const modalViewer = document.getElementById('modalViewer');
const closeModal = document.getElementById('closeModal');
const noFilesMsg = document.getElementById('noFilesMsg');

if (filesInput) {
    filesInput.addEventListener('change', () => {
        const files = Array.from(filesInput.files);
        if (files.length > 0) {
            if (noFilesMsg) noFilesMsg.style.display = 'none';
            if (gridSearch) gridSearch.style.display = 'block';
        } else {
            if (noFilesMsg) noFilesMsg.style.display = 'block';
            if (gridSearch) gridSearch.style.display = 'none';
        }
        renderGrid(files);
    });

    if (gridSearch) {
        gridSearch.addEventListener('input', (e) => {
            const term = e.target.value.toLowerCase();
            const files = Array.from(filesInput.files).filter(f => f.name.toLowerCase().includes(term));
            renderGrid(files);
        });
    }
}

function renderGrid(files) {
    if (!pdfGrid) return;
    pdfGrid.innerHTML = '';
    files.forEach(file => {
        const card = document.createElement('div');
        card.className = 'pdf-card';
        card.innerHTML = `<div class="pdf-card-icon">📄</div><div class="pdf-card-title" title="${file.name}">${file.name}</div>`;
        card.addEventListener('click', () => openModal(file));
        pdfGrid.appendChild(card);
    });
}

function openModal(file) {
    if (!modalViewer || !pdfModal) return;
    modalViewer.data = URL.createObjectURL(file);
    pdfModal.style.display = 'flex';
}

if (closeModal) {
    closeModal.addEventListener('click', () => {
        if (!pdfModal || !modalViewer) return;
        pdfModal.style.display = 'none';
        modalViewer.data = '';
    });
}

// ==========================================
// 4. EXTRACTION LOGIC (CONCURRENT BATCHING)
// ==========================================

const extractBtn = document.getElementById('extractBtn');

if (extractBtn) {
    extractBtn.addEventListener('click', async (e) => {
        e.preventDefault();
        
        const progContainer = document.getElementById('progressContainer');
        const progFill = document.getElementById('progressFill');
        const progStats = document.getElementById('progressStats');
        const progTimer = document.getElementById('progressTimer');
        const finalRunStats = document.getElementById('finalRunStats');
        const resultsContainer = document.getElementById('resultsContainer');
        const detailsBody = document.querySelector('#detailsTable tbody');
        const summaryBody = document.querySelector('#summaryTable tbody');
        const headerRow = document.getElementById('detailsHeaderRow');
        const errorBox = document.getElementById('errorAlertBox');
        const errorText = document.getElementById('errorAlertText');

        let batchId = "BATCH_" + Date.now();
        let customColKeys = [];
        let processed = 0;
        let totalPagesExtracted = 0;
        let startTime = Date.now();
        let timerInterval = null;
        let isStopped = false;

        // DYNAMIC STOP BUTTON CREATION
        let stopBtn = document.getElementById('stopBtn');
        if (!stopBtn) {
            stopBtn = document.createElement('button');
            stopBtn.id = 'stopBtn';
            stopBtn.style = "background-color: #e74c3c; color: white; padding: 10px 15px; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; margin-left: 10px;";
            extractBtn.parentNode.insertBefore(stopBtn, extractBtn.nextSibling);
        }
        stopBtn.innerText = "🛑 Stop & Download Partial";
        stopBtn.style.display = 'inline-block';
        stopBtn.disabled = false;
        
        stopBtn.onclick = () => {
            isStopped = true;
            stopBtn.innerText = "⏳ Stopping after current batch...";
            stopBtn.disabled = true;
        };

        try {
            if (errorBox) errorBox.style.display = 'none';

            const fileInputElement = document.getElementById('pdfFiles');
            if (!fileInputElement || !fileInputElement.files || fileInputElement.files.length === 0) {
                alert("❌ Please upload at least one file in Step 2.");
                stopBtn.style.display = 'none';
                return;
            }
            const files = Array.from(fileInputElement.files);

            const maxPages = document.getElementById('configMaxPages').value || "15";
            const dpi = document.getElementById('configDPI').value || "150";
            // Defaulting concurrency to 15
            const concurrencyLimit = parseInt(document.getElementById('configBatchSize').value) || 15; 
            const aliasesDict = getTableData('aliasTable');
            const customFieldsDict = getTableData('customTable');
            customColKeys = Object.keys(customFieldsDict);

            let baseHeaders = `<th>File Name</th><th>Page #</th><th>Inv #</th><th>Material</th><th>Description</th><th>Qty</th><th>UOM</th><th>Price</th><th>Line Total</th><th>Inv# Conf</th><th>Total Conf</th><th>Variance</th><th>Proc Time</th>`;
            customColKeys.forEach(col => { baseHeaders += `<th style="color: #3498db;">${col}</th>`; });
            if (headerRow) headerRow.innerHTML = baseHeaders;

            extractBtn.disabled = true;
            extractBtn.innerText = "⏳ Extracting... Please Wait";
            if (progContainer) progContainer.style.display = 'block';
            if (finalRunStats) finalRunStats.style.display = 'none';
            if (resultsContainer) resultsContainer.style.display = 'none';
            if (detailsBody) detailsBody.innerHTML = '';
            if (summaryBody) summaryBody.innerHTML = '';
            if (progFill) progFill.style.width = '0%';
            
            const total = files.length;
            
            timerInterval = setInterval(() => { 
                if (progTimer) progTimer.innerText = `Time: ${Math.floor((Date.now() - startTime)/1000)}s`; 
            }, 1000);

            // MAIN LOOP
            for (let i = 0; i < total; i += concurrencyLimit) {
                if (isStopped) {
                    if (progStats) progStats.innerText = `Extraction stopped by user. Saving ${processed} files...`;
                    break;
                }

                const chunk = files.slice(i, i + concurrencyLimit);
                if (progStats) progStats.innerText = `Extracting batch... (${processed} of ${total} finished)`;
                
                const chunkPromises = chunk.map(async (file) => {
                    const formData = new FormData();
                    formData.append('file', file);
                    formData.append('batch_id', batchId);
                    formData.append('max_pages', maxPages);
                    formData.append('dpi', dpi);
                    formData.append('aliases', JSON.stringify(aliasesDict));
                    formData.append('custom_fields', JSON.stringify(customFieldsDict));

                    try {
                        const response = await fetchAPI('/api/extract-single', { method: 'POST', body: formData });
                        
                        if (response.status === 'success' && response.data) {
                            
                            if (response.data.total_file_pages) {
                                totalPagesExtracted += response.data.total_file_pages;
                            }
                            
                            if (summaryBody) {
                                response.data.summary.forEach(row => {
                                    const statusClass = row["Status"].includes('NEEDS HUMAN REVIEW') ? 'status-fail' : 'status-pass';
                                    const tr = document.createElement('tr');
                                    tr.innerHTML = `
                                        <td>${row["File Name"]}</td>
                                        <td>${row["Vendor Name"]}</td>
                                        <td>${row["Invoice #"] || 'N/A'}</td>
                                        <td>${row["Variance"]}</td>
                                        <td>${row["Proc Time"]}</td>
                                        <td style="color: #e67e22; font-weight: bold;">${row["LLM Time"]}</td>
                                        <td style="color: #8e44ad; font-weight: bold;">${row["Sec/Page"]}</td>
                                        <td class="${statusClass}">${row["Status"]}</td>
                                    `;
                                    summaryBody.appendChild(tr);
                                });
                            }
                            
                            if (detailsBody) {
                                response.data.details.forEach(row => {
                                    let customCells = "";
                                    customColKeys.forEach(col => customCells += `<td>${row[col] || '-'}</td>`);
                                    
                                    const tr = document.createElement('tr');
                                    tr.innerHTML = `
                                        <td>${row["File Name"]}</td>
                                        <td>${row["Page #"]||'-'}</td>
                                        <td>${row["Invoice Number"]||'-'}</td>
                                        <td>${row["Material"]||'-'}</td>
                                        <td>${row["Description"]}</td>
                                        <td>${row["Qty"]||'-'}</td>
                                        <td>${row["UOM"]||'-'}</td>
                                        <td>${row["Price"]||'-'}</td>
                                        <td>${row["Line Total"]||'-'}</td>
                                        <td>${row["Inv# Conf"]||'-'}</td>
                                        <td>${row["Total Conf"]||'-'}</td>
                                        <td>$${row["Variance"]||'0.00'}</td>
                                        <td>${row["Proc Time"]}</td>
                                        ${customCells}
                                    `;
                                    detailsBody.appendChild(tr);
                                });
                            }
                        }
                    } catch (err) {
                        console.error(`File ${file.name} failed:`, err);
                    }

                    processed++;
                    if (progFill) progFill.style.width = `${(processed / total) * 100}%`;
                });

                await Promise.all(chunkPromises);
            }

        } catch (fatalError) {
            if (errorText) errorText.innerText = `Critical UI Error: ${fatalError.message}`;
            if (errorBox) errorBox.style.display = 'block';
        } finally {
            // FAIL-SAFE EXCEL GENERATOR (Always runs, even if stopped or crashed mid-way)
            clearInterval(timerInterval);
            const totalTimeSeconds = Math.floor((Date.now() - startTime) / 1000);

            if (processed > 0) {
                if (progStats) progStats.innerText = "Generating Excel File for processed items...";
                
                const excelForm = new FormData();
                excelForm.append('batch_id', batchId);
                excelForm.append('custom_fields', JSON.stringify(customColKeys));
                
                try {
                    await fetchAPI('/api/generate-excel', { method: 'POST', body: excelForm });
                    const downloadBtn = document.getElementById('downloadExcelBtn');
                    if (downloadBtn) downloadBtn.href = `/api/download-excel/${batchId}`;
                } catch(e) {
                    console.error("Excel generation failed on server", e);
                }

                if (progStats) progStats.innerText = isStopped ? "⚠️ Partial Extraction Complete!" : "✅ Extraction Complete!";
                if (finalRunStats) {
                    finalRunStats.innerHTML = `📊 Batch Complete! &nbsp; | &nbsp; Files Processed: ${processed} &nbsp; | &nbsp; Total Pages: ${totalPagesExtracted} &nbsp; | &nbsp; Total Time: ${totalTimeSeconds}s`;
                    finalRunStats.style.display = 'block';
                }
                if (resultsContainer) resultsContainer.style.display = 'block';
            } else {
                if (progStats) progStats.innerText = "Extraction stopped before any files finished.";
            }
            
            // Clean up UI buttons
            extractBtn.innerText = "🚀 Step 4: Start Extraction";
            extractBtn.disabled = false;
            stopBtn.style.display = 'none';
        }
    });
}
