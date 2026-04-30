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

function addRow(tableId) {
    const tbody = document.querySelector(`#${tableId} tbody`);
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td><input type="text" style="width: 100%; box-sizing: border-box;" placeholder="Key..."></td>
        <td><input type="text" style="width: 100%; box-sizing: border-box;" placeholder="Value..."></td>
        <td><button type="button" onclick="this.parentElement.parentElement.remove()" style="color: red; border: none; background: none; cursor: pointer;">X</button></td>
    `;
    tbody.appendChild(tr);
}

function getTableData(tableId) {
    const data = {};
    const rows = document.querySelectorAll(`#${tableId} tbody tr`);
    if (rows) {
        rows.forEach(row => {
            const key = row.cells[0].querySelector('input').value.trim();
            const val = row.cells[1].querySelector('input').value.trim();
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

        try {
            if (errorBox) errorBox.style.display = 'none';

            // 1. Check Files
            const fileInputElement = document.getElementById('pdfFiles');
            if (!fileInputElement || !fileInputElement.files || fileInputElement.files.length === 0) {
                alert("❌ Please upload at least one PDF file in Step 2.");
                return;
            }
            const files = Array.from(fileInputElement.files);

            // 2. Gather Settings
            const maxPages = document.getElementById('configMaxPages').value || "15";
            const dpi = document.getElementById('configDPI').value || "150";
            const concurrencyLimit = parseInt(document.getElementById('configBatchSize').value) || 5; 
            const aliasesDict = getTableData('aliasTable');
            const customFieldsDict = getTableData('customTable');
            const customColKeys = Object.keys(customFieldsDict);

            // 3. Generate Dynamic Headers for Details Table (Includes QC info at line level)
            let baseHeaders = `<th>File Name</th><th>Page #</th><th>Supplier</th><th>Inv #</th><th>Material</th><th>Description</th><th>Qty</th><th>UOM</th><th>Price</th><th>Line Total</th><th>Inv# Conf</th><th>Total Conf</th><th>Variance</th><th>Proc Time</th><th>Status</th>`;
            customColKeys.forEach(col => { baseHeaders += `<th style="color: #3498db;">${col}</th>`; });
            if (headerRow) headerRow.innerHTML = baseHeaders;

            // 4. Lock UI and Reset
            extractBtn.disabled = true;
            extractBtn.innerText = "⏳ Extracting... Please Wait";
            if (progContainer) progContainer.style.display = 'block';
            if (finalRunStats) finalRunStats.style.display = 'none';
            if (resultsContainer) resultsContainer.style.display = 'none';
            if (detailsBody) detailsBody.innerHTML = '';
            if (summaryBody) summaryBody.innerHTML = '';
            if (progFill) progFill.style.width = '0%';
            
            const total = files.length;
            let processed = 0;
            let totalPagesExtracted = 0;
            const batchId = "BATCH_" + Date.now();
            
            // 5. Start Global Timer
            const startTime = Date.now();
            const timerInterval = setInterval(() => { 
                if (progTimer) progTimer.innerText = `Time: ${Math.floor((Date.now() - startTime)/1000)}s`; 
            }, 1000);

            // 6. THE CONCURRENT LOOP
            for (let i = 0; i < total; i += concurrencyLimit) {
                const chunk = files.slice(i, i + concurrencyLimit);
                if (progStats) progStats.innerText = `Extracting batch... (${processed} of ${total} finished)`;
                
                // Map the chunk into an array of Promises
                const chunkPromises = chunk.map(async (file) => {
                    const formData = new FormData();
                    formData.append('file', file);
                    formData.append('batch_id', batchId);
                    formData.append('max_pages', maxPages);
                    formData.append('dpi', dpi);
                    formData.append('aliases', JSON.stringify(aliasesDict));
                    formData.append('custom_fields', JSON.stringify(customFieldsDict));

                    try {
                        const result = await fetchAPI('/api/extract-single', { method: 'POST', body: formData });
                        
                        if (result.status === 'success' && result.data) {
                            
                            // Append Summary
                            if (summaryBody) {
                                result.data.summary.forEach(row => {
                                    if (row["Total Pages"]) totalPagesExtracted += row["Total Pages"];
                                    const tr = document.createElement('tr');
                                    tr.innerHTML = `<td>${row["File Name"]}</td><td>${row["Vendor Name"]}</td><td>${row["Invoice #"] || 'N/A'}</td><td>${row["Variance"]}</td><td>${row["Proc Time"]}</td><td class="${row["Status"].includes('FAIL') ? 'status-fail' : 'status-pass'}">${row["Status"]}</td>`;
                                    summaryBody.appendChild(tr);
                                });
                            }
                            
                            // Append Line Level Details
                            if (detailsBody) {
                                result.data.details.forEach(row => {
                                    let customCells = "";
                                    customColKeys.forEach(col => customCells += `<td>${row[col] || '-'}</td>`);
                                    const statusClass = row["Status"].includes('FAIL') ? 'status-fail' : 'status-pass';
                                    const tr = document.createElement('tr');
                                    tr.innerHTML = `
                                        <td>${row["File Name"]}</td>
                                        <td>${row["Page #"]||'-'}</td>
                                        <td>${row["Original Supplier"]||'-'}</td>
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
                                        <td class="${statusClass}">${row["Status"]}</td>
                                        ${customCells}
                                    `;
                                    detailsBody.appendChild(tr);
                                });
                            }
                        }
                    } catch (err) {
                        console.error(`File ${file.name} failed:`, err);
                        // We intentionally log but don't break the loop, so one bad PDF doesn't kill the batch
                    }

                    // Update Progress visually after each file in the chunk finishes
                    processed++;
                    if (progFill) progFill.style.width = `${(processed / total) * 100}%`;
                });

                // AWAIT all requests in the current chunk before moving to the next chunk
                await Promise.all(chunkPromises);
            }

            // 7. Cleanup & Save Excel
            clearInterval(timerInterval);
            const totalTimeSeconds = Math.floor((Date.now() - startTime) / 1000);

            if (progStats) progStats.innerText = "Generating Excel File...";
            
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

            // 8. Finish UI State
            if (progStats) progStats.innerText = "✅ Extraction Complete!";
            if (finalRunStats) {
                finalRunStats.innerHTML = `📊 Batch Complete! &nbsp; | &nbsp; Total PDFs: ${total} &nbsp; | &nbsp; Total Pages: ${totalPagesExtracted} &nbsp; | &nbsp; Total Time: ${totalTimeSeconds}s`;
                finalRunStats.style.display = 'block';
            }
            if (resultsContainer) resultsContainer.style.display = 'block';
            
            extractBtn.innerText = "🚀 Step 4: Start Extraction";
            extractBtn.disabled = false;

        } catch (fatalError) {
            if (errorText) errorText.innerText = `Critical UI Error: ${fatalError.message}`;
            if (errorBox) errorBox.style.display = 'block';
            extractBtn.innerText = "🚀 Step 4: Start Extraction";
            extractBtn.disabled = false;
        }
    });
}
