// --- Core API Fetcher ---
async function fetchAPI(endpoint, options = {}) {
    const response = await fetch(endpoint, options);
    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Server Error (${response.status}): ${errorText}`);
    }
    return await response.json();
}

// --- Settings Tables Logic ---
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

// --- Document Grid & Viewer Logic ---
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
            noFilesMsg.style.display = 'none';
            gridSearch.style.display = 'block';
        } else {
            noFilesMsg.style.display = 'block';
            gridSearch.style.display = 'none';
        }
        renderGrid(files);
    });

    gridSearch.addEventListener('input', (e) => {
        const term = e.target.value.toLowerCase();
        const files = Array.from(filesInput.files).filter(f => f.name.toLowerCase().includes(term));
        renderGrid(files);
    });
}

function renderGrid(files) {
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
    modalViewer.data = URL.createObjectURL(file);
    pdfModal.style.display = 'flex';
}
if(closeModal) {
    closeModal.addEventListener('click', () => {
        pdfModal.style.display = 'none';
        modalViewer.data = '';
    });
}

// --- Main Extraction Logic ---
const extractBtn = document.getElementById('extractBtn');

if (extractBtn) {
    // Note: We are listening to the BUTTON click now, bypassing form validation completely.
    extractBtn.addEventListener('click', async (e) => {
        e.preventDefault();
        
        try {
            const progContainer = document.getElementById('progressContainer');
            const progFill = document.getElementById('progressFill');
            const progStats = document.getElementById('progressStats');
            const progTimer = document.getElementById('progressTimer');
            const finalRunStats = document.getElementById('finalRunStats');
            const resultsContainer = document.getElementById('resultsContainer');
            const detailsBody = document.querySelector('#detailsTable tbody');
            const summaryBody = document.querySelector('#summaryTable tbody');
            const headerRow = document.getElementById('detailsHeaderRow');
            
            // 1. Check Files
            const fileInputElement = document.getElementById('pdfFiles');
            if (!fileInputElement || !fileInputElement.files || fileInputElement.files.length === 0) {
                alert("❌ Please upload at least one PDF file in Step 2.");
                return;
            }
            const files = Array.from(fileInputElement.files);

            // 2. Gather Settings
            const maxPages = document.getElementById('configMaxPages').value || "15";
            const dpi = document.getElementById('configDPI').value || "300";
            const aliasesDict = getTableData('aliasTable');
            const customFieldsDict = getTableData('customTable');
            const customColKeys = Object.keys(customFieldsDict);

            // 3. Generate Dynamic Headers
            let baseHeaders = `<th>File Name</th><th>Page #</th><th>Supplier</th><th>Inv #</th><th>Material</th><th>Description</th><th>Qty</th><th>UOM</th><th>Price</th><th>Line Total</th><th>Origin</th><th>Dest</th>`;
            customColKeys.forEach(col => baseHeaders += `<th style="color: #3498db;">${col}</th>`);
            headerRow.innerHTML = baseHeaders;

            // 4. Lock UI and show Progress
            extractBtn.disabled = true;
            extractBtn.innerText = "⏳ Extracting... Please Wait";
            progContainer.style.display = 'block';
            finalRunStats.style.display = 'none';
            resultsContainer.style.display = 'none';
            detailsBody.innerHTML = '';
            summaryBody.innerHTML = '';
            progFill.style.width = '0%';
            
            const total = files.length;
            let processed = 0;
            let totalPagesExtracted = 0;
            const batchId = "BATCH_" + Date.now();
            
            // 5. Start Global Timer
            const startTime = Date.now();
            const timerInterval = setInterval(() => { 
                progTimer.innerText = `Time: ${Math.floor((Date.now() - startTime)/1000)}s`; 
            }, 1000);

            // 6. Sequential processing loop
            for (let i = 0; i < total; i++) {
                progStats.innerText = `Extracting ${i + 1} of ${total}: ${files[i].name}...`;
                
                const formData = new FormData();
                formData.append('file', files[i]);
                formData.append('batch_id', batchId);
                formData.append('max_pages', maxPages);
                formData.append('dpi', dpi);
                formData.append('aliases', JSON.stringify(aliasesDict));
                formData.append('custom_fields', JSON.stringify(customFieldsDict));

                try {
                    const response = await fetch('/api/extract-single', { method: 'POST', body: formData });
                    
                    if (!response.ok) {
                        const errDetails = await response.text();
                        console.error(`Failed on ${files[i].name}. Error: ${errDetails}`);
                        continue;
                    }

                    const result = await response.json();
                    
                    if (result.status === 'success' && result.data) {
                        // Append Summary Row
                        result.data.summary.forEach(row => {
                            if (row["Total Pages"]) totalPagesExtracted += row["Total Pages"];
                            
                            const tr = document.createElement('tr');
                            tr.innerHTML = `<td>${row["File Name"]}</td><td>${row["Vendor Name"]}</td><td>${row["Invoice #"] || 'N/A'}</td><td>${row["Variance"]}</td><td>${row["Proc Time"]}</td><td class="${row["Status"].includes('FAIL') ? 'status-fail' : 'status-pass'}">${row["Status"]}</td>`;
                            summaryBody.appendChild(tr);
                        });
                        
                        // Append Details Rows
                        result.data.details.forEach(row => {
                            let customCells = "";
                            customColKeys.forEach(col => customCells += `<td>${row[col] || '-'}</td>`);
                            
                            const tr = document.createElement('tr');
                            tr.innerHTML = `<td>${row["File Name"]}</td><td>${row["Page #"]||'-'}</td><td>${row["Original Supplier"]||'-'}</td><td>${row["Invoice Number"]||'-'}</td><td>${row["Material"]||'-'}</td><td>${row["Description"]}</td><td>${row["Qty"]||'-'}</td><td>${row["UOM"]||'-'}</td><td>${row["Price"]||'-'}</td><td>${row["Line Total"]||'-'}</td><td>${row["Origin"]||'-'}</td><td>${row["Dest"]||'-'}</td>${customCells}`;
                            detailsBody.appendChild(tr);
                        });
                    }
                } catch (err) {
                    console.error("API Call failed for:", files[i].name, err);
                }

                // Update Progress Bar
                processed++;
                progFill.style.width = `${(processed / total) * 100}%`;
            }

            // 7. End processing
            clearInterval(timerInterval);
            const totalTimeSeconds = Math.floor((Date.now() - startTime) / 1000);

            progStats.innerText = "Generating Excel File...";
            
            // 8. Trigger Excel Save
            const excelForm = new FormData();
            excelForm.append('batch_id', batchId);
            excelForm.append('custom_fields', JSON.stringify(customColKeys));
            
            try {
                await fetch('/api/generate-excel', { method: 'POST', body: excelForm });
                document.getElementById('downloadExcelBtn').href = `/api/download-excel/${batchId}`;
            } catch(e) {
                console.error("Excel generation failed", e);
            }

            // 9. Finish UI State
            progStats.innerText = "✅ Complete!";
            finalRunStats.innerHTML = `📊 Batch Complete! &nbsp; | &nbsp; Total PDFs: ${total} &nbsp; | &nbsp; Total Pages: ${totalPagesExtracted} &nbsp; | &nbsp; Total Time: ${totalTimeSeconds}s`;
            finalRunStats.style.display = 'block';
            resultsContainer.style.display = 'block';
            extractBtn.innerText = "🚀 Step 4: Start Extraction";
            extractBtn.disabled = false;

        } catch (fatalError) {
            alert("A fatal UI error occurred:\n" + fatalError.message);
            console.error(fatalError);
            extractBtn.innerText = "🚀 Step 4: Start Extraction";
            extractBtn.disabled = false;
        }
    });
}
