async function fetchAPI(endpoint, options = {}) {
    const response = await fetch(endpoint, options);
    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Server Error (${response.status}): ${errorText}`);
    }
    return await response.json();
}

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
    document.querySelectorAll(`#${tableId} tbody tr`).forEach(row => {
        const key = row.cells[0].querySelector('input').value.trim();
        const val = row.cells[1].querySelector('input').value.trim();
        if (key) data[key] = val;
    });
    return data;
}

// --- Viewer & Grid Logic ---
const filesInput = document.getElementById('pdfFiles');
const pdfGrid = document.getElementById('pdfGrid');
const gridSearch = document.getElementById('gridSearch');
const pdfModal = document.getElementById('pdfModal');
const modalViewer = document.getElementById('modalViewer');

if (filesInput) {
    filesInput.addEventListener('change', () => {
        const files = Array.from(filesInput.files);
        if (files.length > 0) {
            document.getElementById('noFilesMsg').style.display = 'none';
            gridSearch.style.display = 'block';
        } else {
            document.getElementById('noFilesMsg').style.display = 'block';
            gridSearch.style.display = 'none';
        }
        renderGrid(files);
    });

    gridSearch.addEventListener('input', (e) => {
        const term = e.target.value.toLowerCase();
        renderGrid(Array.from(filesInput.files).filter(f => f.name.toLowerCase().includes(term)));
    });
}

function renderGrid(files) {
    pdfGrid.innerHTML = '';
    files.forEach(file => {
        const card = document.createElement('div');
        card.className = 'pdf-card';
        card.innerHTML = `<div class="pdf-card-icon">📄</div><div class="pdf-card-title" title="${file.name}">${file.name}</div>`;
        card.addEventListener('click', () => {
            modalViewer.data = URL.createObjectURL(file);
            pdfModal.style.display = 'flex';
        });
        pdfGrid.appendChild(card);
    });
}

if(document.getElementById('closeModal')) {
    document.getElementById('closeModal').addEventListener('click', () => {
        pdfModal.style.display = 'none';
        modalViewer.data = '';
    });
}

// --- Main Extraction Logic ---
const uploadForm = document.getElementById('uploadForm');
if (uploadForm) {
    uploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        try {
            const btn = document.getElementById('extractBtn');
            const progContainer = document.getElementById('progressContainer');
            const progFill = document.getElementById('progressFill');
            const progStats = document.getElementById('progressStats');
            const progTimer = document.getElementById('progressTimer');
            const finalRunStats = document.getElementById('finalRunStats');
            
            const files = Array.from(document.getElementById('pdfFiles').files);
            if (files.length === 0) {
                alert("Please select at least one PDF file.");
                return;
            }

            // Gather Dynamic Headers
            const customColKeys = Object.keys(getTableData('customTable'));
            let baseHeaders = `<th>File Name</th><th>Page #</th><th>Supplier</th><th>Inv #</th><th>Material</th><th>Description</th><th>Qty</th><th>UOM</th><th>Price</th><th>Line Total</th><th>Origin</th><th>Dest</th>`;
            customColKeys.forEach(col => baseHeaders += `<th style="color: #3498db;">${col}</th>`);
            document.getElementById('detailsHeaderRow').innerHTML = baseHeaders;

            // Reset UI
            btn.disabled = true;
            progContainer.style.display = 'block';
            finalRunStats.style.display = 'none';
            document.getElementById('resultsContainer').style.display = 'none';
            document.querySelector('#detailsTable tbody').innerHTML = '';
            document.querySelector('#summaryTable tbody').innerHTML = '';
            progFill.style.width = '0%';
            
            const total = files.length;
            let processed = 0;
            let totalPagesExtracted = 0; // Track Total Pages
            const batchId = "BATCH_" + Date.now();
            
            // Start Global Timer
            const startTime = Date.now();
            const timerInterval = setInterval(() => { 
                progTimer.innerText = `Time: ${Math.floor((Date.now() - startTime)/1000)}s`; 
            }, 1000);

            // Sequential processing
            for (let i = 0; i < total; i++) {
                progStats.innerText = `Extracting ${i + 1} of ${total}: ${files[i].name}...`;
                
                const formData = new FormData();
                formData.append('file', files[i]);
                formData.append('batch_id', batchId);
                formData.append('max_pages', document.getElementById('configMaxPages').value || 15);
                formData.append('dpi', document.getElementById('configDPI').value || 300);
                formData.append('aliases', JSON.stringify(getTableData('aliasTable')));
                formData.append('custom_fields', JSON.stringify(getTableData('customTable')));

                try {
                    const result = await fetchAPI('/api/extract-single', { method: 'POST', body: formData });
                    
                    if (result.status === 'success' && result.data) {
                        
                        // Append Summary & track pages
                        result.data.summary.forEach(row => {
                            if (row["Total Pages"]) totalPagesExtracted += row["Total Pages"]; // Accumulate pages
                            
                            const tr = document.createElement('tr');
                            tr.innerHTML = `<td>${row["File Name"]}</td><td>${row["Vendor Name"]}</td><td>${row["Invoice #"] || 'N/A'}</td><td>${row["Variance"]}</td><td>${row["Proc Time"]}</td><td class="${row["Status"].includes('FAIL') ? 'status-fail' : 'status-pass'}">${row["Status"]}</td>`;
                            document.querySelector('#summaryTable tbody').appendChild(tr);
                        });
                        
                        // Append Details
                        result.data.details.forEach(row => {
                            let customCells = "";
                            customColKeys.forEach(col => customCells += `<td>${row[col] || '-'}</td>`);
                            const tr = document.createElement('tr');
                            tr.innerHTML = `<td>${row["File Name"]}</td><td>${row["Page #"]||'-'}</td><td>${row["Original Supplier"]||'-'}</td><td>${row["Invoice Number"]||'-'}</td><td>${row["Material"]||'-'}</td><td>${row["Description"]}</td><td>${row["Qty"]||'-'}</td><td>${row["UOM"]||'-'}</td><td>${row["Price"]||'-'}</td><td>${row["Line Total"]||'-'}</td><td>${row["Origin"]||'-'}</td><td>${row["Dest"]||'-'}</td>${customCells}`;
                            document.querySelector('#detailsTable tbody').appendChild(tr);
                        });
                    }
                } catch (err) { 
                    console.error("API Call failed:", err); 
                }

                processed++;
                progFill.style.width = `${(processed / total) * 100}%`;
            }

            // End processing
            clearInterval(timerInterval);
            const totalTimeSeconds = Math.floor((Date.now() - startTime) / 1000);

            progStats.innerText = "Generating Excel File...";
            
            // Trigger Excel Save
            const excelForm = new FormData();
            excelForm.append('batch_id', batchId);
            excelForm.append('custom_fields', JSON.stringify(customColKeys));
            await fetch('/api/generate-excel', { method: 'POST', body: excelForm });
            document.getElementById('downloadExcelBtn').href = `/api/download-excel/${batchId}`;

            // Show Final Stats Card
            progStats.innerText = "✅ Complete!";
            finalRunStats.innerHTML = `📊 Batch Complete! &nbsp; | &nbsp; Total PDFs: ${total} &nbsp; | &nbsp; Total Pages: ${totalPagesExtracted} &nbsp; | &nbsp; Total Time: ${totalTimeSeconds}s`;
            finalRunStats.style.display = 'block';
            
            // Show Tables
            document.getElementById('resultsContainer').style.display = 'block';
            btn.disabled = false;

        } catch (fatalError) {
            alert("A fatal error occurred preventing extraction:\n" + fatalError.message);
            console.error(fatalError);
            if (document.getElementById('extractBtn')) {
                document.getElementById('extractBtn').disabled = false;
            }
        }
    });
}
