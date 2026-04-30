async function fetchAPI(endpoint, options = {}) {
    const response = await fetch(endpoint, options);
    if (!response.ok) throw new Error(`API Error: ${response.status}`);
    return await response.json();
}

// 1. Extraction UI & Document Viewer Logic
const uploadForm = document.getElementById('uploadForm');
if (uploadForm) {
    const filesInput = document.getElementById('pdfFiles');
    const docSelect = document.getElementById('docSelect');
    const pdfViewer = document.getElementById('pdfViewer');
    const pdfContainer = document.getElementById('pdfContainer');
    const pdfFallback = document.getElementById('pdfFallback');

    // Populate Viewer Dropdown when files are selected
    filesInput.addEventListener('change', () => {
        docSelect.innerHTML = '<option value="">Select a document to view...</option>';
        docSelect.style.display = filesInput.files.length > 0 ? 'block' : 'none';
        pdfContainer.style.display = 'none';

        Array.from(filesInput.files).forEach((file, index) => {
            const option = document.createElement('option');
            option.value = index;
            option.textContent = file.name;
            docSelect.appendChild(option);
        });
    });

    // Render the selected PDF
    docSelect.addEventListener('change', (e) => {
        const fileIndex = e.target.value;
        if (fileIndex !== "") {
            const file = filesInput.files[fileIndex];
            const fileURL = URL.createObjectURL(file);
            pdfViewer.data = fileURL;
            pdfFallback.href = fileURL;
            pdfContainer.style.display = 'block';
        } else {
            pdfContainer.style.display = 'none';
        }
    });

    // Handle Form Submission & Table Population
    uploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        const loader = document.getElementById('loader');
        const btn = document.getElementById('extractBtn');
        const resultsContainer = document.getElementById('resultsContainer');
        const detailsBody = document.querySelector('#detailsTable tbody');
        const summaryBody = document.querySelector('#summaryTable tbody');
        const excelBtn = document.getElementById('downloadExcelBtn');
        
        if (!filesInput.files.length) return;

        const formData = new FormData();
        for (const file of filesInput.files) formData.append('files', file);

        // Reset UI
        btn.disabled = true; 
        loader.style.display = 'block'; 
        detailsBody.innerHTML = ''; 
        summaryBody.innerHTML = ''; 
        resultsContainer.style.display = 'none';

        try {
            const result = await fetchAPI('/api/extract', { method: 'POST', body: formData });
            
            if (result.status === 'success') {
                const data = result.data;
                
                // Link the Excel download to the exact batch ID generated on the backend
                excelBtn.href = `/api/download-excel/${data.batch_id}`;

                // Populate Line Items Table
                data.details.forEach(row => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td>${row["File Name"]}</td>
                        <td>${row["Page #"] || '-'}</td>
                        <td>${row["Description"]}</td>
                        <td>${row["Quantity"] || '-'}</td>
                        <td>${row["UOM"] || '-'}</td>
                        <td>${row["Unit Price"] || '-'}</td>
                        <td>${row["Line Total"] || '-'}</td>
                    `;
                    detailsBody.appendChild(tr);
                });

                // Populate QC Summary Table
                data.summary.forEach(row => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td>${row["File Name"]}</td>
                        <td>${row["Vendor Name"]}</td>
                        <td>${row["Invoice #"] || 'N/A'}</td>
                        <td>${row["Variance"]}</td>
                        <td>${row["Proc Time"]}</td>
                        <td class="${row["Status"].includes('FAIL') ? 'status-fail' : 'status-pass'}">${row["Status"]}</td>
                    `;
                    summaryBody.appendChild(tr);
                });

                // Show all results
                resultsContainer.style.display = 'block';
            }
        } catch (error) { 
            alert("Extraction Failed: " + error.message); 
        } finally { 
            btn.disabled = false; 
            loader.style.display = 'none'; 
        }
    });
}

// 2. Analytics UI Logic (Plotly Integration)
if (document.getElementById('chart-container')) {
    fetchAPI('/api/audit-data').then(data => {
        const vendorSpend = {};
        data.forEach(row => {
            if(row.vendor_name && row.vendor_name !== 'ERROR' && row.extracted_total) {
                vendorSpend[row.vendor_name] = (vendorSpend[row.vendor_name] || 0) + row.extracted_total;
            }
        });
        const trace = { x: Object.keys(vendorSpend), y: Object.values(vendorSpend), type: 'bar', marker: { color: '#3498db' }};
        Plotly.newPlot('chart-container', [trace], { title: 'Total Spend by Vendor ($)' });
    });
}
