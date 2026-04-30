// Global API Fetcher with Basic Error Handling
async function fetchAPI(endpoint, options = {}) {
    const response = await fetch(endpoint, options);
    if (!response.ok) throw new Error(`API Error: ${response.status}`);
    return await response.json();
}

// 1. Extraction UI Logic
const uploadForm = document.getElementById('uploadForm');
if (uploadForm) {
    uploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        const filesInput = document.getElementById('pdfFiles');
        const loader = document.getElementById('loader');
        const btn = document.getElementById('extractBtn');
        const tbody = document.querySelector('#resultsTable tbody');
        const resultsCard = document.getElementById('resultsCard');
        
        if (!filesInput.files.length) return;

        const formData = new FormData();
        for (const file of filesInput.files) formData.append('files', file);

        btn.disabled = true; loader.style.display = 'block'; tbody.innerHTML = ''; resultsCard.style.display = 'none';

        try {
            const result = await fetchAPI('/api/extract', { method: 'POST', body: formData });
            if (result.status === 'success') {
                result.data.forEach(row => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `<td>${row.file}</td><td>${row.vendor}</td>
                                    <td class="${row.status.includes('FAIL') ? 'status-fail' : 'status-pass'}">${row.status}</td>`;
                    tbody.appendChild(tr);
                });
                resultsCard.style.display = 'block';
            }
        } catch (error) { alert(error.message); } 
        finally { btn.disabled = false; loader.style.display = 'none'; filesInput.value = ''; }
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

// 3. Batch QA Data Loader
if (document.getElementById('qaTableBody')) {
    fetchAPI('/api/audit-data').then(data => {
        const tbody = document.getElementById('qaTableBody');
        // Get last 50 for QA
        data.slice(-50).reverse().forEach(row => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>${row.file_name || 'N/A'}</td><td>${row.invoice_number || 'N/A'}</td>
                            <td>${row.vendor_name}</td><td>$${(row.variance || 0).toFixed(2)}</td>
                            <td class="${row.status.includes('FAIL') ? 'status-fail' : 'status-pass'}">${row.status}</td>`;
            tbody.appendChild(tr);
        });
    });
}
