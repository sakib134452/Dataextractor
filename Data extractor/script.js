document.addEventListener('DOMContentLoaded', () => {
    const dropArea = document.getElementById('drop-area');
    const fileInput = document.getElementById('excelFile');
    const fileNameSpan = document.getElementById('fileName');
    const columnSelectSection = document.getElementById('column-select-section');
    const columnNameSelect = document.getElementById('columnName');
    const extractBtn = document.getElementById('extractBtn');
    const dataList = document.getElementById('dataList');
    const outputSection = document.getElementById('outputSection');
    const downloadTxtBtn = document.getElementById('downloadTxtBtn');
    const downloadPdfBtn = document.getElementById('downloadPdfBtn');
    
    let extractedData = [];
    let uploadedFile = null;
    let headers = [];

    // Reset UI state
    function resetUI() {
        columnSelectSection.classList.add('hidden');
        columnNameSelect.innerHTML = '<option value="">-- Select a column --</option>';
        extractBtn.disabled = true;
        outputSection.classList.add('hidden');
        dataList.innerHTML = '';
        extractedData = [];
    }

    // Handle file upload/drop
    function handleFile(file) {
        if (!file || !file.name.endsWith('.xlsx')) {
            alert('Please upload a valid Excel file (.xlsx).');
            return;
        }

        uploadedFile = file;
        fileNameSpan.textContent = file.name;
        resetUI();

        const reader = new FileReader();
        reader.onload = function(event) {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                if (jsonData.length > 0) {
                    headers = jsonData[0];
                    columnNameSelect.innerHTML = '<option value="">-- Select a column --</option>';
                    headers.forEach(header => {
                        const option = document.createElement('option');
                        option.value = header;
                        option.textContent = header;
                        columnNameSelect.appendChild(option);
                    });
                    columnSelectSection.classList.remove('hidden');
                } else {
                    alert('No data found in the Excel file.');
                }
            } catch (error) {
                alert('An error occurred while processing the Excel file.');
                console.error(error);
            }
        };
        reader.readAsArrayBuffer(file);
    }
    
    // File input change event
    fileInput.addEventListener('change', function(e) {
        handleFile(e.target.files[0]);
    });

    // Handle drag and drop events
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, () => dropArea.classList.add('highlight'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, () => dropArea.classList.remove('highlight'), false);
    });

    dropArea.addEventListener('drop', function(e) {
        let dt = e.dataTransfer;
        let files = dt.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    }, false);

    dropArea.addEventListener('click', () => fileInput.click());

    // Enable/disable extract button based on selection
    columnNameSelect.addEventListener('change', function() {
        extractBtn.disabled = !this.value;
    });

    // Main extraction logic
    extractBtn.addEventListener('click', function() {
        const selectedColumnName = columnNameSelect.value.trim();
        dataList.innerHTML = '';
        extractedData = [];

        if (!uploadedFile || !selectedColumnName) {
            alert('Please select a file and a column.');
            return;
        }

        const reader = new FileReader();
        reader.onload = function(event) {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                const columnIndex = headers.indexOf(selectedColumnName);

                for (let i = 1; i < jsonData.length; i++) {
                    const row = jsonData[i];
                    const cellData = row[columnIndex];
                    if (cellData !== undefined && cellData !== null) {
                        const listItem = document.createElement('li');
                        listItem.textContent = cellData;
                        dataList.appendChild(listItem);
                        extractedData.push(String(cellData));
                    }
                }

                if (dataList.childElementCount === 0) {
                    dataList.innerHTML = '<li>No data found in this column.</li>';
                }
                outputSection.classList.remove('hidden');
            } catch (error) {
                alert('An error occurred while processing the Excel file.');
                console.error(error);
            }
        };
        reader.readAsArrayBuffer(uploadedFile);
    });

    // Download logic remains the same
    downloadTxtBtn.addEventListener('click', function() {
        if (extractedData.length === 0) {
            alert('No data to download.');
            return;
        }
        const textContent = extractedData.join('\n');
        const blob = new Blob([textContent], { type: 'text/plain' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'extracted_data.txt';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    });

    downloadPdfBtn.addEventListener('click', function() {
        if (extractedData.length === 0) {
            alert('No data to download.');
            return;
        }
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        const pageWidth = doc.internal.pageSize.getWidth();
        const margin = 15;
        const contentWidth = pageWidth - 2 * margin;
        const lineHeight = 7;
        let y = margin + 10;

        doc.setFont('helvetica', 'bold');
        doc.setFontSize(18);
        doc.text('Extracted Data', margin, margin);
        
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(12);
        extractedData.forEach(item => {
            const lines = doc.splitTextToSize(item, contentWidth);
            doc.text(lines, margin, y);
            y += lines.length * lineHeight;
            if (y > doc.internal.pageSize.getHeight() - margin) {
                doc.addPage();
                y = margin;
            }
        });

        doc.save('extracted_data.pdf');
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
});