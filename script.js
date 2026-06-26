// ============================================
// STATE MANAGEMENT
// ============================================
let templateBuffer = null;
let templateName = 'document';
let csvData = [];
let downloadMode = 'merged';

// ============================================
// DOM ELEMENTS
// ============================================
const templateInput = document.getElementById('templateUpload');
const dataInput = document.getElementById('dataUpload');
const templateStatus = document.getElementById('templateStatus');
const dataStatus = document.getElementById('dataStatus');
const generateBtn = document.getElementById('generateBtn');
const statusDiv = document.getElementById('status');
const progressSection = document.getElementById('progressSection');
const progressFill = document.getElementById('progressFill');
const progressLabel = document.getElementById('progressLabel');
const progressPercent = document.getElementById('progressPercent');
const statsDiv = document.getElementById('stats');
const recordCount = document.getElementById('recordCount');
const docCount = document.getElementById('docCount');
const outputType = document.getElementById('outputType');
const clearStorageBtn = document.getElementById('clearStorageBtn');

const templateDrop = document.getElementById('templateDrop');
const dataDrop = document.getElementById('dataDrop');

// Setup download selection options card triggers
const optionCards = document.querySelectorAll('.option-card');
optionCards.forEach(card => {
    card.addEventListener('click', function() {
        optionCards.forEach(c => c.classList.remove('selected'));
        this.classList.add('selected');
        const radio = this.querySelector('input[type="radio"]');
        radio.checked = true;
        downloadMode = radio.value;
        outputType.textContent = downloadMode === 'merged' ? 'Merged' : 'ZIP';
        localStorage.setItem('docGen_downloadMode', downloadMode);
        updateUI();
    });
});

// ============================================
// UTILITY FUNCTIONS FOR BINARY STORAGE
// ============================================
function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
}

function base64ToArrayBuffer(base64) {
    const binary_string = window.atob(base64);
    const len = binary_string.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binary_string.charCodeAt(i);
    }
    return bytes.buffer;
}

// ============================================
// LOCAL STORAGE ENGINE
// ============================================
function loadSavedState() {
    try {
        const savedMode = localStorage.getItem('docGen_downloadMode');
        if (savedMode) {
            downloadMode = savedMode;
            optionCards.forEach(card => {
                const radio = card.querySelector('input[type="radio"]');
                if (radio.value === downloadMode) {
                    optionCards.forEach(c => c.classList.remove('selected'));
                    card.classList.add('selected');
                    radio.checked = true;
                }
            });
            outputType.textContent = downloadMode === 'merged' ? 'Merged' : 'ZIP';
        }

        const savedTemplateName = localStorage.getItem('docGen_templateName');
        const savedTemplateBase64 = localStorage.getItem('docGen_templateBuffer');
        if (savedTemplateName && savedTemplateBase64) {
            templateName = savedTemplateName;
            templateBuffer = base64ToArrayBuffer(savedTemplateBase64);
            templateStatus.textContent = `💾 Recovered: ${templateName}.docx`;
            templateDrop.classList.add('has-file');
            clearStorageBtn.style.display = 'block';
        }

        const savedCsvData = localStorage.getItem('docGen_csvData');
        const savedDataFilename = localStorage.getItem('docGen_dataFilename') || 'cached_data';
        if (savedCsvData) {
            csvData = JSON.parse(savedCsvData);
            dataStatus.textContent = `💾 Recovered: ${csvData.length} rows from ${savedDataFilename}`;
            dataDrop.classList.add('has-file');
            recordCount.textContent = csvData.length;
            docCount.textContent = csvData.length;
            statsDiv.classList.add('show');
            clearStorageBtn.style.display = 'block';
        }
        updateUI();
    } catch (e) {
        console.error("Failed to restore saved local session state:", e);
    }
}

clearStorageBtn.addEventListener('click', function() {
    localStorage.clear();
    location.reload();
});

// ============================================
// FILE HANDLERS
// ============================================
templateInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.endsWith('.docx')) {
        templateStatus.textContent = '❌ Please upload a .docx file';
        templateStatus.className = 'file-status error';
        return;
    }

    const reader = new FileReader();
    reader.onload = function(event) {
        templateBuffer = event.target.result;
        templateName = file.name.replace('.docx', '');
        templateStatus.textContent = `✅ ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        templateStatus.className = 'file-status';
        templateDrop.classList.add('has-file');
        
        try {
            const base64Str = arrayBufferToBase64(templateBuffer);
            localStorage.setItem('docGen_templateBuffer', base64Str);
            localStorage.setItem('docGen_templateName', templateName);
            clearStorageBtn.style.display = 'block';
        } catch(err) {
            console.warn("Template file is too large to fully cache in LocalStorage quota limits.");
        }
        updateUI();
    };
    reader.readAsArrayBuffer(file);
});

dataInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        try {
            let workbook;
            if (file.name.endsWith('.csv')) {
                const text = new TextDecoder().decode(event.target.result);
                const lines = text.split('\n').filter(line => line.trim());
                if (lines.length < 2) throw new Error('CSV must have headers and rows');
                const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
                csvData = lines.slice(1).map(line => {
                    const values = line.split(',').map(v => v.trim().replace(/^"|"$/g, ''));
                    const obj = {};
                    headers.forEach((h, i) => { obj[h] = values[i] || ''; });
                    return obj;
                });
            } else {
                const data = new Uint8Array(event.target.result);
                workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                csvData = XLSX.utils.sheet_to_json(firstSheet);
            }

            if (csvData.length === 0) throw new Error('No data rows found');

            dataStatus.textContent = `✅ ${csvData.length} records loaded from ${file.name}`;
            dataStatus.className = 'file-status';
            dataDrop.classList.add('has-file');
            recordCount.textContent = csvData.length;
            docCount.textContent = csvData.length;
            statsDiv.classList.add('show');

            try {
                localStorage.setItem('docGen_csvData', JSON.stringify(csvData));
                localStorage.setItem('docGen_dataFilename', file.name);
                clearStorageBtn.style.display = 'block';
            } catch(err) {
                console.warn("Dataset exceeds maximum browser LocalStorage quota capacity.");
            }
            updateUI();
        } catch (error) {
            dataStatus.textContent = `❌ Error: ${error.message}`;
            dataStatus.className = 'file-status error';
        }
    };
    reader.readAsArrayBuffer(file);
});

// Drag and Drop listeners
[templateDrop, dataDrop].forEach(section => {
    section.addEventListener('dragover', function(e) {
        e.preventDefault();
        this.style.borderColor = '#667eea';
        this.style.background = '#f0f4ff';
    });
    section.addEventListener('dragleave', function(e) {
        e.preventDefault();
        this.style.borderColor = this.classList.contains('has-file') ? '#48bb78' : '#e2e8f0';
        this.style.background = this.classList.contains('has-file') ? '#f0fff4' : '#f7fafc';
    });
    section.addEventListener('drop', function(e) {
        e.preventDefault();
        this.style.borderColor = '#e2e8f0';
        this.style.background = '#f7fafc';
        const input = this.querySelector('input[type="file"]');
        if (e.dataTransfer.files.length) {
            input.files = e.dataTransfer.files;
            input.dispatchEvent(new Event('change'));
        }
    });
});

function updateUI() {
    const ready = templateBuffer !== null && csvData.length > 0;
    generateBtn.disabled = !ready;
    if (ready) {
        generateBtn.textContent = `🚀 Generate ${downloadMode === 'merged' ? 'Merged Document' : 'ZIP Archive'}`;
    } else {
        generateBtn.textContent = '🚀 Generate Documents';
    }
}

// ============================================
// CORE BATCH GENERATION WORKFLOW
// ============================================
function generateDocuments() {
    if (typeof PizZip === 'undefined' || typeof Docxtemplater === 'undefined') {
        showStatus('❌ Error: Framework parsing runtimes failed loading dependencies.', 'error');
        return;
    }

    generateBtn.disabled = true;
    showStatus('Processing generation tasks...', 'loading');
    progressSection.classList.add('show');
    progressFill.style.width = '0%';
    progressLabel.textContent = 'Initializing engine...';
    progressPercent.textContent = '0%';

    setTimeout(() => {
        try {
            const total = csvData.length;
            if (downloadMode === 'merged') {
                generateMergedFallback(total);
            } else {
                generateZIP(total);
            }
            showStatus(`✅ Generated batch files successfully!`, 'success');
            progressLabel.textContent = 'Complete! ✅';
            progressPercent.textContent = '100%';
            progressFill.style.width = '100%';
        } catch (error) {
            console.error(error);
            showStatus(`❌ Runtime Processing Error: ${error.message}`, 'error');
        } finally {
            generateBtn.disabled = false;
            updateUI();
        }
    }, 100);
}

function generateMergedFallback(total) {
    let combinedContent = '';
    let bodyAttributes = '';

    for (let i = 0; i < total; i++) {
        const zip = new PizZip(templateBuffer);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
        doc.render(csvData[i]);
        const docXml = doc.getZip().file('word/document.xml').asText();

        if (i === 0) {
            const bodyMatch = docXml.match(/<w:body([^>]*)>([\s\S]*?)<\/w:body>/);
            bodyAttributes = bodyMatch ? bodyMatch[1] : '';
            combinedContent = bodyMatch ? bodyMatch[2] : docXml;
        } else {
            const bodyMatch = docXml.match(/<w:body[^>]*>([\s\S]*?)<\/w:body>/);
            if (bodyMatch && bodyMatch[1]) {
                combinedContent += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + bodyMatch[1];
            }
        }
        const progress = ((i + 1) / total) * 100;
        progressFill.style.width = `${progress}%`;
        progressPercent.textContent = `${Math.round(progress)}%`;
        progressLabel.textContent = `Compiling ${i + 1}/${total}`;
    }

    const fullXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body${bodyAttributes}>${combinedContent}</w:body></w:document>`;
    const finalZip = new PizZip(templateBuffer);
    finalZip.file('word/document.xml', fullXml);
    const outBlob = finalZip.generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    saveAs(outBlob, `${templateName}_merged.docx`);
}

function generateZIP(total) {
    const exportZip = new PizZip();

    for (let i = 0; i < total; i++) {
        const row = csvData[i];
        const zip = new PizZip(templateBuffer);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
        doc.render(row);

        const outBlob = doc.getZip().generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        const rawName = row.name || row.Name || row.NAME || `document_${i + 1}`;
        const safeName = String(rawName).replace(/[^a-zA-Z0-9]/g, '_').substring(0, 50);
        exportZip.file(`${safeName}.docx`, outBlob);

        const progress = ((i + 1) / total) * 100;
        progressFill.style.width = `${progress}%`;
        progressPercent.textContent = `${Math.round(progress)}%`;
        progressLabel.textContent = `Archiving ${i + 1}/${total}`;
    }
    const zipContent = exportZip.generate({ type: 'blob' });
    saveAs(zipContent, `${templateName}_archive.zip`);
}

function showStatus(message, type = 'info') {
    statusDiv.textContent = message;
    statusDiv.className = `status show ${type}`;
}

document.addEventListener('DOMContentLoaded', loadSavedState);
generateBtn.addEventListener('click', generateDocuments);