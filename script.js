// ============================================
// IMMUTABLE ES-MODULE RUNTIME IMPORTS
// ============================================
import PizZip from 'https://esm.sh/pizzip@3.1.4';
import Docxtemplater from 'https://esm.sh/docxtemplater@3.42.0';
import * as XLSX from 'https://esm.sh/xlsx@0.18.5';
import { saveAs } from 'https://esm.sh/file-saver@2.0.5';

// ============================================
// CORE DATA STATE
// ============================================
let templateBuffer = null;
let templateName = 'document';
let csvData = [];
let downloadMode = 'merged';

// ============================================
// ELEMENT MAPS
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
const clearStorageBtn = document.getElementById('clearStorageBtn');
const templateDrop = document.getElementById('templateDrop');
const dataDrop = document.getElementById('dataDrop');
const overlay = document.getElementById('appInitOverlay');
const optionCards = document.querySelectorAll('.option-card');

// ============================================
// BASE64 TRANSFORMS FOR LOCAL STORAGE
// ============================================
function bufferToBase64(buf) {
    let binary = '';
    const bytes = new Uint8Array(buf);
    for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

function base64ToBuffer(b64) {
    const binStr = atob(b64);
    const bytes = new Uint8Array(binStr.length);
    for (let i = 0; i < binStr.length; i++) {
        bytes[i] = binStr.charCodeAt(i);
    }
    return bytes.buffer;
}

// ============================================
// APPLICATION LIFECYCLE INITIALIZER
// ============================================
function initApplication() {
    try {
        // Hide initialization guard screen
        overlay.classList.remove('show');
        overlay.style.display = 'none';

        // Unlock input interfaces
        templateInput.disabled = false;
        dataInput.disabled = false;
        templateStatus.textContent = 'No file selected';
        dataStatus.textContent = 'No file selected';

        // Read LocalStorage Persistence Engine
        loadSavedState();
    } catch (err) {
        overlay.className = "status error show";
        overlay.textContent = `🚨 Framework State Error: ${err.message}`;
    }
}

function loadSavedState() {
    const savedMode = localStorage.getItem('docGen_downloadMode');
    if (savedMode) {
        downloadMode = savedMode;
        optionCards.forEach(card => {
            const radio = card.querySelector('input[type="radio"]');
            const isTarget = radio.value === downloadMode;
            card.classList.toggle('selected', isTarget);
            if (isTarget) radio.checked = true;
        });
    }

    const savedTplName = localStorage.getItem('docGen_templateName');
    const savedTplB64 = localStorage.getItem('docGen_templateBuffer');
    if (savedTplName && savedTplB64) {
        templateName = savedTplName;
        templateBuffer = base64ToBuffer(savedTplB64);
        templateStatus.textContent = `💾 Recovered: ${templateName}.docx`;
        templateDrop.classList.add('has-file');
        clearStorageBtn.style.display = 'block';
    }

    const savedCsv = localStorage.getItem('docGen_csvData');
    const savedCsvName = localStorage.getItem('docGen_dataFilename') || 'cached_data';
    if (savedCsv) {
        csvData = JSON.parse(savedCsv);
        dataStatus.textContent = `💾 Recovered: ${csvData.length} records from ${savedCsvName}`;
        dataDrop.classList.add('has-file');
        clearStorageBtn.style.display = 'block';
    }
    updateUI();
}

optionCards.forEach(card => {
    card.addEventListener('click', function() {
        optionCards.forEach(c => c.classList.remove('selected'));
        this.classList.add('selected');
        const radio = this.querySelector('input[type="radio"]');
        radio.checked = true;
        downloadMode = radio.value;
        localStorage.setItem('docGen_downloadMode', downloadMode);
        updateUI();
    });
});

clearStorageBtn.addEventListener('click', () => {
    localStorage.clear();
    location.reload();
});

// ============================================
// EVENT & DATA PIPELINE MANAGEMENT
// ============================================
templateInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    templateBuffer = await file.arrayBuffer();
    templateName = file.name.replace('.docx', '');
    templateStatus.textContent = `✅ ${file.name}`;
    templateDrop.classList.add('has-file');

    try {
        localStorage.setItem('docGen_templateBuffer', bufferToBase64(templateBuffer));
        localStorage.setItem('docGen_templateName', templateName);
        clearStorageBtn.style.display = 'block';
    } catch(err) {
        console.warn("Template tracking skipped cache storage limits.");
    }
    updateUI();
});

dataInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
        const buffer = await file.arrayBuffer();
        if (file.name.endsWith('.csv')) {
            const text = new TextDecoder().decode(buffer);
            const workbook = XLSX.read(text, { type: 'string' });
            csvData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        } else {
            const data = new Uint8Array(buffer);
            const workbook = XLSX.read(data, { type: 'array' });
            csvData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        }

        if (csvData.length === 0) throw new Error('Data file is empty.');

        dataStatus.textContent = `✅ ${csvData.length} records loaded`;
        dataDrop.classList.add('has-file');

        try {
            localStorage.setItem('docGen_csvData', JSON.stringify(csvData));
            localStorage.setItem('docGen_dataFilename', file.name);
            clearStorageBtn.style.display = 'block';
        } catch(err) {
            console.warn("Dataset exceeded quota memory constraints.");
        }
        updateUI();
    } catch (err) {
        dataStatus.textContent = `❌ Error: ${err.message}`;
        dataStatus.className = 'file-status error';
    }
});

function updateUI() {
    const ready = templateBuffer !== null && csvData.length > 0;
    generateBtn.disabled = !ready;
    generateBtn.textContent = ready ? '🚀 Generate Documents' : '🚀 Upload Files to Start';
}

// Drag & Drop
[templateDrop, dataDrop].forEach(section => {
    section.addEventListener('dragover', (e) => { e.preventDefault(); section.style.background = '#f0f4ff'; });
    section.addEventListener('dragleave', () => { section.style.background = section.classList.contains('has-file') ? '#f0fff4' : '#f7fafc'; });
    section.addEventListener('drop', (e) => {
        e.preventDefault();
        section.style.background = '#f7fafc';
        const input = section.querySelector('input[type="file"]');
        if (e.dataTransfer.files.length) {
            input.files = e.dataTransfer.files;
            input.dispatchEvent(new Event('change'));
        }
    });
});

// ============================================
// PROCESSING INTERFACES
// ============================================
function generateDocuments() {
    generateBtn.disabled = true;
    showStatus('Compiling files...', 'loading');
    progressSection.classList.add('show');
    progressFill.style.width = '0%';

    setTimeout(() => {
        try {
            const total = csvData.length;
            if (downloadMode === 'merged') {
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
                    updateProgress(i + 1, total);
                }

                const fullXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body${bodyAttributes}>${combinedContent}</w:body></w:document>`;
                const finalZip = new PizZip(templateBuffer);
                finalZip.file('word/document.xml', fullXml);
                const outBlob = finalZip.generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                saveAs(outBlob, `${templateName}_merged.docx`);
            } else {
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
                    updateProgress(i + 1, total);
                }
                const zipContent = exportZip.generate({ type: 'blob' });
                saveAs(zipContent, `${templateName}_archive.zip`);
            }
            showStatus('✅ Documents created successfully!', 'success');
        } catch (error) {
            showStatus(`❌ Processing Error: ${error.message}`, 'error');
        } finally {
            generateBtn.disabled = false;
            updateUI();
        }
    }, 100);
}

function updateProgress(current, total) {
    const pct = Math.round((current / total) * 100);
    progressFill.style.width = `${pct}%`;
    progressPercent.textContent = `${pct}%`;
    progressLabel.textContent = `Processing document ${current} of ${total}`;
}

function showStatus(msg, type) {
    statusDiv.textContent = msg;
    statusDiv.className = `status show ${type}`;
}

// Run immediately
initApplication();
generateBtn.addEventListener('click', generateDocuments);