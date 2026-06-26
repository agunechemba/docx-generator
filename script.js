// Global State
let templateBuffer = null;
let templateName = 'document';
let csvData = [];
let downloadMode = 'merged';

// Element DOM Mapping
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
const templateDrop = document.getElementById('templateDrop');
const dataDrop = document.getElementById('dataDrop');
const optionCards = document.querySelectorAll('.option-card');

// Handle Mode Selector Buttons
optionCards.forEach(card => {
    card.addEventListener('click', function() {
        optionCards.forEach(c => c.classList.remove('selected'));
        this.classList.add('selected');
        const radio = this.querySelector('input[type="radio"]');
        radio.checked = true;
        downloadMode = radio.value;
        updateUI();
    });
});

// Template Upload Processing
templateInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    templateBuffer = await file.arrayBuffer();
    templateName = file.name.replace('.docx', '');
    templateStatus.textContent = `✅ Loaded: ${file.name}`;
    templateDrop.classList.add('has-file');
    updateUI();
});

// Spreadsheet/CSV Upload Processing
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

        if (csvData.length === 0) throw new Error('File contains no records.');

        dataStatus.textContent = `✅ Loaded: ${csvData.length} rows`;
        dataDrop.classList.add('has-file');
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

// Processing Logic Loop
function generateDocuments() {
    // Crucial safety fallback mapping to the correct window key name
    const DocEngine = window.Docxtemplater || window.docxtemplater;
    
    if (!DocEngine || !window.PizZip || !window.saveAs) {
        showStatus('❌ Critical dependency error. Please refresh your page.', 'error');
        return;
    }

    generateBtn.disabled = true;
    showStatus('Processing generation tasks...', 'loading');
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
                    const doc = new DocEngine(zip, { paragraphLoop: true, linebreaks: true });
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
                    const doc = new DocEngine(zip, { paragraphLoop: true, linebreaks: true });
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
            showStatus(`❌ Error: ${error.message}`, 'error');
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
    progressLabel.textContent = `Processing ${current} of ${total}`;
}

function showStatus(msg, type) {
    statusDiv.textContent = msg;
    statusDiv.className = `status show ${type}`;
}

// Drag & Drop Listeners
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

generateBtn.addEventListener('click', generateDocuments);