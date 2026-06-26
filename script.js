let templateBuffer = null;
let templateName = 'document';
let csvData = [];

const templateInput = document.getElementById('templateUpload');
const dataInput = document.getElementById('dataUpload');
const templateStatus = document.getElementById('templateStatus');
const dataStatus = document.getElementById('dataStatus');
const generateBtn = document.getElementById('generateBtn');
const statusDiv = document.getElementById('status');

// Handle Template uploads
templateInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    templateBuffer = await file.arrayBuffer();
    templateName = file.name.replace('.docx', '');
    templateStatus.textContent = `✅ Loaded: ${file.name}`;
    checkReadyState();
});

// Handle Data spreadsheet uploads
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

        if (csvData.length === 0) throw new Error('File is empty.');
        dataStatus.textContent = `✅ Loaded: ${csvData.length} records`;
        checkReadyState();
    } catch (err) {
        dataStatus.textContent = `❌ Error: ${err.message}`;
    }
});

function checkReadyState() {
    generateBtn.disabled = !(templateBuffer && csvData.length > 0);
}

// Processing Core Engine Loop
generateBtn.addEventListener('click', () => {
    // Fail-safe namespace selector
    const Engine = window.Docxtemplater || window.docxtemplater;
    const ZipEngine = window.PizZip || window.pizzip;

    if (!Engine || !ZipEngine) {
        statusDiv.style.display = 'block';
        statusDiv.textContent = '❌ Core engines failed to load. Please verify your connection.';
        return;
    }

    statusDiv.style.display = 'block';
    statusDiv.textContent = '📦 Compiling separate documents into ZIP archive...';
    generateBtn.disabled = true;

    setTimeout(() => {
        try {
            const exportZip = new ZipEngine();

            csvData.forEach((row, i) => {
                const zip = new ZipEngine(templateBuffer);
                const doc = new Engine(zip, { paragraphLoop: true, linebreaks: true });
                
                doc.render(row);

                const outBlob = doc.getZip().generate({ 
                    type: 'blob', 
                    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
                });
                
                const rawName = row.name || row.Name || row.NAME || `document_${i + 1}`;
                const safeName = String(rawName).replace(/[^a-zA-Z0-9]/g, '_').substring(0, 50);
                exportZip.file(`${safeName}.docx`, outBlob);
            });

            const finalZip = exportZip.generate({ type: 'blob' });
            saveAs(finalZip, `${templateName}_archive.zip`);
            
            statusDiv.textContent = '✅ Success! ZIP archive downloaded.';
        } catch (error) {
            statusDiv.textContent = `❌ Error: ${error.message}`;
        } finally {
            generateBtn.disabled = false;
        }
    }, 100);
});