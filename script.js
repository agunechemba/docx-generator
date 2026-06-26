// Wrap everything in a window load event to ensure CDNs finish loading first
window.addEventListener('load', () => {
    
    // Run an immediate diagnostics check in the browser console
    console.log("XLSX Loaded:", !!window.XLSX);
    console.log("PizZip Loaded:", !!(window.PizZip || window.pizzip));
    console.log("Docxtemplater Loaded:", !!(window.Docxtemplater || window.docxtemplater));
    console.log("FileSaver Loaded:", !!window.saveAs);

    let templateBuffer = null;
    let templateName = 'document';
    let csvData = [];

    const templateInput = document.getElementById('templateUpload');
    const dataInput = document.getElementById('dataUpload');
    const templateStatus = document.getElementById('templateStatus');
    const dataStatus = document.getElementById('dataStatus');
    const generateBtn = document.getElementById('generateBtn');
    const statusDiv = document.getElementById('status');
    const templateDrop = document.getElementById('templateDrop');
    const dataDrop = document.getElementById('dataDrop');

    // Enable inputs immediately if HTML is ready
    if(templateInput) templateInput.disabled = false;
    if(dataInput) dataInput.disabled = false;

    templateInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        templateBuffer = await file.arrayBuffer();
        templateName = file.name.replace('.docx', '');
        templateStatus.textContent = `✅ Loaded: ${file.name}`;
        templateDrop.classList.add('has-file');
        checkReadyState();
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

            if (csvData.length === 0) throw new Error('File is empty.');
            dataStatus.textContent = `✅ Loaded: ${csvData.length} rows`;
            dataDrop.classList.add('has-file');
            checkReadyState();
        } catch (err) {
            dataStatus.textContent = `❌ Error: ${err.message}`;
        }
    });

    function checkReadyState() {
        if (templateBuffer && csvData.length > 0) {
            generateBtn.disabled = false;
            generateBtn.textContent = '🚀 Generate Documents';
        } else {
            generateBtn.disabled = true;
            generateBtn.textContent = '🚀 Upload Files to Start';
        }
    }

    generateBtn.addEventListener('click', () => {
        const Engine = window.Docxtemplater || window.docxtemplater;
        const ZipEngine = window.PizZip || window.pizzip;
        const Saver = window.saveAs;

        // Diagnostic alert detailing exactly what is missing
        if (!Engine || !ZipEngine || !Saver || !window.XLSX) {
            let missing = [];
            if (!window.XLSX) missing.push("Excel Parser (XLSX)");
            if (!ZipEngine) missing.push("Zip Engine (PizZip)");
            if (!Engine) missing.push("Document Template Engine");
            if (!Saver) missing.push("File Saver System");
            
            statusDiv.style.display = 'block';
            statusDiv.className = 'status error';
            statusDiv.innerHTML = `❌ <strong>Engine tracking fault.</strong><br>The browser blocked or failed to load: ${missing.join(', ')}.<br>Please disable ad-blockers or refresh with Ctrl+F5.`;
            return;
        }

        statusDiv.style.display = 'block';
        statusDiv.className = 'status';
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
                Saver(finalZip, `${templateName}_archive.zip`);
                
                statusDiv.textContent = '✅ Success! ZIP archive downloaded.';
            } catch (error) {
                statusDiv.className = 'status error';
                statusDiv.textContent = `❌ Error: ${error.message}`;
            } finally {
                generateBtn.disabled = false;
            }
        }, 100);
    });

    // Drag & Drop
    [templateDrop, dataDrop].forEach(section => {
        if(!section) return;
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
});