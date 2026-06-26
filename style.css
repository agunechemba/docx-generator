/**
 * Bulk Document Generator Engine
 * Purely client-side architecture optimized for GitHub Pages.
 */

(function () {
    'use strict';

    // --- Bulletproof Global Scope Dependecy Resolution ---
    const ExcelParser = window.XLSX;
    const ZipContainer = window.PizZip;
    const FileSaver = window.saveAs;
    const StandaloneZip = window.JSZip;
    
    // Fail-safe namespace resolution for Docxtemplater across CDN distributions
    const DocxEngine = window.Docxtemplater || window.docxtemplater;

    // Critical assertion layer to prevent silent deployment configuration breakages
    if (!ExcelParser || !ZipContainer || !DocxEngine || !FileSaver || !StandaloneZip) {
        console.error("Dependency Check Failure: Missing upstream global CDN libraries.");
        alert("System Error: Critical runtime libraries failed to initialize. Check browser console logs.");
        return;
    }

    // --- State Storage Application Context ---
    let templateArrayBuffer = null;
    let spreadsheetRecords = [];
    let discoveredKeys = [];

    // --- DOM Elements Cache ---
    const docxDropzone = document.getElementById('dropzone-docx');
    const docxInput = document.getElementById('input-docx');
    const docxLabel = document.getElementById('label-docx');
    const docxFileName = document.getElementById('file-name-docx');

    const xlsxDropzone = document.getElementById('dropzone-xlsx');
    const xlsxInput = document.getElementById('input-xlsx');
    const xlsxLabel = document.getElementById('label-xlsx');
    const xlsxFileName = document.getElementById('file-name-xlsx');

    const statusPanel = document.getElementById('status-panel');
    const metaRows = document.getElementById('meta-rows');
    const metaKeys = document.getElementById('meta-keys');
    const namingColumnInput = document.getElementById('naming-column');

    const btnMerge = document.getElementById('btn-merge');
    const btnZip = document.getElementById('btn-zip');

    const progressContainer = document.getElementById('progress-container');
    const progressStatus = document.getElementById('progress-status');
    const progressPercent = document.getElementById('progress-percent');
    const progressBar = document.getElementById('progress-bar');

    // --- Drag & Drop Decorator Operations ---
    function setupDragAndDrop(dropzone, inputElement, callback) {
        dropzone.addEventListener('click', () => inputElement.click());
        dropzone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropzone.classList.add('drop-zone--over');
        });
        ['dragleave', 'drop'].forEach(evt => {
            dropzone.addEventListener(evt, () => dropzone.classList.remove('drop-zone--over'));
        });
        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            if (e.dataTransfer.files.length) {
                inputElement.files = e.dataTransfer.files;
                callback(e.dataTransfer.files[0]);
            }
        });
        inputElement.addEventListener('change', (e) => {
            if (e.target.files.length) callback(e.target.files[0]);
        });
    }

    // Initialize drag/drop features
    setupDragAndDrop(docxDropzone, docxInput, handleDocxTemplateFile);
    setupDragAndDrop(xlsxDropzone, xlsxInput, handleSpreadsheetFile);

    // --- File Processing Implementations ---
    function handleDocxTemplateFile(file) {
        if (!file.name.endsWith('.docx')) {
            alert('Invalid configuration: Selected template must be a valid .docx file.');
            return;
        }
        const reader = new FileReader();
        reader.onload = function (e) {
            templateArrayBuffer = e.target.result;
            docxLabel.textContent = "Template Loaded Successfully";
            docxFileName.textContent = file.name;
            docxFileName.classList.remove('hidden');
            evaluateGenerationState();
        };
        reader.readAsArrayBuffer(file);
    }

    function handleSpreadsheetFile(file) {
        if (!file.name.endsWith('.xlsx')) {
            alert('Invalid configuration: Selected data matrix file must be a valid .xlsx spreadsheet.');
            return;
        }
        const reader = new FileReader();
        reader.onload = function (e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = ExcelParser.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Read and structurally output raw row matrices
                // defval: "" preserves structural integrity for empty table cells
                const sheetJson = ExcelParser.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

                if (sheetJson.length < 2) {
                    alert("Processing Error: Data sheet requires a primary key row index (header) and at least one entry record.");
                    return;
                }

                const headers = sheetJson[0].map(h => String(h).trim());
                discoveredKeys = headers.filter(h => h !== "");
                spreadsheetRecords = [];

                // Parse records mapped cleanly into key/value data objects
                for (let i = 1; i < sheetJson.length; i++) {
                    const row = sheetJson[i];
                    // Skip completely clean empty tracking lines
                    if (row.every(cell => cell === "")) continue; 
                    
                    let record = {};
                    headers.forEach((header, index) => {
                        if (header) {
                            record[header] = row[index] !== undefined ? row[index] : "";
                        }
                    });
                    spreadsheetRecords.push(record);
                }

                // UI feedback updates
                metaRows.textContent = spreadsheetRecords.length;
                metaKeys.textContent = discoveredKeys.join(', ');
                statusPanel.classList.remove('hidden');
                
                if(discoveredKeys.length > 0 && !namingColumnInput.value) {
                    namingColumnInput.value = discoveredKeys[0]; // Set default fallback key dynamically
                }

                xlsxLabel.textContent = "Data Matrix Importer Synced";
                xlsxFileName.textContent = file.name;
                xlsxFileName.classList.remove('hidden');
                evaluateGenerationState();

            } catch (err) {
                console.error(err);
                alert("Fatal Engine Parsing Exception caught reading Excel Workbook.");
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function evaluateGenerationState() {
        if (templateArrayBuffer && spreadsheetRecords.length > 0) {
            btnMerge.removeAttribute('disabled');
            btnZip.removeAttribute('disabled');
        }
    }

    // --- Async Processing Animation Framework ---
    function updateProgress(current, total, statusText) {
        const percent = Math.round((current / total) * 100);
        progressContainer.classList.remove('hidden');
        progressBar.style.width = `${percent}%`;
        progressPercent.textContent = `${percent}%`;
        progressStatus.textContent = statusText;
    }

    function resetProgress() {
        setTimeout(() => {
            progressContainer.classList.add('hidden');
            progressBar.style.width = `0%`;
        }, 2000);
    }

    // --- Core Architecture: Engine Generators ---
    
    // Process records in small asynchronous execution blocks to keep the browser UI ultra-responsive
    function batchProcess(records, processFn, finishFn) {
        let index = 0;
        function nextBatch() {
            const limit = Math.min(index + 5, records.length); // Process 5 records at a time
            while (index < limit) {
                processFn(records[index], index);
                index++;
            }
            if (index < records.length) {
                updateProgress(index, records.length, `Compiling batch structures...`);
                // Release main execution loop thread to draw UI animations
                setTimeout(nextBatch, 0); 
            } else {
                finishFn();
            }
        }
        nextBatch();
    }

    // Action A Strategy: Generate Independent Documents within a Target ZIP Archive
    btnZip.addEventListener('click', function() {
        if(!templateArrayBuffer || !spreadsheetRecords.length) return;
        
        const archive = new StandaloneZip();
        const namingKey = namingColumnInput.value.trim() || Object.keys(spreadsheetRecords[0])[0];

        batchProcess(spreadsheetRecords, (record, i) => {
            try {
                const zip = new ZipContainer(templateArrayBuffer);
                const doc = new DocxEngine(zip, { paragraphLoop: true, linebreaks: true });
                doc.render(record);
                const out = doc.getZip().generate({ type: "blob", compression: "DEFLATE" });
                
                let dynamicName = record[namingKey] ? String(record[namingKey]).replace(/[/\\?%*:|"<>\s]/g, '_') : `Record_${i + 1}`;
                archive.file(`${dynamicName}.docx`, out);
            } catch (ex) {
                console.error(`Error processing record dynamic context at line index ${i}:`, ex);
            }
        }, () => {
            updateProgress(95, 100, "Assembling Compressed ZIP Container File...");
            archive.generateAsync({ type: "blob" }).then(function(content) {
                updateProgress(100, 100, "Download Triggered!");
                FileSaver(content, `Generated_Documents_Bulk_${Date.now()}.zip`);
                resetProgress();
            });
        });
    });

    // Action B Strategy: Single Merged Document Engine via Native XML Document Node Cloning
    btnMerge.addEventListener('click', function() {
        if(!templateArrayBuffer || !spreadsheetRecords.length) return;

        let aggregatedZip = null;
        let masterBodyNode = null;
        let xmlnsText = "";
        let count = 0;

        batchProcess(spreadsheetRecords, (record, i) => {
            try {
                const zip = new ZipContainer(templateArrayBuffer);
                const doc = new DocxEngine(zip, { paragraphLoop: true, linebreaks: true });
                doc.render(record);

                // Fetch raw generated structure back from OpenXML text matrix file mapping
                const rawXmlText = doc.getZip().file("word/document.xml").asText();
                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(rawXmlText, "application/xml");
                const currentBody = xmlDoc.getElementsByTagName("w:body")[0];

                if (count === 0) {
                    // Cache root template structural envelope definition references
                    aggregatedZip = doc.getZip();
                    masterBodyNode = currentBody;
                    // Extract root structural namespaces safely
                    const docElement = xmlDoc.documentElement;
                    xmlnsText = Array.from(docElement.attributes)
                                     .map(attr => `${attr.nodeName}="${attr.nodeValue}"`)
                                     .join(" ");
                } else {
                    // Append page break prior to merging sequential bodies
                    const pageBreakXml = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
                    const pageBreakFragment = parser.parseFromString(
                        `<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${pageBreakXml}</w:root>`, 
                        "application/xml"
                    ).documentElement.firstChild;
                    
                    const importedBreak = xmlDoc.importNode(pageBreakFragment, true);
                    masterBodyNode.insertBefore(importedBreak, masterBodyNode.lastChild);

                    // Import and append body nodes sequentially
                    const childNodes = Array.from(currentBody.childNodes);
                    childNodes.forEach(node => {
                        // Avoid double embedding the terminal section structural layout element block
                        if (node.nodeName !== "w:sectPr") {
                            const importedNode = xmlDoc.importNode(node, true);
                            masterBodyNode.insertBefore(importedNode, masterBodyNode.lastChild);
                        }
                    });
                }
                count++;
            } catch (ex) {
                console.error(`Error merging records context data at line index ${i}:`, ex);
            }
        }, () => {
            try {
                updateProgress(90, 100, "Finalizing Document Structural Core Layout XML Elements...");
                
                // Package the updated master DOM back into structural content format text
                const serializer = new XMLSerializer();
                let bodyXmlContent = serializer.serializeToString(masterBodyNode);
                
                // Reconstruct full legal openXML structural payload envelope packaging container
                const fullDocumentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${xmlnsText}>${bodyXmlContent}</w:document>`;
                
                // Write completely generated compound binary data structure back to memory location
                aggregatedZip.file("word/document.xml", fullDocumentXml);
                const mergedBlob = aggregatedZip.generate({ type: "blob", compression: "DEFLATE" });
                
                updateProgress(100, 100, "Download Triggered!");
                FileSaver(mergedBlob, `Merged_Compilation_Output_${Date.now()}.docx`);
            } catch(e) {
                console.error("Critical operational compilation crash on output serialisation formatting step:", e);
                alert("Compilation Error: XML Serializer encountered an invalid document construction rule.");
            }
            resetProgress();
        });
    });

})();