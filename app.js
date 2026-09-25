    // =======================================================
    // 1. GLOBAL VARIABLES & APP STATE
    // =======================================================
    const slotNames = ["DLP Observation Image", "NMT Captured Image", "Captured in AW image", "Vendor Captured Image"];
    let currentImgs = [null, null, null, null];
    let draftData = { obsId: '', jobId: '', circle: '', partner: '', imgs: [null, null, null, null] };
    let records = [];
    let navIndex = -1; 
    let modalTask = '';
    
    // File System & Backup Variables
    let backupFolderHandle = null;
    let currentActiveFileHandle = null;
    let currentActiveFileName = '';
    let backupJsonFilesData = {}; 
    let backupSearchCache = {}; 
    let activeConversionMode = 'WITH_BOXES';
    let toastTimeout = null;
    let searchGlowTimeout = null;
    
    // EXCEL Data Management Variables
    let excelDatabaseIndex = {}; 
    let globalWorkbook = null;
    let globalRawRows = [];
    let globalExcelFileName = "Exported_OBS_Matrix.xlsx";
    let activeExcelRowIndexes = []; 

    // Search and Undo functionality
    let clearPageBackup = null;
    let searchResults = [];
    let searchCurrentIndex = 0;

    // Default NC Suggestion List
    const defaultNCList = [
        "Tower premises.", 
	"OK as per AW Specs.",
	"Sand Filling not done.", 
	"Lat/Long mismatched.", 
	"RM installed /captured as per AW spec.",
        "CP/Red Stone installed/captured as per AW spec.", 
	"Approved based on latest image attached.",
        "Need to close observation as per NMT remarks.", 
	"Need to capture rectification images as per NMT remarks.",
        "Physical pit need to capture for rectification.", 
	"Provide relevant images as per remark for further validation.",
	"Backfilling filling need to be done.", 
        "Backfilling & compaction need to be done.", 
	"Low depth needs to rectify.",
        "Protection not as per spec.", 
	"GI+PCC required as per depth.",
	"Used Protection not captured- Need to show protection with depth.",
	"Low depth protection required -Need to capture rectification image.",
        "Image over image-Need to capture original site images for further validation.", 
	"Approved based on approval attached in deviation.",
	"Protection not given as per spec -Hight*with of protection not as per specs.",
	"Need to provide supporting images for further validation of NMT remarks.",
        "Critical -- cable visible in surface without any protection, cable Protection required.", 
	"I need a photo showing the height and width of the protection. Please capture it from a slightly higher angle to get a wide shot of the entire protection." ,
	"Top Lid Depth wrongly captured .", 
	"MH Top Lid PCC not done.",
	"MH Top Lid PCC and Backfilling not done.,",
	"Top Lid Depth not measured/captured .", 
	"MH install in NGL/approval attached in deviation.",
        "Base CC not done/visible .", "Airtel embossing not done/Visible .", 
	"Fibers Breaks need to be rectify as per NMT remarks.",
        "Fibers Breaks need to be rectify as per mentioned fiber numbers.", 
	"OTDR Losses need to be rectify as per mentioned fiber numbers.",
        "Optical loss above/below/above acceptable limit- need to be rectify as per NMT remarks."
    ];
    let customNCList = [];

    // =======================================================
    // 2. INITIALIZATION & STORAGE FUNCTIONS
    // =======================================================
    
    // App Init function runs on load
    function init() { 
        loadSavedNCList(); 
        renderSlots();     
        updateCounts();    
        updateCategoryBadge(); 
    }

    // Load custom NC list from local storage
    function loadSavedNCList() {
        const saved = localStorage.getItem("obs_custom_nc_list");
        if (saved) { try { customNCList = JSON.parse(saved); } catch(e) { customNCList = [...defaultNCList]; } } 
        else { customNCList = [...defaultNCList]; }
        updateNCDatalist();
    }

    // Update the dropdown datalist with NC items
    function updateNCDatalist() {
        const datalist = document.getElementById('nc-datalist');
        datalist.innerHTML = customNCList.map(nc => `<option value="${nc}"></option>`).join('');
    }

    function openNCManager() {
        document.getElementById('nc-manager-textarea').value = customNCList.join('\n');
        document.getElementById('nc-manager-modal').classList.remove('hidden');
    }
    function closeNCManager() { document.getElementById('nc-manager-modal').classList.add('hidden'); }

    function resetDefaultNCs() {
        if(confirm("Are you sure you want to revert to the original list? Your custom additions will be lost.")) {
            document.getElementById('nc-manager-textarea').value = defaultNCList.join('\n');
            showToast("Restored Default List", "#FF9500");
        }
    }

    function saveNCList() {
        const text = document.getElementById('nc-manager-textarea').value;
        const newArr = text.split('\n').map(s => s.trim()).filter(s => s.length > 0);
        customNCList = newArr;
        localStorage.setItem("obs_custom_nc_list", JSON.stringify(customNCList));
        updateNCDatalist();
        closeNCManager();
        showToast("NC List Successfully Updated! ✅", "#34C759");
    }

	// Save repeating text fields to session cache automatically
    function saveHeaderFields() {
        sessionStorage.setItem('saved_jobId', document.getElementById('jobId').value);
        sessionStorage.setItem('saved_circle', document.getElementById('circle').value);
        sessionStorage.setItem('saved_partner', document.getElementById('partner').value);
    }

    // Restore text fields from cache (but clear on browser refresh)
    function loadHeaderFields() {
        const isReload = window.performance && window.performance.getEntriesByType("navigation")[0]?.type === "reload";
        if (isReload) {
            sessionStorage.removeItem('saved_jobId');
            sessionStorage.removeItem('saved_circle');
            sessionStorage.removeItem('saved_partner');
            return;
        }
        if(sessionStorage.getItem('saved_jobId')) document.getElementById('jobId').value = sessionStorage.getItem('saved_jobId');
        if(sessionStorage.getItem('saved_circle')) document.getElementById('circle').value = sessionStorage.getItem('saved_circle');
        if(sessionStorage.getItem('saved_partner')) document.getElementById('partner').value = sessionStorage.getItem('saved_partner');
    }
    window.addEventListener('DOMContentLoaded', loadHeaderFields);
    // =======================================================
    // 3. EXCEL FILL FEATURE LOGIC
    // =======================================================
    
    // Open Excel Data modal to update NCs matching current OBS ID
    function openExcelDataModal() {
        const queryObsId = document.getElementById('obsId').value.trim();
        if (!queryObsId) {
            showToast("⚠️ Type an OBS ID first to find target rows!", "#FF3B30");
            document.getElementById('obsId').focus();
            return;
        }

        if (!globalRawRows || globalRawRows.length === 0) {
            showToast("⚠️ Connect Excel Matrix File from Hub First!", "#FF9500");
            return;
        }

        document.getElementById('excel-fill-target-id').innerText = `[${queryObsId}]`;
        const tbody = document.getElementById('excel-fill-table-body');
        tbody.innerHTML = '';
        activeExcelRowIndexes = [];

        let matchFound = false;
        const todayDate = new Date().toLocaleDateString('en-GB'); 

        for (let i = 0; i < globalRawRows.length; i++) {
            const row = globalRawRows[i];
            if (row && row[0] && String(row[0]).trim() === queryObsId) {
                matchFound = true;
                activeExcelRowIndexes.push(i);
                
                const subCat = row[14] ? String(row[14]) : "N/A";
                const currentStatus = row[29] ? String(row[29]) : "";
                const currentRemark = row[30] ? String(row[30]) : "";

                const rowHtml = `
                    <tr class="hover:bg-gray-100/50 dark:hover:bg-[#202024] transition-colors">
			            <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700">
                            <div class="subcat-cell truncate max-w-[8rem] text-[11px] font-semibold text-gray-500" title="Suraj">Suraj</div>
                            <input type="hidden" id="excel-zqh-${i}" value="Suraj">
                        </td>
                        <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700">
                            <div class="subcat-cell truncate max-w-[8rem] text-[11px] font-semibold text-gray-500" title="${todayDate}">${todayDate}</div>
                            <input type="hidden" id="excel-date-${i}" value="${todayDate}">
                        </td>
                        <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700">
                            <div class="subcat-cell truncate max-w-[12rem] text-[11px] font-semibold" title="${subCat}">${subCat}</div>
                        </td>
                        <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700">
                            <select id="excel-status-${i}" class="w-full text-xs p-2 rounded-lg border border-gray-300 dark:border-zinc-600 bg-white dark:bg-[#25252a]">
                                <option value="" disabled hidden>Select Status</option>
                                <option value="Hold" ${currentStatus === 'Hold' ? 'selected' : ''}>Hold</option>
				               <option value="ZQH Closed" ${currentStatus === 'ZQH Closed' ? 'selected' : ''}>ZQH Closed</option>
				               <option value="Vendor WIP" ${!currentStatus || currentStatus === 'Vendor WIP' ? 'selected' : ''}>Vendor WIP</option>
                            </select>
                        </td>
                        <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700">
                            <input type="text" id="excel-remark-${i}" list="nc-datalist" value="${currentRemark}" class="w-full text-xs p-2 rounded-lg border border-gray-300 dark:border-zinc-600 bg-white dark:bg-[#25252a] focus:ring-1 focus:ring-[#007AFF]" placeholder="Type or select NC...">
                        </td>
                        
                        <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700 text-center">
                            <button onclick="copyExcelRemark(${i})" class="bg-gray-200 dark:bg-zinc-700 text-gray-800 dark:text-white px-2.5 py-1.5 rounded-lg text-[10px] font-bold hover:bg-[#007AFF] hover:text-white dark:hover:bg-[#007AFF] transition-all" title="Copy Remark">📋 Copy</button>
                        </td>
                    </tr>
                `;
                tbody.insertAdjacentHTML('beforeend', rowHtml);
            }
        }

        if (!matchFound) { 
            // 4. Excel Fallback: Agar ID missing hai toh ek naya blank row banayein aur render karein
            let newRowIndex = globalRawRows.length;
            let newRow = new Array(33).fill("");
            newRow[0] = queryObsId; 
            newRow[1] = document.getElementById('jobId').value.trim();
            globalRawRows.push(newRow);
            activeExcelRowIndexes.push(newRowIndex);
            
            const rowHtml = `
                <tr class="hover:bg-gray-100/50 dark:hover:bg-[#202024] transition-colors">
                    <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700">
                        <div class="subcat-cell truncate max-w-[8rem] text-[11px] font-semibold text-gray-500" title="New Entry">New Entry</div>
                        <input type="hidden" id="excel-zqh-${newRowIndex}" value="New Entry">
                    </td>
                    <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700">
                        <div class="subcat-cell truncate max-w-[8rem] text-[11px] font-semibold text-gray-500" title="${todayDate}">${todayDate}</div>
                        <input type="hidden" id="excel-date-${newRowIndex}" value="${todayDate}">
                    </td>
                    <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700">
                        <div class="subcat-cell truncate max-w-[12rem] text-[11px] font-semibold" title="N/A">N/A</div>
                    </td>
                    <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700">
                        <select id="excel-status-${newRowIndex}" class="w-full text-xs p-2 rounded-lg border border-gray-300 dark:border-zinc-600 bg-white dark:bg-[#25252a]">
                            <option value="" disabled hidden>Select Status</option>
                            <option value="Hold">Hold</option>
                            <option value="ZQH Closed">ZQH Closed</option>
                            <option value="Vendor WIP" selected>Vendor WIP</option>
                        </select>
                    </td>
                    <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700">
                        <input type="text" id="excel-remark-${newRowIndex}" list="nc-datalist" value="" class="w-full text-xs p-2 rounded-lg border border-gray-300 dark:border-zinc-600 bg-white dark:bg-[#25252a] focus:ring-1 focus:ring-[#007AFF]" placeholder="Type or select NC...">
                    </td>
                    <td class="px-4 py-3 border-b border-gray-200 dark:border-zinc-700 text-center">
                        <button onclick="copyExcelRemark(${newRowIndex})" class="bg-gray-200 dark:bg-zinc-700 text-gray-800 dark:text-white px-2.5 py-1.5 rounded-lg text-[10px] font-bold hover:bg-[#007AFF] hover:text-white dark:hover:bg-[#007AFF] transition-all" title="Copy Remark">📋 Copy</button>
                    </td>
                </tr>
            `;
            tbody.insertAdjacentHTML('beforeend', rowHtml);
            showToast("New OBS ID attached. Ready to fill!", "#34C759");
        }
        document.getElementById('excel-fill-modal').classList.remove('hidden');
    }

    function closeExcelDataModal() { document.getElementById('excel-fill-modal').classList.add('hidden'); }

    // Copy remark field text
    function copyExcelRemark(rowIndex) {
        const input = document.getElementById(`excel-remark-${rowIndex}`);
        if(input && input.value) {
            navigator.clipboard.writeText(input.value).then(() => { showToast("📋 Copied to clipboard!", "#007AFF"); }).catch(() => showToast("❌ Clipboard lock error", "#FF3B30"));
        } else { showToast("Nothing to copy!", "#FF9500"); }
    }

    // Save selected data back to memory
    function handleNCSaveAndModal(actionType) {
        if (globalWorkbook && globalRawRows.length > 0) {
            activeExcelRowIndexes.forEach(rowIndex => {
                const statusVal = document.getElementById(`excel-status-${rowIndex}`).value;
                const remarkVal = document.getElementById(`excel-remark-${rowIndex}`).value;
                const zqhVal = document.getElementById(`excel-zqh-${rowIndex}`).value;
                const dateVal = document.getElementById(`excel-date-${rowIndex}`).value;

                while(globalRawRows[rowIndex].length <= 32) {
                    globalRawRows[rowIndex].push("");
                }

                globalRawRows[rowIndex][29] = statusVal;  // Column AD
                globalRawRows[rowIndex][30] = remarkVal;  // Column AE
                globalRawRows[rowIndex][31] = zqhVal;     // Column AF
                globalRawRows[rowIndex][32] = dateVal;    // Column AG
            });

            try {
                const newWorksheet = XLSX.utils.aoa_to_sheet(globalRawRows);
                globalWorkbook.Sheets[globalWorkbook.SheetNames[0]] = newWorksheet;
                showToast("✅ Excel Data updated in memory!", "#34C759");
                syncToActiveFile();
            } catch (error) {
                console.error(error);
                showToast("❌ Error updating data in memory!", "#FF3B30");
            }
        }
        
        closeExcelDataModal();
        openModal(actionType);
    }

    // Trigger physical export of the updated Excel file
    function exportUpdatedExcel() {
        if (!globalWorkbook) {
            showToast("⚠️ No Excel data found. Please link Excel file first.", "#FF9500");
            return;
        }
        try {
            showToast("⏳ Generating Updated Excel File...", "#007AFF");
            XLSX.writeFile(globalWorkbook, "Updated_" + globalExcelFileName);
            showToast("✅ Excel File Exported Successfully!", "#34C759");
            closeModal();
        } catch (error) {
            console.error(error);
            showToast("❌ File Download Failed!", "#FF3B30");
        }
    }

    // =======================================================
    // 4. THEME & JSON MERGING LOGIC
    // =======================================================
    
    function openMasterConvertModal() { document.getElementById('master-convert-modal').classList.remove('hidden'); }
    function closeMasterConvertModal() { document.getElementById('master-convert-modal').classList.add('hidden'); }
    
    function routeToWordToJson() {
        closeMasterConvertModal();
        openConverterModal();
    }

    function triggerJsonMergeSelect() {
        document.getElementById('merge-json-input').click();
    }

    // Merge multiple JSON logic
    async function processMergeJson(input) {
        if (!input.files || input.files.length === 0) return;
        closeMasterConvertModal();
        
        let mergeFileNamePrompt = prompt("Please enter the name for the merged file:", "Merged_OBS_Data");
        if (!mergeFileNamePrompt) { input.value = ""; return; }
        mergeFileNamePrompt = mergeFileNamePrompt.endsWith(".json") ? mergeFileNamePrompt : mergeFileNamePrompt + ".json";
        
        showToast("⏳ Merging Data efficiently...", "#FF9500");
        
        try {
            let masterRecords = [];
            let masterExcelRows = [];
            let masterExcelName = "Exported_OBS_Matrix.xlsx";
            
            for (let i = 0; i < input.files.length; i++) {
                const text = await input.files[i].text();
                const parsed = JSON.parse(text);
                
                if (masterExcelRows.length === 0 && parsed.version === "2.0" && parsed.excelRawRows && parsed.excelRawRows.length > 0) {
                    masterExcelRows = parsed.excelRawRows;
                    masterExcelName = parsed.excelFileName || masterExcelName;
                }
                
                const recordsArray = Array.isArray(parsed) ? parsed : (parsed.records || []);
                for (let j = 0; j < recordsArray.length; j++) masterRecords.push(recordsArray[j]);
            }
            
            const payload = { version: "2.0", records: masterRecords, excelRawRows: masterExcelRows || [], excelFileName: masterExcelName };
            triggerJsonFileDownload(payload, mergeFileNamePrompt);
            showToast(`✅ Successfully merged ${masterRecords.length} records!`, "#34C759");
        } catch (error) {
            console.error(error);
            showToast("❌ Parsing Error while Merging!", "#FF3B30");
        } finally {
            input.value = ""; 
        }
    }
    
    // Day and Night Theme switch
    function toggleThemeMode() {
        const html = document.documentElement; const body = document.body; const themeBtn = document.getElementById('theme-toggle-btn');
        html.classList.toggle('dark'); body.classList.toggle('dark-mode'); body.classList.toggle('dark'); 
        if(body.classList.contains('dark-mode')) { themeBtn.innerText = "🌞"; showToast("🌚 System Dark Theme Activated", "#1c1c1e"); } 
        else { themeBtn.innerText = "🌚"; showToast("🌞 System Light Theme Activated", "#007AFF"); }
    }

    // Modal Helpers
    function openHubModal() { document.getElementById('hub-modal').classList.remove('hidden'); }
    function closeHubModal() { document.getElementById('hub-modal').classList.add('hidden'); }
    function openConverterModal() { document.getElementById('converter-modal').classList.remove('hidden'); }
    function closeConverterModal() { document.getElementById('converter-modal').classList.add('hidden'); }
    function triggerExcelImport() { document.getElementById('excel-file-input').click(); }

    // Read and parse Excel file
    function loadExcelDatabase(input) {
        if (!input.files.length) return;
        const file = input.files[0];
        globalExcelFileName = file.name;
        const reader = new FileReader();
        showToast("Processing Matrix Grid...", "#FF9500");
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                globalWorkbook = XLSX.read(data, { type: 'array' });
                const worksheet = globalWorkbook.Sheets[globalWorkbook.SheetNames[0]];
                globalRawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                
                excelDatabaseIndex = {}; let count = 0;
                globalRawRows.forEach(row => {
                    if (row && row[0]) {
                        const obsIdKey = String(row[0]).trim();
                        const latVal = row[5] ? String(row[5]).trim() : '';
                        const longVal = row[6] ? String(row[6]).trim() : '';
                        if(obsIdKey && (latVal || longVal)) {
                            excelDatabaseIndex[obsIdKey] = { lat: latVal, lng: longVal }; 
                            count++;
                        }
                    }
                });

                document.getElementById('hub-excel-status').innerText = "🟢 Connected ("+ (globalRawRows.length-1) +" rows)";
                document.getElementById('hub-excel-status').className = "text-[10px] text-emerald-500 font-bold mt-0.5";
                evaluateHubOverallStatus();
                showToast(`Synchronized successfully!`, "#34C759");
            } catch(err) { console.error(err); showToast("Invalid Matrix Layout Structure!", "#FF3B30"); }
        };
        reader.readAsArrayBuffer(file);
    }

    // Coordinates copy functionality
    function searchAndCopyExcelCoordinates(silent = false) {
        const queryObsId = document.getElementById('obsId').value.trim();
        if (!queryObsId) { if(!silent) showToast("⚠️ Missing Primary OBS ID Data Link!", "#FF3B30"); return false; }
        if (Object.keys(excelDatabaseIndex).length === 0) { if(!silent) showToast("⚠️ Excel Link Database Offline!", "#FF9500"); return false; }
        
        const match = excelDatabaseIndex[queryObsId];
        if (!match || !match.lat || !match.lng) { if(!silent) showToast(`❌ Element not found for [${queryObsId}]`, "#FF3B30"); return false; }
        
        const mergedCoordinates = `${match.lat}, ${match.lng}`;
        navigator.clipboard.writeText(mergedCoordinates).then(() => {
            showToast(`✅ Copied Coordinates: ${mergedCoordinates}`, "#34C759");
        }).catch(() => { if(!silent) showToast("❌ Clipboard permission block!", "#FF3B30"); });
        return true;
    }

    function updateButtonProgress(percent, label = "Processing") {
        const btn = document.getElementById('converter-btn');
        if (percent === 0) { btn.innerHTML = `⏳ Parsing...`; btn.className = "bg-[#ff453a] text-white px-3 py-2 rounded-xl font-bold text-xs pointer-events-none"; } 
        else if (percent === 100) { btn.innerHTML = `🧬 Convert`; btn.className = "bg-gray-100 dark:bg-[#25252a] text-gray-800 dark:text-white px-3 py-2 rounded-xl font-bold text-xs hover:bg-[#FF9500] hover:text-white dark:hover:bg-[#FF9500] transition-all flex items-center gap-1"; } 
        else { btn.innerHTML = `⏳ ${label} ${percent}%`; }
    }

    function triggerWordConversion(mode) { activeConversionMode = mode; document.getElementById('word-file-input').click(); }

    // Async Image Array Compressor for Word to JSON mapping
    async function compressImageArray(imgList) {
        let compressed = [];
        for (let img of imgList) {
            if (img) compressed.push(await compressImage(img));
            else compressed.push(null);
        }
        return compressed;
    }

    // Word doc processing function (Extract JSON)
    async function processWordToJson() {
        const fileInput = document.getElementById('word-file-input');
        if (!fileInput.files.length) return;
        const file = fileInput.files[0]; closeConverterModal(); updateButtonProgress(0);
        await new Promise(resolve => requestAnimationFrame(resolve));

        const reader = new FileReader();
        reader.onload = function (loadEvent) {
            mammoth.convertToHtml({ arrayBuffer: loadEvent.target.result }).then(async function (result) {
                const tempDiv = document.createElement('div'); tempDiv.innerHTML = result.value;
                let extractedRecords = [];
                
                if (activeConversionMode === 'WITH_BOXES') {
                    const tables = Array.from(tempDiv.getElementsByTagName('table'));
                    for (let i = 0; i < tables.length; i++) {
                        let pct = Math.floor(((i + 1) / tables.length) * 100); updateButtonProgress(pct, "Box Core");
                        if (tables[i].innerText && tables[i].innerText.includes("OBS Task No.")) {
                            let matchObj = parseMetadataString(tables[i].innerText);
                            if (matchObj) {
                                const imgElements = tables[i].getElementsByTagName('img');
                                let imageArray = [null, null, null, null];
                                // Compress images on the fly during extraction
                                for (let j = 0; j < Math.min(imgElements.length, 4); j++) {
                                    imageArray[j] = await compressImage(imgElements[j].src);
                                }
                                extractedRecords.push({ category: "ZQH Validation Closed", obsId: matchObj.obsId, jobId: matchObj.jobId, circle: matchObj.circle, partner: matchObj.partner, imgs: imageArray });
                            }
                        }
                    }
                } else {
                    const nodes = Array.from(tempDiv.querySelectorAll('p, table, img, div'));
                    let currentRecord = null; let trackingImages = [];
                    for (let i = 0; i < nodes.length; i++) {
                        let pct = Math.floor(((i + 1) / nodes.length) * 100); updateButtonProgress(pct, "Extract");
                        const text = nodes[i].innerText || nodes[i].textContent || "";
                        if (text.includes("OBS Task No.")) {
                            if (currentRecord) { 
                                currentRecord.imgs = await compressImageArray(padOrTrimImages(trackingImages)); 
                                extractedRecords.push(currentRecord); 
                            }
                            trackingImages = []; let matchObj = parseMetadataString(text);
                            if (matchObj) currentRecord = { category: "ZQH Validation Closed", obsId: matchObj.obsId, jobId: matchObj.jobId, circle: matchObj.circle, partner: matchObj.partner, imgs: [null, null, null, null] };
                        } else if (currentRecord) {
                            if (nodes[i].tagName === 'IMG' && nodes[i].src) { if (!trackingImages.includes(nodes[i].src)) trackingImages.push(nodes[i].src); }
                            else {
                                const childImgs = nodes[i].getElementsByTagName('img');
                                for (let g = 0; g < childImgs.length; g++) { if (childImgs[g].src && !trackingImages.includes(childImgs[g].src)) trackingImages.push(childImgs[g].src); }
                            }
                        }
                    }
                    if (currentRecord) { 
                        currentRecord.imgs = await compressImageArray(padOrTrimImages(trackingImages)); 
                        extractedRecords.push(currentRecord); 
                    }
                }

                // Categorize records based on image count
                extractedRecords.forEach(rec => {
                    let totalImgs = rec.imgs.filter(x => x !== null).length;
                    if(totalImgs === 3) rec.category = "ZQH Closed";
                    else if(totalImgs < 3) rec.category = "Vendor WIP";
                    else rec.category = "ZQH Validation Closed";
                });

                updateButtonProgress(100);
                if (extractedRecords.length === 0) { showToast("No metadata objects found!", "#FF3B30"); return; }
                triggerJsonFileDownload(extractedRecords, file.name.replace('.docx', '.json'));
                showToast(`Parsed ${extractedRecords.length} Elements!`, "#34C759");
                fileInput.value = ''; 
            }).catch(() => { updateButtonProgress(100); showToast("Parser Error!", "#FF3B30"); });
        };
        reader.readAsArrayBuffer(file);
    }

    function padOrTrimImages(imgList) {
        let output = [null, null, null, null];
        for (let i = 0; i < Math.min(imgList.length, 4); i++) output[i] = imgList[i];
        return output;
    }

    // Clean word document metadata text strings
    function parseMetadataString(rawStr) {
        try {
            let clean = rawStr.replace("OBS Task No.", "").replace(":", "").trim();
            let parts = clean.split("/");
            if (parts.length >= 4) {
                let candidatePartner = parts[3].split("\n")[0].split("\r")[0].trim();
                return { obsId: parts[0].trim(), jobId: parts[1].trim(), circle: parts[2].trim(), partner: cleanPartnerText(candidatePartner) };
            }
        } catch(e) { console.error("Malformed metadata line parse failure", e); }
        return null;
    }

    function cleanPartnerText(text) {
        const keywords = ["NMT", "Captured", "images", "Image", "Uploaded", "captured in FF1", "Vendor", "capturing", "captured in AW", "vendor captured"];
        let lowerText = text.toLowerCase(); let earliestIndex = text.length;
        keywords.forEach(word => {
            let idx = lowerText.indexOf(word.toLowerCase());
            if (idx !== -1 && idx < earliestIndex) earliestIndex = idx;
        });
        return text.substring(0, earliestIndex).trim().replace(/[:\/,\s-]+$/, "").trim();
    }

    // Blob handler to download JSON
    function triggerJsonFileDownload(objData, exportName) {
        const dataStr = JSON.stringify(objData, null, 2);
        const blob = new Blob([dataStr], { type: "application/json;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        
        const downloadAnchor = document.createElement('a');
        downloadAnchor.setAttribute("href", url); 
        downloadAnchor.setAttribute("download", exportName);
        document.body.appendChild(downloadAnchor); 
        downloadAnchor.click(); 
        downloadAnchor.remove();
        
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    function downloadKML() {
        const jobId = document.getElementById('jobId').value.trim();
        if (!jobId) { showToast("Missing System Job Reference Key!", "#FF3B30"); return; }
        window.open(`https://fibre.airtel.com/fd_images/kml/${jobId}/${jobId}_trenching.kml`, '_blank');
        showToast("Fetching Satellite Vector KML...", "#007AFF");
    }

    // Standard Toast Notification System
    function showToast(msg, color = "#333") {
        const toast = document.getElementById("toast");
        if (toastTimeout) clearTimeout(toastTimeout);
        toast.innerText = msg; toast.style.borderLeft = `4px solid ${color}`;
        toast.classList.add("show");
        toastTimeout = setTimeout(() => { toast.classList.remove("show"); }, 3500);
    }

    function openWorkEngineModal() { document.getElementById('work-engine-modal').classList.remove('hidden'); }
    function closeWorkEngineModal() { document.getElementById('work-engine-modal').classList.add('hidden'); }
    function openDuplicateModal() { document.getElementById('duplicate-modal').classList.remove('hidden'); }
    function closeDuplicateModal() { document.getElementById('duplicate-modal').classList.add('hidden'); }

    // =======================================================
    // 5. STORAGE CONNECTIONS AND LOGIC
    // =======================================================
    function processRestoredData(text, fileName) {
        if (text) {
            const parsed = JSON.parse(text);
            if (Array.isArray(parsed)) {
                records = parsed;
            } else if (parsed.version === "2.0") {
                records = parsed.records || [];
                if (parsed.excelRawRows && parsed.excelRawRows.length > 0) {
                    globalRawRows = parsed.excelRawRows;
                    globalExcelFileName = parsed.excelFileName || "Exported_OBS_Matrix.xlsx";
                    
                    excelDatabaseIndex = {};
                    globalRawRows.forEach(row => {
                        if (row && row[0]) {
                            const obsIdKey = String(row[0]).trim();
                            const latVal = row[5] ? String(row[5]).trim() : '';
                            const longVal = row[6] ? String(row[6]).trim() : '';
                            if(obsIdKey && (latVal || longVal)) excelDatabaseIndex[obsIdKey] = { lat: latVal, lng: longVal };
                        }
                    });
                    
                    const newWorksheet = XLSX.utils.aoa_to_sheet(globalRawRows);
                    globalWorkbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(globalWorkbook, newWorksheet, "Sheet1");
                    
                    document.getElementById('hub-excel-status').innerText = "🟢 Restored From JSON ("+ (globalRawRows.length-1) +" rows)";
                    document.getElementById('hub-excel-status').className = "text-[10px] text-emerald-500 font-bold mt-0.5";
                }
            }
        } else {
            records = [];
        }

        document.getElementById('hub-work-status').innerText = `🟢 Linked: ${fileName.toUpperCase()}`;
        document.getElementById('hub-work-status').className = "text-[10px] text-emerald-500 font-bold mt-0.5";
        evaluateHubOverallStatus(); updateCounts(); updateCategoryBadge();
        showToast(`Mounted Engine Storage Instance`, "#34C759");
    }

   // 🚀 EXTREMELY FAST INDEXING BUILDER
    async function buildFastSearchIndex() {
        const btnStatus = document.getElementById('hub-backup-status');
        backupSearchIndex = {};
        const filenames = Object.keys(backupJsonFilesData);
        for (let i = 0; i < filenames.length; i++) {
            btnStatus.innerText = `⏳ Indexing... (${i+1}/${filenames.length})`;
            try {
                const entry = backupJsonFilesData[filenames[i]];
                const file = await entry.getFile();
                const text = await file.text();
                // REGEX lookup (100x faster than JSON.parse on 50mb files, uses 0 RAM)
                const regex = /"obsId"\s*:\s*"([^"]+)"/g;
                let match;
                while ((match = regex.exec(text)) !== null) {
                    const id = match[1];
                    if (!backupSearchIndex[id]) backupSearchIndex[id] = new Set();
                    backupSearchIndex[id].add(filenames[i]);
                }
            } catch (e) { console.error("Index Skip:", filenames[i]); }
        }
        isIndexBuilt = true;
        btnStatus.innerText = `🟢 Connected (${filenames.length} files Indexed)`;
        btnStatus.className = "text-[10px] text-emerald-500 font-bold mt-0.5";
        showToast("System Repositories Indexed!", "#34C759");
    }

    async function connectBackupFolder() {
        if (window.showDirectoryPicker) {
            try {
                backupFolderHandle = await window.showDirectoryPicker({ mode: 'read' });
                await scanBackupFolder();
                evaluateHubOverallStatus();
                await buildFastSearchIndex();
            } catch (err) { showToast("Local execution block!", "#FF3B30"); }
        } else {
            const input = document.createElement('input'); input.type = 'file'; input.webkitdirectory = true; input.directory = true; input.multiple = true;
            input.onchange = async (e) => {
                const files = e.target.files; backupJsonFilesData = {}; backupSearchIndex = {}; isIndexBuilt = false;
                for (let i = 0; i < files.length; i++) { if (files[i].name.endsWith('.json')) { backupJsonFilesData[files[i].name] = { getFile: async () => files[i] }; } }
                backupFolderHandle = { fallback: true }; evaluateHubOverallStatus();
                await buildFastSearchIndex();
            };
            input.click();
        }
    }

    async function scanBackupFolder() {
        if (!backupFolderHandle) return; backupJsonFilesData = {}; backupSearchIndex = {}; isIndexBuilt = false;
        try { for await (const entry of backupFolderHandle.values()) { if (entry.kind === 'file' && entry.name.endsWith('.json')) { backupJsonFilesData[entry.name] = entry; } } } catch (err) { console.error("Folder scan failed", err); }
    }

    async function workSelectExistingFile() {
        closeWorkEngineModal();
        if (window.showOpenFilePicker) {
            try {
                const [handle] = await window.showOpenFilePicker({ types: [{ description: 'JSON Files', accept: { 'application/json': ['.json'] } }], multiple: false });
                if (handle) { currentActiveFileHandle = handle; currentActiveFileName = handle.name; const file = await currentActiveFileHandle.getFile(); const text = await file.text(); processRestoredData(text, currentActiveFileName); }
            } catch(e) { showToast("Operation Interrupted", "#FF3B30"); }
        } else {
            const input = document.createElement('input'); input.type = 'file'; input.accept = '.json';
            input.onchange = async (e) => {
                if (e.target.files.length > 0) { const file = e.target.files[0]; currentActiveFileHandle = { fallback: true }; currentActiveFileName = file.name; const text = await file.text(); processRestoredData(text, currentActiveFileName); }
            }; input.click();
        }
    }

    async function workCreateNewFile() {
        closeWorkEngineModal();
        if (window.showSaveFilePicker) {
            try {
                const handle = await window.showSaveFilePicker({ suggestedName: 'new_obs_report.json', types: [{ description: 'JSON Files', accept: { 'application/json': ['.json'] } }] });
                if (handle) {
                    currentActiveFileHandle = handle; currentActiveFileName = handle.name; records = []; const writable = await currentActiveFileHandle.createWritable(); const payload = { version: "2.0", records: [], excelRawRows: globalRawRows || [], excelFileName: globalExcelFileName };
                    await writable.write(JSON.stringify(payload, null, 2)); await writable.close(); document.getElementById('hub-work-status').innerText = `🟢 Linked: ${currentActiveFileName.toUpperCase()}`; document.getElementById('hub-work-status').className = "text-[10px] text-emerald-500 font-bold mt-0.5"; evaluateHubOverallStatus(); updateCounts(); updateCategoryBadge(); showToast(`Initialized New Workspace Grid`, "#34C759");
                }
            } catch(e) { showToast("Initialization Aborted", "#FF3B30"); }
        } else {
            currentActiveFileHandle = { fallback: true }; currentActiveFileName = 'new_obs_report.json'; records = []; document.getElementById('hub-work-status').innerText = `🟢 Linked: ${currentActiveFileName.toUpperCase()}`; document.getElementById('hub-work-status').className = "text-[10px] text-emerald-500 font-bold mt-0.5"; evaluateHubOverallStatus(); updateCounts(); updateCategoryBadge(); showToast(`Workspace Initialized (Firefox Compatible)`, "#34C759");
        }
    }

    function evaluateHubOverallStatus() {
        const hasExcel = Object.keys(excelDatabaseIndex).length > 0; const hasBackup = backupFolderHandle !== null; const hasWork = currentActiveFileHandle !== null; const hubBtn = document.getElementById('hub-master-btn');
        if (hasExcel && hasBackup && hasWork) { hubBtn.className = "bg-[#34C759] text-white px-3.5 py-2 rounded-xl font-bold text-xs flex items-center gap-1.5 shadow-md active:scale-95 transition-all whitespace-nowrap"; document.getElementById('hub-master-icon').innerText = "✅"; document.getElementById('hub-master-text').innerText = "LIVE"; } 
        else { hubBtn.className = "bg-gray-100 dark:bg-[#25252a] text-gray-800 dark:text-white px-3.5 py-2 rounded-xl font-bold text-xs flex items-center gap-1.5 shadow-sm hover:bg-[#5856D6] hover:text-white dark:hover:bg-[#5856D6] active:scale-95 transition-all whitespace-nowrap"; document.getElementById('hub-master-icon').innerText = "⛓️‍💥"; document.getElementById('hub-master-text').innerText = "CONNECT DATABASE"; }
    }

    async function syncToActiveFile() {
        if (!currentActiveFileHandle) return;
        try {
            const payload = { version: "2.0", records: records, excelRawRows: globalRawRows || [], excelFileName: globalExcelFileName };
            if (currentActiveFileHandle.fallback) { triggerJsonFileDownload(payload, currentActiveFileName); updateCounts(); } 
            else { const writable = await currentActiveFileHandle.createWritable(); await writable.write(JSON.stringify(payload, null, 2)); await writable.close(); updateCounts(); }
        } catch (err) { showToast("Storage file system lock error!", "#FF3B30"); }
    }

    // =======================================================
    // 6. GLOBAL SEARCH & FILTERING
    // =======================================================

    async function handleObsIdInput() {
        invalidateClearBackup(); 
        const searchBtn = document.querySelector("button[onclick='globalDeepSearch()']");
        if (searchBtn) searchBtn.classList.remove('red-breath');

        // --- NEW AUTO FIND GLOW LOGIC (DRAFT MODE ONLY) ---
        if (navIndex === -1 && searchBtn) {
            const query = document.getElementById('obsId').value.trim();
            searchBtn.classList.remove('search-glow-active');
            if (searchGlowTimeout) clearTimeout(searchGlowTimeout);

            if (query.length > 0) {
                let exists = false;
                // 1. Check current stream
                if (records.some(r => r.obsId === query)) exists = true;
                
                // 2. Check backup index if connected
                if (!exists && typeof isIndexBuilt !== 'undefined' && isIndexBuilt && typeof backupSearchIndex !== 'undefined' && backupSearchIndex[query]) {
                    exists = true;
                }

                if (exists) {
                    searchBtn.classList.add('search-glow-active');
                    searchGlowTimeout = setTimeout(() => {
                        searchBtn.classList.remove('search-glow-active');
                    }, 5000);
                }
            }
        }
    }

    async function globalDeepSearch() {
        const query = document.getElementById('obsId').value.trim();
        if (!query) { showToast("Please input Target Primary ID", "#FF3B30"); return; }
        
        showToast("Searching index...", "#FF9500");
        let masterResults = [];

        // 1. Current Work Stream Memory Search
        records.forEach((record, idx) => { 
            if (record.obsId === query) { masterResults.push({ source: 'CURRENT_WORK_FILE', fileName: currentActiveFileName || 'Active File', index: idx, data: record }); }
        });

        // 2. High-Speed Index Search
        if (typeof isIndexBuilt !== 'undefined' && isIndexBuilt && backupSearchIndex[query]) {
            const filesToOpen = Array.from(backupSearchIndex[query]);
            for (let filename of filesToOpen) {
                if(filename === currentActiveFileName) continue; // Skip if it's our active file
                try {
                    const entry = backupJsonFilesData[filename]; const file = await entry.getFile(); const text = await file.text();
                    const parsedData = JSON.parse(text); 
                    const targetArr = Array.isArray(parsedData) ? parsedData : (parsedData.version === "2.0" ? parsedData.records : []);
                    targetArr.forEach((record, idx) => {
                        if (record.obsId === query) { masterResults.push({ source: 'BACKUP_REPOSITORY', fileName: filename, index: idx, data: record }); }
                    });
                } catch (e) { console.error("Could not extract matched file", e); }
            }
        } else if (typeof isIndexBuilt !== 'undefined' && !isIndexBuilt) {
            showToast("Search Index Not Connected from Hub!", "#FF3B30");
        }

        if (masterResults.length === 0) { showToast("Element record not found globally", "#FF3B30"); return; }
        searchResults = masterResults; searchCurrentIndex = 0; renderPaginatedSearchResult(); openDuplicateModal();
    }

    function renderPaginatedSearchResult() {
        const listContainer = document.getElementById('duplicate-list');
        if (!searchResults.length) { listContainer.innerHTML = ''; return; }
        
        const currentMatch = searchResults[searchCurrentIndex]; const r = currentMatch.data;
        const badgeText = currentMatch.source === 'CURRENT_WORK_FILE' ? 'Active Stream File' : 'Backup Storage';
        const badgeStyle = currentMatch.source === 'CURRENT_WORK_FILE' ? 'bg-emerald-500/10 text-emerald-500' : 'bg-blue-500/10 text-blue-500';
        
        let categoryBadgeColor = "bg-gray-500/10 text-gray-400";
        if(r.category === "ZQH Closed") categoryBadgeColor = "bg-teal-500/10 text-teal-400 dark:text-teal-300"; 
        else if(r.category === "ZQH Validation Closed") categoryBadgeColor = "bg-amber-500/10 text-amber-500 dark:text-amber-400"; 
        else if(r.category === "Vendor WIP") categoryBadgeColor = "bg-blue-500/10 text-blue-500 dark:text-blue-400";
        else if(r.category === "Hold") categoryBadgeColor = "bg-pink-500/10 text-pink-500 dark:text-pink-400";

        const activeImgs = (r.imgs || []).filter(x => x); const subImgs = activeImgs.slice(1); let structuralGridHtml = "";

        if (activeImgs.length === 3) { structuralGridHtml = `<div class="grid grid-cols-2 gap-2 mt-2"><div class="border border-gray-300 dark:border-zinc-700 p-2 text-center bg-gray-50 dark:bg-[#141417] rounded-xl"><div class="text-[10px] font-bold text-gray-400 dark:text-zinc-500 mb-1 uppercase">NMT CAPTURED</div><img src="${subImgs[0]}" class="h-36 w-full object-fill rounded-lg border border-gray-200 dark:border-zinc-800"></div><div class="border border-gray-300 dark:border-zinc-700 p-2 text-center bg-gray-50 dark:bg-[#141417] rounded-xl"><div class="text-[10px] font-bold text-gray-400 dark:text-zinc-500 mb-1 uppercase">CAPTURED IN AW</div><img src="${subImgs[1]}" class="h-36 w-full object-fill rounded-lg border border-gray-200 dark:border-zinc-800"></div></div>`; } 
        else if (activeImgs.length === 4) { structuralGridHtml = `<div class="grid grid-cols-3 gap-2 mt-2"><div class="border border-gray-300 dark:border-zinc-700 p-2 text-center bg-gray-50 dark:bg-[#141417] rounded-xl"><div class="text-[10px] font-bold text-gray-400 dark:text-zinc-500 mb-1 uppercase">NMT</div><img src="${subImgs[0]}" class="h-36 w-full object-fill rounded-lg border border-gray-200 dark:border-zinc-800"></div><div class="border border-gray-300 dark:border-zinc-700 p-2 text-center bg-gray-50 dark:bg-[#141417] rounded-xl"><div class="text-[10px] font-bold text-gray-400 dark:text-zinc-500 mb-1 uppercase">AW</div><img src="${subImgs[1]}" class="h-36 w-full object-fill rounded-lg border border-gray-200 dark:border-zinc-800"></div><div class="border border-gray-300 dark:border-zinc-700 p-2 text-center bg-gray-50 dark:bg-[#141417] rounded-xl"><div class="text-[10px] font-bold text-gray-400 dark:text-zinc-500 mb-1 uppercase">VENDOR</div><img src="${subImgs[2]}" class="h-36 w-full object-fill rounded-lg border border-gray-200 dark:border-zinc-800"></div></div>`; } 
        else { structuralGridHtml = `<div class="grid grid-cols-2 gap-2 mt-2">${subImgs.map((img, index) => `<div class="border border-gray-300 dark:border-zinc-700 p-2 text-center bg-gray-50 dark:bg-[#141417] rounded-xl"><div class="text-[10px] font-bold text-gray-400 dark:text-zinc-500 mb-1 uppercase">IMAGE ${index+2}</div><img src="${img}" class="h-36 w-full object-fill rounded-lg border border-gray-200 dark:border-zinc-800"></div>`).join('')}</div>`; }

        listContainer.innerHTML = `<div class="w-full bg-white dark:bg-[#1a1a1e] p-2 rounded-2xl flex flex-col gap-3"><div class="flex items-start justify-between bg-gray-50 dark:bg-[#25252a] p-3 rounded-xl border border-gray-200/60 dark:border-zinc-800"><div class="flex flex-col truncate"><span class="text-xs font-bold text-gray-900 dark:text-white truncate">📄 File: ${currentMatch.fileName}</span><span class="text-[11px] text-gray-400 dark:text-zinc-500 font-semibold mt-0.5">${r.obsId} / ${r.jobId} / ${r.circle} / ${r.partner}</span></div><div class="flex gap-1.5 shrink-0 ml-2"><span class="px-2 py-0.5 rounded text-[9px] uppercase font-bold ${badgeStyle}">${badgeText}</span><span class="px-2 py-0.5 rounded text-[9px] uppercase font-bold ${categoryBadgeColor}">${r.category}</span></div></div><div class="border border-gray-200 dark:border-zinc-800 p-3 rounded-2xl bg-white dark:bg-[#25252a] shadow-sm"><div class="text-xs font-bold p-2.5 bg-gray-100 dark:bg-[#141417] border border-gray-200 dark:border-zinc-800 rounded-xl text-gray-800 dark:text-gray-200 mb-3 font-mono">OBS Task No. : ${r.obsId} / ${r.jobId} / ${r.circle} / ${r.partner}</div>${activeImgs[0] ? `<div class="border border-gray-200 dark:border-zinc-800 p-2 text-center rounded-xl bg-gray-50 dark:bg-[#141417]"><div class="text-[10px] font-bold text-gray-400 dark:text-zinc-500 mb-1 uppercase">OBSERVATION BASE LAYER</div><img src="${activeImgs[0]}" class="h-28 w-full object-fill rounded-lg border border-gray-200 dark:border-zinc-800 shadow-sm"></div>` : ''}${structuralGridHtml}</div><div class="flex items-center justify-between bg-gray-50 dark:bg-[#25252a] p-2.5 rounded-xl border border-gray-200 dark:border-zinc-800"><button onclick="navigateSearchResult(-1)" class="bg-white dark:bg-[#3a3a42] text-gray-800 dark:text-white px-3 py-1.5 rounded-lg text-xs font-bold shadow-sm hover:bg-[#007AFF] hover:text-white dark:hover:bg-[#007AFF] transition-all disabled:opacity-30 whitespace-nowrap" ${searchCurrentIndex === 0 ? 'disabled' : ''}>◀ Back</button><span class="text-xs font-bold text-gray-500 dark:text-zinc-400">Record ${searchCurrentIndex + 1} of ${searchResults.length}</span><button onclick="navigateSearchResult(1)" class="bg-white dark:bg-[#3a3a42] text-gray-800 dark:text-white px-3 py-1.5 rounded-lg text-xs font-bold shadow-sm hover:bg-[#007AFF] hover:text-white dark:hover:bg-[#007AFF] transition-all disabled:opacity-30 whitespace-nowrap" ${searchCurrentIndex === searchResults.length - 1 ? 'disabled' : ''}>Next ▶</button></div><div class="grid grid-cols-2 gap-3 pt-2"><button onclick="importSelectedSearchResult()" class="w-full bg-[#34C759] text-white text-xs font-bold py-3.5 rounded-xl hover:bg-emerald-600 transition-all active:scale-[0.99] shadow-md">✨ Select / Import</button><button onclick="closeDuplicateModal()" class="w-full bg-gray-100 dark:bg-zinc-800 text-gray-600 dark:text-zinc-300 text-xs font-bold py-3.5 rounded-xl hover:bg-gray-200 dark:hover:bg-zinc-700 transition-all">Cancel</button></div></div>`;
    }

    function navigateSearchResult(direction) {
        searchCurrentIndex += direction; if (searchCurrentIndex < 0) searchCurrentIndex = 0; if (searchCurrentIndex >= searchResults.length) searchCurrentIndex = searchResults.length - 1; renderPaginatedSearchResult();
    }

    function importSelectedSearchResult() {
        if (!searchResults.length) return; loadUniversalTargetMatch(searchResults[searchCurrentIndex]); closeDuplicateModal(); showToast("🎯 Record imported into workspace container", "#34C759");
    }

    function loadUniversalTargetMatch(match) {
        const r = match.data;
        if (match.source === 'CURRENT_WORK_FILE') { navIndex = match.index; document.getElementById('nav-label').innerText = `Rec ${navIndex+1}/${records.length}`; document.getElementById('update-btn').classList.remove('hidden'); document.getElementById('delete-page-btn').classList.remove('hidden'); } 
        else { navIndex = -1; document.getElementById('nav-label').innerText = `Linked File`; document.getElementById('update-btn').classList.add('hidden'); document.getElementById('delete-page-btn').classList.add('hidden'); }
        document.getElementById('obsId').value = r.obsId || ''; document.getElementById('jobId').value = r.jobId || ''; document.getElementById('circle').value = r.circle || ''; document.getElementById('partner').value = r.partner || '';
        const fields = ['obsId', 'jobId', 'circle', 'partner']; fields.forEach(id => document.getElementById(id).classList.remove('input-error'));
        currentImgs = r.imgs && Array.isArray(r.imgs) ? [...r.imgs] : [null, null, null, null]; renderSlots(); invalidateClearBackup(); updateCategoryBadge();
    }

    // =======================================================
    // 7. IMAGE HANDLING & UI RENDER Logic
    // =======================================================

    function updateCategoryBadge() {
        if (navIndex === -1 || !records[navIndex]) { document.getElementById('nav-label').innerText = "DRAFT"; }
        const cardWip = document.getElementById('wrapper-wip');
        const cardHold = document.getElementById('wrapper-hold');
        const cardClosed = document.getElementById('wrapper-closed');
        const cardValidation = document.getElementById('wrapper-validation');
        const baseClass = "bg-white dark:bg-[#3a3a42] px-2.5 py-1.5 rounded-lg flex items-center gap-1 border border-transparent transition-all duration-300";
        if(cardWip) cardWip.className = baseClass;
        if(cardHold) cardHold.className = baseClass;
        if(cardClosed) cardClosed.className = baseClass;
        if(cardValidation) cardValidation.className = baseClass;

        if (navIndex !== -1 && records[navIndex]) {
            const cat = records[navIndex].category.toUpperCase();
            if (cat.includes('WIP')) {
                if(cardWip) cardWip.className = "bg-blue-50/50 dark:bg-[#3a3a42] px-2.5 py-1.5 rounded-lg flex items-center gap-1 border-2 border-[#007AFF] shadow-sm transition-all duration-300 scale-105";
            } else if (cat.includes('HOLD')) {
                if(cardHold) cardHold.className = "bg-pink-50/50 dark:bg-[#3a3a42] px-2.5 py-1.5 rounded-lg flex items-center gap-1 border-2 border-[#FF2D55] shadow-sm transition-all duration-300 scale-105";
            } else if (cat.includes('CLOSED') && !cat.includes('VALIDATION')) {
                if(cardClosed) cardClosed.className = "bg-emerald-50/50 dark:bg-[#3a3a42] px-2.5 py-1.5 rounded-lg flex items-center gap-1 border-2 border-[#34C759] shadow-sm transition-all duration-300 scale-105";
            } else if (cat.includes('VALIDATION')) {
                if(cardValidation) cardValidation.className = "bg-amber-50/50 dark:bg-[#3a3a42] px-2.5 py-1.5 rounded-lg flex items-center gap-1 border-2 border-[#FF9500] shadow-sm transition-all duration-300 scale-105";
            }
        }
    }

    function renderSlots() {
        const grid = document.getElementById('image-grid');
        grid.innerHTML = currentImgs.map((src, i) => {
            const colClass = i === 0 ? 'col-span-3 first-slot' : 'col-span-1';
            let innerHtml = "";
            if (src) {
                // NEW: Open Image In New Tab functionality attached to img
                innerHtml += `<img src="${src}" onclick="event.stopPropagation();openImageInNewTab(${i})"><div class="delete-img" onclick="event.stopPropagation();removeImg(${i})" title="Delete Image">✕</div>`;
                if (i > 0) { innerHtml += `<div class="lens-btn" onclick="event.stopPropagation();cropAndOpenLens(${i})">🔍 Google Lens Search</div>`; }
            } else { innerHtml += `<span>Tap to insert <br><b class="text-gray-700 dark:text-zinc-300 font-semibold">${slotNames[i]}</b></span>`; }
            return `<div id="slot-${i}" class="image-slot ${colClass}" onclick="pasteImg(${i})">${innerHtml}</div>`;
        }).join('');
    }

    // NEW FUNCTION: Open original Image in New Tab
    function openImageInNewTab(index) {
        const src = currentImgs[index]; 
        if (!src) return;
        const win = window.open('', '_blank');
        win.document.write('<html><head><title>Original Image Preview</title></head><body style="margin:0; background:#0e0e11; display:flex; justify-content:center; align-items:center; height:100vh;"><img src="' + src + '" style="max-width:100%; max-height:100%; object-fit:contain;"></body></html>');
        win.document.close();
    }

    function cropAndOpenLens(index) {
        const base64Str = currentImgs[index]; if (!base64Str) return;
        showToast("Cropping Lat/Long Matrix Layer...", "#007AFF"); const targetWindow = window.open('about:blank', '_blank');
        const img = new Image(); img.src = base64Str;
        img.onload = () => {
            const canvas = document.createElement('canvas'); const ctx = canvas.getContext('2d');
            const cropWidth = img.width * 0.90; const cropHeight = img.height * 0.20;
            const cropX = img.width - cropWidth; const cropY = img.height - cropHeight;
            canvas.width = cropWidth; canvas.height = cropHeight;
            ctx.drawImage(img, cropX, cropY, cropWidth, cropHeight, 0, 0, cropWidth, cropHeight);
            
            canvas.toBlob((blob) => {
                if (!blob) { showToast("Crop Generation Error!", "#FF3B30"); targetWindow.close(); return; }
                try {
                    showToast("🚀 Handshaking to Google Lens Hub...", "#34C759");
                    const form = document.createElement('form'); form.method = 'POST'; form.action = 'https://lens.google.com/v3/upload'; form.target = targetWindow.name || '_blank'; form.enctype = 'multipart/form-data';
                    const fileInput = document.createElement('input'); fileInput.type = 'file'; fileInput.name = 'encoded_image';
                    const dataTransfer = new DataTransfer(); const file = new File([blob], 'cropped_corner.jpg', { type: 'image/jpeg' });
                    dataTransfer.items.add(file); fileInput.files = dataTransfer.files;
                    form.appendChild(fileInput); document.body.appendChild(form); form.submit(); form.remove(); 
                } catch (err) { showToast("❌ Auto Upload Failure", "#FF3B30"); targetWindow.location.href = "https://lens.google.com/search?p="; }
            }, 'image/jpeg', 0.9);
        };
    }

    // Heavy Image Size reduction with WebP 0.6 + 800px width limit
    async function compressImage(base64Str) {
        return new Promise((resolve) => {
            if(!base64Str) return resolve(null);
            const img = new Image(); 
            img.src = base64Str;
            img.onload = () => {
                const canvas = document.createElement('canvas'); 
                const MAX_WIDTH = 800; // Optimal max width for preserving visual clarity 
                let scaleSize = MAX_WIDTH / img.width;
                if (scaleSize > 1) scaleSize = 1; 
                
                canvas.width = img.width * scaleSize; 
                canvas.height = img.height * scaleSize;
                
                const ctx = canvas.getContext('2d'); 
                ctx.fillStyle = '#FFFFFF'; 
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.drawImage(img, 0, 0, canvas.width, canvas.height); 
                
                // Extremely efficient size saving format
                resolve(canvas.toDataURL('image/webp', 0.6));
            };
            img.onerror = () => resolve(base64Str); 
        });
    }

    async function pasteImg(i) {
        try {
            const items = await navigator.clipboard.read();
            for (let item of items) {
                if (item.types.some(t => t.startsWith('image/'))) {
                    const blob = await item.getType(item.types.find(t => t.startsWith('image/')));
                    const reader = new FileReader();
                    reader.onload = async (e) => { 
                        const compressed = await compressImage(e.target.result); currentImgs[i] = compressed; 
                        document.getElementById(`slot-${i}`).classList.remove('input-error');
                        if(navIndex === -1) draftData.imgs[i] = compressed; 
                        renderSlots(); invalidateClearBackup();
                        if (i === 1 || i === 3) { searchAndCopyExcelCoordinates(false); }
                    };
                    reader.readAsDataURL(blob);
                }
            }
        } catch(e) { showToast("Clipboard system link locked", "#FF3B30"); }
    }

    function removeImg(i) { currentImgs[i] = null; if(navIndex === -1) draftData.imgs[i] = null; renderSlots(); invalidateClearBackup(); }
    
    // =======================================================
    // 8. FORMS, NAVIGATIONS & FILE SYNC
    // =======================================================

    function openModal(task) { 
        modalTask = task; 
        document.getElementById('action-modal').classList.remove('hidden'); 
        
        const excelBtn = document.getElementById('excel-export-btn');
        const currentExportBtn = document.getElementById('export-current-btn');

        // NEW: Modals Visibility Logic for Export
        if(task === 'export') {
            if(globalWorkbook) excelBtn.classList.remove('hidden');
            else excelBtn.classList.add('hidden');
            
            currentExportBtn.classList.remove('hidden');
        } else {
            if(excelBtn) excelBtn.classList.add('hidden');
            if(currentExportBtn) currentExportBtn.classList.add('hidden');
        }
    }
    
    function closeModal() { document.getElementById('action-modal').classList.add('hidden'); }
    function handleAction(cat) { closeModal(); if (modalTask === 'export') exportReport(cat); else saveRecord(cat); }

    function validateData(cat) {
        let missingFields = [];
        const fields = [{ id: 'obsId', name: 'OBS ID' }, { id: 'jobId', name: 'JOB ID' }, { id: 'circle', name: 'CIRCLE' }, { id: 'partner', name: 'PARTNER' }];
        fields.forEach(f => {
            const el = document.getElementById(f.id);
            if (!el.value.trim()) { el.classList.add('input-error'); missingFields.push(f.name); } else { el.classList.remove('input-error'); }
        });
        document.querySelectorAll('.image-slot').forEach(slot => slot.classList.remove('input-error'));
        
        if (cat === 'Vendor WIP' || cat === 'Hold') {
            let validImgCount = currentImgs.filter(x => x).length;
            if (validImgCount < 3) { [0, 1, 2].forEach(i => { if(!currentImgs[i]) { document.getElementById(`slot-${i}`).classList.add('input-error'); missingFields.push(slotNames[i]); } }); }
        } else if (cat === 'ZQH Closed') {
            [0, 1, 2].forEach(i => { if(!currentImgs[i]) { document.getElementById(`slot-${i}`).classList.add('input-error'); missingFields.push(slotNames[i]); } });
            if(currentImgs[3]){ document.getElementById('slot-3').classList.add('input-error'); showToast("⚠️ ZQH Closed criteria requires exactly 3 images max!", "#FF3B30"); return false; }
        } else if (cat === 'ZQH Validation Closed') {
            [0, 1, 2, 3].forEach(i => { if(!currentImgs[i]) { document.getElementById(`slot-${i}`).classList.add('input-error'); missingFields.push(slotNames[i]); } });
        }
        if (missingFields.length > 0) { showToast(`⚠️ Missing: ${missingFields.join(', ')}`, "#FF3B30"); return false; }
        return true;
    }

    async function saveRecord(cat) {
        if (!validateData(cat)) return;
        if (!currentActiveFileHandle) { showToast("Database Storage Core Offline!", "#FF3B30"); return; }
        const data = { category: cat, obsId: document.getElementById('obsId').value, jobId: document.getElementById('jobId').value, circle: document.getElementById('circle').value, partner: document.getElementById('partner').value, imgs: [...currentImgs] };
        records.push(data); await syncToActiveFile(); showToast(`🟢 Saved to ${cat}`, "#34C759"); invalidateClearBackup();
        if(modalTask === 'continue') { 
            document.getElementById('obsId').value = ''; currentImgs = [null, null, null, null]; 
            const fields = ['obsId', 'jobId', 'circle', 'partner']; fields.forEach(id => document.getElementById(id).classList.remove('input-error'));
            renderSlots(); 
        } else { resetDraft(); }
        updateCategoryBadge();
    }

    function updateCounts() { 
        document.getElementById('count-wip').innerText = records.filter(r => r.category === 'Vendor WIP').length; 
        if (document.getElementById('count-hold')) { document.getElementById('count-hold').innerText = records.filter(r => r.category === 'Hold').length; }
        document.getElementById('count-closed').innerText = records.filter(r => r.category === 'ZQH Closed').length; 
        if (document.getElementById('count-pending')) { document.getElementById('count-pending').innerText = records.filter(r => r.category === 'ZQH Validation Closed').length; }
        document.getElementById('count-total').innerText = records.length;
    }

    function toggleClearOrUndo() {
        const btn = document.getElementById('clear-undo-btn');
        if (clearPageBackup !== null) {
            document.getElementById('obsId').value = clearPageBackup.obsId; document.getElementById('jobId').value = clearPageBackup.jobId; document.getElementById('circle').value = clearPageBackup.circle; document.getElementById('partner').value = clearPageBackup.partner;
            currentImgs = [...clearPageBackup.imgs]; if(navIndex === -1) { draftData = { ...clearPageBackup }; }
            renderSlots(); clearPageBackup = null; btn.innerHTML = "🧹"; btn.className = "bg-[#FF3B30]/10 text-[#FF3B30] border border-[#FF3B30]/20 px-3 py-2 rounded-xl text-xs font-bold hover:bg-[#FF3B30] hover:text-white transition-all active:scale-95"; showToast("↩️ Content Restored Automatically!", "#34C759");
        } else {
            clearPageBackup = { obsId: document.getElementById('obsId').value, jobId: document.getElementById('jobId').value, circle: document.getElementById('circle').value, partner: document.getElementById('partner').value, imgs: [...currentImgs] };
            document.getElementById('obsId').value = ''; document.getElementById('jobId').value = ''; document.getElementById('circle').value = ''; document.getElementById('partner').value = ''; currentImgs = [null, null, null, null];
            if(navIndex === -1) { draftData = { obsId: '', jobId: '', circle: '', partner: '', imgs: [null, null, null, null] }; }
            const fields = ['obsId', 'jobId', 'circle', 'partner']; fields.forEach(id => document.getElementById(id).classList.remove('input-error'));
            renderSlots(); btn.innerHTML = "↩️"; btn.className = "bg-[#FF9500]/10 text-[#FF9500] border border-[#FF9500]/20 px-3 py-2 rounded-xl text-xs font-bold hover:bg-[#FF9500] hover:text-white transition-all active:scale-95"; showToast("🧹 Page Cleared! Click again to Undo.", "#FF9500");
        }
        updateCategoryBadge();
    }

    function invalidateClearBackup() {
        if (clearPageBackup !== null) { clearPageBackup = null; const btn = document.getElementById('clear-undo-btn'); btn.innerHTML = "🧹"; btn.className = "bg-[#FF3B30]/10 text-[#FF3B30] border border-[#FF3B30]/20 px-3 py-2 rounded-xl text-xs font-bold hover:bg-[#FF3B30] hover:text-white transition-all active:scale-95"; }
    }

    function navigate(dir) {
        if(navIndex === -1) { draftData = { obsId: document.getElementById('obsId').value, jobId: document.getElementById('jobId').value, circle: document.getElementById('circle').value, partner: document.getElementById('partner').value, imgs: [...currentImgs] }; }
        if(!records.length) { showToast("No records available to navigate!", "#FF9500"); return; }
        if (dir === -1) { 
            if (navIndex === -1) { navIndex = records.length - 1; } else if (navIndex > 0) { navIndex--; } else { showToast("Reached beginning of saved records!", "#007AFF"); return; }
        } else if (dir === 1) { 
            if (navIndex === -1) { return; } else if (navIndex < records.length - 1) { navIndex++; } else { navIndex = -1; restoreDraft(); return; }
        }
        const r = records[navIndex]; if(!r) return;
        document.getElementById('nav-label').innerText = `Rec ${navIndex+1}/${records.length}`; document.getElementById('obsId').value = r.obsId || ''; document.getElementById('jobId').value = r.jobId || ''; document.getElementById('circle').value = r.circle || ''; document.getElementById('partner').value = r.partner || '';
        const fields = ['obsId', 'jobId', 'circle', 'partner']; fields.forEach(id => document.getElementById(id).classList.remove('input-error'));
        currentImgs = [...r.imgs]; renderSlots(); invalidateClearBackup(); document.getElementById('update-btn').classList.remove('hidden'); document.getElementById('delete-page-btn').classList.remove('hidden'); updateCategoryBadge();
    }

    function restoreDraft() {
        document.getElementById('nav-label').innerText = "DRAFT"; document.getElementById('obsId').value = draftData.obsId; document.getElementById('jobId').value = draftData.jobId; document.getElementById('circle').value = draftData.circle; document.getElementById('partner').value = draftData.partner;
        const fields = ['obsId', 'jobId', 'circle', 'partner']; fields.forEach(id => document.getElementById(id).classList.remove('input-error'));
        currentImgs = [...draftData.imgs]; document.getElementById('update-btn').classList.add('hidden'); document.getElementById('delete-page-btn').classList.add('hidden'); renderSlots(); invalidateClearBackup(); updateCategoryBadge();
    }

    async function updateRecord() {
        if(navIndex === -1) return; const cat = records[navIndex].category; if (!validateData(cat)) return;
        records[navIndex] = { ...records[navIndex], obsId: document.getElementById('obsId').value, jobId: document.getElementById('jobId').value, circle: document.getElementById('circle').value, partner: document.getElementById('partner').value, imgs: [...currentImgs] };
        await syncToActiveFile(); showToast(`✅ Mutated & Updated in ${cat}`, "#FF9500"); invalidateClearBackup(); updateCategoryBadge();
    }

    async function deleteCurrentPage() {
        if (navIndex === -1) return;
        if (confirm(`Delete Record ${navIndex + 1}?`)) {
            records.splice(navIndex, 1); await syncToActiveFile();
            if (records.length === 0) { resetDraft(); } else { if (navIndex >= records.length) navIndex = records.length - 1; loadUniversalTargetMatch({ source: 'CURRENT_WORK_FILE', index: navIndex, data: records[navIndex] }); }
            showToast("🗑️ Node Purged from Instance", "#ff453a"); invalidateClearBackup(); updateCategoryBadge();
        }
    }

    function resetDraft() { 
        document.querySelectorAll('input').forEach(i => { i.value = ''; i.classList.remove('input-error'); }); currentImgs = [null, null, null, null];
        draftData = { obsId: '', jobId: '', circle: '', partner: '', imgs: [null, null, null, null] }; navIndex = -1; renderSlots(); restoreDraft(); invalidateClearBackup(); updateCategoryBadge();
    }

    // Helper: Converts older WebP backups into JPEG specifically for MS Word copy support
    function forceJpeg(base64Str) {
        return new Promise((resolve) => {
            if (!base64Str || !base64Str.startsWith('data:image/webp')) return resolve(base64Str);
            const img = new Image(); 
            img.onload = () => {
                const canvas = document.createElement('canvas'); 
                canvas.width = img.width; canvas.height = img.height;
                const ctx = canvas.getContext('2d'); 
                ctx.fillStyle = '#FFFFFF'; 
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.drawImage(img, 0, 0); 
                resolve(canvas.toDataURL('image/jpeg', 0.8));
            };
            img.onerror = () => resolve(base64Str);
            img.src = base64Str;
        });
    }

    // NEW FUNCTION: Export single current page
    async function exportCurrentPage() {
        closeModal();
        const obsId = document.getElementById('obsId').value.trim();
        if (!obsId) {
            showToast("Nothing to export. Form is empty.", "#FF3B30");
            return;
        }

        const currentRecord = {
            obsId: obsId,
            jobId: document.getElementById('jobId').value,
            circle: document.getElementById('circle').value,
            partner: document.getElementById('partner').value,
            imgs: [...currentImgs]
        };

        await processExport([currentRecord]);
    }

    async function exportReport(cat) {
        const data = records.filter(r => r.category === cat); 
        if(!data.length) { showToast("Empty stack layer parameters", "#FF3B30"); return; }
        await processExport(data);
    }
    
    // REFACTORED EXPORT CORE ENGINE
    async function processExport(data) {
        const pWin = window.open('', '_blank');
        if(!pWin) { showToast("Popup blocked! Please allow popups.", "#FF3B30"); return; }
        
        pWin.document.write('<html style="background:#222;color:#fff;font-family:sans-serif;text-align:center;padding:50px;"><body><h2>⏳ Formatting Images for MS Word Compatibility...</h2></body></html>');
        
        let html = '<div id="report-container" style="font-family: Calibri;">';
        
        for (let r of data) {
            const rawActiveImgs = r.imgs.filter(x => x);
            const activeImgs = [];
            // Dynamically transpile WebP back to JPEG
            for (let img of rawActiveImgs) { activeImgs.push(await forceJpeg(img)); }
            
            const subImgs = activeImgs.slice(1); let subTableHtml = "";
            if (activeImgs.length === 3) {
                subTableHtml = `<tr><td width="280" style="border: 1pt solid black; padding: 10px; text-align: center;"><div class="label-box" style="margin-bottom: 8px;">NMT CAPTURED</div><img src="${subImgs[0]}" width="280" height="520" style="width:2.9in;height:5.4in;border:1pt solid black; display:block; margin:0 auto; object-fit:fill;"></td><td width="280" style="border: 1pt solid black; padding: 10px; text-align: center;"><div class="label-box" style="margin-bottom: 8px;">CAPTURED IN AW</div><img src="${subImgs[1]}" width="280" height="520" style="width:2.9in;height:5.4in;border:1pt solid black; display:block; margin:0 auto; object-fit:fill;"></td></tr>`;
            } else if (activeImgs.length === 4) {
                subTableHtml = `<tr><td width="185" style="border: 1pt solid black; padding: 8px; text-align: center;"><div class="label-box" style="margin-bottom: 8px;">NMT CAPTURED</div><img src="${subImgs[0]}" width="185" height="510" style="width:1.9in;height:5.3in;border:1pt solid black; display:block; margin:0 auto; object-fit:fill;"></td><td width="185" style="border: 1pt solid black; padding: 8px; text-align: center;"><div class="label-box" style="margin-bottom: 8px;">CAPTURED IN AW</div><img src="${subImgs[1]}" width="185" height="510" style="width:1.9in;height:5.3in;border:1pt solid black; display:block; margin:0 auto; object-fit:fill;"></td><td width="185" style="border: 1pt solid black; padding: 8px; text-align: center;"><div class="label-box" style="margin-bottom: 8px;">VENDOR CAPTURED</div><img src="${subImgs[2]}" width="185" height="510" style="width:1.9in;height:5.3in;border:1pt solid black; display:block; margin:0 auto; object-fit:fill;"></td></tr>`;
            }
            html += `<div class="page-break" style="width: 8.27in; padding: 0.3in; margin: 20px auto; background: white;"><table width="620" style="border-collapse: separate; border-spacing: 0 12px; margin: 0 auto;"><tr><td colspan="${activeImgs.length === 3 ? 2 : 3}" style="border: 1.5pt solid black; padding: 12px; background: #f8f8f8; text-align: left;"><span style="font-weight: bold; font-size: 12pt;">OBS Task No. :</span> <span style="font-size: 10pt; font-weight: bold;">${r.obsId} / ${r.jobId} / ${r.circle} / ${r.partner}</span></td></tr><tr><td colspan="${activeImgs.length === 3 ? 2 : 3}" style="border: 1pt solid black; padding: 15px; text-align: center;"><img src="${activeImgs[0]}" width="580" height="200" style="width:6in;height:2.1in;border:1pt solid black; display:block; margin:0 auto; object-fit:fill;"></td></tr>${subTableHtml}</table></div>`;
        }
        html += '</div>';
        
        pWin.document.open();
        pWin.document.write(`<html><head><style>body { background: #555; padding: 20px; } .label-box { font-weight: bold; font-size: 10pt; background: #f0f0f0; border: 1pt solid black; padding: 4pt; text-align: center; font-family: Calibri; } .copy-btn { position: fixed; top: 20px; right: 20px; background: #007AFF; color: white; border: none; padding: 12px 24px; border-radius: 8px; font-weight: bold; cursor: pointer; z-index: 1000; box-shadow: 0 4px 12px rgba(0,0,0,0.3); transition: all 0.2s; } .copy-btn:active { transform: scale(0.95); }</style></head><body><button class="copy-btn" onclick="copyAll()">📋 Copy All for Word</button>${html}<script>function copyAll(){const range=document.createRange();range.selectNode(document.getElementById('report-container'));window.getSelection().removeAllRanges();window.getSelection().addRange(range);document.execCommand('copy');alert('✅ Copied Successfully! Now Paste (Ctrl+V) in MS Word.');}<\/script></body></html>`);
        pWin.document.close();
    }
    
    // Application Bootstrap
    window.onload = init;
