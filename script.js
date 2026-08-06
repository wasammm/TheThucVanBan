document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('file-input');
    const dropZone = document.getElementById('drop-zone');
    const fileNameDisplay = document.getElementById('file-name');
    const btnFormat = document.getElementById('btn-format');
    const btnDownload = document.getElementById('btn-download');
    const docTypeSelect = document.getElementById('doc-type');
    const progressContainer = document.getElementById('progress-container');
    const progressList = document.getElementById('progress-list');
    const statusMsg = document.getElementById('status-message');

    let currentFile = null;
    let formattedBlob = null;
    let processedFileName = "";

    const stepsTemplate = [
        "Đọc và giải nén tệp Word (ZIP)",
        "Áp dụng quy định font chữ & khoảng cách đoạn",
        "Áp dụng cấu hình khổ giấy, lề trang & định dạng tiêu đề",
        "Đóng gói tệp tin định dạng mới"
    ];

    // Tạo hiệu ứng trễ để UI kịp render trạng thái
    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            currentFile = e.target.files[0];
            updateFileInfo();
        }
    });

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            let file = e.dataTransfer.files[0];
            if (file.name.endsWith('.docx')) {
                currentFile = file;
                updateFileInfo();
            } else {
                showMessage("Vui lòng chỉ tải lên file Word (.docx)", "error");
            }
        }
    });

    function updateFileInfo() {
        fileNameDisplay.textContent = `File đã chọn: ${currentFile.name}`;
        btnFormat.disabled = false;
        btnDownload.style.display = 'none';
        progressContainer.style.display = 'none';
        showMessage("", "");
    }

    function showMessage(msg, type) {
        statusMsg.textContent = msg;
        statusMsg.className = 'status-message ' + type;
    }

    function initProgressUI() {
        progressContainer.style.display = 'block';
        progressList.innerHTML = '';
        stepsTemplate.forEach((step, index) => {
            const li = document.createElement('li');
            li.id = `step-${index}`;
            li.textContent = step;
            progressList.appendChild(li);
        });
    }

    function setStepStatus(stepIndex, status) {
        const li = document.getElementById(`step-${stepIndex}`);
        if (li) {
            li.className = status; // active, success, error
        }
    }

    // NÚT 1: ĐỊNH DẠNG
    btnFormat.addEventListener('click', async () => {
        if (!currentFile) return;
        
        btnFormat.disabled = true;
        btnDownload.style.display = 'none';
        initProgressUI();
        showMessage("", "");

        let currentStep = 0;
        try {
            // Bước 0: Đọc tệp
            setStepStatus(currentStep, 'active');
            await sleep(500); 
            const zip = new JSZip();
            const loadedZip = await zip.loadAsync(currentFile);
            setStepStatus(currentStep, 'success');
            currentStep++;

            // Bước 1: Xử lý styles.xml
            setStepStatus(currentStep, 'active');
            await sleep(500);
            const stylesFile = loadedZip.file("word/styles.xml");
            if (!stylesFile) throw new Error("Không tìm thấy cấu trúc styles chuẩn trong file Word này.");
            let stylesXml = await stylesFile.async("string");
            stylesXml = applyGlobalStyles(stylesXml);
            loadedZip.file("word/styles.xml", stylesXml);
            setStepStatus(currentStep, 'success');
            currentStep++;

            // Bước 2: Xử lý document.xml
            setStepStatus(currentStep, 'active');
            await sleep(500);
            const docFile = loadedZip.file("word/document.xml");
            if (!docFile) throw new Error("Không tìm thấy nội dung văn bản chuẩn trong file Word này.");
            let docXml = await docFile.async("string");
            
            let selectedType = docTypeSelect.value;
            if(selectedType === "AUTO") {
                if(docXml.includes("TỜ TRÌNH") || docXml.includes("Tờ trình") || docXml.includes("tờ trình")) selectedType = "TO_TRINH";
                else if(docXml.includes("BÁO CÁO") || docXml.includes("Báo cáo") || docXml.includes("báo cáo")) selectedType = "BAO_CAO";
                else selectedType = "CONG_VAN";
            }

            docXml = applyPageSetup(docXml);
            docXml = applySpecificRules(docXml, selectedType);
            loadedZip.file("word/document.xml", docXml);
            setStepStatus(currentStep, 'success');
            currentStep++;

            // Bước 3: Đóng gói
            setStepStatus(currentStep, 'active');
            await sleep(500);
            formattedBlob = await loadedZip.generateAsync({ type: "blob" });
            processedFileName = currentFile.name.replace('.docx', '_DaDinhDang.docx');
            setStepStatus(currentStep, 'success');

            showMessage("Định dạng hoàn tất! Bạn có thể tải file về.", "success");
            btnFormat.textContent = "Định dạng lại";
            btnFormat.disabled = false;
            
            // Hiển thị nút tải xuống
            btnDownload.style.display = 'block';

        } catch (error) {
            console.error(error);
            setStepStatus(currentStep, 'error');
            showMessage(`Lỗi tại bước ${currentStep + 1}: ${error.message || 'Lỗi không xác định'}`, "error");
            btnFormat.disabled = false;
            btnFormat.textContent = "Thử lại";
        }
    });

    // NÚT 2: TẢI XUỐNG
    btnDownload.addEventListener('click', () => {
        if (formattedBlob && processedFileName) {
            saveAs(formattedBlob, processedFileName);
        }
    });

    // --- Các hàm Core Logic xử lý XML của file Docx ---

    function applyGlobalStyles(xml) {
        const fontDefault = `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>`;
        const paraDefault = `<w:pPr><w:jc w:val="both"/><w:spacing w:line="360" w:lineRule="auto"/><w:ind w:firstLine="720"/></w:pPr>`;

        if (xml.includes('<w:docDefaults>')) {
            xml = xml.replace(/<w:rPrDefault>.*?<\/w:rPrDefault>/, `<w:rPrDefault>${fontDefault}</w:rPrDefault>`);
            xml = xml.replace(/<w:pPrDefault>.*?<\/w:pPrDefault>/, `<w:pPrDefault>${paraDefault}</w:pPrDefault>`);
        } else {
            const defaults = `<w:docDefaults><w:rPrDefault>${fontDefault}</w:rPrDefault><w:pPrDefault>${paraDefault}</w:pPrDefault></w:docDefaults>`;
            xml = xml.replace(/<w:styles[^>]*>/, `$&${defaults}`);
        }
        return xml;
    }

    function applyPageSetup(xml) {
        xml = xml.replace(/<w:pgSz[^>]*\/>/g, '<w:pgSz w:w="11906" w:h="16838"/>');
        // w:right="1134" tương đương 20mm (20 * 56.7 twips)
        xml = xml.replace(/<w:pgMar[^>]*\/>/g, '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>');
        return xml;
    }

    function applySpecificRules(xml, docType) {
        if (docType === "BAO_CAO" || docType === "TO_TRINH") {
            const keyword1 = docType === "BAO_CAO" ? "BÁO CÁO" : "TỜ TRÌNH";
            const keyword2 = docType === "BAO_CAO" ? "Báo cáo" : "Tờ trình";
            
            let paragraphs = xml.split('</w:p>');
            for (let i = 0; i < paragraphs.length; i++) {
                if (paragraphs[i].includes(keyword1) || paragraphs[i].includes(keyword2)) {
                    let pPr = `<w:pPr><w:jc w:val="center"/><w:ind w:left="0" w:firstLine="0"/></w:pPr>`;
                    paragraphs[i] = paragraphs[i].replace(/<w:rPr>/g, '<w:rPr><w:b/><w:caps/>');
                    
                    if (paragraphs[i].includes('<w:pPr>')) {
                        paragraphs[i] = paragraphs[i].replace(/<w:pPr>.*?<\/w:pPr>/, pPr);
                    } else {
                        paragraphs[i] = paragraphs[i].replace(/<w:p[^>]*>/, `$&${pPr}`);
                    }
                }
            }
            xml = paragraphs.join('</w:p>');
        }
        return xml;
    }
});
