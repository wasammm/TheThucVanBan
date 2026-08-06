document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('file-input');
    const dropZone = document.getElementById('drop-zone');
    const fileNameDisplay = document.getElementById('file-name');
    const btnFormat = document.getElementById('btn-format');
    const docTypeSelect = document.getElementById('doc-type');
    const statusMsg = document.getElementById('status-message');

    let currentFile = null;

    // Xử lý sự kiện chọn file
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            currentFile = e.target.files[0];
            updateFileInfo();
        }
    });

    // Kéo thả file
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
        showMessage("", "");
    }

    function showMessage(msg, type) {
        statusMsg.textContent = msg;
        statusMsg.className = 'status-message ' + type;
    }

    // Xử lý khi bấm nút định dạng
    btnFormat.addEventListener('click', async () => {
        if (!currentFile) return;
        
        try {
            btnFormat.disabled = true;
            btnFormat.textContent = "Đang xử lý...";
            showMessage("Đang đọc file...", "");

            // 1. Đọc file zip (docx)
            const zip = new JSZip();
            const loadedZip = await zip.loadAsync(currentFile);
            
            // 2. Chỉnh sửa styles.xml (Định dạng chữ, lùi đầu dòng, căn lề, khoảng cách dòng)
            let stylesXml = await loadedZip.file("word/styles.xml").async("string");
            stylesXml = applyGlobalStyles(stylesXml);
            loadedZip.file("word/styles.xml", stylesXml);

            // 3. Chỉnh sửa document.xml (Khổ giấy, lề trang, tiêu đề TỜ TRÌNH, BÁO CÁO)
            let docXml = await loadedZip.file("word/document.xml").async("string");
            
            let selectedType = docTypeSelect.value;
            // Nếu AUTO, thử tìm từ khóa
            if(selectedType === "AUTO") {
                if(docXml.includes("TỜ TRÌNH") || docXml.includes("Tờ trình") || docXml.includes("tờ trình")) selectedType = "TO_TRINH";
                else if(docXml.includes("BÁO CÁO") || docXml.includes("Báo cáo") || docXml.includes("báo cáo")) selectedType = "BAO_CAO";
                else selectedType = "CONG_VAN";
            }

            docXml = applyPageSetup(docXml);
            docXml = applySpecificRules(docXml, selectedType);
            loadedZip.file("word/document.xml", docXml);

            // 4. Đóng gói lại file và tải xuống
            showMessage("Đang tạo file chuẩn...", "");
            const outBlob = await loadedZip.generateAsync({ type: "blob" });
            
            const newName = currentFile.name.replace('.docx', '_DaDinhDang.docx');
            saveAs(outBlob, newName);

            showMessage(`Thành công! Đã tải xuống file: ${newName}`, "success");
        } catch (error) {
            console.error(error);
            showMessage("Có lỗi xảy ra trong quá trình xử lý file.", "error");
        } finally {
            btnFormat.disabled = false;
            btnFormat.textContent = "Định dạng & Tải xuống";
        }
    });

    // --- Các hàm Core Logic xử lý XML của file Docx ---

    function applyGlobalStyles(xml) {
        // Can thiệp vào Defaults của styles.xml
        // 1. Font Times New Roman, Size 28 (14pt)
        const fontDefault = `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>`;
        
        // 2. Canh đều hai bên (both), lùi đầu dòng 1.27cm (720 twips), khoảng cách dòng 1.5 (360 twips)
        const paraDefault = `<w:pPr><w:jc w:val="both"/><w:spacing w:line="360" w:lineRule="auto"/><w:ind w:firstLine="720"/></w:pPr>`;

        // Chèn vào docDefaults
        if (xml.includes('<w:docDefaults>')) {
            xml = xml.replace(/<w:rPrDefault>.*?<\/w:rPrDefault>/, `<w:rPrDefault>${fontDefault}</w:rPrDefault>`);
            xml = xml.replace(/<w:pPrDefault>.*?<\/w:pPrDefault>/, `<w:pPrDefault>${paraDefault}</w:pPrDefault>`);
        } else {
            // Nếu chưa có docDefaults, thêm vào đầu
            const defaults = `<w:docDefaults><w:rPrDefault>${fontDefault}</w:rPrDefault><w:pPrDefault>${paraDefault}</w:pPrDefault></w:docDefaults>`;
            xml = xml.replace(/<w:styles[^>]*>/, `$&${defaults}`);
        }
        return xml;
    }

    function applyPageSetup(xml) {
        // Thiết lập khổ A4 (w:w="11906" w:h="16838")
        xml = xml.replace(/<w:pgSz[^>]*\/>/g, '<w:pgSz w:w="11906" w:h="16838"/>');
        
        // Thiết lập lề: Trên 20mm (1134), Dưới 20mm (1134), Trái 30mm (1701), Phải 15mm (850)
        xml = xml.replace(/<w:pgMar[^>]*\/>/g, '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>');
        return xml;
    }

    function applySpecificRules(xml, docType) {
        // Hàm này quét qua các thẻ paragraph <w:p>
        // Vì làm việc với XML bằng regex rất phức tạp, ta dùng thao tác string đơn giản để nhận diện chữ
        
        // Nhận diện "BÁO CÁO" hoặc "TỜ TRÌNH" để in hoa, in đậm, canh giữa, KHÔNG lùi đầu dòng
        if (docType === "BAO_CAO" || docType === "TO_TRINH") {
            const keyword1 = docType === "BAO_CAO" ? "BÁO CÁO" : "TỜ TRÌNH";
            const keyword2 = docType === "BAO_CAO" ? "Báo cáo" : "Tờ trình";
            
            let paragraphs = xml.split('</w:p>');
            for (let i = 0; i < paragraphs.length; i++) {
                if (paragraphs[i].includes(keyword1) || paragraphs[i].includes(keyword2)) {
                    // Nếu là đoạn chứa tiêu đề, ép định dạng
                    // Thêm thuộc tính pPr: canh giữa (center), không lùi đầu dòng (ind left=0 firstLine=0)
                    let pPr = `<w:pPr><w:jc w:val="center"/><w:ind w:left="0" w:firstLine="0"/></w:pPr>`;
                    
                    // Thêm thuộc tính rPr: in đậm, in hoa
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
