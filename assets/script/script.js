document.getElementById('exportButton').addEventListener('click', function() {
    const text = document.getElementById('inputText').value.trim();
    const termSeparator = document.getElementById('termSeparator').value || ':';
    const pairSeparator = document.getElementById('pairSeparator').value || '\n';
    const fileName = document.getElementById('fileName').value || 'vocabulary.xlsx';
    const sheetName = document.getElementById('sheetName').value || 'Vocabulary';
    const appendOption = document.getElementById('appendOption').checked;
    const fileInput = document.getElementById('fileInput').files[0];

    // Xử lý dữ liệu đầu vào
    const pairs = text.split(new RegExp(pairSeparator, 'g'));
    const data = [];

    pairs.forEach(pair => {
        const trimmedPair = pair.trim();
        if (trimmedPair) {
            const splitIndex = trimmedPair.indexOf(termSeparator);
            if (splitIndex !== -1) {
                const term = trimmedPair.substring(0, splitIndex).trim();
                const definition = trimmedPair.substring(splitIndex + termSeparator.length).trim();
                if (term && definition) {
                    data.push({
                        Term: term,
                        Definition: definition
                    });
                }
            }
        }
    });

    if (appendOption && fileInput) {
        // Đọc file Excel hiện có
        const reader = new FileReader();
        reader.onload = function(event) {
            const wb = XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
            let ws = wb.Sheets[wb.SheetNames[0]];

            // Đọc dữ liệu từ sheet hiện tại
            const existingData = XLSX.utils.sheet_to_json(ws);
            const newData = existingData.concat(data);

            // Tạo sheet mới với dữ liệu kết hợp
            ws = XLSX.utils.json_to_sheet(newData);

            // Đặt chiều rộng cột
            ws['!cols'] = [{ wpx: 200 }, { wpx: 400 }]; // Điều chỉnh chiều rộng cột

            // Ghi vào file mới
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
            XLSX.writeFile(wb, fileName);
        };
        reader.readAsArrayBuffer(fileInput);
    } else {
        // Tạo workbook và worksheet mới
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(data);

        // Đặt tiêu đề cho cột
        ws['!cols'] = [{ wpx: 150 }, { wpx: 555 }]; // Điều chỉnh chiều rộng cột

        XLSX.utils.book_append_sheet(wb, ws, sheetName);

        // Ghi vào file
        XLSX.writeFile(wb, fileName);
    }
});
