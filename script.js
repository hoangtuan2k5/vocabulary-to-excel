document.getElementById('exportButton').addEventListener('click', function() {
    // Lấy dữ liệu đầu vào
    const text = document.getElementById('inputText').value.trim();
    const termSeparator = document.getElementById('termSeparator').value || ':';
    const pairSeparator = document.getElementById('pairSeparator').value || '\n';

    // Xử lý dữ liệu đầu vào, thay thế các ký tự xuống dòng khác nhau
    const pairs = text.split(new RegExp(pairSeparator, 'g'));
    const data = [];

    pairs.forEach(pair => {
        // Loại bỏ khoảng trắng thừa
        const trimmedPair = pair.trim();
        if (trimmedPair) {
            // Tách thuật ngữ và định nghĩa
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

    // Tạo workbook và worksheet
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);

    // Đặt tiêu đề cho cột
    ws['!cols'] = [{ wpx: 300 }, { wpx: 600 }]; // Điều chỉnh chiều rộng cột
    ws['A1'] = { v: 'Term', t: 's' }; // Tiêu đề cột Thuật ngữ
    ws['B1'] = { v: 'Definition', t: 's' }; // Tiêu đề cột Định nghĩa

    XLSX.utils.book_append_sheet(wb, ws, "Vocabulary");

    // Ghi vào file
    XLSX.writeFile(wb, "vocabulary.xlsx");
});
