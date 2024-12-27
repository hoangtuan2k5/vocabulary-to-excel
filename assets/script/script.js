document.getElementById('exportButton').addEventListener('click', async function () {
    const button = this;
    button.classList.add('is-loading');

    try {
        const text = document.getElementById('inputText').value.trim();
        const termSeparator = document.getElementById('termSeparator').value || ':';
        const pairSeparator = document.getElementById('pairSeparator').value || '\n';
        const fileName = document.getElementById('fileName').value || 'vocabulary.xlsx';
        const sheetName = document.getElementById('sheetName').value || 'Vocabulary';
        const appendOption = document.getElementById('appendOption').checked;
        const fileInput = document.getElementById('fileInput').files[0];

        // Process input data
        const pairs = text.split(new RegExp(pairSeparator, 'g'));
        const data = pairs
            .map(pair => {
                const trimmedPair = pair.trim();
                if (!trimmedPair) return null;

                const splitIndex = trimmedPair.indexOf(termSeparator);
                if (splitIndex === -1) return null;

                const term = trimmedPair.substring(0, splitIndex).trim();
                const definition = trimmedPair.substring(splitIndex + termSeparator.length).trim();

                return term && definition ? { Term: term, Definition: definition } : null;
            })
            .filter(item => item !== null);

        if (data.length === 0) {
            throw new Error('No valid data to export');
        }

        if (appendOption && fileInput) {
            await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = function (event) {
                    try {
                        const wb = XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
                        const ws = XLSX.utils.json_to_sheet(data);
                        ws['!cols'] = [{ wpx: 200 }, { wpx: 400 }];
                        XLSX.utils.book_append_sheet(wb, ws, sheetName);
                        XLSX.writeFile(wb, fileName);
                        resolve();
                    } catch (error) {
                        reject(error);
                    }
                };
                reader.onerror = reject;
                reader.readAsArrayBuffer(fileInput);
            });
        } else {
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.json_to_sheet(data);
            ws['!cols'] = [{ wpx: 200 }, { wpx: 400 }];
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
            XLSX.writeFile(wb, fileName);
        }
    } catch (error) {
        console.error('Error exporting data:', error);
    } finally {
        button.classList.remove('is-loading');
    }
});
