function getDetails() {
    let input = document.getElementById('inputField').value.trim().toLowerCase();
    let resultDiv = document.getElementById('result');

    fetch('data.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            let workbook = XLSX.read(data, { type: 'array' });
            let sheetName = workbook.SheetNames[0];
            let sheet = workbook.Sheets[sheetName];
            let rows = XLSX.utils.sheet_to_json(sheet);

            let matchedUsers = rows.filter(row => {
                let hming = row.Hming ? row.Hming.toString().trim().toLowerCase() : '';
                let mobile = row.Mobile ? row.Mobile.toString().trim() : '';
                let presbytery = row.Presbytery ? row.Presbytery.toString().trim().toLowerCase() : '';
                let kohhran = row.Kohhran ? row.Kohhran.toString().trim().toLowerCase() : '';
                return hming.includes(input) || mobile.includes(input) || presbytery.includes(input) || kohhran.includes(input);
            });

            if (matchedUsers.length > 0) {
                resultDiv.innerHTML = matchedUsers.map(user => `
                    <div class="user-details">
                        <p><strong>Hming:</strong> ${user.Hming}</p>
                        <p><strong>Mobile:</strong> ${user.Mobile}</p>
                        <p><strong>Presbytery:</strong> ${user.Presbytery}</p>
                        <p><strong>Kohhran:</strong> ${user.Kohhran}</p>
                        <p><strong>Thlen dan tur:</strong> ${user['Thlen dan tur']}</p>
                        <p><strong>Thlen In:</strong> ${user['Thlen In']}</p>
                    </div>
                `).join('');
            } else {
                resultDiv.innerHTML = '<p>No details found.</p>';
            }
        })
        .catch(error => {
            console.error('Error fetching or processing Excel file:', error);
            resultDiv.innerHTML = '<p>Error fetching or processing Excel file.</p>';
        });
}
