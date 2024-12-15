form.onsubmit = function (event) {
    event.preventDefault();

    const formData = new FormData(form);

    // Student Details
    const StudentName = formData.get('StudentName');
    const Roll = formData.get('Roll');
    const classTeacher = formData.get('classTeacher');
    const FatherName = formData.get('FatherName');
    const classSection = formData.get('ClassSection');
    const AdmissionNumber = formData.get('AdmissionNumber');
    const MotherName = formData.get('MotherName');
    const dob = formData.get('dob');

    // Files: Marks & Image
    const marksFile = document.getElementById('File').files[0];
    const imageFile = document.getElementById('profilePhoto').files[0];

    // Parse Excel Marks File
    if (marksFile) {
        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Get the first sheet
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // Convert to JSON
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            // Open new window and display content
            openNewWindow(StudentName, Roll, classTeacher, FatherName, classSection, AdmissionNumber, jsonData, imageFile, MotherName, dob);
        };
        reader.readAsArrayBuffer(marksFile);
    }
};

function openNewWindow(StudentName, Roll, classTeacher, FatherName, classSection, AdmissionNumber, tableData, imageFile, MotherName, DOB) {
    // Open new window
    const newWindow = window.open("", "_blank");

    // Start building the content
    let content = `
        <html>
        <head>
            <title>Student Information</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                table, th, td { border: 1px solid black; }
                th, td { padding: 8px; text-align: left; }
                h1, h2 { text-align: center; }
            </style>
        </head>
        <body>
            <h1>Student Information</h1>
            <p><strong>Name:</strong> ${StudentName}</p>
            <p><strong>Roll Number:</strong> ${Roll}</p>
            <p><strong>Class Teacher:</strong> ${classTeacher}</p>
            <p><strong>Father's Name:</strong> ${FatherName}</p>
            <p><strong>Class & Section:</strong> ${classSection}</p>
            <p><strong>Admission Number:</strong> ${AdmissionNumber}</p>
            <p><strong>Mother's Name:</strong> ${MotherName}</p>
            <p><strong>Date of Birth:</strong> ${DOB}</p>
    `;

    // Add image preview if available
    if (imageFile) {
        const reader = new FileReader();
        reader.onload = function (e) {
            content += `<img src="${e.target.result}" alt="Student Photo" style="width: 200px; margin-top: 20px;">`;
            content += "</body></html>";
            newWindow.document.write(content);
            newWindow.document.close();
        };
        reader.readAsDataURL(imageFile);
    }

    // Add table data
    if (tableData) {
        content += `<h2>Marks Table</h2><table>`;
        tableData.forEach((row, rowIndex) => {
            content += `<tr>`;
            row.forEach(cell => {
                content += rowIndex === 0 ? `<th>${cell}</th>` : `<td>${cell}</td>`;
            });
            content += `</tr>`;
        });
        content += `</table>`;
    }

    // Close the content if there's no image
    if (!imageFile) {
        content += "</body></html>";
        newWindow.document.write(content);
        newWindow.document.close();
    }
}
