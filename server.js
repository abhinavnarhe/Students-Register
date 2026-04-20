const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// ============ CONFIGURATION ============
const DATA_DIR = path.join(__dirname, 'student_data');
const EXCEL_FILE = path.join(DATA_DIR, 'All_Students_Master.xlsx');

if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
}

// In-memory storage
let inMemoryStudents = [];

// Load existing data
function loadExistingData() {
    if (fs.existsSync(EXCEL_FILE)) {
        try {
            const workbook = XLSX.readFile(EXCEL_FILE);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            inMemoryStudents = XLSX.utils.sheet_to_json(worksheet);
            console.log(`📚 Loaded ${inMemoryStudents.length} existing students`);
        } catch (err) {
            console.log('No existing data found');
        }
    }
}

// Save to file
function saveToFile() {
    if (inMemoryStudents.length > 0) {
        try {
            const worksheet = XLSX.utils.json_to_sheet(inMemoryStudents);
            autoFitColumns(worksheet);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "All_Students");
            XLSX.writeFile(workbook, EXCEL_FILE);
            console.log(`💾 Saved ${inMemoryStudents.length} students to file`);
        } catch (err) {
            console.log('Error saving to file:', err.message);
        }
    }
}

function autoFitColumns(worksheet) {
    if (!worksheet['!ref']) return;
    
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const colWidths = {};
    
    const headers = ["S.No", "Student Name", "Father's Name", "Address", "Pin Code",
        "Contact No", "Emergency Contact No", "Blood Group",
        "Photo (Base64)", "Signature (Base64)", "Registration Date", "Registration ID"];
    
    headers.forEach((header, idx) => {
        colWidths[idx] = Math.max(header.length + 3, 20);
    });
    
    for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];
            if (cell && cell.v) {
                const cellValue = String(cell.v);
                const maxWidth = 50;
                const cellLength = Math.min(cellValue.length, maxWidth);
                colWidths[col] = Math.max(colWidths[col] || 0, cellLength + 2);
            }
        }
    }
    
    worksheet['!cols'] = [];
    for (let col = range.s.c; col <= range.e.c; col++) {
        let width = colWidths[col] || 20;
        width = Math.min(60, Math.max(15, width));
        worksheet['!cols'].push({ wch: width });
    }
}

loadExistingData();
setInterval(saveToFile, 5 * 60 * 1000);

// HTML Content (embedded)
const HTML_CONTENT = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Student Registration System</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        body { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; padding: 20px; }
        .container { max-width: 1000px; margin: 0 auto; background: white; border-radius: 20px; box-shadow: 0 20px 60px rgba(0,0,0,0.3); overflow: hidden; }
        .header { background: linear-gradient(135deg, #1e3c72, #2a5298); color: white; padding: 30px; text-align: center; }
        .stats { background: #f8f9fa; padding: 15px; text-align: center; border-bottom: 1px solid #dee2e6; }
        .stats span { font-size: 24px; font-weight: bold; color: #2a5298; }
        .form-container { padding: 40px; }
        .form-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 25px; }
        .full-width { grid-column: span 2; }
        .input-group { display: flex; flex-direction: column; gap: 8px; }
        .input-group label { font-weight: 600; color: #333; font-size: 0.9rem; }
        .required { color: red; }
        input, select, textarea { padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px; font-size: 14px; }
        input:focus, select:focus, textarea:focus { outline: none; border-color: #2a5298; }
        button { background: linear-gradient(135deg, #1e3c72, #2a5298); color: white; padding: 15px 30px; border: none; border-radius: 10px; font-size: 16px; font-weight: 600; cursor: pointer; width: 100%; margin-top: 20px; }
        button:disabled { opacity: 0.6; cursor: not-allowed; }
        .btn-download { background: linear-gradient(135deg, #28a745, #20c997); margin-top: 10px; }
        .status { margin-top: 20px; padding: 15px; border-radius: 8px; display: none; }
        .status.success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; display: block; }
        .status.error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; display: block; }
        .status.loading { background: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; display: block; }
        .error-message { color: #e74c3c; font-size: 12px; margin-top: 5px; }
        .file-hint { font-size: 11px; color: #666; margin-top: 5px; }
        .preview-img { max-width: 100px; margin-top: 10px; border-radius: 8px; }
        @media (max-width: 768px) { .form-grid { grid-template-columns: 1fr; } .full-width { grid-column: span 1; } .form-container { padding: 20px; } }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-id-card"></i> Student Registration System</h1>
            <p>Complete all fields to register</p>
        </div>
        <div class="stats">
            <i class="fas fa-users"></i> Total Registered Students: <span id="totalStudents">0</span>
        </div>
        <div class="form-container">
            <form id="registrationForm">
                <div class="form-grid">
                    <div class="input-group">
                        <label><i class="fas fa-camera"></i> Student Photo <span class="required">*</span></label>
                        <input type="file" id="studentPhoto" accept=".jpg,.jpeg" required>
                        <div class="file-hint">Only .jpg files, max 2MB</div>
                        <div class="error-message" id="photoError"></div>
                    </div>
                    <div class="input-group">
                        <label><i class="fas fa-user-graduate"></i> Student Name <span class="required">*</span></label>
                        <input type="text" id="studentName" placeholder="Enter full name" required>
                        <div class="error-message" id="nameError"></div>
                    </div>
                    <div class="input-group">
                        <label><i class="fas fa-user-friends"></i> Father's Name <span class="required">*</span></label>
                        <input type="text" id="fatherName" placeholder="Enter father's name" required>
                        <div class="error-message" id="fatherError"></div>
                    </div>
                    <div class="input-group full-width">
                        <label><i class="fas fa-map-marker-alt"></i> Address <span class="required">*</span></label>
                        <textarea id="address" rows="3" placeholder="Enter complete address" required></textarea>
                        <div class="error-message" id="addressError"></div>
                    </div>
                    <div class="input-group">
                        <label><i class="fas fa-map-pin"></i> Pin Code <span class="required">*</span></label>
                        <input type="text" id="pincode" maxlength="6" placeholder="6 digits" required>
                        <div class="error-message" id="pincodeError"></div>
                    </div>
                    <div class="input-group">
                        <label><i class="fas fa-phone-alt"></i> Contact Number <span class="required">*</span></label>
                        <input type="tel" id="contactNo" maxlength="10" placeholder="10 digits" required>
                        <div class="error-message" id="contactError"></div>
                    </div>
                    <div class="input-group">
                        <label><i class="fas fa-phone-volume"></i> Emergency Contact <span class="required">*</span></label>
                        <input type="tel" id="emergencyContact" maxlength="10" placeholder="10 digits" required>
                        <div class="error-message" id="emergencyError"></div>
                    </div>
                    <div class="input-group">
                        <label><i class="fas fa-tint"></i> Blood Group <span class="required">*</span></label>
                        <select id="bloodGroup" required>
                            <option value="">Select Blood Group</option>
                            <option>A+</option><option>A-</option><option>B+</option><option>B-</option>
                            <option>O+</option><option>O-</option><option>AB+</option><option>AB-</option>
                        </select>
                        <div class="error-message" id="bloodError"></div>
                    </div>
                    <div class="input-group">
                        <label><i class="fas fa-signature"></i> Signature <span class="required">*</span></label>
                        <input type="file" id="signature" accept=".jpg,.jpeg" required>
                        <div class="file-hint">Only .jpg files, max 2MB</div>
                        <div class="error-message" id="signError"></div>
                    </div>
                </div>
                <button type="submit" id="submitBtn"><i class="fas fa-paper-plane"></i> Register Student</button>
                <button type="button" class="btn-download" onclick="downloadExcel()"><i class="fas fa-download"></i> Download All Data (Excel with Images)</button>
            </form>
            <div id="status" class="status"></div>
        </div>
    </div>
    <script>
        async function loadTotalStudents() {
            try {
                const response = await fetch('/api/students');
                const data = await response.json();
                document.getElementById('totalStudents').textContent = data.total || 0;
            } catch (error) { console.error(error); }
        }
        loadTotalStudents();
        
        const form = document.getElementById('registrationForm');
        const statusDiv = document.getElementById('status');
        const submitBtn = document.getElementById('submitBtn');
        
        function clearErrors() { document.querySelectorAll('.error-message').forEach(el => el.textContent = ''); }
        
        function validateFile(file, fieldName) {
            if (!file) return fieldName + ' is required';
            if (!file.name.match(/\\.(jpg|jpeg)$/i)) return fieldName + ' must be a .jpg file';
            if (file.size > 2 * 1024 * 1024) return fieldName + ' must be less than 2MB';
            return null;
        }
        
        function validateForm() {
            let isValid = true;
            clearErrors();
            const studentName = document.getElementById('studentName').value.trim();
            if (!studentName) { document.getElementById('nameError').textContent = 'Student name is required'; isValid = false; }
            const fatherName = document.getElementById('fatherName').value.trim();
            if (!fatherName) { document.getElementById('fatherError').textContent = "Father's name is required"; isValid = false; }
            const address = document.getElementById('address').value.trim();
            if (!address) { document.getElementById('addressError').textContent = 'Address is required'; isValid = false; }
            const pincode = document.getElementById('pincode').value.trim();
            if (!pincode) { document.getElementById('pincodeError').textContent = 'Pin code is required'; isValid = false; }
            else if (!/^\\d{6}$/.test(pincode)) { document.getElementById('pincodeError').textContent = 'Pin code must be 6 digits'; isValid = false; }
            const contactNo = document.getElementById('contactNo').value.trim();
            if (!contactNo) { document.getElementById('contactError').textContent = 'Contact number is required'; isValid = false; }
            else if (!/^\\d{10}$/.test(contactNo)) { document.getElementById('contactError').textContent = 'Contact number must be 10 digits'; isValid = false; }
            const emergencyContact = document.getElementById('emergencyContact').value.trim();
            if (!emergencyContact) { document.getElementById('emergencyError').textContent = 'Emergency contact is required'; isValid = false; }
            else if (!/^\\d{10}$/.test(emergencyContact)) { document.getElementById('emergencyError').textContent = 'Emergency contact must be 10 digits'; isValid = false; }
            const bloodGroup = document.getElementById('bloodGroup').value;
            if (!bloodGroup) { document.getElementById('bloodError').textContent = 'Please select blood group'; isValid = false; }
            const photoFile = document.getElementById('studentPhoto').files[0];
            const photoError = validateFile(photoFile, 'Student photo');
            if (photoError) { document.getElementById('photoError').textContent = photoError; isValid = false; }
            const signFile = document.getElementById('signature').files[0];
            const signError = validateFile(signFile, 'Signature');
            if (signError) { document.getElementById('signError').textContent = signError; isValid = false; }
            return isValid;
        }
        
        function fileToBase64(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.readAsDataURL(file);
                reader.onload = () => resolve(reader.result);
                reader.onerror = error => reject(error);
            });
        }
        
        async function downloadExcel() { window.open('/api/download-excel', '_blank'); }
        
        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!validateForm()) return;
            statusDiv.className = 'status loading';
            statusDiv.innerHTML = '<i class="fas fa-spinner fa-pulse"></i> Processing registration...';
            submitBtn.disabled = true;
            try {
                const photoFile = document.getElementById('studentPhoto').files[0];
                const signatureFile = document.getElementById('signature').files[0];
                const photoBase64 = await fileToBase64(photoFile);
                const signatureBase64 = await fileToBase64(signatureFile);
                const formData = {
                    studentName: document.getElementById('studentName').value.trim(),
                    fatherName: document.getElementById('fatherName').value.trim(),
                    address: document.getElementById('address').value.trim(),
                    pincode: document.getElementById('pincode').value.trim(),
                    contactNo: document.getElementById('contactNo').value.trim(),
                    emergencyContact: document.getElementById('emergencyContact').value.trim(),
                    bloodGroup: document.getElementById('bloodGroup').value,
                    photoBase64: photoBase64,
                    signatureBase64: signatureBase64
                };
                const response = await fetch('/api/submit', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(formData) });
                const result = await response.json();
                if (result.success) {
                    statusDiv.className = 'status success';
                    statusDiv.innerHTML = '✅ ' + result.message;
                    form.reset();
                    loadTotalStudents();
                } else { throw new Error(result.message); }
            } catch (error) {
                statusDiv.className = 'status error';
                statusDiv.innerHTML = '❌ Error: ' + error.message;
            } finally { submitBtn.disabled = false; }
        });
    </script>
</body>
</html>`;

// Serve HTML
app.get('/', (req, res) => {
    res.send(HTML_CONTENT);
});

// Submit registration
app.post('/api/submit', async (req, res) => {
    console.log('\n📥 New registration received');
    try {
        const { studentName, fatherName, address, pincode, contactNo,
                emergencyContact, bloodGroup, photoBase64, signatureBase64 } = req.body;
        
        if (!studentName || !fatherName || !address || !pincode || 
            !contactNo || !emergencyContact || !bloodGroup) {
            return res.status(400).json({ success: false, message: 'Missing required fields' });
        }
        
        console.log(`👤 Student: ${studentName}`);
        
        // Create new entry with images as base64
        const newEntry = {
            "S.No": inMemoryStudents.length + 1,
            "Student Name": studentName,
            "Father's Name": fatherName,
            "Address": address,
            "Pin Code": pincode,
            "Contact No": contactNo,
            "Emergency Contact No": emergencyContact,
            "Blood Group": bloodGroup,
            "Photo (Base64)": photoBase64 || '',
            "Signature (Base64)": signatureBase64 || '',
            "Registration Date": new Date().toLocaleString(),
            "Registration ID": `REG${Date.now()}${Math.floor(Math.random() * 1000)}`
        };
        
        inMemoryStudents.push(newEntry);
        saveToFile();
        
        console.log(`✅ Registered! Total students: ${inMemoryStudents.length}`);
        
        res.json({ 
            success: true, 
            message: `Registration successful! Total students: ${inMemoryStudents.length}`,
            totalStudents: inMemoryStudents.length
        });
        
    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ success: false, message: error.message });
    }
});

// Get all students
app.get('/api/students', (req, res) => {
    const studentsWithoutImages = inMemoryStudents.map(s => ({
        "S.No": s["S.No"],
        "Student Name": s["Student Name"],
        "Father's Name": s["Father's Name"],
        "Address": s["Address"],
        "Pin Code": s["Pin Code"],
        "Contact No": s["Contact No"],
        "Emergency Contact No": s["Emergency Contact No"],
        "Blood Group": s["Blood Group"],
        "Registration Date": s["Registration Date"],
        "Registration ID": s["Registration ID"]
    }));
    res.json({ success: true, students: studentsWithoutImages, total: inMemoryStudents.length });
});

// Download Excel file (with images embedded)
app.get('/api/download-excel', (req, res) => {
    try {
        if (inMemoryStudents.length === 0) {
            return res.status(404).json({ error: 'No data available' });
        }
        
        const worksheet = XLSX.utils.json_to_sheet(inMemoryStudents);
        autoFitColumns(worksheet);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "All_Students");
        
        const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=All_Students_Master.xlsx');
        res.send(excelBuffer);
        
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Health check
app.get('/health', (req, res) => {
    res.json({ 
        status: 'OK', 
        timestamp: new Date().toISOString(),
        totalStudents: inMemoryStudents.length
    });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`
    ╔════════════════════════════════════════════════════════════╗
    ║     🚀 STUDENT REGISTRATION SYSTEM - RENDER READY          ║
    ║                                                            ║
    ║   📡 Server: http://localhost:${PORT}                       ║
    ║   👥 Students Registered: ${inMemoryStudents.length}         ║
    ║   📸 Images stored as Base64 INSIDE Excel                  ║
    ║                                                            ║
    ║   ✅ Photos and signatures are saved in the Excel file!    ║
    ║   📋 Download Excel to see all data with images            ║
    ╚════════════════════════════════════════════════════════════╝
    `);
});