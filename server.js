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
const IMAGES_DIR = path.join(DATA_DIR, 'images');
const EXCEL_FILE = path.join(DATA_DIR, 'All_Students_Master.xlsx');

// Create directories
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(IMAGES_DIR)) fs.mkdirSync(IMAGES_DIR, { recursive: true });

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

// Save to file (without images in Excel)
function saveToFile() {
    if (inMemoryStudents.length > 0) {
        try {
            // Create a copy without images for Excel
            const studentsForExcel = inMemoryStudents.map(s => ({
                "S.No": s["S.No"],
                "Student Name": s["Student Name"],
                "Father's Name": s["Father's Name"],
                "Address": s["Address"],
                "Pin Code": s["Pin Code"],
                "Contact No": s["Contact No"],
                "Emergency Contact No": s["Emergency Contact No"],
                "Blood Group": s["Blood Group"],
                "Photo File": s["Photo File"] || '',
                "Signature File": s["Signature File"] || '',
                "Registration Date": s["Registration Date"],
                "Registration ID": s["Registration ID"]
            }));
            
            const worksheet = XLSX.utils.json_to_sheet(studentsForExcel);
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
        "Photo File", "Signature File", "Registration Date", "Registration ID"];
    
    headers.forEach((header, idx) => {
        colWidths[idx] = Math.max(header.length + 3, 20);
    });
    
    for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];
            if (cell && cell.v) {
                const cellValue = String(cell.v);
                const cellLength = Math.min(cellValue.length, 50);
                colWidths[col] = Math.max(colWidths[col] || 0, cellLength + 2);
            }
        }
    }
    
    worksheet['!cols'] = [];
    for (let col = range.s.c; col <= range.e.c; col++) {
        let width = colWidths[col] || 20;
        width = Math.min(50, Math.max(15, width));
        worksheet['!cols'].push({ wch: width });
    }
    
    // Enable text wrapping
    for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];
            if (cell) {
                if (!cell.s) cell.s = {};
                cell.s.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
            }
        }
    }
}

loadExistingData();
setInterval(saveToFile, 5 * 60 * 1000);

// Save image to disk and return filename
function saveImage(base64Data, studentName, type) {
    try {
        // Remove data URL prefix if present
        const base64String = base64Data.includes(',') ? base64Data.split(',')[1] : base64Data;
        const buffer = Buffer.from(base64String, 'base64');
        
        const safeName = studentName.replace(/[^a-zA-Z0-9]/g, '_');
        const timestamp = Date.now();
        const filename = `${safeName}_${type}_${timestamp}.jpg`;
        const filepath = path.join(IMAGES_DIR, filename);
        
        fs.writeFileSync(filepath, buffer);
        console.log(`💾 Saved image: ${filename}`);
        return filename;
    } catch (err) {
        console.error('Error saving image:', err);
        return null;
    }
}

// Serve static images
app.use('/images', express.static(IMAGES_DIR));

// ============ HTML CONTENT WITH GALLERY ============
const HTML_CONTENT = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Student Registration System</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        body { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; }
        
        .nav-tabs { display: flex; background: white; box-shadow: 0 2px 10px rgba(0,0,0,0.1); position: sticky; top: 0; z-index: 100; }
        .tab-btn { flex: 1; padding: 15px; border: none; background: white; font-size: 16px; font-weight: 600; cursor: pointer; transition: all 0.3s; color: #666; }
        .tab-btn.active { background: linear-gradient(135deg, #1e3c72, #2a5298); color: white; }
        .tab-btn:hover:not(.active) { background: #f0f0f0; }
        
        .tab-content { display: none; padding: 20px; max-width: 1400px; margin: 0 auto; }
        .tab-content.active { display: block; }
        
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
        
        .gallery-header { background: white; border-radius: 15px; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .gallery-stats { display: flex; gap: 20px; flex-wrap: wrap; }
        .stat-card { background: linear-gradient(135deg, #667eea, #764ba2); color: white; padding: 15px 25px; border-radius: 10px; flex: 1; min-width: 150px; text-align: center; }
        .stat-card i { font-size: 30px; margin-bottom: 10px; }
        .stat-card h3 { font-size: 28px; margin: 5px 0; }
        .search-box { margin-top: 20px; display: flex; gap: 10px; }
        .search-box input { flex: 1; padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px; }
        
        .student-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(350px, 1fr)); gap: 25px; margin-top: 20px; }
        .student-card { background: white; border-radius: 15px; overflow: hidden; box-shadow: 0 5px 15px rgba(0,0,0,0.1); transition: transform 0.3s; cursor: pointer; }
        .student-card:hover { transform: translateY(-5px); box-shadow: 0 10px 25px rgba(0,0,0,0.2); }
        .student-images { display: flex; padding: 20px; gap: 15px; background: #f8f9fa; border-bottom: 1px solid #eee; }
        .image-container { flex: 1; text-align: center; }
        .image-container img { width: 100%; height: 150px; object-fit: cover; border-radius: 10px; border: 3px solid white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .image-label { font-size: 11px; color: #666; margin-top: 5px; }
        .student-info { padding: 15px; }
        .student-name { font-size: 18px; font-weight: bold; color: #1e3c72; margin-bottom: 10px; }
        .student-detail { font-size: 13px; color: #555; margin: 5px 0; display: flex; align-items: center; gap: 8px; }
        .student-detail i { width: 20px; color: #2a5298; }
        
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.8); z-index: 1000; justify-content: center; align-items: center; }
        .modal.active { display: flex; }
        .modal-content { background: white; border-radius: 15px; max-width: 90%; max-height: 90%; overflow: auto; padding: 20px; }
        .modal-images { display: flex; gap: 20px; margin-bottom: 20px; flex-wrap: wrap; justify-content: center; }
        .modal-images img { max-width: 300px; border-radius: 10px; }
        .close-modal { position: absolute; top: 20px; right: 30px; font-size: 40px; color: white; cursor: pointer; }
        .loading-spinner { text-align: center; padding: 50px; font-size: 20px; color: white; }
        .no-data { text-align: center; padding: 50px; background: white; border-radius: 15px; color: #666; }
        
        @media (max-width: 768px) {
            .form-grid { grid-template-columns: 1fr; }
            .full-width { grid-column: span 1; }
            .form-container { padding: 20px; }
            .student-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
    <div class="nav-tabs">
        <button class="tab-btn active" onclick="switchTab('register')"><i class="fas fa-user-plus"></i> Register Student</button>
        <button class="tab-btn" onclick="switchTab('gallery')"><i class="fas fa-images"></i> Student Gallery</button>
    </div>
    
    <div id="registerTab" class="tab-content active">
        <div class="container">
            <div class="header"><h1><i class="fas fa-id-card"></i> Student Registration System</h1><p>Complete all fields to register</p></div>
            <div class="stats"><i class="fas fa-users"></i> Total Registered Students: <span id="totalStudents">0</span></div>
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
                    <button type="button" class="btn-download" onclick="downloadExcel()"><i class="fas fa-download"></i> Download All Data (Excel)</button>
                </form>
                <div id="status" class="status"></div>
            </div>
        </div>
    </div>
    
    <div id="galleryTab" class="tab-content">
        <div class="gallery-header">
            <div class="gallery-stats">
                <div class="stat-card"><i class="fas fa-users"></i><h3 id="galleryTotal">0</h3><p>Total Students</p></div>
                <div class="stat-card"><i class="fas fa-tint"></i><h3 id="bloodStats">0</h3><p>Blood Groups</p></div>
            </div>
            <div class="search-box">
                <input type="text" id="searchInput" placeholder="🔍 Search by name, father's name, or blood group..." onkeyup="filterStudents()">
            </div>
        </div>
        <div id="studentGrid" class="student-grid"><div class="loading-spinner"><i class="fas fa-spinner fa-pulse"></i> Loading students...</div></div>
    </div>
    
    <div id="modal" class="modal" onclick="closeModal()">
        <span class="close-modal">&times;</span>
        <div class="modal-content" onclick="event.stopPropagation()">
            <div id="modalContent"></div>
        </div>
    </div>
    
    <script>
        let allStudents = [];
        
        function switchTab(tab) {
            document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            if (tab === 'register') {
                document.getElementById('registerTab').classList.add('active');
                document.querySelector('.tab-btn:first-child').classList.add('active');
            } else {
                document.getElementById('galleryTab').classList.add('active');
                document.querySelector('.tab-btn:last-child').classList.add('active');
                loadGallery();
            }
        }
        
        async function loadGallery() {
            try {
                const response = await fetch('/api/students-with-images');
                const data = await response.json();
                allStudents = data.students;
                document.getElementById('galleryTotal').textContent = allStudents.length;
                const bloodGroups = new Set(allStudents.map(s => s["Blood Group"]).filter(b => b));
                document.getElementById('bloodStats').textContent = bloodGroups.size;
                displayStudents(allStudents);
            } catch (error) {
                document.getElementById('studentGrid').innerHTML = '<div class="no-data"><i class="fas fa-exclamation-triangle"></i> Error loading students</div>';
            }
        }
        
        function displayStudents(students) {
            const grid = document.getElementById('studentGrid');
            if (students.length === 0) {
                grid.innerHTML = '<div class="no-data"><i class="fas fa-user-graduate"></i> No students found</div>';
                return;
            }
            
            grid.innerHTML = students.map(student => {
                const photoUrl = student["Photo File"] ? `/images/${student["Photo File"]}` : null;
                const signUrl = student["Signature File"] ? `/images/${student["Signature File"]}` : null;
                
                return '<div class="student-card" onclick=\'showStudentDetail(' + JSON.stringify(student).replace(/'/g, "&#39;") + ')\'>' +
                    '<div class="student-images">' +
                        '<div class="image-container">' +
                            (photoUrl ? '<img src="' + photoUrl + '" alt="Photo" onerror="this.src=\'https://via.placeholder.com/150?text=No+Photo\'">' : '<img src="https://via.placeholder.com/150?text=No+Photo" alt="No Photo">') +
                            '<div class="image-label">Student Photo</div>' +
                        '</div>' +
                        '<div class="image-container">' +
                            (signUrl ? '<img src="' + signUrl + '" alt="Signature" onerror="this.src=\'https://via.placeholder.com/150?text=No+Signature\'">' : '<img src="https://via.placeholder.com/150?text=No+Signature" alt="No Signature">') +
                            '<div class="image-label">Signature</div>' +
                        '</div>' +
                    '</div>' +
                    '<div class="student-info">' +
                        '<div class="student-name">' + escapeHtml(student["Student Name"]) + '</div>' +
                        '<div class="student-detail"><i class="fas fa-user-friends"></i> ' + escapeHtml(student["Father's Name"] || 'N/A') + '</div>' +
                        '<div class="student-detail"><i class="fas fa-tint"></i> Blood: ' + (student["Blood Group"] || 'N/A') + '</div>' +
                        '<div class="student-detail"><i class="fas fa-phone"></i> ' + (student["Contact No"] || 'N/A') + '</div>' +
                        '<div class="student-detail"><i class="fas fa-calendar"></i> ' + (student["Registration Date"] || 'N/A') + '</div>' +
                    '</div>' +
                '</div>';
            }).join('');
        }
        
        function showStudentDetail(student) {
            const modal = document.getElementById('modal');
            const modalContent = document.getElementById('modalContent');
            const photoUrl = student["Photo File"] ? `/images/${student["Photo File"]}` : null;
            const signUrl = student["Signature File"] ? `/images/${student["Signature File"]}` : null;
            
            modalContent.innerHTML = '<h2 style="color: #1e3c72; margin-bottom: 20px;">' + escapeHtml(student["Student Name"]) + '</h2>' +
                '<div class="modal-images">' +
                    '<div style="text-align: center;"><h4>Student Photo</h4>' + (photoUrl ? '<img src="' + photoUrl + '" alt="Photo">' : '<p>No photo available</p>') + '</div>' +
                    '<div style="text-align: center;"><h4>Signature</h4>' + (signUrl ? '<img src="' + signUrl + '" alt="Signature">' : '<p>No signature available</p>') + '</div>' +
                '</div>' +
                '<div style="margin-top: 20px;">' +
                    '<p><strong>Father\'s Name:</strong> ' + escapeHtml(student["Father's Name"] || 'N/A') + '</p>' +
                    '<p><strong>Address:</strong> ' + escapeHtml(student["Address"] || 'N/A') + '</p>' +
                    '<p><strong>Pin Code:</strong> ' + (student["Pin Code"] || 'N/A') + '</p>' +
                    '<p><strong>Contact No:</strong> ' + (student["Contact No"] || 'N/A') + '</p>' +
                    '<p><strong>Emergency Contact:</strong> ' + (student["Emergency Contact No"] || 'N/A') + '</p>' +
                    '<p><strong>Blood Group:</strong> ' + (student["Blood Group"] || 'N/A') + '</p>' +
                    '<p><strong>Registration ID:</strong> ' + (student["Registration ID"] || 'N/A') + '</p>' +
                    '<p><strong>Registration Date:</strong> ' + (student["Registration Date"] || 'N/A') + '</p>' +
                '</div>';
            
            modal.classList.add('active');
        }
        
        function closeModal() { document.getElementById('modal').classList.remove('active'); }
        
        function filterStudents() {
            const searchTerm = document.getElementById('searchInput').value.toLowerCase();
            const filtered = allStudents.filter(student => 
                (student["Student Name"] || '').toLowerCase().includes(searchTerm) ||
                (student["Father's Name"] || '').toLowerCase().includes(searchTerm) ||
                (student["Blood Group"] || '').toLowerCase().includes(searchTerm) ||
                (student["Contact No"] || '').includes(searchTerm)
            );
            displayStudents(filtered);
        }
        
        function escapeHtml(text) { if (!text) return ''; const div = document.createElement('div'); div.textContent = text; return div.innerHTML; }
        
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
            if (!document.getElementById('studentName').value.trim()) { document.getElementById('nameError').textContent = 'Student name is required'; isValid = false; }
            if (!document.getElementById('fatherName').value.trim()) { document.getElementById('fatherError').textContent = "Father's name is required"; isValid = false; }
            if (!document.getElementById('address').value.trim()) { document.getElementById('addressError').textContent = 'Address is required'; isValid = false; }
            const pincode = document.getElementById('pincode').value.trim();
            if (!pincode) { document.getElementById('pincodeError').textContent = 'Pin code is required'; isValid = false; }
            else if (!/^\\d{6}$/.test(pincode)) { document.getElementById('pincodeError').textContent = 'Pin code must be 6 digits'; isValid = false; }
            const contactNo = document.getElementById('contactNo').value.trim();
            if (!contactNo) { document.getElementById('contactError').textContent = 'Contact number is required'; isValid = false; }
            else if (!/^\\d{10}$/.test(contactNo)) { document.getElementById('contactError').textContent = 'Contact number must be 10 digits'; isValid = false; }
            const emergencyContact = document.getElementById('emergencyContact').value.trim();
            if (!emergencyContact) { document.getElementById('emergencyError').textContent = 'Emergency contact is required'; isValid = false; }
            else if (!/^\\d{10}$/.test(emergencyContact)) { document.getElementById('emergencyError').textContent = 'Emergency contact must be 10 digits'; isValid = false; }
            if (!document.getElementById('bloodGroup').value) { document.getElementById('bloodError').textContent = 'Please select blood group'; isValid = false; }
            const photoError = validateFile(document.getElementById('studentPhoto').files[0], 'Student photo');
            if (photoError) { document.getElementById('photoError').textContent = photoError; isValid = false; }
            const signError = validateFile(document.getElementById('signature').files[0], 'Signature');
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

// ============ API ENDPOINTS ============

app.get('/', (req, res) => {
    res.send(HTML_CONTENT);
});

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
        
        // Save images to disk
        const photoFile = saveImage(photoBase64, studentName, 'photo');
        const signatureFile = saveImage(signatureBase64, studentName, 'sign');
        
        const newEntry = {
            "S.No": inMemoryStudents.length + 1,
            "Student Name": studentName,
            "Father's Name": fatherName,
            "Address": address,
            "Pin Code": pincode,
            "Contact No": contactNo,
            "Emergency Contact No": emergencyContact,
            "Blood Group": bloodGroup,
            "Photo File": photoFile || '',
            "Signature File": signatureFile || '',
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

app.get('/api/students-with-images', (req, res) => {
    res.json({ success: true, students: inMemoryStudents, total: inMemoryStudents.length });
});

app.get('/api/download-excel', (req, res) => {
    try {
        if (inMemoryStudents.length === 0) {
            return res.status(404).json({ error: 'No data available' });
        }
        
        const studentsForExcel = inMemoryStudents.map(s => ({
            "S.No": s["S.No"],
            "Student Name": s["Student Name"],
            "Father's Name": s["Father's Name"],
            "Address": s["Address"],
            "Pin Code": s["Pin Code"],
            "Contact No": s["Contact No"],
            "Emergency Contact No": s["Emergency Contact No"],
            "Blood Group": s["Blood Group"],
            "Photo File": s["Photo File"] || '',
            "Signature File": s["Signature File"] || '',
            "Registration Date": s["Registration Date"],
            "Registration ID": s["Registration ID"]
        }));
        
        const worksheet = XLSX.utils.json_to_sheet(studentsForExcel);
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
    ╔════════════════════════════════════════════════════════════════╗
    ║     🚀 STUDENT REGISTRATION SYSTEM WITH GALLERY - READY        ║
    ║                                                                ║
    ║   📡 Server: http://localhost:${PORT}                           ║
    ║   👥 Students Registered: ${inMemoryStudents.length}             ║
    ║   📸 Images stored in: /student_data/images/                   ║
    ║   📊 Excel stored in: /student_data/                           ║
    ║                                                                ║
    ║   ✅ NO Excel character limit issues!                          ║
    ║   ✅ Images displayed in Gallery tab!                          ║
    ║   ✅ Excel references image filenames!                         ║
    ║                                                                ║
    ║   🌐 Visit: https://your-app.onrender.com                      ║
    ╚════════════════════════════════════════════════════════════════╝
    `);
});
