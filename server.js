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

// Create directories synchronously on startup
try {
    if (!fs.existsSync(DATA_DIR)) {
        fs.mkdirSync(DATA_DIR, { recursive: true });
        console.log('✅ Created data directory');
    }
    if (!fs.existsSync(IMAGES_DIR)) {
        fs.mkdirSync(IMAGES_DIR, { recursive: true });
        console.log('✅ Created images directory');
    }
} catch (err) {
    console.error('Error creating directories:', err.message);
}

// In-memory storage
let inMemoryStudents = [];

// Load existing data
function loadExistingData() {
    try {
        if (fs.existsSync(EXCEL_FILE)) {
            const workbook = XLSX.readFile(EXCEL_FILE);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            inMemoryStudents = XLSX.utils.sheet_to_json(worksheet);
            console.log(`📚 Loaded ${inMemoryStudents.length} existing students`);
        } else {
            console.log('No existing data file found, starting fresh');
        }
    } catch (err) {
        console.log('Error loading data:', err.message);
    }
}

// Save to file
function saveToFile() {
    try {
        if (inMemoryStudents.length > 0) {
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
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "All_Students");
            XLSX.writeFile(workbook, EXCEL_FILE);
            console.log(`💾 Saved ${inMemoryStudents.length} students to file`);
        }
    } catch (err) {
        console.log('Error saving to file:', err.message);
    }
}

// Save image to disk
function saveImage(base64Data, studentName, type) {
    try {
        if (!base64Data) return null;
        const base64String = base64Data.includes(',') ? base64Data.split(',')[1] : base64Data;
        const buffer = Buffer.from(base64String, 'base64');
        const safeName = studentName.replace(/[^a-zA-Z0-9]/g, '_');
        const timestamp = Date.now();
        const filename = `${safeName}_${type}_${timestamp}.jpg`;
        const filepath = path.join(IMAGES_DIR, filename);
        fs.writeFileSync(filepath, buffer);
        return filename;
    } catch (err) {
        console.error('Error saving image:', err.message);
        return null;
    }
}

// Serve static images
app.use('/images', express.static(IMAGES_DIR));

// Load data on startup
loadExistingData();

// Auto-save every 5 minutes
setInterval(saveToFile, 5 * 60 * 1000);

// ============ SIMPLE HTML PAGE (embedded) ============
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
        .input-group label { font-weight: 600; color: #333; }
        .required { color: red; }
        input, select, textarea { padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px; font-size: 14px; }
        button { background: linear-gradient(135deg, #1e3c72, #2a5298); color: white; padding: 15px; border: none; border-radius: 10px; font-size: 16px; font-weight: 600; cursor: pointer; width: 100%; margin-top: 20px; }
        button:disabled { opacity: 0.6; }
        .btn-download { background: linear-gradient(135deg, #28a745, #20c997); margin-top: 10px; }
        .status { margin-top: 20px; padding: 15px; border-radius: 8px; display: none; }
        .status.success { background: #d4edda; color: #155724; display: block; }
        .status.error { background: #f8d7da; color: #721c24; display: block; }
        .status.loading { background: #d1ecf1; color: #0c5460; display: block; }
        .error-message { color: #e74c3c; font-size: 12px; margin-top: 5px; }
        .file-hint { font-size: 11px; color: #666; margin-top: 5px; }
        .gallery-header { background: white; border-radius: 15px; padding: 20px; margin-top: 20px; }
        .student-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 20px; margin-top: 20px; }
        .student-card { background: white; border-radius: 15px; overflow: hidden; box-shadow: 0 2px 10px rgba(0,0,0,0.1); cursor: pointer; }
        .student-card img { width: 100%; height: 150px; object-fit: cover; }
        .student-info { padding: 15px; }
        .student-name { font-weight: bold; font-size: 18px; color: #1e3c72; }
        .nav-tabs { display: flex; margin-bottom: 20px; }
        .tab-btn { background: #eee; color: #333; margin: 0 5px; padding: 10px 20px; border-radius: 10px; cursor: pointer; }
        .tab-btn.active { background: linear-gradient(135deg, #1e3c72, #2a5298); color: white; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        @media (max-width: 768px) { .form-grid { grid-template-columns: 1fr; } .full-width { grid-column: span 1; } }
    </style>
</head>
<body>
<div class="container">
    <div class="header"><h1><i class="fas fa-id-card"></i> Student Registration</h1><p>Complete all fields to register</p></div>
    <div class="stats"><i class="fas fa-users"></i> Total Students: <span id="totalStudents">0</span></div>
    
    <div class="nav-tabs">
        <button class="tab-btn active" onclick="switchTab('register')">Register</button>
        <button class="tab-btn" onclick="switchTab('gallery')">Gallery</button>
    </div>
    
    <div id="registerTab" class="tab-content active">
        <div class="form-container">
            <form id="registrationForm">
                <div class="form-grid">
                    <div class="input-group"><label>Student Photo <span class="required">*</span></label><input type="file" id="studentPhoto" accept=".jpg,.jpeg" required><div class="error-message" id="photoError"></div></div>
                    <div class="input-group"><label>Student Name <span class="required">*</span></label><input type="text" id="studentName" placeholder="Full name" required><div class="error-message" id="nameError"></div></div>
                    <div class="input-group"><label>Father's Name <span class="required">*</span></label><input type="text" id="fatherName" placeholder="Father's name" required><div class="error-message" id="fatherError"></div></div>
                    <div class="input-group full-width"><label>Address <span class="required">*</span></label><textarea id="address" rows="3" placeholder="Complete address" required></textarea><div class="error-message" id="addressError"></div></div>
                    <div class="input-group"><label>Pin Code <span class="required">*</span></label><input type="text" id="pincode" maxlength="6" placeholder="6 digits" required><div class="error-message" id="pincodeError"></div></div>
                    <div class="input-group"><label>Contact No <span class="required">*</span></label><input type="tel" id="contactNo" maxlength="10" placeholder="10 digits" required><div class="error-message" id="contactError"></div></div>
                    <div class="input-group"><label>Emergency Contact <span class="required">*</span></label><input type="tel" id="emergencyContact" maxlength="10" placeholder="10 digits" required><div class="error-message" id="emergencyError"></div></div>
                    <div class="input-group"><label>Blood Group <span class="required">*</span></label><select id="bloodGroup" required><option value="">Select</option><option>A+</option><option>A-</option><option>B+</option><option>B-</option><option>O+</option><option>O-</option><option>AB+</option><option>AB-</option></select><div class="error-message" id="bloodError"></div></div>
                    <div class="input-group"><label>Signature <span class="required">*</span></label><input type="file" id="signature" accept=".jpg,.jpeg" required><div class="error-message" id="signError"></div></div>
                </div>
                <button type="submit" id="submitBtn">Register Student</button>
                <button type="button" class="btn-download" onclick="downloadExcel()">Download Excel</button>
            </form>
            <div id="status" class="status"></div>
        </div>
    </div>
    
    <div id="galleryTab" class="tab-content">
        <div class="gallery-header"><input type="text" id="searchInput" placeholder="Search students..." style="width:100%;padding:10px;" onkeyup="filterStudents()"></div>
        <div id="studentGrid" class="student-grid">Loading...</div>
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
        const res = await fetch('/api/students-with-images');
        const data = await res.json();
        allStudents = data.students;
        document.getElementById('totalStudents').textContent = allStudents.length;
        displayStudents(allStudents);
    } catch(e) { document.getElementById('studentGrid').innerHTML = 'Error loading'; }
}
function displayStudents(students) {
    const grid = document.getElementById('studentGrid');
    if(students.length === 0) { grid.innerHTML = 'No students found'; return; }
    grid.innerHTML = students.map(s => '<div class="student-card" onclick="alert(\'Name: ' + s["Student Name"] + '\\nFather: ' + (s["Father's Name"]||'') + '\\nBlood: ' + (s["Blood Group"]||'') + '\')">' +
        '<div style="display:flex;gap:5px;padding:10px">' +
        (s["Photo File"] ? '<img src="/images/' + s["Photo File"] + '" style="width:50%;height:120px">' : '<div style="width:50%;background:#eee;display:flex;align-items:center;justify-content:center">No Photo</div>') +
        (s["Signature File"] ? '<img src="/images/' + s["Signature File"] + '" style="width:50%;height:120px">' : '<div style="width:50%;background:#eee;display:flex;align-items:center;justify-content:center">No Signature</div>') +
        '</div>' +
        '<div class="student-info"><div class="student-name">' + escapeHtml(s["Student Name"]) + '</div>' +
        '<div>Blood: ' + (s["Blood Group"]||'N/A') + '</div>' +
        '<div>Contact: ' + (s["Contact No"]||'N/A') + '</div></div></div>').join('');
}
function filterStudents() {
    const term = document.getElementById('searchInput').value.toLowerCase();
    const filtered = allStudents.filter(s => (s["Student Name"]||'').toLowerCase().includes(term) || (s["Father's Name"]||'').toLowerCase().includes(term));
    displayStudents(filtered);
}
function escapeHtml(text) { if(!text) return ''; return text.replace(/[&<>]/g, function(m){if(m==='&') return '&amp;'; if(m==='<') return '&lt;'; if(m==='>') return '&gt;'; return m;}); }
async function downloadExcel() { window.open('/api/download-excel', '_blank'); }
async function loadTotal() { try { const res = await fetch('/api/students'); const data = await res.json(); document.getElementById('totalStudents').textContent = data.total || 0; } catch(e){} }
loadTotal();
const form = document.getElementById('registrationForm');
form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const statusDiv = document.getElementById('status');
    statusDiv.className = 'status loading';
    statusDiv.innerHTML = 'Processing...';
    document.getElementById('submitBtn').disabled = true;
    try {
        const photoFile = document.getElementById('studentPhoto').files[0];
        const signFile = document.getElementById('signature').files[0];
        const toBase64 = (file) => new Promise((resolve, reject) => { const reader = new FileReader(); reader.readAsDataURL(file); reader.onload = () => resolve(reader.result); reader.onerror = reject; });
        const formData = {
            studentName: document.getElementById('studentName').value.trim(),
            fatherName: document.getElementById('fatherName').value.trim(),
            address: document.getElementById('address').value.trim(),
            pincode: document.getElementById('pincode').value.trim(),
            contactNo: document.getElementById('contactNo').value.trim(),
            emergencyContact: document.getElementById('emergencyContact').value.trim(),
            bloodGroup: document.getElementById('bloodGroup').value,
            photoBase64: await toBase64(photoFile),
            signatureBase64: await toBase64(signFile)
        };
        const res = await fetch('/api/submit', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(formData) });
        const result = await res.json();
        if(result.success) {
            statusDiv.className = 'status success';
            statusDiv.innerHTML = '✅ ' + result.message;
            form.reset();
            loadTotal();
        } else throw new Error(result.message);
    } catch(err) {
        statusDiv.className = 'status error';
        statusDiv.innerHTML = '❌ Error: ' + err.message;
    } finally { document.getElementById('submitBtn').disabled = false; }
});
</script>
</body>
</html>`;

// ============ API ROUTES ============
app.get('/', (req, res) => res.send(HTML_CONTENT));

app.post('/api/submit', async (req, res) => {
    console.log('📥 New registration');
    try {
        const { studentName, fatherName, address, pincode, contactNo, emergencyContact, bloodGroup, photoBase64, signatureBase64 } = req.body;
        if (!studentName || !fatherName || !address || !pincode || !contactNo || !emergencyContact || !bloodGroup) {
            return res.status(400).json({ success: false, message: 'Missing fields' });
        }
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
            "Registration ID": `REG${Date.now()}`
        };
        inMemoryStudents.push(newEntry);
        saveToFile();
        res.json({ success: true, message: `Registered! Total: ${inMemoryStudents.length}` });
    } catch(err) {
        res.status(500).json({ success: false, message: err.message });
    }
});

app.get('/api/students', (req, res) => {
    res.json({ success: true, students: inMemoryStudents, total: inMemoryStudents.length });
});

app.get('/api/students-with-images', (req, res) => {
    res.json({ success: true, students: inMemoryStudents, total: inMemoryStudents.length });
});

app.get('/api/download-excel', (req, res) => {
    try {
        if (inMemoryStudents.length === 0) return res.status(404).json({ error: 'No data' });
        const exportData = inMemoryStudents.map(s => ({
            "S.No": s["S.No"], "Student Name": s["Student Name"], "Father's Name": s["Father's Name"],
            "Address": s["Address"], "Pin Code": s["Pin Code"], "Contact No": s["Contact No"],
            "Emergency Contact No": s["Emergency Contact No"], "Blood Group": s["Blood Group"],
            "Photo File": s["Photo File"], "Signature File": s["Signature File"],
            "Registration Date": s["Registration Date"], "Registration ID": s["Registration ID"]
        }));
        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Students");
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        res.setHeader('Content-Disposition', 'attachment; filename=students.xlsx');
        res.send(buffer);
    } catch(err) { res.status(500).json({ error: err.message }); }
});

app.get('/health', (req, res) => res.json({ status: 'ok', students: inMemoryStudents.length }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`✅ Server running on port ${PORT}`);
    console.log(`📊 Current students: ${inMemoryStudents.length}`);
});