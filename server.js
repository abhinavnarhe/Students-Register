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
const EXCEL_FILE = path.join(DATA_DIR, 'students.xlsx');

// Create directories
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(IMAGES_DIR)) fs.mkdirSync(IMAGES_DIR, { recursive: true });

// In-memory storage
let students = [];

// Load existing data
try {
    if (fs.existsSync(EXCEL_FILE)) {
        const workbook = XLSX.readFile(EXCEL_FILE);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        students = XLSX.utils.sheet_to_json(sheet);
        console.log(`✅ Loaded ${students.length} existing students`);
    }
} catch (err) {
    console.log('No existing data');
}

// Save to Excel
function saveToExcel() {
    try {
        const exportData = students.map(s => ({
            "S.No": s["S.No"],
            "Student Name": s["Student Name"],
            "Father's Name": s["Father's Name"],
            "Address": s["Address"],
            "Pin Code": s["Pin Code"],
            "Contact No": s["Contact No"],
            "Emergency Contact": s["Emergency Contact"],
            "Blood Group": s["Blood Group"],
            "Photo File": s["Photo File"],
            "Signature File": s["Signature File"],
            "Date": s["Date"],
            "ID": s["ID"]
        }));
        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Students");
        XLSX.writeFile(wb, EXCEL_FILE);
        console.log(`💾 Saved ${students.length} students`);
    } catch (err) {
        console.log('Save error:', err.message);
    }
}

// Save image
function saveImage(base64Data, studentName, type) {
    try {
        if (!base64Data) return null;
        const base64String = base64Data.split(',')[1] || base64Data;
        const buffer = Buffer.from(base64String, 'base64');
        const safeName = studentName.replace(/[^a-z0-9]/gi, '_').substring(0, 30);
        const filename = `${safeName}_${type}_${Date.now()}.jpg`;
        fs.writeFileSync(path.join(IMAGES_DIR, filename), buffer);
        return filename;
    } catch (err) {
        console.log('Image save error:', err.message);
        return null;
    }
}

// Serve static files
app.use('/images', express.static(IMAGES_DIR));

// ============ SIMPLE HTML PAGE ============
const HTML = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Student Registration</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; font-family: Arial, sans-serif; }
        body { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; padding: 20px; }
        .container { max-width: 900px; margin: 0 auto; background: white; border-radius: 20px; overflow: hidden; box-shadow: 0 20px 60px rgba(0,0,0,0.3); }
        .header { background: linear-gradient(135deg, #1e3c72, #2a5298); color: white; padding: 30px; text-align: center; }
        .stats { background: #f0f0f0; padding: 15px; text-align: center; font-size: 18px; }
        .stats span { font-weight: bold; color: #2a5298; font-size: 24px; }
        .tabs { display: flex; }
        .tab { flex: 1; padding: 15px; text-align: center; background: #e0e0e0; cursor: pointer; font-weight: bold; border: none; }
        .tab.active { background: white; color: #2a5298; border-bottom: 3px solid #2a5298; }
        .tab-content { display: none; padding: 30px; }
        .tab-content.active { display: block; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #333; }
        input, select, textarea { width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 8px; font-size: 14px; }
        input:focus, select:focus, textarea:focus { outline: none; border-color: #2a5298; }
        .row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        button { background: linear-gradient(135deg, #1e3c72, #2a5298); color: white; padding: 15px; border: none; border-radius: 10px; font-size: 16px; font-weight: bold; cursor: pointer; width: 100%; margin-top: 10px; }
        .btn-download { background: linear-gradient(135deg, #28a745, #20c997); margin-top: 10px; }
        .error { color: red; font-size: 12px; margin-top: 5px; }
        .status { padding: 15px; border-radius: 8px; margin-top: 20px; display: none; }
        .status.success { background: #d4edda; color: #155724; display: block; }
        .status.error { background: #f8d7da; color: #721c24; display: block; }
        .status.loading { background: #d1ecf1; color: #0c5460; display: block; }
        .student-card { background: #f9f9f9; border-radius: 10px; padding: 15px; margin-bottom: 15px; display: flex; gap: 15px; align-items: center; }
        .student-card img { width: 80px; height: 80px; object-fit: cover; border-radius: 10px; }
        .student-info { flex: 1; }
        .student-name { font-weight: bold; font-size: 18px; color: #1e3c72; }
        .search { width: 100%; padding: 10px; margin-bottom: 20px; border: 2px solid #ddd; border-radius: 8px; }
        .file-hint { font-size: 11px; color: #666; margin-top: 5px; }
        @media (max-width: 600px) { .row { grid-template-columns: 1fr; } .student-card { flex-direction: column; text-align: center; } }
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>📝 Student Registration System</h1>
        <p>Register students with photo and signature</p>
    </div>
    <div class="stats">
        📊 Total Students: <span id="totalCount">0</span>
    </div>
    <div class="tabs">
        <button class="tab active" onclick="showTab('register')">📝 Register</button>
        <button class="tab" onclick="showTab('gallery')">🖼️ Gallery</button>
    </div>
    
    <!-- Register Tab -->
    <div id="registerTab" class="tab-content active">
        <form id="registerForm">
            <div class="row">
                <div class="form-group">
                    <label>📸 Student Photo *</label>
                    <input type="file" id="photo" accept=".jpg,.jpeg" required>
                    <div class="file-hint">Only JPG files, max 2MB</div>
                    <div class="error" id="photoError"></div>
                </div>
                <div class="form-group">
                    <label>👤 Student Name *</label>
                    <input type="text" id="studentName" placeholder="Full name" required>
                    <div class="error" id="nameError"></div>
                </div>
            </div>
            <div class="row">
                <div class="form-group">
                    <label>👨 Father's Name *</label>
                    <input type="text" id="fatherName" placeholder="Father's name" required>
                    <div class="error" id="fatherError"></div>
                </div>
                <div class="form-group">
                    <label>🩸 Blood Group *</label>
                    <select id="bloodGroup" required>
                        <option value="">Select</option>
                        <option>A+</option><option>A-</option><option>B+</option><option>B-</option>
                        <option>O+</option><option>O-</option><option>AB+</option><option>AB-</option>
                    </select>
                    <div class="error" id="bloodError"></div>
                </div>
            </div>
            <div class="form-group">
                <label>📍 Address *</label>
                <textarea id="address" rows="2" placeholder="Complete address" required></textarea>
                <div class="error" id="addressError"></div>
            </div>
            <div class="row">
                <div class="form-group">
                    <label>📮 Pin Code *</label>
                    <input type="text" id="pincode" maxlength="6" placeholder="6 digits" required>
                    <div class="error" id="pincodeError"></div>
                </div>
                <div class="form-group">
                    <label>📞 Contact Number *</label>
                    <input type="tel" id="contact" maxlength="10" placeholder="10 digits" required>
                    <div class="error" id="contactError"></div>
                </div>
            </div>
            <div class="row">
                <div class="form-group">
                    <label>🆘 Emergency Contact *</label>
                    <input type="tel" id="emergency" maxlength="10" placeholder="10 digits" required>
                    <div class="error" id="emergencyError"></div>
                </div>
                <div class="form-group">
                    <label>✍️ Signature *</label>
                    <input type="file" id="signature" accept=".jpg,.jpeg" required>
                    <div class="file-hint">Only JPG files, max 2MB</div>
                    <div class="error" id="signError"></div>
                </div>
            </div>
            <button type="submit">✅ Register Student</button>
            <button type="button" class="btn-download" onclick="downloadExcel()">📥 Download Excel Report</button>
        </form>
        <div id="status" class="status"></div>
    </div>
    
    <!-- Gallery Tab -->
    <div id="galleryTab" class="tab-content">
        <input type="text" id="searchInput" class="search" placeholder="🔍 Search by name or father's name..." onkeyup="filterGallery()">
        <div id="galleryList"></div>
    </div>
</div>

<script>
let allStudents = [];

function showTab(tab) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    if (tab === 'register') {
        document.querySelector('.tab:first-child').classList.add('active');
        document.getElementById('registerTab').classList.add('active');
    } else {
        document.querySelector('.tab:last-child').classList.add('active');
        document.getElementById('galleryTab').classList.add('active');
        loadGallery();
    }
}

async function loadGallery() {
    try {
        const res = await fetch('/api/students');
        const data = await res.json();
        allStudents = data.students;
        document.getElementById('totalCount').innerText = allStudents.length;
        displayGallery(allStudents);
    } catch(e) {
        document.getElementById('galleryList').innerHTML = '<p style="color:red">Error loading students</p>';
    }
}

function displayGallery(students) {
    const container = document.getElementById('galleryList');
    if (students.length === 0) {
        container.innerHTML = '<p style="text-align:center;padding:50px">No students registered yet</p>';
        return;
    }
    container.innerHTML = students.map(s => `
        <div class="student-card">
            <div>
                ${s["Photo File"] ? `<img src="/images/${s["Photo File"]}" alt="Photo">` : '<div style="width:80px;height:80px;background:#ddd;display:flex;align-items:center;justify-content:center">📷</div>'}
                ${s["Signature File"] ? `<img src="/images/${s["Signature File"]}" alt="Signature" style="margin-top:5px">` : ''}
            </div>
            <div class="student-info">
                <div class="student-name">${escapeHtml(s["Student Name"])}</div>
                <div>👨 Father: ${escapeHtml(s["Father's Name"] || 'N/A')}</div>
                <div>🩸 Blood: ${s["Blood Group"] || 'N/A'}</div>
                <div>📞 ${s["Contact No"] || 'N/A'}</div>
                <div>📍 ${escapeHtml(s["Address"] || 'N/A').substring(0, 50)}</div>
                <div>📅 ${s["Date"] || 'N/A'}</div>
            </div>
        </div>
    `).join('');
}

function filterGallery() {
    const term = document.getElementById('searchInput').value.toLowerCase();
    const filtered = allStudents.filter(s => 
        (s["Student Name"] || '').toLowerCase().includes(term) ||
        (s["Father's Name"] || '').toLowerCase().includes(term)
    );
    displayGallery(filtered);
}

function escapeHtml(text) {
    if (!text) return '';
    return text.replace(/[&<>]/g, function(m) {
        if (m === '&') return '&amp;';
        if (m === '<') return '&lt;';
        if (m === '>') return '&gt;';
        return m;
    });
}

async function downloadExcel() {
    window.open('/api/download-excel', '_blank');
}

async function loadTotal() {
    try {
        const res = await fetch('/api/students');
        const data = await res.json();
        document.getElementById('totalCount').innerText = data.total;
    } catch(e) {}
}
loadTotal();

const form = document.getElementById('registerForm');
form.addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const statusDiv = document.getElementById('status');
    statusDiv.className = 'status loading';
    statusDiv.innerHTML = '⏳ Processing registration...';
    document.querySelector('#registerForm button[type="submit"]').disabled = true;
    
    try {
        const photoFile = document.getElementById('photo').files[0];
        const signFile = document.getElementById('signature').files[0];
        
        if (!photoFile || !signFile) throw new Error('Please select both photo and signature');
        
        const toBase64 = (file) => new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => resolve(reader.result);
            reader.onerror = reject;
        });
        
        const data = {
            studentName: document.getElementById('studentName').value.trim(),
            fatherName: document.getElementById('fatherName').value.trim(),
            address: document.getElementById('address').value.trim(),
            pincode: document.getElementById('pincode').value.trim(),
            contactNo: document.getElementById('contact').value.trim(),
            emergencyContact: document.getElementById('emergency').value.trim(),
            bloodGroup: document.getElementById('bloodGroup').value,
            photoBase64: await toBase64(photoFile),
            signatureBase64: await toBase64(signFile)
        };
        
        // Validation
        if (!data.studentName) throw new Error('Student name required');
        if (!data.fatherName) throw new Error('Father name required');
        if (!data.address) throw new Error('Address required');
        if (!/^\d{6}$/.test(data.pincode)) throw new Error('Pincode must be 6 digits');
        if (!/^\d{10}$/.test(data.contactNo)) throw new Error('Contact number must be 10 digits');
        if (!/^\d{10}$/.test(data.emergencyContact)) throw new Error('Emergency contact must be 10 digits');
        if (!data.bloodGroup) throw new Error('Blood group required');
        
        const res = await fetch('/api/submit', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        
        const result = await res.json();
        
        if (result.success) {
            statusDiv.className = 'status success';
            statusDiv.innerHTML = `✅ ${result.message}`;
            form.reset();
            loadTotal();
        } else {
            throw new Error(result.message);
        }
    } catch (err) {
        statusDiv.className = 'status error';
        statusDiv.innerHTML = `❌ Error: ${err.message}`;
    } finally {
        document.querySelector('#registerForm button[type="submit"]').disabled = false;
        setTimeout(() => { statusDiv.className = 'status'; }, 3000);
    }
});
</script>
</body>
</html>`;

// ============ API ROUTES ============

app.get('/', (req, res) => {
    res.send(HTML);
});

app.post('/api/submit', (req, res) => {
    console.log('📥 New registration');
    try {
        const { studentName, fatherName, address, pincode, contactNo, emergencyContact, bloodGroup, photoBase64, signatureBase64 } = req.body;
        
        // Validate
        if (!studentName || !fatherName || !address || !pincode || !contactNo || !emergencyContact || !bloodGroup) {
            return res.status(400).json({ success: false, message: 'Missing required fields' });
        }
        
        // Save images
        const photoFile = saveImage(photoBase64, studentName, 'photo');
        const signFile = saveImage(signatureBase64, studentName, 'sign');
        
        // Add to array
        const newStudent = {
            "S.No": students.length + 1,
            "Student Name": studentName,
            "Father's Name": fatherName,
            "Address": address,
            "Pin Code": pincode,
            "Contact No": contactNo,
            "Emergency Contact": emergencyContact,
            "Blood Group": bloodGroup,
            "Photo File": photoFile || '',
            "Signature File": signFile || '',
            "Date": new Date().toLocaleString(),
            "ID": `STU${Date.now()}`
        };
        
        students.push(newStudent);
        saveToExcel();
        
        console.log(`✅ Registered: ${studentName} (Total: ${students.length})`);
        res.json({ success: true, message: `Registered ${studentName}! Total: ${students.length}` });
        
    } catch (err) {
        console.error('Error:', err);
        res.status(500).json({ success: false, message: err.message });
    }
});

app.get('/api/students', (req, res) => {
    res.json({ students: students, total: students.length });
});

app.get('/api/download-excel', (req, res) => {
    try {
        if (students.length === 0) {
            return res.status(404).send('No data available');
        }
        
        const exportData = students.map(s => ({
            "S.No": s["S.No"],
            "Student Name": s["Student Name"],
            "Father's Name": s["Father's Name"],
            "Address": s["Address"],
            "Pin Code": s["Pin Code"],
            "Contact No": s["Contact No"],
            "Emergency Contact": s["Emergency Contact"],
            "Blood Group": s["Blood Group"],
            "Photo File": s["Photo File"],
            "Signature File": s["Signature File"],
            "Registration Date": s["Date"],
            "ID": s["ID"]
        }));
        
        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Students");
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        
        res.setHeader('Content-Disposition', 'attachment; filename=students.xlsx');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (err) {
        res.status(500).send(err.message);
    }
});

app.get('/health', (req, res) => {
    res.json({ status: 'ok', students: students.length, uptime: process.uptime() });
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`
    ╔════════════════════════════════════════════╗
    ║   ✅ SERVER RUNNING SUCCESSFULLY!         ║
    ║   📡 Port: ${PORT}                          ║
    ║   👥 Students: ${students.length}            ║
    ║   🌐 URL: http://localhost:${PORT}          ║
    ╚════════════════════════════════════════════╝
    `);
});