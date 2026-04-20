const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Serve static files from 'public' directory
app.use(express.static(path.join(__dirname, 'public')));

// ============ CONFIGURATION ============
const DATA_DIR = path.join(__dirname, 'student_data');
const EXCEL_FILE = path.join(DATA_DIR, 'All_Students_Master.xlsx');

// Create directories if they don't exist
if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
}

// In-memory storage for better performance
let inMemoryStudents = [];

// Load existing data from file if available
function loadExistingData() {
    if (fs.existsSync(EXCEL_FILE)) {
        try {
            const workbook = XLSX.readFile(EXCEL_FILE);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            inMemoryStudents = XLSX.utils.sheet_to_json(worksheet);
            console.log(`📚 Loaded ${inMemoryStudents.length} existing students`);
        } catch (err) {
            console.log('No existing data found, starting fresh');
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

// Auto-fit columns function
function autoFitColumns(worksheet) {
    if (!worksheet['!ref']) return;
    
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const colWidths = {};
    
    const headers = ["S.No", "Student Name", "Father's Name", "Address", "Pin Code",
        "Contact No", "Emergency Contact No", "Blood Group",
        "Photo File", "Signature File", "Student Folder", "Registration Date", "Registration ID"];
    
    headers.forEach((header, idx) => {
        colWidths[idx] = Math.max(header.length + 3, 15);
    });
    
    for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];
            if (cell && cell.v) {
                const cellValue = String(cell.v);
                const maxWidth = (col === 3) ? 50 : 35;
                const cellLength = Math.min(cellValue.length, maxWidth);
                colWidths[col] = Math.max(colWidths[col] || 0, cellLength + 2);
            }
        }
    }
    
    worksheet['!cols'] = [];
    for (let col = range.s.c; col <= range.e.c; col++) {
        let width = colWidths[col] || 15;
        width = Math.min(45, Math.max(12, width));
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

// Load existing data on startup
loadExistingData();

// Save data every 5 minutes (optional)
setInterval(saveToFile, 5 * 60 * 1000);

// ============ API ENDPOINTS ============

// Serve the main HTML page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Submit registration
app.post('/api/submit', async (req, res) => {
    console.log('\n📥 New registration received');
    
    try {
        const { studentName, fatherName, address, pincode, contactNo,
                emergencyContact, bloodGroup, photoBase64, signatureBase64 } = req.body;
        
        // Validation
        if (!studentName || !fatherName || !address || !pincode || 
            !contactNo || !emergencyContact || !bloodGroup) {
            return res.status(400).json({ 
                success: false, 
                message: 'Missing required fields' 
            });
        }
        
        console.log(`👤 Student: ${studentName}`);
        
        // Save photos to disk (optional - creates folders)
        let savedFiles = { folderName: 'N/A', photoFileName: 'N/A', signatureFileName: 'N/A' };
        try {
            const safeName = studentName.replace(/[^a-zA-Z0-9]/g, '_');
            const timestamp = Date.now();
            const folderName = `${safeName}_${timestamp}`;
            const studentFolder = path.join(DATA_DIR, folderName);
            
            if (!fs.existsSync(studentFolder)) {
                fs.mkdirSync(studentFolder, { recursive: true });
            }
            
            const photoFileName = `${safeName}_photo.jpg`;
            const signatureFileName = `${safeName}_sign.jpg`;
            
            if (photoBase64) {
                const photoBuffer = Buffer.from(photoBase64, 'base64');
                fs.writeFileSync(path.join(studentFolder, photoFileName), photoBuffer);
            }
            
            if (signatureBase64) {
                const signatureBuffer = Buffer.from(signatureBase64, 'base64');
                fs.writeFileSync(path.join(studentFolder, signatureFileName), signatureBuffer);
            }
            
            savedFiles = { folderName, photoFileName, signatureFileName };
            console.log(`📁 Created folder: ${folderName}`);
        } catch (fileError) {
            console.log('File save note:', fileError.message);
        }
        
        // Create new student entry
        const newEntry = {
            "S.No": inMemoryStudents.length + 1,
            "Student Name": studentName,
            "Father's Name": fatherName,
            "Address": address,
            "Pin Code": pincode,
            "Contact No": contactNo,
            "Emergency Contact No": emergencyContact,
            "Blood Group": bloodGroup,
            "Photo File": savedFiles.photoFileName,
            "Signature File": savedFiles.signatureFileName,
            "Student Folder": savedFiles.folderName,
            "Registration Date": new Date().toLocaleString(),
            "Registration ID": `REG${Date.now()}${Math.floor(Math.random() * 1000)}`
        };
        
        // Add to storage
        inMemoryStudents.push(newEntry);
        
        // Save to file immediately
        saveToFile();
        
        console.log(`✅ Registered! Total students: ${inMemoryStudents.length}`);
        
        res.json({ 
            success: true, 
            message: `✅ Registration successful! Total students: ${inMemoryStudents.length}`,
            totalStudents: inMemoryStudents.length
        });
        
    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ 
            success: false, 
            message: error.message 
        });
    }
});

// Get all students
app.get('/api/students', (req, res) => {
    res.json({ 
        success: true, 
        students: inMemoryStudents, 
        total: inMemoryStudents.length 
    });
});

// Download Excel file
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
        totalStudents: inMemoryStudents.length,
        server: 'Render.com Ready'
    });
});

// Catch-all route to handle client-side routing (optional)
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`
    ╔════════════════════════════════════════════════════════════╗
    ║     🚀 STUDENT REGISTRATION SYSTEM - RENDER READY          ║
    ║                                                            ║
    ║   📡 Server: http://localhost:${PORT}                       ║
    ║   👥 Students Registered: ${inMemoryStudents.length}         ║
    ║   💾 Storage: File + Memory                                ║
    ║                                                            ║
    ║   ✅ Ready for deployment on Render.com                    ║
    ║   📋 Visit: https://your-app.onrender.com                  ║
    ╚════════════════════════════════════════════════════════════╝
    `);
});