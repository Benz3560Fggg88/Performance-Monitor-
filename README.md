# 📊 Performance-Monitor

**Performance-Monitor** เป็นชุดเครื่องมือ Python สำหรับการติดตาม (monitoring) การใช้ **CPU และ RAM** ของโปรเซสที่กำลังรันอยู่ เช่น **MATLAB** หรือ **Python script**  

โครงการนี้ออกแบบมาเพื่อช่วยเก็บข้อมูล resource consumption ระหว่างการเทรนโมเดลหรือรันโปรแกรมเชิงวิจัย โดยสามารถแสดงผลได้ทั้งแบบ **GUI (กราฟ)** และ **CLI (command line)** พร้อมบันทึกเป็น **Excel/CSV** เพื่อการวิเคราะห์ภายหลัง  

---

## ✨ Features

### 🔹 Core
- ตรวจจับ **MATLAB process** (ผ่าน PID file) หรือ **Python script** โดยอัตโนมัติ  
- Sampling rate ปรับได้ (0.1 – 10.0 วินาที)  
- จัดการ process monitoring อัตโนมัติ และหยุดเมื่อโปรเซสจบ  

### 🔹 Display Modes
- **Real-time display** – แสดงผล CPU/RAM ทุก sampling interval  
- **Buffered display** – รวมค่าเฉลี่ยตามช่วงเวลา เพื่อลดความหนาแน่นของ output  
- **Adaptive update interval** – ยืดช่วงการอัปเดตตามเวลาที่ผ่านไป (เช่น จากทุก 2 วินาที → 30 วินาทีเมื่อรันนานขึ้น)  

### 🔹 Export
- Export ข้อมูลการ monitor เป็น  
  - **Excel (.xlsx)** ผ่าน `openpyxl`  
  - **CSV (.csv)`  
- บันทึก **Source/Command** ที่ทำการตรวจจับไว้ท้ายไฟล์  

### 🔹 GUI (PyQt5 + Matplotlib)
- แสดงกราฟ **CPU (%) และ RAM (MB)** แบบ real-time หรือหลัง training เสร็จ  
- ปุ่มควบคุม: Reset table, Export Excel/CSV, Save graph (PNG), Exit  
- Table log ที่ sync กับกราฟ  

### 🔹 CLI (argparse + interactive)
- ใช้งานได้ 2 แบบ  
  1. **Interactive Mode** → ผู้ใช้ตอบคำถามทีละขั้น (sampling rate, display mode, export)  
  2. **CLI Argument Mode** → ใช้ flags เช่น `-s`, `-rt`, `-excel`, `-n`, `-end`  

---

## 📂 Repository Structure

Performance-Monitor/
│
├── CPU_RAM Monitor_GUI by psutil.py # GUI version 
├── CPU_RAM Monitor_CLI by psutil.py # CLI version 
└── README.md # Documentation
