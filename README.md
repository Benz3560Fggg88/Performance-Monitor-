<div align="center">

# 📊 Performance Monitor

**เครื่องมือสำหรับติดตามการใช้งาน CPU & RAM ของโปรเซสเทรนนิ่ง (Python & MATLAB) ที่ออกแบบมาให้ใช้งานง่ายและทรงพลัง**

</div>

<p align="center">
  <img alt="Python Version" src="https://img.shields.io/badge/python-3.8%2B-blue?style=for-the-badge&logo=python">
  <img alt="License" src="https://img.shields.io/badge/license-MIT-green?style=for-the-badge">
  <img alt="Code Style" src="https://img.shields.io/badge/code%20style-black-black?style=for-the-badge">
</p>

---

โปรเจกต์นี้ประกอบด้วยเครื่องมือ 2 เวอร์ชันที่ตอบโจทย์ทุกการใช้งาน ตั้งแต่การแสดงผลแบบกราฟิกสวยงาม ไปจนถึงการทำงานอัตโนมัติผ่าน Command-Line

| GUI Version | CLI Version |
| :---: | :---: |
|  |  |
| **หน้าจอสวยงาม, ใช้งานง่าย, เห็นข้อมูลครบ** | **ยืดหยุ่นสูง, ทำงานอัตโนมัติ, ใช้ทรัพยากรน้อย** |

## ✅ คุณสมบัติเด่น (Key Features)

-   **Auto-Detection:** ตรวจจับโปรเซส `MATLAB` หรือ `Python` ที่กำลังเทรนโดยอัตโนมัติ
-   **Dual Interface:** เลือกใช้ได้ทั้งแบบ **GUI** ที่มีกราฟและตาราง หรือ **CLI** สำหรับการทำงานบนเซิร์ฟเวอร์
-   **Flexible Display:** โหมด CLI สามารถแสดงผลได้ทั้งแบบ **Real-time** และ **Buffered** เพื่อลดภาระหน้าจอ
-   **Data Export:** บันทึกผลลัพธ์การติดตามทั้งหมดเป็นไฟล์ **Excel (.xlsx)** หรือ **CSV (.csv)** ได้อย่างง่ายดาย
-   **Graph Snapshot:** เวอร์ชัน GUI สามารถบันทึกภาพกราฟเป็นไฟล์ **PNG** คุณภาพสูงได้

---

## ⚙️ เริ่มต้นใช้งาน (Getting Started)

### ข้อกำหนดเบื้องต้น

-   Python 3.8 หรือสูงกว่า
-   `pip` (Python package installer)

### การติดตั้ง

เปิด Terminal หรือ Command Prompt แล้วรันคำสั่งเดียวเพื่อติดตั้งไลบรารีที่จำเป็นทั้งหมด:

```bash
pip install psutil openpyxl PyQt5 matplotlib
```
### 
Download CPU_RAM Monitor by psutil : [here](https://github.com/Benz3560Fggg88/Performance-Monitor-/releases/tag/v1.0.0)
---

## 💡 คู่มือการใช้งาน (Usage Guide)

### 🖥️ GUI Version

1.  **รันโปรแกรม:**
    ```bash
    python "CPU_RAM Monitor_GUI by psutil.py"
    ```
2.  **เปิดโหมดตรวจจับ:** ติ๊กที่ช่อง **"Start Detection Automatically"**
3.  **เริ่มงานของคุณ:** โปรแกรมจะรอและเริ่มบันทึกข้อมูลทันทีที่ตรวจพบโปรเซสเป้าหมาย
4.  **จัดการข้อมูล:** เมื่อการเทรนสิ้นสุดลง คุณสามารถ Export ข้อมูลหรือบันทึกกราฟได้จากปุ่มบนหน้าจอ

### ⌨️ CLI Version

รันโปรแกรมผ่าน Terminal โดยสามารถเลือกได้ 2 รูปแบบ:

1.  **โหมดโต้ตอบ (Interactive Mode):**
    เพียงรันสคริปต์โดยไม่ใส่ Argument โปรแกรมจะถามการตั้งค่าทีละขั้นตอน
    ```bash
    python "CPU_RAM Monitor_CLI by psutil.py"
    ```

2.  **โหมด Argument (CLI Mode):**
    ควบคุมทุกอย่างผ่านคำสั่งเดียว เหมาะสำหรับสร้างสคริปต์อัตโนมัติ
    ```bash
    # ตัวอย่าง: ติดตามทุก 0.5 วินาที, แสดงผลแบบ Buffered, และบันทึกเป็นไฟล์ Excel ชื่อ `resnet_log`
    python "CPU_RAM Monitor_CLI by psutil.py" -s 0.5 -bf -excel -n resnet_log
    ```

**ตาราง Arguments:**

| Argument | Alias | รายละเอียด |
| :--- | :--- | :--- |
| `-s` | | **Sampling Rate** (วินาที) |
| `-rt` | | โหมดแสดงผลแบบ **Real-time** |
| `-bf` | | โหมดแสดงผลแบบ **Buffered** |
| `-excel` | | **Export to Excel** หลังจบการทำงาน |
| `-csv` | | **Export to CSV** หลังจบการทำงาน |
| `-n` | | **ชื่อไฟล์** สำหรับ Export (ไม่ต้องใส่นามสกุล) |
| `-end` | | **จบการทำงาน** ทันทีหลัง Export |

---

## 🔗 การเชื่อมต่อกับ MATLAB (MATLAB Integration)

> **สำคัญ:** เพื่อให้โปรแกรมสามารถตรวจจับโปรเซส MATLAB ได้ คุณจำเป็นต้องเพิ่มโค้ด `.m` บางส่วนในสคริปต์ของคุณ เพื่อสร้างไฟล์ `C:\temp\training_pid.txt` สำหรับให้โปรแกรม Python อ่าน

<details>
<summary><strong>คลิกเพื่อดูโค้ด: สำหรับวางไว้ "ก่อน" เริ่มการประมวลผล</strong></summary>

```matlab
% ---------- MATLAB: Start Detection ----------
pid = feature('getpid');  % Get MATLAB's own Process ID
fid = fopen('C:\temp\training_pid.txt', 'w');
if fid == -1
    error('Cannot open C:\temp\training_pid.txt for writing.');
end
fprintf(fid, '%d\n', pid);
fclose(fid);
% -------------------------------------------
```

</details>

<details>
<summary><strong>คลิกเพื่อดูโค้ด: สำหรับวางไว้ "หลัง" จบการประมวลผล</strong></summary>

```matlab
% ---------- MATLAB: End Detection ----------
pause(1);  % Allow a moment for the monitor to catch up
if exist('C:\temp\training_pid.txt', 'file')
    delete('C:\temp\training_pid.txt');
    fprintf('PID file deleted successfully.\n');
end
% -----------------------------------------
```

</details>

---

## 📜 สิทธิ์การใช้งาน (License)

โปรเจกต์นี้อยู่ภายใต้สิทธิ์การใช้งานแบบ **MIT License**
