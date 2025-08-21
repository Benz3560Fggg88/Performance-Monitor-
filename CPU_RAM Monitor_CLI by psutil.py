import time
import psutil
import csv
import os
from openpyxl import Workbook
from datetime import datetime, timedelta
import argparse
import sys

def get_pid():
    """
    ตรวจหา PID ของโปรเซสที่กำลังเทรน
    - ตรวจหาจากไฟล์ C:\\temp\\training_pid.txt สำหรับ MATLAB ก่อน
    - หากไม่เจอ จะค้นหาโปรเซส Python ที่กำลังรันไฟล์ .py
    """
    # --- ตรวจสอบ MATLAB ก่อน ---
    pid_file_path = "C:\\temp\\training_pid.txt"
    try:
        if os.path.exists(pid_file_path):
            with open(pid_file_path, "r") as f:
                pid = int(f.read().strip())
            proc = psutil.Process(pid)
            if proc.is_running() and "matlab" in proc.name().lower():
                return pid, f"MATLAB (PID: {pid}) CMD: {' '.join(proc.cmdline())}"
    except (FileNotFoundError, psutil.NoSuchProcess, ValueError, psutil.AccessDenied):
        pass # หากมีปัญหา ให้ข้ามไปหา Python

    # --- หากไม่เจอ MATLAB ให้ตรวจสอบ Python ---
    current_pid = psutil.Process().pid
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            if proc.pid == current_pid:
                continue

            name = proc.info['name'].lower()
            cmdline_list = proc.info.get('cmdline')
            cmdline = ' '.join(cmdline_list or []).lower()

            if ("python" in name or "python.exe" in name) and ".py" in cmdline:
                return proc.pid, f"Python: {cmdline}"
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    return None, None

def get_update_interval(elapsed):
    """คำนวณช่วงเวลาการแสดงผลแบบ Buffered ตามเวลาที่ผ่านไป"""
    if elapsed <= 10: return 10
    elif elapsed <= 20: return 2
    elif elapsed <= 60: return 5
    elif elapsed <= 300: return 10
    elif elapsed <= 900: return 20
    else: return 30

def monitor(samrate, display_mode):
    """
    ฟังก์ชันหลักสำหรับติดตามและบันทึกข้อมูล CPU/RAM
    """
    print("🔍 Waiting for training process...")
    pid_file_path = "C:\\temp\\training_pid.txt"

    while True:
        pid, source = get_pid()
        if pid:
            break
        time.sleep(1)

    print(f"\n✅ Detected training from: {source}")
    print(f"{'Elapsed Time':<15} {'CPU (%)':<10} {'RAM (MB)':<12} Source")

    training_start = time.time()
    last_display_time = training_start
    data, buffer, samples = [], [], []
    is_matlab = "matlab" in source.lower()

    while True:
        # --- เงื่อนไขการหยุด Monitor ---
        if is_matlab and not os.path.exists(pid_file_path):
            break
        if not psutil.pid_exists(pid):
            print("\nℹ️ Process PID not found. Stopping.")
            break

        # --- เก็บข้อมูล CPU/RAM ---
        try:
            proc = psutil.Process(pid)
            proc.cpu_percent(interval=None)
            time.sleep(0.1)
            cpu = proc.cpu_percent(interval=None) / psutil.cpu_count()
            ram = proc.memory_info().rss / (1024 * 1024)
        except psutil.NoSuchProcess:
            break

        # --- ประมวลผลและแสดงข้อมูล ---
        samples.append((cpu, ram))
        if (len(samples) * 0.1) >= samrate:
            avg_cpu = sum(x[0] for x in samples) / len(samples) if samples else 0
            avg_ram = sum(x[1] for x in samples) / len(samples) if samples else 0
            samples.clear()
            
            elapsed_seconds = time.time() - training_start
            td_str = str(timedelta(seconds=elapsed_seconds))
            try:
                time_part, ms_part = td_str.split('.')
                timestamp = f"{time_part}.{ms_part[:3]}"
            except ValueError:
                timestamp = f"{td_str}.000"

            row = (timestamp, avg_cpu, avg_ram, source)

            if display_mode == 1:
                print(f"{timestamp:<15} {avg_cpu:<10.2f} {avg_ram:<12.2f} {source}")
                data.append(row)
            else:
                buffer.append(row)
                if time.time() - last_display_time >= get_update_interval(time.time() - training_start):
                    for b in buffer:
                        print(f"{b[0]:<15} {b[1]:<10.2f} {b[2]:<12.2f} {b[3]}")
                    data.extend(buffer)
                    buffer.clear()
                    last_display_time = time.time()

    if display_mode == 2 and buffer:
        for b in buffer:
            print(f"{b[0]:<15} {b[1]:<10.2f} {b[2]:<12.2f} {b[3]}")
        data.extend(buffer)

    print("\n⏹️ Training stopped.")
    return data, source

def export_excel(data, source, filename=None):
    """ส่งออกข้อมูลเป็นไฟล์ Excel"""
    if not filename:
        filename = input("Enter Excel filename (without extension): ").strip()
    if not filename:
        filename = f"monitor_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    full_filename = f"{filename}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Monitoring_Log"
    ws.append(["Elapsed Time", "CPU (%)", "RAM (MB)", "Source"])
    for row in data:
        ws.append(row)
    ws.append([])
    ws.append(["Command/Source:", source])
    try:
        wb.save(full_filename)
        print(f"📁 Saved Excel to {os.path.abspath(full_filename)}")
    except Exception as e:
        print(f"❌ Error saving Excel file: {e}")

def export_csv(data, source, filename=None):
    """ส่งออกข้อมูลเป็นไฟล์ CSV"""
    if not filename:
        filename = input("Enter CSV filename (without extension): ").strip()
    if not filename:
        filename = f"monitor_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    full_filename = f"{filename}.csv"
    try:
        with open(full_filename, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["Elapsed Time", "CPU (%)", "RAM (MB)", "Source"])
            writer.writerows(data)
            writer.writerow([])
            writer.writerow(["Command/Source:", source])
        print(f"📁 Saved CSV to {os.path.abspath(full_filename)}")
    except Exception as e:
        print(f"❌ Error saving CSV file: {e}")

def main_cli(args):
    """ฟังก์ชันสำหรับโหมด CLI"""
    s = args.s
    mode = 1 if args.rt else 2
    records, source = monitor(s, mode)

    # --- จัดการ Export ---
    if args.excel:
        export_excel(records, source, args.n)
    elif args.csv:
        export_csv(records, source, args.n)

    # --- จบการทำงานถ้ามี -end ---
    if args.end:
        print("👋 Exiting as requested by -end flag.")
        return
        
    # --- เมนูหลังจบการทำงาน (ถ้าไม่มี -end) ---
    while True:
        print("\n✅ Monitoring finished. What next?")
        print("1. Wait for new training")
        print("2. Export to Excel")
        print("3. Export to CSV")
        print("4. Restart from beginning")
        print("5. Exit")
        post = input("Choice: ").strip()
        if post == '1':
            print("\n" + "-"*40 + "\n")
            records, source = monitor(s, mode)
            continue
        elif post == '2':
            export_excel(records, source)
        elif post == '3':
            export_csv(records, source)
        elif post == '4':
            # -----[ จุดที่แก้ไข ]-----
            print("\n" + "="*40 + "\n")
            main_interactive() # เรียกโหมด Interactive เพื่อเริ่มใหม่ทั้งหมด
            return
        elif post == '5':
            print("👋 Exiting...")
            return
        else:
            print("❌ Invalid choice.")

def main_interactive(prefilled_s=None):
    """ฟังก์ชันสำหรับโหมด Interactive (โต้ตอบกับผู้ใช้)"""
    s = 0.0
    if prefilled_s:
        s = prefilled_s
        print(f"⏱️ Sampling rate set to {s} sec via command line.")
    else:
        while True:
            try:
                s_input = input("⏱️ Set sampling rate (0.1–10.0) sec (recommended: 1.0): ")
                s = float(s_input)
                if 0.1 <= s <= 10.0: break
                else: print("❌ Invalid range. Try again.")
            except ValueError:
                print("❌ Invalid input. Try again.")

    while True: # Display mode loop
        print("\n📺 Select display mode:")
        print("1. Real-time display")
        print("2. Buffered display")
        print("3. Back to sampling rate")
        m = input("Choice: ").strip()

        if m == '3':
            main_interactive()
            return

        if m in ['1', '2']:
            mode = int(m)
            while True: # Action loop
                print("\n▶️ Select action:")
                print("1. Wait for training detection")
                print("2. Back to display mode selection")
                action = input("Choice: ").strip()

                if action == '2':
                    break 

                if action == '1':
                    records, source = monitor(s, mode)
                    while True: # Post-monitoring loop
                        print("\n✅ Monitoring finished. What next?")
                        print("1. Wait for new training")
                        print("2. Export to Excel")
                        print("3. Export to CSV")
                        print("4. Restart from beginning")
                        print("5. Exit")
                        post = input("Choice: ").strip()

                        if post == '1':
                            print("\n" + "-"*40 + "\n")
                            records, source = monitor(s, mode)
                            continue
                        elif post == '2':
                            export_excel(records, source)
                        elif post == '3':
                            export_csv(records, source)
                        elif post == '4':
                            print("\n" + "="*40 + "\n")
                            main_interactive()
                            return 
                        elif post == '5':
                            print("👋 Exiting...")
                            return
                        else:
                            print("❌ Invalid choice.")
                else:
                    print("❌ Invalid choice.")
        else:
            print("❌ Invalid choice.")


def main():
    """ฟังก์ชันหลักในการควบคุมโปรแกรมและจัดการ CLI arguments"""
    parser = argparse.ArgumentParser(
        description="Monitor training process CPU/RAM usage.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("-s", type=float, help="Set sampling rate (range: 0.1–10.0 sec).") 
    group_mode = parser.add_mutually_exclusive_group()
    group_mode.add_argument("-rt", action="store_true", help="Use real-time display mode.")
    group_mode.add_argument("-bf", action="store_true", help="Use buffered display mode.")
    group_export = parser.add_mutually_exclusive_group()
    group_export.add_argument("-excel", action="store_true", help="Export to Excel after monitoring.")
    group_export.add_argument("-csv", action="store_true", help="Export to CSV after monitoring.")
    parser.add_argument("-n", type=str, help="Filename for the export (without extension).")
    parser.add_argument("-end", action="store_true", help="End the program after monitoring and saving.")
    
    if len(sys.argv) == 1:
        main_interactive()
        return

    args, unknown = parser.parse_known_args()

    if unknown:
        print(f"\n❌ Error: Unrecognized arguments: {' '.join(unknown)}")
        print("Here are the valid options:\n")
        parser.print_help()
        print("\n👋 Exiting.")
        return
        
    if args.n and not (args.excel or args.csv):
        print("\n❌ Error: The -n argument can only be used with an export flag (-excel or -csv).")
        print("Here are the valid options:\n")
        parser.print_help()
        print("\n👋 Exiting.")
        return

    if args.s is not None and (args.rt or args.bf):
        if not (0.1 <= args.s <= 10.0):
            print("\n❌ Error: Sampling rate (-s) must be between 0.1 and 10.0.")
            print("Here are the valid options:\n")
            parser.print_help()
            print("\n👋 Exiting.")
            return
        main_cli(args)
        return

    if args.s is not None and not any([args.rt, args.bf, args.excel, args.csv, args.n, args.end]):
        if not (0.1 <= args.s <= 10.0):
            print("\n❌ Error: Sampling rate (-s) must be between 0.1 and 10.0.")
            print("Here are the valid options:\n")
            parser.print_help()
            print("\n👋 Exiting.")
            return
        main_interactive(prefilled_s=args.s)
        return

    print("\n❌ Error: For CLI mode, both sampling rate (-s) and display mode (-rt or -bf) are required.")
    print("Here are the valid options:\n")
    parser.print_help()
    print("\n👋 Exiting.")


if __name__ == "__main__":
    main()