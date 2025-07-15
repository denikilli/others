#-------------------------------------------------------------------------------
# Name:    Disk_Monitor-14
# Purpose: Displaying system information about working infrastructure
# Author:  Zhecheva.Yordanka
# Created: 08/07/2025
# Copyright: (c) Zhecheva.Yordanka 2025
# Licence:   <your licence>
#-------------------------------------------------------------------------------
import os
import wmi
import winrm
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import scrolledtext, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from datetime import datetime

def get_system_info_wmi(conn):
    try:
        os_info = conn.Win32_OperatingSystem()[0]
        cpu_info = conn.Win32_Processor()[0]
        comp_sys = conn.Win32_ComputerSystem()[0]

        os_name = os_info.Caption
        os_version = os_info.Version
        cpu_name = cpu_info.Name
        cpu_cores = cpu_info.NumberOfCores
        total_physical_memory_gb = int(comp_sys.TotalPhysicalMemory) / (1024**3)

        return {
            "OS Name": os_name,
            "OS Version": os_version,
            "CPU": cpu_name,
            "CPU Cores": cpu_cores,
            "Physical Memory GB": f"{total_physical_memory_gb:.2f}"
        }
    except Exception:
        return {
            "OS Name": "Unknown",
            "OS Version": "Unknown",
            "CPU": "Unknown",
            "CPU Cores": "Unknown",
            "Physical Memory GB": "Unknown"
        }

def get_logger_for_ip(ip):
    import logging
    logger = logging.getLogger(ip)
    if not logger.hasHandlers():
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
    return logger

def check_disks_wmi(server_ip, username, password, data_rows, log_output):
    logger = get_logger_for_ip(server_ip)
    try:
        conn = wmi.WMI(computer=server_ip, user=username, password=password)
        sys_info = get_system_info_wmi(conn)
        sys_info_lines = [
            f"OS Name: {sys_info['OS Name']}",
            f"OS Version: {sys_info['OS Version']}",
            f"CPU: {sys_info['CPU']}",
            f"CPU Cores: {sys_info['CPU Cores']}",
            f"Physical Memory(GB): {sys_info['Physical Memory GB']}",
        ]
        for line in sys_info_lines:
            logger.info(line)
            log_output.insert(tk.END, line + "\n")

        data_rows.append({
            "IP": server_ip,
            "Share": "System Info",
            "Drive": "",
            "Total GB": "",
            "Free GB": "",
            "Free %": "",
            "Status": "\n".join(sys_info_lines)
        })

        logical_disks = conn.Win32_LogicalDisk()
        if not logical_disks:
            msg = f"{server_ip} - Няма намерени логически дискове (WMI)"
            logger.warning(msg)
            log_output.insert(tk.END, msg + "\n")
            return

        for disk in logical_disks:
            if disk.DriveType not in [3, 4]:
                continue

            try:
                total_gb = int(disk.Size) / (1024**3)
                free_gb = int(disk.FreeSpace) / (1024**3)
                free_percent = (free_gb / total_gb) * 100 if total_gb > 0 else 0
                status = "OK" if free_percent > 25 else "LOW SPACE"
                row = {
                    "IP": server_ip,
                    "Share": disk.VolumeName if disk.VolumeName else "",
                    "Drive": disk.DeviceID,
                    "Total GB": f"{total_gb:.2f}",
                    "Free GB": f"{free_gb:.2f}",
                    "Free %": f"{free_percent:.2f}%",
                    "Status": status
                }
                data_rows.append(row)
                logger.info(f"{server_ip} - {disk.DeviceID} {status} - Free: {free_percent:.2f}%")
                log_output.insert(tk.END, f"{server_ip} - {disk.DeviceID} {status} - Free: {free_percent:.2f}%\n")
            except Exception as e:
                logger.error(f"{server_ip} - Error processing disk {disk.DeviceID}: {e}")
                log_output.insert(tk.END, f"{server_ip} - Error processing disk {disk.DeviceID}: {e}\n")
    except Exception as e:
        logger.error(f"{server_ip} - WMI error: {e}")
        log_output.insert(tk.END, f"{server_ip} - WMI error: {e}\n")

def generate_pdf(data_rows, output_path, server_ip):
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter
    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(colors.red)
    c.drawString(50, height - 40, f"Disk Monitoring Report for Server: {server_ip}")
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10)
    c.drawString(50, height - 60, f"Date: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}")
    y = height - 90

    sys_info_rows = [row for row in data_rows if row.get("Share") == "System Info"]
    for sys_row in sys_info_rows:
        lines = sys_row.get("Status", "").split("\n")
        for line in lines:
            c.drawString(50, y, line)
            y -= 25

    headers = ["IP", "Share", "Drive", "Total GB", "Free GB", "Free %", "Status"]
    col_widths = [80, 100, 50, 70, 70, 60, 60]

    c.setFont("Helvetica-Bold", 10)
    c.drawString(50, y, headers[0])
    for i, header in enumerate(headers[1:], 1):
        c.drawString(50 + sum(col_widths[:i]), y, header)
    y -= 25

    c.setFont("Helvetica", 10)
    for row in data_rows:
        if row.get("Share") == "System Info":
            continue
        if y < 50:
            c.showPage()
            y = height - 50
            c.setFont("Helvetica", 10)
        for i, key in enumerate(headers):
            c.drawString(50 + sum(col_widths[:i]), y, str(row.get(key, "")))
        y -= 25

    c.save()

def send_email_alert(data_rows, ip, log_output):
    if not (smtp_user and recipient_email): #and smtp_password
        log_output.insert(tk.END, "Имейл настройките не са попълнени. Имейл няма да бъде изпратен.\n")
        return

    low_space_disks = [
        row for row in data_rows
        if row.get("Free %") and row.get("Free %") != "" and float(row["Free %"].strip('%')) < 25
    ]
    if not low_space_disks:
        log_output.insert(tk.END, "Няма дискове с под 25% свободно място. Имейл няма да бъде изпратен.\n")
        return

    subject = f"Disk Monitor ALERT: Low Disk Space on {ip}"
    body = "Следните дискове са под 25% свободно място:\n\n"

    for row in low_space_disks:
        body += f"{row['IP']} - {row['Drive']} ({row['Free %']} свободно) - Статус: {row['Status']}\n"

    message = MIMEMultipart()
    message["From"] = smtp_user
    message["To"] = recipient_email
    message["Subject"] = subject

    message.attach(MIMEText(body, "plain"))

    # Прикачване на PDF
    import os
    from email.mime.base import MIMEBase
    from email import encoders

    reports_dir = os.path.join(os.getcwd(), "reports")
    pdf_filename = f"disk_report_{ip.replace('.', '_')}.pdf"
    pdf_path = os.path.join(reports_dir, pdf_filename)

    if os.path.exists(pdf_path):
        with open(pdf_path, "rb") as f:
            part = MIMEBase("application", "pdf")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={pdf_filename}",
        )
        message.attach(part)
    else:
        log_output.insert(tk.END, f"Внимание: PDF файлът {pdf_filename} не беше намерен и не беше прикачен.\n")

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.sendmail(smtp_user, recipient_email, message.as_string())
        server.quit()
        log_output.insert(tk.END, "Изпратен имейл за предупреждение при ниско дисково пространство.\n")
    except Exception as e:
        log_output.insert(tk.END, f"Грешка при изпращане на имейл: {e}\n")

def run_monitor():
    ips = entry_ip.get().strip().split(',')
    username = entry_user.get().strip()
    password = entry_pass.get().strip()

    global smtp_user, smtp_password, recipient_email, smtp_server, smtp_port
    smtp_user = entry_email_user.get().strip()
    smtp_password = entry_email_pass.get().strip()
    recipient_email = entry_email_recipient.get().strip()
    smtp_server = entry_smtp_server.get().strip()
    smtp_port = int(entry_smtp_port.get().strip())

    if not ips or not username or not password:
        messagebox.showerror("Грешка", "Моля, попълнете всички полета.")
        return

    log_output.delete(1.0, tk.END)
    for ip in ips:
        ip = ip.strip()
        if not ip:
            continue
        data_rows = []
        try:
            check_disks_wmi(ip, username, password, data_rows, log_output)
        except Exception as e:
            log_output.insert(tk.END, f"WMI check failed: {e}\n")

        send_email_alert(data_rows, ip, log_output)

        reports_dir = os.path.join(os.getcwd(), "reports")
        os.makedirs(reports_dir, exist_ok=True)
        pdf_path = os.path.join(reports_dir, f"disk_report_{ip.replace('.', '_')}.pdf")
        generate_pdf(data_rows, pdf_path, ip)

    messagebox.showinfo("Готово", f"Репортите са генерирани в: {reports_dir}")

root = tk.Tk()
root.title("Disk Monitor")

# Credentials and target
tk.Label(root, text="IP адрес(и):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_ip = tk.Entry(root, width=50)
entry_ip.grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="Потребител:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
entry_user = tk.Entry(root, width=50)
entry_user.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Парола:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry_pass = tk.Entry(root, width=50, show="*")
entry_pass.grid(row=2, column=1, padx=5, pady=5)

# Email configuration with fixed default values and disabled
tk.Label(root, text="Имейл (изпращач):").grid(row=3, column=0, padx=5, pady=5, sticky="e")
entry_email_user = tk.Entry(root, width=50)
entry_email_user.insert(0, "Disk_Monitor@neftochim.bg")  # по подразбиране
entry_email_user.config(state='disabled')
entry_email_user.grid(row=3, column=1, padx=5, pady=5)

tk.Label(root, text="Имейл парола:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
entry_email_pass = tk.Entry(root, width=50, show="*")
entry_email_pass.config(state='disabled')
entry_email_pass.grid(row=4, column=1, padx=5, pady=5)

tk.Label(root, text="Имейл (получател):").grid(row=5, column=0, padx=5, pady=5, sticky="e")
entry_email_recipient = tk.Entry(root, width=50)
entry_email_recipient.insert(0, "mail@yourdomain.bg")  # по подразбиране
entry_email_recipient.config(state='disabled')
entry_email_recipient.grid(row=5, column=1, padx=5, pady=5)

tk.Label(root, text="SMTP сървър:").grid(row=6, column=0, padx=5, pady=5, sticky="e")
entry_smtp_server = tk.Entry(root, width=50)
entry_smtp_server.insert(0, "mail.yourdomain.bg")  # по подразбиране Exchange
entry_smtp_server.config(state='disabled')
entry_smtp_server.grid(row=6, column=1, padx=5, pady=5)

tk.Label(root, text="SMTP порт:").grid(row=7, column=0, padx=5, pady=5, sticky="e")
entry_smtp_port = tk.Entry(root, width=50)
entry_smtp_port.insert(0, "25")  # порт 25
entry_smtp_port.config(state='disabled')
entry_smtp_port.grid(row=7, column=1, padx=5, pady=5)

def enable_email_fields():
    entry_email_user.config(state='normal')
    entry_email_pass.config(state='normal')
    entry_email_recipient.config(state='normal')
    entry_smtp_server.config(state='normal')
    entry_smtp_port.config(state='normal')
    btn_edit_email.config(state='disabled')

btn_edit_email = tk.Button(root, text="Редактирай имейл настройки", command=enable_email_fields)
btn_edit_email.grid(row=8, column=0, columnspan=2, pady=5)

log_output = scrolledtext.ScrolledText(root, width=80, height=20)
log_output.grid(row=9, column=0, columnspan=2, padx=10, pady=10)

btn_run = tk.Button(root, text="Стартирай", command=run_monitor)
btn_run.grid(row=10, column=0, columnspan=2, pady=10)

# Глобални променливи за имейл (ще се попълват при Стартирай)
smtp_user = ""
smtp_password = ""
recipient_email = ""
smtp_server = ""
smtp_port = 25

root.mainloop()
