from flask import Flask, request, jsonify
from flask_cors import CORS
import socket
import logging
from datetime import datetime
from urllib.parse import parse_qs, urlparse
import pystray
from PIL import Image, ImageDraw, ImageFont
import threading
import os
import sys
import winreg
import ctypes
import requests
import urllib.parse

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('printer_server2.log'),
        logging.StreamHandler()
    ]
)

app = Flask(__name__)
CORS(app)
server_running = True
connection_type = "IP"

TEMPLATES = {
    '5': {'size': '62mm', 'dpi': '300x300'},  # Network QL series
    '1': {'size': '36mm', 'dpi': '300x300'},  # USB PT-9xx series
    '2': {'size': '24mm', 'dpi': '300x300'}   # USB PT-7xx and/or PT-9xx series
}

def create_label_data(template_format, serial_number, model, model_apn, type_name, username, date):
    return (f"%-12345X@PJL ENTER LANGUAGE=PCL\n"
            f"!12355FQL-820NWB\n"
            f"!1B 1D 72 {template_format}\n"
            # SN text and barcode
            f"^FO20,20^A0N,30,30^FDSN:^FS\n"
            f"^FO100,20^BCN,100,Y,N,N^FD{serial_number}^FS\n"
            # Model heading
            f"^FO20,150^A0N,30,30^FDModel^FS\n"
            f"^FO20,190^A0N,35,35^FD{model}^FS\n"
            # APN heading and value
            f"^FO20,240^A0N,30,30^FDAPN^FS\n"
            f"^FO20,280^A0N,35,35^FD{model_apn}^FS\n"
            # QR code on the right
            f"^FO450,240^BQN,2,4^FD{serial_number}^FS\n"
            # Type heading and value
            f"^FO20,330^A0N,30,30^FDType:^FS\n"
            f"^FO20,370^A0N,35,35^FD{type_name}^FS\n"
            f"^XZ")

def send_to_network_printer(printer_ip, data):
    try:
        printer_port = 9100
        if ':' in printer_ip:
            printer_ip, port = printer_ip.split(':')
            printer_port = int(port)

        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(5)
            s.connect((printer_ip, printer_port))
            logging.info(f"Connection successful to {printer_ip}:{printer_port}")
            
            s.sendall(data.encode('utf-8'))
            logging.info(f"Sentent data: {data}")
            
            logging.info("Print job sent successfully")
            return True
            
    except socket.timeout:
        logging.error(f"Connection timeout to printer at {printer_ip}")
        return False
    except Exception as e:
        logging.error(f"Failed to send data to printer at {printer_ip}: {str(e)}")
        return False

@app.route('/', methods=['GET'])
def health_check():
    return jsonify({"status": "running", "timestamp": datetime.now().isoformat()})

@app.route('/scan', methods=['GET'])
def scan_barcode():
    try:
        format_type = request.args.get('FORMAT', '5')
        barcode_data = request.args.get('DATA')
        
        if not barcode_data:
            return jsonify({"status": "error", "message": "No barcode data provided"}), 400
            
        logging.info(f"Received scan request: FORMAT={format_type}, DATA={barcode_data}")
        
        return jsonify({
            "status": "success",
            "data": barcode_data,
            "format": format_type
        })
            
    except Exception as e:
        logging.error(f"Error processing scan request: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/label', methods=['GET'])
def label():
    try:
        printer_connection = request.args.get('PRINTER', '')
        serial_number = request.args.get('SN')
        model = request.args.get('MODEL', '')
        model_apn = request.args.get('Model_APN', '')
        type_name = request.args.get('TYPE', '')
        mode = request.args.get('MODE', 'IP')
        username = request.args.get('USER', 'Unknown')
        date = datetime.now().strftime('%Y-%m-%d')
        
        # Determine format type based on connection and printer model
        if mode == 'IP':
            format_type = '5'  # 62mm for QL network printers
        else:
            # USB mode - check printer model
            if any(model in printer_connection for model in ['PT-700', 'PT-750','PT-900', 'PT-950']):
                format_type = '2'  # 24mm for PT-7xx or PT-9xx series
            elif any(model in printer_connection for model in ['PT-900', 'PT-950']):
                format_type = '1'  # 36mm for PT-9xx series
            else:
                format_type = '5'  # Default to 62mm for QL series
        
        logging.info(f"Received print request: PRINTER={printer_connection}, SN={serial_number}, MODEL={model}, Model_APN={model_apn}, TYPE={type_name}, MODE={mode}, FORMAT={format_type}")
        
        if not serial_number:
            return jsonify({"status": "error", "message": "No serial number provided"}), 400
            
        data = create_label_data(format_type, serial_number, model, model_apn, type_name, username, date)
        
        if mode == 'IP':
            success = send_to_network_printer(printer_connection, data)
        else:
            printer_name = printer_connection.replace('USB:', '')
            success = sendend_to_network_printer(printer_name, data)
            
        if success:
            return jsonify({"status": "success"})
        return jsonify({"status": "error", "message": f"Failed to send print job to {mode} printer"}), 500
            
    except Exception as e:
        logging.error(f"Error processing print job: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/getLabelSize', methods=['GET'])
def get_label_size():
    try:
        mode = request.args.get('MODE', 'IP')
        format_type = request.args.get('FORMAT', '5')
        template = TEMPLATES.get(format_type, TEMPLATES['5'])
        return jsonify({"labelSize": template['size']})
    except Exception as e:
        logging.error(f"Error getting label size: {str(e)}")
        return jsonify({"status": "error", "message": f"Failed to get label size: {str(e)}"}), 500

def create_square_icon():
    icon_path = r"C:\Users\YOUR_NAME\Desktop\Projects\printer-app\app\icons\icon128.png"
    try:
        if os.path.exists(icon_path):
            return Image.open(icon_path)
    except Exception as e:
        logging.error(f"Failed to load icon: {str(e)}")
    
    # Fallback to creating a default icon if the image fails to load
    size = (64, 64)
    icon = Image.new('RGBA', size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(icon)
    square_bbox = [2, 2, 62, 62]
    draw.rectangle(square_bbox, fill='#31465e', outline=None)
    try:
        font = ImageFont.truetype("seguiemj.ttf", 40)
    except:
        font = ImageFont.load_default()
    text = "🖨️"
    left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
    text_width = right - left
    text_height = bottom - top
    x = (size[0] - text_width) // 2
    y = (size[1] - text_height) // 2
    draw.text((x, y), text, fill="white", font=font)
    return icon

def toggle_server():
    global server_running
    server_running = not server_running
    return "Stop Server" if server_running else "Start Server"

def show_info():
    ctypes.windll.user32.MessageBoxW(0, 
        "Printer Server v3.0\nStatus: Running\nNetwork QL Series: 62mm\nUSB PT-9xx Series: 36mm\nUSB PT-7xx Series: 24mm", 
        "Printer Server Info", 0)

def show_instructions():
    ctypes.windll.user32.MessageBoxW(0,
        "Printer Setup Guide:\n\n"
        "For detailed instructions visit:\n"
        "web_site_suport\n\n"
        "Quick Setup:\n\n"
        "IP Mode (Network QL Series):\n"
        "1. Windows Settings > Devices > Printers & scanners\n"
        "2. Add printer > Add a printer manually\n"
        "3. Add TCP/IP printer > Enter printer IP\n"
        "4. Select 'Brother QL-820NWB' driver\n\n"
        "USB Mode (PT Series):\n"
        "1. Connect printer via USB\n"
        "2. Visit brother.com/inst\n"
        "3. Download & install P-touch driver\n"
        "4. Windows will detect printer automatically\n\n"
        "For support: [your_email]",
        "Printer Setup Instructions", 0)

def create_system_tray():
    icon_image = create_square_icon()
    menu = (
        pystray.MenuItem("Info", show_info),
        pystray.MenuItem("Server Status", toggle_server),
        pystray.MenuItem("Instructions", show_instructions),
        pystray.MenuItem("Exit", lambda: icon.stop())
    )
    icon = pystray.Icon("Printer Server", icon_image, "Printer Server", menu)
    return icon

def add_to_startup():
    try:
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, 
                            winreg.KEY_ALL_ACCESS)
        winreg.SetValueEx(key, "PrinterServer2", 0, winreg.REG_SZ, sys.executable)
        winreg.CloseKey(key)
        logging.info("Successfully added to startup")
    except Exception as e:
        logging.error(f"Failed to add to startup: {str(e)}")
        pass

def run_flask():
    app.run(host='0.0.0.0', port=3000)

if __name__ == '__main__':
    try:
        add_to_startup()
        icon = create_system_tray()
        flask_thread = threading.Thread(target=run_flask, daemon=True)
        flask_thread.start()
        icon.run()
    except Exception as e:
        logging.error(f"Failed to start server: {str(e)}")
