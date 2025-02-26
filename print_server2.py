from flask import Flask, request, jsonify
from flask_cors import CORS
import socket
from brother_ql.raster import BrotherQLRaster
from brother_ql.backends.helpers import send
from brother_ql.backends import backend_factory
import logging
from datetime import datetime
from urllib.parse import parse_qs, urlparse
import pystray
from PIL import Image, ImageDraw, ImageFont
import threading
import os
import sys
import winreg

# Set up logging
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

# Template configurations
TEMPLATES = {
    '2': {'size': '24mm', 'dpi': '180x180'},  # For PT-P7* series (USB only)
    '3': {'size': '36mm', 'dpi': '300x300'},
    '4': {'size': 'A4', 'dpi': '300x300'},
    '6': {'size': '62mm', 'dpi': '300x300'}
}

def create_label_data(template_format, serial_number, model, model_apn, type_name, username, date):
    if template_format == '2':  # 24mm (PT-P7* series)
        return (f"^XA^FO20,10^A0N,20,20^FD{serial_number}^FS"
                f"^FO20,35^BCN,40,N,N,N^FD{serial_number}^FS"
                f"^FO20,80^A0N,15,15^FD{model}^FS^XZ")
    elif template_format == '3':  # 36mm
        return (f"^XA^FO30,10^A0N,25,25^FD{serial_number}^FS"
                f"^FO30,40^BCN,60,N,N,N^FD{serial_number}^FS"
                f"^FO30,110^A0N,20,20^FD{model}^FS^XZ")
    elif template_format == '4':  # A4
        return (f"^XA^FO100,100^A0N,70,70^FD{serial_number}^FS"
                f"^FO100,200^BCN,150,Y,N,N^FD{serial_number}^FS"
                f"^FO100,400^A0N,40,40^FDModel: {model}^FS"
                f"^FO100,450^A0N,40,40^FDAPN: {model_apn}^FS"
                f"^FO100,500^A0N,40,40^FDType: {type_name}^FS"
                f"^FO100,550^A0N,30,30^FDPrinted by: {username}^FS"
                f"^FO100,600^A0N,30,30^FDDate: {date}^FS^XZ")
    else:  # format 6 - 62mm
        return (f"^XA^FO50,50^A0N,50,50^FD{serial_number}^FS"
                f"^FO50,120^BCN,100,Y,N,N^FD{serial_number}^FS"
                f"^FO50,250^A0N,30,30^FDModel: {model}^FS"
                f"^FO50,290^A0N,30,30^FDAPN: {model_apn}^FS"
                f"^FO50,330^A0N,30,30^FDType: {type_name}^FS"
                f"^FO50,370^A0N,20,20^FDPrinted by: {username}^FS"
                f"^FO50,400^A0N,20,20^FDDate: {date}^FS^XZ")

def send_to_network_printer(printer_ip, data):
    try:
        printer_port = 9100
        if ':' in printer_ip:
            printer_ip, port = printer_ip.split(':')
            printer_port = int(port)

        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(5)
            s.connect((printer_ip, printer_port))
            s.sendall(data.encode('utf-8'))
        logging.info(f"Successfully sent data to printer at {printer_ip}:{printer_port}")
        return True
    except socket.timeout:
        logging.error(f"Connection timeout to printer at {printer_ip}")
        return False
    except Exception as e:
        logging.error(f"Failed to send data to printer at {printer_ip}: {str(e)}")
        return False

def send_to_usb_printer(printer_name, data, template_format):
    try:
        is_ptp7_series = printer_name.startswith('Brother PT-P7')
        
        if is_ptp7_series:
            qlr = BrotherQLRaster('PT-P750W')
            template_config = TEMPLATES['2']
        else:
            qlr = BrotherQLRaster('QL-800')
            template_config = TEMPLATES.get(template_format, TEMPLATES['6'])
        
        qlr.exception_on_warning = True
        
        if hasattr(qlr, 'set_label_size'):
            qlr.set_label_size(template_config['size'])
        else:
            qlr.label_size = template_config['size']
        
        qlr.add_text(data)
        
        backend = backend_factory('pyusb')
        
        if is_ptp7_series:
            identifier = 'usb://0x04f9:0x2060'
        else:
            identifier = f"usb://{printer_name}" if printer_name else 'usb://0x04f9:0x209b'
        
        send(qlr.data, backend, printer_identifier=identifier, blocking=True)
        logging.info(f"Successfully sent data to USB printer {printer_name}")
        return True
        
    except Exception as e:
        logging.error(f"Failed to send data to USB printer {printer_name}: {str(e)}")
        return False

@app.route('/', methods=['GET'])
def health_check():
    return jsonify({"status": "running", "timestamp": datetime.now().isoformat()})

@app.route('/label', methods=['GET'])
def label():
    try:
        printer_ip = request.args.get('IP')
        serial_number = request.args.get('SN')
        model = request.args.get('MODEL', '')
        model_apn = request.args.get('Model_APN', '')
        type_name = request.args.get('TYPE', '')
        mode = request.args.get('MODE', 'IP')
        
        # Set format based on mode and printer type
        if mode == 'USB' and printer_ip.startswith('Brother PT-P7'):
            format_type = '2'
        elif mode == 'IP':
            format_type = '6'
        else:
            format_type = request.args.get('FORMAT', '6')
        
        username = request.args.get('USER', 'Unknown')
        date = request.now().strftime('%Y-%m-%d')
        
        if not serial_number:
            return jsonify({"status": "error", "message": "No serial number provided"}), 400
            
        data = create_label_data(format_type, serial_number, model, model_apn, type_name, username, date)
        
        if mode == 'IP':
            success = send_to_network_printer(printer_ip, data)
        else:
            success = send_to_usb_printer(printer_ip, data, format_type)
            
        if success:
            return jsonify({"status": "success"})
        return jsonify({"status": "error", "message": f"Failed to send print job to {mode} printer"}), 500
            
    except Exception as e:
        logging.error(f"Error processing print job: {str(e)}")
        return jsonify({"status": "error", "message": f"Server error: {str(e)}"}), 500

@app.route('/getLabelSize', methods=['GET'])
def get_label_size():
    try:
        printer_ip = request.args.get('printerIP')
        mode = request.args.get('MODE', 'IP')
        format_type = request.args.get('FORMAT', '6')
        
        if mode == 'USB' and printer_ip.startswith('Brother PT-P7'):
            template = TEMPLATES['2']
        else:
            template = TEMPLATES.get(format_type, TEMPLATES['6'])
        
        return jsonify({"labelSize": template['size']})
    except Exception as e:
        logging.error(f"Error getting label size: {str(e)}")
        return jsonify({"status": "error", "message": f"Failed to get label size: {str(e)}"}), 500

def create_square_icon():
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
        "Printer Server v2.0\nStatus: Running\nNetwork (IP): 62mm\nUSB: 24mm (PT-P7* series)", 
        "Printer Server Info", 0)

def show_instructions():
    ctypes.windll.user32.MessageBoxW(0,
        "1. Choose IP or USB mode\n2. Enter printer address/name\n3. Save settings\n4. Print",
        "Instructions", 0)

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
