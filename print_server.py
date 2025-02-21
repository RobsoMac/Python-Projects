from flask import Flask, request, jsonify
from flask_cors import CORS
import socket
from brother_ql.raster import BrotherQLRaster
from brother_ql.backends.helpers import send
from brother_ql.backends import backend_factory
import logging
from datetime import datetime

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('printer_server.log'),
        logging.StreamHandler()
    ]
)

app = Flask(__name__)
CORS(app)

def send_to_network_printer(printer_ip, data):
    try:
        printer_port = 9100
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(5)  # 5 second timeout
            s.connect((printer_ip, printer_port))
            s.sendall(data.encode('utf-8'))
        logging.info(f"Successfully sent data to printer at {printer_ip}")
        return True
    except socket.timeout:
        logging.error(f"Connection timeout to printer at {printer_ip}")
        return False
    except Exception as e:
        logging.error(f"Failed to send data to printer at {printer_ip}: {str(e)}")
        return False

def send_to_usb_printer(printer_model, data):
    try:
        qlr = BrotherQLRaster(printer_model)
        qlr.exception_on_warning = True
        qlr.set_label_size('62')
        qlr.add_text(data)
        backend = backend_factory('pyusb')
        send(qlr.data, backend, printer_identifier='usb://0x04f9:0x209b', blocking=True)
        logging.info(f"Successfully sent data to USB printer {printer_model}")
        return True
    except Exception as e:
        logging.error(f"Failed to send data to USB printer {printer_model}: {str(e)}")
        return False

@app.route('/', methods=['GET'])
def health_check():
    return jsonify({"status": "running", "timestamp": datetime.now().isoformat()})

@app.route('/label', methods=['POST', 'OPTIONS'])
def label():
    if request.method == 'OPTIONS':
        return '', 200
    
    try:
        data = request.get_json()
        logging.info(f"Received print job: {data}")
        
        printer_ip = data.get('printerIP')
        printer_model = data.get('printerModel')
        label_layout = data.get('labelLayout')
        
        if not label_layout:
            return jsonify({
                "status": "error", 
                "message": "No label layout provided"
            }), 400
            
        if printer_ip:
            success = send_to_network_printer(printer_ip, label_layout)
            if success:
                return jsonify({"status": "success"})
            return jsonify({
                "status": "error", 
                "message": "Failed to send print job to network printer"
            }), 500
            
        elif printer_model:
            success = send_to_usb_printer(printer_model, label_layout)
            if success:
                return jsonify({"status": "success"})
            return jsonify({
                "status": "error", 
                "message": "Failed to send print job to USB printer"
            }), 500
            
        return jsonify({
            "status": "error", 
            "message": "No printer specified"
        }), 400
        
    except Exception as e:
        logging.error(f"Error processing print job: {str(e)}")
        return jsonify({
            "status": "error",
            "message": f"Server error: {str(e)}"
        }), 500

@app.route('/getLabelSize', methods=['GET'])
def get_label_size():
    try:
        printer_ip = request.args.get('printerIP')
        # You could add logic here to actually query the printer
        return jsonify({"labelSize": "62mm"})
    except Exception as e:
        logging.error(f"Error getting label size: {str(e)}")
        return jsonify({
            "status": "error",
            "message": f"Failed to get label size: {str(e)}"
        }), 500

if __name__ == '__main__':
    try:
        logging.info("Starting printer server on port 3000...")
        app.run(host='0.0.0.0', port=3000)
    except Exception as e:
        logging.error(f"Failed to start server: {str(e)}")
