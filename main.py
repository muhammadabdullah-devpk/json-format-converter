from datetime import datetime, timezone
from http.server import HTTPServer, BaseHTTPRequestHandler
import json
import os

def convertFromFormat1(jsonObject):
    locationParts = jsonObject["location"].split("/")
    
    result = {
        "deviceID": jsonObject["deviceID"],
        "deviceType": jsonObject["deviceType"],
        "timestamp": jsonObject["timestamp"],
        "location": {
            "country": locationParts[0],
            "city": locationParts[1],
            "area": locationParts[2],
            "factory": locationParts[3],
            "section": locationParts[4]
        },
        "data": {
            "status": jsonObject["operationStatus"],
            "temperature": jsonObject["temp"]
        }
    }
    
    return result

def convertFromFormat2(jsonObject):
    dt = datetime.strptime(jsonObject['timestamp'], "%Y-%m-%dT%H:%M:%S.%fZ")
    dt = dt.replace(tzinfo=timezone.utc)
    
    timestamp = int(dt.timestamp() * 1000)
    
    result = {
        "deviceID": jsonObject["device"]["id"],
        "deviceType": jsonObject["device"]["type"],
        "timestamp": timestamp,
        "location": {
            "country": jsonObject["country"],
            "city": jsonObject["city"],
            "area": jsonObject["area"],
            "factory": jsonObject["factory"],
            "section": jsonObject["section"]
        },
        "data": jsonObject["data"]
    }
    
    return result

def run_tests():
    # Test Format 1
    input_format_1 = {
        "deviceID": "device-123",
        "deviceType": "sensor",
        "timestamp": 1624445837783,
        "location": "japan/tokyo/keiyō-industrial-zone/factory-1/section-2",
        "operationStatus": "active",
        "temp": 22
    }
    
    expected_output_1 = {
        "deviceID": "device-123",
        "deviceType": "sensor",
        "timestamp": 1624445837783,
        "location": {
            "country": "japan",
            "city": "tokyo",
            "area": "keiyō-industrial-zone",
            "factory": "factory-1",
            "section": "section-2"
        },
        "data": {
            "status": "active",
            "temperature": 22
        }
    }
    
    # Test Format 2
    input_format_2 = {
        "device": {
            "id": "device-456",
            "type": "camera"
        },
        "timestamp": "2021-06-23T10:57:17.783Z",
        "country": "japan",
        "city": "tokyo",
        "area": "keiyō-industrial-zone",
        "factory": "factory-1",
        "section": "section-2",
        "data": {
            "status": "idle",
            "temperature": 25
        }
    }
    
    expected_output_2 = {
        "deviceID": "device-456",
        "deviceType": "camera",
        "timestamp": 1624445837783,
        "location": {
            "country": "japan",
            "city": "tokyo",
            "area": "keiyō-industrial-zone",
            "factory": "factory-1",
            "section": "section-2"
        },
        "data": {
            "status": "idle",
            "temperature": 25
        }
    }
    
    assert convertFromFormat1(input_format_1) == expected_output_1, "Format 1 conversion failed!"
    assert convertFromFormat2(input_format_2) == expected_output_2, "Format 2 conversion failed!"
    
    print("All tests passed")

class SimpleHTTPRequestHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        
        # Run tests to verify they pass and return the status
        try:
            run_tests()
            response = {"status": "success", "message": "All tests passed"}
        except AssertionError as e:
            response = {"status": "failed", "message": str(e)}
        except Exception as e:
            response = {"status": "error", "message": str(e)}
            
        self.wfile.write(json.dumps(response).encode('utf-8'))

if __name__ == "__main__":
    # Run tests on local startup
    run_tests()
    
    # Start web server for Replit Deployment
    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(('0.0.0.0', port), SimpleHTTPRequestHandler)
    print(f"Server starting on port {port}...")
    server.serve_forever()
