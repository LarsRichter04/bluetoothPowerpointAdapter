import time

import win32com.client
from bluetooth import *
import os


class Ppt:
    def _startSlideShow(self):
        self.objCom.SlideShowSettings.Run()

    def __init__(self, path):
        self.objCom = app.Presentations.Open(FileName=path.decode('utf-8'),
                                             WithWindow=0)
        self._startSlideShow()

    def nextSlide(self):
        try:
            self.objCom.SlideShowWindow.View.Next()
            return "done"
        except:
            print("Presentation not available")
            return "Error: Presentation Window not open"

    def previousSlide(self):
        try:
            self.objCom.SlideShowWindow.View.Previous()
            return "done"
        except:
            return "Error: Couldn't go to previousSlide"

    def stopSlideShow(self):
        try:
            self.objCom.SlideShowWindow.View.Exit()
            return "done"
        except:
            return "Couldn't close Presentation"


def handle_message(ppt, presentation_name):
    client_socket.send("opened:".encode() + presentation_name.encode())
    while True:
        _data = client_socket.recv(1024)
        if len(_data) == 0:
            break
        print(_data)
        if _data == b'NextSlide':
            response = ppt.nextSlide()
            client_socket.send(response.encode())
        elif _data == b'CloseConnection':
            client_socket.send(ppt.stopSlideShow().encode())
            socket.close()
            break
        elif _data == b'PreviousSlide':
            client_socket.send(ppt.previousSlide().encode())


app = win32com.client.Dispatch("PowerPoint.Application")
socket = BluetoothSocket(RFCOMM)
try:
    socket.bind(("", 0))
except:
    print("Device does not Support Bluetooth or Bluetooth is disabled")
    time.sleep(5)
    quit(42069)
socket.listen(1)

presentations_path = os.getcwd()
allPresentations = os.listdir(presentations_path)
port = socket.getsockname()[1]
advertise_service(socket, name="Sample Server",
                  service_id="5feedd1f-2df3-404c-a1ec-b7f32a6c9b11",
                  service_classes=["5feedd1f-2df3-404c-a1ec-b7f32a6c9b11", SERIAL_PORT_CLASS],
                  profiles=[SERIAL_PORT_PROFILE])

print("Waiting for connection on port", port)

client_socket, client_address = socket.accept()
print("Accepted connection from", client_address)
presentation_string = ""
for presentation in allPresentations:
    if ".pptx" in presentation:
        presentation_string += presentation + ";"
client_socket.send("Connection established:" + presentation_string)

while True:
    data = client_socket.recv(1024)
    if len(data) == 0:
        break
    elif b'openPPTX:' in data:
        print(data)
        data = data.replace(b'openPPTX:', b'')
        print(data.decode())
        decoded = data.decode()
        handle_message(Ppt((presentations_path + "\\").encode() + data), decoded)
        break
