import time
from enum import Enum

import win32com.client
from bluetooth import *
import os

from win32com.universal import com_error


class Answers(Enum):
    OPENPPTX = b'0'
    NEXTSLIDE = b'1'
    CLOSECONNECTION = b'3'
    PREVIOUSLIDE = b'2'


class Commands(Enum):
    CONNECTIONESTABLISHED = b'0'
    DONEACTION = b'1'


class Ppt:
    def __start_slide_show(self):
        self.objCom.SlideShowSettings.Run()

    def __init__(self, path):
        self.objCom = app.Presentations.Open(FileName=presentations_path + "\\" + path,
                                             WithWindow=0)
        self.__start_slide_show()

    def next_slide(self):
        try:
            self.objCom.SlideShowWindow.View.Next()
            return "done"
        except com_error:
            print("Presentation not available")
            return "Error: Presentation Window not open"

    def previous_slide(self):
        try:
            self.objCom.SlideShowWindow.View.Previous()
            return "done"
        except com_error:
            return "Error: Couldn't go to previousSlide"

    def stop_slide_show(self):
        try:
            self.objCom.SlideShowWindow.View.Exit()
            return "done"
        except com_error:
            return "Couldn't close Presentation"


def handle_message(presentation_name):
    ppt = Ppt(presentation_name)
    client = client_socket
    client.send("opened:".encode() + presentation_name.encode())
    while True:
        _data = client.recv(1024)
        if len(_data) == 0:
            break
        match _data:
            case Answers.NEXTSLIDE.value:
                response = ppt.next_slide()
                client.send(response.encode())
            case Answers.PREVIOUSLIDE.value:
                client.send(ppt.previous_slide().encode())
            case Answers.CLOSECONNECTION.value:
                client.send(ppt.stop_slide_show().encode())
                client.close()
                break


app = win32com.client.Dispatch("PowerPoint.Application")
socket = BluetoothSocket(RFCOMM)
try:
    socket.bind(("", 0))
except BluetoothError:
    print("Device does not Support Bluetooth or Bluetooth is disabled")
    time.sleep(5)
    quit(42069)
socket.listen(1)

presentations_path = os.getcwd()
allPresentations = os.listdir(presentations_path)
port = socket.getsockname()[1]
advertise_service(socket, name="Bluetooth Powerpoint Adapter",
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
    elif Answers.OPENPPTX.value in data:
        print(data)
        dataArr = data.split(b':')
        decoded = dataArr[1].decode()
        handle_message(decoded)
        break
    elif Answers.CLOSECONNECTION.value == data:
        client_socket.close()
        break
