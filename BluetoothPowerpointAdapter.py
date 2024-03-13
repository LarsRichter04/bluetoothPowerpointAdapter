import time
from enum import Enum

import win32com.client
from bluetooth import *
import os

from win32com.universal import com_error


class Answers(Enum):
    OPENPPTX = b'0'
    NEXTSLIDE = b'\x01'
    CLOSECONNECTION = b'\x03'
    PREVIOUSLIDE = b'\x02'


class Commands(Enum):
    CONNECTIONESTABLISHED = b'0'
    DONEACTION = b'1'
    OPENED = b'2'
    ERROR = b'3'



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
            return Commands.DONEACTION.value
        except com_error:
            print("Presentation not available")
            return Commands.ERROR.value

    def previous_slide(self):
        try:
            self.objCom.SlideShowWindow.View.Previous()
            return Commands.DONEACTION.value
        except com_error:
            return Commands.ERROR.value

    def stop_slide_show(self):
        try:
            self.objCom.SlideShowWindow.View.Exit()
            return Commands.DONEACTION.value
        except com_error:
            return Commands.ERROR.value


def handle_message(presentation_index):
    presentation_index = int(presentation_index)
    ppt = Ppt(presentations[presentation_index])
    client = client_socket
    client.send(Commands.OPENED.value + b':' + presentations[presentation_index].encode())
    while True:
        _data = client.recv(1024)
        print(_data)
        if len(_data) == 0:
            break
        match _data:
            case Answers.NEXTSLIDE.value:
                response = ppt.next_slide()
                client.send(response)
            case Answers.PREVIOUSLIDE.value:
                client.send(ppt.previous_slide())
            case Answers.CLOSECONNECTION.value:
                client.send(ppt.stop_slide_show())
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
port = socket.getsockname()[1]
uuid = "5feedd1f-2df3-404c-a1ec-b7f32a6c9b11"
advertise_service(socket, name="Bluetooth Powerpoint Adapter",
                  service_id=uuid,
                  service_classes=[uuid, SERIAL_PORT_CLASS],
                  profiles=[SERIAL_PORT_PROFILE])
print("Waiting for connection on port", port)
client_socket, client_address = socket.accept()
presentations_path = os.getcwd()
allPresentations = os.listdir(presentations_path)
presentations = [x for x in allPresentations if ".pptx" in x]
presentation_string = ";".join(presentations)

print(presentation_string)
print("Accepted connection from", client_address)
client_socket.send(Commands.CONNECTIONESTABLISHED.value + b":" + presentation_string.encode())

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
