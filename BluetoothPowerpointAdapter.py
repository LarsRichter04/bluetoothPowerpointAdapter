import time
from bluetooth import *
from enums import Commands, Answers_Pepper, Answers
from ppt import Ppt
import configparser


def read_config():
    _config = configparser.ConfigParser()
    _config.read('config.ini')
    pepper_mac = _config.get('Addresses', 'pepper_mac', fallback="")
    bt_uuid = _config.get('Bluetooth', 'uuid', fallback="")

    config_values = {
        'pepper_mac': pepper_mac,
        'bt_uuid': bt_uuid
    }
    return config_values


def handle_message(presentation_index, client_mac_address):
        presentation_index = int(presentation_index)
        ppt = Ppt(presentations[presentation_index])
        answers = Answers_Pepper if client_mac_address == config['pepper_mac'] else Answers
        client = client_socket
        client.send(Commands.OPENED.value + b':' + presentations[presentation_index].encode())
        while True:
            try:
                _data = client.recv(1024)
                print(_data)
                if len(_data) == 0:
                    break
                elif answers.NEXT_SLIDE.value == _data:
                    client.send(ppt.next_slide())
                elif answers.PREVIOUS_SLIDE.value == _data:
                    client.send(ppt.previous_slide())
                elif answers.GOTO_SLIDE.value in _data:
                    client.send(ppt.goto_slide(_data.split(':')[1]))
                elif answers.CLOSE_CONNECTION.value == _data:
                    client.send(ppt.stop_slide_show())
                    client.close()
                    break
            except Exception as e:
                client.send(answers.CLOSE_CONNECTION)


if __name__ == "__main__":
    socket = BluetoothSocket(RFCOMM)
    config = read_config()
    try:
        socket.bind(("", 0))
    except BluetoothError:
        print("Device does not Support Bluetooth or Bluetooth is disabled")
        time.sleep(5)
        quit(0)
    socket.listen(1)
    uuid = config['bt_uuid']
    try:
        advertise_service(socket, name="Bluetooth Powerpoint Adapter",
                          service_id=uuid,
                          service_classes=[uuid, SERIAL_PORT_CLASS],
                          profiles=[SERIAL_PORT_PROFILE])
    except ValueError:
        print("""Invalid uuid provided. an missing or invalid config.ini might be the issue.\n
                 Please ensure it looks like this.\n
                 [Bluetooth]
                 uuid = <your process uuid>""")
        time.sleep(5)
        quit(0)
    print("Successfully advertised service. Now waiting for connection...")
    client_socket, client_address = socket.accept()
    presentations = [x for x in os.listdir(os.getcwd()) if ".pptx" in x]
    presentation_string = ";".join(presentations)
    print("Accepted connection from", client_address[0])
    client_socket.send(Commands.CONNECTION_ESTABLISHED.value + b":" + presentation_string.encode())

    while True:
        data = client_socket.recv(1024)
        if len(data) == 0:
            break
        elif Answers.OPEN_PPTX.value in data:
            print(data)
            decoded = data.split(b':')[1].decode()
            handle_message(decoded, client_address[0])
            break
        elif Answers.CLOSE_CONNECTION.value == data:
            client_socket.close()
            break
