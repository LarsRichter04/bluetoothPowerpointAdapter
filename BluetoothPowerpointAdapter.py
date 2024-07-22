import threading

from PyQt5.QtWidgets import QApplication
from bluetooth import *

from adapter import background_tasks
from configreader import read_config
from gui import Window

if __name__ == "__main__":

    config = read_config()
    app = QApplication(sys.argv)
    window = Window(config["version"])
    window.show()
    print('Bluetooth Powerpoint Adapter '+config['version'])
    btThread = threading.Thread(target=background_tasks,daemon=True)
    btThread.start()
    app.exec_()
    sys.exit()


