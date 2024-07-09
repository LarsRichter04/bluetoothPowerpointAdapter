import os

from win32com.universal import com_error

from enums import Commands, ERRORS

import win32com.client


class Ppt:
    def __start_slide_show(self):
        self.objCom.SlideShowSettings.Run()

    def __init__(self, path):
        self.app = win32com.client.Dispatch("PowerPoint.Application")
        self.objCom = self.app.Presentations.Open(FileName=os.getcwd() + "\\" + path,
                                                  WithWindow=0)
        self.__start_slide_show()

    def next_slide(self):
        try:
            self.objCom.SlideShowWindow.View.Next()
            return Commands.DONE_ACTION.value
        except com_error:
            print("Presentation not available")
            return Commands.ERROR.value + b':' + ERRORS.SLIDESHOW_NOT_AVAILABLE.value

    def previous_slide(self):
        try:
            self.objCom.SlideShowWindow.View.Previous()
            return Commands.DONE_ACTION.value
        except com_error:
            return Commands.ERROR.value + b':' + ERRORS.SLIDESHOW_NOT_AVAILABLE.value

    def stop_slide_show(self):
        try:
            self.objCom.SlideShowWindow.View.Exit()
            return Commands.DONE_ACTION.value
        except com_error:
            return Commands.ERROR.value + b":" + ERRORS.PRESENTATION_NOT_AVAILABLE.value

    def goto_slide(self, index):
        try:
            self.objCom.SlideShowWindow.View.GotoSlide(index)
            return Commands.DONE_ACTION.value
        except com_error:
            return Commands.ERROR.value + b':' + ERRORS.SLIDESHOW_NOT_AVAILABLE.value
