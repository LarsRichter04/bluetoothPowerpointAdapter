from enum import Enum


class Answers(Enum):
    OPEN_PPTX = b'0'
    NEXT_SLIDE = b'\x01'
    CLOSE_CONNECTION = b'\x03'
    PREVIOUS_SLIDE = b'\x02'
    GOTO_SLIDE = b'\x04'


class Answers_Pepper(Enum):
    OPEN_PPTX = b'0'
    NEXT_SLIDE = b'\\x01'
    CLOSE_CONNECTION = b'\\x03'
    PREVIOUS_SLIDE = b'\\x02'
    GOTO_SLIDE = b'\\x04'


class Commands(Enum):
    CONNECTION_ESTABLISHED = b'0'
    DONE_ACTION = b'1'
    OPENED = b'2'
    ERROR = b'3'


class ERRORS(Enum):
    PRESENTATION_NOT_AVAILABLE = b'0'
    SLIDESHOW_NOT_AVAILABLE = b'1'
