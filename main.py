import time
import logging
import graypy
import getpass
from typing import Optional
from ctypes import windll, create_unicode_buffer


SEARCHING_NAME = 'консультант'


class ConsFilter(logging.Filter):
    active_time = 0

    def __init__(self):
        super().__init__()
        self.username = getpass.getuser()

    def filter(self, record):
        record.username = self.username
        record.tag = "cons.exe"
        record.active_time = ConsFilter.active_time
        return True


conslogger = logging.getLogger("ConsWatchDogLogger")
conslogger.setLevel(logging.DEBUG)
handler = graypy.GELFUDPHandler("172.16.0.39", 12206, debugging_fields=False)
conslogger.addHandler(handler)
conslogger.addFilter(ConsFilter())


def graylog_it(active_time: int) -> None:
    ConsFilter.active_time = active_time
    conslogger.log(logging.DEBUG, f"Окно Консультант+ было открыто {active_time} сек.")


def getForegroundWindowTitle(hwnd: int) -> Optional[str]:
    length = windll.user32.GetWindowTextLengthW(hwnd)
    buf = create_unicode_buffer(length + 1)
    windll.user32.GetWindowTextW(hwnd, buf, length + 1)
    window_name = buf.value.lower()
    return window_name


def get_foreground_window_HWND() -> int:
    hwnd = windll.user32.GetForegroundWindow()
    return hwnd


def counter(hwnd: int) -> int:
    active_time = 0
    while hwnd == get_foreground_window_HWND():
        active_time += 1
        time.sleep(1)
    return active_time


while True:
    hwnd = get_foreground_window_HWND()
    hwnd_window_name = getForegroundWindowTitle(hwnd)
    if SEARCHING_NAME in hwnd_window_name:
        active_time = counter(hwnd)
        graylog_it(active_time)
    time.sleep(5)
