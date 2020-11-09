import socket
import time
from typing import Optional
from ctypes import windll, create_unicode_buffer

import servicemanager
import win32event
import win32service
import win32serviceutil

import getpass
import platform

import logging
import graypy

SEARCHING_NAME = 'консультант'


def getForegroundWindowTitle(hwnd: int) -> Optional[str]:
    length = windll.user32.GetWindowTextLengthW(hwnd)
    buf = create_unicode_buffer(length + 1)
    windll.user32.GetWindowTextW(hwnd, buf, length + 1)
    window_name = buf.value.lower()
    return window_name


def get_foreground_window_HWND() -> int:
    hwnd = windll.user32.GetForegroundWindow()
    return hwnd


def is_HWND_visible(hwnd: int) -> bool:
    return windll.user32.IsWindowVisible(hwnd)


def counter(hwnd: int) -> int:
    active_time = 0
    while is_HWND_visible(hwnd):
        active_time += 1
        time.sleep(1)
    return active_time


def graylog_it(active_time: int) -> None:
    conslogger = logging.getLogger("ConsWatchDogLogger")
    conslogger.setLevel(logging.INFO)
    handler = graypy.GELFUDPHandler("172.16.0.39")
    conslogger.addHandler(handler)
    conslogger.addFilter(ConsFilter(active_time))
    conslogger.debug("Consultant+ window is open")


class ConsFilter(logging.Filter):
    def __init__(self, active_time: int):
        super().__init__()
        self.username = getpass.getuser()
        self.pc_name = platform.uname().node
        self.active_time = active_time

    def filter(self, record):
        record.username = self.username
        record.pc_name = self.pc_name
        record.active_time = self.active_time
        return True


class ConsWatchDogService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'ConsWatchDog'
    _svc_display_name_ = 'ConsWatchDog'
    _svc_description_ = 'Служба мониторинга активности Консультант+'

    @classmethod
    def parse_command_line(cls):
        win32serviceutil.HandleCommandLine(cls)

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        while True:
            hwnd = get_foreground_window_HWND()
            hwnd_window_name = getForegroundWindowTitle(hwnd)
            if SEARCHING_NAME in hwnd_window_name:
                active_time = counter(hwnd)
                graylog_it(active_time)


if __name__ == '__main__':
    ConsWatchDogService.parse_command_line()
