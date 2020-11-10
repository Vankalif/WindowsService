import socket
import time
from datetime import datetime

import servicemanager
import win32event
import win32service
import win32serviceutil
import getpass
import logging
import graypy

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
handler = graypy.GELFUDPHandler("172.16.0.39", 12206)
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


def write_to_file():
    with open("C:\\Users\\turshiev_nm\\Desktop\\WindowsService\\service.log", "a+") as file:
        file.write(str(datetime.now()) + "\n")


class ServiceBaseClass(win32serviceutil.ServiceFramework):

    _svc_name_ = 'pythonService'
    _svc_display_name_ = 'Python Service'
    _svc_description_ = 'Python Service Description'

    @classmethod
    def parse_command_line(cls):
        win32serviceutil.HandleCommandLine(cls)

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        self.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        self.start()
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def start(self):
        pass

    def stop(self):
        pass

    def main(self):
        pass


class ExampleService(ServiceBaseClass):
    _svc_name_ = "ConsWatchDogService"
    _svc_display_name_ = "Сервис мониторинга активности"
    _svc_description_ = "Сервис мониторинга активности Консультант+"

    def start(self):
        self.isrunning = True

    def stop(self):
        self.isrunning = False

    def main(self):
        while self.isrunning:
            while self.isrunning:
                write_to_file()
                time.sleep(5)


if __name__ == '__main__':
    ExampleService.parse_command_line()
