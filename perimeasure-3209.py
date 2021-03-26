#!/usr/bin/python

from pynput.keyboard import Listener as KeyboardListener
from win32gui import GetWindowText, GetForegroundWindow
from pynput.mouse import Listener as MouseListener
from win10toast import ToastNotifier
from datetime import datetime
import threading
import sqlite3
import signal
import psutil
import time
import sys
import gc
import os

lock = threading.Lock()                                     # Used to lock main-thread to prevent race condition
key_buffer      = []                                        # Keyboard buffers: used for storing and counting 'X' tokens
corr_buffer     = []
t_withinkey     = []                                        # Buffer for within and between key intervals
t_betweenkey    = []
t_buffer        = []
click_buffer    = []                                        # Mouse buffers
scroll_buffer   = []
meet_time       = []                                        # Stores seconds in which a participant is using their conference app

corr            =   'Key.backspace', 'Key.delete', 'x1a'    # Corrector keys to be monitored (CTRL-Z is coded as /x1a')

browsers        = [                                         # Used to ignore browsers in conference app detection
                    'Firefox',
                    'Chrome',
                    'Edge',
                    'Explorer'
                  ]


def find_process(process_name):  # Iterates through list of processes to find the specified 'process_name'.
    listOfProcessObjects = []
    for proc in psutil.process_iter():
        try:
            pinfo = proc.as_dict(attrs = ['pid', 'name', 'create_time'])
            # Check if process name contains the given name string.
            if process_name.lower() in pinfo['name'].lower():
                listOfProcessObjects.append(pinfo)
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return listOfProcessObjects


def restart(process_name):  # Before starting the program, this function will terminate older, existing instances.
    listOfProcessIds = find_process(process_name)
    if len(listOfProcessIds) > 0:
        # print('Process Exists | PID and other details are')
        for elem in listOfProcessIds:
            PID = elem['pid']
            PName = elem['name']
            PCreationTime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(elem['create_time']))

            PCreationTime_fmtd = datetime.strptime(PCreationTime, '%Y-%m-%d %H:%M:%S')  # Formatting string to datetime
            tDelta = int((datetime.now() - PCreationTime_fmtd).total_seconds())
            # print((PID, PName, PCreationTime, f"Process started: {tDelta} second(s) ago"))
            if tDelta > 5:  # If existing instance of program existed for over 5 seconds, terminate.
                os.kill(PID, signal.SIGTERM)
                # print(f"Terminated: {PName} | PID: {PID} | Executed: {tDelta} second(s) ago")
    else:
        # print('>>> No such running process found...')
        pass


def get_id():  # Extracts participant ID from software title (e.g., NAME-3209)
    proc_name = os.path.basename(sys.argv[0])
    no_ext = os.path.splitext(proc_name)[0]
    idnr = no_ext.split('-')[-1]
    return idnr


def notify(title, msg):  # Windows 10 'toast' notification
    ToastNotifier().show_toast(title, msg, icon_path='graph.ico', duration=3)


class Timer:  # keyboard metrics and conference-app detection embedded in timer class
    def __init__(self):
        self.t0 = None
        self.window_t0 = None

    def conference_app(self):  # Monitors whether vid-conference app is active and starts counting if true.
        self.window_t0 = time.time()    # Start timer
        while True:
            appID = GetWindowText(GetForegroundWindow())    # Gets the active foreground-window
            # print(appID)
            if not any(i in appID for i in browsers) and ('Microsoft Teams' in appID):
                meet_time.append(1)
                print(">>> Using conference app...")
            time.sleep(1.0 - ((time.time() - self.window_t0) % 1.0))  # Loops def every 1.0s. Locks to system clock.

    def on_press(self, key):
        self.t0 = time.time()           # Start timer
        try:
            t_buffer.append(self.t0)    # Append time to list
            last = t_buffer[-1]         # Get last element in list
            secondlast = t_buffer[-2]   # Get second to last element in list
            diff = secondlast - last    # Calculate difference
            if abs(diff) <= 5:          # If difference is <= 5 seconds, append to list; otherwise disregard
                t_betweenkey.append(round(abs(secondlast-last), 4))     # Get positive interkey interval times
            else:
                # print("Time between keystroke exceeded 5 seconds >>> keystroke disregarded")
                pass
        except IndexError:
            pass
        x = str(key).strip("''")                            # Strip away quotation marks around 'key' variable
        x = x.replace("\\", "")                             # Remove slash from key-code (i.e., /x1a > x1a)
        try:
            key_buffer.append("x")                          # If any key is pressed, place an 'x' inside buffer
            if x in corr:                                   # Detect corrector keys
                corr_buffer.append("x")
        except AttributeError:
            pass

    def on_release(self, key):                              # Function triggered when key is released
        t1 = time.time()                                    # End timer
        t_withinkey.append(round((t1 - self.t0), 4))        # Subtract key-release time from when key was pressed

    def run(self):                                          # Runs conf_app function in separate thread, called in main
        threading.Thread(name='conf-thread', target=self.conference_app).start()


def on_click(x, y, button, pressed):                        # Mouse click; right and left combined
    if pressed:                                             # If R- or L-click is pressed, insert 'x' token in buffer
        click_buffer.append('x')


def on_scroll(x, y, dx, dy):                                # Mouse scroll-wheel registration
    scroll_buffer.append('x')                               # Place 'x' token in buffer for every movement detected


def storage(idnr):                                          # Storing data locally in database
    db = idnr + ".db"                                       # Naming the DB, extracts participant ID number
    try:                                                    # Calculate average within and between keypress intervals
        av_withinkey = round(sum(t_withinkey) / len(t_withinkey), 4)
        av_betweenkey = round(sum(t_betweenkey) / len(t_betweenkey), 4)
    except ZeroDivisionError:                               # Error handling necessary for when program begins
        av_withinkey = 0
        av_betweenkey = 0
        pass

    try:
        con = sqlite3.connect(db)
        c = con.cursor()

        try:
            c.execute("CREATE TABLE db (t, keys, correctors, t_withinkey, t_betweenkey, clicks, scrolls, t_meeting);")
        except:
            pass

        query = "INSERT INTO db (t, " \
                "keys, " \
                "correctors, " \
                "t_withinkey, " \
                "t_betweenkey, " \
                "clicks, " \
                "scrolls," \
                "t_meeting) VALUES (?,?,?,?,?,?,?,?);"
        c.execute(query, (time.strftime('%d/%m/%Y - %H:%M:%S'),
                          len(key_buffer),
                          len(corr_buffer),
                          av_withinkey,
                          av_betweenkey,
                          len(click_buffer),
                          len(scroll_buffer),
                          sum(meet_time)))
        con.commit()
        c.close()
    except sqlite3.Error as error:
        print("Failed to insert data: ", error)
    finally:
        if (con):
            con.close()


def clearing():  # Clearing buffers from memory
    key_buffer.clear()
    corr_buffer.clear()
    t_withinkey.clear()
    t_betweenkey.clear()
    t_buffer.clear()
    click_buffer.clear()
    scroll_buffer.clear()
    meet_time.clear()


def main():
    with lock:                                  # Acquires threading lock preventing race conditions
        threading.Timer(60.0, main).start()     # Main script loops every 60 seconds without time-drift
        storage(get_id())                       # Calls 'storage' function
        clearing()                              # Calls 'clearing' function
        gc.enable()                             # Garbage collection freeing memory from incidental leaks


if __name__ == '__main__':
    restart('perimeasure')
    notify("Perimeasure", "running...")         # Sets title and message in Windows notification
    timer = Timer()
    timer.run()                                 # Runs conference-app detection on separate thread

    keyboard_listener = KeyboardListener(on_press=timer.on_press, on_release=timer.on_release)      # Tracks keyboard events
    mouse_listener = MouseListener(on_click=on_click, on_scroll=on_scroll)                          # Tracks mouse events
    keyboard_listener.start()                                                                       # Starts 'listening' on separate threads
    mouse_listener.start()

    main()
