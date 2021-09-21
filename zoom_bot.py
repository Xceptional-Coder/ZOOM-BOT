from datetime import datetime
import subprocess
import pyautogui
import time
import pandas as pd

# path of zoom app
PATH = "C:/Users/satya/AppData/Roaming/Zoom/bin/Zoom.exe"


# Yeh kholega, Zoom app
def sign_in(meetingid, pswd):
    subprocess.call(PATH)

    time.sleep(1)

    # Yeh dbata hai, join button
    join_btn = pyautogui.locateCenterOnScreen('join_button.png')
    pyautogui.moveTo(join_btn)
    pyautogui.click()

    time.sleep(0.1)

    # Yeh likhega pyara  Meeting ID
    pyautogui.click(x=950, y=473)
    time.sleep(0.1)
    pyautogui.write(meetingid)

    # Camera aur Mic ko disable krega
    pyautogui.click(x=802, y=579)  # Mic disable
    pyautogui.click(x=802, y=608)  # Video disable

    # Join button ko dbayega
    pyautogui.click(x=985, y=648)

    time.sleep(1)

    # Password likhega aur enter dbayega
    pyautogui.click(x=961, y=477)
    time.sleep(0.25)
    pyautogui.write(pswd)
    pyautogui.press('enter')


# reads the schedule
df = pd.read_excel('schedule.XLSX')

while True:
    # dekho abhi time match kr rha hai ki nhi
    now = (datetime.now().strftime("%A:%H:%M"))
    df['timings'] = [x[:-3] for x in df['timings']]
    if now in str(df['timings']):  # agar match kr rha hai to meeting join kro

        # match kiya to join kro fir meeting
        m_id = str(df[df["timings"] == now]["meeting ID"].values)  # meeting ID select krega
        m_id = m_id.strip("['']")  # meeting ID me se bracket aur aur quotation nikalega
        m_pswd = str(df[df["timings"] == now]["pswd"].values)  # password select krega
        m_pswd = m_pswd.strip("['']")  # password me se bracket aur quotation nikalega
        print(m_id)  # just for verification
        print(m_pswd)  # ditto

        sign_in(m_id, m_pswd)  # kr do sign in
        time.sleep(30)
        print('signed in')
