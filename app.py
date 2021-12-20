import time, win32con, win32api, win32gui
from io import BytesIO
import win32clipboard
from PIL import Image
from datetime import datetime
from pynput.keyboard import Key, Controller
import schedule
from urllib.request import Request, urlopen as uReq
from bs4 import BeautifulSoup as soup
import urllib


# # 카톡창 이름, (활성화 상태의 열려있는 창)
kakao_opentalk_name = '👉청년부 잡담방👈'


# # 채팅방에 메시지 전송
def kakao_sendtext(chatroom_name, text):
    # # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow( None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RICHEDIT50W", None)
    # hwndListControl = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    pressPaste()
    time.sleep(0.01)
    returnPaste()

    print("sent")


# # 붙여넣기
def pressPaste():
    keyboard = Controller()

    keyboard.press(Key.ctrl)
    keyboard.press('v')
    time.sleep(0.01)
    keyboard.release(Key.ctrl)
    keyboard.release('v')

# # 엔터
def returnPaste():
    keyboard = Controller()

    keyboard.press(Key.enter)
    time.sleep(0.01)
    keyboard.release(Key.enter)

# # 엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

## 이미지 다운로드
def downloadImg():
    my_url = Request('https://www.bible.com/ko/verse-of-the-day/', headers={'User-Agent': 'Mozilla/5.0'})
    uClient = uReq(my_url)
    page_html = uClient.read()
    uClient.close()
    currentDate = datetime.now().date()
    currentDateTime = datetime.now()
    print("downLoad "+str(currentDateTime))
    page_soup = soup(page_html, "html.parser")
    bible_img = page_soup.findAll("amp-img",{"sizes":"(max-width: 385px) 320px, 640px"})
    bible_url = bible_img[-1]["src"].replace("320x320", "1280x1280")
    urllib.request.urlretrieve("https:"+bible_url, "C:/Users/david/Downloads/todaysbible"+str(currentDate)+".jpg")
    return True


# # 채팅방 열기

def open_chatroom(chatroom_name):
    # # 채팅방 목록 검색하는 Edit (채팅방이 열려있지 않아도 전송 가능하기 위하여)
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx( hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx( hwndkakao_edit1, None, "EVA_Window", None)
    hwndkakao_edit2_2 = win32gui.FindWindowEx( hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
    hwndkakao_edit3 = win32gui.FindWindowEx( hwndkakao_edit2_2, None, "Edit", None)

    # # Edit에 검색 _ 입력되어있는 텍스트가 있어도 덮어쓰기됨
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)   # 안정성 위해 필요
    SendReturn(hwndkakao_edit3)
    time.sleep(1)

def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()

def fileCopy():
    currentDate = datetime.now().date()
    currentDateTime = datetime.now()
    print(currentDateTime)

    filepath = 'C:/Users/david/Downloads/todaysbible' + str(currentDate) + '.jpg'
    image = Image.open(filepath)

    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    send_to_clipboard(win32clipboard.CF_DIB, data)

def main():
    downloadImg() # 이미지 다운
    fileCopy() # 이미지 클리보드 복사
    open_chatroom(kakao_opentalk_name)  # 채팅방 열기

    text = ""
    kakao_sendtext(kakao_opentalk_name, text)    # 메시지 전송

if __name__ == '__main__':
    schedule.every().day.at("07:00").do(main)

    while True:
        schedule.run_pending()
 
        time.sleep(1)
