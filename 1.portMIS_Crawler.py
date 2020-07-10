# 와! 킹자갓부프레임워크! 마우스 키보드 매크로로 해결하자
# 원하는 날짜 세팅해놓고 실행하면 모든 지역별로 다운로드

import pyautogui
import keyboard
import time
import os
from os import rename

DOWNDIR = 'C:\\Users\\nakag\\Downloads\\'
harbor_code = ['020', '030', '031', '050', '200', '201', '203', '204', '206', '300',
               '301', '302', '500', '501', '610', '611', '620', '622', '700', '810',
               '811', '812', '813', '814', '815', '816', '817', '820', '900', '901']
print("크롤링할 항만 갯수:",len(harbor_code))

pos1 = (555, 342) #항만 선택창 열기
pos2 = (464, 409) #항만 검색창
pos3 = (423, 443) #검색 결과 선택
pos4 = (885, 717) #선택
pos5 = (1144, 346) #조회
pos6 = (1353, 339) #액셀 저장
pos7 = (693, 407) #항만 초기화

for code in harbor_code:
    time.sleep(2)
    pyautogui.click(pos1)
    pyautogui.click(pos2)
    keyboard.write(code, delay=0.05)
    keyboard.press_and_release('enter')
    time.sleep(1)
    pyautogui.click(pos3)
    pyautogui.click(pos4)
    pyautogui.click(pos5)
    time.sleep(10)

    pyautogui.click(pos6)#Download
    time.sleep(10)

    # initialize
    pyautogui.click(pos1)
    pyautogui.click(pos7)
    pyautogui.click(pos2)
    keyboard.press_and_release('ctrl+a')
    keyboard.press_and_release('backspace')
    pyautogui.click(pos1)

    directory = os.listdir(DOWNDIR) #다운로드 폴더 내용 리스트로 가져오기
    for f in directory:
        if f[:4] == '컨테이너':
            rename(DOWNDIR+f, DOWNDIR+code+'.xlsx')
    print(code,'Complete')