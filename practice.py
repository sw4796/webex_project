import pyautogui
import pygetwindow as gw
import pywinauto
import imaplib
import email
import time

#
# # get_key_from_mail과 세트입니다. 메일을 읽들 때 사용됩니다.
# def findEncodingInfo(txt):
#     info = email.header.decode_header(txt)
#     s, encoding = info[0]
#     return s, encoding
#
#
# # USER와 PASSWORD를 여러분의 정보로 바꿔주세요
# def get_key_from_mail():
#     global api_key
#     sendto = 'sw2518sw@gmail.com'
#     user = 'sw2518sw@gmail.com'
#     password = "bkrjauxcwowlodxe"
#
#     # 메일서버 로그인
#     imapsrv = "smtp.gmail.com"
#     # 아래행 imap = imaplib.IMAP4_SSL('imap.gmail.com')에서 수정
#     imap = imaplib.IMAP4_SSL(imapsrv, "993")
#     id = user
#     pw = password
#     imap.login(id, pw)
#
#     # 받은 편지함
#     imap.select('inbox')
#
#     # 받은 편지함 모든 메일 검색
#     resp, data = imap.uid('search', None, '(FROM "swlee4796@naver.com")')
#
#     # 여러 메일 읽기 (반복)
#     all_email = data[0].split()
#     all_email.reverse()
#     # del(all_email[0])
#     for mail in all_email:
#
#         # fetch 명령을 통해서 메일 가져오기 (RFC822 Protocol)
#         result, data = imap.uid('fetch', mail, '(RFC822)')
#
#         # 사람이 읽기 힘든 Raw 메세지 (byte)
#         raw_email = data[0][1]
#
#         # 메시지 처리(email 모듈 활용)
#         email_message = email.message_from_bytes(raw_email)
#
#         # 이메일 정보 keys
#         # print(email_message.keys())
#         print('FROM:', email_message['From'])
#         print('SENDER:', email_message['Sender'])
#         print('TO:', email_message['To'])
#         print('DATE:', email_message['Date'])
#
#         email_time = email_message['Date'].replace(':', ' ').split()
#         now_time = time.strftime('%a, %d %b %Y %H %M %S', time.localtime(time.time())).split()
#         print(email_time)
#         print(now_time)
#
#         for i in range(0, 5):
#             if (email_time[i] != now_time[i]) | int(now_time[5])-int(email_time[5]) > 1:
#                 return 0
#
#         b, encode = findEncodingInfo(email_message['Subject'])
#         # 제목
#         print('SUBJECT:', str(b, encode))
#         subject = str(b, encode)
#         if subject != '웹엑스명령':
#             # print(subject)
#             return 0
#         text = ''
#         # 이메일 본문 내용 확인
#         print('[CONTENT]')
#         print('=' * 80)
#         if email_message.is_multipart():
#             for part in email_message.get_payload():
#                 bytes = part.get_payload(decode=True)
#                 encode = part.get_content_charset()
#                 print(str(bytes, encode))
#                 text = str(bytes, encode)
#                 if "꺼버려" in text:
#                     return True
#                 break
#         print('=' * 80)
#         break
#
#     imap.close()
#     imap.logout()


pyautogui.FAILSAFE = False

# 웹엑스 화면 활성화
win = gw.getWindowsWithTitle('webex')[0]

if win.isActive == False:
    pywinauto.application.Application().connect(handle=win._hWnd).top_window().set_focus()
    win.activate()

if win.isMaximized == False:
    win.maximize()

#화면에서 카메라 아이콘 탐색
while(True):
    #화면 활성화, 최대화
    if win.isActive == False:
        pywinauto.application.Application().connect(handle=win._hWnd).top_window().set_focus()
        win.activate()

    if win.isMaximized == False:
        win.maximize()

    test1 = list(pyautogui.locateAllOnScreen("camera_on.png", confidence=0.9))
    print(len(test1))

    # 메인 동작 시작(카메라 아이콘이 3개 이상일때 실행)
    if len(test1) > 2:
        #webex 종료 버튼 누르기
        button_exit = pyautogui.locateCenterOnScreen("exit.png")
        pyautogui.click(button_exit)
        break
    time.sleep(5.0)


# 키값: bkrjauxcwowlodxe
