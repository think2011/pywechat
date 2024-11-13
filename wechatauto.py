'''模块:\n
Messages: 5种类型的发送消息功能包括:单人单条,单人多条,多人单条,多人多条,转发消息:多人同一条\n
Files: 5种类型的发送文件功能包括:单人单个,单人多个,多人单个,多人多个,转发文件:多人同一个\n
FriendSettings: 涵盖了PC微信针对某个好友的全部操作\n
GroupSettings: 涵盖了PC微信针对某个群聊的全部操作\n
Contacts: 获取3种类型通讯录好友的备注与昵称包括:微信好友,企业号微信,群聊\n
Call: 给某个好友打视频或语言电话\n
Auto_response:自动接应微信视频或语言电话\n
WechatSettings: 修改PC微信设置\n
\n
函数:函数为上述模块内的所有方法\n
使用该模块内的方法时,你可以:\n
from pywechat.wechatauto import Messages\n
Messages.send_messages_to_friend()\n
或者\n
from pywechat.wechatauto import send_messages_to_friend\n
send_messages_to_friend()\n
或者\n
from pywechat import wechatauto as wechat\n
wechat.send_messages_to_friend()\n
来进行使用\n 
'''
#########################################依赖环境#####################################
import time
import json
import pyautogui
from pywechat.wechatTools import Tools
from pywechat.wechatTools import mouse
from pywechat.wechatTools import Desktop
from pywechat.winSettings import Systemsettings
from pywechat.Errors import TimeNotCorrectError
from pywechat.Errors import HaveBeenPinnedError
from pywechat.Errors import HaveBeenUnpinnedError
from pywechat.Errors import HaveBeenMutedError
from pywechat.Errors import HaveBeenStickiedError
from pywechat.Errors import HaveBeenUnmutedError
from pywechat.Errors import HaveBeenUnstickiedError
from pywechat.Errors import HaveBeenStaredError
from pywechat.Errors import HaveBeenUnstaredError
from pywechat.Errors import HaveBeenInBlacklistError
from pywechat.Errors import HaveBeenOutofBlacklistError
from pywechat.Errors import NoWechat_number_or_Phone_numberError
from pywechat.Errors import HaveBeenSetChatonlyError
from pywechat.Errors import HaveBeenSetUnseentohimError
from pywechat.Errors import HaveBeenSetDontseehimError
from pywechat.Errors import PrivacytNotCorrectError
from pywechat.Errors import EmptyFileError
from pywechat.Errors import EmptyFolderError
from pywechat.Errors import NotFileError
from pywechat.Errors import NotFolderError
from pywechat.Errors import CantCreateGroupError
from pywechat.Errors import NoPermissionError
from pywechat.Errors import SameNameError
from pywechat.Errors import AlreadyOpenError
from pywechat.Errors import AlreadyCloseError
from pywechat.Errors import AlreadyInContactsError
from pywechat.Errors import EmptyNoteError
##############################################@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
pyautogui.FAILSAFE=False#防止鼠标在屏幕边缘处造成的误触
class Messages():
    @staticmethod
    def send_message_to_friend(friend:str,message:str,wechat_path:str=None,delay:int=2,is_maximize:bool=True):
        '''
        该方法用于给单个好友或群聊发送单条信息\n
        friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
        message:待发送消息。格式:message="消息"\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        '''
        chat,main_window=Tools.find_friend_in_MessageList(friend,wechat_path,is_maximize=False)
        if chat:
            if is_maximize:
                main_window.maximize()
            chat.set_focus()
            chat.click_input()
            chat.type_keys(message,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
        else:
            if is_maximize:
                main_window.maximize()
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            chat.type_keys(message,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()

    @staticmethod
    def send_messages_to_friend(friend:str,messages:list,delay:int=2,wechat_path:str=None,is_maximize:bool=True):
        '''
        该方法用于给单个好友或群聊发送多条信息\n
        friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
        message:待发送消息列表。格式:message=["发给好友的消息1","发给好友的消息2"]\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        '''
        chat,main_window=Tools.find_friend_in_MessageList(friend=friend,wechat_path=wechat_path,is_maximize=False)
        if chat:
            if is_maximize:
                main_window.maximize()
            chat.set_focus()
            chat.click_input()
            for message in messages:
                chat.type_keys(message,with_spaces=True)
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
        else:
            if is_maximize:
                main_window.maximize()
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            for message in messages:
                chat.type_keys(message,with_spaces=True)
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()

    @staticmethod
    def send_messages_to_firends(friends:str,messages:list[list[str]],wechat_path:str=None,delay:int=2,is_maximize:bool=True):
        '''
        该方法用于给多个好友或群聊发送多条信息\n
        friends:好友或群聊备注列表,格式:firends=["好友1","好友2","好友3"]。\n
        messages:待发送消息,格式: message=[[发给好友1的多条消息],[发给好友2的多条消息],[发给好友3的多条信息]]。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        注意!messages与friends长度需一致,并且messages内每一个列表顺序需与friends中好友名称出现顺序一致,否则会出现消息发错的尴尬情况
        '''
        Chats=dict(zip(friends,messages))
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        time.sleep(2)
        for friend in Chats:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend,with_spaces=True)
            time.sleep(2)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            for message in Chats.get(friend):
                chat.type_keys(message,with_spaces=True)
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()

    @staticmethod
    def send_message_to_friends(friends:list,message:list,wechat_path:str=None,delay:int=2,is_maximize:bool=True):
        '''
        该方法用于给多个好友或群聊发送单条信息\n
        friends:好友或群聊备注。格式:friends=["好友1","好友2","好友3"]\n
        message:待发送消息,格式: message=[发给好友1的多条消息,发给好友2的多条消息,发给好友3的多条消息]。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        注意!message与friends长度需一致,并且messages内每一条消息顺序需与friends中好友名称出现顺序一致,否则会出现消息发错的尴尬情况\n
        '''
        Chats=dict(zip(friends,message))
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        time.sleep(2)
        for friend in Chats:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            chat.type_keys(Chats.get(friend),with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()

    @staticmethod
    def forward_message(friends:list,message:str,wechat_path:str=None):
        '''
        该方法用于给多个好友或群聊转发单条信息\n
        friends:好友或群聊备注列表。格式:friends=["好友1","好友2","好友3"]\n
        message:待发送消息,格式: message="转发消息"。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        '''
        messages=[message for _ in range(len(friends))]
        Messages.send_message_to_friends(friends,messages,wechat_path)



class Files():
    @staticmethod
    def send_file_to_friend(friend:str,file_path:str,wechat_path:str=None,delay:int=2,is_maximize:bool=True,with_messages:bool=False,messages:list=[],message_first:bool=False):
        '''
        该方法用于给单个好友或群聊发送单个文件\n
        friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
        file_path:待发送文件绝对路径。
        messages:与文件一同发送的消息。格式:message=["消息1","消息2","消息3"]\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        delay:发送单条信息或文件的延迟,单位:秒/s,默认2s。\n
        with_messages:发送文件时是否给好友发消息。True发送消息,默认为False\n
        messages_first:默认先发送文件后发送消息,messages_first设置为True,先发送消息,后发送文件,\n
        '''
        if Systemsettings.is_empty_file(file_path):
            raise EmptyFileError(f'不能发送空文件！请重新选择文件路径！')
        if not Systemsettings.is_file(file_path):
            raise NotFileError(f'该路径下的内容不是文件,无法发送!')
        chat,main_window=Tools.find_friend_in_MessageList(friend,wechat_path,is_maximize=False)
        if chat:
            chat.set_focus()
            chat.click_input()
            if is_maximize:
                main_window.maximize()
            if with_messages and messages:
                if message_first:
                    for message in messages:
                        chat.type_keys(message)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s')
                    Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
                    pyautogui.hotkey("ctrl","v")
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
                   
                else:
                    Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
                    pyautogui.hotkey("ctrl","v")
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
                    for message in messages:
                        chat.type_keys(message)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s')
            else:
                if is_maximize:
                    main_window.maximize()
                Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')

        else:
            if is_maximize:
                main_window.maximize()
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            if with_messages and messages:
                if message_first:
                    for message in messages:
                        chat.type_keys(message)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s')
                    Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
                    pyautogui.hotkey("ctrl","v")
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')     
                else:
                    Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
                    pyautogui.hotkey("ctrl","v")
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
                    for message in messages:
                        chat.type_keys(message)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s') 
            else:                      
                Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()

        
    @staticmethod
    def send_files_to_friend(friend:str,folder_path:str,wechat_path:str=None,delay:int=2,is_maximize:bool=True,with_messages:bool=False,messages:list=[],messages_first:bool=False):
        '''
        该方法用于给单个好友或群聊发送多个文件\n
        friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
        folder_path:所有待发送文件所处的文件夹的地址。
        messages:与文件一同发送的消息。格式:message=["消息1","消息2","消息3"]\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        delay:发送单条信息或文件的延迟,单位:秒/s,默认2s。\n
        with_messages:发送文件时是否给好友发消息。True发送消息,默认为False\n
        messages_first:默认先发送文件后发送消息,messages_first设置为True,先发送消息,后发送文件,\n
        '''
        if not Systemsettings.is_dirctory(folder_path):
            raise NotFolderError(f'给定路径不是文件夹！若需发送多个文件给好友,请将所有待发送文件置于文件夹内,并在此方法中传入文件夹路径')
        files_in_folder=Systemsettings.get_files_in_folder(folder_path=folder_path)
        if not files_in_folder:
            raise EmptyFolderError(f"文件夹内没有文件！请重新选择！")
        def send_files():
            if len(files_in_folder)<=9:
                Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
            else:
                files_num=len(files_in_folder)
                rem=len(files_in_folder)%9
                for i in range(0,files_num,9):
                    if i+9<files_num:
                        Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder[i:i+9])
                        pyautogui.hotkey("ctrl","v")
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s')
                if rem:
                    Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder[files_num-rem:files_num])
                    pyautogui.hotkey("ctrl","v")
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
        chat,main_window=Tools.find_friend_in_MessageList(friend,wechat_path,is_maximize=False)
        if chat:
            if is_maximize:
                main_window.maximize()
            chat.set_focus()
            chat.click_input()
            if with_messages and messages:
                if messages_first:
                    for message in messages:
                        chat.type_keys(message)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s')
                    send_files()
                else:
                    send_files()
                    for message in messages:
                        chat.type_keys(message)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s') 
            else:
                send_files()

        else:
            if is_maximize:
                main_window.maximize()
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            if with_messages and messages:
                if messages_first:
                    for message in messages:
                        chat.type_keys(message)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s')
                    send_files()    
                else:
                    send_files()
                    for message in messages:
                        chat.type_keys(message)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s') 
            else:                      
                send_files()
        time.sleep(2)
        main_window.close()
    

    @staticmethod
    def send_file_to_friends(friends:list[str],file_paths:list,wechat_path:str=None,delay:int=2,is_maximize:bool=True,with_messages:bool=False,messages:list[list[str]]=[],message_first:bool=False):
        '''
        该方法用于给每个好友或群聊发送单个不同的文件或信息\n
        friends:好友或群聊备注。格式:friends=["好友1","好友2","好友3"]\n
        file_paths:待发送文件,格式: file=[发给好友1的单个文件,发给好友2的文件,发给好友3的文件]。\n
        messages:待发送消息，格式:messages=["发给好友1的单条消息","发给好友2的单条消息","发给好友3的单条消息"]
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        注意!messages,filepaths与friends长度需一致,并且messages内每一条消息顺序需与friends中好友名称出现顺序一致,否则会出现消息发错的尴尬情况
        '''
        for file_path in file_paths:
            if Systemsettings.is_empty_file(file_path):
                raise EmptyFileError(f'不能发送空文件！请重新选择文件路径！')
            if Systemsettings.is_dirctory(file_path):
                raise NotFileError(f'该路径下的内容不是文件,无法发送!')
            if Systemsettings.is_file(file_path):
                raise NotFileError(f'该路径下的内容不是文件,无法发送!')
        Files=dict(zip(friends,file_paths))
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        time.sleep(2)
        if with_messages and messages:
            Chats=dict(zip(friends,messages))
            for friend in Files:
                search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
                search.click_input()
                search.type_keys(friend,with_spaces=True)
                time.sleep(delay)
                pyautogui.hotkey('enter')
                chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
                chat.set_focus()
                chat.click_input()
                if message_first:
                    messages=Chats.get(friend)
                    for message in messages:
                        chat.type_keys(message,with_spaces=True)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s')
                    Systemsettings.copy_file_to_windowsclipboard(Files.get(friend))
                    pyautogui.hotkey('ctrl','v')
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
                else:
                    Systemsettings.copy_file_to_windowsclipboard(Files.get(friend))
                    pyautogui.hotkey('ctrl','v')
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
                    messages=Chats.get(friend)
                    for message in messages:
                        chat.type_keys(message,with_spaces=True)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s')
        else:
            for friend in Files:
                search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
                search.click_input()
                search.type_keys(friend,with_spaces=True)
                time.sleep(delay)
                pyautogui.hotkey('enter')
                chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
                chat.set_focus()
                chat.click_input()
                Systemsettings.copy_file_to_windowsclipboard(Files.get(friend))
                pyautogui.hotkey('ctrl','v')
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()

    @staticmethod
    def send_files_to_friends(friends:list[str],folder_paths:list,wechat_path:str=None,delay:int=2,is_maximize:bool=True,with_messages:bool=False,messages:list[list[str]]=[],message_first:bool=False):
        '''
        该方法用于给多个好友或群聊发送多个不同或相同的文件夹内的所有文件.\n
        friends:好友或群聊备注。格式:friends=["好友1","好友2","好友3"]\n
        folder_paths:待发送文件夹路径列表，每个文件夹内可以存放多个文件,格式: FolderPath_list=["","",""]\n
        message_list:待发送消息，格式:message=[[""],[""],[""]]\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        注意! messages,folder_paths与friends长度需一致,并且messages内每一条消息FolderPath_list每一个文件\n
        顺序需与friends中好友名称出现顺序一致,否则会出现消息发错的尴尬情况\n
        '''
        for folder_path in folder_paths:
            if not Systemsettings.is_dirctory(folder_path):
                raise NotFolderError(f'给定路径不是文件夹！若需发送多个文件给好友,请将所有待发送文件置于文件夹内,并在此方法中传入文件夹路径')
        def send_files(folder_path):
            files_in_folder=Systemsettings.get_files_in_folder(folder_path=folder_path)
            if len(files_in_folder)<=9:
                Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
            else:
                files_num=len(files_in_folder)
                rem=len(files_in_folder)%9
                for i in range(0,files_num,9):
                    if i+9<files_num:
                        Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder[i:i+9])
                        pyautogui.hotkey("ctrl","v")
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s')
                if rem:
                    Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder[files_num-rem:files_num])
                    pyautogui.hotkey("ctrl","v")
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
        Files=dict(zip(friends,folder_paths))
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        if with_messages and messages:
            Chats=dict(zip(friends,messages))
            for friend in Files:
                search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
                search.click_input()
                search.type_keys(friend,with_spaces=True)
                time.sleep(delay)
                pyautogui.hotkey('enter')
                chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
                chat.set_focus()
                chat.click_input()
                if message_first:
                    messages=Chats.get(friend)
                    for message in messages:
                        chat.type_keys(message,with_spaces=True)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s')
                    folder_path=Files.get(friend)
                    send_files(folder_path)
                else:
                    folder_path=Files.get(friend)
                    send_files(folder_path)
                    messages=Chats.get(friend)
                    for message in messages:
                        chat.type_keys(message,with_spaces=True)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s')
        else:
            for friend in Files:
                search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
                search.click_input()
                search.type_keys(friend,with_spaces=True)
                time.sleep(delay)
                pyautogui.hotkey('enter')
                chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
                chat.set_focus()
                chat.click_input()
                folder_path=Files.get(friend)
                send_files(folder_path)
        time.sleep(2)
        main_window.close()

    @staticmethod
    def forward_file(friends:list[str],file_path:str,message:str,wechat_path:str=None,is_maximize:bool=True,with_messages:bool=False,message_first:bool=False):
        '''
        friends:好友或群聊备注列表。格式:friends=["好友1","好友2","好友3"]\n
        file_path:待发送文件,格式: file_path="转发文件路径"。\n
        message:转发消息,格式:messaghe="待转发消息"
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        该方法用来给多个好友或群聊转发文件\n
        '''
        messages=[[message] for _ in len(friends)]
        file_paths=[file_path for _ in range(len(friends))]
        Files.send_file_to_friends(file_paths=file_paths,wechat_path=wechat_path,is_maximize=is_maximize,message_list=messages,with_messages=with_messages,message_first=message_first)

class WechatSettings():
    @staticmethod
    def open_settings(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来打开微信设置界面。\n
        '''    
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(control_type="ToolBar",title='导航')
        setting=Toolbar.children(control_type='Pane',title='')[3].children(control_type='Button',title="设置及其他")[0]
        setting.click_input()
        settings_menu=main_window.child_window(class_name="SetMenuWnd",control_type='Window')
        settings_button=settings_menu.child_window(control_type='Button',title="设置")
        settings_button.click_input() 
        time.sleep(2)
        desktop=Desktop(backend='uia')
        settings_window=desktop.window(title="设置",class_name="SettingWnd",control_type="Window")
        return settings_window
    
    @staticmethod
    def log_out(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来PC微信退出登录。\n
        '''
        settings_window=WechatSettings.open_wechat_settings(wechat_path=wechat_path,is_maximize=is_maximize)
        log_out_button=settings_window.window(title="退出登录",control_type="Button")
        log_out_button.click_input()
        time.sleep(2)
        confirm_button=settings_window.window(title="确定",control_type="Button")
        confirm_button.click_input()



class Call():
    @staticmethod
    def voice_call(friend,wechat_path=None,is_maximize:bool=True):
        '''
        friend:好友备注\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来给好友拨打语音电话
        '''
        main_window=Tools.open_dialog_window(friend,wechat_path,is_maximize=is_maximize)[1]  
        Tool_bar=main_window.child_window(found_index=0,title='',control_type='ToolBar')
        voice_call_button=Tool_bar.children(title='语音聊天',control_type='Button')[0]
        time.sleep(2)
        voice_call_button.click_input()

    def video_call(friend,wechat_path=None,is_maximize:bool=True):
        '''
        friend:好友备注.\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来给好友拨打视频电话
            '''
        main_window=Tools.open_dialog_window(friend,wechat_path,is_maximize=is_maximize)[1]  
        Tool_bar=main_window.child_window(found_index=0,title='',control_type='ToolBar')
        voice_call_button=Tool_bar.children(title='视频聊天',control_type='Button')[0]
        time.sleep(2)
        voice_call_button.click_input()


class FriendSettings():
    '''这个模块包括 修改好友备注,获取聊天记录,删除联系人,设为星标朋友,将好友聊天界面置顶\n
    消息免打扰,置顶聊天,清空聊天记录,加入黑名单,推荐给朋友,取消设为星标朋友,取消消息免打扰,\n
    取消置顶聊天,取消聊天界面置顶,移出黑名单,添加好友,获取单个或多个好友微信号共计18项功能\n'''
    @staticmethod
    def pin_friend(friend:str,wechat_path=None,is_maximize:bool=True):
        '''
        friend:好友备注。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来将好友或群聊置顶
        '''
        main_window,chat_window=Tools.open_dialog_window(friend,wechat_path,is_maximize=is_maximize) 
        Tool_bar=chat_window.child_window(found_index=1,title='',control_type='ToolBar')
        Top_button=Tool_bar.children(title='置顶',control_type='Button')[0]
        if Top_button[0].exists():
            Top_button[0].click_input()
            time.sleep(2)
            main_window.close()
        else:
            main_window.click_input()  
            main_window.close()
            raise HaveBeenPinnedError(f"好友'{friend}'已被置顶,无需操作！")
   
    @staticmethod
    def cancel_pin_friend(friend:str,wechat_path=None,is_maximize:bool=True):
        '''
        friend:好友备注\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来取消将好友或群聊置顶
        '''
        main_window,chat_window=Tools.open_dialog_window(friend,wechat_path,is_maximize=is_maximize)
        Tool_bar=chat_window.child_window(found_index=1,title='',control_type='ToolBar')
        Top_button=Tool_bar.children(title='取消置顶',control_type='Button')[0]
        if Top_button[0].exists():
            Top_button[0].click_input()
            time.sleep(2)
            main_window.close()
        else:
            main_window.click_input()  
            main_window.close()
            raise HaveBeenUnpinnedError(f"好友'{friend}'未被置顶,无需操作！")

    @staticmethod    
    def mute_friend_notifications(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来开启好友的消息免打扰
        '''
        friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        mute_checkbox=friend_settings_window.child_window(title="消息免打扰",control_type="CheckBox")
        if mute_checkbox.get_toggle_state():
            friend_settings_window.close()
            time.sleep(2)
            main_window.click_input()  
            main_window.close()
            raise HaveBeenMutedError(f"好友'{friend}'的消息免打扰已开启,无需再开启消息免打扰")
        else:
            mute_checkbox.click_input()
            time.sleep(2)
            main_window.close()

    @staticmethod
    def sticky_friend_on_top(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注\n 
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来将好友的聊天置顶
        '''
        friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        sticky_on_top_checkbox=friend_settings_window.child_window(title="置顶聊天",control_type="CheckBox")
        if sticky_on_top_checkbox.get_toggle_state():
            friend_settings_window.close()
            time.sleep(2)
            main_window.click_input()  
            main_window.close()
            raise HaveBeenStickiedError(f"好友'{friend}'的置顶聊天已设置,无需再设为置顶聊天")
        else:
            sticky_on_top_checkbox.click_input()
            time.sleep(2)
            main_window.close()

    @staticmethod
    def cancel_mute_friend_notifications(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来取消好友的消息免打扰
        '''
        friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        mute_checkbox=friend_settings_window.child_window(title="消息免打扰",control_type="CheckBox")
        if not mute_checkbox.get_toggle_state():
            friend_settings_window.close()
            time.sleep(2)
            main_window.click_input()  
            main_window.close()
            raise HaveBeenUnmutedError(f"好友'{friend}'的消息免打扰未开启,无需再关闭消息免打扰")
        else:
            mute_checkbox.click_input()
            time.sleep(2)
            main_window.close()
    
    @staticmethod
    def cancel_sticky_friend_on_top(friend:str,wechat_path:str=None,is_maximize:bool=True):
        ''' 
        friend:好友备注\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来取消好友聊天置顶
        '''
        friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        sticky_on_top_checkbox=friend_settings_window.child_window(title="置顶聊天",control_type="CheckBox")
        if not sticky_on_top_checkbox.get_toggle_state():
            friend_settings_window.close()
            time.sleep(2)
            main_window.click_input()  
            main_window.close()
            raise HaveBeenUnstickiedError(f"好友'{friend}'的置顶聊天未开启,无需再取消置顶聊天")
        else:
            sticky_on_top_checkbox.click_input()
            time.sleep(2)
            main_window.close()

    @staticmethod
    def clear_friend_chat_history(friend:str,wechat_path:str=None,is_maximize:bool=True):
        ''' 
        friend:好友备注\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来清楚聊天记录\n
        '''
        friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        clear_chat_history_button=friend_settings_window.child_window(title="清空聊天记录",control_type="Button")
        clear_chat_history_button.click_input()
        confirm_button=main_window.child_window(title="清空",control_type="Button")
        confirm_button.click_input()
        main_window.close()
    
    @staticmethod
    def delete_friend(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来删除好友\n
        '''
        menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        delete_friend_item=menu.child_window(title='删除联系人',control_type='MenuItem')
        delete_friend_item.click_input()
        confirm_window=friend_settings_window.child_window(class_name='WeUIDialog',title="",control_type='Pane')
        confirm_buton=confirm_window.child_window(control_type='Button',title='确定')
        confirm_buton.click_input()
        time.sleep(2)
        main_window.close()
    
    @staticmethod
    def add_new_friend(phone_number:str=None,wechat_number:str=None,request_content:str=None,wechat_path:str=None,is_maximize:bool=True):
        '''
        phone_number:手机号\n
        wechat_number:微信号\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来添加新朋友\n
        '''
        desktop=Desktop(backend='uia')
        main_window=Tools.open_wechat(wechat_path,is_maximize=is_maximize)
        pane=main_window.child_window(found_index=2,control_type='Pane',title="")
        toolbar=pane.child_window(control_type='ToolBar',title='导航')
        contacts=toolbar.child_window(title='通讯录',control_type='Button')
        contacts.set_focus()
        contacts.click_input()
        add_friend_button=main_window.child_window(control_type="Button",title="添加朋友")
        add_friend_button.click_input()
        search_bar=main_window.child_window(control_type='Edit',title='微信号/手机号')
        search_bar.click_input()
        if phone_number and not wechat_number:
            search_bar.type_keys(phone_number)
        elif wechat_number and phone_number:
            search_bar.type_keys(wechat_number)
        elif not phone_number and wechat_number:
            search_bar.type_keys(wechat_number)
        else:
            main_window.close()
            raise NoWechat_number_or_Phone_numberError(f'未输入微信号或手机号,请至少输入二者其中一个！')
        search_pane=main_window.child_window(title_re="@str:IDS_FAV_SEARCH_RESULT",control_type='List')
        search_pane.child_window(title_re="搜索",control_type="Text").click_input()
        profile_pane=desktop.window(class_name='ContactProfileWnd',framework_id="Win32",control_type='Pane',title='微信')
        add_to_contacts=profile_pane.child_window(title='添加到通讯录',control_type='Button')
        if add_to_contacts.exists():
            add_to_contacts.click_input()
            query_window=main_window.child_window(title="添加朋友请求",class_name='WeUIDialog',control_type='Window',framework_id='Win32')
            if query_window.exists():
                if request_content:
                    request_content_edit=query_window.child_window(title_re='我是',control_type='Edit')
                    request_content_edit.click_input()
                    pyautogui.hotkey('ctrl','a')
                    pyautogui.press('backspace')
                request_content_edit=query_window.child_window(title='',control_type='Edit',found_index=0)
                request_content_edit.type_keys(request_content)
                confirm_button=query_window.child_window(title="确定",control_type='Button')
                confirm_button.click_input()
                time.sleep(5)
                main_window.close()
        else:
            time.sleep(2)
            profile_pane.close()
            main_window.close()
            raise AlreadyInContactsError(f"该好友已在通讯录中,无需通过该群聊添加！")

    @staticmethod 
    def change_friend_remark_and_tag(friend:str,remark:str,tag:str=None,description:str=None,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来修改好友备注和标签\n
        '''
        if friend==remark:
            raise SameNameError(f"待修改的备注要与先前的备注不同才可以修改！")
        menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        change_remark=menu.child_window(title='设置备注和标签',control_type='MenuItem')
        change_remark.click_input()
        sessionchat=friend_settings_window.child_window(title='设置备注和标签',class_name='WeUIDialog',framework_id='Win32')
        remark_edit=sessionchat.child_window(title=friend,control_type='Edit')
        remark_edit.click_input()
        pyautogui.hotkey('ctrl','a')
        pyautogui.press('backspace')
        remark_edit=sessionchat.child_window(control_type='Edit',found_index=0)
        remark_edit.type_keys(remark)
        if tag:
           tag_set=sessionchat.child_window(title='点击编辑标签',control_type='Button')
           tag_set.click_input()
           confirm_pane=main_window.child_window(title="设置标签",framework_id='Win32',class_name='StandardConfirmDialog')
           edit=confirm_pane.child_window(title='设置标签',control_type='Edit')
           edit.click_input()
           edit.type_keys(tag)
           confirm_pane.child_window(title='确定',control_type='Button').click_input()
        if description:
            description_edit=sessionchat.child_window(control_type='Edit',found_index=1)
            description_edit.click_input()
            pyautogui.hotkey('ctrl','a')
            pyautogui.press('backspace')
            description_edit.type_keys(description)
        confirm=sessionchat.child_window(title='确定',control_type='Button')
        confirm.click_input()
        friend_settings_window.close()
        main_window.click_input()
        main_window.close()

    @staticmethod
    def add_friend_to_blacklist(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来将好友添加至黑名单\n
        '''
        menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        blacklist=menu.child_window(title='加入黑名单',control_type='MenuItem')
        if blacklist.exists():
            blacklist.click_input()
            confirm_window=friend_settings_window.child_window(class_name='WeUIDialog',title="",control_type='Pane')
            confirm_buton=confirm_window.child_window(control_type='Button',title='确定')
            confirm_buton.click_input()
            friend_settings_window.close()
            time.sleep(2)
            main_window.close()
        else:
            friend_settings_window.close()
            time.sleep(2) 
            main_window.click_input() 
            main_window.close()
            raise HaveBeenInBlacklistError(f'好友"{friend}"已位于黑名单中,无需操作!')
    
    @staticmethod
    def move_friend_out_of_blacklist(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来将好友移出黑名单\n
        '''
        menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        blacklist=menu.child_window(title='移出黑名单',control_type='MenuItem')
        if blacklist.exists():
            blacklist.click_input()
            friend_settings_window.close()
            time.sleep(2)   
            main_window.close()
        else:
            friend_settings_window.close()
            time.sleep(2)
            main_window.click_input()  
            main_window.close()
            raise HaveBeenOutofBlacklistError(f"好友'{friend}'未在黑名单中,无需操作！")
        
    @staticmethod
    def set_friend_as_star_friend(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来将好友设置为星标朋友
        '''
        menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        star=menu.child_window(title='设为星标朋友',control_type='MenuItem')
        if star.exists():
            star.click_input()
            friend_settings_window.close()
            time.sleep(2)
            main_window.close()
        else:
            friend_settings_window.close()
            time.sleep(2)
            main_window.click_input()  
            main_window.close()
            raise HaveBeenStaredError(f"好友'{friend}'已设为星标朋友,无需操作！")
            
    
    @staticmethod
    def cancel_set_friend_as_star_friend(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来不再将好友设置为星标朋友\n
        '''
        menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        star=menu.child_window(title='不再设为星标朋友',control_type='MenuItem')
        if star.exists():
            star.click_input()
            friend_settings_window.close()
            time.sleep(2)
            main_window.close()
        else:
            friend_settings_window.close()
            time.sleep(2)
            main_window.click_input()  
            main_window.close()
            raise HaveBeenUnstaredError(f"好友'{friend}'未被设为星标朋友,无需操作！")
    
    @staticmethod
    def change_friend_privacy(friend:str,privacy:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注。\n
        privacy:好友权限,共有：仅聊天',"聊天、朋友圈、微信运动等",'不让他（她）看',"不看他（她）"四种\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来修改好友权限\n
        '''
        privacy_rights=['仅聊天',"聊天、朋友圈、微信运动等",'不让他（她）看',"不看他（她）"]
        if privacy not in privacy_rights:
            raise PrivacytNotCorrectError(f'权限不存在！请按照 仅聊天;聊天、朋友圈、微信运动等;\n不让他（她）看;不看他（她);的四种格式输入privacy')
        menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        privacy_button=menu.child_window(title='设置朋友权限',control_type='MenuItem')
        privacy_button.click_input()
        privacy_window=friend_settings_window.child_window(title='朋友权限',class_name='WeUIDialog',framework_id='Win32')
        if privacy=="仅聊天":
            only_chat=privacy_window.child_window(title='仅聊天',control_type='CheckBox')
            if only_chat.get_toggle_state():
                privacy_window.close()
                friend_settings_window.close()
                main_window.click_input()
                main_window.close()
                raise HaveBeenSetChatonlyError(f"好友'{friend}'权限已被设置为仅聊天")
            else:
                only_chat.click_input()
                sure_button=privacy_window.child_window(control_type='Button',title='确定')
                sure_button.click_input()
                friend_settings_window.close()
                main_window.close()
        elif  privacy=="聊天、朋友圈、微信运动等":
            open_chat=privacy_window.child_window(title="聊天、朋友圈、微信运动等",control_type='CheckBox')
            if open_chat.get_toggle_state():
                privacy_window.close()
                friend_settings_window.close()
                main_window.click_input()
                main_window.close()
            else:
                open_chat.click_input()
                sure_button=privacy_window.child_window(control_type='Button',title='确定')
                sure_button.click_input()
                friend_settings_window.close()
                main_window.close()
        else:
            if privacy=='不让他（她）看':
                unseen_to_him=privacy_window.child_window(title="",control_type='CheckBox',found_index=0)
                if unseen_to_him.exists():
                    if unseen_to_him.get_toggle_state():
                        privacy_window.close()
                        friend_settings_window.close()
                        main_window.click_input()
                        main_window.close()
                        raise HaveBeenSetUnseentohimError(f"好友 {friend}权限已被设置为不让他（她）看")
                    else:
                        unseen_to_him.click_input()
                        sure_button=privacy_window.child_window(control_type='Button',title='确定')
                        sure_button.click_input()
                        friend_settings_window.close()
                        main_window.close()
                else:
                    privacy_window.close()
                    friend_settings_window.close()
                    main_window.click_input()
                    main_window.close()
                    raise HaveBeenSetChatonlyError(f"好友 {friend}已被设置为仅聊天,无法设置为不让他（她）看\n若需将其设置为不让他（她）看,请先将好友设置为：\n聊天、朋友圈、微信运动等")
            if privacy=="不看他（她）":
                dont_see_him=privacy_window.child_window(title="",control_type='CheckBox',found_index=1)
                if dont_see_him.exists():
                    if dont_see_him.get_toggle_state():
                        privacy_window.close()
                        friend_settings_window.close()
                        main_window.click_input()
                        main_window.close()
                        raise HaveBeenSetDontseehimError(f"好友 {friend}权限已被设置为不看他（她）")
                    else:
                        dont_see_him.click_input()
                        sure_button=privacy_window.child_window(control_type='Button',title='确定')
                        sure_button.click_input()
                        friend_settings_window.close()
                        main_window.close()  
                else:
                    privacy_window.close()
                    friend_settings_window.close()
                    main_window.click_input()
                    main_window.close()
                    raise HaveBeenSetChatonlyError(f"好友 {friend}已被设置为仅聊天,无法设置为不看他（她）\n若需将其设置为不看他（她）,请先将好友设置为：\n聊天、朋友圈、微信运动等")    
    
    @staticmethod
    def share_contact(friend:str,others:list[str],wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:被推荐好友备注\n
        others:推荐人备注列表\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来推荐好友给其他人
        '''
        menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        share_contact_choice1=menu.child_window(title='推荐给朋友',control_type='MenuItem')
        share_contact_choice2=menu.child_window(title='把他推荐给朋友',control_type='MenuItem')
        share_contact_choice3=menu.child_window(title='把她推荐给朋友',control_type='MenuItem')
        if share_contact_choice1.exists():
            share_contact_choice1.click_input()
        if share_contact_choice2.exists():
            share_contact_choice2.click_input()
        if share_contact_choice3.exists():
            share_contact_choice3.click_input()
        select_contact_window=friend_settings_window.child_window(control_type='Window',class_name='SelectContactWnd',framework_id='Win32',title="")
        if len(others)>1:
            multi=select_contact_window.child_window(control_type='Button',title='多选')
            multi.click_input()
            send=select_contact_window.child_window(title_re='分别发送',control_type='Button')
        else:
            send=select_contact_window.child_window(title='发送',control_type='Button')
        search=select_contact_window.child_window(title="搜索",control_type='Edit')
        for other_friend in others:
            search.click_input()
            search.type_keys(other_friend,with_spaces=True)
            time.sleep(0.5)
            pyautogui.press('enter')
            pyautogui.hotkey('ctrl','a')
            pyautogui.press('backspace')
            time.sleep(0.5)
        send.click_input()
        friend_settings_window.close()
        main_window.close()

    @staticmethod
    def get_friend_wechat_number(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法根据微信备注获取单个好友的微信号
        '''
        profile_window,main_window=Tools.open_friend_profile(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        wechat_number=profile_window.child_window(control_type='Text',found_index=4).window_text()
        profile_window.close()
        main_window.close()
        return wechat_number

    @staticmethod
    def get_friends_wechat_numbers(friends:list[str],wechat_path:str=None,is_maximize:bool=True):
        '''
        friends:所有带获取微信号的好友的备注列表。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法根据微信备注获取多个好友微信号
        '''
        wechat_numbers=[]
        for friend in friends:
            profile_window,main_window=Tools.open_friend_profile(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
            wechat_number=profile_window.child_window(control_type='Text',found_index=4).window_text()
            wechat_numbers.append(wechat_number)
            profile_window.close()
        wechat_numbers=dict(zip(friends,wechat_numbers))        
        main_window.close()
        return wechat_numbers 



class GroupSettings():

    @staticmethod
    def create_group_chat(friends:list[str],group_name:str=None,wechat_path:str=None,is_maximize:bool=True,messages:list=[]):
        '''
        friends:新群聊的好友备注列表。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        messages:建群后是否发送消息,messages非空列表,在建群后会发送消息
        该方法用来新建群聊
        '''
        if len(friends)<2:
            raise CantCreateGroupError(f'三人不成群,除自身外最少还需要两人才能建群！')
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        cerate_group_chat_button=main_window.child_window(title="发起群聊",control_type="Button")
        cerate_group_chat_button.click_input()
        Add_member_window=main_window.child_window(title='AddMemberWnd',control_type='Window',framework_id='Win32')
        for member in friends:
            search=Add_member_window.child_window(title='搜索',control_type="Edit")
            search.click_input()
            search.type_keys(member,with_spaces=True)
            pyautogui.press("enter")
            pyautogui.press('backspace')
            time.sleep(2)

        confirm=Add_member_window.child_window(title='完成',control_type='Button')
        confirm.click_input()
        time.sleep(10)
        if messages:
            group_edit=main_window.child_window(control_type='Edit',found_index=1)
            for message in message:
                group_edit.type_keys(message)
                pyautogui.hotkey('alt','s')
        if group_name:
            three_points=main_window.child_window(title='聊天信息',control_type='Button')
            three_points.click_input()
            group_settings_window=main_window.child_window(title='SessionChatRoomDetailWnd',control_type='Pane',framework_id='Win32')
            change_group_name_button=group_settings_window.child_window(title='群聊名称',control_type='Button')
            change_group_name_button.click_input()
            pyautogui.hotkey('ctrl','a')
            pyautogui.press('backspace')
            change_group_name_edit=group_settings_window.child_window(control_type='Edit',class_name='EditWnd',framework_id='Win32')
            change_group_name_edit.type_keys(group_name)
            pyautogui.press('enter')
            group_settings_window.close()
        main_window.close()

    @staticmethod
    def change_group_name(group_name:str,change_name:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        change_name:待修改的名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来修改群聊名称\n
        '''
        if group_name==change_name:
            raise SameNameError(f'待修改的群名需与先前的群名不同才可修改！')
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        text=group_chat_settings_window.child_window(title='仅群主或管理员可以修改',control_type='Text')
        if text.exists():
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close()
            raise NoPermissionError(f"你不是'{group_name}'的群主或管理员,无权修改群聊名称")
        else:
            change_group_name_button=group_chat_settings_window.child_window(title='群聊名称',control_type='Button')
            change_group_name_button.click_input()
            change_group_name_edit=group_chat_settings_window.child_window(control_type='Edit',class_name='EditWnd',framework_id='Win32')
            change_group_name_edit.click_input()
            pyautogui.hotkey('ctrl','a')
            pyautogui.press('backspace')
            time.sleep(0.5)
            change_group_name_edit.type_keys(change_name,with_spaces=True)
            pyautogui.press('enter')
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close()

    @staticmethod
    def change_my_nickname_in_group(group_name:str,my_nickname:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        my_nickname:待修改昵称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来修改我在本群的昵称\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        change_my_nickname_button=group_chat_settings_window.child_window(title='我在本群的昵称',control_type='Button')
        change_my_nickname_button.click_input()
        change_my_nickname_edit=group_chat_settings_window.child_window(control_type='Edit',class_name='EditWnd',framework_id='Win32')
        change_my_nickname_edit.click_input()
        pyautogui.hotkey('ctrl','a')
        pyautogui.press('backspace')
        time.sleep(0.5)
        change_my_nickname_edit.type_keys(my_nickname,with_spaces=True)
        pyautogui.press('enter')
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close()

    @staticmethod
    def change_group_remark(group_name:str,group_remark:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        group_remark:群聊备注\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来修改群聊备注\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        change_group_remark_button=group_chat_settings_window.child_window(title='备注',control_type='Button')
        change_group_remark_button.click_input()
        change_group_remark_edit=group_chat_settings_window.child_window(control_type='Edit',class_name='EditWnd',framework_id='Win32')
        change_group_remark_edit.click_input()
        pyautogui.hotkey('ctrl','a')
        pyautogui.press('backspace')
        time.sleep(0.5)
        change_group_remark_edit.type_keys(group_remark,with_spaces=True)
        pyautogui.press('enter')
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close()
    
    @staticmethod
    def show_group_members_nickname(group_name:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来开启显示群聊成员名称\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        show_group_members_nickname_button=group_chat_settings_window.child_window(title='显示群成员昵称',control_type='CheckBox')
        if not show_group_members_nickname_button.get_toggle_state():
            show_group_members_nickname_button.click_input()
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close()
        else:
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close()
            raise AlreadyOpenError(f"显示群成员昵称功能已开启,无需开启")

    @staticmethod
    def dont_show_group_members_nickname(group_name:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来关闭显示群聊成员名称\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        show_group_members_nickname_button=group_chat_settings_window.child_window(title='显示群成员昵称',control_type='CheckBox')
        if not show_group_members_nickname_button.get_toggle_state():
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close()
            raise AlreadyCloseError(f"显示群成员昵称功能已关闭,无需关闭")
        else:
            show_group_members_nickname_button.click_input()
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close()
            

    @staticmethod
    def mute_group_notifications(group_name:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来开启群聊消息免打扰\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        mute_checkbox=group_chat_settings_window.child_window(title="消息免打扰",control_type="CheckBox")
        if mute_checkbox.get_toggle_state():
            group_chat_settings_window.close()
            main_window.click_input()  
            main_window.close()
            raise HaveBeenMutedError(f"群聊'{group_name}'的消息免打扰已开启,无需再开启消息免打扰")
        else:
            mute_checkbox.click_input()
            group_chat_settings_window.close()
            main_window.close() 

    @staticmethod
    def cancel_mute_group_notifications(group_name:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来关闭群聊消息免打扰\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        mute_checkbox=group_chat_settings_window.child_window(title="消息免打扰",control_type="CheckBox")
        if not mute_checkbox.get_toggle_state():
            group_chat_settings_window.close()
            main_window.click_input()  
            main_window.close()
            raise HaveBeenUnmutedError(f"群聊'{group_name}'的消息免打扰未开启,无需再关闭消息免打扰")
        else:
            mute_checkbox.click_input()
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close() 

    @staticmethod
    def sticky_group_on_top(group_name:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来将微信群聊聊天置顶\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        sticky_on_top_checkbox=group_chat_settings_window.child_window(title="置顶聊天",control_type="CheckBox")
        if not sticky_on_top_checkbox.get_toggle_state():
            sticky_on_top_checkbox.click_input()
            group_chat_settings_window.close()
            main_window.click_input()  
            main_window.close()
        else:
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close() 
            raise HaveBeenStickiedError(f"群聊'{group_name}'的置顶聊天已设置,无需再设为置顶聊天")

    @staticmethod
    def cancel_sticky_group_on_top(group_name:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来取消微信群聊聊天置顶\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        sticky_on_top_checkbox=group_chat_settings_window.child_window(title="置顶聊天",control_type="CheckBox")
        if not sticky_on_top_checkbox.get_toggle_state():
            
            group_chat_settings_window.close()
            main_window.click_input()  
            main_window.close()
            raise HaveBeenUnstickiedError(f"群聊'{group_name}'的置顶聊天未开启,无需再取消置顶聊天")
            
        else:
            sticky_on_top_checkbox.click_input()
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close() 

    @staticmethod            
    def save_group_to_contacts(group_name:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来将群聊保存到通讯录\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        save_to_contacts_checkbox=group_chat_settings_window.child_window(title="保存到通讯录",control_type="CheckBox")
        if not save_to_contacts_checkbox.get_toggle_state():
            save_to_contacts_checkbox.click_input()
            group_chat_settings_window.close()
            main_window.click_input()  
            main_window.close()
        else:
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close() 
            raise AlreadyOpenError(f"群聊'{group_name}'已保存到通讯录,无需再保存到通讯录")

    @staticmethod     
    def cancel_save_group_to_contacts(group_name:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来不再将群聊保存至通讯录\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        save_to_contacts_checkbox=group_chat_settings_window.child_window(title="保存到通讯录",control_type="CheckBox")
        if not save_to_contacts_checkbox.get_toggle_state():
        
            group_chat_settings_window.close()
            main_window.click_input()  
            main_window.close()
            raise AlreadyCloseError(f"群聊'{group_name}'未保存到通讯录,无需再取消保存到通讯录")
        else:
            save_to_contacts_checkbox.click_input()
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close() 

    @staticmethod
    def clear_group_chat_history(group_name:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来清空群聊聊天记录\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        clear_chat_history_button=group_chat_settings_window.child_window(title='清空聊天记录',control_type='Button')
        clear_chat_history_button.click_input()
        confirm_button=main_window.child_window(title="清空",control_type="Button")
        confirm_button.click_input()
        main_window.close()

    @staticmethod
    def quit_group_chat(group_name:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来退出微信群聊\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        quit_group_chat_button=group_chat_settings_window.child_window(title='退出群聊',control_type='Button')
        quit_group_chat_button.click_input()
        confirm_button=main_window.child_window(title="退出",control_type="Button")
        confirm_button.click_input()
        main_window.close()

    @staticmethod
    def invite_others_to_group(group_name:str,friends:list[str],wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        friends:所有待邀请好友备注列表\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来邀请他人至群聊\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        add=group_chat_settings_window.child_window(title='',control_type="Button",found_index=1)
        add.click_input()
        Add_member_window=main_window.child_window(title='AddMemberWnd',control_type='Window',framework_id='Win32')
        for member in friends:
            search=Add_member_window.child_window(title='搜索',control_type="Edit")
            search.click_input()
            search.type_keys(member,with_spaces=True)
            pyautogui.press("enter")
            pyautogui.press('backspace')
            time.sleep(2)
        confirm=Add_member_window.child_window(title='完成',control_type='Button')
        confirm.click_input()
        time.sleep(10)
        group_chat_settings_window.close()
        main_window.close()

    @staticmethod
    def remove_friend_from_group(group_name:str,friends:list[str],wechat_path:str=None,is_maximize:bool=True):
        '''
        group_name:群聊名称\n
        friends:所有移出群聊的成员备注列表\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来将群成员移出群聊\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        delete=group_chat_settings_window.child_window(title='',control_type="Button",found_index=2)
        if not delete.exists():
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close()
            raise NoPermissionError(f"你不是'{group_name}'的群主或管理员,无权将好友移出群聊")
        else:
            delete.click_input()
            delete_member_window=main_window.child_window(title='DeleteMemberWnd',control_type='Window',framework_id='Win32')
            for member in friends:
                search=delete_member_window.child_window(title='搜索',control_type="Edit")
                search.click_input()
                search.type_keys(member,with_spaces=True)
                button=delete_member_window.child_window(title=member,control_type='Button')
                button.click_input()
            confirm=delete_member_window.child_window(title="完成",control_type='Button')
            confirm.click_input()
            confirm_dialog_window=delete_member_window.child_window(class_name='ConfirmDialog',framework_id='Win32')
            delete=confirm_dialog_window.child_window(title="删除",control_type='Button')
            delete.click_input()
            group_chat_settings_window.close()
            main_window.close()

    @staticmethod
    def add_friend_from_group(friend:str,group_name:str,request_content:str=None,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:待添加群聊成员群聊中的名称\n
        group_name:群聊名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来添加群成员为好友\n
        '''
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        search=group_chat_settings_window.child_window(title='搜索群成员',control_type="Edit")
        search.click_input()
        search.type_keys(friend,with_spaces=True)
        friend_butotn=group_chat_settings_window.child_window(title=friend,control_type='Button',found_index=1)
        for _ in range(2):
            friend_butotn.click_input()
        contact_window=group_chat_settings_window.child_window(class_name='ContactProfileWnd',framework_id="Win32")
        add_to_contacts_button=contact_window.child_window(title='添加到通讯录',control_type='Button')
        if add_to_contacts_button.exists():
            add_to_contacts_button.click_input()
            query_window=main_window.child_window(title="添加朋友请求",class_name='WeUIDialog',control_type='Window',framework_id='Win32')
            request_content_edit=query_window.child_window(title_re='我是',control_type='Edit')
            request_content_edit.click_input()
            pyautogui.hotkey('ctrl','a')
            pyautogui.press('backspace')
            request_content_edit=query_window.child_window(title='',control_type='Edit',found_index=0)
            request_content_edit.type_keys(request_content)
            confirm_button=query_window.child_window(title="确定",control_type='Button')
            confirm_button.click_input()
            time.sleep(5)
            main_window.close()
        else:
            group_chat_settings_window.close()
            main_window.close()
            raise AlreadyInContactsError(f"好友'{friend}'已在通讯录中,无需通过该群聊添加！")
    @staticmethod
    def edit_group_notice(group_name:str,content:str,wechat_path:str=None,is_maximize:bool=True):
        desktop=Desktop(backend='uia')
        group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
        edit_group_notice_button=group_chat_settings_window.child_window(title='点击编辑群公告',control_type='Button')
        edit_group_notice_button.click_input()
        edit_group_notice_window=desktop.window(title='群公告',framework_id='Win32',class_name='ChatRoomAnnouncementWnd')
        text=edit_group_notice_window.child_window(title='仅群主和管理员可编辑',control_type='Text')
        if text.exists():
            edit_group_notice_window.close()
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close()
            raise NoPermissionError(f"你不是'{group_name}'的群主或管理员,无权发布群公告")
        else:
            edit_board=edit_group_notice_window.child_window(control_type="Edit",found_index=0)
            edit_board.click_input()
            pyautogui.hotkey('ctrl','a')
            pyautogui.press('backspace')
            edit_board.type_keys(content) 
            confirm_button=edit_group_notice_window.child_window(title="完成",control_type='Button')
            confirm_button.click_input()
            confirm_pane=edit_group_notice_window.child_window(title="",class_name='WeUIDialog',framework_id="Win32")
            forward=confirm_pane.child_window(title="发布",control_type='Button')
            forward.click_input()
            time.sleep(2)
            edit_group_notice_window.close()
            group_chat_settings_window.close()
            main_window.click_input()
            main_window.close()



class Contacts():
    @staticmethod
    def get_friends_info(wechat_path:str=None,is_maximize:bool=False):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来获取通讯录中所有的好友名称与昵称。
        '''
        def get_names(ListItem):
            pane=ListItem.children(title="",control_type="Pane")[0]
            pane=pane.children(title="",control_type="Pane")[0]
            pane=pane.children(title="",control_type="Pane")[0]
            names=(pane.children()[0].window_text(),pane.children()[1].window_text())
            return names
        contacts_settings_window,main_window=Tools.open_contacts_settings(wechat_path=wechat_path,is_maximize=is_maximize)
        pane=contacts_settings_window.child_window(found_index=5,title="",control_type='Pane')
        total_number=pane.children()[1].texts()[0]
        total_number=total_number.replace('(','').replace(')','')
        total_number=int(total_number)#好友总数
        #先点击选中第一个好友，并且按下，只有这样，才可以在按下pagedown之后才可以滚动页面，每页可以记录8人
        pane=contacts_settings_window.child_window(found_index=28,title="",control_type='Pane')
        friends_list=pane.child_window(title='',control_type='List')
        friends=friends_list.children(title='',control_type='ListItem')  
        pane=friends[0].children(title="",control_type="Pane")[0]
        pane=pane.children(title="",control_type="Pane")[0]
        checkbox=pane.children(title="",control_type="CheckBox")[0]
        checkbox.click_input()
        pages=total_number//8#点击选中第一个好友后，每一页只能记录8人，pages是总页数，也是pagedown按钮的按下次数
        res=total_number%8#好友人数不是8的整数倍数时，需要处理余数部分
        Names=[]
        for _ in range(pages):
            pane=contacts_settings_window.child_window(found_index=28,title="",control_type='Pane')
            friends_list=pane.child_window(title='',control_type='List')
            friends=friends_list.children(title='',control_type='ListItem')  
            names=[get_names(friend) for friend in friends]
            time.sleep(1)
            pyautogui.press('pagedown')
            Names.extend(names)
        if res:
        #处理余数部分
            pyautogui.press('pagedown')
            time.sleep(1)
            pane=contacts_settings_window.child_window(found_index=28,title="",control_type='Pane')
            friends_list=pane.child_window(title='',control_type='List')
            friends=friends_list.children(title='',control_type='ListItem')  
            names=[get_names(friend) for friend in friends[8-res:8]]
            Names.extend(names)
            contacts_settings_window.close()
            main_window.close()
            contacts=[{'好友昵称':name[1],'好友备注':name[0]}for name in Names]
            contacts_json=json.dumps(contacts,ensure_ascii=False,indent=4)
            return contacts_json
        else:
            contacts_settings_window.close()
            main_window.close()
            contacts=[{'好友昵称':name[1],'好友备注':name[0]}for name in Names]
            contacts_json=json.dumps(contacts,ensure_ascii=False,indent=4)
            return contacts_json
    
    @staticmethod
    def get_WeCom_friends(max_WeCom_num:int=100,wechat_path:str=None,is_maximize:bool=False):
        '''
        max_WeCom_num:最大企业微信好友数量,微信未给企业微信好友一个单独的分区,所以我们要从通讯录列表中查询\n
        查询时需要根据数量滚动页面,这里默认100\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        '''
        def get_weCom_friends_names(listitem):
            weComcontact=[friend for friend in listitem if '@' in friend.window_text() and len(friend.children()[0].children()[1].children())==2]
            wechat_Names=[friend.children()[0].children()[1].children(control_type='Text')[0].window_text() for friend in weComcontact]
            return wechat_Names
        contacts_settings_window,main_window=Tools.open_contacts_settings(wechat_path=wechat_path,is_maximize=is_maximize)
        pane=contacts_settings_window.child_window(found_index=5,title="",control_type='Pane')
        total_number=pane.children()[1].texts()[0]
        total_number=total_number.replace('(','').replace(')','')
        total_number=int(total_number)#好友总数
        total_number+=max_WeCom_num
        contacts_settings_window.close()
        toolbar=main_window.child_window(control_type='ToolBar',title='导航')
        contacts=toolbar.child_window(title='通讯录',control_type='Button')
        contacts.set_focus()
        contacts.click_input()
        contacts_list=main_window.child_window(title='联系人',control_type='List')
        rec=contacts_list.rectangle()  
        mouse.click(coords=(rec.right-5,rec.top+10))
        pages=total_number//12
        res=total_number%12
        contacts_list=main_window.child_window(title='联系人',control_type='List')
        WeCom_names=[]
        WeCom_names.extend(get_weCom_friends_names(contacts_list))
        for _ in range(pages):
            contacts_list=main_window.child_window(title='联系人',control_type='List')
            WeCom_names.extend(get_weCom_friends_names(contacts_list))
            pyautogui.press('pagedown')
        for _ in range(res):
            contacts_list=main_window.child_window(title='联系人',control_type='List')
            WeCom_names.extend(get_weCom_friends_names(contacts_list))
        WeCom_Contacts=list(zip(WeCom_names,WeCom_names))
        mouse.click(coords=(rec.right-5,rec.top+10))
        for _ in range(pages+res):
            pyautogui.press("pageup")
        main_window.close()
        contacts=[{'好友昵称':name[1],'好友备注':name[0]}for name in WeCom_Contacts]
        WeCom_json=json.dumps(contacts,ensure_ascii=False,indent=4)
        return WeCom_json
    
    @staticmethod
    def get_groups_info(wechat_path:str=None,is_maximize:bool=True,max_goup_numbers:int=99):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        max_group_numbers:由于微信未在Ui界面显示群聊总数,在获取群聊信息时,本方法按照传入的max_group_numbers数量计算
        按下pagedown的次数来获取每页的群聊信息\n
        若你加入的群聊非常多可以将其设置为一较大的数。
        max_group_numbers默认为99.\n
        该方法用来获取通讯录中所有的好友名称与昵称。
        '''
        def get_groups_names(group_chat_list):
            names=[chat.children()[0].children()[0].children(control_type="Button")[0].texts()[0] for chat in group_chat_list]
            numbers=[chat.children()[0].children()[0].children()[1].children()[0].children()[1].texts()[0] for chat in group_chat_list]
            numbers=[number.replace('(','').replace(')','') for number in numbers]
            return names,numbers
        def remove_duplicates(lst):
            seen=set()
            result=[]
            for item in lst:
                if item not in seen:
                    seen.add(item)
                    result.append(item)
            return result
        desktop=Desktop(backend='uia')
        main_window=Tools.open_wechat(wechat_path,is_maximize=is_maximize)
        pane=main_window.child_window(found_index=2,control_type='Pane',title="")
        toolbar=pane.child_window(control_type='ToolBar',title='导航')
        contacts=toolbar.child_window(title='通讯录',control_type='Button')
        contacts.set_focus()
        contacts.click_input()
        contacts_settings=main_window.child_window(title='通讯录管理',control_type='Button')#通讯录管理窗口主界面
        contacts_settings.click_input()
        contacts_settings_window=desktop.window(class_name='ContactManagerWindow',title='通讯录管理')
        recent_group_chat=contacts_settings_window.child_window(control_type="Button",title="最近群聊")
        recent_group_chat.set_focus()
        recent_group_chat.click_input()
        group_chat_list=contacts_settings_window.child_window(control_type="List",found_index=0,title="").children(control_type="ListItem",title="")
        first_group=group_chat_list[0].children()[0].children()[0].children(control_type="Button")[0]
        first_group.click_input()
        groups_names=[]
        groups_members=[]
        i=0
        while i<max_goup_numbers//9:
            time.sleep(2)
            group_chat_list=contacts_settings_window.child_window(control_type="List",found_index=0,title="").children(control_type="ListItem",title="")
            names,numbers=get_groups_names(group_chat_list)
            groups_names.extend(names)
            groups_members.extend(numbers)
            pyautogui.press("pagedown")
            i+=1
        contacts_settings_window.close()
        main_window.close()
        groups_names=remove_duplicates(groups_names)
        groups_members=remove_duplicates(groups_members)
        groups_info={"群聊名称":groups_names,"群聊人数":groups_members}
        groups_info_json=json.dumps(groups_info,indent=4,ensure_ascii=False)
        return groups_info_json



class Auto_response():
    @staticmethod
    def auto_answer_call(duration:str,broadcast_content:str,message:str,times:int,wechat_path:str=None):
        '''
        duration:自动接听功能持续时长,格式:s,min,h分别对应秒,分钟,小时,例:duration='1.5h'持续1.5小时\n
        broadcast_content:windowsAPI语音播报内容\n
        message:语音播报结束挂断后,给呼叫者发送的留言\n
        times:语音播报重复次数\n
        注意！一旦开启自动接听功能后,在设定时间内,你的所有视频语音电话都将优先被PC微信接听,并按照设定的播报与留言内容进行播报和留言。
        '''
        def judge_call(call_interface):
            window_text=call_interface.child_window(found_index=3,control_type='Pane').children(control_type='Button')[0].texts()[0]
            if '视频通话' in window_text:
                index=window_text.index("邀")
                caller_name=window_text[0:index]
                return '视频通话',caller_name
            else:
                index=window_text.index("邀")
                caller_name=window_text[0:index]
                return "语音通话",caller_name
        match duration:
            case duration if "s" in duration:
                try:
                    duration=duration.replace('s','')
                    duration=float(duration)
                except ValueError:
                    print("请输入合法时间长度！")
                    return "出现了错误"
            case duration if 'min' in duration:
                try:
                    duration=duration.replace('min','')
                    duration=float(duration)*60
                except ValueError:
                    print("请输入合法时间长度！")
                    return "出现了错误"
            case duration if 'h' in duration:
                try:
                    duration=duration.replace('h','')
                    duration=float(duration)*60*60
                except ValueError:
                    print("请输入合法的时间长度")
                    return "出现了错误"

            case _:
                raise TimeNotCorrectError("请输入合法时间长度！")
        Systemsettings.open_listening_mode()
        start_time=time.time()
        while True:
            if time.time()-start_time<duration:
                
                desktop=Desktop(backend='uia')
                call_interface1=desktop.window(class_name='VoipTrayWnd',title='微信')
                call_interface2=desktop.window(class_name='ILinkVoipTrayWnd',title='微信')
                if call_interface1.exists():
                    flag,caller_name=judge_call(call_interface1)
                    call_window=call_interface1.child_window(found_index=3,title="",control_type='Pane')
                    accept=call_window.children(title='接受',control_type='Button')[0]
                    if flag=="语音通话":
                        time.sleep(2)
                        accept.click_input()
                        accept_call_window=desktop.window(class_name='AudioWnd',title='微信')
                        if accept_call_window.exists():
                            Systemsettings.speaker(times=times,text=broadcast_content)
                            answering_window=accept_call_window.child_window(found_index=13,control_type='Pane',title='')
                            if answering_window.exists():
                                reject=answering_window.children(title='挂断',control_type='Button')[0]
                                reject.click_input()
                        time.sleep(2)
                        Messages.send_message_to_friend(wechat_path=wechat_path,friend=caller_name,message=message)
                            
                    else:
                        accept=call_window.children(title='接受',control_type='Button')[0]
                        time.sleep(2)
                        accept.click_input()
                        time.sleep(3)
                        Systemsettings.speaker(times=times,text=broadcast_content)
                        accept_call_window=desktop.window(class_name='ILinkVoipWnd',title='微信')
                        accept_call_window.click_input()
                        reject_pane=accept_call_window.children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[0].children(title="",control_type="Pane")[2].children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[0]
                        reject=reject_pane.children(control_type='Button',title='挂断')[0]
                        if reject.is_enabled():
                            reject.click_input()
                            Messages.send_message_to_friend(wechat_path=wechat_path,friend=caller_name,message=message)

                elif call_interface2.exists():
                    call_window=call_interface2.child_window(found_index=4,title="",control_type='Pane')
                    accept=call_window.children(title='接受',control_type='Button')[0]
                    flag,caller_name=judge_call(call_interface2)
                    if flag=="语音通话":
                        accept=call_window.children(title='接受',control_type='Button')[0]
                        time.sleep(2)
                        accept.click_input()
                        time.sleep(3)
                        accept_call_window=desktop.window(class_name='ILinkAudioWnd',title='微信')
                        if accept_call_window.exists():
                            answering_window=accept_call_window.child_window(found_index=13,control_type='Pane',title='')
                            Systemsettings.speaker(times=times,text=broadcast_content)
                            if answering_window.exists():
                                reject=answering_window.children(title='挂断',control_type='Button')[0]
                                reject.click_input()
                        time.sleep(2)
                        Messages.send_message_to_friend(wechat_path=wechat_path,friend=caller_name,message=message)
                    else:
                        accept=call_window.children(title='接受',control_type='Button')[0]
                        time.sleep(2)
                        accept.click_input()
                        time.sleep(3)
                        Systemsettings.speaker(times=times,text=broadcast_content)
                        accept_call_window=desktop.window(class_name='ILinkVoipWnd',title='微信')
                        accept_call_window.click_input()
                        reject_pane=accept_call_window.children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[0].children(title="",control_type="Pane")[2].children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[0]
                        reject=reject_pane.children(control_type='Button',title='挂断')[0]
                        if reject.is_enabled():
                            reject.click_input()
                            Messages.send_message_to_friend(wechat_path=wechat_path,friend=caller_name,message=message)
                        
                else:
                    call_interface1=call_interface2=None
            else:
                break
        Systemsettings.close_listening_mode()
    
    @staticmethod
    def auto_response_message(friend:str,duration:str,content:str,wechat_path:str=None,is_maximize:bool=True):
        '''
    friend:好友或群聊备注\n
    duration:自动回复持续时长,格式:'s','min','h'单位:s/秒,min/分,h/小时\n
    content:指定的回复内容，比如:自动回复[微信机器人]:您好,我当前不在,请您稍后再试。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    '''
        def get_new_message(message_list):
            latest_message_list_len=len(message_list.children())
            if latest_message_list_len!=0:
                latest_message=message_list.children()[latest_message_list_len-1]
                who=latest_message.children()[0].children()[0].window_text()
                content=latest_message.window_text()
                return who,content
            else:
                return None,None
        match duration:
            case duration if "s" in duration:
                try:
                    duration=duration.replace('s','')
                    duration=float(duration)
                except ValueError:
                    print("请输入合法的时间长度！")
                    return "错误"
            case duration if 'min' in duration:
                try:
                    duration=duration.replace('min','')
                    duration=float(duration)*60
                except ValueError:
                    print("请输入合法的时间长度！")
                    return "错误"
            case duration if 'h' in duration:
                try:
                    duration=duration.replace('h','')
                    duration=float(duration)*60*60
                except ValueError:
                    print("请输入合法的时间长度！")
                    return "错误"
            case _:
                raise TimeNotCorrectError('请输入合法的时间长度！') 
        chat,main_window=Tools.open_dialog_window(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        Systemsettings.open_listening_mode()
        start_time=time.time() 
        unresponsed=[]
        responsed=[]
        message_list=main_window.child_window(title="消息",control_type="List")
        who,new_message=get_new_message(message_list=message_list)
        responsed.append(content)
        if new_message:
            responsed.append(new_message)
        while True:
            if time.time()-start_time<duration:
                message_list=main_window.child_window(title="消息",control_type="List")
                who,new_message=get_new_message(message_list=message_list)
                unresponsed.append(new_message)
                if new_message:
                    if new_message in unresponsed and not new_message in responsed and who==friend:
                        chat.click_input()
                        chat.type_keys(content)
                        pyautogui.hotkey('alt','s')
            else:
                break
        Systemsettings.close_listening_mode()
        chat.close()
        

def send_message_to_friend(friend:str,message:str,wechat_path:str=None,delay:int=2,is_maximize:bool=True):
    '''
    给单个好友或群聊发送单条信息\n
    friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
    message:待发送消息。格式:message="消息"\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
  
    '''
    chat,main_window=Tools.find_friend_in_MessageList(friend,wechat_path,is_maximize)
    if chat:
        chat.set_focus()
        chat.click_input()
        chat.type_keys(message,with_spaces=True)
        time.sleep(delay)
        pyautogui.hotkey('alt','s')
    else:
        search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
        search.click_input()
        search.type_keys(friend,with_spaces=True)
        time.sleep(delay)
        pyautogui.hotkey('enter')
        chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
        chat.set_focus()
        chat.click_input()
        chat.type_keys(message,with_spaces=True)
        time.sleep(delay)
        pyautogui.hotkey('alt','s')
    time.sleep(2)
    main_window.close()

def send_messages_to_friend(friend:str,messages:list,delay:int=2,wechat_path:str=None,is_maximize:bool=True):
    '''
    给单个好友或群聊发送多条信息\n
    friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
    message:待发送消息列表。格式:message=["发给好友的消息1","发给好友的消息2"]\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    '''
    chat,main_window=Tools.find_friend_in_MessageList(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    if chat:
        for message in messages:
            chat.set_focus()
            chat.click_input()
            chat.type_keys(message,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
    else:
        search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
        search.click_input()
        search.type_keys(friend,with_spaces=True)
        time.sleep(delay)
        pyautogui.hotkey('enter')
        chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
        chat.set_focus()
        chat.click_input()
        for message in messages:
            chat.type_keys(message,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
    time.sleep(2)
    main_window.close()

def send_message_to_friends(friends:list,message:list,wechat_path:str=None,delay:int=2,is_maximize:bool=True):
    '''
    给多个好友或群聊发送单条信息\n
    friends:好友或群聊备注。格式:friends=["好友1","好友2","好友3"]\n
    message:待发送消息,格式: message=[发给好友1的多条消息,发给好友2的多条消息,发给好友3的多条消息]。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    注意!message与friends长度需一致,并且messages内每一条消息顺序需与friends中好友名称出现顺序一致,否则会出现消息发错的尴尬情况\n
    '''
    Chats=dict(zip(friends,message))
    main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    time.sleep(2)
    for friend in Chats:
        search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
        search.click_input()
        search.type_keys(friend,with_spaces=True)
        time.sleep(delay)
        pyautogui.hotkey('enter')
        chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
        chat.set_focus()
        chat.click_input()
        chat.type_keys(Chats.get(friend),with_spaces=True)
        time.sleep(delay)
        pyautogui.hotkey('alt','s')
    time.sleep(2)
    main_window.close()

def send_messages_to_firends(friends:str,messages:list[list[str]],wechat_path:str=None,delay:int=2,is_maximize:bool=True):
    '''
    给多个好友或群聊发送多条信息\n
    friends:好友或群聊备注列表,格式:firends=["好友1","好友2","好友3"]。\n
    messages:待发送消息,格式: message=[[发给好友1的多条消息],[发给好友2的多条消息],[发给好友3的多条信息]]。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    注意!messages与friends长度需一致,并且messages内每一个列表顺序需与friends中好友名称出现顺序一致,否则会出现消息发错的尴尬情况
    '''
    Chats=dict(zip(friends,messages))
    main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    time.sleep(delay)
    for friend in Chats:
        search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
        search.click_input()
        search.type_keys(friend,with_spaces=True)
        time.sleep(2)
        pyautogui.hotkey('enter')
        chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
        chat.set_focus()
        chat.click_input()
        for message in Chats.get(friend):
            chat.type_keys(message,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
    time.sleep(2)
    main_window.close()

def forward_message(friends:list,message:str,wechat_path:str=None):
    '''
    给多个好友或群聊转发单条信息\n
    friends:好友或群聊备注列表。格式:friends=["好友1","好友2","好友3"]\n
    message:待发送消息,格式: message="转发消息"。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    '''
    messages=[message for _ in range(len(friends))]
    Messages.send_message_to_friends(friends,messages,wechat_path)


def send_file_to_friend(friend:str,file_path:str,wechat_path:str=None,delay:int=2,is_maximize:bool=True,with_messages:bool=False,messages:list=[],message_first:bool=False):
    '''
    给单个好友或群聊发送单个文件\n
    friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
    file_path:待发送文件绝对路径。
    messages:与文件一同发送的消息。格式:message=["消息1","消息2","消息3"]\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    with_messages:发送文件时是否给好友发消息。True发送消息,默认为False\n
    messages_first:默认先发送文件后发送消息,messages_first设置为True,先发送消息,后发送文件,\n
    '''
    if Systemsettings.is_empty_file(file_path):
        raise EmptyFileError(f'不能发送空文件！请重新选择文件路径！')
    if not Systemsettings.is_file(file_path):
        raise NotFileError(f'该路径下的内容不是文件无法发送!')
    chat,main_window=Tools.find_friend_in_MessageList(friend,wechat_path,is_maximize=False)
    if chat:
        if is_maximize:
            main_window.maximize()
        chat.set_focus()
        chat.click_input()
        if with_messages and messages:
            if message_first:
                for message in messages:
                    chat.type_keys(message)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
                Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
                   
            else:
                Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
                for message in messages:
                    chat.type_keys(message)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
        else:
            if is_maximize:
                main_window.maximize()
            chat.set_focus()
            chat.click_input()
            Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
            pyautogui.hotkey("ctrl","v")
            time.sleep(delay)
            pyautogui.hotkey('alt','s')

    else:
        if is_maximize:
            main_window.maximize()
        search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
        search.click_input()
        search.type_keys(friend,with_spaces=True)
        time.sleep(delay)
        pyautogui.hotkey('enter')
        chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
        chat.set_focus()
        chat.click_input()
        if with_messages and messages:
            if message_first:
                for message in messages:
                    chat.type_keys(message)
                    time.sleep(delay)
                pyautogui.hotkey('alt','s')
                Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')     
            else:
                Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
                for message in messages:
                    chat.type_keys(message)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s') 
        else:                      
            Systemsettings.copy_file_to_windowsclipboard(file_path=file_path)
            pyautogui.hotkey("ctrl","v")
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
    time.sleep(2)
    main_window.close()

def send_files_to_friend(friend:str,folder_path:str,wechat_path:str=None,delay:int=2,is_maximize:bool=True,with_messages:bool=False,messages:list=[],messages_first:bool=False):
    '''
    给单个好友或群聊发送多个文件\n
    friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
    folder_path:所有待发送文件所处的文件夹的地址。
    messages:与文件一同发送的消息。格式:message=["消息1","消息2","消息3"]\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    with_messages:发送文件时是否给好友发消息。True发送消息,默认为False\n
    messages_first:默认先发送文件后发送消息,messages_first设置为True,先发送消息,后发送文件,\n
    '''
    if not Systemsettings.is_dirctory(folder_path):
        raise NotFolderError(f'给定路径不是文件夹！若需发送多个文件给好友,请将所有待发送文件置于文件夹内,并在此方法中传入文件夹路径')
    files_in_folder=Systemsettings.get_files_in_folder(folder_path=folder_path)
    if not files_in_folder:
        raise EmptyFolderError(f"文件夹内没有文件！请重新选择！")
    def send_files():
        if len(files_in_folder)<=9:
            Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder)
            pyautogui.hotkey("ctrl","v")
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
        else:
            files_num=len(files_in_folder)
            rem=len(files_in_folder)%9
            for i in range(0,files_num,9):
                if i+9<files_num:
                    Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder[i:i+9])
                    pyautogui.hotkey("ctrl","v")
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
            if rem:
                Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder[files_num-rem:files_num])
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
    chat,main_window=Tools.find_friend_in_MessageList(friend,wechat_path,is_maximize=False)
    if chat:
        if is_maximize:
            main_window.maximize()
        chat.set_focus()
        chat.click_input()
        if with_messages and messages:    
            if messages_first:
                for message in messages:
                    chat.type_keys(message)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
                send_files()
            else:
                send_files()
                for message in messages:
                    chat.type_keys(message)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s') 
        else:
            send_files()
    else:
        if is_maximize:
            main_window.maximize()
        search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
        search.click_input()
        search.type_keys(friend,with_spaces=True)
        time.sleep(delay)
        pyautogui.hotkey('enter')
        chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
        chat.set_focus()
        chat.click_input()
        if with_messages and messages:
            if messages_first:
                for message in messages:
                    chat.type_keys(message)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
                send_files()    
            else:
                send_files()
                for message in messages:
                    chat.type_keys(message)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s') 
        else:                      
            send_files()
    time.sleep(2)
    main_window.close()


def send_file_to_friends(friends:list,file_paths:list,wechat_path:str=None,delay:int=2,is_maximize:bool=True,with_messages:bool=False,messages:list[list[str]]=[],message_first:bool=False):
    '''
    给每个好友或群聊发送单个不同的文件或信息\n
    friends:好友或群聊备注。格式:friends=["好友1","好友2","好友3"]\n
    file_paths:待发送文件,格式: file=[发给好友1的单个文件,发给好友2的文件,发给好友3的文件]。\n
    messages:待发送消息，格式:messages=[["发给好友1的多条消息"],["发给好友2的多条消息"],["发给好友3的多条消息"]]
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    注意!messages,filepaths与friends长度需一致,并且messages内每一条消息顺序需与friends中好友名称出现顺序一致,否则会出现消息发错的尴尬情况\n
    '''
    for file_path in file_paths:
        if Systemsettings.is_empty_file(file_path):
            raise EmptyFileError(f'不能发送空文件！请重新选择文件路径！')
        if Systemsettings.is_dirctory(file_path):
            raise NotFileError(f'该路径下的内容不是文件,无法发送!')
        if Systemsettings.is_file(file_path):
            raise NotFileError(f'该路径下的内容不是文件,无法发送!')
    Files=dict(zip(friends,file_paths))
    main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    time.sleep(2)
    if with_messages and messages:
        Chats=dict(zip(friends,messages))
        for friend in Files:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            if message_first:
                messages=Chats.get(friend)
                for message in messages:
                    chat.type_keys(message,with_spaces=True)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
                Systemsettings.copy_file_to_windowsclipboard(Files.get(friend))
                pyautogui.hotkey('ctrl','v')
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
            else:
                Systemsettings.copy_file_to_windowsclipboard(Files.get(friend))
                pyautogui.hotkey('ctrl','v')
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
                messages=Chats.get(friend)
                for message in messages:
                    chat.type_keys(message,with_spaces=True)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
    else:
        for friend in Files:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            Systemsettings.copy_file_to_windowsclipboard(Files.get(friend))
            pyautogui.hotkey('ctrl','v')
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
    time.sleep(2)
    main_window.close()


def send_files_to_friends(friends:list,folder_paths:list,wechat_path:str=None,delay:int=2,is_maximize:bool=True,with_messages:bool=False,messages:list[list[str]]=[],message_first:bool=False):
    '''
    给多个好友或群聊发送多个不同或相同的文件夹内的所有文件.\n
    friends:好友或群聊备注。格式:friends=["好友1","好友2","好友3"]\n
    folder_paths:待发送文件夹路径列表，每个文件夹内可以存放多个文件,格式: FolderPath_list=["","",""]\n
    message_list:待发送消息，格式:message=[[""],[""],[""]]\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    注意! messages,folder_paths与friends长度需一致,并且messages内每一条消息FolderPath_list每一个文件\n
    顺序需与friends中好友名称出现顺序一致,否则会出现消息发错的尴尬情况\n
    '''
    for folder_path in folder_paths:
            if not Systemsettings.is_dirctory(folder_path):
                raise NotFolderError(f'给定路径不是文件夹！若需发送多个文件给好友,请将所有待发送文件置于文件夹内,并在此方法中传入文件夹路径')
    def send_files(folder_path):
        files_in_folder=Systemsettings.get_files_in_folder(folder_path=folder_path)
        if len(files_in_folder)<=9:
            Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder)
            pyautogui.hotkey("ctrl","v")
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
        else:
            files_num=len(files_in_folder)
            rem=len(files_in_folder)%9
            for i in range(0,files_num,9):
                if i+9<files_num:
                    Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder[i:i+9])
                    pyautogui.hotkey("ctrl","v")
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
            if rem:
                Systemsettings.copy_files_to_windowsclipboard(filepaths_list=files_in_folder[files_num-rem:files_num])
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
    Files=dict(zip(friends,folder_paths))
    main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    if with_messages and messages:
        Chats=dict(zip(friends,messages))
        for friend in Files:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            if message_first:
                messages=Chats.get(friend)
                for message in messages:
                    chat.type_keys(message,with_spaces=True)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
                folder_path=Files.get(friend)
                send_files(folder_path)
            else:
                folder_path=Files.get(friend)
                send_files(folder_path)
                messages=Chats.get(friend)
                for message in messages:
                    chat.type_keys(message,with_spaces=True)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
    else:
        for friend in Files:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            folder_path=Files.get(friend)
            send_files(folder_path)
    time.sleep(2)
    main_window.close()


def forward_File(friends:list,file_path:str,message:str,wechat_path:str=None,is_maximize:bool=True,with_messages:bool=False,message_first:bool=False):
    '''
    给多个好友或群聊转发单个文件\n
    friends:好友或群聊备注列表。格式:friends=["好友1","好友2","好友3"]\n
    file_path:待发送文件,格式: file_path="转发文件路径"。\n
    message:转发消息,格式:messaghe="待转发消息"
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    '''
    messages=[[message] for _ in len(friends)]
    file_paths=[file_path for _ in range(len(friends))]
    Files.send_file_to_friends(file_paths=file_paths,wechat_path=wechat_path,is_maximize=is_maximize,message_list=messages,with_messages=with_messages,message_first=message_first)


def get_friends_info(wechat_path:str=None,is_maximize:bool=True):
    '''
    获取通讯录所有好友名称与昵称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    '''
    def get_names(ListItem):
        pane=ListItem.children(title="",control_type="Pane")[0]
        pane=pane.children(title="",control_type="Pane")[0]
        pane=pane.children(title="",control_type="Pane")[0]
        names=(pane.children()[0].window_text(),pane.children()[1].window_text())
        return names
    contacts_settings_window,main_window=Tools.open_contacts_settings(wechat_path=wechat_path,is_maximize=is_maximize)
    pane=contacts_settings_window.child_window(found_index=5,title="",control_type='Pane')
    total_number=pane.children()[1].texts()[0]
    total_number=total_number.replace('(','').replace(')','')
    total_number=int(total_number)#好友总数
    #先点击选中第一个好友，并且按下，只有这样，才可以在按下pagedown之后才可以滚动页面，每页可以记录8人
    pane=contacts_settings_window.child_window(found_index=28,title="",control_type='Pane')
    friends_list=pane.child_window(title='',control_type='List')
    friends=friends_list.children(title='',control_type='ListItem')  
    pane=friends[0].children(title="",control_type="Pane")[0]
    pane=pane.children(title="",control_type="Pane")[0]
    checkbox=pane.children(title="",control_type="CheckBox")[0]
    checkbox.click_input()
    pages=total_number//8#点击选中第一个好友后，每一页只能记录8人，pages是总页数，也是pagedown按钮的按下次数
    res=total_number%8#好友人数不是8的整数倍数时，需要处理余数部分
    Names=[]
    for _ in range(pages):
        pane=contacts_settings_window.child_window(found_index=28,title="",control_type='Pane')
        friends_list=pane.child_window(title='',control_type='List')
        friends=friends_list.children(title='',control_type='ListItem')  
        names=[get_names(friend) for friend in friends]
        time.sleep(1)
        pyautogui.press('pagedown')
        Names.extend(names)
    if res:
    #处理余数部分
        pyautogui.press('pagedown')
        time.sleep(1)
        pane=contacts_settings_window.child_window(found_index=28,title="",control_type='Pane')
        friends_list=pane.child_window(title='',control_type='List')
        friends=friends_list.children(title='',control_type='ListItem')  
        names=[get_names(friend) for friend in friends[8-res:8]]
        Names.extend(names)
        contacts=[{'好友昵称':name[1],'好友备注':name[0]}for name in Names]
        contacts_json=json.dumps(contacts,ensure_ascii=False,indent=4)
        contacts_settings_window.close()
        main_window.close()
        return contacts_json  
    else:
        contacts=[{'好友昵称':name[1],'好友备注':name[0]}for name in Names]
        contacts_json=json.dumps(contacts,ensure_ascii=False,indent=4)
        contacts_settings_window.close()
        main_window.close()  
        return contacts_json                     
     
def get_WeCom_friends(max_WeCom_num:int=100,wechat_path:str=None,is_maximize:bool=False):
    '''
    max_WeCom_num:最大企业微信好友数量,微信未给企业微信好友一个单独的分区,所以我们要从通讯录列表中查询\n
    查询时需要根据数量滚动页面,这里默认100\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    '''
    def get_weCom_friends_names(listitem):
        weComcontact=[friend for friend in listitem if '@' in friend.window_text() and len(friend.children()[0].children()[1].children())==2]
        wechat_Names=[friend.children()[0].children()[1].children(control_type='Text')[0].window_text() for friend in weComcontact]
        return wechat_Names
    contacts_settings_window,main_window=Tools.open_contacts_settings(wechat_path=wechat_path,is_maximize=is_maximize)
    pane=contacts_settings_window.child_window(found_index=5,title="",control_type='Pane')
    total_number=pane.children()[1].texts()[0]
    total_number=total_number.replace('(','').replace(')','')
    total_number=int(total_number)#好友总数
    total_number+=max_WeCom_num
    contacts_settings_window.close()
    toolbar=main_window.child_window(control_type='ToolBar',title='导航')
    contacts=toolbar.child_window(title='通讯录',control_type='Button')
    contacts.set_focus()
    contacts.click_input()
    contacts_list=main_window.child_window(title='联系人',control_type='List')
    rec=contacts_list.rectangle()  
    mouse.click(coords=(rec.right-5,rec.top+10))
    pages=total_number//12
    res=total_number%12
    contacts_list=main_window.child_window(title='联系人',control_type='List')
    WeCom_names=[]
    WeCom_names.extend(get_weCom_friends_names(contacts_list))
    for _ in range(pages):
        contacts_list=main_window.child_window(title='联系人',control_type='List')
        WeCom_names.extend(get_weCom_friends_names(contacts_list))
        pyautogui.press('pagedown')
    for _ in range(res):
        contacts_list=main_window.child_window(title='联系人',control_type='List')
        WeCom_names.extend(get_weCom_friends_names(contacts_list))
    WeCom_Contacts=list(zip(WeCom_names,WeCom_names))
    mouse.click(coords=(rec.right-5,rec.top+10))
    for _ in range(pages+res):
        pyautogui.press("pageup")
    main_window.close()
    contacts=[{'好友昵称':name[1],'好友备注':name[0]}for name in WeCom_Contacts]
    WeCom_json=json.dumps(contacts,ensure_ascii=False,indent=4)
    return WeCom_json


def get_groups_info(wechat_path:str=None,is_maximize:bool=True,max_goup_numbers:int=99):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    max_group_numbers:由于微信未在Ui界面显示群聊总数,在获取群聊信息时,本方法按照传入的max_group_numbers数量计算
    按下pagedown的次数来获取每页的群聊信息\n
    若你加入的群聊非常多可以将其设置为一较大的数。
    max_group_numbers默认为99.\n
    用来获取通讯录中所有的群聊名称与人数。
    '''
    def get_groups_names(group_chat_list):
        names=[chat.children()[0].children()[0].children(control_type="Button")[0].texts()[0] for chat in group_chat_list]
        numbers=[chat.children()[0].children()[0].children()[1].children()[0].children()[1].texts()[0] for chat in group_chat_list]
        numbers=[number.replace('(','').replace(')','') for number in numbers]
        
        return names,numbers
    def remove_duplicates(lst):
        seen=set()
        result=[]
        for item in lst:
            if item not in seen:
                seen.add(item)
                result.append(item)
        return result
    contacts_settings_window,main_window=Tools.open_contacts_settings(wechat_path=wechat_path,is_maximize=is_maximize)
    recent_group_chat=contacts_settings_window.child_window(control_type="Button",title="最近群聊")
    recent_group_chat.set_focus()
    recent_group_chat.click_input()
    group_chat_list=contacts_settings_window.child_window(control_type="List",found_index=0,title="").children(control_type="ListItem",title="")
    first_group=group_chat_list[0].children()[0].children()[0].children(control_type="Button")[0]
    first_group.click_input()
    groups_names=[]
    groups_members=[]
    i=0
    while i<max_goup_numbers//9:
        time.sleep(2)
        group_chat_list=contacts_settings_window.child_window(control_type="List",found_index=0,title="").children(control_type="ListItem",title="")
        names,numbers=get_groups_names(group_chat_list)
        groups_names.extend(names)
        groups_members.extend(numbers)
        pyautogui.press("pagedown")
        i+=1
    groups_names=remove_duplicates(groups_names)
    groups_members=remove_duplicates(groups_members)
    groups_info={"群聊名称":groups_names,"群聊人数":groups_members}
    groups_info_json=json.dumps(groups_info,indent=4,ensure_ascii=False)
    contacts_settings_window.close()
    main_window.close()
    return groups_info_json


def auto_answer_call(duration:str,broadcast_content:str,message:str,times:int,wechat_path:str=None):
    '''
     duration:自动接听功能持续时长,格式:s,min,h分别对应秒,分钟,小时,例:duration='1.5h'\n
     broadcast_content:windowsAPI语音播报内容\n
     message:语音播报结束挂断后,给呼叫者发送的留言\n
     times:语音播报重复次数\n
     注意！一旦开启自动接听功能后,在设定时间内,你的所有视频语音电话都将优先被PC微信接听,并按照设定的播报与留言内容进行播报和留言。
    '''
    def judge_call(call_interface):
        window_text=call_interface.child_window(found_index=3,control_type='Pane').children(control_type='Button')[0].texts()[0]
        if '视频通话' in window_text:
            index=window_text.index("邀")
            caller_name=window_text[0:index]
            return '视频通话',caller_name
        else:
            index=window_text.index("邀")
            caller_name=window_text[0:index]
            return "语音通话",caller_name
    match duration:
        case duration if "s" in duration:
            try:
                duration=duration.replace('s','')
                duration=float(duration)
            except ValueError:
                print("请输入合法的时间长度！")
                return "错误"
        case duration if 'min' in duration:
            try:
                duration=duration.replace('min','')
                duration=float(duration)*60
            except ValueError:
                print("请输入合法的时间长度！")
                return "错误"
        case duration if 'h' in duration:
            try:
                duration=duration.replace('h','')
                duration=float(duration)*60*60
            except ValueError:
                print("请输入合法的时间长度！")
                return "错误"
        case _:
            raise TimeNotCorrectError('请输入合法的时间长度！') 
    Systemsettings.open_listening_mode()
    start_time=time.time()
    while True:
        if time.time()-start_time<duration:
            desktop=Desktop(backend='uia')
            call_interface1=desktop.window(class_name='VoipTrayWnd',title='微信')
            call_interface2=desktop.window(class_name='ILinkVoipTrayWnd',title='微信')
            if call_interface1.exists():
                flag,caller_name=judge_call(call_interface1)
                call_window=call_interface1.child_window(found_index=3,title="",control_type='Pane')
                accept=call_window.children(title='接受',control_type='Button')[0]
                if flag=="语音通话":
                    time.sleep(2)
                    accept.click_input()
                    accept_call_window=desktop.window(class_name='AudioWnd',title='微信')
                    if accept_call_window.exists():
                        Systemsettings.speaker(times=times,text=broadcast_content)
                        answering_window=accept_call_window.child_window(found_index=13,control_type='Pane',title='')
                        if answering_window.exists():
                            reject=answering_window.children(title='挂断',control_type='Button')[0]
                            reject.click_input()
                    time.sleep(2)
                    Messages.send_message_to_friend(friend=caller_name,message=message)
                        
                else:
                    accept=call_window.children(title='接受',control_type='Button')[0]
                    time.sleep(2)
                    accept.click_input()
                    time.sleep(3)
                    Systemsettings.speaker(times=times,text=broadcast_content)
                    accept_call_window=desktop.window(class_name='ILinkVoipWnd',title='微信')
                    accept_call_window.click_input()
                    reject_pane=accept_call_window.children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[0].children(title="",control_type="Pane")[2].children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[0]
                    reject=reject_pane.children(control_type='Button',title='挂断')[0]
                    if reject.is_enabled():
                        reject.click_input()
                    time.sleep(2)
                    Messages.send_message_to_friend(wechat_path=wechat_path,friend=caller_name,message=message)

            elif call_interface2.exists():
                call_window=call_interface2.child_window(found_index=4,title="",control_type='Pane')
                accept=call_window.children(title='接受',control_type='Button')[0]
                flag,caller_name=judge_call(call_interface2)
                if flag=="语音通话":
                    accept=call_window.children(title='接受',control_type='Button')[0]
                    time.sleep(2)
                    accept.click_input()
                    time.sleep(3)
                    accept_call_window=desktop.window(class_name='ILinkAudioWnd',title='微信')
                    if accept_call_window.exists():
                        answering_window=accept_call_window.child_window(found_index=13,control_type='Pane',title='')
                        Systemsettings.speaker(times=times,text=broadcast_content)
                        if answering_window.exists():
                            reject=answering_window.children(title='挂断',control_type='Button')[0]
                            reject.click_input()
                    time.sleep(2)
                    Messages.send_message_to_friend(wechat_path=wechat_path,friend=caller_name,message=message)
                else:
                    accept=call_window.children(title='接受',control_type='Button')[0]
                    time.sleep(2)
                    accept.click_input()
                    time.sleep(3)
                    Systemsettings.speaker(times=times,text=broadcast_content)
                    accept_call_window=desktop.window(class_name='ILinkVoipWnd',title='微信')
                    accept_call_window.click_input()
                    reject_pane=accept_call_window.children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[0].children(title="",control_type="Pane")[2].children(title="",control_type="Pane")[1].children(title="",control_type="Pane")[0]
                    reject=reject_pane.children(control_type='Button',title='挂断')[0]
                    if reject.is_enabled():
                        reject.click_input()
                    time.sleep(2)
                    Messages.send_message_to_friend(wechat_path=wechat_path,friend=caller_name,message=message)
                    
            else:
                call_interface1=call_interface2=None
        else:
            break
    Systemsettings.close_listening_mode()


def open_settings(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    打开微信设置界面。\n
    '''    
    main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(control_type="ToolBar",title='导航')
    setting=Toolbar.children(control_type='Pane',title='')[3].children(control_type='Button',title="设置及其他")[0]
    setting.click_input()
    settings_menu=main_window.child_window(class_name="SetMenuWnd",control_type='Window')
    settings_button=settings_menu.child_window(control_type='Button',title="设置")
    settings_button.click_input() 
    time.sleep(2)
    desktop=Desktop(backend='uia')
    settings_window=desktop.window(title="设置",class_name="SettingWnd",control_type="Window")
    return settings_window
    

def log_out(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    PC微信退出登录。\n
    '''
    settings_window=WechatSettings.open_wechat_settings(wechat_path=wechat_path,is_maximize=is_maximize)
    log_out_button=settings_window.window(title="退出登录",control_type="Button")
    log_out_button.click_input()
    time.sleep(2)
    confirm_button=settings_window.window(title="确定",control_type="Button")
    confirm_button.click_input()

  
def pin_friend(friend,wechat_path=None,is_maximize:bool=True):
    '''
    friend:好友备注。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来将好友或群聊置顶
    '''
    main_window,chat_window=Tools.open_dialog_window(friend,wechat_path,is_maximize=is_maximize) 
    Tool_bar=chat_window.child_window(found_index=1,title='',control_type='ToolBar')
    Top_button=Tool_bar.children(title='置顶',control_type='Button')[0]
    if Top_button[0].exists():
        Top_button[0].click_input()
        time.sleep(2)
        main_window.close()
    else:
        main_window.click_input()  
        main_window.close()
        raise HaveBeenPinnedError(f"好友'{friend}'已被置顶,无需操作！")
    
def cancel_pin_friend(friend,wechat_path=None,is_maximize:bool=True):
    '''
    friend:好友备注\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来取消将好友或群聊置顶
    '''
    main_window,chat_window=Tools.open_dialog_window(friend,wechat_path,is_maximize=is_maximize)
    Tool_bar=chat_window.child_window(found_index=1,title='',control_type='ToolBar')
    Top_button=Tool_bar.children(title='取消置顶',control_type='Button')[0]
    if Top_button[0].exists():
        Top_button[0].click_input()
        time.sleep(2)
        main_window.close()
    else:
        main_window.click_input()  
        main_window.close()
        raise HaveBeenUnpinnedError(f"好友'{friend}'未被置顶,无需操作！")

        
def mute_friend_notifications(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来开启好友的消息免打扰
    '''
    friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    mute_checkbox=friend_settings_window.child_window(title="消息免打扰",control_type="CheckBox")
    if mute_checkbox.get_toggle_state():
        friend_settings_window.close()
        time.sleep(2)
        main_window.click_input()  
        main_window.close()
        raise HaveBeenMutedError(f"好友'{friend}'的消息免打扰已开启,无需再开启消息免打扰")
    else:
        mute_checkbox.click_input()
        time.sleep(2)
        main_window.close()

    
def sticky_friend_on_top(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注\n 
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来将好友的聊天置顶
    '''
    friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    sticky_on_top_checkbox=friend_settings_window.child_window(title="置顶聊天",control_type="CheckBox")
    if sticky_on_top_checkbox.get_toggle_state():
        friend_settings_window.close()
        time.sleep(2)
        main_window.click_input()  
        main_window.close()
        raise HaveBeenStickiedError(f"好友 {friend}的置顶聊天已设置,无需再设为置顶聊天")
    else:
        sticky_on_top_checkbox.click_input()
        time.sleep(2)
        main_window.close()

    
def cancel_mute_friend_notifications(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来取消好友的消息免打扰
    '''
    friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    mute_checkbox=friend_settings_window.child_window(title="消息免打扰",control_type="CheckBox")
    if not mute_checkbox.get_toggle_state():
        friend_settings_window.close()
        time.sleep(2)
        main_window.click_input()  
        main_window.close()
        raise HaveBeenUnmutedError(f"好友 {friend}的消息免打扰未开启,无需再关闭消息免打扰")
    else:
        mute_checkbox.click_input()
        time.sleep(2)
        main_window.close()
    
   
def cancel_sticky_friend_on_top(friend:str,wechat_path:str=None,is_maximize:bool=True):
    ''' 
    friend:好友备注\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来取消好友聊天置顶
    '''
    friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    sticky_on_top_checkbox=friend_settings_window.child_window(title="置顶聊天",control_type="CheckBox")
    if not sticky_on_top_checkbox.get_toggle_state():
        friend_settings_window.close()
        time.sleep(2)
        main_window.click_input()  
        main_window.close()
        raise HaveBeenUnstickiedError(f"好友'{friend}'的置顶聊天未开启,无需再取消置顶聊天")
    else:
        sticky_on_top_checkbox.click_input()
        time.sleep(2)
        friend_settings_window.close()
        main_window.close()


def clear_friend_chat_history(friend:str,wechat_path:str=None,is_maximize:bool=True):
    ''' 
    friend:好友备注\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来清楚聊天记录\n
    '''
    friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    clear_chat_history_button=friend_settings_window.child_window(title="清空聊天记录",control_type="Button")
    clear_chat_history_button.click_input()
    confirm_button=main_window.child_window(title="清空",control_type="Button")
    confirm_button.click_input()
    main_window.close()
    
    
def delete_friend(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来删除好友\n
    '''
    menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    delete_friend_item=menu.child_window(title='删除联系人',control_type='MenuItem')
    delete_friend_item.click_input()
    confirm_window=friend_settings_window.child_window(class_name='WeUIDialog',title="",control_type='Pane')
    confirm_buton=confirm_window.child_window(control_type='Button',title='确定')
    confirm_buton.click_input()
    main_window.close()
    
    
def add_new_friend(phone_number:str=None,wechat_number:str=None,request_content:str=None,wechat_path:str=None,is_maximize:bool=True):
    '''
    phone_number:手机号\n
    wechat_number:微信号\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来添加新朋友\n
    '''
    desktop=Desktop(backend='uia')
    main_window=Tools.open_wechat(wechat_path,is_maximize=is_maximize)
    pane=main_window.child_window(found_index=2,control_type='Pane',title="")
    toolbar=pane.child_window(control_type='ToolBar',title='导航')
    contacts=toolbar.child_window(title='通讯录',control_type='Button')
    contacts.set_focus()
    contacts.click_input()
    add_friend_button=main_window.child_window(control_type="Button",title="添加朋友")
    add_friend_button.click_input()
    search_bar=main_window.child_window(control_type='Edit',title='微信号/手机号')
    search_bar.click_input()
    if phone_number and not wechat_number:
        search_bar.type_keys(phone_number)
    elif wechat_number and phone_number:
        search_bar.type_keys(wechat_number)
    elif not phone_number and wechat_number:
        search_bar.type_keys(wechat_number)
    else:
        main_window.close()
        raise NoWechat_number_or_Phone_numberError(f'未输入微信号或手机号,请至少输入二者其中一个！')
    search_pane=main_window.child_window(title_re="@str:IDS_FAV_SEARCH_RESULT",control_type='List')
    search_pane.child_window(title_re="搜索",control_type="Text").click_input()
    profile_pane=desktop.window(class_name='ContactProfileWnd',framework_id="Win32",control_type='Pane',title='微信')
    add_to_contacts=profile_pane.child_window(title='添加到通讯录',control_type='Button')
    if add_to_contacts.exists():
        add_to_contacts.click_input()
        query_window=main_window.child_window(title="添加朋友请求",class_name='WeUIDialog',control_type='Window',framework_id='Win32')
        if query_window.exists():
            if request_content:
                request_content_edit=query_window.child_window(title_re='我是',control_type='Edit')
                request_content_edit.click_input()
                pyautogui.hotkey('ctrl','a')
                pyautogui.press('backspace')
            request_content_edit=query_window.child_window(title='',control_type='Edit',found_index=0)
            request_content_edit.type_keys(request_content)
            confirm_button=query_window.child_window(title="确定",control_type='Button')
            confirm_button.click_input()
            time.sleep(5)
            main_window.close()
    else:
        time.sleep(2)
        profile_pane.close()
        main_window.close()
        raise AlreadyInContactsError(f"该好友已在通讯录中,无需通过该群聊添加！")
        

    
def change_friend_remark_and_tag(friend:str,remark:str=None,tag:str=None,description:str=None,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来修改好友备注和标签\n
    '''
    if friend==remark:
        raise SameNameError(f"待修改的备注要与先前的备注不同才可以修改！")
    menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    change_remark=menu.child_window(title='设置备注和标签',control_type='MenuItem')
    change_remark.click_input()
    sessionchat=friend_settings_window.child_window(title='设置备注和标签',class_name='WeUIDialog',framework_id='Win32')
    remark_edit=sessionchat.child_window(title=friend,control_type='Edit')
    remark_edit.click_input()
    pyautogui.hotkey('ctrl','a')
    pyautogui.press('backspace')
    remark_edit=sessionchat.child_window(control_type='Edit',found_index=0)
    remark_edit.type_keys(remark)
    if tag:
        tag_set=sessionchat.child_window(title='点击编辑标签',control_type='Button')
        tag_set.click_input()
        confirm_pane=main_window.child_window(title="设置标签",framework_id='Win32',class_name='StandardConfirmDialog')
        edit=confirm_pane.child_window(title='设置标签',control_type='Edit')
        edit.click_input()
        edit.type_keys(tag)
        confirm_pane.child_window(title='确定',control_type='Button').click_input()
    if description:
        description_edit=sessionchat.child_window(control_type='Edit',found_index=1)
        description_edit.click_input()
        pyautogui.hotkey('ctrl','a')
        pyautogui.press('backspace')
        description_edit.type_keys(description)
    confirm=sessionchat.child_window(title='确定',control_type='Button')
    confirm.click_input()
    friend_settings_window.close()
    main_window.click_input()
    main_window.close()


    
def add_friend_to_blacklist(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来将好友添加至黑名单\n
    '''
    menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    blacklist=menu.child_window(title='加入黑名单',control_type='MenuItem')
    if blacklist.exists():
        blacklist.click_input()
        confirm_window=friend_settings_window.child_window(class_name='WeUIDialog',title="",control_type='Pane')
        confirm_buton=confirm_window.child_window(control_type='Button',title='确定')
        confirm_buton.click_input()
        friend_settings_window.close()
        time.sleep(2)
        main_window.close()
    else:
        friend_settings_window.close()
        time.sleep(2) 
        main_window.click_input() 
        main_window.close()
        raise HaveBeenInBlacklistError(f"好友'{friend}'已位于黑名单中,无需操作!")
    
    
def move_friend_out_of_blacklist(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来将好友移出黑名单\n
    '''
    menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    blacklist=menu.child_window(title='移出黑名单',control_type='MenuItem')
    if blacklist.exists():
        blacklist.click_input()
        friend_settings_window.close()
        time.sleep(2)   
        main_window.close()
    else:
        friend_settings_window.close()
        time.sleep(2)
        main_window.click_input()  
        main_window.close()
        raise HaveBeenOutofBlacklistError(f"好友'{friend}'未在黑名单中,无需操作！")
        
    
def set_friend_as_star_friend(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来将好友设置为星标朋友
    '''
    menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    star=menu.child_window(title='设为星标朋友',control_type='MenuItem')
    if star.exists():
        star.click_input()
        friend_settings_window.close()
        time.sleep(2)
        main_window.close()
    else:
        friend_settings_window.close()
        time.sleep(2)
        main_window.click_input()  
        main_window.close()
        raise HaveBeenStaredError(f"好友'{friend}'已设为星标朋友,无需操作！")
            
    
    
def cancel_set_friend_as_star_friend(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来不再将好友设置为星标朋友\n
    '''
    menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    star=menu.child_window(title='不再设为星标朋友',control_type='MenuItem')
    if star.exists():
        star.click_input()
        friend_settings_window.close()
        time.sleep(2)
        main_window.close()
    else:
        friend_settings_window.close()
        time.sleep(2)
        main_window.click_input()  
        main_window.close()
        raise HaveBeenUnstaredError(f"好友'{friend}'未被设为星标朋友,无需操作！")
    

def change_friend_privacy(friend:str,privacy:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注。\n
    privacy:好友权限,共有：仅聊天',"聊天、朋友圈、微信运动等",'不让他（她）看',"不看他（她）"四种\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来修改好友权限\n
    '''
    privacy_rights=['仅聊天',"聊天、朋友圈、微信运动等",'不让他（她）看',"不看他（她）"]
    if privacy not in privacy_rights:
        raise PrivacytNotCorrectError(f'权限不存在！请按照 仅聊天;聊天、朋友圈、微信运动等;\n不让他（她）看;不看他（她);的四种格式输入privacy')
    menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    privacy_button=menu.child_window(title='设置朋友权限',control_type='MenuItem')
    privacy_button.click_input()
    privacy_window=friend_settings_window.child_window(title='朋友权限',class_name='WeUIDialog',framework_id='Win32')
    if privacy=="仅聊天":
        only_chat=privacy_window.child_window(title='仅聊天',control_type='CheckBox')
        if only_chat.get_toggle_state():
            privacy_window.close()
            friend_settings_window.close()
            main_window.click_input()
            main_window.close()
            raise HaveBeenSetChatonlyError(f"好友'{friend}'权限已被设置为仅聊天")
        else:
            only_chat.click_input()
            sure_button=privacy_window.child_window(control_type='Button',title='确定')
            sure_button.click_input()
            friend_settings_window.close()
            main_window.close()
    elif  privacy=="聊天、朋友圈、微信运动等":
        open_chat=privacy_window.child_window(title="聊天、朋友圈、微信运动等",control_type='CheckBox')
        if open_chat.get_toggle_state():
            privacy_window.close()
            friend_settings_window.close()
            main_window.click_input()
            main_window.close()
        else:
            open_chat.click_input()
            sure_button=privacy_window.child_window(control_type='Button',title='确定')
            sure_button.click_input()
            friend_settings_window.close()
            main_window.close()
    else:
        if privacy=='不让他（她）看':
            unseen_to_him=privacy_window.child_window(title="",control_type='CheckBox',found_index=0)
            if unseen_to_him.exists():
                if unseen_to_him.get_toggle_state():
                    privacy_window.close()
                    friend_settings_window.close()
                    main_window.click_input()
                    main_window.close()
                    raise HaveBeenSetUnseentohimError(f"好友 {friend}权限已被设置为不让他（她）看")
                else:
                    unseen_to_him.click_input()
                    sure_button=privacy_window.child_window(control_type='Button',title='确定')
                    sure_button.click_input()
                    friend_settings_window.close()
                    main_window.close()
            else:
                privacy_window.close()
                friend_settings_window.close()
                main_window.click_input()
                main_window.close()
                raise HaveBeenSetChatonlyError(f"好友 {friend}已被设置为仅聊天,无法设置为不让他（她）看\n若需将其设置为不让他（她）看,请先将好友设置为：\n聊天、朋友圈、微信运动等")
        if privacy=="不看他（她）":
            dont_see_him=privacy_window.child_window(title="",control_type='CheckBox',found_index=1)
            if dont_see_him.exists():
                if dont_see_him.get_toggle_state():
                    privacy_window.close()
                    friend_settings_window.close()
                    main_window.click_input()
                    main_window.close()
                    raise HaveBeenSetDontseehimError(f"好友 {friend}权限已被设置为不看他（她）")
                else:
                    dont_see_him.click_input()
                    sure_button=privacy_window.child_window(control_type='Button',title='确定')
                    sure_button.click_input()
                    friend_settings_window.close()
                    main_window.close()  
            else:
                privacy_window.close()
                friend_settings_window.close()
                main_window.click_input()
                main_window.close()
                raise HaveBeenSetChatonlyError(f"好友 {friend}已被设置为仅聊天,无法设置为不看他（她）\n若需将其设置为不看他（她）,请先将好友设置为：\n聊天、朋友圈、微信运动等")    

def get_friend_wechat_number(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    根据微信备注获取单个好友的微信号
    '''
    profile_window,main_window=Tools.open_friend_profile(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    wechat_number=profile_window.child_window(control_type='Text',found_index=4).window_text()
    profile_window.close()
    main_window.close()
    return wechat_number

def get_friends_wechat_numbers(friends:list[str],wechat_path:str=None,is_maximize:bool=True):
        '''
        friends:所有带获取微信号的好友的备注列表。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        根据微信备注获取多个好友微信号
        '''
        wechat_numbers=[]
        for friend in friends:
            profile_window,main_window=Tools.open_friend_profile(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
            wechat_number=profile_window.child_window(control_type='Text',found_index=4).window_text()
            wechat_numbers.append(wechat_number)
            profile_window.close()
        wechat_numbers=dict(zip(friends,wechat_numbers))        
        main_window.close()
        return wechat_numbers 

def share_contact(friend:str,others:list,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:被推荐好友备注\n
    others:推荐人备注列表\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来推荐好友给其他人
    '''
    menu,friend_settings_window,main_window=Tools.open_friend_settings_menu(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    share_contact_choice1=menu.child_window(title='推荐给朋友',control_type='MenuItem')
    share_contact_choice2=menu.child_window(title='把他推荐给朋友',control_type='MenuItem')
    share_contact_choice3=menu.child_window(title='把她推荐给朋友',control_type='MenuItem')
    if share_contact_choice1.exists():
        share_contact_choice1.click_input()
    if share_contact_choice2.exists():
        share_contact_choice2.click_input()
    if share_contact_choice3.exists():
        share_contact_choice3.click_input()
    select_contact_window=friend_settings_window.child_window(control_type='Window',class_name='SelectContactWnd',framework_id='Win32',title="")
    if len(others)>1:
        multi=select_contact_window.child_window(control_type='Button',title='多选')
        multi.click_input()
        send=select_contact_window.child_window(title_re='分别发送',control_type='Button')
    else:
        send=select_contact_window.child_window(title='发送',control_type='Button')
    search=select_contact_window.child_window(title="搜索",control_type='Edit')
    for other_friend in others:
        search.click_input()
        search.type_keys(other_friend,with_spaces=True)
        time.sleep(0.5)
        pyautogui.press('enter')
        pyautogui.hotkey('ctrl','a')
        pyautogui.press('backspace')
        time.sleep(0.5)
    send.click_input()
    friend_settings_window.close()
    main_window.close()

def create_group_chat(friends:list,group_name:str=None,wechat_path:str=None,is_maximize:bool=True,messages:list=[]):
    '''
    friends:新群聊的好友备注列表。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    messages:建群后是否发送消息,messages非空列表,在建群后会发送消息
    用来新建群聊
    '''
    if len(friends)<2:
        raise CantCreateGroupError(f'三人不成群,除自身外最少还需要两人才能建群！')
    main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    cerate_group_chat_button=main_window.child_window(title="发起群聊",control_type="Button")
    cerate_group_chat_button.click_input()
    Add_member_window=main_window.child_window(title='AddMemberWnd',control_type='Window',framework_id='Win32')
    for member in friends:
        search=Add_member_window.child_window(title='搜索',control_type="Edit")
        search.click_input()
        search.type_keys(member,with_spaces=True)
        pyautogui.press("enter")
        pyautogui.press('backspace')
        time.sleep(2)
    confirm=Add_member_window.child_window(title='完成',control_type='Button')
    confirm.click_input()
    time.sleep(10)
    if messages:
        group_edit=main_window.child_window(control_type='Edit',found_index=1)
        for message in message:
            group_edit.type_keys(message)
            pyautogui.hotkey('alt','s')
    if group_name:
        three_points=main_window.child_window(title='聊天信息',control_type='Button')
        three_points.click_input()
        group_chat_settings_window=main_window.child_window(title='SessionChatRoomDetailWnd',control_type='Pane',framework_id='Win32')
        change_group_name_button=group_chat_settings_window.child_window(title='群聊名称',control_type='Button')
        change_group_name_button.click_input()
        change_group_name_edit=group_chat_settings_window.child_window(control_type='Edit',class_name='EditWnd',framework_id='Win32')
        change_group_name_edit.click_input()
        pyautogui.hotkey('ctrl','a')
        pyautogui.press('backspace')
        time.sleep(.51)
        change_group_name_edit.type_keys(group_name)
        pyautogui.press('enter')
        group_chat_settings_window.close()
    main_window.close()

def change_group_name(group_name:str,change_name:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    change_name:待修改的名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来修改群聊名称\n
    '''
    if group_name==change_name:
        raise SameNameError(f'待修改的群名需与先前的群名不同才可修改！')
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    text=group_chat_settings_window.child_window(title='仅群主或管理员可以修改',control_type='Text')
    if text.exists():
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close()
        raise NoPermissionError(f"你不是'{group_name}'的群主或管理员,无权修改群聊名称")
    else:
        change_group_name_button=group_chat_settings_window.child_window(title='群聊名称',control_type='Button')
        change_group_name_button.click_input()
        change_group_name_edit=group_chat_settings_window.child_window(control_type='Edit',class_name='EditWnd',framework_id='Win32')
        change_group_name_edit.click_input()
        pyautogui.hotkey('ctrl','a')
        pyautogui.press('backspace')
        time.sleep(.51)
        change_group_name_edit.type_keys(change_name,with_spaces=True)
        pyautogui.press('enter')
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close()

def change_my_nickname_in_group(group_name:str,my_nickname:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    my_nickname:待修改昵称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来修改我在本群的昵称\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    change_my_nickname_button=group_chat_settings_window.child_window(title='我在本群的昵称',control_type='Button')
    change_my_nickname_button.click_input()
    change_my_nickname_edit=group_chat_settings_window.child_window(control_type='Edit',class_name='EditWnd',framework_id='Win32')
    change_my_nickname_edit.click_input()
    pyautogui.hotkey('ctrl','a')
    pyautogui.press('backspace')
    time.sleep(0.5)
    change_my_nickname_edit.type_keys(my_nickname,with_spaces=True)
    pyautogui.press('enter')
    group_chat_settings_window.close()
    main_window.click_input()
    main_window.close()

def change_group_remark(group_name:str,group_remark:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    group_remark:群聊备注\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来修改群聊备注\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    change_group_remark_button=group_chat_settings_window.child_window(title='备注',control_type='Button')
    change_group_remark_button.click_input()
    change_group_remark_edit=group_chat_settings_window.child_window(control_type='Edit',class_name='EditWnd',framework_id='Win32')
    change_group_remark_edit.click_input()
    pyautogui.hotkey('ctrl','a')
    pyautogui.press('backspace')
    time.sleep(0.5)
    change_group_remark_edit.type_keys(group_remark,with_spaces=True)
    pyautogui.press('enter')
    group_chat_settings_window.close()
    main_window.click_input()
    main_window.close()

def show_group_members_nickname(group_name:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来开启显示群聊成员名称\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    show_group_members_nickname_button=group_chat_settings_window.child_window(title='显示群成员昵称',control_type='CheckBox')
    if not show_group_members_nickname_button.get_toggle_state():
        show_group_members_nickname_button.click_input()
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close()
    else:
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close()
        raise AlreadyOpenError(f"显示群成员昵称功能已开启,无需开启")


def dont_show_group_members_nickname(group_name:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来关闭显示群聊成员名称\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    show_group_members_nickname_button=group_chat_settings_window.child_window(title='显示群成员昵称',control_type='CheckBox')
    if not show_group_members_nickname_button.get_toggle_state():
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close()
        raise AlreadyCloseError(f"显示群成员昵称功能已关闭,无需关闭")
    else:
        show_group_members_nickname_button.click_input()
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close()
        

def mute_group_notifications(group_name:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来开启群聊消息免打扰\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    mute_checkbox=group_chat_settings_window.child_window(title="消息免打扰",control_type="CheckBox")
    if mute_checkbox.get_toggle_state():
        group_chat_settings_window.close()
        main_window.click_input()  
        main_window.close()
        raise HaveBeenMutedError(f"群聊'{group_name}'的消息免打扰已开启,无需再开启消息免打扰")
    else:
        mute_checkbox.click_input()
        group_chat_settings_window.close()
        main_window.close() 

def cancel_mute_group_notifications(group_name:str,wechat_path:str=None,is_maximize:bool=True):                            
    '''
    group_name:群聊名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来关闭群聊消息免打扰\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    mute_checkbox=group_chat_settings_window.child_window(title="消息免打扰",control_type="CheckBox")
    if not mute_checkbox.get_toggle_state():
        group_chat_settings_window.close()
        main_window.click_input()  
        main_window.close()
        raise HaveBeenUnmutedError(f"群聊'{group_name}'的消息免打扰未开启,无需再关闭消息免打扰")
    else:
        mute_checkbox.click_input()
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close() 

def sticky_group_on_top(group_name:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来将微信群聊聊天置顶\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    sticky_on_top_checkbox=group_chat_settings_window.child_window(title="置顶聊天",control_type="CheckBox")
    if not sticky_on_top_checkbox.get_toggle_state():
        sticky_on_top_checkbox.click_input()
        group_chat_settings_window.close()
        main_window.click_input()  
        main_window.close()
    else:
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close() 
        raise HaveBeenStickiedError(f"群聊'{group_name}'的置顶聊天已设置,无需再设为置顶聊天")


def cancel_sticky_group_on_top(group_name:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来不再将微信群聊聊天置顶\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    sticky_on_top_checkbox=group_chat_settings_window.child_window(title="置顶聊天",control_type="CheckBox")
    if not sticky_on_top_checkbox.get_toggle_state():
        
        group_chat_settings_window.close()
        main_window.click_input()  
        main_window.close()
        raise HaveBeenUnstickiedError(f"群聊'{group_name}'的置顶聊天未开启,无需再取消置顶聊天")
           
    else:
        sticky_on_top_checkbox.click_input()
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close() 
        
def save_group_to_contacts(group_name:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来将群聊保存到通讯录\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    save_to_contacts_checkbox=group_chat_settings_window.child_window(title="保存到通讯录",control_type="CheckBox")
    if not save_to_contacts_checkbox.get_toggle_state():
        save_to_contacts_checkbox.click_input()
        group_chat_settings_window.close()
        main_window.click_input()  
        main_window.close()
    else:
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close() 
        raise AlreadyOpenError(f"群聊'{group_name}'已保存到通讯录,无需再保存到通讯录")
    
def cancel_save_group_to_contacts(group_name:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来不再将群聊保存到通讯录\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    save_to_contacts_checkbox=group_chat_settings_window.child_window(title="保存到通讯录",control_type="CheckBox")
    if not save_to_contacts_checkbox.get_toggle_state():
        group_chat_settings_window.close()
        main_window.click_input()  
        main_window.close()
        raise AlreadyCloseError(f"群聊'{group_name}'未保存到通讯录,无需再取消保存到通讯录")
    else:
        save_to_contacts_checkbox.click_input()
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close() 

def clear_group_chat_history(group_name:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用来清空群聊聊天记录\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    clear_chat_history_button=group_chat_settings_window.child_window(title='清空聊天记录',control_type='Button')
    clear_chat_history_button.click_input()
    confirm_button=main_window.child_window(title="清空",control_type="Button")
    confirm_button.click_input()
    main_window.close()

def quit_group_chat(group_name:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用来退出微信群聊\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    quit_group_chat_button=group_chat_settings_window.child_window(title='退出群聊',control_type='Button')
    quit_group_chat_button.click_input()
    confirm_button=main_window.child_window(title="退出",control_type="Button")
    confirm_button.click_input()
    main_window.close()

def invite_others_to_group(group_name:str,friends:list[str],wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    friends:所有待邀请好友备注列表\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用来邀请他人至群聊\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    add=group_chat_settings_window.child_window(title='',control_type="Button",found_index=1)
    add.click_input()
    Add_member_window=main_window.child_window(title='AddMemberWnd',control_type='Window',framework_id='Win32')
    for member in friends:
        search=Add_member_window.child_window(title='搜索',control_type="Edit")
        search.click_input()
        search.type_keys(member,with_spaces=True)
        pyautogui.press("enter")
        pyautogui.press('backspace')
        time.sleep(2)
    confirm=Add_member_window.child_window(title='完成',control_type='Button')
    confirm.click_input()
    time.sleep(10)
    group_chat_settings_window.close()
    main_window.close()

def remove_friend_from_group(group_name:str,friends:list[str],wechat_path:str=None,is_maximize:bool=True):
    '''
    group_name:群聊名称\n
    friends:所有移出群聊的成员备注列表\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来将群成员移出群聊\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    delete=group_chat_settings_window.child_window(title='',control_type="Button",found_index=2)
    if not delete.exists():
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close()
        raise NoPermissionError(f"你不是'{group_name}'的群主或管理员,无权将好友移出群聊")
    else:
        delete.click_input()
        delete_member_window=main_window.child_window(title='DeleteMemberWnd',control_type='Window',framework_id='Win32')
        for member in friends:
            search=delete_member_window.child_window(title='搜索',control_type="Edit")
            search.click_input()
            search.type_keys(member,with_spaces=True)
            button=delete_member_window.child_window(title=member,control_type='Button')
            button.click_input()
        confirm=delete_member_window.child_window(title="完成",control_type='Button')
        confirm.click_input()
        confirm_dialog_window=delete_member_window.child_window(class_name='ConfirmDialog',framework_id='Win32')
        delete=confirm_dialog_window.child_window(title="删除",control_type='Button')
        delete.click_input()
        group_chat_settings_window.close()
        main_window.close()

def add_friend_from_group(friend:str,group_name:str,request_content:str=None,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:待添加群聊成员群聊中的名称\n
    group_name:群聊名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    用来添加群成员为好友\n
    '''
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    search=group_chat_settings_window.child_window(title='搜索群成员',control_type="Edit")
    search.click_input()
    search.type_keys(friend,with_spaces=True)
    friend_butotn=group_chat_settings_window.child_window(title=friend,control_type='Button',found_index=1)
    for _ in range(2):
        friend_butotn.click_input()
    contact_window=group_chat_settings_window.child_window(class_name='ContactProfileWnd',framework_id="Win32")
    add_to_contacts_button=contact_window.child_window(title='添加到通讯录',control_type='Button')
    if add_to_contacts_button.exists():
        add_to_contacts_button.click_input()
        query_window=main_window.child_window(title="添加朋友请求",class_name='WeUIDialog',control_type='Window',framework_id='Win32')
        request_content_edit=query_window.child_window(title_re='我是',control_type='Edit')
        request_content_edit.click_input()
        pyautogui.hotkey('ctrl','a')
        pyautogui.press('backspace')
        request_content_edit=query_window.child_window(title='',control_type='Edit',found_index=0)
        request_content_edit.type_keys(request_content)
        confirm_button=query_window.child_window(title="确定",control_type='Button')
        confirm_button.click_input()
        time.sleep(5)
        main_window.close()
    else:
        group_chat_settings_window.close()
        main_window.close()
        raise AlreadyInContactsError(f"好友'{friend}'已在通讯录中,无需通过该群聊添加！")
       
def create_an_new_note(content:str=None,file:str=None,wechat_path:str=None,is_maximize:bool=True,content_first:bool=True):
    '''
    content:笔记文本内容\n
    file:笔记文件内容\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    content_first:先写文本内容还是先放置文件\n
    用来创建一个新笔记\n
        '''
    main_window=Tools.open_collections(wechat_path=wechat_path,is_maximize=is_maximize)
    create_an_new_note_button=main_window.child_window(title="新建笔记",control_type="Button")
    create_an_new_note_button.click_input()
    desktop=Desktop(backend="uia")
    note_window=desktop.window(title='笔记',class_name="FavNoteWnd",framework_id="Win32")
    if file and content:
        if content_first:
            edit_note=note_window.child_window(control_type='Edit',found_index=0)
            edit_note.click_input()
            edit_note.type_keys(content)
            if Systemsettings.is_empty_file(file):
                note_window.close()
                main_window.close()
                raise EmptyFileError(f"输入的路径下的文件为空!请重试")
            elif Systemsettings.is_dirctory(file):
                files=Systemsettings.get_files_in_folder(file)
                if len(files)>10:
                    print("笔记中最多只能存放10个文件,已为您存放前10个文件")
                    files=files[0:10]
                Systemsettings.copy_files_to_windowsclipboard(files)
                edit_note.click_input()
                pyautogui.hotkey('ctrl','v')
            else:
                Systemsettings.copy_file_to_windowsclipboard(file)
                pyautogui.press('enter')
                edit_note.click_input()
                pyautogui.hotkey('ctrl','v')
            pyautogui.hotkey('ctrl','s') 
            note_window.close()
            main_window.close()
        if not content_first:
            edit_note=note_window.child_window(control_type='Edit',found_index=0)
            edit_note.click_input()
            if Systemsettings.is_empty_file(file):
                note_window.close()
                main_window.close()
                raise EmptyFileError(f"输入的路径下的文件为空!请重试")
            elif Systemsettings.is_dirctory(file):
                files=Systemsettings.get_files_in_folder(file)
                if len(files)>10:
                    print("笔记中最多只能存放10个文件,已为您存放前10个文件")
                    files=files[0:10]
                Systemsettings.copy_files_to_windowsclipboard(files)
                pyautogui.hotkey('ctrl','v')
            else:
                Systemsettings.copy_file_to_windowsclipboard(file)
                pyautogui.hotkey('ctrl','v')
            pyautogui.press('enter')
            edit_note.click_input()
            edit_note.type_keys(content)
            pyautogui.hotkey('ctrl','s')
            note_window.close()
            main_window.close()
    if  not file and content:
        edit_note=note_window.child_window(control_type='Edit',found_index=0)
        edit_note.click_input()
        edit_note.type_keys(content)
        note_window.close()
        main_window.close()
        pyautogui.hotkey('ctrl','s')
    if file and not content:
        edit_note=note_window.child_window(control_type='Edit',found_index=0)
        edit_note.click_input()
        if Systemsettings.is_empty_file(file):
            note_window.close()
            main_window.close()
            raise EmptyFileError(f"输入的路径下的文件为空!请重试")
        elif Systemsettings.is_dirctory(file):
            files=Systemsettings.get_files_in_folder(file)
            if len(files)>10:
                print("笔记中最多只能存放10个文件,已为您存放前10个文件")
                files=files[0:10]
            edit_note=note_window.child_window(control_type='Edit',found_index=0)
            Systemsettings.copy_files_to_windowsclipboard(files)
            pyautogui.hotkey('ctrl','v')
        else:
            edit_note=note_window.child_window(control_type='Edit',found_index=0)
            Systemsettings.copy_file_to_windowsclipboard(file)
            pyautogui.hotkey('ctrl','v')
        pyautogui.hotkey('ctrl','s')
        note_window.close()
        main_window.close()
    if not file and not content:
        raise EmptyNoteError(f"笔记中至少要有文字和文件中的一个！")

def edit_group_notice(group_name:str,content:str,wechat_path:str=None,is_maximize:bool=True):
    desktop=Desktop(backend='uia')
    group_chat_settings_window,main_window=Tools.open_group_settings(group=group_name,wechat_path=wechat_path,is_maximize=is_maximize)
    edit_group_notice_button=group_chat_settings_window.child_window(title='点击编辑群公告',control_type='Button')
    edit_group_notice_button.click_input()
    edit_group_notice_window=desktop.window(title='群公告',framework_id='Win32',class_name='ChatRoomAnnouncementWnd')
    text=edit_group_notice_window.child_window(title='仅群主和管理员可编辑',control_type='Text')
    if text.exists():
        edit_group_notice_window.close()
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close()
        raise NoPermissionError(f"你不是'{group_name}'的群主或管理员,无权发布群公告")
    else:
        edit_board=edit_group_notice_window.child_window(control_type="Edit",found_index=0)
        edit_board.click_input()
        pyautogui.hotkey('ctrl','a')
        pyautogui.press('backspace')
        edit_board.type_keys(content) 
        confirm_button=edit_group_notice_window.child_window(title="完成",control_type='Button')
        confirm_button.click_input()
        confirm_pane=edit_group_notice_window.child_window(title="",class_name='WeUIDialog',framework_id="Win32")
        forward=confirm_pane.child_window(title="发布",control_type='Button')
        forward.click_input()
        time.sleep(2)
        edit_group_notice_window.close()
        group_chat_settings_window.close()
        main_window.click_input()
        main_window.close()

   




def auto_response_message(friend:str,duration:str,content:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友或群聊备注\n
    duration:自动回复持续时长,格式:'s','min','h'单位:s/秒,min/分,h/小时\n
    content:指定的回复内容，比如:自动回复[微信机器人]:您好,我当前不在,请您稍后再试。\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    '''
    def get_new_message(message_list):
        latest_message_list_len=len(message_list.children())
        if latest_message_list_len!=0:
            latest_message=message_list.children()[latest_message_list_len-1]
            who=latest_message.children()[0].children()[0].window_text()
            content=latest_message.window_text()
            return who,content
        else:
            return None,None
    match duration:
        case duration if "s" in duration:
            try:
                duration=duration.replace('s','')
                duration=float(duration)
            except ValueError:
                print("请输入合法的时间长度！")
                return "错误"
        case duration if 'min' in duration:
            try:
                duration=duration.replace('min','')
                duration=float(duration)*60
            except ValueError:
                print("请输入合法的时间长度！")
                return "错误"
        case duration if 'h' in duration:
            try:
                duration=duration.replace('h','')
                duration=float(duration)*60*60
            except ValueError:
                print("请输入合法的时间长度！")
                return "错误"
        case _:
            raise TimeNotCorrectError('请输入合法的时间长度！') 
    chat,main_window=Tools.open_dialog_window(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    Systemsettings.open_listening_mode()
    start_time=time.time() 
    unresponsed=[]
    responsed=[]
    message_list=main_window.child_window(title="消息",control_type="List")
    who,new_message=get_new_message(message_list=message_list)
    responsed.append(content)
    if new_message:
        responsed.append(new_message)
    while True:
        if time.time()-start_time<duration:
            message_list=main_window.child_window(title="消息",control_type="List")
            who,new_message=get_new_message(message_list=message_list)
            unresponsed.append(new_message)
            if new_message:
                if new_message in unresponsed and not new_message in responsed and who==friend:
                    chat.click_input()
                    chat.type_keys(content)
                    pyautogui.hotkey('alt','s')
        else:
            break
    Systemsettings.close_listening_mode()
    chat.close()
    
