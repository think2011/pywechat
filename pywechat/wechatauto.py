'''目前实现的功能启动微信，发送消息(四种),拨打电话,自动接听(视频不支持自动挂断),获取通讯录信息'''
'''还需设计的功能:发送功能(发送文件),获取聊天记录(语音转文字)，打开小程序(接口)以及上述功能的定时
监听消息(自动接听电话(已实现)，自动回复,自动接受转账，自动领红包),朋友圈(点赞)。'''
#########################################依赖环境
import time
import pyautogui
import pandas as pd
from pywinauto import Desktop
from pywechat.wechatTools import Tools
from pywechat.winSettings import Systemsettings
from pywechat.Errors import TimenotCorrectError
from pywechat.Errors import HavebeenPinnedError
##############################################
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.width', 280)
pd.set_option("display.max_rows",10000)
pd.set_option("display.max_columns",50)
pyautogui.FAILSAFE = False
class Messages():
    @staticmethod
    def send_message_to_friend(friend:str,message:str,wechat_path:str=None,delay:int=2,is_maximize:bool=True):
        '''
        给单个好友或群聊发送单条信息\n
        friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
        message:待发送消息。格式:message="消息"\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        '''
        chat,main_window=Tools.find_friend_in_MessageList(friend,wechat_path,is_maximize)
        if chat:
            chat.set_focus()
            chat.type_keys(message,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
        else:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.type_keys(message,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()
    @staticmethod
    def send_messages_to_friend(friend:str,messages:list,delay:int=2,wechat_path:str=None,is_maximize:bool=True):
        '''
        给单个好友或群聊发送多条信息\n
        friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
        message:待发送消息列表。格式:message=["发给好友的消息1","发给好友的消息2"]\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        '''
        chat,main_window=Tools.find_friend_in_MessageList(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        if chat:
            for message in messages:
                chat.set_focus()
                chat.type_keys(message,with_spaces=True)
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
        else:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            for message in messages:
                chat.set_focus()
                chat.type_keys(message,with_spaces=True)
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()
    @staticmethod
    def send_messages_to_firends(friends:str,messages:list[list[str]],wechat_path:str=None,delay:int=2,is_maximize:bool=True):
        '''
        给多个好友或群聊发送多条信息\n
        friends:好友或群聊备注列表,格式:firends=["好友1","好友2","好友3"]。\n
        messages:待发送消息,格式: message=[[发给好友1的多条消息],[发给好友2的多条消息],[发给好友3的多条信息]]。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        注意!messages与friends长度需一致,并且messages内每一个列表顺序需与friends中好友名称出现顺序一致,否则会出现消息发错的尴尬情况
        '''
        Chats=dict(zip(friends,messages))
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        time.sleep(2)
        for friend in Chats:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend)
            time.sleep(2)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            for message in Chats.get(friend):
                chat.type_keys(message,with_spaces=True)
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()
    @staticmethod
    def send_message_to_friends(friends:list,message:list,wechat_path:str=None,delay:int=2,is_maximize:bool=True):
        '''
        给多个好友或群聊发送单条信息\n
        friends:好友或群聊备注。格式:friends=["好友1","好友2","好友3"]\n
        message:待发送消息,格式: message=[发给好友1的多条消息,发给好友2的多条消息,发给好友3的多条消息]。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        注意!message与friends长度需一致,并且messages内每一条消息顺序需与friends中好友名称出现顺序一致,否则会出现消息发错的尴尬情况\n
        '''
        Chats=dict(zip(friends,message))
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        time.sleep(2)
        for friend in Chats:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.type_keys(Chats.get(friend),with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()
    @staticmethod
    def forward_message(friends:list,message:str,wechat_path:str=None):
        '''
        给多个好友或群聊转发单条信息\n
        friends:好友或群聊备注列表。格式:friends=["好友1","好友2","好友3"]\n
        message:待发送消息,格式: message="转发消息"。\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        '''
        messages=[message for _ in range(len(friends))]
        Messages.send_message_to_friends(friends,messages,wechat_path)

class Files():
    @staticmethod
    def send_File_to_friend(friend:str,file_path:list,wechat_path:str=None,delay:int=2,is_maximize:bool=True):
        '''
        给单个好友或群聊发送信息\n
        friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
        message:待发送消息。格式:message="消息"\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        '''
        chat,main_window=Tools.find_friend_in_MessageList(friend,wechat_path,is_maximize)
        if chat:
            chat.set_focus()
            chat.click_input()
            Systemsettings.copy_files_to_windowsclipboard(filepaths_list=file_path)
            pyautogui.hotkey("ctrl","v")
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
        else:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            Systemsettings.copy_files_to_windowsclipboard(filepaths_list=file_path)
            pyautogui.hotkey("ctrl","v")
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()

class Call():

    def voice_call(friend,wechat_path=None):
        main_window=Tools.open_dialog_window(friend,wechat_path)[1]  
        Tool_bar=main_window.child_window(found_index=0,title='',control_type='ToolBar')
        voice_call_button=Tool_bar.children(title='语音聊天',control_type='Button')[0]
        time.sleep(2)
        voice_call_button.click_input()

    def video_call(friend,wechat_path=None):
        main_window=Tools.open_dialog_window(friend,wechat_path)[1]  
        Tool_bar=main_window.child_window(found_index=0,title='',control_type='ToolBar')
        voice_call_button=Tool_bar.children(title='视频聊天',control_type='Button')[0]
        time.sleep(2)
        voice_call_button.click_input()


class friendSettings():
    '''这个模块包括 将好友置顶，修改好友备注,获取聊天记录,删除联系人,设为星标朋友,消息免打扰,置顶聊天,清空聊天记录,加入黑名单,推荐给朋友10个功能'''
    @staticmethod
    def Pin_friend(friend,wechat_path=None,is_maximize:bool=True):
        '''
        friend:好友或群聊名称。\n
        wechat_path:微信WeChat.exe的文件地址,若已添加到windows环境变量中该参数默认为None,不需要传入该参数。\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        用来将好友或群聊置顶
        '''
        chat_window=Tools.open_dialog_window(friend,wechat_path,is_maximize=is_maximize)[1] 
        Tool_bar=chat_window.child_window(found_index=1,title='',control_type='ToolBar')
        try:
            Top_button=Tool_bar.children(title='置顶',control_type='Button')[0]
            Top_button[0].click_input()
        except IndexError:
            raise HavebeenPinnedError(f"好友{friend}已被置顶,无需操作！")


class Contacts():
    @staticmethod
    def get_friends_info(wechat_path:str=None,is_maximize:bool=True):
        '''打开通讯录,点击通讯录管理,弹出页面后不放大,点击选中第一个好友后。
        最多支持显示8个好友(这与屏幕大小无关),然后一直按pagedown,每次更新8个,记录下来'''
        def get_names(ListItem):
            pane=ListItem.children(title="",control_type="Pane")[0]
            pane=pane.children(title="",control_type="Pane")[0]
            pane=pane.children(title="",control_type="Pane")[0]
            names=(pane.children()[0].window_text(),pane.children()[1].window_text())
            return names
        desktop=Desktop(backend='uia')
        main_window=Tools.open_wechat(wechat_path,is_maximize=is_maximize)
        pane=main_window.child_window(found_index=2,control_type='Pane',title="")
        toolbar=pane.child_window(control_type='ToolBar',title='导航')
        contacts=toolbar.child_window(title='通讯录',control_type='Button')
        contacts.set_focus()
        contacts.click_input()
        time.sleep(2)
        #打开通讯录管理主界面
        contact_pane=pane.children()[1]
        contact_pane=contact_pane.children()[1]
        contact_pane=contact_pane.children()[0]
        contact_list=contact_pane.children()[0]
        contact_pane=contact_list.children()[0].children()[0]
        contacts_settings=contact_pane.children(title='通讯录管理',control_type='Button')[0]#通讯录管理窗口主界面
        contacts_settings.set_focus()
        contacts_settings.click_input()
        #定位全部后边的好友总数
        contacts_settings_window=desktop.window(class_name='ContactManagerWindow',title='通讯录管理')
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
            contacts={'昵称':[name[1] for name in Names],'备注':[name[0] for name in Names]}
            contacts=pd.DataFrame(contacts)
            return contacts
        else:
            contacts_settings_window.close()
            main_window.close()
            contacts={'昵称':[name[1] for name in Names],'备注':[name[0] for name in Names]}
            contacts=pd.DataFrame(contacts)
            return contacts
    @staticmethod
    def get_group_info():
        pass
class auto_response():
    def auto_receive_transfer():
        pass
    def auto_answer_call(duration:str,broadcast_content:str,message:str,times:int):
        '''
        duration:自动接听功能持续时长,格式:s,m,h分别对应秒,分钟,小时,例:duration='1.5h'\n
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
                raise TimenotCorrectError("请输入合法时间长度！")
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
                            Messages.send_message_to_friend(friend=caller_name,message=message)

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
                            Messages.send_message_to_friend(friend=caller_name,message=message)
                        
                else:
                    call_interface1=call_interface2=None
            else:
                break
        Systemsettings.close_listening_mode()
def send_message_to_friend(friend:str,message:str,wechat_path:str=None,delay:int=2,is_maximize:bool=True):
    '''
    给单个好友或群聊发送单条信息\n
    friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
    message:待发送消息。格式:message="消息"\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    '''
    chat,main_window=Tools.find_friend_in_MessageList(friend,wechat_path,is_maximize)
    if chat:
        chat.set_focus()
        chat.type_keys(message,with_spaces=True)
        time.sleep(delay)
        pyautogui.hotkey('alt','s')
    else:
        search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
        search.click_input()
        search.type_keys(friend)
        time.sleep(delay)
        pyautogui.hotkey('enter')
        chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
        chat.set_focus()
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
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    '''
    chat,main_window=Tools.find_friend_in_MessageList(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    if chat:
        for message in messages:
            chat.set_focus()
            chat.type_keys(message,with_spaces=True)
            time.sleep(delay)
            pyautogui.hotkey('alt','s')
    else:
        search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
        search.click_input()
        search.type_keys(friend)
        time.sleep(delay)
        pyautogui.hotkey('enter')
        chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
        for message in messages:
            chat.set_focus()
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
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    注意!message与friends长度需一致,并且messages内每一条消息顺序需与friends中好友名称出现顺序一致,否则会出现消息发错的尴尬情况\n
    '''
    Chats=dict(zip(friends,message))
    main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    time.sleep(2)
    for friend in Chats:
        search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
        search.click_input()
        search.type_keys(friend)
        time.sleep(delay)
        pyautogui.hotkey('enter')
        chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
        chat.set_focus()
        chat.type_keys(Chats.get(friend),with_spaces=True)
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
    delay:发送单条消息延迟,单位:秒/s,默认2s。\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    '''
    messages=[message for _ in range(len(friends))]
    Messages.send_message_to_friends(friends,messages,wechat_path)

def get_contacts(wechat_path:str=None,is_maximize:bool=True):
    '''打开通讯录,点击通讯录管理,弹出页面后不放大,点击选中第一个好友后。
    最多支持显示8个好友(这与屏幕大小无关),然后一直按pagedown,每次更新8个,记录下来'''
    def get_names(ListItem):
        pane=ListItem.children(title="",control_type="Pane")[0]
        pane=pane.children(title="",control_type="Pane")[0]
        pane=pane.children(title="",control_type="Pane")[0]
        names=(pane.children()[0].window_text(),pane.children()[1].window_text())
        return names
    desktop=Desktop(backend='uia')
    main_window=Tools.open_wechat(wechat_path,is_maximize=is_maximize)
    pane=main_window.child_window(found_index=2,control_type='Pane',title="")
    toolbar=pane.child_window(control_type='ToolBar',title='导航')
    contacts=toolbar.child_window(title='通讯录',control_type='Button')
    contacts.set_focus()
    contacts.click_input()
    time.sleep(2)
    #打开通讯录管理主界面
    contact_pane=pane.children()[1]
    contact_pane=contact_pane.children()[1]
    contact_pane=contact_pane.children()[0]
    contact_list=contact_pane.children()[0]
    contact_pane=contact_list.children()[0].children()[0]
    contacts_settings=contact_pane.children(title='通讯录管理',control_type='Button')[0]#通讯录管理窗口主界面
    contacts_settings.set_focus()
    contacts_settings.click_input()
    #定位全部后边的好友总数
    contacts_settings_window=desktop.window(class_name='ContactManagerWindow',title='通讯录管理')
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
        contacts={'昵称':[name[1] for name in Names],'备注':[name[0] for name in Names]}
        contacts=pd.DataFrame(contacts)
        main_window.close()
        return contacts
    else:
        contacts_settings_window.close()
        main_window.close()
        contacts={'昵称':[name[1] for name in Names],'备注':[name[0] for name in Names]}
        contacts=pd.DataFrame(contacts)
        main_window.close()
        return contacts
    
    
def get_group_info():
    pass

def auto_answer_call(duration:str,broadcast_content:str,message:str,times:int):
    '''
     duration:自动接听功能持续时长,格式:s,m,h分别对应秒,分钟,小时,例:duration='1.5h'\n
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
            raise TimenotCorrectError('请输入合法的时间长度！') 
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
                        Messages.send_message_to_friend(friend=caller_name,message=message)

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
                        Messages.send_message_to_friend(friend=caller_name,message=message)
                    
            else:
                call_interface1=call_interface2=None
        else:
            break
    Systemsettings.close_listening_mode()

def send_File_to_friend(friend:str,file_path:list,wechat_path:str=None,delay:int=2,is_maximize:bool=True,with_messages:bool=False,messages:list=[]):
        '''
        给单个好友或群聊发送信息\n
        friend:好友或群聊备注。格式:friend="好友或群聊备注"\n
        message:待发送消息。格式:message="消息"\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        delay:发送单条消息延迟,单位:秒/s,默认2s。\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        with_messages
        '''
        chat,main_window=Tools.find_friend_in_MessageList(friend,wechat_path,is_maximize)
        if chat:
            if with_messages and messages:
                chat.set_focus()
                chat.click_input()
                Systemsettings.copy_files_to_windowsclipboard(filepaths_list=file_path)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
                for message in messages:
                    chat.type_keys(message)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s')
            else:
                chat.set_focus()
                chat.click_input()
                Systemsettings.copy_files_to_windowsclipboard(filepaths_list=file_path)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')

        else:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend)
            time.sleep(delay)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            chat.set_focus()
            chat.click_input()
            if with_messages and messages:
                Systemsettings.copy_files_to_windowsclipboard(filepaths_list=file_path)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
                for message in messages:
                    chat.type_keys(message)
                    time.sleep(delay)
                    pyautogui.hotkey('alt','s') 
            else:                      
                Systemsettings.copy_files_to_windowsclipboard(filepaths_list=file_path)
                pyautogui.hotkey("ctrl","v")
                time.sleep(delay)
                pyautogui.hotkey('alt','s')
        time.sleep(2)
        main_window.close()
import os
os.startfile("E:\OneDrive\Desktop\教程6.png")