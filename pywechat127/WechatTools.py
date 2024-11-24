''''''
############################依赖环境###########################
import os
import time
import psutil
import pyautogui
import win32gui
import subprocess
from pywinauto import mouse
from pywinauto import Desktop
from pywechat127.Errors import PathNotFoundError
from pywechat127.Errors import NoSuchFriendError
from pywechat127.Errors import WechatNotStartError
from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError
from pywechat127.Uielements import Login_window,Main_window,ToolBar,Independent_window
from pywechat127.WinSettings import Systemsettings
#############################################################
pyautogui.FAILSAFE = False#防止鼠标在屏幕边缘处造成的误触
os.environ.update({"__COMPAT_LAYER":"RUnAsInvoker"})#添加管理员权限
class Tools():
    '''该模块中封装了3个关于PC微信的工具,包括:\n
    is_wechat_running:用来判断PC微信是否运行。\n
    open_wechat:打开PC微信。\n
    find_friends_in_Messagelist:在会话列表和当前聊天窗口中查找好友\n
    以及10个open方法用于打开微信主界面内所有能打开的界面\n
    ''' 
    @staticmethod
    def is_wechat_running():
        '''这个函数通过检测当前windows进程中\n
        是否有wechatappex该项进程来判断微信是否在运行'''
        flag=0
        for process in psutil.process_iter(['name']):
            if 'WeChatAppEx' in process.info['name']:
                    flag=1
                    break
        if flag:#count不为0，表明微信在运行，即微信已登录
            return True
        else:#count为0,表明微信不在运行，即微信未登录
            return False
    @staticmethod
    def find_wechat_path():
        wechat_path=None
        for process in psutil.process_iter(['name','exe']):
            if 'WeChat.exe' in process.info['name']:
                wechat_path=process.info['exe']
                Systemsettings.copy_text_to_windowsclipboard(wechat_path)
                break
        if wechat_path:
            print("已将微信程序路径复制到剪贴板")
            return wechat_path
        else:
            pyautogui.press('esc')
            pyautogui.hotkey('alt','f4')
            pyautogui.hotkey('alt','f4')
            raise WechatNotStartError(f'微信尚未启动,无法寻找到路径,请先启动后再试!')
    @staticmethod
    def set_wechat_as_environ_path():
        subprocess.Popen(["SystemPropertiesAdvanced.exe"])
        pyautogui.hotkey('alt','n',interval=0.5)
        pyautogui.hotkey('alt','n',interval=0.5)
        pyautogui.press('shift')
        pyautogui.typewrite('wechatpath')
        Tools.find_wechat_path()
        pyautogui.hotkey('Tab',interval=0.5)
        pyautogui.hotkey('ctrl','v')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('esc')
    @staticmethod
    def judge_wechat_state():
        if Tools.is_wechat_running():
            window=win32gui.FindWindow('WeChatMainWndForPC','微信')
            if win32gui.IsIconic(window):
                return '主界面最小化'
            elif win32gui.IsWindowVisible(window):
                return '主界面可见'
            else:
                return '主界面不可见'
        else:
            return "微信未打开"
    @staticmethod 
    def open_wechat(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
    
        微信的打开分为三种情况:\n
        1.未登录,此时调用该函数会启动wechat_path路径下的wechat.exe\n
        启动并点击登录进入微信(注:需勾选自动登录按钮,否则启动后为扫码登录)\n
        2.已登录，主界面最小化在状态栏，此时调用该函数会直接打开后台中的微信。\n
        3.已登录，主界面关闭，此时调用该函数会打开已经关闭的微信界面。
        '''
        wechat_environ_path=[path for path in dict(os.environ).values() if 'WeChat.exe' in path]#windows环境变量中查找WeChat.exe路径
        if  wechat_environ_path:#已将WeChat.exe设置为windows环境变量
            if not Tools.is_wechat_running():#未登录微信先激活登录界面，点击进入微信，然后打开后返回主界面
                wechat=Application(backend='uia').start(wechat_environ_path[0])
                login_window=wechat.window(**Login_window.LoginWindow)
                login_button=login_window.child_window(**Login_window.LoginButton).wait(wait_for='visible',retry_interval=0.1,timeout=10)
                if login_button.is_enabled():
                    login_button.set_focus()
                    time.sleep(1)
                    login_button.click_input()
                    time.sleep(8)
                    wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                    main_window=wechat.window(**Main_window.MainWindow)
                    if is_maximize:
                        main_window.maximize()
                    return main_window#主界面
            else:
                #微信已启动，但是没有登陆，点击微信icon出现的登录界面
                try:  
                    wechat=Application(backend='uia').start(wechat_environ_path[0])
                    wechat=Application(backend='uia').connect(**Login_window.LoginWindow)
                    login_window=wechat.window(**Login_window.LoginWindow)
                    login_button=login_window.child_window(**Login_window.LoginButton).wait(wait_for='visible',retry_interval=0.1,timeout=10)
                    if login_button.is_enabled():
                        login_button.set_focus()
                        time.sleep(1)
                        login_button.click_input()
                        time.sleep(8)
                        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                        main_window=wechat.window(**Main_window.MainWindow)
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                except ElementNotFoundError:
                    try:#微信已经登录且打开过了主界面，只是挂在后台的情况，这时直接使用pywinauto的connect连接并返回主界面
                        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                        main_window=wechat.window(**Main_window.MainWindow)
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                    except ElementNotFoundError:#微信已经登陆打开过主界面但是将主界面删掉了，这时先start再connect，然后返回主界面
                        wechat=Application(backend='uia').start(wechat_environ_path[0])
                        time.sleep(2)
                        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                        main_window=wechat.window(**Main_window.MainWindow)
                        if is_maximize:
                            main_window.maximize()
                        return main_window
        else:
                if not wechat_path:
                    raise PathNotFoundError(f'未检测到微信文件路径!请输入微信文件路径或将其添加至环境变量中')
                if not Tools.is_wechat_running():#未登录微信先激活登录界面，点击进入微信，然后再找到并返回主界面
                    wechat=Application(backend='uia').start(wechat_path)
                    login_window=wechat.window(title='微信',class_name='WeChatLoginWndForPC')
                    login_button=login_window.child_window(title='进入微信',control_type="Button").wait(wait_for='visible',retry_interval=0.1,timeout=10)
                    if login_button.is_enabled():
                        login_button.set_focus()
                        time.sleep(1)
                        login_button.click_input()
                        time.sleep(8)
                        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                        main_window=wechat.window(**Main_window.MainWindow)
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                else:
                    try:  
                        wechat=Application(backend='uia').start(wechat_path)
                        wechat=Application(backend='uia').connect(title='微信',class_name='WeChatLoginWndForPC')
                        login_window=wechat.window(title='微信',class_name='WeChatLoginWndForPC')
                        login_button=login_window.child_window(title='进入微信',control_type="Button").wait(wait_for='visible',retry_interval=0.1,timeout=10)
                        if login_button.is_enabled():
                            login_button.set_focus()
                            time.sleep(1)
                            login_button.click_input()
                            time.sleep(8)
                            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                            main_window=wechat.window(**Main_window.MainWindow)
                            if is_maximize:
                                main_window.maximize()
                            return main_window
                    except ElementNotFoundError:
                        try:#微信已经登录且打开过了主界面，只是挂在后台的情况，这时直接使用connect，然后找到并返回主界面
                            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                            main_window=wechat.window(**Main_window.MainWindow)
                            if is_maximize:
                                main_window.maximize()
                            return main_window
                        except ElementNotFoundError:#微信已经登陆打开过主界面但是将主界面删掉了，这时先start再connect，然后找到并返回主界面
                            wechat=Application(backend='uia').start(wechat_path)
                            time.sleep(2)
                            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                            main_window=wechat.window(**Main_window.MainWindow)
                            if is_maximize:
                                main_window.maximize()
                            return main_window
    @staticmethod
    def open_settings(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用来打开微信设置界面。\n
        '''   
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize) 
        Toolbar=main_window.child_window(**Main_window.Toolbar)
        setting=Toolbar.children(control_type='Pane',title='')[3].children(control_type='Button',title="设置及其他")[0]
        setting.click_input()
        settings_menu=main_window.child_window(**Main_window.SettingsMenu)
        settings_button=settings_menu.child_window(control_type='Button',title="设置")
        settings_button.click_input() 
        time.sleep(2)
        desktop=Desktop(**Independent_window.Desktop)
        settings_window=desktop.window(**Independent_window.SettingWindow)
        return settings_window
    
    @staticmethod                    
    def open_dialog_window(friend:str,wechat_path:str=None,is_maximize:bool=True): 
        '''
        friend:好友或群聊备注名称,需提供完整名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开某个好友的聊天窗口
        '''
        def is_in_searh_result(friend,search_result):
            listitem=search_result.children(control_type="ListItem")
            names=[item.window_text() for item in listitem]
            if names[0]==friend:
                return True
            else:
                return False

        chat,main_window=Tools.find_friend_in_MessageList(friend=friend,wechat_path=wechat_path,is_maximize=False)
        if is_maximize:
            main_window.maximize()
        if chat:
            return chat,main_window
        else:  
            search=main_window.child_window(**Main_window.Search).wait(wait_for='visible',retry_interval=0.1,timeout=10)
            search.click_input()
            search.type_keys(friend,with_spaces=True)
            time.sleep(2)
            search_result=main_window.child_window(**Main_window.SearchResult)
            if is_in_searh_result(friend=friend,search_result=search_result):
                pyautogui.hotkey('enter')
                chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible',retry_interval=0.1,timeout=10)
                return chat,main_window #同时返回该好友的聊天窗口与主界面！若只需要其中一个需要使用元祖索引获取。
            else:
                main_window.close()
                raise NoSuchFriendError(f"好友或群聊备注有误！查无此人！")

    @staticmethod
    def find_friend_in_MessageList(friend:str,wechat_path:str=None,is_maximize:bool=False):
        '''
        friend:好友或群聊备注名称,需提供完整名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该函数用于在会话列表中查询是否存在待查询好友。\n
        若存在:返回值为 (chat,main_window) 同时返回好友聊天界面与主界面\n
        否则:返回值为 (None,main_window)只返回主界面\n
        该方法用于在会话列表中寻找好友
        '''
        def get_window(friend):
            message_list=message_list_pane.children(control_type='ListItem')
            buttons=[friend.children()[0].children()[0] for friend in message_list]
            friend_button=None
            for button in buttons:
                if friend in button.texts()[0]:
                    friend_button=button
                    #mouse.scroll(wheel_dist=-20)
                    break
            return friend_button
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(**Main_window.Toolbar)
        chat_button=Toolbar.children(control_type="Button",title='聊天')[0]
        chat_button.click_input()
        time.sleep(2)
        chat=None
        current_chat=main_window.child_window(**Main_window.CurrentChatWindow)
        if current_chat.exists() and friend in current_chat.window_text():
            chat=current_chat
            return chat,main_window
        else:
            message_list_pane=main_window.child_window(**Main_window.MessageList)
            message_list_pane.set_focus()
            rectangle=message_list_pane.rectangle()
            mouse.wheel_click(coords=(rectangle.right-5, rectangle.top+20))
            for _ in range(8):
                friend_button=get_window(friend)
                if friend_button:
                    friend_button.click_input()
                    chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible',retry_interval=0.1,timeout=10)
                    break
                else:
                    pyautogui.press("pagedown")
                    time.sleep(2)
            for _ in range(8):
                pyautogui.press("pageup")
            return chat,main_window 
    
    @staticmethod
    def open_friend_settings(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注名称,需提供完整名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开好友设置界面\n
        '''
        main_window=Tools.open_dialog_window(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)[1]
        ChatMessage=main_window.child_window(**Main_window.ChatMessage)
        ChatMessage.click_input()
        friend_settings_window=main_window.child_window(**Main_window.FriendSettingsWindow)
        return friend_settings_window,main_window
 
    @staticmethod
    def open_contacts_settings(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开通讯录设置界面
        '''
        desktop=Desktop(**Independent_window.Desktop)
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        toolbar=main_window.child_window(control_type='ToolBar',title='导航')
        contacts=toolbar.child_window(**ToolBar.Contacts)
        contacts.set_focus()
        contacts.click_input()
        contacts_settings=main_window.child_window(**Main_window.ContactsManage)#通讯录管理窗口按钮 
        contacts_settings.set_focus()
        contacts_settings.click_input()
        contacts_settings_window=desktop.window(**Independent_window.ContactManagerWindow)
        main_window.close()
        return contacts_settings_window,main_window
    
    @staticmethod
    def open_friend_settings_menu(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注名称,需提供完整名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开好友设置界面
        '''
        friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        friend_button=friend_settings_window.child_window(title=friend,control_type="Button",found_index=0)
        friend_button.click_input()
        profile_window=friend_settings_window.child_window(class_name="ContactProfileWnd",control_type="Pane",framework_id='Win32')
        more_button=profile_window.child_window(title='更多',control_type='Button')
        more_button.click_input()
        friend_menu=profile_window.child_window(control_type="Menu",title="",class_name='CMenuWnd',framework_id='Win32')
        return friend_menu,friend_settings_window,main_window
         
    @staticmethod
    def open_collections(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开收藏
        '''
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(**Main_window.Toolbar)
        collections_button=Toolbar.child_window(**ToolBar.Collections)
        collections_button.click_input()
        return main_window
    
    @staticmethod
    def open_group_settings(group:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        group:群聊备注名称,需提供完整名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        '''
        main_window=Tools.open_dialog_window(friend=group,wechat_path=wechat_path,is_maximize=False)[1]
        if is_maximize:
            main_window.maximize()
        ChatMessage=main_window.child_window(**Main_window.ChatMessage)
        ChatMessage.click_input()
        group_settings_window=main_window.child_window(**Main_window.GroupSettingsWindow)
        group_settings_window.child_window(title="群聊名称",control_type="Text").click_input()
        pyautogui.press('pagedown')
        return group_settings_window,main_window

    
    @staticmethod
    def open_wechat_moments(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开微信朋友圈
        '''
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(**Main_window.Toolbar)
        moments_button=Toolbar.child_window(**ToolBar.Moments)
        moments_button.click_input()
        desktop=Desktop(**Independent_window.Desktop)
        moments_window=desktop.window(**Independent_window.MomentsWindow)
        main_window.close()
        return moments_window
    
    @staticmethod
    def open_chat_files(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开聊天文件
        '''
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(**Main_window.Toolbar)
        moments_button=Toolbar.child_window(**ToolBar.ChatFiles)
        moments_button.click_input()
        desktop=Desktop(**Independent_window.Desktop)
        chat_files_window=desktop.window(**Independent_window.ChatFilesWindow)
        main_window.close()
        return chat_files_window
    
    @staticmethod
    def open_friend_profile(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友备注名称,需提供完整名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开好友个人简介界面
        '''
        friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        friend_button=friend_settings_window.child_window(title=friend,control_type="Button",found_index=0)
        friend_button.click_input()
        profile_window=friend_settings_window.child_window(class_name="ContactProfileWnd",control_type="Pane",framework_id='Win32')
        return profile_window,main_window
    
    @staticmethod
    def open_wechat_contacts(wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友或群聊备注名称,需提供完整名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开微信通信录界面
        '''
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        toolbar=main_window.child_window(**Main_window.Toolbar)
        contacts=toolbar.child_window(**ToolBar.Contacts)
        contacts.set_focus()
        contacts.click_input()
        return main_window
    
    @staticmethod
    def open_wechat_settings(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该函数用于打开微信设置界面
        '''
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(**Main_window.Toolbar)
        setting=Toolbar.child_window(**ToolBar.SettingsAndOthers)
        setting.click_input()
        settings_menu=main_window.child_window(**Main_window.SettingsMenu)
        settings_button=settings_menu.child_window(**Main_window.SettingsButton)
        settings_button.click_input() 
        time.sleep(2)
        desktop=Desktop(**Independent_window.Desktop)
        settings_window=desktop.window(**Independent_window.SettingWindow)
        main_window.close()
        return settings_window

class API():
    '''这个模块包括打开小程序,打开微信公众号,打开微信朋友圈,打开视频号并返回其主界面(pywinauto--window)的三个功能\n
    若有其他开发者需要使用到微信内的这三个功能进行自动化操作可调用此模块\n
    '''
    @staticmethod
    def open_wechat_mini_program(program_name:str,load_delay:int=3,wechat_path:str=None,is_maximize:bool=True):
        '''
        program_name:微信小程序名字\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        '''
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(**Main_window.Toolbar)
        program_button=Toolbar.children(control_type="Pane",title="")[1].children(control_type="Pane",title="")[0].children(control_type="Button",title='小程序面板')[0]
        program_button.click_input()
        desktop=Desktop(Desktop)
        program_window=desktop.window(**Independent_window.MiniProgramWindow)
        main_window.close()
        time.sleep(load_delay)
        search=program_window.child_window(control_type="Pane",title="",found_index=1).child_window(title="",control_type="Tab").children(title="",control_type="Pane")[1]
        search.click_input()
        search.type_keys(program_name,with_spaces=True)
        pyautogui.press("enter")
        time.sleep(load_delay)
        search_result=program_window.child_window(control_type="Document",class_name="Chrome_RenderWidgetHostHWND")
        program_button=search_result.child_window(control_type="ListItem",found_index=4)
        program_button.click_input()
        main=search_result.child_window(control_type="Button",title_re=program_name,found_index=0,framework_id="Chrome")
        time.sleep(load_delay)
        main.click_input()
        time.sleep(load_delay)
        program_window=desktop.window(title=program_name,framework_id="Win32",control_type="Pane",class_name="Chrome_WidgetWin_0")
        return program_window
    
    @staticmethod
    def open_wechat_official_account(official_acount_name:str,load_delay:int=3,wechat_path:str=None,is_maximize:bool=True):
        '''
        official_acount_name:微信公众号名称\n
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开微信公众号
        '''
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(**Main_window.Toolbar)
        program_button=Toolbar.children(control_type="Pane",title="")[1].children(control_type="Pane",title="")[0].children(control_type="Button",title='小程序面板')[0]
        program_button.click_input()
        desktop=Desktop(**Independent_window.Desktop)
        program_window=desktop.window(**Independent_window.MiniProgramWindow)
        main_window.close()
        time.sleep(load_delay)
        search=program_window.child_window(control_type="Pane",title="",found_index=1).child_window(title="",control_type="Tab").children(title="",control_type="Pane")[1]
        search.click_input()
        search.type_keys(official_acount_name,with_spaces=True)
        pyautogui.press("enter")
        time.sleep(load_delay)
        search_result=program_window.child_window(control_type="Document",class_name="Chrome_RenderWidgetHostHWND")
        official_acount_button=search_result.child_window(control_type="ListItem",found_index=3)
        official_acount_button.click_input()
        try:
            result=search_result.child_window(control_type="Button",found_index=1,framework_id="Chrome")
            time.sleep(load_delay)
            result.click_input()
            time.sleep(load_delay)
        except ElementNotFoundError:
            print("查无此公众号！")
    
    @staticmethod
    def open_search(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开搜一搜
        '''
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(**Main_window.Toolbar)
        search_button=Toolbar.child_window(**ToolBar.Search)
        search_button.click_input()
        desktop=Desktop(**Independent_window.Desktop)
        search_window=desktop.window(**Independent_window.SearchWindow)
        main_window.close()
        return search_window    

    @staticmethod
    def open_channel(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开视频号\n        
        '''
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(**Main_window.Toolbar)
        channel_button=Toolbar.child_window(**ToolBar.Channel)
        channel_button.click_input()
        desktop=Desktop(**Independent_window.Desktop)
        channel_window=desktop.window(**Independent_window.ChannelWindow)
        main_window.close()
        return channel_window

    @staticmethod
    def open_top_stories(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开看一看
        '''
        if Tools.judge_wechat_state()=='主界面可见':
            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
            main_window=wechat.window(**Main_window.MainWindow)
        else:
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(**Main_window.Toolbar)
        top_stories_button=Toolbar.child_window(**ToolBar.Topstories)
        top_stories_button.click_input()
        desktop=Desktop(backend="uia")
        top_stories_window=desktop.window(**Independent_window.TopStoriesWindow)
        main_window.close()
        return top_stories_window


def is_wechat_running():
    '''这个函数通过检测当前windows进程中\n
    是否有wechatappex该项进程来判断微信是否在运行'''
    flag=0
    for process in psutil.process_iter(['name']):
        if 'WeChatAppEx' in process.info['name']:
            flag=1
            break
    if flag:#flag不为0，表明微信在运行，即微信已登录
         return True
    else:#flag为0,表明微信不在运行，即微信未登录
         return False


def open_wechat(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
        若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        该方法用于打开微信公众号
        微信的打开分为三种情况:\n
        1.未登录,此时调用该函数会启动wechat_path路径下的wechat.exe\n
        启动并点击登录进入微信(注:需勾选自动登录按钮,否则启动后为扫码登录)\n
        2.已登录，主界面最小化在状态栏，此时调用该函数会直接打开后台中的微信。\n
        3.已登录，主界面关闭，此时调用该函数会打开已经关闭的微信界面。
        '''
        wechat_environ_path=[path for path in dict(os.environ).values() if 'WeChat.exe' in path]#windows环境变量中查找WeChat.exe路径
        if  wechat_environ_path:#已将WeChat.exe设置为windows环境变量
            if not Tools.is_wechat_running():#未登录微信先激活登录界面，点击进入微信，然后打开后返回主界面
                wechat=Application(backend='uia').start(wechat_environ_path[0])
                login_window=wechat.window(**Login_window.LoginWindow)
                login_button=login_window.child_window(**Login_window.LoginButton).wait(wait_for='visible',retry_interval=0.1,timeout=10)
                if login_button.is_enabled():
                    login_button.set_focus()
                    time.sleep(1)
                    login_button.click_input()
                    time.sleep(8)
                    wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                    main_window=wechat.window(**Main_window.MainWindow)
                    if is_maximize:
                        main_window.maximize()
                    return main_window#主界面
            else:
                #微信已启动，但是没有登陆，点击微信icon出现的事登录界面
                try:  
                    wechat=Application(backend='uia').start(wechat_environ_path[0])
                    wechat=Application(backend='uia').connect(**Login_window.LoginWindow)
                    login_window=wechat.window(**Login_window.LoginWindow)
                    login_button=login_window.child_window(**Login_window.LoginButton).wait(wait_for='visible',retry_interval=0.1,timeout=10)
                    if login_button.is_enabled():
                        login_button.set_focus()
                        time.sleep(1)
                        login_button.click_input()
                        time.sleep(8)
                        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                        main_window=wechat.window(**Main_window.MainWindow)
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                except ElementNotFoundError:
                    try:#微信已经登录且打开过了主界面，只是挂在后台的情况，这时直接使用pywinauto的connect连接并返回主界面
                        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                        main_window=wechat.window(**Main_window.MainWindow)
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                    except ElementNotFoundError:#微信已经登陆打开过主界面但是将主界面删掉了，这时先start再connect，然后返回主界面
                        wechat=Application(backend='uia').start(wechat_environ_path[0])
                        time.sleep(2)
                        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                        main_window=wechat.window(**Main_window.MainWindow)
                        if is_maximize:
                            main_window.maximize()
                        return main_window
        else:
                if not wechat_path:
                    raise PathNotFoundError(f'未检测到微信文件路径!请输入微信文件路径或将其添加至环境变量中')
                if not Tools.is_wechat_running():#未登录微信先激活登录界面，点击进入微信，然后再找到并返回主界面
                    wechat=Application(backend='uia').start(wechat_path)
                    login_window=wechat.window(**Login_window.LoginWindow)
                    login_button=login_window.child_window(**Login_window.LoginButton).wait(wait_for='visible',retry_interval=0.1,timeout=10)
                    if login_button.is_enabled():
                        login_button.set_focus()
                        time.sleep(1)
                        login_button.click_input()
                        time.sleep(8)
                        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                        main_window=wechat.window(**Main_window.MainWindow)
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                else:
                    try:  
                        wechat=Application(backend='uia').start(wechat_path)
                        wechat=Application(backend='uia').connect(**Login_window.LoginWindow)
                        login_window=wechat.window(**Login_window.LoginWindow)
                        login_button=login_window.child_window(**Login_window.LoginButton).wait(wait_for='visible',retry_interval=0.1,timeout=10)
                        if login_button.is_enabled():
                            login_button.set_focus()
                            time.sleep(1)
                            login_button.click_input()
                            time.sleep(8)
                            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                            main_window=wechat.window(**Main_window.MainWindow)
                            if is_maximize:
                                main_window.maximize()
                            return main_window
                    except ElementNotFoundError:
                        try:#微信已经登录且打开过了主界面，只是挂在后台的情况，这时直接使用connect，然后找到并返回主界面
                            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                            main_window=wechat.window(**Main_window.MainWindow)
                            if is_maximize:
                                main_window.maximize()
                            return main_window
                        except ElementNotFoundError:#微信已经登陆打开过主界面但是将主界面删掉了，这时先start再connect，然后找到并返回主界面
                            wechat=Application(backend='uia').start(wechat_path)
                            time.sleep(2)
                            wechat=Application(backend='uia').connect(**Main_window.MainWindow)
                            main_window=wechat.window(**Main_window.MainWindow)
                            if is_maximize:
                                main_window.maximize()
                            return main_window





            

def find_friend_in_MessageList(friend:str,wechat_path:str=None,is_maximize:bool=False):
    '''
    friend:好友或群聊备注名称,需提供完整名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于在会话列表中查询是否存在待查询好友。\n
    若存在:返回值为 (chat,main_window) 同时返回好友聊天界面与主界面\n
    否则:返回值为 (None,main_window)只返回主界面\n
    该方法用于在会话列表中寻找好友
    '''
    def get_window(friend):
        message_list=message_list_pane.children(control_type='ListItem')
        buttons=[friend.children()[0].children()[0] for friend in message_list]
        friend_button=None
        for button in buttons:
            if friend in button.texts()[0]:
                friend_button=button
                break
        return friend_button
    if Tools.judge_wechat_state()=='主界面可见':
        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
        main_window=wechat.window(**Main_window.MainWindow)
    else:
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(**Main_window.Toolbar)
    chat_button=Toolbar.children(control_type="Button",title='聊天')[0]
    chat_button.click_input()
    time.sleep(2)
    chat=None
    current_chat=main_window.child_window(**Main_window.CurrentChatWindow)
    if current_chat.exists() and friend in current_chat.window_text():
        chat=current_chat
        return chat,main_window
    else:    
        message_list_pane=main_window.child_window(**Main_window.MessageList)
        message_list_pane.set_focus()
        rectangle=message_list_pane.rectangle()
        mouse.wheel_click(coords=(rectangle.right-5, rectangle.top+20))
        for _ in range(5):
            friend_button=get_window(friend)
            if friend_button:
                friend_button.click_input()
                chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible',retry_interval=0.1,timeout=10)
                break
            pyautogui.press("pagedown")
            time.sleep(2)
        for _ in range(5):
            pyautogui.press("pageup")
        return chat,main_window


def open_dialog_window(friend:str,wechat_path:str=None,is_maximize:bool=True): 
    '''
    friend:好友或群聊备注名称,需提供完整名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该方法用于打开某个好友的聊天窗口
    '''
    def is_in_searh_result(friend,search_result):
        listitem=search_result.children(control_type="ListItem")
        names=[item.window_text() for item in listitem]
        if names[0]==friend:
            return True
        else:
            return False

    chat,main_window=Tools.find_friend_in_MessageList(friend=friend,wechat_path=wechat_path,is_maximize=False)
    if is_maximize:
        main_window.maximize()
    if chat:
        return chat,main_window
    else:
        search=main_window.child_window(**Main_window.Search).wait(wait_for='visible',retry_interval=0.1,timeout=10)
        search.click_input()
        search.type_keys(friend,with_spaces=True)
        time.sleep(2)
        search_result=main_window.child_window(**Main_window.SearchResult)
        if is_in_searh_result(friend=friend,search_result=search_result):
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible',retry_interval=0.1,timeout=10)
            return chat,main_window #同时返回该好友的聊天窗口与主界面！若只需要其中一个需要使用元祖索引获取。
        else:
            main_window.close()
            raise NoSuchFriendError(f"好友或群聊备注有误！查无此人！")


def open_wechat_settings(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开微信设置界面
    '''
    if Tools.judge_wechat_state()=='主界面可见':
        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
        main_window=wechat.window(**Main_window.MainWindow)
    else:
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(**Main_window.Toolbar)
    setting=Toolbar.children(control_type='Pane',title='')[3].children(control_type='Button',title="设置及其他")[0]
    setting.click_input()
    settings_menu=main_window.child_window(**Main_window.SettingsMenu)
    settings_button=settings_menu.child_window(**Main_window.SettingsButton)
    settings_button.click_input() 
    time.sleep(2)
    desktop=Desktop(**Independent_window.Desktop)
    setting_window=desktop.window(**Independent_window.SettingWindow)
    main_window.close()
    return setting_window


def open_wechat_moments(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开微信朋友圈
    '''
    if Tools.judge_wechat_state()=='主界面可见':
        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
        main_window=wechat.window(**Main_window.MainWindow)
    else:
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(**Main_window.Toolbar)
    moments_button=Toolbar.children(control_type="Button",title='朋友圈')[0]
    moments_button.click_input()
    desktop=Desktop(**Independent_window.Desktop)
    moments_window=desktop.window(**Independent_window.MomentsWindow)
    main_window.close()
    return moments_window


        
def open_wechat_mini_program(program_name:str,load_delay:int=3,wechat_path:str=None,is_maximize:bool=True):
    '''
    program_name:小程序名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开微信小程序
    '''
    if Tools.judge_wechat_state()=='主界面可见':
        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
        main_window=wechat.window(**Main_window.MainWindow)
    else:
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(**Main_window.Toolbar)
    program_button=Toolbar.child_window(**ToolBar.Miniprogram_pane)
    program_button.click_input()
    desktop=Desktop(**Independent_window.Desktop)
    program_window=desktop.window(**Independent_window.MiniProgramWindow)
    main_window.close()
    time.sleep(load_delay)
    search=program_window.child_window(control_type="Pane",title="",found_index=1).child_window(title="",control_type="Tab").children(title="",control_type="Pane")[1]
    search.click_input()
    search.type_keys(program_name,with_spaces=True)
    pyautogui.press("enter")
    time.sleep(load_delay)
    search_result=program_window.child_window(control_type="Document",class_name="Chrome_RenderWidgetHostHWND")
    program_button=search_result.child_window(control_type="ListItem",found_index=4)
    program_button.click_input()
    main=search_result.child_window(control_type="Button",title_re=program_name,found_index=0,framework_id="Chrome")
    time.sleep(load_delay)
    main.click_input()
    time.sleep(load_delay)
    program_window=desktop.window(title=program_name,framework_id="Win32",control_type="Pane",class_name="Chrome_WidgetWin_0")
    return program_window

def open_wechat_official_account(official_acount_name:str,load_delay:int=3,wechat_path:str=None,is_maximize:bool=True):
    '''
    official_account__name:公众号名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开公众号
    '''
    if Tools.judge_wechat_state()=='主界面可见':
        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
        main_window=wechat.window(**Main_window.MainWindow)
    else:
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(**Main_window.Toolbar)
    program_button=Toolbar.child_window(**ToolBar.Miniprogram_pane)
    program_button.click_input()
    desktop=Desktop(**Independent_window.Desktop)
    program_window=desktop.window(**Independent_window.MiniProgramWindow)
    main_window.close()
    time.sleep(load_delay)
    search=program_window.child_window(control_type="Pane",title="",found_index=1).child_window(title="",control_type="Tab").children(title="",control_type="Pane")[1]
    search.click_input()
    search.type_keys(official_acount_name,with_spaces=True)
    pyautogui.press("enter")
    time.sleep(load_delay)
    search_result=program_window.child_window(control_type="Document",class_name="Chrome_RenderWidgetHostHWND")
    official_acount_button=search_result.child_window(control_type="ListItem",found_index=3)
    official_acount_button.click_input()
    try:
        result=search_result.child_window(control_type="Button",found_index=1,framework_id="Chrome")
        time.sleep(2)
        result.click_input()
        time.sleep(2)
    except ElementNotFoundError:
        print("查无此公众号！")

def open_contacts_settings(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开通讯录管理
    '''
    desktop=Desktop(**Independent_window.Desktop)
    main_window=Tools.open_wechat(wechat_path,is_maximize=is_maximize)
    toolbar=main_window.child_window(**Main_window.Toolbar)
    contacts=toolbar.child_window(**ToolBar.Contacts)
    contacts.set_focus()
    contacts.click_input()
    time.sleep(2)
    contacts_settings=main_window.child_window(**Main_window.ContactsManage)#通讯录管理按钮
    contacts_settings.set_focus()
    contacts_settings.click_input()
    contacts_settings_window=desktop.window(**Independent_window.ContactManagerWindow)
    main_window.close()
    return contacts_settings_window

def open_wechat_contacts(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开通讯录列表\n
    '''
    if Tools.judge_wechat_state()=='主界面可见':
        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
        main_window=wechat.window(**Main_window.MainWindow)
    else:
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    toolbar=main_window.child_window(**Main_window.Toolbar)
    contacts=toolbar.child_window(**ToolBar.Contacts)
    contacts.set_focus()
    contacts.click_input()
    return main_window

def open_friend_settings(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开好友的设置界面
    '''
    main_window=Tools.open_dialog_window(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)[1]
    ChatMessage=main_window.child_window(title="聊天信息",control_type="Button")
    ChatMessage.click_input()
    friend_settings_window=main_window.child_window(**Main_window.FriendSettingsWindow)
    return friend_settings_window,main_window

def open_friend_settings_menu(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开好友的设置菜单
    '''
    friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    friend_button=friend_settings_window.child_window(title=friend,control_type="Button",found_index=0)
    friend_button.click_input()
    profile_window=friend_settings_window.child_window(class_name="ContactProfileWnd",control_type="Pane",framework_id='Win32')
    more_button=profile_window.child_window(title='更多',control_type='Button')
    more_button.click_input()
    friend_menu=profile_window.child_window(control_type="Menu",title="",class_name='CMenuWnd',framework_id='Win32')
    return friend_menu,main_window

def open_channel(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开视频号
    '''
    if Tools.judge_wechat_state()=='主界面可见':
        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
        main_window=wechat.window(**Main_window.MainWindow)
    else:
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(**Main_window.Toolbar)
    channel_button=Toolbar.child_window(**ToolBar.Channel)
    channel_button.click_input()
    desktop=Desktop(**Independent_window.Desktop)
    channel_window=desktop.window(**Independent_window.ChannelWindow)
    return channel_window

def open_top_stories(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开看一看
    '''
    if Tools.judge_wechat_state()=='主界面可见':
        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
        main_window=wechat.window(**Main_window.MainWindow)
    else:
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(**Main_window.Toolbar)
    top_stories_button=Toolbar.child_window(**ToolBar.Topstories)
    top_stories_button.click_input()
    desktop=Desktop(**Independent_window.Desktop)
    top_stories_window=desktop.window(**Independent_window.TopStoriesWindow)
    return top_stories_window

def open_collections(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开收藏
    '''
    if Tools.judge_wechat_state()=='主界面可见':
        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
        main_window=wechat.window(**Main_window.MainWindow)
    else:
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(**Main_window.Toolbar)
    collections_button=Toolbar.child_window(**ToolBar.Collections)
    collections_button.click_input()
    return main_window


def open_group_settings(group:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    group:群聊备注名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开好友的设置界面
    '''
    main_window=Tools.open_dialog_window(friend=group,wechat_path=wechat_path,is_maximize=is_maximize)[1]
    ChatMessage=main_window.child_window(**Main_window.ChatMessage)
    ChatMessage.click_input()
    group_settings_window=main_window.child_window(**Main_window.GroupSettingsWindow)
    return group_settings_window,main_window


    
def open_chat_files(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开聊天界面
    '''
    if Tools.judge_wechat_state()=='主界面可见':
        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
        main_window=wechat.window(**Main_window.MainWindow)
    else:
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(**Main_window.Toolbar)
    moments_button=Toolbar.child_window(**ToolBar.ChatFiles)
    moments_button.click_input()
    desktop=Desktop(**Independent_window.Desktop)
    filelist_window=desktop.window(**Independent_window.ChatFilesWindow)
    main_window.close()
    return filelist_window
    

def open_friend_profile(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''
    friend:好友备注名称\n
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开好友的个人简介界面
    '''
    friend_settings_window,main_window=Tools.open_friend_settings(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    friend_button=friend_settings_window.child_window(title=friend,control_type="Button",found_index=0)
    friend_button.click_input()
    profile_window=friend_settings_window.child_window(class_name="ContactProfileWnd",control_type="Pane",framework_id='Win32')
    return profile_window,main_window

    

def open_search(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    这里强烈建议将微信路径加入到windows环境变量中,因为该函数默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用该函数无需任何参数\n
    若你没有设置微信的Wechat.exe地址为环境变量,那么你需要传入该文件地址作为参数否则会引发错误无法打开微信！\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    该函数用于打开搜一搜
    '''
    if Tools.judge_wechat_state()=='主界面可见':
        wechat=Application(backend='uia').connect(**Main_window.MainWindow)
        main_window=wechat.window(**Main_window.MainWindow)
    else:
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(**Main_window.Toolbar)
    search_button=Toolbar.child_window(**ToolBar.Search)
    search_button.click_input()
    desktop=Desktop(**Independent_window.Desktop)
    search_window=desktop.window(**Independent_window.SearchWindow)
    main_window.close()
    return search_window
 
def find_wechat_path():
    wechat_path=None
    for process in psutil.process_iter(['name','exe']):
        if 'WeChat.exe' in process.info['name']:
            wechat_path=process.info['exe']
            Systemsettings.copy_text_to_windowsclipboard(wechat_path)
            break
    if wechat_path:
        print("已将微信路径复制到剪贴板")
        return wechat_path
    else:
        pyautogui.press('esc')
        pyautogui.hotkey('alt','f4')
        pyautogui.hotkey('alt','f4')
        raise WechatNotStartError(f'微信尚未启动,无法寻找到路径,请先启动后再试!')
def set_wechat_as_environ_path():
    subprocess.Popen(["SystemPropertiesAdvanced.exe"])
    pyautogui.hotkey('alt','n',interval=0.5)
    pyautogui.hotkey('alt','n',interval=0.5)
    pyautogui.press('shift')
    pyautogui.typewrite('wechatpath')
    Tools.find_wechat_path()
    pyautogui.hotkey('Tab',interval=0.5)
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('esc')

