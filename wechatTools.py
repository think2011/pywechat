'''使用该模块的三个函数或方法时候,你可以:\n
 from pywechat.wechatTools import Tools \n
 导入Tools,使用Tools下的四个方法:\n
 Tools.open_wechat();Tools.open_dialog_window();Tools.is_wechat_running(),Tools.find_friend_in_MessageList\n
 或者:\n
 from pywechat.wechatTools import open_wechat,open_dialog_window,is_wechat_running,find_friend_in_MessageList\n
 直接导入该四个函数:\n
 open_wechat();open_dialog_window();is_wechat_running();find_friend_in_MessageList\n
 '''
############################依赖环境
import os
import time
import psutil
import pyautogui
from pywinauto import mouse
from pywinauto import Desktop
from pywechat.Errors import PathNotFoundError
from pywechat.Errors import HavebeenPinnedError
from pywechat.Errors import NoSuchfriendError
from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError
import pandas as pd
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.width', 280)
pd.set_option("display.max_rows",10000)
pd.set_option("display.max_columns",50)
#############################################################
pyautogui.FAILSAFE = False#防止鼠标在屏幕边缘处造成的错误
class Tools():
    '''该模块中封装了3个关于PC微信的工具,包括:\n
    is_wechat_running:用来判断PC微信是否运行。\n
    open_wechat:打开PC微信。\n
    find_friends_in_Messagelist:在会话列表和当前聊天窗口中查找好友\n
    open_dialog_window:打开与某个好友的对话框窗口。
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
                login_window=wechat.window(title='微信',class_name='WeChatLoginWndForPC')
                login_button=login_window.child_window(title='进入微信',control_type="Button").wait(wait_for='visible')
                if login_button.is_enabled():
                    login_button.set_focus()
                    time.sleep(1)
                    login_button.click_input()
                    time.sleep(8)
                    wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                    main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                    if is_maximize:
                        main_window.maximize()
                    return main_window#主界面
            else:
                #微信已启动，但是没有登陆，点击微信icon出现的事登录界面
                try:  
                    wechat=Application(backend='uia').start(wechat_environ_path[0])
                    wechat=Application(backend='uia').connect(title='微信',class_name='WeChatLoginWndForPC')
                    login_window=wechat.window(title='微信',class_name='WeChatLoginWndForPC')
                    login_button=login_window.child_window(title='进入微信',control_type="Button").wait(wait_for='visible')
                    if login_button.is_enabled():
                        login_button.set_focus()
                        time.sleep(1)
                        login_button.click_input()
                        time.sleep(8)
                        wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                        main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                except ElementNotFoundError:
                    try:#微信已经登录且打开过了主界面，只是挂在后台的情况，这时直接使用pywinauto的connect连接并返回主界面
                        wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                        main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                    except ElementNotFoundError:#微信已经登陆打开过主界面但是将主界面删掉了，这时先start再connect，然后返回主界面
                        wechat=Application(backend='uia').start(wechat_environ_path[0])
                        time.sleep(2)
                        wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                        main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                        if is_maximize:
                            main_window.maximize()
                        return main_window
        else:
                if not wechat_path:
                    raise PathNotFoundError(f'未检测到微信文件路径!请输入微信文件路径或将其添加至环境变量中')
                if not Tools.is_wechat_running():#未登录微信先激活登录界面，点击进入微信，然后再找到并返回主界面
                    wechat=Application(backend='uia').start(wechat_path)
                    login_window=wechat.window(title='微信',class_name='WeChatLoginWndForPC')
                    login_button=login_window.child_window(title='进入微信',control_type="Button").wait(wait_for='visible')
                    if login_button.is_enabled():
                        login_button.set_focus()
                        time.sleep(1)
                        login_button.click_input()
                        time.sleep(8)
                        wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                        main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                else:
                    try:  
                        wechat=Application(backend='uia').start(wechat_path)
                        wechat=Application(backend='uia').connect(title='微信',class_name='WeChatLoginWndForPC')
                        login_window=wechat.window(title='微信',class_name='WeChatLoginWndForPC')
                        login_button=login_window.child_window(title='进入微信',control_type="Button").wait(wait_for='visible')
                        if login_button.is_enabled():
                            login_button.set_focus()
                            time.sleep(1)
                            login_button.click_input()
                            time.sleep(8)
                            wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                            main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                            if is_maximize:
                                main_window.maximize()
                            return main_window
                    except ElementNotFoundError:
                        try:#微信已经登录且打开过了主界面，只是挂在后台的情况，这时直接使用connect，然后找到并返回主界面
                            wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                            main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                            if is_maximize:
                                main_window.maximize()
                            return main_window
                        except ElementNotFoundError:#微信已经登陆打开过主界面但是将主界面删掉了，这时先start再connect，然后找到并返回主界面
                            wechat=Application(backend='uia').start(wechat_path)
                            time.sleep(2)
                            wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                            main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                            if is_maximize:
                                main_window.maximize()
                            return main_window
    @staticmethod                    
    def open_dialog_window(friend:str,wechat_path:str=None,is_maximize:bool=True): 
        '''
        friend:好友或群聊备注名称,需提供完整名称\n
        wechat_path:微信WeChat.exe的文件地址,若已添加到windows环境变量中该参数默认为None,不需要传入该参数
        '''
        chat,main_window=find_friend_in_MessageList(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
        if chat:
            return chat,main_window
        else:
            search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
            search.click_input()
            search.type_keys(friend)
            time.sleep(1)
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            return chat,main_window  

    @staticmethod
    def find_friend_in_MessageList(friend:str,wechat_path:str=None,is_maximize:bool=True):
        '''
        friend:好友或群聊备注
        is_maximize:微信界面是否全屏,默认全屏。\n
        该函数用于在会话列表中查询是否存在待查询好友。\n
        若存在:返回值为 (chat,main_window) 同时返回好友聊天界面与主界面\n
        否则:返回值为 (None,main_window)只返回主界面'''
        def get_window(friend):
            message_list=message_list_pane.children(control_type='ListItem')
            buttons=[friend.children()[0].children()[0] for friend in message_list]
            friend_button=None
            for button in buttons:
                if friend==button.texts()[0]:
                    friend_button=button
                    break
            return friend_button
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        current_chat=main_window.child_window(control_type='Edit',found_index=1)
        if current_chat.exists() and current_chat.texts()[0]==friend:
            return current_chat,main_window
        else:
            message_list_pane=main_window.child_window(title="会话",control_type="List")
            message_list_pane.set_focus()
            rectangle=message_list_pane.rectangle()
            mouse.wheel_click(coords=(rectangle.right-5, rectangle.top+20))
            chat=None
            for _ in range(5):
                friend_button=get_window(friend)
                if friend_button:
                    friend_button.click_input()
                    chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
                    break
                pyautogui.press("pagedown")
                time.sleep(2)
            for _ in range(5):
                pyautogui.press("pageup")
            return chat,main_window 
         
class wechatSettings():
    @staticmethod
    def open_settings(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
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
    @staticmethod
    def log_out(wechat_path:str=None,is_maximize:bool=True):
        '''
        wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        PC微信退出登录。\n
        '''
        settings_window=open_wechat_settings(wechat_path=wechat_path,is_maximize=is_maximize)
        log_out_button=settings_window.window(title="退出登录",control_type="Button")
        log_out_button.click_input()
        time.sleep(2)
        confirm_button=settings_window.window(title="确定",control_type="Button")
        confirm_button.click_input()



class API():
    '''这个模块包括打开小程序,打开微信公众号,打开微信朋友圈,打开视频号个人主页'''
    @staticmethod
    def open_wechat_program(program_name,wechat_path:str=None,is_maximize:bool=True):
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(control_type="ToolBar",title='导航')
        program_button=Toolbar.children(control_type="Pane",title="")[1].children(control_type="Pane",title="")[0].children(control_type="Button",title='小程序面板')[0]
        program_button.click_input()
        desktop=Desktop(backend="uia")
        program_window=desktop.window(control_type="Pane",class_name="Chrome_WidgetWin_0",title="微信")
        main_window.close()
        search=program_window.child_window(control_type="Pane",title="",found_index=1).child_window(title="",control_type="Tab").children(title="",control_type="Pane")[1]
        search.click_input()
        search.type_keys(program_name)
        pyautogui.press("enter")
        search_result=program_window.child_window(control_type="Document",class_name="Chrome_RenderWidgetHostHWND")
        program_button=search_result.child_window(control_type="ListItem",found_index=4)
        program_button.click_input()
        main=search_result.child_window(control_type="Button",title_re=program_name,found_index=0,framework_id="Chrome")
        time.sleep(2)
        main.click_input()
        time.sleep(2)
        program_window=desktop.window(title=program_name,framework_id="Win32",control_type="Pane",class_name="Chrome_WidgetWin_0")
        return program_window
    
    @staticmethod
    def open_wechat_official_account(official_acount_name,wechat_path:str=None,is_maximize:bool=True):
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(control_type="ToolBar",title='导航')
        program_button=Toolbar.children(control_type="Pane",title="")[1].children(control_type="Pane",title="")[0].children(control_type="Button",title='小程序面板')[0]
        program_button.click_input()
        desktop=Desktop(backend="uia")
        program_window=desktop.window(control_type="Pane",class_name="Chrome_WidgetWin_0",title="微信")
        main_window.close()
        search=program_window.child_window(control_type="Pane",title="",found_index=1).child_window(title="",control_type="Tab").children(title="",control_type="Pane")[1]
        search.click_input()
        search.type_keys(official_acount_name)
        pyautogui.press("enter")
        search_result=program_window.child_window(control_type="Document",class_name="Chrome_RenderWidgetHostHWND")
        official_acount_button=search_result.child_window(control_type="ListItem",found_index=3)
        official_acount_button.click_input()
        result=search_result.child_window(control_type="Button",title_re=official_acount_name,found_index=0,framework_id="Chrome")
        time.sleep(2)
        result.click_input()
    
    @staticmethod
    def open_wechat_moments(wechat_path:str=None,is_maximize:bool=True):
        main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
        Toolbar=main_window.child_window(control_type="ToolBar",title='导航')
        moments_button=Toolbar.children(control_type="Button",title='朋友圈')[0]
        moments_button.click_input()
        desktop=Desktop(backend="uia")
        moments_window=desktop.window(control_type="Window",class_name="SnsWnd",title="朋友圈")
        main_window.close()
        return moments_window
    
    
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
                login_window=wechat.window(title='微信',class_name='WeChatLoginWndForPC')
                login_button=login_window.child_window(title='进入微信',control_type="Button").wait(wait_for='visible')
                if login_button.is_enabled():
                    login_button.set_focus()
                    time.sleep(1)
                    login_button.click_input()
                    time.sleep(8)
                    wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                    main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                    if is_maximize:
                        main_window.maximize()
                    return main_window#主界面
            else:
                #微信已启动，但是没有登陆，点击微信icon出现的事登录界面
                try:  
                    wechat=Application(backend='uia').start(wechat_environ_path[0])
                    wechat=Application(backend='uia').connect(title='微信',class_name='WeChatLoginWndForPC')
                    login_window=wechat.window(title='微信',class_name='WeChatLoginWndForPC')
                    login_button=login_window.child_window(title='进入微信',control_type="Button").wait(wait_for='visible')
                    if login_button.is_enabled():
                        login_button.set_focus()
                        time.sleep(1)
                        login_button.click_input()
                        time.sleep(8)
                        wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                        main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                except ElementNotFoundError:
                    try:#微信已经登录且打开过了主界面，只是挂在后台的情况，这时直接使用pywinauto的connect连接并返回主界面
                        wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                        main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                    except ElementNotFoundError:#微信已经登陆打开过主界面但是将主界面删掉了，这时先start再connect，然后返回主界面
                        wechat=Application(backend='uia').start(wechat_environ_path[0])
                        time.sleep(2)
                        wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                        main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                        if is_maximize:
                            main_window.maximize()
                        return main_window
        else:
                if not wechat_path:
                    raise PathNotFoundError(f'未检测到微信文件路径!请输入微信文件路径或将其添加至环境变量中')
                if not Tools.is_wechat_running():#未登录微信先激活登录界面，点击进入微信，然后再找到并返回主界面
                    wechat=Application(backend='uia').start(wechat_path)
                    login_window=wechat.window(title='微信',class_name='WeChatLoginWndForPC')
                    login_button=login_window.child_window(title='进入微信',control_type="Button").wait(wait_for='visible')
                    if login_button.is_enabled():
                        login_button.set_focus()
                        time.sleep(1)
                        login_button.click_input()
                        time.sleep(8)
                        wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                        main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                        if is_maximize:
                            main_window.maximize()
                        return main_window
                else:
                    try:  
                        wechat=Application(backend='uia').start(wechat_path)
                        wechat=Application(backend='uia').connect(title='微信',class_name='WeChatLoginWndForPC')
                        login_window=wechat.window(title='微信',class_name='WeChatLoginWndForPC')
                        login_button=login_window.child_window(title='进入微信',control_type="Button").wait(wait_for='visible')
                        if login_button.is_enabled():
                            login_button.set_focus()
                            time.sleep(1)
                            login_button.click_input()
                            time.sleep(8)
                            wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                            main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                            if is_maximize:
                                main_window.maximize()
                            return main_window
                    except ElementNotFoundError:
                        try:#微信已经登录且打开过了主界面，只是挂在后台的情况，这时直接使用connect，然后找到并返回主界面
                            wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                            main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                            if is_maximize:
                                main_window.maximize()
                            return main_window
                        except ElementNotFoundError:#微信已经登陆打开过主界面但是将主界面删掉了，这时先start再connect，然后找到并返回主界面
                            wechat=Application(backend='uia').start(wechat_path)
                            time.sleep(2)
                            wechat=Application(backend='uia').connect(title='微信',class_name='WeChatMainWndForPC')
                            main_window=wechat.window(title='微信',class_name='WeChatMainWndForPC')
                            if is_maximize:
                                main_window.maximize()
                            return main_window




def open_dialog_window(friend:str,wechat_path:str=None,is_maximize:bool=True): 
    def is_in_searh_result(friend,search_result):
        listitem=search_result.children(control_type="ListItem")
        names=[item.texts()[0] for item in listitem]
        if friend in names:
            return True
        else:
            return False
    '''
    friend:好友或群聊备注名称,需提供完整名称\n
    wechat_path:微信WeChat.exe的文件地址,若已添加到windows环境变量中该参数默认为None,不需要传入该参数\n
    is_maximize:是否放大微信主界面
    '''
    chat,main_window=find_friend_in_MessageList(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize)
    if chat:
        return chat,main_window
    else:
        search=main_window.child_window(title='搜索',control_type='Edit').wait(wait_for='visible')
        search.click_input()
        search.type_keys(friend)
        time.sleep(2)
        search_result=main_window.child_window(title="@str:IDS_FAV_SEARCH_RESULT:3780",control_type='List')
        if is_in_searh_result(friend=friend,search_result=search_result):
            pyautogui.hotkey('enter')
            chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
            return chat,main_window #同时返回该好友的聊天窗口与主界面！若只需要其中一个需要使用元祖索引获取。
        else:
            main_window.close()
            raise NoSuchfriendError(f"好友或群聊备注有误！查无此人！")
            

def find_friend_in_MessageList(friend:str,wechat_path:str=None,is_maximize:bool=True):
    '''该函数用于在会话列表中查询是否存在待查询好友。\n
    若存在:返回值为 (chat,main_window) 同时返回好友聊天界面与主界面\n
    否则:返回值为 (None,main_window)只返回主界面'''
    def get_window(friend):
        message_list=message_list_pane.children(control_type='ListItem')
        buttons=[friend.children()[0].children()[0] for friend in message_list]
        friend_button=None
        for button in buttons:
            if friend==button.texts()[0]:
                friend_button=button
                break
        return friend_button
    main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    current_chat=main_window.child_window(control_type='Edit',found_index=1)
    if current_chat.exists() and current_chat.texts()[0]==friend:
        return current_chat
    else:
        message_list_pane=main_window.child_window(title="会话",control_type="List")
        message_list_pane.set_focus()
        rectangle=message_list_pane.rectangle()
        mouse.wheel_click(coords=(rectangle.right-5, rectangle.top+20))
        chat=None
        for _ in range(5):
            friend_button=get_window(friend)
            if friend_button:
                friend_button.click_input()
                chat=main_window.child_window(title=friend,control_type='Edit').wait(wait_for='visible')
                break
            pyautogui.press("pagedown")
            time.sleep(2)
        for _ in range(5):
            pyautogui.press("pageup")
        return chat,main_window
    
def Pin_friend(friend,wechat_path=None,is_maximize:bool=True):
    '''
    friend:好友或群聊名称。\n
    wechat_path:微信WeChat.exe的文件地址,若已添加到windows环境变量中该参数默认为None,不需要传入该参数。\n
    is_maximize:微信界面是否全屏,默认全屏。
    将好友或群聊置顶
    '''
    main_window=Tools.open_dialog_window(friend,wechat_path,is_maximize=is_maximize)[1] 
    Tool_bar=main_window.child_window(found_index=1,title='',control_type='ToolBar')
    try:
        Top_button=Tool_bar.children(title='置顶',control_type='Button')[0]
        Top_button[0].click_input()
    except IndexError:
        raise HavebeenPinnedError(f"好友{friend}已被置顶,无需操作！")


def open_wechat_settings(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
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
    main_window.close()
    return settings_window

def log_out(wechat_path:str=None,is_maximize:bool=True):
    '''
    wechat_path:微信的WeChat.exe文件地址,当微信未登录时,该函数使用此地址启动微信并点击登录完成打开微信界面。\n
    is_maximize:微信界面是否全屏,默认全屏。\n
    PC微信退出登录。\n
    '''
    settings_window=open_wechat_settings(wechat_path=wechat_path,is_maximize=is_maximize)
    log_out_button=settings_window.window(title="退出登录",control_type="Button")
    log_out_button.click_input()
    time.sleep(2)
    confirm_button=settings_window.window(title="确定",control_type="Button")
    confirm_button.click_input()


def open_wechat_moments(wechat_path:str=None,is_maximize:bool=True):
    main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(control_type="ToolBar",title='导航')
    moments_button=Toolbar.children(control_type="Button",title='朋友圈')[0]
    moments_button.click_input()
    desktop=Desktop(backend="uia")
    moments_window=desktop.window(control_type="Window",class_name="SnsWnd",title="朋友圈")
    main_window.close()
    return moments_window

# def wechat_moment_likes(wechat_path:str=None,is_maximize:bool=True):
#     def judge_picture(content):
#         if "图片" in content.window_text():
#             return "图片"#判断是否有图片视频文本
#         elif "视频" in content.window_text():
#             return "视频"
#         return "文字"
#     moments_window=open_wechat_moments(wechat_path=wechat_path,is_maximize=is_maximize)
#     refresh_button=moments_window.child_window(title="刷新",control_type="Button")
#     refresh_button.click_input()
#     rec=moments_window.rectangle()
#     mouse.click(coords=(rec.right-10,rec.top+100))
#     pyautogui.press("pagedown")
#     moments_list=moments_window.child_window(control_type="List",title="朋友圈")
#     for _ in range(10):
#         time.sleep(5)
#         contents=moments_list.children(control_type="ListItem")
#         if judge_picture(contents[0])=="图片" or "视频":
#                     button=contents[0].children()[0].children()[0].children()[1].children()[3].children(title="评论",control_type="Button")[0]
#                     button.click_input()
#                     time.sleep(2)
#                     try:
#                         likes=moments_window.child_window(control_type="Button",title="赞")
#                         likes.click_input()
#                         pyautogui.press("down")
#                     except ElementNotFoundError:
#                         print("已点赞")
#                         pyautogui.press("down")
                     
#         elif judge_picture(contents[0])=="文字":
#                     button=contents[0].children()[0].children()[0].children()[1].children()[2].children(title="评论",control_type="Button")[0]
#                     button.click_input()
#                     time.sleep(2)
#                     try:
#                         likes=moments_window.child_window(control_type="Button",title="赞")
#                         likes.click_input()
#                         pyautogui.press("down")
#                     except ElementNotFoundError:
#                         print("已点赞")
#                         pyautogui.press("down")
       

        # if judge_picture(contents[1]):
        #             button=contents[1].children()[0].children()[0].children()[1].children()[3].children(title="评论",control_type="Button")[0]
        #             button.click_input()
        #             time.sleep(2)
        #             try:
        #                 likes=moments_window.child_window(control_type="Button",title="赞")
        #                 likes.click_input()
        #             except ElementNotFoundError:
        #                 print("已点赞")
        #                 pyautogui.press("down")
        # else:
        #             button=contents[1].children()[0].children()[0].children()[1].children()[2].children(title="评论",control_type="Button")[0]
        #             button.click_input()
        #             time.sleep(2)
        #             try:
        #                 likes=moments_window.child_window(control_type="Button",title="赞")
        #                 likes.click_input()
        #             except ElementNotFoundError:
        #                 print("已点赞")
        #                 pyautogui.press("down")
                   
        
def open_wechat_program(program_name,wechat_path:str=None,is_maximize:bool=True):
    main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(control_type="ToolBar",title='导航')
    program_button=Toolbar.children(control_type="Pane",title="")[1].children(control_type="Pane",title="")[0].children(control_type="Button",title='小程序面板')[0]
    program_button.click_input()
    desktop=Desktop(backend="uia")
    program_window=desktop.window(control_type="Pane",class_name="Chrome_WidgetWin_0",title="微信")
    main_window.close()
    search=program_window.child_window(control_type="Pane",title="",found_index=1).child_window(title="",control_type="Tab").children(title="",control_type="Pane")[1]
    search.click_input()
    search.type_keys(program_name)
    pyautogui.press("enter")
    search_result=program_window.child_window(control_type="Document",class_name="Chrome_RenderWidgetHostHWND")
    program_button=search_result.child_window(control_type="ListItem",found_index=4)
    program_button.click_input()
    main=search_result.child_window(control_type="Button",title_re=program_name,found_index=0,framework_id="Chrome")
    time.sleep(2)
    main.click_input()
    time.sleep(2)
    program_window=desktop.window(title=program_name,framework_id="Win32",control_type="Pane",class_name="Chrome_WidgetWin_0")
    return program_window

def open_wechat_official_account(official_acount_name,wechat_path:str=None,is_maximize:bool=True):
    main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
    Toolbar=main_window.child_window(control_type="ToolBar",title='导航')
    program_button=Toolbar.children(control_type="Pane",title="")[1].children(control_type="Pane",title="")[0].children(control_type="Button",title='小程序面板')[0]
    program_button.click_input()
    desktop=Desktop(backend="uia")
    program_window=desktop.window(control_type="Pane",class_name="Chrome_WidgetWin_0",title="微信")
    main_window.close()
    search=program_window.child_window(control_type="Pane",title="",found_index=1).child_window(title="",control_type="Tab").children(title="",control_type="Pane")[1]
    search.click_input()
    search.type_keys(official_acount_name)
    pyautogui.press("enter")
    search_result=program_window.child_window(control_type="Document",class_name="Chrome_RenderWidgetHostHWND")
    official_acount_button=search_result.child_window(control_type="ListItem",found_index=3)
    official_acount_button.click_input()
    result=search_result.child_window(control_type="Button",title_re=official_acount_name,found_index=0,framework_id="Chrome")
    time.sleep(2)
    result.click_input()
    time.sleep(2)

