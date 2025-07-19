import time
import pyautogui
from functools import wraps
from pywechat.WechatTools import Tools
from pywechat.WinSettings import Systemsettings
from pywechat.WechatTools import match_duration,mouse
from pywechat.Errors import TimeNotCorrectError,NotFriendError
from pywechat.Uielements import Buttons,Main_window,Texts,Edits,SideBar
Buttons=Buttons()
Main_window=Main_window()
Texts=Texts()
Edits=Edits()
SideBar()
language=Tools.language_detector()
def auto_reply_to_friend_decorator(duration:str,friend:str,search_pages:int=5,delay:int=0.2,wechat_path:str=None,is_maximize:bool=True,close_wechat:bool=True):
    '''
    该函数为自动回复指定好友的修饰器\n
    Args:
        friend:好友或群聊备注\n
        duration:自动回复持续时长,格式:'s','min','h',单位:s/秒,min/分,h/小时\n
        search_pages:在会话列表中查询查找好友时滚动列表的次数,默认为5,一次可查询5-12人,当search_pages为0时,直接从顶部搜索栏搜索好友信息打开聊天界面\n
        folder_path:存放聊天记录截屏图片的文件夹路径\n
        wechat_path:微信的WeChat.exe文件地址,主要针对未登录情况而言,一般而言不需要传入该参数,因为pywechat会通过查询环境变量,注册表等一些方法\n
            尽可能地自动找到微信路径,然后实现无论PC微信是否启动都可以实现自动化操作,除非你的微信路径手动修改过,发生了变动的话可能需要\n
            传入该参数。最后,还是建议加入到环境变量里吧,这样方便一些。加入环境变量可调用set_wechat_as_environ_path函数\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        close_wechat:任务结束后是否关闭微信,默认关闭\n
    '''
    def decorator(reply_func):
        @wraps(reply_func)
        def wrapper():
            if not match_duration(duration):#不按照指定的时间格式输入,需要提前中断退出
                raise TimeNotCorrectError
            edit_area,main_window=Tools.open_dialog_window(friend=friend,wechat_path=wechat_path,is_maximize=is_maximize,search_pages=search_pages)
            voice_call_button=main_window.child_window(**Buttons.VoiceCallButton)
            video_call_button=main_window.child_window(**Buttons.VideoCallButton)
            if not voice_call_button.exists():
                #公众号没有语音聊天按钮
                main_window.close()
                raise NotFriendError(f'非正常好友,无法自动回复!')
            if not video_call_button.exists() and voice_call_button.exists():
                main_window.close()
                raise NotFriendError('auto_reply_to_friend只用来自动回复好友,如需自动回复群聊请使用auto_reply_to_group!')
            chatList=main_window.child_window(**Main_window.FriendChatList)#聊天界面内存储所有信息的容器
            initial_last_message=Tools.pull_latest_message(chatList)[0]#刚打开聊天界面时的最后一条消息的listitem   
            Systemsettings.open_listening_mode(full_volume=False)#开启监听模式,此时电脑只要不断电不会息屏 
            start_time=time.time()  
            while True:
                if time.time()-start_time<match_duration(duration):#将's','min','h'转换为秒
                    newMessage,who=Tools.pull_latest_message(chatList)
                    #消息列表内的最后一条消息(listitem)不等于刚打开聊天界面时的最后一条消息(listitem)
                    #并且最后一条消息的发送者是好友时自动回复
                    #这里我们判断的是两条消息(listitem)是否相等,不是文本是否相等,要是文本相等的话,对方一直重复发送
                    #刚打开聊天界面时的最后一条消息的话那就一直不回复了
                    if newMessage!=initial_last_message and who==friend:
                        reply_content=reply_func(newMessage)
                        Systemsettings.copy_text_to_windowsclipboard(reply_content)
                        pyautogui.hotkey('ctrl','v',_pause=False)
                        time.sleep(delay)
                        pyautogui.hotkey('alt','s',_pause=False)
                else:
                    break
            Systemsettings.close_listening_mode()
            if close_wechat:
                main_window.close()
        return wrapper
    return decorator 

def auto_reply_to_group_decorator(duration:str,group_name:str,search_pages:int=5,wechat_path:str=None,at_only:bool=False,maxReply:int=3,at_other:bool=True,is_maximize:bool=True,close_wechat:bool=True):
    '''
    该函数为自动回复指定群聊的修饰器\n
    Args:
        friend:好友或群聊备注\n
        duration:自动回复持续时长,格式:'s','min','h',单位:s/秒,min/分,h/小时\n
        search_pages:在会话列表中查询查找好友时滚动列表的次数,默认为5,一次可查询5-12人,当search_pages为0时,直接从顶部搜索栏搜索好友信息打开聊天界面\n
        folder_path:存放聊天记录截屏图片的文件夹路径\n
        wechat_path:微信的WeChat.exe文件地址,主要针对未登录情况而言,一般而言不需要传入该参数,因为pywechat会通过查询环境变量,注册表等一些方法\n
            尽可能地自动找到微信路径,然后实现无论PC微信是否启动都可以实现自动化操作,除非你的微信路径手动修改过,发生了变动的话可能需要\n
            传入该参数。最后,还是建议加入到环境变量里吧,这样方便一些。加入环境变量可调用set_wechat_as_environ_path函数\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        close_wechat:任务结束后是否关闭微信,默认关闭\n
    '''
    def decorator(reply_func):
        @wraps(reply_func)
        def wrapper():
            def at_others(who):
                edit_area.click_input()
                edit_area.type_keys(f'@{who}')
                pyautogui.press('enter',_pause=False)
            def send_message(newMessage,who,reply_func):
                if at_only:
                    if who!=myname and f'@{myalias}' in newMessage:#如果消息中有@我的字样,那么回复
                        if at_other:
                            at_others(who)
                        reply_content=reply_func(newMessage)
                        Systemsettings.copy_text_to_windowsclipboard(reply_content)
                        pyautogui.hotkey('ctrl','v',_pause=False)
                        pyautogui.hotkey('alt','s',_pause=False)
                    else:#消息中没有@我的字样不回复
                        pass
                if not at_only:#at_only设置为False时,只要有人发新消息就自动回复
                    if who!=myname:
                        if at_other:
                            at_others(who)
                        reply_content=reply_func(newMessage)
                        Systemsettings.copy_text_to_windowsclipboard(reply_content)
                        pyautogui.hotkey('ctrl','v',_pause=False)
                        pyautogui.hotkey('alt','s',_pause=False)
                    else:
                        pass
            if not match_duration(duration):#不按照指定的时间格式输入,需要提前中断退出
                raise TimeNotCorrectError
            #打开好友的对话框,返回值为编辑消息框和主界面
            Systemsettings.set_english_input()
            edit_area,main_window=Tools.open_dialog_window(friend=group_name,wechat_path=wechat_path,is_maximize=is_maximize,search_pages=search_pages)
            myname=main_window.child_window(**Buttons.MySelfButton).window_text()#我的昵称
            chat_history_button=main_window.child_window(**Buttons.ChatHistoryButton)
            #需要判断一下是不是公众号
            if not chat_history_button.exists():
                #公众号没有语音聊天按钮
                main_window.close()
                raise NotFriendError(f'非正常群聊,无法自动回复!')
            #####################################################################################
            #打开群聊右侧的设置界面,看一看我的群昵称是什么,这样是为了判断我是否被@
            ChatMessage=main_window.child_window(**Buttons.ChatMessageButton)
            ChatMessage.click_input()
            group_settings_window=main_window.child_window(**Main_window.GroupSettingsWindow)
            group_settings_window.child_window(**Texts.GroupNameText).click_input()
            group_settings_window.child_window(**Buttons.MyAliasInGroupButton).click_input() 
            change_my_alias_edit=group_settings_window.child_window(**Edits.EditWnd)
            change_my_alias_edit.click_input()
            myalias=change_my_alias_edit.window_text()#我的群昵称
            ########################################################################
            chatList=main_window.child_window(**Main_window.FriendChatList)#聊天界面内存储所有信息的容器
            x,y=chatList.rectangle().left+8,(main_window.rectangle().top+main_window.rectangle().bottom)//2#
            mouse.click(coords=(x,y))
            responsed=[]
            initialMessages=Tools.pull_messages(friend=group_name,number=maxReply,search_pages=search_pages,wechat_path=wechat_path,is_maximize=is_maximize,close_wechat=False,parse=False)
            responsed.extend(initialMessages) 
            Systemsettings.open_listening_mode(full_volume=False)#开启监听模式,此时电脑只要不断电不会息屏 
            start_time=time.time()  
            while True:
                if time.time()-start_time<match_duration(duration):
                    newMessages=Tools.pull_messages(friend=group_name,number=maxReply,search_pages=search_pages,wechat_path=wechat_path,is_maximize=is_maximize,close_wechat=False,parse=False)
                    filtered_newMessages=[newMessage for newMessage in newMessages if newMessage not in responsed]
                    for newMessage in filtered_newMessages:
                        message_sender,message_content,message_type=Tools.parse_message_content(ListItem=newMessage,friendtype='群聊')
                        send_message(message_content,message_sender,reply_func)
                        responsed.append(newMessage)
                else:
                    break
            if close_wechat:
                main_window.close()
        return wrapper
    return decorator
    
from pywechat import *
from functools import wraps
def auto_reply_groups_decorator(duration:str,max_pages:int=5,never_reply:list=[],scroll_delay:int=0,wechat_path:str=None,is_maximize:bool=True,close_wechat:bool=True)->None:
    '''
    该装饰器用来遍历会话列表查找新消息如果是群聊且该群聊新消息中含有@我的字样则自动回复的否则不回复
    最大回复数量=max_pages*(8~10)\n
    Args:
        duration:自动回复持续时长,格式:'s','min','h'单位:s/秒,min/分,h/小时\n
        max_pages:遍历会话列表页数,一页为8~10人,设定持续时间后,将持续在max_pages内循环遍历查找是否有新消息\n
        never_reply:在never_reply列表中的好友即使有新消息时也不会回复\n
        scroll_delay:滚动遍历max_pages页会话列表后暂停秒数,如果你的max_pages很大,且持续时间长,scroll_delay还为0的话,那么一直滚动遍历有可能被微信检测到自动退出登录\n
            该参数只在会话列表可以滚动的情况下生效\n
        wechat_path:微信的WeChat.exe文件地址,主要针对未登录情况而言,一般而言不需要传入该参数,因为pywechat会通过查询环境变量,注册表等一些方法\n
            尽可能地自动找到微信路径,然后实现无论PC微信是否启动都可以实现自动化操作,除非你的微信路径手动修改过,发生了变动的话可能需要\n
            传入该参数。最后,还是建议加入到环境变量里吧,这样方便一些。加入环境变量可调用set_wechat_as_environ_path函数\n
        is_maximize:微信界面是否全屏,默认全屏。\n
        close_wechat:任务结束后是否关闭微信,默认关闭\n
    '''
    def decorator(reply_func):
        @wraps(reply_func)
        def wrapper():
            
            def check_my_alias(chat_name):
                ChatMessage=main_window.child_window(**Buttons.ChatMessageButton)
                ChatMessage.click_input()
                group_settings_window=main_window.child_window(**Main_window.GroupSettingsWindow)
                group_settings_window.child_window(**Texts.GroupNameText).click_input()
                group_settings_window.child_window(**Buttons.MyAliasInGroupButton).click_input() 
                change_my_alias_edit=group_settings_window.child_window(**Edits.EditWnd)
                change_my_alias_edit.click_input()
                myalias=change_my_alias_edit.window_text()#我的群昵称
                alias_map[chat_name]=myalias#使用alias_map这个字典缓存我在不同群聊的群昵称,这样不用每次进入回复过的群聊都看一遍我的群昵称
                main_window.click_input()
                return myalias

            def at_others(edit_area,who):
                edit_area.click_input()
                edit_area.type_keys(f'@{who}')
                pyautogui.press('enter',_pause=False)

            def record():
                time.sleep(1)
                #遍历当前会话列表内可见的所有成员，获取他们的名称和新消息条数，没有新消息的话返回[],[]
                #newMessagefriends为会话列表(List)中所有含有新消息的ListItem
                newMessagefriends=[friend for friend in messageList.items() if '条新消息' in friend.window_text()]
                if newMessagefriends:
                    #newMessageTips为newMessagefriends中每个元素的文本:['测试365 5条新消息','一家人已置顶20条新消息']这样的字符串列表
                    newMessageTips=[friend.window_text() for friend in newMessagefriends]
                    #会话列表中的好友具有Text属性，Text内容为备注名，通过这个按钮的名称获取好友名字
                    names=[friend.descendants(control_type='Text')[0].window_text() for friend in newMessagefriends]
                    #此时filtered_Tips变为：['5条新消息','20条新消息']直接正则匹配就不会出问题了
                    filtered_Tips=[friend.replace(name,'') for name,friend in zip(names,newMessageTips)]
                    nums=[int(re.findall(r'\d+',tip)[0]) for tip in filtered_Tips]
                    return names,nums 
                return [],[]

            #监听并且回复右侧聊天界面
            def listen_on_current_chat():
                voice_call_button=main_window.child_window(**Buttons.VoiceCallButton)
                video_call_button=main_window.child_window(**Buttons.VideoCallButton)
                current_chat=main_window.child_window(**Main_window.CurrentChatWindow)
                #判断好友类型
                if video_call_button.exists() and voice_call_button.exists():#好友
                    type='好友'
                if not video_call_button.exists() and voice_call_button.exists():#好友
                    type='群聊'
                if not video_call_button.exists() and not voice_call_button.exists():#好友
                    type='非好友'

                if type=='群聊' and current_chat.window_text() not in taboo_list:
                    latest_message,who=Tools.pull_latest_message(chatlist)#最新的消息
                    myalias=alias_map.get(current_chat.window_text()) if alias_map.get(current_chat.window_text()) else check_my_alias(current_chat.window_text())
                    if latest_message!=initial_last_message and f'@{myalias}' in latest_message:
                        reply_content=reply_func(latest_message)
                        at_others(current_chat,who)
                        Systemsettings.copy_text_to_windowsclipboard(reply_content)
                        pyautogui.hotkey('ctrl','v',_pause=False)
                        pyautogui.hotkey('alt','s',_pause=False)
                        responsed_groups.add(current_chat.window_text())
                        if scorllable:
                            mouse.click(coords=(x+2,y-6))#点击右上方激活滑块

            #用来回复在会话列表中找到的头顶有红色数字新消息提示的好友
            def reply(names,nums):
                for name,number in dict(zip(names,nums)).items():
                    if name not in taboo_list:       
                        Tools.find_friend_in_MessageList(friend=name,search_pages=search_pages,is_maximize=is_maximize)
                        voice_call_button=main_window.child_window(**Buttons.VoiceCallButton)
                        video_call_button=main_window.child_window(**Buttons.VideoCallButton)
                        #判断好友类型
                        if video_call_button.exists() and voice_call_button.exists():#好友
                            type='好友'
                        if not video_call_button.exists() and voice_call_button.exists():#好友
                            type='群聊'
                        if not video_call_button.exists() and not voice_call_button.exists():#好友
                            type='非好友'
                        if type=='群聊':
                            current_chat=main_window.child_window(**Main_window.CurrentChatWindow)
                            message_contents,message_senders=Tools.pull_messages(friend=name,number=number,close_wechat=False)[:2]
                            myalias=alias_map.get(current_chat.window_text()) if alias_map.get(current_chat.window_text()) else check_my_alias(current_chat.window_text())
                            for message,who in zip(message_contents,message_senders):
                                if f"@{myalias}" in message:
                                    reply_content=reply_func(message)
                                    at_others(current_chat,who)
                                    Systemsettings.copy_text_to_windowsclipboard(reply_content)
                                    pyautogui.hotkey('ctrl','v',_pause=False)
                                    pyautogui.hotkey('alt','s',_pause=False)
                                    responsed_groups.add(name)
                if scorllable:
                    mouse.click(coords=(x,y))#回复完成后点击右上方,激活滑块，继续遍历会话列表

            if not match_duration(duration):
                raise TimeNotCorrectError
            if language=='简体中文':
                taboo_list=['微信团队','微信支付','微信运动','订阅号','腾讯新闻','服务通知','微信游戏']
            if language=='繁体中文':
                taboo_list=['微信团队','微信支付','微信运动','訂閱賬號','騰訊新聞','服務通知','微信游戏']
            if language=='英文':
                taboo_list=['微信团队','微信支付','微信运动','Subscriptions','Tencent News','Service Notifications','微信游戏']
            taboo_list.extend(never_reply)
            responsed_groups=set()
            alias_map={}
            main_window=Tools.open_wechat(wechat_path=wechat_path,is_maximize=is_maximize)
            chatsButton=main_window.child_window(**SideBar.Chats)
            chatsButton.click_input()
            chatlist=main_window.child_window(**Main_window.FriendChatList)#聊天界面内存储所有聊天信息的消息列表
            initial_last_message=Tools.pull_latest_message(chatlist)[0]#刚打开聊天界面时的最后一条消息的listitem 
            messageList=main_window.child_window(**Main_window.ConversationList)#左侧的会话列表
            scorllable=Tools.is_VerticalScrollable(messageList)#只调用一次is_VerticallyScrollable函数来判断会话列表是否可以滚动
            x,y=messageList.rectangle().right-5,messageList.rectangle().top+8#右上方滑块的位置
            if scorllable:
                mouse.click(coords=(x,y))#点击右上方激活滑块
                pyautogui.press('Home')#按下Home健确保从顶部开始
            search_pages=1
            Systemsettings.open_listening_mode(full_volume=False)
            start_time=time.time()
            while time.time()-start_time<=match_duration(duration):
                if chatsButton.legacy_properties().get('Value'):#如果左侧的聊天按钮是红色的就遍历,否则原地等待
                    if scorllable:
                        for _ in range(max_pages+1):
                            names,nums=record()
                            reply(names,nums)
                            pyautogui.press('pagedown',_pause=False)
                            search_pages+=1
                        pyautogui.press('Home')
                        time.sleep(scroll_delay)
                    else:
                        names,nums=record()
                        reply(names,nums)
                listen_on_current_chat()
            Systemsettings.close_listening_mode()
            if responsed_groups:
                print(f"在{duration}内回复了以下好友\n{responsed_groups}")
            if close_wechat:
                main_window.close()
        return wrapper
    return decorator

# @auto_reply_to_friend_decorator(duration='30s',friend='测试ing365')
# def reply_func(newMessage):
#     if '你好' in newMessage:
#         return '你好,有什么需要帮助的吗?'
#     if '在吗' in newMessage:
#         return '不好意思,当前不在,请稍后联系'
#     return '我的微信正在被pywechat控制'   
# reply_func()

#@auto_reply_to_group_decorator(duration='30s',group_name='Pywechat测试群',at_other=True)
# def reply_func(newMessage):
#     if '你好' in newMessage:
#         return '在的,亲😙请问有什么需要帮助的吗？'
#     if '售后' in newMessage:
#         return '您可以点击下方链接，申请售后服务'
#     if '算了' in newMessage:
#         return '很遗憾未能未您提供服务，欢迎下次选购'
#     else:
#         return '欢迎进店选购，祝您生活愉快'   
# reply_func()

#@auto_reply_groups_decorator(duration='1min')
#def reply_func(newmessage):
#   if '你好' in newmessage:
#        return '你好,请问有什么可以帮助你的吗?'
#    if '售后' in newmessage:
#        return '你可以点击以下链接申请售后服务'
#    else:
#        return '不好意思,不明白您的需求'
#reply_func()

# import requests
# import json
# #接入大模型API自动回复
# @auto_reply_to_group_decorator(duration='30s',group_name='Pywechat测试群',at_other=True)
# def reply_func(newmessage):
#     # API URL
#     url = 'https://api.coze.cn/v1/workflow/run'
    
#     # Headers
#     headers = {
#         'Authorization': '',  # 替换为真实的token
#         'Content-Type': 'application/json'
#     }
#     # 请求数据
#     data = {
#         "workflow_id": "",  #替换为真实的workflow_id
#         "parameters": {   
#             "input": f"{newmessage}"
#         }
#     }
#     response=requests.post(url, headers=headers, data=json.dumps(data))
#     return response.json()['data']
# reply_func()

