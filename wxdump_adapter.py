import time
import pyautogui
from pywechat.Errors import NoChatHistoryError
from pywechat.Uielements import Lists
from pywechat.WechatAuto import check_new_message, get_groupmembers_info, save_photos
from pywechat.WechatTools import Tools,mouse

Lists=Lists() #所有列表类型UI

def trigger_file_downloads(friend:str, number:int=100):
    #打开好友的对话框,返回值为编辑消息框和主界面
    chat_history_window,main_window=Tools.open_chat_history(friend=friend,TabItem='文件',close_wechat=True,is_maximize=True,)
    
    fileList=chat_history_window.child_window(**Lists.FileList)
    if not fileList.exists():
        chat_history_window.close()
        if main_window:
            main_window.close()
        print(NoChatHistoryError(f'没有找到好友{friend}的聊天记录'))
        return
    x,y=fileList.rectangle().right-8,fileList.rectangle().top+5
    mouse.click(coords=(x,y))
    pyautogui.press('Home')#回到最顶部
    
    last_selected_item = None
    for i in range(number):
        selected_item = [item for item in fileList.children(control_type='ListItem') if item.is_selected()]
        if not selected_item:
            break
        
        selected_item = selected_item[0]
        if last_selected_item == selected_item:
            print("已到达列表底部")
            break
        last_selected_item = selected_item
        
        print(selected_item)    
        
        # 触发下载
        if selected_item.descendants(control_type='Text', title='未下载'):
            print(f'触发下载: {selected_item}')
            selected_item.click_input()
            time.sleep(1) 
            
        pyautogui.press('down', _pause=False, presses=1)
        
    
    chat_history_window.close()

def main():
    DELAY = 5  # 每次检查新消息的间隔时间（秒）
    while True:
        new_msgs = check_new_message()
        if not new_msgs:
            time.sleep(DELAY)
            continue
        
        for msg_item in new_msgs:
            friend_name = msg_item.get('好友名称')
            number = msg_item.get('新消息条数', 10)
            number = number * 2  # 避免遗漏
            
            # 触发文件和图片下载
            if friend_name.startswith('ISH-'):
                save_photos(friend=friend_name, number=number)
                trigger_file_downloads(friend=friend_name, number=number)
            
        # 继续检查
        time.sleep(DELAY)

def test():
    pass

main()
