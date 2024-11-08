'''使用该模块的四个函数或方法时,你可以:\n
 from wechatauto.winSettings import Systemsettins \n
 导入Systemsettins,使用Systemsettings下的三个方法来修改相关的windows系统设置:\n
 Systemsettings.set_volume_to_100();Systemsettings.open_listening_mode();Systemsettings.close_listening_mode()\n
 或者:\n
 from wechatauto.Systemsettings import set_volume_to_100,open_listening_mode,close_listening_mode\n
 直接导入该三个方法的函数:\n
 set_volume_to_100,open_listening_mode,close_listening_mode\n
 '''
import ctypes
import win32com.client
import win32clipboard
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
devices=AudioUtilities.GetSpeakers()
interface=devices.Activate(IAudioEndpointVolume._iid_,CLSCTX_ALL,None)
volume=cast(interface, POINTER(IAudioEndpointVolume))
ES_DISPLAY_REQUIRED=0x00000002
ES_CONTINUOUS=0x80000000
ES_CONTINUOUS=0x80000000

class Systemsettings():
    '''该模块中有四个修改windows系统设置的方法,包括:  \n  
    set_volume_to_100:将windows系统音量设置为100。      \n
    open_listening_mode:开启监听模式,开启后屏幕保持常亮。  \n
    close_listening_mode:关闭监听模式,开启后结束屏幕保持常亮。  \n
    speaker:调用windows word中朗读文本的API来进行语音播报。  \n
    '''
    @staticmethod
    def set_volume_to_100():
        '''将系统音量设置为100'''
        mute=volume.GetMute()
        if mute==1:
            volume.SetMute(False,None)
        volume.SetMasterVolumeLevel(0.0, None)
    @staticmethod
    def open_listening_mode():
        '''用来开启监听模式,此时电脑音量设置为100,除非断电否则屏幕保持常亮'''
        ES_DISPLAY_REQUIRED=0x00000002
        ES_CONTINUOUS=0x80000000
        Systemsettings.set_volume_to_100()
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS|ES_DISPLAY_REQUIRED)
    @staticmethod    
    def close_listening_mode():
        '''用来关闭监听模式,需要与open_listening_mode函数结合使用,单独使用无意义''' 
        ES_CONTINUOUS=0x80000000
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
    @staticmethod    
    def speaker(text:str,times:int=1):
        '''
        text:朗读文本的内容\n
        times:重复朗读次数\n
        调用windows word中朗读文本的API来进行语音播报'''
        speaker=win32com.client.Dispatch("SAPI.SpVoice")
        for _ in range(times):
            speaker.speak(text)
    @staticmethod
    def copy_files_to_windowsclipboard(filepaths_list:list):
        '''
        filepaths_list:文件路径列表 
        该函数用来将windows系统下所有给定文件复制到剪贴板
        '''
        filepaths_list=[file_path.replace('/','\\') for file_path in filepaths_list]
        class DROPFILES(ctypes.Structure):
            _fields_ = [
                ("pFiles", ctypes.c_uint),
                ("x", ctypes.c_long),
                ("y", ctypes.c_long),
                ("fNC", ctypes.c_int),
                ("fWide", ctypes.c_bool),
            ]
        pDropFiles = DROPFILES()
        pDropFiles.pFiles = ctypes.sizeof(DROPFILES)
        pDropFiles.fWide = True
        #获取文件绝对路径
        files = ("\0".join(filepaths_list)).replace("/", "\\")
        data = files.encode("U16")[2:] + b"\0\0"        #结尾一定要两个\0\0字符，这是规定！
        win32clipboard.OpenClipboard()  #打开剪贴板（独占）
        try:
            #若要将信息放在剪贴板上，首先需要使用 EmptyClipboard 函数清除当前的剪贴板内容
            win32clipboard.EmptyClipboard() #清空当前的剪贴板信息
            win32clipboard.SetClipboardData(win32clipboard.CF_HDROP,bytes(pDropFiles)+data) #设置当前剪贴板数据
        except Exception as e:
            print("复制文件到剪贴板时出错！")
        finally:
            win32clipboard.CloseClipboard() #无论什么情况，都关闭剪贴板

def set_volume_to_100():
        '''将系统音量设置为100'''
        mute=volume.GetMute()
        if mute==1:
            volume.SetMute(False,None)
        volume.SetMasterVolumeLevel(0.0, None)

def open_listening_mode():
        '''用来开启监听模式,此时电脑音量设置为100,除非断电否则屏幕保持常亮'''
        ES_DISPLAY_REQUIRED=0x00000002
        ES_CONTINUOUS=0x80000000
        Systemsettings.set_volume_to_100()
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS|ES_DISPLAY_REQUIRED)
def close_listening_mode():
        '''用来关闭监听模式,需要与open_listening_mode函数结合使用,单独使用无意义''' 
        ES_CONTINUOUS=0x80000000
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
def speaker(text:str,times:int=1):
        '''
        text:朗读文本的内容\n
        times:重复朗读次数\n
        调用windows word中朗读文本的API来进行语音播报'''
        speaker=win32com.client.Dispatch("SAPI.SpVoice")
        for _ in range(times):
            speaker.speak(text)

def copy_file_to_windowsclipboard(filepaths_list:list):
    class DROPFILES(ctypes.Structure):
        _fields_ = [
            ("pFiles", ctypes.c_uint),
            ("x", ctypes.c_long),
            ("y", ctypes.c_long),
            ("fNC", ctypes.c_int),
            ("fWide", ctypes.c_bool),
        ]
    pDropFiles = DROPFILES()
    pDropFiles.pFiles = ctypes.sizeof(DROPFILES)
    pDropFiles.fWide = True
    #获取文件绝对路径
    files = ("\0".join(filepaths_list)).replace("/", "\\")
    data = files.encode("U16")[2:] + b"\0\0"        #结尾一定要两个\0\0字符，这是规定！
    win32clipboard.OpenClipboard()  #打开剪贴板（独占）
    try:
        #若要将信息放在剪贴板上，首先需要使用 EmptyClipboard 函数清除当前的剪贴板内容
        win32clipboard.EmptyClipboard() #清空当前的剪贴板信息
        win32clipboard.SetClipboardData(win32clipboard.CF_HDROP,bytes(pDropFiles)+data) #设置当前剪贴板数据
    except Exception as e:
        print("复制文件到剪贴板时出错！")
    finally:
        win32clipboard.CloseClipboard() #无论什么情况，都关闭剪贴板
