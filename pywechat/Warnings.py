
from warnings import warn
class LongTextWarning(Warning):
    '''消息字数超长警告'''
    #微信字数限制2000字，超长部分不发送，这时需要转化为word发送
    pass
class InputMethodWarning(Warning):#使用pywinauto的type_keys方法时
    #电脑的输入法为微信输入法时无法打字，需要及时转换为ctrl+v复制消息发送，并按下win+空格切换一下输入法
    '''输入法警告'''
    pass
class ChatHistoryNotEnough(Warning):
    '''聊天记录不足指定的个数'''
    pass