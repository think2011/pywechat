# pywechat
![image](https://github.com/Hello-Mr-Crab/pywechat/blob/master/introduction.jpg)
## pywechat是一款基于pywinauto实现的Windows系统下PC微信自动化的Python库。它可以帮助用户实现微信的自动化操作，包括发送消息、发送文件、自动回复以及针对微信好友的所有操作，针对微信群聊的所有操作。
### pywechat的使用非常简单，用户只需要通过调用相应的函数或模块即可实现所需的功能。同时，pywechat还提供了详细的文档和示例代码，方便用户学习和使用。

#### pywechat项目结构：
                            pywechat 
                        /     |       \
                      /	      |         \
                wechatTools wechatauto winSettings   

wechatTools:包含Toools与API两个模块。 Tools中封装了3个关于PC微信的工具,包括:
                                                                  is_wechat_running:用来判断PC微信是否运行。\n
                                                                  open_wechat:打开PC微信。\n
                                                                  find_friends_in_Messagelist:在会话列表和当前聊天窗口中查找好友\n
                                                                  以及10个open方法用于打开微信主界面内所有能打开的界面\n

