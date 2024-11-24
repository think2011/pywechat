# pywechat🥇
![image](https://github.com/Hello-Mr-Crab/pywechat/blob/main/pics/introduction.jpg)
## 🍬🍬pywechat是一款基于pywinauto实现的Windows系统下PC微信自动化的Python项目。它可以帮助用户实现微信的一系列自动化操作，包括发送消息、发送文件、自动回复以及针对微信好友的所有操作，针对微信群聊的所有操作。

### 微信版本:3.9.12.17
### 操作系统:🪟windows 10 🪟windows 11
### python版本🐍:3.1x
### pywechat项目结构：
![image](https://github.com/Hello-Mr-Crab/pywechat/blob/main/pics/pywechat_structure.jpg)
<br>

##     该项目内的函数与方法名称与PC微信英文版各界面与功能英文翻译一致。其中pywechat的open_wechat函数无论微信是否打开，是否登录(需先前登录过,手机端勾选自动登录)均可正常打开微信,你只需要将微信WeChat.exe文件地址传入pywechat各个函数，或添加到windows系统环境变量中即可开启微信自动化之旅。🗺️🗺️
这里强烈建议将微信Wechat.exe文件添加到windows系统环境变量中，因为pywechat默认使用windows环境变量中的Wechat.exe路径启动微信,此时调用其中的每个方法与函数无需传入wechat_path参数即可自动化操作微信。

<br>

### 使用方法:
```
pip install pywechat127
```
<br>

### 添加微信至windows用户环境变量:
#### pywechat已内置自动添加微信至用户环境变量的方法 
```
from pywechat127.wechatTools import Tools
Tools.set_wechat_as_environ_path()
```
<br>

### wechatTools🌪️🌪️
#### 模块包括:
#### Tools:关于PC微信的一些工具,包括3个关于PC微信程序的方法和10个打开PC微信内各个界面的open系列方法。
#### API:打开指定微信小程序，指定公众号,打开视频号的功能，若有其他开发者想自动化操作上述程序可调用此API。
#### 函数:该模块内所有函数为上述模块内的所有方法。
<br>

### wechatauto🛏️🛏️
#### 模块包括：
##### Messages: 5种类型的发送消息方法，包括:单人单条,单人多条,多人单条,多人多条,转发消息:多人同一条。 
##### Files: 5种类型的发送文件方法，包括:单人单个,单人多个,多人单个,多人多个,转发文件:多人同一个。
##### FriendSettings: 涵盖了PC微信针对某个好友的全部操作的方法。
##### GroupSettings: 涵盖了PC微信针对某个群聊的全部操作的方法。
##### Contacts: 获取3种类型通讯录好友的备注与昵称包括:微信好友,企业号微信,群聊名称与人数，数据返回格式为json。
##### Call: 给某个好友打视频或语音电话。
##### Auto_response:自动接应微信视频或语音电话。
#### 函数:该模块内所有函数为上述模块内的所有方法。  
<br>

### winSettings🔹🔹
#### 模块包括：
#### Systemsettings:该模块中提供了7个修改windows系统设置和3个判断文件类型的方法。
#### 函数：该模块内所有函数为上述模块内的所有方法。
<br>

### 使用示例:
#### (注意，微信WeChat.exe路径已添加至windows系统环境变量,故以下方法或函数无需传入wechat_path这一参数)
#### 给某个好友发送多条信息：
```
 from pywechat.wechatauto import Messages
Messages.send_messages_to_friend(friend="文件传输助手",messages=['你好','我正在使用pywechat操控微信给你发消息','收到请回复'])
```
##### 或者
```
import pywechat.wechatauto as wechat
wechat.send_messages_to_friend(friend="文件传输助手",messages=['你好','我正在使用pywechat操控微信给你发消息','收到请回复'])
```
<br>

#### 自动接听语音视频电话:
```
from pywechat.wechatauto import Auto_response
Auto_response.auto_answer_call(broadcast_content='您好，我目前不在线我的PC微信正在由我的微信机器人控制请稍后再试',message='您好，我目前不在线我的PC微信正在由我的微信机器人控制请稍后再试',duration='1h',times=1)
```
##### 或者
```
import pywechat.wechatauto as wechat
wechat.auto_answer_call(broadcast_content='您好，我目前不在线我的PC微信正在由我的微信机器人控制请稍后再试',message='您好，我目前不在线我的PC微信正在由我的微信机器人控制请稍后再试',duration='1h',times=1)
```
