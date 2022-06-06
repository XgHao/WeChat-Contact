# 微信联系人导出 (Export WeChat Contact)

基于AHK v2实现。

# 原理
通过模拟键鼠操作实现微信好友导出，零封号危险。

# 使用步骤
运行exe后，打开PC微信窗口
- （可选）切换到通讯录，手动过一下好友，先把好友信息都加载出来，可以明显提高脚本运行速度
- win+c 运行
- 然后再提示框中输入要导出的微信好友数，微信好友数可以通讯录->通讯录管理中查询
- win+esc 退出
  
![微信好友数](https://github.com/XgHao/WeChat-Contact/blob/main/pic/contact.png?raw=true)

# 版本须知
PC端微信3.7.0版本，对好友页进行了更改，
- 对于>=3.7.0的版本：不支持【来源，签名】信息，但支持【标签】信息。
- 对于<3.7.0的版本：支持【来源，签名】信息，但不支持【标签】信息

> 如果不需要来源和签名信息，建议使用微信新版本，新版本导出效率明显高于老版本
 
**具体区别如下**
![版本区别](https://github.com/XgHao/WeChat-Contact/blob/main/pic/diff.png?raw=true)

# 本机测试环境
- Windows 10 企业版 2004
- office 365
