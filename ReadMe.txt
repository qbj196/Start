名称：Start
版本：1.00
作者：qbj196@qq.com
功能：跳过Windows 8开始屏幕，直接显示桌面，桌面模式下隐藏边角导航和顶部分屏功能，鼠标右键单击开始按钮显示开始屏幕，支持Windows 8、8.1。
安装方法：打开命令提示符(管理员)，输入“regsvr32 (Start.dll所在路径)\Start.dll”，32位系统用32位dll，64位系统用64位dll。
使用方法：鼠标右键单击任务栏，选择“工具栏(T)”，选择“开始(S)”即可，以后启动系统就可以跳过Windows 8开始屏幕直接显示桌面了。
卸载方法：鼠标右键单击任务栏，选择“工具栏(T)”，取消选择“开始(S)”，然后打开命令提示符(管理员)，输入“regsvr32 /u (Start.dll所在路径)\Start.dll”，删除Start.dll文件即可。
截图：https://github.com/qbj196/Screenshot/blob/main/Start-0.99.png
截图：https://github.com/qbj196/Screenshot/blob/main/Start-1.00.png
