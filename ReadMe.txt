名称：Start
版本：0.98
作者：qbj196@qq.com
功能：跳过Windows 8 Metro界面，直接显示桌面。
安装方法：打开命令提示符，输入“regsvr32 (Start.dll所在路径)\Start.dll”，32位系统用32为dll，64位系统用64位dll。
使用方法：鼠标右键单击任务栏，选择“工具栏(T)”，选择“开始(S)”即可，以后启动系统就可以跳过Windows 8 Metro界面直接显示桌面了。
卸载方法：鼠标右键单击任务栏，选择“工具栏(T)”，取消选择“开始(S)”，然后打开命令提示符，输入“regsvr32 /u (Start.dll所在路径)\Start.dll”，删除Start.dll文件即可。