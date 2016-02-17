# Email_My_PC
命令行版 Email My Pc，所有操作通过邮件标题操作，命令参数与命令本体用空格分开，只执行来自白名单邮件列表的命令。

## 支持命令列表
所有命令都可以在 config.ini 中配置

### #say hello
调用 tkinter 组件，弹出弹窗，内容为 hello，并返回回执.

### #screen
调用 ImageGrab 组件，截屏，并发送回命令发布者。

### #cmd
调用 subprocess 组件，执行 参数 ，并返回回执。
如:

    开启默认浏览器，打开页面:    #cmd start http://twitter.com
    通过默认文本编辑器打开文本:  #cmd notepad.exe test.txt
    ...

### #shutdown
调用 subprocess 组件，执行 关机 操作，目前仅支持 windows 系统

## 执行
windows 下面可以直接双击 run_bat.vbe 执行，或者添加该文件的快捷方式到启动列表中。执行时不显示窗口。

### 开启时即发送邮件
可配合配置文件中的 startsend 实现开机的时候自动发送回执到所有的白名单邮件列表。

### 查询间隔
修改配置文件中的 sleep 实现，单位秒。

## Environment

    OS:Windows
    Python edition:Python2.7
    Packages:pywin32/email
