# AutoClose WinCC Info & TaskDaemon

当你没有授权,没有狗,Wincc会定期弹窗</br>

  
<img src="https://i-blog.csdnimg.cn/blog_migrate/2b31e56e380871ba7fae5bc11c625ef2.png"  height="200" /><img src="https://www.ad.siemens.com.cn/service/answer/Uploads/questionimgs/20221205110735_13.png"  height="200" />
<br>

这个小程序托盘运行,自动查找窗口并关闭,窗口名称包括"WinCC Information"和"WinCC 信息" ,每秒循环一次</br>

更新一个功能: 自动守护计划任务的进程</br>

自动检测PdlRt进程(画面)是否存在, 如果存在再检测gscrt进程(计划任务)是否存在</br>

如果不存在自动运行"C:\Program Files (x86)\Siemens\Automation\SCADA-RT_V11\WinCC\bin\gscrt.exe"</br>

有需要的自取</br>


