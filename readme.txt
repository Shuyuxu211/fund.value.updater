如何在 Windows 任务计划程序中设置每天 16:00 自动运行脚本
1. 建议先创建一个简单的 .bat 启动脚本（方便设置工作目录）
假设你的 main.py 放在：D:\fund_auto\main.py，则在同目录下新建一个 run_fund_update.bat，内容如下：
@echo offcd /d D:\fund_autopython main.py
> 注意：
> - D:\fund_auto 替换为你实际存放脚本的目录；
> - 确保 python 已加入系统 PATH（在命令行直接输入 python 能运行），否则需要写全路径，例如：
> C:\Users\你的用户名\AppData\Local\Programs\Python\Python312\python.exe main.py
2. 在任务计划程序中创建计划任务
打开「任务计划程序」：
按 Win + R，输入 taskschd.msc，回车。
在左侧选择「任务计划程序库」；
右侧点击「创建基本任务」：
名称：例如“基金净值自动更新”
描述：可填“每天 16:00 自动更新基金净值 Excel”
触发器：
选择「每天」；
开始时间设置为 16:00:00；
下一步。
操作：
选择「启动程序」；
程序或脚本：选择刚才的 run_fund_update.bat（点击“浏览”）；
起始于（可选）：填写脚本所在目录，例如 D:\fund_auto；
下一步。
确认无误后点击「完成」。
只要电脑在 16:00 处于开机状态（不完全关机），任务就会自动执行脚本，更新 fund_data 文件夹中的当日 Excel。
小提示
首次使用：
在 主推公募、ETF、个人关注基金 的 A 列第2行开始，手动填入你关注的基金代码（6位，如 000001），保存后运行脚本即可。
多次运行：
同一天多次运行会覆盖同一个 fund_tracker_YYYY-MM-DD.xlsx 文件中的数据和“更新时间”。
扩展更多指标：
未来如果要增加「近1月涨跌幅」「今年以来」，可以在 HEADERS 里增加列名，并在 fetch_from_akshare 中多计算几个周期，然后在 update_sheet 里写入对应列即可。