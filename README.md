# 侧边笔记PPT转PDF(Side notes pptx 2 pdf)
Convert PPT to PDF with blank side note areas, customizable page margins and the number of PPT pages per PDF page.
将PPT转化为包含空白侧边笔记区域的PDF，可定制页边距和每页PDF包含的PPT页数
# 使用方法(How to use)
- 双击 **install.bat** 进行安装
- 双击 **run.bat** 启动程序
- 当需要输入路径时，使用 **Windows** 的复制文件路径功能，例如 **"D:\\School\\PPT\\DNA Replication.pptx"**，其它系统的路径按照该路径书写，**必须为绝对路径**，并粘贴到程序中
- **config.json** 中可以设置页边距 **(margin)** 和每页PDF包含的PPT页数 **(pages_in_page)** ，**margin** 数组分别对应上边距、下边距、左边距、右边距
# 注意事项(Precautions)
- Python 脚本**理论上**可用于所有系统，仅测试过 Windows
- 系统**必须已经安装 Office**，仅测试过 MS-Office