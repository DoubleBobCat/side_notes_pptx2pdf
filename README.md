# 侧边笔记PPT转PDF(简体中文|[English](README_en.md))
❤️将PPT转化为包含侧边笔记区域的PDF，较高可定制性![](https://img.shields.io/badge/快捷-易用-易用)
# 使用方法
- 双击 **install.bat** 进行安装
- 双击 **run.bat** 启动程序
- 当需要输入路径时，使用 **Windows** 的复制文件路径功能，例如 **"D:\\School\\PPT\\DNA Replication.pptx"**，其它系统的路径按照该路径书写，**必须为绝对路径**，并粘贴到程序中
# 设置方法
- 所有的配置均在 **config.json** 中，配置文件及解释如下
``` Json
{
    "double_page_printing": true, # 双面打印优化开关
    "line_num": 32, # 笔记区横线数量
    "line_color": [ # 笔记区横线颜色
        0,
        0,
        0
    ],
    "line_thickness": 2, # 笔记区横线高度
    "pages_in_page": 5, # PDF单页包含的PPT页数
    "margin": [ # PDF的边距，依次为上边距、下边距、左边距、右边距
        1.5,
        1.5,
        1.3,
        1.3
    ],
    "add_page_num": false # 页码显示开关(margin过小会被遮住)
}
```
# 注意事项
- ⚠ Python 脚本**理论上**可用于所有系统，但仅测试过 Windows
- 系统**必须已经安装 Office**，仅测试过 MS-Office
# 示例
![After](https://img2.imgtp.com/2024/03/14/l41TcQBa.png)
