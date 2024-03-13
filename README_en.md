# Sidebar Notes PPT to PDF ([简体中文](README.md) | English)

❤️ Convert PPT to PDF with sidebar notes area, high customization![](https://img.shields.io/badge/Fast-Simple-Simple)

# Usage
- Double-click **install.bat** to install
- Double-click **run.bat** to start the program
- When you need to input the path, use the **Windows** file path copy function, for example **"D:\School\PPT\DNA Replication.pptx"**. For other systems, the path should be written in a similar way, **it must be an absolute path**, and paste it into the program

# Settings
- All configurations are in config.json. The configuration file and its explanations are as follows:
```Json
{  
    "double_page_printing": true, # Switch for double-sided printing optimization  
    "line_num": 32, # Number of horizontal lines in the notes area  
    "line_color": [ # Color of the horizontal lines in the notes area  
        0,  
        0,  
        0  
    ],  
    "line_thickness": 2, # Height of the horizontal lines in the notes area  
    "pages_in_page": 5, # Number of PPT pages contained in a single PDF page  
    "margin": [ # Margins of the PDF, in order of top, bottom, left, and right  
        1.5,  
        1.5,  
        1.3,  
        1.3  
    ],  
    "add_page_num": false # Switch for displaying page numbers (they may be covered if the margin is too small)  
}
```

# Precautions
- ⚠ The Python script **in theory** can be used on all systems, but it has only been tested on Windows
- The system **must have Office installed**, and only MS-Office has been tested

# Example
![After](https://img2.imgtp.com/2024/03/14/l41TcQBa.png)