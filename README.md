# BOM_Organized

使用python脚本对公司使用的标准BOM表进行整理
---
**主要功能：** 将xlsx格式的BOM文件以"元件值", "封装" 和 "精度" 这三列为参考依据进行操作,这三列都相同的视为同一元器件，将同一元器件的位号合并到一行中, 并重新统计数量。

**使用方法：** 运行脚本,弹出文件选择对话框, 选择需要整理的xlsx格式的BOM表, 整理好后会在原目录生成一个“原文件名+时间戳”的新文件。