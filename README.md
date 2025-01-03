## 概述

--------

本工具用于格式化Excel文件，包括冻结首行、首行筛选、设置首行字体和颜色、居中对齐等功能。

之所以有这个需求，是因为在日常工作中，在向客户提交具体文档之前进行格式化处理不仅能提升可读性，更能展现专业形象（尽管有时候是表面功夫啦），再加上强迫症作祟，总是忍不住想把文档打磨得更~~完美~~（符合我的期待，期待未必是正确）一些。

## 命令行版本

--------------

### 特点

*彩色的命令行输出
*双进度条（总体进度和当前工作表进度）
*详细的步骤提示
*时间戳
*预计剩余时间
*完成百分比

### 代码

```python
见Github...
```

### 使用说明

1.安装必要的库：`pip install openpyxl tqdm colorama`
2.将代码保存为 `.py` 文件并运行
3.输入Excel文件路径即可开始处理

![image-20250103160644808](https://favorably-7690.oss-cn-beijing.aliyuncs.com/%E8%87%AA%E5%8A%A8%E5%8C%96%20-%20Excel%E6%A0%BC%E5%BC%8F%E5%8C%96%E5%B7%A5%E5%85%B7/202501031639054.png)

![image-20250103160655589](https://favorably-7690.oss-cn-beijing.aliyuncs.com/%E8%87%AA%E5%8A%A8%E5%8C%96%20-%20Excel%E6%A0%BC%E5%BC%8F%E5%8C%96%E5%B7%A5%E5%85%B7/202501031639282.png)

![image-20250103161321283](https://favorably-7690.oss-cn-beijing.aliyuncs.com/%E8%87%AA%E5%8A%A8%E5%8C%96%20-%20Excel%E6%A0%BC%E5%BC%8F%E5%8C%96%E5%B7%A5%E5%85%B7/202501031639250.png)

## 图形界面版本

----------------

### 特点

*图形界面，包含文件选择按钮、总体进度条、当前工作表进度条、详细步骤显示、预计剩余时间和开始/取消按钮
*实现了所有格式要求
*进度显示：实时显示处理步骤、显示预计剩余时间、可以取消处理过程

### 代码

```python
见Github...
```

### 使用说明

1.安装必要的库：`pip install openpyxl`
2.将代码保存为 `.py` 文件并运行
3.选择Excel文件即可开始处理

![image-20250103161430306](https://favorably-7690.oss-cn-beijing.aliyuncs.com/%E8%87%AA%E5%8A%A8%E5%8C%96%20-%20Excel%E6%A0%BC%E5%BC%8F%E5%8C%96%E5%B7%A5%E5%85%B7/202501031639078.png)
