## 1. 程序简介

本程序是一个用于处理PDF文档的工具，主要功能包括：

1. 从PDF文档中提取标题、摘要和结论。
2. 将提取的内容保存到Excel表格中。
3. 使用百度翻译API将提取的内容翻译成中文。
4. 支持批量处理PDF文档。
5. 支持为已提取的文档添加笔记。
6. 支持根据提取的标题重命名PDF文件。

## 2. 目的

该程序旨在帮助用户高效地管理和处理学术论文PDF文档，提取关键信息并进行翻译，以便更好地进行阅读和研究。

## 3. 功能概述

### 3.1 提取标题、摘要和结论

- 从PDF文档中提取标题、摘要和结论部分的内容。
- 支持处理多页PDF文档，最多处理前5页以获取摘要，处理所有页面以获取结论。
- 提供多种方式读取标题，确保提取的准确性。

### 3.2 保存到Excel表格

- 将提取的内容保存到Excel表格中，方便进行后续管理和分析。
- 如果Excel表格不存在，会自动创建一个新的表格。

### 3.3 翻译功能

- 使用百度翻译API将提取的标题、摘要和结论翻译成中文。
- 程序中调用的是百度翻译的专业领域服务，选用的领域是电子领域。
- 如果翻译失败，相应的翻译字段将留空，以便后续进行二次翻译。

### 3.4 批量处理

- 支持一次性处理多个PDF文档，批量提取信息并保存到Excel表格中。
- 支持选择多个文件，自动找出其中的PDF文件进行处理。

### 3.5 添加笔记

- 支持为已提取的文档添加笔记，并将笔记保存到Excel表格中。
- 笔记功能键单个选择PDF，选中后不需要反复重新选择，除非用户自己切换文件记笔记。
- 笔记添加完成后自动清空输入框，方便下次输入。

### 3.6 重命名PDF文件

- 根据提取的标题重命名PDF文件，方便文件管理。
- 支持批量重命名，选择需要重命名的PDF文件进行处理。

### 3.7 API配置

- API凭据单独放在一个txt文件中，支持更改。
- 没有配置API时能正常运行，不进行翻译操作，提示API配置有误，但不影响正常的摘录。

## 4. 使用方法

### 4.1 环境配置

- 确保已安装Python 3.x。

- 安装所需的Python库：

  bash

  ```bash
  pip install PyMuPDF pandas requests openpyxl tk
  ```

### 4.2 配置API凭据

- 在程序目录下创建一个名为`api.txt`的文件，内容如下：

  none

  ```none
  <你的百度翻译API的AppID>
  <你的百度翻译API的SecretKey>
  ```

### 4.3 获取百度翻译API

1. 访问百度翻译开放平台：https://fanyi-api.baidu.com/

2. 注册并登录百度账号。

3. 创建新应用，获取AppID和SecretKey。

4. 将AppID和SecretKey填入`api.txt`文件中，格式如下：

   none

   ```none
   AppID
   SecretKey
   ```

### 4.4 运行程序

#### 4.4.1 使用Python脚本运行

- 运行`main.py`文件启动程序：

  bash

  ```bash
  python main.py
  ```

#### 4.4.2 打包成exe文件运行

- 初次使用可能会比较慢请耐心等待一会

- 安装PyInstaller：

  bash

  ```bash
  pip install pyinstaller
  ```

- 在命令行中导航到程序目录，然后运行以下命令将程序打包成exe文件：

  bash

  ```bash
  pyinstaller --onefile --add-data "api.txt;." --add-data "ReadPaperRecord.xlsx;." main.py
  ```

- 这将生成一个 `dist` 目录，里面包含打包好的可执行文件 `main.exe`。

#### 4.4.3 双击exe文件运行

- 请确保 `api.txt` 文件和 `ReadPaperRecord.xlsx` 文件放在与 `main.exe` 文件同一目录下。
- 双击运行 `dist` 目录中的 `main.exe` 文件，即可启动程序。

### 4.5 使用界面

- 程序启动后，会显示一个图形用户界面（GUI），包含以下按钮：
  - **选择笔记**：选择一个PDF文档，并在界面上方显示其标题。
  - **添加笔记**：为选中的PDF文档添加笔记，并保存到Excel表格中。
  - **批量摘录**：选择多个PDF文档，批量提取信息并保存到Excel表格中。
  - **改文件名**：选择多个PDF文档，根据提取的标题重命名文件。
  - **重新翻译**：选择多个PDF文档，批量翻译提取的内容并保存到Excel表格中。

### 4.6 操作步骤

#### 4.6.1 批量提取信息

1. 点击“批量摘录”按钮。
2. 在弹出的文件选择对话框中，选择要处理的PDF文件。
3. 程序会自动提取每个PDF文件的标题、摘要和结论，并尝试进行翻译。
4. 提取和翻译的结果会保存到程序目录下的 `ReadPaperRecord.xlsx` 文件中。

#### 4.6.2 添加笔记

1. 点击“选择笔记”按钮，选择一个已提取的PDF文件。
2. 在文本框中输入笔记内容。
3. 点击“添加笔记”按钮，笔记会保存到 `ReadPaperRecord.xlsx` 文件中。
4. 笔记添加完成后，输入框会自动清空，方便下次输入。

#### 4.6.3 重命名PDF文件

1. 点击“改文件名”按钮。
2. 在弹出的文件选择对话框中，选择要重命名的PDF文件。
3. 程序会根据提取的标题重命名文件。

#### 4.6.4 重新翻译

1. 点击“重新翻译”按钮。
2. 在弹出的文件选择对话框中，选择需要重新翻译的PDF文件。
3. 程序会重新翻译提取的内容，并更新Excel文件。

## 5. 错误处理

### 5.1 API凭据错误

**问题描述**：如果 `api.txt` 文件未找到或API凭据不正确，程序将无法正常翻译内容。

**解决方案**：

1. 确保 `api.txt` 文件存在，并放在与脚本同一目录下。
2. 检查 `api.txt` 文件内容是否正确，确保第一行为AppID，第二行为SecretKey。

### 5.2 文件处理错误

**问题描述**：文件路径不正确或文件权限不足可能导致文件处理失败。

**解决方案**：

1. 确保文件路径正确。可以在代码中打印路径以调试：

   python

   ```python
   print("Excel Path:", excel_path)
   ```

2. 确保程序有足够的权限创建和写入文件。如果打包成exe文件后运行有问题，尝试以管理员身份运行。

### 5.3 Excel文件错误

**问题描述**：Excel文件未生成或更新，可能是因为路径不正确或文件权限不足。

**解决方案**：

1. 确保路径正确，文件存在并且可以被正确读取和写入。

2. 在代码中添加检查，确保文件存在：

   python

   ```python
   if not os.path.exists(excel_path):
       with open(excel_path, 'w') as f:
           pass  # 创建空的Excel文件
   ```

### 5.4 PDF读取错误

**问题描述**：程序在读取PDF文件时可能会遇到错误，导致无法提取内容。

**解决方案**：

1. 确保PDF文件格式正确且未损坏。
2. 尝试使用其他PDF文件查看是否能够正常提取内容。

### 5.5 翻译API调用错误

**问题描述**：调用百度翻译API时可能会遇到网络问题或API限制，导致翻译失败。

**解决方案**：

1. 检查网络连接是否正常。
2. 检查API调用配额是否已用完，如果是，请等待配额恢复或申请更多配额。

## 6. 常见问题

### 6.1 为什么翻译结果为空？

**可能原因**：

1. API凭据不正确或 `api.txt` 文件未找到。
2. 网络连接问题导致API请求失败。

**解决方案**：

1. 检查 `api.txt` 文件内容是否正确。
2. 确保网络连接正常。

### 6.2 为什么无法提取摘要或结论？

**可能原因**：

1. PDF文件格式不标准，无法正确提取内容。
2. 摘要或结论部分可能使用了不同的标题或格式。

**解决方案**：

1. 检查PDF文件格式，确保文件内容可读。
2. 修改代码中的正则表达式以匹配不同的标题或格式。

### 6.3 如何重新翻译内容？

**操作步骤**：

1. 点击“重新翻译”按钮，选择需要重新翻译的PDF文件。
2. 程序会重新翻译提取的信息，并更新Excel文件。

### 6.4 为什么程序运行缓慢？

**可能原因**：

1. 处理的PDF文件较大或数量较多。
2. 网络连接速度较慢，影响翻译API的调用。

**解决方案**：

1. 尝试减少一次性处理的PDF文件数量。
2. 确保网络连接稳定。

### 6.5 为什么Excel文件没有更新？

**可能原因**：

1. Excel文件路径不正确。
2. 程序没有足够的权限写入文件。

**解决方案**：

1. 确保Excel文件路径正确且文件存在。
2. 以管理员身份运行程序。

## 7. 打包过程中遇到的问题及解决方法

### 7.1 打包指令

```shell
pyinstaller myapp.spec
```

### 7.2 打包过程中遇到的问题

在使用Anaconda和PyInstaller打包过程中，遇到了以下问题：

1. **打包后的程序无法创建Excel文件**：
   - 错误信息：`No module named 'openpyxl.cell._writer'`
   - 解决方案：显式指定更多的隐藏导入，并创建一个hook文件。
2. **权限问题**：
   - 错误信息：`PermissionError: [Errno 13] Permission denied`
   - 解决方案：以管理员身份运行命令提示符，并确保文件和目录的权限设置正确。

### 7.3 解决方法

1. **显式指定更多的隐藏导入**： 创建一个名为`hook-openpyxl.py`的文件，并添加以下内容：

   python

   ```python
   from PyInstaller.utils.hooks import collect_submodules
   
   hiddenimports = collect_submodules('openpyxl')
   ```

   然后在spec文件中指定hook文件路径：

   python

   ```python
   # myapp.spec
   # -*- mode: python ; coding: utf-8 -*-
   
   import os
   
   block_cipher = None
   
   project_dir = '/'
   
   a = Analysis(
       [os.path.join(project_dir, 'main.py')],
       pathex=[project_dir],
       binaries=[],
       datas=[(os.path.join(project_dir, 'api.txt'), 'api.txt')],
       hiddenimports=[],
       hookspath=[os.path.join(project_dir, 'hooks')],
       runtime_hooks=[],
       excludes=[],
       win_no_prefer_redirects=False,
       win_private_assemblies=False,
       cipher=block_cipher,
       noarchive=False,
   )
   pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
   
   exe = EXE(
       pyz,
       a.scripts,
       [],
       exclude_binaries=True,
       name='main',
       debug=False,
       bootloader_ignore_signals=False,
       strip=False,
       upx=True,
       console=False,  # 设置为False，不显示控制台窗口
   )
   coll = COLLECT(
       exe,
       a.binaries,
       a.zipfiles,
       a.datas,
       strip=False,
       upx=True,
       upx_exclude=[],
       name='main',
   )
   ```

2. **以管理员身份运行命令提示符**： 在Windows上，搜索“命令提示符”，右键点击，然后选择“以管理员身份运行”。

3. **确保文件和目录的权限设置正确**： 右键点击项目目录，选择“属性”，在“安全”选项卡中检查并设置权限。

### 7.4 吐槽

虽然Python在开发过程中非常方便，但打包后的程序相当大，这也是Python打包的一大缺点。不推荐在对文件大小有严格要求的项目中使用Python进行打包。

### 7.5 使用Anaconda打包

在使用Anaconda和PyInstaller打包过程中，遇到了打包后无法创建Excel文件的问题。通过添加spec文件和hook-openpyxl.py文件，并以管理员身份运行命令提示符，解决了创建Excel文件的问题。这主要是由于openpyxl库的依赖项没有很好地加载导致的。

### 7.6 项目结构

整个工程的主要结构如下（exe打开的用户无需考虑）：

```none
├── main.py
├── myapp.spec
├── api.txt
└── hooks\
    └── hook-openpyxl.py
```

通过这些步骤，你应该能够成功使用Anaconda和PyInstaller将你的Python脚本打包成一个可执行文件，并确保它能够正常运行和创建Excel文件。**打包成的exe文件会在dist\main文件夹内**，项目仓库有可能会附带这个exe文件。

