# Feishu Auto-Complete Student Info

该程序使用 Selenium 自动在 Feishu 中搜索学生姓名并补全其学号和学院信息，最后将结果保存为 Excel 文件。程序自动打开 Feishu 网页，扫描二维码登录，输入学生姓名并提取相关信息，最终输出一个新的 Excel 文件。

## 功能

- **自动登录**：扫描二维码登录 Feishu。
- **自动搜索**：从 Excel 文件中读取学生姓名并搜索相关信息（学号、学院）。
- **自动补全**：将学生的学号和学院信息自动补全到新 Excel 文件。
- **输出 Excel**：保存结果到新的 Excel 文件中。

## 使用要求

1. **安装 Python**：确保已经安装了 Python 3.13 版本。

2. **安装必要的库**：安装以下 Python 库：

	- `selenium`：用于浏览器自动化。
	- `pandas`：用于数据处理和 Excel 文件读取。

	使用以下命令安装所需库：

	```bash
	pip install selenium pandas
	```

3. **Chrome 浏览器和 ChromeDriver**：

	- 安装 Google Chrome 浏览器。
	- 下载与您 Chrome 版本匹配的 [ChromeDriver](https://sites.google.com/chromium.org/driver/downloads)（115之后的版本在 [Chrome For Testing](https://googlechromelabs.github.io/chrome-for-testing/)）

## 使用步骤

1. **配置文件路径**： 在代码中，你需要设置以下路径：
	- `google_chorme_address`：设置为你本地 ChromeDriver 的路径。
	- `feishu_url`：Feishu 网页链接，默认为 `https://nankai.feishu.cn/next/messenger/`。
	- `original_file_path`：输入的 Excel 文件路径，包含学生姓名列。
	- `new_file_path`：输出的新 Excel 文件路径，包含学生姓名、学号和学院。
2. **运行程序**： 执行程序 `python main.py`，程序将自动完成以下操作：
	- 打开 Feishu 网页并等待扫码登录。
	- 搜索 Excel 文件中的学生姓名，提取学号和学院信息。
	- 将结果保存到新的 Excel 文件中。
3. **查看结果**： 程序完成后，生成的新 Excel 文件将包含学生姓名、学号和学院。

## 代码结构

- **main.py**：主程序文件，包含 Feishu 自动登录、信息搜索和结果输出的功能。
- **old.xlsx**：包含学生姓名信息的输入文件（格式：姓名列）。
- **new.xlsx**：输出的 Excel 文件（格式：姓名、学号、学院列）。

## 错误处理

程序会捕获可能出现的错误并打印相关信息。如果出现登录超时或无法找到学生信息，程序会打印错误并跳过该学生。

## 示例

执行以下命令运行程序：

```bash
python main.py
```

### 预期输出：

```bash
Page opened https://nankai.feishu.cn/next/messenger/
Please scan with your phone...
Logged in successfully!
wait to located...
body window located!
pressed ctrl+k
wait to located...
Search box located!
Entered: 张三
('张三', '123456789', '计算机学院') done
...
Total done!!!
```

程序将生成一个新的 Excel 文件 `new.xlsx`，并保存学生信息。