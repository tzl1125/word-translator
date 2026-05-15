# word-translator
这是一个基于大语言模型的word翻译项目，支持一键多语言翻译，同时最大程度保持word格式不变（包括复杂表格、页眉页脚、EndNote参考文献引用、数学公式、图表标题在内的所有原始格式和元数据）

![Word文档智能翻译](pictures/demo.png)

## 🌟 核心优势

- **✅ 保留文档原始格式**: 字体、字号、上下标、段落、表格、页眉页脚、数学公式等复杂排版在翻译后均能还原。
- **✅ EndNote引用保留**: 完整保留EndNote参考文献的引用标记，学术用户的福音！
- **✅ 自定义翻译风格**: 通过提示词（Prompt）自由控制翻译风格，如“学术严谨”、“简洁口语化”、“法律术语”等。
- **✅ 上下文感知翻译**: 基于大语言模型（LLM），语义连贯性远超传统机器翻译。
- **✅ 极简操作**: 只需上传 `.docx` 文件，一键即可获得高质量、可直接使用的译文文档。

## 🚀 在线体验

该项目已经部署上线，欢迎随时体验！
👉 **[https://www.zlaixy.top/](https://www.zlaixy.top/)**

## 📦 本地部署

### 环境要求

- Python 3.8+

### 安装步骤

1. **克隆仓库**
   ```bash
   git clone https://github.com/tzl1125/word-translator.git
   cd word-translator
   ```

2. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

3. **准备文件**
   - 将您要翻译的 `.docx` 文件放入 `source_files/` 目录。

4. **配置参数**
   打开 `main.py` 文件，修改顶部的配置参数：
   ```python
   INPUT_FILE = "source_files/your_document.docx"  # 您的文档路径
   API_KEY = "your-api-key-here"                 # 您的大模型API密钥
   # ... 其他参数
   ```

5. **运行翻译**
   ```bash
   python main.py
   ```
   翻译完成的文件将出现在 `translated_files/` 目录下。

## 🔧 自定义大模型

本项目采用模块化设计，您可以轻松更换底层的大语言模型。

1. 打开 `src/llm_interface.py` 文件。
2. 修改 `call_llm_api_async` 函数，接入您选择的模型（如 OpenAI GPT, Claude, 本地模型等）。
3. 确保新函数的输入输出签名与原函数保持一致。

## 📂 项目结构

```
word-docx-translator/
├── main.py              # 项目入口
├── README.md            # 本说明文件
├── requirements.txt     # 依赖列表
├── src/                 # 核心源码
│   ├── core.py          # 数据类与公共函数
│   ├── extractor.py     # 文档提取器
│   ├── translator.py    # 翻译控制器
│   ├── injector.py      # 结果注入器
│   └── llm_interface.py # 大模型接口 (可替换)
├── pictures/            # 展示图片
└── source_files/        # 原文档
└── translated_files/    # 翻译后的文档
```

## ☕ 友情赞助

如果您觉得这个项目对您有帮助，欢迎扫码支持！您的赞助将激励我持续维护和开发更多实用功能。

<img src="pictures/wechat_donation_qr.png" alt="赞助二维码" width="250" />
