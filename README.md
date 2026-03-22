# Zotero Word Cross-Reference Macro / Zotero Word 交叉引用宏

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![VBA](https://img.shields.io/badge/VBA-Word%20Macro-blue.svg)](https://docs.microsoft.com/en-us/office/vba/word/concepts/concepts-word-vba-reference)

[English](#english) | [中文](#chinese)

---

<a name="english"></a>
## 🇺🇸 English

A powerful Word VBA macro tool that automatically **adds cross-reference jump links** to Zotero-inserted references, enabling click-to-jump functionality from citations in the main text to the corresponding bibliography entries.

### ✨ Features

- 🔖 **One-click bookmark creation** - Automatically create bookmarks for bibliography entries
- 🔗 **Add jump links** - Add hyperlinks to all Zotero citations in the main text
- 🎯 **Precise positioning** - Support single citation `[1]` and multiple citations `[1,2,3]`
- 🛡️ **Smart handling** - Special handling for the last reference jump issue
- ⚡ **Efficient** - Automatically process the entire document in batch

### 📋 Prerequisites

1. Microsoft Word 2016 or later
2. [Zotero](https://www.zotero.org/) with Word plugin installed
3. References already inserted in the document

### 🚀 Installation

1. Open your Word document
2. Press `Alt + F11` to open the VBA editor
3. Right-click your document in the Project window → **Insert** → **Module**
4. Copy all code from `src/CrossRefMacro.vba` into the module
5. Press `Ctrl + S` to save (choose `.docm` format if prompted)

### 💡 Usage

1. **Select the bibliography area** - Manually select from the first to the last reference
2. **Run the macro** - Press `Alt + F8`, select `RefLink_OneClick`, click **Run**
3. **Done** - The macro will create bookmarks and add jump links to citations

After completion, hold `Ctrl` and click citation numbers (e.g., `[1]`) in the main text to jump to the corresponding reference.

---

<a name="chinese"></a>
## 🇨🇳 中文

一个强大的 Word VBA 宏工具，为 Zotero 插入的参考文献自动添加**交叉引用跳转链接**，实现正文中点击引用编号跳转到对应参考文献的功能。

### ✨ 功能特性

- 🔖 **一键创建书签** - 自动为参考文献列表创建书签
- 🔗 **添加跳转链接** - 为正文中所有 Zotero 引文添加超链接
- 🎯 **精准定位** - 支持单个引用 `[1]` 和多个引用 `[1,2,3]`
- 🛡️ **智能处理** - 解决最后一条引用跳转异常问题
- ⚡ **高效运行** - 自动批量处理整个文档

### 📋 使用前准备

1. Microsoft Word 2016 或更高版本
2. 已安装 [Zotero](https://www.zotero.org/) 和 Word 插件
3. 文档中已插入 Zotero 参考文献

### 🚀 安装方法

1. 打开你的 Word 文档
2. 按 `Alt + F11` 打开 VBA 编辑器
3. 在左侧项目窗口中，右键点击你的文档 → **插入** → **模块**
4. 将 `src/CrossRefMacro.vba` 中的代码全部复制粘贴到模块中
5. 按 `Ctrl + S` 保存（如提示，选择启用宏的格式 `.docm`）

### 💡 使用方法

1. **选中参考文献区域** - 从第一条到最后一条参考文献，手动选中整个区域
2. **运行宏** - 按 `Alt + F8`，选择 `RefLink_OneClick`，点击**运行**
3. **完成** - 宏会自动创建书签并为正文引用添加跳转链接

运行完成后，按住 `Ctrl` 键并点击正文中的引用编号（如 `[1]`），即可跳转到对应的参考文献。

---

## 📁 Project Structure / 项目结构

```
zotero-word-crossref/
├── src/
│   └── CrossRefMacro.vba      # Main macro code / 主宏代码
├── docs/
│   ├── INSTALL.md             # Installation guide / 安装教程
│   ├── TROUBLESHOOTING.md     # Troubleshooting / 常见问题
│   └── CHANGELOG.md           # Changelog / 更新日志
├── examples/                  # Example files / 示例文件
├── LICENSE                    # MIT License / MIT 许可证
└── README.md                  # This file / 本文件
```

## ⚠️ Important Notes / 注意事项

**EN:**
1. **Backup your document** before running the macro
2. **Select the bibliography area** before running
3. **Re-run after Zotero refresh** if citations are added/removed
4. **Supported formats** - `[1]`, `[1,2]`, `[1-3]` etc.

**CN:**
1. **备份文档** - 运行宏前建议先备份原文档
2. **选中区域** - 运行前必须手动选中参考文献列表区域
3. **Zotero 刷新** - 如添加/删除引用，需要重新运行此宏更新链接
4. **兼容格式** - 支持 `[1]`、`[1,2]`、`[1-3]` 等常见引用格式

## 🔧 Compatibility / 兼容性

- Microsoft Word 2016/2019/2021/Microsoft 365
- Zotero 6.x / 7.x
- Windows 10/11

## 📄 License / 许可证

[MIT License](LICENSE)

Copyright (c) 2026 xz-666

---

If this project helps you, please give it a ⭐ Star!

如果这个项目对你有帮助，请给个 ⭐ Star！
