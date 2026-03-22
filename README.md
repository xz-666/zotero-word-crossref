# Zotero Word 交叉引用宏

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![VBA](https://img.shields.io/badge/VBA-Word%20Macro-blue.svg)](https://docs.microsoft.com/en-us/office/vba/word/concepts/concepts-word-vba-reference)

一个强大的 Word VBA 宏工具，为 Zotero 插入的参考文献自动添加**交叉引用跳转链接**，实现正文中点击引用编号跳转到对应参考文献的功能。

## ✨ 功能特性

- 🔖 **一键创建书签** - 自动为参考文献列表创建书签
- 🔗 **添加跳转链接** - 为正文中所有 Zotero 引文添加超链接
- 🎯 **精准定位** - 支持单个引用 `[1]` 和多个引用 `[1,2,3]`
- 🛡️ **智能处理** - 解决最后一条引用跳转异常问题
- ⚡ **高效运行** - 自动批量处理整个文档

## 📋 使用前准备

1. Microsoft Word 2016 或更高版本
2. 已安装 [Zotero](https://www.zotero.org/) 和 Word 插件
3. 文档中已插入 Zotero 参考文献

## 🚀 使用方法

### 1. 安装宏

1. 打开你的 Word 文档
2. 按 `Alt + F11` 打开 VBA 编辑器
3. 在左侧项目窗口中，右键点击你的文档 → **插入** → **模块**
4. 将 `src/CrossRefMacro.vba` 中的代码全部复制粘贴到模块中
5. 按 `Ctrl + S` 保存（如提示，选择启用宏的格式 `.docm`）

### 2. 运行宏

1. **选中参考文献区域** - 从第一条到最后一条参考文献，手动选中整个区域
2. **运行宏** - 按 `Alt + F8`，选择 `RefLink_OneClick`，点击**运行**
3. **完成** - 宏会自动创建书签并为正文引用添加跳转链接

### 3. 使用跳转

运行完成后，按住 `Ctrl` 键并点击正文中的引用编号（如 `[1]`），即可跳转到对应的参考文献。

## 📁 项目结构

```
zotero-word-crossref/
├── src/
│   └── CrossRefMacro.vba      # 主宏代码文件
├── docs/
│   ├── INSTALL.md             # 详细安装教程
│   ├── TROUBLESHOOTING.md     # 常见问题排查
│   └── screenshots/           # 截图说明
├── examples/
│   └── sample-doc.docx        # 示例文档
├── LICENSE                    # MIT 许可证
└── README.md                  # 本文件
```

## ⚠️ 注意事项

1. **备份文档** - 运行宏前建议先备份原文档
2. **选中区域** - 运行前必须手动选中参考文献列表区域
3. **Zotero 刷新** - 如添加/删除引用，需要重新运行此宏更新链接
4. **兼容格式** - 支持 `[1]`、`[1,2]`、`[1-3]` 等常见引用格式

## 🔧 兼容性

- Microsoft Word 2016/2019/2021/Microsoft 365
- Zotero 6.x / 7.x
- Windows 10/11

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

本项目基于 [MIT License](LICENSE) 开源。

## 🙏 致谢

感谢 Zotero 团队提供的优秀文献管理工具。

---

如果这个项目对你有帮助，请给个 ⭐ Star！
