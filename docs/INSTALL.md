# Installation Guide / 安装教程

[English](#english) | [中文](#chinese)

---

<a name="english"></a>
## 🇺🇸 English

## Method 1: Install via VBA Editor (Recommended)

### Step 1: Open VBA Editor

1. Open your Word document
2. Press `Alt + F11` to open the VBA editor

### Step 2: Insert Module

1. In the Project Explorer window on the left, find your document (e.g., `Project (YourDocument)`)
2. Right-click your document → **Insert** → **Module**

### Step 3: Paste Code

1. Open the `src/CrossRefMacro.vba` file
2. Copy all code
3. Paste it into the code window on the right side of the VBA editor

### Step 4: Save

1. Press `Ctrl + S` to save
2. If prompted that saving is not possible, select **No**, then save the document as `.docm` format (Macro-enabled Word Document)

### Step 5: Run Macro

1. Return to the Word document
2. Select the bibliography area (from first to last reference)
3. Press `Alt + F8` to open the Macro dialog
4. Select `RefLink_OneClick`, click **Run**

## Method 2: Install via Normal Template (Global)

If you want to use this macro in all Word documents, install it to the Normal template:

1. Press `Alt + F11` to open the VBA editor
2. Find `Normal` (Global template) on the left
3. Right-click `Normal` → **Insert** → **Module**
4. Paste the code and save
5. Restart Word, the macro will be available in all documents

## Verify Installation

After running the macro, you should see:

1. Gray brackets appear on the left side of each bibliography entry (indicating bookmarks created)
2. Citation numbers in the main text turn blue with underlines (indicating hyperlinks added)
3. A success message box showing the number of bookmarks and links created

## Next Step

- View [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for common issues and solutions

---

<a name="chinese"></a>
## 🇨🇳 中文

## 方法一：通过 VBA 编辑器安装（推荐）

### 步骤 1：打开 VBA 编辑器

1. 打开你的 Word 文档
2. 按 `Alt + F11` 打开 VBA 编辑器

### 步骤 2：插入模块

1. 在左侧的**工程资源管理器**窗口中，找到你的文档（如 `Project (你的文档名)`）
2. 右键点击你的文档 → **插入** → **模块**

### 步骤 3：粘贴代码

1. 打开 `src/CrossRefMacro.vba` 文件
2. 复制全部代码
3. 粘贴到 VBA 编辑器右侧的代码窗口中

### 步骤 4：保存

1. 按 `Ctrl + S` 保存
2. 如果提示无法保存，选择**否**，然后将文档另存为 `.docm` 格式（启用宏的 Word 文档）

### 步骤 5：运行宏

1. 回到 Word 文档
2. 选中参考文献区域（从第一条到最后一条）
3. 按 `Alt + F8` 打开宏对话框
4. 选择 `RefLink_OneClick`，点击**运行**

## 方法二：通过 Normal 模板安装（全局可用）

如果你想在所有 Word 文档中使用此宏，可以将其安装到 Normal 模板：

1. 按 `Alt + F11` 打开 VBA 编辑器
2. 在左侧找到 `Normal`（通用模板）
3. 右键点击 `Normal` → **插入** → **模块**
4. 粘贴代码并保存
5. 重启 Word，宏将在所有文档中可用

## 验证安装

运行宏后，你应该看到：

1. 参考文献列表区域每个条目左侧出现灰色方括号（表示书签已创建）
2. 正文中的引用编号变成蓝色并带下划线（表示超链接已添加）
3. 弹出成功提示框，显示创建的书签数和链接数

## 下一步

- 查看 [TROUBLESHOOTING.md](TROUBLESHOOTING.md) 了解常见问题及解决方法
