# Troubleshooting / 常见问题排查

[English](#english) | [中文](#chinese)

---

<a name="english"></a>
## 🇺🇸 English

## Issue 1: "Please manually select the bibliography area first"

**Cause**: You ran the macro without selecting the bibliography list area

**Solution**:
1. Find the bibliography list in your document
2. Drag from the first to the last reference with your mouse to select the entire area
3. Run the macro again

---

## Issue 2: Last citation jumps to [1] instead of the last item

**Cause**: Word bookmark cache ID conflict

**Solution**:
- This macro has special handling for this issue (`Last_Ref_` marker). If problems persist:
1. Delete all bibliography bookmarks (manually or use cleanup macro)
2. Re-run this macro

---

## Issue 3: Some citations have no links added

**Possible causes**:
1. The citation doesn't exist in the bibliography list
2. Citation format doesn't match expectations (e.g., using Chinese brackets【】)
3. The citation was manually entered instead of inserted by Zotero

**Solution**:
1. Ensure all citations are inserted via Zotero
2. Check if the bibliography list has the corresponding number
3. Manually verify if the citation brackets are English `[]`

---

## Issue 4: Jump position is inaccurate

**Possible causes**:
1. Page breaks or section breaks interfere with the document
2. Bibliography uses special paragraph formatting

**Solution**:
1. Try removing page breaks, use paragraph pagination settings
2. Ensure each bibliography entry is an independent paragraph

---

## Issue 5: Zotero refresh causes links to disappear after macro runs

**Cause**: Zotero refreshes citation fields and overwrites existing hyperlinks

**Solution**:
1. Complete all citation additions and modifications in Zotero first
2. Then run this macro to add jump links
3. Re-run this macro after modifying citations

---

## Issue 6: "No reference numbers recognized"

**Possible causes**:
1. Bibliography numbers use Chinese brackets `【】` instead of English `[]`
2. Number format is not pure numeric (e.g., [1a])
3. Selected area doesn't contain numbers

**Solution**:
1. Ensure numbers use English brackets `[number]`
2. Check if the selected area is correct

---

## Issue 7: "Cannot save because file already exists"

**Cause**: Document is occupied by another program

**Solution**:
1. Close the document and reopen
2. Use "Save As" to save as a new `.docm` file

---

## Issue 8: Macro is disabled and cannot run

**Cause**: Word security settings block macro execution

**Solution**:

### Method 1: Enable Macros
1. Click **File** → **Options** → **Trust Center** → **Trust Center Settings**
2. Select **Macro Settings** → **Enable all macros** (not recommended for long-term use)

### Method 2: Add Trusted Location
1. Click **File** → **Options** → **Trust Center** → **Trust Center Settings**
2. Select **Trusted Locations** → **Add new location**
3. Add the folder where your document is located

---

## Issue 9: Word becomes unresponsive after running macro

**Possible cause**: Document too large or too many citations

**Solution**:
1. Be patient, the macro may take a few seconds when processing many citations
2. Close other programs to free up memory
3. Process in batches: handle some sections first, then others

---

## Other Issues

If the above methods cannot solve your problem, please:

1. Submit an [Issue](https://github.com/xz-666/zotero-word-crossref/issues) on GitHub
2. Include the following in your description:
   - Word version
   - Zotero version
   - Operating system version
   - Detailed description of the issue
   - Screenshots of error messages (if any)

---

<a name="chinese"></a>
## 🇨🇳 中文

## 问题 1：运行宏时提示"请先手动选中参考文献区域"

**原因**：你没有选中参考文献列表区域就运行了宏

**解决方法**：
1. 找到文档中的参考文献列表
2. 用鼠标从第一条文献拖动到最后一条，选中整个区域
3. 再次运行宏

---

## 问题 2：最后一条引用跳转到 [1] 而不是最后一条

**原因**：Word 书签缓存 ID 冲突

**解决方法**：
- 本宏已对此问题做特殊处理（`Last_Ref_` 标记），如果仍有问题：
1. 删除所有参考文献书签（手动或使用清理宏）
2. 重新运行本宏

---

## 问题 3：部分引用没有添加链接

**可能原因**：
1. 该引用在参考文献列表中不存在
2. 引用格式不符合预期（如使用中文括号【】）
3. 该引用是手动输入而非 Zotero 插入

**解决方法**：
1. 确保所有引用都通过 Zotero 插入
2. 检查参考文献列表中是否有对应的编号
3. 手动检查引用的括号是否为英文 `[]`

---

## 问题 4：点击链接跳转位置不准确

**可能原因**：
1. 文档中有分页符或分节符干扰
2. 参考文献使用了特殊的段落格式

**解决方法**：
1. 尝试删除分页符，使用段落的分页设置
2. 确保每条参考文献是一个独立段落

---

## 问题 5：宏运行后 Zotero 刷新导致链接丢失

**原因**：Zotero 刷新引文域时会覆盖原有的超链接

**解决方法**：
1. 在 Zotero 中完成所有引用的添加和修改
2. 然后运行本宏添加跳转链接
3. 如需修改引用，修改后重新运行本宏

---

## 问题 6：提示"未识别到任何参考文献编号"

**可能原因**：
1. 参考文献编号使用了中文括号 `【】` 而非英文 `[]`
2. 编号格式不是纯数字（如 [1a]）
3. 选中的区域不包含编号

**解决方法**：
1. 确保编号使用英文方括号 `[数字]`
2. 检查选中的区域是否正确

---

## 问题 7：保存时提示"无法保存，因为文件已存在"

**原因**：文档已被其他程序占用

**解决方法**：
1. 关闭文档，重新打开
2. 使用"另存为"保存为新的 `.docm` 文件

---

## 问题 8：宏被禁用，无法运行

**原因**：Word 安全设置阻止了宏运行

**解决方法**：

### 方法 1：启用宏
1. 点击**文件** → **选项** → **信任中心** → **信任中心设置**
2. 选择**宏设置** → **启用所有宏**（不推荐长期使用）

### 方法 2：添加受信任位置
1. 点击**文件** → **选项** → **信任中心** → **信任中心设置**
2. 选择**受信任位置** → **添加新位置**
3. 将你的文档所在文件夹添加进去

---

## 问题 9：运行宏后 Word 卡顿或无响应

**可能原因**：文档过大或引文数量过多

**解决方法**：
1. 耐心等待，宏在处理大量引文时可能需要几秒钟
2. 关闭其他程序释放内存
3. 分批处理：先处理部分章节，再处理其他部分

---

## 其他问题

如果以上方法无法解决你的问题，请：

1. 在 GitHub 上提交 [Issue](https://github.com/xz-666/zotero-word-crossref/issues)
2. 描述问题时请包含：
   - Word 版本
   - Zotero 版本
   - 操作系统版本
   - 问题的详细描述
   - 错误提示截图（如有）
