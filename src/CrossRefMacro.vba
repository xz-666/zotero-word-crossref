' ============================================================================
' Zotero Word Cross-Reference Macro / Zotero Word 交叉引用宏
' ============================================================================
' Description / 描述:
'   Automatically create bookmarks for Zotero references and add jump links
'   to citations in the main text.
'   自动为 Zotero 参考文献创建书签，并为正文引文添加跳转链接。
'
' Version / 版本: 1.0.0
' Author / 作者: xz-666
' Date / 日期: 2026-03-23
' License / 许可证: MIT
' ============================================================================

' ============================================================================
' Main Function: One-click bookmark creation + jump link addition
' 主函数：一键创建书签+添加跳转链接
' ============================================================================
Sub RefLink_OneClick()
    Application.ScreenUpdating = False
    On Error Resume Next

    ' ===================== Step 1: Create bookmarks (with special handling for last item) =====================
    ' ===================== 第一步：创建参考文献书签（含最后一条特殊处理） =====================
    Dim step1_success As Boolean: step1_success = False

    ' 1. Check if bibliography area is selected / 检查是否手动选中了参考文献区域
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "❌ Please manually select your bibliography area (from first to last), then run this code!" & vbCrLf & _
               "❌ 请先手动选中你的参考文献区域（从第一条到最后一条），再运行此代码！", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 2. Clear old bookmarks (avoid duplicates/cache issues) / 清空旧的ZotRef书签（避免重复/缓存干扰）
    Dim bk As Bookmark
    For Each bk In ActiveDocument.Bookmarks
        If Left(bk.Name, 6) = "ZotRef_" Or Left(bk.Name, 8) = "Last_Ref_" Then
            bk.Delete
        End If
    Next bk

    ' 3. Traverse selected area, extract [numbers] and create bookmarks / 自动遍历选中区域，提取所有[数字]编号并创建普通书签
    Dim oPara As Paragraph, paraText As String, refNum As String
    Dim successCount As Integer: successCount = 0
    Dim foundNums As String: foundNums = ""
    Dim i As Integer, inBracket As Boolean: inBracket = False
    Dim lastPara As Paragraph, lastRefNum As String: lastRefNum = ""

    For Each oPara In Selection.Paragraphs
        paraText = Trim(oPara.Range.Text)
        refNum = ""
        inBracket = False

        ' Strictly extract pure numbers between [ ] (supports multi-digit like 100) / 严格提取[ ]之间的纯数字（兼容多位数，如100）
        For i = 1 To Len(paraText)
            Dim c As String: c = Mid(paraText, i, 1)
            If c = "[" Then
                inBracket = True
            ElseIf c = "]" Then
                inBracket = False
                If refNum <> "" Then Exit For ' Stop after extracting number / 提取到编号后立即停止
            ElseIf inBracket And IsNumeric(c) Then
                refNum = refNum & c ' Only concatenate numbers inside brackets / 只拼接括号内的数字
            End If
        Next i

        ' Create bookmark for valid numbers / 为有效编号创建普通书签
        If refNum <> "" Then
            ActiveDocument.Bookmarks.Add _
                Name:="ZotRef_" & refNum, _
                Range:=oPara.Range
            successCount = successCount + 1
            foundNums = foundNums & refNum & ","
            ' Update last item info in real-time / 实时更新最后一条的信息
            Set lastPara = oPara
            lastRefNum = refNum
        End If
    Next oPara

    ' 4. Force special handling for the last bibliography entry (fixes jump to [1] issue) / 对最后一条参考文献做【强制特殊处理】（解决跳转[1]的问题）
    If lastRefNum <> "" Then
        ' 4.1 Delete old bookmark of last item (avoid cache) / 删除最后一条的旧书签（避免缓存）
        ActiveDocument.Bookmarks("ZotRef_" & lastRefNum).Delete

        ' 4.2 Precisely locate [number] text of last item / 精准定位最后一条的[编号]文字
        Dim lastRange As Range
        Set lastRange = lastPara.Range
        Dim bracketStart As Integer: bracketStart = InStr(lastRange.Text, "[")
        Dim bracketEnd As Integer: bracketEnd = InStr(lastRange.Text, "]")

        If bracketStart > 0 And bracketEnd > 0 Then
            ' Precisely select [number], convert to document global position / 精准选中[编号]，转换为文档全局位置
            lastRange.SetRange _
                Start:=lastPara.Range.Start + bracketStart - 1, _
                End:=lastPara.Range.Start + bracketEnd
        End If

        ' 4.3 Force bind [number] text bookmark (avoid page break interference) / 强制绑定[编号]文字书签（避免分页符干扰）
        ActiveDocument.Bookmarks.Add _
            Name:="ZotRef_" & lastRefNum, _
            Range:=lastRange

        ' 4.4 Create hard-coded marker (bypass Word ID cache conflict) / 创建硬编码标记（绕开Word ID缓存冲突）
        ActiveDocument.Bookmarks.Add _
            Name:="Last_Ref_" & lastRefNum, _
            Range:=lastPara.Range
    End If

    If successCount > 0 Then
        step1_success = True
    Else
        MsgBox "❌ No reference numbers recognized!" & vbCrLf & _
               "❌ 未识别到任何参考文献编号！", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' ===================== Step 2: Add jump links to main text citations =====================
    ' ===================== 第二步：为正文引文添加跳转链接 =====================
    Dim linkCount As Integer: linkCount = 0
    Dim oField As Field, fieldRng As Range, fieldText As String
    Dim bracketPos As Long, bracketEndPos As Long, contentInBracket As String

    ' Traverse all Zotero citation fields (only process ZOTERO_ITEM type) / 遍历所有Zotero引文域（仅处理ZOTERO_ITEM类型）
    For Each oField In ActiveDocument.Fields
        If InStr(oField.Code.Text, "ZOTERO_ITEM") > 0 Then
            Set fieldRng = oField.Result
            fieldText = Trim(fieldRng.Text)

            ' Only process citations with [] format / 只处理带[]的引文格式
            bracketPos = InStr(fieldText, "[")
            bracketEndPos = InStr(fieldText, "]")
            If bracketPos > 0 And bracketEndPos > bracketPos Then
                contentInBracket = Mid(fieldText, bracketPos + 1, bracketEndPos - bracketPos - 1)
                ' Parse numbers and add precise jumps / 解析数字并添加精准跳转
                linkCount = linkCount + AddLinksToContent(fieldRng, contentInBracket, bracketPos)
            End If
        End If
    Next oField

    ' ===================== Completion message =====================
    ' ===================== 完成提示 =====================
    Application.ScreenUpdating = True
    MsgBox "✅ Processing complete! / 一键处理完成！" & vbCrLf & _
           "📌 Bookmarks created: / 成功创建书签：" & successCount & _
           IIf(successCount > 1, " items", " item") & " / 个" & vbCrLf & _
           "🔗 Jump links added: / 成功添加跳转链接：" & linkCount & _
           IIf(linkCount > 1, " items", " item") & " / 个" & _
           IIf(lastRefNum <> "", vbCrLf & "🔧 Special handling applied to last item (" & lastRefNum & ") / 已对最后一条（" & lastRefNum & "）做特殊处理。", "")
End Sub

' ============================================================================
' Helper Function: Parse numbers and add jumps / 辅助子过程：解析数字并添加跳转
' ============================================================================
Function AddLinksToContent(fieldRng As Range, content As String, bracketStart As Long) As Integer
    Dim i As Integer, char As String
    Dim currentNum As String, linkCount As Integer: linkCount = 0
    Dim numStartInField As Integer ' Starting position of number in field text / 数字在域文本中的「内部起始位置」
    Dim numRng As Range ' Declare once, avoid repetition / 只声明一次，避免重复

    For i = 1 To Len(content)
        char = Mid(content, i, 1)

        ' Encounter digit: concatenate complete number, record field-internal start position / 遇到数字：拼接完整数字，记录域内起始位置
        If IsNumeric(char) Then
            If currentNum = "" Then
                numStartInField = bracketStart + i
            End If
            currentNum = currentNum & char

        ' Encounter non-digit: process concatenated complete number / 遇到非数字：处理已拼接的完整数字
        ElseIf currentNum <> "" Then
            If numStartInField + Len(currentNum) - 1 <= fieldRng.Characters.Count Then
                ' Precisely locate number area within field / 精准定位域内数字区域
                Set numRng = fieldRng.Duplicate
                numRng.Collapse Direction:=wdCollapseStart
                numRng.MoveStart Unit:=wdCharacter, Count:=numStartInField - 1
                numRng.MoveEnd Unit:=wdCharacter, Count:=Len(currentNum)

                ' Add jump after verifying bookmark exists / 验证书签存在后添加跳转
                If BookmarkExists("ZotRef_" & currentNum) Then
                    ActiveDocument.Hyperlinks.Add _
                        Anchor:=numRng, _
                        Address:="", _
                        SubAddress:="ZotRef_" & currentNum, _
                        TextToDisplay:=currentNum
                    linkCount = linkCount + 1
                End If
            End If
            currentNum = "" ' Reset, prepare for next number / 重置，准备下一个数字
        End If
    Next i

    ' Process last number (e.g., 33 in [33]) / 处理最后一个数字（如[33]的33）
    If currentNum <> "" Then
        If numStartInField + Len(currentNum) - 1 <= fieldRng.Characters.Count Then
            ' Precisely locate number area within field / 精准定位域内数字区域
            Set numRng = fieldRng.Duplicate
            numRng.Collapse Direction:=wdCollapseStart
            numRng.MoveStart Unit:=wdCharacter, Count:=numStartInField - 1
            numRng.MoveEnd Unit:=wdCharacter, Count:=Len(currentNum)

            ' Add jump after verifying bookmark exists / 验证书签存在后添加跳转
            If BookmarkExists("ZotRef_" & currentNum) Then
                ActiveDocument.Hyperlinks.Add _
                    Anchor:=numRng, _
                    Address:="", _
                    SubAddress:="ZotRef_" & currentNum, _
                    TextToDisplay:=currentNum
                linkCount = linkCount + 1
            End If
        End If
    End If

    AddLinksToContent = linkCount
End Function

' ============================================================================
' Helper Function: Check if bookmark exists / 辅助函数：检查书签是否存在
' ============================================================================
Function BookmarkExists(bkName As String) As Boolean
    Dim bk As Bookmark
    For Each bk In ActiveDocument.Bookmarks
        If bk.Name = bkName Then
            BookmarkExists = True
            Exit Function
        End If
    Next bk
    BookmarkExists = False
End Function
