Attribute VB_Name = "modDeepSeek"
Option Explicit

'============================================================
' DeepSeek-V4-Pro 对话助手 —— 公共模块
'
' 功能：
'   1. 调用 DeepSeek API（模型 deepseek-v4-pro，OpenAI 兼容接口）
'   2. 读取当前 Word 文档的正文 / 样式 / 评论
'   3. 提供 ShowChatDialog 入口，弹出对话窗体 frmChat
'
' 导入方法（VBA 编辑器 Alt+F11）：
'   文件 -> 导入文件 -> 选择 modDeepSeek.bas 和 frmChat.frm
'   然后在【宏】里运行 ShowChatDialog，或把它加到快速访问工具栏/
'   自定义快捷键上，即可在 Word 界面一键打开对话框。
'============================================================

'---------------- 配置区（按需修改）----------------
' DeepSeek API Key：在 https://platform.deepseek.com 申请，sk- 开头
Private Const DEEPSEEK_API_KEY As String = "sk-REPLACE_ME"

' 若不想把 Key 写死在这里，可以留 sk-REPLACE_ME，然后新建文件
' %APPDATA%\deepseek_key.txt（首行写 sk-xxx），程序会自动读取（见 GetApiKey）

' DeepSeek API 地址（OpenAI 兼容，不必加 /v1）
Private Const DEEPSEEK_URL As String = "https://api.deepseek.com/chat/completions"

' 模型名：DeepSeek-V4-Pro 正式版（2026-08-13 上线的最新版本）
Private Const DEEPSEEK_MODEL As String = "deepseek-v4-pro"

' 单次回答最大输出 token 数
Private Const DEEPSEEK_MAX_TOKENS As Long = 8192

' 系统提示词
Private Const SYSTEM_PROMPT As String = _
    "你是嵌入在 Microsoft Word 中的文档助手。你会收到来自当前 Word 文档的正文、样式和评论信息，" & _
    "请基于这些信息回答用户关于文档的问题（总结、翻译、改写、校对、找问题等）。" & _
    "回答使用简洁、专业、准确的中文；除非用户明确要求，否则不要输出与文档无关的内容。"

' 读取文档时正文文本的最大长度（避免超出上下文）
Private Const MAX_DOC_TEXT As Long = 50000
'============================================================


'================= 入口 =================
Public Sub ShowChatDialog()
    ' 以非模态方式打开，便于用户在对话的同时翻阅/编辑文档
    frmChat.Show vbModeless
End Sub


'================= 对外：读取文档信息 =================

Public Function GetDocumentStyles() As String
    ' 汇总当前文档中正在使用的段落样式（含出现次数）
    On Error Resume Next
    Dim p As Paragraph
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim sn As String
    For Each p In ActiveDocument.Paragraphs
        sn = p.Style.NameLocal
        If Len(sn) = 0 Then sn = "(未命名)"
        If dict.Exists(sn) Then
            dict(sn) = CLng(dict(sn)) + 1
        Else
            dict(sn) = 1
        End If
    Next p

    Dim k As Variant, sb As String
    sb = "文档中正在使用的段落样式（共 " & dict.Count & " 种）：" & vbCrLf
    For Each k In dict.Keys
        sb = sb & "  - " & k & "  (" & dict(k) & " 段)" & vbCrLf
    Next k
    GetDocumentStyles = sb
End Function

Public Function GetDocumentComments() As String
    ' 读取文档所有批注：作者、时间、批注内容、被批注的文本
    On Error Resume Next
    If ActiveDocument.Comments.Count = 0 Then
        GetDocumentComments = "（当前文档没有评论/批注）"
        Exit Function
    End If

    Dim c As Comment
    Dim sb As String
    sb = "文档评论（共 " & ActiveDocument.Comments.Count & " 条）：" & vbCrLf
    For Each c In ActiveDocument.Comments
        sb = sb & "[" & c.Index & "] 作者: " & c.Author & vbCrLf
        sb = sb & "    评论内容: " & Trim(c.Range.Text) & vbCrLf
        On Error Resume Next
        sb = sb & "    被评论文本: " & Trim(c.Scope.Text) & vbCrLf
        sb = sb & vbCrLf
    Next c
    GetDocumentComments = sb
End Function

Public Function GetDocumentText() As String
    ' 读取正文（段落 + 表格），并用[样式]标注标题等结构
    On Error Resume Next
    Dim sb As String
    Dim p As Paragraph
    Dim t As String

    For Each p In ActiveDocument.Paragraphs
        t = Trim(p.Range.Text)
        If Len(t) > 0 Then
            sb = sb & HeadingPrefix(p) & t & vbCrLf
        End If
    Next p

    ' 表格内容单独标注
    Dim tbl As Table, r As Row, cell As Cell
    For Each tbl In ActiveDocument.Tables
        sb = sb & "[表格]" & vbCrLf
        For Each r In tbl.Rows
            Dim rowText As String
            rowText = ""
            For Each cell In r.Cells
                rowText = rowText & Trim(CleanCellText(cell.Range.Text)) & " | "
            Next cell
            sb = sb & "  " & Trim(rowText) & vbCrLf
        Next r
        sb = sb & "[/表格]" & vbCrLf
    Next tbl

    If Len(sb) > MAX_DOC_TEXT Then
        sb = Left(sb, MAX_DOC_TEXT) & vbCrLf & "…（正文过长，已截断）"
    End If
    GetDocumentText = sb
End Function


'================= 对外：调用模型 =================

Public Function BuildRoleMsg(ByVal role As String, ByVal content As String) As String
    ' 把一个 role/content 转成 JSON 消息对象字符串
    BuildRoleMsg = "{""role"":""" & role & """,""content"":""" & JsonEscape(content) & """}"
End Function

Public Function GetSystemPrompt() As String
    GetSystemPrompt = SYSTEM_PROMPT
End Function

Public Function CallDeepSeek(ByVal messagesJson As String, _
                             ByVal enableThinking As Boolean, _
                             ByRef reply As String, _
                             ByRef reasoning As String) As Boolean
    ' 发送一次对话请求。
    ' 参数：
    '   messagesJson : 形如 [{"role":..,"content":..}, ...] 的 JSON 数组字符串
    '   enableThinking: 是否开启深度思考模式
    ' 返回：
    '   reply     : 最终回答文本
    '   reasoning : 思维链内容（思考模式开启时才有）
    Dim thinking As String
    If enableThinking Then
        thinking = """thinking"":{""type"":""enabled""}"
    Else
        thinking = """thinking"":{""type"":""disabled""}"
    End If

    Dim payload As String
    payload = "{" & _
        """model"":""" & DEEPSEEK_MODEL & """," & _
        """messages"":" & messagesJson & "," & _
        thinking & "," & _
        """stream"":false," & _
        """max_tokens"":" & DEEPSEEK_MAX_TOKENS & "}"

    Dim respText As String
    If Not HttpPostJson(DEEPSEEK_URL, payload, respText) Then
        reply = respText
        reasoning = ""
        CallDeepSeek = False
        Exit Function
    End If

    reply = ExtractJsonString(respText, "content")
    reasoning = ExtractJsonString(respText, "reasoning_content")

    ' 若两者都为空，可能是接口返回了错误对象，尝试提取错误信息
    If reply = "" And reasoning = "" Then
        Dim apiMsg As String
        apiMsg = ExtractJsonString(respText, "message")
        If apiMsg <> "" Then
            reply = "API 错误：" & apiMsg
            CallDeepSeek = False
            Exit Function
        End If
    End If
    CallDeepSeek = True
End Function


'================= 内部：HTTP 请求 =================

Private Function GetApiKey() As String
    ' 优先用常量；若未填写则尝试从 %APPDATA%\deepseek_key.txt 读取首行
    If DEEPSEEK_API_KEY <> "" And DEEPSEEK_API_KEY <> "sk-REPLACE_ME" Then
        GetApiKey = DEEPSEEK_API_KEY
        Exit Function
    End If
    On Error Resume Next
    Dim fPath As String, fnum As Integer, line As String
    fPath = Environ("APPDATA") & "\deepseek_key.txt"
    If Dir(fPath) <> "" Then
        fnum = FreeFile
        Open fPath For Input As #fnum
        Line Input #fnum, line
        Close #fnum
        GetApiKey = Trim(line)
    Else
        GetApiKey = ""
    End If
End Function

Private Function HttpPostJson(ByVal url As String, ByVal payload As String, ByRef responseText As String) As Boolean
    On Error GoTo fail
    Dim http As Object
    Dim progIds As Variant, i As Long
    progIds = Array("MSXML2.ServerXMLHTTP.6.0", "MSXML2.XMLHTTP.6.0", _
                    "MSXML2.ServerXMLHTTP", "MSXML2.XMLHTTP", _
                    "WinHttp.WinHttpRequest.5.1")
    For i = LBound(progIds) To UBound(progIds)
        On Error Resume Next
        Set http = CreateObject(progIds(i))
        On Error GoTo fail
        If Not http Is Nothing Then Exit For
    Next i
    If http Is Nothing Then
        responseText = "无法创建 HTTP 对象（请确认系统支持 MSXML/WinHttp）。"
        HttpPostJson = False
        Exit Function
    End If

    ' 超时设置（部分对象不支持 setTimeouts，忽略错误即可）
    On Error Resume Next
    http.setTimeouts 30000, 30000, 60000, 120000
    On Error GoTo fail

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & GetApiKey()
    http.send payload

    Select Case http.Status
        Case 200
            responseText = http.responseText
        Case 401
            responseText = "认证失败（401）：API Key 错误或未填写。"
            HttpPostJson = False
            Exit Function
        Case 402
            responseText = "余额不足（402）：请到 DeepSeek 平台充值。"
            HttpPostJson = False
            Exit Function
        Case 429
            responseText = "请求过于频繁（429）：请稍后再试。"
            HttpPostJson = False
            Exit Function
        Case Else
            responseText = "API 返回错误 " & http.Status & "：" & http.responseText
            HttpPostJson = False
            Exit Function
    End Select
    HttpPostJson = True
    Exit Function

fail:
    responseText = "HTTP 请求失败：" & Err.Number & " - " & Err.Description
    HttpPostJson = False
End Function


'================= 内部：JSON 处理 =================

Private Function JsonEscape(ByVal s As String) As String
    ' 转义成合法的 JSON 字符串内容（不含首尾引号）
    Dim i As Long, c As String, out As String
    out = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case c
            Case """":  out = out & "\"""
            Case "\":   out = out & "\\"
            Case vbCr:  out = out & "\r"
            Case vbLf:  out = out & "\n"
            Case vbTab: out = out & "\t"
            Case Else
                If AscW(c) < 32 Then
                    out = out & "\u" & Right("0000" & Hex(AscW(c)), 4)
                Else
                    out = out & c
                End If
        End Select
    Next i
    JsonEscape = out
End Function

Private Function ExtractJsonString(ByVal json As String, ByVal key As String) As String
    ' 在 JSON 文本中提取 "key":"字符串值"（支持常见转义，返回未转义内容）。
    ' 找不到或值为 null 时返回空字符串。
    Dim marker As String, pos As Long, startPos As Long
    marker = """" & key & """"
    pos = InStr(1, json, marker, vbTextCompare)
    If pos = 0 Then ExtractJsonString = "": Exit Function

    pos = InStr(pos + Len(marker), json, ":")
    If pos = 0 Then ExtractJsonString = "": Exit Function

    startPos = pos + 1
    Do While startPos <= Len(json)
        Dim ch As String
        ch = Mid(json, startPos, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            startPos = startPos + 1
        Else
            Exit Do
        End If
    Loop

    If startPos > Len(json) Then ExtractJsonString = "": Exit Function
    If Mid(json, startPos, 4) = "null" Then ExtractJsonString = "": Exit Function
    If Mid(json, startPos, 1) <> """" Then ExtractJsonString = "": Exit Function

    Dim i As Long, c As String, out As String
    i = startPos + 1
    out = ""
    Do While i <= Len(json)
        c = Mid(json, i, 1)
        If c = "\" Then
            i = i + 1
            If i > Len(json) Then Exit Do
            Select Case Mid(json, i, 1)
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case """", "\", "/": out = out & Mid(json, i, 1)
                Case "u":
                    If i + 4 <= Len(json) Then
                        out = out & ChrW(CLng("&H" & Mid(json, i + 1, 4)))
                        i = i + 4
                    End If
                Case Else: out = out & Mid(json, i, 1)
            End Select
        ElseIf c = """" Then
            Exit Do
        Else
            out = out & c
        End If
        i = i + 1
    Loop
    ExtractJsonString = out
End Function


'================= 内部：结构辅助 =================

Private Function HeadingPrefix(p As Paragraph) As String
    ' 为标题段落加 [样式] 前缀，普通正文返回空
    Dim sn As String
    sn = LCase(p.Style.NameLocal)
    If sn Like "heading*" Or sn Like "标题*" Then
        HeadingPrefix = "[" & p.Style.NameLocal & "] "
    Else
        HeadingPrefix = ""
    End If
End Function

Private Function CleanCellText(ByVal s As String) As String
    ' 去掉单元格文本里的段落/单元格结束标记
    CleanCellText = Replace(Replace(Replace(s, Chr(7), ""), vbCr, " "), vbLf, " ")
End Function