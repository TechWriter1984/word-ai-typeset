VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChat 
   Caption         =   "DeepSeek-V4-Pro 对话助手"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   StartUpPosition =   1  'CenterOwner
   Begin {8BD21D40-EC42-11CE-9E0D-00AA006002F3} txtChat 
      Height          =   3795
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'fmScrollBarsVertical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   8760
   End
   Begin {8BD21D40-EC42-11CE-9E0D-00AA006002F3} txtInput 
      Height          =   960
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'fmScrollBarsVertical
      TabIndex        =   4
      Top             =   4280
      Width           =   8760
   End
   Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} chkThinking 
      Caption         =   "深度思考"
      Height          =   240
      Left            =   4560
      TabIndex        =   2
      Top             =   60
      Value           =   -1  'True
      Width           =   1500
   End
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblStatus 
      Caption         =   "就绪"
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   4200
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} cmdSend 
      Caption         =   "发送"
      Default         =   -1  'True
      Height          =   360
      Left            =   7560
      TabIndex        =   9
      Top             =   5420
      Width           =   1320
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} cmdClear 
      Caption         =   "清空对话"
      Height          =   360
      Left            =   4980
      TabIndex        =   8
      Top             =   5420
      Width           =   1380
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} cmdLoadComments 
      Caption         =   "读取评论"
      Height          =   360
      Left            =   3360
      TabIndex        =   7
      Top             =   5420
      Width           =   1500
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} cmdLoadStyles 
      Caption         =   "读取样式"
      Height          =   360
      Left            =   1740
      TabIndex        =   6
      Top             =   5420
      Width           =   1500
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} cmdLoadDoc 
      Caption         =   "读取文档"
      Height          =   360
      Left            =   120
      TabIndex        =   5
      Top             =   5420
      Width           =   1500
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 对话历史（仅存 user/assistant 消息，system 消息每次发送时实时拼接）
Private mHistory As Collection

Private Sub UserForm_Initialize()
    Set mHistory = New Collection
    chkThinking.Value = True
    txtChat.Text = "你好！我是 DeepSeek-V4-Pro。" & vbCrLf & _
        "可以点击下方【读取文档 / 读取样式 / 读取评论】把当前 Word 文档的信息作为上下文，" & vbCrLf & _
        "然后直接在下方输入问题并回车发送。" & vbCrLf & _
        String(60, "-") & vbCrLf
    lblStatus.Caption = "就绪"
End Sub

Private Sub cmdSend_Click()
    Dim q As String
    q = Trim(txtInput.Text)
    If Len(q) = 0 Then
        MsgBox "请输入要发送的内容。", vbExclamation, "提示"
        txtInput.SetFocus
        Exit Sub
    End If

    ' 记录并显示用户消息
    mHistory.Add BuildRoleMsg("user", q)
    AppendChat "我", q
    txtInput.Text = ""

    SendAndShow
End Sub

Private Sub SendAndShow()
    lblStatus.Caption = "正在请求 DeepSeek-V4-Pro …"
    DoEvents

    ' 组装完整 messages 数组：system + 历史
    Dim msgs As String, i As Long
    msgs = "["
    msgs = msgs & BuildRoleMsg("system", GetSystemPrompt())
    For i = 1 To mHistory.Count
        msgs = msgs & "," & mHistory.Item(i)
    Next i
    msgs = msgs & "]"

    Dim reply As String, reasoning As String
    Dim ok As Boolean
    ok = CallDeepSeek(msgs, CBool(chkThinking.Value), reply, reasoning)

    If Not ok Then
        AppendChat "助手", "(请求失败) " & reply
        lblStatus.Caption = "请求失败"
        Exit Sub
    End If

    ' 记录助手消息（只记最终回答，思维链仅展示不参与后续）
    mHistory.Add BuildRoleMsg("assistant", reply)

    If Len(reasoning) > 0 Then
        AppendChat "深度思考", reasoning
    End If
    AppendChat "助手", reply
    lblStatus.Caption = "就绪"
End Sub

Private Sub cmdLoadDoc_Click()
    Dim t As String
    t = GetDocumentText()
    mHistory.Add BuildRoleMsg("user", "以下是当前 Word 文档的正文内容，请据此回答后续问题：" & vbCrLf & t)
    AppendChat "系统", "已读取文档正文（共 " & Len(t) & " 字符），作为上下文。"
    txtInput.SetFocus
End Sub

Private Sub cmdLoadStyles_Click()
    Dim t As String
    t = GetDocumentStyles()
    mHistory.Add BuildRoleMsg("user", "以下是当前 Word 文档正在使用的段落样式：" & vbCrLf & t)
    AppendChat "系统", "已读取文档样式信息，作为上下文。"
    txtInput.SetFocus
End Sub

Private Sub cmdLoadComments_Click()
    Dim t As String
    t = GetDocumentComments()
    mHistory.Add BuildRoleMsg("user", "以下是当前 Word 文档的评论信息：" & vbCrLf & t)
    AppendChat "系统", "已读取文档评论信息，作为上下文。"
    txtInput.SetFocus
End Sub

Private Sub cmdClear_Click()
    Set mHistory = New Collection
    txtChat.Text = ""
    AppendChat "系统", "对话已清空，可重新开始。"
End Sub

Private Sub AppendChat(ByVal who As String, ByVal text As String)
    ' 防止文本框溢出，超长时裁掉开头
    If Len(txtChat.Text) > 30000 Then
        txtChat.Text = Right(txtChat.Text, 20000)
    End If
    txtChat.Text = txtChat.Text & "【" & who & "】" & vbCrLf & text & vbCrLf & String(60, "-") & vbCrLf
    txtChat.SelStart = Len(txtChat.Text)
End Sub