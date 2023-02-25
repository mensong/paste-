VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paste++"
   ClientHeight    =   1005
   ClientLeft      =   8280
   ClientTop       =   4080
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   6.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   6435
   Begin VB.CheckBox Vaild 
      Caption         =   "是否生效"
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   0
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox AutoPaste 
      Caption         =   "自动粘贴"
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   360
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ctrl + 5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   4440
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ctrl + 4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3360
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ctrl + 3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ctrl + 2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ctrl + 1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    SetHotKey
    
    'SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3 '窗口置顶
    
    Dim i As Integer
    For i = 0 To 4
        Text(i).MousePointer = 1
        Text(i).ToolTipText = "CTRL +" & Str(i + 1)
    Next i
    
    'Me.Width = Screen.Width
    'Frame1.Move Me.Width - Frame1.Width
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim ret As Long
    '取消Message的截取，而使之又只送往原来的Window Procedure
    ret = SetWindowLong(Me.hwnd, GWL_WNDPROC, preWinProc)
    Call UnregisterHotKey(Me.hwnd, uVirtKey)
End Sub


Function Paste(ByVal Index As Integer)
    If Vaild.Value = 1 And Text(Index).Text <> "" Then
        Clipboard.Clear
        Clipboard.SetText Me.Text(Index).Text
        
        '自动粘贴
        If AutoPaste.Value = 1 Then
            SendKeys "^v"
        End If
    End If
End Function

Private Sub Text_DblClick(Index As Integer)
    Paste (Index)
End Sub
