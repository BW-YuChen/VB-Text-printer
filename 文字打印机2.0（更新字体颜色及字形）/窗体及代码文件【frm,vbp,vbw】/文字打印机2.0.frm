VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "文字打印机2.0（宇辰出品）"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   10275
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "文字打印机（支持字体及字体大小）【窗体】.frx":0000
      Left            =   8160
      List            =   "文字打印机（支持字体及字体大小）【窗体】.frx":0010
      TabIndex        =   13
      Text            =   "常规"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "文字打印机（支持字体及字体大小）【窗体】.frx":0030
      Left            =   5880
      List            =   "文字打印机（支持字体及字体大小）【窗体】.frx":004F
      TabIndex        =   11
      Text            =   "蓝"
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Text            =   "宋体"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Text            =   "20"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   5775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   615
      Left            =   6840
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "确认"
      Height          =   615
      Left            =   6840
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "作品源代码请联系作者QQ：3288243945 "
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   5400
      Width           =   9855
   End
   Begin VB.Label Label7 
      Caption         =   "字形（默认常规）"
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "字体颜色（默认蓝色）"
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "文字属性"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "文字打印机2.0（宇辰出品）"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "字体（默认宋体）"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "字体大小（必填，默认为20）"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "请在下方输入文字 "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.ForeColor = &HFF0000
Label1.FontSize = Text2.Text
Label1.FontName = Text3.Text
If Combo1.Text = "红" Then
   Dim color As Long
   color = vbRed
ElseIf Combo1.Text = "蓝" Then
   color = vbBlue
ElseIf Combo1.Text = "橙" Then
   color = vborange
ElseIf Combo1.Text = "黄" Then
   color = vbYellow
ElseIf Combo1.Text = "绿" Then
   color = vbGreen
ElseIf Combo1.Text = "青" Then
   color = vbCyan
ElseIf Combo1.Text = "黑" Then
   color = vbBlack
ElseIf Combo1.Text = "白" Then
   color = vbWhite
ElseIf Combo1.Text = "粉" Then
   color = vbMagenta
End If
If Combo1.Text = "常规" Then
   Dim bold, italic As Boolean
   bold = False
   italic = False
ElseIf Combo2.Text = "粗体" Then
   bold = True
   italic = False
ElseIf Combo2.Text = "倾斜" Then
   bold = False
   italic = True
ElseIf Combo2.Text = "粗偏斜体" Then
   bold = True
   italic = True
End If
Label1.ForeColor = color
Label1.FontBold = bold
Label1.FontItalic = italic
Label1.Caption = Text1.Text
End Sub
Private Sub Command2_Click()
Label1.Caption = " "
End Sub


