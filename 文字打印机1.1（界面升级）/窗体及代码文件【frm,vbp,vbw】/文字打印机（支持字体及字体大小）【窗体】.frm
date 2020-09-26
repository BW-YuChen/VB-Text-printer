VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   7050
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Text            =   "宋体"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Text            =   "20"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "确认"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   2880
      Width           =   975
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
      Caption         =   "文字打印机（宇辰出品）1.0"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "字体大小（必填，默认为20）"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "请在下方输入文字，仅支持蓝色字体 "
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
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   5295
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

Label1.Caption = Text1.Text
End Sub

Private Sub Command2_Click()
Label1.Caption = " "
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Text1_Change()
Label1.Caption = " "
End Sub
