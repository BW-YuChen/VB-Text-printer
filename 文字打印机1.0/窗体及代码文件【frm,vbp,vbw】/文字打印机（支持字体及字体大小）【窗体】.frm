VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   5205
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Text            =   "宋体"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Text            =   "20"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "确认"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "字体（默认宋体）"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "字体大小（必填，默认为20）"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2520
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
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
