VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���ִ�ӡ��2.0�����Ʒ��"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   10275
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "���ִ�ӡ����֧�����弰�����С�������塿.frx":0000
      Left            =   8160
      List            =   "���ִ�ӡ����֧�����弰�����С�������塿.frx":0010
      TabIndex        =   13
      Text            =   "����"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "���ִ�ӡ����֧�����弰�����С�������塿.frx":0030
      Left            =   5880
      List            =   "���ִ�ӡ����֧�����弰�����С�������塿.frx":004F
      TabIndex        =   11
      Text            =   "��"
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Text            =   "����"
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
      Caption         =   "���"
      Height          =   615
      Left            =   6840
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "ȷ��"
      Height          =   615
      Left            =   6840
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "��ƷԴ��������ϵ����QQ��3288243945 "
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���Σ�Ĭ�ϳ��棩"
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "������ɫ��Ĭ����ɫ��"
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ִ�ӡ��2.0�����Ʒ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���壨Ĭ�����壩"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "�����С�����Ĭ��Ϊ20��"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "�����·��������� "
      BeginProperty Font 
         Name            =   "����"
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
If Combo1.Text = "��" Then
   Dim color As Long
   color = vbRed
ElseIf Combo1.Text = "��" Then
   color = vbBlue
ElseIf Combo1.Text = "��" Then
   color = vborange
ElseIf Combo1.Text = "��" Then
   color = vbYellow
ElseIf Combo1.Text = "��" Then
   color = vbGreen
ElseIf Combo1.Text = "��" Then
   color = vbCyan
ElseIf Combo1.Text = "��" Then
   color = vbBlack
ElseIf Combo1.Text = "��" Then
   color = vbWhite
ElseIf Combo1.Text = "��" Then
   color = vbMagenta
End If
If Combo1.Text = "����" Then
   Dim bold, italic As Boolean
   bold = False
   italic = False
ElseIf Combo2.Text = "����" Then
   bold = True
   italic = False
ElseIf Combo2.Text = "��б" Then
   bold = False
   italic = True
ElseIf Combo2.Text = "��ƫб��" Then
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


