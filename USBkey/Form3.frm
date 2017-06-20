VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "注册"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3315
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "继续试用"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "注册"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "http://reg.banma.com/"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "注册码："
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "机器码："
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "注册后，本提示将不再出现"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "您正在使用的是USBkey试用版，请注册"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Command1_Click()
    If Text2.Text = "" Then
        Text2.Text = "请输入注册码"
        Text2.SetFocus
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2.Text)
        Beep
    ElseIf Trim(Text2.Text) = Trim(Str(Val("&H" + Text1.Text))) Then
        Open Form1.sysPath & "\usbreg.ini" For Output As #1
            Write #1, Trim(Text1.Text)
        Close #1
        Unload Form3
    Else
        Text2.Text = "注册码错误"
        Text2.SetFocus
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2.Text)
        Beep
    End If
End Sub

Private Sub Command2_Click()
    Unload Form3
End Sub

Private Sub Form_Load()
    Dim i As Long
    i = SetWindowPos(Form3.hwnd, -1, 0, 0, 0, 0, 3)
    Text1.Text = Hex(Form1.UVolSeri)
End Sub
