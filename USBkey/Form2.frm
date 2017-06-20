VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   465
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   1605
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   465
   ScaleWidth      =   1605
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "本系统已被U盘锁定请插入钥匙盘解锁"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Load()
    Dim i As Long
    i = SetWindowPos(Form2.hwnd, -1, 0, 0, 0, 0, 3)
End Sub
