VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USBreg"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   2400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Éú³É×¢²áÂë"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "×¢²áÂë£º"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "»úÆ÷Âë£º"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text2.Text = Trim(Str(Val("&H" + Trim(Text1.Text))))
End Sub
