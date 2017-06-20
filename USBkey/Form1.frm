VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "U盘电脑锁"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3315
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "状态："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const MAX_PATH = 260
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Const FILE_VOLUME_IS_COMPRESSED = &H8000
Dim removable As String
Public sysPath As String
Public UVolSeri As Long
Dim VolSeri As Long
Dim ifHooked As Boolean
Dim lastIsin As Boolean
Dim lastLen As Long

Private Sub Command1_Click()
    On Error Resume Next
    If Command1.Caption = "关闭" Then
        End
    Else
        'make me a lock
        SetAttr sysPath & "\usbkey.exe", vbNormal
        FileCopy App.path & IIf(Len(App.path) > 3, "\", "") & App.EXEName & ".exe", sysPath & "\usbkey.exe"
        SetAttr sysPath & "\usbkey.exe", vbHidden + vbSystem
        
        Open sysPath & "\usbkey.ini" For Output As #1
            Write #1, VolSeri
        Close #1
        DoEvents
        MsgBox "如果您安装了注册表监测类软件，可能会使本程序不能正常启动，请在监测软件询问时选择“允许修改”", , "注意"
        makeRun
        MsgBox "上锁成功，一旦拔出U盘，系统立即上锁，插入此U盘即解锁", , "U盘电脑锁"
        Shell sysPath & "\usbkey.exe"
        End
    End If
        
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    If Combo1.ListIndex <> -1 Then
        removable = Combo1.List(Combo1.ListIndex)
        Combo1.Visible = False
        Command2.Visible = False
        
        Label1.Caption = "状态：已选择多个移动设备中的一个"
        Command1.Caption = "上锁"
        GetVolInfo (removable)
    Else
        Combo1.Text = ""
    End If
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    App.TaskVisible = False
    If App.PrevInstance Then End 'don't run me again
    
    ifHooked = False
    lastIsin = False
    lastLen = 256
    sysPath = SystemDir
    If CBool(PathFileExists(sysPath & "\usbkey.ini")) Then
        Me.Hide
        
        If App.path = sysPath Then
            Open sysPath & "\usbkey.ini" For Input As #1
                Input #1, UVolSeri
            Close #1
            Timer1.Enabled = True
            If CBool(PathFileExists(sysPath & "\usbreg.ini")) = False Then
                Form3.Show
            End If
        Else
            'remove
            If MsgBox("本系统已上锁，是否永久去除锁定？", vbOKCancel, "U盘电脑锁") = vbOK Then
                unmakeRun
                SetAttr sysPath & "\usbkey.ini", vbNormal
                Kill sysPath & "\usbkey.ini"
                Shell "taskkill /F /IM usbkey.exe", vbHide
                MsgBox "本程序已成功卸载(非XP的系统需要重启才能生效)", , "U盘电脑锁"
            End If
            End
        End If
    Else
        If App.path = sysPath Then
            End
        Else
            'show ui
            FindDriver
            If Combo1.ListCount = 0 Then
                Label1.Caption = "状态：找不到可移动存储设备"
                Command1.Caption = "关闭"
                Label7.Caption = "请确保系统识别您的设备并重启本程序"
            ElseIf Combo1.ListCount = 1 Then
                removable = Combo1.List(0) & "\"
                Label1.Caption = "状态：检测到可移动存储设备"
                Command1.Caption = "上锁"
                GetVolInfo (removable)
            Else
                Label1.Caption = "状态：检测到多个可移动设备"
                Combo1.Visible = True
                Combo1.ListIndex = 0
                Command2.Visible = True
                Command1.Caption = "关闭"
                Label7.Caption = "请选择："
            End If
        End If
    End If
    
End Sub

Public Function SystemDir() As String
    On Error Resume Next
    Dim tmp As String
    tmp = Space$(MAX_PATH)
    SystemDir = Left$(tmp, GetSystemDirectory(tmp, MAX_PATH))
End Function

Public Sub FindDriver()
    On Error Resume Next
    Dim totlen As Long
    Dim buff As String
    Dim i As Long
    Dim diskType As Long
    buff = String(255, 0)
    totlen = GetLogicalDriveStrings(256, buff)
    '取得的值如: "a:\"+Chr(0)+"c:\"+Chr(0) + "d:\"+Chr(0) + Chr(0)
    '而这个例子中传回长度(totlen)是12
    buff = Left(buff, totlen)
    For i = 1 To totlen Step 4
        diskType = GetDriveType(Mid(buff, i, 3))
        If diskType = 2 Then
            Combo1.AddItem (Mid(buff, i, 2))
        End If
    Next i
End Sub

Public Sub GetVolInfo(ByVal path As String)
    On Error Resume Next
    Dim aa As Long
    Dim VolName As String
    Dim fsysName As String
    Dim compress As Long
    Dim Sysflag As Long, Maxlen As Long
    
    VolName = String(255, 0)
    fsysName = String(255, 0)
    aa = GetVolumeInformation(path, VolName, 256, VolSeri, Maxlen, Sysflag, fsysName, 256)
    VolName = Left(VolName, InStr(1, VolName, Chr(0)) - 1)
    fsysName = Left(fsysName, InStr(1, fsysName, Chr(0)) - 1)
    compress = Sysflag And FILE_VOLUME_IS_COMPRESSED
    If compress = 0 Then
        Label2.Caption = "非压缩卷"
    Else
        Label2.Caption = "压缩卷"
    End If
    Label2.Visible = True
    
    Label3.Caption = "卷名：" & VolName
    Label3.Visible = True
    
    Label4.Caption = "磁盘序列号 : " & Hex(VolSeri)
    Label4.Visible = True
    
    Label5.Caption = "文件系统：" & fsysName
    Label5.Visible = True
    
    Label6.Caption = "最大文件名长度：" & Maxlen
    Label6.Visible = True
    
    Label7.Caption = "卷标：" & Left(path, 2)

    Label8.Caption = "此设备将作为本计算机的唯一钥匙，请妥善保管，是否将本计算机上锁？"
    Label8.Visible = True
End Sub

Private Sub makeRun()
    On Error Resume Next
     SetAttr sysPath & "\makeKRun.reg", vbNormal
    Open sysPath & "\makeKRun.reg" For Output As #1
        Print #1, "REGEDIT4"
        Print #1, ""
        Print #1, "[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon]"
        Print #1, """Userinit""=""userinit.exe,usbkey.exe"""
    Close #1
    Shell "regedit.exe /s " & sysPath & "\makeKRun.reg", vbHide
    SetAttr sysPath & "\makeKRun.reg", vbSystem + vbHidden
End Sub

Private Sub unmakeRun()
    On Error Resume Next
    SetAttr sysPath & "\unmakeKR.reg", vbNormal
    Open sysPath & "\unmakeKR.reg" For Output As #1
        Print #1, "REGEDIT4"
        Print #1, ""
        Print #1, "[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon]"
        Print #1, """Userinit""=""userinit.exe,"""
    Close #1
    Shell "regedit.exe /s " & sysPath & "\unmakeKR.reg", vbHide
    SetAttr sysPath & "\unmakeKR.reg", vbSystem + vbHidden
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Dim DrvType As Long
    Dim totlen As Long
    Dim buff As String
    Dim i As Long
    
    buff = String(255, 0)
    totlen = GetLogicalDriveStrings(256, buff)
    '取得的值如: "a:\"+Chr(0)+"c:\"+Chr(0) + "d:\"+Chr(0) + Chr(0)
    '而这个例子中传回长度(totlen)是12
    
    If totlen > lastLen Then '有新插入盘
        If ifHooked = True Then '检查是否该解锁
            buff = Left(buff, totlen)
            For i = 1 To totlen Step 4
                DrvType = GetDriveType(Mid(buff, i, 3))
                If DrvType = 2 Then
                    removable = Mid(buff, i, 3)
                    GetVolInfo (removable)
                    
                    If VolSeri = UVolSeri Then
                        ifHooked = False
                        Call UnHook
                        Form2.Hide
                        Exit For
                    End If
                End If
            Next i
        End If
        lastLen = totlen
    ElseIf totlen < lastLen Then '有盘移出
        If ifHooked = False Then
            buff = Left(buff, totlen)
            Dim toLock As Boolean
            toLock = True
            
            For i = 1 To totlen Step 4
                DrvType = GetDriveType(Mid(buff, i, 3))
                If DrvType = 2 Then
                    removable = Mid(buff, i, 3)
                    GetVolInfo (removable)
                    
                    If VolSeri = UVolSeri Then
                        toLock = False
                        Exit For
                    End If
                End If
            Next i
            
            If toLock = True Then
                ifHooked = True
                Call EnableHook
                Form2.Show
            End If
        End If
        lastLen = totlen
    End If
End Sub
