VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   Caption         =   "Smt Login"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   0
      Picture         =   "frmLogin.frx":29C12
      ScaleHeight     =   1635
      ScaleWidth      =   765
      TabIndex        =   8
      Top             =   60
      Width           =   765
      Begin MSWinsockLib.Winsock Winsock 
         Left            =   240
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.Frame fraMain 
      Height          =   1755
      Left            =   780
      TabIndex        =   0
      Top             =   0
      Width           =   4245
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1350
         TabIndex        =   1
         Top             =   210
         Width           =   2115
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   390
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   390
         Left            =   2790
         TabIndex        =   4
         Top             =   1200
         Width           =   1140
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1350
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   630
         Width           =   2115
      End
      Begin VB.CheckBox chkKeepPassword 
         Caption         =   "&Keep Password"
         Height          =   615
         Left            =   150
         TabIndex        =   5
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   765
         Left            =   3540
         Picture         =   "frmLogin.frx":2A954
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   210
         Width           =   585
      End
      Begin VB.Label lblUserName 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   660
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: Smt_PMS.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: LynnSun
'**日    期: 2007.12.22
'**描    述: DID Header
'
'**EQMS_ID      修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**             Udall       2009.05.25     对于仅有一个厂区的BU，不核对其IP (0001)
'**             Lynn        2010.01.08     修改程式，让QSMS使用SMT传过来的连接字符串 (0002)
'***********************************************************************************/

Option Explicit
Public delright As String
Public Authorized As Boolean

Private Sub cmdCancel_Click()
    'End
    Unload Me
End Sub

Private Sub CmdOk_Click()
Dim Rs As New ADODB.Recordset
Dim strSQL As String
On Error GoTo errHandler
Dim i As Long
    
If delright = "" Then
    strSQL = "select * from UserDetail where username='" & txtUserName & "' and password='" & txtPassword & "'"
    If Rs.State = 1 Then Rs.Close
    Rs.CursorLocation = adUseClient
    Rs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    If Rs.RecordCount <= 0 Then
        MsgBox "User name or Password error,Please check"
        Exit Sub
    Else
        g_userName = Trim(txtUserName)
        Set Rs = Nothing
        '''read user right
        '(1004)
        strSQL = "select userright from UserRight where AppName='QSMS' and username='" & txtUserName & "'"
        Set Rs = Conn.Execute(strSQL)
        ReDim g_userRight(Rs.RecordCount - 1)
        For i = 0 To UBound(g_userRight)
            g_userRight(i) = Rs!userright
            Rs.MoveNext
        Next i
        Set Rs = Nothing
        Unload Me
        ''''Main.Show
    End If
    Exit Sub
Else
    ''''''''''''''''''''''''''Add by Jing  20071126   (0006) ''''''''''''''''''''''''''''
    strSQL = "select * from UserDetail where username='" & txtUserName & "' and password='" & txtPassword & "'"
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF Then
        MsgBox "User name or Password is error,please check it !", vbCritical
        Exit Sub
    Else
    '(1004)
        strSQL = "select userright from UserRight where  AppName='QSMS' and username='" & txtUserName & "' and userright='" & delright & "'"
        Set Rs = Conn.Execute(strSQL)
        If Not Rs.EOF Then
            Authorized = True
            g_delrightUser = Trim(txtUserName)  ''(1016)
            delright = ""
            Unload Me
        Else
            MsgBox ("You have not the Authorize !"), vbCritical
            Unload Me
        End If
    End If
    Exit Sub
End If
    
errHandler:
    If Err.Number = -2147217873 Or Err.Number = -2147217900 Then
        MsgBox "Your name have loggined,please login again!!!", vbCritical, "Tip:"
    Else
        MsgBox Err.Description, vbCritical, "Tip"
    End If
        
End Sub

Private Sub Form_Load()
Dim strStation As String
Dim strLine As String
Dim connStr As String
Dim SMTServer As String
Dim QSMSServer As String
Dim SMTDB As String
Dim QSMSDB As String
Dim sql As String
Dim Rs As New ADODB.Recordset

If App.Title <> App.EXEName Then
    If Command = "" Then
        MsgBox "Please Use MainMenu "
        End
    Else
        If InStr(1, Command, "<LINE=", vbTextCompare) > 0 Then
            strLine = Mid(Mid(Command, InStr(1, Command, "<LINE=", vbTextCompare) + Len("<LINE="), Len(Command)), 1, InStr(1, Mid(Command, InStr(1, Command, "<LINE=", vbTextCompare) + Len("<LINE="), Len(Command)), ">") - 1)
        End If
        If InStr(1, Command, "<STATION=", vbTextCompare) > 0 Then
            strStation = Mid(Mid(Command, InStr(1, Command, "<Station=", vbTextCompare) + Len("<Station="), Len(Command)), 1, InStr(1, Mid(Command, InStr(1, Command, "<STATION=", vbTextCompare) + Len("<STATION="), Len(Command)), ">") - 1)
        End If
        If InStr(1, Command, "<CONN=", vbTextCompare) > 0 Then
            connStr = Mid(Mid(Command, InStr(1, Command, "<CONN=", vbTextCompare) + Len("<CONN="), Len(Command)), 1, InStr(1, Mid(Command, InStr(1, Command, "<CONN=", vbTextCompare) + Len("<CONN="), Len(Command)), ">") - 1)
        End If
    End If
    
Else
    End
    strLine = "All"
    strStation = "QSMS"
    connStr = "Provider=sqloledb;UID=sa;server=192.168.20.39;database=SMT;Network Library=DBMSSOCN;pwd=qms7sa"
    'Call BuildMainConnection
End If
    ProgLine = strLine
    Conn.CommandTimeout = 0
    Conn.CursorLocation = adUseClient
    If Conn.State = 1 Then Conn.Close
    Conn.Open connStr

    '''Get SMT Server (0002)
    SMTServer = GetKeyValue(connStr, "server")
    If SMTServer = "" Then
        MsgBox "Cant't get SMT Server information !! Call QMS please! "
        End
    Else
        sql = "select smt_db,qsms_db,QSMS_Server from QSMS_SMT_DB where smt_server='" & Trim(SMTServer) & "'" ''AND  BU='" & tSettings.BU & "'"
        Set Rs = Conn.Execute(sql)
        SMTDB = Rs!SMT_DB
        QSMSDB = Rs!qsms_db
        QSMSServer = Rs!QSMS_Server
    End If
    ''Get QSMS Server
    If QSMSServer = "" Then
        MsgBox "Can't get QSMS Server information ! Call QMS please! "
        End
    Else
        IP = QSMSServer
        connStr = Replace(connStr, SMTServer, QSMSServer)
        connStr = Replace(connStr, SMTDB, QSMSDB)
        connStr = Replace(connStr, LCase(SMTDB), QSMSDB)
    End If
    ''Connect QSMS DB
    If Conn.State = 1 Then Conn.Close
    Conn.CursorLocation = adUseClient
    Conn.Open connStr
    
    Call GetSettings
     
    chkKeepPassword.Value = GetSetting("SMTUT", "Login", "KeepPassword", 0)
    
    txtUserName = GetSetting("SMTUT", "Login", "UserName")
    If NoKeepPWD <> "Y" Then    ''1199
        txtPassword = GetSetting("SMTUT", "Login", "Password")
    End If
 
    ''''''added by Jing (0028)''''''
    chkQty = ReadIniFile("QSMS", "MaxDIDGroupQty", App.Path & "\set.ini")
    StrBU = ReadIniFile("COMMON", "BU", App.Path & "\set.ini")   'add a flag to NB5 for DeleteMe_Bom  (0010)
    ''''''''''(0008)
    If CheckFacIP = False Then          ''(0014)
        End
    End If
 ''      Call ChkVersion("ALL", "QSMS", App.EXEName & ".exe")
End Sub
Public Sub GetSettings()
    Dim strSQL As String
    Dim Rs As New ADODB.Recordset
    
    strSQL = "select * from QSMS_ProConfig where Line='All' and station='QSMS'"
    Set Rs = Conn.Execute(strSQL)
    
    If Rs.EOF = False Then
        While Not Rs.EOF
            Select Case UCase(Rs!key)
                Case "SCANCOMPPN"
                    ScanCompPN = UCase(Rs!Value)
                Case "SCANMSD"
                    ScanMSD = UCase(Rs!Value)
                Case "CHECKBOMLOGON"
                    CheckBomLogon = UCase(Rs!Value)
                Case UCase("CheckReturnForbiddenPN")
                    CheckReturnForbiddenPN = UCase(Rs!Value)
                Case UCase("ChkOldDIDLabelQty")  ''(0061)
                    ChkOldDIDLabelQty = UCase(Rs!Value)
                Case UCase("ChkOneByOneMaterial")  ''(0076)
                    ChkOneByOneMaterial = UCase(Rs!Value)
                Case UCase("NPMMachineType")  ''(1079)
                    NPMMachineType = Trim(Rs!Value)
            End Select
            Rs.MoveNext
        Wend
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If chkKeepPassword.Value = 1 Then
        SaveSetting "SMTUT", "Login", "KeepPassword", "1"
        SaveSetting "SMTUT", "Login", "UserName", txtUserName
        SaveSetting "SMTUT", "Login", "Password", txtPassword
    Else
        SaveSetting "SMTUT", "Login", "KeepPassword", "0"
        SaveSetting "SMTUT", "Login", "UserName", ""
        SaveSetting "SMTUT", "Login", "Password", ""
    End If
End Sub


Private Sub ChkVersion(strLine As String, strStation As String, EXEName As String)
Dim Rs As New ADODB.Recordset
Dim Sqlstr As String
    Sqlstr = "select * from  Application_List  where AppEXE= '" & EXEName & "'"
    If strLine <> "" And UCase(strLine) <> "ALL" Then
       Sqlstr = Sqlstr & " and Line = '" & Trim(strLine) & "' and StationName = '" & strStation & "' "
    End If
    Rs.Open Sqlstr, Conn, adOpenForwardOnly, adLockReadOnly
    If Rs.EOF = True Then
       MsgBox "The Program Version is Wrong,pls Access through MainMenu or Contact QMS!!", vbCritical
       End
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdOK.SetFocus
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtPassword.SetFocus
End If
End Sub

Private Function CheckFacIP() As Boolean
Dim strIP() As String
Dim Rs As New ADODB.Recordset
Dim strSQL As String
Dim i As Integer, j As Integer
    LocalIP = Winsock.LocalIP
    CheckFacIP = False
    Factory = ""
    CreateDIDFlag = "N"
  ''''''(0014)  Start
    strSQL = "select distinct Factory from Site"
    If Rs.State = 1 Then Rs.Close
    Rs.CursorLocation = adUseClient
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        ReDim FactoryID(Rs.RecordCount, 2)
    Else
        MsgBox "The Factory is empty,please connect with QMS for set the Factory in the Site table!"
        Exit Function
    End If
    If Rs.RecordCount > 1 Then      ''(0001)
        While Rs.EOF = False
             FactoryID(i, 0) = Rs.Fields("Factory")
             FactoryID(i, 1) = ReadIniFile("QSMS", Trim(Rs.Fields("Factory")), App.Path & "\set.ini")
             i = i + 1
             Rs.MoveNext
        Wend
        For i = 0 To UBound(FactoryID)
            If Trim(FactoryID(i, 0) <> "" And Trim(FactoryID(i, 1) = "")) Then
                MsgBox "Your BU produce in " & FactoryID(i, 0) & " factories,please connect with QMS for set the " & FactoryID(i, 0) & " IP!"
                Exit Function
            End If
            strIP = Split(FactoryID(i, 1), ";")
            For j = 0 To UBound(strIP)
                If strIP(j) = Left(LocalIP, Len(strIP(j))) And Trim(strIP(j)) <> "" Then
                    If Trim(Factory) <> "" Then
                        MsgBox "Your IP " & LocalIP & " is exist in different factory,please connect with QMS check!"
                        Exit Function
                    Else
                        Factory = Trim(FactoryID(i, 0))
                        CreateDIDFlag = "Y"
                    End If
                End If
            Next j
        Next i
    Else
        Factory = Trim(Rs.Fields("Factory"))
        CreateDIDFlag = "Y"
    End If
    CheckFacIP = True
    ''''(0014)---------
End Function

Public Function SaveLog(System_Name As String, strIP As String, strUserID As String, StrEventDesc As String)
Dim Rs As New ADODB.Recordset
Dim strSQL As String
strSQL = "Insert into QMS_Log(System_Name,Event_No,SN,User_Name,Desc1,Trans_Date)" & _
            "Select '" & Trim(System_Name) & "','1','" & Trim(strIP) & "','" & Trim(strUserID) & "','('+Host_Name()+')" & Trim(StrEventDesc) & "',dbo.formatdate(getdate(),'YYYYMMDDHHNNSS')"
If Rs.State = 1 Then Rs.Close
Rs.CursorLocation = adUseClient
Set Rs = Conn.Execute(strSQL)
End Function

