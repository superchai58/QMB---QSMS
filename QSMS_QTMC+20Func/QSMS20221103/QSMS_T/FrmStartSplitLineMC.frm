VERSION 5.00
Begin VB.Form FrmStartSplitLineMC 
   Caption         =   "�����ֲ�"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdStartSplitLineMC 
      Caption         =   "����"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "FrmStartSplitLineMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdStartSplitLineMC_Click()  '1181
Dim strSql As String
Dim rs As ADODB.Recordset
    strSql = "Exec QSMS_SplitLineMC '" & Trim(g_userName) & "'"
    Set rs = Conn.Execute(strSql)
    MsgBox ("�������ֲ�!"), vbInformation
End Sub
