VERSION 5.00
Begin VB.Form FrmStartSplitLineMC 
   Caption         =   "启动分仓"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdStartSplitLineMC 
      Caption         =   "启动"
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
    MsgBox ("已启动分仓!"), vbInformation
End Sub
