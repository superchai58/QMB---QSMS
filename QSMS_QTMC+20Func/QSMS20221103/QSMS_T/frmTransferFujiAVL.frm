VERSION 5.00
Begin VB.Form frmTransferFujiAVL 
   Caption         =   "Transfer Replace PN data to FUJI AVL"
   ClientHeight    =   1110
   ClientLeft      =   3510
   ClientTop       =   4095
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtWO 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Transfer"
      Height          =   435
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Work Order:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmTransferFujiAVL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdTransfer_Click()
Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim PCBGroup As String
On Error GoTo ErrHdl:
strSQL = "select * from sap_wo_list where wo='" & Trim(txtWO) & "'"
Set rs = Conn.Execute(strSQL)
If rs.EOF Then
    MsgBox "Can't find this wo in SF system!", vbCritical
    txtWO.SetFocus
    Exit Sub
Else
    PCBGroup = Trim(rs!Group)
    strSQL = "select * from sap_wo_list where [group]='" & PCBGroup & "' and status<>'20' "
    Set rs = Conn.Execute(strSQL)
    If rs.EOF = False Then
        MsgBox "There was some wo didn't check bom pass in this PCB Group", vbCritical
        txtWO.SetFocus
        Exit Sub
    End If
End If
If MsgBox("Are you sure to transfer replace pn data to FUJI AVLList?", vbOKCancel) = vbCancel Then Exit Sub
strSQL = "exec TransferReplacePNToFujiAVLByPCBGroup '" & PCBGroup & "'"
Set rs = Conn.Execute(strSQL)
If rs.EOF = False Then
    MsgBox "Please check the wrong comp in excel", vbCritical
    Call CopyToExcel(rs)
Else
    MsgBox "Transfer ok"
End If

Exit Sub
ErrHdl:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
Dim strSQL As String
Dim rs As New ADODB.Recordset
strSQL = "select "
End Sub
