VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_contratacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARI�VEIS
'================================
Private contratoCRUD As clsContratacaoCRUD
Private rs As ADODB.Recordset
'================================
'FORMUL�RIO PRINCIPAL
'================================
Private Sub Form_Open(Cancel As Integer)
    CarregarForm
End Sub
'================================
'FUN��ES
'================================
'PREENCHE FORMUL�RIO COM REGISTROS
Private Sub CarregarForm()
    Set contratoCRUD = New clsContratacaoCRUD
    Set rs = New ADODB.Recordset
    
    On Error GoTo TrataErro
    Set rs = contratoCRUD.ObterContrato
    If rs.RecordCount = 0 Then
        Set Form.Recordset = Nothing
        Me.rotAviso.Visible = True
    Else
        Me.rotAviso.Visible = False
        Set Form.Recordset = rs
    End If
    modUtils.FecharConexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
