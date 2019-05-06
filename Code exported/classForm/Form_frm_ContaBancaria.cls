VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_ContaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private contaCRUD As clsContaBancariaCRUD
Private rs As ADODB.Recordset
'================================
'FORMULÁRIO PRINCIPAL
'================================
Private Sub Form_Open(Cancel As Integer)
    CarregarForm
End Sub
'================================
'FUNÇÕES
'================================
'PREENCHE FORMULÁRIO COM REGISTROS
Private Sub CarregarForm()
    Set contaCRUD = New clsContaBancariaCRUD
    Set rs = New ADODB.Recordset
    
    On Error GoTo TrataErro
    Set rs = contaCRUD.ObterConta
    If rs.RecordCount = 0 Then
        Set Form.Recordset = Nothing
        Me.txtFocus.SetFocus
        Me.rotAviso.Visible = True 'mostrar aviso
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
