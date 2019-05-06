VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Salario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private salarioCRUD As clsSalarioCRUD
Private rs As ADODB.Recordset
Private ctl As control
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
Sub CarregarForm()
    Set salarioCRUD = New clsSalarioCRUD
    Set rs = New ADODB.Recordset
    
    On Error GoTo TrataErro
    Set rs = salarioCRUD.ObterSalario
    If rs.RecordCount = 0 Then
        Me.rotAviso.Visible = True
        Set Form.Recordset = Nothing
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
