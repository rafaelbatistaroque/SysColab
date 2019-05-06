VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_PrestadorSv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private conta As clsConta
Private contaCRUD As clsContaCRUD
Private prestadorCRUD As clsPrestadorDeSvCRUD
Private strSQL As String
Private rs As ADODB.Recordset
'================================
'FORMULÁRIO PRINCIPAL
'================================
Private Sub Form_Open(Cancel As Integer)
    DoCmd.Close acForm, "frm_Home"
    Call modUtils.CarregarIcons(Me.btnCriarRegistro, Me.btnEditarRegistro, Me.btnDeletarRegistro, , , Me.btnVoltar)
    Call Argumento
    Call Form_Resize
End Sub
'VOLTAR
Private Sub btnVoltar_Click()
    DoCmd.Close
    DoCmd.OpenForm "frm_Home", , , , , , "frm_DashBoard_PrestadorSv"
End Sub
'==================================
'CRUD MENU SUBFORMULÁRIOS
'==================================
'CRIAR NOVO REGISTRO NO FORM NETO
Private Sub btnCriarRegistro_Click()
     Set contaCRUD = New clsContaCRUD
     Set conta = New clsConta
     conta.SetIdEstrangeira = Me.txtId
    
     Call contaCRUD.NovaConta(conta)
End Sub
'EDITAR REGISTRO NO FORM NETO
Private Sub btnEditarRegistro_Click()
    Set contaCRUD = New clsContaCRUD
    contaCRUD.Editar
End Sub
'DELETAR REGISTRO NO FORM NETO
Private Sub btnDeletarRegistro_Click()
    Call modUtils.Deletar("tbl_contas")
    Call modUtils.CarregaSubForm(4, "frm_Conta")
End Sub
'================================
'REDIMENSIONAMENTO DE FORMULÁRIO
'================================
Private Sub Form_Resize()
    Dim larguraForm As Double
    Dim alturaForm As Double
    Dim margem, borda As Integer
    Dim divisor As Double
    
    DoCmd.Maximize
    
    larguraForm = Me.WindowWidth 'Form.InsideWidth
    alturaForm = Me.WindowHeight 'Form.InsideHeight
    margem = 300
    borda = 30
   
    'Formulario neto
    formNeto.Width = larguraForm - (formNeto.Left + margem)
    formNeto.Height = alturaForm - (formNeto.Top + Section(1).Height) - margem
End Sub
'================================
'FUNÇOES
'================================
'INSERE A OS DODOS NOS CAMPOS
Public Sub Argumento()
    On Error GoTo TrataErro
    If Not IsEmpty(Me.OpenArgs) Then
        strSQL = " WHERE ID=" & Me.OpenArgs 'Recebe argumento do "form_DashBoard_Contas"
        
        Set prestadorCRUD = New clsPrestadorDeSvCRUD
        Set rs = New ADODB.Recordset
        Set rs = prestadorCRUD.ObterPrestadorSv(strSQL) 'Filtra Conta conforme Id
        
        With rs
            Me.txtId = rs(1)
            Me.txtDescricao = rs(3)
        End With
        
        Call modUtils.CarregaSubForm(4, "frm_Conta")
    End If
    rs.Close: Set rs = Nothing
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
