VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_funcionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private funcionario As clsFuncionario
Private funcionarioCRUD As clsFuncionarioCRUD
Private salarioCRUD As clsSalarioCRUD
Private salario As clsSalario
Private arquivo As clsArquivoFuncionario
Private arquivoCRUD As clsArquivoFuncionarioCRUD
Private contrato As clsContratacao
Private contratoCRUD As clsContratacaoCRUD
Private conta As clsContaBancaria
Private contaCRUD As clsContaBancariaCRUD
Private ferias As clsFerias
Private feriasCRUD As clsFeriasCRUD
Private btnMenu As control
Private strCaminho As String
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
    DoCmd.OpenForm "frm_Home", , , , , , "frm_DashBoard"
End Sub
'=====================================
'BOTÕES SUPERIOR MENU SUBFORMULÁRIOS
'=====================================
Private Sub btnMenuSalario_Click()
    'CARREGA SUBFORMULARIO SALARIO
    Call EnfaseMenu(btnMenuSalario)
    Call modUtils.CarregaSubForm(1, "frm_Salario")
End Sub
Private Sub btnMenuFerias_Click()
    'CARREGA SUBFORMULARIO FÉRIAS
    Call EnfaseMenu(btnMenuFerias)
    Call modUtils.CarregaSubForm(1, "frm_Ferias")
End Sub
Private Sub btnMenuArquivo_Click()
    'CARREGA SUBFORMULARIO ARQUIVO
    Call EnfaseMenu(btnMenuArquivo)
    Call modUtils.CarregaSubForm(1, "frm_Arquivo")
End Sub
Private Sub btnMenuContrato_Click()
    'CARREGA SUBFORMULARIO CONTRATO
    Call EnfaseMenu(btnMenuContrato)
    Call modUtils.CarregaSubForm(1, "frm_contratacao")
End Sub
Private Sub btnMenuContasBancarias_Click()
    'CARREGA SUBFORMULARIO CONTA BANCÁRIA
    Call EnfaseMenu(btnMenuContasBancarias)
    Call modUtils.CarregaSubForm(1, "frm_ContaBancaria")
End Sub
'==================================
'CRUD MENU INFERIOR SUBFORMULÁRIOS
'==================================
'CRIAR NOVO REGISTRO NO FORM NETO
Private Sub btnCriarRegistro_Click()
    Select Case Me.formNeto.SourceObject
        Case "frm_Salario"
            Set salarioCRUD = New clsSalarioCRUD
            Set salario = New clsSalario
            salario.SetIdEstrangeira = Me.txtId
           
            Call salarioCRUD.NovoSalario(salario)
        Case "frm_ContaBancaria"
            Set contaCRUD = New clsContaBancariaCRUD
            Set conta = New clsContaBancaria
            conta.SetIdEstrangeira = Me.txtId
            
            Call contaCRUD.NovaConta(conta)
        Case "frm_Contratacao"
            Set contratoCRUD = New clsContratacaoCRUD
            Set contrato = New clsContratacao
            contrato.SetIdEstrangeira = Me.txtId
            
            Call contratoCRUD.NovoContrato(contrato)
        Case "frm_Ferias"
            Set feriasCRUD = New clsFeriasCRUD
            Set ferias = New clsFerias
            ferias.SetIdEstrangeira = Me.txtId
            
            Call feriasCRUD.NovaFerias(ferias)
        Case "frm_Arquivo"
            Set arquivoCRUD = New clsArquivoFuncionarioCRUD
            Set arquivo = New clsArquivoFuncionario
            arquivo.SetIdEstrangeira = Me.txtId
            
            Call arquivoCRUD.NovoArquivo(arquivo)
        Case Else
            MsgBox ("Ops! Nenhuma ação encontrada")
    End Select
End Sub
'EDITAR REGISTRO NO FORM NETO
Private Sub btnEditarRegistro_Click()
    Select Case Me.formNeto.SourceObject
        Case "frm_Salario"
            Set salarioCRUD = New clsSalarioCRUD
            salarioCRUD.Editar
        Case "frm_ContaBancaria"
            Set contaCRUD = New clsContaBancariaCRUD
            contaCRUD.Editar
        Case "frm_contratacao"
            Set contratoCRUD = New clsContratacaoCRUD
            contratoCRUD.Editar
        Case "frm_Ferias"
            Set feriasCRUD = New clsFeriasCRUD
            feriasCRUD.Editar
        Case "frm_Arquivo"
            Set arquivoCRUD = New clsArquivoFuncionarioCRUD
            Call arquivoCRUD.Editar(Me.txtId)
        Case Else
            MsgBox ("Ops! Primeiro selecione um formulário")
    End Select
End Sub
'DELETAR REGISTRO NO FORM NETO
Private Sub btnDeletarRegistro_Click()
    Select Case Me.formNeto.SourceObject
        Case "frm_Salario"
            Call modUtils.Deletar("tbl_Salario")
            Call modUtils.CarregaSubForm(1, "frm_Salario")
        Case "frm_ContaBancaria"
            Call modUtils.Deletar("tbl_ContaBancaria")
            Call modUtils.CarregaSubForm(1, "frm_ContaBancaria")
        Case "frm_Contratacao"
            Call modUtils.Deletar("tbl_contratacao")
            Call modUtils.CarregaSubForm(1, "frm_Contratacao")
        Case "frm_Ferias"
            Call modUtils.Deletar("tbl_Ferias")
            Call modUtils.Deletar("frm_Ferias")
        Case "frm_Arquivo"
            Call modUtils.Deletar("tbl_Arquivo")
            Call modUtils.Deletar("frm_Arquivo")
        Case Else
            MsgBox ("Ops! Primeiro ocorreu om problema ao carregar os dados")
    End Select
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
    formNeto.Width = larguraForm - (formNeto.Left + margem) - cxDados.Width
    formNeto.Height = alturaForm - (formNeto.Top + Section(1).Height) - margem
    
    'caixa de dados
    cxDados.Left = formNeto.Width + margem
    rotRg.Left = cxDados.Left + (margem / 2)
End Sub
'================================
'FUNÇOES
'================================
'ÊNFASE DO SUBMENU
Sub EnfaseMenu(btn As CommandButton)
    Me.cxMenu.Left = btn.Left
End Sub
'CLIQUE DO LINK DE TODOS OS ARQUIVOS DO FUNCIONÁRIO
Private Sub txtLink_Global_Click()
    On Error GoTo TrataErro
    If Not IsNull(Me.txtLink_Global) Then
        Application.FollowHyperlink Me.txtLink_Global
    End If
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'INSERE A OS DODOS NOS CAMPOS
Public Sub Argumento()
    On Error GoTo TrataErro
    If Not IsEmpty(Me.OpenArgs) Then
        strSQL = " WHERE ID=" & Me.OpenArgs 'Recebe argumento do "form_DashBoard"
        
        Set funcionarioCRUD = New clsFuncionarioCRUD
        Set rs = New ADODB.Recordset
        Set rs = funcionarioCRUD.ObterFuncionario(strSQL) 'Filtra Funcionário conforme Id
        
        With rs
            Me.txtId = rs(1)
            Me.txtNome = rs(3)
            Me.txtRg = rs(4) & " | " & rs(5)
            Me.txtCpf = rs(6)
            Me.txtData_Nascimento = rs(7)
            Me.txtPis = rs(8)
            Me.txtCtps = "Nr: " & rs(9) & " | Sr: " & rs(10) & " | UF: " & rs(11)
            Me.txtObs_Funcionario = "<font color= #266EAA><b>Obs: </b></font color>" & rs(12)
            Me.txtLink_Global = rs(13)
            Me.txtEmail = rs(14)
        End With
        
        Call modUtils.CarregaSubForm(1, "frm_Salario")
        Call EnfaseMenu(btnMenuSalario)
    End If
    rs.Close: Set rs = Nothing
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
