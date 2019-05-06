VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DashBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private funcionarioCRUD As clsFuncionarioCRUD
Private funcionario As clsFuncionario
Private rs As ADODB.Recordset
Private strCaminho As String

'================================
'FORMULÁRIO PRINCIPAL
'================================
Private Sub Form_Open(Cancel As Integer)
    Set rs = New ADODB.Recordset
    Me.Pesquisa = Empty
    Form_Resize
    Set rs = FiltrarFuncionario(1)
    Call CarregarForm(rs)
    Call modUtils.CarregarIcons(Me.btnCriar, Me.btnEditar, Me.btnDeletar, Me.btnFiltro, Me.txtBtnIr, , , , , , , Me.btnEmail)
End Sub
'SALVA O REGISTRO DA CHECKBOX
Private Sub ckbSelecao_Click()
    DoCmd.RunCommand acCmdSaveRecord
End Sub
'CRIAR NOVO FUNCIONÁRIOS
Private Sub btnCriar_Click()
    Set funcionarioCRUD = New clsFuncionarioCRUD
    funcionarioCRUD.NovoFuncionario
End Sub
'EDITAR FUNCIONÁRIO
Private Sub btnEditar_Click()
    Set funcionarioCRUD = New clsFuncionarioCRUD
    Call funcionarioCRUD.Editar
End Sub
'DELETAR REGISTROS
Private Sub btnDeletar_Click()
    Set rs = New ADODB.Recordset
    Call modUtils.Deletar("Tbl_Funcionarios")
    Set rs = FiltrarFuncionario(1)
    Call CarregarForm(rs)
End Sub
'ENVIAR EMAIL COM HOLERITE
Private Sub btnEmail_Click()
    modUtils.EnviarEmail
End Sub
'BOTÃO IR PARA FUNCIONÁRIO
Private Sub txtBtnIr_Click()
    DoCmd.OpenForm "frm_funcionarios", , , , , , Me.txtId
End Sub
'CARREGA FILTRO DEPOIS DO ENTER
Private Sub Pesquisa_AfterUpdate()
    btnFiltro_Click
End Sub
'ABRE LINK DA NUVEM
Private Sub txtLinkGlobal_Click()
    If Not IsNull(Me.txtLinkGlobal) Then
        Application.FollowHyperlink Me.txtLinkGlobal
    End If
End Sub
'BOTÃO DE FILTRO
Private Sub btnFiltro_Click()
    Set rs = New ADODB.Recordset
    If Not IsNull(Me.Pesquisa) Then
        Set rs = FiltrarFuncionario(2)
        Call CarregarForm(rs)
    Else
        Set rs = FiltrarFuncionario(1)
        Call CarregarForm(rs)
    End If
End Sub
'REDIMENSIONAR
Private Sub Form_Resize()
    Dim larguraForm As Double
    Dim alturaForm As Double
    Dim margem, borda As Integer
    
    larguraForm = Form.InsideWidth
    alturaForm = Form.InsideHeight
    margem = 300
    borda = 30
    
    'TEXTBOXS
    'Me.ckbSelecao.Left = margem
    Me.txtLinha107.Left = margem / 2
    Me.txtLinha107.Width = larguraForm - (margem)
    Me.txtNome.Width = larguraForm / 3.5
    Me.txtCpf.Width = larguraForm / 6.5
    Me.txtDataNasc.Width = larguraForm / 6.5
    Me.txtLinkGlobal.Width = larguraForm - (Me.txtLinkGlobal.Left + margem)
    
    'BOTÕES CRUD
    Me.btnDeletar.Left = margem
    
    'PESQUISA
    Me.btnFiltro.Left = larguraForm - Me.Pesquisa.Width - (btnFiltro.Width + margem)
    
    'AVISO
    Me.rotAviso.Width = larguraForm
End Sub
'================================
'FUNÇÕES
'================================
'PREENCHE FORMULÁRIO COM REGISTROS
Private Sub CarregarForm(vRs As ADODB.Recordset)
    On Error GoTo TrataErro
    'Se não houver registros, oculta os controles e mostra o aviso
    If vRs.RecordCount = 0 Then
        Set Me.Form.Recordset = Nothing 'nenhum registro
        Me.btnFiltro.SetFocus
        Me.rotAviso.Visible = True 'mostrar aviso
    Else
        Set Me.Form.Recordset = vRs 'carregar dados
        Me.rotAviso.Visible = False 'oculta aviso
    End If
        modUtils.FecharConexao
    Exit Sub

TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'FILTRO
Public Function FiltrarFuncionario(Optional escolha As Integer) As ADODB.Recordset
    Set funcionarioCRUD = New clsFuncionarioCRUD
    Set rs = New ADODB.Recordset
    Dim wh As String
    
    On Error GoTo TrataErro
    Select Case escolha
        Case 1
            Set rs = funcionarioCRUD.ObterFuncionario 'Armazena arquivos na variável rs; SEM FILTRO
            Me.Cx1 = rs.RecordCount 'Total de funcionários
        Case 2
            wh = " WHERE NOME like '%" & Me!Pesquisa & "%' OR CPF like '%" & Me!Pesquisa & "%' OR DATA_NASCIMENTO like '%" & Me!Pesquisa & "%'"
            Set rs = funcionarioCRUD.ObterFuncionario(wh) 'Armazena arquivos na variável rs; COM FILTRO
        Case Else
            MsgBox "Ops! Talvez tem algum problema com a busca"
        End Select
    Set FiltrarFuncionario = rs
    Exit Function

TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function
