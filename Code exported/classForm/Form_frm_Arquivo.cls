VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Arquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private arquivoCRUD As clsArquivoFuncionarioCRUD
Private rs As ADODB.Recordset
Private strCaminho As String
Private margem As Integer
Private larguraForm As Long
'================================
'FORMULÁRIO PRINCIPAL
'================================
Private Sub Form_Open(Cancel As Integer)
    Call modUtils.CarregarIcons(, , , Me.btnFiltro)
    Call Form_Resize
    
    Set rs = New ADODB.Recordset
    Set rs = FiltrarArquivo(1)
    Call CarregarForm(rs)
End Sub
'REDIMENSIONAR
Private Sub Form_Resize()
    margem = 600
    larguraForm = Form.InsideWidth
    Me.btnFiltro.Left = larguraForm - (Me.Pesquisa.Width + margem * 2)
    Me.txtLinha107.Width = larguraForm - margem
    Me.txtObs.Width = larguraForm - (Me.txtObs.Left + margem)
    Me.ckbSelecao.Width = larguraForm - margem
End Sub
'ABRE LINK DA NUVEM
Private Sub txtLink_Click()
    If Not IsNull(Me.txtLink) Then
        Application.FollowHyperlink Me.txtLink
    End If
End Sub
'BOTÃO DE FILTRO
Private Sub btnFiltro_Click()
    Set rs = New ADODB.Recordset
    If Not IsNull(Me.Pesquisa) Then
        Set rs = FiltrarArquivo(2)
        Call CarregarForm(rs)
    Else
        Forms!frm_funcionarios.txtFocus.SetFocus
        Set rs = FiltrarArquivo(1)
        Call CarregarForm(rs)
    End If
End Sub
'ENTER DA CAIXA DE PESQUISA
Private Sub Pesquisa_AfterUpdate()
    btnFiltro_Click
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
        Me.rotAviso.Visible = False 'oculta aviso
        Set Form.Recordset = vRs 'carrega dados
    End If
    modUtils.FecharConexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'FILTRO
Private Function FiltrarArquivo(Optional escolha As Integer) As ADODB.Recordset
    Set arquivoCRUD = New clsArquivoFuncionarioCRUD
    Set rs = New ADODB.Recordset
    Dim wh As String
    
    On Error GoTo TrataErro
    Select Case escolha
        Case 1
            Set rs = arquivoCRUD.ObterArquivo 'Armazena arquivos na variável rs; SEM FILTRO
        Case 2
            wh = " AND DATA_CADASTRO like '%" & Me!Pesquisa & "%' " & _
            "OR ID_ESTRANGEIRA=" & Forms!frm_funcionarios.txtId & " AND TIPO_ARQUIVO like '%" & Me!Pesquisa & "%' " & _
            "OR ID_ESTRANGEIRA=" & Forms!frm_funcionarios.txtId & " AND MES_REFERENCIA like '%" & Me!Pesquisa & "%' " & _
            "OR ID_ESTRANGEIRA=" & Forms!frm_funcionarios.txtId & " AND OBS_ARQUIVO like '%" & Me!Pesquisa & "%'"
            Set rs = arquivoCRUD.ObterArquivo(wh) 'Armazena arquivos na variável rs; COM FILTRO
        Case Else
            MsgBox "Ops! Talvez tem algum problema com a busca"
        End Select
    Set FiltrarArquivo = rs
    Exit Function

TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function
