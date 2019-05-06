VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Conta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private contaCRUD As clsContaCRUD
Private rs As ADODB.Recordset
Private margem As Integer
Private larguraForm As Long
'================================
'FORMULÁRIO PRINCIPAL
'================================
Private Sub Form_Open(Cancel As Integer)
    Call modUtils.CarregarIcons(, , , Me.btnFiltro)
    Call Form_Resize
    
    Set rs = New ADODB.Recordset
    Set rs = FiltroConta(1)
    Call CarregarForm(rs)
End Sub
'REDIMENSIONAR
Private Sub Form_Resize()
    margem = 600
    larguraForm = Form.InsideWidth
    Me.ckbSelecao.Width = larguraForm - margem
    Me.btnFiltro.Left = larguraForm - (Me.Pesquisa.Width + margem * 2)
    Me.txtLinha107.Width = larguraForm - margem
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
        Set rs = FiltroConta(2)
        Call CarregarForm(rs)
    Else
        Forms!frm_PrestadorSv.txtFocus.SetFocus
        Set rs = FiltroConta(1)
        Call CarregarForm(rs)
    End If
End Sub
'ENTER DA CAIXA DE PESQUISA
Private Sub Pesquisa_AfterUpdate()
    btnFiltro_Click
End Sub
'TXT STATUS
Private Sub txtSatus_Click()
    Forms!frm_PrestadorSv.txtFocus.SetFocus
End Sub
'TXT DATA VENCIMENTO
Private Sub txtDtVenc_Click()
    Forms!frm_PrestadorSv.txtFocus.SetFocus
End Sub
'================================
'FUNÇÕES
'================================
'PREENCHE FORMULÁRIO COM REGISTROS
Private Sub CarregarForm(vRs As ADODB.Recordset)
On Error GoTo TrataErro
 'Se não houver registros, oculta os controles e mostra o aviso
    If vRs.RecordCount = 0 Then
        Set Form.Recordset = Nothing
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
Private Function FiltroConta(Optional escolha As Integer) As ADODB.Recordset
    Set contaCRUD = New clsContaCRUD
    Set rs = New ADODB.Recordset
    Dim wh As String
    
    On Error GoTo TrataErro
    Select Case escolha
        Case 1
            Set rs = contaCRUD.ObterConta 'Armazena arquivos na variável rs; SEM FILTRO
        Case 2
            wh = " AND DTA_CHEGADA like '%" & Me!Pesquisa & "%' " & _
                 "OR ID_ESTRANGEIRA=" & Forms!frm_PrestadorSv.txtId & " AND MES_REFERENCIA like '%" & Me!Pesquisa & "%' " & _
                 "OR ID_ESTRANGEIRA=" & Forms!frm_PrestadorSv.txtId & " AND VALOR_CONTA like '%" & Me!Pesquisa & "%' " & _
                 "OR ID_ESTRANGEIRA=" & Forms!frm_PrestadorSv.txtId & " AND DTA_VENCIMENTO like '%" & Me!Pesquisa & "%' " & _
                 "OR ID_ESTRANGEIRA=" & Forms!frm_PrestadorSv.txtId & " AND STATUS like '%" & Me!Pesquisa & "%' " & _
                 "OR ID_ESTRANGEIRA=" & Forms!frm_PrestadorSv.txtId & " AND OBS_CONTA like '%" & Me!Pesquisa & "%'"
            Set rs = contaCRUD.ObterConta(wh)  'Armazena arquivos na variável rs; COM FILTRO
        Case Else
            MsgBox "Ops! Talvez tem algum problema com a busca"
        End Select
    Set FiltroConta = rs
    Exit Function

TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function
