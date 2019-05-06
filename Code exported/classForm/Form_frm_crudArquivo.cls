VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_crudArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private arquivo As clsArquivoFuncionario
Private arquivoCRUD As clsArquivoFuncionarioCRUD
'================================
'CONTROLES
'================================
'MOVER FORM
Private Sub CabeçalhoDoFormulário_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoverForm
End Sub
Private Sub rotArq_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoverForm
End Sub
Private Sub Form_Open(Cancel As Integer)
    Call modUtils.CarregarIcons(, , , , , , Me.btnSalvar, Me.btnFechar)
End Sub
'FECHAR
Private Sub btnFechar_Click()
    DoCmd.Close
End Sub
'SALVAR
Private Sub btnSalvar_Click()
    Set arquivoCRUD = New clsArquivoFuncionarioCRUD
    Set arquivo = New clsArquivoFuncionario
    
    With arquivo
        .SetId = Me.txtId
        .SetIdEstrangeira = Me.txtIdEstrangeira
        .SetSelecao = False
        .SetDataCadastro = Me.txtDataCad
        .SetTipoArquivo = Me.txtTipo
        .SetReferencia = Me.txtMesRef
        .SetObsArquivo = Me.txtObs
        .SetLink = Me.txtLink
    End With
    DoCmd.Close
    Call arquivoCRUD.Salvar(arquivo)
    
End Sub
'================================
'FUNÇÕES
'================================
'MOVER FORM
Private Sub MoverForm()
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
'INSERE A CLASSE NOS CAMPOS
Public Sub Argumento(arquivo As clsArquivoFuncionario)
    On Error GoTo TrataErro
    If Not IsEmpty(arquivo.GetId) Then
        With arquivo
                Me.txtIdEstrangeira = .IdEstrangeira
                Me.txtId = .GetId
                Me.ckbSelecao = .Selecao
                Me.txtDataCad = .DataCadastro
                Me.txtMesRef = .Referencia
                Me.txtTipo = .TipoArquivo
                Me.txtObs = .ObsArquivo
                Me.txtLink = .GetLink
                Me.rotArq.Caption = "Editar Arquivo"
        End With
    Else
        Me.txtIdEstrangeira = arquivo.IdEstrangeira
        Me.rotArq.Caption = "Novo Arquivo"
    End If
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'APAGAR ESPAÇOS DEPOIS DO ";"
Private Sub txtObs_AfterUpdate()
    If Not IsNull(txtObs) Then txtObs = Replace(txtObs, "; ", ";")
End Sub
