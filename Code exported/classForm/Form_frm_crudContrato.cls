VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_crudContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private contrato As clsContratacao
Private contratoCRUD As clsContratacaoCRUD
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
    Set contratoCRUD = New clsContratacaoCRUD
    Set contrato = New clsContratacao
    
    With contrato
        .SetId = Me.txtId
        .SetIdEstrangeira = Me.txtIdEstrangeira
        .SetSelecao = False
        .SetEmpresa = Me.txtEmpresa
        .SetCargo = Me.txtCargo
        .SetDtAdmissao = Me.txtDtAdmissao
        .SetDtDemissao = Me.txtDtDemissao
        .SetStatus = Me.txtStatus
    End With
    DoCmd.Close
    Call contratoCRUD.Salvar(contrato)
    
End Sub
'DATA DE ADMISSÃO ALTERA STATUS
Private Sub txtDtAdmissao_AfterUpdate()
    If IsNull(txtDtDemissao) Then
        Me.txtStatus = "Ativo"
    Else
        Me.txtDtDemissao = Null
    End If
End Sub
'DATA DE DEMISSÃO ALTERA STATUS
Private Sub txtDtDemissao_AfterUpdate()
    If IsNull(txtDtAdmissao) Then
        MsgBox "Sem data de Admissão!", vbCritical
        Me.txtDtDemissao = Null
        txtDtAdmissao.SetFocus
        Me.txtStatus = "Ativo"
    Else
        Me.txtStatus = "Inativo"
    End If
End Sub
'MOVER FORM
Private Sub MoverForm()
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
'================================
'FUNÇÕES
'================================
'INSERE A CLASSE NOS CAMPOS
Public Sub Argumento(contrato As clsContratacao)
    On Error GoTo TrataErro
    If Not IsEmpty(contrato.GetId) Then
        With contrato
                Me.txtIdEstrangeira = .IdEstrangeira
                Me.txtId = .GetId
                Me.ckbSelecao = .Selecao
                Me.txtEmpresa = .Empresa
                Me.txtCargo = .Cargo
                Me.txtDtAdmissao = .DtAdmissao
                Me.txtDtDemissao = .DtDemissao
                Me.txtStatus = .Status
                Me.rotArq.Caption = "Editar Contrato"
        End With
    Else
        Me.txtIdEstrangeira = contrato.IdEstrangeira
        Me.rotArq.Caption = "Nova Contratação"
    End If
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
