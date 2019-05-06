VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_crudFuncionario"
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
Private strCaminho As String
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
'ABRIR
Private Sub Form_Open(Cancel As Integer)
    Call modUtils.CarregarIcons(, , , , , , Me.btnSalvar, Me.btnFechar)
End Sub
'FECHAR
Private Sub btnFechar_Click()
    DoCmd.Close
End Sub
'SALVAR
Private Sub btnSalvar_Click()
    Set funcionarioCRUD = New clsFuncionarioCRUD
    Set funcionario = New clsFuncionario

    With funcionario
        .SetIdEstrangeira = Me.txtIdEstrangeira
        .SetId = Me.txtId
        .SetSelecao = False
        .SetNome = Me.txtNome
        .SetRg = Me.txtRg
        .SetOEmissor = Me.txtOemissor
        .SetCpf = Me.txtCpf
        .SetDtNascimento = Me.txtDtNasc
        .SetPis = Me.txtPis
        .SetCtpsNr = Me.txtCtps
        .SetCtpsSr = Me.txtCtpsSr
        .SetCtpsUF = Me.txtCtpsUf
        .SetLink = Me.txtLink
        .SetObs = Me.txtObs
        .SetEmail = Me.txtEmail
    End With
    DoCmd.Close
    Call funcionarioCRUD.Salvar(funcionario)
    
    Call modUtils.CarregaSubForm(2, "frm_DashBoard")
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
Public Sub Argumento(Optional funcionario As clsFuncionario)
    On Error GoTo TrataErro
    If Not IsEmpty(funcionario.GetId) Then
        With funcionario
                Me.txtIdEstrangeira = .IdEstrangeira
                Me.txtId = .GetId
                Me.ckbSelecao = .Selecao
                Me.txtNome = .Nome
                Me.txtRg = .Rg
                Me.txtOemissor = .OEmissor
                Me.txtCpf = .Cpf
                Me.txtDtNasc = .DtNascimento
                Me.txtPis = .Pis
                Me.txtCtps = .CtpsNr
                Me.txtCtpsSr = .CtpsSr
                Me.txtCtpsUf = .CtpsUF
                Me.txtObs = .Obs
                Me.txtLink = .GetLink
                Me.txtEmail = .GetEmail
                Me.rotArq.Caption = "Editar Funcionário"
        End With
    End If
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub

