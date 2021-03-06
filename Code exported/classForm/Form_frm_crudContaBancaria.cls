VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_crudContaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARI�VEIS
'================================
Private conta As clsContaBancaria
Private contaCRUD As clsContaBancariaCRUD
'================================
'CONTROLES
'================================
'MOVER FORM
Private Sub Cabe�alhoDoFormul�rio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    Set contaCRUD = New clsContaBancariaCRUD
    Set conta = New clsContaBancaria
    
    With conta
        .SetId = Me.txtId
        .SetIdEstrangeira = Me.txtIdEstrangeira
        .SetSelecao = False
        .SetBanco = Me.txtBanco
        .SetAgencia = Me.txtAgencia
        .SetOperacao = Me.txtOperacao
        .SetTipoConta = Me.txtTipoCont
        .SetContaNr = Me.txtContaNr
        .SetContaDV = Me.txtContaDV
        .SetFavorecido = Me.txtFavorecido
        .SetObsConta = Me.txtObsConta
    End With
    DoCmd.Close
    Call contaCRUD.Salvar(conta)
End Sub
'MOVER FORM
Private Sub MoverForm()
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
'================================
'FUN��ES
'================================
'INSERE A CLASSE NOS CAMPOS
Public Sub Argumento(conta As clsContaBancaria)
    On Error GoTo TrataErro
    If Not IsEmpty(conta.GetId) Then
        With conta
            Me.txtIdEstrangeira = .IdEstrangeira
            Me.txtId = .GetId
            Me.ckbSelecao = .Selecao
            Me.txtBanco = .Banco
            Me.txtAgencia = .Agencia
            Me.txtOperacao = .Operacao
            Me.txtTipoCont = .TipoConta
            Me.txtContaNr = .ContaNr
            Me.txtContaDV = .ContaDv
            Me.txtFavorecido = .Favorecido
            Me.txtObsConta = .ObsConta
            Me.rotArq.Caption = "Editar Conta Banc�ria"
        End With
    Else
        Me.txtIdEstrangeira = conta.IdEstrangeira
        Me.rotArq.Caption = "Nova Conta Banc�ria"
    End If
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'ATUALIZA COMPO COM TIPO DE CONTA
Private Sub txtOperacao_AfterUpdate()
Select Case Me.txtOperacao
    Case "001"
        Me.txtTipoCont = "Corrente de Pessoa F�sica"
    Case "002"
        Me.txtTipoCont = "Conta Simples de Pessoa F�sica"
    Case "003"
        Me.txtTipoCont = "Conta Corrente de Pessoa Jur�dica"
    Case "006"
        Me.txtTipoCont = "Entidades P�blicas"
    Case "007"
        Me.txtTipoCont = "Dep�sitos Institui��es Financeiras"
    Case "013"
        Me.txtTipoCont = "Poupan�a de Pessoa F�sica"
    Case "022"
        Me.txtTipoCont = "Poupan�a de Pessoa Jur�dica"
    Case "028"
        Me.txtTipoCont = "Poupan�a de Cr�dito Imobili�rio"
    Case "043"
        Me.txtTipoCont = "Dep�sitos Lot�ricos"
    Case Else
        MsgBox "Opera��o n�o encontrada"
End Select
End Sub
