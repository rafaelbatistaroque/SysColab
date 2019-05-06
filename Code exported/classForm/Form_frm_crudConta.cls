VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_crudConta"
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
    On Error GoTo TrataErro
    Set contaCRUD = New clsContaCRUD
    Set conta = New clsConta
    
    If Not IsNull(Me.txtDtChegada) Or Not IsNull(Me.txtValor) Then
        With conta
            .SetId = Me.txtId
            .SetIdEstrangeira = Me.txtIdEstrangeira
            .SetSelecao = False
            .SetDtChegada = Me.txtDtChegada
            .SetMesRef = Me.txtMesRef
            .SetValorConta = Me.txtValor
            .SetDtVencimento = Me.txtDtVencimento
            .SetStatus = Me.txtStatus
            .SetLinkConta = Me.txtLink
            .SetObsConta = Me.txtObs
        End With
        DoCmd.Close
        Call contaCRUD.Salvar(conta)
    Else
        MsgBox "Operação cancelada." & vbCrLf & _
              "Campo Dta Chegada ou Valor vazio"
    End If
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
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
Public Sub Argumento(conta As clsConta)
    On Error GoTo TrataErro
    If Not IsEmpty(conta.GetId) Then
        With conta
                Me.txtIdEstrangeira = .IdEstrangeira
                Me.txtId = .GetId
                Me.ckbSelecao = .Selecao
                Me.txtDtChegada = .DtChegada
                Me.txtMesRef = .MesRef
                Me.txtValor = .ValorConta
                Me.txtDtVencimento = .DtVencimento
                Me.txtStatus = .Status
                Me.txtObs = .ObsConta
                Me.txtLink = .LinkConta
                Me.rotArq.Caption = "Editar Conta"
        End With
    Else
        Me.txtIdEstrangeira = conta.IdEstrangeira
        Me.rotArq.Caption = "Nova Conta"
    End If
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub


