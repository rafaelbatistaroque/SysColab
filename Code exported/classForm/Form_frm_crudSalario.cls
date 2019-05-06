VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_crudSalario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private salario As clsSalario
Private salarioCRUD As clsSalarioCRUD
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
    Set salarioCRUD = New clsSalarioCRUD
    Set salario = New clsSalario
    
    With salario
        .SetId = Me.txtId
        .SetIdEstrangeira = Me.txtIdEstrangeira
        .SetSelecao = False
        .SetValor = Me.txtValor
        .SetData = Me.txtData
        .SetMotivo = Me.txtMotivo
        .SetPeriodo = Me.txtPeriodo
        .SetObsSalario = Me.txtObs
    End With
    DoCmd.Close
    Call salarioCRUD.Salvar(salario)
    
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
Public Sub Argumento(salario As clsSalario)
    On Error GoTo TrataErro
    If Not IsEmpty(salario.GetId) Then
        With salario
                Me.txtIdEstrangeira = .IdEstrangeira
                Me.txtId = .GetId
                Me.ckbSelecao = .Selecao
                Me.txtValor = .Valor
                Me.txtData = .Data
                Me.txtMotivo = .Motivo
                Me.txtPeriodo = .Periodo
                Me.txtObs = .ObsSalario
                Me.rotArq.Caption = "Editar Salario"
        End With
    Else
        Me.txtIdEstrangeira = salario.IdEstrangeira
        Me.rotArq.Caption = "Novo Salario"
    End If
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
