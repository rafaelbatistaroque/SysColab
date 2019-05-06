VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_crudFerias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private ferias As clsFerias
Private feriasCRUD As clsFeriasCRUD
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
    Set feriasCRUD = New clsFeriasCRUD
    Set ferias = New clsFerias
    
    With ferias
        .SetId = Me.txtId
        .SetIdEstrangeira = Me.txtIdEstrangeira
        .SetSelecao = False
        .SetDtInicio = Me.txtDtInicio
        .SetDtFinal = Me.txtDtFinal
        .SetAnoRef = Me.txtAnoRef
        .SetOptante = Me.ckbOp
        .SetObsFerias = Me.txtObsFerias
    End With
    DoCmd.Close
    Call feriasCRUD.Salvar(ferias)
    
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
Public Sub Argumento(ferias As clsFerias)
    On Error GoTo TrataErro
    If Not IsEmpty(ferias.GetId) Then
        With ferias
                Me.txtIdEstrangeira = .IdEstrangeira
                Me.txtId = .GetId
                Me.ckbSelecao = .Selecao
                Me.txtDtInicio = .DtInicio
                Me.txtDtFinal = .DtFinal
                Me.txtAnoRef = .AnoRef
                Me.ckbOp = .Optante
                Me.txtObsFerias = .ObsFerias
                Me.rotArq.Caption = "Editar Férias"
        End With
    Else
        Me.txtIdEstrangeira = ferias.IdEstrangeira
        Me.rotArq.Caption = "Nova Férias"
    End If
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
