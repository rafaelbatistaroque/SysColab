VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_crudPrestadorSv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private prestadorSV As clsPrestadorDeSv
Private prestadorSvCRUD As clsPrestadorDeSvCRUD
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
    CarregarDescricao
    CarregarTipoContrato
End Sub
'FECHAR
Private Sub btnFechar_Click()
    DoCmd.Close
End Sub
'SALVAR
Private Sub btnSalvar_Click()
    On Error GoTo TrataErro
    Set prestadorSvCRUD = New clsPrestadorDeSvCRUD
    Set prestadorSV = New clsPrestadorDeSv
    
    If Not IsNull(Me.txtDescricao) Then
        With prestadorSV
            .SetId = Me.txtId
            .SetIdEstrangeira = Me.txtIdEstrangeira
            .SetSelecao = False
            .SetDescricao = Me.txtDescricao
            .SetContato = Me.txtContato
            .SetTpContrato = Me.txtTipoContrato
            .SetNrContrato = Me.txtNrContr
            .SetObsConta = Me.txtObs
            .SetLinkGlobal = Me.txtLink
        End With
        DoCmd.Close
        Call prestadorSvCRUD.Salvar(prestadorSV)
    Else
        MsgBox "Operação cancelada." & vbCrLf & _
                     "Campo 'Descrição' vazio", vbCritical

        Me.txtDescricao.SetFocus
    End If
    Exit Sub

TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'ATUALIZAR MASK DO NR DE CONTRATO
Private Sub txtTipoContrato_AfterUpdate()
    Select Case Me.txtTipoContrato
        Case "Contrato (Correios)"
            Me.txtNrContr.InputMask = "0000000000/0000;0;_"
        Case "UC (Energisa)"
            Me.txtNrContr.InputMask = "00/000000-0;0;_"
        Case "Matrícula (Águas Guariroba)"
            Me.txtNrContr.InputMask = "00000000-0;0;_"
        Case "Código do Cliente (Tim)"
            Me.txtNrContr.InputMask = "0\.0000000;0;_"
        Case "Código NET"
            Me.txtNrContr.InputMask = "000/000000000;0;_"
        Case "Código do Cliente (Vivo)"
            Me.txtNrContr.InputMask = "0000\ 0000\ 0000 \DV: 0;0;_"
        Case "Geral"
            Me.txtNrContr.InputMask = ""
    End Select
    Me.txtNrContr.SetFocus
End Sub
'CARREGA COMBOBOX TIPO CONTRATO
Private Sub CarregarTipoContrato()
    Dim strTipo() As String
    Dim i As Integer
    
    strTipo = Split("UC (Energisa);" & _
                          "Contrato (Correios);" & _
                          "Inscrição;" & _
                          "Matrícula (Águas Guariroba);" & _
                          "Código do Cliente (Tim);" & _
                          "Código do Cliente (Vivo);" & _
                          "Número do Cliente;" & _
                          "Código NET;" & _
                          "Geral", ";")
    
    For i = 0 To UBound(strTipo)
    Me.txtTipoContrato.AddItem (strTipo(i))
    Next
End Sub
'CARREGA DESCRIÇÃO
Private Sub CarregarDescricao()
    Dim strTipo() As String
    Dim i As Integer
    
    strTipo = Split("Advogados;" & _
                          "Águas Guariroba;" & _
                          "Aluguel;" & _
                          "Certificado Digital;" & _
                          "Contador;" & _
                          "Correios;" & _
                          "Serviços Gerais;" & _
                          "Energisa;" & _
                          "Gráfica Pontual;" & _
                          "LocaWeb;" & _
                          "NET;" & _
                          "Tim;" & _
                          "Única Computadores;" & _
                          "Vivo;" & _
                          "Transporte;" & _
                          "Serviços de Limpeza (Angélica)", ";")
    
    For i = 0 To UBound(strTipo)
    Me.txtDescricao.AddItem (strTipo(i))
    Next
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
Public Sub Argumento(Optional prestadorSV As clsPrestadorDeSv)
    On Error GoTo TrataErro
    If Not IsEmpty(prestadorSV.GetId) Then
        With prestadorSV
            Me.txtIdEstrangeira = .IdEstrangeira
            Me.txtId = .GetId
            Me.ckbSelecao = .Selecao
            Me.txtDescricao = .Descricao
            Me.txtContato = .Contato
            Me.txtTipoContrato = .TpContrato
            Me.txtNrContr = .NrContrato
            Me.txtObs = .ObsConta
            Me.txtLink = .LinkGlobal
            Me.rotArq.Caption = "Editar prestador de Sv"
        End With
    End If
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
