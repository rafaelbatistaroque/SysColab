VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnFechar_Click()
    DoCmd.Quit
End Sub

Private Sub btnRh_Click()
rotTop.Caption = "Recursos Humanos"
    Me.formFilho.SourceObject = "frm_DashBoard"
    Me.formFilho.SetFocus
End Sub

Private Sub btnFinaceiro_Click()
rotTop.Caption = "Financeiro"
formFilho.SourceObject = "frm_DashBoard_PrestadorSv"
Me.formFilho.SetFocus
End Sub

Private Sub btnHome_Click()
rotTop.Caption = "Home"
End Sub
Private Sub Form_Open(Cancel As Integer)
    Form_Resize
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    Call modUtils.CarregarIcons(, , , , , , , , Me.btnHome, Me.btnRh, Me.btnFinaceiro)
    focus.SetFocus
    DoCmd.Close acForm, Form(0).Name
    If IsNull(Me.OpenArgs) Then
    formFilho.SourceObject = "frm_DashBoard"
    Else
    formFilho.SourceObject = Me.OpenArgs
    End If
End Sub

Private Sub Form_Resize()
    Dim larguraForm As Double
    Dim alturaForm As Double
    Dim margem, borda As Integer
    
    margem = 600
    borda = 30
    larguraForm = Form.InsideWidth
    alturaForm = Form.InsideHeight
    
    DoCmd.Maximize
    
    'Caixa Top
    cxTop.Top = 0
    cxTop.Left = 0
    cxTop.Height = margem
    cxTop.Width = larguraForm
    
    'rodape
    rodape.Height = 150
    rodape.Top = alturaForm - rodape.Height
    rodape.Left = 0
    rodape.Width = larguraForm
    
    'Menu Lateral
    cxLateral.Width = margem
    cxLateral.Left = 0
    cxLateral.Height = alturaForm - cxTop.Height - rodape.Height
    cxLateral.Top = cxTop.Height
    
    btnHome.Left = 0
    btnHome.Width = cxLateral.Width
    btnHome.Top = cxTop.Height
    
    'Propriedades form filho
    formFilho.Width = larguraForm - cxLateral.Width
    formFilho.Height = alturaForm - cxTop.Height - rodape.Height
    formFilho.Top = cxTop.Height
    formFilho.Left = margem
    
    
End Sub
