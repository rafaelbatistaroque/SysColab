VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrestadorDeSv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'==============
'ATRIBUTOS
'==============
Private cIdEstrangeira As Variant
Private cId As Variant
Private cSelecao As Variant
Private cDescricao As Variant
Private cContato As Variant
Private cTipoContrato As Variant
Private cNrContrato As Variant
Private cObsConta As Variant
Private cLinkGlobal As Variant
'==============
'PROPRIEDADES
'==============
'ID ESTRANGEIRA
Public Property Get IdEstrangeira() As Variant
    IdEstrangeira = cIdEstrangeira
End Property
Public Property Let SetIdEstrangeira(ByVal vidEstrangeira As Variant)
    cIdEstrangeira = vidEstrangeira
End Property
'ID
Public Property Get GetId() As Variant
    GetId = cId
End Property
Public Property Let SetId(ByVal vId As Variant)
    cId = vId
End Property
'SELEÇÃO
Public Property Get Selecao() As Variant
    Selecao = cSelecao
End Property
Public Property Let SetSelecao(ByVal vSelecao As Variant)
    cSelecao = vSelecao
End Property
'DESCRIÇÃO
Public Property Get Descricao() As Variant
    Descricao = cDescricao
End Property
Public Property Let SetDescricao(ByVal vDescricao As Variant)
    cDescricao = vDescricao
End Property
'CONTATO
Public Property Get Contato() As Variant
    Contato = cContato
End Property
Public Property Let SetContato(ByVal vContato As Variant)
    cContato = vContato
End Property
'TIPO CONTRATO
Public Property Get TpContrato() As Variant
    TpContrato = cTipoContrato
End Property
Public Property Let SetTpContrato(ByVal vTpContrato As Variant)
    cTipoContrato = vTpContrato
End Property
'NR CONTRATO
Public Property Get NrContrato() As Variant
    NrContrato = cNrContrato
End Property
Public Property Let SetNrContrato(ByVal vNrContrato As Variant)
    cNrContrato = vNrContrato
End Property
'OBS CONTA
Public Property Get ObsConta() As Variant
    ObsConta = cObsConta
End Property
Public Property Let SetObsConta(ByVal vObsConta As Variant)
    cObsConta = vObsConta
End Property
'LINK GLOBAL
Public Property Get LinkGlobal() As Variant
    LinkGlobal = cLinkGlobal
End Property
Public Property Let SetLinkGlobal(ByVal vLinkGlobal As Variant)
    cLinkGlobal = vLinkGlobal
End Property
