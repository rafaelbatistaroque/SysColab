VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConta"
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
Private cDtChegada As Variant
Private cTipoDocumento As Variant
Private cMesRef As Variant
Private cValorConta As Variant
Private cDtVencimento As Variant
Private cStatus As Variant
Private cObsConta As Variant
Private cLinkConta As Variant
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
'DATA CHEGADA
Public Property Get DtChegada() As Variant
    DtChegada = cDtChegada
End Property
Public Property Let SetDtChegada(ByVal vDtChegada As Variant)
    cDtChegada = vDtChegada
End Property
'TIPO DE DOCUMENTO
Public Property Get TipoDoc() As Variant
    TipoDoc = cTipoDocumento
End Property
Public Property Let SetTipoDoc(ByVal vTipoDoc As Variant)
    cTipoDocumento = vTipoDoc
End Property
'MÊS REFERÊNCIA
Public Property Get MesRef() As Variant
    MesRef = cMesRef
End Property
Public Property Let SetMesRef(ByVal vMesRef As Variant)
    cMesRef = vMesRef
End Property
'VALOR
Public Property Get ValorConta() As Variant
    ValorConta = cValorConta
End Property
Public Property Let SetValorConta(ByVal vValorConta As Variant)
    cValorConta = vValorConta
End Property
'DATA DE VENCIMENTO
Public Property Get DtVencimento() As Variant
    DtVencimento = cDtVencimento
End Property
Public Property Let SetDtVencimento(ByVal vDtVencimento As Variant)
    cDtVencimento = vDtVencimento
End Property
'STATUS
Public Property Get Status() As Variant
    Status = cStatus
End Property
Public Property Let SetStatus(ByVal vStatus As Variant)
    cStatus = vStatus
End Property
'OBS
Public Property Get ObsConta() As Variant
    ObsConta = cObsConta
End Property
Public Property Let SetObsConta(ByVal vObsConta As Variant)
    cObsConta = vObsConta
End Property
'LINK CONTA
Public Property Get LinkConta() As Variant
    LinkConta = cLinkConta
End Property
Public Property Let SetLinkConta(ByVal vLinkConta As Variant)
    cLinkConta = vLinkConta
End Property
