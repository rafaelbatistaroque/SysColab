VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalario"
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
Private cValor As Variant
Private cData As Variant
Private cMotivo As Variant
Private cPeriodo As Variant
Private cObsSalario As Variant
'==============
'PROPRIEDADES
'==============
'ID estrangeira
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
'VALOR
Public Property Get Valor() As Variant
    Valor = cValor
End Property
Public Property Let SetValor(ByVal vValor As Variant)
    cValor = vValor
End Property
'DATA
Public Property Get Data() As Variant
    Data = cData
End Property
Public Property Let SetData(ByVal vData As Variant)
    cData = vData
End Property
'MOTIVO
Public Property Get Motivo() As Variant
    Motivo = cMotivo
End Property
Public Property Let SetMotivo(ByVal vMotivo As Variant)
    cMotivo = vMotivo
End Property
'PERÍODO
Public Property Get Periodo() As Variant
    Periodo = cPeriodo
End Property
Public Property Let SetPeriodo(ByVal vPeriodo As Variant)
    cPeriodo = vPeriodo
End Property
'OBS
Public Property Get ObsSalario() As Variant
    ObsSalario = cObsSalario
End Property
Public Property Let SetObsSalario(ByVal vObs As Variant)
    cObsSalario = vObs
End Property
