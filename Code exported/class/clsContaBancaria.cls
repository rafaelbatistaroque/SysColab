VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsContaBancaria"
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
Private cBanco As Variant
Private cAgencia As Variant
Private cOperacao As Variant
Private cTipoConta As Variant
Private cContaNr As Variant
Private cContaDV As Variant
Private cFavorecido As Variant
Private cObs As Variant
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
''BANCO
Public Property Get Banco() As Variant
    Banco = cBanco
End Property
Public Property Let SetBanco(ByVal vBanco As Variant)
    cBanco = vBanco
End Property
'AGÊNCIA
Public Property Get Agencia() As Variant
    Agencia = cAgencia
End Property
Public Property Let SetAgencia(ByVal vAgencia As Variant)
    cAgencia = vAgencia
End Property
'OPERAÇÃO
Public Property Get Operacao() As Variant
    Operacao = cOperacao
End Property
Public Property Let SetOperacao(ByVal vOperacao As Variant)
    cOperacao = vOperacao
End Property
'TIPO DA CONTA
Public Property Get TipoConta() As Variant
    TipoConta = cTipoConta
End Property
Public Property Let SetTipoConta(ByVal vTipoConta As Variant)
    cTipoConta = vTipoConta
End Property
'NÚMERO DA CONTA
Public Property Get ContaNr() As Variant
    ContaNr = cContaNr
End Property
Public Property Let SetContaNr(ByVal vContaNr As Variant)
    cContaNr = vContaNr
End Property
'DÍGITO VERIFICADOR
Public Property Get ContaDv() As Variant
    ContaDv = cContaDV
End Property
Public Property Let SetContaDV(ByVal vContaDV As Variant)
    cContaDV = vContaDV
End Property
'FAVORECIDO
Public Property Get Favorecido() As Variant
    Favorecido = cFavorecido
End Property
Public Property Let SetFavorecido(ByVal vFavorecido As Variant)
    cFavorecido = vFavorecido
End Property
'OBS CONTA
Public Property Get ObsConta() As Variant
    ObsConta = cObs
End Property
Public Property Let SetObsConta(ByVal vObsConta As Variant)
    cObs = vObsConta
End Property
