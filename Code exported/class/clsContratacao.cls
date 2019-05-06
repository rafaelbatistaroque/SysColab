VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsContratacao"
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
Private cEmpresa As Variant
Private cCargo As Variant
Private cDtAdmissao As Variant
Private cDtDemissao As Variant
Private cStatus As Variant
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
'EMPRESA
Public Property Get Empresa() As Variant
    Empresa = cEmpresa
End Property
Public Property Let SetEmpresa(ByVal vEmpresa As Variant)
    cEmpresa = vEmpresa
End Property
'CARGO
Public Property Get Cargo() As Variant
    Cargo = cCargo
End Property
Public Property Let SetCargo(ByVal vCargo As Variant)
    cCargo = vCargo
End Property
'DATA ADMISSÃO
Public Property Get DtAdmissao() As Variant
    DtAdmissao = cDtAdmissao
End Property
Public Property Let SetDtAdmissao(ByVal vDtAdmissao As Variant)
    cDtAdmissao = vDtAdmissao
End Property
'DATA DEMISSÃO
Public Property Get DtDemissao() As Variant
    DtDemissao = cDtDemissao
End Property
Public Property Let SetDtDemissao(ByVal vDtDemissao As Variant)
    cDtDemissao = vDtDemissao
End Property
'STATUS
Public Property Get Status() As Variant
    Status = cStatus
End Property
Public Property Let SetStatus(ByVal vStatus As Variant)
    cStatus = vStatus
End Property

