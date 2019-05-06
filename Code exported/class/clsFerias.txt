VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFerias"
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
Private cDtInicio As Variant
Private cDtFinal As Variant
Private cAnoRef As Variant
Private cOptante As Variant
Private cObsFerias As Variant
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
'DATA DE INÍCIO
Public Property Get DtInicio() As Variant
    DtInicio = cDtInicio
End Property
Public Property Let SetDtInicio(ByVal vDtInicio As Variant)
    cDtInicio = vDtInicio
End Property
'DATA FINAL
Public Property Get DtFinal() As Variant
    DtFinal = cDtFinal
End Property
Public Property Let SetDtFinal(ByVal vDtFinal As Variant)
    cDtFinal = vDtFinal
End Property
'ANO DE REFERÊNCIA
Public Property Get AnoRef() As Variant
    AnoRef = cAnoRef
End Property
Public Property Let SetAnoRef(ByVal vAnoRef As Variant)
    cAnoRef = vAnoRef
End Property
'SE GOZOU DA FÉRIAS OU VENDEU
Public Property Get Optante() As Variant
    Optante = cOptante
End Property
Public Property Let SetOptante(ByVal vOptante As Variant)
    cOptante = vOptante
End Property
'OBS FÉRIAS
Public Property Get ObsFerias() As Variant
    ObsFerias = cObsFerias
End Property
Public Property Let SetObsFerias(ByVal vObs As Variant)
    cObsFerias = vObs
End Property

