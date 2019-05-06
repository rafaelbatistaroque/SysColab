VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArquivoFuncionario"
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
Private cDataCad As Variant
Private cTipoArquivo As Variant
Private cReferencia As Variant
Private cObsAquivo As Variant
Private cLink As Variant
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
'DATA DE CADASTRO
Public Property Get DataCadastro() As Variant
    DataCadastro = cDataCad
End Property
Public Property Let SetDataCadastro(ByVal vDataCad As Variant)
    cDataCad = vDataCad
End Property
'TIPO DE ARQUIVO
Public Property Get TipoArquivo() As Variant
    TipoArquivo = cTipoArquivo
End Property
Public Property Let SetTipoArquivo(ByVal vTipoArq As Variant)
    cTipoArquivo = vTipoArq
End Property
'MÊS DE REFERÊNCIA
Public Property Get Referencia() As Variant
    Referencia = cReferencia
End Property
Public Property Let SetReferencia(ByVal vReferencia As Variant)
    cReferencia = vReferencia
End Property
'OBS ARQUIVO
Public Property Get ObsArquivo() As Variant
    ObsArquivo = cObsAquivo
End Property
Public Property Let SetObsArquivo(ByVal vObsArquivo As Variant)
    cObsAquivo = vObsArquivo
End Property
'LINK
Public Property Get GetLink() As Variant
    GetLink = cLink
End Property
Public Property Let SetLink(ByVal vLink As Variant)
    cLink = vLink
End Property
