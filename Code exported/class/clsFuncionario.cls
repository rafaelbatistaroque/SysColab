VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuncionario"
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
Private cNome As Variant
Private cRg As Variant
Private cOEmissor As Variant
Private cCpf As Variant
Private cDtNasc As Variant
Private cPis As Variant
Private cCtpsNr As Variant
Private cCtpsSr As Variant
Private cCtpsUF As Variant
Private cLink As Variant
Private cObs As Variant
Private cEmail As Variant

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
'NOME
Public Property Get Nome() As Variant
    Nome = cNome
End Property
Public Property Let SetNome(ByVal vNome As Variant)
    cNome = vNome
End Property
'RG
Public Property Get Rg() As Variant
    Rg = cRg
End Property
Public Property Let SetRg(ByVal vRg As Variant)
    cRg = vRg
End Property
'ORGÃO EMISSOR
Public Property Get OEmissor() As Variant
    OEmissor = cOEmissor
End Property
Public Property Let SetOEmissor(ByVal vOEmissor As Variant)
    cOEmissor = vOEmissor
End Property
'CPF
Public Property Get Cpf() As Variant
    Cpf = cCpf
End Property
Public Property Let SetCpf(ByVal vCpf As Variant)
    cCpf = vCpf
End Property
'DATA DE NASCIMENTO
Public Property Get DtNascimento() As Variant
    DtNascimento = cDtNasc
End Property
Public Property Let SetDtNascimento(ByVal vDtNasc As Variant)
    cDtNasc = vDtNasc
End Property
'PIS
Public Property Get Pis() As Variant
    Pis = cPis
End Property
Public Property Let SetPis(ByVal vPis As Variant)
    cPis = vPis
End Property
'CTPS NR
Public Property Get CtpsNr() As Variant
    CtpsNr = cCtpsNr
End Property
Public Property Let SetCtpsNr(ByVal vCtpsNr As Variant)
    cCtpsNr = vCtpsNr
End Property
'CTPS SR
Public Property Get CtpsSr() As Variant
    CtpsSr = cCtpsSr
End Property
Public Property Let SetCtpsSr(ByVal vCtpsSr As Variant)
    cCtpsSr = vCtpsSr
End Property
'CTPS UF
Public Property Get CtpsUF() As Variant
    CtpsUF = cCtpsUF
End Property
Public Property Let SetCtpsUF(ByVal vCtpsUF As Variant)
    cCtpsUF = vCtpsUF
End Property
'LINK
Public Property Get GetLink() As Variant
    GetLink = cLink
End Property
Public Property Let SetLink(ByVal vLink As Variant)
    cLink = vLink
End Property
'OBS
Public Property Get Obs() As Variant
    Obs = cObs
End Property
Public Property Let SetObs(ByVal vObs As Variant)
    cObs = vObs
End Property
'EMAIL
Public Property Get GetEmail() As Variant
    GetEmail = cEmail
End Property
Public Property Let SetEmail(ByVal vEmail As Variant)
    cEmail = vEmail
End Property
