VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsContratacaoCRUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private conexao As New ADODB.connection
Private comando As New ADODB.Command
Private rs As ADODB.Recordset
Private contrato As clsContratacao
Private strSQL As String
'================================
'MÉTODOS
'================================
'ABRE FORM PRA NOVO REGISTRO DE CONTRATO
Public Sub NovoContrato(contrato As clsContratacao)
    On Error GoTo TrataErro

    Static frm As Form 'declara
    Set frm = New Form_frm_crudContrato 'instancia
    Call frm.Argumento(contrato) 'chama a função interna do form e passa o Objeto contrato
    frm.Visible = True 'torna visível
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'ABRE FORM PARA EDITAR O REGISTRO DE SALÁRIO
Public Sub Editar()
    Set rs = New ADODB.Recordset
    'busca os registros no BD que está com checkbox TRUE
    strSQL = "SELECT * FROM tbl_Contratacao WHERE SELECAO = -1"
    
    On Error GoTo TrataErro
    'Abre conexao | Criar comando | executa comando
    Set conexao = modUtils.AbrirConexao
    Set comando = modUtils.CriarComando(conexao, strSQL)
    Set rs = modUtils.GetReader(comando)
    'se não encontrar registos para o ID, morre aqui
    If rs.RecordCount = 0 Then
        MsgBox ("Selecione um registro para editar")
        Exit Sub
    End If
    'Seta os registros do BD nas propriedades da classe
    Set contrato = New clsContratacao
    With contrato
        .SetIdEstrangeira = rs(0)
        .SetId = rs(1)
        .SetSelecao = rs(2)
        .SetEmpresa = rs(3)
        .SetCargo = rs(4)
        .SetDtAdmissao = rs(5)
        .SetDtDemissao = rs(6)
        .SetStatus = rs(7)
    End With
    'Criar nova instancia de form para edição
    Call NovoContrato(contrato)
    
    rs.Close: Set rs = Nothing 'Esvazia memória
    modUtils.FecharConexao 'Fecha conexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'SALVA REGISTRO NOVO DE CONTRATO
Public Sub Salvar(contrato As clsContratacao)
    On Error GoTo TrataErro
    If Not IsNull(contrato.GetId) Then
        strSQL = "UPDATE tbl_contratacao SET ID_ESTRANGEIRA = ?, SELECAO = ?, EMPRESA= ?, " & _
                 "CARGO = ?, DT_ADMISSAO = ?, DT_DEMISSAO = ?, STATUS = ? WHERE ID = ?;"
    Else
        strSQL = "INSERT INTO tbl_contratacao (ID_ESTRANGEIRA, SELECAO, EMPRESA, CARGO, " & _
                 "DT_ADMISSAO, DT_DEMISSAO, STATUS)VALUES(?,?,?,?,?,?,?);"
    End If
    
    Set conexao = modUtils.AbrirConexao
    Set comando = modUtils.CriarComando(conexao, strSQL)
    With comando
        .Parameters.Append .CreateParameter("@aIdEst", adInteger, adParamInput, 10, contrato.IdEstrangeira)
        .Parameters.Append .CreateParameter("@aSelecao", adBoolean, adParamInput, 1, contrato.Selecao)
        .Parameters.Append .CreateParameter("@aEmpresa", adVarChar, adParamInput, 70, contrato.Empresa)
        .Parameters.Append .CreateParameter("@aCargo", adVarChar, adParamInput, 50, contrato.Cargo)
        .Parameters.Append .CreateParameter("@aDtAdm", adDate, adParamInput, 10, contrato.DtAdmissao)
        .Parameters.Append .CreateParameter("@aDtDem", adDate, adParamInput, 10, contrato.DtDemissao)
        .Parameters.Append .CreateParameter("@aStatus", adVarChar, adParamInput, 10, contrato.Status)
        .Parameters.Append .CreateParameter("@aId", adInteger, adParamInput, 10, contrato.GetId)
        .Execute
    End With
    modUtils.FecharConexao
    MsgBox "Registro Salvo"
    Call modUtils.CarregaSubForm(1, "frm_Contratacao")
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'================================
'FUNÇÕES
'================================
'OBTER REGISTROS DE CONTRATO
Public Function ObterContrato() As ADODB.Recordset
    On Error GoTo TrataErro
    strSQL = "SELECT * FROM tbl_contratacao WHERE ID_ESTRANGEIRA =" & Forms!frm_funcionarios.txtId
    
    Set rs = New ADODB.Recordset
    Set conexao = modUtils.AbrirConexao
    rs.Open strSQL, conexao, adOpenStatic, adLockOptimistic
    Set ObterContrato = rs 'Return
    Exit Function
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function


