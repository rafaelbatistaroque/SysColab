VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuncionarioCRUD"
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
Private funcionario As clsFuncionario
Private rs As ADODB.Recordset
Private strSQL As String
'================================
'MÉTODOS
'================================
Public Sub NovoFuncionario()
    On Error GoTo TrataErro
    
    Static frm As Form 'declara
    Set frm = New Form_frm_crudFuncionario 'instancia
    frm.Visible = True 'funcionario visível
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'Atualiza registro de funcionários no Banco de Dados
Public Sub Salvar(funcionario As clsFuncionario)
    On Error GoTo TrataErro
    If Not IsNull(funcionario.GetId) Then
        strSQL = "UPDATE Tbl_Funcionarios SET ID_ESTRANGEIRA=?, SELECAO=?, NOME = ?, RG = ?, ORGAO_EMISSOR = ?, " & _
                      "CPF = ?, DATA_NASCIMENTO = ?, PIS = ?, CTPS_NR= ?, CTPS_SERIE = ?, CTPS_UF = ?, " & _
                      "OBS_FUNCIONARIO = ?, LINK_GLOBAL = ?, EMAIL= ? WHERE ID = ?;"
    Else
        strSQL = "INSERT INTO Tbl_Funcionarios (ID_ESTRANGEIRA, SELECAO, NOME, RG, ORGAO_EMISSOR, CPF, DATA_NASCIMENTO, PIS, " & _
                      "CTPS_NR, CTPS_SERIE, CTPS_UF, OBS_FUNCIONARIO, LINK_GLOBAL, EMAIL)VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?);"
    End If
    
    'Abre conexão
    Set conexao = modUtils.AbrirConexao
    'Comando
    Set comando = modUtils.CriarComando(conexao, strSQL)
    With comando
        '.NamedParameters = True
        .Parameters.Append .CreateParameter("@aIdEst", adInteger, adParamInput, 10, funcionario.IdEstrangeira)
        .Parameters.Append .CreateParameter("@aSelecao", adBoolean, adParamInput, 1, funcionario.Selecao)
        .Parameters.Append .CreateParameter("@aNome", adVarChar, adParamInput, 100, funcionario.Nome)
        .Parameters.Append .CreateParameter("@aRg", adVarChar, adParamInput, 15, funcionario.Rg)
        .Parameters.Append .CreateParameter("@aOrgEmissor", adVarChar, adParamInput, 10, funcionario.OEmissor)
        .Parameters.Append .CreateParameter("@aCPF", adVarChar, adParamInput, 15, funcionario.Cpf)
        .Parameters.Append .CreateParameter("@aDtNasc", adDate, adParamInput, 10, funcionario.DtNascimento)
        .Parameters.Append .CreateParameter("@aPis", adVarChar, adParamInput, 15, funcionario.Pis)
        .Parameters.Append .CreateParameter("@aCtpsNr", adVarChar, adParamInput, 10, funcionario.CtpsNr)
        .Parameters.Append .CreateParameter("@aCtpsSr", adVarChar, adParamInput, 10, funcionario.CtpsSr)
        .Parameters.Append .CreateParameter("@aCtpsUf", adVarChar, adParamInput, 2, funcionario.CtpsUF)
        .Parameters.Append .CreateParameter("@aObsFunc", adLongVarChar, adParamInput, 255, funcionario.Obs)
        .Parameters.Append .CreateParameter("@aLink", adVarChar, adParamInput, 100, funcionario.GetLink)
        .Parameters.Append .CreateParameter("@aEmail", adVarChar, adParamInput, 100, funcionario.GetEmail)
        .Parameters.Append .CreateParameter("@aId", adInteger, adParamInput, 10, funcionario.GetId)
        .Execute
    End With
    MsgBox "Registrado com Sucesso!"
    modUtils.FecharConexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'ABRE FORM PARA EDIÇÃO DE REGISTROS
Public Sub Editar()
    Set rs = New ADODB.Recordset
    'busca os registros no BD que está com checkbox TRUE
    strSQL = "SELECT * FROM Tbl_Funcionarios WHERE SELECAO = -1"
    
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
    Set funcionario = New clsFuncionario
    With funcionario
        .SetIdEstrangeira = rs(0)
        .SetId = rs(1)
        .SetSelecao = rs(2)
        .SetNome = rs(3)
        .SetRg = rs(4)
        .SetOEmissor = rs(5)
        .SetCpf = rs(6)
        .SetDtNascimento = rs(7)
        .SetPis = rs(8)
        .SetCtpsNr = rs(9)
        .SetCtpsSr = rs(10)
        .SetCtpsUF = rs(11)
        .SetObs = rs(12)
        .SetLink = rs(13)
        .SetEmail = rs(14)
    End With
    
    'Criar nova instancia de form para edição
    Static frm As Form
    Set frm = New Form_frm_crudFuncionario 'instancia
    Call frm.Argumento(funcionario) 'chama a função interna do form e passa o Objeto funcionario
    frm.Visible = True 'funcionario visível
    
    rs.Close: Set rs = Nothing 'Esvazia memória
    modUtils.FecharConexao 'Fecha conexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'================================
'FUNÇÕES
'================================
'OBTER FUNCIONÁRIO
Public Function ObterFuncionario(Optional filtro As String) As ADODB.Recordset
    On Error GoTo TrataErro
    strSQL = "SELECT * FROM Tbl_Funcionarios" & filtro & " ORDER BY NOME"
    
    Set rs = New ADODB.Recordset
    Set conexao = modUtils.AbrirConexao
    rs.Open strSQL, conexao, adOpenStatic, adLockOptimistic
    Set ObterFuncionario = rs
    Exit Function
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function
