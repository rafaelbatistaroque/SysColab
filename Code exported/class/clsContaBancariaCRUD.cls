VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsContaBancariaCRUD"
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
Private conta As clsContaBancaria
Private strSQL As String
'================================
'MÉTODOS
'================================
'ABRE FORM PRA NOVO REGISTRO DE CONTA
Public Sub NovaConta(conta As clsContaBancaria)
    On Error GoTo TrataErro
    Static frm As Form 'declara
    Set frm = New Form_frm_crudContaBancaria 'instancia
    Call frm.Argumento(conta) 'chama a função interna do form e passa o Objeto contrato
    frm.Visible = True 'torna visível
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'ABRE FORM PARA EDITAR O REGISTRO DE CONTA
Public Sub Editar()
    Set rs = New ADODB.Recordset
    'busca os registros no BD que está com checkbox TRUE
    strSQL = "SELECT * FROM tbl_ContaBancaria WHERE SELECAO = -1"
    
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
    Set conta = New clsContaBancaria
    With conta
        .SetIdEstrangeira = rs(0)
        .SetId = rs(1)
        .SetSelecao = rs(2)
        .SetBanco = rs(3)
        .SetAgencia = rs(4)
        .SetOperacao = rs(5)
        .SetTipoConta = rs(6)
        .SetContaNr = rs(7)
        .SetContaDV = rs(8)
        .SetFavorecido = rs(9)
        .SetObsConta = rs(10)
    End With
    'Criar nova instancia de form para edição
    Call NovaConta(conta)
    
    rs.Close: Set rs = Nothing 'Esvazia memória
    modUtils.FecharConexao 'Fecha conexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'SALVA NOVO REGISTRO DE CONTA
Public Sub Salvar(conta As clsContaBancaria)
    On Error GoTo TrataErro
    If Not IsNull(conta.GetId) Then
        strSQL = "UPDATE tbl_ContaBancaria SET ID_ESTRANGEIRA = ?, SELECAO = ?, BANCO = ?, " & _
                 "AGENCIA = ?, OPERACAO = ?, TIPO_CONTA = ?, CONTA = ?, CONTA_DIGITO = ?, " & _
                 " FAVORECIDO = ?, OBS_CONTA = ? WHERE ID = ?;"
    Else
        strSQL = "INSERT INTO tbl_ContaBancaria (ID_ESTRANGEIRA, SELECAO, BANCO, AGENCIA, " & _
                 "OPERACAO, TIPO_CONTA, CONTA, CONTA_DIGITO, FAVORECIDO, OBS_CONTA)" & _
                 "VALUES(?,?,?,?,?,?,?,?,?,?);"
    End If
    
    Set conexao = modUtils.AbrirConexao
    Set comando = modUtils.CriarComando(conexao, strSQL)
    With comando
        .Parameters.Append .CreateParameter("@aIdEst", adInteger, adParamInput, 10, conta.IdEstrangeira)
        .Parameters.Append .CreateParameter("@aSelecao", adBoolean, adParamInput, 1, conta.Selecao)
        .Parameters.Append .CreateParameter("@aBanco", adVarChar, adParamInput, 40, conta.Banco)
        .Parameters.Append .CreateParameter("@aAgencia", adVarChar, adParamInput, 15, conta.Agencia)
        .Parameters.Append .CreateParameter("@aOperacao", adVarChar, adParamInput, 4, conta.Operacao)
        .Parameters.Append .CreateParameter("@aTipoConta", adVarChar, adParamInput, 40, conta.TipoConta)
        .Parameters.Append .CreateParameter("@aContaNr", adVarChar, adParamInput, 15, conta.ContaNr)
        .Parameters.Append .CreateParameter("@aContaDV", adInteger, adParamInput, 2, conta.ContaDv)
        .Parameters.Append .CreateParameter("@aFavorecido", adVarChar, adParamInput, 50, conta.Favorecido)
        .Parameters.Append .CreateParameter("@aObsConta", adLongVarChar, adParamInput, 255, conta.ObsConta)
        .Parameters.Append .CreateParameter("@aId", adInteger, adParamInput, 10, conta.GetId)
        .Execute
    End With
    modUtils.FecharConexao
    MsgBox "Registro Salvo"
    Call modUtils.CarregaSubForm(1, "frm_ContaBancaria")
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'================================
'FUNÇÕES
'================================
'OBTER REGISTROS DE CONTRATO
Public Function ObterConta() As ADODB.Recordset
    On Error GoTo TrataErro
    strSQL = "SELECT * FROM tbl_ContaBancaria WHERE ID_ESTRANGEIRA =" & Forms!frm_funcionarios.txtId
    
    Set rs = New ADODB.Recordset
    Set conexao = modUtils.AbrirConexao
    rs.Open strSQL, conexao, adOpenStatic, adLockOptimistic
    Set ObterConta = rs 'Return
    Exit Function
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function




