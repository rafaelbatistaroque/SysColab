VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalarioCRUD"
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
Private salario As New clsSalario
Private strSQL As String
'================================
'MÉTODOS
'================================
'ABRE FORM PRA NOVO REGISTRO DE SALARIO
Public Sub NovoSalario(salario As clsSalario)
    On Error GoTo TrataErro
    
    Static frm As Form 'declara
    Set frm = New Form_frm_crudSalario 'instancia
    Call frm.Argumento(salario) 'chama a função interna do form e passa o Objeto contrato
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
    strSQL = "SELECT * FROM tbl_Salario WHERE SELECAO = -1"
    
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
    Set salario = New clsSalario
    With salario
        .SetIdEstrangeira = rs(0)
        .SetId = rs(1)
        .SetSelecao = rs(2)
        .SetValor = rs(3)
        .SetData = rs(4)
        .SetMotivo = rs(5)
        .SetPeriodo = rs(6)
        .SetObsSalario = rs(7)
    End With
    'Criar nova instancia de form para edição
    Call NovoSalario(salario)
    
    rs.Close: Set rs = Nothing 'Esvazia memória
    modUtils.FecharConexao 'Fecha conexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'SALVA REGISTRO NOVO DE SALARIO
Public Sub Salvar(salario As clsSalario)
    On Error GoTo TrataErro
    If Not IsNull(salario.GetId) Then
        strSQL = "UPDATE tbl_Salario SET ID_ESTRANGEIRA = ?, SELECAO = ?, VALOR = ?, " & _
                 "DATA = ?, MOTIVO = ?, PERIODO = ?, OBS_SALARIO = ? WHERE ID = ?;"
    Else
        strSQL = "INSERT INTO tbl_Salario (ID_ESTRANGEIRA, SELECAO, VALOR, DATA, " & _
                 "MOTIVO, PERIODO, OBS_SALARIO)VALUES(?,?,?,?,?,?,?);"
    End If
    
    Set conexao = modUtils.AbrirConexao
    Set comando = modUtils.CriarComando(conexao, strSQL)
    With comando
        .Parameters.Append .CreateParameter("aIdEst", adInteger, adParamInput, 10, salario.IdEstrangeira)
        .Parameters.Append .CreateParameter("aSelecao", adBoolean, adParamInput, 1, salario.Selecao)
        .Parameters.Append .CreateParameter("aValor", adCurrency, adParamInput, 10, salario.Valor)
        .Parameters.Append .CreateParameter("aData", adDate, adParamInput, 10, salario.Data)
        .Parameters.Append .CreateParameter("aMotivo", adVarChar, adParamInput, 30, salario.Motivo)
        .Parameters.Append .CreateParameter("aPeriodo", adVarChar, adParamInput, 15, salario.Periodo)
        .Parameters.Append .CreateParameter("aObs", adLongVarChar, adParamInput, 255, salario.ObsSalario)
        .Parameters.Append .CreateParameter("aId", adInteger, adParamInput, 10, salario.GetId)
        .Execute
    End With
    modUtils.FecharConexao
    MsgBox "Registro Salvo"
    Call modUtils.CarregaSubForm(1, "frm_Salario")
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'================================
'FUNÇÕES
'================================
'OBTER REGISTROS DE CONTRATO
Public Function ObterSalario() As ADODB.Recordset
    On Error GoTo TrataErro
    strSQL = "SELECT * FROM tbl_Salario WHERE ID_ESTRANGEIRA =" & Forms!frm_funcionarios.txtId
    
    Set rs = New ADODB.Recordset
    Set conexao = modUtils.AbrirConexao
    rs.Open strSQL, conexao, adOpenStatic, adLockOptimistic
    Set ObterSalario = rs 'Return
    Exit Function
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function
