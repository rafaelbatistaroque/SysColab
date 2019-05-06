VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFeriasCRUD"
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
Private ferias As New clsFerias
Private strSQL As String
'================================
'MÉTODOS
'================================
'ABRE FORM PRA NOVO REGISTRO DE SALARIO
Public Sub NovaFerias(ferias As clsFerias)
    On Error GoTo TrataErro

    Static frm As Form 'declara
    Set frm = New Form_frm_crudFerias 'instancia
    Call frm.Argumento(ferias) 'chama a função interna do form e passa o Objeto ferias
    frm.Visible = True 'torna visível
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'ABRE FORM PARA EDITAR O REGISTRO DE FÉRIAS
Public Sub Editar()
    Set rs = New ADODB.Recordset
    'busca os registros no BD que está com checkbox TRUE
    strSQL = "SELECT * FROM tbl_Ferias WHERE SELECAO = -1"
    
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
    Set ferias = New clsFerias
    With ferias
        .SetIdEstrangeira = rs(0)
        .SetId = rs(1)
        .SetSelecao = rs(2)
        .SetDtInicio = rs(3)
        .SetDtFinal = rs(4)
        .SetAnoRef = rs(5)
        .SetOptante = rs(6)
        .SetObsFerias = rs(7)
    End With
    'Criar nova instancia de form para edição
    Call NovaFerias(ferias)
    
    rs.Close: Set rs = Nothing 'Esvazia memória
    modUtils.FecharConexao 'Fecha conexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'SALVA REGISTRO NOVO DE SALARIO
Public Sub Salvar(ferias As clsFerias)
    On Error GoTo TrataErro
    If Not IsNull(ferias.GetId) Then
        strSQL = "UPDATE tbl_Ferias SET ID_ESTRANGEIRA = ?, SELECAO = ?, DATA_INICIO = ?, " & _
                 "DATA_FINAL= ?, ANO_REFERENCIA = ?, OPTANTE = ?, OBS_FERIAS = ? WHERE ID = ?;"
    Else
        strSQL = "INSERT INTO tbl_Ferias (ID_ESTRANGEIRA, SELECAO, DATA_INICIO, DATA_FINAL, " & _
                 "ANO_REFERENCIA, OPTANTE, OBS_FERIAS)VALUES(?,?,?,?,?,?,?);"
    End If
    
    Set conexao = modUtils.AbrirConexao
    Set comando = modUtils.CriarComando(conexao, strSQL)
    With comando
        '.NamedParameters = True
        .Parameters.Append .CreateParameter("@aIdEst", adInteger, adParamInput, 10, ferias.IdEstrangeira)
        .Parameters.Append .CreateParameter("@aSelecao", adBoolean, adParamInput, 1, ferias.Selecao)
        .Parameters.Append .CreateParameter("@DtInicio", adDate, adParamInput, 10, ferias.DtInicio)
        .Parameters.Append .CreateParameter("@DtFinal", adDate, adParamInput, 10, ferias.DtFinal)
        .Parameters.Append .CreateParameter("@AnoRef", adVarChar, adParamInput, 30, ferias.AnoRef)
        .Parameters.Append .CreateParameter("@Op", adBoolean, adParamInput, 1, ferias.Optante)
        .Parameters.Append .CreateParameter("@ObsFerias", adLongVarChar, adParamInput, 255, ferias.ObsFerias)
        .Parameters.Append .CreateParameter("@aId", adInteger, adParamInput, 10, ferias.GetId)
        .Execute
    End With
    modUtils.FecharConexao
    MsgBox "Registro Salvo"
    Call modUtils.CarregaSubForm(1, "frm_Ferias")
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'================================
'FUNÇÕES
'================================
'OBTER REGISTROS DE CONTRATO
Public Function ObterFerias() As ADODB.Recordset
    On Error GoTo TrataErro
    strSQL = "SELECT * FROM tbl_Ferias WHERE ID_ESTRANGEIRA =" & Forms!frm_funcionarios.txtId
    
    Set rs = New ADODB.Recordset
    Set conexao = modUtils.AbrirConexao
    rs.Open strSQL, conexao, adOpenStatic, adLockOptimistic
    Set ObterFerias = rs 'Return
    Exit Function
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function


