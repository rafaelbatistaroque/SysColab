VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArquivoFuncionarioCRUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'================================
'VARI�VEIS
'================================
Private conexao As New ADODB.connection
Private comando As New ADODB.Command
Private arquivo As clsArquivoFuncionario
Private rs As ADODB.Recordset
Private strSQL As String
'================================
'M�TODOS
'================================
'ABRE FORM PRA NOVO REGISTRO DE ARQUIVO
Public Sub NovoArquivo(arquivo As clsArquivoFuncionario)
    On Error GoTo TrataErro
    
    Static frm As Form 'declara
    Set frm = New Form_frm_crudArquivo 'instancia
    Call frm.Argumento(arquivo) 'chama a fun��o interna do form e passa o Objeto Arquivo
    frm.Visible = True 'torna vis�vel
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'ABRE FORM PARA EDITAR O REGISTRO DE SAL�RIO
Public Sub Editar(vId As Integer)
    Set rs = New ADODB.Recordset
    'busca os registros no BD que est� com checkbox TRUE
    strSQL = "SELECT * FROM tbl_Arquivo WHERE SELECAO = -1 AND ID_ESTRANGEIRA = " & vId
    
    On Error GoTo TrataErro
    'Abre conexao | Criar comando | executa comando
    Set conexao = modUtils.AbrirConexao
    Set comando = modUtils.CriarComando(conexao, strSQL)
    Set rs = modUtils.GetReader(comando)
    'se n�o encontrar registos para o ID, morre aqui
    If rs.RecordCount = 0 Then
        MsgBox ("Selecione um registro para editar")
        Exit Sub
    End If
    'Seta os registros do BD nas propriedades da classe
    Set arquivo = New clsArquivoFuncionario
    With arquivo
        .SetIdEstrangeira = rs(0)
        .SetId = rs(1)
        .SetSelecao = rs(2)
        .SetDataCadastro = rs(3)
        .SetTipoArquivo = rs(4)
        .SetReferencia = rs(5)
        .SetObsArquivo = rs(6)
        .SetLink = rs(7)
    End With
    'Criar nova instancia de form para edi��o
    Call NovoArquivo(arquivo)
    
    rs.Close: Set rs = Nothing 'Esvazia mem�ria
    modUtils.FecharConexao 'Fecha conexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'SALVA REGISTRO NOVO DE ARQUIVO
Public Sub Salvar(arquivo As clsArquivoFuncionario)
    On Error GoTo TrataErro
    If Not IsNull(arquivo.GetId) Then
        strSQL = "UPDATE tbl_Arquivo SET ID_ESTRANGEIRA = ?, SELECAO = ?, DATA_CADASTRO = ?, " & _
                 "TIPO_ARQUIVO = ?, MES_REFERENCIA = ?, OBS_ARQUIVO = ?, LINK = ? WHERE ID = ?;"
    Else
        strSQL = "INSERT INTO tbl_Arquivo (ID_ESTRANGEIRA, SELECAO, DATA_CADASTRO, TIPO_ARQUIVO, " & _
                 "MES_REFERENCIA, OBS_ARQUIVO, LINK)VALUES(?,?,?,?,?,?,?);"
    End If
    
    Set conexao = modUtils.AbrirConexao
    Set comando = modUtils.CriarComando(conexao, strSQL)
    With comando
        .Parameters.Append .CreateParameter("@aIdEst", adInteger, adParamInput, 10, arquivo.IdEstrangeira)
        .Parameters.Append .CreateParameter("@aSelecao", adBoolean, adParamInput, 1, arquivo.Selecao)
        .Parameters.Append .CreateParameter("@aDtCadastro", adDate, adParamInput, 10, arquivo.DataCadastro)
        .Parameters.Append .CreateParameter("@aTipoArq", adVarChar, adParamInput, 50, arquivo.TipoArquivo)
        .Parameters.Append .CreateParameter("@aMesRef", adVarChar, adParamInput, 10, arquivo.Referencia)
        .Parameters.Append .CreateParameter("@aObs", adLongVarChar, adParamInput, 255, arquivo.ObsArquivo)
        .Parameters.Append .CreateParameter("@aLink", adVarChar, adParamInput, 100, arquivo.GetLink)
        .Parameters.Append .CreateParameter("@aId", adInteger, adParamInput, 10, arquivo.GetId)
        .Execute
    End With
    modUtils.FecharConexao
    MsgBox "Registro Salvo"
    
    Call modUtils.CarregaSubForm(1, "frm_Arquivo")
    Call CriarDiretorio(arquivo)
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'================================
'FUN��ES
'================================
'OBTER REGISTROS DE ARQUIVO
Public Function ObterArquivo(Optional filtro As String) As ADODB.Recordset
    On Error GoTo TrataErro
    strSQL = "SELECT * FROM tbl_Arquivo WHERE ID_ESTRANGEIRA=" & Forms!frm_funcionarios.txtId & filtro & " ORDER BY DATA_CADASTRO DESC"
    
    Set rs = New ADODB.Recordset
    Set conexao = modUtils.AbrirConexao 'Abre conexao
    rs.Open strSQL, conexao, adOpenStatic, adLockOptimistic
    Set ObterArquivo = rs 'Return
    Exit Function
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function
'CRIAR DIRET�RIO
Private Sub CriarDiretorio(arquivo As clsArquivoFuncionario)
    Dim strDestino As String
    Dim strSplit As Variant
    Dim strPastas As Variant
    Dim i As Integer
    
    On Error GoTo TrataErro
    'Consulta que seleciona a empresa e o funcion�rio
    strSQL = "SELECT tbl_Contratacao.empresa, tbl_Contratacao.status, tbl_Funcionarios.nome,  tbl_Contratacao.id_estrangeira " & _
                  "FROM tbl_Contratacao  INNER JOIN  tbl_Funcionarios ON Tbl_Funcionarios.id = tbl_Contratacao.id_estrangeira " & _
                  "WHERE tbl_Contratacao.id_estrangeira =" & arquivo.IdEstrangeira & " AND tbl_Contratacao.status ='Ativo';" 'tbl_Contratacao.id_estrangeira =" & arquivo.IdEstrangeira & " AND"
    
    Set rs = New ADODB.Recordset 'Instancia recordSet
    Set conexao = modUtils.AbrirConexao '
    Set comando = modUtils.CriarComando(conexao, strSQL)
    Set rs = modUtils.GetReader(comando)
    
    strPastas = rs(0) & "|" & rs(2) & "|" & arquivo.TipoArquivo & "|" & "Ref " & Format(arquivo.Referencia, "00-00")
    'strDestino = "C:\Users\Vostro 5480\Desktop\"
    strDestino = CurrentProject.Path & "\SYSCOLAB\ARQUIVO\Funcion�rios\"
    
    'verifica se existe a pasta, caso sim, sobrescreve-a
    On Error Resume Next
    strSplit = Split(strPastas, "|") 'Pega os valores divididos por |
    i = 0
    Do While i <= UBound(strSplit) 'Enquanto for menor que tamanho do Split
        strDestino = strDestino & strSplit(i) & "\"
        If Len(Dir(strDestino, vbDirectory) & "") = 0 Then FileSystem.MkDir (strDestino) 'Cria diret�rio
        i = i + 1
    Loop
    rs.Close: Set rs = Nothing 'Esvazia mom�ria
    modUtils.FecharConexao 'Fecha conexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
