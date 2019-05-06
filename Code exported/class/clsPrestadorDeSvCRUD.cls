VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrestadorDeSvCRUD"
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
Private prestadorSV As clsPrestadorDeSv
Private strSQL As String
'================================
'MÉTODOS
'================================
'ABRE FORM PRA NOVO REGISTRO DE CONTA
Public Sub NovaConta()
    On Error GoTo TrataErro
    
    Static frm As Form 'declara
    Set frm = New Form_frm_crudPrestadorSv 'instancia
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
    strSQL = "SELECT * FROM tbl_PrestadorDeSv WHERE SELECAO = -1"
    
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
    Set prestadorSV = New clsPrestadorDeSv
    With prestadorSV
        .SetIdEstrangeira = rs(0)
        .SetId = rs(1)
        .SetSelecao = rs(2)
        .SetDescricao = rs(3)
        .SetContato = rs(4)
        .SetTpContrato = rs(5)
        .SetNrContrato = rs(6)
        .SetObsConta = rs(7)
        .SetLinkGlobal = rs(8)
    End With
    
    'Criar nova instancia de form para edição
    Static frm As Form
    Set frm = New Form_frm_crudPrestadorSv 'instancia
    Call frm.Argumento(prestadorSV) 'chama a função interna do form e passa o Objeto funcionario
    frm.Visible = True 'funcionario visível
    Exit Sub
    
    rs.Close: Set rs = Nothing 'Esvazia memória
    modUtils.FecharConexao 'Fecha conexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'SALVA NOVO REGISTRO DE CONTA
Public Sub Salvar(prestadorSV As clsPrestadorDeSv)
    On Error GoTo TrataErro
    If Not IsNull(prestadorSV.GetId) Then
        strSQL = "UPDATE tbl_PrestadorDeSv SET ID_ESTRANGEIRA = ?, SELECAO = ?, DESCRICAO = ?, " & _
                 "CONTATO = ?, TIPO_CONTRATO = ?, NR_CONTRATO = ?, OBS_CONTA = ?, LINK_GLOBAL = ?" & _
                 " WHERE ID = ?;"
    Else
        strSQL = "INSERT INTO tbl_PrestadorDeSv(ID_ESTRANGEIRA, SELECAO, DESCRICAO, CONTATO, " & _
                 "TIPO_CONTRATO, NR_CONTRATO, OBS_CONTA, LINK_GLOBAL)" & _
                 "VALUES(?,?,?,?,?,?,?,?);"
    End If
    
    Set conexao = modUtils.AbrirConexao
    Set comando = modUtils.CriarComando(conexao, strSQL)
    With comando
        .Parameters.Append .CreateParameter("@aIdEst", adInteger, adParamInput, 10, prestadorSV.IdEstrangeira)
        .Parameters.Append .CreateParameter("@aSelecao", adBoolean, adParamInput, 1, prestadorSV.Selecao)
        .Parameters.Append .CreateParameter("@aDescr", adVarChar, adParamInput, 40, prestadorSV.Descricao)
        .Parameters.Append .CreateParameter("@aContato", adVarChar, adParamInput, 30, prestadorSV.Contato)
        .Parameters.Append .CreateParameter("@aTpContr", adVarChar, adParamInput, 40, prestadorSV.TpContrato)
        .Parameters.Append .CreateParameter("@aNrContr", adVarChar, adParamInput, 40, prestadorSV.NrContrato)
        .Parameters.Append .CreateParameter("@aObsConta", adLongVarChar, adParamInput, 255, prestadorSV.ObsConta)
        .Parameters.Append .CreateParameter("@aLink", adVarChar, adParamInput, 100, prestadorSV.LinkGlobal)
        .Parameters.Append .CreateParameter("@aId", adInteger, adParamInput, 10, prestadorSV.GetId)
        .Execute
    End With
    modUtils.FecharConexao
    MsgBox "Registro Salvo"
    Call modUtils.CarregaSubForm(3, "frm_DashBoard_PrestadorSv")
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'================================
'FUNÇÕES
'================================
'OBTER REGISTROS DE ARQUIVO
Public Function ObterPrestadorSv(Optional filtro As String) As ADODB.Recordset
    On Error GoTo TrataErro
    strSQL = "SELECT * FROM tbl_PrestadorDeSv" & filtro & " ORDER BY DESCRICAO DESC"
    
    Set rs = New ADODB.Recordset
    Set conexao = modUtils.AbrirConexao 'Abre conexao
    rs.Open strSQL, conexao, adOpenStatic, adLockOptimistic
    Set ObterPrestadorSv = rs 'Return
    Exit Function
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function
