Attribute VB_Name = "modUtils"
Option Compare Database
Option Explicit
'================================
'VARIÁVEIS
'================================
Private conexao As ADODB.connection
Private comando As ADODB.Command
Private rs As ADODB.Recordset
Private conectionString As String
Private strCaminho As String
Private strSQL As String
'================================
'FUNÇÕES
'================================
'ABRIR CONEXAO
Public Function AbrirConexao() As ADODB.connection
    Set conexao = New ADODB.connection
    'String de conexão
    On Error GoTo TrataErro
    conectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & CurrentProject.Path & _
                    "\SYSCOLAB\bd\sysColab_be.accdb;Jet OLEDB:Database Password=010586120130135225"
    conexao.CursorLocation = adUseClient
    conexao.Open conectionString 'Abre conexão
    
    If conexao.State = adStateOpen Then 'Se conexão aberta, Return conexão
        Set AbrirConexao = conexao
    Else
        MsgBox "Sem conexão com Base de Dados"
    End If
    Exit Function

TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function
'OBTER COMANDO
Public Function CriarComando(connection As ADODB.connection, strSQL As String) As ADODB.Command
    Set comando = New ADODB.Command
    
    On Error GoTo TrataErro
        With comando
            .ActiveConnection = connection 'Habilita conexão
            .CommandText = strSQL
            .CommandType = adCmdText
        End With
    Set CriarComando = comando
    Exit Function
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Function
'LER COMANDO
Public Function GetReader(cmd As ADODB.Command) As ADODB.Recordset
    Set GetReader = cmd.Execute
End Function
'================================
'MÉTODOS
'================================
'FECHAR CONEXAO
Public Sub FecharConexao()
    On Error GoTo TrataErro
    If conexao.State = adStateOpen Then conexao.Close: Set conexao = Nothing
    Exit Sub

TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    Exit Sub
End Sub
'CARREGA SUBFORMULÁRIO NETO
Public Sub CarregaSubForm(escolha As Integer, subFrm As String)
    On Error GoTo TrataErro
    Select Case escolha
        Case 1
            Forms!frm_funcionarios.Form.formNeto.SourceObject = subFrm
        Case 2
            Forms!frm_Home.Form.formFilho.SourceObject = subFrm
        Case 3
            Forms!frm_Home.Form.formFilho.SourceObject = subFrm
        Case 4
            Forms!frm_PrestadorSv.Form.formNeto.SourceObject = subFrm
    End Select
    Exit Sub

TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'ICONES BOTÃO
Public Sub CarregarIcons(Optional bNovo As CommandButton, _
                         Optional bEdit As CommandButton, _
                         Optional bDel As CommandButton, _
                         Optional bSearch As CommandButton, _
                         Optional bGo As CommandButton, _
                         Optional bBack As CommandButton, _
                         Optional bSave As CommandButton, _
                         Optional bClose As CommandButton, _
                         Optional bMenuHome As CommandButton, _
                         Optional bMenuRh As CommandButton, _
                         Optional bMenuFin As CommandButton, _
                         Optional bMail As CommandButton)
On Error GoTo TrataErro
                        'Local icones botões
                        strCaminho = CurrentProject.Path & "\SYSCOLAB\icons\"
                        'Verifica se o parâmetro existe, se sim, carrega a imagem no botão
                        If Not bNovo Is Nothing Then bNovo.Picture = strCaminho & "novo.png"
                        If Not bEdit Is Nothing Then bEdit.Picture = strCaminho & "editar.png"
                        If Not bDel Is Nothing Then bDel.Picture = strCaminho & "deletar.png"
                        If Not bSearch Is Nothing Then bSearch.Picture = strCaminho & "pesquisar.png"
                        If Not bGo Is Nothing Then bGo.Picture = strCaminho & "ir.png"
                        If Not bBack Is Nothing Then bBack.Picture = strCaminho & "voltar.png"
                        If Not bSave Is Nothing Then bSave.Picture = strCaminho & "salvar.png"
                        If Not bClose Is Nothing Then bClose.Picture = strCaminho & "close.png"
                        If Not bMenuHome Is Nothing Then bMenuHome.Picture = strCaminho & "menu-home.png"
                        If Not bMenuRh Is Nothing Then bMenuRh.Picture = strCaminho & "menu-rh.png"
                        If Not bMenuFin Is Nothing Then bMenuFin.Picture = strCaminho & "menu-financeiro.png"
                        If Not bMail Is Nothing Then bMail.Picture = strCaminho & "mail.png"
                        Exit Sub
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'DELETAR REGISTRO
Public Sub Deletar(vTblName As String)
    Set rs = New ADODB.Recordset
    'busca os registros no BD que está com checkbox TRUE
    strSQL = "SELECT SELECAO FROM " & vTblName & " WHERE SELECAO = -1"
    
    On Error GoTo TrataErro
    'Abre conexao | Criar comando | executa comando
    Set conexao = modUtils.AbrirConexao
    Set comando = modUtils.CriarComando(conexao, strSQL)
    Set rs = modUtils.GetReader(comando)
    'se não encontrar registos para o ID, Lança msg e morre aqui
    If rs.RecordCount = 0 Then
        MsgBox ("Selecione o registro para excluir")
        Exit Sub
    Else
        'Confirma com o usuário antes de deletar
        If MsgBox(rs.RecordCount & " registro(s) para exclusão." & vbCrLf & "Você tem certeza?", vbYesNo + vbDefaultButton2) = vbYes Then
            strSQL = "DELETE * FROM " & vTblName & " WHERE SELECAO = -1"
            Set comando = modUtils.CriarComando(conexao, strSQL)
            comando.Execute
            MsgBox rs.RecordCount & " Registro(s) excluído"
        Else
            MsgBox "Exclusão cancelada!"
            End
        End If
    End If
    rs.Close: Set rs = Nothing 'Esvazia memória
    modUtils.FecharConexao 'Fecha conexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
'ENVIAR EMAIL
Public Sub EnviarEmail()
    Dim objOut As Outlook.Application
    Dim objMail As Outlook.MailItem
    Dim objConta As Outlook.Accounts
    Dim strMail As String
    
  'Consulta que seleciona a empresa e o funcionário
    strSQL = "SELECT tbl_Funcionarios.email, tbl_Arquivo.selecao, tbl_Arquivo.mes_referencia, tbl_Arquivo.tipo_arquivo, tbl_Arquivo.obs_Arquivo, tbl_Arquivo.link " & _
                  "FROM tbl_Funcionarios INNER JOIN tbl_Arquivo ON tbl_Funcionarios.id = tbl_Arquivo.id_estrangeira " & _
                  "WHERE tbl_Arquivo.selecao = -1;"
    
      'strSQL = "SELECT * FROM tbl_Arquivo WHERE SELECAO = -1"
    
    On Error GoTo TrataErro
    'Abre conexao | Criar comando | executa comando
    Set conexao = modUtils.AbrirConexao
    Set comando = modUtils.CriarComando(conexao, strSQL)
    Set rs = modUtils.GetReader(comando)
    
    rs.MoveFirst
    Do While Not rs.EOF
        '-------------------------------------------------
        'Ativa o outlook e o formulário de envio
        '-------------------------------------------------
        Set objOut = New Outlook.Application
        Set objMail = objOut.CreateItem(olMailItem)
        
        strMail = "<p><font size=1>*Mensagem automática*</font></p>"
        strMail = strMail & "<p>Referência: " & Format(rs(2), "##/##")
        strMail = strMail & "<br>Conteúdo: " & rs(4) & "</p>"
        strMail = strMail & "<p>Link: <a href='" & rs(5) & "'><button>Abrir</button></a></p>"
        
        With objMail
                .To = rs(0) 'Destinatário
                .Subject = "INFO RH: " & rs(3) 'Assunto
                .HTMLBody = strMail 'Corpo do Email
                .Send 'Enviar
        End With
    rs.MoveNext
    Loop
    MsgBox "Mensagem enviada...", vbInformation, "Aviso"
    
    strSQL = "UPDATE TBL_ARQUIVO SET SELECAO = 0 WHERE SELECAO = -1"
    Set comando = modUtils.CriarComando(conexao, strSQL)
    comando.Execute
    
sair:
    '------------------------
    'Limpa a memória
    '------------------------
    Set objMail = Nothing
    Set objOut = Nothing
    rs.Close: Set rs = Nothing 'Esvazia momória
    modUtils.FecharConexao 'Fecha conexao
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then Call modExceptions.TrataErro(Err)
    End
End Sub
