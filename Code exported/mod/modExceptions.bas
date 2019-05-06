Attribute VB_Name = "modExceptions"
Option Compare Database
Option Explicit

Public Function TrataErro(erro As ErrObject)
    Select Case erro
        Case 3021 'Nenhum registro dentro do RecordSet
            MsgBox "Ops! Alguns campos ficaram sem preenchimento para executar esta a��o"
        Case 91 'Valor de vari�vel n�o definida
            MsgBox "Ops! Opera��o n�o executada"
        Case -2147467259 'tabela do back-end aberta em modo Design
            MsgBox "Ops! Ocorreu um erro. Sigua as instru��es: " & vbCrLf & _
                         "1 - Verifique se o banco de dados est� aberto" & vbCrLf & _
                         "2 - Verifique se alguem est� editando o banco de dados"
        Case Else
            MsgBox "Erro Nr: " & erro.Number & vbCrLf & "Descri��o: " & erro.Description
    End Select
    erro.Clear 'limpa erro
End Function
