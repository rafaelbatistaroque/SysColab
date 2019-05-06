Attribute VB_Name = "modExceptions"
Option Compare Database
Option Explicit

Public Function TrataErro(erro As ErrObject)
    Select Case erro
        Case 3021 'Nenhum registro dentro do RecordSet
            MsgBox "Ops! Alguns campos ficaram sem preenchimento para executar esta ação"
        Case 91 'Valor de variável não definida
            MsgBox "Ops! Operação não executada"
        Case -2147467259 'tabela do back-end aberta em modo Design
            MsgBox "Ops! Ocorreu um erro. Sigua as instruções: " & vbCrLf & _
                         "1 - Verifique se o banco de dados está aberto" & vbCrLf & _
                         "2 - Verifique se alguem está editando o banco de dados"
        Case Else
            MsgBox "Erro Nr: " & erro.Number & vbCrLf & "Descrição: " & erro.Description
    End Select
    erro.Clear 'limpa erro
End Function
