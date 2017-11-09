Attribute VB_Name = "MensagensPainel"
Public Function mensagem(resultado As Long) As String
    
    Select Case resultado
        Case 1
            mensagem = "Não autorizado. Por favor, realize a autenticação para utilizar o CapptaGpPlus"
        Case 2
            mensagem = "O CapptaGpPlus esta sendo inicializado, tente novamente em alguns instantes"
        Case 3
            mensagem = "O formato da requisição recebida pelo CapptaGpPlus é inválido"
        Case 4
            mensagem = "Operação cancelada pelo operador"
        Case 5
            mensagem = "Pagamento não autorizado/pendente/não encontrado"
        Case 6
            mensagem = "Pagamento ou cancelamento negados pela rede adquirente"
        Case 7
            mensagem = "Ocorreu um erro interno no CapptaGpPlus"
        Case 8
            mensagem = "Ocorreu um erro na comunicação entre a CappAPI e o CapptaGpPlus"
        Case 9
            mensagem = "Não é possível realizar uma operação sem que se tenha finalizado o último pagamento"
        Case 10
            mensagem = "Uma reimpressão ou cancelamento foi executada dentro de uma sessão multi-cartões"
        Case Else
            mensagem = "Não foi possível realizar a operação."
    End Select
    
End Function
