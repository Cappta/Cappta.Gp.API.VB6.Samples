Attribute VB_Name = "MensagensPainel"
Public Function mensagem(resultado As Long) As String
    
    Select Case resultado
        Case 1
            mensagem = "N�o autorizado. Por favor, realize a autentica��o para utilizar o CapptaGpPlus"
        Case 2
            mensagem = "O CapptaGpPlus esta sendo inicializado, tente novamente em alguns instantes"
        Case 3
            mensagem = "O formato da requisi��o recebida pelo CapptaGpPlus � inv�lido"
        Case 4
            mensagem = "Opera��o cancelada pelo operador"
        Case 5
            mensagem = "Pagamento n�o autorizado/pendente/n�o encontrado"
        Case 6
            mensagem = "Pagamento ou cancelamento negados pela rede adquirente"
        Case 7
            mensagem = "Ocorreu um erro interno no CapptaGpPlus"
        Case 8
            mensagem = "Ocorreu um erro na comunica��o entre a CappAPI e o CapptaGpPlus"
        Case 9
            mensagem = "N�o � poss�vel realizar uma opera��o sem que se tenha finalizado o �ltimo pagamento"
        Case 10
            mensagem = "Uma reimpress�o ou cancelamento foi executada dentro de uma sess�o multi-cart�es"
        Case Else
            mensagem = "N�o foi poss�vel realizar a opera��o."
    End Select
    
End Function
