Attribute VB_Name = "Mensagens"
Public Function Mensagem(M As Long)
    Select Case M
       Case Is = 1
          Mensagem = "N�o autorizado. Por favor, realize a autentica��o para utilizar o CapptaGpPlus"
       Case Is = 2
          Mensagem = "Uma reimpress�o ou cancelamento foi executada dentro de uma sess�o multi-cart�es"
       Case Is = 3
          Mensagem = "O CapptaGpPlus esta sendo inicializado, tente novamente em alguns instantes"
       Case Is = 4
          Mensagem = "O formato da requisi��o recebida pelo CapptaGpPlus � inv�lido"
       Case Is = 5
          Mensagem = "Opera��o cancelada pelo operador"
       Case Is = 6
          Mensagem = "Pagamento n�o autorizado/pendente/n�o encontrado"
       Case Is = 7
          Mensagem = "Pagamento ou cancelamento negados pela rede adquirente"
       Case Is = 8
          Mensagem = "Ocorreu um erro interno no CapptaGpPlus"
       Case Is = 9
          Mensagem = "Ocorreu um erro na comunica��o entre a CappAPI e o CapptaGpPlus"
       Case Is = 10
          Mensagem = "N�o � poss�vel realizar uma opera��o sem que se tenha finalizado o �ltimo pagamento"
    End Select
End Function

