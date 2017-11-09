Attribute VB_Name = "Mensagens"
Public Function Mensagem(M As Long)
    Select Case M
       Case Is = 1
          Mensagem = "Não autorizado. Por favor, realize a autenticação para utilizar o CapptaGpPlus"
       Case Is = 2
          Mensagem = "Uma reimpressão ou cancelamento foi executada dentro de uma sessão multi-cartões"
       Case Is = 3
          Mensagem = "O CapptaGpPlus esta sendo inicializado, tente novamente em alguns instantes"
       Case Is = 4
          Mensagem = "O formato da requisição recebida pelo CapptaGpPlus é inválido"
       Case Is = 5
          Mensagem = "Operação cancelada pelo operador"
       Case Is = 6
          Mensagem = "Pagamento não autorizado/pendente/não encontrado"
       Case Is = 7
          Mensagem = "Pagamento ou cancelamento negados pela rede adquirente"
       Case Is = 8
          Mensagem = "Ocorreu um erro interno no CapptaGpPlus"
       Case Is = 9
          Mensagem = "Ocorreu um erro na comunicação entre a CappAPI e o CapptaGpPlus"
       Case Is = 10
          Mensagem = "Não é possível realizar uma operação sem que se tenha finalizado o último pagamento"
    End Select
End Function

