Integração C#
A Dll da Cappta foi desenvolvida utilizando as melhores práticas de programação e desenvolvimento de software. Utilizamos o padrão COM pensando justamente na integração entre aplicações construídas em várias linguagens.

Obs: Durante a instalação do CapptaGpPlus o mesmo encarrega-se de registrar a DLL em seu computador.

<h1> Etapa 1 </h1>

<h4>Tempo estimado de 01:00 hora</h4>

A primeira etapa consiste na importação do componente (dll) para dentro do projeto.
Com a sua IDE aberta va em Project e depois References e adicione o CapptaGpPlus

A primeira função a ser utilizada é AutenticarPdv().

Para autenticar é necessário os seguintes dados: CNPJ, PDV e chave de autenticação, estes dados são os mesmos fornecidos durante a instalação do GP.

Chave: 795180024C04479982560F61B3C2C06E 

Os dados para autenticar aqui serão colocados inline, indicamos sempre colocar estes dados de forma configuravel para o usuário, para que o mesmo ou algum tecnico tenha facil acesso na alteração do CNPJ e PDV a serem utilizados.

```javascript
Public Sub AutenticarPDV()
    
  Dim resultadoAutenticacao As Long
    
    resultadoAutenticacao = cappta.AutenticarPDV(CNPJ, NumeroPDV, ChavePDV)
    iniciouTef = True
    If resultadoAutenticacao = 0 Then
            
        Exit Sub
    End If
    
    MsgBox (MensagensPainel.mensagem(resultadoAutenticacao))
   
End Sub

```

O resultado para autenticação com sucesso é: 0


<h1>Primeiro esforço.</h1>

Toda vez que realizar uma ação com o GP, vai perceber que ele começa a exibir o código 2 para autenticação, não se preocupe é assim mesmo, para recuperar os estados do GP, vamos direto para a etapa 3.


<h1>Etapa 2 </h1>

Tempo estimado de 00:30 minutos

Temos duas formas de integração, a visivel, onde a interação com o usuário fica por conta da Cappta, e a invisivel onde o form pode ser personalizado.

```javascript
Private Sub ConfigurarModoIntegracao(exibirGp As Boolean)
    
    Dim configs As New Configuracoes
    configs.ExibirInterface = exibirGp
    
    Dim result As Long
    result = cappta.Configurar(configs)
    
    If result <> 0 Then
        CriarMensagem (MensagensPainel.mensagem(result))
    End If
    
End Sub
```
<h1>Etapa 3</h1>

<h4>Tempo estimado de 01:00 hora</h4>

Conforme mencionado acima a Iteração Tef é muito importante para o perfeito funcionamento da integração, toda as ações de venda e administrativas passam por esta função.

```javascript
Public Sub IterarOperacaoTef(objCappta As ClienteCappta)
 
 If OptionUsarMultiTef.Value Then
    DesabilitarControlesMultiTef
 End If
 DesabilitarBotoes

 Dim iteracaoTef As Cappta_Gp_Api_Com.IIteracaoTef

 Do
 
    Set iteracaoTef = objCappta.IterarOperacaoTef()

    If TypeOf iteracaoTef Is IMensagem Then
        Call ExibirMensagem(iteracaoTef)
        Sleep INTERVALO_MILISEGUNDOS
    End If

    If TypeOf iteracaoTef Is IRequisicaoParametro Then
        Call RequisitarParametros(iteracaoTef, objCappta)
    End If

    If TypeOf iteracaoTef Is IRespostaTransacaoPendente Then
        Call ResolverTransacaoPendente(iteracaoTef, objCappta)
    End If

    If TypeOf iteracaoTef Is IRespostaOperacaoRecusada Then
        Call ExibirDadosOperacaoRecusada(iteracaoTef)
    End If

    If TypeOf iteracaoTef Is IRespostaOperacaoAprovada Then
        Call ExibirDadosOperacaoAprovada(iteracaoTef)
        Call FinalizarPagamento(objCappta)
    End If

  Loop While OperacaoNaoFinalizada(iteracaoTef)
  
  If sessaoMultiTefEmAndamento = False Then
    HabilitarControlesMultiTef
  End If
  HabilitarBotoes
  

End Sub

```

Dentro de IterarOperacaoTef() temos alguns métodos:

<h1>Requisitar Parametros</h1>

```javascript

Private Sub RequisitarParametros(requisicaoParametros As IRequisicaoParametro, cappta As ClienteCappta)

    Dim result As Long
    Dim parametro As Long
    Dim entrada As String
    
    entrada = InputBox(requisicaoParametros.mensagem)
    
    If Len(entrada) <= 0 Then
        parametro = 2
    Else
        parametro = 1
    End If
        
    result = cappta.EnviarParametro(entrada, parametro)

End Sub

```
<h1> Resolver Transacao Pendente</h1>

```javascript
Private Sub ResolverTransacaoPendente(resposta As RespostaTransacaoPendente, cappta As ClienteCappta)
    
    Dim result As Long
    Dim acao As Long
    Dim inputString As String
    Dim mensagemTransacaoPendente As String
    Dim pendencia As TransacaoPendente
    
    For Each Item In resposta.ListaTransacoesPendentes
        
        mensagemTransacaoPendente = mensagemTransacaoPendente & " Número de Controle: " & Item.numeroControle & vbNewLine
        mensagemTransacaoPendente = mensagemTransacaoPendente & " Bandeira: " & Item.NomeBandeiraCartao & vbNewLine
        mensagemTransacaoPendente = mensagemTransacaoPendente & " Adquirente: " & Item.NomeAdquirente & vbNewLine
        mensagemTransacaoPendente = mensagemTransacaoPendente & " Valor: " & Item.valor & vbNewLine
        mensagemTransacaoPendente = mensagemTransacaoPendente & " Data: " & Item.DataHoraAutorizacao & vbNewLine
        
    Next
    
    inputString = Interaction.InputBox(mensagemTransacaoPendente)
    
    If Len(inputString) <= 0 Then
        acao = 2
    Else
        acao = 1
    End If
    
    result = cappta.EnviarParametro(inputString, acao)
    
End Sub

```
<h1>Exibir Dados Operacao Aprovada </h1>

```javascript
Private Sub ExibirDadosOperacaoAprovada(resposta As RespostaOperacaoAprovada)

    Dim mensagemAprovada As String

    If Len(resposta.CupomCliente) > 0 Then
        mensagemAprovada = mensagemAprovada & resposta.CupomCliente & vbNewLine
        
    End If

    If Len(resposta.CupomLojista) > 0 Then
        mensagemAprovada = mensagemAprovada & resposta.CupomLojista & vbNewLine
    End If

    If Len(resposta.CupomReduzido) > 0 Then
        mensagemAprovada = mensagemAprovada & resposta.CupomReduzido & vbNewLine
    End If
      
    AtualizarResultado (mensagemAprovada)
    
End Sub

```
<h1> Finalizar Pagamento </h1>

```javascript

Private Sub FinalizarPagamento(cappta As ClienteCappta)
    
    If processandoPagamento = False Then
        Exit Sub
    End If

    If sessaoMultiTefEmAndamento Then
        
        quantidadeCartoes = quantidadeCartoes - 1
        
        If quantidadeCartoes > 0 Then
            Exit Sub
        End If
    
    End If

    Dim mensagem As String
    mensagem = "Clique em OK para confirmar a transação e em Cancelar para desfaze-la"

    processandoPagamento = False
    sessaoMultiTefEmAndamento = False

    Dim resultado As VbMsgBoxResult
    resultado = MsgBox(mensagem, vbOKCancel, "Cappta Api Sample")
    
    If resultado = vbOK Then
        cappta.ConfirmarPagamentos
    Else
        cappta.DesfazerPagamentos
    End If

End Sub

```

<h1>Etapa 4</h1>
<h4>Tempo estimado de 01:00 hora</h4>

Parabéns agora falta pouco, lembrando que a qualquer momento você pode entrar em contato com a equipe tecnica.

Tel: (11) 4302-6179.

Por se tratar de um ambiente de testes, pode ser utilizado cartões reais para as transações, não sera cobrado nada em sua fatura. Se precisar pode utilizar os cartões presentes em nosso roteiro de teste. Lembrando que vendas digitadas é permitido apenas para a modalidade crédito.

Vamos para a elaboração dos metodos para pagamento.

O primeiro é pagamento débito, o mais simples.

```javascript
Private Sub ExecutarDebito_Click()
    
    If DeveIniciarMultiCartoes() Then
        IniciarMultiCartoes
    End If
     
    Dim valor As Double
    valor = CDbl(TxtValorPagamentoDebito.Text)
    
     If DeveIniciarMultiCartoes() Then
        IniciarMultiCartoes
    End If
        
    Dim resultado As Long
    resultado = cappta.PagamentoDebito(valor)
    
    If resultado <> 0 Then
        CriarMensagem (MensagensPainel.mensagem(resultado))
        Exit Sub
    End If
    
    processandoPagamento = True
    Call IterarOperacaoTef(cappta)
    
End Sub

```

<h1>Agora pagamento credito:</h1>

Private Sub ExecutarCredito_Click()
```javascript    
    Dim valor As Double
    valor = CDbl(TxtValorPagamentoCredito.Text)
    
    Dim detalhes As New DetalhesCredito
    
    detalhes.TransacaoParcelada = OptionTransacaoParceladaCreditoSim.Value
    detalhes.QuantidadeParcelas = UpDownNumeroParcelasCredito.Value
    detalhes.TipoParcelamento = TipoParcelamentoSelecionado()
    
    Dim resultado As Long
    resultado = cappta.PagamentoCredito(valor, detalhes)
    
    If resultado <> 0 Then
        CriarMensagem (MensagensPainel.mensagem(resultado))
        Exit Sub
    End If
    
    processandoPagamento = True
    Call IterarOperacaoTef(cappta)
    
End Sub

```

<h1> Crediário </h1>

```javascript
private void OnExecutaPagamentoCrediarioClick(object sender, EventArgs e)
{
	double valor = (double)NumericUpDownValorPagamentoCrediario.Value;
	IDetalhesCrediario detalhes = new DetalhesCrediario
	{
		QuantidadeParcelas = (int)NumericUpDownQuantidadeParcelasPagamentoCrediario.Value,
	};

	if (this.DeveIniciarMultiCartoes()) { this.IniciarMultiCartoes(); }

	int resultado = this.cliente.PagamentoCrediario(valor, detalhes);
	if (resultado != 0) { this.CriarMensagemErroPainel(resultado); return; }

	this.processandoPagamento = true;
	this.IterarOperacaoTef();
}
```

<h1>Etapa 5 </h1>

<h4>Tempo estimado de 01:00 hora </h4>

Funções administrativas

Agora que tratamos as formas de pagamento, podemos partir para as funções administrativas.

Clientes com frequência pedem a reimpressão de um comprovante ou um cancelamento, as funções administrativas tem a função de deixar praticas e acessiveis estas funções.

Para reimpressão
Temos as seguintes formas:
*Reimpressão por número de controle *Reimpressão cupom lojista *Reimpressão cupom cliente *Reimpressão de todas as vias

```javascript
Private Sub ExecutarReimpressao_Click()
    
    If OptionUsarMultiTef.Value = True Then
        CriarMensagem ("Não é possível reimprimir um cupom com uma sessão multitef em andamento.")
        Exit Sub
    End If
    
    Dim resultado As Long
    
    If OptionReimprimirUltimoCupomSim.Value = True Then
    
        resultado = cappta.ReimprimirUltimoCupom(Via)
    
    Else
    
        resultado = cappta.ReimprimirCupom(TxtNumeroControleReimpressao.Text, Via)
    
    End If
    
    If resultado <> 0 Then
        CriarMensagem (MensagensPainel.mensagem(resultado))
        Exit Sub
    End If
    
     Call IterarOperacaoTef(cappta)
    
End Sub

```

<h1>Para Cancelamento</h1>

Para cancelar uma transação é preciso do número de controle e da senha administrativa, esta senha é configurável no Pinpad e por padrão é: <strong>cappta</strong>. O número de controle é informado na resposta da operação aprovada.

```javascript
Private Sub ExecutarCancelamento_Click()
    
    If OptionUsarMultiTef.Value = True Then
        CriarMensagem ("Não é possível reimprimir um cupom com uma sessão multitef em andamento.")
        Exit Sub
    End If
    
    Dim senhaAdministrativa As String
    senhaAdministrativa = TxtSenhaAdministrativaCancelamento.Text
    
    If Len(senhaAdministrativa) <= 0 Then
        CriarMensagem ("A senha administrativa não pode ser vazia")
        Exit Sub
    End If
    
    Dim numeroControle As String
    numeroControle = TxtNumeroControleCancelamento.Text
    
    Dim resultado As Long
    resultado = cappta.CancelarPagamento(senhaAdministrativa, numeroControle)
    
     Call IterarOperacaoTef(cappta)
    
End Sub
```

<h1>Etapa 6</h1>

<h4>Tempo estimado de 00:40 minutos</h4>

Agora que ja fizemos 80% da integração precisamos trabalhar no Multicartões.

Multicartões ou MultiTef é uma forma de passar mais de um cartão em uma transação, nossa forma de realizar esta tarefa é diferente, se cancelarmos uma venda no meio de uma transação multtef todas são canceladas.


```javascript
Private Sub IniciarMultiCartoes()

  quantidadeCartoes = UpDownQuantidadePagamentosMultiTef.Value
  sessaoMultiTefEmAndamento = True
  cappta.IniciarMultiCartoes (quantidadeCartoes)
    
End Sub

```

Para o código completo basta clonar o repositório, qualquer dúvida entre em contato com o time de homologação e parceria Cappta. Quando completar a integração basta acessar nossa documentação e seguir os passos do nosso [roteiro](http://docs.desktop.cappta.com.br/docs).
Configurando e usando:

Basta alterar os dados de CNPJ e PDV

Muito obrigado pela atenção, qualquer dúvida:
Telefone: (11) 4302-6179
Skype: homologa.cappta1
e-mail: homologa@cappta.com.br
