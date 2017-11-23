VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OperacoesTef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const INTERVALO_MILISEGUNDOS As Long = 500
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public TextBoxResultado As TextBox
Public processandoPagamento As Boolean
Public iniciarMultiTef As Boolean
Public quantidadeCartoes As Long


Private Sub IniciarMultiCartoes(objCappta As ClienteCappta)
    
   
    iniciarMultiTef = True
    objCappta.IniciarMultiCartoes (quantidadeCartoes)
    
    
End Sub

Private Sub AtualizarResultado(mensagem As String)
    
    TextBoxResultado.Text = mensagem
    TextBoxResultado.Refresh

End Sub

Private Sub ExibirMensagem(resposta As mensagem)

    AtualizarResultado (resposta.Descricao)

End Sub

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

Private Sub ExibirDadosOperacaoRecusada(ByVal result As RespostaOperacaoRecusada)
    
    If iniciarMultiTef = True Then
        
        quantidadeCartoes = 0
        processandoPagamento = False
        iniciarMultiTef = False
    
    End If
    
    AtualizarResultado (result.Motivo & vbNewLine & " Código do Erro: " & result.CodigoMotivo)
    
End Sub


Private Function OperacaoNaoFinalizada(iteracaoTef As IIteracaoTef) As Boolean

    If iteracaoTef.TipoIteracao <> 1 And iteracaoTef.TipoIteracao <> 2 Then
        OperacaoNaoFinalizada = True
    Else
        OperacaoNaoFinalizada = False
    End If
    
End Function

Private Sub FinalizarPagamento(cappta As ClienteCappta)
    
    If processandoPagamento = False Then
        Exit Sub
    End If

    If iniciarMultiTef = True Then
        
        quantidadeCartoes = quantidadeCartoes - 1
        
        If quantidadeCartoes > 0 Then
            Exit Sub
        End If
    
    End If

    Dim mensagem As String
    mensagem = "Clique em OK para confirmar a transação e em Cancelar para desfaze-la"

    processandoPagamento = False
    iniciarMultiTef = False

    Dim resultado As VbMsgBoxResult
    resultado = MsgBox(mensagem, vbOKCancel, "Cappta Api Sample")
    
    If resultado = vbOK Then
        cappta.ConfirmarPagamentos
    Else
        cappta.DesfazerPagamentos
    End If

End Sub

Public Sub IterarOperacaoTef(objCappta As ClienteCappta, processandoPag As Boolean)
 
    
 
 If OptionUsarMultiTef_Click() Then
  OptionNaoUsarMultiTef_Click
 End If
 
 
 

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

End Sub

Private Sub Form_Load()
    
    Set cappta = CriarClienteCappta.CriarCliente()
    Set tef.TextBoxResultado = TextBoxResultado
    processandoMultiPagamento = False
    
    IniciarControles
    ConfigurarModoIntegracao (OptionExibirInterfaceSim.Value)
    
End Sub

Private Sub CriarMensagem(mensagem As String)
        
    MsgBox (mensagem)
    
End Sub

Private Sub AtualizarResultado(mensagem As String)
    
    TextBoxResultado.Text = mensagem
    TextBoxResultado.Refresh

End Sub

'Metodos Tef *************************************************************************************

Private Sub ExecutarDebito_Click()
    If DeveIniciarMultiCartoes() Then
    Call IniciarMultiCartoes(objCappta)
    End If
    
    
    Dim valor As Double
    valor = CDbl(TxtValorPagamentoDebito.Text)
    
    If DeveIniciarMultiCartoes() Then
    Call IniciarMultiCartoes(objCappta)
    End If
    
    Dim resultado As Long
    resultado = cappta.PagamentoDebito(valor)
    
    If resultado <> 0 Then
        CriarMensagem (MensagensPainel.mensagem(resultado))
        Exit Sub
    End If
    
    Call tef.IterarOperacaoTef(cappta, True)
    MultiTefReset
    
End Sub

Private Sub ExecutarCredito_Click()
    
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
    
    Call tef.IterarOperacaoTef(cappta, True, OptionUsarMultiTef.Value, UpDownQuantidadePagamentosMultiTef.Value, processandoMultiPagamento)
    
End Sub

Private Sub ExecutarCrediario_Click()
        
    Dim valor As Double
    valor = CDbl(TxtValorPagamentoCrediario.Text)
    
    Dim detalhes As New DetalhesCrediario
    detalhes.QuantidadeParcelas = CLng(UpDownNumeroParcelasCrediario.Value)
    
    Dim resultado As Long
    resultado = cappta.PagamentoCrediario(valor, detalhes)
    
    If resultado <> 0 Then
        CriarMensagem (MensagensPainel.mensagem(resultado))
        Exit Sub
    End If
    
    Call tef.IterarOperacaoTef(cappta, True, OptionUsarMultiTef.Value, UpDownQuantidadePagamentosMultiTef.Value, processandoMultiPagamento)
    
End Sub


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
    
    Call tef.IterarOperacaoTef(cappta, False, OptionUsarMultiTef.Value, UpDownQuantidadePagamentosMultiTef.Value, processandoMultiPagamento)
    
End Sub

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
    
    Call tef.IterarOperacaoTef(cappta, False, OptionUsarMultiTef.Value, UpDownQuantidadePagamentosMultiTef.Value, processandoMultiPagamento)
    
End Sub

Private Sub ExecutarTicketCar_Click()
    
    Dim resultado As Long
    Dim valor As Double
    
    valor = TxtValorTicketCar.Text
    
    Dim detalhes As New DetalhesPagamentoTicketCarPessoaFisica
    
    detalhes.NumeroReciboFiscal = TxtNumeroDocTicketCar.Text
    detalhes.NumeroSerialECF = TxtNumeroEcfTicketCar.Text
    
    resultado = cappta.PagamentoTicketCarPessoaFisica(valor, detalhes)
    
    
    If resultado <> 0 Then
        CriarMensagem (MensagensPainel.mensagem(resultado))
        Exit Sub
    End If
    
    Call tef.IterarOperacaoTef(cappta, True, OptionUsarMultiTef.Value, UpDownQuantidadePagamentosMultiTef.Value, processandoMultiPagamento)
    
End Sub


Private Sub ExecutarPinPad_Click()
    
    Dim requisicaoPinPad As New RequisicaoInformacaoPinpad
    requisicaoPinPad.TipoInformacaoPinPad = InformacaoPinPadSelecionada()
       
    Dim InformacaoPinPad As String
    InformacaoPinPad = cappta.SolicitarInformacoesPinpad(requisicaoPinPad)
    AtualizarResultado (InformacaoPinPad)
    
End Sub

Private Sub ConfigurarModoIntegracao(exibirGp As Boolean)
    
    Dim configs As New Configuracoes
    configs.ExibirInterface = exibirGp
    
    Dim result As Long
    result = cappta.Configurar(configs)
    
    If result <> 0 Then
        CriarMensagem (MensagensPainel.mensagem(result))
    End If
    
End Sub

'Metodos Tela (Efeitos, Visibles, Preenchimento de combos *************************************************************************************

Private Sub IniciarControles()

    TipoViaSelecionado
    LabelQuantidadeDePagamentosMultiTef.Caption = "Quantidade de pagamentos: " & UpDownQuantidadePagamentosMultiTef.Value
    PreencherInformacoesPinPad
    PreencherTipoParcelamento
    
End Sub

Private Sub MultiTefReset()
    
    Dim quantidadeCartoes As Long
    quantidadeCartoes = UpDownQuantidadePagamentosMultiTef.Value
    
    If quantidadeCartoes <= 0 Then
        quantidadeCartoes = 2
        UpDownQuantidadePagamentosMultiTef.Min = 2
        processandoMultiPagamento = False
    Else
        UpDownQuantidadePagamentosMultiTef.Min = 0
        quantidadeCartoes = quantidadeCartoes - 1
        processandoMultiPagamento = True
    End If
    
    UpDownQuantidadePagamentosMultiTef.Value = quantidadeCartoes
    
End Sub

Private Sub TipoViaSelecionado()
    
    If OptionViaCliente.Value = True Then
        Via = TipoVia.TIPO_VIA_CLIENTE
    ElseIf OptionViaLoja.Value = True Then
        Via = TipoVia.TIPO_VIA_LOJA
    Else
        Via = TipoVia.TIPO_VIA_TODAS
    End If
    
End Sub


Private Sub PreencherInformacoesPinPad()
    
    ComboBoxTipoEntradaPinPad.AddItem "Solicitar CPF"
    ComboBoxTipoEntradaPinPad.AddItem "Solicitar Telefone"
    ComboBoxTipoEntradaPinPad.AddItem "Solicitar Senha"
    
End Sub

Private Sub PreencherTipoParcelamento()
    
    ComboBoxTransacaoParceladaPagamentoCredito.AddItem "Adminsitrativo"
    ComboBoxTransacaoParceladaPagamentoCredito.AddItem "Loja"
    
End Sub

Private Function InformacaoPinPadSelecionada()
        
    InformacaoPinPadSelecionada = ComboBoxTipoEntradaPinPad.ListIndex + 1
    
End Function

Private Function TipoParcelamentoSelecionado()
    
    TipoParcelamentoSelecionado = ComboBoxTransacaoParceladaPagamentoCredito.ListIndex + 1
    
End Function

Private Sub NaoReimprimirUltimoCupomSelecionado(selecionado As Boolean)
    
    LabelNumeroControleReimpressao.visible = selecionado
    TxtNumeroControleReimpressao.visible = selecionado
    UpDownNumeroControleReimpressao.visible = selecionado
    
End Sub

Private Sub UtilizarMultiTefSelecionado(selecionado As Boolean)
    
    LabelQuantidadeDePagamentosMultiTef.visible = selecionado
    UpDownQuantidadePagamentosMultiTef.visible = selecionado
    
End Sub

Private Sub OptionExibirInterfaceNao_Click()
    
    ConfigurarModoIntegracao (OptionExibirInterfaceSim.Value)
    
End Sub

Private Sub OptionExibirInterfaceSim_Click()
    
    ConfigurarModoIntegracao (OptionExibirInterfaceSim.Value)

End Sub

Private Sub TransacaoParceladaSelecionada(visible As Boolean)
    
    GroupBoxDadosParcelamentoCredito.visible = visible
    
End Sub


Private Sub OptionReimprimirUltimoCupomNao_Click()
    
    NaoReimprimirUltimoCupomSelecionado (OptionReimprimirUltimoCupomNao.Value)
    
End Sub

Private Sub OptionReimprimirUltimoCupomSim_Click()

    NaoReimprimirUltimoCupomSelecionado (OptionReimprimirUltimoCupomNao.Value)
    
End Sub

Private Sub OptionTransacaoParceladaCreditoNao_Click()

    TransacaoParceladaSelecionada (OptionTransacaoParceladaCreditoSim.Value)

End Sub

Private Sub OptionTransacaoParceladaCreditoSim_Click()

    TransacaoParceladaSelecionada (OptionTransacaoParceladaCreditoSim.Value)
    
End Sub

Private Sub OptionViaCliente_Click()
    
    TipoViaSelecionado

End Sub

Private Sub OptionViaLoja_Click()
    
    TipoViaSelecionado

End Sub

Private Sub OptionViaTodas_Click()
    
    TipoViaSelecionado

End Sub

Private Sub OptionNaoUsarMultiTef_Click()
    
    UtilizarMultiTefSelecionado (OptionUsarMultiTef.Value)
    DesabilitarBotoes
    DesabilitarControle
    
End Sub

Private Sub OptionUsarMultiTef_Click()
    
    UtilizarMultiTefSelecionado (OptionUsarMultiTef.Value)

End Sub
Private Sub DesabilitarBotoes()

    Call DesabilitarControle(ExecutarDebito_Click)
    Call DesabilitarControle(ExecutarCrediario_Click)
    Call DesabilitarControle(ExecutarCancelamento_Click)
    Call DesabilitarControle(ExecutarReimpressao_Click)
    Call DesabilitarControle(ExecutarCredito_Click)
    
End Sub

Private Sub DesabilitarControlesMultiTef()
OptionNaoUsarMultiTef_Click
End Sub
Private Sub DesabilitarControle(controle As Control)
controle.Enabled = False
End Sub



Private Sub UpDownQuantidadePagamentosMultiTef_Change()
    
    LabelQuantidadeDePagamentosMultiTef.Caption = "Quantidade de pagamentos: " & UpDownQuantidadePagamentosMultiTef.Value
    
End Sub

Private Sub UpDownNumeroParcelasCrediario_Change()
    
    TxtNumeroParcelasCrediario.Text = UpDownNumeroParcelasCrediario.Value

End Sub

Private Sub UpDownNumeroParcelasCredito_Change()
    
    TxtNumeroParcelasPagamentoCredito.Text = UpDownNumeroParcelasCredito.Value

End Sub




