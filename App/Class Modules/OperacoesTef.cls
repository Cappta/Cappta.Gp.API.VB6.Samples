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
Public sessaoMultiTefEmAndamento As Boolean
Public quantidadeCartoes As Long


Private Sub IniciarMultiCartoes(objCappta As ClienteCappta)
    
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
    
    If sessaoMultiTefEmAndamento = True Then
        
        quantidadeCartoes = 0
        processandoPagamento = False
        sessaoMultiTefEmAndamento = False
    
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

    If sessaoMultiTefEmAndamento = True Then
        
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

Public Sub IterarOperacaoTef(objCappta As ClienteCappta, processandoPag As Boolean, multiSessao As Boolean, cartoes As Long)
 
 processandoPagamento = processandoPag
 sessaoMultiTefEmAndamento = multiSessao
 quantidadeCartoes = cartoes
 
 If sessaoMultiTefEmAndamento = True Then
    
    Call IniciarMultiCartoes(objCappta)
 
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


