Attribute VB_Name = "CriarClienteCappta"
Private Const ChavePDV As String = "795180024C04479982560F61B3C2C06E"
Private Const CNPJ As String = "34555898000186"
Private Const NumeroPDV As Long = 90



Public Function CriarCliente() As ClienteCappta
        
    Dim capptaObj As New ClienteCappta
    Call AutenticarPDV(capptaObj)
    Set CriarCliente = capptaObj
    
End Function


Public Sub AutenticarPDV(objCappta As ClienteCappta)

    Dim resultadoAutenticacao As Long
    
    resultadoAutenticacao = objCappta.AutenticarPDV(CNPJ, NumeroPDV, ChavePDV)
        
    If resultadoAutenticacao <> 0 Then
        
        MsgBox (Mensagens.mensagem(resultadoAutenticacao))
    
    End If
    
End Sub

