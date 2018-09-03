VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormularioSample 
   Caption         =   "VB6 Sample"
   ClientHeight    =   7860
   ClientLeft      =   6075
   ClientTop       =   1410
   ClientWidth     =   14850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox UpDownQuantidadePagamentosMultiTef 
      Height          =   285
      Left            =   8640
      TabIndex        =   67
      Text            =   "2"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   9480
      TabIndex        =   7
      Top             =   240
      Width           =   5175
      Begin VB.OptionButton OptionExibirInterfaceNao 
         Caption         =   "Invis�vel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptionExibirInterfaceSim 
         Caption         =   "Vis�vel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Modo de Integra��o: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.OptionButton OptionNaoUsarMultiTef 
      Caption         =   "N�o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   480
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton OptionUsarMultiTef 
      Caption         =   "Sim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame GroupBoxResultadoPagamentoDebito 
      Caption         =   "Resultado"
      Height          =   6615
      Left            =   9480
      TabIndex        =   1
      Top             =   960
      Width           =   5175
      Begin VB.TextBox TextBoxResultado 
         Height          =   6255
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   4935
      End
   End
   Begin TabDlg.SSTab TabTicketCar 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "D�bito"
      TabPicture(0)   =   "FormularioSample.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "GroupBoxDadosPagamentoDebito"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ExecutarDebito"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Cr�dito"
      TabPicture(1)   =   "FormularioSample.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GroupBoxDadosPagamentoCredito"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ExecutarCredito"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Credi�rio"
      TabPicture(2)   =   "FormularioSample.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GroupBoxPagamentoCrediario"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "ExecutarCrediario"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Reimpress�o"
      TabPicture(3)   =   "FormularioSample.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "UpDown3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "GroupBoxReimpressao"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "ExecutarReimpressao"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Cancelamento"
      TabPicture(4)   =   "FormularioSample.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GroupBoxDadosCancelamento"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "ExecutarCancelamento"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "TicketCar"
      TabPicture(5)   =   "FormularioSample.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "GroupBoxTicketCar"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "ExecutarTicketCar"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "PinPad"
      TabPicture(6)   =   "FormularioSample.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Frame6"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "ExecutarPinPad"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      Begin VB.CommandButton ExecutarPinPad 
         Caption         =   "Executar Opera��o"
         Height          =   495
         Left            =   1200
         TabIndex        =   55
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Frame Frame6 
         Caption         =   "Solicitar Informa��es no Papel"
         Height          =   3255
         Left            =   240
         TabIndex        =   53
         Top             =   840
         Width           =   4335
         Begin VB.ComboBox ComboBoxTipoEntradaPinPad 
            Height          =   315
            Left            =   480
            TabIndex        =   62
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label Label16 
            Caption         =   "Tipo de Entrada pinpad:"
            Height          =   255
            Left            =   495
            TabIndex        =   54
            Top             =   645
            Width           =   2280
         End
      End
      Begin VB.CommandButton ExecutarTicketCar 
         Caption         =   "Executar Opera��o"
         Height          =   495
         Left            =   -72600
         TabIndex        =   52
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Frame GroupBoxTicketCar 
         Caption         =   "Dados de Pagamento Ticket Car"
         Height          =   4455
         Left            =   -74640
         TabIndex        =   45
         Top             =   720
         Width           =   7215
         Begin VB.TextBox TxtNumeroDocTicketCar 
            Height          =   375
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   2000
            Width           =   1935
         End
         Begin VB.TextBox TxtNumeroEcfTicketCar 
            Height          =   375
            Left            =   500
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   2000
            Width           =   1935
         End
         Begin VB.TextBox TxtValorTicketCar 
            Height          =   285
            Left            =   500
            TabIndex        =   46
            Text            =   "0,10"
            Top             =   1000
            Width           =   1935
         End
         Begin VB.Label LabelDocFiscalTicketCar 
            Caption         =   "N�mero Doc. Fiscal"
            Height          =   255
            Left            =   3840
            TabIndex        =   50
            Top             =   1650
            Width           =   2055
         End
         Begin VB.Label LabelNumeroECFTicketCar 
            Caption         =   "N�mero de S�rie do ECF:"
            Height          =   255
            Left            =   500
            TabIndex        =   49
            Top             =   1650
            Width           =   2055
         End
         Begin VB.Label LabelValorTicketCar 
            Caption         =   "Valor: "
            Height          =   255
            Left            =   500
            TabIndex        =   48
            Top             =   650
            Width           =   975
         End
      End
      Begin VB.CommandButton ExecutarCancelamento 
         Caption         =   "Executar Opera��o"
         Height          =   495
         Left            =   -74160
         TabIndex        =   44
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Frame GroupBoxDadosCancelamento 
         Caption         =   "Dados do Cancelamento"
         Height          =   3495
         Left            =   -74760
         TabIndex        =   39
         Top             =   720
         Width           =   3615
         Begin VB.TextBox TxtNumeroControleCancelamento 
            Height          =   375
            Left            =   500
            TabIndex        =   41
            Text            =   "1"
            Top             =   2000
            Width           =   1935
         End
         Begin VB.TextBox TxtSenhaAdministrativaCancelamento 
            Height          =   285
            Left            =   500
            TabIndex        =   40
            Text            =   "cappta"
            Top             =   1000
            Width           =   1935
         End
         Begin VB.Label LabelNumeroControleCancelamento 
            Caption         =   "N�mero de Controle"
            Height          =   255
            Left            =   500
            TabIndex        =   43
            Top             =   1650
            Width           =   2055
         End
         Begin VB.Label LabelSenhaAdministrativaCancelamento 
            Caption         =   "Senha Administrativa"
            Height          =   255
            Left            =   500
            TabIndex        =   42
            Top             =   650
            Width           =   2055
         End
      End
      Begin VB.CommandButton ExecutarReimpressao 
         Caption         =   "Executar Opera��o"
         Height          =   495
         Left            =   -72960
         TabIndex        =   38
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Frame GroupBoxReimpressao 
         Caption         =   "Dados da Reimpress�o"
         Height          =   5295
         Left            =   -74760
         TabIndex        =   28
         Top             =   840
         Width           =   6375
         Begin VB.OptionButton OptionReimprimirUltimoCupomNao 
            Caption         =   "N�o"
            Height          =   255
            Left            =   1680
            TabIndex        =   33
            Top             =   1080
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptionReimprimirUltimoCupomSim 
            Caption         =   "Sim"
            Height          =   255
            Left            =   500
            TabIndex        =   32
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox TxtNumeroControleReimpressao 
            Height          =   375
            Left            =   4000
            TabIndex        =   29
            Text            =   "2"
            Top             =   1950
            Width           =   1815
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            TabIndex        =   63
            Top             =   1680
            Width           =   3495
            Begin VB.OptionButton OptionViaCliente 
               Caption         =   "Cliente"
               Height          =   255
               Left            =   1320
               TabIndex        =   66
               Top             =   250
               Width           =   975
            End
            Begin VB.OptionButton OptionViaTodas 
               Caption         =   "Todas"
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   250
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton OptionViaLoja 
               Caption         =   "Loja"
               Height          =   255
               Left            =   2520
               TabIndex        =   64
               Top             =   250
               Width           =   975
            End
         End
         Begin VB.Label LabelViaReimpressao 
            Caption         =   "Qual via ?"
            Height          =   255
            Left            =   500
            TabIndex        =   37
            Top             =   1650
            Width           =   2055
         End
         Begin VB.Label LabelNumeroControleReimpressao 
            Caption         =   "N�mero do Controle"
            Height          =   255
            Left            =   4000
            TabIndex        =   31
            Top             =   1650
            Width           =   2055
         End
         Begin VB.Label LabelReimprimirUltimoCupom 
            Caption         =   "Reimprimir �ltimo Cupom"
            Height          =   255
            Left            =   500
            TabIndex        =   30
            Top             =   650
            Width           =   2055
         End
      End
      Begin VB.CommandButton ExecutarCrediario 
         Caption         =   "Executar Opera��o"
         Height          =   495
         Left            =   -73560
         TabIndex        =   27
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Frame GroupBoxPagamentoCrediario 
         Caption         =   "Dados do Pagamento Credi�rio"
         Height          =   3615
         Left            =   -74760
         TabIndex        =   22
         Top             =   720
         Width           =   4935
         Begin VB.TextBox TxtValorPagamentoCrediario 
            Height          =   285
            Left            =   500
            TabIndex        =   24
            Text            =   "0,10"
            Top             =   1000
            Width           =   1935
         End
         Begin VB.TextBox TxtNumeroParcelasCrediario 
            Height          =   375
            Left            =   500
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "1"
            Top             =   1900
            Width           =   480
         End
         Begin VB.Label LabelValorCrediario 
            Caption         =   "Valor: "
            Height          =   255
            Left            =   500
            TabIndex        =   26
            Top             =   650
            Width           =   975
         End
         Begin VB.Label LabelParcelasCrediario 
            Caption         =   "Quantidade de Parcelas"
            Height          =   255
            Left            =   500
            TabIndex        =   25
            Top             =   1500
            Width           =   2055
         End
      End
      Begin VB.CommandButton ExecutarCredito 
         Caption         =   "Executar Opera��o"
         Height          =   495
         Left            =   -72720
         TabIndex        =   21
         Top             =   5400
         Width           =   2055
      End
      Begin VB.CommandButton ExecutarDebito 
         Caption         =   "Executar Opera��o"
         Height          =   495
         Left            =   -74280
         TabIndex        =   12
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Frame GroupBoxDadosPagamentoDebito 
         Caption         =   "Dados do Pagamento D�bito "
         Height          =   2895
         Left            =   -74760
         TabIndex        =   11
         Top             =   720
         Width           =   3855
         Begin VB.TextBox TxtValorPagamentoDebito 
            Height          =   285
            Left            =   500
            TabIndex        =   14
            Text            =   "0,10"
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label LabelValorPagamentoDebito 
            Caption         =   "Valor: "
            Height          =   255
            Left            =   500
            TabIndex        =   13
            Top             =   650
            Width           =   975
         End
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   375
         Left            =   -72600
         TabIndex        =   34
         Top             =   2280
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   2
         Max             =   9
         Min             =   2
         Enabled         =   -1  'True
      End
      Begin VB.Frame GroupBoxDadosPagamentoCredito 
         Caption         =   "Dados do Pagamento Cr�dito"
         Height          =   5415
         Left            =   -74760
         TabIndex        =   15
         Top             =   700
         Width           =   6735
         Begin VB.OptionButton OptionTransacaoParceladaCreditoNao 
            Caption         =   "N�o"
            Height          =   255
            Left            =   1400
            TabIndex        =   20
            Top             =   1680
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptionTransacaoParceladaCreditoSim 
            Caption         =   "Sim"
            Height          =   255
            Left            =   480
            TabIndex        =   19
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox TxtValorPagamentoCredito 
            Height          =   285
            Left            =   500
            TabIndex        =   16
            Text            =   "0,10"
            Top             =   840
            Width           =   1935
         End
         Begin VB.Frame GroupBoxDadosParcelamentoCredito 
            Caption         =   "Dados Parcelamento"
            DragMode        =   1  'Automatic
            Height          =   2415
            Left            =   360
            TabIndex        =   57
            Top             =   2040
            Visible         =   0   'False
            Width           =   4815
            Begin VB.TextBox TxtNumeroParcelasPagamentoCredito 
               Height          =   300
               Left            =   360
               Locked          =   -1  'True
               TabIndex        =   61
               Text            =   "2"
               Top             =   1680
               Width           =   615
            End
            Begin VB.ComboBox ComboBoxTransacaoParceladaPagamentoCredito 
               Height          =   315
               ItemData        =   "FormularioSample.frx":00C4
               Left            =   360
               List            =   "FormularioSample.frx":00CE
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   840
               Width           =   3495
            End
            Begin VB.Label Label4 
               Caption         =   "N�mero de Parcelas"
               Height          =   255
               Left            =   360
               TabIndex        =   59
               Top             =   1320
               Width           =   2055
            End
            Begin VB.Label Label1 
               Caption         =   "Transa��o Parcelada ?"
               Height          =   255
               Left            =   350
               TabIndex        =   58
               Top             =   480
               Width           =   2055
            End
         End
         Begin VB.Label LabelTransacaoParceladaPagamentoCredito 
            Caption         =   "Transa��o Parcelada ?"
            Height          =   255
            Left            =   500
            TabIndex        =   18
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label LabelValorPagamentoCredito 
            Caption         =   "Valor: "
            Height          =   255
            Left            =   500
            TabIndex        =   17
            Top             =   480
            Width           =   975
         End
      End
   End
   Begin MSComCtl2.UpDown UpDown4 
      Height          =   375
      Left            =   2400
      TabIndex        =   35
      Top             =   3240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   2
      Max             =   9
      Min             =   2
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown5 
      Height          =   375
      Left            =   4200
      TabIndex        =   36
      Top             =   2040
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   2
      Max             =   9
      Min             =   2
      Enabled         =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Transa��o Parcelada ?"
      Height          =   255
      Left            =   1320
      TabIndex        =   56
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label LabelQuantidadeDePagamentosMultiTef 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantidade de pagamentos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   480
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label LabelUsarMultiTef 
      Caption         =   "Usar MultiTef?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "FormularioSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Via As Long
Dim sessaoMultiTefEmAndamento As Boolean
Dim cappta As New ClienteCappta
Private Const INTERVALO_MILISEGUNDOS As Long = 500
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public processandoPagamento As Boolean
Public quantidadeCartoes As Long
Private Const ChavePDV As String = "795180024C04479982560F61B3C2C06E"
Private Const CNPJ As String = "34555898000186"
Private Const NumeroPDV As Long = 1

Private Sub Form_Load()
   
    AutenticarPDV
    IniciarControles
    ConfigurarModoIntegracao (OptionExibirInterfaceSim.Value)
   
End Sub

Public Sub AutenticarPDV()
    
  Dim resultadoAutenticacao As Long
    
    resultadoAutenticacao = cappta.AutenticarPDV(CNPJ, NumeroPDV, ChavePDV)
    iniciouTef = True
    If resultadoAutenticacao = 0 Then
            
        Exit Sub
    End If
    
    MsgBox (MensagensPainel.Mensagem(resultadoAutenticacao))
   
End Sub

Private Sub CriarMensagem(Mensagem As String)
        
    MsgBox (Mensagem)
    
End Sub



'Metodos Tef *************************************************************************************

Private Sub ExecutarDebito_Click()
    
    If DeveIniciarMultiCartoes() Then
        IniciarMultiCartoes
    End If
     
    Dim valor As Double
    valor = CDbl(TxtValorPagamentoDebito.Text)
      
    Dim resultado As Long
    resultado = cappta.PagamentoDebito(valor)
    
    If resultado <> 0 Then
        CriarMensagem (MensagensPainel.Mensagem(resultado))
        Exit Sub
    End If
    
    processandoPagamento = True
    Call IterarOperacaoTef(cappta)
    
End Sub

Private Sub ExecutarCredito_Click()

    If DeveIniciarMultiCartoes() Then
        IniciarMultiCartoes
    End If
    
    Dim valor As Double
    valor = CDbl(TxtValorPagamentoCredito.Text)
    
    Dim detalhes As New DetalhesCredito
    
    detalhes.TransacaoParcelada = OptionTransacaoParceladaCreditoSim.Value
    detalhes.QuantidadeParcelas = UpDownNumeroParcelasCredito.Value
    detalhes.TipoParcelamento = TipoParcelamentoSelecionado()
    
    Dim resultado As Long
    resultado = cappta.PagamentoCredito(valor, detalhes)
    
    If resultado <> 0 Then
        CriarMensagem (MensagensPainel.Mensagem(resultado))
        Exit Sub
    End If
    
    processandoPagamento = True
    Call IterarOperacaoTef(cappta)
    
End Sub

Private Sub ExecutarCrediario_Click()
        
    Dim valor As Double
    valor = CDbl(TxtValorPagamentoCrediario.Text)
    
    Dim detalhes As New DetalhesCrediario
    detalhes.QuantidadeParcelas = CLng(UpDownNumeroParcelasCrediario.Value)
    
    Dim resultado As Long
    resultado = cappta.PagamentoCrediario(valor, detalhes)
    
    If resultado <> 0 Then
        CriarMensagem (MensagensPainel.Mensagem(resultado))
        Exit Sub
    End If
    
    processandoPagamento = True
    Call IterarOperacaoTef(cappta)
    
End Sub


Private Sub ExecutarReimpressao_Click()
    
    If OptionUsarMultiTef.Value = True Then
        CriarMensagem ("N�o � poss�vel reimprimir um cupom com uma sess�o multitef em andamento.")
        Exit Sub
    End If
    
    Dim resultado As Long
    
    If OptionReimprimirUltimoCupomSim.Value = True Then
    
        resultado = cappta.ReimprimirUltimoCupom(Via)
    
    Else
    
        resultado = cappta.ReimprimirCupom(TxtNumeroControleReimpressao.Text, Via)
    
    End If
    
    If resultado <> 0 Then
        CriarMensagem (MensagensPainel.Mensagem(resultado))
        Exit Sub
    End If
    
     Call IterarOperacaoTef(cappta)
    
End Sub

Private Sub ExecutarCancelamento_Click()
    
    If OptionUsarMultiTef.Value = True Then
        CriarMensagem ("N�o � poss�vel reimprimir um cupom com uma sess�o multitef em andamento.")
        Exit Sub
    End If
    
    Dim senhaAdministrativa As String
    senhaAdministrativa = TxtSenhaAdministrativaCancelamento.Text
    
    If Len(senhaAdministrativa) <= 0 Then
        CriarMensagem ("A senha administrativa n�o pode ser vazia")
        Exit Sub
    End If
    
    Dim numeroControle As String
    numeroControle = TxtNumeroControleCancelamento.Text
    
    Dim resultado As Long
    resultado = cappta.CancelarPagamento(senhaAdministrativa, numeroControle)
    
    Call IterarOperacaoTef(cappta)
    
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
        CriarMensagem (MensagensPainel.Mensagem(resultado))
        Exit Sub
    End If
    
    Call IterarOperacaoTef(cappta)
    
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
        CriarMensagem (MensagensPainel.Mensagem(result))
    End If
    
End Sub

'Metodos Tela (Efeitos, Visibles, Preenchimento de combos *************************************************************************************

Private Function DeveIniciarMultiCartoes() As Boolean
    
        If sessaoMultiTefEmAndamento = False And OptionUsarMultiTef.Value Then
            IniciarMultiCartoes
        
        Else
            Exit Function
        End If
    
End Function

Private Sub IniciarControles()

    TipoViaSelecionado
    LabelQuantidadeDePagamentosMultiTef.Caption = "Quantidade de pagamentos: "
    PreencherInformacoesPinPad
    PreencherTipoParcelamento
    
End Sub

Private Sub DesabilitarControle(controle As Control)
   controle.Enabled = False
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
    
End Sub

Private Sub OptionUsarMultiTef_Click()
    If OptionUsarMultiTef.Value = False Then
        Exit Sub
    End If
    
    UtilizarMultiTefSelecionado (OptionUsarMultiTef.Value)
    
    
    
End Sub

Private Sub UpDownNumeroParcelasCredito_Change()
    
    TxtNumeroParcelasPagamentoCredito.Text = UpDownNumeroParcelasCredito.Value

End Sub


Private Sub IniciarMultiCartoes()

quantidadeCartoes = CInt(UpDownQuantidadePagamentosMultiTef.Text)
sessaoMultiTefEmAndamento = True
cappta.IniciarMultiCartoes (quantidadeCartoes)
    
End Sub



Private Sub AtualizarResultado(Mensagem As String)
    
    TextBoxResultado.Text = Mensagem
    TextBoxResultado.Refresh

End Sub

Private Sub ExibirMensagem(resposta As Mensagem)

    AtualizarResultado (resposta.Descricao)

End Sub

Private Sub RequisitarParametros(requisicaoParametros As IRequisicaoParametro, cappta As ClienteCappta)

    Dim result As Long
    Dim parametro As Long
    Dim entrada As String
    
    entrada = InputBox(requisicaoParametros.Mensagem)
    
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
        
        mensagemTransacaoPendente = mensagemTransacaoPendente & " N�mero de Controle: " & Item.numeroControle & vbNewLine
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
    
    AtualizarResultado (result.Motivo & vbNewLine & " C�digo do Erro: " & result.CodigoMotivo)
    
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

    If sessaoMultiTefEmAndamento Then
        
        quantidadeCartoes = quantidadeCartoes - 1
        
        If quantidadeCartoes > 0 Then
            Exit Sub
        End If
    
    End If

    Dim Mensagem As String
    Mensagem = "Clique em OK para confirmar a transa��o e em Cancelar para desfaze-la"

    processandoPagamento = False
    sessaoMultiTefEmAndamento = False

    Dim resultado As VbMsgBoxResult
    resultado = MsgBox(Mensagem, vbOKCancel, "Cappta Api Sample")
    
    If resultado = vbOK Then
        cappta.ConfirmarPagamentos
    Else
        cappta.DesfazerPagamentos
    End If

End Sub

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

Private Sub HabilitarControle(Control As Control)
    Control.Enabled = True
End Sub

Private Sub DesabilitarBotoes()

   Call DesabilitarControle(ExecutarCancelamento)
   Call DesabilitarControle(ExecutarCrediario)
   Call DesabilitarControle(ExecutarCredito)
   Call DesabilitarControle(ExecutarDebito)
   Call DesabilitarControle(ExecutarReimpressao)
    
End Sub
Private Sub DesabilitarControlesMultiTef()
   Call DesabilitarControle(OptionNaoUsarMultiTef)
   Call DesabilitarControle(OptionUsarMultiTef)
   Call DesabilitarControle(LabelQuantidadeDePagamentosMultiTef)
   Call DesabilitarControle(UpDownQuantidadePagamentosMultiTef)
End Sub


Private Sub HabilitarBotoes()
  Call HabilitarControle(ExecutarCancelamento)
  Call HabilitarControle(ExecutarCrediario)
  Call HabilitarControle(ExecutarCredito)
  Call HabilitarControle(ExecutarDebito)
  Call HabilitarControle(ExecutarReimpressao)
End Sub

Private Sub HabilitarControlesMultiTef()

  Call HabilitarControle(OptionNaoUsarMultiTef)
  Call HabilitarControle(OptionUsarMultiTef)
  Call HabilitarControle(LabelQuantidadeDePagamentosMultiTef)
  Call HabilitarControle(UpDownQuantidadePagamentosMultiTef)
    
End Sub




