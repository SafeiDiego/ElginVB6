VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmPrincipal 
   Caption         =   "Testa Dll Elgin"
   ClientHeight    =   9690
   ClientLeft      =   4380
   ClientTop       =   2325
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   9975
   Begin RichTextLib.RichTextBox rtbLog 
      Height          =   5175
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9128
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmPrincipal.frx":0000
   End
   Begin VB.CommandButton btnLimparLog 
      Caption         =   "Limpar Log"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton btnPesquisar 
      Caption         =   "Pesquisar"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox edtPesquisa 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   5775
   End
   Begin VB.CommandButton btnExecutar 
      Caption         =   "Executa Comando"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.ListBox lstComandos 
      BackColor       =   &H80000004&
      Height          =   2400
      ItemData        =   "frmPrincipal.frx":0082
      Left            =   360
      List            =   "frmPrincipal.frx":03E3
      TabIndex        =   1
      Top             =   600
      Width           =   7455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comandos:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   795
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExecutar_Click()
    Dim iNumComando As Integer
    Dim iCodErro As Integer
    Dim iRetorno01 As Integer
    Dim iRetorno02 As Integer
    Dim iRetorno03 As Integer
    Dim iRetorno04 As Integer
    Dim iResultado As Integer
    Dim strErrorMsg As String
    Dim strRetorno01 As String
    Dim strRetorno02 As String
    Dim strRetorno03 As String
    Dim strDataTempoInicial As String
    Dim strDataTempoFinal As String
    Dim Duracao As Long
    strErrorMsg = ""
    iCodErro = -999
    iResultado = -999
    iRetorno01 = 0
    iRetorno02 = 0
    iRetorno03 = 0
    iRetorno04 = 0
    strRetorno01 = Space(100)
    strRetorno02 = Space(100)
    strRetorno03 = Space(100)
    
    strDataTempoInicial = Time
    AdicionaTexto ("Início: " & strDataTempoInicial)
    
    Select Case lstComandos.ListIndex
    '=================FUNÇÕES ESPECÍFICAS MFD ELGIN===============
    Case 2
        AdicionaTexto ("Chamando Elgin_CancelaImpressaoCheque() Integer")
        iResultado = Elgin_CancelaImpressaoCheque()
      
    Case 3
        AdicionaTexto ("Chamando Elgin_ImprimeCheque(""0273"",""10,00"",""Elgin"",""Manaus"",""15/09/06"",""Bom para dia 01/10"") Integer")
        iResultado = Elgin_ImprimeCheque("0273", "10,00", "Elgin", "Manaus", "15/09/06", "Bom para dia 01/10")
      
    Case 4
        AdicionaTexto ("Chamando Elgin_ImprimeCopiaCheque( ) Integer")
        iResultado = Elgin_ImprimeCopiaCheque()
      
    Case 5
        AdicionaTexto ("Chamando Elgin_IncluiCidadeFavorecido( ""Manaus"", ""Elgin"" ) Integer")
        iResultado = Elgin_IncluiCidadeFavorecido("Manaus", "Mirian")
      
    Case 6
        AdicionaTexto ("Chamando Elgin_ProgramaMoedaPlural(""Reais"") Integer")
        iResultado = Elgin_ProgramaMoedaPlural("Reais")
      
    Case 7
        AdicionaTexto ("Chamando Elgin_ProgramaMoedaSingular( ""Real"" ) Integer")
        iResultado = Elgin_ProgramaMoedaSingular("Real")
      
    Case 8
        AdicionaTexto ("Chamando Elgin_VerificaStatusCheque(iStatusCheque) Integer")
        iResultado = Elgin_VerificaStatusCheque(iRetorno01)
        AdicionaTexto ("iStatusCheque " + CStr(iRetorno01))
      
    Case 9
        AdicionaTexto ("Chamando Elgin_LeituraChequeMFD(strCodigoCMC7)Integer")
        strRetorno01 = Space(36)
        iResultado = Elgin_LeituraCheque(strRetorno01)
        AdicionaTexto ("Retorno")
        AdicionaTexto ("strCodigoCMC7" + strRetorno01)
      

    '///////////////// FUNÇÕES ESPECÍFICAS ELGIN ///////////////////
    Case 13
        AdicionaTexto ("Chamando Elgin_VendaBruta(strVendaBruta) Integer")
        strRetorno01 = Space(18)
        iResultado = Elgin_VendaBruta(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strVendaBruta " + strRetorno01)
      
    Case 14
        AdicionaTexto ("Chamando Elgin_VendaLiquida(strVendaLiquida) Integer")
        strRetorno01 = Space(18)
        iResultado = Elgin_VendaLiquida(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strVendaLiquida " + strRetorno01)
      
    Case 15
        AdicionaTexto ("Chamando Elgin_TotalDocTroco(strDocTroco) Integer")
        strRetorno01 = Space(18)
        iResultado = Elgin_TotalDocTroco(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strDocTroco " + strRetorno01)
      
    Case 16
        AdicionaTexto ("Chamando Elgin_TotalDiaTroco(strDiaTroco) Integer")
        strRetorno01 = Space(18)
        iResultado = Elgin_TotalDiaTroco(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strDiaTroco " + strRetorno01)
      

    '//////////////// 18 - Funções de Inicialização /////////////////////////////////

    Case 20
        AdicionaTexto ("Chamando Elgin_AlteraSimboloMoeda(""R$"") Integer")
        iResultado = Elgin_AlteraSimboloMoeda("R$")
      
    Case 21
        AdicionaTexto ("Chamando Elgin_ProgramaAliquota(""5,01%"",0) Integer")
        iResultado = Elgin_ProgramaAliquota("5,01%", 0)
      
    Case 22
        AdicionaTexto ("Chamando Elgin_ProgramaHorarioVerao() Integer")
        iResultado = Elgin_ProgramaHorarioVerao()
      
    Case 23
        AdicionaTexto ("Chamando Elgin_NomeiaDepartamento(5,""Vendas"") Integer")
        iResultado = Elgin_NomeiaDepartamento(5, "Vendas")
      
    Case 24
        AdicionaTexto ("Chamando Elgin_NomeiaTotalizadorNaoSujeitoIcms(9,""Conta de Luz"") Integer")
        iResultado = Elgin_NomeiaTotalizadorNaoSujeitoIcms(13, "relatorio")
      
    Case 25
        AdicionaTexto ("Chamando Elgin_ProgramaArredondamento() Integer")
        iResultado = Elgin_ProgramaArredondamento()
      
    Case 26
        AdicionaTexto ("Chamando Elgin_ProgramaTruncamento() Integer")
        iResultado = Elgin_ProgramaTruncamento()
      
    Case 27
        AdicionaTexto ("Chamando Elgin_LinhasEntreCupons(8) Integer")
        iResultado = Elgin_LinhasEntreCupons(8)
      
    Case 28
        AdicionaTexto ("Chamando Elgin_EspacoEntreLinhas(10) Integer")
        iResultado = Elgin_EspacoEntreLinhas(10)
      
    '//////////////// 30 - Funções de Cupom Fiscal /////////////////////////////////

    Case 32
          AdicionaTexto ("Chamando Elgin_AbreCupom(""12844858000102"" ) Integer")
          iResultado = Elgin_AbreCupom("12844858000102")
        
    Case 33
          AdicionaTexto ("Chamando Elgin_VendeItem(""123"", ""Coca-Cola"", ""-3"", ""I"", ""10"", 2, ""5,00"", ""%"", ""5,00"") Integer")
          iResultado = Elgin_VendeItem("000001", "Tabua ganga super " + Chr(34) + "G" + Chr(34) + " madeira  (PC)    X    Preco(R$/PC)            Valor", "-4", "I", "1", 3, "109,250", "%", "")
        
    Case 34
          AdicionaTexto ("Chamando Elgin_VendeItemDepartamento(""55"", ""Gasolina"", ""12"", ""2,50"", ""10"", ""0"", ""0"", ""1"", ""l"") Integer")
          iResultado = Elgin_VendeItemDepartamento("55", "Gasolina", "12", "2,50", "10", "0", "0", "1", "l")
        
    Case 35
          AdicionaTexto ("Chamando Elgin_CancelaItemAnterior()  Integer")
          iResultado = Elgin_CancelaItemAnterior()
        
    Case 36
          AdicionaTexto ("Chamando Elgin_CancelaItemGenerico(""2"")  Integer")
          iResultado = Elgin_CancelaItemGenerico("2")
        
    Case 37
          AdicionaTexto ("Chamando Elgin_CancelaCupom()  Integer")
          iResultado = Elgin_CancelaCupom()
        
    Case 38
          AdicionaTexto ("Chamando Elgin_FechaCupomResumido(""01"", ""Agradecemos a Preferencia"")  Integer")
          iResultado = Elgin_FechaCupomResumido("01", "Agradecemos a Preferencia")
        
    Case 39
          AdicionaTexto ("Chamando Elgin_FechaCupom(""Débito"", "" "", "" "", "" "",""150,00"", ""Agradecemos e Preferencia"")  Integer")
          iResultado = Elgin_FechaCupom("Débito", "", "", "", "150,00", "Agradecemos e Preferencia")
        
    Case 40
          AdicionaTexto ("Chamando Elgin_ResetaImpressora()  Integer")
          iResultado = Elgin_ResetaImpressora()
        
    Case 41
          AdicionaTexto ("Chamando Elgin_IniciaFechamentoCupom(""D"",""$"",""00"")  Integer")
          iResultado = Elgin_IniciaFechamentoCupom("D", "$", "00")
        
    Case 42
          AdicionaTexto ("Chamando Elgin_EfetuaFormaPagamento(""Vale"", ""100,00"")  Integer")
          iResultado = Elgin_EfetuaFormaPagamento("Vale", "100,00")
        
    Case 43
          AdicionaTexto ("Chamando Elgin_EfetuaFormaPagamentoDescricaoForma(""-2"",""50,00"",""Dinheiro  "")  Integer")
          iResultado = Elgin_EfetuaFormaPagamentoDescricaoForma("-2", "50,00", "Dinheiro  ")
        
    Case 44
          AdicionaTexto ("Chamando Elgin_TerminaFechamentoCupom(""Elgin Agradece"")  Integer")
          Elgin_TerminaFechamentoCupom ("iResultado = Elgin Agradece")
        
    Case 45
          AdicionaTexto ("Chamando Elgin_EstornoFormasPagamento(""-2"",""1"",""100,00"")  Integer")
          iResultado = Elgin_EstornoFormasPagamento("-2", "1", "100,00")
        

    '///////////////////// 47 - FUNÇÕES DE RELATÓRIO ////////////////////////////

    Case 49
        AdicionaTexto ("Chamando Elgin_LeituraX() Integer")
        iResultado = Elgin_LeituraX()
      
    Case 50
        AdicionaTexto ("Chamando Elgin_ReducaoZ(""17/09/06"",""0855"")Integer")
        iResultado = Elgin_ReducaoZ("17/09/06", "0855")
      
    Case 51
        AdicionaTexto ("Chamando Elgin_RelatorioGerencial(""Relatório Gerencial Elgin.dll"")Integer")
        iResultado = Elgin_RelatorioGerencial("Relatório Gerencial Elgin.dll")
      
    Case 52
        AdicionaTexto ("Chamando Elgin_FechaRelatorioGerencial() Integer")
        iResultado = Elgin_FechaRelatorioGerencial()
      
    Case 53
        AdicionaTexto ("Chamando Elgin_LeituraMemoriaFiscalData(""01/09/06"", ""29/09/06"",""c"")Integer")
        iResultado = Elgin_LeituraMemoriaFiscalData("01/09/06", "29/09/06", "c")
      
    Case 54
        AdicionaTexto ("Chamando Elgin_LeituraMemoriaFiscalReducaoMFD(""0001"", ""0010"",""c"")Integer")
        iResultado = Elgin_LeituraMemoriaFiscalReducao("0001", "0010", "c")
      
    Case 55
        AdicionaTexto ("Chamando Elgin_LeituraMemoriaFiscalSerialDataMFD(""01/09/06"", ""29/09/06"",""c"")Integer")
        iResultado = Elgin_LeituraMemoriaFiscalSerialData("01/09/06", "29/09/06", "c")
      
    Case 56
        AdicionaTexto ("Chamando Elgin_LeituraMemoriaFiscalSerialReducaoMFD(""0001"", ""0010"",""c"")Integer")
        iResultado = Elgin_LeituraMemoriaFiscalSerialReducao("0001", "0010", "c")
      
    Case 57
        AdicionaTexto ("Chamando Elgin_AbreRelatorioGerencial(""0001"", ""0010"",""c"")Integer")
        iResultado = Elgin_AbreRelatorioGerencial
      

    '  ============== 59 - Funções das Operações Não Fiscais ==============

    Case 61
        AdicionaTexto ("Chamando Elgin_RecebimentoNaoFiscal(""01"",""200"",""-2"")Integer")
        iResultado = Elgin_RecebimentoNaoFiscal("01", "200", "-2")
      
    Case 62
        AdicionaTexto ("Chamando Elgin_AbreComprovanteNaoFiscalVinculado(""TEF"",""1,10"",""000922"")Integer")
        iResultado = Elgin_AbreComprovanteNaoFiscalVinculado("TEF", "1,10", "000922") '// Obs. É necessário mudar o número do cupom
       
    Case 63
        AdicionaTexto ("Chamando Elgin_UsaComprovanteNaoFiscalVinculado(""Teste Elgin.Dll"")Integer")
        iResultado = Elgin_UsaComprovanteNaoFiscalVinculado("Teste Elgin.Dll")
        
    Case 64
        AdicionaTexto ("Chamando Elgin_FechaComprovanteNaoFiscalVinculado()Integer")
        iResultado = Elgin_FechaComprovanteNaoFiscalVinculado()
      
    Case 65
        AdicionaTexto ("Chamando Elgin_Sangria(""50,00"")Integer")
        iResultado = Elgin_Sangria("50,00")
      
    Case 66
        AdicionaTexto ("Chamando Elgin_Suprimento(""100,00"", ""-2"")Integer")
        iResultado = Elgin_Suprimento("100,00", "-2")
      
    Case 67
        AdicionaTexto ("Chamando Elgin_CancelaItemNaoFiscalMFD (""1"") Integer")
        iResultado = Elgin_CancelaItemNaoFiscalMFD("1")
      
    Case 68
        AdicionaTexto ("Chamando Elgin_CancelaAcrescimoNaoFiscalMFD(strACK, strST1, strST2)Integer")
        iResultado = Elgin_CancelaAcrescimoNaoFiscalMFD("009", "A")
      
    Case 69
        AdicionaTexto ("Chamando Elgin_NumeroOperacoesNaoFiscais(strNumOperacoesNaoFiscais)Integer")
        strRetorno01 = Space(6)
        iResultado = Elgin_NumeroOperacoesNaoFiscais(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strNumOperacoesNaoFiscais " + strRetorno01)
      

    ' =============== 71 - Funções de Informações da Impressora ===========

    Case 73
        AdicionaTexto ("Chamando Elgin_NumeroSerie(strNumSerie)Integer")
        strRetorno01 = Space(20)
        iResultado = Elgin_NumeroSerie(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strNumSerie " + strRetorno01)
      
    Case 74
        AdicionaTexto ("Chamando Elgin_SubTotal(strSubTotal)Integer")
        strRetorno01 = Space(14)
        iResultado = Elgin_SubTotal(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strSubTotal " + strRetorno01)
      
    Case 75
        AdicionaTexto ("Chamando Elgin_NumeroCupom(strNumCupom)Integer")
        strRetorno01 = Space(6)
        iResultado = Elgin_NumeroCupom(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strNumCupom " + strRetorno01)
      
    Case 76
        AdicionaTexto ("Chamando Elgin_LeituraXSerial()Integer")
        iResultado = Elgin_LeituraXSerial()
      
    Case 77
        AdicionaTexto ("Chamando Elgin_VersaoFirmware(strVersaoFirmware)Integer")
        strRetorno01 = Space(8)
        iResultado = Elgin_VersaoFirmware(strRetorno01) ' Essa função nao esta declarada como Elgin_VersaoFirmware
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strVersaoFirmware " + strRetorno01)
      
    Case 78
        AdicionaTexto ("Chamando Elgin_CGC_IE(strCGC, strIE)Integer")
        strRetorno01 = Space(8)
        strRetorno01 = Space(15)
        iResultado = Elgin_CGC_IE(strRetorno01, strRetorno02)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strCGC " + strRetorno01)
        AdicionaTexto ("strIE " + strRetorno02)
      
    Case 79
        AdicionaTexto ("Chamando Elgin_GrandeTotal(strGrandeTotal)Integer")
        strRetorno01 = Space(18)
        iResultado = Elgin_GrandeTotal(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strGrandeTotal " + strRetorno01)
      
    Case 80
        AdicionaTexto ("Chamando Elgin_Cancelamentos(strValorCancelado)Integer")
        strRetorno01 = Space(14)
        iResultado = Elgin_Cancelamentos(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strValorCancelado " + strRetorno01)
      
    Case 81
        AdicionaTexto ("Chamando Elgin_Descontos(strValorDescontos)Integer")
        strRetorno01 = Space(14)
        iResultado = Elgin_Descontos(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strValorDescontos " + strRetorno01)
      
    Case 82
        AdicionaTexto ("Chamando Elgin_NumeroOperacoesNaoFiscais(strNumeroOperacoeNFs)Integer")
        strRetorno01 = Space(18)
        iResultado = Elgin_NumeroOperacoesNaoFiscais(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strNumeroOperacoeNFs " + strRetorno01)
      
    Case 83
        AdicionaTexto ("Chamando Elgin_NumeroCuponsCancelados(strNumeroCancelamentos)Integer")
        strRetorno01 = Space(18)
        iResultado = Elgin_NumeroCuponsCancelados(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strNumeroCancelamentos " + strRetorno01)
      
    Case 84
        AdicionaTexto ("Chamando Elgin_NumeroIntervencoes(strNumIntervencoes)Integer")
        strRetorno01 = Space(4)
        iResultado = Elgin_NumeroIntervencoes(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strNumIntervencoes " + strRetorno01)
      
    Case 85
        AdicionaTexto ("Chamando Elgin_NumeroReducoes(strNumReducoes)Integer")
        strRetorno01 = Space(4)
        iResultado = Elgin_NumeroReducoes(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strNumReducoes " + strRetorno01)
      
    Case 86
        AdicionaTexto ("Chamando Elgin_NumeroSubstituicoesProprietario(strNumSubstituicaoProrietarios)Integer")
        strRetorno01 = Space(4)
        iResultado = Elgin_NumeroSubstituicoesProprietario(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strNumSubstituicaoProrietarios " + strRetorno01)
      
    Case 87
        AdicionaTexto ("Chamando Elgin_UltimoItemVendido(strNumItem)Integer")
        strRetorno01 = Space(4)
        iResultado = Elgin_UltimoItemVendido(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strNumItem " + strRetorno01)
      
    Case 88
        AdicionaTexto ("Chamando Elgin_ClicheProprietario(strCliche)Integer")
        strRetorno01 = Space(186)
        iResultado = Elgin_ClicheProprietario(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strCliche " + strRetorno01)
      
    Case 89
        AdicionaTexto ("Chamando Elgin_NumeroCaixa(strNumCaixa)Integer")
        strRetorno01 = Space(4)
        iResultado = Elgin_NumeroCaixa(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strNumCaixa " + strRetorno01)
      
    Case 90
        AdicionaTexto ("Chamando Elgin_NumeroLoja(strNumLoja)Integer")
        strRetorno01 = Space(4)
        iResultado = Elgin_NumeroLoja(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strNumLoja " + strRetorno01)
      
    Case 91
        AdicionaTexto ("Chamando Elgin_SimboloMoeda(strGrandeTotal)Integer")
        strRetorno01 = Space(2)
        iResultado = Elgin_SimboloMoeda(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strGrandeTotal " + strRetorno01)
      
    Case 92
        AdicionaTexto ("Chamando Elgin_MinutosLigada(strMinutosLigada)Integer")
        strRetorno01 = Space(4)
        iResultado = Elgin_MinutosLigada(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strMinutosLigada " + strRetorno01)
      
    Case 93
        AdicionaTexto ("Chamando Elgin_MinutosImprimindo(strMinutosImprimindo)Integer")
        strRetorno01 = Space(4)
        iResultado = Elgin_MinutosImprimindo(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strMinutosImprimindo " + strRetorno01)
      
    Case 94
        AdicionaTexto ("Chamando Elgin_VerificaModoOperacao(strModoOperacao)Integer")
        strRetorno01 = Space(1)
        iResultado = Elgin_VerificaModoOperacao(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strModoOperacao " + strRetorno01)
      
    Case 95
        AdicionaTexto ("Chamando Elgin_FlagsFiscais(iFlags)Integer")
        strRetorno01 = Space(4)
        iResultado = Elgin_FlagsFiscais(iRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("iFlags " + CStr(iRetorno01))
      
    Case 96
        AdicionaTexto ("Chamando Elgin_ValorPagoUltimoCupom(strValorCupom)Integer")
        strRetorno01 = Space(14)
        iResultado = Elgin_ValorPagoUltimoCupom(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strValorCupom " + strRetorno01)
      
    Case 97
        AdicionaTexto ("Chamando Elgin_DataHoraImpressora(strMinutosLigada)Integer")
        strRetorno01 = Space(6)
        strRetorno02 = Space(6)
        iResultado = Elgin_DataHoraImpressora(strRetorno01, strRetorno02)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strData " + strRetorno01)
        AdicionaTexto ("strHora " + strRetorno02)
      
    Case 98
        AdicionaTexto ("Chamando Elgin_ContadoresTotalizadoresNaoFiscais(strContadores)Integer")
        strRetorno01 = Space(44)
        iResultado = Elgin_ContadoresTotalizadoresNaoFiscais(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strContadores " + strRetorno01)
      
    Case 99
        AdicionaTexto ("Chamando Elgin_VerificaTotalizadoresNaoFiscais(strTotalizadores)Integer")
        strRetorno01 = Space(179)
        iResultado = Elgin_VerificaTotalizadoresNaoFiscais(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strTotalizadores " + strRetorno01)
      
    Case 100
        AdicionaTexto ("Chamando Elgin_DataHoraReducao(strData,strHora)Integer")
        strRetorno01 = Space(6)
        strRetorno02 = Space(6)
        iResultado = Elgin_DataHoraReducao(strRetorno01, strRetorno02)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strData " + strRetorno01)
        AdicionaTexto ("strHora " + strRetorno02)
      
    Case 101
        AdicionaTexto ("Chamando Elgin_DataMovimento(strData)Integer")
        strRetorno01 = Space(6)
        iResultado = Elgin_DataMovimento(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strData " + strRetorno01)
      
    Case 102
        AdicionaTexto ("Chamando Elgin_VerificaTruncamento(strTruncamento)Integer")
        strRetorno01 = Space(1)
        iResultado = Elgin_VerificaTruncamento(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strTruncamento " + strRetorno01)
      
    Case 103
        AdicionaTexto ("Chamando Elgin_Acrescimos(strValorAcrecimos)Integer")
        strRetorno01 = Space(14)
        iResultado = Elgin_Acrescimos(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strValorAcrecimos " + strRetorno01)
      
    Case 104
        AdicionaTexto ("Chamando Elgin_VerificaAliquotasIss(strAliquotaIss)Integer")
        strRetorno01 = Space(79)
        iResultado = Elgin_VerificaAliquotasIss(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strAliquotaIss " + strRetorno01)
      
    Case 105
        AdicionaTexto ("Chamando Elgin_VerificaFormasPagamento(strFormasPagamento)Integer")
        strRetorno01 = Space(3016)
        iResultado = Elgin_VerificaFormasPagamento(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strFormasPagamento " + strRetorno01)
      
    Case 106
        AdicionaTexto ("Chamando Elgin_VerificaRecebimentoNaoFiscal(strRecebimentos)Integer")
        strRetorno01 = Space(2200)
        iResultado = Elgin_VerificaRecebimentoNaoFiscal(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strRecebimentos " + strRetorno01)
      
    Case 107
        AdicionaTexto ("Chamando Elgin_VerificaDepartamentos(strDepartamentos)Integer")
        strRetorno01 = Space(1019)
        iResultado = Elgin_VerificaDepartamentos(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strDepartamentos " + strRetorno01)
      
    Case 108
        AdicionaTexto ("Chamando Elgin_VerificaTipoImpressora(iTipoImpressora)Integer")
        iResultado = Elgin_VerificaTipoImpressora(iRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("iTipoImpressora " + CStr(iRetorno01))
      
    Case 109
        AdicionaTexto ("Chamando Elgin_VerificaTotalizadoresParciais(strTotalizadores)Integer")
        strRetorno01 = Space(445)
        iResultado = Elgin_VerificaTotalizadoresParciais(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strTotalizadores " + strRetorno01)
      
    Case 110
        AdicionaTexto ("Chamando Elgin_RetornoAliquotas(strAliquota)Integer")
        strRetorno01 = Space(79)
        iResultado = Elgin_RetornoAliquotas(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strAliquota " + strRetorno01)
      
    Case 111
        AdicionaTexto ("Chamando Elgin_VerificaEstadoImpressora(iACK, iST1, iST2)Integer")
        iResultado = Elgin_VerificaEstadoImpressora(iRetorno01, iRetorno02, iRetorno03)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("iACK " + CStr(iRetorno01))
        AdicionaTexto ("iST1 " + CStr(iRetorno02))
        AdicionaTexto ("iST2 " + CStr(iRetorno03))
      
    Case 112
        AdicionaTexto ("Chamando Elgin_DadosUltimaReducao(strDadosReducao)Integer")
        strRetorno01 = Space(631)
        iResultado = Elgin_DadosUltimaReducao(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strDadosReducao " + strRetorno01)
      
    Case 113
        AdicionaTexto ("Chamando Elgin_VerificaIndiceAliquotasIss(strAliquotasIss)Integer")
        strRetorno01 = Space(79)
        iResultado = Elgin_VerificaIndiceAliquotasIss(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strAliquotasIss " + strRetorno01)
      
    Case 114
        AdicionaTexto ("Chamando Elgin_ValorFormaPagamento(""Dinheiro  "", strValor)Integer")
        strRetorno01 = Space(14)
        iResultado = Elgin_ValorFormaPagamento("Dinheiro  ", strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strValor " + strRetorno01)
      
    Case 115
        AdicionaTexto ("Chamando Elgin_ValorTotalizadorNaoFiscal(""Estorno"",strValor)Integer")
        strRetorno01 = Space(14)
        iResultado = Elgin_ValorTotalizadorNaoFiscal("01", strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strValor " + strRetorno01)
      
    Case 116
        AdicionaTexto ("Chamando Elgin_RetornoImpressora( iCodErro, strErrorMsg ) Integer")
        iResultado = Elgin_RetornoImpressora(iCodErro, strErrorMsg)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("iCodErro " + CStr(iCodErro))
        AdicionaTexto ("strErrorMsg " + strErrorMsg)
      
    Case 117
        AdicionaTexto ("Chamando Elgin_CNPJ_IE(strCNPJ, strIE)Integer")
        strRetorno01 = Space(18)
        strRetorno02 = Space(15)
        iResultado = Elgin_CNPJ_IE(strRetorno01, strRetorno02)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strCNPJ " + (strRetorno01))
        AdicionaTexto ("strIE " + (strRetorno02))
     
    Case 118
        AdicionaTexto ("Chamando Elgin_FlagsFiscaisStr(strFlagFiscal)Integer")
        strRetorno01 = Space(3)
        iResultado = Elgin_FlagsFiscaisStr(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strFlagFiscal " + (strRetorno01))
      
    Case 119
        AdicionaTexto ("Chamando Elgin_VerificaEstadoImpressoraStr(strACK, strST1, strST2)Integer")
        strRetorno01 = Space(10)
        strRetorno02 = Space(10)
        strRetorno03 = Space(10)
        iResultado = Elgin_VerificaEstadoImpressoraStr(strRetorno01, strRetorno02, strRetorno03)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strACK " + (strRetorno01))
        AdicionaTexto ("strST1 " + (strRetorno02))
        AdicionaTexto ("strST2 " + (strRetorno03))
    
    Case 120
        AdicionaTexto ("Chamando Elgin_VerificaTipoImpressoraStr(strTipoImpressora)Integer")
        strRetorno01 = Space(128)
        iResultado = Elgin_VerificaTipoImpressoraStr(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strTipoImpressora " + (strRetorno01))
    

    ' ====== 122 - FUNÇOES DE AUTENTICAÇÃO E GAVETA DE DINHEIRO ========

    Case 124
        AdicionaTexto ("Chamando Elgin_Autenticacao()Integer")
        iResultado = Elgin_Autenticacao()
      
    Case 125
        AdicionaTexto ("Chamando Elgin_ProgramaCaracterAutenticacao(""001,002,004,008,016,032,064,128,064, 032,016,008,004,002,129,129,129,129"") Integer")
        iResultado = Elgin_ProgramaCaracterAutenticacao("001,002,004,008,016,032,064,128,064, 032,016,008,004,002,129,129,129,129")
      
    Case 126
        AdicionaTexto ("Chamando Elgin_AcionaGaveta()Integer")
        iResultado = Elgin_AcionaGaveta()
      
    Case 127
        AdicionaTexto ("Chamando Elgin_VerificaEstadoGaveta(strEstadoGaveta)Integer")
        iResultado = Elgin_VerificaEstadoGaveta(iRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strEstadoGaveta " + CStr(iRetorno01))
      

    ' ===================== 129 - OUTRAS FUNÇÕES ========================

    Case 131
        AdicionaTexto ("Chamando Elgin_AbrePortaSerial()Integer")
        iResultado = Elgin_AbrePortaSerial()
      
    Case 132
        AdicionaTexto ("Chamando Elgin_FechaPortaSerial()Integer")
        iResultado = Elgin_FechaPortaSerial()
      
    Case 133
        AdicionaTexto ("Chamando Elgin_MapaResumo()Integer")
        iResultado = Elgin_MapaResumo()
      
    Case 134
        AdicionaTexto ("Chamando Elgin_AberturaDoDia(""100,00"", ""Dinheiro"")Integer")
        iResultado = Elgin_AberturaDoDia("100,00", "Dinheiro")
      
    Case 135
        AdicionaTexto ("Chamando Elgin_FechamentoDoDia()Integer")
        iResultado = Elgin_FechamentoDoDia()
      
    Case 136
        AdicionaTexto ("Chamando Elgin_ImprimeConfiguracoesImpressora()Integer")
        iResultado = Elgin_ImprimeConfiguracoesImpressora()
      
    Case 137
        AdicionaTexto ("Chamando Elgin_ImprimeDepartamentos()Integer")
        iResultado = Elgin_ImprimeDepartamentos()
      
    Case 138
        AdicionaTexto ("Chamando Elgin_RelatorioTipo60Analitico()Integer")
        iResultado = Elgin_RelatorioTipo60Analitico()
      
    Case 139
        AdicionaTexto ("Chamando Elgin_RelatorioTipo60Mestre()Integer")
        iResultado = Elgin_RelatorioTipo60Mestre()
      
    Case 140
        AdicionaTexto ("Chamando Elgin_VerificaImpressoraLigada()Integer")
        iRetorno01 = Elgin_VerificaImpressoraLigada()
        AdicionaTexto ("Retornos" + CStr(iRetorno01))
        If iRetorno01 = 1 Then
          mmLogComandos.Lines.Append ("Impressora Ligada")
        ElseIf iRetorno01 = -4 Then
          mmLogComandos.Lines.Append ("O arquivo de inicialização Elgin.ini não foi encontrado no diretório de sistema do Windows.")
        ElseIf iRetorno01 = -5 Then
          mmLogComandos.Lines.Append ("Erro ao abrir a porta de comunicação.")
        ElseIf iRetorno01 = -6 Then
          AdicionaTexto ("Impressora desligada ou cabo de comunicação desconectado.")
        End If
      
    Case 141
        AdicionaTexto ("Chamando Elgin_DadosSintegra(""20/09/06"", ""20/09/06"")Integer")
        iResultado = Elgin_DadosSintegra("20/09/06", "20/09/06")
      
    Case 142
        AdicionaTexto ("Chamando Elgin_LeArquivoRetorno(strRetorno) Integer")
        strRetorno01 = Space(6)
        strRetorno02 = Space(1024)
        strRetorno03 = Space(6)
        iResultado = Elgin_NumeroCupom(strRetorno01)
        iResultado = Elgin_RetornoImpressora(iRetorno01, strRetorno02)
        iResultado = Elgin_LeArquivoRetorno(strRetorno03)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strRetorno " + strRetorno03)
      

      ' ================ 144 - FUNÇÕES PARA IMPRESSORAS MFD ==============

    Case 146
         AdicionaTexto ("Chamando Elgin_AbreCupomMFD(""02.844.344/0001-02"", ""Fund. Paulo Feitoza"", ""Gov. Danilo de Matos Areosa, s/nº - lote 164 Distrito "") Integer")
         iResultado = Elgin_AbreCupomMFD("02.844.344/0001-02", "Fund. Paulo Feitoza", "Gov. Danilo de Matos Areosa, s/nº - lote 164 Distrito ")
       
    Case 147
         AdicionaTexto ("Chamando Elgin_CancelaCupomMFD(""02.844.344/0001-02"", ""Fund. Paulo Feitoza"", ""Gov. Danilo de Matos Areosa, s/nº - lote 164 Distrito "") Integer")
         iResultado = Elgin_CancelaCupomMFD("02.844.344/0001-02", "Fund. Paulo Feitoza", "Gov. Danilo de Matos Areosa, s/nº - lote 164 Distrito ")
       
    Case 148
         AdicionaTexto ("Chamando Elgin_ProgramaFormaPagamentoMFD(""Dinheiro"", ""0"")  Integer")
         iResultado = Elgin_ProgramaFormaPagamentoMFD("Dinheiro", "0")
       
    Case 149
         AdicionaTexto ("Chamando Elgin_EfetuaFormaPagamentoMFD(""Dinheiro"", ""100,00"", ""0"", ""Compra à vista"") Integer")
         iResultado = Elgin_EfetuaFormaPagamentoMFD("Dinheiro", "100,00", "0", "Compra à vista")
       
    Case 150
         AdicionaTexto ("Chamando Elgin_CupomAdicionalMFD() Integer")
         iResultado = Elgin_CupomAdicionalMFD()
       
    Case 151
         AdicionaTexto ("Chamando Elgin_AcrescimoDescontoItemMFD (""002"", ""D"",""%"", ""1000"") Integer")
         iResultado = Elgin_AcrescimoDescontoItemMFD("002", "D", "%", "1000")
       
    Case 152
         AdicionaTexto ("Chamando Elgin_NomeiaRelatorioGerencialMFD (""1"", ""Resumo de Vendas"") Integer")
         iResultado = Elgin_NomeiaRelatorioGerencialMFD("1", "Resumo de Vendas")
       
    Case 153
         AdicionaTexto ("Chamando Elgin_AbreComprovanteNaoFiscalVinculadoMFD(""Convênio"", ""100,00"", ""000100"", ""1.111.111-1"", ""Fulano de Tal"", ""R. Sem Fim, 1000"") Integer")
         iResultado = Elgin_AbreComprovanteNaoFiscalVinculadoMFD("Convenio", "100,00", "000386", "1.111.111-1", "Fulano de Tal", "R. Sem Fim, 1000")
       
    Case 154
         AdicionaTexto ("Chamando Elgin_ReimpressaoNaoFiscalVinculadoMFD()  Integer")
         iResultado = Elgin_ReimpressaoNaoFiscalVinculadoMFD()
       
    Case 155
         AdicionaTexto ("Chamando Elgin_AbreRecebimentoNaoFiscalMFD(""02.844.344/0001-02"", ""Fund. Paulo Feitoza"", ""Gov. Danilo de Matos Areosa, s/nº - lote 164 Distrito "") Integer")
         iResultado = Elgin_AbreRecebimentoNaoFiscalMFD("02.844.344/0001-02", "Fund. Paulo Feitoza", "Gov. Danilo de Matos Areosa, s/nº - lote 164 Distrito ")
       
    Case 156
         AdicionaTexto ("Chamando Elgin_EfetuaRecebimentoNaoFiscalMFD(""RGB"", ""50,00"") Integer")
         iResultado = Elgin_EfetuaRecebimentoNaoFiscalMFD("RGB", "50,00")
       
    Case 157
         AdicionaTexto ("Chamando Elgin_IniciaFechamentoCupomMFD(""A"",""%"", ""1000"", ""0000"") Integer")
         iResultado = Elgin_IniciaFechamentoCupomMFD("A", "%", "1000", "0000")
       
    Case 158
         AdicionaTexto ("Chamando Elgin_IniciaFechamentoRecebimentoNaoFiscalMFD(""X"",""%"", ""0000"", ""0000""  string) Integer")
         iResultado = Elgin_IniciaFechamentoRecebimentoNaoFiscalMFD("X", "%", "0000", "0000")
       
    Case 159
         AdicionaTexto ("Chamando Elgin_FechaRecebimentoNaoFiscalMFD(""Obrigado, volte sempre !!!"") Integer")
         iResultado = Elgin_FechaRecebimentoNaoFiscalMFD("Obrigado, volte sempre !!!")
           
    Case 160
         AdicionaTexto ("Chamando Elgin_CancelaRecebimentoNaoFiscalMFD(""02.844.344/0001-02"", ""Fund. Paulo Feitoza"", ""Gov. Danilo de Matos Areosa, s/nº - lote 164 Distrito "") Integer")
         iResultado = Elgin_CancelaRecebimentoNaoFiscalMFD("02.844.344/0001-02", "Fund. Paulo Feitoza", "Gov. Danilo de Matos Areosa, s/nº - lote 164 Distrito")
       
    Case 161
         AdicionaTexto ("Chamando Elgin_AbreRelatorioGerencialMFD(""01"") Integer")
         iResultado = Elgin_AbreRelatorioGerencialMFD("01")
       
    Case 162
         AdicionaTexto ("Chamando Elgin_UsaRelatorioGerencialMFD(""Entre com o texto aqui !!!"") Integer")
         iResultado = Elgin_UsaRelatorioGerencialMFD("Entre com o texto aqui !!!")
       
    Case 163
         AdicionaTexto ("Chamando Elgin_SegundaViaNaoFiscalVinculadoMFD() Integer")
         iResultado = Elgin_SegundaViaNaoFiscalVinculadoMFD()
       
    Case 164
         AdicionaTexto ("Chamando Elgin_VersaoFirmware(""strVerFirmware"") Integer")
         strRetorno01 = Space(20)
         iResultado = Elgin_VersaoFirmware(strRetorno01)
         AdicionaTexto ("strVerFirmware " + strRetorno01)
       
    Case 165
         AdicionaTexto ("Chamando Elgin_CNPJMFD(""strCNPJMFD"") Integer")
         strRetorno01 = Space(20)
         iResultado = Elgin_CNPJMFD(strRetorno01)
         AdicionaTexto ("strCNPJMFD " + strRetorno01)
       
    Case 166
         AdicionaTexto ("Chamando Elgin_InscricaoEstadualMFD(""strIE"") Integer")
         strRetorno01 = Space(20)
         iResultado = Elgin_InscricaoEstadualMFD(strRetorno01)
         AdicionaTexto ("strIE " + strRetorno01)
       
    Case 167
         AdicionaTexto ("Chamando Elgin_InscricaoMunicipalMFD(""strInscMunicipal"") Integer")
         strRetorno01 = Space(20)
         iResultado = Elgin_InscricaoMunicipalMFD(strRetorno01)
         AdicionaTexto ("strInscMunicipal " + strRetorno01)
       
    Case 168
         AdicionaTexto ("Chamando Elgin_TempoOperacionalMFD(""strTempoOp"") Integer")
         strRetorno01 = Space(4)
         iResultado = Elgin_TempoOperacionalMFD(strRetorno01)
         AdicionaTexto ("strTempoOp " + strRetorno01)
       
    Case 169
         AdicionaTexto ("Chamando Elgin_MinutosEmitindoDocumentosFiscaisMFD(""strTempo"") Integer")
         strRetorno01 = Space(4)
         iResultado = Elgin_MinutosEmitindoDocumentosFiscaisMFD(strRetorno01)
         AdicionaTexto ("strTempo " + strRetorno01)
       
    Case 170
         AdicionaTexto ("Chamando Elgin_ContadoresTotalizadoresNaoFiscaisMFD(""strContadores"") Integer")
         strRetorno01 = Space(599)
         iResultado = Elgin_ContadoresTotalizadoresNaoFiscaisMFD(strRetorno01)
         AdicionaTexto ("strContadores " + strRetorno01)
       
    Case 171
         AdicionaTexto ("Chamando Elgin_VerificaTotalizadoresNaoFiscaisMFD(""strTotalizadores"") Integer")
         strRetorno01 = Space(599)
         iResultado = Elgin_VerificaTotalizadoresNaoFiscaisMFD(strRetorno01)
         AdicionaTexto ("strTotalizadores " + strRetorno01)
       
    Case 172
         AdicionaTexto ("Chamando Elgin_VerificaFormasPagamentoMFD(""strFormasPagamento"") Integer")
         strRetorno01 = Space(919)
         iResultado = Elgin_VerificaFormasPagamentoMFD(strRetorno01)
         AdicionaTexto ("strFormasPagamento " + strRetorno01)
       
    Case 173
         AdicionaTexto ("Chamando Elgin_VerificaRecebimentoNaoFiscalMFD(""strRecNaoFiscal"") Integer")
         strRetorno01 = Space(1077)
         iResultado = Elgin_VerificaRecebimentoNaoFiscalMFD(strRetorno01)
         AdicionaTexto ("strRecNaoFiscal " + strRetorno01)
       
    Case 174
         AdicionaTexto ("Chamando Elgin_VerificaRelatorioGerencialMFD(""strRelatorioGerencialMFD"") Integer")
         strRetorno01 = Space(659)
         iResultado = Elgin_VerificaRelatorioGerencialMFD(strRetorno01)
         AdicionaTexto ("strRelatorioGerencialMFD " + strRetorno01)
       
    Case 175
         AdicionaTexto ("Chamando Elgin_ContadorComprovantesCreditoMFD(""strContComprovCreditoMFD"") Integer")
         strRetorno01 = Space(4)
         iResultado = Elgin_ContadorComprovantesCreditoMFD(strRetorno01)
         AdicionaTexto ("strContComprovCreditoMFD " + strRetorno01)
       
    Case 176
         AdicionaTexto ("Chamando Elgin_ContadorOperacoesNaoFiscaisCanceladasMFD(""strContOpNaoFiscaisCanceladasMFD"") Integer")
         strRetorno01 = Space(4)
         iResultado = Elgin_ContadorOperacoesNaoFiscaisCanceladasMFD(strRetorno01)
         AdicionaTexto ("strContOpNaoFiscaisCanceladasMFD " + strRetorno01)
       
    Case 177
         AdicionaTexto ("Chamando Elgin_ContadorRelatoriosGerenciaisMFD (""strContRelGerencialMFD"") Integer")
         strRetorno01 = Space(6)
         iResultado = Elgin_ContadorRelatoriosGerenciaisMFD(strRetorno01)
         AdicionaTexto ("strContRelGerencialMFD " + strRetorno01)
       
    Case 178
         AdicionaTexto ("Chamando Elgin_ContadorFitaDetalheMFD (""strContFitaDetalheMFD"") Integer")
         strRetorno01 = Space(6)
         iResultado = Elgin_ContadorFitaDetalheMFD(strRetorno01)
         AdicionaTexto ("strContFitaDetalheMFD " + strRetorno01)
       
    Case 179
         AdicionaTexto ("Chamando Elgin_ComprovantesNaoFiscaisNaoEmitidosMFD (""strCompNaoFiscalNaoEmitidoMFD"") Integer")
         strRetorno01 = Space(4)
         iResultado = Elgin_ComprovantesNaoFiscaisNaoEmitidosMFD(strRetorno01)
         AdicionaTexto ("strCompNaoFiscalNaoEmitidoMFD " + strRetorno01)
       
    Case 180
         AdicionaTexto ("Chamando Elgin_NumeroSerieMemoriaMFD (""strNumSerieMemoriaMFD"") Integer")
         strRetorno01 = Space(20)
         iResultado = Elgin_NumeroSerieMemoriaMFD(strRetorno01)
         AdicionaTexto ("strNumSerieMemoriaMFD " + strRetorno01)
       
    Case 181
        AdicionaTexto ("Chamando Elgin_MarcaModeloTipoImpressoraMFD(strMarca, strModelo, strTipo)Integer")
        strRetorno01 = Space(15)
        strRetorno02 = Space(20)
        strRetorno03 = Space(7)
        iResultado = Elgin_MarcaModeloTipoImpressoraMFD(strRetorno01, strRetorno02, strRetorno03)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strMarca " + strRetorno01)
        AdicionaTexto ("strModelo " + strRetorno02)
        AdicionaTexto ("strTipo " + strRetorno03)
      
    Case 182
        AdicionaTexto ("Chamando Elgin_ReducoesRestantesMFD(strReducoesRestantes)Integer")
        strRetorno01 = Space(4)
        iResultado = Elgin_ReducoesRestantesMFD(strRetorno01)
        AdicionaTexto ("Retornos")
        mmLogComandos.Lines.Append ("strReducoesRestantes " + strRetorno01)
      
    Case 183
        AdicionaTexto ("Chamando Elgin_VerificaTotalizadoresParciaisMFD(strTotalizadores)Integer")
        strRetorno01 = Space(889)
        iResultado = Elgin_VerificaTotalizadoresParciaisMFD(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strTotalizadores " + strRetorno01)
      
    Case 184
        AdicionaTexto ("Chamando Elgin_DadosUltimaReducaoMFD(strDadosReducao)Integer")
        strRetorno01 = Space(1278)
        iResultado = Elgin_DadosUltimaReducaoMFD(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strDadosReducao " + strRetorno01)
      
'    Case 185
'        AdicionaTexto ("Chamando Elgin_RegistrosTipo60(""1"")Integer")
'        iResultado = Elgin_HabilitaDesabilitaRetornoEstendidoMFD("1")
      
    Case 185
        AdicionaTexto ("Chamando Elgin_AtivaDesativaVendaUmaLinhaMFD(""1"")Integer")
        iResultado = Elgin_AtivaDesativaVendaUmaLinhaMFD("1")
        
    Case 186
        AdicionaTexto ("Chamando Elgin_CancelaAcrescimoDescontoItemMFD(""D"", ""01"")Integer")
        iResultado = Elgin_CancelaAcrescimoDescontoItemMFD("D", "01")
      
    Case 187
        AdicionaTexto ("Chamando Elgin_TotalLivreMFD(strTamanhoTotalLivreMFD)Integer")
        strRetorno01 = Space(10)
        iResultado = Elgin_TotalLivreMFD(strRetorno01)
        AdicionaTexto ("strTamanhoTotalLivreMFD " + strRetorno01)
      
    Case 188
        AdicionaTexto ("Chamando Elgin_TamanhoTotalMFD(strTamanhoTotalMFD)Integer")
        strRetorno01 = Space(10)
        iResultado = Elgin_TamanhoTotalMFD(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strTamanhoTotalMFD " + strRetorno01)
      
    Case 189
        AdicionaTexto ("Chamando Elgin_RegistrosTipo60(""D"", ""%"",""0"")Integer")
        iResultado = Elgin_AcrescimoDescontoSubtotalRecebimentoMFD("D", "%", "10")
      
    Case 190
        AdicionaTexto ("Chamando Elgin_AcrescimoDescontoSubtotalMFD(""D"", ""%"",""0"")Integer")
        iResultado = Elgin_AcrescimoDescontoSubtotalMFD("D", "%", "10")
      
    Case 191
        AdicionaTexto ("Chamando Elgin_CancelaAcrescimoDescontoSubtotalMFD(""D"")Integer")
        iResultado = Elgin_CancelaAcrescimoDescontoSubtotalMFD("D")
      
    Case 192
        AdicionaTexto ("Chamando Elgin_CancelaAcrescimoDescontoSubtotalRecebimentoMFD(""D"")Integer")
        iResultado = Elgin_CancelaAcrescimoDescontoSubtotalRecebimentoMFD("D")
      
    Case 193
        AdicionaTexto ("Chamando Elgin_PercentualLivreMFD(strPercentualLivre)Integer")
        strRetorno01 = Space(12)
        iResultado = Elgin_PercentualLivreMFD(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strPercentualLivre " + strRetorno01)
      
    Case 194
        AdicionaTexto ("Chamando Elgin_DataHoraUltimoDocumentoMFD(strDataHora)Integer")
        strRetorno01 = Space(12)
        iResultado = Elgin_DataHoraUltimoDocumentoMFD(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strDataHora " + strRetorno01)
      
    Case 195
        AdicionaTexto ("Chamando Elgin_MapaResumoMFD()Integer")
        iResultado = Elgin_MapaResumoMFD()
      
    Case 196
        AdicionaTexto ("Chamando Elgin_RelatorioTipo60AnaliticoMFD()Integer")
        iResultado = Elgin_RelatorioTipo60AnaliticoMFD()
      
    Case 197
        AdicionaTexto ("Chamando Elgin_ValorFormaPagamentoMFD(""Dinheiro"" strValor)Integer")
        strRetorno01 = Space(20)
        iResultado = Elgin_ValorFormaPagamento("Dinheiro", strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strValor " + strRetorno01)
      
    Case 198
        AdicionaTexto ("Chamando Elgin_ValorTotalizadorNaoFiscalMFD(""Sangria"", strRetorno01)Integer")
        strRetorno01 = Space(14)
        iResultado = Elgin_ValorTotalizadorNaoFiscalMFD("Sangria", strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strValor " + strRetorno01)
      
    Case 199
        AdicionaTexto ("Chamando Elgin_VerificaEstadoImpressoraMFD(iACK, iST1, iST2, iST3)Integer")
        iResultado = Elgin_VerificaEstadoImpressoraMFD(iRetorno01, iRetorno02, iRetorno03, iRetorno04)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("iACK " + CStr(iRetorno01))
        AdicionaTexto ("iST1 " + CStr(iRetorno02))
        AdicionaTexto ("iST2 " + CStr(iRetorno03))
        AdicionaTexto ("iST3 " + CStr(iRetorno04))
      
    Case 200
        AdicionaTexto ("Chamando Elgin_RelatorioSintegraMFD(28,""c\Saida.txt"",""09"",""2010"",""ELGIN S/A"",""Av. Danilo Areosa"", ""s/n""," & _
                                      """ "",""Distrito Industrial"", ""Manaus-AM"",""6900-000"",""2123-9784"",""2123-9797"",""Claudio"")Integer")
        iResultado = Elgin_RelatorioSintegraMFD(28, "c:\Saida.txt", "09", "2010", "ELGIN S/A", "Av. Danilo Areosa", "s/n", _
                                      "", "Distrito Industrial", "Manaus-AM", "6900-000", "2123-9784", "2123-9797", "Claudio")
      
    Case 201
        AdicionaTexto ("Chamando Elgin_DownloadMFD(""Saida.txt"",""0"",""0001"",""0010"",""1"")Integer")
        iResultado = Elgin_DownloadMFD("Saida.txt", "0", "0001", "0010", "1")
        
      
    Case 202
        AdicionaTexto ("Chamando Elgin_RegistrosTipo60()Integer")
        iResultado = Elgin_RegistrosTipo60()
      
    Case 203
        AdicionaTexto ("Chamando Elgin_FormatoDadosMFD(""c\Donwload.txt"",""c\Saida.txt"",""0"",""0"",""0001"",""0010"",""1"")Integer")
        iResultado = Elgin_FormatoDadosMFD("c\Donwload.txt", "c\Saida.txt", "0", "0", "0001", "0010", "1")
      
    Case 204
        AdicionaTexto ("Chamando Elgin_ContadorRelatoriosGerenciaisMFD(""strContRelatoriosGerenciaisMFD"") Integer")
        strRetorno01 = Space(4)
        iResultado = Elgin_ContadorRelatoriosGerenciaisMFD(strRetorno01)
        AdicionaTexto ("strContRelatoriosGerenciaisMFD " + strRetorno01)
    
    Case 205
        AdicionaTexto ("Chamando Elgin_DownloadMF(""c\Saida.txt"")Integer")
        iResultado = Elgin_DownloadMF("c\Saida.txt")
      
    Case 206
        AdicionaTexto ("Chamando Elgin_GrandeTotalUltimaReducaoMFD(strGrandeTotal)Integer")
        strRetorno01 = Space(18)
        iResultado = Elgin_GrandeTotalUltimaReducaoMFD(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strGrandeTotal " + strRetorno01)
      
    Case 207
        AdicionaTexto ("Chamando Elgin_InicioFimCOOsMFD(strCOOInicial, strCOOFinal)Integer")
        strRetorno01 = Space(6)
        strRetorno02 = Space(6)
        iResultado = Elgin_InicioFimCOOsMFD(strRetorno01, strRetorno02)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strCOOInicial " + strRetorno01)
        AdicionaTexto ("strCOOFinal " + strRetorno02)
      
    Case 208
        AdicionaTexto ("Chamando Elgin_InicioFimGTsMFD(strGTInicial, strGTFinal)Integer")
        strRetorno01 = Space(18)
        strRetorno02 = Space(18)
        iResultado = Elgin_InicioFimGTsMFD(strRetorno01, strRetorno02)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strGTInicial " + strRetorno01)
        AdicionaTexto ("strGTFinal " + strRetorno02)
      
    Case 209
        AdicionaTexto ("Chamando Elgin_StatusEstendidoMFD(iStatus)Integer")
        iResultado = Elgin_StatusEstendidoMFD(iRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("iStatus " + CStr(iRetorno01))
      
    Case 210
        AdicionaTexto ("Chamando Elgin_SubTotalComprovanteNaoFiscalMFD(strSubTotal) Integer")
        strRetorno01 = Space(14)
        AdicionaTexto ("Retornos")
        iResultado = Elgin_SubTotalComprovanteNaoFiscalMFD(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strSubTotal " + strRetorno01)
    
    Case 211
        AdicionaTexto ("Chamando Elgin_VerificaSensorPoucoPapelMFD(strFlag)Integer")
        strRetorno01 = Space(2)
        iResultado = Elgin_VerificaSensorPoucoPapelMFD(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strFlag " + strRetorno01)
      
    Case 212
        AdicionaTexto ("Chamando Elgin_AcrescimoItemNaoFiscalMFD(""01"",""D"",""%"",""5,00"")Integer")
        iResultado = Elgin_AcrescimoItemNaoFiscalMFD("01", "D", "%", "5,00")
      

    ' ====== 214 - FUNÇÕES PARA A IMPRESSÃO DE CÓDIGO DE BARRAS ========

    Case 216
            AdicionaTexto ("Chamando Elgin_TerminaFechamentoCupomCodigoBarrasMFD(""Elgin Teste"", ""EAN13"", ""000000000013"", 120, 2, 2, 0, 0, 4, 6 )Integer")
            iResultado = Elgin_TerminaFechamentoCupomCodigoBarrasMFD("Elgin Teste", "EAN13", "000000000013", 120, 2, 2, 0, 0, 4, 6)
         
    Case 217
        AdicionaTexto ("Chamando Elgin_CodigoBarrasCODABARMFD(""123-ABC/001"")Integer")
        iResultado = Elgin_CodigoBarrasCODABARMFD("123-ABC/001")
      
    Case 218
        AdicionaTexto ("Chamando Elgin_CodigoBarrasCODE128MFD(""Elgin SA"")Integer")
        iResultado = Elgin_CodigoBarrasCODE128MFD("Elgin SA")
      
    Case 219
        AdicionaTexto ("Chamando Elgin_CodigoBarrasCODE39MFD(""abc-123"")Integer")
        iResultado = Elgin_CodigoBarrasCODE39MFD("abc-123")
      
    Case 220
        AdicionaTexto ("Chamando Elgin_CodigoBarrasCODE93MFD(""123-ABC"")Integer")
        iResultado = Elgin_CodigoBarrasCODE93MFD("123-ABC")
      
    Case 221
        AdicionaTexto ("Chamando Elgin_CodigoBarrasEAN13MFD(""123456789012"")Integer")
        iResultado = Elgin_CodigoBarrasEAN13MFD("123456789012")
      
    Case 222
        AdicionaTexto ("Chamando Elgin_CodigoBarrasEAN8MFD(""1234567"")Integer")
        iResultado = Elgin_CodigoBarrasEAN8MFD("1234567")
      
    Case 223
        AdicionaTexto ("Chamando Elgin_CodigoBarrasISBNMFD(""1-56592-292-X 90000"")Integer")
        iResultado = Elgin_CodigoBarrasISBNMFD("1-56592-292-X 90000")
      
    Case 224
        AdicionaTexto ("Chamando Elgin_CodigoBarrasITFMFD(""0123456789012345"")Integer")
        iResultado = Elgin_CodigoBarrasITFMFD("0123456789012345")
      
    Case 225
        AdicionaTexto ("Chamando Elgin_CodigoBarrasMSIMFD(""123"")Integer")
        iResultado = Elgin_CodigoBarrasMSIMFD("123")
      
    Case 226
        AdicionaTexto ("Chamando Elgin_CodigoBarrasPLESSEYMFD(""123-ABC"")Integer")
        iResultado = Elgin_CodigoBarrasPLESSEYMFD("123-ABC")
      
    Case 227
        AdicionaTexto ("Chamando Elgin_CodigoBarrasUPCAMFD(""12345678901"")Integer")
        iResultado = Elgin_CodigoBarrasUPCAMFD("12345678901")
      
    Case 228
        AdicionaTexto ("Chamando Elgin_CodigoBarrasUPCEMFD(""123456"")Integer")
        iResultado = Elgin_CodigoBarrasUPCEMFD("123456")
      

    '============= 230 - FUNÇÕES ADICIONAIS ============================

    Case 232
        AdicionaTexto ("Chamando Elgin_FlagsFiscaisStr(strFlag)Integer")
        strRetorno01 = Space(3)
        iResultado = Elgin_FlagsFiscaisStr(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strFlag " + (strRetorno01))
      
    Case 233
        AdicionaTexto ("Chamando Elgin_VerificaEstadoGavetaStr(strEstadoGaveta)Integer")
        strRetorno01 = Space(1)
        iResultado = Elgin_VerificaEstadoGavetaStr(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strEstadoGaveta " + (strRetorno01))
      
    Case 234
        AdicionaTexto ("Chamando Elgin_VerificaEstadoImpressoraStr(strACK, strST1, strST2)Integer")
        strRetorno01 = Space(3)
        strRetorno02 = Space(3)
        strRetorno03 = Space(3)
        iResultado = Elgin_VerificaEstadoImpressoraStr(strRetorno01, strRetorno02, strRetorno03)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strACK " + (strRetorno01))
        AdicionaTexto ("strST1 " + (strRetorno02))
        AdicionaTexto ("strST2 " + (strRetorno03))
      
    Case 235
        AdicionaTexto ("Chamando Elgin_VerificaSensorPoucoPapelMFD(strFlag)Integer")
        strRetorno01 = Space(1)
        iResultado = Elgin_VerificaSensorPoucoPapelMFD(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strFlag " + (strRetorno01))
      
    Case 236
        AdicionaTexto ("Chamando Elgin_VerificaTipoImpressoraStr(strTipoImpressora)Integer")
        strRetorno01 = Space(1)
        iResultado = Elgin_VerificaTipoImpressoraStr(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strTipoImpressora " + (strRetorno01))
      
    Case 237
        AdicionaTexto ("Chamando Elgin_CancelaItemNaoFiscalMFD(""001"")Integer")
        iResultado = Elgin_CancelaItemNaoFiscalMFD("001")
      
    Case 238
        AdicionaTexto ("Chamando Elgin_AcrescimoItemNaoFiscalMFD(""001"",""A"",""$"",""1,00"")Integer")
        strRetorno01 = Space(1)
        iResultado = Elgin_AcrescimoItemNaoFiscalMFD("001", "A", "$", "1,00")
      
    Case 239
        AdicionaTexto ("Chamando Elgin_CancelaAcrescimoNaoFiscalMFD(""001"",""A"")Integer")
        strRetorno01 = Space(1)
        iResultado = Elgin_CancelaAcrescimoNaoFiscalMFD("001", "A")
      
    Case 240
        AdicionaTexto ("Chamando Elgin_LeArquivoRetorno(strRetorno)Integer")
        strRetorno01 = Space(1024 * 10)
        iResultado = Elgin_LeArquivoRetorno(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strRetorno " + (strRetorno01))
      

   ' =============== 242 - Funções da Wind ===========

    Case 244
        AdicionaTexto ("Wind_AcionaGaveta()")
        iResultado = Wind_AcionaGaveta()
      
    Case 245
        AdicionaTexto ("Chamando Wind_AcionaGuilhotina(1)")
        iResultado = Wind_AcionaGuilhotina(1)
      
    Case 246
        AdicionaTexto ("Chamando Wind_ConfiguraCodigoBarras(100,1,1,0,0)")
        iResultado = Wind_ConfiguraCodigoBarras(100, 1, 1, 0, 0)
      
    Case 247
        AdicionaTexto ("TESTE ELGIN WIND + Chr(10)")
        iResultado = Wind_EnviaBuffer("TESTE ELGIN WIND" + Chr(10))
      
    Case 248
        AdicionaTexto ("Chamando Wind_EnviaBufferFormatado(""Texto em tipo de letra 1, Italico, sublinhado, expandido, enfatizado"",1,1,1,1,1")
        iResultado = Wind_EnviaBufferFormatado("Texto em tipo de letra 1, Italico, sublinhado, expandido, enfatizado" + Chr(10), 1, 1, 1, 1, 1)
      
    Case 249
        AdicionaTexto ("Chamando Wind_EnviaComando(Chr(2),1")
        iResultado = Wind_EnviaComando(Chr(2), 1)
      
    Case 250
        AdicionaTexto ("Chamando Wind_AjustaLarguraPapel(8000)")
        iResultado = Wind_AjustaLarguraPapel(8000)
      
    Case 251
        AdicionaTexto ("Wind_ImprimeCodigoBarrasCODABAR(""9876543210"")")
        iResultado = Wind_ImprimeCodigoBarrasCODABAR("9876543210")
      
    Case 252
        AdicionaTexto ("Chamando Wind_ImprimeCodigoBarrasCODE128(""{AABC"")")
        iResultado = Wind_ImprimeCodigoBarrasCODE128("{AABC")
      
    Case 253
        AdicionaTexto ("Chamando Wind_ImprimeCodigoBarrasCODE39(""ab*TEST80"")")
        iResultado = Wind_ImprimeCodigoBarrasCODE39("ab*TEST80")
      
    Case 254
        AdicionaTexto ("Chamando Wind_ImprimeCodigoBarrasCODE93(""AbCdEfGh12"")")
        iResultado = Wind_ImprimeCodigoBarrasCODE93("AbCdEfGh12")
      
    Case 255
        AdicionaTexto ("Wind_ImprimeCodigoBarrasEAN13(""789123456789"")")
        iResultado = Wind_ImprimeCodigoBarrasEAN13("789123456789")
      
    Case 256
        AdicionaTexto ("Wind_ImprimeCodigoBarrasEAN8(""1234567"")")
        iResultado = Wind_ImprimeCodigoBarrasEAN8("1234567")
      
    Case 257
        AdicionaTexto ("Chamando Wind_ImprimeCodigoBarrasISBN(""1234-56-789 00000"")")
        iResultado = Wind_ImprimeCodigoBarrasISBN("1234-56-789 00000")
      
    Case 258
        AdicionaTexto ("Chamando Wind_ImprimeCodigoBarrasITF(""0123456789012345"")")
        iResultado = Wind_ImprimeCodigoBarrasITF("0123456789012345")
      
    Case 259
        AdicionaTexto ("Chamando Wind_ImprimeCodigoBarrasMSI(""9876543"")")
        iResultado = Wind_ImprimeCodigoBarrasMSI("9876543")
      
    Case 260
        AdicionaTexto ("Chamando Wind_ImprimeCodigoBarrasPDF417(4, 3, 2, 0, ""ELGIN E VC TUDO HAVER!"")")
        iResultado = Wind_ImprimeCodigoBarrasPDF417(4, 3, 2, 0, "ELGIN E VC TUDO HAVER!")
      
    Case 261
        AdicionaTexto ("Chamando Wind_ImprimeCodigoBarrasPLESSEY(""ABC0123"")")
        iResultado = Wind_ImprimeCodigoBarrasPLESSEY("ABC0123")
      
    Case 262
        AdicionaTexto ("Wind_ImprimeCodigoBarrasUPCA(""12345678901"")")
        iResultado = Wind_ImprimeCodigoBarrasUPCA("12345678901")
      
    Case 263
        AdicionaTexto ("Chamando Wind_ImprimeCodigoBarrasUPCE(""04210000526"")")
        iResultado = Wind_ImprimeCodigoBarrasUPCE("04210000526")
                     
    Case 264
        AdicionaTexto ("Chamando Wind_VerificaEstadoGaveta()")
        iResultado = Wind_VerificaEstadoGaveta()
    
    Case 265
        AdicionaTexto ("Chamando Wind_VerificaPoucoPapel()")
        iResultado = Wind_VerificaPoucoPapel()
    
    Case 266
        AdicionaTexto ("Chamando Wind_VerificaFimPapel()")
        iResultado = Wind_VerificaFimPapel()
        
        ' =============== 266 - Funções Adicionais ===========
        
    Case 270
        AdicionaTexto ("Chamando Elgin_DataMovimentoUltimaReducaoMFD(strDataReducao)")
        strRetorno01 = Space(6)
        iResultado = Elgin_DataMovimentoUltimaReducaoMFD(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strDataReducao: " + (strRetorno01))
    Case 271
        AdicionaTexto ("Chamando Elgin_VerificaZPendente(iZPendente)")
        iResultado = Elgin_VerificaZPendente(iRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("iZPendente: " & CStr(iRetorno01))
    Case 272
        AdicionaTexto ("Chamando Elgin_LeIndicadores(iIndicadores)")
        iResultado = Elgin_LeIndicadores(iRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("iIndicadores: " & CStr(iRetorno01))
    Case 273
        AdicionaTexto ("Chamando Elgin_VerificaAliquotasICMS(strAliquotasICMS)")
        iResultado = Elgin_VerificaAliquotasICMS(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("strAliquotasICMS: " & strRetorno01)
    Case 274
        AdicionaTexto ("Chamando RFD_ConvertedaMFDData('01/04/2008', '01/05/2008')")
        iResultado = RFD_ConvertedaMFDData("01/04/2008", "01/05/2008")
    Case 275
        AdicionaTexto ("Chamando Elgin_ExecutaComando('AvancaPapel','Avanco=100')")
        iResultado = Elgin_ExecutaComando("AvancaPapel", "Avanco=100")
    Case 276
        AdicionaTexto ("Chamando Elgin_ExecutaLeitura('LeInteiro', 'NomeInteiro='EspacamentoDocumentos', strRetorno01)")
        strRetorno01 = Space(256)
        iResultado = Elgin_ExecutaLeitura("LeInteiro", "NomeInteiro=" & Chr(34) & "EspacamentoDocumentos" & Chr(34), strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("Retorno: " & strRetorno01)
    Case 277
        AdicionaTexto ("Chamando Elgin_LeCodigoNacionalIdentificacaoECF(CNI)")
        strRetorno01 = Space(6)
        iResultado = Elgin_LeCodigoNacionalIdentificacaoECF(strRetorno01)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("CNI: " & strRetorno01)
    Case 278
        AdicionaTexto ("Chamando Elgin_LeParametrosPAF(CNPJ,Data,Hora,NumSerie,NumECF,GrandeTotal)")
        Dim strRetorno04 As String
        Dim strRetorno05 As String
        Dim strRetorno06 As String
        strRetorno01 = Space(18)
        strRetorno02 = Space(8)
        strRetorno03 = Space(8)
        strRetorno04 = Space(21)
        strRetorno05 = Space(4)
        strRetorno06 = Space(20)
        iResultado = Elgin_LeParametrosPAF(strRetorno01, strRetorno02, strRetorno03, strRetorno04, strRetorno05, strRetorno06)
        AdicionaTexto ("Retornos")
        AdicionaTexto ("CNPJ: " & Trim(strRetorno01))
        AdicionaTexto ("Data: " & Trim(strRetorno02))
        AdicionaTexto ("Hora: " & strRetorno03)
        AdicionaTexto ("Num. de Serie: " & strRetorno04)
        AdicionaTexto ("Num. do ECF: " & strRetorno05)
        AdicionaTexto ("Grande Total: " & strRetorno06)
    Case 279
        AdicionaTexto ("Chamando Elgin_TotalIcmsCupom(TotalICMS)")
        strRetorno01 = Space(14)
        iResultado = Elgin_TotalIcmsCupom(strRetorno01)
        AdicionaTexto ("Retorno")
        AdicionaTexto ("TotalICMS: " & strRetorno01)
    Case 280
        AdicionaTexto ("Chamando Elgin_LeMemoriasBinario(szNomeArquivo,szSerieECF, bArguardaConcluirLeitura)")
        strRetorno01 = Space(21)
        Elgin_NumeroSerie (strRetorno01)
        strRetorno01 = RTrim(strRetorno01)
        iResultado = Elgin_LeMemoriasBinario("Memorias.tdm", strRetorno01, True)
    Case 281
        AdicionaTexto ("Chamando Elgin_GeraArquivoATO17Binario(szArquivoBinario, szArquivoTexto,szPeriodoIni,szPeriodoFIM,tipoPeriodo, szUsuario,szTipoLeitura)")
        iResultado = Elgin_GeraArquivoATO17Binario("Memorias.tdm", "RFD.txt", "20100301", "20100330", Asc("M"), "01", "TDM")
    Case 282
        AdicionaTexto ("Chamando Elgin_LeStatusGeraBinario(nSituacaoAtual, nCodigoErro, nTamanhoLeitura, nProgressoLeitura, strSituacaoAtual)")
        Dim lTotal As Long
        Dim lProgresso As Long
        iResultado = Elgin_LeStatusGeraBinario(iRetorno01, iRetorno02, lTotal, lProgresso, strRetorno01)
        strRetorno01 = Space(256)
        AdicionaTexto ("Retorno")
        AdicionaTexto ("Situaçao: " & Str(iRetorno01))
        AdicionaTexto ("CodErro: " & Str(iRetorno02))
        AdicionaTexto ("TamanhoLeitura: " & Str(lTotal))
        AdicionaTexto ("ProgressoLeitura: " & Str(lProgresso))
        AdicionaTexto ("strSituacaoAtual: " & strRetorno01)
    Case 283
        AdicionaTexto ("Chamando Elgin_CancelaLeituraBinario()")
        iResultado = Elgin_CancelaLeituraBinario
    Case 284
        AdicionaTexto ("Chamando Elgin_GeraRFDBinario( periodoInicial,  periodoFinal,  tipoPeriodo,  tipoLeitura, nomeArquivo)")
        strRetorno01 = Space(256)
        iResultado = Elgin_GeraRFDBinario("20100301", "20100330", 1, 0, strRetorno01)
        AdicionaTexto ("Retorno")
        AdicionaTexto ("nomeArquivo: " & strRetorno01)
    Case 285
        AdicionaTexto ("Chamando Elgin_ConverteATO17ParaPAFRJ( arquivoATO17)")
        iResultado = Elgin_ConverteATO17ParaPAFRJ("RFD.txt")
    Case 286
        AdicionaTexto ("Chamando Elgin_GeraRFDBinarioRJ( periodoInicial,  periodoFinal,  tipoPeriodo)")
        strRetorno01 = Space(14)
        iResultado = Elgin_GeraRFDBinarioRJ("20100301", "20100330", 1)
 Case Else
        AdicionaTexto ("Posição " & CStr(lstComandos.ListIndex) & " - o comando não pode ser executado.")
 End Select
    
 AdicionaTexto ("Retorno da Função: " & CStr(iResultado))
 strErrorMsg = Space(100)
 Elgin_RetornoImpressora iCodErro, strErrorMsg
 AdicionaTexto ("Retorno da Impressora Código: " & CStr(iCodErro) & "  - Mensagem " & strErrorMsg)
 strDataTempoFinal = Time
 AdicionaTexto ("Fim: " & strDataTempoFinal)
 AdicionaTexto ("")
     
    
End Sub
Private Sub AdicionaTexto(ByVal strTexto As String)
    rtbLog.Text = rtbLog.Text & strTexto & Chr$(13)
End Sub
Private Sub btnLimparLog_Click()
    rtbLog.Text = ""
End Sub

Private Sub btnPesquisar_Click()
    Dim iItemPesquisadoPos As Integer
    Dim bAchado As Boolean
    Dim strPesquisa As String
    Dim strItem As String
    
    strPesquisa = UCase((edtPesquisa.Text))

    If (strPesquisa <> "") Then
        iItemPesquisadoPos = lstComandos.ListIndex + 1
    End If
    If (iItemPesquisadoPos >= lstComandos.ListCount - 1) Then
        iItemPesquisadoPos = 0
    End If

    Do
        strItem = UCase(lstComandos.List(iItemPesquisadoPos))
        bAchado = InStr(strItem, strPesquisa) > 0
        If Not bAchado Then
            iItemPesquisadoPos = iItemPesquisadoPos + 1
        End If
            
    Loop Until bAchado Or (iItemPesquisadoPos > lstComandos.ListCount - 1)

    If (Not bAchado) Then
        iItemPesquisadoPos = 0
        lstComandos.ListIndex = 0
    End If

    If (bAchado) Then
        lstComandos.ListIndex = iItemPesquisadoPos
    End If
    
End Sub


Private Sub edtPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
    btnPesquisar_Click
End Sub

Private Sub lstComandos_DblClick()
    btnExecutar_Click
End Sub
