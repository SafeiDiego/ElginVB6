Attribute VB_Name = "ModuloDeclaracoes"
'===============================================================================
'********************************************************************************
'
'                      DECLARAÇÃO DAS FUNÇÕES DA Elgin.DLL
'                      Ultima atualização: 17/12/2010
'                      Elgin.dll v.1.0.0.8
'
'********************************************************************************
'===============================================================================}
Declare Function Elgin_AberturaDoDia Lib "Elgin.dll" (ByVal ValorCompra As String, ByVal FormaPagamento As String) As Integer
Declare Function Elgin_AbreComprovanteNaoFiscalVinculado Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NumeroCupom As String) As Integer
Declare Function Elgin_AbreComprovanteNaoFiscalVinculadoMFD Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal Valor As String, ByVal NumeroCupom As String, ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_AbreCupom Lib "Elgin.dll" (ByVal CGC_CPF As String) As Integer
Declare Function Elgin_AbreCupomMFD Lib "Elgin.dll" (ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_AbrePortaSerial Lib "Elgin.dll" () As Integer
Declare Function Elgin_AbreRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_AbreRelatorioGerencial Lib "Elgin.dll" () As Integer
Declare Function Elgin_AbreRelatorioGerencialMFD Lib "Elgin.dll" (ByVal Indice As String) As Integer
Declare Function Elgin_AcionaGaveta Lib "Elgin.dll" () As Integer
Declare Function Elgin_AcionaGuilhotinaMFD Lib "Elgin.dll" (ByVal TipoCorte As Integer) As Integer
Declare Function Elgin_AcrescimoDescontoItemMFD Lib "Elgin.dll" (ByVal Item As String, ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Declare Function Elgin_AcrescimoDescontoSubtotalMFD Lib "Elgin.dll" (ByVal cFlag As String, ByVal cTipo As String, ByVal cValor As String) As Integer
Declare Function Elgin_AcrescimoDescontoSubtotalRecebimentoMFD Lib "Elgin.dll" (ByVal cFlag As String, ByVal cTipo As String, ByVal cValor As String) As Integer
Declare Function Elgin_AcrescimoItemNaoFiscalMFD Lib "Elgin.dll" (ByVal strNroItem As String, ByVal strAcrescDesc As String, ByVal strTipoAcrescDesc As String, ByVal strValor As String) As Integer
Declare Function Elgin_Acrescimos Lib "Elgin.dll" (ByVal ValorAcrescimos As String) As Integer
Declare Function Elgin_AlteraSimboloMoeda Lib "Elgin.dll" (ByVal SimboloMoeda As String) As Integer
Declare Function Elgin_AtivaDesativaVendaUmaLinhaMFD Lib "Elgin.dll" (ByVal iFlag As Integer) As Integer
Declare Function Elgin_AumentaDescricaoItem Lib "Elgin.dll" (ByVal Descricao As String) As Integer
Declare Function Elgin_Autenticacao Lib "Elgin.dll" () As Integer
Declare Function Elgin_AutenticacaoMFD Lib "Elgin.dll" (ByVal Linhas As String, ByVal Texto As String) As Integer
Declare Function Elgin_CancelaAcrescimoDescontoItemMFD Lib "Elgin.dll" (ByVal cFlag As String, ByVal cItem As String) As Integer
Declare Function Elgin_CancelaAcrescimoDescontoSubtotalMFD Lib "Elgin.dll" (ByVal cFlag As String) As Integer
Declare Function Elgin_CancelaAcrescimoDescontoSubtotalRecebimentoMFD Lib "Elgin.dll" (ByVal cFlag As String) As Integer
Declare Function Elgin_CancelaAcrescimoNaoFiscalMFD Lib "Elgin.dll" (ByVal strNumeroItem As String, ByVal strAcrecDesc As String) As Integer
Declare Function Elgin_CancelaCupom Lib "Elgin.dll" () As Integer
Declare Function Elgin_CancelaCupomMFD Lib "Elgin.dll" (ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_CancelaImpressaoCheque Lib "Elgin.dll" () As Integer
Declare Function Elgin_CancelaItemAnterior Lib "Elgin.dll" () As Integer
Declare Function Elgin_CancelaItemGenerico Lib "Elgin.dll" (ByVal NumeroItem As String) As Integer
Declare Function Elgin_CancelaItemNaoFiscalMFD Lib "Elgin.dll" (ByVal strNroItem As String) As Integer
Declare Function Elgin_CancelaLeituraBinario Lib "Elgin.dll" () As Integer
Declare Function Elgin_Cancelamentos Lib "Elgin.dll" (ByVal ValorCancelamentos As String) As Integer
Declare Function Elgin_CancelaRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_CGC_IE Lib "Elgin.dll" (ByVal CGC As String, ByVal IE As String) As Integer
Declare Function Elgin_ClicheProprietario Lib "Elgin.dll" (ByVal Cliche As String) As Integer
Declare Function Elgin_CNPJ_IE Lib "Elgin.dll" (ByVal CNPJ As String, ByVal IE As String) As Integer
Declare Function Elgin_CNPJMFD Lib "Elgin.dll" (ByVal CNPJ As String) As Integer
Declare Function Elgin_CodigoBarrasCODABARMFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_CodigoBarrasCODE128MFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_CodigoBarrasCODE39MFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_CodigoBarrasCODE93MFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_CodigoBarrasEAN13MFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_CodigoBarrasEAN8MFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_CodigoBarrasISBNMFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_CodigoBarrasITFMFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_CodigoBarrasMSIMFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_CodigoBarrasPLESSEYMFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_CodigoBarrasUPCAMFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_CodigoBarrasUPCEMFD Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Elgin_ComprovantesNaoFiscaisNaoEmitidosMFD Lib "Elgin.dll" (ByVal Comprovantes As String) As Integer
Declare Function Elgin_ConfiguraCodigoBarrasMFD Lib "Elgin.dll" (ByVal Altura As Integer, ByVal Largura As Integer, ByVal pos As Integer, ByVal Fonte As Integer, ByVal Margem As Integer) As Integer
Declare Function Elgin_ContadorComprovantesCreditoMFD Lib "Elgin.dll" (ByVal Comprovantes As String) As Integer
Declare Function Elgin_ContadorCupomFiscalMFD Lib "Elgin.dll" (ByVal CuponsEmitidos As String) As Integer
Declare Function Elgin_ContadoresTotalizadoresNaoFiscais Lib "Elgin.dll" (ByVal Contadores As String) As Integer
Declare Function Elgin_ContadoresTotalizadoresNaoFiscaisMFD Lib "Elgin.dll" (ByVal Contadores As String) As Integer
Declare Function Elgin_ContadorFitaDetalheMFD Lib "Elgin.dll" (ByVal ContadorFita As String) As Integer
Declare Function Elgin_ContadorOperacoesNaoFiscaisCanceladasMFD Lib "Elgin.dll" (ByVal OperacoesCanceladas As String) As Integer
Declare Function Elgin_ContadorRelatoriosGerenciaisMFD Lib "Elgin.dll" (ByVal Relatorios As String) As Integer
Declare Function Elgin_ConverteATO17ParaPAFRJ Lib "Elgin.dll" (ByVal arquivoATO17 As String) As Integer
Declare Function Elgin_CupomAdicionalMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_DadosSintegra Lib "Elgin.dll" (ByVal DataInicial As String, ByVal DataFinal As String) As Integer
Declare Function Elgin_DadosUltimaReducao Lib "Elgin.dll" (ByVal DadosReducao As String) As Integer
Declare Function Elgin_DadosUltimaReducaoMFD Lib "Elgin.dll" (ByVal DadosReducao As String) As Integer
Declare Function Elgin_DataHoraImpressora Lib "Elgin.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Declare Function Elgin_DataHoraReducao Lib "Elgin.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Declare Function Elgin_DataHoraSoftwareBasico Lib "Elgin.dll" (ByVal DataSW As String, ByVal HoraSW As String) As Integer
Declare Function Elgin_DataHoraUltimoDocumentoMFD Lib "Elgin.dll" (ByVal cDataHora As String) As Integer
Declare Function Elgin_DataMovimento Lib "Elgin.dll" (ByVal Data As String) As Integer
Declare Function Elgin_DataMovimentoUltimaReducaoMFD Lib "Elgin.dll" (ByVal cDataMovimento As String) As Integer
Declare Function Elgin_Descontos Lib "Elgin.dll" (ByVal ValorDescontos As String) As Integer
Declare Function Elgin_DownloadMF Lib "Elgin.dll" (ByVal Arquivo As String) As Integer
Declare Function Elgin_DownloadMFD Lib "Elgin.dll" (ByVal Arquivo As String, ByVal TipoDownload As String, ByVal ParametroInicial As String, ByVal ParametroFinal As String, ByVal UsuarioECF As String) As Integer
Declare Function Elgin_EfetuaFormaPagamento Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String) As Integer
Declare Function Elgin_EfetuaFormaPagamentoDescricaoForma Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal DescricaoFormaPagto As String) As Integer
Declare Function Elgin_EfetuaFormaPagamentoMFD Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal ValorFormaPagamento As String, ByVal Parcelas As String, ByVal DescricaoFormaPagto As String) As Integer
Declare Function Elgin_EfetuaRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal IndiceTotalizador As String, ByVal ValorRecebimento As String) As Integer
Declare Function Elgin_EspacoEntreLinhas Lib "Elgin.dll" (ByVal Dots As Integer) As Integer
Declare Function Elgin_EstornoFormasPagamento Lib "Elgin.dll" (ByVal FormaOrigem As String, ByVal FormaDestino As String, ByVal Valor As String) As Integer
Declare Function Elgin_EstornoNaoFiscalVinculadoMFD Lib "Elgin.dll" (ByVal CGC As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_ExecutaComando Lib "Elgin.dll" (ByVal Comando As String, ByVal parametros As String) As Integer
Declare Function Elgin_ExecutaLeitura Lib "Elgin.dll" (ByVal Comando As String, ByVal parametros As String, ByVal retorno As String) As Integer
Declare Function Elgin_FechaComprovanteNaoFiscalVinculado Lib "Elgin.dll" () As Integer
Declare Function Elgin_FechaCupom Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String, ByVal ValorPago As String, ByVal Mensagem As String) As Integer
Declare Function Elgin_FechaCupomResumido Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal Mensagem As String) As Integer
Declare Function Elgin_FechamentoDoDia Lib "Elgin.dll" () As Integer
Declare Function Elgin_FechaPortaSerial Lib "Elgin.dll" () As Integer
Declare Function Elgin_FechaRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal Mensagem As String) As Integer
Declare Function Elgin_FechaRelatorioGerencial Lib "Elgin.dll" () As Integer
Declare Function Elgin_FlagsFiscais Lib "Elgin.dll" (ByRef Flag As Integer) As Integer
Declare Function Elgin_FlagsFiscaisStr Lib "Elgin.dll" (ByVal FlagFiscal As String) As Integer
Declare Function Elgin_FormatoDadosMFD Lib "Elgin.dll" (ByVal ArquivoOrigem As String, ByVal ArquivoDestino As String, ByVal TipoFormato As String, ByVal TipoDownload As String, ByVal ParametroInicial As String, ByVal ParametroFinal As String, ByVal UsuarioECF As String) As Integer
Declare Function Elgin_GeraArquivoATO17Binario Lib "Elgin.dll" (ByVal szArquivoBinario As String, ByVal szArquivoTexto As String, ByVal szPeriodoIni As String, ByVal szPeriodoFIM As String, ByVal tipoPeriodo As Byte, ByVal szUsuario As String, ByVal szTipoLeitura As String) As Integer
Declare Function Elgin_GeraRFDBinario Lib "Elgin.dll" (ByVal periodoInicial As String, ByVal periodoFinal As String, ByVal tipoPeriodo As Integer, ByVal tipoLeitura As Integer, ByVal nomeArquivo As String) As Integer
Declare Function Elgin_GeraRFDBinarioRJ Lib "Elgin.dll" (ByVal periodoInicial As String, ByVal periodoFinal As String, ByVal tipoPeriodo As Integer) As Integer
Declare Function Elgin_GrandeTotal Lib "Elgin.dll" (ByVal GrandeTotal As String) As Integer
Declare Function Elgin_GrandeTotalUltimaReducaoMFD Lib "Elgin.dll" (ByVal cGT As String) As Integer
Declare Function Elgin_HabilitaDesabilitaRetornoEstendidoMFD Lib "Elgin.dll" (ByVal FlagRetorno As String) As Integer
Declare Function Elgin_IdentificaConsumidor Lib "Elgin.dll" (ByVal CNPJ_CPF As String, ByVal Nome As String, ByVal Endereco As String) As Integer
Declare Function Elgin_ImprimeCheque Lib "Elgin.dll" (ByVal Banco As String, ByVal Valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal Data As String, ByVal Mensagem As String) As Integer
Declare Function Elgin_ImprimeChequeMFD Lib "Elgin.dll" (ByVal NumeroBanco As String, ByVal Valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal Data As String, ByVal Mensagem As String, ByVal ImpressaoVerso As String, ByVal Linhas As String) As Integer
Declare Function Elgin_ImprimeConfiguracoesImpressora Lib "Elgin.dll" () As Integer
Declare Function Elgin_ImprimeCopiaCheque Lib "Elgin.dll" () As Integer
Declare Function Elgin_ImprimeDepartamentos Lib "Elgin.dll" () As Integer
Declare Function Elgin_IncluiCidadeFavorecido Lib "Elgin.dll" (ByVal Cidade As String, ByVal Favorecido As String) As Integer
Declare Function Elgin_IniciaFechamentoCupom Lib "Elgin.dll" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimoDesconto As String) As Integer
Declare Function Elgin_IniciaFechamentoCupomMFD Lib "Elgin.dll" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimo As String, ByVal ValorDesconto As String) As Integer
Declare Function Elgin_IniciaFechamentoRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal AcrescimoDesconto As String, ByVal TipoAcrescimoDesconto As String, ByVal ValorAcrescimo As String, ByVal ValorDesconto As String) As Integer
Declare Function Elgin_InicioFimCOOsMFD Lib "Elgin.dll" (ByVal cCOOIni As String, ByVal cCOOFim As String) As Integer
Declare Function Elgin_InicioFimGTsMFD Lib "Elgin.dll" (ByVal cGTIni, ByVal cGTFim As String) As Integer
Declare Function Elgin_InscricaoEstadualMFD Lib "Elgin.dll" (ByVal InscricaoEstadual As String) As Integer
Declare Function Elgin_InscricaoMunicipalMFD Lib "Elgin.dll" (ByVal InscricaoMunicipal As String) As Integer
Declare Function Elgin_LeArquivoRetorno Lib "Elgin.dll" (ByVal sCupom As String) As Integer
Declare Function Elgin_LeCodigoNacionalIdentificacaoECF Lib "Elgin.dll" (ByVal CNI As String) As Integer
Declare Function Elgin_LeIndicadores Lib "Elgin.dll" (ByRef indicador As Integer) As Integer
Declare Function Elgin_LeituraCheque Lib "Elgin.dll" (ByVal CodigoCMC7 As String) As Integer
Declare Function Elgin_LeituraMemoriaFiscalData Lib "Elgin.dll" (ByVal DataInicial As String, ByVal DataFinal As String, ByVal FlagLeitura As String) As Integer
Declare Function Elgin_LeituraMemoriaFiscalReducao Lib "Elgin.dll" (ByVal ReducaoInicial As String, ByVal ReducaoFinal As String, ByVal FlagLeitura As String) As Integer
Declare Function Elgin_LeituraMemoriaFiscalSerialData Lib "Elgin.dll" (ByVal DataInicial As String, ByVal DataFinal As String, ByVal FlagLeitura As String) As Integer
Declare Function Elgin_LeituraMemoriaFiscalSerialReducao Lib "Elgin.dll" (ByVal ReducaoInicial As String, ByVal ReducaoFinal As String, ByVal FlagLeitura As String) As Integer
Declare Function Elgin_LeituraX Lib "Elgin.dll" () As Integer
Declare Function Elgin_LeituraXSerial Lib "Elgin.dll" () As Integer
Declare Function Elgin_LeMemoriasBinario Lib "Elgin.dll" (ByVal szNomeArquivo As String, ByVal szSerieECF As String, ByVal bAguardaConcluirLeitura As Boolean) As Integer
Declare Function Elgin_LeNomeRelatorioGerencial Lib "Elgin.dll" (ByVal Codigo As String, ByVal NomeRelatorio As String) As Integer
Declare Function Elgin_LeParametrosPAF Lib "Elgin.dll" (ByVal CNPJ As String, ByVal Data As String, ByVal Hora As String, ByVal NumSerie As String, ByVal NumECF As String, ByVal GrandeTotal As String) As Integer
Declare Function Elgin_LeStatusGeraBinario Lib "Elgin.dll" (ByRef nSituacaoAtual As Integer, ByRef nCodigoErro As Integer, ByRef nTamanhoLeitura As Long, ByRef nProgressoLeitura As Long, ByVal strSituacaoAtual As String) As Integer
Declare Function Elgin_LinhasEntreCupons Lib "Elgin.dll" (ByVal Linhas As Integer) As Integer
Declare Function Elgin_MapaResumo Lib "Elgin.dll" () As Integer
Declare Function Elgin_MapaResumoMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_MarcaModeloTipoImpressoraMFD Lib "Elgin.dll" (ByVal Marca As String, ByVal Modelo As String, ByVal Tipo As String) As Integer
Declare Function Elgin_MinutosEmitindoDocumentosFiscaisMFD Lib "Elgin.dll" (ByVal Minutos As String) As Integer
Declare Function Elgin_MinutosImprimindo Lib "Elgin.dll" (ByVal Minutos As String) As Integer
Declare Function Elgin_MinutosLigada Lib "Elgin.dll" (ByVal Minutos As String) As Integer
Declare Function Elgin_ModeloImpressora Lib "Elgin.dll" (ByVal ModeloImpressora As String) As Integer
Declare Function Elgin_NomeiaDepartamento Lib "Elgin.dll" (ByVal Indice As Integer, ByVal Departamento As String) As Integer
Declare Function Elgin_NomeiaRelatorioGerencialMFD Lib "Elgin.dll" (ByVal Indice As String, ByVal Descricao As String) As Integer
Declare Function Elgin_NomeiaTotalizadorNaoSujeitoIcms Lib "Elgin.dll" (ByVal Indice As Integer, ByVal Totalizador As String) As Integer
Declare Function Elgin_NumeroCaixa Lib "Elgin.dll" (ByVal NumeroCaixa As String) As Integer
Declare Function Elgin_NumeroCupom Lib "Elgin.dll" (ByVal NumeroCupom As String) As Integer
Declare Function Elgin_NumeroCuponsCancelados Lib "Elgin.dll" (ByVal NumeroCancelamentos As String) As Integer
Declare Function Elgin_NumeroIntervencoes Lib "Elgin.dll" (ByVal NumeroIntervencoes As String) As Integer
Declare Function Elgin_NumeroLoja Lib "Elgin.dll" (ByVal NumeroLoja As String) As Integer
Declare Function Elgin_NumeroOperacoesNaoFiscais Lib "Elgin.dll" (ByVal NumeroOperacoes As String) As Integer
Declare Function Elgin_NumeroReducoes Lib "Elgin.dll" (ByVal NumeroReducoes As String) As Integer
Declare Function Elgin_NumeroSerie Lib "Elgin.dll" (ByVal NumeroSerie As String) As Integer
Declare Function Elgin_NumeroSerieCriptografado Lib "Elgin.dll" (ByVal NumeroSerie As String) As Integer
Declare Function Elgin_NumeroSerieDescriptografado Lib "Elgin.dll" (ByVal NumeroSerieCriptografado As String, ByVal NumeroSerieDesCriptografado As String) As Integer
Declare Function Elgin_NumeroSerieMemoriaMFD Lib "Elgin.dll" (ByVal NumeroSerieMFD As String) As Integer
Declare Function Elgin_NumeroSubstituicoesProprietario Lib "Elgin.dll" (ByVal NumeroSubstituicoes As String) As Integer
Declare Function Elgin_PercentualLivreMFD Lib "Elgin.dll" (ByVal cMemoriaLivre As String) As Integer
Declare Function Elgin_ProgramaAliquota Lib "Elgin.dll" (ByVal Aliquota As String, ByVal ICMS_ISS As Integer) As Integer
Declare Function Elgin_ProgramaArredondamento Lib "Elgin.dll" () As Integer
Declare Function Elgin_ProgramaBaudRate Lib "Elgin.dll" (ByVal BaudRate As String) As Integer
Declare Function Elgin_ProgramaCaracterAutenticacao Lib "Elgin.dll" (ByVal parametros As String) As Integer
Declare Function Elgin_ProgramaFormaPagamentoMFD Lib "Elgin.dll" (ByVal FormaPagto As String, ByVal OperacaoTef As String) As Integer
Declare Function Elgin_ProgramaHorarioVerao Lib "Elgin.dll" () As Integer
Declare Function Elgin_ProgramaIdAplicativoMFD Lib "Elgin.dll" (ByVal NomeAplicativo As String) As Integer
Declare Function Elgin_ProgramaMoedaPlural Lib "Elgin.dll" (ByVal MoedaPlural As String) As Integer
Declare Function Elgin_ProgramaMoedaSingular Lib "Elgin.dll" (ByVal MoedaSingular As String) As Integer
Declare Function Elgin_ProgramaOperador Lib "Elgin.dll" (ByVal NomeOperador As String) As Integer
Declare Function Elgin_ProgramaTruncamento Lib "Elgin.dll" () As Integer
Declare Function Elgin_RecebimentoNaoFiscal Lib "Elgin.dll" (ByVal IndiceTotalizador As String, ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Declare Function Elgin_ReducaoZ Lib "Elgin.dll" (ByVal Data As String, ByVal Hora As String) As Integer
Declare Function Elgin_ReducoesRestantesMFD Lib "Elgin.dll" (ByVal Reducoes As String) As Integer
Declare Function Elgin_RegistrosTipo60 Lib "Elgin.dll" () As Integer
Declare Function Elgin_ReimpressaoNaoFiscalVinculadoMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_RelatorioGerencial Lib "Elgin.dll" (ByVal Texto As String) As Integer
Declare Function Elgin_RelatorioSintegraMFD Lib "Elgin.dll" (ByVal iRelatorios As Integer, ByVal cArquivo As String, ByVal cMes As String, ByVal cAno As String, ByVal cRazaoSocial As String, ByVal cEndereco As String, ByVal cNumero As String, ByVal cComplemento As String, ByVal cBairro As String, ByVal cCidade As String, ByVal cCEP As String, ByVal cTelefone As String, ByVal cFax As String, ByVal cContato As String) As Integer
Declare Function Elgin_RelatorioTipo60Analitico Lib "Elgin.dll" () As Integer
Declare Function Elgin_RelatorioTipo60AnaliticoMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_RelatorioTipo60Mestre Lib "Elgin.dll" () As Integer
Declare Function Elgin_ResetaImpressora Lib "Elgin.dll" () As Integer
Declare Function Elgin_RetornoAliquotas Lib "Elgin.dll" (ByVal Aliquotas As String) As Integer
Declare Function Elgin_RetornoImpressora Lib "Elgin.dll" (ByRef i As Integer, ByVal ErrorMsg As String) As Integer
Declare Function Elgin_Sangria Lib "Elgin.dll" (ByVal Valor As String) As Integer
Declare Function Elgin_SegundaViaNaoFiscalVinculadoMFD Lib "Elgin.dll" () As Integer
Declare Function Elgin_SimboloMoeda Lib "Elgin.dll" (ByVal SimboloMoeda As String) As Integer
Declare Function Elgin_StatusEstendidoMFD Lib "Elgin.dll" (ByRef iStatus As Integer) As Integer
Declare Function Elgin_SubTotal Lib "Elgin.dll" (ByVal SubTotal As String) As Integer
Declare Function Elgin_SubTotalComprovanteNaoFiscalMFD Lib "Elgin.dll" (ByVal cSubTotal As String) As Integer
Declare Function Elgin_Suprimento Lib "Elgin.dll" (ByVal Valor As String, ByVal FormaPagamento As String) As Integer
Declare Function Elgin_TamanhoTotalMFD Lib "Elgin.dll" (ByVal cTamanhoMFD As String) As Integer
Declare Function Elgin_TempoOperacionalMFD Lib "Elgin.dll" (ByVal TempoOperacional As String) As Integer
Declare Function Elgin_TerminaFechamentoCupom Lib "Elgin.dll" (ByVal Mensagem As String) As Integer
Declare Function Elgin_TerminaFechamentoCupomCodigoBarrasMFD Lib "Elgin.dll" (ByVal cMensagem As String, ByVal cTipoCodigo As String, ByVal cCodigo As String, ByVal iAltura As Integer, ByVal iLargura As Integer, ByVal iPosicaoCaracteres As Integer, ByVal iFonte As Integer, ByVal iMargem As Integer, ByVal iCorrecaoErros As Integer, ByVal iColunas As Integer) As Integer
Declare Function Elgin_TotalDiaTroco Lib "Elgin.dll" (ByVal TotalDiaTroco As String) As Integer
Declare Function Elgin_TotalDocTroco Lib "Elgin.dll" (ByVal TotalDocTroco As String) As Integer
Declare Function Elgin_TotalIcmsCupom Lib "Elgin.dll" (ByVal TotalICMS As String) As Integer
Declare Function Elgin_TotalLivreMFD Lib "Elgin.dll" (ByVal cMemoriaLivre As String) As Integer
Declare Function Elgin_UltimoItemVendido Lib "Elgin.dll" (ByVal NumeroItem As String) As Integer
Declare Function Elgin_UsaComprovanteNaoFiscalVinculado Lib "Elgin.dll" (ByVal Texto As String) As Integer
Declare Function Elgin_UsaRelatorioGerencialMFD Lib "Elgin.dll" (ByVal Texto As String) As Integer
Declare Function Elgin_ValorFormaPagamento Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal Valor As String) As Integer
Declare Function Elgin_ValorFormaPagamentoMFD Lib "Elgin.dll" (ByVal FormaPagamento As String, ByVal Valor As String) As Integer
Declare Function Elgin_ValorPagoUltimoCupom Lib "Elgin.dll" (ByVal ValorCupom As String) As Integer
Declare Function Elgin_ValorTotalizadorNaoFiscal Lib "Elgin.dll" (ByVal Totalizador As String, ByVal Valor As String) As Integer
Declare Function Elgin_ValorTotalizadorNaoFiscalMFD Lib "Elgin.dll" (ByVal Totalizador As String, ByVal Valor As String) As Integer
Declare Function Elgin_VendaBruta Lib "Elgin.dll" (ByVal VendaBruta As String) As Integer
Declare Function Elgin_VendaLiquida Lib "Elgin.dll" (ByVal VendaLiquida As String) As Integer
Declare Function Elgin_VendeItem Lib "Elgin.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal TipoQuantidade As String, ByVal Quantidade As String, ByVal CasasDecimais As Integer, ByVal ValorUnitario As String, ByVal TipoDesconto As String, ByVal Desconto As String) As Integer
Declare Function Elgin_VendeItemDepartamento Lib "Elgin.dll" (ByVal Codigo As String, ByVal Descricao As String, ByVal Aliquota As String, ByVal ValorUnitario As String, ByVal Quantidade As String, ByVal Acrescimo As String, ByVal Desconto As String, ByVal IndiceDepartamento As String, ByVal UnidadeMedida As String) As Integer
Declare Function Elgin_VerificaAliquotasICMS Lib "Elgin.dll" (ByVal Flag As String) As Integer
Declare Function Elgin_VerificaAliquotasIss Lib "Elgin.dll" (ByVal Flag As String) As Integer
Declare Function Elgin_VerificaDepartamentos Lib "Elgin.dll" (ByVal Departamentos As String) As Integer
Declare Function Elgin_VerificaEstadoGaveta Lib "Elgin.dll" (ByRef EstadoGaveta As Integer) As Integer
Declare Function Elgin_VerificaEstadoGavetaStr Lib "Elgin.dll" (ByVal EstadoGaveta As String) As Integer
Declare Function Elgin_VerificaEstadoImpressora Lib "Elgin.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer) As Integer
Declare Function Elgin_VerificaEstadoImpressoraMFD Lib "Elgin.dll" (ByRef ACK As Integer, ByRef ST1 As Integer, ByRef ST2 As Integer, ByRef ST3 As Integer) As Integer
Declare Function Elgin_VerificaEstadoImpressoraStr Lib "Elgin.dll" (ByVal ACK As String, ByVal ST1 As String, ByVal ST2 As String) As Integer
Declare Function Elgin_VerificaFormasPagamento Lib "Elgin.dll" (ByVal Formas As String) As Integer
Declare Function Elgin_VerificaFormasPagamentoMFD Lib "Elgin.dll" (ByVal FormasPagamento As String) As Integer
Declare Function Elgin_VerificaImpressoraLigada Lib "Elgin.dll" () As Integer
Declare Function Elgin_VerificaIndiceAliquotasICMS Lib "Elgin.dll" (ByVal Flag As String) As Integer
Declare Function Elgin_VerificaIndiceAliquotasIss Lib "Elgin.dll" (ByVal Flag As String) As Integer
Declare Function Elgin_VerificaModoOperacao Lib "Elgin.dll" (ByVal Modo As String) As Integer
Declare Function Elgin_VerificaRecebimentoNaoFiscal Lib "Elgin.dll" (ByVal Recebimentos As String) As Integer
Declare Function Elgin_VerificaRecebimentoNaoFiscalMFD Lib "Elgin.dll" (ByVal Recebimentos As String) As Integer
Declare Function Elgin_VerificaReducaoZAutomatica Lib "Elgin.dll" (ByRef Flag As Integer) As Integer
Declare Function Elgin_VerificaRelatorioGerencialMFD Lib "Elgin.dll" (ByVal Relatorios As String) As Integer
Declare Function Elgin_VerificaSensorPoucoPapelMFD Lib "Elgin.dll" (ByVal Flag As String) As Integer
Declare Function Elgin_VerificaStatusCheque Lib "Elgin.dll" (ByRef StatusCheque As Integer) As Integer
Declare Function Elgin_VerificaTipoImpressora Lib "Elgin.dll" (ByRef TipoImpressora As Integer) As Integer
Declare Function Elgin_VerificaTipoImpressoraStr Lib "Elgin.dll" (ByVal TipoImpressora As String) As Integer
Declare Function Elgin_VerificaTotalizadoresNaoFiscais Lib "Elgin.dll" (ByVal Totalizadores As String) As Integer
Declare Function Elgin_VerificaTotalizadoresNaoFiscaisMFD Lib "Elgin.dll" (ByVal Totalizadores As String) As Integer
Declare Function Elgin_VerificaTotalizadoresParciais Lib "Elgin.dll" (ByVal Totalizadores As String) As Integer
Declare Function Elgin_VerificaTotalizadoresParciaisMFD Lib "Elgin.dll" (ByVal Totalizadores As String) As Integer
Declare Function Elgin_VerificaTruncamento Lib "Elgin.dll" (ByVal Flag As String) As Integer
Declare Function Elgin_VerificaZPendente Lib "Elgin.dll" (ByRef Flag As Integer) As Integer
Declare Function Elgin_VersaoFirmware Lib "Elgin.dll" (ByVal VersaoFirmware As String) As Integer
Declare Function RFD_ConvertedaMFD Lib "Elgin.dll" (ByVal CRZ As String) As Integer
Declare Function RFD_ConvertedaMFDData Lib "Elgin.dll" (ByVal DataInicial As String, ByVal DataFinal As String) As Integer
Declare Function Wind_AcionaGaveta Lib "Elgin.dll" () As Integer
Declare Function Wind_AcionaGuilhotina Lib "Elgin.dll" (ByVal Modo As Integer) As Integer
Declare Function Wind_AcionaGuilhotinaParcial Lib "Elgin.dll" (ByVal Modo As Integer) As Integer
Declare Function Wind_AjustaLarguraPapel Lib "Elgin.dll" (ByVal LarguraPapel As Integer) As Integer
Declare Function Wind_ConfiguraCodigoBarras Lib "Elgin.dll" (ByVal Altura As Integer, ByVal Largura As Integer, ByVal PosicaoCaracteres As Integer, ByVal Fonte As Integer, ByVal Margem As Integer) As Integer
Declare Function Wind_EnviaBuffer Lib "Elgin.dll" (ByVal Buffer As String) As Integer
Declare Function Wind_EnviaBufferFormatado Lib "Elgin.dll" (ByVal Buffer As String, ByVal TipoLetra As Integer, ByVal Italico As Integer, ByVal Sublinhado As Integer, ByVal Expandido As Integer, ByVal Enfatizado As Integer) As Integer
Declare Function Wind_EnviaComando Lib "Elgin.dll" (ByVal Buffer As String, ByVal TamanhoBuffer As Integer) As Integer
Declare Function Wind_ImprimeBitmap Lib "Elgin.dll" (ByVal nomeArquivo As String, ByVal Modo As Integer) As Integer
Declare Function Wind_ImprimeBmpEspecial Lib "Elgin.dll" (ByVal nomeArquivo As String, ByVal EscalaX As Integer, ByVal EscalaY As Integer, ByVal Angulo As Integer) As Integer
Declare Function Wind_ImprimeCodigoBarrasCODABAR Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasCODE128 Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasCODE39 Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasCODE93 Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasEAN13 Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasEAN8 Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasISBN Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasITF Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasMSI Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasPDF417 Lib "Elgin.dll" (ByVal NivelCorrecaoErros As Integer, ByVal Altura As Integer, ByVal Largura As Integer, ByVal Colunas As Integer, ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasPLESSEY Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasUPCA Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_ImprimeCodigoBarrasUPCE Lib "Elgin.dll" (ByVal Codigo As String) As Integer
Declare Function Wind_VerificaEstadoGaveta Lib "Elgin.dll" () As Integer
Declare Function Wind_VerificaFimPapel Lib "Elgin.dll" () As Integer
Declare Function Wind_VerificaPoucoPapel Lib "Elgin.dll" () As Integer

'===============================================================================
'********************************************************************************
'
'                   DECLARAÇÃO DAS FUNÇÕES GLOBAIS DO DemoFit
'
'********************************************************************************
'===============================================================================}
Public Function TrataRetorno(ByVal iRetorno As Integer, Optional ByVal bMostraMsg As Boolean = True) As Boolean
    Dim strMsgErro As String
    Dim bRetorno As Boolean
    
    bRetorno = False
    If (iRetorno <> 1) Then
        Select Case iRetorno
            Case 0
                If (ObtemRetornoECF(strMsgErro) And bMostraMsg) Then
                    MsgBox strMsgErro, vbCritical, "DemoFit32 - Erro na comunicação."
                ElseIf (bMostraMsg) Then
                    MsgBox "Erro na comunicação.", vbCritical, "DemoFit32"
                End If
            Case -2
                MsgBox "Parâmetro inválido na função.", vbCritical, "DemoFit32"
            Case -4
                MsgBox "O arquivo de inicialização Elgin.ini não foi encontrado no diretório de sistema do Windows.", vbCritical, "DemoFit32"
            Case -5
                MsgBox "Erro ao abrir a porta de comunicação.", vbCritical, "DemoFit32"
            Case -27
                MsgBox "Status da impressora diferente de 6,0,0 (ACK, ByVal ST1 e ST2).", vbCritical, "DemoFit32"
            Case Else
                MsgBox "Ocorreu um erro desconhecido. Erro nº " & CStr(iRetorno), vbCritical, "DemoFit32"
        End Select
    ElseIf (bMostraMsg) Then
        MsgBox "Operação realizada com sucesso", vbOKOnly, "DemoFit"
        bRetorno = True
    End If

    AtualizaStatus
    TrataRetorno = bRetorno
End Function

Public Function ObtemRetornoECF(ByRef strMensagemErro As String) As Boolean
    Dim iRetorno As Integer
    Dim iCodErro As Integer
    Dim strErroMsg As String
    Dim bSucesso As Boolean
    
    strErroMsg = Space(100)
        
    iRetorno = Elgin_RetornoImpressora(iCodErro, strErroMsg)
    
    strMensagemErro = "Erro nº: " & CStr(iCodErro) & " - " & strErroMsg
    
    If (iRetorno = 1) Then
        bSucesso = True
    Else
        bSucesso = False
    End If
    
    ObtemRetornoECF = bSucesso

End Function

Public Sub AtualizaStatus()
    Dim iRetorno As Integer
    Dim iST1 As Integer
    Dim iST2 As Integer
    Dim iACK As Integer
    
    iRetorno = Elgin.VerificaStatusImpressora(iACK, iST1, iST2)
           
    frmDemoFit.stsBarraStatus.Panels.Item(2).Text = "ACK = " & Str(iACK) & ", ST1 = " _
    & Str(iST1) & ", ST2 = " & Str(iST2)
End Sub
