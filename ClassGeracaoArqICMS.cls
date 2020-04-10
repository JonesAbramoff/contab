VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoArqICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
''Option Explicit
''
'''???? a Inscricao Estadual está vindo com a Formatação
'''OK ??? reg 50/51/...: CGC é do remetente nas entradas e destinatario nas saidas
''
'''COLUNA NO LIVRO PARA TRIBUTACAO ICMS
''Const TIPO_TRIBUTADO = 1
''Const TIPO_ISENTO_NAO_TRIBUTADO = 2
''Const TIPO_OUTRAS = 3
''
'''CONTANTE DO TIPO DE COMPLEMENTO
''Const TIPO_COMPL_NORMAL = 0
''Const TIPO_COMPL_FRETE = 1
''Const TIPO_COMPL_SEGURO = 2
''Const TIPO_COMPL_DESCONTO = 3
''Const TIPO_COMPL_OUTRAS = 4
''
''
'''nao deve poder ter produto com codigo maior que 14 posicoes
'''
'''esta classe existe para permitir a criacao de arquivos com dados de ICMS e IPI exigida de
'''Empresas com saidas acima de R$ 1.200.000,00 com processamento eletronico de dados
'''o layout do arquivo está detalhado atraves dos convenios 57/95 e outros.
'''
'''na 1a versao vamos jogar em GeracaoArqICMSProd os registros p/evitar duplicidade desnecessaria de registros do tipo 75
'''uma versao otimizada poderia manter pool de produtos recentemente utilizados
'''O indice da tabela GeracaoArqICMS garantirá que os registros fiquem ordenados de acordo com o item 8 pg 8.
'''
'''documentacao em: m:\documentacao\sge\trb\tributos\icms\sefis
'''
'''Arquivo de ICMS:
'''Além do arquivo "detalhado" para a secretaria de fazenda local, enviar:
'''    um arquivo p/cada estado destinatario de mercadoria para as operacoes interestaduais: ffufaamm.icm onde ff é a filial, uf é a unidade de federacao(RJ,SP,...), aamm o ano e o mes, com dados a nivel de nf em que as UFs foram destinatarios.
'''        os arquivos serao criados no subdiretorio \DADOS
'''        periodo trimestral (anterior) (pg 2 e 3 da apostila)
'''
'''CFO ( Codigo Fiscal de Operacoes e Prestacoes): natureza de operacao
'''
'''STFO ( Situacao Tributaria Federal da Operacao):
''' obtida atraves do cadastro de produtos com codigos da IN SRF 142
'''    ex: produto nacional de fabricacao propria, produto nacional
'''
'''Observacoes:
'''    a principio nao vamos tratar empresas que sao substitutos tributarios pois teriamos que obter dados da GNR.
'''        Onde obter data, codigo do bco, codigo da agencia, numero e valor recolhid da GNR (guia nacional de recolhimento) (pg 3 apostila)
'''    Desfazimento, retificacao total, aditiva ou corretiva: a principio nao faremos.
'''    tratamento que ficará pendente até a implementacao do modulo de frente de loja.
'''       Por dia por maquina qdo se tratar de cupom fiscal de ECF, PDV, registradora,…
'''           As info deste tipo de operacao ainda nao estao no sistema.
'''       por dia:
'''           Nota fiscal de venda a consumidor (modelo 2): nao tem info de icms
'''
'''Registros com 126 bytes +CRLF.
'''
'''estrutura:
'''registro tipo 10
'''registro tipo 11
'''Registros 50, 51, …
'''
'''na pag 8 tem as posicoes p/sort
'''
'''registro 75: tabela de produtos
'''registro 90: totalizador.
'''
'''Registro 10:
'''    Confirmar se posso usar no campo 10 o codigo 03 para passar tudo.
'''Registro 50: p/contrib ICMS
'''    Um registro por (aliquota, CFOP)
'''atencao ao status de cancelado
'''se o "outro" nao tiver inscricao colocar "ISENTO"
'''para exterior colocar "EX"
'''conferir se series/subseries nao terao problema:
'''    no futuro podemos colocar flag p/indicar se series sao unicas, se tem subseries e se estas sao unicas.
'''
'''Nas entradas a data que vale é a de recebimento (entrada).
'''
'''Registro 51: sobre IPI, somente p/contrib deste imposto.
'''
'''Registro 53: subst trib:
'''    Somente p/o subst tribut: ex. fabrica de bebidas e nao o botequim.
'''
'''Registro 54:
'''    Registro item a item da nf
'''    Colocar 99 c/item "desconto"
'''
'''Registro 55: GNR:
'''    Somente p/o subst tribut: ex. fabrica de bebidas e nao o botequim.
'''
'''transferir esta classe e suas auxiliares p/vbp
''
''Private mvardtDataInicial As Date
''Private mvardtDataFinal As Date
''Private mvarsContato As String
''Private mvarsTelContato As String
''Private mvarsNomeArq As String
''Private mvarsEndereco As String
''Private mvarlNumero As Long
''Private mvarsComplemento As String
''Private mvarsCGCEmpresa As String
''Private mvarsNomeEmpresa As String
''
'''acumuladores p/reg do tipo 90
''Private mvarlNumReg50 As Long
''Private mvarlNumReg51 As Long
''Private mvarlNumReg53 As Long
''Private mvarlNumReg54 As Long
''Private mvarlNumReg75 As Long
''Private mvarlNumRegTotal As Long
''Private lNumIntReg As Long
''Private lNumIntGeracao As Long
''Private dTotalBach As Double
''
''Private colInfoAux As New Collection 'de objetos do tipo ClassGeracaoArqICMSAux
''
''Public Property Get dtDataInicial() As Date
''    dtDataInicial = mvardtDataInicial
''End Property
''
''Public Property Get dtDataFinal() As Date
''    dtDataFinal = mvardtDataFinal
''End Property
''
''Public Property Get sNomeArq() As String
''    sNomeArq = mvarsNomeArq
''End Property
''
''Public Property Get sContato() As String
''    sContato = mvarsContato
''End Property
''
''Public Property Get sTelContato() As String
''    sTelContato = mvarsTelContato
''End Property
''
''Public Property Get sEndereco() As String
''    sEndereco = mvarsEndereco
''End Property
''
''Public Property Get lNumero() As Long
''    lNumero = mvarlNumero
''End Property
''
''Public Property Get sComplemento() As String
''    sComplemento = mvarsComplemento
''End Property
''
''Public Property Get sCGCEmpresa() As String
''    sCGCEmpresa = mvarsCGCEmpresa
''End Property
''
''Public Property Get sNomeEmpresa() As String
''    sNomeEmpresa = mvarsNomeEmpresa
''End Property
''
''Public Function CriarArquivo(dtDataInicial As Date, dtDataFinal As Date, sContato As String, sTelContato As String, sEndereco As String, lNumero As Long, sComplemento As String, sNomeArquivo As String, sEmpresa As String, sNomeArqParam As String)
'''Rotina de Inicialização da Criação do Arquivo ICMS
''
''Dim sCRLF As String
''Dim lErro As Long
''Dim objClassBatchCPR As New ClassBatchCPR
''
''On Error GoTo Erro_CriarArquivo
''
''    sCRLF = Chr(13) & Chr(10)
''    mvardtDataInicial = dtDataInicial
''    mvardtDataFinal = dtDataFinal
''    mvarsContato = sContato
''    mvarsTelContato = sTelContato
''    mvarsEndereco = sEndereco
''    mvarlNumero = lNumero
''    mvarsComplemento = sComplemento
''    mvarsNomeArq = sNomeArquivo
''    mvarsNomeEmpresa = sEmpresa
''
''    lErro = objClassBatchCPR.Rotina_Inicia_Batch_ICMS(sNomeArqParam, Me)
''    If lErro <> SUCESSO Then Error 64001
''
''    CriarArquivo = SUCESSO
''
''    Exit Function
''
''Erro_CriarArquivo:
''
''    CriarArquivo = Err
''
''    Select Case Err
''
''        Case 64001
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150530)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Rotina_Gerar_ICMS() As Long
'''Rotina que Criará o Arquivo ICMS (ABRE A TRANSACAO)
''
''Dim objFilial As New AdmFiliais, sBuffer As String, sCRLF As String
''Dim lErro As Long
''Dim lTransacao As Long
''
''On Error GoTo Erro_Rotina_Gerar_ICMS
''
''    'abrir transacao
''    lTransacao = Transacao_Abrir()
''    If lTransacao = 0 Then Error 58782
''
''    'preencher objFilial
''    objFilial.iCodFilial = giFilialEmpresa
''
''    lErro = CF("FilialEmpresa_Le",objFilial)
''    If lErro <> SUCESSO Then Error 58783
''
''    'Quarda o CGC da Empresa
''    mvarsCGCEmpresa = objFilial.sCgc
''
''    'filtrar NFs do periodo e filial empresa selecionados, nao excluidas
''
''    'percorrer as nfs e seus itens obtendo info p/montagem de regs 50, 51, 54 e 75.
''    lErro = ProcessarNFs()
''    If lErro <> SUCESSO Then Error 58784
''
''    lErro = InserirRegsArqICMS(objFilial)
''    If lErro <> SUCESSO Then Error 58785
''
''    'fazer rollback: defaz todas as insercoes nas tabelas auxiliares
''    Call Transacao_Rollback
''
''    Rotina_Gerar_ICMS = SUCESSO
''
''    Exit Function
''
''Erro_Rotina_Gerar_ICMS:
''
''    Rotina_Gerar_ICMS = Err
''
''    Select Case Err
''
''        Case 58782
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
''
''        Case 58783, 58784, 58785
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150531)
''
''    End Select
''
''    'fazer rollback: defaz todas as insercoes nas tabelas auxiliares
''    Call Transacao_Rollback
''
''    Exit Function
''
''End Function
''
''Private Function MontarReg10(sBuffer As String, objFilial As AdmFiliais) As Long
'''preenche sBuffer com registro do tipo 10
''
''Dim lErro As Long
''
''On Error GoTo Erro_MontarReg10
''
''    'usar objFilial para obter CGC, Inscr Est, Nome, Municipio, UF, Fax
''    objFilial.iCodFilial = giFilialEmpresa
''
''    lErro = CF("FilialEmpresa_Le",objFilial)
''    If lErro <> SUCESSO Then Error 53086
''
''    'OK ??? Ler a razao social no BD. O nome da filial nao é a razao social da empresa
''
''    'data inicial e final da propria classe
''    'codigo de Identificacao do Convenio: 1
''    'Codigo da identificacao do natureza das operacoes:3
''    'finalidade: 1
''    sBuffer = "10" & FormataCpoNum(objFilial.sCgc, 14) & FormataCpoTexto(objFilial.sInscricaoEstadual, 14) & FormataCpoTexto(mvarsNomeEmpresa, 35) & FormataCpoTexto(objFilial.objEndereco.sCidade, 30) & FormataCpoTexto(objFilial.objEndereco.sSiglaEstado, 2) & FormataCpoNum(objFilial.objEndereco.sFax, 10) & FormataCpoData(mvardtDataInicial) & FormataCpoData(mvardtDataFinal) & "1" & "3" & "1"  'OK ??? acho que o campo 10 deveria ter o valor 1 e vc colocou 0. OK
''
''    MontarReg10 = SUCESSO
''
''    Exit Function
''
''Erro_MontarReg10:
''
''    Select Case Err
''
''        Case 53086
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150532)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Private Function MontarReg11(sBuffer As String, objFilial As AdmFiliais) As Long
'''preenche sBuffer com registro do tipo 11
''
''Dim lErro As Long
''
''On Error GoTo Erro_MontarReg11
''
''    'usar objFilial para obter CEP e endereco
''    'obter contato e telefone da classe ClassGeracaoArqICMS (esta classe)
''
''    objFilial.iCodFilial = giFilialEmpresa
''
''    lErro = CF("FilialEmpresa_Le",objFilial)
''    If lErro <> SUCESSO Then Error 53085
''
''    sBuffer = "11" & FormataCpoTexto(mvarsEndereco, 34) & FormataCpoNum(mvarlNumero, 5) & FormataCpoTexto(mvarsComplemento, 22) & FormataCpoTexto(objFilial.objEndereco.sBairro, 15) & FormataCpoTexto(objFilial.objEndereco.sCEP, 8) & FormataCpoTexto(mvarsContato, 28) & FormataCpoNum(mvarsTelContato, 12)
''
''    MontarReg11 = SUCESSO
''
''    Exit Function
''
''Erro_MontarReg11:
''
''    Select Case Err
''
''        Case 53085
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150533)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Private Function MontarReg50(sBuffer As String, sChaveOrd As String, tTipo50 As typeTipo50) As Long
'''preenche sBuffer com registro do tipo 50
''
''' info sobre ICMS ref a NF, NF de entrada
'''percorrer NFiscal e NFiscalBaixadas filialfornecedor ou filialcliente, dependendo de qual estiver preenchido.
'''Obter o cgc e inscr estadual na tabela de
'''se o "outro" nao tiver inscricao colocar "ISENTO"
'''se nao tiver cgc, zera-lo
'''se a data de entrada <> DATA_NULA, usa-la, senao usar data de emissao (esta data será utilizada para ver se o movimento está dentro do periodo)
'''para exterior colocar "EX"
'''Um registro por (aliquota, CFOP)
'''atencao ao status de cancelado
'''conferir se series/subseries nao terao problema:
'''    no futuro podemos colocar flag p/indicar se series sao unicas, se tem subseries e se estas sao unicas.
'''modelo: 01 (nfiscal) ou 03 nfiscal de entrada: pegar coluna de tipodocinfo
'''nao considerar subst tributaria
''
''    sBuffer = "50" & FormataCpoNum(tTipo50.sCgc, 14) & FormataCpoTexto(tTipo50.sInscricaoEstadual, 14) & FormataCpoData(tTipo50.dtDataEmissao) & FormataCpoTexto(tTipo50.sUnidadeFederacao, 2) & FormataCpoNum(tTipo50.iModelo, 2) & FormataCpoTexto(tTipo50.sSerie, 2) & "   " & FormataCpoNum(tTipo50.lNumeroNF, 6) & FormataCpoNum(tTipo50.sCFOP, 3) & FormataCpoValor(tTipo50.dValorTotal, 13) & FormataCpoValor(tTipo50.dBaseICMS, 13) & FormataCpoValor(tTipo50.dValorICMS, 13) & FormataCpoValor(tTipo50.dIsentaNTributada, 13) & FormataCpoValor(tTipo50.dOutras, 13) & FormataCpoValor(tTipo50.dAliquotaICMS * 100, 4) & FormataCpoTexto(tTipo50.sSituacao, 1)
''
''    sChaveOrd = FormataCpoData(tTipo50.dtDataEmissao)
''
''    MontarReg50 = SUCESSO
''
''End Function
''
''Private Function MontarReg51(sBuffer As String, sChaveOrd As String, tTipo51 As typeTipo51) As Long
'''preenche sBuffer com registro do tipo 51
''
''    sBuffer = "51" & FormataCpoNum(tTipo51.sCgc, 14) & FormataCpoTexto(tTipo51.sInscricaoEstadual, 14) & FormataCpoData(tTipo51.dtDataEmissao) & FormataCpoTexto(tTipo51.sUnidadeFederacao, 2) & FormataCpoTexto(tTipo51.sSerie, 2) & FormataCpoTexto("", 2) & FormataCpoNum(tTipo51.lNumeroNF, 6) & FormataCpoNum(tTipo51.sCFOP, 3) & FormataCpoValor(tTipo51.dValorTotal, 13) & FormataCpoValor(tTipo51.dValorIPI, 13) & FormataCpoValor(tTipo51.dIsentaNTributada, 13) & FormataCpoValor(tTipo51.dOutras, 13) & FormataCpoTexto("", 20) & FormataCpoTexto(tTipo51.sSituacao, 1)
''
''    sChaveOrd = FormataCpoData(tTipo51.dtDataEmissao)
''
''    MontarReg51 = SUCESSO
''
''End Function
''
''Private Function MontarReg54(sBuffer As String, sChaveOrd As String, tTipo54 As typeTipo54) As Long
'''preenche sBuffer com registro do tipo 54
''
''    sBuffer = "54" & FormataCpoNum(tTipo54.sCgc, 14) & FormataCpoNum(tTipo54.iModelo, 2) & FormataCpoTexto(tTipo54.sSerie, 2) & "   " & FormataCpoNum(tTipo54.lNumeroNF, 6) & FormataCpoNum(tTipo54.sCFOP, 3) & FormataCpoNum(tTipo54.iItem, 3) & FormataCpoTexto(tTipo54.sCodigoProduto, 14) & FormataCpoQuantidade(tTipo54.dQuantidade, 13) & FormataCpoValor(tTipo54.dValorProduto, 12) & FormataCpoValor(tTipo54.dValorDescontos, 12) & FormataCpoValor(tTipo54.dBaseICMS, 12) & FormataCpoValor(tTipo54.dBaseICMSSubst, 12) & FormataCpoValor(tTipo54.dValorIPI, 12) & FormataCpoValor(tTipo54.dAliquotaICMS * 100, 4)
''    sChaveOrd = FormataCpoNum(tTipo54.sCgc, 14) & FormataCpoTexto(tTipo54.sSerie, 3) & FormataCpoNum(tTipo54.lNumeroNF, 6) & FormataCpoNum(tTipo54.iItem, 3)
''
''    MontarReg54 = SUCESSO
''
''End Function
''
''Private Function MontarReg75(sBuffer As String, tTipo75 As typeTipo75) As Long
'''preenche sBuffer com registro do tipo 75
''
''    sBuffer = "75" & FormataCpoData(mvardtDataInicial) & FormataCpoData(mvardtDataFinal) & FormataCpoTexto(tTipo75.sCodigoProduto, 14) & FormataCpoTexto(tTipo75.sCodigoNCM, 8) & FormataCpoTexto(tTipo75.sDescricaoProduto, 53) & FormataCpoTexto(tTipo75.sUnidadeMedida, 6) & tTipo75.sSituacaoTributaria & "0" & FormataCpoNum(tTipo75.dAliquotaIPI, 4) & FormataCpoNum(tTipo75.dAliquotaICMS, 4) & FormataCpoNum(tTipo75.dReducaoBaseICMS, 4) & FormataCpoValor(tTipo75.dBaseICMSSubst, 12)
''
''    MontarReg75 = SUCESSO
''
''End Function
''
''Private Function MontarReg90(sBuffer As String, sCgc As String, sInscricaostadual As String) As Long
'''preenche sBuffer com registro do tipo 90
''
''    sBuffer = "90" & FormataCpoNum(sCgc, 14) & FormataCpoTexto(sInscricaostadual, 14)
''
''    If mvarlNumReg50 > 0 Then
''        sBuffer = sBuffer & "50" & FormataCpoNum(mvarlNumReg50, 8)
''    End If
''
''    If mvarlNumReg51 > 0 Then
''        sBuffer = sBuffer & "51" & FormataCpoNum(mvarlNumReg51, 8)
''    End If
''
''    If mvarlNumReg54 > 0 Then
''        sBuffer = sBuffer & "54" & FormataCpoNum(mvarlNumReg54, 8)
''    End If
''
''    If mvarlNumReg75 > 0 Then
''        sBuffer = sBuffer & "75" & FormataCpoNum(mvarlNumReg75, 8)
''    End If
''
''    mvarlNumRegTotal = mvarlNumReg50 + mvarlNumReg51 + mvarlNumReg54 + mvarlNumReg75 + 3
''
''    sBuffer = sBuffer & "99" & FormataCpoNum(mvarlNumRegTotal, 8)
''
''    sBuffer = FormataCpoTexto(sBuffer, 125)
''    sBuffer = sBuffer & "1"
''
''    MontarReg90 = SUCESSO
''
''End Function
''
''Private Function ProcessarNFs() As Long
'''percorrer as nfs e seus itens obtendo info p/montagem de regs 50, 51, 54 e 75.
'''a cada registro lido criar um reg54 e, se nao existir, incluir reg em GeracaoArqICMSProd e reg75 em GeracaoArqICMS
'''o(s) registro(s) 50 de uma nf somente poderão ser gerados após a leitura de todos os regs dos itens pois só entao saberei o total de cada aliquota
'''ir mantendo em colInfoAux os acumuladores por (aliquota, CFOP)
'''apos ler o ultimo reg da nf criar os regs 50 e o reg 51
''
''Dim lErro As Long
''
''On Error GoTo Erro_ProcessarNFs
''
''    lNumIntReg = 0
''    lNumIntGeracao = 0
''    dTotalBach = 0
''
''    'Processa as Notas para os Clientes
''    lErro = ProcessarNFs_Cliente()
''    If lErro <> SUCESSO Then Error 58817
''
''    'Processa as Notas para os Fornecedores
''    lErro = ProcessarNFs_Fornecedor()
''    If lErro <> SUCESSO Then Error 58818
''
''    ProcessarNFs = SUCESSO
''
''    Exit Function
''
''Erro_ProcessarNFs:
''
''    ProcessarNFs = Err
''
''    Select Case Err
''
''        Case 58817, 58818
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150534)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Private Function ProcessarNFs_Cliente() As Long
''
''Dim lErro As Long
''Dim alComando(6) As Long
''Dim tTipoGenericoCliente As typeTipo
''Dim tTipo50 As typeTipo50
''Dim tTipo51 As typeTipo51
''Dim tTipo54 As typeTipo54
''Dim sPais As String
''Dim colInfoAux As Collection
''Dim objGeracaoArqICMSAux As ClassGeracaoArqICMSAux
''Dim objGeracaoArqICMSProd As New ClassGeracaoArqICMSProd
''Dim sBuffer As String
''Dim sChaveOrd As String
''Dim iIndice As Integer
''Dim iCancelou As Integer
''Dim vbMsgBox As VbMsgBoxResult
''
''On Error GoTo Erro_ProcessarNFs_Cliente
''
''    'Abertura comandos
''    For iIndice = LBound(alComando) To UBound(alComando)
''        alComando(iIndice) = Comando_Abrir()
''        If alComando(iIndice) = 0 Then Error 58786
''    Next
''
''    Set colInfoAux = New Collection
''
''    tTipoGenericoCliente.sSerie = String(STRING_SERIE, 0)
''    tTipoGenericoCliente.sCgc = String(STRING_CGC, 0)
''    tTipoGenericoCliente.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
''    tTipoGenericoCliente.sUnidadeFederacao = String(STRING_ESTADO, 0)
''
''    'OK ??? por que o order by no SELECT abaixo ? OK
''    'OK ??? por que nao le a data de entrada que será necessaria p/recebimentos ? OK
''    'OK ??? Acho que o criterio de datas para selecao das nfs deve considerar a data de emissao p/saidas e data de entrada p/entradas OK
''
''    'Primeiro faz um select nas Notas que tenham cliente <> 0 dentro do periodo...
''    lErro = Comando_Executar(alComando(1), "SELECT NFiscal.NumIntDoc, NFiscal.Serie, NFiscal.NumNotaFiscal, NFiscal.ValorDesconto, NFiscal.ValorTotal, NFiscal.ValorFrete, NFiscal.ValorSeguro, NFiscal.ValorOutrasDespesas, NFiscal.DataEmissao, NFiscal.Status, FiliaisClientes.CGC, FiliaisClientes.InscricaoEstadual, Enderecos.SiglaEstado, Enderecos.CodigoPais, TiposDocInfo.ModeloArqICMS, TiposDocInfo.Emitente, TiposDocInfo.Destinatario, TiposDocInfo.Tipo FROM NFiscal, FiliaisClientes, Enderecos, TiposDocInfo WHERE NFiscal.Cliente <> ? AND NFiscal.FilialEmpresa = ? AND ((TiposDocInfo.Tipo = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ?) OR ((TiposDocInfo.Tipo = ? OR TiposDocInfo.Tipo = ?) AND NFiscal.DataEntrada >= ? AND NFiscal.DataEntrada <= ?)) AND NFiscal.Cliente = FiliaisClientes.CodCliente AND NFiscal.FilialCli = FiliaisClientes.CodFilial AND FiliaisClientes.Endereco = Enderecos.Codigo AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo", _
''            tTipoGenericoCliente.lNumIntDocNF, tTipoGenericoCliente.sSerie, tTipoGenericoCliente.lNumeroNF, tTipoGenericoCliente.dValorDescontoTotal, tTipoGenericoCliente.dValorTotal, tTipoGenericoCliente.dValorFrete, tTipoGenericoCliente.dValorSeguro, tTipoGenericoCliente.dValorOutras, tTipoGenericoCliente.dtDataEmissao, tTipoGenericoCliente.iStatus, tTipoGenericoCliente.sCgc, tTipoGenericoCliente.sInscricaoEstadual, tTipoGenericoCliente.sUnidadeFederacao, tTipoGenericoCliente.iCodigoPais, tTipoGenericoCliente.iModelo, tTipoGenericoCliente.iEmitente, tTipoGenericoCliente.iDestinatario, tTipoGenericoCliente.iTipo, 0, giFilialEmpresa, TIPODOCINFO_TIPO_NFIS, mvardtDataInicial, mvardtDataFinal, TIPODOCINFO_TIPO_NFIE, TIPODOCINFO_TIPO_NFEXT, mvardtDataInicial, mvardtDataFinal)
''    If lErro <> AD_SQL_SUCESSO Then Error 58787
''
''    lErro = Comando_BuscarPrimeiro(alComando(1))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58788
''
''    Do While lErro = SUCESSO
''
''        dTotalBach = dTotalBach + 1
''
''        Call TelaAcompanhaBatchCPR.Acompanha_Batch_CPR(dTotalBach, iCancelou)
''
''        If iCancelou = CANCELA_BATCH Then
''            vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_GERACAO_ARQ_ICMS")
''            If vbMsgBox = vbYes Then Error 61491
''        End If
''
''        tTipoGenericoCliente.sUnidMedida = String(STRING_UM_SIGLA, 0)
''        tTipoGenericoCliente.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
''        tTipoGenericoCliente.sDescricaoProduto = String(STRING_PRODUTO_DESCRICAO, 0)
''        tTipoGenericoCliente.sCodigoProduto = String(STRING_PRODUTO, 0)
''        tTipoGenericoCliente.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
''
''        'Lê os itens da Nota para os Clientes
''        lErro = Comando_Executar(alComando(3), "SELECT ItensNFiscal.ValorDesconto, ItensNFiscal.Quantidade, ItensNFiscal.Item, ItensNFiscal.UnidadeMed,  ItensNFiscal.PrecoUnitario, TributacaoItemNF.NaturezaOp, TributacaoItemNF.ICMSBase, TributacaoItemNF.ICMSValor, TributacaoItemNF.ICMSSubstBase, TributacaoItemNF.IPIValor, TributacaoItemNF.ICMSAliquota , TiposTribICMS.ColunaNoLivroEntrada, TiposTribICMS.Tipo,  Produtos.IPICodigo, Produtos.Codigo, Produtos.Descricao, TiposTribIPI.ColunaNoLivro FROM ItensNFiscal, TributacaoItemNF, TiposTribICMS, Produtos, TiposTribIPI  WHERE ItensNFiscal.NumIntNF = ?  AND ItensNFiscal.NumIntDoc = TributacaoItemNF.NumIntDoc AND TributacaoItemNF.ICMSTipo = TiposTribICMS.Tipo AND TributacaoItemNF.ICMSTipo = TiposTribIPI.Tipo AND Produtos.Codigo = ItensNFiscal.Produto ORDER BY ItensNFiscal.NumIntDoc ASC", _
''                tTipoGenericoCliente.dValorDescontoItem, tTipoGenericoCliente.dQuantidade, tTipoGenericoCliente.iItem, tTipoGenericoCliente.sUnidMedida, tTipoGenericoCliente.dPrecoUnitario, tTipoGenericoCliente.sNaturezaOp, tTipoGenericoCliente.dBaseICMS, tTipoGenericoCliente.dValorICMS, tTipoGenericoCliente.dICMSSubstBase, tTipoGenericoCliente.dValorIPI, tTipoGenericoCliente.dICMSAliquota, tTipoGenericoCliente.iColunaNoLivroICMS, tTipoGenericoCliente.iTipoTribICMS, tTipoGenericoCliente.sIPICodigo, tTipoGenericoCliente.sCodigoProduto, tTipoGenericoCliente.sDescricaoProduto, tTipoGenericoCliente.iColunaNoLivroIPI, tTipoGenericoCliente.lNumIntDocNF)
''        If lErro <> AD_SQL_SUCESSO Then Error 58789
''
''        lErro = Comando_BuscarPrimeiro(alComando(3))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58790
''
''        Do While lErro = SUCESSO
''
''            'Transfere os dados do tTipoGenericoCliente para tTipo54
''            Call Organiza_TypeTipo54(tTipo54, tTipoGenericoCliente, TIPO_COMPL_NORMAL)
''
''            'e acumula os valores que convem
''            Call Acumula_Valores(colInfoAux, tTipoGenericoCliente, TIPO_COMPL_NORMAL)
''
''            'Gera o Tipo 54
''            Call MontarReg54(sBuffer, sChaveOrd, tTipo54)
''
''            'Gera o NumInt do tipo
''            lNumIntReg = lNumIntReg + 1
''            'Conta o tipo
''            mvarlNumReg54 = mvarlNumReg54 + 1
''
''            'OK ??? sugiro manter comando(s) aberto(s) p/chamar funcao abaixo OK
''            'e Grava no BD o Tipo 54
''            lErro = GeracaoArqICMS_Grava(sBuffer, sChaveOrd, 54, lNumIntReg, alComando(4))
''            If lErro <> SUCESSO Then Error 58791
''
''            'e Grava o prototipo do Tipo 75 no BD para evitar repeticoes
''            objGeracaoArqICMSProd.sCodigo = tTipoGenericoCliente.sCodigoProduto
''            objGeracaoArqICMSProd.sDescricao = tTipoGenericoCliente.sDescricaoProduto
''            objGeracaoArqICMSProd.sUnidadeMedida = tTipoGenericoCliente.sUnidMedida
''            objGeracaoArqICMSProd.iTipoTribICMS = tTipoGenericoCliente.iTipoTribICMS
''
''            'OK ??? sugiro manter comando(s) aberto(s) p/chamar funcao abaixo
''            lErro = CF("GeracaoArqICMSProd_Grava",lNumIntGeracao, objGeracaoArqICMSProd, alComando(5), alComando(6))
''            If lErro <> SUCESSO Then Error 58792
''
''            'Busca o proximo Item
''            lErro = Comando_BuscarProximo(alComando(3))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58793
''
''        Loop
''
''        'Limpa alguns valores
''        Call Limpa_Tipo(tTipoGenericoCliente)
''
''        tTipoGenericoCliente.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
''
''        'Lê os Complementos
''        lErro = Comando_Executar(alComando(2), "SELECT TributacaoComplNF.NaturezaOp, TributacaoComplNF.Tipo, TributacaoComplNF.ICMSValor, TributacaoComplNF.ICMSAliquota, TributacaoComplNF.ICMSBase, TributacaoComplNF.ICMSSubstBase, TributacaoComplNF.IPIValor, TiposTribICMS.ColunaNoLivroEntrada, TiposTribIPI.ColunaNoLivro FROM TributacaoComplNF, TiposTribICMS, TiposTribIPI WHERE TributacaoComplNF.NumIntDoc = ? AND TributacaoComplNF.ICMSTipo = TiposTribICMS.Tipo AND TributacaoComplNF.IPITipo = TiposTribIPI.Tipo AND TributacaoComplNF.ICMSBase > 0 ORDER BY TributacaoComplNF.NumIntDoc ASC", tTipoGenericoCliente.sNaturezaOp, tTipoGenericoCliente.iTipoComplemento, tTipoGenericoCliente.dValorICMS, tTipoGenericoCliente.dICMSAliquota, tTipoGenericoCliente.dBaseICMS, tTipoGenericoCliente.dICMSSubstBase, tTipoGenericoCliente.dValorIPI, tTipoGenericoCliente.iColunaNoLivroICMS, tTipoGenericoCliente.iColunaNoLivroIPI, tTipoGenericoCliente.lNumIntDocNF)
''        If lErro <> AD_SQL_SUCESSO Then Error 58794
''
''        lErro = Comando_BuscarPrimeiro(alComando(2))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58795
''
''        Do While lErro = SUCESSO
''
''            'Transfere os dados do tTipoGenericoCliente para tTipo54
''            Call Organiza_TypeTipo54(tTipo54, tTipoGenericoCliente, tTipoGenericoCliente.iTipoComplemento)
''
''            'e acumula os valores que convem para a geracao do tipo 50 e 51
''            Call Acumula_Valores(colInfoAux, tTipoGenericoCliente, tTipoGenericoCliente.iTipoComplemento)
''
''            'Gera o Tipo 54
''            Call MontarReg54(sBuffer, sChaveOrd, tTipo54)
''
''            'Gera o NumInt do Registro
''            lNumIntReg = lNumIntReg + 1
''
''            'Conta o Tipo 54 para geracao do Tipo 90
''            mvarlNumReg54 = mvarlNumReg54 + 1
''
''            'OK ??? sugiro manter comando(s) aberto(s) p/chamar funcao abaixo
''
''            'e Grava no BD o Tipo 54
''            lErro = GeracaoArqICMS_Grava(sBuffer, sChaveOrd, 54, lNumIntReg, alComando(4))
''            If lErro <> SUCESSO Then Error 58796
''
''            lErro = Comando_BuscarProximo(alComando(2))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58797
''
''        Loop
''
''        'Depois que percorreu todos os itens e seus complementos gera os tipos 50, 51 e pega a proxima nota
''
''        'Para cada objGeracaoArqICMSAux da colecao gera um tipo 50 e 51
''        For Each objGeracaoArqICMSAux In colInfoAux
''
''            'Transfere do tipoGenericoCliente e objGeracaoArqICMSAux para o Tipo 50
''            Call Organiza_TypeTipo50(tTipo50, tTipoGenericoCliente, objGeracaoArqICMSAux)
''
''            'Monta o Buffer
''            Call MontarReg50(sBuffer, sChaveOrd, tTipo50)
''
''            'Gera o NunInt do Registro
''            lNumIntReg = lNumIntReg + 1
''
''            'Conta o Reistro 50 para gerar o Tipo 90
''            mvarlNumReg50 = mvarlNumReg50 + 1
''
''            'OK ??? sugiro manter comando(s) aberto(s) p/chamar funcao abaixo
''
''            'e Grava no BD o Tipo 50
''            lErro = GeracaoArqICMS_Grava(sBuffer, sChaveOrd, 50, lNumIntReg, alComando(4))
''            If lErro <> SUCESSO Then Error 58799
''
''            'Transfere do tipoGenericoCliente e objGeracaoArqICMSAux para o Tipo 51
''            Call Organiza_TypeTipo51(tTipo51, tTipoGenericoCliente, objGeracaoArqICMSAux)
''
''            'Monta o Buffer
''            Call MontarReg51(sBuffer, sChaveOrd, tTipo51)
''
''            'Gera o NunInt do Registro
''            lNumIntReg = lNumIntReg + 1
''
''            'Conta o Reistro 51 para gerar o Tipo 90
''            mvarlNumReg51 = mvarlNumReg51 + 1
''
''            'OK ??? sugiro manter comando(s) aberto(s) p/chamar funcao abaixo
''
''            'e Grava no BD o Tipo 51
''            lErro = GeracaoArqICMS_Grava(sBuffer, sChaveOrd, 51, lNumIntReg, alComando(4))
''            If lErro <> SUCESSO Then Error 58800
''
''        Next
''
''        Set colInfoAux = New Collection
''
''        'Pega a Proxima Nota Fiscal
''        lErro = Comando_BuscarProximo(alComando(1))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58801
''
''    Loop
''
''
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    ProcessarNFs_Cliente = SUCESSO
''
''    Exit Function
''
''Erro_ProcessarNFs_Cliente:
''
''    ProcessarNFs_Cliente = Err
''
''    Select Case Err
''
''        Case 58786
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 58787, 58788, 58801
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)
''
''        Case 58789, 58790, 58793
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", Err)
''
''        Case 58791, 58792, 58796, 58799, 58800
''
''        Case 58794, 58795, 58797
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLNF", Err)
''
''        Case 61491 'Cancelou o batch
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150535)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Private Function ProcessarNFs_Fornecedor() As Long
''
''Dim lErro As Long
''Dim alComando(6) As Long
''Dim tTipoGenericoFornecedores As typeTipo
''Dim tTipo50 As typeTipo50
''Dim tTipo51 As typeTipo51
''Dim tTipo54 As typeTipo54
''Dim sPais As String
''Dim colInfoAux As Collection
''Dim objGeracaoArqICMSAux As ClassGeracaoArqICMSAux
''Dim objGeracaoArqICMSProd As New ClassGeracaoArqICMSProd
''Dim sBuffer As String
''Dim sChaveOrd As String
''Dim iIndice As Integer
''Dim iCancelou As Integer
''Dim vbMsgBox As VbMsgBoxResult
''
''On Error GoTo Erro_ProcessarNFs_Fornecedor
''
''    'Abertura comandos
''    For iIndice = LBound(alComando) To UBound(alComando)
''        alComando(iIndice) = Comando_Abrir()
''        If alComando(iIndice) = 0 Then Error 58802
''    Next
''
''    Set colInfoAux = New Collection
''
''    tTipoGenericoFornecedores.sSerie = String(STRING_SERIE, 0)
''    tTipoGenericoFornecedores.sCgc = String(STRING_CGC, 0)
''    tTipoGenericoFornecedores.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
''    tTipoGenericoFornecedores.sUnidadeFederacao = String(STRING_ESTADO, 0)
''
''    '??? OK por que o order by no SELECT abaixo? OK
''    '??? OK por que nao le a data de entrada que será necessaria p/recebimentos? OK
''    '??? OK Acho que o criterio de datas para selecao das nfs deve considerar a data de emissao p/saidas e data de entrada p/entradas. OK
''
''    'Primeiro faz um select na Nota que tenham Fornecedores <> 0
''    lErro = Comando_Executar(alComando(1), "SELECT NFiscal.NumIntDoc, NFiscal.Serie, NFiscal.NumNotaFiscal, NFiscal.ValorDesconto, NFiscal.ValorTotal, NFiscal.ValorFrete, NFiscal.ValorSeguro, NFiscal.ValorOutrasDespesas, NFiscal.DataEmissao, NFiscal.Status, FiliaisFornecedores.CGC, FiliaisFornecedores.InscricaoEstadual, Enderecos.SiglaEstado, Enderecos.CodigoPais, TiposDocInfo.ModeloArqICMS, TiposDocInfo.Emitente, TiposDocInfo.Destinatario, TiposDocInfo.Tipo FROM NFiscal, FiliaisFornecedores, Enderecos, TiposDocInfo WHERE NFiscal.Fornecedor <> ? AND NFiscal.FilialEmpresa = ? AND ((TiposDocInfo.Tipo = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ?) OR ((TiposDocInfo.Tipo = ? OR TiposDocInfo.Tipo = ?) AND NFiscal.DataEntrada >= ? AND NFiscal.DataEntrada <= ?)) AND NFiscal.Fornecedor = FiliaisFornecedores.CodFornecedor AND NFiscal.FilialForn = FiliaisFornecedores.CodFilial AND FiliaisFornecedores.Endereco = Enderecos.Codigo AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo", _
''            tTipoGenericoFornecedores.lNumIntDocNF, tTipoGenericoFornecedores.sSerie, tTipoGenericoFornecedores.lNumeroNF, tTipoGenericoFornecedores.dValorDescontoTotal, tTipoGenericoFornecedores.dValorTotal, tTipoGenericoFornecedores.dValorFrete, tTipoGenericoFornecedores.dValorSeguro, tTipoGenericoFornecedores.dValorOutras, tTipoGenericoFornecedores.dtDataEmissao, tTipoGenericoFornecedores.iStatus, tTipoGenericoFornecedores.sCgc, tTipoGenericoFornecedores.sInscricaoEstadual, tTipoGenericoFornecedores.sUnidadeFederacao, tTipoGenericoFornecedores.iCodigoPais, tTipoGenericoFornecedores.iModelo, tTipoGenericoFornecedores.iEmitente, tTipoGenericoFornecedores.iDestinatario, tTipoGenericoFornecedores.iTipo, 0, giFilialEmpresa, TIPODOCINFO_TIPO_NFIS, mvardtDataInicial, mvardtDataFinal, TIPODOCINFO_TIPO_NFIE, TIPODOCINFO_TIPO_NFEXT, mvardtDataInicial, mvardtDataFinal)
''    If lErro <> AD_SQL_SUCESSO Then Error 58803
''
''    lErro = Comando_BuscarPrimeiro(alComando(1))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58804
''
''    Do While lErro = SUCESSO
''
''        dTotalBach = dTotalBach + 1
''
''        Call TelaAcompanhaBatchCPR.Acompanha_Batch_CPR(dTotalBach, iCancelou)
''
''        If iCancelou = CANCELA_BATCH Then
''            vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_GERACAO_ARQ_ICMS")
''            If vbMsgBox = vbYes Then Error 61492
''        End If
''
''        tTipoGenericoFornecedores.sUnidMedida = String(STRING_UM_SIGLA, 0)
''        tTipoGenericoFornecedores.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
''        tTipoGenericoFornecedores.sDescricaoProduto = String(STRING_PRODUTO_DESCRICAO, 0)
''        tTipoGenericoFornecedores.sCodigoProduto = String(STRING_PRODUTO, 0)
''        tTipoGenericoFornecedores.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
''
''        'Lê os itens da Nota
''        lErro = Comando_Executar(alComando(3), "SELECT ItensNFiscal.ValorDesconto, ItensNFiscal.Quantidade, ItensNFiscal.Item, ItensNFiscal.UnidadeMed,  ItensNFiscal.PrecoUnitario, TributacaoItemNF.NaturezaOp, TributacaoItemNF.ICMSBase, TributacaoItemNF.ICMSValor, TributacaoItemNF.ICMSSubstBase, TributacaoItemNF.IPIValor, TributacaoItemNF.ICMSAliquota , TiposTribICMS.ColunaNoLivroSaida, TiposTribICMS.Tipo,  Produtos.IPICodigo, Produtos.Codigo, Produtos.Descricao, TiposTribIPI.ColunaNoLivro FROM ItensNFiscal, TributacaoItemNF, TiposTribICMS, Produtos, TiposTribIPI  WHERE ItensNFiscal.NumIntNF = ?  AND ItensNFiscal.NumIntDoc = TributacaoItemNF.NumIntDoc AND TributacaoItemNF.ICMSTipo = TiposTribICMS.Tipo AND TributacaoItemNF.ICMSTipo = TiposTribIPI.Tipo AND Produtos.Codigo = ItensNFiscal.Produto ORDER BY ItensNFiscal.NumIntDoc ASC", _
''                tTipoGenericoFornecedores.dValorDescontoItem, tTipoGenericoFornecedores.dQuantidade, tTipoGenericoFornecedores.iItem, tTipoGenericoFornecedores.sUnidMedida, tTipoGenericoFornecedores.dPrecoUnitario, tTipoGenericoFornecedores.sNaturezaOp, tTipoGenericoFornecedores.dBaseICMS, tTipoGenericoFornecedores.dValorICMS, tTipoGenericoFornecedores.dICMSSubstBase, tTipoGenericoFornecedores.dValorIPI, tTipoGenericoFornecedores.dICMSAliquota, tTipoGenericoFornecedores.iColunaNoLivroICMS, tTipoGenericoFornecedores.iTipoTribICMS, tTipoGenericoFornecedores.sIPICodigo, tTipoGenericoFornecedores.sCodigoProduto, tTipoGenericoFornecedores.sDescricaoProduto, tTipoGenericoFornecedores.iColunaNoLivroIPI, tTipoGenericoFornecedores.lNumIntDocNF)
''        If lErro <> AD_SQL_SUCESSO Then Error 58805
''
''        lErro = Comando_BuscarPrimeiro(alComando(3))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58806
''
''        Do While lErro = SUCESSO
''
''            'Transfere os dados do tTipoGenericoFornecedores para tTipo54
''            Call Organiza_TypeTipo54(tTipo54, tTipoGenericoFornecedores, TIPO_COMPL_NORMAL)
''
''            'Acumula Valores para a Geracao do Tipo 50 e 51
''            Call Acumula_Valores(colInfoAux, tTipoGenericoFornecedores, TIPO_COMPL_NORMAL)
''
''            'Gera o Tipo 54
''            Call MontarReg54(sBuffer, sChaveOrd, tTipo54)
''
''            'Gera o NumInt do Registro
''            lNumIntReg = lNumIntReg + 1
''
''            'Conta o registro 54
''            mvarlNumReg54 = mvarlNumReg54 + 1
''
''            'OK ??? sugiro manter comando(s) aberto(s) p/chamar funcao abaixo
''
''            'e Grava no BD o Tipo 54
''            lErro = GeracaoArqICMS_Grava(sBuffer, sChaveOrd, 54, lNumIntReg, alComando(4))
''            If lErro <> SUCESSO Then Error 58807
''
''            'e Grava o prototipo do Tipo 75 no BD para evitar repeticoes
''            objGeracaoArqICMSProd.sCodigo = tTipoGenericoFornecedores.sCodigoProduto
''            objGeracaoArqICMSProd.sDescricao = tTipoGenericoFornecedores.sDescricaoProduto
''            objGeracaoArqICMSProd.sUnidadeMedida = tTipoGenericoFornecedores.sUnidMedida
''            objGeracaoArqICMSProd.iTipoTribICMS = tTipoGenericoFornecedores.iTipoTribICMS
''
''            'OK ??? sugiro manter comando(s) aberto(s) p/chamar funcao abaixo OK
''
''            lErro = CF("GeracaoArqICMSProd_Grava",lNumIntGeracao, objGeracaoArqICMSProd, alComando(5), alComando(6))
''            If lErro <> SUCESSO Then Error 58808
''
''            lErro = Comando_BuscarProximo(alComando(3))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58809
''
''        Loop
''
''        'Limpa alguns valores
''        Call Limpa_Tipo(tTipoGenericoFornecedores)
''
''        tTipoGenericoFornecedores.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
''
''        'Lê os Complementos
''        lErro = Comando_Executar(alComando(2), "SELECT TributacaoComplNF.NaturezaOp, TributacaoComplNF.Tipo, TributacaoComplNF.ICMSValor, TributacaoComplNF.ICMSAliquota, TributacaoComplNF.ICMSBase, TributacaoComplNF.ICMSSubstBase, TributacaoComplNF.IPIValor, TiposTribICMS.ColunaNoLivroEntrada, TiposTribIPI.ColunaNoLivro FROM TributacaoComplNF, TiposTribICMS, TiposTribIPI WHERE TributacaoComplNF.NumIntDoc = ? AND TributacaoComplNF.ICMSTipo = TiposTribICMS.Tipo AND TributacaoComplNF.IPITipo = TiposTribIPI.Tipo AND TributacaoComplNF.ICMSBase >0 ORDER BY TributacaoComplNF.NumIntDoc ASC", tTipoGenericoFornecedores.sNaturezaOp, tTipoGenericoFornecedores.iTipoComplemento, tTipoGenericoFornecedores.dValorICMS, tTipoGenericoFornecedores.dICMSAliquota, tTipoGenericoFornecedores.dBaseICMS, tTipoGenericoFornecedores.dICMSSubstBase, tTipoGenericoFornecedores.dValorIPI, tTipoGenericoFornecedores.iColunaNoLivroICMS, tTipoGenericoFornecedores.iColunaNoLivroIPI, tTipoGenericoFornecedores.lNumIntDocNF)
''        If lErro <> AD_SQL_SUCESSO Then Error 58810
''
''        lErro = Comando_BuscarPrimeiro(alComando(2))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58811
''
''        Do While lErro = SUCESSO
''
''            'Transfere os dados do tTipoGenericoFornecedores para tTipo54
''            Call Organiza_TypeTipo54(tTipo54, tTipoGenericoFornecedores, tTipoGenericoFornecedores.iTipoComplemento)
''
''            'Acumula os Valores para a Geracao do Tipo 50 e 51
''            Call Acumula_Valores(colInfoAux, tTipoGenericoFornecedores, tTipoGenericoFornecedores.iTipoComplemento)
''
''            'Gera o Tipo 54
''            Call MontarReg54(sBuffer, sChaveOrd, tTipo54)
''
''            ' Gera o NumInt do Registro
''            lNumIntReg = lNumIntReg + 1
''
''            'Conta o
''            mvarlNumReg54 = mvarlNumReg54 + 1
''
''            'OK ??? sugiro manter comando(s) aberto(s) p/chamar funcao abaixo
''
''            'e Grava no BD o Tipo 54
''            lErro = GeracaoArqICMS_Grava(sBuffer, sChaveOrd, 54, lNumIntReg, alComando(4))
''            If lErro <> SUCESSO Then Error 58812
''
''            lErro = Comando_BuscarProximo(alComando(2))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58813
''
''        Loop
''
''        'gera o tipo 50 e 51
''        'Para cada objGeracaoArqICMSAux da colecao gera  um tipo 50 e 51
''        For Each objGeracaoArqICMSAux In colInfoAux
''
''            'Passa os Valores de tTipoGenericoFornecedores, objGeracaoArqICMSAux para tTipo50
''            Call Organiza_TypeTipo50(tTipo50, tTipoGenericoFornecedores, objGeracaoArqICMSAux)
''
''            'Monta o Buffer
''            Call MontarReg50(sBuffer, sChaveOrd, tTipo50)
''
''            'Gera o NumInt do Registro
''            lNumIntReg = lNumIntReg + 1
''
''            'Conta o Tipo 50 para a Geracao do Tipo 90
''            mvarlNumReg50 = mvarlNumReg50 + 1
''
''            'OK ??? sugiro manter comando(s) aberto(s) p/chamar funcao abaixo
''
''            'e Grava no BD o Tipo 50
''            lErro = GeracaoArqICMS_Grava(sBuffer, sChaveOrd, 50, lNumIntReg, alComando(4))
''            If lErro <> SUCESSO Then Error 58814
''
''            'Passa os Valores de tTipoGenericoFornecedores, objGeracaoArqICMSAux para tTipo51
''            Call Organiza_TypeTipo51(tTipo51, tTipoGenericoFornecedores, objGeracaoArqICMSAux)
''
''            'Monta o Buffer
''            Call MontarReg51(sBuffer, sChaveOrd, tTipo51)
''
''            'Gera o NumInt do Registro
''            lNumIntReg = lNumIntReg + 1
''
''            'Conta o Tipo 51 para a Geracao do Tipo 90
''            mvarlNumReg51 = mvarlNumReg51 + 1
''
''            'OK ??? sugiro manter comando(s) aberto(s) p/chamar funcao abaixo
''
''            'e Grava no BD o Tipo 51
''            lErro = GeracaoArqICMS_Grava(sBuffer, sChaveOrd, 51, lNumIntReg, alComando(4))
''            If lErro <> SUCESSO Then Error 58815
''
''        Next
''
''        Set colInfoAux = New Collection
''
''        'Busca a Proxina Nota
''        lErro = Comando_BuscarProximo(alComando(1))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58816
''
''    Loop
''
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    ProcessarNFs_Fornecedor = SUCESSO
''
''    Exit Function
''
''Erro_ProcessarNFs_Fornecedor:
''
''    ProcessarNFs_Fornecedor = Err
''
''    Select Case Err
''
''        Case 58802
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 58803, 58804, 58816
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)
''
''        Case 58805, 58806, 58809
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", Err)
''
''        Case 58807, 58808, 58812, 58814, 58815
''
''        Case 58810, 58811, 58813
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBUTACAOCOMPLNF", Err)
''
''        Case 61492 'Cancelou Batch
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150536)
''
''    End Select
''
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    Exit Function
''
''End Function
''
''Private Function InserirRegsArqICMS(objFilial As AdmFiliais) As Long
''
''Dim sBuffer As String
''Dim iIndice As Integer
''Dim lErro As Long
''
''On Error GoTo Erro_InserirRegsArqICMS
''
''    Open mvarsNomeArq For Output As #1
''
''    lErro = MontarReg10(sBuffer, objFilial)
''    If lErro <> SUCESSO Then Error 58819
''
''    'incluir reg tipo 10
''    Print #1, sBuffer
''
''    lErro = MontarReg11(sBuffer, objFilial)
''    If lErro <> SUCESSO Then Error 58820
''
''    'incluir reg tipo 11
''    Print #1, sBuffer
''
''    'percorrer a tabela GeracaoArqICMS onde ficaram armazenados os regs montados acima e criar o arquivo
''    lErro = GeracaoArqICMS_Inclui_registros()
''    If lErro <> SUCESSO Then Error 58821
''
''    'Monta e insere os Registro Tipo 75
''    lErro = GeracaoArqICMSProd_Inclui_Tipos75()
''    If lErro <> SUCESSO Then Error 58822
''
''    'Inclui o Reg 90
''    lErro = MontarReg90(sBuffer, objFilial.sCgc, objFilial.sInscricaoEstadual)
''    If lErro <> SUCESSO Then Error 58823
''
''    Print #1, sBuffer
''
''    Close
''
''    InserirRegsArqICMS = SUCESSO
''
''    Exit Function
''
''Erro_InserirRegsArqICMS:
''
''    InserirRegsArqICMS = Err
''
''    Select Case Err
''
''        Case 58819, 58820, 58821, 58822, 58823
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150537)
''
''    End Select
''
''    Close
''
''    Exit Function
''
''End Function
''
'''FORMATACAO DOS CAMPOS
''Private Function FormataCpoData(dtData As Date) As String
'''formata campo data como AAAAMMDD
''
''    FormataCpoData = Format(dtData, "yyyymmdd")
''
''End Function
''
''Private Function FormataCpoQuantidade(vData As Variant, iTam As Integer) As String
'''formata campo Quantidade alinhado-o à direita sem ponto e decimais, com zeros a esquerda
''
''Dim sValor As String
''Dim iTamanhoValor As Integer
''Dim dValor As Double
''Dim sDecimais As String
''Dim sInteiro As String
''Dim sNumero As String
''Dim sEspacos As String
''Dim iData As Integer
''
''    sValor = Format(vData, "#,##0.000")
''    sDecimais = Mid(sValor, Len(sValor) - 2, 3)
''    sInteiro = Mid(sValor, 1, Len(sValor) - 4)
''    sNumero = sInteiro & sDecimais
''
''    iData = iTam - Len(sNumero)
''    sEspacos = String(iData, "0")
''
''    FormataCpoQuantidade = sEspacos & sInteiro & sDecimais
''
''End Function
''
''Private Sub Acumula_Valores(colInfoAux As Collection, tTipoGenerico As typeTipo, iColuna As Integer)
''
''Dim objGeracaoArqICMSAux As New ClassGeracaoArqICMSAux
''Dim bEncontrou As Boolean
''Dim iIndice As Integer
''
''    bEncontrou = False
''
''    'Procura para ver se já tem Nat. de Operacao na coleção
''    For Each objGeracaoArqICMSAux In colInfoAux
''
''        iIndice = iIndice + 1
''
''        If objGeracaoArqICMSAux.sNatOperacao = tTipoGenerico.sNaturezaOp And objGeracaoArqICMSAux.dAliquota = tTipoGenerico.dICMSAliquota Then
''
''            objGeracaoArqICMSAux.dAliquota = tTipoGenerico.dICMSAliquota
''            objGeracaoArqICMSAux.dBaseICMS = objGeracaoArqICMSAux.dBaseICMS + tTipoGenerico.dBaseICMS
''
''            Select Case tTipoGenerico.iColunaNoLivroICMS
''
''                Case TIPO_TRIBUTADO
''                    objGeracaoArqICMSAux.dValorICMS = objGeracaoArqICMSAux.dValorICMS + tTipoGenerico.dValorICMS
''                Case TIPO_ISENTO_NAO_TRIBUTADO
''                    objGeracaoArqICMSAux.dIsentoNaoTribICMS = objGeracaoArqICMSAux.dIsentoNaoTribICMS + tTipoGenerico.dValorICMS
''                Case TIPO_OUTRAS
''                    objGeracaoArqICMSAux.dOutrasICMS = objGeracaoArqICMSAux.dOutrasICMS + tTipoGenerico.dValorICMS
''
''            End Select
''
''            Select Case tTipoGenerico.iColunaNoLivroIPI
''
''                Case TIPO_TRIBUTADO
''                    objGeracaoArqICMSAux.dValorIPI = objGeracaoArqICMSAux.dValorIPI + tTipoGenerico.dValorIPI
''                Case TIPO_ISENTO_NAO_TRIBUTADO
''                    objGeracaoArqICMSAux.dIsentoNaoTribIPI = objGeracaoArqICMSAux.dIsentoNaoTribIPI + tTipoGenerico.dValorIPI
''                Case TIPO_OUTRAS
''                    objGeracaoArqICMSAux.dOutrasIPI = objGeracaoArqICMSAux.dOutrasIPI + tTipoGenerico.dValorIPI
''
''            End Select
''
''            objGeracaoArqICMSAux.dValorTotal = objGeracaoArqICMSAux.dValorTotal + (tTipoGenerico.dQuantidade * tTipoGenerico.dPrecoUnitario)
''
''            Select Case iColuna
''
''                Case TIPO_COMPL_FRETE
''                    objGeracaoArqICMSAux.dValorTotal = objGeracaoArqICMSAux.dValorTotal + tTipoGenerico.dValorFrete
''                Case TIPO_COMPL_SEGURO
''                    objGeracaoArqICMSAux.dValorTotal = objGeracaoArqICMSAux.dValorTotal + tTipoGenerico.dValorSeguro
''                Case TIPO_COMPL_DESCONTO
''                    objGeracaoArqICMSAux.dValorTotal = objGeracaoArqICMSAux.dValorTotal - tTipoGenerico.dValorDescontoTotal
''                Case TIPO_COMPL_OUTRAS
''                    objGeracaoArqICMSAux.dValorTotal = objGeracaoArqICMSAux.dValorTotal + tTipoGenerico.dValorOutras
''
''            End Select
''
''            bEncontrou = True
''            Exit For
''
''        End If
''
''    Next
''
''    If bEncontrou = False Then
''
''        Set objGeracaoArqICMSAux = New ClassGeracaoArqICMSAux
''
''        objGeracaoArqICMSAux.dAliquota = tTipoGenerico.dICMSAliquota
''        objGeracaoArqICMSAux.dBaseICMS = tTipoGenerico.dBaseICMS
''
''        Select Case tTipoGenerico.iColunaNoLivroICMS
''
''            Case TIPO_TRIBUTADO
''                objGeracaoArqICMSAux.dValorICMS = tTipoGenerico.dValorICMS
''            Case TIPO_ISENTO_NAO_TRIBUTADO
''                objGeracaoArqICMSAux.dIsentoNaoTribICMS = tTipoGenerico.dValorICMS
''            Case TIPO_OUTRAS
''                objGeracaoArqICMSAux.dOutrasICMS = tTipoGenerico.dValorICMS
''
''        End Select
''
''        Select Case tTipoGenerico.iColunaNoLivroIPI
''
''            Case TIPO_TRIBUTADO
''                objGeracaoArqICMSAux.dValorIPI = tTipoGenerico.dValorIPI
''            Case TIPO_ISENTO_NAO_TRIBUTADO
''                objGeracaoArqICMSAux.dIsentoNaoTribIPI = tTipoGenerico.dValorIPI
''            Case TIPO_OUTRAS
''                objGeracaoArqICMSAux.dOutrasIPI = tTipoGenerico.dValorIPI
''
''        End Select
''
''        objGeracaoArqICMSAux.dValorTotal = (tTipoGenerico.dQuantidade * tTipoGenerico.dPrecoUnitario)
''        objGeracaoArqICMSAux.sNatOperacao = tTipoGenerico.sNaturezaOp
''
''        colInfoAux.Add objGeracaoArqICMSAux
''
''    End If
''
''End Sub
''
''Private Sub Organiza_TypeTipo54(tTipo54 As typeTipo54, tTipoGenerico As typeTipo, iTipoTributacao As Integer)
'''Transfere os dados do tTipoGenerico para tTipo54 para a montagem do arquivo
''
''    tTipo54.dAliquotaICMS = tTipoGenerico.dICMSAliquota
''    tTipo54.dBaseICMS = tTipoGenerico.dBaseICMS
''    tTipo54.dBaseICMSSubst = tTipoGenerico.dICMSSubstBase
''    tTipo54.dValorIPI = tTipoGenerico.dValorIPI
''    tTipo54.iModelo = tTipoGenerico.iModelo
''
''    Select Case iTipoTributacao
''
''        Case TIPO_COMPL_NORMAL
''            tTipo54.iItem = tTipoGenerico.iItem
''            tTipo54.dQuantidade = tTipoGenerico.dQuantidade
''            tTipo54.dValorProduto = tTipoGenerico.dQuantidade * tTipoGenerico.dPrecoUnitario
''            tTipo54.sCodigoProduto = tTipoGenerico.sCodigoProduto
''            If tTipoGenerico.dValorTotal > 0 Then
''                tTipo54.dValorDescontos = tTipoGenerico.dValorDescontoItem + ((tTipoGenerico.dQuantidade * tTipoGenerico.dPrecoUnitario * tTipoGenerico.dValorDescontoTotal) / tTipoGenerico.dValorTotal)
''            Else
''                tTipo54.dValorDescontos = tTipoGenerico.dValorDescontoItem
''            End If
''
''        Case TIPO_COMPL_FRETE
''            tTipo54.iItem = 991
''            tTipo54.dQuantidade = 0
''            tTipo54.dValorProduto = 0
''            tTipo54.dValorDescontos = tTipoGenerico.dValorFrete
''            tTipo54.sCodigoProduto = ""
''
''        Case TIPO_COMPL_SEGURO
''            tTipo54.iItem = 992
''            tTipo54.dQuantidade = 0
''            tTipo54.dValorProduto = 0
''            tTipo54.dValorDescontos = tTipoGenerico.dValorSeguro
''            tTipo54.sCodigoProduto = ""
''
''        Case TIPO_COMPL_OUTRAS
''            tTipo54.iItem = 999
''            tTipo54.dQuantidade = 0
''            tTipo54.dValorProduto = 0
''            tTipo54.dValorDescontos = tTipoGenerico.dValorOutras
''            tTipo54.sCodigoProduto = ""
''
''    End Select
''
''    tTipo54.lNumeroNF = tTipoGenerico.lNumeroNF
''    tTipo54.sCFOP = tTipoGenerico.sNaturezaOp
''    tTipo54.sCgc = tTipoGenerico.sCgc
''    tTipo54.sSerie = tTipoGenerico.sSerie
''
''End Sub
''
''Private Sub Organiza_TypeTipo50(tTipo50 As typeTipo50, tTipoGenerico As typeTipo, objGeracaoArqICMSAux As ClassGeracaoArqICMSAux)
'''Transfere os dados do tTipoGenerico e os totais de objGeracaoArqICMSAux para tTipo50 para a montagem do arquivo
''
''    tTipo50.dAliquotaICMS = objGeracaoArqICMSAux.dAliquota
''    tTipo50.dBaseICMS = objGeracaoArqICMSAux.dBaseICMS
''    tTipo50.dIsentaNTributada = objGeracaoArqICMSAux.dIsentoNaoTribICMS
''    tTipo50.dOutras = objGeracaoArqICMSAux.dOutrasICMS
''    tTipo50.dtDataEmissao = tTipoGenerico.dtDataEmissao
''    tTipo50.dValorICMS = objGeracaoArqICMSAux.dValorICMS
''    tTipo50.dValorTotal = objGeracaoArqICMSAux.dValorTotal
''    tTipo50.iModelo = tTipoGenerico.iModelo
''    tTipo50.lNumeroNF = tTipoGenerico.lNumeroNF
''    tTipo50.sCFOP = objGeracaoArqICMSAux.sNatOperacao
''    tTipo50.sSerie = tTipoGenerico.sSerie
''
''    If tTipoGenerico.iStatus = STATUS_CANCELADO Then
''        tTipo50.sSituacao = "S" 'OK ??? agora existe status de cancelado
''    Else
''        tTipo50.sSituacao = "N"
''    End If
''
''    If Len(tTipoGenerico.sInscricaoEstadual) = 0 Or tTipoGenerico.iCodigoPais <> 1 Then
''        tTipo50.sInscricaoEstadual = "ISENTO"
''    Else
''        tTipo50.sInscricaoEstadual = tTipoGenerico.sInscricaoEstadual
''    End If
''
''    If Len(tTipoGenerico.sCgc) = 0 Or tTipoGenerico.iCodigoPais <> 1 Then
''        tTipo50.sCgc = String(14, "0")
''    Else
''        'Se for NF saída usar Destinatário
''        If tTipoGenerico.iTipo = TIPODOCINFO_TIPO_NFIS Then
''            'Se for a Empresa, usar o CGC da Empresa
''            If tTipoGenerico.iDestinatario = EMITENTE_EMPRESA Then
''                tTipo50.sCgc = mvarsCGCEmpresa
''            'Se for Cliente ou Fornecedor
''            Else
''                tTipo50.sCgc = tTipoGenerico.sCgc
''            End If
''
''        'Se for NF Entrada usar Remetente ou Emitente
''        ElseIf tTipoGenerico.iTipo = TIPODOCINFO_TIPO_NFIE Or tTipoGenerico.iTipo = TIPODOCINFO_TIPO_NFEXT Then
''            'Se for a Empresa, usar o CGC da Empresa
''            If tTipoGenerico.iEmitente = EMITENTE_EMPRESA Then
''                tTipo50.sCgc = mvarsCGCEmpresa
''            'Se for Cliente ou Fornecedor
''            Else
''                tTipo50.sCgc = tTipoGenerico.sCgc
''            End If
''        End If
''    End If
''
''    'Caso seja exterior
''    If tTipoGenerico.iCodigoPais <> 1 Then
''        tTipo50.sUnidadeFederacao = "EX"
''    Else
''        tTipo50.sUnidadeFederacao = tTipoGenerico.sUnidadeFederacao
''    End If
''
''End Sub
''
''Private Sub Organiza_TypeTipo51(tTipo51 As typeTipo51, tTipoGenerico As typeTipo, objGeracaoArqICMSAux As ClassGeracaoArqICMSAux)
'''Transfere os dados do tTipoGenerico e os totais de objGeracaoArqICMSAux para tTipo50 para a montagem do arquivo
''
''    tTipo51.dIsentaNTributada = objGeracaoArqICMSAux.dIsentoNaoTribICMS
''    tTipo51.dOutras = objGeracaoArqICMSAux.dOutrasICMS
''    tTipo51.dtDataEmissao = tTipoGenerico.dtDataEmissao
''    tTipo51.dValorIPI = objGeracaoArqICMSAux.dValorIPI
''    tTipo51.dValorTotal = objGeracaoArqICMSAux.dValorTotal
''    tTipo51.lNumeroNF = tTipoGenerico.lNumeroNF
''    tTipo51.sCFOP = objGeracaoArqICMSAux.sNatOperacao
''    tTipo51.sSerie = tTipoGenerico.sSerie
''    tTipo51.sSituacao = "N"
''
''    If Len(tTipoGenerico.sInscricaoEstadual) = 0 Or tTipoGenerico.iCodigoPais <> 1 Then
''        tTipo51.sInscricaoEstadual = "ISENTO"
''    Else
''        tTipo51.sInscricaoEstadual = tTipoGenerico.sInscricaoEstadual
''    End If
''
''    If Len(tTipoGenerico.sCgc) = 0 Or tTipoGenerico.iCodigoPais <> 1 Then
''        tTipo51.sCgc = String(14, "0")
''    Else
''        tTipo51.sCgc = tTipoGenerico.sCgc
''    End If
''
''    'Caso seja exterior
''    If tTipoGenerico.iCodigoPais <> 1 Then
''        tTipo51.sUnidadeFederacao = "EX"
''    Else
''        tTipo51.sUnidadeFederacao = tTipoGenerico.sUnidadeFederacao
''    End If
''
''End Sub
''
''Private Sub Limpa_Tipo(tTipoGenerico As typeTipo)
''
''    tTipoGenerico.dBaseICMS = 0
''    tTipoGenerico.dICMSAliquota = 0
''    tTipoGenerico.dICMSSubstBase = 0
''    tTipoGenerico.dPrecoUnitario = 0
''    tTipoGenerico.dQuantidade = 0
''    tTipoGenerico.dValorDescontoItem = 0
''
''End Sub
''
''Function GeracaoArqICMS_Inclui_registros() As Long
'''Le todos os Tipos armazenados no BD no, em Ordem
''
''Dim lErro As Long
''Dim sBuffer As String
''Dim lComando As Long
''
''On Error GoTo Erro_GeracaoArqICMS_Inclui_registros
''
''    'Abre o Comando
''    lComando = Comando_Abrir()
''    If lComando = 0 Then Error 58824
''
''    sBuffer = String(127, 0)
''
''    'Seleciona todos os Tipos
''    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM GeracaoArqICMS ORDER BY Tipo, ChaveOrd ASC", sBuffer)
''    If lErro <> SUCESSO Then Error 58825
''
''    lErro = Comando_BuscarPrimeiro(lComando)
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58826
''
''    Do While lErro = AD_SQL_SUCESSO
''
''        Print #1, sBuffer
''
''        lErro = Comando_BuscarProximo(lComando)
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58827
''
''    Loop
''
''    Call Comando_Fechar(lComando)
''
''    GeracaoArqICMS_Inclui_registros = SUCESSO
''
''    Exit Function
''
''Erro_GeracaoArqICMS_Inclui_registros:
''
''    GeracaoArqICMS_Inclui_registros = Err
''
''    Select Case Err
''
''        Case 58824
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 58825, 58826, 58827
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_GERACAOARQICMS", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150538)
''
''    End Select
''
''    Call Comando_Fechar(lComando)
''
''    Exit Function
''
''End Function
''
''Function GeracaoArqICMS_Grava(sBuffer As String, sChaveOrd As String, iTipo As Integer, lNumIntReg As Long, lComando As Long) As Long
'''Grava o Tipo no BD
'''Tem que ser chamado  em transacao
''
''Dim lErro As Long
''
''On Error GoTo Erro_GeracaoArqICMS_Grava
''
''    'Insere no Bd o Tipo passado
''    lErro = Comando_Executar(lComando, "INSERT INTO GeracaoArqICMS (NumIntReg, Tipo, ChaveOrd, Conteudo) VALUES (?,?,?,?)", lNumIntReg, iTipo, sChaveOrd, sBuffer)
''    If lErro <> AD_SQL_SUCESSO Then Error 58829
''
''    GeracaoArqICMS_Grava = SUCESSO
''
''    Exit Function
''
''Erro_GeracaoArqICMS_Grava:
''
''    GeracaoArqICMS_Grava = Err
''
''    Select Case Err
''
''        Case 58828
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 58829
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TABELA_GERACAOARQICMS", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150539)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''
''Function GeracaoArqICMSProd_Inclui_Tipos75() As Long
'''Le todos os Tipos 75 armazenados no BD no, em Ordem e inclui no Arquivo
''
''Dim lErro As Long
''Dim sBuffer As String
''Dim lComando As Long
''Dim tTipo75 As typeTipo75
''Dim iOrigemMercadoria As Integer
''Dim iTipoTribCST As Integer
''
''On Error GoTo Erro_GeracaoArqICMSProd_Inclui_Tipos75
''
''    'Abre o Comando
''    lComando = Comando_Abrir()
''    If lComando = 0 Then Error 58835
''
''    tTipo75.sCodigoProduto = String(STRING_PRODUTO, 0)
''    tTipo75.sDescricaoProduto = String(STRING_PRODUTO_DESCRICAO, 0)
''    tTipo75.sUnidadeMedida = String(STRING_UM_SIGLA, 0)
''    tTipo75.sCodigoNCM = String(STRING_PRODUTO_IPI_CODIGO, 0)
''
''    'Seleciona todos os Tipos
''    lErro = Comando_Executar(lComando, "SELECT GeracaoArqICMSProd.Codigo, GeracaoArqICMSProd.Descricao, GeracaoArqICMSProd.UnidadeMedida, Produtos.OrigemMercadoria, TiposTribICMS.TipoTribCST, Produtos.IPICodigo FROM GeracaoArqICMSProd, Produtos, TiposTribICMS WHERE Produtos.Codigo = GeracaoArqICMSProd.Codigo AND TiposTribICMS.Tipo = GeracaoArqICMSProd.TipoTribICMS ORDER BY GeracaoArqICMSProd.Codigo ASC", tTipo75.sCodigoProduto, tTipo75.sDescricaoProduto, tTipo75.sUnidadeMedida, iOrigemMercadoria, iTipoTribCST, tTipo75.sCodigoNCM)
''    If lErro <> SUCESSO Then Error 58836
''
''    lErro = Comando_BuscarPrimeiro(lComando)
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58837
''
''    Do While lErro = AD_SQL_SUCESSO
''
''        'SituacaoTributaria
''        tTipo75.sSituacaoTributaria = CStr(iOrigemMercadoria & iTipoTribCST)
''
''        'Conta o Tipo 75 para a Geracao do Tipo 90
''        mvarlNumReg75 = mvarlNumReg75 + 1
''
''        'JONES - Falta passar alguns valores(AliquotaICMS, Reducao de Base ICMS, Base de Calculo do ICMS de Substituicao Tributaria estes valores só entram quando o produto tem este valores fixos)
''        Call MontarReg75(sBuffer, tTipo75)
''
''        Print #1, sBuffer
''
''        lErro = Comando_BuscarProximo(lComando)
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58838
''
''    Loop
''
''    Call Comando_Fechar(lComando)
''
''    GeracaoArqICMSProd_Inclui_Tipos75 = SUCESSO
''
''    Exit Function
''
''Erro_GeracaoArqICMSProd_Inclui_Tipos75:
''
''    GeracaoArqICMSProd_Inclui_Tipos75 = Err
''
''    Select Case Err
''
''        Case 58835
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 58836, 58837, 58838
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_GERACAOARQICMSPROD", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150540)
''
''    End Select
''
''    Call Comando_Fechar(lComando)
''
''    Exit Function
''
''End Function
''
