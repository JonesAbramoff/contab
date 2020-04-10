Attribute VB_Name = "ErrosFIS"
Option Explicit

'C�digos de Erro - Reservado de 13300 a 13499
Public Const ERRO_DATAINICIO_MAIOR_DATAFIM = 13300 'Sem par�metros
'A data inicial n�o pode ser maior que a data final.
Public Const ERRO_TRIBUTO_NAO_PREENCHIDO = 13301 'Sem par�metros
'� obrigat�rio o preenchimento de Tributo.
Public Const ERRO_LIVRO_NAO_PREENCHIDO = 13302 'Sem par�metros
'� obrigat�rio o preenchimento do Livro Fiscal.
Public Const ERRO_PERIODICIDADE_NAO_PREENCHIDO = 13303 'Sem par�metros
'� obrigat�rio o preenchimento da Periodicidade.
Public Const ERRO_NUMEROLIVRO_NAO_PREENCHIDO = 13304 'Sem par�metros
'� obrigat�rio o preenchimento do n�mero do Livro.
Public Const ERRO_FOLHA_NAO_PREENCHIDA = 13305 'Sem par�metros
'� obrigat�rio o preenchimento do n�mero do Folha.
Public Const ERRO_LEITURA_TRIBUTOS = 13306 'Sem par�metros
'Erro na leitura da tabela Tributos.
Public Const ERRO_LEITURA_PERIODICIDADESLIVROSFISC = 13307 'Sem par�metros
'Erro na leitura da tabela PeriodiciddadesLivrosFisc.
Public Const ERRO_LEITURA_LIVROFISCAL = 13308 'Sem par�metros
'Erro na leitura da tabela LivrosFiscais.
Public Const ERRO_DATAINICIAL_MAIOR_DATAFINALLIVRO = 13309 'Par�metros: dtDataFinal, dtDataInicial
'A Data Inicial %s do Livro Fiscal tem que ser maior do que a data
'final do Livro que j� Fechado, que � %s.
Public Const ERRO_LOCK_LIVROSFILIAL = 13310 'Par�metros: iCodLivro, iFilialEmpresa
'Erro na tentativa de fazer "Lock" em LivrosFilial com Livro de c�digo %s e Filial Empresa de c�digo %s.
Public Const ERRO_EXCLUSAO_LIVROSFILIAL = 13311 'Par�metros: iCodLivro, iFilialEmpresa
'Erro na exclus�o de LivrosFilial com Livro de c�digo %s e Filial Empresa de c�digo %s.
Public Const ERRO_ATUALIZACAO_LIVROSFILIAL = 13312 'Par�metros: iCodLivro, iFilialEmpresa
'Erro na atualiza��o de LivrosFilial com Livro de c�digo %s e Filial Empresa de c�digo %s.
Public Const ERRO_INSERCAO_LIVROSFILIAL = 13313 'Par�metros: iCodLivro, iFilialEmpresa
'Erro na inser��o de LivrosFilial com Livro de c�digo %s e Filial Empresa de c�digo %s.
Public Const ERRO_LIVROFISCAL_NAO_CADASTRADO = 13314 'Par�metros: iCodigo
'O Livro Fiscal de c�digo %s n�o est� cadastrado no Banco de dados.
Public Const ERRO_LIVROFILIAL_VINCULADO_REGAPURACAOICMS = 13315 'Par�metros: iCodLivro, iFilialEmpresa, dtDataInicial, dtDataFinal
'O Livro Fiscal de c�digo %s da Filial Empresa de c�digo %s est� vinculado a um Registro de Apura��o
'de ICMS de Data Inicial %s e Data Final %s.
Public Const ERRO_LIVROFILIAL_VINCULADO_REGAPURACAOIPI = 13316 'Par�metros: iCodLivro, iFilialEmpresa, dtDataInicial, dtDataFinal
'O Livro Fiscal de c�digo %s da Filial Empresa de c�digo %s est� vinculado a um Registro de Apura��o
'de IPI de Data Inicial %s e Data Final %s.
Public Const ERRO_LEITURA_REGAPURACAOIPI = 13317 'Sem par�metros
'Erro na leitura da tabela RegApuracaoIPI.
Public Const ERRO_LEITURA_LIVROSFECHADOS = 13318 'Sem par�metros
'Erro na leitura da tabela LivrosFechados.
Public Const ERRO_SECAO_NAO_PREENCHIDA = 13319 'Sem par�metros
'� obrigat�rio o preenchimento da Se��o.
Public Const ERRO_LEITURA_TIPOREGAPURACAOICMS = 13320 'Sem par�metros
'Erro na leitura da tabela TipoRegApuracaoICMS.
Public Const ERRO_TIPOREGAPURACAOICMS_NAO_CADASTRADA = 13321 'Par�metros: iCodigo
'O tipo de registro para apura��o ICMS de c�digo %s n�o est� cadastrado.
Public Const ERRO_LOCK_TIPOREGAPURACAOICMS = 13322 'Par�metros: iCodigo
'Erro na tentativa de fazer "lock" no Tipo de Registro de apura��o ICMS de c�digo %s.
Public Const ERRO_TIPOREGAPURACAOICMS_PRE_CADASTRADO = 13323 'Par�metros: iCodigo
'O Tipo  de Registro de Apura��o de ICMS de c�digo %s � pr�-cadastrado
'e n�o pode ser excluido nem alterado.
Public Const ERRO_ATUALIZACAO_TIPOREGAPURACAOICMS = 13324 'Par�metros: iCodigo
'Erro na atualiza��o do Tipo de Registro de apura��o ICMS de c�digo %s.
Public Const ERRO_INSERCAO_TIPOREGAPURACAOICMS = 13325 'Par�metros: iCodigo
'Erro na inser��o do Tipo de Registro de apura��o ICMS de c�digo %s.
Public Const ERRO_EXCLUSAO_TIPOREGAPURACAOICMS = 13326 'Par�metros: iCodigo
'Erro na exclus�o do Tipo de Registro de apura��o ICMS de c�digo %s.
Public Const ERRO_LEITURA_REGAPURACAOICMSITEM = 13327 'Sem par�metros
'Erro na leitura da tabela RegApuracaoICMSItem.
Public Const ERRO_TIPOREGAPURACAO_VINCULADO_REGAPURACAOICMSITEM = 13328 'Par�metros: iCodigo, lNumIntDoc
'O Tipo de Registro de Apura��o de ICMS de c�digo %s n�o pode ser excluido pois est� vinculado
'ao Item de Registro de Apura��o ICMS.
Public Const ERRO_ATUALIZACAO_REGAPURACAOICMSITEM1 = 13329 'Sem Par�metros
'Erro na atualiza��o do Item de Apura��o de ICMS.
Public Const ERRO_ATUALIZACAO_REGAPURACAOIPI = 13330 'Par�metros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na atualiza��o do Registro de apura��o de IPI de per�odo %s at� o per�odo %s da Filial Empresa de c�digo %s.
Public Const ERRO_ATUALIZACAO_REGAPURACAOIPIITEM1 = 13331 'Sem Par�metros
'Erro na atualiza��o do Item de Apura��o de IPI.
Public Const ERRO_LEITURA_REGAPURACAOIPIITEM = 13332 'Sem par�metros
'Erro na leitura da tabela RegApuracaoIPIItem.
Public Const ERRO_NFISCAL_NAO_SELECIONADA = 13333 'Sem Parametros
'� necess�rio selecionar uma Nota Fiscal.
Public Const ERRO_GRIDLANCAMENTO_VAZIO = 13334 'Sem Parametros
'� nescess�rio ter pelo menos uma linha no Grid de Lan�amentos.
Public Const ERRO_FISCAL_GRID_NAO_PREENCHIDO = 13335 'Parametro: iLinha
'O Campo Fiscal da Linha %s n�o foi preenchido.
Public Const ERRO_UF_GRID_NAO_PREENCHIDO = 13336 'Parametro: iLinha
'O Campo UF da Linha %s n�o foi preenchido.
Public Const ERRO_VALORCONTABIL_GRID_NAO_PREENCHIDO = 13337 'Parametro: iLinha
'O Campo Valor Cont�bil da Linha %s n�o foi preenchido.
Public Const ERRO_LEITURA_LIVREGES = 13338 'Sem Parametros
'Erro na Leitura da Tabela de LivRegES.
Public Const ERRO_LEITURA_LIVREGESITEM = 13339 'Sem Parametros
'Erro na Leitura da Tabela de LivRegESItem.
Public Const ERRO_LEITURA_LIVREGCADPROD = 13340 'Sem Parametros
'Erro na Leitura da Tabela de LivRegESCadProd.
Public Const ERRO_LIVREGCADPROD_NAO_ENCONTRADO = 13341 'Sem Parametros
'N�o foi encontrado o registro em LivRegESCadProd para o Item de Registro de Entrada ou Sa�da.
Public Const ERRO_LEITURA_LIVREGESLINHA = 13342 'Sem Parametros
'Erro na Leitura da Tabela de LivRegESLinha.
Public Const ERRO_LOCK_LIVREGESLINHA = 13343
'Ocorreu um Erro na tentativa de fazer "lock" na Tabela de LivRegESLinha.
Public Const ERRO_EXCLUSAO_LIVREGESLINHA = 13344 'Sem Parametros
'Erro na tentativa de excluir LivRegESLinha.
Public Const ERRO_INSERCAO_LIVREGESLINHA = 13345
'Erro na tentiva de inserir um registro na tabela LivRegESLinha.
Public Const ERRO_LOCK_LIVREGESITEM = 13346 'Sem Parametros
'Ocorreu um Erro na tentativa de fazer "lock" na Tabela de LivRegESItem.
Public Const ERRO_ATUALIZACAO_LIVREGESITEM = 13347 'Sem Parametros
'Erro na tentativa de Atualizar LivRegESItem.
Public Const ERRO_EXCLUSAO_LIVREGESITEM = 13348 'Sem Parametros
'Erro na tentativa de excluir LivRegESItem.
Public Const ERRO_LIVROENTRADA_NAO_CADASTRADO_NFISCAL = 13349  'Parametros = sSerie, lNumNotaFiscal
'Livro de Registro de Entrada n�o encontrado para a Nota Fiscal com S�rie = %s e N�mero %s.
Public Const ERRO_LOCK_LIVREGES = 13350 'Sem Parametros
'Ocorreu um Erro na tentativa de fazer "lock" na Tabela de LivRegES.
Public Const ERRO_ATUALIZACAO_LIVREGES = 13351 'Sem Parametros
'Erro na tentativa de Atualizar LivRegES.
Public Const ERRO_LIVROREGES_JA_FECHADO = 13352 'Sem Parametros
'O Livro Fiscal para a Nota Fiscal (Serie = %s e N�mero = %s) j� foi fechado, por isso n�o � poss�vel fazer altera��es.
Public Const ERRO_LEITURA_GERACAOARQICMS = 13354 'Sem Parametros
'Erro na Leitura da Tabela de GeracaoArqICMS.
Public Const ERRO_LEITURA_GERACAOARQICMSPROD = 13355 'Sem Parametros
'Erro na Leitura da Tabela de GeracaoArqICMSProd.
Public Const ERRO_EXCLUSAO_GERACAOARQICMS = 13356 'Sem Parametros
'Erro na Exclus�o na Tabela de GeracaoArqICMS.
Public Const ERRO_EXCLUSAO_GERACAOARQICMSPROD = 13357 'Sem Parametros
'Erro na Exclus�o na Tabela de GeracaoArqICMSProd.
Public Const ERRO_LEITURA_INFOARQICMS = 13358 'Sem Parametros
'Erro na Leitura da Tabela de InfoArqICMS.
Public Const ERRO_INFOARQICMS_JA_CADASTRADO = 13359 'Data Inicial, Data Final, Data Inicial Lido, Data Final Lido
'N�o � possivel gerar o Arquivo de ICMS para as Datas de %s at� %s, pois j� existe um arquivo para as datas de %s at� %s.
Public Const ERRO_ATUALIZACAO_INFOARQICMS = 13360 'Sem Parametros
'Erro na tentativa de Atualizar InfoArqICMS.
Public Const ERRO_DATA_PAGAMENTO_NAO_PREENCHIDO = 13361 'Sem par�metros
'O preenchimento de data de pagamento � obrigat�rio.
Public Const ERRO_TIPOGNR_NAO_PREENCHIDO = 13362 'Sem par�metros
'O preenchimento do Tipo de Guia � obrigat�rio.
Public Const ERRO_DATA_VENCIMENTO_NAO_PREENCHIDA = 13363 'Sem par�metros
'O preenchimento da Data de Vencimento � obrigat�ria.
Public Const ERRO_DATA_REFERENCIA_NAO_PREENCHIDA = 13364 'Sem par�metros
 'O preenchimento da Data de Refer�ncia � obrigat�ria.
Public Const ERRO_GNRICMS_VINCULADO_APURACAOICMS = 13365 'Parametro: lCodigo
'A Guia de Recolhimento %s j� est� associada a uma Apura��o ICMS.
Public Const ERRO_GNRICMS_VINCULADO_ARQICMS = 13366 'Parametro: lCodigo
'A Guia de Recolhimento %s j� est� associada a um Arquivo de ICMS.
Public Const ERRO_LEITURA_GNRICMS = 13367 'Sem par�metros
'Erro na leitura da tabela GNRICMS.
Public Const ERRO_GNRICMS_NAO_CADASTRADA = 13368 'Par�metros: lCodigo
'A Guia de ICMS de c�digo %s n�o est� cadastrada no banco de dados.
Public Const ERRO_EXCLUSAO_GNRICMS = 13369 'Par�metros: lCodigo
'Erro na exclus�o da Guia de ICMS de c�digo %s.
Public Const ERRO_ATUALIZACAO_GNRICMS = 13370 'Par�metros: lCodigo
'Erro na atualiza��o da Guia de ICMS de c�digo %s.
Public Const ERRO_INSERCAO_GNRICMS = 13371 'Par�metros: lCodigo
'Erro na inser��o da Guia de ICMS de c�digo %s.
Public Const ERRO_LOCK_GNRICMS = 13372 'Par�metros: lCodigo
'Erro na tentativa de fazer "lock" em GNRICMS com Guia ICMS de c�digo %s.
Public Const ERRO_NOME_EMPRESA_NAO_PREENCHIDO = 13373 'Sem par�metros
'O Nome da Empresa n�o foi preenchido.
Public Const ERRO_LEITURA_LIVROSFILIAL = 13374 'Sem par�metros
'Erro na leitura da tabela LivrosFilial.
Public Const ERRO_LIVROFILIAL_NAO_CADASTRADO = 13375 'Par�metros: iTipoLivro, iCodigoFilial
'N�o existe Livro Fiscal aberto com c�digo %s para Filial Empresa %s.
Public Const ERRO_LEITURA_REGAPURACAOICMS = 13376 'Sem par�metros
'Erro na leitura da tabela RegApuracaoICMS.
Public Const ERRO_NENHUMA_APURACAOICMS_CADASTRADA = 13377 'Sem par�metros
'N�o h� Apura��es de ICMS anteriores Fechadas cadastradas no Banco de dados.
Public Const ERRO_REGAPURACAOICMS_NAO_CADASTRADA = 13378 'Par�metros: dtDataInicial, dtDataFinal, iFilialEmpresa
'O Registro de apura��o de ICMS de per�odo %s at� o per�odo %s da Filial Empresa %s n�o est�
'cadastrado no Banco de dados.
Public Const ERRO_LOCK_REGAPURACAOICMS = 13379 'Par�metros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro no "Lock" do Registro de apura��o de ICMS de per�odo %s at� o per�odo %s da Filial Empresa de c�digo %s.
Public Const ERRO_EXCLUSAO_REGAPURACAOICMS = 13380 'Par�metros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na exclus�o do Registro de apura��o de ICMS de per�odo %s at� o per�odo %s da Filial Empresa de c�digo %s.
Public Const ERRO_ATUALIZACAO_REGAPURACAOICMS = 13381 'Par�metros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na atualiza��o do Registro de apura��o de ICMS de per�odo %s at� o per�odo %s da Filial Empresa de c�digo %s.
Public Const ERRO_INSERCAO_REGAPURACAOICMS = 13382 'Par�metros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na inser��o do Registro de apura��o de ICMS de per�odo %s at� o per�odo %s da Filial Empresa de c�digo %s.
Public Const ERRO_REGAPURACAOICMS_FECHADO = 13383 'Par�metros: DataInicial, DataFinal, FilialEmpresa
'O Livro Fiscal relacionado ao Registro de Apura��o de ICMS de data inicial %s e data final %s da
'Filial Empresa %s j� foi fechado.
Public Const ERRO_DATA_APURACAO_DIFERENTE_LIVROFILIAL = 13384 'Par�metros: iFilialEmpresa, dtDataInicial, dtDataFinal
'As datas Inicial %s e Final %s do Livro Fiscal relacionados a Filial Empresa de c�digo %s n�o s�o iguais ao
'da Apura��o em quest�o.
Public Const ERRO_TIPOAPURACAO_NAO_PREENCHIDA = 13385 'Sem par�metros
'O tipo de apura��o n�o foi preenchido.
Public Const ERRO_REGAPURACAOICMSITEM_FECHADO = 13386 'Par�metros: iTipo, sDescricao, dtData
'O Livro Fiscal relacionado ao Item de Apura��o ICMS de Tipo %s, descri��o %s e data %s j� foi fechado.
Public Const ERRO_LOCK_REGAPURACAOICMSITEM = 13387 'Par�metros: iTipo, sDescricao, dtData
'Erro na tentativa de fazer lock no Item de Apura��o ICMS de Tipo %s, descri��o %s e data %s.
Public Const ERRO_EXCLUSAO_REGAPURACAOICMSITEM = 13388 'Par�metros: iTipo, sDescricao, dtData
'Erro na tentativa de excluir o Item de Apura��o de ICMS de de tipo %s, descri��o %s e data %s.
Public Const ERRO_REGAPURACAOICMSITEM_NAO_CADASTRADA = 13389 'Par�metros: iTipo, sDescricao, dtData
'O item de Apura��o ICMS de tipo %s, descri��o %s e Data %s n�o est� cadastrado no Banco de dados.
Public Const ERRO_ATUALIZACAO_REGAPURACAOICMSITEM = 13390 'Par�metros: iTipo, sDescricao, dtData
'Erro na atualiza��o do Item de Apura��o de ICMS de de tipo %s, descri��o %s e data %s.
Public Const ERRO_INSERCAO_REGAPURACAOICMSITEM = 13391 'Par�metros: iTipo, sDescricao, dtData
'Erro na inser��o do Item de Apura��o de ICMS de de tipo %s, descri��o %s e data %s.
Public Const ERRO_REGINVENTARIO_EXISTENTE = 13392 'Par�metros: dtData
'J� existe um Registro de Invent�rio para a data %s.
Public Const ERRO_INSERCAO_REGINVENTARIO = 13393 'Par�metros: iFilialEmpresa, sProduto, dtData
'Erro na inser��o do Registro de Invent�rio com Filial Empresa de c�digo %s, Produto %s e Data %s.
Public Const ERRO_REGINVENTARIO_NAO_CADASTRADO = 13394 'Par�metros: sProduto, dtData, iFilialEmpresa
'O Registro de invent�rio do Produto %s, de Data %s da Filial Empresa de c�digo %s, n�o est�
'cadastrado no Banco de dados.
Public Const ERRO_LEITURA_REGINVENTARIO = 13395 'Sem par�metros
'Erro na leitura da tabela RegInventario.
Public Const ERRO_REGINVENTARIO_FECHADO = 13396 'Par�metros: dtData, iFilialEmpresa
'O Livro Fiscal relativo ao Registro de invent�rio de Data %s da Filial Empresa
'de c�digo %s j� foi fechado.
Public Const ERRO_LOCK_REGINVENTARIO = 13397 'Par�metros: sProduto, dtData, iFilialEmpresa
'Erro no "lock" de Registro de invent�rio com Produto %s, Data %s e Filial Empresa
'de c�digo %s.
Public Const ERRO_EXCLUSAO_REGINVENTARIO = 13398 'Par�metros: sProduto, dtData, iFilialEmpresa
'Erro na exclus�o do Registro de invent�rio com Produto %s, Data %s e Filial Empresa
'de c�digo %s.
Public Const ERRO_ATUALIZACAO_REGINVENTARIO = 13399 'Par�metros: sProduto, dtData, iFilialEmpresa
'Erro na atualiza��o do Registro de invent�rio com Produto %s, Data %s e Filial Empresa
'de c�digo %s.
Public Const ERRO_SLDDIAESTALM_NAO_CADASTRADO = 13400 'Par�metros: sProduto, iAlmoxarifado
'N�o foram encontrados registros em SldDiaEstAlm com Produto %s e Almoxarifado de c�digo %s.
Public Const ERRO_SLDDIAEST_NAO_CADASTRADO = 13401 'Par�metros: sProduto
'N�o foram econtrados registros em SldDiaEst para o Produto %s.
Public Const ERRO_ESTOQUEMES_ABERTO_NAOAPURADO = 13402 'Par�metros: iMes, iFilialEmpresa
'Para gerar o Registro de invent�rio, o m�s %s da FilialEmpresa %s tem que estar fechado e apurado.
Public Const ERRO_REGINVENTARIO_NAO_CADASTRADO1 = 13403 'Par�metros: dtData
'N�o existe Registro de Invent�rio para a data %s.
Public Const ERRO_INTERVALO_DATA_DIFERENTE_LIVROFISCAL = 13404 'Par�metros: dtDataInicial, dtDataFinal
'As data inicial %s e a data final %s tem que estar dentro do intervalo das datas do Livro Fiscal aberto,
'ou de um livro fiscal fechado.
Public Const ERRO_LIVRO_FISCAL_ABERTO_INEXISTENTE = 13405 'Par�metros: iCodLivro
'N�o existe Livro Fiscal de c�digo %i aberto.
Public Const ERRO_REGAPURACAOIPI_NAO_CADASTRADA = 13406 'Par�metros: dtDataInicial, dtDataFinal, iFilialEmpresa
'O Registro de apura��o de IPI de per�odo %s at� o per�odo %s da Filial Empresa %s n�o est�
'cadastrado no Banco de dados.
Public Const ERRO_ESTADO_INICIAL_MAIOR = 13407 'Sem par�metros
'O Estado inicial n�o pode ser maior que o Estado final.
Public Const ERRO_DATA_REGINVENTARIO_FORA_INTERVALO = 13408 'Par�metros: dtData
'N�o existe Livro Fiscal de Registro de invent�rio aberto que possua
'a data %s dentro do seu intervalo de data inicial e final.
Public Const ERRO_LEITURA_REGINVENTARIOALMOX = 13409 'Sem par�metros
'Erro na leitura da tabela de RegInventarioAlmox.
Public Const ERRO_REGINVENTARIOALMOX_NAO_CADASTRADO = 13410 'Par�metros: sProduto, iAlmoxarifado, dtData
'N�o foi encontrado registros em RegInventarioAlmox com Produto %s, Almoxarifado de c�digo %s
'e data %s.
Public Const ERRO_LOCK_REGINVENTARIOALMOX = 13411 'Par�metros: sProduto, iAlmoxarifado, dtData
'Erro no "lock" de RegInventarioAlmox com Produto %s, Almoxarifado de c�digo %s
'e data %s.
Public Const ERRO_EXCLUSAO_REGINVENTARIOALMOX = 13412 'Par�metros: sProduto, iAlmoxarifado, dtData
'Erro na exclus�o de RegInventarioAlmox com Produto %s, Almoxarifado de c�digo %s
'e data %s.
Public Const ERRO_LOCK_REGINVENTARIO1 = 13413 'Par�metros: dtData
'Erro na tentativa de fazer "lock" em RegInventario com data %s.
Public Const ERRO_EXCLUSAO_REGINVENTARIO1 = 13414 'Par�metros: dtData
'Erro na tentativa de excluir Registro de invent�rio com data %s.
Public Const ERRO_EXCLUSAO_REGINVENTARIOALMOX1 = 13415 'Par�metros: dtData
'Erro na tentativa de excluir registros de RegInventarioAlmox com data %s.
Public Const ERRO_LOCK_REGINVENTARIOALMOX1 = 13416 'Par�metros: dtData
'Erro na tentativa de fazer "lock" em RegInventarioAlmox com data %s.
Public Const ERRO_LIVROFILIAL_VINCULADO_REGINVENTARIO = 13417 'Par�metros: dtData
'O Livro Fiscal de Registro de Invent�rio est� vinculado a um Registro de
'invent�rio gerado para a data %s.
Public Const ERRO_LOCK_LIVROSFECHADOS = 13418 'Par�metros: iCodLivro, iFilialEmpresa, dtDataInicial, dtDataFinal
'Erro na tentativa de fazer "Lock" em LivrosFechados. C�digo do Livro = %s, C�digo da Filial Empresa  %s, Data Inicial  = %s e Data Final = %s.
Public Const ERRO_ATUALIZACAO_LIVROSFECHADOS = 13419 'Par�metros: iCodLivro, iFilialEmpresa
'Erro na atualiza��o de LivrosFechados C�digo do Livro = %s, C�digo da Filial Empresa  %s, Data Inicial  = %s e Data Final = %s.
Public Const ERRO_REG_INVENTARIO_SEM_PRODUTO_DATA = 13420 'Par�metro: dtData
'N�o existe cadastro de Produto no estoque antes da data #.
Public Const ERRO_LEITURA_LIVREGESEMITENTES = 13421
'Erro na leitura da tabela de LivRegESEmitentes.
Public Const ERRO_INSERCAO_LIVREGESEMITENTES = 13422
'Erro na Inser��o na Tabela de LivRegESEmitentes.
Public Const ERRO_INSERCAO_LIVREGES = 13423
'Erro na Inser��o na Tabela de LivRegES.
Public Const ERRO_INSERCAO_LIVREGESITEMNF = 13426
'Erro na Inser��o na Tabela de LivRegESItemNF.
Public Const ERRO_MNEMONICO_NAO_ENCONTRADO = 13427  'Parametro: sMnemonico
'O Mnemonico %s n�o foi encontrado.


'jones 28/10

Public Const ERRO_INFOARQICMS_NAO_CADASTRADO = 13428 'Parametros: DataInicial e DataFinal
'Arquivo de ICMS n�o cadastrado. Data Inicial: %s e Data Final %s.
Public Const ERRO_LOCK_INFOARQICMS = 13429 'Sem Par�metros
'Erro na tentativa de fazer "lock" em InfoArqICMS.
Public Const ERRO_EXCLUSAO_INFOARQICMS = 13430 'Sem Par�metros
'Erro na tentativa de excluir registros de InfoArqICMS.
Public Const ERRO_TIPOREGAPURACAOICMS_NAO_ACEITA_LANCAMENTO = 13431 'Par�metros: iCodigo
'O tipo de registro para apura��o ICMS de c�digo %s n�o aceita lancamentos.
Public Const ERRO_LIVROFILIAL_NAO_CONFIGURADO = 13432 'Par�metros: sNomeLivro, iCodigoFilial
'O Livro %s ainda n�o foi configurado para a Filial Empresa %s.
Public Const ERRO_TIPOREGAPURACAOIPI_NAO_CADASTRADA = 13433 'Par�metros: iCodigo
'O tipo de registro para apura��o IPI de c�digo %s n�o est� cadastrado.
Public Const ERRO_LEITURA_TIPOREGAPURACAOIPI = 13434 'Sem par�metros
'Erro na leitura da tabela TipoRegApuracaoIPI.
Public Const ERRO_TIPOREGAPURACAOIPI_PRE_CADASTRADO = 13435 'Par�metros: iCodigo
'O Tipo  de Registro de Apura��o de IPI de c�digo %s � pr�-cadastrado e n�o pode ser excluido nem alterado.
Public Const ERRO_LOCK_TIPOREGAPURACAOIPI = 13436 'Par�metros: iCodigo
'Erro na tentativa de fazer "lock" no Tipo de Registro de apura��o IPI de c�digo %s.
Public Const ERRO_EXCLUSAO_TIPOREGAPURACAOIPI = 13437 'Par�metros: iCodigo
'Erro na exclus�o do Tipo de Registro de apura��o IPI de c�digo %s.
Public Const ERRO_TIPOREGAPURACAO_VINCULADO_REGAPURACAOIPIITEM = 13438 'Par�metros: iCodigo, lNumIntDoc
'O Tipo de Registro de Apura��o de IPI de c�digo %s n�o pode ser excluido pois est� vinculado ao Item de Registro de Apura��o IPI.
Public Const ERRO_ATUALIZACAO_TIPOREGAPURACAOIPI = 13439 'Par�metros: iCodigo
'Erro na atualiza��o do Tipo de Registro de apura��o IPI de c�digo %s.
Public Const ERRO_INSERCAO_TIPOREGAPURACAOIPI = 13440 'Par�metros: iCodigo
'Erro na inser��o do Tipo de Registro de apura��o IPI de c�digo %s.
Public Const ERRO_REGAPURACAOIPIITEM_NAO_CADASTRADA = 13441 'Par�metros: iTipo, sDescricao, dtData
'O item de Apura��o IPI de tipo %s, descri��o %s e Data %s n�o est� cadastrado no Banco de dados.
Public Const ERRO_REGAPURACAOIPIITEM_FECHADO = 13442 'Par�metros: iTipo, sDescricao, dtData
'O Livro Fiscal relacionado ao Item de Apura��o IPI de Tipo %s, descri��o %s e data %s j� foi fechado.
Public Const ERRO_LOCK_REGAPURACAOIPIITEM = 13443 'Par�metros: iTipo, sDescricao, dtData
'Erro na tentativa de fazer lock no Item de Apura��o IPI de Tipo %s, descri��o %s e data %s.
Public Const ERRO_EXCLUSAO_REGAPURACAOIPIITEM = 13444 'Par�metros: iTipo, sDescricao, dtData
'Erro na tentativa de excluir o Item de Apura��o de IPI de de tipo %s, descri��o %s e data %s.
Public Const ERRO_ATUALIZACAO_REGAPURACAOIPIITEM = 13445 'Par�metros: iTipo, sDescricao, dtData
'Erro na atualiza��o do Item de Apura��o de IPI de de tipo %s, descri��o %s e data %s.
Public Const ERRO_INSERCAO_REGAPURACAOIPIITEM = 13446 'Par�metros: iTipo, sDescricao, dtData
'Erro na inser��o do Item de Apura��o de IPI de de tipo %s, descri��o %s e data %s.
Public Const ERRO_TIPOREGAPURACAOIPI_NAO_ACEITA_LANCAMENTO = 13447 'Par�metros: iCodigo
'O tipo de registro para apura��o ICMS de c�digo %s n�o aceita lancamentos.
Public Const ERRO_NENHUMA_APURACAOIPI_CADASTRADA = 13448 'Sem par�metros
'N�o h� Apura��es de IPI anteriores Fechadas cadastradas no Banco de dados.
Public Const ERRO_REGAPURACAOIPI_FECHADO = 13449 'Par�metros: DataInicial, DataFinal, FilialEmpresa
'O Livro Fiscal relacionado ao Registro de Apura��o de IPI de data inicial %s e data final %s da Filial Empresa %s j� foi fechado.
Public Const ERRO_LOCK_REGAPURACAOIPI = 13450 'Par�metros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro no "Lock" do Registro de apura��o de IPI de per�odo %s at� o per�odo %s da Filial Empresa de c�digo %s.
Public Const ERRO_EXCLUSAO_REGAPURACAOIPI = 13451 'Par�metros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na exclus�o do Registro de apura��o de IPI de per�odo %s at� o per�odo %s da Filial Empresa de c�digo %s.
Public Const ERRO_INSERCAO_REGAPURACAOIPI = 13452 'Par�metros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na inser��o do Registro de apura��o de IPI de per�odo %s at� o per�odo %s da Filial Empresa de c�digo %s.
Public Const ERRO_LIVROSAIDA_NAO_CADASTRADO_NFISCAL = 13453 'Parametros = sSerie, lNumNotaFiscal
'Livro de Registro de Sa�da n�o encontrado para a Nota Fiscal com S�rie = %s e N�mero %s.
Public Const ERRO_DATAINVENTARIO_FORA_PERIODO = 13454 'Parametro: dtData
'A Data %s tem que estar dentro de um per�odo de um Livro Aberto ou Fechado.
Public Const ERRO_LEITURA_FISCONFIG = 13455
'Erro na leitura da tabela FisConfig
Public Const ERRO_REGISTRO_FIS_CONFIG_NAO_ENCONTRADO = 13456 'Parametros sCodigo,iFilialEmpresa
'Registro na tabela FISConfig com C�digo=%s e FilialEmpresa=%i n�o foi encontrado.
Public Const ERRO_ATUALIZACAO_FISCONFIG = 13457
'Erro na Atualiza��o da tabela FisConfig




'C�digos de Avisos - Reservado de 15300 a 15399
Public Const AVISO_EXCLUSAO_LIVROSFILIAL = 15300 'Par�metros: iCodLivro, iFilialEmpresa
'Confirma e exclus�o do Livro Fiscal de c�digo %s da Filial Empresa %s?
Public Const AVISO_LIVRO_APURACAO_ICMS_ALTERAR_LIVRO_FOLHA = 15301 'Sem Par�metros
'Somente os campos: Livro e Folha podem ser alterados pois o Livro de Apura��o ICMS
'est� vinculado com uma Apura��o ICMS. Deseja Continuar?
Public Const AVISO_LIVRO_APURACAO_IPI_ALTERAR_LIVRO_FOLHA = 15302 'Sem Par�metros
'Somente os campos: Livro e Folha podem ser alterados pois o Livro de Apura��o IPI
'est� vinculado com uma Apura��o ICMS. Deseja Continuar?
Public Const AVISO_EXCLUSAO_TIPOREGAPURACAOICMS = 15303 'Par�metros: iCodigo
'Confirma a Exclus�o o Tipo de Registro de apura��o ICMS de c�digo %s?
Public Const AVISO_EXCLUSAO_GNRICMS = 15304 'Par�metros: lCodigo
'Confirma a exclus�o da Guia de ICMS de c�digo %s?
Public Const AVISO_EXCLUSAO_REGAPURACAOICMS = 15305 'Par�metros: iFilialEmpresa, dtDataInicial, dtDataFinal
'Confirma a exclus�o do Registro de Apura��o da Filial Empresa %s, Data Inicial %s e Data Final %s?
Public Const AVISO_EXCLUSAO_REGAPURACAOICMSITEM = 15306 'Par�metros: iTipo, sDescricao, dtData
'Confirma a exclus�o do Item de Registro de Apura��o ICMS de Tipo %s, Descri��o %s e Data %s?
Public Const AVISO_CRIAR_TIPOAPURACAOICMS = 15307 'Par�metros: iCodigo
'O Tipo de Registro de Apura��o ICMS de c�digo %s n�o est� cadastrado. Deseja criar?
Public Const AVISO_EXCLUSAO_REGIVENTARIO = 15308 'Par�metros: sProduto, dtData, iFilialEmpresa
'Confirma exclus�o do Registro de invent�rio do Produto %s, de Data %s da Filial Empresa de c�digo %s?
Public Const AVISO_EXCLUSAO_REGIVENTARIOTODOS = 15309 'Par�metros: dtData
'Confirma a exclus�o de todos os registros invent�rios da data %s?
Public Const AVISO_LIVRO_REG_INVENTARIO_ALTERAR_LIVRO_FOLHA = 15310 'Sem par�metros
'Somente os campos: Livro e Folha podem ser alterados pois o Livro de Registro de Invent�rio
'est� vinculado com um Registro de Invent�rio. Deseja Continuar?
Public Const AVISO_CRIAR_REGINVENTARIO = 15311 'Par�metros: dtData
'N�o existem Registros de Invent�rio para a data %s. Desja criar?
Public Const AVISO_EXCLUSAO_INFOARQICMS = 15312 'Par�metros: dtDataInicial, dtDataFinal
'Confirma a exclus�o do ArquivoICMS com a data de %s at� %s. ?
Public Const AVISO_EXCLUSAO_TIPOREGAPURACAOIPI = 15313 'Par�metros: iCodigo
'Confirma a Exclus�o o Tipo de Registro de apura��o IPI de c�digo %s?
Public Const AVISO_CRIAR_TIPOAPURACAOIPI = 15314 'Par�metros: iCodigo
'O Tipo de Registro de Apura��o IPI de c�digo %s n�o est� cadastrado. Deseja criar?
Public Const AVISO_EXCLUSAO_REGAPURACAOIPIITEM = 15315 'Par�metros: iTipo, sDescricao, dtData
'Confirma a exclus�o do Item de Registro de Apura��o IPI de Tipo %s, Descri��o %s e Data %s?
Public Const AVISO_EXCLUSAO_REGAPURACAOIPI = 15316 'Par�metros: iFilialEmpresa, dtDataInicial, dtDataFinal
'Confirma a exclus�o do Registro de Apura��o da Filial Empresa %s, Data Inicial %s e Data Final %s?

