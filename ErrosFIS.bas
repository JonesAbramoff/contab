Attribute VB_Name = "ErrosFIS"
Option Explicit

'Códigos de Erro - Reservado de 13300 a 13499
Public Const ERRO_DATAINICIO_MAIOR_DATAFIM = 13300 'Sem parâmetros
'A data inicial não pode ser maior que a data final.
Public Const ERRO_TRIBUTO_NAO_PREENCHIDO = 13301 'Sem parâmetros
'É obrigatório o preenchimento de Tributo.
Public Const ERRO_LIVRO_NAO_PREENCHIDO = 13302 'Sem parâmetros
'É obrigatório o preenchimento do Livro Fiscal.
Public Const ERRO_PERIODICIDADE_NAO_PREENCHIDO = 13303 'Sem parâmetros
'É obrigatório o preenchimento da Periodicidade.
Public Const ERRO_NUMEROLIVRO_NAO_PREENCHIDO = 13304 'Sem parâmetros
'É obrigatório o preenchimento do número do Livro.
Public Const ERRO_FOLHA_NAO_PREENCHIDA = 13305 'Sem parâmetros
'É obrigatório o preenchimento do número do Folha.
Public Const ERRO_LEITURA_TRIBUTOS = 13306 'Sem parâmetros
'Erro na leitura da tabela Tributos.
Public Const ERRO_LEITURA_PERIODICIDADESLIVROSFISC = 13307 'Sem parâmetros
'Erro na leitura da tabela PeriodiciddadesLivrosFisc.
Public Const ERRO_LEITURA_LIVROFISCAL = 13308 'Sem parâmetros
'Erro na leitura da tabela LivrosFiscais.
Public Const ERRO_DATAINICIAL_MAIOR_DATAFINALLIVRO = 13309 'Parâmetros: dtDataFinal, dtDataInicial
'A Data Inicial %s do Livro Fiscal tem que ser maior do que a data
'final do Livro que já Fechado, que é %s.
Public Const ERRO_LOCK_LIVROSFILIAL = 13310 'Parâmetros: iCodLivro, iFilialEmpresa
'Erro na tentativa de fazer "Lock" em LivrosFilial com Livro de código %s e Filial Empresa de código %s.
Public Const ERRO_EXCLUSAO_LIVROSFILIAL = 13311 'Parâmetros: iCodLivro, iFilialEmpresa
'Erro na exclusão de LivrosFilial com Livro de código %s e Filial Empresa de código %s.
Public Const ERRO_ATUALIZACAO_LIVROSFILIAL = 13312 'Parâmetros: iCodLivro, iFilialEmpresa
'Erro na atualização de LivrosFilial com Livro de código %s e Filial Empresa de código %s.
Public Const ERRO_INSERCAO_LIVROSFILIAL = 13313 'Parâmetros: iCodLivro, iFilialEmpresa
'Erro na inserção de LivrosFilial com Livro de código %s e Filial Empresa de código %s.
Public Const ERRO_LIVROFISCAL_NAO_CADASTRADO = 13314 'Parâmetros: iCodigo
'O Livro Fiscal de código %s não está cadastrado no Banco de dados.
Public Const ERRO_LIVROFILIAL_VINCULADO_REGAPURACAOICMS = 13315 'Parâmetros: iCodLivro, iFilialEmpresa, dtDataInicial, dtDataFinal
'O Livro Fiscal de código %s da Filial Empresa de código %s está vinculado a um Registro de Apuração
'de ICMS de Data Inicial %s e Data Final %s.
Public Const ERRO_LIVROFILIAL_VINCULADO_REGAPURACAOIPI = 13316 'Parâmetros: iCodLivro, iFilialEmpresa, dtDataInicial, dtDataFinal
'O Livro Fiscal de código %s da Filial Empresa de código %s está vinculado a um Registro de Apuração
'de IPI de Data Inicial %s e Data Final %s.
Public Const ERRO_LEITURA_REGAPURACAOIPI = 13317 'Sem parâmetros
'Erro na leitura da tabela RegApuracaoIPI.
Public Const ERRO_LEITURA_LIVROSFECHADOS = 13318 'Sem parâmetros
'Erro na leitura da tabela LivrosFechados.
Public Const ERRO_SECAO_NAO_PREENCHIDA = 13319 'Sem parâmetros
'É obrigatório o preenchimento da Seção.
Public Const ERRO_LEITURA_TIPOREGAPURACAOICMS = 13320 'Sem parâmetros
'Erro na leitura da tabela TipoRegApuracaoICMS.
Public Const ERRO_TIPOREGAPURACAOICMS_NAO_CADASTRADA = 13321 'Parâmetros: iCodigo
'O tipo de registro para apuração ICMS de código %s não está cadastrado.
Public Const ERRO_LOCK_TIPOREGAPURACAOICMS = 13322 'Parâmetros: iCodigo
'Erro na tentativa de fazer "lock" no Tipo de Registro de apuração ICMS de código %s.
Public Const ERRO_TIPOREGAPURACAOICMS_PRE_CADASTRADO = 13323 'Parâmetros: iCodigo
'O Tipo  de Registro de Apuração de ICMS de código %s é pré-cadastrado
'e não pode ser excluido nem alterado.
Public Const ERRO_ATUALIZACAO_TIPOREGAPURACAOICMS = 13324 'Parâmetros: iCodigo
'Erro na atualização do Tipo de Registro de apuração ICMS de código %s.
Public Const ERRO_INSERCAO_TIPOREGAPURACAOICMS = 13325 'Parâmetros: iCodigo
'Erro na inserção do Tipo de Registro de apuração ICMS de código %s.
Public Const ERRO_EXCLUSAO_TIPOREGAPURACAOICMS = 13326 'Parâmetros: iCodigo
'Erro na exclusão do Tipo de Registro de apuração ICMS de código %s.
Public Const ERRO_LEITURA_REGAPURACAOICMSITEM = 13327 'Sem parâmetros
'Erro na leitura da tabela RegApuracaoICMSItem.
Public Const ERRO_TIPOREGAPURACAO_VINCULADO_REGAPURACAOICMSITEM = 13328 'Parâmetros: iCodigo, lNumIntDoc
'O Tipo de Registro de Apuração de ICMS de código %s não pode ser excluido pois está vinculado
'ao Item de Registro de Apuração ICMS.
Public Const ERRO_ATUALIZACAO_REGAPURACAOICMSITEM1 = 13329 'Sem Parâmetros
'Erro na atualização do Item de Apuração de ICMS.
Public Const ERRO_ATUALIZACAO_REGAPURACAOIPI = 13330 'Parâmetros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na atualização do Registro de apuração de IPI de período %s até o período %s da Filial Empresa de código %s.
Public Const ERRO_ATUALIZACAO_REGAPURACAOIPIITEM1 = 13331 'Sem Parâmetros
'Erro na atualização do Item de Apuração de IPI.
Public Const ERRO_LEITURA_REGAPURACAOIPIITEM = 13332 'Sem parâmetros
'Erro na leitura da tabela RegApuracaoIPIItem.
Public Const ERRO_NFISCAL_NAO_SELECIONADA = 13333 'Sem Parametros
'É necessário selecionar uma Nota Fiscal.
Public Const ERRO_GRIDLANCAMENTO_VAZIO = 13334 'Sem Parametros
'É nescessário ter pelo menos uma linha no Grid de Lançamentos.
Public Const ERRO_FISCAL_GRID_NAO_PREENCHIDO = 13335 'Parametro: iLinha
'O Campo Fiscal da Linha %s não foi preenchido.
Public Const ERRO_UF_GRID_NAO_PREENCHIDO = 13336 'Parametro: iLinha
'O Campo UF da Linha %s não foi preenchido.
Public Const ERRO_VALORCONTABIL_GRID_NAO_PREENCHIDO = 13337 'Parametro: iLinha
'O Campo Valor Contábil da Linha %s não foi preenchido.
Public Const ERRO_LEITURA_LIVREGES = 13338 'Sem Parametros
'Erro na Leitura da Tabela de LivRegES.
Public Const ERRO_LEITURA_LIVREGESITEM = 13339 'Sem Parametros
'Erro na Leitura da Tabela de LivRegESItem.
Public Const ERRO_LEITURA_LIVREGCADPROD = 13340 'Sem Parametros
'Erro na Leitura da Tabela de LivRegESCadProd.
Public Const ERRO_LIVREGCADPROD_NAO_ENCONTRADO = 13341 'Sem Parametros
'Não foi encontrado o registro em LivRegESCadProd para o Item de Registro de Entrada ou Saída.
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
'Livro de Registro de Entrada não encontrado para a Nota Fiscal com Série = %s e Número %s.
Public Const ERRO_LOCK_LIVREGES = 13350 'Sem Parametros
'Ocorreu um Erro na tentativa de fazer "lock" na Tabela de LivRegES.
Public Const ERRO_ATUALIZACAO_LIVREGES = 13351 'Sem Parametros
'Erro na tentativa de Atualizar LivRegES.
Public Const ERRO_LIVROREGES_JA_FECHADO = 13352 'Sem Parametros
'O Livro Fiscal para a Nota Fiscal (Serie = %s e Número = %s) já foi fechado, por isso não é possível fazer alterações.
Public Const ERRO_LEITURA_GERACAOARQICMS = 13354 'Sem Parametros
'Erro na Leitura da Tabela de GeracaoArqICMS.
Public Const ERRO_LEITURA_GERACAOARQICMSPROD = 13355 'Sem Parametros
'Erro na Leitura da Tabela de GeracaoArqICMSProd.
Public Const ERRO_EXCLUSAO_GERACAOARQICMS = 13356 'Sem Parametros
'Erro na Exclusão na Tabela de GeracaoArqICMS.
Public Const ERRO_EXCLUSAO_GERACAOARQICMSPROD = 13357 'Sem Parametros
'Erro na Exclusão na Tabela de GeracaoArqICMSProd.
Public Const ERRO_LEITURA_INFOARQICMS = 13358 'Sem Parametros
'Erro na Leitura da Tabela de InfoArqICMS.
Public Const ERRO_INFOARQICMS_JA_CADASTRADO = 13359 'Data Inicial, Data Final, Data Inicial Lido, Data Final Lido
'Não é possivel gerar o Arquivo de ICMS para as Datas de %s até %s, pois já existe um arquivo para as datas de %s até %s.
Public Const ERRO_ATUALIZACAO_INFOARQICMS = 13360 'Sem Parametros
'Erro na tentativa de Atualizar InfoArqICMS.
Public Const ERRO_DATA_PAGAMENTO_NAO_PREENCHIDO = 13361 'Sem parâmetros
'O preenchimento de data de pagamento é obrigatório.
Public Const ERRO_TIPOGNR_NAO_PREENCHIDO = 13362 'Sem parâmetros
'O preenchimento do Tipo de Guia é obrigatório.
Public Const ERRO_DATA_VENCIMENTO_NAO_PREENCHIDA = 13363 'Sem parâmetros
'O preenchimento da Data de Vencimento é obrigatória.
Public Const ERRO_DATA_REFERENCIA_NAO_PREENCHIDA = 13364 'Sem parâmetros
 'O preenchimento da Data de Referência é obrigatória.
Public Const ERRO_GNRICMS_VINCULADO_APURACAOICMS = 13365 'Parametro: lCodigo
'A Guia de Recolhimento %s já está associada a uma Apuração ICMS.
Public Const ERRO_GNRICMS_VINCULADO_ARQICMS = 13366 'Parametro: lCodigo
'A Guia de Recolhimento %s já está associada a um Arquivo de ICMS.
Public Const ERRO_LEITURA_GNRICMS = 13367 'Sem parâmetros
'Erro na leitura da tabela GNRICMS.
Public Const ERRO_GNRICMS_NAO_CADASTRADA = 13368 'Parâmetros: lCodigo
'A Guia de ICMS de código %s não está cadastrada no banco de dados.
Public Const ERRO_EXCLUSAO_GNRICMS = 13369 'Parâmetros: lCodigo
'Erro na exclusão da Guia de ICMS de código %s.
Public Const ERRO_ATUALIZACAO_GNRICMS = 13370 'Parâmetros: lCodigo
'Erro na atualização da Guia de ICMS de código %s.
Public Const ERRO_INSERCAO_GNRICMS = 13371 'Parâmetros: lCodigo
'Erro na inserção da Guia de ICMS de código %s.
Public Const ERRO_LOCK_GNRICMS = 13372 'Parâmetros: lCodigo
'Erro na tentativa de fazer "lock" em GNRICMS com Guia ICMS de código %s.
Public Const ERRO_NOME_EMPRESA_NAO_PREENCHIDO = 13373 'Sem parâmetros
'O Nome da Empresa não foi preenchido.
Public Const ERRO_LEITURA_LIVROSFILIAL = 13374 'Sem parâmetros
'Erro na leitura da tabela LivrosFilial.
Public Const ERRO_LIVROFILIAL_NAO_CADASTRADO = 13375 'Parâmetros: iTipoLivro, iCodigoFilial
'Não existe Livro Fiscal aberto com código %s para Filial Empresa %s.
Public Const ERRO_LEITURA_REGAPURACAOICMS = 13376 'Sem parâmetros
'Erro na leitura da tabela RegApuracaoICMS.
Public Const ERRO_NENHUMA_APURACAOICMS_CADASTRADA = 13377 'Sem parâmetros
'Não há Apurações de ICMS anteriores Fechadas cadastradas no Banco de dados.
Public Const ERRO_REGAPURACAOICMS_NAO_CADASTRADA = 13378 'Parâmetros: dtDataInicial, dtDataFinal, iFilialEmpresa
'O Registro de apuração de ICMS de período %s até o período %s da Filial Empresa %s não está
'cadastrado no Banco de dados.
Public Const ERRO_LOCK_REGAPURACAOICMS = 13379 'Parâmetros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro no "Lock" do Registro de apuração de ICMS de período %s até o período %s da Filial Empresa de código %s.
Public Const ERRO_EXCLUSAO_REGAPURACAOICMS = 13380 'Parâmetros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na exclusão do Registro de apuração de ICMS de período %s até o período %s da Filial Empresa de código %s.
Public Const ERRO_ATUALIZACAO_REGAPURACAOICMS = 13381 'Parâmetros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na atualização do Registro de apuração de ICMS de período %s até o período %s da Filial Empresa de código %s.
Public Const ERRO_INSERCAO_REGAPURACAOICMS = 13382 'Parâmetros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na inserção do Registro de apuração de ICMS de período %s até o período %s da Filial Empresa de código %s.
Public Const ERRO_REGAPURACAOICMS_FECHADO = 13383 'Parâmetros: DataInicial, DataFinal, FilialEmpresa
'O Livro Fiscal relacionado ao Registro de Apuração de ICMS de data inicial %s e data final %s da
'Filial Empresa %s já foi fechado.
Public Const ERRO_DATA_APURACAO_DIFERENTE_LIVROFILIAL = 13384 'Parâmetros: iFilialEmpresa, dtDataInicial, dtDataFinal
'As datas Inicial %s e Final %s do Livro Fiscal relacionados a Filial Empresa de código %s não são iguais ao
'da Apuração em questão.
Public Const ERRO_TIPOAPURACAO_NAO_PREENCHIDA = 13385 'Sem parâmetros
'O tipo de apuração não foi preenchido.
Public Const ERRO_REGAPURACAOICMSITEM_FECHADO = 13386 'Parâmetros: iTipo, sDescricao, dtData
'O Livro Fiscal relacionado ao Item de Apuração ICMS de Tipo %s, descrição %s e data %s já foi fechado.
Public Const ERRO_LOCK_REGAPURACAOICMSITEM = 13387 'Parâmetros: iTipo, sDescricao, dtData
'Erro na tentativa de fazer lock no Item de Apuração ICMS de Tipo %s, descrição %s e data %s.
Public Const ERRO_EXCLUSAO_REGAPURACAOICMSITEM = 13388 'Parâmetros: iTipo, sDescricao, dtData
'Erro na tentativa de excluir o Item de Apuração de ICMS de de tipo %s, descrição %s e data %s.
Public Const ERRO_REGAPURACAOICMSITEM_NAO_CADASTRADA = 13389 'Parâmetros: iTipo, sDescricao, dtData
'O item de Apuração ICMS de tipo %s, descrição %s e Data %s não está cadastrado no Banco de dados.
Public Const ERRO_ATUALIZACAO_REGAPURACAOICMSITEM = 13390 'Parâmetros: iTipo, sDescricao, dtData
'Erro na atualização do Item de Apuração de ICMS de de tipo %s, descrição %s e data %s.
Public Const ERRO_INSERCAO_REGAPURACAOICMSITEM = 13391 'Parâmetros: iTipo, sDescricao, dtData
'Erro na inserção do Item de Apuração de ICMS de de tipo %s, descrição %s e data %s.
Public Const ERRO_REGINVENTARIO_EXISTENTE = 13392 'Parâmetros: dtData
'Já existe um Registro de Inventário para a data %s.
Public Const ERRO_INSERCAO_REGINVENTARIO = 13393 'Parâmetros: iFilialEmpresa, sProduto, dtData
'Erro na inserção do Registro de Inventário com Filial Empresa de código %s, Produto %s e Data %s.
Public Const ERRO_REGINVENTARIO_NAO_CADASTRADO = 13394 'Parâmetros: sProduto, dtData, iFilialEmpresa
'O Registro de inventário do Produto %s, de Data %s da Filial Empresa de código %s, não está
'cadastrado no Banco de dados.
Public Const ERRO_LEITURA_REGINVENTARIO = 13395 'Sem parâmetros
'Erro na leitura da tabela RegInventario.
Public Const ERRO_REGINVENTARIO_FECHADO = 13396 'Parâmetros: dtData, iFilialEmpresa
'O Livro Fiscal relativo ao Registro de inventário de Data %s da Filial Empresa
'de código %s já foi fechado.
Public Const ERRO_LOCK_REGINVENTARIO = 13397 'Parâmetros: sProduto, dtData, iFilialEmpresa
'Erro no "lock" de Registro de inventário com Produto %s, Data %s e Filial Empresa
'de código %s.
Public Const ERRO_EXCLUSAO_REGINVENTARIO = 13398 'Parâmetros: sProduto, dtData, iFilialEmpresa
'Erro na exclusão do Registro de inventário com Produto %s, Data %s e Filial Empresa
'de código %s.
Public Const ERRO_ATUALIZACAO_REGINVENTARIO = 13399 'Parâmetros: sProduto, dtData, iFilialEmpresa
'Erro na atualização do Registro de inventário com Produto %s, Data %s e Filial Empresa
'de código %s.
Public Const ERRO_SLDDIAESTALM_NAO_CADASTRADO = 13400 'Parâmetros: sProduto, iAlmoxarifado
'Não foram encontrados registros em SldDiaEstAlm com Produto %s e Almoxarifado de código %s.
Public Const ERRO_SLDDIAEST_NAO_CADASTRADO = 13401 'Parâmetros: sProduto
'Não foram econtrados registros em SldDiaEst para o Produto %s.
Public Const ERRO_ESTOQUEMES_ABERTO_NAOAPURADO = 13402 'Parâmetros: iMes, iFilialEmpresa
'Para gerar o Registro de inventário, o mês %s da FilialEmpresa %s tem que estar fechado e apurado.
Public Const ERRO_REGINVENTARIO_NAO_CADASTRADO1 = 13403 'Parâmetros: dtData
'Não existe Registro de Inventário para a data %s.
Public Const ERRO_INTERVALO_DATA_DIFERENTE_LIVROFISCAL = 13404 'Parâmetros: dtDataInicial, dtDataFinal
'As data inicial %s e a data final %s tem que estar dentro do intervalo das datas do Livro Fiscal aberto,
'ou de um livro fiscal fechado.
Public Const ERRO_LIVRO_FISCAL_ABERTO_INEXISTENTE = 13405 'Parâmetros: iCodLivro
'Não existe Livro Fiscal de código %i aberto.
Public Const ERRO_REGAPURACAOIPI_NAO_CADASTRADA = 13406 'Parâmetros: dtDataInicial, dtDataFinal, iFilialEmpresa
'O Registro de apuração de IPI de período %s até o período %s da Filial Empresa %s não está
'cadastrado no Banco de dados.
Public Const ERRO_ESTADO_INICIAL_MAIOR = 13407 'Sem parâmetros
'O Estado inicial não pode ser maior que o Estado final.
Public Const ERRO_DATA_REGINVENTARIO_FORA_INTERVALO = 13408 'Parâmetros: dtData
'Não existe Livro Fiscal de Registro de inventário aberto que possua
'a data %s dentro do seu intervalo de data inicial e final.
Public Const ERRO_LEITURA_REGINVENTARIOALMOX = 13409 'Sem parâmetros
'Erro na leitura da tabela de RegInventarioAlmox.
Public Const ERRO_REGINVENTARIOALMOX_NAO_CADASTRADO = 13410 'Parâmetros: sProduto, iAlmoxarifado, dtData
'Não foi encontrado registros em RegInventarioAlmox com Produto %s, Almoxarifado de código %s
'e data %s.
Public Const ERRO_LOCK_REGINVENTARIOALMOX = 13411 'Parâmetros: sProduto, iAlmoxarifado, dtData
'Erro no "lock" de RegInventarioAlmox com Produto %s, Almoxarifado de código %s
'e data %s.
Public Const ERRO_EXCLUSAO_REGINVENTARIOALMOX = 13412 'Parâmetros: sProduto, iAlmoxarifado, dtData
'Erro na exclusão de RegInventarioAlmox com Produto %s, Almoxarifado de código %s
'e data %s.
Public Const ERRO_LOCK_REGINVENTARIO1 = 13413 'Parâmetros: dtData
'Erro na tentativa de fazer "lock" em RegInventario com data %s.
Public Const ERRO_EXCLUSAO_REGINVENTARIO1 = 13414 'Parâmetros: dtData
'Erro na tentativa de excluir Registro de inventário com data %s.
Public Const ERRO_EXCLUSAO_REGINVENTARIOALMOX1 = 13415 'Parâmetros: dtData
'Erro na tentativa de excluir registros de RegInventarioAlmox com data %s.
Public Const ERRO_LOCK_REGINVENTARIOALMOX1 = 13416 'Parâmetros: dtData
'Erro na tentativa de fazer "lock" em RegInventarioAlmox com data %s.
Public Const ERRO_LIVROFILIAL_VINCULADO_REGINVENTARIO = 13417 'Parâmetros: dtData
'O Livro Fiscal de Registro de Inventário está vinculado a um Registro de
'inventário gerado para a data %s.
Public Const ERRO_LOCK_LIVROSFECHADOS = 13418 'Parâmetros: iCodLivro, iFilialEmpresa, dtDataInicial, dtDataFinal
'Erro na tentativa de fazer "Lock" em LivrosFechados. Código do Livro = %s, Código da Filial Empresa  %s, Data Inicial  = %s e Data Final = %s.
Public Const ERRO_ATUALIZACAO_LIVROSFECHADOS = 13419 'Parâmetros: iCodLivro, iFilialEmpresa
'Erro na atualização de LivrosFechados Código do Livro = %s, Código da Filial Empresa  %s, Data Inicial  = %s e Data Final = %s.
Public Const ERRO_REG_INVENTARIO_SEM_PRODUTO_DATA = 13420 'Parâmetro: dtData
'Não existe cadastro de Produto no estoque antes da data #.
Public Const ERRO_LEITURA_LIVREGESEMITENTES = 13421
'Erro na leitura da tabela de LivRegESEmitentes.
Public Const ERRO_INSERCAO_LIVREGESEMITENTES = 13422
'Erro na Inserção na Tabela de LivRegESEmitentes.
Public Const ERRO_INSERCAO_LIVREGES = 13423
'Erro na Inserção na Tabela de LivRegES.
Public Const ERRO_INSERCAO_LIVREGESITEMNF = 13426
'Erro na Inserção na Tabela de LivRegESItemNF.
Public Const ERRO_MNEMONICO_NAO_ENCONTRADO = 13427  'Parametro: sMnemonico
'O Mnemonico %s não foi encontrado.


'jones 28/10

Public Const ERRO_INFOARQICMS_NAO_CADASTRADO = 13428 'Parametros: DataInicial e DataFinal
'Arquivo de ICMS não cadastrado. Data Inicial: %s e Data Final %s.
Public Const ERRO_LOCK_INFOARQICMS = 13429 'Sem Parâmetros
'Erro na tentativa de fazer "lock" em InfoArqICMS.
Public Const ERRO_EXCLUSAO_INFOARQICMS = 13430 'Sem Parâmetros
'Erro na tentativa de excluir registros de InfoArqICMS.
Public Const ERRO_TIPOREGAPURACAOICMS_NAO_ACEITA_LANCAMENTO = 13431 'Parâmetros: iCodigo
'O tipo de registro para apuração ICMS de código %s não aceita lancamentos.
Public Const ERRO_LIVROFILIAL_NAO_CONFIGURADO = 13432 'Parâmetros: sNomeLivro, iCodigoFilial
'O Livro %s ainda não foi configurado para a Filial Empresa %s.
Public Const ERRO_TIPOREGAPURACAOIPI_NAO_CADASTRADA = 13433 'Parâmetros: iCodigo
'O tipo de registro para apuração IPI de código %s não está cadastrado.
Public Const ERRO_LEITURA_TIPOREGAPURACAOIPI = 13434 'Sem parâmetros
'Erro na leitura da tabela TipoRegApuracaoIPI.
Public Const ERRO_TIPOREGAPURACAOIPI_PRE_CADASTRADO = 13435 'Parâmetros: iCodigo
'O Tipo  de Registro de Apuração de IPI de código %s é pré-cadastrado e não pode ser excluido nem alterado.
Public Const ERRO_LOCK_TIPOREGAPURACAOIPI = 13436 'Parâmetros: iCodigo
'Erro na tentativa de fazer "lock" no Tipo de Registro de apuração IPI de código %s.
Public Const ERRO_EXCLUSAO_TIPOREGAPURACAOIPI = 13437 'Parâmetros: iCodigo
'Erro na exclusão do Tipo de Registro de apuração IPI de código %s.
Public Const ERRO_TIPOREGAPURACAO_VINCULADO_REGAPURACAOIPIITEM = 13438 'Parâmetros: iCodigo, lNumIntDoc
'O Tipo de Registro de Apuração de IPI de código %s não pode ser excluido pois está vinculado ao Item de Registro de Apuração IPI.
Public Const ERRO_ATUALIZACAO_TIPOREGAPURACAOIPI = 13439 'Parâmetros: iCodigo
'Erro na atualização do Tipo de Registro de apuração IPI de código %s.
Public Const ERRO_INSERCAO_TIPOREGAPURACAOIPI = 13440 'Parâmetros: iCodigo
'Erro na inserção do Tipo de Registro de apuração IPI de código %s.
Public Const ERRO_REGAPURACAOIPIITEM_NAO_CADASTRADA = 13441 'Parâmetros: iTipo, sDescricao, dtData
'O item de Apuração IPI de tipo %s, descrição %s e Data %s não está cadastrado no Banco de dados.
Public Const ERRO_REGAPURACAOIPIITEM_FECHADO = 13442 'Parâmetros: iTipo, sDescricao, dtData
'O Livro Fiscal relacionado ao Item de Apuração IPI de Tipo %s, descrição %s e data %s já foi fechado.
Public Const ERRO_LOCK_REGAPURACAOIPIITEM = 13443 'Parâmetros: iTipo, sDescricao, dtData
'Erro na tentativa de fazer lock no Item de Apuração IPI de Tipo %s, descrição %s e data %s.
Public Const ERRO_EXCLUSAO_REGAPURACAOIPIITEM = 13444 'Parâmetros: iTipo, sDescricao, dtData
'Erro na tentativa de excluir o Item de Apuração de IPI de de tipo %s, descrição %s e data %s.
Public Const ERRO_ATUALIZACAO_REGAPURACAOIPIITEM = 13445 'Parâmetros: iTipo, sDescricao, dtData
'Erro na atualização do Item de Apuração de IPI de de tipo %s, descrição %s e data %s.
Public Const ERRO_INSERCAO_REGAPURACAOIPIITEM = 13446 'Parâmetros: iTipo, sDescricao, dtData
'Erro na inserção do Item de Apuração de IPI de de tipo %s, descrição %s e data %s.
Public Const ERRO_TIPOREGAPURACAOIPI_NAO_ACEITA_LANCAMENTO = 13447 'Parâmetros: iCodigo
'O tipo de registro para apuração ICMS de código %s não aceita lancamentos.
Public Const ERRO_NENHUMA_APURACAOIPI_CADASTRADA = 13448 'Sem parâmetros
'Não há Apurações de IPI anteriores Fechadas cadastradas no Banco de dados.
Public Const ERRO_REGAPURACAOIPI_FECHADO = 13449 'Parâmetros: DataInicial, DataFinal, FilialEmpresa
'O Livro Fiscal relacionado ao Registro de Apuração de IPI de data inicial %s e data final %s da Filial Empresa %s já foi fechado.
Public Const ERRO_LOCK_REGAPURACAOIPI = 13450 'Parâmetros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro no "Lock" do Registro de apuração de IPI de período %s até o período %s da Filial Empresa de código %s.
Public Const ERRO_EXCLUSAO_REGAPURACAOIPI = 13451 'Parâmetros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na exclusão do Registro de apuração de IPI de período %s até o período %s da Filial Empresa de código %s.
Public Const ERRO_INSERCAO_REGAPURACAOIPI = 13452 'Parâmetros: dtDataInicial, dtDataFinal, iFilialEmpresa
'Erro na inserção do Registro de apuração de IPI de período %s até o período %s da Filial Empresa de código %s.
Public Const ERRO_LIVROSAIDA_NAO_CADASTRADO_NFISCAL = 13453 'Parametros = sSerie, lNumNotaFiscal
'Livro de Registro de Saída não encontrado para a Nota Fiscal com Série = %s e Número %s.
Public Const ERRO_DATAINVENTARIO_FORA_PERIODO = 13454 'Parametro: dtData
'A Data %s tem que estar dentro de um período de um Livro Aberto ou Fechado.
Public Const ERRO_LEITURA_FISCONFIG = 13455
'Erro na leitura da tabela FisConfig
Public Const ERRO_REGISTRO_FIS_CONFIG_NAO_ENCONTRADO = 13456 'Parametros sCodigo,iFilialEmpresa
'Registro na tabela FISConfig com Código=%s e FilialEmpresa=%i não foi encontrado.
Public Const ERRO_ATUALIZACAO_FISCONFIG = 13457
'Erro na Atualização da tabela FisConfig




'Códigos de Avisos - Reservado de 15300 a 15399
Public Const AVISO_EXCLUSAO_LIVROSFILIAL = 15300 'Parâmetros: iCodLivro, iFilialEmpresa
'Confirma e exclusão do Livro Fiscal de código %s da Filial Empresa %s?
Public Const AVISO_LIVRO_APURACAO_ICMS_ALTERAR_LIVRO_FOLHA = 15301 'Sem Parâmetros
'Somente os campos: Livro e Folha podem ser alterados pois o Livro de Apuração ICMS
'está vinculado com uma Apuração ICMS. Deseja Continuar?
Public Const AVISO_LIVRO_APURACAO_IPI_ALTERAR_LIVRO_FOLHA = 15302 'Sem Parâmetros
'Somente os campos: Livro e Folha podem ser alterados pois o Livro de Apuração IPI
'está vinculado com uma Apuração ICMS. Deseja Continuar?
Public Const AVISO_EXCLUSAO_TIPOREGAPURACAOICMS = 15303 'Parâmetros: iCodigo
'Confirma a Exclusão o Tipo de Registro de apuração ICMS de código %s?
Public Const AVISO_EXCLUSAO_GNRICMS = 15304 'Parâmetros: lCodigo
'Confirma a exclusão da Guia de ICMS de código %s?
Public Const AVISO_EXCLUSAO_REGAPURACAOICMS = 15305 'Parâmetros: iFilialEmpresa, dtDataInicial, dtDataFinal
'Confirma a exclusão do Registro de Apuração da Filial Empresa %s, Data Inicial %s e Data Final %s?
Public Const AVISO_EXCLUSAO_REGAPURACAOICMSITEM = 15306 'Parâmetros: iTipo, sDescricao, dtData
'Confirma a exclusão do Item de Registro de Apuração ICMS de Tipo %s, Descrição %s e Data %s?
Public Const AVISO_CRIAR_TIPOAPURACAOICMS = 15307 'Parâmetros: iCodigo
'O Tipo de Registro de Apuração ICMS de código %s não está cadastrado. Deseja criar?
Public Const AVISO_EXCLUSAO_REGIVENTARIO = 15308 'Parâmetros: sProduto, dtData, iFilialEmpresa
'Confirma exclusão do Registro de inventário do Produto %s, de Data %s da Filial Empresa de código %s?
Public Const AVISO_EXCLUSAO_REGIVENTARIOTODOS = 15309 'Parâmetros: dtData
'Confirma a exclusão de todos os registros inventários da data %s?
Public Const AVISO_LIVRO_REG_INVENTARIO_ALTERAR_LIVRO_FOLHA = 15310 'Sem parâmetros
'Somente os campos: Livro e Folha podem ser alterados pois o Livro de Registro de Inventário
'está vinculado com um Registro de Inventário. Deseja Continuar?
Public Const AVISO_CRIAR_REGINVENTARIO = 15311 'Parâmetros: dtData
'Não existem Registros de Inventário para a data %s. Desja criar?
Public Const AVISO_EXCLUSAO_INFOARQICMS = 15312 'Parâmetros: dtDataInicial, dtDataFinal
'Confirma a exclusão do ArquivoICMS com a data de %s até %s. ?
Public Const AVISO_EXCLUSAO_TIPOREGAPURACAOIPI = 15313 'Parâmetros: iCodigo
'Confirma a Exclusão o Tipo de Registro de apuração IPI de código %s?
Public Const AVISO_CRIAR_TIPOAPURACAOIPI = 15314 'Parâmetros: iCodigo
'O Tipo de Registro de Apuração IPI de código %s não está cadastrado. Deseja criar?
Public Const AVISO_EXCLUSAO_REGAPURACAOIPIITEM = 15315 'Parâmetros: iTipo, sDescricao, dtData
'Confirma a exclusão do Item de Registro de Apuração IPI de Tipo %s, Descrição %s e Data %s?
Public Const AVISO_EXCLUSAO_REGAPURACAOIPI = 15316 'Parâmetros: iFilialEmpresa, dtDataInicial, dtDataFinal
'Confirma a exclusão do Registro de Apuração da Filial Empresa %s, Data Inicial %s e Data Final %s?

