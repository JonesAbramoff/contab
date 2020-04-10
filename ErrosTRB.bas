Attribute VB_Name = "ErrosTRB"
Option Explicit

'C�digos de erro - Reservado de 7000 at� 7099
Public Const ERRO_ICMS_ALIQ_INT_INEXISTENTE = 7000
'Al�quota interna inexistente
Public Const ERRO_ICMS_ALIQ_INTERSTADUAL_INEXISTENTE = 7001
'Al�quota InterEstadual inexistente.
Public Const ERRO_LEITURA_PADROES_TRIBUTACAO = 7002
'Erro na leitura da tabela de Padr�es de Tributa��o.
Public Const ERRO_LEITURA_ICMS_ALIQ_EXT = 7004 'sem parametros
'Erro na leitura da tabela de aliquotas externas para o c�lculo de ICMS
Public Const ERRO_LEITURA_ICMS_EXCECOES = 7005 'sem parametros
'Erro na leitura da tabela de exce��es para o c�lculo de ICMS
Public Const ERRO_LEITURA_TIPOSTRIBICMS = 7006 'sem parametros
'Erro na leitura da tabela de tipos de tributa��o para o c�lculo de ICMS
Public Const ERRO_TIPO_TRIB_ICMS_INEXISTENTE = 7007 'parametro tipo ICMS
'O tipo de tributa��o ICMS %d n�o � valido
Public Const ERRO_LEITURA_TIPOSTRIBIPI = 7008 'Sem Parametros
'Erro na leitura da tabela de tipos de tributa��o para o c�lculo de IPI.
Public Const ERRO_TIPO_TRIB_IPI_INEXISTENTE = 7009 'parametro tipo IPI
'O tipo de tributa��o IPI %d n�o � valido
Public Const ERRO_LEITURA_IPI_EXCECOES = 7010  'Sem parametros
'Erro na leitura da tabela de exce��es para o c�lculo de IPI.
Public Const ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO = 7011 'Sem par�metros
'O peenchimento do Tipo de Tributa��o � obrigat�rio.
Public Const ERRO_DESCRICAO_TIPO_TRIBUTACAO_NAO_PREENCHIDO = 7012 'Sem par�metros
'O preenchimento da Descri��o do Tipo de Tributa��o � obrigat�rio.
Public Const ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBICMS1 = 7014 'Par�metro: TipoTributacaoICMS.Text
'O Tipo de Tributa��o ICMS %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBIPI1 = 7015 'Par�metro: TipoTributacaoIPI.Text
'O Tipo de Tributa��o IPI %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LOCK_TIPOSDETRIBUTACAOMOVTO = 7016 'Par�metro: iTipo
'Erro na tentativa de fazer "lock" num registro da tabela TiposDeTributacao com Tipo %i.
Public Const ERRO_ATUALIZACAO_TIPOSDETRIBUTACAOMOVTO = 7017 'Par�metro: iTipo
'Erro na tentativa de atualizar um registro da tabela TiposDeTributacao com Tipo %i.
Public Const ERRO_INSERCAO_TIPOSDETRIBUTACAOMOVTO = 7018 'Par�metro: iTipo
'Erro na tentativa de inserir um novo registro na tabela TiposDeTributacao com Tipo %i.
Public Const ERRO_EXCLUSAO_TIPOSDETRIBUTACAOMOVTO = 7019 'Par�metro: iTipo
'Erro na tentativa de excluir um registro da tabela TiposDeTributacao com Tipo %i.
Public Const ERRO_LEITURA_PADROESTRIBUTACAO = 7020 'Par�metro: iTipo
'Erro na leitura de tabela de padr�es de tributa��o com Tipo de Tributa��o Padr�o %i.
Public Const ERRO_EXCLUSAO_TIPO_TRIBUTACAO_PADRAO = 7021 'Par�metro: iTipo
'N�o � permitido excluir o Tipo de Tributa��o %i porque � utilizado em padr�o de tributa��o.
Public Const ERRO_TIPO_TRIBUTACAO_ICMS_NAO_PREENCHIDO = 7022 'Sem par�metros
'O preenchimento de Tipo de Tributa��o ICMS � obrigat�rio.
Public Const ERRO_TIPO_TRIBUTACAO_IPI_NAO_PREENCHIDO = 7023 'Sem par�metros
'O preenchimento de Tipo de Tributa��o IPI � obrigat�rio.
Public Const ERRO_LEITURA_PADROESTRIBENTRADA = 7024 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Erro na leitura da tabela PadroesTribEntrada com Natureza Opera��o %s, Tipo de Documento %s, Categoria Produto %s e Item %s.
Public Const ERRO_LOCK_PADROESTRIBENTRADA = 7025 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Erro na tentativa de fazer "lock" na tabela PadroesTribEntrada com Natureza Opera��o %s, Tipo de Documento %s, Categoria Produto %s e Item %s.
Public Const ERRO_ATUALIZACAO_PADROESTRIBENTRADA = 7026 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Erro na tentativa de atualizar um registro na tabela PadroesTribEntrada com Natureza Opera��o %s, Tipo de Documento %s, Categoria Produto %s e Item %s.
Public Const ERRO_INSERCAO_PADROESTRIBENTRADA = 7027 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Erro na tentativa de inserir um novo registro na tabela PadroesTribEntrada com Natureza %s, Tipo de Documento %s, Categoria Produto %s e Item %s.
Public Const ERRO_CAMPOS_PADRAO_TRIBUTACAO_ENTRADA_NAO_PREENCHIDOS = 7028 'Sem par�metros
'Para Todos os Produtos deve ser preenchido o campo Natureza da Opera��o ou Tipo de Documento ou ambos.
Public Const ERRO_PADRAOTRIBENTRADA_NAO_CADASTRADO = 7029 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'O Padr�o de Tributa��o de Entrada com Natureza %s, Tipo de Documento %s, Categoria Produto %s e Item %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_EXCLUSAO_PADROESTRIBENTRADA = 7030 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Erro na tentativa de excluir um registro na tabela PadroesTribEntrada com Natureza Opera��o %s, Tipo de Documento %s, Categoria Produto %s e Item %s.
Public Const ERRO_CAMPOS_PADRAO_TRIBUTACAO_NAO_PREENCHIDOS = 7031 'Sem par�metros
'Para Todos os Clientes deve ser preenchido o campo Natureza da Opera��o ou Tipo de Documento ou ambos.
Public Const ERRO_LEITURA_PADROESTRIBUTACAO1 = 7032 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Erro na leitura da tabela PadroesTribSaida com Natureza Opera��o %s, Tipo Documento %s, Categoria Cliente %s e Item %s.
Public Const ERRO_LOCK_PADROESTRIBUTACAO = 7033 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Erro na tentativa de fazer "lock" na tabela PadroesTribSaida com Natureza Opera��o %s, Tipo Documento %s, Categoria Cliente %s e Item %s.
Public Const ERRO_ATUALIZACAO_PADROESTRIBUTACAO = 7034 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Erro na tentativa de atualizar um registro na tabela PadroesTribSaida com Natureza Opera��o %s, Tipo Documento %s, Categoria Cliente %s e Item %s.
Public Const ERRO_EXCLUSAO_PADROESTRIBUTACAO = 7035 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Erro na tentativa de excluir um registro na tabela PadroesTribuSaida com Natureza Opera��o %s, Tipo Documento %s, Categoria Cliente %s e Item %s.
Public Const ERRO_PADRAO_TRIBUTACAO_NAO_CADASTRADO = 7036 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'O Padr�o de Tributa��o com Natureza %s, Tipo de Documento %s, Categoria Client %s e Item %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_INSERCAO_PADROESTRIBUTACAO = 7037 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Erro na tentativa de inserir um novo registro na tabela PadroesTribSaida com Natureza %s, Documento %s, Categoria Cliente %s e Item %s.
Public Const ERRO_TIPODOCINFO_NAO_CADASTRADO = 7038 'Par�metro: sSiglaMovto
'O Tipo de Documento com Sigla %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_TRIBUTACAOPV = 7039
'Erro na leitura da tabela de TributacaoPV.
Public Const ERRO_EXCLUSAO_TIPO_TRIBUTACAO_NF = 7040 'Par�metro: iTipo
'N�o � permitido excluir o Tipo de Tributa��o %i porque � utilizado em TributacaoNF.
Public Const ERRO_EXCLUSAO_TIPO_TRIBUTACAO_PV = 7041 'Par�metro: iTipo
'N�o � permitido excluir o Tipo de Tributa��o %i porque � utilizado em TributacaoPV.
Public Const ERRO_EXCLUSAO_TIPO_TRIBUTACAO_TIPODOCINFO = 7042 'Par�metro: iTipo
'N�o � permitido excluir o Tipo de Tributa��o %i porque � utilizado em Tipos de Documento.
Public Const ERRO_LEITURA_TRIBPEDIDOBAIXADO = 7043
'Erro na leitura da tabela de Tributa��o de Pedido de Venda Baixado.
Public Const ERRO_EXCLUSAO_TRIBPEDIDOBAIXADO = 7044
'Erro na exclus�o de registro da tabela de Tributa��o de Pedidos de Venda Baixados.
Public Const ERRO_LEITURA_TRIBITEMPEDIDOBAIXADO = 7045
'Erro na leitura da tabela de Tributa��o de Itens de Pedido de Venda Baixados.
Public Const ERRO_TRIBITEMPEDIDO_NAO_ENCONTRADA = 7046 'Par�metro: lNumIntItem
'N�o foi encontrado nenhum registro de Tributa��o para o Item de Pedido com n�mero interno %l.
Public Const ERRO_EXCLUSAO_TRIBITEMPEDIDOBAIXADO = 7047
'Erro na exclus�o na tabela de Tributa��o de Itens de Pedido de Venda Baixados.
Public Const ERRO_LEITURA_TRIBCOMPLPEDIDOBAIXADO = 7048
'Erro na leitura da tabela de Tributa��o de Complemento de Pedido de Venda Baixado.
Public Const ERRO_EXCLUSAO_TRIBCOMPLPEDIDOBAIXADO = 7049
'Erro na exclus�o  de Itens de Pedido de Venda da tabela de Tributa��o Baixado.
Public Const ERRO_TIPOTRIB_INCOMPAT_ENTRADA = 7051 'sem parametros
'Tipo de tributa��o incompat�vel com opera��o de entrada
Public Const ERRO_TIPOTRIB_INCOMPAT_SAIDA = 7052 'sem parametros
'Tipo de tributa��o incompat�vel com opera��o de sa�da
Public Const ERRO_TIPOTRIBMOV_ALT_ENTSAI = 7053 'sem parametros
'N�o pode alterar um tipo quanto a ser entrada ou sa�da
Public Const ERRO_NUMERO_PRIMEIRO_TIPOTRIB_USUARIO = 7054 'Par�metro: NUMERO_PRIMEIRO_TIPOTRIB_USUARIO
'Os tipos de tributa��o cadastrados pelo usu�rio devem ter c�digo superior a %s.



'VEIO DE ERROS FAT
Public Const ERRO_EXCLUSAO_IPIEXCECOES = 8030 'Par�metros: sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'Erro na tentativa de exclus�o de registro da tabela de exce��es de IPI.Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_INSERCAO_IPIEXCECOES = 8031 'Par�metros: sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'Erro na tentativa de inser��o de registro da tabela de exce��es de IPI.Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_ATUALIZACAO_IPIEXCECOES = 8032 'Par�metros: sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'Erro na tentativa de atualiza��o de registro da tabela de exce��es de IPI.Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_IPIEXCECOES_INEXISTENTE = 8034 'Par�metros: sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'O registro n�o existe na tabela de exce��es de IPI.Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_INSERCAO_TRIBITEMNFISCAL = 8062
'Erro na inser��o de registro na tebela de Tributa��o de Itens de N.Fiscal.
Public Const ERRO_INSERCAO_TRIBNFISCAL = 8063
'Erro na inser��o de registro na tebela de Tributa��o de N.Fiscal.
Public Const ERRO_INSERCAO_TRIBCOMPLNFISCAL = 8064
'Erro na inser��o de registro na tabela de Tributa��o de Complemento de N.Fiscal.


'C�digos de Avisos - Reservado de 5500 at� 5599
Public Const AVISO_EXCLUSAO_EXCECAO = 5500 'Sem parametros
'Confirma exclus�o da Exce��o de ICMS?
Public Const AVISO_CONFIRMA_EXCLUSAO_TIPO_TRIBUTACAO = 5501 'Par�metro: iTipo
'Confirma exclus�o do Tipo de Tributa��o %i?
Public Const AVISO_CONFIRMA_EXCLUSAO_PADRAOTRIBENTRADA = 5502 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Confirma exclus�o do Padr�o de Tributa��o Entrada com Natureza %s, Tipo de Documento %s, Categoria Produto %s e Item %s ?
Public Const AVISO_CONFIRMA_EXCLUSAO_PADRAO_TRIBUTACAO = 5503 'Par�metros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Confirma a exclus�o do Padr�o de Tributa��o Sa�da com Natureza %s, Tipo de Documento %s, Categoria Cliente %s e item %s ?
Public Const AVISO_EXCLUSAO_EXCECAO_IPI = 5504 'Sem parametros
'Confirma exclus�o da Exce��o de IPI?

'C�digos de Erro  RESERVADO de 13000 a 13099
Public Const ERRO_ALIQUOTA_IPI_NAO_PREENCHIDA = 13000
'O campo Al�quota de IPI n�o foi preenchido.
Public Const ERRO_CLAS_FISCAL_NAO_PREENCHIDA = 13001
'O campo Clas. Fiscal de IPI n�o foi preenchido.

