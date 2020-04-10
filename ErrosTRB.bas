Attribute VB_Name = "ErrosTRB"
Option Explicit

'Códigos de erro - Reservado de 7000 até 7099
Public Const ERRO_ICMS_ALIQ_INT_INEXISTENTE = 7000
'Alíquota interna inexistente
Public Const ERRO_ICMS_ALIQ_INTERSTADUAL_INEXISTENTE = 7001
'Alíquota InterEstadual inexistente.
Public Const ERRO_LEITURA_PADROES_TRIBUTACAO = 7002
'Erro na leitura da tabela de Padrões de Tributação.
Public Const ERRO_LEITURA_ICMS_ALIQ_EXT = 7004 'sem parametros
'Erro na leitura da tabela de aliquotas externas para o cálculo de ICMS
Public Const ERRO_LEITURA_ICMS_EXCECOES = 7005 'sem parametros
'Erro na leitura da tabela de exceções para o cálculo de ICMS
Public Const ERRO_LEITURA_TIPOSTRIBICMS = 7006 'sem parametros
'Erro na leitura da tabela de tipos de tributação para o cálculo de ICMS
Public Const ERRO_TIPO_TRIB_ICMS_INEXISTENTE = 7007 'parametro tipo ICMS
'O tipo de tributação ICMS %d não é valido
Public Const ERRO_LEITURA_TIPOSTRIBIPI = 7008 'Sem Parametros
'Erro na leitura da tabela de tipos de tributação para o cálculo de IPI.
Public Const ERRO_TIPO_TRIB_IPI_INEXISTENTE = 7009 'parametro tipo IPI
'O tipo de tributação IPI %d não é valido
Public Const ERRO_LEITURA_IPI_EXCECOES = 7010  'Sem parametros
'Erro na leitura da tabela de exceções para o cálculo de IPI.
Public Const ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO = 7011 'Sem parâmetros
'O peenchimento do Tipo de Tributação é obrigatório.
Public Const ERRO_DESCRICAO_TIPO_TRIBUTACAO_NAO_PREENCHIDO = 7012 'Sem parâmetros
'O preenchimento da Descrição do Tipo de Tributação é obrigatório.
Public Const ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBICMS1 = 7014 'Parâmetro: TipoTributacaoICMS.Text
'O Tipo de Tributação ICMS %s não está cadastrado no Banco de Dados.
Public Const ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBIPI1 = 7015 'Parâmetro: TipoTributacaoIPI.Text
'O Tipo de Tributação IPI %s não está cadastrado no Banco de Dados.
Public Const ERRO_LOCK_TIPOSDETRIBUTACAOMOVTO = 7016 'Parâmetro: iTipo
'Erro na tentativa de fazer "lock" num registro da tabela TiposDeTributacao com Tipo %i.
Public Const ERRO_ATUALIZACAO_TIPOSDETRIBUTACAOMOVTO = 7017 'Parâmetro: iTipo
'Erro na tentativa de atualizar um registro da tabela TiposDeTributacao com Tipo %i.
Public Const ERRO_INSERCAO_TIPOSDETRIBUTACAOMOVTO = 7018 'Parâmetro: iTipo
'Erro na tentativa de inserir um novo registro na tabela TiposDeTributacao com Tipo %i.
Public Const ERRO_EXCLUSAO_TIPOSDETRIBUTACAOMOVTO = 7019 'Parâmetro: iTipo
'Erro na tentativa de excluir um registro da tabela TiposDeTributacao com Tipo %i.
Public Const ERRO_LEITURA_PADROESTRIBUTACAO = 7020 'Parâmetro: iTipo
'Erro na leitura de tabela de padrões de tributação com Tipo de Tributação Padrão %i.
Public Const ERRO_EXCLUSAO_TIPO_TRIBUTACAO_PADRAO = 7021 'Parâmetro: iTipo
'Não é permitido excluir o Tipo de Tributação %i porque é utilizado em padrão de tributação.
Public Const ERRO_TIPO_TRIBUTACAO_ICMS_NAO_PREENCHIDO = 7022 'Sem parâmetros
'O preenchimento de Tipo de Tributação ICMS é obrigatório.
Public Const ERRO_TIPO_TRIBUTACAO_IPI_NAO_PREENCHIDO = 7023 'Sem parâmetros
'O preenchimento de Tipo de Tributação IPI é obrigatório.
Public Const ERRO_LEITURA_PADROESTRIBENTRADA = 7024 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Erro na leitura da tabela PadroesTribEntrada com Natureza Operação %s, Tipo de Documento %s, Categoria Produto %s e Item %s.
Public Const ERRO_LOCK_PADROESTRIBENTRADA = 7025 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Erro na tentativa de fazer "lock" na tabela PadroesTribEntrada com Natureza Operação %s, Tipo de Documento %s, Categoria Produto %s e Item %s.
Public Const ERRO_ATUALIZACAO_PADROESTRIBENTRADA = 7026 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Erro na tentativa de atualizar um registro na tabela PadroesTribEntrada com Natureza Operação %s, Tipo de Documento %s, Categoria Produto %s e Item %s.
Public Const ERRO_INSERCAO_PADROESTRIBENTRADA = 7027 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Erro na tentativa de inserir um novo registro na tabela PadroesTribEntrada com Natureza %s, Tipo de Documento %s, Categoria Produto %s e Item %s.
Public Const ERRO_CAMPOS_PADRAO_TRIBUTACAO_ENTRADA_NAO_PREENCHIDOS = 7028 'Sem parâmetros
'Para Todos os Produtos deve ser preenchido o campo Natureza da Operação ou Tipo de Documento ou ambos.
Public Const ERRO_PADRAOTRIBENTRADA_NAO_CADASTRADO = 7029 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'O Padrão de Tributação de Entrada com Natureza %s, Tipo de Documento %s, Categoria Produto %s e Item %s não está cadastrado no Banco de Dados.
Public Const ERRO_EXCLUSAO_PADROESTRIBENTRADA = 7030 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Erro na tentativa de excluir um registro na tabela PadroesTribEntrada com Natureza Operação %s, Tipo de Documento %s, Categoria Produto %s e Item %s.
Public Const ERRO_CAMPOS_PADRAO_TRIBUTACAO_NAO_PREENCHIDOS = 7031 'Sem parâmetros
'Para Todos os Clientes deve ser preenchido o campo Natureza da Operação ou Tipo de Documento ou ambos.
Public Const ERRO_LEITURA_PADROESTRIBUTACAO1 = 7032 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Erro na leitura da tabela PadroesTribSaida com Natureza Operação %s, Tipo Documento %s, Categoria Cliente %s e Item %s.
Public Const ERRO_LOCK_PADROESTRIBUTACAO = 7033 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Erro na tentativa de fazer "lock" na tabela PadroesTribSaida com Natureza Operação %s, Tipo Documento %s, Categoria Cliente %s e Item %s.
Public Const ERRO_ATUALIZACAO_PADROESTRIBUTACAO = 7034 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Erro na tentativa de atualizar um registro na tabela PadroesTribSaida com Natureza Operação %s, Tipo Documento %s, Categoria Cliente %s e Item %s.
Public Const ERRO_EXCLUSAO_PADROESTRIBUTACAO = 7035 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Erro na tentativa de excluir um registro na tabela PadroesTribuSaida com Natureza Operação %s, Tipo Documento %s, Categoria Cliente %s e Item %s.
Public Const ERRO_PADRAO_TRIBUTACAO_NAO_CADASTRADO = 7036 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'O Padrão de Tributação com Natureza %s, Tipo de Documento %s, Categoria Client %s e Item %s não está cadastrado no Banco de Dados.
Public Const ERRO_INSERCAO_PADROESTRIBUTACAO = 7037 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Erro na tentativa de inserir um novo registro na tabela PadroesTribSaida com Natureza %s, Documento %s, Categoria Cliente %s e Item %s.
Public Const ERRO_TIPODOCINFO_NAO_CADASTRADO = 7038 'Parâmetro: sSiglaMovto
'O Tipo de Documento com Sigla %s não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_TRIBUTACAOPV = 7039
'Erro na leitura da tabela de TributacaoPV.
Public Const ERRO_EXCLUSAO_TIPO_TRIBUTACAO_NF = 7040 'Parâmetro: iTipo
'Não é permitido excluir o Tipo de Tributação %i porque é utilizado em TributacaoNF.
Public Const ERRO_EXCLUSAO_TIPO_TRIBUTACAO_PV = 7041 'Parâmetro: iTipo
'Não é permitido excluir o Tipo de Tributação %i porque é utilizado em TributacaoPV.
Public Const ERRO_EXCLUSAO_TIPO_TRIBUTACAO_TIPODOCINFO = 7042 'Parâmetro: iTipo
'Não é permitido excluir o Tipo de Tributação %i porque é utilizado em Tipos de Documento.
Public Const ERRO_LEITURA_TRIBPEDIDOBAIXADO = 7043
'Erro na leitura da tabela de Tributação de Pedido de Venda Baixado.
Public Const ERRO_EXCLUSAO_TRIBPEDIDOBAIXADO = 7044
'Erro na exclusão de registro da tabela de Tributação de Pedidos de Venda Baixados.
Public Const ERRO_LEITURA_TRIBITEMPEDIDOBAIXADO = 7045
'Erro na leitura da tabela de Tributação de Itens de Pedido de Venda Baixados.
Public Const ERRO_TRIBITEMPEDIDO_NAO_ENCONTRADA = 7046 'Parâmetro: lNumIntItem
'Não foi encontrado nenhum registro de Tributação para o Item de Pedido com número interno %l.
Public Const ERRO_EXCLUSAO_TRIBITEMPEDIDOBAIXADO = 7047
'Erro na exclusão na tabela de Tributação de Itens de Pedido de Venda Baixados.
Public Const ERRO_LEITURA_TRIBCOMPLPEDIDOBAIXADO = 7048
'Erro na leitura da tabela de Tributação de Complemento de Pedido de Venda Baixado.
Public Const ERRO_EXCLUSAO_TRIBCOMPLPEDIDOBAIXADO = 7049
'Erro na exclusão  de Itens de Pedido de Venda da tabela de Tributação Baixado.
Public Const ERRO_TIPOTRIB_INCOMPAT_ENTRADA = 7051 'sem parametros
'Tipo de tributação incompatível com operação de entrada
Public Const ERRO_TIPOTRIB_INCOMPAT_SAIDA = 7052 'sem parametros
'Tipo de tributação incompatível com operação de saída
Public Const ERRO_TIPOTRIBMOV_ALT_ENTSAI = 7053 'sem parametros
'Não pode alterar um tipo quanto a ser entrada ou saída
Public Const ERRO_NUMERO_PRIMEIRO_TIPOTRIB_USUARIO = 7054 'Parâmetro: NUMERO_PRIMEIRO_TIPOTRIB_USUARIO
'Os tipos de tributação cadastrados pelo usuário devem ter código superior a %s.



'VEIO DE ERROS FAT
Public Const ERRO_EXCLUSAO_IPIEXCECOES = 8030 'Parâmetros: sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'Erro na tentativa de exclusão de registro da tabela de exceções de IPI.Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_INSERCAO_IPIEXCECOES = 8031 'Parâmetros: sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'Erro na tentativa de inserção de registro da tabela de exceções de IPI.Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_ATUALIZACAO_IPIEXCECOES = 8032 'Parâmetros: sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'Erro na tentativa de atualização de registro da tabela de exceções de IPI.Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_IPIEXCECOES_INEXISTENTE = 8034 'Parâmetros: sCategoriaCliente, sCategoriaClienteItem, sCategoriaProduto,sCategoriaProdutoItem
'O registro não existe na tabela de exceções de IPI.Categoria do Cliente: %s, Valor: %s, Categoria do Produto: %s, Valor %s.
Public Const ERRO_INSERCAO_TRIBITEMNFISCAL = 8062
'Erro na inserção de registro na tebela de Tributação de Itens de N.Fiscal.
Public Const ERRO_INSERCAO_TRIBNFISCAL = 8063
'Erro na inserção de registro na tebela de Tributação de N.Fiscal.
Public Const ERRO_INSERCAO_TRIBCOMPLNFISCAL = 8064
'Erro na inserção de registro na tabela de Tributação de Complemento de N.Fiscal.


'Códigos de Avisos - Reservado de 5500 até 5599
Public Const AVISO_EXCLUSAO_EXCECAO = 5500 'Sem parametros
'Confirma exclusão da Exceção de ICMS?
Public Const AVISO_CONFIRMA_EXCLUSAO_TIPO_TRIBUTACAO = 5501 'Parâmetro: iTipo
'Confirma exclusão do Tipo de Tributação %i?
Public Const AVISO_CONFIRMA_EXCLUSAO_PADRAOTRIBENTRADA = 5502 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaProduto, sItemCategoria
'Confirma exclusão do Padrão de Tributação Entrada com Natureza %s, Tipo de Documento %s, Categoria Produto %s e Item %s ?
Public Const AVISO_CONFIRMA_EXCLUSAO_PADRAO_TRIBUTACAO = 5503 'Parâmetros: sNaturezaOperacao, sSiglaMovto, sCategoriaFilialCliente, sItemCategoria
'Confirma a exclusão do Padrão de Tributação Saída com Natureza %s, Tipo de Documento %s, Categoria Cliente %s e item %s ?
Public Const AVISO_EXCLUSAO_EXCECAO_IPI = 5504 'Sem parametros
'Confirma exclusão da Exceção de IPI?

'Códigos de Erro  RESERVADO de 13000 a 13099
Public Const ERRO_ALIQUOTA_IPI_NAO_PREENCHIDA = 13000
'O campo Alíquota de IPI não foi preenchido.
Public Const ERRO_CLAS_FISCAL_NAO_PREENCHIDA = 13001
'O campo Clas. Fiscal de IPI não foi preenchido.

