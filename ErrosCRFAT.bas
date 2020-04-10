Attribute VB_Name = "ErrosCRFAT"
 Option Explicit

'Códigos de erro - Reservado de 6000 até 6999
Public Const ERRO_ITEM_REPETIDO_NO_GRID = 6000 'Sem parâmetros
'Não pode haver repetição do Item de uma Categoria.
Public Const ERRO_LEITURA_CATEGORIACLIENTE = 6001 'Sem parâmetros
'Erro na leitura da tabela de CategoriaCliente.
Public Const ERRO_LOCK_CATEGORIACLIENTE = 6002 'Parâmetro: sCategoria
'Erro na tentativa de fazer 'lock' na tabela de CategoriaCliente com Categoria %s.
Public Const ERRO_INSERCAO_COMISSOES = 6003 'Parametro: iTipoTituo, lNumTitulo
'Erro na tentetiva de inserir comissões de um Documento do Tipo %i com Número Interno = %l.
Public Const ERRO_NOTA_FISCAL_ASSOCIADA_A_FATURA = 6004 'Parametro: número da nota fiscal
'A nota fiscal %l já está associada a uma fatura. Não é possível realizar operação.
Public Const ERRO_LOCK_NFSREC = 6005 'Parametro : lNumnotaFiscal
' Erro na tentativa de "lock" da tabela NfsRec na Nota Fiscal número %l.
Public Const ERRO_LEITURA_COMISSOES = 6006 'Parametro: lNumIntDoc
'Erro na leitura da tabela de Comissoes com número de documento %l.
Public Const ERRO_LEITURA_COMISSOESNF = 6007 'Sem parâmetro
'Erro na leitura de registros de comissões de notas fiscais, da tabela ComissoesNF.
Public Const ERRO_LEITURA_COMISSOESPEDVENDAS = 6008 'Sem parâmetro
'Erro na leitura de registros de comissões de pedidos de venda
Public Const ERRO_COMISSOES_BAIXADA = 6009 'Parametros:lNumIntDoc e iTipoTitulo
'Erro na tentativa de excluir registro da tabela de Comissões, do documento %l, do tipo %i.A comissão já foi baixada.
Public Const ERRO_ATUALIZACAO_COMISSOES = 6010 'Parametro: lNumIntDoc, iTipoTitulo
'Erro na tentativa atualizar as Comissões do Documento de Tipo %i e Número Interno %l na Tabela de Comissões
Public Const ERRO_ATUALIZACAO_COMISSOESNF = 6011 'Parametro: lNumNotaFiscal
'Erro na tentativa de atualizar as Comissões da Nota Fiscal de Número %l
Public Const ERRO_EXCLUSAO_COMISSOESNF = 6012 'Parametros:lNumIntDoc
'Erro na tentativa de excluir registro da tabela de Comissões de Notas Fiscais, do documento %l.
Public Const ERRO_EXCLUSAO_COMISSOES = 6013 'Parametros:lNumIntDoc e iTipoTitulo
'Erro na tentativa de excluir registro da tabela de Comissões, do documento %l, do tipo %i.
Public Const ERRO_INSERCAO_COMISSOESNF = 6014 'Sem parâmetro
'Erro na tentetiva de inserir um registro na tabela de comissões de Notas Fiscais.
Public Const ERRO_NFREC_NAO_CADASTRADA = 6015 'Parametro: número da nota fiscal, série
'A Nota Fiscal com número %l da série %s não está cadastrada.
Public Const ERRO_LEITURA_TIPODOCINFO = 6016 'parametro sigla do doc
'Erro na leitura do tipo de documento %s
Public Const ERRO_LEITURA_ESTADOS1 = 6017 'Parâmetro: sSigla
'Erro na leitura da tabela Estados com o Estado %s.
Public Const ERRO_SIGLA_ESTADO_NAO_CADASTRADA = 6018 'Parâmetro: Sigla.Text
'O Estado %s não está cadastrado.
Public Const ERRO_SIGLA_ESTADO_NAO_PREENCHIDA = 6019 'Sem parâmetro
'O preenchimento da Sigla do Estado é obrigatório.
Public Const ERRO_LEITURA_CATEGS_CLI = 6020  'parametros codigo do cliente e codigo da filial
'Erro na leitura das categorias associadas ao cliente %ld filial %d
Public Const ERRO_REGIAO_VENDA_NAO_CADASTRADA = 6021 'Parametro: iCodigo
'Região de Venda com código %i não está cadastrada no Banco de Dados.
Public Const ERRO_LEITURA_CPRCONFIG = 6022 'Parametro sCodigo
'Erro na leitura do codigo %s na tabela CPRConfig.
Public Const ERRO_LOCK_CPRCONFIG = 6023 'Parametro sCodigo
'Erro no "lock" da tabela CPRConfig. Codigo = %s
Public Const ERRO_VALOR_PORCENTAGEM_JUROS = 6024 'Parametros dValor, lPorcentMaxima
'Porcentagem de Juros %d não está entre 0 e %l.
Public Const ERRO_LEITURA_FILIAISCLIENTES = 6025 'Parametros: lCodCliente, iCodFilial
'Erro na leitura da tabela FiliaisClientes. CodCliente=%l, CodFilial=%i.
Public Const ERRO_LEITURA_CLIENTES = 6026
'Erro na leitura da tabela Clientes.
Public Const ERRO_LEITURA_TIPOSCLIENTE = 6028 'Parametro codigo
'Erro na leitura da tabela TiposDeClientes , codigo = %i.
Public Const ERRO_NOME_RED_DUPLICADO = 6029
'Erro na tentativa de cadastrar novo Cliente com o Nome Reduzido ja existente.
Public Const ERRO_INSERCAO_CLIENTES = 6030 'Parametro codigo Cliente
'Erro na tentativa de gravar Cliente de codigo %s.
Public Const ERRO_INSERCAO_FILIAISCLIENTES = 6031
'Erro na tentativa de gravar na tabela FiliaisClientes.
Public Const ERRO_MODIFICACAO_CLIENTE = 6032 'Parametro codigo do cliente
'Erro na tentativa de modificar tabela de Clientes com codigo = %s.
Public Const ERRO_LOCK_TIPOSCLIENTE = 6033
'Erro na tentativa de "lock" na tabela TiposDeCliente
Public Const ERRO_LEITURA_TABELAPRECO = 6034 'Parametro codigo da tabela Preco
'Erro na leitura da tabela de Precos , %i.
Public Const ERRO_LOCK_TABELAPRECO = 6035
'Erro na tentativa de "lock" na tabela de Precos.
Public Const ERRO_LEITURA_CONDICAOPAGTO = 6036 'Parametro codigo da condicao de Pagto
'Erro na leitura da tabela de Condicao de Pagto , no codigo %i.
Public Const ERRO_LOCK_CONDICAOPAGTO = 6037
'Erro na tentativa de "lock" na tabela de Condicao de Pagto.
Public Const ERRO_LEITURA_MENSAGEM = 6038 'Parametro codigo da Mensagem
'Erro na leitura da tabela de Mensagem , no codigo %i.
Public Const ERRO_LOCK_MENSAGEM = 6039
'Erro na tentativa de "lock" na tabela de Mensagem.
Public Const ERRO_LEITURA_COBRADOR = 6040 'Parametro codigo do Cobrador
'Erro na leitura da tabela de Cobrador , no codigo %i.
Public Const ERRO_LOCK_COBRADOR = 6041
'Erro na tentativa de "lock" na tabela de Cobradores.
Public Const ERRO_LEITURA_TRANSPORTADORA = 6042 'Parametro codigo da Transportadora
'Erro na leitura da tabela de Transportadora , no codigo %i.
Public Const ERRO_LOCK_TRANSPORTADORA = 6043
'Erro na tentativa de "lock" na tabela de Transportadoras.
Public Const ERRO_LEITURA_VENDEDOR = 6044 'Parametro: iCodigo
'Erro na leitura do Vendedor com código %i na tabela de Vendedores.
Public Const ERRO_LOCK_VENDEDOR = 6045
'Erro na tentativa de "lock" na tabela de Vendedores.
Public Const ERRO_LEITURA_REGIAO = 6046 'Parametro codigo da Regiao
'Erro na leitura da tabela de Regioes de Vendas , no codigo %i.
Public Const ERRO_LOCK_REGIAO = 6047
'Erro na tentativa de "lock" na tabela de Regioes de Vendas.
Public Const ERRO_CLIENTE_REL_NF_REC_PEND = 6049 'Parametro: lCodigo
'Erro na exclusão do Cliente com código %l. Está relacionado com Nota Fiscal a Receber Pendente.
Public Const ERRO_CLIENTE_REL_TITULOS_REC = 6050 'Parametro: lCodigo
'Erro na exclusão do Cliente com código %l. Está relacionado com Títulos a Receber.
Public Const ERRO_LEITURA_TITULOS_REC = 6051
'Erro na leitura da tabela TitulosRec.
Public Const ERRO_LEITURA_DEBITOSRECCLI = 6052
'Erro na leitura da tabela DebitosRecCli .
Public Const ERRO_CLIENTE_REL_DEBITOS = 6053 'Parametro: lCodigo
'Não é permitido a exclusão do Cliente com código %l, pois está relacionado com Crédito a Receber.
Public Const ERRO_LEITURA_RECEB_ANTEC = 6054
'Erro na leitura da tabela de RecebAntecipados.
Public Const ERRO_CLIENTE_REL_RECEB_ANTEC = 6055 'Parametro: lCodigo
'Não é permitido a exclusão do Cliente com código %l, pois está relacionado com Recebimento Antecipado.
Public Const ERRO_CLIENTE_SEM_FILIAL = 6056 'Parametro codigo do cliente
'O cliente %l nao está vinculado a nenhuma filial.
Public Const ERRO_EXCLUSAO_CLIENTE = 6057 'Parametro codigo do cliente
'Erro na tentativa de excluir o cliente %l.
Public Const ERRO_EXCLUSAO_FILIAISCLIENTES = 6058 'Parametro codigo do cliente
'Erro na exclusao das filiais do cliente %l.
Public Const ERRO_MODIFICACAO_FILIAISCLIENTES = 6059
'Erro na tentativa de modificacao na tabela de FiliaisClientes.
Public Const ERRO_LOCK_CLIENTES = 6060 'Parametro lCodigoCliente
'Erro na tentativa de "lock" da tabela Clientes, Código Cliente = %l .
Public Const ERRO_FILIAL_DESASSOCIADA_CLIENTE = 6061 'Parametro: sCGC
'Filial de Cliente com CGC %s desassociada de Fornecedor.
Public Const ERRO_CONDICAO_PAGTO_NAO_CADASTRADA = 6062 'Parametro iCodigo
'A Condição de Pagamento com código %i não está cadastrada no Banco de Dados.
Public Const ERRO_LEITURA_PADRAO_COBRANCA = 6063 ' Parametro iPadraoCobranca
'Erro na leitura do Código = %i da tabela Padrões Cobranca .
Public Const ERRO_LEITURA_NOTAS_FISCAIS_REC = 6064
'Erro na leitura da tabela de Notas Fiscais a Receber.
Public Const ERRO_TIPO_CLIENTE_NAO_CADASTRADO = 6065 'Parametro: iCodigo
'Tipo de Cliente com código %i não está cadastrado.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_REC_PEND = 6066
'Erro na leitura da tabela de Notas Fiscais a Receber Pendentes.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_REC_BAIXADAS = 6067
'Erro na leitura da tabela de Notas Fiscais a Receber Baixadas.
Public Const ERRO_CLIENTE_REL_NF_REC = 6068 'Parametro: lCodigo
'Erro na exclusão do Cliente com código %l. Está relacionado com Nota Fiscal a Receber.
Public Const ERRO_CLIENTE_REL_NF_REC_BAIXADA = 6069 'Parametro: lCodigo
'Erro na exclusão do Cliente com código %l. Está relacionado com Nota Fiscal a Receber Baixada.
Public Const ERRO_LEITURA_TITULOS_REC_BAIXADOS = 6070
'Erro na leitura da tabela TitulosRecBaixados.
Public Const ERRO_LEITURA_TITULOS_REC_PEND = 6071
'Erro na leitura da tabela TitulosRecPend.
Public Const ERRO_CLIENTE_REL_TITULOS_REC_PEND = 6072 'Parametro: lCodigo
'Erro na exclusão do Cliente com código %l. Está relacionado com Títulos a Receber Pendentes.
Public Const ERRO_CLIENTE_REL_TITULOS_REC_BAIXADOS = 6073 'Parametro: lCodigo
'Erro na exclusão do Cliente com código %l. Está relacionado com Títulos a Receber Baixados.
Public Const ERRO_LEITURA_MVPERCLI2 = 6074 'Sem parametros
'Erro de leitura na tabela MvPerCli.
Public Const ERRO_LOCK_MVPERCLI = 6075 'Parâmetros: iFilialEmpresa, iExercicio, lCliente, iFilial
'Erro na tentativa de "lock" na tabela MvPerCli. FilialEmpresa=%i, Exercicio=%i, Cliente=%l, Filial=%i.
Public Const ERRO_EXCLUSAO_MVPERCLI = 6076 'Parâmetros: iFilialEmpresa, iExercicio, lCliente, iFilial
'Erro na exclusão de registro na tabela MvPerCli. FilialEmpresa=%i, Exercicio=%i, Cliente=%l, Filial=%i.
Public Const ERRO_LOCK_MVDIACLI = 6078 'Parâmetros: iFilialEmpresa, lCliente, iFilial, dtData
'Erro na tentativa de "lock" na tabela MvDiaCli. FilialEmpresa=%i, Cliente=%l, Filial=%i, Data=%dt.
Public Const ERRO_EXCLUSAO_MVDIACLI = 6079 'Parâmetros: iFilialEmpresa, lCliente, iFilial, dtData
'Erro na exclusão de registro na tabela MvPerCli. FilialEmpresa=%i, Cliente=%l, Filial=%i, Data=%dt.
Public Const ERRO_LEITURA_CARTEIRAS_COBRADOR = 6080
'Erro na leitura da tabela Carteiras Cobrador
Public Const ERRO_PADRAO_COBRANCA_INVALIDO = 6081
'O Padrão de Cobrança %d é inválido.
Public Const ERRO_CARTEIRA_COBRADOR_INEXISTENTE = 6082
'Carteira Cobrador inexistente.
Public Const ERRO_LEITURA_PADRAO_COBRANCA2 = 6083
'Erro na leitura da tabela de Padrões de Cobranca.
Public Const ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO = 6084 'Parametro: iCodigo
'Condição de Pagamento com código %i não é de Contas a Receber.
Public Const ERRO_MENSAGEM_NAO_CADASTRADA = 6085 'Parametro: iCodMensagem
'A Mensagem com código %i não está cadastrada no Banco de Dados.
Public Const ERRO_COBRADOR_NAO_CADASTRADO = 6086 'Parametro iCodCobrador
'O Cobrador com código %i não está cadastrado no Banco de Dados.
Public Const ERRO_TRANSPORTADORA_NAO_CADASTRADA = 6087 'Parametro: iCodTransportadora
'A Transportadora com código %i não está cadastrada no Banco de Dados.
Public Const ERRO_VENDEDOR_NAO_CADASTRADO = 6088 'Parametro: iCodVendedor
'O Vendedor com código %i não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_COMISSOES_BAIXA = 6089    'Sem parâmetro
'Erro na leitura de registros de comissões a serem baixadas
Public Const ERRO_LOCK_CATEGORIACLIENTEITEM = 6090  'Sem parametro
'Erro na tentativa de fazer 'lock' na tabela de Estados.
Public Const ERRO_CATEGORIACLIENTEITEM_INEXISTENTE = 6091 'Parametro: sItem, sCategoria
'O Item %s da Categoria %s de Cliente não existe.
Public Const ERRO_CATEGORIACLIENTE_INEXISTENTE = 6092 'Parametro: sCategoria
'A Categoria %s de Cliente não existe.
Public Const ERRO_LEITURA_CATEGORIACLIENTE1 = 6093 'Parâmetro: sCategoria
'Erro na leitura da tabela de CategoriaCliente com Categoria %s.
Public Const ERRO_LEITURA_CATEGORIACLIENTEITEM = 6094  'Parâmetro: sCategoria
'Erro na leitura da tabela de CategoriaClienteItem com Categoria %s.
Public Const ERRO_CATEGORIACLIENTE_NAO_INFORMADA = 6095 'Sem parâmetros
'O preenchimento de Categoria é obrigatório.
Public Const ERRO_FALTA_ITEM_CATEGORIACLIENTE = 6096 'Sem parâmetros
'Somente a Descrição do Item foi informada. O Item não foi informado.
Public Const ERRO_MODIFICACAO_CATEGORIACLIENTE = 6097 'Parâmetro: sCategoria
'Erro na modificação da tabela de CategoriaCliente com Categoria %s.
Public Const ERRO_MODIFICACAO_CATEGORIACLIENTEITEM = 6098 'Parâmetro: sCategoria
'Erro na modificação da tabela de CategoriaClienteItem com Categoria %s.
Public Const ERRO_EXCLUSAO_CATEGORIACLIENTEITEM = 6099 'Parâmetro: sCategoria
'Erro na exclusão de registro na tabela CategoriaClienteItem com Categoria %s.
Public Const ERRO_INSERCAO_CATEGORIACLIENTE = 6100 'Parâmetro: sCategoria
'Erro na inserção da Categoria %s de Cliente na tabela CategoriaCliente.
Public Const ERRO_INSERCAO_CATEGORIACLIENTEITEM = 6101 'Parâmetro: sCategoria
'Erro na inserção de registro na tabela de CategoriaClienteItem com Categoria %s.
Public Const ERRO_LEITURA_FILIALCLIENTECATEGORIAS = 6102 'Parâmetros: sCategoria, sItem
'Erro na leitura da tabela FilialClienteCategorias com Categoria %s e Item %s.
Public Const ERRO_LEITURA_FILIALCLIENTECATEGORIAS1 = 6103 'Parâmetro: sCategoria
'Erro na leitura da tabela FilialClienteCategorias com Categoria %s.
Public Const ERRO_CATEGORIACLIENTEITEM_UTILIZADA = 6104  'Parâmetros: lCliente, sItem, sCategoria
'O Cliente %l está associado ao Ítem %s da Categoria %s.
Public Const ERRO_CATEGORIACLIENTE_UTILIZADA = 6105 'Parâmetro: sCategoria, lCliente
'A Categoria %s já foi utilizada pelo Cliente %l.
Public Const ERRO_CATEGORIACLIENTE_NAO_CADASTRADA = 6106 'Parâmetro: sCategoria
'A Categoria %s não está cadastrada no Banco de Dados.
Public Const ERRO_EXCLUSAO_CATEGORIACLIENTE = 6107  'Parâmetro: sCategoria
'Erro na exclusão de uma Categoria %s de Cliente.
Public Const ERRO_LEITURA_TIPODECLIENTECATEGORIAS1 = 6108 'Parâmetro: sCategoria
'Erro de leitura da tabela TipoDeClienteCategorias Com Categoria %s.
Public Const ERRO_EXCLUSAO_CATEGORIACLIENTE_UTILIZADA = 6109 'Parâmetro: sCategoria, iTipo
'Não é permitido excluir a Categoria %s porque está sendo utilizada pelo Tipo de Cliente %i.
Public Const ERRO_LEITURA_TIPODECLIENTECATEGORIAS = 6110 'Parâmetro: iTipoCliente
'Erro na leitura da tabela TipoDeClienteCategorias com Tipo de Cliente %i.
Public Const ERRO_LEITURA_CATEGORIACLIENTEITEM1 = 6111 'Parâmetro: sItem, sCategoria
'Erro na leitura da tabela CategoriaClienteItem, cuja Categoria é %s e Item %s.
Public Const ERRO_CODIGO_NAO_PREENCHIDO = 6112 'Sem parametros
' Preenchimento do código é obrigatório.
Public Const ERRO_INSERCAO_FILIALCLIENTECATEGORIAS = 6113 'Parâmero: lCodigo
'Erro na tentativa de inserir um registro na tabela FilialClienteCategorias com código do Cliente %l.
Public Const ERRO_LEITURA_FILIALCLIENTECATEGORIAS2 = 6114 'Parâmetros: lCliente
'Erro na leitura da tabela FilialClienteCategorias com Cliente %l.
Public Const ERRO_ALTERACAO_FILIALCLIENTECATEGORIAS = 6115 'Parâmetros: lCliente
'Erro na tentativa deatualizar um registro na tabela FilialClienteCategorias com Cliente %l.
Public Const ERRO_EXCLUSAO_FILIALCLIENTECATEGORIAS = 6116 'Parâmetro: lCodCliente
'Erro na exclusão das Categorias das Filiais do Cliente %l.
Public Const ERRO_EXCLUSAO_FILIALCLIENTECATEGORIAS1 = 6117 'Parâmetro: lCodCliente, iFilial
'Erro na exclusão das Categorias das Filiais do Cliente %l e Filial %i.
Public Const ERRO_LOCK_CATEGORIACLIENTEITEM2 = 6118 'Parâmetros: sCategoria, sItem
'Erro na tentativa de fazer "lock" na tabela CategoriaClienteItem com Categoria %s e Item %s.
Public Const ERRO_ITEM_CATEGORIA_NAO_CADASTRADO = 6119 'Parâmetros: sItem, sCategoria
'O Item %s da Categoria %s não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_CRFATCONFIG = 6120 'Parametro sCodigo
'Erro na leitura do codigo %s na tabela CRFATConfig.
Public Const ERRO_FILIALCLIENTE_NOME_DUPLICADO = 6121 'Parametro sNomeFilial
'O nome %s já está sendo usado em outra Filial de Cliente.
Public Const ERRO_LOCK_CRFATCONFIG = 6122 'Parametro sCodigo
'Erro no "lock" da tabela CRFATConfig. Codigo = %s
Public Const ERRO_ATUALIZACAO_CRFATCONFIG = 6124 'Parametro sCodigo
'Erro ao atualizar o registro de configuração que possui o codigo %s na tabela CRFATConfig.
Public Const ERRO_CATEGORIACLIENTE_TAMMAX = 6126 'parametros: tam max da categoria
'A categoria deve ter no máximo %i caracteres.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_CLI = 6127
'Erro na leitura da nota fiscal.
Public Const ERRO_EXISTEM_NOTAS_FISCAIS_CLI = 6128
'Existe ao menos uma nota fiscal cadastrada para este cliente.
Public Const ERRO_LEITURA_TIPODOCINFO1 = 6129 'Sem parametro
'Erro na leitura da tabela dos tipos de documentos.
Public Const ERRO_LEITURA_NOTA_FISCAL_NUM_SERIE = 6130 'Parametros: Série, Número e Filial
'Erro na leitura da nota fiscal série %s número %l da filial %d
Public Const ERRO_NFISCAL_NUM_SERIE_NAO_CADASTRADA = 6131 'Parametros: Série, Número e Filial
'A nota fiscal série %s número %l da filial %d não está cadastrada ou já foi baixada.
Public Const ERRO_ATUALIZACAO_ESTADOS = 6132  'Parâmetro: sSigla
'Erro na tentativa de atualizar um registro na tabela de Estados com Estado de sigla %s.
Public Const ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO1 = 6133 'Parametro: Tipo.Text
'O Tipo de Documento %s não está cadastrado.
Public Const ERRO_LOCK_TIPOSDOCINFO = 6134 'Parâmetro: sSiglaMovto
'Erro na tentativa de fazer "lock" na tabela TiposDocInfo com Documento %s.
Public Const ERRO_LEITURA_NATUREZAOP = 6135 'Parametro: sCodigo
'Erro na leitura da tabela de Naturezas de Operação %s.
Public Const ERRO_LOCK_NATUREZAOP = 6136 'Parametro: sCodigo
'Não conseguiu fazer o lock da Natureza de Operação %s.
Public Const ERRO_ATUALIZACAO_NATUREZAOP = 6137 'Parametro sCodigo
'Erro de atualização da Natureza de Operação %s.
Public Const ERRO_INSERCAO_NATUREZAOP = 6138 'Parametro: sCodigo
'Erro na inserção da Natureza de Operação %s.
Public Const ERRO_NATUREZAOP_INEXISTENTE = 6139 'Parametro: sCodigo
'A Natureza de Operação %s não está cadastrada.
Public Const ERRO_EXCLUSAO_NATUREZAOP = 6140 'Parametro: sCodigo
'Houve um erro na exclusão da Natureza de Operação %s.
Public Const ERRO_LEITURA_NATUREZAOP1 = 6141 'Sem Parametros
'Erro na leitura da tabela de Naturezas de Operação.
Public Const ERRO_CATEGORIA_CLIENTE_NAO_PREENCHIDA = 6142 'Sem parâmetro
'O preenchimento da Categoria do Cliente é obrigatório.
Public Const ERRO_CATEGORIA_CLIENTE_ITEM_NAO_PREENCHIDA = 6143 'Sem parâmetro
'O preenchimento do item da Categoria do Cliente é obrigatório.
Public Const ERRO_TITRECPEND_JA_CADASTRADO = 6144  'Parametros: lNumTitulo, sSiglaDocumento
'Atenção. Já existe um título a receber pendente cadastrado com esta identificação. Título = %l, Sigla = %s.
Public Const ERRO_TITRECBAIXA_JA_CADASTRADO = 6145  'Parametros: lNumTitulo, sSiglaDocumento
'Atenção. Já existe um título a receber baixado cadastrado com esta identificação. Título = %l, Sigla = %s.
Public Const ERRO_TITREC_JA_CADASTRADO = 6146  'Parametros: lNumTitulo, sSiglaDocumento
'Atenção. Já existe um título a receber cadastrado com esta identificação. Título = %l, Sigla = %s.
Public Const ERRO_NFISCAL_NUMINTDOCCPR_NAO_ZERO = 6147 'Parametros: Série, Número e Filial
'A nota fiscal série %s número %l da filial %i já está associada a um título.
Public Const ERRO_ATUALIZACAO_NFISCAL = 6148  'Parametros: Série, Número e Filial
'Ocorreu um erro na tentativa de atualizar um registro na tabela de Notas Fiscais. Série = %s, Número = %l, Filial = %i
Public Const ERRO_FILIALCLIENTE_NAO_CADASTRADA = 6149 'Parametros: lCodCliente, iCodFilial
'A Filial %i do Cliente com código %l não está cadastrada no Banco de Dados.
Public Const ERRO_LOCK_FILIAISCLIENTES = 6150 'Parametros: lCodCliente, iCodFilial
'Erro na tentativa de "lock" da tabela FiliaisClientes. CodCliente = %l, CodFilial = %i.
Public Const ERRO_TIPO_NAO_SELECIONADO = 6151 'Parâmetro: iCodigo
'Tipo com código %i está na List e não foi selecionado.
Public Const ERRO_LEITURA_CLIENTE_PEDIDOS_DE_VENDA = 6152 'Parâmetro: lCodigo
'Erro na leitura das tabelas PedidosDeVenda e PedidosDeVendabaixados com Cliente %l.
Public Const ERRO_CLIENTE_REL_PED_VENDA = 6153 'Parâmetros: lCodigo, iFilialEmpresa, lPedido
'O Cliente %l está relacionada com Pedido de Venda com Filial Empresa = %i, Código = %l.
Public Const ERRO_CLIENTE_REL_CHEQUE_PRE = 6154 'Parâmetros: lCodigo, lNumIntCheque
'O Cliente %l está associado com Cheque Pré.
Public Const ERRO_LEITURA_CHEQUEPRE2 = 6155 'Parâmetro: lCodigo
'Erro na leitura da tabela de ChequePre com Cliente %l.
Public Const ERRO_CATEGORIA_SEM_VALOR_CORRESPONDENTE = 6156 'Parâmetro: sCategoria
'A Categoria %s não tem um valor correspondente.
Public Const ERRO_CATEGORIACLIENTE_REPETIDA_NO_GRID = 6157 'Sem parâmetro
'Não pode haver repetição de Categorias de Cliente no Grid.
Public Const ERRO_LEITURA_TIPOSDECLIENTE1 = 6158 'Parâmetro: iCodigo
'Erro na leitura da tabela TipoDeCliente com Código %i.
Public Const ERRO_LEITURA_CONDICAOPAGTO_PEDIDOS_DE_VENDA = 6159 'Parâmetro: iCodigo
'Erro na leitura das tabelas PedidosDeVenda e PedidosDeVendabaixados com Condição de Pagamento %i.
Public Const ERRO_CONDICAOPAGTO_REL_PED_VENDA = 6160 'Parâmetros: iCodigo, lPedido, iFilial
'A Condição de Pagamento %i está vinculada a Pedido de Venda %l da Filial Empresa %i.
Public Const ERRO_LEITURA_CONDICOESPAGTO = 6161 'Sem parâmetros
'Erro na leitura da tabela CondicoesPagto.
Public Const ERRO_DIA_DO_MES_INVALIDO = 6162 'Sem parâmetros
'Dia do Mês tem que estar entre 1 e 30.
Public Const ERRO_DESCRICAO_REDUZIDA_NAO_PREENCHIDA = 6163 'Sem parâmetros
'Descricão Reduzida deve ser preenchida.
Public Const ERRO_NUMERO_PARCELAS_NAO_PREENCHIDA = 6164 'Sem parâmetros
'Número de Parcelas deve ser preenchido.
Public Const ERRO_DIAS_PARA_PRIMEIRA_PARCELA_NAO_PREENCHIDA = 6165 'Sem parâmetros
'Dias para Primeira Parcela deve ser preenchido.
Public Const ERRO_DIA_DO_MES_NAO_PREENCHIDO = 6166 'Sem parâmetros
'Dia do Mês deve ser preenchido.
Public Const ERRO_INTERVALO_ENTRE_PARCELAS_NAO_PREENCHIDO = 6167 'Sem parâmetros
'Intervalo entre Parcelas deve ser preenchido.
Public Const ERRO_DESCRICAO_REDUZIDA_CONDICAOPAGTO_REPETIDA = 6168 'Sem parâmetros
'Descrição Reduzida é atributo de outra Condição de Pagamento.
Public Const ERRO_INSERCAO_CONDICAOPAGTO = 6169 'Parâmetro: iCodigo
'Erro na inserção da Condição de Pagamento %i.
Public Const ERRO_ATUALIZACAO_CONDICAOPAGTO = 6170 'Parâmetro: iCodigo
'Erro na atualização da Condição de Pagamento %i.
Public Const ERRO_CONDICAOPAGTO_RELACIONADA_COM_CLIENTE = 6171 'Parâmetro: lTotal
'Condição de Pagamento está relacionada com %l Cliente(s).
Public Const ERRO_CONDICAOPAGTO_RELACIONADA_COM_FORNECEDOR = 6172 'Parâmetro: lTotal
'Condição de Pagamento está relacionada com %l Fornecedores.
Public Const ERRO_CONDICAOPAGTO_RELACIONADA_COM_TIPOSDECLIENTE = 6173 'Parâmetro: lTotal
'Condição de Pagamento está relacionada com %l Tipo(s) de Cliente(s).
Public Const ERRO_CONDICAOPAGTO_RELACIONADA_COM_TIPOSDEFORNECEDOR = 6174 'Parâmetro: lTotal
'Condição de Pagamento está relacionada com %s Tipo(s) de Fornecedor.
Public Const ERRO_EXCLUSAO_CONDICAOPAGTO = 6175 'Parâmetro: iCodigo
'Erro na exclusão da Condição de Pagamento %i.
Public Const ERRO_NUM_PARCELAS_EXCESSIVO = 6176 'Parâmetros: sNumero, iNumMaximo
'O número de parcelas %s ultrapassou o limite máximo %i.
Public Const ERRO_LEITURA_FILIALCLIENTE_PEDIDOS_DE_VENDA = 6177 'Parâmetro: lCodCliente, iCodFilial
'Erro na leitura das tabelas PedidosDeVenda e PedidosDeVendabaixados com Cliente %l e Filial %i.
Public Const ERRO_FILIALCLIENTE_REL_PED_VENDA = 6178 'Parâmetros: lCodCliente, iCodFilial, iFilial, lPedido
'Filial Cliente com chave codCliente=%l, codFilial=%i está relacionada com Pedido de Venda com chave Filial Empresa = %i, Código = %l.
Public Const ERRO_PERCENTAGEM_EMISSAO_NAO_PREENCHIDA = 6179 'Sem parâmetros
'O Percentual de Comissão de Emissão deve ser preenchido.
Public Const ERRO_SOMA_EMISSAO_MAIS_BAIXA = 6180 'Sem parâmetros
'O Percentual de Comissão de Emissão mais o de Baixa não dá 100%.
Public Const ERRO_LEITURA_TIPOSDEVENDEDOR = 6181 'Parâmetro: iCodigo
'Erro na leitura do Tipo de Vendedor %i.
Public Const ERRO_TIPODEVENDEDOR_NAO_CADASTRADO = 6182 'Parâmetro: iCodigo
'Tipo de Vendedor %i não está cadastrado.
Public Const ERRO_INSERCAO_TIPOSDEVENDEDOR = 6183 'Parâmetro: iCodigo
'Erro na inserção do Tipo de Vendedor %i.
Public Const ERRO_ATUALIZACAO_TIPOSDEVENDEDOR = 6184 'Parâmetro: iCodigo
'Erro na atualização do Tipo de Vendedor %i.
Public Const ERRO_EXCLUSAO_TIPOSDEVENDEDOR = 6185 'Parâmetro: iCodigo
'Erro na exclusão do Tipo de Vendedor %i.
Public Const ERRO_LOCK_TIPOSDEVENDEDOR = 6186 'Parâmetro: iCodigo
'Não conseguiu fazer o lock do Tipo de Vendedor %i.
Public Const ERRO_TIPODEVENDEDOR_RELACIONADO_VENDEDOR = 6187 'Parâmetros: iCodTipoVendedor, iCodVendedor
'Tipo de Vendedor %i está relacionado com Vendedor %i.
Public Const ERRO_DESCRICAO_TIPO_VENDEDOR_REPETIDA = 6188 'Parâmetro: iCodigo
'Tipo de Vendedor %i tem a mesma descrição.
Public Const ERRO_DESCRICAO_NAO_PREENCHIDA = 6189 'Sem parâmetros
'Preenchimento da descrição é obrigatório.
Public Const ERRO_LEITURA_FILIAISCLIENTES1 = 6190 'Parâmetro: iCodigo
'Erro na leitura da tabela FiliaisClientes com Vendedor %i.
Public Const ERRO_LEITURA_COMISSOESPEDVENDAS1 = 6191 'Parâmetro: iCodigo
'Erro na leitura da tabela ComissoesPedVendas com Vendedor %i.
Public Const ERRO_LEITURA_COMISSOESPEDVENDASBAIXADOS1 = 6192 'Parâmetro: iCodigo
'Erro na leitura da tabela ComissoesPedVendasBaixados com Vendedor %i.
Public Const ERRO_LEITURA_COMISSOESNF1 = 6193 'Parâmetro: iCodigo
'Erro na leitura da tabela ComissoesNF com Vendedor %i.
Public Const ERRO_VENDEDOR_REL_FILIALCLIENTE = 6194 'Parâmetros: iCodigo, lCliente, iFilial
'O Vendedor %i está relacionado com Cliente %l da Filial %i.
Public Const ERRO_VENDEDOR_REL_COMISSAOPEDVENDA = 6195 'Parâmetros: iCodigo, lPedido
'O Vendedor %i está relacionado com Comissão de Pedido de Venda %l.
Public Const ERRO_VENDEDOR_REL_COMISSAOPEDVENDABAIXADA = 6196 'Parâmetros: iCodigo, lPedido
'O Vendedor %i está relacionado com Comissão de Pedido de Venda Baixado %l.
Public Const ERRO_VENDEDOR_REL_COMISSAONF = 6197 'Parâmetros: iCodigo, lNumInt
'O Vendedor %i está relacionado com Comissão de Nota Fiscal.
Public Const ERRO_TIPO_VENDEDOR_NAO_ENCONTRADO = 6198 'Parametro sTipoVendedor
'Tipo Vendedor com descrição %s não foi encontrado.
Public Const ERRO_REGIAO_VENDA_NAO_ENCONTRADA = 6199 'Parametro sRegiaoVenda
'Região de Venda com descrição %s não foi encontrada.
Public Const ERRO_NOME_REDUZIDO_VENDEDOR_REPETIDO = 6200 'Parametro: iCodigo
'Vendedor %i tem o mesmo Nome Reduzido.
Public Const ERRO_INSERCAO_VENDEDOR = 6201 'Parametro iCodigo
'Erro na inserção do Vendedor %i.
Public Const ERRO_INSERCAO_ENDERECO = 6202 'Parametro lCodigo
'Erro na inserção do Endereço %l.
Public Const ERRO_ATUALIZACAO_VENDEDOR = 6203 'Parametro iCodigo
'Erro na atualização do Vendedor %i.
Public Const ERRO_ATUALIZACAO_ENDERECO = 6204 'Parametro iCodigo
'Erro na atualização do Endereço %i.
Public Const ERRO_VENDEDOR_RELACIONADO_COMISSAO = 6205 'Parametro iCodigo, lNumIntCom
'Vendedor %i está relacionado com Comissão com número interno %l.
Public Const ERRO_EXCLUSAO_VENDEDOR = 6206 'Parametro iCodigo
'Erro na exclusão do Vendedor %i.
Public Const ERRO_TIPO_VENDEDOR_NAO_PREENCHIDO = 6207 'Sem Parametros
'Preenchimento do Tipo de Vendedor é obrigatório.
Public Const ERRO_FORNECEDOR_VINCULADO_COND_PAGTO = 6209 'Parametro: iCodCondicaoPagto
'A Condição de Pagamento com código %i não pode deixar de ser usada em Contas a Pagar. Existem Fornecedores vinculados a ela.
Public Const ERRO_LEITURA_TIPOSFORNECEDOR = 6210
'Erro na leitura da tabela TiposdeFornecedor.
Public Const ERRO_TIPO_FORNECEDOR_VINCULADO_COND_PAGTO = 6211 'Parametro: iCodCondicaoPagto
'A Condição de Pagamento com código %i não pode deixar de ser usada em Contas a Pagar. Existem Tipos de Fornecedor vinculados a ela.
Public Const ERRO_CLIENTE_VINCULADO_COND_PAGTO = 6212 'Parametro: iCodCondicaoPagto
'A Condição de Pagamento com código %i não pode deixar de ser usada em Contas a Receber. Existem Clientes vinculados a ela.
Public Const ERRO_LEITURA_TIPOCLIENTE = 6213 'Sem parametros
'Erro na leitura da tabela TiposDeCliente.
Public Const ERRO_TIPO_CLIENTE_VINCULADO_COND_PAGTO = 6214 'Parametro: iCodCondicaoPagto
'A Condição de Pagamento com código %i não pode deixar de ser usada em Contas a Receber. Existem Tipos de Cliente vinculados a ela.
Public Const ERRO_FILIALCLIENTE_REL_DEBITOS = 6215 'Parametros: lCodCliente, iCodFilial
'Não é permitido a exclusão do Cliente %l com código de Filial = %l, pois está relaconada com Crédito a Receber.
Public Const ERRO_FILIALCLIENTE_REL_RECEB_ANTEC = 6216 'Parametros: lCodCliente, iCodFilial
'Não é permitido a exclusão do Cliente %l com código de Filial = %l, pois está relacionada com Recebimento Antecipado.
Public Const ERRO_EXCLUSAO_FILIALCLIENTE = 6217 'Parametros: lCodCliente, iCodFilial
'Erro na tentativa de excluir a Filial Cliente. CodCliente = %l, CodFilial = %i.
Public Const ERRO_LEITURA_FILIALCLIENTECATEGORIAS3 = 6218 'Parâmetros: lCodCliente, iCodFilial
'Erro vna leitura da tabela FilialClienteCategorias com Cliente %l e Filial %i.
Public Const ERRO_LEITURA_REGIOESVENDAS = 6219 'Parametro iCodigo
' Erro na leitura da Região de Venda %i.
Public Const ERRO_LOCK_REGIOESVENDAS = 6220 'Parametro
' Não conseguiu fazer o lock da Região de Venda %i.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_FIL_CLI = 6221
'Erro na leitura da nota fiscal do cliente.
Public Const ERRO_EXISTEM_NOTAS_FISCAIS_FIL_CLI = 6222
'Existe ao menos uma nota fiscal cadastrada para este cliente.
Public Const ERRO_FILIALCLIENTE_REL_TITULOS_REC_PEND = 6223 'Parametros: lCodCliente, iCodFilial
'Erro na exclusão da Filial de Cliente com CodCliente=%l, CodFilial=%i. Está relacionada com Títulos a Receber Pendentes.
Public Const ERRO_FILIALCLIENTE_REL_TITULOS_REC = 6224 'Parametros: lCodCliente, iCodFilial
'Erro na exclusão da Filial de Cliente com CodCliente=%l, CodFilial=%i. Está relacionada com Títulos a Receber.
Public Const ERRO_FILIALCLIENTE_REL_TITULOS_REC_BAIXADOS = 6225 'Parametros: lCodCliente, iCodFilial
'Erro na exclusão da Filial de Cliente com CodCliente=%l, CodFilial=%i. Está relacionada com Títulos a Receber Baixados.
Public Const ERRO_LEITURA_VENDEDORES = 6226 'Sem parametros
'Erro na leitura da tabela Vendedores.
Public Const ERRO_CATEGORIA_SEM_ITEM_CORRESPONDENTE = 6227
'Categoria sem item.
Public Const ERRO_TITULO_REL_NFISCAL = 6228 'Sem Parâmetro
'Não é permitido excluir o Título, porque está relacionado com Nota Fiscal.
Public Const ERRO_CLIENTE_NAO_CADASTRADO1 = 6229 'Parâmetro: sCliente
'O Cliente %s não está cadastrado no Banco de Dados.
Public Const ERRO_VALOR_NAO_PREENCHIDO1 = 6230
'O campo Valor não foi prenchido.
Public Const ERRO_TIPODESCONTO_NAO_ENCONTRADO = 6231  'Parametro: iCodigo
'O Tipo de Desconto %i não foi encontrado.
Public Const ERRO_TIPODESCONTO_NAO_ENCONTRADO1 = 6232  'Parametro: sTipoDesconto
'O Tipo de Desconto %s não foi encontrado.
Public Const ERRO_TITULORECEBER_NAO_CADASTRADO = 6233 'lNumIntDoc
'O Título a Receber com Número Interno %l não está cadastrado.
Public Const ERRO_CHEQUEPRE_NAO_CADASTRADO = 6234 'lNumIntCheque
'O ChequePre com Número Interno = %l não foi encontrado.
Public Const ERRO_TITULOREC_FILIALEMPRESA_DIFERENTE = 6235 'Parâmetro: lNumTitulo, sSiglaDocumento
'Não é possível modificar oTítulo a Receber do Tipo %s e de Número %l. Ela pertence a outra Filial da Empresa.
Public Const ERRO_VALOR_COMISSAO_NAO_INFORMADO = 6236 'iComissao
'O Valor da Comissao %i do Titulo não foi informado.
Public Const ERRO_LEITURA_TITULOSRECBAIXADOS = 6237 'Parametro: lNumTitulo, sSiglaDocumento
'Erro na tentativa de leitura do Titulo Receber Número %l do Tipo %s na Tabela TitulosRecBaixado
Public Const ERRO_LEITURA_TITULOSRECPEND = 6238 'Parametros: lNumTitulo, sSiglaDocumento
'Erro na tentativa de leitura do Titulo Receber Número %l do Tipo %s na Tabela TitulosRecPend
Public Const ERRO_LEITURA_TITULOSREC = 6239 'Parametros: lNumTitulo,  sSiglaDocumento
'Erro na tentativa de leitura do Titulo Receber Número %l do Tipo %s na Tabela TitulosRec
Public Const ERRO_LEITURA_TITULOSREC1 = 6240 'Parametros: lNumIntDoc
'Erro na tentativa de leitura do Titulo Receber com Número Interno %l na Tabela TitulosRec
Public Const ERRO_TITULOREC_PENDENTE_MODIFICACAO = 6241 'Parametro: lNumTitulo, sSiglaDocumento
'Não é possível modificar o Titulo a Receber Número %l do Tipo %s pois ele faz parte de um Lote Pendente.
Public Const ERRO_TITULOREC_BAIXADO_MODIFICACAO = 6242 'Parametro: lNumTitulo, sSiglaDocumento
'Não é possível modificar o Titulo a Receber Número %l do Tipo %s pois ele está baixado.
Public Const ERRO_LOCK_TITULOSREC = 6243 'Parametro: lNumTitulo, sSiglaDocumento
'Erro na tentativa de fazer lock no Titulo Receber Número %l do Tipo %s na Tabela TitulosRec
Public Const ERRO_VENDEDOR_COMISSAO_PARCELA_NAO_INFORMADO = 6244 'Parametro: iComissao,iParcela
'O Vendedor da Comissão %i da Parcela %i não foi informado
Public Const ERRO_VALORISS_MAIOR = 6245 'Parâmetros: sValorISS, sValor
'Valor do ISS não pode ser maior do que o Valor do Titulo
Public Const ERRO_SOMA_PARCELAS_DIFERENTE = 6246 'Parâmetros: dValorParcelas, dValorTitulo
'A soma das Parcelas é diferente do Valor do Titulo
Public Const ERRO_VALORTITULO_MENOS_IMPOSTOS = 6247
'Valor do Título menos Impostos retidos deve ser positivo
Public Const ERRO_TITULORECEBER_SEM_PARCELAS = 6248 'Parâmetro: lNumIntDoc
'Título a Receber com número interno %l não tem Parcelas associadas.
Public Const ERRO_LEITURA_PARCELASREC1 = 6249  'lNumIntTitulo
'Erro na tentativa de ler as Parcelas referenes ao Título de Número Interno %l na tabela de Parcelas a Receber.
Public Const ERRO_VALORBASE_COMISSAO_PARCELA_NAO_INFORMADO = 6250 'Parâmetros: iComissao, iParcela
'O Valor Base da Comissão %i da Parcela %i não foi informado.
Public Const ERRO_VALOR_CHQPRE_PARCELA_NAO_PREENCHIDO = 6251  'Parametro: iParcela
'O Valor do ChequePre da Parcela %i não foi preenchido.
Public Const ERRO_PERCENTUAL_COMISSAO_PARCELA_NAO_INFORMADO = 6252 'iComissao, iParcela
'O Percentual da Comissao %i da Parcela %i não foi informado
Public Const ERRO_PARCELA_RECEBER_NAO_CADASTRADA = 6253 'lNumIntTitulo, iParcela
'A Parcela %i do Titulo a Receber com Número Interno %l não foi encontrada
Public Const ERRO_ATUALIZACAO_CHEQUESPRE = 6254 'lNumIntCheque
'Erro na tentativa de atualizar o CHequePre de Número Interno %l na tabela  ChequesPre.
Public Const ERRO_TITULORECEBER_PENDENTE_EXCLUSAO = 6255 'lNumTitulo, sSiglaDocumento
'Não é possível excluir o Titulo a Receber do Tipo %s e Número %l por que ele faz parte de um Lote pendente.
Public Const ERRO_TITULORECEBER_BAIXADO_EXCLUSAO = 6256 'lNumTitulo, sSiglaDocumento
'Não é possível excluir o Titulo a Receber do Tipo %s e Número %l por que ele está baixado.
Public Const ERRO_TITULORECEBER_NAO_CADASTRADO1 = 6257 'sSiglaDocumento,lNumTitulo
'O titulo a Receber do Tipo %s e Número %l não foi encontrado.
Public Const ERRO_VALOR_CHQPRE_DIFERENTE_DESCONTO = 6258 'iParcela, dDesconto
'O Valor do ChequePre da Parcela %i é diferente do Valor da Parcela com desconto que é %d
Public Const ERRO_DATADEPOSITO_CHQPRE_PARCELA_NAO_PREENCHIDA = 6259 'iParcela
'A Data de Deposito de ChequePre da Parcela %i não foi preenchida
Public Const ERRO_VALOR_COMISSAO_PARCELA_NAO_INFORMADO = 6260 'iComissao, iParcela
'O Valor da Comissão %i da Parcela %i não foi preenchido.
Public Const ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA = 6261 'iComissao, iParcela
'A Data do Desconto %i da Parcela %i não foi preenchida.
Public Const ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO = 6262 'iComissao, iParcela
'O Valor do Desconto %i da Parcela %i não foi preenchido.
Public Const ERRO_CODIGO_DESCONTO_PARCELA_NAO_PREENCHIDO = 6263 'iComissao, iParcela
'O Código do Desconto %i da Parcela %i não foi preenchido.
Public Const ERRO_BANCO_CHQPRE_PARCELA_NAO_PREENCHIDO = 6264 'iParcela
'O Banco do ChequePre da Parcela %i não foi preenchido
Public Const ERRO_NUMERO_CHQPRE_PARCELA_NAO_PREENCHIDO = 6265 'iParcela
'O Número do ChequePre da Parcela %i não foi preenchido
Public Const ERRO_AGENCIA_CHQPRE_PARCELA_NAO_PREENCHIDA = 6266 'iParcela
'A Agência do Cheque Pré da Parcela %i não foi preenchida.
Public Const ERRO_CONTA_CHQPRE_PARCELA_NAO_PREENCHIDA = 6267 'iParcela
'A Conta Corrente do Cheque Pré da Parcela %i não foi preenchida.
Public Const ERRO_SOMA_COMISSOES_PARCELA = 6268 'iParcela
'A soma das Comissões da Parcela %i é maior ou igual ao valor da Parcela.
Public Const ERRO_SOMA_COMISSOES_EMISSAO = 6269
'A soma das Comissões é maior ou igual ao valor do Título.
Public Const ERRO_CARTEIRA_COBRADOR_NAO_INFORMADA = 6270 'Sem parâmetro
'A Carteira do Cobrador deve ser informada.
Public Const ERRO_LEITURA_CARTEIRAS_COBRADOR1 = 6271 'Parâmetro iCobrador
'Erro na leitura da tabela CarteirasCobrador. Cobrador %i.
Public Const ERRO_COBRADOR_SEM_CARTEIRA = 6272 'Parâmetro iCobrador
'O Cobrador %i não possui Carteiras cadastradas
Public Const ERRO_LEITURA_ESTADOS = 6273 'Parâmetro sEstado
'Erro na leitura da tabela Estados. Estado: %s
Public Const ERRO_TIPOCLIENTE_INEXISTENTE1 = 6274 'Parâmetro: sTipoCliente
'Esse Tipo de Cliente não está cadastrado.
Public Const ERRO_LEITURA_TABELA_CARTEIRASCOBRANCA1 = 6275 'Parametro ObjCarteiraCobranca.iCodigo
'Erro na leitura da tabela de Carteiras de Cobrança. Carteira Nº: %s
Public Const ERRO_PARAMETRO_OBRIGATORIO = 6276 'Sem Parametro
'Parametro é obrigatório
Public Const ERRO_EXCLUSAO_PADRAO_COBRANCA = 6277 'Parametro iCodigo
'Erro na exclusão do Padrão de Cobrança %i.
Public Const ERRO_LEITURA_CONTASCORRENTES = 6278
'Erro na leitura da tabela de contas correntes internas
Public Const ERRO_COBRADOR_USADO_BORDEROCOBRANCA = 6279
'O Cobrador está sendo utilizado em um Borderô de Cobrança.
Public Const ERRO_COBRADOR_USADO_OCORRENCIA = 6280
'O Cobrador está sendo utilizado em uma Ocorrência.
Public Const ERRO_COBRADOR_MESMO_NOMEREDUZIDO = 6281 'parâmetro: sNomeReduzido
'Já existe no BD um Cobrador com o Nome Reduzido %s.
Public Const ERRO_COBRADOR_USADO_BAIXASPARCREC = 6282
'O Cobrador foi utilizado em uma Baixa a Receber.
Public Const ERRO_COBRADOR_USADO_FILIAISCLIENTE = 6283
'O Cobrador está sendo utilizado por uma Filial Cliente.
Public Const ERRO_LEITURA_OCORRENCIASREMPARCREC = 6284
'Erro na leitura da tabela OcorrenciasRemParcRec.
Public Const ERRO_COBRADOR_NAO_INFORMADO = 6286 'Sem parametros
'Erro Cobrador não foi Informado
Public Const ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO = 6287 'Parametro iCodigo
'Erro carteira cobrador %i não cadastrada
Public Const ERRO_EXCLUSAO_CARTEIRASCOBRADOR = 6288 'Parametros iCarteira, iCobrador
'Erro exclusão da Carteira %i do cobrador %i
Public Const ERRO_LEITURA_TABELA_COBRADOR = 6289 'Sem parametros
'Erro na leitura da tabela de Cobradores.
Public Const ERRO_LOCK_PADRAO_COBRANCA = 6290 'Parametro iCodigo
'Não conseguiu fazer o lock do Padrão de Cobrança %i.
Public Const ERRO_PADRAO_COBRANCA_RELACIONADO_COM_TIPOS_DE_CLIENTE = 6291 'Parametro lTotal
'Padrão de Cobrança está relacionado com %l Tipos de Cliente.
Public Const ERRO_CONTACORRENTE_INEXISTENTE1 = 6292 'Parametro: CodContaCorrente.Text
'A Conta Corrente %s nao está cadastrada.
Public Const ERRO_DESCRICAO_PADRAO_COBRANCA_REPETIDA = 6293 'Sem Parametros
'Descrição é atributo de outro Padrão de Cobrança.
Public Const ERRO_INSERCAO_PADRAO_COBRANCA = 6294 'Parametro iCodigo
'Erro na inserção do Padrão de Cobrança %i.
Public Const ERRO_PADRAO_COBRANCA_NAO_CADASTRADO = 6295 'Parametro sPadraoCobranca
'O Padrao Cobrança %s não está cadastrado no Banco de Dados.
Public Const ERRO_ATUALIZACAO_PADRAO_COBRANCA = 6296 'Parametro iCodigo
'Erro na atualização do Padrão de Cobrança %i.
Public Const ERRO_BANCO_INEXISTENTE1 = 6297 'Parametro: Banco.Text
'O Banco %s nao está cadastrado.
Public Const ERRO_ATUALIZACAO_CARTEIRASCOBRADOR = 6298 'Parametro objCarteiraCobrador.iCodCarteiraCobranca
'Erro de atualizaçao da Carteira %i
Public Const ERRO_LEITURA_TABELA_CARTEIRASCOBRADOR1 = 6299 'Sem parametros
'Erro na leitura da tabela de Carteiras de Cobrador.
Public Const ERRO_INSERCAO_CARTEIRASCOBRADOR = 6300 'Parametro objCarteiraCobrador.iCodCarteiraCobranca
'Erro de inserção na Carteira %i
Public Const ERRO_LEITURA_TABELA_CARTEIRASCOBRADOR = 6301 'Parametro objcarteiracobrador.icobrador
'Erro de leitura na tabela carteiras cobrador para o cobrador %s
Public Const ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA = 6302 'Parametro objCarteiraCobranca.iCodigo
'Carteira Cobrança não está cadastrada.
Public Const ERRO_LEITURA_BAIXASPARCREC = 6303
'Erro na leitura da tabela BaixasParcRec
Public Const ERRO_BANCO_NAO_INFORMADO = 6304
'O Banco deve ser informado.
Public Const ERRO_CARTEIRA_COBANCA_NAO_INFORMADA = 6305
'A Carteira de Cobrança deve ser informada.
Public Const ERRO_LEITURA_BORDERO_COBRANCA = 6306
'Erro na leitura da tabela de Bordero de Cobrança.
Public Const ERRO_CARTEIRACOBRANCA_VINCULADA_PARCELAS = 6307 'iCarteira
'A Carteira cobrança %i não pode ser excluida por está sendo utilizada por alguma parcela a receber.
Public Const ERRO_LOCK_CARTEIRASCOBRADOR = 6308
'Erro an tentativa de fazer "lock" na tabela CarteirasCobrador.
Public Const ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL = 6309 'Parametro: sContaCorrente, iFilial
'A Conta Corrente %s não pertence à Filial selecionada. Filial: %i.
Public Const ERRO_TIPO_COBRANCA_NAO_INFORMADO = 6310 'Sem Parâmetro
'O Tipo de cobrança deve ser informado.
Public Const ERRO_NUMBORDERO_NAO_INFORMADO = 6311 'Sem Parâmetro
'O Número do Bordero deve ser informado.
Public Const ERRO_LEITURA_TIPOSDECOBRANCA = 6312 'Sem Parâmetro
'Erro na leitura da tabela TiposDeCobrança.
Public Const ERRO_TIPOCOBRANCA_NAO_VALE_BORDERO = 6313 'Parâmetro : sTipoCobranca
'O Tipo de cobrança %s não vale para Bordero.
Public Const ERRO_TIPOCOBRANCA_INEXISTENTE = 6314 'Parâmetro : sTipoCobranca
'O Tipo de cobrança %s não está cadastrado.
Public Const ERRO_CARTEIRA_COBRANCA_NAO_INFORMADA = 6315 'Sem parametros
'Codigo não foi informado
Public Const ERRO_LEITURA_CARTEIRASCOBRADOR = 6316
'Erro na leitura da tabela de Carteiras do Cobrador.
Public Const ERRO_CARTEIRA_COM_COBRADOR = 6317 'Parametro iCodigo
'Não foi possível excluir carteira  %i , existem um ou mais cobradores associados a mesma .
Public Const ERRO_CODIGOCARTCOBR_NAO_PREENCHIDO = 6318 'Sem parametros
'O Código da Carteira de Cobrança não foi informado
Public Const ERRO_LEITURA_NFS_TRANSPORTADORA = 6319
'Erro de leitura na tabela de Notas Fiscais.
Public Const ERRO_EXISTEM_NFS_TRANSPORTADORA = 6320
'Existe Nota Fiscal para esta transportadora.
Public Const ERRO_TRANSPORTADORA_RELACIONADA_FILIAISCLIENTES = 6321 'Parametro: lTotal
'Transportadora esta relacionada com %l Filiais de Clientes
Public Const ERRO_CODTRANSPORTADORA_NAO_PREENCHIDO = 6322
'O código da Transportadora não foi preenchido.
Public Const ERRO_INSERCAO_TRANSPORTADORA = 6323 'Sem parametro
'Erro na inserção da Transportadora.
Public Const ERRO_MODIFICACAO_TRANSPORTADORA = 6324 'Sem parametro
'Erro  na modificação da Transportadora.
Public Const ERRO_ENDERECO_NAO_CADASTRADO = 6325 'Sem parametro
'O Endereco  não está cadastrado no Banco de Dados.
Public Const ERRO_EXCLUSAO_TRANSPORTADORA = 6326 'Parametro codigo da Transportadora
'Erro  na tentativa de excluir Transportadora.
Public Const ERRO_TRANSPORTADORA_REL_NF_REC_PEND = 6327 'Parametro: iCodigo
'Erro na exclusão da Transportadora com código %l, relacionado com Nota Fiscal a Receber Pendente.
Public Const ERRO_TRANSPORTADORA_REL_NF_REC = 6328 'Parametro: iCodigo
'Erro na exclusão da Transportadora com código %l, relacionado com Nota Fiscal a Receber.
Public Const ERRO_TRANSPORTADORA_REL_NF_REC_BAIXADA = 6329 'Parametro: iCodigo
'Erro na exclusão da Transportadora com código %l, relacionado com Nota Fiscal a Receber Baixada.
Public Const ERRO_CONTA_INEXISTENTE = 6330 'Parametro sContaContábil
'A Conta Contabil %s nao existe.
Public Const ERRO_ATUALIZACAO_CARTEIRASCOBRANCA = 6331 'Parametro: Código
'Erro na atualização da Carteira Cobrança %s.
Public Const ERRO_INSERCAO_CARTEIRASCOBRANCA = 6332 'Parametro: Código
'Erro na inclusão da Carteira Cobrança %s.
Public Const ERRO_CARTEIRACOBRANCA_NAO_CADASTRADO = 6333 'Parametro: Código
'A Carteira Cobrança %s não está cadastrada.
Public Const ERRO_LOCK_CARTEIRASCOBRANCA = 6334 'Parametro: Código
'Erro na tentativa de fazer Lock da Carteira Cobrança %s.
Public Const ERRO_EXCLUSAO_CARTEIRASCOBRANCA = 6335 'Parametro: Código
'Erro na exclusão da Carteira Cobrança %s.
Public Const ERRO_NUMERO_PARCELAS_ALTERADO = 6336 'Parâmetros: iNumParcelasTela, iNumParcelasBD
'Não é possível alterar o número de parcelas de um Título lançado. Na Tela: %i. No Banco de Dados: %i.
Public Const ERRO_LOCK_PARCELAS_REC = 6337
'Erro na tentativa de "lock" na Tabela de Parcelas a Receber.
Public Const ERRO_LOCK_COMISSOES = 6338 'Parametro: lNumIntDoc, iTipoTitulo
'Erro na tentativa de fazer lock nas Comissões do Documento de Tipo %i e Número Interno %l na Tabela de Comissões
Public Const ERRO_LEITURA_PARCELASREC = 6339 'Sem parametro
'Erro na tentativa de ler registro na tabela ParcelasRec.
Public Const ERRO_ATUALIZACAO_PARCELASREC = 6340 'Parametro: lNumIntParc
'Erro na Atualizacao da Parcela %s do Título %s da tabela de ParcelasRec.
Public Const ERRO_LEITURA_CHEQUEPRE = 6341 'Parametro: lNumIntCheque
'Erro na leitura da tabela de ChequePre com Número %l.
Public Const ERRO_EXCLUSAO_CHEQUESPRE = 6342 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'Erro na tentativa de excluir o ChequePre com Banco %i, Agência %s, ContaCorrente %s e Número %l da tabela de ChequesPre.
Public Const ERRO_INSERCAO_PARCELAS_REC = 6343
'Erro na inserção de um registro na tabela de Parcelas a Receber.
Public Const ERRO_LOCK_TITULOS_REC = 6344
'Erro na tentativa de "lock" na Tabela Títulos a Receber.
Public Const ERRO_EXCLUSAO_TITULOS_RECEBER = 6345
'Erro na exclusão de um registro da tabela de Títulos a Receber.
Public Const ERRO_PARCELA_COM_BAIXA = 6346 'Parâmetro: iNumParcela, lNumTitulo
'Erro na exclusão da Parcela %i do Título %l. A Parcela tem baixa.
Public Const ERRO_INSERCAO_TITULOS_REC = 6347
'Erro na inserção de um registro na tabela de Títulos a Receber.
Public Const ERRO_LEITURA_TIPOS_INSTRUCAO_COBRANCA = 6348 'Sem Parametros
'Erro na leitura da tabela TipoInstrCobranca
Public Const ERRO_LOCK_INSTRUCAO_COBRANCA = 6349 'Parametro: iInstrucao
'Não conseguiu fazer o lock da Instrução de Cobrança %i.
Public Const ERRO_INSERCAO_CHEQUESPRE = 6350 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'Erro na inserção de ChequePre na tabela de ChequesPre com Banco %i, Agência %s, ContaCorrente %s e Número %l já foi utilizado.
Public Const ERRO_LEITURA_TABELA_CARTEIRASCOBRANCA = 6351 'Sem parametro
'Erro na leitura da tabela de Carteiras de Cobrança.
Public Const ERRO_LEITURA_TIPO_INSTRUCAO_COBRANCA = 6352 'Parametro iCodigo
'Erro na leitura do Tipo Instrução Cobrança %i.
Public Const ERRO_LEITURA_NFISCAL2 = 6354 'Parâmetros: lCodFornecedor, iCodFilial
'Erro na leitura da tabela NFiscal com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_TIPODOCUMENTO1 = 6355 'Sem parâmetros
'Erro na leitura da tabela de TiposdeDocumento.
Public Const ERRO_LOCK_CANALVENDA = 6357 'Parâmetro: iCanalVenda
'Não foi possível fazer o Lock do Cana de Venda %i da tabela CanalVenda.
Public Const ERRO_LEITURA_SERIE1 = 6358 'Parâmetro: sSerie
'Erro de leitura da tabela Serie com Série %s.
Public Const ERRO_LOCK_SERIE1 = 6359 'Parâmetro: sSerie
'Não foi possível fazer o Lock da Série %s da tabela Serie.
Public Const ERRO_LOCK_TRANSPORTADORA1 = 6360 'Parâmetro: iTransportadora
'Não foi possível fazer o Lock da Transportadora %i da tabela Transportadoras.
Public Const ERRO_LEITURA_PARCELAS_REC_NF = 6361 'parametro: lNumNotaFiscal
'Ocorreu um erro na leitura das parcelas a receber vinculadas à nota fiscal %l.
Public Const ERRO_LEITURA_PARCELAS_REC_BAIXADAS_NF = 6362 'parametro: lNumNotaFiscal
'Ocorreu um erro na leitura das parcelas a receber baixadas vinculadas à nota fiscal %l.
Public Const ERRO_LEITURA_TIPOSDECLIENTE2 = 6363 'Parâmetro: iCodigo
'Erro de leitura da tabela TiposDeCliente com Vendedor %i.
Public Const ERRO_VENDEDOR_REL_TIPOCLIENTE = 6364 'Parâmetros: iCodigo, iTipoCliente
'O Vendedor %i está relacionado com o Tipo de Cliente %i.
Public Const ERRO_DATA_SAIDA_MENOR_DATA_EMISSAO = 6365 'Parâmetro: dtDataSaida, dtDataEmissao
'A Data de Saída %dt é anterior a Data de Emissão %dt.
Public Const ERRO_TIPODOC_DIFERENTE_NF_VENDA = 6366 'iTipoDocInfo
'Tipo de Documento %i não é Nota Fiscal de Venda.
Public Const ERRO_FALTA_LOCALIZACAO = 6367 'Parâmetro: sProduto
'O Produto %s não foi localizado. Não é possível gravar a Nota Fiscal.
Public Const ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL = 6368 'Parâmetros: iItem, dDesconto, dPrecoTotal
'Para o Item %i o Desconto %d é maior ou igual ao Preço %d.
Public Const ERRO_VALOR_DESCONTO_ULTRAPASSOU_SOMA_VALORES = 6369 'dDesconto, dSomaValores
'Desconto = %d não pode ultrapassar a soma de Produtos + Frete + Seguro + Despesas = %d.
Public Const ERRO_LOCALIZACAO_ITEM_INEXISTENTE = 6370 'iItem
'Não foi feita a localização do item %i da Nota Fiscal.
Public Const ERRO_LOCALIZACAO_ITEM_INCOMPLETA = 6371 'iItem, dQuantVendida, dQuantAlocada
'A localização do item %i da Nota Fiscal está incompleta. Quantidade vendida: %d. Quantidade localizada: %d.
Public Const ERRO_DATASAIDA_ANTERIOR_DATAEMISSAO = 6372 'dtDataSaida, dtDataEmissao)
'A Data de Saida %dt é anterior a Data de Emissão %dt.
Public Const ERRO_INSERCAO_NFISCAL_SAIDA = 6373 'Parâmetro: lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal
'Erro na inserção da Nota Fiscal com os dados Código do Cliente =%l, Código da Filial =%i, Tipo =%i, Serie =%s e Número NF =%l na tabela de Notas Fiscais.
Public Const ERRO_ALTERACAO_NFISCAL_SAIDA = 6374 'Parâmetros: lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal com os dados Código do Cliente = %l, Código da Filial = %i, Tipo = %i, Serie = %s, Número NF = %l, Data Emissao = %dt está cadastrada no Banco de Dados. Não é possível alterar.
Public Const ERRO_LEITURA_NFISCAL_SAIDA_BAIXADA = 6375 'Parâmetros: iTipoNFiscal, lCliente, iFilialCli, sSerie, lNumNotaFiscal
'Erro na leitura da tabela NFiscalBaixadas na Nota Fiscal com Tipo = %i, Cliente = %l, Filial = %i, Serie = %s e Número = %l.
Public Const ERRO_LEITURA_NFISCAL3 = 6376 'Sem parametros
'Ocorreu um erro na leitura da tabela de Notas Fiscais.
Public Const ERRO_CHEQUEPRE_PARCELAREC_INVALIDO = 6377 'parametro: iParcela
'A parcela %i não pode ser associada a um cheque-pre, pois já está vinculada a outra carteira cobrador.
Public Const ERRO_INSERCAO_TRANSFCARTCOBR = 6378 'sem parametros
'Erro na inserção de registro de transferência de carteira de cobrança
Public Const ERRO_CARTEIRA_COBRADOR_SALDO_NEG = 6379 'parametros: codigo da carteira de cobranca e do cobrador
'A carteira %i do cobrador %i não pode ter saldo negativo em valor
Public Const ERRO_CARTEIRA_COBRADOR_QTDE_NEG = 6380 'parametros: codigo da carteira de cobranca e do cobrador
'A carteira %i do cobrador %i não pode ter número negativo de títulos
Public Const ERRO_LEITURA_CARTEIRA_COBRADOR = 6381 'parametros: codigo da carteira de cobranca e do cobrador
'Erro na leitura da carteira %i do cobrador %i
Public Const ERRO_LEITURA_TIPOCLIENTECATEGORIAS = 6383 'Sem Parâmetros
'Erro de Leitura na Tabela TipoDeClienteCategorias.
Public Const ERRO_CATEGORIACLIENTEITEM_TIPOCLIENTECATEGORIAS = 6386 'Parâmetros: CategoriaCliente e CategoriaClienteItem
'Categoria Cliente %s e Categoria Cliente Item %s são usados na tabela TipoDeClienteCategorias.
Public Const ERRO_CATEGORIACLIENTE_ICMSEXCECOES = 6387 'Parâmetro: CategoriaCliente
'Categoria Cliente %s é usada na Tabela ICMSExcecoes.
Public Const ERRO_CATEGORIACLIENTE_IPIEXCECOES = 6388 'Parâmetro: CategoriaCliente
'Categoria Cliente %s é usada na Tabela IPIExcecoes.
Public Const ERRO_CATEGORIACLIENTE_TIPOCLIENTECATEGORIAS = 6389 'Parâmetro: CategoriaCliente
'Categoria Cliente %s é usada na Tabela TipoDeClienteCategorias.
Public Const ERRO_CATEGORIACLIENTE_FILIALCLIENTECATEGORIAS = 6390 'Parâmetro: CategoriaCliente
'Categoria Cliente %s é usada na tabela FilialClienteCategorias.
Public Const ERRO_LEITURA_PEDIDODEVENDAS1 = 6391 'Sem Parâmetros
'Erro de Leitura na Tabela PedidoDeVendas.
Public Const ERRO_TRANSPORTADORA_PEDIDODEVENDAS = 6392 'Parâmetro: Código da Transportadora
'Transportadora %i é utilizada na tabela PedidoDeVendas.
Public Const ERRO_LEITURA_PEDIDODEVENDASBAIXADOS = 6393 'Sem Parâmetros
'Erro de Leitura na Tabela PedidoDeVendasBaixados.
Public Const ERRO_TRANSPORTADORA_PEDIDODEVENDASBAIXADOS = 6394 'Parâmetro: Código da Transportadora
'Transportadora %i é utilizada na tabela PedidoDeVendasBaixados.
Public Const ERRO_TRANSPORTADORA_FILIAISCLIENTES = 6395 'Parâmetro: Código da Transportadora
'Transportadora %i é utilizada na tabela FiliaisClientes
Public Const ERRO_MODIFICACAO_CARTEIRAS_COBRADOR = 6396
'Erro na atualização na tabela Carteiras Cobrador.
Public Const ERRO_EXCLUSAO_TIPODECLIENTECATEGORIAS = 6397 'Parâmetro: iCodigo
'Erro na tentativa de excluir a Categoria do Tipo de Cliente %i da tabela TiposDeCliente.
Public Const ERRO_INSERCAO_TIPOSDECLIENTE = 6398 'Parâmetro: iCodigo
'Erro na tentativa de inserir um registro na tabela TiposDeCliente com Código %i.
Public Const ERRO_PADRAO_COBRANCA_NAO_CADASTRADO1 = 6399 'Parâmetro: iCodigo
'O Padrão de Cobrança %1 não está cadastrado no Banco de Dados.
Public Const ERRO_INSERCAO_TIPODECLIENTECATEGORIAs = 6400 'Parametro: iCodigo
'Erro na tentativa de inserir o Tipo de Cliente %i na tabela TipoDeClienteCategorias.
Public Const ERRO_MODIFICACAO_TIPOSDECLIENTE = 6401 'Parâmetro: iCodigo
'Erro na modificação da tabela TiposDeCliente com Código %i.
Public Const ERRO_MODIFICACAO_TIPODECLIENTECATEGORIAS = 6402 'Parâmetros: sCategoria, sItem
'Erro na modificação da tabela TipoDeClienteCategorias com Categoria %s e Ítem %i.
Public Const ERRO_PADRAO_COBRANCA_NAO_CADASTRADA = 6406 'Parâmetro: .iCodigo
'O Padrão de Cobrança %i não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_PARCELASREC_NUMINT = 6407 'Parâmetros: lNumIntTitulo, iNumParcela
'Erro na leitura da tabela ParcelasRec com número interno do título %l e número da Parcela %i.
Public Const ERRO_LEITURA_TIPOSINSTRCOBRANCA = 6408 'Sem parâmetros
'Erro na leitura da tabela TiposInstrCobranca.
Public Const ERRO_LEITURA_TITULOSREC2 = 6409 'Parâmetros: iFilialEmpresa, lCliente, iFilial, sSiglaDocumento, lNumTitulo
'Erro na tentativa de leitura do Titulo Receber da Filial Empresa %i, Cliente %l , Filial %i, Sigla do Documento %s e Número do Título %l na tabela TitulosRec.
Public Const ERRO_VALOR_COMISSAO_GRID_NAO_PREENCHIDO = 6410 'Parâmetro: iComissao
'O Valor da comissao %i do Grid de Comissões não está preenchido.
Public Const ERRO_TIPODOC_DIFERENTE_NF_FATURA = 6411 'Parâmetro: iTipoDocInfo
'Tipo de Documento %i não é Nota Fiscal Fatura.
Public Const ERRO_DATA_DESCONTO_INFERIOR_DATA_EMISSAO = 6412 'dtDataDesconto
'Data do Desconto = %dt não pode ser inferior à Data de Emissão da Nota Fiscal.
Public Const ERRO_DATA_DESCONTO_SUPERIOR_DATA_VENCIMENTO = 6413 'dtDataDesconto
'Data de Desconto = %dt não pode ser superior à data de Vencimento
Public Const ERRO_ALTERACAO_NFISCAL_INTERNA = 6414 'sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal Interna com os dados Série = %s , Número = %l, Data Emissão = %dt está cadastrada no Banco de Dados.
Public Const ERRO_PADRAO_COBRANCA_RELACIONADO_COM_CLIENTE = 6415
'O Padrão de Cobrança não pode ser excluído pois está relacionado com um cliente
Public Const ERRO_LOCK_ALCADAFAT = 6416 'Parametro: sCodUsuario
'Ocorreu um erro na tentativa de fazer um lock de um registro da tabela de Alçada Fat. Usuário = %s.
Public Const ERRO_LOCK_VALORLIBERADOCREDITO = 6417 'Parametros: sCodUsuario, iAno
'Ocorreu um erro ao tentar fazer o "lock" de um registro da tabela ValorLiberadoCredito. Usuário = %s, Ano = %i.
'Public Const ERRO_TABELAPRECO_UTILIZADA_TABELAPRECOITEM = 6418 'Parâmetro: iCodTabela
'A Tabela %i não pode ser excluida pois está sendo utilizada em outras Filiais.
'Public Const ERRO_TABELAPRECO_UTILIZADA_NOTAS_FICAIS = 6419 'Parâmetro: iCodTabela
'A Tabela %i não pode ser excluida pois está sendo utilizada em Notas Fiscais.
'Public Const ERRO_TABELAPRECO_UTILIZADA_PEDIDOS_VENDA = 6420 'Parâmetro: iCodTabela
'A Tabela %i não pode ser excluida pois está sendo utilizada em Pedidos de Venda.
Public Const ERRO_LEITURA_NOTAS_FISCAIS = 6421 'Sem Parâmetros
'Erro na leitura de Notas Fiscais.
Public Const ERRO_LEITURA_PEDIDOS_VENDA = 6422 'Sem Parâmetros
'Erro na leitura de Pedidos de Venda.
Public Const ERRO_LEITURA_TABELASDEPRECO = 6423 'Sem parametros
'Erro na leitura da tabela de Tabelas de Preço.
Public Const ERRO_ITEM_NAO_CADASTRADO = 6424 'Parametros: iCodTabela, sCodProduto
'Ítem de Tabela de Preço não está cadastrado no Banco de Dados. Código da Tabela %i e Código do Produto %s.
Public Const ERRO_TABELAPRECO_NAO_PREENCHIDA = 6425 'Sem parametros
'A Tabela deve estar preenchida.
Public Const ERRO_CLIENTE_TABELAPRECO = 6426 'Parametro: iCodigo
'Não é possível excluir a Tabela de Preço %i. Está associada a Clientes.
Public Const ERRO_LOCK_TABELASDEPRECOITENS1 = 6427 'Parametro: iCodigo
'Não conseguiu fazer o lock na tabela de TabelasDePrecoItens com Código da Tabela %i.
Public Const ERRO_EXCLUSAO_TABELASDEPRECO = 6428 'Parametro: iCodigo
'Erro na tentativa de excluir a Tabela de Preço com código %i da tabela de TabelasDePrecoItens.
Public Const ERRO_FILIALCLIENTE_NAO_CADASTRADA2 = 6430 'Parametros: sNomeRedCliente, iCodFilial
'A Filial %i do Cliente %s não está cadastrada no Banco de Dados.
Public Const ERRO_ATUALIZACAO_FILIALCLIENTEFILEMP = 6431 'Sem Parametros
'Erro na Atualização da tabela FilialClienteFilEmp.
Public Const ERRO_FILIALCLIENTEFILEMP_NAO_CADASTRADA = 6432 'Parametros iFilialEmpresa,lCodCliente,iCodFilial
'O cliente %l, Filial %i com a filialEmpresa %i não estão cadastrado na tabela FilialClienteFilEmp.
Public Const ERRO_LEITURA_FILIALCLIENTEFILEMP = 6433 'Sem Parametro
'Erro na leitura da Tabela FilialClienteFilEmp.
Public Const ERRO_LOCK_FILIALCLIENTEFILEMP = 6434 'SEM pARAMETROS
'Erro na tentativa de fazer lock na tabela FilialClienteFilEmp.
Public Const ERRO_LEITURA_PARCELASREC_TITULOSREC_SALDO = 6435 'Parametros : lCliente
'Erro na leitura das tabelas TitulosRec e ParcelasRec para conseguir a soma dos saldos do Cliente %l.
Public Const ERRO_LEITURA_ITEMPEDIDO_PEDIDOVENDA = 6436 'Sem PArametros
'Erro na Leitura das Tabelas ItensPedidoDeVenda e PedidoDeVenda.
Public Const ERRO_LEITURA_PEDIDOS_VENDA_BLOQUEIOSPV = 6437 'Sem Parametros
'Erro na Leitura das Tabelas BloqueiosPV e PedidoDeVenda.
Public Const ERRO_LEITURA_TITULOSREC3 = 6438  'Parametros: lCodCliente ,giFilialEmpresa
'Erro na leitura da Tabela de TitulosRec com o cliente:%l e FilialEmpresa:%i.
Public Const ERRO_CRFATCONFIG_INEXISTENTE = 6439 '%s chave %d FilialEmpresa
'Não foi encontrado registro em CRFATConfig. Codigo = %s Filial = %i
Public Const ERRO_LEITURA_CRFATCONFIG2 = 6440 '%s chave %d FilialEmpresa
'Erro na leitura da tabela CRFATConfig. Codigo = %s Filial = %i
Public Const ERRO_LEITURA_CRFATCONFIG1 = 6441 'Parametros: sCodigo, iFilial.
'Ocorreu um erro na leitura do codigo %s da Filial %i na tabela CRFATConfig.
Public Const ERRO_DATA_EMISSAO_NAO_PREENCHIDA = 6442 'Sem parâmetros
'É obrigatório o preenchimento da Data de Emissão.
Public Const ERRO_ATUALIZACAO_FILIALFORNFILEMP = 6443 'Sem Parametros
'Erro na Atualização da tabela FilialFornFilEmp.
Public Const ERRO_FILIALFORNFILEMP_NAO_CADASTRADA = 6444 'Parametros iFilialEmpresa,lCodFornecedor,iCodFilial
'O Fornecedor %l, Filial %i com a filialEmpresa %i não estão cadastrado na tabela FilialFornFilEmp.
Public Const ERRO_LEITURA_FILIALFORNFILEMP = 6445 'Sem Parametro
'Erro na leitura da Tabela FilialFornecedorFilEmp.
Public Const ERRO_LOCK_FILIALFORNFILEMP = 6446 'Sem Parametros
'Erro na tentativa de fazer lock na tabela FilialFornFilEmp.
Public Const ERRO_INSERIR_FILIALFORNFILEMP = 6447 'Sem Parametros
'Erro na tentativa de inserir na tabela FilialFornFilEmp.
'Public Const ERRO_TABELAPRECO_UTILIZADA_TIPOCLIENTE = 6448 'Parâmetro: iCodTabela
'A Tabela %i não pode ser excluida pois está sendo utilizada em Tipos de Cliente.
Public Const ERRO_ALCADA_NAO_CADASTRADA2 = 6449 'Parametro: Código do Usuário
'A alçada do usuário %s não está cadastrada.
Public Const ERRO_TIPO_CLIENTE_NAO_ENCONTRADO2 = 6450 'Parametro sTipoCliente
'O tipo de cliente %s não foi encontrado.
Public Const ERRO_TIPO_FORNECEDOR_NAO_ENCONTRADO2 = 6451 'Parametro sTipoFornecedor
'O tipo de fornecedor %s não foi encontrado.
Public Const ERRO_TIPO_VENDEDOR_NAO_ENCONTRADO2 = 6452 'Parametro sTipoVendedor
'O tipo de vendedor %s não foi encontrado.
Public Const ERRO_CONTA_CORRENTE_NAO_ENCONTRADA = 6453 'Parametro: sContaCorrente
'A conta corrente %s não foi encontrada.
Public Const ERRO_TRANSP_NOME_RED_DUPLICADO = 6454 'Parametro: iCodigo
'A transportadora %i tem o mesmo Nome Reduzido.
Public Const ERRO_TRANSPORTADORA_TIPOCLIENTE = 6455 'Parametro: iCodigoTipo
'Erro a Transportadora %i é utilizada na tabela de Tipo de Cliente.
Public Const ERRO_COND_PAGAMENTO_TITULOSPAG = 6456 'Parametros: iCodigo
'Erro a Condição de Pagamento %i é utilizada na tabela TitulosPag.
Public Const ERRO_COND_PAGAMENTO_TITULOSPAG_BAIXADOS = 6457 'Parametros: iCodigo
'Erro a Condição de Pagamento %i é utilizada na tabela TitulosPagBaixados.
Public Const ERRO_CATEGORIACLIENTE_PADROESTRIBSAIDA = 6458  'Parametros: sCategoriaCliente
'Erro a Categoria Cliente %s foi utilizada na tabela de PadroesTribSaida.
Public Const ERRO_CATEGORIACLIENTEITEM_PADROESTRIBSAIDA = 6459  'Parametros: sCategoria , sItemCategoria
'Categoria Cliente %s e Categoria Cliente Item %s são usados na tabela de PadroesTribSaida.
Public Const ERRO_FILIALCLIENTE_CHEQUEPRE = 6460 'Parametros: iFilialCliente,lCodigoCliente
'Erro na Exclusão da Filial %i do Cliente %l que está associado com Cheque Pré.
Public Const ERRO_COND_PAGAMENTO_TITULOSREC = 6461 'Parametros: iCodigo
'Erro a Condição de Pagamento %i é utilizada na tabela TitulosRec.
Public Const ERRO_COND_PAGAMENTO_TITULOSREC_BAIXADOS = 6462 'Parametros: iCodigo
'Erro a Condição de Pagamento %i é utilizada na tabela TitulosRecBaixados.
Public Const ERRO_TABELAPRECO_RELACIONADA_CLIENTE = 6463 'Parametros iTabelaPreco, lCodigoCliente
'Não é possível excluir a Tabela de Preço %i pois está sendo utilizada pelo Cliente %l.
Public Const ERRO_TABELAPRECO_RELACIONADA_NFISCAL = 6464 'Parametros iTabelaPreco, sSerie, lNumeroNF, iFilialNF
'Não é possível excluir a Tabela de Preço %i pois está sendo utilizada pela Nota Fiscal: Série = %s, Numero = %l e FilialEmpresa = %i.
Public Const ERRO_TABELAPRECO_RELACIONADA_TIPOSDECLIENTE = 6465 'Parametros iTabelaPreco, iCodigoTipo
'Não é possível excluir a Tabela de Preço %i pois está sendo utilizada por Tipo de Cliente: Codigo = %i.
Public Const ERRO_TABELAPRECO_RELACIONADA_PEDVENDA = 6466 'Parametros iTabelaPreco, lCodigoPV, iFilialEmpresaPV
'Não é possível excluir a Tabela de Preço %i pois está sendo utilizada pelo Pedido de Venda: Codigo = %l e FilialEmpresa = %i.
Public Const ERRO_TABELAPRECO_RELACIONADA_PEDVENDA_BAIXADO = 6467 'Parametros iTabelaPreco, lCodigoPVBaixado, iFilialEmpresaPVBaixado
'Não é possível excluir a Tabela de Preço %i pois está sendo utilizada pelo Pedido de Venda Baixado: Codigo = %l e FilialEmpresa = %i.
Public Const ERRO_COMISSOES_BAIXADA_NFISCAL = 6468 'Parametros: sSerie, lNumNota, iFilialEmpresa
'Erro na tentativa de excluir registro da tabela de Comissões, da Nota: Série = %s, Número = %l e FilialEmpresa %i. A comissão já foi baixada.
Public Const ERRO_COMISSOES_BAIXADA_PARCELA = 6469 'Parametros: lNumeroTítulo, iParcela, iFilialEmpresa
'Erro na tentativa de excluir registro da tabela de Comissões, da Parcela: Número do Título = %l, Número da Parcela = %i e FilialEmpresa %i. A comissão já foi baixada.
Public Const ERRO_COMISSOES_BAIXADA_DEBITOS = 6470  'Parametros : lNumTitulo, sSiglaDocumento, lCliente,iFilial
'Erro na tentativa de excluir registro da tabela de Comissões, do Débito: Número do Título = %l, Sigla do Documento = %s, Código da Cliente = %l Filial do Cliente = %i. A comissão já foi baixada.
Public Const ERRO_COMISSOES_BAIXADA_TITULO = 6471  'Parametros: lNumTitulo, iFilialEmpresa
'Erro na tentativa de excluir registro da tabela de Comissões, do Título : Número do Título = %l e FilialEmpresa %i. A comissão já foi baixada.
Public Const ERRO_NOTAFISCAL_NAO_CADASTRADO_COMISSOES = 6472 'Sem Parametros
'A Nota Fiscal associada a esta comissão não foi encontrada.
Public Const ERRO_PARCELAREC_NAO_CADASTRADO_COMISSOES = 6473 'Sem Parametros
'A Parcela associada a esta Comissão não foi encontrada.
Public Const ERRO_TITULOREC_NAO_CADASTRADO_COMISSOES = 6474 'Sem Parametros
'O Titulo a Receber associado a esta Comissão não foi encontrado.
Public Const ERRO_DEBITOREC_NAO_CADASTRADO_COMISSOES = 6475 'Sem Parametros
'O Débito a Receber associado a esta Comissão não foi encontrado.
Public Const ERRO_TRANSF_MANUAL_COBR_ELETRONICA = 6476 'sem parametro
'Não pode transferir título de/para cobrador com cobrança eletrônica
Public Const ERRO_ALTERE_VCTO_INSTR_COB_ELETR = 6477 'sem parametros
'Para alterar o vencimento de uma parcela em cobrança eletrônica use a tela de instruções para cobrança eletrônica
Public Const ERRO_LEITURA_CHEQUEPRE3 = 6478 'sem parametros
'Erro na leitura da tabela de cheques pré-datados
Public Const ERRO_CHEQUEPRE_DEPOSITADO = 6479 'sem parametros
'Este cheque pré-datado já foi depositado
Public Const ERRO_PARCREC_DE_CARTEIRA_PARA_CHEQUEPRE = 6480
'Para associar um cheque pré datado a uma parcela esta deve estar em carteira.
Public Const ERRO_PARCELAREC_OUTRO_CHEQUEPRE = 6481 'parametros: banco ag, cta e num do cheque
'Esta parcela já está associada ao cheque pré identificado por: banco %s, agencia %s, conta %s e número %s. Pode-se excluir o cheque anterior e cadastrar um novo.
Public Const ERRO_CHEQUEPRE_REPETIDO = 6482 'sem parametros
'O mesmo cheque pré-datado está associado a mais de uma parcela
Public Const ERRO_CHEQUEPRE_DUPLICADO = 6483 'parametro: codigo de cliente
'Este cheque pré-datado já está registrado para o cliente com código %s.
Public Const ERRO_CHEQUEPRE_OUTROTITULO = 6484 'sem parametros
'Este cheque pré-datado está associado a outro título já registrado
Public Const ERRO_PARCELA_COBRANCA_EMPRESA = 6485 'Parâmetro: lNumTitulo, iNumParcela
'Erro na exclusão do Titulo %l. A parcela %i não está em cobrança na própria Empresa.
Public Const ERRO_ALTERACAO_CARTEIRA_EMPRESA = 6486 'sem parametros
'Não se pode incluir ou alterar uma carteira de cobrança de uso restrito à própria Empresa.
Public Const ERRO_NFISCALINTERNA_COM_NUMERO = 6487 'Sem Parametros
'Não é possivel gravar uma Nota Fiscal Interna com seu número preenchido.
'SOLUCAO: Limpe o campo Número com o botão ao lado, pois o sistema irá gerar o Número na gravação.
Public Const ERRO_NFISCAL_COM_NUMERO = 6488 'Sem Parametros
'Não é possivel gravar uma Nota Fiscal com seu número preenchido.
'SOLUCAO: Limpe o campo Número com o botão ao lado, pois o sistema irá gerar o Número na gravação.
Public Const ERRO_NOTA_FISCAL_CANCELADA = 6489 'Parâmetro: sSerieNF, lNumNotaFiscal
'A Nota Fiscal com a série %s e número %l já está cancelada.
Public Const ERRO_NFISCAL_VINCULADA_CANCELAR = 6490 'Parâmetro: lNumNotaFiscalOrig, iTipoNFOrig
'A Nota Fiscal em questão não pode ser cancelada pois está vinculada a
'Nota Fiscal número %l e tipo %i que não está cancelada.
Public Const ERRO_ALTERACAO_ITEMNF = 6491
'Erro na tentativa de alterar os dados de um item de nota fiscal.
Public Const ERRO_NF_JA_DEVOLVIDA = 6492
'Ja existe uma nota fiscal de devolução para a nota em questão.
Public Const ERRO_NFISCAL_OUTRA_FILIAL = 6493
'A nota fiscal em questão percente a outra filialempresa.
Public Const ERRO_NFISCAL_SEM_ITENS = 6494 'lCodPedido
'A Nota Fiscal com o Código %l não possui Ítens associados à ele.
Public Const ERRO_REGIAO_INICIAL_MAIOR = 6495 'Sem Parametros
'A Região inicial não pode ser maior que a Região final.
Public Const ERRO_COBRADOR_INATIVO = 6496 'Parâmetros: iCodCobrador
'O Cobrador de código %i é inativo.
Public Const ERRO_PADRAO_COBRANCA_INATIVO = 6497 'Parâmetros: iCodPadraoCobranca
'O Padrão de Cobrança de código %i é inativo.
Public Const ERRO_CRIACAO_NFR_COM_FATURAMENTO = 6498
'A criação de um título de Nota Fatura a Receber ou Nota Fatura a Receber de Serviço é criado automaticamente após o cadastro da Nota Fiscal.
Public Const ERRO_LEITURA_TRANSPORTADORA2 = 6499 'Sem Parametros
'Erro na leitura da tabela de Transportadoras.
Public Const ERRO_LIMITE_TRANSP_VLIGHT = 6500 'Parametros : iNumeroMaxTransportadoras
'Número máximo de Transportadoras desta versão é %i.
Public Const ERRO_FATURA_ATE_IMPRESSAO_NAO_PREENCHIDO = 6501 'Sem Parametros
'É obrigatório informar até que Fatura foi obtida boa Impressão.
Public Const ERRO_FATURA_ATE_MAIOR_ANTERIOR = 6502 'Sem Parametros
'A Fatura Até deve ser menor ou igual ao que foi mandado para Impressão.
Public Const ERRO_FATURA_ATE_MENOR_NUMERO_DE = 6503 'Sem Parametros
'A Fatura Até não pode ser menor que a Fatura De.
Public Const ERRO_NUMERO_ATE_IMPRESSAO_NAO_PREENCHIDO = 6504 'Sem Parametros
'É obrigatório informar até que Nota Fiscal foi obtida boa Impressão.
Public Const ERRO_UNLOCK_SERIE_IMPRESSAO_NF = 6505 'Parametros : sSerie
'A Série %s não está lockada para Impressão, por isso não pode ser feito unlock.
Public Const ERRO_NUMERO_ATE_MAIOR_ANTERIOR = 6506 'Sem Parametros
'O Número Até deve ser menor ou igual ao que foi mandado para Impressão.
Public Const ERRO_NUMERO_ATE_MENOR_NUMERO_DE = 6507 'Sem Parametros
'O Numero Até não pode ser menor que o Número De.
Public Const ERRO_TIPO_FORMULARIO_IMCOMPATIVEL = 6508 'Parametro: sSerie
'O Tipo de Formulário da Série %s está imcompativel.
Public Const ERRO_ATUALIZACAO_NFISCAL1 = 6509
'Erro na atualização de registros na tabela de notas fiscais.
Public Const ERRO_LEITURA_ANTECIPRECS = 6510  'sem parametros
'Erro na leitura de adiantamentos de clientes
Public Const ERRO_FALTA_MESANO_ESTOQUE = 6511 'sem parametros
'Preencha o mes/ano para o módulo de estoque.



'VEIO DE ERROS CPR
Public Const ERRO_LEITURA_FILIAISFORNECEDORES = 2026
'Erro na leitura da tabela Filiais Fornecedores.
Public Const ERRO_LOCK_FILIAISFORNECEDORES = 2040 'Parametro Codigo Fornecedor
'Erro na tentativa de fazer "lock" na tabela de FiliaisFornecedores para Codigo Fornecedor = %l .
Public Const ERRO_LEITURA_TITULOS_PAGAR = 2042
'Erro na leitura da tabela Titulos a Pagar.
Public Const ERRO_FILIAL_FORNECEDOR_INEXISTENTE = 2047 'Parametro Codigo Filial Fornecedor e Codigo Fornecedor
'A Filial Fornecedor %s do Fornecedor %s , não esta cadastrada no Banco De Dados.
Public Const ERRO_TIPOCLIENTE_INEXISTENTE = 2099
'Esse Tipo de Cliente não existe , ou então já foi excluído.
Public Const ERRO_TIPOCLIENTE_DESCR_DUPLICATA = 2102
'A Descrição %s já esta sendo utilizada para outro Tipo De Cliente.
Public Const ERRO_EXCLUSAO_TIPOCLIENTE_RELACIONADO = 2105 'Parametro Codigo do Cliente
'Esse Tipo Cliente não pode ser excluído pois se encontra relacionado com o Cliente %l.
Public Const ERRO_LOCK_TIPOCLIENTE = 2106 'Parametro Codigo do Tipo de Cliente
'Erro na tentativa de fazer "lock" no Tipo de Cliente = %i na tabela TiposDeCliente.
Public Const ERRO_EXCLUSAO_TIPOCLIENTE = 2107 'Parametro Codigo do Tipo de Cliente
'Erro na tentativa de excluir o Tipo De Cliente = %i.
Public Const ERRO_LEITURA_PARCELAS_PAG1 = 2229
'Erro na leitura da tabela de Parcelas a Pagar e Titulos a Pagar.
Public Const ERRO_LEITURA_PARCELAS_REC = 2234
'Erro na leitura da tabela de Parcelas a Receber.
Public Const ERRO_LEITURA_FILIAISFORNECEDORES1 = 2292 'Parâmetros : lFornecedor, iFilial
'Erro na leitura do Fornecedor %l e Filial %i na tabela FiliaisFornecedores.
Public Const ERRO_LOCK_FILIAISFORNECEDORES1 = 2294 'Parametros: lFornecedor e iFilial
'Erro na tentativa de fazer "lock" na tabela de FiliaisFornecedores para Fornecedor = %l  e Filial = %i
Public Const ERRO_LEITURA_NFSPAG1 = 2299 'Parametro: lNumNotafiscal
'Erro na tentativa de leitura da Nota Fiscal número %l na tabela NFsPag.
Public Const ERRO_ATUALIZACAO_NFSPAG = 2306 'Parâmetro: lNumnotafiscal
'Erro na tentativa de atualização da Nota Fiscal número %l na tabela de NfsPag.
Public Const ERRO_LOCK_NFSPAG = 2316 'Parametro : lNumnotaFiscal
' Erro na tentativa de "lock" da tabela NfsPag na Nota Fiscal número %l.
Public Const ERRO_NF_VINCULADA = 2317 'Parâmetros: lNumNotafiscal, lNumTitulo
'Não é possível excluir a Nota Fiscal número %l pois está vinculada à Fatura %l.
Public Const ERRO_VENDEDOR_NAO_CADASTRADO1 = 2383 'Parametro: sNomeReduzido
'O Vendedor com Nome Reduzido %s não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_TITULOS_PAGAR_BAIXADO = 2466
'Erro na leitura da tabela de Títulos a Pagar Baixados.
Public Const ERRO_NUMBORDERO_CHEQUEPRE_DEPOSITADO = 2478 'Sem parametros
'Não se pode excluir um cheque pré-datado que tenha sido depositado
Public Const ERRO_LEITURA_CHEQUEPRE1 = 2480 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'Erro na leitura da tabela de ChequePre com Banco %i, Agência %s, ContaCorrente %s e Número %l.
Public Const ERRO_CHEQUEPRE_JA_UTILIZADO = 2490 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'O ChequePre com Banco %i, Agência %s, ContaCorrente %s e Número %l já foi utilizado.
Public Const ERRO_LOCK_CHEQUESPRE = 2491 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'Erro na tentativa de fazer "lock" na tabela de ChequesPre com Banco %i, Agência %s, ContaCorrente %s e Número %l.
Public Const ERRO_REFERENCIA_OUTRO_CHEQUE = 2492 'Parametro: lNumIntParc
'Esta Parcela %l faz referência a outro Cheque Pre.
Public Const ERRO_CHEQUEPRE_NAO_ENCONTRADO = 2498 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'O ChequePre com Banco %i, Agência %s, ContaCorrente %s e Número %l não foi encontrado na tabela de ChequesPre.
Public Const ERRO_CLINTE_FILIAL_NAO_CONFEREM = 2504 'Parametro: lCliente, iFilial
'O Cliente %l ou a Filial %i não conferem com o BD.
Public Const ERRO_PARCELAS_VINCULADAS_CHQPRE = 2505 'Parametro: lNumIntParc
'Não existe Parcelas associadas ao Cheque Pre %l.
Public Const ERRO_VENDEDOR_JA_EXISTENTE = 2508 'Parametro: sNomeReduzido
'O Vendedor %s já existe no Grid de Comissões. Duas Comissões não podem ter o mesmo Vendedor.
Public Const ERRO_VENDEDOR_COMISSAO_NAO_INFORMADO = 2509 'iComissao
'O Vendedor da comissao %i do Título não foi informado.
Public Const ERRO_LEITURA_CREDITOSPAGFORN = 2587 'Parâmetros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Erro na leitura  da tabela de CreditosPagForn com Fornecedor l%, Filial i%, Tipo de Documento %s, Número l% e Data de Emissão %dt.
Public Const ERRO_LEITURA_TIPODOCUMENTO = 2589 'Parâmetro: sSigla
'Erro na leitura da tabela de TiposDeDocumento com o Tipo de Documento com Sigla %s.
Public Const ERRO_TIPO_NAO_PREENCHIDO = 2592 'Sem parâmetros
'Preenchimento do Tipo de Documento é obrigatório.
Public Const ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO = 2593 'Parâmetro: sSiglaTipoDocumento
'O Tipo de Documento %s não está cadastrado no Banco de Dados.
Public Const ERRO_INSERCAO_CREDITOSPAGFORN = 2596 'Pârametros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo
'Erro na tentativa de inserir um novo registro na tabela CreditosPagForn com Fornecedor %l, Filial %i Tipo de Documento %s e Número %l.
Public Const ERRO_CLIENTE_NAO_PREENCHIDO = 2676 'Sem parametro
'O preenchimento de Cliente é obrigatório.
Public Const ERRO_INSERCAO_DEBITOSRECCLI = 2678 'Parametro: lNumProxDebitoRecCli
'Erro na tentativa de inserir um novo registro número %l na tabela de DebitosRecCli.
Public Const ERRO_LOCK_VENDEDOR1 = 2687 'Parâmetro: iVendedor
'Não foi possível fazer o Lock do Vendedor %i da tabela Vendedores.
Public Const ERRO_LEITURA_TIPOSDEDOCUMENTO = 2688
'Erro na leitura da tabela de Tipo de Documento.
Public Const ERRO_PARCELA_VINCULADA_CHQPRE = 2760
'A parcela a receber selecionada já está vinculada a outro cheque-pré.
Public Const ERRO_PARCELA_VINCULADA_CHEQUEPRE = 2762 'Parâmetros: iNumParcela, lNumIntCheque
'A Parcela com número interno %i já está associada ao Cheque-Pré com número interno %l.
Public Const ERRO_LEITURA_COBRADOR1 = 2763 'Parâmetro: sNomeReduzido
'Erro na leitura da tabela de Cobrador com Nome Reduzido %s.
Public Const ERRO_LEITURA_PARCELAS_REC_BAIXADAS = 2862
'Erro na leitura da tabela de Parcelas a Receber Baixadas.
Public Const ERRO_PARCELA_RECEBER_NAO_CADASTRADA1 = 2890 'Parâmetro: lNumIntParc
'A Parcela a Receber com o Número Interno %l não foi encontrada no Banco de Dados
Public Const ERRO_VENDEDOR_INICIAL_MAIOR = 2916 'Sem parametros
'Código do Vendedor Inicial é maior
Public Const ERRO_CLIENTE_INICIAL_MAIOR = 2938
'O Cliente Inicial é maior que o final.'
Public Const ERRO_CLIENTE_NAO_CADASTRADO_2 = 2939
'O Cliente não está cadastrado.'
Public Const AVISO_CREDITO_PAGAR_NUMERO_REPETIDO = 5285 'Parâmetros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo e dtDataEmissao
'No Banco de Dados já existe Crédito com Fornecedor com os dados abaixo: Número: %l, Fornecedor: %l, Filial: %i, Tipo: %s e Data de Emissão: %dt. Deseja prosseguir na inserção de um novo Crédito com mesmo Número e Fornecedor?
Public Const AVISO_DEBITORECCLI_JA_EXISTENTE = 5287 'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo e dtDataEmissao
'No Banco de Dados já existe Débito com Cliente com os dados abaixo: Número: %l, Cliente: %l, Filial: %i, Tipo: %s e Data de Emissao: dt%. Deseja prosseguir na inserção de novo Débito do mesmo tipo com o mesmo número ?
Public Const ERRO_CODIGO_INVALIDO1 = 16010 'Sem Parametros
'O Código tem que ser um valor inteiro positivo.



''VEIO DE ERROS FAT
Public Const ERRO_CODIGO_INVALIDO = 8009 'Parametro: Codigo.Text
'O preenchimento de Código deve ser um número maior que 99.
Public Const ERRO_SERIE_NAO_CADASTRADA = 8010 'Parametro: sSerie
'A Serie %s não está cadastrada no Banco de Dados.
Public Const ERRO_LEITURA_SERIE = 8011 'Sem Parametro
'Erro de leitura no Banco de Dados.
Public Const ERRO_LEITURA_CANALVENDA1 = 8048 'Parametros objCanal.iCodigo
'Erro na leitura do canal %i da tabela CanalVenda
Public Const ERRO_LIMITEMENSAL_NAO_INFORMADO = 8107
'O Limite Mensal deve ser informado.
Public Const ERRO_LIMITEOPERACAO_NAO_INFORMADO = 8108
'O Limite de Operação deve ser informado.
Public Const ERRO_ALCADA_NAO_CADASTRADA = 8110 'Parametro NomeReduzido
'A alçada do usuário %s não está cadastrada.
Public Const ERRO_DATADE_MAIOR_DATAATE = 8129 'Sem parâmetro
'"Data De" deve ser menor que "Data Até".
Public Const ERRO_LEITURA_PEDIDODEVENDA = 8138
'Erro na leitura da tabela de PedidoDeVenda.
Public Const ERRO_LEITURA_PEDIDODEVENDABAIXADOS = 8139
'Erro na leitura da tabela de PedidoDeVendaBaixados.
Public Const ERRO_LEITURA_PADRAOTRIBSAIDA = 8140
'Erro na leitura da tabela de PadraoTribSaida.
Public Const ERRO_LEITURA_PADRAOTRIBENTRADA = 8141
'Erro na leitura da tabela de PadraoTribEntrada.
Public Const ERRO_LEITURA_TRIBUTACAONF = 8142
'Erro na leitura da tabela de TributacaoNF.
Public Const ERRO_LEITURA_TRIBUTACAOITEMPV = 8143
'Erro na leitura da tabela de TributacaoItemPV.
Public Const ERRO_LEITURA_TRIBUTACAOCOMPLNF = 8144
'Erro na leitura da tabela de TributacaoComplNF.
Public Const ERRO_LEITURA_TRIBUTACAOCOMPLPV = 8145
'Erro na leitura da tabela de TributacaoComplPV.
Public Const ERRO_LEITURA_TRIBUTACAOITEMNF = 8146
'Erro na leitura da tabela de TributacaoItemNF.
Public Const ERRO_NATUREZAOP_USADO_PEDIDODEVENDA = 8147 'Parametro : sCodigo da Natureza
'Não é permitido excluir a Natureza de Operação %s, pois está vinculada com Pedido de Venda %l da Filial Empresa %i.
Public Const ERRO_NATUREZAOP_USADO_PEDIDODEVENDABAIXADO = 8148 'Parametro : sCodigo da Natureza
'Não é permitido excluir a Natureza de Operação %s, pois está vinculada com Pedido de Venda Baixado %l da Filial Empresa %i.
Public Const ERRO_NATUREZAOP_USADO_NFISCAL = 8149 'Parametro : sCodigo da Natureza
'Não é permitido excluir a Natureza de Operação %s, pois está vinculada com a Nota Fiscal - Série = %s, Número = %l e Filial Empresa = %l.
Public Const ERRO_NATUREZAOP_USADO_NFISCALBAIXADA = 8150 'Parametro : sCodigo da Natureza
'Erro Natureza %s está sendo utilizada em uma Nota Fiscal baixada.
Public Const ERRO_NATUREZAOP_USADO_PADRAOTRIBSAIDA = 8151 'Parametro : sCodigo da Natureza
'Erro Natureza %s está sendo utilizada em um Padrão Tributação Saida.
Public Const ERRO_NATUREZAOP_USADO_PADRAOTRIBENTRADA = 8152 'Parametro : sCodigo da Natureza
'Erro Natureza %s está sendo utilizada em um Padrão Tributação Entrada.
Public Const ERRO_NATUREZAOP_USADO_TIPODOCINFO = 8153 'Parametro : sCodigo da Natureza
'Não é permitido excluir a Natureza de Operação %s, pois está vinculada com Tipo de Documento (Sigla = %s).
Public Const ERRO_NATUREZAOP_USADO_TRIBUTACAONF = 8154 'Parametro : sCodigo da Natureza
'Erro Natureza %s está sendo utilizada em uma Tributação N.F.
Public Const ERRO_NATUREZAOP_USADO_TRIBUTACAOITEMPV = 8155 'Parametro : sCodigo da Natureza
'Erro Natureza %s está sendo utilizada em uma Tributação Item P.V.
Public Const ERRO_NATUREZAOP_USADO_TRIBUTACAOCOMPLNF = 8156 'Parametro : sCodigo da Natureza
'Erro Natureza %s está sendo utilizada em uma Tributação Complemento N.F..
Public Const ERRO_NATUREZAOP_USADO_TRIBUTACAOCOMPLPV = 8157 'Parametro : sCodigo da Natureza
'Erro Natureza %s está sendo utilizada em uma Tributação Complemento P.V..
Public Const ERRO_NATUREZAOP_USADO_TRIBUTACAOITEMNF = 8158 'Parametro : sCodigo da Natureza
'Erro Natureza %s está sendo utilizada em uma Tributação  Item N.F.
Public Const ERRO_NOTA_FISCAL_INTERNA_ENTRADA_NAO_CADASTRADA = 8160 'Parametros sSerie, lNumNotaFiscal
'Nota Fiscal Interna de Entrada com série %s e número %l não está cadastrada no Banco de Dados.
Public Const ERRO_ATUALIZACAO_TABELASDEPRECO = 8191 'Parametros: iCodigo
'Erro na atualização de registro na tabela TabelasDePreco com código da Tabela %i.
Public Const ERRO_DESCRICAO_TABELAPRECO_JA_EXISTENTE = 8192 'Parâmetro: Descrição
'A Descrição %s já é utilizada por outra Tabela de Preço.
Public Const ERRO_TABELAPRECO_JA_EXISTENTE = 8193 'Parametro: iCodigo
'A Tabela de Preço com o código %i já existe no Banco de Dados.
Public Const ERRO_INSERCAO_TABELASDEPRECO = 8194 'Parametros: iCodigo
'Erro na inserção de registro na tabela TabelasDePreco com código da Tabela %i.
Public Const ERRO_AUSENCIA_PEDIDO_BAIXAR = 8196 'Sem parâmetros
'Deve haver pelo menos um Pedido marcado para ser baixado.
Public Const ERRO_BLOQUEIOPV_REPETIDO = 8254 'iTipoBloqueio
'Já existe no grid um bloqueio com o tipo %i.
Public Const ERRO_ITEM_ARVORE_CONSULTA_NAO_SELECIONADO = 8296 'Parametros
'É necessário selecionar um ítem na árvore de Consultas.
Public Const ERRO_LEITURA_CONSULTAS = 8297 'Sem parâmetros
'Erro na leitura da tabela Consultas.



''VEIO DE ERROS MAT
Public Const ERRO_LEITURA_NFISCAL4 = 7639 'Sem Parâmetros
'Erro na leitura da tabela de Notas Fiscais
Public Const ERRO_ATUALIZACAO_SERIE = 7640 'Parametro: sSerie
'Erro na atualização da Série %s na tabela de Séries.
Public Const ERRO_VALOR_TOTAL_COMISSAO_INVALIDO = 7907 'Parâmetro: dValorTotalComissao, dValorTotal
'O total de valores de comissões = %d não pode ultrapassar o Valor Total = %d.
Public Const ERRO_VENDEDOR_COMISSAO_GRID_NAO_INFORMADO = 7908 'Parâmetro: iComissao
'O Vendedor da comissão %i do Grid de Comissões não está preenchido.


''VEIO DE ERROS TRB
Public Const ERRO_TIPO_TRIBUTACAO_NAO_CADASTRADO = 7013 'Parâmetro: iTipo
'O Tipo %i de Tributação não está cadastrado no Banco de Dados.


'Veio de ErrosCOM
Public Const ERRO_REGISTRO_COMPRAS_CONFIG_NAO_ENCONTRADO = 12004 'Parametros sCodigo,iFilialEmpresa
'Registro na tabela ComprasConfig com Código=%s e FilialEmpresa=%i não foi encontrado.
Public Const ERRO_LEITURA_PEDIDOCOMPRATODOS = 12277 'Sem parâmetros
'Erro na leitura da tabela PedidoCompraTodos.
Public Const ERRO_LEITURA_PEDIDOCOTACAOTODOS = 12361 'Sem parâmetros
'Erro na leitura de PedidoCotacaoTodos.




'Códigos de Avisos - Reservado de 5400 até 5499
Public Const AVISO_CRIAR_VENDEDOR = 5400
'Deseja cadastrar novo Vendedor?
Public Const AVISO_CRIAR_VENDEDOR1 = 5401 'Parametro: iCodigo
'Vendedor com código %i não está cadastrado. Deseja cadastrar?
Public Const AVISO_CRIAR_VENDEDOR2 = 5402 'Parametro: sNomeReduzido
'Vendedor com Nome Reduzido %s não está cadastrado. Deseja cadastrar?
Public Const AVISO_CRIAR_CLIENTE = 5403
'Deseja cadastrar novo Cliente?
Public Const AVISO_CRIAR_CLIENTE_1 = 5404 'Parametro: sNomeReduzido
'Cliente %s não está cadastrado. Deseja criar?
Public Const AVISO_CRIAR_CLIENTE_2 = 5405 'Parametro: lCodigo
'Cliente com código %s não está cadastrado. Deseja criar?
Public Const AVISO_CRIAR_CLIENTE_3 = 5406 'Parametro: sCGC
'Cliente com CGC/CPF %s não está cadastrado. Deseja criar?
Public Const AVISO_EXCLUIR_CATEGORIACLIENTE = 5407 'ParÂmetro: sCategoria
'Confirma exclusão da Categoria %s?
Public Const AVISO_DESEJA_CRIAR_CATEGORIACLIENTE = 5408 'Sem parâmetro
'Confirma a criação de uma nova Categoria de Cliente?
Public Const AVISO_DESEJA_CRIAR_CATEGORIACLIENTEITEM = 5409 'Sem parâmetros
'Confirma a criação de um novo Ítem de Categoria de Cliente?
Public Const AVISO_CONFIRMA_EXCLUSAO_CONDICAOPAGTO = 5410 'Parâmetro: iCodigo
'A Condição de Pagamento %i será excluída. Confirma exclusão?
Public Const AVISO_CONFIRMA_EXCLUSAO_TIPODEVENDEDOR = 5411 'Parâmetro: iCodigo
'O Tipo de Vendedor %i será excluído. Confirma exclusão?
Public Const AVISO_CONFIRMA_EXCLUSAO_VENDEDOR = 5412 'Parametro iCodigo
'O Vendedor %i será excluído. Confirma exclusão?
Public Const AVISO_CRIAR_TIPO_VENDEDOR = 5413
'Deseja criar um novo Tipo de Vendedor?
Public Const AVISO_COMISSAO_EMISSAO_PAGA = 5414 'Parametro: tTituloReceber.lNumIntDoc
'O Número de Comissões na Emissão do Título com número %l não pode ser alterado por que já existe comissao paga.
Public Const AVISO_DATAVENCIMENTO_PARCELAS_ALTERAVEIS = 5415
'Este Título já está lançado, portanto só é permitido alterar os campos referentes a Parcelas (Datas de Vencimento), Descontos, Cheque-Pré e Comissões. Deseja prosseguir?
Public Const AVISO_EXCLUSAO_TITULORECEBER = 5416 'Parametro : lNumTitulo
'Confirma a exclusão do Título a Receber número %l ?
Public Const AVISO_PARCELA_COM_BAIXA_NAO_ALTERAVEL = 5417 'Parametro: iParcela
'A Parcela %i com baixa não pode ter o número de descontos alterado. Deseja prosseguir na alteração para os campos alteráveis?
Public Const AVISO_PARCELA_COM_BAIXA_DESCONTO_INALTERAVEL = 5418 'iParcela
'A Parcela %i com baixa não pode ter o descontos alterados. Deseja prosseguir na alteração para os campos alteráveis?
Public Const AVISO_COMISSOES_EMISSAO_NAO_ALTERAVEIS = 5419 'lNumIntDoc
'O número de comissões na emissão do Título com Número Interno %l não pode ser alterado porque existe comissão paga. Deseja prosseguir na alteração para os campos alteráveis?
Public Const AVISO_DATA_VALOR_CHEQUEPRE_NAO_ALTERAVEIS = 5420 'iParcela
'O Valor e Data de ChequePre da Parcela %i com baixa não podem ser alterados.Deseja prosseguir na alteração para os campos alteráveis?
Public Const AVISO_NUM_COMISSOES_NAO_ALTERAVEL = 5421 'iParcela
'O Número de Comissões da Parcela %i não pode ser alterado porque existe comissão paga. Deseja prosseguir na alteração para os campos alteráveis?
Public Const AVISO_COMISSAO_PARCELA_PAGA = 5422 'iParcela
's comissões da Parcela %i não poderão ser alteradas porque existe comissão paga. Deseja prosseguir na alteração para os campos alteráveis?
Public Const AVISO_TITULORECEBER_PENDENTE_MESMO_NUMERO = 5423 'sSiglaDocumento, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Titulo a Receber Pendente com os dados Tipo = %s, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de Novo Titulo a Receber com o mesmo número?
Public Const AVISO_TITULORECEBER_MESMO_NUMERO = 5424 'sSiglaDocumento, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Titulo a Receber com os dados Tipo = %s, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de novo Titulo a Receber com o mesmo número?
Public Const AVISO_TITULORECEBER_BAIXADO_MESMO_NUMERO = 5425 'sSiglaDocumento, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Titulo a Receber Baixado com os dados Tipo = %s, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de Novo Titulo a Receber com o mesmo número?
Public Const AVISO_CARTEIRA_JA_ADICIONADA = 5426 'sCarteira
'A Carteira  %s  já está adicionada.
Public Const AVISO_BANCO_INEXISTENTE = 5427 'Parametro objBanco.iCodBanco
'O Banco %i não está cadastrado
Public Const AVISO_EXCLUIR_COBRADOR = 5428 'Parametro objcobrador.icodigo
'Aviso excluir cobrador %i da tabela cobradores
Public Const AVISO_CONFIRMA_EXCLUSAO_CARTEIRACOBRANCA = 5429
'Confirma exclusão de Carteira Cobrança ?
Public Const AVISO_EXCLUIR_TRANSPORTADORA = 5430 'Parametro: iCodigo
'Confirma exclusão da transportadora com código %i?
Public Const AVISO_VALOR_DESCONTO_MAIOR1 = 5431 'dDesconto, dSomaValores
'Desconto = %d não pode ultrapassar a soma de Produtos + ICMSSubst + IPIValor + Frete + Seguro + Despesas = %d. Desconto será zerado.
Public Const AVISO_EXISTENCIA_NOTA_FISCAL_SAIDA = 5432 'Parâmetros: lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'No Banco de Dados existe Nota Fiscal de Saida com os Dados Código do Cliente =%l, Código da Filial =%i, Tipo =%i,  Série NF =%s, Número NF =%l, Data Emisão =%dt.
Public Const AVISO_EXISTENCIA_NOTA_FISCAL_SAIDA_BAIXADA = 5433 'Parâmetro: lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'No Banco de Dados existe Nota Fiscal de Saida Baixada com os Dados Código do Cliente =%l, Código da Filial =%i, Tipo =%i,  Série NF =%s, Número NF =%l, Data Emisão =%dt.
Public Const AVISO_IR_FONTE_MAIOR_VALOR_TOTAL = 5434
'IR Fonte maior que o valor total.
Public Const AVISO_NF_INTERNA_DATA_PROXIMA = 5435 'sSerie, lNumNotaFiscal, dtDataEmissao
'No Banco de Dados existe Nota Fiscal Interna com os dados Série =%s, Número =%l, Data Emissão =%dt. Deseja prosseguir na inserção de nova Nota Fiscal com o mesmo número?
Public Const AVISO_EXCLUSAO_TABELA_DE_PRECO = 5436 'Parametro: iCodigo
'Confirma a exclusão da Tabela de Preço com código %i ? Será excluído a tabela para todas as Filiais.
Public Const AVISO_EXCLUSAO_ITEM_TABELA_DE_PRECO = 5437 'Parametros: sCodProduto, iCodTabela
'Confirma a exclusão de Item com código %s da Tabela de Preço com código %i ?
Public Const AVISO_CLIENTE_CGC_IGUAL = 5438 'Parametro: sCGC
'Já existe um outro Cliente Cadastrado com o CGC %s, deseja continuar a Gravação?
Public Const AVISO_INFORMA_NUMERO_NOTA_GRAVADA = 5439  'Parametros: lNumNotaFiscal
'A Nota Fiscal foi gravada com o Número %l.
Public Const AVISO_INFORMA_NUMERO_FATURA = 5440  'Parametros: lNumFatura
'A Fatura a Receber foi gravada com o Número %l.
Public Const AVISO_FATURA_LOCKADA = 5441 'Sem Parametros
'A Impressão da Fatura a Receber está bloqueada. Está havendo uma impressão ou houve erro anterior.
'Continue somente em caso de Erro anterior. Deseja Continuar?
Public Const AVISO_FATURA_REIMPRESSA = 5442 'Parametros : lNumeroFaturaInicial, lNumFaturaFinal
'Confirma a reimpressão das Faturas a Receber de %l até %l ?
Public Const AVISO_CRIAR_PAIS = 5443 'Parametros: iCodPais
'O País %i não existe, deseja criá-lo?


'VEIO DE ERROS EST
Public Const AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB = 5603 'sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal Interna com os dados Série = %s , Número = %l, Data Emissão = %dt está cadastrada no Banco de Dados, só é possivel alterar os dados relativos a contabilidade. Deseja proseguir na alteração?



'Erros CRFAT 2
'Códigos de Erros - Reservado de 13100 a 13299
Public Const ERRO_LEITURA_CONTRATOFORNECIMENTO = 13100 'Sem parametros
'Erro na leitura da tabela de ContratoFornecimento.
Public Const ERRO_LEITURA_ITENS_CONTRATO = 13101 'Sem parametros
'Erro na leitura da tabela de ItensContrato.
Public Const ERRO_FORNECEDOR_REL_PEDIDOCOMPRA = 13102 'Parâmetros: lCodigoForn, lCodigoPedidoCompra
'O Fornecedor %l está relacionado com o Pedido de Compra %l.
Public Const ERRO_FORNECEDOR_REL_CONTRATOFORNECIMENTO = 13103 'Parâmetros: lCodigoForn, lCodigoContrato
'O Fornecedor %l está relacionado com o Contrato de Fornecimento %l.
Public Const ERRO_FORNECEDOR_REL_ITEMCONCORRENCIA = 13104 'Parâmetros: lCodigoForn
'O Fornecedor %l está relacionado com Item de Concorrência.
Public Const ERRO_FORNECEDOR_REL_CONCORRENCIA = 13105 'Parâmetros: lCodigoForn, lCodigoConcorrencia
'O Fornecedor %l está relacionado com a Concorrência %l.
Public Const ERRO_FORNECEDOR_REL_REQCOMPRA = 13106 'Parametro: lCodFornecedor, lCodRequisicaoCompra
'O Fornecedor %l está relacionado com a Requisição de Compra %l.
Public Const ERRO_FORNECEDOR_REL_PEDIDOCOTACAO = 13107 'Parametro: lCodFornecedor, lCodPedidoCotacao
'O Fornecedor %l está relacionado com o Pedido de Cotação %l.
Public Const ERRO_FORNECEDOR_REL_COTACAO = 13108 'Parametro: lCodFornecedor, lCodCotacao
'O Fornecedor %l está relacionado com a Cotação %l.
Public Const ERRO_FORNECEDOR_REL_ITEMREQCOMPRA = 13109 'Parametro: lCodFornecedor
'O Fornecedor %l está relacionado com um item de Requisição de Compra.
Public Const ERRO_FORNECEDOR_REL_REQMODELO = 13110 'Parametro: lCodFornecedor, lCodRequisicaoModelo
'O Fornecedor %l está relacionado com a Requisição Modelo %l.
Public Const ERRO_FORNECEDOR_REL_COTACAOPRODUTO = 13111 'Parametro: lCodFornecedor
'O Fornecedor %l está relacionado com Cotação Produto.
Public Const ERRO_LEITURA_ITENSREQMODELO2 = 13112 'sem parâmetros
'Erro na leitura da tabela ItensReqModelo.
Public Const ERRO_FORNECEDOR_REL_ITENSREQMODELO = 13113 'Parametro: lCodFornecedor
'O Fornecedor %l está relacionado com um item de Requisição Modelo.
Public Const ERRO_LEITURA_COTACAOPRODUTOTODAS = 13114 'Sem parametros
'Erro na leitura da tabela CotacaoProdutoTodas.
Public Const ERRO_LEITURA_CONCORRENCIATODAS = 13115 'Sem Parâmetros
'Erro de leitura na tabela de ConcorrenciaTodas.
Public Const ERRO_FORNECEDOR_REL_FORNECEDOR_PRODUTOFF = 13116 'Parâmetros: lCodFornecedor, sProduto
'O Fornecedor %l está relacionado com o Produto %s.
Public Const ERRO_LEITURA_IMPORTCLI = 13117 'Sem parametros
'Erro na leitura da tabela ImportCli.




