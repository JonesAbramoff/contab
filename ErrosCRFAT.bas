Attribute VB_Name = "ErrosCRFAT"
 Option Explicit

'C�digos de erro - Reservado de 6000 at� 6999
Public Const ERRO_ITEM_REPETIDO_NO_GRID = 6000 'Sem par�metros
'N�o pode haver repeti��o do Item de uma Categoria.
Public Const ERRO_LEITURA_CATEGORIACLIENTE = 6001 'Sem par�metros
'Erro na leitura da tabela de CategoriaCliente.
Public Const ERRO_LOCK_CATEGORIACLIENTE = 6002 'Par�metro: sCategoria
'Erro na tentativa de fazer 'lock' na tabela de CategoriaCliente com Categoria %s.
Public Const ERRO_INSERCAO_COMISSOES = 6003 'Parametro: iTipoTituo, lNumTitulo
'Erro na tentetiva de inserir comiss�es de um Documento do Tipo %i com N�mero Interno = %l.
Public Const ERRO_NOTA_FISCAL_ASSOCIADA_A_FATURA = 6004 'Parametro: n�mero da nota fiscal
'A nota fiscal %l j� est� associada a uma fatura. N�o � poss�vel realizar opera��o.
Public Const ERRO_LOCK_NFSREC = 6005 'Parametro : lNumnotaFiscal
' Erro na tentativa de "lock" da tabela NfsRec na Nota Fiscal n�mero %l.
Public Const ERRO_LEITURA_COMISSOES = 6006 'Parametro: lNumIntDoc
'Erro na leitura da tabela de Comissoes com n�mero de documento %l.
Public Const ERRO_LEITURA_COMISSOESNF = 6007 'Sem par�metro
'Erro na leitura de registros de comiss�es de notas fiscais, da tabela ComissoesNF.
Public Const ERRO_LEITURA_COMISSOESPEDVENDAS = 6008 'Sem par�metro
'Erro na leitura de registros de comiss�es de pedidos de venda
Public Const ERRO_COMISSOES_BAIXADA = 6009 'Parametros:lNumIntDoc e iTipoTitulo
'Erro na tentativa de excluir registro da tabela de Comiss�es, do documento %l, do tipo %i.A comiss�o j� foi baixada.
Public Const ERRO_ATUALIZACAO_COMISSOES = 6010 'Parametro: lNumIntDoc, iTipoTitulo
'Erro na tentativa atualizar as Comiss�es do Documento de Tipo %i e N�mero Interno %l na Tabela de Comiss�es
Public Const ERRO_ATUALIZACAO_COMISSOESNF = 6011 'Parametro: lNumNotaFiscal
'Erro na tentativa de atualizar as Comiss�es da Nota Fiscal de N�mero %l
Public Const ERRO_EXCLUSAO_COMISSOESNF = 6012 'Parametros:lNumIntDoc
'Erro na tentativa de excluir registro da tabela de Comiss�es de Notas Fiscais, do documento %l.
Public Const ERRO_EXCLUSAO_COMISSOES = 6013 'Parametros:lNumIntDoc e iTipoTitulo
'Erro na tentativa de excluir registro da tabela de Comiss�es, do documento %l, do tipo %i.
Public Const ERRO_INSERCAO_COMISSOESNF = 6014 'Sem par�metro
'Erro na tentetiva de inserir um registro na tabela de comiss�es de Notas Fiscais.
Public Const ERRO_NFREC_NAO_CADASTRADA = 6015 'Parametro: n�mero da nota fiscal, s�rie
'A Nota Fiscal com n�mero %l da s�rie %s n�o est� cadastrada.
Public Const ERRO_LEITURA_TIPODOCINFO = 6016 'parametro sigla do doc
'Erro na leitura do tipo de documento %s
Public Const ERRO_LEITURA_ESTADOS1 = 6017 'Par�metro: sSigla
'Erro na leitura da tabela Estados com o Estado %s.
Public Const ERRO_SIGLA_ESTADO_NAO_CADASTRADA = 6018 'Par�metro: Sigla.Text
'O Estado %s n�o est� cadastrado.
Public Const ERRO_SIGLA_ESTADO_NAO_PREENCHIDA = 6019 'Sem par�metro
'O preenchimento da Sigla do Estado � obrigat�rio.
Public Const ERRO_LEITURA_CATEGS_CLI = 6020  'parametros codigo do cliente e codigo da filial
'Erro na leitura das categorias associadas ao cliente %ld filial %d
Public Const ERRO_REGIAO_VENDA_NAO_CADASTRADA = 6021 'Parametro: iCodigo
'Regi�o de Venda com c�digo %i n�o est� cadastrada no Banco de Dados.
Public Const ERRO_LEITURA_CPRCONFIG = 6022 'Parametro sCodigo
'Erro na leitura do codigo %s na tabela CPRConfig.
Public Const ERRO_LOCK_CPRCONFIG = 6023 'Parametro sCodigo
'Erro no "lock" da tabela CPRConfig. Codigo = %s
Public Const ERRO_VALOR_PORCENTAGEM_JUROS = 6024 'Parametros dValor, lPorcentMaxima
'Porcentagem de Juros %d n�o est� entre 0 e %l.
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
'Erro na leitura do Vendedor com c�digo %i na tabela de Vendedores.
Public Const ERRO_LOCK_VENDEDOR = 6045
'Erro na tentativa de "lock" na tabela de Vendedores.
Public Const ERRO_LEITURA_REGIAO = 6046 'Parametro codigo da Regiao
'Erro na leitura da tabela de Regioes de Vendas , no codigo %i.
Public Const ERRO_LOCK_REGIAO = 6047
'Erro na tentativa de "lock" na tabela de Regioes de Vendas.
Public Const ERRO_CLIENTE_REL_NF_REC_PEND = 6049 'Parametro: lCodigo
'Erro na exclus�o do Cliente com c�digo %l. Est� relacionado com Nota Fiscal a Receber Pendente.
Public Const ERRO_CLIENTE_REL_TITULOS_REC = 6050 'Parametro: lCodigo
'Erro na exclus�o do Cliente com c�digo %l. Est� relacionado com T�tulos a Receber.
Public Const ERRO_LEITURA_TITULOS_REC = 6051
'Erro na leitura da tabela TitulosRec.
Public Const ERRO_LEITURA_DEBITOSRECCLI = 6052
'Erro na leitura da tabela DebitosRecCli .
Public Const ERRO_CLIENTE_REL_DEBITOS = 6053 'Parametro: lCodigo
'N�o � permitido a exclus�o do Cliente com c�digo %l, pois est� relacionado com Cr�dito a Receber.
Public Const ERRO_LEITURA_RECEB_ANTEC = 6054
'Erro na leitura da tabela de RecebAntecipados.
Public Const ERRO_CLIENTE_REL_RECEB_ANTEC = 6055 'Parametro: lCodigo
'N�o � permitido a exclus�o do Cliente com c�digo %l, pois est� relacionado com Recebimento Antecipado.
Public Const ERRO_CLIENTE_SEM_FILIAL = 6056 'Parametro codigo do cliente
'O cliente %l nao est� vinculado a nenhuma filial.
Public Const ERRO_EXCLUSAO_CLIENTE = 6057 'Parametro codigo do cliente
'Erro na tentativa de excluir o cliente %l.
Public Const ERRO_EXCLUSAO_FILIAISCLIENTES = 6058 'Parametro codigo do cliente
'Erro na exclusao das filiais do cliente %l.
Public Const ERRO_MODIFICACAO_FILIAISCLIENTES = 6059
'Erro na tentativa de modificacao na tabela de FiliaisClientes.
Public Const ERRO_LOCK_CLIENTES = 6060 'Parametro lCodigoCliente
'Erro na tentativa de "lock" da tabela Clientes, C�digo Cliente = %l .
Public Const ERRO_FILIAL_DESASSOCIADA_CLIENTE = 6061 'Parametro: sCGC
'Filial de Cliente com CGC %s desassociada de Fornecedor.
Public Const ERRO_CONDICAO_PAGTO_NAO_CADASTRADA = 6062 'Parametro iCodigo
'A Condi��o de Pagamento com c�digo %i n�o est� cadastrada no Banco de Dados.
Public Const ERRO_LEITURA_PADRAO_COBRANCA = 6063 ' Parametro iPadraoCobranca
'Erro na leitura do C�digo = %i da tabela Padr�es Cobranca .
Public Const ERRO_LEITURA_NOTAS_FISCAIS_REC = 6064
'Erro na leitura da tabela de Notas Fiscais a Receber.
Public Const ERRO_TIPO_CLIENTE_NAO_CADASTRADO = 6065 'Parametro: iCodigo
'Tipo de Cliente com c�digo %i n�o est� cadastrado.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_REC_PEND = 6066
'Erro na leitura da tabela de Notas Fiscais a Receber Pendentes.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_REC_BAIXADAS = 6067
'Erro na leitura da tabela de Notas Fiscais a Receber Baixadas.
Public Const ERRO_CLIENTE_REL_NF_REC = 6068 'Parametro: lCodigo
'Erro na exclus�o do Cliente com c�digo %l. Est� relacionado com Nota Fiscal a Receber.
Public Const ERRO_CLIENTE_REL_NF_REC_BAIXADA = 6069 'Parametro: lCodigo
'Erro na exclus�o do Cliente com c�digo %l. Est� relacionado com Nota Fiscal a Receber Baixada.
Public Const ERRO_LEITURA_TITULOS_REC_BAIXADOS = 6070
'Erro na leitura da tabela TitulosRecBaixados.
Public Const ERRO_LEITURA_TITULOS_REC_PEND = 6071
'Erro na leitura da tabela TitulosRecPend.
Public Const ERRO_CLIENTE_REL_TITULOS_REC_PEND = 6072 'Parametro: lCodigo
'Erro na exclus�o do Cliente com c�digo %l. Est� relacionado com T�tulos a Receber Pendentes.
Public Const ERRO_CLIENTE_REL_TITULOS_REC_BAIXADOS = 6073 'Parametro: lCodigo
'Erro na exclus�o do Cliente com c�digo %l. Est� relacionado com T�tulos a Receber Baixados.
Public Const ERRO_LEITURA_MVPERCLI2 = 6074 'Sem parametros
'Erro de leitura na tabela MvPerCli.
Public Const ERRO_LOCK_MVPERCLI = 6075 'Par�metros: iFilialEmpresa, iExercicio, lCliente, iFilial
'Erro na tentativa de "lock" na tabela MvPerCli. FilialEmpresa=%i, Exercicio=%i, Cliente=%l, Filial=%i.
Public Const ERRO_EXCLUSAO_MVPERCLI = 6076 'Par�metros: iFilialEmpresa, iExercicio, lCliente, iFilial
'Erro na exclus�o de registro na tabela MvPerCli. FilialEmpresa=%i, Exercicio=%i, Cliente=%l, Filial=%i.
Public Const ERRO_LOCK_MVDIACLI = 6078 'Par�metros: iFilialEmpresa, lCliente, iFilial, dtData
'Erro na tentativa de "lock" na tabela MvDiaCli. FilialEmpresa=%i, Cliente=%l, Filial=%i, Data=%dt.
Public Const ERRO_EXCLUSAO_MVDIACLI = 6079 'Par�metros: iFilialEmpresa, lCliente, iFilial, dtData
'Erro na exclus�o de registro na tabela MvPerCli. FilialEmpresa=%i, Cliente=%l, Filial=%i, Data=%dt.
Public Const ERRO_LEITURA_CARTEIRAS_COBRADOR = 6080
'Erro na leitura da tabela Carteiras Cobrador
Public Const ERRO_PADRAO_COBRANCA_INVALIDO = 6081
'O Padr�o de Cobran�a %d � inv�lido.
Public Const ERRO_CARTEIRA_COBRADOR_INEXISTENTE = 6082
'Carteira Cobrador inexistente.
Public Const ERRO_LEITURA_PADRAO_COBRANCA2 = 6083
'Erro na leitura da tabela de Padr�es de Cobranca.
Public Const ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO = 6084 'Parametro: iCodigo
'Condi��o de Pagamento com c�digo %i n�o � de Contas a Receber.
Public Const ERRO_MENSAGEM_NAO_CADASTRADA = 6085 'Parametro: iCodMensagem
'A Mensagem com c�digo %i n�o est� cadastrada no Banco de Dados.
Public Const ERRO_COBRADOR_NAO_CADASTRADO = 6086 'Parametro iCodCobrador
'O Cobrador com c�digo %i n�o est� cadastrado no Banco de Dados.
Public Const ERRO_TRANSPORTADORA_NAO_CADASTRADA = 6087 'Parametro: iCodTransportadora
'A Transportadora com c�digo %i n�o est� cadastrada no Banco de Dados.
Public Const ERRO_VENDEDOR_NAO_CADASTRADO = 6088 'Parametro: iCodVendedor
'O Vendedor com c�digo %i n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_COMISSOES_BAIXA = 6089    'Sem par�metro
'Erro na leitura de registros de comiss�es a serem baixadas
Public Const ERRO_LOCK_CATEGORIACLIENTEITEM = 6090  'Sem parametro
'Erro na tentativa de fazer 'lock' na tabela de Estados.
Public Const ERRO_CATEGORIACLIENTEITEM_INEXISTENTE = 6091 'Parametro: sItem, sCategoria
'O Item %s da Categoria %s de Cliente n�o existe.
Public Const ERRO_CATEGORIACLIENTE_INEXISTENTE = 6092 'Parametro: sCategoria
'A Categoria %s de Cliente n�o existe.
Public Const ERRO_LEITURA_CATEGORIACLIENTE1 = 6093 'Par�metro: sCategoria
'Erro na leitura da tabela de CategoriaCliente com Categoria %s.
Public Const ERRO_LEITURA_CATEGORIACLIENTEITEM = 6094  'Par�metro: sCategoria
'Erro na leitura da tabela de CategoriaClienteItem com Categoria %s.
Public Const ERRO_CATEGORIACLIENTE_NAO_INFORMADA = 6095 'Sem par�metros
'O preenchimento de Categoria � obrigat�rio.
Public Const ERRO_FALTA_ITEM_CATEGORIACLIENTE = 6096 'Sem par�metros
'Somente a Descri��o do Item foi informada. O Item n�o foi informado.
Public Const ERRO_MODIFICACAO_CATEGORIACLIENTE = 6097 'Par�metro: sCategoria
'Erro na modifica��o da tabela de CategoriaCliente com Categoria %s.
Public Const ERRO_MODIFICACAO_CATEGORIACLIENTEITEM = 6098 'Par�metro: sCategoria
'Erro na modifica��o da tabela de CategoriaClienteItem com Categoria %s.
Public Const ERRO_EXCLUSAO_CATEGORIACLIENTEITEM = 6099 'Par�metro: sCategoria
'Erro na exclus�o de registro na tabela CategoriaClienteItem com Categoria %s.
Public Const ERRO_INSERCAO_CATEGORIACLIENTE = 6100 'Par�metro: sCategoria
'Erro na inser��o da Categoria %s de Cliente na tabela CategoriaCliente.
Public Const ERRO_INSERCAO_CATEGORIACLIENTEITEM = 6101 'Par�metro: sCategoria
'Erro na inser��o de registro na tabela de CategoriaClienteItem com Categoria %s.
Public Const ERRO_LEITURA_FILIALCLIENTECATEGORIAS = 6102 'Par�metros: sCategoria, sItem
'Erro na leitura da tabela FilialClienteCategorias com Categoria %s e Item %s.
Public Const ERRO_LEITURA_FILIALCLIENTECATEGORIAS1 = 6103 'Par�metro: sCategoria
'Erro na leitura da tabela FilialClienteCategorias com Categoria %s.
Public Const ERRO_CATEGORIACLIENTEITEM_UTILIZADA = 6104  'Par�metros: lCliente, sItem, sCategoria
'O Cliente %l est� associado ao �tem %s da Categoria %s.
Public Const ERRO_CATEGORIACLIENTE_UTILIZADA = 6105 'Par�metro: sCategoria, lCliente
'A Categoria %s j� foi utilizada pelo Cliente %l.
Public Const ERRO_CATEGORIACLIENTE_NAO_CADASTRADA = 6106 'Par�metro: sCategoria
'A Categoria %s n�o est� cadastrada no Banco de Dados.
Public Const ERRO_EXCLUSAO_CATEGORIACLIENTE = 6107  'Par�metro: sCategoria
'Erro na exclus�o de uma Categoria %s de Cliente.
Public Const ERRO_LEITURA_TIPODECLIENTECATEGORIAS1 = 6108 'Par�metro: sCategoria
'Erro de leitura da tabela TipoDeClienteCategorias Com Categoria %s.
Public Const ERRO_EXCLUSAO_CATEGORIACLIENTE_UTILIZADA = 6109 'Par�metro: sCategoria, iTipo
'N�o � permitido excluir a Categoria %s porque est� sendo utilizada pelo Tipo de Cliente %i.
Public Const ERRO_LEITURA_TIPODECLIENTECATEGORIAS = 6110 'Par�metro: iTipoCliente
'Erro na leitura da tabela TipoDeClienteCategorias com Tipo de Cliente %i.
Public Const ERRO_LEITURA_CATEGORIACLIENTEITEM1 = 6111 'Par�metro: sItem, sCategoria
'Erro na leitura da tabela CategoriaClienteItem, cuja Categoria � %s e Item %s.
Public Const ERRO_CODIGO_NAO_PREENCHIDO = 6112 'Sem parametros
' Preenchimento do c�digo � obrigat�rio.
Public Const ERRO_INSERCAO_FILIALCLIENTECATEGORIAS = 6113 'Par�mero: lCodigo
'Erro na tentativa de inserir um registro na tabela FilialClienteCategorias com c�digo do Cliente %l.
Public Const ERRO_LEITURA_FILIALCLIENTECATEGORIAS2 = 6114 'Par�metros: lCliente
'Erro na leitura da tabela FilialClienteCategorias com Cliente %l.
Public Const ERRO_ALTERACAO_FILIALCLIENTECATEGORIAS = 6115 'Par�metros: lCliente
'Erro na tentativa deatualizar um registro na tabela FilialClienteCategorias com Cliente %l.
Public Const ERRO_EXCLUSAO_FILIALCLIENTECATEGORIAS = 6116 'Par�metro: lCodCliente
'Erro na exclus�o das Categorias das Filiais do Cliente %l.
Public Const ERRO_EXCLUSAO_FILIALCLIENTECATEGORIAS1 = 6117 'Par�metro: lCodCliente, iFilial
'Erro na exclus�o das Categorias das Filiais do Cliente %l e Filial %i.
Public Const ERRO_LOCK_CATEGORIACLIENTEITEM2 = 6118 'Par�metros: sCategoria, sItem
'Erro na tentativa de fazer "lock" na tabela CategoriaClienteItem com Categoria %s e Item %s.
Public Const ERRO_ITEM_CATEGORIA_NAO_CADASTRADO = 6119 'Par�metros: sItem, sCategoria
'O Item %s da Categoria %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_CRFATCONFIG = 6120 'Parametro sCodigo
'Erro na leitura do codigo %s na tabela CRFATConfig.
Public Const ERRO_FILIALCLIENTE_NOME_DUPLICADO = 6121 'Parametro sNomeFilial
'O nome %s j� est� sendo usado em outra Filial de Cliente.
Public Const ERRO_LOCK_CRFATCONFIG = 6122 'Parametro sCodigo
'Erro no "lock" da tabela CRFATConfig. Codigo = %s
Public Const ERRO_ATUALIZACAO_CRFATCONFIG = 6124 'Parametro sCodigo
'Erro ao atualizar o registro de configura��o que possui o codigo %s na tabela CRFATConfig.
Public Const ERRO_CATEGORIACLIENTE_TAMMAX = 6126 'parametros: tam max da categoria
'A categoria deve ter no m�ximo %i caracteres.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_CLI = 6127
'Erro na leitura da nota fiscal.
Public Const ERRO_EXISTEM_NOTAS_FISCAIS_CLI = 6128
'Existe ao menos uma nota fiscal cadastrada para este cliente.
Public Const ERRO_LEITURA_TIPODOCINFO1 = 6129 'Sem parametro
'Erro na leitura da tabela dos tipos de documentos.
Public Const ERRO_LEITURA_NOTA_FISCAL_NUM_SERIE = 6130 'Parametros: S�rie, N�mero e Filial
'Erro na leitura da nota fiscal s�rie %s n�mero %l da filial %d
Public Const ERRO_NFISCAL_NUM_SERIE_NAO_CADASTRADA = 6131 'Parametros: S�rie, N�mero e Filial
'A nota fiscal s�rie %s n�mero %l da filial %d n�o est� cadastrada ou j� foi baixada.
Public Const ERRO_ATUALIZACAO_ESTADOS = 6132  'Par�metro: sSigla
'Erro na tentativa de atualizar um registro na tabela de Estados com Estado de sigla %s.
Public Const ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO1 = 6133 'Parametro: Tipo.Text
'O Tipo de Documento %s n�o est� cadastrado.
Public Const ERRO_LOCK_TIPOSDOCINFO = 6134 'Par�metro: sSiglaMovto
'Erro na tentativa de fazer "lock" na tabela TiposDocInfo com Documento %s.
Public Const ERRO_LEITURA_NATUREZAOP = 6135 'Parametro: sCodigo
'Erro na leitura da tabela de Naturezas de Opera��o %s.
Public Const ERRO_LOCK_NATUREZAOP = 6136 'Parametro: sCodigo
'N�o conseguiu fazer o lock da Natureza de Opera��o %s.
Public Const ERRO_ATUALIZACAO_NATUREZAOP = 6137 'Parametro sCodigo
'Erro de atualiza��o da Natureza de Opera��o %s.
Public Const ERRO_INSERCAO_NATUREZAOP = 6138 'Parametro: sCodigo
'Erro na inser��o da Natureza de Opera��o %s.
Public Const ERRO_NATUREZAOP_INEXISTENTE = 6139 'Parametro: sCodigo
'A Natureza de Opera��o %s n�o est� cadastrada.
Public Const ERRO_EXCLUSAO_NATUREZAOP = 6140 'Parametro: sCodigo
'Houve um erro na exclus�o da Natureza de Opera��o %s.
Public Const ERRO_LEITURA_NATUREZAOP1 = 6141 'Sem Parametros
'Erro na leitura da tabela de Naturezas de Opera��o.
Public Const ERRO_CATEGORIA_CLIENTE_NAO_PREENCHIDA = 6142 'Sem par�metro
'O preenchimento da Categoria do Cliente � obrigat�rio.
Public Const ERRO_CATEGORIA_CLIENTE_ITEM_NAO_PREENCHIDA = 6143 'Sem par�metro
'O preenchimento do item da Categoria do Cliente � obrigat�rio.
Public Const ERRO_TITRECPEND_JA_CADASTRADO = 6144  'Parametros: lNumTitulo, sSiglaDocumento
'Aten��o. J� existe um t�tulo a receber pendente cadastrado com esta identifica��o. T�tulo = %l, Sigla = %s.
Public Const ERRO_TITRECBAIXA_JA_CADASTRADO = 6145  'Parametros: lNumTitulo, sSiglaDocumento
'Aten��o. J� existe um t�tulo a receber baixado cadastrado com esta identifica��o. T�tulo = %l, Sigla = %s.
Public Const ERRO_TITREC_JA_CADASTRADO = 6146  'Parametros: lNumTitulo, sSiglaDocumento
'Aten��o. J� existe um t�tulo a receber cadastrado com esta identifica��o. T�tulo = %l, Sigla = %s.
Public Const ERRO_NFISCAL_NUMINTDOCCPR_NAO_ZERO = 6147 'Parametros: S�rie, N�mero e Filial
'A nota fiscal s�rie %s n�mero %l da filial %i j� est� associada a um t�tulo.
Public Const ERRO_ATUALIZACAO_NFISCAL = 6148  'Parametros: S�rie, N�mero e Filial
'Ocorreu um erro na tentativa de atualizar um registro na tabela de Notas Fiscais. S�rie = %s, N�mero = %l, Filial = %i
Public Const ERRO_FILIALCLIENTE_NAO_CADASTRADA = 6149 'Parametros: lCodCliente, iCodFilial
'A Filial %i do Cliente com c�digo %l n�o est� cadastrada no Banco de Dados.
Public Const ERRO_LOCK_FILIAISCLIENTES = 6150 'Parametros: lCodCliente, iCodFilial
'Erro na tentativa de "lock" da tabela FiliaisClientes. CodCliente = %l, CodFilial = %i.
Public Const ERRO_TIPO_NAO_SELECIONADO = 6151 'Par�metro: iCodigo
'Tipo com c�digo %i est� na List e n�o foi selecionado.
Public Const ERRO_LEITURA_CLIENTE_PEDIDOS_DE_VENDA = 6152 'Par�metro: lCodigo
'Erro na leitura das tabelas PedidosDeVenda e PedidosDeVendabaixados com Cliente %l.
Public Const ERRO_CLIENTE_REL_PED_VENDA = 6153 'Par�metros: lCodigo, iFilialEmpresa, lPedido
'O Cliente %l est� relacionada com Pedido de Venda com Filial Empresa = %i, C�digo = %l.
Public Const ERRO_CLIENTE_REL_CHEQUE_PRE = 6154 'Par�metros: lCodigo, lNumIntCheque
'O Cliente %l est� associado com Cheque Pr�.
Public Const ERRO_LEITURA_CHEQUEPRE2 = 6155 'Par�metro: lCodigo
'Erro na leitura da tabela de ChequePre com Cliente %l.
Public Const ERRO_CATEGORIA_SEM_VALOR_CORRESPONDENTE = 6156 'Par�metro: sCategoria
'A Categoria %s n�o tem um valor correspondente.
Public Const ERRO_CATEGORIACLIENTE_REPETIDA_NO_GRID = 6157 'Sem par�metro
'N�o pode haver repeti��o de Categorias de Cliente no Grid.
Public Const ERRO_LEITURA_TIPOSDECLIENTE1 = 6158 'Par�metro: iCodigo
'Erro na leitura da tabela TipoDeCliente com C�digo %i.
Public Const ERRO_LEITURA_CONDICAOPAGTO_PEDIDOS_DE_VENDA = 6159 'Par�metro: iCodigo
'Erro na leitura das tabelas PedidosDeVenda e PedidosDeVendabaixados com Condi��o de Pagamento %i.
Public Const ERRO_CONDICAOPAGTO_REL_PED_VENDA = 6160 'Par�metros: iCodigo, lPedido, iFilial
'A Condi��o de Pagamento %i est� vinculada a Pedido de Venda %l da Filial Empresa %i.
Public Const ERRO_LEITURA_CONDICOESPAGTO = 6161 'Sem par�metros
'Erro na leitura da tabela CondicoesPagto.
Public Const ERRO_DIA_DO_MES_INVALIDO = 6162 'Sem par�metros
'Dia do M�s tem que estar entre 1 e 30.
Public Const ERRO_DESCRICAO_REDUZIDA_NAO_PREENCHIDA = 6163 'Sem par�metros
'Descric�o Reduzida deve ser preenchida.
Public Const ERRO_NUMERO_PARCELAS_NAO_PREENCHIDA = 6164 'Sem par�metros
'N�mero de Parcelas deve ser preenchido.
Public Const ERRO_DIAS_PARA_PRIMEIRA_PARCELA_NAO_PREENCHIDA = 6165 'Sem par�metros
'Dias para Primeira Parcela deve ser preenchido.
Public Const ERRO_DIA_DO_MES_NAO_PREENCHIDO = 6166 'Sem par�metros
'Dia do M�s deve ser preenchido.
Public Const ERRO_INTERVALO_ENTRE_PARCELAS_NAO_PREENCHIDO = 6167 'Sem par�metros
'Intervalo entre Parcelas deve ser preenchido.
Public Const ERRO_DESCRICAO_REDUZIDA_CONDICAOPAGTO_REPETIDA = 6168 'Sem par�metros
'Descri��o Reduzida � atributo de outra Condi��o de Pagamento.
Public Const ERRO_INSERCAO_CONDICAOPAGTO = 6169 'Par�metro: iCodigo
'Erro na inser��o da Condi��o de Pagamento %i.
Public Const ERRO_ATUALIZACAO_CONDICAOPAGTO = 6170 'Par�metro: iCodigo
'Erro na atualiza��o da Condi��o de Pagamento %i.
Public Const ERRO_CONDICAOPAGTO_RELACIONADA_COM_CLIENTE = 6171 'Par�metro: lTotal
'Condi��o de Pagamento est� relacionada com %l Cliente(s).
Public Const ERRO_CONDICAOPAGTO_RELACIONADA_COM_FORNECEDOR = 6172 'Par�metro: lTotal
'Condi��o de Pagamento est� relacionada com %l Fornecedores.
Public Const ERRO_CONDICAOPAGTO_RELACIONADA_COM_TIPOSDECLIENTE = 6173 'Par�metro: lTotal
'Condi��o de Pagamento est� relacionada com %l Tipo(s) de Cliente(s).
Public Const ERRO_CONDICAOPAGTO_RELACIONADA_COM_TIPOSDEFORNECEDOR = 6174 'Par�metro: lTotal
'Condi��o de Pagamento est� relacionada com %s Tipo(s) de Fornecedor.
Public Const ERRO_EXCLUSAO_CONDICAOPAGTO = 6175 'Par�metro: iCodigo
'Erro na exclus�o da Condi��o de Pagamento %i.
Public Const ERRO_NUM_PARCELAS_EXCESSIVO = 6176 'Par�metros: sNumero, iNumMaximo
'O n�mero de parcelas %s ultrapassou o limite m�ximo %i.
Public Const ERRO_LEITURA_FILIALCLIENTE_PEDIDOS_DE_VENDA = 6177 'Par�metro: lCodCliente, iCodFilial
'Erro na leitura das tabelas PedidosDeVenda e PedidosDeVendabaixados com Cliente %l e Filial %i.
Public Const ERRO_FILIALCLIENTE_REL_PED_VENDA = 6178 'Par�metros: lCodCliente, iCodFilial, iFilial, lPedido
'Filial Cliente com chave codCliente=%l, codFilial=%i est� relacionada com Pedido de Venda com chave Filial Empresa = %i, C�digo = %l.
Public Const ERRO_PERCENTAGEM_EMISSAO_NAO_PREENCHIDA = 6179 'Sem par�metros
'O Percentual de Comiss�o de Emiss�o deve ser preenchido.
Public Const ERRO_SOMA_EMISSAO_MAIS_BAIXA = 6180 'Sem par�metros
'O Percentual de Comiss�o de Emiss�o mais o de Baixa n�o d� 100%.
Public Const ERRO_LEITURA_TIPOSDEVENDEDOR = 6181 'Par�metro: iCodigo
'Erro na leitura do Tipo de Vendedor %i.
Public Const ERRO_TIPODEVENDEDOR_NAO_CADASTRADO = 6182 'Par�metro: iCodigo
'Tipo de Vendedor %i n�o est� cadastrado.
Public Const ERRO_INSERCAO_TIPOSDEVENDEDOR = 6183 'Par�metro: iCodigo
'Erro na inser��o do Tipo de Vendedor %i.
Public Const ERRO_ATUALIZACAO_TIPOSDEVENDEDOR = 6184 'Par�metro: iCodigo
'Erro na atualiza��o do Tipo de Vendedor %i.
Public Const ERRO_EXCLUSAO_TIPOSDEVENDEDOR = 6185 'Par�metro: iCodigo
'Erro na exclus�o do Tipo de Vendedor %i.
Public Const ERRO_LOCK_TIPOSDEVENDEDOR = 6186 'Par�metro: iCodigo
'N�o conseguiu fazer o lock do Tipo de Vendedor %i.
Public Const ERRO_TIPODEVENDEDOR_RELACIONADO_VENDEDOR = 6187 'Par�metros: iCodTipoVendedor, iCodVendedor
'Tipo de Vendedor %i est� relacionado com Vendedor %i.
Public Const ERRO_DESCRICAO_TIPO_VENDEDOR_REPETIDA = 6188 'Par�metro: iCodigo
'Tipo de Vendedor %i tem a mesma descri��o.
Public Const ERRO_DESCRICAO_NAO_PREENCHIDA = 6189 'Sem par�metros
'Preenchimento da descri��o � obrigat�rio.
Public Const ERRO_LEITURA_FILIAISCLIENTES1 = 6190 'Par�metro: iCodigo
'Erro na leitura da tabela FiliaisClientes com Vendedor %i.
Public Const ERRO_LEITURA_COMISSOESPEDVENDAS1 = 6191 'Par�metro: iCodigo
'Erro na leitura da tabela ComissoesPedVendas com Vendedor %i.
Public Const ERRO_LEITURA_COMISSOESPEDVENDASBAIXADOS1 = 6192 'Par�metro: iCodigo
'Erro na leitura da tabela ComissoesPedVendasBaixados com Vendedor %i.
Public Const ERRO_LEITURA_COMISSOESNF1 = 6193 'Par�metro: iCodigo
'Erro na leitura da tabela ComissoesNF com Vendedor %i.
Public Const ERRO_VENDEDOR_REL_FILIALCLIENTE = 6194 'Par�metros: iCodigo, lCliente, iFilial
'O Vendedor %i est� relacionado com Cliente %l da Filial %i.
Public Const ERRO_VENDEDOR_REL_COMISSAOPEDVENDA = 6195 'Par�metros: iCodigo, lPedido
'O Vendedor %i est� relacionado com Comiss�o de Pedido de Venda %l.
Public Const ERRO_VENDEDOR_REL_COMISSAOPEDVENDABAIXADA = 6196 'Par�metros: iCodigo, lPedido
'O Vendedor %i est� relacionado com Comiss�o de Pedido de Venda Baixado %l.
Public Const ERRO_VENDEDOR_REL_COMISSAONF = 6197 'Par�metros: iCodigo, lNumInt
'O Vendedor %i est� relacionado com Comiss�o de Nota Fiscal.
Public Const ERRO_TIPO_VENDEDOR_NAO_ENCONTRADO = 6198 'Parametro sTipoVendedor
'Tipo Vendedor com descri��o %s n�o foi encontrado.
Public Const ERRO_REGIAO_VENDA_NAO_ENCONTRADA = 6199 'Parametro sRegiaoVenda
'Regi�o de Venda com descri��o %s n�o foi encontrada.
Public Const ERRO_NOME_REDUZIDO_VENDEDOR_REPETIDO = 6200 'Parametro: iCodigo
'Vendedor %i tem o mesmo Nome Reduzido.
Public Const ERRO_INSERCAO_VENDEDOR = 6201 'Parametro iCodigo
'Erro na inser��o do Vendedor %i.
Public Const ERRO_INSERCAO_ENDERECO = 6202 'Parametro lCodigo
'Erro na inser��o do Endere�o %l.
Public Const ERRO_ATUALIZACAO_VENDEDOR = 6203 'Parametro iCodigo
'Erro na atualiza��o do Vendedor %i.
Public Const ERRO_ATUALIZACAO_ENDERECO = 6204 'Parametro iCodigo
'Erro na atualiza��o do Endere�o %i.
Public Const ERRO_VENDEDOR_RELACIONADO_COMISSAO = 6205 'Parametro iCodigo, lNumIntCom
'Vendedor %i est� relacionado com Comiss�o com n�mero interno %l.
Public Const ERRO_EXCLUSAO_VENDEDOR = 6206 'Parametro iCodigo
'Erro na exclus�o do Vendedor %i.
Public Const ERRO_TIPO_VENDEDOR_NAO_PREENCHIDO = 6207 'Sem Parametros
'Preenchimento do Tipo de Vendedor � obrigat�rio.
Public Const ERRO_FORNECEDOR_VINCULADO_COND_PAGTO = 6209 'Parametro: iCodCondicaoPagto
'A Condi��o de Pagamento com c�digo %i n�o pode deixar de ser usada em Contas a Pagar. Existem Fornecedores vinculados a ela.
Public Const ERRO_LEITURA_TIPOSFORNECEDOR = 6210
'Erro na leitura da tabela TiposdeFornecedor.
Public Const ERRO_TIPO_FORNECEDOR_VINCULADO_COND_PAGTO = 6211 'Parametro: iCodCondicaoPagto
'A Condi��o de Pagamento com c�digo %i n�o pode deixar de ser usada em Contas a Pagar. Existem Tipos de Fornecedor vinculados a ela.
Public Const ERRO_CLIENTE_VINCULADO_COND_PAGTO = 6212 'Parametro: iCodCondicaoPagto
'A Condi��o de Pagamento com c�digo %i n�o pode deixar de ser usada em Contas a Receber. Existem Clientes vinculados a ela.
Public Const ERRO_LEITURA_TIPOCLIENTE = 6213 'Sem parametros
'Erro na leitura da tabela TiposDeCliente.
Public Const ERRO_TIPO_CLIENTE_VINCULADO_COND_PAGTO = 6214 'Parametro: iCodCondicaoPagto
'A Condi��o de Pagamento com c�digo %i n�o pode deixar de ser usada em Contas a Receber. Existem Tipos de Cliente vinculados a ela.
Public Const ERRO_FILIALCLIENTE_REL_DEBITOS = 6215 'Parametros: lCodCliente, iCodFilial
'N�o � permitido a exclus�o do Cliente %l com c�digo de Filial = %l, pois est� relaconada com Cr�dito a Receber.
Public Const ERRO_FILIALCLIENTE_REL_RECEB_ANTEC = 6216 'Parametros: lCodCliente, iCodFilial
'N�o � permitido a exclus�o do Cliente %l com c�digo de Filial = %l, pois est� relacionada com Recebimento Antecipado.
Public Const ERRO_EXCLUSAO_FILIALCLIENTE = 6217 'Parametros: lCodCliente, iCodFilial
'Erro na tentativa de excluir a Filial Cliente. CodCliente = %l, CodFilial = %i.
Public Const ERRO_LEITURA_FILIALCLIENTECATEGORIAS3 = 6218 'Par�metros: lCodCliente, iCodFilial
'Erro vna leitura da tabela FilialClienteCategorias com Cliente %l e Filial %i.
Public Const ERRO_LEITURA_REGIOESVENDAS = 6219 'Parametro iCodigo
' Erro na leitura da Regi�o de Venda %i.
Public Const ERRO_LOCK_REGIOESVENDAS = 6220 'Parametro
' N�o conseguiu fazer o lock da Regi�o de Venda %i.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_FIL_CLI = 6221
'Erro na leitura da nota fiscal do cliente.
Public Const ERRO_EXISTEM_NOTAS_FISCAIS_FIL_CLI = 6222
'Existe ao menos uma nota fiscal cadastrada para este cliente.
Public Const ERRO_FILIALCLIENTE_REL_TITULOS_REC_PEND = 6223 'Parametros: lCodCliente, iCodFilial
'Erro na exclus�o da Filial de Cliente com CodCliente=%l, CodFilial=%i. Est� relacionada com T�tulos a Receber Pendentes.
Public Const ERRO_FILIALCLIENTE_REL_TITULOS_REC = 6224 'Parametros: lCodCliente, iCodFilial
'Erro na exclus�o da Filial de Cliente com CodCliente=%l, CodFilial=%i. Est� relacionada com T�tulos a Receber.
Public Const ERRO_FILIALCLIENTE_REL_TITULOS_REC_BAIXADOS = 6225 'Parametros: lCodCliente, iCodFilial
'Erro na exclus�o da Filial de Cliente com CodCliente=%l, CodFilial=%i. Est� relacionada com T�tulos a Receber Baixados.
Public Const ERRO_LEITURA_VENDEDORES = 6226 'Sem parametros
'Erro na leitura da tabela Vendedores.
Public Const ERRO_CATEGORIA_SEM_ITEM_CORRESPONDENTE = 6227
'Categoria sem item.
Public Const ERRO_TITULO_REL_NFISCAL = 6228 'Sem Par�metro
'N�o � permitido excluir o T�tulo, porque est� relacionado com Nota Fiscal.
Public Const ERRO_CLIENTE_NAO_CADASTRADO1 = 6229 'Par�metro: sCliente
'O Cliente %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_VALOR_NAO_PREENCHIDO1 = 6230
'O campo Valor n�o foi prenchido.
Public Const ERRO_TIPODESCONTO_NAO_ENCONTRADO = 6231  'Parametro: iCodigo
'O Tipo de Desconto %i n�o foi encontrado.
Public Const ERRO_TIPODESCONTO_NAO_ENCONTRADO1 = 6232  'Parametro: sTipoDesconto
'O Tipo de Desconto %s n�o foi encontrado.
Public Const ERRO_TITULORECEBER_NAO_CADASTRADO = 6233 'lNumIntDoc
'O T�tulo a Receber com N�mero Interno %l n�o est� cadastrado.
Public Const ERRO_CHEQUEPRE_NAO_CADASTRADO = 6234 'lNumIntCheque
'O ChequePre com N�mero Interno = %l n�o foi encontrado.
Public Const ERRO_TITULOREC_FILIALEMPRESA_DIFERENTE = 6235 'Par�metro: lNumTitulo, sSiglaDocumento
'N�o � poss�vel modificar oT�tulo a Receber do Tipo %s e de N�mero %l. Ela pertence a outra Filial da Empresa.
Public Const ERRO_VALOR_COMISSAO_NAO_INFORMADO = 6236 'iComissao
'O Valor da Comissao %i do Titulo n�o foi informado.
Public Const ERRO_LEITURA_TITULOSRECBAIXADOS = 6237 'Parametro: lNumTitulo, sSiglaDocumento
'Erro na tentativa de leitura do Titulo Receber N�mero %l do Tipo %s na Tabela TitulosRecBaixado
Public Const ERRO_LEITURA_TITULOSRECPEND = 6238 'Parametros: lNumTitulo, sSiglaDocumento
'Erro na tentativa de leitura do Titulo Receber N�mero %l do Tipo %s na Tabela TitulosRecPend
Public Const ERRO_LEITURA_TITULOSREC = 6239 'Parametros: lNumTitulo,  sSiglaDocumento
'Erro na tentativa de leitura do Titulo Receber N�mero %l do Tipo %s na Tabela TitulosRec
Public Const ERRO_LEITURA_TITULOSREC1 = 6240 'Parametros: lNumIntDoc
'Erro na tentativa de leitura do Titulo Receber com N�mero Interno %l na Tabela TitulosRec
Public Const ERRO_TITULOREC_PENDENTE_MODIFICACAO = 6241 'Parametro: lNumTitulo, sSiglaDocumento
'N�o � poss�vel modificar o Titulo a Receber N�mero %l do Tipo %s pois ele faz parte de um Lote Pendente.
Public Const ERRO_TITULOREC_BAIXADO_MODIFICACAO = 6242 'Parametro: lNumTitulo, sSiglaDocumento
'N�o � poss�vel modificar o Titulo a Receber N�mero %l do Tipo %s pois ele est� baixado.
Public Const ERRO_LOCK_TITULOSREC = 6243 'Parametro: lNumTitulo, sSiglaDocumento
'Erro na tentativa de fazer lock no Titulo Receber N�mero %l do Tipo %s na Tabela TitulosRec
Public Const ERRO_VENDEDOR_COMISSAO_PARCELA_NAO_INFORMADO = 6244 'Parametro: iComissao,iParcela
'O Vendedor da Comiss�o %i da Parcela %i n�o foi informado
Public Const ERRO_VALORISS_MAIOR = 6245 'Par�metros: sValorISS, sValor
'Valor do ISS n�o pode ser maior do que o Valor do Titulo
Public Const ERRO_SOMA_PARCELAS_DIFERENTE = 6246 'Par�metros: dValorParcelas, dValorTitulo
'A soma das Parcelas � diferente do Valor do Titulo
Public Const ERRO_VALORTITULO_MENOS_IMPOSTOS = 6247
'Valor do T�tulo menos Impostos retidos deve ser positivo
Public Const ERRO_TITULORECEBER_SEM_PARCELAS = 6248 'Par�metro: lNumIntDoc
'T�tulo a Receber com n�mero interno %l n�o tem Parcelas associadas.
Public Const ERRO_LEITURA_PARCELASREC1 = 6249  'lNumIntTitulo
'Erro na tentativa de ler as Parcelas referenes ao T�tulo de N�mero Interno %l na tabela de Parcelas a Receber.
Public Const ERRO_VALORBASE_COMISSAO_PARCELA_NAO_INFORMADO = 6250 'Par�metros: iComissao, iParcela
'O Valor Base da Comiss�o %i da Parcela %i n�o foi informado.
Public Const ERRO_VALOR_CHQPRE_PARCELA_NAO_PREENCHIDO = 6251  'Parametro: iParcela
'O Valor do ChequePre da Parcela %i n�o foi preenchido.
Public Const ERRO_PERCENTUAL_COMISSAO_PARCELA_NAO_INFORMADO = 6252 'iComissao, iParcela
'O Percentual da Comissao %i da Parcela %i n�o foi informado
Public Const ERRO_PARCELA_RECEBER_NAO_CADASTRADA = 6253 'lNumIntTitulo, iParcela
'A Parcela %i do Titulo a Receber com N�mero Interno %l n�o foi encontrada
Public Const ERRO_ATUALIZACAO_CHEQUESPRE = 6254 'lNumIntCheque
'Erro na tentativa de atualizar o CHequePre de N�mero Interno %l na tabela  ChequesPre.
Public Const ERRO_TITULORECEBER_PENDENTE_EXCLUSAO = 6255 'lNumTitulo, sSiglaDocumento
'N�o � poss�vel excluir o Titulo a Receber do Tipo %s e N�mero %l por que ele faz parte de um Lote pendente.
Public Const ERRO_TITULORECEBER_BAIXADO_EXCLUSAO = 6256 'lNumTitulo, sSiglaDocumento
'N�o � poss�vel excluir o Titulo a Receber do Tipo %s e N�mero %l por que ele est� baixado.
Public Const ERRO_TITULORECEBER_NAO_CADASTRADO1 = 6257 'sSiglaDocumento,lNumTitulo
'O titulo a Receber do Tipo %s e N�mero %l n�o foi encontrado.
Public Const ERRO_VALOR_CHQPRE_DIFERENTE_DESCONTO = 6258 'iParcela, dDesconto
'O Valor do ChequePre da Parcela %i � diferente do Valor da Parcela com desconto que � %d
Public Const ERRO_DATADEPOSITO_CHQPRE_PARCELA_NAO_PREENCHIDA = 6259 'iParcela
'A Data de Deposito de ChequePre da Parcela %i n�o foi preenchida
Public Const ERRO_VALOR_COMISSAO_PARCELA_NAO_INFORMADO = 6260 'iComissao, iParcela
'O Valor da Comiss�o %i da Parcela %i n�o foi preenchido.
Public Const ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA = 6261 'iComissao, iParcela
'A Data do Desconto %i da Parcela %i n�o foi preenchida.
Public Const ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO = 6262 'iComissao, iParcela
'O Valor do Desconto %i da Parcela %i n�o foi preenchido.
Public Const ERRO_CODIGO_DESCONTO_PARCELA_NAO_PREENCHIDO = 6263 'iComissao, iParcela
'O C�digo do Desconto %i da Parcela %i n�o foi preenchido.
Public Const ERRO_BANCO_CHQPRE_PARCELA_NAO_PREENCHIDO = 6264 'iParcela
'O Banco do ChequePre da Parcela %i n�o foi preenchido
Public Const ERRO_NUMERO_CHQPRE_PARCELA_NAO_PREENCHIDO = 6265 'iParcela
'O N�mero do ChequePre da Parcela %i n�o foi preenchido
Public Const ERRO_AGENCIA_CHQPRE_PARCELA_NAO_PREENCHIDA = 6266 'iParcela
'A Ag�ncia do Cheque Pr� da Parcela %i n�o foi preenchida.
Public Const ERRO_CONTA_CHQPRE_PARCELA_NAO_PREENCHIDA = 6267 'iParcela
'A Conta Corrente do Cheque Pr� da Parcela %i n�o foi preenchida.
Public Const ERRO_SOMA_COMISSOES_PARCELA = 6268 'iParcela
'A soma das Comiss�es da Parcela %i � maior ou igual ao valor da Parcela.
Public Const ERRO_SOMA_COMISSOES_EMISSAO = 6269
'A soma das Comiss�es � maior ou igual ao valor do T�tulo.
Public Const ERRO_CARTEIRA_COBRADOR_NAO_INFORMADA = 6270 'Sem par�metro
'A Carteira do Cobrador deve ser informada.
Public Const ERRO_LEITURA_CARTEIRAS_COBRADOR1 = 6271 'Par�metro iCobrador
'Erro na leitura da tabela CarteirasCobrador. Cobrador %i.
Public Const ERRO_COBRADOR_SEM_CARTEIRA = 6272 'Par�metro iCobrador
'O Cobrador %i n�o possui Carteiras cadastradas
Public Const ERRO_LEITURA_ESTADOS = 6273 'Par�metro sEstado
'Erro na leitura da tabela Estados. Estado: %s
Public Const ERRO_TIPOCLIENTE_INEXISTENTE1 = 6274 'Par�metro: sTipoCliente
'Esse Tipo de Cliente n�o est� cadastrado.
Public Const ERRO_LEITURA_TABELA_CARTEIRASCOBRANCA1 = 6275 'Parametro ObjCarteiraCobranca.iCodigo
'Erro na leitura da tabela de Carteiras de Cobran�a. Carteira N�: %s
Public Const ERRO_PARAMETRO_OBRIGATORIO = 6276 'Sem Parametro
'Parametro � obrigat�rio
Public Const ERRO_EXCLUSAO_PADRAO_COBRANCA = 6277 'Parametro iCodigo
'Erro na exclus�o do Padr�o de Cobran�a %i.
Public Const ERRO_LEITURA_CONTASCORRENTES = 6278
'Erro na leitura da tabela de contas correntes internas
Public Const ERRO_COBRADOR_USADO_BORDEROCOBRANCA = 6279
'O Cobrador est� sendo utilizado em um Border� de Cobran�a.
Public Const ERRO_COBRADOR_USADO_OCORRENCIA = 6280
'O Cobrador est� sendo utilizado em uma Ocorr�ncia.
Public Const ERRO_COBRADOR_MESMO_NOMEREDUZIDO = 6281 'par�metro: sNomeReduzido
'J� existe no BD um Cobrador com o Nome Reduzido %s.
Public Const ERRO_COBRADOR_USADO_BAIXASPARCREC = 6282
'O Cobrador foi utilizado em uma Baixa a Receber.
Public Const ERRO_COBRADOR_USADO_FILIAISCLIENTE = 6283
'O Cobrador est� sendo utilizado por uma Filial Cliente.
Public Const ERRO_LEITURA_OCORRENCIASREMPARCREC = 6284
'Erro na leitura da tabela OcorrenciasRemParcRec.
Public Const ERRO_COBRADOR_NAO_INFORMADO = 6286 'Sem parametros
'Erro Cobrador n�o foi Informado
Public Const ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO = 6287 'Parametro iCodigo
'Erro carteira cobrador %i n�o cadastrada
Public Const ERRO_EXCLUSAO_CARTEIRASCOBRADOR = 6288 'Parametros iCarteira, iCobrador
'Erro exclus�o da Carteira %i do cobrador %i
Public Const ERRO_LEITURA_TABELA_COBRADOR = 6289 'Sem parametros
'Erro na leitura da tabela de Cobradores.
Public Const ERRO_LOCK_PADRAO_COBRANCA = 6290 'Parametro iCodigo
'N�o conseguiu fazer o lock do Padr�o de Cobran�a %i.
Public Const ERRO_PADRAO_COBRANCA_RELACIONADO_COM_TIPOS_DE_CLIENTE = 6291 'Parametro lTotal
'Padr�o de Cobran�a est� relacionado com %l Tipos de Cliente.
Public Const ERRO_CONTACORRENTE_INEXISTENTE1 = 6292 'Parametro: CodContaCorrente.Text
'A Conta Corrente %s nao est� cadastrada.
Public Const ERRO_DESCRICAO_PADRAO_COBRANCA_REPETIDA = 6293 'Sem Parametros
'Descri��o � atributo de outro Padr�o de Cobran�a.
Public Const ERRO_INSERCAO_PADRAO_COBRANCA = 6294 'Parametro iCodigo
'Erro na inser��o do Padr�o de Cobran�a %i.
Public Const ERRO_PADRAO_COBRANCA_NAO_CADASTRADO = 6295 'Parametro sPadraoCobranca
'O Padrao Cobran�a %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_ATUALIZACAO_PADRAO_COBRANCA = 6296 'Parametro iCodigo
'Erro na atualiza��o do Padr�o de Cobran�a %i.
Public Const ERRO_BANCO_INEXISTENTE1 = 6297 'Parametro: Banco.Text
'O Banco %s nao est� cadastrado.
Public Const ERRO_ATUALIZACAO_CARTEIRASCOBRADOR = 6298 'Parametro objCarteiraCobrador.iCodCarteiraCobranca
'Erro de atualiza�ao da Carteira %i
Public Const ERRO_LEITURA_TABELA_CARTEIRASCOBRADOR1 = 6299 'Sem parametros
'Erro na leitura da tabela de Carteiras de Cobrador.
Public Const ERRO_INSERCAO_CARTEIRASCOBRADOR = 6300 'Parametro objCarteiraCobrador.iCodCarteiraCobranca
'Erro de inser��o na Carteira %i
Public Const ERRO_LEITURA_TABELA_CARTEIRASCOBRADOR = 6301 'Parametro objcarteiracobrador.icobrador
'Erro de leitura na tabela carteiras cobrador para o cobrador %s
Public Const ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA = 6302 'Parametro objCarteiraCobranca.iCodigo
'Carteira Cobran�a n�o est� cadastrada.
Public Const ERRO_LEITURA_BAIXASPARCREC = 6303
'Erro na leitura da tabela BaixasParcRec
Public Const ERRO_BANCO_NAO_INFORMADO = 6304
'O Banco deve ser informado.
Public Const ERRO_CARTEIRA_COBANCA_NAO_INFORMADA = 6305
'A Carteira de Cobran�a deve ser informada.
Public Const ERRO_LEITURA_BORDERO_COBRANCA = 6306
'Erro na leitura da tabela de Bordero de Cobran�a.
Public Const ERRO_CARTEIRACOBRANCA_VINCULADA_PARCELAS = 6307 'iCarteira
'A Carteira cobran�a %i n�o pode ser excluida por est� sendo utilizada por alguma parcela a receber.
Public Const ERRO_LOCK_CARTEIRASCOBRADOR = 6308
'Erro an tentativa de fazer "lock" na tabela CarteirasCobrador.
Public Const ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL = 6309 'Parametro: sContaCorrente, iFilial
'A Conta Corrente %s n�o pertence � Filial selecionada. Filial: %i.
Public Const ERRO_TIPO_COBRANCA_NAO_INFORMADO = 6310 'Sem Par�metro
'O Tipo de cobran�a deve ser informado.
Public Const ERRO_NUMBORDERO_NAO_INFORMADO = 6311 'Sem Par�metro
'O N�mero do Bordero deve ser informado.
Public Const ERRO_LEITURA_TIPOSDECOBRANCA = 6312 'Sem Par�metro
'Erro na leitura da tabela TiposDeCobran�a.
Public Const ERRO_TIPOCOBRANCA_NAO_VALE_BORDERO = 6313 'Par�metro : sTipoCobranca
'O Tipo de cobran�a %s n�o vale para Bordero.
Public Const ERRO_TIPOCOBRANCA_INEXISTENTE = 6314 'Par�metro : sTipoCobranca
'O Tipo de cobran�a %s n�o est� cadastrado.
Public Const ERRO_CARTEIRA_COBRANCA_NAO_INFORMADA = 6315 'Sem parametros
'Codigo n�o foi informado
Public Const ERRO_LEITURA_CARTEIRASCOBRADOR = 6316
'Erro na leitura da tabela de Carteiras do Cobrador.
Public Const ERRO_CARTEIRA_COM_COBRADOR = 6317 'Parametro iCodigo
'N�o foi poss�vel excluir carteira  %i , existem um ou mais cobradores associados a mesma .
Public Const ERRO_CODIGOCARTCOBR_NAO_PREENCHIDO = 6318 'Sem parametros
'O C�digo da Carteira de Cobran�a n�o foi informado
Public Const ERRO_LEITURA_NFS_TRANSPORTADORA = 6319
'Erro de leitura na tabela de Notas Fiscais.
Public Const ERRO_EXISTEM_NFS_TRANSPORTADORA = 6320
'Existe Nota Fiscal para esta transportadora.
Public Const ERRO_TRANSPORTADORA_RELACIONADA_FILIAISCLIENTES = 6321 'Parametro: lTotal
'Transportadora esta relacionada com %l Filiais de Clientes
Public Const ERRO_CODTRANSPORTADORA_NAO_PREENCHIDO = 6322
'O c�digo da Transportadora n�o foi preenchido.
Public Const ERRO_INSERCAO_TRANSPORTADORA = 6323 'Sem parametro
'Erro na inser��o da Transportadora.
Public Const ERRO_MODIFICACAO_TRANSPORTADORA = 6324 'Sem parametro
'Erro  na modifica��o da Transportadora.
Public Const ERRO_ENDERECO_NAO_CADASTRADO = 6325 'Sem parametro
'O Endereco  n�o est� cadastrado no Banco de Dados.
Public Const ERRO_EXCLUSAO_TRANSPORTADORA = 6326 'Parametro codigo da Transportadora
'Erro  na tentativa de excluir Transportadora.
Public Const ERRO_TRANSPORTADORA_REL_NF_REC_PEND = 6327 'Parametro: iCodigo
'Erro na exclus�o da Transportadora com c�digo %l, relacionado com Nota Fiscal a Receber Pendente.
Public Const ERRO_TRANSPORTADORA_REL_NF_REC = 6328 'Parametro: iCodigo
'Erro na exclus�o da Transportadora com c�digo %l, relacionado com Nota Fiscal a Receber.
Public Const ERRO_TRANSPORTADORA_REL_NF_REC_BAIXADA = 6329 'Parametro: iCodigo
'Erro na exclus�o da Transportadora com c�digo %l, relacionado com Nota Fiscal a Receber Baixada.
Public Const ERRO_CONTA_INEXISTENTE = 6330 'Parametro sContaCont�bil
'A Conta Contabil %s nao existe.
Public Const ERRO_ATUALIZACAO_CARTEIRASCOBRANCA = 6331 'Parametro: C�digo
'Erro na atualiza��o da Carteira Cobran�a %s.
Public Const ERRO_INSERCAO_CARTEIRASCOBRANCA = 6332 'Parametro: C�digo
'Erro na inclus�o da Carteira Cobran�a %s.
Public Const ERRO_CARTEIRACOBRANCA_NAO_CADASTRADO = 6333 'Parametro: C�digo
'A Carteira Cobran�a %s n�o est� cadastrada.
Public Const ERRO_LOCK_CARTEIRASCOBRANCA = 6334 'Parametro: C�digo
'Erro na tentativa de fazer Lock da Carteira Cobran�a %s.
Public Const ERRO_EXCLUSAO_CARTEIRASCOBRANCA = 6335 'Parametro: C�digo
'Erro na exclus�o da Carteira Cobran�a %s.
Public Const ERRO_NUMERO_PARCELAS_ALTERADO = 6336 'Par�metros: iNumParcelasTela, iNumParcelasBD
'N�o � poss�vel alterar o n�mero de parcelas de um T�tulo lan�ado. Na Tela: %i. No Banco de Dados: %i.
Public Const ERRO_LOCK_PARCELAS_REC = 6337
'Erro na tentativa de "lock" na Tabela de Parcelas a Receber.
Public Const ERRO_LOCK_COMISSOES = 6338 'Parametro: lNumIntDoc, iTipoTitulo
'Erro na tentativa de fazer lock nas Comiss�es do Documento de Tipo %i e N�mero Interno %l na Tabela de Comiss�es
Public Const ERRO_LEITURA_PARCELASREC = 6339 'Sem parametro
'Erro na tentativa de ler registro na tabela ParcelasRec.
Public Const ERRO_ATUALIZACAO_PARCELASREC = 6340 'Parametro: lNumIntParc
'Erro na Atualizacao da Parcela %s do T�tulo %s da tabela de ParcelasRec.
Public Const ERRO_LEITURA_CHEQUEPRE = 6341 'Parametro: lNumIntCheque
'Erro na leitura da tabela de ChequePre com N�mero %l.
Public Const ERRO_EXCLUSAO_CHEQUESPRE = 6342 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'Erro na tentativa de excluir o ChequePre com Banco %i, Ag�ncia %s, ContaCorrente %s e N�mero %l da tabela de ChequesPre.
Public Const ERRO_INSERCAO_PARCELAS_REC = 6343
'Erro na inser��o de um registro na tabela de Parcelas a Receber.
Public Const ERRO_LOCK_TITULOS_REC = 6344
'Erro na tentativa de "lock" na Tabela T�tulos a Receber.
Public Const ERRO_EXCLUSAO_TITULOS_RECEBER = 6345
'Erro na exclus�o de um registro da tabela de T�tulos a Receber.
Public Const ERRO_PARCELA_COM_BAIXA = 6346 'Par�metro: iNumParcela, lNumTitulo
'Erro na exclus�o da Parcela %i do T�tulo %l. A Parcela tem baixa.
Public Const ERRO_INSERCAO_TITULOS_REC = 6347
'Erro na inser��o de um registro na tabela de T�tulos a Receber.
Public Const ERRO_LEITURA_TIPOS_INSTRUCAO_COBRANCA = 6348 'Sem Parametros
'Erro na leitura da tabela TipoInstrCobranca
Public Const ERRO_LOCK_INSTRUCAO_COBRANCA = 6349 'Parametro: iInstrucao
'N�o conseguiu fazer o lock da Instru��o de Cobran�a %i.
Public Const ERRO_INSERCAO_CHEQUESPRE = 6350 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'Erro na inser��o de ChequePre na tabela de ChequesPre com Banco %i, Ag�ncia %s, ContaCorrente %s e N�mero %l j� foi utilizado.
Public Const ERRO_LEITURA_TABELA_CARTEIRASCOBRANCA = 6351 'Sem parametro
'Erro na leitura da tabela de Carteiras de Cobran�a.
Public Const ERRO_LEITURA_TIPO_INSTRUCAO_COBRANCA = 6352 'Parametro iCodigo
'Erro na leitura do Tipo Instru��o Cobran�a %i.
Public Const ERRO_LEITURA_NFISCAL2 = 6354 'Par�metros: lCodFornecedor, iCodFilial
'Erro na leitura da tabela NFiscal com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_TIPODOCUMENTO1 = 6355 'Sem par�metros
'Erro na leitura da tabela de TiposdeDocumento.
Public Const ERRO_LOCK_CANALVENDA = 6357 'Par�metro: iCanalVenda
'N�o foi poss�vel fazer o Lock do Cana de Venda %i da tabela CanalVenda.
Public Const ERRO_LEITURA_SERIE1 = 6358 'Par�metro: sSerie
'Erro de leitura da tabela Serie com S�rie %s.
Public Const ERRO_LOCK_SERIE1 = 6359 'Par�metro: sSerie
'N�o foi poss�vel fazer o Lock da S�rie %s da tabela Serie.
Public Const ERRO_LOCK_TRANSPORTADORA1 = 6360 'Par�metro: iTransportadora
'N�o foi poss�vel fazer o Lock da Transportadora %i da tabela Transportadoras.
Public Const ERRO_LEITURA_PARCELAS_REC_NF = 6361 'parametro: lNumNotaFiscal
'Ocorreu um erro na leitura das parcelas a receber vinculadas � nota fiscal %l.
Public Const ERRO_LEITURA_PARCELAS_REC_BAIXADAS_NF = 6362 'parametro: lNumNotaFiscal
'Ocorreu um erro na leitura das parcelas a receber baixadas vinculadas � nota fiscal %l.
Public Const ERRO_LEITURA_TIPOSDECLIENTE2 = 6363 'Par�metro: iCodigo
'Erro de leitura da tabela TiposDeCliente com Vendedor %i.
Public Const ERRO_VENDEDOR_REL_TIPOCLIENTE = 6364 'Par�metros: iCodigo, iTipoCliente
'O Vendedor %i est� relacionado com o Tipo de Cliente %i.
Public Const ERRO_DATA_SAIDA_MENOR_DATA_EMISSAO = 6365 'Par�metro: dtDataSaida, dtDataEmissao
'A Data de Sa�da %dt � anterior a Data de Emiss�o %dt.
Public Const ERRO_TIPODOC_DIFERENTE_NF_VENDA = 6366 'iTipoDocInfo
'Tipo de Documento %i n�o � Nota Fiscal de Venda.
Public Const ERRO_FALTA_LOCALIZACAO = 6367 'Par�metro: sProduto
'O Produto %s n�o foi localizado. N�o � poss�vel gravar a Nota Fiscal.
Public Const ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL = 6368 'Par�metros: iItem, dDesconto, dPrecoTotal
'Para o Item %i o Desconto %d � maior ou igual ao Pre�o %d.
Public Const ERRO_VALOR_DESCONTO_ULTRAPASSOU_SOMA_VALORES = 6369 'dDesconto, dSomaValores
'Desconto = %d n�o pode ultrapassar a soma de Produtos + Frete + Seguro + Despesas = %d.
Public Const ERRO_LOCALIZACAO_ITEM_INEXISTENTE = 6370 'iItem
'N�o foi feita a localiza��o do item %i da Nota Fiscal.
Public Const ERRO_LOCALIZACAO_ITEM_INCOMPLETA = 6371 'iItem, dQuantVendida, dQuantAlocada
'A localiza��o do item %i da Nota Fiscal est� incompleta. Quantidade vendida: %d. Quantidade localizada: %d.
Public Const ERRO_DATASAIDA_ANTERIOR_DATAEMISSAO = 6372 'dtDataSaida, dtDataEmissao)
'A Data de Saida %dt � anterior a Data de Emiss�o %dt.
Public Const ERRO_INSERCAO_NFISCAL_SAIDA = 6373 'Par�metro: lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal
'Erro na inser��o da Nota Fiscal com os dados C�digo do Cliente =%l, C�digo da Filial =%i, Tipo =%i, Serie =%s e N�mero NF =%l na tabela de Notas Fiscais.
Public Const ERRO_ALTERACAO_NFISCAL_SAIDA = 6374 'Par�metros: lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal com os dados C�digo do Cliente = %l, C�digo da Filial = %i, Tipo = %i, Serie = %s, N�mero NF = %l, Data Emissao = %dt est� cadastrada no Banco de Dados. N�o � poss�vel alterar.
Public Const ERRO_LEITURA_NFISCAL_SAIDA_BAIXADA = 6375 'Par�metros: iTipoNFiscal, lCliente, iFilialCli, sSerie, lNumNotaFiscal
'Erro na leitura da tabela NFiscalBaixadas na Nota Fiscal com Tipo = %i, Cliente = %l, Filial = %i, Serie = %s e N�mero = %l.
Public Const ERRO_LEITURA_NFISCAL3 = 6376 'Sem parametros
'Ocorreu um erro na leitura da tabela de Notas Fiscais.
Public Const ERRO_CHEQUEPRE_PARCELAREC_INVALIDO = 6377 'parametro: iParcela
'A parcela %i n�o pode ser associada a um cheque-pre, pois j� est� vinculada a outra carteira cobrador.
Public Const ERRO_INSERCAO_TRANSFCARTCOBR = 6378 'sem parametros
'Erro na inser��o de registro de transfer�ncia de carteira de cobran�a
Public Const ERRO_CARTEIRA_COBRADOR_SALDO_NEG = 6379 'parametros: codigo da carteira de cobranca e do cobrador
'A carteira %i do cobrador %i n�o pode ter saldo negativo em valor
Public Const ERRO_CARTEIRA_COBRADOR_QTDE_NEG = 6380 'parametros: codigo da carteira de cobranca e do cobrador
'A carteira %i do cobrador %i n�o pode ter n�mero negativo de t�tulos
Public Const ERRO_LEITURA_CARTEIRA_COBRADOR = 6381 'parametros: codigo da carteira de cobranca e do cobrador
'Erro na leitura da carteira %i do cobrador %i
Public Const ERRO_LEITURA_TIPOCLIENTECATEGORIAS = 6383 'Sem Par�metros
'Erro de Leitura na Tabela TipoDeClienteCategorias.
Public Const ERRO_CATEGORIACLIENTEITEM_TIPOCLIENTECATEGORIAS = 6386 'Par�metros: CategoriaCliente e CategoriaClienteItem
'Categoria Cliente %s e Categoria Cliente Item %s s�o usados na tabela TipoDeClienteCategorias.
Public Const ERRO_CATEGORIACLIENTE_ICMSEXCECOES = 6387 'Par�metro: CategoriaCliente
'Categoria Cliente %s � usada na Tabela ICMSExcecoes.
Public Const ERRO_CATEGORIACLIENTE_IPIEXCECOES = 6388 'Par�metro: CategoriaCliente
'Categoria Cliente %s � usada na Tabela IPIExcecoes.
Public Const ERRO_CATEGORIACLIENTE_TIPOCLIENTECATEGORIAS = 6389 'Par�metro: CategoriaCliente
'Categoria Cliente %s � usada na Tabela TipoDeClienteCategorias.
Public Const ERRO_CATEGORIACLIENTE_FILIALCLIENTECATEGORIAS = 6390 'Par�metro: CategoriaCliente
'Categoria Cliente %s � usada na tabela FilialClienteCategorias.
Public Const ERRO_LEITURA_PEDIDODEVENDAS1 = 6391 'Sem Par�metros
'Erro de Leitura na Tabela PedidoDeVendas.
Public Const ERRO_TRANSPORTADORA_PEDIDODEVENDAS = 6392 'Par�metro: C�digo da Transportadora
'Transportadora %i � utilizada na tabela PedidoDeVendas.
Public Const ERRO_LEITURA_PEDIDODEVENDASBAIXADOS = 6393 'Sem Par�metros
'Erro de Leitura na Tabela PedidoDeVendasBaixados.
Public Const ERRO_TRANSPORTADORA_PEDIDODEVENDASBAIXADOS = 6394 'Par�metro: C�digo da Transportadora
'Transportadora %i � utilizada na tabela PedidoDeVendasBaixados.
Public Const ERRO_TRANSPORTADORA_FILIAISCLIENTES = 6395 'Par�metro: C�digo da Transportadora
'Transportadora %i � utilizada na tabela FiliaisClientes
Public Const ERRO_MODIFICACAO_CARTEIRAS_COBRADOR = 6396
'Erro na atualiza��o na tabela Carteiras Cobrador.
Public Const ERRO_EXCLUSAO_TIPODECLIENTECATEGORIAS = 6397 'Par�metro: iCodigo
'Erro na tentativa de excluir a Categoria do Tipo de Cliente %i da tabela TiposDeCliente.
Public Const ERRO_INSERCAO_TIPOSDECLIENTE = 6398 'Par�metro: iCodigo
'Erro na tentativa de inserir um registro na tabela TiposDeCliente com C�digo %i.
Public Const ERRO_PADRAO_COBRANCA_NAO_CADASTRADO1 = 6399 'Par�metro: iCodigo
'O Padr�o de Cobran�a %1 n�o est� cadastrado no Banco de Dados.
Public Const ERRO_INSERCAO_TIPODECLIENTECATEGORIAs = 6400 'Parametro: iCodigo
'Erro na tentativa de inserir o Tipo de Cliente %i na tabela TipoDeClienteCategorias.
Public Const ERRO_MODIFICACAO_TIPOSDECLIENTE = 6401 'Par�metro: iCodigo
'Erro na modifica��o da tabela TiposDeCliente com C�digo %i.
Public Const ERRO_MODIFICACAO_TIPODECLIENTECATEGORIAS = 6402 'Par�metros: sCategoria, sItem
'Erro na modifica��o da tabela TipoDeClienteCategorias com Categoria %s e �tem %i.
Public Const ERRO_PADRAO_COBRANCA_NAO_CADASTRADA = 6406 'Par�metro: .iCodigo
'O Padr�o de Cobran�a %i n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_PARCELASREC_NUMINT = 6407 'Par�metros: lNumIntTitulo, iNumParcela
'Erro na leitura da tabela ParcelasRec com n�mero interno do t�tulo %l e n�mero da Parcela %i.
Public Const ERRO_LEITURA_TIPOSINSTRCOBRANCA = 6408 'Sem par�metros
'Erro na leitura da tabela TiposInstrCobranca.
Public Const ERRO_LEITURA_TITULOSREC2 = 6409 'Par�metros: iFilialEmpresa, lCliente, iFilial, sSiglaDocumento, lNumTitulo
'Erro na tentativa de leitura do Titulo Receber da Filial Empresa %i, Cliente %l , Filial %i, Sigla do Documento %s e N�mero do T�tulo %l na tabela TitulosRec.
Public Const ERRO_VALOR_COMISSAO_GRID_NAO_PREENCHIDO = 6410 'Par�metro: iComissao
'O Valor da comissao %i do Grid de Comiss�es n�o est� preenchido.
Public Const ERRO_TIPODOC_DIFERENTE_NF_FATURA = 6411 'Par�metro: iTipoDocInfo
'Tipo de Documento %i n�o � Nota Fiscal Fatura.
Public Const ERRO_DATA_DESCONTO_INFERIOR_DATA_EMISSAO = 6412 'dtDataDesconto
'Data do Desconto = %dt n�o pode ser inferior � Data de Emiss�o da Nota Fiscal.
Public Const ERRO_DATA_DESCONTO_SUPERIOR_DATA_VENCIMENTO = 6413 'dtDataDesconto
'Data de Desconto = %dt n�o pode ser superior � data de Vencimento
Public Const ERRO_ALTERACAO_NFISCAL_INTERNA = 6414 'sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal Interna com os dados S�rie = %s , N�mero = %l, Data Emiss�o = %dt est� cadastrada no Banco de Dados.
Public Const ERRO_PADRAO_COBRANCA_RELACIONADO_COM_CLIENTE = 6415
'O Padr�o de Cobran�a n�o pode ser exclu�do pois est� relacionado com um cliente
Public Const ERRO_LOCK_ALCADAFAT = 6416 'Parametro: sCodUsuario
'Ocorreu um erro na tentativa de fazer um lock de um registro da tabela de Al�ada Fat. Usu�rio = %s.
Public Const ERRO_LOCK_VALORLIBERADOCREDITO = 6417 'Parametros: sCodUsuario, iAno
'Ocorreu um erro ao tentar fazer o "lock" de um registro da tabela ValorLiberadoCredito. Usu�rio = %s, Ano = %i.
'Public Const ERRO_TABELAPRECO_UTILIZADA_TABELAPRECOITEM = 6418 'Par�metro: iCodTabela
'A Tabela %i n�o pode ser excluida pois est� sendo utilizada em outras Filiais.
'Public Const ERRO_TABELAPRECO_UTILIZADA_NOTAS_FICAIS = 6419 'Par�metro: iCodTabela
'A Tabela %i n�o pode ser excluida pois est� sendo utilizada em Notas Fiscais.
'Public Const ERRO_TABELAPRECO_UTILIZADA_PEDIDOS_VENDA = 6420 'Par�metro: iCodTabela
'A Tabela %i n�o pode ser excluida pois est� sendo utilizada em Pedidos de Venda.
Public Const ERRO_LEITURA_NOTAS_FISCAIS = 6421 'Sem Par�metros
'Erro na leitura de Notas Fiscais.
Public Const ERRO_LEITURA_PEDIDOS_VENDA = 6422 'Sem Par�metros
'Erro na leitura de Pedidos de Venda.
Public Const ERRO_LEITURA_TABELASDEPRECO = 6423 'Sem parametros
'Erro na leitura da tabela de Tabelas de Pre�o.
Public Const ERRO_ITEM_NAO_CADASTRADO = 6424 'Parametros: iCodTabela, sCodProduto
'�tem de Tabela de Pre�o n�o est� cadastrado no Banco de Dados. C�digo da Tabela %i e C�digo do Produto %s.
Public Const ERRO_TABELAPRECO_NAO_PREENCHIDA = 6425 'Sem parametros
'A Tabela deve estar preenchida.
Public Const ERRO_CLIENTE_TABELAPRECO = 6426 'Parametro: iCodigo
'N�o � poss�vel excluir a Tabela de Pre�o %i. Est� associada a Clientes.
Public Const ERRO_LOCK_TABELASDEPRECOITENS1 = 6427 'Parametro: iCodigo
'N�o conseguiu fazer o lock na tabela de TabelasDePrecoItens com C�digo da Tabela %i.
Public Const ERRO_EXCLUSAO_TABELASDEPRECO = 6428 'Parametro: iCodigo
'Erro na tentativa de excluir a Tabela de Pre�o com c�digo %i da tabela de TabelasDePrecoItens.
Public Const ERRO_FILIALCLIENTE_NAO_CADASTRADA2 = 6430 'Parametros: sNomeRedCliente, iCodFilial
'A Filial %i do Cliente %s n�o est� cadastrada no Banco de Dados.
Public Const ERRO_ATUALIZACAO_FILIALCLIENTEFILEMP = 6431 'Sem Parametros
'Erro na Atualiza��o da tabela FilialClienteFilEmp.
Public Const ERRO_FILIALCLIENTEFILEMP_NAO_CADASTRADA = 6432 'Parametros iFilialEmpresa,lCodCliente,iCodFilial
'O cliente %l, Filial %i com a filialEmpresa %i n�o est�o cadastrado na tabela FilialClienteFilEmp.
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
'N�o foi encontrado registro em CRFATConfig. Codigo = %s Filial = %i
Public Const ERRO_LEITURA_CRFATCONFIG2 = 6440 '%s chave %d FilialEmpresa
'Erro na leitura da tabela CRFATConfig. Codigo = %s Filial = %i
Public Const ERRO_LEITURA_CRFATCONFIG1 = 6441 'Parametros: sCodigo, iFilial.
'Ocorreu um erro na leitura do codigo %s da Filial %i na tabela CRFATConfig.
Public Const ERRO_DATA_EMISSAO_NAO_PREENCHIDA = 6442 'Sem par�metros
'� obrigat�rio o preenchimento da Data de Emiss�o.
Public Const ERRO_ATUALIZACAO_FILIALFORNFILEMP = 6443 'Sem Parametros
'Erro na Atualiza��o da tabela FilialFornFilEmp.
Public Const ERRO_FILIALFORNFILEMP_NAO_CADASTRADA = 6444 'Parametros iFilialEmpresa,lCodFornecedor,iCodFilial
'O Fornecedor %l, Filial %i com a filialEmpresa %i n�o est�o cadastrado na tabela FilialFornFilEmp.
Public Const ERRO_LEITURA_FILIALFORNFILEMP = 6445 'Sem Parametro
'Erro na leitura da Tabela FilialFornecedorFilEmp.
Public Const ERRO_LOCK_FILIALFORNFILEMP = 6446 'Sem Parametros
'Erro na tentativa de fazer lock na tabela FilialFornFilEmp.
Public Const ERRO_INSERIR_FILIALFORNFILEMP = 6447 'Sem Parametros
'Erro na tentativa de inserir na tabela FilialFornFilEmp.
'Public Const ERRO_TABELAPRECO_UTILIZADA_TIPOCLIENTE = 6448 'Par�metro: iCodTabela
'A Tabela %i n�o pode ser excluida pois est� sendo utilizada em Tipos de Cliente.
Public Const ERRO_ALCADA_NAO_CADASTRADA2 = 6449 'Parametro: C�digo do Usu�rio
'A al�ada do usu�rio %s n�o est� cadastrada.
Public Const ERRO_TIPO_CLIENTE_NAO_ENCONTRADO2 = 6450 'Parametro sTipoCliente
'O tipo de cliente %s n�o foi encontrado.
Public Const ERRO_TIPO_FORNECEDOR_NAO_ENCONTRADO2 = 6451 'Parametro sTipoFornecedor
'O tipo de fornecedor %s n�o foi encontrado.
Public Const ERRO_TIPO_VENDEDOR_NAO_ENCONTRADO2 = 6452 'Parametro sTipoVendedor
'O tipo de vendedor %s n�o foi encontrado.
Public Const ERRO_CONTA_CORRENTE_NAO_ENCONTRADA = 6453 'Parametro: sContaCorrente
'A conta corrente %s n�o foi encontrada.
Public Const ERRO_TRANSP_NOME_RED_DUPLICADO = 6454 'Parametro: iCodigo
'A transportadora %i tem o mesmo Nome Reduzido.
Public Const ERRO_TRANSPORTADORA_TIPOCLIENTE = 6455 'Parametro: iCodigoTipo
'Erro a Transportadora %i � utilizada na tabela de Tipo de Cliente.
Public Const ERRO_COND_PAGAMENTO_TITULOSPAG = 6456 'Parametros: iCodigo
'Erro a Condi��o de Pagamento %i � utilizada na tabela TitulosPag.
Public Const ERRO_COND_PAGAMENTO_TITULOSPAG_BAIXADOS = 6457 'Parametros: iCodigo
'Erro a Condi��o de Pagamento %i � utilizada na tabela TitulosPagBaixados.
Public Const ERRO_CATEGORIACLIENTE_PADROESTRIBSAIDA = 6458  'Parametros: sCategoriaCliente
'Erro a Categoria Cliente %s foi utilizada na tabela de PadroesTribSaida.
Public Const ERRO_CATEGORIACLIENTEITEM_PADROESTRIBSAIDA = 6459  'Parametros: sCategoria , sItemCategoria
'Categoria Cliente %s e Categoria Cliente Item %s s�o usados na tabela de PadroesTribSaida.
Public Const ERRO_FILIALCLIENTE_CHEQUEPRE = 6460 'Parametros: iFilialCliente,lCodigoCliente
'Erro na Exclus�o da Filial %i do Cliente %l que est� associado com Cheque Pr�.
Public Const ERRO_COND_PAGAMENTO_TITULOSREC = 6461 'Parametros: iCodigo
'Erro a Condi��o de Pagamento %i � utilizada na tabela TitulosRec.
Public Const ERRO_COND_PAGAMENTO_TITULOSREC_BAIXADOS = 6462 'Parametros: iCodigo
'Erro a Condi��o de Pagamento %i � utilizada na tabela TitulosRecBaixados.
Public Const ERRO_TABELAPRECO_RELACIONADA_CLIENTE = 6463 'Parametros iTabelaPreco, lCodigoCliente
'N�o � poss�vel excluir a Tabela de Pre�o %i pois est� sendo utilizada pelo Cliente %l.
Public Const ERRO_TABELAPRECO_RELACIONADA_NFISCAL = 6464 'Parametros iTabelaPreco, sSerie, lNumeroNF, iFilialNF
'N�o � poss�vel excluir a Tabela de Pre�o %i pois est� sendo utilizada pela Nota Fiscal: S�rie = %s, Numero = %l e FilialEmpresa = %i.
Public Const ERRO_TABELAPRECO_RELACIONADA_TIPOSDECLIENTE = 6465 'Parametros iTabelaPreco, iCodigoTipo
'N�o � poss�vel excluir a Tabela de Pre�o %i pois est� sendo utilizada por Tipo de Cliente: Codigo = %i.
Public Const ERRO_TABELAPRECO_RELACIONADA_PEDVENDA = 6466 'Parametros iTabelaPreco, lCodigoPV, iFilialEmpresaPV
'N�o � poss�vel excluir a Tabela de Pre�o %i pois est� sendo utilizada pelo Pedido de Venda: Codigo = %l e FilialEmpresa = %i.
Public Const ERRO_TABELAPRECO_RELACIONADA_PEDVENDA_BAIXADO = 6467 'Parametros iTabelaPreco, lCodigoPVBaixado, iFilialEmpresaPVBaixado
'N�o � poss�vel excluir a Tabela de Pre�o %i pois est� sendo utilizada pelo Pedido de Venda Baixado: Codigo = %l e FilialEmpresa = %i.
Public Const ERRO_COMISSOES_BAIXADA_NFISCAL = 6468 'Parametros: sSerie, lNumNota, iFilialEmpresa
'Erro na tentativa de excluir registro da tabela de Comiss�es, da Nota: S�rie = %s, N�mero = %l e FilialEmpresa %i. A comiss�o j� foi baixada.
Public Const ERRO_COMISSOES_BAIXADA_PARCELA = 6469 'Parametros: lNumeroT�tulo, iParcela, iFilialEmpresa
'Erro na tentativa de excluir registro da tabela de Comiss�es, da Parcela: N�mero do T�tulo = %l, N�mero da Parcela = %i e FilialEmpresa %i. A comiss�o j� foi baixada.
Public Const ERRO_COMISSOES_BAIXADA_DEBITOS = 6470  'Parametros : lNumTitulo, sSiglaDocumento, lCliente,iFilial
'Erro na tentativa de excluir registro da tabela de Comiss�es, do D�bito: N�mero do T�tulo = %l, Sigla do Documento = %s, C�digo da Cliente = %l Filial do Cliente = %i. A comiss�o j� foi baixada.
Public Const ERRO_COMISSOES_BAIXADA_TITULO = 6471  'Parametros: lNumTitulo, iFilialEmpresa
'Erro na tentativa de excluir registro da tabela de Comiss�es, do T�tulo : N�mero do T�tulo = %l e FilialEmpresa %i. A comiss�o j� foi baixada.
Public Const ERRO_NOTAFISCAL_NAO_CADASTRADO_COMISSOES = 6472 'Sem Parametros
'A Nota Fiscal associada a esta comiss�o n�o foi encontrada.
Public Const ERRO_PARCELAREC_NAO_CADASTRADO_COMISSOES = 6473 'Sem Parametros
'A Parcela associada a esta Comiss�o n�o foi encontrada.
Public Const ERRO_TITULOREC_NAO_CADASTRADO_COMISSOES = 6474 'Sem Parametros
'O Titulo a Receber associado a esta Comiss�o n�o foi encontrado.
Public Const ERRO_DEBITOREC_NAO_CADASTRADO_COMISSOES = 6475 'Sem Parametros
'O D�bito a Receber associado a esta Comiss�o n�o foi encontrado.
Public Const ERRO_TRANSF_MANUAL_COBR_ELETRONICA = 6476 'sem parametro
'N�o pode transferir t�tulo de/para cobrador com cobran�a eletr�nica
Public Const ERRO_ALTERE_VCTO_INSTR_COB_ELETR = 6477 'sem parametros
'Para alterar o vencimento de uma parcela em cobran�a eletr�nica use a tela de instru��es para cobran�a eletr�nica
Public Const ERRO_LEITURA_CHEQUEPRE3 = 6478 'sem parametros
'Erro na leitura da tabela de cheques pr�-datados
Public Const ERRO_CHEQUEPRE_DEPOSITADO = 6479 'sem parametros
'Este cheque pr�-datado j� foi depositado
Public Const ERRO_PARCREC_DE_CARTEIRA_PARA_CHEQUEPRE = 6480
'Para associar um cheque pr� datado a uma parcela esta deve estar em carteira.
Public Const ERRO_PARCELAREC_OUTRO_CHEQUEPRE = 6481 'parametros: banco ag, cta e num do cheque
'Esta parcela j� est� associada ao cheque pr� identificado por: banco %s, agencia %s, conta %s e n�mero %s. Pode-se excluir o cheque anterior e cadastrar um novo.
Public Const ERRO_CHEQUEPRE_REPETIDO = 6482 'sem parametros
'O mesmo cheque pr�-datado est� associado a mais de uma parcela
Public Const ERRO_CHEQUEPRE_DUPLICADO = 6483 'parametro: codigo de cliente
'Este cheque pr�-datado j� est� registrado para o cliente com c�digo %s.
Public Const ERRO_CHEQUEPRE_OUTROTITULO = 6484 'sem parametros
'Este cheque pr�-datado est� associado a outro t�tulo j� registrado
Public Const ERRO_PARCELA_COBRANCA_EMPRESA = 6485 'Par�metro: lNumTitulo, iNumParcela
'Erro na exclus�o do Titulo %l. A parcela %i n�o est� em cobran�a na pr�pria Empresa.
Public Const ERRO_ALTERACAO_CARTEIRA_EMPRESA = 6486 'sem parametros
'N�o se pode incluir ou alterar uma carteira de cobran�a de uso restrito � pr�pria Empresa.
Public Const ERRO_NFISCALINTERNA_COM_NUMERO = 6487 'Sem Parametros
'N�o � possivel gravar uma Nota Fiscal Interna com seu n�mero preenchido.
'SOLUCAO: Limpe o campo N�mero com o bot�o ao lado, pois o sistema ir� gerar o N�mero na grava��o.
Public Const ERRO_NFISCAL_COM_NUMERO = 6488 'Sem Parametros
'N�o � possivel gravar uma Nota Fiscal com seu n�mero preenchido.
'SOLUCAO: Limpe o campo N�mero com o bot�o ao lado, pois o sistema ir� gerar o N�mero na grava��o.
Public Const ERRO_NOTA_FISCAL_CANCELADA = 6489 'Par�metro: sSerieNF, lNumNotaFiscal
'A Nota Fiscal com a s�rie %s e n�mero %l j� est� cancelada.
Public Const ERRO_NFISCAL_VINCULADA_CANCELAR = 6490 'Par�metro: lNumNotaFiscalOrig, iTipoNFOrig
'A Nota Fiscal em quest�o n�o pode ser cancelada pois est� vinculada a
'Nota Fiscal n�mero %l e tipo %i que n�o est� cancelada.
Public Const ERRO_ALTERACAO_ITEMNF = 6491
'Erro na tentativa de alterar os dados de um item de nota fiscal.
Public Const ERRO_NF_JA_DEVOLVIDA = 6492
'Ja existe uma nota fiscal de devolu��o para a nota em quest�o.
Public Const ERRO_NFISCAL_OUTRA_FILIAL = 6493
'A nota fiscal em quest�o percente a outra filialempresa.
Public Const ERRO_NFISCAL_SEM_ITENS = 6494 'lCodPedido
'A Nota Fiscal com o C�digo %l n�o possui �tens associados � ele.
Public Const ERRO_REGIAO_INICIAL_MAIOR = 6495 'Sem Parametros
'A Regi�o inicial n�o pode ser maior que a Regi�o final.
Public Const ERRO_COBRADOR_INATIVO = 6496 'Par�metros: iCodCobrador
'O Cobrador de c�digo %i � inativo.
Public Const ERRO_PADRAO_COBRANCA_INATIVO = 6497 'Par�metros: iCodPadraoCobranca
'O Padr�o de Cobran�a de c�digo %i � inativo.
Public Const ERRO_CRIACAO_NFR_COM_FATURAMENTO = 6498
'A cria��o de um t�tulo de Nota Fatura a Receber ou Nota Fatura a Receber de Servi�o � criado automaticamente ap�s o cadastro da Nota Fiscal.
Public Const ERRO_LEITURA_TRANSPORTADORA2 = 6499 'Sem Parametros
'Erro na leitura da tabela de Transportadoras.
Public Const ERRO_LIMITE_TRANSP_VLIGHT = 6500 'Parametros : iNumeroMaxTransportadoras
'N�mero m�ximo de Transportadoras desta vers�o � %i.
Public Const ERRO_FATURA_ATE_IMPRESSAO_NAO_PREENCHIDO = 6501 'Sem Parametros
'� obrigat�rio informar at� que Fatura foi obtida boa Impress�o.
Public Const ERRO_FATURA_ATE_MAIOR_ANTERIOR = 6502 'Sem Parametros
'A Fatura At� deve ser menor ou igual ao que foi mandado para Impress�o.
Public Const ERRO_FATURA_ATE_MENOR_NUMERO_DE = 6503 'Sem Parametros
'A Fatura At� n�o pode ser menor que a Fatura De.
Public Const ERRO_NUMERO_ATE_IMPRESSAO_NAO_PREENCHIDO = 6504 'Sem Parametros
'� obrigat�rio informar at� que Nota Fiscal foi obtida boa Impress�o.
Public Const ERRO_UNLOCK_SERIE_IMPRESSAO_NF = 6505 'Parametros : sSerie
'A S�rie %s n�o est� lockada para Impress�o, por isso n�o pode ser feito unlock.
Public Const ERRO_NUMERO_ATE_MAIOR_ANTERIOR = 6506 'Sem Parametros
'O N�mero At� deve ser menor ou igual ao que foi mandado para Impress�o.
Public Const ERRO_NUMERO_ATE_MENOR_NUMERO_DE = 6507 'Sem Parametros
'O Numero At� n�o pode ser menor que o N�mero De.
Public Const ERRO_TIPO_FORMULARIO_IMCOMPATIVEL = 6508 'Parametro: sSerie
'O Tipo de Formul�rio da S�rie %s est� imcompativel.
Public Const ERRO_ATUALIZACAO_NFISCAL1 = 6509
'Erro na atualiza��o de registros na tabela de notas fiscais.
Public Const ERRO_LEITURA_ANTECIPRECS = 6510  'sem parametros
'Erro na leitura de adiantamentos de clientes
Public Const ERRO_FALTA_MESANO_ESTOQUE = 6511 'sem parametros
'Preencha o mes/ano para o m�dulo de estoque.



'VEIO DE ERROS CPR
Public Const ERRO_LEITURA_FILIAISFORNECEDORES = 2026
'Erro na leitura da tabela Filiais Fornecedores.
Public Const ERRO_LOCK_FILIAISFORNECEDORES = 2040 'Parametro Codigo Fornecedor
'Erro na tentativa de fazer "lock" na tabela de FiliaisFornecedores para Codigo Fornecedor = %l .
Public Const ERRO_LEITURA_TITULOS_PAGAR = 2042
'Erro na leitura da tabela Titulos a Pagar.
Public Const ERRO_FILIAL_FORNECEDOR_INEXISTENTE = 2047 'Parametro Codigo Filial Fornecedor e Codigo Fornecedor
'A Filial Fornecedor %s do Fornecedor %s , n�o esta cadastrada no Banco De Dados.
Public Const ERRO_TIPOCLIENTE_INEXISTENTE = 2099
'Esse Tipo de Cliente n�o existe , ou ent�o j� foi exclu�do.
Public Const ERRO_TIPOCLIENTE_DESCR_DUPLICATA = 2102
'A Descri��o %s j� esta sendo utilizada para outro Tipo De Cliente.
Public Const ERRO_EXCLUSAO_TIPOCLIENTE_RELACIONADO = 2105 'Parametro Codigo do Cliente
'Esse Tipo Cliente n�o pode ser exclu�do pois se encontra relacionado com o Cliente %l.
Public Const ERRO_LOCK_TIPOCLIENTE = 2106 'Parametro Codigo do Tipo de Cliente
'Erro na tentativa de fazer "lock" no Tipo de Cliente = %i na tabela TiposDeCliente.
Public Const ERRO_EXCLUSAO_TIPOCLIENTE = 2107 'Parametro Codigo do Tipo de Cliente
'Erro na tentativa de excluir o Tipo De Cliente = %i.
Public Const ERRO_LEITURA_PARCELAS_PAG1 = 2229
'Erro na leitura da tabela de Parcelas a Pagar e Titulos a Pagar.
Public Const ERRO_LEITURA_PARCELAS_REC = 2234
'Erro na leitura da tabela de Parcelas a Receber.
Public Const ERRO_LEITURA_FILIAISFORNECEDORES1 = 2292 'Par�metros : lFornecedor, iFilial
'Erro na leitura do Fornecedor %l e Filial %i na tabela FiliaisFornecedores.
Public Const ERRO_LOCK_FILIAISFORNECEDORES1 = 2294 'Parametros: lFornecedor e iFilial
'Erro na tentativa de fazer "lock" na tabela de FiliaisFornecedores para Fornecedor = %l  e Filial = %i
Public Const ERRO_LEITURA_NFSPAG1 = 2299 'Parametro: lNumNotafiscal
'Erro na tentativa de leitura da Nota Fiscal n�mero %l na tabela NFsPag.
Public Const ERRO_ATUALIZACAO_NFSPAG = 2306 'Par�metro: lNumnotafiscal
'Erro na tentativa de atualiza��o da Nota Fiscal n�mero %l na tabela de NfsPag.
Public Const ERRO_LOCK_NFSPAG = 2316 'Parametro : lNumnotaFiscal
' Erro na tentativa de "lock" da tabela NfsPag na Nota Fiscal n�mero %l.
Public Const ERRO_NF_VINCULADA = 2317 'Par�metros: lNumNotafiscal, lNumTitulo
'N�o � poss�vel excluir a Nota Fiscal n�mero %l pois est� vinculada � Fatura %l.
Public Const ERRO_VENDEDOR_NAO_CADASTRADO1 = 2383 'Parametro: sNomeReduzido
'O Vendedor com Nome Reduzido %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_TITULOS_PAGAR_BAIXADO = 2466
'Erro na leitura da tabela de T�tulos a Pagar Baixados.
Public Const ERRO_NUMBORDERO_CHEQUEPRE_DEPOSITADO = 2478 'Sem parametros
'N�o se pode excluir um cheque pr�-datado que tenha sido depositado
Public Const ERRO_LEITURA_CHEQUEPRE1 = 2480 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'Erro na leitura da tabela de ChequePre com Banco %i, Ag�ncia %s, ContaCorrente %s e N�mero %l.
Public Const ERRO_CHEQUEPRE_JA_UTILIZADO = 2490 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'O ChequePre com Banco %i, Ag�ncia %s, ContaCorrente %s e N�mero %l j� foi utilizado.
Public Const ERRO_LOCK_CHEQUESPRE = 2491 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'Erro na tentativa de fazer "lock" na tabela de ChequesPre com Banco %i, Ag�ncia %s, ContaCorrente %s e N�mero %l.
Public Const ERRO_REFERENCIA_OUTRO_CHEQUE = 2492 'Parametro: lNumIntParc
'Esta Parcela %l faz refer�ncia a outro Cheque Pre.
Public Const ERRO_CHEQUEPRE_NAO_ENCONTRADO = 2498 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'O ChequePre com Banco %i, Ag�ncia %s, ContaCorrente %s e N�mero %l n�o foi encontrado na tabela de ChequesPre.
Public Const ERRO_CLINTE_FILIAL_NAO_CONFEREM = 2504 'Parametro: lCliente, iFilial
'O Cliente %l ou a Filial %i n�o conferem com o BD.
Public Const ERRO_PARCELAS_VINCULADAS_CHQPRE = 2505 'Parametro: lNumIntParc
'N�o existe Parcelas associadas ao Cheque Pre %l.
Public Const ERRO_VENDEDOR_JA_EXISTENTE = 2508 'Parametro: sNomeReduzido
'O Vendedor %s j� existe no Grid de Comiss�es. Duas Comiss�es n�o podem ter o mesmo Vendedor.
Public Const ERRO_VENDEDOR_COMISSAO_NAO_INFORMADO = 2509 'iComissao
'O Vendedor da comissao %i do T�tulo n�o foi informado.
Public Const ERRO_LEITURA_CREDITOSPAGFORN = 2587 'Par�metros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Erro na leitura  da tabela de CreditosPagForn com Fornecedor l%, Filial i%, Tipo de Documento %s, N�mero l% e Data de Emiss�o %dt.
Public Const ERRO_LEITURA_TIPODOCUMENTO = 2589 'Par�metro: sSigla
'Erro na leitura da tabela de TiposDeDocumento com o Tipo de Documento com Sigla %s.
Public Const ERRO_TIPO_NAO_PREENCHIDO = 2592 'Sem par�metros
'Preenchimento do Tipo de Documento � obrigat�rio.
Public Const ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO = 2593 'Par�metro: sSiglaTipoDocumento
'O Tipo de Documento %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_INSERCAO_CREDITOSPAGFORN = 2596 'P�rametros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo
'Erro na tentativa de inserir um novo registro na tabela CreditosPagForn com Fornecedor %l, Filial %i Tipo de Documento %s e N�mero %l.
Public Const ERRO_CLIENTE_NAO_PREENCHIDO = 2676 'Sem parametro
'O preenchimento de Cliente � obrigat�rio.
Public Const ERRO_INSERCAO_DEBITOSRECCLI = 2678 'Parametro: lNumProxDebitoRecCli
'Erro na tentativa de inserir um novo registro n�mero %l na tabela de DebitosRecCli.
Public Const ERRO_LOCK_VENDEDOR1 = 2687 'Par�metro: iVendedor
'N�o foi poss�vel fazer o Lock do Vendedor %i da tabela Vendedores.
Public Const ERRO_LEITURA_TIPOSDEDOCUMENTO = 2688
'Erro na leitura da tabela de Tipo de Documento.
Public Const ERRO_PARCELA_VINCULADA_CHQPRE = 2760
'A parcela a receber selecionada j� est� vinculada a outro cheque-pr�.
Public Const ERRO_PARCELA_VINCULADA_CHEQUEPRE = 2762 'Par�metros: iNumParcela, lNumIntCheque
'A Parcela com n�mero interno %i j� est� associada ao Cheque-Pr� com n�mero interno %l.
Public Const ERRO_LEITURA_COBRADOR1 = 2763 'Par�metro: sNomeReduzido
'Erro na leitura da tabela de Cobrador com Nome Reduzido %s.
Public Const ERRO_LEITURA_PARCELAS_REC_BAIXADAS = 2862
'Erro na leitura da tabela de Parcelas a Receber Baixadas.
Public Const ERRO_PARCELA_RECEBER_NAO_CADASTRADA1 = 2890 'Par�metro: lNumIntParc
'A Parcela a Receber com o N�mero Interno %l n�o foi encontrada no Banco de Dados
Public Const ERRO_VENDEDOR_INICIAL_MAIOR = 2916 'Sem parametros
'C�digo do Vendedor Inicial � maior
Public Const ERRO_CLIENTE_INICIAL_MAIOR = 2938
'O Cliente Inicial � maior que o final.'
Public Const ERRO_CLIENTE_NAO_CADASTRADO_2 = 2939
'O Cliente n�o est� cadastrado.'
Public Const AVISO_CREDITO_PAGAR_NUMERO_REPETIDO = 5285 'Par�metros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo e dtDataEmissao
'No Banco de Dados j� existe Cr�dito com Fornecedor com os dados abaixo: N�mero: %l, Fornecedor: %l, Filial: %i, Tipo: %s e Data de Emiss�o: %dt. Deseja prosseguir na inser��o de um novo Cr�dito com mesmo N�mero e Fornecedor?
Public Const AVISO_DEBITORECCLI_JA_EXISTENTE = 5287 'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo e dtDataEmissao
'No Banco de Dados j� existe D�bito com Cliente com os dados abaixo: N�mero: %l, Cliente: %l, Filial: %i, Tipo: %s e Data de Emissao: dt%. Deseja prosseguir na inser��o de novo D�bito do mesmo tipo com o mesmo n�mero ?
Public Const ERRO_CODIGO_INVALIDO1 = 16010 'Sem Parametros
'O C�digo tem que ser um valor inteiro positivo.



''VEIO DE ERROS FAT
Public Const ERRO_CODIGO_INVALIDO = 8009 'Parametro: Codigo.Text
'O preenchimento de C�digo deve ser um n�mero maior que 99.
Public Const ERRO_SERIE_NAO_CADASTRADA = 8010 'Parametro: sSerie
'A Serie %s n�o est� cadastrada no Banco de Dados.
Public Const ERRO_LEITURA_SERIE = 8011 'Sem Parametro
'Erro de leitura no Banco de Dados.
Public Const ERRO_LEITURA_CANALVENDA1 = 8048 'Parametros objCanal.iCodigo
'Erro na leitura do canal %i da tabela CanalVenda
Public Const ERRO_LIMITEMENSAL_NAO_INFORMADO = 8107
'O Limite Mensal deve ser informado.
Public Const ERRO_LIMITEOPERACAO_NAO_INFORMADO = 8108
'O Limite de Opera��o deve ser informado.
Public Const ERRO_ALCADA_NAO_CADASTRADA = 8110 'Parametro NomeReduzido
'A al�ada do usu�rio %s n�o est� cadastrada.
Public Const ERRO_DATADE_MAIOR_DATAATE = 8129 'Sem par�metro
'"Data De" deve ser menor que "Data At�".
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
'N�o � permitido excluir a Natureza de Opera��o %s, pois est� vinculada com Pedido de Venda %l da Filial Empresa %i.
Public Const ERRO_NATUREZAOP_USADO_PEDIDODEVENDABAIXADO = 8148 'Parametro : sCodigo da Natureza
'N�o � permitido excluir a Natureza de Opera��o %s, pois est� vinculada com Pedido de Venda Baixado %l da Filial Empresa %i.
Public Const ERRO_NATUREZAOP_USADO_NFISCAL = 8149 'Parametro : sCodigo da Natureza
'N�o � permitido excluir a Natureza de Opera��o %s, pois est� vinculada com a Nota Fiscal - S�rie = %s, N�mero = %l e Filial Empresa = %l.
Public Const ERRO_NATUREZAOP_USADO_NFISCALBAIXADA = 8150 'Parametro : sCodigo da Natureza
'Erro Natureza %s est� sendo utilizada em uma Nota Fiscal baixada.
Public Const ERRO_NATUREZAOP_USADO_PADRAOTRIBSAIDA = 8151 'Parametro : sCodigo da Natureza
'Erro Natureza %s est� sendo utilizada em um Padr�o Tributa��o Saida.
Public Const ERRO_NATUREZAOP_USADO_PADRAOTRIBENTRADA = 8152 'Parametro : sCodigo da Natureza
'Erro Natureza %s est� sendo utilizada em um Padr�o Tributa��o Entrada.
Public Const ERRO_NATUREZAOP_USADO_TIPODOCINFO = 8153 'Parametro : sCodigo da Natureza
'N�o � permitido excluir a Natureza de Opera��o %s, pois est� vinculada com Tipo de Documento (Sigla = %s).
Public Const ERRO_NATUREZAOP_USADO_TRIBUTACAONF = 8154 'Parametro : sCodigo da Natureza
'Erro Natureza %s est� sendo utilizada em uma Tributa��o N.F.
Public Const ERRO_NATUREZAOP_USADO_TRIBUTACAOITEMPV = 8155 'Parametro : sCodigo da Natureza
'Erro Natureza %s est� sendo utilizada em uma Tributa��o Item P.V.
Public Const ERRO_NATUREZAOP_USADO_TRIBUTACAOCOMPLNF = 8156 'Parametro : sCodigo da Natureza
'Erro Natureza %s est� sendo utilizada em uma Tributa��o Complemento N.F..
Public Const ERRO_NATUREZAOP_USADO_TRIBUTACAOCOMPLPV = 8157 'Parametro : sCodigo da Natureza
'Erro Natureza %s est� sendo utilizada em uma Tributa��o Complemento P.V..
Public Const ERRO_NATUREZAOP_USADO_TRIBUTACAOITEMNF = 8158 'Parametro : sCodigo da Natureza
'Erro Natureza %s est� sendo utilizada em uma Tributa��o  Item N.F.
Public Const ERRO_NOTA_FISCAL_INTERNA_ENTRADA_NAO_CADASTRADA = 8160 'Parametros sSerie, lNumNotaFiscal
'Nota Fiscal Interna de Entrada com s�rie %s e n�mero %l n�o est� cadastrada no Banco de Dados.
Public Const ERRO_ATUALIZACAO_TABELASDEPRECO = 8191 'Parametros: iCodigo
'Erro na atualiza��o de registro na tabela TabelasDePreco com c�digo da Tabela %i.
Public Const ERRO_DESCRICAO_TABELAPRECO_JA_EXISTENTE = 8192 'Par�metro: Descri��o
'A Descri��o %s j� � utilizada por outra Tabela de Pre�o.
Public Const ERRO_TABELAPRECO_JA_EXISTENTE = 8193 'Parametro: iCodigo
'A Tabela de Pre�o com o c�digo %i j� existe no Banco de Dados.
Public Const ERRO_INSERCAO_TABELASDEPRECO = 8194 'Parametros: iCodigo
'Erro na inser��o de registro na tabela TabelasDePreco com c�digo da Tabela %i.
Public Const ERRO_AUSENCIA_PEDIDO_BAIXAR = 8196 'Sem par�metros
'Deve haver pelo menos um Pedido marcado para ser baixado.
Public Const ERRO_BLOQUEIOPV_REPETIDO = 8254 'iTipoBloqueio
'J� existe no grid um bloqueio com o tipo %i.
Public Const ERRO_ITEM_ARVORE_CONSULTA_NAO_SELECIONADO = 8296 'Parametros
'� necess�rio selecionar um �tem na �rvore de Consultas.
Public Const ERRO_LEITURA_CONSULTAS = 8297 'Sem par�metros
'Erro na leitura da tabela Consultas.



''VEIO DE ERROS MAT
Public Const ERRO_LEITURA_NFISCAL4 = 7639 'Sem Par�metros
'Erro na leitura da tabela de Notas Fiscais
Public Const ERRO_ATUALIZACAO_SERIE = 7640 'Parametro: sSerie
'Erro na atualiza��o da S�rie %s na tabela de S�ries.
Public Const ERRO_VALOR_TOTAL_COMISSAO_INVALIDO = 7907 'Par�metro: dValorTotalComissao, dValorTotal
'O total de valores de comiss�es = %d n�o pode ultrapassar o Valor Total = %d.
Public Const ERRO_VENDEDOR_COMISSAO_GRID_NAO_INFORMADO = 7908 'Par�metro: iComissao
'O Vendedor da comiss�o %i do Grid de Comiss�es n�o est� preenchido.


''VEIO DE ERROS TRB
Public Const ERRO_TIPO_TRIBUTACAO_NAO_CADASTRADO = 7013 'Par�metro: iTipo
'O Tipo %i de Tributa��o n�o est� cadastrado no Banco de Dados.


'Veio de ErrosCOM
Public Const ERRO_REGISTRO_COMPRAS_CONFIG_NAO_ENCONTRADO = 12004 'Parametros sCodigo,iFilialEmpresa
'Registro na tabela ComprasConfig com C�digo=%s e FilialEmpresa=%i n�o foi encontrado.
Public Const ERRO_LEITURA_PEDIDOCOMPRATODOS = 12277 'Sem par�metros
'Erro na leitura da tabela PedidoCompraTodos.
Public Const ERRO_LEITURA_PEDIDOCOTACAOTODOS = 12361 'Sem par�metros
'Erro na leitura de PedidoCotacaoTodos.




'C�digos de Avisos - Reservado de 5400 at� 5499
Public Const AVISO_CRIAR_VENDEDOR = 5400
'Deseja cadastrar novo Vendedor?
Public Const AVISO_CRIAR_VENDEDOR1 = 5401 'Parametro: iCodigo
'Vendedor com c�digo %i n�o est� cadastrado. Deseja cadastrar?
Public Const AVISO_CRIAR_VENDEDOR2 = 5402 'Parametro: sNomeReduzido
'Vendedor com Nome Reduzido %s n�o est� cadastrado. Deseja cadastrar?
Public Const AVISO_CRIAR_CLIENTE = 5403
'Deseja cadastrar novo Cliente?
Public Const AVISO_CRIAR_CLIENTE_1 = 5404 'Parametro: sNomeReduzido
'Cliente %s n�o est� cadastrado. Deseja criar?
Public Const AVISO_CRIAR_CLIENTE_2 = 5405 'Parametro: lCodigo
'Cliente com c�digo %s n�o est� cadastrado. Deseja criar?
Public Const AVISO_CRIAR_CLIENTE_3 = 5406 'Parametro: sCGC
'Cliente com CGC/CPF %s n�o est� cadastrado. Deseja criar?
Public Const AVISO_EXCLUIR_CATEGORIACLIENTE = 5407 'Par�metro: sCategoria
'Confirma exclus�o da Categoria %s?
Public Const AVISO_DESEJA_CRIAR_CATEGORIACLIENTE = 5408 'Sem par�metro
'Confirma a cria��o de uma nova Categoria de Cliente?
Public Const AVISO_DESEJA_CRIAR_CATEGORIACLIENTEITEM = 5409 'Sem par�metros
'Confirma a cria��o de um novo �tem de Categoria de Cliente?
Public Const AVISO_CONFIRMA_EXCLUSAO_CONDICAOPAGTO = 5410 'Par�metro: iCodigo
'A Condi��o de Pagamento %i ser� exclu�da. Confirma exclus�o?
Public Const AVISO_CONFIRMA_EXCLUSAO_TIPODEVENDEDOR = 5411 'Par�metro: iCodigo
'O Tipo de Vendedor %i ser� exclu�do. Confirma exclus�o?
Public Const AVISO_CONFIRMA_EXCLUSAO_VENDEDOR = 5412 'Parametro iCodigo
'O Vendedor %i ser� exclu�do. Confirma exclus�o?
Public Const AVISO_CRIAR_TIPO_VENDEDOR = 5413
'Deseja criar um novo Tipo de Vendedor?
Public Const AVISO_COMISSAO_EMISSAO_PAGA = 5414 'Parametro: tTituloReceber.lNumIntDoc
'O N�mero de Comiss�es na Emiss�o do T�tulo com n�mero %l n�o pode ser alterado por que j� existe comissao paga.
Public Const AVISO_DATAVENCIMENTO_PARCELAS_ALTERAVEIS = 5415
'Este T�tulo j� est� lan�ado, portanto s� � permitido alterar os campos referentes a Parcelas (Datas de Vencimento), Descontos, Cheque-Pr� e Comiss�es. Deseja prosseguir?
Public Const AVISO_EXCLUSAO_TITULORECEBER = 5416 'Parametro : lNumTitulo
'Confirma a exclus�o do T�tulo a Receber n�mero %l ?
Public Const AVISO_PARCELA_COM_BAIXA_NAO_ALTERAVEL = 5417 'Parametro: iParcela
'A Parcela %i com baixa n�o pode ter o n�mero de descontos alterado. Deseja prosseguir na altera��o para os campos alter�veis?
Public Const AVISO_PARCELA_COM_BAIXA_DESCONTO_INALTERAVEL = 5418 'iParcela
'A Parcela %i com baixa n�o pode ter o descontos alterados. Deseja prosseguir na altera��o para os campos alter�veis?
Public Const AVISO_COMISSOES_EMISSAO_NAO_ALTERAVEIS = 5419 'lNumIntDoc
'O n�mero de comiss�es na emiss�o do T�tulo com N�mero Interno %l n�o pode ser alterado porque existe comiss�o paga. Deseja prosseguir na altera��o para os campos alter�veis?
Public Const AVISO_DATA_VALOR_CHEQUEPRE_NAO_ALTERAVEIS = 5420 'iParcela
'O Valor e Data de ChequePre da Parcela %i com baixa n�o podem ser alterados.Deseja prosseguir na altera��o para os campos alter�veis?
Public Const AVISO_NUM_COMISSOES_NAO_ALTERAVEL = 5421 'iParcela
'O N�mero de Comiss�es da Parcela %i n�o pode ser alterado porque existe comiss�o paga. Deseja prosseguir na altera��o para os campos alter�veis?
Public Const AVISO_COMISSAO_PARCELA_PAGA = 5422 'iParcela
's comiss�es da Parcela %i n�o poder�o ser alteradas porque existe comiss�o paga. Deseja prosseguir na altera��o para os campos alter�veis?
Public Const AVISO_TITULORECEBER_PENDENTE_MESMO_NUMERO = 5423 'sSiglaDocumento, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Titulo a Receber Pendente com os dados Tipo = %s, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de Novo Titulo a Receber com o mesmo n�mero?
Public Const AVISO_TITULORECEBER_MESMO_NUMERO = 5424 'sSiglaDocumento, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Titulo a Receber com os dados Tipo = %s, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de novo Titulo a Receber com o mesmo n�mero?
Public Const AVISO_TITULORECEBER_BAIXADO_MESMO_NUMERO = 5425 'sSiglaDocumento, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Titulo a Receber Baixado com os dados Tipo = %s, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de Novo Titulo a Receber com o mesmo n�mero?
Public Const AVISO_CARTEIRA_JA_ADICIONADA = 5426 'sCarteira
'A Carteira  %s  j� est� adicionada.
Public Const AVISO_BANCO_INEXISTENTE = 5427 'Parametro objBanco.iCodBanco
'O Banco %i n�o est� cadastrado
Public Const AVISO_EXCLUIR_COBRADOR = 5428 'Parametro objcobrador.icodigo
'Aviso excluir cobrador %i da tabela cobradores
Public Const AVISO_CONFIRMA_EXCLUSAO_CARTEIRACOBRANCA = 5429
'Confirma exclus�o de Carteira Cobran�a ?
Public Const AVISO_EXCLUIR_TRANSPORTADORA = 5430 'Parametro: iCodigo
'Confirma exclus�o da transportadora com c�digo %i?
Public Const AVISO_VALOR_DESCONTO_MAIOR1 = 5431 'dDesconto, dSomaValores
'Desconto = %d n�o pode ultrapassar a soma de Produtos + ICMSSubst + IPIValor + Frete + Seguro + Despesas = %d. Desconto ser� zerado.
Public Const AVISO_EXISTENCIA_NOTA_FISCAL_SAIDA = 5432 'Par�metros: lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'No Banco de Dados existe Nota Fiscal de Saida com os Dados C�digo do Cliente =%l, C�digo da Filial =%i, Tipo =%i,  S�rie NF =%s, N�mero NF =%l, Data Emis�o =%dt.
Public Const AVISO_EXISTENCIA_NOTA_FISCAL_SAIDA_BAIXADA = 5433 'Par�metro: lCliente, iFilialCli, iTipoNFiscal, sSerie, lNumNotaFiscal, dtDataEmissao
'No Banco de Dados existe Nota Fiscal de Saida Baixada com os Dados C�digo do Cliente =%l, C�digo da Filial =%i, Tipo =%i,  S�rie NF =%s, N�mero NF =%l, Data Emis�o =%dt.
Public Const AVISO_IR_FONTE_MAIOR_VALOR_TOTAL = 5434
'IR Fonte maior que o valor total.
Public Const AVISO_NF_INTERNA_DATA_PROXIMA = 5435 'sSerie, lNumNotaFiscal, dtDataEmissao
'No Banco de Dados existe Nota Fiscal Interna com os dados S�rie =%s, N�mero =%l, Data Emiss�o =%dt. Deseja prosseguir na inser��o de nova Nota Fiscal com o mesmo n�mero?
Public Const AVISO_EXCLUSAO_TABELA_DE_PRECO = 5436 'Parametro: iCodigo
'Confirma a exclus�o da Tabela de Pre�o com c�digo %i ? Ser� exclu�do a tabela para todas as Filiais.
Public Const AVISO_EXCLUSAO_ITEM_TABELA_DE_PRECO = 5437 'Parametros: sCodProduto, iCodTabela
'Confirma a exclus�o de Item com c�digo %s da Tabela de Pre�o com c�digo %i ?
Public Const AVISO_CLIENTE_CGC_IGUAL = 5438 'Parametro: sCGC
'J� existe um outro Cliente Cadastrado com o CGC %s, deseja continuar a Grava��o?
Public Const AVISO_INFORMA_NUMERO_NOTA_GRAVADA = 5439  'Parametros: lNumNotaFiscal
'A Nota Fiscal foi gravada com o N�mero %l.
Public Const AVISO_INFORMA_NUMERO_FATURA = 5440  'Parametros: lNumFatura
'A Fatura a Receber foi gravada com o N�mero %l.
Public Const AVISO_FATURA_LOCKADA = 5441 'Sem Parametros
'A Impress�o da Fatura a Receber est� bloqueada. Est� havendo uma impress�o ou houve erro anterior.
'Continue somente em caso de Erro anterior. Deseja Continuar?
Public Const AVISO_FATURA_REIMPRESSA = 5442 'Parametros : lNumeroFaturaInicial, lNumFaturaFinal
'Confirma a reimpress�o das Faturas a Receber de %l at� %l ?
Public Const AVISO_CRIAR_PAIS = 5443 'Parametros: iCodPais
'O Pa�s %i n�o existe, deseja cri�-lo?


'VEIO DE ERROS EST
Public Const AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB = 5603 'sSerie, lNumNotaFiscal, dtDataEmissao
'Nota Fiscal Interna com os dados S�rie = %s , N�mero = %l, Data Emiss�o = %dt est� cadastrada no Banco de Dados, s� � possivel alterar os dados relativos a contabilidade. Deseja proseguir na altera��o?



'Erros CRFAT 2
'C�digos de Erros - Reservado de 13100 a 13299
Public Const ERRO_LEITURA_CONTRATOFORNECIMENTO = 13100 'Sem parametros
'Erro na leitura da tabela de ContratoFornecimento.
Public Const ERRO_LEITURA_ITENS_CONTRATO = 13101 'Sem parametros
'Erro na leitura da tabela de ItensContrato.
Public Const ERRO_FORNECEDOR_REL_PEDIDOCOMPRA = 13102 'Par�metros: lCodigoForn, lCodigoPedidoCompra
'O Fornecedor %l est� relacionado com o Pedido de Compra %l.
Public Const ERRO_FORNECEDOR_REL_CONTRATOFORNECIMENTO = 13103 'Par�metros: lCodigoForn, lCodigoContrato
'O Fornecedor %l est� relacionado com o Contrato de Fornecimento %l.
Public Const ERRO_FORNECEDOR_REL_ITEMCONCORRENCIA = 13104 'Par�metros: lCodigoForn
'O Fornecedor %l est� relacionado com Item de Concorr�ncia.
Public Const ERRO_FORNECEDOR_REL_CONCORRENCIA = 13105 'Par�metros: lCodigoForn, lCodigoConcorrencia
'O Fornecedor %l est� relacionado com a Concorr�ncia %l.
Public Const ERRO_FORNECEDOR_REL_REQCOMPRA = 13106 'Parametro: lCodFornecedor, lCodRequisicaoCompra
'O Fornecedor %l est� relacionado com a Requisi��o de Compra %l.
Public Const ERRO_FORNECEDOR_REL_PEDIDOCOTACAO = 13107 'Parametro: lCodFornecedor, lCodPedidoCotacao
'O Fornecedor %l est� relacionado com o Pedido de Cota��o %l.
Public Const ERRO_FORNECEDOR_REL_COTACAO = 13108 'Parametro: lCodFornecedor, lCodCotacao
'O Fornecedor %l est� relacionado com a Cota��o %l.
Public Const ERRO_FORNECEDOR_REL_ITEMREQCOMPRA = 13109 'Parametro: lCodFornecedor
'O Fornecedor %l est� relacionado com um item de Requisi��o de Compra.
Public Const ERRO_FORNECEDOR_REL_REQMODELO = 13110 'Parametro: lCodFornecedor, lCodRequisicaoModelo
'O Fornecedor %l est� relacionado com a Requisi��o Modelo %l.
Public Const ERRO_FORNECEDOR_REL_COTACAOPRODUTO = 13111 'Parametro: lCodFornecedor
'O Fornecedor %l est� relacionado com Cota��o Produto.
Public Const ERRO_LEITURA_ITENSREQMODELO2 = 13112 'sem par�metros
'Erro na leitura da tabela ItensReqModelo.
Public Const ERRO_FORNECEDOR_REL_ITENSREQMODELO = 13113 'Parametro: lCodFornecedor
'O Fornecedor %l est� relacionado com um item de Requisi��o Modelo.
Public Const ERRO_LEITURA_COTACAOPRODUTOTODAS = 13114 'Sem parametros
'Erro na leitura da tabela CotacaoProdutoTodas.
Public Const ERRO_LEITURA_CONCORRENCIATODAS = 13115 'Sem Par�metros
'Erro de leitura na tabela de ConcorrenciaTodas.
Public Const ERRO_FORNECEDOR_REL_FORNECEDOR_PRODUTOFF = 13116 'Par�metros: lCodFornecedor, sProduto
'O Fornecedor %l est� relacionado com o Produto %s.
Public Const ERRO_LEITURA_IMPORTCLI = 13117 'Sem parametros
'Erro na leitura da tabela ImportCli.




