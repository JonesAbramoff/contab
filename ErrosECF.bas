Attribute VB_Name = "ErrosECF"
Option Explicit

Public GL_ConexaoOrc As Long

Public Const CORPORATOR_ECF_VERSAO_PGM = "537"
Public Const CORPORATOR_ECF_VERSAO_BD_ECF = "8350"
Public Const CORPORATOR_ECF_VERSAO_BD_ORC = "8137"

Public Const TEF_PAYGO_CERTIFICADO = "03252" 'C�digo obtido junto � NTK Solutions no in�cio do processo de certifica��o da Automa��o Comercial

Private Declare Function Sistema_ObterTipoCliente Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterTipoCliente" (ByVal lID_Sistema As Long) As Long
Private Declare Function Rotina_Erro_Carregar Lib "ADCRTL.DLL" (ByVal sTipoErro As String, ByVal lLocalErro As Long, ByVal lpMsgErro As String, ByVal sParam1 As String, ByVal sParam2 As String, ByVal sParam3 As String, ByVal sParam4 As String, ByVal sParam5 As String, ByVal sParam6 As String, ByVal sParam7 As String, ByVal sParam8 As String, ByVal sParam9 As String, ByVal sParam10 As String) As Long

Private Const AD_SIST_NORMAL = 0
Private Const AD_SIST_RELLIB = 1
Private Const AD_SIST_BATCH = 2

'Erros
Public Const ERRO_EM_PROCESSAMENTO_SEFAZ = "Aguardando processamento do envio da nfce pela sefaz"
Public Const ERRO_NFCEINFO_NAO_CADASTRADO = "A NFCe com chave de acesso %s n�o foi encontrada no banco de dados."
Public Const ERRO_NFCEINFO_CANC_NAO_CADASTRADO = "A NFCe com chave de acesso %s n�o est� registrada como cancelada."
Public Const ERRO_LEITURA_NFCEINFO = "Erro na leitura da tabela NFCeInfo."
Public Const ERRO_NFCE_OFFLINE_NAO_TRANSMITIDA = "Precisa transmitir para a retaguarda os arquivos xml das nfce offline antes de autoriz�-las. V� em Fun��es/Opera��es de Arquivo/Transmitir."
Public Const ERRO_XJUST_INVALIDO = "O motivo da contingencia precisa ter no m�nimo 15 caracteres."
Public Const AVISO_CANCELA_ITEM_CUPOM = "Deseja realmente cancelar um item desta venda ?"
Public Const ERRO_DOWNLOAD_ARQUIVO_FTP1 = "Erro na recep��o do arquivo via ftp: %s"
Public Const ERRO_UPLOAD_ARQUIVO_FTP1 = "Erro no envio do arquivo via ftp: %s"
Public Const ERRO_COMUNICACAO_FTP = "Erro no download do arquivo via ftp"
Public Const ERRO_ARQUIVO_FTP_NAO_ENCONTRADO = "arquivo n�o econtrado no servidor ftp: %s"
Public Const ERRO_EMAIL_INVALIDO = "O email digitado n�o � v�lido."
Public Const ERRO_SAT_CODIGO_ATIVACAO = "O c�digo de ativa��o possui entre 8 e 32 caracteres"
Public Const ERRO_LOCAL_ARQUIVO_NAO_CONFIGURADO = "O local do arquivo %s n�o foi configurado." 'Inclu�do por Luiz Nogueira em 16/04/04
Public Const ERRO_FORNECIDO_PELO_VB_1 = "Erro fornecido pelo VB: %s. (%s)"
Public Const ERRO_ARQUIVO_NAO_RETRANSMITIDO = "O arquivo n�o pode ser retransmitido pois o arquivo j� existe. Delete-o e tente novamente."
Public Const ERRO_FIGURA_INVALIDO = "O arquivo referente a Figura %s deste produto n�o existe no sistema."
Public Const ERRO_CAIXA_NAO_EXISTENTE = "Caixa Padr�o n�o encontrado. Verifique o arquivo de transfer�ncia."
Public Const ERRO_ARQUIVO_ABERTO = "O Arquivo est� aberto. Tente novamente"
Public Const ARQUIVO_INICIALIZACAO_NAO_EXISTENTE = "O Arquivo de inicializa��o %s n�o foi encontrado."
Public Const ERRO_ARQUIVO_DIFERENTE = "O Arquivo do caixa est� com dados pertencentes a outro caixa."
Public Const ERRO_ITEM_NAO_PREENCHIDO1 = "O Item deve ser preenchido."
Public Const ERRO_ITEM_CANCELADO = "O Item j� foi cancelado."
Public Const ERRO_ITEM_NAO_EXISTENTE1 = "O Item n�o existe neste cupom."
Public Const ERRO_QUANTIDADE_INVALIDA = "O valor da quantidade n�o pode superar 4 d�gitos inteiros."
Public Const ERRO_PARCELAMENTO_NAO_EXISTENTE = "O parcelamento selecionado n�o existe, selecione outro e repita a opera��o."
Public Const ERRO_ADMMEIOPAGTO_NAO_EXISTENTE = "O AdmMeioPagto selecionado n�o existe, selecione outro e repita a opera��o."
Public Const ERRO_VALOR_JA_PAGO = "N�o existe valor para ser pago por TEF."
Public Const ERRO_DDD_NAO_PREENCHIDO = "O DDD n�o foi preenchido"
Public Const ERRO_OPERADORA_NAO_PREENCHIDA = "A Operadora n�o foi preenchida"
Public Const ERRO_MEIOSPAG_ULTRAPASSAM = "Os meios de pagamento informados ultrapassam o valor a ser pago."
Public Const ERRO_CUPOM_NAO_CANCELADO = "N�o existe cupom anterior a ser cancelado."
Public Const ERRO_VENDEDOR_NAO_PREENCHIDO = "O vendedor n�o foi preenchido"
Public Const ERRO_NUMEROSERIE_NAO_PREENCHIDO = "O N� de S�rie deve estar preenchido."
Public Const ERRO_DATABOMPARA_MENOR = "A data para dep�sito do cheque n�o pode ser menor que a data Atual."
Public Const ERRO_SANGRIA_MAIOR = "O valor da Sangria %s n�o pode ser maior do que o saldo %s em caixa."
Public Const ERRO_REDUCAO_JA_EXECUTADA = "Foi executada a Redu��o Z para a data de %s(hoje), nenhum movimento pode ser executado, s� ser�o permitidas consultas."
Public Const ERRO_NOME_TIPOMEIOPAGTO_NAO_EXISTENTE = "O Tipo Meio Pagto com nome: %s n�o foi cadastrado."
Public Const ERRO_CODIGO_TIPOMEIOPAGTO_NAO_EXISTENTE = "O Tipo Meio Pagto com c�digo: %s n�o foi cadastrado."
Public Const ERRO_NOME_CARTAO_NAO_EXISTENTE = "O Cart�o com nome: %s n�o foi cadastrado."
Public Const ERRO_CODIGO_CARTAO_NAO_EXISTENTE = "O Cart�o com c�digo: %s n�o foi cadastrado."
Public Const ERRO_NOME_PARCELAMENTO_NAO_EXISTENTE = "O Parcelamento com nome: %s n�o foi cadastrado."
Public Const ERRO_CODIGO_PARCELAMENTO_NAO_EXISTENTE = "O Parcelamento com c�digo: %s n�o foi cadastrado."
Public Const ERRO_NOME_VALETICKET_NAO_EXISTENTE = "O Vale/Ticket com nome: %s n�o foi cadastrado."
Public Const ERRO_CODIGO_VALETICKET_NAO_EXISTENTE = "O Vale/Ticket com c�digo: %s n�o foi cadastrado."
Public Const ERRO_TAMANHO_CGC_CPF = "O Tamanho do CPF/CGC do cliente est� incorreto, ele deve ter 11 ou 14 caracteres."
Public Const ERRO_TIPOMEIOPAGTODE_NAO_PREENCHIDO = "O Tipo Meio Pagto de origem n�o foi preenchido."
Public Const ERRO_TIPOMEIOPAGTOPARA_NAO_PREENCHIDO = "O Tipo Meio Pagto de destino n�o foi preenchido."
Public Const ERRO_DATABOMPARA_NAO_PREENCHIDA = "A Data de Dep�sito (Bom Para) n�o foi preenchida."
Public Const ERRO_GRUPOCHQDE_NAO_PREENCHIDO = "Se algum dos campos do agrupamento de origem: Banco, Ag�ncia, Conta, Cliente ou N�mero estiverem preenchidos, todos os campos desse agrupamento devem ser preenchidos."
Public Const ERRO_GRUPOCHQPARA_NAO_PREENCHIDO = "Se algum dos campos do agrupamento de destino: Banco, Ag�ncia, Conta, Cliente ou N�mero estiverem preenchidos, todos os campos desse agrupamento devem ser preenchidos."
Public Const ERRO_GRUPOCRTDE_NAO_PREENCHIDO = "Se algum dos campos do agrupamento de origem: Cart�o, Parcelamento estiverem preenchidos, todos os campos desse agrupamento devem ser preenchidos."
Public Const ERRO_GRUPOCRTPARA_NAO_PREENCHIDO = "Se algum dos campos do agrupamento de destino: Cart�o, Parcelamento estiverem preenchidos, todos os campos desse agrupamento devem ser preenchidos."
Public Const ERRO_VALOR_NAO_PREENCHIDO_ORIGEM = "O Valor n�o foi preenchido."
Public Const ERRO_TRANSFCAIXA_NAO_PREENCHIDO = "O C�digo da transfer�ncia deve ser preenchido."
Public Const ERRO_NECESSARIO_FECHAR_APP = "Ocorreu um erro na grava��o dos arquivos. Clique em OK para encerrar a aplica��o."
Public Const ERRO_VALORDIN_MAIOR_SALDODIN = "O Valor da Transfer�ncia em quest�o � maior que o saldo em dinheiro. Valor: %s    Saldo: %s"
Public Const ERRO_VALORCHQ_MAIOR_SALDOCHQ = "O Valor da Transfer�ncia em quest�o � maior que o saldo dos cheques. Valor: %s    Saldo: %s"
Public Const ERRO_CHQESPECIFIC_VINCULADO_BANCO = "O cheque em quest�o n�o foi especificado, mas no entanto possui v�nculo com o banco: %s"
Public Const ERRO_VALORCRT_MAIOR_SALDOCRT = "O Valor da Transfer�ncia em quest�o � maior que o saldo do Cart�o. Valor: %s    Saldo: %s        AdmMeioPagto: %s   Parcelamento: %s"
Public Const ERRO_VALORADM_MAIOR_SALDOADM = "O Valor da Transfer�ncia em quest�o � maior que o saldo do  Meio Pagto. Valor: %s    Saldo: %s        AdmMeioPagto: %s   Parcelamento: %s"
Public Const ERRO_VALORTKT_MAIOR_SALDOTKT = "O Valor da Transfer�ncia em quest�o � maior que o saldo do Ticket. Valor: %s    Saldo: %s        AdmMeioPagto: %s   Parcelamento: %s"
Public Const ERRO_TRANSFERENCIA_NAO_EXISTENTE = "A Transfer�ncia de Caixa com c�digo: %s  n�o foi cadastrada."
Public Const ERRO_VALOR_CHQESPEC_DIF_VALORTELA = "O Valor do cheque especificado em quest�o deve ser %s, pois o cheque deve ser transferido integralmente."
Public Const ERRO_VALORSANGRIA_NAO_INFORMADO = "O valor da sangria n�o foi informado."
Public Const ERRO_VALORSANGRIA_NAO_INFORMADO_GRID = "O valor da sangria n�o foi informado na linha %s."
Public Const ERRO_CARNE_NAO_EXISTENTE = "O carne n�o existe."
Public Const ERRO_PARCELA_NAO_SELECIONADA = "Deveria ter pelo menos uma Parcela selecionada no grid."
Public Const ERRO_CGC_OVERFLOW1 = "Valor %s ultrapassa o limite de CGC."
Public Const ERRO_NUMERO_NAO_INTEIRO1 = "O n�mero %s n�o � inteiro."
Public Const ERRO_NUMERO_NAO_POSITIVO1 = "O n�mero %s n�o � positivo."
Public Const ERRO_CLIENTE_NAO_EXISTENTE = "O Cliente n�o foi encontrado."
Public Const ERRO_VALOR_SANGRIACHEQUE_INEXISTENTE = "n�o exitem cheques a serem sangrados."
Public Const ERRO_CODIGO_NAO_PREENCHIDO1 = "O campo obrigat�rio c�digo n�o esta preenchido."
Public Const ERRO_DESCONTO_MAIOR = "O desconto dado � maior que o total."
Public Const ERRO_ACRESCIMO_MAIOR = "O acr�scimo dado � maior que o total."
Public Const ERRO_NAO_EXISTE_ITEM = "Deve existir pelo menos um item para ser cancelado."
Public Const ERRO_ITEM_NAO_EXISTENTE = "Deve existir pelo menos um item de Venda."
Public Const ERRO_ABERTURA_TRANSACAO1 = "N�o conseguiu abrir a transa��o."
Public Const ERRO_COMMIT_TRANSACAO1 = "N�o confirmou a transa��o."
Public Const ERRO_ORCAMENTO_NAO_PREENCHIDO = "O N�mero do Or�amento deve estar preenchido."
Public Const ERRO_IMPRESSORA_NAO_RESPONDE = "A Impressora n�o responde."
Public Const ERRO_TEF_NAO_ATIVO = "O Gerenciador Padr�o n�o esta ativo."
Public Const ERRO_SANGRIA_SUPRIMENTO_NAO_SELECIONADO = "Erro deve ser selecionado alguma opera��o, de sangria ou de suprimento."
Public Const ERRO_CAIXA_FECHADO = "O caixa %s est� fechado."
Public Const ERRO_TRANSACAO_CAIXA_ABERTA = "J� Existe uma Transa��o Aberta para o caixa de C�digo %s."
Public Const ERRO_NUMERO_INTERVALO_REDUCAO_INICIAL_MAIOR = "O n�mero correspondente ao primeiro intervalo de redu��o n�o pode ser maior que o segundo."
Public Const ERRO_DATA_INICIAL_MAIOR1 = "Data inicial n�o pode ser maior que a data final."
Public Const ERRO_DATAS_MEMORIAFISCAL_NAO_PREENCHIDA = "A leitura de mem�ria fiscal por intervalo de datas n�o est� preenchido, os intervalos s�o campos obrigat�rios."
Public Const ERRO_REDUCAO_NAO_PREENCHIDA = "A leitura de mem�ria fiscal por intervalo de redu��o n�o est� preenchido, os intervalos s�o campos obrigat�rios."
Public Const ERRO_OPERADOR_NAO_SELECIONADO = "Nenhum Operador foi selecionado."
Public Const ERRO_SENHA_NAO_PREENCHIDA1 = "O preenchimento da Senha � obrigat�rio."
Public Const ERRO_SENHA_INVALIDA1 = "Senha Inv�lida."
Public Const ERRO_TROCO_NAO_ESPECIFICADO = "O valor do troco deve estar especificado."
Public Const ERRO_ADMMEIOPAGTO_NAO_SELECIONADO = "Um AdmMeioPagto deve estar selecionado."
Public Const ERRO_PARCELAMENTO_NAO_SELECIONADO = "Um Parcelamento deve estar selecionado."
Public Const ERRO_VALOR_NAO_PREENCHIDO2 = "O campo Valor n�o foi prenchido."
Public Const ERRO_DATA_NAO_PREENCHIDA1 = "O preenchimento da Data � obrigat�rio."
Public Const ERRO_QUANTIDADE_NAO_PREENCHIDO1 = "A quantidade n�o foi informada."
Public Const ERRO_TICKET_NAO_SELECIONADO = "Um Ticket deve estar selecionado."
Public Const ERRO_VALOR_NAO_EXISTENTE = "N�o existe valor para ser pago na tela."
Public Const ERRO_TROCO_MAIOR = "O Valor do Troco � maior que o devido."
Public Const ERRO_VENDA_ANDAMENTO = "A Sess�o n�o pode ser suspensa. Existe uma venda em andamento."
Public Const ERRO_TERMINAL_NAO_SELECIONADO = "Um terminal deve estar selecionado."
Public Const ERRO_SESSAO_ABERTA = "J� existe uma sess�o aberta no caixa %s para o operador %s."
Public Const ERRO_SESSAO_SUSPENSA = "A sess�o do caixa %s est� suspensa."
Public Const ERRO_SESSAO_ABERTA_INEXISTENTE = "N�o existe sess�o aberta para o caixa %s."
Public Const ERRO_CAIXA_ABERTO = "O caixa %s j� esta aberto."
Public Const ERRO_TROCO_DIFERENTE = "O subtotal � diferente do Troco que deve ser informado."
Public Const ERRO_CARNE_MAIOR = "O Valor do Carn� � maior que o que deve ser cobrado."
Public Const ERRO_VENDEDOR_NAO_CADASTRADO2 = "O Vendedor n�o est� cadastrado no Banco de Dados."
Public Const ERRO_PRODUTO_NAO_CADASTRADO1 = "O Produto %s n�o est� cadastrado na tabela Produtos."
Public Const ERRO_PRODUTO_NAO_PREENCHIDO1 = "O Produto deve estar preenchido."
Public Const ERRO_PRODUTO_SEM_TABELAPRECO_PADRAO1 = "N�o existe pre�o cadastrado para o produto em quest�o."
Public Const ERRO_PRODUTO_SEM_PRECO = "O Produto %s n�o tem um pre�o associado na Tabela de Pre�os do Loja."
Public Const ERRO_TECLADOPADRAO_NAO_EXISTE = "N�o existe um Teclado Padr�o."
Public Const ERRO_VALOR_INSUFICIENTE = "O Valor � insuficiente para o pagamento da compra."
Public Const ERRO_SALDO_INSUFICIENTE_SANGRIA = " O saldo %s do caixa %s n�o � sucifiente para efetuar a sangria desejada %s ."
Public Const ERRO_CHEQUE_NAO_SELECIONADO = "Pelo menos um cheque deve estar selecionado."
Public Const ERRO_BANCO_NAO_PREENCHIDO = "O campo Banco deve estar preenchido."
Public Const ERRO_AGENCIA_NAO_PREENCHIDA = "O campo Agencia deve estar preenchida."
Public Const ERRO_NUMERO_NAO_PREENCHIDO = "O campo N�mero deve estar preenchido."
Public Const ERRO_CLIENTE_NAO_PREENCHIDO1 = "O campo Cliente deve estar preenchido."
Public Const ERRO_CONTA_NAO_PREENCHIDA = "O campo Conta deve estar preenchida."
Public Const ITEM_CUPOM_NAO_SELECIONADO = "Um Item do Cupom deve estar selecionado."
Public Const ERRO_DATAEMISSAO_MAIOR = "A Data de Emiss�o � maior que a Data Final."
Public Const ERRO_AUTORIZACAO_NAO_PREENCHIDA = "O campo Autoriza��o deve estar preenchido."
Public Const ERRO_VALORSANGRIA_NAO_DISPONIVEL = "O valor determinado para a Sangria %s � maior do que o valor total em caixa %s."
Public Const ERRO_LINHA_REPETIDA = " Existe uma linha duplicada no grid."
Public Const ERRO_SANGRIA_NAOPODE_SER_EXECUTADA = "A sangria n�o pode ser executada com valor zero."
Public Const ERRO_MOVIMENTO_INEXISTENTE = "O Movimento %s n�o existe no sistema."
Public Const ERRO_CARREGA_DADOS_TIPOSMEIOPAGTO = "Erro no carregamento de dados de Tipos de Meio de Pagamento para o Caixa ECF"
Public Const ERRO_CARREGAMENTO_CAIXA_CONFIG = "Erro no carregamento dos dados de configura��o do Caixa"
Public Const ERRO_PARCELAMENTO_NAO_PREENCHIDO = "O campo Parcelamento deve estar preenchido."
Public Const ERRO_CLIENTE_NAO_CADASTRADO2 = "O cliente em quest�o n�o est� cadastrado."
Public Const ERRO_TIPOMEIOPAGTO_NAO_CARTAO = "O Tipo Meio Pagto deveria ser cart�o de cr�dito ou d�bito, mas outro valor foi encontrado. Tipo = %s."
Public Const ERRO_PREENCHIMENTO_ARQUIVO_CONFIG = "O item %s da se��o %s do arquivo de configura��o %s n�o foi preenchido corretamente"
Public Const ERRO_INTERVALO_NAO_PREENCHIDO = "Alguma data n�o est� preenchida, os intervalos s�o campos obrigat�rios."
Public Const ERRO_ECF_NAO_PREENCHIDO = "O ECF n�o foi selecionado."
Public Const ERRO_TEF_NAO_RESPONDE = "TEF n�o responde."
Public Const ERRO_INCONSISTENCIA_ARQUIVO_TEF = "Inconsistencia no campo %s do arquivo %s gerado pelo TEF. %s."
Public Const ERRO_TEF_TRANSACAO_NEGADA = "%s."
Public Const ERRO_TEF_FALHA_IMPRESSAO = "Transa��o TEF cancelada. Rede: %s  NSU: %s  Valor: %s ."

Public Const ERRO_CHEQUE_CUPOMFISCAL_NAOENCONTRADO = "N�o foi encontrado o cheque associado ao cupom fiscal %s"
Public Const ERRO_NOME_OUTROS_NAO_EXISTENTE = "O Meio de Pagamento com nome: %s n�o foi cadastrado."
Public Const ERRO_CODIGO_OUTROS_NAO_EXISTENTE = "O Meio de Pagamento com c�digo: %s n�o foi cadastrado."
Public Const ERRO_OUTROS_NAO_SELECIONADO = "Um meio de pagamento deve estar selecionado."
Public Const ERRO_VALOR_NAO_PREENCHIDO_DESTINO = "O Campo Valor de destino n�o foi preenchido."
Public Const ERRO_CUPOMFISCAL_NAO_PREENCHIDO = "O cupom fiscal deve estar preenchido."
Public Const ERRO_CHEQUEPRE_NAO_ENCONTRADO2 = "O Cheque com sequencial %s n�o foi encontrado."
Public Const ERRO_TRANSFCAIXA_NAO_ENCONTRADO = "A transfer�ncia de caixa com c�digo %s n�o foi encontrada."
Public Const ERRO_CHEQUEPRE_OUTRA_TRANSFERENCIA = "O cheque referente � transfer�ncia %s n�o corresponde ao que est� na tela."
Public Const ERRO_CUPOMFISCAL_NAO_ENCONTRADO = "O cupom fiscal %s do ECF %s n�o foi encontrado."
Public Const ERRO_CHEQUEPRE_EXCLUIDO = "O Cheque com sequencial %s est� marcado como excluido. "
Public Const ERRO_TIPOMOVTOCAIXA_INVALIDO = "O Tipo de Movimento de Caixa em quest�o � inv�lido neste contexto. Tipo = %s "
Public Const ERRO_MOVIMENTO_NAO_DINHEIRO = "O Movimento %s n�o � uma movimenta��o de dinheiro."
Public Const ERRO_MOVIMENTO_NAO_CARTAO = "O Movimento %s n�o � uma movimenta��o de Cart�o."
Public Const ERRO_MOVIMENTO_NAO_CHEQUE = "O Movimento %s n�o � uma movimenta��o de cheques."
Public Const ERRO_MOVIMENTO_NAO_TICKET = "O Movimento %s n�o � uma movimenta��o de Tickets."
Public Const ERRO_MOVIMENTO_NAO_OUTROS = "O Movimento %s n�o � uma movimenta��o de Outros."
Public Const ERRO_MOVIMENTO_JA_TRANSMITIDO = "Essa opera��o n�o pode ser efetuada pois este movimento j� foi transferido para o Caixa Central."
Public Const ERRO_CHEQUE_SANGRADO = "O Cheque com sequencial %s j� foi sangrado. "
Public Const ERRO_ARQUIVO_EXISTENTE = "O arquivo %s j� existe."
Public Const ERRO_CAIXA_FECHADO_SESSAO_NAO = "O caixa est� fechado, no entanto a sess�o esta indicando que n�o."
Public Const ERRO_VALOR_JA_PAGO1 = "N�o existe valor para ser pago."
Public Const ERRO_TROCO_MAIOR_DINHEIRO = "O Valor do Troco � maior que o valor pago em dinheiro."
Public Const ERRO_CHEQUE_MAIOR_FALTA = "O Valor do Cheque ultrapassa a quantidade que falta ser paga = %s."
Public Const ERRO_PRECO_NAO_PREENCHIDO = "O pre�o unit�rio do produto n�o est� preenchido."
Public Const ERRO_DAV_NOME_CLIENTE_NAO_PREENCHIDO = "Para o DAV, o nome do cliente tem que estar preenchido, antes de gravar."
Public Const ERRO_DAV_CPFCNPJ_CLIENTE_NAO_PREENCHIDO = "Para o DAV, o cpf/cnpj do cliente tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_NOME_CLIENTE_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o nome do cliente tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_CPFCNPJ_CLIENTE_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o cpf/cnpj do cliente tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_LOGRADOURO_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o logradouro do endere�o entrega tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_NUMERO_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o numero do endere�o de entrega tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_COMPL_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o complemento do endere�o de entrega tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_BAIRRO_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o bairro do endere�o de entrega tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_CIDADE_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, a cidade do endere�o de entrega tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_UF_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, a UF do endere�o de entrega tem que estar preenchido, antes de gravar."


Public Const AVISO_NFCE_XMLS_PENDENTES_ENVIADOS = "Os arquivos xml de nfce emitidos em contingencia offline foram enviados para a sefaz."
Public Const AVISO_SEQUENCIAL_NAO_ENCONTRADO_ARQUIVAO = "O sequencial %s n�o foi encontrado no arquivo %s e portanto n�o poder� ser retransmitido."
Public Const AVISO_EXCLUSAO_TRANSFCAIXA = "Confirma a exclus�o da transfer�ncia com c�digo %s.?"
Public Const AVISO_ALTERACAO_MOVIMENTOCAIXA = "O Movimento de Caixa %s j� existe no Sistema, Deseja Realmente efetuar a altera��o ?"
Public Const AVISO_EXCLUSAO_MOVIMENTOCAIXA = "Deseja Realmente Excluir o movimento de caixa %s ? "
Public Const AVISO_ALTERACAO_MOVIMENTOCAIXA1 = "Deseja Salvar as Altera��es ?"
Public Const AVISO_ORCAMENTO_INEXISTENTE = "O Or�amento n�o existe."
Public Const AVISO_DESEJA_ABRIR_SESSAO = " N�o Existe sess�o aberta para o caixa %s. Deseja efetuar a abertura de sess�o ?"
Public Const AVISO_ORCAMENTO_VENDA = "Deseja converter o Or�amento em Venda?"
Public Const AVISO_CANCELA_CUPOM = "Deseja cancelar o Cupom em andamento?"
Public Const AVISO_CANCELA_VENDA = "Deseja cancelar a venda em andamento?"
Public Const AVISO_CANCELA_ORCAMENTO = "O or�amento vai ter as informa��es n�o gravadas perdidas. Deseja continuar?"
Public Const AVISO_CANCELAR_COMPRA = "Deseja cancelar a compra?"
Public Const AVISO_CANCELAR_TICKET = "Deseja realmente cancelar os tickets especificados?"
Public Const AVISO_APROVEITA_PAGAMENTO = "Deseja que os pagamentos sejem aproveitados?"
Public Const AVISO_DESEJA_ABRIR_CAIXA = "O caixa %s ainda n�o foi aberto. Deseja efetuar a abertura?"
Public Const AVISO_DESEJA_FAZER_SANGRIA_COMPLETA = "Deseja efetuar uma sangria completa do caixa %s ?"
Public Const AVISO_SESSAO_ABERTA = "J� existe uma sess�o aberta no caixa %s para o Operador %s."
Public Const AVISO_REDUCAOZ_ENCERRA_DIA = "A execu��o da Redu��oZ n�o permitir� o registro de novos movimentos nesta data (%s) para o caixa %s. Deseja Proceguir?"
Public Const AVISO_ALTERACAO_TRANSFCAIXA = "Tem certeza de que deseja alterar a Transfer�ncia de c�digo: %s."
Public Const AVISO_SEM_REGISTRO1 = "N�o existe nenhum registro para a sele��o executada."
Public Const AVISO_SALVAR_ALTERACOES_CARNE = "Deseja salvar os Cancelamentos dos Carn�s selecionados?"
Public Const AVISO_SEM_REGISTRO = "N�o existe nenhum registro para a sele��o executada."
Public Const AVISO_NOVAS_CONFIGURACOES = "Deseja atualizar as parcelas com o novo filtro de Cliente?"
Public Const AVISO_LIMPAR_TELA = "Deseja realmente limpar e perder todas as informa��es presentes na tela? "
Public Const AVISO_IMPRESSORA_NAO_RESPONDE = "A Impressora n�o responde. Deseja tentar novamente?"
Public Const AVISO_CANCELA_CUPOM_TELA = "Deseja realmente cancelar o cupom da tela?"
Public Const AVISO_CANCELA_CUPOM_ANTERIOR = "Deseja realmente cancelar o cupom ANTERIOR?"
Public Const AVISO_DESEJA_REDUCAOZ = "A Redu��o Z implica no fechamento do Caixa e mais nenhuma venda poder� ser efetuada no dia %s. Deseja realmente prosseguir?"
Public Const AVISO_FECHAR_SESSAO_ABERTA = "Existe uma Sess�o aberta. Deseja fech�-la?"
Public Const AVISO_CONTINUAR_IMPRESSAO = "Deseja reimprimir o TEF interrompido?"
Public Const AVISO_ORCAMENTO_IMPRESSAO = "Deseja imprimir o Or�amento?"
Public Const AVISO_INICIALIZAR_SISTEMA_AGORA = "O sistema vai ser reinicializado automaticamente a partir de agora."
Public Const AVISO_INICIALIZAR_SISTEMA = "Devido a mudan�a de dia ser� necess�rio reinicializar o sistema, caso contr�rio isto ser� feito ap�s � 2:00."
Public Const AVISO_NAO_INICIALIZADO_TERMINAL = "O Terminal n�o foi reconfigurado."
Public Const AVISO_RECONFIGURAR_VISANET = "Deseja realmente inicializar o TEF do Visanet?(Isto s� deve ser feito em caso de reconfigura��o)"
Public Const AVISO_ARQUIVO_NAO_ENCONTRADO = "O arquivo %s n�o foi encontrado. Deseja prosseguir?"
Public Const AVISO_ARQUIVO_TRANSMITIDO = "O arquivo %s foi criado com sucesso."
Public Const AVISO_NUM_CHEQUES_MAIOR_NUM_MAX_GRID = "O n�mero de cheques no caixa ultrapassa o limite de %s do grid. Fa�a a sangria de alguns para poder ver os demais. Total de Cheque = %s."
Public Const AVISO_ARQUIVOS_RETRANSMITIDOS = "O(s) arquivo(s):%s foi(ram) regerado(s) com sucesso."
Public Const AVISO_TABELAS_ATUALIZADAS = "Tabelas atualizadas com sucesso."
Public Const AVISO_CANCELAR_ACOMPANHAMENTO = "A gera��o do arquivo %s ir� continuar. Deseja cancelar o acompanhamento da gera��o?"
Public Const AVISO_ARQUIVO_GERADO = "O arquivo %s foi gerado."
Public Const AVISO_NAO_INDICETECNICO = "Este PAF-ECF n�o executa fun��es de baixa de estoque com base em �ndices t�cnicos de produ��o, n�o podendo ser utilizando por estabelecimento que necessite deste recurso."
Public Const AVISO_DESEJA_PROSSEGUIR_ACOMPANHAMENTO_ARQUIVO = "Deseja continuar o acompanhamento da gera��o do arquivo %s ?"
Public Const AVISO_CANCELAR_ACOMPANHAMENTO_ARQUIVO = "Confirma o cancelamento do acompanhamento da gera��o do arquivo %s ?"
Public Const AVISO_DESEJA_CPF_NA_NOTA = "Deseja colocar CPF na nota?"
Public Const AVISO_CAIXA_ABERTA = "Caixa Aberta."
Public Const AVISO_CANCELA_CUPOM_TELA_SUCESSO = "A venda foi cancelada com sucesso."

Public Const ERRO_NFCE_SEM_OFFLINE_PENDENTE = "N�o h� xml de nfce pendente de envio para a sefaz."
Public Const ERRO_NFCE_INFO_LEITURA = "Erro na leitura de informa��es de nfce"
Public Const ERRO_NFCE_INFO_UPDATE = "Erro na atualiza��o de informa��es de nfce com a chave %s"
Public Const ERRO_NFCE_OFFLINE_PENDENTE = "N�o foi encontrado xml pendente de nfce com a chave %s"
Public Const ERRO_CHEQUE_SANGRIA_NAOENCONTRADO = "N�o foi encontrado nenhum cheque associado a sangria em quest�o"
Public Const ERRO_ARQUIVO_NAO_ENCONTRADO1 = "N�o conseguiu gravar dados na se��o %s item %s do arquivo %s."
Public Const ERRO_PARCELAMENTO_NAO_EXISTENTE1 = "O parcelamento %s n�o existe."
Public Const ERRO_PARCELAMENTO_NAO_PREENCHIDO_GRID = "O Parcelamento n�o foi preenchido na linha %s."
Public Const ERRO_ADMINISTRADORA_NAO_PREENCHIDO_GRID = "A Administradora n�o foi preenchida na linha %s."
Public Const ERRO_SANGRIA_NAOESP_NAO_DISPONIVEL = "O valor para a Sangria de boletos n�o especificados %s � maior do que o valor total em caixa %s."
Public Const ERRO_CHEQUE_SANGRADO_GRID = "O Cheque da linha %s do grid j� foi sangrado pelo movimento de caixa %s."
Public Const ERRO_MOVIMENTO_NAO_SANGRIA_CHEQUE = "O Movimento %s n�o � uma sangria de cheques."
Public Const ERRO_MOVIMENTO_NAO_SANGRIA_BOLETO_CC = "O Movimento %s n�o � uma sangria de boletos de cart�o de cr�dito."
Public Const ERRO_MOVIMENTO_NAO_SANGRIA_OUTROS = "O Movimento %s n�o � uma sangria de outros meios de pagamento."
Public Const ERRO_MOVIMENTO_NAO_SANGRIA_TICKET = "O Movimento %s n�o � uma sangria de ticket."
Public Const ERRO_ADMMEIOPAGTO_NAO_EXISTENTE1 = "O meio de pagamento %s n�o existe."
Public Const ERRO_VALORSANGRIA_NAO_DISPONIVEL_GRID = "O valor determinado para a Sangria %s � maior do que o valor total em caixa %s na linha %s."
Public Const AVISO_EXCLUSAO_ORCAMENTO_ECF = "O or�amento vai ser desprezado. Deseja continuar?"
Public Const ERRO_NAO_HA_DADOS_TRANSMITIR = "N�o h� dados que precisem ser transmitidos"
Public Const ERRO_ARQUIVO_ORCAMENTO_INVALIDO = "Este arquivo de or�amento n�o cont�m os dados que deveria no primeiro registro. Arquivo = %s."
Public Const ERRO_ORCAMENTO_NAO_CADASTRADO1 = "O Or�amento %s n�o foi encontrado."
Public Const ERRO_CAIXA_SO_ORCAMENTO = "Esta fun��o n�o est� dispon�vel pois este terminal est� configurado para s� trabalhar com or�amentos."
Public Const ERRO_ORCAMENTO_SEM_ITEM = "O or�amento deve ter pelo menos um item."
Public Const ERRO_ARQUIVO_LOCADO = "O arquivo %s est� sendo acessado por um outro programa e n�o pode ser acessado no momento."
Public Const ERRO_GETWINDOWSDIRECTORY = "Ocorreu um erro na execu��o da fun��o GetWindowsDirectory."
Public Const ERRO_CAIXACONFIGINI_NAO_ENCONTRADO = "O arquivo caixaconfig.ini n�o foi encontrado. Arquivo = %s."
Public Const ERRO_LIMITE_DESCONTO_ULTRAPASSADO = "O Limite de Desconto do Produto %s que � %s foi ultrapassado. Percentual Desconto Total = %s."
Public Const ERRO_CLIENTE_NAO_CADASTRADO3 = "O cliente em quest�o n�o est� cadastrado. Cliente = %s."
Public Const ERRO_TAMANHO_CGC_CPF1 = "O tamanho do campo tem que ser 11 caracteres para CPF ou 14 para CGC."
Public Const ERRO_CPFCNPJ_NAO_CADASTRADO = "O cliente em quest�o n�o est� cadastrado. CPF/CNPJ = %s."
Public Const ERRO_TIPOCF_NAO_ESCOLHIDO = "� necess�rio selecionar uma das op��es: Cupom Fiscal ou Or�amento."
Public Const ERRO_VALORMINIMO_CONDPAGTO = "O Valor m�nimo da condi��o de pagamento escolhida n�o foi atingido. Valor M�nimo = %s, Valor = %s."
Public Const ERRO_ARQUIVO_NAO_ENCONTRADO2 = "N�o conseguiu encontrar o arquivo %s."
Public Const ERRO_CONDPAGTO_NAO_CADASTRADO = "A condi��o de pagamento n�o est� cadastrada. Rede = %s, Administradora = %s, Num.Parcelas = %s, Meio Pagamento = %s, Tipo Meio Pagto = %s ."
Public Const ERRO_VALIDAR_CHEQUE_SO_FUNCIONA_VENDA = "A valida��o de cheques s� funciona para vendas."
Public Const ERRO_VALOR_TEF = "O valor do TEF n�o pode superar o valor a pagar."
Public Const ERRO_BOTAO_AG_NAO_TEF = "Ao acionar o bot�o Abrir Gaveta, n�o pode ter valor no campo TEF."
Public Const ERRO_BOTAO_TEF_SEM_FALTA = "Ao acionar o bot�o TEF e ter preenchido o campo TEF, n�o pode ter valor no campo Falta."
Public Const ERRO_VALOR_TEF_SUPERIOR_FALTA = "O valor n�o pode superar o valor que falta pagar."
Public Const ERRO_SITEF_NAO_ATIVO = "Erro de Comunica��o - Cliente SiTef. Favor ativ�-lo."
Public Const ERRO_SITEF_TRN_NAO_EFETUADA = "Transa��o n�o efetuada. Favor reter o Cupom."
Public Const ERRO_ARQUIVO_CONFIGURACAO_ALTERADO = "O arquivo de configura��o (orcamento.ini) foi alterado. Entre em contato com o fabricante."
Public Const ERRO_USAIMPRESSORAFISCAL_INVALIDO = "Orcamento.ini esta setado para o uso de impressora fiscal mas o modelo de impressora escolhido n�o � fiscal."
Public Const ERRO_USAIMPRESSORANAOFISCAL_INVALIDO = "Orcamento.ini esta setado para o uso de impressora n�o fiscal mas o modelo de impressora escolhido � fiscal."
Public Const AVISO_NAO_E_POSSIVEL_EXLUIR_ORCAMENTO = "N�o � poss�vel excluir/cancelar or�amentos."
Public Const ERRO_ORCAMENTO_COM_CUPOM = "O or�camento em quest�o j� existe e j� teve seu cupom fiscal correspondente emitido. Orcamento = %s."
Public Const ERRO_TAMANHO_DO_PAPEL_MENOR_QUE_MINIMO = "O tamanho do papel da impressora � menor de 210 mm x 148 mm ou 240 mm x 140 mm."
Public Const ERRO_ARQUIVO_DAV_INVALIDO = "Este arquivo de DAV n�o cont�m os dados que deveria no primeiro registro. Arquivo = %s."
Public Const ERRO_DATAS_NAO_PREENCHIDAS = "As datas precisam estar preenchidas."
Public Const ERRO_GERACAO_ASSINATURA = "Ocorreu um erro na gera��o da assinatura RSA."
Public Const ERRO_CODIGO_CODIGOCFE_INVALIDO = "O meio de pagamento est� inv�lido."
Public Const ERRO_FUNCAO_NAO_DISPONIVEL_NAOFISCAL = "A fun��o em quest�o n�o est� dispon�vel por se tratar de impressora n�o fiscal."
Public Const ERRO_COO_NAO_PREENCHIDO = "A faixa de COO n�o est� preenchido, os intervalos s�o campos obrigat�rios."
Public Const ERRO_COO_INICIAL_MAIOR_FINAL = "O COO Inicial � maior que o final."
Public Const ERRO_ARQUIVO_ECFCORPORATOR_ALTERADO = "O arquivo ecfcorporator.crp foi alterado. Entre em contato com o fabricante."
Public Const ERRO_DATAS_DIFEREM = "A data do ecf = %s est� diferente da data do sistema = %s."
Public Const ERRO_HORAS_DIFEREM = "A diferen�a da data do sistema para a data do ecf n�o pode ultrapassar 60 minutos. hora do ecf = %s,  hora do sistema = %s."
Public Const ERRO_ABERTURA_COMANDO = "Ocorreu um erro ao tentar abrir os comandos de banco de dados."
Public Const ERRO_LEITURA_ECFCONFIG = "Ocorreu um erro ao tentar ler um registro da tabela Ecfconfig.  Codigo = %s."
Public Const ERRO_ALTERACAO_ECFCONFIG = "Ocorreu um erro ao tentar alterar um registro da tabela Ecfconfig.  Codigo = %s."
Public Const ERRO_LEITURA_MOVIMENTOCAIXA = "Ocorreu um erro ao tentar ler um registro da tabela MovimentoCaixa."
Public Const ERRO_INCLUSAO_MOVIMENTOCAIXA = "Ocorreu um erro ao tentar incluir um registro na tabela MovimentoCaixa.  NumIntDoc = %s."
Public Const ERRO_COMMIT = "Ocorreu um erro ao confirmar uma transacao (commit). "
Public Const ERRO_LEITURA_ARQUIVOSEQ = "Ocorreu um erro ao tentar ler um registro da tabela ArquivoSeq."
Public Const ERRO_INCLUSAO_ARQUIVOSEQ = "Ocorreu um erro ao tentar incluir um registro na tabela ArquivoSeq."
Public Const ERRO_ARQUIVOSEQ_NAO_CADASTRADO = "o Sequencial %s da tabela ArquivoSeq n�o esta cadastrado."
Public Const ERRO_INCLUSAO_ORCAMENTO = "Ocorreu um erro ao tentar incluir um registro na tabela Orcamento."
Public Const ERRO_EXCLUSAO_ORCAMENTO = "Ocorreu um erro ao tentar excluir um registro da tabela Orcamento."
Public Const ERRO_LEITURA_ORCAMENTO = "Ocorreu um erro ao tentar ler um registro da tabela Orcamento."
Public Const ERRO_INCLUSAO_ORCAMENTOBAIXADO = "Ocorreu um erro ao tentar incluir um registro na tabela OrcamentoBaixado."
Public Const ERRO_LEITURA_CAIXAATIVO = "Ocorreu um erro ao tentar ler um registro da tabela CaixaAtivo."
Public Const ERRO_INCLUSAO_CAIXAATIVO = "Ocorreu um erro ao tentar incluir um registro na tabela CaixaAtivo."
Public Const ERRO_EXCLUSAO_CAIXAATIVO = "Ocorreu um erro ao tentar excluir um registro da tabela CaixaAtivo."
Public Const ERRO_ORCAMENTO_JA_CADASTRADO = "O or�amento %s j� est� cadastrado. N�o � poss�vel alter�-lo nem exclui-lo."
Public Const ERRO_IMPRESSAO_NAO_PERMITIDA = "N�o � poss�vel imprimir um or�amento que n�o seja do tipo DAV."
Public Const ERRO_ECF_LEITURA_TABELA_CONFIG = "Ocorreu um erro ao tentar ler um registro da tabela %s.  Codigo = %s."
Public Const ERRO_ECF_LOCK_TABELA_CONFIG = "Ocorreu um erro ao tentar fazer um lock em um registro da tabela %s.  Codigo = %s."
Public Const ERRO_ECF_LEITURA_TABELA = "Ocorreu um erro ao tentar ler um registro da tabela %s."
Public Const ERRO_ALTERACAO_TABELA_CONFIG = "Ocorreu um erro ao tentar alterar um registro da tabela %s.  Codigo = %s."
Public Const ERRO_QTDE_INVALIDA_NUMINT = "Ocorreu um erro. A quantidade � negativa."
'Public Const ERRO_ARQUIVOSEQ_NAO_CADASTRADO = "O registro da tabela ArquivoSeq n�o est� cadastrado. Sequencial = %s."
Public Const ERRO_LEITURA_ORCAMENTOCONFIG = "Ocorreu um erro ao tentar ler um registro da tabela OrcamentoConfig.  Codigo = %s."
Public Const ERRO_ORCAMENTOCONFIG_NAO_CADASTRADO = "O registro da tabela OrcamentoConfig n�o est� cadastrado. Codigo = %s."
Public Const ERRO_LEITURA_REDUCAOZ = "Ocorreu um erro ao tentar ler um registro da tabela ReducaoZ."
Public Const ERRO_INCLUSAO_REDUCAOZ = "Ocorreu um erro ao tentar incluir um registro na tabela ReducaoZ."
Public Const ERRO_LEITURA_R01 = "Ocorreu um erro ao tentar ler um registro da tabela R01."
Public Const ERRO_LEITURA_R02 = "Ocorreu um erro ao tentar ler um registro da tabela R02."
Public Const ERRO_LEITURA_R03 = "Ocorreu um erro ao tentar ler um registro da tabela R03."
Public Const ERRO_LEITURA_R04 = "Ocorreu um erro ao tentar ler um registro da tabela R04."
Public Const ERRO_LEITURA_R05 = "Ocorreu um erro ao tentar ler um registro da tabela R05."
Public Const ERRO_LEITURA_R06 = "Ocorreu um erro ao tentar ler um registro da tabela R06."
Public Const ERRO_LEITURA_R07 = "Ocorreu um erro ao tentar ler um registro da tabela R07."
Public Const ERRO_ALTERACAO_R04 = "Ocorreu um erro ao tentar alterar um registro na tabela R04."
Public Const ERRO_ALTERACAO_R07 = "Ocorreu um erro ao tentar alterar um registro na tabela R07."

Public Const ERRO_NFCE_NAO_AUTORIZADA = "Erro na tentativa de gravar nfce n�o autorizada."
Public Const ERRO_PRODUTO_NAO_SELECIONADO = "Nenhum produto foi selecionado."
Public Const ERRO_PRODUTO_JA_SELECIONADO_LOJA = "Este produto j� foi selecionado."
Public Const ERRO_ORCAMENTO_EAD = "O checksum da tabela or�amento n�o est� correto. Codigo = %s."
Public Const ERRO_INCLUSAO_E2 = "Ocorreu um erro ao tentar gravar um registro na tabela E2."
Public Const ERRO_INCLUSAO_E1 = "Ocorreu um erro ao tentar gravar um registro na tabela E1."
Public Const ERRO_INCLUSAO_E3 = "Ocorreu um erro ao tentar gravar um registro na tabela E3."
Public Const ERRO_INCLUSAO_R01 = "Ocorreu um erro ao tentar gravar um registro na tabela R01."
Public Const ERRO_INCLUSAO_R02 = "Ocorreu um erro ao tentar gravar um registro na tabela R02."
Public Const ERRO_INCLUSAO_R03 = "Ocorreu um erro ao tentar gravar um registro na tabela R03."
Public Const ERRO_INCLUSAO_R04 = "Ocorreu um erro ao tentar gravar um registro na tabela R04."
Public Const ERRO_INCLUSAO_R05 = "Ocorreu um erro ao tentar gravar um registro na tabela R05."
Public Const ERRO_INCLUSAO_R06 = "Ocorreu um erro ao tentar gravar um registro na tabela R06."
Public Const ERRO_INCLUSAO_R07 = "Ocorreu um erro ao tentar gravar um registro na tabela R07."
Public Const ERRO_LEITURA_E1 = "Ocorreu um erro ao tentar ler um registro da tabela E1."
Public Const ERRO_LEITURA_E2 = "Ocorreu um erro ao tentar ler um registro da tabela E2."
Public Const ERRO_LEITURA_E3 = "Ocorreu um erro ao tentar ler um registro da tabela E3."
Public Const ERRO_E1_NAO_CADASTRADO = "O estoque n�o esta cadastrado na tabela E1."
Public Const ERRO_E2_NAO_CADASTRADO = "O estoque n�o est� cadastrado na tabela E2. Data = %s, Produto = %s"
Public Const ERRO_E3_NAO_CADASTRADO = "O estoque n�o est� cadastrado na tabela E3."
Public Const ERRO_INCLUSAO_D1 = "Ocorreu um erro ao tentar gravar um registro na tabela D1."
Public Const ERRO_INCLUSAO_D2 = "Ocorreu um erro ao tentar gravar um registro na tabela D2."
Public Const ERRO_INCLUSAO_D3 = "Ocorreu um erro ao tentar gravar um registro na tabela D3."
Public Const ERRO_INCLUSAO_D4 = "Ocorreu um erro ao tentar gravar um registro na tabela D4."
Public Const ERRO_EXCLUSAO_D1 = "Ocorreu um erro ao tentar excluir um registro da tabela D1."
Public Const ERRO_EXCLUSAO_D2 = "Ocorreu um erro ao tentar excluir um registro da tabela D2."
Public Const ERRO_EXCLUSAO_D3 = "Ocorreu um erro ao tentar excluir um registro na tabela D3."
Public Const ERRO_ALTERACAO_D1 = "Ocorreu um erro ao tentar alterar um registro na tabela D1."
Public Const ERRO_LEITURA_D1 = "Ocorreu um erro ao tentar ler um registro da tabela D1."
Public Const ERRO_LEITURA_D2 = "Ocorreu um erro ao tentar ler um registro da tabela D2."
Public Const ERRO_LEITURA_D3 = "Ocorreu um erro ao tentar ler um registro da tabela D3."
Public Const ERRO_LEITURA_D4 = "Ocorreu um erro ao tentar ler um registro da tabela D4."
Public Const ERRO_LEITURA_P1 = "Ocorreu um erro ao tentar ler um registro da tabela P1."
Public Const ERRO_LEITURA_P2 = "Ocorreu um erro ao tentar ler um registro da tabela P2."
Public Const ERRO_INCLUSAO_P1 = "Ocorreu um erro ao tentar gravar um registro na tabela P1."
Public Const ERRO_INCLUSAO_P2 = "Ocorreu um erro ao tentar gravar um registro na tabela P2."
Public Const ERRO_EXCLUSAO_P1 = "Ocorreu um erro ao tentar excluir um registro da tabela P1."
Public Const ERRO_EXCLUSAO_P2 = "Ocorreu um erro ao tentar excluir um registro da tabela P2."
Public Const ERRO_P1_NAO_CADASTRADO = "N�o h� registro cadastrado na tabela P1."
Public Const ERRO_LEITURA_MEIOPAGAMENTO = "Ocorreu um erro ao tentar ler um registro da tabela MeioPagamento."
Public Const ERRO_INCLUSAO_MEIOPAGAMENTO = "Ocorreu um erro ao tentar inserir um registro na tabela MeioPagamento."
Public Const ERRO_ALTERACAO_MEIOPAGAMENTO = "Ocorreu um erro ao tentar alterar um registro da tabela MeioPagamento."
Public Const ERRO_NAO_EXISTE_R01_PERIODO = "N�o existe registro R01 no periodo em quest�o."
Public Const ERRO_LEITURA_ARQUIVOECF = "Ocorreu um erro ao tentar ler um registro da tabela ArquivoECF."
Public Const ERRO_R02_NAO_CADASTRADO = "N�o h� registro cadastrado na tabela R02.  GT Arquivo = %s, GT ECF = %s."
Public Const ERRO_GT_NAO_CONFERE_BLOQUEIO = "O GT n�o confere com o o que esta no arquivo. GT Arquivo = %s, GT ECF = %s."
Public Const ERRO_NUMFAB_NAO_CADASTRADO_BLOQUEIO = "O numero de serie %s do ecf nao esta cadastrado no arquivo."
Public Const ERRO_LEITURA_CONFIGURACAOECF = "Ocorreu um erro ao tentar ler um registro da tabela ConfiguracaoECF."
Public Const ERRO_CONFIGURACAOECF_NAO_CADASTRADO = "Nao ha registro na tabela ConfiguracaoECF."
Public Const ERRO_ORCAMENTO_BAIXADO = "O Or�amento %s j� est� baixado."
Public Const ERRO_DAV_NAO_ALTERADO_DEPOIS_DE_IMPRESSO = "O DAV n�o pode ser alterado depois de impresso."
Public Const ERRO_DAV_NAO_PODE_SER_REIMPRESSO = "O DAV n�o pode ser reimpresso."
Public Const ERRO_ITEM_NAO_ENCONTRADO_CANCELAR = "N�o foi encontrado nenhum item para cancelar."
Public Const ERRO_PREVENDA_DAV_SIMULTANEOS = "As op��es de DAV e Pr� Venda estao acionadas simultaneamente."
Public Const ERRO_PREVENDA_SEM_IMPRESSORAFISCAL = "A op��o de Pr� Venda est� ligada mas o uso de impressora fiscal est� desligado."
Public Const ERRO_R07_NAO_CADASTRADO = "N�o h� registro cadastrado na tabela R07 para os dados em quest�o."
Public Const ERRO_ARQ_CRIPTOGRAFADO_INEXISTENTE = "O arquivo criptografado %s n�o foi encontrado."
Public Const ERRO_PREVENDA_CUPOM_SIMULTANEOS = "As op��es de Cupom Fiscal e Pr� Venda estao acionadas simultaneamente."
Public Const AVISO_CAIXA_SO_ORCAMENTO = "Este � um caixa que s� faz or�amento. N�o � poss�vel executar esta fun��o."
Public Const ERRO_ALTERACAO_CONFIGURACAOECF = "Ocorreu um erro ao tentar alterar um registro na tabela ConfiguracaoECF."
Public Const AVISO_EXCLUSAO_NFD2 = "Confirma a exclus�o da nota fiscal %s.?"
Public Const AVISO_ALTERACAO_NFD2 = "A nota fiscal %s j� existe no Sistema. Deseja Realmente efetuar a altera��o ?"
Public Const ERRO_EXCLUSAO_NFD2 = "Ocorreu um erro na exclus�o na nota fiscal."
Public Const ERRO_GRAVACAO_NFD2 = "Ocorreu um erro na grava��o na nota fiscal."
Public Const ERRO_LEITURA_NFD2 = "Ocorreu um erro na leitura na nota fiscal."
Public Const ERRO_NFD2_NAO_LOCALIZADA = "A nota fiscal s�rie %s n�mero %s emitida em %s n�o foi localizada."
Public Const ERRO_SERIE_NAO_PREENCHIDA = "O campo S�rie deve estar preenchido."
Public Const ERRO_NENHUM_ITEM_GRID = "Nenhum item foi inclu�do"
Public Const ERRO_DISCRIMINACAO_NAO_PREENCHIDA = "A discrimina��o da mercadoria tem que ser preenchida"
Public Const ERRO_NFD2_JA_EXISTENTE = "Esta Nota Fiscal j� est� cadastrada no Sistema."
Public Const ERRO_ORCAMENTO_NAO_PERMITE_INCLUSAO_ITENS = "Or�amento ao ser transformado para cupom n�o admite adi��o de itens."
Public Const ERRO_ITEM_ORCAMENTO_CUPOM_NAO_PODE_CANCELAR = "Or�amento ao ser transformado para cupom n�o admite cancelamento de itens."
Public Const ERRO_NAO_PERMITIDO_IMPRIMIR_DAV_NAO_FISCAL = "N�o � permitido imprimir DAV em impressora n�o fiscal."
Public Const ERRO_MD5_ALTERADO = "O MD5 do arquivo criptografado foi alterado."
Public Const ERRO_HOUVE_PERDA_DADOS_ARQ_CRIPTO = "Houve perda de dados no arquivo criptografado e o mesmmo n�o pode ser recomposto."
Public Const ERRO_NFD2_DESABILITADA = "S� pode registrar nota manual ap�s redu��o Z ou com o ECF com problema."
Public Const ERRO_LEITURA_CODNACID = "Ocorreu um erro ao tentar ler um registro da tabela CodNacId."
Public Const ERRO_CODNACID_NAO_CADASTRADO = "O CodNacId da Marca = %s , Modelo = %s, VersaoSB = %s n�o est� cadastrado."
Public Const ERRO_DIRETORIO_INVALIDO = "O diret�rio escolhido n�o existe. %s"
Public Const ERRO_VERSAO_PGM_INCOMPATIVEL_PGM = "Existe uma incompatibilidade entre a vers�o do programa instalada e o banco de dados. Pgm: %s Pgm BD: %s"
Public Const ERRO_VERSAO_PGM_INCOMPATIVEL_BD_ECF = "Existe uma incompatibilidade entre a vers�o do programa instalada e o banco de dados. BD ECF Pgm: %s BD ECF: %s"
Public Const ERRO_VERSAO_PGM_INCOMPATIVEL_BD_ORC = "Existe uma incompatibilidade entre a vers�o do programa instalada e o banco de dados. BD Orc Pgm: %s BD Orc: %s"
Public Const ERRO_NAO_PERMITIDO_CANCELAR_VARIOS_VINC = "N�o � permitido cancelar cupom com mais de um cupom vinculado."
Public Const ERRO_LEITURA_CONFIGURACAOSAT = "Ocorreu um erro ao tentar ler um registro da tabela ConfiguracaoSAT."
Public Const ERRO_ALTERACAO_CONFIGURACAOSAT = "Ocorreu um erro ao tentar alterar um registro na tabela ConfiguracaoSAT."
Public Const ERRO_LEITURA_CONFIGURACAONFE = "Ocorreu um erro ao tentar ler um registro da tabela ConfiguracaoNFe."
Public Const ERRO_ALTERACAO_CONFIGURACAONFE = "Ocorreu um erro ao tentar alterar um registro na tabela ConfiguracaoNFe."
Public Const ERRO_INCLUSAO_NFCEINFO = "Ocorreu um erro ao tentar inserir um registro na tabela NFCeInfo."
Public Const ERRO_CONFIGURACAOSAT_NAO_CADASTRADO = "N�o h� ConfiguracaoSAT cadastrado."
Public Const ERRO_LEITURA_PRODUTOCODBARRAS = "Ocorreu um erro ao tentar ler um registro da tabela ProdutoCodBarras.  Codigo de Barras = %s."
Public Const ERRO_LEITURA_TRIBUTACAODOCITEM = "Ocorreu um erro ao tentar ler um registro da tabela TributacaoDocItem. Produto = %s. "
Public Const ERRO_LEITURA_PRODUTOS_NOME = "Ocorreu um erro ao tentar ler um registro da tabela Produtos.  Produto = %s."
Public Const ERRO_LEITURA_PRODUTOS_REFERENCIA = "Ocorreu um erro ao tentar ler um registro da tabela Produtos.  Referencia = %s."
Public Const ERRO_INCLUSAO_PRODUTOS = "Ocorreu um erro ao tentar gravar um registro na tabela Produtos."
Public Const ERRO_EXCLUSAO_PRODUTOS1 = "Ocorreu um erro ao tentar excluir um registro na tabela Produtos."
Public Const ERRO_INCLUSAO_TRIBUTACAODOCITEM = "Ocorreu um erro ao tentar gravar um registro na tabela TributacaoDocItem."
Public Const ERRO_EXCLUSAO_TRIBUTACAODOCITEM = "Ocorreu um erro ao tentar excluir um registro na tabela TributacaoDocItem."
Public Const ERRO_LEITURA_PRODUTOS_0 = "Ocorreu um erro ao tentar ler um registro da tabela Produtos."
Public Const ERRO_EXCLUSAO_PRODUTOCODBARRAS = "Ocorreu um erro ao tentar excluir um registro na tabela ProdutoCodBarras."
Public Const ERRO_INCLUSAO_PRODUTOCODBARRAS = "Ocorreu um erro ao tentar gravar um registro na tabela ProdutoCodBarras."
Public Const ERRO_ARQ_CRIPTOGRAFADO_VAZIO = "O arquivo criptografado %s est� vazio."
Public Const ERRO_INCLUSAO_CLIENTES = "Ocorreu um erro ao tentar gravar um registro na tabela Clientes."
Public Const ERRO_EXCLUSAO_CLIENTES1 = "Ocorreu um erro ao tentar excluir um registro na tabela Clientes."
Public Const ERRO_LEITURA_CLIENTE_NOMEREDUZIDO = "Ocorreu um erro ao tentar ler um registro da tabela Clientes.  NomeReduzido = %s."
Public Const ERRO_LEITURA_CLIENTES_0 = "Ocorreu um erro ao tentar ler um registro da tabela Clientes."
Public Const ERRO_LEITURA_CLIENTE_CODIGO = "Ocorreu um erro ao tentar ler um registro da tabela Clientes.  Codigo = %s."
Public Const ERRO_INCLUSAO_VENDEDORES = "Ocorreu um erro ao tentar gravar um registro na tabela Vendedores."
Public Const ERRO_LEITURA_VENDEDOR_CODIGO = "Ocorreu um erro ao tentar ler um registro da tabela Vendedores.  Codigo = %s."
Public Const ERRO_EXCLUSAO_VENDEDORES = "Ocorreu um erro ao tentar excluir um registro da tabela Vendedores."
Public Const ERRO_TROCO_DINHEIRO_NEGATIVO = "O Valor do troco em dinheiro n�o pode ser negativo."
Public Const ERRO_TROCO_CONTRAVALE_NEGATIVO = "O Valor do troco em contra vale n�o pode ser negativo."
Public Const ERRO_TROCO_TICKET_NEGATIVO = "O Valor do troco em ticket n�o pode ser negativo."

Public Const ERRO_LEITURA_BACKUPCONFIG = "Ocorreu um erro na leitura da tabela de configura��o de backup"
Public Const ERRO_UPDATE_BACKUPCONFIG = "Ocorreu um erro na altera��o de um registro da tabela de configura��o de backup"
Public Const ERRO_INSERCAO_BACKUPCONFIG = "Ocorreu um erro na inclus�o de um registro da tabela de configura��o de backup"
Public Const ERRO_BACKUP_TOKEN_LIBERAR = "Ocorreu um erro ao tentar liberar o token para backup"
Public Const ERRO_BACKUPCONFIG_NAO_CADASTRADO = "A configura��o de backup com c�digo %s n�o foi configurada."
Public Const ERRO_LEITURA_BDSINFO = "Ocorreu um erro na leitura da tabela dos BDs para backup (BDsInfo)"
Public Const ERRO_BDSINFO_NAO_CADASTRADO = "N�o foi encontrado nenhum BD para backup na tabela BDsInfo."
Public Const ERRO_BACKUP_BDSINFO = "N�o foi poss�vel realizar o backuo. %s"
Public Const AVISO_ERRO_AO_COMPACTAR_O_BACKUP = "Ocorreu um erro na tentativa de compactar o backup"
Public Const AVISO_ERRO_AO_FAZER_O_UPLOAD_DO_BACKUP = "Ocorreu um erro na tentativa fazer o upaload do backup"
Public Const AVISO_ERRO_AO_APAGAR_BACKUPS_ANTIGOS_FTP = "Ocorreu um erro na tentativa de apagar um backup mais antigo"
Public Const ERRO_LEITURA_BACKUPLOG = "Ocorreu um erro na leitura da tabela de Log de backup"
Public Const ERRO_UPDATE_BACKUPLOG = "Ocorreu um erro na altera��o de um registro da tabela de Log de backup"

Public Const ERRO_NENHUMARQ_BKP_ENCONTRADO_FTP = "Nenhum arquivo para download foi encontrado no diret�rio FTP"

Public Const ERRO_LEITURA_SALDOEMDINHEIRO = "Ocorreu um erro na leitura da tabela SaldoEmDinheiro"
Public Const ERRO_INSERT_SALDOEMDINHEIRO = "Ocorreu um erro na inclus�o de um registro na tabela SaldoEmDinheiro"

Public Const ERRO_INCLUSAO_TABELAPRECOITENS = "Ocorreu um erro ao tentar gravar um registro na tabela TabelaPrecoItens."
Public Const ERRO_EXCLUSAO_TABELAPRECOITENS = "Ocorreu um erro ao tentar excluir um registro na tabela TabelaPrecoItens."
Public Const ERRO_LEITURA_TABELAPRECOITENS = "Ocorreu um erro ao tentar ler um registro na tabela TabelaPrecoItens."
Public Const ERRO_INCLUSAO_PRODUTODESCONTO = "Ocorreu um erro ao tentar gravar um registro na tabela ProdutoDesconto."
Public Const ERRO_EXCLUSAO_PRODUTODESCONTO = "Ocorreu um erro ao tentar excluir um registro na tabela ProdutoDesconto."
Public Const ERRO_LEITURA_PRODUTODESCONTO = "Ocorreu um erro ao tentar ler um registro na tabela ProdutoDesconto."

Function Rotina_ErroECF(ByVal MsgBoxTipo As Integer, ByVal sErroId As String, ByVal lCodigo As Long, Optional vParam1 As Variant, Optional vParam2 As Variant, Optional vParam3 As Variant, Optional vParam4 As Variant, Optional vParam5 As Variant, Optional vParam6 As Variant, Optional vParam7 As Variant, Optional vParam8 As Variant, Optional vParam9 As Variant, Optional vParam10 As Variant) As Long

Dim sErro As String, X As Object, Y As New AdmSQL
Dim iMsgBoxRet As Integer, lCodigoInt As Long, sTipoErro As String
Dim lErroAux As Integer, iTipoCliente As Integer, Z As New ClassECFConfig, lTrans As Long

On Error GoTo Erro_Rotina_ErroECF

    iTipoCliente = Sistema_ObterTipoCliente(GL_lSistema)
    
    If InStr(UCase(App.EXEName), "BATCH") <> 0 Then iTipoCliente = AD_SIST_BATCH
    
    If iTipoCliente = AD_SIST_BATCH Then
        lCodigoInt = 0
    Else
        lCodigoInt = lCodigo
    End If
    
    'para depurar o batchest2 como uma dll o trecho abaixo at� os asteriscos deve estar comentado
    'se h� alguma transacao aberta vou forcar o rollback
    If GL_lTransacao <> 0 And iTipoCliente <> AD_SIST_BATCH Then
        Call Y.Transacao_Rollback
    End If
'****

    If Z.glTransacaoPAFECF <> 0 Then
    
        lTrans = Z.glTransacaoPAFECF
        
        'Desfaz Transa��o
        Call Transacao_RollbackExt(lTrans)
        
        Z.glTransacaoPAFECF = 0
    
    End If
    
    If Z.glTransacaoOrcPAFECF <> 0 Then
    
        lTrans = Z.glTransacaoOrcPAFECF
    
        'Desfaz Transa��o
        Call Transacao_RollbackExt(lTrans)
        
        Z.glTransacaoOrcPAFECF = 0
    
    End If
    
    GL_lUltimoErro = lCodigo

    Rotina_ErroECF = 0
    sErro = String(1024, 0)

    'para depurar o batchest como uma dll o comando abaixo deve estar descomentado
'    iTipoCliente = AD_SIST_BATCH
'****
    'Se o erro trazido � ERRO_FORNECIDO_PELO_VB_1
    If sErroId = "1020" Then sErroId = "Erro fornecido pelo VB: %s."
    
    If Not IsMissing(vParam10) Then
        lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), CStr(vParam10))
        If lErroAux Then Exit Function
    Else
        If Not IsMissing(vParam9) Then
            lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), "")
            If lErroAux Then Exit Function
        Else
            If Not IsMissing(vParam8) Then
                lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), "", "")
                If lErroAux Then Exit Function
            Else
                If Not IsMissing(vParam7) Then
                    lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), "", "", "")
                    If lErroAux Then Exit Function
                Else
                    If Not IsMissing(vParam6) Then
                        lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), "", "", "", "")
                        If lErroAux Then Exit Function
                    Else
                        If Not IsMissing(vParam5) Then
                            lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), "", "", "", "", "")
                            If lErroAux Then Exit Function
                        Else
                            If Not IsMissing(vParam4) Then
                                lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), "", "", "", "", "", "")
                                If lErroAux Then Exit Function
                            Else
                                If Not IsMissing(vParam3) Then
                                    lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), "", "", "", "", "", "", "")
                                    If lErroAux Then Exit Function
                                Else
                                    If Not IsMissing(vParam2) Then
                                        lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), CStr(vParam2), "", "", "", "", "", "", "", "")
                                        If lErroAux Then Exit Function
                                    Else
                                        If Not IsMissing(vParam1) Then
                                            lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, CStr(vParam1), "", "", "", "", "", "", "", "", "")
                                            If lErroAux Then Exit Function
                                        Else
                                            lErroAux = Rotina_Erro_Carregar(sErroId, lCodigoInt, sErro, "", "", "", "", "", "", "", "", "", "")
                                            If lErroAux Then Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
'    Call Sistema_RegistrarOcorrencia(GL_lSistema, "Empresa: " & gsNomeEmpresa & " Filial: " & gsNomeFilialEmpresa & " - " & Format(Now(), "General Date") & " - " & "Usu�rio: " & gsUsuario)
'    Call Sistema_RegistrarOcorrencia(GL_lSistema, "   ERRO: " & sTipoErro & " Local: " & CStr(lCodigo) & " Descri��o: " & sErro)
'
    If iTipoCliente <> AD_SIST_BATCH Then
        
        Set X = CreateObject("TelasAdm.ClassTelasAdm")
        
        If X Is Nothing Then Error 58452
        If X.objFormMsgErro Is Nothing Then Error 58453
        
        X.objFormMsgErro.sErro = sErro
        X.objFormMsgErro.lLocalErro = lCodigo
        X.objFormMsgErro.sTipoErro = sErroId
        
        X.objFormMsgErro.Show vbModal
        'Rotina_Erro = MsgBox(sErro, MsgBoxTipo, "SGE - Forprint")
        
    End If
    
    Call ECF_Grava_Log("Codigo: " & CStr(lCodigo) & " ErroId: " & sErroId & " Msg: " & sErro, "ERRO")
    
    Rotina_ErroECF = vbOK

    Exit Function
    
Erro_Rotina_ErroECF:
    
    Select Case Err
    
        Case 58452, 58453
        
        Case Else
            Call MsgBox(Error$, vbOKOnly, "SGE - Forprint")
             
    End Select
        
    Exit Function

End Function

Function Rotina_AvisoECF(ByVal MsgBoxTipo As VbMsgBoxStyle, sErroId As String, Optional vParam1 As Variant, Optional vParam2 As Variant, Optional vParam3 As Variant, Optional vParam4 As Variant, Optional vParam5 As Variant, Optional vParam6 As Variant, Optional vParam7 As Variant, Optional vParam8 As Variant, Optional vParam9 As Variant, Optional vParam10 As Variant) As VbMsgBoxResult
''   Rotina_Aviso = MsgBox("Confirma a opera��o?", MsgBoxTipo, "Rotina temporaria de aviso")
Dim sErro As String
Dim lErroAux As Integer
Dim X As Object, sTipoErro As String
Dim iErro As Integer

On Error GoTo Erro_Rotina_AvisoECF
    
    Rotina_AvisoECF = 0
    sErro = String(1024, 0)
       
    If Not IsMissing(vParam10) Then
        lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), CStr(vParam10))
        If lErroAux Then Exit Function
    Else
        If Not IsMissing(vParam9) Then
            lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), CStr(vParam9), "")
            If lErroAux Then Exit Function
        Else
            If Not IsMissing(vParam8) Then
                lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), CStr(vParam8), "", "")
                If lErroAux Then Exit Function
            Else
                If Not IsMissing(vParam7) Then
                    lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), CStr(vParam7), "", "", "")
                    If lErroAux Then Exit Function
                Else
                    If Not IsMissing(vParam6) Then
                        lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), CStr(vParam6), "", "", "", "")
                        If lErroAux Then Exit Function
                    Else
                        If Not IsMissing(vParam5) Then
                            lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), CStr(vParam5), "", "", "", "", "")
                            If lErroAux Then Exit Function
                        Else
                            If Not IsMissing(vParam4) Then
                                lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), CStr(vParam4), "", "", "", "", "", "")
                                If lErroAux Then Exit Function
                            Else
                                If Not IsMissing(vParam3) Then
                                    lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), CStr(vParam3), "", "", "", "", "", "", "")
                                    If lErroAux Then Exit Function
                                Else
                                    If Not IsMissing(vParam2) Then
                                        lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), CStr(vParam2), "", "", "", "", "", "", "", "")
                                        If lErroAux Then Exit Function
                                    Else
                                        If Not IsMissing(vParam1) Then
                                            lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, CStr(vParam1), "", "", "", "", "", "", "", "", "")
                                            If lErroAux Then Exit Function
                                        Else
                                            lErroAux = Rotina_Erro_Carregar(sErroId, 0, sErro, "", "", "", "", "", "", "", "", "", "")
                                            If lErroAux Then Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Set X = CreateObject("TelasAdm.ClassTelasAdm")
    
    If X Is Nothing Then Error 58450
    If X.objFormMsgAviso Is Nothing Then Error 58451
    
    'Passa o Erro e o Tipo dos Bot�es
    X.objFormMsgAviso.sErro = sErro
    X.objFormMsgAviso.MsgBoxTipo = MsgBoxTipo
    
    X.objFormMsgAviso.Show vbModal
    
    'Retorna o Resultado da escolha
    Rotina_AvisoECF = X.objFormMsgAviso.MsgBoxResultado
    
    Call ECF_Grava_Log(sErro & "|Retorno: " & CStr(Rotina_AvisoECF), "AVISO")
    
'    Rotina_AvisoECF = MsgBox(sErro, MsgBoxTipo, "SGE - Forprint")

    Exit Function
    
Erro_Rotina_AvisoECF:
    
    Select Case Err
    
        Case 58450, 58451
            Call MsgBox("Erro na exibi��o de uma pergunta (ou aviso).", vbOKOnly, "SGE - Forprint")
        
        Case Else
            Call MsgBox(Error$, vbOKOnly, "SGE - Forprint")
             
    End Select
        
    Error 59060 'para que quem chamou nao prossiga "normalmente" na operacao
    
    Exit Function
    
End Function

Public Function ECF_Grava_Log(ByVal sTexto As String, Optional ByVal sTipo As String = "INFO") As Long

Dim bAbriuArq As Boolean
Dim sNomeArq As String
Dim lFN As Long, iIndice As Integer

On Error GoTo Erro_ECF_Grava_Log

    bAbriuArq = False

    sNomeArq = CurDir & "\ECFLog.txt"
    
    lFN = FreeFile
    Open sNomeArq For Append As lFN
    bAbriuArq = True
    
    sTexto = Replace(sTexto, vbNewLine, "-")
    sTexto = Replace(sTexto, Chr$(10), "")
    sTexto = Replace(sTexto, Chr$(13), "")
    sTexto = Replace(sTexto, Chr$(0), "")
    sTexto = Trim(sTexto)
    
    Print #lFN, Format(Now(), "General Date") & " |" & sTipo & "| " & sTexto
    

    Close #lFN

    ECF_Grava_Log = SUCESSO

    Exit Function
    
Erro_ECF_Grava_Log:

    If bAbriuArq Then Close #lFN
    
    ECF_Grava_Log = Err
    
    Exit Function
    
End Function

Public Function ValidEmail(ByVal strCheck As String) As Boolean
'Created by Chad M. Kovac
'Tech Knowledgey, Inc.
'http://www.TechKnowledgeyInc.com

Dim bCK As Boolean
Dim strDomainType As String
Dim strDomainName As String
Const sInvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
Dim i As Integer

bCK = Not InStr(1, strCheck, Chr(34)) > 0 'Check to see if there is a double quote
If Not bCK Then GoTo ExitFunction

bCK = Not InStr(1, strCheck, "..") > 0 'Check to see if there are consecutive dots
If Not bCK Then GoTo ExitFunction

' Check for invalid characters.
If Len(strCheck) > Len(sInvalidChars) Then
    For i = 1 To Len(sInvalidChars)
        If InStr(strCheck, Mid(sInvalidChars, i, 1)) > 0 Then
            bCK = False
            GoTo ExitFunction
        End If
    Next
Else
    For i = 1 To Len(strCheck)
        If InStr(sInvalidChars, Mid(strCheck, i, 1)) > 0 Then
            bCK = False
            GoTo ExitFunction
        End If
    Next
End If

If InStr(1, strCheck, "@") > 1 Then 'Check for an @ symbol
    bCK = Len(left(strCheck, InStr(1, strCheck, "@") - 1)) > 0
Else
    bCK = False
End If
If Not bCK Then GoTo ExitFunction

strCheck = right(strCheck, Len(strCheck) - InStr(1, strCheck, "@"))
bCK = Not InStr(1, strCheck, "@") > 0 'Check to see if there are too many @'s
If Not bCK Then GoTo ExitFunction

strDomainType = right(strCheck, Len(strCheck) - InStr(1, strCheck, "."))
bCK = Len(strDomainType) > 0 And InStr(1, strCheck, ".") > 0 And InStr(1, strCheck, ".") < Len(strCheck)
If Not bCK Then GoTo ExitFunction

strCheck = left(strCheck, Len(strCheck) - Len(strDomainType) - 1)
Do Until InStr(1, strCheck, ".") <= 1
    If Len(strCheck) >= InStr(1, strCheck, ".") Then
        strCheck = left(strCheck, Len(strCheck) - (InStr(1, strCheck, ".") - 1))
    Else
        bCK = False
        GoTo ExitFunction
    End If
Loop
If strCheck = "." Or Len(strCheck) = 0 Then bCK = False

ExitFunction:
    ValidEmail = bCK
End Function

Public Function TrocaFoco(objTela As Object, objControle As Object) As Boolean

    Dim bTrocaFoco As Boolean, objAux As Object
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")

    bTrocaFoco = False
    
    If Not (objTela.ActiveControl Is Nothing) Then
    
        Set objAux = objTela.ActiveControl
        Call WshShell.SendKeys("{Tab}", True)
        DoEvents
        If Not (objTela.ActiveControl Is Nothing) Then
            If objAux.Name <> objTela.ActiveControl.Name Then
                bTrocaFoco = True
                If Not (objControle Is Nothing) Then objControle.SetFocus
            End If
        End If
    End If
    
    TrocaFoco = bTrocaFoco

End Function

