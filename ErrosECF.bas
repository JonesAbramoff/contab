Attribute VB_Name = "ErrosECF"
Option Explicit

Public GL_ConexaoOrc As Long

Public Const CORPORATOR_ECF_VERSAO_PGM = "537"
Public Const CORPORATOR_ECF_VERSAO_BD_ECF = "8350"
Public Const CORPORATOR_ECF_VERSAO_BD_ORC = "8137"

Public Const TEF_PAYGO_CERTIFICADO = "03252" 'Código obtido junto à NTK Solutions no início do processo de certificação da Automação Comercial

Private Declare Function Sistema_ObterTipoCliente Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterTipoCliente" (ByVal lID_Sistema As Long) As Long
Private Declare Function Rotina_Erro_Carregar Lib "ADCRTL.DLL" (ByVal sTipoErro As String, ByVal lLocalErro As Long, ByVal lpMsgErro As String, ByVal sParam1 As String, ByVal sParam2 As String, ByVal sParam3 As String, ByVal sParam4 As String, ByVal sParam5 As String, ByVal sParam6 As String, ByVal sParam7 As String, ByVal sParam8 As String, ByVal sParam9 As String, ByVal sParam10 As String) As Long

Private Const AD_SIST_NORMAL = 0
Private Const AD_SIST_RELLIB = 1
Private Const AD_SIST_BATCH = 2

'Erros
Public Const ERRO_EM_PROCESSAMENTO_SEFAZ = "Aguardando processamento do envio da nfce pela sefaz"
Public Const ERRO_NFCEINFO_NAO_CADASTRADO = "A NFCe com chave de acesso %s não foi encontrada no banco de dados."
Public Const ERRO_NFCEINFO_CANC_NAO_CADASTRADO = "A NFCe com chave de acesso %s não está registrada como cancelada."
Public Const ERRO_LEITURA_NFCEINFO = "Erro na leitura da tabela NFCeInfo."
Public Const ERRO_NFCE_OFFLINE_NAO_TRANSMITIDA = "Precisa transmitir para a retaguarda os arquivos xml das nfce offline antes de autorizá-las. Vá em Funções/Operações de Arquivo/Transmitir."
Public Const ERRO_XJUST_INVALIDO = "O motivo da contingencia precisa ter no mínimo 15 caracteres."
Public Const AVISO_CANCELA_ITEM_CUPOM = "Deseja realmente cancelar um item desta venda ?"
Public Const ERRO_DOWNLOAD_ARQUIVO_FTP1 = "Erro na recepção do arquivo via ftp: %s"
Public Const ERRO_UPLOAD_ARQUIVO_FTP1 = "Erro no envio do arquivo via ftp: %s"
Public Const ERRO_COMUNICACAO_FTP = "Erro no download do arquivo via ftp"
Public Const ERRO_ARQUIVO_FTP_NAO_ENCONTRADO = "arquivo não econtrado no servidor ftp: %s"
Public Const ERRO_EMAIL_INVALIDO = "O email digitado não é válido."
Public Const ERRO_SAT_CODIGO_ATIVACAO = "O código de ativação possui entre 8 e 32 caracteres"
Public Const ERRO_LOCAL_ARQUIVO_NAO_CONFIGURADO = "O local do arquivo %s não foi configurado." 'Incluído por Luiz Nogueira em 16/04/04
Public Const ERRO_FORNECIDO_PELO_VB_1 = "Erro fornecido pelo VB: %s. (%s)"
Public Const ERRO_ARQUIVO_NAO_RETRANSMITIDO = "O arquivo não pode ser retransmitido pois o arquivo já existe. Delete-o e tente novamente."
Public Const ERRO_FIGURA_INVALIDO = "O arquivo referente a Figura %s deste produto não existe no sistema."
Public Const ERRO_CAIXA_NAO_EXISTENTE = "Caixa Padrão não encontrado. Verifique o arquivo de transferência."
Public Const ERRO_ARQUIVO_ABERTO = "O Arquivo está aberto. Tente novamente"
Public Const ARQUIVO_INICIALIZACAO_NAO_EXISTENTE = "O Arquivo de inicialização %s não foi encontrado."
Public Const ERRO_ARQUIVO_DIFERENTE = "O Arquivo do caixa está com dados pertencentes a outro caixa."
Public Const ERRO_ITEM_NAO_PREENCHIDO1 = "O Item deve ser preenchido."
Public Const ERRO_ITEM_CANCELADO = "O Item já foi cancelado."
Public Const ERRO_ITEM_NAO_EXISTENTE1 = "O Item não existe neste cupom."
Public Const ERRO_QUANTIDADE_INVALIDA = "O valor da quantidade não pode superar 4 dígitos inteiros."
Public Const ERRO_PARCELAMENTO_NAO_EXISTENTE = "O parcelamento selecionado não existe, selecione outro e repita a operação."
Public Const ERRO_ADMMEIOPAGTO_NAO_EXISTENTE = "O AdmMeioPagto selecionado não existe, selecione outro e repita a operação."
Public Const ERRO_VALOR_JA_PAGO = "Não existe valor para ser pago por TEF."
Public Const ERRO_DDD_NAO_PREENCHIDO = "O DDD não foi preenchido"
Public Const ERRO_OPERADORA_NAO_PREENCHIDA = "A Operadora não foi preenchida"
Public Const ERRO_MEIOSPAG_ULTRAPASSAM = "Os meios de pagamento informados ultrapassam o valor a ser pago."
Public Const ERRO_CUPOM_NAO_CANCELADO = "Não existe cupom anterior a ser cancelado."
Public Const ERRO_VENDEDOR_NAO_PREENCHIDO = "O vendedor não foi preenchido"
Public Const ERRO_NUMEROSERIE_NAO_PREENCHIDO = "O Nº de Série deve estar preenchido."
Public Const ERRO_DATABOMPARA_MENOR = "A data para depósito do cheque não pode ser menor que a data Atual."
Public Const ERRO_SANGRIA_MAIOR = "O valor da Sangria %s não pode ser maior do que o saldo %s em caixa."
Public Const ERRO_REDUCAO_JA_EXECUTADA = "Foi executada a Redução Z para a data de %s(hoje), nenhum movimento pode ser executado, só serão permitidas consultas."
Public Const ERRO_NOME_TIPOMEIOPAGTO_NAO_EXISTENTE = "O Tipo Meio Pagto com nome: %s não foi cadastrado."
Public Const ERRO_CODIGO_TIPOMEIOPAGTO_NAO_EXISTENTE = "O Tipo Meio Pagto com código: %s não foi cadastrado."
Public Const ERRO_NOME_CARTAO_NAO_EXISTENTE = "O Cartão com nome: %s não foi cadastrado."
Public Const ERRO_CODIGO_CARTAO_NAO_EXISTENTE = "O Cartão com código: %s não foi cadastrado."
Public Const ERRO_NOME_PARCELAMENTO_NAO_EXISTENTE = "O Parcelamento com nome: %s não foi cadastrado."
Public Const ERRO_CODIGO_PARCELAMENTO_NAO_EXISTENTE = "O Parcelamento com código: %s não foi cadastrado."
Public Const ERRO_NOME_VALETICKET_NAO_EXISTENTE = "O Vale/Ticket com nome: %s não foi cadastrado."
Public Const ERRO_CODIGO_VALETICKET_NAO_EXISTENTE = "O Vale/Ticket com código: %s não foi cadastrado."
Public Const ERRO_TAMANHO_CGC_CPF = "O Tamanho do CPF/CGC do cliente está incorreto, ele deve ter 11 ou 14 caracteres."
Public Const ERRO_TIPOMEIOPAGTODE_NAO_PREENCHIDO = "O Tipo Meio Pagto de origem não foi preenchido."
Public Const ERRO_TIPOMEIOPAGTOPARA_NAO_PREENCHIDO = "O Tipo Meio Pagto de destino não foi preenchido."
Public Const ERRO_DATABOMPARA_NAO_PREENCHIDA = "A Data de Depósito (Bom Para) não foi preenchida."
Public Const ERRO_GRUPOCHQDE_NAO_PREENCHIDO = "Se algum dos campos do agrupamento de origem: Banco, Agência, Conta, Cliente ou Número estiverem preenchidos, todos os campos desse agrupamento devem ser preenchidos."
Public Const ERRO_GRUPOCHQPARA_NAO_PREENCHIDO = "Se algum dos campos do agrupamento de destino: Banco, Agência, Conta, Cliente ou Número estiverem preenchidos, todos os campos desse agrupamento devem ser preenchidos."
Public Const ERRO_GRUPOCRTDE_NAO_PREENCHIDO = "Se algum dos campos do agrupamento de origem: Cartão, Parcelamento estiverem preenchidos, todos os campos desse agrupamento devem ser preenchidos."
Public Const ERRO_GRUPOCRTPARA_NAO_PREENCHIDO = "Se algum dos campos do agrupamento de destino: Cartão, Parcelamento estiverem preenchidos, todos os campos desse agrupamento devem ser preenchidos."
Public Const ERRO_VALOR_NAO_PREENCHIDO_ORIGEM = "O Valor não foi preenchido."
Public Const ERRO_TRANSFCAIXA_NAO_PREENCHIDO = "O Código da transferência deve ser preenchido."
Public Const ERRO_NECESSARIO_FECHAR_APP = "Ocorreu um erro na gravação dos arquivos. Clique em OK para encerrar a aplicação."
Public Const ERRO_VALORDIN_MAIOR_SALDODIN = "O Valor da Transferência em questão é maior que o saldo em dinheiro. Valor: %s    Saldo: %s"
Public Const ERRO_VALORCHQ_MAIOR_SALDOCHQ = "O Valor da Transferência em questão é maior que o saldo dos cheques. Valor: %s    Saldo: %s"
Public Const ERRO_CHQESPECIFIC_VINCULADO_BANCO = "O cheque em questão não foi especificado, mas no entanto possui vínculo com o banco: %s"
Public Const ERRO_VALORCRT_MAIOR_SALDOCRT = "O Valor da Transferência em questão é maior que o saldo do Cartão. Valor: %s    Saldo: %s        AdmMeioPagto: %s   Parcelamento: %s"
Public Const ERRO_VALORADM_MAIOR_SALDOADM = "O Valor da Transferência em questão é maior que o saldo do  Meio Pagto. Valor: %s    Saldo: %s        AdmMeioPagto: %s   Parcelamento: %s"
Public Const ERRO_VALORTKT_MAIOR_SALDOTKT = "O Valor da Transferência em questão é maior que o saldo do Ticket. Valor: %s    Saldo: %s        AdmMeioPagto: %s   Parcelamento: %s"
Public Const ERRO_TRANSFERENCIA_NAO_EXISTENTE = "A Transferência de Caixa com código: %s  não foi cadastrada."
Public Const ERRO_VALOR_CHQESPEC_DIF_VALORTELA = "O Valor do cheque especificado em questão deve ser %s, pois o cheque deve ser transferido integralmente."
Public Const ERRO_VALORSANGRIA_NAO_INFORMADO = "O valor da sangria não foi informado."
Public Const ERRO_VALORSANGRIA_NAO_INFORMADO_GRID = "O valor da sangria não foi informado na linha %s."
Public Const ERRO_CARNE_NAO_EXISTENTE = "O carne não existe."
Public Const ERRO_PARCELA_NAO_SELECIONADA = "Deveria ter pelo menos uma Parcela selecionada no grid."
Public Const ERRO_CGC_OVERFLOW1 = "Valor %s ultrapassa o limite de CGC."
Public Const ERRO_NUMERO_NAO_INTEIRO1 = "O número %s não é inteiro."
Public Const ERRO_NUMERO_NAO_POSITIVO1 = "O número %s não é positivo."
Public Const ERRO_CLIENTE_NAO_EXISTENTE = "O Cliente não foi encontrado."
Public Const ERRO_VALOR_SANGRIACHEQUE_INEXISTENTE = "não exitem cheques a serem sangrados."
Public Const ERRO_CODIGO_NAO_PREENCHIDO1 = "O campo obrigatório código não esta preenchido."
Public Const ERRO_DESCONTO_MAIOR = "O desconto dado é maior que o total."
Public Const ERRO_ACRESCIMO_MAIOR = "O acréscimo dado é maior que o total."
Public Const ERRO_NAO_EXISTE_ITEM = "Deve existir pelo menos um item para ser cancelado."
Public Const ERRO_ITEM_NAO_EXISTENTE = "Deve existir pelo menos um item de Venda."
Public Const ERRO_ABERTURA_TRANSACAO1 = "Não conseguiu abrir a transação."
Public Const ERRO_COMMIT_TRANSACAO1 = "Não confirmou a transação."
Public Const ERRO_ORCAMENTO_NAO_PREENCHIDO = "O Número do Orçamento deve estar preenchido."
Public Const ERRO_IMPRESSORA_NAO_RESPONDE = "A Impressora não responde."
Public Const ERRO_TEF_NAO_ATIVO = "O Gerenciador Padrão não esta ativo."
Public Const ERRO_SANGRIA_SUPRIMENTO_NAO_SELECIONADO = "Erro deve ser selecionado alguma operação, de sangria ou de suprimento."
Public Const ERRO_CAIXA_FECHADO = "O caixa %s está fechado."
Public Const ERRO_TRANSACAO_CAIXA_ABERTA = "Já Existe uma Transação Aberta para o caixa de Código %s."
Public Const ERRO_NUMERO_INTERVALO_REDUCAO_INICIAL_MAIOR = "O número correspondente ao primeiro intervalo de redução não pode ser maior que o segundo."
Public Const ERRO_DATA_INICIAL_MAIOR1 = "Data inicial não pode ser maior que a data final."
Public Const ERRO_DATAS_MEMORIAFISCAL_NAO_PREENCHIDA = "A leitura de memória fiscal por intervalo de datas não está preenchido, os intervalos são campos obrigatórios."
Public Const ERRO_REDUCAO_NAO_PREENCHIDA = "A leitura de memória fiscal por intervalo de redução não está preenchido, os intervalos são campos obrigatórios."
Public Const ERRO_OPERADOR_NAO_SELECIONADO = "Nenhum Operador foi selecionado."
Public Const ERRO_SENHA_NAO_PREENCHIDA1 = "O preenchimento da Senha é obrigatório."
Public Const ERRO_SENHA_INVALIDA1 = "Senha Inválida."
Public Const ERRO_TROCO_NAO_ESPECIFICADO = "O valor do troco deve estar especificado."
Public Const ERRO_ADMMEIOPAGTO_NAO_SELECIONADO = "Um AdmMeioPagto deve estar selecionado."
Public Const ERRO_PARCELAMENTO_NAO_SELECIONADO = "Um Parcelamento deve estar selecionado."
Public Const ERRO_VALOR_NAO_PREENCHIDO2 = "O campo Valor não foi prenchido."
Public Const ERRO_DATA_NAO_PREENCHIDA1 = "O preenchimento da Data é obrigatório."
Public Const ERRO_QUANTIDADE_NAO_PREENCHIDO1 = "A quantidade não foi informada."
Public Const ERRO_TICKET_NAO_SELECIONADO = "Um Ticket deve estar selecionado."
Public Const ERRO_VALOR_NAO_EXISTENTE = "Não existe valor para ser pago na tela."
Public Const ERRO_TROCO_MAIOR = "O Valor do Troco é maior que o devido."
Public Const ERRO_VENDA_ANDAMENTO = "A Sessão não pode ser suspensa. Existe uma venda em andamento."
Public Const ERRO_TERMINAL_NAO_SELECIONADO = "Um terminal deve estar selecionado."
Public Const ERRO_SESSAO_ABERTA = "Já existe uma sessão aberta no caixa %s para o operador %s."
Public Const ERRO_SESSAO_SUSPENSA = "A sessão do caixa %s está suspensa."
Public Const ERRO_SESSAO_ABERTA_INEXISTENTE = "Não existe sessão aberta para o caixa %s."
Public Const ERRO_CAIXA_ABERTO = "O caixa %s já esta aberto."
Public Const ERRO_TROCO_DIFERENTE = "O subtotal é diferente do Troco que deve ser informado."
Public Const ERRO_CARNE_MAIOR = "O Valor do Carnê é maior que o que deve ser cobrado."
Public Const ERRO_VENDEDOR_NAO_CADASTRADO2 = "O Vendedor não está cadastrado no Banco de Dados."
Public Const ERRO_PRODUTO_NAO_CADASTRADO1 = "O Produto %s não está cadastrado na tabela Produtos."
Public Const ERRO_PRODUTO_NAO_PREENCHIDO1 = "O Produto deve estar preenchido."
Public Const ERRO_PRODUTO_SEM_TABELAPRECO_PADRAO1 = "Não existe preço cadastrado para o produto em questão."
Public Const ERRO_PRODUTO_SEM_PRECO = "O Produto %s não tem um preço associado na Tabela de Preços do Loja."
Public Const ERRO_TECLADOPADRAO_NAO_EXISTE = "Não existe um Teclado Padrão."
Public Const ERRO_VALOR_INSUFICIENTE = "O Valor é insuficiente para o pagamento da compra."
Public Const ERRO_SALDO_INSUFICIENTE_SANGRIA = " O saldo %s do caixa %s não é sucifiente para efetuar a sangria desejada %s ."
Public Const ERRO_CHEQUE_NAO_SELECIONADO = "Pelo menos um cheque deve estar selecionado."
Public Const ERRO_BANCO_NAO_PREENCHIDO = "O campo Banco deve estar preenchido."
Public Const ERRO_AGENCIA_NAO_PREENCHIDA = "O campo Agencia deve estar preenchida."
Public Const ERRO_NUMERO_NAO_PREENCHIDO = "O campo Número deve estar preenchido."
Public Const ERRO_CLIENTE_NAO_PREENCHIDO1 = "O campo Cliente deve estar preenchido."
Public Const ERRO_CONTA_NAO_PREENCHIDA = "O campo Conta deve estar preenchida."
Public Const ITEM_CUPOM_NAO_SELECIONADO = "Um Item do Cupom deve estar selecionado."
Public Const ERRO_DATAEMISSAO_MAIOR = "A Data de Emissão é maior que a Data Final."
Public Const ERRO_AUTORIZACAO_NAO_PREENCHIDA = "O campo Autorização deve estar preenchido."
Public Const ERRO_VALORSANGRIA_NAO_DISPONIVEL = "O valor determinado para a Sangria %s é maior do que o valor total em caixa %s."
Public Const ERRO_LINHA_REPETIDA = " Existe uma linha duplicada no grid."
Public Const ERRO_SANGRIA_NAOPODE_SER_EXECUTADA = "A sangria não pode ser executada com valor zero."
Public Const ERRO_MOVIMENTO_INEXISTENTE = "O Movimento %s não existe no sistema."
Public Const ERRO_CARREGA_DADOS_TIPOSMEIOPAGTO = "Erro no carregamento de dados de Tipos de Meio de Pagamento para o Caixa ECF"
Public Const ERRO_CARREGAMENTO_CAIXA_CONFIG = "Erro no carregamento dos dados de configuração do Caixa"
Public Const ERRO_PARCELAMENTO_NAO_PREENCHIDO = "O campo Parcelamento deve estar preenchido."
Public Const ERRO_CLIENTE_NAO_CADASTRADO2 = "O cliente em questão não está cadastrado."
Public Const ERRO_TIPOMEIOPAGTO_NAO_CARTAO = "O Tipo Meio Pagto deveria ser cartão de crédito ou débito, mas outro valor foi encontrado. Tipo = %s."
Public Const ERRO_PREENCHIMENTO_ARQUIVO_CONFIG = "O item %s da seção %s do arquivo de configuração %s não foi preenchido corretamente"
Public Const ERRO_INTERVALO_NAO_PREENCHIDO = "Alguma data não está preenchida, os intervalos são campos obrigatórios."
Public Const ERRO_ECF_NAO_PREENCHIDO = "O ECF não foi selecionado."
Public Const ERRO_TEF_NAO_RESPONDE = "TEF não responde."
Public Const ERRO_INCONSISTENCIA_ARQUIVO_TEF = "Inconsistencia no campo %s do arquivo %s gerado pelo TEF. %s."
Public Const ERRO_TEF_TRANSACAO_NEGADA = "%s."
Public Const ERRO_TEF_FALHA_IMPRESSAO = "Transação TEF cancelada. Rede: %s  NSU: %s  Valor: %s ."

Public Const ERRO_CHEQUE_CUPOMFISCAL_NAOENCONTRADO = "Não foi encontrado o cheque associado ao cupom fiscal %s"
Public Const ERRO_NOME_OUTROS_NAO_EXISTENTE = "O Meio de Pagamento com nome: %s não foi cadastrado."
Public Const ERRO_CODIGO_OUTROS_NAO_EXISTENTE = "O Meio de Pagamento com código: %s não foi cadastrado."
Public Const ERRO_OUTROS_NAO_SELECIONADO = "Um meio de pagamento deve estar selecionado."
Public Const ERRO_VALOR_NAO_PREENCHIDO_DESTINO = "O Campo Valor de destino não foi preenchido."
Public Const ERRO_CUPOMFISCAL_NAO_PREENCHIDO = "O cupom fiscal deve estar preenchido."
Public Const ERRO_CHEQUEPRE_NAO_ENCONTRADO2 = "O Cheque com sequencial %s não foi encontrado."
Public Const ERRO_TRANSFCAIXA_NAO_ENCONTRADO = "A transferência de caixa com código %s não foi encontrada."
Public Const ERRO_CHEQUEPRE_OUTRA_TRANSFERENCIA = "O cheque referente à transferência %s não corresponde ao que está na tela."
Public Const ERRO_CUPOMFISCAL_NAO_ENCONTRADO = "O cupom fiscal %s do ECF %s não foi encontrado."
Public Const ERRO_CHEQUEPRE_EXCLUIDO = "O Cheque com sequencial %s está marcado como excluido. "
Public Const ERRO_TIPOMOVTOCAIXA_INVALIDO = "O Tipo de Movimento de Caixa em questão é inválido neste contexto. Tipo = %s "
Public Const ERRO_MOVIMENTO_NAO_DINHEIRO = "O Movimento %s não é uma movimentação de dinheiro."
Public Const ERRO_MOVIMENTO_NAO_CARTAO = "O Movimento %s não é uma movimentação de Cartão."
Public Const ERRO_MOVIMENTO_NAO_CHEQUE = "O Movimento %s não é uma movimentação de cheques."
Public Const ERRO_MOVIMENTO_NAO_TICKET = "O Movimento %s não é uma movimentação de Tickets."
Public Const ERRO_MOVIMENTO_NAO_OUTROS = "O Movimento %s não é uma movimentação de Outros."
Public Const ERRO_MOVIMENTO_JA_TRANSMITIDO = "Essa operação não pode ser efetuada pois este movimento já foi transferido para o Caixa Central."
Public Const ERRO_CHEQUE_SANGRADO = "O Cheque com sequencial %s já foi sangrado. "
Public Const ERRO_ARQUIVO_EXISTENTE = "O arquivo %s já existe."
Public Const ERRO_CAIXA_FECHADO_SESSAO_NAO = "O caixa está fechado, no entanto a sessão esta indicando que não."
Public Const ERRO_VALOR_JA_PAGO1 = "Não existe valor para ser pago."
Public Const ERRO_TROCO_MAIOR_DINHEIRO = "O Valor do Troco é maior que o valor pago em dinheiro."
Public Const ERRO_CHEQUE_MAIOR_FALTA = "O Valor do Cheque ultrapassa a quantidade que falta ser paga = %s."
Public Const ERRO_PRECO_NAO_PREENCHIDO = "O preço unitário do produto não está preenchido."
Public Const ERRO_DAV_NOME_CLIENTE_NAO_PREENCHIDO = "Para o DAV, o nome do cliente tem que estar preenchido, antes de gravar."
Public Const ERRO_DAV_CPFCNPJ_CLIENTE_NAO_PREENCHIDO = "Para o DAV, o cpf/cnpj do cliente tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_NOME_CLIENTE_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o nome do cliente tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_CPFCNPJ_CLIENTE_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o cpf/cnpj do cliente tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_LOGRADOURO_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o logradouro do endereço entrega tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_NUMERO_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o numero do endereço de entrega tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_COMPL_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o complemento do endereço de entrega tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_BAIRRO_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, o bairro do endereço de entrega tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_CIDADE_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, a cidade do endereço de entrega tem que estar preenchido, antes de gravar."
Public Const ERRO_NFE_UF_ENTREGA_NAO_PREENCHIDO = "Para a Nota Fiscal Eletronica, a UF do endereço de entrega tem que estar preenchido, antes de gravar."


Public Const AVISO_NFCE_XMLS_PENDENTES_ENVIADOS = "Os arquivos xml de nfce emitidos em contingencia offline foram enviados para a sefaz."
Public Const AVISO_SEQUENCIAL_NAO_ENCONTRADO_ARQUIVAO = "O sequencial %s não foi encontrado no arquivo %s e portanto não poderá ser retransmitido."
Public Const AVISO_EXCLUSAO_TRANSFCAIXA = "Confirma a exclusão da transferência com código %s.?"
Public Const AVISO_ALTERACAO_MOVIMENTOCAIXA = "O Movimento de Caixa %s já existe no Sistema, Deseja Realmente efetuar a alteração ?"
Public Const AVISO_EXCLUSAO_MOVIMENTOCAIXA = "Deseja Realmente Excluir o movimento de caixa %s ? "
Public Const AVISO_ALTERACAO_MOVIMENTOCAIXA1 = "Deseja Salvar as Alterações ?"
Public Const AVISO_ORCAMENTO_INEXISTENTE = "O Orçamento não existe."
Public Const AVISO_DESEJA_ABRIR_SESSAO = " Não Existe sessão aberta para o caixa %s. Deseja efetuar a abertura de sessão ?"
Public Const AVISO_ORCAMENTO_VENDA = "Deseja converter o Orçamento em Venda?"
Public Const AVISO_CANCELA_CUPOM = "Deseja cancelar o Cupom em andamento?"
Public Const AVISO_CANCELA_VENDA = "Deseja cancelar a venda em andamento?"
Public Const AVISO_CANCELA_ORCAMENTO = "O orçamento vai ter as informações não gravadas perdidas. Deseja continuar?"
Public Const AVISO_CANCELAR_COMPRA = "Deseja cancelar a compra?"
Public Const AVISO_CANCELAR_TICKET = "Deseja realmente cancelar os tickets especificados?"
Public Const AVISO_APROVEITA_PAGAMENTO = "Deseja que os pagamentos sejem aproveitados?"
Public Const AVISO_DESEJA_ABRIR_CAIXA = "O caixa %s ainda não foi aberto. Deseja efetuar a abertura?"
Public Const AVISO_DESEJA_FAZER_SANGRIA_COMPLETA = "Deseja efetuar uma sangria completa do caixa %s ?"
Public Const AVISO_SESSAO_ABERTA = "Já existe uma sessão aberta no caixa %s para o Operador %s."
Public Const AVISO_REDUCAOZ_ENCERRA_DIA = "A execução da ReduçãoZ não permitirá o registro de novos movimentos nesta data (%s) para o caixa %s. Deseja Proceguir?"
Public Const AVISO_ALTERACAO_TRANSFCAIXA = "Tem certeza de que deseja alterar a Transferência de código: %s."
Public Const AVISO_SEM_REGISTRO1 = "Não existe nenhum registro para a seleção executada."
Public Const AVISO_SALVAR_ALTERACOES_CARNE = "Deseja salvar os Cancelamentos dos Carnês selecionados?"
Public Const AVISO_SEM_REGISTRO = "Não existe nenhum registro para a seleção executada."
Public Const AVISO_NOVAS_CONFIGURACOES = "Deseja atualizar as parcelas com o novo filtro de Cliente?"
Public Const AVISO_LIMPAR_TELA = "Deseja realmente limpar e perder todas as informações presentes na tela? "
Public Const AVISO_IMPRESSORA_NAO_RESPONDE = "A Impressora não responde. Deseja tentar novamente?"
Public Const AVISO_CANCELA_CUPOM_TELA = "Deseja realmente cancelar o cupom da tela?"
Public Const AVISO_CANCELA_CUPOM_ANTERIOR = "Deseja realmente cancelar o cupom ANTERIOR?"
Public Const AVISO_DESEJA_REDUCAOZ = "A Redução Z implica no fechamento do Caixa e mais nenhuma venda poderá ser efetuada no dia %s. Deseja realmente prosseguir?"
Public Const AVISO_FECHAR_SESSAO_ABERTA = "Existe uma Sessão aberta. Deseja fechá-la?"
Public Const AVISO_CONTINUAR_IMPRESSAO = "Deseja reimprimir o TEF interrompido?"
Public Const AVISO_ORCAMENTO_IMPRESSAO = "Deseja imprimir o Orçamento?"
Public Const AVISO_INICIALIZAR_SISTEMA_AGORA = "O sistema vai ser reinicializado automaticamente a partir de agora."
Public Const AVISO_INICIALIZAR_SISTEMA = "Devido a mudança de dia será necessário reinicializar o sistema, caso contrário isto será feito após à 2:00."
Public Const AVISO_NAO_INICIALIZADO_TERMINAL = "O Terminal não foi reconfigurado."
Public Const AVISO_RECONFIGURAR_VISANET = "Deseja realmente inicializar o TEF do Visanet?(Isto só deve ser feito em caso de reconfiguração)"
Public Const AVISO_ARQUIVO_NAO_ENCONTRADO = "O arquivo %s não foi encontrado. Deseja prosseguir?"
Public Const AVISO_ARQUIVO_TRANSMITIDO = "O arquivo %s foi criado com sucesso."
Public Const AVISO_NUM_CHEQUES_MAIOR_NUM_MAX_GRID = "O número de cheques no caixa ultrapassa o limite de %s do grid. Faça a sangria de alguns para poder ver os demais. Total de Cheque = %s."
Public Const AVISO_ARQUIVOS_RETRANSMITIDOS = "O(s) arquivo(s):%s foi(ram) regerado(s) com sucesso."
Public Const AVISO_TABELAS_ATUALIZADAS = "Tabelas atualizadas com sucesso."
Public Const AVISO_CANCELAR_ACOMPANHAMENTO = "A geração do arquivo %s irá continuar. Deseja cancelar o acompanhamento da geração?"
Public Const AVISO_ARQUIVO_GERADO = "O arquivo %s foi gerado."
Public Const AVISO_NAO_INDICETECNICO = "Este PAF-ECF não executa funções de baixa de estoque com base em índices técnicos de produção, não podendo ser utilizando por estabelecimento que necessite deste recurso."
Public Const AVISO_DESEJA_PROSSEGUIR_ACOMPANHAMENTO_ARQUIVO = "Deseja continuar o acompanhamento da geração do arquivo %s ?"
Public Const AVISO_CANCELAR_ACOMPANHAMENTO_ARQUIVO = "Confirma o cancelamento do acompanhamento da geração do arquivo %s ?"
Public Const AVISO_DESEJA_CPF_NA_NOTA = "Deseja colocar CPF na nota?"
Public Const AVISO_CAIXA_ABERTA = "Caixa Aberta."
Public Const AVISO_CANCELA_CUPOM_TELA_SUCESSO = "A venda foi cancelada com sucesso."

Public Const ERRO_NFCE_SEM_OFFLINE_PENDENTE = "Não há xml de nfce pendente de envio para a sefaz."
Public Const ERRO_NFCE_INFO_LEITURA = "Erro na leitura de informações de nfce"
Public Const ERRO_NFCE_INFO_UPDATE = "Erro na atualização de informações de nfce com a chave %s"
Public Const ERRO_NFCE_OFFLINE_PENDENTE = "Não foi encontrado xml pendente de nfce com a chave %s"
Public Const ERRO_CHEQUE_SANGRIA_NAOENCONTRADO = "Não foi encontrado nenhum cheque associado a sangria em questão"
Public Const ERRO_ARQUIVO_NAO_ENCONTRADO1 = "Não conseguiu gravar dados na seção %s item %s do arquivo %s."
Public Const ERRO_PARCELAMENTO_NAO_EXISTENTE1 = "O parcelamento %s não existe."
Public Const ERRO_PARCELAMENTO_NAO_PREENCHIDO_GRID = "O Parcelamento não foi preenchido na linha %s."
Public Const ERRO_ADMINISTRADORA_NAO_PREENCHIDO_GRID = "A Administradora não foi preenchida na linha %s."
Public Const ERRO_SANGRIA_NAOESP_NAO_DISPONIVEL = "O valor para a Sangria de boletos não especificados %s é maior do que o valor total em caixa %s."
Public Const ERRO_CHEQUE_SANGRADO_GRID = "O Cheque da linha %s do grid já foi sangrado pelo movimento de caixa %s."
Public Const ERRO_MOVIMENTO_NAO_SANGRIA_CHEQUE = "O Movimento %s não é uma sangria de cheques."
Public Const ERRO_MOVIMENTO_NAO_SANGRIA_BOLETO_CC = "O Movimento %s não é uma sangria de boletos de cartão de crédito."
Public Const ERRO_MOVIMENTO_NAO_SANGRIA_OUTROS = "O Movimento %s não é uma sangria de outros meios de pagamento."
Public Const ERRO_MOVIMENTO_NAO_SANGRIA_TICKET = "O Movimento %s não é uma sangria de ticket."
Public Const ERRO_ADMMEIOPAGTO_NAO_EXISTENTE1 = "O meio de pagamento %s não existe."
Public Const ERRO_VALORSANGRIA_NAO_DISPONIVEL_GRID = "O valor determinado para a Sangria %s é maior do que o valor total em caixa %s na linha %s."
Public Const AVISO_EXCLUSAO_ORCAMENTO_ECF = "O orçamento vai ser desprezado. Deseja continuar?"
Public Const ERRO_NAO_HA_DADOS_TRANSMITIR = "Não há dados que precisem ser transmitidos"
Public Const ERRO_ARQUIVO_ORCAMENTO_INVALIDO = "Este arquivo de orçamento não contém os dados que deveria no primeiro registro. Arquivo = %s."
Public Const ERRO_ORCAMENTO_NAO_CADASTRADO1 = "O Orçamento %s não foi encontrado."
Public Const ERRO_CAIXA_SO_ORCAMENTO = "Esta função não está disponível pois este terminal está configurado para só trabalhar com orçamentos."
Public Const ERRO_ORCAMENTO_SEM_ITEM = "O orçamento deve ter pelo menos um item."
Public Const ERRO_ARQUIVO_LOCADO = "O arquivo %s está sendo acessado por um outro programa e não pode ser acessado no momento."
Public Const ERRO_GETWINDOWSDIRECTORY = "Ocorreu um erro na execução da função GetWindowsDirectory."
Public Const ERRO_CAIXACONFIGINI_NAO_ENCONTRADO = "O arquivo caixaconfig.ini não foi encontrado. Arquivo = %s."
Public Const ERRO_LIMITE_DESCONTO_ULTRAPASSADO = "O Limite de Desconto do Produto %s que é %s foi ultrapassado. Percentual Desconto Total = %s."
Public Const ERRO_CLIENTE_NAO_CADASTRADO3 = "O cliente em questão não está cadastrado. Cliente = %s."
Public Const ERRO_TAMANHO_CGC_CPF1 = "O tamanho do campo tem que ser 11 caracteres para CPF ou 14 para CGC."
Public Const ERRO_CPFCNPJ_NAO_CADASTRADO = "O cliente em questão não está cadastrado. CPF/CNPJ = %s."
Public Const ERRO_TIPOCF_NAO_ESCOLHIDO = "É necessário selecionar uma das opções: Cupom Fiscal ou Orçamento."
Public Const ERRO_VALORMINIMO_CONDPAGTO = "O Valor mínimo da condição de pagamento escolhida não foi atingido. Valor Mínimo = %s, Valor = %s."
Public Const ERRO_ARQUIVO_NAO_ENCONTRADO2 = "Não conseguiu encontrar o arquivo %s."
Public Const ERRO_CONDPAGTO_NAO_CADASTRADO = "A condição de pagamento não está cadastrada. Rede = %s, Administradora = %s, Num.Parcelas = %s, Meio Pagamento = %s, Tipo Meio Pagto = %s ."
Public Const ERRO_VALIDAR_CHEQUE_SO_FUNCIONA_VENDA = "A validação de cheques só funciona para vendas."
Public Const ERRO_VALOR_TEF = "O valor do TEF não pode superar o valor a pagar."
Public Const ERRO_BOTAO_AG_NAO_TEF = "Ao acionar o botão Abrir Gaveta, não pode ter valor no campo TEF."
Public Const ERRO_BOTAO_TEF_SEM_FALTA = "Ao acionar o botão TEF e ter preenchido o campo TEF, não pode ter valor no campo Falta."
Public Const ERRO_VALOR_TEF_SUPERIOR_FALTA = "O valor não pode superar o valor que falta pagar."
Public Const ERRO_SITEF_NAO_ATIVO = "Erro de Comunicação - Cliente SiTef. Favor ativá-lo."
Public Const ERRO_SITEF_TRN_NAO_EFETUADA = "Transação não efetuada. Favor reter o Cupom."
Public Const ERRO_ARQUIVO_CONFIGURACAO_ALTERADO = "O arquivo de configuração (orcamento.ini) foi alterado. Entre em contato com o fabricante."
Public Const ERRO_USAIMPRESSORAFISCAL_INVALIDO = "Orcamento.ini esta setado para o uso de impressora fiscal mas o modelo de impressora escolhido não é fiscal."
Public Const ERRO_USAIMPRESSORANAOFISCAL_INVALIDO = "Orcamento.ini esta setado para o uso de impressora não fiscal mas o modelo de impressora escolhido é fiscal."
Public Const AVISO_NAO_E_POSSIVEL_EXLUIR_ORCAMENTO = "Não é possível excluir/cancelar orçamentos."
Public Const ERRO_ORCAMENTO_COM_CUPOM = "O orçcamento em questão já existe e já teve seu cupom fiscal correspondente emitido. Orcamento = %s."
Public Const ERRO_TAMANHO_DO_PAPEL_MENOR_QUE_MINIMO = "O tamanho do papel da impressora é menor de 210 mm x 148 mm ou 240 mm x 140 mm."
Public Const ERRO_ARQUIVO_DAV_INVALIDO = "Este arquivo de DAV não contém os dados que deveria no primeiro registro. Arquivo = %s."
Public Const ERRO_DATAS_NAO_PREENCHIDAS = "As datas precisam estar preenchidas."
Public Const ERRO_GERACAO_ASSINATURA = "Ocorreu um erro na geração da assinatura RSA."
Public Const ERRO_CODIGO_CODIGOCFE_INVALIDO = "O meio de pagamento está inválido."
Public Const ERRO_FUNCAO_NAO_DISPONIVEL_NAOFISCAL = "A função em questão não está disponível por se tratar de impressora não fiscal."
Public Const ERRO_COO_NAO_PREENCHIDO = "A faixa de COO não está preenchido, os intervalos são campos obrigatórios."
Public Const ERRO_COO_INICIAL_MAIOR_FINAL = "O COO Inicial é maior que o final."
Public Const ERRO_ARQUIVO_ECFCORPORATOR_ALTERADO = "O arquivo ecfcorporator.crp foi alterado. Entre em contato com o fabricante."
Public Const ERRO_DATAS_DIFEREM = "A data do ecf = %s está diferente da data do sistema = %s."
Public Const ERRO_HORAS_DIFEREM = "A diferença da data do sistema para a data do ecf não pode ultrapassar 60 minutos. hora do ecf = %s,  hora do sistema = %s."
Public Const ERRO_ABERTURA_COMANDO = "Ocorreu um erro ao tentar abrir os comandos de banco de dados."
Public Const ERRO_LEITURA_ECFCONFIG = "Ocorreu um erro ao tentar ler um registro da tabela Ecfconfig.  Codigo = %s."
Public Const ERRO_ALTERACAO_ECFCONFIG = "Ocorreu um erro ao tentar alterar um registro da tabela Ecfconfig.  Codigo = %s."
Public Const ERRO_LEITURA_MOVIMENTOCAIXA = "Ocorreu um erro ao tentar ler um registro da tabela MovimentoCaixa."
Public Const ERRO_INCLUSAO_MOVIMENTOCAIXA = "Ocorreu um erro ao tentar incluir um registro na tabela MovimentoCaixa.  NumIntDoc = %s."
Public Const ERRO_COMMIT = "Ocorreu um erro ao confirmar uma transacao (commit). "
Public Const ERRO_LEITURA_ARQUIVOSEQ = "Ocorreu um erro ao tentar ler um registro da tabela ArquivoSeq."
Public Const ERRO_INCLUSAO_ARQUIVOSEQ = "Ocorreu um erro ao tentar incluir um registro na tabela ArquivoSeq."
Public Const ERRO_ARQUIVOSEQ_NAO_CADASTRADO = "o Sequencial %s da tabela ArquivoSeq não esta cadastrado."
Public Const ERRO_INCLUSAO_ORCAMENTO = "Ocorreu um erro ao tentar incluir um registro na tabela Orcamento."
Public Const ERRO_EXCLUSAO_ORCAMENTO = "Ocorreu um erro ao tentar excluir um registro da tabela Orcamento."
Public Const ERRO_LEITURA_ORCAMENTO = "Ocorreu um erro ao tentar ler um registro da tabela Orcamento."
Public Const ERRO_INCLUSAO_ORCAMENTOBAIXADO = "Ocorreu um erro ao tentar incluir um registro na tabela OrcamentoBaixado."
Public Const ERRO_LEITURA_CAIXAATIVO = "Ocorreu um erro ao tentar ler um registro da tabela CaixaAtivo."
Public Const ERRO_INCLUSAO_CAIXAATIVO = "Ocorreu um erro ao tentar incluir um registro na tabela CaixaAtivo."
Public Const ERRO_EXCLUSAO_CAIXAATIVO = "Ocorreu um erro ao tentar excluir um registro da tabela CaixaAtivo."
Public Const ERRO_ORCAMENTO_JA_CADASTRADO = "O orçamento %s já está cadastrado. Não é possível alterá-lo nem exclui-lo."
Public Const ERRO_IMPRESSAO_NAO_PERMITIDA = "Não é possível imprimir um orçamento que não seja do tipo DAV."
Public Const ERRO_ECF_LEITURA_TABELA_CONFIG = "Ocorreu um erro ao tentar ler um registro da tabela %s.  Codigo = %s."
Public Const ERRO_ECF_LOCK_TABELA_CONFIG = "Ocorreu um erro ao tentar fazer um lock em um registro da tabela %s.  Codigo = %s."
Public Const ERRO_ECF_LEITURA_TABELA = "Ocorreu um erro ao tentar ler um registro da tabela %s."
Public Const ERRO_ALTERACAO_TABELA_CONFIG = "Ocorreu um erro ao tentar alterar um registro da tabela %s.  Codigo = %s."
Public Const ERRO_QTDE_INVALIDA_NUMINT = "Ocorreu um erro. A quantidade é negativa."
'Public Const ERRO_ARQUIVOSEQ_NAO_CADASTRADO = "O registro da tabela ArquivoSeq não está cadastrado. Sequencial = %s."
Public Const ERRO_LEITURA_ORCAMENTOCONFIG = "Ocorreu um erro ao tentar ler um registro da tabela OrcamentoConfig.  Codigo = %s."
Public Const ERRO_ORCAMENTOCONFIG_NAO_CADASTRADO = "O registro da tabela OrcamentoConfig não está cadastrado. Codigo = %s."
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

Public Const ERRO_NFCE_NAO_AUTORIZADA = "Erro na tentativa de gravar nfce não autorizada."
Public Const ERRO_PRODUTO_NAO_SELECIONADO = "Nenhum produto foi selecionado."
Public Const ERRO_PRODUTO_JA_SELECIONADO_LOJA = "Este produto já foi selecionado."
Public Const ERRO_ORCAMENTO_EAD = "O checksum da tabela orçamento não está correto. Codigo = %s."
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
Public Const ERRO_E1_NAO_CADASTRADO = "O estoque não esta cadastrado na tabela E1."
Public Const ERRO_E2_NAO_CADASTRADO = "O estoque não está cadastrado na tabela E2. Data = %s, Produto = %s"
Public Const ERRO_E3_NAO_CADASTRADO = "O estoque não está cadastrado na tabela E3."
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
Public Const ERRO_P1_NAO_CADASTRADO = "Não há registro cadastrado na tabela P1."
Public Const ERRO_LEITURA_MEIOPAGAMENTO = "Ocorreu um erro ao tentar ler um registro da tabela MeioPagamento."
Public Const ERRO_INCLUSAO_MEIOPAGAMENTO = "Ocorreu um erro ao tentar inserir um registro na tabela MeioPagamento."
Public Const ERRO_ALTERACAO_MEIOPAGAMENTO = "Ocorreu um erro ao tentar alterar um registro da tabela MeioPagamento."
Public Const ERRO_NAO_EXISTE_R01_PERIODO = "Não existe registro R01 no periodo em questão."
Public Const ERRO_LEITURA_ARQUIVOECF = "Ocorreu um erro ao tentar ler um registro da tabela ArquivoECF."
Public Const ERRO_R02_NAO_CADASTRADO = "Não há registro cadastrado na tabela R02.  GT Arquivo = %s, GT ECF = %s."
Public Const ERRO_GT_NAO_CONFERE_BLOQUEIO = "O GT não confere com o o que esta no arquivo. GT Arquivo = %s, GT ECF = %s."
Public Const ERRO_NUMFAB_NAO_CADASTRADO_BLOQUEIO = "O numero de serie %s do ecf nao esta cadastrado no arquivo."
Public Const ERRO_LEITURA_CONFIGURACAOECF = "Ocorreu um erro ao tentar ler um registro da tabela ConfiguracaoECF."
Public Const ERRO_CONFIGURACAOECF_NAO_CADASTRADO = "Nao ha registro na tabela ConfiguracaoECF."
Public Const ERRO_ORCAMENTO_BAIXADO = "O Orçamento %s já está baixado."
Public Const ERRO_DAV_NAO_ALTERADO_DEPOIS_DE_IMPRESSO = "O DAV não pode ser alterado depois de impresso."
Public Const ERRO_DAV_NAO_PODE_SER_REIMPRESSO = "O DAV não pode ser reimpresso."
Public Const ERRO_ITEM_NAO_ENCONTRADO_CANCELAR = "Não foi encontrado nenhum item para cancelar."
Public Const ERRO_PREVENDA_DAV_SIMULTANEOS = "As opções de DAV e Pré Venda estao acionadas simultaneamente."
Public Const ERRO_PREVENDA_SEM_IMPRESSORAFISCAL = "A opção de Pré Venda está ligada mas o uso de impressora fiscal está desligado."
Public Const ERRO_R07_NAO_CADASTRADO = "Não há registro cadastrado na tabela R07 para os dados em questão."
Public Const ERRO_ARQ_CRIPTOGRAFADO_INEXISTENTE = "O arquivo criptografado %s não foi encontrado."
Public Const ERRO_PREVENDA_CUPOM_SIMULTANEOS = "As opções de Cupom Fiscal e Pré Venda estao acionadas simultaneamente."
Public Const AVISO_CAIXA_SO_ORCAMENTO = "Este é um caixa que só faz orçamento. Não é possível executar esta função."
Public Const ERRO_ALTERACAO_CONFIGURACAOECF = "Ocorreu um erro ao tentar alterar um registro na tabela ConfiguracaoECF."
Public Const AVISO_EXCLUSAO_NFD2 = "Confirma a exclusão da nota fiscal %s.?"
Public Const AVISO_ALTERACAO_NFD2 = "A nota fiscal %s já existe no Sistema. Deseja Realmente efetuar a alteração ?"
Public Const ERRO_EXCLUSAO_NFD2 = "Ocorreu um erro na exclusão na nota fiscal."
Public Const ERRO_GRAVACAO_NFD2 = "Ocorreu um erro na gravação na nota fiscal."
Public Const ERRO_LEITURA_NFD2 = "Ocorreu um erro na leitura na nota fiscal."
Public Const ERRO_NFD2_NAO_LOCALIZADA = "A nota fiscal série %s número %s emitida em %s não foi localizada."
Public Const ERRO_SERIE_NAO_PREENCHIDA = "O campo Série deve estar preenchido."
Public Const ERRO_NENHUM_ITEM_GRID = "Nenhum item foi incluído"
Public Const ERRO_DISCRIMINACAO_NAO_PREENCHIDA = "A discriminação da mercadoria tem que ser preenchida"
Public Const ERRO_NFD2_JA_EXISTENTE = "Esta Nota Fiscal já está cadastrada no Sistema."
Public Const ERRO_ORCAMENTO_NAO_PERMITE_INCLUSAO_ITENS = "Orçamento ao ser transformado para cupom não admite adição de itens."
Public Const ERRO_ITEM_ORCAMENTO_CUPOM_NAO_PODE_CANCELAR = "Orçamento ao ser transformado para cupom não admite cancelamento de itens."
Public Const ERRO_NAO_PERMITIDO_IMPRIMIR_DAV_NAO_FISCAL = "Não é permitido imprimir DAV em impressora não fiscal."
Public Const ERRO_MD5_ALTERADO = "O MD5 do arquivo criptografado foi alterado."
Public Const ERRO_HOUVE_PERDA_DADOS_ARQ_CRIPTO = "Houve perda de dados no arquivo criptografado e o mesmmo não pode ser recomposto."
Public Const ERRO_NFD2_DESABILITADA = "Só pode registrar nota manual após redução Z ou com o ECF com problema."
Public Const ERRO_LEITURA_CODNACID = "Ocorreu um erro ao tentar ler um registro da tabela CodNacId."
Public Const ERRO_CODNACID_NAO_CADASTRADO = "O CodNacId da Marca = %s , Modelo = %s, VersaoSB = %s não está cadastrado."
Public Const ERRO_DIRETORIO_INVALIDO = "O diretório escolhido não existe. %s"
Public Const ERRO_VERSAO_PGM_INCOMPATIVEL_PGM = "Existe uma incompatibilidade entre a versão do programa instalada e o banco de dados. Pgm: %s Pgm BD: %s"
Public Const ERRO_VERSAO_PGM_INCOMPATIVEL_BD_ECF = "Existe uma incompatibilidade entre a versão do programa instalada e o banco de dados. BD ECF Pgm: %s BD ECF: %s"
Public Const ERRO_VERSAO_PGM_INCOMPATIVEL_BD_ORC = "Existe uma incompatibilidade entre a versão do programa instalada e o banco de dados. BD Orc Pgm: %s BD Orc: %s"
Public Const ERRO_NAO_PERMITIDO_CANCELAR_VARIOS_VINC = "Não é permitido cancelar cupom com mais de um cupom vinculado."
Public Const ERRO_LEITURA_CONFIGURACAOSAT = "Ocorreu um erro ao tentar ler um registro da tabela ConfiguracaoSAT."
Public Const ERRO_ALTERACAO_CONFIGURACAOSAT = "Ocorreu um erro ao tentar alterar um registro na tabela ConfiguracaoSAT."
Public Const ERRO_LEITURA_CONFIGURACAONFE = "Ocorreu um erro ao tentar ler um registro da tabela ConfiguracaoNFe."
Public Const ERRO_ALTERACAO_CONFIGURACAONFE = "Ocorreu um erro ao tentar alterar um registro na tabela ConfiguracaoNFe."
Public Const ERRO_INCLUSAO_NFCEINFO = "Ocorreu um erro ao tentar inserir um registro na tabela NFCeInfo."
Public Const ERRO_CONFIGURACAOSAT_NAO_CADASTRADO = "Não há ConfiguracaoSAT cadastrado."
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
Public Const ERRO_ARQ_CRIPTOGRAFADO_VAZIO = "O arquivo criptografado %s está vazio."
Public Const ERRO_INCLUSAO_CLIENTES = "Ocorreu um erro ao tentar gravar um registro na tabela Clientes."
Public Const ERRO_EXCLUSAO_CLIENTES1 = "Ocorreu um erro ao tentar excluir um registro na tabela Clientes."
Public Const ERRO_LEITURA_CLIENTE_NOMEREDUZIDO = "Ocorreu um erro ao tentar ler um registro da tabela Clientes.  NomeReduzido = %s."
Public Const ERRO_LEITURA_CLIENTES_0 = "Ocorreu um erro ao tentar ler um registro da tabela Clientes."
Public Const ERRO_LEITURA_CLIENTE_CODIGO = "Ocorreu um erro ao tentar ler um registro da tabela Clientes.  Codigo = %s."
Public Const ERRO_INCLUSAO_VENDEDORES = "Ocorreu um erro ao tentar gravar um registro na tabela Vendedores."
Public Const ERRO_LEITURA_VENDEDOR_CODIGO = "Ocorreu um erro ao tentar ler um registro da tabela Vendedores.  Codigo = %s."
Public Const ERRO_EXCLUSAO_VENDEDORES = "Ocorreu um erro ao tentar excluir um registro da tabela Vendedores."
Public Const ERRO_TROCO_DINHEIRO_NEGATIVO = "O Valor do troco em dinheiro não pode ser negativo."
Public Const ERRO_TROCO_CONTRAVALE_NEGATIVO = "O Valor do troco em contra vale não pode ser negativo."
Public Const ERRO_TROCO_TICKET_NEGATIVO = "O Valor do troco em ticket não pode ser negativo."

Public Const ERRO_LEITURA_BACKUPCONFIG = "Ocorreu um erro na leitura da tabela de configuração de backup"
Public Const ERRO_UPDATE_BACKUPCONFIG = "Ocorreu um erro na alteração de um registro da tabela de configuração de backup"
Public Const ERRO_INSERCAO_BACKUPCONFIG = "Ocorreu um erro na inclusão de um registro da tabela de configuração de backup"
Public Const ERRO_BACKUP_TOKEN_LIBERAR = "Ocorreu um erro ao tentar liberar o token para backup"
Public Const ERRO_BACKUPCONFIG_NAO_CADASTRADO = "A configuração de backup com código %s não foi configurada."
Public Const ERRO_LEITURA_BDSINFO = "Ocorreu um erro na leitura da tabela dos BDs para backup (BDsInfo)"
Public Const ERRO_BDSINFO_NAO_CADASTRADO = "Não foi encontrado nenhum BD para backup na tabela BDsInfo."
Public Const ERRO_BACKUP_BDSINFO = "Não foi possível realizar o backuo. %s"
Public Const AVISO_ERRO_AO_COMPACTAR_O_BACKUP = "Ocorreu um erro na tentativa de compactar o backup"
Public Const AVISO_ERRO_AO_FAZER_O_UPLOAD_DO_BACKUP = "Ocorreu um erro na tentativa fazer o upaload do backup"
Public Const AVISO_ERRO_AO_APAGAR_BACKUPS_ANTIGOS_FTP = "Ocorreu um erro na tentativa de apagar um backup mais antigo"
Public Const ERRO_LEITURA_BACKUPLOG = "Ocorreu um erro na leitura da tabela de Log de backup"
Public Const ERRO_UPDATE_BACKUPLOG = "Ocorreu um erro na alteração de um registro da tabela de Log de backup"

Public Const ERRO_NENHUMARQ_BKP_ENCONTRADO_FTP = "Nenhum arquivo para download foi encontrado no diretório FTP"

Public Const ERRO_LEITURA_SALDOEMDINHEIRO = "Ocorreu um erro na leitura da tabela SaldoEmDinheiro"
Public Const ERRO_INSERT_SALDOEMDINHEIRO = "Ocorreu um erro na inclusão de um registro na tabela SaldoEmDinheiro"

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
    
    'para depurar o batchest2 como uma dll o trecho abaixo até os asteriscos deve estar comentado
    'se há alguma transacao aberta vou forcar o rollback
    If GL_lTransacao <> 0 And iTipoCliente <> AD_SIST_BATCH Then
        Call Y.Transacao_Rollback
    End If
'****

    If Z.glTransacaoPAFECF <> 0 Then
    
        lTrans = Z.glTransacaoPAFECF
        
        'Desfaz Transação
        Call Transacao_RollbackExt(lTrans)
        
        Z.glTransacaoPAFECF = 0
    
    End If
    
    If Z.glTransacaoOrcPAFECF <> 0 Then
    
        lTrans = Z.glTransacaoOrcPAFECF
    
        'Desfaz Transação
        Call Transacao_RollbackExt(lTrans)
        
        Z.glTransacaoOrcPAFECF = 0
    
    End If
    
    GL_lUltimoErro = lCodigo

    Rotina_ErroECF = 0
    sErro = String(1024, 0)

    'para depurar o batchest como uma dll o comando abaixo deve estar descomentado
'    iTipoCliente = AD_SIST_BATCH
'****
    'Se o erro trazido é ERRO_FORNECIDO_PELO_VB_1
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
    
'    Call Sistema_RegistrarOcorrencia(GL_lSistema, "Empresa: " & gsNomeEmpresa & " Filial: " & gsNomeFilialEmpresa & " - " & Format(Now(), "General Date") & " - " & "Usuário: " & gsUsuario)
'    Call Sistema_RegistrarOcorrencia(GL_lSistema, "   ERRO: " & sTipoErro & " Local: " & CStr(lCodigo) & " Descrição: " & sErro)
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
''   Rotina_Aviso = MsgBox("Confirma a operação?", MsgBoxTipo, "Rotina temporaria de aviso")
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
    
    'Passa o Erro e o Tipo dos Botões
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
            Call MsgBox("Erro na exibição de uma pergunta (ou aviso).", vbOKOnly, "SGE - Forprint")
        
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

