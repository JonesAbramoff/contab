Attribute VB_Name = "ECFElgin"
'**********************************************************
' Nome do modulo: << MdlMain.Bas >>
' Data Inicio: 28/06/00
' Cristiane Silva
'
' Cinq Coerente Informatica
'**********************************************************
Option Explicit
'********************
'* Constantes de erro
'********************

Global Const CIF_OK = 0                   ' Sucesso
Global Const CIF_PPAPEL = 1               ' Sucesso, detectado pouco papel
Global Const CIF_CANCCUP = 2              ' Sucesso, cancelando cupom
Global Const CIF_CUPNF = 3                ' Sucesso, abrindo cupom rel gerencial
Global Const CIF_ERR_FIMPAPEL = -83       ' Ocorreu Fim de Papel
Global Const CIF_ERR = -84                ' Falha Geral
Global Const CIF_EMEXECUCAO = -85         ' Comando nao recebido pelo ECF
Global Const CIF_ERR_CONFIG = -86         ' Erro no Cif.ini
Global Const CIF_ERR_SERIAL = -87         ' Falha na abertura da serial
Global Const CIF_ERR_SYS = -88            ' Erro na alocacao de recursos do windows.
Global Const CIF_ERR_ANSWER = -89         ' Retorno nao identificado
Global Const CIF_ERR_READSER = -90        ' Erro de TimeOut na Read Serial

Global Const CIF_ERR_TEMP = -91           ' Temperatura Alta
Global Const CIF_ERR_PPAPEL = -92         ' Detectado pouco papel

Global Const CIF_IRRECUPERAVEL = -94      ' Erro Irrecuperavel
Global Const CIF_ERR_MECANICO = -95       ' Erro Mecanico
Global Const CIF_ERR_TABERTA = -96        ' Tampa Aberta
Global Const CIF_SEMRETORNO = -97         ' Sem Retorno
Global Const CIF_OVERFLOW = -98           ' Overflow
Global Const CIF_TIMEOUT = -99            ' TimeOut na execucao do comando


'*****************************
' Definir funcoes para ECF32M
'*****************************

Declare Sub Elgin_CloseCif Lib "ECF32M.DLL" Alias "CloseCif" ()
Declare Function Elgin_OpenCif Lib "ECF32M.DLL" Alias "OpenCif" () As Long


'**********************
' Funcoes da impressora
'**********************

Declare Function Elgin_ModoChequeValidacao Lib "ECF32M.DLL" Alias "ModoChequeValidacao" (ByVal Tipo As Byte, ByVal load As Byte) As Long
Declare Function Elgin_ImprimeCheque Lib "ECF32M.DLL" Alias "ImprimeCheque" (ByVal l1 As Byte, ByVal c1 As Byte, ByVal l2 As Byte, ByVal c2 As Byte, ByVal l3 As Byte, ByVal c3 As Byte, ByVal l4 As Byte, ByVal l5 As Byte, ByVal c5 As Byte, ByVal l6 As Byte, ByVal l7 As Byte, ByVal c8 As Byte, ByVal Valor As String, ByVal Favorecido As String, ByVal Cidade As String, ByVal SetAno As Byte, ByVal Data As String, ByVal Com1 As String, ByVal Com2 As String) As Long
Declare Function Elgin_ImprimeValidacao Lib "ECF32M.DLL" Alias "ImprimeValidacao" (ByVal leg As String, ByVal LinhaOp As String) As Long
Declare Function Elgin_CancelaChequeValidacao Lib "ECF32M.DLL" Alias "CancelaChequeValidacao" () As Long

'*****************************
' Funcoes de operacoes fiscais
'*****************************

Declare Function Elgin_TotalizarCupomParcial Lib "ECF32M.DLL" Alias "TotalizarCupomParcial" () As Long
Declare Function Elgin_TotalizarCupom Lib "ECF32M.DLL" Alias "TotalizarCupom" (ByVal oper As Byte, ByVal toper As Byte, ByVal Valor As String, ByVal legendaOp As String) As Long
Declare Function Elgin_TotCupomSemDescAcres Lib "ECF32M.DLL" Alias "TotCupomSemDescAcres" () As Long
Declare Function Elgin_TotCupomAcresValor Lib "ECF32M.DLL" Alias "TotCupomAcresValor" (ByVal Valor As String, ByVal legendaOp As String) As Long
Declare Function Elgin_TotCupomAcresPorcentagem Lib "ECF32M.DLL" Alias "TotCupomAcresPorcentagem" (ByVal porcentagem As String, ByVal legendaOp As String) As Long
Declare Function Elgin_TotCupomDescValor Lib "ECF32M.DLL" Alias "TotCupomDescValor" (ByVal Valor As String, ByVal legendaOp As String) As Long
Declare Function Elgin_TotCupomDescPorcentagem Lib "ECF32M.DLL" Alias "TotCupomDescPorcentagem" (ByVal porcentagem As String, ByVal legendaOp As String) As Long
Declare Function Elgin_Pagamento Lib "ECF32M.DLL" Alias "Pagamento" (ByVal reg As String, ByVal vpgto As String, ByVal subtr As Byte) As Long
Declare Function Elgin_PagamentoComTexto Lib "ECF32M.DLL" Alias "PagamentoComTexto" (ByVal reg As String, ByVal vpgto As String, ByVal parm As Byte, ByVal comentario As String) As Long
Declare Function Elgin_AbreCupomFiscal Lib "ECF32M.DLL" Alias "AbreCupomFiscal" () As Long
Declare Function Elgin_VendaItemStr Lib "ECF32M.DLL" Alias "VendaItemStr" (ByVal fmt As String, ByVal qtd As String, ByVal punit As String, ByVal trib As String, ByVal TDesc As String, ByVal Valor As String, ByVal unid As String, ByVal cod As String, ByVal ex As String, ByVal descr As String, ByVal legendaOp As String) As Long
Declare Function Elgin_CancelamentoItem Lib "ECF32M.DLL" Alias "CancelamentoItem" (ByVal numitem As String) As Long
Declare Function Elgin_DescontoItem Lib "ECF32M.DLL" Alias "DescontoItem" (ByVal toper As Byte, ByVal Valor As String, ByVal legop As String) As Long
Declare Function Elgin_DescontoItemPorcentagem Lib "ECF32M.DLL" Alias "DescontoItemPorcentagem" (ByVal porcentagem As String, ByVal legop As String) As Long
Declare Function Elgin_DescontoItemValor Lib "ECF32M.DLL" Alias "DescontoItemValor" (ByVal Valor As String, ByVal legop As String) As Long
Declare Function Elgin_FechaCupomFiscal Lib "ECF32M.DLL" Alias "FechaCupomFiscal" (ByVal tam_msg As String, ByVal Msg As String) As Long
Declare Function Elgin_CancelaCupomFiscal Lib "ECF32M.DLL" Alias "CancelaCupomFiscal" () As Long
Declare Function Elgin_LeituraX Lib "ECF32M.DLL" Alias "LeituraX" (ByVal relGer As Byte) As Long
Declare Function Elgin_ReducaoZ Lib "ECF32M.DLL" Alias "ReducaoZ" (ByVal relGer As Byte) As Long
Declare Function Elgin_LeituraXComRelGer Lib "ECF32M.DLL" Alias "LeituraXComRelGer" () As Long
Declare Function Elgin_ReducaoZComRelGer Lib "ECF32M.DLL" Alias "ReducaoZComRelGer" () As Long
Declare Function Elgin_LeMemFiscalData Lib "ECF32M.DLL" Alias "LeMemFiscalData" (ByVal datai As String, ByVal dataf As String, ByVal res As Byte) As Long
Declare Function Elgin_LeMemFiscalReducao Lib "ECF32M.DLL" Alias "LeMemFiscalReducao" (ByVal redi As String, ByVal redf As String, ByVal res As Byte) As Long
Declare Function Elgin_AbreCupomFiscalCPF_CNPJ Lib "ECF32M.DLL" Alias "AbreCupomFiscalCPF_CNPJ" (ByVal sCPF As String) As Long

'******************************************
' Funcoes de operacoes nao sujeitas ao ICMS
'******************************************

Declare Function Elgin_AbreCupomVinculado Lib "ECF32M.DLL" Alias "AbreCupomVinculado" () As Long
Declare Function Elgin_AbreCupomNaoVinculado Lib "ECF32M.DLL" Alias "AbreCupomNaoVinculado" () As Long
Declare Function Elgin_EncerraCupomNaoFiscal Lib "ECF32M.DLL" Alias "EncerraCupomNaoFiscal" () As Long
Declare Function Elgin_CancelaCupomNaoFiscal Lib "ECF32M.DLL" Alias "CancelaCupomNaoFiscal" () As Long
Declare Function Elgin_OperRegNaoVinculado Lib "ECF32M.DLL" Alias "OperRegNaoVinculado" (ByVal reg As String, ByVal Valor As String, ByVal oper As Byte, ByVal toper As Byte, ByVal valorop As String, ByVal legop As String) As Long
Declare Function Elgin_AbrirCupom Lib "ECF32M.DLL" Alias "AbrirCupom" (ByVal reg As String, ByVal Valor As String, ByVal oper As String, ByVal toper As String, ByVal valorop As String, ByVal legop As String, ByVal buffRet As String) As Long
Declare Function Elgin_ImprimeLinhaNaoFiscal Lib "ECF32M.DLL" Alias "ImprimeLinhaNaoFiscal" (ByVal par As Byte, ByVal str As String) As Long
Declare Function Elgin_ImprimeLinhaNaoFiscalTexto Lib "ECF32M.DLL" Alias "ImprimeLinhaNaoFiscalTexto" (ByVal par As String, ByVal str As String) As Long
Declare Function Elgin_ProgramaLegenda Lib "ECF32M.DLL" Alias "ProgramaLegenda" (ByVal reg As String, ByVal leg As String) As Long
'Declare Function Elgin_OpRegNaoVinculado Lib "ECF32M.DLL" Alias "OperRegNaoVinculado" (ByVal reg As String, ByVal Valor As String, ByVal oper As String, ByVal toper As String, ByVal valorop As String, ByVal legop As String) As Long

'*****************
' Funcoes diversas
'*****************

Declare Function Elgin_AcionarGaveta Lib "ECF32M.DLL" Alias "AcionarGaveta" () As Long
Declare Function Elgin_ProgramaHorarioVeraoStr Lib "ECF32M.DLL" Alias "ProgramaHorarioVeraoStr" (ByVal hv As String) As Long
Declare Function Elgin_ImprimeTotalizadores Lib "ECF32M.DLL" Alias "ImprimeTotalizadores" (ByVal reg As String) As Long
Declare Function Elgin_TransTabAliquotas Lib "ECF32M.DLL" Alias "TransTabAliquotas" () As Long
Declare Function Elgin_TransTotCont Lib "ECF32M.DLL" Alias "TransTotCont" () As Long
Declare Function Elgin_TransStatus Lib "ECF32M.DLL" Alias "TransStatus" (ByVal BitTest As Long, ByVal BufStat As String) As Long
Declare Function Elgin_TransDataHora Lib "ECF32M.DLL" Alias "TransDataHora" () As Long
Declare Function Elgin_EcfPar Lib "ECF32M.DLL" Alias "EcfPar" (ByVal par As String) As Long
Declare Function Elgin_ProgLinhaAdicional Lib "ECF32M.DLL" Alias "ProgLinhaAdicional" (ByVal reg As String) As Long
Declare Function Elgin_AjusteHora Lib "ECF32M.DLL" Alias "AjusteHora" (ByVal dir As Byte, ByVal Hora As String) As Long
Declare Function Elgin_EcfID Lib "ECF32M.DLL" Alias "EcfID" () As Long
Declare Function Elgin_EsperaResposta Lib "ECF32M.DLL" Alias "EsperaResposta" (ByVal buf_ret As String) As Long

'*********************************
' Funcoes de controle de impressao
'*********************************

Declare Function Elgin_ImprimeNaoFiscal Lib "ECF32M.DLL" Alias "ImprimeNaoFiscal" (ByVal NroImp As Long, ByVal Buf_Imp As String) As Long
Declare Function Elgin_SELECIONAATRIBUTO Lib "ECF32M.DLL" Alias "SELECIONAATRIBUTO" (ByVal Modo As String) As Long
Declare Function Elgin_MODOSUBLINHADO Lib "ECF32M.DLL" Alias "MODOSUBLINHADO" (ByVal Modo As String) As Long
Declare Function Elgin_HOME Lib "ECF32M.DLL" Alias "HOME" () As Long

'*******************************
' Funcoes de intervencao tecnica
'*******************************

Declare Function Elgin_ProgRelogio Lib "ECF32M.DLL" Alias "ProgRelogio" (ByVal Hora As String, ByVal Data As String) As Long
Declare Function Elgin_GravaDados Lib "ECF32M.DLL" Alias "GravaDados" (ByVal CGC As String, ByVal IE As String, ByVal ccm As String) As Long
Declare Function Elgin_RecompoeDadosNOVRAM Lib "ECF32M.DLL" Alias "RecompoeDadosNOVRAM" () As Long
Declare Function Elgin_ProgNumSerie Lib "ECF32M.DLL" Alias "ProgNumSerie" (ByVal numserie As String, ByVal modelo As String) As Long
Declare Function Elgin_ProgAliquotas Lib "ECF32M.DLL" Alias "ProgAliquotas" (ByVal tot As String, ByVal aliq As String) As Long
Declare Function Elgin_ProgSimbolo Lib "ECF32M.DLL" Alias "ProgSimbolo" (ByVal s1 As Byte, ByVal s2 As Byte, ByVal s3 As Byte, ByVal s4 As Byte, ByVal s5 As Byte, ByVal s6 As Byte, ByVal s7 As Byte, ByVal s8 As Byte, ByVal s9 As Byte, ByVal s10 As Byte, ByVal s11 As Byte) As Long
Declare Function Elgin_ProgRazaoSocial Lib "ECF32M.DLL" Alias "ProgRazaoSocial" (ByVal razsoc As String, ByVal numseq As String) As Long
Declare Function Elgin_Prog_Moeda Lib "ECF32M.DLL" Alias "Prog_Moeda" (ByVal sing As String, ByVal plur As String) As Long
Declare Function Elgin_ProgArredondamento Lib "ECF32M.DLL" Alias "ProgArredondamento" (ByVal par As Byte) As Long
Declare Function Elgin_ProgAliquotasICMS_ISS Lib "ECF32M.DLL" Alias "ProgAliquotasICMS_ISS" (ByVal tot As String, ByVal aliq As String, ByVal Tipo As Byte) As Long

Function Elgin_TraduzCodigoRetorno(ByVal intretorno As Integer) As String

    Dim strMsg As String
    
    Select Case intretorno
    
        '---------------------------
        ' Codigo de retorno dos comandos da impressora
        '
        Case -1
            strMsg = "Cabeçalho contém caracteres inválidos"
        Case -2
            strMsg = "Comando inexistente"
        Case -3
            strMsg = "Valor não numérico em campo numérico"
        Case -4
            strMsg = "Valor fora da faixa entre 20h e 7Fh"
        Case -5
            strMsg = "Campo deve iniciar com @, & ou %"
        Case -6
            strMsg = "Troco já realizado."
        Case -7
            strMsg = "O intervalo é inconsistente. No caso de datas, valores anteriores a " & _
                   "01/01/95 serão consideradas como ano 2000 a 2094"
        Case -9
            strMsg = "A string TOTAL não é aceita"
        Case -10
            strMsg = "A sintaxe do comando está errada"
        Case -11
            strMsg = "Excedeu número máximo de linhas permitidas pelo comando"
        Case -12
            strMsg = "O terminador enviado não está obedecendo o protocolo de comunicação"
        Case -13
            strMsg = "O checksum está incorreto"
        Case -15
            strMsg = "A situação tributária deve iniciar com T, F, I ou N"
        Case -16
            strMsg = "Data inválida"
        Case -17
            strMsg = "Hora inválida"
        Case -18
            strMsg = "Alíquota não programada ou fora do intervalo"
        Case -19
            strMsg = "O campo de sinal está incorreto"
        Case -20
            strMsg = "Comando só aceito em Intervenção Fiscal"
        Case -21
            strMsg = "Comando só aceito em Modo Normal"
        Case -22
            strMsg = "É necessário abrir o Cupom Fiscal"
        Case -23
            strMsg = "Comando não aceito durante Cupom Fiscal"
        Case -24
            strMsg = "É necessário abrir Cupom Não Fiscal"
        Case -25
            strMsg = "Comando não aceito durante Cupom Não Fiscal"
        Case -26
            strMsg = "O relógio já está em horário de verão"
        Case -27
            strMsg = "O relógio não está em horário de verão"
        Case -28
            strMsg = "Necessário realizar Redução Z"
        Case -29
            strMsg = "Fechamento do dia (Redução Z) já executado"
        Case -30
            strMsg = "Necessário programar legenda"
        Case -31
            strMsg = "Item inexistente ou já cancelado"
        Case -32
            strMsg = "O cupom anterior não pode ser cancelado"
        Case -33
            strMsg = "Detectado falta de papel. Verifique a impressora."
        Case -36
            strMsg = "Necessário programar os dados do estabelecimento"
        Case -37
            strMsg = "Necessário realizar Intervenção Fiscal."
        Case -38
        strMsg = "Memória Fiscal não permite mais realizar vendas. Apenas é possível realizar LeituraX " & _
               "ou Leitura da Memória Fiscal."
        Case -39
            strMsg = "Memória Fiscal não permite mais realizar vendas. Apenas é possível realizar LeituraX " & _
                   "ou Leitura da Memória Fiscal, deve haver algum problema na memória NOVRAM. Será " & _
                   "necessário realizar Intervenção Fiscal."
        Case -40
            strMsg = "Necessário programar a data do relógio"
        Case -41
            strMsg = "Número máximo de itens por cupom ultrapassado"
        Case -42
            strMsg = "Já foi realizado o Ajuste de Hora Diário"
        Case -43
            strMsg = "Comando válido ainda em execução -43"
        Case -44
            strMsg = "Está em estado de Impressão de Cheques"
        Case -45
            strMsg = "Não está em estado de Impressão de Cheques"
        Case -46
            strMsg = "Necessário inserir o cheque"
        Case -47
            strMsg = "Necessário inserir nova bobina"
        Case -48
            strMsg = "Necessário executar uma Leitura X"
        Case -49
            strMsg = "Detectado algum problema na impressora (Paper jam, sobretensão, etc)."
        Case -50
            strMsg = "Cupom já totalizado"
        Case -51
            strMsg = "Necessário totalizar cupom antes de fechar"
        Case -52
            strMsg = "Necessário finalizar Cupom com comando correto"
        Case -53
            strMsg = "Ocorreu erro de gravação na Memória Fiscal"
        Case -54
            strMsg = "Excedeu número máximo de estabelecimentos"
        Case -55
            strMsg = "Memória fiscal não inicializada"
        Case -56
            strMsg = "Ultrapassou valor do pagamento"
        Case -57
            strMsg = "Registrador não programado ou troco já realizado"
        Case -58
            strMsg = "Falta completar valor do pagamento"
        Case -59
            strMsg = "Campo somente de caracteres não numéricos"
        Case -60
            strMsg = "Excedeu campo máximo de caracteres"
        Case -61
            strMsg = "Troco não realizado"
        Case -62
            strMsg = "Comando desabilitado"
        Case CIF_OK
            strMsg = "Operação efetuada com sucesso"
        Case CIF_PPAPEL
            strMsg = "Sucesso, detectado pouco papel"
        Case CIF_CANCCUP
            strMsg = "Sucesso, cancelando cupom"
        Case CIF_CUPNF
            strMsg = "Sucesso, abrindo cupom rel gerencial"
        Case CIF_ERR_FIMPAPEL
            strMsg = "Ocorreu Fim de Papel"
        Case CIF_ERR
            strMsg = "Falha geral na execução da DLL"
        Case CIF_EMEXECUCAO
            strMsg = "Comando válido ainda em execução -85"
        Case CIF_ERR_CONFIG
            strMsg = "Erro no arquivo CIF.INI"
        Case CIF_ERR_SERIAL
            strMsg = "Erro na abertura da serial"
        Case CIF_ERR_SYS
            strMsg = "Falha na alocação de recursos do Windows"
        Case CIF_ERR_ANSWER
            strMsg = "Retorno nao reconhecido"
        Case CIF_ERR_READSER
            strMsg = "Falha na leitura da serial"
        Case CIF_ERR_TEMP
            strMsg = "Temperatura da cabeça de impressão alta"
        Case CIF_ERR_PPAPEL
            strMsg = "Pouco papel"
        Case CIF_IRRECUPERAVEL
            strMsg = "Erro irrecuperável"
        Case CIF_ERR_MECANICO
            strMsg = "Erro mecânico"
        Case CIF_ERR_TABERTA
            strMsg = "Tampa aberta"
        Case CIF_SEMRETORNO
            strMsg = "Operação sem retorno"
        Case CIF_OVERFLOW
            strMsg = "Buffer overflow. Tamanho da mensagem enviada pelo ECF é maior do que o buffer fornecido pela aplicação"
        Case CIF_TIMEOUT
            strMsg = "TimeOut na execucao do comando"
        Case Else
            strMsg = "Código de retorno inexistente" + intretorno
    End Select
    Elgin_TraduzCodigoRetorno = strMsg
    
End Function
