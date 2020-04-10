VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVariavelCalculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'armazena o valor de uma variavel identificada por uma string
'mantem o controle se o valor é valido e quem o "gerou"
'um valor "gerado" pelo usuario nao sera sobreposto por outro "automatico" (calculado pelo sistema)

Public sIdentificador As String
Public vValor As Variant
Public iStatus As Integer  'nao preenchido, preenchido automaticamente, preenchido manualmente (pelo usuario)

Function IGUAL(objVariavelCalculo As ClassVariavelCalculo)
'retorna SUCESSO se o conteudo for identico ao existente em "Me"
'??? falta implementar, ainda nao foi necessaria e talvez nem seja
End Function

Sub Copiar(objVariavelCalculo As ClassVariavelCalculo)
'??? verificar se é mesmo necessaria. P/objetos está criando uma referencia adicional e nao copiando o conteudo
'copiar o conteudo de objVariavelCalculo p/"Me"
    
    If (IsObject(objVariavelCalculo.vValor)) Then
        Set Me.vValor = objVariavelCalculo.vValor
    Else
        Me.vValor = objVariavelCalculo.vValor
    End If
    
    Me.iStatus = objVariavelCalculo.iStatus
    
End Sub

''''tipos de iStatus
'''VAR_CALC_NAO_PREENCHIDA         'nao preenchido
'''VAR_CALC_PREENCHIDA_MANUAL      'preenchido manualmente (pelo usuario)
'''VAR_CALC_PREENCHIDA_AUTO        'preenchido automaticamente pelo sistema
'''
''''valores definidos p/sIdentificador
'''
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''a nivel de doc
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''
'''TIPO_DOC_SIGLA
'''TIPO_DOC_CODIGO
'''NAT_OPERACAO
'''NAT_OPERACAO_EXTERNA (p/NFs de entrada na Empresa)
'''TIPO_TRIB
'''CLIENTE_CODIGO
'''CLIENTE_UF
'''FILIAL_CLIENTE_OBJ
'''FILIAL_CLIENTE_CODIGO
'''FILIAL_CLIENTE_INSC_EST
'''CLIENTE_ENDERECO_OBJ
'''FORN_UF
'''FORN_CODIGO
'''FILIAL_FORN_OBJ
'''FILIAL_FORN_CODIGO
'''FILIAL_FORN_INSC_EST
'''FORN_ENDERECO_OBJ
'''FRETE_VALOR
'''SEGURO_VALOR
'''DESP_ACESS_VALOR
'''NATOPPADRAO_OBJ

''''para ICMS
'''ICMS_BASE
'''ICMS_VALOR
'''ICMS_SUBST_BASE
'''ICMS_SUBST_VALOR
'''ICMS_VALOR_CREDITO

''''para IPI
'''IPI_BASE
'''IPI_VALOR
'''IPI_VALOR_CREDITO
'''
''''para ISS
'''ISS_INCLUSO
'''ISS_BASE
'''ISS_ALIQUOTA
'''ISS_VALOR
'''
''''para IR
'''IR_BASE
'''IR_ALIQUOTA
'''IR_VALOR
'''
''''para INSS
'''INSS_BASE
'''INSS_VALOR_DEDUCOES
'''INSS_VALOR
'''INSS_RETIDO
'''
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''a nivel de docitem
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''
'''ITEM_TIPO (ITEM_TIPO_NORMAL, ITEM_TIPO_FRETE, ITEM_TIPO_SEGURO, ITEM_TIPO_DESCONTO, ITEM_TIPO_OUTRAS_DESP)
'''NAT_OPERACAO
'''TIPO_TRIB
'''PRODUTO_CODIGO
'''PRODUTO_VALOR
'''PRODUTO_QTDE
'''PRODUTO_OBJ
'''
''''para ICMS
'''ICMS_TIPO
'''ICMS_BASE
'''ICMS_RED_BASE
'''ICMS_ALIQUOTA
'''ICMS_VALOR
'''ICMS_MARG_SUBST
'''ICMS_BASE_SUBST
'''ICMS_ALIQUOTA_SUBST
'''ICMS_VALOR_SUBST
'''ICMS_VALOR_CREDITO
'''
''''para IPI
'''IPI_TIPO
'''IPI_BASE
'''IPI_RED_BASE
'''IPI_ALIQUOTA
'''IPI_VALOR
'''IPI_VALOR_CREDITO
'''
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''
''''identificadores internos
'''
'''I_FILIAL_CLIENTE_COL_CATEG 'col de categorias associadas a uma filial cliente incluindo o valor em cada categoria
'''I_PRODUTO_COL_CATEG 'col de categorias associadas a um produto incluindo o valor em cada categoria
'''
''''para funcao ObterNaturezaOpPadrao
'''I_NOP_NAT_OP            'natureza de operacao padrao
'''I_NOP_TIPO_DOC          'sigla do tipo doc
'''I_NOP_UF
'''
''''para funcao ObterTribPadrao
'''I_OTP_TIPO_DOC_INFO
'''I_OTP_NAT_OP
'''I_OTP_CLIENTE_CODIGO
'''I_OTP_CLIENTE_FILIAL
'''
''''para funcao ObterColCategFilCli
'''I_CCC_CLIENTE_CODIGO
'''I_CCC_CLIENTE_FILIAL
'''
''''para funcao ObterColCategProd
'''I_CCP_PRODUTO

''''para funcao ObterObjFilialCliente
'''I_OFC_CLIENTE_CODIGO
'''I_OFC_FILIAL_CLIENTE_CODIGO

'''para funcao ObterUFDestino
'''I_OUD_CLIENTE_ENDERECO

'''para funcao ObterUFOrigem
'''I_OUO_ENDERECO

''''para funcao ObterObjFilialForn
'''I_OFF_FORN_CODIGO
'''I_OFF_FILIAL_FORN_CODIGO

'''para funcao ObterObjNatOpPadrao
'''I_NOPP_TIPOOPERTRIB
'''I_NOPP_TIPOATIVTRIB
