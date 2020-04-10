VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalEstUniao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'VARIÁVEIS USADAS POR TODAS NFS
'''iFornecedorAlterado
'''iNaturezaOpAlterada
'''iSerieAlterada
'''iFrameAtual
'''iEmitenteNF
'''iValorDescontoAlterado
'''iValorFreteAlterado
'''iValorSeguroAlterado
'''iValorDespesasAlterado
'''iValorIRRFAlterado
'''iAlterado
'''giPosCargaOk

Private mvariClienteAlterado As Integer         'NFED, NFER,     ,
Private mvardtDataEntrada As Date               '    , NFER, NFFE, NFE
Private mvariFornecedorBenefAlterado As Integer '    ,     , NFFE, NFE

Public Property Let iFornecedorBenefAlterado(ByVal vData As Integer)
    mvariFornecedorBenefAlterado = vData
End Property

Public Property Get iFornecedorBenefAlterado() As Integer
    iFornecedorBenefAlterado = mvariFornecedorBenefAlterado
End Property

Public Property Let dtDataEntrada(ByVal vData As Date)
    mvardtDataEntrada = vData
End Property

Public Property Get dtDataEntrada() As Date
    dtDataEntrada = mvardtDataEntrada
End Property

Public Property Let iClienteAlterado(ByVal vData As Integer)
    mvariClienteAlterado = vData
End Property

Public Property Get iClienteAlterado() As Integer
    iClienteAlterado = mvariClienteAlterado
End Property

'@@@@@ para ser customizada
Function NFEnt_Traz_NFiscal_Tela(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long
    NFEnt_Traz_NFiscal_Tela = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFEnt_Move_Tela_Memoria(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long
    NFEnt_Move_Tela_Memoria = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFEnt_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long
    NFEnt_Form_Unload = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFEnt_Form_Load(ByVal objCT As Object) As Long
    NFEnt_Form_Load = SUCESSO
End Function

'################################
'Inserido por Wagner
Function NFEntC_Traz_NFiscal_Tela(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long
    NFEntC_Traz_NFiscal_Tela = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFEntC_Move_Tela_Memoria(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long
    NFEntC_Move_Tela_Memoria = SUCESSO
End Function

Function NFEntC_Form_Load(ByVal objCT As Object) As Long
    NFEntC_Form_Load = SUCESSO
End Function

Function NFEntC_Limpa_Tela(ByVal objCT As Object) As Long
    NFEntC_Limpa_Tela = SUCESSO
End Function
'################################

