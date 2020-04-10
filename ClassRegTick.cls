VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRegTick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlTick_Indice As Long
Private mvarsTick_Sequencial As String
Private mvarsTick_PlacaCarreta As String
Private mvarsTick_PlacaVeiculo As String
Private mvarsTick_CodTransportadora As String
Private mvarsTick_CodEmissor As String
Private mvarsTick_CodItem As String
Private mvarsTick_SttFim As String
Private mvarsTick_Status As String
Private mvarsTick_RecExp As String
Private mvardTick_SeqRecExp As Double
Private mvarsTick_DescrStt As String
Private mvarsTick_DescrRecExp As String
Private mvardTick_PesoDosagem As Double
Private mvarsTick_RazSocTrans As String
Private mvarsTick_RazSocEmissor As String
Private mvarsTick_DescricaoItem As String
Private mvardTick_PesoInicial As Double
Private mvarsTick_OpPesoInicial As String
Private mvardtTick_DtHrPesoInicial As Date
Private mvarsTick_BalPesoInicial As String
Private mvardTick_PesoFinal As Double
Private mvarsTick_OpPesoFinal As String
Private mvardtTick_DtHrPesoFinal As Date
Private mvarsTick_BalPesoFinal As String
Private mvardTick_PesoLiquido As Double
Private mvardTick_FatCorrecao As Double
Private mvardTick_FatConversao As Double
Private mvardTick_LiquidoCorrigido As Double
Private mvarsTick_UnidadeAposConversao As String
Private mvarsTick_DescDocTot As String
Private mvardTick_PesoTotDoc As Double
Private mvardTick_DifOrigemRealPc As Double
Private mvardTick_DifOrigemRealKg As Double
Private mvardTick_PesoLiquido1 As Double
Private mvardTick_PesoLiquido2 As Double
Private mvardTick_PesoLiquido3 As Double
Private mvardTick_PesoLiquido4 As Double
Private mvardTick_PesoLiquido5 As Double
Private mvardTick_FatCorrecao1 As Double
Private mvardTick_FatCorrecao2 As Double
Private mvardTick_FatCorrecao3 As Double
Private mvardTick_FatCorrecao4 As Double
Private mvardTick_FatCorrecao5 As Double
Private mvardTick_FatCorrecao6 As Double
Private mvarsTick_Observacao1 As String
Private mvarsTick_Observacao2 As String
Private mvarsTick_SenhaPreEntrada As String
Private mvardtTick_DtHrPreEntrada As Date
Private mvarsTick_OpPreEntrada As String
Private mvarsTick_MemoObs As String
Private mvardTick_PesoCavInic As Double
Private mvarsTick_PlacaCavInic As String
Private mvardTick_PesoCavFinal As Double
Private mvarsTick_PlacaCavFinal As String
Private mvarlTick_QtdeTara1 As Long
Private mvardTick_PesoTara1 As Double
Private mvardTick_TotTara1 As Double
Private mvarlTick_QtdeTara2 As Long
Private mvardTick_PesoTara2 As Double
Private mvardTick_TotTara2 As Double
Private mvardTick_TotTaras As Double
Private mvardTick_PesoM3 As Double
Private mvariTick_HabCavalo As Integer
Private mvarlTick_Compartimentos As Long
Private mvardtTick_DtHrPosPesa As Date
Private mvarsTick_TmpVeicEmpr As String
Private mvarsTick_NumTransp As String
Private mvarsTick_CampoUsu1 As String
Private mvarsTick_CampoUsu2 As String
Private mvarsTick_CampoUsu3 As String
Private mvarsTick_CampoUsu4 As String
Private mvarsTick_CampoUsu5 As String
Private mvarsTick_CampoUsu6 As String
Private mvarsTick_CampoUsu7 As String
Private mvarsTick_CampoUsu8 As String
Private mvarsTick_CampoUsu9 As String
Private mvarsTick_CampoUsu10 As String
Private mvardTick_PesoLiqCorrUsu As Double
Private mvarlTick_NumOcupante As Long
Private mvarsCGC As String

Public Property Let sCGC(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCGC() As String
    sCGC = mvarsCGC
End Property

Public Property Let lTick_Indice(ByVal vData As Long)
    mvarlTick_Indice = vData
End Property

Public Property Get lTick_Indice() As Long
    lTick_Indice = mvarlTick_Indice
End Property

Public Property Let sTick_Sequencial(ByVal vData As String)
    mvarsTick_Sequencial = vData
End Property

Public Property Get sTick_Sequencial() As String
    sTick_Sequencial = mvarsTick_Sequencial
End Property

Public Property Let sTick_PlacaCarreta(ByVal vData As String)
    mvarsTick_PlacaCarreta = vData
End Property

Public Property Get sTick_PlacaCarreta() As String
    sTick_PlacaCarreta = mvarsTick_PlacaCarreta
End Property

Public Property Let sTick_PlacaVeiculo(ByVal vData As String)
    mvarsTick_PlacaVeiculo = vData
End Property

Public Property Get sTick_PlacaVeiculo() As String
    sTick_PlacaVeiculo = mvarsTick_PlacaVeiculo
End Property

Public Property Let sTick_CodTransportadora(ByVal vData As String)
    mvarsTick_CodTransportadora = vData
End Property

Public Property Get sTick_CodTransportadora() As String
    sTick_CodTransportadora = mvarsTick_CodTransportadora
End Property

Public Property Let sTick_CodEmissor(ByVal vData As String)
    mvarsTick_CodEmissor = vData
End Property

Public Property Get sTick_CodEmissor() As String
    sTick_CodEmissor = mvarsTick_CodEmissor
End Property

Public Property Let sTick_CodItem(ByVal vData As String)
    mvarsTick_CodItem = vData
End Property

Public Property Get sTick_CodItem() As String
    sTick_CodItem = mvarsTick_CodItem
End Property

Public Property Let sTick_SttFim(ByVal vData As String)
    mvarsTick_SttFim = vData
End Property

Public Property Get sTick_SttFim() As String
    sTick_SttFim = mvarsTick_SttFim
End Property

Public Property Let sTick_Status(ByVal vData As String)
    mvarsTick_Status = vData
End Property

Public Property Get sTick_Status() As String
    sTick_Status = mvarsTick_Status
End Property

Public Property Let sTick_RecExp(ByVal vData As String)
    mvarsTick_RecExp = vData
End Property

Public Property Get sTick_RecExp() As String
    sTick_RecExp = mvarsTick_RecExp
End Property

Public Property Let dTick_SeqRecExp(ByVal vData As Double)
    mvardTick_SeqRecExp = vData
End Property

Public Property Get dTick_SeqRecExp() As Double
    dTick_SeqRecExp = mvardTick_SeqRecExp
End Property

Public Property Let sTick_DescrStt(ByVal vData As String)
    mvarsTick_DescrStt = vData
End Property

Public Property Get sTick_DescrStt() As String
    sTick_DescrStt = mvarsTick_DescrStt
End Property

Public Property Let sTick_DescrRecExp(ByVal vData As String)
    mvarsTick_DescrRecExp = vData
End Property

Public Property Get sTick_DescrRecExp() As String
    sTick_DescrRecExp = mvarsTick_DescrRecExp
End Property

Public Property Let dTick_PesoDosagem(ByVal vData As Double)
    mvardTick_PesoDosagem = vData
End Property

Public Property Get dTick_PesoDosagem() As Double
    dTick_PesoDosagem = mvardTick_PesoDosagem
End Property

Public Property Let sTick_RazSocTrans(ByVal vData As String)
    mvarsTick_RazSocTrans = vData
End Property

Public Property Get sTick_RazSocTrans() As String
    sTick_RazSocTrans = mvarsTick_RazSocTrans
End Property

Public Property Let sTick_RazSocEmissor(ByVal vData As String)
    mvarsTick_RazSocEmissor = vData
End Property

Public Property Get sTick_RazSocEmissor() As String
    sTick_RazSocEmissor = mvarsTick_RazSocEmissor
End Property

Public Property Let sTick_DescricaoItem(ByVal vData As String)
    mvarsTick_DescricaoItem = vData
End Property

Public Property Get sTick_DescricaoItem() As String
    sTick_DescricaoItem = mvarsTick_DescricaoItem
End Property

Public Property Let dTick_PesoInicial(ByVal vData As Double)
    mvardTick_PesoInicial = vData
End Property

Public Property Get dTick_PesoInicial() As Double
    dTick_PesoInicial = mvardTick_PesoInicial
End Property

Public Property Let sTick_OpPesoInicial(ByVal vData As String)
    mvarsTick_OpPesoInicial = vData
End Property

Public Property Get sTick_OpPesoInicial() As String
    sTick_OpPesoInicial = mvarsTick_OpPesoInicial
End Property

Public Property Let dtTick_DtHrPesoInicial(ByVal vData As Date)
    mvardtTick_DtHrPesoInicial = vData
End Property

Public Property Get dtTick_DtHrPesoInicial() As Date
    dtTick_DtHrPesoInicial = mvardtTick_DtHrPesoInicial
End Property

Public Property Let sTick_BalPesoInicial(ByVal vData As String)
    mvarsTick_BalPesoInicial = vData
End Property

Public Property Get sTick_BalPesoInicial() As String
    sTick_BalPesoInicial = mvarsTick_BalPesoInicial
End Property

Public Property Let dTick_PesoFinal(ByVal vData As Double)
    mvardTick_PesoFinal = vData
End Property

Public Property Get dTick_PesoFinal() As Double
    dTick_PesoFinal = mvardTick_PesoFinal
End Property

Public Property Let sTick_OpPesoFinal(ByVal vData As String)
    mvarsTick_OpPesoFinal = vData
End Property

Public Property Get sTick_OpPesoFinal() As String
    sTick_OpPesoFinal = mvarsTick_OpPesoFinal
End Property

Public Property Let dtTick_DtHrPesoFinal(ByVal vData As Date)
    mvardtTick_DtHrPesoFinal = vData
End Property

Public Property Get dtTick_DtHrPesoFinal() As Date
    dtTick_DtHrPesoFinal = mvardtTick_DtHrPesoFinal
End Property

Public Property Let sTick_BalPesoFinal(ByVal vData As String)
    mvarsTick_BalPesoFinal = vData
End Property

Public Property Get sTick_BalPesoFinal() As String
    sTick_BalPesoFinal = mvarsTick_BalPesoFinal
End Property

Public Property Let dTick_PesoLiquido(ByVal vData As Double)
    mvardTick_PesoLiquido = vData
End Property

Public Property Get dTick_PesoLiquido() As Double
    dTick_PesoLiquido = mvardTick_PesoLiquido
End Property

Public Property Let dTick_FatCorrecao(ByVal vData As Double)
    mvardTick_FatCorrecao = vData
End Property

Public Property Get dTick_FatCorrecao() As Double
    dTick_FatCorrecao = mvardTick_FatCorrecao
End Property

Public Property Let dTick_FatConversao(ByVal vData As Double)
    mvardTick_FatConversao = vData
End Property

Public Property Get dTick_FatConversao() As Double
    dTick_FatConversao = mvardTick_FatConversao
End Property

Public Property Let dTick_LiquidoCorrigido(ByVal vData As Double)
    mvardTick_LiquidoCorrigido = vData
End Property

Public Property Get dTick_LiquidoCorrigido() As Double
    dTick_LiquidoCorrigido = mvardTick_LiquidoCorrigido
End Property

Public Property Let sTick_UnidadeAposConversao(ByVal vData As String)
    mvarsTick_UnidadeAposConversao = vData
End Property

Public Property Get sTick_UnidadeAposConversao() As String
    sTick_UnidadeAposConversao = mvarsTick_UnidadeAposConversao
End Property

Public Property Let sTick_DescDocTot(ByVal vData As String)
    mvarsTick_DescDocTot = vData
End Property

Public Property Get sTick_DescDocTot() As String
    sTick_DescDocTot = mvarsTick_DescDocTot
End Property

Public Property Let dTick_PesoTotDoc(ByVal vData As Double)
    mvardTick_PesoTotDoc = vData
End Property

Public Property Get dTick_PesoTotDoc() As Double
    dTick_PesoTotDoc = mvardTick_PesoTotDoc
End Property

Public Property Let dTick_DifOrigemRealPc(ByVal vData As Double)
    mvardTick_DifOrigemRealPc = vData
End Property

Public Property Get dTick_DifOrigemRealPc() As Double
    dTick_DifOrigemRealPc = mvardTick_DifOrigemRealPc
End Property

Public Property Let dTick_DifOrigemRealKg(ByVal vData As Double)
    mvardTick_DifOrigemRealKg = vData
End Property

Public Property Get dTick_DifOrigemRealKg() As Double
    dTick_DifOrigemRealKg = mvardTick_DifOrigemRealKg
End Property

Public Property Let dTick_PesoLiquido1(ByVal vData As Double)
    mvardTick_PesoLiquido1 = vData
End Property

Public Property Get dTick_PesoLiquido1() As Double
    dTick_PesoLiquido1 = mvardTick_PesoLiquido1
End Property

Public Property Let dTick_PesoLiquido2(ByVal vData As Double)
    mvardTick_PesoLiquido2 = vData
End Property

Public Property Get dTick_PesoLiquido2() As Double
    dTick_PesoLiquido2 = mvardTick_PesoLiquido2
End Property

Public Property Let dTick_PesoLiquido3(ByVal vData As Double)
    mvardTick_PesoLiquido3 = vData
End Property

Public Property Get dTick_PesoLiquido3() As Double
    dTick_PesoLiquido3 = mvardTick_PesoLiquido3
End Property

Public Property Let dTick_PesoLiquido4(ByVal vData As Double)
    mvardTick_PesoLiquido4 = vData
End Property

Public Property Get dTick_PesoLiquido4() As Double
    dTick_PesoLiquido4 = mvardTick_PesoLiquido4
End Property

Public Property Let dTick_PesoLiquido5(ByVal vData As Double)
    mvardTick_PesoLiquido5 = vData
End Property

Public Property Get dTick_PesoLiquido5() As Double
    dTick_PesoLiquido5 = mvardTick_PesoLiquido5
End Property

Public Property Let dTick_FatCorrecao1(ByVal vData As Double)
    mvardTick_FatCorrecao1 = vData
End Property

Public Property Get dTick_FatCorrecao1() As Double
    dTick_FatCorrecao1 = mvardTick_FatCorrecao1
End Property

Public Property Let dTick_FatCorrecao2(ByVal vData As Double)
    mvardTick_FatCorrecao2 = vData
End Property

Public Property Get dTick_FatCorrecao2() As Double
    dTick_FatCorrecao2 = mvardTick_FatCorrecao2
End Property

Public Property Let dTick_FatCorrecao3(ByVal vData As Double)
    mvardTick_FatCorrecao3 = vData
End Property

Public Property Get dTick_FatCorrecao3() As Double
    dTick_FatCorrecao3 = mvardTick_FatCorrecao3
End Property

Public Property Let dTick_FatCorrecao4(ByVal vData As Double)
    mvardTick_FatCorrecao4 = vData
End Property

Public Property Get dTick_FatCorrecao4() As Double
    dTick_FatCorrecao4 = mvardTick_FatCorrecao4
End Property

Public Property Let dTick_FatCorrecao5(ByVal vData As Double)
    mvardTick_FatCorrecao5 = vData
End Property

Public Property Get dTick_FatCorrecao5() As Double
    dTick_FatCorrecao5 = mvardTick_FatCorrecao5
End Property

Public Property Let dTick_FatCorrecao6(ByVal vData As Double)
    mvardTick_FatCorrecao6 = vData
End Property

Public Property Get dTick_FatCorrecao6() As Double
    dTick_FatCorrecao6 = mvardTick_FatCorrecao6
End Property

Public Property Let sTick_Observacao1(ByVal vData As String)
    mvarsTick_Observacao1 = vData
End Property

Public Property Get sTick_Observacao1() As String
    sTick_Observacao1 = mvarsTick_Observacao1
End Property

Public Property Let sTick_Observacao2(ByVal vData As String)
    mvarsTick_Observacao2 = vData
End Property

Public Property Get sTick_Observacao2() As String
    sTick_Observacao2 = mvarsTick_Observacao2
End Property

Public Property Let sTick_SenhaPreEntrada(ByVal vData As String)
    mvarsTick_SenhaPreEntrada = vData
End Property

Public Property Get sTick_SenhaPreEntrada() As String
    sTick_SenhaPreEntrada = mvarsTick_SenhaPreEntrada
End Property

Public Property Let dtTick_DtHrPreEntrada(ByVal vData As Date)
    mvardtTick_DtHrPreEntrada = vData
End Property

Public Property Get dtTick_DtHrPreEntrada() As Date
    dtTick_DtHrPreEntrada = mvardtTick_DtHrPreEntrada
End Property

Public Property Let sTick_OpPreEntrada(ByVal vData As String)
    mvarsTick_OpPreEntrada = vData
End Property

Public Property Get sTick_OpPreEntrada() As String
    sTick_OpPreEntrada = mvarsTick_OpPreEntrada
End Property

Public Property Let sTick_MemoObs(ByVal vData As String)
    mvarsTick_MemoObs = vData
End Property

Public Property Get sTick_MemoObs() As String
    sTick_MemoObs = mvarsTick_MemoObs
End Property

Public Property Let dTick_PesoCavInic(ByVal vData As Double)
    mvardTick_PesoCavInic = vData
End Property

Public Property Get dTick_PesoCavInic() As Double
    dTick_PesoCavInic = mvardTick_PesoCavInic
End Property

Public Property Let sTick_PlacaCavInic(ByVal vData As String)
    mvarsTick_PlacaCavInic = vData
End Property

Public Property Get sTick_PlacaCavInic() As String
    sTick_PlacaCavInic = mvarsTick_PlacaCavInic
End Property

Public Property Let dTick_PesoCavFinal(ByVal vData As Double)
    mvardTick_PesoCavFinal = vData
End Property

Public Property Get dTick_PesoCavFinal() As Double
    dTick_PesoCavFinal = mvardTick_PesoCavFinal
End Property

Public Property Let sTick_PlacaCavFinal(ByVal vData As String)
    mvarsTick_PlacaCavFinal = vData
End Property

Public Property Get sTick_PlacaCavFinal() As String
    sTick_PlacaCavFinal = mvarsTick_PlacaCavFinal
End Property

Public Property Let lTick_QtdeTara1(ByVal vData As Long)
    mvarlTick_QtdeTara1 = vData
End Property

Public Property Get lTick_QtdeTara1() As Long
    lTick_QtdeTara1 = mvarlTick_QtdeTara1
End Property

Public Property Let dTick_PesoTara1(ByVal vData As Double)
    mvardTick_PesoTara1 = vData
End Property

Public Property Get dTick_PesoTara1() As Double
    dTick_PesoTara1 = mvardTick_PesoTara1
End Property

Public Property Let dTick_TotTara1(ByVal vData As Double)
    mvardTick_TotTara1 = vData
End Property

Public Property Get dTick_TotTara1() As Double
    dTick_TotTara1 = mvardTick_TotTara1
End Property

Public Property Let lTick_QtdeTara2(ByVal vData As Long)
    mvarlTick_QtdeTara2 = vData
End Property

Public Property Get lTick_QtdeTara2() As Long
    lTick_QtdeTara2 = mvarlTick_QtdeTara2
End Property

Public Property Let dTick_PesoTara2(ByVal vData As Double)
    mvardTick_PesoTara2 = vData
End Property

Public Property Get dTick_PesoTara2() As Double
    dTick_PesoTara2 = mvardTick_PesoTara2
End Property

Public Property Let dTick_TotTara2(ByVal vData As Double)
    mvardTick_TotTara2 = vData
End Property

Public Property Get dTick_TotTara2() As Double
    dTick_TotTara2 = mvardTick_TotTara2
End Property

Public Property Let dTick_TotTaras(ByVal vData As Double)
    mvardTick_TotTaras = vData
End Property

Public Property Get dTick_TotTaras() As Double
    dTick_TotTaras = mvardTick_TotTaras
End Property

Public Property Let dTick_PesoM3(ByVal vData As Double)
    mvardTick_PesoM3 = vData
End Property

Public Property Get dTick_PesoM3() As Double
    dTick_PesoM3 = mvardTick_PesoM3
End Property

Public Property Let iTick_HabCavalo(ByVal vData As Integer)
    mvariTick_HabCavalo = vData
End Property

Public Property Get iTick_HabCavalo() As Integer
    iTick_HabCavalo = mvariTick_HabCavalo
End Property

Public Property Let lTick_Compartimentos(ByVal vData As Long)
    mvarlTick_Compartimentos = vData
End Property

Public Property Get lTick_Compartimentos() As Long
    lTick_Compartimentos = mvarlTick_Compartimentos
End Property

Public Property Let dtTick_DtHrPosPesa(ByVal vData As Date)
    mvardtTick_DtHrPosPesa = vData
End Property

Public Property Get dtTick_DtHrPosPesa() As Date
    dtTick_DtHrPosPesa = mvardtTick_DtHrPosPesa
End Property

Public Property Let sTick_TmpVeicEmpr(ByVal vData As String)
    mvarsTick_TmpVeicEmpr = vData
End Property

Public Property Get sTick_TmpVeicEmpr() As String
    sTick_TmpVeicEmpr = mvarsTick_TmpVeicEmpr
End Property

Public Property Let sTick_NumTransp(ByVal vData As String)
    mvarsTick_NumTransp = vData
End Property

Public Property Get sTick_NumTransp() As String
    sTick_NumTransp = mvarsTick_NumTransp
End Property

Public Property Let sTick_CampoUsu1(ByVal vData As String)
    mvarsTick_CampoUsu1 = vData
End Property

Public Property Get sTick_CampoUsu1() As String
    sTick_CampoUsu1 = mvarsTick_CampoUsu1
End Property

Public Property Let sTick_CampoUsu2(ByVal vData As String)
    mvarsTick_CampoUsu2 = vData
End Property

Public Property Get sTick_CampoUsu2() As String
    sTick_CampoUsu2 = mvarsTick_CampoUsu2
End Property

Public Property Let sTick_CampoUsu3(ByVal vData As String)
    mvarsTick_CampoUsu3 = vData
End Property

Public Property Get sTick_CampoUsu3() As String
    sTick_CampoUsu3 = mvarsTick_CampoUsu3
End Property

Public Property Let sTick_CampoUsu4(ByVal vData As String)
    mvarsTick_CampoUsu4 = vData
End Property

Public Property Get sTick_CampoUsu4() As String
    sTick_CampoUsu4 = mvarsTick_CampoUsu4
End Property

Public Property Let sTick_CampoUsu5(ByVal vData As String)
    mvarsTick_CampoUsu5 = vData
End Property

Public Property Get sTick_CampoUsu5() As String
    sTick_CampoUsu5 = mvarsTick_CampoUsu5
End Property

Public Property Let sTick_CampoUsu6(ByVal vData As String)
    mvarsTick_CampoUsu6 = vData
End Property

Public Property Get sTick_CampoUsu6() As String
    sTick_CampoUsu6 = mvarsTick_CampoUsu6
End Property

Public Property Let sTick_CampoUsu7(ByVal vData As String)
    mvarsTick_CampoUsu7 = vData
End Property

Public Property Get sTick_CampoUsu7() As String
    sTick_CampoUsu7 = mvarsTick_CampoUsu7
End Property

Public Property Let sTick_CampoUsu8(ByVal vData As String)
    mvarsTick_CampoUsu8 = vData
End Property

Public Property Get sTick_CampoUsu8() As String
    sTick_CampoUsu8 = mvarsTick_CampoUsu8
End Property

Public Property Let sTick_CampoUsu9(ByVal vData As String)
    mvarsTick_CampoUsu9 = vData
End Property

Public Property Get sTick_CampoUsu9() As String
    sTick_CampoUsu9 = mvarsTick_CampoUsu9
End Property

Public Property Let sTick_CampoUsu10(ByVal vData As String)
    mvarsTick_CampoUsu10 = vData
End Property

Public Property Get sTick_CampoUsu10() As String
    sTick_CampoUsu10 = mvarsTick_CampoUsu10
End Property

Public Property Let dTick_PesoLiqCorrUsu(ByVal vData As Double)
    mvardTick_PesoLiqCorrUsu = vData
End Property

Public Property Get dTick_PesoLiqCorrUsu() As Double
    dTick_PesoLiqCorrUsu = mvardTick_PesoLiqCorrUsu
End Property

Public Property Let lTick_NumOcupante(ByVal vData As Long)
    mvarlTick_NumOcupante = vData
End Property

Public Property Get lTick_NumOcupante() As Long
    lTick_NumOcupante = mvarlTick_NumOcupante
End Property
