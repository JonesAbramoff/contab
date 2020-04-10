VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContab"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function Mascara_MascararContaInt Lib "ADCUSR.DLL" Alias "AD_Mascara_MascaraContaExt" (ByVal lID_Mascara As Long, ByVal sContaNaoMascarada As String, ByVal sContaMascarada As String) As Long
Private Declare Function Mascara_MascararCclInt Lib "ADCUSR.DLL" Alias "AD_Mascara_MascaraCclExt" (ByVal lID_Mascara As Long, ByVal sCclNaoMascarado As String, ByVal sCclMascarado As String) As Long
Private Declare Function Mascara_RetornaContaNoNivelInt Lib "ADCUSR.DLL" Alias "AD_Mascara_RetornaContaNoNivelExt" (ByVal lID_Mascara As Long, ByVal i_nivel As Integer, ByVal lp_conta As String, ByVal lp_conta_no_nivel As String) As Long
Private Declare Function Mascara_RetornaCclNoNivelInt Lib "ADCUSR.DLL" Alias "AD_Mascara_RetornaCclNoNivelExt" (ByVal lID_Mascara As Long, ByVal i_nivel As Integer, ByVal lp_ccl As String, ByVal lp_ccl_no_nivel As String) As Long
Private Declare Function Mascara_RetornaContaPaiInt Lib "ADCUSR.DLL" Alias "AD_Mascara_RetornaContaPaiExt" (ByVal lID_Mascara As Long, ByVal lp_conta As String, ByVal lp_conta_pai As String) As Long
Private Declare Function Mascara_RetornaCclPaiInt Lib "ADCUSR.DLL" Alias "AD_Mascara_RetornaCclPaiExt" (ByVal lID_Mascara As Long, ByVal lp_ccl As String, ByVal lp_ccl_pai As String) As Long
Private Declare Function Mascara_RetornaContaEnxutaInt Lib "ADCUSR.DLL" Alias "AD_Mascara_RetornaContaEnxuta" (ByVal lID_Mascara As Long, ByVal lp_conta As String, ByVal lpContaEnxuta As String) As Long
Private Declare Function Mascara_RetornaContaTelaInt Lib "ADCUSR.DLL" Alias "AD_Mascara_RetornaContaTela" (ByVal lID_Mascara As Long, ByVal lp_conta As String, ByVal lpContaTela As String) As Long
Private Declare Function Mascara_RetornaCclEnxutaInt Lib "ADCUSR.DLL" Alias "AD_Mascara_RetornaCclEnxuta" (ByVal lID_Mascara As Long, ByVal lp_ccl As String, ByVal lpCclEnxuta As String) As Long
Private Declare Function Mascara_RetornaMaiorContaInt Lib "ADCUSR.DLL" Alias "AD_Mascara_RetornaMaiorConta" (ByVal lID_Mascara As Long, ByVal lp_conta As String, ByVal lpMaiorConta As String) As Long
Private Declare Function Mascara_RetornaCclInt Lib "ADCUSR.DLL" Alias "AD_Mascara_RetornaCclExt" (ByVal lID_Mascara As Long, ByVal lp_conta As String, ByVal lpCcl As String) As Long
Private Declare Function Mascara_ExisteProxNivelCtaInt Lib "ADCUSR.DLL" Alias "AD_Mascara_ExisteProxNivelCta" (ByVal lID_Mascara As Long, ByVal i_nivel As Integer, ByVal lp_conta As String) As Long
Private Declare Function Mascara_ExisteProxNivelCclInt Lib "ADCUSR.DLL" Alias "AD_Mascara_ExisteProxNivelCcl" (ByVal lID_Mascara As Long, ByVal i_nivel As Integer, ByVal lp_ccl As String) As Long
Private Declare Function Mascara_ObterInt Lib "ADCUSR.DLL" Alias "AD_Mascara_ObterExt" (ByVal lID_Mascara As Long, ByVal sMascaraConta As String) As Long
Private Declare Function Mascara_ObterCclInt Lib "ADCUSR.DLL" Alias "AD_Mascara_ObterExtCcl" (ByVal lID_Mascara As Long, ByVal sMascaraCcl As String) As Long
Private Declare Function Mascara_Conta_ObterNivelInt Lib "ADCUSR.DLL" Alias "AD_Mascara_CtaObterNivel" (ByVal lID_Mascara As Long, ByVal lp_conta As String, i_nivel As Integer) As Long
Private Declare Function Mascara_Ccl_ObterNivelInt Lib "ADCUSR.DLL" Alias "AD_Mascara_CclObterNivel" (ByVal lID_Mascara As Long, ByVal lp_ccl As String, i_nivel As Integer) As Long
Private Declare Function NumSegContaInt Lib "ADCUSR.DLL" Alias "AD_Mascara_NumSegConta" (ByVal lID_Mascara As Long) As Long
Private Declare Function NumSegCclInt Lib "ADCUSR.DLL" Alias "AD_Mascara_NumSegCcl" (ByVal lID_Mascara As Long) As Long
Private Declare Function SegmentoContaInt Lib "ADCUSR.DLL" Alias "AD_Mascara_SegmentoConta" (ByVal lID_Mascara As Long, ByVal iNivel As Integer, iPreenchimento As Integer, iTamanho As Integer, iTipo As Integer, ByVal sDelimitador As String) As Long
Private Declare Function SegmentoCclInt Lib "ADCUSR.DLL" Alias "AD_Mascara_SegmentoCcl" (ByVal lID_Mascara As Long, ByVal iNivel As Integer, iPreenchimento As Integer, iTamanho As Integer, iTipo As Integer, ByVal sDelimitador As String) As Long

Public Property Get iESTAglutinaLancamPorDia() As Integer
    iESTAglutinaLancamPorDia = GL_objContabInt.mvariESTAglutinaLancamPorDia
End Property

Public Property Get iESTGeraLoteAutomatico() As Integer
    iESTGeraLoteAutomatico = GL_objContabInt.mvariESTGeraLoteAutomatico
End Property

Public Property Get iESTTamanhoMinimoContaGrid() As Integer
    iESTTamanhoMinimoContaGrid = GL_objContabInt.mvariESTTamanhoMinimoContaGrid
End Property

Public Property Let iESTTamanhoMinimoContaGrid(ByVal vData As Integer)
    GL_objContabInt.mvariESTTamanhoMinimoContaGrid = vData
End Property

Public Property Get iESTTamanhoMinimoProdutoGrid() As Integer
    iESTTamanhoMinimoProdutoGrid = GL_objContabInt.mvariESTTamanhoMinimoProdutoGrid
End Property

Public Property Let iESTTamanhoMinimoProdutoGrid(ByVal vData As Integer)
    GL_objContabInt.mvariESTTamanhoMinimoProdutoGrid = vData
End Property

Public Property Get iFATAglutinaLancamPorDia() As Integer
    iFATAglutinaLancamPorDia = GL_objContabInt.mvariFATAglutinaLancamPorDia
End Property

Public Property Get iFATGeraLoteAutomatico() As Integer
    iFATGeraLoteAutomatico = GL_objContabInt.mvariFATGeraLoteAutomatico
End Property

Public Property Get iTESAglutinaLancamPorDia() As Integer
    iTESAglutinaLancamPorDia = GL_objContabInt.mvariTESAglutinaLancamPorDia
End Property

Public Property Get iTESGeraLoteAutomatico() As Integer
    iTESGeraLoteAutomatico = GL_objContabInt.mvariTESGeraLoteAutomatico
End Property

Public Property Get iCRAglutinaLancamPorDia() As Integer
    iCRAglutinaLancamPorDia = GL_objContabInt.mvariCRAglutinaLancamPorDia
End Property

Public Property Get iCRGeraLoteAutomatico() As Integer
    iCRGeraLoteAutomatico = GL_objContabInt.mvariCRGeraLoteAutomatico
End Property

Public Property Get iCPAglutinaLancamPorDia() As Integer
    iCPAglutinaLancamPorDia = GL_objContabInt.mvariCPAglutinaLancamPorDia
End Property

Public Property Get iCPGeraLoteAutomatico() As Integer
    iCPGeraLoteAutomatico = GL_objContabInt.mvariCPGeraLoteAutomatico
End Property

Public Property Get giRateioOn() As Integer
        giRateioOn = GL_objContabInt.mvargiRateioOn
End Property

Public Property Set gobjRateioOff(ByVal vData As Object)
Attribute gobjRateioOff.VB_Description = "Se preenchido, contém a chave de um rateio off-line."
    Set GL_objContabInt.mvargobjRateioOff = vData
End Property

Public Property Get gobjRateioOff() As ClassRateioOff
    Set gobjRateioOff = GL_objContabInt.mvargobjRateioOff
End Property

Public Property Let giRateioOn(ByVal vData As Integer)
Attribute giRateioOn.VB_Description = "Se estiver preenchido, contém um código de rateio on-line."
    GL_objContabInt.mvargiRateioOn = vData
End Property

'Public Property Get giRateioOn() As Integer
'        giRateioOn = mvargiRateioOn
'End Property

Public Property Let giNaturezaConta(ByVal vData As Integer)
Attribute giNaturezaConta.VB_Description = "Contém a natureza default da conta. 1=crédito 2=débito."
    GL_objContabInt.mvargiNaturezaConta = vData
End Property

Public Property Get giNaturezaConta() As Integer
    giNaturezaConta = GL_objContabInt.mvargiNaturezaConta
End Property

Public Property Let giTipoConta(ByVal vData As Integer)
Attribute giTipoConta.VB_Description = "Contém o tipo de conta default. 1=Sintética 3=Analítica"
    GL_objContabInt.mvargiTipoConta = vData
End Property

Public Property Get giTipoConta() As Integer
    giTipoConta = GL_objContabInt.mvargiTipoConta
End Property

Public Property Let giHistPadrao(ByVal vData As Integer)
Attribute giHistPadrao.VB_Description = "se estiver preenchido, contém um código de histórico padrão."
    GL_objContabInt.mvargiHistPadrao = vData
End Property

Public Property Get giHistPadrao() As Integer
    giHistPadrao = GL_objContabInt.mvargiHistPadrao
End Property

Public Property Let gsCcl(ByVal vData As String)
Attribute gsCcl.VB_Description = "se estiver preenchido, contém um código de centro de custo/lucro"
    GL_objContabInt.mvargsCcl = vData
End Property

Public Property Get gsCcl() As String
    gsCcl = GL_objContabInt.mvargsCcl
End Property

Public Property Let gsConta(ByVal vData As String)
Attribute gsConta.VB_Description = "se estiver preenchido, contem uma conta contábil"
    GL_objContabInt.mvargsConta = vData
End Property

Public Property Get gsConta() As String
    gsConta = GL_objContabInt.mvargsConta
End Property

Public Property Get giSetupUsoCcl() As Integer
Attribute giSetupUsoCcl.VB_Description = "Indica o uso do Ccl. 0 = o sistema não usa ccl 1= o sistema usa ccl contabil 2= o sistema usa ccl extra contabil"
    giSetupUsoCcl = GL_objContabInt.mvargiSetupUsoCcl
End Property

Public Property Get giSetupDocPorPeriodo() As Integer
Attribute giSetupDocPorPeriodo.VB_Description = "Indica se a numeração do documento(voucher) é zerado por periodo ou não.                                                                  1=zerado por periodo                                                              0=zerado por exercicio   "
    giSetupDocPorPeriodo = GL_objContabInt.mvargiSetupDocPorPeriodo
End Property

Public Property Get giSetupLotePorPeriodo() As Integer
Attribute giSetupLotePorPeriodo.VB_Description = "Indica se a numeração do lote é zerado por periodo ou não.  1=zerado por periodo                                                              0=zerado por exercicio   "
    giSetupLotePorPeriodo = GL_objContabInt.mvargiSetupLotePorPeriodo
End Property

Public Property Let giPeriodoAtual(ByVal vData As Integer)
Attribute giPeriodoAtual.VB_Description = "contém o último periodo utilizado"
    GL_objContabInt.mvargiPeriodoAtual = vData
End Property

Public Property Get giPeriodoAtual() As Integer
    giPeriodoAtual = GL_objContabInt.mvargiPeriodoAtual
End Property

Public Property Let gsOrigemAtual(ByVal vData As String)
Attribute gsOrigemAtual.VB_Description = "contém a ultima origem utilizada"
    GL_objContabInt.mvargsOrigemAtual = vData
End Property

Public Property Get gsOrigemAtual() As String
    gsOrigemAtual = GL_objContabInt.mvargsOrigemAtual
End Property

Public Property Let giContabCentralizada(ByVal vData As Integer)
    GL_objContabInt.mvargiContabCentralizada = vData
End Property

Public Property Get giContabCentralizada() As Integer
    giContabCentralizada = GL_objContabInt.mvargiContabCentralizada
End Property

Public Property Get gobjColOrigem() As ClassColOrigemContab
Attribute gobjColOrigem.VB_Description = "Contem as Origens editáveis pelo módulo de contabilidade"
    Set gobjColOrigem = GL_objContabInt.mvargobjColOrigem
End Property

Public Property Get gobjColNaturezaConta() As ClassColNaturezaConta
    Set gobjColNaturezaConta = GL_objContabInt.mvargobjColNaturezaConta
End Property

Public Property Get gobjColTipoConta() As ClassColTipoConta
    Set gobjColTipoConta = GL_objContabInt.mvargobjColTipoConta
End Property

Public Property Get gobjColCodigoSegmento() As ClassColCodigoSegmento
    Set gobjColCodigoSegmento = GL_objContabInt.mvargobjColCodigoSegmento
End Property

Public Property Get gobjColPreenchimento() As ClassColPreenchimento
    Set gobjColPreenchimento = GL_objContabInt.mvargobjColPreenchimento
End Property

Public Property Get gobjColTipoSegmento() As ClassColTipoSegmento
    Set gobjColTipoSegmento = GL_objContabInt.mvargobjColTipoSegmento
End Property

Public Property Let giExercicioAtual(ByVal vData As Integer)
Attribute giExercicioAtual.VB_Description = "Contem o ultimo exercicio utilizado"
    GL_objContabInt.mvargiExercicioAtual = vData
End Property

Public Property Get giExercicioAtual() As Integer
    giExercicioAtual = GL_objContabInt.mvargiExercicioAtual
End Property

Public Property Let giHistoricoObrigatorio(ByVal vData As Integer)
    GL_objContabInt.mvargiHistoricoObrigatorio = vData
End Property

Public Property Get giHistoricoObrigatorio() As Integer
    giHistoricoObrigatorio = GL_objContabInt.mvargiHistoricoObrigatorio
End Property

Public Property Let giValidaCtaCcl(ByVal vData As Integer)
    GL_objContabInt.mvargiValidaCtaCcl = vData
End Property

Public Property Get giValidaCtaCcl() As Integer
    giValidaCtaCcl = GL_objContabInt.mvargiValidaCtaCcl
End Property

Function NumSegConta() As Long
    NumSegConta = NumSegContaInt(GL_lMascara)
End Function

Function NumSegCcl() As Long
    NumSegCcl = NumSegCclInt(GL_lMascara)
End Function

Public Property Let giCtbLctoExcluiEstorno(ByVal vData As Integer)
    GL_objContabInt.mvargiCtbLctoExcluiEstorno = vData
End Property

Public Property Get giCtbLctoExcluiEstorno() As Integer
    giCtbLctoExcluiEstorno = GL_objContabInt.mvargiCtbLctoExcluiEstorno
End Property


Function SegmentoConta(iNivel As Integer, objSegmento As ClassSegmento) As Long
    
Dim lErro As Long
Dim tSegmento As typeSegmento
    
On Error GoTo Erro_SegmentoConta

    tSegmento.sDelimitador = " "
    
    lErro = SegmentoContaInt(GL_lMascara, iNivel, tSegmento.iPreenchimento, tSegmento.iTamanho, tSegmento.iTipo, tSegmento.sDelimitador)
    If lErro <> AD_BOOL_TRUE Then Error 5814
    
    objSegmento.iPreenchimento = tSegmento.iPreenchimento
    objSegmento.iTamanho = tSegmento.iTamanho
    objSegmento.iTipo = tSegmento.iTipo
    objSegmento.sDelimitador = tSegmento.sDelimitador
    
    SegmentoConta = SUCESSO
    
    Exit Function
    
Erro_SegmentoConta:

    SegmentoConta = Err
    
    Select Case Err
    
        Case 5814
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147768)
    
    End Select
    
    Exit Function
    
End Function

Function SegmentoCcl(iNivel As Integer, objSegmento As ClassSegmento) As Long

Dim lErro As Long
Dim tSegmento As typeSegmento

On Error GoTo Erro_SegmentoCcl

    tSegmento.sDelimitador = " "
    
    lErro = SegmentoCclInt(GL_lMascara, iNivel, tSegmento.iPreenchimento, tSegmento.iTamanho, tSegmento.iTipo, tSegmento.sDelimitador)
    If lErro <> AD_BOOL_TRUE Then Error 5813
    
    objSegmento.iPreenchimento = tSegmento.iPreenchimento
    objSegmento.iTamanho = tSegmento.iTamanho
    objSegmento.iTipo = tSegmento.iTipo
    objSegmento.sDelimitador = tSegmento.sDelimitador
    
    SegmentoCcl = SUCESSO
    
    Exit Function
    
Erro_SegmentoCcl:

    SegmentoCcl = Err
    
    Select Case Err
    
        Case 5813
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147769)
    
    End Select
    
    Exit Function

End Function

Function Mascara_MascararConta(sConta As String, sContaMascarada As String) As Long
    
Dim lErro As Long

On Error GoTo Erro_Mascara_MascararConta

    sContaMascarada = String(STRING_CONTA * 2, 0) 'o retorno pode ser maior por causa dos delimitadores
    
    lErro = Mascara_MascararContaInt(GL_lMascara, sConta, sContaMascarada)
    If lErro <> AD_BOOL_TRUE Then Error 5807
    
    sContaMascarada = StringZ(sContaMascarada)
    
    Mascara_MascararConta = SUCESSO
    
    Exit Function
    
Erro_Mascara_MascararConta:

    Mascara_MascararConta = Err
    
    Select Case Err
    
        Case 5807
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147770)
    
    End Select
    
    Exit Function

End Function

Function Mascara_MascararCcl(sCcl As String, sCclMascarado As String) As Long

Dim lErro As Long

On Error GoTo Erro_Mascara_MascararCcl

    sCclMascarado = String(STRING_CCL * 2, 0) 'o retorno pode ser maior por causa dos delimitadores
    
    lErro = Mascara_MascararCclInt(GL_lMascara, sCcl, sCclMascarado)
    If lErro <> AD_BOOL_TRUE Then Error 5808
    
    sCclMascarado = StringZ(sCclMascarado)
    
    Mascara_MascararCcl = SUCESSO
    
    Exit Function
    
Erro_Mascara_MascararCcl:

    Mascara_MascararCcl = Err
    
    Select Case Err
    
        Case 5808
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147771)
    
    End Select
    
    Exit Function

End Function

Function Mascara_RetornaContaNoNivel(ByVal i_nivel As Integer, ByVal lp_conta As String, lp_conta_no_nivel As String) As Long

Dim lErro As Long

On Error GoTo Erro_Mascara_RetornaContaNoNivel

    lp_conta_no_nivel = String(STRING_CONTA, 0)
    
    lErro = Mascara_RetornaContaNoNivelInt(GL_lMascara, i_nivel, lp_conta, lp_conta_no_nivel)
    If lErro <> AD_BOOL_TRUE Then Error 5809
    
    lp_conta_no_nivel = StringZ(lp_conta_no_nivel)
    Mascara_RetornaContaNoNivel = SUCESSO
    
    Exit Function
    
Erro_Mascara_RetornaContaNoNivel:

    Mascara_RetornaContaNoNivel = Err
    
    Select Case Err
    
        Case 5809
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147772)
    
    End Select
    
    Exit Function

End Function

Function Mascara_RetornaCclNoNivel(ByVal i_nivel As Integer, ByVal lp_ccl As String, lp_ccl_no_nivel As String) As Long

Dim lErro As Long

On Error GoTo Erro_Mascara_RetornaCclNoNivel

    lp_ccl_no_nivel = String(STRING_CCL, 0)
    
    lErro = Mascara_RetornaCclNoNivelInt(GL_lMascara, i_nivel, lp_ccl, lp_ccl_no_nivel)
    If lErro <> AD_BOOL_TRUE Then Error 7252
    
    lp_ccl_no_nivel = StringZ(lp_ccl_no_nivel)
    Mascara_RetornaCclNoNivel = SUCESSO
    
    Exit Function
    
Erro_Mascara_RetornaCclNoNivel:

    Mascara_RetornaCclNoNivel = Err
    
    Select Case Err
    
        Case 7252
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147773)
    
    End Select
    
    Exit Function

End Function

Function Mascara_RetornaContaPai(ByVal lp_conta As String, lpContaPai As String) As Long
'coloca em lpContaPai a conta pai de lp_conta. Se nao houver coloca "".

Dim lErro As Long
    
On Error GoTo Erro_Mascara_RetornaContaPai

    lpContaPai = String(STRING_CONTA, 0)
    
    lErro = Mascara_RetornaContaPaiInt(GL_lMascara, lp_conta, lpContaPai)
    If lErro <> AD_BOOL_TRUE Then Error 5810
    
    lpContaPai = StringZ(lpContaPai)
    Mascara_RetornaContaPai = SUCESSO
    
    Exit Function
    
Erro_Mascara_RetornaContaPai:

    Mascara_RetornaContaPai = Err
    
    Select Case Err
    
        Case 5810
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147774)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_RetornaCclPai(ByVal lp_ccl As String, lp_ccl_pai As String) As Long
'coloca em lp_ccl_pai o ccl pai de lp_ccl. Se nao houver coloca "".

Dim lErro As Long
    
On Error GoTo Erro_Mascara_RetornaCclPai

    lp_ccl_pai = String(STRING_CCL, 0)
    
    lErro = Mascara_RetornaCclPaiInt(GL_lMascara, lp_ccl, lp_ccl_pai)
    If lErro <> AD_BOOL_TRUE Then Error 7253
    
    lp_ccl_pai = StringZ(lp_ccl_pai)
    Mascara_RetornaCclPai = SUCESSO
    
    Exit Function
    
Erro_Mascara_RetornaCclPai:

    Mascara_RetornaCclPai = Err
    
    Select Case Err
    
        Case 7253
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147775)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_RetornaCclEnxuta(ByVal lp_ccl As String, lpCclEnxuta As String) As Long
'pega o ccl no formato do banco de dados e retorna somente com os segmentos preenchidos

Dim lErro As Long
    
On Error GoTo ERRO_MASCARA_RETORNACCLENXUTA

    lpCclEnxuta = String(STRING_CCL, 0)
    
    lErro = Mascara_RetornaCclEnxutaInt(GL_lMascara, lp_ccl, lpCclEnxuta)
    If lErro <> AD_BOOL_TRUE Then Error 7259
    
    lpCclEnxuta = StringZ(lpCclEnxuta)
    Mascara_RetornaCclEnxuta = SUCESSO
    
    Exit Function
    
ERRO_MASCARA_RETORNACCLENXUTA:

    Mascara_RetornaCclEnxuta = Err
    
    Select Case Err
    
        Case 7259
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147776)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_RetornaContaEnxuta(ByVal lp_conta As String, lpContaEnxuta As String) As Long
Attribute Mascara_RetornaContaEnxuta.VB_Description = "pega a conta no formato do banco de dados e retorna mascarada somente com os segmentos preenchidos\r\n"
'pega a conta no formato do banco de dados e retorna somente com os segmentos preenchidos

Dim lErro As Long
    
On Error GoTo ERRO_MASCARA_RETORNACONTAENXUTA

    lpContaEnxuta = String(STRING_CONTA, 0)
    
    lErro = Mascara_RetornaContaEnxutaInt(GL_lMascara, lp_conta, lpContaEnxuta)
    If lErro <> AD_BOOL_TRUE Then Error 5811
    
    lpContaEnxuta = StringZ(lpContaEnxuta)
    Mascara_RetornaContaEnxuta = SUCESSO
    
    Exit Function
    
ERRO_MASCARA_RETORNACONTAENXUTA:

    Mascara_RetornaContaEnxuta = Err
    
    Select Case Err
    
        Case 5811
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147777)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_RetornaContaTela(ByVal lp_conta As String, lpContaTela As String) As Long
'pega a conta no formato do banco de dados e retorna como se tivesse sido colocada numa masked edit

Dim lErro As Long
    
On Error GoTo Erro_Mascara_RetornaContaTela

    lpContaTela = String(STRING_CONTA, 0)
    
    lErro = Mascara_RetornaContaTelaInt(GL_lMascara, lp_conta, lpContaTela)
    If lErro <> AD_BOOL_TRUE Then Error 32180
    
    lpContaTela = StringZ(lpContaTela)
    
    Mascara_RetornaContaTela = SUCESSO
    
    Exit Function
    
Erro_Mascara_RetornaContaTela:

    Mascara_RetornaContaTela = Err
    
    Select Case Err
    
        Case 32180
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147778)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_RetornaUltimaConta(ByVal lp_conta As String, lpMaiorConta As String) As Long
'pega a conta no formato do banco de dados e retorna a maior conta possivel no nivel dela,
'para a conta pai (no nivel imediatamente acima)

Dim lErro As Long
    
On Error GoTo ERRO_MASCARA_RETORNAULTIMACONTA

    lpMaiorConta = String(STRING_CONTA, 0)
    
    lErro = Mascara_RetornaMaiorContaInt(GL_lMascara, lp_conta, lpMaiorConta)
    If lErro <> AD_BOOL_TRUE Then Error 9814
    
    lpMaiorConta = StringZ(lpMaiorConta)
    Mascara_RetornaUltimaConta = SUCESSO
    
    Exit Function
    
ERRO_MASCARA_RETORNAULTIMACONTA:

    Mascara_RetornaUltimaConta = Err
    
    Select Case Err
    
        Case 9814
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147779)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_RetornaCcl(ByVal lp_conta As String, lpCcl As String) As Long
'coloca em lpCcl o centro de custo em lp_conta. Se nao houver coloca "".

Dim lErro As Long
    
On Error GoTo Erro_Mascara_RetornaCcl

    lpCcl = String(STRING_CCL, 0)
    
    lErro = Mascara_RetornaCclInt(GL_lMascara, lp_conta, lpCcl)
    If lErro <> AD_BOOL_TRUE Then Error 7900
    
    lpCcl = StringZ(lpCcl)
    Mascara_RetornaCcl = SUCESSO
    
    Exit Function
    
Erro_Mascara_RetornaCcl:

    Mascara_RetornaCcl = Err
    
    Select Case Err
    
        Case 7900
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147780)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_ExisteProxNivelCcl(ByVal i_nivel As Integer, ByVal lp_ccl As String) As Long

Dim lErro As Long
    
On Error GoTo Erro_Mascara_ExisteProxNivelCcl

    lErro = Mascara_ExisteProxNivelCclInt(GL_lMascara, i_nivel, lp_ccl)
    If lErro <> AD_BOOL_TRUE Then Error 7258
    
    Mascara_ExisteProxNivelCcl = SUCESSO
    
    Exit Function
    
Erro_Mascara_ExisteProxNivelCcl:

    Mascara_ExisteProxNivelCcl = Err
    
    Select Case Err
    
        Case 7258
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147781)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_ExisteProxNivelCta(ByVal i_nivel As Integer, ByVal lp_conta As String) As Long

Dim lErro As Long
    
On Error GoTo Erro_Mascara_ExisteProxNivelCta

    lErro = Mascara_ExisteProxNivelCtaInt(GL_lMascara, i_nivel, lp_conta)
    If lErro <> AD_BOOL_TRUE Then Error 5812
    
    Mascara_ExisteProxNivelCta = SUCESSO
    
    Exit Function
    
Erro_Mascara_ExisteProxNivelCta:

    Mascara_ExisteProxNivelCta = Err
    
    Select Case Err
    
        Case 5812
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147782)
    
    End Select
    
    Exit Function
    
End Function

Function MascaraConta(sMascara As String) As Long
'coloca em sMascara a mascara de cta.

Dim lErro As Long

On Error GoTo Erro_MascaraConta

    sMascara = String(STRING_CONTA * 2, 0) 'o retorno pode ser maior por causa dos delimitadores

    lErro = Mascara_ObterInt(GL_lMascara, sMascara)
    If lErro <> AD_BOOL_TRUE Then Error 5815
    
    sMascara = StringZ(sMascara)
    
    MascaraConta = SUCESSO
    
    Exit Function
    
Erro_MascaraConta:

    MascaraConta = Err
    
    Select Case Err
    
        Case 5815
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147783)
    
    End Select
    
    Exit Function

End Function

Function MascaraCcl(sMascaraCcl As String) As Long
'coloca em sMascaraCcl a mascara do centro de custo.

Dim lErro As Long

On Error GoTo Erro_MascaraCcl

    sMascaraCcl = String(STRING_CCL * 2, 0) 'o retorno pode ser maior por causa dos delimitadores
    
    lErro = Mascara_ObterCclInt(GL_lMascara, sMascaraCcl)
    If lErro <> AD_BOOL_TRUE Then Error 5816
    
    sMascaraCcl = StringZ(sMascaraCcl)
    
    MascaraCcl = SUCESSO
    
    Exit Function
    
Erro_MascaraCcl:

    MascaraCcl = Err
    
    Select Case Err
    
        Case 5816
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147784)
    
    End Select
    
    Exit Function

End Function

Function Mascara_Conta_ObterNivel(sConta As String, i_nivel As Integer) As Long
'retorna nivel da conta incluindo centro de custo
'contabil valido como um nivel, se houver
Dim lErro As Long

On Error GoTo Erro_Mascara_ContaObterNivel

    lErro = Mascara_Conta_ObterNivelInt(GL_lMascara, sConta, i_nivel)
    If lErro <> AD_BOOL_TRUE Then Error 5817
    
    Mascara_Conta_ObterNivel = SUCESSO
    
    Exit Function
    
Erro_Mascara_ContaObterNivel:

    Mascara_Conta_ObterNivel = Err
    
    Select Case Err
    
        Case 5817
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147785)
    
    End Select
    
    Exit Function
    
End Function

Function Mascara_Ccl_ObterNivel(sCcl As String, i_nivel As Integer) As Long
'retorna nivel do centro de custo
Dim lErro As Long

On Error GoTo Erro_Mascara_CclObterNivel

    lErro = Mascara_Ccl_ObterNivelInt(GL_lMascara, sCcl, i_nivel)
    If lErro <> AD_BOOL_TRUE Then Error 7254
    
    Mascara_Ccl_ObterNivel = SUCESSO
    
    Exit Function
    
Erro_Mascara_CclObterNivel:

    Mascara_Ccl_ObterNivel = Err
    
    Select Case Err
    
        Case 7254
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147786)
    
    End Select
    
    Exit Function
    
End Function

Function CTB_Inicializa_Periodo_ExercicioAtual(dtData As Date) As Long

Dim objPeriodo As New ClassPeriodo
Dim lErro As Long
Dim colExerciciosAbertos As New Collection
Dim objExercicio As ClassExercicio
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CTB_Inicializa_Periodo_ExercicioAtual

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        'Coloca o periodo relativo a data na tela
        lErro = Periodo_Le2(dtData, objPeriodo, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 83475 Then gError 44658
        
        'se o periodo de dtData ainda não estiver cadastrado, tenta encontrar o ultimo exercicio aberto e colocar a ultima data valida como data do sistema
        If lErro = 83475 Then
            
            'le todos os exercicios abertos e coloca-os em colExerciciosAbertos
            lErro = CF("Exercicios_Abertos_Le_Todos", colExerciciosAbertos)
            If lErro <> SUCESSO Then gError 83476
            
            If colExerciciosAbertos.Count = 0 Then gError 83477
            
            'pega o ultimo exercicio aberto
            Set objExercicio = colExerciciosAbertos.Item(colExerciciosAbertos.Count)
            
            'Coloca o periodo relativo a data fim do ultimo exercicio aberto na tela
            lErro = Periodo_Le2(objExercicio.dtDataFim, objPeriodo, giFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 83475 Then gError 83477
            
            'se não foi encontrado um periodo para esta data
            If lErro = 83475 Then gError 83478
            
            vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_DATA_TROCADA_ULTIMO_PERIODO_ABERTO", CStr(dtData), CStr(objExercicio.dtDataFim), CStr(dtData))
            
            dtData = objExercicio.dtDataFim
            
        End If
        
        giExercicioAtual = objPeriodo.iExercicio
        giPeriodoAtual = objPeriodo.iPeriodo
        
    End If

    CTB_Inicializa_Periodo_ExercicioAtual = SUCESSO
    
    Exit Function

Erro_CTB_Inicializa_Periodo_ExercicioAtual:

    CTB_Inicializa_Periodo_ExercicioAtual = gErr
    
    Select Case gErr
    
        Case 44658, 83476
        
        Case 83477
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_HA_EXERCICIO_CONTABIL_ABERTO", gErr)
    
        Case 83478
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_NAO_ENCONTRADO_DATA", gErr, CStr(dtData))
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147787)
    
    End Select
    
    Exit Function
    
End Function

'????? Transferir para ClassSelect.
'???? ATenção!!! Quando transferir, a chamada de Periodo_Le2 na rotina CTB_Inicializa_Periodo_ExercicioAtual tem que ser prefixada de "X.". Obrigado. Mario (02/01/01)
Function Periodo_Le2(ByVal dtData As Date, objPeriodo As ClassPeriodo, iFilialEmpresa As Integer) As Long
'le o registro referente ao periodo em questão e coloca os dados em objPeriodo

Dim lComando As Long
Dim lErro As Long
Dim tPeriodo As typePeriodo

On Error GoTo Erro_Periodo_Le2

    tPeriodo.sNomeExterno = String(STRING_PERIODO_NOME_EXTERNO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 83472
    
    lErro = Comando_Executar(lComando, "SELECT Periodo.Exercicio, Periodo.Periodo, NomeExterno, Lancamentos, Apurado, DataInicio, DataFim, Lote, Fechado FROM Periodo, PeriodosFilial WHERE FilialEmpresa = ? AND Periodo.DataInicio<=? AND Periodo.DataFim>=? AND Periodo.Exercicio = PeriodosFilial.Exercicio AND Periodo.Periodo = PeriodosFilial.Periodo", tPeriodo.iExercicio, tPeriodo.iPeriodo, tPeriodo.sNomeExterno, tPeriodo.iLancamentos, tPeriodo.iApurado, tPeriodo.dtDataInicio, tPeriodo.dtDataFim, tPeriodo.iLote, tPeriodo.iFechado, iFilialEmpresa, dtData, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 83473
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83474
    
    If lErro = AD_SQL_SEM_DADOS Then gError 83475
    
    Call Comando_Fechar(lComando)
    
    objPeriodo.iExercicio = tPeriodo.iExercicio
    objPeriodo.iPeriodo = tPeriodo.iPeriodo
    objPeriodo.sNomeExterno = tPeriodo.sNomeExterno
    objPeriodo.iLancamentos = tPeriodo.iLancamentos
    objPeriodo.iApurado = tPeriodo.iApurado
    objPeriodo.dtDataInicio = tPeriodo.dtDataInicio
    objPeriodo.dtDataFim = tPeriodo.dtDataFim
    objPeriodo.iLote = tPeriodo.iLote
    objPeriodo.iFechado = tPeriodo.iFechado
    
    Periodo_Le2 = SUCESSO

    Exit Function

Erro_Periodo_Le2:

    Periodo_Le2 = gErr
    
    Select Case gErr
    
    Case 83472
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
    Case 83473, 83474
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO", gErr)
    
    Case 83475
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147788)
        
    End Select
    
    Call Comando_Fechar(lComando)
       
    Exit Function
    
End Function

Function Conta_VaziaFormatoBD() As String
'retorna conta vazia no formato do bd

Dim iInd1 As Integer
Dim iNumSegmentos As Integer
Dim objSegmento As New ClassSegmento
Dim sSegmentoFormatado As String 'Segmento sendo formatado
Dim sContaFormatada As String
Dim lErro As Long

    sContaFormatada = ""

    iNumSegmentos = NumSegConta()
    
    For iInd1 = 1 To iNumSegmentos
        
        lErro = SegmentoConta(iInd1, objSegmento)
        If lErro <> SUCESSO Then Error 6072
        
        sSegmentoFormatado = ""

        Select Case objSegmento.iTipo
        
            Case SEGMENTO_NUMERICO
                            
                sSegmentoFormatado = String(objSegmento.iTamanho, "0")
                
            Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII
        
                sSegmentoFormatado = String(objSegmento.iTamanho, " ")

        End Select

        sContaFormatada = sContaFormatada & sSegmentoFormatado

    Next
    
    Conta_VaziaFormatoBD = sContaFormatada

End Function

Function Conta_Formata_Importacao(ByVal sConta As String, sContaFormatada As String, iContaPreenchida As Integer) As Long
'Devolve Formato BD: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
'A String de Entrada vem do GRID com os caracteres delimitadores e brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'iContaPreenchida retornará CONTA_VAZIA se nenhum segmento da conta for preenchido,
'caso contrário retornará CONTA_PREENCHIDA
    
Dim lErro As Long
Dim iNumSegmentos As Integer
Dim iInd1 As Integer
Dim iInd2 As Integer
Dim objSegmento As New ClassSegmento
Dim sContaBruta As String
Dim sSegmentoBruto As String 'Segmento sendo formatado
Dim sSegmentoFormatado As String 'Segmento sendo formatado
Dim sCaracter As String 'caracter sendo tratado
Dim sCcl As String
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim iSegmentoVazio As Integer
Dim iRestoVazio As Integer
Dim iIndice As Integer

On Error GoTo Erro_Conta_Formata_Importacao
            
    iContaPreenchida = CONTA_VAZIA
    sContaFormatada = ""

    'se a conta veio com os delimitadores
    If InStr(1, sConta, ".") <> 0 Then
    
        '??? copiado de Conta_Formata
        
        iRestoVazio = SEGMENTOS_RESTANTES_CHEIOS
        sContaBruta = sConta
        iNumSegmentos = NumSegConta()
        
        For iInd1 = 1 To iNumSegmentos
            
            lErro = SegmentoConta(iInd1, objSegmento)
            If lErro <> SUCESSO Then Error 6072
            
            sSegmentoBruto = left(sContaBruta, objSegmento.iTamanho)
            sContaBruta = Mid(sContaBruta, objSegmento.iTamanho + Len(objSegmento.sDelimitador) + 1)
            sSegmentoFormatado = ""
    
            iSegmentoVazio = SEGMENTO_VAZIO
            
            Select Case objSegmento.iTipo
            
                Case SEGMENTO_NUMERICO
            
                    For iInd2 = 1 To objSegmento.iTamanho
            
                        sCaracter = Mid(sSegmentoBruto, iInd2, 1)
            
                        If sCaracter = " " Or sCaracter = "" Then
                            sSegmentoFormatado = "0" & sSegmentoFormatado
                            
                        ElseIf sCaracter = "0" Then
                            sSegmentoFormatado = sSegmentoFormatado & sCaracter
                            
                        ElseIf sCaracter = "-" Then
                            Error 55689
                        Else
                            iSegmentoVazio = SEGMENTO_CHEIO
                            sSegmentoFormatado = sSegmentoFormatado & sCaracter
                        End If
            
                    Next
                    
                    
                Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII
            
                    For iInd2 = objSegmento.iTamanho To 1 Step -1
            
                        sCaracter = Mid(sSegmentoBruto, iInd2, 1)
            
                        If sCaracter = " " Or sCaracter = "" Then
                            sSegmentoFormatado = sSegmentoFormatado & " "
                        Else
                            iSegmentoVazio = SEGMENTO_CHEIO
                            sSegmentoFormatado = sCaracter & sSegmentoFormatado
                        End If
            
                    Next
    
            End Select
    
            If iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS And iSegmentoVazio <> SEGMENTO_VAZIO Then Error 5954
                
            If iSegmentoVazio = SEGMENTO_VAZIO Then
                iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS
            Else
                iContaPreenchida = CONTA_PREENCHIDA
            End If
    
            sContaFormatada = sContaFormatada & sSegmentoFormatado
    
        Next
        
        If giSetupUsoCcl = CCL_USA_CONTABIL Then
        
            Error 5954 '??? nao tratado
           
        End If
    
    Else
    
        If Len(Trim(sConta)) <> 0 Then
        
            sContaFormatada = Trim(sConta) & Mid(Conta_VaziaFormatoBD, Len(Trim(sConta)) + 1)
            iContaPreenchida = CONTA_PREENCHIDA
            
        End If
    
    End If
    
    If iContaPreenchida = CONTA_VAZIA Then sContaFormatada = ""
    
    Conta_Formata_Importacao = SUCESSO
    
    Exit Function
    
Erro_Conta_Formata_Importacao:

    Conta_Formata_Importacao = Err

    Select Case Err
    
        Case 5819, 5820, 6072 'Vazio porque o erro já foi tratado
        
        Case 5954
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_SEG_MEIO_NAO_PREENCHIDOS", Err)
        
        Case 55689
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_SEG_NUM_CARACTER_INVALIDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150477)

    End Select

    Exit Function
    
End Function

Function Ccl_VazioFormatoBD() As String
'retorna ccl vazio no formato do bd

Dim iInd1 As Integer
Dim iNumSegmentos As Integer
Dim objSegmento As New ClassSegmento
Dim sSegmentoFormatado As String 'Segmento sendo formatado
Dim sCclFormatada As String
Dim lErro As Long

    sCclFormatada = ""
    
    iNumSegmentos = NumSegCcl()
    
    For iInd1 = 1 To iNumSegmentos
        
        lErro = SegmentoCcl(iInd1, objSegmento)
        If lErro <> SUCESSO Then Error 6072
        
        sSegmentoFormatado = ""

        Select Case objSegmento.iTipo
        
            Case SEGMENTO_NUMERICO
                            
                sSegmentoFormatado = String(objSegmento.iTamanho, "0")
                
            Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII
        
                sSegmentoFormatado = String(objSegmento.iTamanho, " ")

        End Select

        sCclFormatada = sCclFormatada & sSegmentoFormatado

    Next
    
    Ccl_VazioFormatoBD = sCclFormatada

End Function

Function Ccl_Formata_Importacao(ByVal sCcl As String, sCclFormatada As String, iCclPreenchida As Integer) As Long
'Devolve Formato BD: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
'A String de Entrada vem do GRID com brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'iCclPreenchida retornará CCL_VAZIA se nenhum segmento do centro de custo for preenchido,
'caso contrário retornará CCL_PREENCHIDA

Dim lErro As Long
Dim iNumSegmentos As Integer
Dim iInd1 As Integer
Dim iInd2 As Integer
Dim objSegmento As New ClassSegmento
Dim sCclBruta As String
Dim sSegmentoBruto As String 'Segmento sendo formatado
Dim sSegmentoFormatado As String 'Segmento sendo formatado
Dim sCaracter As String 'caracter sendo tratado
Dim iSegmentoVazio As Integer
Dim iRestoVazio As Integer
    
On Error GoTo Erro_Ccl_Formata_Importacao
            
    iCclPreenchida = CCL_VAZIA
    iRestoVazio = SEGMENTOS_RESTANTES_CHEIOS
    sCclBruta = sCcl
    sCclFormatada = ""
    
    iNumSegmentos = NumSegCcl()
    
    If InStr(1, sCcl, ".") <> 0 Then
    
        iRestoVazio = SEGMENTOS_RESTANTES_CHEIOS
        sCclBruta = sCcl
    
        For iInd1 = 1 To iNumSegmentos
            
            lErro = SegmentoCcl(iInd1, objSegmento)
            If lErro <> SUCESSO Then Error 6122
            
            sSegmentoBruto = left(sCclBruta, objSegmento.iTamanho)
            sCclBruta = Mid(sCclBruta, objSegmento.iTamanho + Len(objSegmento.sDelimitador) + 1)
            sSegmentoFormatado = ""
    
            iSegmentoVazio = SEGMENTO_VAZIO
    
            Select Case objSegmento.iTipo
            Case SEGMENTO_NUMERICO
    
                For iInd2 = 1 To objSegmento.iTamanho
    
                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
    
                    If sCaracter = " " Or sCaracter = "" Then
                        sSegmentoFormatado = "0" & sSegmentoFormatado
                        
                    ElseIf sCaracter = "0" Then
                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
                    ElseIf sCaracter = "-" Then
                        Error 55690
                    Else
                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
                        iSegmentoVazio = SEGMENTO_CHEIO
                    End If
    
                Next
    
            Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII
    
                For iInd2 = objSegmento.iTamanho To 1 Step -1
    
                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
    
                    If sCaracter = " " Or sCaracter = "" Then
                        sSegmentoFormatado = sSegmentoFormatado & " "
                    Else
                        sSegmentoFormatado = sCaracter & sSegmentoFormatado
                        iSegmentoVazio = SEGMENTO_CHEIO
                    End If
    
                Next
    
            End Select
    
            If iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS And iSegmentoVazio <> SEGMENTO_VAZIO Then Error 5955
                
            If iSegmentoVazio = SEGMENTO_VAZIO Then
                iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS
            Else
                iCclPreenchida = CCL_PREENCHIDA
            End If
    
            sCclFormatada = sCclFormatada & sSegmentoFormatado
    
        Next

    Else
    
        If Len(Trim(sCcl)) <> 0 Then
        
            If Len(Trim(sCcl)) >= Len(Ccl_VazioFormatoBD) Then
                
                sCclFormatada = sCcl
                
            Else
            
                If iNumSegmentos = 1 Then
                
                    lErro = SegmentoCcl(1, objSegmento)
                    If lErro <> SUCESSO Then Error 6122
                
                    If objSegmento.iTipo = SEGMENTO_NUMERICO Then
                        sCclFormatada = String(Len(Ccl_VazioFormatoBD) - Len(Trim(sCcl)), "0") & Trim(sCcl)
                    Else
                        sCclFormatada = Trim(sCcl) & String(Len(Ccl_VazioFormatoBD) - Len(Trim(sCcl)), " ")
                    End If
                    
                Else
                
                    sCclFormatada = Trim(sCcl) & Mid(Ccl_VazioFormatoBD, Len(Trim(sCcl)))
                
                End If
        
            End If
        
            iCclPreenchida = CCL_PREENCHIDA
        
        End If
    
    End If
    
    If iCclPreenchida = CCL_VAZIA Then sCclFormatada = ""

    Ccl_Formata_Importacao = SUCESSO
    
    Exit Function
    
Erro_Ccl_Formata_Importacao:

    Ccl_Formata_Importacao = Err

    Select Case Err

        Case 5955
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_SEG_MEIO_NAO_PREENCHIDOS", Err)

        Case 6122 'Vazio porque o erro já foi tratado
        
        Case 55690
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_SEG_NUM_CARACTER_INVALIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150475)
        

    End Select

    Exit Function
    
End Function

''Function Mascara_RetornaCclPai(sCcl As String, sCclPai As String) As Long
''
''    sCclPai = ""
''    Mascara_RetornaCclPai = SUCESSO
''
''End Function
