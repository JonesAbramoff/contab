VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotGeraComiInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Conexao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Private Declare Function Conexao_FecharExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long


Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Private glConexaoTVI As Long

Public gobjTRVVoucher As ClassTRVVouchers
Public gobjTRVProdTafCalc As ClassTRVProdTafCalc
Public gdValorBase As Double
Public gdAporteValor As Variant
Public giTemCMA As Integer
Public giTemCMCC As Integer
Public giTipoCliente As Variant
Public giAporteMoeda As Variant
Public giVendedorCallCenter As Integer

'inicia objeto associado a GridComissoes
Public objTabComissoes As New ClassTabComissoes

'declara o objeto da execucao das regras de comissoes e
'o objeto do calculo de comissoes
Public objComissoesRegrasCalc As Object
Public objMnemonicoComissCalc As ClassMnemonicoComissCalc
Public objMnemonicoComissCalcAux As ClassMneComissCalcTRV

Public Function TRVGerComiInt_Gera(objTRVGerComiInt As ClassTRVGerComiInt, ByVal iRelatorio As Integer) As Long
'Gera as comissões internas

Dim lErro As Long
Dim vNumReg As Variant
Dim alComando(1 To 10) As Long
Dim vNumVou As Variant
Dim vTipoDoc As Variant
Dim vSerie As Variant
Dim tTRVVoucher As typeTRVVouchers
Dim tTRVVoucherInfo As typeTRVVoucherInfoN
Dim dValorBase As Double
Dim vCliente As Variant
Dim colComissoes As New Collection
Dim bCalculou As Boolean
Dim lNumIntTRVGerComiInt As Long
Dim dtDataGeracao As Date
Dim dHoraGeracao As Double
Dim lSeq As Long
Dim dValorDolarAcum As Double
Dim iVendedor As Integer
Dim iRegiao As Integer
Dim colTRVAcumVendedor As New Collection
Dim colTRVAcumSupervisor As New Collection
Dim colTRVAcumGerente As New Collection
Dim colTRVAcumDiretor As New Collection
Dim dValorDolarBase As Double
Dim iIndice As Integer
Dim sSelecao As String
Dim sComando_SQL As String
Dim lTransacao As Long
Dim colNumIntTRVVoucherInfo As Collection
Dim lNumIntTRVGerComiIntDet As Long
Dim objTRVAcumAporte As ClassTRVAcumAporte
Dim dValorDe As Double
Dim dValorAte As Double
Dim dAporteValorAux As Double
Dim objTRVAcumVendedor As ClassTRVAcumVendedor
Dim iAchouVendedor As Integer
Dim vsTipoDoc As Variant
Dim viTipoDocDestino As Variant
Dim vdtDataEmiDe As Variant
Dim vdtDataEmiAte As Variant
Dim vdtDataPagDe As Variant
Dim vdtDataPagAte As Variant
Dim vsTipVou As Variant
Dim vlNumVou As Variant
Dim vsSerie As Variant
Dim vlCliente As Variant
Dim vdCambio As Variant
Dim vsProduto As Variant
Dim vidiasantc As Variant
Dim objComissaoGenerica As Object
Dim lNumIntInfoGer As Long
Dim lNumIntRel As Long
Dim dValorVoucherInfo As Double
Dim lNumIntRelComiInt As Long
Dim objTRVVoucherInfo As ClassTRVVoucherInfoN
Dim objRelatorio As New AdmRelatorio
Dim vlPromotor As Variant
Dim vlEmpresaPai As Variant
Dim viAporte As Variant
Dim vdtData As Variant
Dim lEmpresaPai As Long
Dim objEmpresa As New ClassDicEmpresa
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String
Dim iLenParamOut As Integer
Dim alComandoTVI(1 To 10) As Long
Dim alComandoTVA(1 To 10) As Long
Dim iEmpresa As Integer
Dim vdtDataEmiAux As Date
Dim vdtDataPagAux As Date
Dim vdTarifaUnitaria As Variant
Dim vdPercComiProd As Variant
Dim viVendProd As Variant
Dim viVendCli As Variant
Dim viVigencia As Variant
Dim colTarifas As New Collection
Dim objTRVProdTafCalc As ClassTRVProdTafCalc
Dim bAchou As Boolean
Dim dValorTarif As Double
Dim vdTarifaUNNormal As Variant
Dim objVendedor As New ClassVendedor
Dim objFilialCliCategoria As New ClassFilialCliCategoria
Dim iCallCenter As Integer
Dim iRecursivo As Integer
Dim vdtDataFat As Date
Dim vsNomeRedCli As String
Dim vsNomeRedProd As String
Dim vlNumFatCoinfo As Long
Dim vsMoeda As String
Dim vdValorCambio As Double
Dim vdValorBruto As Double
Dim objPlanilha As New ClassPlanilhaExcel
Dim bCliEhCallCenter As Boolean
Dim vsUsuRespCallCenter As String

On Error GoTo Erro_TRVGerComiInt_Gera


    TelaAcompanhaBatchFAT1.dValorTotal = 0
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 137413

    'desativa os locks dos comandos a seguir
    Call Conexao_DesativarLocks(DESATIVAR_LOCKS)

    If iRelatorio <> 0 Then
        
        lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_PROX_NUMINTREL", lNumIntRel)
        If lErro <> SUCESSO Then gError 197469
        
        TelaAcompanhaBatchFAT1.Observacao.Caption = "Prévia n. " & CStr(lNumIntRel)
        
        If iRelatorio = 3 Then
        
            lErro = Exporta_Cabecalho_Excel(objPlanilha)
            If lErro <> SUCESSO Then gError 199589
        
        End If
        
        
    Else
        
        lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVGERCOMIINT", lNumIntTRVGerComiInt)
        If lErro <> SUCESSO Then gError 197886
        
    End If

    objVendedor.sNomeReduzido = "CallCenter"
    lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
    If lErro <> SUCESSO And lErro <> 25008 Then gError 199397
    
    If lErro = SUCESSO Then
        giVendedorCallCenter = objVendedor.iCodigo
    End If

    'instancia a classe q executa as regras
    Set objComissoesRegrasCalc = CreateObject("RotinasContab.ClassComissoesRegrasCalc")
    
    'instancia a classe q calcula os mnemonicos
    Set objMnemonicoComissCalc = New ClassMnemonicoComissCalc
    Set objMnemonicoComissCalcAux = New ClassMneComissCalcTRV
    

    'setar o objtela para as classes que calculam mnemonicos
    Set objMnemonicoComissCalc.objTela = Me
    Set objMnemonicoComissCalcAux.objTela = Me

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197382
        alComandoTVA(iIndice) = alComando(iIndice)
    Next

    If glEmpresa <> 1 Then gError 197889

    objEmpresa.lCodigo = 2 'Travel Ace

    lErro = Empresa_Le(objEmpresa)
    If lErro <> SUCESSO Then gError 189756

    sDSN = objEmpresa.sStringConexao
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    glConexaoTVI = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If glConexaoTVI = 0 Then gError 197890

    For iIndice = LBound(alComandoTVI) To UBound(alComandoTVI)
        alComandoTVI(iIndice) = Comando_AbrirExt(glConexaoTVI)
        If alComandoTVI(iIndice) = 0 Then gError 197891
    Next
   
    lErro = CF("TRVProdTarifa_Le", colTarifas)
    If lErro <> SUCESSO Then gError 197469

'    sSelecao = "WHERE TRVVouchers.TipoDoc = ? AND TRVVouchers.TipoDocDestino = ? AND TRVVouchers.NumIntDocDestino = ParcelasRecBaixadas.NumIntTitulo AND BaixasParcRec.NumIntParcela = ParcelasRecBaixadas.NumINtDoc AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND "
'    sSelecao = sSelecao & "ParcelasRecBaixadas.NumParcela = 1 AND BaixasParcRec.Sequencial = 1 AND Clientes.Codigo = TRVVouchers.ClienteVou AND TRVVouchers.ClienteVou = FiliaisClientesTRV.CodCliente AND FiliaisClientesTRV.CodFilial = 1 AND Clientes.Tipo <> 7 AND EXISTS (SELECT * FROM TRVVoucherInfo WHERE NumIntDocComiInt = 0 AND TRVVoucherInfo.NumVou = TRVVouchers.NumVou AND TRVVoucherInfo.Serie = TRVVouchers.Serie)"

    sSelecao = "WHERE V.TipoDoc = ? AND V.TipoDocDestino = ? AND V.NumIntDocDestino = B.NumIntTitulo AND V.Produto = P.Codigo AND V.Produto = PT.Codigo AND V.ClienteVou = FC.CodCliente AND FC.CodFilial = 1 "
    sSelecao = sSelecao & "AND C.Codigo = V.ClienteVou AND V.ClienteVou = FT.CodCliente AND FT.CodFilial = 1 AND EXISTS (SELECT I.NumIntDoc FROM TRVVoucherInfo AS I WHERE I.NumIntDocComiInt = 0 AND I.NumVou = V.NumVou AND I.Serie = V.Serie AND I.TipVou = V.TipVou) "
    sSelecao = sSelecao & " AND NOT (V.Status = 7 AND V.Data < '09-15-2008') "

    If objTRVGerComiInt.dtDataEmiDe <> DATA_NULA Then
        sSelecao = sSelecao & " AND V.Data >= ? "
    End If
    
    If objTRVGerComiInt.dtDataEmiAte <> DATA_NULA Then
        sSelecao = sSelecao & " AND ( V.Data <= ? "
    End If

    If objTRVGerComiInt.dtDataPagtoDe <> DATA_NULA Then
        sSelecao = sSelecao & " AND B.DataBaixa >= ? "
    End If
    
    If objTRVGerComiInt.dtDataPagtoAte <> DATA_NULA Then
        sSelecao = sSelecao & " AND B.DataBaixa <= ? "
    End If
    
    If objTRVGerComiInt.dtDataEmiAte <> DATA_NULA Then
        
        vdtDataEmiAux = StrParaDate("01/" & CStr(Month(objTRVGerComiInt.dtDataEmiAte)) & "/" & CStr(Year(objTRVGerComiInt.dtDataEmiAte)))
        vdtDataPagAux = objTRVGerComiInt.dtDataPagtoDe
    
        sSelecao = sSelecao & " OR ( V.Data >= ? AND B.DataBaixa < ? ))"
    End If

    sComando_SQL = "SELECT Count(*) "
    sComando_SQL = sComando_SQL & "FROM PrimeiraBaixaTitRec AS B, ClientesTVA AS C, FiliaisClientesTRVTVA AS FT, TRVVouchers AS V, ProdutosTVA AS P, FiliaisClientesTVA AS FC, TRVProdutosTVA PT " & sSelecao

    'sComando_SQL = "SELECT Count(*) FROM TRVVouchers, ParcelasRecBaixadas, BaixasParcRec, BaixasRec, Clientes, FiliaisClientesTRV " & sSelecao

    vsTipoDoc = TRV_TIPODOC_VOU_TEXTO
    viTipoDocDestino = TRV_TIPO_DOC_DESTINO_TITREC
    vdtDataEmiDe = objTRVGerComiInt.dtDataEmiDe
    vdtDataEmiAte = objTRVGerComiInt.dtDataEmiAte
    vdtDataPagDe = objTRVGerComiInt.dtDataPagtoDe
    vdtDataPagAte = objTRVGerComiInt.dtDataPagtoAte

    For iEmpresa = 1 To 2
    
        lErro = Comando_PrepararInt(alComando(1), sComando_SQL)
        If lErro <> AD_SQL_SUCESSO Then gError 197383
    
        vNumReg = CLng(0)
        lErro = TRVGerComiInt_Bind_2(alComando(1), vNumReg)
        If lErro <> AD_SQL_SUCESSO Then gError 197385
        
        
        lErro = TRVGerComiInt_Bind_1(objTRVGerComiInt, alComando(1), vsTipoDoc, viTipoDocDestino, vdtDataEmiDe, vdtDataEmiAte, vdtDataPagDe, vdtDataPagAte, vdtDataEmiAux, vdtDataPagAux)
        If lErro <> AD_SQL_SUCESSO Then gError 197385
    
        'Pesquisa os lotes de inventário pendentes no banco de dados
        lErro = Comando_ExecutarInt(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 197386
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197387
    
        'Inicializa barra de progressão
        TelaAcompanhaBatchFAT1.dValorTotal = TelaAcompanhaBatchFAT1.dValorTotal + vNumReg
    
        alComando(1) = alComandoTVI(1)
    
    Next
    
    alComando(1) = alComandoTVA(1)
    
    TelaAcompanhaBatchFAT1.dValorAtual = 0
      
    dtDataGeracao = gdtDataAtual
    dHoraGeracao = Time

    If TelaAcompanhaBatchFAT1.dValorTotal > 0 Then
      
        lSeq = 1
      
        For iEmpresa = 1 To 2
      
            vsTipVou = String(STRING_TRV_VOU_INFO_TIPVOU, 0)
            vlNumVou = CLng(0)
            vsSerie = String(STRING_TRV_VOU_INFO_SERIE, 0)
            vlCliente = CLng(0)
            vdCambio = CDbl(0)
            gdAporteValor = CDbl(0)
            giTipoCliente = CInt(0)
            vsProduto = String(STRING_PRODUTO, 0)
            vidiasantc = CInt(0)
            giAporteMoeda = CInt(0)
            vlPromotor = CLng(0)
            vlEmpresaPai = CLng(0)
            viAporte = CInt(0)
            vdtData = CDate(DATA_NULA)
            vdTarifaUnitaria = CDbl(0)
            vdPercComiProd = CDbl(0)
            viVendProd = CInt(0)
            viVendCli = CInt(0)
            viVigencia = CInt(0)
            vdTarifaUNNormal = CDbl(0)
            vdtDataFat = CDate(DATA_NULA)
            vsNomeRedCli = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
            vsNomeRedProd = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
            vlNumFatCoinfo = CLng(0)
            vsMoeda = String(STRING_SIMBOLO_MOEDA, 0)
            vdValorCambio = CDbl(0)
            vdValorBruto = CDbl(0)
            vsUsuRespCallCenter = String(255, 0)
          
            lErro = TRVVoucher_Le_GerComiInt(alComando(2), sSelecao, objTRVGerComiInt, vsTipVou, vlNumVou, vsSerie, vlCliente, vdCambio, gdAporteValor, giTipoCliente, vsProduto, vidiasantc, giAporteMoeda, vlPromotor, vlEmpresaPai, viAporte, vdtData, vdTarifaUnitaria, vdPercComiProd, viVendProd, viVendCli, viVigencia, vdTarifaUNNormal, vdtDataFat, vsNomeRedCli, vsNomeRedProd, vlNumFatCoinfo, vsMoeda, vdValorCambio, vdValorBruto, vsUsuRespCallCenter)
            If lErro <> SUCESSO Then gError 197389
        
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197390
              
            Do While lErro = AD_SQL_SUCESSO
            
                gdValorBase = 0
                giTemCMA = 0
                giTemCMCC = 0
            
                Set gobjTRVVoucher = New ClassTRVVouchers
                gobjTRVVoucher.lCliente = vlCliente
                gobjTRVVoucher.sProduto = vsProduto
                gobjTRVVoucher.idiasantc = vidiasantc
                gobjTRVVoucher.sTipVou = vsTipVou
                gobjTRVVoucher.sSerie = vsSerie
                gobjTRVVoucher.lNumVou = vlNumVou
'                gobjTRVVoucher.iAporteMoeda = giAporteMoeda
                gobjTRVVoucher.iTipoCliente = giTipoCliente
                gobjTRVVoucher.lPromotor = vlPromotor
                gobjTRVVoucher.dtData = vdtData
                gobjTRVVoucher.dCambio = vdCambio
                gobjTRVVoucher.dTarifaUnitaria = vdTarifaUnitaria
                gobjTRVVoucher.iVendCli = viVendCli
                gobjTRVVoucher.iVendProd = viVendProd
                gobjTRVVoucher.dPercComissaoProd = vdPercComiProd
                gobjTRVVoucher.iVigencia = viVigencia
                gobjTRVVoucher.dtDataFat = vdtDataFat
                gobjTRVVoucher.sNomeRedCli = vsNomeRedCli
                gobjTRVVoucher.sNomeRedProd = vsNomeRedProd
                gobjTRVVoucher.lNumFatCoinfo = vlNumFatCoinfo
                gobjTRVVoucher.sMoeda = vsMoeda
                gobjTRVVoucher.dValorCambio = vdValorCambio
                gobjTRVVoucher.dValorBruto = vdValorBruto
                gobjTRVVoucher.sUsuRespCallCenter = vsUsuRespCallCenter
                
'                If gobjTRVVoucher.lNumVou = 309519 Or gobjTRVVoucher.lNumVou = 309505 Then
'                    MsgBox (gobjTRVVoucher.lNumVou)
'                End If
'
'                If gobjTRVVoucher.lNumVou = 309530 Or gobjTRVVoucher.lNumVou = 319360 Then
'                    MsgBox (gobjTRVVoucher.lNumVou)
'                End If
                
                bAchou = False
                For Each objTRVProdTafCalc In colTarifas
                    If UCase(objTRVProdTafCalc.sProduto) = UCase(gobjTRVVoucher.sProduto) Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If bAchou Then
                    Set gobjTRVProdTafCalc = objTRVProdTafCalc

                    Call gobjTRVProdTafCalc.Calcula_Tarifa(Year(gobjTRVVoucher.dtData), gobjTRVVoucher.iVigencia, dValorTarif)
                    gobjTRVVoucher.dTarifaUNNormal = dValorTarif
                Else
                    Set gobjTRVProdTafCalc = New ClassTRVProdTafCalc
                    gobjTRVVoucher.dTarifaUNNormal = gobjTRVVoucher.dTarifaUnitaria
                End If

                If vdTarifaUNNormal <> 0 Then
                    gobjTRVVoucher.dTarifaUNNormal = vdTarifaUNNormal
                End If
                
                lEmpresaPai = vlEmpresaPai
                
                If viAporte = 0 Then
            
                    gobjTRVVoucher.dAporteValor = 0
            
                Else
            
                    lErro = TRVVoucher_Le_Aportes(alComando(8), gobjTRVVoucher, lEmpresaPai)
                    If lErro <> SUCESSO Then gError 197883
                    
                End If
            
                gdAporteValor = gobjTRVVoucher.dAporteValor
            
                Set colNumIntTRVVoucherInfo = New Collection
            
                tTRVVoucherInfo.sTipoDoc = String(STRING_TRV_VOU_INFO_TIPODOC, 0)
        
                sComando_SQL = "SELECT NumIntDoc, TipoDoc, Valor FROM TRVVoucherInfo WHERE NumIntDocComiInt = 0 AND NumVou = ? AND Serie = ? AND TipVou = ?"
            
                lErro = Comando_ExecutarPos(alComando(3), sComando_SQL, 0, tTRVVoucherInfo.lNumIntDoc, tTRVVoucherInfo.sTipoDoc, tTRVVoucherInfo.dValor, vlNumVou, vsSerie, vsTipVou)
                If lErro <> AD_SQL_SUCESSO Then gError 197391
              
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197392
              
                If lErro = AD_SQL_SUCESSO Then
                    
                    If iRelatorio = 2 Then

                        lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_PROX_TRVRELCOMIINT", lNumIntRelComiInt)
                        If lErro <> SUCESSO Then gError 197474

                    End If
            
                End If
                
                Do While lErro = AD_SQL_SUCESSO
                
                    Select Case tTRVVoucherInfo.sTipoDoc
                    
                        Case TRV_TIPODOC_BRUTO_TEXTO, TRV_TIPODOC_OCR_TEXTO, TRV_TIPODOC_NVL_TEXTO
                            gdValorBase = gdValorBase + tTRVVoucherInfo.dValor
                            dValorVoucherInfo = tTRVVoucherInfo.dValor
                            
                        Case TRV_TIPODOC_CMA_TEXTO, TRV_TIPODOC_CMC_TEXTO, TRV_TIPODOC_CMR_TEXTO, TRV_TIPODOC_OVER_TEXTO, TRV_TIPODOC_CMCC_TEXTO
                            gdValorBase = gdValorBase - tTRVVoucherInfo.dValor
                            dValorVoucherInfo = -tTRVVoucherInfo.dValor
                            'se tiver qq tipo de desconto ==> giTemCMA = 1
                            giTemCMA = 1
                            If tTRVVoucherInfo.sTipoDoc = TRV_TIPODOC_CMCC_TEXTO Then giTemCMCC = 1
                    End Select
                    
                    If iRelatorio = 0 Then
                    
                        lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRVVoucherInfo SET NumIntDocComiInt = ?", alComando(3), lNumIntTRVGerComiInt)
                        If lErro <> AD_SQL_SUCESSO Then gError 197391
                    
                    ElseIf iRelatorio = 2 Then
                
                        lErro = Comando_Executar(alComando(7), "INSERT INTO TRVRelVoucherInfo (NumIntRel, NumIntDoc, TipoDoc, Valor, NumIntRelComiInt, Serie, NumVou, Empresa) VALUES (?,?,?,?,?,?,?,?)", _
                        lNumIntRel, tTRVVoucherInfo.lNumIntDoc, tTRVVoucherInfo.sTipoDoc, dValorVoucherInfo, lNumIntRelComiInt, gobjTRVVoucher.sSerie, gobjTRVVoucher.lNumVou, iEmpresa)
                        If lErro <> AD_SQL_SUCESSO Then gError 197473
                
                    End If
                    
                    Set objTRVVoucherInfo = New ClassTRVVoucherInfoN
                    
                    objTRVVoucherInfo.sTipoDoc = tTRVVoucherInfo.sTipoDoc
                    objTRVVoucherInfo.lNumIntDoc = tTRVVoucherInfo.lNumIntDoc
                    objTRVVoucherInfo.dValor = dValorVoucherInfo
                    
                    
                    colNumIntTRVVoucherInfo.Add objTRVVoucherInfo
                    
                    lErro = Comando_BuscarProximo(alComando(3))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197393
                
                Loop
               
                'antes dava erro mas agora tem q aceitar pois tem produtos em reais que vem com este campo zerado
                If vdCambio = 0 Then vdCambio = 1 ' gError 197468
              
                dValorDolarBase = gdValorBase / vdCambio
                
'                objFilialCliCategoria.lCliente = gobjTRVVoucher.lCliente
'                objFilialCliCategoria.iFilial = 1
'                objFilialCliCategoria.sCategoria = "Responsável"
'                objFilialCliCategoria.sItem = "Call Center"
'
                iCallCenter = 0
                bCliEhCallCenter = False
'
'                If gobjTRVVoucher.dtData >= #7/1/2008# Then
'
'                lErro = CF("FilialClienteCategorias_Le", objFilialCliCategoria)
'                If lErro <> SUCESSO And lErro <> 199401 Then gError 199403

                lErro = CF("TRVVoucher_Verifica_CallCenter", gobjTRVVoucher.lCliente, gobjTRVVoucher.dtData, bCliEhCallCenter)
                If lErro <> SUCESSO Then gError 199403

'                If lErro = SUCESSO Then
                If bCliEhCallCenter Then
                    iCallCenter = 1
                    gobjTRVVoucher.iAporte = 2
                End If

'                End If
              

                bCalculou = False
              
                'se nao for venda direta
                If 1 = 1 Then
                
                    '*****************************************************************
                    'se for cliente > 1.000.000 significa que é cliente MyAssistance
                    '*****************************************************************
                    If gobjTRVVoucher.lCliente > 1000000 Then
                
                        Set colComissoes = New Collection
                        
                        lErro = objComissoesRegrasCalc.Comissoes_Calcula_Regras_Int(Me, bCalculou, colComissoes)
                        If lErro <> SUCESSO Then gError 199431
                        
                        If bCalculou Then
                        
                            Set objComissaoGenerica = colComissoes(1)
                            
                            iVendedor = objComissaoGenerica.iCodVendedor
                            
                            lErro = TRVGerComiIntDet_Acum_Comissoes(alComando(10), gobjTRVVoucher, iVendedor, gdValorBase, dValorDolarBase, gdAporteValor, giAporteMoeda, vdCambio, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, iRelatorio, lNumIntRel, iCallCenter)
                            If lErro <> SUCESSO Then gError 199858
                            
                            If iRelatorio = 0 Then
                            
                                lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOUCHERINFOGER", lNumIntInfoGer)
                                If lErro <> SUCESSO Then gError 199859
                                
                            End If
                            
                            iRecursivo = 0
                            
                            If iRelatorio = 0 Then
                            
                                For Each objTRVAcumVendedor In colTRVAcumVendedor
                            
                                    If objTRVAcumVendedor.iVendedor = iVendedor Then
                                    
                                        For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
                                        
                                            lErro = Comando_Executar(alComando(9), "INSERT INTO TRVVoucherInfoGer (NumIntVoucherInfo, NumIntInfoGer, Empresa) VALUES (?,?,?)", _
                                            objTRVVoucherInfo.lNumIntDoc, objTRVAcumVendedor.lNumIntInfoGerGrava, iEmpresa)
                                            If lErro <> AD_SQL_SUCESSO Then gError 199860
                            
                                        Next
                            
                                        Exit For
                                        
                                    End If
                            
                                Next
                            
                            End If
                            
                            
                            lErro = TRVGerComiIntDet_Gravar_Comissoes(alComando(10), iEmpresa, colComissoes, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, colNumIntTRVVoucherInfo, lNumIntTRVGerComiIntDet, lNumIntInfoGer, iRelatorio, lNumIntRel, gobjTRVVoucher, lNumIntRelComiInt, 1, objPlanilha, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, iCallCenter)
                            If lErro <> SUCESSO Then gError 199861
                            
                            If iRelatorio = 2 Then
                                                    
                                iRecursivo = -1
                                                    
                                'vai gravar o TRVRelVoucherInfo associado ao promotor
                                lErro = TRVRelVoucher_Gravar_1(VENDEDOR_CARGO_TRV_PROMOTOR, iRecursivo, iEmpresa, iVendedor, colNumIntTRVVoucherInfo, iRelatorio, lNumIntRel, gobjTRVVoucher, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, colComissoes(1).dValorBase)
                                If lErro <> SUCESSO Then gError 199862
                            
                            End If
                            
                        End If
              
                    Else
              
                        If iCallCenter = 0 Then
                  
                            Set colComissoes = New Collection
                      
                            lErro = objComissoesRegrasCalc.Comissoes_Calcula_Regras_Int(Me, bCalculou, colComissoes)
                            If lErro <> SUCESSO Then gError 197394
                          
                            If bCalculou Then
                            
                                Set objComissaoGenerica = colComissoes(1)
                                
                                
                                
                                
                                objComissaoGenerica.dValor = Arredonda_Moeda(objComissaoGenerica.dValor)
                            
                                iVendedor = objComissaoGenerica.iCodVendedor
                            
                                If iRelatorio = 0 Then
                            
                                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVGERCOMIINTDET", lNumIntTRVGerComiIntDet)
                                    If lErro <> SUCESSO Then gError 197428
                                    
                                End If
                            
                                'serve para indicar que tem aporte para quem nao tem call center
                                If gdAporteValor > 2000 Then
                                    gobjTRVVoucher.iAporte = 1
                                    objComissaoGenerica.dPercentual = 0.015
                                    objComissaoGenerica.dValor = objComissaoGenerica.dValorBase * objComissaoGenerica.dPercentual
                                End If
                            
                            
                                lErro = TRVGerComiIntDet_Acum_Comissoes(alComando(10), gobjTRVVoucher, iVendedor, gdValorBase, dValorDolarBase, gdAporteValor, giAporteMoeda, vdCambio, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, iRelatorio, lNumIntRel, iCallCenter)
                                If lErro <> SUCESSO Then gError 197395
                            
    '                            For Each objTRVAcumVendedor In colTRVAcumVendedor
    '
    '                                If objTRVAcumVendedor.iVendedor = iVendedor Then
    '
    '                                    iAchouVendedor = 1
    '
    '                                    Exit For
    '
    '                                End If
    '
    '                            Next
    '
    '                            If (objTRVAcumVendedor Is Nothing) Then Set objTRVAcumVendedor = New ClassTRVAcumVendedor
    '                            If (objTRVAcumVendedor.colAcumAporte Is Nothing) Then Set objTRVAcumVendedor.colAcumAporte = New Collection
    '
    '                            For Each objTRVAcumAporte In objTRVAcumVendedor.colAcumAporte
    '
    '                                dAporteValorAux = gdAporteValor
    '                                dValorDe = objTRVAcumAporte.dValorDe
    '                                dValorAte = objTRVAcumAporte.dValorAte
    '
    '                                If giAporteMoeda <> objTRVAcumAporte.iMoeda Then
    '                                    If giAporteMoeda = MOEDA_DOLAR Then
    '                                        dAporteValorAux = dAporteValorAux * vdCambio
    '                                    Else
    '                                        dValorDe = dValorDe * vdCambio
    '                                        dValorAte = dValorAte * vdCambio
    '                                    End If
    '                                End If
    '
    '                                If dAporteValorAux >= dValorDe And (dAporteValorAux <= dValorAte Or dValorAte = 0) Then
    '                                    colComissoes(1).dPercentual = objTRVAcumAporte.dPercentual
    '                                    colComissoes(1).dValor = colComissoes(1).dValorBase * colComissoes(1).dPercentual
    '                                    gobjTRVVoucher.iAporte = 1
    '                                    Exit For
    '                                End If
    '                            Next
                            
                                If iRelatorio = 0 Then
                            
                                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOUCHERINFOGER", lNumIntInfoGer)
                                    If lErro <> SUCESSO Then gError 197467
                                    
                                End If
                            
                                lErro = TRVGerComiIntDet_Gravar_Comissoes(alComando(10), iEmpresa, colComissoes, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, colNumIntTRVVoucherInfo, lNumIntTRVGerComiIntDet, lNumIntInfoGer, iRelatorio, lNumIntRel, gobjTRVVoucher, lNumIntRelComiInt, 1, objPlanilha, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, iCallCenter)
                                If lErro <> SUCESSO Then gError 197395
                                
                            End If
                            
                        Else
                                  
                            '************************************************
                            'se o cliente for responsabilidade do call center
                            '************************************************
                            
                            Set colComissoes = New Collection
                      
                            lErro = objComissoesRegrasCalc.Comissoes_Calcula_Regras_Int(Me, bCalculou, colComissoes)
                            If lErro <> SUCESSO Then gError 199431
                          
                            If bCalculou Then
                  
                                Set objComissaoGenerica = colComissoes(1)
                                
                                iVendedor = objComissaoGenerica.iCodVendedor
                  
                                lErro = TRVGerComiIntDet_Acum_Comissoes(alComando(10), gobjTRVVoucher, iVendedor, gdValorBase, dValorDolarBase, gdAporteValor, giAporteMoeda, vdCambio, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, iRelatorio, lNumIntRel, iCallCenter)
                                If lErro <> SUCESSO Then gError 199404
                                
                                If iRelatorio = 0 Then
                            
                                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOUCHERINFOGER", lNumIntInfoGer)
                                    If lErro <> SUCESSO Then gError 199425
                                    
                                End If
                            
                                iRecursivo = 0
                            
                                If iRelatorio = 0 Then
                            
                                    For Each objTRVAcumVendedor In colTRVAcumVendedor
                            
                                        If objTRVAcumVendedor.iVendedor = giVendedorCallCenter Then
                                        
                                            For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
                                            
                                                lErro = Comando_Executar(alComando(9), "INSERT INTO TRVVoucherInfoGer (NumIntVoucherInfo, NumIntInfoGer, Empresa) VALUES (?,?,?)", _
                                                objTRVVoucherInfo.lNumIntDoc, objTRVAcumVendedor.lNumIntInfoGerGrava, iEmpresa)
                                                If lErro <> AD_SQL_SUCESSO Then gError 199426
                            
                                            Next
                            
                                            Exit For
                                            
                                        End If
                            
                                    Next
                            
                                    For Each objTRVAcumVendedor In colTRVAcumVendedor
                            
                                        If objTRVAcumVendedor.iVendedor = iVendedor Then
                                        
                                            For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
                                            
                                                lErro = Comando_Executar(alComando(9), "INSERT INTO TRVVoucherInfoGer (NumIntVoucherInfo, NumIntInfoGer, Empresa) VALUES (?,?,?)", _
                                                objTRVVoucherInfo.lNumIntDoc, objTRVAcumVendedor.lNumIntInfoGerGrava, iEmpresa)
                                                If lErro <> AD_SQL_SUCESSO Then gError 199427
                            
                                            Next
                            
                                            Exit For
                                            
                                        End If
                            
                                    Next
                            
                                End If
                            
                            
                                lErro = TRVGerComiIntDet_Gravar_Comissoes(alComando(10), iEmpresa, colComissoes, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, colNumIntTRVVoucherInfo, lNumIntTRVGerComiIntDet, lNumIntInfoGer, iRelatorio, lNumIntRel, gobjTRVVoucher, lNumIntRelComiInt, 1, objPlanilha, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, iCallCenter)
                                If lErro <> SUCESSO Then gError 199428
                                
                                If iRelatorio = 2 Then
                                                        
                                    iRecursivo = -1
                                                        
                                    'vai gravar o TRVRelVoucherInfo associado ao CALL CENTER
                                    lErro = TRVRelVoucher_Gravar_1(VENDEDOR_CARGO_TRV_PROMOTOR, iRecursivo, iEmpresa, giVendedorCallCenter, colNumIntTRVVoucherInfo, iRelatorio, lNumIntRel, gobjTRVVoucher, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, colComissoes(1).dValorBase)
                                    If lErro <> SUCESSO Then gError 199429
                                    
                                    'vai gravar o TRVRelVoucherInfo associado ao promotor
                                    lErro = TRVRelVoucher_Gravar_1(VENDEDOR_CARGO_TRV_PROMOTOR, iRecursivo, iEmpresa, iVendedor, colNumIntTRVVoucherInfo, iRelatorio, lNumIntRel, gobjTRVVoucher, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, colComissoes(1).dValorBase)
                                    If lErro <> SUCESSO Then gError 199430
                                
                                End If
                                
                            End If
                        
                        End If
              
                    End If
              
                Else
                    
                    'se for venda direta ==> descobre a regiao de venda do cliente
                    'e a partir dela descobre o gerente e diretor que vao acumular o valor para comissao
                    
                    lErro = TRVGerComiIntDet_VendaDireta(iEmpresa, gobjTRVVoucher, colNumIntTRVVoucherInfo, gdValorBase, dValorDolarBase, gdAporteValor, giAporteMoeda, vdCambio, colTRVAcumGerente, colTRVAcumDiretor, iRelatorio, lNumIntRel, iCallCenter)
                    If lErro <> SUCESSO Then gError 197396
              
                End If
              
                'Atualiza Barra de Progressão
                lErro = Rotina_GerComiInt_AtualizaTelaBatch
                If lErro <> SUCESSO Then gError 197397
              
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197461
              
            Loop
      
            alComando(2) = alComandoTVI(2)
            alComando(3) = alComandoTVI(3)
            alComando(5) = alComandoTVI(5)
            alComando(8) = alComandoTVI(8)
            alComando(10) = alComandoTVI(10)
      
        Next
      
    End If
      
    lErro = TRVGerComiIntDet_Gravar_Comissoes_Super(alComando(10), colTRVAcumVendedor, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, iRelatorio, lNumIntRel)
    If lErro <> SUCESSO Then gError 199405
      
''''    lErro = TRVGerComiIntDet_Gravar_Comissoes_Super(colTRVAcumSupervisor, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, iRelatorio, lNumIntRel)
''''    If lErro <> SUCESSO Then gError 199406
''''
''''    lErro = TRVGerComiIntDet_Gravar_Comissoes_Super(colTRVAcumGerente, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, iRelatorio, lNumIntRel)
''''    If lErro <> SUCESSO Then gError 199407
''''
''''    lErro = TRVGerComiIntDet_Gravar_Comissoes_Super(colTRVAcumDiretor, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, iRelatorio, lNumIntRel)
''''    If lErro <> SUCESSO Then gError 197399

    If iRelatorio = 0 Then
      
        lErro = Comando_Executar(alComando(4), "INSERT INTO TRVGerComiInt (NumIntDoc, Codigo, DataGeracao, HoraGeracao, Usuario, DataPagtoDe, DataPagtoAte, DataEmiDe, DataEmiAte) VALUES (?,?,?,?,?,?,?,?,?)", _
        lNumIntTRVGerComiInt, objTRVGerComiInt.lCodigo, objTRVGerComiInt.dtDataGeracao, objTRVGerComiInt.dHoraGeracao, objTRVGerComiInt.sUsuario, objTRVGerComiInt.dtDataPagtoDe, objTRVGerComiInt.dtDataPagtoAte, objTRVGerComiInt.dtDataEmiDe, objTRVGerComiInt.dtDataEmiAte)
        If lErro <> AD_SQL_SUCESSO Then gError 197412
      
    ElseIf iRelatorio = 1 Or iRelatorio = 2 Then
    
        lErro = Comando_Executar(alComando(4), "INSERT INTO TRVRelGerComiInt (NumIntRel, Codigo, DataGeracao, HoraGeracao, Usuario, DataPagtoDe, DataPagtoAte, DataEmiDe, DataEmiAte) VALUES (?,?,?,?,?,?,?,?,?)", _
        lNumIntRel, objTRVGerComiInt.lCodigo, objTRVGerComiInt.dtDataGeracao, objTRVGerComiInt.dHoraGeracao, objTRVGerComiInt.sUsuario, objTRVGerComiInt.dtDataPagtoDe, objTRVGerComiInt.dtDataPagtoAte, objTRVGerComiInt.dtDataEmiDe, objTRVGerComiInt.dtDataEmiAte)
        If lErro <> AD_SQL_SUCESSO Then gError 199483
      
    ElseIf iRelatorio = 3 Then
    
        lErro = CF("Excel_Gera_Planilha", objPlanilha)
        If lErro <> SUCESSO Then gError 199615
      
    End If
      
    'Fecha Comando
    For iIndice = LBound(alComandoTVA) To UBound(alComandoTVA)
        Call Comando_Fechar(alComandoTVA(iIndice))
    Next
                
    For iIndice = LBound(alComandoTVI) To UBound(alComandoTVI)
        Call Comando_Fechar(alComandoTVI(iIndice))
    Next
                
    Call Conexao_DesativarLocks(REATIVAR_LOCKS)
                
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 137414
                
'    If iRelatorio = 1 Then
'
'        lErro = objRelatorio.ExecutarDireto("Previa Comissionamento Interno", "", 0, "", "NNUMINTREL", CStr(lNumIntRel), "DBAIXADE", CStr(objTRVGerComiInt.dtDataPagtoDe), "DBAIXAATE", CStr(objTRVGerComiInt.dtDataPagtoAte), "DEMISSAODE", CStr(objTRVGerComiInt.dtDataEmiDe), "DEMISSAOATE", CStr(objTRVGerComiInt.dtDataEmiAte))
'        If lErro <> SUCESSO Then gError 197470
'
'    ElseIf iRelatorio = 2 Then
'
'        lErro = objRelatorio.ExecutarDireto("Previa Comissionamento Int Detalhado", "", 0, "", "NNUMINTREL", CStr(lNumIntRel), "DBAIXADE", CStr(objTRVGerComiInt.dtDataPagtoDe), "DBAIXAATE", CStr(objTRVGerComiInt.dtDataPagtoAte), "DEMISSAODE", CStr(objTRVGerComiInt.dtDataEmiDe), "DEMISSAOATE", CStr(objTRVGerComiInt.dtDataEmiAte))
'        If lErro <> SUCESSO Then gError 197492
'
'    End If
                
    If iRelatorio <> 0 Then
        Call Rotina_Aviso(vbOKOnly, "Prévia n. " & CStr(lNumIntRel) & " gerada com sucesso.")
    End If
                
    TRVGerComiInt_Gera = SUCESSO

    Exit Function

Erro_TRVGerComiInt_Gera:

    TRVGerComiInt_Gera = gErr

    Select Case gErr

        Case 197382, 197891
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197383 To 197387, 197390
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

        Case 197388, 197389, 197394 To 197401, 197412, 197428, 197469, 197470, 197474, 197492, 197883, 197886, 197890, 199397, 199404 To 199407, 199415, 199425 To 199431, 199483, 199589, 199615, 199858 To 199862

        Case 197391 To 197393
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERINFO", gErr)

        Case 137413
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 137414
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 197468
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVVOUCHERS_CAMBIO_ZERADO", gErr, vlNumVou, vsSerie)

        Case 197473
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVRELVOUCHERINFO", gErr)

        Case 197889
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTGERCOMI_EXECUCAO_EMPRESA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197402)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComandoTVA) To UBound(alComandoTVA)
        Call Comando_Fechar(alComandoTVA(iIndice))
    Next
                
    For iIndice = LBound(alComandoTVI) To UBound(alComandoTVI)
        Call Comando_Fechar(alComandoTVI(iIndice))
    Next

    Call Conexao_DesativarLocks(REATIVAR_LOCKS)

    Call Transacao_Rollback

    Exit Function

End Function

Private Function TRVGerComiInt_Bind_1(ByVal objTRVGerComiInt As ClassTRVGerComiInt, ByVal lComando As Long, vParam1 As Variant, vParam2 As Variant, vParam3 As Variant, vParam4 As Variant, vParam5 As Variant, vParam6 As Variant, vdtDataEmiAux As Variant, vdtDataPagAux As Variant) As Long

Dim lErro As Long

On Error GoTo Erro_TRVGerComiInt_Bind_1

    lErro = Comando_BindVarInt(lComando, vParam1)
    If lErro <> AD_SQL_SUCESSO Then gError 197315

    lErro = Comando_BindVarInt(lComando, vParam2)
    If lErro <> AD_SQL_SUCESSO Then gError 197316
    
    If vParam3 <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vParam3)
        If lErro <> AD_SQL_SUCESSO Then gError 197317
    End If

    If vParam4 <> DATA_NULA Then
        vParam4 = objTRVGerComiInt.dtDataEmiAte
        lErro = Comando_BindVarInt(lComando, vParam4)
        If lErro <> AD_SQL_SUCESSO Then gError 197318
    End If

    If vParam5 <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vParam5)
        If lErro <> AD_SQL_SUCESSO Then gError 197319
    End If

    If vParam6 <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vParam6)
        If lErro <> AD_SQL_SUCESSO Then gError 197320
    End If
    
    If vParam4 <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataEmiAux)
        If lErro <> AD_SQL_SUCESSO Then gError 197320
    
        lErro = Comando_BindVarInt(lComando, vdtDataPagAux)
        If lErro <> AD_SQL_SUCESSO Then gError 197320
    End If

    TRVGerComiInt_Bind_1 = SUCESSO
    
    Exit Function
    
Erro_TRVGerComiInt_Bind_1:

    TRVGerComiInt_Bind_1 = gErr
    
    Select Case gErr

        Case 197315 To 197320
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197321)

    End Select

    Exit Function
    
End Function

Private Function TRVGerComiInt_Bind_2(ByVal lComando As Long, vNumReg As Variant) As Long

Dim lErro As Long

On Error GoTo Erro_TRVGerComiInt_Bind_2

    lErro = Comando_BindVarInt(lComando, vNumReg)
    If lErro <> AD_SQL_SUCESSO Then gError 197458

    TRVGerComiInt_Bind_2 = SUCESSO
    
    Exit Function
    
Erro_TRVGerComiInt_Bind_2:

    TRVGerComiInt_Bind_2 = gErr
    
    Select Case gErr

        Case 197458
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197459)

    End Select

    Exit Function
    
End Function

Private Function TRVGerComiIntDet_Gravar_Comissoes(ByVal lComando As Long, ByVal iEmpresa As Integer, ByVal colComissoes As Collection, ByVal lNumIntTRVGerComiInt As Long, lSeq As Long, ByVal dtDataGeracao As Date, ByVal colNumIntTRVVoucherInfo As Collection, ByVal lNumIntTRVGerComiIntDet As Long, ByVal lNumIntInfoGer As Long, ByVal iRelatorio As Integer, ByVal lNumIntRel As Long, ByVal objTRVVoucher As ClassTRVVouchers, ByVal lNumIntRelComiInt As Long, ByVal iImprime As Integer, ByVal objPlanilha As ClassPlanilhaExcel, Optional ByVal colTRVAcumVendedor As Collection, Optional ByVal colTRVAcumSupervisor As Collection, Optional ByVal colTRVAcumGerente As Collection, Optional ByVal colTRVAcumDiretor As Collection, Optional ByVal iCallCenter As Integer = 0) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objComissaoGenerica As Object
Dim alComando(1 To 8) As Long
Dim lNumIntCom As Long
Dim objTRVAcumVendedor As ClassTRVAcumVendedor
Dim objTRVAcumSupervisor As ClassTRVAcumVendedor
Dim objTRVAcumGerente As ClassTRVAcumVendedor
Dim objTRVAcumDiretor As ClassTRVAcumVendedor
Dim iCargo As Integer
Dim iSuperior As Integer
Dim iVendedor As Integer
Dim vNumIntDoc As Variant
Dim objVendedor As New ClassVendedor
Dim objVendedorTRV As ClassVendedorTRV
Dim objTRVVoucherInfo As ClassTRVVoucherInfoN
Dim sSupervisor As String
Dim sGerente As String
Dim sDiretor As String
Dim sPromotor As String
Dim colVendedores As New Collection

On Error GoTo Erro_TRVGerComiIntDet_Gravar_Comissoes

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197330
    Next

    Set objComissaoGenerica = colComissoes(1)

    'Gera o próximo número interno para a nova Comissão
    lErro = CF("Comissao_Automatico1", alComando(1), alComando(2), lNumIntCom)
    If lErro <> SUCESSO Then gError 197332

    iVendedor = objComissaoGenerica.iCodVendedor

    If iRelatorio = 0 Then

        If iCallCenter = 0 And objTRVVoucher.lCliente < 1000000 Then

        '    VAI GERAR UMA COMISSAO POR VOUCHER PARA PROMOTOR DEVIDO AO VALOR BASE E O PERCENTUAL
            'Insere a comissao no BD
            lErro = Comando_Executar(alComando(3), "INSERT INTO Comissoes (NumIntCom, FilialEmpresa, Status, TipoTitulo, NumIntDoc, CodVendedor, DataGeracao, DataBaixa, Percentual, ValorBase, Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntCom, giFilialEmpresa, STATUS_LIBERADO, TIPO_COMISSAO_TRV, lNumIntTRVGerComiIntDet, objComissaoGenerica.iCodVendedor, dtDataGeracao, DATA_NULA, objComissaoGenerica.dPercentual, objComissaoGenerica.dValorBase, objComissaoGenerica.dValor)
            If lErro <> AD_SQL_SUCESSO Then gError 197333
        
            lErro = Comando_Executar(alComando(4), "INSERT INTO TRVGerComiIntDet (NumIntDoc, NumIntDocGerComi, Seq, NumIntDocComi, ValorBase, ValorComissao, Vendedor, NumIntInfoGer) VALUES (?,?,?,?,?,?,?,?)", _
            lNumIntTRVGerComiIntDet, lNumIntTRVGerComiInt, lSeq, lNumIntCom, objComissaoGenerica.dValorBase, objComissaoGenerica.dValor, objComissaoGenerica.iCodVendedor, lNumIntInfoGer)
            If lErro <> AD_SQL_SUCESSO Then gError 197334
    
        End If
    
        For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
        
            lErro = Comando_Executar(alComando(5), "INSERT INTO TRVVoucherInfoGer (NumIntVoucherInfo, NumIntInfoGer, Empresa) VALUES (?,?,?)", _
            objTRVVoucherInfo.lNumIntDoc, lNumIntInfoGer, iEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 197433
        
'            lErro = Comando_Executar(alComando(6), "UPDATE TRVVoucherInfo SET NumIntDocComiInt = ? WHERE NumIntDoc = ?", lNumIntTRVGerComiInt, objTRVVoucherInfo.lNumIntDoc)
'            If lErro <> AD_SQL_SUCESSO Then gError 197437
        
            If Not (colTRVAcumVendedor Is Nothing) Then
            
                For Each objTRVAcumVendedor In colTRVAcumVendedor
                
                    If objTRVAcumVendedor.iVendedor = iVendedor Then
                        
                        iSuperior = objTRVAcumVendedor.iSuperior
                        
                        objVendedor.iCodigo = iSuperior
                        
                        lErro = CF("Vendedor_Le_Customizado1", objVendedor)
                        If lErro <> SUCESSO Then gError 197462
                        
                        Set objVendedorTRV = objVendedor.objInfoUsu
                        
                        iCargo = objVendedorTRV.iCargo
                        Exit For
                    End If
                    
                Next
                
                lErro = TRVGerComiIntDet_Gravar_Com(lComando, alComando(5), objTRVVoucherInfo, iEmpresa, iCargo, iSuperior, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, objTRVVoucher)
                If lErro <> SUCESSO Then gError 199408
                
'                If iCargo = VENDEDOR_CARGO_TRV_SUPERVISOR Then
'
'                    For Each objTRVAcumSupervisor In colTRVAcumSupervisor
'
'                        If objTRVAcumSupervisor.iVendedor = iSuperior Then
'                            iSuperior = objTRVAcumSupervisor.iSuperior
'                            objVendedor.iCodigo = iSuperior
'
'                            lErro = CF("Vendedor_Le_Customizado1", objVendedor)
'                            If lErro <> SUCESSO Then gError 197463
'
'                            Set objVendedorTRV = objVendedor.objInfoUsu
'
'                            iCargo = objVendedorTRV.iCargo
'
'                            lErro = Comando_Executar(alComando(5), "INSERT INTO TRVVoucherInfoGer (NumIntVoucherInfo, NumIntInfoGer, Empresa) VALUES (?,?,?)", _
'                            objTRVVoucherInfo.lNumIntDoc, objTRVAcumSupervisor.lNumIntInfoGerGrava, iEmpresa)
'                            If lErro <> AD_SQL_SUCESSO Then gError 197434
'
'                            Exit For
'                        End If
'
'                    Next
'                End If
'
'                If iCargo = VENDEDOR_CARGO_TRV_GERENTE Then
'
'                    For Each objTRVAcumGerente In colTRVAcumGerente
'
'                        If objTRVAcumGerente.iVendedor = iSuperior Then
'                            iSuperior = objTRVAcumGerente.iSuperior
'                            objVendedor.iCodigo = iSuperior
'
'                            lErro = CF("Vendedor_Le_Customizado1", objVendedor)
'                            If lErro <> SUCESSO Then gError 197464
'
'                            Set objVendedorTRV = objVendedor.objInfoUsu
'
'                            iCargo = objVendedorTRV.iCargo
'
'                            lErro = Comando_Executar(alComando(5), "INSERT INTO TRVVoucherInfoGer (NumIntVoucherInfo, NumIntInfoGer, Empresa) VALUES (?,?,?)", _
'                            objTRVVoucherInfo.lNumIntDoc, objTRVAcumGerente.lNumIntInfoGerGrava, iEmpresa)
'                            If lErro <> AD_SQL_SUCESSO Then gError 197435
'
'                            Exit For
'                        End If
'
'                    Next
'                End If
'
'
'                If iCargo = VENDEDOR_CARGO_TRV_DIRETOR Then
'
'                    For Each objTRVAcumDiretor In colTRVAcumDiretor
'
'                        If objTRVAcumDiretor.iVendedor = iSuperior Then
'
'                            lErro = Comando_Executar(alComando(5), "INSERT INTO TRVVoucherInfoGer (NumIntVoucherInfo, NumIntInfoGer, Empresa) VALUES (?,?,?)", _
'                            objTRVVoucherInfo.lNumIntDoc, objTRVAcumDiretor.lNumIntInfoGerGrava, iEmpresa)
'                            If lErro <> AD_SQL_SUCESSO Then gError 197436
'
'                            Exit For
'                        End If
'
'                    Next
'                End If
            End If
        Next

    ElseIf iRelatorio = 1 Or iRelatorio = 2 Then

        If iCallCenter = 0 And objTRVVoucher.lCliente < 1000000 Then

            lErro = Comando_Executar(alComando(7), "INSERT INTO TRVRelComiInt (NumIntRel, Seq, CodVendedor, Percentual, ValorBase, Valor, NumIntRelComiInt, ValorAcum, ValorDolarAcum, ValorAcumAporte, ValorDolarAcumAporte, Imprime) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntRel, lSeq, objComissaoGenerica.iCodVendedor, objComissaoGenerica.dPercentual, objComissaoGenerica.dValorBase, objComissaoGenerica.dValor, lNumIntRelComiInt, objTRVVoucher.dValorAcum, objTRVVoucher.dValorDolarAcum, objTRVVoucher.dValorAcumAporte, objTRVVoucher.dValorDolarAcumAporte, iImprime)
            If lErro <> AD_SQL_SUCESSO Then gError 197471
            
        End If

        If iRelatorio = 2 Then
            
            If objTRVVoucher.lNumVou <> 0 Then
    
                If iCallCenter = 0 And objTRVVoucher.lCliente < 1000000 Then
    
                    lErro = Comando_Executar(alComando(8), "INSERT INTO TRVRelVoucher (NumIntRel, Cliente, Produto, DiasAntc, TipVou, Serie, NumVou, AporteValor, AporteMoeda, TipoCliente, NumIntRelComiInt, Empresa, Data, DataFat, Aporte, ValorBase) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    lNumIntRel, objTRVVoucher.lCliente, objTRVVoucher.sProduto, objTRVVoucher.idiasantc, objTRVVoucher.sTipVou, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, objTRVVoucher.dAporteValor, objTRVVoucher.iAporteMoeda, objTRVVoucher.iTipoCliente, lNumIntRelComiInt, iEmpresa, objTRVVoucher.dtData, objTRVVoucher.dtDataFat, objTRVVoucher.iAporte, objComissaoGenerica.dValorBase)
                    If lErro <> AD_SQL_SUCESSO Then gError 197472
                
                End If
                
                lErro = TRVRelVoucher_Gravar(lComando, iEmpresa, iVendedor, colNumIntTRVVoucherInfo, iRelatorio, lNumIntRel, objTRVVoucher, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, objComissaoGenerica.dValorBase)
                If lErro <> SUCESSO Then gError 197487
                
                
            End If
        End If
        
    ElseIf iRelatorio = 3 Then
    
        objVendedor.iCodigo = iVendedor
    
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then gError 199586
        
'        If lErro <> SUCESSO Then gError 199587
        
        sPromotor = objVendedor.sNomeReduzido
        
        lErro = CF("TRVVouVendedores_Le", colVendedores, objTRVVoucher, lComando)
        If lErro <> SUCESSO Then gError 199513
    
        For Each objVendedor In colVendedores
        
            For Each objTRVAcumVendedor In colTRVAcumVendedor

                If objTRVAcumVendedor.iVendedor = objVendedor.iCodigo Then
                    If objTRVAcumVendedor.iCargo = VENDEDOR_CARGO_TRV_SUPERVISOR Then
                        sSupervisor = objVendedor.sNomeReduzido
                    ElseIf objTRVAcumVendedor.iCargo = VENDEDOR_CARGO_TRV_GERENTE Then
                        sGerente = objVendedor.sNomeReduzido
                    ElseIf objTRVAcumVendedor.iCargo = VENDEDOR_CARGO_TRV_DIRETOR Then
                        sDiretor = objVendedor.sNomeReduzido
                    End If
                    Exit For
                
                End If
            
            Next
            
        Next
        
        lErro = Cria_Linha_Excel(objPlanilha, iEmpresa, sPromotor, sSupervisor, sGerente, sDiretor, objTRVVoucher, objComissaoGenerica)
        If lErro <> SUCESSO Then gError 199614
        
    End If


    lSeq = lSeq + 1
        
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVGerComiIntDet_Gravar_Comissoes = SUCESSO
    
    Exit Function
    
Erro_TRVGerComiIntDet_Gravar_Comissoes:

    TRVGerComiIntDet_Gravar_Comissoes = gErr

    Select Case gErr

        Case 197330
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 197331, 197332, 197462, 197463, 197464, 197487, 199408, 199586, 199614
        
        Case 197333
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMISSOES", gErr, TIPO_COMISSAO_TRV, lNumIntCom)
        
        Case 197334
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVGERCOMIINTDET", gErr, lNumIntTRVGerComiIntDet)
        
        Case 197433, 197434, 197435, 197436
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVVOUCHERINFOGER", gErr)
        
        Case 197437
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_TRVVOUCHERINFO", gErr, vNumIntDoc)
        
        Case 197471
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVRELCOMIINT", gErr)
        
        Case 197472, 197476, 197478, 197480
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVRELVOUCHER", gErr)
        
        Case 197477, 197479, 197481
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVRELVOUCHERINFO", gErr)
        
        Case 199587
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objTRVVoucher.lNumVou)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197335)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function TRVGerComiIntDet_Gravar_Com(ByVal lComando1 As Long, ByVal lComando As Long, ByVal objTRVVoucherInfo As ClassTRVVoucherInfoN, iEmpresa As Integer, ByVal iCargo As Integer, ByVal iSuperior As Integer, ByVal colTRVAcumPromotor As Collection, ByVal colTRVAcumSupervisor As Collection, ByVal colTRVAcumGerente As Collection, ByVal colTRVAcumDiretor As Collection, ByVal objTRVVoucher As ClassTRVVouchers) As Long

Dim objVendedor As New ClassVendedor
Dim objTRVAcumVendedor As ClassTRVAcumVendedor
Dim colTRVAcumVendedor As Collection
Dim objVendedorTRV As ClassVendedorTRV
Dim lErro As Long
Dim colVendedores As New Collection
Dim vVendedor As Variant


On Error GoTo Erro_TRVGerComiIntDet_Gravar_Com

''''    If iSuperior <> 0 Then
''''
''''        If iCargo = VENDEDOR_CARGO_TRV_PROMOTOR Then
''''            Set colTRVAcumVendedor = colTRVAcumPromotor
''''        ElseIf iCargo = VENDEDOR_CARGO_TRV_SUPERVISOR Then
''''            Set colTRVAcumVendedor = colTRVAcumSupervisor
''''        ElseIf iCargo = VENDEDOR_CARGO_TRV_GERENTE Then
''''            Set colTRVAcumVendedor = colTRVAcumGerente
''''        ElseIf iCargo = VENDEDOR_CARGO_TRV_DIRETOR Then
''''            Set colTRVAcumVendedor = colTRVAcumDiretor
''''        End If
''''
''''        If Not colTRVAcumVendedor Is Nothing Then
''''
''''            For Each objTRVAcumVendedor In colTRVAcumVendedor
''''
''''                If objTRVAcumVendedor.iVendedor = iSuperior Then
''''
''''                    iSuperior = objTRVAcumVendedor.iSuperior
''''
''''                    If iSuperior <> 0 Then
''''
''''                        objVendedor.iCodigo = iSuperior
''''
''''                        lErro = CF("Vendedor_Le_Customizado1", objVendedor)
''''                        If lErro <> SUCESSO Then gError 199409
''''
''''                        Set objVendedorTRV = objVendedor.objInfoUsu
''''
''''                        iCargo = objVendedorTRV.iCargo
''''
''''                    Else
''''                        iCargo = 0
''''                    End If
''''
''''                    lErro = Comando_Executar(lComando, "INSERT INTO TRVVoucherInfoGer (NumIntVoucherInfo, NumIntInfoGer, Empresa) VALUES (?,?,?)", _
''''                    objTRVVoucherInfo.lNumIntDoc, objTRVAcumVendedor.lNumIntInfoGerGrava, iEmpresa)
''''                    If lErro <> AD_SQL_SUCESSO Then gError 199410
''''
''''                    Exit For
''''                End If
''''
''''            Next
''''
''''            lErro = TRVGerComiIntDet_Gravar_Com(lComando, objTRVVoucherInfo, iEmpresa, iCargo, iSuperior, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor)
''''            If lErro <> SUCESSO Then gError 199411
''''
''''        End If
''''
''''    End If

    lErro = CF("TRVVouVendedores_Le", colVendedores, objTRVVoucher, lComando1)
    If lErro <> SUCESSO Then gError 199520
    
    For Each objVendedor In colVendedores
        
        If Not colTRVAcumVendedor Is Nothing Then

            For Each objTRVAcumVendedor In colTRVAcumPromotor

                If objTRVAcumVendedor.iVendedor = objVendedor.iCodigo Then
        
                    lErro = Comando_Executar(lComando, "INSERT INTO TRVVoucherInfoGer (NumIntVoucherInfo, NumIntInfoGer, Empresa) VALUES (?,?,?)", _
                    objTRVVoucherInfo.lNumIntDoc, objTRVAcumVendedor.lNumIntInfoGerGrava, iEmpresa)
                    If lErro <> AD_SQL_SUCESSO Then gError 199410
                    
                    Exit For
                    
                End If
                
            Next
            
        End If

    Next

    TRVGerComiIntDet_Gravar_Com = SUCESSO
    
    Exit Function
    
Erro_TRVGerComiIntDet_Gravar_Com:

    TRVGerComiIntDet_Gravar_Com = gErr

    Select Case gErr

        Case 199409, 199411, 199520
        
        Case 199410
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVVOUCHERINFOGER", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199412)

    End Select

    Exit Function

End Function

Private Function TRVRelVoucher_Gravar(ByVal lComando As Long, ByVal iEmpresa As Integer, ByVal iVendedor As Integer, ByVal colNumIntTRVVoucherInfo As Collection, ByVal iRelatorio As Integer, ByVal lNumIntRel As Long, ByVal objTRVVoucher As ClassTRVVouchers, ByVal colTRVAcumVendedor As Collection, ByVal colTRVAcumSupervisor As Collection, ByVal colTRVAcumGerente As Collection, ByVal colTRVAcumDiretor As Collection, ByVal dValorBase As Double) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim objTRVAcumVendedor As ClassTRVAcumVendedor
Dim objTRVAcumSupervisor As ClassTRVAcumVendedor
Dim objTRVAcumGerente As ClassTRVAcumVendedor
Dim objTRVAcumDiretor As ClassTRVAcumVendedor
Dim iCargo As Integer
Dim iSuperior As Integer
Dim objVendedor As New ClassVendedor
Dim objVendedorTRV As ClassVendedorTRV
Dim objTRVVoucherInfo As ClassTRVVoucherInfoN
Dim iRecursivo As Integer
Dim vVendedor As Variant
Dim colVendedores As New Collection

On Error GoTo Erro_TRVRelVoucher_Gravar

''''    If Not (colTRVAcumVendedor Is Nothing) Then
''''
''''        For Each objTRVAcumVendedor In colTRVAcumVendedor
''''
''''            If objTRVAcumVendedor.iVendedor = iVendedor Then
''''
''''                iSuperior = objTRVAcumVendedor.iSuperior
''''
''''                objVendedor.iCodigo = iSuperior
''''
''''                lErro = CF("Vendedor_Le_Customizado1", objVendedor)
''''                If lErro <> SUCESSO Then gError 197462
''''
''''                Set objVendedorTRV = objVendedor.objInfoUsu
''''
''''                If Not objVendedorTRV Is Nothing Then
''''                    iCargo = objVendedorTRV.iCargo
''''                Else
''''                    iCargo = 0
''''                End If
''''                Exit For
''''            End If
''''
''''        Next
        
    lErro = CF("TRVVouVendedores_Le", colVendedores, objTRVVoucher, lComando)
    If lErro <> SUCESSO Then gError 199513
    
    For Each objVendedor In colVendedores
        
            lErro = TRVRelVoucher_Gravar_1(iCargo, iRecursivo, iEmpresa, objVendedor.iCodigo, colNumIntTRVVoucherInfo, iRelatorio, lNumIntRel, objTRVVoucher, colTRVAcumVendedor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, dValorBase)
            If lErro <> SUCESSO Then gError 199387
            
    Next

    TRVRelVoucher_Gravar = SUCESSO
    
    Exit Function
    
Erro_TRVRelVoucher_Gravar:

    TRVRelVoucher_Gravar = gErr

    Select Case gErr

        Case 197462, 199387, 199513
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197486)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Private Function TRVVoucher_Le_GerComiInt(ByVal lComando As Long, ByVal sSelecao As String, objTRVGerComiInt As ClassTRVGerComiInt, vTipVou As Variant, vNumVou As Variant, vSerie As Variant, vCliente As Variant, vCambio As Variant, vAporteValor As Variant, vClienteTipo As Variant, vProduto As Variant, vDiasAntc As Variant, vAporteMoeda As Variant, vPromotor As Variant, vEmpresaPai As Variant, vAporte As Variant, vData As Variant, vdTarifaUnitaria As Variant, vdPercComiProd As Variant, viVendProd As Variant, viVendCli As Variant, viVigencia As Variant, vdTarifaUNNormal As Variant, vdtDataFat As Variant, vsNomeRedCli As Variant, vsNomeRedProd As Variant, vlNumFatCoinfo As Variant, vsMoeda As Variant, vdValorCambio As Variant, vdValorBruto As Variant, vsUsuRespCallCenter As Variant) As Long

Dim sComando_SQL As String
Dim lErro As Long
Dim vdtDataEmiDe As Variant
Dim vdtDataEmiAte As Variant
Dim vdtDataPagDe As Variant
Dim vdtDataPagAte As Variant
Dim vsTipoDoc As Variant
Dim viTipoDocDestino As Variant
Dim vdtDataEmiAux As Variant
Dim vdtDataPagAux As Variant

On Error GoTo Erro_TRVVoucher_Le_GerComiInt

    sComando_SQL = "SELECT V.TipVou, V.NumVou, V.Serie, V.ClienteVou, V.Cambio, C.Tipo, V.Produto, V.DiasAntc, V.Promotor, FT.EmpresaPai, FT.Aporte, V.Data, V.TarifaUnitaria , FC.Vendedor, PT.Vendedor, PT.PercComissao, V.Vigencia, V.TarifaUNNormal, B.DataBaixa, C.NomeReduzido, P.NomeReduzido, V.NumFatCoinfo, V.Moeda, V.ValorCambio, V.ValorBruto, C.UsuRespCallCenter "
    sComando_SQL = sComando_SQL & "FROM PrimeiraBaixaTitRec AS B, ClientesTVA AS C, FiliaisClientesTRVTVA AS FT, TRVVouchers AS V, ProdutosTVA AS P, FiliaisClientesTVA AS FC, TRVProdutosTVA PT " & sSelecao & " ORDER BY V.NumVou"

    'sComando_SQL = "SELECT TipVou, NumVou, Serie, TRVVouchers.ClienteVou, Cambio, Clientes.Tipo, Produto, DiasAntc, TRVVouchers.Promotor, FiliaisClientesTRV.EmpresaPai, FiliaisClientesTRV.Aporte, TRVVouchers.Data, TRVVouchers.TarifaUnitaria FROM ParcelasRecBaixadas, BaixasParcRec, BaixasRec, Clientes, FiliaisClientesTRV, TRVVouchers  " & sSelecao

'From ParcelasRecBaixadas, BaixasParcRec, BaixasRec, Clientes, TRVVouchers
'    WHERE TRVVouchers.TipoDoc = ? AND TRVVouchers.TipoDocDestino = ? AND TRVVouchers.NumIntDocDestino = ParcelasRecBaixadas.NumIntTitulo AND BaixasParcRec.NumIntParcela = ParcelasRecBaixadas.NumINtDoc AND BaixasParcRec.NumIntBaixa = BaixasRec.NumIntBaixa AND "
'    sSelecao = sSelecao & "ParcelasRecBaixadas.NumParcela = 1 AND BaixasParcRec.Sequencial = 1 AND Clientes.Codigo = TRVVouchers.Cliente AND Clientes.Tipo <> 7 AND EXISTS (SELECT * FROM TRVVoucherInfo WHERE NumIntDocComiInt = 0 AND TRVVoucherInfo.NumVou = TRVVouchers.NumVou AND TRVVoucherInfo.Serie = TRVVouchers.Serie)"

    lErro = Comando_PrepararInt(lComando, sComando_SQL)
    If lErro <> AD_SQL_SUCESSO Then gError 197336

    lErro = Comando_BindVarInt(lComando, vTipVou)
    If lErro <> AD_SQL_SUCESSO Then gError 197408

    lErro = Comando_BindVarInt(lComando, vNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 197337
    
    lErro = Comando_BindVarInt(lComando, vSerie)
    If lErro <> AD_SQL_SUCESSO Then gError 197338
    
    lErro = Comando_BindVarInt(lComando, vCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 197339
    
    lErro = Comando_BindVarInt(lComando, vCambio)
    If lErro <> AD_SQL_SUCESSO Then gError 197340
    
    lErro = Comando_BindVarInt(lComando, vClienteTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 197373

    lErro = Comando_BindVarInt(lComando, vProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 197405

    lErro = Comando_BindVarInt(lComando, vDiasAntc)
    If lErro <> AD_SQL_SUCESSO Then gError 197407

    lErro = Comando_BindVarInt(lComando, vPromotor)
    If lErro <> AD_SQL_SUCESSO Then gError 197875

    lErro = Comando_BindVarInt(lComando, vEmpresaPai)
    If lErro <> AD_SQL_SUCESSO Then gError 197876

    lErro = Comando_BindVarInt(lComando, vAporte)
    If lErro <> AD_SQL_SUCESSO Then gError 197877

    lErro = Comando_BindVarInt(lComando, vData)
    If lErro <> AD_SQL_SUCESSO Then gError 197878
    
    lErro = Comando_BindVarInt(lComando, vdTarifaUnitaria)
    If lErro <> AD_SQL_SUCESSO Then gError 197878
          
    lErro = Comando_BindVarInt(lComando, viVendCli)
    If lErro <> AD_SQL_SUCESSO Then gError 197878

    lErro = Comando_BindVarInt(lComando, viVendProd)
    If lErro <> AD_SQL_SUCESSO Then gError 197878

    lErro = Comando_BindVarInt(lComando, vdPercComiProd)
    If lErro <> AD_SQL_SUCESSO Then gError 197878

    lErro = Comando_BindVarInt(lComando, viVigencia)
    If lErro <> AD_SQL_SUCESSO Then gError 197878
    
    lErro = Comando_BindVarInt(lComando, vdTarifaUNNormal)
    If lErro <> AD_SQL_SUCESSO Then gError 197878
    
    lErro = Comando_BindVarInt(lComando, vdtDataFat)
    If lErro <> AD_SQL_SUCESSO Then gError 197878
    
    lErro = Comando_BindVarInt(lComando, vsNomeRedCli)
    If lErro <> AD_SQL_SUCESSO Then gError 197878
    
    lErro = Comando_BindVarInt(lComando, vsNomeRedProd)
    If lErro <> AD_SQL_SUCESSO Then gError 197878
    
    lErro = Comando_BindVarInt(lComando, vlNumFatCoinfo)
    If lErro <> AD_SQL_SUCESSO Then gError 197878
    
    lErro = Comando_BindVarInt(lComando, vsMoeda)
    If lErro <> AD_SQL_SUCESSO Then gError 197878
    
    lErro = Comando_BindVarInt(lComando, vdValorCambio)
    If lErro <> AD_SQL_SUCESSO Then gError 197878
    
    lErro = Comando_BindVarInt(lComando, vdValorBruto)
    If lErro <> AD_SQL_SUCESSO Then gError 197878
    
    lErro = Comando_BindVarInt(lComando, vsUsuRespCallCenter)
    If lErro <> AD_SQL_SUCESSO Then gError 207555
    
    vsTipoDoc = TRV_TIPODOC_VOU_TEXTO
    viTipoDocDestino = TRV_TIPO_DOC_DESTINO_TITREC
    vdtDataEmiDe = objTRVGerComiInt.dtDataEmiDe
    vdtDataEmiAte = objTRVGerComiInt.dtDataEmiAte
    vdtDataPagDe = objTRVGerComiInt.dtDataPagtoDe
    vdtDataPagAte = objTRVGerComiInt.dtDataPagtoAte
    
    vdtDataEmiAux = StrParaDate("01/" & CStr(Month(objTRVGerComiInt.dtDataEmiAte)) & "/" & CStr(Year(objTRVGerComiInt.dtDataEmiAte)))
    vdtDataPagAux = objTRVGerComiInt.dtDataPagtoDe

    lErro = TRVGerComiInt_Bind_1(objTRVGerComiInt, lComando, vsTipoDoc, viTipoDocDestino, vdtDataEmiDe, vdtDataEmiAte, vdtDataPagDe, vdtDataPagAte, vdtDataEmiAux, vdtDataPagAux)
    If lErro <> AD_SQL_SUCESSO Then gError 197341

    lErro = Comando_ExecutarInt(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 197342

    TRVVoucher_Le_GerComiInt = SUCESSO
    
    Exit Function
    
Erro_TRVVoucher_Le_GerComiInt:

    TRVVoucher_Le_GerComiInt = gErr

    Select Case gErr

        Case 197336 To 197342, 197344, 197373, 197405, 197407, 197408, 197432, 197878, 207555
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197343)

    End Select

    Exit Function

End Function

Private Function TRVGerComiIntDet_Gravar_Comissoes_Super(lComando1 As Long, colTRVAcum As Collection, ByVal lNumIntTRVGerComiInt As Long, lSeq As Long, ByVal dtDataGeracao As Date, ByVal iRelatorio As Integer, ByVal lNumIntRel As Long) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntTRVGerComiIntDet As Long
Dim lNumIntCom As Long
Dim objTRVAcum As ClassTRVAcumVendedor
Dim objTRVAcumGerente As ClassTRVAcumVendedor
Dim objVendedorTRV As ClassVendedorTRV
Dim objVNDComissaoTRV As ClassVNDComissaoTRV
Dim objVNDReducaoTRV As ClassVNDReducaoTRV
Dim dPercentualReducao As Double
Dim objComissao As ClassComissao
Dim colComissao As Collection
Dim iVendedor As Integer
Dim dCambioMedio As Double
Dim dValorBase As Double
Dim dPercentual As Double
Dim objVendedor As New ClassVendedor
Dim objTRVAcumAporte As ClassTRVAcumAporte
Dim colNumIntTRVVoucherInfo As New Collection
Dim dValorTotal As Double
Dim dValorTotalDolar As Double
Dim dValorTotalReal As Double
Dim iMoedaFaixa As Integer
Dim dValorFaixa As Double
Dim dAux As Double
Dim objTRVVoucher As New ClassTRVVouchers
Dim dValorAcumAporte As Double
Dim dValorDolarAcumAporte As Double
Dim iImprime As Integer
Dim objTRVAcumAux As ClassTRVAcumAux
Dim objPlanilha As New ClassPlanilhaExcel

On Error GoTo Erro_TRVGerComiIntDet_Gravar_Comissoes_Super

    If iRelatorio <> 3 Then

        For Each objTRVAcum In colTRVAcum
    
            iImprime = 1
    
            If objTRVAcum.dValorAcumCallCenter <> 0 Then
            
                Set objComissao = New ClassComissao
                objComissao.iCodVendedor = objTRVAcum.iVendedor
                objComissao.dPercentual = objTRVAcum.dPercCallCenter
                objComissao.dValorBase = objTRVAcum.dValorAcumCallCenter
                objComissao.dValor = objComissao.dValorBase * objComissao.dPercentual
                
                Set colComissao = New Collection
                
                colComissao.Add objComissao
                
                If iRelatorio = 0 Then
                
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVGERCOMIINTDET", lNumIntTRVGerComiIntDet)
                    If lErro <> SUCESSO Then gError 197887
                    
                End If
                
                lErro = TRVGerComiIntDet_Gravar_Comissoes(lComando1, 0, colComissao, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, colNumIntTRVVoucherInfo, lNumIntTRVGerComiIntDet, objTRVAcum.lNumIntInfoGerCallCenter, iRelatorio, lNumIntRel, objTRVVoucher, objTRVAcum.lNumIntRelComiIntCallCenter, iImprime, objPlanilha)
                If lErro <> SUCESSO Then gError 197888
            
            End If
            
            
            If objTRVAcum.dValorAcumMyAssistance <> 0 Then
            
                If objTRVAcum.iCargo = VENDEDOR_CARGO_TRV_PROMOTOR Or objTRVAcum.iCargo = VENDEDOR_CARGO_TRV_SUPERVISOR Or objTRVAcum.iCargo = VENDEDOR_CARGO_TRV_GERENTE Then
            
                    Set objComissao = New ClassComissao
                    objComissao.iCodVendedor = objTRVAcum.iVendedor
                    
                    If objTRVAcum.iCargo = VENDEDOR_CARGO_TRV_PROMOTOR Then
                        objComissao.dPercentual = 0.01
                    ElseIf objTRVAcum.iCargo = VENDEDOR_CARGO_TRV_SUPERVISOR Then
                        objComissao.dPercentual = 0.005
                    ElseIf objTRVAcum.iCargo = VENDEDOR_CARGO_TRV_GERENTE Then
                        objComissao.dPercentual = 0.003
                    End If
                    
                    objComissao.dValorBase = objTRVAcum.dValorAcumMyAssistance
                    objComissao.dValor = objComissao.dValorBase * objComissao.dPercentual
                    
                    Set colComissao = New Collection
                    
                    colComissao.Add objComissao
                    
                    If iRelatorio = 0 Then
                    
                        lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVGERCOMIINTDET", lNumIntTRVGerComiIntDet)
                        If lErro <> SUCESSO Then gError 197887
                        
                    End If
                    
                    lErro = TRVGerComiIntDet_Gravar_Comissoes(lComando1, 0, colComissao, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, colNumIntTRVVoucherInfo, lNumIntTRVGerComiIntDet, objTRVAcum.lNumIntInfoGerCallCenter, iRelatorio, lNumIntRel, objTRVVoucher, objTRVAcum.lNumIntRelComiIntCallCenter, iImprime, objPlanilha)
                    If lErro <> SUCESSO Then gError 197888
                    
                End If
            
            End If
            
    
            objVendedor.iCodigo = objTRVAcum.iVendedor
            
            lErro = CF("Vendedor_Le_Customizado1", objVendedor)
            If lErro <> SUCESSO Then gError 197360
            
            Set objVendedorTRV = objVendedor.objInfoUsu
            
            dValorAcumAporte = 0
            dValorDolarAcumAporte = 0
            
            For Each objTRVAcumAporte In objTRVAcum.colAcumAporte
                dValorAcumAporte = dValorAcumAporte + objTRVAcumAporte.dValorAcumAporte
                dValorDolarAcumAporte = dValorDolarAcumAporte + objTRVAcumAporte.dValorDolarAcumAporte
                
            Next
            
            
            dValorTotalReal = dValorAcumAporte + objTRVAcum.dValorAcum
            dValorTotalDolar = dValorDolarAcumAporte + objTRVAcum.dValorDolarAcum
            
            objTRVVoucher.dValorAcum = objTRVAcum.dValorAcum
            objTRVVoucher.dValorDolarAcum = objTRVAcum.dValorDolarAcum
            objTRVVoucher.dValorAcumAporte = dValorAcumAporte
            objTRVVoucher.dValorDolarAcumAporte = dValorDolarAcumAporte
            
            If objVendedorTRV.colVNDComissaoTRV.Count > 0 Then
            
                iMoedaFaixa = objVendedorTRV.colVNDComissaoTRV(1).iMoeda
            
                If iMoedaFaixa = MOEDA_REAL Then
                    dValorTotal = dValorTotalReal
                Else
                    dValorTotal = dValorTotalDolar
                End If
                
                For Each objVNDComissaoTRV In objVendedorTRV.colVNDComissaoTRV
        
                    If dValorTotal <= 0 Then Exit For
                    
                    If dValorTotal > objVNDComissaoTRV.dValorDe Then
                                
                        dAux = Round((objVNDComissaoTRV.dValorDe - DELTA_VALORMONETARIO), 2)
                                
                        dValorFaixa = dValorTotal - IIf(dAux < 0, 0, dAux)
                            
                        'tenta alocar primeiro os aportes pelas faixas mais altas
                        For Each objTRVAcumAporte In objTRVAcum.colAcumAporte
                        
                            If dValorFaixa = 0 Then Exit For
                        
                            If iMoedaFaixa = MOEDA_REAL Then
                                
                                If objTRVAcumAporte.dValorAcumAporte > 0 Then
                                
                                    If objTRVAcumAporte.dValorAcumAporte > dValorFaixa Then
                                        dValorBase = dValorFaixa
                                        objTRVAcumAporte.dValorAcumAporte = objTRVAcumAporte.dValorAcumAporte - dValorBase
                                        dValorFaixa = 0
                                    Else
                                        dValorBase = objTRVAcumAporte.dValorAcumAporte
                                        objTRVAcumAporte.dValorAcumAporte = 0
                                        dValorFaixa = dValorFaixa - dValorBase
                                    End If
                                    
                                    dValorTotal = dValorTotal - dValorBase
                                    
                                End If
                                    
                            Else
                            
                                If objTRVAcumAporte.dValorDolarAcumAporte > 0 Then
                                
                                    If objTRVAcumAporte.dValorDolarAcumAporte > dValorFaixa Then
                                        dValorBase = dValorFaixa
                                        objTRVAcumAporte.dValorDolarAcumAporte = objTRVAcumAporte.dValorDolarAcumAporte - dValorBase
                                        dValorFaixa = 0
                                    Else
                                        dValorBase = objTRVAcumAporte.dValorDolarAcumAporte
                                        objTRVAcumAporte.dValorAcumAporte = 0
                                        dValorFaixa = dValorFaixa - dValorBase
                                    End If
                                
                                    dValorTotal = dValorTotal - dValorBase
                                
                                    dValorBase = dValorBase * objTRVAcumAporte.dCambioMedio
                                
                                End If
                                
                            End If
                                    
                            dPercentual = IIf(objTRVAcumAporte.dPercentual < objVNDComissaoTRV.dPercComissao, objTRVAcumAporte.dPercentual, objVNDComissaoTRV.dPercComissao)
                            
                            If dPercentual <> 0 And dValorBase <> 0 Then
                            
                                Set objComissao = New ClassComissao
                                objComissao.iCodVendedor = objTRVAcum.iVendedor
                                objComissao.dPercentual = dPercentual
                                objComissao.dValorBase = dValorBase
                                objComissao.dValor = objComissao.dValorBase * objComissao.dPercentual
                                
                                Set colComissao = New Collection
                            
                                colComissao.Add objComissao
                            
                                If iRelatorio = 0 Then
                                
                                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVGERCOMIINTDET", lNumIntTRVGerComiIntDet)
                                    If lErro <> SUCESSO Then gError 197465
                                    
                                End If
                                
                                lErro = TRVGerComiIntDet_Gravar_Comissoes(lComando1, 0, colComissao, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, colNumIntTRVVoucherInfo, lNumIntTRVGerComiIntDet, objTRVAcum.lNumIntInfoGer, iRelatorio, lNumIntRel, objTRVVoucher, objTRVAcum.lNumIntRelComiInt, iImprime, objPlanilha)
                                If lErro <> SUCESSO Then gError 197361
                                    
                                iImprime = 0
                                
                            End If
                        Next
                            
                        If dValorFaixa > 0 Then
                        
                            If iMoedaFaixa = MOEDA_REAL Then
                        
                                If objTRVAcum.dValorAcum > 0 Then
                                    
                                    If objTRVAcum.dValorAcum > dValorFaixa Then
                                        dValorBase = dValorFaixa
                                        objTRVAcum.dValorAcum = objTRVAcum.dValorAcum - dValorBase
                                        dValorFaixa = 0
                                    Else
                                        dValorBase = objTRVAcum.dValorAcum
                                        objTRVAcum.dValorAcum = 0
                                        dValorFaixa = 0
                                    End If
                                    
                                    dValorTotal = dValorTotal - dValorBase
                                    
                                End If
                                    
                            Else
                                    
                                If objTRVAcum.dValorDolarAcum > 0 Then
                                    
                                    If objTRVAcum.dValorDolarAcum > dValorFaixa Then
                                        dValorBase = dValorFaixa
                                        objTRVAcum.dValorDolarAcum = objTRVAcum.dValorDolarAcum - dValorBase
                                        dValorFaixa = 0
                                    Else
                                        dValorBase = objTRVAcum.dValorDolarAcum
                                        objTRVAcum.dValorAcum = 0
                                        dValorFaixa = 0
                                    End If
                                    
                                    dValorTotal = dValorTotal - dValorBase
                                
                                    dValorBase = dValorBase * objTRVAcum.dCambioMedio
                                    
                                End If
                                    
                            End If
                                    
                            If objVNDComissaoTRV.dPercComissao <> 0 And dValorBase <> 0 Then
                                    
                                Set objComissao = New ClassComissao
                                objComissao.iCodVendedor = objTRVAcum.iVendedor
                                objComissao.dPercentual = objVNDComissaoTRV.dPercComissao
                                objComissao.dValorBase = dValorBase
                                objComissao.dValor = objComissao.dValorBase * objComissao.dPercentual
                                
                                Set colComissao = New Collection
                                
                                colComissao.Add objComissao
                                
                                If iRelatorio = 0 Then
                                
                                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVGERCOMIINTDET", lNumIntTRVGerComiIntDet)
                                    If lErro <> SUCESSO Then gError 197466
                                    
                                End If
                                
                                lErro = TRVGerComiIntDet_Gravar_Comissoes(lComando1, 0, colComissao, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, colNumIntTRVVoucherInfo, lNumIntTRVGerComiIntDet, objTRVAcum.lNumIntInfoGer, iRelatorio, lNumIntRel, objTRVVoucher, objTRVAcum.lNumIntRelComiInt, iImprime, objPlanilha)
                                If lErro <> SUCESSO Then gError 197362
                                
                                iImprime = 0
                                    
                            End If
                                    
                        End If
                                
                    End If
                    
                Next
                
            End If
                
            For Each objTRVAcumAux In objTRVAcum.colTRVAcumAux
                
                Set objComissao = New ClassComissao
                objComissao.iCodVendedor = objTRVAcum.iVendedor
                objComissao.dPercentual = objTRVAcumAux.dPercComissao
                objComissao.dValorBase = objTRVAcumAux.dAcumValorBase
                objComissao.dValor = objComissao.dValorBase * objComissao.dPercentual
                
                Set colComissao = New Collection
                
                colComissao.Add objComissao
                
                If iRelatorio = 0 Then
                
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVGERCOMIINTDET", lNumIntTRVGerComiIntDet)
                    If lErro <> SUCESSO Then gError 197887
                    
                End If
                
                lErro = TRVGerComiIntDet_Gravar_Comissoes(lComando1, 0, colComissao, lNumIntTRVGerComiInt, lSeq, dtDataGeracao, colNumIntTRVVoucherInfo, lNumIntTRVGerComiIntDet, objTRVAcumAux.lNumIntInfoGer, iRelatorio, lNumIntRel, objTRVVoucher, objTRVAcumAux.lNumIntRelComiInt, iImprime, objPlanilha)
                If lErro <> SUCESSO Then gError 197888
                
            Next
                
        Next
        
    End If
    
    TRVGerComiIntDet_Gravar_Comissoes_Super = SUCESSO
    
    Exit Function
    
Erro_TRVGerComiIntDet_Gravar_Comissoes_Super:

    TRVGerComiIntDet_Gravar_Comissoes_Super = gErr

    Select Case gErr

        Case 197360 To 197362, 197465, 197466, 197887, 197888
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197363)

    End Select

    Exit Function

End Function

Private Function TRVGerComiIntDet_VendaDireta(ByVal iEmpresa As Integer, ByVal objTRVVoucher As ClassTRVVouchers, ByVal colNumIntTRVVoucherInfo As Collection, ByVal dValorBase As Double, ByVal dValorDolarBase As Double, ByVal dAporteValor As Double, ByVal iAporteMoeda As Integer, ByVal dCambio As Double, ByVal colTRVAcumGerente As Collection, ByVal colTRVAcumDiretor As Collection, ByVal iRelatorio As Long, ByVal lNumIntRel As Long, ByVal iCallCenter As Integer) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iGerente As Integer
Dim iVendedor As Integer
Dim iCargo As Integer
Dim iAchouGerente As Integer
Dim iAchouDiretor As Integer
Dim objTRVAcumGerente As ClassTRVAcumVendedor
Dim objTRVAcumDiretor As ClassTRVAcumVendedor
Dim iIndice As Integer
Dim alComando(1 To 5) As Long
Dim iRegiao As Integer
Dim iRegiaoPai As Integer
Dim iDiretor As Integer
Dim sComando_SQL As String
Dim iSuperior As Integer
Dim objTRVVoucherInfo As ClassTRVVoucherInfoN

On Error GoTo Erro_TRVGerComiIntDet_VendaDireta

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197374
    Next

    objCliente.lCodigo = objTRVVoucher.lCliente

    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO Then gError 197375

    If objCliente.iRegiao <> 0 Then
    
        iRegiao = objCliente.iRegiao
    
        Do While iRegiao <> 0 And (iDiretor = 0 Or iGerente = 0)
        
            sComando_SQL = "SELECT VendedorTRV.Vendedor, VendedorTRV.Cargo FROM VendedorTRVRegVenda, VendedorTRV WHERE VendedorTRVRegVenda.Vendedor = VendedorTRV.Vendedor " & _
                           " AND (VendedorTRV.Cargo = ? OR VendedorTRV.Cargo = ?) AND VendedorTRVRegVenda.RegiaoVenda = ?"
        
            lErro = Comando_Executar(alComando(1), sComando_SQL, iVendedor, iCargo, VENDEDOR_CARGO_TRV_GERENTE, VENDEDOR_CARGO_TRV_DIRETOR, iRegiao)
            If lErro <> AD_SQL_SUCESSO Then gError 197376

            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197377

            Do While lErro = AD_SQL_SUCESSO And (iGerente = 0 Or iDiretor = 0)
            
                If iCargo = VENDEDOR_CARGO_TRV_GERENTE And iGerente = 0 Then
                
                    iGerente = iVendedor
                
                    lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumGerente, iGerente, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, iCallCenter)
                    If lErro <> SUCESSO Then gError 197428

                    If iRelatorio = 0 Then
                    
                        For Each objTRVAcumGerente In colTRVAcumGerente
                        
                            If objTRVAcumGerente.iVendedor = iGerente Then
                    
                                For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
                        
                                    lErro = Comando_Executar(alComando(5), "INSERT INTO TRVVoucherInfoGer (NumIntVoucherInfo, NumIntInfoGer, Empresa) VALUES (?,?,?)", _
                                    objTRVVoucherInfo.lNumIntDoc, objTRVAcumGerente.lNumIntInfoGerGrava, iEmpresa)
                                    If lErro <> AD_SQL_SUCESSO Then gError 197884
                                    
                                Next
                                
                            End If
                    
                        Next
                    

                    End If


                    If iRelatorio = 2 Then

                        For Each objTRVAcumGerente In colTRVAcumGerente
                        
                            If objTRVAcumGerente.iVendedor = iGerente Then
                                
                                lErro = Comando_Executar(alComando(3), "INSERT INTO TRVRelVoucher (NumIntRel, Cliente, Produto, DiasAntc, TipVou, Serie, NumVou, AporteValor, AporteMoeda, TipoCliente, NumIntRelComiInt, Empresa, Data, DataFat) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                lNumIntRel, objTRVVoucher.lCliente, objTRVVoucher.sProduto, objTRVVoucher.idiasantc, objTRVVoucher.sTipVou, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, objTRVVoucher.dAporteValor, objTRVVoucher.iAporteMoeda, objTRVVoucher.iTipoCliente, objTRVAcumGerente.lNumIntRelComiIntGrava, iEmpresa, objTRVVoucher.dtData, objTRVVoucher.dtDataFat)
                                If lErro <> AD_SQL_SUCESSO Then gError 197488
                                
                                For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
                                
                                    lErro = Comando_Executar(alComando(4), "INSERT INTO TRVRelVoucherInfo (NumIntRel, NumIntDoc, TipoDoc, Valor, NumIntRelComiInt, Serie, NumVou, Empresa) VALUES (?,?,?,?,?,?,?,?)", _
                                    lNumIntRel, objTRVVoucherInfo.lNumIntDoc, objTRVVoucherInfo.sTipoDoc, objTRVVoucherInfo.dValor, objTRVAcumGerente.lNumIntRelComiIntGrava, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, iEmpresa)
                                    If lErro <> AD_SQL_SUCESSO Then gError 197489
            
                                Next
                                
                                Exit For
                            End If
    
                        Next

                    End If

                ElseIf iCargo = VENDEDOR_CARGO_TRV_DIRETOR And iDiretor = 0 Then

                    iDiretor = iVendedor
                
                    lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumDiretor, iDiretor, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, iCallCenter)
                    If lErro <> SUCESSO Then gError 197429

                    If iRelatorio = 0 Then
                    
                        For Each objTRVAcumDiretor In colTRVAcumDiretor
                        
                            If objTRVAcumDiretor.iVendedor = iDiretor Then
                    
                                For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
                        
                                    lErro = Comando_Executar(alComando(5), "INSERT INTO TRVVoucherInfoGer (NumIntVoucherInfo, NumIntInfoGer, Empresa) VALUES (?,?,?)", _
                                    objTRVVoucherInfo.lNumIntDoc, objTRVAcumDiretor.lNumIntInfoGerGrava, iEmpresa)
                                    If lErro <> AD_SQL_SUCESSO Then gError 197885
                                    
                                Next
                                
                            End If
                    
                        Next
                    

                    End If

                    If iRelatorio = 2 Then

                        For Each objTRVAcumDiretor In colTRVAcumDiretor
                        
                            If objTRVAcumDiretor.iVendedor = iDiretor Then
                                
                                lErro = Comando_Executar(alComando(3), "INSERT INTO TRVRelVoucher (NumIntRel, Cliente, Produto, DiasAntc, TipVou, Serie, NumVou, AporteValor, AporteMoeda, TipoCliente, NumIntRelComiInt, Empresa, Data, DataFat) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                lNumIntRel, objTRVVoucher.lCliente, objTRVVoucher.sProduto, objTRVVoucher.idiasantc, objTRVVoucher.sTipVou, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, objTRVVoucher.dAporteValor, objTRVVoucher.iAporteMoeda, objTRVVoucher.iTipoCliente, objTRVAcumDiretor.lNumIntRelComiIntGrava, iEmpresa, objTRVVoucher.dtData, objTRVVoucher.dtDataFat)
                                If lErro <> AD_SQL_SUCESSO Then gError 197490
                                
                                For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
                                
                                    lErro = Comando_Executar(alComando(4), "INSERT INTO TRVRelVoucherInfo (NumIntRel, NumIntDoc, TipoDoc, Valor, NumIntRelComiInt, Serie, NumVou, Empresa) VALUES (?,?,?,?,?,?,?,?)", _
                                    lNumIntRel, objTRVVoucherInfo.lNumIntDoc, objTRVVoucherInfo.sTipoDoc, objTRVVoucherInfo.dValor, objTRVAcumDiretor.lNumIntRelComiIntGrava, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, iEmpresa)
                                    If lErro <> AD_SQL_SUCESSO Then gError 197491
            
                                Next
                                
                                Exit For
                            End If
    
                        Next

                    End If

                End If
                
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197378
                
            Loop

            If iGerente = 0 Or iDiretor = 0 Then
                
                'pega o pai da regiao atual
                lErro = Comando_ExecutarPos(alComando(2), "SELECT CodRegiaoPai FROM TRVRegiaoVenda WHERE Codigo = ?", 0, iRegiaoPai, iRegiao)
                If lErro <> AD_SQL_SUCESSO Then gError 197379
    
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197380
                
                If lErro = AD_SQL_SUCESSO Then
                    
                    iRegiao = iRegiaoPai
                    
                Else
                
                    iRegiao = 0
                    
                End If
                
            End If

        Loop

    End If

    TRVGerComiIntDet_VendaDireta = SUCESSO
    
    Exit Function
    
Erro_TRVGerComiIntDet_VendaDireta:

    TRVGerComiIntDet_VendaDireta = gErr

    Select Case gErr

        Case 197374
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197375, 197428, 197429, 197884, 197885
        
        Case 197376, 197377, 197378
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORTRV", gErr)

        Case 197379, 197380
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVREGIAOVENDA", gErr)

        Case 197488, 197490
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVRELVOUCHER", gErr)
        
        Case 197489, 197491
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVRELVOUCHERINFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197381)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function TRVGerComiIntDet_Acum_Comissoes(ByVal lComando As Long, ByVal objTRVVoucher As ClassTRVVouchers, ByVal iPromotor As Integer, ByVal dValorBase As Double, ByVal dValorDolarBase As Double, ByVal dAporteValor As Double, ByVal iAporteMoeda As Integer, ByVal dCambio As Double, ByVal colTRVAcumPromotor As Collection, ByVal colTRVAcumSupervisor As Collection, ByVal colTRVAcumGerente As Collection, ByVal colTRVAcumDiretor As Collection, ByVal iRelatorio As Integer, ByVal lNumIntRel As Long, ByVal iCallCenter As Long) As Long

Dim lErro As Long
Dim iCargo As Integer
Dim iSuperior As Integer
Dim iVendedor As Integer
Dim iSupervisor As Integer
Dim iGerente As Integer
Dim iDiretor As Integer
Dim iRecursivo As Integer

On Error GoTo Erro_TRVGerComiIntDet_Acum_Comissoes

    'se for cliente MyAssistance
    If objTRVVoucher.lCliente > 1000000 Then

        'se for call center acumula os valores pois o promotor ganha um percentual fixo nestes casos
        lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumPromotor, iPromotor, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, 0)
        If lErro <> SUCESSO Then gError 199854

    ElseIf iCallCenter = 1 Then

        iVendedor = giVendedorCallCenter
    
        'se for call center acumula o valor no promotor call center pois ele ganha um percentual fixo
        lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumPromotor, iVendedor, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, 0)
        If lErro <> SUCESSO Then gError 197422

        'se for call center acumula os valores pois o promotor ganha um percentual fixo nestes casos
        lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumPromotor, iPromotor, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, iCallCenter)
        If lErro <> SUCESSO Then gError 199521

    End If

    lErro = Acum_Vend(lComando, iRecursivo, objTRVVoucher, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, colTRVAcumPromotor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, iCallCenter)
    If lErro <> SUCESSO Then gError 197423
        
    
    
'    objVendedor.iCodigo = iSuperior
'
'    lErro = CF("Vendedor_Le_Customizado1", objVendedor)
'    If lErro <> SUCESSO Then gError 199227
'
'    Set objVendedorTRV = objVendedor.objInfoUsu
'
'    iCargo = objVendedorTRV.iCargo
'
'    If iCargo = VENDEDOR_CARGO_TRV_SUPERVISOR Then
'
'        iSupervisor = iSuperior
'
'        lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumSupervisor, iSupervisor, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel)
'        If lErro <> SUCESSO Then gError 197423
'
'    End If
'
    
    
'    objVendedor.iCodigo = iSuperior
'
'    lErro = CF("Vendedor_Le_Customizado1", objVendedor)
'    If lErro <> SUCESSO Then gError 199228
'
'    Set objVendedorTRV = objVendedor.objInfoUsu
'
'    iCargo = objVendedorTRV.iCargo
'
'    If iCargo = VENDEDOR_CARGO_TRV_GERENTE Then
'
'        iGerente = iSuperior
'
'        lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumGerente, iGerente, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel)
'        If lErro <> SUCESSO Then gError 197424
'
'    End If
'
'    objVendedor.iCodigo = iSuperior
'
'    lErro = CF("Vendedor_Le_Customizado1", objVendedor)
'    If lErro <> SUCESSO Then gError 199229
'
'    Set objVendedorTRV = objVendedor.objInfoUsu
'
'    iCargo = objVendedorTRV.iCargo
'
'    If iCargo = VENDEDOR_CARGO_TRV_DIRETOR Then
'
'        iDiretor = iSuperior
'
'        lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumDiretor, iDiretor, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel)
'        If lErro <> SUCESSO Then gError 197425
'
'    End If
            
    TRVGerComiIntDet_Acum_Comissoes = SUCESSO
    
    Exit Function
    
Erro_TRVGerComiIntDet_Acum_Comissoes:

    TRVGerComiIntDet_Acum_Comissoes = gErr

    Select Case gErr

        Case 197422 To 197423, 199521, 199854
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197426)

    End Select

    Exit Function

End Function

Private Function Acum_Vendedor1(ByVal objTRVVoucher As ClassTRVVouchers, colTRVAcumVendedor As Collection, ByVal iVendedor As Integer, iSuperior As Integer, iCargo As Integer, ByVal dAporteValor As Double, ByVal dValorBase As Double, ByVal dValorDolarBase As Double, ByVal iAporteMoeda As Integer, ByVal dCambio As Double, ByVal iRelatorio As Integer, ByVal lNumIntRel As Long, ByVal iCallCenter As Integer) As Long

Dim lErro As Long
Dim iAchouVendedor As Integer
Dim objTRVAcumVendedor As ClassTRVAcumVendedor
Dim objVendedor As New ClassVendedor
Dim objVendedorTRV As ClassVendedorTRV
Dim lNumIntInfoGer As Long
Dim dPercentualReducao As Double
Dim objVNDReducaoTRV As ClassVNDReducaoTRV
Dim objTRVAcumAporte As ClassTRVAcumAporte
Dim iAporte As Integer
Dim dValorDe As Double
Dim dValorAte As Double
Dim dAporteValorAux As Double
Dim lNumIntRelComiInt As Long
Dim objCliente As New ClassCliente
Dim sComando_SQL As String
Dim iIndice As Integer
Dim alComando(1 To 1) As Long
Dim iRegiao As Integer
Dim dPercComissao As Double
Dim iAchouPerc As Integer
Dim objTRVAcumAux As ClassTRVAcumAux

On Error GoTo Erro_Acum_Vendedor1

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 199417
    Next


    If iVendedor <> 0 Then

        For Each objTRVAcumVendedor In colTRVAcumVendedor
        
            If objTRVAcumVendedor.iVendedor = iVendedor Then
                
                iAchouVendedor = 1
                
                Exit For
                    
            End If
            
        Next
        
        If iAchouVendedor = 0 Then
        
            objVendedor.iCodigo = iVendedor
            
            lErro = CF("Vendedor_Le", objVendedor)
            If lErro <> SUCESSO And lErro <> 12582 Then gError 199584
            
            If lErro <> SUCESSO Then gError 199585
            
            lErro = CF("Vendedor_Le_Customizado1", objVendedor)
            If lErro <> SUCESSO Then gError 199418
            
            Set objVendedorTRV = objVendedor.objInfoUsu
        
            Set objTRVAcumVendedor = New ClassTRVAcumVendedor
            objTRVAcumVendedor.iVendedor = iVendedor
            objTRVAcumVendedor.iCargo = objVendedorTRV.iCargo
            objTRVAcumVendedor.iSuperior = objVendedorTRV.iSuperior
            objTRVAcumVendedor.dPercCallCenter = objVendedorTRV.dPercCallCenter
            objTRVAcumVendedor.sNomeReduzido = objVendedor.sNomeReduzido
            
            If iRelatorio = 0 Then
            
                'para tratar vouchers relacionados a myassistance
                lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOUCHERINFOGER", lNumIntInfoGer)
                If lErro <> SUCESSO Then gError 199855
        
                objTRVAcumVendedor.lNumIntInfoGerMyAssistance = lNumIntInfoGer
                    
            ElseIf iRelatorio = 2 Then
                    
                lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_PROX_TRVRELCOMIINT", lNumIntRelComiInt)
                If lErro <> SUCESSO Then gError 199856
                
                objTRVAcumVendedor.lNumIntRelComiIntMyAssistance = lNumIntRelComiInt
                
            End If
            
            If objTRVAcumVendedor.iCargo <> 0 Then
            
                If iRelatorio = 0 Then
            
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOUCHERINFOGER", lNumIntInfoGer)
                    If lErro <> SUCESSO Then gError 199416
            
                    objTRVAcumVendedor.lNumIntInfoGer = lNumIntInfoGer
                    
                    'para tratar vouchers relacionados a call center
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOUCHERINFOGER", lNumIntInfoGer)
                    If lErro <> SUCESSO Then gError 199413
            
                    objTRVAcumVendedor.lNumIntInfoGerCallCenter = lNumIntInfoGer
                    
                End If
                
                If iRelatorio = 2 Then
                
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_PROX_TRVRELCOMIINT", lNumIntRelComiInt)
                    If lErro <> SUCESSO Then gError 199415
                    
                    objTRVAcumVendedor.lNumIntRelComiInt = lNumIntRelComiInt
                
                    lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_PROX_TRVRELCOMIINT", lNumIntRelComiInt)
                    If lErro <> SUCESSO Then gError 199414
                    
                    objTRVAcumVendedor.lNumIntRelComiIntCallCenter = lNumIntRelComiInt
                
                End If
                
            End If
            
            'serve para indicar que um determinado voucher esta associado a aporte
            objTRVAcumVendedor.iAporte = 0
            
            objTRVAcumVendedor.lNumIntInfoGerGrava = objTRVAcumVendedor.lNumIntInfoGer
            objTRVAcumVendedor.lNumIntRelComiIntGrava = objTRVAcumVendedor.lNumIntRelComiInt
            
            For Each objVNDReducaoTRV In objVendedorTRV.colVNDReducaoTRV
                
                Set objTRVAcumAporte = New ClassTRVAcumAporte
                
                objTRVAcumAporte.dPercentual = objVNDReducaoTRV.dPercComissaoMax
                objTRVAcumAporte.dValorDe = objVNDReducaoTRV.dValorDe
                objTRVAcumAporte.dValorAte = objVNDReducaoTRV.dValorAte
                objTRVAcumAporte.iMoeda = objVNDReducaoTRV.iMoeda
                
                objTRVAcumVendedor.colAcumAporte.Add objTRVAcumAporte
            
            Next
            
            colTRVAcumVendedor.Add objTRVAcumVendedor
            
        End If
    
        If objTRVVoucher.lCliente > 1000000 Then
        
            objTRVAcumVendedor.dValorAcumMyAssistance = objTRVAcumVendedor.dValorAcumMyAssistance + dValorBase
            objTRVAcumVendedor.lNumIntInfoGerGrava = objTRVAcumVendedor.lNumIntInfoGerMyAssistance
            objTRVAcumVendedor.lNumIntRelComiIntGrava = objTRVAcumVendedor.lNumIntRelComiIntMyAssistance
        
        ElseIf iCallCenter = 1 Then
            objTRVAcumVendedor.dValorAcumCallCenter = objTRVAcumVendedor.dValorAcumCallCenter + dValorBase
            objTRVAcumVendedor.lNumIntInfoGerGrava = objTRVAcumVendedor.lNumIntInfoGerCallCenter
            objTRVAcumVendedor.lNumIntRelComiIntGrava = objTRVAcumVendedor.lNumIntRelComiIntCallCenter
        Else
    
    
            'se nao for um promotor (é um supervisor, gerente ou diretor) e na regiao do cliente o vendedor
            'tiver valor fixo de percentual, este prevalece sobre os demais criterios para este profissional
            If objTRVAcumVendedor.iCargo <> 0 And iVendedor <> giVendedorCallCenter Then
            
                objCliente.lCodigo = objTRVVoucher.lCliente
    
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO Then gError 199419
    
                iRegiao = objCliente.iRegiao
                
                If iRegiao <> 0 Then
                
                    sComando_SQL = "SELECT PercComissao FROM VendedorTRVRegVenda WHERE Vendedor = ? And RegiaoVenda = ?"
                
                    lErro = Comando_ExecutarPos(alComando(1), sComando_SQL, 0, dPercComissao, iVendedor, iRegiao)
                    If lErro <> AD_SQL_SUCESSO Then gError 199420
        
                    lErro = Comando_BuscarPrimeiro(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199421
                
                    If dPercComissao > 0 Then
                    
                        For Each objTRVAcumAux In objTRVAcumVendedor.colTRVAcumAux
                        
                            If objTRVAcumAux.dPercComissao = dPercComissao Then
                            
                                iAchouPerc = 1
                                
                                Exit For
                    
                            End If
                            
                        Next
                        
                        If iAchouPerc = 0 Then
                        
                            Set objTRVAcumAux = New ClassTRVAcumAux
                            
                            objTRVAcumAux.dPercComissao = dPercComissao
                    
                            objTRVAcumAux.dAcumValorBase = dValorBase
                            
                            If iRelatorio = 0 Then
                        
                                lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_INT_PROX_TRVVOUCHERINFOGER", lNumIntInfoGer)
                                If lErro <> SUCESSO Then gError 199422
                        
                                objTRVAcumAux.lNumIntInfoGer = lNumIntInfoGer
                                
                            End If
                            
                            If iRelatorio = 2 Then
                            
                                lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_PROX_TRVRELCOMIINT", lNumIntRelComiInt)
                                If lErro <> SUCESSO Then gError 199423
                                
                                objTRVAcumAux.lNumIntRelComiInt = lNumIntRelComiInt
                            
                            End If
                    
                            objTRVAcumVendedor.colTRVAcumAux.Add objTRVAcumAux
                            
                            iAchouPerc = 1
                    
                        Else
                        
                            objTRVAcumAux.dAcumValorBase = objTRVAcumAux.dAcumValorBase + dValorBase
        
                        End If
        
        
                    End If
                End If
        
            End If
        
            If iAchouPerc = 0 Then
        
                For Each objTRVAcumAporte In objTRVAcumVendedor.colAcumAporte
                
                    dAporteValorAux = dAporteValor
                    dValorDe = objTRVAcumAporte.dValorDe
                    dValorAte = objTRVAcumAporte.dValorAte
                    
                    If iAporteMoeda <> objTRVAcumAporte.iMoeda Then
                        If iAporteMoeda = MOEDA_DOLAR Then
                            dAporteValorAux = dAporteValorAux * dCambio
                        Else
                            dValorDe = dValorDe * dCambio
                            dValorAte = dValorAte * dCambio
                        End If
                    End If
                
                    If dAporteValorAux >= dValorDe And (dAporteValorAux <= dValorAte Or dValorAte = 0) Then
                    
                        If (objTRVAcumAporte.dValorDolarAcumAporte + dValorDolarBase) <> 0 Then
                            objTRVAcumAporte.dCambioMedio = ((objTRVAcumAporte.dValorDolarAcumAporte * objTRVAcumAporte.dCambioMedio) + dValorDolarBase * dCambio) / (objTRVAcumAporte.dValorDolarAcumAporte + dValorDolarBase)
                        End If
                        
                        objTRVAcumAporte.dValorAcumAporte = objTRVAcumAporte.dValorAcumAporte + dValorBase
                        objTRVAcumAporte.dValorDolarAcumAporte = objTRVAcumAporte.dValorDolarAcumAporte + dValorDolarBase
                        iAporte = 1
                        
                        'serve para indicar que um determinado voucher esta associado a aporte
                        objTRVAcumVendedor.iAporte = 1
                        
                        Exit For
                    End If
                Next
            
                If iAporte = 0 Then
                    If (objTRVAcumVendedor.dValorDolarAcum + dValorDolarBase) <> 0 Then
                        objTRVAcumVendedor.dCambioMedio = ((objTRVAcumVendedor.dValorDolarAcum * objTRVAcumVendedor.dCambioMedio) + dValorDolarBase * dCambio) / (objTRVAcumVendedor.dValorDolarAcum + dValorDolarBase)
                    End If
                    objTRVAcumVendedor.dValorAcum = objTRVAcumVendedor.dValorAcum + dValorBase
                    objTRVAcumVendedor.dValorDolarAcum = objTRVAcumVendedor.dValorDolarAcum + dValorDolarBase
                End If
        
            Else
                objTRVAcumVendedor.lNumIntInfoGerGrava = objTRVAcumAux.lNumIntInfoGer
                objTRVAcumVendedor.lNumIntRelComiIntGrava = objTRVAcumAux.lNumIntRelComiInt
            End If
        
        End If
    
        iCargo = objTRVAcumVendedor.iCargo
        iSuperior = objTRVAcumVendedor.iSuperior
    
    Else
        iCargo = 0
        iSuperior = 0
    
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Acum_Vendedor1 = SUCESSO
    
    Exit Function
    
Erro_Acum_Vendedor1:

    Acum_Vendedor1 = gErr

    Select Case gErr

        Case 199413 To 199423, 199584, 199855, 199856
        
        Case 199585
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199424)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Rotina_GerComiInt(ByVal sNomeArqParam As String, ByVal objTRVGerComiInt As ClassTRVGerComiInt, Optional ByVal iRelatorio As Integer = 0, Optional lNumIntRel As Long) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_GerComiInt


    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 197415

    Set gcolModulo = New AdmColModulo
    
    lErro = CF("Modulos_Le_Empresa_Filial", glEmpresa, giFilialEmpresa, gcolModulo)
    If lErro <> SUCESSO Then gError 197416

    lErro = CF("Retorna_ColFiliais")
    If lErro <> SUCESSO Then gError 197417

    GL_lUltimoErro = SUCESSO

    TelaAcompanhaBatchFAT1.sNomeArqParam = sNomeArqParam
    Set TelaAcompanhaBatchFAT1.gobjTRVGerComiInt = objTRVGerComiInt
    TelaAcompanhaBatchFAT1.iRotinaBatch = ROTINA_TRVGERACOMIINT_BATCH
    TelaAcompanhaBatchFAT1.giRelatorio = iRelatorio

    TelaAcompanhaBatchFAT1.Show

    Rotina_GerComiInt = SUCESSO
    
    Exit Function

Erro_Rotina_GerComiInt:

    Rotina_GerComiInt = gErr

    Select Case gErr

        Case 197415, 197416, 197417
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197418)

    End Select
    
    Exit Function

End Function

Private Function TRVVoucher_Le_Aportes(ByVal lComando As Long, ByVal objTRVVoucher As ClassTRVVouchers, ByVal lCliente As Long) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iMoeda As Integer
Dim dValor As Double
Dim sComando_SQL As String

On Error GoTo Erro_TRVVoucher_Le_Aportes

    sComando_SQL = "SELECT Moeda, Valor FROM TRVAportes, TRVAportePagtoFat WHERE Cliente = ? AND TRVAportes.NumIntDoc = NumIntDocAporte AND ValidadeDe <= ? AND ValidadeAte >= ? "

    lErro = Comando_Executar(lComando, sComando_SQL, iMoeda, dValor, lCliente, objTRVVoucher.dtData, objTRVVoucher.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 197880
  
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197881

    Do While lErro = AD_SQL_SUCESSO
    
        If iMoeda = MOEDA_DOLAR Then
            objTRVVoucher.dAporteValor = objTRVVoucher.dAporteValor + dValor * objTRVVoucher.dCambio
        Else
            objTRVVoucher.dAporteValor = objTRVVoucher.dAporteValor + dValor
        End If
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197882
    
    Loop

    sComando_SQL = "SELECT Moeda, Valor FROM TRVAportes, TRVAportePagtoDireto WHERE Cliente = ? AND TRVAportes.NumIntDoc = NumIntDocAporte AND DataEmissao <= ? AND DataEmissao+365 >= ? "

    lErro = Comando_Executar(lComando, sComando_SQL, iMoeda, dValor, lCliente, objTRVVoucher.dtData, objTRVVoucher.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 199392

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199393

    Do While lErro = AD_SQL_SUCESSO

        If iMoeda = MOEDA_DOLAR Then
            objTRVVoucher.dAporteValor = objTRVVoucher.dAporteValor + dValor * objTRVVoucher.dCambio
        Else
            objTRVVoucher.dAporteValor = objTRVVoucher.dAporteValor + dValor
        End If

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199394

    Loop
                
    TRVVoucher_Le_Aportes = SUCESSO
    
    Exit Function
    
Erro_TRVVoucher_Le_Aportes:

    TRVVoucher_Le_Aportes = gErr

    Select Case gErr

        Case 197869
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197880 To 197882, 199392 To 199394
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVAPORTES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197883)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()
    Set gobjTRVProdTafCalc = Nothing
    Set gobjTRVVoucher = Nothing
End Sub

Private Function Acum_Vend(ByVal lComando As Long, ByVal iRecursivo As Integer, ByVal objTRVVoucher As ClassTRVVouchers, ByVal iSuperior As Integer, ByVal iCargo As Integer, ByVal dAporteValor As Double, ByVal dValorBase As Double, ByVal dValorDolarBase As Double, ByVal iAporteMoeda As Integer, ByVal dCambio As Double, ByVal iRelatorio As Integer, ByVal lNumIntRel As Long, ByVal colTRVAcumPromotor As Collection, ByVal colTRVAcumSupervisor As Collection, ByVal colTRVAcumGerente As Collection, ByVal colTRVAcumDiretor As Collection, ByVal iCallCenter As Integer) As Long

Dim objVendedor As New ClassVendedor
Dim objVendedorTRV As ClassVendedorTRV
Dim lErro As Long
Dim vVendedor As Variant
Dim colVendedores As New Collection

On Error GoTo Erro_Acum_Vend

''''    If iSuperior <> 0 Then
''''
''''        iRecursivo = iRecursivo + 1
''''
''''        If iRecursivo > 10 Then gError 197424
''''
''''        objVendedor.iCodigo = iSuperior
''''
''''        lErro = CF("Vendedor_Le_Customizado1", objVendedor)
''''        If lErro <> SUCESSO Then gError 197425
''''
''''        Set objVendedorTRV = objVendedor.objInfoUsu
''''
''''        iCargo = objVendedorTRV.iCargo
''''
''''        Select Case iCargo
''''
''''            Case VENDEDOR_CARGO_TRV_PROMOTOR
''''
''''                lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumPromotor, iSuperior, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, iCallCenter)
''''                If lErro <> SUCESSO Then gError 199384
''''
''''            Case VENDEDOR_CARGO_TRV_SUPERVISOR
''''
''''                lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumSupervisor, iSuperior, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, iCallCenter)
''''                If lErro <> SUCESSO Then gError 199227
''''
''''            Case VENDEDOR_CARGO_TRV_GERENTE
''''
''''                lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumGerente, iSuperior, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, iCallCenter)
''''                If lErro <> SUCESSO Then gError 199228
''''
''''            Case VENDEDOR_CARGO_TRV_DIRETOR
''''
''''                lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumDiretor, iSuperior, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, iCallCenter)
''''                If lErro <> SUCESSO Then gError 199229
''''
''''        End Select
''''
''''        lErro = Acum_Vend(iRecursivo, objTRVVoucher, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, colTRVAcumPromotor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, iCallCenter)
''''        If lErro <> SUCESSO Then gError 199382
''''
''''    End If
    
    
    lErro = CF("TRVVouVendedores_Le", colVendedores, objTRVVoucher, lComando)
    If lErro <> SUCESSO Then gError 199519
    
    For Each objVendedor In colVendedores
        
        lErro = Acum_Vendedor1(objTRVVoucher, colTRVAcumPromotor, objVendedor.iCodigo, iSuperior, iCargo, dAporteValor, dValorBase, dValorDolarBase, iAporteMoeda, dCambio, iRelatorio, lNumIntRel, iCallCenter)
        If lErro <> SUCESSO Then gError 199228
        
    Next
    
    
    Acum_Vend = SUCESSO
    
    Exit Function
    
Erro_Acum_Vend:

    Acum_Vend = gErr

    Select Case gErr

        Case 197424
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_RECURSIVIDADE", gErr)

        Case 197425, 199227 To 199229, 199382, 199384, 199519

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199383)

    End Select

    Exit Function

End Function

Private Function TRVRelVoucher_Gravar_1(ByVal iCargo As Integer, ByVal iRecursivo As Integer, ByVal iEmpresa As Integer, ByVal iVendedor As Integer, ByVal colNumIntTRVVoucherInfo As Collection, ByVal iRelatorio As Integer, ByVal lNumIntRel As Long, ByVal objTRVVoucher As ClassTRVVouchers, ByVal colTRVAcumPromotor As Collection, ByVal colTRVAcumSupervisor As Collection, ByVal colTRVAcumGerente As Collection, ByVal colTRVAcumDiretor As Collection, ByVal dValorBase As Double)

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim objTRVAcumPromotor As ClassTRVAcumVendedor
Dim objTRVAcumSupervisor As ClassTRVAcumVendedor
Dim objTRVAcumGerente As ClassTRVAcumVendedor
Dim objTRVAcumDiretor As ClassTRVAcumVendedor
Dim objVendedor As New ClassVendedor
Dim objVendedorTRV As ClassVendedorTRV
Dim objTRVVoucherInfo As ClassTRVVoucherInfoN


On Error GoTo Erro_TRVRelVoucher_Gravar_1

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197485
    Next

    For Each objTRVAcumPromotor In colTRVAcumPromotor

        If objTRVAcumPromotor.iVendedor = iVendedor Then

            lErro = Comando_Executar(alComando(1), "INSERT INTO TRVRelVoucher (NumIntRel, Cliente, Produto, DiasAntc, TipVou, Serie, NumVou, AporteValor, AporteMoeda, TipoCliente, NumIntRelComiInt, Empresa, Data, DataFat, Aporte, ValorBase) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntRel, objTRVVoucher.lCliente, objTRVVoucher.sProduto, objTRVVoucher.idiasantc, objTRVVoucher.sTipVou, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, objTRVVoucher.dAporteValor, objTRVVoucher.iAporteMoeda, objTRVVoucher.iTipoCliente, objTRVAcumPromotor.lNumIntRelComiIntGrava, iEmpresa, objTRVVoucher.dtData, objTRVVoucher.dtDataFat, objTRVVoucher.iAporte, dValorBase)
            If lErro <> AD_SQL_SUCESSO Then gError 199389

            For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo

                lErro = Comando_Executar(alComando(2), "INSERT INTO TRVRelVoucherInfo (NumIntRel, NumIntDoc, TipoDoc, Valor, NumIntRelComiInt, Serie, NumVou, Empresa) VALUES (?,?,?,?,?,?,?,?)", _
                lNumIntRel, objTRVVoucherInfo.lNumIntDoc, objTRVVoucherInfo.sTipoDoc, objTRVVoucherInfo.dValor, objTRVAcumPromotor.lNumIntRelComiIntGrava, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, iEmpresa)
                If lErro <> AD_SQL_SUCESSO Then gError 199390

            Next
            
            Exit For

        End If

    Next


''''    If iSuperior <> 0 Then
''''
''''        iRecursivo = iRecursivo + 1
''''
''''        If iRecursivo > 10 Then gError 199385
    
    
    
''''        Select Case iCargo
''''
''''            Case VENDEDOR_CARGO_TRV_PROMOTOR
''''
''''                For Each objTRVAcumPromotor In colTRVAcumPromotor
''''
''''                    If objTRVAcumPromotor.iVendedor = iSuperior Then
''''                        iSuperior = objTRVAcumPromotor.iSuperior
''''                        objVendedor.iCodigo = iSuperior
''''
''''                        If iSuperior <> 0 Then
''''
''''                            lErro = CF("Vendedor_Le_Customizado1", objVendedor)
''''                            If lErro <> SUCESSO Then gError 199388
''''
''''                            Set objVendedorTRV = objVendedor.objInfoUsu
''''
''''                            iCargo = objVendedorTRV.iCargo
''''
''''                        Else
''''
''''                            iCargo = 0
''''                        End If
''''
''''                        lErro = Comando_Executar(alComando(1), "INSERT INTO TRVRelVoucher (NumIntRel, Cliente, Produto, DiasAntc, TipVou, Serie, NumVou, AporteValor, AporteMoeda, TipoCliente, NumIntRelComiInt, Empresa, Data, DataFat, Aporte, ValorBase) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
''''                        lNumIntRel, objTRVVoucher.lCliente, objTRVVoucher.sProduto, objTRVVoucher.idiasantc, objTRVVoucher.sTipVou, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, objTRVVoucher.dAporteValor, objTRVVoucher.iAporteMoeda, objTRVVoucher.iTipoCliente, objTRVAcumPromotor.lNumIntRelComiIntGrava, iEmpresa, objTRVVoucher.dtData, objTRVVoucher.dtdatafat, objTRVVoucher.iAporte, dValorBase)
''''                        If lErro <> AD_SQL_SUCESSO Then gError 199389
''''
''''                        For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
''''
''''                            lErro = Comando_Executar(alComando(2), "INSERT INTO TRVRelVoucherInfo (NumIntRel, NumIntDoc, TipoDoc, Valor, NumIntRelComiInt, Serie, NumVou, Empresa) VALUES (?,?,?,?,?,?,?,?)", _
''''                            lNumIntRel, objTRVVoucherInfo.lNumIntDoc, objTRVVoucherInfo.sTipoDoc, objTRVVoucherInfo.dValor, objTRVAcumPromotor.lNumIntRelComiIntGrava, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, iEmpresa)
''''                            If lErro <> AD_SQL_SUCESSO Then gError 199390
''''
''''                        Next
''''
''''                        Exit For
''''
''''                    End If
''''
''''                Next
''''
''''            Case VENDEDOR_CARGO_TRV_SUPERVISOR
''''
''''                For Each objTRVAcumSupervisor In colTRVAcumSupervisor
''''
''''                    If objTRVAcumSupervisor.iVendedor = iSuperior Then
''''                        iSuperior = objTRVAcumSupervisor.iSuperior
''''                        objVendedor.iCodigo = iSuperior
''''
''''                        If iSuperior <> 0 Then
''''
''''                            lErro = CF("Vendedor_Le_Customizado1", objVendedor)
''''                            If lErro <> SUCESSO Then gError 197463
''''
''''                            Set objVendedorTRV = objVendedor.objInfoUsu
''''
''''                            iCargo = objVendedorTRV.iCargo
''''
''''                        Else
''''                            iCargo = 0
''''                        End If
''''
''''
''''                        lErro = Comando_Executar(alComando(1), "INSERT INTO TRVRelVoucher (NumIntRel, Cliente, Produto, DiasAntc, TipVou, Serie, NumVou, AporteValor, AporteMoeda, TipoCliente, NumIntRelComiInt, Empresa, Data, DataFat, Aporte, ValorBase) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
''''                        lNumIntRel, objTRVVoucher.lCliente, objTRVVoucher.sProduto, objTRVVoucher.idiasantc, objTRVVoucher.sTipVou, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, objTRVVoucher.dAporteValor, objTRVVoucher.iAporteMoeda, objTRVVoucher.iTipoCliente, objTRVAcumSupervisor.lNumIntRelComiIntGrava, iEmpresa, objTRVVoucher.dtData, objTRVVoucher.dtdatafat, objTRVVoucher.iAporte, dValorBase)
''''                        If lErro <> AD_SQL_SUCESSO Then gError 197476
''''
''''                        For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
''''
''''                            lErro = Comando_Executar(alComando(2), "INSERT INTO TRVRelVoucherInfo (NumIntRel, NumIntDoc, TipoDoc, Valor, NumIntRelComiInt, Serie, NumVou, Empresa) VALUES (?,?,?,?,?,?,?,?)", _
''''                            lNumIntRel, objTRVVoucherInfo.lNumIntDoc, objTRVVoucherInfo.sTipoDoc, objTRVVoucherInfo.dValor, objTRVAcumSupervisor.lNumIntRelComiIntGrava, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, iEmpresa)
''''                            If lErro <> AD_SQL_SUCESSO Then gError 197477
''''
''''                        Next
''''
''''                        Exit For
''''
''''                    End If
''''
''''                Next
''''
''''            Case VENDEDOR_CARGO_TRV_GERENTE
''''
''''                For Each objTRVAcumGerente In colTRVAcumGerente
''''
''''                    If objTRVAcumGerente.iVendedor = iSuperior Then
''''                        iSuperior = objTRVAcumGerente.iSuperior
''''                        objVendedor.iCodigo = iSuperior
''''
''''                        If iSuperior <> 0 Then
''''
''''                            lErro = CF("Vendedor_Le_Customizado1", objVendedor)
''''                            If lErro <> SUCESSO Then gError 197464
''''
''''                            Set objVendedorTRV = objVendedor.objInfoUsu
''''
''''                            iCargo = objVendedorTRV.iCargo
''''
''''                        Else
''''                            iCargo = 0
''''                        End If
''''
''''                        lErro = Comando_Executar(alComando(1), "INSERT INTO TRVRelVoucher (NumIntRel, Cliente, Produto, DiasAntc, TipVou, Serie, NumVou, AporteValor, AporteMoeda, TipoCliente, NumIntRelComiInt, Empresa, Data, DataFat, Aporte, ValorBase) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
''''                        lNumIntRel, objTRVVoucher.lCliente, objTRVVoucher.sProduto, objTRVVoucher.idiasantc, objTRVVoucher.sTipVou, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, objTRVVoucher.dAporteValor, objTRVVoucher.iAporteMoeda, objTRVVoucher.iTipoCliente, objTRVAcumGerente.lNumIntRelComiIntGrava, iEmpresa, objTRVVoucher.dtData, objTRVVoucher.dtdatafat, objTRVVoucher.iAporte, dValorBase)
''''                        If lErro <> AD_SQL_SUCESSO Then gError 197478
''''
''''                        For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
''''
''''                            lErro = Comando_Executar(alComando(2), "INSERT INTO TRVRelVoucherInfo (NumIntRel, NumIntDoc, TipoDoc, Valor, NumIntRelComiInt, Serie, NumVou, Empresa) VALUES (?,?,?,?,?,?,?,?)", _
''''                            lNumIntRel, objTRVVoucherInfo.lNumIntDoc, objTRVVoucherInfo.sTipoDoc, objTRVVoucherInfo.dValor, objTRVAcumGerente.lNumIntRelComiIntGrava, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, iEmpresa)
''''                            If lErro <> AD_SQL_SUCESSO Then gError 197479
''''
''''                        Next
''''
''''                        Exit For
''''                    End If
''''
''''                Next
''''
''''
''''            Case VENDEDOR_CARGO_TRV_DIRETOR
''''
''''                For Each objTRVAcumDiretor In colTRVAcumDiretor
''''
''''                    If objTRVAcumDiretor.iVendedor = iSuperior Then
''''
''''                        iSuperior = objTRVAcumDiretor.iSuperior
''''                        objVendedor.iCodigo = iSuperior
''''
''''                        If iSuperior <> 0 Then
''''
''''                            lErro = CF("Vendedor_Le_Customizado1", objVendedor)
''''                            If lErro <> SUCESSO Then gError 199391
''''
''''                            Set objVendedorTRV = objVendedor.objInfoUsu
''''
''''                            iCargo = objVendedorTRV.iCargo
''''
''''                        Else
''''                            iCargo = 0
''''                        End If
''''
''''                        lErro = Comando_Executar(alComando(1), "INSERT INTO TRVRelVoucher (NumIntRel, Cliente, Produto, DiasAntc, TipVou, Serie, NumVou, AporteValor, AporteMoeda, TipoCliente, NumIntRelComiInt, Empresa, Data, DataFat, Aporte, ValorBase) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
''''                        lNumIntRel, objTRVVoucher.lCliente, objTRVVoucher.sProduto, objTRVVoucher.idiasantc, objTRVVoucher.sTipVou, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, objTRVVoucher.dAporteValor, objTRVVoucher.iAporteMoeda, objTRVVoucher.iTipoCliente, objTRVAcumDiretor.lNumIntRelComiIntGrava, iEmpresa, objTRVVoucher.dtData, objTRVVoucher.dtdatafat, objTRVVoucher.iAporte, dValorBase)
''''                        If lErro <> AD_SQL_SUCESSO Then gError 197480
''''
''''                        For Each objTRVVoucherInfo In colNumIntTRVVoucherInfo
''''
''''                            lErro = Comando_Executar(alComando(2), "INSERT INTO TRVRelVoucherInfo (NumIntRel, NumIntDoc, TipoDoc, Valor, NumIntRelComiInt, Serie, NumVou, Empresa) VALUES (?,?,?,?,?,?,?,?)", _
''''                            lNumIntRel, objTRVVoucherInfo.lNumIntDoc, objTRVVoucherInfo.sTipoDoc, objTRVVoucherInfo.dValor, objTRVAcumDiretor.lNumIntRelComiIntGrava, objTRVVoucher.sSerie, objTRVVoucher.lNumVou, iEmpresa)
''''                            If lErro <> AD_SQL_SUCESSO Then gError 197481
''''
''''                        Next
''''
''''                        Exit For
''''                    End If
''''
''''                Next
''''
''''        End Select
''''
''''        If iRecursivo > 0 Then
''''
''''            lErro = TRVRelVoucher_Gravar_1(iCargo, iRecursivo, iEmpresa, iSuperior, colNumIntTRVVoucherInfo, iRelatorio, lNumIntRel, objTRVVoucher, colTRVAcumPromotor, colTRVAcumSupervisor, colTRVAcumGerente, colTRVAcumDiretor, dValorBase)
''''            If lErro <> SUCESSO Then gError 199387
''''
''''        End If
''''
''''    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TRVRelVoucher_Gravar_1 = SUCESSO
    
    Exit Function
    
Erro_TRVRelVoucher_Gravar_1:

    TRVRelVoucher_Gravar_1 = gErr
    
    Select Case gErr

        Case 197463, 197464, 199388, 199391, 199513
        
        Case 197472, 197476, 197478, 197480, 199389
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVRELVOUCHER", gErr)
        
        Case 197477, 197479, 197481, 199390
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVRELVOUCHERINFO", gErr)
        
        Case 197485
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 199385
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_RECURSIVIDADE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199386)


    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Exporta_Cabecalho_Excel(ByVal objPlanilha As ClassPlanilhaExcel) As Long
'Transfere os dados de objBrowse para objPlanilha no formato esperado pelo Excel para geração da planilha
'objBrowse RECEBE(Input) os dados que serão transferidos
'objPlanilha RETORNA(Output) os dados no formato esperado pelo Excel

Dim iColuna As Integer
Dim objColunas As ClassColunasExcel
Dim objCelulas As ClassCelulasExcel
Dim dLarguraColuna As Double
Dim lErro As Long
Dim iIndiceCol As Integer
Dim iIndiceLin As Integer
'##################################
'Inserido por Wagner 26/01/2006
Dim iIndiceColdeFato As Integer
Dim iIndice As Integer
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim vValorCampo As Variant
Dim iExercicio As Integer
Dim iIndiceColAux As Integer
'##################################

On Error GoTo Erro_Exporta_Cabecalho_Excel

    'Guarda o nome que será exibido para a planilha
    objPlanilha.sNomePlanilha = "Comissionamento Interno"
    
    lErro = Cria_Coluna_Excel(objPlanilha, "Est")
    If lErro <> SUCESSO Then gError 199593
    
    lErro = Cria_Coluna_Excel(objPlanilha, "Prom")
    If lErro <> SUCESSO Then gError 199594
    
    lErro = Cria_Coluna_Excel(objPlanilha, "Sup")
    If lErro <> SUCESSO Then gError 199595
    
    lErro = Cria_Coluna_Excel(objPlanilha, "Ger")
    If lErro <> SUCESSO Then gError 199596
    
    lErro = Cria_Coluna_Excel(objPlanilha, "Dir")
    If lErro <> SUCESSO Then gError 199597
    
    lErro = Cria_Coluna_Excel(objPlanilha, "CodEmp")
    If lErro <> SUCESSO Then gError 199598
    
    lErro = Cria_Coluna_Excel(objPlanilha, "Fantasia")
    If lErro <> SUCESSO Then gError 199599
    
    lErro = Cria_Coluna_Excel(objPlanilha, "Produto")
    If lErro <> SUCESSO Then gError 199600
    
    lErro = Cria_Coluna_Excel(objPlanilha, "PercProm")
    If lErro <> SUCESSO Then gError 199601
    
    lErro = Cria_Coluna_Excel(objPlanilha, "NumVou")
    If lErro <> SUCESSO Then gError 199602
    
    lErro = Cria_Coluna_Excel(objPlanilha, "DatEmi")
    If lErro <> SUCESSO Then gError 199603
    
    lErro = Cria_Coluna_Excel(objPlanilha, "DatPag")
    If lErro <> SUCESSO Then gError 199604
    
    lErro = Cria_Coluna_Excel(objPlanilha, "NumFat")
    If lErro <> SUCESSO Then gError 199605
    
    lErro = Cria_Coluna_Excel(objPlanilha, "Moeda")
    If lErro <> SUCESSO Then gError 199606
    
    lErro = Cria_Coluna_Excel(objPlanilha, "TarifMoe")
    If lErro <> SUCESSO Then gError 199607
    
    lErro = Cria_Coluna_Excel(objPlanilha, "Cambio")
    If lErro <> SUCESSO Then gError 199608
    
    lErro = Cria_Coluna_Excel(objPlanilha, "TarifMoeBr")
    If lErro <> SUCESSO Then gError 199609
    
    lErro = Cria_Coluna_Excel(objPlanilha, "ValNet")
    If lErro <> SUCESSO Then gError 199610
    
    lErro = Cria_Coluna_Excel(objPlanilha, "ValNetBr")
    If lErro <> SUCESSO Then gError 199611
    
    lErro = Cria_Coluna_Excel(objPlanilha, "ValCmPrm")
    If lErro <> SUCESSO Then gError 199612
    
    lErro = Cria_Coluna_Excel(objPlanilha, "Aporte")
    If lErro <> SUCESSO Then gError 199613
    
    lErro = Cria_Coluna_Excel(objPlanilha, "Re. Call Center")
    If lErro <> SUCESSO Then gError 207554
    
    
'    ' *** Transfere os dados retornados pela função de leitura acima para o formato que
'    ' será passado para o Excel ***
'
'    '***** Trecho adicionado por Rafael Menezes em 03/09/2002
'    For iIndiceCol = 1 To objBrowse.colBrowseUsuarioCampo.Count
'
'        '################################################################
'        'Inserido por Wagner 26/01/2006
'        'Após uma modificação na posição dos campos a coleção colregistros
'        'não tem sua ordem alterada, por isso é necessário confrontar a ordem inicial
'        'com a ordem atual para fazer a correlação correta entre os campos
'        iIndice = 0
'        For Each objBrowseUsuarioCampo In objBrowse.colBrowseUsuarioCampo
'            iIndice = iIndice + 1
'            If objBrowseUsuarioCampo.iPosicaoTela = iIndiceCol Then
'                iIndiceColdeFato = iIndice
'                Exit For
'            End If
'        Next
'        '################################################################
'
'        For iIndiceLin = 1 To objBrowse.colRegistros.Count
'
'            'aponta para a coluna em questão
'            Set objColunas = objPlanilha.colColunas(iIndiceCol)
'
'            'Instancia um novo objeto para armazenar os dados da célula
'            Set objCelulas = New ClassCelulasExcel
'
'            'se não for data_Nula
'            'If objBrowse.colRegistros.Item(iIndiceLin)(iIndiceCol) <> DATA_NULA Then
'            If objBrowse.colRegistros.Item(iIndiceLin)(iIndiceColdeFato) <> DATA_NULA Then
'
'                'guarda o valor do registro dakela linha, dakela coluna
'                'objCelulas.vValor = objBrowse.colRegistros.Item(iIndiceLin)(iIndiceCol)
'                objCelulas.vValor = objBrowse.colRegistros.Item(iIndiceLin)(iIndiceColdeFato)
'
'                'se for uma data, formata para "dd/mm/yyyy", caso contrário, o Excel inverte o mês com o dia e o dado perde a semântica
'                'passando a ser tratado como um texto qualquer
'                Select Case TypeName(objCelulas.vValor)
'                    Case "Date"
'                        objCelulas.vValor = StrParaDate(objCelulas.vValor)
'
'                    Case "Double"
'                        objCelulas.vValor = StrParaDbl(objCelulas.vValor)
'
'                        Select Case objBrowse.colValorCampo.Item(iIndiceColdeFato).iSubTipo
'
'                            Case ADM_SUBTIPO_PERCENTUAL
'                                objCelulas.sNumberFormat = "0.00%"
'
'                            Case ADM_SUBTIPO_HORA
'                                objCelulas.vValor = CDate(objCelulas.vValor)
'                                objCelulas.sNumberFormat = "hh:mm:ss"
'
'                            Case Else
'                                objCelulas.sNumberFormat = "#,##0.00##"
'
'                        End Select
'
'                    Case Else
'
'                        '##################################################
'                        'Inserido por Wagner 17/07/2006
'
'                        If Left(objBrowseUsuarioCampo.sNome, 7) = "Periodo" And Left(objBrowseUsuarioCampo.sNome, 9) <> "PeriodoDe" And Left(objBrowseUsuarioCampo.sNome, 10) <> "PeriodoAte" Then
'                            For iIndiceColAux = 1 To objBrowse.colBrowseUsuarioCampo.Count
'                                If UCase(objBrowse.colBrowseUsuarioCampo.Item(iIndiceColAux).sNome) = "EXERCICIO" Or UCase(objBrowse.colBrowseUsuarioCampo.Item(iIndiceColAux).sNome) = "EXERCÍCIO" Then
'                                    iExercicio = objBrowse.colRegistros.Item(iIndiceLin)(iIndiceColAux)
'                                    Exit For
'                                End If
'                            Next
'                        End If
'
'                        lErro = CF("Exclusao_Valida_Mascara", objBrowse.colValorCampo.Item(iIndiceColdeFato).iTipo, objBrowse.colValorCampo.Item(iIndiceColdeFato).iSubTipo, objCelulas.vValor, vValorCampo, True, objBrowse.lComando2, objBrowse.lComando3, iExercicio)
'                        If lErro <> SUCESSO Then gError 181223
'
'                        objCelulas.vValor = vValorCampo
'                        '##################################################
'
'                End Select
'
'            End If
'
'            'adiciona a célula à coleção de células da coluna em questão
'            objColunas.colCelulas.Add objCelulas
'
'        Next
'
'    Next
    '***** fim do Trecho adicionado por Rafael Menezes em 03/09/2002

    Exporta_Cabecalho_Excel = SUCESSO
    
    Exit Function
    
Erro_Exporta_Cabecalho_Excel:
    
    Exporta_Cabecalho_Excel = gErr
    
    Select Case gErr

        Case 102098, 181223, 207554
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144015)

    End Select

    Exit Function
    
End Function

Private Function Cria_Coluna_Excel(ByVal objPlanilha As ClassPlanilhaExcel, ByVal sTitulo As String) As Long

Dim objColunas As ClassColunasExcel
Dim objCelulas As ClassCelulasExcel

On Error GoTo Erro_Cria_Coluna_Excel

    'Instancia um novo obj para armazenar os dados da coluna
    Set objColunas = New ClassColunasExcel
    
    'Guarda no objeto a largura dessa coluna
    objColunas.dLarguraColuna = 10
    
    'Adiciona a coluna à coleção de colunas
    objPlanilha.colColunas.Add objColunas

    'Instancia um novo objeto para armazenar os dados da primeira célula da coluna em questão
    Set objCelulas = New ClassCelulasExcel

    'Guarda o título da coluna em questão que será exibido na primeira célula
    objCelulas.vValor = sTitulo
    objCelulas.bFonteNegrito = True
     
    'Adiciona a célula à coleção de células
    objColunas.colCelulas.Add objCelulas

    Cria_Coluna_Excel = SUCESSO
    
    Exit Function
    
Erro_Cria_Coluna_Excel:
    
    Cria_Coluna_Excel = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199591)

    End Select

    Exit Function

End Function

Private Function Cria_Linha_Excel(ByVal objPlanilha As ClassPlanilhaExcel, ByVal iEmpresa As Integer, ByVal sPromotor As String, ByVal sSupervisor As String, ByVal sGerente As String, ByVal sDiretor As String, ByVal objTRVVoucher As ClassTRVVouchers, ByVal objComissao As Object) As Long

Dim objColunas As ClassColunasExcel
Dim objCelulas As ClassCelulasExcel
Dim iColuna As Integer

On Error GoTo Erro_Cria_Linha_Excel

    For iColuna = 1 To objPlanilha.colColunas.Count

        'aponta para a coluna em questão
        Set objColunas = objPlanilha.colColunas(iColuna)

        'Instancia um novo objeto para armazenar os dados da célula
        Set objCelulas = New ClassCelulasExcel

        'adiciona a célula à coleção de células da coluna em questão
        objColunas.colCelulas.Add objCelulas

        Select Case iColuna
        
            Case 1 'Est
            
                If objTRVVoucher.lCliente > 1000000 Then
                    objCelulas.vValor = "MYA"
                Else
                    If iEmpresa = 1 Then
                        objCelulas.vValor = "TVA"
                    Else
                        objCelulas.vValor = "TVI"
                    End If
                End If
                
            Case 2 'Prom
                objCelulas.vValor = sPromotor
                
            Case 3 'Sup
                objCelulas.vValor = sSupervisor
                
            Case 4 'Ger
                objCelulas.vValor = sGerente
                    
            Case 5 'Dir
                objCelulas.vValor = sDiretor

            Case 6 ' Codigo do Clente
                objCelulas.vValor = objTRVVoucher.lCliente
                    
            Case 7 'Nome do Cliente
                objCelulas.vValor = objTRVVoucher.sNomeRedCli
                
            Case 8 'Nome do Produto
                objCelulas.vValor = objTRVVoucher.sNomeRedProd
                
            Case 9 'Percentual de Comissao do Promotor
                If objTRVVoucher.iAporte = 2 Or objTRVVoucher.lCliente > 1000000 Then 'se for call center ou MyAssistance
                    objCelulas.vValor = 0.01
                Else
                    objCelulas.vValor = objComissao.dPercentual
                End If
                objCelulas.sNumberFormat = "0.00%"

            Case 10 'Numero do Voucher
                objCelulas.vValor = objTRVVoucher.lNumVou
                
            Case 11 'Data de emissao
                objCelulas.vValor = objTRVVoucher.dtData
                
            Case 12 'Data de faturamento
                objCelulas.vValor = objTRVVoucher.dtDataFat

            Case 13 'Numero da Fatura
                objCelulas.vValor = objTRVVoucher.lNumFatCoinfo

            Case 14 'Moeda
                objCelulas.vValor = objTRVVoucher.sMoeda

            Case 15 'Valor bruto na moeda usada
                objCelulas.vValor = objTRVVoucher.dValorCambio
                
            Case 16 'Cambio
                objCelulas.vValor = objTRVVoucher.dCambio
                
            Case 17 'Valor Bruto em Reais
                objCelulas.vValor = objTRVVoucher.dValorBruto

            Case 18 'valor net base na moeda usada
                objCelulas.vValor = Round(objComissao.dValorBase / objTRVVoucher.dCambio, 2)

            Case 19 'valor net base em reais
                objCelulas.vValor = objComissao.dValorBase

            Case 20 'valor da comissao
                If objTRVVoucher.iAporte = 2 Or objTRVVoucher.lCliente > 1000000 Then 'se for call center ou MyAssistance
                    objCelulas.vValor = objComissao.dValorBase * 0.01
                Else
                    objCelulas.vValor = objComissao.dValor
                End If
                
            Case 21 'aporte
'
'                If objTRVVoucher.lCliente > 1000000 Then
'                    objCelulas.vValor = "Nao"
'                Else
                    If objTRVVoucher.iAporte = 1 Then
                        objCelulas.vValor = "Sim"
                    ElseIf objTRVVoucher.iAporte = 2 Then
                        objCelulas.vValor = "Call"
                    Else
                        objCelulas.vValor = "Nao"
                    End If
'                End If
                    
            Case 22 'UsuRespCallCenter
                If objTRVVoucher.iAporte = 2 Then
                    objCelulas.vValor = objTRVVoucher.sUsuRespCallCenter
                End If
                
                
        End Select

    Next
    
    Cria_Linha_Excel = SUCESSO
    
    Exit Function
    
Erro_Cria_Linha_Excel:
    
    Cria_Linha_Excel = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199592)
    
    End Select
    
    Call Rotina_Erro(vbOKOnly, "Ocorreu um erro na função Cria_Linha_Excel, coluna " & CStr(iColuna) & " para o voucher " & objTRVVoucher.sTipVou & objTRVVoucher.sSerie & CStr(objTRVVoucher.lNumVou) & " de valor " & Format(objTRVVoucher.dValor, "STANDARD") & " e câmbio " & Format(objTRVVoucher.dCambio, "STANDARD"), 0)
    
    Exit Function

End Function

