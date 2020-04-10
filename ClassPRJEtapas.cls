VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJEtapas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocPRJ As Long
Private mvarsCodigo As String
Private mvarsCodigoAnt As String
Private mvarsCodigoPaiOrg As String
Private mvarsReferencia As String
Private mvarsNomeReduzido As String
Private mvarsDescricao As String
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvarsResponsavel As String
Private mvarsObjetivo As String
Private mvarsJustificativa As String
Private mvarsObservacao As String
Private mvardtDataInicio As Date
Private mvardtDataFim As Date
Private mvarlNumIntDocEtapaPaiOrg As Long
Private mvardtDataInicioReal As Date
Private mvardtDataFimReal As Date
Private mvardPercentualComplet As Double
Private mvarlNumIntDocEscopo As Long
Private mvarobjEscopo As New ClassPRJEscopo
Private mvarobjDadosCalc As New ClassPRJDadosCalc
Private mvarobjTiposCamposCust As New ClassTiposCamposCust
Private mvarobjCamposCust As New ClassCamposCustomizados
Private mvariIndiceTvw As Integer
Private mvarsChaveTvw As String
Private mvarcolPredecessores As New Collection
Private mvarcolMaquinas As New Collection
Private mvarcolMaoDeObra As New Collection
Private mvarcolMateriaPrima As New Collection
Private mvarcolItensProduzidos As New Collection
Private mvariSeq As Integer
Private mvariNivel As Integer
Private mvariNovo As Integer
Private mvariPosicao As Integer
Private mvardPreco As Double
Private mvardCustoInfoPrev As Double
Private mvardCustoCalcPrev As Double
Private mvardCustoInfoReal As Double
Private mvardCustoCalcReal As Double
Private mvariTvwChecked As Integer
Private mvarobjTela As Object
Private mvardtDataVistoria As Date
Private mvardtValidadeVistoria As Date

Public Property Let dtDataVistoria(ByVal vData As Date)
    mvardtDataVistoria = vData
End Property

Public Property Get dtDataVistoria() As Date
    dtDataVistoria = mvardtDataVistoria
End Property

Public Property Let dtValidadeVistoria(ByVal vData As Date)
    mvardtValidadeVistoria = vData
End Property

Public Property Get dtValidadeVistoria() As Date
    dtValidadeVistoria = mvardtValidadeVistoria
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let iTvwChecked(ByVal vData As Integer)
    mvariTvwChecked = vData
End Property

Public Property Get iTvwChecked() As Integer
    iTvwChecked = mvariTvwChecked
End Property

Public Property Let dCustoCalcReal(ByVal vData As Double)
    mvardCustoCalcReal = vData
End Property

Public Property Get dCustoCalcReal() As Double
    dCustoCalcReal = mvardCustoCalcReal
End Property

Public Property Let dCustoInfoReal(ByVal vData As Double)
    mvardCustoInfoReal = vData
End Property

Public Property Get dCustoInfoReal() As Double
    dCustoInfoReal = mvardCustoInfoReal
End Property

Public Property Let dCustoCalcPrev(ByVal vData As Double)
    mvardCustoCalcPrev = vData
End Property

Public Property Get dCustoCalcPrev() As Double
    dCustoCalcPrev = mvardCustoCalcPrev
End Property

Public Property Let dCustoInfoPrev(ByVal vData As Double)
    mvardCustoInfoPrev = vData
End Property

Public Property Get dCustoInfoPrev() As Double
    dCustoInfoPrev = mvardCustoInfoPrev
End Property

Public Property Let dPreco(ByVal vData As Double)
    mvardPreco = vData
End Property

Public Property Get dPreco() As Double
    dPreco = mvardPreco
End Property

Public Property Let iPosicao(ByVal vData As Integer)
    mvariPosicao = vData
End Property

Public Property Get iPosicao() As Integer
    iPosicao = mvariPosicao
End Property

Public Property Let iNovo(ByVal vData As Integer)
    mvariNovo = vData
End Property

Public Property Get iNovo() As Integer
    iNovo = mvariNovo
End Property

Public Property Let iNivel(ByVal vData As Integer)
    mvariNivel = vData
End Property

Public Property Get iNivel() As Integer
    iNivel = mvariNivel
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Set colItensProduzidos(ByVal vData As Collection)
    Set mvarcolItensProduzidos = vData
End Property

Public Property Get colItensProduzidos() As Collection
    Set colItensProduzidos = mvarcolItensProduzidos
End Property

Public Property Set colMateriaPrima(ByVal vData As Collection)
    Set mvarcolMateriaPrima = vData
End Property

Public Property Get colMateriaPrima() As Collection
    Set colMateriaPrima = mvarcolMateriaPrima
End Property

Public Property Set colMaoDeObra(ByVal vData As Collection)
    Set mvarcolMaoDeObra = vData
End Property

Public Property Get colMaoDeObra() As Collection
    Set colMaoDeObra = mvarcolMaoDeObra
End Property

Public Property Set colMaquinas(ByVal vData As Collection)
    Set mvarcolMaquinas = vData
End Property

Public Property Get colMaquinas() As Collection
    Set colMaquinas = mvarcolMaquinas
End Property

Public Property Set colPredecessores(ByVal vData As Collection)
    Set mvarcolPredecessores = vData
End Property

Public Property Get colPredecessores() As Collection
    Set colPredecessores = mvarcolPredecessores
End Property

Public Property Let sChaveTvw(ByVal vData As String)
    mvarsChaveTvw = vData
End Property

Public Property Get sChaveTvw() As String
    sChaveTvw = mvarsChaveTvw
End Property

Public Property Let iIndiceTvw(ByVal vData As Integer)
    mvariIndiceTvw = vData
End Property

Public Property Get iIndiceTvw() As Integer
    iIndiceTvw = mvariIndiceTvw
End Property

Public Property Let objCamposCust(ByVal vData As ClassCamposCustomizados)
    Set mvarobjCamposCust = vData
End Property

Public Property Get objCamposCust() As ClassCamposCustomizados
    Set objCamposCust = mvarobjCamposCust
End Property

Public Property Let objTiposCamposCust(ByVal vData As ClassTiposCamposCust)
    Set mvarobjTiposCamposCust = vData
End Property

Public Property Get objTiposCamposCust() As ClassTiposCamposCust
    Set objTiposCamposCust = mvarobjTiposCamposCust
End Property

Public Property Let objDadosCalc(ByVal vData As ClassPRJDadosCalc)
    Set mvarobjDadosCalc = vData
End Property

Public Property Get objDadosCalc() As ClassPRJDadosCalc
    Set objDadosCalc = mvarobjDadosCalc
End Property

Public Property Let objEscopo(ByVal vData As ClassPRJEscopo)
    Set mvarobjEscopo = vData
End Property

Public Property Get objEscopo() As ClassPRJEscopo
    Set objEscopo = mvarobjEscopo
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocPRJ(ByVal vData As Long)
    mvarlNumIntDocPRJ = vData
End Property

Public Property Get lNumIntDocPRJ() As Long
    lNumIntDocPRJ = mvarlNumIntDocPRJ
End Property

Public Property Let sCodigoAnt(ByVal vData As String)
    mvarsCodigoAnt = vData
End Property

Public Property Get sCodigoAnt() As String
    sCodigoAnt = mvarsCodigoAnt
End Property

Public Property Let sCodigoPaiOrg(ByVal vData As String)
    mvarsCodigoPaiOrg = vData
End Property

Public Property Get sCodigoPaiOrg() As String
    sCodigoPaiOrg = mvarsCodigoPaiOrg
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let sReferencia(ByVal vData As String)
    mvarsReferencia = vData
End Property

Public Property Get sReferencia() As String
    sReferencia = mvarsReferencia
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let sResponsavel(ByVal vData As String)
    mvarsResponsavel = vData
End Property

Public Property Get sResponsavel() As String
    sResponsavel = mvarsResponsavel
End Property

Public Property Let sObjetivo(ByVal vData As String)
    mvarsObjetivo = vData
End Property

Public Property Get sObjetivo() As String
    sObjetivo = mvarsObjetivo
End Property

Public Property Let sJustificativa(ByVal vData As String)
    mvarsJustificativa = vData
End Property

Public Property Get sJustificativa() As String
    sJustificativa = mvarsJustificativa
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Let lNumIntDocEtapaPaiOrg(ByVal vData As Long)
    mvarlNumIntDocEtapaPaiOrg = vData
End Property

Public Property Get lNumIntDocEtapaPaiOrg() As Long
    lNumIntDocEtapaPaiOrg = mvarlNumIntDocEtapaPaiOrg
End Property

Public Property Let dtDataInicioReal(ByVal vData As Date)
    mvardtDataInicioReal = vData
End Property

Public Property Get dtDataInicioReal() As Date
    dtDataInicioReal = mvardtDataInicioReal
End Property

Public Property Let dtDataFimReal(ByVal vData As Date)
    mvardtDataFimReal = vData
End Property

Public Property Get dtDataFimReal() As Date
    dtDataFimReal = mvardtDataFimReal
End Property

Public Property Let dPercentualComplet(ByVal vData As Double)
    mvardPercentualComplet = vData
End Property

Public Property Get dPercentualComplet() As Double
    dPercentualComplet = mvardPercentualComplet
End Property

Public Property Let lNumIntDocEscopo(ByVal vData As Long)
    mvarlNumIntDocEscopo = vData
End Property

Public Property Get lNumIntDocEscopo() As Long
    lNumIntDocEscopo = mvarlNumIntDocEscopo
End Property

Function Obtem_Custo() As Long

Dim lErro As Long
Dim objMP As ClassPRJEtapaMateriais
Dim objMO As ClassPRJEtapaMO
Dim objMaq As ClassPRJEtapaMaquinas
Dim dCusto As Double
Dim iTipo As Integer

On Error GoTo Erro_Obtem_Custo

    For iTipo = 1 To 4
        dCusto = 0
        For Each objMP In colMateriaPrima
            If objMP.iTipo = iTipo Then
                dCusto = dCusto + objMP.dCusto
            End If
        Next
        For Each objMaq In colMaquinas
            If objMaq.iTipo = iTipo Then
                dCusto = dCusto + objMaq.dCusto
            End If
        Next
        For Each objMO In colMaoDeObra
            If objMO.iTipo = iTipo Then
                dCusto = dCusto + objMO.dCusto
            End If
        Next
        
        Select Case iTipo
        
            Case INDICE_CALC_PREV
                dCustoCalcPrev = dCusto
            Case INDICE_CALC_REAL
                dCustoCalcReal = dCusto
            Case INDICE_INF_PREV
                dCustoInfoPrev = dCusto
            Case INDICE_INF_REAL
                dCustoInfoReal = dCusto
        
        End Select
        
    Next

    Obtem_Custo = SUCESSO

    Exit Function

Erro_Obtem_Custo:

    Obtem_Custo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187475)

    End Select

    Exit Function

End Function

Function Cria_Copia(ByVal objEtapa As ClassPRJEtapas) As Long

Dim lErro As Long

On Error GoTo Erro_Cria_Copia

    Set objEtapa.colItensProduzidos = colItensProduzidos
    Set objEtapa.colMaoDeObra = colMaoDeObra
    Set objEtapa.colMaquinas = colMaquinas
    Set objEtapa.colMateriaPrima = colMateriaPrima
    Set objEtapa.colPredecessores = colPredecessores
    
    objEtapa.dCustoCalcPrev = dCustoCalcPrev
    objEtapa.dCustoCalcReal = dCustoCalcReal
    objEtapa.dCustoInfoPrev = dCustoInfoPrev
    objEtapa.dCustoInfoReal = dCustoInfoReal
    objEtapa.dPercentualComplet = dPercentualComplet
    objEtapa.dPreco = dPreco
    objEtapa.dtDataFim = dtDataFim
    objEtapa.dtDataFimReal = dtDataFimReal
    objEtapa.dtDataInicio = dtDataInicio
    objEtapa.dtDataInicioReal = dtDataInicioReal
    objEtapa.iFilialCliente = iFilialCliente
    objEtapa.iIndiceTvw = iIndiceTvw
    objEtapa.iNivel = iNivel
    objEtapa.iNovo = iNovo
    objEtapa.iPosicao = iPosicao
    objEtapa.iSeq = iSeq
    objEtapa.iTvwChecked = iTvwChecked
    objEtapa.lCliente = lCliente
    objEtapa.lNumIntDoc = lNumIntDoc
    objEtapa.lNumIntDocEscopo = lNumIntDocEscopo
    objEtapa.lNumIntDocEtapaPaiOrg = lNumIntDocEtapaPaiOrg
    objEtapa.lNumIntDocPRJ = lNumIntDocPRJ
    objEtapa.sChaveTvw = sChaveTvw
    objEtapa.sCodigo = sCodigo
    objEtapa.sCodigoAnt = sCodigoAnt
    objEtapa.sCodigoPaiOrg = sCodigoPaiOrg
    objEtapa.sDescricao = sDescricao
    objEtapa.sJustificativa = sJustificativa
    objEtapa.sNomeReduzido = sNomeReduzido
    objEtapa.sObjetivo = sObjetivo
    objEtapa.sObservacao = sObservacao
    objEtapa.sReferencia = sReferencia
    objEtapa.sResponsavel = sResponsavel

    Cria_Copia = SUCESSO

    Exit Function

Erro_Cria_Copia:

    Cria_Copia = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187475)

    End Select

    Exit Function

End Function
