VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

Public objInfoUsu As Object 'Inserido por Wagner 20/03/2006

Private Const INICIO_PROJETO = "Inicio_Projeto"
Private Const INICIO_ETAPA = "Inicio_Etapa"
Private Const PERC_COMPL_PRJ = "Perc_Compl_PRJ"
Private Const PERC_COMPL_ETAPA = "Perc_Compl_Etapa"

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsCodigo As String
Private mvariFilialEmpresa As Integer
Private mvarsNomeReduzido As String
Private mvarsDescricao As String
Private mvardtDataCriacao As Date
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvarsResponsavel As String
Private mvarsObjetivo As String
Private mvarsJustificativa As String
Private mvarsObservacao As String
Private mvardtDataInicio As Date
Private mvardtDataFim As Date
Private mvardtDataInicioReal As Date
Private mvardtDataFimReal As Date
Private mvardPercentualComplet As Double
Private mvarlNumIntDocEscopo As Long
Private mvarobjEscopo As New ClassPRJEscopo
Private mvarobjDadosCalc As New ClassPRJDadosCalc
Private mvarobjTiposCamposCust As New ClassTiposCamposCust
Private mvarobjCamposCust As New ClassCamposCustomizados
Private mvarcolEtapas As New Collection
Private mvarcolRecebPagto As New Collection
Private mvarobjTela As Object
Private mvarobjAnexos As Object
Private mvarsSegmento As String

Public Property Let sSegmento(ByVal vData As String)
    mvarsSegmento = vData
End Property

Public Property Get sSegmento() As String
    sSegmento = mvarsSegmento
End Property

Public Property Set objAnexos(ByVal vData As Object)
    Set mvarobjAnexos = vData
End Property

Public Property Get objAnexos() As Object
    Set objAnexos = mvarobjAnexos
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Set colRecebPagto(ByVal vData As Collection)
    Set mvarcolRecebPagto = vData
End Property

Public Property Get colRecebPagto() As Collection
    Set colRecebPagto = mvarcolRecebPagto
End Property

Public Property Set colEtapas(ByVal vData As Collection)
    Set mvarcolEtapas = vData
End Property

Public Property Get colEtapas() As Collection
    Set colEtapas = mvarcolEtapas
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

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
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

Public Property Let dtDataCriacao(ByVal vData As Date)
    mvardtDataCriacao = vData
End Property

Public Property Get dtDataCriacao() As Date
    dtDataCriacao = mvardtDataCriacao
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

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim bAchou As Boolean
Dim objEtapa As New ClassPRJEtapas
Dim dPerc As Double
Dim iNumDias As Integer

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
                                                      
        Case INICIO_PROJETO
            objMnemonicoValor.colValor.Add dtDataInicio
        
        Case INICIO_ETAPA
        
            bAchou = False
            For Each objEtapa In colEtapas
                If objEtapa.sCodigo = objMnemonicoValor.vParam(1) Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then gError 187991
        
            objMnemonicoValor.colValor.Add objEtapa.dtDataInicio

        Case PERC_COMPL_PRJ
            
            dPerc = objMnemonicoValor.vParam(1)
            
            iNumDias = DateDiff("d", dtDataInicio, dtDataFim)
            
            iNumDias = iNumDias * dPerc / 100
        
            objMnemonicoValor.colValor.Add DateAdd("d", dtDataInicio, iNumDias)
  
        Case PERC_COMPL_ETAPA
            
            bAchou = False
            For Each objEtapa In colEtapas
                If objEtapa.sCodigo = objMnemonicoValor.vParam(1) Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then gError 187992
            
            dPerc = objMnemonicoValor.vParam(2)
            
            iNumDias = DateDiff("d", objEtapa.dtDataInicio, objEtapa.dtDataFim)
            
            iNumDias = iNumDias * dPerc / 100
        
            objMnemonicoValor.colValor.Add DateAdd("d", objEtapa.dtDataInicio, iNumDias)
  
        Case Else
            gError 187163

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 187163
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            
        Case 187991, 187992
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAS_NAO_CADASTRADO2", gErr, objMnemonicoValor.vParam(1), sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187164)

    End Select

    Exit Function

End Function

Function Calcula_Dados_Calculados(Optional ByVal colEtp As Collection = Nothing) As Long

Dim lErro As Long
Dim colEtpAux As Collection
Dim objEtapa As ClassPRJEtapas
Dim objEtapaAux As ClassPRJEtapas
Dim iNivel As Integer
Dim iUltimoNivel As Integer
Dim dtDataIniP As Date
Dim dtDataFimP As Date
Dim dtDataIniR As Date
Dim dtDataFimR As Date
Dim dFator As Double
Dim dValor As Double
Dim bEtapaIncompleta As Boolean
Dim bFolha As Boolean

On Error GoTo Erro_Calcula_Dados_Calculados

    If colEtp Is Nothing Then
        Set colEtpAux = colEtapas
    Else
        Set colEtpAux = colEtp
    End If

    For Each objEtapa In colEtpAux
        If iUltimoNivel < objEtapa.iNivel Then
            iUltimoNivel = objEtapa.iNivel
        End If
    Next
    
    For iNivel = iUltimoNivel To 0 Step -1
    
        For Each objEtapa In colEtpAux
        
            If iNivel = objEtapa.iNivel Or iNivel = 0 Then
            
                If iNivel = 0 Then
                    Set objEtapa = New ClassPRJEtapas
                    objEtapa.objDadosCalc = objDadosCalc
                End If
                
                bFolha = True
                For Each objEtapaAux In colEtpAux
                    If objEtapa.lNumIntDoc = objEtapaAux.lNumIntDocEtapaPaiOrg Then
                        bFolha = False
                        Exit For
                    End If
                Next
            
                If bFolha Then
                
                    objEtapa.objDadosCalc.dPercentualComplet = objEtapa.dPercentualComplet
                    objEtapa.objDadosCalc.dtDataFimPrev = objEtapa.dtDataFim
                    objEtapa.objDadosCalc.dtDataFimReal = objEtapa.dtDataFimReal
                    objEtapa.objDadosCalc.dtDataIniPrev = objEtapa.dtDataInicio
                    objEtapa.objDadosCalc.dtDataIniReal = objEtapa.dtDataInicioReal
                Else
                
                    dtDataIniP = DATA_NULA
                    dtDataFimP = DATA_NULA
                    dtDataIniR = DATA_NULA
                    dtDataFimR = DATA_NULA
                    bEtapaIncompleta = False
                    dValor = 0
                    dFator = 0
                    For Each objEtapaAux In colEtpAux
                        If objEtapaAux.iNivel - 1 = objEtapa.iNivel And objEtapa.lNumIntDoc = objEtapaAux.lNumIntDocEtapaPaiOrg Then
                            If dtDataIniP = DATA_NULA Or (dtDataIniP > objEtapaAux.objDadosCalc.dtDataIniPrev And objEtapaAux.objDadosCalc.dtDataIniPrev <> DATA_NULA) Then
                                dtDataIniP = objEtapaAux.objDadosCalc.dtDataIniPrev
                            End If
                            If dtDataIniR = DATA_NULA Or (dtDataIniR > objEtapaAux.objDadosCalc.dtDataIniReal And objEtapaAux.objDadosCalc.dtDataIniReal <> DATA_NULA) Then
                                dtDataIniR = objEtapaAux.objDadosCalc.dtDataIniReal
                            End If
                            If dtDataFimP = DATA_NULA Or dtDataFimP < objEtapaAux.objDadosCalc.dtDataFimPrev Then
                                dtDataFimP = objEtapaAux.objDadosCalc.dtDataFimPrev
                            End If
                            If dtDataFimR = DATA_NULA Or dtDataFimR < objEtapaAux.objDadosCalc.dtDataFimReal Then
                                dtDataFimR = objEtapaAux.objDadosCalc.dtDataFimReal
                            End If
                            
                            If objEtapaAux.objDadosCalc.dtDataFimReal = DATA_NULA Then
                                bEtapaIncompleta = True
                            End If
                            
                            dValor = dValor + objEtapaAux.objDadosCalc.dPercentualComplet * (1 + DateDiff("d", objEtapaAux.dtDataInicio, objEtapaAux.dtDataFim))
                            dFator = dFator + (1 + DateDiff("d", objEtapaAux.dtDataInicio, objEtapaAux.dtDataFim))
                        End If
                    Next
                    
                    'Se existe uma etapa filha não finalizada
                    'é porque esse etapa ainda não foi concluída
                    If bEtapaIncompleta Then
                        dtDataFimR = DATA_NULA
                    End If

                    If dFator <> 0 Then
                        objEtapa.objDadosCalc.dPercentualComplet = dValor / dFator
                    Else
                        objEtapa.objDadosCalc.dPercentualComplet = 0
                    End If
                    
                    objEtapa.objDadosCalc.dtDataFimPrev = dtDataFimP
                    objEtapa.objDadosCalc.dtDataFimReal = dtDataFimR
                    objEtapa.objDadosCalc.dtDataIniPrev = dtDataIniP
                    objEtapa.objDadosCalc.dtDataIniReal = dtDataIniR
                    
                    If iNivel = 0 Then Exit For
               
                End If
            End If
        Next
    Next
    
    Calcula_Dados_Calculados = SUCESSO

    Exit Function

Erro_Calcula_Dados_Calculados:

    Calcula_Dados_Calculados = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189253)

    End Select

    Exit Function

End Function
