VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoDeTributacaoMovto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariTipo As Integer
Private mvarsDescricao As String
Private mvariEntrada As Integer
Private mvariICMSIncide As Integer
Private mvariICMSTipo As Integer
Private mvariICMSBaseComIPI As Integer
Private mvariICMSCredita As Integer
Private mvariIPICredita As Integer
Private mvariIPIIncide As Integer
Private mvariIPITipo As Integer
Private mvariIPIFrete As Integer
Private mvariIPIDestaca As Integer
Private mvariISSIncide As Integer
Private mvariIRIncide As Integer
Private mvardIRAliquota As Double
Private mvariINSSIncide As Integer
Private mvardINSSRetencaoMinima As Double
Private mvardINSSAliquota As Double
Private mvariPISCredita As Integer
Private mvariPISRetencao As Integer
Private mvariCOFINSCredita As Integer
Private mvariCOFINSRetencao As Integer
Private mvariCSLLRetencao As Integer
Private mvariISSRetencao As Integer
Private mvariISSTipo As Integer
Private mvariPISTipo As Integer
Private mvariCOFINSTipo As Integer
Private mvariICMSSimplesTipo As Integer
Private mvariRegimeTributario As Integer
Private mvarsNatBCCred As String

'nfe 3.10
Public iISSIndExigibilidade As Integer
Private mvarsIPICodEnq As String

Public Property Let sIPICodEnq(ByVal vData As String)
    mvarsIPICodEnq = vData
End Property

Public Property Get sIPICodEnq() As String
    sIPICodEnq = mvarsIPICodEnq
End Property

'**** ATENCAO: QUANDO INCLUIR NOVO CAMPO ATUALIZAR A SUB COPIAR ******'

Public Property Let sNatBCCred(ByVal vData As String)
    mvarsNatBCCred = vData
End Property

Public Property Get sNatBCCred() As String
    sNatBCCred = mvarsNatBCCred
End Property

Public Property Let iRegimeTributario(ByVal vData As Integer)
    mvariRegimeTributario = vData
End Property

Public Property Get iRegimeTributario() As Integer
    iRegimeTributario = mvariRegimeTributario
End Property

Public Property Let iICMSSimplesTipo(ByVal vData As Integer)
    mvariICMSSimplesTipo = vData
End Property

Public Property Get iICMSSimplesTipo() As Integer
    iICMSSimplesTipo = mvariICMSSimplesTipo
End Property

Public Property Let iISSTipo(ByVal vData As Integer)
    mvariISSTipo = vData
End Property

Public Property Get iISSTipo() As Integer
    iISSTipo = mvariISSTipo
End Property

Public Property Let iPISTipo(ByVal vData As Integer)
    mvariPISTipo = vData
End Property

Public Property Get iPISTipo() As Integer
    iPISTipo = mvariPISTipo
End Property

Public Property Let iCOFINSTipo(ByVal vData As Integer)
    mvariCOFINSTipo = vData
End Property

Public Property Get iCOFINSTipo() As Integer
    iCOFINSTipo = mvariCOFINSTipo
End Property

Public Property Let iISSRetencao(ByVal vData As Integer)
    mvariISSRetencao = vData
End Property

Public Property Get iISSRetencao() As Integer
    iISSRetencao = mvariISSRetencao
End Property

Public Property Let iCSLLRetencao(ByVal vData As Integer)
    mvariCSLLRetencao = vData
End Property

Public Property Get iCSLLRetencao() As Integer
    iCSLLRetencao = mvariCSLLRetencao
End Property

Public Property Let iCOFINSRetencao(ByVal vData As Integer)
    mvariCOFINSRetencao = vData
End Property

Public Property Get iCOFINSRetencao() As Integer
    iCOFINSRetencao = mvariCOFINSRetencao
End Property

Public Property Let iCOFINSCredita(ByVal vData As Integer)
    mvariCOFINSCredita = vData
End Property

Public Property Get iCOFINSCredita() As Integer
    iCOFINSCredita = mvariCOFINSCredita
End Property

Public Property Let iPISRetencao(ByVal vData As Integer)
    mvariPISRetencao = vData
End Property

Public Property Get iPISRetencao() As Integer
    iPISRetencao = mvariPISRetencao
End Property

Public Property Let iPISCredita(ByVal vData As Integer)
    mvariPISCredita = vData
End Property

Public Property Get iPISCredita() As Integer
    iPISCredita = mvariPISCredita
End Property

Public Property Let dINSSRetencaoMinima(ByVal vData As Double)
    mvardINSSRetencaoMinima = vData
End Property

Public Property Get dINSSRetencaoMinima() As Double
    dINSSRetencaoMinima = mvardINSSRetencaoMinima
End Property

Public Property Let dIRAliquota(ByVal vData As Double)
    mvardIRAliquota = vData
End Property

Public Property Get dIRAliquota() As Double
    dIRAliquota = mvardIRAliquota
End Property

Public Property Let dINSSAliquota(ByVal vData As Double)
    mvardINSSAliquota = vData
End Property

Public Property Get dINSSAliquota() As Double
    dINSSAliquota = mvardINSSAliquota
End Property

Public Property Let iINSSIncide(ByVal vData As Integer)
    mvariINSSIncide = vData
End Property

Public Property Get iINSSIncide() As Integer
    iINSSIncide = mvariINSSIncide
End Property

Public Property Let iEntrada(ByVal vData As Integer)
    mvariEntrada = vData
End Property

Public Property Get iEntrada() As Integer
    iEntrada = mvariEntrada
End Property

Public Property Let iIPICredita(ByVal vData As Integer)
    mvariIPICredita = vData
End Property

Public Property Get iIPICredita() As Integer
    iIPICredita = mvariIPICredita
End Property

Public Property Let iIRIncide(ByVal vData As Integer)
    mvariIRIncide = vData
End Property

Public Property Get iIRIncide() As Integer
    iIRIncide = mvariIRIncide
End Property

Public Property Let iISSIncide(ByVal vData As Integer)
    mvariISSIncide = vData
End Property

Public Property Get iISSIncide() As Integer
    iISSIncide = mvariISSIncide
End Property

Public Property Let iIPIDestaca(ByVal vData As Integer)
    mvariIPIDestaca = vData
End Property

Public Property Get iIPIDestaca() As Integer
    iIPIDestaca = mvariIPIDestaca
End Property

Public Property Let iIPIFrete(ByVal vData As Integer)
    mvariIPIFrete = vData
End Property

Public Property Get iIPIFrete() As Integer
    iIPIFrete = mvariIPIFrete
End Property

'Public Property Let iIPICredita(ByVal vData As Integer)
'    mvariIPICredita = vData
'End Property
'
'Public Property Get iIPICredita() As Integer
'    iIPICredita = mvariIPICredita
'End Property
'
Public Property Let iIPIIncide(ByVal vData As Integer)
    mvariIPIIncide = vData
End Property

Public Property Get iIPIIncide() As Integer
    iIPIIncide = mvariIPIIncide
End Property

Public Property Let iIPITipo(ByVal vData As Integer)
    mvariIPITipo = vData
End Property

Public Property Get iIPITipo() As Integer
    iIPITipo = mvariIPITipo
End Property

Public Property Let iICMSTipo(ByVal vData As Integer)
    mvariICMSTipo = vData
End Property

Public Property Get iICMSTipo() As Integer
    iICMSTipo = mvariICMSTipo
End Property

Public Property Let iICMSBaseComIPI(ByVal vData As Integer)
    mvariICMSBaseComIPI = vData
End Property

Public Property Get iICMSBaseComIPI() As Integer
    iICMSBaseComIPI = mvariICMSBaseComIPI
End Property

Public Property Let iICMSCredita(ByVal vData As Integer)
    mvariICMSCredita = vData
End Property

Public Property Get iICMSCredita() As Integer
    iICMSCredita = mvariICMSCredita
End Property

Public Property Let iICMSIncide(ByVal vData As Integer)
    mvariICMSIncide = vData
End Property

Public Property Get iICMSIncide() As Integer
    iICMSIncide = mvariICMSIncide
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Sub Copiar(objTributacaoTipo As ClassTipoDeTributacaoMovto)
'copia os dados de objTributacaoTipo p/"Me"

    mvariTipo = objTributacaoTipo.iTipo
    mvarsDescricao = objTributacaoTipo.sDescricao
    mvariEntrada = objTributacaoTipo.iEntrada
    mvariICMSIncide = objTributacaoTipo.iICMSIncide
    mvariICMSTipo = objTributacaoTipo.iICMSTipo
    mvariICMSCredita = objTributacaoTipo.iICMSCredita
    mvariICMSBaseComIPI = objTributacaoTipo.iICMSBaseComIPI
    mvariIPIIncide = objTributacaoTipo.iIPIIncide
    mvariIPITipo = objTributacaoTipo.iIPITipo
    mvariIPICredita = objTributacaoTipo.iIPICredita
    mvariIPIFrete = objTributacaoTipo.iIPIFrete
    mvariIPIDestaca = objTributacaoTipo.iIPIDestaca
    mvariISSIncide = objTributacaoTipo.iISSIncide
    mvariIRIncide = objTributacaoTipo.iIRIncide
    mvardIRAliquota = objTributacaoTipo.dIRAliquota
    mvariINSSIncide = objTributacaoTipo.iINSSIncide
    mvardINSSRetencaoMinima = objTributacaoTipo.dINSSRetencaoMinima
    mvardINSSAliquota = objTributacaoTipo.dINSSAliquota
    mvariPISCredita = objTributacaoTipo.iPISCredita
    mvariPISRetencao = objTributacaoTipo.iPISRetencao
    mvariCOFINSCredita = objTributacaoTipo.iCOFINSCredita
    mvariCOFINSRetencao = objTributacaoTipo.iCOFINSRetencao
    mvariCSLLRetencao = objTributacaoTipo.iCSLLRetencao
    
    mvariISSTipo = objTributacaoTipo.iISSTipo
    mvariPISTipo = objTributacaoTipo.iPISTipo
    mvariCOFINSTipo = objTributacaoTipo.iCOFINSTipo
    mvariICMSSimplesTipo = objTributacaoTipo.iICMSSimplesTipo
    mvariRegimeTributario = objTributacaoTipo.iRegimeTributario
    mvarsNatBCCred = objTributacaoTipo.sNatBCCred

    iISSIndExigibilidade = objTributacaoTipo.iISSIndExigibilidade
    mvarsIPICodEnq = objTributacaoTipo.sIPICodEnq

End Sub

 
