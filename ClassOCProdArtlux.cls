VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOCProdArtlux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocOC As Long
Private mvariSeq As Integer
Private mvarlNumIntDocMovEst As Long
Private mvarsUsuMontagem As String
Private mvardtDataIniMontagem As Date
Private mvardtDataFimMontagem As Date
Private mvardQuantidadePreProd As Double
Private mvardQuantidadeProd As Double

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocOC(ByVal vData As Long)
    mvarlNumIntDocOC = vData
End Property

Public Property Get lNumIntDocOC() As Long
    lNumIntDocOC = mvarlNumIntDocOC
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let lNumIntDocMovEst(ByVal vData As Long)
    mvarlNumIntDocMovEst = vData
End Property

Public Property Get lNumIntDocMovEst() As Long
    lNumIntDocMovEst = mvarlNumIntDocMovEst
End Property

Public Property Let sUsuMontagem(ByVal vData As String)
    mvarsUsuMontagem = vData
End Property

Public Property Get sUsuMontagem() As String
    sUsuMontagem = mvarsUsuMontagem
End Property

Public Property Let dtDataIniMontagem(ByVal vData As Date)
    mvardtDataIniMontagem = vData
End Property

Public Property Get dtDataIniMontagem() As Date
    dtDataIniMontagem = mvardtDataIniMontagem
End Property

Public Property Let dtDataFimMontagem(ByVal vData As Date)
    mvardtDataFimMontagem = vData
End Property

Public Property Get dtDataFimMontagem() As Date
    dtDataFimMontagem = mvardtDataFimMontagem
End Property

Public Property Let dQuantidadePreProd(ByVal vData As Double)
    mvardQuantidadePreProd = vData
End Property

Public Property Get dQuantidadePreProd() As Double
    dQuantidadePreProd = mvardQuantidadePreProd
End Property

Public Property Let dQuantidadeProd(ByVal vData As Double)
    mvardQuantidadeProd = vData
End Property

Public Property Get dQuantidadeProd() As Double
    dQuantidadeProd = mvardQuantidadeProd
End Property

Public Sub Copiar(ByVal objOCProd As ClassOCProdArtlux)
    mvarlNumIntDoc = objOCProd.lNumIntDoc
    mvarlNumIntDocOC = objOCProd.lNumIntDocOC
    mvariSeq = objOCProd.iSeq
    mvarlNumIntDocMovEst = objOCProd.lNumIntDocMovEst
    mvarsUsuMontagem = objOCProd.sUsuMontagem
    mvardtDataIniMontagem = objOCProd.dtDataIniMontagem
    mvardtDataFimMontagem = objOCProd.dtDataFimMontagem
    mvardQuantidadePreProd = objOCProd.dQuantidadePreProd
    mvardQuantidadeProd = objOCProd.dQuantidadeProd
End Sub
