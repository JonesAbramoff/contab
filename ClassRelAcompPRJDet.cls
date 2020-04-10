VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelAcompPRJDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel As Long
Private mvariSeqAcomp As Integer
Private mvariSeq As Integer
Private mvariTipo As Integer
Private mvarsTipoTexto As String
Private mvarsItem As String
Private mvarsDescricao As String
Private mvardCustoPrev As Double
Private mvardCustoReal As Double
Private mvarsUM As String
Private mvardQtdPrev As Double
Private mvardQtdReal As Double
Private mvarsObs As String

Public Property Let sObs(ByVal vData As String)
    mvarsObs = vData
End Property

Public Property Get sObs() As String
    sObs = mvarsObs
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let iSeqAcomp(ByVal vData As Integer)
    mvariSeqAcomp = vData
End Property

Public Property Get iSeqAcomp() As Integer
    iSeqAcomp = mvariSeqAcomp
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sTipoTexto(ByVal vData As String)
    mvarsTipoTexto = vData
End Property

Public Property Get sTipoTexto() As String
    sTipoTexto = mvarsTipoTexto
End Property

Public Property Let sItem(ByVal vData As String)
    mvarsItem = vData
End Property

Public Property Get sItem() As String
    sItem = mvarsItem
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dCustoPrev(ByVal vData As Double)
    mvardCustoPrev = vData
End Property

Public Property Get dCustoPrev() As Double
    dCustoPrev = mvardCustoPrev
End Property

Public Property Let dCustoReal(ByVal vData As Double)
    mvardCustoReal = vData
End Property

Public Property Get dCustoReal() As Double
    dCustoReal = mvardCustoReal
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dQtdPrev(ByVal vData As Double)
    mvardQtdPrev = vData
End Property

Public Property Get dQtdPrev() As Double
    dQtdPrev = mvardQtdPrev
End Property

Public Property Let dQtdReal(ByVal vData As Double)
    mvardQtdReal = vData
End Property

Public Property Get dQtdReal() As Double
    dQtdReal = mvardQtdReal
End Property
