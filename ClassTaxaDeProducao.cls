VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTaxaDeProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarsProduto As String 'local copy
Private mvarlNumIntDocMaq As Long 'local copy
Private mvarlNumIntDocCompet As Long 'local copy
Private mvardLoteMax As Double 'local copy
Private mvardLoteMin As Double 'local copy
Private mvardLotePadrao As Double 'local copy
Private mvardTempoPreparacao As Double 'local copy
Private mvardTempoMovimentacao As Double 'local copy
Private mvardTempoDescarga As Double 'local copy
Private mvariTipo As Integer 'local copy
Private mvardQuantidade As Double 'local copy
Private mvarsUMProduto As String 'local copy
Private mvardTempoOperacao As Double 'local copy
Private mvarsUMTempo As String 'local copy
Private mvariAtivo As Integer
Private mvardtData As Date
Private mvardtDataDesativacao As Date

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dtDataDesativacao(ByVal vData As Date)
    mvardtDataDesativacao = vData
End Property

Public Property Get dtDataDesativacao() As Date
    dtDataDesativacao = mvardtDataDesativacao
End Property

Public Property Let sUMTempo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMTempo = 5
    mvarsUMTempo = vData
End Property


Public Property Get sUMTempo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMTempo
    sUMTempo = mvarsUMTempo
End Property



Public Property Let dTempoOperacao(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTempoOperacao = 5
    mvardTempoOperacao = vData
End Property


Public Property Get dTempoOperacao() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTempoOperacao
    dTempoOperacao = mvardTempoOperacao
End Property



Public Property Let sUMProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMProduto = 5
    mvarsUMProduto = vData
End Property


Public Property Get sUMProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMProduto
    sUMProduto = mvarsUMProduto
End Property



Public Property Let dQuantidade(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade = 5
    mvardQuantidade = vData
End Property


Public Property Get dQuantidade() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade
    dQuantidade = mvardQuantidade
End Property



Public Property Let iTipo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipo = 5
    mvariTipo = vData
End Property


Public Property Get iTipo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipo
    iTipo = mvariTipo
End Property



Public Property Let dTempoDescarga(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTempoDescarga = 5
    mvardTempoDescarga = vData
End Property


Public Property Get dTempoDescarga() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTempoDescarga
    dTempoDescarga = mvardTempoDescarga
End Property



Public Property Let dTempoMovimentacao(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTempoMovimentacao = 5
    mvardTempoMovimentacao = vData
End Property


Public Property Get dTempoMovimentacao() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTempoMovimentacao
    dTempoMovimentacao = mvardTempoMovimentacao
End Property



Public Property Let dTempoPreparacao(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTempoPreparacao = 5
    mvardTempoPreparacao = vData
End Property


Public Property Get dTempoPreparacao() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTempoPreparacao
    dTempoPreparacao = mvardTempoPreparacao
End Property


Public Property Let dLoteMin(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dLoteMin = 5
    mvardLoteMin = vData
End Property


Public Property Get dLoteMin() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dLoteMin
    dLoteMin = mvardLoteMin
End Property


Public Property Let dLotePadrao(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dLotePadrao = 5
    mvardLotePadrao = vData
End Property


Public Property Get dLotePadrao() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dLotePadrao
    dLotePadrao = mvardLotePadrao
End Property


Public Property Let dLoteMax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dLoteMax = 5
    mvardLoteMax = vData
End Property


Public Property Get dLoteMax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dLoteMax
    dLoteMax = mvardLoteMax
End Property



Public Property Let lNumIntDocCompet(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocCompet = 5
    mvarlNumIntDocCompet = vData
End Property


Public Property Get lNumIntDocCompet() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocCompet
    lNumIntDocCompet = mvarlNumIntDocCompet
End Property



Public Property Let lNumIntDocMaq(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocMaq = 5
    mvarlNumIntDocMaq = vData
End Property


Public Property Get lNumIntDocMaq() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocMaq
    lNumIntDocMaq = mvarlNumIntDocMaq
End Property



Public Property Let sProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProduto = 5
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProduto
    sProduto = mvarsProduto
End Property



Public Property Let lNumIntDoc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDoc = 5
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDoc
    lNumIntDoc = mvarlNumIntDoc
End Property



