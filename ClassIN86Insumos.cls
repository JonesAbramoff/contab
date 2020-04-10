VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIN86Insumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvardtData As Date 'local copy
Private mvarsProdutoPai As String 'local copy
Private mvarsVersao As String 'local copy
Private mvariNivel As Integer 'local copy
Private mvariSeq As Integer 'local copy
Private mvarsProduto As String 'local copy
Private mvariSeqPai As Integer 'local copy
Private mvariComposicao As Integer 'local copy
Private mvardQuantidade As Double 'local copy
Private mvarsUnidadeMedInsumo As String 'local copy
Private mvardPercentualPerda As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarsUnidadeMedPai As String 'local copy
Public Property Let sUnidadeMedPai(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUnidadeMedPai = 5
    mvarsUnidadeMedPai = vData
End Property


Public Property Get sUnidadeMedPai() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUnidadeMedPai
    sUnidadeMedPai = mvarsUnidadeMedPai
End Property




Public Property Let dPercentualPerda(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercentualPerda = 5
    mvardPercentualPerda = vData
End Property


Public Property Get dPercentualPerda() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPercentualPerda
    dPercentualPerda = mvardPercentualPerda
End Property



Public Property Let sUnidadeMedInsumo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUnidadeMed = 5
    mvarsUnidadeMedInsumo = vData
End Property


Public Property Get sUnidadeMedInsumo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUnidadeMed
    sUnidadeMedInsumo = mvarsUnidadeMedInsumo
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



Public Property Let iComposicao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iComposicao = 5
    mvariComposicao = vData
End Property


Public Property Get iComposicao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iComposicao
    iComposicao = mvariComposicao
End Property



Public Property Let iSeqPai(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSeqPai = 5
    mvariSeqPai = vData
End Property


Public Property Get iSeqPai() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSeqPai
    iSeqPai = mvariSeqPai
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



Public Property Let iSeq(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSeq = 5
    mvariSeq = vData
End Property


Public Property Get iSeq() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSeq
    iSeq = mvariSeq
End Property



Public Property Let iNivel(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iNivel = 5
    mvariNivel = vData
End Property


Public Property Get iNivel() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iNivel
    iNivel = mvariNivel
End Property



Public Property Let sVersao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sVersao = 5
    mvarsVersao = vData
End Property


Public Property Get sVersao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sVersao
    sVersao = mvarsVersao
End Property



Public Property Let sProdutoPai(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProdutoPai = 5
    mvarsProdutoPai = vData
End Property


Public Property Get sProdutoPai() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProdutoPai
    sProdutoPai = mvarsProdutoPai
End Property

Public Property Let dtData(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtData = 5
    mvardtData = vData
End Property


Public Property Get dtData() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtData
    dtData = mvardtData
End Property
