VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMaquinaOperadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarlNumIntDocMaq As Long 'local copy
Private mvariTipoMaoDeObra As Integer 'local copy
Private mvariQuantidade As Integer 'local copy
Private mvardPercentualUso As Double 'local copy
Public Property Let dPercentualUso(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercentualUso = 5
    mvardPercentualUso = vData
End Property


Public Property Get dPercentualUso() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPercentualUso
    dPercentualUso = mvardPercentualUso
End Property



Public Property Let iQuantidade(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iQuantidade = 5
    mvariQuantidade = vData
End Property


Public Property Get iQuantidade() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iQuantidade
    iQuantidade = mvariQuantidade
End Property



Public Property Let iTipoMaoDeObra(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoMaoDeObra = 5
    mvariTipoMaoDeObra = vData
End Property


Public Property Get iTipoMaoDeObra() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoMaoDeObra
    iTipoMaoDeObra = mvariTipoMaoDeObra
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



