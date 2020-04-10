VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTecladoProdutoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsProduto As String
Private mvarsTitulo As String
Private mvarlColor As String
Private mvariTecla As String
Private mvariIndice As String
Private mvarsReferencia As String
Private mvarsCodigoBarras As String
Private mvarsArvoreKey As String

Public Property Let sArvoreKey(ByVal vData As String)
    mvarsArvoreKey = vData
End Property

Public Property Get sArvoreKey() As String
    sArvoreKey = mvarsArvoreKey
End Property

 
Public Property Let sReferencia(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCaption = 5
    mvarsReferencia = vData
End Property

Public Property Get sReferencia() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCaption
    sReferencia = mvarsReferencia
End Property
 
Public Property Let sCodigoBarras(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCaption = 5
    mvarsCodigoBarras = vData
End Property

Public Property Get sCodigoBarras() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCaption
    sCodigoBarras = mvarsCodigoBarras
End Property
 
Public Property Let sTitulo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCaption = 5
    mvarsTitulo = vData
End Property

Public Property Get sTitulo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCaption
    sTitulo = mvarsTitulo
End Property

Public Property Let iTecla(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTecla2 = 5
    mvariTecla = vData
End Property

Public Property Get iTecla() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTecla2
    iTecla = mvariTecla
End Property

Public Property Let iIndice(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iIndice2 = 5
    mvariIndice = vData
End Property

Public Property Get iIndice() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iIndice2
    iIndice = mvariIndice
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

Public Property Let lColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lColor = 5
    mvarlColor = vData
End Property

Public Property Get lColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lColor
    lColor = mvarlColor
End Property


