VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCclHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvardtDataAtualizacao As Date 'local copy
Private mvarsCodigoCcl As String 'local copy
Private mvariTipoCcl As Integer 'local copy
Private mvarsDescCcl As String 'local copy
Public Property Let sDescCcl(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescCcl = 5
    mvarsDescCcl = vData
End Property


Public Property Get sDescCcl() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescCcl
    sDescCcl = mvarsDescCcl
End Property



Public Property Let iTipoCcl(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoCcl = 5
    mvariTipoCcl = vData
End Property


Public Property Get iTipoCcl() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoCcl
    iTipoCcl = mvariTipoCcl
End Property



Public Property Let sCodigoCcl(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodigoCcl = 5
    mvarsCodigoCcl = vData
End Property


Public Property Get sCodigoCcl() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodigoCcl
    sCodigoCcl = mvarsCodigoCcl
End Property



Public Property Let dtDataAtualizacao(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao = 5
    mvardtDataAtualizacao = vData
End Property


Public Property Get dtDataAtualizacao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao
    dtDataAtualizacao = mvardtDataAtualizacao
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



