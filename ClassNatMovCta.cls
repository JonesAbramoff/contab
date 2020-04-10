VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNatMovCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsCodigo As String 'local copy
Private mvarsDescricao As String 'local copy
Private mvariTipo As Integer 'local copy
Private mvariFluxoCaixa As Integer 'local copy
Private mvarlGrupo As Long

'AUXILIARES
Private mvarsDescGrupo As String
Private mvardCredito As Double
Private mvardDebito As Double

Public Property Let sDescGrupo(ByVal vData As String)
    mvarsDescGrupo = vData
End Property
Public Property Get sDescGrupo() As String
    sDescGrupo = mvarsDescGrupo
End Property

Public Property Let dCredito(ByVal vData As Double)
    mvardCredito = vData
End Property
Public Property Get dCredito() As Double
    dCredito = mvardCredito
End Property

Public Property Let dDebito(ByVal vData As Double)
    mvardDebito = vData
End Property
Public Property Get dDebito() As Double
    dDebito = mvardDebito
End Property

'AUXILIARES

Public Property Let lGrupo(ByVal vData As Long)
    mvarlGrupo = vData
End Property
Public Property Get lGrupo() As Long
    lGrupo = mvarlGrupo
End Property

Public Property Let iFluxoCaixa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFluxoCaixa = 5
    mvariFluxoCaixa = vData
End Property


Public Property Get iFluxoCaixa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFluxoCaixa
    iFluxoCaixa = mvariFluxoCaixa
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



Public Property Let sDescricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsDescricao = vData
End Property


Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sDescricao = mvarsDescricao
End Property



Public Property Let sCodigo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodigo = 5
    mvarsCodigo = vData
End Property


Public Property Get sCodigo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodigo
    sCodigo = mvarsCodigo
End Property




