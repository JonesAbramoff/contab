VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFluxoCxItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable(s) to hold property value(s)
Private mvariSequencial As Integer 'local copy
Private mvariTipo As Integer 'local copy
Private mvariFluxoCaixa As Integer 'local copy
Private mvarsDescricao As String 'local copy
Private mvardValor As Double 'local copy
Private mvariNumParcelas As Integer 'local copy
Private mvariNumParcela As Integer 'local copy
Private mvariTipoNumIntDocOrigem As Integer 'local copy
Private mvarlNumIntDocOrigem As Long 'local copy
Public Property Let lNumIntDocOrigem(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocOrigem = 5
    mvarlNumIntDocOrigem = vData
End Property


Public Property Get lNumIntDocOrigem() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocOrigem
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property



Public Property Let iTipoNumIntDocOrigem(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoNumIntDocOrigem = 5
    mvariTipoNumIntDocOrigem = vData
End Property


Public Property Get iTipoNumIntDocOrigem() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoNumIntDocOrigem
    iTipoNumIntDocOrigem = mvariTipoNumIntDocOrigem
End Property



Public Property Let iNumParcela(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iNumParcela = 5
    mvariNumParcela = vData
End Property


Public Property Get iNumParcela() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iNumParcela
    iNumParcela = mvariNumParcela
End Property



Public Property Let iNumParcelas(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iNumParcelas = 5
    mvariNumParcelas = vData
End Property


Public Property Get iNumParcelas() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iNumParcelas
    iNumParcelas = mvariNumParcelas
End Property



Public Property Let dValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor = 5
    mvardValor = vData
End Property


Public Property Get dValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor
    dValor = mvardValor
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



Public Property Let iSequencial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSequencial = 5
    mvariSequencial = vData
End Property


Public Property Get iSequencial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSequencial
    iSequencial = mvariSequencial
End Property




