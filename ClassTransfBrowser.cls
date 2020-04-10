VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTransfBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlTransferencia As Long 'local copy
Private mvariFilialEmpresa As Integer 'local copy
Private mvardtDataMovimento As Date 'local copy
Private mvardValor As Double 'local copy
Private mvariTipoMovtoOrigem As Integer 'local copy
Private mvariTipoMovtoDestino As Integer 'local copy
Private mvarsTipoMovtoOrigemNomeRed As String 'local copy
Private mvarsTipoMovtoDestinoNomeRed As String 'local copy

Public Property Let iTipoMovtoDestino(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoMovtoDestino = 5
    mvariTipoMovtoDestino = vData
End Property


Public Property Get iTipoMovtoDestino() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoMovtoDestino
    iTipoMovtoDestino = mvariTipoMovtoDestino
End Property

Public Property Let sTipoMovtoDestinoNomeRed(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
    mvarsTipoMovtoDestinoNomeRed = vData
End Property

Public Property Get sTipoMovtoDestinoNomeRed() As String
'used when retrieving value of a property, on the right side of an assignment.
    sTipoMovtoDestinoNomeRed = mvarsTipoMovtoDestinoNomeRed
End Property

Public Property Let sTipoMovtoOrigemNomeRed(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
    mvarsTipoMovtoOrigemNomeRed = vData
End Property

Public Property Get sTipoMovtoOrigemNomeRed() As String
'used when retrieving value of a property, on the right side of an assignment.
    sTipoMovtoOrigemNomeRed = mvarsTipoMovtoOrigemNomeRed
End Property

Public Property Let iTipoMovtoOrigem(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoMovtoOrigem = 5
    mvariTipoMovtoOrigem = vData
End Property


Public Property Get iTipoMovtoOrigem() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoMovtoOrigem
    iTipoMovtoOrigem = mvariTipoMovtoOrigem
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



Public Property Let dtDataMovimento(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataMovimento = 5
    mvardtDataMovimento = vData
End Property


Public Property Get dtDataMovimento() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataMovimento
    dtDataMovimento = mvardtDataMovimento
End Property



Public Property Let iFilialEmpresa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property


Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property



Public Property Let lTransferencia(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lTransferencia = 5
    mvarlTransferencia = vData
End Property


Public Property Get lTransferencia() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lTransferencia
    lTransferencia = mvarlTransferencia
End Property



