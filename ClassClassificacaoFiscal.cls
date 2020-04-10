VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassClassificacaoFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsDescricao As String 'local copy
Private mvarsCodigo As String 'local copy

Private mvardIIAliquota As Double 'aliquota de imposto de importacao
Private mvardCOFINSAliquota As Double 'aliquota de imposto de importacao
Private mvardPISAliquota As Double 'aliquota de imposto de importacao
Private mvardICMSAliquota As Double 'aliquota de imposto de importacao
Private mvardIPIAliquota As Double 'aliquota de imposto de importacao

Public Property Let dIPIAliquota(ByVal vData As Double)
    mvardIPIAliquota = vData
End Property

Public Property Get dIPIAliquota() As Double
    dIPIAliquota = mvardIPIAliquota
End Property

Public Property Let dICMSAliquota(ByVal vData As Double)
    mvardICMSAliquota = vData
End Property

Public Property Get dICMSAliquota() As Double
    dICMSAliquota = mvardICMSAliquota
End Property

Public Property Let dPISAliquota(ByVal vData As Double)
    mvardPISAliquota = vData
End Property

Public Property Get dPISAliquota() As Double
    dPISAliquota = mvardPISAliquota
End Property

Public Property Let dCOFINSAliquota(ByVal vData As Double)
    mvardCOFINSAliquota = vData
End Property

Public Property Get dCOFINSAliquota() As Double
    dCOFINSAliquota = mvardCOFINSAliquota
End Property

Public Property Let dIIAliquota(ByVal vData As Double)
    mvardIIAliquota = vData
End Property

Public Property Get dIIAliquota() As Double
    dIIAliquota = mvardIIAliquota
End Property

Public Property Let sCodigo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodigo = 5
    mvarsCodigo = vData
End Property


Public Property Get sCodigo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodigo
    sCodigo = mvarsCodigo
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



