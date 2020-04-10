VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTribComplNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvariTipo As Integer 'local copy
Private mvarsNaturezaOp As String 'local copy
Private mvariTipoTributacao As Integer 'local copy
Private mvariIPITipo As Integer 'local copy
'Private mvarsIPICodProduto As String 'local copy
Private mvardIPIBaseCalculo As Double 'local copy
Private mvardIPIPercRedBase As Double 'local copy
Private mvardIPIAliquota As Double 'local copy
Private mvardIPIValor As Double 'local copy
Private mvardIPICredito As Double 'local copy
Private mvariICMSTipo As Integer 'local copy
Private mvardICMSBase As Double 'local copy
Private mvardICMSPercRedBase As Double 'local copy
Private mvardICMSAliquota As Double 'local copy
Private mvardICMSValor As Double 'local copy
Private mvardICMSCredito As Double 'local copy
Private mvardICMSSubstBase As Double 'local copy
Private mvardICMSSubstAliquota As Double 'local copy
Private mvardICMSSubstValor As Double 'local copy

Private mvariNaturezaOpManual As Integer
Private mvariTipoTributacaoManual As Integer
Private mvariIPITipoManual As Integer
Private mvariIPIBaseManual As Integer
Private mvariIPIPercRedBaseManual As Integer
Private mvariIPIAliquotaManual As Integer
Private mvariIPIValorManual As Integer
Private mvariICMSTipoManual As Integer
Private mvariICMSBaseManual As Integer
Private mvariICMSPercRedBaseManual As Integer
Private mvariICMSAliquotaManual As Integer
Private mvariICMSValorManual As Integer
Private mvariICMSSubstBaseManual As Integer
Private mvariICMSSubstAliquotaManual As Integer
Private mvariICMSSubstValorManual As Integer

Private mvardPISCredito As Double
Private mvardCOFINSCredito As Double

Public Property Let dCOFINSCredito(ByVal vData As Double)
    mvardCOFINSCredito = vData
End Property

Public Property Get dCOFINSCredito() As Double
    dCOFINSCredito = mvardCOFINSCredito
End Property

Public Property Let dPISCredito(ByVal vData As Double)
    mvardPISCredito = vData
End Property

Public Property Get dPISCredito() As Double
    dPISCredito = mvardPISCredito
End Property

Public Property Let iICMSSubstAliquotaManual(ByVal vData As Integer)
    mvariICMSSubstAliquotaManual = vData
End Property

Public Property Get iICMSSubstAliquotaManual() As Integer
    iICMSSubstAliquotaManual = mvariICMSSubstAliquotaManual
End Property

Public Property Let iICMSSubstValorManual(ByVal vData As Integer)
    mvariICMSSubstValorManual = vData
End Property

Public Property Get iICMSSubstValorManual() As Integer
    iICMSSubstValorManual = mvariICMSSubstValorManual
End Property

Public Property Let iICMSSubstBaseManual(ByVal vData As Integer)
    mvariICMSSubstBaseManual = vData
End Property

Public Property Get iICMSSubstBaseManual() As Integer
    iICMSSubstBaseManual = mvariICMSSubstBaseManual
End Property

Public Property Let iICMSValorManual(ByVal vData As Integer)
    mvariICMSValorManual = vData
End Property

Public Property Get iICMSValorManual() As Integer
    iICMSValorManual = mvariICMSValorManual
End Property

Public Property Let iICMSAliquotaManual(ByVal vData As Integer)
    mvariICMSAliquotaManual = vData
End Property

Public Property Get iICMSAliquotaManual() As Integer
    iICMSAliquotaManual = mvariICMSAliquotaManual
End Property

Public Property Let iICMSPercRedBaseManual(ByVal vData As Integer)
    mvariICMSPercRedBaseManual = vData
End Property

Public Property Get iICMSPercRedBaseManual() As Integer
    iICMSPercRedBaseManual = mvariICMSPercRedBaseManual
End Property

Public Property Let iICMSBaseManual(ByVal vData As Integer)
    mvariICMSBaseManual = vData
End Property

Public Property Get iICMSBaseManual() As Integer
    iICMSBaseManual = mvariICMSBaseManual
End Property

Public Property Let iICMSTipoManual(ByVal vData As Integer)
    mvariICMSTipoManual = vData
End Property

Public Property Get iICMSTipoManual() As Integer
    iICMSTipoManual = mvariICMSTipoManual
End Property

Public Property Let iIPIValorManual(ByVal vData As Integer)
    mvariIPIValorManual = vData
End Property

Public Property Get iIPIValorManual() As Integer
    iIPIValorManual = mvariIPIValorManual
End Property

Public Property Let iIPIAliquotaManual(ByVal vData As Integer)
    mvariIPIAliquotaManual = vData
End Property

Public Property Get iIPIAliquotaManual() As Integer
    iIPIAliquotaManual = mvariIPIAliquotaManual
End Property

Public Property Let iIPITipoManual(ByVal vData As Integer)
    mvariIPITipoManual = vData
End Property

Public Property Get iIPITipoManual() As Integer
    iIPITipoManual = mvariIPITipoManual
End Property

Public Property Let iTipoTributacaoManual(ByVal vData As Integer)
    mvariTipoTributacaoManual = vData
End Property

Public Property Get iTipoTributacaoManual() As Integer
    iTipoTributacaoManual = mvariTipoTributacaoManual
End Property

Public Property Let iIPIPercRedBaseManual(ByVal vData As Integer)
    mvariIPIPercRedBaseManual = vData
End Property

Public Property Get iIPIPercRedBaseManual() As Integer
    iIPIPercRedBaseManual = mvariIPIPercRedBaseManual
End Property

Public Property Let iNaturezaOpManual(ByVal vData As Integer)
    mvariNaturezaOpManual = vData
End Property

Public Property Get iNaturezaOpManual() As Integer
    iNaturezaOpManual = mvariNaturezaOpManual
End Property

Public Property Let iIPIBaseManual(ByVal vData As Integer)
    mvariIPIBaseManual = vData
End Property

Public Property Get iIPIBaseManual() As Integer
    iIPIBaseManual = mvariIPIBaseManual
End Property

Public Property Let dICMSSubstValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSSubstValor = 5
    mvardICMSSubstValor = vData
End Property

Public Property Get dICMSSubstValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSSubstValor
    dICMSSubstValor = mvardICMSSubstValor
End Property

Public Property Let dICMSSubstAliquota(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSSubstAliquota = 5
    mvardICMSSubstAliquota = vData
End Property


Public Property Get dICMSSubstAliquota() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSSubstAliquota
    dICMSSubstAliquota = mvardICMSSubstAliquota
End Property



Public Property Let dICMSSubstBase(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSSubstBase = 5
    mvardICMSSubstBase = vData
End Property


Public Property Get dICMSSubstBase() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSSubstBase
    dICMSSubstBase = mvardICMSSubstBase
End Property



'Public Property Let iICMSCredita(ByVal vData As Integer)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.iICMSCredita = 5
'    mvariICMSCredita = vData
'End Property
'
'
'Public Property Get iICMSCredita() As Integer
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.iICMSCredita
'    iICMSCredita = mvariICMSCredita
'End Property



Public Property Let dICMSValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSValor = 5
    mvardICMSValor = vData
End Property


Public Property Get dICMSValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSValor
    dICMSValor = mvardICMSValor
End Property



Public Property Let dICMSAliquota(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSAliquota = 5
    mvardICMSAliquota = vData
End Property


Public Property Get dICMSAliquota() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSAliquota
    dICMSAliquota = mvardICMSAliquota
End Property



Public Property Let dICMSPercRedBase(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSPercRedBase = 5
    mvardICMSPercRedBase = vData
End Property


Public Property Get dICMSPercRedBase() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSPercRedBase
    dICMSPercRedBase = mvardICMSPercRedBase
End Property



Public Property Let dICMSBase(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dICMSBase = 5
    mvardICMSBase = vData
End Property


Public Property Get dICMSBase() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dICMSBase
    dICMSBase = mvardICMSBase
End Property



Public Property Let iICMSTipo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iICMSTipo = 5
    mvariICMSTipo = vData
End Property


Public Property Get iICMSTipo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iICMSTipo
    iICMSTipo = mvariICMSTipo
End Property



Public Property Let dIPICredito(ByVal vData As Double)
    mvardIPICredito = vData
End Property


Public Property Get dIPICredito() As Double
    dIPICredito = mvardIPICredito
End Property

Public Property Let dICMSCredito(ByVal vData As Double)
    mvardICMSCredito = vData
End Property


Public Property Get dICMSCredito() As Double
    dICMSCredito = mvardICMSCredito
End Property



Public Property Let dIPIValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dIPIValor = 5
    mvardIPIValor = vData
End Property


Public Property Get dIPIValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dIPIValor
    dIPIValor = mvardIPIValor
End Property



Public Property Let dIPIAliquota(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dIPIAliquota = 5
    mvardIPIAliquota = vData
End Property


Public Property Get dIPIAliquota() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dIPIAliquota
    dIPIAliquota = mvardIPIAliquota
End Property



Public Property Let dIPIPercRedBase(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dIPIPercRedBase = 5
    mvardIPIPercRedBase = vData
End Property


Public Property Get dIPIPercRedBase() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dIPIPercRedBase
    dIPIPercRedBase = mvardIPIPercRedBase
End Property



Public Property Let dIPIBaseCalculo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dIPIBaseCalculo = 5
    mvardIPIBaseCalculo = vData
End Property


Public Property Get dIPIBaseCalculo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dIPIBaseCalculo
    dIPIBaseCalculo = mvardIPIBaseCalculo
End Property



''Public Property Let sIPICodProduto(ByVal vData As String)
'''used when assigning a value to the property, on the left side of an assignment.
'''Syntax: X.sIPICodProduto = 5
''    mvarsIPICodProduto = vData
''End Property
''
''
''Public Property Get sIPICodProduto() As String
'''used when retrieving value of a property, on the right side of an assignment.
'''Syntax: Debug.Print X.sIPICodProduto
''    sIPICodProduto = mvarsIPICodProduto
''End Property
''


Public Property Let iIPITipo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iIPITipo = 5
    mvariIPITipo = vData
End Property


Public Property Get iIPITipo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iIPITipo
    iIPITipo = mvariIPITipo
End Property



Public Property Let iTipoTributacao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoTributacao = 5
    mvariTipoTributacao = vData
End Property


Public Property Get iTipoTributacao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoTributacao
    iTipoTributacao = mvariTipoTributacao
End Property



Public Property Let sNaturezaOp(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNaturezaOp = 5
    mvarsNaturezaOp = vData
End Property


Public Property Get sNaturezaOp() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNaturezaOp
    sNaturezaOp = mvarsNaturezaOp
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



