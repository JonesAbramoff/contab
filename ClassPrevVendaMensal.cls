VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPrevVendaMensal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Private mvarsCodigo As String 'local copy
Private mvariAno As Integer 'local copy
Private mvariCodRegiao As Integer 'local copy
Private mvarlCliente As Long 'local copy
Private mvariFilial As Integer 'local copy
Private mvarsProduto As String 'local copy
Private mvardQuantidade1 As Double 'local copy
Private mvardValor1 As Double 'local copy
Private mvardQuantidade2 As Double 'local copy
Private mvardQuantidade3 As Double 'local copy
Private mvardQuantidade4 As Double 'local copy
Private mvardQuantidade5 As Double 'local copy
Private mvardQuantidade6 As Double 'local copy
Private mvardQuantidade7 As Double 'local copy
Private mvardQuantidade8 As Double 'local copy
Private mvardQuantidade9 As Double 'local copy
Private mvardQuantidade10 As Double 'local copy
Private mvardQuantidade11 As Double 'local copy
Private mvardQuantidade12 As Double 'local copy
Private mvardValor2 As Double 'local copy
Private mvardValor3 As Double 'local copy
Private mvardValor4 As Double 'local copy
Private mvardValor5 As Double 'local copy
Private mvardValor6 As Double 'local copy
Private mvardValor7 As Double 'local copy
Private mvardValor8 As Double 'local copy
Private mvardValor9 As Double 'local copy
Private mvardValor10 As Double 'local copy
Private mvardValor11 As Double 'local copy
Private mvardValor12 As Double 'local copy
Private mvardtDataAtualizacao1 As Date 'local copy
Private mvardtDataAtualizacao2 As Date 'local copy
Private mvardtDataAtualizacao3 As Date 'local copy
Private mvardtDataAtualizacao5 As Date 'local copy
Private mvardtDataAtualizacao6 As Date 'local copy
Private mvardtDataAtualizacao4 As Date 'local copy
Private mvardtDataAtualizacao7 As Date 'local copy
Private mvardtDataAtualizacao8 As Date 'local copy
Private mvardtDataAtualizacao9 As Date 'local copy
Private mvardtDataAtualizacao10 As Date 'local copy
Private mvardtDataAtualizacao11 As Date 'local copy
Private mvardtDataAtualizacao12 As Date 'local copy
Public Property Let dtDataAtualizacao12(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao12 = 5
    mvardtDataAtualizacao12 = vData
End Property


Public Property Get dtDataAtualizacao12() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao12
    dtDataAtualizacao12 = mvardtDataAtualizacao12
End Property



Public Property Let dtDataAtualizacao11(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao11 = 5
    mvardtDataAtualizacao11 = vData
End Property


Public Property Get dtDataAtualizacao11() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao11
    dtDataAtualizacao11 = mvardtDataAtualizacao11
End Property



Public Property Let dtDataAtualizacao10(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao10 = 5
    mvardtDataAtualizacao10 = vData
End Property


Public Property Get dtDataAtualizacao10() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao10
    dtDataAtualizacao10 = mvardtDataAtualizacao10
End Property



Public Property Let dtDataAtualizacao9(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao9 = 5
    mvardtDataAtualizacao9 = vData
End Property


Public Property Get dtDataAtualizacao9() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao9
    dtDataAtualizacao9 = mvardtDataAtualizacao9
End Property



Public Property Let dtDataAtualizacao8(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao8 = 5
    mvardtDataAtualizacao8 = vData
End Property


Public Property Get dtDataAtualizacao8() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao8
    dtDataAtualizacao8 = mvardtDataAtualizacao8
End Property



Public Property Let dtDataAtualizacao7(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao7 = 5
    mvardtDataAtualizacao7 = vData
End Property


Public Property Get dtDataAtualizacao7() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao7
    dtDataAtualizacao7 = mvardtDataAtualizacao7
End Property



Public Property Let dtDataAtualizacao4(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao4 = 5
    mvardtDataAtualizacao4 = vData
End Property


Public Property Get dtDataAtualizacao4() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao4
    dtDataAtualizacao4 = mvardtDataAtualizacao4
End Property



Public Property Let dtDataAtualizacao6(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao6 = 5
    mvardtDataAtualizacao6 = vData
End Property


Public Property Get dtDataAtualizacao6() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao6
    dtDataAtualizacao6 = mvardtDataAtualizacao6
End Property



Public Property Let dtDataAtualizacao5(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao5 = 5
    mvardtDataAtualizacao5 = vData
End Property


Public Property Get dtDataAtualizacao5() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao5
    dtDataAtualizacao5 = mvardtDataAtualizacao5
End Property



Public Property Let dtDataAtualizacao3(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao3 = 5
    mvardtDataAtualizacao3 = vData
End Property


Public Property Get dtDataAtualizacao3() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao3
    dtDataAtualizacao3 = mvardtDataAtualizacao3
End Property



Public Property Let dtDataAtualizacao2(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao2 = 5
    mvardtDataAtualizacao2 = vData
End Property


Public Property Get dtDataAtualizacao2() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao2
    dtDataAtualizacao2 = mvardtDataAtualizacao2
End Property



Public Property Let dtDataAtualizacao1(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAtualizacao1 = 5
    mvardtDataAtualizacao1 = vData
End Property


Public Property Get dtDataAtualizacao1() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAtualizacao1
    dtDataAtualizacao1 = mvardtDataAtualizacao1
End Property



Public Property Let dValor12(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor12 = 5
    mvardValor12 = vData
End Property


Public Property Get dValor12() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor12
    dValor12 = mvardValor12
End Property



Public Property Let dValor11(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor11 = 5
    mvardValor11 = vData
End Property


Public Property Get dValor11() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor11
    dValor11 = mvardValor11
End Property



Public Property Let dValor10(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor10 = 5
    mvardValor10 = vData
End Property


Public Property Get dValor10() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor10
    dValor10 = mvardValor10
End Property



Public Property Let dValor9(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor9 = 5
    mvardValor9 = vData
End Property


Public Property Get dValor9() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor9
    dValor9 = mvardValor9
End Property



Public Property Let dValor8(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor8 = 5
    mvardValor8 = vData
End Property


Public Property Get dValor8() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor8
    dValor8 = mvardValor8
End Property



Public Property Let dValor7(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor7 = 5
    mvardValor7 = vData
End Property


Public Property Get dValor7() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor7
    dValor7 = mvardValor7
End Property



Public Property Let dValor6(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor6 = 5
    mvardValor6 = vData
End Property


Public Property Get dValor6() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor6
    dValor6 = mvardValor6
End Property



Public Property Let dValor5(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor5 = 5
    mvardValor5 = vData
End Property


Public Property Get dValor5() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor5
    dValor5 = mvardValor5
End Property



Public Property Let dValor4(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor4 = 5
    mvardValor4 = vData
End Property


Public Property Get dValor4() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor4
    dValor4 = mvardValor4
End Property



Public Property Let dValor3(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor3 = 5
    mvardValor3 = vData
End Property


Public Property Get dValor3() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor3
    dValor3 = mvardValor3
End Property



Public Property Let dValor2(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor2 = 5
    mvardValor2 = vData
End Property


Public Property Get dValor2() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor2
    dValor2 = mvardValor2
End Property



Public Property Let dQuantidade12(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade12 = 5
    mvardQuantidade12 = vData
End Property


Public Property Get dQuantidade12() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade12
    dQuantidade12 = mvardQuantidade12
End Property



Public Property Let dQuantidade11(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade11 = 5
    mvardQuantidade11 = vData
End Property


Public Property Get dQuantidade11() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade11
    dQuantidade11 = mvardQuantidade11
End Property



Public Property Let dQuantidade10(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade10 = 5
    mvardQuantidade10 = vData
End Property


Public Property Get dQuantidade10() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade10
    dQuantidade10 = mvardQuantidade10
End Property



Public Property Let dQuantidade9(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade9 = 5
    mvardQuantidade9 = vData
End Property


Public Property Get dQuantidade9() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade9
    dQuantidade9 = mvardQuantidade9
End Property



Public Property Let dQuantidade8(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade8 = 5
    mvardQuantidade8 = vData
End Property


Public Property Get dQuantidade8() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade8
    dQuantidade8 = mvardQuantidade8
End Property



Public Property Let dQuantidade7(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade7 = 5
    mvardQuantidade7 = vData
End Property


Public Property Get dQuantidade7() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade7
    dQuantidade7 = mvardQuantidade7
End Property



Public Property Let dQuantidade6(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade6 = 5
    mvardQuantidade6 = vData
End Property


Public Property Get dQuantidade6() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade6
    dQuantidade6 = mvardQuantidade6
End Property



Public Property Let dQuantidade5(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade5 = 5
    mvardQuantidade5 = vData
End Property


Public Property Get dQuantidade5() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade5
    dQuantidade5 = mvardQuantidade5
End Property



Public Property Let dQuantidade4(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade4 = 5
    mvardQuantidade4 = vData
End Property


Public Property Get dQuantidade4() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade4
    dQuantidade4 = mvardQuantidade4
End Property



Public Property Let dQuantidade3(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade3 = 5
    mvardQuantidade3 = vData
End Property


Public Property Get dQuantidade3() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade3
    dQuantidade3 = mvardQuantidade3
End Property



Public Property Let dQuantidade2(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade2 = 5
    mvardQuantidade2 = vData
End Property


Public Property Get dQuantidade2() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade2
    dQuantidade2 = mvardQuantidade2
End Property



Public Property Let dValor1(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor1 = 5
    mvardValor1 = vData
End Property


Public Property Get dValor1() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor1
    dValor1 = mvardValor1
End Property



Public Property Let dQuantidade1(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade1 = 5
    mvardQuantidade1 = vData
End Property


Public Property Get dQuantidade1() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade1
    dQuantidade1 = mvardQuantidade1
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



Public Property Let iFilial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilial = 5
    mvariFilial = vData
End Property


Public Property Get iFilial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilial
    iFilial = mvariFilial
End Property



Public Property Let lCliente(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCliente = 5
    mvarlCliente = vData
End Property


Public Property Get lCliente() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCliente
    lCliente = mvarlCliente
End Property



Public Property Let iCodRegiao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCodRegiao = 5
    mvariCodRegiao = vData
End Property


Public Property Get iCodRegiao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCodRegiao
    iCodRegiao = mvariCodRegiao
End Property



Public Property Let iAno(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAno = 5
    mvariAno = vData
End Property


Public Property Get iAno() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAno
    iAno = mvariAno
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



