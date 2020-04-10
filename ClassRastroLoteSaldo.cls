VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRastroLoteSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsProduto As String
Private mvarsLote As String
Private mvariFilialOP As Integer
Private mvariStatus As Integer
Private mvardtDataValidade As Date
Private mvardtDataEntrada As Date
Private mvardtDataFabricacao As Date
Private mvarsObservacao As String
Private mvariAlmoxarifado As Integer
Private mvardQuantDispNossa As Double
Private mvardQuantConsig3 As Double
Private mvardQuantConsig As Double
Private mvardQuantDemo3 As Double
Private mvardQuantDemo As Double
Private mvardQuantConserto3 As Double
Private mvardQuantConserto As Double
Private mvardQuantOutras3 As Double
Private mvardQuantOutras As Double
Private mvardQuantBenef As Double
Private mvardQuantBenef3 As Double

Public Property Let dQuantBenef3(ByVal vData As Double)
    mvardQuantBenef3 = vData
End Property

Public Property Get dQuantBenef3() As Double
    dQuantBenef3 = mvardQuantBenef3
End Property

Public Property Let dQuantBenef(ByVal vData As Double)
    mvardQuantBenef = vData
End Property

Public Property Get dQuantBenef() As Double
    dQuantBenef = mvardQuantBenef
End Property

Public Property Let dQuantOutras(ByVal vData As Double)
    mvardQuantOutras = vData
End Property

Public Property Get dQuantOutras() As Double
    dQuantOutras = mvardQuantOutras
End Property

Public Property Let dQuantOutras3(ByVal vData As Double)
    mvardQuantOutras3 = vData
End Property

Public Property Get dQuantOutras3() As Double
    dQuantOutras3 = mvardQuantOutras3
End Property

Public Property Let dQuantConserto(ByVal vData As Double)
    mvardQuantConserto = vData
End Property

Public Property Get dQuantConserto() As Double
    dQuantConserto = mvardQuantConserto
End Property

Public Property Let dQuantConserto3(ByVal vData As Double)
    mvardQuantConserto3 = vData
End Property

Public Property Get dQuantConserto3() As Double
    dQuantConserto3 = mvardQuantConserto3
End Property

Public Property Let dQuantDemo(ByVal vData As Double)
    mvardQuantDemo = vData
End Property

Public Property Get dQuantDemo() As Double
    dQuantDemo = mvardQuantDemo
End Property

Public Property Let dQuantDemo3(ByVal vData As Double)
    mvardQuantDemo3 = vData
End Property

Public Property Get dQuantDemo3() As Double
    dQuantDemo3 = mvardQuantDemo3
End Property

Public Property Let dQuantConsig(ByVal vData As Double)
    mvardQuantConsig = vData
End Property

Public Property Get dQuantConsig() As Double
    dQuantConsig = mvardQuantConsig
End Property

Public Property Let dQuantConsig3(ByVal vData As Double)
    mvardQuantConsig3 = vData
End Property

Public Property Get dQuantConsig3() As Double
    dQuantConsig3 = mvardQuantConsig3
End Property

Public Property Let dQuantDispNossa(ByVal vData As Double)
    mvardQuantDispNossa = vData
End Property

Public Property Get dQuantDispNossa() As Double
    dQuantDispNossa = mvardQuantDispNossa
End Property

Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dtDataFabricacao(ByVal vData As Date)
    mvardtDataFabricacao = vData
End Property

Public Property Get dtDataFabricacao() As Date
    dtDataFabricacao = mvardtDataFabricacao
End Property

Public Property Let dtDataEntrada(ByVal vData As Date)
    mvardtDataEntrada = vData
End Property

Public Property Get dtDataEntrada() As Date
    dtDataEntrada = mvardtDataEntrada
End Property

Public Property Let dtDataValidade(ByVal vData As Date)
    mvardtDataValidade = vData
End Property

Public Property Get dtDataValidade() As Date
    dtDataValidade = mvardtDataValidade
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iFilialOP(ByVal vData As Integer)
    mvariFilialOP = vData
End Property

Public Property Get iFilialOP() As Integer
    iFilialOP = mvariFilialOP
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property



