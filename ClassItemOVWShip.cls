VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOVWShip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardPrecoUnitImp As Double
Private mvariTabelaPrecoImp As Integer

Public Property Let iTabelaPrecoImp(ByVal vData As Integer)
    mvariTabelaPrecoImp = vData
End Property

Public Property Get iTabelaPrecoImp() As Integer
    iTabelaPrecoImp = mvariTabelaPrecoImp
End Property

Public Property Let dPrecoUnitImp(ByVal vData As Double)
    mvardPrecoUnitImp = vData
End Property

Public Property Get dPrecoUnitImp() As Double
    dPrecoUnitImp = mvardPrecoUnitImp
End Property

