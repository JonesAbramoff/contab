VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassReservaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjItemOP As ClassItemOP
Private mvarobjItemMovEst As ClassItemMovEstoque
Private mvarobjProduto As ClassProduto

Public Property Set objItemOP(ByVal vData As ClassItemOP)
    Set mvarobjItemOP = vData
End Property

Public Property Get objItemOP() As ClassItemOP
    Set objItemOP = mvarobjItemOP
End Property

Public Property Set objItemMovEst(ByVal vData As ClassItemMovEstoque)
    Set mvarobjItemMovEst = vData
End Property

Public Property Get objItemMovEst() As ClassItemMovEstoque
    Set objItemMovEst = mvarobjItemMovEst
End Property

Public Property Set objProduto(ByVal vData As ClassProduto)
    Set mvarobjProduto = vData
End Property

Public Property Get objProduto() As ClassProduto
    Set objProduto = mvarobjProduto
End Property
