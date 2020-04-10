VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelDMPLFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariLinha As Integer
Private mvariColuna As Integer
Private mvariItem As Integer
Private mvariOperacao As Integer
Private mvarsModelo As String
Private mvariLinhaFormula As Integer
Private mvariColunaFormula As Integer

Public Property Let iColunaFormula(ByVal vData As Integer)
    mvariColunaFormula = vData
End Property

Public Property Get iColunaFormula() As Integer
    iColunaFormula = mvariColunaFormula
End Property

Public Property Let iLinhaFormula(ByVal vData As Integer)
    mvariLinhaFormula = vData
End Property

Public Property Get iLinhaFormula() As Integer
    iLinhaFormula = mvariLinhaFormula
End Property

Public Property Let sModelo(ByVal vData As String)
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let iOperacao(ByVal vData As Integer)
    mvariOperacao = vData
End Property

Public Property Get iOperacao() As Integer
    iOperacao = mvariOperacao
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let iColuna(ByVal vData As Integer)
    mvariColuna = vData
End Property

Public Property Get iColuna() As Integer
    iColuna = mvariColuna
End Property

Public Property Let iLinha(ByVal vData As Integer)
    mvariLinha = vData
End Property

Public Property Get iLinha() As Integer
    iLinha = mvariLinha
End Property

