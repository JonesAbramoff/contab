VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelDMPLConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariLinha As Integer
Private mvariColuna As Integer
Private mvariItem As Integer
Private mvarsContaInicial As String
Private mvarsContaFinal As String
Private mvarsModelo As String

Public Property Let sModelo(ByVal vData As String)
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let sContaFinal(ByVal vData As String)
    mvarsContaFinal = vData
End Property

Public Property Get sContaFinal() As String
    sContaFinal = mvarsContaFinal
End Property

Public Property Let sContaInicial(ByVal vData As String)
    mvarsContaInicial = vData
End Property

Public Property Get sContaInicial() As String
    sContaInicial = mvarsContaInicial
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

