VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPlanMargContrLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariLinha As Integer
Private mvarsDescricao As String
Private mvarsFormulaGeral As String
Private mvarsFormulaL1 As String
Private mvariFormato As Integer
Private mvariEditavel As Integer

Public Property Let iEditavel(ByVal vData As Integer)
    mvariEditavel = vData
End Property

Public Property Get iEditavel() As Integer
    iEditavel = mvariEditavel
End Property

Public Property Let iFormato(ByVal vData As Integer)
    mvariFormato = vData
End Property

Public Property Get iFormato() As Integer
    iFormato = mvariFormato
End Property

Public Property Let sFormulaL1(ByVal vData As String)
    mvarsFormulaL1 = vData
End Property

Public Property Get sFormulaL1() As String
    sFormulaL1 = mvarsFormulaL1
End Property

Public Property Let sFormulaGeral(ByVal vData As String)
    mvarsFormulaGeral = vData
End Property

Public Property Get sFormulaGeral() As String
    sFormulaGeral = mvarsFormulaGeral
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iLinha(ByVal vData As Integer)
    mvariLinha = vData
End Property

Public Property Get iLinha() As Integer
    iLinha = mvariLinha
End Property
