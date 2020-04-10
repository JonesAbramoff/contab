VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoKitCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarsGrupoPesagem As String
Private mvariTipoCarga As Integer

Public Property Let iTipoCarga(ByVal vData As Integer)
    mvariTipoCarga = vData
End Property

Public Property Get iTipoCarga() As Integer
    iTipoCarga = mvariTipoCarga
End Property

Public Property Let sGrupoPesagem(ByVal vData As String)
    mvarsGrupoPesagem = vData
End Property

Public Property Get sGrupoPesagem() As String
    sGrupoPesagem = mvarsGrupoPesagem
End Property



