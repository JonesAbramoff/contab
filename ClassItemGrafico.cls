VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNomeColuna As String
Private mvardValorColuna As Double
Private mvarLegendText As String
Public Property Let LegendText(ByVal vData As String)
Attribute LegendText.VB_Description = "Texto que aparecerá no quadro de legenda"
    mvarLegendText = vData
End Property

Public Property Get LegendText() As String
    LegendText = mvarLegendText
End Property

Public Property Let dValorColuna(ByVal vData As Double)
    mvardValorColuna = vData
End Property

Public Property Get dValorColuna() As Double
    dValorColuna = mvardValorColuna
End Property

Public Property Let sNomeColuna(ByVal vData As String)
    mvarsNomeColuna = vData
End Property

Public Property Get sNomeColuna() As String
    sNomeColuna = mvarsNomeColuna
End Property
