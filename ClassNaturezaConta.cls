VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNaturezaConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariNatureza As Integer
Private mvarsDescricao As String
Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property


Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property



Public Property Let iNatureza(ByVal vData As Integer)
    mvariNatureza = vData
End Property


Public Property Get iNatureza() As Integer
    iNatureza = mvariNatureza
End Property


