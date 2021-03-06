VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoSegmento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Cont�m o tipo do segmento (formato n�merico) tal qual � armazenado no banco de dados e sua descri��o, que pode ser exibida em tela."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsDescricao As String
Private mvariTipo As Integer

Public Property Let iTipo(ByVal vData As Integer)
Attribute iTipo.VB_Description = "Tipo do segmento. Formato com que o tipo � armazenado no BD. "
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descri��o do Tipo do Segmento. Ex: Alfanum�rico, Num�rico, etc."
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property



