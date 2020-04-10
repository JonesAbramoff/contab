VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvariApuracaoPeriodicidade As Integer

Public Property Let iApuracaoPeriodicidade(ByVal vData As Integer)
Attribute iApuracaoPeriodicidade.VB_Description = "default para configuracao de filial (Não está sendo usado)"
    mvariApuracaoPeriodicidade = vData
End Property

Public Property Get iApuracaoPeriodicidade() As Integer
    iApuracaoPeriodicidade = mvariApuracaoPeriodicidade
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descrição do Tributo"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iCodigo(ByVal vData As Integer)
Attribute iCodigo.VB_Description = "Código do Tributo"
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property
