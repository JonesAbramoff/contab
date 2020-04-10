VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassRateioOffBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariLote As Integer
Private mvardtData As Date
Private mvariPeriodoInicial As Integer
Private mvariPeriodoFinal As Integer
Private mvarcolRateios As Collection
Private mvariFilialEmpresa As Integer
Private mvarsHistorico As String

Public Property Let sHistorico(ByVal vData As String)
Attribute sHistorico.VB_Description = "Historico a ser usado nos lançamentos."
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Set colRateios(ByVal vData As Object)
Attribute colRateios.VB_Description = "Contém os codigos dos rateios que serão processados."
    Set mvarcolRateios = vData
End Property

Public Property Get colRateios() As Collection
    Set colRateios = mvarcolRateios
End Property

Public Property Let iPeriodoInicial(ByVal vData As Integer)
    mvariPeriodoInicial = vData
End Property

Public Property Get iPeriodoInicial() As Integer
    iPeriodoInicial = mvariPeriodoInicial
End Property

Public Property Let iPeriodoFinal(ByVal vData As Integer)
    mvariPeriodoFinal = vData
End Property

Public Property Get iPeriodoFinal() As Integer
    iPeriodoFinal = mvariPeriodoFinal
End Property

Public Property Let dtData(ByVal vData As Date)
Attribute dtData.VB_Description = "Data Contábil. Será utilizada nos lançamentos que serão gerados."
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iLote(ByVal vData As Integer)
Attribute iLote.VB_Description = "Lote Contábil"
    mvariLote = vData
End Property

Public Property Get iLote() As Integer
    iLote = mvariLote
End Property



