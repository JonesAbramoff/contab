VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFiltroImportCoinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariGerarComissaoRetroativo As Integer
Private mvariContabilizarVouchers As Integer
Private mvariGerarOver As Integer
Private mvariAtualizarClientesRetroativo As Integer
Private mvariGerarMovEst As Integer
Private mvariImportarNVLRetroativo As Integer
Private mvariConsiderarFatSigav As Integer
Private mvariImportArqsNovos As Integer
Private mvariContabilizarFaturas As Integer
Private mvariContabilizarNFs As Integer
Private mvariSoArqExport As Integer
Private mvariGerarBol As Integer

Public Property Let iGerarBol(ByVal vData As Integer)
    mvariGerarBol = vData
End Property

Public Property Get iGerarBol() As Integer
    iGerarBol = mvariGerarBol
End Property

Public Property Let iSoArqExport(ByVal vData As Integer)
    mvariSoArqExport = vData
End Property

Public Property Get iSoArqExport() As Integer
    iSoArqExport = mvariSoArqExport
End Property

Public Property Let iContabilizarNFs(ByVal vData As Integer)
    mvariContabilizarNFs = vData
End Property

Public Property Get iContabilizarNFs() As Integer
    iContabilizarNFs = mvariContabilizarNFs
End Property

Public Property Let iContabilizarFaturas(ByVal vData As Integer)
    mvariContabilizarFaturas = vData
End Property

Public Property Get iContabilizarFaturas() As Integer
    iContabilizarFaturas = mvariContabilizarFaturas
End Property

Public Property Let iImportArqsNovos(ByVal vData As Integer)
    mvariImportArqsNovos = vData
End Property

Public Property Get iImportArqsNovos() As Integer
    iImportArqsNovos = mvariImportArqsNovos
End Property

Public Property Let iConsiderarFatSigav(ByVal vData As Integer)
    mvariConsiderarFatSigav = vData
End Property

Public Property Get iConsiderarFatSigav() As Integer
    iConsiderarFatSigav = mvariConsiderarFatSigav
End Property

Public Property Let iImportarNVLRetroativo(ByVal vData As Integer)
    mvariImportarNVLRetroativo = vData
End Property

Public Property Get iImportarNVLRetroativo() As Integer
    iImportarNVLRetroativo = mvariImportarNVLRetroativo
End Property

Public Property Let iGerarMovEst(ByVal vData As Integer)
    mvariGerarMovEst = vData
End Property

Public Property Get iGerarMovEst() As Integer
    iGerarMovEst = mvariGerarMovEst
End Property

Public Property Let iAtualizarClientesRetroativo(ByVal vData As Integer)
    mvariAtualizarClientesRetroativo = vData
End Property

Public Property Get iAtualizarClientesRetroativo() As Integer
    iAtualizarClientesRetroativo = mvariAtualizarClientesRetroativo
End Property

Public Property Let iGerarOver(ByVal vData As Integer)
    mvariGerarOver = vData
End Property

Public Property Get iGerarOver() As Integer
    iGerarOver = mvariGerarOver
End Property

Public Property Let iContabilizarVouchers(ByVal vData As Integer)
    mvariContabilizarVouchers = vData
End Property

Public Property Get iContabilizarVouchers() As Integer
    iContabilizarVouchers = mvariContabilizarVouchers
End Property

Public Property Let iGerarComissaoRetroativo(ByVal vData As Integer)
    mvariGerarComissaoRetroativo = vData
End Property

Public Property Get iGerarComissaoRetroativo() As Integer
    iGerarComissaoRetroativo = mvariGerarComissaoRetroativo
End Property
