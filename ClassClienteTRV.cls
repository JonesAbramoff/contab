VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassClienteTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolCliEmissoresTRV As New Collection
Private mvarlCodRepresentante As Long
Private mvardPercComiRep As Double
Private mvarlCodCorrentista As Long
Private mvardPercComiCorr As Double
Private mvardPercComiAg As Double
Private mvariConsiderarAporte As Integer
Private mvarlEmpresaPai As Long
Private mvariFilialNF As Integer
Private mvariFilialEmpresa As Integer
Private mvariFilialCoinfo As Integer
Private mvarsGrupo As String
Private mvariFilialFat As Integer
Private mvariCondPagtoCC As Integer
Private mvarcolExcRep As New Collection
Private mvarcolExcCor As New Collection
Private mvarcolExcComissaoCli As New Collection
Private mvarcolDataCallCenter As New Collection
Private mvardPercFatorDevCMCC As Double

Public Property Let dPercFatorDevCMCC(ByVal vData As Double)
    mvardPercFatorDevCMCC = vData
End Property

Public Property Get dPercFatorDevCMCC() As Double
    dPercFatorDevCMCC = mvardPercFatorDevCMCC
End Property

Public Property Set colDataCallCenter(ByVal vData As Collection)
    Set mvarcolDataCallCenter = vData
End Property

Public Property Get colDataCallCenter() As Collection
    Set colDataCallCenter = mvarcolDataCallCenter
End Property

Public Property Set colExcComissaoCli(ByVal vData As Collection)
    Set mvarcolExcComissaoCli = vData
End Property

Public Property Get colExcComissaoCli() As Collection
    Set colExcComissaoCli = mvarcolExcComissaoCli
End Property

Public Property Set colExcCor(ByVal vData As Collection)
    Set mvarcolExcCor = vData
End Property

Public Property Get colExcCor() As Collection
    Set colExcCor = mvarcolExcCor
End Property

Public Property Set colExcRep(ByVal vData As Collection)
    Set mvarcolExcRep = vData
End Property

Public Property Get colExcRep() As Collection
    Set colExcRep = mvarcolExcRep
End Property

Public Property Let iCondPagtoCC(ByVal vData As Integer)
    mvariCondPagtoCC = vData
End Property

Public Property Get iCondPagtoCC() As Integer
    iCondPagtoCC = mvariCondPagtoCC
End Property

Public Property Let iFilialFat(ByVal vData As Integer)
    mvariFilialFat = vData
End Property

Public Property Get iFilialFat() As Integer
    iFilialFat = mvariFilialFat
End Property

Public Property Let sGrupo(ByVal vData As String)
    mvarsGrupo = vData
End Property

Public Property Get sGrupo() As String
    sGrupo = mvarsGrupo
End Property

Public Property Let iFilialCoinfo(ByVal vData As Integer)
    mvariFilialCoinfo = vData
End Property

Public Property Get iFilialCoinfo() As Integer
    iFilialCoinfo = mvariFilialCoinfo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iFilialNF(ByVal vData As Integer)
    mvariFilialNF = vData
End Property

Public Property Get iFilialNF() As Integer
    iFilialNF = mvariFilialNF
End Property

Public Property Let lEmpresaPai(ByVal vData As Long)
    mvarlEmpresaPai = vData
End Property

Public Property Get lEmpresaPai() As Long
    lEmpresaPai = mvarlEmpresaPai
End Property

Public Property Let iConsiderarAporte(ByVal vData As Integer)
    mvariConsiderarAporte = vData
End Property

Public Property Get iConsiderarAporte() As Integer
    iConsiderarAporte = mvariConsiderarAporte
End Property

Public Property Let dPercComiAg(ByVal vData As Double)
    mvardPercComiAg = vData
End Property

Public Property Get dPercComiAg() As Double
    dPercComiAg = mvardPercComiAg
End Property

Public Property Let dPercComiCorr(ByVal vData As Double)
    mvardPercComiCorr = vData
End Property

Public Property Get dPercComiCorr() As Double
    dPercComiCorr = mvardPercComiCorr
End Property

Public Property Let lCodCorrentista(ByVal vData As Long)
    mvarlCodCorrentista = vData
End Property

Public Property Get lCodCorrentista() As Long
    lCodCorrentista = mvarlCodCorrentista
End Property

Public Property Let dPercComiRep(ByVal vData As Double)
    mvardPercComiRep = vData
End Property

Public Property Get dPercComiRep() As Double
    dPercComiRep = mvardPercComiRep
End Property

Public Property Let lCodRepresentante(ByVal vData As Long)
    mvarlCodRepresentante = vData
End Property

Public Property Get lCodRepresentante() As Long
    lCodRepresentante = mvarlCodRepresentante
End Property

Public Property Set colCliEmissoresTRV(ByVal vData As Collection)
    Set mvarcolCliEmissoresTRV = vData
End Property

Public Property Get colCliEmissoresTRV() As Collection
    Set colCliEmissoresTRV = mvarcolCliEmissoresTRV
End Property



