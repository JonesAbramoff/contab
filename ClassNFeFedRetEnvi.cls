VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFeFedRetEnvi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarsnRec As String
Private mvariTpAmb As Integer
Private mvarsverAplic As String
Private mvarsversao As String
Private mvarscStat As String
Private mvarsxMotivo As String
Private mvarscUF As String
Private mvardtData As Date
Private mvarstMed As String
Private mvarsLote As String
Private mvardHora As Double

Public Property Let dHora(ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora() As Double
    dHora = mvardHora
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let stMed(ByVal vData As String)
    mvarstMed = vData
End Property

Public Property Get stMed() As String
    stMed = mvarstMed
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let scUF(ByVal vData As String)
    mvarscUF = vData
End Property

Public Property Get scUF() As String
    scUF = mvarscUF
End Property

Public Property Let sxMotivo(ByVal vData As String)
    mvarsxMotivo = vData
End Property

Public Property Get sxMotivo() As String
    sxMotivo = mvarsxMotivo
End Property

Public Property Let scStat(ByVal vData As String)
    mvarscStat = vData
End Property

Public Property Get scStat() As String
    scStat = mvarscStat
End Property

Public Property Let sversao(ByVal vData As String)
    mvarsversao = vData
End Property

Public Property Get sversao() As String
    sversao = mvarsversao
End Property

Public Property Let sverAplic(ByVal vData As String)
    mvarsverAplic = vData
End Property

Public Property Get sverAplic() As String
    sverAplic = mvarsverAplic
End Property

Public Property Let iTpAmb(ByVal vData As Integer)
    mvariTpAmb = vData
End Property

Public Property Get iTpAmb() As Integer
    iTpAmb = mvariTpAmb
End Property

Public Property Let snRec(ByVal vData As String)
    mvarsnRec = vData
End Property

Public Property Get snRec() As String
    snRec = mvarsnRec
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property


