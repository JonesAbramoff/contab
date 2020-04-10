VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFeFedScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlOcorrencia As Long
Private mvardtDataEntrada As Date
Private mvardHoraEntrada As Double
Private mvardtDataSaida As Date
Private mvardHoraSaida As Double
Private mvarsJustificativa As String
Private mvariFilialEmpresa As Integer

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sJustificativa(ByVal vData As String)
    mvarsJustificativa = vData
End Property

Public Property Get sJustificativa() As String
    sJustificativa = mvarsJustificativa
End Property

Public Property Let dHoraSaida(ByVal vData As Double)
    mvardHoraSaida = vData
End Property

Public Property Get dHoraSaida() As Double
    dHoraSaida = mvardHoraSaida
End Property

Public Property Let dtDataSaida(ByVal vData As Date)
    mvardtDataSaida = vData
End Property

Public Property Get dtDataSaida() As Date
    dtDataSaida = mvardtDataSaida
End Property

Public Property Let dHoraEntrada(ByVal vData As Double)
    mvardHoraEntrada = vData
End Property

Public Property Get dHoraEntrada() As Double
    dHoraEntrada = mvardHoraEntrada
End Property

Public Property Let dtDataEntrada(ByVal vData As Date)
    mvardtDataEntrada = vData
End Property

Public Property Get dtDataEntrada() As Date
    dtDataEntrada = mvardtDataEntrada
End Property

Public Property Let lOcorrencia(ByVal vData As Long)
    mvarlOcorrencia = vData
End Property

Public Property Get lOcorrencia() As Long
    lOcorrencia = mvarlOcorrencia
End Property


