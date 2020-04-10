VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCondicaoPagtoParc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCodigo As Integer
Private mvariSeq As Integer
Private mvariTipoDataBase As Integer
Private mvariTipoIntervalo As Integer
Private mvariDias As Integer
Private mvariModificador As Integer
Private mvardPercReceb As Double

'valores calculados
Private mvardtVencimento As Date
Private mvardValor As Double

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtVencimento(ByVal vData As Date)
    mvardtVencimento = vData
End Property

Public Property Get dtVencimento() As Date
    dtVencimento = mvardtVencimento
End Property

Public Property Let dPercReceb(ByVal vData As Double)
    mvardPercReceb = vData
End Property

Public Property Get dPercReceb() As Double
    dPercReceb = mvardPercReceb
End Property

Public Property Let iModificador(ByVal vData As Integer)
    mvariModificador = vData
End Property

Public Property Get iModificador() As Integer
    iModificador = mvariModificador
End Property

Public Property Let iDias(ByVal vData As Integer)
    mvariDias = vData
End Property

Public Property Get iDias() As Integer
    iDias = mvariDias
End Property

Public Property Let iTipoIntervalo(ByVal vData As Integer)
    mvariTipoIntervalo = vData
End Property

Public Property Get iTipoIntervalo() As Integer
    iTipoIntervalo = mvariTipoIntervalo
End Property

Public Property Let iTipoDataBase(ByVal vData As Integer)
    mvariTipoDataBase = vData
End Property

Public Property Get iTipoDataBase() As Integer
    iTipoDataBase = mvariTipoDataBase
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property


