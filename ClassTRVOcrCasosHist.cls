VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVOcrCasosHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocOcrCaso As Long
Private mvariSeq As Integer
Private mvardtData As Date
Private mvardHora As Double
Private mvardtDataReg As Date
Private mvardHoraReg As Double
Private mvarsUsuario As String
Private mvariOrigem As Integer
Private mvarsAssunto As String

Public Property Let sAssunto(ByVal vData As String)
    mvarsAssunto = vData
End Property

Public Property Get sAssunto() As String
    sAssunto = mvarsAssunto
End Property

Public Property Let lNumIntDocOcrCaso(ByVal vData As Long)
    mvarlNumIntDocOcrCaso = vData
End Property

Public Property Get lNumIntDocOcrCaso() As Long
    lNumIntDocOcrCaso = mvarlNumIntDocOcrCaso
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dHora(ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora() As Double
    dHora = mvardHora
End Property

Public Property Let dtdatareg(ByVal vData As Date)
    mvardtDataReg = vData
End Property

Public Property Get dtdatareg() As Date
    dtdatareg = mvardtDataReg
End Property

Public Property Let dHoraReg(ByVal vData As Double)
    mvardHoraReg = vData
End Property

Public Property Get dHoraReg() As Double
    dHoraReg = mvardHoraReg
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let iOrigem(ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property
