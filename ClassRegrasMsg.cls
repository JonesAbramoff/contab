VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRegrasMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariTipoDoc As Integer
Private mvariSeq As Integer
Private mvariTipoMsg As Integer
Private mvarsRegra1 As String
Private mvarsRegra2 As String
Private mvarsRegra3 As String
Private mvarsRegra4 As String
Private mvarsRegra5 As String
Private mvarsMensagem As String
Private mvarsMensagemRet As String
Private mvardRegraRet As Double

Public Property Let dRegraRet(ByVal vData As Double)
    mvardRegraRet = vData
End Property

Public Property Get dRegraRet() As Double
    dRegraRet = mvardRegraRet
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iTipoMsg(ByVal vData As Integer)
    mvariTipoMsg = vData
End Property

Public Property Get iTipoMsg() As Integer
    iTipoMsg = mvariTipoMsg
End Property

Public Property Let sRegra1(ByVal vData As String)
    mvarsRegra1 = vData
End Property

Public Property Get sRegra1() As String
    sRegra1 = mvarsRegra1
End Property

Public Property Let sRegra2(ByVal vData As String)
    mvarsRegra2 = vData
End Property

Public Property Get sRegra2() As String
    sRegra2 = mvarsRegra2
End Property

Public Property Let sRegra3(ByVal vData As String)
    mvarsRegra3 = vData
End Property

Public Property Get sRegra3() As String
    sRegra3 = mvarsRegra3
End Property

Public Property Let sRegra4(ByVal vData As String)
    mvarsRegra4 = vData
End Property

Public Property Get sRegra4() As String
    sRegra4 = mvarsRegra4
End Property

Public Property Let sRegra5(ByVal vData As String)
    mvarsRegra5 = vData
End Property

Public Property Get sRegra5() As String
    sRegra5 = mvarsRegra5
End Property

Public Property Let sMensagem(ByVal vData As String)
    mvarsMensagem = vData
End Property

Public Property Get sMensagem() As String
    sMensagem = mvarsMensagem
End Property

Public Property Let sMensagemRet(ByVal vData As String)
    mvarsMensagemRet = vData
End Property

Public Property Get sMensagemRet() As String
    sMensagemRet = mvarsMensagemRet
End Property
