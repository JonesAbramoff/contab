VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRegApuracaoPISCofins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvariTipo As Integer
Private mvariAno As Integer
Private mvariMes As Integer
Private mvariOrigCred As Integer
Private mvarsCNPJCedCred As String
Private mvarsCodCred As String
Private mvardVlCredApu As Double
Private mvardVlCredExtApu As Double
Private mvardVlTotCredApu As Double
Private mvardVlCredDescPAAnt As Double
Private mvardVlCredPerPAAnt As Double
Private mvardVlCredDCompPAAnt As Double
Private mvardSdCredDispEFD As Double
Private mvardVlCredDescEFD As Double
Private mvardVlCredPerEFD As Double
Private mvardVlCredDCompEFD As Double
Private mvardVlCredTrans As Double
Private mvardVlCredOut As Double
Private mvardSdCredFim As Double

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iMes(ByVal vData As Integer)
    mvariMes = vData
End Property

Public Property Get iMes() As Integer
    iMes = mvariMes
End Property

Public Property Let iOrigCred(ByVal vData As Integer)
    mvariOrigCred = vData
End Property

Public Property Get iOrigCred() As Integer
    iOrigCred = mvariOrigCred
End Property

Public Property Let sCNPJCedCred(ByVal vData As String)
    mvarsCNPJCedCred = vData
End Property

Public Property Get sCNPJCedCred() As String
    sCNPJCedCred = mvarsCNPJCedCred
End Property

Public Property Let sCodCred(ByVal vData As String)
    mvarsCodCred = vData
End Property

Public Property Get sCodCred() As String
    sCodCred = mvarsCodCred
End Property

Public Property Let dVlCredApu(ByVal vData As Double)
    mvardVlCredApu = vData
End Property

Public Property Get dVlCredApu() As Double
    dVlCredApu = mvardVlCredApu
End Property

Public Property Let dVlCredExtApu(ByVal vData As Double)
    mvardVlCredExtApu = vData
End Property

Public Property Get dVlCredExtApu() As Double
    dVlCredExtApu = mvardVlCredExtApu
End Property

Public Property Let dVlTotCredApu(ByVal vData As Double)
    mvardVlTotCredApu = vData
End Property

Public Property Get dVlTotCredApu() As Double
    dVlTotCredApu = mvardVlTotCredApu
End Property

Public Property Let dVlCredDescPAAnt(ByVal vData As Double)
    mvardVlCredDescPAAnt = vData
End Property

Public Property Get dVlCredDescPAAnt() As Double
    dVlCredDescPAAnt = mvardVlCredDescPAAnt
End Property

Public Property Let dVlCredPerPAAnt(ByVal vData As Double)
    mvardVlCredPerPAAnt = vData
End Property

Public Property Get dVlCredPerPAAnt() As Double
    dVlCredPerPAAnt = mvardVlCredPerPAAnt
End Property

Public Property Let dVlCredDCompPAAnt(ByVal vData As Double)
    mvardVlCredDCompPAAnt = vData
End Property

Public Property Get dVlCredDCompPAAnt() As Double
    dVlCredDCompPAAnt = mvardVlCredDCompPAAnt
End Property

Public Property Let dSdCredDispEFD(ByVal vData As Double)
    mvardSdCredDispEFD = vData
End Property

Public Property Get dSdCredDispEFD() As Double
    dSdCredDispEFD = mvardSdCredDispEFD
End Property

Public Property Let dVlCredDescEFD(ByVal vData As Double)
    mvardVlCredDescEFD = vData
End Property

Public Property Get dVlCredDescEFD() As Double
    dVlCredDescEFD = mvardVlCredDescEFD
End Property

Public Property Let dVlCredPerEFD(ByVal vData As Double)
    mvardVlCredPerEFD = vData
End Property

Public Property Get dVlCredPerEFD() As Double
    dVlCredPerEFD = mvardVlCredPerEFD
End Property

Public Property Let dVlCredDCompEFD(ByVal vData As Double)
    mvardVlCredDCompEFD = vData
End Property

Public Property Get dVlCredDCompEFD() As Double
    dVlCredDCompEFD = mvardVlCredDCompEFD
End Property

Public Property Let dVlCredTrans(ByVal vData As Double)
    mvardVlCredTrans = vData
End Property

Public Property Get dVlCredTrans() As Double
    dVlCredTrans = mvardVlCredTrans
End Property

Public Property Let dVlCredOut(ByVal vData As Double)
    mvardVlCredOut = vData
End Property

Public Property Get dVlCredOut() As Double
    dVlCredOut = mvardVlCredOut
End Property

Public Property Let dSdCredFim(ByVal vData As Double)
    mvardSdCredFim = vData
End Property

Public Property Get dSdCredFim() As Double
    dSdCredFim = mvardSdCredFim
End Property
