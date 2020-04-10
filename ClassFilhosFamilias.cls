VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFilhosFamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodFamilia As Long
Private mvariSeqFilho As Integer
Private mvarsNome As String
Private mvarsNomeHebr As String
Private mvardtDtNasc As Date
Private mvariDtNascNoite As Integer
Private mvardtDtFal As Date
Private mvariDtFalNoite As Integer
Private mvarsTelefone As String
Private mvarsEmail As String

Public iAN As Integer
Public iCD As Integer
Public iCH As Integer
Public iCJ As Integer
Public iH As Integer
Public iH1 As Integer
Public iH2 As Integer
Public iLE As Integer
Public iLR As Integer
Public iPA As Integer
Public iRE As Integer
Public iSH As Integer
Public iSI As Integer
Public iTH As Integer
Public iVF As Integer

Public Property Let sTelefone(ByVal vData As String)
    mvarsTelefone = vData
End Property

Public Property Get sTelefone() As String
    sTelefone = mvarsTelefone
End Property

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Get sDtNascHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtDtNasc = DATA_NULA Then
        sDtNascHebr = ""
    Else
        sDtNascHebr = GregToHeb(mvardtDtNasc + mvariDtNascNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Get sDtFalHebr() As String
Dim iAnoHebr As Integer, iMesHebr As Integer, iDiaHebr As Integer

    If mvardtDtFal = DATA_NULA Then
        sDtFalHebr = ""
    Else
        sDtFalHebr = GregToHeb(mvardtDtFal + mvariDtFalNoite, iAnoHebr, iMesHebr, iDiaHebr)
    End If
    
End Property

Public Property Let dtDtFal(ByVal vData As Date)
    mvardtDtFal = vData
End Property

Public Property Get dtDtFal() As Date
    dtDtFal = mvardtDtFal
End Property

Public Property Let iDtFalNoite(ByVal vData As Integer)
    mvariDtFalNoite = vData
End Property

Public Property Get iDtFalNoite() As Integer
    iDtFalNoite = mvariDtFalNoite
End Property

Public Property Let lCodFamilia(ByVal vData As Long)
    mvarlCodFamilia = vData
End Property

Public Property Get lCodFamilia() As Long
    lCodFamilia = mvarlCodFamilia
End Property

Public Property Let iSeqFilho(ByVal vData As Integer)
    mvariSeqFilho = vData
End Property

Public Property Get iSeqFilho() As Integer
    iSeqFilho = mvariSeqFilho
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sNomeHebr(ByVal vData As String)
    mvarsNomeHebr = vData
End Property

Public Property Get sNomeHebr() As String
    sNomeHebr = mvarsNomeHebr
End Property

Public Property Let dtDtNasc(ByVal vData As Date)
    mvardtDtNasc = vData
End Property

Public Property Get dtDtNasc() As Date
    dtDtNasc = mvardtDtNasc
End Property

Public Property Let iDtNascNoite(ByVal vData As Integer)
    mvariDtNascNoite = vData
End Property

Public Property Get iDtNascNoite() As Integer
    iDtNascNoite = mvariDtNascNoite
End Property
