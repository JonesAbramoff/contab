VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVOcrCasoImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntArq As Long
Private mvarsCodigo As String
Private mvardtData As Date
Private mvarsTCliente As String
Private mvarsNome As String
Private mvarsSobrenome As String
Private mvarsChaveVou As String
Private mvarsTipVou As String
Private mvarsSerie As String
Private mvarlNumVou As Long
Private mvarsCidadeOCR As String
Private mvarsEstadoOCR As String
Private mvarsPaisOCR As String
Private mvarsPrestador As String
Private mvarsCarater As String
Private mvarsGrauSatisfacao As String
Private mvarsTelefone As String
Private mvarlSeq As Long
Private mvarsNumVouTexto As String

Public Property Let sNumVouTexto(ByVal vData As String)
    mvarsNumVouTexto = vData
End Property

Public Property Get sNumVouTexto() As String
    sNumVouTexto = mvarsNumVouTexto
End Property

Public Property Let lSeq(ByVal vData As Long)
    mvarlSeq = vData
End Property

Public Property Get lSeq() As Long
    lSeq = mvarlSeq
End Property

Public Property Let lNumIntArq(ByVal vData As Long)
    mvarlNumIntArq = vData
End Property

Public Property Get lNumIntArq() As Long
    lNumIntArq = mvarlNumIntArq
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sTCliente(ByVal vData As String)
    mvarsTCliente = vData
End Property

Public Property Get sTCliente() As String
    sTCliente = mvarsTCliente
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sSobrenome(ByVal vData As String)
    mvarsSobrenome = vData
End Property

Public Property Get sSobrenome() As String
    sSobrenome = mvarsSobrenome
End Property

Public Property Let sChaveVou(ByVal vData As String)
    mvarsChaveVou = vData
End Property

Public Property Get sChaveVou() As String
    sChaveVou = mvarsChaveVou
End Property

Public Property Let sTipVou(ByVal vData As String)
    mvarsTipVou = vData
End Property

Public Property Get sTipVou() As String
    sTipVou = mvarsTipVou
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let lNumVou(ByVal vData As Long)
    mvarlNumVou = vData
End Property

Public Property Get lNumVou() As Long
    lNumVou = mvarlNumVou
End Property

Public Property Let sCidadeOCR(ByVal vData As String)
    mvarsCidadeOCR = vData
End Property

Public Property Get sCidadeOCR() As String
    sCidadeOCR = mvarsCidadeOCR
End Property

Public Property Let sEstadoOCR(ByVal vData As String)
    mvarsEstadoOCR = vData
End Property

Public Property Get sEstadoOCR() As String
    sEstadoOCR = mvarsEstadoOCR
End Property

Public Property Let sPaisOCR(ByVal vData As String)
    mvarsPaisOCR = vData
End Property

Public Property Get sPaisOCR() As String
    sPaisOCR = mvarsPaisOCR
End Property

Public Property Let sPrestador(ByVal vData As String)
    mvarsPrestador = vData
End Property

Public Property Get sPrestador() As String
    sPrestador = mvarsPrestador
End Property

Public Property Let sCarater(ByVal vData As String)
    mvarsCarater = vData
End Property

Public Property Get sCarater() As String
    sCarater = mvarsCarater
End Property

Public Property Let sGrauSatisfacao(ByVal vData As String)
    mvarsGrauSatisfacao = vData
End Property

Public Property Get sGrauSatisfacao() As String
    sGrauSatisfacao = mvarsGrauSatisfacao
End Property

Public Property Let sTelefone(ByVal vData As String)
    mvarsTelefone = vData
End Property

Public Property Get sTelefone() As String
    sTelefone = mvarsTelefone
End Property
