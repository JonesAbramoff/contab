VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassRetPagtoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocRet AS Long
Private mvariSeq AS Integer
Private mvariLote AS Integer
Private mvariSeqLote AS Integer
Private mvariTipoMov AS Integer
Private mvariCodInstMov AS Integer
Private mvariBanco AS Integer
Private mvarsAgencia AS String
Private mvarsConta AS String
Private mvarsNomeFavorecido AS String
Private mvarsSeuNumero AS String
Private mvarsNossoNumero AS String
Private mvardtDataPagto AS Date
Private mvardValorPagto AS Double
Private mvardtDataReal AS Date
Private mvardValorReal AS Double
Private mvarsFinalidade AS String
Private mvarsCodOCR1 AS String
Private mvarsCodOCR2 AS String
Private mvarsCodOCR3 AS String
Private mvarsCodOCR4 AS String
Private mvarsCodOCR5 AS String
Private mvariTipo AS Integer
Private mvarsCodigoBarras AS String

Public Property Let lNumIntDocRet (ByVal vData As Long)
    mvarlNumIntDocRet = vData
End Property

Public Property Get lNumIntDocRet () AS Long
    lNumIntDocRet= mvarlNumIntDocRet
End Property

Public Property Let iSeq (ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq () AS Integer
    iSeq= mvariSeq
End Property

Public Property Let iLote (ByVal vData As Integer)
    mvariLote = vData
End Property

Public Property Get iLote () AS Integer
    iLote= mvariLote
End Property

Public Property Let iSeqLote (ByVal vData As Integer)
    mvariSeqLote = vData
End Property

Public Property Get iSeqLote () AS Integer
    iSeqLote= mvariSeqLote
End Property

Public Property Let iTipoMov (ByVal vData As Integer)
    mvariTipoMov = vData
End Property

Public Property Get iTipoMov () AS Integer
    iTipoMov= mvariTipoMov
End Property

Public Property Let iCodInstMov (ByVal vData As Integer)
    mvariCodInstMov = vData
End Property

Public Property Get iCodInstMov () AS Integer
    iCodInstMov= mvariCodInstMov
End Property

Public Property Let iBanco (ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco () AS Integer
    iBanco= mvariBanco
End Property

Public Property Let sAgencia (ByVal vData As String)
    mvarsAgencia = vData
End Property

Public Property Get sAgencia () AS String
    sAgencia= mvarsAgencia
End Property

Public Property Let sConta (ByVal vData As String)
    mvarsConta = vData
End Property

Public Property Get sConta () AS String
    sConta= mvarsConta
End Property

Public Property Let sNomeFavorecido (ByVal vData As String)
    mvarsNomeFavorecido = vData
End Property

Public Property Get sNomeFavorecido () AS String
    sNomeFavorecido= mvarsNomeFavorecido
End Property

Public Property Let sSeuNumero (ByVal vData As String)
    mvarsSeuNumero = vData
End Property

Public Property Get sSeuNumero () AS String
    sSeuNumero= mvarsSeuNumero
End Property

Public Property Let sNossoNumero (ByVal vData As String)
    mvarsNossoNumero = vData
End Property

Public Property Get sNossoNumero () AS String
    sNossoNumero= mvarsNossoNumero
End Property

Public Property Let dtDataPagto (ByVal vData As Date)
    mvardtDataPagto = vData
End Property

Public Property Get dtDataPagto () AS Date
    dtDataPagto= mvardtDataPagto
End Property

Public Property Let dValorPagto (ByVal vData As Double)
    mvardValorPagto = vData
End Property

Public Property Get dValorPagto () AS Double
    dValorPagto= mvardValorPagto
End Property

Public Property Let dtDataReal (ByVal vData As Date)
    mvardtDataReal = vData
End Property

Public Property Get dtDataReal () AS Date
    dtDataReal= mvardtDataReal
End Property

Public Property Let dValorReal (ByVal vData As Double)
    mvardValorReal = vData
End Property

Public Property Get dValorReal () AS Double
    dValorReal= mvardValorReal
End Property

Public Property Let sFinalidade (ByVal vData As String)
    mvarsFinalidade = vData
End Property

Public Property Get sFinalidade () AS String
    sFinalidade= mvarsFinalidade
End Property

Public Property Let sCodOCR1 (ByVal vData As String)
    mvarsCodOCR1 = vData
End Property

Public Property Get sCodOCR1 () AS String
    sCodOCR1= mvarsCodOCR1
End Property

Public Property Let sCodOCR2 (ByVal vData As String)
    mvarsCodOCR2 = vData
End Property

Public Property Get sCodOCR2 () AS String
    sCodOCR2= mvarsCodOCR2
End Property

Public Property Let sCodOCR3 (ByVal vData As String)
    mvarsCodOCR3 = vData
End Property

Public Property Get sCodOCR3 () AS String
    sCodOCR3= mvarsCodOCR3
End Property

Public Property Let sCodOCR4 (ByVal vData As String)
    mvarsCodOCR4 = vData
End Property

Public Property Get sCodOCR4 () AS String
    sCodOCR4= mvarsCodOCR4
End Property

Public Property Let sCodOCR5 (ByVal vData As String)
    mvarsCodOCR5 = vData
End Property

Public Property Get sCodOCR5 () AS String
    sCodOCR5= mvarsCodOCR5
End Property

Public Property Let iTipo (ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo () AS Integer
    iTipo= mvariTipo
End Property

Public Property Let sCodigoBarras (ByVal vData As String)
    mvarsCodigoBarras = vData
End Property

Public Property Get sCodigoBarras () AS String
    sCodigoBarras= mvarsCodigoBarras
End Property
