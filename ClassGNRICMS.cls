VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGNRICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvardtDataPagto As Date
Private mvarsCGCSubstTrib As String
Private mvarsInscricaoEstadual As String
Private mvarsUFSubstTrib As String
Private mvarsUFDestino As String
Private mvariBanco As Integer
Private mvariAgencia As Integer
Private mvarsNumero As String
Private mvardValor As Double
Private mvardtVencimento As Date
Private mvardtDataRef As Date
Private mvarsConvenio As String
Private mvarlNumIntRegApuracaoICMS As Long
Private mvarlNumIntArqICMS As Long
Private mvariTipo As Integer
Private mvarlCodigo As Long
'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Private mvarsCodReceita As String 'local copy
Private mvarsCodObrigRecolher As String 'local copy

Public Property Let sCodObrigRecolher(ByVal vData As String)
    mvarsCodObrigRecolher = vData
End Property

Public Property Get sCodObrigRecolher() As String
    sCodObrigRecolher = mvarsCodObrigRecolher
End Property

Public Property Let sCodReceita(ByVal vData As String)
    mvarsCodReceita = vData
End Property

Public Property Get sCodReceita() As String
    sCodReceita = mvarsCodReceita
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property




Public Property Let lCodigo(ByVal vData As Long)
Attribute lCodigo.VB_Description = "Codigo Automatico da Guia"
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iTipo(ByVal vData As Integer)
Attribute iTipo.VB_Description = "1 - Arquivo ICMS, 2 -  Apuração ICMS, ???? 3 - Apuração IPI"
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lNumIntArqICMS(ByVal vData As Long)
Attribute lNumIntArqICMS.VB_Description = "identifica o ""arquivo"" gerado com info s/ICMS cnvenio 57/95"
    mvarlNumIntArqICMS = vData
End Property

Public Property Get lNumIntArqICMS() As Long
    lNumIntArqICMS = mvarlNumIntArqICMS
End Property

Public Property Let lNumIntRegApuracaoICMS(ByVal vData As Long)
Attribute lNumIntRegApuracaoICMS.VB_Description = "identifica a apuracao a que se refere este recolhimento"
    mvarlNumIntRegApuracaoICMS = vData
End Property

Public Property Get lNumIntRegApuracaoICMS() As Long
    lNumIntRegApuracaoICMS = mvarlNumIntRegApuracaoICMS
End Property

Public Property Let sConvenio(ByVal vData As String)
    mvarsConvenio = vData
End Property

Public Property Get sConvenio() As String
    sConvenio = mvarsConvenio
End Property

Public Property Let dtDataRef(ByVal vData As Date)
    mvardtDataRef = vData
End Property

Public Property Get dtDataRef() As Date
    dtDataRef = mvardtDataRef
End Property

Public Property Let dtVencimento(ByVal vData As Date)
    mvardtVencimento = vData
End Property

Public Property Get dtVencimento() As Date
    dtVencimento = mvardtVencimento
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let sNumero(ByVal vData As String)
    mvarsNumero = vData
End Property

Public Property Get sNumero() As String
    sNumero = mvarsNumero
End Property

Public Property Let iAgencia(ByVal vData As Integer)
    mvariAgencia = vData
End Property

Public Property Get iAgencia() As Integer
    iAgencia = mvariAgencia
End Property

Public Property Let iBanco(ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco() As Integer
    iBanco = mvariBanco
End Property

Public Property Let sUFDestino(ByVal vData As String)
    mvarsUFDestino = vData
End Property

Public Property Get sUFDestino() As String
    sUFDestino = mvarsUFDestino
End Property

Public Property Let sUFSubstTrib(ByVal vData As String)
    mvarsUFSubstTrib = vData
End Property

Public Property Get sUFSubstTrib() As String
    sUFSubstTrib = mvarsUFSubstTrib
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let sCGCSubstTrib(ByVal vData As String)
    mvarsCGCSubstTrib = vData
End Property

Public Property Get sCGCSubstTrib() As String
    sCGCSubstTrib = mvarsCGCSubstTrib
End Property

Public Property Let dtDataPagto(ByVal vData As Date)
Attribute dtDataPagto.VB_Description = "Data do Pagamento do Documento de Arrecadação"
    mvardtDataPagto = vData
End Property

Public Property Get dtDataPagto() As Date
    dtDataPagto = mvardtDataPagto
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
