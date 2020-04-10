VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAntecipReceb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCliente As Long
Private mvariFilial As Integer
Private mvariFilialEmpresa As Integer
Private mvariCodConta As Integer
Private mvarlSequencial As Long
Private mvardtData As Date
Private mvardValor As Double
Private mvariTipoMeioPagto As Integer
Private mvarsHistorico As String
Private mvarlNumMovto As Long
Private mvarsNumRefExterna As String
Private mvarlNumIntRec As Long
Private mvardSaldoNaoApropriado As Double
Private mvariExcluido As Integer
Private mvardValorBaixado As Double
Private mvarsNatureza As String

Public Property Let sNatureza(ByVal vData As String)
    mvarsNatureza = vData
End Property

Public Property Get sNatureza() As String
    sNatureza = mvarsNatureza
End Property

Public Property Let dValorBaixado(ByVal vData As Double)
    mvardValorBaixado = vData
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Let iExcluido(ByVal vData As Integer)
Attribute iExcluido.VB_Description = "0=Válido 1=Excluido"
    mvariExcluido = vData
End Property

Public Property Get iExcluido() As Integer
    iExcluido = mvariExcluido
End Property

Public Property Let dSaldoNaoApropriado(ByVal vData As Double)
Attribute dSaldoNaoApropriado.VB_Description = "valor ainda nao utilizado c/abatimento"
    mvardSaldoNaoApropriado = vData
End Property

Public Property Get dSaldoNaoApropriado() As Double
    dSaldoNaoApropriado = mvardSaldoNaoApropriado
End Property

Public Property Let lNumIntRec(ByVal vData As Long)
Attribute lNumIntRec.VB_Description = "Identificador do movimento"
    mvarlNumIntRec = vData
End Property

Public Property Get lNumIntRec() As Long
    lNumIntRec = mvarlNumIntRec
End Property

Public Property Let sNumRefExterna(ByVal vData As String)
    mvarsNumRefExterna = vData
End Property

Public Property Get sNumRefExterna() As String
    sNumRefExterna = mvarsNumRefExterna
End Property

Public Property Let lNumMovto(ByVal vData As Long)
    mvarlNumMovto = vData
End Property

Public Property Get lNumMovto() As Long
    lNumMovto = mvarlNumMovto
End Property

Public Property Let sHistorico(ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lSequencial(ByVal vData As Long)
    mvarlSequencial = vData
End Property

Public Property Get lSequencial() As Long
    lSequencial = mvarlSequencial
End Property

Public Property Let iCodConta(ByVal vData As Integer)
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

