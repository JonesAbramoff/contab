VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntBaixa As Variant
Private mvariStatus As Integer
Private mvarsHistorico As String
Private mvariMotivo As Integer
Private mvardtData As Date
Private mvardtDataContabil As Date
Private mvardtDataRegistro As Date
Private mvarlNumMovCta As Long
Private mvarlNumIntDoc As Long
Private mvarcolBaixaParcRec As colBaixaParcRec
Private mvarlNumIntCheque As Long
Private mvariFilialEmpresaCheque As Integer
Private mvarlNumIntDocBaixaAgrupada As Long

Public Property Let lNumIntDocBaixaAgrupada(ByVal vData As Long)
    mvarlNumIntDocBaixaAgrupada = vData
End Property

Public Property Get lNumIntDocBaixaAgrupada() As Long
    lNumIntDocBaixaAgrupada = mvarlNumIntDocBaixaAgrupada
End Property

Public Property Let iFilialEmpresaCheque(ByVal vData As Integer)
    mvariFilialEmpresaCheque = vData
End Property

Public Property Get iFilialEmpresaCheque() As Integer
    iFilialEmpresaCheque = mvariFilialEmpresaCheque
End Property

Public Property Let lNumIntCheque(ByVal vData As Long)
    mvarlNumIntCheque = vData
End Property

Public Property Get lNumIntCheque() As Long
    lNumIntCheque = mvarlNumIntCheque
End Property

Public Property Get colBaixaParcRec() As colBaixaParcRec
    Set colBaixaParcRec = mvarcolBaixaParcRec
End Property

Public Property Set colBaixaParcRec(ByVal vData As colBaixaParcRec)
    Set mvarcolBaixaParcRec = vData
End Property

Public Property Let dtDataContabil(ByVal vData As Date)
    mvardtDataContabil = vData
End Property

Public Property Get dtDataContabil() As Date
    dtDataContabil = mvardtDataContabil
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
Attribute lNumIntDoc.VB_Description = "identifica reg em DebitosRecCli"
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumMovCta(ByVal vData As Long)
Attribute lNumMovCta.VB_Description = "numero que identifica um movto de cta corrente"
    mvarlNumMovCta = vData
End Property

Public Property Get lNumMovCta() As Long
    lNumMovCta = mvarlNumMovCta
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iMotivo(ByVal vData As Integer)
Attribute iMotivo.VB_Description = "normal, devolucao, dacao"
    mvariMotivo = vData
End Property

Public Property Get iMotivo() As Integer
    iMotivo = mvariMotivo
End Property

Public Property Let sHistorico(ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let iStatus(ByVal vData As Integer)
Attribute iStatus.VB_Description = "1:lancado, 2:excluido"
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lNumIntBaixa(ByVal vData As Variant)
    mvarlNumIntBaixa = vData
End Property

Public Property Set lNumIntBaixa(ByVal vData As Object)
    Set mvarlNumIntBaixa = vData
End Property

Public Property Get lNumIntBaixa() As Variant
    If IsObject(mvarlNumIntBaixa) Then
        Set lNumIntBaixa = mvarlNumIntBaixa
    Else
        lNumIntBaixa = mvarlNumIntBaixa
    End If
End Property



