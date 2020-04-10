VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColNFsPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassNFsPag"
Attribute VB_Ext_KEY = "Member0" ,"ClassNFsPag"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(lNumIntDoc As Long, iFilialEmpresa As Integer, lFornecedor As Long, iFilial As Integer, lNumNotaFiscal As Long, dtDataEmissao As Date, iStatus As Integer, lNumIntTitPag As Long, dtDataVencimento As Date, dValorTotal As Double, dValorSeguro As Double, dValorFrete As Double, dOutrasDespesas As Double, dValorProdutos As Double, dValorIRRF As Double, dValorICMS As Double, dValorICMSSubst As Double, iCreditoICMS As Integer, dValorIPI As Double, iCreditoIPI As Integer, dPISRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double, Optional sKey As String) As ClassNFsPag
    
Dim objNewMember As ClassNFsPag
Set objNewMember = New ClassNFsPag

    objNewMember.lNumIntDoc = lNumIntDoc
    objNewMember.iFilialEmpresa = iFilialEmpresa
    objNewMember.lFornecedor = lFornecedor
    objNewMember.iFilial = iFilial
    objNewMember.lNumNotaFiscal = lNumNotaFiscal
    objNewMember.dtDataEmissao = dtDataEmissao
    objNewMember.iStatus = iStatus
    objNewMember.lNumIntTitPag = lNumIntTitPag
    objNewMember.dtDataVencimento = dtDataVencimento
    objNewMember.dValorTotal = dValorTotal
    objNewMember.dValorSeguro = dValorSeguro
    objNewMember.dValorFrete = dValorFrete
    objNewMember.dOutrasDespesas = dOutrasDespesas
    objNewMember.dValorProdutos = dValorProdutos
    objNewMember.dValorIRRF = dValorIRRF
    objNewMember.dValorICMS = dValorICMS
    objNewMember.dValorICMSSubst = dValorICMSSubst
    objNewMember.iCreditoICMS = iCreditoICMS
    objNewMember.dValorIPI = dValorIPI
    objNewMember.iCreditoIPI = iCreditoIPI
    objNewMember.dPISRetido = dPISRetido
    objNewMember.dCOFINSRetido = dCOFINSRetido
    objNewMember.dCSLLRetido = dCSLLRetido
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassNFsPag
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

