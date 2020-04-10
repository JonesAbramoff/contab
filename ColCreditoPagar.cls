VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColCreditoPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassCreditoPagar"
Attribute VB_Ext_KEY = "Member0" ,"ClassCreditoPagar"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(dtDataEmissao As Date, iFilial As Integer, lNumIntDoc As Long, lFornecedor As Long, iStatus As Integer, sObservacao As String, lNumTitulo As Long, iFilialEmpresa As Integer, dValorTotal As Double, dSaldo As Double, sSiglaDocumento As String, dValorSeguro As Double, dValorFrete As Double, dOutrasDespesas As Double, dValorProdutos As Double, dValorICMS As Double, dValorICMSSubst As Double, dValorIPI As Double, dValorIRRF As Double, iDebitoICMS As Integer, iDebitoIPI As Integer, dValorBaixado As Double, Optional sKey As String) As ClassCreditoPagar
    
Dim objNewMember As ClassCreditoPagar
Set objNewMember = New ClassCreditoPagar

    objNewMember.dtDataEmissao = dtDataEmissao
    objNewMember.iFilial = iFilial
    objNewMember.lNumIntDoc = lNumIntDoc
    objNewMember.lFornecedor = lFornecedor
    objNewMember.iStatus = iStatus
    objNewMember.sObservacao = sObservacao
    objNewMember.lNumTitulo = lNumTitulo
    objNewMember.iFilialEmpresa = iFilialEmpresa
    objNewMember.dValorTotal = dValorTotal
    objNewMember.dSaldo = dSaldo
    objNewMember.sSiglaDocumento = sSiglaDocumento
    objNewMember.dValorSeguro = dValorSeguro
    objNewMember.dValorFrete = dValorFrete
    objNewMember.dOutrasDespesas = dOutrasDespesas
    objNewMember.dValorProdutos = dValorProdutos
    objNewMember.dValorICMS = dValorICMS
    objNewMember.dValorICMSSubst = dValorICMSSubst
    objNewMember.dValorIPI = dValorIPI
    objNewMember.dValorIRRF = dValorIRRF
    objNewMember.iDebitoICMS = iDebitoICMS
    objNewMember.iDebitoIPI = iDebitoIPI
    objNewMember.dValorBaixado = dValorBaixado

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassCreditoPagar
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

