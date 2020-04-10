VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColTituloPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassTituloPagar"
Attribute VB_Ext_KEY = "Member0" ,"ClassTituloPagar"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'
'Private mCol As Collection
'
'Public Function Add(dtDataEmissao As Date, iFilial As Integer, lNumIntDoc As Long, dOutrasDespesas As Double, dValorFrete As Double, dValorProdutos As Double, dValorSeguro As Double, dValorTotal As Double, lFornecedor As Long, iStatus As Integer, dValorIRRF As Double, dValorICMS As Double, dValorICMSSubst As Double, dValorIPI As Double, dValorINSS As Double, sObservacao As String, lNumTitulo As Long, lNumIntDocOriginal As Long, iFilialEmpresa As Integer, iCreditoIPI As Integer, sSiglaDocumento As String, dSaldo As Double, iNumParcelas As Integer, iCreditoICMS As Integer, iFilialPedCompra As Integer, lNumPedCompra As Long, Optional sKey As String) As ClassTituloPagar
'
'Dim objNewMember As ClassTituloPagar
'Set objNewMember = New ClassTituloPagar
'
'    objNewMember.dtDataEmissao = dtDataEmissao
'    objNewMember.iFilial = iFilial
'    objNewMember.lNumIntDoc = lNumIntDoc
'    objNewMember.dOutrasDespesas = dOutrasDespesas
'    objNewMember.dValorFrete = dValorFrete
'    objNewMember.dValorProdutos = dValorProdutos
'    objNewMember.dValorSeguro = dValorSeguro
'    objNewMember.dValorTotal = dValorTotal
'    objNewMember.lFornecedor = lFornecedor
'    objNewMember.iStatus = iStatus
'    objNewMember.dValorIRRF = dValorIRRF
'    objNewMember.dValorICMS = dValorICMS
'    objNewMember.dValorICMSSubst = dValorICMSSubst
'    objNewMember.dValorIPI = dValorIPI
'    objNewMember.dValorINSS = dValorINSS
'    objNewMember.sObservacao = sObservacao
'    objNewMember.lNumTitulo = lNumTitulo
'    objNewMember.lNumIntDocOriginal = lNumIntDocOriginal
'    objNewMember.iFilialEmpresa = iFilialEmpresa
'    objNewMember.iCreditoIPI = iCreditoIPI
'    objNewMember.sSiglaDocumento = sSiglaDocumento
'    objNewMember.dSaldo = dSaldo
'    objNewMember.iNumParcelas = iNumParcelas
'    objNewMember.iCreditoICMS = iCreditoICMS
'    objNewMember.iFilialPedCompra = iFilialPedCompra
'    objNewMember.lNumPedCompra = lNumPedCompra
'
'    If Len(sKey) = 0 Then
'        mCol.Add objNewMember
'    Else
'        mCol.Add objNewMember, sKey
'    End If
'
'    Set Add = objNewMember
'    Set objNewMember = Nothing
'
'End Function
'
'Public Property Get Item(vntIndexKey As Variant) As ClassTituloPagar
'  Set Item = mCol(vntIndexKey)
'End Property
'
'Public Property Get Count() As Long
'    Count = mCol.Count
'End Property
'
'Public Sub Remove(vntIndexKey As Variant)
'    mCol.Remove vntIndexKey
'End Sub
'
'Public Property Get NewEnum() As IUnknown
'    Set NewEnum = mCol.[_NewEnum]
'End Property
'
'Private Sub Class_Initialize()
'    Set mCol = New Collection
'End Sub
'
'Private Sub Class_Terminate()
'    Set mCol = Nothing
'End Sub
'
