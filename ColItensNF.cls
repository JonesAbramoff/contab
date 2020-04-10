VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColItensNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassItemNF"
Attribute VB_Ext_KEY = "Member0" ,"ClassItemNF"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(lNumIntNF As Long, iItem As Integer, sProduto As String, sUnidadeMed As String, dQuantidade As Double, dPrecoUnitario As Double, dPercDesc As Double, dValorDesconto As Double, dtDataEntrega As Date, sDescricaoItem As String, dValorAbatComissao As Double, lNumIntPedVenda As Long, lNumIntDoc As Long, lNumIntTrib As Long, lNumIntItemPedVenda As Long, ColAlocacoes As ColAlocacoesItemNF, iAlmoxarifado As Integer, sAlmoxarifadoNomeRed As String, sCCL As String, iStatus As Integer, lNumIntDocOrig As Long, sSerieNFOrig As String, lNumNFOrig As Long, iItemNFOrig As Integer, iControleEstoque As Integer, sUMEstoque As String, iClasseUM As Integer, dPrecoUnitarioMoeda As Double, Optional sKey As String, Optional colRastreamento As Collection) As ClassItemNF

Dim objNewMember As ClassItemNF
Set objNewMember = New ClassItemNF

    objNewMember.lNumIntNF = lNumIntNF
    objNewMember.iItem = iItem
    objNewMember.sProduto = sProduto
    objNewMember.sUnidadeMed = sUnidadeMed
    objNewMember.dQuantidade = dQuantidade
    objNewMember.dPrecoUnitario = dPrecoUnitario
    objNewMember.dPercDesc = dPercDesc
    objNewMember.dValorDesconto = dValorDesconto
    objNewMember.dtDataEntrega = dtDataEntrega
    objNewMember.sDescricaoItem = sDescricaoItem
    objNewMember.dValorAbatComissao = dValorAbatComissao
    objNewMember.lNumIntPedVenda = lNumIntPedVenda
    objNewMember.lNumIntDoc = lNumIntDoc
    objNewMember.lNumIntTrib = lNumIntTrib
    objNewMember.lNumIntItemPedVenda = lNumIntItemPedVenda
    Set objNewMember.ColAlocacoes = ColAlocacoes
    objNewMember.iAlmoxarifado = iAlmoxarifado
    objNewMember.sAlmoxarifadoNomeRed = sAlmoxarifadoNomeRed
    objNewMember.iStatus = iStatus
    objNewMember.lNumIntDocOrig = lNumIntDocOrig
    objNewMember.sSerieNFOrig = sSerieNFOrig
    objNewMember.lNumNFOrig = lNumNFOrig
    objNewMember.iItemNFOrig = iItemNFOrig
    objNewMember.iControleEstoque = iControleEstoque
    objNewMember.sUMEstoque = sUMEstoque
    objNewMember.iClasseUM = iClasseUM
    objNewMember.sCCL = sCCL
    objNewMember.dPrecoUnitarioMoeda = dPrecoUnitarioMoeda

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    If Not IsMissing(colRastreamento) Then
        Set objNewMember.colRastreamento = colRastreamento
    End If
    
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Function Add1(objItemNF As ClassItemNF, Optional sKey As String) As ClassItemNF

    If Len(sKey) = 0 Then
        mCol.Add objItemNF
    Else
        mCol.Add objItemNF, sKey
    End If

End Function


Public Property Get Item(vntIndexKey As Variant) As ClassItemNF
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

