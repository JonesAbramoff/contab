VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColItensMovEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassItemMovEstoque"
Attribute VB_Ext_KEY = "Member0" ,"ClassItemMovEstoque"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(lNumIntDoc As Long, iTipoMov As Integer, dCusto As Double, iApropriacao As Integer, sProduto As String, sProdutoDesc As String, sSiglaUM As String, dQuantidade As Double, iAlmoxarifado As Integer, sAlmoxarifadoNomeRed As String, lNumIntDocOrigem As Long, sCcl As String, lNumIntDocEst As Long, sOPCodigo As String, sProdutoOP As String, sContaContabilAplic As String, sContaContabilEst As String, lHorasMaquina As Long, colRastreamentoMovto As Collection, colApropriacaoInsumo As Collection, dtDataInicioProducao As Date, Optional sKey As String) As ClassItemMovEstoque

Dim objNewMember As ClassItemMovEstoque
    
Set objNewMember = New ClassItemMovEstoque

    objNewMember.lNumIntDoc = lNumIntDoc
    objNewMember.iTipoMov = iTipoMov
    objNewMember.dCusto = dCusto
    objNewMember.iApropriacao = iApropriacao
    objNewMember.sProduto = sProduto
    objNewMember.sProdutoDesc = sProdutoDesc
    objNewMember.sSiglaUM = sSiglaUM
    objNewMember.dQuantidade = dQuantidade
    objNewMember.iAlmoxarifado = iAlmoxarifado
    objNewMember.sAlmoxarifadoNomeRed = sAlmoxarifadoNomeRed
    objNewMember.lNumIntDocOrigem = lNumIntDocOrigem
    objNewMember.sCcl = sCcl
    objNewMember.lNumIntDocEst = lNumIntDocEst
    objNewMember.sOPCodigo = sOPCodigo
    objNewMember.sProdutoOP = sProdutoOP
    objNewMember.sContaContabilAplic = sContaContabilAplic
    objNewMember.sContaContabilEst = sContaContabilEst
    objNewMember.lHorasMaquina = lHorasMaquina
    Set objNewMember.colRastreamentoMovto = colRastreamentoMovto
    Set objNewMember.colApropriacaoInsumo = colApropriacaoInsumo
    objNewMember.dtDataInicioProducao = dtDataInicioProducao

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Sub Add1(objItemMovEstoque As ClassItemMovEstoque)
        mCol.Add objItemMovEstoque
End Sub

Public Property Get Item(vntIndexKey As Variant) As ClassItemMovEstoque
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


