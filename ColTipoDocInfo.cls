VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColTipoDocInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassTipoDocInfo"
Attribute VB_Ext_KEY = "Member0" ,"ClassTipoDocInfo"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(sSigla As String, sDescricao As String, iTipoMovtoEstoque As Integer, iTipoMovtoEstoque2 As Integer, iTipoMovtoEstoqueBenef As Integer, sNaturezaOperacaoPadrao As String, iInfoContabilizacao As Integer, sTipoDocCPR As String, iCodigo As Integer, sNomeReduzido As String, sNomeTelaNFiscal As String, sTituloTelaNFiscal As String, iFaturavel As Integer, iComissao As Integer, iEmitente As Integer, iDestinatario As Integer, iOrigem As Integer, iComplementar As Integer, iTipo As Integer, sSiglaNFOriginal As String, iFaturamento As Integer, iModeloArqICMS As Integer, iPadrao As Integer, sNatOpExtPadrao As String, iNFFatura As Integer, iTipoOperacaoTrib As Integer, Optional sKey As String) As ClassTipoDocInfo
    
Dim objNewMember As ClassTipoDocInfo
Set objNewMember = New ClassTipoDocInfo

    objNewMember.sSigla = sSigla
    objNewMember.sDescricao = sDescricao
    objNewMember.iTipoMovtoEstoque = iTipoMovtoEstoque
    objNewMember.iTipoMovtoEstoque2 = iTipoMovtoEstoque2
    objNewMember.iTipoMovtoEstoqueBenef = iTipoMovtoEstoqueBenef
    objNewMember.sNaturezaOperacaoPadrao = sNaturezaOperacaoPadrao
    objNewMember.iInfoContabilizacao = iInfoContabilizacao
    objNewMember.sTipoDocCPR = sTipoDocCPR
    objNewMember.iCodigo = iCodigo
    objNewMember.sNomeReduzido = sNomeReduzido
    objNewMember.sNomeTelaNFiscal = sNomeTelaNFiscal
    objNewMember.sTituloTelaNFiscal = sTituloTelaNFiscal
    objNewMember.iFaturavel = iFaturavel
    objNewMember.iComissao = iComissao
    objNewMember.iEmitente = iEmitente
    objNewMember.iDestinatario = iDestinatario
    objNewMember.iOrigem = iOrigem
    objNewMember.iComplementar = iComplementar
    objNewMember.iTipo = iTipo
    objNewMember.sSiglaNFOriginal = sSiglaNFOriginal
    objNewMember.iFaturamento = iFaturamento
    objNewMember.iModeloArqICMS = iModeloArqICMS
    objNewMember.iPadrao = iPadrao
    objNewMember.sNatOpExtPadrao = sNatOpExtPadrao
    objNewMember.iNFFatura = iNFFatura
    objNewMember.iTipoOperacaoTrib = iTipoOperacaoTrib

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassTipoDocInfo
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

