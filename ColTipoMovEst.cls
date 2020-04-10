VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColTipoMovEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassTipoMovEst"
Attribute VB_Ext_KEY = "Member0" ,"ClassTipoMovEst"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(iCodigo As Integer, sDescricao As String, sEntradaOuSaida As String, iInativo As Integer, iAtualizaConsumo As Integer, iAtualizaVenda As Integer, iEditavel As Integer, iValidoMovInt As Integer, iAtualizaCusto As Integer, iAtualizaConsig As Integer, iAtualizaDemo As Integer, iAtualizaConserto As Integer, iProdutoDeTerc As Integer, iAtualizaOutras As Integer, iAtualizaIndOutras As Integer, iAtualizaNossaDisp As Integer, iAtualizaDefeituosa As Integer, iAtualizaRecebIndisp As Integer, iTransferencia As Integer, iAtualizaCMRProd As Integer, iAtualizaBenef As Integer, iCustoInformado As Integer, iValidaDataUltMov As Integer, _
    iAtualizaMovEstoque As Integer, iAtualizaReserva As Integer, _
    iAtualizaCompra As Integer, sDescrNumIntDocOrigem As String, iTipoNumIntDocOrigem As Integer, iAtualizaOP As Integer, iKardex As Integer, _
    sSigla As String, sNomeTela As String, iLivroMod3 As Integer, sEntradaSaidaCMP As String, iNFDevolucao As Integer, _
    Optional sKey As String) As ClassTipoMovEst

Dim objNewMember As ClassTipoMovEst
Set objNewMember = New ClassTipoMovEst

    objNewMember.iCodigo = iCodigo
    objNewMember.sDescricao = sDescricao
    objNewMember.sEntradaOuSaida = sEntradaOuSaida
    objNewMember.iInativo = iInativo
    objNewMember.iAtualizaConsumo = iAtualizaConsumo
    objNewMember.iAtualizaVenda = iAtualizaVenda
    objNewMember.iEditavel = iEditavel
    objNewMember.iValidoMovInt = iValidoMovInt
    objNewMember.iAtualizaCusto = iAtualizaCusto
    objNewMember.iAtualizaConsig = iAtualizaConsig
    objNewMember.iAtualizaDemo = iAtualizaDemo
    objNewMember.iAtualizaConserto = iAtualizaConserto
    objNewMember.iProdutoDeTerc = iProdutoDeTerc
    objNewMember.iAtualizaOutras = iAtualizaOutras
    objNewMember.iAtualizaIndOutras = iAtualizaIndOutras
    objNewMember.iAtualizaNossaDisp = iAtualizaNossaDisp
    objNewMember.iAtualizaDefeituosa = iAtualizaDefeituosa
    objNewMember.iAtualizaRecebIndisp = iAtualizaRecebIndisp
    objNewMember.iTransferencia = iTransferencia
    objNewMember.iAtualizaCMRProd = iAtualizaCMRProd
    objNewMember.iCustoInformado = iCustoInformado
    objNewMember.iValidaDataUltMov = iValidaDataUltMov
    objNewMember.iAtualizaBenef = iAtualizaBenef
    
    objNewMember.iAtualizaMovEstoque = iAtualizaMovEstoque
    objNewMember.iAtualizaReserva = iAtualizaReserva
    objNewMember.iAtualizaCompra = iAtualizaCompra
    objNewMember.sDescrNumIntDocOrigem = sDescrNumIntDocOrigem
    objNewMember.iTipoNumIntDocOrigem = iTipoNumIntDocOrigem
    objNewMember.iAtualizaOP = iAtualizaOP
    objNewMember.iKardex = iKardex
    objNewMember.sSigla = sSigla
    objNewMember.sNomeTela = sNomeTela
    objNewMember.iLivroMod3 = iLivroMod3
    objNewMember.sEntradaSaidaCMP = sEntradaSaidaCMP
    objNewMember.iNFDevolucao = iNFDevolucao
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassTipoMovEst
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

