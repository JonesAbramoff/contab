VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassTipoDocumento"
Attribute VB_Ext_KEY = "Member0" ,"ClassTipoDocumento"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"ClassTipoDocumento"
Option Explicit

Private mCol As Collection

Public Function Add(sDescricao As String, iContabiliza As Integer, iAumentaValorPagto As Integer, iEmNFFatPag As Integer, iEmCreditoPagForn As Integer, iEmTituloRec As Integer, iEmDebitosRecCli As Integer, sSigla As String, sDescricaoReduzida As String, Optional sKey As String) As ClassTipoDocumento

Dim objNewMember As ClassTipoDocumento
Set objNewMember = New ClassTipoDocumento

    objNewMember.sDescricao = sDescricao
    objNewMember.iContabiliza = iContabiliza
    objNewMember.iAumentaValorPagto = iAumentaValorPagto
    objNewMember.iEmNFFatPag = iEmNFFatPag
    objNewMember.iEmCreditoPagForn = iEmCreditoPagForn
    objNewMember.iEmTituloRec = iEmTituloRec
    objNewMember.iEmDebitosRecCli = iEmDebitosRecCli
    objNewMember.sSigla = sSigla
    objNewMember.sDescricaoReduzida = sDescricaoReduzida
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassTipoDocumento
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

