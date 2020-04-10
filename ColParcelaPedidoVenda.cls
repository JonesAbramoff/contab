VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColParcelaPedidoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassParcelaPedidoVenda"
Attribute VB_Ext_KEY = "Member0" ,"ClassParcelaPedidoVenda"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(dValor As Double, dtDataVencimento As Date, iNumParcela As Integer, iDesconto1Codigo As Integer, dtDesconto1Ate As Date, dDesconto1Valor As Double, iDesconto2Codigo As Integer, dtDesconto2Ate As Date, dDesconto2Valor As Double, dtDesconto3Ate As Date, dDesconto3Valor As Double, iDesconto3Codigo As Integer, Optional sKey As String) As ClassParcelaPedidoVenda

Dim objNewMember As ClassParcelaPedidoVenda
Set objNewMember = New ClassParcelaPedidoVenda

    objNewMember.dValor = dValor
    objNewMember.dtDataVencimento = dtDataVencimento
    objNewMember.iNumParcela = iNumParcela
    objNewMember.iDesconto1Codigo = iDesconto1Codigo
    objNewMember.dtDesconto1Ate = dtDesconto1Ate
    objNewMember.dDesconto1Valor = dDesconto1Valor
    objNewMember.iDesconto2Codigo = iDesconto2Codigo
    objNewMember.dtDesconto2Ate = dtDesconto2Ate
    objNewMember.dDesconto2Valor = dDesconto2Valor
    objNewMember.dtDesconto3Ate = dtDesconto3Ate
    objNewMember.dDesconto3Valor = dDesconto3Valor
    objNewMember.iDesconto3Codigo = iDesconto3Codigo

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassParcelaPedidoVenda
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

