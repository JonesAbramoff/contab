VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColItemPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassItemPedido"
Attribute VB_Ext_KEY = "Member0" ,"ClassItemPedido"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(iFilialEmpresa As Integer, lCodPedido As Long, iItem As Integer, sProduto As String, dQuantidade As Double, dPrecoUnitario As Double, dPrecoTotal As Double, dValorDesconto As Double, dtDataEntrega As Date, sProdutoDescricao As String, dValorAbatComissao As Double, dQuantCancelada As Double, dQuantReservada As Double, colReserva As colReserva, sProdutoNomeReduzido As String, sUnidadeMed As String, sLote As String, sUMEstoque As String, iClasseUM As Integer, dQuantFaturada As Double, dQuantAFaturar As Double, sDescricao As String, iStatus As Integer, iControleEstoque As Integer, dQuantOP As Double, dQuantSC As Double, dPercDesc1 As Double, iTipoDesc1 As Integer, dPercDesc2 As Double, iTipoDesc2 As Integer, dPercDesc3 As Double, iTipoDesc3 As Integer, lNumIntDoc As Long, Optional sKey As String, Optional ByVal dPrecoBase As Double) As ClassItemPedido

Dim objNewMember As ClassItemPedido
Dim lErro As Long
On Error GoTo Erro_Add

    Set objNewMember = New ClassItemPedido

    objNewMember.iFilialEmpresa = iFilialEmpresa
    objNewMember.lCodPedido = lCodPedido
    objNewMember.iItem = iItem
    objNewMember.sProduto = sProduto
    objNewMember.dQuantidade = dQuantidade
    objNewMember.dPrecoUnitario = dPrecoUnitario
    objNewMember.dPrecoTotal = dPrecoTotal
    objNewMember.dValorDesconto = dValorDesconto
    objNewMember.dtDataEntrega = dtDataEntrega
    objNewMember.sProdutoDescricao = sProdutoDescricao
    objNewMember.dValorAbatComissao = dValorAbatComissao
    objNewMember.dQuantCancelada = dQuantCancelada
    objNewMember.dQuantReservada = dQuantReservada
    Set objNewMember.colReserva = colReserva
    objNewMember.sProdutoNomeReduzido = sProdutoNomeReduzido
    objNewMember.sUnidadeMed = sUnidadeMed
    objNewMember.sLote = sLote
    objNewMember.sUMEstoque = sUMEstoque
    objNewMember.iClasseUM = iClasseUM
    objNewMember.dQuantFaturada = dQuantFaturada
    objNewMember.dQuantAFaturar = dQuantAFaturar
    objNewMember.sDescricao = sDescricao
    objNewMember.iStatus = iStatus
    objNewMember.iControleEstoque = iControleEstoque
    objNewMember.dQuantOP = dQuantOP
    objNewMember.dQuantSC = dQuantSC
    
    objNewMember.dPercDesc1 = dPercDesc1
    objNewMember.iTipoDesc1 = iTipoDesc1
    objNewMember.dPercDesc2 = dPercDesc2
    objNewMember.iTipoDesc2 = iTipoDesc2
    objNewMember.dPercDesc3 = dPercDesc3
    objNewMember.iTipoDesc3 = iTipoDesc3
    
    objNewMember.lNumIntDoc = lNumIntDoc

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    If Not IsMissing(dPrecoBase) Then
        objNewMember.dPrecoBase = dPrecoBase
    End If

    Set Add = objNewMember

    Exit Function

Erro_Add:

    Set Add = Nothing

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154330)

    End Select

    Exit Function

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassItemPedido
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

Public Function AddObj(objNewMember As ClassItemPedido, Optional sKey As String) As ClassItemPedido

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set AddObj = objNewMember

End Function

