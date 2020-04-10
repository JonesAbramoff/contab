VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColChequePre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassChequePre"
Attribute VB_Ext_KEY = "Member0" ,"ClassChequePre"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"ClassChequePre"
Option Explicit

Private mCol As Collection

Public Function Add(lNumIntCheque As Long, lCliente As Long, iFilial As Integer, iBanco As Integer, sAgencia As String, sContaCorrente As String, lNumero As Long, dtDataDeposito As Date, dValor As Double, lNumBordero As Long, Optional sKey As String) As ClassChequePre
    
Dim objNewMember As ClassChequePre
Set objNewMember = New ClassChequePre

    objNewMember.lNumIntCheque = lNumIntCheque
    objNewMember.lCliente = lCliente
    objNewMember.iFilial = iFilial
    objNewMember.iBanco = iBanco
    objNewMember.sAgencia = sAgencia
    objNewMember.sContaCorrente = sContaCorrente
    objNewMember.lNumero = lNumero
    objNewMember.dtDataDeposito = dtDataDeposito
    objNewMember.dValor = dValor
    objNewMember.lNumBordero = lNumBordero

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassChequePre
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

