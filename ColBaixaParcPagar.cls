VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColBaixaParcPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassBaixaParcPagar"
Attribute VB_Ext_KEY = "Member0" ,"ClassBaixaParcPagar"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(lNumIntDoc As Long, lNumIntBaixa As Long, lNumIntParcela As Long, iSequencial As Integer, dValorMulta As Double, dValorJuros As Double, dValorDesconto As Double, dValorBaixado As Double, iStatus As Integer, Optional sKey As String, Optional dValorDiferenca As Double, Optional iMotivoDiferenca As Integer) As ClassBaixaParcPagar

Dim objNewMember As ClassBaixaParcPagar
Set objNewMember = New ClassBaixaParcPagar

    objNewMember.lNumIntDoc = lNumIntDoc
    objNewMember.lNumIntBaixa = lNumIntBaixa
    objNewMember.lNumIntParcela = lNumIntParcela
    objNewMember.iSequencial = iSequencial
    objNewMember.dValorMulta = dValorMulta
    objNewMember.dValorJuros = dValorJuros
    objNewMember.dValorDesconto = dValorDesconto
    objNewMember.dValorBaixado = dValorBaixado
    objNewMember.iStatus = iStatus
   
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    If Not IsMissing(dValorDiferenca) Then
        objNewMember.dValorDiferenca = dValorDiferenca
    End If
    
    If Not IsMissing(iMotivoDiferenca) Then
        objNewMember.iMotivoDiferenca = iMotivoDiferenca
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassBaixaParcPagar
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

