VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassSerie"
Attribute VB_Ext_KEY = "Member0" ,"ClassSerie"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(iFilialEmpresa As Integer, sSerie As String, lProxNumNFiscal As Long, lProxNumNFiscalEntrada As Long, lProxNumNFiscalImpressa As Long, iLockImpressao As Integer, iImprimindo As Integer, iTipoFormulario As Integer, Optional sKey As String) As ClassSerie
    
Dim objNewMember As ClassSerie
Set objNewMember = New ClassSerie

    objNewMember.iFilialEmpresa = iFilialEmpresa
    objNewMember.sSerie = sSerie
    objNewMember.lProxNumNFiscal = lProxNumNFiscal
    objNewMember.lProxNumNFiscalEntrada = lProxNumNFiscalEntrada
    objNewMember.lProxNumNFiscalImpressa = lProxNumNFiscalImpressa
    objNewMember.iLockImpressao = iLockImpressao
    objNewMember.iImprimindo = iImprimindo
    objNewMember.iTipoFormulario = iTipoFormulario
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassSerie
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

