VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColEstoqueProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassEstoqueProduto"
Attribute VB_Ext_KEY = "Member0" ,"ClassEstoqueProduto"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(sProduto As String, iAlmoxarifado As Integer, sAlmoxarifadoNomeReduzido As String, dQuantDispNossa As Double, dSaldo As Double, sLocalizacaoFisica As String, sContaContabil As String, dQuantReservada As Double, dtDataInventario As Date, _
    dQuantidadeInicial As Double, dSaldoInicial As Double, dtDataInicial As Date, dQuantEmpenhada As Double, _
    dQuantRecIndl As Double, dQuantInd As Double, dQuantDefeituosa As Double, _
    dQuantConsig As Double, dQuantConsig3 As Double, dQuantDemo As Double, dQuantDemo3 As Double, _
    dQuantConserto3 As Double, dQuantConserto As Double, dQuantOutras As Double, dQuantOutras3 As Double, _
    dQuantBenef As Double, dQuantBenef3 As Double, _
    Optional sKey As String) As ClassEstoqueProduto

Dim objNewMember As ClassEstoqueProduto
Set objNewMember = New ClassEstoqueProduto
    
    With objNewMember
        .sProduto = sProduto
        .iAlmoxarifado = iAlmoxarifado
        .sAlmoxarifadoNomeReduzido = sAlmoxarifadoNomeReduzido
        .dQuantDispNossa = dQuantDispNossa
        .dSaldo = dSaldo
        .sLocalizacaoFisica = sLocalizacaoFisica
        .sContaContabil = sContaContabil
        .dQuantReservada = dQuantReservada
        .dtDataInventario = dtDataInventario
        .dQuantidadeInicial = dQuantidadeInicial
        .dSaldoInicial = dSaldoInicial
        .dtDataInicial = dtDataInicial
        .dQuantEmpenhada = dQuantEmpenhada
        .dQuantRecIndl = dQuantRecIndl
        .dQuantInd = dQuantInd
        .dQuantDefeituosa = dQuantDefeituosa
        .dQuantConsig = dQuantConsig
        .dQuantConsig3 = dQuantConsig3
        .dQuantDemo = dQuantDemo
        .dQuantDemo3 = dQuantDemo3
        .dQuantConserto3 = dQuantConserto3
        .dQuantConserto = dQuantConserto
        .dQuantOutras = dQuantOutras
        .dQuantOutras3 = dQuantOutras3
        .dQuantBenef = dQuantBenef
        .dQuantBenef3 = dQuantBenef3
    End With
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassEstoqueProduto
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

