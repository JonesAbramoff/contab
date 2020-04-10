VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColTipoMovtoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Public Function Add(iCodigo As Integer, _
                    sDescricao As String, _
                    sNomeReduzido As String, _
                    sSigla As String, _
                    sSignificadoRefInterna As String, _
                    iCredito As Integer, _
                    iPodeCaixaECF As Integer, _
                    iPodeCaixaCentral As Integer, _
                    iPodeCaixaGeral As Integer, _
                    iRelMovCaixa As Integer, _
                    iTipoMeioPagto As Integer, _
                    Optional sKey As String) As ClassTipoMovtoCaixa
                    
Dim objNewMember As New ClassTipoMovtoCaixa

    With objNewMember
    
        .iCodigo = iCodigo
        .iCredito = iCredito
        .iPodeCaixaCentral = iPodeCaixaCentral
        .iPodeCaixaECF = iPodeCaixaECF
        .iPodeCaixaGeral = iPodeCaixaGeral
        .iRelMovCaixa = iRelMovCaixa
        .iTMPLoja = iTipoMeioPagto
        .sDescricao = sDescricao
        .sNomeReduzido = sNomeReduzido
        .sSigla = sSigla
        .sSignificadoRefInterna = sSignificadoRefInterna
    
    End With
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As ClassTipoMovtoCaixa
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
