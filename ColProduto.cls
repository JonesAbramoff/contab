VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassProduto"
Attribute VB_Ext_KEY = "Member0" ,"ClassProduto"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public Function Add(sCodigo As String, sDescricao As String, iGerencial As Integer, iAtivo As Integer, iClasseUM As Integer, sSiglaUMEstoque As String, sSiglaUMCompra As String, sSiglaUMVenda As String, sSubstituto1 As String, sSubstituto2 As String, iTipo As Integer, sNomeReduzido As String, sModelo As String, iNivel As Integer, iPrazoValidade As Integer, sCodigoBarras As String, iEtiquetasCodBarras As Integer, dPesoLiq As Double, dPesoBruto As Double, dComprimento As Double, dEspessura As Double, dLargura As Double, sCor As String, sObsFisica As String, iFaturamento As Integer, iCompras As Integer, iPCP As Integer, iKitBasico As Integer, iKitInt As Integer, colCategoriaItem As Collection, dIPIAliquota As Double, sIPICodigo As String, sIPICodDIPI As String, iControleEstoque As Integer, iIPIAgregaCusto As Integer, iICMSAgregaCusto As Integer, iFreteAgregaCusto As Integer, iApropriacaoCusto As Integer, Optional sKey As String) As ClassProduto
    'create a new object
    Dim objNewMember As ClassProduto
    Set objNewMember = New ClassProduto


    'set the properties passed into the method
    objNewMember.sCodigo = sCodigo
    objNewMember.sDescricao = sDescricao
    objNewMember.iGerencial = iGerencial
    objNewMember.iAtivo = iAtivo
    objNewMember.iClasseUM = iClasseUM
    objNewMember.sSiglaUMEstoque = sSiglaUMEstoque
    objNewMember.sSiglaUMCompra = sSiglaUMCompra
    objNewMember.sSiglaUMVenda = sSiglaUMVenda
    objNewMember.sSubstituto1 = sSubstituto1
    objNewMember.sSubstituto2 = sSubstituto2
    objNewMember.iTipo = iTipo
    objNewMember.sNomeReduzido = sNomeReduzido
    objNewMember.sModelo = sModelo
    objNewMember.iNivel = iNivel
    objNewMember.iPrazoValidade = iPrazoValidade
    objNewMember.sCodigoBarras = sCodigoBarras
    objNewMember.iEtiquetasCodBarras = iEtiquetasCodBarras
    objNewMember.dPesoLiq = dPesoLiq
    objNewMember.dPesoBruto = dPesoBruto
    objNewMember.dComprimento = dComprimento
    objNewMember.dEspessura = dEspessura
    objNewMember.dLargura = dLargura
    objNewMember.sCor = sCor
    objNewMember.sObsFisica = sObsFisica
    objNewMember.iFaturamento = iFaturamento
    objNewMember.iCompras = iCompras
    objNewMember.iPCP = iPCP
    objNewMember.iKitBasico = iKitBasico
    objNewMember.iKitInt = iKitInt
    Set objNewMember.colCategoriaItem = colCategoriaItem
    objNewMember.dIPIAliquota = dIPIAliquota
    objNewMember.sIPICodigo = sIPICodigo
    objNewMember.sIPICodDIPI = sIPICodDIPI
    objNewMember.iControleEstoque = iControleEstoque
    objNewMember.iIPIAgregaCusto = iIPIAgregaCusto
    objNewMember.iICMSAgregaCusto = iICMSAgregaCusto
    objNewMember.iFreteAgregaCusto = iFreteAgregaCusto
    objNewMember.iApropriacaoCusto = iApropriacaoCusto




    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As ClassProduto
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

