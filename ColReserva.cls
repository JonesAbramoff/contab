VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClassReserva"
Attribute VB_Ext_KEY = "Member0" ,"ClassReserva"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public Function Add(iFilialEmpresa As Integer, lCodigo As Long, sProduto As String, iAlmoxarifado As Integer, iTipoDoc As Integer, lDocOrigem As Long, lNumIntOrigem As Long, dQuantidade As Double, dtDataReserva As Date, dtDataValidade As Date, sCodUsuario As String, sResponsavel As String, lNumIntDoc As Long, sAlmoxarifado As String, Optional sKey As String) As ClassReserva
    'create a new object
    Dim objNewMember As ClassReserva
    Set objNewMember = New ClassReserva


    'set the properties passed into the method
    objNewMember.iFilialEmpresa = iFilialEmpresa
    objNewMember.lCodigo = lCodigo
    objNewMember.sProduto = sProduto
    objNewMember.iAlmoxarifado = iAlmoxarifado
    objNewMember.iTipoDoc = iTipoDoc
    objNewMember.lDocOrigem = lDocOrigem
    objNewMember.lNumIntOrigem = lNumIntOrigem
    objNewMember.dQuantidade = dQuantidade
    objNewMember.dtDataReserva = dtDataReserva
    objNewMember.dtDataValidade = dtDataValidade
    objNewMember.sCodUsuario = sCodUsuario
    objNewMember.sResponsavel = sResponsavel
    objNewMember.lNumIntDoc = lNumIntDoc
    objNewMember.sAlmoxarifado = sAlmoxarifado




    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As ClassReserva
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

