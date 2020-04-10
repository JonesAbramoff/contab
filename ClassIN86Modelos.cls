VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIN86Modelos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsModelo As String
Private mvardtDataInicio As Date
Private mvardtDataFim As Date
Private mvariMeioEntrega As Integer
Private mvariEtiquetas As Integer
Private mvarcolIN86Arquivos As New Collection

Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let sModelo(ByVal vData As String)
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Set colIN86Arquivos(ByVal vData As Object)
    Set mvarcolIN86Arquivos = vData
End Property

Public Property Get colIN86Arquivos() As Collection
    Set colIN86Arquivos = mvarcolIN86Arquivos
End Property

Public Property Let iMeioEntrega(ByVal vData As Integer)
    mvariMeioEntrega = vData
End Property

Public Property Get iMeioEntrega() As Integer
    iMeioEntrega = mvariMeioEntrega
End Property

Public Property Let iEtiquetas(ByVal vData As Integer)
    mvariEtiquetas = vData
End Property

Public Property Get iEtiquetas() As Integer
    iEtiquetas = mvariEtiquetas
End Property

Public Sub Copiar(objOriginal As Object)

Dim objAux As Object, objIN86Arquivo As ClassIN86Arquivos

    With Me
        .iCodigo = objOriginal.iCodigo
        .sModelo = objOriginal.sModelo
        .dtDataInicio = objOriginal.dtDataInicio
        .dtDataFim = objOriginal.dtDataFim
        .iMeioEntrega = objOriginal.iMeioEntrega
        .iEtiquetas = objOriginal.iEtiquetas
    End With
    
    Set Me.colIN86Arquivos = New Collection
        
    For Each objAux In objOriginal.colIN86Arquivos

        Set objIN86Arquivo = New ClassIN86Arquivos
        
        Call objIN86Arquivo.Copiar(objAux)
        
        Call Me.colIN86Arquivos.Add(objIN86Arquivo)
        
    Next
        
End Sub
