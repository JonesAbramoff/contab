VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIN86Arquivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariModelo As Integer
Private mvariTipo As Integer
Private mvariSelecionado As Integer
Private mvarsNome As String
Private mvariDUMP As Integer
Private mvariRelatAcompanhamento As Integer
Private mvariLayout As Integer
Private mvariFilialEmpresa As Integer
Private mvariNumEtiqueta As Integer

'*** PROPERTY iModelo - INÍCIO ***
Public Property Let iModelo(ByVal vData As Integer)
    mvariModelo = vData
End Property

Public Property Get iModelo() As Integer
    iModelo = mvariModelo
End Property
'*** PROPERTY iModelo - FIM ***

'*** PROPERTY iTipo - INÍCIO ***
Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property
'*** PROPERTY iTipo - FIM ***

'*** PROPERTY iSelecionado - INÍCIO ***
Public Property Let iSelecionado(ByVal vData As Integer)
    mvariSelecionado = vData
End Property

Public Property Get iSelecionado() As Integer
    iSelecionado = mvariSelecionado
End Property
'*** PROPERTY iSelecionado - FIM ***

'*** PROPERTY sNome - INÍCIO ***
Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property
'*** PROPERTY sNome - FIM ***

'*** PROPERTY iDUMP - INÍCIO ***
Public Property Let iDUMP(ByVal vData As Integer)
    mvariDUMP = vData
End Property

Public Property Get iDUMP() As Integer
    iDUMP = mvariDUMP
End Property
'*** PROPERTY iDUMP - FIM ***

'*** PROPERTY iRelatAcompanhamento - INÍCIO ***
Public Property Let iRelatAcompanhamento(ByVal vData As Integer)
    mvariRelatAcompanhamento = vData
End Property

Public Property Get iRelatAcompanhamento() As Integer
    iRelatAcompanhamento = mvariRelatAcompanhamento
End Property
'*** PROPERTY iRelatAcompanhamento - FIM ***

'*** PROPERTY iLayout - INÍCIO ***
Public Property Let iLayout(ByVal vData As Integer)
    mvariLayout = vData
End Property

Public Property Get iLayout() As Integer
    iLayout = mvariLayout
End Property
'*** PROPERTY iRelatAcompanhamento - FIM ***

'*** PROPERTY iFilialEmpresa - INÍCIO ***
Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property
'*** PROPERTY iFilialEmpresa - FIM ***

'*** PROPERTY iNumEtiqueta - INÍCIO ***
Public Property Let iNumEtiqueta(ByVal vData As Integer)
    mvariNumEtiqueta = vData
End Property

Public Property Get iNumEtiqueta() As Integer
    iNumEtiqueta = mvariNumEtiqueta
End Property
'*** PROPERTY iNumEtiqueta - FIM ***

Sub Copiar(objOriginal As Object)

    With Me
    
        .iModelo = objOriginal.iModelo
        .iTipo = objOriginal.iTipo
        .iSelecionado = objOriginal.iSelecionado
        .sNome = objOriginal.sNome
        .iDUMP = objOriginal.iDUMP
        .iRelatAcompanhamento = objOriginal.iRelatAcompanhamento
        .iLayout = objOriginal.iLayout
        .iFilialEmpresa = objOriginal.iFilialEmpresa
        .iNumEtiqueta = objOriginal.iNumEtiqueta
        
    End With

End Sub
