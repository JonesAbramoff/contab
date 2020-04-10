VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassArqCNABPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjFonte As Object

Private mvarlNumIntBordero As Long

Private mvarlQuantRegistros As Long

Private mvarcolDetArqCNABPag As New Collection

Private mvarlSequencialLote As Long

Private mvardValorTotal As Double

Public Property Let lNumIntBordero(ByVal vData As Long)
    mvarlNumIntBordero = vData
End Property

Public Property Get lNumIntBordero() As Long
    lNumIntBordero = mvarlNumIntBordero
End Property

Public Sub AcumulaValor(ByVal dValor As Double)

    mvardValorTotal = mvardValorTotal + dValor
    
End Sub

Public Property Get lSequencialLote() As Long
    lSequencialLote = mvarlSequencialLote
End Property

Public Property Get colDetArqCNABPag() As Collection
    Set colDetArqCNABPag = mvarcolDetArqCNABPag
End Property

Public Property Get lQuantRegistros() As Long
    lQuantRegistros = mvarlQuantRegistros
End Property

Public Function ObterProxReg() As Long

    mvarlQuantRegistros = mvarlQuantRegistros + 1
    ObterProxReg = mvarlQuantRegistros
    
End Function

Public Function IniciarLote(ByVal objLote As ClassLoteArqCNABPag) As Long
    
Dim lErro As Long

On Error GoTo Erro_IniciarLote

    mvarlSequencialLote = mvarlSequencialLote + 1
    
    lErro = objFonte.Montar_Registro_HeaderLote(objLote)
    If lErro <> SUCESSO Then gError 106564
    
    IniciarLote = SUCESSO
     
    Exit Function
    
Erro_IniciarLote:

    IniciarLote = gErr
     
    Select Case gErr
          
        Case 106564
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144633)
     
    End Select
     
    Exit Function

End Function

Public Function FinalizarLote(ByVal objLote As ClassLoteArqCNABPag) As Long

Dim lErro As Long

On Error GoTo Erro_FinalizarLote

    lErro = objFonte.Montar_Registro_TrailerLote(objLote)
    If lErro <> SUCESSO Then gError 106565
    
    FinalizarLote = SUCESSO
     
    Exit Function
    
Erro_FinalizarLote:

    FinalizarLote = gErr
     
    Select Case gErr
          
        Case 106565
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144634)
     
    End Select
     
    Exit Function
    
End Function

Public Property Set objFonte(ByVal vData As Object)
    Set mvarobjFonte = vData
End Property

Public Property Get objFonte() As Object
    Set objFonte = mvarobjFonte
End Property

Public Function Iniciar(ByVal objFontePar As Object, ByVal lNumIntBordero As Long) As Long

Dim lErro As Long, objAux As ClassArqCNABPag

On Error GoTo Erro_Iniciar

    Set mvarobjFonte = objFontePar
    mvarlNumIntBordero = lNumIntBordero
    
    'Lê as baixas ligadas ao bordero
    Set objAux = Me '??? apenas para evitar erro 438 que dava qdo passava 'Me' diretamente por CF.
    lErro = CF("ArqCNABPag_Carregar_DetPag", objAux)
    If lErro <> SUCESSO Then gError 51726
    
    'Monta o registro de Header do arquivo
    lErro = objFonte.Montar_RegistroHeaderCNABPag(Me)
    If lErro <> SUCESSO Then gError 51724
    
    Iniciar = SUCESSO
     
    Exit Function
    
Erro_Iniciar:

    Iniciar = gErr
     
    Select Case gErr
          
        Case 51724, 51726
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144635)
     
    End Select
     
    Exit Function

End Function

Public Function ProcessarDetalhe(ByVal objDetArqCNABPag As ClassDetArqCNABPag) As Long

    ProcessarDetalhe = objFonte.BorderoPagto_ProcessarDetPag(objDetArqCNABPag)
    
End Function

Public Function Finalizar() As Long

Dim lErro As Long

On Error GoTo Erro_Finalizar

    'Monta o registro de Header do arquivo
    lErro = objFonte.Montar_Registro_TrailerArq(Me)
    If lErro <> SUCESSO Then gError 51724
    
    Set objFonte = Nothing
    
    Finalizar = SUCESSO
     
    Exit Function
    
Erro_Finalizar:

    Finalizar = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144636)
     
    End Select
     
    Exit Function

End Function

