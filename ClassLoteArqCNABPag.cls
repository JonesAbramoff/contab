VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLoteArqCNABPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjArq As ClassArqCNABPag
Private mvarlQuantRegistros As Long
Private mvardValorTotal As Double
Private mvariTipoCobranca As Integer
Private mvariBancoCobrador As Integer
Private mvarsFormaPagamento As String

Private mvarlSequencialDet As Long

Public Property Let sFormaPagamento(ByVal vData As String)
    mvarsFormaPagamento = vData
End Property

Public Property Get sFormaPagamento() As String
    sFormaPagamento = mvarsFormaPagamento
End Property


Public Property Get lSequencialDet() As Long
'obs.:o sequencial é o mesmo para todos os segmentos de um detalhe
    
    lSequencialDet = mvarlSequencialDet
    
End Property

Public Property Let iBancoCobrador(ByVal vData As Integer)
    mvariBancoCobrador = vData
End Property

Public Property Get iBancoCobrador() As Integer
    iBancoCobrador = mvariBancoCobrador
End Property

Public Property Let iTipoCobranca(ByVal vData As Integer)
    mvariTipoCobranca = vData
End Property

Public Property Get iTipoCobranca() As Integer
    iTipoCobranca = mvariTipoCobranca
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Get lQuantRegistros() As Long
    lQuantRegistros = mvarlQuantRegistros
End Property

Function IniciarLocal(ByVal objDet As ClassDetArqCNABPag) As Long

Dim lErro As Long

On Error GoTo Erro_IniciarLocal

    mvarlSequencialDet = 0
    mvarlQuantRegistros = 0
    mvardValorTotal = 0
    mvariTipoCobranca = objDet.iTipoCobranca
    mvariBancoCobrador = objDet.iBancoCobrador
    mvarsFormaPagamento = objDet.sFormaPagamento
    
    lErro = objArq.IniciarLote(Me)
    If lErro <> SUCESSO Then gError 106566
    
    IniciarLocal = SUCESSO
     
    Exit Function
    
Erro_IniciarLocal:

    IniciarLocal = gErr
     
    Select Case gErr
          
        Case 106566
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151532)
     
    End Select
     
    Exit Function

End Function

Public Function Iniciar(ByVal objDet As ClassDetArqCNABPag, ByVal objArqPar As ClassArqCNABPag) As Long

Dim lErro As Long

On Error GoTo Erro_Iniciar

    Set mvarobjArq = objArqPar
    
    lErro = IniciarLocal(objDet)
    If lErro <> SUCESSO Then gError 106567
    
    Iniciar = SUCESSO
     
    Exit Function
    
Erro_Iniciar:

    Iniciar = gErr
     
    Select Case gErr
              
        Case 106567
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151533)
     
    End Select
     
    Exit Function

End Function

Public Function Trocou(ByVal objDet As ClassDetArqCNABPag) As Boolean
'retorna TRUE se o registro informado nao irá pertencer ao mesmo lote
        
    Trocou = mvarsFormaPagamento <> objDet.sFormaPagamento

End Function

Public Property Get objArq() As ClassArqCNABPag
    Set objArq = mvarobjArq
End Property

Public Function ProcessarDetalhe(ByVal objDet As ClassDetArqCNABPag) As Long
'
Dim lErro As Long

On Error GoTo Erro_ProcessarDetalhe

    'o sequencial é o mesmo para todos os segmentos de um detalhe
    mvarlSequencialDet = mvarlSequencialDet + 1
    
    lErro = objDet.Iniciar(Me)
    If lErro <> SUCESSO Then gError 106568
    
    lErro = objArq.ProcessarDetalhe(objDet)
    If lErro <> SUCESSO Then gError 106569
        
    ProcessarDetalhe = SUCESSO
     
    Exit Function
    
Erro_ProcessarDetalhe:

    ProcessarDetalhe = gErr
     
    Select Case gErr
          
        Case 106568, 106569
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151534)
     
    End Select
     
    Exit Function

End Function

Public Function Reiniciar(ByVal objDet As ClassDetArqCNABPag) As Long

Dim lErro As Long

On Error GoTo Erro_Reiniciar

    lErro = objArq.FinalizarLote(Me)
    If lErro <> SUCESSO Then gError 106572
    
    lErro = IniciarLocal(objDet)
    If lErro <> SUCESSO Then gError 106571
    
    Reiniciar = SUCESSO
     
    Exit Function
    
Erro_Reiniciar:

    Reiniciar = gErr
     
    Select Case gErr
          
        Case 106571, 106572
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151535)
     
    End Select
     
    Exit Function

End Function

Public Function Finalizar() As Long

Dim lErro As Long

On Error GoTo Erro_Finalizar

    lErro = objArq.FinalizarLote(Me)
    If lErro <> SUCESSO Then gError 106572
    
    Set mvarobjArq = Nothing
        
    Finalizar = SUCESSO
     
    Exit Function
    
Erro_Finalizar:

    Finalizar = gErr
     
    Select Case gErr
          
        Case 106572
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151536)
     
    End Select
     
    Exit Function

End Function

Public Property Get lSequencial() As Long

    lSequencial = objArq.lSequencialLote
    
End Property

Public Sub AcumulaValor(ByVal dValor As Double)

    mvardValorTotal = mvardValorTotal + dValor
    
    Call objArq.AcumulaValor(dValor)
    
End Sub

Public Function ObterProxReg() As Long
'incrementa contador de registros no lote
'cada segmento de um detalhe é contado separadamente

    mvarlQuantRegistros = mvarlQuantRegistros + 1
    
    Call objArq.ObterProxReg
    
    ObterProxReg = mvarlQuantRegistros
    
End Function


