VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelPedVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Criada para atender a pedidos de vendas

'para obter dados das parcelas do pedido de vendas
Private lComandoParc As Long
Private lCodigoAntParc As Long
Private iFilialEmpresaAntParc As Long
Private colParc As ColParcelaPedidoVenda

'para obter a qtde de bloqueios ativos
Dim lComandoBloq As Long
Dim lCodigoAntBloq As Long
Dim iFilialEmpresaAntBloq As Long
Dim iBloqueiosAnt As Long

Private Sub Class_Terminate()
    If lComandoBloq <> 0 Then Call Comando_Fechar(lComandoBloq)
    If lComandoParc <> 0 Then Call Comando_Fechar(lComandoParc)
End Sub

Public Function Obter_Bloqueios_PV(iBloqueios As Integer, ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer) As Long
'Le a quantidade de bloqueios não liberados

Dim lErro As Long

On Error GoTo Erro_Obter_Bloqueios_PV

    If lCodigo <> lCodigoAntBloq Or iFilialEmpresa <> iFilialEmpresaAntBloq Then
        
        lErro = CF("PV_Verifica_Bloqueios", lCodigo, iFilialEmpresa, iBloqueios, lComandoBloq)
        If lErro <> SUCESSO Then gError 141740
        
        lCodigoAntBloq = lCodigo
        iFilialEmpresaAntBloq = iFilialEmpresa
        iBloqueiosAnt = iBloqueios
        
    Else
        
        iBloqueios = iBloqueiosAnt
    
    End If
        
    Obter_Bloqueios_PV = SUCESSO
    
    Exit Function
    
Erro_Obter_Bloqueios_PV:
        
    Obter_Bloqueios_PV = Err
    
    Select Case Err
        
        Case 141740 'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 130463)
    
    End Select

    Exit Function
    
End Function

Public Function ObterDataParcRec(dtDataParcRec As Date, ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer, iParcela As Integer) As Long
'para obter a data de vencimento da parcela a receber

Dim lErro As Long

On Error GoTo Erro_ObterDataParcRec

    If lCodigo <> lCodigoAntParc Or iFilialEmpresa <> iFilialEmpresaAntParc Then
        
        Set colParc = New ColParcelaPedidoVenda
        
        lErro = CF("ParcelasPV_Le_Rel", iFilialEmpresa, lCodigo, colParc, lComandoParc)
        If lErro <> SUCESSO Then Error 58223
        
        lCodigoAntParc = lCodigo
        iFilialEmpresaAntParc = iFilialEmpresa
        
    End If
    
    If iParcela <= colParc.Count Then
        dtDataParcRec = colParc(iParcela).dtDataVencimento
    Else
        dtDataParcRec = DATA_NULA
    End If
    
    ObterDataParcRec = SUCESSO
    
    Exit Function
    
Erro_ObterDataParcRec:

    ObterDataParcRec = Err
    
    Select Case Err

        Case 58223 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 130466)

    End Select

    Exit Function

End Function

Public Function ObterValorParcRec(dValor As Double, ByVal lCodigo As Long, ByVal iFilialEmpresa As Integer, iParcela As Integer) As Long
'para obter o valor da parcela a receber

Dim lErro As Long

On Error GoTo Erro_ObterValorParcRec

    If lCodigo <> lCodigoAntParc Or iFilialEmpresa <> iFilialEmpresaAntParc Then
        
        Set colParc = New ColParcelaPedidoVenda
        
        lErro = CF("ParcelasPV_Le_Rel", iFilialEmpresa, lCodigo, colParc, lComandoParc)
        If lErro <> SUCESSO Then Error 58224
        
        lCodigoAntParc = lCodigo
        iFilialEmpresaAntParc = iFilialEmpresa
        
    End If
    
    If iParcela <= colParc.Count Then
        dValor = colParc(iParcela).dValor
    Else
        dValor = 0
    End If
        
    ObterValorParcRec = SUCESSO
    
    Exit Function
    
Erro_ObterValorParcRec:

    ObterValorParcRec = Err
    
    Select Case Err

        Case 58224 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 130467)

    End Select

    Exit Function

End Function

