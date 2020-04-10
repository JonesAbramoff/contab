VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelTitRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Criada para atender emissão de Notas Fiscais

Private lNumIntUltimoTitulo As Long
Private colParcRec As New ColParcelaReceber
Private lComando As Long

Public Function ObterDataParcRec(dtDataParcRec As Date, lNumIntTitulo As Long, iParcela As Integer) As Long
'para obter a data de vencimento da parcela a receber

Dim lErro As Long

On Error GoTo Erro_ObterDataParcRec

    If lNumIntTitulo <> lNumIntUltimoTitulo Then
        
        Set colParcRec = New ColParcelaReceber
        
        If lNumIntTitulo <> 0 Then
        
            lErro = CF("ParcelasRec_Le_DataValor", lNumIntTitulo, colParcRec, lComando)
            If lErro <> SUCESSO Then Error 58223
        
        End If
        
        lNumIntUltimoTitulo = lNumIntTitulo
        
    End If
    
    If iParcela <= colParcRec.Count Then
        dtDataParcRec = colParcRec(iParcela).dtDataVencimento
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153377)

    End Select

    Exit Function

End Function

Public Function ObterValorParcRec(dValor As Double, lNumIntTitulo As Long, iParcela As Integer) As Long
'para obter o valor da parcela a receber

Dim lErro As Long

On Error GoTo Erro_ObterValorParcRec

    If lNumIntTitulo <> lNumIntUltimoTitulo Then
        
        Set colParcRec = New ColParcelaReceber
        
        If lNumIntTitulo <> 0 Then
        
            lErro = CF("ParcelasRec_Le_DataValor", lNumIntTitulo, colParcRec, lComando)
            If lErro <> SUCESSO Then Error 58224
            
        End If
        
        lNumIntUltimoTitulo = lNumIntTitulo
        
    End If
    
    If iParcela <= colParcRec.Count Then
        dValor = colParcRec(iParcela).dValor 'Original
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153378)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub

Public Function ObterSaldoParcRec(dSaldo As Double, lNumIntTitulo As Long, iParcela As Integer) As Long
'para obter o valor da parcela a receber

Dim lErro As Long

On Error GoTo Erro_ObterSaldoParcRec

    If lNumIntTitulo <> lNumIntUltimoTitulo Then
        
        Set colParcRec = New ColParcelaReceber
        
        If lNumIntTitulo <> 0 Then
        
            lErro = CF("ParcelasRec_Le_DataValor", lNumIntTitulo, colParcRec, lComando)
            If lErro <> SUCESSO Then Error 58224
            
        End If
        
        lNumIntUltimoTitulo = lNumIntTitulo
        
    End If
    
    If iParcela <= colParcRec.Count Then
        dSaldo = colParcRec(iParcela).dSaldo
    Else
        dSaldo = 0
    End If
        
    ObterSaldoParcRec = SUCESSO
    
    Exit Function
    
Erro_ObterSaldoParcRec:

    ObterSaldoParcRec = Err
    
    Select Case Err

        Case 58224 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153378)

    End Select

    Exit Function

End Function

