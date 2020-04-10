VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelUltVendasPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MAX_LINHAS_ULTVENDASPV = 100

Private alComando(0 To 3) As Long

Private lUltNumIntPV As Long

Private iNumLinhas As Integer
Private asLinhas(0 To MAX_LINHAS_ULTVENDASPV) As String

Function UltVendasPV_ObtemLinha(sRetorno As String, ByVal iLinha As Integer, ByVal lNumIntPV As Long) As Long

Dim lErro As Long

On Error GoTo Erro_UltVendasPV_ObtemLinha

    If lNumIntPV <> lUltNumIntPV Then
    
        lErro = UltVendasPV_ObtemLinhas(lNumIntPV)
        If lErro <> SUCESSO Then gError 130784
        
    End If
    
    If iLinha <= iNumLinhas And iLinha > 0 Then
        sRetorno = asLinhas(iLinha)
    Else
        sRetorno = ""
    End If
    
    UltVendasPV_ObtemLinha = SUCESSO
     
    Exit Function
    
Erro_UltVendasPV_ObtemLinha:

    UltVendasPV_ObtemLinha = gErr
     
    Select Case gErr
          
        Case 130784
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Function UltVendasPV_ObtemLinhas(ByVal lNumIntPV As Long) As Long

Dim lErro As Long, lErro2 As Long, iIndice As Integer, sProduto As String
Dim iNumLinhas2, asLinhas2(0 To MAX_LINHAS_ULTVENDASPV) As String, iNFsProd As Integer, dtDataEmissaoPV As Date
Dim iFilialEmpresa As Integer, iFilialEmpresaPV As Integer, lCodigoPV As Long, lCliente As Long, iFilialCli As Integer
Dim dtDataEmissaoNF As Date, lNumNotaFiscal As Long, dValorTotalNF As Double, dPrecoUnitario As Double, dQuantidade As Double, dPrecoDesde As Date, dPrecoUnitInicial As Date

On Error GoTo Erro_UltVendasPV_ObtemLinhas

    If alComando(UBound(alComando)) = 0 Then
    
        For iIndice = LBound(alComando) To UBound(alComando)
            
            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
            If alComando(iIndice) = 0 Then gError 130784
            
        Next
        
    End If
    
    sProduto = String(STRING_PRODUTO, 0)
    lErro = Comando_Executar(alComando(0), "SELECT PedidosDeVenda.FilialEmpresaFaturamento, PedidosDeVenda.FilialEmpresa, PedidosDeVenda.Codigo, Cliente, Filial, Produto, DataEmissao FROM ItensPedidoDeVenda, PedidosDeVenda WHERE ItensPedidoDeVenda.FilialEmpresa = PedidosDeVenda.FilialEmpresa AND ItensPedidoDeVenda.CodPedido = PedidosDeVenda.Codigo AND PedidosDeVenda.NumIntDoc = ? ORDER BY ItensPedidoDeVenda.NumIntDoc", _
        iFilialEmpresa, iFilialEmpresaPV, lCodigoPV, lCliente, iFilialCli, sProduto, dtDataEmissaoPV, lNumIntPV)
    If lErro <> AD_SQL_SUCESSO Then gError 130785

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130786
    
    Do While lErro = AD_SQL_SUCESSO
    
        iNFsProd = 0
        dPrecoDesde = DATA_NULA
        dPrecoUnitInicial = 0
        
        lErro2 = Comando_Executar(alComando(1), "SELECT NFiscal.DataEmissao, NFiscal.NumNotaFiscal, NFiscal.ValorTotal, ItensNFiscal.PrecoUnitario, ItensNFiscal.Quantidade FROM NFiscal, ItensNFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.Status <> 7 AND TiposDocInfo.Faturamento=1 AND NFiscal.FilialEmpresa = ? AND NFiscal.Cliente = ? AND NFiscal.FilialCli = ? AND ItensNFiscal.Produto = ? AND NFiscal.DataEmissao < ? AND (NFiscal.NumPedidoVenda <> ? OR NFiscal.FilialPedido <> ?) ORDER BY DataEmissao DESC", _
            dtDataEmissaoNF, lNumNotaFiscal, dValorTotalNF, dPrecoUnitario, dQuantidade, iFilialEmpresa, lCliente, iFilialCli, sProduto, dtDataEmissaoPV, lCodigoPV, iFilialEmpresaPV)
        If lErro2 <> AD_SQL_SUCESSO Then gError 130787
        
        lErro2 = Comando_BuscarProximo(alComando(1))
        If lErro2 <> AD_SQL_SUCESSO And lErro2 <> AD_SQL_SEM_DADOS Then gError 130788
    
        Do While lErro2 = AD_SQL_SUCESSO
        
            If iNFsProd = 2 Then
                
                If Abs(dPrecoUnitario - dPrecoUnitInicial) > DELTA_VALORMONETARIO Then
                    asLinhas2(iNumLinhas2) = asLinhas2(iNumLinhas2) & " " & Format(dPrecoDesde, "dd/mm/yy")
                    dPrecoDesde = DATA_NULA
                    Exit Do
                Else
                    dPrecoDesde = dtDataEmissaoNF
                End If
                            
            Else
            
                'montar linha
                iNumLinhas2 = iNumLinhas2 + 1
                asLinhas2(iNumLinhas2) = sProduto & " " & Format(dtDataEmissaoNF, "dd/mm/yy") & " " & FormataCpoTexto(CStr(lNumNotaFiscal), 6) & " " & FormataCpoTexto(Format(dPrecoUnitario, "standard"), 8) & " " & FormataCpoTexto(Format(dQuantidade, "standard"), 8)
                
                'limitar a 2 nfs por produto
                iNFsProd = iNFsProd + 1
                
                If iNFsProd = 2 Then
                    dPrecoUnitInicial = dPrecoUnitario
                    dPrecoDesde = dtDataEmissaoNF
                End If
            
            End If
            
            lErro2 = Comando_BuscarProximo(alComando(1))
            If lErro2 <> AD_SQL_SUCESSO And lErro2 <> AD_SQL_SEM_DADOS Then gError 130789
        
        Loop
        
        If iNFsProd = 2 And dPrecoDesde <> DATA_NULA Then
                
            asLinhas2(iNumLinhas2) = asLinhas2(iNumLinhas2) & " " & Format(dPrecoDesde, "dd/mm/yy")
        
        End If
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130790
    
    Loop

    iNumLinhas = iNumLinhas2
    For iIndice = 1 To iNumLinhas2
        asLinhas(iIndice) = asLinhas2(iIndice)
    Next
    
    UltVendasPV_ObtemLinhas = SUCESSO
     
    Exit Function
    
Erro_UltVendasPV_ObtemLinhas:

    UltVendasPV_ObtemLinhas = gErr
     
    Select Case gErr
          
        Case 130784
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130785 To 130790
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ULTVENDASPV", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Sub Class_Terminate()

Dim iIndice As Integer

    For iIndice = LBound(alComando) To UBound(alComando)
        If alComando(iIndice) <> 0 Then Call Comando_Fechar(alComando(iIndice))
    Next
    
End Sub
