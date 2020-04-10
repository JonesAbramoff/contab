VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private alComando(1 To 3) As Long

Private lUltNumIntDoc As Long
Private gcolNF As Collection

Public Function Fatura_ObterNFNumero(lNumNotaFiscal As Long, lNumIntDoc As Long, iNumNF As Integer) As Long
'retorna o valor da enesima nota fiscal cobrada pela fatura identificada por lNumIntDoc da tabela de titulosrec (n = iNumNF)
'a posicao das notas fiscais dentro da fatura é dada pela ordem crescente de numeracao

Dim lErro As Long, objNFiscal As ClassNFiscal

On Error GoTo Erro_Fatura_ObterNFNumero

    If lNumIntDoc <> lUltNumIntDoc Then
    
        lErro = RelFatura_ObterNFs(alComando, lNumIntDoc, gcolNF)
        If lErro <> SUCESSO Then gError 81821
    
        lUltNumIntDoc = lNumIntDoc
        
    End If
    
    If iNumNF >= 1 And iNumNF <= gcolNF.Count Then
    
        lNumNotaFiscal = gcolNF(iNumNF).lNumNotaFiscal
            
    Else
        
        lNumNotaFiscal = 0
            
    End If
    
    Fatura_ObterNFNumero = SUCESSO
     
    Exit Function
    
Erro_Fatura_ObterNFNumero:

    Fatura_ObterNFNumero = gErr
     
    Select Case gErr
          
        Case 81821
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153119)
     
    End Select
     
    Exit Function

End Function

Public Function Fatura_ObterNFValor(dValorNF As Double, lNumIntDoc As Long, iNumNF As Integer) As Long
'retorna o valor da enesima nota fiscal cobrada pela fatura identificada por lNumIntDoc da tabela de titulosrec (n = iNumNF)
'a posicao das notas fiscais dentro da fatura é dada pela ordem crescente de numeracao

Dim lErro As Long, objNFiscal As ClassNFiscal

On Error GoTo Erro_Fatura_ObterNFValor

    If lNumIntDoc <> lUltNumIntDoc Then
    
        lErro = RelFatura_ObterNFs(alComando, lNumIntDoc, gcolNF)
        If lErro <> SUCESSO Then gError 81822
    
        lUltNumIntDoc = lNumIntDoc
        
    End If
    
    If iNumNF >= 1 And iNumNF <= gcolNF.Count Then
    
        Set objNFiscal = gcolNF(iNumNF)
        
        If objNFiscal.iTipoNFiscal <> TIPODOCINFO_CONHECIMENTOFRETE_FATURA And objNFiscal.iTipoNFiscal <> TIPODOCINFO_CONHECIMENTOFRETE Then
            dValorNF = Round(objNFiscal.dValorTotal - objNFiscal.objTributacaoNF.dIRRFValor - IIf(objNFiscal.objTributacaoNF.iINSSRetido <> 0, objNFiscal.objTributacaoNF.dValorINSS, 0), 2)
        Else
            dValorNF = Round(objNFiscal.dValorTotal - IIf(objNFiscal.objConhecimentoFrete.iINSSRetido <> 0, objNFiscal.objConhecimentoFrete.dValorINSS, 0), 2)
        End If
            
    Else
        
        dValorNF = 0
            
    End If
    
    Fatura_ObterNFValor = SUCESSO
     
    Exit Function
    
Erro_Fatura_ObterNFValor:

    Fatura_ObterNFValor = gErr
     
    Select Case gErr
          
        Case 81822
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153120)
     
    End Select
     
    Exit Function

End Function

Private Sub Class_Initialize()
    lUltNumIntDoc = -1
End Sub

Private Sub Class_Terminate()
Dim iIndice As Integer

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
End Sub

Public Function RelFatura_ObterNFs(alComando() As Long, lNumIntDoc As Long, colNF As Collection) As Long
'Preenche colNF com informacoes das notas fiscais associadas à fatura a receber identificada por lNumIntDoc, por ordem de numero de nota fiscal
'Obs.: uma solucao + otimizada pode utilizar uma view de nfisca com outer joins p/conhec de frete e tributacaonf

Dim lErro As Long, iIndice As Integer, iTipoNFiscal As Integer
Dim lNumIntNF As Long, lNumNotaFiscal As Long, dtDataEmissao As Date, dValorTotal As Double
Dim iINSSRetido As Integer, dValorINSS As Double, dIRRFValor As Double
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_RelFatura_ObterNFs
    
    Set colNF = New Collection
    
    For iIndice = LBound(alComando) To UBound(alComando)
        
        If alComando(iIndice) = 0 Then
    
            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
            If alComando(iIndice) = 0 Then gError 81811
    
        End If
        
    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, NumNotaFiscal, DataEmissao, ValorTotal, TipoNFiscal FROM NFiscal WHERE NumIntDocCPR = ? AND ClasseDocCPR = ? ORDER BY NumNotaFiscal", lNumIntNF, lNumNotaFiscal, dtDataEmissao, dValorTotal, iTipoNFiscal, lNumIntDoc, CPR_TITULO_RECEBER)
    If lErro <> AD_SQL_SUCESSO Then gError 81812
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81813
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objNFiscal = New ClassNFiscal
        
        With objNFiscal
        
            .lNumIntDoc = lNumIntNF
            .lNumNotaFiscal = lNumNotaFiscal
            .dtDataEmissao = dtDataEmissao
            .dValorTotal = dValorTotal
            .iTipoNFiscal = iTipoNFiscal
            
        End With
        
        Select Case iTipoNFiscal
        
            Case TIPODOCINFO_CONHECIMENTOFRETE, TIPODOCINFO_CONHECIMENTOFRETE_FATURA
            
                lErro = Comando_Executar(alComando(2), "SELECT INSSRetido, ValorINSS FROM ConhecimentoFrete WHERE NumIntNFiscal = ?", iINSSRetido, dValorINSS, lNumIntNF)
                If lErro <> AD_SQL_SUCESSO Then gError 81815
                
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81816
                If lErro <> AD_SQL_SUCESSO Then gError 81817
            
                Set objNFiscal.objConhecimentoFrete = New ClassConhecimentoFrete
                
                objNFiscal.objConhecimentoFrete.iINSSRetido = iINSSRetido
                objNFiscal.objConhecimentoFrete.dValorINSS = dValorINSS
                
            Case Else
        
                lErro = Comando_Executar(alComando(3), "SELECT INSSRetido, ValorINSS, IRRFValor FROM TributacaoNF WHERE NumIntDoc = ?", iINSSRetido, dValorINSS, dIRRFValor, lNumIntNF)
                If lErro <> AD_SQL_SUCESSO Then gError 81818
                
                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81819
                If lErro <> AD_SQL_SUCESSO Then gError 81820
        
                Set objNFiscal.objTributacaoNF = New ClassTributacaoDoc
                
                objNFiscal.objTributacaoNF.iINSSRetido = iINSSRetido
                objNFiscal.objTributacaoNF.dValorINSS = dValorINSS
                objNFiscal.objTributacaoNF.dIRRFValor = dIRRFValor
            
        End Select
    
        Call colNF.Add(objNFiscal)
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81814
    
    Loop
    
    RelFatura_ObterNFs = SUCESSO
     
    Exit Function
    
Erro_RelFatura_ObterNFs:

    RelFatura_ObterNFs = gErr
     
    Select Case gErr
          
        Case 81811
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error)
            
        Case 81812, 81813, 81814
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr, Error)
        
        Case 81815, 81816, 81817
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONHECIMENTO_FRETE", gErr, Error)
        
        Case 81818, 81819, 81820
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRIBNFISCAL", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153121)
     
    End Select
     
    Exit Function
    
End Function
