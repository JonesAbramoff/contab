VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotBxCartCielo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'RO = Resumo de Operacao
'CV = Comprovante de Venda

Implements ClassRotBxCartRede

Private mvarobjAux As New ClassRotBxCartRedeAux

Private objArqRO As ClassArqCieloRO

Private lNumRegArq As Long

Private iTipoRegAnterior As Integer

Private Function ClassRotBxCartRede_ImportarArqFim(alComando() As Long) As Long

Dim lErro As Long, objMov As ClassAdmExtFinMov, objRO As ClassArqCieloRO, objCV As ClassArqCieloCV
Dim lNumIntMov As Long, lNumIntRO As Long, lNumIntMovDet As Long, objMovDet As ClassAdmExtFinMovDet

On Error GoTo Erro_ClassRotBxCartRede_ImportarArqFim

    'fazer loop pelos movtos
    For Each objMov In mvarobjAux.colMov
    
        'obter um numint para o registro
        lErro = CF("Config_ObterNumInt", "CRConfig", "NUM_PROX_ROTBXCARTAO", lNumIntMov)
        If lErro <> SUCESSO Then gError 201076
        
        objMov.lNumIntDoc = lNumIntMov
        
        'inserir registro em AdmExtFinMov
        With objMov
            lErro = Comando_Executar(alComando(5), "INSERT INTO AdmExtFinMov (NumIntDoc, NumIntArq, Estabelecimento, Tipo, CodConta, Data, ValorBruto, ValorComissao, ValorRejeitado, ValorLiq, FilialEmpresa, NumMovto, Bandeira) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                .lNumIntDoc, mvarobjAux.lNumIntArq, .sEstabelecimento, ADMEXTFIN_TIPOMOV_ROBXCARTAO, .iCodConta, .dtData, Arredonda_Moeda(.dValorBruto / 100), Arredonda_Moeda(.dValorComissao / 100), Arredonda_Moeda(.dValorRejeitado / 100), Arredonda_Moeda(.dValorLiq / 100), .iFilialEmpresa, 0, .iBandeira)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 201077
        
        'loop pelos resumos de operacao
        For Each objRO In objMov.colRO
        
            'obter um numint para o registro
            lErro = CF("Config_ObterNumInt", "CRConfig", "NUM_PROX_ROTBXCARTAO", lNumIntRO)
            If lErro <> SUCESSO Then gError 201076

            objRO.lNumIntDoc = lNumIntRO

            'inserir registro em AdmExtFinCieloRO
            With objRO
                lErro = Comando_Executar(alComando(6), "INSERT INTO AdmExtFinCieloRO (NumIntDoc, NumIntArq, NumRegArq, Estabelecimento, RO, Banco, Agencia, Conta, Produto, Status, DtDeposito, DtPrevPag, DtEnvBco, ValorBruto, ValorComissao, ValorRejeitado, ValorLiq, NumIntMov) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    .lNumIntDoc, mvarobjAux.lNumIntArq, .lNumRegArq, .sEstabelecimento, .sRO, .iBanco, .iAgencia, .sConta, .sProduto, .sStatus, .dtDeposito, .dtPrevPag, .dtEnvBco, .dValorBruto, .dValorComissao, .dValorRejeitado, .dValorLiq, objMov.lNumIntDoc)
            End With
            If lErro <> AD_SQL_SUCESSO Then gError 201078

            'loop pelos comprovantes de vendas
            For Each objCV In objRO.colCV
            
                'obter um numint para o registro
                lErro = CF("Config_ObterNumInt", "CRConfig", "NUM_PROX_ROTBXCARTAO", lNumIntMovDet)
                If lErro <> SUCESSO Then gError 201079

                Set objMovDet = New ClassAdmExtFinMovDet
                Call objCV.PreencherMovDet(objMovDet)

                With objMovDet
                    .lNumIntDoc = lNumIntMovDet
                    .lNumIntMov = objMov.lNumIntDoc
                End With

                'inserir registro em AdmExtFinMovDet
                With objMovDet
                    lErro = Comando_Executar(alComando(7), "INSERT INTO AdmExtFinMovDet (NumIntDoc, NumIntMov, Tipo, Valor, NumCartao, DataCompra, NumParcela, Autorizacao, RO, NSU, Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                        .lNumIntDoc, .lNumIntMov, .iTipo, .dValor, .sNumCartao, .dtDataCompra, .iNumParcela, .sAutorizacao, .sRO, .sNSU, "")
                End With
                If lErro <> AD_SQL_SUCESSO Then gError 201080

                'inserir registro em AdmExtFinCieloCV
                With objCV
                    lErro = Comando_Executar(alComando(8), "INSERT INTO AdmExtFinCieloCV (NumIntDoc, NumIntArq, NumRegArq, NumIntRO, TID, TotalParcelas, MotivoRejeicao) VALUES (?,?,?,?,?,?,?)", _
                        objMovDet.lNumIntDoc, mvarobjAux.lNumIntArq, .lNumRegArq, objRO.lNumIntDoc, .sTID, .iTotalParcelas, .sMotivoRejeicao)
                End With
                If lErro <> AD_SQL_SUCESSO Then gError 201081

            Next
        
        Next
        
    Next
    
    ClassRotBxCartRede_ImportarArqFim = SUCESSO
    
    Exit Function
    
Erro_ClassRotBxCartRede_ImportarArqFim:

    ClassRotBxCartRede_ImportarArqFim = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201100)

    End Select
    
    Exit Function

End Function

Private Function ClassRotBxCartRede_ImportarArqInicio(alComando() As Long) As Long

    lNumRegArq = 0
    iTipoRegAnterior = -1
    Set objArqRO = Nothing
    
    ClassRotBxCartRede_ImportarArqInicio = SUCESSO

End Function

Private Function ClassRotBxCartRede_ImportarArqTrataRegistro(ByVal sRegistro As String, alComando() As Long) As Long

Dim lErro As Long, lSeq As Long, dtProc As Date, sEstabelecimento As String
Dim sRO As String, sROStatus As String, sTipoReg As String, iCodConta As Integer
Dim objCV As ClassArqCieloCV, iFilEmpCCI As Integer
Dim dtProcAnterior As Date, objRO As ClassArqCieloRO
Dim objMov As ClassAdmExtFinMov, bAchou As Boolean

On Error GoTo Erro_ClassRotBxCartRede_ImportarArqTrataRegistro

    lNumRegArq = lNumRegArq + 1
    
    sTipoReg = left(sRegistro, 1)
    
    'se é o 1o registro
    If iTipoRegAnterior = -1 Then
        
        If sTipoReg <> "0" Then gError 206886
        
        If UCase(Mid(sRegistro, 43, 5)) <> "CIELO" Then gError 206887
        
        sEstabelecimento = Mid(sRegistro, 2, 10)
        lSeq = StrParaLong(Mid(sRegistro, 36, 7))
        dtProc = AAAAMMDD_ParaDate(Mid(sRegistro, 12, 8))
        
        'Pagamentos com CV
        If Mid(sRegistro, 48, 2) = "04" Then
            
            mvarobjAux.bPularArquivo = False
        
            'verificar se o arquivo com este sequencial já foi importado
            lErro = Comando_Executar(alComando(3), "SELECT V.DataProcessamento FROM AdmExtFinCieloArq As V, AdmExtFinArqsLidos As A WHERE A.NumIntDoc = V.NumIntArq AND A.NaoAtualizar = 0 AND V.Estabelecimento = ? AND V.Seq = ?", _
                dtProcAnterior, sEstabelecimento, lSeq)
            If lErro <> AD_SQL_SUCESSO Then gError 206888
            
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206889
            If lErro <> AD_SQL_SEM_DADOS Then gError 206890
            
            'inserir registro em AdmExtFinCieloArq
            lErro = Comando_Executar(alComando(4), "INSERT INTO AdmExtFinCieloArq (NumIntArq, Estabelecimento, DataProcessamento, Seq) VALUES (?,?,?,?)", _
                mvarobjAux.lNumIntArq, sEstabelecimento, dtProc, lSeq)
            If lErro <> AD_SQL_SUCESSO Then gError 206891
        
        Else
        
            mvarobjAux.bPularArquivo = True
        
        End If
        
        iTipoRegAnterior = 0
    
    Else
    
        Select Case sTipoReg
        
            Case "1" 'RO = Resumo de Operacao
                
                Set objRO = New ClassArqCieloRO
                objRO.sReg = sRegistro
                
                'se nao é credito a vista/parcelado
                If objRO.sProduto = " " Or StrParaLong(objRO.sConta) = 0 Then
                    Set objArqRO = Nothing
                Else
                    Set objArqRO = objRO
                    
                    'verificar bco/agencia e conta
                    lErro = Comando_Executar(alComando(9), "SELECT C.Codigo, C.FilialEmpresa FROM ContasCorrentesInternas C, Redes R WHERE CodBanco = ? AND CONVERT(SMALLINT,Agencia) = ? AND (NumConta = ? OR (NumConta + DVNumConta = ?)) AND R.Estabelecimento = ? AND C.FilialEmpresa = R.FilialEmpresa", _
                        iCodConta, iFilEmpCCI, objRO.iBanco, objRO.iAgencia, CStr(StrParaLong(objRO.sConta)), CStr(StrParaLong(objRO.sConta)), objRO.sEstabelecimento)
                    If lErro <> AD_SQL_SUCESSO Then gError 206892
                    
                    lErro = Comando_BuscarProximo(alComando(9))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206893
                    If lErro <> AD_SQL_SUCESSO Then gError 206894
                    
                    objRO.iCodConta = iCodConta
                    objRO.lNumRegArq = lNumRegArq
                    
                    bAchou = False
                    For Each objMov In mvarobjAux.colMov
                        If objMov.dtData = objRO.dtPrevPag And objMov.iCodConta = objRO.iCodConta And objMov.sEstabelecimento = objRO.sEstabelecimento And objMov.iBandeira = objRO.iBandeira Then
                        
                            bAchou = True
                            Exit For
                            
                        End If
                    Next
                    
                    If bAchou = False Then
                    
                        Set objMov = New ClassAdmExtFinMov
                        objMov.iCodConta = iCodConta
                        objMov.iFilialEmpresa = iFilEmpCCI
                        objMov.dtData = objRO.dtPrevPag
                        objMov.sEstabelecimento = objRO.sEstabelecimento
                        objMov.iBandeira = objRO.iBandeira
                        
                        mvarobjAux.colMov.Add objMov
                        
                    End If
                    
                    objMov.dValorBruto = objMov.dValorBruto + objRO.dValorBruto
                    objMov.dValorComissao = objMov.dValorComissao + objRO.dValorComissao
                    objMov.dValorRejeitado = objMov.dValorRejeitado + objRO.dValorRejeitado
                    objMov.dValorLiq = objMov.dValorLiq + objRO.dValorLiq
                            
                    objMov.colRO.Add objRO
                       
                End If
                
            Case "2" 'CV = comprovante de venda
            
                If Not (objArqRO Is Nothing) Then
                
                    Set objCV = New ClassArqCieloCV
                    objCV.sReg = sRegistro
                
                    objCV.lNumRegArq = lNumRegArq
                    
'                    'só adicionar em objArqRO.colCV se nao houver motivo de rejeicao e se o valor for >0
'                    If Len(Trim(objCV.sMotivoRejeicao)) = 0 And objCV.dValorParcela > DELTA_VALORMONETARIO Then
                        objArqRO.colCV.Add objCV
'                    End If
                    
                End If
            
            Case "9" 'trailer de arquivo
                If (lNumRegArq - 2) <> StrParaInt(Mid(sRegistro, 2, 11)) Then gError 206895
            
            Case Else
                gError 201104
                
        End Select
    
    End If
    
    ClassRotBxCartRede_ImportarArqTrataRegistro = SUCESSO
    
    Exit Function
    
Erro_ClassRotBxCartRede_ImportarArqTrataRegistro:

    ClassRotBxCartRede_ImportarArqTrataRegistro = gErr

    Select Case gErr
    
        Case 206886
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_IMPORT_EXT_TIPOREG_ERRADO", gErr, sTipoReg, "0")
    
        Case 206887
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_IMPORT_EXT_REDE_ERRADA", gErr, Mid(sRegistro, 46, 7), "Cielo")
    
        Case 206888, 206889
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMEXTFINVISANETARQ", gErr)
            
        Case 206890
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_ARQ_JA_IMPORTADO", gErr, lSeq, sEstabelecimento, dtProcAnterior)
        
        Case 206891
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_ADMEXTFINVISANETARQ", gErr)
            
        Case 206892, 206893
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTASCORRENTESINTERNAS", gErr)
    
        Case 206894
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_CTACORRENTE_NAO_CADASTRADO", gErr, objRO.iBanco, objRO.iAgencia, CStr(CLng(objRO.sConta)), objRO.sEstabelecimento)

        Case 206895
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_IMPORT_EXT_TRAILER_ERRADA", gErr, CStr(lNumRegArq - 2), Mid(sRegistro, 2, 11))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206896)

    End Select
    
    Exit Function
    
End Function

Private Property Get ClassRotBxCartRede_objAux() As ClassRotBxCartRedeAux

    Set ClassRotBxCartRede_objAux = mvarobjAux

End Property
