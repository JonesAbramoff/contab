VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotBxCartAmex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'SOC = Resumo de Vendas
'ROC = Registro de Vendas

Implements ClassRotBxCartRede

Private mvarobjAux As New ClassRotBxCartRedeAux

Private colPagtos As New Collection

Private lNumRegArq As Long

Private Function ClassRotBxCartRede_ImportarArqFim(alComando() As Long) As Long

Dim lErro As Long, objMov As ClassAdmExtFinMov
Dim lNumIntMov As Long, lNumIntSOC As Long, lNumIntMovDet As Long, objMovDet As ClassAdmExtFinMovDet
Dim objPgto As ClassArqAmexPgto, objSOC As ClassArqAmexSOC, objROC As ClassArqAmexROC

On Error GoTo Erro_ClassRotBxCartRede_ImportarArqFim

    'fazer loop pelos pagtos
    For Each objPgto In colPagtos

        If objPgto.iPagtoFechado <> 0 Then
        
            'obter um numint para o registro
            lErro = CF("Config_ObterNumInt", "CRConfig", "NUM_PROX_ROTBXCARTAO", lNumIntMov)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            Set objMov = New ClassAdmExtFinMov
            With objMov
                .lNumIntDoc = lNumIntMov
                .sEstabelecimento = objPgto.sEstabelecimento
                .iCodConta = objPgto.iCodConta
                .dtData = objPgto.dtDataCredito
                .dValorLiq = objPgto.dValorCredito
                .iFilialEmpresa = objPgto.iFilEmpCCI
                .dValorBruto = objPgto.dValorBruto
                .dValorComissao = objPgto.dValorComissao
            End With
    
            'inserir registro em AdmExtFinMov
            With objMov
                lErro = Comando_Executar(alComando(5), "INSERT INTO AdmExtFinMov (NumIntDoc, NumIntArq, Estabelecimento, Tipo, CodConta, Data, ValorBruto, ValorComissao, ValorRejeitado, ValorLiq, FilialEmpresa, NumMovto) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
                    .lNumIntDoc, mvarobjAux.lNumIntArq, .sEstabelecimento, ADMEXTFIN_TIPOMOV_ROBXCARTAO, .iCodConta, .dtData, Arredonda_Moeda(.dValorBruto / 100), Arredonda_Moeda(.dValorComissao / 100), Arredonda_Moeda(.dValorRejeitado / 100), Arredonda_Moeda(.dValorLiq / 100), .iFilialEmpresa, 0)
            End With
            If lErro <> AD_SQL_SUCESSO Then gError 216146
    
            With objPgto
                lErro = Comando_Executar(alComando(5), "INSERT INTO AdmExtFinAmexPgto (NumIntDoc, NumRegArq, Banco, Agencia, Conta, SeqPgto) VALUES (?,?,?,?,?,?)", _
                    objMov.lNumIntDoc, .lNumRegArq, .iBanco, .sAgencia, .sConta, .sSeqPgto)
            End With
            If lErro <> AD_SQL_SUCESSO Then gError 216147
            
            'loop pelos resumos de operacao
            For Each objSOC In objPgto.colSOC
    
                'obter um numint para o registro
                lErro = CF("Config_ObterNumInt", "CRConfig", "NUM_PROX_ROTBXCARTAO", lNumIntSOC)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                objSOC.lNumIntDoc = lNumIntSOC
    
                'inserir registro em AdmExtFinAmexSOC
                With objSOC
                    lErro = Comando_Executar(alComando(6), "INSERT INTO AdmExtFinAmexSOC (NumIntDoc, NumIntArq, NumRegArq, NumSOC, DataPagto, DataVenda, ValorLiqSemAntec, DataPagto24hs, ValorLiqComAntec, NumIntMov) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                        .lNumIntDoc, mvarobjAux.lNumIntArq, .lNumRegArq, .sNumSOC, .dtDataPagto, .dtDataVenda, .dValorLiqSemAntec, .dtDataPagto24hs, .dValorLiqComAntec, objMov.lNumIntDoc)
                End With
                If lErro <> AD_SQL_SUCESSO Then gError 216148
    
                'loop pelos comprovantes de vendas
                For Each objROC In objSOC.colROC
    
                    'obter um numint para o registro
                    lErro = CF("Config_ObterNumInt", "CRConfig", "NUM_PROX_ROTBXCARTAO", lNumIntMovDet)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                    Set objMovDet = New ClassAdmExtFinMovDet
                    Call objROC.PreencherMovDet(objMovDet)
    
                    With objMovDet
                        .lNumIntDoc = lNumIntMovDet
                        .lNumIntMov = objMov.lNumIntDoc
                    End With
    
                    'inserir registro em AdmExtFinMovDet
                    With objMovDet
                        lErro = Comando_Executar(alComando(7), "INSERT INTO AdmExtFinMovDet (NumIntDoc, NumIntMov, Tipo, Valor, NumCartao, DataCompra, NumParcela, Autorizacao, RO, NSU, Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                            .lNumIntDoc, .lNumIntMov, .iTipo, .dValor, .sNumCartao, .dtDataCompra, .iNumParcela, .sAutorizacao, .sRO, .sNSU, "")
                    End With
                    If lErro <> AD_SQL_SUCESSO Then gError 216149
    
                    'inserir registro em AdmExtFinAmexROC
                    With objROC
                        lErro = Comando_Executar(alComando(8), "INSERT INTO AdmExtFinAmexROC (NumIntDoc, NumIntArq, NumRegArq, NumIntSOC, ValorTotal, ValorPrimcParc, ValorOutrasParc, QtdeParcelas, NSUInt, NumFatROC) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                            objMovDet.lNumIntDoc, mvarobjAux.lNumIntArq, .lNumRegArq, objSOC.lNumIntDoc, .dValorTotal, .dValorPrimParc, .dValorOutrasParc, .iQtdeParcelas, .sNSUInt, .sNumFatROC)
                    End With
                    If lErro <> AD_SQL_SUCESSO Then gError 216150
    
                Next
    
            Next

        End If
        
    Next

    ClassRotBxCartRede_ImportarArqFim = SUCESSO

    Exit Function

Erro_ClassRotBxCartRede_ImportarArqFim:

    ClassRotBxCartRede_ImportarArqFim = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 216146
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_ADMEXTFINMOV", gErr)

        Case 216147
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_ADMEXTFINAMEXPAGTO", gErr)

        Case 216148
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_ADMEXTFINAMEXSOC", gErr)

        Case 216149
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_ADMEXTFINMOVDET", gErr)

        Case 216150
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_ADMEXTFINAMEXROC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216151)

    End Select

    Exit Function

End Function

Private Function ClassRotBxCartRede_ImportarArqInicio(alComando() As Long) As Long

    lNumRegArq = 0

    ClassRotBxCartRede_ImportarArqInicio = SUCESSO

End Function

Private Function ClassRotBxCartRede_ImportarArqTrataRegistro(ByVal sRegistro As String, alComando() As Long) As Long

Dim lErro As Long, lNumIntArq As Long, dtProc As Date, sEstabelecimento As String
Dim sRO As String, sROStatus As String, sTipoReg As String, iCodConta As Integer
Dim iFilEmpCCI As Integer
Dim dtArq As Date, objHeader As New ClassArqAmexHeader
Dim objMov As ClassAdmExtFinMov, bAchou As Boolean, objTrailer As New ClassArqAmexTrailer
Dim objPgto As ClassArqAmexPgto, objSOC As ClassArqAmexSOC, objROC As ClassArqAmexROC
Dim objAjuste As ClassArqAmexAjuste, sNomeArq As String

On Error GoTo Erro_ClassRotBxCartRede_ImportarArqTrataRegistro

    sTipoReg = Mid(sRegistro, 45, 1)

    Select Case sTipoReg

        Case "0" 'header
            lNumRegArq = lNumRegArq + 1

            objHeader.sReg = sRegistro

            'verificar se o arquivo com esta data já foi importado
            lErro = Comando_Executar(alComando(3), "SELECT NumIntArq FROM AdmExtFinAmexArq WHERE Estabelecimento = ? AND DataCriacao = ? AND Seq = ?", _
                lNumIntArq, objHeader.sEstabelecimento, objHeader.dtDataCriacao, objHeader.lSeq)
            If lErro <> AD_SQL_SUCESSO Then gError 216135
            
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216136
            If lErro <> AD_SQL_SEM_DADOS Then
            
                sNomeArq = String(STRING_MAXIMO, 0)
                
                lErro = Comando_Executar(alComando(4), "SELECT NomeArq FROM AdmExtFinArqsLidos WHERE NumIntDoc = ? AND NaoAtualizar = 0", _
                    sNomeArq, lNumIntArq)
                If lErro <> AD_SQL_SUCESSO Then gError 216143
                
                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216144
            
                If lErro = AD_SQL_SUCESSO Then gError 216137
            End If
            
            'inserir registro em AdmExtFinAmexArq
            lErro = Comando_Executar(alComando(4), "INSERT INTO AdmExtFinAmexArq (NumIntArq, Estabelecimento, DataCriacao, Seq) VALUES (?,?,?,?)", _
                mvarobjAux.lNumIntArq, objHeader.sEstabelecimento, objHeader.dtDataCriacao, objHeader.lSeq)
            If lErro <> AD_SQL_SUCESSO Then gError 216138
        
        Case "1" 'payment
            lNumRegArq = lNumRegArq + 1
            
            Set objPgto = New ClassArqAmexPgto
            objPgto.sReg = sRegistro
            objPgto.lNumRegArq = lNumRegArq
            
            'verificar bco/agencia e conta
            lErro = Comando_Executar(alComando(9), "SELECT C.Codigo, C.FilialEmpresa FROM ContasCorrentesInternas C, Redes R WHERE CodBanco = ? AND CONVERT(smallint,Agencia) = ? AND NumConta = ? AND R.Estabelecimento = ? AND C.FilialEmpresa = R.FilialEmpresa", _
                iCodConta, iFilEmpCCI, objPgto.iBanco, StrParaInt(objPgto.sAgencia), objPgto.sConta, objPgto.sEstabelecimento)
            If lErro <> AD_SQL_SUCESSO Then gError 216139

            lErro = Comando_BuscarProximo(alComando(9))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216140
            If lErro <> AD_SQL_SUCESSO Then gError 216141

            objPgto.iCodConta = iCodConta
            objPgto.iFilEmpCCI = iFilEmpCCI
        
            colPagtos.Add objPgto

        Case "3" 'SOC
            lNumRegArq = lNumRegArq + 1
            
            Set objSOC = New ClassArqAmexSOC
            objSOC.sReg = sRegistro
            objSOC.lNumRegArq = lNumRegArq
            
            Set objPgto = colPagtos.Item(colPagtos.Count)
            objPgto.colSOC.Add objSOC
            
        Case "4" 'ROC
            lNumRegArq = lNumRegArq + 1
            
            Set objROC = New ClassArqAmexROC
            objROC.sReg = sRegistro
            objROC.lNumRegArq = lNumRegArq
            
            If objROC.iQtdeParcelas = 0 Or objROC.iQtdeParcelas >= objROC.iNumParcela Then
            
                Set objPgto = colPagtos.Item(colPagtos.Count)
                Set objSOC = objPgto.colSOC.Item(objPgto.colSOC.Count)
                
                objROC.sNumSOC = objSOC.sNumSOC
                
                objSOC.colROC.Add objROC

            End If

        Case "5" 'adjustment
            lNumRegArq = lNumRegArq + 1
            Set objAjuste = New ClassArqAmexAjuste
            objAjuste.sReg = sRegistro
            objAjuste.lNumRegArq = lNumRegArq
            Call objAjuste.Processa(colPagtos)

        Case "6" 'pricing
            lNumRegArq = lNumRegArq + 1
        
        Case "9" 'trailer
            lNumRegArq = lNumRegArq + 1
            
            objTrailer.sReg = sRegistro
            
            If lNumRegArq <> objTrailer.lQtdeRegsArq Then gError 216145

    End Select

    ClassRotBxCartRede_ImportarArqTrataRegistro = SUCESSO

    Exit Function

Erro_ClassRotBxCartRede_ImportarArqTrataRegistro:

    ClassRotBxCartRede_ImportarArqTrataRegistro = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 216135, 216136
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMEXTFINAMEXARQ", gErr)

        Case 216137
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMEXTFINAMEXARQ_JA_IMPORTADO", gErr, objHeader.lSeq, objHeader.sEstabelecimento, objHeader.dtDataCriacao, sNomeArq)

        Case 216138
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_ADMEXTFINAMEXARQ", gErr)

        Case 216139, 216140
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REDE1", gErr)

        Case 216141
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_ENCONTRADA2", gErr, objPgto.iBanco, StrParaInt(objPgto.sAgencia), objPgto.sConta, objPgto.sEstabelecimento)

        Case 216143, 216144
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMEXTFINARQSLIDOS", gErr)
            
        Case 216145
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMEXTFINAMEXARQ_DIF_QTDE_REG", gErr, objTrailer.lQtdeRegsArq, lNumRegArq)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216142)

    End Select

    Exit Function

End Function

Private Property Get ClassRotBxCartRede_objAux() As ClassRotBxCartRedeAux

    Set ClassRotBxCartRede_objAux = mvarobjAux

End Property


