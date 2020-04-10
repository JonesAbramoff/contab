VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATGravaNena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function RelPosFinanc_Prepara(ByVal iFilialEmpresa As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iLayout As Integer, ByVal iCodConta As Integer, lNumIntRel As Long) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 20) As Long, sSQL As String
Dim alComandoAux1(0 To 4) As Long
Dim alComandoAux2(0 To 10) As Long
Dim iCredito As Integer, sNomeCta As String
Dim tMovCta As typeMovContaCorrente, iTipoGrupo As Integer, iSeq As Integer
Dim sTextoGrupoValor(1 To 3) As String, dValor(1 To 3) As Double, sTextoGrupo As String
Dim sTextoDescricao As String, sDescricao As String
Dim dSaldoCta As Double, dSaldoTotalIni As Double, dSaldoTotalFim As Double
Dim objNat As ClassNatMovCta, colNat As New Collection, sNatCta As String
Dim objCGGrupoNat As New ClassCamposGenericos, bAchou As Boolean
Dim objCGVGrupo As ClassCamposGenericosValores, sDescNatureza As String
Dim lFornecedor As Long, sFornecedor As String, iCount As Integer
Dim sNomeForn As String, sDescGrupoNat As String, dValorTotal(1 To 3) As Double
Dim colRet As Collection, objNatAux As ClassNatMovCtaAux

Const POS_FINANC_LAYOUT_DIARIO = 1
Const POS_FINANC_LAYOUT_MENSAL = 2

On Error GoTo Erro_RelPosFinanc_Prepara

    'abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211866
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211867
    Next
    
    For iIndice = LBound(alComandoAux1) To UBound(alComandoAux1)
        alComandoAux1(iIndice) = Comando_Abrir()
        If alComandoAux1(iIndice) = 0 Then gError 211868
    Next
    
    For iIndice = LBound(alComandoAux2) To UBound(alComandoAux2)
        alComandoAux2(iIndice) = Comando_Abrir()
        If alComandoAux2(iIndice) = 0 Then gError 211869
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUMINTREL_POSFINANCNENA", lNumIntRel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objCGGrupoNat.lCodigo = CAMPOSGENERICOS_NATCTA_GRUPO
    
    lErro = CF("CamposGenericosValores_Le_CodCampo", objCGGrupoNat)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iSeq = 0
    dSaldoTotalIni = 0
    dSaldoTotalFim = 0
    
    iTipoGrupo = 0
    sTextoGrupo = "" ' Informações sobre as contas correntes
    sTextoDescricao = "Banco"
    sTextoGrupoValor(1) = "Limite"
    sTextoGrupoValor(2) = "Saldo"
    sTextoGrupoValor(3) = "Disponível"
    
    dValorTotal(1) = 0
    dValorTotal(2) = 0
    dValorTotal(3) = 0
    
    sNomeCta = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT M.CodConta, C.NomeReduzido, SUM(((2 * T.Credito)-1) * M.Valor) AS Valor FROM MovimentosContaCorrente As M, TiposMovtoCtaCorrente AS T, ContasCorrentesInternas AS C WHERE M.Tipo = T.Codigo AND C.Codigo = M.CodConta AND M.Excluido <> 1 AND M.FilialEmpresa = ? AND M.DataMovimento BETWEEN ? AND ? AND (C.Codigo = ? OR ? = 0) GROUP BY M.CodConta, C.NomeReduzido ORDER BY C.NomeReduzido", _
    tMovCta.iCodConta, sNomeCta, tMovCta.dValor, iFilialEmpresa, dtDataDe, dtDataAte, iCodConta, iCodConta)
    If lErro <> AD_SQL_SUCESSO Then gError 211870
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211871
    
    Do While lErro = AD_SQL_SUCESSO
    
        iSeq = iSeq + 1
                
        lErro = CF("CCI_ObterRelTes", tMovCta.iCodConta, DateAdd("d", -1, dtDataDe), dSaldoCta, alComandoAux1())
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dSaldoTotalIni = dSaldoTotalIni + dSaldoCta
        
        sDescricao = sNomeCta
        dValor(1) = 0
        dValor(2) = tMovCta.dValor
        dValor(3) = dValor(2) - dValor(1)
    
        dValorTotal(1) = dValorTotal(1) + dValor(1)
        dValorTotal(2) = dValorTotal(2) + dValor(2)
        dValorTotal(3) = dValorTotal(3) + dValor(3)
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO RelPosFinanc (NumIntRel, Seq, TipoGrupo, TextoGrupo, TextoDescricao, TextoValor1, TextoValor2, TextoValor3, Descricao, Valor1, Valor2, Valor3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                                lNumIntRel, iSeq, iTipoGrupo, sTextoGrupo, sTextoDescricao, sTextoGrupoValor(1), sTextoGrupoValor(2), sTextoGrupoValor(3), sDescricao, dValor(1), dValor(2), dValor(3))
        If lErro <> AD_SQL_SUCESSO Then gError 211872
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211873
    
    Loop
    
    iSeq = iSeq + 1
    
    sDescricao = "Total:"
    dValor(1) = dValorTotal(1)
    dValor(2) = dValorTotal(2)
    dValor(3) = dValorTotal(3)
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO RelPosFinanc (NumIntRel, Seq, TipoGrupo, TextoGrupo, TextoDescricao, TextoValor1, TextoValor2, TextoValor3, Descricao, Valor1, Valor2, Valor3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                            lNumIntRel, iSeq, iTipoGrupo, sTextoGrupo, sTextoDescricao, sTextoGrupoValor(1), sTextoGrupoValor(2), sTextoGrupoValor(3), sDescricao, dValor(1), dValor(2), dValor(3))
    If lErro <> AD_SQL_SUCESSO Then gError 211874
    
    iTipoGrupo = 1
    sTextoGrupo = "Saldos" ' Saldo Anterior + Total Entradas + Total Saídas + Saldo Final
    sTextoDescricao = ""
    sTextoGrupoValor(1) = ""
    sTextoGrupoValor(2) = ""
    sTextoGrupoValor(3) = ""
    
    iSeq = iSeq + 1
    
    sDescricao = "Saldo Anterior em " & Format(DateAdd("d", -1, dtDataDe), "dd/mm/yyyy")
    dValor(1) = 0
    dValor(2) = 0
    dValor(3) = dSaldoTotalIni
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO RelPosFinanc (NumIntRel, Seq, TipoGrupo, TextoGrupo, TextoDescricao, TextoValor1, TextoValor2, TextoValor3, Descricao, Valor1, Valor2, Valor3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                            lNumIntRel, iSeq, iTipoGrupo, sTextoGrupo, sTextoDescricao, sTextoGrupoValor(1), sTextoGrupoValor(2), sTextoGrupoValor(3), sDescricao, dValor(1), dValor(2), dValor(3))
    If lErro <> AD_SQL_SUCESSO Then gError 211875
    
    lErro = Comando_Executar(alComando(2), "SELECT T.Credito, SUM(M.Valor) FROM MovimentosContaCorrente As M, TiposMovtoCtaCorrente AS T WHERE M.Tipo = T.Codigo AND M.Excluido <> 1 AND M.FilialEmpresa = ? AND M.DataMovimento BETWEEN ? AND ? AND (M.CodConta = ? OR ? = 0) GROUP BY T.Credito ORDER BY T.Credito DESC ", _
    iCredito, tMovCta.dValor, iFilialEmpresa, dtDataDe, dtDataAte, iCodConta, iCodConta)
    If lErro <> AD_SQL_SUCESSO Then gError 211876
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211877
    
    Do While lErro = AD_SQL_SUCESSO
    
        iSeq = iSeq + 1
        
        If iCredito = 0 Then
            sDescricao = "Total Saídas"
            dSaldoTotalFim = dSaldoTotalFim - tMovCta.dValor
        Else
            sDescricao = "Total Entradas"
            dSaldoTotalFim = dSaldoTotalFim + tMovCta.dValor
        End If
        dValor(1) = 0
        dValor(2) = 0
        dValor(3) = tMovCta.dValor
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO RelPosFinanc (NumIntRel, Seq, TipoGrupo, TextoGrupo, TextoDescricao, TextoValor1, TextoValor2, TextoValor3, Descricao, Valor1, Valor2, Valor3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                                lNumIntRel, iSeq, iTipoGrupo, sTextoGrupo, sTextoDescricao, sTextoGrupoValor(1), sTextoGrupoValor(2), sTextoGrupoValor(3), sDescricao, dValor(1), dValor(2), dValor(3))
        If lErro <> AD_SQL_SUCESSO Then gError 211878
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211879
    
    Loop
    
    iSeq = iSeq + 1
    
    sDescricao = "Saldo para " & Format(DateAdd("d", 1, dtDataAte), "dd/mm/yyyy")
    dValor(1) = 0
    dValor(2) = 0
    dValor(3) = dSaldoTotalFim
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO RelPosFinanc (NumIntRel, Seq, TipoGrupo, TextoGrupo, TextoDescricao, TextoValor1, TextoValor2, TextoValor3, Descricao, Valor1, Valor2, Valor3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                            lNumIntRel, iSeq, iTipoGrupo, sTextoGrupo, sTextoDescricao, sTextoGrupoValor(1), sTextoGrupoValor(2), sTextoGrupoValor(3), sDescricao, dValor(1), dValor(2), dValor(3))
    If lErro <> AD_SQL_SUCESSO Then gError 211880
    
    '******************************
    'Vai gerar uma tabela auxiliar por natureza/grupo de natureza/fornecedor
    
    tMovCta.sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)

    lErro = Comando_Executar(alComando(3), "SELECT M.NumMovto, T.Credito, T.Codigo, M.NumRefInterna, M.Natureza, M.Valor FROM MovimentosContaCorrente As M, TiposMovtoCtaCorrente AS T WHERE M.Tipo = T.Codigo AND M.Excluido <> 1 AND M.FilialEmpresa = ? AND M.DataMovimento BETWEEN ? AND ? AND (M.CodConta = ? OR ? = 0)", _
    tMovCta.lNumMovto, iCredito, tMovCta.iTipo, tMovCta.lNumRefInterna, tMovCta.sNatureza, tMovCta.dValor, iFilialEmpresa, dtDataDe, dtDataAte, iCodConta, iCodConta)
    If lErro <> AD_SQL_SUCESSO Then gError 211881
    
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211882
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set colRet = New Collection
    
        lErro = CF("RelPosFinanc_Obtem_Natureza", alComandoAux2(), tMovCta.iTipo, tMovCta.lNumMovto, tMovCta.sNatureza, tMovCta.dValor, colRet)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        For Each objNatAux In colRet
                
            Set objNat = New ClassNatMovCta
            
            If objNatAux.sNatureza = "" Then
                objNat.sDescricao = "NÃO PREENCHIDA"
            Else
                objNat.sCodigo = objNatAux.sNatureza
                lErro = CF("NatMovCta_Le", objNat)
                If lErro <> SUCESSO And lErro <> 122786 Then gError ERRO_SEM_MENSAGEM
                
                If objNat.lGrupo <> 0 Then
                    For Each objCGVGrupo In objCGGrupoNat.colCamposGenericosValores
                        If objCGVGrupo.lCodValor = objNat.lGrupo Then
                            objNat.sDescGrupo = objCGVGrupo.sValor
                            Exit For
                        End If
                    Next
                End If
            End If
            If iCredito = MARCADO Then
                objNat.dCredito = objNatAux.dValor
                objNat.dDebito = 0
            Else
                objNat.dDebito = objNatAux.dValor
                objNat.dCredito = 0
            End If
        
            lErro = Comando_Executar(alComando(4), "INSERT INTO RelPosFinancAux (NumIntRel, NumMovto, TipoDoc, NumIntDoc, GrupoNat, DescGrupoNat, Natureza, DescNatureza, Credito, Debito, TipoCliForn, CliForn, NomeCliForn) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                                    lNumIntRel, tMovCta.lNumMovto, objNatAux.iTipoDoc, objNatAux.lNumIntDoc, objNat.lGrupo, objNat.sDescGrupo, objNat.sCodigo, objNat.sDescricao, objNat.dCredito, objNat.dDebito, objNatAux.iTipoCliForn, objNatAux.lCliForn, objNatAux.sNomeCliForn)
            If lErro <> AD_SQL_SUCESSO Then gError 211883

        Next

        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211884
     
    Loop
    '******************************
    
    iTipoGrupo = 2
    sTextoGrupo = "Entradas" ' Entradas por natureza
    sTextoDescricao = ""
    sTextoGrupoValor(1) = ""
    sTextoGrupoValor(2) = ""
    sTextoGrupoValor(3) = ""
    
    sDescNatureza = String(STRING_NATMOVCTA_DESCRICAO, 0)

    lErro = Comando_Executar(alComando(5), "SELECT DescNatureza, SUM(Credito) FROM RelPosFinancAux WHERE NumIntRel = ? AND Natureza <> '' GROUP BY DescNatureza HAVING SUM(Credito) > 0 ORDER BY DescNatureza ", _
    sDescNatureza, tMovCta.dValor, lNumIntRel)
    If lErro <> AD_SQL_SUCESSO Then gError 211885
    
    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211886
    
    Do While lErro = AD_SQL_SUCESSO

        iSeq = iSeq + 1
        
        sDescricao = sDescNatureza
        dValor(1) = 0
        dValor(2) = 0
        dValor(3) = tMovCta.dValor
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO RelPosFinanc (NumIntRel, Seq, TipoGrupo, TextoGrupo, TextoDescricao, TextoValor1, TextoValor2, TextoValor3, Descricao, Valor1, Valor2, Valor3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                                lNumIntRel, iSeq, iTipoGrupo, sTextoGrupo, sTextoDescricao, sTextoGrupoValor(1), sTextoGrupoValor(2), sTextoGrupoValor(3), sDescricao, dValor(1), dValor(2), dValor(3))
        If lErro <> AD_SQL_SUCESSO Then gError 211887

        lErro = Comando_BuscarProximo(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211888
     
    Loop
    
    iTipoGrupo = 3
    sTextoGrupo = "Saídas" ' Saídas por natureza
    sTextoDescricao = ""
    sTextoGrupoValor(1) = ""
    sTextoGrupoValor(2) = ""
    sTextoGrupoValor(3) = ""
       
    If iLayout = POS_FINANC_LAYOUT_DIARIO Then 'Por Natureza com top fornecedores
    
        sNomeForn = String(STRING_MAXIMO, 0)
    
        'Exibe o fornecedor para algumas naturezas
        lErro = Comando_Executar(alComando(6), "SELECT CONVERT(VARCHAR,CliForn) + '-' + NomeCliForn, SUM(Debito) FROM RelPosFinancAux WHERE NumIntRel = ?  AND Natureza <> '' AND DescGrupoNAT IN ('COMPRA DE MAT. PRIMA','DESP. DE PRODUCAO','EMBALAGENS') GROUP BY CONVERT(VARCHAR,CliForn) + '-' + NomeCliForn HAVING SUM(Debito) > 0 ORDER BY SUM(Debito) DESC ", _
        sNomeForn, tMovCta.dValor, lNumIntRel)
        If lErro <> AD_SQL_SUCESSO Then gError 211889
        
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211890
        
        Do While lErro = AD_SQL_SUCESSO
    
            iSeq = iSeq + 1
            
            sDescricao = sNomeForn
            dValor(1) = 0
            dValor(2) = 0
            dValor(3) = tMovCta.dValor
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelPosFinanc (NumIntRel, Seq, TipoGrupo, TextoGrupo, TextoDescricao, TextoValor1, TextoValor2, TextoValor3, Descricao, Valor1, Valor2, Valor3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                                    lNumIntRel, iSeq, iTipoGrupo, sTextoGrupo, sTextoDescricao, sTextoGrupoValor(1), sTextoGrupoValor(2), sTextoGrupoValor(3), sDescricao, dValor(1), dValor(2), dValor(3))
            If lErro <> AD_SQL_SUCESSO Then gError 211891
    
            lErro = Comando_BuscarProximo(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211892
         
        Loop

        sDescNatureza = String(STRING_MAXIMO, 0)

        'Exibe por natureza para as demais
        lErro = Comando_Executar(alComando(7), "SELECT DescNatureza, SUM(Debito) FROM RelPosFinancAux WHERE NumIntRel = ? AND Natureza <> '' AND DescGrupoNAT NOT IN ('COMPRA DE MAT. PRIMA','DESP. DE PRODUCAO','EMBALAGENS') GROUP BY DescNatureza HAVING SUM(Debito) > 0 ORDER BY DescNatureza ", _
        sDescNatureza, tMovCta.dValor, lNumIntRel)
        If lErro <> AD_SQL_SUCESSO Then gError 211893
        
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211894
        
        Do While lErro = AD_SQL_SUCESSO
    
            iSeq = iSeq + 1
            
            sDescricao = sDescNatureza
            dValor(1) = 0
            dValor(2) = 0
            dValor(3) = tMovCta.dValor
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelPosFinanc (NumIntRel, Seq, TipoGrupo, TextoGrupo, TextoDescricao, TextoValor1, TextoValor2, TextoValor3, Descricao, Valor1, Valor2, Valor3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                                    lNumIntRel, iSeq, iTipoGrupo, sTextoGrupo, sTextoDescricao, sTextoGrupoValor(1), sTextoGrupoValor(2), sTextoGrupoValor(3), sDescricao, dValor(1), dValor(2), dValor(3))
            If lErro <> AD_SQL_SUCESSO Then gError 211895
    
            lErro = Comando_BuscarProximo(alComando(7))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211896
         
        Loop
       
    Else 'Por grupo de natureza
    
        sDescGrupoNat = String(STRING_MAXIMO, 0)

        lErro = Comando_Executar(alComando(8), "SELECT DescGrupoNat, SUM(Debito) FROM RelPosFinancAux WHERE NumIntRel = ? AND Natureza <> '' GROUP BY DescGrupoNat HAVING SUM(Debito) > 0 ORDER BY DescGrupoNat ", _
        sDescGrupoNat, tMovCta.dValor, lNumIntRel)
        If lErro <> AD_SQL_SUCESSO Then gError 211897
        
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211898
        
        Do While lErro = AD_SQL_SUCESSO
    
            iSeq = iSeq + 1
            
            sDescricao = sDescGrupoNat
            dValor(1) = 0
            dValor(2) = 0
            dValor(3) = tMovCta.dValor
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelPosFinanc (NumIntRel, Seq, TipoGrupo, TextoGrupo, TextoDescricao, TextoValor1, TextoValor2, TextoValor3, Descricao, Valor1, Valor2, Valor3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                                    lNumIntRel, iSeq, iTipoGrupo, sTextoGrupo, sTextoDescricao, sTextoGrupoValor(1), sTextoGrupoValor(2), sTextoGrupoValor(3), sDescricao, dValor(1), dValor(2), dValor(3))
            If lErro <> AD_SQL_SUCESSO Then gError 211899
            
            sDescNatureza = String(STRING_MAXIMO, 0)
            
            lErro = Comando_Executar(alComando(9), "SELECT DescNatureza, SUM(Debito) FROM RelPosFinancAux WHERE NumIntRel = ?  AND Natureza <> '' AND DescGrupoNat = ? GROUP BY DescNatureza HAVING SUM(Debito) > 0 ORDER BY DescNatureza ", _
            sDescNatureza, tMovCta.dValor, lNumIntRel, sDescGrupoNat)
            If lErro <> AD_SQL_SUCESSO Then gError 211900
            
            lErro = Comando_BuscarPrimeiro(alComando(9))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211901
            
            Do While lErro = AD_SQL_SUCESSO
        
                iSeq = iSeq + 1
                
                sDescricao = "     " & sDescNatureza
                dValor(1) = 0
                dValor(2) = 0
                dValor(3) = tMovCta.dValor
            
                lErro = Comando_Executar(alComando(1), "INSERT INTO RelPosFinanc (NumIntRel, Seq, TipoGrupo, TextoGrupo, TextoDescricao, TextoValor1, TextoValor2, TextoValor3, Descricao, Valor1, Valor2, Valor3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ", _
                                                        lNumIntRel, iSeq, iTipoGrupo, sTextoGrupo, sTextoDescricao, sTextoGrupoValor(1), sTextoGrupoValor(2), sTextoGrupoValor(3), sDescricao, dValor(1), dValor(2), dValor(3))
                If lErro <> AD_SQL_SUCESSO Then gError 211902
        
        
                lErro = Comando_BuscarProximo(alComando(9))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211903
             
            Loop
    
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211904
         
        Loop
        
    End If

    lErro = Comando_Executar(alComando(10), "DELETE RelPosFinancAux WHERE NumIntRel = ? ", lNumIntRel)
    If lErro <> AD_SQL_SUCESSO Then gError 211905

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    For iIndice = LBound(alComandoAux1) To UBound(alComandoAux1)
        Call Comando_Fechar(alComandoAux1(iIndice))
    Next
    
    For iIndice = LBound(alComandoAux2) To UBound(alComandoAux2)
        Call Comando_Fechar(alComandoAux2(iIndice))
    Next
    
     'fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211906

    RelPosFinanc_Prepara = SUCESSO

    Exit Function

Erro_RelPosFinanc_Prepara:

    RelPosFinanc_Prepara = gErr

    Select Case gErr

        Case 211866
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 211867 To 211869
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211870, 211871, 211873, 211876, 211877, 211879, 211881, 211882, 211884
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOSCONTACORRENTE", gErr)

        Case 211872, 211874, 211875, 211878, 211880, 211883, 211887, 211891, 211895, 211899, 211902
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELPOSFINANC", gErr)
        
        Case 211885, 211886, 211888, 211889, 211890, 211892, 211893, 211894, 211896, 211897, 211898, 211900, 211901, 211903, 211904
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELPOSFINANCAUX", gErr)
        
        Case 211905
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_RELPOSFINANCAUX", gErr)
        
        Case 211906
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211907)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    For iIndice = LBound(alComandoAux1) To UBound(alComandoAux1)
        Call Comando_Fechar(alComandoAux1(iIndice))
    Next
    For iIndice = LBound(alComandoAux2) To UBound(alComandoAux2)
        Call Comando_Fechar(alComandoAux2(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function RelPosFinanc_Obtem_Natureza(alComando() As Long, ByVal iTipoMov As Integer, ByVal lNumMovto As Long, ByVal sNaturezaMovto As String, ByVal dValorMovto As Double, ByVal colRet As Collection) As Long

Dim lErro As Long
Dim objNatAux As ClassNatMovCtaAux
Dim sSQL As String
Dim sNatureza As String, dValor As Double, dValorJuros As Double, lNumIntDoc As Long
Dim dValorJurosTotal As Double, lCliForn As Long, sNomeCliForn As String

On Error GoTo Erro_RelPosFinanc_Obtem_Natureza

    Select Case iTipoMov
    
        Case 6 'Adiantamento para Fornecedor
               
            sNomeCliForn = String(STRING_MAXIMO, 0)
               
            lErro = Comando_Executar(alComando(1), "SELECT P.Fornecedor, F.NomeReduzido FROM PagtosAntecipados AS P, Fornecedores AS F WHERE P.Fornecedor = F.Codigo AND NumMovto = ?", _
            lCliForn, sNomeCliForn, lNumMovto)
            If lErro <> AD_SQL_SUCESSO Then gError 211908
            
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211909
            
            Set objNatAux = New ClassNatMovCtaAux
            colRet.Add objNatAux
            objNatAux.iTipoDoc = 0
            objNatAux.lNumIntDoc = 0
            objNatAux.sNatureza = sNaturezaMovto
            objNatAux.dValor = dValorMovto
            objNatAux.iTipoCliForn = 2
            objNatAux.lCliForn = lCliForn
            objNatAux.sNomeCliForn = sNomeCliForn
        
        Case 7 'Adiantamento de Cliente
        
            sNomeCliForn = String(STRING_MAXIMO, 0)
        
            lErro = Comando_Executar(alComando(2), "SELECT R.Cliente, C.NomeReduzido FROM RecebAntecipados AS R, Clientes AS C WHERE R.Cliente = C.Codigo AND NumMovto = ?", _
            lCliForn, sNomeCliForn, lNumMovto)
            If lErro <> AD_SQL_SUCESSO Then gError 211910
            
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211911
        
            Set objNatAux = New ClassNatMovCtaAux
            colRet.Add objNatAux
            objNatAux.iTipoDoc = 0
            objNatAux.lNumIntDoc = 0
            objNatAux.sNatureza = sNaturezaMovto
            objNatAux.dValor = dValorMovto
            objNatAux.iTipoCliForn = 1
            objNatAux.lCliForn = lCliForn
            objNatAux.sNomeCliForn = sNomeCliForn
            
        Case 8 'Recebimento de cliente - > BaixasRec
        
            dValorJurosTotal = 0
        
            sNatureza = String(STRING_MAXIMO, 0)
            sNomeCliForn = String(STRING_MAXIMO, 0)
        
            sSQL = "SELECT T.Cliente, C.NomeReduzido, T.Natureza, BP.NumIntDoc, BP.ValorRecebido, BP.ValorJuros + BP.ValorMulta FROM TitulosRecTodos AS T, ParcelasRecTodas AS P, BaixasParcRec As BP, Clientes AS C, BaixasRec As B"
            sSQL = sSQL & " " & "WHERE T.NumIntDoc = P.NumIntTitulo And P.NumIntDoc = BP.NumIntParcela AND T.Cliente = C.Codigo AND B.NumIntBaixa = BP.NumIntBaixa AND BP.Status <> 5 AND B.Status <> 5"
            sSQL = sSQL & " " & "AND B.NumMovCta = ?"
        
            lErro = Comando_Executar(alComando(3), sSQL, _
            lCliForn, sNomeCliForn, sNatureza, lNumIntDoc, dValor, dValorJuros, lNumMovto)
            If lErro <> AD_SQL_SUCESSO Then gError 211912
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211913
            
            Do While lErro = AD_SQL_SUCESSO
        
                dValorJurosTotal = dValorJurosTotal + dValorJuros
                
                Set objNatAux = New ClassNatMovCtaAux
                colRet.Add objNatAux
                objNatAux.iTipoDoc = TRANSACAOCTBORIGEM_BAIXASPARCREC
                objNatAux.lNumIntDoc = lNumIntDoc
                objNatAux.sNatureza = sNatureza
                objNatAux.iTipoCliForn = 1
                objNatAux.lCliForn = lCliForn
                objNatAux.sNomeCliForn = sNomeCliForn
                objNatAux.dValor = dValor - dValorJuros
        
                If dValorJuros > DELTA_VALORMONETARIO Then
                    Set objNatAux = New ClassNatMovCtaAux
                    colRet.Add objNatAux
                    objNatAux.iTipoDoc = TRANSACAOCTBORIGEM_BAIXASPARCREC
                    objNatAux.lNumIntDoc = lNumIntDoc
                    objNatAux.sNatureza = "030008    "
                    objNatAux.dValor = dValorJuros
                    objNatAux.iTipoCliForn = 1
                    objNatAux.lCliForn = lCliForn
                    objNatAux.sNomeCliForn = sNomeCliForn
                End If
        
                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211914
             
            Loop
        
        Case 9, 10, 12 'Pagto de Título por Cheque/Borderô/Dinheiro - > BaixasPag
        
            dValorJurosTotal = 0
        
            sNatureza = String(STRING_MAXIMO, 0)
            sNomeCliForn = String(STRING_MAXIMO, 0)
        
            sSQL = "SELECT T.Fornecedor, C.NomeReduzido, T.Natureza, BP.NumIntDoc, BP.ValorBaixado FROM TitulosPagTodos AS T, ParcelasPagTodas AS P, BaixasParcPag As BP, Fornecedores AS C, BaixasPag As B"
            sSQL = sSQL & " " & "WHERE T.NumIntDoc = P.NumIntTitulo And P.NumIntDoc = BP.NumIntParcela AND T.Fornecedor = C.Codigo AND B.NumIntBaixa = BP.NumIntBaixa AND BP.Status <> 5 AND B.Status <> 5"
            sSQL = sSQL & " " & "AND B.NumMovCta = ?"
        
            lErro = Comando_Executar(alComando(4), sSQL, _
            lCliForn, sNomeCliForn, sNatureza, lNumIntDoc, dValor, lNumMovto)
            If lErro <> AD_SQL_SUCESSO Then gError 211915
            
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211916
            
            Do While lErro = AD_SQL_SUCESSO
        
                dValorJurosTotal = dValorJurosTotal + dValorJuros
                
                Set objNatAux = New ClassNatMovCtaAux
                colRet.Add objNatAux
                objNatAux.iTipoDoc = TRANSACAOCTBORIGEM_BAIXASPARCPAG
                objNatAux.lNumIntDoc = lNumIntDoc
                objNatAux.sNatureza = sNatureza
                objNatAux.dValor = dValor
                objNatAux.iTipoCliForn = 2
                objNatAux.lCliForn = lCliForn
                objNatAux.sNomeCliForn = sNomeCliForn
        
                lErro = Comando_BuscarProximo(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211917
             
            Loop
            
        Case Else
            Set objNatAux = New ClassNatMovCtaAux
            colRet.Add objNatAux
            objNatAux.iTipoDoc = 0
            objNatAux.lNumIntDoc = 0
            objNatAux.sNatureza = sNaturezaMovto
            objNatAux.dValor = dValorMovto
    
    End Select

    RelPosFinanc_Obtem_Natureza = SUCESSO

    Exit Function

Erro_RelPosFinanc_Obtem_Natureza:

    RelPosFinanc_Obtem_Natureza = gErr

    Select Case gErr

        Case 211908, 211909
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PAGTO_ANTECIPADO2", gErr)

        Case 211910, 211911
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RECEB_ANTEC", gErr)

        Case 211911 To 211914
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", gErr)

        Case 211914 To 211917
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_PAG2", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211918)

    End Select
    
    Exit Function
    
End Function
