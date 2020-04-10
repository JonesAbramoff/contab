VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFormata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Ccl_Formata(ByVal sCcl As String, sCclFormatada As String, iCclPreenchida As Integer) As Long
'Devolve Formato BD: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
'A String de Entrada vem do GRID com brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'iCclPreenchida retornará CCL_VAZIA se nenhum segmento do centro de custo for preenchido,
'caso contrário retornará CCL_PREENCHIDA

Dim lErro As Long
Dim iNumSegmentos As Integer
Dim iInd1 As Integer
Dim iInd2 As Integer
Dim objSegmento As New ClassSegmento
Dim sCclBruta As String
Dim sSegmentoBruto As String 'Segmento sendo formatado
Dim sSegmentoFormatado As String 'Segmento sendo formatado
Dim sCaracter As String 'caracter sendo tratado
Dim iSegmentoVazio As Integer
Dim iRestoVazio As Integer
    
On Error GoTo Erro_Ccl_Formata
            
    iCclPreenchida = CCL_VAZIA
    iRestoVazio = SEGMENTOS_RESTANTES_CHEIOS
    sCclBruta = sCcl
    sCclFormatada = ""
    iNumSegmentos = NumSegCcl()

    For iInd1 = 1 To iNumSegmentos
        
        lErro = SegmentoCcl(iInd1, objSegmento)
        If lErro <> SUCESSO Then Error 6122
        
        sSegmentoBruto = Left(sCclBruta, objSegmento.iTamanho)
        sCclBruta = Mid(sCclBruta, objSegmento.iTamanho + Len(objSegmento.sDelimitador) + 1)
        sSegmentoFormatado = ""

        iSegmentoVazio = SEGMENTO_VAZIO

        Select Case objSegmento.iTipo
        Case SEGMENTO_NUMERICO

            For iInd2 = 1 To objSegmento.iTamanho

                sCaracter = Mid(sSegmentoBruto, iInd2, 1)

                If sCaracter = " " Or sCaracter = "" Then
                    sSegmentoFormatado = "0" & sSegmentoFormatado
                    
                ElseIf sCaracter = "0" Then
                    sSegmentoFormatado = sSegmentoFormatado & sCaracter
                ElseIf sCaracter = "-" Then
                    Error 55690
                Else
                    sSegmentoFormatado = sSegmentoFormatado & sCaracter
                    iSegmentoVazio = SEGMENTO_CHEIO
                End If

            Next

        Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII

            For iInd2 = objSegmento.iTamanho To 1 Step -1

                sCaracter = Mid(sSegmentoBruto, iInd2, 1)

                If sCaracter = " " Or sCaracter = "" Then
                    sSegmentoFormatado = sSegmentoFormatado & " "
                Else
                    sSegmentoFormatado = sCaracter & sSegmentoFormatado
                    iSegmentoVazio = SEGMENTO_CHEIO
                End If

            Next

        End Select

        If iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS And iSegmentoVazio <> SEGMENTO_VAZIO Then Error 5955
            
        If iSegmentoVazio = SEGMENTO_VAZIO Then
            iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS
        Else
            iCclPreenchida = CCL_PREENCHIDA
        End If

        sCclFormatada = sCclFormatada & sSegmentoFormatado

    Next

    If iCclPreenchida = CCL_VAZIA Then sCclFormatada = ""

    Ccl_Formata = SUCESSO
    
    Exit Function
    
Erro_Ccl_Formata:

    Ccl_Formata = Err

    Select Case Err

        Case 5955
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_SEG_MEIO_NAO_PREENCHIDOS", Err)

        Case 6122 'Vazio porque o erro já foi tratado
        
        Case 55690
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_SEG_NUM_CARACTER_INVALIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150475)
        

    End Select

    Exit Function
    
End Function

Private Function Ccl_Formata1(ByVal sCcl As String, sCclFormatada As String, iCclPreenchida As Integer) As Long
'Utilizada somente pela função Conta_Formata
'Devolve Formato BD: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
'A String de Entrada vem do GRID com brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'iCclPreenchida retornará CCL_VAZIA se nenhum segmento do centro de custo for preenchido,
'caso contrário retornará CCL_PREENCHIDA

Dim lErro As Long
Dim iNumSegmentos As Integer
Dim iInd1 As Integer
Dim iInd2 As Integer
Dim objSegmento As New ClassSegmento
Dim sCclBruta As String
Dim sSegmentoBruto As String 'Segmento sendo formatado
Dim sSegmentoFormatado As String 'Segmento sendo formatado
Dim sCaracter As String 'caracter sendo tratado
Dim iSegmentoVazio As Integer
Dim iRestoVazio As Integer
    
On Error GoTo Erro_Ccl_Formata1
            
            
    iRestoVazio = SEGMENTOS_RESTANTES_CHEIOS
    sCclBruta = sCcl
    sCclFormatada = ""
    iNumSegmentos = NumSegCcl()

    For iInd1 = 1 To iNumSegmentos
        
        lErro = SegmentoCcl(iInd1, objSegmento)
        If lErro <> SUCESSO Then Error 55950
        
        sSegmentoBruto = Left(sCclBruta, objSegmento.iTamanho)
        sCclBruta = Mid(sCclBruta, objSegmento.iTamanho + Len(objSegmento.sDelimitador) + 1)
        sSegmentoFormatado = ""

        iSegmentoVazio = SEGMENTO_VAZIO

        Select Case objSegmento.iTipo
        Case SEGMENTO_NUMERICO

            For iInd2 = 1 To objSegmento.iTamanho

                sCaracter = Mid(sSegmentoBruto, iInd2, 1)

                If sCaracter = " " Or sCaracter = "" Then
                    sSegmentoFormatado = "0" & sSegmentoFormatado
                    
                ElseIf sCaracter = "0" Then
                    sSegmentoFormatado = sSegmentoFormatado & sCaracter
                ElseIf sCaracter = "-" Then
                    Error 55951
                Else
                    sSegmentoFormatado = sSegmentoFormatado & sCaracter
                    iSegmentoVazio = SEGMENTO_CHEIO
                End If

            Next

        Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII

            For iInd2 = objSegmento.iTamanho To 1 Step -1

                sCaracter = Mid(sSegmentoBruto, iInd2, 1)

                If sCaracter = " " Or sCaracter = "" Then
                    sSegmentoFormatado = sSegmentoFormatado & " "
                Else
                    sSegmentoFormatado = sCaracter & sSegmentoFormatado
                    iSegmentoVazio = SEGMENTO_CHEIO
                End If

            Next

        End Select

        If iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS And iSegmentoVazio <> SEGMENTO_VAZIO Then Error 55952
            
        If iSegmentoVazio = SEGMENTO_VAZIO Then
            iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS
        Else
            iCclPreenchida = CCL_PREENCHIDA
        End If

        sCclFormatada = sCclFormatada & sSegmentoFormatado

    Next

    If iCclPreenchida = CCL_VAZIA Then sCclFormatada = ""

    Ccl_Formata1 = SUCESSO
    
    Exit Function
    
Erro_Ccl_Formata1:

    Ccl_Formata1 = Err

    Select Case Err

        Case 55950
        
        Case 55951
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_SEG_NUM_CARACTER_INVALIDO", Err)

        Case 55952
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_SEG_MEIO_NAO_PREENCHIDOS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150476)
        

    End Select

    Exit Function
    
End Function

Function Conta_Formata(ByVal sConta As String, sContaFormatada As String, iContaPreenchida As Integer) As Long
'Devolve Formato BD: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
'A String de Entrada vem do GRID com os caracteres delimitadores e brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'iContaPreenchida retornará CONTA_VAZIA se nenhum segmento da conta for preenchido,
'caso contrário retornará CONTA_PREENCHIDA
    
Dim lErro As Long
Dim iNumSegmentos As Integer
Dim iInd1 As Integer
Dim iInd2 As Integer
Dim objSegmento As New ClassSegmento
Dim sContaBruta As String
Dim sSegmentoBruto As String 'Segmento sendo formatado
Dim sSegmentoFormatado As String 'Segmento sendo formatado
Dim sCaracter As String 'caracter sendo tratado
Dim sCcl As String
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim iSegmentoVazio As Integer
Dim iRestoVazio As Integer
    
On Error GoTo Erro_Conta_Formata
            
    iRestoVazio = SEGMENTOS_RESTANTES_CHEIOS
    iContaPreenchida = CONTA_VAZIA
    sContaBruta = sConta
    sContaFormatada = ""
    iNumSegmentos = NumSegConta()

    For iInd1 = 1 To iNumSegmentos
        
        lErro = SegmentoConta(iInd1, objSegmento)
        If lErro <> SUCESSO Then Error 6072
        
        sSegmentoBruto = Left(sContaBruta, objSegmento.iTamanho)
        sContaBruta = Mid(sContaBruta, objSegmento.iTamanho + Len(objSegmento.sDelimitador) + 1)
        sSegmentoFormatado = ""

        iSegmentoVazio = SEGMENTO_VAZIO
        
        Select Case objSegmento.iTipo
        
            Case SEGMENTO_NUMERICO
        
                For iInd2 = 1 To objSegmento.iTamanho
        
                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
        
                    If sCaracter = " " Or sCaracter = "" Then
                        sSegmentoFormatado = "0" & sSegmentoFormatado
                        
                    ElseIf sCaracter = "0" Then
                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
                        
                    ElseIf sCaracter = "-" Then
                        Error 55689
                    Else
                        iSegmentoVazio = SEGMENTO_CHEIO
                        sSegmentoFormatado = sSegmentoFormatado & sCaracter
                    End If
        
                Next
                
                
            Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII
        
                For iInd2 = objSegmento.iTamanho To 1 Step -1
        
                    sCaracter = Mid(sSegmentoBruto, iInd2, 1)
        
                    If sCaracter = " " Or sCaracter = "" Then
                        sSegmentoFormatado = sSegmentoFormatado & " "
                    Else
                        iSegmentoVazio = SEGMENTO_CHEIO
                        sSegmentoFormatado = sCaracter & sSegmentoFormatado
                    End If
        
                Next

        End Select

        If iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS And iSegmentoVazio <> SEGMENTO_VAZIO Then Error 5954
            
        If iSegmentoVazio = SEGMENTO_VAZIO Then
            iRestoVazio = SEGMENTOS_RESTANTES_VAZIOS
        Else
            iContaPreenchida = CONTA_PREENCHIDA
        End If

        sContaFormatada = sContaFormatada & sSegmentoFormatado

    Next
    
    If giSetupUsoCcl = CCL_USA_CONTABIL Then
    
        lErro = SegmentoCcl(1, objSegmento)
        If lErro <> SUCESSO Then Error 5820
        
        sCcl = sContaBruta
        
        If iContaPreenchida = CONTA_PREENCHIDA Then iCclPreenchida = CCL_PREENCHIDA
        
        lErro = Ccl_Formata1(sCcl, sCclFormatada, iCclPreenchida)
        If lErro Then Error 5819
       
        If iCclPreenchida = CCL_PREENCHIDA Then iContaPreenchida = CONTA_PREENCHIDA
            
        sContaFormatada = sContaFormatada & sCclFormatada
        
    End If
    
    If iContaPreenchida = CONTA_VAZIA Then sContaFormatada = ""
    
    Conta_Formata = SUCESSO
    
    Exit Function
    
Erro_Conta_Formata:

    Conta_Formata = Err

    Select Case Err
    
        Case 5819, 5820, 6072 'Vazio porque o erro já foi tratado
        
        Case 5954
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_SEG_MEIO_NAO_PREENCHIDOS", Err)
        
        Case 55689
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_SEG_NUM_CARACTER_INVALIDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150477)

    End Select

    Exit Function
    
End Function

Function Conta_Formata_SemCcl(ByVal sConta As String, sContaFormatada As String, iContaPreenchida As Integer) As Long
'Devolve a conta no Formato do BD sem centro de custo: zeros a esquerda em segmentos numericos e brancos a direita para segmentos alfabeticos
'A String de Entrada vem do GRID com os caracteres delimitadores e brancos(se for o caso). O comprimento dela é o comprimento total da máscara.
'iContaPreenchida retornará CONTA_VAZIA se nenhum segmento da conta for preenchido,
'caso contrário retornará CONTA_PREENCHIDA
    
Dim lErro As Long
Dim iNumSegmentos As Integer
Dim iInd1 As Integer
Dim iInd2 As Integer
Dim objSegmento As New ClassSegmento
Dim sContaBruta As String
Dim sSegmentoBruto As String 'Segmento sendo formatado
Dim sSegmentoFormatado As String 'Segmento sendo formatado
Dim sCaracter As String 'caracter sendo tratado
Dim sCcl As String
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
    
    On Error GoTo Erro_Conta_Formata_SemCcl
            
    iContaPreenchida = 0
    sContaBruta = sConta
    sContaFormatada = ""
    iNumSegmentos = NumSegConta()

    For iInd1 = 1 To iNumSegmentos
        
        lErro = SegmentoConta(iInd1, objSegmento)
        If lErro <> SUCESSO Then Error 6072
        
        sSegmentoBruto = Left(sContaBruta, objSegmento.iTamanho)
        sContaBruta = Mid(sContaBruta, objSegmento.iTamanho + Len(objSegmento.sDelimitador) + 1)
        sSegmentoFormatado = ""

        Select Case objSegmento.iTipo
        Case SEGMENTO_NUMERICO

            For iInd2 = 1 To objSegmento.iTamanho

                sCaracter = Mid(sSegmentoBruto, iInd2, 1)

                If sCaracter = " " Or sCaracter = "" Then
                    sSegmentoFormatado = "0" & sSegmentoFormatado
                ElseIf sCaracter = "-" Then
                    Error 55691
                Else
                    iContaPreenchida = CONTA_PREENCHIDA
                    sSegmentoFormatado = sSegmentoFormatado & sCaracter
                End If

            Next

        Case SEGMENTO_ALFANUMERICO, SEGMENTO_ASCII

            For iInd2 = objSegmento.iTamanho To 1 Step -1

                sCaracter = Mid(sSegmentoBruto, iInd2, 1)

                If sCaracter = " " Or sCaracter = "" Then
                    sSegmentoFormatado = sSegmentoFormatado & " "
                Else
                    iContaPreenchida = CONTA_PREENCHIDA
                    sSegmentoFormatado = sCaracter & sSegmentoFormatado
                End If

            Next

        End Select

        sContaFormatada = sContaFormatada & sSegmentoFormatado

    Next
    
    Conta_Formata_SemCcl = SUCESSO
    
    Exit Function
    
Erro_Conta_Formata_SemCcl:

    Conta_Formata_SemCcl = Err


    Select Case Err

        Case 5858 'Vazio porque o erro já foi tratado

        Case 55691
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_SEG_NUM_CARACTER_INVALIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150478)

    End Select

    Exit Function
    
End Function


Function Str_Limpa(sString As String, sChar As String) As String
'Retorna sString sem sChar (caracter)

    Dim sResult As String
    Dim sRabo As String
    Dim sSeg As String
    Dim iPos As Integer
    
    sResult = ""
    sRabo = sString
    iPos = InStr(sRabo, sChar)
    Do While Not iPos = 0
        If iPos > 1 Then
            sSeg = Left$(sRabo, iPos - 1)
            sResult = sResult & sSeg
            sRabo = Mid$(sRabo, iPos + 1)
        Else
            sRabo = Mid$(sRabo, iPos + 1)
        End If
        iPos = InStr(sRabo, sChar)
    Loop
    sResult = sResult & sRabo
    Str_Limpa = sResult
End Function


Function Str_Subst(sString As String, sChar1 As String, sChar2 As String) As String
'Substitui sChar1 por sChar2 em sString

    Dim sResult As String
    Dim sRabo As String
    Dim sSeg As String
    Dim iPos As Integer
    
    sResult = ""
    sRabo = sString
    iPos = InStr(sRabo, sChar1)
    Do While Not iPos = 0
        sSeg = Left$(sRabo, iPos - 1) & sChar2
        sResult = sResult & sSeg
        sRabo = Mid$(sRabo, iPos + 1)
        iPos = InStr(sRabo, sChar1)
    Loop
    sResult = sResult & sRabo
    Str_Subst = sResult
End Function

Function Browse_Formata_Campo(objBrowse As AdmBrowse, ByVal colRegCampo As Collection, ByVal objValorCampo As AdmValorCampo, vCampo As Variant, iProximoCampo As Integer, iLinha As Integer, iIndice As Integer, iExercicio As Integer, iPeriodo As Integer, sNomeExterno As String) As Long

Dim sContaMascarada As String
Dim sCclMascarado As String
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo
Dim sConta As String
Dim sCcl As String
Dim lErro As Long
Dim sData As String
Dim sTipoContaCcl As String
Dim sNaturezaConta As String
Dim sTipoRateio As String
Dim sAglutina As String

On Error GoTo Erro_Browse_Formata_Campo

    Select Case objValorCampo.iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case objValorCampo.iSubtipo
        
                Case ADM_SUBTIPO_PERCENTUAL
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Percent")
        
                Case ADM_SUBTIPO_HORA
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CDate(vCampo)
        
                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Standard")
                    
            End Select
        
        Case ADM_TIPO_DATE
            sData = Format(vCampo, "dd/mm/yyyy")
        
            If sData = "07/09/1822" Then sData = ""
        
            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sData
    
        Case ADM_TIPO_SMALLINT
    
            Select Case objValorCampo.iSubtipo
        
                Case ADM_SUBTIPO_TIPOCONTACCL
                    Call CF("Converte_TipoContaCcl", vCampo, sTipoContaCcl)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoContaCcl
        
                Case ADM_SUBTIPO_NATUREZACONTA
                    Call CF("Converte_NaturezaConta", vCampo, sNaturezaConta)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sNaturezaConta
        
                Case ADM_SUBTIPO_TIPORATEIO
                    Call Converte_TipoRateio(vCampo, sTipoRateio)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoRateio
        
                Case ADM_SUBTIPO_EXERCICIO
                
                    iExercicio = vCampo
                
                    'le o exercicio do BD e exibe o nome externo
                    lErro = CF("Exercicio_Le_NomeExterno", objBrowse.lComando2, iExercicio, sNomeExterno)
                    If lErro <> SUCESSO Then Error 9261
                    
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sNomeExterno
                
                Case ADM_SUBTIPO_PERIODO
                
                    'indica que na proxima passagem do for o campo não deve ser exibido
                    iProximoCampo = 1
                    
                    'coloca o exercicio neste campo
                    iExercicio = colRegCampo.Item(iIndice + 1)
                    iPeriodo = vCampo
                
                    'le o periodo do BD e exibe o nome externo
                    lErro = CF("Periodo_Le_NomeExterno", objBrowse.lComando3, iExercicio, iPeriodo, sNomeExterno)
                    If lErro <> SUCESSO Then Error 9269
                                            
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sNomeExterno
        
                Case ADM_SUBTIPO_AGLUTINA
                    Call Converte_Aglutina(vCampo, sAglutina)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sAglutina
      
                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
    
        Case Else
            
            Select Case objValorCampo.iSubtipo
        
                Case ADM_SUBTIPO_CONTA
                
                    If Len(vCampo) > 0 Then
                
                        'se o campo for uma conta, mascara a conta antes de exibi-la
                        sContaMascarada = String(STRING_CONTA, 0)
                
                        sConta = vCampo
                    
                        lErro = Mascara_MascararConta(sConta, sContaMascarada)
                        If lErro <> SUCESSO Then Error 9259
                        
                    Else
                        sContaMascarada = ""
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sContaMascarada
        
                Case ADM_SUBTIPO_CCL
            
                    If Len(vCampo) > 0 Then
            
                        'mascara o centro de custo
                        sCclMascarado = String(STRING_CCL, 0)
                    
                        sCcl = vCampo
                        
                        lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
                        If lErro Then Error 9260
                
                    Else
                    
                        sCclMascarado = ""
                        
                    End If
        
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCclMascarado
                    
                    
                Case Else
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
            
    End Select
    
    Browse_Formata_Campo = SUCESSO
    
    Exit Function
    
Erro_Browse_Formata_Campo:

    Browse_Formata_Campo = Err

    Select Case Err
        
        Case 9259
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, CStr(vCampo))

        Case 9260
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, CStr(vCampo))
            
        Case 9261, 9269
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150479)

    End Select

    Exit Function

End Function

''Function Formata_Contas_Inicial_Final(sCtaMasc_I As String, sCtaMasc_F As String, sCtaNaoMasc_I As String, sCtaNaoMasc_F As String) As Long
'''tranforma contas do formato externo (com delimitadores) para o formato do bd
'''Verifica se a conta inicial é maior que a conta final
'''retorna 7071 p/erro na cta inicial, 7072 p/erro na cta final
'''7073 p/cta inicial > final
''Dim iCtaPreenchida_I As Integer, iCtaPreenchida_F As Integer
''Dim lErro As Long
''
''On Error GoTo Erro_Formata_Contas_Inicial_Final
''
''    'formata a Conta Inicial
''    lErro = CF("Conta_Formata",sCtaMasc_I, sCtaNaoMasc_I, iCtaPreenchida_I)
''    If lErro <> SUCESSO Then Error 7071
''
''    'formata a Conta Final
''    lErro = CF("Conta_Formata",sCtaMasc_F, sCtaNaoMasc_F, iCtaPreenchida_F)
''    If lErro <> SUCESSO Then Error 7072
''
''    'se ambas as contas estão preenchidas, a conta inicial não pode ser maior que a final
''    If iCtaPreenchida_I = CONTA_PREENCHIDA And iCtaPreenchida_F = CONTA_PREENCHIDA Then
''        If sCta_I > sCta_F Then Error 7073
''    End If
''
''    Formata_Contas_Inicial_Final = SUCESSO
''
''    Exit Function
''
''Erro_Formata_Contas_Inicial_Final:
''
''    Formata_Contas_Inicial_Final = Err
''
''    Select Case Err
''
''        Case 7071
''        Case 7072
''        Case 7073
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150480)
''
''    End Select
''
''    Exit Function
''
''End Function

Sub Converte_TipoContaCcl(ByVal iTipo As Integer, sTipo As String)
'converte o tipo de conta/centro de custo do valor inteiro para o seu correspondente textual (sintetico/analitico)

    If iTipo = CONTA_SINTETICA Or iTipo = CCL_SINTETICA Then
        sTipo = "Sintético"
    ElseIf iTipo = CONTA_ANALITICA Or iTipo = CCL_ANALITICA Then
        sTipo = "Analítico"
    End If

End Sub

Sub Converte_NaturezaConta(ByVal iNatureza As Integer, sNatureza As String)
'converte a natureza de conta do valor inteiro para o seu correspondente textual (credito/debito)

    If iNatureza = CONTA_CREDITO Then
        sNatureza = "Crédito"
    ElseIf iNatureza = CONTA_DEBITO Then
        sNatureza = "Débito"
    End If

End Sub

Sub Converte_Aglutina(ByVal iAglutina As Integer, sAglutina As String)
'converte a natureza de conta do valor inteiro para o seu correspondente textual (credito/debito)

    If iAglutina = LANCAMENTO_NAO_AGLUTINADO Then
        sAglutina = "Não Aglutinado"
    ElseIf iAglutina = LANCAMENTO_AGLUTINA Then
        sAglutina = "Aglutinado"
    End If

End Sub

Function UM_Conversao_Trans(ByVal iClasseUM As Integer, ByVal sUnidade1 As String, ByVal sUnidade2 As String, dFator As Double) As Long
'Calcula o fator de conversão entre 2 Unidades de Medida
'De sUnidade1 para sUnidade2

'??? se for a mesma unidade retornar 1,
'??? se uma das unidades for a base basta uma leitura

Dim lErro As Long
Dim objUnidade1 As New ClassUnidadeDeMedida
Dim objUnidade2 As New ClassUnidadeDeMedida

On Error GoTo Erro_UM_Conversao_Trans

    'Se forem iguais retorna 1
    If sUnidade1 = sUnidade2 Then
    
        dFator = 1
         
    Else 'Se forem diferentes.....
        
        'Preenche chave de objUnidade1
        objUnidade1.iClasse = iClasseUM
        objUnidade1.sSigla = sUnidade1
    
        'Lê objUnidade1
        lErro = CF("UM_Le", objUnidade1)
        If lErro <> SUCESSO And lErro <> 23775 Then Error 23770
        
        If lErro = 23775 Then Error 55990
        
        'Se Editavel de objUnidade1 for UM_EDITAVEL, atualiza objUnidade1 no BD para Editavel=UM_NAO_EDITAVEL
        If objUnidade1.iEditavel = UM_EDITAVEL Then
                        
            'Chama UM_Grava_NaoEditavel(objUnidade1)
            lErro = UM_Grava_NaoEditavel_Trans(objUnidade1)
            If lErro <> SUCESSO Then Error 33568
            
        End If
        
        'Preenche chave de objUnidade2
        objUnidade2.iClasse = iClasseUM
        objUnidade2.sSigla = sUnidade2
    
        'Lê objUnidade2
        lErro = CF("UM_Le", objUnidade2)
        If lErro <> SUCESSO And lErro <> 23775 Then Error 23771
    
        If lErro = 23775 Then Error 55991
    
        'Se Editavel de objUnidade2 for UM_EDITAVEL, atualiza objUnidade2 no BD para Editavel=UM_NAO_EDITAVEL
        If objUnidade2.iEditavel = UM_EDITAVEL Then
        
            'Chama UM_Grava_NaoEditavel(objUnidade2)
            lErro = UM_Grava_NaoEditavel_Trans(objUnidade2)
            If lErro <> SUCESSO Then Error 33569
            
        End If
        
        'Calcula dFator
        dFator = (objUnidade1.dQuantidade) / (objUnidade2.dQuantidade)

    End If
    
    UM_Conversao_Trans = SUCESSO

    Exit Function

Erro_UM_Conversao_Trans:

    UM_Conversao_Trans = Err

    Select Case Err

        Case 23770, 23771, 33568, 33569

        Case 55990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA", Err, objUnidade1.iClasse, objUnidade1.sSigla)

        Case 55991
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA", Err, objUnidade2.iClasse, objUnidade2.sSigla)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 150481)

    End Select

    Exit Function

End Function

Private Function UM_Grava_NaoEditavel_Trans(ByVal objUnidadeMedida As ClassUnidadeDeMedida) As Long
'Atualiza o campo Editável da tabela UnidadesDeMedida (em Transação)

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim iEditavel As Integer

On Error GoTo Erro_UM_Grava_NaoEditavel_Trans

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33570

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 33571

    'se ainda não está em transacao ==> abre a transacao
    If GL_lTransacao = 0 Then

        'Inicia a Transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then Error 33572

    End If
    
    'Pesquisa no BD a Unidade de Medida em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Editavel FROM UnidadesdeMedida WHERE Classe=? AND Sigla=?", 0, iEditavel, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 33573

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33574

    If lErro = AD_SQL_SEM_DADOS Then Error 33575
    
    If iEditavel = UM_EDITAVEL Then
             
        'Alterar o campo editável p/ não editável
        lErro = Comando_ExecutarPos(lComando1, "UPDATE UnidadesDeMedida SET Editavel = ?", lComando, UM_NAO_EDITAVEL)
        If lErro <> AD_SQL_SUCESSO Then Error 33576
                
    End If
    
    'se abriu a transacao nesta funcao ==> pode fazer o commit
    If lTransacao <> 0 Then
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then Error 33577
        
    End If

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    UM_Grava_NaoEditavel_Trans = SUCESSO

    Exit Function

Erro_UM_Grava_NaoEditavel_Trans:

    UM_Grava_NaoEditavel_Trans = Err

    Select Case Err

        Case 33570, 33571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 33572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 33573, 33574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_UNIDADESDEMEDIDA", Err, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)

        Case 33575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA", Err, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)

        Case 33576
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_UNIDADESDEMEDIDA", Err)
        
        Case 33577
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150482)

    End Select

    'se abriu a transacao nesta funcao ==> pode fazer o rollback
    If lTransacao <> 0 Then Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function
        
End Function

Sub Converte_TipoRateio(ByVal iTipoRateio As Integer, sTipoRateio As String)
'converte o tipo de rateio do valor inteiro para o seu correspondente textual (credito/debito)

    If iTipoRateio = RATEIO_MENSAL Then
        sTipoRateio = "Mensal"
    ElseIf iTipoRateio = RATEIO_ACUMULADO Then
        sTipoRateio = "Períodos Acumulados"
    End If

End Sub

