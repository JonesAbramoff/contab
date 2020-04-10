VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCOMFormata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Dim glFornecedor As Long

Function Browse_Formata_Campo(objBrowse As AdmBrowse, colRegCampo As Collection, objValorCampo As AdmValorCampo, vCampo As Variant, iProximoCampo As Integer, iLinha As Integer, iIndice As Integer, iExercicio As Integer, iPeriodo As Integer, sNomeExterno As String) As Long

Dim sContaMascarada As String
Dim sCclMascarado As String
Dim sConta As String
Dim sCcl As String
Dim lErro As Long
Dim sData As String
Dim sTipoContaCcl As String
Dim sNaturezaConta As String
Dim sNomeFilial As String
Dim objFilialFornecedor As ClassFilialFornecedor
Dim objFornecedor As ClassFornecedor
Dim objCondicaoPagto As ClassCondicaoPagto
Dim objComprador As ClassComprador
Dim dValorMensal As Double
Dim sProdutoMascarado As String
Dim sProduto As String
Dim objAdmFiliais As New AdmFiliais
Dim sPodeSerCompravel As String
Dim sGerencial As String
Dim sNatureza As String

On Error GoTo Erro_Browse_Formata_Campo

    Select Case objValorCampo.iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_PERCENTUAL
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Percent")
        
                Case ADM_SUBTIPO_HORA
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CDate(vCampo)
        
                Case ADM_SUBTIPO_LIBERADO
                                                
                    lErro = AlcadaUsuario_Le_ValorMensal(dValorMensal)
                    If lErro <> SUCESSO Then gError 91600
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(dValorMensal, "Standard")
                    
                Case ADM_SUBTIPO_PRECOUNITARIO
                    If Len(Trim(gobjCOM.sFormatoPrecoUnitario)) > 0 Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, gobjCOM.sFormatoPrecoUnitario)
                    Else
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Standard")
                    End If
                    
                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Standard")
                    
            End Select
                            
        Case ADM_TIPO_DATE
            sData = Format(vCampo, "dd/mm/yyyy")
        
            If sData = "07/09/1822" Then sData = ""
        
            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sData
    
        Case ADM_TIPO_SMALLINT
    
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_TIPOCONTACCL
                    Call CF("Converte_TipoContaCcl", vCampo, sTipoContaCcl)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoContaCcl
        
                Case ADM_SUBTIPO_NATUREZACONTA
                    Call CF("Converte_NaturezaConta", vCampo, sNaturezaConta)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sNaturezaConta
                        
                Case ADM_SUBTIPO_COMPRAS 'William
                    Call Retorna_Tipo_Compravel(vCampo, sPodeSerCompravel)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sPodeSerCompravel
                        
                Case ADM_SUBTIPO_GERENCIAL 'William
                    Call Retorna_Tipo_Gerencial(vCampo, sGerencial)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sGerencial
                
                Case ADM_SUBTIPO_NATUREZA 'William
                    Call Retorna_Tipo_Natureza(vCampo, sNatureza)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sNatureza
                          
                Case ADM_SUBTIPO_CONDPAGTO
                    If Len(vCampo) > 0 Then
                        Set objCondicaoPagto = New ClassCondicaoPagto
                        objCondicaoPagto.iCodigo = CInt(vCampo)
                        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
                        If lErro <> SUCESSO Then gError 87552
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = objCondicaoPagto.sDescReduzida
                    End If
                                       
                Case ADM_SUBTIPO_COMPRADOR
                    If Len(vCampo) > 0 Then
                        Set objComprador = New ClassComprador
                        objComprador.iCodigo = CInt(vCampo)
                        objComprador.iFilialEmpresa = giFilialEmpresa
                        lErro = CF("Comprador_Le", objComprador)
                        If lErro <> SUCESSO Then gError 87553
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = objComprador.sCodUsuario
                    End If
                                            
                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
    
        Case ADM_TIPO_INTEGER
    
            Select Case objValorCampo.iSubTipo
                            
                '?????? Substituir por Constantes. Mario - Ok
                Case ADM_SUBTIPO_URGENTE
                    If vCampo = 0 Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = ADM_SUBTIPO_NAO
                    ElseIf vCampo = 1 Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = ADM_SUBTIPO_SIM
                    End If

                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select

        Case Else
            
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_TIPOFRETE
                    If Len(Trim(vCampo)) >= 0 Then
                        
                        Select Case vCampo
                        
                            '???? Substituir por Constantes. Mario
                            Case TIPO_CIF
                                objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "CIF"
                            
                            Case TIPO_FOB
                                objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = "FOB"
                            
                        End Select
                        
                    End If
        
                Case ADM_SUBTIPO_CONTA
                
                    If Len(vCampo) > 0 Then
                    
                        'se o campo for uma conta, mascara a conta antes de exibi-la
                        sContaMascarada = String(STRING_CONTA, 0)
                
                        sConta = vCampo
                    
                        lErro = Mascara_MascararConta(sConta, sContaMascarada)
                        If lErro <> SUCESSO Then gError 57244
                        
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
                        If lErro <> SUCESSO Then gError 57245
                        
                    Else
                    
                        sCclMascarado = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCclMascarado
                    
                Case ADM_SUBTIPO_PRODUTO
                    
                    If Len(vCampo) > 0 Then
                        
                        sProdutoMascarado = String(STRING_PRODUTO, 0)
                    
                        sProduto = vCampo
                            
                        lErro = Mascara_MascararProduto(sProduto, sProdutoMascarado)
                        If lErro <> SUCESSO Then gError 93396
                    
                    Else
                        
                        sProdutoMascarado = ""
                            
                    End If
                    
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sProdutoMascarado
                                            
                Case Else
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
            
    End Select
    
    Browse_Formata_Campo = SUCESSO
    
    Exit Function
    
Erro_Browse_Formata_Campo:

    Browse_Formata_Campo = gErr

    Select Case gErr
        
        Case 57244
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", gErr, CStr(vCampo))

        Case 57245
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, CStr(vCampo))
            
        Case 87548, 87550, 87552, 87553, 91600, 93396, 93397
            'Tratado na rotina chamadora
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147437)

    End Select

    Exit Function

End Function


Public Function AlcadaUsuario_Le_ValorMensal(dValorMensal As Double) As Long
'Le o Valor do mês Corrente do ano atual, monta o Select dinamicamente

Dim lErro As Long
Dim lComando As Long
Dim dValor As Double
Dim iAno As Integer
Dim iMes As Integer
Dim sMes As String
Dim sSelect As String

On Error GoTo Erro_AlcadaUsuario_Le_ValorMensal

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 91601
    
    'Recolhe o mês e o ano Atual
    iAno = Year(gdtDataAtual)
    iMes = Month(gdtDataAtual)
    
    'Concatena o mês para formar o nome da coluna da view a ser lida
    sMes = "Liberado" & CStr(iMes)
    
    'Monta o Select
    sSelect = "SELECT " & sMes & " FROM AlcadaUsuario WHERE Ano = ?"
    
    'Faz a leitura de acordo com o Select Montado
    lErro = Comando_Executar(lComando, sSelect, dValorMensal, iAno)
    If lErro <> AD_SQL_SUCESSO Then gError 91602
    
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91603
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    AlcadaUsuario_Le_ValorMensal = SUCESSO
    
    Exit Function
    
Erro_AlcadaUsuario_Le_ValorMensal:

    AlcadaUsuario_Le_ValorMensal = gErr
    
    Select Case gErr
    
        Case 91601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 91602, 91603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALCADAUSUARIO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147438)
            
    End Select
    
    'Fecha o Comando ---> saída por erro
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


'William
Public Function Retorna_Tipo_Compravel(vCampo As Variant, sPodeSerCompravel As String)

    If Len(Trim(vCampo)) = 0 Then
        sPodeSerCompravel = ADM_NAO_PARTICIPA_COMPRAS
    Else
        sPodeSerCompravel = ADM_PARTICIPA_COMPRAS
    End If
        
End Function

'William
Public Function Retorna_Tipo_Gerencial(vCampo As Variant, sGerencial As String)

    If Len(Trim(vCampo)) = 0 Then
        sGerencial = ADM_ANALITICO
    Else
        sGerencial = ADM_GERENCIAL
    End If
        
End Function

'William
Public Function Retorna_Tipo_Natureza(vCampo As Variant, sNatureza As String)

    If Len(vCampo) > 0 Then
    
        Select Case vCampo

            Case ADM_SUBTIPO_MATERIAPRIMA
                sNatureza = ADM_TIPO_MATERIAPRIMA
            
            Case ADM_SUBTIPO_PRODUTOINTERMD
                sNatureza = ADM_TIPO_INTERMEDIARIO
            
            Case ADM_SUBTIPO_EMBALAGEM
                sNatureza = ADM_TIPO_EMBALAGEM
            
            Case ADM_SUBTIPO_PRODUTOACABADO
                sNatureza = ADM_TIPO_ACABADO
            
            Case ADM_SUBTIPO_PRODUTOREVENDA
                sNatureza = ADM_TIPO_REVENDA
            
            Case ADM_SUBTIPO_PRODUTOREPARO
                sNatureza = ADM_TIPO_REPARO
            
            Case ADM_SUBTIPO_OUTROS
                sNatureza = ADM_TIPO_OUTROS

        End Select
    
    End If

End Function


