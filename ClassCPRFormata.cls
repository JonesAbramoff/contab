VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRFormata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Browse_Formata_Campo(objBrowse As AdmBrowse, ByVal colRegCampo As Collection, objValorCampo As AdmValorCampo, vCampo As Variant, iProximoCampo As Integer, iLinha As Integer, iIndice As Integer, iExercicio As Integer, iPeriodo As Integer, sNomeExterno As String, Optional iVermelho As Integer = -1) As Long

Dim sContaMascarada As String
Dim sCclMascarado As String
Dim sCodigoMascarado As String
Dim sConta As String
Dim sCcl As String
Dim sCodigo As String
Dim lErro As Long
Dim sData As String
Dim sTipoContaCcl As String
Dim sNaturezaConta As String
Dim sTipoMeioPagto As String
Dim sCgc As String
Dim sStatusTitPag As String, sStatus As String

On Error GoTo Erro_Browse_Formata_Campo

        
    If giCliAtrasoDestacar <> 0 Then
    
        If iVermelho <> -1 Then
        
            objBrowse.objGrid.Row = iLinha
            objBrowse.objGrid.Col = objValorCampo.iPosicaoTela - 1
        
            'VERMELHO
            If iVermelho = 1 Then
        
                objBrowse.objGrid.CellForeColor = vbRed
        
            Else
            
                objBrowse.objGrid.CellForeColor = vbWindowText
            
            End If

        End If

    End If

    Select Case objValorCampo.iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case objValorCampo.iSubTipo
        
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
    
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_TIPOCONTACCL
                    Call CF("Converte_TipoContaCcl", vCampo, sTipoContaCcl)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoContaCcl
        
                Case ADM_SUBTIPO_NATUREZACONTA
                    Call CF("Converte_NaturezaConta", vCampo, sNaturezaConta)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sNaturezaConta
        
                
                Case ADM_SUBTIPO_TIPOMEIOPAGTO
                    Call Converte_TipoMeioPagto(vCampo, sTipoMeioPagto)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoMeioPagto

                Case ADM_SUBTIPO_STATUS
                    Call CF("Converte_Status", vCampo, sStatus)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sStatus

                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
    
        Case Else
            
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_CONTA
                
                    If Len(vCampo) > 0 Then
                    
                        'se o campo for uma conta, mascara a conta antes de exibi-la
                        sContaMascarada = String(STRING_CONTA, 0)
                
                        sConta = vCampo
                    
                        lErro = Mascara_MascararConta(sConta, sContaMascarada)
                        If lErro <> SUCESSO Then Error 10038
                        
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
                        If lErro <> SUCESSO Then Error 10039
                        
                    Else
                    
                        sCclMascarado = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCclMascarado
                                                        
                Case ADM_SUBTIPO_CODIGONATMOVCTA

                    If Len(vCampo) > 0 Then

                        'mascara o Codigo
                        sCodigoMascarado = String(STRING_NATMOVCTA_CODIGO, 0)

                        sCodigo = vCampo

                        lErro = Mascara_MascararItem(SEGMENTO_NATMOVCTA, sCodigo, sCodigoMascarado)
                        If lErro <> SUCESSO Then Error 10040

                    Else

                        sCodigoMascarado = ""

                    End If

                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCodigoMascarado
                
                Case ADM_SUBTIPO_CGC
                
                    Select Case Len(Trim(vCampo))
                    
                        Case STRING_CPF 'CPF
                                        
                            sCgc = Format(vCampo, "000\.000\.000-00; ; ; ")
                    
                        Case STRING_CGC 'CGC
                    
                            sCgc = Format(vCampo, "00\.000\.000\/0000-00; ; ; ")
                            
                    End Select
                    
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCgc
                
                '#########################################
                'Inserido por Wagner
                Case ADM_SUBTIPO_STATUS_TITPAG
                
                    If Len(vCampo) > 0 Then
            
                        Select Case vCampo
                        
                            Case TITULOPAG_CHEQUE_STATUS_ABERTO
                                sStatusTitPag = TITULOPAG_CHEQUE_STATUS_ABERTO & SEPARADOR & STRING_TITULOPAG_CHEQUE_STATUS_ABERTO
                            
                            Case TITULOPAG_CHEQUE_STATUS_LIQUIDADO
                                sStatusTitPag = TITULOPAG_CHEQUE_STATUS_LIQUIDADO & SEPARADOR & STRING_TITULOPAG_CHEQUE_STATUS_LIQUIDADO
                   
                            Case TITULOPAG_CHEQUE_STATUS_PAGO
                               sStatusTitPag = TITULOPAG_CHEQUE_STATUS_LIQUIDADO & SEPARADOR & STRING_TITULOPAG_CHEQUE_STATUS_PAGO
                        
                        End Select
                        
                    Else
                    
                        sStatusTitPag = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sStatusTitPag
                '#########################################
                
                Case Else
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
            
    End Select
    
    Browse_Formata_Campo = SUCESSO
    
    Exit Function
    
Erro_Browse_Formata_Campo:

    Browse_Formata_Campo = Err

    Select Case Err
        
        Case 10038
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, CStr(vCampo))

        Case 10039
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, CStr(vCampo))
            
        Case 10040
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararNatMovCta", Err, CStr(vCampo))
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147902)

    End Select

    Exit Function

End Function

Private Sub Converte_TipoMeioPagto(ByVal iTipo As Integer, sTipo As String)
'converte o tipo meio de pagamento do valor inteiro para o seu correspondente textual (dinheiro/cheque/bordero)

    If iTipo = DINHEIRO Then
        sTipo = "Dinheiro"
    ElseIf iTipo = Cheque Then
        sTipo = "Cheque"
    ElseIf iTipo = BORDERO Then
        sTipo = "Borderô"
    End If

End Sub


'Preenche ComboBox Filial com códigos-nome em colCodigoNome
Sub Filial_Preenche(ByVal FilialCombo As Object, ByVal colCodigoNome As AdmColCodigoNome)
            
Dim Filial As ComboBox
Dim objCodigoNome As AdmCodigoNome
Dim sComboBoxItem As String
Dim lErro As Long
    
On Error GoTo Erro_Filial_Preenche
    
    Set Filial = FilialCombo
    
    Filial.Clear
    For Each objCodigoNome In colCodigoNome
    
        sComboBoxItem = CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Filial.AddItem sComboBoxItem
        Filial.ItemData(Filial.NewIndex) = objCodigoNome.iCodigo
                
    Next
    
    Exit Sub
    
Erro_Filial_Preenche:

    lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147903)

    Exit Sub

End Sub
'Seleciona Filial com código iCodFilial na Combo FilialCombo
Sub Filial_Seleciona(ByVal FilialCombo As Object, ByVal iCodFilial As Integer)

Dim Filial As ComboBox
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Filial_Seleciona
    
    Set Filial = FilialCombo
    
    For iIndice = 0 To Filial.ListCount - 1
    
        If Filial.ItemData(iIndice) = iCodFilial Then
            Filial.ListIndex = iIndice
            Exit For
        End If
        
    Next
    
    Exit Sub
    
Erro_Filial_Seleciona:

    lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147904)

    Exit Sub
        
End Sub


Private Function FormataCpoTexto(ByVal sTexto As String, ByVal iTam As Integer) As String
'formata campo texto completando com brancos à direita

Dim iTexto As Integer

    iTexto = Len(sTexto)

    If iTexto >= iTam Then
        FormataCpoTexto = Mid$(sTexto, 1, iTam)
    Else
        FormataCpoTexto = sTexto & Space(iTam - iTexto)
    End If

    Exit Function

End Function


Private Function FormataCpoValor(vData As Variant, iTam As Integer) As String
'formata campo Valor alinhado-o à direita sem ponto e decimais, com zeros a esquerda

Dim sValor As String
Dim iTamanhoValor As Integer
Dim dValor As Double
Dim sDecimais As String
Dim sInteiro As String
Dim sNumero As String
Dim sEspacos As String
Dim iData As Integer

    sValor = Format(vData, "Fixed")
    sDecimais = Mid(sValor, Len(sValor) - 1, 2)
    sInteiro = Mid(sValor, 1, Len(sValor) - 3)
    sNumero = sInteiro & sDecimais

    iData = iTam - Len(sNumero)
    sEspacos = String(iData, "0")

    FormataCpoValor = sEspacos & sInteiro & sDecimais

End Function
Private Function FormataCpoNum(ByVal vData As Variant, ByVal iTam As Integer) As String
'formata campo numerico alinhado-o à direita sem ponto e decimais, com zeros a esquerda

Dim iData As Integer
Dim sData As String

    If Len(vData) = iTam Then

        FormataCpoNum = vData
        Exit Function

    End If

    iData = iTam - Len(vData)
    
    If iData > 0 Then sData = String(iData, "0")

    FormataCpoNum = sData & vData

    Exit Function

End Function
Function Formata_Cep(ByVal vCep As Variant, ByVal iTamanho As Integer) As String
'Formata o Cep passado completando com zeros no final até atingir o tamanho indicado

Dim iDiferenca As Integer

    iDiferenca = iTamanho - Len(vCep)

    If iDiferenca > 0 Then
        Formata_Cep = vCep & FormataCpoNum("", iDiferenca)
    Else
        Formata_Cep = vCep
    End If

    Exit Function

End Function


