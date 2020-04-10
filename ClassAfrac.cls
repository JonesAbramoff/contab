VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAfrac"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'C:\acbr3wc\trunk2\Fontes\ACBrSerial

Public Sub Afrac_UF_ObtemPerfil(sPerfil As String)

Dim X As New ClassECFConfig
    
    Select Case X.gsUF
    
        Case "AM", "PE", "RJ", "DF", "GO", "MA", "MS", "RR", "SC", "TO"
        
            sPerfil = "W"
            
        Case "BA"
            sPerfil = "Z"
        
        Case "PB", "ES"
            sPerfil = "Y"
            
        Case "MG"
            sPerfil = "X"
        
        Case Else
            sPerfil = "ND"
        
    End Select
 
End Sub

Public Function AFRAC_TipoMeioPagtoTicket(ByVal iTipoMeioPagto As Integer) As Boolean

    Select Case iTipoMeioPagto
    
        Case TIPOMEIOPAGTOLOJA_VALE_TICKET, TIPOMEIOPAGTOLOJA_VALE_REFEICAO, TIPOMEIOPAGTOLOJA_VALE_PRESENTE, TIPOMEIOPAGTOLOJA_VALE_COMBUSTIVEL
            AFRAC_TipoMeioPagtoTicket = True
        
        Case Else
            AFRAC_TipoMeioPagtoTicket = False
            
    End Select

End Function

Public Function AFRAC_ImpressoraCFe(ByVal iModeloECF As Integer) As Boolean
'informa se a "impressora" corresponde à cupom fiscal eletronico (SAT ou NFCe)

    Select Case iModeloECF
            
        Case IMPRESSORA_NFCE, IMPRESSORA_SAT_2_5_15, IMPRESSORA_NFE
            AFRAC_ImpressoraCFe = True
    
        Case Else
            AFRAC_ImpressoraCFe = False
        
    End Select
    
End Function

Public Function AFRAC_FechaPortaSerial() As Integer
'
'Fecha a porta serial de comunicação entre o micro e a impressora
Dim X As New ClassECFConfig
Dim sRetorno As String
Dim iStatus As Integer

On Error GoTo Erro_AFRAC_FechaPortaSerial

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            AFRAC_FechaPortaSerial = Bematech_FI_FechaPortaSerial()

            If AFRAC_FechaPortaSerial = 1 Then
                AFRAC_FechaPortaSerial = SUCESSO
            ElseIf AFRAC_FechaPortaSerial = 0 Then
                AFRAC_FechaPortaSerial = 1
            End If
        
        Case IMPRESSORA_DARUMA_FS600
        
'                AFRAC_FechaPortaSerial = Daruma_FI_FechaPortaSerial()
            AFRAC_FechaPortaSerial = eFecharSerial_Daruma()

            If AFRAC_FechaPortaSerial = 1 Then
                AFRAC_FechaPortaSerial = SUCESSO
            ElseIf AFRAC_FechaPortaSerial = 0 Then
                AFRAC_FechaPortaSerial = 1
            End If

        Case IMPRESSORA_ELGIN
                Call Elgin_CloseCif
                
                sRetorno = Space(50)
                iStatus = Elgin_EsperaResposta(sRetorno)
                
                If iStatus = SUCESSO Then
                    AFRAC_FechaPortaSerial = SUCESSO
                Else
                    AFRAC_FechaPortaSerial = 1
                End If
                
        Case Else
            AFRAC_FechaPortaSerial = SUCESSO
                

    End Select

   Exit Function
   
Erro_AFRAC_FechaPortaSerial:

    AFRAC_FechaPortaSerial = 1
    
    Exit Function


End Function

Public Function AFRAC_AbrePortaSerial() As Integer

'Abre a porta serial de comunicação entre o micro e a impressora
Dim X As New ClassECFConfig
Dim sRetorno As String
Dim iStatus As Integer

On Error GoTo Erro_AFRAC_AbrePortaSerial

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            AFRAC_AbrePortaSerial = Bematech_FI_AbrePortaSerial()
            
            If AFRAC_AbrePortaSerial = 1 Then
                AFRAC_AbrePortaSerial = SUCESSO
            ElseIf AFRAC_AbrePortaSerial = 0 Then
                AFRAC_AbrePortaSerial = 1
            End If
            

            
        Case IMPRESSORA_DARUMA_FS600
            
'                AFRAC_AbrePortaSerial = Daruma_FI_AbrePortaSerial()
            AFRAC_AbrePortaSerial = eDefinirProduto_Daruma("ECF")
            If AFRAC_AbrePortaSerial = 1 Then
                AFRAC_AbrePortaSerial = eBuscarPortaVelocidade_ECF_Daruma()
            End If
            
            If AFRAC_AbrePortaSerial = 1 Then
                AFRAC_AbrePortaSerial = SUCESSO
            ElseIf AFRAC_AbrePortaSerial = 0 Then
                AFRAC_AbrePortaSerial = 1
            End If
            
        Case IMPRESSORA_ELGIN
                AFRAC_AbrePortaSerial = Elgin_OpenCif
            
                sRetorno = Space(50)
                iStatus = Elgin_EsperaResposta(sRetorno)
            
                If iStatus = SUCESSO Then
                    AFRAC_AbrePortaSerial = SUCESSO
                Else
                    AFRAC_AbrePortaSerial = 1
                End If
            
        Case Else
            AFRAC_AbrePortaSerial = SUCESSO
            
    End Select

   Exit Function
   
Erro_AFRAC_AbrePortaSerial:

    AFRAC_AbrePortaSerial = 1
    
    Exit Function

End Function

Public Function AFRAC_LerIndiceRelatorioGerencial(sIndice As String, sDescricao As String) As Integer

'Le o índice do tipo de relatorio gerencial

    AFRAC_LerIndiceRelatorioGerencial = SUCESSO


End Function


Public Function AFRAC_ProgramarTributacao(sTributacao As String) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim iFlag As Integer
Dim sBuffer As String
Dim sBuffer1 As String
Dim iIndice As Integer
Dim sRetorno As String
Dim iStatus1 As Integer
Dim sFlags As String
Dim iISS As Integer
Dim sAliquota As String

'Programa uma alíquota de ICMS ou ISS
Dim sICMS_ISS As String

On Error GoTo Erro_AFRAC_ProgramarTributacao

    AFRAC_ProgramarTributacao = 1

    Select Case X.giCodModeloECF

        Case 1 To 2, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            iStatus = Bematech_FI_FlagsFiscais(iFlag)
        
            'se retorno ok e o cupom nao estiver aberto
            If iStatus = 1 And Not (iFlag And 1) Then
        
                If left(sTributacao, 1) = "S" Then
                    iISS = 1
                Else
                    iISS = 0
                End If
        
                sAliquota = Mid(sTributacao, 4, 4)
        
                AFRAC_ProgramarTributacao = Bematech_FI_ProgramaAliquota(sAliquota, CStr(iISS))
    
                If AFRAC_ProgramarTributacao = 1 Then
                    AFRAC_ProgramarTributacao = SUCESSO
                ElseIf AFRAC_ProgramarTributacao = 0 Then
                    AFRAC_ProgramarTributacao = 1
                End If
            End If
           
        Case IMPRESSORA_DARUMA_FS600
        
 '               iStatus = Daruma_FI_FlagsFiscais(iFlag)
 
            sFlags = Space(18)
                
            iStatus = rStatusImpressoraBinario_ECF_Daruma(sFlags)
            If iStatus = 1 Then
                iFlag = CInt(Mid(sFlags, 12, 1))
            End If
        
            'se retorno ok e o cupom nao estiver aberto
            If iStatus = 1 And Not (iFlag And 1) Then
        
                '  AFRAC_ProgramarTributacao = Daruma_FI_ProgramaAliquota(sTributacao, CInt(sICMS_ISS))
                
                If Mid(sTributacao, 1, 1) = "T" Then sTributacao = Mid(sTributacao, 2)
                AFRAC_ProgramarTributacao = confCadastrar_ECF_Daruma("ALIQUOTA", sTributacao, "|")
    
                If AFRAC_ProgramarTributacao = 1 Then
                    AFRAC_ProgramarTributacao = SUCESSO
                ElseIf AFRAC_ProgramarTributacao = 0 Then
                    AFRAC_ProgramarTributacao = 1
                End If
                
            End If
           
        Case IMPRESSORA_ELGIN
            
            sBuffer = Space(41)
            
            iStatus = Elgin_TransStatus(0, sBuffer)
            
            sRetorno = Space(50)
            iStatus1 = Elgin_EsperaResposta(sRetorno)
            
            If Mid(sTributacao, 1, 1) = "T" Then
                sICMS_ISS = 0
            Else
                sICMS_ISS = 1
            End If
            
            sTributacao = Mid(sTributacao, 4, 4)
            
            'testa se o cupom esta fechado
            If iStatus = SUCESSO And Mid(sBuffer, 8, 1) <> "1" Then
            
                iStatus = Elgin_TransTabAliquotas()
                
                sBuffer1 = Space(100)
                iStatus = Elgin_EsperaResposta(sBuffer1)
                                        
                If iStatus = SUCESSO Then
                    sBuffer1 = Mid(sBuffer1, 6)
                    For iIndice = 1 To 15
                        If Mid(sBuffer1, (iIndice * 4), 4) = "0000" Then
                            iStatus = Elgin_ProgAliquotasICMS_ISS("T" & Format(iIndice, "00"), sTributacao, CByte(sICMS_ISS))
                            
                            sRetorno = Space(50)
                            iStatus = Elgin_EsperaResposta(sRetorno)
                            
                            Exit For
                        End If
                    Next
                End If
                    
            End If
            
            If iStatus = SUCESSO Or iStatus = -18 Then
                AFRAC_ProgramarTributacao = SUCESSO
            Else
                AFRAC_ProgramarTributacao = 1
            End If
            
            
            'EM PRINCIPIO NAO VAMOS FAZER NADA POIS ESTA DLL NAO TEM OUTRAS FUNCOES NECESSARIAS PARA PROGRAMAR A TRIBUTACAO
            'AFRAC_ProgramarTributacao = Elgin_ProgAliquotas(Mid(sTributacao, 1, 3), Mid(sTributacao, 4, 4))
           'ProgAliquotas("T07","1500")
'            trib: Indica o tipo e a posição do totalizador parcial a ser programado. Este parâmetro deve
'            respeitar o seguinte formato: <"Xnn>, onde:
'            · X – Representa o tipo do totalizador: (1c)
'            o T – Tributado;
'            o I – Isento;
'            o F – Substituição tributária e
'            o N – Não incidência.
'            · nn – Representa a posição do totalizador, ou alíquota, "00" a "15". (2n)
'            valor: Valor da alíquota; este parâmetro deve respeitar o seguinte formato: nnnn. (4n) Por
'            exemplo: "1000" – alíquota de dez porcento.
           
           
           
        Case Else
            AFRAC_ProgramarTributacao = SUCESSO
        
           
           
    End Select

   Exit Function
   
Erro_AFRAC_ProgramarTributacao:

    AFRAC_ProgramarTributacao = 1
    
    Exit Function

End Function

Public Function AFRAC_SairHorarioVerao() As Integer
'Retira o ECF no horário de Verão

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim iFlag As Integer
Dim sRetorno As String
Dim iRZPendente As Integer
Dim lErro As Long
Dim sDataRed As String
Dim sHoraRed As String
Dim sDataImp As String
Dim sHoraImp As String
Dim dtDataRed As Date
Dim dtHoraRed As Date
Dim dtDataImp As Date
Dim dtHoraImp As Date
Dim dtDataHoraRed As Date
Dim dtDataHoraImp As Date

On Error GoTo Erro_AFRAC_SairHorarioVerao

        If InStr(UCase(X.gsNumSerie), "EMULADOR") = 0 Then

            AFRAC_SairHorarioVerao = 1
        
            lErro = AFRAC_DataHoraReducao(sDataRed, sHoraRed)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Data/Hora Ultima Reducao Z")
            If lErro <> SUCESSO Then gError 214500
        
        
            'verifica se ha diferenca de data, hora > 1 hora - PAFECF
            lErro = AFRAC_DataHoraImpressora(sDataImp, sHoraImp)
            If lErro <> SUCESSO Then
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Data/Hora Impressora")
                If lErro <> SUCESSO Then gError 214501
            End If
    
    
            dtDataRed = left(sDataRed, 2) & "/" & Mid(sDataRed, 3, 2) & "/" & right(sDataRed, 2)
    
            dtHoraRed = left(sHoraRed, 2) & ":" & Mid(sHoraRed, 3, 2) & ":" & right(sHoraRed, 2)
            
            dtDataHoraRed = dtDataRed + dtHoraRed
            
            dtDataImp = left(sDataImp, 2) & "/" & Mid(sDataImp, 3, 2) & "/" & right(sDataImp, 2)
    
            dtHoraImp = left(sHoraImp, 2) & ":" & Mid(sHoraImp, 3, 2) & ":" & right(sHoraImp, 2)
            
            dtDataHoraImp = dtDataImp + dtHoraImp
            
    
            'so pode desativar o horario de verao apos uma hora da reducao Z e sem emitir cupons fiscais
            
            If DateDiff("n", dtDataHoraRed, dtDataHoraImp) > 60 Then
        
        
                Select Case X.giCodModeloECF
            
                    Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
                        iStatus = Bematech_FI_FlagsFiscais(iFlag)
                        
                        If iStatus = 1 And (iFlag And 4) Then
            
                            iStatus = Bematech_FI_ProgramaHorarioVerao
                                
                        End If
            
                        If iStatus = 1 Then
                            AFRAC_SairHorarioVerao = SUCESSO
                        ElseIf iStatus = 0 Then
                            AFRAC_SairHorarioVerao = 1
                        Else
                            AFRAC_SairHorarioVerao = iStatus
                        End If
                    
                    Case IMPRESSORA_DARUMA_FS600
            
                         '   iStatus = Daruma_FI_FlagsFiscais(iFlag)
                         ' nao precisa testar o flag pois agora tem funcoes especificas para habilitar e desabilitar antes era uma funcao unica liga-desliga
                         iStatus = 1
                         iFlag = 4
                        
                        If iStatus = 1 And (iFlag And 4) Then
            
            '                    AFRAC_SairHorarioVerao = Daruma_FI_ProgramaHorarioVerao
                            AFRAC_SairHorarioVerao = confDesabilitarHorarioVerao_ECF_Daruma()
                                
                            AFRAC_SairHorarioVerao = SUCESSO
            
                        End If
            
                    Case IMPRESSORA_ELGIN
                        AFRAC_SairHorarioVerao = Elgin_ProgramaHorarioVeraoStr("-")
            
                        sRetorno = Space(50)
                        iStatus = Elgin_EsperaResposta(sRetorno)
                    
                        AFRAC_SairHorarioVerao = SUCESSO
            
                    Case Else
                        AFRAC_SairHorarioVerao = SUCESSO
            
            
                End Select
                
    
            Else
    
                AFRAC_SairHorarioVerao = SUCESSO
    
            End If
            
        Else
            AFRAC_SairHorarioVerao = SUCESSO
        End If

   Exit Function
   
Erro_AFRAC_SairHorarioVerao:

    AFRAC_SairHorarioVerao = 1
    
    Exit Function

End Function

Public Function AFRAC_EntrarHorarioVerao() As Integer
'Coloca o ECF no horário de Verão

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim iFlag As Integer
Dim sRetorno As String
Dim iRZPendente As Integer
Dim lErro As Long
Dim sDataRed As String
Dim sHoraRed As String
Dim sDataImp As String
Dim sHoraImp As String
Dim dtDataRed As Date
Dim dtHoraRed As Date
Dim dtDataImp As Date
Dim dtHoraImp As Date
Dim dtDataHoraRed As Date
Dim dtDataHoraImp As Date


On Error GoTo Erro_AFRAC_EntrarHorarioVerao

            AFRAC_EntrarHorarioVerao = 1


            Select Case X.giCodModeloECF
        
                Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
                    If InStr(UCase(X.gsNumSerie), "EMULADOR") = 0 Then
        
        
                        iStatus = Bematech_FI_FlagsFiscais(iFlag)
                        
                        If iStatus = 1 And Not (iFlag And 4) Then
            
                            iStatus = Bematech_FI_ProgramaHorarioVerao
                            
                        End If
                        
                        If iStatus = 1 Then
                            AFRAC_EntrarHorarioVerao = SUCESSO
                        ElseIf iStatus = 0 Then
                            AFRAC_EntrarHorarioVerao = 1
                        Else
                            AFRAC_EntrarHorarioVerao = iStatus
                        End If
                    Else
                    
                        AFRAC_EntrarHorarioVerao = SUCESSO
                        
                    End If
                
                Case IMPRESSORA_DARUMA_FS600
        
                     '   iStatus = Daruma_FI_FlagsFiscais(iFlag)
                    iStatus = 1
                    iFlag = 0
                    
                    If iStatus = 1 And Not (iFlag And 4) Then
        
        '                    AFRAC_EntrarHorarioVerao = Daruma_FI_ProgramaHorarioVerao
                        AFRAC_EntrarHorarioVerao = confHabilitarHorarioVerao_ECF_Daruma()
        
                        AFRAC_EntrarHorarioVerao = SUCESSO
        
                    End If
                    
                Case IMPRESSORA_ELGIN
                    AFRAC_EntrarHorarioVerao = Elgin_ProgramaHorarioVeraoStr("+")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                
                    If iStatus = SUCESSO Or iStatus = -26 Or iStatus = -27 Then
                        AFRAC_EntrarHorarioVerao = SUCESSO
                    Else
                        AFRAC_EntrarHorarioVerao = 1
                    End If
                    
                Case Else
                    AFRAC_EntrarHorarioVerao = SUCESSO
                    
        
            End Select
    
    
    Exit Function
   
Erro_AFRAC_EntrarHorarioVerao:

    AFRAC_EntrarHorarioVerao = 1
    
    Exit Function

End Function

Public Function AFRAC_ProgramarTotalizadorNaoSujeitoICMS(sIndice As String, sNomeTotalizador As String, sTipoOperacao As String, bVinculado As Boolean, bAceitaFormPag As Boolean, sIndiceFormaPagamento As String) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String

'Programa os totalizadores não sujeitos a ICMS

On Error GoTo Erro_AFRAC_ProgramarTotalizadorNaoSujeitoICMS

    AFRAC_ProgramarTotalizadorNaoSujeitoICMS = 1

    Select Case X.giCodModeloECF

        Case 1 To 2, 5, IMPRESSORA_BEMATECH_MP2100_TH_FI

            AFRAC_ProgramarTotalizadorNaoSujeitoICMS = Bematech_FI_NomeiaTotalizadorNaoSujeitoIcms(CInt(sIndice), sNomeTotalizador)
            
            If AFRAC_ProgramarTotalizadorNaoSujeitoICMS = 1 Then
                AFRAC_ProgramarTotalizadorNaoSujeitoICMS = SUCESSO
            ElseIf AFRAC_ProgramarTotalizadorNaoSujeitoICMS = 0 Then
                AFRAC_ProgramarTotalizadorNaoSujeitoICMS = 1
            End If

        Case IMPRESSORA_BEMATECH_MP4200_TH_FI

            If sIndice <> X.giTotalizador_Sangria Then
            
                AFRAC_ProgramarTotalizadorNaoSujeitoICMS = Bematech_FI_NomeiaTotalizadorNaoSujeitoIcms(CInt(sIndice), sNomeTotalizador)
                
                If AFRAC_ProgramarTotalizadorNaoSujeitoICMS = 1 Then
                    AFRAC_ProgramarTotalizadorNaoSujeitoICMS = SUCESSO
                ElseIf AFRAC_ProgramarTotalizadorNaoSujeitoICMS = 0 Then
                    AFRAC_ProgramarTotalizadorNaoSujeitoICMS = 1
                End If
            Else
                   AFRAC_ProgramarTotalizadorNaoSujeitoICMS = SUCESSO



            End If


        Case IMPRESSORA_DARUMA_FS600

            'o Daruma ja faz configura 1=Sangria 2=Suprimento
            If CInt(sIndice) > 2 Then
                ' AFRAC_ProgramarTotalizadorNaoSujeitoICMS = Daruma_FI_NomeiaTotalizadorNaoSujeitoIcms(CInt(sIndice), sNomeTotalizador)
                AFRAC_ProgramarTotalizadorNaoSujeitoICMS = confCadastrar_ECF_Daruma("TNF", sNomeTotalizador, ";")
            End If
            
            If AFRAC_ProgramarTotalizadorNaoSujeitoICMS = 1 Then
                AFRAC_ProgramarTotalizadorNaoSujeitoICMS = SUCESSO
            ElseIf AFRAC_ProgramarTotalizadorNaoSujeitoICMS = 0 Then
                AFRAC_ProgramarTotalizadorNaoSujeitoICMS = 1
            End If

         Case IMPRESSORA_ELGIN
            AFRAC_ProgramarTotalizadorNaoSujeitoICMS = Elgin_ProgramaLegenda(CInt(sIndice) + 15, sNomeTotalizador)

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)
        
            AFRAC_ProgramarTotalizadorNaoSujeitoICMS = SUCESSO

    
        Case Else
            AFRAC_ProgramarTotalizadorNaoSujeitoICMS = SUCESSO
        
    

    End Select

   Exit Function
   
Erro_AFRAC_ProgramarTotalizadorNaoSujeitoICMS:

    AFRAC_ProgramarTotalizadorNaoSujeitoICMS = 1
    
    Exit Function

End Function

Public Function AFRAC_ArredondarTruncar(iMetodo As Integer) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim iErro As Integer

'Arredonda ou trunca os resultados de cálculo

On Error GoTo Erro_AFRAC_ArredondarTruncar

    AFRAC_ArredondarTruncar = 1

    Select Case X.giCodModeloECF

        Case 1 To 2, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            If iMetodo = 0 Then
                AFRAC_ArredondarTruncar = Bematech_FI_ProgramaArredondamento()
            Else
                AFRAC_ArredondarTruncar = Bematech_FI_ProgramaTruncamento()
            End If

            If AFRAC_ArredondarTruncar = 1 Then
                AFRAC_ArredondarTruncar = SUCESSO
            ElseIf AFRAC_ArredondarTruncar = 0 Then
                AFRAC_ArredondarTruncar = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

'                iErro = Daruma_Registry_AplMensagem1("Corporator - ECF 3.2")
'                iErro = Daruma_Registry_AplMensagem2("")
                iErro = regAlterarValor_Daruma("ECF\MensagemApl1", "Corporator - ECF 3.2")
                iErro = regAlterarValor_Daruma("ECF\MensagemApl2", "")


            If iMetodo = 0 Then
                    'AFRAC_ArredondarTruncar = Daruma_FI_ProgramaArredondamento()
                    AFRAC_ArredondarTruncar = regAlterarValor_Daruma("ECF\ArredondarTruncar", "A")
            Else
                    'AFRAC_ArredondarTruncar = Daruma_FI_ProgramaTruncamento()
                    AFRAC_ArredondarTruncar = regAlterarValor_Daruma("ECF\ArredondarTruncar", "T")
            End If

            If AFRAC_ArredondarTruncar = 1 Then
                AFRAC_ArredondarTruncar = SUCESSO
            ElseIf AFRAC_ArredondarTruncar = 0 Then
                AFRAC_ArredondarTruncar = 1
            End If
            
        Case IMPRESSORA_ELGIN
            If iMetodo = 0 Then
                AFRAC_ArredondarTruncar = Elgin_ProgArredondamento(1)
            Else
                AFRAC_ArredondarTruncar = Elgin_ProgArredondamento(0)
            End If
            '0' (30h) ou 0 (00h) – Truncamento.
            '1' (31h) ou 1 (01h) – Arredondamento.
            '2' (32h) ou 2 (02h) – Arredondamento ABNT.

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)
        
            If iStatus = SUCESSO Or iStatus = -20 Then
                AFRAC_ArredondarTruncar = SUCESSO
            Else
                AFRAC_ArredondarTruncar = 1
            End If


        Case Else
            AFRAC_ArredondarTruncar = SUCESSO
                    
    End Select

   Exit Function
   
Erro_AFRAC_ArredondarTruncar:

    AFRAC_ArredondarTruncar = 1
    
    Exit Function

End Function

Public Function AFRAC_ConfigurarLinhasEntreCupons(iQuantLinhas As Integer) As Integer

Dim X As New ClassECFConfig
Dim sQuantLinhas As String
Dim sBuffer As String
Dim iStatus As Integer

'Configura a quantidade de linhas entre o final de um cupom e o início do clichê de outro.

On Error GoTo Erro_AFRAC_ConfigurarLinhasEntreCupons

    AFRAC_ConfigurarLinhasEntreCupons = 1

    Select Case X.giCodModeloECF

        Case 1 To 2, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            sQuantLinhas = iQuantLinhas
            
            AFRAC_ConfigurarLinhasEntreCupons = Bematech_FI_LinhasEntreCupons(CInt(sQuantLinhas))
            
            If AFRAC_ConfigurarLinhasEntreCupons = 1 Then
                AFRAC_ConfigurarLinhasEntreCupons = SUCESSO
            ElseIf AFRAC_ConfigurarLinhasEntreCupons = 0 Then
                AFRAC_ConfigurarLinhasEntreCupons = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

            sQuantLinhas = iQuantLinhas
            
            '     AFRAC_ConfigurarLinhasEntreCupons = Daruma_FI_LinhasEntreCupons(CInt(sQuantLinhas))
            
            sBuffer = Space(1)
            iStatus = rRetornarInformacao_ECF_Daruma("57", sBuffer)
            
            'se o cupom estiver fechado
            If sBuffer = "0" Then
        
                AFRAC_ConfigurarLinhasEntreCupons = confProgramarAvancoPapel_ECF_Daruma("25", sQuantLinhas, sQuantLinhas, "1", "1")
                
            Else
                AFRAC_ConfigurarLinhasEntreCupons = 1
                
            End If

            If AFRAC_ConfigurarLinhasEntreCupons = 1 Then
                AFRAC_ConfigurarLinhasEntreCupons = SUCESSO
            ElseIf AFRAC_ConfigurarLinhasEntreCupons = 0 Then
                AFRAC_ConfigurarLinhasEntreCupons = 1
            End If
            
        Case Else
            AFRAC_ConfigurarLinhasEntreCupons = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_ConfigurarLinhasEntreCupons:

    AFRAC_ConfigurarLinhasEntreCupons = 1
    
    Exit Function

End Function

Public Function AFRAC_ConfigurarEspacoEntreLinhas(lEspaco As Long) As Integer

Dim X As New ClassECFConfig
Dim sEspaco As String

'Configura o espaço entre as linhas

On Error GoTo Erro_AFRAC_ConfigurarEspacoEntreLinhas

    AFRAC_ConfigurarEspacoEntreLinhas = 1

    Select Case X.giCodModeloECF

        Case 1 To 2, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            sEspaco = lEspaco
            
            AFRAC_ConfigurarEspacoEntreLinhas = Bematech_FI_EspacoEntreLinhas(CInt(FormataCpoNum(sEspaco, 3)))
            
            If AFRAC_ConfigurarEspacoEntreLinhas = 1 Then
                AFRAC_ConfigurarEspacoEntreLinhas = SUCESSO
            ElseIf AFRAC_ConfigurarEspacoEntreLinhas = 0 Then
                AFRAC_ConfigurarEspacoEntreLinhas = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

            sEspaco = lEspaco
            
'                 AFRAC_ConfigurarEspacoEntreLinhas = Daruma_FI_EspacoEntreLinhas(CInt(FormataCpoNum(sEspaco, 3)))
            AFRAC_ConfigurarEspacoEntreLinhas = 1
            
            If AFRAC_ConfigurarEspacoEntreLinhas = 1 Then
                AFRAC_ConfigurarEspacoEntreLinhas = SUCESSO
            ElseIf AFRAC_ConfigurarEspacoEntreLinhas = 0 Then
                AFRAC_ConfigurarEspacoEntreLinhas = 1
            End If
            
        Case Else
            AFRAC_ConfigurarEspacoEntreLinhas = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_ConfigurarEspacoEntreLinhas:

    AFRAC_ConfigurarEspacoEntreLinhas = 1
    
    Exit Function


End Function

Public Function AFRAC_ConfigurarMoeda(sMoeda As String) As Integer

Dim X As New ClassECFConfig
Dim iPos As Integer

'Configura a moeda

On Error GoTo Erro_AFRAC_ConfigurarMoeda

    AFRAC_ConfigurarMoeda = 1

    Select Case X.giCodModeloECF

        Case 1 To 2, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            iPos = InStr(1, sMoeda, "$")
            If iPos > 0 Then
                sMoeda = left(sMoeda, Len(sMoeda) - 1)
            End If

            AFRAC_ConfigurarMoeda = Bematech_FI_AlteraSimboloMoeda(sMoeda)
            If AFRAC_ConfigurarMoeda = 1 Then
                AFRAC_ConfigurarMoeda = SUCESSO
            ElseIf AFRAC_ConfigurarMoeda = 0 Then
                AFRAC_ConfigurarMoeda = 1
            End If
            
        Case Else
                AFRAC_ConfigurarMoeda = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_ConfigurarMoeda:

    AFRAC_ConfigurarMoeda = 1
    
    Exit Function


End Function

Public Function AFRAC_ForcarImpactoAgulhas(bForcar As Boolean) As Integer

Dim X As New ClassECFConfig

'Indica e deve forcar as agulhas
Dim sForcar As String

On Error GoTo Erro_AFRAC_ForcarImpactoAgulhas

    AFRAC_ForcarImpactoAgulhas = 1

    Select Case X.giCodModeloECF

        Case 1 To 2, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            If bForcar = True Then
                sForcar = 3
            Else
                sForcar = 1
            End If

            AFRAC_ForcarImpactoAgulhas = Bematech_FI_ForcaImpactoAgulhas(CInt(sForcar))
            If AFRAC_ForcarImpactoAgulhas = 1 Then
                AFRAC_ForcarImpactoAgulhas = SUCESSO
            ElseIf AFRAC_ForcarImpactoAgulhas = 0 Then
                AFRAC_ForcarImpactoAgulhas = 1
            End If
            
        Case Else
            AFRAC_ForcarImpactoAgulhas = SUCESSO

    End Select

   Exit Function
   
Erro_AFRAC_ForcarImpactoAgulhas:

    AFRAC_ForcarImpactoAgulhas = 1
    
    Exit Function

End Function

Public Function AFRAC_AjustarRelogio(dHora As Double) As Integer

'Ajusta o horário da impressora
    
    AFRAC_AjustarRelogio = SUCESSO


End Function

Public Function AFRAC_ProgramarRelatorioGerencial(sIndice As String, sDescricao As String) As Integer

'Programa o índice do tipo de relatorio gerencial

'Programar as formas de pagamento

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim objTiposMeiosPagtos As ClassTMPLoja
Dim sRelatorios As String
Dim iIndice As Integer
Dim iPos As Integer

On Error GoTo Erro_AFRAC_ProgramarRelatorioGerencial

    AFRAC_ProgramarRelatorioGerencial = 1

    Select Case X.giCodModeloECF

        Case 1, 2, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
'            iStatus = Bematech_FI_ProgramaFormaPagamentoMFD(sDescricao, "1")
            
            iStatus = AFRAC_LerRelatorioGerencial(sRelatorios)
            If iStatus = SUCESSO Then
            
                iPos = InStr(sRelatorios, sDescricao)
            
                If iPos = 0 Then
                    iStatus = Bematech_FI_NomeiaRelatorioGerencialMFD(sIndice, sDescricao)
                
                Else
                    'retorna o indice que atualmente esta programado o relatorio gerencial
                    iIndice = (iPos \ 22) + 1
                    sIndice = iIndice
                    iStatus = 1
                End If
            
            
            Else
                iStatus = -1
            End If
            
            If iStatus = 1 Then
                AFRAC_ProgramarRelatorioGerencial = SUCESSO
            ElseIf iStatus = 0 Then
                AFRAC_ProgramarRelatorioGerencial = 1
            Else
                AFRAC_ProgramarRelatorioGerencial = iStatus
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            iStatus = AFRAC_LerRelatorioGerencial(sRelatorios)
            If iStatus = SUCESSO Then
            
                iPos = InStr(sRelatorios, sDescricao)
            
                If iPos = 0 Then
                    iStatus = confCadastrarPadrao_ECF_Daruma("RG", left(sDescricao, 15))
                
                Else
                    'retorna o indice que atualmente esta programado o relatorio gerencial
                    iIndice = (iPos \ 16) + 1
                    sIndice = iIndice
                    iStatus = 1
                End If
            
            
            Else
                iStatus = -1
            End If
        
        
            If iStatus = 1 Then
                AFRAC_ProgramarRelatorioGerencial = SUCESSO
            ElseIf iStatus = 0 Then
                AFRAC_ProgramarRelatorioGerencial = 1
            Else
                AFRAC_ProgramarRelatorioGerencial = iStatus
            End If
        
        Case IMPRESSORA_ELGIN
            For Each objTiposMeiosPagtos In X.gcolTiposMeiosPagtos
                If CInt(sIndice) = objTiposMeiosPagtos.iTipo Then
                    iStatus = Elgin_ProgramaLegenda(Format(objTiposMeiosPagtos.iIndice, "00"), DesacentuaTexto(Format(left(sDescricao, 16), "!@@@@@@@@@@@@@@@@")))
                    
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                
                    Exit For
                End If
            Next
        
            
            If iStatus = SUCESSO Or iStatus = -20 Then
                AFRAC_ProgramarRelatorioGerencial = SUCESSO
            Else
                AFRAC_ProgramarRelatorioGerencial = 1
            End If
            
        Case Else
            AFRAC_ProgramarRelatorioGerencial = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_ProgramarRelatorioGerencial:

    AFRAC_ProgramarRelatorioGerencial = 1
    
    Exit Function


End Function

Public Function AFRAC_ProgramarFormasDePagamento(sIndice As String, sDescricao As String, bVinculado As Boolean) As Integer

'Programar as formas de pagamento

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim objTiposMeiosPagtos As ClassTMPLoja
Dim sFgto As String
Dim iPos As Integer
Dim iIndice As Integer
Dim iPos1 As Integer

On Error GoTo Erro_AFRAC_ProgramarFormasDePagamento

    AFRAC_ProgramarFormasDePagamento = 1

    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            iStatus = AFRAC_LerFormasPagamento(sFgto)
            If iStatus = SUCESSO Then
            
                iPos = InStr(sFgto, left(sDescricao, 15))
            
                If iPos = 0 Then
                    If UCase(DesacentuaTexto(sDescricao)) = "CARTAO CREDITO" Or UCase(DesacentuaTexto(sDescricao)) = "CARTAO DEBITO" Or UCase(DesacentuaTexto(sDescricao)) = "TEF" Then
                        iStatus = Bematech_FI_ProgramaFormaPagamentoMFD(left(sDescricao, 15), "1")
                    Else
                        iStatus = Bematech_FI_ProgramaFormaPagamentoMFD(left(sDescricao, 15), "0")
                    End If

                Else
                    'retorna o indice que atualmente esta programada a forma de pagamento
                    iIndice = (iPos \ 46) + 1
                    sIndice = iIndice
                    iStatus = 1
                End If
            
            
            Else
                iStatus = -1
            End If
            
            If iStatus = 1 Then
                AFRAC_ProgramarFormasDePagamento = SUCESSO
            ElseIf iStatus = 0 Then
                AFRAC_ProgramarFormasDePagamento = 1
            Else
                AFRAC_ProgramarFormasDePagamento = iStatus
            End If
        
            
        Case IMPRESSORA_DARUMA_FS600
        
            iStatus = AFRAC_LerFormasPagamento(sFgto)
            If iStatus = SUCESSO Then
            
                iPos = InStr(sFgto, left(sDescricao, 15))
            
                If iPos = 0 Then
                     iStatus = confCadastrarPadrao_ECF_Daruma("FPGTO", left(sDescricao, 15))

                Else
                    
                    'retorna o indice que atualmente esta programada a forma de pagamento
                    iPos1 = 1
                    iIndice = 0
                    
                    Do While iPos1 < iPos And iPos1 <> 0
                    
                        iIndice = iIndice + 1
                        iPos1 = InStr(iPos1, sFgto, ";")
                        
                        If iPos1 = 0 Then
                            iIndice = 0
                            iStatus = -1
                            Exit Do
                        End If
                            
                        iPos1 = iPos1 + 1
                        
                    Loop
                                        
                    If iPos1 <> 0 Then
                        sIndice = iIndice
                        iStatus = 1
                    End If
                End If
            
            
            Else
                iStatus = -1
            End If
            
            If iStatus = 1 Then
                AFRAC_ProgramarFormasDePagamento = SUCESSO
            ElseIf iStatus = 0 Then
                AFRAC_ProgramarFormasDePagamento = 1
            Else
                AFRAC_ProgramarFormasDePagamento = iStatus
            End If
        
        
        Case IMPRESSORA_ELGIN
            For Each objTiposMeiosPagtos In X.gcolTiposMeiosPagtos
                If CInt(sIndice) = objTiposMeiosPagtos.iTipo Then
                    iStatus = Elgin_ProgramaLegenda(Format(objTiposMeiosPagtos.iIndice, "00"), DesacentuaTexto(Format(left(sDescricao, 16), "!@@@@@@@@@@@@@@@@")))
                    
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                
                    Exit For
                End If
            Next
        
            
            If iStatus = SUCESSO Or iStatus = -20 Then
                AFRAC_ProgramarFormasDePagamento = SUCESSO
            Else
                AFRAC_ProgramarFormasDePagamento = 1
            End If
            
        Case Else
            AFRAC_ProgramarFormasDePagamento = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_ProgramarFormasDePagamento:

    AFRAC_ProgramarFormasDePagamento = 1
    
    Exit Function

End Function

Public Function AFRAC_LerFormaDePagamento(sIndice As String, sDescricao As String) As Integer

'Leitura ds formas de pagamento encontradas na impressora

    AFRAC_LerFormaDePagamento = SUCESSO

End Function

Public Function AFRAC_AbrirDia(dtData As Date) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String

'Função que Faz a AberTura do dia

On Error GoTo Erro_AFRAC_AbrirDia

    AFRAC_AbrirDia = 1

    Select Case X.giCodModeloECF

        Case 1 To 2
        
            AFRAC_AbrirDia = Bematech_FI_AberturaDoDia("0", "")
            
            If AFRAC_AbrirDia = 1 Then
                AFRAC_AbrirDia = SUCESSO
            ElseIf AFRAC_AbrirDia = 0 Then
                AFRAC_AbrirDia = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
'                AFRAC_AbrirDia = Daruma_FI_AberturaDoDia("0", "")
            AFRAC_AbrirDia = 1
            
            If AFRAC_AbrirDia = 1 Then
                AFRAC_AbrirDia = SUCESSO
            ElseIf AFRAC_AbrirDia = 0 Then
                AFRAC_AbrirDia = 1
            End If
            
        Case IMPRESSORA_ELGIN
            'iStatus = Elgin_LeituraX(0)
            
            'sRetorno = Space(50)
            'iStatus = Elgin_EsperaResposta(sRetorno)
        
            'If iStatus = SUCESSO Then
            '    AFRAC_AbrirDia = SUCESSO
            'Else
            '    AFRAC_AbrirDia = 1
            'End If
            
            AFRAC_AbrirDia = SUCESSO
            
        Case Else
            AFRAC_AbrirDia = SUCESSO
        
        
    End Select

   Exit Function
   
Erro_AFRAC_AbrirDia:

    AFRAC_AbrirDia = 1
    
    Exit Function

End Function

Public Function AFRAC_FechaDia() As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim iRetorno1 As Integer
Dim iRZPendenteAnt As Long
Dim iRZPendenteAtual As Long
Dim sDataProxRZ As String
Dim dtDataProxRZ As Date

'Função que Faz a AberTura do dia

On Error GoTo Erro_AFRAC_FechaDia

    AFRAC_FechaDia = 1

    Select Case X.giCodModeloECF

        Case 1, 2, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            iStatus = AFRAC_DataMovimentoProximaReducao(sDataProxRZ)
            
            If iStatus = SUCESSO Then
            
            
                dtDataProxRZ = CDate(Mid(sDataProxRZ, 1, 2) & "/" & Mid(sDataProxRZ, 3, 2) & "/" & Mid(sDataProxRZ, 5, 4))
            
                'verifica se da data da proxima reducao Z está setada, ou seja, se tem reducao Z pendente
                If dtDataProxRZ <> DATA_NULA Then
            
                    AFRAC_FechaDia = Bematech_FI_FechamentoDoDia()
                    
                    If AFRAC_FechaDia = 1 Then
                        AFRAC_FechaDia = SUCESSO
                    ElseIf AFRAC_FechaDia = 0 Then
                        AFRAC_FechaDia = 1
                    End If
                    
                Else
                    AFRAC_FechaDia = SUCESSO
                    
                End If
            
            Else
            
                AFRAC_FechaDia = 1
                
            End If
            
            
        Case IMPRESSORA_DARUMA_FS600
            
'                AFRAC_FechaDia = Daruma_FI_FechamentoDoDia()
'                AFRAC_FechaDia = Daruma_FI_ReducaoZ("", "")

            'se ha rz pendente do dia anterior = 1
            iRetorno1 = rConsultaStatusImpressoraInt_ECF_Daruma(6, iRZPendenteAnt)
            If iRetorno1 <> 1 Then gError 210891
                
            'se ainda nao emitiu reducao Z hoje  = 0
            iRetorno1 = rConsultaStatusImpressoraInt_ECF_Daruma(4, iRZPendenteAtual)
            If iRetorno1 <> 1 Then gError 210892
                
            'se tem uma reducao Z pendente anterior e ainda nao emitiu reducao Z hoje
            If iRZPendenteAnt = 1 Or iRZPendenteAtual = 1 Then
            
                AFRAC_FechaDia = iReducaoZ_ECF_Daruma("", "")
                
            Else
                AFRAC_FechaDia = 1
                
            End If

            If AFRAC_FechaDia = 1 Then
                AFRAC_FechaDia = SUCESSO
            ElseIf AFRAC_FechaDia = 0 Then
                AFRAC_FechaDia = 1
            End If
            
        Case IMPRESSORA_ELGIN
            iStatus = Elgin_ReducaoZ("0")
            
            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)
            
            If iStatus = SUCESSO Or iStatus = -29 Then
                AFRAC_FechaDia = 1
            Else
                AFRAC_FechaDia = SUCESSO
            End If
            
        Case Else
            AFRAC_FechaDia = SUCESSO
            
    End Select


   Exit Function
   
Erro_AFRAC_FechaDia:

    AFRAC_FechaDia = 1
    
    Exit Function

End Function

'Funções de Cumpom Fiscal

Public Function AFRAC_AbrirCupom(sCPF As String, sNomeCliente As String, sEndereco As String) As Integer

Dim X As New ClassECFConfig
Dim sRetorno As String
Dim lErro As Long
Dim lCOO As Long
Dim lTamanho As Long
Dim iStatus As Integer
Dim iFlag As Integer
Dim sBuffer As String
Dim iStatus1 As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_AFRAC_AbrirCupom

'Função que Faz a AberTura de Cupom da Afrac

    AFRAC_AbrirCupom = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            iStatus = Bematech_FI_FlagsFiscais(iFlag)
        
            'se o cupom anterior nao estiver fechado...cancela
            If iStatus = 1 And (iFlag And 1) Then
                lErro = AFRAC_CancelarCupom(Nothing, Nothing)

            End If
        
            iStatus = Bematech_FI_FlagsFiscais(iFlag)
        
            'se retorno ok e o cupom nao estiver aberto
            If iStatus = 1 And Not (iFlag And 1) Then
            
                If X.giCodModeloECF = 5 Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP4200_TH_FI Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP2100_TH_FI Then
                    AFRAC_AbrirCupom = Bematech_FI_AbreCupomMFD(sCPF, sNomeCliente, sEndereco)
                Else
                    AFRAC_AbrirCupom = Bematech_FI_AbreCupom(sCPF)
                End If
                
                If AFRAC_AbrirCupom = 1 Then
                    AFRAC_AbrirCupom = SUCESSO
                ElseIf AFRAC_AbrirCupom = 0 Then
                    AFRAC_AbrirCupom = 1
                End If
            
            End If
            
        Case IMPRESSORA_DARUMA_FS600
            
'                iStatus = Daruma_FI_FlagsFiscais(iFlag)
                
            sBuffer = Space(1)
            iStatus = rRetornarInformacao_ECF_Daruma("57", sBuffer)
            
            iFlag = 0
            
            If sBuffer <> "0" Then
                iFlag = 1
            End If
                
            'se o cupom anterior nao estiver fechado...cancela
            If iStatus = 1 And (iFlag And 1) Then
                lErro = AFRAC_CancelarCupom(Nothing, Nothing)

            End If
        
'                iStatus = Daruma_FI_FlagsFiscais(iFlag)
            
            sBuffer = Space(1)
            iStatus = rRetornarInformacao_ECF_Daruma("57", sBuffer)
            
            If sBuffer = "0" Then
                iFlag = 0
            Else
                iFlag = 1
            End If
            
            'se retorno ok e o cupom nao estiver aberto
            If iStatus = 1 And Not (iFlag And 1) Then
            
'                    AFRAC_AbrirCupom = Daruma_FI_AbreCupom(sCPF)
                If Len(sCPF) = 0 Then sCPF = " "
                If Len(sNomeCliente) = 0 Then sNomeCliente = " "
                If Len(sEndereco) = 0 Then sEndereco = " "
                    
                AFRAC_AbrirCupom = iCFAbrir_ECF_Daruma(sCPF, sNomeCliente, sEndereco)

                If AFRAC_AbrirCupom = 1 Then
                    AFRAC_AbrirCupom = SUCESSO
                ElseIf AFRAC_AbrirCupom = 0 Then
                    AFRAC_AbrirCupom = 1
                End If
            
            End If
            
        Case IMPRESSORA_NAO_FISCAL
        
            lTamanho = 6
            sRetorno = String(lTamanho, 0)
        
            Call GetPrivateProfileString(APLICACAO_CAIXA, "COO", CONSTANTE_ERRO, sRetorno, lTamanho + 1, NOME_ARQUIVO_CAIXA)
            If sRetorno <> String(lTamanho, 0) Then
            
                lCOO = StrParaLong(sRetorno)
        
                lCOO = lCOO + 1
        
                'Atualiza o arquivo
                lErro = WritePrivateProfileString(APLICACAO_CAIXA, "COO", CStr(lCOO), NOME_ARQUIVO_CAIXA)
                If lErro <> 0 Then
                    AFRAC_AbrirCupom = SUCESSO
                Else
                    AFRAC_AbrirCupom = 1
                End If
        
            Else
                AFRAC_AbrirCupom = 1
            End If
            
        Case IMPRESSORA_ELGIN
        
            AFRAC_AbrirCupom = 1
        
            sBuffer = Space(41)
            
            iStatus = Elgin_TransStatus(0, sBuffer)
            
            sRetorno = Space(50)
            iStatus1 = Elgin_EsperaResposta(sRetorno)
            
            'se o cupom anterior nao estiver fechado...cancela
            If iStatus = SUCESSO And Mid(sBuffer, 8, 1) = "1" Then
        
                lErro = AFRAC_CancelarCupom(Nothing, Nothing)
        
                sRetorno = Space(50)
                iStatus = Elgin_EsperaResposta(sRetorno)
            
                If iStatus = SUCESSO Then
                    AFRAC_AbrirCupom = SUCESSO
                Else
                    AFRAC_AbrirCupom = 1
                End If
        
            End If
            
            sBuffer = Space(41)
            
            iStatus = Elgin_TransStatus(0, sBuffer)
            
            sRetorno = Space(50)
            iStatus1 = Elgin_EsperaResposta(sRetorno)
        
            If iStatus = SUCESSO Then
                AFRAC_AbrirCupom = SUCESSO
            Else
                AFRAC_AbrirCupom = 1
            End If
            
            'se retorno ok e o cupom nao estiver aberto
            If iStatus = SUCESSO And Mid(sBuffer, 8, 1) <> "1" Then
        
                iStatus = Elgin_AbreCupomFiscalCPF_CNPJ(sCPF)
                
                sRetorno = Space(50)
                iStatus = Elgin_EsperaResposta(sRetorno)
            
                If iStatus = SUCESSO Then
                    AFRAC_AbrirCupom = SUCESSO
                Else
                    AFRAC_AbrirCupom = 1
                End If
            
            End If
        
        Case Else
            AFRAC_AbrirCupom = SUCESSO
        
        
    End Select

   Exit Function
   
Erro_AFRAC_AbrirCupom:

    AFRAC_AbrirCupom = 1
    
    Exit Function

End Function

Public Function AFRAC_VenderItem(sCodigo As String, sDescricao As String, dQuantidade As Double, sValor_Unitario As String, iAcres_Desc As Integer, iPerc_valor As Integer, dValor_AcresDesc As Double, dValor_Total As Double, sAliquota As String, sUnidade As String, bForcarImpreUmaLinha As Boolean) As Integer
'Imprime o Item de Venda

Dim X As New ClassECFConfig
Dim sDescricao1 As String
Dim sAliquotasISS As String
Dim sIndiceAliquotasISS As String
Dim iAchou As Integer
Dim iInicio As Integer
Dim iFim As Integer
Dim sTrib As String
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim lErro As Long
Dim sQuantidade As String
Dim sPreco As String
Dim iStatus As Integer
Dim sBuffer1 As String
Dim sRetorno As String
Dim sTipoDescAcresc As String


On Error GoTo Erro_AFRAC_VenderItem

    AFRAC_VenderItem = 1

   Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            If sAliquota = TIPOTRIBICMS_SITUACAOTRIBECF_NAO_TRIBUTADO Then
                sAliquota = "NN"
            ElseIf sAliquota = TIPOTRIBICMS_SITUACAOTRIBECF_ISENTA Then
                sAliquota = "II"
            ElseIf sAliquota = TIPOTRIBICMS_SITUACAOTRIBECF_TRIB_SUBST Then
                sAliquota = "FF"
            ElseIf sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_NAO_TRIBUTADO Then
                sAliquota = "SN"
            ElseIf sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_ISENTA Then
                sAliquota = "SI"
            ElseIf sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_TRIB_SUBST Then
                sAliquota = "SF"
            Else
                If left(sAliquota, 1) = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL Then
                    
                    sAliquota = Mid(sAliquota, 2)
                                    
                    sAliquotasISS = String(79, 0)
                                    
                    lErro = AFRAC_VerificaAliquotasIss(sAliquotasISS)
                    If lErro <> SUCESSO Then
                        gError 133673
                    End If
                                    
                    sIndiceAliquotasISS = String(48, 0)
                                    
                    lErro = AFRAC_VerificaIndiceAliquotasIss(sIndiceAliquotasISS)
                    If lErro <> SUCESSO Then
                        gError 133673
                    End If
                    
                    iIndice = 0
                    iAchou = 0
                    iInicio = 1
                    iFim = InStr(iInicio, sAliquotasISS, ",")
                    
                    If iFim = 0 And Mid(sAliquotasISS, iInicio, 1) <> Chr(0) Then
                        iFim = InStr(iInicio, sAliquotasISS, Chr(0))
                    End If
                    
                    Do While iFim <> 0
                        iIndice = iIndice + 1
                        sTrib = Mid(sAliquotasISS, iInicio, iFim - iInicio)
                        If sTrib = sAliquota Then
                            iAchou = 1
                            Exit Do
                        End If
                        iInicio = iFim + 1
                        iFim = InStr(iInicio, sAliquotasISS, ",")
                        If iFim = 0 And Mid(sAliquotasISS, iInicio, 1) <> Chr(0) Then iFim = InStr(iInicio, sAliquotasISS, Chr(0))
                    Loop
                
                    iInicio = 1
                
                    If iAchou = 1 Then
                
                        If iIndice > 0 Then
                            
                            For iIndice1 = 1 To iIndice - 1
                                iInicio = InStr(iInicio, sIndiceAliquotasISS, ",")
                                iInicio = iInicio + 1
                            Next
                
                            sAliquota = Mid(sIndiceAliquotasISS, iInicio, 2)
                    
                        End If
                        
                    End If
                    
                End If
            
            End If
                
            sDescricao1 = Mid(sDescricao, 1, 29)
            
            
            'se quiser que apareça a unidade de medida então use Bematech_FI_VendeItemDepartamento
            If iPerc_valor = 0 Then
                AFRAC_VenderItem = Bematech_FI_VendeItem(sCodigo, sDescricao1, sAliquota, IIf(dQuantidade - Int(dQuantidade) = 0, "I", "F"), CStr(dQuantidade), 2, sValor_Unitario, "%", Format(dValor_AcresDesc * 100, "0000"))
            Else
                AFRAC_VenderItem = Bematech_FI_VendeItem(sCodigo, sDescricao1, sAliquota, IIf(dQuantidade - Int(dQuantidade) = 0, "I", "F"), CStr(dQuantidade), 2, sValor_Unitario, "$", Format(dValor_AcresDesc, "standard"))
            End If
            
            If AFRAC_VenderItem = 1 Then
                AFRAC_VenderItem = SUCESSO
            ElseIf AFRAC_VenderItem = 0 Then
                AFRAC_VenderItem = 1
            End If

        Case IMPRESSORA_DARUMA_FS600

            If sAliquota = TIPOTRIBICMS_SITUACAOTRIBECF_NAO_TRIBUTADO Then
                sAliquota = "NN"
            ElseIf sAliquota = TIPOTRIBICMS_SITUACAOTRIBECF_ISENTA Then
                sAliquota = "II"
            ElseIf sAliquota = TIPOTRIBICMS_SITUACAOTRIBECF_TRIB_SUBST Then
                sAliquota = "FF"
            ElseIf sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_NAO_TRIBUTADO Then
                sAliquota = "NS"
            ElseIf sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_ISENTA Then
                sAliquota = "ISS"
            ElseIf sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_TRIB_SUBST Then
                sAliquota = "FS"
            End If
                
            sDescricao1 = Mid(sDescricao, 1, 29)
            
            If iPerc_valor = 0 And dValor_AcresDesc >= 0 Then
                sTipoDescAcresc = "D%"
            ElseIf iPerc_valor = 0 And dValor_AcresDesc < 0 Then
                sTipoDescAcresc = "A%"
            ElseIf iPerc_valor = 1 And dValor_AcresDesc >= 0 Then
                sTipoDescAcresc = "D$"
            ElseIf iPerc_valor = 1 And dValor_AcresDesc < 0 Then
                sTipoDescAcresc = "A$"
            End If
'                    AFRAC_VenderItem = Daruma_FI_VendeItem(sCodigo, sDescricao1, sAliquota, "I", CStr(dQuantidade), 2, sValor_Unitario, "%", Format(dValor_AcresDesc * 100, "0000"))
                    AFRAC_VenderItem = iCFVender_ECF_Daruma(sAliquota, CStr(dQuantidade), sValor_Unitario, sTipoDescAcresc, Format(Abs(dValor_AcresDesc), "standard"), sCodigo, left(sUnidade, 3), sDescricao1)
'                Else
'                    AFRAC_VenderItem = Daruma_FI_VendeItem(sCodigo, sDescricao1, sAliquota, "I", CStr(dQuantidade), 2, sValor_Unitario, "$", Format(dValor_AcresDesc, "standard"))
'                End If
            
            If AFRAC_VenderItem = 1 Then
                AFRAC_VenderItem = SUCESSO
            ElseIf AFRAC_VenderItem = 0 Then
                AFRAC_VenderItem = 1
            End If

        Case IMPRESSORA_ELGIN
        'CONFIRMAR
'                sAliquota = "" '?????
                '· "Tnn": Tributado como ICMS ou ISS, com "nn" variando entre "00" e "15" conforme
                'o índice da alíquota. O totalizador de situação tributária "T00" é específico para
                'realizar operações sujeitas ao ISS (Imposto Sobre Serviços).
                '· "F00": Substituição Tributária
                '· "I00": Isenção Tributária
                '· "N00": Não Tributado (não incidência)
        
        
            If sAliquota = TIPOTRIBICMS_SITUACAOTRIBECF_NAO_TRIBUTADO Then
                sAliquota = "N00"
            ElseIf sAliquota = TIPOTRIBICMS_SITUACAOTRIBECF_ISENTA Then
                sAliquota = "I00"
            ElseIf sAliquota = TIPOTRIBICMS_SITUACAOTRIBECF_TRIB_SUBST Then
                sAliquota = "F00))"
            ElseIf sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_NAO_TRIBUTADO Then
                sAliquota = "N00"
            ElseIf sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_ISENTA Then
                sAliquota = "I00"
            ElseIf sAliquota = TIPOTRIBISS_SITUACAOTRIBECF_TRIB_SUBST Then
                sAliquota = "F00"
            Else
                
                If left(sAliquota, 1) = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL Then
                    
                    sAliquota = Mid(sAliquota, 2)
                End If
                
                iStatus = Elgin_TransTabAliquotas()
                
                
                sBuffer1 = Space(100)
                iStatus = Elgin_EsperaResposta(sBuffer1)
                                        
                If iStatus = SUCESSO Then
                    sBuffer1 = Mid(sBuffer1, 6)
                    For iIndice = 0 To 15
                        If Mid(sBuffer1, (iIndice * 4) + 1, 4) = sAliquota Then
                            sAliquota = "T" & Format(iIndice, "00")
                            Exit For
                        End If
                    Next
                End If
                
            End If
        
        
            sQuantidade = Format(Round(dQuantidade, 3) * 1000, "000000")
            
            sPreco = Format(CDbl(sValor_Unitario) * 100, "00000000000")
            
            If iPerc_valor = 0 Then
                 iStatus = Elgin_VendaItemStr("", sQuantidade, sPreco, sAliquota, "%", Format(Round(dValor_AcresDesc, 2) * 100, "0000"), "un", Format(left(sCodigo, 13), "@@@@@@@@@@@@@"), "", DesacentuaTexto(left(sDescricao, 20)), "")
             Else
                 iStatus = Elgin_VendaItemStr("", sQuantidade, sPreco, sAliquota, "1", Format(Round(dValor_AcresDesc, 2) * 100, "000000000000000"), "un", Format(left(sCodigo, 13), "@@@@@@@@@@@@@"), "", DesacentuaTexto(left(sDescricao, 20)), "")
             End If

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_VenderItem = SUCESSO
            Else
                AFRAC_VenderItem = 1
            End If
            
        Case Else
            AFRAC_VenderItem = SUCESSO
        


   End Select
   
   Exit Function
   
Erro_AFRAC_VenderItem:

    AFRAC_VenderItem = 1
    
    Exit Function

End Function

Public Function AFRAC_AcrescimoDescontoCupom(sValorAcrescimo As String, sValorDesconto As String, sMsgDesconto As String, sMsgAcrescimo As String) As Integer

Dim X As New ClassECFConfig

'Imprime o Desconto ou acréscimo

Dim sValor As String
Dim sAcres_Desc As String
Dim iValor As Integer
Dim dValorAcrescimo As Double
Dim dValorDesconto As Double
Dim dValor As Double
Dim Str_Informacao As String
Dim Int_Retorno As Integer
Dim iStatus As Integer
Dim sRetorno As String

On Error GoTo Erro_AFRAC_AcrescimoDescontoCupom

'    Str_Informacao = Space(275)
'
'     Int_Retorno = Daruma_FI_VerificaDescricaoFormasPagamento(Str_Informacao)
 

    AFRAC_AcrescimoDescontoCupom = 1

    dValorAcrescimo = StrParaDbl(sValorAcrescimo)
    
    dValorDesconto = StrParaDbl(sValorDesconto)


    If dValorAcrescimo > dValorDesconto Then
        dValor = dValorAcrescimo - dValorDesconto
        sAcres_Desc = "A"
    Else
        dValor = dValorDesconto - dValorAcrescimo
        sAcres_Desc = "D"
    End If

   dValor = Round(dValor, 2) * 100
   
   Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            AFRAC_AcrescimoDescontoCupom = Bematech_FI_IniciaFechamentoCupom(sAcres_Desc, "$", CStr(dValor))
            
            If AFRAC_AcrescimoDescontoCupom = 1 Then
                AFRAC_AcrescimoDescontoCupom = SUCESSO
            ElseIf AFRAC_AcrescimoDescontoCupom = 0 Then
                AFRAC_AcrescimoDescontoCupom = 1
            End If

        Case IMPRESSORA_DARUMA_FS600

            ' AFRAC_AcrescimoDescontoCupom = Daruma_FI_IniciaFechamentoCupom(sAcres_Desc, "$", CStr(dValor))
            AFRAC_AcrescimoDescontoCupom = iCFTotalizarCupom_ECF_Daruma(sAcres_Desc & "$", CStr(dValor))
            
            If AFRAC_AcrescimoDescontoCupom = 1 Then
                AFRAC_AcrescimoDescontoCupom = SUCESSO
            ElseIf AFRAC_AcrescimoDescontoCupom = 0 Then
                AFRAC_AcrescimoDescontoCupom = 1
            End If

        Case IMPRESSORA_ELGIN
        
            If dValor = 0 Then
                iStatus = Elgin_TotCupomSemDescAcres
            ElseIf sAcres_Desc = "A" Then
                iStatus = Elgin_TotCupomAcresValor(Format(dValor, "000000000000000"), "")
            ElseIf sAcres_Desc = "D" Then
                iStatus = Elgin_TotCupomDescValor(Format(dValor, "000000000000000"), "")
            End If
                

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)


            If iStatus = SUCESSO Then
                AFRAC_AcrescimoDescontoCupom = SUCESSO
            Else
                AFRAC_AcrescimoDescontoCupom = 1
            End If
            
        Case Else
            AFRAC_AcrescimoDescontoCupom = SUCESSO
        

   End Select

   Exit Function
   
Erro_AFRAC_AcrescimoDescontoCupom:

    AFRAC_AcrescimoDescontoCupom = 1
    
    Exit Function

End Function

Public Function AFRAC_FecharAcrescimoDesconto(sMsgDesconto As String, sMsgAcrescimo As String, iValor As Integer) As Integer


'Fechar o Desconto ou acréscimo

    AFRAC_FecharAcrescimoDesconto = SUCESSO

End Function

Public Function AFRAC_FormaPagamento(sDescForma As String, sIndice As String, sValor As String, Optional sMsg As String) As Integer

Dim X As New ClassECFConfig
Dim sFormaPagto As String
Dim iPos As Integer
Dim objTiposMeiosPagtos As ClassTMPLoja
Dim sRetorno As String
Dim iStatus As Integer


'Registra a Forma da Pagamento, Este comando pode ser Executado n vezes para registrar n formas de Pagamento.

On Error GoTo Erro_AFRAC_FormaPagamento

    AFRAC_FormaPagamento = 1

    Select Case X.giCodModeloECF

        Case 1 To 2, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sFormaPagto = sDescForma
        
            If Len(Trim(sMsg)) > 0 Then
                AFRAC_FormaPagamento = Bematech_FI_EfetuaFormaPagamentoDescricaoForma(Mid(sDescForma, 1, 16), sValor, Mid(sMsg, 1, 80))
            Else
               AFRAC_FormaPagamento = Bematech_FI_EfetuaFormaPagamento(Mid(sDescForma, 1, 16), sValor)
            End If
            
            If AFRAC_FormaPagamento = 1 Then
                AFRAC_FormaPagamento = SUCESSO
            ElseIf AFRAC_FormaPagamento = 0 Then
                AFRAC_FormaPagamento = 1
            End If
        
        Case IMPRESSORA_DARUMA_FS600
        
            sFormaPagto = sDescForma
        
'                If left(sDescForma, 6) = "Cartão" Then
'                    sFormaPagto = "Cartão"
'                Else
'                    sFormaPagto = sDescForma
'                End If
            
            
                If Len(Trim(sMsg)) > 0 Then
'                    AFRAC_FormaPagamento = Daruma_FI_EfetuaFormaPagamentoDescricaoForma(Mid(sFormaPagto, 1, 16), sValor, Mid(sMsg, 1, 48))
                    AFRAC_FormaPagamento = iCFEfetuarPagamento_ECF_Daruma(Mid(sFormaPagto, 1, 15), sValor, Mid(sMsg, 1, 84))
                Else
'                   AFRAC_FormaPagamento = Daruma_FI_EfetuaFormaPagamento(Mid(sFormaPagto, 1, 16), sValor)
                   AFRAC_FormaPagamento = iCFEfetuarPagamentoFormatado_ECF_Daruma(Mid(sFormaPagto, 1, 15), sValor)
                End If
            
            If AFRAC_FormaPagamento = 1 Then
                AFRAC_FormaPagamento = SUCESSO
            ElseIf AFRAC_FormaPagamento = 0 Then
                AFRAC_FormaPagamento = 1
            End If
        
        Case 3
            
            AFRAC_FormaPagamento = Bematech_FI_EfetuaFormaPagamento(Mid(sDescForma, 1, 16), sValor)
            
            If AFRAC_FormaPagamento = 1 Then
                AFRAC_FormaPagamento = SUCESSO
            ElseIf AFRAC_FormaPagamento = 0 Then
                AFRAC_FormaPagamento = 1
            End If
   
        Case IMPRESSORA_ELGIN
            For Each objTiposMeiosPagtos In X.gcolTiposMeiosPagtos
                If CInt(sIndice) = objTiposMeiosPagtos.iTipo Then
                    iStatus = Elgin_Pagamento(Format(objTiposMeiosPagtos.iIndice, "00"), Format(CDbl(sValor) * 100, "000000000000000"), "1")
                    
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    Exit For
                End If
            Next
   
            If iStatus = SUCESSO Then
                AFRAC_FormaPagamento = SUCESSO
            Else
                AFRAC_FormaPagamento = 1
            End If
   
        Case Else
            AFRAC_FormaPagamento = SUCESSO
        
   
   End Select

   Exit Function
   
Erro_AFRAC_FormaPagamento:

    AFRAC_FormaPagamento = 1
    
    Exit Function

End Function

Public Function AFRAC_CancelarItem(iNumeroItem As Integer) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String

'Cancela o Item Iformado no Parâmetro
On Error GoTo Erro_AFRAC_CancelarItem

    AFRAC_CancelarItem = 1

    Select Case X.giCodModeloECF
 
        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            If iNumeroItem = 0 Then
                AFRAC_CancelarItem = Bematech_FI_CancelaItemAnterior()
            Else
                AFRAC_CancelarItem = Bematech_FI_CancelaItemGenerico(CStr(iNumeroItem))
            End If
            
            If AFRAC_CancelarItem = 1 Then
                AFRAC_CancelarItem = SUCESSO
            ElseIf AFRAC_CancelarItem = 0 Then
                AFRAC_CancelarItem = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

            If iNumeroItem = 0 Then
'                    AFRAC_CancelarItem = Daruma_FI_CancelaItemAnterior()
                AFRAC_CancelarItem = iCFCancelarUltimoItem_ECF_Daruma()
                
            Else
'                    AFRAC_CancelarItem = Daruma_FI_CancelaItemGenerico(CStr(iNumeroItem))
                AFRAC_CancelarItem = iCFCancelarItem_ECF_Daruma(CStr(iNumeroItem))
            End If
            
            If AFRAC_CancelarItem = 1 Then
                AFRAC_CancelarItem = SUCESSO
            ElseIf AFRAC_CancelarItem = 0 Then
                AFRAC_CancelarItem = 1
            End If
            
            
        Case IMPRESSORA_ELGIN
            iStatus = Elgin_CancelamentoItem(Format(iNumeroItem, "000"))

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_CancelarItem = SUCESSO
            Else
                AFRAC_CancelarItem = 1
            End If
            
        Case Else
            AFRAC_CancelarItem = SUCESSO
        
            
   End Select

   Exit Function
   
Erro_AFRAC_CancelarItem:

    AFRAC_CancelarItem = 1
    
    Exit Function

End Function

Public Function AFRAC_FecharCupom(ByVal objTela As Object, ByVal objVenda As Object, ByVal bVinculado As Boolean, ByVal sCPFCNPJ As String, ByVal sNomeCliente As String, ByVal sEndereco As String, Optional ByVal bCupomAdicional As Boolean, Optional ByVal sMsg As String, Optional ByVal sVendedor As String) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim iFlag As Integer
Dim sRetorno As String
Dim sBuffer As String
Dim objSAT As Object, lErro As Long, Y As New ClassSATGlobal
Dim sMsg1 As String
Dim sData As String
Dim sHora As String
Dim sData1 As String
Dim sSequencialECF As String
Dim objNFe As Object, sMsgTotTrib As String, sNFCeMensagem As String
Dim iCodModeloECF As Integer

'Finaliza o Cupom Fiscal
On Error GoTo Erro_AFRAC_FecharCupom

    AFRAC_FecharCupom = 1
    
    If X.gsUF = "MG" Then
    
        lErro = AFRAC_DataHoraUltimoDocumento(sData, sHora)
        If lErro <> SUCESSO Then
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DataHoraUltimoDocumento")
            If lErro <> SUCESSO Then gError 214128
        End If

    
        sMsg1 = Chr(10) & Chr(13) & Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "MINAS LEGAL: " & X.gsCNPJ & " " & sData & " " & Format(CDbl(Format(objVenda.objCupomFiscal.dValorTotal, "Fixed")) * 100, "########"))
        
    ElseIf X.gsUF = "RJ" Then
    
        lErro = AFRAC_DataHoraUltimoDocumento(sData, sHora)
        If lErro <> SUCESSO Then
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DataHoraUltimoDocumento")
            If lErro <> SUCESSO Then gError 214129
        End If
    
        sData1 = left(sData, 2) & Mid(sData, 3, 2) & right(sData, 2)
    
        lErro = AFRAC_SequencialECF(sSequencialECF)
        If lErro <> SUCESSO Then
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Retorna Sequencial ECF")
            If lErro <> SUCESSO Then gError 214130
        End If
    
'        sMsg1 = Chr(10) & Chr(13) & Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "CUPOM MANIA, CONCORRA A PREMIOS")
'        sMsg1 = sMsg1 & Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "ENVIE SMS P/6789: " & Format(X.gsInscricaoEstadual, "00000000") & sData1 & Format(objVenda.objCupomFiscal.lNumero, "000000") & Format(sSequencialECF, "000"))
        sMsg1 = Chr(10) & Chr(13) & Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "PROCON - R. da Ajuda 5 - RJ - (21) 151")
        sMsg1 = sMsg1 & Chr(10) & Chr(13) & Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "ALERJ - Rua 1o de Março - RJ - (21) 25881418")
        
        sNFCeMensagem = "PROCON - R. da Ajuda 5 - RJ - (21) 151. ALERJ - Rua 1o de Março - RJ - (21) 25881418."
    
    ElseIf X.gsUF = "PB" Then
    
        lErro = AFRAC_DataHoraUltimoDocumento(sData, sHora)
        If lErro <> SUCESSO Then
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DataHoraUltimoDocumento")
            If lErro <> SUCESSO Then gError 214131
        End If
    
        sData1 = left(sData, 2) & Mid(sData, 3, 2) & right(sData, 2)
    
        lErro = AFRAC_SequencialECF(sSequencialECF)
        If lErro <> SUCESSO Then
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Retorna Sequencial ECF")
            If lErro <> SUCESSO Then gError 214132
        End If
    
        sMsg1 = Chr(10) & Chr(13) & Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "PARAIBA LEGAL - RECEITA CIDADÃ")
        sMsg1 = sMsg1 & Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", "TORPEDO PREMIADO:")
        sMsg1 = sMsg1 & Formata_Campo(ALINHAMENTO_DIREITA, 48, " ", Format(X.gsCNPJ, "00000000000000") & " " & sData & " " & Format(CDbl(Format(objVenda.objCupomFiscal.dValorTotal, "Fixed")) * 100, "########") & " " & objVenda.objCupomFiscal.sCPFCGC)
    
    End If
    
    If objVenda.objCupomFiscal.objNF.lNumNotaFiscal <> 0 Then
        sMsg1 = sMsg1 & "NF:" & CStr(objVenda.objCupomFiscal.objNF.lNumNotaFiscal)
    End If

    'Mensagem referente ao total aproximado dos tributos lei 12741
    Call CF_ECF("Venda_MsgValAproxTrib", objVenda, sMsgTotTrib)
    sMsg1 = sMsg1 & Chr(10) & Chr(13) & sMsgTotTrib

    If objVenda.iCodModeloECF = 0 Then
        iCodModeloECF = X.giCodModeloECF
    Else
        iCodModeloECF = objVenda.iCodModeloECF
    End If

    Select Case iCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            iStatus = Bematech_FI_FlagsFiscais(iFlag)
        
            'se retorno ok e o cupom estiver aberto
            If iStatus = 1 And (iFlag And 1) Then

                AFRAC_FecharCupom = Bematech_FI_TerminaFechamentoCupom(left(Bematech_AjustaTextoImpressora("MD5:" & Formata_Campo(ALINHAMENTO_DIREITA, 44, " ", X.gsMD5PAFECF)) & IIf(Len(sMsg) <> 0, Bematech_AjustaTextoImpressora(sMsg), "") & Bematech_AjustaTextoImpressora(sMsg1) & IIf(Len(sVendedor) <> 0, Bematech_AjustaTextoImpressora(sVendedor), "") & IIf(Len(X.gsMensagemCupom) <> 0, Bematech_AjustaTextoImpressora(X.gsMensagemCupom), ""), 342))
'                AFRAC_FecharCupom = Bematech_FI_TerminaFechamentoCupom("MD5:" & Formata_Campo(ALINHAMENTO_DIREITA, 44, " ", X.gsMD5PAFECF)) & IIf(Len(sMsg) <> 0, sMsg, "") & IIf(Len(sMsg1) <> 0, sMsg1, "") & IIf(Len(X.gsMensagemCupom) <> 0, X.gsMensagemCupom, "")

                If AFRAC_FecharCupom = 1 Then
                    AFRAC_FecharCupom = SUCESSO
                ElseIf AFRAC_FecharCupom = 0 Then
                    AFRAC_FecharCupom = 1
                End If
                
            End If
            
    Case IMPRESSORA_DARUMA_FS600
            
'            iStatus = Daruma_FI_FlagsFiscais(iFlag)

                sBuffer = Space(1)
                iStatus = rRetornarInformacao_ECF_Daruma("57", sBuffer)
                
                iFlag = 0
                
                If sBuffer <> "0" Then
                    iFlag = 1
                End If


            'se retorno ok e o cupom estiver aberto
            If iStatus = 1 And (iFlag = 1) Then

                If Len(sCPFCNPJ) = 0 Then sCPFCNPJ = " "
                If Len(sNomeCliente) = 0 Then sNomeCliente = " "
                If Len(sEndereco) = 0 Then sEndereco = " "

'                iStatus = Daruma_FI_IdentificaConsumidor(sNomeCliente, sEndereco, sCPFCNPJ)
                iStatus = iCFIdentificarConsumidor_ECF_Daruma(left(sNomeCliente, 30), sEndereco, sCPFCNPJ)


                'o formata campo foi colocado para preencher o requisito IX que requer as duas primeiras linhas exclusivas para o MD5
'                AFRAC_FecharCupom = Daruma_FI_TerminaFechamentoCupom(Formata_Campo(ALINHAMENTO_DIREITA, 92, " ", X.gsMD5PAFECF) & sMsg & X.gsMensagemCupom)
                AFRAC_FecharCupom = iCFEncerrar_ECF_Daruma("0", left("MD5:" & Formata_Campo(ALINHAMENTO_DIREITA, 44, " ", X.gsMD5PAFECF) & sMsg & sMsg1 & sVendedor & X.gsMensagemCupom, 384))
                
                If AFRAC_FecharCupom = 1 Then
                    AFRAC_FecharCupom = SUCESSO
                ElseIf AFRAC_FecharCupom = 0 Then
                    AFRAC_FecharCupom = 1
                End If
            End If
            
        Case IMPRESSORA_ELGIN
            sMsg = "A"
            
            iStatus = Elgin_FechaCupomFiscal("S" & Format(Len(sMsg & X.gsMensagemCupom), "000"), sMsg & sMsg1 & sVendedor & X.gsMensagemCupom)
            
            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)
            
            If iStatus = SUCESSO Then
                AFRAC_FecharCupom = SUCESSO
            Else
                AFRAC_FecharCupom = 1
            End If
            
        Case IMPRESSORA_SAT_2_5_15
                
            AFRAC_FecharCupom = SUCESSO
            
            If objVenda.iTipo = OPTION_CF Then
            
                If Not (objTela Is Nothing) Then
                    
                    If Len(objVenda.objCupomFiscal.sSATChaveAcesso) = 0 Then
                    
                        'completa objVenda para ser enviado ao SAT
                        lErro = CF_ECF("Venda_Prepara_SAT", objVenda)
                        If lErro <> SUCESSO Then
                            AFRAC_FecharCupom = 1
                        Else
                        
                            Set objSAT = CreateObject("SGESAT.SATGeral")
                        
                            Call objSAT.Inicializar(Y.gobjSATInfo.sNomeArqDLL, Y.gobjSATInfo.sCodigoDeAtivacao)
                            Call objSAT.EnviarDadosVenda(objVenda)
                        
                            Set objSAT = Nothing
                            
                            If Len(objVenda.objCupomFiscal.sSATChaveAcesso) = 0 Then
                            
                                If objTela.Name <> "Pagamento" Then
                                    objTela.Controls("Exibe").Caption = "Erro no envio do SAT-CFe..."
                                    'DoEvents
                                End If
                                
                                AFRAC_FecharCupom = 1
                                
                            Else
                            
                                lErro = CF_ECF("Venda_Grava_SAT", objVenda)
                                If lErro <> SUCESSO Then
                                    AFRAC_FecharCupom = 1
                                End If
                                
                            End If
                        
                        End If
                                        
                    End If
                
                    If AFRAC_FecharCupom = SUCESSO And Len(objVenda.objCupomFiscal.sSATChaveAcesso) <> 0 Then
                        
                        If objTela.Name <> "Pagamento" Then
                            objTela.Controls("Exibe").Caption = "Imprimindo SAT-CFe..."
                            'DoEvents
                        End If
                        
                        Call CF_ECF("SAT_Imprime", objTela, objVenda)
                        
                    End If
            
                End If
        
            End If
        
        Case IMPRESSORA_NFCE
        
            AFRAC_FecharCupom = SUCESSO
        
            If objVenda.iTipo = OPTION_CF Then
            
                If Not (objTela Is Nothing) Then
                    
                    If Len(objVenda.objCupomFiscal.sNFeChaveAcesso) = 0 Then
                    
                        objVenda.objCupomFiscal.sNFCeMensagem = IIf(Len(Trim(sMsg)) > 0, Trim(sMsg) & ".", "") & IIf(Len(Trim(sVendedor)) > 0, Trim(sVendedor) & ".", "") & sNFCeMensagem
                    
                        'completa objVenda para ser enviado
                        objVenda.objCupomFiscal.bEditavel = False
                        lErro = CF_ECF("Venda_Prepara_NFCE", objVenda)
                        If lErro <> SUCESSO Then
                            AFRAC_FecharCupom = 1
                        Else
                        
                            If objTela.Name <> "Pagamento" Then
                                objTela.Controls("Exibe").Caption = "Autorizando NFCe..."
                                'DoEvents
                            End If
                            
                            Set objNFe = CreateObject(NFeGlob_objNFe.PGM_SGENFE & ".NFeEnvio")
                        
                            lErro = objNFe.Envia_NFCE(objVenda)
                            If lErro <> SUCESSO Then
                                
                                If objTela.Name <> "Pagamento" Then
                                    objTela.Controls("Exibe").Caption = "Erro no envio da NFCe..."
                                    'DoEvents
                                End If
                                
                                Call CF_ECF("Grava_Venda_Atual", objVenda)
                                
                                AFRAC_FecharCupom = 1
                                
                            End If
                            
                            Set objNFe = Nothing
                        
                            If Len(objVenda.objCupomFiscal.sNFeChaveAcesso) <> 0 And AFRAC_FecharCupom = SUCESSO Then
                            
                                lErro = CF_ECF("Venda_Grava_NFCE", objVenda)
                                If lErro <> SUCESSO Then
                                    AFRAC_FecharCupom = 1
                                End If
                            
                            Else
                            
                                AFRAC_FecharCupom = 1
                            
                            End If
                            
                        End If
                                        
                    End If
                    
                    If objVenda.objCupomFiscal.bNFCEImprimir And AFRAC_FecharCupom = SUCESSO And Len(objVenda.objCupomFiscal.sNFeChaveAcesso) <> 0 Then
                                                        
                        If objTela.Name <> "Pagamento" Then
                            objTela.Controls("Exibe").Caption = "Imprimindo NFCe..."
                            'DoEvents
                        End If
                        
                        Call CF_ECF("NFCE_Imprime", objTela, objVenda)
                        
                    End If
                        
                End If
        
            End If
            
        Case IMPRESSORA_NFE
            
            AFRAC_FecharCupom = SUCESSO
            
            If objVenda.iTipo = OPTION_CF Then
            
                If Not (objTela Is Nothing) Then
                    
                    If Len(objVenda.objCupomFiscal.sNFeChaveAcesso) = 0 Then
                    
                        'completa objVenda para ser enviado
                        lErro = CF_ECF("Venda_Prepara_NFe", objVenda)
                        If lErro <> SUCESSO Then
                            AFRAC_FecharCupom = 1
                        Else
                        
                            Set objNFe = CreateObject(NFeGlob_objNFe.PGM_SGENFE & ".NFeEnvio")
                        
                            Call objNFe.Envia_NFe(objVenda)
                            
                            Set objNFe = Nothing
                        
                            If Len(objVenda.objCupomFiscal.sNFeChaveAcesso) <> 0 Then
                            
                                lErro = CF_ECF("Venda_Grava_NFe", objVenda)
                                If lErro <> SUCESSO Then
                                    AFRAC_FecharCupom = 1
                                End If
                            
                            Else
                            
                                AFRAC_FecharCupom = 1
                                    
                            End If
                            
                        End If
                                        
                    End If
                
                    If AFRAC_FecharCupom = SUCESSO And Len(objVenda.objCupomFiscal.sNFeChaveAcesso) <> 0 Then Call CF_ECF("NFe_Imprime", objTela, objVenda)
            
                End If
        
            End If
            
        Case Else
            
            AFRAC_FecharCupom = SUCESSO
        
            
   End Select

   Exit Function
   
Erro_AFRAC_FecharCupom:

    AFRAC_FecharCupom = 1
    
    Exit Function

End Function

Public Function AFRAC_CancelarCupom(ByVal objTela As Object, ByVal objVenda As Object) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String, objNFe As Object
Dim objSAT As Object, lErro As Long, Y As New ClassSATGlobal
Dim Z As New ClassNFeGlobal
Dim iCodModeloECF As Integer

On Error GoTo Erro_AFRAC_CancelarCupom

'O senão for passado o Numero do Cupom Fiscal o ECF vai Cancelar o Ultimo Cupom ou o Cupom Em Aberto

    AFRAC_CancelarCupom = 1

    If objVenda.iCodModeloECF = 0 Then
        iCodModeloECF = X.giCodModeloECF
    Else
        iCodModeloECF = objVenda.iCodModeloECF
    End If

    Select Case iCodModeloECF

        Case 1 To 3, 5
        
            AFRAC_CancelarCupom = Bematech_FI_CancelaCupom()
            
            If AFRAC_CancelarCupom = 1 Then
                AFRAC_CancelarCupom = SUCESSO
                
            ElseIf AFRAC_CancelarCupom = 0 Then
                AFRAC_CancelarCupom = 1
            End If
            
            
        Case IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            'estorna cupom vinculado se houver
            AFRAC_CancelarCupom = Bematech_FI_EstornoNaoFiscalVinculadoMFD("", "", "")
            AFRAC_CancelarCupom = Bematech_FI_FechaComprovanteNaoFiscalVinculado()

            AFRAC_CancelarCupom = Bematech_FI_CancelaCupom()
            
            If AFRAC_CancelarCupom = 1 Then
                AFRAC_CancelarCupom = SUCESSO
                
            ElseIf AFRAC_CancelarCupom = 0 Then
                AFRAC_CancelarCupom = 1
            End If
            

            
            
        Case IMPRESSORA_DARUMA_FS600

'                AFRAC_CancelarCupom = Daruma_FI_CancelaCupom()
            AFRAC_CancelarCupom = iCFCancelar_ECF_Daruma()
            
            If AFRAC_CancelarCupom = 1 Then
                AFRAC_CancelarCupom = SUCESSO
                
            ElseIf AFRAC_CancelarCupom = 0 Then
                AFRAC_CancelarCupom = 1
            End If
            
        Case IMPRESSORA_ELGIN
        
            iStatus = Elgin_CancelaCupomFiscal()
            
            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)
            
            If iStatus = SUCESSO Then
                AFRAC_CancelarCupom = SUCESSO
            Else
                AFRAC_CancelarCupom = 1
            End If
            
        Case IMPRESSORA_NFCE
        
            AFRAC_CancelarCupom = SUCESSO
            
            If objVenda.iTipo = OPTION_CF Then

                If Not (objTela Is Nothing) Then
                    
                    If Len(objVenda.objCupomFiscal.sNFeChaveAcesso) <> 0 Then
                    
                        'se estiver em contingencia offline
                        If Z.gobjNFeInfo.iEmContingencia <> 0 Then
                        
                            objVenda.objCupomFiscal.iNFeCancPendente = 1
                        
                            lErro = CF_ECF("NFCE_Canc_Grava", objVenda)
                            If lErro <> SUCESSO Then AFRAC_CancelarCupom = 1
                            
                            Call CF_ECF("NFCE_Imprime_Canc", objTela, objVenda)
            
                        Else
                        
                            Set objNFe = CreateObject(NFeGlob_objNFe.PGM_SGENFE & ".NFeCancela")
                        
                            lErro = objNFe.NFCE_Cancela(objVenda)
                            If lErro <> SUCESSO Then AFRAC_CancelarCupom = 1
                            
                            Set objNFe = Nothing
                            
                            If objVenda.objCupomFiscal.iNFeCancHomologado = 0 Or AFRAC_CancelarCupom = 1 Then
                            
                                AFRAC_CancelarCupom = 1
                                
                            Else
                            
                                lErro = CF_ECF("NFCE_Canc_Grava", objVenda)
                                If lErro <> SUCESSO Then
                                
                                    AFRAC_CancelarCupom = 1
                                    
                                Else
                                
                                    Call CF_ECF("NFCE_Imprime_Canc", objTela, objVenda)
                                    
                                End If
                            
                            End If
                            
                        End If
                        
                    End If
                
                End If
        
            End If
        
        Case IMPRESSORA_SAT_2_5_15
                
            AFRAC_CancelarCupom = SUCESSO
            
            If objVenda.iTipo = OPTION_CF Then

                If Not (objTela Is Nothing) Then
                    
                    If Len(objVenda.objCupomFiscal.sSATChaveAcesso) <> 0 Then
                    
                        Set objSAT = CreateObject("SGESAT.SATGeral")
                    
                        Call objSAT.Inicializar(Y.gobjSATInfo.sNomeArqDLL, Y.gobjSATInfo.sCodigoDeAtivacao)
                        Call objSAT.CancelarUltimaVenda(objVenda)
                        
                        Set objSAT = Nothing
                        
                        If Len(Trim(objVenda.objCupomFiscal.sSATCancChaveAcesso)) = 0 Then
                        
                            AFRAC_CancelarCupom = 1
                        
                        Else
                        
                            lErro = CF_ECF("SAT_Canc_Grava", objVenda)
                            If lErro <> SUCESSO Then
                            
                                AFRAC_CancelarCupom = 1
                                
                            Else
                            
                                Call CF_ECF("SAT_Imprime_Canc", objTela, objVenda)
                                
                            End If
                        
                        End If
                        
                    End If
                
                End If
        
            End If
        
            AFRAC_CancelarCupom = SUCESSO
        
        Case Else
            AFRAC_CancelarCupom = SUCESSO
            
   End Select

    If AFRAC_CancelarCupom = SUCESSO Then
        Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)
    End If

   Exit Function
   
Erro_AFRAC_CancelarCupom:

    AFRAC_CancelarCupom = 1
    
    Exit Function

End Function

Public Function AFRAC_AcrecimoDescontoItem(iAcres_Desc As Integer, iPerc_valor As Integer, dValor_AcresDesc As Double, sDescricao As String) As Integer


'Registra o acrescimo ou Desconto ao Item Anterior

    AFRAC_AcrecimoDescontoItem = SUCESSO


End Function

Public Function AFRAC_InformarCNPJIE(sCNPJ As String, sIE As String) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sResp As String
Dim lTamanho As Long
Dim sRetorno As String

On Error GoTo Erro_AFRAC_InformarCNPJIE

    AFRAC_InformarCNPJIE = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
    
            sCNPJ = String(18, 0)
            sIE = String(15, 0)
            
            AFRAC_InformarCNPJIE = Bematech_FI_CGC_IE(sCNPJ, sIE)
            
            sCNPJ = StringZ(sCNPJ)
            sIE = StringZ(sIE)
            
            If AFRAC_InformarCNPJIE = 1 Then
                AFRAC_InformarCNPJIE = SUCESSO
            ElseIf AFRAC_InformarCNPJIE = 0 Then
                AFRAC_InformarCNPJIE = 1
            End If
            
            
        Case IMPRESSORA_DARUMA_FS600
    
            sCNPJ = String(18, 0)
            sIE = String(15, 0)
            
'                AFRAC_InformarCNPJIE = Daruma_FI_CGC_IE(sCNPJ, sIE)
            sCNPJ = String(20, 0)
            
            AFRAC_InformarCNPJIE = rRetornarInformacao_ECF_Daruma("90", sCNPJ)
            
            If AFRAC_InformarCNPJIE = 1 Then
            
                sIE = String(20, 0)

                AFRAC_InformarCNPJIE = rRetornarInformacao_ECF_Daruma("91", sIE)
            End If
                    
            If AFRAC_InformarCNPJIE = 1 Then
                AFRAC_InformarCNPJIE = SUCESSO
            ElseIf AFRAC_InformarCNPJIE = 0 Then
                AFRAC_InformarCNPJIE = 1
            End If
            
            
        Case IMPRESSORA_ELGIN
        
            lTamanho = 20
            sCNPJ = String(lTamanho, 0)
        
        
'            Call GetPrivateProfileString(APLICACAO_EMPRESA, "CNPJ", CONSTANTE_ERRO, sCNPJ, lTamanho + 1, NOME_ARQUIVO_CAIXA)
'            If sCNPJ <> String(lTamanho, 0) Then
'                sCNPJ = StringZ(sCNPJ)
'                AFRAC_InformarCNPJIE = SUCESSO
'            Else
'                AFRAC_InformarCNPJIE = 1
'            End If
                    
        
            iStatus = Elgin_EcfPar("88")

            sResp = Space(100)
            iStatus = Elgin_EsperaResposta(sResp)

            If iStatus = SUCESSO Then
                sResp = Mid(sResp, 6)

                sCNPJ = Mid(sResp, 1, 18)
                sIE = Mid(sResp, 19, 15)

                AFRAC_InformarCNPJIE = SUCESSO
            Else
                AFRAC_InformarCNPJIE = 1
            End If
            
        Case Else
            AFRAC_InformarCNPJIE = SUCESSO
        
   End Select

   Exit Function
   
Erro_AFRAC_InformarCNPJIE:

    AFRAC_InformarCNPJIE = 1
    
    Exit Function

End Function

Public Function AFRAC_InformarCNPJ(sCNPJ As String) As Integer


'Informa ao sistema o CNPJ do cliente

    AFRAC_InformarCNPJ = SUCESSO


End Function

Public Function AFRAC_InformarInscricao(sIE As String) As Integer


'Informa ao ECF a inscrição

    AFRAC_InformarInscricao = SUCESSO


End Function

Public Function AFRAC_InformarRazaoSocial(sRazaoSocial As String) As Integer


'Informa ao ECF a Razão Soicial ou o Nome do Cliente

    AFRAC_InformarRazaoSocial = SUCESSO


End Function


Public Function AFRAC_InformarEndereco(sEndereco As String) As Integer


'Informa ao ECF o Endereço do Cliente

    AFRAC_InformarEndereco = SUCESSO


End Function

Public Function AFRAC_InformarOperador(sNomeOperador As String) As Integer


'Informa o Nome do Operador

    AFRAC_InformarOperador = SUCESSO


End Function

Public Function AFRAC_InformarGerente(sNomeGerente As String) As Integer


'Informa o Nome do Gernete

    AFRAC_InformarGerente = SUCESSO


End Function

Public Function AFRAC_InformarVendedor(sNomeVendedor As String) As Integer


'Informa o Nome do Vendedor

    AFRAC_InformarVendedor = SUCESSO


End Function

Public Function AFRAC_InformarMensagemCupom(sLinhaMenssagem As String) As Integer


'Informa A menssagem Impressa no Cupom Fiscal

    AFRAC_InformarMensagemCupom = SUCESSO


End Function

'Relatórios Fiscais

Public Function AFRAC_ReducaoZ(dtData As Date) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String

'Imprime o Relatório da ReducaoZ

'    Select Case X.giCodModeloECF
'
'        Case 1 To 3, 5, IMPRESSORA_DARUMA_FS600
'
'            If X.giCodModeloECF = IMPRESSORA_DARUMA_FS600 Then
'                AFRAC_ReducaoZ = Daruma_FI_ReducaoZ(Format(dtData, "dd/mm/yy"), CStr(Time))
'            Else
'                AFRAC_ReducaoZ = Bematech_FI_ReducaoZ(Format(dtData, "dd/mm/yy"), CStr(Time))
'            End If
'
'            If AFRAC_ReducaoZ = 1 Then
'                AFRAC_ReducaoZ = SUCESSO
'            ElseIf AFRAC_ReducaoZ = 0 Then
'                AFRAC_ReducaoZ = 1
'            End If
'
'
'        Case IMPRESSORA_ELGIN
'            iStatus = Elgin_ReducaoZ("0")
'
'            sRetorno = Space(50)
'            iStatus = Elgin_EsperaResposta(sRetorno)
'
'            If iStatus = SUCESSO Or iStatus = -29 Then
'                AFRAC_ReducaoZ = SUCESSO
'            Else
'                AFRAC_ReducaoZ = 1
'            End If
'
'
'   End Select

End Function

Public Function AFRAC_LeituraX() As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim lErro As Long

'Imprime o Relatório da Leitura X
On Error GoTo Erro_AFRAC_LeituraX

    AFRAC_LeituraX = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            AFRAC_LeituraX = Bematech_FI_LeituraX()
            
            If AFRAC_LeituraX = 1 Then
                AFRAC_LeituraX = SUCESSO
            ElseIf AFRAC_LeituraX = 0 Then
                AFRAC_LeituraX = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

'                AFRAC_LeituraX = Daruma_FI_LeituraX()
            AFRAC_LeituraX = iLeituraX_ECF_Daruma()

            If AFRAC_LeituraX = 1 Then
                AFRAC_LeituraX = SUCESSO
            ElseIf AFRAC_LeituraX = 0 Then
                AFRAC_LeituraX = 1
            End If
            
        Case IMPRESSORA_ELGIN
            iStatus = Elgin_LeituraX(0)
            
            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)
            
            If iStatus = SUCESSO Then
                AFRAC_LeituraX = SUCESSO
            Else
                AFRAC_LeituraX = 1
            End If
            
        Case Else
            AFRAC_LeituraX = SUCESSO
        
            
   End Select


   Exit Function
   
Erro_AFRAC_LeituraX:

    AFRAC_LeituraX = 1
    
    Exit Function

End Function

Public Function AFRAC_EmitirLeituraMemoriaFiscal(ByVal iTipoLeitura As Integer, sInicio As String, sFinal As String, ByVal iTipo As Integer, ByVal iArquivo As Integer) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim sRegistroEAD As String
Dim vbMsgRes As VbMsgBoxResult
Dim dtData As Date
Dim iCompleto As Integer


'Emite Leitura da Memória Fiscal

On Error GoTo Erro_AFRAC_EmitirLeituraMemoriaFiscal

    AFRAC_EmitirLeituraMemoriaFiscal = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            'Bematech matricial antigas
            If X.giCodModeloECF = 1 Or X.giCodModeloECF = 2 Or X.giCodModeloECF = 3 Then
                
                If iTipo = LEITURA_COMPLETA Then
                
                    If iArquivo = 0 Then
                        If iTipoLeitura = LEITURA_COMPLETA Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalData("01012000", Format(Date, "ddmmyyyy"))
                        ElseIf iTipoLeitura = LEITURA_DATAS Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalData(sInicio, sFinal)
                        ElseIf iTipoLeitura = LEITURA_REDUCOES Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalReducao(sInicio, sFinal)
                        End If
                    Else
                    
                        If iTipoLeitura = LEITURA_COMPLETA Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalSerialData("01012000", Format(Date, "ddmmyyyy"))
                        ElseIf iTipoLeitura = LEITURA_DATAS Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalSerialData(sInicio, sFinal)
                        ElseIf iTipoLeitura = LEITURA_REDUCOES Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalSerialReducao(sInicio, sFinal)
                        End If
                        
                        If AFRAC_EmitirLeituraMemoriaFiscal = 1 Then
                
                            If Len(dir(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt")) > 0 Then Kill CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt"
                
                            Name "c:\retorno.txt" As CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt"
                
                            AFRAC_EmitirLeituraMemoriaFiscal = AFRAC_CriaAssinaturaRSA(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt", sRegistroEAD)
                
                            If AFRAC_EmitirLeituraMemoriaFiscal = SUCESSO Then
                
                                vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt")
                                
                                AFRAC_EmitirLeituraMemoriaFiscal = 1
                                
                            End If
                
                
                        End If
                    
                    End If
                    
                    
                Else
                
                    MsgBox "FUNÇÃO NÃO SUPORTADA PELO MODELO DE ECF UTILIZADO"
                    
                End If
                
            'Bematech matricial nova e termicas
            ElseIf (X.giCodModeloECF = 5 Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP4200_TH_FI Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP2100_TH_FI) Then
                
                If iTipo = LEITURA_COMPLETA Then
                
                    If iArquivo = 0 Then
                
                        If iTipoLeitura = LEITURA_COMPLETA Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalDataMFD("01012000", Format(Date, "ddmmyyyy"), "c")
                        ElseIf iTipoLeitura = LEITURA_DATAS Then
                            'leitura memoria fiscal completa por data relatorio gerencial. Tela LMFCPorData
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalDataMFD(sInicio, sFinal, "c")
                        ElseIf iTipoLeitura = LEITURA_REDUCOES Then
                            'leitura memoria fiscal completa por reducoes relatorio gerencial. Tela LMFCPorReducaoZ
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalReducaoMFD(sInicio, sFinal, "c")
                        End If
                        
                    Else
                    
                        If iTipoLeitura = LEITURA_COMPLETA Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalSerialDataMFD("01012000", Format(Date, "ddmmyyyy"), "c")
                        ElseIf iTipoLeitura = LEITURA_DATAS Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalSerialDataMFD(sInicio, sFinal, "c")
                        ElseIf iTipoLeitura = LEITURA_REDUCOES Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalSerialReducaoMFD(sInicio, sFinal, "c")
                        End If
                        
                        If AFRAC_EmitirLeituraMemoriaFiscal = 1 Then
                        
                            AFRAC_EmitirLeituraMemoriaFiscal = AFRAC_CriaAssinaturaRSA("c:\Retorno.txt", sRegistroEAD)
                            
                            If AFRAC_EmitirLeituraMemoriaFiscal = SUCESSO Then
                                vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt")
                                AFRAC_EmitirLeituraMemoriaFiscal = 1
                            End If
                            
                        End If
                    
                    End If
                    
                Else
                
                    If iArquivo = 0 Then
                
                        If iTipoLeitura = LEITURA_COMPLETA Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalDataMFD("01012000", Format(Date, "ddmmyyyy"), "s")
                        ElseIf iTipoLeitura = LEITURA_DATAS Then
                            'leitura memoria fiscal simples por data relatorio gerencial. Tela LMFSPorData
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalDataMFD(sInicio, sFinal, "s")
                            'leitura memoria fiscal simples por reducoes relatorio gerencial. Tela LMFSPorReducaoZ
                        ElseIf iTipoLeitura = LEITURA_REDUCOES Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalReducaoMFD(sInicio, sFinal, "s")
                        End If
                        
                    Else
                    
                        If iTipoLeitura = LEITURA_COMPLETA Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalSerialDataMFD("01012000", Format(Date, "ddmmyyyy"), "s")
                        ElseIf iTipoLeitura = LEITURA_DATAS Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalSerialDataMFD(sInicio, sFinal, "s")
                        ElseIf iTipoLeitura = LEITURA_REDUCOES Then
                            AFRAC_EmitirLeituraMemoriaFiscal = Bematech_FI_LeituraMemoriaFiscalSerialReducaoMFD(sInicio, sFinal, "s")
                        End If
                    
                        If AFRAC_EmitirLeituraMemoriaFiscal = 1 Then
                
                            If Len(dir(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt")) > 0 Then Kill CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt"
                
                            Name "c:\retorno.txt" As CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt"
                
                            AFRAC_EmitirLeituraMemoriaFiscal = AFRAC_CriaAssinaturaRSA(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt", sRegistroEAD)
                
                            If AFRAC_EmitirLeituraMemoriaFiscal = SUCESSO Then
                
                                vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt")
                                
                                AFRAC_EmitirLeituraMemoriaFiscal = 1
                            End If
                
                
                        End If
                    
                    End If
                
                End If

            End If

            If AFRAC_EmitirLeituraMemoriaFiscal = 1 Then
                AFRAC_EmitirLeituraMemoriaFiscal = SUCESSO
            ElseIf AFRAC_EmitirLeituraMemoriaFiscal = 0 Then
                AFRAC_EmitirLeituraMemoriaFiscal = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

            If iTipo = LEITURA_COMPLETA Then
'                    Call Daruma_Registry_MFD_LeituraMFCompleta("1")
                 Call regAlterarValor_Daruma("ECF\LMFCompleta", "1")
                 iCompleto = 1
            ElseIf iTipo = LEITURA_SIMPLES Then
'                    Call Daruma_Registry_MFD_LeituraMFCompleta("0")
                 Call regAlterarValor_Daruma("ECF\LMFCompleta", "0")
                 iCompleto = 0
            End If
            
            If iArquivo = 1 Then
            
            
                If Len(dir(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt")) > 0 Then Kill CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt"
            
                If iTipoLeitura = LEITURA_COMPLETA Then
'                        AFRAC_EmitirLeituraMemoriaFiscal = Daruma_FI_LeituraMemoriaFiscalSerialData("01012000", Format(gdtDataHoje, "ddmmyyyy"))
                    AFRAC_EmitirLeituraMemoriaFiscal = iMFLerSerial_ECF_Daruma("010100", Format(Date, "ddmmyy"))
                ElseIf iTipoLeitura = LEITURA_DATAS Then
'                        AFRAC_EmitirLeituraMemoriaFiscal = Daruma_FI_LeituraMemoriaFiscalSerialData(sInicio, sFinal)
                    AFRAC_EmitirLeituraMemoriaFiscal = iMFLerSerial_ECF_Daruma(sInicio, sFinal)
                ElseIf iTipoLeitura = LEITURA_REDUCOES Then
'                        AFRAC_EmitirLeituraMemoriaFiscal = Daruma_FI_LeituraMemoriaFiscalSerialReducao(sInicio, sFinal)
                    AFRAC_EmitirLeituraMemoriaFiscal = iMFLerSerial_ECF_Daruma(sInicio, sFinal)
                End If
            
                If AFRAC_EmitirLeituraMemoriaFiscal = 1 Then
            
                    AFRAC_EmitirLeituraMemoriaFiscal = AFRAC_CriaAssinaturaRSA(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt", sRegistroEAD)
            
                    If AFRAC_EmitirLeituraMemoriaFiscal = SUCESSO Then
                        vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt")
                        AFRAC_EmitirLeituraMemoriaFiscal = 1
                    End If
            
            
                End If
                
                
            '*** ATO COTEPE *****
            ElseIf iArquivo = 2 Then
            
            
                If iTipoLeitura = LEITURA_COMPLETA Then
'                        AFRAC_EmitirLeituraMemoriaFiscal = Daruma_FI_LeituraMemoriaFiscalSerialData("01012000", Format(gdtDataHoje, "ddmmyyyy"))
                    'AFRAC_EmitirLeituraMemoriaFiscal = rGerarRelatorio_ECF_Daruma("MF+", "DATAM", "01012000", Format(gdtDataHoje, "ddmmyyyy"))
                    If Len(dir(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_DATA.TXT")) > 0 Then Kill CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_DATA.TXT"
                    
                    AFRAC_EmitirLeituraMemoriaFiscal = eMemoriaFiscal_ECF_Daruma("01012000", Format(Date, "ddmmyyyy"), iCompleto, "ATO")
                
                    If AFRAC_EmitirLeituraMemoriaFiscal = 1 Then
                
                        AFRAC_EmitirLeituraMemoriaFiscal = AFRAC_CriaAssinaturaRSA(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_DATA.TXT", sRegistroEAD)
                
                        If AFRAC_EmitirLeituraMemoriaFiscal = SUCESSO Then
                            vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_DATA.TXT")
                            AFRAC_EmitirLeituraMemoriaFiscal = 1
                        End If
                
                
                    End If
                
                
                ElseIf iTipoLeitura = LEITURA_DATAS Then
'                        dtData = StrParaDate(left(sInicio, 2) & "/" & Mid(sInicio, 3, 2) & "/" & Mid(sInicio, 5, 2))
'                        sInicio = Format(dtData, "ddmmyyyy")
'
'                        dtData = StrParaDate(left(sFinal, 2) & "/" & Mid(sFinal, 3, 2) & "/" & Mid(sFinal, 5, 2))
'                        sFinal = Format(dtData, "ddmmyyyy")
                                            
                                            
'                        AFRAC_EmitirLeituraMemoriaFiscal = Daruma_FI_LeituraMemoriaFiscalSerialData(sInicio, sFinal)
                    'AFRAC_EmitirLeituraMemoriaFiscal = rGerarRelatorio_ECF_Daruma("MF+", "DATAM", sInicio, sFinal)
                    If Len(dir(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_DATA.TXT")) > 0 Then Kill CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_DATA.TXT"
                    
                    AFRAC_EmitirLeituraMemoriaFiscal = eMemoriaFiscal_ECF_Daruma(sInicio, sFinal, iCompleto, "ATO")
                    
                    If AFRAC_EmitirLeituraMemoriaFiscal = 1 Then
                
                        AFRAC_EmitirLeituraMemoriaFiscal = AFRAC_CriaAssinaturaRSA(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_DATA.TXT", sRegistroEAD)
                
                        If AFRAC_EmitirLeituraMemoriaFiscal = SUCESSO Then
                            vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_DATA.TXT")
                            AFRAC_EmitirLeituraMemoriaFiscal = 1
                        End If
                
                
                    End If
                    
                    
                ElseIf iTipoLeitura = LEITURA_REDUCOES Then
'                        AFRAC_EmitirLeituraMemoriaFiscal = Daruma_FI_LeituraMemoriaFiscalSerialReducao(sInicio, sFinal)
                    'AFRAC_EmitirLeituraMemoriaFiscal = rGerarRelatorio_ECF_Daruma("MF+", "DATAM", sInicio, sFinal)
                    If Len(dir(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_CRZ.TXT")) > 0 Then Kill CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_CRZ.TXT"
                    
                    
                    AFRAC_EmitirLeituraMemoriaFiscal = eMemoriaFiscal_ECF_Daruma(sInicio, sFinal, iCompleto, "ATO")
                    
                    If AFRAC_EmitirLeituraMemoriaFiscal = 1 Then
                
                        AFRAC_EmitirLeituraMemoriaFiscal = AFRAC_CriaAssinaturaRSA(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_CRZ.TXT", sRegistroEAD)
                
                        If AFRAC_EmitirLeituraMemoriaFiscal = SUCESSO Then
                            vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MF_CRZ.TXT")
                            AFRAC_EmitirLeituraMemoriaFiscal = 1
                        End If
                
                
                    End If
                    
                End If
            
            Else
            
                If iTipoLeitura = LEITURA_COMPLETA Then
'                       AFRAC_EmitirLeituraMemoriaFiscal = Daruma_FI_LeituraMemoriaFiscalData("01012000", Format(gdtDataHoje, "ddmmyyyy"))
                   AFRAC_EmitirLeituraMemoriaFiscal = iMFLer_ECF_Daruma("01012000", Format(Date, "ddmmyyyy"))
                 ElseIf iTipoLeitura = LEITURA_DATAS Then
'                        AFRAC_EmitirLeituraMemoriaFiscal = Daruma_FI_LeituraMemoriaFiscalData(sInicio, sFinal)
                    AFRAC_EmitirLeituraMemoriaFiscal = iMFLer_ECF_Daruma(sInicio, sFinal)
                ElseIf iTipoLeitura = LEITURA_REDUCOES Then
'                        AFRAC_EmitirLeituraMemoriaFiscal = Daruma_FI_LeituraMemoriaFiscalReducao(sInicio, sFinal)
                    AFRAC_EmitirLeituraMemoriaFiscal = iMFLer_ECF_Daruma(sInicio, sFinal)
                End If
                
            End If

            If AFRAC_EmitirLeituraMemoriaFiscal = 1 Then
                AFRAC_EmitirLeituraMemoriaFiscal = SUCESSO
            ElseIf AFRAC_EmitirLeituraMemoriaFiscal = 0 Then
                AFRAC_EmitirLeituraMemoriaFiscal = 1
            End If
            
        Case IMPRESSORA_ELGIN
            If iTipoLeitura = LEITURA_COMPLETA Then
                iStatus = Elgin_LeMemFiscalData("01012000", Format(Date, "ddmmyy"), "0")
            ElseIf iTipoLeitura = LEITURA_DATAS Then
                iStatus = Elgin_LeMemFiscalData(Format(CDate(sInicio), "ddmmyy"), Format(CDate(sFinal), "ddmmyy"), "0")
            ElseIf iTipoLeitura = LEITURA_REDUCOES Then
                iStatus = Elgin_LeMemFiscalReducao(Mid(sInicio, Len(sInicio) - 3, Len(sInicio)), Mid(sFinal, Len(sFinal) - 3, Len(sFinal)), "0")
            End If
   
            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)
   
            If iStatus = SUCESSO Then
                AFRAC_EmitirLeituraMemoriaFiscal = SUCESSO
            Else
                AFRAC_EmitirLeituraMemoriaFiscal = 1
            End If
            
        Case Else
            AFRAC_EmitirLeituraMemoriaFiscal = SUCESSO
            
   End Select

   Exit Function
   
Erro_AFRAC_EmitirLeituraMemoriaFiscal:

    AFRAC_EmitirLeituraMemoriaFiscal = 1
    
    Exit Function

End Function

Public Function AFRAC_GravarLeituraX(sNomeCaminhoArquivo As String) As Integer


'Grava A LeituraX , obtida Atrvés da Transmissão do Serial , no arquivo passado

    AFRAC_GravarLeituraX = SUCESSO


End Function

Public Function AFRAC_GravarLeituraMemoriaFiscal(sTipo As String, sInicio As String, sFinal As String, sNomeCaminhoArquivo As String) As Integer


'Grava A Leitura da Memória Fiscal , no arquivo passado

    AFRAC_GravarLeituraMemoriaFiscal = SUCESSO


End Function

'Operações não fiscais

Public Function AFRAC_AbrirVinculado(lCOO As Long, sDescFormaPag As String, dValor As Double, ByVal iNumParcelas As Integer) As Integer

Dim X As New ClassECFConfig
Dim sValor As String
Dim sCOO As String
Dim iStatus As Integer
Dim sRetorno As String
Dim sNumParcelas As String


'Função que Faz a AberTura de Vinculado da Afrac

On Error GoTo Erro_AFRAC_AbrirVinculado

    AFRAC_AbrirVinculado = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            sValor = CStr(Format(dValor, "standard"))
            sCOO = CStr(lCOO)
            sNumParcelas = CStr(iNumParcelas)
            
            AFRAC_AbrirVinculado = Bematech_FI_AbreComprovanteNaoFiscalVinculado(left(sDescFormaPag, 16), sValor, sCOO)
            
            If AFRAC_AbrirVinculado = 1 Then
                'Atualiza o arquivo
                Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)
                AFRAC_AbrirVinculado = SUCESSO
            ElseIf AFRAC_AbrirVinculado = 0 Then
                AFRAC_AbrirVinculado = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

            sValor = CStr(Format(dValor, "standard"))
            sCOO = CStr(lCOO)
            sNumParcelas = CStr(iNumParcelas)
            
'                AFRAC_AbrirVinculado = Daruma_FI_AbreComprovanteNaoFiscalVinculado(left(sDescFormaPag, 16), sValor, sCOO)
            AFRAC_AbrirVinculado = iCCDAbrirSimplificado_ECF_Daruma(left(sDescFormaPag, 20), sNumParcelas, sCOO, sValor)
            
            If AFRAC_AbrirVinculado = 1 Then
                'Atualiza o arquivo
                Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)
                AFRAC_AbrirVinculado = SUCESSO
            ElseIf AFRAC_AbrirVinculado = 0 Then
                AFRAC_AbrirVinculado = 1
            End If
            
            
        Case IMPRESSORA_ELGIN
               
            iStatus = Elgin_AbreCupomVinculado

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_AbrirVinculado = SUCESSO
            Else
                AFRAC_AbrirVinculado = 1
            End If
            
        Case Else
            AFRAC_AbrirVinculado = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_AbrirVinculado:

    AFRAC_AbrirVinculado = 1
    
    Exit Function

End Function

Public Function AFRAC_ImprimirVinculado(sLinha1 As String, Optional sLinha2 As String) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer, Y As New ClassSATGlobal, Z As New ClassNFeGlobal
Dim sRetorno As String, iModeloImpressora As Integer
Dim objRel As Object, sPortaImpressora As String

'imprime uma linha no cupom fiscal não vinculado

On Error GoTo Erro_AFRAC_ImprimirVinculado

    AFRAC_ImprimirVinculado = 1

    iModeloImpressora = X.giCodModeloECF
    
    If AFRAC_ImpressoraCFe(iModeloImpressora) Then
        
        Select Case iModeloImpressora
            
            Case IMPRESSORA_SAT_2_5_15
                iModeloImpressora = Y.gobjSATInfo.iModeloImpressora
                sPortaImpressora = Y.gobjSATInfo.sPortaImpressora
                
            Case IMPRESSORA_NFCE
                iModeloImpressora = Z.gobjNFeInfo.iModeloImpressora
                sPortaImpressora = Z.gobjNFeInfo.sPortaImpressora
                
            Case Else
                iModeloImpressora = IMPRESSORA_NAO_FISCAL
                
        End Select
        
    End If
    
    Select Case iModeloImpressora

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            AFRAC_ImprimirVinculado = Bematech_FI_UsaComprovanteNaoFiscalVinculado(Bematech_AjustaTextoImpressora(sLinha1))
            
            If AFRAC_ImprimirVinculado = 1 Then
                AFRAC_ImprimirVinculado = SUCESSO
            ElseIf AFRAC_ImprimirVinculado = 0 Then
                AFRAC_ImprimirVinculado = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

'                AFRAC_ImprimirVinculado = Daruma_FI_UsaComprovanteNaoFiscalVinculado(sLinha1)
            AFRAC_ImprimirVinculado = iCCDImprimirTexto_ECF_Daruma(sLinha1)
            
            If AFRAC_ImprimirVinculado = 1 Then
                AFRAC_ImprimirVinculado = SUCESSO
            ElseIf AFRAC_ImprimirVinculado = 0 Then
                AFRAC_ImprimirVinculado = 1
            End If
            
        Case IMPRESSORA_ELGIN
               
            iStatus = Elgin_ImprimeLinhaNaoFiscal(CByte(Asc("0")), sLinha1)

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_ImprimirVinculado = SUCESSO
            Else
                AFRAC_ImprimirVinculado = 1
            End If
            
        Case IMPRESSORA_DARUMA_DR700
            iStatus = iImprimirTexto_DUAL_DarumaFramework(sLinha1, 0)
        
            If iStatus = 1 Then
                AFRAC_ImprimirVinculado = SUCESSO
            Else
                AFRAC_ImprimirVinculado = 1
            End If
        
        Case IMPRESSORA_BEMATECH_MP4000_TH
            iStatus = BematechNaoFiscal_BematechTX(sLinha1)
            If iStatus = 1 Then
                AFRAC_ImprimirVinculado = SUCESSO
            Else
                AFRAC_ImprimirVinculado = 1
            End If
            
        Case IMPRESSORA_EPSON_TM_T20
            iStatus = EpsonNaoFiscal_ImprimeTexto(sLinha1)
            If iStatus = 1 Then
                AFRAC_ImprimirVinculado = SUCESSO
            Else
                AFRAC_ImprimirVinculado = SUCESSO '???? jones acre
            End If
            
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS
            Set objRel = New ClassImpressoraESC
            Call objRel.Define_Impressora(sPortaImpressora, iModeloImpressora)
            iStatus = objRel.ImprimeNormal(sLinha1)
            iStatus = objRel.TerminaImpressao
            Set objRel = Nothing
            If iStatus = 1 Then
                AFRAC_ImprimirVinculado = SUCESSO
            Else
                AFRAC_ImprimirVinculado = 1
            End If
        
        Case Else
            AFRAC_ImprimirVinculado = SUCESSO
            
    End Select

   Exit Function
   
Erro_AFRAC_ImprimirVinculado:

    AFRAC_ImprimirVinculado = 1
    
    Exit Function

End Function

Public Function AFRAC_FecharVinculado() As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim lErro As Long
'fecha Vinculado

On Error GoTo Erro_AFRAC_FecharVinculado

    AFRAC_FecharVinculado = 1
    
    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            AFRAC_FecharVinculado = Bematech_FI_FechaComprovanteNaoFiscalVinculado()
            
            If AFRAC_FecharVinculado = 1 Then
                AFRAC_FecharVinculado = SUCESSO
                lErro = CF_ECF("Grava_R06", "CC")
                If lErro <> SUCESSO Then gError 210807
            ElseIf AFRAC_FecharVinculado = 0 Then
                AFRAC_FecharVinculado = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

'                AFRAC_FecharVinculado = Daruma_FI_FechaComprovanteNaoFiscalVinculado()
            AFRAC_FecharVinculado = iCCDFechar_ECF_Daruma()
            
            If AFRAC_FecharVinculado = 1 Then
                AFRAC_FecharVinculado = SUCESSO
                lErro = CF_ECF("Grava_R06", "CC")
                If lErro <> SUCESSO Then gError 210807
            ElseIf AFRAC_FecharVinculado = 0 Then
                AFRAC_FecharVinculado = 1
            End If
            
        Case IMPRESSORA_ELGIN
               
            iStatus = Elgin_EncerraCupomNaoFiscal

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_FecharVinculado = SUCESSO
            Else
                AFRAC_FecharVinculado = 1
            End If
            
        Case Else
                AFRAC_FecharVinculado = SUCESSO
            
    End Select

   Exit Function
   
Erro_AFRAC_FecharVinculado:

    AFRAC_FecharVinculado = 1
    
    Exit Function

End Function

Public Function AFRAC_AbrirNaoFiscalNaoVinculado() As Integer
'Abre não Fiscal não Vinculado
'NAO USADO PELO PAF-ECF

Dim X As New ClassECFConfig

    AFRAC_AbrirNaoFiscalNaoVinculado = 1
    
    Select Case X.giCodModeloECF

        Case IMPRESSORA_DARUMA_FS600

            AFRAC_AbrirNaoFiscalNaoVinculado = iCNFAbrirPadrao_ECF_Daruma()
            
            If AFRAC_AbrirNaoFiscalNaoVinculado = 1 Then
                AFRAC_AbrirNaoFiscalNaoVinculado = SUCESSO
            ElseIf AFRAC_AbrirNaoFiscalNaoVinculado = 0 Then
                AFRAC_AbrirNaoFiscalNaoVinculado = 1
            End If
            
        Case Else
                AFRAC_AbrirNaoFiscalNaoVinculado = SUCESSO
            
    End Select

   Exit Function
   
Erro_AFRAC_AbrirNaoFiscalNaoVinculado:

    AFRAC_AbrirNaoFiscalNaoVinculado = 1
    
    Exit Function

End Function

Public Function AFRAC_RegistrarNaoFiscal(sIndice As String, dValor As Double, sMensagem As String) As Integer
'Registra um valor em um determinado totalizador

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim lTamanho1 As Long
Dim sRetorno As String

'Sangria ou suprimento

On Error GoTo Erro_AFRAC_RegistrarNaoFiscal

    AFRAC_RegistrarNaoFiscal = 1
    
    Select Case X.giCodModeloECF

        Case IMPRESSORA_ELGIN

            If sIndice = X.giTotalizador_Sangria Then
                lTamanho1 = 6
                sRetorno = String(lTamanho1, 0)
            
                Call GetPrivateProfileString(APLICACAO_ECF, "TOTALIZADORSANGRIA", CONSTANTE_ERRO, sRetorno, lTamanho1 + 1, NOME_ARQUIVO_CAIXA)
                If sRetorno <> String(lTamanho1, 0) Then sIndice = sRetorno
                
            ElseIf sIndice = X.giTotalizador_Suprimento Then
                lTamanho1 = 6
                sRetorno = String(lTamanho1, 0)
            
                Call GetPrivateProfileString(APLICACAO_ECF, "TOTALIZADORSUPRIMENTO", CONSTANTE_ERRO, sRetorno, lTamanho1 + 1, NOME_ARQUIVO_CAIXA)
                If sRetorno <> String(lTamanho1, 0) Then sIndice = sRetorno
               
            End If
            
            iStatus = Elgin_OperRegNaoVinculado(CStr(CInt(sIndice) + 15), Format(dValor * 100, "000000000000000"), CByte(Asc("@")), CByte(Asc("&")), "000000000000000", "")

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_RegistrarNaoFiscal = SUCESSO
            Else
                AFRAC_RegistrarNaoFiscal = 1
            End If
            
        Case Else
            AFRAC_RegistrarNaoFiscal = SUCESSO
        
            
    End Select

   Exit Function
   
Erro_AFRAC_RegistrarNaoFiscal:

    AFRAC_RegistrarNaoFiscal = 1
    
    Exit Function

End Function

Public Function AFRAC_CancelarVinculado() As Integer
'Cancela Vinculado

    AFRAC_CancelarVinculado = SUCESSO


End Function

Public Function AFRAC_SangriaSuprimento(bSangria As Boolean, dValor As Double) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer

'Sangria ou suprimento

On Error GoTo Erro_AFRAC_SangriaSuprimento

    AFRAC_SangriaSuprimento = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            If bSangria = True Then
                AFRAC_SangriaSuprimento = Bematech_FI_Sangria(CStr(Format(dValor, "standard")))
            Else
                AFRAC_SangriaSuprimento = Bematech_FI_Suprimento(CStr(Format(dValor, "standard")), "")
            End If

            If AFRAC_SangriaSuprimento = 1 Then
                AFRAC_SangriaSuprimento = SUCESSO
                Call CF_ECF("Grava_R06", "CN")
            ElseIf AFRAC_SangriaSuprimento = 0 Then
                AFRAC_SangriaSuprimento = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

            If bSangria = True Then
'                    AFRAC_SangriaSuprimento = Daruma_FI_Sangria(CStr(Format(dValor, "standard")))
                AFRAC_SangriaSuprimento = iSangria_ECF_Daruma(CStr(Format(dValor, "standard")), "")
            Else
'                    AFRAC_SangriaSuprimento = Daruma_FI_Suprimento(CStr(Format(dValor, "standard")), "")
                AFRAC_SangriaSuprimento = iSuprimento_ECF_Daruma(CStr(Format(dValor, "standard")), "")
            End If

            If AFRAC_SangriaSuprimento = 1 Then
                AFRAC_SangriaSuprimento = SUCESSO
                Call CF_ECF("Grava_R06", "CN")
            ElseIf AFRAC_SangriaSuprimento = 0 Then
                AFRAC_SangriaSuprimento = 1
            End If
            
        Case Else
            AFRAC_SangriaSuprimento = SUCESSO
            
    End Select

   Exit Function
   
Erro_AFRAC_SangriaSuprimento:

    AFRAC_SangriaSuprimento = 1
    
    Exit Function

End Function

Public Function AFRAC_AbrirRelatorioGerencial2(sIndice As String, ByVal objTela As Object, Optional ByRef objRel As Object) As Integer
'Abre um relatorio gerencial
    
Dim X As New ClassECFConfig
Dim iStatus As Integer, Y As New ClassSATGlobal, Z As New ClassNFeGlobal, iRetornoBematech As Integer
Dim sRetorno As String, iRetornoEpson As Integer
Dim lErro As Long, iModeloImpressora As Integer
Dim sNome As String, iRetornoDual As Integer, sPortaImpressora As String

'Sangria ou suprimento
On Error GoTo Erro_AFRAC_AbrirRelatorioGerencial2

    AFRAC_AbrirRelatorioGerencial2 = 1

    iModeloImpressora = X.giCodModeloECF
    
    If AFRAC_ImpressoraCFe(iModeloImpressora) Then
        
        Select Case iModeloImpressora
            
            Case IMPRESSORA_SAT_2_5_15
                iModeloImpressora = Y.gobjSATInfo.iModeloImpressora
                sPortaImpressora = Y.gobjSATInfo.sPortaImpressora
                
            Case IMPRESSORA_NFCE
                iModeloImpressora = Z.gobjNFeInfo.iModeloImpressora
                sPortaImpressora = Z.gobjNFeInfo.sPortaImpressora
                
            Case Else
                iModeloImpressora = IMPRESSORA_NAO_FISCAL
                
        End Select
    
    End If

    Select Case iModeloImpressora
    
        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            
            If sIndice = RELGER_IDENTPAFECF Or sIndice = RELGER_PARAMCONFIG Then
                AFRAC_AbrirRelatorioGerencial2 = Bematech_FI_AbreRelatorioGerencialMFD(sIndice)

            End If

            If AFRAC_AbrirRelatorioGerencial2 = 1 Then
                AFRAC_AbrirRelatorioGerencial2 = SUCESSO
            ElseIf AFRAC_AbrirRelatorioGerencial2 = 0 Then
                    AFRAC_AbrirRelatorioGerencial2 = 1
            End If
    
    
        Case IMPRESSORA_DARUMA_DR700
'            iRetornoDual = regAlterarValor_Daruma("DUAL\PortaComunicacao", sPortaImpressora)
'            If iRetornoDual = 1 Then
'                Call regAlterarValor_Daruma("DUAL\Velocidade", "115200")
'                iRetornoDual = rStatusImpressora_DUAL_DarumaFramework
'            End If
            
'            If iRetornoDual <> 1 Then
                iRetornoDual = eBuscarPortaVelocidade_DUAL_DarumaFramework
'            End If
            
            If iRetornoDual = 1 Then
                AFRAC_AbrirRelatorioGerencial2 = SUCESSO
            ElseIf iRetornoDual = 0 Then
                AFRAC_AbrirRelatorioGerencial2 = 1
            Else
                AFRAC_AbrirRelatorioGerencial2 = iRetornoDual
            End If
            
            Set objRel = New ClassDarumaNaoFiscalImp
        
        Case IMPRESSORA_DARUMA_FS600
    
            Select Case sIndice
            
                Case CStr(RELGER_IDENTPAFECF)
                    AFRAC_AbrirRelatorioGerencial2 = iRGAbrir_ECF_Daruma("Ident. PAF-ECF")
                
                Case CStr(RELGER_PARAMCONFIG)
                    AFRAC_AbrirRelatorioGerencial2 = iRGAbrir_ECF_Daruma("Param. Config.")
    
            End Select
    
            If AFRAC_AbrirRelatorioGerencial2 = 1 Then
                AFRAC_AbrirRelatorioGerencial2 = SUCESSO
            ElseIf AFRAC_AbrirRelatorioGerencial2 = 0 Then
                    AFRAC_AbrirRelatorioGerencial2 = 1
            End If
    
        Case IMPRESSORA_NAO_FISCAL
        
            lErro = CF_ECF("Configura_Impressora_ECF", objTela)
            If lErro <> SUCESSO Then gError 204311
        
            AFRAC_AbrirRelatorioGerencial2 = SUCESSO
        
        Case IMPRESSORA_ELGIN

            iStatus = Elgin_LeituraXComRelGer()
            
            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)
            
            If iStatus = SUCESSO Then
                AFRAC_AbrirRelatorioGerencial2 = SUCESSO
            Else
                AFRAC_AbrirRelatorioGerencial2 = 1
            End If
        
        Case IMPRESSORA_BEMATECH_MP4000_TH
            iRetornoBematech = BematechNaoFiscal_ConfiguraModeloImpressora(5)
            If iRetornoBematech = 1 Then
                iRetornoBematech = BematechNaoFiscal_IniciaPorta(sPortaImpressora)
            End If
            
            If iRetornoBematech = 1 Then
                AFRAC_AbrirRelatorioGerencial2 = SUCESSO
            ElseIf iRetornoBematech = 0 Then
                AFRAC_AbrirRelatorioGerencial2 = 1
            Else
                AFRAC_AbrirRelatorioGerencial2 = iRetornoBematech
            End If
        
            Set objRel = New ClassBematechNaoFiscalImp
        
        Case IMPRESSORA_EPSON_TM_T20
        
            Call EpsonNaoFiscal_ConfiguraTaxaSerial(115200)
            
            iRetornoEpson = EpsonNaoFiscal_IniciaPorta(sPortaImpressora)
            If iRetornoEpson = 1 Then
            
                AFRAC_AbrirRelatorioGerencial2 = SUCESSO
                
            Else
            
                'a porta pode ter ficado aberta numa impressao anterior
                Call EpsonNaoFiscal_FechaPorta
                
                iRetornoEpson = EpsonNaoFiscal_IniciaPorta(sPortaImpressora)
                If iRetornoEpson = 1 Then
                    AFRAC_AbrirRelatorioGerencial2 = SUCESSO
                ElseIf iRetornoEpson = 0 Then
                    AFRAC_AbrirRelatorioGerencial2 = 1
                Else
                    AFRAC_AbrirRelatorioGerencial2 = iRetornoEpson
                End If
                
            End If
                            
            Set objRel = New ClassEpsonNaoFiscalImp
        
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS
            Set objRel = New ClassImpressoraESC
            Call objRel.Define_Impressora(sPortaImpressora, iModeloImpressora)
            AFRAC_AbrirRelatorioGerencial2 = SUCESSO
        
        Case Else
            AFRAC_AbrirRelatorioGerencial2 = SUCESSO
            
    End Select
    
    Exit Function
    
Erro_AFRAC_AbrirRelatorioGerencial2:

    Select Case gErr
    
        Case 204311
            AFRAC_AbrirRelatorioGerencial2 = 1
    
        Case Else
            AFRAC_AbrirRelatorioGerencial2 = 2
    End Select
    
End Function

Public Function AFRAC_ImprimirRelatorioGerencial2(sTexto As String, ByVal objTela As Object, Optional ByVal objRel As Object) As Integer
'Imprime o texto passado como parametro

Dim X As New ClassECFConfig
Dim iStatus As Integer, Y As New ClassSATGlobal, Z As New ClassNFeGlobal
Dim sRetorno As String, iModeloImpressora As Integer

'Sangria ou suprimento

On Error GoTo Erro_AFRAC_ImprimirRelatorioGerencial2

    AFRAC_ImprimirRelatorioGerencial2 = 1

    iModeloImpressora = X.giCodModeloECF
    
    If AFRAC_ImpressoraCFe(iModeloImpressora) Then
        
        Select Case iModeloImpressora
            
            Case IMPRESSORA_SAT_2_5_15
                iModeloImpressora = Y.gobjSATInfo.iModeloImpressora
                
            Case IMPRESSORA_NFCE
                iModeloImpressora = Z.gobjNFeInfo.iModeloImpressora
                
            Case Else
                iModeloImpressora = IMPRESSORA_NAO_FISCAL
                
        End Select
        
    End If
    
    Select Case iModeloImpressora

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            'Atualiza o arquivo
            Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)

            AFRAC_ImprimirRelatorioGerencial2 = Bematech_FI_RelatorioGerencial(Bematech_AjustaTextoImpressora(sTexto))
            
            If AFRAC_ImprimirRelatorioGerencial2 = 1 Then
                AFRAC_ImprimirRelatorioGerencial2 = SUCESSO
            ElseIf AFRAC_ImprimirRelatorioGerencial2 = 0 Then
                AFRAC_ImprimirRelatorioGerencial2 = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

            'Atualiza o arquivo
            Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)

'                AFRAC_ImprimirRelatorioGerencial2 = Daruma_FI_RelatorioGerencial(sTexto)
            AFRAC_ImprimirRelatorioGerencial2 = iRGImprimirTexto_ECF_Daruma(sTexto)
            
            If AFRAC_ImprimirRelatorioGerencial2 = 1 Then
                AFRAC_ImprimirRelatorioGerencial2 = SUCESSO
            ElseIf AFRAC_ImprimirRelatorioGerencial2 = 0 Then
                AFRAC_ImprimirRelatorioGerencial2 = 1
            End If
            
        Case IMPRESSORA_NAO_FISCAL
            If Not (objTela Is Nothing) Then objTela.Controls("RT1").Text = objTela.Controls("RT1").Text & sTexto & Chr(13) & Chr(10)
        
            AFRAC_ImprimirRelatorioGerencial2 = SUCESSO
        
        Case IMPRESSORA_ELGIN

            'Atualiza o arquivo
            Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)

            iStatus = Elgin_ImprimeLinhaNaoFiscal(0, left(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", DesacentuaTexto(sTexto)), 48))

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_ImprimirRelatorioGerencial2 = SUCESSO
            Else
                AFRAC_ImprimirRelatorioGerencial2 = 1
            End If
        
        Case IMPRESSORA_DARUMA_DR700
            If objRel Is Nothing Then
                iStatus = iImprimirTexto_DUAL_DarumaFramework(sTexto, 0)
            Else
                iStatus = objRel.ImprimeNormal(sTexto)
            End If
        
            If iStatus = 1 Then
                AFRAC_ImprimirRelatorioGerencial2 = SUCESSO
            Else
                AFRAC_ImprimirRelatorioGerencial2 = 1
            End If
        
        Case IMPRESSORA_BEMATECH_MP4000_TH
            
            If objRel Is Nothing Then
                iStatus = BematechNaoFiscal_BematechTX(sTexto)
            Else
                iStatus = objRel.ImprimeNormal(sTexto)
            End If
            
            If iStatus = 1 Then
                AFRAC_ImprimirRelatorioGerencial2 = SUCESSO
            Else
                AFRAC_ImprimirRelatorioGerencial2 = 1
            End If
            
        Case IMPRESSORA_EPSON_TM_T20
            If objRel Is Nothing Then
                iStatus = EpsonNaoFiscal_ImprimeTexto(sTexto)
            Else
                iStatus = objRel.ImprimeNormal(sTexto)
            End If
                
            If iStatus = 1 Then
                AFRAC_ImprimirRelatorioGerencial2 = SUCESSO
            Else
                AFRAC_ImprimirRelatorioGerencial2 = 1
            End If
            
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS
            iStatus = objRel.ImprimeNormal(sTexto)
                
            If iStatus = 1 Then
                AFRAC_ImprimirRelatorioGerencial2 = SUCESSO
            Else
                AFRAC_ImprimirRelatorioGerencial2 = 1
            End If
            
        Case Else
            AFRAC_ImprimirRelatorioGerencial2 = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_ImprimirRelatorioGerencial2:

    AFRAC_ImprimirRelatorioGerencial2 = 1
    
    Exit Function

End Function

Public Function AFRAC_FecharRelatorioGerencial2(ByVal objTela As Object, Optional ByVal objRel As Object) As Integer
'Encerra relatorio gerencial

Dim X As New ClassECFConfig
Dim iStatus As Integer, Y As New ClassSATGlobal, Z As New ClassNFeGlobal
Dim sRetorno As String, iModeloImpressora As Integer

On Error GoTo Erro_AFRAC_FecharRelatorioGerencial2

    AFRAC_FecharRelatorioGerencial2 = 1

    iModeloImpressora = X.giCodModeloECF
    
    If AFRAC_ImpressoraCFe(iModeloImpressora) Then
        
        Select Case iModeloImpressora
            
            Case IMPRESSORA_SAT_2_5_15
                iModeloImpressora = Y.gobjSATInfo.iModeloImpressora
                
            Case IMPRESSORA_NFCE
                iModeloImpressora = Z.gobjNFeInfo.iModeloImpressora
                
            Case Else
                iModeloImpressora = IMPRESSORA_NAO_FISCAL
                
        End Select
    
    End If
    
    Select Case iModeloImpressora

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            AFRAC_FecharRelatorioGerencial2 = Bematech_FI_FechaRelatorioGerencial()

            If AFRAC_FecharRelatorioGerencial2 = 1 Then
                AFRAC_FecharRelatorioGerencial2 = SUCESSO
                Call CF_ECF("Grava_R06", "RG")
            ElseIf AFRAC_FecharRelatorioGerencial2 = 0 Then
                AFRAC_FecharRelatorioGerencial2 = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

'                AFRAC_FecharRelatorioGerencial2 = Daruma_FI_FechaRelatorioGerencial()
            AFRAC_FecharRelatorioGerencial2 = iRGFechar_ECF_Daruma()

            If AFRAC_FecharRelatorioGerencial2 = 1 Then
                AFRAC_FecharRelatorioGerencial2 = SUCESSO
                Call CF_ECF("Grava_R06", "RG")
            ElseIf AFRAC_FecharRelatorioGerencial2 = 0 Then
                AFRAC_FecharRelatorioGerencial2 = 1
            End If
            
        Case IMPRESSORA_NAO_FISCAL
            If Not (objTela Is Nothing) Then
                objTela.Controls("RT1").SelStart = 0
                objTela.Controls("RT1").SelLength = Len(objTela.Controls("RT1").Text)
                objTela.Controls("RT1").SelFontName = "Courier New"
'                If objTela.Controls("CD1").hdc > 0 Then objTela.Controls("RT1").SelPrint objTela.Controls("CD1").hdc
                objTela.Controls("RT1").SelFontSize = 10
                Call objTela.Controls("RT1").SelPrint(objTela.Controls("CD1").hdc)

            End If
            
            AFRAC_FecharRelatorioGerencial2 = SUCESSO
            
        Case IMPRESSORA_ELGIN

            iStatus = Elgin_EncerraCupomNaoFiscal()

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_FecharRelatorioGerencial2 = SUCESSO
            Else
                AFRAC_FecharRelatorioGerencial2 = 1
            End If
            
        Case IMPRESSORA_BEMATECH_MP4000_TH
            
            If Not (objRel Is Nothing) Then
            
                iStatus = objRel.TerminaImpressao
                If iStatus = 1 Then
                    AFRAC_FecharRelatorioGerencial2 = SUCESSO
                Else
                    AFRAC_FecharRelatorioGerencial2 = 1
                End If
                
                Set objRel = Nothing
            
            End If
            
            iStatus = BematechNaoFiscal_FechaPorta
            If iStatus = 1 Then
                AFRAC_FecharRelatorioGerencial2 = SUCESSO
            Else
                AFRAC_FecharRelatorioGerencial2 = 1
            End If
        
        Case IMPRESSORA_EPSON_TM_T20
            
            If Not (objRel Is Nothing) Then
            
                iStatus = objRel.TerminaImpressao
                If iStatus = 1 Then
                    AFRAC_FecharRelatorioGerencial2 = SUCESSO
                Else
                    AFRAC_FecharRelatorioGerencial2 = 1
                End If
                
                Set objRel = Nothing
            
            End If
            
            iStatus = EpsonNaoFiscal_FechaPorta
            If iStatus = 1 Then
                AFRAC_FecharRelatorioGerencial2 = SUCESSO
            Else
                AFRAC_FecharRelatorioGerencial2 = 1
            End If
        
        Case IMPRESSORA_DARUMA_DR700
            If objRel Is Nothing Then
            
                AFRAC_FecharRelatorioGerencial2 = SUCESSO
        
            Else
            
                iStatus = objRel.TerminaImpressao
                If iStatus = 1 Then
                    AFRAC_FecharRelatorioGerencial2 = SUCESSO
                Else
                    AFRAC_FecharRelatorioGerencial2 = 1
                End If
                
                Set objRel = Nothing
            
            End If
        
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS
            
            If Not (objRel Is Nothing) Then
            
                iStatus = objRel.TerminaImpressao
                If iStatus = 1 Then
                    AFRAC_FecharRelatorioGerencial2 = SUCESSO
                Else
                    AFRAC_FecharRelatorioGerencial2 = 1
                End If
                
                Set objRel = Nothing
            
            End If
            
        Case Else
            AFRAC_FecharRelatorioGerencial2 = SUCESSO
            
    End Select

   Exit Function
   
Erro_AFRAC_FecharRelatorioGerencial2:

    AFRAC_FecharRelatorioGerencial2 = 1
    
    Exit Function

End Function

'Grupo Autenticação

Public Function AFRAC_Autenticar(sLinha As String, Optional sMensagem As String) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String

'Autentica ( slinha no maximo 15 e em alguns ECF's 5)

On Error GoTo Erro_AFRAC_Autenticar

    AFRAC_Autenticar = 1
    
    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            AFRAC_Autenticar = Bematech_FI_Autenticacao
            
            If AFRAC_Autenticar = 1 Then
                AFRAC_Autenticar = SUCESSO
            ElseIf AFRAC_Autenticar = 0 Then
                AFRAC_Autenticar = 1
            End If
        
        Case IMPRESSORA_DARUMA_FS600
            
            'AFRAC_Autenticar = Daruma_FI_Autenticacao
            
            If AFRAC_Autenticar = 1 Then
                AFRAC_Autenticar = SUCESSO
            ElseIf AFRAC_Autenticar = 0 Then
                AFRAC_Autenticar = 1
            End If
        
        Case IMPRESSORA_ELGIN

            iStatus = Elgin_ImprimeValidacao("PAGO", "")

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_Autenticar = SUCESSO
            Else
                AFRAC_Autenticar = 1
            End If
            
        Case Else
            AFRAC_Autenticar = SUCESSO
            

    End Select

   Exit Function
   
Erro_AFRAC_Autenticar:

    AFRAC_Autenticar = 1
    
    Exit Function

End Function

Public Function AFRAC_RepetirAutenticacao() As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String

'ReAutentica

On Error GoTo Erro_AFRAC_RepetirAutenticacao
    
    AFRAC_RepetirAutenticacao = 1
    
    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            AFRAC_RepetirAutenticacao = Bematech_FI_Autenticacao
            
            If AFRAC_RepetirAutenticacao = 1 Then
                AFRAC_RepetirAutenticacao = SUCESSO
            ElseIf AFRAC_RepetirAutenticacao = 0 Then
                AFRAC_RepetirAutenticacao = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

'                AFRAC_RepetirAutenticacao = Daruma_FI_Autenticacao
            
            If AFRAC_RepetirAutenticacao = 1 Then
                AFRAC_RepetirAutenticacao = SUCESSO
            ElseIf AFRAC_RepetirAutenticacao = 0 Then
                AFRAC_RepetirAutenticacao = 1
            End If
            
        Case IMPRESSORA_ELGIN

            iStatus = Elgin_ImprimeValidacao("PAGO", "")

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_RepetirAutenticacao = SUCESSO
            Else
                AFRAC_RepetirAutenticacao = 1
            End If
            
        Case Else
            AFRAC_RepetirAutenticacao = SUCESSO
            

    End Select

   Exit Function
   
Erro_AFRAC_RepetirAutenticacao:

    AFRAC_RepetirAutenticacao = 1
    
    Exit Function
            
End Function

'gaveta de Dinheiro

Public Function AFRAC_AbrirGaveta() As Integer

Dim X As New ClassECFConfig, Y As New ClassSATGlobal
Dim iStatus As Integer, Z As New ClassNFeGlobal
Dim sRetorno As String, iModeloImpressora As Integer, iRetornoEpson As Integer, sPortaImpressora As String
Dim objRel As ClassImpressoraESC

'Abri a gaveta

On Error GoTo Erro_AFRAC_AbrirGaveta

    AFRAC_AbrirGaveta = 1

    iModeloImpressora = X.giCodModeloECF
    
    If AFRAC_ImpressoraCFe(iModeloImpressora) Then
        
        Select Case iModeloImpressora
            
            Case IMPRESSORA_SAT_2_5_15
                iModeloImpressora = Y.gobjSATInfo.iModeloImpressora
                sPortaImpressora = Y.gobjSATInfo.sPortaImpressora
                
            Case IMPRESSORA_NFCE
                iModeloImpressora = Z.gobjNFeInfo.iModeloImpressora
                sPortaImpressora = Z.gobjNFeInfo.sPortaImpressora
                
            Case Else
                iModeloImpressora = IMPRESSORA_NAO_FISCAL
                
        End Select
    
    End If

    Select Case iModeloImpressora

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            AFRAC_AbrirGaveta = Bematech_FI_AcionaGaveta
            
            If AFRAC_AbrirGaveta = 1 Then
                AFRAC_AbrirGaveta = SUCESSO
            ElseIf AFRAC_AbrirGaveta = 0 Then
                AFRAC_AbrirGaveta = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

'                AFRAC_AbrirGaveta = Daruma_FI_AcionaGaveta
            AFRAC_AbrirGaveta = eAbrirGaveta_ECF_Daruma()
            
            If AFRAC_AbrirGaveta = 1 Then
                AFRAC_AbrirGaveta = SUCESSO
            ElseIf AFRAC_AbrirGaveta = 0 Then
                AFRAC_AbrirGaveta = 1
            End If
            
        Case IMPRESSORA_EPSON_TM_T20
        
            If Len(Trim(sPortaImpressora)) <> 0 Then
            
                iRetornoEpson = EpsonNaoFiscal_ConfiguraTaxaSerial(115200)
                If iRetornoEpson = 1 Then
                
                    iRetornoEpson = EpsonNaoFiscal_IniciaPorta(sPortaImpressora)
                    
                    If iRetornoEpson = 1 Then
                    
                        Call EpsonNaoFiscal_AcionaGaveta 'ou EpsonNaoFiscal_ImprimeTextoTag("<g>")
                        Call EpsonNaoFiscal_FechaPorta
                
                    End If
                
                End If
            
            End If
            
            AFRAC_AbrirGaveta = SUCESSO
        
        Case IMPRESSORA_ELGIN

            iStatus = Elgin_AcionarGaveta()

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_AbrirGaveta = SUCESSO
            Else
                AFRAC_AbrirGaveta = 1
            End If
            
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS
            Set objRel = New ClassImpressoraESC
            Call objRel.Define_Impressora(sPortaImpressora, iModeloImpressora)
            Call objRel.AbrirGaveta
            iStatus = objRel.TerminaImpressao
            Set objRel = Nothing
            If iStatus = 1 Then
                AFRAC_AbrirGaveta = SUCESSO
            Else
                AFRAC_AbrirGaveta = 1
            End If
        
        Case Else
            AFRAC_AbrirGaveta = SUCESSO

    End Select

   Exit Function
   
Erro_AFRAC_AbrirGaveta:

    AFRAC_AbrirGaveta = 1
    
    Exit Function

End Function

Public Function AFRAC_VerificarGaveta(sStatus As String) As Integer

Dim X As New ClassECFConfig
Dim sBuffer As String
Dim iStatus As Integer
Dim sRetorno As String
Dim iStatus1 As Integer

'Verifica o Status da gaveta
 '1-Aberta e 0-Fechada

On Error GoTo Erro_AFRAC_VerificarGaveta

    AFRAC_VerificarGaveta = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            AFRAC_VerificarGaveta = Bematech_FI_VerificaEstadoGaveta(iStatus)
            
            If AFRAC_VerificarGaveta = 1 Then
                sStatus = iStatus
                AFRAC_VerificarGaveta = SUCESSO
            ElseIf AFRAC_VerificarGaveta = 0 Then
                AFRAC_VerificarGaveta = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

'                AFRAC_VerificarGaveta = Daruma_FI_VerificaEstadoGaveta(iStatus)
            AFRAC_VerificarGaveta = rStatusGaveta_ECF_Daruma(iStatus)
            
            If AFRAC_VerificarGaveta = 1 Then
                sStatus = iStatus
                AFRAC_VerificarGaveta = SUCESSO
            ElseIf AFRAC_VerificarGaveta = 0 Then
                AFRAC_VerificarGaveta = 1
            End If
            
        Case IMPRESSORA_ELGIN

            sBuffer = Space(41)

            iStatus = Elgin_TransStatus(0, sBuffer)

            sRetorno = Space(50)
            iStatus1 = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                sStatus = Mid(sBuffer, 23, 1)
                AFRAC_VerificarGaveta = SUCESSO
            Else
                AFRAC_VerificarGaveta = 1
            End If
        
        Case Else
            AFRAC_VerificarGaveta = SUCESSO

    End Select

   Exit Function
   
Erro_AFRAC_VerificarGaveta:

    AFRAC_VerificarGaveta = 1
    
    Exit Function

End Function

'Informações Impressora

Public Function AFRAC_LerInformacaoImpressora(sCodInformacao As String, sRetorno As String) As Integer

'Obtém várias Informações sobre o ECF
Dim X As New ClassECFConfig
Dim sCGC As String
Dim sIE As String
Dim sData As String
Dim sHora As String
Dim lTamanho As Long
Dim iStatus As Integer
Dim sBuffer As String
Dim iStatus1 As Integer
Dim sDataUsuario As String
Dim sDataSWBasico As String
Dim sMFAdicional As String

On Error GoTo Erro_AFRAC_LerInformacaoImpressora

    AFRAC_LerInformacaoImpressora = 1

    Select Case X.giCodModeloECF

        Case 1 To 3

            Select Case sCodInformacao

                Case "001"
                    sRetorno = String(14, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_SubTotal(sRetorno)

                Case "002"
                    sRetorno = String(15, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroSerie(sRetorno)

                Case "003"
                    sRetorno = String(4, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_VersaoFirmware(sRetorno)

                Case "004"
                    sCGC = String(18, 0)
                    sIE = String(15, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_CGC_IE(sCGC, sIE)
                    sRetorno = sCGC

                Case "005"
                    sCGC = String(18, 0)
                    sIE = String(15, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_CGC_IE(sCGC, sIE)
                    sRetorno = sIE

                Case "007"
                    sRetorno = String(18, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_GrandeTotal(sRetorno)

                Case "008"
                    sRetorno = String(14, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_Acrescimos(sRetorno)

                Case "009"
                    sRetorno = String(14, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_Descontos(sRetorno)

                Case "010"
                    sRetorno = String(14, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_Cancelamentos(sRetorno)

                Case "011"
                    sRetorno = String(4, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroIntervencoes(sRetorno)

                Case "012"
                    sRetorno = String(4, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroCuponsCancelados(sRetorno)

                Case "013"
                    sRetorno = String(4, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroReducoes(sRetorno)

                Case "014"
                    sRetorno = String(186, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_ClicheProprietario(sRetorno)

                Case "015"
                    sRetorno = String(4, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroCaixa(sRetorno)

                Case "016"
                    sRetorno = String(4, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroLoja(sRetorno)

                Case "017"
                    sData = String(6, 0)
                    sHora = String(6, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_DataHoraImpressora(sData, sHora)
                    sRetorno = sData & sHora

                Case "018"
                    sData = String(6, 0)
                    sHora = String(6, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_DataHoraReducao(sData, sHora)
                    sRetorno = sData & sHora

                Case "019"
                    sRetorno = String(6, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_DataMovimento(sRetorno)

                Case "020"
                    sRetorno = String(1, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_VerificaTruncamento(sRetorno)
                    If sRetorno = 1 Then
                        sRetorno = 0
                    Else
                        sRetorno = 1
                    End If

                Case "023"
                    sRetorno = String(6, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroCupom(sRetorno)

                Case "026"
                    sRetorno = String(4, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroIntervencoes(sRetorno)

                Case "050"
'                    'Retorna MF Adicional se houver
'                    sDataUsuario = Space(20)
'                    sDataSWBasico = Space(20)
'                    sMFAdicional = Space(5)
'
'                    If InStr(UCase(X.gsNumSerie), "EMULADOR") = 0 Then
'                        AFRAC_LerInformacaoImpressora = Bematech_FI_DataHoraGravacaoUsuarioSWBasicoMFAdicional(sDataUsuario, sDataSWBasico, sMFAdicional)
'                    Else
'                        AFRAC_LerInformacaoImpressora = 1
'                    End If
'
'                    sMFAdicional = StringZ(sMFAdicional)
'
'                    If Len(sMFAdicional) > 0 Then
'                        sRetorno = left(sMFAdicional, 1)
'                    Else
'                        sRetorno = " "
'                    End If

                    sRetorno = " "
                    AFRAC_LerInformacaoImpressora = 1

                
'                    If X.gsHoraSWBasico = "" Then
'
'                        'Retorna MF Adicional se houver
'                        sDataUsuario = Space(20)
'                        sDataSWBasico = Space(20)
'                        sMFAdicional = Space(5)
'
'                        If InStr(UCase(X.gsNumSerie), "EMULADOR") = 0 Then
'                            AFRAC_LerInformacaoImpressora = Bematech_FI_DataHoraGravacaoUsuarioSWBasicoMFAdicional(sDataUsuario, sDataSWBasico, sMFAdicional)
'                        Else
'                            sDataSWBasico = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
'                            AFRAC_LerInformacaoImpressora = 1
'                        End If
'
'
'                        X.gsMFAdicional = StringZ(sMFAdicional)
'
'                        If Len(sMFAdicional) > 0 Then
'                            sRetorno = left(X.gsMFAdicional, 1)
'                        Else
'                            sRetorno = " "
'                        End If
'
'
'                        sDataSWBasico = StringZ(sDataSWBasico)
'
'                        sData = left(sDataSWBasico, 10)
'
'                        X.gdtDataSWBasico = CDate(sData)
'
'                        X.gsHoraSWBasico = Mid(sDataSWBasico, 12, 2) & Mid(sDataSWBasico, 15, 2) & Mid(sDataSWBasico, 18, 2)
'
'                    Else
'
'                        sRetorno = X.gsMFAdicional
'
'                    End If


                Case "062"
                    sRetorno = String(14, 0)
'                    AFRAC_LerInformacaoImpressora = Bematech_FI_SubTotal(sRetorno)

            End Select

            If AFRAC_LerInformacaoImpressora = 1 Then
                AFRAC_LerInformacaoImpressora = SUCESSO
            ElseIf AFRAC_LerInformacaoImpressora = 0 Then
                AFRAC_LerInformacaoImpressora = 1
            End If


        Case IMPRESSORA_NAO_FISCAL

            Select Case sCodInformacao
            
                Case "023"
                    lTamanho = 6
                    sRetorno = String(lTamanho, 0)
                
                    Call GetPrivateProfileString(APLICACAO_CAIXA, "COO", CONSTANTE_ERRO, sRetorno, lTamanho + 1, NOME_ARQUIVO_CAIXA)
    
                    If sRetorno <> String(lTamanho, 0) Then
                        AFRAC_LerInformacaoImpressora = SUCESSO
                    Else
                        AFRAC_LerInformacaoImpressora = 1
                    End If
                    
                Case Else
                    AFRAC_LerInformacaoImpressora = SUCESSO

            End Select
            
        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            Select Case sCodInformacao

                Case "001"
                    sRetorno = String(15, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_SubTotal(sRetorno)

                Case "002"
                    sRetorno = Space(21)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroSerieMFD(sRetorno)

                Case "003"
                    sRetorno = Space(7)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_VersaoFirmwareMFD(sRetorno)


                Case "004"
                    sCGC = String(19, 0)
                    sIE = String(16, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_CGC_IE(sCGC, sIE)
                    sRetorno = sCGC

                Case "005"
                    sCGC = String(19, 0)
                    sIE = String(16, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_CGC_IE(sCGC, sIE)
                    sRetorno = sIE

                Case "007"
                    sRetorno = String(19, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_GrandeTotal(sRetorno)

                Case "008"
                    sRetorno = String(15, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_Acrescimos(sRetorno)

                Case "009"
                    sRetorno = String(15, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_Descontos(sRetorno)

                Case "010"
                    sRetorno = String(15, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_Cancelamentos(sRetorno)

                Case "011"
                    sRetorno = String(5, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroIntervencoes(sRetorno)

                Case "012"
                    sRetorno = String(5, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroCuponsCancelados(sRetorno)

                Case "013"
                    sRetorno = String(5, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroReducoes(sRetorno)

                Case "014"
                    sRetorno = String(187, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_ClicheProprietario(sRetorno)

                Case "015"
                    sRetorno = String(5, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroCaixa(sRetorno)

                Case "016"
                    sRetorno = String(5, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroLoja(sRetorno)

                Case "017"
                    sData = String(7, 0)
                    sHora = String(7, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_DataHoraImpressora(sData, sHora)
                    sData = StringZ(sData)
                    sHora = StringZ(sHora)
                    sRetorno = sData & sHora

                Case "018"
                    sData = String(7, 0)
                    sHora = String(7, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_DataHoraReducao(sData, sHora)
                    sRetorno = sData & sHora

                Case "019"
                    sRetorno = String(13, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_DataMovimento(sRetorno)
    
                Case "020"
                    sRetorno = String(1, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_VerificaTruncamento(sRetorno)
                    If sRetorno = 1 Then
                        sRetorno = 1
                    Else
                        sRetorno = 0
                    End If


                Case "023"
                    sRetorno = String(7, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroCupom(sRetorno)

                Case "026"
                    sRetorno = String(5, 0)
                    AFRAC_LerInformacaoImpressora = Bematech_FI_NumeroIntervencoes(sRetorno)

                Case "050"
'                    'Retorna MF Adicional se houver
'                    sDataUsuario = Space(21)
'                    sDataSWBasico = Space(21)
'                    sMFAdicional = Space(2)
'
'                    If InStr(UCase(X.gsNumSerie), "EMULADOR") = 0 Then
'                        AFRAC_LerInformacaoImpressora = Bematech_FI_DataHoraGravacaoUsuarioSWBasicoMFAdicional(sDataUsuario, sDataSWBasico, sMFAdicional)
'                    Else
'                        AFRAC_LerInformacaoImpressora = 1
'                    End If
'
'                    sMFAdicional = StringZ(sMFAdicional)
'
'                    If Len(sMFAdicional) > 0 Then
'                        sRetorno = left(sMFAdicional, 1)
'                    Else
'                        sRetorno = " "
'                    End If
'
'                    AFRAC_LerInformacaoImpressora = 1
                    
                    sRetorno = " "
                    AFRAC_LerInformacaoImpressora = 1
                    
                    
'                    If X.gsHoraSWBasico = "" Then
'
'                        'Retorna MF Adicional se houver
'                        sDataUsuario = Space(21)
'                        sDataSWBasico = Space(21)
'                        sMFAdicional = Space(2)
'
'                        If InStr(UCase(X.gsNumSerie), "EMULADOR") = 0 Then
'                            AFRAC_LerInformacaoImpressora = Bematech_FI_DataHoraGravacaoUsuarioSWBasicoMFAdicional(sDataUsuario, sDataSWBasico, sMFAdicional)
'                        Else
'                            sDataSWBasico = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
'                            AFRAC_LerInformacaoImpressora = 1
'                        End If
'
'                        X.gsMFAdicional = StringZ(sMFAdicional)
'
'                        If Len(sMFAdicional) > 0 Then
'                            sRetorno = left(X.gsMFAdicional, 1)
'                        Else
'                            sRetorno = " "
'                        End If
'
'
'                        sDataSWBasico = StringZ(sDataSWBasico)
'
'                        sData = left(sDataSWBasico, 10)
'
'                        X.gdtDataSWBasico = CDate(sData)
'
'                        X.gsHoraSWBasico = Mid(sDataSWBasico, 12, 2) & Mid(sDataSWBasico, 15, 2) & Mid(sDataSWBasico, 18, 2)
'
'                    Else
'
'                        sRetorno = X.gsMFAdicional
'
'                    End If


                Case "062"
                    sRetorno = String(14, 0)
'                    AFRAC_LerInformacaoImpressora = Bematech_FI_SubTotal(sRetorno)


            End Select

            sRetorno = StringZ(sRetorno)

            If AFRAC_LerInformacaoImpressora = 1 Then
                AFRAC_LerInformacaoImpressora = SUCESSO
            ElseIf AFRAC_LerInformacaoImpressora = 0 Then
                AFRAC_LerInformacaoImpressora = 1
            End If
            
            
        Case IMPRESSORA_DARUMA_FS600

            Select Case sCodInformacao

                Case "001"
                    'retorna o subtotal do cupom fiscal atual
                    sRetorno = String(14, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_SubTotal(sRetorno)
                    AFRAC_LerInformacaoImpressora = rCFSubTotal_ECF_Daruma(sRetorno)

                Case "002"
                    'retorna o numero de fabricacao (numero de serie)
                    sRetorno = String(21, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FIMFD_RetornaInformacao("78", sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("78", sRetorno)
                    sRetorno = left(sRetorno, 20)
                Case "003"
                    'Versao do SB
                    sRetorno = String(6, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FIMFD_RetornaInformacao("83", sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("83", sRetorno)
                                
                Case "004"
                    sCGC = String(20, 0)
 '                   sIE = String(15, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_CGC_IE(sCGC, sIE)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("90", sCGC)
                    sRetorno = Trim(sCGC)

                Case "005"
'                    sCGC = String(18, 0)
                    sIE = String(20, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_CGC_IE(sCGC, sIE)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("91", sIE)
                    sRetorno = Trim(sIE)

                Case "007"
                    sRetorno = String(18, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_GrandeTotal(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("1", sRetorno)

                Case "008"
                    sRetorno = String(13, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_Acrescimos(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("12", sRetorno)

                Case "009"
                    sRetorno = String(13, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_Descontos(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("11", sRetorno)

                Case "010"
                    sRetorno = String(13, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_Cancelamentos(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("13", sRetorno)

                Case "011"
                    sRetorno = String(3, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_NumeroIntervencoes(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("23", sRetorno)

                Case "012"
                    'retorna o CFC (contador de cupom fiscal cancelado)
                    sRetorno = String(4, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_NumeroCuponsCancelados(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("39", sRetorno)

                Case "013"
                    sRetorno = String(4, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_NumeroReducoes(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("24", sRetorno)

                Case "014"
                    sRetorno = String(219, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_ClicheProprietario(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("132", sRetorno)

                Case "015"
                    sRetorno = String(3, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_NumeroCaixa(sRetorno)
                    'numero do ECF
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("107", sRetorno)

                Case "016"
                    sRetorno = String(4, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_NumeroLoja(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("129", sRetorno)

                Case "017"
                    sData = String(8, 0)
                    sHora = String(6, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_DataHoraImpressora(sData, sHora)
                    AFRAC_LerInformacaoImpressora = rDataHoraImpressora_ECF_Daruma(sData, sHora)
                    sData = Mid(sData, 1, 4) & Mid(sData, 7, 2)
                    sRetorno = sData & sHora

                Case "018"
'                    sData = String(6, 0)
'                    sHora = String(6, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_DataHoraReducao(sData, sHora)
                    sRetorno = String(14, 0)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("154", sRetorno)
                    sData = Mid(sRetorno, 1, 4) & Mid(sRetorno, 7, 2)
                    sHora = right(sRetorno, 6)
                    sRetorno = sData & sHora

                Case "019"
                    sRetorno = String(8, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_DataMovimento(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("70", sRetorno)
                    sRetorno = Mid(sRetorno, 1, 4) & Mid(sRetorno, 7, 2)

                Case "020"
                    sRetorno = String(1, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_VerificaTruncamento(sRetorno)
                    AFRAC_LerInformacaoImpressora = regRetornaValorChave_DarumaFramework("ECF", "ArredondarTruncar", sRetorno)
                    If sRetorno = "A" Then
                        sRetorno = 0
                    Else
                        sRetorno = 1
                    End If

                Case "023"
                    sRetorno = String(6, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_NumeroCupom(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("26", sRetorno)
                    
                Case "026"
                    sRetorno = String(3, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FI_NumeroIntervencoes(sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("23", sRetorno)


                Case "050"
                    'Retorna MF Adicional se houver
                    sRetorno = String(21, 0)
                    'AFRAC_LerInformacaoImpressora = Daruma_FIMFD_RetornaInformacao("78", sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("78", sRetorno)
                    sRetorno = right(Trim(sRetorno), 1)
                    
                Case "051"
                    'Retorna MF Adicional se houver
                    sRetorno = String(6, 0)
                    'AFRAC_LerInformacaoImpressora = Daruma_FIMFD_RetornaInformacao("78", sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("50", sRetorno)
                    
                Case "062"
                    'Retorna valor liquido do ultimo item inserido
                    sRetorno = String(11, 0)
'                    AFRAC_LerInformacaoImpressora = Daruma_FIMFD_RetornaInformacao("62", sRetorno)
                    AFRAC_LerInformacaoImpressora = rRetornarInformacao_ECF_Daruma("62", sRetorno)
                    


            End Select

            If AFRAC_LerInformacaoImpressora = 1 Then
                AFRAC_LerInformacaoImpressora = SUCESSO
            ElseIf AFRAC_LerInformacaoImpressora = 0 Then
                AFRAC_LerInformacaoImpressora = 1
            End If


        Case IMPRESSORA_ELGIN

            Select Case sCodInformacao

                Case "001"
                
                    iStatus = Elgin_EcfPar("96")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 15)
                    End If
                

                Case "002"

                    iStatus = Elgin_EcfPar("49")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 10)
                    End If


                Case "003"

                    iStatus = Elgin_EcfPar("47")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                        
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 7)
                    End If


                Case "004"

                    iStatus = Elgin_EcfPar("88")
        
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 18)
                    End If



                Case "005"
                
                    iStatus = Elgin_EcfPar("88")
        
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 24, 15)
                    End If
                
                
                Case "007"

                    iStatus = Elgin_EcfPar("39")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 19)
                    End If


                Case "008"
                    iStatus = Elgin_EcfPar("37")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 15)
                    End If


                Case "009"
                    iStatus = Elgin_EcfPar("36")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 15)
                    End If

                Case "010"
                    iStatus = Elgin_EcfPar("38")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 15)
                    End If

                Case "011"

                    iStatus = Elgin_EcfPar("43")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 4)
                    End If

                Case "012"
                    iStatus = Elgin_EcfPar("44")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 4)
                    End If

                Case "013"
                    iStatus = Elgin_EcfPar("42")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 4)
                    End If

                Case "014"

                Case "015"
                    iStatus = Elgin_EcfPar("48")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 6)
                    End If

                Case "016"
                    iStatus = Elgin_EcfPar("46")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 2)
                    End If

                Case "017"


                    iStatus = Elgin_TransDataHora()
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 6) & Mid(sRetorno, 13, 6)
                    End If



                Case "018"


                    iStatus = Elgin_EcfPar("83")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 6)
                    End If


                Case "019"

                Case "020"

                    
                    sBuffer = Space(41)
                    
                    iStatus = Elgin_TransStatus(0, sBuffer)
        
                    sRetorno = Space(50)
                    iStatus1 = Elgin_EsperaResposta(sRetorno)
        
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sBuffer, 38, 1)
                        
                        If sRetorno = "1" Then
                            sRetorno = "0"
                        Else
                            sRetorno = "1"
                        End If
                    End If
                    

                Case "023"
                    iStatus = Elgin_EcfPar("41")
    
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 6)
                    End If

                Case "026"
                    iStatus = Elgin_EcfPar("43")
        
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                    
                    If iStatus = SUCESSO Then
                        sRetorno = Mid(sRetorno, 6, 4)
                    End If


            End Select

            If iStatus = SUCESSO Then
                AFRAC_LerInformacaoImpressora = SUCESSO
            Else
                AFRAC_LerInformacaoImpressora = 1
            End If


        Case Else
                AFRAC_LerInformacaoImpressora = SUCESSO

    End Select

   Exit Function
   
Erro_AFRAC_LerInformacaoImpressora:

    AFRAC_LerInformacaoImpressora = 1
    
    Exit Function


End Function

Public Function AFRAC_LerAliquotas(sRetorno As String) As Integer

Dim X As New ClassECFConfig
Dim sResp As String
Dim iStatus As Integer
Dim iIndice As Integer

'Retorno a Ser Preenchido com as Informações das Aliquotas Programadas no ECF

On Error GoTo Erro_AFRAC_LerAliquotas

    AFRAC_LerAliquotas = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            sRetorno = Space(150)

            AFRAC_LerAliquotas = Bematech_FI_RetornoAliquotas(sRetorno)
            
            sRetorno = StringZ(sRetorno)
            
            If AFRAC_LerAliquotas = 1 Then
                AFRAC_LerAliquotas = SUCESSO
            ElseIf AFRAC_LerAliquotas = 0 Then
                AFRAC_LerAliquotas = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

            sRetorno = Space(150)

'                AFRAC_LerAliquotas = Daruma_FI_RetornoAliquotas(sRetorno)
            AFRAC_LerAliquotas = rLerAliquotas_ECF_Daruma(sRetorno)
            
            If AFRAC_LerAliquotas = 1 Then
                AFRAC_LerAliquotas = SUCESSO
            ElseIf AFRAC_LerAliquotas = 0 Then
                AFRAC_LerAliquotas = 1
            End If
            
        Case IMPRESSORA_ELGIN
                iStatus = Elgin_TransTabAliquotas()
                
                sResp = Space(100)
                iStatus = Elgin_EsperaResposta(sResp)
                
                If iStatus = SUCESSO Then
                    sResp = Mid(sResp, 6)
                    sRetorno = ""
                    For iIndice = 0 To 15
                        sRetorno = sRetorno & Mid(sResp, (iIndice * 4) + 1, 4)
                        If iIndice < 15 Then
                            sRetorno = sRetorno & ","
                        End If
                    Next
                End If
                
                If iStatus = SUCESSO Then
                    AFRAC_LerAliquotas = SUCESSO
                Else
                    AFRAC_LerAliquotas = 1
                End If
            
            
        Case Else
            AFRAC_LerAliquotas = SUCESSO

    End Select

   Exit Function
   
Erro_AFRAC_LerAliquotas:

    AFRAC_LerAliquotas = 1
    
    Exit Function

End Function

Public Function AFRAC_LerValorTotalAliquotas(sRetorno As String) As Integer


'Retorno a Ser Preenchido com as Informações das Aliquotas Programadas no ECF

    AFRAC_LerValorTotalAliquotas = SUCESSO


End Function

Public Function AFRAC_LerTotalizadoresNSICMS(sRetorno As String) As Integer


'Retorno a Ser Preenchido com as Informações dos Totalizadores Programadas no ECF

    AFRAC_LerTotalizadoresNSICMS = SUCESSO


End Function


'Fabricante

Public Function AFRAC_Fabricante(sEntrada As String, sSaida As String) As Integer


'Comandos para a Impressora que Foram Implementadas pelo Fabricante.

    AFRAC_Fabricante = SUCESSO


End Function

'Grupo impressao de cheques
Public Function AFRAC_ChequeProgramarMoeda(sDescMoedaSingular As String, sDescMoedaPlural As String) As Integer

Dim X As New ClassECFConfig

'Configura a moeda para impressão de cheque.

On Error GoTo Erro_AFRAC_ChequeProgramarMoeda

    AFRAC_ChequeProgramarMoeda = 1
    
    Select Case X.giCodModeloECF

        Case 1 To 3

            AFRAC_ChequeProgramarMoeda = Bematech_FI_ProgramaMoedaPlural(sDescMoedaPlural)
            AFRAC_ChequeProgramarMoeda = Bematech_FI_ProgramaMoedaSingular(sDescMoedaSingular)

            If AFRAC_ChequeProgramarMoeda = 1 Then
                AFRAC_ChequeProgramarMoeda = SUCESSO
            ElseIf AFRAC_ChequeProgramarMoeda = 0 Then
                AFRAC_ChequeProgramarMoeda = 1
            End If

        Case Else
            AFRAC_ChequeProgramarMoeda = SUCESSO

    End Select

   Exit Function
   
Erro_AFRAC_ChequeProgramarMoeda:

    AFRAC_ChequeProgramarMoeda = 1
    
    Exit Function

End Function

Public Function AFRAC_ChequeImprimir(sNumeroBanco As String, dValor As Double, sFavorecido As String, sCidade As String, dtData As Date, sMensagem As String) As Integer

Dim X As New ClassECFConfig

'Imprimir Cheque no ECF

On Error GoTo Erro_AFRAC_ChequeImprimir

    AFRAC_ChequeImprimir = 1

    Select Case X.giCodModeloECF

        Case 1 To 3

            AFRAC_ChequeImprimir = Bematech_FI_ImprimeCheque(sNumeroBanco, CStr(Format(dValor, "standard")), sFavorecido, sCidade, Format(dtData, "dd/mm/yy"), sMensagem)
            
            If AFRAC_ChequeImprimir = 1 Then
                AFRAC_ChequeImprimir = SUCESSO
            ElseIf AFRAC_ChequeImprimir = 0 Then
                AFRAC_ChequeImprimir = 1
            End If

        Case IMPRESSORA_DARUMA_FS600

'                AFRAC_ChequeImprimir = Daruma_FI2000_ImprimirCheque(sNumeroBanco, sCidade, Format(dtData, "dd/mm/yy"), sFavorecido, CStr(Format(dValor, "standard")), "H")
            
            If AFRAC_ChequeImprimir = 1 Then
                AFRAC_ChequeImprimir = SUCESSO
            ElseIf AFRAC_ChequeImprimir = 0 Then
                AFRAC_ChequeImprimir = 1
            End If

        Case Else
            AFRAC_ChequeImprimir = SUCESSO

    End Select

    Exit Function
   
Erro_AFRAC_ChequeImprimir:

    AFRAC_ChequeImprimir = 1
    
    Exit Function

End Function

Public Function AFRAC_ChequeConfiguraFormatacao(sFormatoValor As String, sFormatoExtenso As Double, sFormatoFavorecido As String, sFormatoCidade As String, sFormatoData As String, sFormatoBomPara As String, bCruzarCheque As Boolean) As Integer


'Texto de Formatação para Impressora 4 Caracteres

    AFRAC_ChequeConfiguraFormatacao = SUCESSO


End Function

Public Function AFRAC_ChequeImprimirVerso(sTextoPrimeiraLinha As String, sTextoSegundaLinha As String) As Integer


'Imprime MSG no Verso do cheque

    AFRAC_ChequeImprimirVerso = SUCESSO


End Function

Public Function AFRAC_ChequeLerCMC7(sCMC7Lido As String) As Integer


'Le o CMC7 de Um Cheque

    AFRAC_ChequeLerCMC7 = SUCESSO


End Function

Public Function AFRAC_ChequeExpulsar() As Integer


'Expulsa o Cheque que se encontra no ECF

    AFRAC_ChequeExpulsar = SUCESSO


End Function

'Grupo Outras

Public Function AFRAC_PegarCodigoErro(sCodigoErro As String, sMsg As String, iAcao As Integer) As Long

Dim X As New ClassECFConfig
Dim iErro As Integer
Dim sResp As String
Dim iStatus As Integer
Dim Str_Msg_NumErro As String
Dim Str_Msg_NumAviso As String
Dim sMsg1 As String
Dim vbMsgRes As VbMsgBoxResult
Dim Str_Msg_NumRetorno As String
Dim sMsg2 As String
Dim lIntervalo As Long

'Retorna a mensagem de erro e o código correnspondente. A ação são algumas ações sugeridas para solucionao o problema.

On Error GoTo Erro_AFRAC_PegarCodigoErro

    AFRAC_PegarCodigoErro = 1
    
    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            iErro = CInt(sCodigoErro)
            AFRAC_PegarCodigoErro = VerificaRetornoImpressora("", "", iErro, sMsg)
            sCodigoErro = iErro
            
        Case IMPRESSORA_DARUMA_FS600
        
            Str_Msg_NumErro = Space(200)
            Str_Msg_NumAviso = Space(200)
            Str_Msg_NumRetorno = Space(200)
            
            iErro = eInterpretarRetorno_ECF_Daruma(sCodigoErro, Str_Msg_NumRetorno)
       
            iErro = eRetornarAvisoErroUltimoCMD_ECF_Daruma(Str_Msg_NumAviso, Str_Msg_NumErro)
                      
            sMsg1 = Trim(Str_Msg_NumRetorno)
            
            sMsg2 = IIf(Trim(Str_Msg_NumAviso) <> "Sem Aviso", " " & Trim(Str_Msg_NumAviso), "") & IIf(Trim(Str_Msg_NumErro) <> "Sem Erro", " " & Trim(Str_Msg_NumErro), "")
                      
                      
            If InStr(sMsg2, "Papel acabando") And InStr(Str_Msg_NumErro, "Fim do papel") = 0 Then
                lIntervalo = DateDiff("n", X.gdtDataHoraFimPapel, Now)
                If lIntervalo <= X.giIntervaloPapelAcabando Then
                    iAcao = -1
                    sMsg1 = ""
                    sMsg2 = ""
                End If
            End If
                      
                      
            If sMsg1 <> "" Then
                'desbloqueia mouse e teclado
                Call AFRAC_DesbloqueiaTecladoMouse
                
                vbMsgRes = Rotina_AvisoECF(vbOKOnly, sMsg & sMsg1)
                
            End If
                      
            If sMsg2 <> "" Then
                'desbloqueia mouse e teclado
                Call AFRAC_DesbloqueiaTecladoMouse
                
                vbMsgRes = Rotina_AvisoECF(vbOKOnly, sMsg2)
                
                'iAcao = -1 ==> nao faz nada, só avisa
                If InStr(sMsg2, "Papel acabando") And InStr(Str_Msg_NumErro, "Fim do papel") = 0 Then
                    iAcao = -1
                    X.gdtDataHoraFimPapel = Now
                End If
                
            End If
                      
                      
            If iErro = 1 Then
                AFRAC_PegarCodigoErro = SUCESSO
            Else
                AFRAC_PegarCodigoErro = 1
            End If
                      
                                
        Case IMPRESSORA_ELGIN
            
'            sResp = Space(2048)
'            iStatus = Elgin_EsperaResposta(sResp)
'
'            sCodigoErro = iStatus
'
'            Select Case iStatus
'
'                Case CIF_TIMEOUT
'                    sMsg = "TIMEOUT"
'
'                Case CIF_ERR_ANSWER
'                    sMsg = sResp
'
'                Case CIF_ERR_READSER
'                    sMsg = "Erro Leitura Serial"
'
'                Case CIF_OVERFLOW
'                    sMsg = sResp
'
'                Case CIF_ERR_TABERTA, CIF_IRRECUPERAVEL, CIF_ERR_MECANICO, CIF_ERR_TEMP, CIF_ERR_PPAPEL, CIF_ERR_SYS
'
'                    sMsg = sResp
'
'                Case Else
'                    sMsg = sResp
'
'            End Select
            

        Case Else
            AFRAC_PegarCodigoErro = SUCESSO

    End Select

    Exit Function
   
Erro_AFRAC_PegarCodigoErro:

    AFRAC_PegarCodigoErro = 1
    
    Exit Function

End Function

Public Function AFRAC_SetaPadrao(sVersao As String) As Integer


'Determina Uma Versão em que a Biblioteca deverá Trabalhar.

    AFRAC_SetaPadrao = SUCESSO


End Function

'Funções TEF

Public Function AFRAC_VendaCartao(lNumProxIdentificacao As Long, sValorTotal As String) As Integer
'chama função do cartão

Dim X As New ClassECFConfig

On Error GoTo Erro_AFRAC_VendaCartao

    AFRAC_VendaCartao = 1

    Select Case X.giCodModeloECF

        Case 1 To 3

'            AFRAC_VendaCartao = Bematech_FITEF_VendaCartao(CStr(lNumProxIdentificacao), sValorTotal)
            
            If AFRAC_VendaCartao = 1 Then
                AFRAC_VendaCartao = SUCESSO
            ElseIf AFRAC_VendaCartao = 0 Then
                AFRAC_VendaCartao = 1
            End If

        Case Else
            AFRAC_VendaCartao = SUCESSO

    End Select

    Exit Function
   
Erro_AFRAC_VendaCartao:

    AFRAC_VendaCartao = 1
    
    Exit Function

End Function

Public Function AFRAC_ImprimeTEF(lNumProxIdentificacao As Long, sDescricao As String, dPagTEF As Double) As Integer
'chama função de impressão de TEF

Dim X As New ClassECFConfig


On Error GoTo Erro_AFRAC_ImprimeTEF

    AFRAC_ImprimeTEF = 1

    Select Case X.giCodModeloECF

        Case 1 To 3

'            AFRAC_ImprimeTEF = Bematech_FITEF_ImprimeTEF(CStr(lNumProxIdentificacao), sDescricao, CStr(Format(dPagTEF, "standard")))
            If AFRAC_ImprimeTEF = 1 Then
'                AFRAC_ImprimeTEF = Bematech_FITEF_ConfirmaVenda(CStr(lNumProxIdentificacao), CStr(Format(dPagTEF, "standard")), "CNF")
            ElseIf AFRAC_ImprimeTEF = 0 Then
'                AFRAC_ImprimeTEF = Bematech_FITEF_NaoConfirmaVendaImpressao(CStr(lNumProxIdentificacao), CStr(Format(dPagTEF, "standard")))
                If AFRAC_ImprimeTEF = 1 Then
                    AFRAC_ImprimeTEF = SUCESSO
                ElseIf AFRAC_ImprimeTEF = 0 Then
                    AFRAC_ImprimeTEF = 1
                End If
            End If

        Case Else
            AFRAC_ImprimeTEF = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_ImprimeTEF:

    AFRAC_ImprimeTEF = 1
    
    Exit Function

End Function

Public Function AFRAC_NaoConfirmaVendaImpressao(lNumProxIdentificacao As Long, dPagTEF As Double) As Integer
'chama função de não confirmação de venda ou impressão

Dim X As New ClassECFConfig

On Error GoTo Erro_AFRAC_NaoConfirmaVendaImpressao

    AFRAC_NaoConfirmaVendaImpressao = 1
     
    Select Case X.giCodModeloECF

        Case 1 To 3

'            AFRAC_NaoConfirmaVendaImpressao = Bematech_FITEF_NaoConfirmaVendaImpressao(CStr(lNumProxIdentificacao), CStr(Format(dPagTEF, "standard")))
            If AFRAC_NaoConfirmaVendaImpressao = 1 Then
                AFRAC_NaoConfirmaVendaImpressao = SUCESSO
            ElseIf AFRAC_NaoConfirmaVendaImpressao = 0 Then
                AFRAC_NaoConfirmaVendaImpressao = 1
            End If

        Case Else
            AFRAC_NaoConfirmaVendaImpressao = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_NaoConfirmaVendaImpressao:

    AFRAC_NaoConfirmaVendaImpressao = 1
    
    Exit Function

End Function

Public Function VerificaRetornoImpressora(Label As String, RetornoFuncao As String, iRetorno As Integer, TituloJanela As String) As Long

Dim ACK As Integer
Dim ST1 As Integer
Dim ST2 As Integer
Dim ST3 As Integer
Dim RetornaMensagem As Integer
Dim StringRetorno As String
Dim ValorRetorno As String
Dim iRetornoStatus As Integer
Dim Mensagem As String
Dim iSt1 As Integer
Dim iSt2 As Integer
Dim vbMsg As VbMsgBoxResult

Dim X As New ClassECFConfig

On Error GoTo Erro_VerificaRetornoImpressora


    If iRetorno = 1 Then
    
        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse
    
        MsgBox "Erro de comunicação com a impressora.", vbOKOnly + vbCritical, TituloJanela
        
        gError 117665

    ElseIf iRetorno = -27 Then
    
        If X.giCodModeloECF = IMPRESSORA_DARUMA_FS600 Then
            iRetornoStatus = Daruma_FI_RetornoImpressora(ACK, ST1, ST2)
        ElseIf X.giCodModeloECF = 1 Or X.giCodModeloECF = 2 Or X.giCodModeloECF = 3 Then
            iRetornoStatus = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
        ElseIf X.giCodModeloECF = 5 Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP4200_TH_FI Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP2100_TH_FI Then
            iRetornoStatus = Bematech_FI_RetornoImpressoraMFD(ACK, ST1, ST2, ST3)
        End If
        
        If iRetornoStatus = 1 Then
        
        
            ValorRetorno = str(ACK) & "," & str(ST1) & "," & str(ST2)
            iSt1 = ST1
            iSt2 = ST2
            
            If Label <> "" And RetornoFuncao <> "" Then
                RetornaMensagem = 1
            End If
    
            If ACK = 21 Then
                MsgBox "Status da Impressora: 21" & vbCr & vbLf & "Comando não executado", vbOKOnly + vbInformation, TituloJanela
                gError 117666
            End If
    
            If ST1 And 64 Then ST1 = ST1 Xor 64
    
            If (ST1 <> 0 Or ST2 <> 0) Then
            
                If ST1 <> 2 Or TituloJanela <> "Abre Cupom" Then
            
                    If (ST1 >= 128) Then
                        StringRetorno = "Fim de Papel" & vbCr
                        ST1 = ST1 - 128
                    End If
    
                    If (ST1 >= 64) Then
                        StringRetorno = StringRetorno & "Pouco Papel" & vbCr
                        ST1 = ST1 - 64
                    End If
    
                    If (ST1 >= 32) Then
                        StringRetorno = StringRetorno & "Erro no relógio" & vbCr
                        ST1 = ST1 - 32
                    End If
    
                    If (ST1 >= 16) Then
                        StringRetorno = StringRetorno & "Impressora em erro" & vbCr
                        ST1 = ST1 - 16
                    End If
    
                    If (ST1 >= 8) Then
                        StringRetorno = StringRetorno & "Primeiro dado do comando não foi Esc" & vbCr
                        ST1 = ST1 - 8
                    End If
    
                    If (ST1 >= 4) Then
                        StringRetorno = StringRetorno & "Comando inexistente" & vbCr
                        ST1 = ST1 - 4
                    End If
    
                    If (ST1 >= 2) Then
                        StringRetorno = StringRetorno & "Cupom fiscal aberto" & vbCr
                        ST1 = ST1 - 2
                        Call Bematech_FI_CancelaCupom
                    End If
    
                    If (ST1 >= 1) Then
                        StringRetorno = StringRetorno & "Número de parâmetros inválido no comando" & vbCr
                        ST1 = ST1 - 1
                    End If
    
                    If (ST2 >= 128) Then
                        StringRetorno = "Tipo de Parâmetro de comando inválido" & vbCr
                        ST2 = ST2 - 128
                    End If
    
                    If (ST2 >= 64) Then
                        StringRetorno = StringRetorno & "Memória fiscal lotada" & vbCr
                        ST2 = ST2 - 64
                    End If
    
                    If (ST2 >= 32) Then
                        StringRetorno = StringRetorno & "Erro na CMOS" & vbCr
                        ST2 = ST2 - 32
                    End If
    
                    If (ST2 >= 16) Then
                        StringRetorno = StringRetorno & "Alíquota não programada" & vbCr
                        ST2 = ST2 - 16
                    End If
    
                    If (ST2 >= 8) Then
                        StringRetorno = StringRetorno & "Capacidade de alíquota programáveis lotada" & vbCr
                        ST2 = ST2 - 8
                    End If
    
                    If (ST2 >= 4) Then
                        StringRetorno = StringRetorno & "Cancelamento não permitido" & vbCr
                        ST2 = ST2 - 4
                    End If
    
                    If (ST2 >= 2) Then
                        StringRetorno = StringRetorno & "CGC/IE do proprietário não programados" & vbCr
                        ST2 = ST2 - 2
                    End If
    
                    If (ST2 >= 1) Then
                        StringRetorno = StringRetorno & "Comando não executado" & vbCr
                        ST2 = ST2 - 1
                    End If
    
                    If RetornaMensagem Then
                        Mensagem = "Status da Impressora: " & ValorRetorno & _
                               vbCr & vbLf & StringRetorno & vbCr & vbLf & _
                               Label & RetornoFuncao
                    Else
                        Mensagem = "Status da Impressora: " & ValorRetorno & _
                           vbCr & vbLf & StringRetorno
                    End If
    '
    '                MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
    '                If iSt1 <> 64 Or ST2 <> 0 Then gError 117667
            
                End If
            
            End If 'fim do ST1 <> 0 and ST2 <> 0
    
            If RetornaMensagem Then
                Mensagem = Label & RetornoFuncao
            End If
    
            If X.giCodModeloECF = 5 Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP4200_TH_FI Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP2100_TH_FI Then
            
                Mensagem = Mensagem & "ST3 = " & ST3
            
            End If
    
            If Mensagem <> "" Then
                'desbloqueia mouse e teclado
                Call AFRAC_DesbloqueiaTecladoMouse
                
                MsgBox Mensagem, vbOKOnly + vbInformation, TituloJanela
                If iSt1 <> 64 Or iSt2 <> 0 Then gError 117668
            End If
            
        Else
            
        End If
    
    
    ElseIf iRetorno = -1 Then
        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Erro de execução da função.", vbOKOnly + vbCritical, TituloJanela
        gError 117669

    ElseIf iRetorno = -2 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Parâmetro inválido na função.", vbOKOnly + vbExclamation, TituloJanela
        gError 126800

    ElseIf iRetorno = -3 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Alíquota não programada.", vbOKOnly + vbExclamation, TituloJanela
        gError 117670

    ElseIf iRetorno = -4 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse


        MsgBox "O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório default. " + vbCr + "Por favor, copie esse arquivo para o diretório de sistema do Windows." + vbCr + "Se for o Windows 95 ou 98 é o diretório 'System' se for o Windows NT é o diretório 'System32'.", vbOKOnly + vbExclamation, TituloJanela
        gError 117671

    ElseIf iRetorno = -5 Then
        MsgBox "Erro ao abrir a porta de comunicação.", vbOKOnly + vbExclamation, TituloJanela
        gError 117672

    ElseIf iRetorno = -6 Then
        MsgBox "Impressora desligada ou cabo de comunicação desconectado.", vbOKOnly + vbExclamation, TituloJanela
        gError 117673

    ElseIf iRetorno = -7 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Banco não encontrado no arquivo BemaFI32.ini.", vbOKOnly + vbExclamation, TituloJanela
        gError 117674

    ElseIf iRetorno = -8 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Erro ao criar ou gravar no arquivo status.txt ou iRetorno.txt.", vbOKOnly + vbExclamation, TituloJanela
        gError 117675

    ElseIf iRetorno = -18 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Não foi possível abrir arquivo INTPOS.001 !", vbOKOnly + vbExclamation, TituloJanela
        gError 117676

    ElseIf iRetorno = -19 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Parâmetro diferentes !", vbOKOnly + vbExclamation, TituloJanela
        gError 117677

    ElseIf iRetorno = -20 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Transação cancelada pelo Operador !", vbOKOnly + vbExclamation, TituloJanela
        gError 117678

    ElseIf iRetorno = -21 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "A Transação não foi aprovada !", vbOKOnly + vbExclamation, TituloJanela
        gError 117679

    ElseIf iRetorno = -22 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Não foi possível terminal a Impressão !", vbOKOnly + vbExclamation, TituloJanela
        gError 117680

    ElseIf iRetorno = -23 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Não foi possível terminal a Operação !", vbOKOnly + vbExclamation, TituloJanela
        gError 117681

    ElseIf iRetorno = -24 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Forma de pagamento não programada.", vbOKOnly + vbExclamation, TituloJanela
        gError 117682

    ElseIf iRetorno = -25 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Totalizador não fiscal não programado.", vbOKOnly + vbExclamation, TituloJanela
        gError 117683

    ElseIf iRetorno = -26 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Transação já realizada.", vbOKOnly + vbExclamation, TituloJanela
        gError 117684

'    ElseIf iRetorno = -27 Then
'
'        'desbloqueia mouse e teclado
'        Call AFRAC_DesbloqueiaTecladoMouse
'
'        MsgBox "Status diferente de 6,0,0.", vbOKOnly + vbExclamation, TituloJanela
'        gError 117685

    ElseIf iRetorno = -28 Then

        'desbloqueia mouse e teclado
        Call AFRAC_DesbloqueiaTecladoMouse

        MsgBox "Não há dados para serem impressos.", vbOKOnly + vbExclamation, TituloJanela
        gError 117686
    End If

    VerificaRetornoImpressora = SUCESSO
    
    Exit Function
    
Erro_VerificaRetornoImpressora:

    VerificaRetornoImpressora = gErr
    
    Select Case gErr
    
        Case 117665 To 117686, 126800
        
            If X.giCodModeloECF = IMPRESSORA_DARUMA_FS600 Then
                Call Daruma_FI_ResetaImpressora
            Else
                Call Bematech_FI_ResetaImpressora
            End If
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144631)
    
    End Select

    Exit Function

End Function

Public Function AFRAC_VerificaImpressoraLigada() As Integer

'Verifica se a impressora esta ligada
Dim X As New ClassECFConfig
Dim sCodigoErro As String
Dim sMsg As String
Dim iAcao As Integer
Dim lErro As Long
Dim iStatus As Integer
Dim sPath As String

On Error GoTo Erro_AFRAC_VerificaImpressoraLigada

    AFRAC_VerificaImpressoraLigada = 1

    X.giTotalizador_Sangria = 1
    X.giTotalizador_Suprimento = 2
    X.giTotalizador_Orcamento = 3

    Select Case X.giCodModeloECF

        Case 1 To 2, 5, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            AFRAC_VerificaImpressoraLigada = Bematech_FI_VerificaImpressoraLigada()
            
            If AFRAC_VerificaImpressoraLigada = 1 Then
                sPath = IIf(right(CurDir(), 1) <> "\", CurDir() & "\", CurDir())
                lErro = WritePrivateProfileString(SECAO_SISTEMA, "Path", sPath, BEMATECH_INI_32)
                lErro = WritePrivateProfileString(SECAO_SISTEMA, "Path", sPath, BEMATECH_INI_64)
                
                If X.giCodModeloECF = 5 Then
                    'habilita o retorno extendido em caso de erro (ST3) nas impressoras termicas
                    iStatus = Bematech_FI_HabilitaDesabilitaRetornoEstendidoMFD("1")
                End If
                
                AFRAC_VerificaImpressoraLigada = SUCESSO
            ElseIf AFRAC_VerificaImpressoraLigada = 0 Then
                AFRAC_VerificaImpressoraLigada = 1
            End If
            
        'a impressora MP4200 tem pre-definido a sangria como 2 e o troco como 1, vide help de Bematech_FI_NomeiaTotalizadorNaoSujeitoIcms
        Case IMPRESSORA_BEMATECH_MP4200_TH_FI
        
            AFRAC_VerificaImpressoraLigada = Bematech_FI_VerificaImpressoraLigada()
            
            If AFRAC_VerificaImpressoraLigada = 1 Then
                sPath = IIf(right(CurDir(), 1) <> "\", CurDir() & "\", CurDir())
                lErro = WritePrivateProfileString(SECAO_SISTEMA, "Path", sPath, BEMATECH_INI_32)
                lErro = WritePrivateProfileString(SECAO_SISTEMA, "Path", sPath, BEMATECH_INI_64)
                iStatus = Bematech_FI_HabilitaDesabilitaRetornoEstendidoMFD("1")
                X.giTotalizador_Sangria = 2
                X.giTotalizador_Suprimento = 3
                X.giTotalizador_Orcamento = 4
                AFRAC_VerificaImpressoraLigada = SUCESSO
            ElseIf AFRAC_VerificaImpressoraLigada = 0 Then
                AFRAC_VerificaImpressoraLigada = 1
            End If
            
            
        Case IMPRESSORA_DARUMA_FS600
        
            AFRAC_VerificaImpressoraLigada = rVerificarImpressoraLigada_ECF_Daruma()

'                AFRAC_VerificaImpressoraLigada = Daruma_FI_VerificaImpressoraLigada()
            
            
            If AFRAC_VerificaImpressoraLigada = 1 Then
'                iStatus = Bematech_FI_HabilitaDesabilitaRetornoEstendidoMFD("1")
                AFRAC_VerificaImpressoraLigada = regAlterarValor_Daruma("START\LocalArquivosRelatorios", IIf(right(CurDir(), 1) <> "\", CurDir() & "\", CurDir()))
                AFRAC_VerificaImpressoraLigada = regAlterarValor_Daruma("START\LocalArquivos", IIf(right(CurDir(), 1) <> "\", CurDir() & "\", CurDir()))
                AFRAC_VerificaImpressoraLigada = regAlterarValor_Daruma("ECF\ReducaoZAutomatica", "1")
            End If
            
            If AFRAC_VerificaImpressoraLigada = 1 Then
                AFRAC_VerificaImpressoraLigada = SUCESSO
            ElseIf AFRAC_VerificaImpressoraLigada = 0 Then
                AFRAC_VerificaImpressoraLigada = 1
            End If
            
        Case Else
            AFRAC_VerificaImpressoraLigada = SUCESSO
            
    End Select


    Exit Function
   
Erro_AFRAC_VerificaImpressoraLigada:

    AFRAC_VerificaImpressoraLigada = 1
    
    Exit Function
    

End Function

Public Function AFRAC_MapaResumo() As Integer

Dim X As New ClassECFConfig
Dim vbMsgRes As VbMsgBoxResult

'Função que gera o maparesumo ECF conforme convenio ICMS_85

On Error GoTo Erro_AFRAC_MapaResumo

    AFRAC_MapaResumo = 1

    Select Case X.giCodModeloECF

        Case 1 To 3
        
            AFRAC_MapaResumo = Bematech_FI_MapaResumo()
            
            If AFRAC_MapaResumo = 1 Then
                AFRAC_MapaResumo = SUCESSO
            ElseIf AFRAC_MapaResumo = 0 Then
                AFRAC_MapaResumo = 1
            End If

        Case IMPRESSORA_DARUMA_FS600
        
'                AFRAC_MapaResumo = Daruma_FI_MapaResumo()
            AFRAC_MapaResumo = rGerarMapaResumo_ECF_Daruma()
            
            If AFRAC_MapaResumo = 1 Then
            
                vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "MAPA_RESUMO.txt")
            
            End If
            
            If AFRAC_MapaResumo = 1 Then
                AFRAC_MapaResumo = SUCESSO
            ElseIf AFRAC_MapaResumo = 0 Then
                AFRAC_MapaResumo = 1
            End If

        Case Else
            AFRAC_MapaResumo = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_MapaResumo:

    AFRAC_MapaResumo = 1
    
    Exit Function

End Function

'Public Function AFRAC_VerificaTotalizadoresParciais(sTotalizadores As String) As Integer
'
'Dim X As New ClassECFConfig
'
'
'    Select Case X.giCodModeloECF
'
'        Case 1 To 3, 5, IMPRESSORA_DARUMA_FS600
'
'            If X.giCodModeloECF = IMPRESSORA_DARUMA_FS600 Then
'                AFRAC_VerificaTotalizadoresParciais = Daruma_FI_VerificaTotalizadoresParciais(sTotalizadores)
'            Else
'                AFRAC_VerificaTotalizadoresParciais = Bematech_FI_VerificaTotalizadoresParciais(sTotalizadores)
'            End If
'
'            If AFRAC_VerificaTotalizadoresParciais = 1 Then
'                AFRAC_VerificaTotalizadoresParciais = SUCESSO
'            Else
'                AFRAC_VerificaTotalizadoresParciais = 1
'            End If
'
'        Case IMPRESSORA_ELGIN
'
'
'    End Select
'
'End Function

Public Function AFRAC_RelatorioTipo60Mestre() As Integer

Dim X As New ClassECFConfig

'Função que gera o maparesumo ECF conforme convenio ICMS_85

On Error GoTo Erro_AFRAC_RelatorioTipo60Mestre

    AFRAC_RelatorioTipo60Mestre = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            AFRAC_RelatorioTipo60Mestre = Bematech_FI_RelatorioTipo60Mestre()
            
            If AFRAC_RelatorioTipo60Mestre = 1 Then
                AFRAC_RelatorioTipo60Mestre = SUCESSO
            ElseIf AFRAC_RelatorioTipo60Mestre = 0 Then
                AFRAC_RelatorioTipo60Mestre = 1
            End If

        Case IMPRESSORA_DARUMA_FS600
        
'            AFRAC_RelatorioTipo60Mestre = Daruma_FI_RelatorioTipo60Mestre()
            
            If AFRAC_RelatorioTipo60Mestre = 1 Then
                AFRAC_RelatorioTipo60Mestre = SUCESSO
            ElseIf AFRAC_RelatorioTipo60Mestre = 0 Then
                AFRAC_RelatorioTipo60Mestre = 1
            End If

        Case Else
            AFRAC_RelatorioTipo60Mestre = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_RelatorioTipo60Mestre:

    AFRAC_RelatorioTipo60Mestre = 1
    
    Exit Function


End Function

'Public Function AFRAC_RelatorioTipo60Analitico() As Integer
'
'Dim X As New ClassECFConfig
'
''Função que gera o maparesumo ECF conforme convenio ICMS_85
'
'    Select Case X.giCodModeloECF
'
'        Case 1 To 3, 5, IMPRESSORA_DARUMA_FS600
'
'            If X.giCodModeloECF = IMPRESSORA_DARUMA_FS600 Then
'                AFRAC_RelatorioTipo60Analitico = Daruma_FI_RelatorioTipo60Analitico()
'            Else
'                AFRAC_RelatorioTipo60Analitico = Bematech_FI_RelatorioTipo60Analitico()
'            End If
'
'            If AFRAC_RelatorioTipo60Analitico = 1 Then
'                AFRAC_RelatorioTipo60Analitico = SUCESSO
'            Else
'                AFRAC_RelatorioTipo60Analitico = 1
'            End If
'    End Select
'
'End Function

'Public Function AFRAC_Carrega_Dados_MestreAnalitico(ByVal objLojaArqFisMestre As ClassLojaArqFisMestre, ByVal colLojaArqFisAnalitico As Collection) As Long
'
'Dim X As New ClassECFConfig
'Dim lErro As Long
'Dim lTamanho As Long
'Dim sRetorno As String
'Dim sRetorno1 As String
'Dim sArquivo As String
'Dim iPosFim As Integer
'Dim iPosIni As Integer
'Dim sChar As String
'Dim objLojaArqFisAnalitico As ClassLojaArqFisAnalitico
'Dim dtData As Date
'Dim sRegistro As String
'Dim sRegistro1 As String
'Dim iStatus As Integer
'Dim sData As String
'Dim iCaracter As Integer
'Dim iIndice As Integer
'Dim dTotalISS As Double
'Dim iCaracter1 As Integer
'Dim iCaracter2 As Integer
'Dim iCaracter3 As Integer
'Dim iCaracter4 As Integer
'Dim sCaracter1 As String
'Dim sCaracter2 As String
'Dim sCaracter3 As String
'Dim sCaracter4 As String
'Dim lTotal As Long
'
'On Error GoTo Erro_AFRAC_Carrega_Dados_MestreAnalitico
'
'    Select Case X.giCodModeloECF
'
'        Case 1 To 3, 5, IMPRESSORA_DARUMA_FS600
'
'            lErro = AFRAC_RelatorioTipo60Mestre()
'            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Relatorio Tipo 60 Mestre")
'            If lErro <> SUCESSO Then gError 133664
'
'            If X.giCodModeloECF = 1 Or X.giCodModeloECF = 2 Or X.giCodModeloECF = 3 Or X.giCodModeloECF = 5 Then
'
'                lTamanho = 255
'                sRetorno1 = String(lTamanho, 0)
'
'                Call GetSystemDirectory(sRetorno1, lTamanho)
'
'                'Retira os espaços no final da string
'                sRetorno1 = StringZ(sRetorno1)
'
'                'se o diretorio nao for terminado por \  ===> acrescentar
'                If right(sRetorno1, 1) <> "\" Then sRetorno1 = sRetorno1 & "\"
'
'                lTamanho = 255
'                sRetorno = String(lTamanho, 0)
'
'                'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
'                Call GetPrivateProfileString("Sistema", "Path", CONSTANTE_ERRO, sRetorno, lTamanho, sRetorno1 & "BemaFI32.ini")
'
'                'Retira os espaços no final da string
'                sRetorno = StringZ(sRetorno)
'
'                'Se não encontrou
'                If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 133666
'
'                'se o diretorio nao for terminado por \  ===> acrescentar
'                If right(sRetorno, 1) <> "\" Then sRetorno = sRetorno & "\"
'
'                sArquivo = sRetorno & "RETORNO.TXT"
'
'            ElseIf X.giCodModeloECF = IMPRESSORA_DARUMA_FS600 Then
'
'                sArquivo = "c:\Retorno.txt"
'
'            End If
'
'            Open sArquivo For Input As #1
'
'            Do While Not EOF(1)
'
'                'Busca o próximo registro do arquivo
'                Line Input #1, sRegistro
'                Line Input #1, sRegistro1
'
'                iPosFim = Len(sRegistro) + 1
'                iPosIni = iPosFim
'                sChar = Mid(sRegistro, iPosIni - 1, 1)
'                Do While sChar <> " "
'                    iPosIni = iPosIni - 1
'                    sChar = Mid(sRegistro, iPosIni - 1, 1)
'                Loop
'
'                If left(sRegistro, 15) = "Data de emissão" Then
'                    If Mid(sRegistro, iPosIni, iPosFim - iPosIni) <> "00/00/00" And Mid(sRegistro, iPosIni, iPosFim - iPosIni) <> "//" Then
'                        objLojaArqFisMestre.dtData = StrParaDate(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                    End If
'                ElseIf left(sRegistro, 15) = "Número de série" Then
'                    objLojaArqFisMestre.sNumSerieECF = Mid(sRegistro, iPosIni, iPosFim - iPosIni)
'                ElseIf left(sRegistro, 21) = "Número do equipamento" Then
'                    objLojaArqFisMestre.iNumEquip = StrParaInt(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                ElseIf left(sRegistro, 11) = "COO inicial" Then
'                    objLojaArqFisMestre.lCOOIni = StrParaLong(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                ElseIf left(sRegistro, 9) = "COO final" Then
'                    objLojaArqFisMestre.lCOOFim = StrParaLong(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                ElseIf left(sRegistro, 20) = "Contador de reduções" Then
'                    objLojaArqFisMestre.lCRZ = StrParaLong(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                ElseIf left(sRegistro, 20) = "Reinicio de Operação" Then
'                    objLojaArqFisMestre.lCRO = StrParaLong(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                ElseIf left(sRegistro, 11) = "Venda Bruta" Then
'                    objLojaArqFisMestre.dVendaBruta = StrParaDbl(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                ElseIf left(sRegistro, 17) = "Totalizador geral" Then
'                    objLojaArqFisMestre.dGrandeTotal = StrParaDbl(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                End If
'
'            Loop
'
'            Close #1
'
'            lErro = AFRAC_RelatorioTipo60Analitico()
'            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Relatorio Tipo 60 Mestre")
'            If lErro <> SUCESSO Then gError 133667
'
'            Open sArquivo For Input As #1
'
'            Do While Not EOF(1)
'
'                'Busca o próximo registro do arquivo
'                Line Input #1, sRegistro
'                Line Input #1, sRegistro1
'
'                iPosFim = Len(sRegistro) + 1
'                iPosIni = iPosFim
'                sChar = Mid(sRegistro, iPosIni - 1, 1)
'                Do While sChar <> " "
'                    iPosIni = iPosIni - 1
'                    sChar = Mid(sRegistro, iPosIni - 1, 1)
'                Loop
'                If left(sRegistro, 17) = "Tipo do relatório" Then
'                ElseIf left(sRegistro, 7) = "Subtipo" Then
'                ElseIf left(sRegistro, 15) = "Data de emissão" Then
'                    If Mid(sRegistro, iPosIni, iPosFim - iPosIni) <> "00/00/00" And Mid(sRegistro, iPosIni, iPosFim - iPosIni) <> "//" Then
'                        dtData = StrParaDate(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                    End If
'                ElseIf left(sRegistro, 15) = "Número de série" Then
'                ElseIf left(sRegistro, 13) = "Cancelamentos" Then
'                    Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                    objLojaArqFisAnalitico.dtData = dtData
'                    objLojaArqFisAnalitico.sSituacaoTrib = "CANC"
'                    objLojaArqFisAnalitico.dTotalizador = StrParaDbl(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                    colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'                ElseIf left(sRegistro, 9) = "Descontos" Then
'                    Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                    objLojaArqFisAnalitico.dtData = dtData
'                    objLojaArqFisAnalitico.sSituacaoTrib = "DESC"
'                    objLojaArqFisAnalitico.dTotalizador = StrParaDbl(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                    colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'                ElseIf left(sRegistro, 2) = "F." Then
'                    Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                    objLojaArqFisAnalitico.dtData = dtData
'                    objLojaArqFisAnalitico.sSituacaoTrib = "F"
'                    objLojaArqFisAnalitico.dTotalizador = StrParaDbl(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                    colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'                ElseIf left(sRegistro, 3) = "ISS" Then
'                    Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                    objLojaArqFisAnalitico.dtData = dtData
'                    objLojaArqFisAnalitico.sSituacaoTrib = "ISS"
'                    objLojaArqFisAnalitico.dTotalizador = StrParaDbl(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                    colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'                ElseIf left(sRegistro, 2) = "I." Then
'                    Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                    objLojaArqFisAnalitico.dtData = dtData
'                    objLojaArqFisAnalitico.sSituacaoTrib = "I"
'                    objLojaArqFisAnalitico.dTotalizador = StrParaDbl(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                    colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'                ElseIf left(sRegistro, 2) = "N." Then
'                    Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                    objLojaArqFisAnalitico.dtData = dtData
'                    objLojaArqFisAnalitico.sSituacaoTrib = "N"
'                    objLojaArqFisAnalitico.dTotalizador = StrParaDbl(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                    colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'                ElseIf IsNumeric(left(sRegistro, 4)) Then
'                    Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                    objLojaArqFisAnalitico.dtData = dtData
'                    objLojaArqFisAnalitico.sSituacaoTrib = left(sRegistro, 4)
'                    objLojaArqFisAnalitico.dTotalizador = StrParaDbl(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                    colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'                End If
'
'            Loop
'
'            Close #1
'
'        Case IMPRESSORA_ELGIN
'
'            lErro = AFRAC_LerInformacaoImpressora("018", sData)
'
'            sData = Mid(sData, 1, 2) & "/" & Mid(sData, 3, 2) & "/" & Mid(sData, 5, 2)
'
'            objLojaArqFisMestre.dtData = StrParaDate(sData)
'
'
'            lErro = AFRAC_LerInformacaoImpressora("002", sRetorno)
'
'            objLojaArqFisMestre.sNumSerieECF = sRetorno
'
'            lErro = AFRAC_LerInformacaoImpressora("015", sRetorno)
'
'            objLojaArqFisMestre.iNumEquip = sRetorno
'
'            iStatus = Elgin_EcfPar("82")
'
'            sRetorno = Space(50)
'
'            iStatus = Elgin_EsperaResposta(sRetorno)
'
'            If iStatus = SUCESSO Then
'                sRetorno = Mid(sRetorno, 6, 12)
'            End If
'
'            objLojaArqFisMestre.lCOOIni = StrParaLong(Mid(sRetorno, 7, 6))
'            objLojaArqFisMestre.lCOOFim = StrParaLong(Mid(sRetorno, 1, 6))
'
'            lErro = AFRAC_LerInformacaoImpressora("013", sRetorno)
'
'
'            objLojaArqFisMestre.lCRZ = StrParaLong(sRetorno)
'
'            lErro = AFRAC_LerInformacaoImpressora("011", sRetorno)
'
'
'            objLojaArqFisMestre.lCRO = StrParaLong(sRetorno)
'
'            iStatus = Elgin_EcfPar("40")
'
'            sRetorno = Space(50)
'            iStatus = Elgin_EsperaResposta(sRetorno)
'
'            If iStatus = SUCESSO Then
'                sRetorno = Mid(sRetorno, 6, 19)
'                objLojaArqFisMestre.dVendaBruta = StrParaDbl(sRetorno) / 100
'            End If
'
'            lErro = AFRAC_LerInformacaoImpressora("007", sRetorno)
'
'            objLojaArqFisMestre.dGrandeTotal = StrParaDbl(sRetorno) / 100
'
'
'            'pega o total de cancelamentos
'
'            iStatus = Elgin_EcfPar("38")
'
'            sRetorno = Space(50)
'            iStatus = Elgin_EsperaResposta(sRetorno)
'
'
'            If iStatus = SUCESSO Then
'                sRetorno = Mid(sRetorno, 6, 15)
'                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                objLojaArqFisAnalitico.dtData = objLojaArqFisMestre.dtData
'                objLojaArqFisAnalitico.sSituacaoTrib = "CANC"
'                objLojaArqFisAnalitico.dTotalizador = StrParaDbl(sRetorno) / 100
'                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'            End If
'
'
'            'pega o total de descontos
'
'            iStatus = Elgin_EcfPar("36")
'
'            sRetorno = Space(50)
'            iStatus = Elgin_EsperaResposta(sRetorno)
'
'            If iStatus = SUCESSO Then
'                sRetorno = Mid(sRetorno, 6, 15)
'                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                objLojaArqFisAnalitico.dtData = objLojaArqFisMestre.dtData
'                objLojaArqFisAnalitico.sSituacaoTrib = "DESC"
'                objLojaArqFisAnalitico.dTotalizador = StrParaDbl(sRetorno) / 100
'                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'            End If
'
'            'pega o total de tributacao na fonte
'
'            iStatus = Elgin_EcfPar("34")
'
'            sRetorno = Space(50)
'            iStatus = Elgin_EsperaResposta(sRetorno)
'
'            If iStatus = SUCESSO Then
'                sRetorno = Mid(sRetorno, 6, 15)
'                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                objLojaArqFisAnalitico.dtData = objLojaArqFisMestre.dtData
'                objLojaArqFisAnalitico.sSituacaoTrib = "F"
'                objLojaArqFisAnalitico.dTotalizador = StrParaDbl(sRetorno) / 100
'                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'            End If
'
'            'pega o total de isentos
'
'            iStatus = Elgin_EcfPar("32")
'
'            sRetorno = Space(50)
'            iStatus = Elgin_EsperaResposta(sRetorno)
'
'            If iStatus = SUCESSO Then
'                sRetorno = Mid(sRetorno, 6, 15)
'                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                objLojaArqFisAnalitico.dtData = objLojaArqFisMestre.dtData
'                objLojaArqFisAnalitico.sSituacaoTrib = "I"
'                objLojaArqFisAnalitico.dTotalizador = StrParaDbl(sRetorno) / 100
'                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'            End If
'
'            'pega o total de nao tributados
'
'            iStatus = Elgin_EcfPar("33")
'
'            sRetorno = Space(50)
'            iStatus = Elgin_EsperaResposta(sRetorno)
'
'            If iStatus = SUCESSO Then
'                sRetorno = Mid(sRetorno, 6, 15)
'                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                objLojaArqFisAnalitico.dtData = objLojaArqFisMestre.dtData
'                objLojaArqFisAnalitico.sSituacaoTrib = "N"
'                objLojaArqFisAnalitico.dTotalizador = StrParaDbl(sRetorno) / 100
'                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'            End If
'
'
'            'pega o total de indicador de totalizadores de iss/icms
'
'            iStatus = Elgin_EcfPar("86")
'
'            sRetorno = Space(50)
'            iStatus = Elgin_EsperaResposta(sRetorno)
'
'            If iStatus = SUCESSO Then
'                sRetorno = Mid(sRetorno, 8, 4)
'
'                sCaracter1 = Mid(sRetorno, 4, 1)
'                sCaracter2 = Mid(sRetorno, 3, 1)
'                sCaracter3 = Mid(sRetorno, 2, 1)
'                sCaracter4 = Mid(sRetorno, 1, 1)
'
'                Call StringHex_To_Int(sCaracter1, iCaracter1)
'                Call StringHex_To_Int(sCaracter2, iCaracter2)
'                Call StringHex_To_Int(sCaracter3, iCaracter3)
'                Call StringHex_To_Int(sCaracter4, iCaracter4)
'
'                lTotal = iCaracter1 + (iCaracter2 * 16) + (iCaracter3 * 256) + (iCaracter4 * 4096)
'
'                dTotalISS = 0
'
'                For iIndice = 0 To 15
'
'                    lErro = Trata_LojaArqFisAnalitico(objLojaArqFisMestre.dtData, lTotal, iIndice, dTotalISS, colLojaArqFisAnalitico)
'                    If lErro <> SUCESSO Then gError 204597
'
'                Next
'
'                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
'                objLojaArqFisAnalitico.dtData = objLojaArqFisMestre.dtData
'                objLojaArqFisAnalitico.sSituacaoTrib = "ISS"
'                objLojaArqFisAnalitico.dTotalizador = dTotalISS
'                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
'
'            End If
'
'    End Select
'
'    AFRAC_Carrega_Dados_MestreAnalitico = SUCESSO
'
'    Exit Function
'
'Erro_AFRAC_Carrega_Dados_MestreAnalitico:
'
'    AFRAC_Carrega_Dados_MestreAnalitico = gErr
'
'    Select Case gErr
'
'        Case 133664, 133667, 204597
'
'        Case 133666
'            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "Path", "Sistema", "BemaFI32.ini")
'
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 144632)
'
'    End Select
'
'    Close #1
'
'    Exit Function
'
'End Function

Public Function AFRAC_VerificaTotalizadoresNaoFiscais(sTotalizadores As String) As Integer

Dim X As New ClassECFConfig
Dim iIndice As Integer
Dim iStatus As Integer
Dim sRetorno As String

On Error GoTo Erro_AFRAC_VerificaTotalizadoresNaoFiscais

    AFRAC_VerificaTotalizadoresNaoFiscais = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sTotalizadores = Space(180)
            'so retorna os primeiros 9 totalizadores separados por virgula
            AFRAC_VerificaTotalizadoresNaoFiscais = Bematech_FI_VerificaTotalizadoresNaoFiscais(sTotalizadores)
            
            sTotalizadores = StringZ(sTotalizadores)
            
            If AFRAC_VerificaTotalizadoresNaoFiscais = 1 Then
                AFRAC_VerificaTotalizadoresNaoFiscais = SUCESSO
            ElseIf AFRAC_VerificaTotalizadoresNaoFiscais = 0 Then
                AFRAC_VerificaTotalizadoresNaoFiscais = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            sTotalizadores = Space(360)
            AFRAC_VerificaTotalizadoresNaoFiscais = rLerCNF_ECF_Daruma(sTotalizadores)
            
            If AFRAC_VerificaTotalizadoresNaoFiscais = 1 Then
                AFRAC_VerificaTotalizadoresNaoFiscais = SUCESSO
            ElseIf AFRAC_VerificaTotalizadoresNaoFiscais = 0 Then
                AFRAC_VerificaTotalizadoresNaoFiscais = 1
            End If
            
        Case IMPRESSORA_ELGIN
            For iIndice = 66 To 81
                
                iStatus = Elgin_EcfPar(Format(iIndice, "00"))
    
                sRetorno = Space(50)
                iStatus = Elgin_EsperaResposta(sRetorno)
                
                If iStatus = SUCESSO Then
                    sRetorno = Trim(Mid(sRetorno, 6, 16))
                    If Len(sRetorno) > 0 Then
                        sTotalizadores = sTotalizadores & sRetorno & ","
                    End If
                End If
                            
                    
            Next
            
            If Len(sTotalizadores) > 0 Then
                    
                sTotalizadores = left(sTotalizadores, Len(sTotalizadores) - 1)
            End If
            
            AFRAC_VerificaTotalizadoresNaoFiscais = SUCESSO
        
        Case Else
            AFRAC_VerificaTotalizadoresNaoFiscais = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_VerificaTotalizadoresNaoFiscais:

    AFRAC_VerificaTotalizadoresNaoFiscais = 1
    
    Exit Function
    

End Function

Public Function AFRAC_RecebimentoNaoFiscal(ByVal sIndiceTotalizador As String, ByVal sValor As String, ByVal sFormaPagamento As String) As Integer
'NAO ANALISADO PARA VERSAO PAFECF
Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String

'Função que gera o maparesumo ECF conforme convenio ICMS_85

On Error GoTo Erro_AFRAC_RecebimentoNaoFiscal

    AFRAC_RecebimentoNaoFiscal = 1
    
    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            AFRAC_RecebimentoNaoFiscal = Bematech_FI_RecebimentoNaoFiscal(sIndiceTotalizador, sValor, sFormaPagamento)
            
            If AFRAC_RecebimentoNaoFiscal = 1 Then
                AFRAC_RecebimentoNaoFiscal = SUCESSO
            ElseIf AFRAC_RecebimentoNaoFiscal = 0 Then
                AFRAC_RecebimentoNaoFiscal = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
'                AFRAC_RecebimentoNaoFiscal = Daruma_FI_RecebimentoNaoFiscal(sIndiceTotalizador, sValor, sFormaPagamento)
            AFRAC_RecebimentoNaoFiscal = iCNFAbrirPadrao_ECF_Daruma()
            If AFRAC_RecebimentoNaoFiscal = 1 Then
                AFRAC_RecebimentoNaoFiscal = iCNFReceberSemDesc_ECF_Daruma(sIndiceTotalizador, sValor)
                If AFRAC_RecebimentoNaoFiscal = 1 Then
                    AFRAC_RecebimentoNaoFiscal = iCNFTotalizarComprovantePadrao_ECF_Daruma()
                    If AFRAC_RecebimentoNaoFiscal = 1 Then
                        AFRAC_RecebimentoNaoFiscal = iCNFEfetuarPagamento_ECF_Daruma(sFormaPagamento, sValor, "Obrigado. Volte Sempre.")
                        If AFRAC_RecebimentoNaoFiscal = 1 Then
                            AFRAC_RecebimentoNaoFiscal = iCNFEncerrarPadrao_ECF_Daruma()
                        End If
                    End If
                End If
            End If
            
            If AFRAC_RecebimentoNaoFiscal = 1 Then
                AFRAC_RecebimentoNaoFiscal = SUCESSO
            ElseIf AFRAC_RecebimentoNaoFiscal = 0 Then
                AFRAC_RecebimentoNaoFiscal = 1
            End If
            
        Case IMPRESSORA_ELGIN
            iStatus = Elgin_OperRegNaoVinculado(Format(StrParaInt(sIndiceTotalizador), "00"), Format(StrParaDbl(sValor), "000000000000000"), "@", "&", "000000000000000", "")
            
            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)
            
            If iStatus = SUCESSO Then
                AFRAC_RecebimentoNaoFiscal = SUCESSO
            Else
                AFRAC_RecebimentoNaoFiscal = 1
            End If
        
        Case Else
            AFRAC_RecebimentoNaoFiscal = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_RecebimentoNaoFiscal:

    AFRAC_RecebimentoNaoFiscal = 1
    
    Exit Function
    

End Function

Public Function AFRAC_ValorTotalizadorNaoFiscal(ByVal sTotalizador As String, sValor As String) As Integer
'NAO ANALISADO PARA VERSAO PAFECF
Dim X As New ClassECFConfig
Dim iIndice As Integer
Dim iStatus As Integer
Dim sRetorno As String

'Função que gera o maparesumo ECF conforme convenio ICMS_85

On Error GoTo Erro_AFRAC_ValorTotalizadorNaoFiscal

    AFRAC_ValorTotalizadorNaoFiscal = 1
    
    Select Case X.giCodModeloECF

        Case 1 To 3
        
            AFRAC_ValorTotalizadorNaoFiscal = Bematech_FI_ValorTotalizadorNaoFiscal(sTotalizador, sValor)
            
            If AFRAC_ValorTotalizadorNaoFiscal = 1 Then
                AFRAC_ValorTotalizadorNaoFiscal = SUCESSO
            ElseIf AFRAC_ValorTotalizadorNaoFiscal = 0 Then
                AFRAC_ValorTotalizadorNaoFiscal = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
'                AFRAC_ValorTotalizadorNaoFiscal = Daruma_FI_ValorTotalizadorNaoFiscal(sTotalizador, sValor)
            
            If AFRAC_ValorTotalizadorNaoFiscal = 1 Then
                AFRAC_ValorTotalizadorNaoFiscal = SUCESSO
            ElseIf AFRAC_ValorTotalizadorNaoFiscal = 0 Then
                AFRAC_ValorTotalizadorNaoFiscal = 1
            End If
            
        Case IMPRESSORA_ELGIN
            For iIndice = 66 To 81
                
                iStatus = Elgin_EcfPar(Format(iIndice, "00"))
    
                sRetorno = Space(50)
                iStatus = Elgin_EsperaResposta(sRetorno)
                
                If iStatus = SUCESSO Then
                    sRetorno = Trim(Mid(sRetorno, 6, 16))
                    If sRetorno = sTotalizador Then
                        sValor = Mid(sRetorno, 22, 15)
                        Exit For
                    End If
                        
                End If
                            
            Next
                
            AFRAC_ValorTotalizadorNaoFiscal = SUCESSO
            
            
        Case Else
            AFRAC_ValorTotalizadorNaoFiscal = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_ValorTotalizadorNaoFiscal:

    AFRAC_ValorTotalizadorNaoFiscal = 1
    
    Exit Function

End Function

Public Function AFRAC_VerificaAliquotasIss(sAliquotasISS As String) As Integer

Dim X As New ClassECFConfig
Dim lErro As Long


On Error GoTo Erro_AFRAC_VerificaAliquotasIss

    AFRAC_VerificaAliquotasIss = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sAliquotasISS = Space(80)
            AFRAC_VerificaAliquotasIss = Bematech_FI_VerificaAliquotasIss(sAliquotasISS)
            
            sAliquotasISS = StringZ(sAliquotasISS)
            
            If AFRAC_VerificaAliquotasIss = 1 Then
                AFRAC_VerificaAliquotasIss = SUCESSO
            ElseIf AFRAC_VerificaAliquotasIss = 0 Then
                AFRAC_VerificaAliquotasIss = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
'                AFRAC_VerificaAliquotasIss = Daruma_FI_VerificaAliquotasIss(sAliquotasISS)
            
            If AFRAC_VerificaAliquotasIss = 1 Then
                AFRAC_VerificaAliquotasIss = SUCESSO
            ElseIf AFRAC_VerificaAliquotasIss = 0 Then
                AFRAC_VerificaAliquotasIss = 1
            End If
            
        Case IMPRESSORA_ELGIN
        
            lErro = Trata_AliquotaISS(sAliquotasISS)
            
            If lErro = SUCESSO Then
                AFRAC_VerificaAliquotasIss = SUCESSO
            Else
                AFRAC_VerificaAliquotasIss = 1
            End If
            
        Case Else
            AFRAC_VerificaAliquotasIss = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_VerificaAliquotasIss:

    AFRAC_VerificaAliquotasIss = 1
    
    Exit Function
    

End Function

Public Function AFRAC_VerificaIndiceAliquotasIss(sIndiceAliquotasISS As String) As Integer

Dim X As New ClassECFConfig
Dim lErro As Long

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sIndiceAliquotasISS = Space(49)
            AFRAC_VerificaIndiceAliquotasIss = Bematech_FI_VerificaIndiceAliquotasIss(sIndiceAliquotasISS)
            
            sIndiceAliquotasISS = StringZ(sIndiceAliquotasISS)
            
            If AFRAC_VerificaIndiceAliquotasIss = 1 Then
                AFRAC_VerificaIndiceAliquotasIss = SUCESSO
            ElseIf AFRAC_VerificaIndiceAliquotasIss = 0 Then
                AFRAC_VerificaIndiceAliquotasIss = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
'                AFRAC_VerificaIndiceAliquotasIss = Daruma_FI_VerificaIndiceAliquotasIss(sIndiceAliquotasISS)
            
            If AFRAC_VerificaIndiceAliquotasIss = 1 Then
                AFRAC_VerificaIndiceAliquotasIss = SUCESSO
            ElseIf AFRAC_VerificaIndiceAliquotasIss = 0 Then
                AFRAC_VerificaIndiceAliquotasIss = 1
            End If
            
        Case IMPRESSORA_ELGIN
        
            sIndiceAliquotasISS = ""
            lErro = Trata_IndiceAliquotaISS(sIndiceAliquotasISS)
        
            If lErro = SUCESSO Then
                AFRAC_VerificaIndiceAliquotasIss = SUCESSO
            Else
                AFRAC_VerificaIndiceAliquotasIss = 1
            End If
        
    End Select

End Function

Public Function AFRAC_BloqueiaTecladoMouse() As Integer

Dim X As New ClassECFConfig

On Error GoTo Erro_AFRAC_BloqueiaTecladoMouse

    AFRAC_BloqueiaTecladoMouse = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_DARUMA_FS600, IMPRESSORA_ELGIN, IMPRESSORA_BEMATECH_MP2100_TH_FI
            BlockInput True   'Bloqueia o teclado
            AFRAC_BloqueiaTecladoMouse = SUCESSO
            
'            AFRAC_BloqueiaTecladoMouse = Bematech_FI_IniciaModoTEF()
'            If AFRAC_BloqueiaTecladoMouse = 1 Then
'                AFRAC_BloqueiaTecladoMouse = SUCESSO
'            Else
'                AFRAC_BloqueiaTecladoMouse = 1
'            End If

        Case Else
            AFRAC_BloqueiaTecladoMouse = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_BloqueiaTecladoMouse:

    AFRAC_BloqueiaTecladoMouse = 1
    
    Exit Function

End Function

Public Function AFRAC_DesbloqueiaTecladoMouse() As Integer

Dim X As New ClassECFConfig

On Error GoTo Erro_AFRAC_DesbloqueiaTecladoMouse

    AFRAC_DesbloqueiaTecladoMouse = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_DARUMA_FS600, IMPRESSORA_ELGIN, IMPRESSORA_BEMATECH_MP2100_TH_FI
            BlockInput False  'Desbloqueia o teclado
            AFRAC_DesbloqueiaTecladoMouse = SUCESSO

        
'            AFRAC_DesbloqueiaTecladoMouse = Bematech_FI_FinalizaModoTEF()
'            If AFRAC_DesbloqueiaTecladoMouse = 1 Then
'                AFRAC_DesbloqueiaTecladoMouse = SUCESSO
'            Else
'                AFRAC_DesbloqueiaTecladoMouse = 1
'            End If
        
        Case Else
            AFRAC_DesbloqueiaTecladoMouse = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_DesbloqueiaTecladoMouse:

    AFRAC_DesbloqueiaTecladoMouse = 1
    
    Exit Function

End Function

Public Function AFRAC_Le_Balanca(ByVal sNomeModelo As String, ByVal sPorta As String, ByVal iModelo As Integer, sPeso As String, sPrecoKilo As String, sPrecoTotal As String) As Integer

Dim X As New ClassECFConfig

On Error GoTo Erro_AFRAC_Le_Balanca

    AFRAC_Le_Balanca = 1

    Select Case sNomeModelo
        
        Case "BP6", "CS15", "SA-110 Prot 0", "SA-110 Prot 4"

            sPeso = String(20, 0)
            sPrecoKilo = String(20, 0)
            sPrecoTotal = String(20, 0)
        
            'para tratar balancas bematech
            AFRAC_Le_Balanca = Bematech_FI_InfoBalanca(sPorta, iModelo, sPeso, sPrecoKilo, sPrecoTotal)
            If AFRAC_Le_Balanca = 1 Then
                AFRAC_Le_Balanca = SUCESSO
            ElseIf AFRAC_Le_Balanca = 0 Then
                AFRAC_Le_Balanca = 1
            End If
            
            sPeso = StringZ(sPeso)
            sPrecoKilo = StringZ(sPrecoKilo)
            sPrecoTotal = StringZ(sPrecoTotal)

        Case Else
            AFRAC_Le_Balanca = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_Le_Balanca:

    AFRAC_Le_Balanca = 1
    
    Exit Function

End Function

Public Function AFRAC_Le_FlagsFiscais(iFlag As Integer) As Integer

Dim X As New ClassECFConfig
Dim sBuffer As String
Dim iStatus As Integer
Dim sRetorno As String
Dim iStatus1 As Integer
Dim sFlags As String

On Error GoTo Erro_AFRAC_Le_FlagsFiscais

    AFRAC_Le_FlagsFiscais = 1

    Select Case X.giCodModeloECF

        Case 1 To 2, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            AFRAC_Le_FlagsFiscais = Bematech_FI_FlagsFiscais(iFlag)
            
            If AFRAC_Le_FlagsFiscais = 1 Then
                AFRAC_Le_FlagsFiscais = SUCESSO
            ElseIf AFRAC_Le_FlagsFiscais = 0 Then
                AFRAC_Le_FlagsFiscais = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            '    iStatus = Daruma_FI_FlagsFiscais(iFlag)
            
            sFlags = Space(20)
            AFRAC_Le_FlagsFiscais = rStatusImpressoraBinario_ECF_Daruma(sFlags)
            If AFRAC_Le_FlagsFiscais = 1 Then
                iFlag = CInt(Mid(sFlags, 12, 1))
            End If

            If AFRAC_Le_FlagsFiscais = 1 Then
                AFRAC_Le_FlagsFiscais = SUCESSO
            ElseIf AFRAC_Le_FlagsFiscais = 0 Then
                AFRAC_Le_FlagsFiscais = 1
            End If
            
        Case IMPRESSORA_ELGIN
        
            sBuffer = Space(41)
            
            iStatus = Elgin_TransStatus(0, sBuffer)
            
            sRetorno = Space(50)
            iStatus1 = Elgin_EsperaResposta(sRetorno)
            
            'se o cupom anterior nao estiver fechado...cancela
            If iStatus = SUCESSO Then
            
                If Mid(sBuffer, 8, 1) = "1" Then
                    iFlag = iFlag + 1
                End If
                
                If Mid(sBuffer, 12, 1) = "1" Then
                    iFlag = iFlag + 2
                End If
                
                If Mid(sBuffer, 2, 1) = "1" Then
                    iFlag = iFlag + 4
                End If
                
                If Mid(sBuffer, 6, 1) = "1" Then
                    iFlag = iFlag + 8
                End If
                
                If Mid(sBuffer, 1, 1) = "1" Then
                    iFlag = iFlag + 32
                End If
                
                If Mid(sBuffer, 22, 1) = "1" Then
                    iFlag = iFlag + 128
                End If
            
            End If
            
            If iStatus = SUCESSO Then
                AFRAC_Le_FlagsFiscais = SUCESSO
            Else
                AFRAC_Le_FlagsFiscais = 1
            End If
            
        Case Else
            AFRAC_Le_FlagsFiscais = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_Le_FlagsFiscais:

    AFRAC_Le_FlagsFiscais = 1
    
    Exit Function


End Function

Function Trata_LojaArqFisAnalitico(ByVal dtData As Date, ByVal lTotal As Long, ByVal iIndice As Integer, dTotalISS As Double, ByVal colLojaArqFisAnalitico As Collection) As Long

Dim iStatus As Integer
Dim sRetorno As String
Dim objLojaArqFisAnalitico As ClassLojaArqFisAnalitico
Dim iCaracterValor As Integer

On Error GoTo Erro_Trata_LojaArqFisAnalitico

    'pega o totalizador
    iStatus = Elgin_EcfPar(Format(iIndice, "00"))

    sRetorno = Space(50)
    iStatus = Elgin_EsperaResposta(sRetorno)
    
    If iStatus = SUCESSO Then
        sRetorno = Mid(sRetorno, 6, 15)
        
        'se o primeiro bit estiver setado ==> o totalizador 0 é vinculado ao ISS
        If lTotal And (2 ^ iIndice) Then
            dTotalISS = dTotalISS + StrParaDbl(sRetorno) / 100
        Else
            Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
            objLojaArqFisAnalitico.dtData = dtData
            objLojaArqFisAnalitico.dTotalizador = StrParaDbl(sRetorno) / 100
        
            iStatus = Elgin_EcfPar(Format(iIndice + 16, "00"))

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)
            
            If iStatus = SUCESSO Then
                objLojaArqFisAnalitico.sSituacaoTrib = Mid(sRetorno, 6, 4)
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
            End If
        End If
    End If

    Trata_LojaArqFisAnalitico = SUCESSO

    Exit Function

Erro_Trata_LojaArqFisAnalitico:

    Trata_LojaArqFisAnalitico = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204596)

    End Select
    
    Exit Function


End Function


Function Trata_AliquotaISS(sAliquotas As String) As Long

Dim iStatus As Integer
Dim sRetorno As String
Dim objLojaArqFisAnalitico As ClassLojaArqFisAnalitico
Dim iCaracterValor As Integer
Dim lTotal As Long
Dim iIndice As Integer
Dim dTotalISS As Double
Dim iCaracter1 As Integer
Dim iCaracter2 As Integer
Dim iCaracter3 As Integer
Dim iCaracter4 As Integer
Dim sCaracter1 As String
Dim sCaracter2 As String
Dim sCaracter3 As String
Dim sCaracter4 As String


On Error GoTo Erro_Trata_AliquotaISS


    'pega o total de indicador de totalizadores de iss/icms

    iStatus = Elgin_EcfPar("86")

    sRetorno = Space(50)
    iStatus = Elgin_EsperaResposta(sRetorno)
    
    If iStatus = SUCESSO Then
        sRetorno = Mid(sRetorno, 8, 4)
        
        sCaracter1 = Mid(sRetorno, 4, 1)
        sCaracter2 = Mid(sRetorno, 3, 1)
        sCaracter3 = Mid(sRetorno, 2, 1)
        sCaracter4 = Mid(sRetorno, 1, 1)
        
        Call StringHex_To_Int(sCaracter1, iCaracter1)
        Call StringHex_To_Int(sCaracter2, iCaracter2)
        Call StringHex_To_Int(sCaracter3, iCaracter3)
        Call StringHex_To_Int(sCaracter4, iCaracter4)
        
        lTotal = iCaracter1 + (iCaracter2 * 16) + (iCaracter3 * 256) + (iCaracter4 * 4096)
        
        dTotalISS = 0
        
        For iIndice = 0 To 15
                    
            'se o primeiro bit estiver setado ==> o totalizador 0 é vinculado ao ISS
            If lTotal And (2 ^ iIndice) Then
            
                iStatus = Elgin_EcfPar(Format(iIndice + 16, "00"))
    
                sRetorno = Space(50)
                iStatus = Elgin_EsperaResposta(sRetorno)
                
                If iStatus = SUCESSO Then
                    sAliquotas = sAliquotas & Mid(sRetorno, 6, 4) & ","
                    
                End If
            End If
        Next
    End If
    
    If Len(sAliquotas) > 0 Then
        sAliquotas = left(sAliquotas, Len(sAliquotas) - 1)
    End If

    Trata_AliquotaISS = SUCESSO

    Exit Function

Erro_Trata_AliquotaISS:

    Trata_AliquotaISS = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204598)

    End Select
    
    Exit Function


End Function

Function Trata_IndiceAliquotaISS(sIndiceAliquotas As String) As Long

Dim iStatus As Integer
Dim sRetorno As String
Dim objLojaArqFisAnalitico As ClassLojaArqFisAnalitico
Dim iCaracterValor As Integer
Dim iIndice As Integer
Dim iCaracter1 As Integer
Dim iCaracter2 As Integer
Dim iCaracter3 As Integer
Dim iCaracter4 As Integer
Dim sCaracter1 As String
Dim sCaracter2 As String
Dim sCaracter3 As String
Dim sCaracter4 As String
Dim lTotal As Long
Dim dTotalISS As Double


On Error GoTo Erro_Trata_IndiceAliquotaISS


    'pega o total de indicador de totalizadores de iss/icms

    iStatus = Elgin_EcfPar("86")

    sRetorno = Space(50)
    iStatus = Elgin_EsperaResposta(sRetorno)
    
    If iStatus = SUCESSO Then
        sRetorno = Mid(sRetorno, 8, 4)
        
        sCaracter1 = Mid(sRetorno, 4, 1)
        sCaracter2 = Mid(sRetorno, 3, 1)
        sCaracter3 = Mid(sRetorno, 2, 1)
        sCaracter4 = Mid(sRetorno, 1, 1)
        
        Call StringHex_To_Int(sCaracter1, iCaracter1)
        Call StringHex_To_Int(sCaracter2, iCaracter2)
        Call StringHex_To_Int(sCaracter3, iCaracter3)
        Call StringHex_To_Int(sCaracter4, iCaracter4)
        
        lTotal = iCaracter1 + (iCaracter2 * 16) + (iCaracter3 * 256) + (iCaracter4 * 4096)
        
        dTotalISS = 0
        
        For iIndice = 0 To 15
                    
            'se o primeiro bit estiver setado ==> o totalizador 0 é vinculado ao ISS
            If lTotal And (2 ^ iIndice) Then
                sIndiceAliquotas = sIndiceAliquotas & Format(iIndice, "00") & ","
            End If
        Next
    End If
    
    If Len(sIndiceAliquotas) > 0 Then
        sIndiceAliquotas = left(sIndiceAliquotas, Len(sIndiceAliquotas) - 1)
    End If

    Trata_IndiceAliquotaISS = SUCESSO

    Exit Function

Erro_Trata_IndiceAliquotaISS:

    Trata_IndiceAliquotaISS = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204599)

    End Select
    
    Exit Function


End Function

Sub StringHex_To_Int(ByVal sCaracter1 As String, iCaracter1 As Integer)

    If sCaracter1 = "F" Then
        iCaracter1 = 15
    ElseIf sCaracter1 = "E" Then
        iCaracter1 = 14
    ElseIf sCaracter1 = "D" Then
        iCaracter1 = 13
    ElseIf sCaracter1 = "C" Then
        iCaracter1 = 12
    ElseIf sCaracter1 = "B" Then
        iCaracter1 = 11
    ElseIf sCaracter1 = "A" Then
        iCaracter1 = 10
    Else
        iCaracter1 = StrParaInt(sCaracter1)
    End If

End Sub

Public Function AFRAC_DadosUltimaReducao(sReducao As String) As Integer
'
'Fecha a porta serial de comunicação entre o micro e a impressora
Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sDadosReducao As String

On Error GoTo Erro_AFRAC_DadosUltimaReducao

    AFRAC_DadosUltimaReducao = 1

    Select Case X.giCodModeloECF

        Case 1 To 3
        
            sReducao = Space(632)
    
            AFRAC_DadosUltimaReducao = Bematech_FI_DadosUltimaReducao(sReducao)

            sReducao = StringZ(sReducao)

            If AFRAC_DadosUltimaReducao = 1 Then
                AFRAC_DadosUltimaReducao = SUCESSO
            ElseIf AFRAC_DadosUltimaReducao = 0 Then
                AFRAC_DadosUltimaReducao = 1
            End If



        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sReducao = Space(1278)
    
            If giDebug = 1 Then MsgBox ("Bematech_FI_DadosUltimaReducaoMFD-1")
    
            AFRAC_DadosUltimaReducao = Bematech_FI_DadosUltimaReducaoMFD(sReducao)
            
            If giDebug = 1 Then MsgBox ("Bematech_FI_DadosUltimaReducaoMFD-2 " & CStr(AFRAC_DadosUltimaReducao))

            sReducao = StringZ(sReducao)

            If AFRAC_DadosUltimaReducao = 1 Then
                AFRAC_DadosUltimaReducao = SUCESSO
            ElseIf AFRAC_DadosUltimaReducao = 0 Then
                AFRAC_DadosUltimaReducao = 1
            End If


        Case IMPRESSORA_DARUMA_FS600
            sReducao = Space(1209)
        
'            AFRAC_DadosUltimaReducao = Daruma_FIMFD_RetornaInformacao("140", sReducao)
            AFRAC_DadosUltimaReducao = rRetornarDadosReducaoZ_ECF_Daruma(sReducao)

            If AFRAC_DadosUltimaReducao = 1 Then
                AFRAC_DadosUltimaReducao = SUCESSO
            ElseIf AFRAC_DadosUltimaReducao = 0 Then
                AFRAC_DadosUltimaReducao = 1
            End If


        Case Else
            AFRAC_DadosUltimaReducao = SUCESSO
            
    End Select

    Exit Function
   
Erro_AFRAC_DadosUltimaReducao:

    AFRAC_DadosUltimaReducao = 1
    
    Exit Function


End Function

Public Function AFRAC_Carrega_Dados_UltimaReducao(objUltimaReducao As ClassUltimaReducao, objLojaArqFisMestre As ClassLojaArqFisMestre, colLojaArqFisAnalitico As Collection) As Long

Dim X As New ClassECFConfig
Dim lErro As Long

Dim sReducao As String
Dim sTotalizadores As String
Dim iIndice As Integer
Dim dTot As Double
Dim objAliquota As ClassAliquotaICMS
Dim sData As String
Dim sIncideDescontoISS As String
Dim dVendaBruta As Double
Dim sAliquotas As String
Dim sHora As String
Dim sDataMovimento As String
Dim sSequencialECF As String
Dim objLojaArqFisAnalitico As ClassLojaArqFisAnalitico
Dim dTotFiscal As Double
Dim dTotFiscalICMS As Double
Dim dTotFiscalISS As Double
'Dim sCOOInicial As String
'Dim sCOOFinal As String

On Error GoTo Erro_AFRAC_Carrega_Dados_UltimaReducao
    
    Set objUltimaReducao = New ClassUltimaReducao
    Set objLojaArqFisMestre = New ClassLojaArqFisMestre
    Set colLojaArqFisAnalitico = New Collection
    
    Select Case X.giCodModeloECF

        Case 1 To 3
        
            lErro = AFRAC_DadosUltimaReducao(sReducao)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Dados Ultima Reducao Z")
            If lErro <> SUCESSO Then gError 204521

            lErro = AFRAC_DataHoraReducao(sData, sHora)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Data Hora Ultima Reducao Z")
            If lErro <> SUCESSO Then gError 210439

            If Len(Trim(sData)) > 0 Then

                sTotalizadores = Mid(sReducao, 334, 224)
                
                For iIndice = 0 To X.gcolAliquotasTotal.Count - 1
                    
                    dTot = StrParaDbl(Mid(sTotalizadores, (iIndice * 14) + 1, 14)) / 100
                    
                    Set objAliquota = X.gcolAliquotasTotal.Item(iIndice + 1)
                    
                    objAliquota.dValorTotalizadoLoja = dTot
                        
                    objUltimaReducao.colAliquotas.Add objAliquota
                    
                    If objAliquota.iISS = 0 Then
                        dTotFiscalICMS = dTotFiscalICMS + objAliquota.dValorTotalizadoLoja
                    Else
                        dTotFiscalISS = dTotFiscalISS + objAliquota.dValorTotalizadoLoja
                    End If
                    
                Next
                
                objUltimaReducao.lCOOReducaoZ = StrParaLong(Mid(sReducao, 11, 6))
                objUltimaReducao.dSubstTribICMS = StrParaDbl(Mid(sReducao, 589, 14)) / 100
                objUltimaReducao.dIsentoICMS = StrParaDbl(Mid(sReducao, 559, 14)) / 100
                objUltimaReducao.dNaoIncideICMS = StrParaDbl(Mid(sReducao, 574, 14)) / 100
                objUltimaReducao.dSubstTribISS = StrParaDbl(Mid(sReducao, 634, 14)) / 100
                objUltimaReducao.dIsentoISS = StrParaDbl(Mid(sReducao, 604, 14)) / 100
                objUltimaReducao.dNaoIncideISS = StrParaDbl(Mid(sReducao, 619, 14)) / 100
                
                objLojaArqFisMestre.lCRZ = StrParaLong(Mid(sReducao, 955, 4))
                objLojaArqFisMestre.lCRO = StrParaLong(Mid(sReducao, 950, 4))
                objLojaArqFisMestre.dGrandeTotal = StrParaDbl(Mid(sReducao, 10, 18)) / 100
                objLojaArqFisMestre.sNumSerieECF = X.gsNumSerie
                objLojaArqFisMestre.dtData = StrParaDate(Mid(sReducao, 1, 2) & "/" & Mid(sReducao, 3, 2) & "/" & Mid(sReducao, 5, 4))
                
                lErro = AFRAC_SequencialECF(sSequencialECF)
                If lErro <> SUCESSO Then
                    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Retorna Sequencial ECF")
                    gError 210808
                End If
    
                objLojaArqFisMestre.iNumEquip = StrParaInt(sSequencialECF)
                
                
                sTotalizadores = Mid(sReducao, 739, 392)
                
                dTot = 0
                
                For iIndice = 0 To 27
                    dTot = dTot + StrParaDbl(Mid(sTotalizadores, (iIndice * 14) + 1, 14)) / 100
                Next
                
                objUltimaReducao.dTotalNaoFiscal = dTot
                
                objUltimaReducao.dDescontoICMS = StrParaDbl(Mid(sReducao, 649, 14)) / 100
                objUltimaReducao.dDescontoISS = StrParaDbl(Mid(sReducao, 664, 14)) / 100
                objUltimaReducao.dDescontoNaoFiscal = StrParaDbl(Mid(sReducao, 1177, 14)) / 100
                
                objUltimaReducao.dAcrescimoICMS = StrParaDbl(Mid(sReducao, 679, 14)) / 100
                objUltimaReducao.dAcrescimoISS = StrParaDbl(Mid(sReducao, 694, 14)) / 100
                objUltimaReducao.dAcrescimoNaoFiscal = StrParaDbl(Mid(sReducao, 1192, 14)) / 100
                
                objUltimaReducao.dCancelamentoICMS = StrParaDbl(Mid(sReducao, 709, 14)) / 100
                objUltimaReducao.dCancelamentoISS = StrParaDbl(Mid(sReducao, 724, 14)) / 100
                objUltimaReducao.dCancelamentoNaoFiscal = StrParaDbl(Mid(sReducao, 1162, 14)) / 100
                
                sDataMovimento = StrParaDate(Mid(sReducao, 1273, 6))
                
                If sDataMovimento = "000000" Then
                
                    objUltimaReducao.dtDataMovimento = DATA_NULA
                    
                    objUltimaReducao.dtDataReducao = DATA_NULA
                    objUltimaReducao.sHoraReducao = "00:00:00"
                    
                Else
                
                    objUltimaReducao.dtDataMovimento = left(sDataMovimento, 2) & "/" & Mid(sDataMovimento, 3, 2) & "/" & right(sDataMovimento, 2)
                    
                    objUltimaReducao.dtDataReducao = left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 2)
                    objUltimaReducao.sHoraReducao = sHora
                            
                End If
                
                lErro = AFRAC_IncideDescontoISS(sIncideDescontoISS)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Incide Desconto ISS")
                If lErro <> SUCESSO Then gError 204519
                
                objUltimaReducao.sIncidenciaDescontoISS = sIncideDescontoISS
                
                lErro = AFRAC_VendaBrutaDiaria(dVendaBruta)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Venda Bruta Diaria")
                If lErro <> SUCESSO Then gError 204523
                
                objUltimaReducao.dVendaBruta = dVendaBruta
                objLojaArqFisMestre.dVendaBruta = dVendaBruta
                
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'isento
                objLojaArqFisAnalitico.sSituacaoTrib = "I"
                objLojaArqFisAnalitico.dTotalizador = objUltimaReducao.dIsentoICMS + objUltimaReducao.dIsentoISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'nao tributado
                objLojaArqFisAnalitico.sSituacaoTrib = "N"
                objLojaArqFisAnalitico.dTotalizador = objUltimaReducao.dNaoIncideICMS + objUltimaReducao.dNaoIncideISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'substituicao tributaria
                objLojaArqFisAnalitico.sSituacaoTrib = "F"
                objLojaArqFisAnalitico.dTotalizador = objUltimaReducao.dSubstTribICMS + objUltimaReducao.dSubstTribISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'tributado ICMS
                objLojaArqFisAnalitico.sSituacaoTrib = "T"
                objLojaArqFisAnalitico.dTotalizador = dTotFiscalICMS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'tributado ISS
                objLojaArqFisAnalitico.sSituacaoTrib = "S"
                objLojaArqFisAnalitico.dTotalizador = dTotFiscalISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
            
            End If
            
            
        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            lErro = AFRAC_DadosUltimaReducao(sReducao)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Dados Ultima Reducao Z")
            If lErro <> SUCESSO Then gError 204521

            lErro = AFRAC_DataHoraReducao(sData, sHora)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Dados Ultima Reducao Z")
            If lErro <> SUCESSO Then gError 210439

            If Len(Trim(sData)) > 0 Then


                sTotalizadores = Mid(sReducao, 335, 224)
                
                For iIndice = 0 To X.gcolAliquotasTotal.Count - 1
                    
                    dTot = StrParaDbl(Mid(sTotalizadores, (iIndice * 14) + 1, 14)) / 100
                    
                    Set objAliquota = X.gcolAliquotasTotal.Item(iIndice + 1)
                    
                    objAliquota.dValorTotalizadoLoja = dTot
                        
                    objUltimaReducao.colAliquotas.Add objAliquota
                    
                    If objAliquota.iISS = 0 Then
                        dTotFiscalICMS = dTotFiscalICMS + objAliquota.dValorTotalizadoLoja
                    Else
                        dTotFiscalISS = dTotFiscalISS + objAliquota.dValorTotalizadoLoja
                    End If
                    
                    
                Next
                
                objUltimaReducao.lCOOReducaoZ = StrParaLong(Mid(sReducao, 14, 6))
                objUltimaReducao.dSubstTribICMS = StrParaDbl(Mid(sReducao, 590, 14)) / 100
                objUltimaReducao.dIsentoICMS = StrParaDbl(Mid(sReducao, 560, 14)) / 100
                objUltimaReducao.dNaoIncideICMS = StrParaDbl(Mid(sReducao, 575, 14)) / 100
                objUltimaReducao.dSubstTribISS = StrParaDbl(Mid(sReducao, 635, 14)) / 100
                objUltimaReducao.dIsentoISS = StrParaDbl(Mid(sReducao, 605, 14)) / 100
                objUltimaReducao.dNaoIncideISS = StrParaDbl(Mid(sReducao, 620, 14)) / 100
                
                objLojaArqFisMestre.lCRZ = StrParaLong(Mid(sReducao, 9, 4))
                objLojaArqFisMestre.lCRO = StrParaLong(Mid(sReducao, 4, 4))
                objLojaArqFisMestre.dGrandeTotal = StrParaDbl(Mid(sReducao, 316, 18)) / 100
                objLojaArqFisMestre.sNumSerieECF = X.gsNumSerie
                
                lErro = AFRAC_SequencialECF(sSequencialECF)
                If lErro <> SUCESSO Then
                    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Retorna Sequencial ECF")
                    gError 210808
                End If
    
                objLojaArqFisMestre.iNumEquip = StrParaInt(sSequencialECF)
               
                sTotalizadores = Mid(sReducao, 740, 392)
                
                dTot = 0
                
                For iIndice = 0 To 27
                    dTot = dTot + StrParaDbl(Mid(sTotalizadores, (iIndice * 14) + 1, 14)) / 100
                Next
                
                objUltimaReducao.dTotalNaoFiscal = dTot
                
                objUltimaReducao.dDescontoICMS = StrParaDbl(Mid(sReducao, 650, 14)) / 100
                objUltimaReducao.dDescontoISS = StrParaDbl(Mid(sReducao, 665, 14)) / 100
                objUltimaReducao.dDescontoNaoFiscal = StrParaDbl(Mid(sReducao, 1163, 14)) / 100
                
                objUltimaReducao.dAcrescimoICMS = StrParaDbl(Mid(sReducao, 680, 14)) / 100
                objUltimaReducao.dAcrescimoISS = StrParaDbl(Mid(sReducao, 695, 14)) / 100
                objUltimaReducao.dAcrescimoNaoFiscal = StrParaDbl(Mid(sReducao, 1178, 14)) / 100
                
                objUltimaReducao.dCancelamentoICMS = StrParaDbl(Mid(sReducao, 710, 14)) / 100
                objUltimaReducao.dCancelamentoISS = StrParaDbl(Mid(sReducao, 725, 14)) / 100
                objUltimaReducao.dCancelamentoNaoFiscal = StrParaDbl(Mid(sReducao, 1193, 14)) / 100
                
                
                sDataMovimento = Mid(sReducao, 1273, 6)
                If sDataMovimento = "000000" Then
                
                    objUltimaReducao.dtDataMovimento = DATA_NULA
                Else
                
                    objUltimaReducao.dtDataMovimento = left(sDataMovimento, 2) & "/" & Mid(sDataMovimento, 3, 2) & "/" & right(sDataMovimento, 2)
                    
                End If
                
                objLojaArqFisMestre.dtData = objUltimaReducao.dtDataMovimento
                
                objUltimaReducao.dtDataReducao = left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 2)
                objUltimaReducao.sHoraReducao = sHora
                            
                
                lErro = AFRAC_IncideDescontoISS(sIncideDescontoISS)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Incide Desconto ISS")
                If lErro <> SUCESSO Then gError 204519
                
                objUltimaReducao.sIncidenciaDescontoISS = sIncideDescontoISS
                
                lErro = AFRAC_VendaBrutaDiaria(dVendaBruta)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Venda Bruta Diaria")
                If lErro <> SUCESSO Then gError 204523
                
                objUltimaReducao.dVendaBruta = dVendaBruta
                objLojaArqFisMestre.dVendaBruta = dVendaBruta
                
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'isento
                objLojaArqFisAnalitico.sSituacaoTrib = "I"
                objLojaArqFisAnalitico.dTotalizador = objUltimaReducao.dIsentoICMS + objUltimaReducao.dIsentoISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'nao tributado
                objLojaArqFisAnalitico.sSituacaoTrib = "N"
                objLojaArqFisAnalitico.dTotalizador = objUltimaReducao.dNaoIncideICMS + objUltimaReducao.dNaoIncideISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'substituicao tributaria
                objLojaArqFisAnalitico.sSituacaoTrib = "F"
                objLojaArqFisAnalitico.dTotalizador = objUltimaReducao.dSubstTribICMS + objUltimaReducao.dSubstTribISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'tributado ICMS
                objLojaArqFisAnalitico.sSituacaoTrib = "T"
                objLojaArqFisAnalitico.dTotalizador = dTotFiscalICMS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'tributado ISS
                objLojaArqFisAnalitico.sSituacaoTrib = "S"
                objLojaArqFisAnalitico.dTotalizador = dTotFiscalISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
            
            End If
            
            
        Case IMPRESSORA_DARUMA_FS600
        
            lErro = AFRAC_DadosUltimaReducao(sReducao)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Dados Ultima Reducao Z")
            If lErro <> SUCESSO Then gError 204522

            lErro = AFRAC_DataHoraReducao(sData, sHora)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Dados Ultima Reducao Z")
            If lErro <> SUCESSO Then gError 210440

            If Len(Trim(sData)) > 0 Then

                sAliquotas = Mid(sReducao, 869, 80)
                sTotalizadores = Mid(sReducao, 138, 224)
                
                For iIndice = 0 To 15
                
                    Set objAliquota = New ClassAliquotaICMS
                
                    If Mid(sAliquotas, iIndice * 5 + 1, 1) = "1" Then
                
                        objAliquota.iISS = 0
                        
                    ElseIf Mid(sAliquotas, iIndice * 5 + 1, 1) = "2" Then
                
                        objAliquota.iISS = 1
                        
                    Else
                    
                        Exit For
                        
                    End If
                
                    objAliquota.dAliquota = CDbl(Mid(sAliquotas, iIndice * 5 + 2, 4)) / 100
                
                    If objAliquota.dAliquota = 0 Then Exit For
                    
                    
                    objAliquota.dValorTotalizadoLoja = CDbl(Mid(sTotalizadores, iIndice * 14 + 1, 14)) / 100
                    
                    objUltimaReducao.colAliquotas.Add objAliquota
                    
                    If objAliquota.iISS = 0 Then
                        dTotFiscalICMS = dTotFiscalICMS + objAliquota.dValorTotalizadoLoja
                    Else
                        dTotFiscalISS = dTotFiscalISS + objAliquota.dValorTotalizadoLoja
                    End If
                
                
                Next
    
                
                objUltimaReducao.lCOOReducaoZ = StrParaLong(Mid(sReducao, 965, 6))
                objUltimaReducao.dSubstTribICMS = StrParaDbl(Mid(sReducao, 363, 14)) / 100
                objUltimaReducao.dIsentoICMS = StrParaDbl(Mid(sReducao, 393, 14)) / 100
                objUltimaReducao.dNaoIncideICMS = StrParaDbl(Mid(sReducao, 423, 14)) / 100
                objUltimaReducao.dSubstTribISS = StrParaDbl(Mid(sReducao, 453, 14)) / 100
                objUltimaReducao.dIsentoISS = StrParaDbl(Mid(sReducao, 483, 14)) / 100
                objUltimaReducao.dNaoIncideISS = StrParaDbl(Mid(sReducao, 513, 14)) / 100
                
                objLojaArqFisMestre.lCRZ = StrParaLong(Mid(sReducao, 955, 4))
                objLojaArqFisMestre.lCRO = StrParaLong(Mid(sReducao, 950, 4))
                objLojaArqFisMestre.dGrandeTotal = StrParaDbl(Mid(sReducao, 10, 18)) / 100
                objLojaArqFisMestre.sNumSerieECF = X.gsNumSerie
                objLojaArqFisMestre.dtData = StrParaDate(Mid(sReducao, 1, 2) & "/" & Mid(sReducao, 3, 2) & "/" & Mid(sReducao, 5, 4))
                    
    '            lErro = AFRAC_COOInicialFinal(sCOOInicial, sCOOFinal)
    '            If lErro <> SUCESSO Then
    '                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Retorna COO Inicial Final")
    '                gError 210919
    '            End If
    '
    '            objLojaArqFisMestre.lCOOIni = StrParaLong(sCOOInicial)
    '            objLojaArqFisMestre.lCOOFim = StrParaLong(sCOOFinal)
                    
                    
                lErro = AFRAC_SequencialECF(sSequencialECF)
                If lErro <> SUCESSO Then
                    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Retorna Sequencial ECF")
                    gError 210808
                End If
    
                
                objLojaArqFisMestre.iNumEquip = StrParaInt(sSequencialECF)
                
                sTotalizadores = Mid(sReducao, 543, 280)
                
                dTot = 0
                
                For iIndice = 0 To 19
                    dTot = dTot + StrParaDbl(Mid(sTotalizadores, (iIndice * 14) + 1, 14)) / 100
                Next
                
                objUltimaReducao.dTotalNaoFiscal = dTot
                
                objUltimaReducao.dDescontoICMS = StrParaDbl(Mid(sReducao, 48, 14)) / 100
                objUltimaReducao.dDescontoISS = StrParaDbl(Mid(sReducao, 63, 14)) / 100
                objUltimaReducao.dDescontoNaoFiscal = StrParaDbl(Mid(sReducao, 824, 14)) / 100
                
                objUltimaReducao.dAcrescimoICMS = StrParaDbl(Mid(sReducao, 108, 14)) / 100
                objUltimaReducao.dAcrescimoISS = StrParaDbl(Mid(sReducao, 123, 14)) / 100
                objUltimaReducao.dAcrescimoNaoFiscal = StrParaDbl(Mid(sReducao, 854, 14)) / 100
                
                objUltimaReducao.dCancelamentoICMS = StrParaDbl(Mid(sReducao, 78, 14)) / 100
                objUltimaReducao.dCancelamentoISS = StrParaDbl(Mid(sReducao, 93, 14)) / 100
                objUltimaReducao.dCancelamentoNaoFiscal = StrParaDbl(Mid(sReducao, 839, 14)) / 100
                
                sDataMovimento = Mid(sReducao, 1, 8)
                objUltimaReducao.dtDataMovimento = left(sDataMovimento, 2) & "/" & Mid(sDataMovimento, 3, 2) & "/" & right(sDataMovimento, 2)
                
                objUltimaReducao.dtDataReducao = left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 2)
                objUltimaReducao.sHoraReducao = sHora
                
                
                lErro = AFRAC_IncideDescontoISS(sIncideDescontoISS)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Incide Desconto ISS")
                If lErro <> SUCESSO Then gError 204520
                
                objUltimaReducao.sIncidenciaDescontoISS = sIncideDescontoISS
                
                lErro = AFRAC_VendaBrutaDiaria(dVendaBruta)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Venda Bruta Diaria")
                If lErro <> SUCESSO Then gError 204524
                
                objUltimaReducao.dVendaBruta = dVendaBruta
                objLojaArqFisMestre.dVendaBruta = dVendaBruta
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'isento
                objLojaArqFisAnalitico.sSituacaoTrib = "I"
                objLojaArqFisAnalitico.dTotalizador = objUltimaReducao.dIsentoICMS + objUltimaReducao.dIsentoISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'nao tributado
                objLojaArqFisAnalitico.sSituacaoTrib = "N"
                objLojaArqFisAnalitico.dTotalizador = objUltimaReducao.dNaoIncideICMS + objUltimaReducao.dNaoIncideISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'substituicao tributaria
                objLojaArqFisAnalitico.sSituacaoTrib = "F"
                objLojaArqFisAnalitico.dTotalizador = objUltimaReducao.dSubstTribICMS + objUltimaReducao.dSubstTribISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'tributado ICMS
                objLojaArqFisAnalitico.sSituacaoTrib = "T"
                objLojaArqFisAnalitico.dTotalizador = dTotFiscalICMS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
                
                Set objLojaArqFisAnalitico = New ClassLojaArqFisAnalitico
                
                'tributado ISS
                objLojaArqFisAnalitico.sSituacaoTrib = "S"
                objLojaArqFisAnalitico.dTotalizador = dTotFiscalISS
                
                colLojaArqFisAnalitico.Add objLojaArqFisAnalitico
            
            End If
            
    End Select

    AFRAC_Carrega_Dados_UltimaReducao = SUCESSO

    Exit Function

Erro_AFRAC_Carrega_Dados_UltimaReducao:

    AFRAC_Carrega_Dados_UltimaReducao = gErr
    
    
    Exit Function

End Function

Public Function AFRAC_AplMensagem1(sMensagem As String) As Integer

Dim X As New ClassECFConfig
Dim iErro As Integer

'Arredonda ou trunca os resultados de cálculo

On Error GoTo Erro_AFRAC_AplMensagem1

    AFRAC_AplMensagem1 = 1

    Select Case X.giCodModeloECF

        Case IMPRESSORA_DARUMA_FS600

'            AFRAC_AplMensagem1 = Daruma_Registry_AplMensagem1(sMensagem)
            AFRAC_AplMensagem1 = regAlterarValor_Daruma("ECF\MensagemApl1", sMensagem)

            If AFRAC_AplMensagem1 = 1 Then
                AFRAC_AplMensagem1 = SUCESSO
            ElseIf AFRAC_AplMensagem1 = 0 Then
                AFRAC_AplMensagem1 = 1
            End If
            
        Case Else
            AFRAC_AplMensagem1 = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_AplMensagem1:

    AFRAC_AplMensagem1 = 1
    
    Exit Function


End Function


Public Function AFRAC_RetornaMarcaECF(sMarca As String) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer

On Error GoTo Erro_AFRAC_RetornaMarcaECF

    AFRAC_RetornaMarcaECF = 1

    Select Case X.giCodModeloECF

        Case 1, 2
        
            sMarca = "Bematech"
            AFRAC_RetornaMarcaECF = SUCESSO
            
        Case 3
        
            sMarca = "Yanco"
            AFRAC_RetornaMarcaECF = SUCESSO
            
        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sMarca = Space(16)
            
            sModelo = Space(21)
            
            sTipo = Space(8)

            AFRAC_RetornaMarcaECF = Bematech_FI_MarcaModeloTipoImpressoraMFD(sMarca, sModelo, sTipo)

            sMarca = StringZ(sMarca)

            If AFRAC_RetornaMarcaECF = 1 Then
                AFRAC_RetornaMarcaECF = SUCESSO
            ElseIf AFRAC_RetornaMarcaECF = 0 Then
                AFRAC_RetornaMarcaECF = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            sMarca = Space(20)
            
'            iStatus = Daruma_FIMFD_RetornaInformacao("80", sMarca)
            AFRAC_RetornaMarcaECF = rRetornarInformacao_ECF_Daruma("80", sMarca)
                
            If AFRAC_RetornaMarcaECF = 1 Then
                AFRAC_RetornaMarcaECF = SUCESSO
            ElseIf AFRAC_RetornaMarcaECF = 0 Then
                AFRAC_RetornaMarcaECF = 1
            End If
            

        Case Else
            AFRAC_RetornaMarcaECF = SUCESSO
                    
    End Select

    sMarca = Trim(sMarca)

    Exit Function
   
Erro_AFRAC_RetornaMarcaECF:

    AFRAC_RetornaMarcaECF = 1
    
    Exit Function

End Function

Public Function AFRAC_RetornaTipoECF(sTipo As String) As Integer

Dim sMarca As String
Dim sModelo As String
Dim iStatus As Integer
Dim X As New ClassECFConfig

On Error GoTo Erro_AFRAC_RetornaTipoECF

    AFRAC_RetornaTipoECF = 1

    Select Case X.giCodModeloECF

        Case 1, 2, 3
            sTipo = "ECF-IF"
            AFRAC_RetornaTipoECF = SUCESSO
        
        
        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            sMarca = Space(16)
            
            sModelo = Space(21)
            
            sTipo = Space(8)

            AFRAC_RetornaTipoECF = Bematech_FI_MarcaModeloTipoImpressoraMFD(sMarca, sModelo, sTipo)

            sTipo = StringZ(sTipo)

            If AFRAC_RetornaTipoECF = 1 Then
                AFRAC_RetornaTipoECF = SUCESSO
            ElseIf AFRAC_RetornaTipoECF = 0 Then
                AFRAC_RetornaTipoECF = 1
            End If
        
        Case IMPRESSORA_DARUMA_FS600
            
            sTipo = Space(7)
            
'            iStatus = Daruma_FIMFD_RetornaInformacao("79", sTipo)
            AFRAC_RetornaTipoECF = rRetornarInformacao_ECF_Daruma("79", sTipo)
                
            If AFRAC_RetornaTipoECF = 1 Then
                AFRAC_RetornaTipoECF = SUCESSO
            ElseIf AFRAC_RetornaTipoECF = 0 Then
                AFRAC_RetornaTipoECF = 1
            End If
            
        Case Else
            AFRAC_RetornaTipoECF = SUCESSO
                    
    End Select

    sTipo = Trim(sTipo)

    Exit Function
   
Erro_AFRAC_RetornaTipoECF:

    AFRAC_RetornaTipoECF = 1
    
    Exit Function


End Function

Public Function AFRAC_RetornaModeloECF(sModelo As String) As Integer

'Abre a porta serial de comunicação entre o micro e a impressora
Dim X As New ClassECFConfig
Dim iStatus As Integer

On Error GoTo Erro_AFRAC_RetornaModeloECF

    AFRAC_RetornaModeloECF = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            sModelo = Space(11)
            AFRAC_RetornaModeloECF = Bematech_FI_ModeloImpressora(sModelo)
    
            sModelo = StringZ(sModelo)
    
            If AFRAC_RetornaModeloECF = 1 Then
                AFRAC_RetornaModeloECF = SUCESSO
            ElseIf AFRAC_RetornaModeloECF = 0 Then
                AFRAC_RetornaModeloECF = 1
            End If
            
            
        Case IMPRESSORA_DARUMA_FS600
                
            sModelo = Space(20)
            
'            iStatus = Daruma_FIMFD_RetornaInformacao("81", sModelo)
            AFRAC_RetornaModeloECF = rRetornarInformacao_ECF_Daruma("81", sModelo)
                
            If AFRAC_RetornaModeloECF = 1 Then
                AFRAC_RetornaModeloECF = SUCESSO
            ElseIf AFRAC_RetornaModeloECF = 0 Then
                AFRAC_RetornaModeloECF = 1
            End If
                
        Case Else
            AFRAC_RetornaModeloECF = SUCESSO
                    
    End Select

    sModelo = Trim(sModelo)

    Exit Function
   
Erro_AFRAC_RetornaModeloECF:

    AFRAC_RetornaModeloECF = 1
    
    Exit Function

End Function

Public Function AFRAC_CriaAssinaturaRSA(sArquivo As String, sRegistroEAD As String) As Integer

Dim X As New ClassECFConfig
Dim iRetorno As Integer
Dim sChavePublica As String
Dim sExpoente As String

On Error GoTo Erro_AFRAC_CriaAssinaturaRSA

'    Select Case X.giCodModeloECF
'
'        Case 1, 2, 3, 5, IMPRESSORA_DARUMA_FS600

'        sChavePublica = Space(256)
'        sExpoente = Space(256)
'
'       iRetorno = rRSAChavePublica_ECF_Daruma("c:\daruma\assinatura\corporatorprivate.key", sChavePublica, sExpoente)
'
'    iRetorno = eRSAAssinarArquivo_ECF_Daruma(sArquivo, "c:\daruma\assinatura\corporatorprivate.key")
        
        
        
    sRegistroEAD = Space(256)

    iRetorno = generateEAD(sArquivo, X.gsChavePublica, X.gsChavePrivada, sRegistroEAD, 1)

    If iRetorno = 1 Then
        iRetorno = validateFile(sArquivo, X.gsChavePublica, X.gsChavePrivada)
    End If
    
'    End Select
            
    If iRetorno = 1 Then
        AFRAC_CriaAssinaturaRSA = SUCESSO
    Else
        AFRAC_CriaAssinaturaRSA = 1
    End If
        
            
    Exit Function
            
Erro_AFRAC_CriaAssinaturaRSA:

        AFRAC_CriaAssinaturaRSA = 1

End Function

Function AFRAC_EspelhoMFD(iTipoParam As Integer, sInicio As String, sFinal As String) As Long
'NAO ANALISADO PARA VERSAO PAFECF

Dim X As New ClassECFConfig
Dim sRetorno As String
Dim lTamanho As Long
Dim sUsuario As String
Dim iStatus As Integer
Dim iNumeroUsuario As Integer
Dim sRegistroEAD As String
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_AFRAC_EspelhoMFD

    AFRAC_EspelhoMFD = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, 9, IMPRESSORA_DARUMA_FS600, IMPRESSORA_BEMATECH_MP2100_TH_FI

            If X.giCodModeloECF = IMPRESSORA_DARUMA_FS600 Then
            
                If iTipoParam = LEITURA_DATAS Then
            
                    sInicio = Replace(sInicio, "/", "")
                    sFinal = Replace(sFinal, "/", "")
                    AFRAC_EspelhoMFD = rGerarEspelhoMFD_ECF_Daruma("3", sInicio, sFinal)
           '         DarumaFramework_Mostrar_Retorno_ECF (0)

                    'AFRAC_EspelhoMFD = Daruma_FIMFD_DownloadDaMFD(sInicio, sFinal)
                
                ElseIf iTipoParam = LEITURA_COO Then
                
'                    sInicio = "COO - " & sInicio
                    
'                    sFinal = "COO - " & sFinal
                
'                    AFRAC_EspelhoMFD = Daruma_FIMFD_DownloadDaMFD(sInicio, sFinal)
                    AFRAC_EspelhoMFD = rGerarEspelhoMFD_ECF_Daruma("2", sInicio, sFinal)
                
                End If
                
                
                
                If AFRAC_EspelhoMFD = 1 Then
                
                    lErro = CF_ECF("Copia_Arquivo_EspelhoMFD_Daruma")
                    If lErro <> SUCESSO Then gError 210797
                
                    AFRAC_EspelhoMFD = AFRAC_CriaAssinaturaRSA(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Espelho_MFD.txt", sRegistroEAD)
                    
                End If
                
'                Kill CurDir() & "\Retorno.txt"
                
'                Name "c:\retorno.txt" As CurDir() & "\Retorno.txt"
                
                If AFRAC_EspelhoMFD = SUCESSO Then
                
                
                    vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Espelho_MFD.txt")
                
                End If
                
            'Bematech matricial antigas
            ElseIf X.giCodModeloECF = 1 Or X.giCodModeloECF = 2 Or X.giCodModeloECF = 3 Or X.giCodModeloECF = 5 Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP4200_TH_FI Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP2100_TH_FI Then
                
                    MsgBox "FUNÇÃO NÃO SUPORTADA PELO MODELO DE ECF UTILIZADO"
                    
                    AFRAC_EspelhoMFD = SUCESSO
                
            'deveria ser um modelo termico da bematech
            ElseIf X.giCodModeloECF = 9 Then
            
                iStatus = AFRAC_NumeroUsuario(iNumeroUsuario)
                
                If iStatus = SUCESSO Then
                    sUsuario = iNumeroUsuario
                End If
            
                If iTipoParam = LEITURA_DATAS Then
            
                    AFRAC_EspelhoMFD = Bematech_FI_DownloadMFD("c:\download.mfd", "1", Format(sInicio, "ddmmyy"), Format(sFinal, "ddmmyy"), sUsuario)
                    
                    If AFRAC_EspelhoMFD = 1 Then
                        AFRAC_EspelhoMFD = Bematech_FI_FormatoDadosMFD("c:\download.mfd", "c:\retorno.txt", "0", "1", Format(sInicio, "ddmmyy"), Format(sFinal, "ddmmyy"), sUsuario)
                    End If
                    
                ElseIf iTipoParam = LEITURA_COO Then
                
                    AFRAC_EspelhoMFD = Bematech_FI_DownloadMFD("c:\download.mfd", "2", sInicio, sFinal, sUsuario)
                
                    If AFRAC_EspelhoMFD = 1 Then
                        AFRAC_EspelhoMFD = Bematech_FI_FormatoDadosMFD("c:\download.mfd", "c:\retorno.txt", "0", "2", sInicio, sFinal, sUsuario)
                    End If
                
                
                End If
                
                If AFRAC_EspelhoMFD = 1 Then
        
                    If Len(dir(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt")) > 0 Then Kill CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt"
        
                    Name "c:\retorno.txt" As CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt"
        
                    AFRAC_EspelhoMFD = AFRAC_CriaAssinaturaRSA(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt", sRegistroEAD)
        
        
                    If AFRAC_EspelhoMFD = SUCESSO Then
                    
                        vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Retorno.txt")
                        
                        AFRAC_EspelhoMFD = 1
                        
                    End If

                End If
                
                If AFRAC_EspelhoMFD = 1 Then
                    AFRAC_EspelhoMFD = SUCESSO
                ElseIf AFRAC_EspelhoMFD = 0 Then
                    AFRAC_EspelhoMFD = 1
                End If
                
            End If
                                
        Case Else
            AFRAC_EspelhoMFD = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_EspelhoMFD:

    AFRAC_EspelhoMFD = 1
    
    Exit Function

End Function

Function AFRAC_ArqMFD(iTipoParam As Integer, sInicio As String, sFinal As String) As Long

Dim X As New ClassECFConfig
Dim sRetorno As String
Dim lTamanho As Long
Dim sUsuario As String
Dim iStatus As Integer
Dim iNumeroUsuario As Integer
Dim sRegistroEAD As String
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataIni As Date
Dim dtDataFim As Date
Dim sDataIni As String
Dim sDataFim As String
Dim sArquivo As String
Dim i As Integer
Dim sArquivo1 As String

On Error GoTo Erro_AFRAC_ArqMFD

    AFRAC_ArqMFD = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, 9, IMPRESSORA_DARUMA_FS600, IMPRESSORA_BEMATECH_MP2100_TH_FI

            If X.giCodModeloECF = IMPRESSORA_DARUMA_FS600 Then
        
                sArquivo = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "arqmfd.mfd"
                If Len(dir(sArquivo)) > 0 Then Kill sArquivo
                
                sArquivo1 = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "arqmfd.txt"
                If Len(dir(sArquivo1)) > 0 Then Kill sArquivo1
                
                sArquivo = "arqmfd.mfd"
        
       
                If iTipoParam = LEITURA_DATAS Then
        
                    dtDataIni = CDate(sInicio)
                    dtDataFim = CDate(sFinal)
                    
                    sDataIni = Format(dtDataIni, "ddmmyyyy")
                    sDataFim = Format(dtDataFim, "ddmmyyyy")
                    
        
                    For i = 1 To 20
        
                        AFRAC_ArqMFD = rVerificarImpressoraLigada_ECF_Daruma()
                    
                        If AFRAC_ArqMFD <> 1 Then Exit For
        
        
                        ' AFRAC_ArqMFD = rGerarRelatorio_ECF_Daruma("MFD", "DATAM", sDataIni, sDataFim)
                        AFRAC_ArqMFD = rEfetuarDownloadMFD_ECF_Daruma("DATAM", sDataIni, sDataFim, sArquivo)
                         
                        If AFRAC_ArqMFD = 1 Then
                         
                            sArquivo = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "arqmfd.mfd"
                         
                            AFRAC_ArqMFD = AFRAC_CriaAssinaturaRSA1(sArquivo, sRegistroEAD)
                             
                            Open sArquivo1 For Append Lock Read Write As #1
                            
                            Print #1, "EAD" & sRegistroEAD
                            
                            Close #1
                             
                             
                            AFRAC_ArqMFD = SUCESSO
                             
                            vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, sArquivo)
                            vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, sArquivo1)
                            
                            Exit For
                             
                        End If
                         
                    Next
                    
                    
'                ElseIf iTipoParam = LEITURA_COO Then
'
'                    AFRAC_ArqMFD = rGerarRelatorio_ECF_Daruma("MFD", "COO", Format(sInicio, "000000"), Format(sFinal, "000000"))
'
'                    If AFRAC_ArqMFD = 1 Then
'
'                        AFRAC_ArqMFD = AFRAC_CriaAssinaturaRSA(CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MFD_COO.TXT", sRegistroEAD)
'
'                        If AFRAC_ArqMFD = SUCESSO Then
'
'                            vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "ATO_MFD_COO.TXT")
'
'                        End If
'
'                    End If
                    
                End If
            
            
            
            
            'Bematech matricial antigas
            ElseIf X.giCodModeloECF = 1 Or X.giCodModeloECF = 2 Or X.giCodModeloECF = 3 Or X.giCodModeloECF = 5 Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP4200_TH_FI Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP2100_TH_FI Then
                
                    MsgBox "FUNÇÃO NÃO SUPORTADA PELO MODELO DE ECF UTILIZADO"
                
                    AFRAC_ArqMFD = SUCESSO
                
            'deveria ser um modelo termico da bematech
            ElseIf X.giCodModeloECF = 5 Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP4200_TH_FI Or X.giCodModeloECF = 9 Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP2100_TH_FI Then
            
                iStatus = AFRAC_NumeroUsuario(iNumeroUsuario)
                
                If iStatus = SUCESSO Then
                    sUsuario = iNumeroUsuario
                End If
            
                
                sArquivo = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "arqmfd.mfd"
                If Len(dir(sArquivo)) > 0 Then Kill sArquivo
                
                'nao pode colocar path pois da erro
                sArquivo1 = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "arqmfd.txt"
                If Len(dir(sArquivo1)) > 0 Then Kill sArquivo1
            
                If iTipoParam = LEITURA_DATAS Then
            
                    AFRAC_ArqMFD = Bematech_FI_DownloadMFD(sArquivo, "1", Format(sInicio, "ddmmyy"), Format(sFinal, "ddmmyy"), sUsuario)
                     
                    
                    If AFRAC_ArqMFD = 1 Then
                    
                    
                            AFRAC_ArqMFD = AFRAC_CriaAssinaturaRSA1(sArquivo, sRegistroEAD)
                             
                            Open sArquivo1 For Append Lock Read Write As #1
                            
                            Print #1, "EAD" & sRegistroEAD
                            
                            Close #1
                        
                            AFRAC_ArqMFD = SUCESSO
                             
                            vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, sArquivo)
                            vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, sArquivo1)
                        
                        
                    End If
                        
                    
                ElseIf iTipoParam = LEITURA_COO Then
                
                    MsgBox "FUNÇÃO NÃO SUPORTADA PELO MODELO DE ECF UTILIZADO"
                    AFRAC_ArqMFD = SUCESSO
                
                End If
                
            End If
                                
        Case Else
            AFRAC_ArqMFD = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_ArqMFD:

    AFRAC_ArqMFD = 1
    
    Exit Function

End Function

Public Function AFRAC_DataInstalacaoSB(dtData As Date) As Integer


Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sDataUsuario As String
Dim sDataSWBasico As String
Dim sMFAdicional As String
Dim sData As String


On Error GoTo Erro_AFRAC_DataInstalacaoSB

    AFRAC_DataInstalacaoSB = 1
    
    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
        
        
        
'            sDataUsuario = Space(21)
'
'            sDataSWBasico = Space(21)
'
'            sMFAdicional = Space(2)
'
'            If InStr(UCase(X.gsNumSerie), "EMULADOR") = 0 Then
'                AFRAC_DataInstalacaoSB = Bematech_FI_DataHoraGravacaoUsuarioSWBasicoMFAdicional(sDataUsuario, sDataSWBasico, sMFAdicional)
'            Else
'                sDataSWBasico = DATA_NULA
'                AFRAC_DataInstalacaoSB = 1
'            End If
'
'
'            sDataSWBasico = StringZ(sDataSWBasico)
'
'            sData = left(sDataSWBasico, 10)
'
'            dtData = CDate(sData)
'
'            AFRAC_DataInstalacaoSB = 1
'
'            If AFRAC_DataInstalacaoSB = 1 Then
'                AFRAC_DataInstalacaoSB = SUCESSO
'            ElseIf AFRAC_DataInstalacaoSB = 0 Then
'                AFRAC_DataInstalacaoSB = 1
'            End If
            
            dtData = Date

            AFRAC_DataInstalacaoSB = SUCESSO
            
            
'            If X.gsHoraSWBasico = "" Then
'
'
'                sDataUsuario = Space(21)
'
'                sDataSWBasico = Space(21)
'
'                sMFAdicional = Space(2)
'
'                If InStr(UCase(X.gsNumSerie), "EMULADOR") = 0 Then
'                    AFRAC_DataInstalacaoSB = Bematech_FI_DataHoraGravacaoUsuarioSWBasicoMFAdicional(sDataUsuario, sDataSWBasico, sMFAdicional)
'                Else
'                    sDataSWBasico = DATA_NULA
'                    AFRAC_DataInstalacaoSB = 1
'                End If
'
'
'                sDataSWBasico = StringZ(sDataSWBasico)
'
'                sData = left(sDataSWBasico, 10)
'
'                dtData = CDate(sData)
'
'                X.gdtDataSWBasico = dtData
'
'            Else
'
'                dtData = X.gdtDataSWBasico
'
'            End If
'
'            AFRAC_DataInstalacaoSB = 1
'
'            If AFRAC_DataInstalacaoSB = 1 Then
'                AFRAC_DataInstalacaoSB = SUCESSO
'            ElseIf AFRAC_DataInstalacaoSB = 0 Then
'                AFRAC_DataInstalacaoSB = 1
'            End If
            
            
        Case IMPRESSORA_DARUMA_FS600
                
            sData = Space(14)
            
'            iStatus = Daruma_FIMFD_RetornaInformacao("85", sData)
            AFRAC_DataInstalacaoSB = rRetornarInformacao_ECF_Daruma("85", sData)
            
            sData = left(sData, 8)
                
            dtData = left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 4)
                
            If AFRAC_DataInstalacaoSB = 1 Then
                AFRAC_DataInstalacaoSB = SUCESSO
            ElseIf AFRAC_DataInstalacaoSB = 0 Then
                AFRAC_DataInstalacaoSB = 1
            End If
                
        Case IMPRESSORA_NAO_FISCAL, IMPRESSORA_SAT_2_5_15, IMPRESSORA_NFCE, IMPRESSORA_NFE
            dtData = Date
            AFRAC_DataInstalacaoSB = SUCESSO
        
        Case Else
            AFRAC_DataInstalacaoSB = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_DataInstalacaoSB:

    AFRAC_DataInstalacaoSB = SUCESSO
    
    Exit Function
    

End Function


Public Function AFRAC_HoraInstalacaoSB(sHora As String) As Integer

'Abre a porta serial de comunicação entre o micro e a impressora
Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sDataUsuario As String
Dim sDataSWBasico As String
Dim sMFAdicional As String
Dim sData As String

On Error GoTo Erro_AFRAC_HoraInstalacaoSB

    AFRAC_HoraInstalacaoSB = 1

    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
'            sDataUsuario = Space(21)
'
'            sDataSWBasico = Space(21)
'
'            sMFAdicional = Space(2)
'
'            If InStr(UCase(X.gsNumSerie), "EMULADOR") = 0 Then
'                AFRAC_HoraInstalacaoSB = Bematech_FI_DataHoraGravacaoUsuarioSWBasicoMFAdicional(sDataUsuario, sDataSWBasico, sMFAdicional)
'            Else
'                sDataSWBasico = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
'                AFRAC_HoraInstalacaoSB = 1
'            End If
'
'            sDataSWBasico = StringZ(sDataSWBasico)
'
'            sHora = Mid(sDataSWBasico, 12, 2) & Mid(sDataSWBasico, 15, 2) & Mid(sDataSWBasico, 18, 2)
'
'            AFRAC_HoraInstalacaoSB = 1
'
'
'            If AFRAC_HoraInstalacaoSB = 1 Then
'                AFRAC_HoraInstalacaoSB = SUCESSO
'            ElseIf AFRAC_HoraInstalacaoSB = 0 Then
'                AFRAC_HoraInstalacaoSB = 1
'            End If
            
            
            sHora = Format(Time, "hhmmss")
            AFRAC_HoraInstalacaoSB = SUCESSO
            
'            If X.gsHoraSWBasico = "" Then
'
'
'                sDataUsuario = Space(21)
'
'                sDataSWBasico = Space(21)
'
'                sMFAdicional = Space(2)
'
'                If InStr(UCase(X.gsNumSerie), "EMULADOR") = 0 Then
'                    AFRAC_HoraInstalacaoSB = Bematech_FI_DataHoraGravacaoUsuarioSWBasicoMFAdicional(sDataUsuario, sDataSWBasico, sMFAdicional)
'                Else
'                    sDataSWBasico = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss")
'                    AFRAC_HoraInstalacaoSB = 1
'                End If
'
'                sDataSWBasico = StringZ(sDataSWBasico)
'
'                sHora = Mid(sDataSWBasico, 12, 2) & Mid(sDataSWBasico, 15, 2) & Mid(sDataSWBasico, 18, 2)
'
'                X.gsHoraSWBasico = sHora
'
'            Else
'
'                sHora = X.gsHoraSWBasico
'
'            End If
'
'            AFRAC_HoraInstalacaoSB = 1
'
'
'            If AFRAC_HoraInstalacaoSB = 1 Then
'                AFRAC_HoraInstalacaoSB = SUCESSO
'            ElseIf AFRAC_HoraInstalacaoSB = 0 Then
'                AFRAC_HoraInstalacaoSB = 1
'            End If
            
            
            
        Case IMPRESSORA_DARUMA_FS600
                
            sData = Space(14)
            
'            iStatus = Daruma_FIMFD_RetornaInformacao("85", sData)
            AFRAC_HoraInstalacaoSB = rRetornarInformacao_ECF_Daruma("85", sData)
            
            sHora = Mid(sData, 9, 6)
                
            If AFRAC_HoraInstalacaoSB = 1 Then
                AFRAC_HoraInstalacaoSB = SUCESSO
            ElseIf AFRAC_HoraInstalacaoSB = 0 Then
                AFRAC_HoraInstalacaoSB = 1
            End If
                
        Case IMPRESSORA_NAO_FISCAL, IMPRESSORA_SAT_2_5_15, IMPRESSORA_NFCE, IMPRESSORA_NFE
            sHora = Format(Time, "hhmmss")
            AFRAC_HoraInstalacaoSB = SUCESSO

        Case Else
            AFRAC_HoraInstalacaoSB = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_HoraInstalacaoSB:

    AFRAC_HoraInstalacaoSB = 1
    
    Exit Function


End Function

Public Function AFRAC_SequencialECF(sSequencial As String) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sDataInicial As String
Dim sDataFinal As String
Dim sRegistro As String
Dim sRetorno As String


On Error GoTo Erro_AFRAC_SequencialECF

    AFRAC_SequencialECF = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_DARUMA_FS600, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
        
'            sDataFinal = Format(gdtDataAtual, "ddmmyyyy")
'            sDataInicial = Format(gdtDataAtual - 10, "ddmmyyyy")
'
'            iStatus = Bematech_FI_DadosSintegra(sDataFinal, sDataInicial)
'
'            Open "c:\retorno.txt" For Input As #1
'
'            If Not EOF(1) Then
'
'                'Busca o próximo registro do arquivo
'                Line Input #1, sRegistro
'
'                sSequencial = Mid(sRegistro, 29, 3)
'
'                Close #1
'
'            Else
'
'                iStatus = 0
'
'            End If

            AFRAC_SequencialECF = AFRAC_LerInformacaoImpressora("015", sSequencial)
            
            
'        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
'
'
'            sDataFinal = Format(gdtDataAtual, "ddmmyyyy")
'            sDataInicial = Format(gdtDataAtual - 10, "ddmmyyyy")
'
'            iStatus = Bematech_FI_DadosSintegra(sDataFinal, sDataInicial)
'
'            Open "c:\retorno.txt" For Input As #1
'
'            If Not EOF(1) Then
'
'                'Busca o próximo registro do arquivo
'                Line Input #1, sRegistro
'
'                sSequencial = Mid(sRegistro, 29, 3)
'
'                Close #1
'
'            Else
'
'                iStatus = 0
'
'            End If
'
'            If iStatus = 1 Then
'                AFRAC_SequencialECF = SUCESSO
'            Else
'                AFRAC_SequencialECF = 1
'            End If
            
            
'        Case IMPRESSORA_DARUMA_FS600
'
'
'            sSequencial = Space(3)
'
''            iStatus = Daruma_FIMFD_RetornaInformacao("107", sSequencial)
'            iStatus = rRetornarInformacao_ECF_Daruma("107", sSequencial)
'
'            If iStatus = 1 Then
'                AFRAC_SequencialECF = SUCESSO
'            Else
'                AFRAC_SequencialECF = 1
'            End If
                
        Case Else
            AFRAC_SequencialECF = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_SequencialECF:

    AFRAC_SequencialECF = 1
    
    Exit Function

End Function


Public Function AFRAC_NumeroUsuario(iNumeroUsuario As Integer) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String

On Error GoTo Erro_AFRAC_NumeroUsuario

    AFRAC_NumeroUsuario = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sRetorno = Space(5)
    
            AFRAC_NumeroUsuario = Bematech_FI_NumeroSubstituicoesProprietario(sRetorno)
    
            sRetorno = StringZ(sRetorno)
    
            iNumeroUsuario = StrParaInt(sRetorno)
        
            If AFRAC_NumeroUsuario = 1 Then
                AFRAC_NumeroUsuario = SUCESSO
            ElseIf AFRAC_NumeroUsuario = 0 Then
                AFRAC_NumeroUsuario = 1
            End If
            
            
        Case IMPRESSORA_DARUMA_FS600
                
            sRetorno = Space(2)
            
'            iStatus = Daruma_FIMFD_RetornaInformacao("94", sSubstituicoes)
            AFRAC_NumeroUsuario = rRetornarInformacao_ECF_Daruma("94", sRetorno)
            
            iNumeroUsuario = StrParaInt(sRetorno)
            
            If AFRAC_NumeroUsuario = 1 Then
                AFRAC_NumeroUsuario = SUCESSO
            ElseIf AFRAC_NumeroUsuario = 0 Then
                AFRAC_NumeroUsuario = 1
            End If
            
        Case Else
            AFRAC_NumeroUsuario = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_NumeroUsuario:

    AFRAC_NumeroUsuario = 1
    
    Exit Function

End Function


Public Function AFRAC_CCF(sCCF As String) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer

On Error GoTo Erro_AFRAC_CCF

    AFRAC_CCF = 1

    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sCCF = Space(7)
            
            AFRAC_CCF = Bematech_FI_ContadorCupomFiscalMFD(sCCF)
            
            sCCF = StringZ(sCCF)
            
            If AFRAC_CCF = 1 Then
                AFRAC_CCF = SUCESSO
            ElseIf AFRAC_CCF = 0 Then
                AFRAC_CCF = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            sCCF = Space(6)
            
'            iStatus = Daruma_FIMFD_RetornaInformacao("30", sCCF)
            AFRAC_CCF = rRetornarInformacao_ECF_Daruma("30", sCCF)
                
            If AFRAC_CCF = 1 Then
                AFRAC_CCF = SUCESSO
            ElseIf AFRAC_CCF = 0 Then
                AFRAC_CCF = 1
            End If
            
        Case Else
            AFRAC_CCF = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_CCF:

    AFRAC_CCF = 1
    
    Exit Function

End Function

Public Function AFRAC_DataHoraImpressora(sData As String, sHora As String) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer

On Error GoTo Erro_AFRAC_DataHoraImpressora

    AFRAC_DataHoraImpressora = 1

    Select Case X.giCodModeloECF

        Case 1, 2, 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sData = Space(7)
            sHora = Space(7)
            AFRAC_DataHoraImpressora = Bematech_FI_DataHoraImpressora(sData, sHora)
            
            sData = StringZ(sData)
            sHora = StringZ(sHora)
            
            If AFRAC_DataHoraImpressora = 1 Then
                AFRAC_DataHoraImpressora = SUCESSO
            ElseIf AFRAC_DataHoraImpressora = 0 Then
                AFRAC_DataHoraImpressora = 1
            End If

            
        Case IMPRESSORA_DARUMA_FS600
        
            sData = Space(8)
            sHora = Space(6)
            AFRAC_DataHoraImpressora = rDataHoraImpressora_ECF_Daruma(sData, sHora)
        
            If AFRAC_DataHoraImpressora = 1 Then
                sData = left(sData, 4) & right(sData, 2)
            End If

            If AFRAC_DataHoraImpressora = 1 Then
                AFRAC_DataHoraImpressora = SUCESSO
            ElseIf AFRAC_DataHoraImpressora = 0 Then
                AFRAC_DataHoraImpressora = 1
            End If

        Case IMPRESSORA_NAO_FISCAL, IMPRESSORA_SAT_2_5_15, IMPRESSORA_NFCE, IMPRESSORA_NFE
            sData = Format(Date, "ddmmyyyy")
            sHora = Format(Time, "hhmmss")
            AFRAC_DataHoraImpressora = SUCESSO
            
        Case Else
            AFRAC_DataHoraImpressora = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_DataHoraImpressora:

    AFRAC_DataHoraImpressora = 1
    
    Exit Function

End Function

Public Function AFRAC_CodigoNacionalIdentificacao(sIdentificacao As String) As Integer
'os cases existem pq tem q usar a tabela naocional de codigos de identificacao de ecf

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer
Dim sData As String
Dim sHora As String
Dim sVersaoSB As String
Dim sMarca As String
Dim lErro As Long

On Error GoTo Erro_AFRAC_CodigoNacionalIdentificacao

    AFRAC_CodigoNacionalIdentificacao = 1

    AFRAC_CodigoNacionalIdentificacao = AFRAC_RetornaModeloECF(sModelo)
    If AFRAC_CodigoNacionalIdentificacao = SUCESSO Then

        AFRAC_CodigoNacionalIdentificacao = AFRAC_RetornaMarcaECF(sMarca)
        If AFRAC_CodigoNacionalIdentificacao = SUCESSO Then
            
            AFRAC_CodigoNacionalIdentificacao = AFRAC_LerInformacaoImpressora("003", sVersaoSB)
            If AFRAC_CodigoNacionalIdentificacao = SUCESSO Then

                lErro = CF_ECF("CodNacId_Le", sMarca, sModelo, sVersaoSB, sIdentificacao)
                If lErro <> SUCESSO Then gError 214511


                AFRAC_CodigoNacionalIdentificacao = SUCESSO
                
            End If
        End If
    End If

'        Select Case X.giCodModeloECF
'
'            Case 1
'
'                sVersaoFirmware = Space(4)
'
'                iStatus = Bematech_FI_VersaoFirmware(sVersaoFirmware)
'
'                If iStatus = 1 Then
'
'                    Select Case sVersaoFirmware
'
'                        Case "0300"
'                            sIdentificacao = "030501"
'
'                        Case "0305"
'                            sIdentificacao = "030502"
'
'                        Case "0310"
'                            sIdentificacao = "030503"
'
'                        Case "0315"
'                            sIdentificacao = "030504"
'
'                        Case "0322"
'                            sIdentificacao = "030505"
'
'                        Case "0325"
'                            sIdentificacao = "030506"
'
'                        Case "0326"
'                            sIdentificacao = "030507"
'
'                        Case "0330"
'                            sIdentificacao = "030509"
'
'                        Case "0331"
'                            sIdentificacao = "030508"
'
'                    End Select
'                End If
'
'                If iStatus = 1 Then
'                    AFRAC_CodigoNacionalIdentificacao = SUCESSO
'                Else
'                    AFRAC_CodigoNacionalIdentificacao = 1
'                End If
'
'            Case 2
'
'                sVersaoFirmware = Space(4)
'
'                iStatus = Bematech_FI_VersaoFirmware(sVersaoFirmware)
'
'                If iStatus = 1 Then
'
'                    Select Case sVersaoFirmware
'
'                        Case "0300"
'                            sIdentificacao = "031301"
'
'                        Case "0305"
'                            sIdentificacao = "031302"
'
'                        Case "0310"
'                            sIdentificacao = "031303"
'
'                        Case "0321"
'                            sIdentificacao = "031304"
'
'                    End Select
'                End If
'
'                If iStatus = 1 Then
'                    AFRAC_CodigoNacionalIdentificacao = SUCESSO
'                Else
'                    AFRAC_CodigoNacionalIdentificacao = 1
'                End If
'
'
'            Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI
'
'                sVersaoFirmware = Space(6)
'
'                iStatus = Bematech_FI_VersaoFirmwareMFD(sVersaoFirmware)
'
'                If iStatus = 1 Then
'
'                    If UCase(sModelo) = "MP-4200 TH FI" Then
'
'                        Select Case sVersaoFirmware
'
'                            Case "010000"
'                                sIdentificacao = "032201"
'    '
'    '                        Case "010101"
'    '                            sIdentificacao = "031002"
'    '
'    '                        Case "010202"
'    '                            sIdentificacao = "031003"
'    '
'    '                        Case "010302"
'    '                            sIdentificacao = "031004"
'
'                        End Select
'                    End If
'                End If
'
'
'                If iStatus = 1 Then
'                    AFRAC_CodigoNacionalIdentificacao = SUCESSO
'                Else
'                    AFRAC_CodigoNacionalIdentificacao = 1
'                End If
'
'            Case IMPRESSORA_DARUMA_FS600
'
'
'                sVersaoFirmware = Space(6)
'
'    '            iStatus = Daruma_FIMFD_RetornaInformacao("83", sVersaoFirmware)
'                iStatus = rRetornarInformacao_ECF_Daruma("83", sVersaoFirmware)
'
'                If iStatus = 1 Then
'
'                    If UCase(modelo) = "FS600" Then
'
'                        Select Case sVersaoFirmware
'
'                            Case "010100"
'                                sIdentificacao = "080801"
'
'                            Case "010200"
'                                sIdentificacao = "080802"
'
'                            Case "010300"
'                                sIdentificacao = "080803"
'
'                            Case "010400"
'                                sIdentificacao = "080804"
'
'                            Case "010500"
'                                sIdentificacao = "080805"
'
'
'                        End Select
'
'                    End If
'
'                End If
'
'                If iStatus = 1 Then
'                    AFRAC_CodigoNacionalIdentificacao = SUCESSO
'                Else
'                    AFRAC_CodigoNacionalIdentificacao = 1
'                End If
'
'
'            Case Else
'                AFRAC_CodigoNacionalIdentificacao = SUCESSO
'
'        End Select
'
'    End If

    Exit Function
   
Erro_AFRAC_CodigoNacionalIdentificacao:

    AFRAC_CodigoNacionalIdentificacao = 1
    
    Exit Function

End Function

Public Function AFRAC_IncideDescontoISS(sIncideDescontoISS As String) As Integer

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String

On Error GoTo Erro_AFRAC_IncideDescontoISS

    AFRAC_IncideDescontoISS = 1

    Select Case X.giCodModeloECF

        Case 1, 2, 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sIncideDescontoISS = "N"
            iStatus = 1
            
            If iStatus = 1 Then
                AFRAC_IncideDescontoISS = SUCESSO
            Else
                AFRAC_IncideDescontoISS = 1
            End If

            
            
        Case IMPRESSORA_DARUMA_FS600
                
             sRetorno = Space(1)
             
'            iStatus = Daruma_FIMFD_RetornaInformacao("109", iRetorno)
            AFRAC_IncideDescontoISS = rRetornarInformacao_ECF_Daruma("109", sRetorno)
                
            If AFRAC_IncideDescontoISS = 1 Then
                If CInt(sRetorno) = 1 Then
                    sIncideDescontoISS = "S"
                Else
                    sIncideDescontoISS = "N"
                End If
            End If
            
            If AFRAC_IncideDescontoISS = 1 Then
                AFRAC_IncideDescontoISS = SUCESSO
            ElseIf AFRAC_IncideDescontoISS = 0 Then
                AFRAC_IncideDescontoISS = 1
            End If

        Case Else
            AFRAC_IncideDescontoISS = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_IncideDescontoISS:

    AFRAC_IncideDescontoISS = 1
    
    Exit Function

End Function

Public Function AFRAC_VendaBrutaDiaria(dVendaBruta As Double) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer
Dim sRetorno As String
Dim dGrandeTotal As Double
Dim dGrandeTotalInicial As Double
Dim dtData As Date
Dim sDataInicial As String
Dim sDataFinal As String
Dim lErro As Long
Dim sRegistro As String
Dim sRegistro1 As String
Dim sRetorno1 As String
Dim lTamanho As Long
Dim iPosIni As Integer
Dim sChar As String
Dim sArquivo As String
Dim iPosFim As Integer

On Error GoTo Erro_AFRAC_VendaBrutaDiaria

    AFRAC_VendaBrutaDiaria = 1

    Select Case X.giCodModeloECF


        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            sRetorno = Space(19)

            AFRAC_VendaBrutaDiaria = Bematech_FI_GrandeTotalUltimaReducaoMFD(sRetorno)

            sRetorno = StringZ(sRetorno)

            dGrandeTotal = StrParaDbl(sRetorno) / 100

            If AFRAC_VendaBrutaDiaria = 1 Then
                AFRAC_VendaBrutaDiaria = SUCESSO
            ElseIf AFRAC_VendaBrutaDiaria = 0 Then
                AFRAC_VendaBrutaDiaria = 1
            End If
           
           
           
'            lErro = AFRAC_RelatorioTipo60Mestre()
'            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Relatorio Tipo 60 Mestre")
'            If lErro <> SUCESSO Then gError 133664
'
'            lTamanho = 255
'            sRetorno1 = String(lTamanho, 0)
'
'            Call GetSystemDirectory(sRetorno1, lTamanho)
'
'            'Retira os espaços no final da string
'            sRetorno1 = StringZ(sRetorno1)
'
'            'se o diretorio nao for terminado por \  ===> acrescentar
'            If right(sRetorno1, 1) <> "\" Then sRetorno1 = sRetorno1 & "\"
'
'            lTamanho = 255
'            sRetorno = String(lTamanho, 0)
'
'            'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
'            Call GetPrivateProfileString("Sistema", "Path", CONSTANTE_ERRO, sRetorno, lTamanho, sRetorno1 & "BemaFI32.ini")
'
'            'Retira os espaços no final da string
'            sRetorno = StringZ(sRetorno)
'
'            'Se não encontrou
'            If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 133666
'
'            'se o diretorio nao for terminado por \  ===> acrescentar
'            If right(sRetorno, 1) <> "\" Then sRetorno = sRetorno & "\"
'
'            sArquivo = sRetorno & "RETORNO.TXT"
'
'            Open sArquivo For Input As #1
'
'            Do While Not EOF(1)
'
'                'Busca o próximo registro do arquivo
'                Line Input #1, sRegistro
'                Line Input #1, sRegistro1
'
'                iPosFim = Len(sRegistro) + 1
'                iPosIni = iPosFim
'                sChar = Mid(sRegistro, iPosIni - 1, 1)
'                Do While sChar <> " "
'                    iPosIni = iPosIni - 1
'                    sChar = Mid(sRegistro, iPosIni - 1, 1)
'                Loop
'
'                If left(sRegistro, 11) = "Venda Bruta" Then
'                    dVendaBruta = StrParaDbl(Mid(sRegistro, iPosIni, iPosFim - iPosIni))
'                End If
'
'            Loop
'
'            Close #1
'
'            AFRAC_VendaBrutaDiaria = SUCESSO
           
        Case IMPRESSORA_DARUMA_FS600
                
            sRetorno = Space(1200)
                
'            iStatus = Daruma_FIMFD_RetornaInformacao("140", sRetorno)

            AFRAC_VendaBrutaDiaria = rRetornarInformacao_ECF_Daruma("140", sRetorno)
            
            If AFRAC_VendaBrutaDiaria = 1 Then
                dGrandeTotal = StrParaDbl(Mid(sRetorno, 9, 18)) / 100
                dGrandeTotalInicial = StrParaDbl(Mid(sRetorno, 27, 18)) / 100
                
                dVendaBruta = dGrandeTotal - dGrandeTotalInicial
            End If
    
            If AFRAC_VendaBrutaDiaria = 1 Then
                AFRAC_VendaBrutaDiaria = SUCESSO
            ElseIf AFRAC_VendaBrutaDiaria = 0 Then
                AFRAC_VendaBrutaDiaria = 1
            End If
            
        Case Else
            AFRAC_VendaBrutaDiaria = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_VendaBrutaDiaria:

    AFRAC_VendaBrutaDiaria = 1
    
'    Close #1
    
    Exit Function


End Function

Public Function AFRAC_EmitirLeituraMemoriaFiscal1(sInicio As String, sFinal As String) As Integer

Dim X As New ClassECFConfig

'Emite Leitura da Memória Fiscal
'é para imprimir a primeira leitura fiscal do mes
'na Daruma isto é feito automaticamente, entao nao precisa

On Error GoTo Erro_AFRAC_EmitirLeituraMemoriaFiscal1

    AFRAC_EmitirLeituraMemoriaFiscal1 = 1

    Select Case X.giCodModeloECF


        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

'            'Bematech matricial antigas
'            If X.giCodModeloECF = 1 Or X.giCodModeloECF = 2 Or X.giCodModeloECF = 3 Then
'
'                AFRAC_EmitirLeituraMemoriaFiscal1 = Bematech_FI_LeituraMemoriaFiscalData(sInicio, sFinal)
'
'            'Bematech matricial nova e termicas
'            ElseIf X.giCodModeloECF = 5 Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP4200_TH_FI Then
'
'                AFRAC_EmitirLeituraMemoriaFiscal1 = Bematech_FI_LeituraMemoriaFiscalDataMFD(sInicio, sFinal, "c")
'
'            End If
            
            If AFRAC_EmitirLeituraMemoriaFiscal1 = 1 Then
                AFRAC_EmitirLeituraMemoriaFiscal1 = SUCESSO
            ElseIf AFRAC_EmitirLeituraMemoriaFiscal1 = 0 Then
                AFRAC_EmitirLeituraMemoriaFiscal1 = 1
            End If

        Case IMPRESSORA_DARUMA_FS600

            
'                Call Daruma_Registry_MFD_LeituraMFCompleta("1")
'
'                AFRAC_EmitirLeituraMemoriaFiscal1 = Daruma_FI_LeituraMemoriaFiscalData(sInicio, sFinal)
                
            If AFRAC_EmitirLeituraMemoriaFiscal1 = 1 Then
                AFRAC_EmitirLeituraMemoriaFiscal1 = SUCESSO
            ElseIf AFRAC_EmitirLeituraMemoriaFiscal1 = 0 Then
                AFRAC_EmitirLeituraMemoriaFiscal1 = 1
            End If

        Case Else
            AFRAC_EmitirLeituraMemoriaFiscal1 = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_EmitirLeituraMemoriaFiscal1:

    AFRAC_EmitirLeituraMemoriaFiscal1 = 1
    
    Exit Function

End Function

Public Function AFRAC_GNF(sGNF As String) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer

On Error GoTo Erro_AFRAC_GNF

    AFRAC_GNF = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sGNF = Space(7)
            
            AFRAC_GNF = Bematech_FI_NumeroOperacoesNaoFiscais(sGNF)
            
            sGNF = StringZ(sGNF)
            
            If AFRAC_GNF = 1 Then
                AFRAC_GNF = SUCESSO
            ElseIf AFRAC_GNF = 0 Then
                AFRAC_GNF = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            sGNF = Space(6)
            
'            iStatus = Daruma_FIMFD_RetornaInformacao("28", sGNF)
            AFRAC_GNF = rRetornarInformacao_ECF_Daruma("28", sGNF)
                
            If AFRAC_GNF = 1 Then
                AFRAC_GNF = SUCESSO
            ElseIf AFRAC_GNF = 0 Then
                AFRAC_GNF = 1
            End If
            

        Case Else
            AFRAC_GNF = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_GNF:

    AFRAC_GNF = 1
    
    Exit Function

End Function

Public Function AFRAC_GRG(sGRG As String) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer

On Error GoTo Erro_AFRAC_GRG

    AFRAC_GRG = 1

    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sGRG = Space(7)
            
            AFRAC_GRG = Bematech_FI_ContadorRelatoriosGerenciaisMFD(sGRG)
            
            sGRG = StringZ(sGRG)
            
            If AFRAC_GRG = 1 Then
                AFRAC_GRG = SUCESSO
            ElseIf AFRAC_GRG = 0 Then
                AFRAC_GRG = 1
            End If

            
            
        Case IMPRESSORA_DARUMA_FS600
        
            sGRG = Space(6)
            
'            iStatus = Daruma_FIMFD_RetornaInformacao("33", sGRG)
            AFRAC_GRG = rRetornarInformacao_ECF_Daruma("33", sGRG)
                
            
            If AFRAC_GRG = 1 Then
                AFRAC_GRG = SUCESSO
            ElseIf AFRAC_GRG = 0 Then
                AFRAC_GRG = 1
            End If
            
        Case Else
            AFRAC_GRG = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_GRG:

    AFRAC_GRG = 1
    
    Exit Function

End Function

Public Function AFRAC_DataHoraUltimoDocumento(sData As String, sHora As String) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer
Dim sDataHora As String
Dim dtData As Date

On Error GoTo Erro_AFRAC_DataHoraUltimoDocumento

    AFRAC_DataHoraUltimoDocumento = 1

    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sDataHora = Space(13)
            AFRAC_DataHoraUltimoDocumento = Bematech_FI_DataHoraUltimoDocumentoMFD(sDataHora)
            
            
            If AFRAC_DataHoraUltimoDocumento = 1 Then
                sDataHora = StringZ(sDataHora)
                
                sData = left(sDataHora, 2) & "/" & Mid(sDataHora, 3, 2) & "/" & Mid(sDataHora, 5, 2)
                dtData = StrParaDate(sData)
                sData = Format(dtData, "ddmmyyyy")
                sHora = right(sDataHora, 6)
                AFRAC_DataHoraUltimoDocumento = SUCESSO
                
            ElseIf AFRAC_DataHoraUltimoDocumento = 0 Then
                AFRAC_DataHoraUltimoDocumento = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            sDataHora = Space(14)
'                    iStatus = Daruma_FIMFD_RetornaInformacao("72", sDataHora)
            AFRAC_DataHoraUltimoDocumento = rRetornarInformacao_ECF_Daruma("72", sDataHora)
            
                
            If AFRAC_DataHoraUltimoDocumento = 1 Then
                sData = left(sDataHora, 8)
                sHora = right(sDataHora, 6)
                AFRAC_DataHoraUltimoDocumento = SUCESSO
            ElseIf AFRAC_DataHoraUltimoDocumento = 0 Then
                AFRAC_DataHoraUltimoDocumento = 1
            End If
            
        Case IMPRESSORA_NAO_FISCAL, IMPRESSORA_SAT_2_5_15, IMPRESSORA_NFCE, IMPRESSORA_NFE
            sData = Format(Date, "ddmmyyyy")
            sHora = Format(Time, "hhmmss")
            AFRAC_DataHoraUltimoDocumento = SUCESSO
        
        Case Else
            AFRAC_DataHoraUltimoDocumento = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_DataHoraUltimoDocumento:

    AFRAC_DataHoraUltimoDocumento = 1
    
    Exit Function


End Function

Public Function AFRAC_MD5(sArquivo As String, sMD5 As String) As Integer

Dim X As New ClassECFConfig
Dim iRetorno As Integer
Dim sRegistroEAD As String
Dim sArq As String

Dim sMD5Ascii As String

On Error GoTo Erro_AFRAC_MD5

    AFRAC_MD5 = 1

    Select Case X.giCodModeloECF

        Case IMPRESSORA_DARUMA_FS600
        
            sMD5Ascii = Space(32)
        
            sMD5 = Space(32)
            sRegistroEAD = Space(256)
            
'            iRetorno = md5FromFile(sArquivo, sMD5)


            iRetorno = rCalcularMD5_ECF_Daruma(sArquivo, sMD5, sMD5Ascii)

            sMD5 = left(sMD5, 32)
                
'            iRetorno = 1
                
            If iRetorno = 1 Then
                AFRAC_MD5 = SUCESSO
            Else
                AFRAC_MD5 = 1
            End If
            
        Case 1, 2, 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sMD5Ascii = Space(32)
        
        
            sMD5 = Space(33)
            sRegistroEAD = Space(256)
            
            sArq = dir(sArquivo)
            
            If sArq <> "REQUISITO_XI.txt" Then
                sArquivo = left(sArquivo, Len(sArquivo) - Len(sArq)) & "md5\" & sArq
            End If
            
            iRetorno = md5FromFile(sArquivo, sMD5)

            sMD5 = left(sMD5, 32)
                
            If giDebug = 1 Then MsgBox ("AFRAC_MD5 iRetorno = " & CStr(iRetorno) & "  MD5 = " & sMD5)
                
'            iRetorno = 1
                
            If iRetorno = 1 Then
                AFRAC_MD5 = SUCESSO
            Else
                AFRAC_MD5 = 1
            End If
            
    Case Else
        AFRAC_MD5 = SUCESSO
                
    End Select
    
    Exit Function
            
Erro_AFRAC_MD5:

        AFRAC_MD5 = 1
        
        Exit Function

End Function

Public Function AFRAC_CRZ(sCRZ As String) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer
Dim sRetorno As String

On Error GoTo Erro_AFRAC_CRZ

    AFRAC_CRZ = 1

    Select Case X.giCodModeloECF

        Case 1, 2, 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
'            sRetorno = Space(1300)
                
'            iStatus = Bematech_FI_DadosUltimaReducaoMFD(sRetorno)
'
'            If iStatus = 1 Then
'                sCRZ = Mid(sRetorno, 7, 4)
'            End If
            
            sCRZ = Space(5)
    
            AFRAC_CRZ = Bematech_FI_NumeroReducoes(sCRZ)
            
            sCRZ = StringZ(sCRZ)
            
            If AFRAC_CRZ = 1 Then
                AFRAC_CRZ = SUCESSO
            ElseIf AFRAC_CRZ = 0 Then
                AFRAC_CRZ = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            sCRZ = Space(4)
            
'            iStatus = Daruma_FI_RetornaCRZ(sCRZ)
            AFRAC_CRZ = rRetornarInformacao_ECF_Daruma("24", sCRZ)
                
            If AFRAC_CRZ = 1 Then
                AFRAC_CRZ = SUCESSO
            ElseIf AFRAC_CRZ = 0 Then
                AFRAC_CRZ = 1
            End If

    Case Else
        AFRAC_CRZ = SUCESSO
                
    End Select
    
    Exit Function
            
Erro_AFRAC_CRZ:

    AFRAC_CRZ = 1
    
    Exit Function

End Function

Public Function AFRAC_CRO(sCRO As String) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer
Dim sRetorno As String

On Error GoTo Erro_AFRAC_CRO

    AFRAC_CRO = 1

    Select Case X.giCodModeloECF

        Case 1, 2, 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
'            sRetorno = Space(1300)
'
'            iStatus = Bematech_FI_DadosUltimaReducaoMFD(sRetorno)
'
'            If iStatus = 1 Then
'                sCRO = Mid(sRetorno, 3, 4)
'            End If
            
            sCRO = Space(5)

            AFRAC_CRO = Bematech_FI_NumeroIntervencoes(sCRO)
            
            sCRO = StringZ(sCRO)
            
            
            If AFRAC_CRO = 1 Then
                AFRAC_CRO = SUCESSO
            ElseIf AFRAC_CRO = 0 Then
                AFRAC_CRO = 1
            End If

            
        Case IMPRESSORA_DARUMA_FS600
        
            sCRO = Space(4)
            
'            iStatus = Daruma_FI_RetornaCRO(sCRO)
            AFRAC_CRO = rRetornarInformacao_ECF_Daruma("23", sCRO)
                
            If AFRAC_CRO = 1 Then
                AFRAC_CRO = SUCESSO
            ElseIf AFRAC_CRO = 0 Then
                AFRAC_CRO = 1
            End If

        Case Else
            AFRAC_CRO = SUCESSO
                
    End Select
    
    Exit Function
            
Erro_AFRAC_CRO:

    AFRAC_CRO = 1
    
    Exit Function

End Function


'        lErro = AFRAC_EmitirLeituraMemoriaFiscal(LEITURA_COMPLETA, "", "", LEITURA_COMPLETA, iArquivo)

'Public Function AFRAC_VendaPeriodo_SPED(ByVal dtDataDe As Date, ByVal dtDataAte As Date) As Integer
'
'Dim x As New ClassECFConfig
'Dim sModelo As String
'Dim sTipo As String
'Dim iStatus As Integer
'Dim sDataDe As String
'Dim sDataAte As String
'
'    Select Case x.giCodModeloECF
'
'        Case 1, 2, 3, 5
'
''            sRetorno = Space(1300)
''
''            iStatus = Bematech_FI_DadosUltimaReducaoMFD(sRetorno)
''
''            If iStatus = 1 Then
''                sCRO = Mid(sRetorno, 3, 4)
''            End If
'
'
'        Case IMPRESSORA_DARUMA_FS600
'
'            sDataDe = Format(dtDataDe, "ddmmyyyy")
'            sDataAte = Format(dtDataAte, "ddmmyyyy")
'
'            iStatus = rGerarRelatorio_ECF_Daruma("SPED", "DATAM", sDataDe, sDataAte)
'
'
'   End Select
'
'
'    If iStatus = 1 Then
'        AFRAC_VendaPeriodo_SPED = SUCESSO
'    Else
'        AFRAC_VendaPeriodo_SPED = 1
'    End If
'
'
'End Function
'
'Public Function AFRAC_VendaPeriodo_SINTEGRA(ByVal dtDataDe As Date, ByVal dtDataAte As Date) As Integer
'
'Dim x As New ClassECFConfig
'Dim sModelo As String
'Dim sTipo As String
'Dim iStatus As Integer
'Dim sDataDe As String
'Dim sDataAte As String
'
'    Select Case x.giCodModeloECF
'
'        Case 1, 2, 3, 5
'
''            sRetorno = Space(1300)
''
''            iStatus = Bematech_FI_DadosUltimaReducaoMFD(sRetorno)
''
''            If iStatus = 1 Then
''                sCRO = Mid(sRetorno, 3, 4)
''            End If
'
'
'        Case IMPRESSORA_DARUMA_FS600
'
'            sDataDe = Format(dtDataDe, "ddmmyyyy")
'            sDataAte = Format(dtDataAte, "ddmmyyyy")
'
'            iStatus = rGerarRelatorio_ECF_Daruma("SINTEGRA", "DATAM", sDataDe, sDataAte)
'
'
'   End Select
'
'
'    If iStatus = 1 Then
'        AFRAC_VendaPeriodo_SINTEGRA = SUCESSO
'    Else
'        AFRAC_VendaPeriodo_SINTEGRA = 1
'    End If
'
'
'End Function
'
'

Public Function AFRAC_DataHoraReducao(sData As String, sHora As String) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer
Dim sRetorno As String

On Error GoTo Erro_AFRAC_DataHoraReducao

    AFRAC_DataHoraReducao = 1

    Select Case X.giCodModeloECF

        Case 1, 2, 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sData = Space(7)
            sHora = Space(7)
            AFRAC_DataHoraReducao = Bematech_FI_DataHoraReducao(sData, sHora)
            
            If AFRAC_DataHoraReducao = 1 Then
                sData = Trim(StringZ(sData))
                sHora = Trim(StringZ(sHora))
                
                If sHora = "" Then sHora = Format(Time, "hhmmss")
            
                AFRAC_DataHoraReducao = SUCESSO
            ElseIf AFRAC_DataHoraReducao = 0 Then
                AFRAC_DataHoraReducao = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
'                    sData = Space(6)
'                    sHora = Space(6)
'                    iStatus = rRetornarInformacao_ECF_Daruma("154", sRetorno)
                
            
            sRetorno = String(14, 0)
            AFRAC_DataHoraReducao = rRetornarInformacao_ECF_Daruma("154", sRetorno)
            
            If AFRAC_DataHoraReducao = 1 Then
                sData = Mid(sRetorno, 1, 4) & Mid(sRetorno, 7, 2)
                sHora = right(sRetorno, 6)
                AFRAC_DataHoraReducao = SUCESSO
            ElseIf AFRAC_DataHoraReducao = 0 Then
                AFRAC_DataHoraReducao = 1
            End If

            
        Case Else
            AFRAC_DataHoraReducao = SUCESSO
                
    End Select
    
    Exit Function
            
Erro_AFRAC_DataHoraReducao:

    AFRAC_DataHoraReducao = 1
    
    Exit Function

End Function

Public Function AFRAC_ProgramarAliquotas() As Integer

Dim lErro As Long
Dim sRetornoICMS As String
Dim sRetornoISS As String
Dim objAliquota As ClassAliquotaICMS
Dim iPos As Integer
Dim sAliquota As String
Dim bAchou As Boolean
Dim sTrib As String
Dim iInicio As Integer
Dim iFim As Integer
Dim sTributacao As String
Dim iIndice As Integer
Dim sIndiceAliquotasISS As String
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iInicioIndice As Integer
Dim iFimIndice As Integer
Dim sRetorno As String
Dim X As New ClassECFConfig
Dim sTribISS As String

On Error GoTo Erro_AFRAC_ProgramarAliquotas


    Select Case X.giCodModeloECF

        Case 1, 2, 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI


            iIndice = 1
        
            sRetornoICMS = String(79, 0)
            lErro = AFRAC_LerAliquotas(sRetornoICMS)
            If lErro <> SUCESSO Then
                Call CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Ler Aliquotas")
                gError 112414
            End If
            
            sRetornoISS = String(79, 0)
            lErro = AFRAC_VerificaAliquotasIss(sRetornoISS)
            If lErro <> SUCESSO Then
                Call CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Verifica Aliquotas Iss")
                gError 133672
            End If
            
            sIndiceAliquotasISS = String(48, 0)
                            
            lErro = AFRAC_VerificaIndiceAliquotasIss(sIndiceAliquotasISS)
            If lErro <> SUCESSO Then
                Call CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Verifica Indice Aliquotas Iss")
                gError 133674
            End If
            
            For Each objAliquota In X.gobjLojaECF.colAliquotaICMS
                
                iPos = InStr(1, objAliquota.dAliquota, ",")
                sAliquota = objAliquota.dAliquota
                If objAliquota.iISS = ALIQUOTA_ISS Then
                    sTributacao = "S" & FormataCpoNum(iIndice, 2) & Formata_Campo(ALINHAMENTO_DIREITA, 4, "0", (Mid(sAliquota, iPos + 1, Len(sAliquota) - iPos)))
                Else
                    sTributacao = "T" & FormataCpoNum(iIndice, 2) & Formata_Campo(ALINHAMENTO_DIREITA, 4, "0", (Mid(sAliquota, iPos + 1, Len(sAliquota) - iPos)))
                End If
                
                
                If objAliquota.iISS = ALIQUOTA_ICMS Then
                
                    bAchou = False
                    iInicio = 1
                    iIndice1 = 0
                    iFim = InStr(iInicio, sRetornoICMS, ",")
                    
                    If iFim = 0 And Len(sRetornoICMS) > 0 Then
                        iFim = Len(sRetornoICMS) + 1
                    End If
                    
                    Do While iInicio < Len(sRetornoICMS)
                        iIndice1 = iIndice1 + 1
                        sTrib = Mid(sRetornoICMS, iInicio, iFim - iInicio)
                        If sTrib = Mid(sTributacao, 4, 4) Then
                            
                            bAchou = True
                            
                            'a rotina a seguir exclui as aliquotas de ISS da pesquisa
                            iInicioIndice = 1
                            
                            iFimIndice = InStr(iInicioIndice, sRetornoISS, ",")
                                                
                            If iFimIndice = 0 Then
                                iFimIndice = Len(sRetornoISS) + 1
                            End If
                                                                    
                            Do While iInicioIndice < Len(sRetornoISS)
                                
                                sTribISS = Mid(sRetornoISS, iInicioIndice, iFimIndice - iInicioIndice)
                                If sTrib = sTribISS Then
                                    bAchou = False
                                    Exit Do
                                End If
                                
                                iInicioIndice = iFimIndice + 1
                            
                                iFimIndice = InStr(iInicioIndice, sRetornoISS, ",")
                                If iFimIndice = 0 Then iFimIndice = Len(sRetornoISS) + 1
                            
                            Loop
                            
                            If bAchou = True Then Exit Do
                            
                        End If
                        
                        iInicio = iFim + 1
                        iFim = InStr(iInicio, sRetornoICMS, ",")
                        If iFim = 0 Then iFim = Len(sRetornoICMS) + 1
                    Loop
                
                ElseIf objAliquota.iISS = ALIQUOTA_ISS Then
                
                    bAchou = False
                    iInicio = 1
                    iFim = InStr(iInicio, sRetornoISS, ",")
                    
                    If iFim = 0 Then
                        iFim = Len(sRetornoISS) + 1
                    End If
                    
                    Do While iInicio < Len(sRetornoISS)
                        sTrib = Mid(sRetornoISS, iInicio, iFim - iInicio)
                        If sTrib = Mid(sTributacao, 4, 4) Then
                            bAchou = True
                            Exit Do
                        End If
                        iInicio = iFim + 1
                        iFim = InStr(iInicio, sRetornoISS, ",")
                        If iFim = 0 Then iFim = Len(sRetornoISS) + 1
                    Loop
                
                End If
                
                If Not (bAchou) Then
                
                
                    lErro = AFRAC_ProgramarTributacao(sTributacao)
                    If lErro <> SUCESSO Then
                        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Programar Tributação")
                        gError 99872
                    End If
                    
                    If giDebug = 1 Then MsgBox ("Programou Aliquota " & sTributacao)
                    
                End If
                
                
                X.gcolAliquotasTotal.Add objAliquota
            
            Next


        Case IMPRESSORA_DARUMA_FS600
        
            iIndice = 1
        
            sRetornoICMS = String(79, 0)
            lErro = AFRAC_LerAliquotas(sRetorno)
            If lErro <> SUCESSO Then
                Call CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Ler Aliquotas")
                gError 210793
            End If
                
            
            For Each objAliquota In X.gobjLojaECF.colAliquotaICMS
                
                If objAliquota.iISS = ALIQUOTA_ISS Then
                    sTributacao = "S" & Format(objAliquota.dAliquota * 10000, "0000")
                Else
                    sTributacao = "T" & Format(objAliquota.dAliquota * 10000, "0000")
                End If
                
                bAchou = False
                iInicio = 1
                iFim = InStr(iInicio, sRetorno, ";")
                
                If iFim = 0 And Mid(sRetorno, iInicio, 1) <> Chr(0) Then
                    iFim = InStr(iInicio, sRetorno, Chr(0))
                End If
                
                Do While iFim <> 0
                    sTrib = Mid(sRetorno, iInicio, iFim - iInicio)
                    If sTrib = sTributacao Then
                        bAchou = True
                        Exit Do
                    End If
                    iInicio = iFim + 1
                    iFim = InStr(iInicio, sRetorno, ";")
                    If iFim = 0 And Mid(sRetorno, iInicio, 1) <> Chr(0) Then iFim = InStr(iInicio, sRetorno, Chr(0))
                Loop
                
                If Not (bAchou) Then
                    lErro = AFRAC_ProgramarTributacao(sTributacao)
                    If lErro <> SUCESSO Then
                        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Programar Tributação")
                        gError 210792
                    End If
                End If
                
                X.gcolAliquotasTotal.Add objAliquota
            
            Next
            
        Case IMPRESSORA_NAO_FISCAL
            
            For Each objAliquota In X.gobjLojaECF.colAliquotaICMS
                
                X.gcolAliquotasTotal.Add objAliquota
            
            Next
            
            
            
   End Select
   
    AFRAC_ProgramarAliquotas = SUCESSO
    
    Exit Function

Erro_AFRAC_ProgramarAliquotas:
    
    AFRAC_ProgramarAliquotas = 1
    
'    Select Case gErr
'
'        Case 99872, 112414, 133672, 133674, 210792, 210793
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165214)
'
'    End Select
    
    Exit Function

End Function

Public Function AFRAC_CDC(sCDC As String) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer

On Error GoTo Erro_AFRAC_CDC

    AFRAC_CDC = 1

    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sCDC = Space(5)
            
            AFRAC_CDC = Bematech_FI_ContadorComprovantesCreditoMFD(sCDC)

            sCDC = StringZ(sCDC)
    
            If AFRAC_CDC = 1 Then
                AFRAC_CDC = SUCESSO
            ElseIf AFRAC_CDC = 0 Then
                AFRAC_CDC = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            sCDC = Space(6)
            
'            iStatus = Daruma_FIMFD_RetornaInformacao("30", sCCF)
            AFRAC_CDC = rRetornarInformacao_ECF_Daruma("45", sCDC)
                
            If AFRAC_CDC = 1 Then
                AFRAC_CDC = SUCESSO
            ElseIf AFRAC_CDC = 0 Then
                AFRAC_CDC = 1
            End If
            
        Case Else
            AFRAC_CDC = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_CDC:

    AFRAC_CDC = 1
    
    Exit Function

End Function

Public Function AFRAC_COOInicialFinal(sCOOInicial As String, sCOOFinal As String) As Integer

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer
Dim iStatus1 As Integer


On Error GoTo Erro_AFRAC_COOInicialFinal

    AFRAC_COOInicialFinal = 1

    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            sCOOInicial = Space(7)
            sCOOFinal = Space(7)
            
            AFRAC_COOInicialFinal = Bematech_FI_InicioFimCOOsMFD(sCOOInicial, sCOOFinal)

            sCOOInicial = StringZ(sCOOInicial)
            sCOOFinal = StringZ(sCOOFinal)
            
            If AFRAC_COOInicialFinal = 1 Then
                AFRAC_COOInicialFinal = SUCESSO
            ElseIf AFRAC_COOInicialFinal = 0 Then
                AFRAC_COOInicialFinal = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            sCOOInicial = Space(6)
            sCOOFinal = Space(6)
            
'            iStatus = Daruma_FIMFD_RetornaInformacao("28", sGNF)
            iStatus = rRetornarInformacao_ECF_Daruma("27", sCOOInicial)
            iStatus1 = rRetornarInformacao_ECF_Daruma("26", sCOOFinal)
                
            If iStatus = 1 And iStatus1 = 1 Then
                AFRAC_COOInicialFinal = SUCESSO
            Else
                AFRAC_COOInicialFinal = 1
            End If
            
        Case Else
            AFRAC_COOInicialFinal = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_COOInicialFinal:

    AFRAC_COOInicialFinal = 1
    
    Exit Function

End Function

Public Function AFRAC_FecharNaoFiscalNaoVinculado() As Integer
'NAO USADO PELO PAF-ECF

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim lErro As Long
'fecha Vinculado

On Error GoTo Erro_AFRAC_FecharNaoFiscalNaoVinculado

    AFRAC_FecharNaoFiscalNaoVinculado = 1
    
    Select Case X.giCodModeloECF

        Case IMPRESSORA_DARUMA_FS600

            AFRAC_FecharNaoFiscalNaoVinculado = iCNFEncerrarPadrao_ECF_Daruma()
            
            If AFRAC_FecharNaoFiscalNaoVinculado = 1 Then
                AFRAC_FecharNaoFiscalNaoVinculado = SUCESSO
                lErro = CF_ECF("Grava_R06", "CN")
                If lErro <> SUCESSO Then gError 210807
            ElseIf AFRAC_FecharNaoFiscalNaoVinculado = 0 Then
                AFRAC_FecharNaoFiscalNaoVinculado = 1
            End If
            
        Case Else
                AFRAC_FecharNaoFiscalNaoVinculado = SUCESSO
            
    End Select

   Exit Function
   
Erro_AFRAC_FecharNaoFiscalNaoVinculado:

    AFRAC_FecharNaoFiscalNaoVinculado = 1
    
    Exit Function

End Function

'Public Function AFRAC_RZPendenteDiaAnterior(iRZPendente As Integer) As Integer
''indica se algum dia anterior tem reducao z pendente
'
'Dim X As New ClassECFConfig
'Dim iStatus As Integer
'Dim sRetorno As String
'Dim lErro As Long
'Dim lRZPendente As Long
''fecha Vinculado
'
'On Error GoTo Erro_AFRAC_RZPendenteDiaAnterior
'
'    Select Case X.giCodModeloECF
'
'        Case IMPRESSORA_DARUMA_FS600
'
'            AFRAC_RZPendenteDiaAnterior = rConsultaStatusImpressoraInt_ECF_Daruma(6, lRZPendente)
'
'            If AFRAC_RZPendenteDiaAnterior = 1 Then
'                iRZPendente = lRZPendente
'                AFRAC_RZPendenteDiaAnterior = SUCESSO
'            Else
'                AFRAC_RZPendenteDiaAnterior = 1
'            End If
'
'        Case Else
'                AFRAC_RZPendenteDiaAnterior = SUCESSO
'
'    End Select
'
'   Exit Function
'
'Erro_AFRAC_RZPendenteDiaAnterior:
'
'    AFRAC_RZPendenteDiaAnterior = 1
'
'    Exit Function
'
'End Function


Public Function AFRAC_RZPendente(iRZPendente As Integer) As Integer
'indica se esta em jornada fiscal

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim lErro As Long
Dim lRZPendente As Long
Dim sStatusRZ As String
'fecha Vinculado
Dim iZPendenteAnt As Integer, iZPendenteAtual As Integer, iBloqueadoRZ As Integer
Dim sDataProxRZ As String
Dim dtUltimaReducaoECF As Date
Dim sData As String
Dim dtDataUltimoMovimento As Date

On Error GoTo Erro_AFRAC_RZPendente

    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            If InStr(UCase(X.gsNumSerie), "EMULADOR") = 0 Then

'                iStatus = AFRAC_DataMovimentoProximaReducao(sDataProxRZ)
'
'                If iStatus = SUCESSO Then
            
                lErro = AFRAC_DataReducao(dtUltimaReducaoECF)
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Data Ultima Reducao Z")
                If lErro <> SUCESSO Then gError 214612
            
                lErro = AFRAC_DataMovimentoProximaReducao(sData)
                If lErro <> SUCESSO Then
                    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DataMovimentoProximaReducao")
                    If lErro <> SUCESSO Then gError 214613
                End If
                
                dtDataUltimoMovimento = StrParaDate(left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 4))
            
                If dtDataUltimoMovimento > dtUltimaReducaoECF Then
                    iRZPendente = 1
                Else
                    iRZPendente = 0
                End If
                    
                AFRAC_RZPendente = SUCESSO
            
'                dtDataProxRZ = CDate(Mid(sDataProxRZ, 1, 2) & "/" & Mid(sDataProxRZ, 3, 2) & "/" & Mid(sDataProxRZ, 5, 4))
'
'
'                sStatusRZ = Space(1)
'
'                AFRAC_RZPendente = Bematech_FI_VerificaZPendente(sStatusRZ)
'
'                If AFRAC_RZPendente = 1 Then
'                    iRZPendente = StrParaInt(sStatusRZ)
'                    AFRAC_RZPendente = SUCESSO
'                ElseIf AFRAC_RZPendente = 0 Then
'                    AFRAC_RZPendente = 1
'                End If

            Else
            
                AFRAC_RZPendente = Bematech_Emulador_ZPendente(iZPendenteAnt, iZPendenteAtual, iBloqueadoRZ)
            
                If AFRAC_RZPendente = 1 Then
                    iRZPendente = IIf(iZPendenteAnt <> 0 Or iZPendenteAtual <> 0, 1, 0)
                    AFRAC_RZPendente = SUCESSO
                ElseIf AFRAC_RZPendente = 0 Then
                    AFRAC_RZPendente = 1
                End If
            
            
            End If

        Case IMPRESSORA_DARUMA_FS600

'            sRetorno = Space(1)

'            AFRAC_RZPendente = rVerificarReducaoZ_ECF_Daruma(sRetorno)
            AFRAC_RZPendente = rConsultaStatusImpressoraInt_ECF_Daruma(4, lRZPendente)
            
            If AFRAC_RZPendente = 1 Then
                iRZPendente = lRZPendente
                AFRAC_RZPendente = SUCESSO
            ElseIf AFRAC_RZPendente = 0 Then
                AFRAC_RZPendente = 1
            End If
            
            
        Case Else
                AFRAC_RZPendente = SUCESSO
            
    End Select

   Exit Function
   
Erro_AFRAC_RZPendente:

    AFRAC_RZPendente = 1
    
    Exit Function

End Function

Public Function AFRAC_DataReducao(dtData As Date) As Integer
'retorna a data dos movimentos a que se refere a ultima reducao. Não é a data em que foi feita a redução.

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer
Dim sRetorno As String
Dim sReducao As String
Dim sDataMovimento As String
Dim lErro As Long

On Error GoTo Erro_AFRAC_DataReducao

    AFRAC_DataReducao = 1

    Select Case X.giCodModeloECF

        Case 1, 2, 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
        
            sRetorno = Space(7)
'            lErro = Bematech_FI_DataMovimentoUltimaReducaoMFD(sRetorno)
            lErro = AFRAC_DadosUltimaReducao(sReducao)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Data Ultima Reducao MFD")
            If lErro <> SUCESSO Then gError 214014
            
            sDataMovimento = Mid(sReducao, 1273, 6)
            
            If StrParaDbl(sDataMovimento) > 0 Then
                dtData = left(sDataMovimento, 2) & "/" & Mid(sDataMovimento, 3, 2) & "/" & right(sDataMovimento, 2)
            Else
                dtData = DATA_NULA
                
            End If
        
            AFRAC_DataReducao = SUCESSO
        
        
'            sData = Space(6)
'            sHora = Space(6)
'            iStatus = Bematech_FI_DataHoraReducao(sData, sHora)
'
'            If iStatus = 1 Then
'                AFRAC_DataHoraReducao = SUCESSO
'            Else
'                AFRAC_DataHoraReducao = 1
'            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
                
            
            lErro = AFRAC_DadosUltimaReducao(sReducao)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Dados Ultima Reducao Z")
            If lErro <> SUCESSO Then gError 214014
            
            sDataMovimento = Mid(sReducao, 1, 8)
            If StrParaDbl(sDataMovimento) > 0 Then
                dtData = left(sDataMovimento, 2) & "/" & Mid(sDataMovimento, 3, 2) & "/" & right(sDataMovimento, 2)
            Else
                dtData = DATA_NULA
                
            End If
            
            AFRAC_DataReducao = SUCESSO

            
        Case Else
            AFRAC_DataReducao = SUCESSO
                
    End Select
    
    Exit Function
            
Erro_AFRAC_DataReducao:

    AFRAC_DataReducao = 1
    
    Exit Function

End Function


Public Function AFRAC_DataMovimentoProximaReducao(sDataMovimento As String) As Integer
'retorna a data dos movimentos a que se refere a proxima reducao.

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim iStatus As Integer
Dim sRetorno As String
Dim sReducao As String
Dim lErro As Long
Dim dtData As Date
Dim dtDataUltRZ As Date

On Error GoTo Erro_AFRAC_DataMovimentoProximaReducao

    AFRAC_DataMovimentoProximaReducao = 1

    Select Case X.giCodModeloECF

        Case 1, 2, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            iStatus = AFRAC_DataReducao(dtDataUltRZ)

            If iStatus = SUCESSO Then
            
                'le a data movimento do ultimo movimento do ecf
                iStatus = AFRAC_LerInformacaoImpressora("019", sRetorno)
        
                If iStatus = SUCESSO Then
        
                    If sRetorno = "000000" Then
                        dtData = DATA_NULA
                    Else
                        dtData = StrParaDate(Mid(sRetorno, 1, 2) & "/" & Mid(sRetorno, 3, 2) & "/" & Mid(sRetorno, 5, 2))
                    End If
                
                    If dtData > dtDataUltRZ Then
                        sDataMovimento = Format(dtData, "ddmmyyyy")
                    Else
                        sDataMovimento = Format(DATA_NULA, "ddmmyyyy")
                    End If
            
                    AFRAC_DataMovimentoProximaReducao = SUCESSO
            
            
                End If
                
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            
            iStatus = AFRAC_DataReducao(dtDataUltRZ)

            If iStatus = SUCESSO Then
            
                'le a data movimento do ultimo movimento do ecf
                iStatus = AFRAC_LerInformacaoImpressora("019", sRetorno)
        
                If iStatus = SUCESSO Then
        
                    If sRetorno = "000000" Then
                        dtData = DATA_NULA
                    Else
                        dtData = StrParaDate(Mid(sRetorno, 1, 2) & "/" & Mid(sRetorno, 3, 2) & "/" & Mid(sRetorno, 5, 2))
                    End If
                
                
                    If dtData > dtDataUltRZ Then
                        sDataMovimento = Format(dtData, "ddmmyyyy")
                    Else
                        sDataMovimento = Format(DATA_NULA, "ddmmyyyy")
                    End If
            
                    AFRAC_DataMovimentoProximaReducao = SUCESSO
            
                End If
            End If
            
'            sDataMovimento = Space(8)
'            iStatus = rRetornarInformacao_ECF_Daruma("69", sDataMovimento)
'            If iStatus = 1 Then
'                dtData = StrParaDate(left(sDataMovimento, 2) & "/" & Mid(sDataMovimento, 3, 2) & "/" & Mid(sDataMovimento, 7, 2))
'                dtData = dtData - 1
'                sDataMovimento = Format(dtData, "ddmmyyyy")
'                AFRAC_DataMovimentoProximaReducao = SUCESSO
'            Else
'                sDataMovimento = DATA_NULA
'                AFRAC_DataMovimentoProximaReducao = 1
'            End If
            
            
            
            
        Case Else
            AFRAC_DataMovimentoProximaReducao = SUCESSO
                
    End Select
    
    Exit Function
            
Erro_AFRAC_DataMovimentoProximaReducao:

    AFRAC_DataMovimentoProximaReducao = 1
    
    Select Case gErr
    
        Case 214014
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 214015)

    End Select
    
    Exit Function

End Function

Function AFRAC_ArqMF() As Long

Dim X As New ClassECFConfig
Dim sRetorno As String
Dim lTamanho As Long
Dim sUsuario As String
Dim iStatus As Integer
Dim iNumeroUsuario As Integer
Dim sRegistroEAD As String
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataIni As Date
Dim dtDataFim As Date
Dim sDataIni As String
Dim sDataFim As String
Dim sArquivo As String
Dim i As Integer
Dim sArquivo1 As String

On Error GoTo Erro_AFRAC_ArqMF

    AFRAC_ArqMF = 1

    Select Case X.giCodModeloECF

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, 9, IMPRESSORA_BEMATECH_MP2100_TH_FI

            'Bematech matricial antigas
            If X.giCodModeloECF = 1 Or X.giCodModeloECF = 2 Or X.giCodModeloECF = 3 Then
                
                MsgBox "FUNÇÃO NÃO SUPORTADA PELO MODELO DE ECF UTILIZADO"
            
                AFRAC_ArqMF = SUCESSO
                
            'deveria ser um modelo termico da bematech
            ElseIf X.giCodModeloECF = 5 Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP4200_TH_FI Or X.giCodModeloECF = IMPRESSORA_BEMATECH_MP2100_TH_FI Then
            
'                iStatus = AFRAC_NumeroUsuario(iNumeroUsuario)
'
'                If iStatus = 1 Then
'                    sUsuario = iNumeroUsuario
'                End If
'
'
'                If iTipoParam = LEITURA_DATAS Then
'
'                    AFRAC_ArqMF = Bematech_FI_DownloadMFD("c:\retorno.mfd", "1", Format(sInicio, "ddmmyy"), Format(sFinal, "ddmmyy"), sUsuario)
'
'                    If AFRAC_ArqMF = 1 Then
'
''                        AFRAC_ArqMF = BemaGeraRegistrosTipoE("c:\retorno.mfd", "c:\retorno.txt", Format(sInicio, "ddmmyyyy"), Format(sFinal, "ddmmyyyy"), gsNomeEmpresa, gsEndereco, "", "2", "", "", "", "", "", "", "", "", "", "", "", "", "")
'
'                    End If
'
'                    If AFRAC_ArqMFD = 1 Then
'
'                        AFRAC_ArqMF = AFRAC_CriaAssinaturaRSA("c:\Retorno.txt", sRegistroEAD)
'
'                    End If
'
'
'                ElseIf iTipoParam = LEITURA_COO Then
'
'                    MsgBox "FUNÇÃO NÃO SUPORTADA PELO MODELO DE ECF UTILIZADO"
'                    AFRAC_ArqMF = SUCESSO
'
'                End If
'
            
                'nao pode colocar path pois da erro
                sArquivo = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Arqmf.mf"
                If Len(dir(sArquivo)) > 0 Then Kill sArquivo
                
                sArquivo = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Arqmf.txt"
                If Len(dir(sArquivo)) > 0 Then Kill sArquivo
                
        
                    AFRAC_ArqMF = Bematech_FI_DownloadMF(sArquivo)
    
                    If AFRAC_ArqMF = 1 Then
                        
                        sArquivo = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Arqmf.mf"
                        
                        AFRAC_ArqMF = AFRAC_CriaAssinaturaRSA1(sArquivo, sRegistroEAD)
                        
                        sArquivo1 = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Arqmf.txt"
                        
                        Open sArquivo1 For Append Lock Read Write As #1
                        
                        Print #1, "EAD" & sRegistroEAD
                        
                        Close #1
                        
                        AFRAC_ArqMF = SUCESSO
                                
                        vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, sArquivo)
                        vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, sArquivo1)
                        
                    End If
                    
            
            
            
            End If
                                
        Case IMPRESSORA_DARUMA_FS600

            'nao pode colocar path pois da erro
            sArquivo = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Arqmf.mf"
            If Len(dir(sArquivo)) > 0 Then Kill sArquivo
            
            sArquivo = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Arqmf.txt"
            If Len(dir(sArquivo)) > 0 Then Kill sArquivo
            
            
            sArquivo = "Arqmf.mf"
    
            For i = 1 To 20
    
                AFRAC_ArqMF = rVerificarImpressoraLigada_ECF_Daruma()
                
                If AFRAC_ArqMF <> 1 Then Exit For
    
                AFRAC_ArqMF = rEfetuarDownloadMF_ECF_Daruma(sArquivo)
                
                If AFRAC_ArqMF = 1 Then
                    
                    sArquivo = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Arqmf.mf"
                    
                    AFRAC_ArqMF = AFRAC_CriaAssinaturaRSA1(sArquivo, sRegistroEAD)
                    
                    sArquivo1 = CurDir() & IIf(right(CurDir, 1) <> "\", "\", "") & "Arqmf.txt"
                    
                    Open sArquivo1 For Append Lock Read Write As #1
                    
                    Print #1, "EAD" & sRegistroEAD
                    
                    Close #1
                    
                    AFRAC_ArqMF = SUCESSO
                            
                    vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, sArquivo)
                    vbMsgRes = Rotina_AvisoECF(vbOKOnly, AVISO_ARQUIVO_GERADO, sArquivo1)
                    
                    Exit For
                
                End If
                
            Next
                                
        Case Else
            AFRAC_ArqMF = SUCESSO
                    
    End Select

    Exit Function
   
Erro_AFRAC_ArqMF:

    AFRAC_ArqMF = 1
    
    Exit Function

End Function

Public Function AFRAC_ProgramarRelGer(sIndice As String, sDescricao As String, bVinculado As Boolean) As Integer

'Programar as formas de pagamento

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim objTiposMeiosPagtos As ClassTMPLoja
Dim iIndice As Integer

On Error GoTo Erro_AFRAC_ProgramarRelGer

    AFRAC_ProgramarRelGer = 1
    
    
    iIndice = CInt(sIndice)
    
    Select Case iIndice
    
        Case RELGER_SANGRIA_CHEQUE
            sDescricao = RELGER_DESC_SANGRIA_CHEQUE
    
        Case RELGER_SANGRIA_TICKET
            sDescricao = RELGER_DESC_SANGRIA_TICKET
    
        Case RELGER_SANGRIA_CARTAO_CREDITO
            sDescricao = RELGER_DESC_SANGRIA_CARTAO_CREDITO
    
        Case RELGER_SANGRIA_OUTROS
            sDescricao = RELGER_DESC_SANGRIA_OUTROS
    
        Case RELGER_ORCAMENTO
            sDescricao = RELGER_DESC_ORCAMENTO
    
        Case RELGER_SANGRIASUPRIMENTO_DINHEIRO
            sDescricao = RELGER_DESC_SANGRIASUPRIMENTO_DINHEIRO
    
        Case RELGER_COMPROVANTE_TEF
            sDescricao = RELGER_DESC_COMPROVANTE_TEF
    
        Case RELGER_DAV
            sDescricao = RELGER_DESC_DAV
    
        Case RELGER_MEIOPAGAMENTO
            sDescricao = RELGER_DESC_MEIOPAGAMENTO
    
        Case RELGER_IDENTPAFECF
            sDescricao = RELGER_DESC_IDENTPAFECF
    
        Case RELGER_PARAMCONFIG
            sDescricao = RELGER_DESC_PARAMCONFIG
    
        Case RELGER_SAT
            sDescricao = RELGER_DESC_SAT
    
        Case RELGER_NFCE
            sDescricao = RELGER_DESC_NFCE
    
        Case RELGER_NFE
            sDescricao = RELGER_DESC_NFE
    
        Case RELGER_SAT_CANC
            sDescricao = RELGER_DESC_SAT_CANC
    
    End Select

    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            AFRAC_ProgramarRelGer = Bematech_FI_NomeiaRelatorioGerencialMFD(sIndice, sDescricao)
            
            
            If AFRAC_ProgramarRelGer = 1 Then
                AFRAC_ProgramarRelGer = SUCESSO
            ElseIf AFRAC_ProgramarRelGer = 0 Then
                AFRAC_ProgramarRelGer = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
'            iStatus = Daruma_FI_ProgramaFormasPagamento(sDescricao)
             AFRAC_ProgramarRelGer = confCadastrarPadrao_ECF_Daruma("RG", left(sDescricao, 15))

        
            If AFRAC_ProgramarRelGer = 1 Or AFRAC_ProgramarRelGer = -1 Then
                AFRAC_ProgramarRelGer = SUCESSO
            ElseIf AFRAC_ProgramarRelGer = 0 Then
                AFRAC_ProgramarRelGer = 1
            End If
        
        
        Case IMPRESSORA_ELGIN
            For Each objTiposMeiosPagtos In X.gcolTiposMeiosPagtos
                If CInt(sIndice) = objTiposMeiosPagtos.iTipo Then
                    iStatus = Elgin_ProgramaLegenda(Format(objTiposMeiosPagtos.iIndice, "00"), DesacentuaTexto(Format(left(sDescricao, 16), "!@@@@@@@@@@@@@@@@")))
                    
                    sRetorno = Space(50)
                    iStatus = Elgin_EsperaResposta(sRetorno)
                
                    Exit For
                End If
            Next
        
            
            If iStatus = SUCESSO Or iStatus = -20 Then
                AFRAC_ProgramarRelGer = SUCESSO
            Else
                AFRAC_ProgramarRelGer = 1
            End If
            
        Case Else
            AFRAC_ProgramarRelGer = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_ProgramarRelGer:

    AFRAC_ProgramarRelGer = 1
    
    Exit Function

End Function

Public Function AFRAC_CriaAssinaturaRSA1(sArquivo As String, sRegistroEAD As String) As Integer

Dim X As New ClassECFConfig
Dim iRetorno As Integer
Dim sChavePublica As String
Dim sExpoente As String

On Error GoTo Erro_AFRAC_CriaAssinaturaRSA1

'    Select Case X.giCodModeloECF
'
'        Case 1, 2, 3, 5, IMPRESSORA_DARUMA_FS600

'        sChavePublica = Space(256)
'        sExpoente = Space(256)
'
'       iRetorno = rRSAChavePublica_ECF_Daruma("c:\daruma\assinatura\corporatorprivate.key", sChavePublica, sExpoente)
'
'    iRetorno = eRSAAssinarArquivo_ECF_Daruma(sArquivo, "c:\daruma\assinatura\corporatorprivate.key")
        
        
        
    sRegistroEAD = Space(256)

    iRetorno = generateEAD(sArquivo, X.gsChavePublica, X.gsChavePrivada, sRegistroEAD, 0)

    If iRetorno = 1 Then
        iRetorno = validateFile(sArquivo, X.gsChavePublica, X.gsChavePrivada)
    End If
    
'    End Select
            
    If iRetorno = 1 Then
        AFRAC_CriaAssinaturaRSA1 = SUCESSO
    Else
        AFRAC_CriaAssinaturaRSA1 = 1
    End If
        
            
    Exit Function
            
Erro_AFRAC_CriaAssinaturaRSA1:

        AFRAC_CriaAssinaturaRSA1 = 1

End Function

Public Function AFRAC_ECFComProblema(bComProblema As Boolean) As Integer
'retorna a data dos movimentos a que se refere a proxima reducao.

Dim X As New ClassECFConfig
Dim sModelo As String
Dim sTipo As String
Dim sRetorno As String
Dim lErro As Long
Dim iStatus As Integer

On Error GoTo Erro_AFRAC_ECFComProblema

    bComProblema = False

    AFRAC_ECFComProblema = 1

    Select Case X.giCodModeloECF
            
        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            iStatus = Bematech_FI_VerificaImpressoraLigada()
            
            If iStatus <> 1 Then bComProblema = True
            
            AFRAC_ECFComProblema = SUCESSO
            
            
        Case IMPRESSORA_DARUMA_FS600
        
            iStatus = rVerificarImpressoraLigada_ECF_Daruma()
            
            If iStatus <> 1 Then bComProblema = True
                
            AFRAC_ECFComProblema = SUCESSO
            
        Case Else
            AFRAC_ECFComProblema = SUCESSO
                
    End Select
    
    Exit Function
            
Erro_AFRAC_ECFComProblema:

    AFRAC_ECFComProblema = 1
    
    Select Case gErr
    
        Case 214014
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 214015)

    End Select
    
    Exit Function

End Function

Public Function AFRAC_LerRelatorioGerencial(sRelatorios As String) As Integer
'Le os nomes dos  relatorios gerenciais cadastrados

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim objTiposMeiosPagtos As ClassTMPLoja

On Error GoTo Erro_AFRAC_LerRelatorioGerencial

    AFRAC_LerRelatorioGerencial = 1

    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            sRelatorios = Space(660)

            AFRAC_LerRelatorioGerencial = Bematech_FI_VerificaRelatorioGerencialMFD(sRelatorios)
            
            sRelatorios = StringZ(sRelatorios)
            
            If AFRAC_LerRelatorioGerencial = 1 Then
                AFRAC_LerRelatorioGerencial = SUCESSO
            Else
                AFRAC_LerRelatorioGerencial = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            sRelatorios = Space(360)

            AFRAC_LerRelatorioGerencial = rLerRG_ECF_Daruma(sRelatorios)

            If AFRAC_LerRelatorioGerencial = 1 Then
                AFRAC_LerRelatorioGerencial = SUCESSO
            Else
                AFRAC_LerRelatorioGerencial = 1
            End If
        
        Case IMPRESSORA_ELGIN
'            For Each objTiposMeiosPagtos In X.gcolTiposMeiosPagtos
'                If CInt(sIndice) = objTiposMeiosPagtos.iTipo Then
'                    iStatus = Elgin_ProgramaLegenda(Format(objTiposMeiosPagtos.iIndice, "00"), DesacentuaTexto(Format(left(sDescricao, 16), "!@@@@@@@@@@@@@@@@")))
'
'                    sRetorno = Space(50)
'                    iStatus = Elgin_EsperaResposta(sRetorno)
'
'                    Exit For
'                End If
'            Next
'
'
'            If iStatus = SUCESSO Or iStatus = -20 Then
'                AFRAC_LerRelatorioGerencial = SUCESSO
'            Else
'                AFRAC_LerRelatorioGerencial = 1
'            End If
            
        Case Else
            AFRAC_LerRelatorioGerencial = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_LerRelatorioGerencial:

    AFRAC_LerRelatorioGerencial = 1
    
    Exit Function

End Function

Public Function AFRAC_LerFormasPagamento(sFgto As String) As Integer
'Retorno a Ser Preenchido com as Formas de Pagto Programadas no ECF

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim objTiposMeiosPagtos As ClassTMPLoja

On Error GoTo Erro_AFRAC_LerFormasPagamento

    AFRAC_LerFormasPagamento = 1

    Select Case X.giCodModeloECF

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            ' Exemplo em Visual Basic
            
            sFgto = Space(920)
            'le os meios de pagamento com tamanho de 46 para cada meio de pagamento
            AFRAC_LerFormasPagamento = Bematech_FI_VerificaFormasPagamentoMFD(sFgto)
            
            sFgto = StringZ(sFgto)
            
            If AFRAC_LerFormasPagamento = 1 Then
                AFRAC_LerFormasPagamento = SUCESSO
            ElseIf AFRAC_LerFormasPagamento = 0 Then
                AFRAC_LerFormasPagamento = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            sFgto = Space(330)
            'le os meios de pagamento separados por ponto e virgula
            AFRAC_LerFormasPagamento = rLerMeiosPagto_ECF_Daruma(sFgto)

        
            If AFRAC_LerFormasPagamento = 1 Then
                AFRAC_LerFormasPagamento = SUCESSO
            ElseIf AFRAC_LerFormasPagamento = 0 Then
                AFRAC_LerFormasPagamento = 1
            End If
        
        Case Else
            AFRAC_LerFormasPagamento = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_LerFormasPagamento:

    AFRAC_LerFormasPagamento = 1
    
    Exit Function


End Function

Private Function Bematech_StringParaData(ByVal sData As String) As Date
    
    Dim sAux As String, dt As Date
    
    sAux = Trim(sData)
    
    If Len(sAux) = 0 Or left(sAux, 2) = "00" Then
        dt = DATA_NULA
    Else
        dt = CDate(left(sAux, 2) & "/" & Mid(sAux, 3, 2) & "/" & Mid(sAux, 5, 2))
    End If
    
    Bematech_StringParaData = dt
    
End Function

Private Function Bematech_Emulador_ZPendente(iZPendenteAnt As Integer, iZPendenteAtual As Integer, iBloqueadoRZ As Integer) As Integer
'http://partners.bematech.com.br/2011/09/edicao-85-verificando-se-ha-reducao-z-pendente/

Dim sDataMovAtual As String, iRetorno As Integer, sDataMovUltZ As String
Dim sDataReducao As String, sHoraReducao As String
Dim sDataAtual As String, sHoraAtual As String

Dim dtMovAtual As Date, dtMovUltZ As Date
Dim dtReducao As Date, dtAtual As Date

Dim FlagFiscal As Integer
    
On Error GoTo Erro_Bematech_Emulador_ZPendente

    iBloqueadoRZ = 0
    iZPendenteAnt = 0
    iZPendenteAtual = 0

    sDataMovAtual = Space(7)
    iRetorno = Bematech_FI_DataMovimento(sDataMovAtual)
    If iRetorno <> 1 Then Error 1600
    dtMovAtual = Bematech_StringParaData(sDataMovAtual)

    sDataMovUltZ = Space(7)
    iRetorno = Bematech_FI_DataMovimentoUltimaReducaoMFD(sDataMovUltZ)
    If iRetorno <> 1 Then Error 1600
    dtMovUltZ = Bematech_StringParaData(sDataMovUltZ)

    sDataReducao = Space(7)
    sHoraReducao = Space(7)
    iRetorno = Bematech_FI_DataHoraReducao(sDataReducao, sHoraReducao)
    If iRetorno <> 1 Then Error 1600
    dtReducao = Bematech_StringParaData(sDataReducao)

    sDataAtual = Space(7)
    sHoraAtual = Space(7)
    iRetorno = Bematech_FI_DataHoraImpressora(sDataAtual, sHoraAtual)
    If iRetorno <> 1 Then Error 1600
    dtAtual = Bematech_StringParaData(sDataAtual)

    '1o. caso
    If dtMovUltZ = DATA_NULA And dtMovAtual <> DATA_NULA Then
        If dtMovAtual < dtAtual Then
            iZPendenteAnt = 1
        Else
            iZPendenteAtual = 1
        End If
    ElseIf dtMovUltZ <> DATA_NULA And dtMovAtual <> DATA_NULA Then
        If dtMovUltZ < dtMovAtual Then
            iZPendenteAtual = 1
        End If
    ElseIf dtMovUltZ <> DATA_NULA And dtMovAtual = DATA_NULA Then
        iZPendenteAnt = 0
    Else
        FlagFiscal = 0
        
        iRetorno = Bematech_FI_FlagsFiscais(FlagFiscal)
        If iRetorno <> 1 Then Error 1600
        'já houve reducao z no dia
        If (FlagFiscal >= 8) Then
            iBloqueadoRZ = 1
        End If
        
    End If

    Bematech_Emulador_ZPendente = 1
   
    Exit Function
   
Erro_Bematech_Emulador_ZPendente:

    Bematech_Emulador_ZPendente = 0
    
    Exit Function

End Function

Public Function AFRAC_AcionarGuilhotina(ByVal sCorte As String) As Integer
'Retorno a Ser Preenchido com as Formas de Pagto Programadas no ECF

Dim X As New ClassECFConfig
Dim iStatus As Integer
Dim sRetorno As String
Dim iCorte As Integer
Dim Y As New ClassSATGlobal, Z As New ClassNFeGlobal
Dim iModeloImpressora As Integer, sPortaImpressora As String, objRel As Object

On Error GoTo Erro_AFRAC_AcionarGuilhotina

    AFRAC_AcionarGuilhotina = 1

    iModeloImpressora = X.giCodModeloECF
    
    If AFRAC_ImpressoraCFe(iModeloImpressora) Then
        
        Select Case iModeloImpressora
            
            Case IMPRESSORA_SAT_2_5_15
                iModeloImpressora = Y.gobjSATInfo.iModeloImpressora
                sPortaImpressora = Y.gobjSATInfo.sPortaImpressora
                
            Case IMPRESSORA_NFCE
                iModeloImpressora = Z.gobjNFeInfo.iModeloImpressora
                sPortaImpressora = Z.gobjNFeInfo.sPortaImpressora
                
            Case Else
                iModeloImpressora = IMPRESSORA_NAO_FISCAL
                
        End Select
        
    End If
    
    Select Case iModeloImpressora

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            ' Exemplo em Visual Basic
            
            If sCorte = "T" Then
                iCorte = 1
            Else
                iCorte = 0
            End If
            
            AFRAC_AcionarGuilhotina = Bematech_FI_AcionaGuilhotinaMFD(iCorte)
            
            If AFRAC_AcionarGuilhotina = 1 Then
                AFRAC_AcionarGuilhotina = SUCESSO
            Else
                AFRAC_AcionarGuilhotina = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            If sCorte = "T" Then
                sCorte = "0"
            Else
                sCorte = "1"
            End If
        
        
            'le os meios de pagamento separados por ponto e virgula
            AFRAC_AcionarGuilhotina = eAcionarGuilhotina_ECF_Daruma(sCorte)

        
            If AFRAC_AcionarGuilhotina = 1 Then
                AFRAC_AcionarGuilhotina = SUCESSO
            Else
                AFRAC_AcionarGuilhotina = 1
            End If
        
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS
            Set objRel = New ClassImpressoraESC
            Call objRel.Define_Impressora(sPortaImpressora, iModeloImpressora)
            iStatus = objRel.AcionarGuilhotina(sCorte)
            iStatus = objRel.TerminaImpressao
            Set objRel = Nothing
            If iStatus = 1 Then
                AFRAC_AcionarGuilhotina = SUCESSO
            Else
                AFRAC_AcionarGuilhotina = 1
            End If
        
        Case Else
            AFRAC_AcionarGuilhotina = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_AcionarGuilhotina:

    AFRAC_AcionarGuilhotina = 1
    
    Exit Function


End Function

Public Function AFRAC_AbrirRelatorioGerencial(sIndice As String, ByVal objTela As Object) As Integer
'Abre um relatorio gerencial
    
Dim X As New ClassECFConfig
Dim iStatus As Integer, Y As New ClassSATGlobal, Z As New ClassNFeGlobal, iRetornoBematech As Integer
Dim sRetorno As String, iRetornoEpson As Integer
Dim lErro As Long, iModeloImpressora As Integer
Dim sNome As String, iRetornoDual As Integer, sPortaImpressora As String

'Sangria ou suprimento
On Error GoTo Erro_AFRAC_AbrirRelatorioGerencial

    AFRAC_AbrirRelatorioGerencial = 1

    iModeloImpressora = X.giCodModeloECF
    
    If AFRAC_ImpressoraCFe(iModeloImpressora) Then
        
        Select Case iModeloImpressora
            
            Case IMPRESSORA_SAT_2_5_15
                iModeloImpressora = Y.gobjSATInfo.iModeloImpressora
                sPortaImpressora = Y.gobjSATInfo.sPortaImpressora
                
            Case IMPRESSORA_NFCE
                iModeloImpressora = Z.gobjNFeInfo.iModeloImpressora
                sPortaImpressora = Z.gobjNFeInfo.sPortaImpressora
                
            Case Else
                iModeloImpressora = IMPRESSORA_NAO_FISCAL
                
        End Select
    
    End If

    Select Case iModeloImpressora
    
        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            
            If sIndice = RELGER_IDENTPAFECF Or sIndice = RELGER_PARAMCONFIG Then
                AFRAC_AbrirRelatorioGerencial = Bematech_FI_AbreRelatorioGerencialMFD(sIndice)

            End If

            If AFRAC_AbrirRelatorioGerencial = 1 Then
                AFRAC_AbrirRelatorioGerencial = SUCESSO
            ElseIf AFRAC_AbrirRelatorioGerencial = 0 Then
                    AFRAC_AbrirRelatorioGerencial = 1
            End If
    
    
        Case IMPRESSORA_DARUMA_DR700
'            iRetornoDual = regAlterarValor_Daruma("DUAL\PortaComunicacao", sPortaImpressora)
'            If iRetornoDual = 1 Then
'                Call regAlterarValor_Daruma("DUAL\Velocidade", "115200")
'                iRetornoDual = rStatusImpressora_DUAL_DarumaFramework
'            End If
            
'            If iRetornoDual <> 1 Then
                iRetornoDual = eBuscarPortaVelocidade_DUAL_DarumaFramework
'            End If
            
            If iRetornoDual = 1 Then
                AFRAC_AbrirRelatorioGerencial = SUCESSO
            ElseIf iRetornoDual = 0 Then
                AFRAC_AbrirRelatorioGerencial = 1
            Else
                AFRAC_AbrirRelatorioGerencial = iRetornoDual
            End If
            
        Case IMPRESSORA_DARUMA_FS600
    
            Select Case sIndice
            
                Case CStr(RELGER_IDENTPAFECF)
                    AFRAC_AbrirRelatorioGerencial = iRGAbrir_ECF_Daruma("Ident. PAF-ECF")
                
                Case CStr(RELGER_PARAMCONFIG)
                    AFRAC_AbrirRelatorioGerencial = iRGAbrir_ECF_Daruma("Param. Config.")
    
            End Select
    
            If AFRAC_AbrirRelatorioGerencial = 1 Then
                AFRAC_AbrirRelatorioGerencial = SUCESSO
            ElseIf AFRAC_AbrirRelatorioGerencial = 0 Then
                    AFRAC_AbrirRelatorioGerencial = 1
            End If
    
        Case IMPRESSORA_NAO_FISCAL
        
            lErro = CF_ECF("Configura_Impressora_ECF", objTela)
            If lErro <> SUCESSO Then gError 204311
        
            AFRAC_AbrirRelatorioGerencial = SUCESSO
        
        Case IMPRESSORA_ELGIN

            iStatus = Elgin_LeituraXComRelGer()
            
            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)
            
            If iStatus = SUCESSO Then
                AFRAC_AbrirRelatorioGerencial = SUCESSO
            Else
                AFRAC_AbrirRelatorioGerencial = 1
            End If
        
        Case IMPRESSORA_BEMATECH_MP4000_TH
            iRetornoBematech = BematechNaoFiscal_ConfiguraModeloImpressora(5)
            If iRetornoBematech = 1 Then
                iRetornoBematech = BematechNaoFiscal_IniciaPorta(sPortaImpressora)
            End If
            
            If iRetornoBematech = 1 Then
                AFRAC_AbrirRelatorioGerencial = SUCESSO
            ElseIf iRetornoBematech = 0 Then
                AFRAC_AbrirRelatorioGerencial = 1
            Else
                AFRAC_AbrirRelatorioGerencial = iRetornoBematech
            End If
        
        Case IMPRESSORA_EPSON_TM_T20
        
            Call EpsonNaoFiscal_ConfiguraTaxaSerial(115200)
            
            iRetornoEpson = EpsonNaoFiscal_IniciaPorta(sPortaImpressora)
            If iRetornoEpson = 1 Then
            
                AFRAC_AbrirRelatorioGerencial = SUCESSO
                
            Else
            
                'a porta pode ter ficado aberta numa impressao anterior
                Call EpsonNaoFiscal_FechaPorta
                
                iRetornoEpson = EpsonNaoFiscal_IniciaPorta(sPortaImpressora)
                If iRetornoEpson = 1 Then
                    AFRAC_AbrirRelatorioGerencial = SUCESSO
                ElseIf iRetornoEpson = 0 Then
                    AFRAC_AbrirRelatorioGerencial = 1
                Else
                    AFRAC_AbrirRelatorioGerencial = iRetornoEpson
                End If
                
            End If
                            
        
        Case Else
            AFRAC_AbrirRelatorioGerencial = SUCESSO
            
    End Select
    
    Exit Function
    
Erro_AFRAC_AbrirRelatorioGerencial:

    Select Case gErr
    
        Case 204311
            AFRAC_AbrirRelatorioGerencial = 1
    
        Case Else
            AFRAC_AbrirRelatorioGerencial = 2
    End Select
    
End Function

Public Function AFRAC_ImprimirRelatorioGerencial(sTexto As String, ByVal objTela As Object) As Integer
'Imprime o texto passado como parametro

Dim X As New ClassECFConfig
Dim iStatus As Integer, Y As New ClassSATGlobal, Z As New ClassNFeGlobal
Dim sRetorno As String, iModeloImpressora As Integer
Dim objRel As Object, sPortaImpressora As String

'Sangria ou suprimento

On Error GoTo Erro_AFRAC_ImprimirRelatorioGerencial

    AFRAC_ImprimirRelatorioGerencial = 1

    iModeloImpressora = X.giCodModeloECF
    
    If AFRAC_ImpressoraCFe(iModeloImpressora) Then
        
        Select Case iModeloImpressora
            
            Case IMPRESSORA_SAT_2_5_15
                iModeloImpressora = Y.gobjSATInfo.iModeloImpressora
                sPortaImpressora = Y.gobjSATInfo.sPortaImpressora
                
            Case IMPRESSORA_NFCE
                iModeloImpressora = Z.gobjNFeInfo.iModeloImpressora
                sPortaImpressora = Z.gobjNFeInfo.sPortaImpressora
                
            Case Else
                iModeloImpressora = IMPRESSORA_NAO_FISCAL
                
        End Select
        
    End If
    
    Select Case iModeloImpressora

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            'Atualiza o arquivo
            Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)

            AFRAC_ImprimirRelatorioGerencial = Bematech_FI_RelatorioGerencial(Bematech_AjustaTextoImpressora(sTexto))
            
            If AFRAC_ImprimirRelatorioGerencial = 1 Then
                AFRAC_ImprimirRelatorioGerencial = SUCESSO
            ElseIf AFRAC_ImprimirRelatorioGerencial = 0 Then
                AFRAC_ImprimirRelatorioGerencial = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

            'Atualiza o arquivo
            Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)

'                AFRAC_ImprimirRelatorioGerencial = Daruma_FI_RelatorioGerencial(sTexto)
            AFRAC_ImprimirRelatorioGerencial = iRGImprimirTexto_ECF_Daruma(sTexto)
            
            If AFRAC_ImprimirRelatorioGerencial = 1 Then
                AFRAC_ImprimirRelatorioGerencial = SUCESSO
            ElseIf AFRAC_ImprimirRelatorioGerencial = 0 Then
                AFRAC_ImprimirRelatorioGerencial = 1
            End If
            
        Case IMPRESSORA_NAO_FISCAL
            If Not (objTela Is Nothing) Then objTela.Controls("RT1").Text = objTela.Controls("RT1").Text & sTexto & Chr(13) & Chr(10)
        
            AFRAC_ImprimirRelatorioGerencial = SUCESSO
        
        Case IMPRESSORA_ELGIN

            'Atualiza o arquivo
            Call WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)

            iStatus = Elgin_ImprimeLinhaNaoFiscal(0, left(Formata_Campo(ALINHAMENTO_CENTRALIZADO, 48, " ", DesacentuaTexto(sTexto)), 48))

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_ImprimirRelatorioGerencial = SUCESSO
            Else
                AFRAC_ImprimirRelatorioGerencial = 1
            End If
        
        Case IMPRESSORA_DARUMA_DR700
            iStatus = iImprimirTexto_DUAL_DarumaFramework(sTexto, 0)
        
            If iStatus = 1 Then
                AFRAC_ImprimirRelatorioGerencial = SUCESSO
            Else
                AFRAC_ImprimirRelatorioGerencial = 1
            End If
        
        Case IMPRESSORA_BEMATECH_MP4000_TH
            iStatus = BematechNaoFiscal_BematechTX(sTexto)
            If iStatus = 1 Then
                AFRAC_ImprimirRelatorioGerencial = SUCESSO
            Else
                AFRAC_ImprimirRelatorioGerencial = 1
            End If
            
        Case IMPRESSORA_EPSON_TM_T20
            iStatus = EpsonNaoFiscal_ImprimeTexto(sTexto)
            If iStatus = 1 Then
                AFRAC_ImprimirRelatorioGerencial = SUCESSO
            Else
                AFRAC_ImprimirRelatorioGerencial = 1
            End If
            
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS
            Set objRel = New ClassImpressoraESC
            Call objRel.Define_Impressora(sPortaImpressora, iModeloImpressora)
            iStatus = objRel.ImprimeNormal(sTexto)
            iStatus = objRel.TerminaImpressao
            Set objRel = Nothing
            If iStatus = 1 Then
                AFRAC_ImprimirRelatorioGerencial = SUCESSO
            Else
                AFRAC_ImprimirRelatorioGerencial = 1
            End If
        
        Case Else
            AFRAC_ImprimirRelatorioGerencial = SUCESSO
        
        
        
    End Select

   Exit Function
   
Erro_AFRAC_ImprimirRelatorioGerencial:

    AFRAC_ImprimirRelatorioGerencial = 1
    
    Exit Function

End Function

Public Function AFRAC_FecharRelatorioGerencial(ByVal objTela As Object) As Integer
'Encerra relatorio gerencial

Dim X As New ClassECFConfig
Dim iStatus As Integer, Y As New ClassSATGlobal, Z As New ClassNFeGlobal
Dim sRetorno As String, iModeloImpressora As Integer
Dim objRel As Object, sPortaImpressora As String

On Error GoTo Erro_AFRAC_FecharRelatorioGerencial

    AFRAC_FecharRelatorioGerencial = 1

    iModeloImpressora = X.giCodModeloECF
    
    If AFRAC_ImpressoraCFe(iModeloImpressora) Then
        
        Select Case iModeloImpressora
            
            Case IMPRESSORA_SAT_2_5_15
                iModeloImpressora = Y.gobjSATInfo.iModeloImpressora
                sPortaImpressora = Y.gobjSATInfo.sPortaImpressora
                
            Case IMPRESSORA_NFCE
                iModeloImpressora = Z.gobjNFeInfo.iModeloImpressora
                sPortaImpressora = Z.gobjNFeInfo.sPortaImpressora
                
            Case Else
                iModeloImpressora = IMPRESSORA_NAO_FISCAL
                
        End Select
    
    End If
    
    Select Case iModeloImpressora

        Case 1 To 3, 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI

            AFRAC_FecharRelatorioGerencial = Bematech_FI_FechaRelatorioGerencial()

            If AFRAC_FecharRelatorioGerencial = 1 Then
                AFRAC_FecharRelatorioGerencial = SUCESSO
                Call CF_ECF("Grava_R06", "RG")
            ElseIf AFRAC_FecharRelatorioGerencial = 0 Then
                AFRAC_FecharRelatorioGerencial = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600

'                AFRAC_FecharRelatorioGerencial = Daruma_FI_FechaRelatorioGerencial()
            AFRAC_FecharRelatorioGerencial = iRGFechar_ECF_Daruma()

            If AFRAC_FecharRelatorioGerencial = 1 Then
                AFRAC_FecharRelatorioGerencial = SUCESSO
                Call CF_ECF("Grava_R06", "RG")
            ElseIf AFRAC_FecharRelatorioGerencial = 0 Then
                AFRAC_FecharRelatorioGerencial = 1
            End If
            
        Case IMPRESSORA_NAO_FISCAL
            If Not (objTela Is Nothing) Then
                objTela.Controls("RT1").SelStart = 0
                objTela.Controls("RT1").SelLength = Len(objTela.Controls("RT1").Text)
                objTela.Controls("RT1").SelFontName = "Courier New"
'                If objTela.Controls("CD1").hdc > 0 Then objTela.Controls("RT1").SelPrint objTela.Controls("CD1").hdc
                objTela.Controls("RT1").SelFontSize = 10
                Call objTela.Controls("RT1").SelPrint(objTela.Controls("CD1").hdc)

            End If
            
            AFRAC_FecharRelatorioGerencial = SUCESSO
            
        Case IMPRESSORA_ELGIN

            iStatus = Elgin_EncerraCupomNaoFiscal()

            sRetorno = Space(50)
            iStatus = Elgin_EsperaResposta(sRetorno)

            If iStatus = SUCESSO Then
                AFRAC_FecharRelatorioGerencial = SUCESSO
            Else
                AFRAC_FecharRelatorioGerencial = 1
            End If
            
        Case IMPRESSORA_BEMATECH_MP4000_TH
            iStatus = BematechNaoFiscal_FechaPorta
            If iStatus = 1 Then
                AFRAC_FecharRelatorioGerencial = SUCESSO
            Else
                AFRAC_FecharRelatorioGerencial = 1
            End If
        
        Case IMPRESSORA_EPSON_TM_T20
            iStatus = EpsonNaoFiscal_FechaPorta
            If iStatus = 1 Then
                AFRAC_FecharRelatorioGerencial = SUCESSO
            Else
                AFRAC_FecharRelatorioGerencial = 1
            End If
        
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS
            Set objRel = New ClassImpressoraESC
            Call objRel.Define_Impressora(sPortaImpressora, iModeloImpressora)

            'pular algumas linhas para poder cortar
            iStatus = objRel.ImprimeNormal("")
            iStatus = objRel.ImprimeNormal("")
            iStatus = objRel.ImprimeNormal("")
            iStatus = objRel.ImprimeNormal("")
            iStatus = objRel.ImprimeNormal("")
            iStatus = objRel.ImprimeNormal("")
            iStatus = objRel.ImprimeNormal("")
            iStatus = objRel.TerminaImpressao
            Set objRel = Nothing
            If iStatus = 1 Then
                AFRAC_FecharRelatorioGerencial = SUCESSO
            Else
                AFRAC_FecharRelatorioGerencial = 1
            End If
        
        Case Else
            AFRAC_FecharRelatorioGerencial = SUCESSO
        
            
    End Select

   Exit Function
   
Erro_AFRAC_FecharRelatorioGerencial:

    AFRAC_FecharRelatorioGerencial = 1
    
    Exit Function

End Function

Public Function AFRAC_AcionarGuilhotina2(ByVal sCorte As String, ByVal objRel As Object) As Integer
'Retorno a Ser Preenchido com as Formas de Pagto Programadas no ECF

Dim X As New ClassECFConfig
Dim iStatus As Integer, Y As New ClassSATGlobal, Z As New ClassNFeGlobal
Dim sRetorno As String, iModeloImpressora As Integer
Dim iCorte As Integer, sTexto As String

On Error GoTo Erro_AFRAC_AcionarGuilhotina2

    AFRAC_AcionarGuilhotina2 = 1

    iModeloImpressora = X.giCodModeloECF
    
    If AFRAC_ImpressoraCFe(iModeloImpressora) Then
        
        Select Case iModeloImpressora
            
            Case IMPRESSORA_SAT_2_5_15
                iModeloImpressora = Y.gobjSATInfo.iModeloImpressora
                
            Case IMPRESSORA_NFCE
                iModeloImpressora = Z.gobjNFeInfo.iModeloImpressora
                
            Case Else
                iModeloImpressora = IMPRESSORA_NAO_FISCAL
                
        End Select
    
    End If
    
    Select Case iModeloImpressora

        Case 5, IMPRESSORA_BEMATECH_MP4200_TH_FI, IMPRESSORA_BEMATECH_MP2100_TH_FI
            
            ' Exemplo em Visual Basic
            
            If sCorte = "T" Then
                iCorte = 1
            Else
                iCorte = 0
            End If
            
            AFRAC_AcionarGuilhotina2 = Bematech_FI_AcionaGuilhotinaMFD(iCorte)
            
            If AFRAC_AcionarGuilhotina2 = 1 Then
                AFRAC_AcionarGuilhotina2 = SUCESSO
            Else
                AFRAC_AcionarGuilhotina2 = 1
            End If
            
        Case IMPRESSORA_DARUMA_FS600
        
            If sCorte = "T" Then
                sCorte = "0"
            Else
                sCorte = "1"
            End If
        
        
            'le os meios de pagamento separados por ponto e virgula
            AFRAC_AcionarGuilhotina2 = eAcionarGuilhotina_ECF_Daruma(sCorte)

        
            If AFRAC_AcionarGuilhotina2 = 1 Then
                AFRAC_AcionarGuilhotina2 = SUCESSO
            Else
                AFRAC_AcionarGuilhotina2 = 1
            End If
        
        Case IMPRESSORA_DARUMA_DR700
        
            sTexto = "<gui><confgui>" & sCorte & "</confgui></gui>"
            
            If objRel Is Nothing Then
                iStatus = iImprimirTexto_DUAL_DarumaFramework(sTexto, 0)
            Else
                iStatus = objRel.ImprimeNormal(sTexto)
            End If
        
            If iStatus = 1 Then
                AFRAC_AcionarGuilhotina2 = SUCESSO
            Else
                AFRAC_AcionarGuilhotina2 = 1
            End If
        
        Case IMPRESSORA_BEMATECH_MP4000_TH
            
            If sCorte = "T" Then
                iCorte = 1
            Else
                iCorte = 0
            End If
            
            If Not (objRel Is Nothing) Then objRel.DescarregaBuffer
            
            iStatus = BematechNaoFiscal_AcionaGuilhotina(iCorte)
        
            If iStatus = 1 Then
                AFRAC_AcionarGuilhotina2 = SUCESSO
            Else
                AFRAC_AcionarGuilhotina2 = 1
            End If
            
        Case IMPRESSORA_EPSON_TM_T20
        
            If sCorte = "T" Then
                iCorte = 1
            Else
                iCorte = 0
            End If
            
            If Not (objRel Is Nothing) Then objRel.DescarregaBuffer
        
            iStatus = EpsonNaoFiscal_AcionaGuilhotina(iCorte)
        
            If iStatus = 1 Then
                AFRAC_AcionarGuilhotina2 = SUCESSO
            Else
                AFRAC_AcionarGuilhotina2 = 1
            End If
        
        Case IMPRESSORA_ELGIN_VOX, IMPRESSORA_BEMATECH_ESCBEMA, IMPRESSORA_EPSON_ESCPOS, IMPRESSORA_DARUMA_ESCPOS

            If Not (objRel Is Nothing) Then
                Call objRel.AcionarGuilhotina(sCorte)
            End If
            AFRAC_AcionarGuilhotina2 = SUCESSO
        
        Case Else
            AFRAC_AcionarGuilhotina2 = SUCESSO
        
    End Select

   Exit Function
   
Erro_AFRAC_AcionarGuilhotina2:

    AFRAC_AcionarGuilhotina2 = 1
    
    Exit Function


End Function


