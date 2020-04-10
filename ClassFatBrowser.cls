VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFatBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Function NFiscalSaidaTodasLista_BotaoEdita(objNFiscalSaida As ClassNFiscal, lErro As Long) As Long

Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_NFiscalSaidaTodasLista_BotaoEdita

    If lErro = SUCESSO Then
    
        If objNFiscalSaida.iTipoNFiscal = 0 Then
            lErro = CF("NFiscal_Le", objNFiscalSaida)
            If lErro <> SUCESSO And lErro <> 31442 Then gError 63056
        End If

        Set objTipoDocInfo = New ClassTipoDocInfo
    
        objTipoDocInfo.iCodigo = objNFiscalSaida.iTipoNFiscal
    
        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then Error 63056
        If lErro = 31415 Then Error 63057 'Não encontrou
    
'        If giTipoVersao = VERSAO_FULL Then
        If UCase(objTipoDocInfo.sNomeTelaNFiscal) <> "VENDAP" And UCase(objTipoDocInfo.sNomeTelaNFiscal) <> "VENDAM" Then
            Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscalSaida)
        Else
            lErro = CF("NFiscal_Le_ChvNFe", objNFiscalSaida)
            If lErro <> SUCESSO Then gError 63056
        
            If objNFiscalSaida.sQRCode <> "" Then
                ShellExecute 0, "open", objNFiscalSaida.sQRCode, vbNullString, vbNullString, 2
            Else
                Call Rotina_Aviso(vbOKOnly, "As NFCes somente podem ser consultadas no site do Sefaz.")
            End If
            Error 63055 'Só para manter a tela aberta
        
        End If
'        ElseIf giTipoVersao = VERSAO_LIGHT Then
'            Call Chama_Tela("NFiscalFatura", objNFiscalSaida)
'        End If

    Else
    
            Call Chama_Tela("NFiscalFatura")
    
    End If
    
    NFiscalSaidaTodasLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_NFiscalSaidaTodasLista_BotaoEdita:

    NFiscalSaidaTodasLista_BotaoEdita = Err

    Select Case Err

        Case 63055, 63056

        Case 63057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149731)

    End Select

    Exit Function

End Function

Function NFiscalInternaSaidaLista_BotaoEdita(objNFiscal As ClassNFiscal, lErro As Long) As Long

Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_NFiscalInternaSaidaLista_BotaoEdita
   
    If lErro = SUCESSO Then
    
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
        'Lê o Tipo de Documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 95000
    
        'Se não achou o Tipo de Documento --> erro
        If lErro = 31415 Then gError 95001
    
        If objTipoDocInfo.sNomeTelaNFiscal = "NFiscal" Then
        
            Call Chama_Tela("NFiscal", objNFiscal)
                    
        ElseIf objTipoDocInfo.sNomeTelaNFiscal = "NFiscalRem" Then
        
            Call Chama_Tela("NFiscalRem", objNFiscal)
                    
        ElseIf objTipoDocInfo.sNomeTelaNFiscal = "NFiscalDev" Then
        
            Call Chama_Tela("NFiscalDev", objNFiscal)
                    
        ElseIf objTipoDocInfo.sNomeTelaNFiscal = "NFiscalFatura" Then
        
            Call Chama_Tela("NFiscalFatura", objNFiscal)
                    
        ElseIf objTipoDocInfo.sNomeTelaNFiscal = "NFiscalPedido" Then
        
            Call Chama_Tela("NFiscalPedido", objNFiscal)
                    
        ElseIf objTipoDocInfo.sNomeTelaNFiscal = "NFiscalFaturaPedido" Then
        
            Call Chama_Tela("NFiscalFaturaPedido", objNFiscal)
        
        End If
    Else
        Call Chama_Tela("NFiscal")
    End If
    
    NFiscalInternaSaidaLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_NFiscalInternaSaidaLista_BotaoEdita:

    NFiscalInternaSaidaLista_BotaoEdita = Err

    Select Case Err

        Case 95000
        
        Case 95001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149732)

    End Select

    Exit Function

End Function

Function NFConhecimentoFreteLista_BotaoEdita(objNFiscal As ClassNFiscal, lErro As Long) As Long

On Error GoTo Erro_NFConhecimentoFreteLista_BotaoEdita

    If objNFiscal Is Nothing Then Call Chama_Tela("ConhecimentoFreteFatura")

    If objNFiscal.iTipoNFiscal = TIPODOCINFO_CONHECIMENTOFRETE Then
        Call Chama_Tela("ConhecimentoFrete", objNFiscal)

    ElseIf objNFiscal.iTipoNFiscal = TIPODOCINFO_CONHECIMENTOFRETE_FATURA Then
        Call Chama_Tela("ConhecimentoFreteFatura", objNFiscal)

    End If
    
    NFConhecimentoFreteLista_BotaoEdita = SUCESSO

    Exit Function

Erro_NFConhecimentoFreteLista_BotaoEdita:

    NFConhecimentoFreteLista_BotaoEdita = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149733)

    End Select

    Exit Function

End Function

Function FilialClienteCompleto_BotaoEdita(objFilialCliente As ClassFilialCliente, lErro As Long) As Long

Dim objCliente As ClassCliente

On Error GoTo Erro_FilialClienteCompleto_BotaoEdita

    If Not (objFilialCliente Is Nothing) Then

        If objFilialCliente.iCodFilial = FILIAL_MATRIZ Then
        
            Set objCliente = New ClassCliente
            
            'passa para objCliente o codigo do cliente correspondente em objFilialCliente
            objCliente.lCodigo = objFilialCliente.lCodCliente
            
            'le os dados do cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 123517
            If lErro = 12293 Then gError 123518
            
            Call Chama_Tela("Clientes", objCliente)
            
        Else
        
            Call Chama_Tela("FiliaisClientes", objFilialCliente)
        
        End If
    
    Else
    
            Call Chama_Tela("Clientes")
    
    End If
    
    FilialClienteCompleto_BotaoEdita = SUCESSO
    
    Exit Function
    
Erro_FilialClienteCompleto_BotaoEdita:

    FilialClienteCompleto_BotaoEdita = gErr

    Select Case gErr
    
    Case 123517
    
    Case 123518
    
    Case Else
       Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149734)

    End Select

    Exit Function

End Function


Function NF_FaturaReceberLista_BotaoEdita(objNFiscal As ClassNFiscal, lErro As Long) As Long

Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_NF_FaturaReceberLista_BotaoEdita

        If objNFiscal Is Nothing Then gError 95002

        Set objTipoDocInfo = New ClassTipoDocInfo
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 59677
        If lErro = 31415 Then gError 59676 'Não encontrou

        Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscal)
        
    NF_FaturaReceberLista_BotaoEdita = SUCESSO

    Exit Function

Erro_NF_FaturaReceberLista_BotaoEdita:

    NF_FaturaReceberLista_BotaoEdita = gErr

    Select Case gErr

        Case 59677

        Case 59676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case 95002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NAO_SELECIONADA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149735)

    End Select

    Exit Function

End Function

Function NFConhecimentoFreteListaTodos_BotaoEdita(objNFiscal As ClassNFiscal, lErro As Long) As Long

On Error GoTo Erro_NFConhecimentoFreteListaTodos_BotaoEdita

    If lErro = SUCESSO Then
    
        If objNFiscal.iTipoNFiscal = TIPODOCINFO_CONHECIMENTOFRETE Then
            Call Chama_Tela("ConhecimentoFrete", objNFiscal)
        
        ElseIf objNFiscal.iTipoNFiscal = TIPODOCINFO_CONHECIMENTOFRETE_FATURA Then
            Call Chama_Tela("ConhecimentoFreteFatura", objNFiscal)
        
        End If
    
    Else
    
            Call Chama_Tela("ConhecimentoFreteFatura")
    
    End If
    
    
    NFConhecimentoFreteListaTodos_BotaoEdita = SUCESSO
    
    Exit Function

Erro_NFConhecimentoFreteListaTodos_BotaoEdita:

    NFConhecimentoFreteListaTodos_BotaoEdita = gErr

    Select Case gErr

        Case 91303
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149736)

    End Select

    Exit Function
'Janaina
End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function CamposGenericos_Lista_TrataParametros(ByVal objTela As Object, ByVal colSelecao As Collection) As Long

On Error GoTo Erro_CamposGenericos_Lista_TrataParametros

    Select Case colSelecao(1)
        
        Case CAMPOSGENERICOS_VOLUMEESPECIE
            objTela.Caption = "Lista de Espécies de Volumes"
            
        Case CAMPOSGENERICOS_VOLUMEMARCA
            objTela.Caption = "Lista de Marcas de Volumes"
        
        Case CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES
            objTela.Caption = "Lista de Tipos de Relacionamentos com clientes"
            
    End Select
    
    CamposGenericos_Lista_TrataParametros = SUCESSO
    
    Exit Function
    
Erro_CamposGenericos_Lista_TrataParametros:

    CamposGenericos_Lista_TrataParametros = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149737)

    End Select
    
End Function

'###########################################################################
'Inserido por Wagner
Function MedicaoFornItensAPagarLista_BotaoEdita(ByVal objItemMedicao As ClassItensMedCtr, lErro As Long) As Long

Dim objMedicaoContrato As New ClassMedicaoContrato

On Error GoTo Erro_MedicaoFornItensAPagarLista_BotaoEdita

    If lErro = SUCESSO Then
    
        objMedicaoContrato.lCodigo = objItemMedicao.lMedicao
    
        Call Chama_Tela("ContratoMedicaoPag", objMedicaoContrato)

    Else
        Call Chama_Tela("ContratoMedicaoPag")
    End If
    
    MedicaoFornItensAPagarLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_MedicaoFornItensAPagarLista_BotaoEdita:

    MedicaoFornItensAPagarLista_BotaoEdita = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149738)

    End Select

    Exit Function

End Function

Function MedicaoCliItensAFaturarLista_BotaoEdita(ByVal objItemMedicao As ClassItensMedCtr, lErro As Long) As Long

Dim objMedicaoContrato As New ClassMedicaoContrato

On Error GoTo Erro_MedicaoCliItensAFaturarLista_BotaoEdita

    If lErro = SUCESSO Then
    
        objMedicaoContrato.lCodigo = objItemMedicao.lMedicao
    
        Call Chama_Tela("ContratoMedicao", objMedicaoContrato)

    Else
        Call Chama_Tela("ContratoMedicao")
    End If
    
    MedicaoCliItensAFaturarLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_MedicaoCliItensAFaturarLista_BotaoEdita:

    MedicaoCliItensAFaturarLista_BotaoEdita = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149739)

    End Select

    Exit Function

End Function

Function ContratosCliItensLista_BotaoEdita(ByVal objItemContrato As ClassItensDeContrato, lErro As Long) As Long

Dim objContrato As New ClassContrato

On Error GoTo Erro_ContratosCliItensLista_BotaoEdita

    If lErro = SUCESSO Then
    
        objContrato.lNumIntDoc = objItemContrato.lNumIntContrato
    
        Call Chama_Tela("ContratoCadastro", objContrato)

    Else
        Call Chama_Tela("ContratoCadastro")
    End If
    
    ContratosCliItensLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_ContratosCliItensLista_BotaoEdita:

    ContratosCliItensLista_BotaoEdita = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149740)

    End Select

    Exit Function

End Function

Function ContratosFornItensLista_BotaoEdita(ByVal objItemContrato As ClassItensDeContrato, lErro As Long) As Long

Dim objContrato As New ClassContrato

On Error GoTo Erro_ContratosFornItensLista_BotaoEdita

    If lErro = SUCESSO Then
    
        objContrato.lNumIntDoc = objItemContrato.lNumIntContrato
    
        Call Chama_Tela("ContratoPagar", objContrato)

    Else
        Call Chama_Tela("ContratoPagar")
    End If
    
    ContratosFornItensLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_ContratosFornItensLista_BotaoEdita:

    ContratosFornItensLista_BotaoEdita = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149741)

    End Select

    Exit Function

End Function
'###########################################################################

Public Function ProjetoInfoLista_BotaoEdita(ByVal objProjetoInfo As ClassProjetoInfo, lErro As Long) As Long

Dim sDocOrigem As String
Dim sNomeTela As String
Dim objOrigem As Object
Dim iFilialDoc As Integer
Dim sDocOrigemRed As String

On Error GoTo Erro_ProjetoInfoLista_BotaoEdita

    lErro = CF("Projeto_Le_DocInfo", objProjetoInfo.lNumIntDoc, sDocOrigem, objOrigem, sNomeTela, sDocOrigemRed, iFilialDoc)
    If lErro <> SUCESSO Then gError 182053
    
    If iFilialDoc <> giFilialEmpresa And iFilialDoc <> 0 Then gError 182054
    
    Call Chama_Tela(sNomeTela, objOrigem)
    
    ProjetoInfoLista_BotaoEdita = SUCESSO

    Exit Function

Erro_ProjetoInfoLista_BotaoEdita:

    ProjetoInfoLista_BotaoEdita = gErr

    Select Case gErr
    
        Case 182053
        
        Case 182054
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMP_DOC_DIF_FILIALCORRENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182059)

    End Select

    Exit Function

End Function

Public Function DocFaturadosLista_BotaoConsulta(ByVal objVoucher As Object, lErro As Long) As Long

On Error GoTo Erro_DocFaturadosLista_BotaoConsulta

    lErro = CF("DocFaturadosLista_BotaoConsulta", objVoucher)
    If lErro <> SUCESSO Then gError 192877

    Exit Function

Erro_DocFaturadosLista_BotaoConsulta:

    DocFaturadosLista_BotaoConsulta = gErr

    Select Case gErr
    
        Case 192877

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192878)

    End Select

    Exit Function

End Function

Public Function TRVAPortesPagtoFatHistResLista_BotaoConsulta(ByVal objVoucher As Object, lErro As Long) As Long

On Error GoTo Erro_TRVAPortesPagtoFatHistResLista_BotaoConsulta

    lErro = CF("TRVAPortesPagtoFatHistResLista_BotaoConsulta", objVoucher)
    If lErro <> SUCESSO Then gError 192877

    Exit Function

Erro_TRVAPortesPagtoFatHistResLista_BotaoConsulta:

    TRVAPortesPagtoFatHistResLista_BotaoConsulta = gErr

    Select Case gErr
    
        Case 192877

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192878)

    End Select

    Exit Function

End Function


Function RPSLista_BotaoEdita(objRPS As ClassRPS, lErro As Long) As Long

Dim objTipoDocInfo As ClassTipoDocInfo
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_RPS_BotaoEdita

    If lErro = SUCESSO Then
    
        objNFiscal.lNumIntDoc = objRPS.lNumIntDocNF

        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 192925

        Set objTipoDocInfo = New ClassTipoDocInfo
    
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 192926
        If lErro = 31415 Then gError 192927 'Não encontrou
    
'        If giTipoVersao = VERSAO_FULL Then
            Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscal)
'        ElseIf giTipoVersao = VERSAO_LIGHT Then
'            Call Chama_Tela("NFiscalFatura", objNFiscal)
'        End If

    Else
    
        Call Chama_Tela("NFiscalFatura")
    
    End If
    
    Exit Function

Erro_RPS_BotaoEdita:

    Select Case gErr

        Case 192925, 192926

        Case 192927
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192928)

    End Select

    Exit Function

End Function

Function NFeLista_BotaoEdita(objNFe As ClassNFe, lErro As Long) As Long

Dim objTipoDocInfo As ClassTipoDocInfo
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_NFeLista_BotaoEdita

    If lErro = SUCESSO Then
    
        objNFiscal.lNumIntDoc = objNFe.lNumIntDocNF

        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 192929

        Set objTipoDocInfo = New ClassTipoDocInfo
    
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 192930
        If lErro = 31415 Then gError 192931 'Não encontrou
    
'        If giTipoVersao = VERSAO_FULL Then
            Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscal)
'        ElseIf giTipoVersao = VERSAO_LIGHT Then
'            Call Chama_Tela("NFiscalFatura", objNFiscal)
'        End If

    Else
    
        Call Chama_Tela("NFiscalFatura")
    
    End If
    
    Exit Function

Erro_NFeLista_BotaoEdita:

    Select Case gErr

        Case 192929, 192930

        Case 192931
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192932)

    End Select

    Exit Function

End Function

Function ItensNFiscalTodosSaida_Lista_BotaoEdita(objItemNFiscal As ClassItemNF, lErro As Long) As Long

Dim objTipoDocInfo As ClassTipoDocInfo
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_ItensNFiscalTodosSaida_Lista_BotaoEdita

    If lErro = SUCESSO Then
                   
        objNFiscal.lNumIntDoc = objItemNFiscal.lNumIntNF
                
        'lê a Nota Fiscal
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 80491
        If lErro = 31442 Then gError 80494 'Não encontrou
        
        Set objTipoDocInfo = New ClassTipoDocInfo
    
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 80492
        If lErro = 31415 Then gError 80493 'Não encontrou
    
        Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscal)
            
    End If
    
    ItensNFiscalTodosSaida_Lista_BotaoEdita = SUCESSO
    
    Exit Function
    
Erro_ItensNFiscalTodosSaida_Lista_BotaoEdita:

    ItensNFiscalTodosSaida_Lista_BotaoEdita = gErr

    Select Case gErr

        Case 80491, 80492
        
        Case 80493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
            
        Case 80494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149471)

    End Select

    Exit Function

End Function
