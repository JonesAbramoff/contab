VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalFaturaMgz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get BotaoCarta(ByVal objCT As Object) As Object
     Set BotaoCarta = objCT.objUserControl.Controls("BotaoCarta")
End Property

Public Property Get IDCorrespondencia(ByVal objCT As Object) As Object
     Set IDCorrespondencia = objCT.objUserControl.Controls("IDCorrespondencia")
End Property

Public Property Get Contrato(ByVal objCT As Object) As Object
     Set Contrato = objCT.objUserControl.Controls("ContratoMgz")
End Property

Public Property Get Processo(ByVal objCT As Object) As Object
     Set Processo = objCT.objUserControl.Controls("Processo")
End Property

Public Property Get Natureza(ByVal objCT As Object) As Object
     Set Natureza = objCT.objUserControl.Controls("Natureza")
End Property

Public Property Get LabelNatureza(ByVal objCT As Object) As Object
     Set LabelNatureza = objCT.objUserControl.Controls("LabelNatureza")
End Property

Public Property Get LabelNaturezaDesc(ByVal objCT As Object) As Object
     Set LabelNaturezaDesc = objCT.objUserControl.Controls("LabelNaturezaDesc")
End Property

Public Function NFF_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long
    Call objCT.gobjInfoUsu.objEventoNatureza_Nothing
    NFF_Form_Unload = SUCESSO
End Function

Public Function NFF_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFF_Form_Load

    Call objCT.gobjInfoUsu.objEventoNatureza_New
    
    'Inicializa a mascara de Natureza
    lErro = Inicializa_Mascara_Natureza(objCT)
    If lErro <> SUCESSO Then gError 42127
    
    NFF_Form_Load = SUCESSO
    
    Exit Function
    
Erro_NFF_Form_Load:

    NFF_Form_Load = gErr
     
    Select Case gErr
          
        Case 42127
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179287)
     
    End Select
     
    Exit Function
    
End Function

Public Function NFF_Limpa_Tela(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFF_Limpa_Tela
   
    Natureza(objCT).PromptInclude = False
    Natureza(objCT).Text = ""
    Natureza(objCT).PromptInclude = True
    LabelNaturezaDesc(objCT).Caption = ""
    
    NFF_Limpa_Tela = SUCESSO
    
    Exit Function
    
Erro_NFF_Limpa_Tela:

    NFF_Limpa_Tela = gErr
     
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179287)
     
    End Select
     
    Exit Function
    
End Function

Public Function NFF_Grava_Registro(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long

Dim lErro As Long
Dim objTitRecPrevisao As New ClassTituloReceber
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NFF_Grava_Registro

'    'busca titulo com status previsao
'    lErro = CF("NFiscal_BuscaPrevisaoRec", objNFiscal, objTitRecPrevisao)
'    If lErro <> SUCESSO And lErro <> 106975 Then gError 124001
'
'    If lErro = SUCESSO Then
'
'        'pede confirmacao p/excluir
'        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_PREVISAO")
'        If vbMsgRes = vbYes Then
'
'            objTitRecPrevisao.iExclusao = MARCADO
'
'            'exclui previsao
'            lErro = CF("TituloReceber_Exclui", objTitRecPrevisao, objContabil)
'            If lErro <> SUCESSO Then gError 124002
'
'        End If
'
'    End If
    
    NFF_Grava_Registro = SUCESSO
    
    Exit Function
    
Erro_NFF_Grava_Registro:

    NFF_Grava_Registro = gErr
     
    Select Case gErr
    
        Case 124001, 124002
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179287)
     
    End Select
     
    Exit Function
    
End Function


Public Function NFF_Move_Tela_Memoria(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim objNFiscalMgz As New ClassNFiscalMgz

On Error GoTo Erro_NFF_Move_Tela_Memoria

    objNFiscalMgz.sContrato = Contrato(objCT).Text
    objNFiscalMgz.sProcesso = Processo(objCT).Text
    objNFiscalMgz.sIDCorrespondencia = IDCorrespondencia(objCT).Text
    
    Set objNFiscal.objInfoUsu = objNFiscalMgz
    
    sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza(objCT).Text, sNaturezaFormatada, iNaturezaPreenchida)
    If lErro <> SUCESSO Then gError 106960
    
    objNFiscal.sNatureza = sNaturezaFormatada
    
    NFF_Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_NFF_Move_Tela_Memoria:

    NFF_Move_Tela_Memoria = gErr
     
    Select Case gErr
    
        Case 106960
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179287)
     
    End Select
     
    Exit Function
    
End Function

Public Function NFF_Traz_NFiscal_Tela(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal, ByVal objTituloRec As ClassTituloReceber) As Long

Dim lErro As Long
Dim sNaturezaEnxuta As String
Dim objNFiscalMgz As New ClassNFiscalMgz

On Error GoTo Erro_NFF_Traz_NFiscal_Tela

    If Not (objNFiscal.objInfoUsu Is Nothing) Then

        Set objNFiscalMgz = objNFiscal.objInfoUsu
    
        Contrato(objCT).Text = objNFiscalMgz.sContrato
        Processo(objCT).Text = objNFiscalMgz.sProcesso
        IDCorrespondencia(objCT).Text = objNFiscalMgz.sIDCorrespondencia
        
    End If
    
    If Len(Trim(objTituloRec.sNatureza)) > 0 Then
        objNFiscal.sNatureza = objTituloRec.sNatureza
    End If
    
    If Len(Trim(objNFiscal.sNatureza)) <> 0 Then
    
        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objNFiscal.sNatureza, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 106961
    
        Natureza(objCT).PromptInclude = False
        Natureza(objCT).Text = sNaturezaEnxuta
        Natureza(objCT).PromptInclude = True
        
    Else
    
        Natureza(objCT).PromptInclude = False
        Natureza(objCT).Text = ""
        Natureza(objCT).PromptInclude = True
        
    End If
        
    Call Natureza_Validate(objCT, bSGECancelDummy)
    
    NFF_Traz_NFiscal_Tela = SUCESSO
    
    Exit Function
    
Erro_NFF_Traz_NFiscal_Tela:

    NFF_Traz_NFiscal_Tela = gErr
     
    Select Case gErr
    
        Case 106961
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179287)
     
    End Select
     
    Exit Function
    
End Function

Sub Natureza_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Sub LabelNatureza_Click(ByVal objCT As Object)
    Call objCT.gobjInfoUsu.LabelNatureza_Click(objCT, Me)
End Sub

Sub Natureza_Validate(ByVal objCT As Object, Cancel As Boolean)
     
Dim lErro As Long
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_Natureza_Validate

    If Len(Natureza(objCT).ClipText) > 0 Then

        sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)

        'critica o formato da Natureza
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza(objCT).Text, sNaturezaFormatada, iNaturezaPreenchida)
        If lErro <> SUCESSO Then gError 122826
        
        'Obj recebe código
        objNatMovCta.sCodigo = sNaturezaFormatada
        
        'Verifica se a Natureza é analítica e se seu Tipo Corresponde
        lErro = CF("Natureza_Critica", objNatMovCta, NATUREZA_TIPO_RECEBIMENTO)
        If lErro <> SUCESSO Then gError 122843
        
        'Coloca a Descrição da Natureza na Tela
        LabelNaturezaDesc(objCT).Caption = objNatMovCta.sDescricao
        
    Else
    
        LabelNaturezaDesc(objCT).Caption = ""
    
    End If
    
    Exit Sub
    
Erro_Natureza_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122826, 122843
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Sub
    
End Sub

Private Function Inicializa_Mascara_Natureza(ByVal objCT As Object) As Long
'inicializa a mascara da Natureza

Dim sMascaraNatureza As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Natureza

    'Inicializa a máscara da Natureza
    sMascaraNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Armazena em sMascaraNatureza a mascara a ser a ser exibida no campo Natureza
    lErro = MascaraItem(SEGMENTO_NATMOVCTA, sMascaraNatureza)
    If lErro <> SUCESSO Then gError 122836
    
    'coloca a mascara na tela.
    Natureza(objCT).Mask = sMascaraNatureza
    
    Inicializa_Mascara_Natureza = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Natureza:

    Inicializa_Mascara_Natureza = gErr
    
    Select Case gErr
    
        Case 122836
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Function

End Function

Public Sub Contrato_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Contrato_Validate(ByVal objCT As Object, Cancel As Boolean)
'verifica se existe contrato p/o cliente

Dim lErro As Long
Dim objContrato As New ClassContratoMgz
Dim objCliente As New ClassCliente

On Error GoTo Erro_Contrato_Validate

    If Len(Trim(objCT.Cliente.Text)) <> 0 And Len(Trim(Contrato(objCT).Text)) <> 0 Then
    
'        objCliente.sNomeReduzido = objCT.Cliente.Text
'        'Lê o Cliente
'        lErro = CF("Cliente_Le_Codigo_NomeReduzido", objCliente)
'        If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
'
'        objContrato.lCliente = objCliente.lCodigo
'        objContrato.sContrato = Contrato(objCT).Text
'
'        lErro = CF("ContratoMgz_Le", objContrato)
'        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
'        If lErro <> SUCESSO Then gError 99999
    
    End If
    
    Exit Sub
     
Erro_Contrato_Validate:

    Cancel = True
    
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Processo_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Processo_Validate(ByVal objCT As Object, Cancel As Boolean)
'verifica se existe o par (contrato,processo) p/o cliente

Dim lErro As Long
Dim objContrato As New ClassContratoMgz
Dim objCliente As New ClassCliente
Dim bAchou As Boolean
Dim objProc As New ClassProcContratoMgz

On Error GoTo Erro_Processo_Validate

    If Len(Trim(objCT.Cliente.Text)) <> 0 And Len(Trim(Contrato(objCT).Text)) <> 0 And Len(Trim(Processo(objCT).Text)) <> 0 Then
    
'        objCliente.sNomeReduzido = objCT.Cliente.Text
'        'Lê o Cliente
'        lErro = CF("Cliente_Le_Codigo_NomeReduzido", objCliente)
'        If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
'
'        objProc.lCliente = objCliente.lCodigo
'        objProc.sContrato = Contrato(objCT).Text
'        objProc.sProcesso = Processo(objCT).Text
'
'        lErro = CF("ProcessoContrato_Le", objProc)
'        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
'        If lErro <> SUCESSO Then gError 99999
    
    End If
    
    Exit Sub
     
Erro_Processo_Validate:

    Cancel = True
    
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Sub

End Sub

Public Sub IDCorrespondencia_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub IDCorrespondencia_Validate(ByVal objCT As Object, Cancel As Boolean)
    '???
End Sub

Public Sub BotaoCarta_Click(ByVal objCT As Object)

Dim lErro As Long, objNFiscal As New ClassNFiscal
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoCarta_Click

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(objCT.Serie.Text)) = 0 Then gError 57653
    If Len(Trim(objCT.NFiscal.Caption)) = 0 Then gError 57654
    If Len(Trim(objCT.TipoNFiscal.Text)) = 0 Then gError 57655
    If Len(Trim(objCT.DataEmissao.ClipText)) = 0 Then gError 57656

    objNFiscal.sSerie = objCT.Serie.Text
    objNFiscal.lNumNotaFiscal = StrParaLong(objCT.NFiscal.Caption)
    objNFiscal.iTipoNFiscal = Codigo_Extrai(objCT.TipoNFiscal.Text)
    objNFiscal.dtDataEmissao = StrParaDate(objCT.DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa

    'Tenta ler a nota Fiscal passada por parametro
    lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
    If lErro <> SUCESSO And lErro <> 43676 Then gError ERRO_SEM_MENSAGEM
    If lErro = 43676 Then gError 57660

    lErro = objRelatorio.ExecutarDireto("Carta para Nota Fiscal", "", 0, "", "NNUMINTDOC", objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub
    
Erro_BotaoCarta_Click:

    Select Case gErr
    
        Case 57653
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 57654
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)
        
        Case 57655
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 57656
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
    
        Case 57660
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
    
        Case ERRO_SEM_MENSAGEM
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Sub

End Sub

