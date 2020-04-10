VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTFilialEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objTela As Object
Dim m_objUserControl As Object

Event Unload()

Public iAlterado As Integer
Public gobjTabEnd As New ClassTabEndereco

Dim iFrameAtual As Integer
Dim glCodigoAnterior As Long

Dim objGridAdmCad As AdmGrid
Dim iGrid_AdmCadCodigo_Col As Integer
Dim iGrid_AdmCadInscricao_Col As Integer

Dim objGridContrib As AdmGrid
Dim iGrid_ContribUF_Col As Integer
Dim iGrid_ContribIE_Col As Integer

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Modulos = 2
Private Const TAB_Complementos = 3
Private Const TAB_Enderecos = 4
Private Const TAB_Tributacao = 5
Private Const TAB_SPED = 6

'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objTela.Frame1
End Property

Public Property Get Filiais() As Object
     Set Filiais = objTela.Filiais
End Property

Public Property Get CodigoEmpresa() As Object
     Set CodigoEmpresa = objTela.CodigoEmpresa
End Property

Public Property Get EmpresaLabel() As Object
     Set EmpresaLabel = objTela.EmpresaLabel
End Property

Public Property Get Nome() As Object
     Set Nome = objTela.Nome
End Property

Public Property Get Codigo() As Object
     Set Codigo = objTela.Codigo
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objTela.BotaoGravar
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objTela.BotaoExcluir
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objTela.BotaoLimpar
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objTela.BotaoFechar
End Property

Public Property Get Frame9() As Object
     Set Frame9 = objTela.Frame9
End Property

Public Property Get Ramo() As Object
     Set Ramo = objTela.Ramo
End Property

Public Property Get CertificadoA1A3() As Object
     Set CertificadoA1A3 = objTela.CertificadoA1A3
End Property

Public Property Get Frame8() As Object
     Set Frame8 = objTela.Frame8
End Property

Public Property Get DataJucerja() As Object
     Set DataJucerja = objTela.DataJucerja
End Property

Public Property Get UpDownDataJucerja() As Object
     Set UpDownDataJucerja = objTela.UpDownDataJucerja
End Property

Public Property Get Jucerja() As Object
     Set Jucerja = objTela.Jucerja
End Property

Public Property Get Frame7() As Object
     Set Frame7 = objTela.Frame7
End Property

Public Property Get CGC() As Object
     Set CGC = objTela.CGC
End Property

Public Property Get InscricaoEstadual() As Object
     Set InscricaoEstadual = objTela.InscricaoEstadual
End Property

Public Property Get InscricaoMunicipal() As Object
     Set InscricaoMunicipal = objTela.InscricaoMunicipal
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objTela.Frame6
End Property

Public Property Get ContadorCRC() As Object
     Set ContadorCRC = objTela.ContadorCRC
End Property

Public Property Get ContadorNome() As Object
     Set ContadorNome = objTela.ContadorNome
End Property

Public Property Get ContadorCPF() As Object
     Set ContadorCPF = objTela.ContadorCPF
End Property

Public Property Get FilialLabel() As Object
     Set FilialLabel = objTela.FilialLabel
End Property

Public Property Get Frame10() As Object
     Set Frame10 = objTela.Frame10
End Property

Public Property Get AtivTrib() As Object
     Set AtivTrib = objTela.AtivTrib
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objTela.Frame3
End Property

Public Property Get IRPercPadrao() As Object
     Set IRPercPadrao = objTela.IRPercPadrao
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objTela.Frame4
End Property

Public Property Get ICMSPorEstimativa() As Object
     Set ICMSPorEstimativa = objTela.ICMSPorEstimativa
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objTela.Frame5
End Property

Public Property Get IPINaoContrib() As Object
     Set IPINaoContrib = objTela.IPINaoContrib
End Property

Public Property Get IPIContrib() As Object
     Set IPIContrib = objTela.IPIContrib
End Property

Public Property Get IPIContrib50() As Object
     Set IPIContrib50 = objTela.IPIContrib50
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objTela.Frame2
End Property

Public Property Get ISSIncluso() As Object
     Set ISSIncluso = objTela.ISSIncluso
End Property

Public Property Get ISSPercPadrao() As Object
     Set ISSPercPadrao = objTela.ISSPercPadrao
End Property

Public Property Get FrameEndPrincipal() As Object
     Set FrameEndPrincipal = objTela.FrameEndPrincipal
End Property

Public Property Get FrameEndEntrega() As Object
     Set FrameEndEntrega = objTela.FrameEndEntrega
End Property

Public Property Get EndPrincipal() As Object
     Set EndPrincipal = objTela.EndPrincipal
End Property

Public Property Get EndEntrega() As Object
     Set EndEntrega = objTela.EndEntrega
End Property

Public Property Get BotaoMarcarTodos() As Object
     Set BotaoMarcarTodos = objTela.BotaoMarcarTodos
End Property

Public Property Get BotaoDesmarcarTodos() As Object
     Set BotaoDesmarcarTodos = objTela.BotaoDesmarcarTodos
End Property

Public Property Get Modulos() As Object
     Set Modulos = objTela.Modulos
End Property

Public Property Get Opcao() As Object
     Set Opcao = objTela.Opcao
End Property

Public Property Get SimplesFederal() As Object
     Set SimplesFederal = objTela.SimplesFederal
End Property

Public Property Get SimplesFederalAliq() As Object
     Set SimplesFederalAliq = objTela.SimplesFederalAliq
End Property

Public Property Get SimplesFederalTeto() As Object
     Set SimplesFederalTeto = objTela.SimplesFederalTeto
End Property

Public Property Get SuperSimples() As Object
     Set SuperSimples = objTela.SuperSimples
End Property

Public Property Get PISNaoCumulativo() As Object
     Set PISNaoCumulativo = objTela.PISNaoCumulativo
End Property

Public Property Get LucroPresumido() As Object
     Set LucroPresumido = objTela.LucroPresumido
End Property

Public Property Get NFeAmbiente() As Object
     Set NFeAmbiente = objTela.NFeAmbiente
End Property

Public Property Get RPSAmbiente() As Object
     Set RPSAmbiente = objTela.RPSAmbiente
End Property

Public Property Get CNAE() As Object
     Set CNAE = objTela.CNAE
End Property

Public Property Get DescCNAE() As Object
     Set DescCNAE = objTela.DescCNAE
End Property

Public Property Get SignatNome() As Object
     Set SignatNome = objTela.SignatNome
End Property

Public Property Get SignatQualif() As Object
     Set SignatQualif = objTela.SignatQualif
End Property

Public Property Get SignatCPF() As Object
     Set SignatCPF = objTela.SignatCPF
End Property

Public Property Get AdmCadInscricao() As Object
     Set AdmCadInscricao = objTela.AdmCadInscricao
End Property

Public Property Get AdmCadCodigo() As Object
     Set AdmCadCodigo = objTela.AdmCadCodigo
End Property

Public Property Get GridAdmCad() As Object
     Set GridAdmCad = objTela.GridAdmCad
End Property

Public Property Get GridContrib() As Object
     Set GridContrib = objTela.GridContrib
End Property

Public Property Get ContribIE() As Object
     Set ContribIE = objTela.ContribIE
End Property

Public Property Get ContribUF() As Object
     Set ContribUF = objTela.ContribUF
End Property

Public Property Get SpedFiscalPerfil() As Object
     Set SpedFiscalPerfil = objTela.SpedFiscalPerfil
End Property

Public Property Get CNPJCTB() As Object
     Set CNPJCTB = objTela.CNPJCTB
End Property

Public Property Get NomeReduzido() As Object
     Set NomeReduzido = objTela.NomeReduzido
End Property

Public Property Get TabEnd() As Object
     Set TabEnd = objTela.TabEnd
End Property

Public Property Get RegimeTrib() As Object
     Set RegimeTrib = objTela.RegimeTrib
End Property

Public Property Get RegimeEspecialTrib() As Object
     Set RegimeEspecialTrib = objTela.RegimeEspecialTrib
End Property

Public Property Get NFCECSC() As Object
     Set NFCECSC = objTela.NFCECSC
End Property

Public Property Get idNFCECSC() As Object
     Set idNFCECSC = objTela.idNFCECSC
End Property

Public Property Get indSincPadrao() As Object
     Set indSincPadrao = objTela.indSincPadrao
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoDesmarcarTodos_Click()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Error_BotaoDesmarcarTodos_Click
    'percorre todas as linhas da ListBox de Módulos
    For iIndice = 0 To Modulos.ListCount - 1
    
        'marca cada checkbox da ListBox
        Modulos.Selected(iIndice) = False
    
    Next
    
    Exit Sub
    
Error_BotaoDesmarcarTodos_Click:
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155858)
            
    End Select
    
    Exit Sub
End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objEmpresa As New ClassDicEmpresa
Dim vbMsgRet As VbMsgBoxResult
Dim iCodFilial As Integer

On Error GoTo Erro_BotaoExcluir_Click
    
    'Verifica se Códigos de Empresa e Filial foram informados
    If Len(Trim(CodigoEmpresa.Text)) = 0 Then Error 25581
    If Len(Trim(Codigo.Text)) = 0 Then Error 25582
    
    objFilialEmpresa.lCodEmpresa = CLng(CodigoEmpresa.Text)
    objFilialEmpresa.iCodFilial = CInt(Codigo.Text)
    
    'Verifica se a FilialEmpresa existe
    lErro = FilialEmpresa_Le2(objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 25576 Then Error 25583
    If lErro = 25576 Then Error 25584 'FilialEmpresa não cadastrada
    
    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_FILIALEMPRESA", objFilialEmpresa.iCodFilial, objFilialEmpresa.lCodEmpresa)
    If vbMsgRet = vbYes Then
        
        'Exclui a FilialEmpresa do BD
        lErro = FilialEmpresa_Exclui(objFilialEmpresa)
        If lErro <> SUCESSO Then Error 25585
        'Remove FilialEmpresa da TreeView Filiais
        Call FiliaisList_Remove(objFilialEmpresa)
    
        'Limpa a Tela
        Call Limpa_Tela_FilialEmpresa
        
        objEmpresa.lCodigo = objFilialEmpresa.lCodEmpresa
        
        'Calcula nº auto para nova Filial
        lErro = FilialEmpresa_Automatica(objEmpresa, iCodFilial)
        If lErro <> SUCESSO Then Error 25586
        
        'Coloca na tela
        Codigo = CStr(iCodFilial)
    
    End If
Exit Sub
Erro_BotaoExcluir_Click:
    Select Case Err
        Case 25581, 25582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_EMPRESA_NAO_INFORMADO", Err)
        
        Case 25583, 25585, 25586  'Tratados nas rotinas chamadas
        Case 25584
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 155859)
    End Select
    Exit Sub
End Sub
Public Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objEmpresa As New ClassDicEmpresa
Dim iCodFilial As Integer

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Verifica_Dados_Obrigatorios()
    If lErro <> SUCESSO Then Error 25921
    
    lErro = Traz_Dados_Memoria(objFilialEmpresa)
    If lErro <> SUCESSO Then Error 25922
    
    'Grava FilialEmpresa no banco de dados
    lErro = FilialEmpresa_Grava(objFilialEmpresa)
    If lErro Then Error 25572
    
    'Remove e insere FilialEmpresa da TreeView Filiais
    Call FiliaisList_Remove(objFilialEmpresa)
    Call FiliaisList_Adiciona(objFilialEmpresa)
    
    'Limpa a Tela
    Call Limpa_Tela_FilialEmpresa
    
    objEmpresa.lCodigo = objFilialEmpresa.lCodEmpresa
    
    'Calcula nº auto para nova Filial
    lErro = FilialEmpresa_Automatica(objEmpresa, iCodFilial)
    If lErro <> SUCESSO Then Error 25575
    
    'Coloca na tela
    Codigo = CStr(iCodFilial)
    
Exit Sub

Erro_BotaoGravar_Click:
    
    Select Case Err
        
        Case 25572, 25575, 25921, 25922  'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155860)
     
     End Select
     
     Exit Sub

End Sub

Public Function Traz_Dados_Memoria(objFilialEmpresa As AdmFiliais, Optional objEmpresa As ClassDicEmpresa) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objModulo As New AdmModulo
Dim colEndereco As New Collection
Dim objFilialAdmCad As AdmFiliaisAdmCad
Dim objFilialAdmCadAux As AdmFiliaisAdmCad
Dim objFilialContrib As AdmFiliaisContrib
Dim objFilialContribAux As AdmFiliaisContrib
Dim iIndice2 As Integer

On Error GoTo Erro_Traz_Dados_Memoria

    With objFilialEmpresa
    
'        If giTipoVersao = VERSAO_FULL Then
            
            'Passa dados da tela de Empresa e Filial
            .lCodEmpresa = CLng(CodigoEmpresa.Text)
            .iCodFilial = CInt(Codigo.Text)
            .sNome = Trim(Nome.Text)
        
'        ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'            'Passa dados de objEmpresa de Empresa e Filial
'            .lCodEmpresa = objEmpresa.lCodigo
'            .iCodFilial = FILIAL_MATRIZ
'            .sNome = MATRIZ
'
'        End If
    
        'Pega na listBox Modulos as siglas dos Modulos ativos
        .colModulos.Add MODULO_ADM
        
        For iIndice = 0 To Modulos.ListCount - 1
    
            If Modulos.Selected(iIndice) = True Then
    
                objModulo.sSigla = Trim(left(Modulos.List(iIndice), InStr(Modulos.List(iIndice), SEPARADOR) - 1))
                .colModulos.Add objModulo.sSigla
    
            End If
    
        Next
    
        'Passa dados complementares
        .sCgc = Trim(CGC)
        .sInscricaoEstadual = InscricaoEstadual
        .sInscricaoMunicipal = InscricaoMunicipal
        .sJucerja = Jucerja
        .dtDataJucerja = StrParaDate(DataJucerja)
        .sContador = ContadorNome
        .sCPFContador = ContadorCPF
        .sCRCContador = ContadorCRC
        .sRamoEmpresa = Ramo
        .sCertificadoA1A3 = CertificadoA1A3.Text
        .iNFeAmbiente = NFeAmbiente.ItemData(NFeAmbiente.ListIndex)
        
        .iRPSAmbiente = RPSAmbiente.ItemData(RPSAmbiente.ListIndex)
        
        .sCNAE = CNAE.ClipText
        .sCodQualiSigCTB = SCodigo_Extrai(SignatQualif.Text)
        .sSignatarioCTB = SignatNome.Text
        .sCPFSignatarioCTB = SignatCPF.Text
        
    End With
        
    For iIndice = 1 To objGridAdmCad.iLinhasExistentes
    
        Set objFilialAdmCad = New AdmFiliaisAdmCad
    
        objFilialAdmCad.iFilialEmpresa = objFilialEmpresa.iCodFilial
        objFilialAdmCad.sCodigo = SCodigo_Extrai(GridAdmCad.TextMatrix(iIndice, iGrid_AdmCadCodigo_Col))
        objFilialAdmCad.sInscricao = GridAdmCad.TextMatrix(iIndice, iGrid_AdmCadInscricao_Col)
        
        iIndice2 = 0
        For Each objFilialAdmCadAux In objFilialEmpresa.colAdmCad
            iIndice2 = iIndice2 + 1
            If objFilialAdmCadAux.sCodigo = objFilialAdmCad.sCodigo Then gError 202791
        Next
               
        objFilialEmpresa.colAdmCad.Add objFilialAdmCad
    Next
    
    For iIndice = 1 To objGridContrib.iLinhasExistentes
    
        Set objFilialContrib = New AdmFiliaisContrib
    
        objFilialContrib.iFilialEmpresa = objFilialEmpresa.iCodFilial
        objFilialContrib.sUF = GridContrib.TextMatrix(iIndice, iGrid_ContribUF_Col)
        objFilialContrib.sInscricao = GridContrib.TextMatrix(iIndice, iGrid_ContribIE_Col)

        iIndice2 = 0
        For Each objFilialContribAux In objFilialEmpresa.colContrib
            iIndice2 = iIndice2 + 1
            If objFilialContribAux.sUF = objFilialContrib.sUF Then gError 202792
        Next
        
        objFilialEmpresa.colContrib.Add objFilialContrib
    Next
    
    objFilialEmpresa.sSpedFiscalPerfil = SpedFiscalPerfil.Text
    objFilialEmpresa.sCNPJContador = Trim(CNPJCTB.Text)
    objFilialEmpresa.sNomeReduzido = Trim(NomeReduzido.Text)
        
    'Le Endereços da tela
    lErro = gobjTabEnd.Move_Endereco_Memoria(colEndereco)
    If lErro <> SUCESSO Then gError 25571
        
    With objFilialEmpresa
        
        'Passa para objFilialEmpresa
        Set .objEndereco = colEndereco(1)
        Set .objEnderecoEntrega = colEndereco(2)
        Set .objEnderecoContador = colEndereco(3)
        
        'Para Endereço entrega preenchido, critica se Estado e Pais foram informados
        If Len(Trim(.objEnderecoEntrega.sEndereco)) > 0 And Len(Trim(.objEnderecoEntrega.sSiglaEstado)) = 0 Then gError 25574
        If Len(Trim(.objEnderecoEntrega.sEndereco)) > 0 And .objEnderecoEntrega.iCodigoPais = 0 Then gError 25573
        
        'Pega dados de Tributacao
        .dISSPercPadrao = StrParaDbl(ISSPercPadrao) / 100
        .iISSIncluso = ISSIncluso
        .dIRPercPadrao = StrParaDbl(IRPercPadrao) / 100
        .iICMSPorEstimativa = ICMSPorEstimativa
        
        If IPINaoContrib = True Then
            .iContribuinteIPI = IPI_NAO_CONTRIBUINTE
        ElseIf IPIContrib = True Then
            .iContribuinteIPI = IPI_CONTRIBUINTE
        ElseIf IPIContrib50 = True Then
            .iContribuinteIPI = IPI_CONTRIBUINTE_50
        End If
        
        For iIndice = 1 To 3
        
            If AtivTrib(iIndice).Value = True Then .iTipoTribAtividade = iIndice
            
        Next
        
        '#########################################
        'Inserido por Wagner 20/03/2006
        .iSimplesFederal = SimplesFederal.Value
        .dSimplesFederalAliq = StrParaDbl(SimplesFederalAliq.Text) / 100
        .dSimplesFederalTeto = StrParaDbl(SimplesFederalTeto.Text)
        '#########################################
        
        .iSuperSimples = SuperSimples.Value
        .iPISNaoCumulativo = PISNaoCumulativo.Value
        .iCOFINSNaoCumulativo = .iPISNaoCumulativo
        .iLucroPresumido = LucroPresumido.Value
        
        .iRegimeTrib = Codigo_Extrai(RegimeTrib.Text)
        .iRegimeEspecialTrib = Codigo_Extrai(RegimeEspecialTrib.Text)
    
        .sNFCECSC = NFCECSC.Text
        .sidNFCECSC = idNFCECSC.Text
        
        If indSincPadrao.ListIndex = -1 Then
            .iindSincPadrao = -1
        Else
            .iindSincPadrao = indSincPadrao.ItemData(indSincPadrao.ListIndex)
        End If
        
    End With
    
    Traz_Dados_Memoria = SUCESSO

Exit Function

Erro_Traz_Dados_Memoria:
    
    Traz_Dados_Memoria = gErr
    
    Select Case gErr
        
        Case 25571 'tratado na rotina chamada
        
        Case 25573
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_INFORMADO_ENTREGA", gErr)
        
        Case 25574
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_INFORMADO_ENTREGA", gErr)
            
        Case 202791
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMCAD_REPETIDO", gErr, iIndice2, iIndice)
        
        Case 202792
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRIB_REPETIDO", gErr, iIndice2, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155861)
     
     End Select
     
     Exit Function

End Function

Public Function Verifica_Dados_Obrigatorios() As Long

Dim lErro As Long

On Error GoTo Erro_Verifica_Dados_Obrigatorios

    'Verifica se dados obrigatórios de FilialEmpresa foram informados
'    If giTipoVersao = VERSAO_FULL Then
        
        If Len(CodigoEmpresa.Text) = 0 Then gError 25565
        If CLng(CodigoEmpresa.Text) = 0 Then gError 25566
        If Len(Codigo.Text) = 0 Then gError 25567
        If CLng(Codigo.Text) = 0 Then gError 25568
        If Len(Trim(Nome.Text)) = 0 Then gError 25569
    
'    End If
    
    If Len(Trim(CGC.Text)) = 0 Then gError 25570
'    If Estado(Estado.LBound).ListIndex = -1 Then Error 25603
'    If Pais(Pais.LBound).ListIndex = -1 Then Error 25604

    'se está no Simples nao pode estar marcado ICMS por estimativa
    If SuperSimples.Value = vbChecked Then
        If ICMSPorEstimativa.Value = vbChecked Then gError 209760
    End If

    If Len(Trim(SpedFiscalPerfil.Text)) = 0 Then gError 216092

    Verifica_Dados_Obrigatorios = SUCESSO

Exit Function

Erro_Verifica_Dados_Obrigatorios:
    
    Verifica_Dados_Obrigatorios = gErr
    
    Select Case gErr
        
        Case 25565, 25566
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_EMPRESA_NAO_INFORMADO", gErr)
        
        Case 25567, 25568
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_FILIALEMPRESA_NAO_INFORMADO", gErr)
        
        Case 25569
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_FILIALEMPRESA_NAO_INFORMADO", gErr)
        
        Case 25570
            Call Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_INFORMADO", gErr)
        
        Case 25603
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_INFORMADO_PRINCIPAL", gErr)
        
        Case 25604
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_INFORMADO_PRINCIPAL", gErr)
        
        Case 209760
            Call Rotina_Erro(vbOKOnly, "ERRO_SIMPLES_COM_ICMS_POR_ESTIMATIVA", gErr)
            
        Case 216092
            Call Rotina_Erro(vbOKOnly, "ERRO_PERFILSPED_NAO_INFORMADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155862)
     
     End Select
     
     Exit Function

End Function

Private Sub FiliaisList_Remove(objFilialEmpresa As AdmFiliais)
'Remove FilialEmpresa caso ela exista na TreeView Filiais
Dim objNode As Node
Dim sKeyEmpresa, sKeyFilial As String
    sKeyEmpresa = KEY_CARACTER & CStr(objFilialEmpresa.lCodEmpresa)
    sKeyFilial = sKeyEmpresa & SEPARADOR & CStr(objFilialEmpresa.iCodFilial)
    
    For Each objNode In Filiais.Nodes
    
        If objNode.Key = sKeyFilial Then
    
            Call Filiais.Nodes.Remove(objNode.Key)
            Exit For
    
        End If
    
    Next
End Sub

Private Sub FiliaisList_Adiciona(objFilialEmpresa As AdmFiliais)
'Inclui FilialEmpresa na TreeView Filiais
Dim sKeyEmpresa, sKeyFilial As String
    
    sKeyEmpresa = KEY_CARACTER & CStr(objFilialEmpresa.lCodEmpresa)
    sKeyFilial = sKeyEmpresa & SEPARADOR & CStr(objFilialEmpresa.iCodFilial)
    
    Call Filiais.Nodes.Add(sKeyEmpresa, tvwChild, sKeyFilial, objFilialEmpresa.sNome)
End Sub
'
'Private Function Le_Dados_Enderecos(colEndereco As Collection) As Long
''Lê os dados relativos ao endereco(0 = principal , 1 = entrega) da FilialEmpresa e coloca em colEndereco
'Dim iIndice As Integer
'Dim objEndereco As ClassEndereco
'Dim lErro As Long
'On Error GoTo Erro_Le_Dados_Enderecos
'
'    For iIndice = Endereco.LBound To Endereco.UBound
'
'        Set objEndereco = New ClassEndereco
'
'        objEndereco.sEndereco = Trim(Endereco(iIndice).Text)
'        objEndereco.sBairro = Trim(Bairro(iIndice).Text)
'        objEndereco.sCidade = Trim(Cidade(iIndice).Text)
'        objEndereco.sCEP = Trim(CEP(iIndice).Text)
'        objEndereco.sSiglaEstado = Trim(Estado(iIndice).Text)
'
'        If Len(Trim(Pais(iIndice).Text)) = 0 Then
'            objEndereco.iCodigoPais = 0
'        Else
'            objEndereco.iCodigoPais = Codigo_Extrai(Pais(iIndice).Text)
'        End If
'
'        objEndereco.sTelefone1 = Trim(Telefone1(iIndice).Text)
'        objEndereco.sTelefone2 = Trim(Telefone2(iIndice).Text)
'        objEndereco.sFax = Trim(Fax(iIndice).Text)
'        objEndereco.sEmail = Trim(Email(iIndice).Text)
'        objEndereco.sContato = Trim(Contato(iIndice).Text)
'
'        colEndereco.Add objEndereco
'
'    Next
'    Le_Dados_Enderecos = SUCESSO
'
'    Exit Function
'Erro_Le_Dados_Enderecos:
'
'    Le_Dados_Enderecos = Err
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155863)
'    End Select
'    Exit Function
'End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long
Dim objEmpresa As New ClassDicEmpresa
Dim iCodFilial As Integer

On Error GoTo Erro_BotaoLimpar_Click
    
    'Limpa a Tela
     Call Limpa_Tela_FilialEmpresa

     'Se código de Empresa estiver preenchido
     If Len(Trim(CodigoEmpresa)) > 0 Then
     
        objEmpresa.lCodigo = CLng(CodigoEmpresa)
        
        'Calcula nº auto para nova Filial
        lErro = FilialEmpresa_Automatica(objEmpresa, iCodFilial)
        If lErro <> SUCESSO Then Error 25580
        
        'Coloca na tela
        Codigo = CStr(iCodFilial)
        
    End If
    
     Exit Sub
Erro_BotaoLimpar_Click:
    
    Select Case Err
        Case 25580 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155864)
    End Select
    Exit Sub
End Sub

Public Sub BotaoMarcarTodos_Click()
Dim iIndice As Integer
Dim lErro As Long
On Error GoTo Error_BotaoMarcarTodos_Click
    'percorre todas as linhas da ListBox de Módulos
    For iIndice = 0 To Modulos.ListCount - 1
    
        'marca cada checkbox da ListBox
        Modulos.Selected(iIndice) = True
    
    Next
    
    Modulos.Refresh
    
    Exit Sub
    
Error_BotaoMarcarTodos_Click:
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155865)
            
    End Select
    
    Exit Sub
End Sub

'Public Sub CEP_GotFocus(Index As Integer)
'    If Me.ActiveControl Is CEP(0) Then
'        Call MaskEdBox_TrataGotFocus(CEP(0))
'    Else
'        Call MaskEdBox_TrataGotFocus(CEP(1))
'    End If
'
'End Sub

Public Sub CGC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGC)
End Sub

Public Sub CGC_Validate(Cancel As Boolean)
Dim lErro As Long
On Error GoTo Erro_CGC_Validate
    If Len(Trim(CGC.Text)) = 0 Then Exit Sub
    
    If Len(Trim(CGC.Text)) <> STRING_CGC Then Error 25559
    lErro = Cgc_Critica(CGC.Text)
    If lErro <> SUCESSO Then Error 25556
    Exit Sub
Erro_CGC_Validate:
    
    Cancel = True
    
    Select Case Err
        Case 25556  'tratado na rotina chamada
        Case 25559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155866)
    End Select
    Exit Sub
End Sub

Public Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo)
End Sub

Public Sub Codigo_Validate(Cancel As Boolean)
Dim lErro As Long
On Error GoTo Erro_Codigo_Validate
    
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub
    lErro = Inteiro_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 54808
    
    Exit Sub
Erro_Codigo_Validate:
    Cancel = True
    
    Select Case Err
    
        Case 54808 'Erro tratado na rotina chamadora
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155867)
    
    End Select
    Exit Sub
    
End Sub

Public Sub CodigoEmpresa_GotFocus()
    'Guarda código anterior
    glCodigoAnterior = StrParaLong(CodigoEmpresa)
    Call MaskEdBox_TrataGotFocus(CodigoEmpresa)
    
End Sub

Public Sub CodigoEmpresa_Validate(Cancel As Boolean)
Dim lErro As Long
Dim colModulos As New Collection
Dim objModulo As AdmModulo
Dim objEmpresa As New ClassDicEmpresa
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
On Error GoTo Erro_CodigoEmpresa_Validate
    'Se estiver vazio
    If Len(Trim(CodigoEmpresa)) = 0 Then
        
        'Limpa ListBox de Modulos
        Modulos.Clear
        
        'Limpa labels de Empresa
        For iIndice = EmpresaLabel.LBound To EmpresaLabel.UBound
            EmpresaLabel(iIndice).Caption = ""
        Next
        
        lErro = Limpa_Tela_FilialEmpresa
        If lErro <> SUCESSO Then Error 57500
        
        Exit Sub
    
    End If
    
    'Critica se é inteiro positivo
    lErro = Inteiro_Critica(CodigoEmpresa)
    If lErro <> SUCESSO Then Error 25530
    
    objEmpresa.lCodigo = CLng(CodigoEmpresa)
    
    'Tenta ler Empresa
    lErro = Empresa_Le(objEmpresa)
    If lErro <> SUCESSO And lErro <> 6179 Then Error 25534
    'Se não existe Empresa no BD
    If lErro = 6179 Then Error 25536
    'Se houve alteração no código da Empresa
    If objEmpresa.lCodigo <> glCodigoAnterior Then
            
        'Traz dados da Empresa para tela sem limpar o resto
        lErro = Traz_Empresa_Tela(objEmpresa, TELA_LIMPA)
        If lErro <> SUCESSO Then Error 25597
    
    End If
    
    Exit Sub
Erro_CodigoEmpresa_Validate:
    Cancel = True
    
    Select Case Err
        Case 25534, 25597  'tratado na rotina chamada
        Case 25530
        
        Case 25536
         
            'Pergunta se deseja criar nova Empresa
             vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_EMPRESA", Trim(CodigoEmpresa.Text))
    
            If vbMsgRes = vbYes Then
    
                EmpresaTela.Show
    
            Else
                'Segura o foco
                CodigoEmpresa.SetFocus
    
            End If
    
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155868)
    End Select
    Exit Sub
    
End Sub

Function Carrega_ListModulosEmpresa(objEmpresa As ClassDicEmpresa) As Long

'carrega módulos ATIVOS da Empresa na LIST e os SELECIONA
Dim lErro As Long
Dim colModulos As New Collection
Dim objModulo As AdmModulo
On Error GoTo Erro_Carrega_ListModulosEmpresa
    
    'Limpa a list
    Modulos.Clear
    
    'Lê módulos associados a Empresa
    lErro = Modulos_Le_Empresa1(objEmpresa.lCodigo, colModulos)
    If lErro <> SUCESSO Then Error 25532
    
    'Coloca os módulos na ListBox Modulos
    For Each objModulo In colModulos
      
        If objModulo.iAtivo = 1 And objModulo.sSigla <> MODULO_ADM Then
            
            'Adiciona na listBox
            Modulos.AddItem objModulo.sSigla & SEPARADOR & objModulo.sNome
            'Seleciona todos os módulos
            Modulos.Selected(Modulos.NewIndex) = True
        
        End If
    
    Next
    
    Carrega_ListModulosEmpresa = SUCESSO
    
    Exit Function

Erro_Carrega_ListModulosEmpresa:
    
    Carrega_ListModulosEmpresa = Err
    
    Select Case Err
        
        Case 25532 'tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155869)
    
    End Select
    
    Exit Function

End Function

Public Function Limpa_Tela_FilialEmpresa()

'Limpa a tela
Dim lErro As Long
Dim iIndice, iIndice2 As Integer
Dim sCodEmpresa As String
Dim objEmpresa As New ClassDicEmpresa
Dim iCodFilial As Integer

On Error GoTo Erro_Limpa_Tela_FilialEmpresa
    
'    If giTipoVersao = VERSAO_FULL Then
    
        'Guarda codigo de Empresa
        sCodEmpresa = Trim(CodigoEmpresa)
    
'    End If
    
    'Limpa a Tela
    lErro = Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridAdmCad)
    Call Grid_Limpa(objGridContrib)
    
    DescCNAE.Caption = ""
    
    RPSAmbiente.ListIndex = 1
    
    SignatQualif.ListIndex = -1
    SpedFiscalPerfil.ListIndex = -1
    
    RegimeEspecialTrib.ListIndex = -1
    RegimeTrib.ListIndex = -1
    indSincPadrao.ListIndex = -1
        
'    If giTipoVersao = VERSAO_FULL Then
    
        'Recoloca Codigo de Empresa
        CodigoEmpresa = sCodEmpresa
    
        'Limpa labels de Filial
        For iIndice = FilialLabel.LBound To FilialLabel.UBound
            FilialLabel(iIndice).Caption = ""
        Next
     
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        'Limpa labels de Empresa
'        For iIndice = EmpresaLabel.LBound To EmpresaLabel.UBound
'            EmpresaLabel(iIndice).Caption = ""
'        Next
'
'    End If
     
     
    'Limpa checkBoxes e OptionButtons
    ISSIncluso.Value = vbChecked
    ICMSPorEstimativa.Value = vbUnchecked
    IPIContrib.Value = False
    IPINaoContrib.Value = True
    IPIContrib50.Value = False
    AtivTrib(1).Value = True
    
    'Traz p/ frente endereço principal
    EndEntrega.Value = False
    EndPrincipal.Value = True
'
'    'Desseleciona Estado
'    For iIndice = Estado.LBound To Estado.UBound
'        Estado(iIndice).Text = ""
'    Next
'
'    'Seleciona Brasil se Combo de Paises tiver List preenchida
'    For iIndice = Pais.LBound To Pais.UBound
'
'        If Pais(iIndice).ListCount > 0 Then
'
'            For iIndice2 = 0 To Pais(iIndice).ListCount - 1
'
'                If right(Pais(iIndice).List(iIndice2), 6) = "Brasil" Then
'                    Pais(iIndice).ListIndex = iIndice2
'                    Exit For
'                End If
'
'            Next
'
'        End If
'
'    Next
    
'    If giTipoVersao = VERSAO_FULL Then
    
        'Se código de Empresa estiver vazio
        If Len(sCodEmpresa) = 0 Then
        
            'Limpa list de Módulos
            Modulos.Clear
            
        Else 'código de Empresa está preenchido
            
            objEmpresa.lCodigo = CLng(sCodEmpresa)
            
            'Calcula nº auto para nova Filial
            lErro = FilialEmpresa_Automatica(objEmpresa, iCodFilial)
            If lErro <> SUCESSO Then Error 25602
        
            'Coloca na tela
            Codigo = CStr(iCodFilial)
            
            'Seleciona os Módulos nas check boxes
            For iIndice = 0 To Modulos.ListCount - 1
                Modulos.Selected(iIndice) = True
            Next
            
            'Desseleciona (não fica faixa azul)
            Modulos.ListIndex = -1
            
        End If
        
'    End If
    
    SimplesFederal.Value = vbUnchecked
    SuperSimples.Value = vbUnchecked
    PISNaoCumulativo = vbChecked
    LucroPresumido = vbUnchecked
    
    Call gobjTabEnd.Limpa_Tela
    
    Limpa_Tela_FilialEmpresa = SUCESSO
    
    Exit Function

Erro_Limpa_Tela_FilialEmpresa:
    
    Limpa_Tela_FilialEmpresa = Err
    
    Select Case Err
        Case 25602  'tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155870)
    End Select
    Exit Function
End Function
Public Sub ContadorCPF_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ContadorCPF)
End Sub
Public Sub ContadorCPF_Validate(Cancel As Boolean)
Dim lErro As Long
On Error GoTo Erro_ContadorCPF_Validate
    If Len(Trim(ContadorCPF.Text)) = 0 Then Exit Sub
    
    If Len(Trim(ContadorCPF.Text)) <> STRING_CPF Then Error 25558
    
    lErro = Cpf_Critica(ContadorCPF.Text)
    If lErro <> SUCESSO Then Error 25557
    Exit Sub
Erro_ContadorCPF_Validate:
    Cancel = True
    
    Select Case Err
        Case 25557  'tratado na rotina chamada
    
        Case 25558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CPF", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155871)
    End Select
    Exit Sub
End Sub
Public Sub DataJucerja_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataJucerja)
End Sub
Public Sub DataJucerja_Validate(Cancel As Boolean)
Dim lErro As Long
On Error GoTo Erro_DataJucerja_Validate
    'Verifica se foi preenchida a Data
    If Len(Trim(DataJucerja.ClipText)) = 0 Then Exit Sub
    'Critica a Data
    lErro = Data_Critica(DataJucerja.Text)
    If lErro <> SUCESSO Then Error 25598
    Exit Sub
Erro_DataJucerja_Validate:
    Cancel = True
    
    Select Case Err
    Case 25598
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155872)
    End Select
    Exit Sub
    
End Sub
Public Sub EndEntrega_Click()
    
    FrameEndPrincipal.Visible = False
    FrameEndEntrega.Visible = True
    
End Sub
Public Sub EndPrincipal_Click()
    FrameEndPrincipal.Visible = True
    FrameEndEntrega.Visible = False
End Sub

Public Sub Filiais_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim objEmpresa As New ClassDicEmpresa
Dim lPosicao As Long
Dim objFilialEmpresa As New AdmFiliais
Dim iCodFilial As Integer

On Error GoTo Erro_Filiais_NodeClick
    
    'Se for nó de 1º nível (Empresa)
    If Node.Parent Is Nothing Then
    
        objEmpresa.lCodigo = CLng(Mid(Node.Key, 2))
    
        'Traz dados da Empresa para tela
        lErro = Traz_Empresa_Tela(objEmpresa)
        If lErro <> SUCESSO Then Error 25540
        
        'Calcula nº auto para nova Filial
        lErro = FilialEmpresa_Automatica(objEmpresa, iCodFilial)
        If lErro <> SUCESSO Then Error 25552
        
        'Coloca na tela
        Codigo = CStr(iCodFilial)
        
    Else 'nó é de 2º nível (FilialEmpresa)
    
        'Extrai código da Empresa
        lPosicao = InStr(Node.Key, SEPARADOR)
        objEmpresa.lCodigo = CLng(Mid(Node.Key, 2, lPosicao - 2))
        
        'Traz dados da Empresa para tela
        lErro = Traz_Empresa_Tela(objEmpresa)
        If lErro <> SUCESSO Then Error 25541
        
        'Extrai código e nome da Filial
        objFilialEmpresa.iCodFilial = CInt(Mid(Node.Key, lPosicao + 1))
        objFilialEmpresa.sNome = Node.Text
        objFilialEmpresa.lCodEmpresa = objEmpresa.lCodigo
        
        'Traz dados da FilialEmpresa para tela (inclusive seleção dos módulos)
        lErro = Traz_Filial_Tela(objFilialEmpresa)
        If lErro <> SUCESSO Then Error 25551
        
    End If
    
    Exit Sub
Erro_Filiais_NodeClick:
    Select Case Err
        Case 25540, 25541, 25551, 25552  'tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155874)
    End Select
    Exit Sub
End Sub
Public Sub IRPercPadrao_Validate(Cancel As Boolean)
Dim lErro As Long
On Error GoTo Erro_IRPercPadrao_Validate
    'Verifica se foi preenchido o percentual
    If Len(Trim(IRPercPadrao.Text)) = 0 Then Exit Sub
    'Critica se é porcentagem
    lErro = Porcentagem_Critica(IRPercPadrao.Text)
    If lErro <> SUCESSO Then Error 25601
    'Formata
    IRPercPadrao.Text = Format(IRPercPadrao.Text, "Fixed")
    Exit Sub
Erro_IRPercPadrao_Validate:
    Cancel = True
    
    Select Case Err
        Case 25601 'Tratado na rotina chamada
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155875)
    
    End Select
    Exit Sub
End Sub
Public Sub ISSPercPadrao_Validate(Cancel As Boolean)
Dim lErro As Long
On Error GoTo Erro_ISSPercPadrao_Validate
    'Verifica se foi preenchido o percentual
    If Len(Trim(ISSPercPadrao.Text)) = 0 Then Exit Sub
    'Critica se é porcentagem
    lErro = Porcentagem_Critica(ISSPercPadrao.Text)
    If lErro <> SUCESSO Then Error 25600
    'Formata
    ISSPercPadrao.Text = Format(ISSPercPadrao.Text, "Fixed")
    Exit Sub
Erro_ISSPercPadrao_Validate:
    Cancel = True
    
    Select Case Err
        Case 25600 'Tratado na rotina chamada
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155876)
    
    End Select
    Exit Sub
End Sub
Public Sub Nome_Validate(Cancel As Boolean)
Dim lErro As Long
Dim iIndice As Integer
On Error GoTo Erro_Nome_Validate
    'Se estiver vazio
    If Len(Trim(Nome)) = 0 Then
        
        'Limpa labels de Filial
        For iIndice = FilialLabel.LBound To FilialLabel.UBound
            FilialLabel(iIndice).Caption = ""
        Next
        
    Else 'foi preenchido
    
        'Preenche labels de Filial
        For iIndice = FilialLabel.LBound To FilialLabel.UBound
            FilialLabel(iIndice).Caption = Nome
        Next
    
    End If
    
    Exit Sub
Erro_Nome_Validate:
    Cancel = True
    
    Select Case Err
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155877)
    End Select
    Exit Sub
    
End Sub

Public Sub Opcao_Click()
    
    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then
        
        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me, 0) <> SUCESSO Then Exit Sub
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
    
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                FilialEmpresa.HelpContextID = IDH_FILIAL_EMPRESA_ID
                
            Case TAB_Modulos
                FilialEmpresa.HelpContextID = IDH_FILIAL_EMPRESA_MODULOS
                        
            Case TAB_Complementos
                FilialEmpresa.HelpContextID = IDH_FILIAL_EMPRESA_COMPLEM
                
            Case TAB_Enderecos
                FilialEmpresa.HelpContextID = IDH_FILIAL_EMPRESA_ENDERECOS
                
            Case TAB_Tributacao
                FilialEmpresa.HelpContextID = IDH_FILIAL_EMPRESA_TRIBUTACAO
                
        End Select

    
    End If

End Sub

Public Function Traz_Filial_Tela(objFilialEmpresa As AdmFiliais) As Long

'Le no BD e traz dados da FilialEmpresa para tela
Dim lErro As Long
Dim iIndice As Integer
Dim objEndereco As ClassEndereco
Dim colModuloFilEmp As Collection 'coleção para ler registros em ModuloFilEmp
Dim objModuloFilEmp As ClassModuloFilEmp
Dim objModulo As New AdmModulo
Dim bCancel As Boolean
Dim objFilialAdmCad As AdmFiliaisAdmCad
Dim objFilialContrib As AdmFiliaisContrib
Dim colEnderecos As New Collection

On Error GoTo Erro_Traz_Filial_Tela
    
    'Faz a leitura de FilialEmpresa (aqui necessita estar ativado o SGE)
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
    If lErro <> SUCESSO And lErro <> 27378 Then Error 25547
    If lErro = 27378 Then Error 25548
    
    lErro = CF("SpedCadastro_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then Error 25547
    
    lErro = CF("ContribuinteSubstituto_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then Error 25547
    
    'Coloca dados da Filial na tela
    With objFilialEmpresa
    
'        If giTipoVersao = VERSAO_FULL Then
            
            Codigo = CStr(.iCodFilial)
            Nome = CStr(.sNome)
    
            For iIndice = FilialLabel.LBound To FilialLabel.UBound
                FilialLabel(iIndice) = CStr(.sNome)
            Next
            
'        End If
        
        CNAE.Text = objFilialEmpresa.sCNAE
        Call CNAE_Validate(bSGECancelDummy)
            
        CGC = .sCgc
        InscricaoEstadual = .sInscricaoEstadual
        InscricaoMunicipal = .sInscricaoMunicipal
        Jucerja = .sJucerja
        DataJucerja = IIf(.dtDataJucerja = DATA_NULA, "  /  /  ", Format(.dtDataJucerja, "dd/mm/yy"))
        ContadorNome = .sContador
        ContadorCPF = .sCPFContador
        ContadorCRC = .sCRCContador
        Ramo = .sRamoEmpresa
        CertificadoA1A3.Text = .sCertificadoA1A3
        For iIndice = 0 To NFeAmbiente.ListCount - 1
            If NFeAmbiente.ItemData(iIndice) = .iNFeAmbiente Then
                NFeAmbiente.ListIndex = iIndice
                Exit For
            End If
        Next
        For iIndice = 0 To RPSAmbiente.ListCount - 1
            If RPSAmbiente.ItemData(iIndice) = .iRPSAmbiente Then
                RPSAmbiente.ListIndex = iIndice
                Exit For
            End If
        Next
        CNPJCTB.Text = .sCNPJContador
        Call CF("SCombo_Seleciona2", SpedFiscalPerfil, .sSpedFiscalPerfil)
        NomeReduzido.Text = .sNomeReduzido
        
        colEnderecos.Add .objEndereco
        colEnderecos.Add .objEnderecoEntrega
        If .objEnderecoContador.iCodigoPais = 0 Then
            .objEnderecoContador.iCodigoPais = 1
            .objEnderecoContador.sSiglaEstado = .objEndereco.sSiglaEstado
        End If
        colEnderecos.Add .objEnderecoContador
         
        lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
        If lErro <> SUCESSO Then Error 25547

        ISSPercPadrao = CStr(.dISSPercPadrao * 100)
        ISSIncluso = .iISSIncluso
        ICMSPorEstimativa = .iICMSPorEstimativa
        IRPercPadrao = CStr(.dIRPercPadrao * 100)
        
        Select Case .iContribuinteIPI
            Case IPI_NAO_CONTRIBUINTE
                IPINaoContrib = True
            Case IPI_CONTRIBUINTE
                IPIContrib = True
            Case IPI_CONTRIBUINTE_50
                IPIContrib50 = True
        End Select
   
        If .iTipoTribAtividade <> 0 Then
            AtivTrib(.iTipoTribAtividade).Value = True
        Else
            AtivTrib(1).Value = True
        End If
           
        SimplesFederal.Value = .iSimplesFederal
        
        If .dSimplesFederalAliq <> 0 Then
            SimplesFederalAliq.Text = CStr(.dSimplesFederalAliq * 100)
        End If
        
        If .dSimplesFederalTeto <> 0 Then
            SimplesFederalTeto.Text = Format(.dSimplesFederalTeto, "STANDARD")
        End If
        
        Call SimplesFederal_Click
           
        SuperSimples.Value = .iSuperSimples
        LucroPresumido.Value = .iLucroPresumido
        PISNaoCumulativo.Value = .iPISNaoCumulativo
        
        SignatNome.Text = .sSignatarioCTB
        SignatCPF.Text = .sCPFSignatarioCTB
        
        Call CF("SCombo_Seleciona2", SignatQualif, .sCodQualiSigCTB)
        
        Call Combo_Seleciona_ItemData(RegimeTrib, .iRegimeTrib)
        Call Combo_Seleciona_ItemData(RegimeEspecialTrib, .iRegimeEspecialTrib)
        
        NFCECSC.Text = .sNFCECSC
        idNFCECSC.Text = .sidNFCECSC
        
        Call Combo_Seleciona_ItemData(indSincPadrao, .iindSincPadrao)
       
        iIndice = 0
        For Each objFilialAdmCad In .colAdmCad
            iIndice = iIndice + 1
            
            Call CF("SCombo_Seleciona2", AdmCadCodigo, objFilialAdmCad.sCodigo)
            
            GridAdmCad.TextMatrix(iIndice, iGrid_AdmCadCodigo_Col) = AdmCadCodigo.Text
            GridAdmCad.TextMatrix(iIndice, iGrid_AdmCadInscricao_Col) = objFilialAdmCad.sInscricao

        Next
        objGridAdmCad.iLinhasExistentes = .colAdmCad.Count
        
        iIndice = 0
        For Each objFilialContrib In .colContrib
            iIndice = iIndice + 1
            
            Call CF("SCombo_Seleciona2", ContribUF, objFilialContrib.sUF)
            
            GridContrib.TextMatrix(iIndice, iGrid_ContribUF_Col) = ContribUF.Text
            GridContrib.TextMatrix(iIndice, iGrid_ContribIE_Col) = objFilialContrib.sInscricao

        Next
        objGridContrib.iLinhasExistentes = .colContrib.Count
        
    End With
    
'    If giTipoVersao = VERSAO_FULL Then
    
        'Seleciona na list Modulos os ativos nessa Filial
        lErro = Seleciona_Modulos_Filial(objFilialEmpresa)
        If lErro <> SUCESSO Then Error 25550
    
'    End If
       
    Traz_Filial_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Filial_Tela:
    
    Traz_Filial_Tela = Err
    
    Select Case Err
        Case 25547, 25550   'tratado na rotina chamada
        
        Case 25548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA1", Err, objFilialEmpresa.lCodEmpresa, objFilialEmpresa.iCodFilial)
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155879)
    End Select
    Exit Function
End Function
Private Function Seleciona_Modulos_Filial(objFilialEmpresa As AdmFiliais) As Long
'Seleciona Modulos ativos nessa filial na ListBox Modulos

Dim lErro As Long
Dim lPosicao As Long
Dim colModuloFilEmp As New Collection
Dim objModuloFilEmp As ClassModuloFilEmp
Dim iIndice As Integer
Dim sSiglaModulo As String

On Error GoTo Erro_Seleciona_Modulos_Filial
    
    'Lê registros de ModuloFilEmp vinculados a Filial Empresa
    lErro = ModuloFilEmp_Le_EmpresaFilial(objFilialEmpresa.lCodEmpresa, objFilialEmpresa.iCodFilial, colModuloFilEmp)
    If lErro <> SUCESSO Then Error 25549
    
    For iIndice = 0 To Modulos.ListCount - 1
    
        'Desseleciona o Modulo na List
        Modulos.Selected(iIndice) = False
        
        'Extrai sigla do Módulo da ListBox Modulos
        sSiglaModulo = Modulos.List(iIndice)
        lPosicao = InStr(sSiglaModulo, SEPARADOR)
        sSiglaModulo = Mid(sSiglaModulo, 1, lPosicao - 1)
        
        'Compara com os lidos no BD
        For Each objModuloFilEmp In colModuloFilEmp
    
            'Se módulo na listBox coincide com módulo lido no BD, seleciona
            If sSiglaModulo = objModuloFilEmp.sSiglaModulo Then
            
                Modulos.Selected(iIndice) = True
                Exit For
                
            End If
    
        Next
    
    Next
    
    Seleciona_Modulos_Filial = SUCESSO
    
    Exit Function
Erro_Seleciona_Modulos_Filial:
    Seleciona_Modulos_Filial = Err
    
    Select Case Err
        Case 25549  'tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155880)
    End Select
    Exit Function
End Function
Private Function Traz_Empresa_Tela(objEmpresa As ClassDicEmpresa, Optional iLimpa As Integer = TELA_LIMPA) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_Empresa_Tela
    
    'Tenta ler Empresa
    lErro = Empresa_Le(objEmpresa)
    If lErro <> SUCESSO And lErro <> 6179 Then Error 25535
    'Se não existe Empresa no BD
    If lErro = 6179 Then Error 25537
    
    'Limpa a Tela
    If iLimpa = TELA_LIMPA Then
        lErro = Limpa_Tela_FilialEmpresa()
        If lErro <> SUCESSO Then Error 25538
    End If
    
    'Para poder fazer transacao no bd da empresa
    lErro = Sistema_DefEmpresa(objEmpresa.sNome, objEmpresa.lCodigo, EMPRESA_TODA_NOME, EMPRESA_TODA)
    If lErro <> AD_BOOL_TRUE Then Error 25446
    
    'Traz dados da Empresa para tela
    CodigoEmpresa = CStr(objEmpresa.lCodigo)
    
    For iIndice = EmpresaLabel.LBound To EmpresaLabel.UBound
        EmpresaLabel(iIndice).Caption = CStr(objEmpresa.sNomeReduzido)
    Next
        
    'Carrega e seleciona módulos ATIVOS associados a Empresa
    lErro = Carrega_ListModulosEmpresa(objEmpresa)
    If lErro <> SUCESSO Then Error 25539
    
    Traz_Empresa_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Empresa_Tela:
    
    Traz_Empresa_Tela = Err
    
    Select Case Err
        Case 25535, 25538, 25539, 25446 'tratado na rotina chamada
    
        Case 25537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_CADASTRADA", Err, objEmpresa.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155881)
    End Select
    Exit Function
End Function
    
Public Sub Form_Unload(Cancel As Integer)
    Set objGridAdmCad = Nothing
    Set objGridContrib = Nothing
    Call gobjTabEnd.Finaliza
    Set gobjTabEnd = Nothing
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objEmpresa As ClassDicEmpresa
Dim lCodigo As Long
Dim colFilialEmpresa As New Collection
Dim objFilialEmpresa As ClassFilialEmpresa
Dim objNode As Node
Dim sKeyEmpresa As String
Dim sKeyFilial As String
Dim lEmpresaAnterior As Long
Dim colEmpresa As New Collection
Dim objTelaAux As Object

On Error GoTo Erro_FilialEmpresa_Form_Load

    Set objGridAdmCad = New AdmGrid
    Set objGridContrib = New AdmGrid
    
    FilialEmpresa.HelpContextID = IDH_FILIAL_EMPRESA_ID
    
    'Começa com Frame 1
    iFrameAtual = 1
    
    'Lê Empresas
    lErro = Empresas_Le_Todas(colEmpresa)
    If lErro <> SUCESSO And lErro <> 6179 Then Error 25523
    If lErro = 6179 Then Error 25524
    
    'Preenche TreeView com Empresas
    For Each objEmpresa In colEmpresa
        
        sKeyEmpresa = KEY_CARACTER & CStr(objEmpresa.lCodigo)
        Set objNode = Filiais.Nodes.Add(, , sKeyEmpresa, objEmpresa.sNomeReduzido)
        objNode.Sorted = True
    
    Next
    
    'Lê dados de FiliaisEmpresas no BD
    lErro = FiliaisEmpresas_Le(colFilialEmpresa)
    If lErro <> SUCESSO Then Error 6189
    
    'Preenche TreeView com Filiais
    For Each objFilialEmpresa In colFilialEmpresa
        sKeyEmpresa = KEY_CARACTER & CStr(objFilialEmpresa.lCodEmpresa)
        sKeyFilial = sKeyEmpresa & SEPARADOR & CStr(objFilialEmpresa.iCodFilial)
        Call Filiais.Nodes.Add(sKeyEmpresa, tvwChild, sKeyFilial, objFilialEmpresa.sNomeFilial)
        
    Next
 
    'Expande os nós
    For Each objNode In Filiais.Nodes
        objNode.Expanded = True
    Next
    
    'Seleciona primeira Empresa
    Set Filiais.SelectedItem = Filiais.Nodes(1)
    
    'Para poder fazer transacao no bd da empresa
    lErro = Sistema_DefEmpresa(colEmpresa(1).sNome, colEmpresa(1).lCodigo, EMPRESA_TODA_NOME, EMPRESA_TODA)
    If lErro <> AD_BOOL_TRUE Then Error 25596
    
    'Carrega combos de Estados
    Set objTelaAux = Me
    lErro = gobjTabEnd.Inicializa(objTelaAux, TabEnd(0), TabEnd(1), TabEnd(2))
    If lErro <> SUCESSO Then Error 25929
    
    'Carrega combos de Estados
    lErro = Carrega_Estados()
    If lErro <> SUCESSO Then Error 25929
    
    lErro = Inicializa_Grid_AdmCad(objGridAdmCad)
    If lErro <> SUCESSO Then Error 25930
    
    lErro = Inicializa_Grid_Contrib(objGridContrib)
    If lErro <> SUCESSO Then Error 25930
    
    lErro = CF("Carrega_Combo", SignatQualif, "SpedQualificacoes", "Codigo", TIPO_LONG, "Descricao", TIPO_STR)
    If lErro <> SUCESSO Then Error 25930
    
    lErro = CF("Carrega_Combo", AdmCadCodigo, "SpedAdmCad", "Codigo", TIPO_STR, "Descricao", TIPO_STR)
    If lErro <> SUCESSO Then Error 25930
    
    NFeAmbiente.ListIndex = 1
    RPSAmbiente.ListIndex = 1
 
    lErro_Chama_Tela = SUCESSO
 
    Exit Sub

Erro_FilialEmpresa_Form_Load:
    
    lErro_Chama_Tela = Err
    
    Select Case Err
        
        Case 6189, 25523, 25596, 25929, 25930  'tratado na rotina chamada
        
        Case 25524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_EMPRESAS", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155883)
    
    End Select
    
    Exit Sub

End Sub

Public Property Get objTela() As Object
    Set objTela = m_objTela
End Property

Public Property Set objTela(ByVal vData As Object)
    Set m_objTela = vData
End Property

Public Property Get Controls() As Object
    Set Controls = objTela.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objTela.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objTela.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objTela.Enabled = New_Enabled
End Property

Public Sub Show()
    objTela.Show
    objTela.SetFocus
End Sub

Public Sub Unload(objMe As Object)
    RaiseEvent Unload
End Sub

Public Property Get Name() As String
    Name = objTela.Name
End Property

'################################################
'Inserido por Wagner 20/03/2006
Public Sub SimplesFederal_Click()

Dim lErro As Long

On Error GoTo Erro_SimplesFederal_Validate

    If SimplesFederal.Value = vbChecked Then
        
        SimplesFederalTeto.Enabled = True
        SimplesFederalAliq.Enabled = True
        'SuperSimples.Value = vbUnchecked
        PISNaoCumulativo.Value = vbUnchecked
    
    Else
    
        SimplesFederalTeto.Enabled = False
        SimplesFederalAliq.Enabled = False
    
        SimplesFederalTeto.Text = ""
        SimplesFederalAliq.Text = ""
    
    End If

    Exit Sub
    
Erro_SimplesFederal_Validate:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 141974)
    
    End Select
    
    Exit Sub

End Sub

Public Sub SimplesFederalAliq_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SimplesFederalAliq_Validate

    'Verifica se foi preenchido o percentual
    If Len(Trim(SimplesFederalAliq.Text)) <> 0 Then
    
        'Critica se é porcentagem
        lErro = Porcentagem_Critica(SimplesFederalAliq.Text)
        If lErro <> SUCESSO Then gError 141972
        
        'Formata
        SimplesFederalAliq.Text = Format(SimplesFederalAliq.Text, "Fixed")
    
    End If
    
    Exit Sub
    
Erro_SimplesFederalAliq_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 141972 'Tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 141973)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub SimplesFederalTeto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SimplesFederalTeto_Validate

    'Verifica se foi preenchido o percentual
    If Len(Trim(SimplesFederalTeto.Text)) <> 0 Then
    
        'Critica se é porcentagem
        lErro = Valor_NaoNegativo_Critica(SimplesFederalTeto.Text)
        If lErro <> SUCESSO Then gError 141975
        
        'Formata
        SimplesFederalTeto.Text = Format(SimplesFederalTeto.Text, "STANDARD")
    
    End If
    
    Exit Sub
    
Erro_SimplesFederalTeto_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 141975 'Tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 141976)
    
    End Select
    
    Exit Sub
    
End Sub
'################################################

Public Sub SuperSimples_Click()

Dim lErro As Long

On Error GoTo Erro_SuperSimples_Click

    If SuperSimples.Value = vbChecked Then
        
'        SimplesFederal.Value = vbUnchecked
        LucroPresumido.Value = vbUnchecked
        PISNaoCumulativo.Value = vbUnchecked
    
    End If
    

    Exit Sub
    
Erro_SuperSimples_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 141974)
    
    End Select
    
    Exit Sub

End Sub

Public Sub PISNaoCumulativo_Click()
    '
End Sub

Public Sub LucroPresumido_Click()
    PISNaoCumulativo.Value = vbUnchecked
End Sub

Public Sub ICMSPorEstimativa_Click()
    '
End Sub

Public Sub CNAE_GotFocus()
    Call MaskEdBox_TrataGotFocus(CNAE)
End Sub

Public Sub CNAE_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sDescricao As String

On Error GoTo Erro_CNAE_Validate

    If Len(Trim(CNAE.Text)) <> 0 Then
    
        lErro = CF("CNAE_Le", CNAE.Text, sDescricao)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 202236
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 202237
        
    End If
    
    DescCNAE.Caption = sDescricao
    
    Exit Sub
    
Erro_CNAE_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 202236  'tratado na rotina chamada
        
        Case 202237
            Call Rotina_Erro(vbOKOnly, "ERRO_CNAE_NAO_CADASTRADO", gErr, CNAE.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202238)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub GridAdmCad_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAdmCad, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAdmCad, iAlterado)
    End If

End Sub

Public Sub GridAdmCad_GotFocus()
    Call Grid_Recebe_Foco(objGridAdmCad)
End Sub

Public Sub GridAdmCad_EnterCell()
    Call Grid_Entrada_Celula(objGridAdmCad, iAlterado)
End Sub

Public Sub GridAdmCad_LeaveCell()
    If objGridAdmCad.iSaidaCelula = 1 Then Call Saida_Celula(objGridAdmCad)
End Sub

Public Sub GridAdmCad_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridAdmCad)
End Sub

Public Sub GridAdmCad_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAdmCad, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAdmCad, iAlterado)
    End If

End Sub

Public Sub GridAdmCad_LostFocus()
    Call Grid_Libera_Foco(objGridAdmCad)
End Sub

Public Sub GridAdmCad_RowColChange()
    Call Grid_RowColChange(objGridAdmCad)
End Sub

Public Sub GridAdmCad_Scroll()
    Call Grid_Scroll(objGridAdmCad)
End Sub

Public Sub AdmCadCodigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdmCad)
End Sub

Public Sub AdmCadCodigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdmCad)
End Sub

Public Sub AdmCadCodigo_LostFocus()
    Set objGridAdmCad.objControle = AdmCadCodigo
    Call Grid_Campo_Libera_Foco(objGridAdmCad)
End Sub

Public Sub AdmCadInscricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAdmCad)
End Sub

Public Sub AdmCadInscricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdmCad)
End Sub

Public Sub AdmCadInscricao_LostFocus()
    Set objGridAdmCad.objControle = AdmCadInscricao
    Call Grid_Campo_Libera_Foco(objGridAdmCad)
End Sub

Public Sub SignatCPF_GotFocus()
    Call MaskEdBox_TrataGotFocus(SignatCPF)
End Sub

Public Sub SignatCPF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SignatCPF_Validate

    If Len(Trim(SignatCPF.Text)) = 0 Then Exit Sub
    
    If Len(Trim(SignatCPF.Text)) <> STRING_CPF Then Error 25558
    
    lErro = Cpf_Critica(SignatCPF.Text)
    If lErro <> SUCESSO Then Error 25557
    
    Exit Sub
    
Erro_SignatCPF_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 25557  'tratado na rotina chamada
    
        Case 25558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CPF", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155871)
            
    End Select
    
    Exit Sub
    
End Sub

Private Function Inicializa_Grid_AdmCad(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid AdmCad

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_AdmCad

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Código")
    objGridInt.colColuna.Add ("Inscrição")
    
    'campos de edição do grid
    objGridInt.colCampo.Add (AdmCadCodigo.Name)
    objGridInt.colCampo.Add (AdmCadInscricao.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_AdmCadCodigo_Col = 1
    iGrid_AdmCadInscricao_Col = 2

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridAdmCad

    'Linhas do grid
    objGridInt.objGrid.Rows = 30 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 250

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_AdmCad = SUCESSO

    Exit Function

Erro_Inicializa_Grid_AdmCad:

    Inicializa_Grid_AdmCad = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190633)

    End Select

    Exit Function

End Function

Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridAdmCad.Name Then

            Select Case objGridInt.objGrid.Col
    
                Case iGrid_AdmCadCodigo_Col
                    Set objGridInt.objControle = AdmCadCodigo
    
                    'critica da coluna 1 vazia
                    lErro = Grid_Abandona_Celula(objGridInt)
                    If lErro Then Error 32100
    
                Case iGrid_AdmCadInscricao_Col
                    Set objGridInt.objControle = AdmCadInscricao
    
                    'critica da coluna 2 vazia
                    lErro = Grid_Abandona_Celula(objGridInt)
                    If lErro Then Error 32101
    
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridContrib.Name Then

             Select Case objGridInt.objGrid.Col
    
                Case iGrid_ContribUF_Col
                    Set objGridInt.objControle = ContribUF
    
                    lErro = Grid_Abandona_Celula(objGridInt)
                    If lErro Then Error 32100
    
                Case iGrid_ContribIE_Col
    
                    lErro = Saida_Celula_IE(objGridInt)
                    If lErro <> SUCESSO Then Error 32100
    
            End Select
        
        End If
        
        If Len(Trim(objGridInt.objControle.Text)) > 0 Then
            If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then Error 32102

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 32100, 32101, 32102
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161729)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Contrib(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Contrib

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Contrib

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("UF")
    objGridInt.colColuna.Add ("Inscrição Estadual")
    
    'campos de edição do grid
    objGridInt.colCampo.Add (ContribUF.Name)
    objGridInt.colCampo.Add (ContribIE.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_ContribUF_Col = 1
    iGrid_ContribIE_Col = 2

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridContrib

    'Linhas do grid
    objGridInt.objGrid.Rows = 30 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 3
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 250
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Contrib = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Contrib:

    Inicializa_Grid_Contrib = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190633)

    End Select

    Exit Function

End Function


Public Sub GridContrib_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridContrib, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContrib, iAlterado)
    End If

End Sub

Public Sub GridContrib_GotFocus()
    Call Grid_Recebe_Foco(objGridContrib)
End Sub

Public Sub GridContrib_EnterCell()
    Call Grid_Entrada_Celula(objGridContrib, iAlterado)
End Sub

Public Sub GridContrib_LeaveCell()
    If objGridContrib.iSaidaCelula = 1 Then Call Saida_Celula(objGridContrib)
End Sub

Public Sub GridContrib_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridContrib)
End Sub

Public Sub GridContrib_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridContrib, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridContrib, iAlterado)
    End If

End Sub

Public Sub GridContrib_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridContrib)
End Sub

Public Sub GridContrib_RowColChange()
    Call Grid_RowColChange(objGridContrib)
End Sub

Public Sub GridContrib_Scroll()
    Call Grid_Scroll(objGridContrib)
End Sub

Public Sub ContribIE_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContrib)
End Sub

Public Sub ContribIE_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContrib)
End Sub

Public Sub ContribIE_Validate(Cancel As Boolean)
    Set objGridContrib.objControle = ContribIE
    Call Grid_Campo_Libera_Foco(objGridContrib)
End Sub

Public Sub ContribUF_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridContrib)
End Sub

Public Sub ContribUF_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridContrib)
End Sub

Public Sub ContribUF_Validate(Cancel As Boolean)
    Set objGridContrib.objControle = ContribUF
    Call Grid_Campo_Libera_Foco(objGridContrib)
End Sub

Public Sub CNPJCTB_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CNPJCTB)
End Sub

Public Sub CNPJCTB_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CNPJCTB_Validate

    If Len(Trim(CNPJCTB.Text)) = 0 Then Exit Sub
    
    If Len(Trim(CNPJCTB.Text)) <> STRING_CGC Then gError 25559
    
    lErro = Cgc_Critica(CNPJCTB.Text)
    If lErro <> SUCESSO Then gError 25556
    
    Exit Sub
    
Erro_CNPJCTB_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 25556  'tratado na rotina chamada
        
        Case 25559
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155866)
            
    End Select
    
    Exit Sub
End Sub

Private Function Saida_Celula_IE(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_IE

    Set objGridInt.objControle = ContribIE
    
    If Len(Trim(ContribIE.Text)) > 0 Then
    
        lErro = CF("InscricaoEstadual_Critica", ContribIE.Text, GridContrib.TextMatrix(GridContrib.Row, iGrid_ContribUF_Col))
        If lErro <> SUCESSO Then gError 190688
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190689

    Saida_Celula_IE = SUCESSO

    Exit Function

Erro_Saida_Celula_IE:

    Saida_Celula_IE = gErr

    Select Case gErr

        Case 190688, 190689
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190690)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
              
    Select Case objControl.Name
    
        Case ContribUF.Name
            If Len(Trim(GridContrib.TextMatrix(iLinha, iGrid_ContribUF_Col))) <> 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
        Case ContribIE.Name
            If Len(Trim(GridContrib.TextMatrix(iLinha, iGrid_ContribUF_Col))) <> 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190665)

    End Select

    Exit Sub

End Sub

Public Sub ContribUF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ContribUF_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ContribIE_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Function Carrega_Estados() As Long
'Carrega combos de Estados

Dim lErro As Long
Dim iIndice As Integer
Dim vCodigo As Variant
Dim colCodigo As New Collection

On Error GoTo Erro_Carrega_Estados

    'Lê cada codigo da tabela Estados
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then Error 25593
    
    For Each vCodigo In colCodigo
        ContribUF.AddItem vCodigo
    Next

    Carrega_Estados = SUCESSO

Exit Function

Erro_Carrega_Estados:
    
    Carrega_Estados = Err
    
    Select Case Err
        
        Case 25593 'tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155884)
     
     End Select
     
     Exit Function

End Function

Public Sub RegimeEspecialTrib_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub RegimeTrib_Click()
    iAlterado = REGISTRO_ALTERADO
    
    SuperSimples.Value = vbUnchecked
    LucroPresumido.Value = vbUnchecked
    PISNaoCumulativo.Value = vbUnchecked
    
    Select Case Codigo_Extrai(RegimeTrib.Text)
        Case 1
            SuperSimples.Value = vbChecked
        Case 2
            LucroPresumido.Value = vbChecked
        Case 3
            PISNaoCumulativo.Value = vbChecked
        Case 4
        
    End Select
    '--1-Simples
    '--2-Lucro Presumido
    '--3-Lucro Real - Não Cumulativo
    '--4-Lucro Real - Cumulativo
End Sub

Public Sub idNFCECSC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub idNFCECSC_Validate(Cancel As Boolean)
'
End Sub

Public Sub NFCECSC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NFCECSC_Validate(Cancel As Boolean)
'
End Sub

Public Sub indSincPadrao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub indSincPadrao_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub


