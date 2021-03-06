VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTesteVBNET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

#Const VERIFICA_ADLOCK = 1

'###################################
'Inserido por Wagner 20/10/2005
Private Const LIMITA_DATA_USO = 0 '0- N�o Limita, 1- Limita Data Emiss�o NF, 2 - Limita Qtds NF e 3 - Limita Data e Qtds NFs
Private Const DATA_INICIAL_REFERENCIA = "2010-07-20"
Private Const MESES_DE_USO = 12 'Qtds de meses a mais que se pode usar
Private Const MEDIA_NFS_MES = 0 '0-Obt�m a m�dia dinamicamente, Outro Valor - For�a a M�dia a ser o valor informado
'###################################

Sub IniciarSGE(ByVal sTexto As String)

Dim lSistema As Long, lErro As Long, sCodTela As String, sProjeto As String, sClasse As String
Dim objUsuarioEmpresa As New ClassUsuarioEmpresa
Dim bSplashFormLoaded As Boolean
Dim objFilialEmpresa As New ClassFilialEmpresa
Dim sNomeArqParam As String
Dim objObject As Object
Dim bDemo As Boolean
Dim Y As New ClassConstCust

On Error GoTo Erro_Main

    MsgBox (sTexto)
    
    bDemo = False
    
    'Call Inicializa_Tamanhos_String
    
    gsNomePrinc = "SGEPrinc"
    
    frmSplashSGEPrinc.Show
    DoEvents
    
    bSplashFormLoaded = True

    App.HelpFile = App.Path & "\sgeprinc.hlp"
    
    'para permitir acessar o dicionario de dados
    lSistema = Sistema_Abrir()
    If lSistema = 0 Then gError 41615
    
    lErro = Single_Logon
    If lErro = 187464 Then gError 41615

    If bDemo Then
    
        objUsuarioEmpresa.sNomeEmpresa = "Corporator"
        objUsuarioEmpresa.sNomeFilial = "Matriz"
        objUsuarioEmpresa.sSenha = "usuario1"
        objUsuarioEmpresa.lCodEmpresa = 1
        objUsuarioEmpresa.iCodFilial = 1
        objUsuarioEmpresa.sCodUsuario = "usuario1"
        objUsuarioEmpresa.iTelaOK = True

    Else
    
        'carrega a tela p/identificacao do usuario e selecao da Empresa e filial
        Load UsuarioEmpresa
    
        lErro = UsuarioEmpresa.Trata_Parametros(objUsuarioEmpresa)
        If lErro <> SUCESSO Then gError 41616
    
        UsuarioEmpresa.Show vbModal
    
        DoEvents
    
        If objUsuarioEmpresa.iTelaOK = False Then gError 41617
    
    End If
    
    Set gobjCheckboxChecked = LoadPicture("checkboxchecked.bmp")
    Set gobjCheckboxUnchecked = LoadPicture("checkboxunchecked.bmp")
    Set gobjCheckboxGrayed = LoadPicture("checkboxgrayed.bmp")
    Set gobjOptionButtonChecked = LoadPicture("optionbuttonchecked.bmp")
    Set gobjOptionButtonUnChecked = LoadPicture("optionbuttonunchecked.bmp")
    Set gobjButton = LoadPicture("botao.bmp")

    'faz login utilizando o codigo do usuario e a senha
''''    lErro = Sistema_Login(objUsuarioEmpresa.sCodUsuario, objUsuarioEmpresa.sSenha)
    lErro = Usuario_Executa_Login(objUsuarioEmpresa.sCodUsuario, objUsuarioEmpresa.sSenha)
    If lErro <> SUCESSO Then gError 41618

    Call Y.Inicializa_Tamanhos_String
        
    objFilialEmpresa.iCodFilial = objUsuarioEmpresa.iCodFilial
    objFilialEmpresa.sNomeFilial = objUsuarioEmpresa.sNomeFilial
    objFilialEmpresa.lCodEmpresa = objUsuarioEmpresa.lCodEmpresa
    objFilialEmpresa.sNomeEmpresa = objUsuarioEmpresa.sNomeEmpresa
    
    'Configura Empresa e Filial inclusive conex�o
    lErro = Empresa_Filial_Configura(objFilialEmpresa)
    If lErro <> SUCESSO Then gError 25875
    
    DoEvents
       
    'apenas para agilizar cargas futuras de telas
    Call Tela_ObterFuncao(sCodTela, sProjeto, sClasse)
    
    Set gcolModulo = New AdmColModulo
    
    'Carrega em gcolModulo os m�dulos indicando ativadade p/ FilialEmpresa
    lErro = CF("Modulos_Le_Empresa_Filial", objUsuarioEmpresa.lCodEmpresa, objUsuarioEmpresa.iCodFilial, gcolModulo)
    If lErro <> SUCESSO Then gError 44984
        
    lErro = CF("Verifica_Configuracoes", LIMITA_DATA_USO, DATA_INICIAL_REFERENCIA, MESES_DE_USO, MEDIA_NFS_MES)
    If lErro <> SUCESSO Then gError 140519
    
    lErro = CF("Atualiza_Versao")
    If lErro <> SUCESSO Then gError 180581
        
#If VERIFICA_ADLOCK = 1 Then
    
    lErro = CF("Valida_Controle")
    If lErro <> SUCESSO Then gError 141618

#End If

    lErro = CF("PV_Exclui_Reservas")
    If lErro <> SUCESSO Then gError 180581
    
'    lErro = CF("Tributacao_Atualiza_Versao")
'    If lErro <> SUCESSO Then gError 180581

    'carrega a tela principal
    lErro_Chama_Tela = SUCESSO
    PrincipalNovo.Show
    If lErro_Chama_Tela <> SUCESSO Then Unload PrincipalNovo
    
    If gobjLoja.lIntervaloTrans > 0 Then

        'Prepara para chamar rotina batch
        lErro = Sistema_Preparar_Batch(sNomeArqParam)
        If lErro <> SUCESSO Then gError 133520

        gobjLoja.sNomeArqParam = sNomeArqParam

        Set gobjLoja.colModulo = gcolModulo

        Set objObject = gobjLoja

        lErro = CF("Rotina_FTP_Recepcao_CC", objObject)
        If lErro <> SUCESSO And lErro <> 133628 Then gError 133521

        If lErro <> SUCESSO Then Call Rotina_Aviso(vbOKOnly, "AVISO_NAO_CARREGOU_ROTINA_RECEPCAO")

    End If
    
    Unload frmSplashSGEPrinc
    
'' 'codigo comentado pertence a GSilva
''    lErro = Pede_CotacaoMoeda_Dia()
''    If lErro <> SUCESSO Then gError 84726
    
    Exit Sub
    
Erro_Main:

    If bSplashFormLoaded Then Unload frmSplashSGEPrinc
    
    Select Case gErr
    
        Case 25875, 41615 To 41618, 44984, 84726, 133520, 133521, 140519, 141618, 180581 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165215)

    End Select
    
    If lSistema <> 0 Then Call Sistema_Fechar
    
    Exit Sub

End Sub

Private Function Rotina_Configuracao_Empresa(bConfigurouEmpresa As Boolean) As Long
'faz a configuracao a nivel de empresa

Dim colModuloFilEmp As New Collection
Dim objModuloFilEmp As ClassModuloFilEmp
Dim iConfigurarSGE As Integer
Dim objConfiguraADM As New ClassConfiguraADM
Dim colModuloFilial As New Collection
Dim lErro As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Rotina_Configuracao_Empresa

    'le todos os objetos ModuloFilEmp para a empresa em quest�o e coloca-os em colModuloFilEmp
    lErro = ModuloFilEmp_Le_EmpresaFilial(glEmpresa, EMPRESA_TODA, colModuloFilEmp)
    If lErro <> SUCESSO Then Error 44858
    
    iConfigurarSGE = True
    
    'pesquisa se h� algum m�dulo a configurar que necessita passar pela tela de configura��o
    For Each objModuloFilEmp In colModuloFilEmp
        If objModuloFilEmp.iConfigurado = NAO_CONFIGURADO Then
                objConfiguraADM.colModulosConfigurar.Add objModuloFilEmp.sSiglaModulo
        End If
        'pesquisa se h� algum modulo da empresa j� configurado ==> significa que a configura��o geral da empresa j� foi feita
        If objModuloFilEmp.iConfigurado = CONFIGURADO Then
            iConfigurarSGE = False
        End If
    Next
    
    If iConfigurarSGE = True Then objConfiguraADM.colModulosConfigurar.Add SISTEMA_SGE
    
    If objConfiguraADM.colModulosConfigurar.Count > 0 Then
    
        Call Carrega_ColFiliais_EmpresaToda
        
        'carrega o wizard de configura��o da empresa
        objConfiguraADM.iConfiguracaoOK = False
    
        Call Chama_Tela("frmWizardEmpresa", objConfiguraADM)
    
        If objConfiguraADM.iConfiguracaoOK = False Then Error 44859
    
        lErro = CF("Retorna_ColFiliais")
        If lErro <> SUCESSO Then Error 44944
    
        bConfigurouEmpresa = True
    
    End If
    
    Rotina_Configuracao_Empresa = SUCESSO
    
    Exit Function
    
Erro_Rotina_Configuracao_Empresa:

    Rotina_Configuracao_Empresa = Err
    
    Select Case Err
    
        Case 44858, 44859, 44944
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165216)

    End Select
    
    Exit Function
    
End Function

Private Function Rotina_Configuracao_Filial(bConfigurouFilial As Boolean) As Long
'faz a configuracao a nivel de filial

Dim colModuloFilEmp As New Collection
Dim objModuloFilEmp As ClassModuloFilEmp
Dim objConfiguraADM As New ClassConfiguraADM
Dim colModuloFilial As New Collection
Dim lErro As Long

On Error GoTo Erro_Rotina_Configuracao_Filial
    
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        'le todos os objetos ModuloFilEmp para a filial em quest�o e coloca-os em colModuloFilEmp
        lErro = ModuloFilEmp_Le_EmpresaFilial(glEmpresa, giFilialEmpresa, colModuloFilEmp)
        If lErro <> SUCESSO Then gError 44860
    
        'pesquisa se h� algum m�dulo a configurar que necessita passar pela tela de configura��o
        For Each objModuloFilEmp In colModuloFilEmp
            If objModuloFilEmp.iConfigurado = NAO_CONFIGURADO Then
                objConfiguraADM.colModulosConfigurar.Add objModuloFilEmp.sSiglaModulo
                'seleciona os m�dulos que necessitam passar por tela de configura��o.
                If objModuloFilEmp.sSiglaModulo = MODULO_ESTOQUE Then
                    colModuloFilial.Add objModuloFilEmp.sSiglaModulo
                End If
                    
            End If
        Next
    
        If colModuloFilial.Count > 0 Then
            
            Call Carrega_ColFiliais_Filial(objConfiguraADM)
            
            objConfiguraADM.iConfiguracaoOK = False
        
            'carrega o wizard de configura��o da filial
            Call Chama_Tela("frmWizardFilial", objConfiguraADM)
        
            If objConfiguraADM.iConfiguracaoOK = False Then gError 44861
            
            lErro = CF("Retorna_ColFiliais")
            If lErro <> SUCESSO Then gError 44946
        
            bConfigurouFilial = True
        
        ElseIf objConfiguraADM.colModulosConfigurar.Count > 0 Then
        
            Call Carrega_ColFiliais_Filial(objConfiguraADM)
        
            lErro = Gravar_Registro(objConfiguraADM.colModulosConfigurar)
            If lErro <> SUCESSO Then gError 44875
        
            lErro = CF("Retorna_ColFiliais")
            If lErro <> SUCESSO Then gError 44947
        
            bConfigurouFilial = True
        
        End If
    
    End If

    'se configurou a filial
    If bConfigurouFilial = True Then
        'cria os registros nos arquivos config (ESTConfig, FATCOnfig, etc) que dependem de filial. Aqueles que tem o campo PorFilial = POR_FILIAL
        lErro = CF("Config_Instalacao_Filial", giFilialEmpresa, objConfiguraADM.colModulosConfigurar)
        If lErro <> SUCESSO Then gError 110029
    End If

    Rotina_Configuracao_Filial = SUCESSO
    
    Exit Function
    
Erro_Rotina_Configuracao_Filial:

    Rotina_Configuracao_Filial = gErr
    
    Select Case gErr
    
        Case 44860, 44861, 44875, 44946, 44947, 110029
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165217)

    End Select
    
    Exit Function
    
End Function

Private Sub Carrega_ColFiliais_EmpresaToda()

Dim objFiliais As AdmFiliais

    'coloca gcolFiliais como uma cole��o de filiais composta somente pela empresa toda
    Set gcolFiliais = New Collection
    
    Set objFiliais = New AdmFiliais
    
    objFiliais.sNome = gsNomeEmpresa
    objFiliais.iCodFilial = EMPRESA_TODA

    'coloca a filial lida na cole��o
    gcolFiliais.Add objFiliais

End Sub


Private Sub Carrega_ColFiliais_Filial(objConfiguraADM As ClassConfiguraADM)

Dim objFiliais As AdmFiliais

    'coloca gcolFiliais como uma cole��o de filiais composta somente pela empresa toda
    Set gcolFiliais = New Collection
    
    Set objFiliais = New AdmFiliais
    
    objFiliais.sNome = gsNomeFilialEmpresa
    objFiliais.iCodFilial = giFilialEmpresa
    Set objFiliais.colModulos = objConfiguraADM.colModulosConfigurar

    'coloca a filial lida na cole��o
    gcolFiliais.Add objFiliais

End Sub

Private Function Valida_Step(sModulo As String, colModulosConfigurar As Collection) As Long

Dim vModulo As Variant

    For Each vModulo In colModulosConfigurar

        If sModulo = vModulo Then
            Valida_Step = SUCESSO
            Exit Function
        End If
        
    Next
    
    Valida_Step = 44870

End Function

Private Function Gravar_Registro(colModulosConfigurar As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim lTransacaoDic As Long
Dim lConexao As Long

On Error GoTo Erro_Gravar_Registro
    
    lConexao = GL_lConexaoDic
    
    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirDic
    If lTransacaoDic = 0 Then Error 44963
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 44871
    
    lErro = CTB_Exercicio_Gravar_Registro(colModulosConfigurar)
    If lErro <> SUCESSO Then Error 44872
    
    lErro = CR_Filial_Gravar_Registro(colModulosConfigurar)
    If lErro <> SUCESSO Then Error 41927
    
    lErro = EST_Filial_Gravar_Registro(colModulosConfigurar)
    If lErro <> SUCESSO Then Error 41928
    
    lErro = FAT_Filial_Gravar_Registro(colModulosConfigurar)
    If lErro <> SUCESSO Then Error 41929
    
    lErro = CF("ModuloFilEmp_Atualiza_Configurado", glEmpresa, giFilialEmpresa, colModulosConfigurar)
    If lErro <> SUCESSO Then Error 44956
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 44873
    
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 44964
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
    
    Gravar_Registro = Err
    
    Select Case Err

        Case 44871
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 44872, 44956, 44963, 44964, 41927, 41928, 41929

        Case 44873
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165218)

    End Select

    If Err <> 44964 Then Call Transacao_Rollback
    Call Transacao_RollbackDic

    Exit Function
    
End Function

Private Function CTB_Exercicio_Gravar_Registro(colModulosConfigurar As Collection) As Long

Dim lErro As Long

On Error GoTo Erro_CTB_Exercicio_Gravar_Registro

    lErro = Valida_Step(MODULO_CONTABILIDADE, colModulosConfigurar)

    If lErro = SUCESSO Then
        
        lErro = CF("Exercicio_Instalacao_Filial", giFilialEmpresa)
        If lErro <> SUCESSO Then Error 44874
        
    End If
    
    CTB_Exercicio_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_CTB_Exercicio_Gravar_Registro:
    
    CTB_Exercicio_Gravar_Registro = Err
    
    Select Case Err

        Case 44874

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165219)

    End Select

    Exit Function
    
End Function

Private Function CR_Filial_Gravar_Registro(colModulosConfigurar As Collection) As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_CR_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_CONTASARECEBER, colModulosConfigurar)

    If lErro = SUCESSO Then
        
        lErro = CF("CR_Instalacao_Filial", giFilialEmpresa)
        If lErro <> SUCESSO Then Error 41913
        
    End If
    
    CR_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_CR_Filial_Gravar_Registro:
    
    CR_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41913

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165220)

    End Select

    Exit Function
    
End Function

Private Function EST_Filial_Gravar_Registro(colModulosConfigurar As Collection) As Long

Dim lErro As Long
Dim colSegmentos As Collection
Dim sIntervaloProducao As String

On Error GoTo Erro_EST_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_ESTOQUE, colModulosConfigurar)

    If lErro = SUCESSO Then
        
        sIntervaloProducao = "0"
        lErro = CF("EST_Instalacao_Filial", giFilialEmpresa, sIntervaloProducao)
        If lErro <> SUCESSO Then Error 41914
        
    End If
    
    EST_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_EST_Filial_Gravar_Registro:
    
    EST_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41914

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165221)

    End Select

    Exit Function
    
End Function

Private Function FAT_Filial_Gravar_Registro(colModulosConfigurar As Collection) As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_FAT_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_FATURAMENTO, colModulosConfigurar)

    If lErro = SUCESSO Then
        
        lErro = CF("FAT_Instalacao_Filial", giFilialEmpresa)
        If lErro <> SUCESSO Then Error 41915
        
    End If
    
    FAT_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_FAT_Filial_Gravar_Registro:
    
    FAT_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41915

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165222)

    End Select

    Exit Function
    
End Function

Public Function Empresa_Filial_Configura(objFilialEmpresa As ClassFilialEmpresa) As Long

Dim bConfigurouEmpresa As Boolean
Dim bConfigurouFilial As Boolean
Dim lErro As Long

On Error GoTo Erro_Empresa_Filial_Configura

    'seleciona a Empresa e filial
    lErro = Sistema_DefEmpresa(objFilialEmpresa.sNomeEmpresa, objFilialEmpresa.lCodEmpresa, objFilialEmpresa.sNomeFilial, objFilialEmpresa.iCodFilial)
    If lErro <> AD_BOOL_TRUE Then Error 41619
    
    glEmpresa = objFilialEmpresa.lCodEmpresa
    
    bConfigurouEmpresa = False
    
    lErro = Rotina_Configuracao_Empresa(bConfigurouEmpresa)
    If lErro <> SUCESSO Then Error 44876
    
    bConfigurouFilial = False
    
    lErro = Rotina_Configuracao_Filial(bConfigurouFilial)
    If lErro <> SUCESSO Then Error 44877
    
    'se houve configuracao de modulo
    If bConfigurouEmpresa = True Or bConfigurouFilial = True Then
    
        'for�a a reinicializacao dos modulos, por exemplo para pegar a nova mascara de conta contabil
        If Sistema_Inicializa_Modulos <> SUCESSO Then Error 56601

    End If
    
    Empresa_Filial_Configura = SUCESSO
    
    Exit Function
    
Erro_Empresa_Filial_Configura:

    Empresa_Filial_Configura = Err
    
    Select Case Err
    
        Case 41619, 44876, 44877, 56601  'tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165223)

    End Select
    
    Exit Function
 
End Function

'??? transferir p/rotinasadm
Private Function Pede_CotacaoMoeda_Dia() As Long

'''' rotinas comentadas pertencem a GSilva

'Dim lErro As Long
'Dim lComando As Long
'Dim dValor As Double
'Dim iTipoMoeda As Integer
'
'On Error GoTo Erro_Pede_CotacaoMoeda_Dia
'
'    'Abre Comandos
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 84722
'
'    'T�tulo da moeda Estrangeira
'    iTipoMoeda = 1 'DOLAR
'
'    'Faz sele��o do campo valor passando data e moeda como param�tros
'    lErro = Comando_Executar(lComando, "SELECT Valor FROM CotacoesMoeda WHERE Data = ? AND Moeda = ?", dValor, gdtDataHoje, iTipoMoeda)
'    If lErro <> AD_SQL_SUCESSO Then gError 84723
'
'    'Tenta encontrar o registro
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84724
'
'    If lErro = AD_SQL_SUCESSO Then
'
'        Pede_CotacaoMoeda_Dia = SUCESSO
'
'        Exit Function
'
'    End If
'
'    lErro = Chama_Tela("CotacaoMoeda")
'    If lErro <> SUCESSO Then gError 84725
'
'    Call Comando_Fechar(lComando)
'
'    Pede_CotacaoMoeda_Dia = SUCESSO
'
'    Exit Function
'
'Erro_Pede_CotacaoMoeda_Dia:
'
'    Pede_CotacaoMoeda_Dia = gErr
'
'    Select Case gErr
'
'        Case 84722
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 84725
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CHAMA_TELA_COTACAOMOEDA", gErr)
'
'        Case 84723,84724
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACOESMOEDA2", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165224)
'
'    End Select
'
'    'Fecha Comandos
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function

End Function


Private Sub Inicializa_Tamanhos_String()

    STRING_ENDERECO = GetPrivateProfileInt("Tamanhos", "STRING_ENDERECO", STRING_ENDERECO, NOME_ARQUIVO_ADM)
    STRING_BAIRRO = GetPrivateProfileInt("Tamanhos", "STRING_BAIRRO", STRING_BAIRRO, NOME_ARQUIVO_ADM)
    STRING_CIDADE = GetPrivateProfileInt("Tamanhos", "STRING_CIDADE", STRING_CIDADE, NOME_ARQUIVO_ADM)
    STRING_CLIENTE_RAZAO_SOCIAL = GetPrivateProfileInt("Tamanhos", "STRING_CLIENTE_RAZAO_SOCIAL", STRING_CLIENTE_RAZAO_SOCIAL, NOME_ARQUIVO_ADM)
    STRING_CLIENTE_NOME_REDUZIDO = GetPrivateProfileInt("Tamanhos", "STRING_CLIENTE_NOME_REDUZIDO", STRING_CLIENTE_NOME_REDUZIDO, NOME_ARQUIVO_ADM)
    STRING_TRANSPORTADORA_NOME = GetPrivateProfileInt("Tamanhos", "STRING_TRANSPORTADORA_NOME", STRING_TRANSPORTADORA_NOME, NOME_ARQUIVO_ADM)
    STRING_TRANSPORTADORA_NOME_REDUZIDO = GetPrivateProfileInt("Tamanhos", "STRING_TRANSPORTADORA_NOME_REDUZIDO", STRING_TRANSPORTADORA_NOME_REDUZIDO, NOME_ARQUIVO_ADM)
    
End Sub



