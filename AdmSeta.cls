VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmSeta"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum enumSeta
    BOTAO_PROXIMO
    BOTAO_ANTERIOR
    BOTAO_PRIMEIRO
    BOTAO_ULTIMO
    BOTAO_CONSULTA
End Enum

Public Enum enumSetaComparacao
    SETA_COMP_IGUAL
    SETA_COMP_DIFERENTE
End Enum

'local variable(s) to hold property value(s)
Private mvargs_ST_TelaIndice As String 'local copy
'local variable(s) to hold property value(s)
Private mvargs_ST_TelaTabela As String 'local copy
'local variable(s) to hold property value(s)
Private mvargcol_ST_CampoValor As New AdmColCampoValor 'local copy
Public Property Let gi_ST_Ordem_StrCmp(ByVal vData As Integer)
Attribute gi_ST_Ordem_StrCmp.VB_Description = "Ordem lexicografica para comparacao de strings no BD. 0 -> ASCII, case sensitive, 1 -> case insensitive, 2 -> comparacao windows case sensitive, 3 -> comparacao windows case insensitive"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gi_ST_Ordem_StrCmp = 5
    i_ST_Ordem_StrCmp = vData
End Property

Public Property Get gi_ST_Ordem_StrCmp() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gi_ST_Ordem_StrCmp
    gi_ST_Ordem_StrCmp = i_ST_Ordem_StrCmp
End Property

Public Property Set gcol_ST_CampoValor(ByVal vData As Object)
Attribute gcol_ST_CampoValor.VB_Description = "Coleção campo-valor dos campos da tela."
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.gcol_ST_CampoValor = Form1
    Set mvargcol_ST_CampoValor = vData
End Property
Public Property Get gcol_ST_CampoValor() As AdmColCampoValor
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gcol_ST_CampoValor
    Set gcol_ST_CampoValor = mvargcol_ST_CampoValor
End Property
Public Property Set gobj_ST_TelaAtiva(ByVal vData As Object)
Attribute gobj_ST_TelaAtiva.VB_Description = "Tela ativa (última form sujeita a setas a receber foco)"
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.gobj_ST_TelaAtiva = Form1
    Set obj_ST_TelaAtiva = vData
End Property
Public Property Get gobj_ST_TelaAtiva() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gobj_ST_TelaAtiva
    Set gobj_ST_TelaAtiva = obj_ST_TelaAtiva
End Property
Public Function ComandoSeta_Fechar(ByVal sTela As String) As Long
Attribute ComandoSeta_Fechar.VB_Description = "Fecha comando das setas se sTela for a tela que ativou o comando."
'Fecha o comando de setas se estiver aberto e for essa tela que o tiver ativado

'Dim lErro As Long
'
'    If gs_ST_TelaSetaClick = sTela And gl_ST_ComandoSeta <> 0 Then
'
'        lErro = Comando_Fechar(gl_ST_ComandoSeta)
'        gl_ST_ComandoSeta = 0
'
'    End If
'
'    ComandoSeta_Fechar = SUCESSO

End Function
Public Function ComandoSeta_Fechar2()
'Fecha comando de setas se estiver aberto qualquer que seja a tela que o originou

Dim lErro As Long

    If gl_ST_ComandoSeta <> 0 Then
    
        lErro = Comando_Fechar(gl_ST_ComandoSeta)
        gl_ST_ComandoSeta = 0
    
    End If

    ComandoSeta_Fechar2 = SUCESSO

End Function
Public Function ComandoSeta_Liberar(ByVal sTela As String) As Long
'Libera a referencia a essa tela e fecha o comando de setas se for essa tela que o tiver disparado

Dim lErro As Long

    'Fecha comando setas se essa tela o originou
    lErro = ComandoSeta_Fechar(sTela)
    
    'Libera referencia a sTela
    Set gobj_ST_TelaAtiva = Nothing
    
    'Passa a ignorar click
    gi_ST_SetaIgnoraClick = 1
    
    ComandoSeta_Liberar = SUCESSO

End Function

Public Property Let gs_ST_TelaTabela(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gs_ST_TelaTabela = 5
    mvargs_ST_TelaTabela = vData
End Property


Public Property Get gs_ST_TelaTabela() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gs_ST_TelaTabela
    gs_ST_TelaTabela = mvargs_ST_TelaTabela
End Property

Public Property Let gs_ST_TelaIndice(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gs_ST_TelaIndice = 5
    mvargs_ST_TelaIndice = vData
End Property
Public Property Get gs_ST_TelaIndice() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gs_ST_TelaIndice
    gs_ST_TelaIndice = mvargs_ST_TelaIndice
End Property
Public Property Let gs_ST_TelaSetaClick(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTela = 5
    s_ST_TelaSetaClick = vData
End Property
Public Property Get gs_ST_TelaSetaClick() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTela
    gs_ST_TelaSetaClick = s_ST_TelaSetaClick
End Property
Public Property Let gl_ST_ComandoSeta(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lComandoAnterior = 5
    l_ST_ComandoSeta = vData
End Property
Public Property Get gl_ST_ComandoSeta() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lComandoAnterior
    gl_ST_ComandoSeta = l_ST_ComandoSeta
End Property
Public Property Let gi_ST_SetaIgnoraClick(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.gi_ST_SetaIgnoraClick = 5
    
    i_ST_SetaIgnoraClick = vData
    
    If Not (gobj_ST_TelaPrincipal Is Nothing) Then
        If vData = 1 Then
            'Limpa Combo de Indices
            gobj_ST_TelaPrincipal.Indice.Clear
        End If
    End If
    
End Property
Public Property Get gi_ST_SetaIgnoraClick() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gi_ST_SetaIgnoraClick
    gi_ST_SetaIgnoraClick = i_ST_SetaIgnoraClick
End Property
Public Property Set gobj_ST_TelaPrincipal(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.gobj_ST_TelaPrincipal = Form1
    Set obj_ST_TelaPrincipal = vData
End Property
Public Property Get gobj_ST_TelaPrincipal() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.gobj_ST_TelaPrincipal
    Set gobj_ST_TelaPrincipal = obj_ST_TelaPrincipal
End Property
'Seleciona os nomes de usuario dos indices da tela e preenche Combo de índices na tela principal
Public Sub TelaIndice_Preenche(objFormAtiva As Object)

Dim lErro As Long
Dim colTelaIndice As New AdmColCodigoNome
Dim objTelaIndice As AdmCodigoNome

On Error GoTo Erro_TelaIndice_Preenche

    'Armazena referencia global da tela ativa
    Set gobj_ST_TelaAtiva = objFormAtiva
    
    'Setas não serão ignoradas
    gi_ST_SetaIgnoraClick = 0
   
    'Limpa Combo de Indices
    If Not (gobj_ST_TelaPrincipal Is Nothing) Then
        gobj_ST_TelaPrincipal.Indice.Clear
    End If

    'Le os indices da tela
    lErro = TelaIndices_Le(objFormAtiva.Name, colTelaIndice)
    If lErro Then Error 6593

    If Not (gobj_ST_TelaPrincipal Is Nothing) Then
    
        'Preenche combo de índices
        For Each objTelaIndice In colTelaIndice
    
            gobj_ST_TelaPrincipal.Indice.AddItem objTelaIndice.sNome
            gobj_ST_TelaPrincipal.Indice.ItemData(gobj_ST_TelaPrincipal.Indice.NewIndex) = objTelaIndice.iCodigo
    
        Next
    
        gobj_ST_TelaPrincipal.Indice.ListIndex = 0
    
    End If

    Exit Sub

Erro_TelaIndice_Preenche:

    Select Case Err

        Case 6593  'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142619)

    End Select

    Exit Sub

End Sub

Public Function TelaIndices_Le(ByVal sNomeTela As String, colCodigoNome As AdmColCodigoNome) As Long
'Lê nº seq e nomes externos de índices associados a sNomeTela. Retorna em colCodigoNome

Dim lConexao As Long
Dim lComando As Long
Dim lErro As Long
Dim iIndice As Integer 'Numero sequencial do índice
Dim sIndice As String  'Nome Externo do Índice

On Error GoTo Erro_TelaIndices_Le

    sIndice = Space(STRING_TELAINDICE_NOME_EXTERNO)

    lConexao = GL_lConexaoDic
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 6589

    'seleciona no BD todos os nomes externos de Indices associados a sNomeTela
    lErro = Comando_Executar(lComando, "SELECT Indice, NomeExterno FROM TelaIndice WHERE NomeTela=? ORDER BY Indice", iIndice, sIndice, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then Error 6590

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6591

    If lErro = AD_SQL_SEM_DADOS Then Error 60777

    'Preenche a coleção
    Do While lErro = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colCodigoNome.Add iIndice, sIndice
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6592

    Loop

    lErro = Comando_Fechar(lComando)

    TelaIndices_Le = SUCESSO

    Exit Function

Erro_TelaIndices_Le:

    TelaIndices_Le = Err

    Select Case Err

        Case 6589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6590, 6591, 6592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELAINDICE", Err)
            lErro = Comando_Fechar(lComando)

        Case 60777
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INDICES_DA_TELA_INEXISTENTES", Err, sNomeTela)
            lErro = Comando_Fechar(lComando)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142620)

    End Select

    Exit Function

End Function

'Retorna em colCampo os campos associados ao indice de tela dado por objTelaIndice
Public Function TelaIndiceCampos_Le(objTelaIndice As AdmTelaIndice, colCampoIndiceValor As AdmColCampoIndiceValor) As Long

Dim lConexao As Long
Dim lComando As Long
Dim lErro As Long
Dim sCampo As String

On Error GoTo Erro_TelaIndiceCampos_Le

    sCampo = String(STRING_TELAINDICECAMPO_NOME_CAMPO, 0)

    lConexao = GL_lConexaoDic
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 6598

    'seleciona no BD todos os campos associados a objTelaIndice
    lErro = Comando_Executar(lComando, "SELECT NomeCampo FROM TelaIndiceCampo WHERE NomeTela=? AND Indice=? ORDER BY Sequencial", sCampo, objTelaIndice.sNomeTela, objTelaIndice.iIndice)
    If lErro <> AD_SQL_SUCESSO Then Error 6599

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6600

    If lErro = AD_SQL_SEM_DADOS Then Error 60778

    'Preenche a coleção
    Do While lErro = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colCampoIndiceValor.Add sCampo, 0, 0
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6601

    Loop

    Call Comando_Fechar(lComando)

    TelaIndiceCampos_Le = SUCESSO

    Exit Function

Erro_TelaIndiceCampos_Le:

    TelaIndiceCampos_Le = Err

    Select Case Err

        Case 6598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6599, 6600, 6601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELAINDICECAMPO", Err)
            lErro = Comando_Fechar(lComando)

        Case 60778
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPOS_DO_INDICE_INEXISTENTES", Err, objTelaIndice.sNomeTela, objTelaIndice.iIndice)
            lErro = Comando_Fechar(lComando)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142621)

    End Select

    Exit Function

End Function
Public Function Compara_Campos_Indice(colCampoIndiceValor As AdmColCampoIndiceValor, colCampoValor As AdmColCampoValor, eComparacao As enumSetaComparacao) As Long
'Compara valores dos campos que compõem o índice anteriores e atuais

Dim lErro As Long
Dim objCampoIndiceValor As AdmCampoIndiceValor

On Error GoTo Erro_Compara_Campos_Indice
    
    eComparacao = SETA_COMP_IGUAL
    
    'Percorre os campos do Indice em uso
    For Each objCampoIndiceValor In colCampoIndiceValor

        'Verifica se valores atual e anterior são distintos
        If colCampoValor(objCampoIndiceValor.sCampo).vValor <> gcol_ST_CampoValor(objCampoIndiceValor.sCampo).vValor Then
            eComparacao = SETA_COMP_DIFERENTE
            Exit For
        End If
    
    Next
    
    Compara_Campos_Indice = SUCESSO

    Exit Function

Erro_Compara_Campos_Indice:

    Compara_Campos_Indice = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142622)

    End Select

    Exit Function

End Function

'Faz o movimento da seta e armazena em gcol_ST_CampoValor SE o comando já foi executado e é válido
Public Function Seta_Le(ByVal lComandoAtivo As Long, ByVal eSeta As enumSeta) As Long

Dim lComando As Long
Dim lErro As Long
Dim vCampo As Variant
Dim iIndice As Integer

On Error GoTo Erro_Seta_Le

    lComando = lComandoAtivo
    If lComando = 0 Then Error 6617
    
    Select Case eSeta
        
        Case BOTAO_PROXIMO

            'Busca o próximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro = AD_SQL_SEM_DADOS Then
                
                'Se não encontrou busca o último
                lErro = Comando_BuscarUltimo(lComando)
                If lErro = AD_SQL_SEM_DADOS Then Error 6618
                If lErro Then Error 25658
                
            End If
            
            If lErro Then Error 6619
            
        Case BOTAO_ANTERIOR

            'Busca o anterior
            lErro = Comando_BuscarAnterior(lComando)
            If lErro = AD_SQL_SEM_DADOS Then
                
                'Se não encontrou busca o primeiro
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro = AD_SQL_SEM_DADOS Then Error 6620
                If lErro Then Error 25659
                
            End If
            
            If lErro Then Error 6621
            
        Case BOTAO_PRIMEIRO
        
            'Busca o primeiro
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro = AD_SQL_SEM_DADOS Then Error 6637
            If lErro Then Error 6638
            
        Case BOTAO_ULTIMO
        
            'Busca o último
            lErro = Comando_BuscarUltimo(lComando)
            If lErro = AD_SQL_SEM_DADOS Then Error 6639
            If lErro Then Error 6640
            
    End Select

    'Passa os valores das variáveis bindadas vCampoSelect para gcol_ST_CampoValor
    For iIndice = 1 To gcol_ST_CampoValor.Count
        gcol_ST_CampoValor.Item(iIndice).vValor = vCampoSelect(iIndice)
    Next

    Seta_Le = SUCESSO

    Exit Function

Erro_Seta_Le:

    Seta_Le = Err

    Select Case Err

        Case 6617
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REFERENCIA_COMANDO_ABERTO", Err)

        Case 6618, 6620, 6637, 6639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_VAZIA", Err, gs_ST_TelaTabela)
            lErro = Comando_Fechar(lComando)
            gl_ST_ComandoSeta = 0

        Case 6619, 6621, 6638, 6640, 25658, 25659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, gs_ST_TelaTabela)
            lErro = Comando_Fechar(lComando)
            gl_ST_ComandoSeta = 0

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142623)

    End Select

    Exit Function

End Function
Public Function Seta_Le_AbreComando(eSeta As enumSeta, sTabela As String, colCampoIndiceValor As AdmColCampoIndiceValor, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro, objTelaIndice As AdmTelaIndice) As Long
'Lê seguindo a seta abrindo novo comando

Dim lErro As Long
Dim iIndice As Integer
Dim objCampoValor As AdmCampoValor

On Error GoTo Erro_Seta_Le_AbreComando

    'Fecha comando de seta que possa estar aberto
    lErro = ComandoSeta_Fechar2()
    
    'Cria novo comando conforme a escolha do botão
    'Usa colCampoIndiceValor para montar ORDER BY e colCampoValor para armazenar valores
    Select Case eSeta
    
        Case BOTAO_ANTERIOR, BOTAO_PROXIMO, BOTAO_CONSULTA
    
            lErro = Proximo_Anterior_Le(eSeta, sTabela, colCampoIndiceValor, colCampoValor, colSelecao)
            If lErro Then Error 6612
    
        Case BOTAO_PRIMEIRO, BOTAO_ULTIMO
    
            lErro = Primeiro_Ultimo_Le(eSeta, sTabela, colCampoIndiceValor, colCampoValor, colSelecao)
            If lErro Then Error 6622
    
    End Select
    
    'Limpa coleção global
    For iIndice = 1 To gcol_ST_CampoValor.Count
        gcol_ST_CampoValor.Remove (1)
    Next
    
    'Armazena colCampoValor na coleção global
    For Each objCampoValor In colCampoValor
        gcol_ST_CampoValor.Add objCampoValor.sCampo, objCampoValor.vValor, objCampoValor.iTamanhoString, objCampoValor.sCampo
    Next
    
    'Armazena o nome da tela que sofreu o (último) click
    gs_ST_TelaSetaClick = gobj_ST_TelaAtiva.Name
    'Armazena o nome da tabela associada
    gs_ST_TelaTabela = sTabela
    'Armazena o nome do indice usado
    gs_ST_TelaIndice = objTelaIndice.sNomeExterno
    
    Seta_Le_AbreComando = SUCESSO

    Exit Function

Erro_Seta_Le_AbreComando:

    Seta_Le_AbreComando = Err

    Select Case Err

        Case 6612, 6622  'Erro tratado na rotina chamada ou não faz nada (ausencia de seguintes)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142624)

    End Select

    Exit Function

End Function


'Monta e executa comando SQL. Campos do SELECT são dados por colCampoValor,
'do ORDER BY por colCampoIndice. Armazena o próximo registro em colCampoValor
Public Function Proximo_Anterior_Le(ByVal eSeta As enumSeta, ByVal sTabela As String, colCampoIndiceValor As AdmColCampoIndiceValor, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lConexao As Long
Dim lComando As Long
Dim lErro As Long
Dim sComandoSQL As String  'Comando Select
Dim objCampoValor As AdmCampoValor
Dim objCampoIndiceValor As AdmCampoIndiceValor
Dim objFiltro As AdmFiltro  'objeto do tipo Filtro
Dim iIndice As Integer
Dim iTotalElementos As Integer
'Dimensiona a ARRAY vCampoSelect (usada para receber os campos do SELECT)
'vCampoSelect tem Dim em adcapi.bas para ficar permanente enquanto o ACTIVEX estiver carregado.
'Não pode pertencer à classe porque é array e convém ter visibilidade apenas do ACTIVEX.
ReDim vCampoSelect(1 To colCampoValor.Count)
Dim vValorFiltro() As Variant
Dim sFiltro As String

On Error GoTo Erro_Proximo_Anterior_Le

    'Monta comando SQL, retorna o comando e o filtro usado no comando
    lErro = Monta_ComandoSQL(sTabela, colCampoIndiceValor, colCampoValor, colSelecao, sComandoSQL, sFiltro)
    If lErro Then Error 6746
   
    'Prepara vCampoSelect para receber valores pelo SELECT
    iIndice = 0
    For Each objCampoValor In colCampoValor
        
        iIndice = iIndice + 1
        
        If VarType(objCampoValor.vValor) = vbString Then
            vCampoSelect(iIndice) = Space(objCampoValor.iTamanhoString)
        Else
            vCampoSelect(iIndice) = objCampoValor.vValor
        End If
    Next
    
    If colSelecao.Count = 0 Then
    
        'ReDimensiona a ARRAY de valores de filtro com índice -1 se a coleção estiver vazia
        ReDim vValorFiltro(-1 To -1)
        
    Else
    
        'ReDimensiona a ARRAY de valores de filtro
        ReDim vValorFiltro(1 To colSelecao.Count)
        
        iIndice = 0
        'Preenche a ARRAY vValorFiltro
        For Each objFiltro In colSelecao
        
            iIndice = iIndice + 1
            'Preenche ARRAY vValorFiltro
            vValorFiltro(iIndice) = objFiltro.vValor
            
        Next
        
    End If

    lConexao = GL_lConexaoBrowse
    gl_ST_ComandoSeta = Comando_AbrirExt(lConexao)
    lComando = gl_ST_ComandoSeta
    If lComando = 0 Then Error 6603
    
    'Executa o SELECT (com ou sem filtro)
    lErro = Comando_Executar2(lComando, sComandoSQL, vCampoSelect(), vValorFiltro())
    If lErro <> AD_SQL_SUCESSO Then Error 6747
   
   'Passa os valores dos campos da tela e suas posições na ARRAY de variáveis bindadas (campos do índice) para a coleção colCampoIndiceValor
    For Each objCampoIndiceValor In colCampoIndiceValor

        'Compara os nomes de campos em colCampoValor e em objCampoIndiceValor
        For iIndice = 1 To colCampoValor.Count
            If colCampoValor.Item(iIndice).sCampo = objCampoIndiceValor.sCampo Then
                Exit For
            End If
        Next

        If iIndice > colCampoValor.Count Then Error 6602

        'Alimenta objCampoIndiceValor com o valor do campo em colCampoValor e com o índice correpondente na ARRAY bindada
        objCampoIndiceValor.vValor = colCampoValor.Item(iIndice).vValor
        objCampoIndiceValor.iIndice = iIndice

    Next
    
    'Posiciona o CURSOR no registro correspondente aos valores de colCampoIndiceValor
    lErro = Posiciona_Cursor(lComando, sTabela, sFiltro, vValorFiltro(), colCampoIndiceValor)
    If lErro = 6614 Then Error 25227  'Não tem registros na Tabela
    If lErro <> SUCESSO Then Error 6623
    
    'Busca o proximo ou anterior
    Select Case eSeta
        
        Case BOTAO_CONSULTA
            'já está OK
            
        Case BOTAO_ANTERIOR
            lErro = Comando_BuscarAnterior(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6625
                        
            'Se não encontrou
            If lErro = AD_SQL_SEM_DADOS Then
            
                'Busca o primeiro
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro = AD_SQL_SEM_DADOS Then Error 25230
                If lErro <> AD_SQL_SUCESSO Then Error 25231
            
            End If
    
        Case BOTAO_PROXIMO
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6627
            
            'Se não encontrou
            If lErro = AD_SQL_SEM_DADOS Then
            
                'Busca o último
                lErro = Comando_BuscarUltimo(lComando)
                If lErro = AD_SQL_SEM_DADOS Then Error 25232
                If lErro <> AD_SQL_SUCESSO Then Error 25233
            
            End If
   
    End Select

    'Passa os valores das variáveis bindadas vCampoSelect para colCampoValor
    For iIndice = 1 To colCampoValor.Count
        colCampoValor.Item(iIndice).vValor = vCampoSelect(iIndice)
    Next
    
    Proximo_Anterior_Le = SUCESSO

    Exit Function

Erro_Proximo_Anterior_Le:

    Proximo_Anterior_Le = Err

    Select Case Err

        Case 6602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INEXISTE_CAMPO_TELA_IGUAL_CAMPO_INDICE", Err, objCampoIndiceValor.sCampo)

        Case 6603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6625, 6627, 6747, 25231, 25233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabela)
            lErro = Comando_Fechar(lComando)
            gl_ST_ComandoSeta = 0
            
        Case 6623
            lErro = Comando_Fechar(lComando)
            gl_ST_ComandoSeta = 0
            
        Case 6746  'Tratado na rotina chamada

        Case 25227, 25230, 25232   'Não tem registros na TABELA
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_VAZIA", Err, sTabela)
            lErro = Comando_Fechar(lComando)
            gl_ST_ComandoSeta = 0
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142625)

    End Select

    Exit Function

End Function
Private Function Posiciona_Cursor(ByVal lComando As Long, ByVal sTabela As String, ByVal sFiltro As String, vValorFiltro() As Variant, colCampoIndiceValor As AdmColCampoIndiceValor) As Long

Dim lConexao As Long
Dim lComando2 As Long
Dim lErro As Long
Dim objCampoIndiceValor As AdmCampoIndiceValor
Dim iIndice As Integer
Dim iExpoenteMaximo As Integer
Dim iExpoente As Integer
Dim lPosicao As Long
Dim iComparacao As Integer
Dim lTotalRegistros As Long
Dim sComandoSQL As String
Dim vTotalRegistros(1 To 1)

On Error GoTo Erro_Posiciona_Cursor
    
    lConexao = GL_lConexaoBrowse
    
    lComando2 = Comando_AbrirExt(lConexao)
    If lComando2 = 0 Then Error 6605
    
    If Len(sFiltro) = 0 Then 'Não existem filtros
    
        sComandoSQL = "SELECT COUNT(*) FROM " & sTabela
        'sComandoSQL = "SELECT DISTINCT COUNT(*) FROM " & sTabela

        lErro = Comando_Executar(lComando2, sComandoSQL, lTotalRegistros)
        If lErro <> AD_SQL_SUCESSO Then Error 6606
        
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro Then Error 6629
        
    Else 'Existem filtros
    
        'Prepara vTotalRegistros para receber o total (tipa)
        'Este passo é necessário porque Comando_Executar2 trabalha com ARRAYS de variants para bindar
        vTotalRegistros(1) = lTotalRegistros
        sComandoSQL = "SELECT COUNT(*) FROM " & sTabela & " WHERE " & sFiltro
        'sComandoSQL = "SELECT DISTINCT COUNT(*) FROM " & sTabela & " WHERE " & sFiltro
        
        lErro = Comando_Executar2(lComando2, sComandoSQL, vTotalRegistros(), vValorFiltro())
        If lErro <> AD_SQL_SUCESSO Then Error 6604
        
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro Then Error 6632
        
        'Volta da ARRAY unitária para a variável tipo long
        lTotalRegistros = vTotalRegistros(1)
    
    End If
    
    If lTotalRegistros = 0 Then Error 6614

    iExpoenteMaximo = IntLog2(lTotalRegistros)
    
    If 2& ^ iExpoenteMaximo = lTotalRegistros Then
        iExpoenteMaximo = iExpoenteMaximo + 1
    End If
    
    iExpoente = iExpoenteMaximo - 1
    lPosicao = CLng(2& ^ CLng(iExpoente))

    For iIndice = 1 To iExpoenteMaximo + 1
      
        If lPosicao > lTotalRegistros Then
            
            iExpoente = iExpoente - 1
            lPosicao = lPosicao - CLng(2& ^ CLng(iExpoente))
            
        Else
            
            'Este comando alimenta array bindada vCampoSelect (definida em adcapi.bas)
            lErro = Comando_BuscarAbsoluto(lComando, lPosicao)
            If lErro Then Error 6628
            
            iComparacao = IGUAL
            
            'Compara variáveis bindadas com valores passados em colCampoIndiceValor
            For Each objCampoIndiceValor In colCampoIndiceValor
                
                'Se for comparacao de STRINGS usa StrComp (por causa da ordem lexicografica)
                If VarType(objCampoIndiceValor.vValor) = vbString Then
                
                    iComparacao = StrComp(UCase(vCampoSelect(objCampoIndiceValor.iIndice)), UCase(objCampoIndiceValor.vValor), gi_ST_Ordem_StrCmp)
                
                Else 'Se nao for STRING, usa comparacao usual
                
                    Select Case vCampoSelect(objCampoIndiceValor.iIndice)
                    
                        'Encontrou registro MAIOR
                        Case Is > objCampoIndiceValor.vValor
                            iComparacao = MAIOR
                    
                        'Encontrou registro MENOR
                        Case Is < objCampoIndiceValor.vValor
                            iComparacao = MENOR
                    
                    End Select
                    
                End If
                    
                'Se houve desigualdade sai do For
                If iComparacao <> IGUAL Then Exit For
            
            Next
        
            Select Case iComparacao
            
                Case IGUAL
                    Exit For
                    
                Case MAIOR
                    iExpoente = iExpoente - 1
                    lPosicao = lPosicao - CLng(2& ^ CLng(iExpoente))
                    
                Case MENOR
                    iExpoente = iExpoente - 1
                    If iExpoente < 0 Then iExpoente = 0
                    lPosicao = lPosicao + CLng(2& ^ CLng(iExpoente))
                    
            End Select
        
        End If
    
    Next
    
    lErro = Comando_Fechar(lComando2)
    
    Posiciona_Cursor = SUCESSO
       
    Exit Function
    
Erro_Posiciona_Cursor:

    Posiciona_Cursor = Err

    Select Case Err

        Case 6605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6604, 6606, 6628, 6629, 6632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabela)
            lErro = Comando_Fechar(lComando2)

        Case 6614   'Será tratado pela rotina chamadora
            lErro = Comando_Fechar(lComando2)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142626)

    End Select

    Exit Function

End Function
Private Function Monta_ComandoSQL(sTabela As String, colCampoIndiceValor As AdmColCampoIndiceValor, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro, sComandoSQL As String, sFiltro As String) As Long

Dim lErro As Long
Dim sCampos As String 'Campos do Select
Dim sOrdem As String 'Ordem do Select
Dim sOperador As String 'Operador em String
Dim objCampoValor As AdmCampoValor
Dim objCampoIndiceValor As AdmCampoIndiceValor
Dim objFiltro As AdmFiltro

On Error GoTo Erro_Monta_ComandoSQL
    
    'Monta sCampos
    For Each objCampoValor In colCampoValor
        sCampos = sCampos & objCampoValor.sCampo & ", "
    Next

    sCampos = Left(sCampos, Len(sCampos) - 2)

    'Monta sOrdem
    For Each objCampoIndiceValor In colCampoIndiceValor
        sOrdem = sOrdem & objCampoIndiceValor.sCampo & ", "
    Next

    sOrdem = Left(sOrdem, Len(sOrdem) - 2)
    
    If colSelecao.Count = 0 Then
    
        'Filtro vazio
        sFiltro = ""
    
        'Monta sComandoSQL
        sComandoSQL = "SELECT " & sCampos & " FROM " & sTabela & " ORDER BY " & sOrdem
        'sComandoSQL = "SELECT DISTINCT " & sCampos & " FROM " & sTabela & " ORDER BY " & sOrdem
    Else
        
        'Monta sFiltro
        For Each objFiltro In colSelecao
        
            'Transforma o operador em STRING
            Select Case objFiltro.eOperador
            
                Case OP_IGUAL
                    sOperador = "="
                Case OP_MAIOR
                    sOperador = ">"
                Case OP_MENOR
                    sOperador = "<"
                Case OP_MAIOR_OU_IGUAL
                    sOperador = ">="
                Case OP_MENOR_OU_IGUAL
                    sOperador = "<="
                Case OP_DIFERENTE
                    sOperador = "<>"
            
            End Select

            sFiltro = sFiltro & objFiltro.sCampo & " " & sOperador & " ?" & " AND "
        Next
    
        sFiltro = Left(sFiltro, Len(sFiltro) - 5)
        
        'Monta sComandoSQL
        sComandoSQL = "SELECT " & sCampos & " FROM " & sTabela & " WHERE " & sFiltro & " ORDER BY " & sOrdem
        'sComandoSQL = "SELECT DISTINCT " & sCampos & " FROM " & sTabela & " WHERE " & sFiltro & " ORDER BY " & sOrdem
    End If
    
    Monta_ComandoSQL = SUCESSO

    Exit Function

Erro_Monta_ComandoSQL:

    Monta_ComandoSQL = Err

    lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142627)

    Exit Function

End Function
'Monta e executa comando SQL. Campos do SELECT são dados por colCampoValor,
'do ORDER BY por colCampoIndice. Armazena o próximo registro em colCampoValor
Public Function Primeiro_Ultimo_Le(ByVal eSeta As enumSeta, ByVal sTabela As String, colCampoIndiceValor As AdmColCampoIndiceValor, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lConexao As Long
Dim lComando As Long
Dim lErro As Long
Dim sComandoSQL As String  'Comando Select
Dim objCampoValor As AdmCampoValor
Dim objCampoIndiceValor As AdmCampoIndiceValor
Dim objFiltro As AdmFiltro
Dim iIndice As Integer
'Dimensiona a ARRAY vCampoSelect (usada para receber os campos do SELECT)
'vCampoSelect tem Dim em adcapi.bas para ficar permanente enquanto o ACTIVEX estiver carregado.
'Não pode pertencer à classe porque é array e convém ter visibilidade apenas do ACTIVEX.
ReDim vCampoSelect(1 To colCampoValor.Count)
Dim vValorFiltro() As Variant
Dim sFiltro As String

On Error GoTo Erro_Primeiro_Ultimo_Le
    
    'Monta comando SQL, retorna o comando e o filtro usado no comando
    lErro = Monta_ComandoSQL(sTabela, colCampoIndiceValor, colCampoValor, colSelecao, sComandoSQL, sFiltro)
    If lErro Then Error 6745
  
    'Prepara vCampoSelect para receber valores pelo SELECT
    iIndice = 0
    For Each objCampoValor In colCampoValor
        
        iIndice = iIndice + 1
        
        If VarType(objCampoValor.vValor) = vbString Then
            vCampoSelect(iIndice) = Space(objCampoValor.iTamanhoString)
        Else
            vCampoSelect(iIndice) = objCampoValor.vValor
        End If
    
    Next

    If colSelecao.Count = 0 Then
    
        'ReDimension ARRAY de valores de filtro com índice -1
        ReDim vValorFiltro(-1 To -1)
    
    Else 'existe coleção de filtros
    
        'ReDimensiona a ARRAY de valores de filtro
        ReDim vValorFiltro(1 To colSelecao.Count)
        
        iIndice = 0
        'Preenche a ARRAY vValorFiltro
        For Each objFiltro In colSelecao
        
            iIndice = iIndice + 1
            'Preenche ARRAY vValorFiltro
            vValorFiltro(iIndice) = objFiltro.vValor
            
        Next
        
    End If
            
    lConexao = GL_lConexaoBrowse
    gl_ST_ComandoSeta = Comando_AbrirExt(lConexao)
    lComando = gl_ST_ComandoSeta
    If lComando = 0 Then Error 6631

    'Executa o SELECT (com ou sem filtro)
    lErro = Comando_Executar2(lComando, sComandoSQL, vCampoSelect(), vValorFiltro())
    If lErro <> AD_SQL_SUCESSO Then Error 6744

    Select Case eSeta

        Case BOTAO_PRIMEIRO
            'Busca o primeiro
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro = AD_SQL_SEM_DADOS Then Error 25228
            If lErro <> AD_SQL_SUCESSO Then Error 6634

        Case BOTAO_ULTIMO
            'Busca o último
            lErro = Comando_BuscarUltimo(lComando)
            If lErro = AD_SQL_SEM_DADOS Then Error 25229
            If lErro <> AD_SQL_SUCESSO Then Error 6636

    End Select

    If lErro = AD_SQL_SUCESSO Then

        'Passa os valores das variáveis bindadas vCampoSelect para colCampoValor
        For iIndice = 1 To colCampoValor.Count
            colCampoValor.Item(iIndice).vValor = vCampoSelect(iIndice)
        Next
    
    End If

    Primeiro_Ultimo_Le = SUCESSO

    Exit Function

Erro_Primeiro_Ultimo_Le:

    Primeiro_Ultimo_Le = Err

    Select Case Err

        Case 6631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6634, 6636, 6744
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabela)
            lErro = Comando_Fechar(lComando)
            gl_ST_ComandoSeta = 0
            
        Case 6745  'Já tratado na rotina chamada

        Case 25228, 25229  'Não tem registros na TABELA
           lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_VAZIA", Err, sTabela)
           lErro = Comando_Fechar(lComando)
           gl_ST_ComandoSeta = 0
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142628)

    End Select

    Exit Function

End Function

'Retorna o menor inteiro maior ou igual ao log2(dNumero)
Private Function IntLog2(lNumero As Long) As Long

    Dim dResultado As Double

        dResultado = Log(CDbl(lNumero)) / Log(2#)

        If dResultado > Int(dResultado) Then
            IntLog2 = CInt(Int(dResultado) + 1)
        Else
            IntLog2 = CInt(Int(dResultado))
        End If

End Function

