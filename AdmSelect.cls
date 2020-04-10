VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt1 Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt1 Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt1 Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long
Private Declare Function Conexao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Private Declare Function Conexao_FecharExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long

Const STRING_BROWSEUSUARIO_NOMETELA = 50
Const STRING_BROWSEUSUARIO_CODUSUARIO = 50
Const STRING_TELAUSUARIO_NOMETELA = 50
Const STRING_TELAUSUARIO_CODUSUARIO = 50

Function BrowseUsuario_Le(ByVal objBrowseUsuario As AdmBrowseUsuario) As Long
'le os campos da tabela browseusuario e coloca os resultados em objBrowseUsuario

Dim lErro As Long
Dim lComando As Long
Dim tBrowseUsuario As typeBrowseUsuario
Dim colBrowseUsuario As New Collection

On Error GoTo Erro_BrowseUsuario_Le

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 91289

    tBrowseUsuario.sNomeTela = String(STRING_BROWSEUSUARIO_NOMETELA, 0)
    tBrowseUsuario.sCodUsuario = String(STRING_BROWSEUSUARIO_CODUSUARIO, 0)

    lErro = Comando_Executar(lComando, "SELECT NomeTela, CodUsuario, Topo, Esquerda, Largura, Altura FROM BrowseUsuario WHERE NomeTela=? AND CodUsuario=?", tBrowseUsuario.sNomeTela, tBrowseUsuario.sCodUsuario, tBrowseUsuario.lTopo, tBrowseUsuario.lEsquerda, tBrowseUsuario.lLargura, tBrowseUsuario.lAltura, objBrowseUsuario.sNomeTela, objBrowseUsuario.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 91291

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91292

    If lErro = AD_SQL_SEM_DADOS Then gError 91302
        
    objBrowseUsuario.sNomeTela = tBrowseUsuario.sNomeTela
    objBrowseUsuario.sCodUsuario = tBrowseUsuario.sCodUsuario
    objBrowseUsuario.lTopo = tBrowseUsuario.lTopo
    objBrowseUsuario.lEsquerda = tBrowseUsuario.lEsquerda
    objBrowseUsuario.lLargura = tBrowseUsuario.lLargura
    objBrowseUsuario.lAltura = tBrowseUsuario.lAltura
        
    Call Comando_Fechar(lComando)
    
    BrowseUsuario_Le = SUCESSO
    
    Exit Function

Erro_BrowseUsuario_Le:

    BrowseUsuario_Le = gErr

    Select Case gErr
   
        Case 91289
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case 91291, 91292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEUSUARIO", gErr)

        Case 91302
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142504)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BrowseUsuarioCampo_Le(ByVal sNomeTela As String, ByVal sCodUsuario As String, ByVal colBrowseUsuarioCampo As Collection) As Long
'le todos os campos da tela de browse para o usuario especificado e coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim tBrowseUsuarioCampo As typeBrowseUsuarioCampo
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
    
On Error GoTo Erro_BrowseUsuarioCampo_Le

    tBrowseUsuarioCampo.sCodUsuario = String(STRING_USUARIO, 0)
    tBrowseUsuarioCampo.sNome = String(STRING_NOME_CAMPO, 0)
    tBrowseUsuarioCampo.sNomeArq = String(STRING_NOME_TABELA, 0)
    tBrowseUsuarioCampo.sNomeTela = String(STRING_NOME_TELA, 0)
    tBrowseUsuarioCampo.sTitulo = String(STRING_TITULO_CAMPO, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 9031

    lErro = Comando_Executar(lComando, "SELECT NomeTela, CodUsuario, NomeArq, Nome, PosicaoTela, Titulo, Largura FROM BrowseUsuarioCampo WHERE NomeTela=? AND CodUsuario=? ORDER BY PosicaoTela", tBrowseUsuarioCampo.sNomeTela, tBrowseUsuarioCampo.sCodUsuario, tBrowseUsuarioCampo.sNomeArq, tBrowseUsuarioCampo.sNome, tBrowseUsuarioCampo.iPosicaoTela, tBrowseUsuarioCampo.sTitulo, tBrowseUsuarioCampo.lLargura, sNomeTela, sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 9032
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9033
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objBrowseUsuarioCampo = New AdmBrowseUsuarioCampo
    
        objBrowseUsuarioCampo.sNomeTela = tBrowseUsuarioCampo.sNomeTela
        objBrowseUsuarioCampo.sCodUsuario = tBrowseUsuarioCampo.sCodUsuario
        objBrowseUsuarioCampo.sNomeArq = tBrowseUsuarioCampo.sNomeArq
        objBrowseUsuarioCampo.sNome = tBrowseUsuarioCampo.sNome
        objBrowseUsuarioCampo.iPosicaoTela = tBrowseUsuarioCampo.iPosicaoTela
        objBrowseUsuarioCampo.sTitulo = tBrowseUsuarioCampo.sTitulo
        objBrowseUsuarioCampo.lLargura = tBrowseUsuarioCampo.lLargura

        colBrowseUsuarioCampo.Add objBrowseUsuarioCampo
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9034
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    BrowseUsuarioCampo_Le = SUCESSO
    
    Exit Function
    
    
Erro_BrowseUsuarioCampo_Le:

    BrowseUsuarioCampo_Le = Err

    Select Case Err
    
        Case 9031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 9032, 9033, 9034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEUSUARIOCAMPO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142505)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BrowseUsuarioOrdenacao_Le(ByVal objBrowseUsuarioOrdenacao As AdmBrowseUsuarioOrdenacao) As Long
'le a ordenacao escolhida pelo usuario

Dim lComando As Long
Dim lErro As Long
Dim tBrowseUsuarioOrdenacao As typeBrowseUsuarioOrdenacao
    
On Error GoTo Erro_BrowseUsuarioOrdenacao_Le

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 9041

    tBrowseUsuarioOrdenacao.sSelecaoSQL1 = String(STRING_SELECAO_SQL, 0)
    tBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario = String(STRING_SELECAO_SQL, 0)
    tBrowseUsuarioOrdenacao.sNomeIndice = String(STRING_NOME_INDICE, 0)

    lErro = Comando_Executar(lComando, "SELECT Indice, SelecaoSQL1, SelecaoSQL1Usuario, NomeIndice FROM BrowseUsuarioOrdenacao WHERE NomeTela=? AND CodUsuario=?", tBrowseUsuarioOrdenacao.iIndice, tBrowseUsuarioOrdenacao.sSelecaoSQL1, tBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario, tBrowseUsuarioOrdenacao.sNomeIndice, objBrowseUsuarioOrdenacao.sNomeTela, objBrowseUsuarioOrdenacao.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 9042
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9043
    
    If lErro = AD_SQL_SEM_DADOS Then Error 9044
    
    objBrowseUsuarioOrdenacao.iIndice = tBrowseUsuarioOrdenacao.iIndice
    objBrowseUsuarioOrdenacao.sSelecaoSQL1 = tBrowseUsuarioOrdenacao.sSelecaoSQL1
    objBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario = tBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario
    objBrowseUsuarioOrdenacao.sNomeIndice = tBrowseUsuarioOrdenacao.sNomeIndice
    
    Call Comando_Fechar(lComando)
    
    BrowseUsuarioOrdenacao_Le = SUCESSO
    
    Exit Function
    
Erro_BrowseUsuarioOrdenacao_Le:

    BrowseUsuarioOrdenacao_Le = Err

    Select Case Err
    
        Case 9041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 9042, 9043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEUSUARIOORDENACAO", Err)
        
        Case 9044
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142506)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BrowseIndice_Le(ByVal sNomeTela As String, ByVal colBrowseIndice As Collection) As Long
'le todos os indices referentes a tela especificada e coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim tBrowseIndice As typeBrowseIndice
Dim objBrowseIndice As AdmBrowseIndice
    
On Error GoTo Erro_BrowseIndice_Le

    tBrowseIndice.sNomeTela = String(STRING_NOME_TELA, 0)
    tBrowseIndice.sNomeIndice = String(STRING_NOME_INDICE, 0)
    tBrowseIndice.sOrdenacaoSQL = String(STRING_ORDENACAO_SQL, 0)
    tBrowseIndice.sSelecaoSQL = String(STRING_SELECAO_SQL, 0)

    lComando = 0
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 9048

    lErro = Comando_Executar(lComando, "SELECT Indice, NomeIndice, OrdenacaoSQL, SelecaoSQL FROM BrowseIndice WHERE NomeTela=?", tBrowseIndice.iIndice, tBrowseIndice.sNomeIndice, tBrowseIndice.sOrdenacaoSQL, tBrowseIndice.sSelecaoSQL, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then Error 9049
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9050
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objBrowseIndice = New AdmBrowseIndice
    
        objBrowseIndice.sNomeTela = sNomeTela
        objBrowseIndice.iIndice = tBrowseIndice.iIndice
        objBrowseIndice.sNomeIndice = tBrowseIndice.sNomeIndice
        objBrowseIndice.sOrdenacaoSQL = tBrowseIndice.sOrdenacaoSQL
        objBrowseIndice.sSelecaoSQL = tBrowseIndice.sSelecaoSQL

        colBrowseIndice.Add objBrowseIndice
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9051
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    BrowseIndice_Le = SUCESSO
    
    Exit Function
    
Erro_BrowseIndice_Le:

    BrowseIndice_Le = Err

    Select Case Err
    
        Case 9048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 9049, 9050, 9051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEINDICE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142507)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Browse_Prepara_Comando_SQL(ByVal sSQL As String, ByVal iNumCampo As Integer, ByVal lComando As Long, avCampo() As Variant, ByVal colSelecao As Collection, ByVal iBind As Integer) As Long
'prepara o comando SQL passado como parametro

Dim lErro As Long
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim objCampo As New AdmCampos
Dim objValorCampo As AdmValorCampo
Dim iCampo As Integer
Dim vSelecao As Variant

On Error GoTo Erro_Browse_Prepara_Comando_SQL


    lErro = Comando_PrepararInt1(lComando, sSQL)
    If lErro <> AD_SQL_SUCESSO Then Error 9057
    
    
    If iBind = 1 Then
    
        For iCampo = 1 To iNumCampo
            
            lErro = Comando_BindVarInt1(lComando, avCampo(iCampo))
            If lErro <> AD_SQL_SUCESSO Then Error 9065
        
        Next
    End If
        
        iCampo = iNumCampo
        
        If Not (colSelecao Is Nothing) Then
        
            For Each vSelecao In colSelecao
            
                iCampo = iCampo + 1
                
                avCampo(iCampo) = vSelecao
                
                lErro = Comando_BindVarInt1(lComando, avCampo(iCampo))
                If lErro <> AD_SQL_SUCESSO Then Error 9251
                
            Next
            
        End If
    
'    Else
'
'        iCampo = iNumCampo
'
'        If Not (colSelecao Is Nothing) Then
'
'            For Each vSelecao In colSelecao
'
'                iCampo = iCampo + 1
'
'                avCampo(iCampo) = vSelecao
'
'            Next
'
'        End If
'
'    End If
    
    lErro = Comando_ExecutarInt1(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9072
    
    Browse_Prepara_Comando_SQL = SUCESSO
    
    Exit Function

Erro_Browse_Prepara_Comando_SQL:

    Browse_Prepara_Comando_SQL = Err
    
    Select Case Err
    
        Case 9057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", Err, sSQL)
        
        Case 9065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BIND_CAMPO", Err, objCampo.sNomeArq, objCampo.sNome)
            
        Case 9072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", Err, sSQL)
            
        Case 9251
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BIND_CAMPO1", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142508)
        
    End Select

    Exit Function
        
End Function

Function Campos_Le(ByVal objCampo As AdmCampos) As Long
'le a ordenacao escolhida pelo usuario

Dim lComando As Long
Dim lErro As Long
Dim tCampos As typeCampos
Dim objCampo1 As AdmCampos
    
On Error GoTo Erro_Campos_Le

    tCampos.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)
    tCampos.sFormatacao = String(STRING_FORMATACAO_CAMPO, 0)
    tCampos.sNome = String(STRING_NOME_CAMPO, 0)
    tCampos.sNomeArq = String(STRING_NOME_TABELA, 0)
    tCampos.sTituloEntradaDados = String(STRING_TITULO_ENTRADA_DADOS_CAMPO, 0)
    tCampos.sTituloGrid = String(STRING_TITULO_GRID_CAMPO, 0)
    tCampos.sValDefault = String(STRING_VALOR_DEFAULT_CAMPO, 0)
    tCampos.sValidacao = String(STRING_VALIDACAO_CAMPO, 0)

    lComando = 0

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 9058

    lErro = Comando_Executar(lComando, "SELECT NomeArq, Nome, Descricao, Obrigatorio, Imexivel, Ativo, ValDefault, Validacao, Formatacao, Tipo, Tamanho, Precisao, Decimais, TamExibicao, TituloEntradaDados, TituloGrid, Subtipo, Alinhamento FROM Campos WHERE NomeArq=? AND Nome=?", tCampos.sNomeArq, tCampos.sNome, tCampos.sDescricao, tCampos.iObrigatorio, tCampos.iImexivel, tCampos.iAtivo, tCampos.sValDefault, tCampos.sValidacao, tCampos.sFormatacao, tCampos.iTipo, tCampos.iTamanho, tCampos.iPrecisao, tCampos.iDecimais, tCampos.iTamExibicao, tCampos.sTituloEntradaDados, tCampos.sTituloGrid, tCampos.iSubTipo, tCampos.iAlinhamento, objCampo.sNomeArq, objCampo.sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 9059

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9060

    If lErro = AD_SQL_SEM_DADOS Then Error 9061

    objCampo.iAtivo = tCampos.iAtivo
    objCampo.iDecimais = tCampos.iDecimais
    objCampo.iImexivel = tCampos.iImexivel
    objCampo.iObrigatorio = tCampos.iObrigatorio
    objCampo.iPrecisao = tCampos.iPrecisao
    objCampo.iTamanho = tCampos.iTamanho
    objCampo.iTamExibicao = tCampos.iTamExibicao
    objCampo.iTipo = tCampos.iTipo
    objCampo.sDescricao = tCampos.sDescricao
    objCampo.sFormatacao = tCampos.sFormatacao
    objCampo.sTituloEntradaDados = tCampos.sTituloEntradaDados
    objCampo.sTituloGrid = tCampos.sTituloGrid
    objCampo.sValDefault = tCampos.sValDefault
    objCampo.sValidacao = tCampos.sValidacao
    objCampo.iSubTipo = tCampos.iSubTipo
    objCampo.iAlinhamento = tCampos.iAlinhamento

    Call Comando_Fechar(lComando)
    
' retirado pois a otimização de carga dos campos para uma colecao no inicio do programa foi retirado
'    Set objCampo = gcolCampos.Item(objCampo.sNomeArq + objCampo.sNome)
    
    Campos_Le = SUCESSO
    
    Exit Function
    
Erro_Campos_Le:

    Campos_Le = Err

    Select Case Err
    
        Case 9058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 9059, 9060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOS", Err)
        
        Case 9061
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142509)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Campos_Le1(ByVal objCampo As AdmCampos, ByVal lComando As Long) As Long
'le o titulo do grid do campo

Dim lErro As Long
Dim tCampos As typeCampos
' retirado pois a otimização de carga dos campos para uma colecao no inicio do programa foi retirado
'Dim objCampo1 As New AdmCampos
    
On Error GoTo Erro_Campos_Le1

    tCampos.sTituloGrid = String(STRING_TITULO_GRID_CAMPO, 0)

    lErro = Comando_Executar(lComando, "SELECT TituloGrid FROM Campos WHERE NomeArq=? AND Nome=?", tCampos.sTituloGrid, objCampo.sNomeArq, objCampo.sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 9177

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9178

    If lErro = AD_SQL_SEM_DADOS Then Error 9179

    objCampo.sTituloGrid = tCampos.sTituloGrid
    
' retirado pois a otimização de carga dos campos para uma colecao no inicio do programa foi retirado
'    Set objCampo1 = gcolCampos.Item(objCampo.sNomeArq + objCampo.sNome)
    
'    objCampo.sTituloGrid = objCampo1.sTituloGrid
    
    Campos_Le1 = SUCESSO
    
    Exit Function
    
Erro_Campos_Le1:

    Campos_Le1 = Err

    Select Case Err
    
        Case 9177, 9178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOS", Err)
        
        Case 9179
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142510)
        
    End Select

    Exit Function

End Function

Function Campos_Le2(ByVal objCampo As AdmCampos, ByVal lComando As Long) As Long
'le o tipo e o tamanho do campo

Dim lErro As Long
Dim tCampos As typeCampos
Dim objCampo1 As AdmCampos
    
On Error GoTo Erro_Campos_Le2

    lErro = Comando_Executar(lComando, "SELECT Tipo, Tamanho, Subtipo, Alinhamento FROM Campos WHERE NomeArq=? AND Nome=?", tCampos.iTipo, tCampos.iTamanho, tCampos.iSubTipo, tCampos.iAlinhamento, objCampo.sNomeArq, objCampo.sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 9182

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9183

    If lErro = AD_SQL_SEM_DADOS Then Error 9184

    objCampo.iTamanho = tCampos.iTamanho
    objCampo.iTipo = tCampos.iTipo
    objCampo.iSubTipo = tCampos.iSubTipo
    objCampo.iAlinhamento = tCampos.iAlinhamento
    
' retirado pois a otimização de carga dos campos para uma colecao no inicio do programa foi retirado
'    Set objCampo1 = gcolCampos.Item(objCampo.sNomeArq + objCampo.sNome)
'
'    objCampo.iTamanho = objCampo1.iTamanho
'    objCampo.iTipo = objCampo1.iTipo
'    objCampo.iSubTipo = objCampo1.iSubTipo
'    objCampo.iAlinhamento = objCampo1.iAlinhamento
    
    Campos_Le2 = SUCESSO
    
    Exit Function
    
Erro_Campos_Le2:

    Campos_Le2 = Err

    Select Case Err
    
        Case 9182, 9183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOS", Err)
        
        Case 9184
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142511)
        
    End Select

    Exit Function

End Function

Function Browse_Le_NumRegistros(lNumReg As Long, ByVal lComando As Long, ByVal sSQL As String, avCampo() As Variant) As Long
'retorna o número de registros contidos na tabela sNomeTabela

Dim lErro As Long

On Error GoTo Erro_Browse_Le_NumRegistros

    If lComando = 0 Then Error 9008

    'Le o número de registro na tabela
    lErro = Comando_ReExecutar(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9009

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9010
    
    lNumReg = avCampo(1)
    
    'mario 1/4/99. Colocado para não deixar que a scrollbar esconda um registro.
    lNumReg = lNumReg + 1
    
    Browse_Le_NumRegistros = SUCESSO

    Exit Function

Erro_Browse_Le_NumRegistros:

    Browse_Le_NumRegistros = Err

    Select Case Err

        Case 9008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9009, 9010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142512)

    End Select

    Exit Function

End Function

Function Browse_Le_Posicionado(ByVal colReg As Collection, ByVal lPosicao As Long, ByVal lNumReg As Long, ByVal lComando As Long, ByVal colValorCampo As Collection, ByVal sSQL As String, avCampo() As Variant) As Long
'le os registros lNumReg a partir da posicao lPosicao e coloca-os e colReg.

Dim lErro As Long
Dim colRegCampo As Collection
Dim lNumRegLido As Long
Dim objValorCampo As AdmValorCampo
Dim iCampo As Integer

On Error GoTo Erro_Browse_Le_Posicionado

    lErro = Comando_ReExecutar(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9014

    lErro = Comando_BuscarAbsoluto(lComando, lPosicao)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9015

    lNumRegLido = 1

    Do While lErro <> AD_SQL_SEM_DADOS And lNumRegLido <= lNumReg

        Set colRegCampo = New Collection

        iCampo = 0

        For Each objValorCampo In colValorCampo
        
            iCampo = iCampo + 1
            
            colRegCampo.Add avCampo(iCampo)
            
        Next
        
        colReg.Add colRegCampo

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9016

        lNumRegLido = lNumRegLido + 1
    Loop
    
    Browse_Le_Posicionado = SUCESSO

    Exit Function

Erro_Browse_Le_Posicionado:

    Browse_Le_Posicionado = Err

    Select Case Err
    
        Case 9014, 9015, 9016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", Err, sSQL)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142513)

    End Select
    
    Exit Function

End Function

Function GrupoBrowseCampo_Le(ByVal sCodGrupo As String, ByVal sNomeTela As String, ByVal colGrupoBrowseCampo As Collection) As Long
'le os campos selecionados para o grupo x tela em questão em coloca-os na coleção

Dim lComando As Long
Dim lErro As Long
Dim tGrupoBrowseCampo As typeGrupoBrowseCampo
Dim objGrupoBrowseCampo As AdmGrupoBrowseCampo
    
On Error GoTo Erro_GrupoBrowseCampo_Le

    tGrupoBrowseCampo.sCodGrupo = String(STRING_GRUPO, 0)
    tGrupoBrowseCampo.sNome = String(STRING_NOME_CAMPO, 0)
    tGrupoBrowseCampo.sNomeArq = String(STRING_NOME_TABELA, 0)
    tGrupoBrowseCampo.sNomeTela = String(STRING_NOME_TELA, 0)

    lComando = 0
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 9077

    lErro = Comando_Executar(lComando, "SELECT NomeArq, Nome FROM GrupoBrowseCampo WHERE CodGrupo=? AND NomeTela=?", tGrupoBrowseCampo.sNomeArq, tGrupoBrowseCampo.sNome, sCodGrupo, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then Error 9078
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9079
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objGrupoBrowseCampo = New AdmGrupoBrowseCampo
    
        objGrupoBrowseCampo.sCodGrupo = sCodGrupo
        objGrupoBrowseCampo.sNomeTela = sNomeTela
        objGrupoBrowseCampo.sNomeArq = tGrupoBrowseCampo.sNomeArq
        objGrupoBrowseCampo.sNome = tGrupoBrowseCampo.sNome
        
        colGrupoBrowseCampo.Add objGrupoBrowseCampo
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9080
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    GrupoBrowseCampo_Le = SUCESSO
    
    Exit Function
    
Erro_GrupoBrowseCampo_Le:

    GrupoBrowseCampo_Le = Err

    Select Case Err
    
        Case 9077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 9078, 9079, 9080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPOBROWSECAMPO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142514)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BrowseIndiceSegmentos_Le(ByVal sNomeTela As String, ByVal iIndice As Integer, ByVal colBrowseIndiceSegmentos As Collection, ByVal colBrowseIndice As Collection) As Long
'le todos os segmentos do indice especificado e coloca os resultados na coleção

Dim lErro As Long
Dim objBrowseIndiceSegmentos As AdmBrowseIndiceSegmentos
Dim objBrowseIndice As AdmBrowseIndice
Dim iPosicaoCampo As Integer
Dim iPosicaoFimCampo As Integer
Dim iPosicao As Integer
    
On Error GoTo Erro_BrowseIndiceSegmentos_Le

    For Each objBrowseIndice In colBrowseIndice
        If objBrowseIndice.iIndice = iIndice Then Exit For
    Next
    
    iPosicaoCampo = 0
    iPosicaoFimCampo = 1
    iPosicao = 1
    
    Do While iPosicaoFimCampo <> 0
    
        iPosicaoCampo = iPosicaoCampo + 1
    
        Set objBrowseIndiceSegmentos = New AdmBrowseIndiceSegmentos
    
        objBrowseIndiceSegmentos.sNomeTela = sNomeTela
        objBrowseIndiceSegmentos.iIndice = iIndice
        objBrowseIndiceSegmentos.iPosicaoCampo = iPosicaoCampo
        iPosicaoFimCampo = InStr(iPosicao, objBrowseIndice.sOrdenacaoSQL, ",")
        If iPosicaoFimCampo <> 0 Then
            objBrowseIndiceSegmentos.sNomeCampo = Trim(Mid(objBrowseIndice.sOrdenacaoSQL, iPosicao, iPosicaoFimCampo - iPosicao))
            iPosicaoFimCampo = iPosicaoFimCampo + 1
            iPosicao = iPosicaoFimCampo
        Else
            objBrowseIndiceSegmentos.sNomeCampo = Trim(Mid(objBrowseIndice.sOrdenacaoSQL, iPosicao))
        End If

        colBrowseIndiceSegmentos.Add objBrowseIndiceSegmentos
    
    Loop
    
    BrowseIndiceSegmentos_Le = SUCESSO
    
    Exit Function
    
Erro_BrowseIndiceSegmentos_Le:

    BrowseIndiceSegmentos_Le = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142515)
        
    End Select
    
    Exit Function

End Function

Function Browse_Le_NumRegistros1(ByVal sSQL As String, lNumReg As Long, ByVal colSegmentos As Collection, ByVal colSelecao As Collection, ByVal objBrowse As AdmBrowse) As Long
'descobre o número de registros contidos no comandoSQL montado a partir dos parametros
'colSegmentos contém os valores atuais da seleção

Dim lErro As Long
Dim vSegmento As Variant
Dim lComando As Long
Dim iIndice As Integer
Dim avCampos(1 To BROWSER_NUM_MAX_CAMPOS) As Variant
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim lTeste As Long
Dim vNumReg As Variant
Dim vSelecao As Variant


On Error GoTo Erro_Browse_Le_NumRegistros1

    If objBrowse.iBancoDados = BD_DICDADOS Then

        lComando = Comando_AbrirExt(GL_lConexaoDicBrowse)
        If lComando = 0 Then Error 20707

    Else

        lComando = Comando_AbrirExt(GL_lConexaoBrowse)
        If lComando = 0 Then Error 9148

    End If
    
    lErro = Comando_PrepararInt1(lComando, "SELECT COUNT(*) FROM " & sSQL)
    If lErro <> AD_SQL_SUCESSO Then Error 9149
    
    vNumReg = lNumReg
    
    lErro = Comando_BindVarInt1(lComando, vNumReg)
    If lErro <> AD_SQL_SUCESSO Then Error 20729
    
    iIndice2 = 0
    
    If Not (colSelecao Is Nothing) Then
    
        For Each vSelecao In colSelecao
        
            iIndice2 = iIndice2 + 1
            
            avCampos(iIndice2) = vSelecao
            
            lErro = Comando_BindVarInt1(lComando, avCampos(iIndice2))
            If lErro <> AD_SQL_SUCESSO Then Error 9253
            
        Next
        
    End If
    
    For iIndice1 = 1 To colSegmentos.Count
    
        For iIndice = 1 To iIndice1
                             
            iIndice2 = iIndice2 + 1
                     
            avCampos(iIndice2) = colSegmentos.Item(iIndice)
                 
            lErro = Comando_BindVarInt1(lComando, avCampos(iIndice2))
            If lErro <> AD_SQL_SUCESSO Then Error 9150
            
        
        Next
        
    Next
    
    
    lErro = Comando_ExecutarInt1(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9151
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9165
    
    lNumReg = vNumReg
   
    'mario 1/4/99. Colocado para não deixar que a scrollbar esconda um registro.
'    lNumReg = lNumReg + 1
   
    Call Comando_Fechar(lComando)
    
    Browse_Le_NumRegistros1 = SUCESSO
    
    Exit Function

Erro_Browse_Le_NumRegistros1:

    Browse_Le_NumRegistros1 = Err
    
    Select Case Err
    
        Case 9148, 20707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", Err, "SELECT Count(*) FROM " & sSQL)
        
        Case 9150, 9253, 20729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BIND_CAMPO1", Err)
            
        Case 9151, 9165
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", Err, "SELECT Count(*) FROM " & sSQL)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142516)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
        
End Function

Function Browse_Le_Posicao(ByVal lPosicao As Long, ByVal sSQL As String, ByVal colBrowseValorCampo As Collection, ByVal colSelecao As Collection, ByVal objBrowse As AdmBrowse) As Long
'le o registro referente a posicao passada como parametro

Dim lErro As Long
Dim objBrowseValorCampo As AdmBrowseValorCampo
Dim iIndice As Integer
Dim lComando As Long
Dim avCampo(1 To BROWSER_NUM_MAX_CAMPOS) As Variant
Dim vSelecao As Variant

On Error GoTo Erro_Browse_Le_Posicao


    If objBrowse.iBancoDados = BD_DICDADOS Then

        lComando = Comando_AbrirExt(GL_lConexaoDicBrowse)
        If lComando = 0 Then Error 20708

    Else
    
        lComando = Comando_AbrirExt(GL_lConexaoBrowse)
        If lComando = 0 Then Error 9155

    End If

    lErro = Comando_PrepararInt1(lComando, sSQL)
    If lErro <> AD_SQL_SUCESSO Then Error 9156
    
    iIndice = 0
    
    For Each objBrowseValorCampo In colBrowseValorCampo
            
        iIndice = iIndice + 1
        
        avCampo(iIndice) = objBrowseValorCampo.vValorCampo
            
        lErro = Comando_BindVarInt1(lComando, avCampo(iIndice))
        If lErro <> AD_SQL_SUCESSO Then Error 9157
        
    Next
    
    If Not (colSelecao Is Nothing) Then
    
        For Each vSelecao In colSelecao
        
            iIndice = iIndice + 1
            
            avCampo(iIndice) = vSelecao
            
            lErro = Comando_BindVarInt1(lComando, avCampo(iIndice))
            If lErro <> AD_SQL_SUCESSO Then Error 9256
            
        Next
        
    End If
    
    'PARA USAR A ROTINA ANTIGA BASTA INSERIR UM REGISTRO NA TABELA CONTROLE COM Codigo=2 e Conteudo=1
    If giBrowsePosicaoAntigo = 0 Then
    
    
        'adicionado para poder selecionar o registro que esta sendo editado
        If Not (objBrowse.colReg Is Nothing) Then
        
            If objBrowse.colReg.Count >= (objBrowse.objGrid.Row - objBrowse.objGrid.TopRow + 1) Then
        
                For Each vSelecao In objBrowse.colReg.Item(objBrowse.objGrid.Row - objBrowse.objGrid.TopRow + 1)
                
                    iIndice = iIndice + 1
                    
                    If VarType(vSelecao) = vbDouble Then
                        If Abs(vSelecao) < DELTA_VALORMONETARIO2 Then vSelecao = 0
                    End If
                    
                    avCampo(iIndice) = vSelecao
                    
                    lErro = Comando_BindVarInt1(lComando, avCampo(iIndice))
                    If lErro <> AD_SQL_SUCESSO Then Error 9256
                    
                Next
            
            End If
            
        End If
    
        lErro = Comando_ExecutarInt1(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 9158
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9159
    
    
    Else
    
        lErro = Comando_ExecutarInt1(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 9158
    
        lErro = Comando_BuscarAbsoluto(lComando, lPosicao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9159
    
    End If
    
    
    'mario 1/4/99. Colocado para não selecionar quando não tiver encontrado dados.
    If lErro = AD_SQL_SEM_DADOS Then Error 20710

    iIndice = 0
    
    For Each objBrowseValorCampo In colBrowseValorCampo
    
        iIndice = iIndice + 1
        
        objBrowseValorCampo.vValorCampo = avCampo(iIndice)
        
    Next
    
    Call Comando_Fechar(lComando)
    
    Browse_Le_Posicao = SUCESSO
    
    Exit Function

Erro_Browse_Le_Posicao:

    Browse_Le_Posicao = Err
    
    Select Case Err
    
        Case 9155, 20708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9156
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", Err, sSQL)
        
        Case 9157, 9256
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BIND_CAMPO1", Err)
            
        Case 9158, 9159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", Err, sSQL)
            
        Case 20710
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142517)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
        
End Function

Function BrowseArquivo_Le(ByVal sNomeTela As String, ByVal objBrowseArquivo As AdmBrowseArquivo) As Long
'le todos os campos da tela de browse para o usuario especificado e coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim tBrowseArquivo As typeBrowseArquivo
    
On Error GoTo Erro_BrowseArquivo_Le

    tBrowseArquivo.sClasse = String(NOME_CLASSE, 0)
    tBrowseArquivo.sProjeto = String(NOME_PROJETO, 0)
    tBrowseArquivo.sNomeArq = String(STRING_NOME_TABELA, 0)
    tBrowseArquivo.sSelecaoSQL = String(STRING_SELECAO_SQL, 0)
    tBrowseArquivo.sClasseBrowser = String(STRING_NOME_CLASSEBROWSER, 0)
    tBrowseArquivo.sTrataParametros = String(STRING_NOME_TRATAPARAMETROS, 0)
    tBrowseArquivo.sTituloBrowser = String(STRING_TITULOBROWSER, 0)
    tBrowseArquivo.sRotinaBotaoEdita = String(STRING_NOME_ROTINABOTAOEDITA, 0)
    tBrowseArquivo.sRotinaBotaoSeleciona = String(STRING_NOME_ROTINABOTAOSELECIONA, 0)
    tBrowseArquivo.sRotinaBotaoConsulta = String(STRING_NOME_ROTINABOTAOCONSULTA, 0)
    tBrowseArquivo.sClasseObjeto = String(NOME_CLASSE, 0)
    tBrowseArquivo.sProjetoObjeto = String(NOME_PROJETO, 0)
    tBrowseArquivo.sNomeTelaConsulta = String(STRING_NOME_TELA, 0)
    tBrowseArquivo.sNomeTelaEdita = String(STRING_NOME_TELA, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 9252

    lErro = Comando_Executar(lComando, "SELECT Projeto, Classe, NomeArq, SelecaoSQL, ClasseBrowser, TrataParametros, TituloBrowser, RotinaBotaoEdita, RotinaBotaoSeleciona, RotinaBotaoConsulta, BotaoSeleciona, BotaoEdita, BotaoConsulta, ProjetoObjeto, ClasseObjeto, NomeTelaConsulta, NomeTelaEdita, BancoDados FROM BrowseArquivo WHERE NomeTela=?", tBrowseArquivo.sProjeto, tBrowseArquivo.sClasse, tBrowseArquivo.sNomeArq, tBrowseArquivo.sSelecaoSQL, tBrowseArquivo.sClasseBrowser, tBrowseArquivo.sTrataParametros, tBrowseArquivo.sTituloBrowser, tBrowseArquivo.sRotinaBotaoEdita, tBrowseArquivo.sRotinaBotaoSeleciona, tBrowseArquivo.sRotinaBotaoConsulta, tBrowseArquivo.iBotaoSeleciona, tBrowseArquivo.iBotaoEdita, tBrowseArquivo.iBotaoConsulta, tBrowseArquivo.sProjetoObjeto, tBrowseArquivo.sClasseObjeto, tBrowseArquivo.sNomeTelaConsulta, tBrowseArquivo.sNomeTelaEdita, tBrowseArquivo.iBancoDados, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then Error 9253
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9254
    
    objBrowseArquivo.sNomeTela = sNomeTela
    objBrowseArquivo.sClasse = tBrowseArquivo.sClasse
    objBrowseArquivo.sProjeto = tBrowseArquivo.sProjeto
    objBrowseArquivo.sNomeArq = tBrowseArquivo.sNomeArq
    objBrowseArquivo.sSelecaoSQL = tBrowseArquivo.sSelecaoSQL
    objBrowseArquivo.sClasseBrowser = tBrowseArquivo.sClasseBrowser
    objBrowseArquivo.sTrataParametros = tBrowseArquivo.sTrataParametros
    objBrowseArquivo.sTituloBrowser = tBrowseArquivo.sTituloBrowser
    objBrowseArquivo.sRotinaBotaoEdita = tBrowseArquivo.sRotinaBotaoEdita
    objBrowseArquivo.sRotinaBotaoSeleciona = tBrowseArquivo.sRotinaBotaoSeleciona
    objBrowseArquivo.sRotinaBotaoConsulta = tBrowseArquivo.sRotinaBotaoConsulta
    objBrowseArquivo.iBotaoSeleciona = tBrowseArquivo.iBotaoSeleciona
    objBrowseArquivo.iBotaoEdita = tBrowseArquivo.iBotaoEdita
    objBrowseArquivo.iBotaoConsulta = tBrowseArquivo.iBotaoConsulta
    objBrowseArquivo.sProjetoObjeto = tBrowseArquivo.sProjetoObjeto
    objBrowseArquivo.sClasseObjeto = tBrowseArquivo.sClasseObjeto
    objBrowseArquivo.sNomeTelaConsulta = tBrowseArquivo.sNomeTelaConsulta
    objBrowseArquivo.sNomeTelaEdita = tBrowseArquivo.sNomeTelaEdita
    objBrowseArquivo.iBancoDados = tBrowseArquivo.iBancoDados
    
    Call Comando_Fechar(lComando)
    
    BrowseArquivo_Le = SUCESSO
    
    Exit Function
    
Erro_BrowseArquivo_Le:

    BrowseArquivo_Le = Err

    Select Case Err
    
        Case 9252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 9253, 9254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEARQUIVO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142518)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BrowseCampo_Le(ByVal sNomeTela As String, ByVal colBrowseCampo As Collection) As Long
'le todos os campos que podem ser utilizado na tela de browse e coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim tBrowseCampo As typeBrowseCampo
Dim objBrowseCampo As AdmBrowseCampo
    
On Error GoTo Erro_BrowseCampo_Le

    tBrowseCampo.sNomeTela = String(STRING_NOME_TELA, 0)
    tBrowseCampo.sNomeCampo = String(STRING_NOME_CAMPO, 0)
    tBrowseCampo.sNome = String(STRING_NOME_CAMPO, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 89970

    lErro = Comando_Executar(lComando, "SELECT NomeCampo, Nome FROM BrowseCampo WHERE NomeTela=?", tBrowseCampo.sNomeCampo, tBrowseCampo.sNome, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 89971
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89972
    
    Do While lErro = SUCESSO
    
        Set objBrowseCampo = New AdmBrowseCampo
    
        objBrowseCampo.sNomeTela = sNomeTela
        objBrowseCampo.sNomeCampo = tBrowseCampo.sNomeCampo
        objBrowseCampo.sNome = tBrowseCampo.sNome
        
        colBrowseCampo.Add objBrowseCampo
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89973
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    BrowseCampo_Le = SUCESSO
    
    Exit Function
    
Erro_BrowseCampo_Le:

    BrowseCampo_Le = gErr

    Select Case gErr
    
        Case 89970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89971, 89972, 89973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSECAMPO", gErr, sNomeTela)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142519)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Relatorio_Le(ByVal objRelatorio As AdmRelatorio) As Long
'retorna os dados do relatorio referentes ao Codigo ja passado por parametro em objRelatorio

Dim lErro As Long, iTipoDeAcesso As Integer
Dim iFim_de_Arquivo As Integer
Dim lComando As Long
Dim sNomeTsk As String, sNomeTskCustomizado As String, sNomeTskCustomizadoGrupo As String
Dim sAutor As String
Dim sDescricao As String
Dim sTelaAuxiliar As String
Dim sUltOpcaoUtilizada As String, iOrigem As Integer
Dim iTipo As Integer, iLandscape As Integer, iPrevia As Integer
Dim iDispositivoDeSaida As Integer
Dim iOutOfProcess As Integer

On Error GoTo Erro_Relatorio_Le
    
    'inicializa as variaveis referentes as colunas da tabela Relatorios no BD
    sNomeTsk = String(STRING_NOME_ARQ_COMPLETO, 0)
    sNomeTskCustomizado = String(STRING_NOME_ARQ_COMPLETO, 0)
    sNomeTskCustomizadoGrupo = String(STRING_NOME_ARQ_COMPLETO, 0)
    sAutor = String(STRING_CODIGO_RELATORIO, 0)
    sDescricao = String(STRING_DESC_RELATORIO, 0)
    sTelaAuxiliar = String(STRING_NOME_TELA, 0)
    sUltOpcaoUtilizada = String(STRING_CODIGO_OPCAO_RELATORIO, 0)
          
    'abre conexao com o BD
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 12007
    
    'faz a selecao dos dados na tabela Relatorios
    lErro = Comando_Executar(lComando, "SELECT TipoDeAcesso, NomeTsk, Relatorios.NomeTskCustomizado, GrupoRelatorios.NomeTskCustomizado, Autor, Relatorios.Descricao, TelaObterParam, UltOpcaoUtilizada, Tipo, DispositivoDeSaida, Landscape, Previa, Origem, OutOfProcess FROM Relatorios, GrupoRelatorios, GruposDeUsuarios, Usuarios WHERE Relatorios.CodRel = ? AND CodUsuario = ? AND Relatorios.CodRel = GrupoRelatorios.CodRel AND GrupoRelatorios.CodGrupo = GruposDeUsuarios.CodGrupo AND GruposDeUsuarios.CodGrupo = Usuarios.CodGrupo", iTipoDeAcesso, sNomeTsk, sNomeTskCustomizado, sNomeTskCustomizadoGrupo, sAutor, sDescricao, sTelaAuxiliar, sUltOpcaoUtilizada, iTipo, iDispositivoDeSaida, iLandscape, iPrevia, iOrigem, iOutOfProcess, objRelatorio.sCodRel, gsUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 12008
    
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    
    If iFim_de_Arquivo <> AD_SQL_SUCESSO Then Error 12009
    
    If iTipoDeAcesso = 0 Then Error 12010
           
    With objRelatorio
        If Len(Trim(sNomeTskCustomizadoGrupo)) <> 0 Then
            .sNomeTsk = sNomeTskCustomizadoGrupo
        Else
            If Len(Trim(sNomeTskCustomizado)) <> 0 Then
                .sNomeTsk = sNomeTskCustomizado
            Else
                .sNomeTsk = sNomeTsk
            End If
        End If
        .sAutor = sAutor
        .sDescricao = sDescricao
        .sTelaAuxiliar = sTelaAuxiliar
        .sUltOpcaoUtilizada = sUltOpcaoUtilizada
        .iTipo = iTipo
        .iDispositivoDeSaida = iDispositivoDeSaida
        .iLandscape = iLandscape
        .iPrevia = iPrevia
        .iOrigem = iOrigem
        .iOutOfProcess = iOutOfProcess
    End With
    
    Call Comando_Fechar(lComando)
    
    Relatorio_Le = SUCESSO
    
    Exit Function

Erro_Relatorio_Le:

    Relatorio_Le = Err
    
    Select Case Err
                
        Case 12007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 12008, 12009
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
            
        Case 12010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_SEM_PERMISSAO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142520)
        
    End Select
        
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Function Relatorios_Le_GrupoModulo(ByVal colRelatorio As Collection, ByVal sGrupo As String, sModulo As String) As Long
'retorna os relatorios referentes ao grupo e modulo que o usuario tem acesso
'os codigos dos relatorios sao retornados em colRelatorio e o grupo e modulo sao passados como parametros

Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim lComando As Long
Dim sCodRel As String

On Error GoTo Erro_Relatorios_Le_GrupoModulo
       
    sCodRel = String(STRING_CODIGO_RELATORIO, 0)
    
    lComando = 0

    'abre conexao com o BD
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 12003
    
    'seleciona os relatorios que o usuario do grupo tem acesso para o modulo corrente
    lErro = Comando_Executar(lComando, "SELECT Relatorios.CodRel FROM Relatorios , GrupoRelatorios , RelatoriosModulo WHERE Relatorios.CodRel = GrupoRelatorios.CodRel AND Relatorios.CodRel = RelatoriosModulo.CodRel AND GrupoRelatorios.TipoDeAcesso = 1 AND GrupoRelatorios.CodGrupo = ? AND RelatoriosModulo.Modulo = ?", sCodRel, sGrupo, sModulo)
    If lErro <> AD_SQL_SUCESSO Then Error 12004
    
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
        
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO
                
        'Adiciona à Coleção
        colRelatorio.Add (sCodRel)

        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    'Se o erro não foi fim de arquivo desvia para rotina de erro
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 12005
        
    Call Comando_Fechar(lComando)
    
    Relatorios_Le_GrupoModulo = SUCESSO
    
    Exit Function

Erro_Relatorios_Le_GrupoModulo:

    Relatorios_Le_GrupoModulo = Err
    
    Select Case Err
                
        Case 12003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 12004, 12005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
                
        Case 12006
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142521)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142522)
        
    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Function RelOpcoes_Le(ByVal objRelOpcoes As AdmRelOpcoes) As Long
'le uma opção de relatório contida na tabela RelatorioOpcoes
'objRelOpcoes.sCodRel e objRelOpcoes.sNome tem que vir preenchidos

Dim lComando As Long
Dim lErro As Long
Dim sArqParametros As String
    
On Error GoTo Erro_RelOpcoes_Le
    
    sArqParametros = String(STRING_NOME_ARQ_COMPLETO, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 13000

    'le a tabela de opções de relatório
    lErro = Comando_Executar(lComando, "SELECT ArqParametros FROM RelatorioOpcoes WHERE CodRel = ? AND Nome = ?", sArqParametros, objRelOpcoes.sCodRel, objRelOpcoes.sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 13001

    'lê a primeira opção de relatório
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 13002
    
    objRelOpcoes.sArqParam = sArqParametros
    
    Call Comando_Fechar(lComando)
    
    RelOpcoes_Le = SUCESSO

    Exit Function

Erro_RelOpcoes_Le:

    RelOpcoes_Le = Err

    Select Case Err

        Case 13000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 13001, 13002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIOOPCOES", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142523)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RelOpcoes_Le_Todos(ByVal sCodRel As String, ByVal colRelParametros As Collection) As Long
'le todos as opções de relatório contidas na tabela RelatorioOpcoes e coloca-os em colRelParametros

Dim lComando As Long
Dim lErro As Long
Dim objRelOpcoes As AdmRelOpcoes
Dim sNome As String, sArqParametros As String
    
On Error GoTo Erro_RelOpcoes_Le_Todos
    
    sNome = String(STRING_CODIGO_OPCAO_RELATORIO, 0)
    sArqParametros = String(STRING_NOME_ARQ_COMPLETO, 0)
    
    lComando = 0

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 13026

    'le a tabela de opções de relatório
    lErro = Comando_Executar(lComando, "SELECT Nome, ArqParametros FROM RelatorioOpcoes WHERE CodRel = ?", sNome, sArqParametros, sCodRel)
    If lErro <> AD_SQL_SUCESSO Then Error 13027

    'le a primeira opção de relatório
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 13028
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objRelOpcoes = New AdmRelOpcoes
        
        objRelOpcoes.sCodRel = sCodRel
        objRelOpcoes.sNome = sNome
        objRelOpcoes.sArqParam = sArqParametros
        
        'coloca a opção do relatório lida na coleção
        colRelParametros.Add objRelOpcoes
        
        'le a proxima opção de relatório
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 13029

    Loop

    Call Comando_Fechar(lComando)
    
    RelOpcoes_Le_Todos = SUCESSO

    Exit Function

Erro_RelOpcoes_Le_Todos:

    RelOpcoes_Le_Todos = Err

    Select Case Err

        Case 13026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 13027, 13028, 13029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIOOPCOES", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142524)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Codigos_Le(ByVal sTabela As String, ByVal sCampo_Codigo As String, ByVal iCodigo_Tipo As enumTipoCodigo, ByVal colCodigo As Collection, Optional ByVal iTamanho_Nome As Integer) As Long
'Le todos os codigos do campo sCampo_Codigo de sTabela , de acordo com o tipo (iCodigo_Tipo) e devolve na colecao colCodigo.
'Se iCodigoTipo for TIPO_STR tem que vir preenchido iTamanho_Nome

Dim lComando As Long
Dim lErro As Long
Dim iCodigo As Integer
Dim lCodigo As Long
Dim sCodigo As String

On Error GoTo Erro_Codigos_Le

    lComando = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12284

    Select Case iCodigo_Tipo

        Case TIPO_INT

            lErro = Comando_Executar(lComando, "SELECT " & sCampo_Codigo & " FROM " & sTabela & " ORDER BY " & sCampo_Codigo, iCodigo)
            If lErro <> AD_SQL_SUCESSO Then Error 12285

        Case TIPO_LONG

            lErro = Comando_Executar(lComando, "SELECT " & sCampo_Codigo & " FROM " & sTabela & " ORDER BY " & sCampo_Codigo, lCodigo)
            If lErro <> AD_SQL_SUCESSO Then Error 12286

        Case TIPO_STR

            sCodigo = String(iTamanho_Nome, 0)

            lErro = Comando_Executar(lComando, "SELECT " & sCampo_Codigo & " FROM " & sTabela & " ORDER BY " & sCampo_Codigo, sCodigo)
            If lErro <> AD_SQL_SUCESSO Then Error 12287

    End Select

    'le o primeiro codigo de sTabela
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12288

    Do While lErro <> AD_SQL_SEM_DADOS

        Select Case iCodigo_Tipo

            Case TIPO_INT

                colCodigo.Add iCodigo

            Case TIPO_LONG

                colCodigo.Add lCodigo

            Case TIPO_STR

                colCodigo.Add sCodigo

        End Select

        'le o proximo registro da tabela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12289

    Loop

    Call Comando_Fechar(lComando)

    Codigos_Le = SUCESSO

    Exit Function

Erro_Codigos_Le:

    Codigos_Le = Err

    Select Case Err

        Case 12284
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12285, 12286, 12287
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabela)

        Case 12288, 12289
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabela)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142525)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cod_Nomes_Le(ByVal sTabela As String, ByVal sCampo_Codigo As String, ByVal sCampo_Nome As String, ByVal iTamanho_Nome As Integer, ByVal colCodigoDescricao As AdmColCodigoNome, Optional ByVal sFiltro As String = "") As Long
'Le todos os campos sCampo_Codigo e sCampo_Nome (de tamanho = iTamanho_Nome) da tabela sTabela e coloca na colecao
'O sCampo_Codigo deve ter tipo Inteiro.

Dim lComando As Long
Dim lErro As Long
Dim iCodigo As Integer
Dim sNome As String

On Error GoTo Erro_Cod_Nomes_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12278

    sNome = String(IIf(iTamanho_Nome < 255, 255, iTamanho_Nome), 0)

    If sFiltro = "" Then
        lErro = Comando_Executar(lComando, "SELECT " & sCampo_Codigo & ", " & sCampo_Nome & " FROM " & sTabela & " ORDER BY " & sCampo_Codigo, iCodigo, sNome)
    Else
        lErro = Comando_Executar(lComando, "SELECT " & sCampo_Codigo & ", " & sCampo_Nome & " FROM " & sTabela & " WHERE " & sFiltro & " ORDER BY " & sCampo_Codigo, iCodigo, sNome)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 12279

    'le o primeiro sCampo_Codigo e sCampo_Nome de sTabela
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12280

    Do While lErro <> AD_SQL_SEM_DADOS

        'coloca o objCodigoDescricao lido na coleção
        colCodigoDescricao.Add iCodigo, sNome

        'le o proximo registro da tabela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12281

    Loop

    Call Comando_Fechar(lComando)

    Cod_Nomes_Le = SUCESSO

    Exit Function

Erro_Cod_Nomes_Le:

    Cod_Nomes_Le = Err

    Select Case Err

        Case 12278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12279, 12280, 12281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabela)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142526)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function LCod_Nomes_Le(ByVal sTabela As String, ByVal sCampo_Codigo As String, ByVal sCampo_Nome As String, ByVal iTamanho_Nome As Integer, ByVal colCodigoNome As AdmCollCodigoNome) As Long
'Le todos os campos sCampo_Codigo e sCampo_Nome (de tamanho = iTamanho_Nome) da tabela sTabela e coloca na colecao
'O sCampo_Codigo tem que ser do tipo Long.

Dim lComando As Long
Dim lErro As Long
Dim lCodigo As Long
Dim sNome As String

On Error GoTo Erro_LCod_Nomes_Le

    lComando = 0

    sNome = String(iTamanho_Nome, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12531

    lErro = Comando_Executar(lComando, "SELECT " & sCampo_Codigo & ", " & sCampo_Nome & " FROM " & sTabela & " ORDER BY " & sCampo_Codigo, lCodigo, sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 12532

    'le o primeiro sCampo_Codigo e sCampo_Nome de sTabela
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12533

    Do While lErro <> AD_SQL_SEM_DADOS

        'coloca o objCodigoDescricao lido na coleção
        colCodigoNome.Add lCodigo, sNome

        'le o proximo registro da tabela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12534

    Loop

    Call Comando_Fechar(lComando)

    LCod_Nomes_Le = SUCESSO

    Exit Function

Erro_LCod_Nomes_Le:

    LCod_Nomes_Le = Err

    Select Case Err

        Case 12531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12532, 12533, 12534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabela)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142527)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Modulos_Le_Todos(ByVal colModulos As Collection) As Long
'Le todos os modulos contidos na tabela Modulos no dicionario de dados
'ordernados por Versao, Sigla.

Dim lComando As Long
Dim lErro As Long
Dim objModulo As AdmModulo
Dim tModulo As typeModulo
Dim lConexao As Long

On Error GoTo Erro_Modulos_Le_Todos

    lConexao = GL_lConexaoDic

    tModulo.sSigla = String(STRING_MODULO_SIGLA, 0)
    tModulo.sNome = String(STRING_MODULO_NOME, 0)
    tModulo.sDescricao = String(STRING_MODULO_DESCRICAO, 0)
    tModulo.sVersao = String(STRING_MODULO_VERSAO, 0)

    'abre comando
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 62334
       
    'Lê todos os registros da tabela Modulos em DIC ordenados por versão,sigla
    lErro = Comando_Executar(lComando, "SELECT Sigla, Nome, Descricao, Versao FROM Modulos, Versao WHERE Modulos.Versao = Versao.Codigo ORDER BY Versao.Data, Modulos.Sigla ", tModulo.sSigla, tModulo.sNome, tModulo.sDescricao, tModulo.sVersao)
    If lErro <> AD_SQL_SUCESSO Then Error 62335

    'le o primeiro modulo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62336

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objModulo = New AdmModulo

        objModulo.sSigla = tModulo.sSigla
        objModulo.sNome = tModulo.sNome
        objModulo.sDescricao = tModulo.sDescricao
        objModulo.sVersao = tModulo.sVersao
        
        'coloca o modulo lido na coleção
        colModulos.Add objModulo

        'le o proximo modulo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62337

    Loop

    Call Comando_Fechar(lComando)

    Modulos_Le_Todos = SUCESSO

    Exit Function

Erro_Modulos_Le_Todos:

    Modulos_Le_Todos = Err

    Select Case Err

        Case 62334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 62335, 62336, 62337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142528)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FiliaisEmpresas_Le_Empresa(ByVal lCodEmpresa As Long, ByVal colFiliais As Collection) As Long
'Lê código e nome de todas as filiais configuradas da empresa em questão e coloca-as em colFiliais

Dim lComando As Long
Dim lErro As Long
Dim objFiliais As AdmFiliais
Dim tFiliais As typeFiliais
Dim sSiglaModulo As String
Dim iCodFilial As Integer

On Error GoTo Erro_FiliaisEmpresas_Le_Empresa

    lComando = 0

    tFiliais.sNome = String(STRING_FILIAL_NOME, 0)
    sSiglaModulo = String(STRING_MODULO_SIGLA, 0)

    'abre conexao com o BD
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 10197
    
    'le todas as filiais da empresa atual
    lErro = Comando_Executar(lComando, "SELECT FiliaisEmpresas.Inativa, FiliaisEmpresas.CodFilial, FiliaisEmpresas.Nome, ModuloFilEmp.SiglaModulo FROM FiliaisEmpresas, ModuloFilEmp WHERE FiliaisEmpresas.CodEmpresa = ? AND FiliaisEmpresas.CodEmpresa = ModuloFilEmp.CodEmpresa AND FiliaisEmpresas.CodFilial = ModuloFilEmp.CodFilial AND ModuloFilEmp.Configurado = ? ORDER BY FiliaisEmpresas.CodFilial", tFiliais.iInativa, tFiliais.iCodFilial, tFiliais.sNome, sSiglaModulo, lCodEmpresa, CONFIGURADO)
    If lErro <> AD_SQL_SUCESSO Then Error 10198
    
    'le a primeira filial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10199
    
    Do While lErro <> AD_SQL_SEM_DADOS

        If iCodFilial <> tFiliais.iCodFilial Then

            iCodFilial = tFiliais.iCodFilial

            Set objFiliais = New AdmFiliais
        
            objFiliais.iCodFilial = tFiliais.iCodFilial
            objFiliais.sNome = tFiliais.sNome
            objFiliais.iInativa = tFiliais.iInativa
        
            'coloca a filial lida na coleção
            colFiliais.Add objFiliais
        
        End If
        
        objFiliais.colModulos.Add sSiglaModulo
        
        'le a proxima filial
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10200

    Loop

    Set objFiliais = New AdmFiliais
    
    objFiliais.sNome = gsNomeEmpresa
    objFiliais.iCodFilial = EMPRESA_TODA

    'coloca a filial lida na coleção
    colFiliais.Add objFiliais

    Call Comando_Fechar(lComando)
    
    FiliaisEmpresas_Le_Empresa = SUCESSO

    Exit Function

Erro_FiliaisEmpresas_Le_Empresa:

    FiliaisEmpresas_Le_Empresa = Err
    
    Select Case Err

        Case 10197
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10198, 10199, 10200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAL", Err, glEmpresa)
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142529)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function FilialEmpresa_Le(ByVal objFilialEmpresa As AdmFiliais, Optional ByVal bDiretoNoBD As Boolean = False) As Long
'carrega dados da FilialEmpresa da tabela FiliaisEmpresa e de Enderecos
'Se objFilialEmpresa.iCodFilial = 0 entao pegar dados de giFilialEmpresa
'LE no SGE e no DIC
'??? atencao: está funcao tem implementacao atipica por ser destinada a minimizar I/O.

Dim lErro As Long, objFilialEmp As AdmFiliais, bAchou As Boolean
Dim iFilialEmpresa As Integer, lCodEmpresa As Long

On Error GoTo Erro_FilialEmpresa_Le

    If bDiretoNoBD Then
    
        'obs. p/retornar numeracao de erro da funcao FilialEmpresa_Le_BD
        FilialEmpresa_Le = FilialEmpresa_Le_BD(objFilialEmpresa)
        
        Exit Function
        
    Else
    
        'Se objFilialEmpresa.lCodEmpresa = 0 entao pegar dados de glEmpresa
        If objFilialEmpresa.lCodEmpresa = 0 Then
            lCodEmpresa = glEmpresa
        Else
            lCodEmpresa = objFilialEmpresa.lCodEmpresa
        End If
        
        'Se objFilialEmpresa.iCodFilial = 0 entao pegar dados de giFilialEmpresa
        If objFilialEmpresa.iCodFilial = 0 Then
            iFilialEmpresa = giFilialEmpresa
        Else
            iFilialEmpresa = objFilialEmpresa.iCodFilial
        End If
    
        bAchou = False
        
        If Not (gcolFiliaisEmpresa Is Nothing) Then
        
            For Each objFilialEmp In gcolFiliaisEmpresa
            
                If objFilialEmp.iCodFilial = iFilialEmpresa And objFilialEmp.lCodEmpresa = lCodEmpresa Then
                
                    bAchou = True
                    Exit For
                    
                End If
                
            Next
        
        End If
        
        If bAchou = False Then
        
            Set objFilialEmp = New AdmFiliais
            objFilialEmp.lCodEmpresa = lCodEmpresa
            objFilialEmp.iCodFilial = iFilialEmpresa
            
            lErro = FilialEmpresa_Le_BD(objFilialEmp)
            If lErro <> SUCESSO Then
            
                FilialEmpresa_Le = lErro
                Exit Function
                
            End If
            
            If gcolFiliaisEmpresa Is Nothing Then Set gcolFiliaisEmpresa = New Collection
            
            Call gcolFiliaisEmpresa.Add(objFilialEmp)
        
        End If
        
        Call objFilialEmpresa.Copiar(objFilialEmp)
        
        FilialEmpresa_Le = SUCESSO
     
    End If
    
    Exit Function
    
Erro_FilialEmpresa_Le:

    FilialEmpresa_Le = Err
     
    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142530)
     
    End Select
     
    Exit Function

End Function

Private Function FilialEmpresa_Le_BD(ByVal objFilialEmpresa As AdmFiliais) As Long
'carrega dados da FilialEmpresa da tabela FiliaisEmpresa e de Enderecos
'Se objFilialEmpresa.iCodFilial = 0 entao pegar dados de giFilialEmpresa
'LE no SGE e no DIC

Dim lErro As Long, lComando As Long, lComando1 As Long, tFilialEmpresa As typeFiliais
Dim iFilialEmpresa As Integer, objEndereco As New ClassEndereco, objEndEntrega As New ClassEndereco, objEndCTB As New ClassEndereco
Dim lCodEmpresa As Long

On Error GoTo Erro_FilialEmpresa_Le_BD

    'Abre comandos
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 25542
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27375
    
    'Se objFilialEmpresa.lCodEmpresa = 0 entao pegar dados de glEmpresa
    If objFilialEmpresa.lCodEmpresa = 0 Then
        lCodEmpresa = glEmpresa
    Else
        lCodEmpresa = objFilialEmpresa.lCodEmpresa
    End If
    
    'Se objFilialEmpresa.iCodFilial = 0 entao pegar dados de giFilialEmpresa
    If objFilialEmpresa.iCodFilial = 0 Then
        iFilialEmpresa = giFilialEmpresa
    Else
        iFilialEmpresa = objFilialEmpresa.iCodFilial
    End If
    
    With tFilialEmpresa
        .sCgc = String(STRING_CGC, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sISSCodigoPadrao = String(STRING_ISS_CODIGO, 0)
        .sNome = String(STRING_FILIAL_NOME, 0)
        .sRamoEmpresa = String(STRING_RAMO_EMPRESA, 0)
        .sJucerja = String(STRING_JUCERJA, 0)
        .sContador = String(STRING_CONTADOR, 0)
        .sCRCContador = String(STRING_CRC_CONTADOR, 0)
        .sCPFContador = String(STRING_CPF_CONTADOR, 0)
        .sCertificadoA1A3 = String(STRING_CERTIFICADOA1A3, 0)
        .sCNAE = String(STRING_CNAE_CODIGO, 0)
        .sSignatarioCTB = String(STRING_SIGNATARIO_CTB, 0)
        .sCodQualiSigCTB = String(STRING_CODQUALISIG_CTB, 0)
        .sCPFSignatarioCTB = String(STRING_CPF, 0)
        .sCNPJContador = String(STRING_CGC, 0)
        .sSpedFiscalPerfil = String(STRING_SPEDFISCAL_PERFIL, 0)
        .sNomeReduzido = String(STRING_FILIAL_NOME, 0)
                                
        .sNFCECSC = String(STRING_NFCE_CSC, 0)
        .sidNFCECSC = String(STRING_NFCE_ID_CSC, 0)
                                
        'Lê atributos no SGE
        lErro = Comando_Executar(lComando, "SELECT Nome, CGC, Endereco, EnderecoEntrega, AlmoxarifadoPadrao, InscricaoEstadual, ICMSPorEstimativa, InscricaoMunicipal, ISSPercPadrao, ISSCodigoPadrao, ISSIncluso, ContribuinteIPI, IRPercPadrao, RamoEmpresa, Jucerja, DataJucerja, Contador, CRCContador, CPFContador, TipoTribAtividade, SimplesFederal, SimplesFederalAliq, SimplesFederalTeto, SuperSimples, PISNaoCumulativo, COFINSNaoCumulativo, LucroPresumido, CertificadoA1A3, NFeAmbiente, CNAE, SignatarioCTB, CodQualiSigCTB, CPFSignatarioCTB, CNPJContador, SpedFiscalPerfil, NomeReduzido, EnderecoContador, RPSAmbiente, PisCofinsRegCumTipo, ContribPrevPorRecBruta, AliquotaPadraoContribPrev, PisCofinsIndAproCred, NFCECSC, idNFCECSC, indSincPadrao, RegimeTrib, RegimeEspecialTrib FROM FiliaisEmpresa WHERE FilialEmpresa = ?", _
            .sNome, .sCgc, .lEndereco, .lEndEntrega, .iAlmoxarifadoPadrao, .sInscricaoEstadual, .iICMSPorEstimativa, .sInscricaoMunicipal, .dISSPercPadrao, .sISSCodigoPadrao, .iISSIncluso, _
            .iContribuinteIPI, .dIRPercPadrao, .sRamoEmpresa, .sJucerja, .dtDataJucerja, .sContador, .sCRCContador, .sCPFContador, .iTipoTribAtividade, .iSimplesFederal, .dSimplesFederalAliq, .dSimplesFederalTeto, .iSuperSimples, .iPISNaoCumulativo, .iCOFINSNaoCumulativo, .iLucroPresumido, .sCertificadoA1A3, .iNFeAmbiente, .sCNAE, _
            .sSignatarioCTB, .sCodQualiSigCTB, .sCPFSignatarioCTB, .sCNPJContador, .sSpedFiscalPerfil, .sNomeReduzido, .lEnderecoContador, .iRPSAmbiente, .iPisCofinsRegCumTipo, .iContribPrevPorRecBruta, .dAliquotaPadraoContribPrev, .iPisCofinsIndAproCred, .sNFCECSC, .sidNFCECSC, .iindSincPadrao, .iRegimeTrib, .iRegimeEspecialTrib, iFilialEmpresa)
    End With
    If lErro <> AD_SQL_SUCESSO Then Error 27376
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27377
    
    If lErro <> AD_SQL_SUCESSO Then Error 27378
    
    'Lê endereços
    objEndereco.lCodigo = tFilialEmpresa.lEndereco
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO Then Error 27379
    
    If tFilialEmpresa.lEndEntrega <> 0 Then
    
        objEndEntrega.lCodigo = tFilialEmpresa.lEndEntrega
        lErro = CF("Endereco_Le", objEndEntrega)
        If lErro <> SUCESSO Then Error 25543
    
    End If
    
    If tFilialEmpresa.lEnderecoContador <> 0 Then
    
        objEndCTB.lCodigo = tFilialEmpresa.lEnderecoContador
        lErro = CF("Endereco_Le", objEndCTB)
        If lErro <> SUCESSO Then Error 25543
    
    End If
    
    'Lê atividade no DIC
    lErro = Comando_Executar(lComando1, "SELECT Inativa FROM FiliaisEmpresas WHERE CodEmpresa = ? AND CodFilial= ? ", _
        tFilialEmpresa.iInativa, lCodEmpresa, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 25544
    
    lErro = Comando_BuscarProximo(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25545
    If lErro <> AD_SQL_SUCESSO Then Error 27378 'Codigo igual de propósito
    
    If tFilialEmpresa.iInativa = 1 Then Error 27378 'Codigo igual de propósito
    
    'mover campos p/o objeto
    With objFilialEmpresa
        
        .dISSPercPadrao = tFilialEmpresa.dISSPercPadrao
        .iAlmoxarifadoPadrao = tFilialEmpresa.iAlmoxarifadoPadrao
        .iCodFilial = iFilialEmpresa
        .iContribuinteIPI = tFilialEmpresa.iContribuinteIPI
        .iICMSPorEstimativa = tFilialEmpresa.iICMSPorEstimativa
        .iISSIncluso = tFilialEmpresa.iISSIncluso
        .lCodEmpresa = lCodEmpresa
        Set .objEndereco = objEndereco
        Set .objEnderecoEntrega = objEndEntrega
        Set .objEnderecoContador = objEndCTB
        .sCgc = tFilialEmpresa.sCgc
        .sInscricaoEstadual = tFilialEmpresa.sInscricaoEstadual
        .sInscricaoMunicipal = tFilialEmpresa.sInscricaoMunicipal
        .sISSCodigoPadrao = tFilialEmpresa.sISSCodigoPadrao
        .sNome = tFilialEmpresa.sNome
        .dIRPercPadrao = tFilialEmpresa.dIRPercPadrao
        .sRamoEmpresa = tFilialEmpresa.sRamoEmpresa
        .sJucerja = tFilialEmpresa.sJucerja
        .dtDataJucerja = tFilialEmpresa.dtDataJucerja
        .sContador = tFilialEmpresa.sContador
        .sCRCContador = tFilialEmpresa.sCRCContador
        .sCPFContador = tFilialEmpresa.sCPFContador
        .iTipoTribAtividade = tFilialEmpresa.iTipoTribAtividade
        .iSimplesFederal = tFilialEmpresa.iSimplesFederal
        .dSimplesFederalAliq = tFilialEmpresa.dSimplesFederalAliq
        .dSimplesFederalTeto = tFilialEmpresa.dSimplesFederalTeto
        .iSuperSimples = tFilialEmpresa.iSuperSimples
        .iPISNaoCumulativo = tFilialEmpresa.iPISNaoCumulativo
        .iCOFINSNaoCumulativo = tFilialEmpresa.iCOFINSNaoCumulativo
        .iLucroPresumido = tFilialEmpresa.iLucroPresumido
        .sCertificadoA1A3 = tFilialEmpresa.sCertificadoA1A3
        .iNFeAmbiente = tFilialEmpresa.iNFeAmbiente
        .iRPSAmbiente = tFilialEmpresa.iRPSAmbiente
        .sCNAE = tFilialEmpresa.sCNAE
        .sSignatarioCTB = tFilialEmpresa.sSignatarioCTB
        .sCodQualiSigCTB = tFilialEmpresa.sCodQualiSigCTB
        .sCPFSignatarioCTB = tFilialEmpresa.sCPFSignatarioCTB
        .sCNPJContador = tFilialEmpresa.sCNPJContador
        .sSpedFiscalPerfil = tFilialEmpresa.sSpedFiscalPerfil
        .sNomeReduzido = tFilialEmpresa.sNomeReduzido
        .iPisCofinsRegCumTipo = tFilialEmpresa.iPisCofinsRegCumTipo
        .iPisCofinsIndAproCred = tFilialEmpresa.iPisCofinsIndAproCred

        .iContribPrevPorRecBruta = tFilialEmpresa.iContribPrevPorRecBruta
        .dAliquotaPadraoContribPrev = tFilialEmpresa.dAliquotaPadraoContribPrev
        
        .sNFCECSC = tFilialEmpresa.sNFCECSC
        .sidNFCECSC = tFilialEmpresa.sidNFCECSC
        .iindSincPadrao = tFilialEmpresa.iindSincPadrao
        .iRegimeTrib = tFilialEmpresa.iRegimeTrib
        .iRegimeEspecialTrib = tFilialEmpresa.iRegimeEspecialTrib
        
    End With
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    FilialEmpresa_Le_BD = SUCESSO

    Exit Function

Erro_FilialEmpresa_Le_BD:

    FilialEmpresa_Le_BD = Err

    Select Case Err

        Case 27379, 25543
        
        Case 25542, 27375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 27376, 27377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALEMPRESA", Err, iFilialEmpresa)
        
        Case 25544, 25545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALEMPRESA1", Err, lCodEmpresa, iFilialEmpresa)
        
        Case 27378 'filial nao encontrada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142531)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

'Function Usuario_Le(objUsuario As ClassUsuario) As Long
''le os dados do usuario passado como parametro e devolve os dados em objUsuario
''??? incompleta, nao le todos os dados
'Dim lErro As Long
'Dim lComando As Long
'Dim tUsuario As typeUsuario
'
'On Error GoTo Erro_Usuario_Le
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then Error 36344
'
'    lErro = Comando_Executar(lComando, "SELECT Lote FROM Usuario WHERE CodUsuario = ?", tUsuario.iLote, objUsuario.sCodUsuario)
'    If lErro <> AD_SQL_SUCESSO Then Error 36345
'
'    lErro = Comando_BuscarProximo(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36346
'
'    If lErro <> AD_SQL_SUCESSO Then Error 36347
'
'    objUsuario.iLote = tUsuario.iLote
'
'    Call Comando_Fechar(lComando)
'
'    Usuario_Le = SUCESSO
'
'    Exit Function
'
'Erro_Usuario_Le:
'
'    Usuario_Le = Err
'
'    Select Case Err
'
'        Case 36344
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 36345, 36346
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO", Err, objUsuario.sCodUsuario)
'
'        Case 36347 'Usuario nao encontrado
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142532)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function


Function Produto_Le2(ByVal lComando As Long, ByVal objProduto As ClassProduto) As Long
'Lê os dados do produto, cujo codigo foi passado como parâmetro

Dim lErro As Long
Dim tProduto As typeProduto

On Error GoTo Erro_Produto_Le2

    'Inicializa as variáveis do tipo String
    With tProduto
        
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)
        .sReferencia = String(STRING_PRODUTO_REFERENCIA, 0)
        .sFigura = String(STRING_NOME_ARQ_COMPLETO, 0)
        .sGrade = String(STRING_GRADE_CODIGO, 0)
                        
        'Lê o Produto
        lErro = Comando_Executar(lComando, "SELECT  Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza ,CustoReposicao, OrigemMercadoria, TabelaPreco, PercentMaisQuantCotAnt, PercentMenosQuantCotAnt, ConsideraQuantCotAnt, TemFaixaReceb, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa,TempoProducao,Rastro,HorasMaquina,PesoEspecifico, Embalagem, Referencia, Figura, INSSPercBase, KitVendaComp, Grade FROM Produtos WHERE Codigo = ?", _
        .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, .dCustoReposicao, .iOrigemMercadoria, .iTabelaPreco, .dPercentMaisQuantCotAnt, .dPercentMenosQuantCotAnt, .iConsideraQuantCotAnt, .iTemFaixaReceb, .dPercentMaisReceb, .dPercentMenosReceb, .iRecebForaFaixa, .iTempoProducao, .iRastro, .lHorasMaquina, .dPesoEspecifico, .iEmbalagem, .sReferencia, .sFigura, .dINSSPercBase, .iKitVendaComp, .sGrade, objProduto.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 20577

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20578

        'Se não encontrou o Produto --> Erro
        If lErro = AD_SQL_SEM_DADOS Then Error 20579

        'Coloca os dados lidos no objProduto
        objProduto.sDescricao = .sDescricao
        objProduto.iGerencial = .iGerencial
        objProduto.iAtivo = .iAtivo
        objProduto.iFaturamento = .iFaturamento
        objProduto.iClasseUM = .iClasseUM
        objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
        objProduto.sSiglaUMCompra = .sSiglaUMCompra
        objProduto.sSiglaUMVenda = .sSiglaUMVenda
        objProduto.sSubstituto1 = .sSubstituto1
        objProduto.sSubstituto2 = .sSubstituto2
        objProduto.iTipo = .iTipo
        objProduto.sNomeReduzido = .sNomeReduzido
        objProduto.sModelo = .sModelo
        objProduto.iNivel = .iNivel
        objProduto.iPrazoValidade = .iPrazoValidade
        objProduto.sCodigoBarras = .sCodigoBarras
        objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
        objProduto.dPesoLiq = .dPesoLiq
        objProduto.dPesoBruto = .dPesoBruto
        objProduto.dComprimento = .dComprimento
        objProduto.dEspessura = .dEspessura
        objProduto.dLargura = .dLargura
        objProduto.sCor = .sCor
        objProduto.sObsFisica = .sObsFisica
        objProduto.iCompras = .iCompras
        objProduto.iPCP = .iPCP
        objProduto.iKitBasico = .iKitBasico
        objProduto.iKitInt = .iKitInt
        objProduto.dIPIAliquota = .dIPIAliquota
        objProduto.sIPICodigo = .sIPICodigo
        objProduto.sIPICodDIPI = .sIPICodDIPI
        objProduto.iControleEstoque = .iControleEstoque
        objProduto.iApropriacaoCusto = .iApropriacaoCusto
        objProduto.sContaContabil = .sContaContabil
        objProduto.sContaContabilProducao = .sContaContabilProducao
        objProduto.dResiduo = .dResiduo
        objProduto.iNatureza = .iNatureza
        objProduto.dCustoReposicao = .dCustoReposicao
        objProduto.iOrigemMercadoria = .iOrigemMercadoria
        objProduto.iTabelaPreco = .iTabelaPreco
        objProduto.dPercentMaisQuantCotAnt = .dPercentMaisQuantCotAnt
        objProduto.dPercentMenosQuantCotAnt = .dPercentMenosQuantCotAnt
        objProduto.iConsideraQuantCotAnt = .iConsideraQuantCotAnt
        objProduto.iTemFaixaReceb = .iTemFaixaReceb
        objProduto.dPercentMaisReceb = .dPercentMaisReceb
        objProduto.dPercentMenosReceb = .dPercentMenosReceb
        objProduto.iRecebForaFaixa = .iRecebForaFaixa
        objProduto.iTempoProducao = .iTempoProducao
        objProduto.iRastro = .iRastro
        objProduto.lHorasMaquina = .lHorasMaquina
        objProduto.dPesoEspecifico = .dPesoEspecifico
        objProduto.iEmbalagem = .iEmbalagem
        objProduto.sReferencia = .sReferencia
        objProduto.sFigura = .sFigura
        objProduto.dINSSPercBase = .dINSSPercBase
        objProduto.iKitVendaComp = .iKitVendaComp
        objProduto.sGrade = .sGrade
        
    End With

    Produto_Le2 = SUCESSO

    Exit Function

Erro_Produto_Le2:

    Produto_Le2 = Err

    Select Case Err

        Case 20577, 20578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", Err, objProduto.sCodigo)

        Case 20579
            'Produto não encontrado. Será tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142533)

    End Select

    Exit Function

End Function

Function Usuarios_Le_Todos(ByVal colUsuarios As Collection) As Long
'le todos os usuários contidos na tabela de Usuarios e coloca os dados em colUsuarios

Dim lComando As Long
Dim lErro As Long
Dim tUsuarios As typeUsuarios
Dim objUsuarios As ClassUsuarios
    
On Error GoTo Erro_Usuarios_Le_Todos

    lComando = 0

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 36999

    tUsuarios.sCodUsuario = String(STRING_SIGLA_USUARIO, 0)
    tUsuarios.sNome = String(STRING_USUARIO_NOME, 0)
    tUsuarios.sNomeReduzido = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    tUsuarios.sCodGrupo = String(STRING_GRUPO, 0) 'Inserido por Wagner

    lErro = Comando_Executar(lComando, "SELECT CodUsuario, Nome, NomeReduzido, CodGrupo FROM Usuarios WHERE Ativo=?", tUsuarios.sCodUsuario, tUsuarios.sNome, tUsuarios.sNomeReduzido, tUsuarios.sCodGrupo, USUARIO_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then Error 44000
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44001
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objUsuarios = New ClassUsuarios
        
        objUsuarios.sCodUsuario = tUsuarios.sCodUsuario
        objUsuarios.sNome = tUsuarios.sNome
        objUsuarios.sNomeReduzido = tUsuarios.sNomeReduzido
        objUsuarios.sCodGrupo = tUsuarios.sCodGrupo 'Inserido por Wagner
        
        colUsuarios.Add objUsuarios
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44002
            
    Loop
            
    Call Comando_Fechar(lComando)
    
    Usuarios_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_Usuarios_Le_Todos:

    Usuarios_Le_Todos = Err

    Select Case Err
    
        Case 36999
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 44000, 44001, 44002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142534)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Niveis0e1(ByVal colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto
Dim objProduto As ClassProduto

On Error GoTo Erro_Produto_Le_Niveis0e1

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44126

    'Inicializa as variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)

        'Lê os Produto de Nivel 1 e 2
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza,TempoProducao,Rastro,HorasMaquina,PesoEspecifico FROM Produtos WHERE Nivel = 1 OR Nivel = 2 ORDER BY Codigo", _
            .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, .iTempoProducao, .iRastro, .lHorasMaquina, .dPesoEspecifico)
        If lErro <> AD_SQL_SUCESSO Then Error 44127

        'le o primeiro produto
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44128

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto
    
            'Coloca os dados lidos no objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
            objProduto.sContaContabilProducao = .sContaContabilProducao
            objProduto.dResiduo = .dResiduo
            objProduto.iNatureza = .iNatureza
            objProduto.iTempoProducao = .iTempoProducao
            objProduto.iRastro = .iRastro
            objProduto.lHorasMaquina = .lHorasMaquina
            objProduto.dPesoEspecifico = .dPesoEspecifico
            
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44129

        Loop

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Produto_Le_Niveis0e1 = SUCESSO

    Exit Function

Erro_Produto_Le_Niveis0e1:

    Produto_Le_Niveis0e1 = Err

    Select Case Err

        Case 44126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44127, 44128, 44129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142535)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Netos(ByVal sProduto As String, ByVal colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim lComando As Long
Dim lErro As Long
Dim objProduto As ClassProduto
Dim tProduto As typeProduto
Dim sProduto1 As String
Dim sProdutoEnxuto As String
Dim iNivel As Integer

On Error GoTo Erro_Produto_Le_Netos

    'Inicializa as Variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 44133

        lErro = Mascara_Produto_ObterNivel(sProduto, iNivel)
        If lErro <> SUCESSO Then Error 44134

        sProduto1 = sProdutoEnxuto + String(STRING_PRODUTO - Len(sProdutoEnxuto), "Z")

        lComando = Comando_Abrir()
        If lComando = 0 Then Error 44135
    
        'Lê produtos ordenados por Codigo (apenas os atributos que interessam)
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza,TempoProducao,Rastro,HorasMaquina,PesoEspecifico FROM Produtos WHERE Nivel = ? AND Codigo > ? AND Codigo < ? ORDER BY Codigo", .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, .iTempoProducao, .iRastro, .lHorasMaquina, .dPesoEspecifico, iNivel + 2, sProduto, sProduto1)
        If lErro <> AD_SQL_SUCESSO Then Error 44136

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44137

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto

            'Guarda so dados lidos em objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
            objProduto.sContaContabilProducao = .sContaContabilProducao
            objProduto.dResiduo = .dResiduo
            objProduto.iNatureza = .iNatureza
            objProduto.iTempoProducao = .iTempoProducao
            objProduto.iRastro = .iRastro
            objProduto.lHorasMaquina = .lHorasMaquina
            objProduto.dPesoEspecifico = .dPesoEspecifico
            
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44138

        Loop

    End With

    Call Comando_Fechar(lComando)
    
    Produto_Le_Netos = SUCESSO

    Exit Function

Erro_Produto_Le_Netos:

    Produto_Le_Netos = Err
    
    Select Case Err

        Case 44133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, sProduto)

        Case 44134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, sProduto)

        Case 44135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44136, 44137, 44138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142536)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function
    
Function Produto_Le_Niveis0e1_Reserva(ByVal colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto
Dim objProduto As ClassProduto

On Error GoTo Erro_Produto_Le_Niveis0e1_Reserva

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44146

    'Inicializa as variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)

        'Lê os Produto de Nivel 1 e 2
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza FROM ProdutosProdutosFilial WHERE ((FilialEmpresa = ? And ControleEstoque=?) OR Gerencial=?) AND Ativo=? AND (Nivel = 1 OR Nivel = 2) ORDER BY Codigo", _
            .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, giFilialEmpresa, PRODUTO_CONTROLE_RESERVA, GERENCIAL, Ativo)
        If lErro <> AD_SQL_SUCESSO Then Error 44147

        'le o primeiro produto
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44148

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto
    
            'Coloca os dados lidos no objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
            objProduto.sContaContabilProducao = .sContaContabilProducao
            objProduto.dResiduo = .dResiduo
            objProduto.iNatureza = .iNatureza
            
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44149

        Loop

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Produto_Le_Niveis0e1_Reserva = SUCESSO

    Exit Function

Erro_Produto_Le_Niveis0e1_Reserva:

    Produto_Le_Niveis0e1_Reserva = Err

    Select Case Err

        Case 44146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44147, 44148, 44149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142537)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Netos_Reserva(ByVal sProduto As String, ByVal colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim lComando As Long
Dim lErro As Long
Dim objProduto As ClassProduto
Dim tProduto As typeProduto
Dim sProduto1 As String
Dim sProdutoEnxuto As String
Dim iNivel As Integer

On Error GoTo Erro_Produto_Le_Netos_Reserva

    'Inicializa as Variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 44150

        lErro = Mascara_Produto_ObterNivel(sProduto, iNivel)
        If lErro <> SUCESSO Then Error 44151

        sProduto1 = sProdutoEnxuto + String(STRING_PRODUTO - Len(sProdutoEnxuto), "Z")

        lComando = Comando_Abrir()
        If lComando = 0 Then Error 44152
    
        'Lê produtos ordenados por Codigo (apenas os atributos que interessam)
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza FROM ProdutosProdutosFilial WHERE ((FilialEmpresa = ? AND ControleEstoque=?) OR Gerencial=?) AND Ativo=? AND Nivel = ? AND Codigo > ? AND Codigo < ? ORDER BY Codigo", _
                .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, giFilialEmpresa, PRODUTO_CONTROLE_RESERVA, GERENCIAL, Ativo, iNivel + 2, sProduto, sProduto1)
        If lErro <> AD_SQL_SUCESSO Then Error 44153

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44154

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto

            'Guarda so dados lidos em objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
            objProduto.sContaContabilProducao = .sContaContabilProducao
            objProduto.dResiduo = .dResiduo
            objProduto.iNatureza = .iNatureza
    
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44155

        Loop

    End With

    Call Comando_Fechar(lComando)
    
    Produto_Le_Netos_Reserva = SUCESSO

    Exit Function

Erro_Produto_Le_Netos_Reserva:

    Produto_Le_Netos_Reserva = Err
    
    Select Case Err

        Case 44150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, sProduto)

        Case 44151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, sProduto)

        Case 44152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44153, 44154, 44155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142538)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function FilialEmpresa_Le1(ByVal lComando As Long, ByVal objFilialEmpresa As AdmFiliais) As Long
'carrega dados da FilialEmpresa da tabela FiliaisEmpresa
'Le APENAS no SGE

Dim lErro As Long
Dim tFilialEmpresa As typeFiliais
Dim lEndereco As Long

On Error GoTo Erro_FilialEmpresa_Le1

    tFilialEmpresa.sCgc = String(STRING_CGC, 0)
    tFilialEmpresa.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tFilialEmpresa.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tFilialEmpresa.sISSCodigoPadrao = String(STRING_ISS_CODIGO, 0)
    tFilialEmpresa.sNome = String(STRING_FILIAL_NOME, 0)
    tFilialEmpresa.sNFCECSC = String(STRING_NFCE_CSC, 0)
    tFilialEmpresa.sidNFCECSC = String(STRING_NFCE_ID_CSC, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Nome, CGC, Endereco, AlmoxarifadoPadrao, InscricaoEstadual, ICMSPorEstimativa, InscricaoMunicipal, ISSPercPadrao, ISSCodigoPadrao, ISSIncluso, ContribuinteIPI, IRPercPadrao, TipoTribAtividade, SimplesFederal, SimplesFederalAliq, SimplesFederalTeto, SuperSimples, PISNaoCumulativo, COFINSNaoCumulativo, LucroPresumido, NFCECSC, idNFCECSC, indSincPadrao, RegimeTrib, RegimeEspecialTrib FROM FiliaisEmpresa WHERE FilialEmpresa = ?", _
        tFilialEmpresa.sNome, tFilialEmpresa.sCgc, lEndereco, tFilialEmpresa.iAlmoxarifadoPadrao, tFilialEmpresa.sInscricaoEstadual, tFilialEmpresa.iICMSPorEstimativa, tFilialEmpresa.sInscricaoMunicipal, tFilialEmpresa.dISSPercPadrao, tFilialEmpresa.sISSCodigoPadrao, tFilialEmpresa.iISSIncluso, tFilialEmpresa.iContribuinteIPI, tFilialEmpresa.dIRPercPadrao, tFilialEmpresa.iTipoTribAtividade, tFilialEmpresa.iSimplesFederal, tFilialEmpresa.dSimplesFederalAliq, tFilialEmpresa.dSimplesFederalTeto, _
        tFilialEmpresa.iSuperSimples, tFilialEmpresa.iPISNaoCumulativo, tFilialEmpresa.iCOFINSNaoCumulativo, tFilialEmpresa.iLucroPresumido, tFilialEmpresa.sNFCECSC, tFilialEmpresa.sidNFCECSC, tFilialEmpresa.iindSincPadrao, tFilialEmpresa.iRegimeTrib, tFilialEmpresa.iRegimeEspecialTrib, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 44210
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44211
    
    If lErro <> AD_SQL_SUCESSO Then Error 44212
    
    'mover campos p/o objeto
    objFilialEmpresa.dISSPercPadrao = tFilialEmpresa.dISSPercPadrao
    objFilialEmpresa.iAlmoxarifadoPadrao = tFilialEmpresa.iAlmoxarifadoPadrao
    objFilialEmpresa.iContribuinteIPI = tFilialEmpresa.iContribuinteIPI
    objFilialEmpresa.iICMSPorEstimativa = tFilialEmpresa.iICMSPorEstimativa
    objFilialEmpresa.iISSIncluso = tFilialEmpresa.iISSIncluso
    objFilialEmpresa.lCodEmpresa = tFilialEmpresa.lCodEmpresa
    objFilialEmpresa.sCgc = tFilialEmpresa.sCgc
    objFilialEmpresa.sInscricaoEstadual = tFilialEmpresa.sInscricaoEstadual
    objFilialEmpresa.sInscricaoMunicipal = tFilialEmpresa.sInscricaoMunicipal
    objFilialEmpresa.sISSCodigoPadrao = tFilialEmpresa.sISSCodigoPadrao
    objFilialEmpresa.sNome = tFilialEmpresa.sNome
    objFilialEmpresa.dIRPercPadrao = tFilialEmpresa.dIRPercPadrao
    objFilialEmpresa.iTipoTribAtividade = tFilialEmpresa.iTipoTribAtividade
    objFilialEmpresa.iSimplesFederal = tFilialEmpresa.iSimplesFederal
    objFilialEmpresa.dSimplesFederalAliq = tFilialEmpresa.dSimplesFederalAliq
    objFilialEmpresa.dSimplesFederalTeto = tFilialEmpresa.dSimplesFederalTeto
    objFilialEmpresa.iSuperSimples = tFilialEmpresa.iSuperSimples
    objFilialEmpresa.iPISNaoCumulativo = tFilialEmpresa.iPISNaoCumulativo
    objFilialEmpresa.iCOFINSNaoCumulativo = tFilialEmpresa.iCOFINSNaoCumulativo
    objFilialEmpresa.iLucroPresumido = tFilialEmpresa.iLucroPresumido
    
    objFilialEmpresa.sNFCECSC = tFilialEmpresa.sNFCECSC
    objFilialEmpresa.sidNFCECSC = tFilialEmpresa.sidNFCECSC
    objFilialEmpresa.iindSincPadrao = tFilialEmpresa.iindSincPadrao
    objFilialEmpresa.iRegimeTrib = tFilialEmpresa.iRegimeTrib
    objFilialEmpresa.iRegimeEspecialTrib = tFilialEmpresa.iRegimeEspecialTrib
    
    FilialEmpresa_Le1 = SUCESSO

    Exit Function

Erro_FilialEmpresa_Le1:

    FilialEmpresa_Le1 = Err

    Select Case Err

        Case 44210, 44211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALEMPRESA", Err, objFilialEmpresa.iCodFilial)
        
        Case 44212 'filial nao encontrada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142539)

    End Select

    Exit Function

End Function

Function UsuarioModulo_Le_UsuarioEmpresa(ByVal colUsuarioModulo As Collection, ByVal sUsuario As String, ByVal lEmpresa As Long, ByVal iFilialEmpresa As Integer) As Long
'le todos os modulos validos para o usuario/empresa/filial passados como parametro e coloca-os em colModulo

Dim lComando As Long
Dim lErro As Long
Dim objUsuarioModulo As ClassUsuarioModulo
Dim tUsuarioModulo As typeUsuarioModulo

On Error GoTo Erro_UsuarioModulo_Le_UsuarioEmpresa

    lComando = 0

    tUsuarioModulo.sNomeModulo = String(STRING_MODULO_NOME, 0)
    tUsuarioModulo.sSiglaModulo = String(STRING_MODULO_SIGLA, 0)
    tUsuarioModulo.sVersaoModulo = String(STRING_MODULO_VERSAO, 0)

    'abre conexao com o BD
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 44415
    
    'le todos os registros da view UsuarioModulo
    lErro = Comando_Executar(lComando, "SELECT Sigla, Nome, Versao, DataValidade, Menu FROM UsuarioModulo WHERE CodUsuario=? AND CodEmpresa=? AND CodFilial=?", tUsuarioModulo.sSiglaModulo, tUsuarioModulo.sNomeModulo, tUsuarioModulo.sVersaoModulo, tUsuarioModulo.dtDataValidade, tUsuarioModulo.iMenuModulo, sUsuario, lEmpresa, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 44416
    
    'le o primeiro modulo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44417
    
    Do While lErro <> AD_SQL_SEM_DADOS

        If tUsuarioModulo.dtDataValidade >= gdtDataAtual Or tUsuarioModulo.dtDataValidade = DATA_NULA Then

            Set objUsuarioModulo = New AdmModulo
        
            objUsuarioModulo.dtDataValidade = tUsuarioModulo.dtDataValidade
            objUsuarioModulo.iCodFilial = iFilialEmpresa
            objUsuarioModulo.iMenuModulo = tUsuarioModulo.iMenuModulo
            objUsuarioModulo.lCodEmpresa = lEmpresa
            objUsuarioModulo.sCodUsuario = sUsuario
            objUsuarioModulo.sNomeModulo = tUsuarioModulo.sNomeModulo
            objUsuarioModulo.sSiglaModulo = tUsuarioModulo.sSiglaModulo
            objUsuarioModulo.sVersaoModulo = tUsuarioModulo.sVersaoModulo
            
            'coloca o UsuarioModulo lido na coleção
            colUsuarioModulo.Add objUsuarioModulo
        
        End If
        
        'le o proximo modulo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44418

    Loop

    Call Comando_Fechar(lComando)
    
    UsuarioModulo_Le_UsuarioEmpresa = SUCESSO

    Exit Function

Erro_UsuarioModulo_Le_UsuarioEmpresa:

    UsuarioModulo_Le_UsuarioEmpresa = Err
    
    Select Case Err

        Case 44415
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44416, 44417, 44418
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOMODULO", Err, sUsuario, lEmpresa, iFilialEmpresa)
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142540)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Paises_Le(ByVal objPais As ClassPais) As Long
'Lê o Pais a partir da chave fornecida em (Código) e retorna os dados em objPais

Dim lComando As Long
Dim lErro As Long
Dim sNome As String
Dim iCodBacen As Integer

On Error GoTo Erro_Paises_Le

    lComando = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 47873

    sNome = String(STRING_PAISES_NOME, 0)

    'Seleciona o Pais cujo Código está no Obj
    lErro = Comando_Executar(lComando, "SELECT Nome, CodBacen FROM Paises WHERE Codigo = ?", sNome, iCodBacen, objPais.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 47874

    'Posiciona no primeiro Pais dos que foram lidos
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 47875

    'Se não encontrou nenhum Pais com o Código em questão
    If lErro = AD_SQL_SEM_DADOS Then Error 47876

    'Passa o Nome do Pais lido para obj
    objPais.sNome = sNome
    objPais.iCodBacen = iCodBacen

    Call Comando_Fechar(lComando)

    Paises_Le = SUCESSO

    Exit Function

Erro_Paises_Le:

    Paises_Le = Err

    Select Case Err

        Case 47873
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 47874, 47875
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_PAISES", Err)

        Case 47876

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142541)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function UsuarioModulos_Le(ByVal objUsuarioModulo As ClassUsuarioModulo, ByVal collCodigoNome As AdmCollCodigoNome) As Long
'Lê os Módulos acessíveis ao usuário, empresa e filial passados como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim sNome As String
Dim iMenu As Integer
Dim objlCodigoNome As AdmlCodigoNome

On Error GoTo Erro_UsuarioModulos_Le

    'Iniciliza comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 43631

    sNome = String(STRING_MODULO_NOME, 0)

    lErro = Comando_Executar(lComando, "SELECT Nome, Menu FROM UsuarioModulo WHERE Menu <> 0 AND CodUsuario = ? AND CodEmpresa = ? AND CodFilial = ? AND (DataValidade >= ? OR DataValidade = ?) ORDER BY Nome", sNome, iMenu, objUsuarioModulo.sCodUsuario, objUsuarioModulo.lCodEmpresa, objUsuarioModulo.iCodFilial, objUsuarioModulo.dtDataValidade, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then Error 43632

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43633

    Do While lErro = AD_SQL_SUCESSO

        If sNome <> "Geral" Then
        
            Set objlCodigoNome = New AdmlCodigoNome
    
            objlCodigoNome.sNome = sNome
            objlCodigoNome.lCodigo = iMenu
    
            collCodigoNome.Add objlCodigoNome.lCodigo, objlCodigoNome.sNome
        
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43645

    Loop

    Call Comando_Fechar(lComando)

    UsuarioModulos_Le = SUCESSO

    Exit Function

Erro_UsuarioModulos_Le:

    UsuarioModulos_Le = Err

    Select Case Err

        Case 43631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 43632, 43633, 43645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOMODULO", Err, objUsuarioModulo.sCodUsuario, objUsuarioModulo.lCodEmpresa, objUsuarioModulo.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142542)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function UsuarioItensMenu_Le(ByVal objUsuItensMenu As ClassUsuarioItensMenu, ByVal colUsuarioItensMenu As Collection) As Long
'Lê os itens de menu acessíveis ao usuário, empresa e filial passados como parâmetro ordenados por nome e indice do controle pai

Dim lErro As Long
Dim lComando As Long
Dim sNomeControle As String
Dim iIndiceControle As Integer
Dim sNomeControlePai As String
Dim sNomeTela As String
Dim iIndiceControlePai As Integer
Dim objUsuarioItensMenu As ClassUsuarioItensMenu

On Error GoTo Erro_UsuarioItensMenu_Le

    'Iniciliza comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 43640

    sNomeControle = String(STRING_BUFFER_MAX_TEXTO, 0)
    sNomeControlePai = String(STRING_BUFFER_MAX_TEXTO, 0)
    sNomeTela = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'Verifica se é EMPRESA_TODA
    If objUsuItensMenu.iCodFilial = EMPRESA_TODA Or objUsuItensMenu.iCodFilial = Abs(giFilialAuxiliar) Then

        lErro = Comando_Executar(lComando, "SELECT NomeTela, NomeControle, IndiceControle, NomeControlePai, IndiceControlePai FROM UsuarioItensMenu WHERE CodUsuario = ? AND CodEmpresa = ? AND CodFilial = ? AND (DataValidade >= ? OR DataValidade = ?) AND TipoDeAcesso = ? AND ((FilialEmpresa = ?) OR (FilialEmpresa = ?)) ORDER BY NomeControlePai, IndiceControlePai", _
            sNomeTela, sNomeControle, iIndiceControle, sNomeControlePai, iIndiceControlePai, objUsuItensMenu.sCodUsuario, objUsuItensMenu.lCodEmpresa, objUsuItensMenu.iCodFilial, objUsuItensMenu.dtDataValidade, DATA_NULA, USUARIO_TEM_ACESSO, EMPRESA_TODA, FILIAL_EMPRESA_TODA)
        If lErro <> AD_SQL_SUCESSO Then Error 43641
    Else
        lErro = Comando_Executar(lComando, "SELECT NomeTela, NomeControle, IndiceControle, NomeControlePai, IndiceControlePai FROM UsuarioItensMenu WHERE CodUsuario = ? AND CodEmpresa = ? AND CodFilial = ? AND (DataValidade >= ? OR DataValidade = ?) AND TipoDeAcesso = ? AND ((FilialEmpresa = ?) OR (FilialEmpresa = ?)) ORDER BY NomeControlePai, IndiceControlePai", _
            sNomeTela, sNomeControle, iIndiceControle, sNomeControlePai, iIndiceControlePai, objUsuItensMenu.sCodUsuario, objUsuItensMenu.lCodEmpresa, objUsuItensMenu.iCodFilial, objUsuItensMenu.dtDataValidade, DATA_NULA, USUARIO_TEM_ACESSO, FILIAL_EMPRESA, FILIAL_EMPRESA_TODA)
        If lErro <> AD_SQL_SUCESSO Then Error 43642
    
    End If
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43643

    Do While lErro = AD_SQL_SUCESSO

        Set objUsuarioItensMenu = New ClassUsuarioItensMenu

        objUsuarioItensMenu.sNomeControle = sNomeControle
        objUsuarioItensMenu.iIndiceControle = iIndiceControle
        objUsuarioItensMenu.sNomeControlePai = sNomeControlePai
        objUsuarioItensMenu.iIndiceControlePai = iIndiceControlePai
        objUsuarioItensMenu.sNomeTela = sNomeTela

        colUsuarioItensMenu.Add objUsuarioItensMenu

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43644

    Loop

    Call Comando_Fechar(lComando)

    UsuarioItensMenu_Le = SUCESSO

    Exit Function

Erro_UsuarioItensMenu_Le:

    UsuarioItensMenu_Le = Err

    Select Case Err

        Case 43640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 43641, 43642, 43643, 43644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOITENSMENU", Err, objUsuItensMenu.sCodUsuario, objUsuItensMenu.lCodEmpresa, objUsuItensMenu.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142543)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MenuItens_Le(ByVal colMenuItens As Collection) As Long
'Lê os itens de menu ordenados por nome, indice do controle pai, posicao do controle em relacao ao pai e nivel (descendente)

Dim tMenuItens As typeMenuItens
Dim objMenuItens As ClassMenuItens
Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_MenuItens_Le

    'Iniciliza comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 44630

    tMenuItens.sNomeTela = String(STRING_NOME_TELA, 0)
    tMenuItens.sNomeControle = String(STRING_NOME_CONTROLE, 0)
    tMenuItens.sNomeControlePai = String(STRING_NOME_CONTROLE, 0)
    tMenuItens.sSiglaRotina = String(STRING_SIGLA_ROTINA, 0)
    tMenuItens.sTitulo = String(STRING_TITULO_MENU, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Identificador, Titulo, SiglaRotina, NomeTela, NomeControle, IndiceControle, NomeControlePai, IndiceControlePai, IndiceControleNoPai, Separador FROM MenuItens ORDER BY Nivel DESC, NomeControlePai, IndiceControlePai, IndiceControleNoPai", _
        tMenuItens.iIdentificador, tMenuItens.sTitulo, tMenuItens.sSiglaRotina, tMenuItens.sNomeTela, tMenuItens.sNomeControle, tMenuItens.iIndiceControle, tMenuItens.sNomeControlePai, tMenuItens.iIndiceControlePai, tMenuItens.iIndiceControleNoPai, tMenuItens.iSeparador)
    If lErro <> AD_SQL_SUCESSO Then Error 44627

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44628

    Do While lErro = AD_SQL_SUCESSO

        Set objMenuItens = New ClassMenuItens

        With objMenuItens
            .iIdentificador = tMenuItens.iIdentificador
            .sTitulo = tMenuItens.sTitulo
            .sSiglaRotina = tMenuItens.sSiglaRotina
            .sNomeTela = tMenuItens.sNomeTela
            .sNomeControle = tMenuItens.sNomeControle
            .iIndiceControle = tMenuItens.iIndiceControle
            .sNomeControlePai = tMenuItens.sNomeControlePai
            .iIndiceControlePai = tMenuItens.iIndiceControlePai
            .iIndiceControleNoPai = tMenuItens.iIndiceControleNoPai
            .iSeparador = tMenuItens.iSeparador
        End With
        
        colMenuItens.Add objMenuItens

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44629

    Loop

    Call Comando_Fechar(lComando)

    MenuItens_Le = SUCESSO

    Exit Function

Erro_MenuItens_Le:

    MenuItens_Le = Err

    Select Case Err

        Case 44630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44627, 44628, 44629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MENUITENS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142544)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'*** INICIO

Function Usuarios_Le(ByVal objUsuarios As ClassUsuarios, Optional ByVal bSomenteAtivos = True) As Long
'le  o usuário contido na tabela de Usuario pelo parametro CodUsuario e coloca no objeto

Dim lComando As Long
Dim lErro As Long
Dim tUsuarios As typeUsuarios
    
On Error GoTo Erro_Usuarios_Le

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 40829
    
    tUsuarios.sNome = String(STRING_USUARIO, 0)
    tUsuarios.sSenha = String(STRING_USUARIO_SENHA, 0)
    tUsuarios.sCodGrupo = String(STRING_GRUPO, 0)
    tUsuarios.sNomeReduzido = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    tUsuarios.sEmail = String(64, 0)
        
    If bSomenteAtivos Then
        lErro = Comando_Executar(lComando, "SELECT Nome, NomeReduzido, CodGrupo, Senha, DataValidade, Ativo, WorkflowAtivo, Email FROM Usuarios WHERE CodUsuario=? AND Ativo=? AND (DataValidade=? OR DataValidade>=?)", _
            tUsuarios.sNome, tUsuarios.sNomeReduzido, tUsuarios.sCodGrupo, tUsuarios.sSenha, tUsuarios.dtDataValidade, tUsuarios.iAtivo, tUsuarios.iWorkFlowAtivo, tUsuarios.sEmail, objUsuarios.sCodUsuario, USUARIO_ATIVO, DATA_NULA, gdtDataAtual)
    Else
        lErro = Comando_Executar(lComando, "SELECT Nome, NomeReduzido, CodGrupo, Senha, DataValidade, Ativo, WorkflowAtivo, Email FROM Usuarios WHERE CodUsuario=?", _
            tUsuarios.sNome, tUsuarios.sNomeReduzido, tUsuarios.sCodGrupo, tUsuarios.sSenha, tUsuarios.dtDataValidade, tUsuarios.iAtivo, tUsuarios.iWorkFlowAtivo, tUsuarios.sEmail, objUsuarios.sCodUsuario)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 40830
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 40831
    
    If lErro = AD_SQL_SEM_DADOS Then Error 40832
    
    objUsuarios.sCodGrupo = tUsuarios.sCodGrupo
    objUsuarios.sNome = tUsuarios.sNome
    objUsuarios.sNomeReduzido = tUsuarios.sNomeReduzido
    objUsuarios.sSenha = tUsuarios.sSenha
    objUsuarios.dtDataValidade = tUsuarios.dtDataValidade
    objUsuarios.iAtivo = tUsuarios.iAtivo
    objUsuarios.iWorkFlowAtivo = tUsuarios.iWorkFlowAtivo
    objUsuarios.sEmail = tUsuarios.sEmail
               
    Call Comando_Fechar(lComando)
    
    Usuarios_Le = SUCESSO
    
    Exit Function
    
Erro_Usuarios_Le:

    Usuarios_Le = Err

    Select Case Err
    
        Case 40829
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 40830, 40831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", Err)
        
        Case 40832
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142545)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Feriado_Le(ByVal objFeriado As ClassFeriado) As Long
'Lê o Feriado à partir dos dados passados como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim sDescricao As String
    
On Error GoTo Erro_Feriado_Le

    sDescricao = String(STRING_FERIADO, 0)
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43376
    
    'Lê o Feriado
    lErro = Comando_Executar(lComando, "SELECT Descricao FROM Feriados WHERE FilialEmpresa = ? AND Data = ?", sDescricao, objFeriado.iFilialEmpresa, objFeriado.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 43377
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43378
    
    'Se não encontrou o Feriado --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 43379
    
    objFeriado.sDescricao = sDescricao
    
    Call Comando_Fechar(lComando)
    
    Feriado_Le = SUCESSO
    
    Exit Function
    
Erro_Feriado_Le:

    Feriado_Le = Err

    Select Case Err
    
        Case 43376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 43377, 43378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FERIADOS", Err, objFeriado.iFilialEmpresa, objFeriado.dtData)
        
        Case 43379 'a data não é feriado para a filial informada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142546)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Feriado_Le_Todos(ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal colFeriados As Collection) As Long
'Lê todos os Feriados contidos na tabela de Feriados

Dim lErro As Long
Dim lComando As Long
Dim objFeriado As ClassFeriado
Dim dtData As Date
Dim sDescrFeriado As String
    
On Error GoTo Erro_Feriado_Le_Todos

    sDescrFeriado = String(STRING_FERIADO, 0)

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43412

    'Lê a tabela de Feriados
    lErro = Comando_Executar(lComando, "SELECT Data, Descricao FROM Feriados WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? ORDER BY Data", dtData, sDescrFeriado, giFilialEmpresa, dtDataInicial, dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then Error 43413

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43414
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objFeriado = New ClassFeriado
        
        objFeriado.dtData = dtData
        objFeriado.sDescricao = sDescrFeriado
        
        'Coloca o Feriado lido na coleção
        colFeriados.Add objFeriado
        
        'Lê o próximo Feriado
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43415

    Loop
    
    Call Comando_Fechar(lComando)
    
    Feriado_Le_Todos = SUCESSO

    Exit Function

Erro_Feriado_Le_Todos:

    Feriado_Le_Todos = Err
    
    Select Case Err

        Case 43412
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 43413, 43414, 43415
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FERIADOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142547)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Relatorio_Le_Reimpressao(ByVal objRelatorio As AdmRelatorio) As Long
'Retorna o nome do Arquivo de Reimpressão sse houver

Dim lErro As Long
Dim lComando As Long
Dim iFim_de_Arquivo As Integer
Dim sDescricao As String
Dim sNomeArqReimp As String
    
On Error GoTo Erro_Relatorio_Le_Reimpressao
    
    sNomeArqReimp = String(STRING_NOME_ARQ_REIMP, 0)
    sDescricao = String(STRING_DESC_RELATORIO, 0)
          
    'Abre conexão com o BD
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 43307
    
    'Faz a seleção dos dados na tabela Relatorios
    lErro = Comando_Executar(lComando, "SELECT  Descricao, NomeArqReimp FROM Relatorios WHERE CodRel = ? ORDER BY CodRel", sDescricao, sNomeArqReimp, objRelatorio.sCodRel)
    If lErro <> AD_SQL_SUCESSO Then Error 43308
    
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    
    If iFim_de_Arquivo <> AD_SQL_SUCESSO Then Error 43309
           
    objRelatorio.sDescricao = sDescricao
    objRelatorio.sNomeArqReimp = sNomeArqReimp
    
    Call Comando_Fechar(lComando)
    
    Relatorio_Le_Reimpressao = SUCESSO
    
    Exit Function

Erro_Relatorio_Le_Reimpressao:

    Relatorio_Le_Reimpressao = Err
    
    Select Case Err
                
        Case 43307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 43308, 43309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142548)
        
    End Select
        
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Function ModuloFilEmp_Quer_Config(ByVal lCodEmpresa As Long, ByVal iFilialEmpresa As Integer, ByVal colModulos As Collection) As Long
'devolve os módulos que devem ser configurados para a empresa/filial em questãp

Dim lComando As Long
Dim lErro As Long
Dim sSiglaModulo As String
    
On Error GoTo Erro_ModuloFilEmp_Quer_Config

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 44848

    sSiglaModulo = String(STRING_MODULO_SIGLA, 0)

    lErro = Comando_Executar(lComando, "SELECT SiglaModulo FROM ModuloFilEmp  WHERE CodEmpresa=? AND CodFilial=? AND Configurado=?", sSiglaModulo, lCodEmpresa, iFilialEmpresa, NAO_CONFIGURADO)
    If lErro <> AD_SQL_SUCESSO Then Error 44849
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44850
    
    Do While lErro = AD_SQL_SUCESSO
    
        colModulos.Add sSiglaModulo
               
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44851
               
    Loop
               
    Call Comando_Fechar(lComando)
    
    ModuloFilEmp_Quer_Config = SUCESSO
    
    Exit Function
    
Erro_ModuloFilEmp_Quer_Config:

    ModuloFilEmp_Quer_Config = Err

    Select Case Err
    
        Case 44848
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 44849, 44850, 44851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULOFILEMP_DIC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142549)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Niveis0e1_Componentes(ByVal colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto
'produtos que podem participar de kit

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto
Dim objProduto As ClassProduto

On Error GoTo Erro_Produto_Le_Niveis0e1_Componentes

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49953

    'Inicializa as variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)

        'Lê os Produto de Nivel 1 e 2
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza,TempoProducao FROM Produtos WHERE (KitBasico=? OR KitInt =? OR Gerencial=? ) AND Ativo=? AND (Nivel = 1 OR Nivel = 2) ORDER BY Codigo", _
            .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, .iTempoProducao, KIT_BASICO, KIT_INTERMEDIARIO, GERENCIAL, Ativo)
        If lErro <> AD_SQL_SUCESSO Then Error 49954

        'le o primeiro produto
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49955

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto
    
            'Coloca os dados lidos no objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
            objProduto.sContaContabilProducao = .sContaContabilProducao
            objProduto.dResiduo = .dResiduo
            objProduto.iNatureza = .iNatureza
            objProduto.iTempoProducao = .iTempoProducao
            
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49956

        Loop

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Produto_Le_Niveis0e1_Componentes = SUCESSO

    Exit Function

Erro_Produto_Le_Niveis0e1_Componentes:

    Produto_Le_Niveis0e1_Componentes = Err

    Select Case Err

        Case 49953
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49954, 49955, 49956
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142550)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Netos_Componentes(ByVal sProduto As String, ByVal colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto
'produtos que podem ser Kit(Intermediario ou basico)

Dim lComando As Long
Dim lErro As Long
Dim objProduto As ClassProduto
Dim tProduto As typeProduto
Dim sProduto1 As String
Dim sProdutoEnxuto As String
Dim iNivel As Integer

On Error GoTo Erro_Produto_Le_Netos_Componentes

    'Inicializa as Variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 49960

        lErro = Mascara_Produto_ObterNivel(sProduto, iNivel)
        If lErro <> SUCESSO Then Error 49961

        sProduto1 = sProdutoEnxuto + String(STRING_PRODUTO - Len(sProdutoEnxuto), "Z")

        lComando = Comando_Abrir()
        If lComando = 0 Then Error 49962
    
        'Lê produtos ordenados por Codigo (apenas os atributos que interessam)
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza,TempoProducao,Rastro,HorasMaquina,PesoEspecifico FROM Produtos WHERE (KitBasico=? OR KitInt=? OR Gerencial=?) AND Ativo=? AND Nivel = ? AND Codigo > ? AND Codigo < ? ORDER BY Codigo", _
            .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, .iTempoProducao, .iRastro, .lHorasMaquina, .dPesoEspecifico, KIT_BASICO, KIT_INTERMEDIARIO, GERENCIAL, Ativo, iNivel + 2, sProduto, sProduto1)
        If lErro <> AD_SQL_SUCESSO Then Error 49963

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49964

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto

            'Guarda so dados lidos em objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
            objProduto.sContaContabilProducao = .sContaContabilProducao
            objProduto.dResiduo = .dResiduo
            objProduto.iNatureza = .iNatureza
            objProduto.iTempoProducao = .iTempoProducao
            objProduto.iRastro = .iRastro
            objProduto.lHorasMaquina = .lHorasMaquina
            objProduto.dPesoEspecifico = .dPesoEspecifico
            
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49965

        Loop

    End With

    Call Comando_Fechar(lComando)
    
    Produto_Le_Netos_Componentes = SUCESSO

    Exit Function

Erro_Produto_Le_Netos_Componentes:

    Produto_Le_Netos_Componentes = Err
    
    Select Case Err

        Case 49960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, sProduto)

        Case 49961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, sProduto)

        Case 49962
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49963, 49964, 49965
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142551)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function



Public Function Modulos_Le_Empresa_Filial(ByVal lEmpresa As Long, ByVal iFilialEmpresa As Integer, ByVal colModulo As AdmColModulo) As Long
'Le todos os modulos associados a lEmpresa/iFilialEmpresa contidos na tabela Modulos no dicionario de dados e coloca-os em colModulo

Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tModulo As typeModulo
Dim lConexao As Long

On Error GoTo Erro_Modulos_Le_Empresa_Filial

    lConexao = GL_lConexaoDic

    tModulo.sSigla = String(STRING_MODULO_SIGLA, 0)
    tModulo.sNome = String(STRING_MODULO_NOME, 0)
    tModulo.sDescricao = String(STRING_MODULO_DESCRICAO, 0)
    tModulo.sVersao = String(STRING_MODULO_VERSAO, 0)
    tModulo.sOrigemCusto = String(STRING_ORIGEM, 0)
    tModulo.sOrigemAglutina = String(STRING_ORIGEM, 0)

    'abre comandos para DIC
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 25463

    lComando2 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 25880
    
    'le todos os Modulos que estão ativos na Filial
    lErro = Comando_Executar(lComando1, "SELECT Sigla, Nome, Descricao, Versao, OrigemCusto, OrigemAglutina, Ativo FROM FilialEmpresaModulos WHERE CodEmpresa = ? AND CodFilial= ? ORDER BY Sigla", tModulo.sSigla, tModulo.sNome, tModulo.sDescricao, tModulo.sVersao, tModulo.sOrigemCusto, tModulo.sOrigemAglutina, tModulo.iAtivo, lEmpresa, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 9642
    
    'le o primeiro modulo
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9643
    
    Do While lErro <> AD_SQL_SEM_DADOS

        'coloca o modulo lido na coleção
        colModulo.Add tModulo.sSigla, tModulo.sNome, tModulo.sDescricao, tModulo.sVersao, tModulo.sOrigemCusto, tModulo.sOrigemAglutina, tModulo.iAtivo, tModulo.sSigla
        
        'le o proximo modulo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25465
        
    Loop
    
    tModulo.sSigla = String(STRING_MODULO_SIGLA, 0)
    tModulo.sNome = String(STRING_MODULO_NOME, 0)
    tModulo.sDescricao = String(STRING_MODULO_DESCRICAO, 0)
    tModulo.sVersao = String(STRING_MODULO_VERSAO, 0)
    tModulo.sOrigemCusto = String(STRING_ORIGEM, 0)
    tModulo.sOrigemAglutina = String(STRING_ORIGEM, 0)
    
    'le todos os Modulos que não estão ativos na Filial
    If giSQLTipoOrdParamDic = SQL_ORD_PARAM_ESQ_DIR Then
        lErro = Comando_Executar(lComando2, "SELECT DISTINCT Sigla, Nome, Descricao, Versao, OrigemCusto, OrigemAglutina FROM FilialEmpresaModulos WHERE CodEmpresa = ? AND Sigla Not In (SELECT Sigla From FilialEmpresaModulos WHERE CodEmpresa=? AND CodFilial=?) ORDER BY Sigla", _
            tModulo.sSigla, tModulo.sNome, tModulo.sDescricao, tModulo.sVersao, tModulo.sOrigemCusto, tModulo.sOrigemAglutina, lEmpresa, lEmpresa, iFilialEmpresa)
    Else
        lErro = Comando_Executar(lComando2, "SELECT DISTINCT Sigla, Nome, Descricao, Versao, OrigemCusto, OrigemAglutina FROM FilialEmpresaModulos WHERE CodEmpresa = ? AND Sigla Not In (SELECT Sigla From FilialEmpresaModulos WHERE CodEmpresa=? AND CodFilial=?) ORDER BY Sigla", _
            tModulo.sSigla, tModulo.sNome, tModulo.sDescricao, tModulo.sVersao, tModulo.sOrigemCusto, tModulo.sOrigemAglutina, lEmpresa, iFilialEmpresa, lEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 25878
    
    'le o primeiro modulo
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25877
    
    'Modulos inativos p/ essa Filial
    tModulo.iAtivo = 0
    
    Do While lErro <> AD_SQL_SEM_DADOS

        'coloca o modulo lido na coleção
        colModulo.Add tModulo.sSigla, tModulo.sNome, tModulo.sDescricao, tModulo.sVersao, tModulo.sOrigemCusto, tModulo.sOrigemAglutina, tModulo.iAtivo, tModulo.sSigla
        
        'le o proximo modulo
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25879
        
    Loop
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Modulos_Le_Empresa_Filial = SUCESSO

    Exit Function

Erro_Modulos_Le_Empresa_Filial:

    Modulos_Le_Empresa_Filial = Err
    
    Select Case Err

        Case 25463, 25880
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9642, 9643, 25465, 25877, 25878, 25879
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULO", Err)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142552)

    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function
    
End Function

Function Produto_Le_Niveis0e1_Inventariado(ByVal colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto
Dim objProduto As ClassProduto

On Error GoTo Erro_Produto_Le_Niveis0e1_Inventariado

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 52135

    'Inicializa as variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sDescricao = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMEstoque = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMCompra = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMVenda = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSubstituto1 = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSubstituto2 = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNomeReduzido = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sModelo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCodigoBarras = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCor = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sIPICodDIPI = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sIPICodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sObsFisica = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabil = String(STRING_BUFFER_MAX_TEXTO, 0)


        'Lê os Produto de Nivel 1 e 2
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil FROM ProdutosProdutosFilial WHERE ((FilialEmpresa = ? AND ControleEstoque<>?) OR Gerencial=?) AND Ativo=? AND (Nivel = 1 OR Nivel = 2) ORDER BY Codigo", .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, giFilialEmpresa, PRODUTO_CONTROLE_SEM_ESTOQUE, GERENCIAL, Ativo)
        If lErro <> AD_SQL_SUCESSO Then Error 52136

        'le o primeiro produto
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52137
        
        If lErro = AD_SQL_SEM_DADOS Then Error 52138
        
        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto
    
            'Coloca os dados lidos no objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
            
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52139

        Loop

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Produto_Le_Niveis0e1_Inventariado = SUCESSO

    Exit Function

Erro_Produto_Le_Niveis0e1_Inventariado:

    Produto_Le_Niveis0e1_Inventariado = Err

    Select Case Err

        Case 52135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 52136, 52137, 52139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)
        
        Case 52138
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142553)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Netos_Inventariado(ByVal sProduto As String, ByVal colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim lComando As Long
Dim lErro As Long
Dim objProduto As ClassProduto
Dim tProduto As typeProduto
Dim sProduto1 As String
Dim sProdutoEnxuto As String
Dim iNivel As Integer

On Error GoTo Erro_Produto_Le_Netos_Inventariado

    'Inicializa as Variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sDescricao = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMEstoque = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMCompra = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMVenda = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSubstituto1 = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSubstituto2 = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNomeReduzido = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sModelo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCodigoBarras = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCor = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sIPICodDIPI = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sIPICodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sObsFisica = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabil = String(STRING_BUFFER_MAX_TEXTO, 0)

        lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 52146

        lErro = Mascara_Produto_ObterNivel(sProduto, iNivel)
        If lErro <> SUCESSO Then Error 52147

        sProduto1 = sProdutoEnxuto + String(STRING_PRODUTO - Len(sProdutoEnxuto), "Z")

        lComando = Comando_Abrir()
        If lComando = 0 Then Error 52148
    
        'Lê produtos ordenados por Codigo (apenas os atributos que interessam)
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil FROM ProdutosProdutosFilial WHERE ((FilialEmpresa = ? AND ControleEstoque<>?) OR Gerencial=?) AND Ativo=? AND Nivel = ? AND Codigo > ? AND Codigo < ? ORDER BY Codigo", .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, giFilialEmpresa, PRODUTO_CONTROLE_SEM_ESTOQUE, GERENCIAL, Ativo, iNivel + 2, sProduto, sProduto1)
        If lErro <> AD_SQL_SUCESSO Then Error 52149

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52150

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto

            'Guarda so dados lidos em objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
    
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52151

        Loop

    End With

    Call Comando_Fechar(lComando)
    
    Produto_Le_Netos_Inventariado = SUCESSO

    Exit Function

Erro_Produto_Le_Netos_Inventariado:

    Produto_Le_Netos_Inventariado = Err
    
    Select Case Err

        Case 52146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, sProduto)

        Case 52147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, sProduto)

        Case 52148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 52149, 52150, 52151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142554)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Produto_Le_Niveis0e1_Venda(ByVal colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto
Dim objProduto As ClassProduto

On Error GoTo Erro_Produto_Le_Niveis0e1_Venda

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58061

    'Inicializa as variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)

        'Lê os Produto de Nivel 1 e 2
        lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza FROM ProdutosProdutosFilial WHERE ((FilialEmpresa = ? AND Faturamento=?) OR Gerencial=?) AND Ativo=? AND (Nivel = 1 OR Nivel = 2) ORDER BY Codigo", _
            .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, giFilialEmpresa, PRODUTO_VENDAVEL, GERENCIAL, Ativo)
        If lErro <> AD_SQL_SUCESSO Then Error 58062

        'le o primeiro produto
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58063

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto
    
            'Coloca os dados lidos no objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
            objProduto.sContaContabilProducao = .sContaContabilProducao
            objProduto.dResiduo = .dResiduo
            objProduto.iNatureza = .iNatureza
            
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58064

        Loop

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Produto_Le_Niveis0e1_Venda = SUCESSO

    Exit Function

Erro_Produto_Le_Niveis0e1_Venda:

    Produto_Le_Niveis0e1_Venda = Err

    Select Case Err

        Case 58061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 58062, 58063, 58064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142555)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Netos_Venda(ByVal sProduto As String, ByVal colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim lComando As Long
Dim lErro As Long
Dim objProduto As ClassProduto
Dim tProduto As typeProduto
Dim sProduto1 As String
Dim sProdutoEnxuto As String
Dim iNivel As Integer

On Error GoTo Erro_Produto_Le_Netos_Venda

    'Inicializa as Variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 58064

        lErro = Mascara_Produto_ObterNivel(sProduto, iNivel)
        If lErro <> SUCESSO Then Error 58065

        sProduto1 = sProdutoEnxuto + String(STRING_PRODUTO - Len(sProdutoEnxuto), "Z")

        lComando = Comando_Abrir()
        If lComando = 0 Then Error 58066
    
        'Lê produtos ordenados por Codigo (apenas os atributos que interessam)
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza FROM ProdutosProdutosFilial WHERE ((FilialEmpresa = ? AND Faturamento =?) OR Gerencial=?) AND Ativo=? AND Nivel = ? AND Codigo > ? AND Codigo < ? ORDER BY Codigo", _
                .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, giFilialEmpresa, PRODUTO_VENDAVEL, GERENCIAL, Ativo, iNivel + 2, sProduto, sProduto1)
        If lErro <> AD_SQL_SUCESSO Then Error 58067

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58068

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto

            'Guarda so dados lidos em objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
            objProduto.sContaContabilProducao = .sContaContabilProducao
            objProduto.dResiduo = .dResiduo
            objProduto.iNatureza = .iNatureza
    
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58069

        Loop

    End With

    Call Comando_Fechar(lComando)
    
    Produto_Le_Netos_Venda = SUCESSO

    Exit Function

Erro_Produto_Le_Netos_Venda:

    Produto_Le_Netos_Venda = Err
    
    Select Case Err

        Case 58064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, sProduto)

        Case 58065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, sProduto)

        Case 58066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 58067, 58068, 58069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142556)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Dias_Uteis_Antes(ByVal dtDataEntrada As Date, ByVal iDias As Integer, dtDataSaida As Date) As Long
'Retorna a data iDias úteis anterior a data de entrada
'???? está longe de ser uma implementacao eficiente. NAO OTIMIZAR AGORA.

Dim lErro As Long, iDiasAcum As Integer
Dim lComando As Long, dtData As Date

On Error GoTo Erro_Dias_Uteis_Antes

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56585
        
    iDiasAcum = 0
    
    dtData = dtDataEntrada - 1
    dtDataSaida = dtData
    
    'Loop que faz a contagem de dias úteis
    Do While iDiasAcum < iDias
    
        'Se for sábado ou domingo passamos para sexta
        If Weekday(dtData) = vbSaturday Then dtData = dtData - 1
        If Weekday(dtData) = vbSunday Then dtData = dtData - 2
        
        'Verifica se a Data é um feriado
        lErro = Comando_Executar(lComando, "SELECT Data FROM Feriados WHERE Data = ? AND FilialEmpresa = ? ", dtData, dtData, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 56586
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56587
        
        'Se não for conta a Data como sendo mais um Dia útil
        If lErro = AD_SQL_SEM_DADOS Then
            
            iDiasAcum = iDiasAcum + 1
            dtDataSaida = dtData

        End If
        
        'decrementa-se a data
        dtData = dtData - 1
            
    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Dias_Uteis_Antes = SUCESSO
    
    Exit Function
    
Erro_Dias_Uteis_Antes:

    Dias_Uteis_Antes = Err
    
    Select Case Err
    
        Case 56585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 56586, 56587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FERIADOS", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142557)
            
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Dias_Uteis(ByVal dtDataMenor As Date, ByVal dtDataMaior As Date, iDias As Integer) As Long
'Retorna o número de dias úteis que existem entre as datas passadas

Dim lErro As Long
Dim lComando As Long
Dim dtData As Date

On Error GoTo Erro_Dias_Uteis

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26209
        
    iDias = 0
    
    dtData = dtDataMenor
    
    'Loop que faz a contagem de dias úteis
    Do While dtData < dtDataMaior
    
        'Se for sábado ou domingo passamos para segunda
        If Weekday(dtData) = vbSaturday Then dtData = dtData + 2
        If Weekday(dtData) = vbSunday Then dtData = dtData + 1
        
        'Se ultrapassou o dia Maior comparado sai do Loop
        If dtData >= dtDataMaior Then Exit Do
        
        'Verifica se a Data é um feriado
        lErro = Comando_Executar(lComando, "SELECT Data FROM Feriados WHERE Data = ? AND FilialEmpresa = ? ", dtData, dtData, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 26210
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26211
        
        'Se não for conta a Data como sendo mais um Dia útil
        If lErro = AD_SQL_SEM_DADOS Then iDias = iDias + 1

        'incrementa-se a data
        dtData = dtData + 1
            
    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Dias_Uteis = SUCESSO
    
    Exit Function
    
Erro_Dias_Uteis:

    Dias_Uteis = Err
    
    Select Case Err
    
        Case 26209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 26210, 26211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FERIADOS", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142558)
            
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Estado_Seleciona(ByVal objEstado As Object) As Long
'Seleciona Estado da FilialEmpresa na(s) Combo(s) passada(s)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objEndereco As New ClassEndereco
Dim iSigla As Integer
Dim iIndice As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_Estado_Seleciona
    
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    
    'Lê a Filial Empresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then Error 19148

    'Se não encontrou a Filial da Empresa --> Erro
    If lErro <> SUCESSO Then Error 19149

    Set objEndereco = objFilialEmpresa.objEndereco

    If objEndereco.sSiglaEstado <> "" Then

        iSigla = Len(Trim(objEndereco.sSiglaEstado))
        
        'Se controle for ControlArray
        If TypeName(objEstado) = "Object" Then
        
            For iIndice = objEstado.LBound To objEstado.UBound
                'Seleciona o Estado "default" p/ a Filial se existir
                For iIndice2 = 0 To objEstado(iIndice).ListCount - 1
                    If right(objEstado(iIndice).List(iIndice2), iSigla) = objEndereco.sSiglaEstado Then
                        objEstado(iIndice).ListIndex = iIndice2
                        Exit For
                    End If
                Next
            Next

        'Se Controle for ComboBox
        ElseIf TypeName(objEstado) = "ComboBox" Then
                
            'Seleciona o Estado "default" p/ a Filial se existir
            For iIndice2 = 0 To objEstado.ListCount - 1
                If right(objEstado.List(iIndice2), iSigla) = objEndereco.sSiglaEstado Then
                    objEstado.ListIndex = iIndice2
                    Exit For
                End If
            Next
        
        End If
    
    End If

    Estado_Seleciona = SUCESSO

    Exit Function

Erro_Estado_Seleciona:

    Estado_Seleciona = Err

    Select Case Err

        Case 19148  'Tratado na rotina chamada

        Case 19149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142559)

    End Select
    
    Exit Function

End Function

Function Produto_Le_Niveis0e1_Estoque(ByVal colProduto As Collection) As Long
'le todos os produtos de nível 0 e 1 e coloca-os em colProduto

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto
Dim objProduto As ClassProduto

On Error GoTo Erro_Produto_Le_Niveis0e1_Estoque

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58750

    'Inicializa as variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sDescricao = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMEstoque = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMCompra = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMVenda = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSubstituto1 = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSubstituto2 = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNomeReduzido = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sModelo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCodigoBarras = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCor = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sIPICodDIPI = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sIPICodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sObsFisica = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabil = String(STRING_BUFFER_MAX_TEXTO, 0)


        'Lê os Produto de Nivel 1 e 2
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil FROM Produtos WHERE (ControleEstoque<>? OR Gerencial=?) AND Ativo=? AND (Nivel = 1 OR Nivel = 2) ORDER BY Codigo", .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, PRODUTO_CONTROLE_SEM_ESTOQUE, GERENCIAL, Ativo)
        If lErro <> AD_SQL_SUCESSO Then Error 58751

        'le o primeiro produto
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58752

        If lErro = AD_SQL_SEM_DADOS Then Error 58753

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto

            'Coloca os dados lidos no objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil

            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58754

        Loop

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Produto_Le_Niveis0e1_Estoque = SUCESSO

    Exit Function

Erro_Produto_Le_Niveis0e1_Estoque:

    Produto_Le_Niveis0e1_Estoque = Err

    Select Case Err

        Case 58750
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 58751, 58752, 58754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case 58753 'Não encontrou

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142560)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Netos_Estoque(ByVal sProduto As String, ByVal colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim lComando As Long
Dim lErro As Long
Dim objProduto As ClassProduto
Dim tProduto As typeProduto
Dim sProduto1 As String
Dim sProdutoEnxuto As String
Dim iNivel As Integer

On Error GoTo Erro_Produto_Le_Netos_Estoque

    'Inicializa as Variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sDescricao = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMEstoque = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMCompra = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUMVenda = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSubstituto1 = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSubstituto2 = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNomeReduzido = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sModelo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCodigoBarras = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCor = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sIPICodDIPI = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sIPICodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sObsFisica = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabil = String(STRING_BUFFER_MAX_TEXTO, 0)

        lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 58755

        lErro = Mascara_Produto_ObterNivel(sProduto, iNivel)
        If lErro <> SUCESSO Then Error 58756

        sProduto1 = sProdutoEnxuto + String(STRING_PRODUTO - Len(sProdutoEnxuto), "Z")

        lComando = Comando_Abrir()
        If lComando = 0 Then Error 58757

        'Lê produtos ordenados por Codigo (apenas os atributos que interessam)
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil FROM ProdutosProdutosFilial WHERE (ControleEstoque<>? OR Gerencial=?) AND Ativo=? AND Nivel = ? AND Codigo > ? AND Codigo < ? ORDER BY Codigo", .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, PRODUTO_CONTROLE_SEM_ESTOQUE, GERENCIAL, Ativo, iNivel + 2, sProduto, sProduto1)
        If lErro <> AD_SQL_SUCESSO Then Error 58758

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58759

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto

            'Guarda so dados lidos em objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil

            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58760

        Loop

    End With

    Call Comando_Fechar(lComando)

    Produto_Le_Netos_Estoque = SUCESSO

    Exit Function

Erro_Produto_Le_Netos_Estoque:

    Produto_Le_Netos_Estoque = Err

    Select Case Err

        Case 58755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, sProduto)

        Case 58756
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, sProduto)

        Case 58757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 58758, 58759, 58760
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142561)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Relatorios_Le_Outros(ByVal colRelatorio As Collection, ByVal sGrupo As String, ByVal sModulo As String) As Long
'retorna os relatorios referentes ao grupo e modulo que o usuario tem acesso e devem aparecer no dialogo de "outros"
'os codigos dos relatorios sao retornados em colRelatorio e o grupo e modulo sao passados como parametros

Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim lComando As Long
Dim sCodRel As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Relatorios_Le_Outros
       
    'abre conexao com o BD
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 59089
    
    sCodRel = String(STRING_CODIGO_RELATORIO, 0)
    
    'seleciona os relatorios que o usuario do grupo tem acesso para o modulo corrente
    lErro = Comando_Executar(lComando, "SELECT Relatorios.CodRel, Relatorios.FilialEmpresa FROM Relatorios , GrupoRelatorios , RelatoriosModulo WHERE Relatorios.CodRel = GrupoRelatorios.CodRel AND Relatorios.CodRel = RelatoriosModulo.CodRel AND GrupoRelatorios.TipoDeAcesso = 1 AND GrupoRelatorios.CodGrupo = ? AND RelatoriosModulo.Modulo = ? AND OutrosRelatorios = 1", sCodRel, iFilialEmpresa, sGrupo, sModulo)
    If lErro <> AD_SQL_SUCESSO Then Error 59090
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59091
        
    Do While lErro = AD_SQL_SUCESSO
                    
        'Se tiver como Empresa Toda
        If giFilialEmpresa = EMPRESA_TODA Then
            'Só não entra na colecao quem é de Somente Filial
            If iFilialEmpresa <> FILIAL_EMPRESA Then
                'Adiciona à Coleção
                colRelatorio.Add (sCodRel)
            End If
        Else 'se tiver como Filial
            'Entra Filial e Ambas menos FilialEmpresaToda
            If iFilialEmpresa <> FILIAL_EMPRESA_TODA Then
                'Adiciona à Coleção
                colRelatorio.Add (sCodRel)
            End If
        End If
                    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59092
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    Relatorios_Le_Outros = SUCESSO
    
    Exit Function

Erro_Relatorios_Le_Outros:

    Relatorios_Le_Outros = Err
    
    Select Case Err
                
        Case 59089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 59090, 59091, 59092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142562)
        
    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Function Campos_Le_Todos(ByVal colCampos As Collection) As Long
'le todos os campos e carrega-os na coleção passada como parametro

Dim lComando As Long
Dim lErro As Long
Dim tCampos As typeCampos
Dim objCampo As New AdmCampos
    
On Error GoTo Erro_Campos_Le_Todos

    Set colCampos = New Collection

    tCampos.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)
    tCampos.sFormatacao = String(STRING_FORMATACAO_CAMPO, 0)
    tCampos.sNome = String(STRING_NOME_CAMPO, 0)
    tCampos.sNomeArq = String(STRING_NOME_TABELA, 0)
    tCampos.sTituloEntradaDados = String(STRING_TITULO_ENTRADA_DADOS_CAMPO, 0)
    tCampos.sTituloGrid = String(STRING_TITULO_GRID_CAMPO, 0)
    tCampos.sValDefault = String(STRING_VALOR_DEFAULT_CAMPO, 0)
    tCampos.sValidacao = String(STRING_VALIDACAO_CAMPO, 0)
    
    lComando = 0
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 55984

    lErro = Comando_Executar(lComando, "SELECT NomeArq, Nome, Descricao, Obrigatorio, Imexivel, Ativo, ValDefault, Validacao, Formatacao, Tipo, Tamanho, Precisao, Decimais, TamExibicao, TituloEntradaDados, TituloGrid, Subtipo, Alinhamento FROM Campos", tCampos.sNomeArq, tCampos.sNome, tCampos.sDescricao, tCampos.iObrigatorio, tCampos.iImexivel, tCampos.iAtivo, tCampos.sValDefault, tCampos.sValidacao, tCampos.sFormatacao, tCampos.iTipo, tCampos.iTamanho, tCampos.iPrecisao, tCampos.iDecimais, tCampos.iTamExibicao, tCampos.sTituloEntradaDados, tCampos.sTituloGrid, tCampos.iSubTipo, tCampos.iAlinhamento)
    If lErro <> AD_SQL_SUCESSO Then Error 55985
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55986
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objCampo = New AdmCampos
    
        objCampo.sNome = tCampos.sNome
        objCampo.sNomeArq = tCampos.sNomeArq
        objCampo.iAtivo = tCampos.iAtivo
        objCampo.iDecimais = tCampos.iDecimais
        objCampo.iImexivel = tCampos.iImexivel
        objCampo.iObrigatorio = tCampos.iObrigatorio
        objCampo.iPrecisao = tCampos.iPrecisao
        objCampo.iTamanho = tCampos.iTamanho
        objCampo.iTamExibicao = tCampos.iTamExibicao
        objCampo.iTipo = tCampos.iTipo
        objCampo.sDescricao = tCampos.sDescricao
        objCampo.sFormatacao = tCampos.sFormatacao
        objCampo.sTituloEntradaDados = tCampos.sTituloEntradaDados
        objCampo.sTituloGrid = tCampos.sTituloGrid
        objCampo.sValDefault = tCampos.sValDefault
        objCampo.sValidacao = tCampos.sValidacao
        objCampo.iSubTipo = tCampos.iSubTipo
        objCampo.iAlinhamento = tCampos.iAlinhamento
        
        colCampos.Add objCampo, objCampo.sNomeArq + objCampo.sNome
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55989
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    Campos_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_Campos_Le_Todos:

    Campos_Le_Todos = Err

    Select Case Err
    
        Case 55984
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 55985, 55986, 55989
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOS", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142563)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Retorna_ColFiliais() As Long

Dim lErro As Long

On Error GoTo Erro_Retorna_ColFiliais

    Set gcolFiliais = New Collection
    
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, gcolFiliais)
    If lErro <> SUCESSO Then Error 44945

    Retorna_ColFiliais = SUCESSO
    
    Exit Function

Erro_Retorna_ColFiliais:

    Retorna_ColFiliais = Err

    Select Case Err
    
        Case 44945
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142564)

    End Select

    Exit Function

End Function

Function BrowseIndiceUsuario_Le(ByVal sNomeTela As String, ByVal sCodUsuario As String, ByVal colBrowseIndice As Collection) As Long
'le todos os indices criados pelo usuario referentes a tela especificada e coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim tBrowseIndice As typeBrowseIndice
Dim objBrowseIndice As AdmBrowseIndice
    
On Error GoTo Erro_BrowseIndiceUsuario_Le

    tBrowseIndice.sNomeTela = String(STRING_NOME_TELA, 0)
    tBrowseIndice.sNomeIndice = String(STRING_NOME_INDICE, 0)
    tBrowseIndice.sOrdenacaoSQL = String(STRING_ORDENACAO_SQL, 0)
    tBrowseIndice.sSelecaoSQL = String(STRING_SELECAO_SQL, 0)

    lComando = 0
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 60816

    lErro = Comando_Executar(lComando, "SELECT Indice, NomeIndice, OrdenacaoSQL, SelecaoSQL FROM BrowseIndiceUsuario WHERE NomeTela=? AND CodUsuario = ?", tBrowseIndice.iIndice, tBrowseIndice.sNomeIndice, tBrowseIndice.sOrdenacaoSQL, tBrowseIndice.sSelecaoSQL, sNomeTela, sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 60817
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60818
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objBrowseIndice = New AdmBrowseIndice
    
        objBrowseIndice.sNomeTela = sNomeTela
        objBrowseIndice.iIndice = tBrowseIndice.iIndice
        objBrowseIndice.sNomeIndice = tBrowseIndice.sNomeIndice
        objBrowseIndice.sOrdenacaoSQL = tBrowseIndice.sOrdenacaoSQL
        objBrowseIndice.sSelecaoSQL = tBrowseIndice.sSelecaoSQL

        colBrowseIndice.Add objBrowseIndice
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60819
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    BrowseIndiceUsuario_Le = SUCESSO
    
    Exit Function
    
Erro_BrowseIndiceUsuario_Le:

    BrowseIndiceUsuario_Le = Err

    Select Case Err
    
        Case 60816
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 60817, 60818, 60819
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEINDICEUSUARIO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142565)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Niveis0e1_Venda_EmpresaToda(ByVal colProduto As Collection) As Long
'Lê todos os produtos que podem ser vendidos para a Empresa Toda e que sejam de nível 0 e 1 e coloca-os em colProduto

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto
Dim objProduto As ClassProduto

On Error GoTo Erro_Produto_Le_Niveis0e1_Venda_EmpresaToda

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61171

    'Inicializa as variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)

        'Lê os Produto de Nivel 1 e 2
        lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza FROM Produtos WHERE (Faturamento=? OR Gerencial=?) AND Ativo=? AND (Nivel = 1 OR Nivel = 2) ORDER BY Codigo", _
            .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, PRODUTO_VENDAVEL, GERENCIAL, Ativo)
        If lErro <> AD_SQL_SUCESSO Then Error 61172

        'le o primeiro produto
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61173

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto
    
            'Coloca os dados lidos no objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
            objProduto.sContaContabilProducao = .sContaContabilProducao
            objProduto.dResiduo = .dResiduo
            objProduto.iNatureza = .iNatureza
            
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61174

        Loop

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Produto_Le_Niveis0e1_Venda_EmpresaToda = SUCESSO

    Exit Function

Erro_Produto_Le_Niveis0e1_Venda_EmpresaToda:

    Produto_Le_Niveis0e1_Venda_EmpresaToda = Err

    Select Case Err

        Case 61171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61172, 61173, 61174
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142566)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Netos_Venda_EmpresaToda(ByVal sProduto As String, ByVal colProduto As Collection) As Long
'le todos os netos de sProduto e coloca-os em colProduto

Dim lComando As Long
Dim lErro As Long
Dim objProduto As ClassProduto
Dim tProduto As typeProduto
Dim sProduto1 As String
Dim sProdutoEnxuto As String
Dim iNivel As Integer

On Error GoTo Erro_Produto_Le_Netos_Venda_EmpresaToda

    'Inicializa as Variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 61175

        lErro = Mascara_Produto_ObterNivel(sProduto, iNivel)
        If lErro <> SUCESSO Then Error 61176

        sProduto1 = sProdutoEnxuto + String(STRING_PRODUTO - Len(sProdutoEnxuto), "Z")

        lComando = Comando_Abrir()
        If lComando = 0 Then Error 61177
    
        'Lê produtos ordenados por Codigo (apenas os atributos que interessam)
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza FROM Produtos WHERE (Faturamento =? OR Gerencial=?) AND Ativo=? AND Nivel = ? AND Codigo > ? AND Codigo < ? ORDER BY Codigo", _
                .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, PRODUTO_VENDAVEL, GERENCIAL, Ativo, iNivel + 2, sProduto, sProduto1)
        If lErro <> AD_SQL_SUCESSO Then Error 61178

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61179

        Do While lErro = AD_SQL_SUCESSO

            Set objProduto = New ClassProduto

            'Guarda so dados lidos em objProduto
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iGerencial = .iGerencial
            objProduto.iAtivo = .iAtivo
            objProduto.iFaturamento = .iFaturamento
            objProduto.iClasseUM = .iClasseUM
            objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
            objProduto.sSiglaUMCompra = .sSiglaUMCompra
            objProduto.sSiglaUMVenda = .sSiglaUMVenda
            objProduto.sSubstituto1 = .sSubstituto1
            objProduto.sSubstituto2 = .sSubstituto2
            objProduto.iTipo = .iTipo
            objProduto.sNomeReduzido = .sNomeReduzido
            objProduto.sModelo = .sModelo
            objProduto.iNivel = .iNivel
            objProduto.iPrazoValidade = .iPrazoValidade
            objProduto.sCodigoBarras = .sCodigoBarras
            objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
            objProduto.dPesoLiq = .dPesoLiq
            objProduto.dPesoBruto = .dPesoBruto
            objProduto.dComprimento = .dComprimento
            objProduto.dEspessura = .dEspessura
            objProduto.dLargura = .dLargura
            objProduto.sCor = .sCor
            objProduto.sObsFisica = .sObsFisica
            objProduto.iCompras = .iCompras
            objProduto.iPCP = .iPCP
            objProduto.iKitBasico = .iKitBasico
            objProduto.iKitInt = .iKitInt
            objProduto.dIPIAliquota = .dIPIAliquota
            objProduto.sIPICodigo = .sIPICodigo
            objProduto.sIPICodDIPI = .sIPICodDIPI
            objProduto.iControleEstoque = .iControleEstoque
            objProduto.iApropriacaoCusto = .iApropriacaoCusto
            objProduto.sContaContabil = .sContaContabil
            objProduto.sContaContabilProducao = .sContaContabilProducao
            objProduto.dResiduo = .dResiduo
            objProduto.iNatureza = .iNatureza
    
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61180

        Loop

    End With

    Call Comando_Fechar(lComando)
    
    Produto_Le_Netos_Venda_EmpresaToda = SUCESSO

    Exit Function

Erro_Produto_Le_Netos_Venda_EmpresaToda:

    Produto_Le_Netos_Venda_EmpresaToda = Err
    
    Select Case Err

        Case 61175
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, sProduto)

        Case 61176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, sProduto)

        Case 61177
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61178, 61179, 61180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142567)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function EdicaoTela_Le() As Long
'Le as propriedades das Telas no dicionário de Dados

Dim lErro As Long
Dim tEdicaoTela As typeEdicaoTela
Dim objEdicaoTela_Tela As New ClassEdicaoTela_Tela
Dim sNomeTelaAnterior As String
Dim objEdicaoTela_Controle As New ClassEdicaoTela_Controle
Dim lComando As Long
Dim objUsu As New ClassUsuarios

On Error GoTo Erro_EdicaoTela_Le
    
    'abre o comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 64017

    'Inserido por Wagner
    '########################
    objUsu.sCodUsuario = gsUsuario
    lErro = Usuarios_Le(objUsu)
    If lErro <> SUCESSO Then gError 129297
    '#########################

    sNomeTelaAnterior = ""
    tEdicaoTela.sNomeTela = String(STRING_NOME_TELA, 0)
    tEdicaoTela.sNomeControle = String(STRING_NOME_CONTROLE, 0)
    tEdicaoTela.sContainer = String(STRING_EDICAOTELA_CONTAINER, 0)
    tEdicaoTela.sTitulo = String(STRING_EDICAOTELA_TITULO, 0)
    tEdicaoTela.sGrupoUsuarios = String(50, 0) 'Inserido por Wagner
    
    'Lê os controles
    lErro = Comando_Executar(lComando, "SELECT NomeTela, NomeControle, Titulo, Visivel, Largura, Altura, Esquerda, Topo, TabStop, Container, IndiceContainer, TabIndex, GrupoUsuarios, zOrder, Habilitado FROM EdicaoTela WHERE GrupoUsuarios = ? ORDER BY NomeTela", tEdicaoTela.sNomeTela, tEdicaoTela.sNomeControle, tEdicaoTela.sTitulo, tEdicaoTela.iVisivel, tEdicaoTela.iLargura, tEdicaoTela.iAltura, tEdicaoTela.iEsquerda, tEdicaoTela.iTopo, tEdicaoTela.iTabStop, tEdicaoTela.sContainer, tEdicaoTela.iIndiceContainer, tEdicaoTela.iTabIndex, tEdicaoTela.sGrupoUsuarios, tEdicaoTela.izOrder, tEdicaoTela.iHabilitado, objUsu.sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then gError 64018
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64019
        
    Set gcolEdicaoTela = New Collection
    
    Do While lErro = AD_SQL_SUCESSO
                
        'Seta um novo controle
        Set objEdicaoTela_Controle = New ClassEdicaoTela_Controle
        
        'Passa os dados para o obj
        objEdicaoTela_Controle.iAltura = tEdicaoTela.iAltura
        objEdicaoTela_Controle.iEsquerda = tEdicaoTela.iEsquerda
        objEdicaoTela_Controle.iLargura = tEdicaoTela.iLargura
        objEdicaoTela_Controle.iTabStop = tEdicaoTela.iTabStop
        objEdicaoTela_Controle.iTopo = tEdicaoTela.iTopo
        objEdicaoTela_Controle.iVisivel = tEdicaoTela.iVisivel
        objEdicaoTela_Controle.sContainer = tEdicaoTela.sContainer
        objEdicaoTela_Controle.iIndiceContainer = tEdicaoTela.iIndiceContainer
        objEdicaoTela_Controle.sNomeControle = tEdicaoTela.sNomeControle
        objEdicaoTela_Controle.sNomeTela = tEdicaoTela.sNomeTela
        objEdicaoTela_Controle.sTitulo = tEdicaoTela.sTitulo
        objEdicaoTela_Controle.iTabIndex = tEdicaoTela.iTabIndex
        objEdicaoTela_Controle.sGrupoUsuarios = tEdicaoTela.sGrupoUsuarios 'Inserido por Wagner
        objEdicaoTela_Controle.izOrder = tEdicaoTela.izOrder 'Inserido por Wagner
        objEdicaoTela_Controle.iHabilitado = tEdicaoTela.iHabilitado 'Inserido por Wagner
        
        'Se sNomeTelaAnterior <> Tela Atual então é uma Nova Tela
        If sNomeTelaAnterior <> tEdicaoTela.sNomeTela Then
            
            'Se está Vazia então não existe Tela Anterior Lida
            If Trim(sNomeTelaAnterior) = "" Then
                'Cria uma Nova Tela e adiciona o controle
                Set objEdicaoTela_Tela = New ClassEdicaoTela_Tela
                objEdicaoTela_Tela.sNomeTela = tEdicaoTela.sNomeTela
                sNomeTelaAnterior = tEdicaoTela.sNomeTela
                Set objEdicaoTela_Tela.colEdicaoTela_Controle = New Collection
                objEdicaoTela_Tela.colEdicaoTela_Controle.Add objEdicaoTela_Controle, objEdicaoTela_Controle.sNomeControle
                gcolEdicaoTela.Add objEdicaoTela_Tela, objEdicaoTela_Tela.sNomeTela
            Else
                'Adiciona a Tela Anterior, cria-se uma Tela Nova e Adiciona o Controle
                Set objEdicaoTela_Tela = New ClassEdicaoTela_Tela
                Set objEdicaoTela_Tela.colEdicaoTela_Controle = New Collection
                objEdicaoTela_Tela.sNomeTela = tEdicaoTela.sNomeTela
                sNomeTelaAnterior = tEdicaoTela.sNomeTela
                objEdicaoTela_Tela.colEdicaoTela_Controle.Add objEdicaoTela_Controle, objEdicaoTela_Controle.sNomeControle
                gcolEdicaoTela.Add objEdicaoTela_Tela, objEdicaoTela_Tela.sNomeTela
            End If
            
        Else
            
            If objEdicaoTela_Tela.colEdicaoTela_Controle Is Nothing Then Set objEdicaoTela_Tela.colEdicaoTela_Controle = New Collection
               
            'Adiciona o controle
            objEdicaoTela_Tela.colEdicaoTela_Controle.Add objEdicaoTela_Controle, objEdicaoTela_Controle.sNomeControle
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64020
        
    Loop
        
    'Le os Tab Index
    lErro = TabIndex_Le()
    If lErro <> SUCESSO Then gError 64034
        
    Call Comando_Fechar(lComando)
    
    EdicaoTela_Le = SUCESSO
    
    Exit Function
    
Erro_EdicaoTela_Le:

    EdicaoTela_Le = gErr

    Select Case gErr
    
        Case 64017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 64018, 64019, 64020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EDICAOTELA", gErr)
        
        Case 64034, 129297
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142568)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TabIndex_Le() As Long
'Lê todos os Tab Index das Telas do Dicionário de dados

Dim lErro As Long
Dim sNomeTela As String
Dim sTabIndex1 As String
Dim sTabIndex2 As String
Dim sTabIndex3 As String
Dim sTabIndex4 As String
Dim sGrupoUsu As String
Dim objEdicaoTela_Tela As New ClassEdicaoTela_Tela
Dim objEdicaoTela_Controle As New ClassEdicaoTela_Controle
Dim lComando As Long
Dim bEncontrou As Boolean
Dim objUsu As New ClassUsuarios

On Error GoTo Erro_TabIndex_Le

    'abre o comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 64030
        
    sNomeTela = String(STRING_NOME_TELA, 0)
    sTabIndex1 = String(STRING_BUFFER_MAX_TEXTO, 0)
    sTabIndex2 = String(STRING_BUFFER_MAX_TEXTO, 0)
    sTabIndex3 = String(STRING_BUFFER_MAX_TEXTO, 0)
    sTabIndex4 = String(STRING_BUFFER_MAX_TEXTO, 0)
    sGrupoUsu = String(STRING_GRUPO, 0)
    
    'Inserido por Wagner
    '########################
    objUsu.sCodUsuario = gsUsuario
    lErro = Usuarios_Le(objUsu)
    If lErro <> SUCESSO Then gError 129297
    '#########################
    
    'Lê as Telas e Tabs Index
    lErro = Comando_Executar(lComando, "SELECT NomeTela, TabIndex1, TabIndex2, TabIndex3, TabIndex4, GrupoUsuarios FROM TabIndex WHERE GrupoUsuarios = ? ORDER BY NomeTela", sNomeTela, sTabIndex1, sTabIndex2, sTabIndex3, sTabIndex4, sGrupoUsu, objUsu.sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then gError 64031
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64032
        
    Do While lErro = AD_SQL_SUCESSO
                
        bEncontrou = False
                
        For Each objEdicaoTela_Tela In gcolEdicaoTela
            
            If objEdicaoTela_Tela.sNomeTela = sNomeTela Then
                objEdicaoTela_Tela.sTabIndex = sTabIndex1 & sTabIndex2 & sTabIndex3 & sTabIndex4
                bEncontrou = True
                Exit For
            End If
        Next
        
        If bEncontrou = False Then
            Set objEdicaoTela_Tela = New ClassEdicaoTela_Tela
            Set objEdicaoTela_Tela.colEdicaoTela_Controle = New Collection
            objEdicaoTela_Tela.sNomeTela = sNomeTela
            objEdicaoTela_Tela.sTabIndex = sTabIndex1 & sTabIndex2 & sTabIndex3 & sTabIndex4
            objEdicaoTela_Tela.sGrupoUsuarios = sGrupoUsu 'Inserido por Wagner
            gcolEdicaoTela.Add objEdicaoTela_Tela, objEdicaoTela_Tela.sNomeTela
        End If
                    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64033
        
    Loop
                
    Call Comando_Fechar(lComando)

    TabIndex_Le = SUCESSO
    
    Exit Function
    
Erro_TabIndex_Le:

    TabIndex_Le = gErr

    Select Case gErr
    
        Case 64030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 64031, 64032, 64033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABINDEX", gErr)
            
        Case 129297
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142569)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Function RelOpcoes_Le_Parametros(ByVal sCodRel As String, ByVal sNome As String, colParametros As Collection) As Long
'Lê os Parametros do Relatorio para a Opção

Dim lComando As Long
Dim lErro As Long
Dim sParametro As String
Dim sParamValor As String
Dim objParametro As AdmRelParametro

On Error GoTo Erro_RelOpcoes_Le_Parametros
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 61267
    
    Set colParametros = New Collection
    
    sParametro = String(STRING_REL_PARAM, 0)
    sParamValor = String(STRING_REL_PARAM_VALOR, 0)
    
    'le a tabela de opções de relatório
    lErro = Comando_Executar(lComando, "SELECT Parametro, ParamValor FROM RelatorioParametros WHERE CodRel = ? AND Nome = ?", sParametro, sParamValor, sCodRel, sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 61268

    'lê a primeira opção de relatório
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61269
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objParametro = New AdmRelParametro
        
        objParametro.sParametro = sParametro
        objParametro.sParamValor = sParamValor
        
        colParametros.Add objParametro
        
        'lê a primeira opção de relatório
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61270
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    RelOpcoes_Le_Parametros = SUCESSO

    Exit Function

Erro_RelOpcoes_Le_Parametros:

    RelOpcoes_Le_Parametros = Err

    Select Case Err

        Case 61267
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 61268, 61269, 61270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIOPARAMETROS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142570)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function RelTela_LeCampos(ByVal sCodRel As String, ByVal colCampos As ColRelTelaCampos) As Long
'carrega colecao com definicao dos campos que compoem um registro para relatorio tipo "printscreen"

Dim lErro As Long, lComando As Long
Dim tCampo As typeRelTelaCampo

On Error GoTo Erro_RelTela_LeCampos
        
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 41560
    
    tCampo.sNome = String(STRING_NOME_CAMPO, 0)
    lErro = Comando_Executar(lComando, "SELECT Sequencial, Nome, Tipo, Tamanho FROM RelTelaCampos WHERE CodRel = ? ORDER BY Sequencial", tCampo.iSequencial, tCampo.sNome, tCampo.iTipo, tCampo.iTamanho, sCodRel)
    If lErro <> AD_SQL_SUCESSO Then Error 41561
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41562
    
    Do While (lErro <> AD_SQL_SEM_DADOS)
        
        colCampos.Add sCodRel, tCampo.iSequencial, tCampo.sNome, tCampo.iTipo, tCampo.iTamanho
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41563
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    RelTela_LeCampos = SUCESSO
     
    Exit Function
    
Erro_RelTela_LeCampos:

    RelTela_LeCampos = Err
     
    Select Case Err
          
        Case 41560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 41561, 41562, 41563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELTELACAMPOS", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142571)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function MenuItem_Le_Titulo(ByVal sTitulo As String, ByVal objMenuItem As ClassMenuItens) As Long
'Lê item de menu a partir do nome da tela

Dim tMenuItens As typeMenuItens
Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_MenuItem_Le_Titulo

    'Iniciliza comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 25961

    tMenuItens.sNomeTela = String(STRING_NOME_TELA, 0)
    tMenuItens.sNomeControle = String(STRING_NOME_CONTROLE, 0)
    tMenuItens.sNomeControlePai = String(STRING_NOME_CONTROLE, 0)
    tMenuItens.sSiglaRotina = String(STRING_SIGLA_ROTINA, 0)
    tMenuItens.sTitulo = String(STRING_TITULO_MENU, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Identificador, Titulo, SiglaRotina, NomeTela, NomeControle, IndiceControle, NomeControlePai, IndiceControlePai FROM MenuItens WHERE Titulo = ? ", _
        tMenuItens.iIdentificador, tMenuItens.sTitulo, tMenuItens.sSiglaRotina, tMenuItens.sNomeTela, tMenuItens.sNomeControle, tMenuItens.iIndiceControle, tMenuItens.sNomeControlePai, tMenuItens.iIndiceControlePai, sTitulo)
    If lErro <> AD_SQL_SUCESSO Then Error 25962

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25963
    If lErro = AD_SQL_SEM_DADOS Then Error 25964

    objMenuItem.iIdentificador = tMenuItens.iIdentificador
    objMenuItem.sTitulo = tMenuItens.sTitulo
    objMenuItem.sSiglaRotina = tMenuItens.sSiglaRotina
    objMenuItem.sNomeTela = tMenuItens.sNomeTela
    objMenuItem.sNomeControle = tMenuItens.sNomeControle
    objMenuItem.iIndiceControle = tMenuItens.iIndiceControle
    objMenuItem.sNomeControlePai = tMenuItens.sNomeControlePai
    objMenuItem.iIndiceControlePai = tMenuItens.iIndiceControlePai

    Call Comando_Fechar(lComando)

    MenuItem_Le_Titulo = SUCESSO

    Exit Function

Erro_MenuItem_Le_Titulo:

    MenuItem_Le_Titulo = Err

    Select Case Err

        Case 25961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25962, 25963
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MENUITENS", Err)

        Case 25964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENUITEM_NAO_CADASTRADO", Err, sTitulo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142572)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le(ByVal objProduto As ClassProduto) As Long
'Lê os dados do produto, cujo codigo foi passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto

On Error GoTo Erro_Produto_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28027

    'Inicializa as variáveis do tipo String
    With tProduto
        
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilProducao = String(STRING_CONTA, 0)
        .sReferencia = String(STRING_PRODUTO_REFERENCIA, 0)
        .sFigura = String(STRING_NOME_ARQ_COMPLETO, 0)
        .sGrade = String(STRING_GRADE_CODIGO, 0)
        .sSerieProx = String(STRING_SERIE_RASTREAMENTO, 0)
        .sSiglaUMTrib = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sISSQN = String(STRING_ISSQN_CODIGO, 0)
        .sGenero = String(STRING_PRODUTOGENERO_CODIGO, 0)
        .sFCI = String(STRING_FCI, 0)
        .sNBS = String(STRING_NBS, 0)
        .sCEST = String(STRING_CEST, 0)
        .scProdANVISA = String(STRING_MAXIMO, 0)
                        
        'Lê o Produto
        lErro = Comando_Executar(lComando, "SELECT  Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao, ContaContabil, ContaContabilProducao, Residuo, Natureza ,CustoReposicao, OrigemMercadoria, TabelaPreco, PercentMaisQuantCotAnt, PercentMenosQuantCotAnt, ConsideraQuantCotAnt, TemFaixaReceb, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa,TempoProducao,Rastro,HorasMaquina,PesoEspecifico, Embalagem, Referencia, Figura, INSSPercBase, KitVendaComp, Grade, UsaBalanca, SerieProx, SerieParteNum,ExTIPI,Genero,ISSQN,SiglaUMTrib,ProdutoEspecifico,PrecoMaxConsumidor, Fabricante, IPIIncide, QtdeEmbBase, FatorAjuste, FCI, NBS, PercComissao, MetaComissao, CEST, ProdEmEscalaRelev, cProdANVISA " & _
        "FROM Produtos WHERE Codigo = ?", _
        .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .sContaContabil, .sContaContabilProducao, .dResiduo, .iNatureza, .dCustoReposicao, .iOrigemMercadoria, .iTabelaPreco, .dPercentMaisQuantCotAnt, .dPercentMenosQuantCotAnt, .iConsideraQuantCotAnt, .iTemFaixaReceb, .dPercentMaisReceb, .dPercentMenosReceb, .iRecebForaFaixa, .iTempoProducao, .iRastro, .lHorasMaquina, .dPesoEspecifico, .iEmbalagem, .sReferencia, .sFigura, .dINSSPercBase, .iKitVendaComp, .sGrade, .iUsaBalanca, .sSerieProx, .iSerieParteNum, .iExTIPI, .sGenero, .sISSQN, _
        .sSiglaUMTrib, .iProdutoEspecifico, .dPrecoMaxConsumidor, .lFabricante, .iIPIIncide, .dQtdeEmbBase, .dFatorAjuste, .sFCI, .sNBS, .dPercComissao, .dMetaComissao, .sCEST, .iProdEmEscalaRelev, .scProdANVISA, objProduto.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 28028

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28029

        'Se não encontrou o Produto --> Erro
        If lErro = AD_SQL_SEM_DADOS Then Error 28030

        'Coloca os dados lidos no objProduto
        objProduto.sDescricao = .sDescricao
        objProduto.iGerencial = .iGerencial
        objProduto.iAtivo = .iAtivo
        objProduto.iFaturamento = .iFaturamento
        objProduto.iClasseUM = .iClasseUM
        objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
        objProduto.sSiglaUMCompra = .sSiglaUMCompra
        objProduto.sSiglaUMVenda = .sSiglaUMVenda
        objProduto.sSubstituto1 = .sSubstituto1
        objProduto.sSubstituto2 = .sSubstituto2
        objProduto.iTipo = .iTipo
        objProduto.sNomeReduzido = .sNomeReduzido
        objProduto.sModelo = .sModelo
        objProduto.iNivel = .iNivel
        objProduto.iPrazoValidade = .iPrazoValidade
        objProduto.sCodigoBarras = .sCodigoBarras
        objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
        objProduto.dPesoLiq = .dPesoLiq
        objProduto.dPesoBruto = .dPesoBruto
        objProduto.dComprimento = .dComprimento
        objProduto.dEspessura = .dEspessura
        objProduto.dLargura = .dLargura
        objProduto.sCor = .sCor
        objProduto.sObsFisica = .sObsFisica
        objProduto.iCompras = .iCompras
        objProduto.iPCP = .iPCP
        objProduto.iKitBasico = .iKitBasico
        objProduto.iKitInt = .iKitInt
        objProduto.dIPIAliquota = .dIPIAliquota
        objProduto.sIPICodigo = .sIPICodigo
        objProduto.sIPICodDIPI = .sIPICodDIPI
        objProduto.iControleEstoque = .iControleEstoque
        objProduto.iApropriacaoCusto = .iApropriacaoCusto
        objProduto.sContaContabil = .sContaContabil
        objProduto.sContaContabilProducao = .sContaContabilProducao
        objProduto.dResiduo = .dResiduo
        objProduto.iNatureza = .iNatureza
        objProduto.dCustoReposicao = .dCustoReposicao
        objProduto.iOrigemMercadoria = .iOrigemMercadoria
        objProduto.iTabelaPreco = .iTabelaPreco
        objProduto.dPercentMaisQuantCotAnt = .dPercentMaisQuantCotAnt
        objProduto.dPercentMenosQuantCotAnt = .dPercentMenosQuantCotAnt
        objProduto.iConsideraQuantCotAnt = .iConsideraQuantCotAnt
        objProduto.iTemFaixaReceb = .iTemFaixaReceb
        objProduto.dPercentMaisReceb = .dPercentMaisReceb
        objProduto.dPercentMenosReceb = .dPercentMenosReceb
        objProduto.iRecebForaFaixa = .iRecebForaFaixa
        objProduto.iTempoProducao = .iTempoProducao
        objProduto.iRastro = .iRastro
        objProduto.lHorasMaquina = .lHorasMaquina
        objProduto.dPesoEspecifico = .dPesoEspecifico
        objProduto.iEmbalagem = .iEmbalagem
        objProduto.sReferencia = .sReferencia
        objProduto.sFigura = .sFigura
        objProduto.dINSSPercBase = .dINSSPercBase
        objProduto.iKitVendaComp = .iKitVendaComp
        objProduto.sGrade = .sGrade
        objProduto.iUsaBalanca = .iUsaBalanca
        objProduto.sSerieProx = .sSerieProx
        objProduto.iSerieParteNum = .iSerieParteNum
        
        objProduto.iExTIPI = .iExTIPI
        objProduto.sGenero = .sGenero
        objProduto.sISSQN = .sISSQN
        objProduto.sSiglaUMTrib = .sSiglaUMTrib
        objProduto.iProdutoEspecifico = .iProdutoEspecifico
        objProduto.dPrecoMaxConsumidor = .dPrecoMaxConsumidor
        objProduto.lFabricante = .lFabricante
        objProduto.iIPIIncide = .iIPIIncide
        objProduto.dQtdeEmbBase = .dQtdeEmbBase
        objProduto.dFatorAjuste = .dFatorAjuste
        objProduto.sFCI = .sFCI
        objProduto.sNBS = .sNBS
        
        objProduto.dPercComissao = .dPercComissao
        objProduto.dMetaComissao = .dMetaComissao
    
        objProduto.sCEST = .sCEST
        objProduto.iProdEmEscalaRelev = .iProdEmEscalaRelev
    
        objProduto.scProdANVISA = .scProdANVISA
    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Produto_Le = SUCESSO

    Exit Function

Erro_Produto_Le:

    Produto_Le = Err

    Select Case Err

        Case 28027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28028, 28029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", Err, objProduto.sCodigo)

        Case 28030
            'Produto não encontrado. Será tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142573)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function VerificaRegistro_ObjetosBD(ByVal objObjetosBD As ClassObjetoBD, avParametros() As Variant) As Long

Dim vQuantReg As Variant
Dim sSelect As String
Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer

On Error GoTo Erro_VerificaRegistro_ObjetosBD

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80451

    'Monta instrução SQL
    sSelect = "SELECT COUNT(*) FROM " & objObjetosBD.sNomeArquivo & " WHERE " & objObjetosBD.sSelecaoSQL

    'Monta SQL Dinâmico
    lErro = Comando_PrepararInt1(lComando, sSelect)
    If lErro <> AD_SQL_SUCESSO Then gError 80444
    
    vQuantReg = 0
    lErro = Comando_BindVarInt1(lComando, vQuantReg)
    If (lErro <> AD_SQL_SUCESSO) Then gError 80447
    
    For iIndice = LBound(avParametros) To UBound(avParametros)
        lErro = Comando_BindVarInt1(lComando, avParametros(iIndice))
        If (lErro <> AD_SQL_SUCESSO) Then gError 80448
    Next

    lErro = Comando_ExecutarInt1(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 80449
    
    'Busca primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80445
    
    'Se não encontrou fornece número de erro
    If lErro = AD_SQL_SEM_DADOS Or vQuantReg = 0 Then gError 80446
    
    Call Comando_Fechar(lComando)
    
    VerificaRegistro_ObjetosBD = SUCESSO

    Exit Function

Erro_VerificaRegistro_ObjetosBD:

    VerificaRegistro_ObjetosBD = gErr

    Select Case gErr

        Case 80446

        Case 80444, 80445, 80447, 80448, 80449
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, objObjetosBD.sNomeArquivo)
    
        Case 80452
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142574)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function ObjetosBD_Le(ByVal objObjetoBD As ClassObjetoBD) As Long

Dim lErro As Long
Dim lComando As Long
Dim tObjetosBD As typeObjetosBd

On Error GoTo Erro_ObjetosBD_Le

    'Abre Comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 80428
        
    'Iniciliza os campos Strings
    tObjetosBD.sNomeArquivo = String(STRING_NOME_ARQUIVO, 0)
    tObjetosBD.sClasseObjeto = String(STRING_CLASSE_OBJETO, 0)
    tObjetosBD.sSelecaoSQL = String(STRING_SELECAO_SQL, 0)
    tObjetosBD.sNomeObjetoMSG = String(STRING_NOME_OBJETO_MSG, 0)
    
    'Executa leitura na tabela ObjetosBD
    lErro = Comando_Executar(lComando, "SELECT NomeArquivo, Tipo, SelecaoSQL, AvisaSobrePosicao, NomeObjetoMsg FROM ObjetosBD WHERE ClasseObjeto = ?", tObjetosBD.sNomeArquivo, tObjetosBD.iTipo, tObjetosBD.sSelecaoSQL, tObjetosBD.iAvisaSobrePosicao, tObjetosBD.sNomeObjetoMSG, objObjetoBD.sClasseObjeto)
    If lErro <> AD_SQL_SUCESSO Then gError 80429

    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80430
        
    'Atribui valores do type para o Obj
    objObjetoBD.iAvisaSobrePosicao = tObjetosBD.iAvisaSobrePosicao
    objObjetoBD.iTipo = tObjetosBD.iTipo
    objObjetoBD.sNomeArquivo = tObjetosBD.sNomeArquivo
    objObjetoBD.sNomeObjetoMSG = tObjetosBD.sNomeObjetoMSG
    objObjetoBD.sSelecaoSQL = tObjetosBD.sSelecaoSQL
        
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    ObjetosBD_Le = SUCESSO

    Exit Function

Erro_ObjetosBD_Le:

    ObjetosBD_Le = gErr
    
    Select Case gErr
    
        Case 80428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error)
    
        Case 80429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OBJETOBD", gErr, Error)
        
        Case 80430
            'Erro Tratado na rotina chamadora
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142575)
        
    End Select
        
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function BrowseParamSelecao_Le(ByVal sNomeTela As String, ByVal colBrowseParamSelecao As Collection) As Long
'le os valores dos parametros de selecao relacionados a sNomeTela coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim tBrowseParamSelecao As typeBrowseParamSelecao
    
On Error GoTo Erro_BrowseParamSelecao_Le

    tBrowseParamSelecao.sNomeTela = String(STRING_NOME_TELA, 0)
    tBrowseParamSelecao.sClasse = String(NOME_CLASSE, 0)
    tBrowseParamSelecao.sProjeto = String(NOME_PROJETO, 0)
    tBrowseParamSelecao.sProperty = String(STRING_NOME_CAMPO, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 89978

    lErro = Comando_Executar(lComando, "SELECT Projeto, Classe, Property FROM BrowseParamSelecao WHERE NomeTela=? ORDER BY Ordem", tBrowseParamSelecao.sProjeto, tBrowseParamSelecao.sClasse, tBrowseParamSelecao.sProperty, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 89979
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89980
    
    Do While lErro = SUCESSO
    
        colBrowseParamSelecao.Add CallByName(CreateObject(tBrowseParamSelecao.sProjeto & "." & tBrowseParamSelecao.sClasse), tBrowseParamSelecao.sProperty, VbGet)
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89981
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    BrowseParamSelecao_Le = SUCESSO
    
    Exit Function
    
Erro_BrowseParamSelecao_Le:

    BrowseParamSelecao_Le = gErr

    Select Case gErr
    
        Case 89978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89979, 89980, 89981
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEPARAMSELECAO", gErr, sNomeTela)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142576)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoCategoria_Le(ByVal sProduto As String, ByVal sCategoria As String, sItem As String) As Long
'retorna o item do produto/categoria

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_ProdutoCategoria_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92402
    
    sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Item FROM ProdutoCategoria WHERE Produto = ? AND Categoria = ?", sItem, sProduto, sCategoria)
    If lErro <> AD_SQL_SUCESSO Then gError 92403
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92404
        
    If lErro = AD_SQL_SEM_DADOS Then gError 92405
    
    Call Comando_Fechar(lComando)
    
    ProdutoCategoria_Le = SUCESSO

    Exit Function

Erro_ProdutoCategoria_Le:

    ProdutoCategoria_Le = gErr

    Select Case gErr

        Case 92402
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 92403, 92404
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTO", Err, sProduto, sCategoria)
        
        Case 92405
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142577)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Browse_Executa_SQL(ByVal objBrowse As AdmBrowse) As Long
'Executa o comando SQL guarda em objBrowse.sSQL
'objBrowse RECEBE(Input) as informações do browse que serão necessárias para execução do select
'objBrowse RETORNA(Output) os registros retornados pela leitura dentro da coleção colRegistros

Dim iIndice As Integer
Dim lComando As Long
Dim lErro As Long
Dim avCampo(1 To BROWSER_NUM_MAX_CAMPOS) As Variant
Dim avRegistro() As Variant

On Error GoTo Erro_Browse_Executa_SQL

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102088
    
    'Prepara as variáveis que serão usadas no select para receber os valores lidos
    lErro = objBrowse.Browse_Reinicializa_avCampo2(objBrowse, avCampo)
    If lErro <> SUCESSO Then gError 102089
    
    'Executa o select
    lErro = CF("Browse_Prepara_Comando_SQL", objBrowse.sSQL, objBrowse.colBrowseUsuarioCampo.Count, lComando, avCampo, objBrowse.colSelecao, BROWSE_EXECUTA_BIND)
    If lErro <> SUCESSO Then gError 102090
    
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102091
    
    'Se não encontrou=> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102092
    
    'cria um array com o tamanho de um registro visível no browser
    ReDim avRegistro(1 To objBrowse.colBrowseUsuarioCampo.Count)
    
    'Enquanto houver registros
    Do While lErro = AD_SQL_SUCESSO
    
        'Para cada campo acessível ao usuário no browse, ou seja, cada campo que está sendo lido no BD
        For iIndice = 1 To objBrowse.colBrowseUsuarioCampo.Count
        
            'guarda cada valor dentro do array, fingindo q o avRegistro seria um obj e q cada elemento seu, uma property desse obj
            avRegistro(iIndice) = avCampo(iIndice)
        
        Next
        
        'adiciona o array à coleção de registros
        objBrowse.colRegistros.Add avRegistro
        
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102093
    
    Loop

    Call Comando_Fechar(lComando)
    
    Browse_Executa_SQL = SUCESSO
    
    Exit Function
    
Erro_Browse_Executa_SQL:

    Browse_Executa_SQL = gErr

    Select Case gErr
    
        Case 102088
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 102089, 102090
        
        Case 102091, 102093
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, objBrowse.objBrowseArquivo.sNomeArq)
        
        Case 102092 'Sem dados... deve ser tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142578)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Browse_Le_NumRegistros2(ByVal objBrowse As AdmBrowse, lNumRegistros As Long) As Long
'Lê no banco de dados o número de registros conforme o comando sql passado em objBrowse
'objBrowse RECEBE(Input) o comando SQL que deve ser utilizado para obter o número de registros
'lNumRegistros RETORNA(Output) a quantidade total de registros

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim vNumRegistros As Variant
Dim vSelecao As Variant
Dim avCampos(1 To BROWSER_NUM_MAX_CAMPOS) As Variant

On Error GoTo Erro_Browse_Le_NumRegistros2

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102099
    
    'Prepara a execução do select  que retorna o número de registros
    lErro = Comando_PrepararInt1(lComando, objBrowse.sSQLCount)
    If lErro <> AD_SQL_SUCESSO Then gError 86243
    
    vNumRegistros = lNumRegistros
    
    'Faz o bind da variável que irá retornar o número de registros
    lErro = Comando_BindVarInt1(lComando, vNumRegistros)
    If lErro <> AD_SQL_SUCESSO Then gError 86242
    
    iIndice = 0
    
    'Se colSelecao está instanciado
    If Not (objBrowse.colSelecao Is Nothing) Then
    
        'Para cada parâmetro de seleção na coleção
        For Each vSelecao In objBrowse.colSelecao
        
            'Incrementa iIndice
            iIndice = iIndice + 1
            
            'Guarda o parâmetro de seleção em um array de variants
            avCampos(iIndice) = vSelecao
            
            'Faz o bind do array
            lErro = Comando_BindVarInt1(lComando, avCampos(iIndice))
            If lErro <> AD_SQL_SUCESSO Then gError 86241
            
        Next
        
    End If

    'Executa o select
    lErro = Comando_ExecutarInt1(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 86244
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 86245
    
    'Início do trecho alterado por rafael menezes em 19/09/2002
    lNumRegistros = vNumRegistros
    'Fim do trecho alterado por rafael menezes em 19/09/2002

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Browse_Le_NumRegistros2 = SUCESSO
    
    Exit Function
    
Erro_Browse_Le_NumRegistros2:

    Browse_Le_NumRegistros2 = gErr
    
    Select Case gErr

        Case 102099
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 102100, 102101, 86241 To 86245
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, objBrowse.objBrowseArquivo.sNomeArq)
        
        Case 102102 'Sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142579)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'??? Daqui p/ baixo leo

Function TelasCampos_Le_Todos(ByVal colTelasCampos As Collection, ByVal objTelasOpcoes As ClassTelasOpcoes) As Long
'Le todos os registros de TelasCampos de acordo com os filtros recebidos como parâmetro.

Dim lErro As Long
Dim objTelasCampos As ClassTelasCampos
Dim alComando(0) As Long
Dim iIndice As Integer
Dim sNomeCampo As String

On Error GoTo Erro_TelasCampos_Le_Todos

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 103241
    Next
     
    sNomeCampo = String(STRING_TELASCAMPOS_NOMECAMPO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT NomeCampo FROM TelasCampos WHERE NomeTela = ? AND TituloTela = ?", _
    sNomeCampo, objTelasOpcoes.sNomeTela, objTelasOpcoes.sTituloTela)
    If lErro <> AD_SQL_SUCESSO Then gError 103242
         
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103243
    
    Do While lErro = AD_SQL_SUCESSO
                 
        Set objTelasCampos = New ClassTelasCampos
        
        objTelasCampos.sTituloTela = objTelasOpcoes.sTituloTela
        objTelasCampos.sNomeTela = objTelasOpcoes.sNomeTela
        objTelasCampos.sNomeCampo = sNomeCampo
        
        colTelasCampos.Add objTelasCampos
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103244
         
    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   TelasCampos_Le_Todos = SUCESSO
   
   Exit Function
   
Erro_TelasCampos_Le_Todos:

   TelasCampos_Le_Todos = gErr

    Select Case gErr

        Case 103241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 103242, 103243, 103244
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELASCAMPOS", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142580)

    End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Carregar_Opcoes(ByVal objTelasOpcoes As ClassTelasOpcoes) As Long
'Carrega a combo de opçoes com as opcoes referentes a tela

Dim lErro As Long
Dim objTelasOpcoesAux As ClassTelasOpcoes
Dim colTelasOpcoes As New Collection

On Error GoTo Erro_Carregar_Opcoes

    'Lê todas as Opções referentes a tela em questão.
    lErro = CF("TelasOpcoes_Le", objTelasOpcoes, colTelasOpcoes)
    If lErro <> SUCESSO Then gError 103245
    
    For Each objTelasOpcoesAux In colTelasOpcoes
            
        objTelasOpcoes.objComboOpcoes.AddItem objTelasOpcoesAux.sNomeOpcao
            
    Next
   
   Carregar_Opcoes = SUCESSO
   
   Exit Function
   
Erro_Carregar_Opcoes:

   Carregar_Opcoes = gErr

    Select Case gErr
        
        Case 103245
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142581)

    End Select
    
    Exit Function

End Function

Function TelasOpcoes_Le(ByVal objTelasOpcoes As ClassTelasOpcoes, ByVal colTelasOpcoes As Collection) As Long
'Le todas as opcoes referente a tela

Dim lErro As Long
Dim objTelasOpcoesAux As ClassTelasOpcoes
Dim alComando(0) As Long
Dim iIndice As Integer
Dim sNomeOpcao As String

On Error GoTo Erro_TelasOpcoes_Le

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 103246
    Next
     
    sNomeOpcao = String(STRING_TELASOPCOESVALORES_NOMEOPCAO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT NomeOpcao FROM TelasOpcoes WHERE NomeTela = ? AND TituloTela = ?", _
    sNomeOpcao, objTelasOpcoes.sNomeTela, objTelasOpcoes.sTituloTela)
    If lErro <> AD_SQL_SUCESSO Then gError 103247
         
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103248
    
    Do While lErro = AD_SQL_SUCESSO
                 
        Set objTelasOpcoesAux = New ClassTelasOpcoes
        
        objTelasOpcoesAux.sTituloTela = objTelasOpcoes.sTituloTela
        objTelasOpcoesAux.sNomeTela = objTelasOpcoes.sNomeTela
        objTelasOpcoesAux.sNomeOpcao = sNomeOpcao
        
        colTelasOpcoes.Add objTelasOpcoesAux
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103249
    
    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   TelasOpcoes_Le = SUCESSO
   
   Exit Function
   
Erro_TelasOpcoes_Le:

   TelasOpcoes_Le = gErr

    Select Case gErr

        Case 103246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 103247, 103248, 103249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELASOPCOES", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142582)

    End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function TelasOpcoesValores_Le(ByVal objTelasOpcoesValores As ClassTelasOpcoesValores) As Long
'Le o layout da tela referente a opcao

Dim lErro As Long
Dim alComando(0) As Long
Dim iIndice As Integer
Dim sValorCampo As String

On Error GoTo Erro_TelasOpcoesValores_Le

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 103250
    Next
    
    sValorCampo = String(STRING_TELASOPCOESVALORES_VALORCAMPO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT ValorCampo FROM TelasOpcoesValores WHERE NomeTela = ? AND TituloTela = ? AND NomeOpcao = ? AND NomeCampo = ?", _
    sValorCampo, objTelasOpcoesValores.sNomeTela, objTelasOpcoesValores.sTituloTela, objTelasOpcoesValores.sNomeOpcao, objTelasOpcoesValores.sNomeCampo)
    If lErro <> AD_SQL_SUCESSO Then gError 103251
         
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103252
    
    If lErro <> AD_SQL_SUCESSO Then gError 103253
    
    objTelasOpcoesValores.sValorCampo = sValorCampo
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   TelasOpcoesValores_Le = SUCESSO
   
   Exit Function
   
Erro_TelasOpcoesValores_Le:

   TelasOpcoesValores_Le = gErr

    Select Case gErr

        Case 103250
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 103251, 103252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELASOPCOESVALORES", gErr)
    
        Case 103253 'tratado da rotina chamadora
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142583)

    End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Traz_Opcao_Tela(ByVal objTelasOpcoes As ClassTelasOpcoes) As Long
'Preenche a tela com a opcao selecionada

Dim lErro As Long
Dim objControle As Object
Dim colTelasCampos As New Collection
Dim objTelasOpcoesValores As New ClassTelasOpcoesValores
Dim iCont As Integer
Dim objTelasCampos As ClassTelasCampos

On Error GoTo Erro_Traz_Opcao_Tela
    
    objTelasOpcoes.sNomeOpcao = objTelasOpcoes.objComboOpcoes.Text
        
    lErro = CF("TelasCampos_Le_Todos", colTelasCampos, objTelasOpcoes)
    If lErro <> SUCESSO Then gError 103286
    
    For Each objTelasCampos In colTelasCampos
            
        objTelasOpcoesValores.sTituloTela = objTelasOpcoes.sTituloTela
        objTelasOpcoesValores.sNomeTela = objTelasOpcoes.sNomeTela
        objTelasOpcoesValores.sNomeOpcao = objTelasOpcoes.sNomeOpcao
        objTelasOpcoesValores.sNomeCampo = objTelasCampos.sNomeCampo
        objTelasOpcoesValores.sValorCampo = ""
        
        lErro = CF("TelasOpcoesValores_Le", objTelasOpcoesValores)
        If lErro <> SUCESSO And lErro <> 103253 Then gError 103252
        
        If lErro <> SUCESSO Then Set objTelasOpcoesValores = New ClassTelasOpcoesValores
        
        Set objControle = objTelasOpcoes.objTela.Controls(objTelasCampos.sNomeCampo)
        
        Select Case TypeName(objControle)
        
            Case "TextBox"
                            
                objControle.Text = objTelasOpcoesValores.sValorCampo
                
                Call CallByName(objTelasOpcoes.objTela, objTelasCampos.sNomeCampo & "_Validate", VbMethod, bSGECancelDummy)
                
            Case "MaskEdBox"
                            
                objControle.PromptInclude = False
                objControle.Text = objTelasOpcoesValores.sValorCampo
                objControle.PromptInclude = True
                
                Call CallByName(objTelasOpcoes.objTela, objTelasCampos.sNomeCampo & "_Validate", VbMethod, bSGECancelDummy)
                
            Case "ComboBox"
            
                For iCont = 0 To objControle.ListCount - 1
                        
                    If objControle.List(iCont) = objTelasOpcoesValores.sValorCampo Then
                    
                        objControle.ListIndex = iCont
                        Exit For
                        
                    End If
                    
                Next
                    
            Case "CheckBox"
                
                If Len(Trim(objTelasOpcoesValores.sValorCampo)) = 0 Then
                    objControle.Value = DESMARCADO
                Else
                    objControle.Value = objTelasOpcoesValores.sValorCampo
                End If
                
            Case "OptionButton"
                
                If Len(Trim(objTelasOpcoesValores.sValorCampo)) = 0 Then
                    objControle.Value = False
                Else
                    objControle.Value = objTelasOpcoesValores.sValorCampo
                End If
                
        End Select
        
    Next
    
   Traz_Opcao_Tela = SUCESSO
   
   Exit Function
   
Erro_Traz_Opcao_Tela:

   Traz_Opcao_Tela = gErr

    Select Case gErr

        Case 103252
        
        Case 103286
        
        Case 438
            Resume Next
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142584)

    End Select
   
    Exit Function

End Function

Public Function CotacaoMoeda_Le(ByVal objCotacaoMoeda As ClassCotacaoMoeda) As Long

Dim lComando As Long
Dim lErro As Long
Dim tCotacaoMoeda As typeCotacaoMoeda

On Error GoTo Erro_CotacaoMoeda_Le

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80264

    'Verifica se existe cotação para a data informada.
    lErro = Comando_Executar(lComando, "SELECT Valor FROM CotacoesMoeda WHERE Data = ? AND Moeda = ?", tCotacaoMoeda.dValor, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)
    If lErro <> AD_SQL_SUCESSO Then gError 80265

    '===> Caso exista seleciona
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80266

    '===> Se não exite Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 80267

    'Carrega objCotacaoMoeda
    objCotacaoMoeda.dValor = tCotacaoMoeda.dValor
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    CotacaoMoeda_Le = SUCESSO

Exit Function

Erro_CotacaoMoeda_Le:

    CotacaoMoeda_Le = gErr

    Select Case gErr

        Case 80264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 80265, 80266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACOESMOEDA", gErr, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)

        Case 80267 'CotacaoMoeda não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142585)

    End Select

    Call Comando_Fechar(lComando)

End Function

Function CotacaoMoeda_Le_UltimasCotacoes(ByVal objCotacaoMoeda As ClassCotacaoMoeda, objCotacaoMoedaAnterior As ClassCotacaoMoeda) As Long
'Le a cotação de data anterior a atual exibida na tela

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim dValor As Double
Dim dtData As Date
Dim dValorAnterior As Double

On Error GoTo Erro_CotacaoMoeda_Le_UltimasCotacoes

    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80230
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 80231

    'Seleciona a data atual no sistema
    lErro = Comando_Executar(lComando, "SELECT Valor FROM CotacoesMoeda WHERE Data = ? AND Moeda = ?", dValor, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)
    If lErro <> AD_SQL_SUCESSO Then gError 80232

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80233

    dtData = DATA_NULA 'para identificar que nao achou cotacao em data anterior
    
    'Seleciona a data anterior em ordem decrescente a partir da data atual contida em dtData
    lErro = Comando_Executar(lComando1, "SELECT Data, Valor FROM CotacoesMoeda WHERE Data < ? AND Moeda = ? Order By Data DESC", dtData, dValorAnterior, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)
    If lErro <> AD_SQL_SUCESSO Then gError 80234
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80235
        
    'Atribui valores ao objCotacaoMoedaAnterior
    objCotacaoMoedaAnterior.dValor = dValorAnterior
    objCotacaoMoedaAnterior.dtData = dtData
    objCotacaoMoeda.dValor = dValor
    
    'Fecha Comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    CotacaoMoeda_Le_UltimasCotacoes = SUCESSO
    
    Exit Function
    
Erro_CotacaoMoeda_Le_UltimasCotacoes:
    
    CotacaoMoeda_Le_UltimasCotacoes = gErr
    
    Select Case gErr
    
        Case 80230, 80231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 80232, 80233, 80234, 80235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACOESMOEDA", gErr, objCotacaoMoeda.dtData)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142586)
    
    End Select
    
    'Fecha Comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
    
End Function

Function Moeda_Critica_Exclusao(objMoedas As ClassMoedas) As Long
'Verifica se a moeda já foi utlizada no sistema

Dim lErro As Long
Dim iMoeda As Integer
Dim alComando(2) As Long
Dim iIndice As Integer

On Error GoTo Erro_Moeda_Critica_Exclusao

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 114516
    Next
    
    'Verifica na tabela de PedidoCompraN o uso da moeda em questao passada no obj
    lErro = Comando_Executar(alComando(0), "SELECT Moeda FROM PedidoCompraN WHERE Moeda=?", iMoeda, objMoedas.iCodigo)
    If lErro <> SUCESSO Then gError 114517
    
    'Busca o primeiro registro - Se encontrar algum -> não pode excluir a moeda
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 114570
    If lErro <> AD_SQL_SEM_DADOS Then gError 114518
    
    'Verifica na tabela de ItensCotacaoN o uso da moeda em questao passada no obj
    lErro = Comando_Executar(alComando(1), "SELECT Moeda FROM ItensCotacaoN WHERE Moeda=?", iMoeda, objMoedas.iCodigo)
    If lErro <> SUCESSO Then gError 114521
    
    'Busca o primeiro registro - Se encontrar algum -> não pode excluir a moeda
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 114571
    If lErro <> AD_SQL_SEM_DADOS Then gError 114522
    
    'Verifica na tabela de CotacoesMoeda o uso da moeda em questao passada no obj
    lErro = Comando_Executar(alComando(2), "SELECT Moeda FROM CotacoesMoeda WHERE Moeda=?", iMoeda, objMoedas.iCodigo)
    If lErro <> SUCESSO Then gError 114523
    
    'Busca o primeiro registro - Se encontrar algum -> não pode excluir a moeda
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 114572
    If lErro <> AD_SQL_SEM_DADOS Then gError 114524
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Moeda_Critica_Exclusao = SUCESSO
    
    Exit Function
    
Erro_Moeda_Critica_Exclusao:

    Moeda_Critica_Exclusao = gErr
    
    Select Case gErr
    
        Case 114516
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 114517
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
        
        Case 114518
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_UTILIZADA_PEDIDOCOMPRA", gErr, objMoedas.sNome)
        
        Case 114522
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_UTILIZADA_ITENSCOTACAO", gErr, objMoedas.sNome)
        
        Case 114524
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_UTILIZADA_COTACOESMOEDA", gErr, objMoedas.sNome)
            
        Case 114521
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCOTACAO", gErr)
        
        Case 114523
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACOESMOEDA", gErr)
            
        Case 114570 To 114572
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142587)
    
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Grades_Le_Todas(colGrade As Collection) As Long
'Lê todas as grades do Banco de Dados

Dim lComando As Long
Dim objGrade As ClassGrade
Dim tGrade As typeGrade
Dim lErro As Long

On Error GoTo Erro_Grades_Le_Todas

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86264
    
    tGrade.sCodigo = String(STRING_GRADE_CODIGO, 0)
    tGrade.sDescricao = String(STRING_GRADE_DESCRICAO, 0)

    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao FROM Grade", tGrade.sCodigo, tGrade.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 86265
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86266
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objGrade = New ClassGrade
        
        objGrade.sCodigo = tGrade.sCodigo
        objGrade.sDescricao = tGrade.sDescricao
        
        colGrade.Add objGrade

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86267
        
    Loop
    
    Call Comando_Fechar(lComando)
    
    Grades_Le_Todas = SUCESSO
    
    Exit Function

Erro_Grades_Le_Todas:

    Grades_Le_Todas = gErr
    
    Select Case gErr
    
        Case 86264
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 86265, 86266, 86267
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRADE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142588)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Grade_Le(objGrade As ClassGrade) As Long
'Lê a Grade com o código passado em objGrade

Dim lErro As Long
Dim lComando As Long
Dim tGrade As typeGrade

On Error GoTo Erro_Grade_Le

    'ABre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86272
    
    'Inicializa a string de Descricao
    tGrade.sDescricao = String(STRING_GRADE_DESCRICAO, 0)
    
    'Seleciona no BD a Grade com o Código passado
    lErro = Comando_Executar(lComando, "SELECT Descricao, Layout FROM Grade WHERE Codigo = ?", tGrade.sDescricao, tGrade.iLayout, objGrade.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 86273
    
    'Verifica se existe a Grade
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86274
    
    'Se não existir --> Retorna Erro
    If lErro <> AD_SQL_SUCESSO Then gError 86275
    
    'Guarda a Descrição da Grade lida
    objGrade.sDescricao = tGrade.sDescricao
    objGrade.iLayout = tGrade.iLayout
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Grade_Le = SUCESSO
    
    Exit Function
    
Erro_Grade_Le:

    Grade_Le = gErr
    
    Select Case gErr
    
        Case 86272
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 86273, 86274
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRADE1", gErr)
        
        Case 86275
            'Sinaliza que não encontrou a grade com o código passado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142589)
    
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function


''''Function GradeCategoria_Le(objGrade As ClassGrade) As Long
'''''Lê as categorias para a Grade com o código passado em objGrade
''''
''''Dim lErro As Long
''''Dim lComando As Long
''''Dim sCategoria As String
''''
''''On Error GoTo Erro_GradeCategoria_Le
''''
''''    'ABre o Comando
''''    lComando = Comando_Abrir()
''''    If lComando = 0 Then gError 86272
''''    ''????cyntia
''''    'Inicializa a coleção onde ficarão armazenadas as Categorias Lidas
''''    'Set objGrade.colCategoria = New Collection
''''
''''    'Inicializa o buffer da variável que vai ser usada para ler a categoria
''''    sCategoria = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
''''
''''    'Seleciona no BD as Categorias da Grade com o Código passado
''''    lErro = Comando_Executar(lComando, "SELECT Categoria FROM GradeCategoria WHERE Grade = ?", sCategoria, objGrade.sCodigo)
''''    If lErro <> AD_SQL_SUCESSO Then gError 86273
''''
''''    'Verifica se existe a Grade
''''    lErro = Comando_BuscarPrimeiro(lComando)
''''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86274
''''
''''    Do While lErro <> AD_SQL_SEM_DADOS
''''    ''????cyntia
'''''        'Guarda a Descrição da Grade lida
'''''        objGrade.colCategoria.Add sCategoria
'''''
''''        lErro = Comando_BuscarProximo(lComando)
''''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86274
''''
''''    Loop
''''
''''
''''    'Fecha o comando
''''    Call Comando_Fechar(lComando)
''''
''''    GradeCategoria_Le = SUCESSO
''''
''''    Exit Function
''''
''''Erro_GradeCategoria_Le:
''''
''''    GradeCategoria_Le = gErr
''''
''''    Select Case gErr
''''
''''        Case 86272
''''            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
''''
''''        Case 86273, 86274
''''            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRADE1", gErr)
''''
''''        Case 86275
''''            'Sinaliza que não encontrou a grade com o código passado
''''
''''        Case Else
''''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142590)
''''
''''    End Select
''''
''''    Call Comando_Fechar(lComando)
''''
''''    Exit Function
''''
''''End Function

Public Function Anotacoes_Le_NumIntDoc(ByVal objAnotacoes As ClassAnotacoes) As Long
'busca o numintdoc da tabela anotacoes a partir da data/titulo/id/origem

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_Anotacoes_Le_NumIntDoc

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 119854

    'busca os registros da tabela anotacoes a partir dos demais dados
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM Anotacoes WHERE DataAlteracao = ? AND Origem = ? AND ID = ? AND Titulo = ?", lNumIntDoc, objAnotacoes.dtDataAlteracao, objAnotacoes.iTipoDocOrigem, objAnotacoes.sID, objAnotacoes.sTitulo)
    If lErro <> AD_SQL_SUCESSO Then gError 119855

    'busca o 1ºregistro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119856

    'erro sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 119857
     
    'carrega o obj c/ o numintdoc lido
    objAnotacoes.lNumIntDoc = lNumIntDoc
    
    'Fechamento do comando
    Call Comando_Fechar(lComando)

   Anotacoes_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_Anotacoes_Le_NumIntDoc:

    Anotacoes_Le_NumIntDoc = gErr

    Select Case gErr

        Case 119857  'sem dados

        Case 119854
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)

        Case 119855, 119856
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ANOTACOES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142591)

    End Select

    'Fechamento do comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Anotacoes_Le_OrigemID(ByVal objAnotacoes As ClassAnotacoes) As Long
'le a todos as anotações referentes ao Id/Origem passados

Dim lErro As Long
Dim lComando As Long
Dim tAnotacoes As typeAnotacoes

On Error GoTo Erro_Anotacoes_Le_OrigemID

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 119841

    'inicializa a string p/ receber o titulo
    tAnotacoes.sTitulo = String(STRING_ANOTACOES_TITULO, 0)

    'busca os registros da tabela anotacoes a partir do Id/Origem passadas
    lErro = Comando_Executar(lComando, "SELECT TOP 1 NumIntDoc, Titulo, DataAlteracao FROM Anotacoes WHERE Origem = ? AND ID = ? ORDER BY DataAlteracao DESC, NumIntDoc DESC", _
        tAnotacoes.lNumIntDoc, tAnotacoes.sTitulo, tAnotacoes.dtDataAtualizacao, objAnotacoes.iTipoDocOrigem, objAnotacoes.sID)
    If lErro <> AD_SQL_SUCESSO Then gError 119842

    'busca o 1ºregistro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119843

    'erro sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 119844
    
    'carrega o obj com o registro + recente da tabela anotacoes
    With tAnotacoes
        
        objAnotacoes.dtDataAlteracao = .dtDataAtualizacao
        objAnotacoes.lNumIntDoc = .lNumIntDoc
        objAnotacoes.sTitulo = .sTitulo
        
    End With
            
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

   Anotacoes_Le_OrigemID = SUCESSO

    Exit Function

Erro_Anotacoes_Le_OrigemID:

    Anotacoes_Le_OrigemID = gErr

    Select Case gErr

        Case 119841
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)

        Case 119842, 119843
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ANOTACOES", gErr)
        
        Case 119844  'sem dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 142592)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function AnotacoesLinha_Le_Texto(colAnotacoesLinha As Collection, ByVal lNumIntDoc As Long) As Long
'busca os "pedaços" do texto em anotacoeslinha a partir do NumIntDoc passado como parametro
'devolve todos as "partes" do texto e a seq. na collection

Dim lErro As Long
Dim lComando As Long
Dim iSeq As Integer
Dim sTexto As String
Dim objAnotacoesLinha As ClassAnotacoesLinha

On Error GoTo Erro_AnotacoesLinha_Le_Texto

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 119845

    'inicializa a string p/ receber as partes do texto
    sTexto = String(STRING_ANOTACOESLINHA_TEXTO, 0)

    'busca os registros da tabela anotacoeslinha a partir do NumIntDoc
    lErro = Comando_Executar(lComando, "SELECT Seq, Texto FROM AnotacoesLinha WHERE NumIntDoc = ? ORDER BY Seq", iSeq, sTexto, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 119846

    'busc o 1ºregistro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119847

    'erro sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 119848
    
    'enquanto houverem "pedaços" do texto
    Do While lErro = SUCESSO
    
        'instancia o obj
        Set objAnotacoesLinha = New ClassAnotacoesLinha
            
        'carrega o obj c/ a seq. de "montagem" e a parte correspondente do texto
        objAnotacoesLinha.iSeq = iSeq
        objAnotacoesLinha.sTexto = sTexto
              
        'adiciona o obj na collection
        colAnotacoesLinha.Add objAnotacoesLinha
        
        'busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119849
              
     Loop
     
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

   AnotacoesLinha_Le_Texto = SUCESSO

    Exit Function

Erro_AnotacoesLinha_Le_Texto:

    AnotacoesLinha_Le_Texto = gErr

    Select Case gErr

        Case 119848  'sem dados

        Case 119845
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)

        Case 119846, 119847, 119849
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ANOTACOESLINHA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142593)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function OrigemAnotacoes_Le(ByVal objOrigemAnotacoes As ClassOrigemAnotacoes) As Long
'preenche o obj c/ todas as informações lidas a partir do código de origem

Dim lErro As Long
Dim lComando As Long
Dim tOrigemAnotacoes As typeOrigemAnotacoes

On Error GoTo Erro_OrigemAnotacoes_Le

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 119850

    'inicializa as strings
    tOrigemAnotacoes.sDescricao = String(STRING_ORIGEMANOTACOES_DESCRICAO, 0)
    tOrigemAnotacoes.sNomeTabela = String(STRING_ORIGEMANOTACOES_NOMETABELA, 0)

    'busca os registros da tabela origensanotacoes a partir do codigo
    lErro = Comando_Executar(lComando, "SELECT Descricao, NomeTabela FROM OrigensAnotacoes WHERE Codigo = ?", tOrigemAnotacoes.sDescricao, tOrigemAnotacoes.sNomeTabela, objOrigemAnotacoes.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 119851

    'busca o 1ºregistro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119852

    'erro sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 119853
     
    'carrega o obj c/ os dados lidos
    objOrigemAnotacoes.sDescricao = tOrigemAnotacoes.sDescricao
    objOrigemAnotacoes.sNomeTabela = tOrigemAnotacoes.sNomeTabela
     
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

   OrigemAnotacoes_Le = SUCESSO

    Exit Function

Erro_OrigemAnotacoes_Le:

    OrigemAnotacoes_Le = gErr

    Select Case gErr

        Case 119853  'sem dados

        Case 119850
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)

        Case 119851, 119852
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGENSANOTACOES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142594)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function OrigemAnotacoes_Verifica_ID(ByVal sNomeTabela As String, ByVal sID As String) As Long
'verifica se o ID passado existe na tabela de origem

Dim lErro As Long
Dim lComando As Long
Dim iCount As Integer

On Error GoTo Erro_OrigemAnotacoes_Verifica_ID
    
    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 119863

    'verifica se o Id passado existe na tabela referente a origem
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM " & sNomeTabela & " WHERE " & sID, iCount)
    If lErro <> AD_SQL_SUCESSO Then gError 119864

    'busca o 1ºregistro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119865

    'erro sem dados
    If iCount = 0 Then gError 119866
     
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    OrigemAnotacoes_Verifica_ID = SUCESSO

    Exit Function

Erro_OrigemAnotacoes_Verifica_ID:

    OrigemAnotacoes_Verifica_ID = gErr

    Select Case gErr
        
        Case 119863
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)
        
        Case 119864, 119865
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGENSANOTACOES", gErr)
        
        Case 119866 'sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142595)
            
    End Select
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Public Function AnotacoesLinha_Le_Texto1(sTextoCompleto As String, ByVal lNumIntDoc As Long) As Long
'busca os "pedaços" do texto em anotacoeslinha a partir do NumIntDoc passado como parametro
'devolve todos as "partes" do texto e a seq. na collection

Dim lErro As Long
Dim lComando As Long
Dim sTextoParcial As String

On Error GoTo Erro_AnotacoesLinha_Le_Texto1

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 119877

    'inicializa a string p/ receber as partes do texto
    sTextoParcial = String(STRING_ANOTACOESLINHA_TEXTO, 0)

    'busca os registros da tabela anotacoeslinha a partir do NumIntDoc
    lErro = Comando_Executar(lComando, "SELECT Texto FROM AnotacoesLinha WHERE NumIntDoc = ? ORDER BY Seq", sTextoParcial, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 119873

    'busc o 1ºregistro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119874

    'erro sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 119875
    
    'enquanto houverem "pedaços" do texto
    Do While lErro = SUCESSO
    
        'monta o texto novamente na string passada como parametro
        sTextoCompleto = sTextoCompleto & sTextoParcial
                      
        'busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119876
              
     Loop
     
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

   AnotacoesLinha_Le_Texto1 = SUCESSO

    Exit Function

Erro_AnotacoesLinha_Le_Texto1:

    AnotacoesLinha_Le_Texto1 = gErr

    Select Case gErr

        Case 119875  'sem dados

        Case 119877
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)

        Case 119874, 119873, 119876
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ANOTACOESLINHA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142596)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Cidade_Le(ByVal objCidades As ClassCidades) As Long
'Le as informações sobre o código passado

Dim lErro As Long
Dim lComando As Long
Dim tCidades As typeCidades

On Error GoTo Erro_Cidade_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 125038
    
    'Inicializa as strings
    tCidades.sDescricao = String(STRING_CIDADE, 0)
    tCidades.sCodIBGE = String(STRING_CIDADE_CODIBGE, 0)
    
    'Realiza a seleção através do Código que foi passado
    lErro = Comando_Executar(lComando, "SELECT Descricao, CodIBGE FROM Cidades WHERE Codigo = ?", tCidades.sDescricao, tCidades.sCodIBGE, objCidades.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 125039
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125040
    
    If lErro <> AD_SQL_SUCESSO Then gError 125041
    
    objCidades.sDescricao = tCidades.sDescricao
    objCidades.sCodIBGE = tCidades.sCodIBGE
        
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Cidade_Le = SUCESSO
    
    Exit Function
    
Erro_Cidade_Le:

    Cidade_Le = gErr
    
    Select Case gErr

        Case 125038
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 125039, 125040
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CIDADES", gErr, objCidades.iCodigo)
        
        Case 125041 'Código não encontrado no Banco de Dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142597)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Cidade_LeTodos(colCidades As Collection) As Long
'Le as informações sobre todos os códigos

Dim lErro As Long
Dim lComando As Long
Dim tCidades As typeCidades
Dim objCidades As ClassCidades

On Error GoTo Erro_Cidade_LeTodos

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 125042
    
    'Inicializa as strings
    tCidades.sDescricao = String(STRING_CIDADE, 0)
    tCidades.sCodIBGE = String(STRING_CIDADE_CODIBGE, 0)
    
    'Realiza a seleção através do Código que foi passado
    lErro = Comando_Executar(lComando, "SELECT Descricao, Codigo, CodIBGE FROM Cidades", tCidades.sDescricao, tCidades.iCodigo, tCidades.sCodIBGE)
    If lErro <> AD_SQL_SUCESSO Then gError 125043
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125044
    
    Do While lErro = SUCESSO
    
        Set objCidades = New ClassCidades
        
        objCidades.iCodigo = tCidades.iCodigo
        objCidades.sDescricao = tCidades.sDescricao
        objCidades.sCodIBGE = tCidades.sCodIBGE
            
        colCidades.Add objCidades
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125045

    Loop

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Cidade_LeTodos = SUCESSO
    
    Exit Function
    
Erro_Cidade_LeTodos:

    Cidade_LeTodos = gErr
    
    Select Case gErr

        Case 125042
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 125043, 125044, 125045
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CIDADES", gErr, objCidades.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142598)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Cidade_Le_Nome(ByVal objCidades As ClassCidades) As Long
'Le as informações sobre a cidade com o nome passado

Dim lErro As Long
Dim lComando As Long
Dim tCidades As typeCidades

On Error GoTo Erro_Cidade_Le_Nome

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 124084
    
    tCidades.sCodIBGE = String(STRING_CIDADE_CODIBGE, 0)
    
    'Realiza a seleção através do nome que foi passado
    lErro = Comando_Executar(lComando, "SELECT Codigo, CodIBGE FROM Cidades WHERE Descricao = ?", tCidades.iCodigo, tCidades.sCodIBGE, objCidades.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 124085
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124086
    
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_OBJETO_NAO_CADASTRADO
    
    objCidades.iCodigo = tCidades.iCodigo
    objCidades.sCodIBGE = tCidades.sCodIBGE
        
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Cidade_Le_Nome = SUCESSO
    
    Exit Function
    
Erro_Cidade_Le_Nome:

    Cidade_Le_Nome = gErr
    
    Select Case gErr

        Case 124084
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 124085, 124086
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CIDADES2", gErr, objCidades.sDescricao)
        
        Case ERRO_OBJETO_NAO_CADASTRADO 'Código não encontrado no Banco de Dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142599)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Browser_Trata_Parametros_Customizado(ByVal sNomeTela As String, ByVal objBrowse1 As Object) As Long
'le os valores dos parametros de selecao relacionados a sNomeTela coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim tBrowseParamSelecao As typeBrowseParamSelecao
    
On Error GoTo Erro_Browser_Trata_Parametros_Customizado

    'carrega os valores dos parametros de selecao que ainda faltam
    lErro = CF("BrowseParamSelecao_Le", sNomeTela, objBrowse1.colSelecao)
    If lErro <> SUCESSO Then gError 89977

    Browser_Trata_Parametros_Customizado = SUCESSO
    
    Exit Function
    
Erro_Browser_Trata_Parametros_Customizado:

    Browser_Trata_Parametros_Customizado = gErr

    Select Case gErr
    
        Case 89977
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142600)
        
    End Select
    
    Exit Function

End Function


Function GrupoUsuarios_Le_Todos(ByVal colGrupoUsu As Collection) As Long

Dim lComando As Long
Dim lErro As Long
Dim tGrupoUsu As typeGrupoUsuario
Dim objGrupoUsu As ClassGrupoUsuarios
    
On Error GoTo Erro_GrupoUsuarios_Le_Todos

    lComando = 0

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 129294

    tGrupoUsu.sCodGrupo = String(10, 0)
    tGrupoUsu.sDescricao = String(50, 0)

    lErro = Comando_Executar(lComando, "SELECT CodGrupo, Descricao, DataValidade, LogAtividade  FROM GruposDeUsuarios", _
                                        tGrupoUsu.sCodGrupo, tGrupoUsu.sDescricao, tGrupoUsu.dtDataValidade, tGrupoUsu.iLogAtividade)
    If lErro <> AD_SQL_SUCESSO Then gError 129295
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129296
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objGrupoUsu = New ClassGrupoUsuarios
        
        objGrupoUsu.sCodGrupo = tGrupoUsu.sCodGrupo
        objGrupoUsu.sDescricao = tGrupoUsu.sDescricao
        objGrupoUsu.dtDataValidade = tGrupoUsu.dtDataValidade
        objGrupoUsu.iLogAtividade = tGrupoUsu.iLogAtividade
        
        colGrupoUsu.Add objGrupoUsu
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129297
            
    Loop
            
    Call Comando_Fechar(lComando)
    
    GrupoUsuarios_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_GrupoUsuarios_Le_Todos:

    GrupoUsuarios_Le_Todos = Err

    Select Case Err
    
        Case 129294
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 129295 To 129297
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142601)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Relatorio_Customiza_ExecutarAux(ByVal objRelOpcoes As AdmRelOpcoes, sNomeTsk As String) As Long

Dim lErro As Long

On Error GoTo Erro_Relatorio_Customiza_ExecutarAux

    lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(giFilialEmpresa))
    If lErro <> AD_BOOL_TRUE Then gError 126935

    Relatorio_Customiza_ExecutarAux = SUCESSO
    
    Exit Function
    
Erro_Relatorio_Customiza_ExecutarAux:

    Relatorio_Customiza_ExecutarAux = gErr
    
    Select Case gErr
    
        Case 126935
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142602)
        
    End Select
    
    Exit Function

End Function

Function FilialEmpresa_Customiza(iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_FilialEmpresa_Customiza

    FilialEmpresa_Customiza = SUCESSO
    
    Exit Function
    
Erro_FilialEmpresa_Customiza:

    FilialEmpresa_Customiza = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142603)
        
    End Select
    
    Exit Function

End Function

Public Function Anotacao_Le_ComTextoCompleto(ByVal objAnotacao As ClassAnotacoes) As Long
'busca os "pedaços" do texto em anotacoeslinha a partir de Origem/Id/Titulo em objAnotacao como parametro
'e devolve o texto completo em objAnotacao, alem de preencher outros dados de objAnotacao

Dim lErro As Long, lNumIntDoc As Long, sTextoCompleto As String
Dim lComando As Long, dtDataAlteracao As Date, iSeq As Integer
Dim sTextoParcial As String, objAnotacoesLinha As ClassAnotacoesLinha

On Error GoTo Erro_Anotacao_Le_ComTextoCompleto

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 119877

    'inicializa a string p/ receber as partes do texto
    sTextoParcial = String(STRING_ANOTACOESLINHA_TEXTO, 0)

    'busca os registros
    lErro = Comando_Executar(lComando, "SELECT Anotacoes.NumIntDoc, DataAlteracao, Seq, Texto FROM Anotacoes, AnotacoesLinha WHERE Anotacoes.NumIntDoc = AnotacoesLinha.NumIntDoc AND Origem = ? AND ID = ? AND Titulo = ? ORDER BY Seq", _
        lNumIntDoc, dtDataAlteracao, iSeq, sTextoParcial, objAnotacao.iTipoDocOrigem, objAnotacao.sID, objAnotacao.sTitulo)
    If lErro <> AD_SQL_SUCESSO Then gError 119873

    'busc o 1ºregistro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119874

    'erro sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 119875
    
    With objAnotacao
        .lNumIntDoc = lNumIntDoc
        .dtDataAlteracao = dtDataAlteracao
    End With
         
    Set objAnotacao.colTexto = New Collection
    
    'enquanto houverem "pedaços" do texto
    Do While lErro = SUCESSO
    
        'monta o texto novamente na string passada como parametro
        sTextoCompleto = sTextoCompleto & sTextoParcial
                      
        'instancia o obj
        Set objAnotacoesLinha = New ClassAnotacoesLinha
            
        'carrega o obj c/ a seq. de "montagem" e a parte correspondente do texto
        objAnotacoesLinha.iSeq = iSeq
        objAnotacoesLinha.sTexto = sTextoParcial
              
        'adiciona o obj na collection
        objAnotacao.colTexto.Add objAnotacoesLinha
        
        'busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119876
              
     Loop
     
    objAnotacao.sTextoCompleto = sTextoCompleto
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

   Anotacao_Le_ComTextoCompleto = SUCESSO

    Exit Function

Erro_Anotacao_Le_ComTextoCompleto:

    Anotacao_Le_ComTextoCompleto = gErr

    Select Case gErr

        Case 119875  'sem dados

        Case 119877
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)

        Case 119874, 119873, 119876
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ANOTACOESLINHA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142604)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'####################################################################
'Inserido por Wagner
Function Verifica_Configuracoes(ByVal iLimiteUso As Integer, ByVal dtDataReferencia As Date, ByVal iMeses As Integer, ByVal iMediaNFs As Integer) As Long
'Verifica se pode usar o sistema ou não

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim dtDataEmissao As Date
Dim dtDataLimite As Date
Dim dtDataLimiteAviso As Date
Dim dtDataPeriodoIni As Date
Dim dtDataPeriodoFim As Date
Dim lQtdNFs As Long
Dim lQtdNFsMes As Long
Dim bJaAvisou As Boolean
Dim iDiasAviso As Integer
    
On Error GoTo Erro_Verifica_Configuracoes

    iDiasAviso = 10 'Quantidade de dias antes do bloqueio que vai exibir um aviso
    bJaAvisou = False


    'Se não é para limitar o uso => não faz nada
    If iLimiteUso <> LIMITA_DATA_USO_NAO Then

        'Abre Comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 140510
        Next
        
        'Se for para limitar pela Data de Emissão da Nota Fiscal
        If iLimiteUso = LIMITA_DATA_USO_DATA Or iLimiteUso = LIMITA_DATA_USO_DATA_QTD Then
        
            'Acha a data limite para emissão
            dtDataLimite = DateAdd("m", iMeses, dtDataReferencia)
            dtDataLimiteAviso = DateAdd("d", -iDiasAviso, dtDataLimite)
        
            'Lê a maior data de emissão de uma NF
            lErro = Comando_Executar(alComando(1), "SELECT DataEmissao FROM NFiscal ORDER BY DataEmissao DESC", dtDataEmissao)
            If lErro <> AD_SQL_SUCESSO Then gError 140511
            
            'Busca o primeiro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140512
            
            'Se encontrou
            If lErro = AD_SQL_SUCESSO Then
            
                'Limita a Data de Emissão
                If dtDataLimite < dtDataEmissao Then gError 140513
                
                'Apenas Avisa se faltar menos de 10 dias
                If dtDataLimiteAviso < dtDataEmissao And Not bJaAvisou Then
                    Call Rotina_Aviso(vbOKOnly, "ERRO_LEITURA_CONFIGURACOES")
                    bJaAvisou = True
                End If
            
            End If
            
        End If
        
        'Se for para limitar pela Qtds de Nota Fiscal
        If iLimiteUso = LIMITA_DATA_USO_QTD Or iLimiteUso = LIMITA_DATA_USO_DATA_QTD Then
        
            dtDataPeriodoIni = DateAdd("m", -1, dtDataReferencia)
            dtDataPeriodoFim = DateAdd("d", -1, dtDataReferencia)
        
            'Se a média não foi forçada
            If iMediaNFs <> 0 Then
        
                'Lê a Qtds de NFs do mês anterior data de referência (média por mês)
                lErro = Comando_Executar(alComando(1), "SELECT Count(*) FROM NFiscal WHERE DataEmissao Between ? AND ? ", lQtdNFsMes, dtDataPeriodoIni, dtDataPeriodoFim)
                If lErro <> AD_SQL_SUCESSO Then gError 140514
                
                lErro = Comando_BuscarPrimeiro(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140515


                
            Else
                
                lQtdNFsMes = iMediaNFs

            
            End If

        
            'Lê a Qtds de NFs após a data de referência (Base)
            lErro = Comando_Executar(alComando(2), "SELECT Count(*) FROM NFiscal WHERE DataEmissao >= ? ", lQtdNFs, dtDataReferencia)
            If lErro <> AD_SQL_SUCESSO Then gError 140516
            
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140517


            'Se a quantidade emitida for maior que o número de meses passados vezes a média mês -> Erro
            If (lQtdNFsMes * iMeses) < lQtdNFs Then gError 140518
            
            'Avisa um pouco antes
            If (lQtdNFsMes * iMeses) < (lQtdNFs + (lQtdNFsMes * iDiasAviso / 30)) And Not bJaAvisou Then
                Call Rotina_Aviso(vbOKOnly, "ERRO_LEITURA_CONFIGURACOES")
                bJaAvisou = True
            End If
                
        End If
        
        'Fecha o comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If

    
    Verifica_Configuracoes = SUCESSO
    
    Exit Function
    
Erro_Verifica_Configuracoes:

    Verifica_Configuracoes = gErr

    Select Case gErr
    
        Case 140510
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 140511, 140512, 140514, 140515, 140516, 140517
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIMITE", gErr)
        
        Case 140513, 140518
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACOES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142605)
        
    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function
'##############################################################################

'##############################################################################
'Inserido por Wagner
Function Planilhas_Le_GrupoModulo(ByVal colPlanilhas As Collection, ByVal sGrupo As String, sModulo As String) As Long
'retorna as Planilhas referentes ao grupo e modulo que o usuario tem acesso
'os codigos das planilhas sao retornados em colPla e o grupo e modulo sao passados como parametros

Dim lErro As Long
Dim lComando As Long
Dim sCodPla As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Planilhas_Le_GrupoModulo
       
    sCodPla = String(STRING_CODIGO_RELATORIO, 0)
    
    'abre conexao com o BD
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 12003
    
    'seleciona os relatorios que o usuario do grupo tem acesso para o modulo corrente
    lErro = Comando_Executar(lComando, "SELECT Planilhas.CodPla, Planilhas.FilialEmpresa FROM Planilhas , GrupoPlanilhas , PlanilhasModulo WHERE Planilhas.CodPla = GrupoPlanilhas.CodPla AND Planilhas.CodPla = PlanilhasModulo.CodPla AND GrupoPlanilhas.TipoDeAcesso = 1 AND GrupoPlanilhas.CodGrupo = ? AND PlanilhasModulo.Modulo = ?", sCodPla, iFilialEmpresa, sGrupo, sModulo)
    If lErro <> AD_SQL_SUCESSO Then gError 12004
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 12004
        
    Do While lErro = AD_SQL_SUCESSO
                
        'Se tiver como Empresa Toda
        If giFilialEmpresa = EMPRESA_TODA Then
            'Só não entra na colecao quem é de Somente Filial
            If iFilialEmpresa <> FILIAL_EMPRESA Then
                'Adiciona à Coleção
                colPlanilhas.Add (sCodPla)
            End If
        Else 'se tiver como Filial
            'Entra Filial e Ambas menos FilialEmpresaToda
            If iFilialEmpresa <> FILIAL_EMPRESA_TODA Then
                'Adiciona à Coleção
                colPlanilhas.Add (sCodPla)
            End If
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 12004

    Loop
    
    Call Comando_Fechar(lComando)
    
    Planilhas_Le_GrupoModulo = SUCESSO
    
    Exit Function

Erro_Planilhas_Le_GrupoModulo:

    Planilhas_Le_GrupoModulo = gErr
    
    Select Case gErr
                
        Case 12003
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 12004, 12005
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANILHAS2", gErr)
                
        Case 12006
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 142606)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 142607)
        
    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Function Planilhas_Le_CodPla(ByVal objRelPlanilhas As AdmRelPlanilha) As Long
'retorna os dados do relatorio referentes ao Codigo ja passado por parametro em objRelatorio

Dim lErro As Long
Dim lComando As Long
Dim sNome As String
Dim sNomeCustomizado As String
Dim sAutor As String
Dim sDescricao As String

On Error GoTo Erro_Planilhas_Le_CodPla
    
    'inicializa as variaveis referentes as colunas da tabela Relatorios no BD
    sNome = String(STRING_NOME_ARQ_COMPLETO, 0)
    sNomeCustomizado = String(STRING_NOME_ARQ_COMPLETO, 0)
    sAutor = String(STRING_CODIGO_RELATORIO, 0)
    sDescricao = String(STRING_DESC_RELATORIO, 0)
          
    'abre conexao com o BD
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 12007
    
    'faz a selecao dos dados na tabela Relatorios
    lErro = Comando_Executar(lComando, "SELECT  Nome, NomeCustomizado, Autor, Descricao FROM Planilhas WHERE CodPla = ?", sNome, sNomeCustomizado, sAutor, sDescricao, objRelPlanilhas.sCodPla)
    If lErro <> AD_SQL_SUCESSO Then gError 12008
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 12009
           
    With objRelPlanilhas
    
        .sNome = IIf(Len(Trim(sNomeCustomizado)) <> 0, sNomeCustomizado, sNome)
        .sAutor = sAutor
        .sDescricao = sDescricao
        
    End With
    
    Call Comando_Fechar(lComando)
    
    Planilhas_Le_CodPla = SUCESSO
    
    Exit Function

Erro_Planilhas_Le_CodPla:

    Planilhas_Le_CodPla = gErr
    
    Select Case gErr
                
        Case 12007
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 12008, 12009
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANILHAS2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 142608)
        
    End Select
        
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function
'###################################################################################

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ValidaExclusoes_Le(ByVal sCodigo As String, ByVal colValidaExclusoes As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tValidaExclusoes As typeValidaExclusoes
Dim objValidaExclusoes As ClassValidaExclusoes

On Error GoTo Erro_ValidaExclusoes_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 177931

    'Alocação de espaço no buffer
    tValidaExclusoes.sCodigo = String(STRING_VALIDAEXCLUSAO_CODIGO, 0)
    tValidaExclusoes.sTabela = String(STRING_VALIDAEXCLUSAO_TABELA, 0)
    tValidaExclusoes.sCampo = String(STRING_VALIDAEXCLUSAO_CAMPO, 0)
    tValidaExclusoes.sCampoLer = String(STRING_VALIDAEXCLUSAO_CAMPOLER, 0)
    tValidaExclusoes.sMsgErro1 = String(STRING_VALIDAEXCLUSAO_MSGERRO1, 0)
    tValidaExclusoes.sMsgErro2 = String(STRING_VALIDAEXCLUSAO_MSGERRO2, 0)
    tValidaExclusoes.sMsgErroLer = String(STRING_VALIDAEXCLUSAO_MSGERROLER, 0)

    'Le a tabelaValidaExclusoes
    lErro = Comando_Executar(lComando, "SELECT Codigo, Tabela, Campo, CampoLer, MsgErro1, MsgErro2, TipoCampoLer, TamanhoCampoLer, MsgErroLer,SubTipoCampoLer,SubTipoCampoProc,GeneroMsgErro FROM ValidaExclusoes WHERE Codigo= ? ", _
                tValidaExclusoes.sCodigo, tValidaExclusoes.sTabela, tValidaExclusoes.sCampo, tValidaExclusoes.sCampoLer, _
                tValidaExclusoes.sMsgErro1, tValidaExclusoes.sMsgErro2, tValidaExclusoes.iTipoCampoLer, tValidaExclusoes.iTamanhoCampoLer, tValidaExclusoes.sMsgErroLer, tValidaExclusoes.iSubTipoCampoLer, tValidaExclusoes.iSubTipoCampoProc, tValidaExclusoes.iGeneroMsgErro, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 177932

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 177933

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objValidaExclusoes = New ClassValidaExclusoes

        objValidaExclusoes.sCodigo = tValidaExclusoes.sCodigo
        objValidaExclusoes.sTabela = tValidaExclusoes.sTabela
        objValidaExclusoes.sCampo = tValidaExclusoes.sCampo
        objValidaExclusoes.sCampoLer = tValidaExclusoes.sCampoLer
        objValidaExclusoes.sMsgErro1 = tValidaExclusoes.sMsgErro1
        objValidaExclusoes.sMsgErro2 = tValidaExclusoes.sMsgErro2
        objValidaExclusoes.iTamanhoCampoLer = tValidaExclusoes.iTamanhoCampoLer
        objValidaExclusoes.iTipoCampoLer = tValidaExclusoes.iTipoCampoLer
        objValidaExclusoes.sMsgErroLer = tValidaExclusoes.sMsgErroLer
        objValidaExclusoes.iSubTipoCampoLer = tValidaExclusoes.iSubTipoCampoLer
        objValidaExclusoes.iSubTipoCampoProc = tValidaExclusoes.iSubTipoCampoProc
        objValidaExclusoes.iGeneroMsgErro = tValidaExclusoes.iGeneroMsgErro
    
        colValidaExclusoes.Add objValidaExclusoes
    
        'Busca Primeiro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 177934

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ValidaExclusoes_Le = SUCESSO

    Exit Function

Erro_ValidaExclusoes_Le:

    ValidaExclusoes_Le = gErr

    Select Case gErr

        Case 177931
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 177932 To 177934
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VALIDAEXCLUSOES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177934)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TransacaoWFW_Le_Todos(ByVal sSiglaModulo As String, ByVal colTransacao As Collection) As Long
'Le todas as transacoes relativas ao modulo em questão e carrega-as em colTransacao

Dim lErro As Long
Dim lComando As Long
Dim objTransacao As ClassTransacaoWFW
Dim tTransWFW As typeTransacaoWFW
    
On Error GoTo Erro_TransacaoWFW_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178016
    
    tTransWFW.sSigla = String(STRING_TRANSACAOWFW_SIGLA, 0)
    tTransWFW.sTransacao = String(STRING_TRANSACAOWFW_TRANSACAO, 0)
    tTransWFW.sTransacaoTela = String(STRING_TRANSACAOWFW_TRANSACAOTELA, 0)
    tTransWFW.sOrigem = String(STRING_TRANSACAOWFW_ORIGEM, 0)
    tTransWFW.sObservacao = String(STRING_TRANSACAOWFW_OBSERVACAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Sigla, Transacao, TransacaoTela, Origem, Observacao, Codigo FROM TransacaoWFW WHERE Sigla = ?", _
    tTransWFW.sSigla, tTransWFW.sTransacao, tTransWFW.sTransacaoTela, tTransWFW.sOrigem, tTransWFW.sObservacao, tTransWFW.iCodigo, sSiglaModulo)
    If lErro <> AD_SQL_SUCESSO Then gError 178013
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178014
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objTransacao = New ClassTransacaoWFW
        
        objTransacao.sSigla = tTransWFW.sSigla
        objTransacao.sTransacao = tTransWFW.sTransacao
        objTransacao.sTransacaoTela = tTransWFW.sTransacaoTela
        objTransacao.sOrigem = tTransWFW.sOrigem
        objTransacao.sObservacao = tTransWFW.sObservacao
        objTransacao.iCodigo = tTransWFW.iCodigo
    
        colTransacao.Add objTransacao

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178015

    Loop
    
    Call Comando_Fechar(lComando)
    
    TransacaoWFW_Le_Todos = SUCESSO
    
    Exit Function

Erro_TransacaoWFW_Le_Todos:

    TransacaoWFW_Le_Todos = gErr
    
    Select Case gErr
                
        Case 178016
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178013 To 178015
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSACAOWFW", gErr, sSiglaModulo)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 178019)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function TransacaoWFW_Le_Todos2(ByVal sSiglaModulo As String, ByVal colTransacao As Collection) As Long
'Le todas as transacoes

Dim lErro As Long
Dim lComando As Long
Dim objTransacao As ClassTransacaoWFW
Dim tTransWFW As typeTransacaoWFW
    
On Error GoTo Erro_TransacaoWFW_Le_Todos2

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 180112
    
    tTransWFW.sSigla = String(STRING_TRANSACAOWFW_SIGLA, 0)
    tTransWFW.sTransacao = String(STRING_TRANSACAOWFW_TRANSACAO, 0)
    tTransWFW.sTransacaoTela = String(STRING_TRANSACAOWFW_TRANSACAOTELA, 0)
    tTransWFW.sOrigem = String(STRING_TRANSACAOWFW_ORIGEM, 0)
    tTransWFW.sObservacao = String(STRING_TRANSACAOWFW_OBSERVACAO, 0)
    
    If Len(Trim(sSiglaModulo)) <> 0 Then
        lErro = Comando_Executar(lComando, "SELECT Sigla, Transacao, TransacaoTela, Origem, Observacao, Codigo FROM TransacaoWFW WHERE Sigla = ? ORDER BY Transacao", _
        tTransWFW.sSigla, tTransWFW.sTransacao, tTransWFW.sTransacaoTela, tTransWFW.sOrigem, tTransWFW.sObservacao, tTransWFW.iCodigo, sSiglaModulo)
    Else
        lErro = Comando_Executar(lComando, "SELECT Sigla, Transacao, TransacaoTela, Origem, Observacao, Codigo FROM TransacaoWFW ORDER BY Transacao", _
        tTransWFW.sSigla, tTransWFW.sTransacao, tTransWFW.sTransacaoTela, tTransWFW.sOrigem, tTransWFW.sObservacao, tTransWFW.iCodigo)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 180113
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180114
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objTransacao = New ClassTransacaoWFW
        
        objTransacao.sSigla = tTransWFW.sSigla
        objTransacao.sTransacao = tTransWFW.sTransacao
        objTransacao.sTransacaoTela = tTransWFW.sTransacaoTela
        objTransacao.sOrigem = tTransWFW.sOrigem
        objTransacao.sObservacao = tTransWFW.sObservacao
        objTransacao.iCodigo = tTransWFW.iCodigo
    
        colTransacao.Add objTransacao

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180115

    Loop
    
    Call Comando_Fechar(lComando)
    
    TransacaoWFW_Le_Todos2 = SUCESSO
    
    Exit Function

Erro_TransacaoWFW_Le_Todos2:

    TransacaoWFW_Le_Todos2 = gErr
    
    Select Case gErr
                
        Case 180112
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 180113 To 180115
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSACAOWFW", gErr, sSiglaModulo)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 180116)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function MnemonicoWFW_Le(ByVal sModulo As String, ByVal iTransacao As Integer, ByVal colMnemonico As Collection) As Long
'retorna os mnemonicos associados ao modulo/transacao passados como parametro. Os mnemonicos são retornados em colMnemonico

Dim lComando As Long
Dim tMnemonicoWFW As typeMnemonicoWFW
Dim lErro As Long
Dim objMnemonico As ClassMnemonicoWFW

On Error GoTo Erro_MnemonicoWFW_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178026
    
    tMnemonicoWFW.sMnemonico = String(STRING_MNEMONICOWFW_MNEMONICO, 0)
    tMnemonicoWFW.sNomeGrid = String(STRING_NOME_CONTROLE, 0)
    tMnemonicoWFW.sMnemonicoCombo = String(STRING_MNEMONICOWFW_MNEMONICOCOMBO, 0)
    tMnemonicoWFW.sMnemonicoDesc = String(STRING_MNEMONICOWFW_MNEMONICODESC, 0)
    
    'seleciona os mnemonicos do Modulo/Transacao passado como parametro
    lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, NomeGrid, MnemonicoCombo, MnemonicoDesc FROM MnemonicoWFW WHERE Modulo=? AND Transacao=?", _
    tMnemonicoWFW.sMnemonico, tMnemonicoWFW.iTipo, tMnemonicoWFW.iNumParam, tMnemonicoWFW.iParam1, tMnemonicoWFW.iParam2, tMnemonicoWFW.iParam3, tMnemonicoWFW.sNomeGrid, tMnemonicoWFW.sMnemonicoCombo, tMnemonicoWFW.sMnemonicoDesc, sModulo, iTransacao)
    If lErro <> AD_SQL_SUCESSO Then gError 178027
    
    'le o primeiro mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178028
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objMnemonico = New ClassMnemonicoWFW
    
        objMnemonico.sModulo = sModulo
        objMnemonico.iTransacao = iTransacao
        objMnemonico.sMnemonico = tMnemonicoWFW.sMnemonico
        objMnemonico.iTipo = tMnemonicoWFW.iTipo
        objMnemonico.iNumParam = tMnemonicoWFW.iNumParam
        objMnemonico.iParam1 = tMnemonicoWFW.iParam1
        objMnemonico.iParam2 = tMnemonicoWFW.iParam2
        objMnemonico.iParam3 = tMnemonicoWFW.iParam3
        objMnemonico.sGrid = tMnemonicoWFW.sNomeGrid
        objMnemonico.sMnemonicoCombo = tMnemonicoWFW.sMnemonicoCombo
        objMnemonico.sMnemonicoDesc = tMnemonicoWFW.sMnemonicoDesc
        
        colMnemonico.Add objMnemonico
        
        'le o proximo mnemonico
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178029
        
    Loop
    
    tMnemonicoWFW.sMnemonico = String(STRING_MNEMONICOWFW_MNEMONICO, 0)
    tMnemonicoWFW.sNomeGrid = String(STRING_NOME_CONTROLE, 0)
    tMnemonicoWFW.sMnemonicoCombo = String(STRING_MNEMONICOWFW_MNEMONICOCOMBO, 0)
    tMnemonicoWFW.sMnemonicoDesc = String(STRING_MNEMONICOWFW_MNEMONICODESC, 0)
    
    
    'seleciona os mnemonicos do Modulo/Transacao passado como parametro
    lErro = Comando_Executar(lComando, "SELECT Mnemonico, Tipo, NumParam, Param1, Param2, Param3, NomeGrid, MnemonicoCombo, MnemonicoDesc FROM MnemonicoCTB WHERE Modulo=? AND Transacao=?", _
    tMnemonicoWFW.sMnemonico, tMnemonicoWFW.iTipo, tMnemonicoWFW.iNumParam, tMnemonicoWFW.iParam1, tMnemonicoWFW.iParam2, tMnemonicoWFW.iParam3, tMnemonicoWFW.sNomeGrid, tMnemonicoWFW.sMnemonicoCombo, tMnemonicoWFW.sMnemonicoDesc, sModulo, iTransacao)
    If lErro <> AD_SQL_SUCESSO Then gError 178150
    
    'le o primeiro mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178151
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objMnemonico = New ClassMnemonicoWFW
    
        objMnemonico.sModulo = sModulo
        objMnemonico.iTransacao = iTransacao
        objMnemonico.sMnemonico = tMnemonicoWFW.sMnemonico
        objMnemonico.iTipo = tMnemonicoWFW.iTipo
        objMnemonico.iNumParam = tMnemonicoWFW.iNumParam
        objMnemonico.iParam1 = tMnemonicoWFW.iParam1
        objMnemonico.iParam2 = tMnemonicoWFW.iParam2
        objMnemonico.iParam3 = tMnemonicoWFW.iParam3
        objMnemonico.sGrid = tMnemonicoWFW.sNomeGrid
        objMnemonico.sMnemonicoCombo = tMnemonicoWFW.sMnemonicoCombo
        objMnemonico.sMnemonicoDesc = tMnemonicoWFW.sMnemonicoDesc
        
        colMnemonico.Add objMnemonico
        
        'le o proximo mnemonico
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178152
        
    Loop
    
    
    Call Comando_Fechar(lComando)
    
    MnemonicoWFW_Le = SUCESSO

    Exit Function

Erro_MnemonicoWFW_Le:

    MnemonicoWFW_Le = gErr
    
    Select Case gErr
    
        Case 178026
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 178027 To 178029
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOWFW", gErr, sModulo, iTransacao)
            
        Case 178150 To 178152
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOCTB", Err, sModulo, iTransacao)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178030)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RegraWFW_Le_Transacao(ByVal objRegraWFW As ClassRegraWFW, ByVal colRegraWFW As Collection) As Long
'retorna as regras associados ao modulo/transacao passados como parametro.

Dim lErro As Long
Dim lComando As Long
Dim tRegraWFW As typeRegraWFW
Dim objRegraWFW1 As ClassRegraWFW, iSeq As Integer
    
On Error GoTo Erro_RegraWFW_Le_Transacao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178075
    
    tRegraWFW.sRegra = String(STRING_REGRAWFW_REGRA, 0)
    tRegraWFW.sEmailPara = String(STRING_REGRAWFW_EMAILPARA, 0)
    tRegraWFW.sEmailAssunto = String(STRING_REGRAWFW_EMAILASSUNTO, 0)
    tRegraWFW.sEmailMsg = String(STRING_REGRAWFW_EMAILMSG, 0)
    tRegraWFW.sAvisoMsg = String(STRING_REGRAWFW_AVISOMSG, 0)
    tRegraWFW.sLogDoc = String(STRING_REGRAWFW_LOGDOC, 0)
    tRegraWFW.sLogMsg = String(STRING_REGRAWFW_LOGMSG, 0)
    tRegraWFW.sRelModulo = String(STRING_MODULO_SIGLA, 0)
    tRegraWFW.sRelNome = String(STRING_CODIGO_RELATORIO, 0)
    tRegraWFW.sRelOpcao = String(STRING_CODIGO_OPCAO_RELATORIO, 0)
    tRegraWFW.sBrowseModulo = String(STRING_MODULO_SIGLA, 0)
    tRegraWFW.sBrowseNome = String(STRING_CODIGO_RELATORIO, 0)
    tRegraWFW.sBrowseOpcao = String(STRING_CODIGO_OPCAO_RELATORIO, 0)
    tRegraWFW.sRelSel = String(STRING_REGRAWFW_RELSEL, 0)
    tRegraWFW.sRelAnexo = String(STRING_REGRAWFW_RELANEXO, 0)
    tRegraWFW.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT CodUsuario, Item, Regra, TipoBloqueio, EmailPara, EmailAssunto, EmailMsg, AvisoMsg, LogDoc, LogMsg, DataUltExec, HoraUltExec, RelModulo, RelNome, RelOpcao, BrowseModulo, BrowseNome, BrowseOpcao, RelPorEmail, RelSel, RelAnexo FROM RegraWFW WHERE Modulo = ? AND Transacao = ? And (CodUsuario = ? OR CodUsuario = ?) ORDER BY CodUsuario, Item", _
        tRegraWFW.sCodUsuario, tRegraWFW.iItem, tRegraWFW.sRegra, tRegraWFW.iTipoBloqueio, tRegraWFW.sEmailPara, tRegraWFW.sEmailAssunto, tRegraWFW.sEmailMsg, tRegraWFW.sAvisoMsg, tRegraWFW.sLogDoc, tRegraWFW.sLogMsg, tRegraWFW.dtDataUltExec, tRegraWFW.dHoraUltExec, tRegraWFW.sRelModulo, tRegraWFW.sRelNome, tRegraWFW.sRelOpcao, tRegraWFW.sBrowseModulo, tRegraWFW.sBrowseNome, tRegraWFW.sBrowseOpcao, tRegraWFW.iRelPorEmail, tRegraWFW.sRelSel, tRegraWFW.sRelAnexo, objRegraWFW.sModulo, objRegraWFW.iTransacao, objRegraWFW.sUsuario, "")
    If lErro <> AD_SQL_SUCESSO Then gError 178076
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178077
    
    iSeq = 0
    Do While lErro = AD_SQL_SUCESSO
    
        iSeq = iSeq + 1
    
        Set objRegraWFW1 = New ClassRegraWFW
        
        objRegraWFW1.sModulo = objRegraWFW.sModulo
        objRegraWFW1.iTransacao = objRegraWFW.iTransacao
        objRegraWFW1.iItem = iSeq ' tRegraWFW.iItem
        objRegraWFW1.sUsuario = tRegraWFW.sCodUsuario 'objRegraWFW.sUsuario
        objRegraWFW1.sRegra = tRegraWFW.sRegra
        objRegraWFW1.iTipoBloqueio = tRegraWFW.iTipoBloqueio
        objRegraWFW1.sEmailPara = tRegraWFW.sEmailPara
        objRegraWFW1.sEmailAssunto = tRegraWFW.sEmailAssunto
        objRegraWFW1.sEmailMsg = tRegraWFW.sEmailMsg
        objRegraWFW1.sAvisoMsg = tRegraWFW.sAvisoMsg
        objRegraWFW1.sLogDoc = tRegraWFW.sLogDoc
        objRegraWFW1.sLogMsg = tRegraWFW.sLogMsg
        objRegraWFW1.dtDataUltExec = tRegraWFW.dtDataUltExec
        objRegraWFW1.dHoraUltExec = tRegraWFW.dHoraUltExec
        objRegraWFW1.sRelModulo = tRegraWFW.sRelModulo
        objRegraWFW1.sRelNome = tRegraWFW.sRelNome
        objRegraWFW1.sRelOpcao = tRegraWFW.sRelOpcao
        objRegraWFW1.sBrowseModulo = tRegraWFW.sBrowseModulo
        objRegraWFW1.sBrowseNome = tRegraWFW.sBrowseNome
        objRegraWFW1.sBrowseOpcao = tRegraWFW.sBrowseOpcao
        objRegraWFW1.iRelPorEmail = tRegraWFW.iRelPorEmail
        objRegraWFW1.sRelSel = tRegraWFW.sRelSel
        objRegraWFW1.sRelAnexo = tRegraWFW.sRelAnexo
    
        colRegraWFW.Add objRegraWFW1

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178078

    Loop
    
    Call Comando_Fechar(lComando)
    
    RegraWFW_Le_Transacao = SUCESSO
    
    Exit Function

Erro_RegraWFW_Le_Transacao:

    RegraWFW_Le_Transacao = gErr
    
    Select Case gErr
                
        Case 178075
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178076 To 178078
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGRAWFW", gErr, objRegraWFW.sModulo, objRegraWFW.iTransacao)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178079)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function AvisoUsuWFW_Le_Item(ByVal objRegraWFW As ClassRegraWFW) As Long
'retorna os usuarios associados a regra em questao.

Dim lErro As Long
Dim lComando As Long
Dim sCodUsuario As String
    
On Error GoTo Erro_AvisoUsuWFW_Le_Item

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178082
    
    sCodUsuario = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT CodUsuario FROM AvisoUsuWFW WHERE Modulo = ? AND Transacao = ? AND Item = ? AND (CodUsuarioOrigem = ? OR CodUsuarioOrigem = ?)", _
    sCodUsuario, objRegraWFW.sModulo, objRegraWFW.iTransacao, objRegraWFW.iItem, objRegraWFW.sUsuario, "")
    If lErro <> AD_SQL_SUCESSO Then gError 178083
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178084
    
    Do While lErro = AD_SQL_SUCESSO
    
        objRegraWFW.colUsuarios.Add sCodUsuario

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178085

    Loop
    
    Call Comando_Fechar(lComando)
    
    AvisoUsuWFW_Le_Item = SUCESSO
    
    Exit Function

Erro_AvisoUsuWFW_Le_Item:

    AvisoUsuWFW_Le_Item = gErr
    
    Select Case gErr
                
        Case 178082
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178083 To 178085
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_AVISOUSUWFW", gErr, objRegraWFW.sModulo, objRegraWFW.iTransacao, objRegraWFW.iItem)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178086)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function MnemonicoWFW_Le_Mnemonico(ByVal objMnemonico As ClassMnemonicoWFW) As Long
'retorna o mnemonico em questao

Dim lComando As Long
Dim tMnemonicoWFW As typeMnemonicoWFW
Dim lErro As Long

On Error GoTo Erro_MnemonicoWFW_Le_Mnemonico

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178114
    
    tMnemonicoWFW.sNomeGrid = String(STRING_NOME_CONTROLE, 0)
    tMnemonicoWFW.sMnemonicoCombo = String(STRING_MNEMONICOWFW_MNEMONICOCOMBO, 0)
    tMnemonicoWFW.sMnemonicoDesc = String(STRING_MNEMONICOWFW_MNEMONICODESC, 0)
    
    'seleciona os mnemonicos do Modulo/Transacao passado como parametro
    lErro = Comando_Executar(lComando, "SELECT Tipo, NumParam, Param1, Param2, Param3, NomeGrid, MnemonicoCombo, MnemonicoDesc FROM MnemonicoWFW WHERE Modulo=? AND Transacao=? AND MnemonicoCombo=?", _
    tMnemonicoWFW.iTipo, tMnemonicoWFW.iNumParam, tMnemonicoWFW.iParam1, tMnemonicoWFW.iParam2, tMnemonicoWFW.iParam3, tMnemonicoWFW.sNomeGrid, tMnemonicoWFW.sMnemonicoCombo, tMnemonicoWFW.sMnemonicoDesc, objMnemonico.sModulo, objMnemonico.iTransacao, objMnemonico.sMnemonicoCombo)
    If lErro <> AD_SQL_SUCESSO Then gError 178115
    
    'le o primeiro mnemonico
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178116
    
    If lErro = AD_SQL_SEM_DADOS Then gError 178118
    
    objMnemonico.iTipo = tMnemonicoWFW.iTipo
    objMnemonico.iNumParam = tMnemonicoWFW.iNumParam
    objMnemonico.iParam1 = tMnemonicoWFW.iParam1
    objMnemonico.iParam2 = tMnemonicoWFW.iParam2
    objMnemonico.iParam3 = tMnemonicoWFW.iParam3
    objMnemonico.sGrid = tMnemonicoWFW.sNomeGrid
    objMnemonico.sMnemonicoCombo = tMnemonicoWFW.sMnemonicoCombo
    objMnemonico.sMnemonicoDesc = tMnemonicoWFW.sMnemonicoDesc
        
    Call Comando_Fechar(lComando)
    
    MnemonicoWFW_Le_Mnemonico = SUCESSO

    Exit Function

Erro_MnemonicoWFW_Le_Mnemonico:

    MnemonicoWFW_Le_Mnemonico = gErr
    
    Select Case gErr
    
        Case 178114
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 178115, 178116
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOWFW1", gErr, objMnemonico.sModulo, objMnemonico.iTransacao, objMnemonico.sMnemonico)
            
        Case 178118
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178117)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RegraWFW_Le_Transacao1(ByVal sSigla As String, ByVal sTransacaoTela As String, ByVal sCodUsuario As String, colRegraWFW As Collection, Optional ByVal iCodigoTransacao As Integer = 0) As Long

Dim lComando As Long
Dim lErro As Long
Dim tRegraWFW As typeRegraWFW
Dim objRegraWFW1 As ClassRegraWFW, iSeq As Integer

On Error GoTo Erro_RegraWFW_Le_Transacao1
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178141
    
    tRegraWFW.sRegra = String(STRING_REGRAWFW_REGRA, 0)
    tRegraWFW.sEmailPara = String(STRING_REGRAWFW_EMAILPARA, 0)
    tRegraWFW.sEmailAssunto = String(STRING_REGRAWFW_EMAILASSUNTO, 0)
    tRegraWFW.sEmailMsg = String(STRING_REGRAWFW_EMAILMSG, 0)
    tRegraWFW.sAvisoMsg = String(STRING_REGRAWFW_AVISOMSG, 0)
    tRegraWFW.sLogDoc = String(STRING_REGRAWFW_LOGDOC, 0)
    tRegraWFW.sLogMsg = String(STRING_REGRAWFW_LOGMSG, 0)
    tRegraWFW.sRelModulo = String(STRING_MODULO_SIGLA, 0)
    tRegraWFW.sRelNome = String(STRING_CODIGO_RELATORIO, 0)
    tRegraWFW.sRelOpcao = String(STRING_CODIGO_OPCAO_RELATORIO, 0)
    tRegraWFW.sBrowseModulo = String(STRING_MODULO_SIGLA, 0)
    tRegraWFW.sBrowseNome = String(STRING_CODIGO_RELATORIO, 0)
    tRegraWFW.sBrowseOpcao = String(STRING_CODIGO_OPCAO_RELATORIO, 0)
    tRegraWFW.sRelSel = String(STRING_REGRAWFW_RELSEL, 0)
    tRegraWFW.sRelAnexo = String(STRING_REGRAWFW_RELANEXO, 0)
    tRegraWFW.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    
    If iCodigoTransacao = 0 Then
        lErro = Comando_Executar(lComando, "SELECT RegraWFW.Transacao, CodUsuario, Item, Regra, TipoBloqueio, EmailPara, EmailAssunto, EmailMsg, AvisoMsg, LogDoc, LogMsg, DataUltExec, HoraUltExec, RelModulo, RelNome, RelOpcao, BrowseModulo, BrowseNome, BrowseOpcao, RelPorEmail, RelSel, RelAnexo FROM RegraWFW, TransacaoWFW WHERE TransacaoWFW.Sigla = ? AND TransacaoWFW.Transacao = ? AND RegraWFW.Modulo = TransacaoWFW.Sigla AND TransacaoWFW.Codigo = RegraWFW.Transacao AND (RegraWFW.CodUsuario = ? OR RegraWFW.CodUsuario = ?) ORDER BY RegraWFW.CodUsuario, Item", _
        tRegraWFW.iTransacao, tRegraWFW.sCodUsuario, tRegraWFW.iItem, tRegraWFW.sRegra, tRegraWFW.iTipoBloqueio, tRegraWFW.sEmailPara, tRegraWFW.sEmailAssunto, tRegraWFW.sEmailMsg, tRegraWFW.sAvisoMsg, tRegraWFW.sLogDoc, tRegraWFW.sLogMsg, tRegraWFW.dtDataUltExec, tRegraWFW.dHoraUltExec, tRegraWFW.sRelModulo, tRegraWFW.sRelNome, tRegraWFW.sRelOpcao, tRegraWFW.sBrowseModulo, tRegraWFW.sBrowseNome, tRegraWFW.sBrowseOpcao, tRegraWFW.iRelPorEmail, tRegraWFW.sRelSel, tRegraWFW.sRelAnexo, sSigla, sTransacaoTela, sCodUsuario, "")
    Else
        lErro = Comando_Executar(lComando, "SELECT RegraWFW.Transacao, CodUsuario, Item, Regra, TipoBloqueio, EmailPara, EmailAssunto, EmailMsg, AvisoMsg, LogDoc, LogMsg, DataUltExec, HoraUltExec, RelModulo, RelNome, RelOpcao, BrowseModulo, BrowseNome, BrowseOpcao, RelPorEmail, RelSel, RelAnexo FROM RegraWFW WHERE RegraWFW.Transacao = ? AND (RegraWFW.CodUsuario = ? OR RegraWFW.CodUsuario = ?) ORDER BY RegraWFW.CodUsuario, Item", _
        tRegraWFW.iTransacao, tRegraWFW.sCodUsuario, tRegraWFW.iItem, tRegraWFW.sRegra, tRegraWFW.iTipoBloqueio, tRegraWFW.sEmailPara, tRegraWFW.sEmailAssunto, tRegraWFW.sEmailMsg, tRegraWFW.sAvisoMsg, tRegraWFW.sLogDoc, tRegraWFW.sLogMsg, tRegraWFW.dtDataUltExec, tRegraWFW.dHoraUltExec, tRegraWFW.sRelModulo, tRegraWFW.sRelNome, tRegraWFW.sRelOpcao, tRegraWFW.sBrowseModulo, tRegraWFW.sBrowseNome, tRegraWFW.sBrowseOpcao, tRegraWFW.iRelPorEmail, tRegraWFW.sRelSel, tRegraWFW.sRelAnexo, iCodigoTransacao, sCodUsuario, "")
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 178142
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178143
    
    iSeq = 0
    Do While lErro = AD_SQL_SUCESSO
    
        iSeq = iSeq + 1
    
        Set objRegraWFW1 = New ClassRegraWFW
        
        objRegraWFW1.sModulo = sSigla
        objRegraWFW1.iTransacao = tRegraWFW.iTransacao
        objRegraWFW1.iItem = iSeq 'tRegraWFW.iItem
        objRegraWFW1.sUsuario = tRegraWFW.sCodUsuario
        objRegraWFW1.sRegra = tRegraWFW.sRegra
        objRegraWFW1.iTipoBloqueio = tRegraWFW.iTipoBloqueio
        objRegraWFW1.sEmailPara = tRegraWFW.sEmailPara
        objRegraWFW1.sEmailAssunto = tRegraWFW.sEmailAssunto
        objRegraWFW1.sEmailMsg = tRegraWFW.sEmailMsg
        objRegraWFW1.sAvisoMsg = tRegraWFW.sAvisoMsg
        objRegraWFW1.sLogDoc = tRegraWFW.sLogDoc
        objRegraWFW1.sLogMsg = tRegraWFW.sLogMsg
        objRegraWFW1.dtDataUltExec = tRegraWFW.dtDataUltExec
        objRegraWFW1.dHoraUltExec = tRegraWFW.dHoraUltExec
        objRegraWFW1.sRelModulo = tRegraWFW.sRelModulo
        objRegraWFW1.sRelNome = tRegraWFW.sRelNome
        objRegraWFW1.sRelOpcao = tRegraWFW.sRelOpcao
        objRegraWFW1.sBrowseModulo = tRegraWFW.sBrowseModulo
        objRegraWFW1.sBrowseNome = tRegraWFW.sBrowseNome
        objRegraWFW1.sBrowseOpcao = tRegraWFW.sBrowseOpcao
        objRegraWFW1.iRelPorEmail = tRegraWFW.iRelPorEmail
        objRegraWFW1.sRelSel = tRegraWFW.sRelSel
        objRegraWFW1.sRelAnexo = tRegraWFW.sRelAnexo
    
        colRegraWFW.Add objRegraWFW1

        lErro = AvisoUsuWFW_Le_Item(objRegraWFW1)
        If lErro <> SUCESSO Then gError 178180

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178144

    Loop

    Call Comando_Fechar(lComando)

    RegraWFW_Le_Transacao1 = SUCESSO

    Exit Function

Erro_RegraWFW_Le_Transacao1:

    RegraWFW_Le_Transacao1 = gErr
    
    Select Case gErr
    
        Case 178141
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178142 To 178144
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGRAWFW", gErr, sSigla, sTransacaoTela)
    
        Case 178180
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178145)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function AvisoWFW_Le_Usuario(ByVal sUsuario As String, ByVal colAvisoWFW As Collection) As Long
'retorna os avisos para um determinado usuario

Dim lErro As Long
Dim lComando As Long
Dim tAvisoWFW As typeAvisoWFW
Dim objAvisoWFW As ClassAvisoWFW

On Error GoTo Erro_AvisoWFW_Le_Usuario

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178185
    
    tAvisoWFW.sMsg = String(STRING_AVISOWFW_MSG, 0)
    tAvisoWFW.sUsuarioOrig = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    tAvisoWFW.sTransacaoTela = String(STRING_TRANSACAOWFW_TRANSACAOTELA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Msg, Data, Hora, TransacaoWFW.TransacaoTela, DataUltAviso, HoraUltAviso, Intervalo, UsuarioOrig, UMIntervalo FROM AvisoWFW, TransacaoWFW WHERE AvisoWFW.Transacao = TransacaoWFW.Codigo AND Usuario = ? ORDER BY Data, Hora", _
    tAvisoWFW.lNumIntDoc, tAvisoWFW.sMsg, tAvisoWFW.dtData, tAvisoWFW.dHora, tAvisoWFW.sTransacaoTela, tAvisoWFW.dtDataUltAviso, tAvisoWFW.dHoraUltAviso, tAvisoWFW.dIntervalo, tAvisoWFW.sUsuarioOrig, tAvisoWFW.iUMIntervalo, sUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 178186
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178187
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objAvisoWFW = New ClassAvisoWFW
    
        objAvisoWFW.lNumIntDoc = tAvisoWFW.lNumIntDoc
        objAvisoWFW.sMsg = tAvisoWFW.sMsg
        objAvisoWFW.dtData = tAvisoWFW.dtData
        objAvisoWFW.dHora = tAvisoWFW.dHora
        objAvisoWFW.sTransacaoTela = tAvisoWFW.sTransacaoTela
        objAvisoWFW.dtDataUltAviso = tAvisoWFW.dtDataUltAviso
        objAvisoWFW.dHoraUltAviso = tAvisoWFW.dHoraUltAviso
        objAvisoWFW.dIntervalo = tAvisoWFW.dIntervalo
        objAvisoWFW.sUsuarioOrig = tAvisoWFW.sUsuarioOrig
        objAvisoWFW.iUMIntervalo = tAvisoWFW.iUMIntervalo
    
        colAvisoWFW.Add objAvisoWFW

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178188

    Loop
    
    Call Comando_Fechar(lComando)
    
    AvisoWFW_Le_Usuario = SUCESSO
    
    Exit Function

Erro_AvisoWFW_Le_Usuario:

    AvisoWFW_Le_Usuario = gErr
    
    Select Case gErr
                
        Case 178185
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178186 To 178188
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_AVISOWFW", gErr, sUsuario)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178189)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function AvisoWFW_Le_NumIntDoc(ByVal lNumIntDoc As Long, ByVal objAvisoWFW As ClassAvisoWFW) As Long
'retorna o aviso para um determinado NumIntDoc

Dim lErro As Long
Dim lComando As Long
Dim tAvisoWFW As typeAvisoWFW

On Error GoTo Erro_AvisoWFW_Le_NumIntDoc

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178217
    
    tAvisoWFW.sMsg = String(STRING_AVISOWFW_MSG, 0)
    tAvisoWFW.sUsuarioOrig = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    tAvisoWFW.sTransacaoTela = String(STRING_TRANSACAOWFW_TRANSACAOTELA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Msg, Data, Hora, TransacaoWFW.TransacaoTela, DataUltAviso, HoraUltAviso, Intervalo, UsuarioOrig, UMIntervalo FROM AvisoWFW, TransacaoWFW WHERE AvisoWFW.Transacao = TransacaoWFW.Codigo AND NumIntDoc = ? ORDER BY Data, Hora", _
    tAvisoWFW.lNumIntDoc, tAvisoWFW.sMsg, tAvisoWFW.dtData, tAvisoWFW.dHora, tAvisoWFW.sTransacaoTela, tAvisoWFW.dtDataUltAviso, tAvisoWFW.dHoraUltAviso, tAvisoWFW.dIntervalo, tAvisoWFW.sUsuarioOrig, tAvisoWFW.iUMIntervalo, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 178218
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178219
    
    If lErro = AD_SQL_SEM_DADOS Then gError 178220
    
    objAvisoWFW.lNumIntDoc = tAvisoWFW.lNumIntDoc
    objAvisoWFW.sMsg = tAvisoWFW.sMsg
    objAvisoWFW.dtData = tAvisoWFW.dtData
    objAvisoWFW.dHora = tAvisoWFW.dHora
    objAvisoWFW.sTransacaoTela = tAvisoWFW.sTransacaoTela
    objAvisoWFW.dtDataUltAviso = tAvisoWFW.dtDataUltAviso
    objAvisoWFW.dHoraUltAviso = tAvisoWFW.dHoraUltAviso
    objAvisoWFW.dIntervalo = tAvisoWFW.dIntervalo
    objAvisoWFW.sUsuarioOrig = tAvisoWFW.sUsuarioOrig
    objAvisoWFW.iUMIntervalo = tAvisoWFW.iUMIntervalo
    
    Call Comando_Fechar(lComando)
    
    AvisoWFW_Le_NumIntDoc = SUCESSO
    
    Exit Function

Erro_AvisoWFW_Le_NumIntDoc:

    AvisoWFW_Le_NumIntDoc = gErr
    
    Select Case gErr
                
        Case 178217
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178218, 178219
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_AVISOWFW1", gErr, lNumIntDoc)
                  
        Case 178220
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178221)
        
        End Select
        
        Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function BrowseOpcaoOrdenacao_Le_Opcoes(ByVal sNomeTela As String, ByVal colOpcoes As Collection) As Long

Dim lComando1 As Long
Dim lErro As Long
Dim sOpcao As String

On Error GoTo Erro_BrowseOpcaoOrdenacao_Le_Opcoes

    lComando1 = 0

    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then gError 178331

    sOpcao = String(STRING_BROWSER_OPCAO, 0)

    'pesquisa no BD as informacoes de ordenacao
    lErro = Comando_Executar(lComando1, "SELECT Opcao FROM BrowseOpcaoOrdenacao WHERE NomeTela=?", sOpcao, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 178332

    'le a ordenacao, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178333

    'encontrou a ordenacao ==> isto é uma alteração
    Do While lErro = AD_SQL_SUCESSO

        colOpcoes.Add sOpcao

        'le a ordenacao, se estiver no BD
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178334

    Loop

    Call Comando_Fechar(lComando1)

    BrowseOpcaoOrdenacao_Le_Opcoes = SUCESSO

    Exit Function

Erro_BrowseOpcaoOrdenacao_Le_Opcoes:

    BrowseOpcaoOrdenacao_Le_Opcoes = gErr
    
    Select Case gErr

        Case 178331
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178332, 178333, 178334
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEOPCAOORDENACAO", gErr, sNomeTela, sOpcao)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178335)

    End Select
    
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function BrowseOpcaoOrdenacao_Le(ByVal sOpcao As String, sNomeTela As String, objBrowseUsuarioOrdenacao As AdmBrowseUsuarioOrdenacao) As Long

Dim lComando1 As Long
Dim lErro As Long
Dim sOpcao1 As String
Dim tBrowseUsuarioOrdenacao As typeBrowseUsuarioOrdenacao

On Error GoTo Erro_BrowseOpcaoOrdenacao_Le

    lComando1 = 0

    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then gError 178338

    tBrowseUsuarioOrdenacao.sNomeTela = String(STRING_BROWSEUSUARIO_NOMETELA, 0)
    tBrowseUsuarioOrdenacao.sSelecaoSQL1 = String(STRING_SELECAO_SQL, 0)
    tBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario = String(STRING_SELECAO_SQL, 0)
    tBrowseUsuarioOrdenacao.sNomeIndice = String(STRING_NOME_INDICE, 0)

    'pesquisa no BD as informacoes de ordenacao
    lErro = Comando_Executar(lComando1, "SELECT NomeTela, Indice, SelecaoSQL1, SelecaoSQL1Usuario, NomeIndice FROM BrowseOpcaoOrdenacao WHERE NomeTela=? And Opcao=?", _
    tBrowseUsuarioOrdenacao.sNomeTela, tBrowseUsuarioOrdenacao.iIndice, _
    tBrowseUsuarioOrdenacao.sSelecaoSQL1, tBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario, tBrowseUsuarioOrdenacao.sNomeIndice, sNomeTela, sOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 178339

    'le a ordenacao, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178340

    If lErro <> AD_SQL_SUCESSO Then gError 178341

    objBrowseUsuarioOrdenacao.sNomeTela = tBrowseUsuarioOrdenacao.sNomeTela
    objBrowseUsuarioOrdenacao.iIndice = tBrowseUsuarioOrdenacao.iIndice
    objBrowseUsuarioOrdenacao.sSelecaoSQL1 = tBrowseUsuarioOrdenacao.sSelecaoSQL1
    objBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario = tBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario
    objBrowseUsuarioOrdenacao.sNomeIndice = tBrowseUsuarioOrdenacao.sNomeIndice

    Call Comando_Fechar(lComando1)

    BrowseOpcaoOrdenacao_Le = SUCESSO

    Exit Function

Erro_BrowseOpcaoOrdenacao_Le:

    BrowseOpcaoOrdenacao_Le = gErr
    
    Select Case gErr

        Case 178338
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178339, 178340
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEOPCAOORDENACAO", gErr, sNomeTela, sOpcao)
            
        Case 178341
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178335)

    End Select
    
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function BrowseOpcao_Le(ByVal sOpcao As String, ByVal objBrowseUsuario As AdmBrowseUsuario) As Long
'le os campos da tabela browseopcao e coloca os resultados em objBrowseUsuario

Dim lErro As Long
Dim lComando As Long
Dim tBrowseUsuario As typeBrowseUsuario
Dim colBrowseUsuario As New Collection

On Error GoTo Erro_BrowseOpcao_Le

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 178344

    tBrowseUsuario.sCodUsuario = String(STRING_BROWSEUSUARIO_CODUSUARIO, 0)
    tBrowseUsuario.sNomeTela = String(STRING_NOME_TELA, 0)

    lErro = Comando_Executar(lComando, "SELECT NomeTela, Topo, Esquerda, Largura, Altura FROM BrowseOpcao WHERE NomeTela=? AND Opcao=?", tBrowseUsuario.sNomeTela, tBrowseUsuario.lTopo, tBrowseUsuario.lEsquerda, tBrowseUsuario.lLargura, tBrowseUsuario.lAltura, objBrowseUsuario.sNomeTela, sOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 178345

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178346

    If lErro = AD_SQL_SEM_DADOS Then gError 178347
        
    objBrowseUsuario.sNomeTela = tBrowseUsuario.sNomeTela
    objBrowseUsuario.lTopo = tBrowseUsuario.lTopo
    objBrowseUsuario.lEsquerda = tBrowseUsuario.lEsquerda
    objBrowseUsuario.lLargura = tBrowseUsuario.lLargura
    objBrowseUsuario.lAltura = tBrowseUsuario.lAltura
        
    Call Comando_Fechar(lComando)
    
    BrowseOpcao_Le = SUCESSO
    
    Exit Function

Erro_BrowseOpcao_Le:

    BrowseOpcao_Le = gErr

    Select Case gErr
   
        Case 178344
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case 178345, 178346
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEOPCAO", gErr)

        Case 178347
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178348)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BrowseOpcaoCampo_Le(ByVal sOpcao As String, ByVal sNomeTela As String, ByVal colBrowseUsuarioCampo As Collection) As Long
'le todos os campos da tela de browse para a opcao especificada e coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim tBrowseUsuarioCampo As typeBrowseUsuarioCampo
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
    
On Error GoTo Erro_BrowseOpcaoCampo_Le

    tBrowseUsuarioCampo.sNome = String(STRING_NOME_CAMPO, 0)
    tBrowseUsuarioCampo.sNomeArq = String(STRING_NOME_TABELA, 0)
    tBrowseUsuarioCampo.sNomeTela = String(STRING_NOME_TELA, 0)
    tBrowseUsuarioCampo.sTitulo = String(STRING_TITULO_CAMPO, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 178349

    lErro = Comando_Executar(lComando, "SELECT NomeTela, NomeArq, Nome, PosicaoTela, Titulo, Largura FROM BrowseOpcaoCampo WHERE NomeTela=? AND Opcao=? ORDER BY PosicaoTela", tBrowseUsuarioCampo.sNomeTela, tBrowseUsuarioCampo.sNomeArq, tBrowseUsuarioCampo.sNome, tBrowseUsuarioCampo.iPosicaoTela, tBrowseUsuarioCampo.sTitulo, tBrowseUsuarioCampo.lLargura, sNomeTela, sOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 178350
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178351
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objBrowseUsuarioCampo = New AdmBrowseUsuarioCampo
    
        objBrowseUsuarioCampo.sNomeTela = tBrowseUsuarioCampo.sNomeTela
        objBrowseUsuarioCampo.sNomeArq = tBrowseUsuarioCampo.sNomeArq
        objBrowseUsuarioCampo.sNome = tBrowseUsuarioCampo.sNome
        objBrowseUsuarioCampo.iPosicaoTela = tBrowseUsuarioCampo.iPosicaoTela
        objBrowseUsuarioCampo.sTitulo = tBrowseUsuarioCampo.sTitulo
        objBrowseUsuarioCampo.lLargura = tBrowseUsuarioCampo.lLargura

        colBrowseUsuarioCampo.Add objBrowseUsuarioCampo
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178352
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    BrowseOpcaoCampo_Le = SUCESSO
    
    Exit Function
    
Erro_BrowseOpcaoCampo_Le:

    BrowseOpcaoCampo_Le = gErr

    Select Case gErr
    
        Case 178349
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 178350 To 178352
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEOPCAOCAMPO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178353)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FilialEmpresaAlmox_Customiza(iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_FilialEmpresaAlmox_Customiza

    FilialEmpresaAlmox_Customiza = SUCESSO
    
    Exit Function
    
Erro_FilialEmpresaAlmox_Customiza:

    FilialEmpresaAlmox_Customiza = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142603)
        
    End Select
    
    Exit Function

End Function

Public Function FilialEmpresa_Real(iExiste As Integer) As Long

Dim objFilialEmp As AdmFiliais

    iExiste = 0

    If Not (gcolFiliaisEmpresa Is Nothing) Then

        For Each objFilialEmp In gcolFiliaisEmpresa

            If objFilialEmp.iCodFilial > 50 And objFilialEmp.lCodEmpresa = glEmpresa Then

                iExiste = 1
                Exit For

            End If

        Next

    End If

End Function

Function Le_Campo_Tabela(ByVal sTabela As String, ByVal sCampo As String, ByVal iTipoCampo As Integer, ByVal sCampoFiltro As String, ByVal vValorCampoFiltro As Variant, vValorCampo As Variant, Optional ByVal objLabel As Object = Nothing) As Long

Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_Le_Campo_Tabela

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200977

    Select Case iTipoCampo

        Case TIPO_INT
            vValorCampo = CInt(0)

        Case TIPO_LONG
            vValorCampo = CLng(0)

        Case TIPO_STR
            vValorCampo = String(255, 0)

    End Select

    lErro = Comando_Executar(lComando, "SELECT " & sCampo & " FROM " & sTabela & " WHERE " & sCampoFiltro & " = ? ", vValorCampo, vValorCampoFiltro)
    If lErro <> AD_SQL_SUCESSO Then gError 200978

    'le o primeiro sCampo_Codigo e sCampo_Nome de sTabela
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200979

    If Not (objLabel Is Nothing) Then
        If lErro <> AD_SQL_SEM_DADOS Then
            objLabel.Caption = CStr(vValorCampo)
        Else
            objLabel.Caption = ""
        End If
    End If

    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    Call Comando_Fechar(lComando)

    Le_Campo_Tabela = SUCESSO

    Exit Function

Erro_Le_Campo_Tabela:

    Le_Campo_Tabela = gErr

    Select Case gErr

        Case 200977
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200978, 200979
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, sTabela)
            
        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200980)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function BrowseExcel_Le(ByVal objBrowse As AdmBrowse, ByVal sOpcao As String, ByVal objBrowseExcel As AdmBrowseExcel) As Long

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim sTitulo As String
Dim iTabelaDinamica As Integer
Dim iGrafico As Integer
Dim iTipoGrafico As Integer
Dim iFormula As Integer
Dim objBrowseExcelAux As AdmBrowseExcelAux
Dim iPosicao As Integer
Dim sCampo As String
Dim iFormato As Integer
Dim sLocalizacaoCSV As String
Dim sArquivo As String

On Error GoTo Erro_BrowseExcel_Le

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 202188
    Next
     
    sTitulo = String(STRING_BROWSEREXCEL_TITULO, 0)
    sLocalizacaoCSV = String(STRING_BROWSEREXCEL_LOCALIZACAOCSV, 0)
    sArquivo = String(STRING_BROWSEREXCEL_ARQUIVO, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT Titulo, TabelaDinamica, Grafico, TipoGrafico, Formato, LocalizacaoCSV, Arquivo FROM BrowseExcel WHERE NomeTela = ? AND Opcao = ?", _
    sTitulo, iTabelaDinamica, iGrafico, iTipoGrafico, iFormato, sLocalizacaoCSV, sArquivo, objBrowse.objForm.Name, sOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 202189
         
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202190
    
    If lErro <> AD_SQL_SEM_DADOS Then
    
        objBrowseExcel.iIncluirGrafico = iGrafico
        objBrowseExcel.iTabelaDinamica = iTabelaDinamica
        objBrowseExcel.iTipoGrafico = iTipoGrafico
        objBrowseExcel.sTitulo = sTitulo
        objBrowseExcel.iFormato = iFormato
        objBrowseExcel.sLocalizacaoCSV = sLocalizacaoCSV
        objBrowseExcel.sArquivo = sArquivo
        
        sCampo = String(STRING_NOME_CAMPO, 0)
    
        lErro = Comando_Executar(alComando(2), "SELECT Campo, Formula FROM BrowseExcelFormulas WHERE NomeTela = ? AND Opcao = ? ORDER BY Seq", _
        sCampo, iFormula, objBrowse.objForm.Name, sOpcao)
        If lErro <> AD_SQL_SUCESSO Then gError 202191
             
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202192
        
        Do While lErro = AD_SQL_SUCESSO
                     
            Set objBrowseExcelAux = New AdmBrowseExcelAux
            
            objBrowseExcelAux.iFormula = iFormula
            objBrowseExcelAux.sCampo = sCampo
            
            objBrowseExcel.colFormulas.Add objBrowseExcelAux
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202193
             
        Loop
        
        sCampo = String(STRING_NOME_CAMPO, 0)
    
        lErro = Comando_Executar(alComando(3), "SELECT Campo,Posicao, Formula FROM BrowseExcelCampos WHERE NomeTela = ? AND Opcao = ? ORDER BY Seq", _
        sCampo, iPosicao, iFormula, objBrowse.objForm.Name, sOpcao)
        If lErro <> AD_SQL_SUCESSO Then gError 202194
             
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202195
        
        Do While lErro = AD_SQL_SUCESSO
                     
            Set objBrowseExcelAux = New AdmBrowseExcelAux
            
            objBrowseExcelAux.iFormula = iFormula
            objBrowseExcelAux.iPosicao = iPosicao
            objBrowseExcelAux.sCampo = sCampo
            
            objBrowseExcel.colCampos.Add objBrowseExcelAux
            
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202196
             
        Loop
        
    End If
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   BrowseExcel_Le = SUCESSO
   
   Exit Function
   
Erro_BrowseExcel_Le:

   BrowseExcel_Le = gErr

    Select Case gErr

        Case 202188
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 202189 To 202196
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEEXCEL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202197)

    End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Tamanhos_Le(ByVal sStringCampo As String, iTamanho As Integer) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Tamanhos_Le

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 202221
    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT Tamanho FROM Tamanhos WHERE StringCampo = ?", _
    iTamanho, sStringCampo)
    If lErro <> AD_SQL_SUCESSO Then gError 202222
         
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202223
    
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   Tamanhos_Le = SUCESSO
   
   Exit Function
   
Erro_Tamanhos_Le:

   Tamanhos_Le = gErr

    Select Case gErr

        Case 202221
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 202222, 202223
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAMANHOS", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202224)

    End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function CNAE_Le(ByVal sCodigo As String, sDescricao As String) As Long

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_CNAE_Le

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 202232
    
    sDescricao = String(STRING_CNAE_DESCRICAO, 0)
    
    sCodigo = Replace(sCodigo, ".", "")
    sCodigo = Replace(sCodigo, "/", "")
    sCodigo = Replace(sCodigo, "\", "")
    sCodigo = Replace(sCodigo, "-", "")
    sCodigo = Replace(sCodigo, " ", "")
    sCodigo = Replace(sCodigo, "_", "")
    If Len(sCodigo) <> 7 Then sCodigo = String(7 - Len(sCodigo), "0") & sCodigo

    'busca os registros da tabela CNAE
    lErro = Comando_Executar(lComando, "SELECT Descricao FROM CNAE WHERE Codigo = ? ", sDescricao, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 202233

    'busca o 1ºregistro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202234

    'erro sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    'Fechamento do comando
    Call Comando_Fechar(lComando)

   CNAE_Le = SUCESSO

    Exit Function

Erro_CNAE_Le:

    CNAE_Le = gErr

    Select Case gErr

        Case ERRO_LEITURA_SEM_DADOS  'sem dados

        Case 202232
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)

        Case 202233, 202234
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CNAE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202235)

    End Select

    'Fechamento do comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Pais_Le_Nome(ByVal objPais As ClassPais) As Long
'Lê o Pais a partir da chave fornecida em (Nome) e retorna os dados em objPais

Dim lComando As Long
Dim lErro As Long
Dim iCodBacen As Integer
Dim iCodPais As Integer

On Error GoTo Erro_Pais_Le_Nome

    lComando = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 203264

    'Seleciona o Pais cujo Código está no Obj
    lErro = Comando_Executar(lComando, "SELECT Codigo, CodBacen FROM Paises WHERE Nome = ?", iCodPais, iCodBacen, objPais.sNome)
    If lErro <> AD_SQL_SUCESSO Then gError 203265

    'Posiciona no primeiro Pais dos que foram lidos
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203266

    'Se não encontrou nenhum Pais com o Código em questão
    If lErro = AD_SQL_SEM_DADOS Then gError 203267

    objPais.iCodigo = iCodPais
    objPais.iCodBacen = iCodBacen

    Call Comando_Fechar(lComando)

    Pais_Le_Nome = SUCESSO

    Exit Function

Erro_Pais_Le_Nome:

    Pais_Le_Nome = gErr

    Select Case gErr

        Case 203265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 203265, 203266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_PAISES", gErr)

        Case 203267

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203268)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function LocalizarGrid_Abre(ByVal objGrid As AdmGrid) As Long
    Call Chama_Tela("LocalizarGrid", objGrid)
End Function

Function Carrega_Combo(ByVal objCombo As Object, ByVal sTabela As String, ByVal sCampo1 As String, ByVal iTipo1 As Integer, ByVal sCampo2 As String, ByVal iTipo2 As Integer, Optional ByVal sFiltro As String) As Long

Dim lErro As Long
Dim lComando As Long
Dim vValorCampo1 As Variant
Dim vValorCampo2 As Variant
Dim bItemData As Boolean

On Error GoTo Erro_Carrega_Combo

    lComando = 0
    bItemData = False

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 202564
    
    Select Case iTipo1

        Case TIPO_INT
            vValorCampo1 = CInt(0)
            bItemData = True

        Case TIPO_LONG
            vValorCampo1 = CLng(0)
            bItemData = True

        Case TIPO_STR
            vValorCampo1 = String(500, 0)

    End Select
    
    Select Case iTipo2

        Case TIPO_INT
            vValorCampo2 = CInt(0)

        Case TIPO_LONG
            vValorCampo2 = CLng(0)

        Case TIPO_STR
            vValorCampo2 = String(500, 0)

    End Select

    objCombo.Clear

    lErro = Comando_Executar(lComando, "SELECT " & sCampo1 & ", " & sCampo2 & " FROM " & sTabela & IIf(Len(Trim(sFiltro)) > 0, " WHERE ", " ") & sFiltro & " ORDER BY " & sCampo1, vValorCampo1, vValorCampo2)
    If lErro <> AD_SQL_SUCESSO Then gError 202565

    'le o primeiro codigo de sTabela
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202566

    Do While lErro <> AD_SQL_SEM_DADOS

        objCombo.AddItem vValorCampo1 & SEPARADOR & vValorCampo2
        If bItemData Then objCombo.ItemData(objCombo.NewIndex) = vValorCampo1

        'le o proximo registro da tabela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202567

    Loop

    Call Comando_Fechar(lComando)

    Carrega_Combo = SUCESSO

    Exit Function

Erro_Carrega_Combo:

    Carrega_Combo = gErr

    Select Case gErr

        Case 202564
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202565 To 202567
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, sTabela)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202568)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function SpedCadastro_Le(ByVal objFilial As AdmFiliais) As Long

Dim lErro As Long
Dim lComando As Long
Dim sCodigo As String
Dim sInscricao As String
Dim objFilialAdmCad As AdmFiliaisAdmCad

On Error GoTo Erro_SpedCadastro_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 202582

    'Alocação de espaço no buffer
    sInscricao = String(STRING_SPEDCADASTRO_CODIGOCAD, 0)
    sCodigo = String(STRING_SPEDCADASTRO_CODIGOINST, 0)

    'Le a tabelaSpedCadastro
    lErro = Comando_Executar(lComando, "SELECT CodigoInst, CodigoCad FROM SpedCadastro WHERE FilialEmpresa= ? ", _
                sCodigo, sInscricao, _
                objFilial.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 202583

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202584

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objFilialAdmCad = New AdmFiliaisAdmCad

        objFilialAdmCad.iFilialEmpresa = objFilial.iCodFilial
        objFilialAdmCad.sInscricao = sInscricao
        objFilialAdmCad.sCodigo = sCodigo
        
        objFilial.colAdmCad.Add objFilialAdmCad
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202585
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    SpedCadastro_Le = SUCESSO

    Exit Function

Erro_SpedCadastro_Le:

    SpedCadastro_Le = gErr

    Select Case gErr

        Case 202582
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202583 To 202585
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDCADASTRO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202586)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function GrupoEmp_AbreComandos(ByVal iGrupoEmp As Integer, ByVal lEmpresa As Long, ByVal iFilialEmpresa As Integer, ByVal iNumCmd As Integer, colGrupoEmpAux As Collection) As Long

Dim lErro As Long, lComando As Long, lComandoDic As Long, sStringConexao As String, lConexaoAux As Long
Dim objGrupoEmpAux As ClassGrupoEmpAux, iIndice As Integer, lCodEmp As Long, iFilial As Integer, lCodEmpAnt As Long
Dim sDSN As String, iLenDSN As Integer, sParamOut As String, iLenParamOut As Integer

On Error GoTo Erro_GrupoEmp_AbreComandos

    Set colGrupoEmpAux = New Collection
    
    lComandoDic = Comando_AbrirExt(GL_lConexaoDic)
    If lComandoDic = 0 Then gError 201157
    
    'incluir a empresa corrente
    Set objGrupoEmpAux = New ClassGrupoEmpAux
    With objGrupoEmpAux
        .iGrupoEmp = iGrupoEmp
        .lEmpresa = glEmpresa
        .iFilialEmpresa = iFilialEmpresa
    End With
    For iIndice = 1 To iNumCmd
        lComando = Comando_Abrir
        If lComando = 0 Then gError 201158
        
        objGrupoEmpAux.alComando(iIndice) = lComando
    Next
    objGrupoEmpAux.iNumCmd = iNumCmd
    colGrupoEmpAux.Add objGrupoEmpAux
    
    If iGrupoEmp <> 0 Then
    
        lCodEmpAnt = 0
        lConexaoAux = 0
        
        sStringConexao = String(255, 0)
        lErro = Comando_Executar(lComandoDic, "SELECT CodEmpresa, CodFilial, StringConexao FROM GrupoEmpEmpFilAux, Empresas WHERE CodGrupoEmp = ? AND CodEmpBase = ? AND CodFilialBase = ? AND CodEmpresa <> 0 AND GrupoEmpEmpFilAux.CodEmpresa = Empresas.Codigo ORDER BY CodEmpresa, CodFilial", _
            lCodEmp, iFilial, sStringConexao, iGrupoEmp, lEmpresa, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 201159
        
        lErro = Comando_BuscarProximo(lComandoDic)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201160
        
        Do While lErro = AD_SQL_SUCESSO
        
            If lCodEmp <> lCodEmpAnt Then
            
                sDSN = sStringConexao
                iLenDSN = Len(sDSN)
                sParamOut = String(1024, 0)
                iLenParamOut = 1024
            
                lConexaoAux = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
                If lConexaoAux = 0 Then gError 201161
                    
            End If
            
            Set objGrupoEmpAux = New ClassGrupoEmpAux
            With objGrupoEmpAux
                .iGrupoEmp = iGrupoEmp
                .lEmpresa = lCodEmp
                .iFilialEmpresa = iFilial
                .lConexao = lConexaoAux
            End With
            For iIndice = 1 To iNumCmd
                lComando = Comando_AbrirExt(lConexaoAux)
                If lComando = 0 Then gError 201162
                
                objGrupoEmpAux.alComando(iIndice) = lComando
            Next
            objGrupoEmpAux.iNumCmd = iNumCmd
            colGrupoEmpAux.Add objGrupoEmpAux
            
            lErro = Comando_BuscarProximo(lComandoDic)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201163
        
        Loop
    
    End If
    
    Call Comando_Fechar(lComandoDic)
    
    GrupoEmp_AbreComandos = SUCESSO
    
    Exit Function
    
Erro_GrupoEmp_AbreComandos:

    GrupoEmp_AbreComandos = gErr

    Select Case gErr

        Case 201157, 201158, 201162
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 201159, 201160, 201163
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPOEMP", gErr)

        Case 201161
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_CONEXAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201156)

    End Select
    
    Call Comando_Fechar(lComandoDic)
    
    Exit Function

End Function

Function GrupoEmp_FechaComandos(colGrupoEmpAux As Collection) As Long

Dim lErro As Long, objGrupoEmpAux As ClassGrupoEmpAux, iIndice As Integer, lEmpresaAnt As Long

On Error GoTo Erro_GrupoEmp_FechaComandos

    If Not (colGrupoEmpAux Is Nothing) Then
    
        For Each objGrupoEmpAux In colGrupoEmpAux
        
            For iIndice = 1 To objGrupoEmpAux.iNumCmd
                Call Comando_Fechar(objGrupoEmpAux.alComando(iIndice))
                objGrupoEmpAux.alComando(iIndice) = 0
            Next
            
        Next
        
        lEmpresaAnt = 0
        For Each objGrupoEmpAux In colGrupoEmpAux
        
            If objGrupoEmpAux.lEmpresa <> glEmpresa Then
                If lEmpresaAnt <> objGrupoEmpAux.lEmpresa Then
                    Call Conexao_FecharExt(objGrupoEmpAux.lConexao)
                    objGrupoEmpAux.lConexao = 0
                    lEmpresaAnt = objGrupoEmpAux.lEmpresa
                End If
            End If
        
        Next
    
    End If
    
    GrupoEmp_FechaComandos = SUCESSO
    
    Exit Function
    
Erro_GrupoEmp_FechaComandos:

    GrupoEmp_FechaComandos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201155)

    End Select
    
    Exit Function

End Function

Function GrupoEmp_CarregaCombo(ByVal objCombo As Object) As Long

Dim lErro As Long, iCodGrupoEmp As Integer
Dim lComando As Long, sNomeRed As String

On Error GoTo Erro_GrupoEmp_CarregaCombo

    'Executa a abertura do Comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 201151
    
    objCombo.AddItem ""
    objCombo.ItemData(objCombo.NewIndex) = 0
    
    sNomeRed = String(STRING_GRUPOEMP_NOMERED, 0)

    lErro = Comando_Executar(lComando, "SELECT CodGrupoEmp, NomeReduzido FROM GrupoEmp ORDER BY CodGrupoEmp", iCodGrupoEmp, sNomeRed)
    If lErro <> AD_SQL_SUCESSO Then gError 201152
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201153
    
    Do While lErro = AD_SQL_SUCESSO
    
        objCombo.AddItem sNomeRed
        objCombo.ItemData(objCombo.NewIndex) = iCodGrupoEmp
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201154
    
    Loop
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    GrupoEmp_CarregaCombo = SUCESSO
    
    Exit Function
    
Erro_GrupoEmp_CarregaCombo:

    GrupoEmp_CarregaCombo = gErr

    Select Case gErr

        Case 201151
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 201152 To 201154
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPOEMP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201150)

    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ContribuinteSubstituto_Le(ByVal objFilial As AdmFiliais) As Long

Dim lErro As Long
Dim lComando As Long
Dim sUF As String
Dim sInscricao As String
Dim objFilialContrib As AdmFiliaisContrib

On Error GoTo Erro_ContribuinteSubstituto_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 202582

    'Alocação de espaço no buffer
    sInscricao = String(STRING_INSCR_EST, 0)
    sUF = String(STRING_ESTADO, 0)

    'Le a tabelaContribuinteSubstituto
    lErro = Comando_Executar(lComando, "SELECT UF, InscricaoEstadual FROM ContribuinteSubstituto WHERE FilialEmpresa= ? ", _
                sUF, sInscricao, _
                objFilial.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 202583

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202584

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objFilialContrib = New AdmFiliaisContrib

        objFilialContrib.iFilialEmpresa = objFilial.iCodFilial
        objFilialContrib.sInscricao = sInscricao
        objFilialContrib.sUF = sUF
        
        objFilial.colContrib.Add objFilialContrib
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202585
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ContribuinteSubstituto_Le = SUCESSO

    Exit Function

Erro_ContribuinteSubstituto_Le:

    ContribuinteSubstituto_Le = gErr

    Select Case gErr

        Case 202582
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202583 To 202585
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTRIBUINTESUBSTITUTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202586)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function NFSE_Obter_EXE(ByVal iFilialEmpresa As Integer, sNFSEEXE As String, Optional ByVal sFuncao As String = "") As Long

Dim lErro As Long, sCodIBGE As String
Dim lComando As Long
Dim sDiretorio As String
Dim lRetorno As Long

On Error GoTo Erro_NFSE_Obter_EXE

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 201188

    sCodIBGE = String(STRING_CIDADE_CODIBGE, 0)
    lErro = Comando_Executar(lComando, "SELECT CodIBGE FROM Cidades, Enderecos, FiliaisEmpresa WHERE Cidades.Descricao = Enderecos.Cidade AND FiliaisEmpresa.Endereco = Enderecos.Codigo AND FiliaisEmpresa.FilialEmpresa = ?", _
        sCodIBGE, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 201189

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201190
    
    If lErro <> SUCESSO Then gError 201191
    
    Select Case Trim(sCodIBGE)
    
        'Curitiba, Recife
        Case "4106902", "2611606"
            sNFSEEXE = "sgenfse.exe"
    
        'nota carioca
        Case "3304557"
            sNFSEEXE = "nfserj.exe"
    
        'nota belo horizonte, salvador e resende(rj)
        Case "3106200", "2927408", "3304201"
            sNFSEEXE = "nfseabrasf.exe"
        
        Case "3548708", "3518800" 'Sao Bernardo do Campo, Guarulhos (sao os municipio que usam o ginfes)
            If sFuncao = "Cancelar" Then
                sNFSEEXE = "nfsetatui2.exe"
            Else
                sNFSEEXE = "nfsetatui.exe"
            End If
        
        Case "3554003" 'Tatuí
            sNFSEEXE = "nfseabrasf2.exe"
        
        Case Else
            gError 201192
            
    End Select
    
    sDiretorio = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
    sDiretorio = left(sDiretorio, lRetorno)
            
    sNFSEEXE = sDiretorio & sNFSEEXE
            
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    NFSE_Obter_EXE = SUCESSO
    
    Exit Function
    
Erro_NFSE_Obter_EXE:

    NFSE_Obter_EXE = gErr

    Select Case gErr

        Case 201188
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 20189 To 201191
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSE_EXE", gErr)

        Case 201192
            Call Rotina_Erro(vbOKOnly, "ERRO_NFSE_NAO_IMPLEM_MUNICIPIO", gErr, sCodIBGE)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201187)

    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Config_Le(ByVal sTabela As String, ByVal sCodigo As String, ByVal iFilialEmpresa As Integer, sConteudo As String) As Long

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Config_Le
    
    'Abrir comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208276

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)
        
    'ler registo
    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM " & sTabela & " WHERE Codigo = ? AND FilialEmpresa = ?", sConteudo, sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 208277
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208278
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then
        sConteudo = ""
        gError 208279
    End If
    
    'Fechar comando
    Call Comando_Fechar(lComando)

    Config_Le = SUCESSO
    
    Exit Function

Erro_Config_Le:
        
    Config_Le = gErr
    
    Select Case gErr
    
        Case 208276
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208277, 208278
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIG", gErr)

        Case 208279
            'Call Rotina_Erro(vbOKOnly, "ERRO_CONFIG_INEXISTENTE", gErr, sCodigo, iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208280)
     
    End Select
     
    'Fechar comando
    Call Comando_Fechar(lComando)
 
    Exit Function
    
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function BackupConfig_Le(ByVal objBackupConfig As ClassBackupConfig) As Long

Dim lErro As Long
Dim lComando As Long
Dim tBackupConfig As typeBackupConfig

On Error GoTo Erro_BackupConfig_Le

    'Executa a abertura do Comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 208987

    'Alocação de espaço no buffer
    tBackupConfig.sDescricao = String(STRING_MAXIMO, 0)
    tBackupConfig.sDiretorio = String(STRING_MAXIMO, 0)
    tBackupConfig.sFTPURL = String(STRING_MAXIMO, 0)
    tBackupConfig.sFTPUsu = String(STRING_MAXIMO, 0)
    tBackupConfig.sFTPSenha = String(STRING_MAXIMO, 0)
    tBackupConfig.sFTPDir = String(STRING_MAXIMO, 0)
    tBackupConfig.sDirDownload = String(STRING_MAXIMO, 0)

    'Le a tabelaBackupConfig
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Habilitado, DataInicio, Hora, RepetirDias, " & _
                "Diretorio, IncluirDataNomeArq, DataUltBkp, DataProxBkp, Compactar, TransfFTP, FTPURL, FTPUsu, FTPSenha, FTPDir, DirDownload FROM BackupConfig WHERE Codigo= ? ", _
                tBackupConfig.lCodigo, tBackupConfig.sDescricao, tBackupConfig.iHabilitado, tBackupConfig.dtDataInicio, _
                tBackupConfig.dHora, tBackupConfig.iRepetirDias, tBackupConfig.sDiretorio, tBackupConfig.iIncluirDataNomeArq, tBackupConfig.dtDataUltBkp, _
                tBackupConfig.dtDataProxBkp, tBackupConfig.iCompactar, tBackupConfig.iTransfFTP, tBackupConfig.sFTPURL, tBackupConfig.sFTPUsu, tBackupConfig.sFTPSenha, tBackupConfig.sFTPDir, tBackupConfig.sDirDownload, _
                objBackupConfig.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 208988

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208989

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objBackupConfig.lCodigo = tBackupConfig.lCodigo
    objBackupConfig.sDescricao = tBackupConfig.sDescricao
    objBackupConfig.iHabilitado = tBackupConfig.iHabilitado
    objBackupConfig.dtDataInicio = tBackupConfig.dtDataInicio
    objBackupConfig.dHora = tBackupConfig.dHora
    objBackupConfig.iRepetirDias = tBackupConfig.iRepetirDias
    objBackupConfig.sDiretorio = tBackupConfig.sDiretorio
    objBackupConfig.iIncluirDataNomeArq = tBackupConfig.iIncluirDataNomeArq
    objBackupConfig.dtDataUltBkp = tBackupConfig.dtDataUltBkp
    objBackupConfig.dtDataProxBkp = tBackupConfig.dtDataProxBkp

    objBackupConfig.iTransfFTP = tBackupConfig.iTransfFTP
    objBackupConfig.iCompactar = tBackupConfig.iCompactar
    objBackupConfig.sFTPURL = tBackupConfig.sFTPURL
    objBackupConfig.sFTPUsu = tBackupConfig.sFTPUsu
    objBackupConfig.sFTPSenha = tBackupConfig.sFTPSenha
    objBackupConfig.sFTPDir = tBackupConfig.sFTPDir
    objBackupConfig.sDirDownload = tBackupConfig.sDirDownload

    'Fecha Comando
    Call Comando_Fechar(lComando)

    BackupConfig_Le = SUCESSO

    Exit Function

Erro_BackupConfig_Le:

    BackupConfig_Le = gErr

    Select Case gErr

        Case 208987
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208988, 208989
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BACKUPCONFIG", gErr)

        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208990)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Backup_Verifica_Habilitado(iBKPAtivo As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim iAux As Integer

On Error GoTo Erro_Backup_Verifica_Habilitado

    iBKPAtivo = DESMARCADO

    'Executa a abertura do Comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 209025

    'Le a tabelaBackupConfig
    lErro = Comando_Executar(lComando, "SELECT Habilitado FROM BackupConfig WHERE DataProxBkp <> ? AND Habilitado <> 0 ", _
                iAux, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 209026

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209027

    'Sem Dados
    If lErro <> AD_SQL_SEM_DADOS Then iBKPAtivo = MARCADO
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    Backup_Verifica_Habilitado = SUCESSO

    Exit Function

Erro_Backup_Verifica_Habilitado:

    Backup_Verifica_Habilitado = gErr

    Select Case gErr

        Case 209025
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209026, 209027
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BACKUPCONFIG", gErr)

        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209028)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Tabela_Le_Campo(ByVal sTabela As String, ByVal sCampo As String, ByVal sFiltro As String, ByVal iTipo As Integer, vValor As Variant) As Long
'Le as informações sobre o código passado

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Tabela_Le_Campo

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 125038
    
    If iTipo = TIPO_CAMPO_STRING Then If Len(vValor) < 500 Then vValor = String(500, 0)
    
    'Realiza a seleção através do Código que foi passado
    If iTipo = TIPO_CAMPO_DOUBLE Then
        lErro = Comando_Executar(lComando, "SELECT {fn CONVERT(" & sCampo & ", SQL_DOUBLE)} FROM " & sTabela & IIf(Len(Trim(sFiltro)) <> 0, " WHERE ", "") & sFiltro & "", vValor)
    Else
        lErro = Comando_Executar(lComando, "SELECT " & sCampo & " FROM " & sTabela & IIf(Len(Trim(sFiltro)) <> 0, " WHERE ", "") & sFiltro & "", vValor)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 125039
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125040
    
    If lErro <> AD_SQL_SUCESSO Then
        Select Case iTipo
            Case TIPO_CAMPO_STRING
                vValor = ""
            Case TIPO_CAMPO_DOUBLE
                vValor = 0
            Case TIPO_CAMPO_DATE
                vValor = DATA_NULA
        End Select
    End If
        
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Tabela_Le_Campo = SUCESSO
    
    Exit Function
    
Erro_Tabela_Le_Campo:

    Tabela_Le_Campo = gErr
    
    Select Case gErr

        Case 125038
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 125039, 125040
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, sTabela)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142597)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function UsuConfig_Le(ByVal sCodUsuario As String, ByVal lCodConfig As Long, sValor As String) As Long
'Le as informações sobre o código passado

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim sValorPadrao As String, sValorUsu As String

On Error GoTo Erro_UsuConfig_Le

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211161
    Next
    
    sValor = ""
    
    sValorPadrao = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT ValorPadrao FROM UsuConfig WHERE Codigo = ?", sValorPadrao, lCodConfig)
    If lErro <> AD_SQL_SUCESSO Then gError 211162
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211163
    
    If lErro <> AD_SQL_SUCESSO Then gError 211164
    
    sValorUsu = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(2), "SELECT Valor FROM UsuConfigValor WHERE CodUsuario = ? AND CodConfig = ?", sValorUsu, sCodUsuario, lCodConfig)
    If lErro <> AD_SQL_SUCESSO Then gError 211165
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211166
    
    If lErro <> AD_SQL_SUCESSO Then
        sValor = sValorPadrao 'Se não achou a configuração para o usuário pega o valor padrão
    Else
        sValor = sValorUsu
    End If
        
    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    UsuConfig_Le = SUCESSO
    
    Exit Function
    
Erro_UsuConfig_Le:

    UsuConfig_Le = gErr
    
    Select Case gErr

        Case 211161
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211162, 211163, 211165, 211166
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA1", gErr, "UsuConfig")
        
        Case 211164
            Call Rotina_Erro(vbOKOnly, "ERRO_USUCONFIG_NAO_CADASTRADO", gErr, lCodConfig)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211167)

    End Select

    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TelaUsuario_Le(ByVal objTelaUsuario As ClassTelaUsuario) As Long
'le os campos da tabela browseusuario e coloca os resultados em objBrowseUsuario

Dim lErro As Long
Dim lComando As Long
Dim tTelaUsuario As typeTelaUsuario

On Error GoTo Erro_TelaUsuario_Le

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 210535

    tTelaUsuario.sNomeTela = String(STRING_TELAUSUARIO_NOMETELA, 0)
    tTelaUsuario.sCodUsuario = String(STRING_TELAUSUARIO_CODUSUARIO, 0)

    'alterei para só retornar se houver algum controle alterado pelo modo de edicao - Jones 06/11/2013
    lErro = Comando_Executar(lComando, "SELECT NomeTela, CodUsuario, Topo, Esquerda, Largura, Altura FROM TelaUsuario T WHERE NomeTela=? AND CodUsuario=? AND ((EXISTS (SELECT * FROM EdicaoTela E WHERE E.NomeTela = T.NomeTela AND E.GrupoUsuarios = (SELECT CodGrupo FROM Usuarios WHERE CodUsuario = T.CodUsuario))) OR (EXISTS (SELECT * FROM TelaControlLargUsu AS L WHERE L.Usuario = T.CodUsuario AND L.NomeTela = T.NomeTela AND L.Fator > 0 )))", tTelaUsuario.sNomeTela, tTelaUsuario.sCodUsuario, tTelaUsuario.lTopo, tTelaUsuario.lEsquerda, tTelaUsuario.lLargura, tTelaUsuario.lAltura, objTelaUsuario.sNomeTela, objTelaUsuario.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 210536

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210537

    If lErro = AD_SQL_SEM_DADOS Then gError 210538
        
    objTelaUsuario.sNomeTela = tTelaUsuario.sNomeTela
    objTelaUsuario.sCodUsuario = tTelaUsuario.sCodUsuario
    objTelaUsuario.lTopo = tTelaUsuario.lTopo
    objTelaUsuario.lEsquerda = tTelaUsuario.lEsquerda
    objTelaUsuario.lLargura = tTelaUsuario.lLargura
    objTelaUsuario.lAltura = tTelaUsuario.lAltura
        
    Call Comando_Fechar(lComando)
    
    TelaUsuario_Le = SUCESSO
    
    Exit Function

Erro_TelaUsuario_Le:

    TelaUsuario_Le = gErr

    Select Case gErr
   
        Case 210535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case 210536, 210537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELAUSUARIO", gErr, objTelaUsuario.sNomeTela, objTelaUsuario.sCodUsuario)

        Case 210538
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210539)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function Anexos_Le(ByVal objAnexos As ClassAnexos) As Long

Dim lErro As Long
Dim lComando As Long
Dim tAnexos As typeAnexos
Dim objAnexosArq As ClassAnexosArq

On Error GoTo Erro_Anexos_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213145

    'Alocação de espaço no buffer
    tAnexos.sArquivo = String(STRING_MAXIMO, 0)
    tAnexos.sDescricao = String(STRING_MAXIMO, 0)

    'Le a tabelaAnexos
    lErro = Comando_Executar(lComando, "SELECT TipoDoc, NumIntDoc, Seq, Arquivo, Descricao FROM Anexos WHERE TipoDoc= ? AND NumIntDoc= ? ", _
                tAnexos.iTipoDoc, tAnexos.lNumIntDoc, tAnexos.iSeq, tAnexos.sArquivo, _
                tAnexos.sDescricao, _
                objAnexos.iTipoDoc, objAnexos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 213146

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213147

    Set objAnexos.colArq = New Collection
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objAnexosArq = New ClassAnexosArq

        objAnexosArq.iTipoDoc = tAnexos.iTipoDoc
        objAnexosArq.lNumIntDoc = tAnexos.lNumIntDoc
        objAnexosArq.iSeq = tAnexos.iSeq
        objAnexosArq.sArquivo = tAnexos.sArquivo
        objAnexosArq.sDescricao = tAnexos.sDescricao
    
        objAnexos.colArq.Add objAnexosArq
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213148

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Anexos_Le = SUCESSO

    Exit Function

Erro_Anexos_Le:

    Anexos_Le = gErr

    Select Case gErr

        Case 213146
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213147 To 213148
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ANEXOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213149)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function Idiomas_Le(ByVal colIdiomas As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim objIdioma As ClassIdioma
Dim tIdiomas As typeIdiomas

On Error GoTo Erro_Idiomas_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213154

    'Alocação de espaço no buffer
    tIdiomas.sDescricao = String(STRING_MAXIMO, 0)
    tIdiomas.sSigla = String(STRING_MAXIMO, 0)

    'Le a tabelaIdiomas
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Sigla, Padrao FROM Idiomas ORDER BY Codigo", _
                tIdiomas.iCodigo, tIdiomas.sDescricao, tIdiomas.sSigla, tIdiomas.iPadrao)
    If lErro <> AD_SQL_SUCESSO Then gError 213155

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213156

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objIdioma = New ClassIdioma
        
        objIdioma.iCodigo = tIdiomas.iCodigo
        objIdioma.sDescricao = tIdiomas.sDescricao
        objIdioma.sSigla = tIdiomas.sSigla
        objIdioma.iPadrao = tIdiomas.iPadrao
        
        colIdiomas.Add objIdioma
        
        'Busca Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213157
        
    Loop
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    Idiomas_Le = SUCESSO

    Exit Function

Erro_Idiomas_Le:

    Idiomas_Le = gErr

    Select Case gErr

        Case 213154
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213155 To 213157
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IDIOMAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213158)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function IdiomaTelaControles_Le(ByVal sNomeTela As String, ByVal colControles As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tControles As typeIdiomaTelaControles
Dim objControles As ClassIdiomaTelaControls

On Error GoTo Erro_IdiomaTelaControles_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213159

    'Alocação de espaço no buffer
    tControles.sNomeTabela = String(STRING_MAXIMO, 0)
    tControles.sNomeCampo = String(STRING_MAXIMO, 0)
    tControles.sNomeTela = String(STRING_MAXIMO, 0)
    tControles.sNomeControle = String(STRING_MAXIMO, 0)
    tControles.sNomeTelaExibicao = String(STRING_MAXIMO, 0)
    tControles.sNomeControleExibicao = String(STRING_MAXIMO, 0)

    'Le a tabelaIdiomaTelaControles
    lErro = Comando_Executar(lComando, "SELECT NomeTabela, NomeCampo, NomeTela, NomeControle, NomeTelaExibicao, NomeControleExibicao, ComMaxLen, ComMultiLine, " & _
                "EmGrid FROM IdiomaTelaControles WHERE NomeTela= ? ", _
                tControles.sNomeTabela, tControles.sNomeCampo, tControles.sNomeTela, tControles.sNomeControle, tControles.sNomeTelaExibicao, tControles.sNomeControleExibicao, _
                tControles.iComMaxLen, tControles.iComMultiLine, tControles.iEmGrid, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 213160

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213161

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objControles = New ClassIdiomaTelaControls

        objControles.sNomeTabela = tControles.sNomeTabela
        objControles.sNomeCampo = tControles.sNomeCampo
        objControles.sNomeTela = tControles.sNomeTela
        objControles.sNomeControle = tControles.sNomeControle
        objControles.sNomeTelaExibicao = tControles.sNomeTelaExibicao
        objControles.sNomeControleExibicao = tControles.sNomeControleExibicao
        objControles.iComMaxLen = tControles.iComMaxLen
        objControles.iComMultiLine = tControles.iComMultiLine
        objControles.iEmGrid = tControles.iEmGrid
        
        colControles.Add objControles
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213162
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    IdiomaTelaControles_Le = SUCESSO

    Exit Function

Erro_IdiomaTelaControles_Le:

    IdiomaTelaControles_Le = gErr

    Select Case gErr

        Case 213159
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213160 To 213162
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IDIOMATELACONTROLES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213163)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function IdiomaTextos_Le(ByVal objIdiomaTela As ClassIdiomaTela) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTextos As typeIdiomaTextos
Dim objTexto As ClassIdiomaTabCampoTexto
Dim objTab As ClassIdiomaTab
Dim objCampo As ClassIdiomaTabCampo

On Error GoTo Erro_IdiomaTextos_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213164
    
    For Each objTab In objIdiomaTela.colTabelas
    
        For Each objCampo In objTab.colCampos
    
            For Each objTexto In objCampo.colTextos
            
                Call objTexto.Reiniciar
    
                'Alocação de espaço no buffer
                tTextos.sTexto = String(STRING_MAXIMO, 0)
            
                'Le a tabelaIdiomaTextos
                lErro = Comando_Executar(lComando, "SELECT Seq, Texto FROM IdiomaTextos WHERE NomeTabela = ? AND NomeCampo = ? AND ChaveDocS = ? AND ChaveDocL = ? AND ChaveDocI = ? AND Idioma = ? ", _
                            tTextos.iSeq, tTextos.sTexto, _
                            objTab.sNomeTabela, objCampo.sNomeCampo, objTab.sChaveDocS, objTab.lChaveDocL, objTab.iChaveDocI, objTexto.iIdioma)
                If lErro <> AD_SQL_SUCESSO Then gError 213165
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213166
            
                'Sem Dados
                Do While lErro <> AD_SQL_SEM_DADOS
                
                    objTexto.sTexto = objTexto.sTexto & tTextos.sTexto
                
                    lErro = Comando_BuscarProximo(lComando)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213167
                
                Loop
                
            Next
        Next
    Next

    'Fecha Comando
    Call Comando_Fechar(lComando)

    IdiomaTextos_Le = SUCESSO

    Exit Function

Erro_IdiomaTextos_Le:

    IdiomaTextos_Le = gErr

    Select Case gErr

        Case 213164
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213165 To 213167
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IDIOMATEXTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213168)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function IdiomaTextos_Le_Texto(ByVal sNomeTabela As String, ByVal sNomeCampo As String, ByVal sChaveDocS As String, ByVal lChaveDocL As Long, ByVal iChaveDocI As Integer, ByVal iIdioma As Integer, sRetorno As String) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTextos As typeIdiomaTextos
Dim sVal As String

On Error GoTo Erro_IdiomaTextos_Le_Texto

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213169
   
    'Alocação de espaço no buffer
    tTextos.sTexto = String(STRING_MAXIMO, 0)

    'Le a tabelaIdiomaTextos
    lErro = Comando_Executar(lComando, "SELECT Seq, Texto FROM IdiomaTextos WHERE NomeTabela = ? AND NomeCampo = ? AND ChaveDocS = ? AND ChaveDocL = ? AND ChaveDocI = ? AND Idioma = ? ", _
                tTextos.iSeq, tTextos.sTexto, _
                sNomeTabela, sNomeCampo, sChaveDocS, lChaveDocL, iChaveDocI, iIdioma)
    If lErro <> AD_SQL_SUCESSO Then gError 213170

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213171

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        sVal = sVal & tTextos.sTexto
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213172
    
    Loop
    
    sRetorno = sVal

    'Fecha Comando
    Call Comando_Fechar(lComando)

    IdiomaTextos_Le_Texto = SUCESSO

    Exit Function

Erro_IdiomaTextos_Le_Texto:

    IdiomaTextos_Le_Texto = gErr

    Select Case gErr

        Case 213169
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213170 To 213172
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IDIOMATEXTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213173)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Rotina_Grid_Enable_BloqueiaAcesso(ByVal sTela As String, ByVal objControle As Object) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim iTipoControle As Integer
Dim sAux As String

Const TELA_CONTROLE_TIPO_ACESSO_NAO_MUDA = 0
Const TELA_CONTROLE_TIPO_ACESSO_BLOQUEIA = 1
Const TELA_CONTROLE_TIPO_ACESSO_BLOQUEIA_POR_USUARIO = 2
Const TELA_CONTROLE_TIPO_ACESSO_LIBERA_POR_USUARIO = 3

On Error GoTo Erro_Rotina_Grid_Enable_BloqueiaAcesso

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 213214
    Next
    
    lErro = Comando_Executar(alComando(0), "SELECT TipoControle FROM TelaControlBloqAcesso WHERE NomeTela = ? AND NomeControle = ?", iTipoControle, sTela, objControle.Name)
    If lErro <> AD_SQL_SUCESSO Then gError 213215
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213216
    
    If lErro = AD_SQL_SUCESSO Then
    
        Select Case iTipoControle
        
            Case TELA_CONTROLE_TIPO_ACESSO_NAO_MUDA
                'Nada faz
                
            Case TELA_CONTROLE_TIPO_ACESSO_BLOQUEIA
                objControle.Enabled = False
        
            Case TELA_CONTROLE_TIPO_ACESSO_BLOQUEIA_POR_USUARIO
            
                If objControle.Enabled Then
                
                    sAux = String(STRING_MAXIMO, 0)
        
                    lErro = Comando_Executar(alComando(1), "SELECT Usuario FROM TelaControlBloqAcessoUsu WHERE NomeTela = ? AND NomeControle = ? AND Usuario = ? ", sAux, sTela, objControle.Name, gsUsuario)
                    If lErro <> AD_SQL_SUCESSO Then gError 213217
                    
                    lErro = Comando_BuscarPrimeiro(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213218
        
                    If lErro = AD_SQL_SUCESSO Then objControle.Enabled = False
                    
                End If
        
            Case TELA_CONTROLE_TIPO_ACESSO_LIBERA_POR_USUARIO
        
                If objControle.Enabled Then
                
                    sAux = String(STRING_MAXIMO, 0)
        
                    lErro = Comando_Executar(alComando(1), "SELECT Usuario FROM TelaControlBloqAcessoUsu WHERE NomeTela = ? AND NomeControle = ? AND Usuario = ? ", sAux, sTela, objControle.Name, gsUsuario)
                    If lErro <> AD_SQL_SUCESSO Then gError 213219
                    
                    lErro = Comando_BuscarPrimeiro(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213220
        
                    If lErro <> AD_SQL_SUCESSO Then objControle.Enabled = False
                    
                End If
                
        End Select
    
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Rotina_Grid_Enable_BloqueiaAcesso = SUCESSO

    Exit Function

Erro_Rotina_Grid_Enable_BloqueiaAcesso:

    Rotina_Grid_Enable_BloqueiaAcesso = gErr

    Select Case gErr

        Case 213214
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213215 To 213216
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELACONTROLBLOQACESSO", gErr)

        Case 213217 To 213220
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELACONTROLBLOQACESSOUSU", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213221)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Telas_Trata_NomeExibicao(ByVal objTela As Object) As Long

Dim lErro As Long
Dim lComando As Long
Dim sNomeExibicao As String

On Error GoTo Erro_Telas_Trata_NomeExibicao

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 213230

    sNomeExibicao = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(lComando, "SELECT NomeExibicao FROM Telas WHERE Nome=? AND NomeExibicao <> ?", sNomeExibicao, objTela.Name, "")
    If lErro <> AD_SQL_SUCESSO Then gError 213231

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213232

    If lErro <> AD_SQL_SEM_DADOS Then
        objTela.Caption = sNomeExibicao
    End If
        
    Call Comando_Fechar(lComando)
    
    Telas_Trata_NomeExibicao = SUCESSO
    
    Exit Function

Erro_Telas_Trata_NomeExibicao:

    Telas_Trata_NomeExibicao = gErr

    Select Case gErr
   
        Case 213230
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case 213231, 213232
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELAS", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213233)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MenuItens_Trata_NomeExibicao(ByVal objPrinc As Object) As Long

Dim lErro As Long
Dim lComando As Long
Dim sNomeTela As String, sNomeExibicao As String
Dim sNomeControle As String, iIndiceControle As Integer
Dim objObjeto As Object, iIndice As Integer

On Error GoTo Erro_MenuItens_Trata_NomeExibicao

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 213234

    sNomeExibicao = String(STRING_MAXIMO, 0)
    sNomeControle = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(lComando, "SELECT NomeControle, IndiceControle, NomeExibicao FROM MenuItens WHERE NomeExibicao <> ? ", sNomeControle, iIndiceControle, sNomeExibicao, "")
    If lErro <> AD_SQL_SUCESSO Then gError 213235

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213236

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objObjeto = objPrinc.Controls(sNomeControle)
        If iIndiceControle <> 0 Then
            objObjeto(iIndiceControle).Caption = sNomeExibicao
        Else
            objObjeto.Caption = sNomeExibicao
        End If
                    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213237
                    
    Loop
        
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    MenuItens_Trata_NomeExibicao = SUCESSO
    
    Exit Function

Erro_MenuItens_Trata_NomeExibicao:

    MenuItens_Trata_NomeExibicao = gErr

    Select Case gErr
   
        Case 213234
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case 213235 To 213237
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELAS", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213238)
        
    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ConfigOutros_Le(ByVal sModulo As String, ByVal colConfig As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tConfigOutros As typeConfigOutros
Dim objConfigOutros As ClassConfigOutros
Dim sConteudo As String

On Error GoTo Erro_ConfigOutros_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213357

    'Alocação de espaço no buffer
    tConfigOutros.sTabela = String(STRING_MAXIMO, 0)
    tConfigOutros.sCodigo = String(STRING_MAXIMO, 0)
    tConfigOutros.sModuloExibicao = String(STRING_MAXIMO, 0)
    tConfigOutros.sNomeObj = String(STRING_MAXIMO, 0)
    tConfigOutros.sNomeProperty = String(STRING_MAXIMO, 0)
    tConfigOutros.sDescricaoGrid = String(STRING_MAXIMO, 0)
    tConfigOutros.sTipoControle = String(STRING_MAXIMO, 0)

    'Le a tabelaConfigOutros
    lErro = Comando_Executar(lComando, "SELECT Tabela, Codigo, ModuloExibicao, NomeObj, NomeProperty, DescricaoGrid, " & _
                "TipoControle, CodVlrValidos FROM ConfigOutros WHERE ModuloExibicao = ? ORDER BY OrdemExibicao", _
                tConfigOutros.sTabela, tConfigOutros.sCodigo, tConfigOutros.sModuloExibicao, tConfigOutros.sNomeObj, _
                tConfigOutros.sNomeProperty, tConfigOutros.sDescricaoGrid, tConfigOutros.sTipoControle, tConfigOutros.lCodVlrValidos, _
                sModulo)
    If lErro <> AD_SQL_SUCESSO Then gError 213358

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213359

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objConfigOutros = New ClassConfigOutros

        objConfigOutros.sTabela = tConfigOutros.sTabela
        objConfigOutros.sCodigo = tConfigOutros.sCodigo
        objConfigOutros.sModuloExibicao = tConfigOutros.sModuloExibicao
        objConfigOutros.sNomeObj = tConfigOutros.sNomeObj
        objConfigOutros.sNomeProperty = tConfigOutros.sNomeProperty
        objConfigOutros.sDescricaoGrid = tConfigOutros.sDescricaoGrid
        objConfigOutros.sTipoControle = tConfigOutros.sTipoControle
        objConfigOutros.lCodVlrValidos = tConfigOutros.lCodVlrValidos
        objConfigOutros.iFilialEmpresa = giFilialEmpresa
        
        lErro = Config_Le(objConfigOutros.sTabela, objConfigOutros.sCodigo, objConfigOutros.iFilialEmpresa, sConteudo)
        If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
        
        If lErro <> SUCESSO And giFilialEmpresa <> EMPRESA_TODA Then
        
            objConfigOutros.iFilialEmpresa = EMPRESA_TODA
            lErro = Config_Le(objConfigOutros.sTabela, objConfigOutros.sCodigo, objConfigOutros.iFilialEmpresa, sConteudo)
            If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        'Se não achou o Codigo na Tabela -> Ignora esse registro
        If lErro = SUCESSO Then
        
            objConfigOutros.sConteudo = sConteudo
            objConfigOutros.sConteudoNovo = sConteudo
            
            lErro = ConfigValoresValidos_Le(objConfigOutros)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            colConfig.Add objConfigOutros
            
        End If

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213360

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ConfigOutros_Le = SUCESSO

    Exit Function

Erro_ConfigOutros_Le:

    ConfigOutros_Le = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 213357
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213358 To 213360
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGOUTROS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213361)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ConfigValoresValidos_Le(ByVal objConfigOutros As ClassConfigOutros) As Long

Dim lErro As Long
Dim lComando As Long
Dim tConfigValoresValidos As typeConfigValoresValidos
Dim objConfigValoresValidos As ClassConfigValoresValidos
Dim bAchou As Boolean

On Error GoTo Erro_ConfigValoresValidos_Le

    If objConfigOutros.lCodVlrValidos > 0 Then

        'Executa a abertura do Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 213362
    
        'Alocação de espaço no buffer
        tConfigValoresValidos.sValor = String(STRING_MAXIMO, 0)
        tConfigValoresValidos.sDescricao = String(STRING_MAXIMO, 0)
    
        'Le a tabelaConfigValoresValidos
        lErro = Comando_Executar(lComando, "SELECT Codigo, Seq, Valor, Descricao FROM ConfigValoresValidos WHERE Codigo= ? ", _
                    tConfigValoresValidos.lCodigo, tConfigValoresValidos.iSeq, tConfigValoresValidos.sValor, tConfigValoresValidos.sDescricao, _
                    objConfigOutros.lCodVlrValidos)
        If lErro <> AD_SQL_SUCESSO Then gError 213363
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213364
    
        bAchou = False
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objConfigValoresValidos = New ClassConfigValoresValidos
    
            objConfigValoresValidos.lCodigo = tConfigValoresValidos.lCodigo
            objConfigValoresValidos.iSeq = tConfigValoresValidos.iSeq
            objConfigValoresValidos.sValor = tConfigValoresValidos.sValor
            objConfigValoresValidos.sDescricao = tConfigValoresValidos.sDescricao
            
            objConfigOutros.colValores.Add objConfigValoresValidos
            
            If UCase(Trim(objConfigValoresValidos.sValor)) = UCase(Trim(objConfigOutros.sConteudo)) Then
                bAchou = True
            End If
                
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213365
            
        Loop
        
        If Not bAchou And objConfigOutros.colValores.Count > 0 Then
        
            Set objConfigValoresValidos = New ClassConfigValoresValidos
    
            objConfigValoresValidos.lCodigo = objConfigOutros.lCodVlrValidos
            objConfigValoresValidos.iSeq = 0
            objConfigValoresValidos.sValor = objConfigOutros.sConteudo
            objConfigValoresValidos.sDescricao = "Outro"
            
            objConfigOutros.colValores.Add objConfigValoresValidos
            
        End If

        'Fecha Comando
        Call Comando_Fechar(lComando)
        
    End If

    ConfigValoresValidos_Le = SUCESSO

    Exit Function

Erro_ConfigValoresValidos_Le:

    ConfigValoresValidos_Le = gErr

    Select Case gErr

        Case 213362
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213363 To 213365
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGVALORESVALIDOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213366)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function Avisos_Obtem_Status(iNumAvisosNovos As Integer, iForcaAberturaTela As Integer) As Long

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Avisos_Obtem_Status

    'Executa a abertura do Comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 213513

    'Le a tabelaAvisos
    lErro = Comando_Executar(lComando, "SELECT COUNT(*), MAX(ForcaAberturaTela) FROM Avisos WHERE Codigo NOT IN (SELECT CodAviso FROM AvisosUsu WHERE CodUsuario = ?) ", _
                iNumAvisosNovos, iForcaAberturaTela, gsUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 213514

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213515

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Avisos_Obtem_Status = SUCESSO

    Exit Function

Erro_Avisos_Obtem_Status:

    Avisos_Obtem_Status = gErr

    Select Case gErr

        Case 213513
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213514, 213515
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_AVISOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213516)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function Avisos_Le(ByVal colAvisos As Collection) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim tAvisos As typeAvisos
Dim tAvisosUsu As typeAvisosUsu
Dim objAviso As ClassAvisos

On Error GoTo Erro_Avisos_Le

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 213517
    Next

    'Alocação de espaço no buffer
    tAvisos.sAssunto = String(STRING_MAXIMO, 0)
    tAvisos.sLink = String(STRING_MAXIMO, 0)

    'Le a tabela Avisos
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, Data, Assunto, Link, Prioridade, ForcaAberturaTela FROM Avisos ORDER BY Codigo DESC", _
                tAvisos.lCodigo, tAvisos.dtData, tAvisos.sAssunto, tAvisos.sLink, tAvisos.iPrioridade, tAvisos.iForcaAberturaTela)
    If lErro <> AD_SQL_SUCESSO Then gError 213518

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213519

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objAviso = New ClassAvisos

        objAviso.lCodigo = tAvisos.lCodigo
        objAviso.dtData = tAvisos.dtData
        objAviso.sAssunto = tAvisos.sAssunto
        objAviso.sLink = tAvisos.sLink
        objAviso.iPrioridade = tAvisos.iPrioridade
        objAviso.iForcaAberturaTela = tAvisos.iForcaAberturaTela
            
        lErro = Comando_Executar(alComando(2), "SELECT Lido, Excluido FROM AvisosUsu WHERE CodAviso = ? AND CodUsuario = ?", tAvisosUsu.iLido, tAvisosUsu.iExcluido, objAviso.lCodigo, gsUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 213520
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213521
        
        If lErro <> AD_SQL_SUCESSO Then
            objAviso.iNovo = MARCADO
        Else
            objAviso.iExcluido = tAvisosUsu.iExcluido
            objAviso.iLido = tAvisosUsu.iLido
        End If
        
        If objAviso.iExcluido = DESMARCADO Then
            colAvisos.Add objAviso
        End If
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213522
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Avisos_Le = SUCESSO

    Exit Function

Erro_Avisos_Le:

    Avisos_Le = gErr

    Select Case gErr

        Case 213517
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213518 To 213522
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_AVISOS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213523)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TelaControlLarg_Le(ByVal sNomeTela As String, ByVal colControles As AdmColCampoIndiceValor) As Long

Dim lErro As Long
Dim lComando As Long
Dim sNomeControle As String
Dim iIndiceControle As Integer
Dim iBase As Integer

On Error GoTo Erro_TelaControlLarg_Le

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 213716

    sNomeControle = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(lComando, "SELECT NomeControle, IndiceControle, GridBase FROM TelaControlLarg WHERE NomeTela=? ", sNomeControle, iIndiceControle, iBase, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 213717

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213718

    Do While lErro <> AD_SQL_SEM_DADOS
    
        colControles.Add sNomeControle, iBase, iIndiceControle
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213719
    
    Loop
        
    Call Comando_Fechar(lComando)
    
    TelaControlLarg_Le = SUCESSO
    
    Exit Function

Erro_TelaControlLarg_Le:

    TelaControlLarg_Le = gErr

    Select Case gErr
   
        Case 213716
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case 213717 To 213719
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "TelaControlLarg com NomeTela = " & sNomeTela)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213720)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TelaControlLargUsu_Le_Fator(ByVal sNomeTela As String, iFator As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim iFatorAux As Integer

On Error GoTo Erro_TelaControlLargUsu_Le_Fator

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 213721

    lErro = Comando_Executar(lComando, "SELECT Fator FROM TelaControlLargUsu WHERE NomeTela=? AND Usuario = ?", iFatorAux, sNomeTela, gsUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 213722

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213723

    iFator = iFatorAux
           
    Call Comando_Fechar(lComando)
    
    TelaControlLargUsu_Le_Fator = SUCESSO
    
    Exit Function

Erro_TelaControlLargUsu_Le_Fator:

    TelaControlLargUsu_Le_Fator = gErr

    Select Case gErr
   
        Case 213721
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case 213722, 213723
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "TelaControlLargUsu com NomeTela = " & sNomeTela & " e Usuário = " & gsUsuario)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213724)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Controle_Le(ByVal lCodigo As Long, sConteudo As String) As Long

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Controle_Le

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 216162

    sConteudo = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM Controle WHERE Codigo=?", sConteudo, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 216163

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216164

    If lErro = AD_SQL_SEM_DADOS Then sConteudo = ""
        
    Call Comando_Fechar(lComando)
    
    Controle_Le = SUCESSO
    
    Exit Function

Erro_Controle_Le:

    Controle_Le = gErr

    Select Case gErr
   
        Case 216162
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case 216163, 216164
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTROLE", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216165)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function
