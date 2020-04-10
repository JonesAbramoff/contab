VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private avCampo(1 To BROWSER_NUM_MAX_CAMPOS) As Variant
Private avCampo1(1 To BROWSER_NUM_MAX_CAMPOS) As Variant
Private lNumReg As Long
Private iExercicio As Integer
Private iPeriodo As Integer
Private sNomeExterno As String
Private mvarlComando As Long
Private mvarlComando1 As Long
Private mvarlOldScrollValue As Long
Private mvarlPosicao As Long
Private mvariTeclado As Integer
Private mvarcolBrowseUsuarioCampo As Collection
Private mvarsSQL As String
Private mvarcolBrowseIndiceUsuario As Collection
Private mvarcolBrowseIndice As Collection
Private mvarsCodUsuario As String
Private mvarcolValorCampo As Collection
Private mvarsNomeTabela As String
Private mvariSelecaoExterna As Integer
Private mvarcolBrowseValorCampo As Collection
Private mvarobjGrid As Object
Private mvarobjEvento As AdmEvento
Private mvarobjBrowseUsuarioOrdenacao As AdmBrowseUsuarioOrdenacao
Private mvarobjForm As Object
Private mvarobjOrdenacao As ComboBox
Private mvarobjBarraScroll As VScrollBar
Private mvariGrid_LinhaInicial As Integer
Private mvariGrid_LinhaFinal As Integer
Private mvariGrid_LinhasExibidas As Integer
Private mvariAlterado As Integer
Private mvarcolSelecao As Collection
Private mvarsSelecaoSQL As String
Private mvarsSelecaoSQL1 As String 'formato ODBC
Private mvarsSelecaoSQL2 As String 'contem a parte da selecao passada como parametro pelo programador
Private mvarsSelecaoSQL3 As String 'contem a salva da parte da selecao passada como parametro pelo programador. Se este campo for diferente de  sSelecaoSQL2 e ibind = 0 ==> ibind vai ser recolocado com 1 para fazer o bind pois a expressao SQL mudou
Private mvarsSelecaoSQL1Usuario As String 'formato visivel p/o usuario
Private mvarsSQLCount As String
Private mvarobjChamaTela As AdmChamaTela
Private mvarobjRotinaModulo As Object
Private mvarlComando2 As Long
Private mvarlComando3 As Long
Private mvariBind As Integer
Private mvariPesquisa As Integer 'indica que a tela de configuracao foi chamada atraves do botao pesquisa
Private mvariBancoDados As Integer 'indica o banco de dados que está sendo usado. BD_ADM=contab97 BD_DICDADOS=dicdados
Private mvariAlteradoOrdenacao As Integer
Private mvarobjBrowseArquivo As AdmBrowseArquivo
Private mvarcolBrowseCampo As Collection
Private mvarobjPrincMDIChild As Object
Private mvarcolRegistros As Collection
Private mvarsOpcao As String
Private mvarobjComboOpcoes As Object
Private mvarcolReg As Collection 'é a colecao atual dos registros exibidos no browser
Private mvarobjBrowseExcel As New AdmBrowseExcel

Public Property Set objBrowseExcel(ByVal vData As Object)
    Set mvarobjBrowseExcel = vData
End Property

Public Property Get objBrowseExcel() As AdmBrowseExcel
    Set objBrowseExcel = mvarobjBrowseExcel
End Property

Public Property Set colReg(ByVal vData As Object)
    Set mvarcolReg = vData
End Property

Public Property Get colReg() As Collection
    Set colReg = mvarcolReg
End Property

Public Property Set objComboOpcoes(ByVal vData As Object)
    Set mvarobjComboOpcoes = vData
End Property

Public Property Get objComboOpcoes() As Object
    Set objComboOpcoes = mvarobjComboOpcoes
End Property

Public Property Let sOpcao(ByVal vData As String)
    mvarsOpcao = vData
End Property

Public Property Get sOpcao() As String
    sOpcao = mvarsOpcao
End Property

Public Property Set colRegistros(ByVal vData As Object)
    Set mvarcolRegistros = vData
End Property

Public Property Get colRegistros() As Collection
    Set colRegistros = mvarcolRegistros
End Property

Public Property Set objPrincMDIChild(ByVal vData As Object)
    Set mvarobjPrincMDIChild = vData
End Property

Public Property Get objPrincMDIChild() As Object
    Set objPrincMDIChild = mvarobjPrincMDIChild
End Property

Public Property Set colBrowseCampo(ByVal vData As Object)
    Set mvarcolBrowseCampo = vData
End Property

Public Property Get colBrowseCampo() As Collection
    Set colBrowseCampo = mvarcolBrowseCampo
End Property

Public Property Set objBrowseArquivo(ByVal vData As Object)
    Set mvarobjBrowseArquivo = vData
End Property

Public Property Get objBrowseArquivo() As AdmBrowseArquivo
    Set objBrowseArquivo = mvarobjBrowseArquivo
End Property

Public Property Let iAlteradoOrdenacao(ByVal vData As Integer)
    mvariAlteradoOrdenacao = vData
End Property

Public Property Get iAlteradoOrdenacao() As Integer
    iAlteradoOrdenacao = mvariAlteradoOrdenacao
End Property

Public Property Let iBancoDados(ByVal vData As Integer)
    mvariBancoDados = vData
End Property

Public Property Get iBancoDados() As Integer
    iBancoDados = mvariBancoDados
End Property

Public Property Let iPesquisa(ByVal vData As Integer)
    mvariPesquisa = vData
End Property

Public Property Get iPesquisa() As Integer
    iPesquisa = mvariPesquisa
End Property

Public Property Let iBind(ByVal vData As Integer)
Attribute iBind.VB_Description = "indica se deve ocorrer o bind na rotina ... ou nao 1=Bind 0=Nao bind"
    mvariBind = vData
End Property

Public Property Get iBind() As Integer
    iBind = mvariBind
End Property

Public Property Let lComando3(ByVal vData As Long)
    mvarlComando3 = vData
End Property

Public Property Get lComando3() As Long
    lComando3 = mvarlComando3
End Property

Public Property Let lComando2(ByVal vData As Long)
    mvarlComando2 = vData
End Property

Public Property Get lComando2() As Long
    lComando2 = mvarlComando2
End Property

Public Property Set objRotinaModulo(ByVal vData As Object)
    Set mvarobjRotinaModulo = vData
End Property

Public Property Get objRotinaModulo() As Object
    Set objRotinaModulo = mvarobjRotinaModulo
End Property

Public Property Set objChamaTela(ByVal vData As Object)
    Set mvarobjChamaTela = vData
End Property

Public Property Get objChamaTela() As AdmChamaTela
    Set objChamaTela = mvarobjChamaTela
End Property

Public Property Let sSQLCount(ByVal vData As String)
Attribute sSQLCount.VB_Description = "contém o comando SQL que devolve o numero de registros de uma tabela"
    mvarsSQLCount = vData
End Property

Public Property Get sSQLCount() As String
    sSQLCount = mvarsSQLCount
End Property

Public Property Let sSelecaoSQL(ByVal vData As String)
Attribute sSelecaoSQL.VB_Description = "Contém a condição de seleção SQL, se houver"
    mvarsSelecaoSQL = vData
End Property

Public Property Get sSelecaoSQL() As String
    sSelecaoSQL = mvarsSelecaoSQL
End Property

Public Property Let sSelecaoSQL2(ByVal vData As String)
    mvarsSelecaoSQL2 = vData
End Property

Public Property Get sSelecaoSQL2() As String
    sSelecaoSQL2 = mvarsSelecaoSQL2
End Property

Public Property Let sSelecaoSQL3(ByVal vData As String)
    mvarsSelecaoSQL3 = vData
End Property

Public Property Get sSelecaoSQL3() As String
    sSelecaoSQL3 = mvarsSelecaoSQL3
End Property

Public Property Let sSelecaoSQL1(ByVal vData As String)
    mvarsSelecaoSQL1 = vData
End Property

Public Property Get sSelecaoSQL1() As String
    sSelecaoSQL1 = mvarsSelecaoSQL1
End Property

Public Property Let sSelecaoSQL1Usuario(ByVal vData As String)
    mvarsSelecaoSQL1Usuario = vData
End Property

Public Property Get sSelecaoSQL1Usuario() As String
    sSelecaoSQL1Usuario = mvarsSelecaoSQL1Usuario
End Property

Public Property Set colSelecao(ByVal vData As Object)
Attribute colSelecao.VB_Description = "Coleção dos valores que serão usados para montar a condição de seleção do comando SQL"
    Set mvarcolSelecao = vData
End Property

Public Property Get colSelecao() As Collection
    Set colSelecao = mvarcolSelecao
End Property

Public Property Let iAlterado(ByVal vData As Integer)
Attribute iAlterado.VB_Description = "1=indica que houve alguma alteração na tela de browse que deve ser salva."
    mvariAlterado = vData
End Property

Public Property Get iAlterado() As Integer
    iAlterado = mvariAlterado
End Property

Public Property Let iGrid_LinhasExibidas(ByVal vData As Integer)
Attribute iGrid_LinhasExibidas.VB_Description = "Número de Linhas do grid que exibem registros"
    mvariGrid_LinhasExibidas = vData
End Property

Public Property Get iGrid_LinhasExibidas() As Integer
    iGrid_LinhasExibidas = mvariGrid_LinhasExibidas
End Property

Public Property Let iGrid_LinhaFinal(ByVal vData As Integer)
Attribute iGrid_LinhaFinal.VB_Description = "Linha Final do Grid"
    mvariGrid_LinhaFinal = vData
End Property

Public Property Get iGrid_LinhaFinal() As Integer
    iGrid_LinhaFinal = mvariGrid_LinhaFinal
End Property

Public Property Let iGrid_LinhaInicial(ByVal vData As Integer)
Attribute iGrid_LinhaInicial.VB_Description = "Linha Inicial do grid"
    mvariGrid_LinhaInicial = vData
End Property

Public Property Get iGrid_LinhaInicial() As Integer
    iGrid_LinhaInicial = mvariGrid_LinhaInicial
End Property

Public Property Set objBarraScroll(ByVal vData As Object)
    Set mvarobjBarraScroll = vData
End Property

Public Property Get objBarraScroll() As Object
    Set objBarraScroll = mvarobjBarraScroll
End Property

Public Property Set objOrdenacao(ByVal vData As Object)
    Set mvarobjOrdenacao = vData
End Property

Public Property Get objOrdenacao() As Object
    Set objOrdenacao = mvarobjOrdenacao
End Property

Public Property Set objForm(ByVal vData As Object)
    Set mvarobjForm = vData
End Property

Public Property Get objForm() As Object
    Set objForm = mvarobjForm
End Property

Public Property Set objBrowseUsuarioOrdenacao(ByVal vData As Object)
Attribute objBrowseUsuarioOrdenacao.VB_Description = "contém as informações relativas a ordenacao selecionada pelo usuario"
    Set mvarobjBrowseUsuarioOrdenacao = vData
End Property

Public Property Get objBrowseUsuarioOrdenacao() As AdmBrowseUsuarioOrdenacao
    Set objBrowseUsuarioOrdenacao = mvarobjBrowseUsuarioOrdenacao
End Property

Public Property Set objEvento(ByVal vData As Object)
    Set mvarobjEvento = vData
End Property

Public Property Get objEvento() As AdmEvento
    Set objEvento = mvarobjEvento
End Property

Public Property Set objGrid(ByVal vData As Object)
Attribute objGrid.VB_Description = "guarda o handl do grid"
    Set mvarobjGrid = vData
End Property

Public Property Get objGrid() As Object
    Set objGrid = mvarobjGrid
End Property

Public Property Set colBrowseValorCampo(ByVal vData As Object)
Attribute colBrowseValorCampo.VB_Description = "Contém os  campos (nomes e valores)  da tela de edição associada ao browse"
    Set mvarcolBrowseValorCampo = vData
End Property

Public Property Get colBrowseValorCampo() As Collection
    Set colBrowseValorCampo = mvarcolBrowseValorCampo
End Property

Public Property Let iSelecaoExterna(ByVal vData As Integer)
Attribute iSelecaoExterna.VB_Description = "1=indica que foi feita uma seleção através de uma tela externa. 0=caso contrário"
    mvariSelecaoExterna = vData
End Property

Public Property Get iSelecaoExterna() As Integer
    iSelecaoExterna = mvariSelecaoExterna
End Property

Public Property Let sNomeTabela(ByVal vData As String)
Attribute sNomeTabela.VB_Description = "Nome da Tabela sendo listada"
    mvarsNomeTabela = vData
End Property

Public Property Get sNomeTabela() As String
    sNomeTabela = mvarsNomeTabela
End Property

Public Property Set colValorCampo(ByVal vData As Object)
Attribute colValorCampo.VB_Description = "contém o tipo, posição dos campos usados no browse"
    Set mvarcolValorCampo = vData
End Property

Public Property Get colValorCampo() As Collection
    Set colValorCampo = mvarcolValorCampo
End Property

Public Property Let sCodUsuario(ByVal vData As String)
Attribute sCodUsuario.VB_Description = "Contem o código do usuário"
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Set colBrowseIndice(ByVal vData As Object)
Attribute colBrowseIndice.VB_Description = "contém todos os índices usados pelo browse"
    Set mvarcolBrowseIndice = vData
End Property

Public Property Get colBrowseIndice() As Collection
    Set colBrowseIndice = mvarcolBrowseIndice
End Property

Public Property Set colBrowseIndiceUsuario(ByVal vData As Object)
    Set mvarcolBrowseIndiceUsuario = vData
End Property

Public Property Get colBrowseIndiceUsuario() As Collection
    Set colBrowseIndiceUsuario = mvarcolBrowseIndiceUsuario
End Property


Public Property Let sSQL(ByVal vData As String)
Attribute sSQL.VB_Description = "Contém o comando SQL usado para exibir o browse"
    mvarsSQL = vData
End Property

Public Property Get sSQL() As String
    sSQL = mvarsSQL
End Property

Public Property Set colBrowseUsuarioCampo(ByVal vData As Object)
Attribute colBrowseUsuarioCampo.VB_Description = "armazena os campos selecionados pelo usuário que serão exibidos pelo browse"
    Set mvarcolBrowseUsuarioCampo = vData
End Property

Public Property Get colBrowseUsuarioCampo() As Collection
    Set colBrowseUsuarioCampo = mvarcolBrowseUsuarioCampo
End Property

Public Property Let iTeclado(ByVal vData As Integer)
Attribute iTeclado.VB_Description = "Indice se está usando o teclado para se movimentar no browse ou o mouse. 1=teclado 0=mouse"
    mvariTeclado = vData
End Property

Public Property Get iTeclado() As Integer
    iTeclado = mvariTeclado
End Property

Public Property Let lPosicao(ByVal vData As Long)
Attribute lPosicao.VB_Description = "Armazena a  posição a partir da qual os registros são listados."
    mvarlPosicao = vData
End Property

Public Property Get lPosicao() As Long
    lPosicao = mvarlPosicao
End Property

Public Property Let lOldScrollValue(ByVal vData As Long)
Attribute lOldScrollValue.VB_Description = "Armazena o valor da posição da scroll bar do browse"
    mvarlOldScrollValue = vData
End Property

Public Property Get lOldScrollValue() As Long
    lOldScrollValue = mvarlOldScrollValue
End Property

Public Property Let lComando1(ByVal vData As Long)
Attribute lComando1.VB_Description = "Armazena o handle do comando que consegue o número de registros da tabela que está sendo listada"
    mvarlComando1 = vData
End Property

Public Property Get lComando1() As Long
    lComando1 = mvarlComando1
End Property

Public Property Let lComando(ByVal vData As Long)
Attribute lComando.VB_Description = "armazena o handle do comando que lista os registros da tabela em questão"
    mvarlComando = vData
End Property

Public Property Get lComando() As Long
    lComando = mvarlComando
End Property

Private Sub Class_Initialize()

    Set mvarcolBrowseUsuarioCampo = New Collection
    Set mvarobjBrowseUsuarioOrdenacao = New AdmBrowseUsuarioOrdenacao
    Set mvarobjBrowseArquivo = New AdmBrowseArquivo
    Set mvarcolBrowseIndice = New Collection
    Set mvarcolBrowseIndiceUsuario = New Collection
    Set mvarobjChamaTela = New AdmChamaTela
    Set mvarcolBrowseCampo = New Collection
    
End Sub

Function Browse_BotaoConfigura_Click(objBrowse As AdmBrowse) As Long

Dim lErro As Long
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim objBrowseConfigura As New AdmBrowseConfigura
Dim objBrowseIndice As AdmBrowseIndice

On Error GoTo Erro_Browse_BotaoConfigura_Click

    For Each objBrowseUsuarioCampo In objBrowse.colBrowseUsuarioCampo
        If objBrowse.objGrid.ColWidth(objBrowseUsuarioCampo.iPosicaoTela - 1) <> objBrowseUsuarioCampo.lLargura Then
            objBrowseUsuarioCampo.lLargura = objBrowse.objGrid.ColWidth(objBrowseUsuarioCampo.iPosicaoTela - 1)
        End If
    Next
    
    Set objBrowseConfigura.objBrowseExcel = objBrowse.objBrowseExcel

    objBrowseConfigura.sNomeTela = objBrowse.objForm.Name
    Set objBrowseConfigura.colBrowseUsuarioCampo = objBrowse.colBrowseUsuarioCampo
    
    'passa a expressao de selecao criada pelo usuario como parametro para possivel alteracao
    objBrowseConfigura.sSelecaoSQL1Usuario = objBrowse.sSelecaoSQL1Usuario

    objBrowseConfigura.iPesquisa = objBrowse.iPesquisa
    
    objBrowseConfigura.iAlteradoOrdenacao = objBrowse.iAlteradoOrdenacao

    Set objBrowseConfigura.colBrowseIndiceUsuario = New Collection

    For Each objBrowseIndice In objBrowse.colBrowseIndiceUsuario
        objBrowseConfigura.colBrowseIndiceUsuario.Add objBrowseIndice
    Next

    Set objBrowseConfigura.colBrowseIndice = objBrowse.colBrowseIndice
        
    objBrowse.objForm.MousePointer = vbHourglass

    Call objBrowse.objChamaTela.Chama_Tela("BrowseConfigura", objBrowseConfigura)

    objBrowse.objForm.MousePointer = vbDefault
    
    If objBrowseConfigura.iTelaOK = 1 Then
    
        objBrowse.iAlterado = 1
        
        objBrowse.iAlteradoOrdenacao = objBrowseConfigura.iAlteradoOrdenacao
    
        'retorna a expressao de selecao criada pelo usuario, possivelmente alterada
        objBrowse.sSelecaoSQL1Usuario = objBrowseConfigura.sSelecaoSQL1Usuario
        objBrowse.sSelecaoSQL1 = objBrowseConfigura.sSelecaoSQL1
        Set objBrowse.objBrowseExcel = objBrowseConfigura.objBrowseExcel
    
        Set objBrowse.colBrowseIndiceUsuario = objBrowseConfigura.colBrowseIndiceUsuario
    
        lErro = Browse_Reinicializa_Ordenacao(objBrowse)
        If lErro <> SUCESSO Then Error 60828
    
        lErro = Browse_Inicializa_Campos(objBrowse)
        If lErro <> SUCESSO Then Error 9119
    
        lErro = Browse_Inicializa_Comando_SQL(objBrowse)
        If lErro <> SUCESSO Then Error 9120
        
        'inicializa o comando SQL que será usado para contar o número de registros da tabela
        lErro = Browse_Inicializa_Comando_SQL_Count(objBrowse)
        If lErro <> SUCESSO Then Error 20669
        
        objBrowse.iBind = 0
        
        lErro = Browse_Inicializa_Grid_Lote(objBrowse)
        If lErro <> SUCESSO Then Error 9122
        
        lErro = Browse_Inicializa_ScrollBar_Lote(objBrowse)
        If lErro <> SUCESSO Then Error 9121

        objBrowse.objGrid.Row = 1
        
        objBrowse.objBarraScroll.Max = objBrowse.objBarraScroll.Min
        
        Call Browse_BarraScroll_Change(objBrowse)
   
    End If
    
    Browse_BotaoConfigura_Click = SUCESSO
    
    Exit Function
    
Erro_Browse_BotaoConfigura_Click:

    Browse_BotaoConfigura_Click = Err

    Select Case Err

        Case 9119, 9120, 9121, 9122, 60828
        
        Case 3010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COLUNAS_EXCEDENTES_BROWSE", Err, objBrowseUsuarioCampo.iPosicaoTela, objBrowse.objGrid.Cols)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142055)

    End Select

    Exit Function
    
End Function

Function Browse_BotaoEdita_Click(objBrowse As AdmBrowse)

Dim lErro As Long
Dim sSQL As String
Dim lPosicao1 As Long

On Error GoTo Erro_Browse_BotaoEdita_Click

    lPosicao1 = objBrowse.lPosicao + objBrowse.objGrid.Row - objBrowse.objGrid.TopRow

    lErro = Browse_Monta_Comando_SQL(objBrowse, sSQL)
    If lErro <> SUCESSO Then Error 9162
    
    lErro = Browse_Inicializa_Campos1(objBrowse)
    If lErro <> SUCESSO Then Error 9174
    
    lErro = CF("Browse_Le_Posicao", lPosicao1, sSQL, objBrowse.colBrowseValorCampo, objBrowse.colSelecao, objBrowse)
    If lErro <> SUCESSO And lErro <> 20710 Then Error 9163

    If lErro = 20710 Then Error 55693

    Browse_BotaoEdita_Click = SUCESSO

    Exit Function

Erro_Browse_BotaoEdita_Click:

    Browse_BotaoEdita_Click = Err

    Select Case Err

        Case 9162, 9163, 9174, 55693
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142056)

    End Select

    Exit Function

End Function

Function Browse_Inicializa_Campos1(objBrowse As AdmBrowse) As Long
'inicializa todos os campos relacionados em colBrowseValorCampo

Dim lErro As Long
Dim objBrowseValorCampo As AdmBrowseValorCampo
Dim objCampo As New AdmCampos
Dim lComando3 As Long

On Error GoTo Erro_Browse_Inicializa_Campos1

    lComando3 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando3 = 0 Then Error 9181

    For Each objBrowseValorCampo In objBrowse.colBrowseValorCampo
            
        objCampo.sNomeArq = objBrowse.sNomeTabela
        objCampo.sNome = objBrowseValorCampo.sNomeCampo

        lErro = CF("Campos_Le2", objCampo, lComando3)
        If lErro <> SUCESSO And lErro <> 9184 Then Error 9170
        
        'se o campo não estiver cadastrado
        If lErro = 9184 Then Error 9171
        
        Select Case objCampo.iTipo
        
            Case ADM_TIPO_SMALLINT
                objBrowseValorCampo.vValorCampo = CInt(0)
                
            Case ADM_TIPO_INTEGER
                objBrowseValorCampo.vValorCampo = CLng(0)
                
            Case ADM_TIPO_DOUBLE
                objBrowseValorCampo.vValorCampo = CDbl(0)
                
            '###############################################
            'ALTERADO POR WAGNER
            Case ADM_TIPO_VARCHAR
                If objCampo.iTamanho < 500 Then
                    objBrowseValorCampo.vValorCampo = String(500, 0)
                Else
                    objBrowseValorCampo.vValorCampo = String(objCampo.iTamanho, 0)
                End If
            '###############################################
                            
            Case ADM_TIPO_DATE
                objBrowseValorCampo.vValorCampo = CDate("1/1/1997")
            
            Case Else
                Error 9172
                
        End Select
        
    Next

    Call Comando_Fechar(lComando3)
    
    Browse_Inicializa_Campos1 = SUCESSO
    
    Exit Function

Erro_Browse_Inicializa_Campos1:

    Browse_Inicializa_Campos1 = Err
    
    Select Case Err
    
        Case 9170
        
        Case 9171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPO_NAO_CADASTRADO", Err, objCampo.sNome, objCampo.sNomeArq)
        
        Case 9172
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CAMPO_INVALIDO", Err, objCampo.iTipo)
            
        Case 9181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142057)
        
    End Select

    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Browse_Monta_Comando_SQL(objBrowse As AdmBrowse, sSQL As String) As Long

Dim lErro As Long
Dim objBrowseValorCampo As AdmBrowseValorCampo
Dim objBrowseIndice As AdmBrowseIndice
Dim iUsuario As Integer
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim iCampoExercicio As Integer
Dim sSQL1 As String
Dim iIndice As Integer

    sSQL = "SELECT"

    For Each objBrowseValorCampo In objBrowse.colBrowseValorCampo

        sSQL = sSQL & " " & objBrowseValorCampo.sNomeCampo & ","

    Next

    sSQL = left(sSQL, Len(sSQL) - 1)

    sSQL = sSQL & " FROM " & objBrowse.sNomeTabela
    
    If Len(objBrowse.sSelecaoSQL) > 0 And Len(objBrowse.sSelecaoSQL1) > 0 Then
        sSQL = sSQL & " WHERE (" & objBrowse.sSelecaoSQL & ") AND (" & objBrowse.sSelecaoSQL1 & ")"
    ElseIf Len(objBrowse.sSelecaoSQL) > 0 Then
        sSQL = sSQL & " WHERE (" & objBrowse.sSelecaoSQL & ")"
    ElseIf Len(objBrowse.sSelecaoSQL1) > 0 Then
        sSQL = sSQL & " WHERE (" & objBrowse.sSelecaoSQL1 & ")"
    End If




    


    'colocado por Mario para resolver o problema de posicionamento do registro escolhido para edicao
    'sera feito uma clausula WHERE com os valores dos campos exibidos na tela

    'PARA USAR A ROTINA ANTIGA BASTA INSERIR UM REGISTRO NA TABELA CONTROLE COM Codigo=2 e Conteudo=1
    If giBrowsePosicaoAntigo = 0 Then


        For Each objBrowseUsuarioCampo In objBrowse.colBrowseUsuarioCampo
    
            
    
            iIndice = iIndice + 1
            
            'adicionado para poder selecionar o registro que esta sendo editado
            If Not (objBrowse.colReg Is Nothing) Then
        
                If objBrowse.colReg.Count >= (objBrowse.objGrid.Row - objBrowse.objGrid.TopRow + 1) Then
        
                
                    If objBrowse.colReg.Item(objBrowse.objGrid.Row - objBrowse.objGrid.TopRow + 1).Item(iIndice) = "" Or objBrowse.colReg.Item(objBrowse.objGrid.Row - objBrowse.objGrid.TopRow + 1).Item(iIndice) = 0 Or objBrowse.colReg.Item(objBrowse.objGrid.Row - objBrowse.objGrid.TopRow + 1).Item(iIndice) = DATA_NULA Then
                
                        sSQL1 = sSQL1 & " (" & objBrowseUsuarioCampo.sNome & "=? "
                        
                        sSQL1 = sSQL1 & "OR " & objBrowseUsuarioCampo.sNome & " IS NULL) AND"
                
                    Else
                    
                        sSQL1 = sSQL1 & " " & objBrowseUsuarioCampo.sNome & "=? AND"
                        
                    End If
                
                    If left(objBrowseUsuarioCampo.sNome, 7) = "Periodo" And left(objBrowseUsuarioCampo.sNome, 9) <> "PeriodoDe" And left(objBrowseUsuarioCampo.sNome, 10) <> "PeriodoAte" Then
                        sSQL1 = sSQL1 & " Exercicio=? AND"
                    End If
            
                End If
                
            End If
            
            
    '        'se o campo Periodo for incluido, então o campo Exercicio é incluido em seguida para
    '        'que possa posteriormente pegar o nome externo do campo.
    '        If Left(objBrowseUsuarioCampo.sNome, 7) = "Periodo" And Left(objBrowseUsuarioCampo.sNome, 9) <> "PeriodoDe" And Left(objBrowseUsuarioCampo.sNome, 10) <> "PeriodoAte" Then
    '            sSQL1 = sSQL1 & " Exercicio AS Exercicio" & CStr(iCampoExercicio) & ","
    '        End If
            
        Next
        
        If Len(sSQL1) <> 0 Then
        
            sSQL1 = left(sSQL1, Len(sSQL1) - 3)
             
            If Len(objBrowse.sSelecaoSQL) > 0 Or Len(objBrowse.sSelecaoSQL1) > 0 Then
            
                sSQL = sSQL & " AND (" & sSQL1 & ")"
                 
            Else
             
                sSQL = sSQL & " WHERE " & sSQL1
             
            End If
            
        End If
        
    End If








    iUsuario = 1

    For Each objBrowseIndice In objBrowse.colBrowseIndice
        If objBrowseIndice.iIndice = objBrowse.objOrdenacao.ItemData(objBrowse.objOrdenacao.ListIndex) Then
            iUsuario = -1
            Exit For
        End If
    Next
    
    If iUsuario = 1 Then
    
        For Each objBrowseIndice In objBrowse.colBrowseIndiceUsuario
            If objBrowseIndice.iIndice + objBrowse.colBrowseIndice.Count = objBrowse.objOrdenacao.ItemData(objBrowse.objOrdenacao.ListIndex) Then Exit For
        Next
    
    End If
    
    sSQL = sSQL & " ORDER BY " & objBrowseIndice.sOrdenacaoSQL
    
    Browse_Monta_Comando_SQL = SUCESSO

End Function

Function Browse_Inicializa(objBrowse As AdmBrowse) As Long

Dim lErro As Long
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim sCodUsuario As String
Dim iNumeroBotoes As Integer
Dim iPosicaoBotao As Integer
Dim objBrowseUsuario As New AdmBrowseUsuario

On Error GoTo Erro_Browse_Inicializa

    objBrowse.iSelecaoExterna = 0
   
    objBrowse.iBind = 0

    'obtem os dados de BrowseArquivo referentes a tela em questão
    lErro = CF("BrowseArquivo_Le", objBrowse.objForm.Name, objBrowse.objBrowseArquivo)
    If lErro <> SUCESSO Then gError 9250

    objBrowse.iBancoDados = objBrowse.objBrowseArquivo.iBancoDados

    If objBrowse.iBancoDados = BD_DICDADOS Then

        objBrowse.lComando = Comando_AbrirExt(GL_lConexaoDicBrowse)
        If objBrowse.lComando = 0 Then gError 20703

        objBrowse.lComando1 = Comando_AbrirExt(GL_lConexaoDicBrowse)
        If objBrowse.lComando1 = 0 Then gError 20704

        objBrowse.lComando2 = Comando_AbrirExt(GL_lConexaoDicBrowse)
        If objBrowse.lComando2 = 0 Then gError 20705

        objBrowse.lComando3 = Comando_AbrirExt(GL_lConexaoDicBrowse)
        If objBrowse.lComando3 = 0 Then gError 20706

    Else
        objBrowse.lComando = Comando_AbrirExt(GL_lConexaoBrowse)
        If objBrowse.lComando = 0 Then gError 9025
    
        objBrowse.lComando1 = Comando_AbrirExt(GL_lConexaoBrowse)
        If objBrowse.lComando1 = 0 Then gError 9026
    
        objBrowse.lComando2 = Comando_AbrirExt(GL_lConexaoBrowse)
        If objBrowse.lComando2 = 0 Then gError 9270
    
        objBrowse.lComando3 = Comando_AbrirExt(GL_lConexaoBrowse)
        If objBrowse.lComando3 = 0 Then gError 9271

    End If

    'le todos os campos que podem ser utilizado na tela de browse e coloca os resultados na coleção
    lErro = CF("BrowseCampo_Le", objBrowse.objForm.Name, objBrowse.colBrowseCampo)
    If lErro <> SUCESSO Then gError 89974

    sCodUsuario = gsUsuario

    objBrowse.sCodUsuario = sCodUsuario

    objBrowseUsuario.sNomeTela = objBrowse.objForm.Name
    objBrowseUsuario.sCodUsuario = objBrowse.sCodUsuario

    If Len(Trim(objBrowse.sOpcao)) = 0 Then

        'obtem os dados de BrowseUsuario referentes a tela em questão
        lErro = CF("BrowseUsuario_Le", objBrowseUsuario)
        If lErro <> SUCESSO And lErro <> 91302 Then gError 89989

    Else
    
        'obtem os dados de BrowseOpcao referentes a tela em questão
        lErro = CF("BrowseOpcao_Le", objBrowse.sOpcao, objBrowseUsuario)
        If lErro <> SUCESSO And lErro <> 178347 Then gError 178354
    
    End If

    If lErro = SUCESSO Then
        objBrowse.objPrincMDIChild.Move objBrowseUsuario.lEsquerda, objBrowseUsuario.lTopo, objBrowseUsuario.lLargura, objBrowseUsuario.lAltura
    Else
        objBrowseUsuario.lLargura = objBrowse.objForm.Width
    End If

    lErro = Browse_Inicializa_colBrowseUsuarioCampo(objBrowse)
    If lErro <> SUCESSO Then gError 9085

    Set objBrowseUsuarioCampo = objBrowse.colBrowseUsuarioCampo.Item(1)

    objBrowse.sNomeTabela = objBrowseUsuarioCampo.sNomeArq

    lErro = Browse_Inicializa_Ordenacao(objBrowse)
    If lErro <> SUCESSO Then gError 9055

    lErro = Browse_Inicializa_Campos(objBrowse)
    If lErro <> SUCESSO Then gError 9113
    
    lErro = CF("BrowseExcel_Le", objBrowse, objBrowse.sOpcao, objBrowse.objBrowseExcel)
    If lErro <> SUCESSO Then gError 9113
    
    If colBrowseCampo.Count > 0 Then
    
        iNumeroBotoes = BROWSE_NUMERO_BOTOES + objBrowse.objBrowseArquivo.iBotaoSeleciona + objBrowse.objBrowseArquivo.iBotaoEdita + objBrowse.objBrowseArquivo.iBotaoConsulta
        
        iPosicaoBotao = ESPACO_ENTRE_BOTOES
        
        objBrowse.objForm.BotaoPlanilha.left = iPosicaoBotao
        
        iPosicaoBotao = iPosicaoBotao + objBrowse.objForm.BotaoPlanilha.Width + ESPACO_ENTRE_BOTOES
        
        If objBrowse.objBrowseArquivo.iBotaoSeleciona = 1 Then
        
            objBrowse.objForm.BotaoSeleciona.left = iPosicaoBotao
            objBrowse.objForm.BotaoSeleciona.Visible = True
            
            iPosicaoBotao = iPosicaoBotao + objBrowse.objForm.BotaoSeleciona.Width + ESPACO_ENTRE_BOTOES
            
        End If
        
        If objBrowse.objBrowseArquivo.iBotaoEdita = 1 Then
        
            objBrowse.objForm.BotaoEdita.left = iPosicaoBotao
            objBrowse.objForm.BotaoEdita.Visible = True
            iPosicaoBotao = iPosicaoBotao + objBrowse.objForm.BotaoEdita.Width + ESPACO_ENTRE_BOTOES
            
        End If
        
        If objBrowse.objBrowseArquivo.iBotaoConsulta = 1 Then
        
            objBrowse.objForm.BotaoConsulta.left = iPosicaoBotao
            objBrowse.objForm.BotaoConsulta.Visible = True
            
            iPosicaoBotao = iPosicaoBotao + objBrowse.objForm.BotaoConsulta.Width + ESPACO_ENTRE_BOTOES
            
        End If
        
        objBrowse.objForm.BotaoConfigura.left = iPosicaoBotao
        
        iPosicaoBotao = iPosicaoBotao + objBrowse.objForm.BotaoConfigura.Width + ESPACO_ENTRE_BOTOES
        
        objBrowse.objForm.BotaoPesquisa.left = iPosicaoBotao
        
        iPosicaoBotao = iPosicaoBotao + objBrowse.objForm.BotaoPesquisa.Width + ESPACO_ENTRE_BOTOES
        
        objBrowse.objForm.BotaoAtualiza.left = iPosicaoBotao
        
        iPosicaoBotao = iPosicaoBotao + objBrowse.objForm.BotaoAtualiza.Width + ESPACO_ENTRE_BOTOES
        
        objBrowse.objForm.BotaoFecha.left = iPosicaoBotao
        
        iPosicaoBotao = iPosicaoBotao + objBrowse.objForm.BotaoFecha.Width + ESPACO_ENTRE_BOTOES
        
        objBrowse.objForm.Picture2.Width = iPosicaoBotao
        
'        objBrowse.objForm.Picture2.Left = (objBrowseUsuario.lLargura - (1440 * 2 * GetSystemMetrics(SM_CXFIXEDFRAME)) - objBrowse.objForm.Picture2.Width) / 2
         
    End If
    
    Browse_Inicializa = SUCESSO

    Exit Function

Erro_Browse_Inicializa:

    Browse_Inicializa = gErr

    Select Case gErr

        Case 9055, 9085, 9113, 9250, 89974, 89989, 178354

        Case 9025, 9026, 9270, 9271, 20703, 20704, 20705, 20706
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 9036
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OBTENCAO_CODIGO_USUARIO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142058)

    End Select

    Call Comando_Fechar(objBrowse.lComando)
    objBrowse.lComando = 0
    Call Comando_Fechar(objBrowse.lComando1)
    objBrowse.lComando1 = 0
    Call Comando_Fechar(objBrowse.lComando2)
    objBrowse.lComando2 = 0
    Call Comando_Fechar(objBrowse.lComando3)
    objBrowse.lComando3 = 0

    Exit Function

End Function


Function Browse_Inicializa_colBrowseUsuarioCampo(objBrowse As AdmBrowse)
'inicializa a coleção de campos disponiveis para a tela x usuario

Dim lErro As Long
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim sCodGrupo As String
Dim colGrupoBrowseCampo As New Collection
Dim objGrupoBrowseCampo As AdmGrupoBrowseCampo
Dim objCampo As AdmCampos
Dim iIndice As Integer
Dim lComando2 As Long
Dim colBrowseUsuarioCampoSupervisor As New Collection
Dim objBrowseUsuarioCampoSup As AdmBrowseUsuarioCampo

On Error GoTo Erro_Browse_Inicializa_colBrowseUsuarioCampo

    lComando2 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando2 = 0 Then Error 9176

    'obtem os campos que serão exibidos para o usuario
    lErro = CF("BrowseUsuarioCampo_Le", objBrowse.objForm.Name, objBrowse.sCodUsuario, objBrowse.colBrowseUsuarioCampo)
    If lErro <> SUCESSO Then Error 9037

    'se ainda não houver nenhum campo selecionado para o usuario
    If objBrowse.colBrowseUsuarioCampo.Count < 1 Then

        sCodGrupo = String(STRING_GRUPO, 0)

        'obtem o codigo do grupo
        lErro = Obter_Grupo(sCodGrupo)
        If lErro <> SUCESSO Then Error 9075

        'le os campos disponiveis para o grupo
        lErro = CF("GrupoBrowseCampo_Le", sCodGrupo, objBrowse.objForm.Name, colGrupoBrowseCampo)
        If lErro <> SUCESSO Then Error 9076

        iIndice = 0

        If colGrupoBrowseCampo.Count > UBound(avCampo) Then Error 9084

        'obtem os campos que são exibidos para o supervisor
        lErro = CF("BrowseUsuarioCampo_Le", objBrowse.objForm.Name, CODUSUARIO_SUPERVISOR, colBrowseUsuarioCampoSupervisor)
        If lErro <> SUCESSO Then Error 55668

        For Each objBrowseUsuarioCampoSup In colBrowseUsuarioCampoSupervisor

            'pesquisa os campos do grupo disponiveis para o usuario
            For Each objGrupoBrowseCampo In colGrupoBrowseCampo
            
                If objBrowseUsuarioCampoSup.sNome = objGrupoBrowseCampo.sNome And objBrowseUsuarioCampoSup.sNomeArq = objGrupoBrowseCampo.sNomeArq Then

                    iIndice = iIndice + 1

                    Set objBrowseUsuarioCampo = New AdmBrowseUsuarioCampo

                    objBrowseUsuarioCampo.sNomeTela = objBrowseUsuarioCampoSup.sNomeTela
                    objBrowseUsuarioCampo.sCodUsuario = objBrowse.sCodUsuario
                    objBrowseUsuarioCampo.sNomeArq = objBrowseUsuarioCampoSup.sNomeArq
                    objBrowseUsuarioCampo.sNome = objBrowseUsuarioCampoSup.sNome
                    objBrowseUsuarioCampo.iPosicaoTela = iIndice
                    objBrowseUsuarioCampo.sTitulo = objBrowseUsuarioCampoSup.sTitulo
                    objBrowseUsuarioCampo.lLargura = objBrowseUsuarioCampoSup.lLargura

                    objBrowse.colBrowseUsuarioCampo.Add objBrowseUsuarioCampo

                    Exit For
                    
                End If

            Next
            
        Next
        
        If objBrowse.colBrowseUsuarioCampo.Count = 0 Then
        
            Set objCampo = New AdmCampos
            
            Set objGrupoBrowseCampo = colGrupoBrowseCampo.Item(1)

            objCampo.sNomeArq = objGrupoBrowseCampo.sNomeArq
            objCampo.sNome = objGrupoBrowseCampo.sNome

            lErro = CF("Campos_Le1", objCampo, lComando2)
            If lErro <> SUCESSO And lErro <> 9179 Then Error 9082

            'se o campo não estiver cadastrado
            If lErro = 9179 Then Error 9083

            Set objBrowseUsuarioCampo = New AdmBrowseUsuarioCampo

            objBrowseUsuarioCampo.sNomeTela = objBrowse.objForm.Name
            objBrowseUsuarioCampo.sCodUsuario = objBrowse.sCodUsuario
            objBrowseUsuarioCampo.sNomeArq = objCampo.sNomeArq
            objBrowseUsuarioCampo.sNome = objCampo.sNome
            objBrowseUsuarioCampo.iPosicaoTela = 1
            objBrowseUsuarioCampo.sTitulo = objCampo.sTituloGrid
            objBrowseUsuarioCampo.lLargura = 1000

            objBrowse.colBrowseUsuarioCampo.Add objBrowseUsuarioCampo
        
        End If
        
    End If

    If objBrowse.colBrowseUsuarioCampo.Count < 1 Then Error 9039

    Call Comando_Fechar(lComando2)

    Browse_Inicializa_colBrowseUsuarioCampo = SUCESSO

    Exit Function

Erro_Browse_Inicializa_colBrowseUsuarioCampo:

    Browse_Inicializa_colBrowseUsuarioCampo = Err

    Select Case Err

        Case 9037, 9076, 9082, 55668

        Case 9039
            Call Rotina_Erro(vbOKOnly, "ERRO_BROWSE_SEM_COLUNAS", Err)
            
        Case 9075
            Call Rotina_Erro(vbOKOnly, "ERRO_OBTENCAO_CODIGO_GRUPO", Err)

        Case 9083
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPO_NAO_CADASTRADO", Err, objCampo.sNome, objCampo.sNomeArq)

        Case 9084
            Call Rotina_Erro(vbOKOnly, "ERRO_BROWSE_EXCEDEU_MAXIMO_COLUNAS", Err)

        Case 9176
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142059)

    End Select

    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Browse_Inicializa_Campos(objBrowse As AdmBrowse) As Long

Dim lErro As Long
Dim iCampo As Integer
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim objCampo As New AdmCampos
Dim objValorCampo As AdmValorCampo
Dim lComando3 As Long

On Error GoTo Erro_Browse_Inicializa_Campos

    Set objBrowse.colValorCampo = New Collection

    iCampo = 0

    lComando3 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando3 = 0 Then Error 9193

    For Each objBrowseUsuarioCampo In objBrowse.colBrowseUsuarioCampo

        objCampo.sNomeArq = objBrowseUsuarioCampo.sNomeArq
        objCampo.sNome = objBrowseUsuarioCampo.sNome

        lErro = CF("Campos_Le2", objCampo, lComando3)
        If lErro <> SUCESSO And lErro <> 9184 Then Error 9063

        'se o campo não estiver cadastrado
        If lErro = 9184 Then Error 9064

        Set objValorCampo = New AdmValorCampo

        objValorCampo.iTipo = objCampo.iTipo
        objValorCampo.iPosicaoTela = objBrowseUsuarioCampo.iPosicaoTela
        objValorCampo.iSubTipo = objCampo.iSubTipo
        objBrowseUsuarioCampo.iAlinhamento = objCampo.iAlinhamento
        objValorCampo.iTamanho = objCampo.iTamanho

        iCampo = iCampo + 1

        Select Case objValorCampo.iTipo

            Case ADM_TIPO_SMALLINT
                avCampo(iCampo) = CInt(0)

            Case ADM_TIPO_INTEGER
                avCampo(iCampo) = CLng(0)

            Case ADM_TIPO_DOUBLE
                avCampo(iCampo) = CDbl(0)

            '###############################################
            'ALTERADO POR WAGNER
            Case ADM_TIPO_VARCHAR
                If objCampo.iTamanho < 500 Then
                    avCampo(iCampo) = String(500, 0)
                Else
                    avCampo(iCampo) = String(objCampo.iTamanho, 0)
                End If
            '###############################################

            Case ADM_TIPO_DATE
                avCampo(iCampo) = CDate("1/1/1997")

            Case Else
                Error 9071

        End Select

        objBrowse.colValorCampo.Add objValorCampo
        
        'se o campo que acabou de ser processado foi Periodo
        'acrescenta a inicializaçao do campo Exercicio que
        'será sempre incluido após o Periodo. Isto é necessário
        'já que é preciso pegar o nome externo do periodo
        If left(objBrowseUsuarioCampo.sNome, 7) = "Periodo" And left(objBrowseUsuarioCampo.sNome, 9) <> "PeriodoDe" And left(objBrowseUsuarioCampo.sNome, 10) <> "PeriodoAte" Then
            Set objValorCampo = New AdmValorCampo
            iCampo = iCampo + 1
            avCampo(iCampo) = CInt(0)
            objValorCampo.iTipo = ADM_TIPO_INTEGER
            objBrowse.colValorCampo.Add objValorCampo
        End If
        
        
        
    Next

    'VERMELHO
    If objBrowse.objForm.Name = "ClientesLista" And giCliAtrasoDestacar <> 0 Then
        Set objValorCampo = New AdmValorCampo
        iCampo = iCampo + 1
        avCampo(iCampo) = CLng(0)
        objValorCampo.iTipo = ADM_TIPO_INTEGER
        objBrowse.colValorCampo.Add objValorCampo
    End If


    Call Comando_Fechar(lComando3)
    
    objBrowse.iBind = 1

    Browse_Inicializa_Campos = SUCESSO

    Exit Function

Erro_Browse_Inicializa_Campos:

    Browse_Inicializa_Campos = Err

    Select Case Err

        Case 9063

        Case 9064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPO_NAO_CADASTRADO", Err, objCampo.sNome, objCampo.sNomeArq)

        Case 9071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CAMPO_INVALIDO", Err, objValorCampo.iTipo)

        Case 9193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142060)

    End Select

    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Private Function Browse_Reinicialza_Avcampo(ByVal objBrowse As AdmBrowse) As Long
'reinicializa os campos que vão receber os dados do bd já que vai ocorrer um novo bind

Dim lErro As Long
Dim iCampo As Integer
Dim objValorCampo As AdmValorCampo

On Error GoTo Erro_Browse_Reinicialza_Avcampo

    iCampo = 0

    For Each objValorCampo In objBrowse.colValorCampo

        iCampo = iCampo + 1

        Select Case objValorCampo.iTipo

            Case ADM_TIPO_SMALLINT
                avCampo(iCampo) = CInt(0)

            Case ADM_TIPO_INTEGER
                avCampo(iCampo) = CLng(0)

            Case ADM_TIPO_DOUBLE
                avCampo(iCampo) = CDbl(0)

            '###############################################
            'ALTERADO POR WAGNER
            Case ADM_TIPO_VARCHAR
                If objValorCampo.iTamanho < 500 Then
                    avCampo(iCampo) = String(500, 0)
                Else
                    avCampo(iCampo) = String(objValorCampo.iTamanho, 0)
                End If
            '###############################################

            Case ADM_TIPO_DATE
                avCampo(iCampo) = CDate("1/1/1997")

            Case Else
                Error 20804

        End Select

    Next

    Browse_Reinicialza_Avcampo = SUCESSO

    Exit Function

Erro_Browse_Reinicialza_Avcampo:

    Browse_Reinicialza_Avcampo = Err

    Select Case Err

        Case 20804
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CAMPO_INVALIDO", Err, objValorCampo.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142061)

    End Select

    Exit Function

End Function

Public Function Browse_Reinicializa_avCampo2(ByVal objBrowse As AdmBrowse, avCampoAux() As Variant) As Long
'reinicializa os campos que vão receber os dados do bd já que vai ocorrer um novo bind

Dim lErro As Long
Dim iCampo As Integer
Dim objValorCampo As AdmValorCampo

On Error GoTo Erro_Browse_Reinicializa_avCampo2

    iCampo = 0

    For Each objValorCampo In objBrowse.colValorCampo

        iCampo = iCampo + 1

        Select Case objValorCampo.iTipo

            Case ADM_TIPO_SMALLINT
                avCampoAux(iCampo) = CInt(0)

            Case ADM_TIPO_INTEGER
                avCampoAux(iCampo) = CLng(0)

            Case ADM_TIPO_DOUBLE
                avCampoAux(iCampo) = CDbl(0)

            '###############################################
            'ALTERADO POR WAGNER
            Case ADM_TIPO_VARCHAR
                If objValorCampo.iTamanho < 500 Then
                    avCampoAux(iCampo) = String(500, 0)
                Else
                    avCampoAux(iCampo) = String(objValorCampo.iTamanho, 0)
                End If
            '###############################################

            Case ADM_TIPO_DATE
                avCampoAux(iCampo) = CDate("1/1/1997")

            Case Else
                gError 102094

        End Select

    Next

    Browse_Reinicializa_avCampo2 = SUCESSO

    Exit Function

Erro_Browse_Reinicializa_avCampo2:

    Browse_Reinicializa_avCampo2 = gErr

    Select Case gErr

        Case 102094
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CAMPO_INVALIDO", gErr, objValorCampo.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142062)

    End Select

    Exit Function

End Function

Function Browse_Inicializa_Comando_SQL(objBrowse As AdmBrowse) As Long

Dim lErro As Long
Dim objBrowseIndice As AdmBrowseIndice
Dim iCampoExercicio As Integer
Dim iUsuario As Integer
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo

On Error GoTo Erro_Browse_Inicializa_Comando_SQL

    iCampoExercicio = 0

    objBrowse.sSQL = "SELECT"

    For Each objBrowseUsuarioCampo In objBrowse.colBrowseUsuarioCampo

        objBrowse.sSQL = objBrowse.sSQL & " " & objBrowseUsuarioCampo.sNome & ","
        
        'se o campo Periodo for incluido, então o campo Exercicio é incluido em seguida para
        'que possa posteriormente pegar o nome externo do campo.
        If left(objBrowseUsuarioCampo.sNome, 7) = "Periodo" And left(objBrowseUsuarioCampo.sNome, 9) <> "PeriodoDe" And left(objBrowseUsuarioCampo.sNome, 10) <> "PeriodoAte" Then
            iCampoExercicio = iCampoExercicio + 1
            objBrowse.sSQL = objBrowse.sSQL & " Exercicio AS Exercicio" & CStr(iCampoExercicio) & ","
        End If
        
    Next

    'VERMELHO
    If objBrowse.objForm.Name = "ClientesLista" And giCliAtrasoDestacar <> 0 Then
            objBrowse.sSQL = objBrowse.sSQL & " ClienteCod AS ClienteCod1,"
            iCampoExercicio = iCampoExercicio + 1
    End If

    objBrowse.sSQL = left(objBrowse.sSQL, Len(objBrowse.sSQL) - 1)

    Set objBrowseUsuarioCampo = objBrowse.colBrowseUsuarioCampo.Item(1)

    objBrowse.sSQL = objBrowse.sSQL & " FROM " & objBrowseUsuarioCampo.sNomeArq

    If Len(objBrowse.sSelecaoSQL) > 0 And Len(objBrowse.sSelecaoSQL1) > 0 Then
        objBrowse.sSQL = objBrowse.sSQL & " WHERE (" & objBrowse.sSelecaoSQL & ") AND (" & objBrowse.sSelecaoSQL1 & ")"
    ElseIf Len(objBrowse.sSelecaoSQL) > 0 Then
        objBrowse.sSQL = objBrowse.sSQL & " WHERE (" & objBrowse.sSelecaoSQL & ")"
    ElseIf Len(objBrowse.sSelecaoSQL1) > 0 Then
        objBrowse.sSQL = objBrowse.sSQL & " WHERE (" & objBrowse.sSelecaoSQL1 & ")"
    End If

    iUsuario = 1

    For Each objBrowseIndice In objBrowse.colBrowseIndice
        If objBrowseIndice.sNomeIndice = objBrowse.objOrdenacao.Text Then
            objBrowse.sSQL = objBrowse.sSQL & " ORDER BY " & objBrowseIndice.sOrdenacaoSQL
            iUsuario = -1
            Exit For
        End If
    Next

    If iUsuario = 1 Then

        For Each objBrowseIndice In objBrowse.colBrowseIndiceUsuario
            If objBrowseIndice.sNomeIndice = objBrowse.objOrdenacao.Text Then
                objBrowse.sSQL = objBrowse.sSQL & " ORDER BY " & objBrowseIndice.sOrdenacaoSQL
                Exit For
            End If
        Next

    End If
    
    lErro = CF("Browse_Prepara_Comando_SQL", objBrowse.sSQL, objBrowse.colBrowseUsuarioCampo.Count + iCampoExercicio, objBrowse.lComando, avCampo, objBrowse.colSelecao, objBrowse.iBind)
    If lErro <> SUCESSO Then Error 9056

    Browse_Inicializa_Comando_SQL = SUCESSO

    Exit Function

Erro_Browse_Inicializa_Comando_SQL:

    Browse_Inicializa_Comando_SQL = Err

    Select Case Err

        Case 9056

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142063)

    End Select

    Exit Function

End Function

Function Browse_Inicializa_Comando_SQL_Count(objBrowse As AdmBrowse) As Long

Dim lErro As Long
Dim colCampo As New Collection

On Error GoTo Erro_Browse_Inicializa_Comando_SQL_Count


    objBrowse.sSQLCount = "SELECT Count(*) FROM " & objBrowse.sNomeTabela

    If Len(objBrowse.sSelecaoSQL) > 0 And Len(objBrowse.sSelecaoSQL1) > 0 Then
        objBrowse.sSQLCount = objBrowse.sSQLCount & " WHERE (" & objBrowse.sSelecaoSQL & ") AND (" & objBrowse.sSelecaoSQL1 & ")"
    ElseIf Len(objBrowse.sSelecaoSQL) > 0 Then
        objBrowse.sSQLCount = objBrowse.sSQLCount & " WHERE (" & objBrowse.sSelecaoSQL & ")"
    ElseIf Len(objBrowse.sSelecaoSQL1) > 0 Then
        objBrowse.sSQLCount = objBrowse.sSQLCount & " WHERE (" & objBrowse.sSelecaoSQL1 & ")"
    End If

    avCampo1(1) = lNumReg

    lErro = CF("Browse_Prepara_Comando_SQL", objBrowse.sSQLCount, 1, objBrowse.lComando1, avCampo1, objBrowse.colSelecao, 1)
    If lErro <> SUCESSO Then Error 20728

    Browse_Inicializa_Comando_SQL_Count = SUCESSO

    Exit Function

Erro_Browse_Inicializa_Comando_SQL_Count:

    Browse_Inicializa_Comando_SQL_Count = Err

    Select Case Err

        Case 20728

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142064)

    End Select

    Exit Function

End Function

Function Browse_Inicializa_Ordenacao(objBrowse As AdmBrowse) As Long
'inicializa a ordenacao para o usuario em questão

Dim lErro As Long
Dim objBrowseIndice As AdmBrowseIndice
Dim iIndice As Integer
Dim objBrowseUsuarioOrdenacao As New AdmBrowseUsuarioOrdenacao
Dim iAchou As Integer

On Error GoTo Erro_Browse_Inicializa_Ordenacao

    objBrowse.objOrdenacao.Clear

    lErro = CF("BrowseIndice_Le", objBrowse.objForm.Name, objBrowse.colBrowseIndice)
    If lErro <> SUCESSO Then Error 9053

    If objBrowse.colBrowseIndice.Count < 1 Then Error 9054

    For Each objBrowseIndice In objBrowse.colBrowseIndice
        objBrowse.objOrdenacao.AddItem objBrowseIndice.sNomeIndice
        objBrowse.objOrdenacao.ItemData(objBrowse.objOrdenacao.NewIndex) = objBrowseIndice.iIndice
    Next

    lErro = CF("BrowseIndiceUsuario_Le", objBrowse.objForm.Name, objBrowse.sCodUsuario, objBrowse.colBrowseIndiceUsuario)
    If lErro <> SUCESSO Then Error 60820

    For Each objBrowseIndice In objBrowse.colBrowseIndiceUsuario
        objBrowse.objOrdenacao.AddItem objBrowseIndice.sNomeIndice
        objBrowse.objOrdenacao.ItemData(objBrowse.objOrdenacao.NewIndex) = objBrowseIndice.iIndice + objBrowse.colBrowseIndice.Count
    Next

    objBrowseUsuarioOrdenacao.sNomeTela = objBrowse.objForm.Name
    objBrowseUsuarioOrdenacao.sCodUsuario = objBrowse.sCodUsuario

    lErro = CF("BrowseUsuarioOrdenacao_Le", objBrowseUsuarioOrdenacao)
    If lErro <> SUCESSO And lErro <> 9044 Then Error 9046

    'se não há uma ordenacao selecionada para este usuario nesta tela ==> seleciona a primeira ordenação
    If lErro = 9044 Then

        objBrowse.objOrdenacao.ListIndex = 0

    Else

        iAchou = 0

        For iIndice = 0 To objBrowse.objOrdenacao.ListCount - 1
            If objBrowse.objOrdenacao.List(iIndice) = objBrowseUsuarioOrdenacao.sNomeIndice Then
                objBrowse.objOrdenacao.ListIndex = iIndice
                iAchou = 1
                Exit For
            End If
        Next

        If iAchou = 0 Then

            For iIndice = 0 To objBrowse.objOrdenacao.ListCount - 1
                If objBrowse.objOrdenacao.ItemData(iIndice) = objBrowseUsuarioOrdenacao.iIndice Then
                    objBrowse.objOrdenacao.ListIndex = iIndice
                    iAchou = 1
                    Exit For
                End If
            Next

        End If

        If iAchou = 0 Then objBrowse.objOrdenacao.ListIndex = 0

    End If

    'transfere a pesquisa que o usuario fez pela ultima vez
    objBrowse.sSelecaoSQL1Usuario = objBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario
    objBrowse.sSelecaoSQL1 = objBrowseUsuarioOrdenacao.sSelecaoSQL1

    Browse_Inicializa_Ordenacao = SUCESSO

    Exit Function

Erro_Browse_Inicializa_Ordenacao:

    Browse_Inicializa_Ordenacao = Err

    Select Case Err

        Case 9046, 9053

        Case 9054
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BROWSE_SEM_ORDENACAO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142065)

    End Select

    Exit Function

End Function

Function Browse_Inicializa_Grid_Lote(objBrowse As AdmBrowse) As Long

Dim iIndice As Integer
Dim iLinha As Integer
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim lErro As Long

On Error GoTo Erro_Browse_Inicializa_Grid_Lote


    objBrowse.objGrid.Cols = objBrowse.colBrowseUsuarioCampo.Count

    For Each objBrowseUsuarioCampo In objBrowse.colBrowseUsuarioCampo
        objBrowse.objGrid.ColWidth(objBrowseUsuarioCampo.iPosicaoTela - 1) = objBrowseUsuarioCampo.lLargura
        objBrowse.objGrid.TextMatrix(0, objBrowseUsuarioCampo.iPosicaoTela - 1) = objBrowseUsuarioCampo.sTitulo
        If objBrowseUsuarioCampo.iAlinhamento = BROWSER_CAMPO_ALINHAMENTO_DIREITA Then
            objBrowse.objGrid.ColAlignment(objBrowseUsuarioCampo.iPosicaoTela - 1) = flexAlignRightTop
        Else
            objBrowse.objGrid.ColAlignment(objBrowseUsuarioCampo.iPosicaoTela - 1) = flexAlignLeftTop
        End If
        
    Next

    objBrowse.objGrid.TopRow = objBrowse.iGrid_LinhaInicial
    

    Browse_Inicializa_Grid_Lote = SUCESSO

    Exit Function

Erro_Browse_Inicializa_Grid_Lote:

    Browse_Inicializa_Grid_Lote = Err

    Select Case Err

        Case 30010
            MsgBox ("Verifique na tabela BrowseUsuarioCampo para o arquivo " & objBrowseUsuarioCampo.sNomeArq & " , campo " & objBrowseUsuarioCampo.sNome & " e campos anteriores por ordem de posicao na tela. Deve haver alguma posicao faltando.")
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142066)

    End Select

    Exit Function

End Function

Function Browse_Inicializa_ScrollBar_Lote(objBrowse As AdmBrowse) As Long
'inicializa a scrollbar que acompanha o grid

Dim lErro As Long

On Error GoTo Erro_Browse_Inicializa_ScrollBar_Lote

    objBrowse.objBarraScroll.left = objBrowse.objGrid.left + objBrowse.objGrid.Width - 300
    objBrowse.objBarraScroll.top = objBrowse.objGrid.top + 50
    objBrowse.objBarraScroll.Height = objBrowse.objGrid.Height - 100

    objBrowse.objBarraScroll.Min = -32000
    objBrowse.lOldScrollValue = objBrowse.objBarraScroll.Min

    Browse_Inicializa_ScrollBar_Lote = SUCESSO

    Exit Function

Erro_Browse_Inicializa_ScrollBar_Lote:

    Browse_Inicializa_ScrollBar_Lote = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142067)

    End Select

    Exit Function

End Function


Function Browse_Form_Unload(objBrowse As AdmBrowse) As Long

Dim lErro As Long
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
Dim objBrowseUsuarioOrdenacao As New AdmBrowseUsuarioOrdenacao
Dim objBrowseUsuario As New AdmBrowseUsuario

On Error GoTo Erro_Browse_Form_Unload
 
    Call Comando_Fechar(objBrowse.lComando)
    objBrowse.lComando = 0
    Call Comando_Fechar(objBrowse.lComando1)
    objBrowse.lComando1 = 0
    Call Comando_Fechar(objBrowse.lComando2)
    objBrowse.lComando2 = 0
    Call Comando_Fechar(objBrowse.lComando3)
    objBrowse.lComando3 = 0
    
    If Len(Trim(objBrowse.objComboOpcoes.Text)) = 0 Then
    
       'se não está fazendo a carga inicial da tela
       If objBrowse.iAlterado >= 0 Then
    
           For Each objBrowseUsuarioCampo In objBrowse.colBrowseUsuarioCampo
       
               If objBrowse.objGrid.ColWidth(objBrowseUsuarioCampo.iPosicaoTela - 1) <> objBrowseUsuarioCampo.lLargura Then
                   objBrowseUsuarioCampo.lLargura = objBrowse.objGrid.ColWidth(objBrowseUsuarioCampo.iPosicaoTela - 1)
                   objBrowse.iAlterado = 1
               End If
           Next
    
           objBrowseUsuarioOrdenacao.sNomeTela = objBrowse.objForm.Name
           objBrowseUsuarioOrdenacao.sCodUsuario = objBrowse.sCodUsuario
           objBrowseUsuarioOrdenacao.iIndice = objBrowse.objOrdenacao.ItemData(objBrowse.objOrdenacao.ListIndex)
           objBrowseUsuarioOrdenacao.sNomeIndice = objBrowse.objOrdenacao.Text
           objBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario = objBrowse.sSelecaoSQL1Usuario
           objBrowseUsuarioOrdenacao.sSelecaoSQL1 = objBrowse.sSelecaoSQL1
           
           lErro = CF("BrowseConfiguracoes_Grava", objBrowse, objBrowseUsuarioOrdenacao)
           If lErro <> SUCESSO Then gError 60832
           
       End If
       
       If Not objBrowse.objPrincMDIChild Is Nothing Then
       
           'se a janela está com tamanho normal, isto é, nao se está maximizada ou minimizada
           If objBrowse.objPrincMDIChild.WindowState = 0 Then
           
               objBrowseUsuario.sNomeTela = objBrowse.objForm.Name
               objBrowseUsuario.sCodUsuario = objBrowse.sCodUsuario
               
               objBrowseUsuario.lEsquerda = objBrowse.objPrincMDIChild.left
               objBrowseUsuario.lTopo = objBrowse.objPrincMDIChild.top
               objBrowseUsuario.lLargura = objBrowse.objPrincMDIChild.Width
               objBrowseUsuario.lAltura = objBrowse.objPrincMDIChild.Height
               
               lErro = CF("BrowseUsuario_Atualiza", objBrowseUsuario)
               If lErro <> SUCESSO Then gError 89990
               
           End If
       
       End If
    
    End If
    
    Browse_Form_Unload = SUCESSO

    Exit Function

Erro_Browse_Form_Unload:

    Browse_Form_Unload = gErr

    Select Case gErr

        Case 60832, 89990

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142068)

    End Select

    Call Comando_Fechar(objBrowse.lComando)
    objBrowse.lComando = 0
    Call Comando_Fechar(objBrowse.lComando1)
    objBrowse.lComando1 = 0
    Call Comando_Fechar(objBrowse.lComando2)
    objBrowse.lComando2 = 0
    Call Comando_Fechar(objBrowse.lComando3)
    objBrowse.lComando3 = 0

    Exit Function

End Function

Function Browse_GridBrowse_RowColChange(objBrowse As AdmBrowse) As Long

Dim lErro As Long
Dim lLargura As Long
Dim iIndice As Integer


    If objBrowse.objGrid.Row < objBrowse.iGrid_LinhaInicial Or objBrowse.objGrid.Row > objBrowse.iGrid_LinhaFinal Then

        objBrowse.iTeclado = 1

        If objBrowse.objGrid.Row < (objBrowse.iGrid_LinhaInicial - objBrowse.objBarraScroll.LargeChange) Then

            objBrowse.objBarraScroll.Value = objBrowse.objBarraScroll.Min
            objBrowse.objGrid.Row = objBrowse.iGrid_LinhaInicial

        ElseIf objBrowse.objGrid.Row > (objBrowse.iGrid_LinhaFinal + objBrowse.objBarraScroll.LargeChange) Then

            objBrowse.objBarraScroll.Value = objBrowse.objBarraScroll.Max
            objBrowse.objGrid.Row = objBrowse.iGrid_LinhaFinal

        ElseIf objBrowse.objGrid.Row < objBrowse.iGrid_LinhaInicial Then

            If objBrowse.objBarraScroll.Value + (objBrowse.objGrid.Row - objBrowse.iGrid_LinhaInicial) < objBrowse.objBarraScroll.Min Then
                objBrowse.objBarraScroll.Value = objBrowse.objBarraScroll.Min
            Else
                objBrowse.objBarraScroll.Value = objBrowse.objBarraScroll.Value + (objBrowse.objGrid.Row - objBrowse.iGrid_LinhaInicial)
            End If

            objBrowse.objGrid.Row = objBrowse.iGrid_LinhaInicial

        ElseIf objBrowse.objGrid.Row > objBrowse.iGrid_LinhaFinal Then

            If objBrowse.objBarraScroll.Value + (objBrowse.objGrid.Row - objBrowse.iGrid_LinhaFinal) > objBrowse.objBarraScroll.Max Then
                objBrowse.objBarraScroll.Value = objBrowse.objBarraScroll.Max
            Else
                objBrowse.objBarraScroll.Value = objBrowse.objBarraScroll.Value + (objBrowse.objGrid.Row - objBrowse.iGrid_LinhaFinal)
            End If

            objBrowse.objGrid.Row = objBrowse.iGrid_LinhaFinal

        End If
        
        objBrowse.objGrid.ColSel = objBrowse.objGrid.Cols - 1
        
    End If

    Browse_GridBrowse_RowColChange = SUCESSO

End Function

Function Browse_Ordenacao_Click(objBrowse As AdmBrowse) As Long

Dim lErro As Long

On Error GoTo Erro_Browse_Ordenacao_Click

    'se não está fazendo a carga inicial da tela e nem a reinicialização da ordenacao
    If objBrowse.iAlterado >= 0 Then

        objBrowse.objForm.MousePointer = vbHourglass

        lErro = Browse_Inicializa_Campos(objBrowse)
        If lErro <> SUCESSO Then Error 9124

        lErro = Browse_Inicializa_Comando_SQL(objBrowse)
        If lErro <> SUCESSO Then Error 9111

        objBrowse.iBind = 0

        If objBrowse.objBarraScroll.Value = objBrowse.objBarraScroll.Min Then
            objBrowse.objGrid.Row = 1
            Call Browse_BarraScroll_Change(objBrowse)
        Else
            objBrowse.objGrid.Row = 1
        End If

        objBrowse.objGrid.ColSel = objBrowse.objGrid.Cols - 1

        objBrowse.objForm.MousePointer = vbDefault

    End If

    Browse_Ordenacao_Click = SUCESSO

    Exit Function

Erro_Browse_Ordenacao_Click:

    Browse_Ordenacao_Click = Err

    Select Case Err

        Case 9111, 9124
            objBrowse.objForm.MousePointer = vbDefault

        Case Else
            objBrowse.objForm.MousePointer = vbDefault
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142069)

    End Select

    Exit Function

End Function


Function Browse_GridBrowse_Scroll(objBrowse As AdmBrowse) As Long

    If objBrowse.objGrid.TopRow <> objBrowse.iGrid_LinhaInicial Then
        objBrowse.objGrid.TopRow = objBrowse.iGrid_LinhaInicial
    End If
    
    Browse_GridBrowse_Scroll = SUCESSO
    
End Function


Function Browse_BarraScroll_Change(objBrowse As AdmBrowse) As Long

Dim lErro As Long
Dim lIncremento As Long

On Error GoTo Erro_Browse_BarraScroll_Change

    If objBrowse.iSelecaoExterna <> 1 Then

        'retorna o número de lotes pendentes
        lErro = CF("Browse_Le_NumRegistros", lNumReg, objBrowse.lComando1, objBrowse.sSQLCount, avCampo1)
        If lErro <> SUCESSO Then Error 9027

        If lNumReg > 64000 Then
            objBrowse.objBarraScroll.Max = 32000
        Else
            If lNumReg > objBrowse.objBarraScroll.LargeChange Then
                objBrowse.objBarraScroll.Max = objBrowse.objBarraScroll.Min + lNumReg - objBrowse.objBarraScroll.LargeChange
            Else
                objBrowse.objBarraScroll.Max = objBrowse.objBarraScroll.Min
            End If
        End If

        lIncremento = Abs(objBrowse.lOldScrollValue - objBrowse.objBarraScroll.Value)

        'calcula o local da tabela a partir do qual os registros serão exibidos
        If lIncremento = objBrowse.objBarraScroll.SmallChange Or lIncremento = objBrowse.objBarraScroll.LargeChange Then
            objBrowse.lPosicao = objBrowse.lPosicao + (objBrowse.objBarraScroll.Value - objBrowse.lOldScrollValue)
        Else

            Select Case objBrowse.objBarraScroll.Value

                Case objBrowse.objBarraScroll.Max
                    objBrowse.lPosicao = lNumReg - objBrowse.objBarraScroll.LargeChange + 1
                    If objBrowse.lPosicao <= 0 Then objBrowse.lPosicao = 1

                Case objBrowse.objBarraScroll.Min
                    objBrowse.lPosicao = 1

                Case Else
                    objBrowse.lPosicao = CLng((objBrowse.objBarraScroll.Value - objBrowse.objBarraScroll.Min) * lNumReg / (objBrowse.objBarraScroll.Max + objBrowse.objBarraScroll.LargeChange - objBrowse.objBarraScroll.Min)) + 1

            End Select

        End If

        lIncremento = objBrowse.objBarraScroll.Value - objBrowse.lOldScrollValue

        'se estiver fazendo scroll pela barra de scroll,
        'fazer com que a seleção acompanhe o elemento selecionado
        If objBrowse.iTeclado <> 1 Then
            If objBrowse.objGrid.Row - lIncremento < objBrowse.iGrid_LinhaInicial Then
                objBrowse.objGrid.Row = objBrowse.iGrid_LinhaInicial
            ElseIf objBrowse.objGrid.Row - lIncremento > objBrowse.iGrid_LinhaFinal Then
                objBrowse.objGrid.Row = objBrowse.iGrid_LinhaFinal
            Else
                objBrowse.objGrid.Row = objBrowse.objGrid.Row - lIncremento
            End If
            objBrowse.objGrid.ColSel = objBrowse.objGrid.Cols - 1
        Else
            objBrowse.iTeclado = 0
        End If

        objBrowse.lOldScrollValue = objBrowse.objBarraScroll.Value

        lErro = Browse_Exibe_Registros(objBrowse)
        If lErro <> SUCESSO Then Error 9168

    End If
    
    Browse_BarraScroll_Change = SUCESSO
    
    Exit Function

Erro_Browse_BarraScroll_Change:

    Browse_BarraScroll_Change = Err

    Select Case Err

        Case 9027, 9168

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142070)

    End Select

    Exit Function

End Function

Function Browse_Trata_Parametros(objBrowse As AdmBrowse, Optional ByVal iPesquisa As Integer = 0) As Long

Dim lErro As Long
Dim lPosAux As Long
Dim lPos As Long

On Error GoTo Erro_Browse_Trata_Parametros

    If iPesquisa = 0 Then

        'inicializa as strings dos comandos SQL utilizados no Browse
        lErro = Browse_Trata_Parametros1(objBrowse)
        If lErro <> SUCESSO Then Error 9254

    End If
    
    objBrowse.iSelecaoExterna = 1

    'retorna o número de lotes pendentes
    lErro = CF("Browse_Le_NumRegistros", lNumReg, objBrowse.lComando1, objBrowse.sSQLCount, avCampo1)
    If lErro <> SUCESSO Then Error 9153
    
    lErro = Browse_Retorna_Posicao_Registro(objBrowse)
    If lErro <> SUCESSO Then Error 9154
        
    If lNumReg > 64000 Then
        objBrowse.objBarraScroll.Max = 32000
    Else
        If lNumReg > objBrowse.objBarraScroll.LargeChange Then
            objBrowse.objBarraScroll.Max = objBrowse.objBarraScroll.Min + lNumReg - objBrowse.objBarraScroll.LargeChange
        Else
            objBrowse.objBarraScroll.Max = objBrowse.objBarraScroll.Min
        End If
    End If
    
    If objBrowse.lPosicao > lNumReg Then objBrowse.lPosicao = lNumReg
    
    If lNumReg < objBrowse.objBarraScroll.LargeChange Then
        objBrowse.objGrid.Row = objBrowse.objGrid.TopRow + objBrowse.lPosicao - 1
        objBrowse.lPosicao = 1
    ElseIf objBrowse.lPosicao > lNumReg - objBrowse.objBarraScroll.LargeChange + 1 Then
        lPosAux = objBrowse.lPosicao - (lNumReg - objBrowse.objBarraScroll.LargeChange + 1)
        objBrowse.objGrid.Row = objBrowse.objGrid.TopRow + lPosAux
        objBrowse.lPosicao = lNumReg - objBrowse.objBarraScroll.LargeChange + 1
    Else
        objBrowse.objGrid.Row = objBrowse.objGrid.TopRow
    End If
        
    objBrowse.objGrid.ColSel = objBrowse.objGrid.Cols - 1
        
    If lNumReg > objBrowse.objBarraScroll.LargeChange Then
        lPos = CInt(objBrowse.objBarraScroll.Min + ((objBrowse.lPosicao - 1) * (objBrowse.objBarraScroll.Max + objBrowse.objBarraScroll.LargeChange - objBrowse.objBarraScroll.Min) / lNumReg))
        If lPos > objBrowse.objBarraScroll.Max Then lPos = objBrowse.objBarraScroll.Max
        objBrowse.objBarraScroll.Value = lPos
    Else
        objBrowse.objBarraScroll.Value = objBrowse.objBarraScroll.Min
    End If
        
    lErro = Browse_Exibe_Registros(objBrowse)
    If lErro <> SUCESSO Then Error 9169
        
    objBrowse.lOldScrollValue = objBrowse.objBarraScroll.Value
            
    objBrowse.iSelecaoExterna = 0

    Browse_Trata_Parametros = SUCESSO

    Exit Function

Erro_Browse_Trata_Parametros:

    Browse_Trata_Parametros = Err

    objBrowse.iSelecaoExterna = 0

    Select Case Err

        Case 9153, 9154, 9169, 9254

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142071)

    End Select

    If Err <> 9254 Then
        Call Comando_Fechar(objBrowse.lComando)
        objBrowse.lComando = 0
        Call Comando_Fechar(objBrowse.lComando1)
        objBrowse.lComando1 = 0
        Call Comando_Fechar(objBrowse.lComando2)
        objBrowse.lComando2 = 0
        Call Comando_Fechar(objBrowse.lComando3)
        objBrowse.lComando3 = 0
    End If

    Exit Function

End Function

Function Browse_Trata_Parametros1(objBrowse As AdmBrowse) As Long

Dim lErro As Long
Dim lPosAux As Long

On Error GoTo Erro_Browse_Trata_Parametros1

    Set objBrowse.objRotinaModulo = CreateObject(objBrowse.objBrowseArquivo.sProjeto & "." & objBrowse.objBrowseArquivo.sClasse)

    If Len(objBrowse.objBrowseArquivo.sSelecaoSQL) > 0 And Len(objBrowse.sSelecaoSQL2) > 0 Then
        objBrowse.sSelecaoSQL = "(" & objBrowse.sSelecaoSQL2 & ") AND (" & objBrowse.objBrowseArquivo.sSelecaoSQL & ")"
    ElseIf Len(objBrowse.sSelecaoSQL2) > 0 Then
        objBrowse.sSelecaoSQL = objBrowse.sSelecaoSQL2
    Else
        objBrowse.sSelecaoSQL = objBrowse.objBrowseArquivo.sSelecaoSQL
    End If
    
    If objBrowse.sSelecaoSQL <> objBrowse.sSelecaoSQL3 Then
    
        objBrowse.sSelecaoSQL3 = objBrowse.sSelecaoSQL
        
        If iBind = 0 Then
        
            'reinicializa os campos que vão receber os dados do bd já que vai ocorrer um novo bind
            lErro = Browse_Reinicialza_Avcampo(objBrowse)
            If lErro <> SUCESSO Then Error 20805
                
            objBrowse.iBind = 1
    
        End If
    
    End If
    
    objBrowse.sSelecaoSQL2 = ""
    
    'inicializa o comando SQL que será usado para exibir os registros da tabela
    lErro = Browse_Inicializa_Comando_SQL(objBrowse)
    If lErro <> SUCESSO Then Error 9040
    
    objBrowse.iBind = 0

    'inicializa o comando SQL que será usado para contar o número de registros da tabela
    lErro = Browse_Inicializa_Comando_SQL_Count(objBrowse)
    If lErro <> SUCESSO Then Error 9252

    lErro = Browse_Inicializa_Grid_Lote(objBrowse)
    If lErro <> SUCESSO Then Error 9020

    lErro = Browse_Inicializa_ScrollBar_Lote(objBrowse)
    If lErro <> SUCESSO Then Error 9019

    Browse_Trata_Parametros1 = SUCESSO

    Exit Function

Erro_Browse_Trata_Parametros1:

    Browse_Trata_Parametros1 = Err

    Select Case Err

        Case 9019, 9020, 9040, 9252, 20805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142072)

    End Select

    Call Comando_Fechar(objBrowse.lComando)
    objBrowse.lComando = 0
    Call Comando_Fechar(objBrowse.lComando1)
    objBrowse.lComando1 = 0
    Call Comando_Fechar(objBrowse.lComando2)
    objBrowse.lComando2 = 0
    Call Comando_Fechar(objBrowse.lComando3)
    objBrowse.lComando3 = 0

    Exit Function

End Function

Function Browse_Retorna_Posicao_Registro(objBrowse As AdmBrowse) As Long

Dim lErro As Long
Dim objBrowseIndice As AdmBrowseIndice
Dim colBrowseIndiceSegmentos As New Collection
Dim sSQL As String
Dim iIndice As Integer
Dim objBrowseIndiceSegmentos As AdmBrowseIndiceSegmentos
Dim colSegmentos As New Collection
Dim objBrowseValorCampo As AdmBrowseValorCampo
Dim lPosicao As Long
Dim iUsuario As Integer
Dim iIndice_Faltando As Integer '*** se algum dos segmentos de ordenacao não faz parte dos parametros passados pela tela ==> não pode utiliza-lo como posicionador

On Error GoTo Erro_Browse_Retorna_Posicao_Registro

    iUsuario = 1

    For Each objBrowseIndice In objBrowse.colBrowseIndice
        If objBrowseIndice.iIndice = objBrowse.objOrdenacao.ItemData(objBrowse.objOrdenacao.ListIndex) Then
            iUsuario = -1
            Exit For
        End If
    Next

    'se for uma ordenacao cadastrada pelo usuario
    If iUsuario = 1 Then
    
        For Each objBrowseIndice In objBrowse.colBrowseIndiceUsuario
            If objBrowseIndice.iIndice + objBrowse.colBrowseIndice.Count = objBrowse.objOrdenacao.ItemData(objBrowse.objOrdenacao.ListIndex) Then Exit For
        Next
        
    End If

    '*** estava posicionado abaixo do preenchimento de sSQL e foi trazido para cima pois quando um dos segmentos do indice de ordenação não fizer parte dos parametros da tela ==> não pode utilizar a ordenação na expressao SQL para posicionar o registro

    'se é uma ordenação pre-cadastrada
    If iUsuario = -1 Then

        lErro = CF("BrowseIndiceSegmentos_Le", objBrowse.objForm.Name, objBrowseIndice.iIndice, colBrowseIndiceSegmentos, objBrowse.colBrowseIndice)
        If lErro <> SUCESSO Then Error 9147
        
    Else
    
        'se for uma ordenacao cadastrada pelo usuario
        lErro = CF("BrowseIndiceSegmentos_Le", objBrowse.objForm.Name, objBrowseIndice.iIndice, colBrowseIndiceSegmentos, objBrowse.colBrowseIndiceUsuario)
        If lErro <> SUCESSO Then Error 60821
    
    End If
    
    For iIndice = 1 To colBrowseIndiceSegmentos.Count
    
        For Each objBrowseIndiceSegmentos In colBrowseIndiceSegmentos
        
            If objBrowseIndiceSegmentos.iPosicaoCampo = iIndice Then Exit For
            
        Next
    
        For Each objBrowseValorCampo In objBrowse.colBrowseValorCampo
        
            If objBrowseValorCampo.sNomeCampo = objBrowseIndiceSegmentos.sNomeCampo Then Exit For
    
        Next
        
        If objBrowseValorCampo Is Nothing Then
            iIndice_Faltando = 1
            Set colSegmentos = New Collection
            Exit For
        End If
        
        colSegmentos.Add objBrowseValorCampo.vValorCampo
        
    Next

    If iIndice_Faltando = 0 Then
        If Len(objBrowse.sSelecaoSQL) > 0 And Len(objBrowse.sSelecaoSQL1) > 0 Then
            sSQL = objBrowse.sNomeTabela & " WHERE (" & objBrowse.sSelecaoSQL & ") AND (" & objBrowse.sSelecaoSQL1 & ") AND (" & objBrowseIndice.sSelecaoSQL & ")"
        ElseIf Len(objBrowse.sSelecaoSQL) > 0 Then
            sSQL = objBrowse.sNomeTabela & " WHERE (" & objBrowse.sSelecaoSQL & ") AND (" & objBrowseIndice.sSelecaoSQL & ")"
        ElseIf Len(objBrowse.sSelecaoSQL1) > 0 Then
            sSQL = objBrowse.sNomeTabela & " WHERE (" & objBrowse.sSelecaoSQL1 & ") AND (" & objBrowseIndice.sSelecaoSQL & ")"
        Else
            sSQL = objBrowse.sNomeTabela & " WHERE (" & objBrowseIndice.sSelecaoSQL & ")"
        End If
    Else
        If Len(objBrowse.sSelecaoSQL) > 0 And Len(objBrowse.sSelecaoSQL1) > 0 Then
            sSQL = objBrowse.sNomeTabela & " WHERE (" & objBrowse.sSelecaoSQL & ") AND (" & objBrowse.sSelecaoSQL1 & ")"
        ElseIf Len(objBrowse.sSelecaoSQL) > 0 Then
            sSQL = objBrowse.sNomeTabela & " WHERE (" & objBrowse.sSelecaoSQL & ")"
        ElseIf Len(objBrowse.sSelecaoSQL1) > 0 Then
            sSQL = objBrowse.sNomeTabela & " WHERE (" & objBrowse.sSelecaoSQL1 & ")"
        Else
            sSQL = objBrowse.sNomeTabela
        End If
    End If
    lErro = CF("Browse_Le_NumRegistros1", sSQL, lPosicao, colSegmentos, objBrowse.colSelecao, objBrowse)
    If lErro <> SUCESSO Then Error 9152
    
    objBrowse.lPosicao = lPosicao + 1
    
    Browse_Retorna_Posicao_Registro = SUCESSO

    Exit Function

Erro_Browse_Retorna_Posicao_Registro:

    Browse_Retorna_Posicao_Registro = Err

    Select Case Err

        Case 9147, 9152, 60821

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142073)

    End Select

    Exit Function

End Function

Function Browse_Exibe_Registros(objBrowse As AdmBrowse) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim colReg As New Collection
Dim colRegCampo As Collection
Dim iLinha1 As Integer
Dim objValorCampo As AdmValorCampo
Dim vCampo As Variant
Dim vCampo1 As Variant
Dim iIndice As Integer
Dim iProximoCampo As Integer
Dim dSaldoAtraso As Double
Dim lCliente As Long
Dim iVermelho As Integer
        
On Error GoTo Erro_Browse_Exibe_Registros
        
    iProximoCampo = 0
        
    objBrowse.objGrid.Redraw = False
        
    'le os registros a partir da posicao selecionada
    lErro = CF("Browse_Le_Posicionado", colReg, objBrowse.lPosicao, objBrowse.iGrid_LinhasExibidas, objBrowse.lComando, objBrowse.colValorCampo, objBrowse.sSQL, avCampo)
    If lErro <> SUCESSO Then Error 9018

    Set objBrowse.colReg = colReg

    iLinha = objBrowse.objGrid.TopRow
    
    For Each colRegCampo In colReg
    
        'VERMELHO
        If objBrowse.objForm.Name = "ClientesLista" And giCliAtrasoDestacar <> 0 Then
        
            iVermelho = 0
            dSaldoAtraso = 0
        
            lCliente = colRegCampo(colRegCampo.Count)

            lErro = CF("Cliente_EmAtraso_Destacar2", lCliente, iVermelho)
            If lErro <> SUCESSO Then gError 207915
'
'            lErro = CF("Cliente_Obtem_SaldoAtraso", lCliente, dSaldoAtraso)
'            If lErro <> SUCESSO Then gError 207915
'
'            If dSaldoAtraso > DELTA_VALORMONETARIO Then
'                iVermelho = 1
'            End If
            
            colRegCampo.Remove (colRegCampo.Count)
            
        End If
    
    
        iIndice = 0
       
        For Each vCampo In colRegCampo
    
            iIndice = iIndice + 1
    
            If iProximoCampo = 1 Then
            
                iProximoCampo = 0
    
            Else
            
                Set objValorCampo = objBrowse.colValorCampo.Item(iIndice)

                'VERMELHO
                If objBrowse.objForm.Name = "ClientesLista" And giCliAtrasoDestacar <> 0 Then

                    lErro = objBrowse.objRotinaModulo.Browse_Formata_Campo(objBrowse, colRegCampo, objValorCampo, vCampo, iProximoCampo, iLinha, iIndice, iExercicio, iPeriodo, sNomeExterno, iVermelho)
                    If lErro <> SUCESSO Then gError 207916
                    
                Else
                
                    lErro = objBrowse.objRotinaModulo.Browse_Formata_Campo(objBrowse, colRegCampo, objValorCampo, vCampo, iProximoCampo, iLinha, iIndice, iExercicio, iPeriodo, sNomeExterno)
                    If lErro <> SUCESSO Then gError 9262
                
                End If


            End If
        
        Next
        
        iLinha = iLinha + 1
       
    Next

    For iLinha1 = iLinha To objBrowse.iGrid_LinhaFinal
        For iIndice = 0 To objBrowse.objGrid.Cols - 1
            objBrowse.objGrid.TextMatrix(iLinha1, iIndice) = ""
        Next
    Next
        
    objBrowse.objGrid.Redraw = True
        
    Browse_Exibe_Registros = SUCESSO
        
    Exit Function

Erro_Browse_Exibe_Registros:

    Browse_Exibe_Registros = Err

    Select Case Err

        Case 9018, 9262
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142074)

    End Select

    Exit Function
        
End Function

Private Function Browse_Reinicializa_Ordenacao(objBrowse As AdmBrowse) As Long
'reinicializa a ordenacao para o usuario em questão

Dim lErro As Long
Dim objBrowseIndice As AdmBrowseIndice
Dim iIndice As Integer
Dim objBrowseUsuarioOrdenacao As New AdmBrowseUsuarioOrdenacao
Dim iAchou As Integer
Dim iAlterado As Integer

On Error GoTo Erro_Browse_Reinicializa_Ordenacao

    If objBrowse.iAlteradoOrdenacao = 1 Then
    
        iAlterado = objBrowse.iAlterado
    
        objBrowse.iAlterado = -1
        
        iIndice = objBrowse.objOrdenacao.ListIndex
    
        objBrowse.objOrdenacao.Clear
    
        For Each objBrowseIndice In objBrowse.colBrowseIndice
            objBrowse.objOrdenacao.AddItem objBrowseIndice.sNomeIndice
            objBrowse.objOrdenacao.ItemData(objBrowse.objOrdenacao.NewIndex) = objBrowseIndice.iIndice
        Next
    
        For Each objBrowseIndice In objBrowse.colBrowseIndiceUsuario
            objBrowse.objOrdenacao.AddItem objBrowseIndice.sNomeIndice
            objBrowse.objOrdenacao.ItemData(objBrowse.objOrdenacao.NewIndex) = objBrowseIndice.iIndice + objBrowse.colBrowseIndice.Count
        Next
    
        If objBrowse.objOrdenacao.ListCount > iIndice Then
            objBrowse.objOrdenacao.ListIndex = iIndice
        Else
            objBrowse.objOrdenacao.ListIndex = 0
        End If
    
        objBrowse.iAlterado = iAlterado

    End If

    Browse_Reinicializa_Ordenacao = SUCESSO

    Exit Function

Erro_Browse_Reinicializa_Ordenacao:

    Browse_Reinicializa_Ordenacao = Err

    objBrowse.iAlterado = iAlterado

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142075)

    End Select

    Exit Function

End Function

