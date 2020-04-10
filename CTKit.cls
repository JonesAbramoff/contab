VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTKit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

'Formato para quantidades de Produtos
Const FORMATO_ESTOQUE_KIT = "#,##0.0####"


Dim iAlterado As Integer
Dim iClasseAtual As Integer
Dim sProdutoSelAnterior As String
Public colComponentes As New Collection 'cada elemento é objProdutoKit e guarda informacoes correspondentes aos dados de cada nó da treeview EstruturaProduto
Dim iProxChave As Integer      'cada elemento da colecao acima recebe um identificador composto por "X" concantenado com um inteiro.
                                'Os nodes da treeview EstruturaProduto tem no seu tag este identificador, o que permite manter a correlacao entre a treeview e a Collection
                                'iMaiorChave guarda o maior numero já utilizado para geracao de um identificador (na estrutura que está sendo editada.)
                                'na key de cada node fica "X" & nivel (em dois caracteres*) & seq + 1 (em dois caracteres*) & codigo do produto no formato do bd
                                '       * = se for menor que dez completar c/ CSTR(0))

'variaveis auxiliares para recalculo de nivel e sequencial
Dim aNivelSequencial(NIVEL_MAXIMO_KIT) As Integer 'para cada nivel guarda o maior sequencial
Dim iUltimoNivel As Integer
Dim sProdutoAnt As String


Private WithEvents objEventoKit As AdmEvento
Attribute objEventoKit.VB_VarHelpID = -1
Private WithEvents objEventoVersao As AdmEvento
Attribute objEventoVersao.VB_VarHelpID = -1
Private WithEvents objEventoVersaoKitComp As AdmEvento 'Inserido por Wagner
Attribute objEventoVersaoKitComp.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoComponente As AdmEvento
Attribute objEventoComponente.VB_VarHelpID = -1

Sub Recalcula_Nivel_Sequencial()
'(re)calcula niveis e sequencias de toda a estrutura
'deve ser chamada apos a remocao de algum node

Dim iIndice As Integer

    If EstruturaProduto.Nodes.Count = 0 Then Exit Sub

    For iIndice = LBound(aNivelSequencial) To UBound(aNivelSequencial)
        aNivelSequencial(iIndice) = 0
    Next

    iUltimoNivel = 0

    'chamar rotina que recalcula recursivamente os campos nivel e sequencial
    Call Calcula_Nivel_Sequencial(EstruturaProduto.Nodes.Item(1), 0, 0)

End Sub

Sub Calcula_Nivel_Sequencial(objNode As Node, iNivel As Integer, iPosicaoAtual As Integer)
'parte recursiva do recalculo de nivel e sequencial, atuando a partir do node passado
'iNivel informa o nivel deste node

Dim objProdutoKit As New ClassProdutoKit
Dim sChave1 As String

    sChave1 = objNode.Tag

    Set objProdutoKit = colComponentes.Item(sChave1)

    aNivelSequencial(iNivel) = aNivelSequencial(iNivel) + 1

    objProdutoKit.iSeq = aNivelSequencial(iNivel)

    If iNivel > 0 Then
        objProdutoKit.iSeqPai = aNivelSequencial(iNivel - 1)
    Else
        objProdutoKit.iSeqPai = 0
    End If
    
    objProdutoKit.iPosicaoArvore = iPosicaoAtual
    iPosicaoAtual = iPosicaoAtual + 1

    objProdutoKit.iNivel = iNivel
    
    colComponentes.Remove sChave1
    colComponentes.Add objProdutoKit, sChave1

    If objNode.Children > 0 Then
        Call Calcula_Nivel_Sequencial(objNode.Child, iNivel + 1, iPosicaoAtual)
    End If

    If objNode.Index <> objNode.LastSibling.Index Then Call Calcula_Nivel_Sequencial(objNode.Next, iNivel, iPosicaoAtual)

    If iNivel > iUltimoNivel Then iUltimoNivel = iNivel
   
End Sub

Sub Calcula_Proxima_Chave(sChave As String)

Dim iNumero As Integer

    iNumero = iProxChave
    iProxChave = iProxChave + 1
    sChave = "X" & right$(CStr(100000 + iNumero), 5)

End Sub

Sub Remove_Filhos(objNode As Node)
'remove objNode, seus irmaos e filhos de colComponentes

        colComponentes.Remove (objNode.Tag)

        If objNode.Children > 0 Then

            Call Remove_Filhos(objNode.Child)

        End If

        If objNode <> objNode.LastSibling Then Call Remove_Filhos(objNode.Next)

    Exit Sub

End Sub

Function Extrai_ProdutoKit(objProdutoKit As ClassProdutoKit) As Long
'preenche objProdutoKit à partir dos dados do frame de "componentes"
'serve para inclusao ou alteracao de componente

Dim lErro As Long, iProdutoPreenchido As Integer
Dim sProdutoFormatado As String, objProduto As New ClassProduto
Dim sChave As String
Dim objKit As New ClassKit 'Inserido por Wagner

On Error GoTo Erro_Extrai_ProdutoKit

    'Formata o código do Produto como no BD
    lErro = CF("Produto_Formata", ProdutoSel.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 19378

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 19379

    If Len(Trim(Quantidade.Text)) = 0 Then gError 19380
    If Composicao.ListIndex = -1 Then gError 19381
    If Len(Trim(SiglaUM.Text)) = 0 Then gError 19382
    
    objProdutoKit.sProduto = sProdutoFormatado
    objProdutoKit.dQuantidade = CDbl(Quantidade.Text)
    objProdutoKit.dCustoStandard = StrParaDbl(CustoStandard.Text)
    objProdutoKit.iComposicao = Composicao.ItemData(Composicao.ListIndex)
    objProdutoKit.sUnidadeMed = SiglaUM.Text
    objProdutoKit.sProdutoDesc = DescProdutoSel.Caption
    
    If Len(Trim(Nivel.Caption)) > 0 Then objProdutoKit.iNivel = CInt(Nivel.Caption)
    If Len(Trim(Sequencial.Caption)) > 0 Then objProdutoKit.iSeq = CInt(Sequencial.Caption)
    If Len(Trim(PercentualPerda.Text)) > 0 Then objProdutoKit.dPercentualPerda = StrParaDbl(PercentualPerda / 100)
   
    '############################
    'Inserido por Wagner 23/11/04
    If Len(Trim(VersaoKitComp.Text)) > 0 And objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
    
        objKit.sVersao = VersaoKitComp.Text
        objKit.sProdutoRaiz = sProdutoFormatado
    
        lErro = CF("Kit_Le_Componentes", objKit)
        If lErro <> SUCESSO And lErro <> 21831 Then gError 131300
        
        'Se o Kit não está cadastrado => Erro
        If lErro <> SUCESSO Then gError 131301
    
        objProdutoKit.sVersaoKitComp = VersaoKitComp.Text
        
    End If
    '############################
   
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 19383
    
    'se o produto não estiver cadastrado
    If lErro = 28030 Then gError 19384

    objProdutoKit.iClasseUM = objProduto.iClasseUM

    lErro = CF2(Me, "KT_Extrai_ProdutoKit", objProdutoKit)
    If lErro <> SUCESSO Then gError 19378
    
    Extrai_ProdutoKit = SUCESSO

    Exit Function

Erro_Extrai_ProdutoKit:

    Extrai_ProdutoKit = gErr

    Select Case gErr

        Case 19378, 19383, 131300 'Alterado por Wagner
        
        Case 19379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTOKIT_NAO_PREENCHIDO", gErr)

        Case 19380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_INFORMADA", gErr)
            
        Case 19381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPOSICAO_NAO_INFORMADA", gErr)
            
        Case 19382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLAUM_NAO_INFORMADA", gErr)
            
        Case 19384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case 131301 'Inserido por Wagner
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_INEXISTENTE", gErr, objKit.sVersao, objKit.sProdutoRaiz)
            
        Case 131308 'Inserido por Wagner
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VERSAO_KIT_NAO_PREENCHIDA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162075)

    End Select

    Exit Function

End Function

Function Preenche_ProdutoKit(objProdutoKit As ClassProdutoKit) As Long
'preenche o frame de "componentes" à partir dos dados de objProdutoKit

Dim lErro As Long, iClasseUM As Integer
Dim objProduto As New ClassProduto, objClasse As New ClassClasseUM
Dim colSiglas As New Collection, objUnidadeDeMedida As New ClassUnidadeDeMedida

On Error GoTo Erro_Preenche_ProdutoKit

    lErro = CF("Traz_Produto_MaskEd", objProdutoKit.sProduto, ProdutoSel, DescProdutoSel)
    If lErro <> SUCESSO Then gError 19377
    'Formata_Estoque(objProdutoKit.dQuantidade)
    Quantidade.Text = Format(objProdutoKit.dQuantidade, FORMATO_ESTOQUE_KIT)
    CustoStandard.Text = CStr(objProdutoKit.dCustoStandard)
    Composicao.ListIndex = objProdutoKit.iComposicao
    PercentualPerda.Text = CStr(objProdutoKit.dPercentualPerda * 100)
    iClasseUM = objProdutoKit.iClasseUM
    
    '#####################################
    'Inserido por Wagner - 23/11/04
    If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
        VersaoKitComp.Clear
        VersaoKitComp.Visible = False
        LabelVersaoComp.Visible = False
    Else
        VersaoKitComp.Visible = True
        LabelVersaoComp.Visible = True
        Call Carrega_ComboVersoes(objProdutoKit.sProduto)
        Call VersaoKitComp_Seleciona(objProdutoKit.sVersaoKitComp)
    
    End If
    '#####################################
    
    'se a classe UM ainda nao está preenchida
    If iClasseUM = 0 Then

        objProduto.sCodigo = objProdutoKit.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 19376

        If lErro = 28030 Then gError 21804

        iClasseUM = objProduto.iClasseUM

    End If

    objClasse.iClasse = iClasseUM

    'preencher a combo de UMs
    lErro = Preenche_Combo_UMs(iClasseUM)
    If lErro <> SUCESSO Then gError 19399

    SiglaUM.Text = objProdutoKit.sUnidadeMed

    'obter a descricao da UM
    objUnidadeDeMedida.sSigla = objProdutoKit.sUnidadeMed
    objUnidadeDeMedida.iClasse = iClasseUM

    lErro = CF("UM_Le", objUnidadeDeMedida)
    If lErro <> SUCESSO And lErro <> 23775 Then gError 21808

    If lErro = 23775 Then gError 21809

    NomeUM.Caption = objUnidadeDeMedida.sNome

    Nivel.Caption = CStr(objProdutoKit.iNivel)
    Sequencial.Caption = CStr(objProdutoKit.iSeq)

    lErro = CF2(Me, "KT_Preenche_ProdutoKit", objProdutoKit)
    If lErro <> SUCESSO Then gError 19399
    
    sProdutoSelAnterior = ProdutoSel.Text

    iAlterado = 0

    Preenche_ProdutoKit = SUCESSO

    Exit Function

Erro_Preenche_ProdutoKit:

    Preenche_ProdutoKit = gErr

    Select Case gErr

        Case 19376, 19377, 21806, 21808, 19399

        Case 21804
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 21807
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_INEXISTENTE", gErr, objClasse.iClasse)

        Case 21809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_SIGLAUM_INEXISTENTE", gErr, objUnidadeDeMedida.iClasse, objUnidadeDeMedida.sSigla)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162076)

    End Select

    Exit Function

End Function

'--- inicio dos properties get dos controles da tela

Public Property Get LabelFator() As Object
     Set LabelFator = objUserControl.Controls("LabelFator")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get MaoDeObra() As Object
     Set MaoDeObra = objUserControl.Controls("MaoDeObra")
End Property

Public Property Get Energia() As Object
     Set Energia = objUserControl.Controls("Energia")
End Property

Public Property Get BPF() As Object
     Set BPF = objUserControl.Controls("BPF")
End Property

Public Property Get Agua() As Object
     Set Agua = objUserControl.Controls("Agua")
End Property

Public Property Get PesoFator5() As Object
     Set PesoFator5 = objUserControl.Controls("PesoFator5")
End Property

Public Property Get PesoFator6() As Object
     Set PesoFator6 = objUserControl.Controls("PesoFator6")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get VersaoKitComp() As Object
     Set VersaoKitComp = objUserControl.Controls("VersaoKitComp")
End Property

Public Property Get Composicao() As Object
     Set Composicao = objUserControl.Controls("Composicao")
End Property

Public Property Get SiglaUM() As Object
     Set SiglaUM = objUserControl.Controls("SiglaUM")
End Property

Public Property Get BotaoAlterar() As Object
     Set BotaoAlterar = objUserControl.Controls("BotaoAlterar")
End Property

Public Property Get BotaoRemover() As Object
     Set BotaoRemover = objUserControl.Controls("BotaoRemover")
End Property

Public Property Get BotaoIncluir() As Object
     Set BotaoIncluir = objUserControl.Controls("BotaoIncluir")
End Property

Public Property Get ProdutoSel() As Object
     Set ProdutoSel = objUserControl.Controls("ProdutoSel")
End Property

Public Property Get Quantidade() As Object
     Set Quantidade = objUserControl.Controls("Quantidade")
End Property

Public Property Get PercentualPerda() As Object
     Set PercentualPerda = objUserControl.Controls("PercentualPerda")
End Property

Public Property Get CustoStandard() As Object
     Set CustoStandard = objUserControl.Controls("CustoStandard")
End Property

Public Property Get LabelVersaoComp() As Object
     Set LabelVersaoComp = objUserControl.Controls("LabelVersaoComp")
End Property

Public Property Get Nivel() As Object
     Set Nivel = objUserControl.Controls("Nivel")
End Property

Public Property Get NomeUM() As Object
     Set NomeUM = objUserControl.Controls("NomeUM")
End Property

Public Property Get DescProdutoSel() As Object
     Set DescProdutoSel = objUserControl.Controls("DescProdutoSel")
End Property

Public Property Get Sequencial() As Object
     Set Sequencial = objUserControl.Controls("Sequencial")
End Property

Public Property Get ComponenteLabel() As Object
     Set ComponenteLabel = objUserControl.Controls("ComponenteLabel")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get EstruturaProduto() As Object
     Set EstruturaProduto = objUserControl.Controls("EstruturaProduto")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get BotaoRoteiros() As Object
     Set BotaoRoteiros = objUserControl.Controls("BotaoRoteiros")
End Property

Public Property Get Situacao() As Object
     Set Situacao = objUserControl.Controls("Situacao")
End Property

Public Property Get BotaoKits() As Object
     Set BotaoKits = objUserControl.Controls("BotaoKits")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get VersaoFormPreco() As Object
     Set VersaoFormPreco = objUserControl.Controls("VersaoFormPreco")
End Property

Public Property Get UpDown1() As Object
     Set UpDown1 = objUserControl.Controls("UpDown1")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

Public Property Get Versao() As Object
     Set Versao = objUserControl.Controls("Versao")
End Property

Public Property Get LabelObeservações() As Object
     Set LabelObeservações = objUserControl.Controls("LabelObeservações")
End Property

Public Property Get LabelVersao() As Object
     Set LabelVersao = objUserControl.Controls("LabelVersao")
End Property

Public Property Get Descricao() As Object
     Set Descricao = objUserControl.Controls("Descricao")
End Property

Public Property Get ProdutoLbl() As Object
     Set ProdutoLbl = objUserControl.Controls("ProdutoLbl")
End Property

Public Property Get BotaoImprimir() As Object
     Set BotaoImprimir = objUserControl.Controls("BotaoImprimir")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Agua_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Agua_GotFocus()

    Call MaskEdBox_TrataGotFocus(Agua)

End Sub

Public Sub PesoFator5_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PesoFator5_GotFocus()

    Call MaskEdBox_TrataGotFocus(PesoFator5)

End Sub

Public Sub PesoFator6_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PesoFator6_GotFocus()

    Call MaskEdBox_TrataGotFocus(PesoFator6)

End Sub

Public Sub BotaoAlterar_Click()

Dim lErro As Long
Dim objNode As Node
Dim sChave As String
Dim objProdutoKitNovo As New ClassProdutoKit, objProdutoKit As ClassProdutoKit

Dim iIndice As Integer 'Inserido por Wagner
Dim objProduto As New ClassProduto
Dim sProdutoMascarado As String

On Error GoTo Erro_BotaoAlterar_Click

    Set objNode = EstruturaProduto.SelectedItem

    If objNode Is Nothing Then Error 21727
    If objNode.Selected = False Then Error 19400

    'preenche objProdutoKit à partir dos dados do frame de "componentes"
    lErro = Extrai_ProdutoKit(objProdutoKitNovo)
    If lErro <> SUCESSO Then Error 19401

    sChave = objNode.Tag
    Set objProdutoKit = colComponentes.Item(sChave)
    
    'preenche a seq pai
    objProdutoKitNovo.iSeqPai = objProdutoKit.iSeqPai
    objProdutoKitNovo.sProdutoRaiz = objProdutoKit.sProdutoRaiz
    
    lErro = CF2(Me, "KT_BotaoAlterar_Click", objProdutoKitNovo, objProdutoKit)
    If lErro <> SUCESSO Then Error 19401

    '#######################################################
    'ALTERADO POR WAGNER
'    'verificar se é o mesmo produto
    If objProdutoKit.sProduto <> objProdutoKitNovo.sProduto Then 'Error 19402

        If objProdutoKitNovo.iNivel = KIT_NIVEL_RAIZ Then
        
            For iIndice = 1 To EstruturaProduto.Nodes.Count
                colComponentes.Item(EstruturaProduto.Nodes.Item(iIndice).Tag).sProdutoRaiz = objProdutoKitNovo.sProduto
            Next
            
            objProdutoKitNovo.sProdutoRaiz = objProdutoKitNovo.sProduto
            
            lErro = CF("Traz_Produto_MaskEd", objProdutoKitNovo.sProduto, Produto, Descricao)
            If lErro <> SUCESSO Then Error 19401
            
            sProdutoAnt = objProdutoKitNovo.sProduto
        
        End If
        
        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(objProdutoKitNovo.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 19401

        objProduto.sCodigo = objProdutoKit.sProduto

        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 19401

        If lErro = 28030 Then Error 19401
        
        objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao & SEPARADOR & " " & Format(objProdutoKitNovo.dQuantidade * 100, FORMATO_ESTOQUE_KIT) & "%"
                
    End If
    '#######################################################

    colComponentes.Remove (sChave)
    colComponentes.Add objProdutoKitNovo, sChave

    Call Recalcula_Nivel_Sequencial

    'Limpa a parte de Componente da Tela Kit
    lErro = Limpa_Componente_Kit()
    If lErro <> SUCESSO Then Error 55196

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoAlterar_Click:

    Select Case Err

        Case 19400, 21727
            lErro = Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_PRODUTO", Err)

        Case 19401, 55196

        Case 19402
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CONFERE_SEL", Err)
            ProdutoSel.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162077)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub CampoSel_Limpa()
    
    PercentualPerda.Text = ""
    DescProdutoSel.Caption = ""
    Quantidade.Text = "1"
    CustoStandard.Text = ""
    SiglaUM.Text = ""
    NomeUM.Caption = ""
    Composicao.ListIndex = -1
    Nivel.Caption = ""
    Sequencial.Caption = ""
    VersaoKitComp.Clear 'Inserido por Wagner

    Call CF2(Me, "KT_CampoSel_Limpa")
    
End Sub

Function Limpa_Tela_Kit() As Long
'Limpa a Tela Kit

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_Kit

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then Error 21801

    Call Limpa_Tela(Me)

    Descricao.Caption = ""
    DescProdutoSel.Caption = ""
    Nivel.Caption = ""
    Sequencial.Caption = ""

    Quantidade.Text = "1"
    CustoStandard.Text = ""
    
    NomeUM.Caption = ""
    SiglaUM.Clear
    Call Composicao_Seleciona
    iClasseAtual = 0

    EstruturaProduto.Nodes.Clear
    Set colComponentes = Nothing
    sProdutoSelAnterior = ""
    
    iProxChave = 1

    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    '*****************************************************
    'Alteracao Daniel em 10/07/2002
    Situacao.ListIndex = -1
    '*****************************************************
    
    '###############################
    'Inserido por Wagner
    VersaoKitComp.Clear
    VersaoKitComp.Visible = True
    LabelVersaoComp.Visible = True
    '###############################
    VersaoFormPreco.Value = vbUnchecked
    
    Call CF2(Me, "KT_Limpa_Tela_Kit")
    
    sProdutoAnt = ""
   
    iAlterado = 0

    Limpa_Tela_Kit = SUCESSO

    Exit Function

Erro_Limpa_Tela_Kit:

    Limpa_Tela_Kit = Err

    Select Case Err

        Case 21801

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162078)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 21750

    'Limpa Tela
    Call Limpa_Tela_Kit

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 21750

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162079)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objKit As New ClassKit
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

'****************************************
'Alteracao Daniel em 12/07/2002
Dim iSituacaoAlterada As Integer
'****************************************
Dim vbMsgResult As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de campos obrigatorios
    If Len(Trim(Produto.ClipText)) = 0 Then gError 21751
    If Len(Trim(Versao.Text)) = 0 Then gError 21753       'alterado por Jorge Specian (Trim)
    If Len(Trim(Data.ClipText)) = 0 Then gError 19367
    If EstruturaProduto.Nodes.Count < 2 Then gError 21758
    If Len(Trim(Situacao.Text)) = 0 Then gError 106240
    
    lErro = Move_Tela_Memoria(objKit)
    If lErro <> SUCESSO Then gError 19403

    If gobjEST.iAltDataKitNaGravacao <> DESMARCADO Then
        If objKit.dtData <> Date Then
            If gobjEST.iAltDataKitNaGravacao = 1 Then
                vbMsgResult = Rotina_Aviso(vbOKCancel, "AVISO_KIT_DATA_NAO_ATUAL")
            Else
                vbMsgResult = vbOK
            End If
            If vbMsgResult = vbOK Then objKit.dtData = Date
        End If
    End If
    
    lErro = CF2(Me, "KT_GravarRegistro", objKit)
    If lErro <> SUCESSO Then gError 19403
    
    lErro = Trata_Alteracao(objKit, objKit.sProdutoRaiz, objKit.sVersao)
    If lErro <> SUCESSO Then gError 32307
    
    '****************************************
    'Alteracao Daniel em 12/07/2002
    'Guarda a Situacao a ser gravada
    iSituacaoAlterada = objKit.iSituacao
    '****************************************
    
    '##########################################################
    'INCLUÍDO POR WAGNER
    'Teste que verifica se a quantidade do Raiz é maior que a soma dos produtos Kit
    lErro = CF("Kit_Valida_Quantidade", objKit)
    If lErro <> SUCESSO Then gError 131680
    
    'Verifica recursividade circular, ou seja, o produto final não pode estar contido como insumo de um dos componentes, direta ou indiretamente desde que e igual ou maior quantidade
    'ntre outros
    lErro = CF("Kit_Valida", objKit)
    If lErro <> SUCESSO Then gError 131680
    'FIM
    '##########################################################
    
    lErro = CF("Kit_Grava", objKit)
    If lErro <> SUCESSO And lErro <> 21779 And lErro <> 21780 Then gError 21759

    If (lErro = 21779 Or lErro = 21780) Then gError 21755

    '****************************************
    'Alteracao Daniel em 12/07/2002
    'Se a situacao foi alterada => avisa o usuario
    If iSituacaoAlterada <> objKit.iSituacao Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_KIT_SITUACAO_PADRAO")
    End If
    '****************************************

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 19367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 19403, 21752, 21755, 21756, 21759, 21760, 32307

        Case 21751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO", gErr)

        Case 21753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VERSAO_KIT_NAO_PREENCHIDA", gErr)

        Case 21758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_APENAS_COM_PRODUTORAIZ", gErr)
            
        Case 106240
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_SITUACAO", gErr)
            
        Case 131680 'Inserido por Wagner

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162080)

    End Select

    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objKit As New ClassKit
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Produto.ClipText)) = 0 Then Error 21730

    If Len(Versao.Text) = 0 Then Error 21735

    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 Then Error 21731

    If lErro = 25041 Then Error 21732

    objKit.sProdutoRaiz = objProduto.sCodigo
    objKit.sVersao = Trim(Versao.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_KIT", objKit.sProdutoRaiz, objKit.sVersao)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    lErro = CF("Kit_Exclui", objKit)
    If lErro <> SUCESSO And lErro <> 21745 And lErro <> 21741 Then Error 21733

    If lErro = 21741 Then Error 21749

    If lErro = 21745 Then Error 21750

    lErro = Limpa_Tela_Kit
    If lErro <> SUCESSO Then Error 21734

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 21730
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO", Err)

        Case 21731, 21734, 21733

        Case 21732
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, sProduto)

        Case 21735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VERSAO_KIT_NAO_PREENCHIDA", Err)

        Case 21749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_INEXISTENTE", Err, objKit.sVersao, objKit.sProdutoRaiz)

        Case 21750
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTOKIT_INEXISTENTE", Err, objKit.sVersao, objKit.sProdutoRaiz)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162081)

    End Select

    Exit Sub

End Sub

Function Testa_Problema_Pai_Na_Inclusao(objNodePai As Node) As Long
'verifica se o produto pai pode ter componentes e outros testes
'retorna SUCESSO se nao houver problema

Dim lErro As Long, iProdutoPreenchido As Integer, objProdutoKit As ClassProdutoKit
Dim objProduto As New ClassProduto

On Error GoTo Erro_Testa_Problema_Pai_Na_Inclusao

    objProduto.sCodigo = right(objNodePai.Key, Len(objNodePai.Key) - 6) 'Alterado por Wagner
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 19381
    
    If lErro = 28030 Then Error 19382

    Set objProdutoKit = colComponentes.Item(objNodePai.Tag)

    'ver se o pai não é o raiz e está marcado como NAO_KIT_INTERMEDIARIO
    If objProdutoKit.iNivel <> 0 And objProduto.iKitInt = NAO_KIT_INTERMEDIARIO Then Error 21788
    
    'Verifica se produto pai já está no último nível permitido
    If objProdutoKit.iNivel = NIVEL_MAXIMO_KIT Then Error 21792

    Testa_Problema_Pai_Na_Inclusao = SUCESSO

    Exit Function

Erro_Testa_Problema_Pai_Na_Inclusao:

    Testa_Problema_Pai_Na_Inclusao = Err

    Select Case Err

        Case 19381, 21789

        Case 21792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NIVEL_MAXIMO_KIT", Err, objProdutoKit.sProduto, objProdutoKit.iNivel)

        Case 19382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 21796
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_FAZ_PARTE_CAMINHO_KIT", Err)

        Case 21788
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_KIT_INTERMEDIARIO", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162082)

    End Select

    Exit Function

End Function

Function Testa_Componente_Recursivo(objNodePai As Node, sTexto As String) As Long
'verifica se de objNode para cima existe algum com sTexto como node.text
'retorna SUCESSO se nao houver

Dim objNode1 As Node, lErro As Long
Dim sProduto As String

On Error GoTo Erro_Testa_Componente_Recursivo

    'verifica se existem pais com o mesmo produto

    Set objNode1 = objNodePai

    Do

        sProduto = right(objNode1.Key, Len(objNode1.Key) - 6) 'Alterado por Wagner

        If sProduto = sTexto Then Error 21796

        Set objNode1 = objNode1.Parent

    Loop Until objNode1 Is Nothing

    'Verifica se existem irmãos com o mesmo produto

    Set objNode1 = objNodePai.Child

    Do While Not (objNode1 Is Nothing)

        sProduto = right(objNode1.Key, Len(objNode1.Key) - 6) 'Alterado por Wagner

        If sProduto = sTexto Then Error 41416

        Set objNode1 = objNode1.Next

    Loop

    Testa_Componente_Recursivo = SUCESSO

    Exit Function

Erro_Testa_Componente_Recursivo:

    Testa_Componente_Recursivo = Err

    Select Case Err

        Case 21796, 41416
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_FAZ_PARTE_CAMINHO_KIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162083)

    End Select

    Exit Function

End Function

Public Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim objRelatorio As New AdmRelatorio
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoImprimir_Click


    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(Produto.Text)) = 0 Then gError 71408
    
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 71412
    
    If Len(Trim(Versao.Text)) = 0 Then gError 71409
    If Len(Trim(Data.ClipText)) = 0 Then gError 71410
    
    objKit.sProdutoRaiz = sProdutoFormatado
    objKit.sVersao = Versao.Text
    objKit.dtData = StrParaDate(Data.Text)
    
'    lErro = CF("OrdemDeProducao_TestaExistencia",objOrdemDeProducao)
'    If lErro <> SUCESSO And lErro <> 57721 Then Error 57724
'    If lErro = 57721 Then Error 57725
    
    lErro = objRelatorio.ExecutarDireto("Relação de Kits", "ProdutoRaiz = @TPRODRAIZ E Versao = @TVERSAO E DataKit = @DDATAKIT", 1, "Kit", "TPRODRAIZ", objKit.sProdutoRaiz, "TVERSAO", objKit.sVersao, "DDATAKIT", objKit.dtData)
    If lErro <> SUCESSO Then gError 71411
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 71408
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 71409
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VERSAO_KIT_NAO_PREENCHIDA", gErr)
        
        Case 71410
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
    
'        Case 57724, 57726
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_INEXISTENTE", Err, objOrdemDeProducao.sCodigo)
            
        Case 71411, 71412
                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162084)

    End Select
    
    Exit Sub

End Sub

Public Sub BotaoIncluir_Click()

Dim lErro As Long, iIndice As Integer, iProdutoPreenchido As Integer
Dim sChave As String, sTexto As String
Dim sProduto As String, sProdutoFormatado As String, sProdutoMascarado As String
Dim objProduto As New ClassProduto
Dim objKit As New ClassKit, objProdutoKit As New ClassProdutoKit, objProdutoKit1 As New ClassProdutoKit
Dim objNode As Node, objNodePai As Node, sChaveTvw As String

On Error GoTo Erro_BotaoIncluir_Click

    If Len(Produto.ClipText) = 0 Then Error 21781

    Set objNodePai = EstruturaProduto.SelectedItem

    If objNodePai Is Nothing Then Error 19375
    If objNodePai.Selected = False Then Error 19376

    'verifica se o produto pai pode ter componentes e outros testes
    lErro = Testa_Problema_Pai_Na_Inclusao(objNodePai)
    If lErro <> SUCESSO Then Error 19377

    lErro = Extrai_ProdutoKit(objProdutoKit)
    If lErro <> SUCESSO Then Error 19378

    sProdutoMascarado = String(STRING_PRODUTO, 0)

    sProduto = objProdutoKit.sProduto

    'coloca o Produto no formato que é exibida na tela
    lErro = Mascara_MascararProduto(sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then Error 19379

    'prepara texto que identificará o novo componente que está sendo incluido
    sTexto = Trim(sProdutoMascarado) & SEPARADOR & objProdutoKit.sProdutoDesc

    '*****************************************
    'Alteracao Daniel em 12/07/2002
    
        'Agora um produto pode fazer parte da composicao dele mesmo e poderá aparecer
        'mais de uma vez no mesmo kit. Portanto o trecho que verifica a sua existencia no kit
        'foi comentado.
        
        'verifica se o produto já está incluído no caminho
        'lErro = Testa_Componente_Recursivo(objNodePai, objProdutoKit.sProduto)
        'If lErro <> SUCESSO Then Error 19380
        
    '*****************************************

    'prepara uma chave para relacionar colComponentes ao node que está sendo incluido
    Call Calcula_Proxima_Chave(sChaveTvw)
    
    sChave = sChaveTvw
    sChaveTvw = sChaveTvw & objProdutoKit.sProduto
    
    'inclui o node na treeview
    Set objNode = EstruturaProduto.Nodes.Add(objNodePai.Index, tvwChild, sChaveTvw, sTexto)

    colComponentes.Add objProdutoKit, sChave
    objNode.Tag = sChave

    Call Recalcula_Nivel_Sequencial

    'Limpa a parte de Componente da Tela Kit
    lErro = Limpa_Componente_Kit()
    If lErro <> SUCESSO Then Error 55195

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case Err

        Case 19375, 19376
            lErro = Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_PRODUTO", Err)

        Case 19377, 19378, 19380, 55195

        Case 19379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, sProduto)

        Case 21781
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162085)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 21797

    Call Limpa_Tela_Kit

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 21797

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162086)

    End Select

    Exit Sub

End Sub

Public Sub BotaoRemover_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objNode As Node
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoRemover_Click

    Set objNode = EstruturaProduto.SelectedItem

    If objNode Is Nothing Then Error 21799
    If objNode.Selected = False Then Error 19384

    'nao pode remover o raiz
    If EstruturaProduto.SelectedItem.Index = 1 Then Error 21800

    If objNode.Children > 0 Then

        'Envia aviso perguntando se realmente deseja excluir
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_TEM_FILHOS")

        If vbMsgRes = vbNo Then Exit Sub

        'chama rotina que exclui filhos
        Call Remove_Filhos(objNode.Child)

    End If

    colComponentes.Remove (objNode.Tag)
    EstruturaProduto.Nodes.Remove (objNode.Index)

    Call Recalcula_Nivel_Sequencial

    'Limpa a parte de Componente da Tela Kit
    lErro = Limpa_Componente_Kit()
    If lErro <> SUCESSO Then Error 55197

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoRemover_Click:

    Select Case Err

        Case 19384, 21799
            lErro = Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_PRODUTO", Err)

        Case 21800
            lErro = Rotina_Erro(vbOKOnly, "AVISO_TROCAR_PRODUTO", Err)

        Case 55197

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162087)

    End Select

    Exit Sub

End Sub

Public Sub BPF_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BPF_GotFocus()

    Call MaskEdBox_TrataGotFocus(BPF)

End Sub


Public Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data)

End Sub

Public Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 21802

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 21802

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162088)

    End Select

    Exit Sub

End Sub

Public Sub Energia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Energia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Energia)
    
End Sub

Public Sub EstruturaProduto_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim objProdutoKit As New ClassProdutoKit
Dim objNode As Node

On Error GoTo Erro_EstruturaProduto_NodeClick

    Set objNode = EstruturaProduto.SelectedItem

    Set objProdutoKit = colComponentes.Item(objNode.Tag)

    lErro = Preenche_ProdutoKit(objProdutoKit)
    If lErro <> SUCESSO Then Error 19375

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_EstruturaProduto_NodeClick:

    Select Case Err

        Case 19375

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162089)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long, objCamposGenericos As New ClassCamposGenericos

On Error GoTo Erro_Form_Load
    
    Set objEventoKit = New AdmEvento
    Set objEventoVersao = New AdmEvento
    Set objEventoVersaoKitComp = New AdmEvento 'Inserido por Wagner
    Set objEventoProduto = New AdmEvento
    Set objEventoComponente = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 21810

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoSel)
    If lErro <> SUCESSO Then Error 21811

'    'Inicializa a Lista de Produtos
'    lErro = CF("Carga_Arvore_Produto_Produziveis",TvwProduto.Nodes)
'    If lErro <> SUCESSO Then Error 21812
'
'    'Inicializa a Lista de Produtos
'    lErro = CF("Carga_Arvore_Produto_Componentes",TvwComponentes.Nodes)
'    If lErro <> SUCESSO Then Error 55201
    
    objCamposGenericos.lCodigo = CAMPOSGENERICOS_KIT_FATOR
    lErro = CF("CamposGenericosValores_Le_CodCampo", objCamposGenericos)
    If lErro <> SUCESSO Then Error 21811
    
    LabelFator(1) = objCamposGenericos.colCamposGenericosValores.Item(1).sValor & ":"
    LabelFator(2) = objCamposGenericos.colCamposGenericosValores.Item(2).sValor & ":"
    LabelFator(3) = objCamposGenericos.colCamposGenericosValores.Item(3).sValor & ":"
    LabelFator(4) = objCamposGenericos.colCamposGenericosValores.Item(4).sValor & ":"
    LabelFator(5) = objCamposGenericos.colCamposGenericosValores.Item(5).sValor & ":"
    LabelFator(6) = objCamposGenericos.colCamposGenericosValores.Item(6).sValor & ":"
    
    iClasseAtual = 0
    iAlterado = 0
    sProdutoSelAnterior = ""

    Quantidade.Format = FORMATO_ESTOQUE_KIT
    CustoStandard.Format = FORMATO_CUSTO
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

'    EstruturaProduto.ImageList = ImageList1

    '*****************************************************
    'Alteracao Daniel em 10/07/2002
    Call Preenche_Combo_Situacao
    '*****************************************************
    
    lErro = CF2(Me, "KT_Form_Load")
    If lErro <> SUCESSO Then Error 21811
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 21810, 21811, 21812, 55201

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162090)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoKit = Nothing
    Set objEventoVersao = Nothing
    Set objEventoVersaoKitComp = Nothing 'Inserido por Wagner
    Set objEventoProduto = Nothing
    Set objEventoComponente = Nothing
    
    Set colComponentes = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    If lErro <> SUCESSO Then Error 21813

    Exit Sub

Erro_Form_Unload:

    Select Case Err

        Case 21813

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162091)

    End Select

    Exit Sub

End Sub

Public Sub MaoDeObra_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub MaoDeObra_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaoDeObra)

End Sub

Private Sub objEventoVersao_evSelecao(obj1 As Object)

Dim objKit As ClassKit
Dim lErro As Long

On Error GoTo Erro_objEventoVersao_evSelecao

    Set objKit = obj1

    'Chama Traz_ClasseUM_Tela
    lErro = Traz_Tela_Produto_Kit(objKit)
    If lErro <> SUCESSO Then Error 52114
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoVersao_evSelecao:

    Select Case Err
    
        Case 52114

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162092)

    End Select

    Exit Sub
    
End Sub

Public Sub PercentualPerda_Change()
    
    iAlterado = REGISTRO_ALTERADO
        
End Sub

Public Sub PercentualPerda_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dAliquota As Double

On Error GoTo Erro_PercentualPerda_Validate

    'Verifica se esta preenchida
    If Len(Trim(PercentualPerda.Text)) = 0 Then Exit Sub

    'Critica
    lErro = Porcentagem_Critica(PercentualPerda.Text)
    If lErro <> SUCESSO Then Error 26898

    If CDbl(PercentualPerda.Text) = 100# Then Error 52286

    Exit Sub

Erro_PercentualPerda_Validate:

    Cancel = True


    Select Case Err

        Case 26898

        Case 52286
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_PERCENTUALPERDA_INVALIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162093)

    End Select

    Exit Sub

End Sub

Public Sub Produto_GotFocus()
    
'    'mostra a arvore de produtos
'    LabelProdutos.Visible = True
'    TvwProduto.Visible = True
'
'    'esconde a arvore dos produto que sao Kit
'    LabelComponentes.Visible = False
'    TvwComponentes.Visible = False
    
End Sub

Public Sub ProdutoLbl_Click()
'abrir browse de produtosKit

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoLbl_Click

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 41302

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProduto)
    
    Exit Sub

Erro_ProdutoLbl_Click:

    Select Case gErr

        Case 41302

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162094)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 83075

    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True
    
    Call Produto_Validate(bSGECancelDummy)
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 83075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162095)

    End Select

    Exit Sub

End Sub

Public Sub ComponenteLabel_Click()
'abrir browse de produtosKit

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ComponenteLabel_Click

    lErro = CF("Produto_Formata", ProdutoSel.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 83076

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutosKitLista", colSelecao, objProduto, objEventoComponente)
    
    Exit Sub

Erro_ComponenteLabel_Click:

    Select Case gErr

        Case 83076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, ProdutoSel.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162096)

    End Select

    Exit Sub

End Sub

Private Sub objEventoComponente_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_objEventoComponente_evSelecao

    Set objProduto = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 83077

    ProdutoSel.PromptInclude = False
    ProdutoSel.Text = sProduto
    ProdutoSel.PromptInclude = True
    
    Call ProdutoSel_Validate(bSGECancelDummy)
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoComponente_evSelecao:

    Select Case gErr

        Case 83077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162097)

    End Select

    Exit Sub

End Sub


Private Sub objEventoKit_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objKit As ClassKit

On Error GoTo Erro_objEventoKit_evSelecao

    Set objKit = obj1

    'Chama Traz_ClasseUM_Tela
    lErro = Traz_Tela_Produto_Kit(objKit)
    If lErro <> SUCESSO Then Error 21815

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoKit_evSelecao:

    Select Case Err

        Case 21815

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162098)

    End Select

    Exit Sub

End Sub

Public Sub LabelVersao_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_LabelVersao_Click

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 41302

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        objKit.sProdutoRaiz = sProdutoFormatado
        If Len(Trim(Versao.ClipText)) > 0 Then objKit.sVersao = Versao.Text
            
        colSelecao.Add sProdutoFormatado
        
        Call Chama_Tela("KitVersaoLista", colSelecao, objKit, objEventoVersao)
    
    Else
         gError 49970
         
    End If

    
    Exit Sub

Erro_LabelVersao_Click:

    Select Case gErr

        Case 41302
        
        Case 49970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO2", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162099)

    End Select

    Exit Sub

End Sub

Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim objKit As New ClassKit

On Error GoTo Erro_Produto_Validate
   
    If Len(Trim(Produto.ClipText)) = 0 Then
        Descricao.Caption = ""
        Exit Sub
    End If

    sProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then Error 19385

    'se o produto não estiver cadastrado ==> erro
    If lErro = 25041 Then Error 19386

    If gobjMAT.iAceitaKitProdGerencial <> MARCADO Then
        'se o produto for gerencial, não pode fazer parte de um kit
        If objProduto.iGerencial = GERENCIAL Then Error 19387
    End If
    
    If objProduto.sCodigo <> sProdutoAnt Then
    
        sProdutoAnt = objProduto.sCodigo

        Descricao.Caption = objProduto.sDescricao
        
        'Incluído por Wagner
        objKit.sProdutoRaiz = objProduto.sCodigo
        
        lErro = Traz_Tela_Produto_Kit(objKit)
        If lErro <> SUCESSO And lErro <> 21860 Then Error 32297
        'Fim alteração
    
        If colComponentes.Count > 1 Then
    
            If colComponentes.Item(1).sProduto <> objProduto.sCodigo Then
                
                Call Limpa_Arvore_Kit
                
                lErro = Cria_Raiz(objProduto)
                If lErro <> SUCESSO Then Error 19397
                
            End If
    
        Else
            
            lErro = Limpa_Arvore_Kit()
            If lErro <> SUCESSO Then Error 52124
                    
            lErro = Cria_Raiz(objProduto)
            If lErro <> SUCESSO Then Error 55202
    
        End If
        
    End If
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case Err

        Case 19385, 19397, 32297, 52124, 55202

        Case 19386
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 19387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162100)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoSel_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Function Preenche_Combo_UMs(iClasseUM As Integer) As Long

Dim lErro As Long, objClasse As New ClassClasseUM, colSiglas As New Collection
Dim objUnidadeDeMedida As ClassUnidadeDeMedida

On Error GoTo Erro_Preenche_Combo_UMs

    If iClasseAtual <> iClasseUM Then

        objClasse.iClasse = iClasseUM

        lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasse, colSiglas)
        If lErro <> SUCESSO And lErro <> 22539 Then Error 19389

        If lErro = 22539 Then Error 19390

        SiglaUM.Clear
        NomeUM.Caption = ""

        For Each objUnidadeDeMedida In colSiglas

            SiglaUM.AddItem objUnidadeDeMedida.sSigla

        Next

        iClasseAtual = iClasseUM

    End If

    Preenche_Combo_UMs = SUCESSO

    Exit Function

Erro_Preenche_Combo_UMs:

    Preenche_Combo_UMs = Err

    Select Case Err

        Case 19389

        Case 19390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_INEXISTENTE", Err, objClasse.iClasse)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162101)

    End Select

    Exit Function

End Function

Public Sub Composicao_Seleciona()
Dim iIndice As Integer

    Composicao.ListIndex = -1
    For iIndice = 0 To Composicao.ListCount - 1
        If Composicao.List(iIndice) = COMPOSICAO_VARIAVEL Then
            Composicao.ListIndex = iIndice
            Exit For
        End If
    Next

End Sub

Public Sub ProdutoSel_GotFocus()
            
    'mostra a arvore dos produto que são Kit
'    TvwComponentes.Visible = True
'    LabelComponentes.Visible = True
'
'    'e esconde a arvore de todos os produtos
'    TvwProduto.Visible = False
'    LabelProdutos.Visible = False
    
End Sub

Public Sub ProdutoSel_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objUnidadeDeMedida As New ClassUnidadeDeMedida
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim sProduto As String, iIndice As Integer
Dim objKit As New ClassProdutoKit
Dim sProdutoFormatado As String

On Error GoTo Erro_ProdutoSel_Validate

    'Verifica preenchimento de Produto Selecionado
    If Len(Trim(ProdutoSel.ClipText)) > 0 Then

        If sProdutoSelAnterior <> ProdutoSel.Text Then

            sProduto = ProdutoSel.Text

            'Critica o formato do Produto e se existe no BD
            lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 21838

            If lErro = 25041 Then gError 21839

            If gobjMAT.iAceitaKitProdGerencial <> MARCADO Then
                If objProduto.iGerencial = GERENCIAL Then gError 21844
            End If

            'Preenche Descricao com Descrição do Produto
            DescProdutoSel.Caption = objProduto.sDescricao
                    
            lErro = Preenche_Combo_UMs(objProduto.iClasseUM)

            'Mostra Unidade de Medida na tela
            SiglaUM.Text = objProduto.sSiglaUMEstoque

            objUnidadeDeMedida.sSigla = objProduto.sSiglaUMEstoque
            objUnidadeDeMedida.iClasse = objProduto.iClasseUM

            lErro = CF("UM_Le", objUnidadeDeMedida)
            If lErro <> SUCESSO And lErro <> 23775 Then gError 21840

            If lErro = 23775 Then gError 21841

            NomeUM.Caption = objUnidadeDeMedida.sNome

            '######################################################
            'Inserido por Wagner
            Call Carrega_ComboVersoes(objProduto.sCodigo)
            
            lErro = VersaoKitComp_SelecionaPadrao(objProduto.sCodigo)
            If lErro <> SUCESSO And lErro <> 131307 Then gError 131308
            '######################################################

            Call Composicao_Seleciona

        End If

        sProdutoSelAnterior = ProdutoSel.Text

    Else

        Call CampoSel_Limpa

    End If

    Exit Sub

Erro_ProdutoSel_Validate:

    Cancel = True

    Select Case gErr

        Case 21838, 21840, 55193, 131303, 131308 'Alterado por Wagner

        Case 21839
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 21841
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_SIGLAUM_INEXISTENTE", gErr, objUnidadeDeMedida.iClasse, objUnidadeDeMedida.sSigla)

        Case 21844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162102)

    End Select

    Exit Sub

End Sub

Public Sub CustoStandard_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub CustoStandard_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_CustoStandard_Validate

    If Len(Trim(CustoStandard.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(CustoStandard.Text)
        If lErro <> SUCESSO Then Error 52881
        
    End If
    
    Exit Sub
    
Erro_CustoStandard_Validate:

    Cancel = True


    Select Case Err
        
        Case 52881
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162103)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sValor As String

On Error GoTo Erro_Quantidade_Validate

    sValor = Quantidade.Text

    If Len(Trim(sValor)) > 0 Then

        lErro = Valor_Positivo_Critica_Double(sValor)
        If lErro <> SUCESSO Then Error 41418

        Quantidade.Text = Format(sValor, FORMATO_ESTOQUE_KIT)

    End If

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True


    Select Case Err

        Case 41418

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162104)

    End Select

    Exit Sub

End Sub

Public Sub SiglaUM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub SiglaUM_Click()

Dim lErro As Long
Dim objUnidadeDeMedida As New ClassUnidadeDeMedida

On Error GoTo Erro_SiglaUM_Click

    If SiglaUM.ListIndex <> -1 Then

        objUnidadeDeMedida.sSigla = SiglaUM.Text
        objUnidadeDeMedida.iClasse = iClasseAtual

        lErro = CF("UM_Le", objUnidadeDeMedida)
        If lErro <> SUCESSO And lErro <> 23775 Then Error 21845

        If lErro = 23775 Then Error 21846

        NomeUM.Caption = objUnidadeDeMedida.sNome

    Else

        NomeUM.Caption = ""

    End If

    Exit Sub

Erro_SiglaUM_Click:

    Select Case Err

        Case 21845

        Case 21846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_SIGLAUM_INEXISTENTE", Err, objUnidadeDeMedida.iClasse, objUnidadeDeMedida.sSigla)
            SiglaUM.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162105)

    End Select

    Exit Sub

End Sub

Public Sub SiglaUM_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SiglaUM_Validate

    If Len(SiglaUM.Text) > 0 Then

        lErro = Combo_Item_Igual(SiglaUM)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 19391
        
        If lErro = 12253 Then Error 19392

    Else

        NomeUM.Caption = ""

    End If

    Exit Sub

Erro_SiglaUM_Validate:

    Cancel = True


    Select Case Err

        Case 19391

        Case 19392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_SIGLAUM_INEXISTENTE", Err, iClasseAtual, SiglaUM.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162106)

    End Select

    Exit Sub

End Sub

'Private Sub TvwComponentes_Expand(ByVal Node As MSComctlLib.Node)
'
'Dim lErro As Long
'
'On Error GoTo Erro_TvwComponentes_Expand
'
'    If Node.Tag <> NETOS_NA_ARVORE Then
'
'        lErro = CF("Carga_Arvore_Produto_Netos_Componentes",Node, TvwComponentes.Nodes)
'        If lErro <> SUCESSO Then Error 49969
'
'    End If
'
'    Exit Sub
'
'Erro_TvwComponentes_Expand:
'
'    Select Case Err
'
'        Case 49969
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162107)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub TvwComponentes_NodeClick(ByVal objNode As MSComctlLib.Node)
'
'Dim lErro As Long
'Dim sProduto As String
'
'On Error GoTo Erro_TvwComponentes_NodeClick
'
'    'obtem produto à partir do node
'    sProduto = Right(objNode.Key, Len(objNode.Key) - 1)
'
'    'Verifica se produto tem filhos
'    If objNode.Children > 0 Then Error 19393
'
'    ProdutoSel.PromptInclude = False
'    ProdutoSel.Text = sProduto
'    ProdutoSel.PromptInclude = True
'
'    Call ProdutoSel_Validate(bSGECancelDummy)
'
'    'Fecha comando de setas se estiver aberto
'    lErro = ComandoSeta_Fechar(Me.Name)
'
'    Exit Sub
'
'Erro_TvwComponentes_NodeClick:
'
'    Select Case Err
'
'        Case 19393
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, sProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162108)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub TvwProduto_Expand(ByVal objNode As MSComctlLib.Node)
'
'Dim lErro As Long
'
'On Error GoTo Erro_TvwProduto_Expand
'
'    If objNode.Tag <> NETOS_NA_ARVORE Then
'
'        'move os dados do plano de contas do banco de dados para a arvore colNodes.
'        lErro = CF("Carga_Arvore_Produto_Netos",objNode, TvwProduto.Nodes)
'        If lErro <> SUCESSO Then Error 48085
'
'    End If
'
'    Exit Sub
'
'Erro_TvwProduto_Expand:
'
'    Select Case Err
'
'        Case 48085
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162109)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub TvwProduto_NodeClick(ByVal Node As MSComctlLib.Node)
'
'Dim lErro As Long
'Dim sProduto As String
'Dim objProduto As New ClassProduto
'Dim objControle As Control
'
'On Error GoTo Erro_TvwProduto_NodeClick
'
'    'obtem produto à partir do node
'    sProduto = Right(Node.Key, Len(Node.Key) - 1)
'
'    'Verifica se produto tem filhos
'    If Node.Children > 0 Then Error 19393
'
'    Produto.PromptInclude = False
'    Produto.Text = sProduto
'    Produto.PromptInclude = True
'
'    Call Produto_Validate(bSGECancelDummy)
'
'    'Fecha comando de setas se estiver aberto
'    lErro = ComandoSeta_Fechar(Me.Name)
'
'    Exit Sub
'
'Erro_TvwProduto_NodeClick:
'
'    Select Case Err
'
'        Case 19393
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, objProduto.sCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162110)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Function Trata_Parametros(Optional objKit As ClassKit) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objKit Is Nothing) Then

        lErro = Traz_Tela_Produto_Kit(objKit)
        If lErro <> SUCESSO And lErro <> 21861 Then Error 21854

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 21854

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162111)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Carrega_Arvore(objKit As ClassKit) As Long
'preenche a treeview EstruturaProduto com a composicao de objKit

Dim objNode As Node
Dim objProduto As New ClassProduto
Dim lErro As Long, sChave As String
Dim iProdutoPreenchido As Integer
Dim sProdutoMascarado As String
Dim sProduto As String, sTexto As String
Dim objProdutoKit As ClassProdutoKit

On Error GoTo Erro_Carrega_Arvore

    'leitura dos Kits dos Produtos no BD
    lErro = CF("Kit_Le_Componentes", objKit)
    If lErro <> SUCESSO And lErro <> 21831 Then Error 21833

    If lErro = 21831 Then Error 21834

    For Each objProdutoKit In objKit.colComponentes

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        sProduto = objProdutoKit.sProduto

        'coloca o Produto no formato que é exibida na tela
        lErro = Mascara_MascararProduto(sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then Error 21835

        objProduto.sCodigo = objProdutoKit.sProduto

        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 21836

        If lErro = 28030 Then Error 21837

        'Alterado por Wagner
        sTexto = sProdutoMascarado & SEPARADOR & objProduto.sDescricao & SEPARADOR & " " & Format(objProdutoKit.dQuantidade * 100, FORMATO_ESTOQUE_KIT) & "%"
        'Fim alteração
        
        If objProdutoKit.iNivel = 0 Then

            sChave = "X" & CStr(0) & String(2 - Len(CStr(0)), 0) & CStr(1) & String(3 - Len(CStr(1)), 0) & objProduto.sCodigo 'Alterado por Wagner

            Set objNode = EstruturaProduto.Nodes.Add(, tvwFirst, sChave, sTexto)
            
            'Incluído por Wagner
            EstruturaProduto.Nodes.Item(objNode.Index).Expanded = True
            'Fim da Alteração

        Else

            sChave = "X" & CStr(objProdutoKit.iNivel) & String(2 - Len(CStr(objProdutoKit.iNivel)), 0) & CStr(objProdutoKit.iSeq) & String(3 - Len(CStr(objProdutoKit.iSeq)), 0) & objProduto.sCodigo 'Alterado por Wagner

            Set objNode = EstruturaProduto.Nodes.Add(objProdutoKit.iNivel + objProdutoKit.iSeqPai - 1, tvwChild, sChave, sTexto)

        End If

        Call Calcula_Proxima_Chave(sChave)
        
        colComponentes.Add objProdutoKit, sChave
        
        objNode.Tag = sChave
        
    Next
    
    Carrega_Arvore = SUCESSO

    Exit Function

Erro_Carrega_Arvore:

    Carrega_Arvore = Err

    Select Case Err

        Case 21833, 21834, 21836

        Case 21835
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, sProduto)

        Case 21837
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162112)

    End Select

    Exit Function

End Function

Sub EstruturaProduto_Limpa()

    Do While EstruturaProduto.Nodes.Count > 0
        EstruturaProduto.Nodes.Remove (1)
    Loop

End Sub

Function Traz_Tela_Produto_Kit(objKit As ClassKit) As Long
'traz para a tela dados da composicao de um produto
'se a versao nao estiver prenchida traz o padrao

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_Tela_Produto_Kit

    If objKit.sVersao = "" Then

        lErro = CF("Kit_Le_Padrao", objKit)
        If lErro <> SUCESSO And lErro <> 106304 Then Error 21859

        If lErro = 106304 Then Error 21860

    End If

    Versao.Text = objKit.sVersao

    lErro = CF("Traz_Produto_MaskEd", objKit.sProdutoRaiz, Produto, Descricao)
    If lErro <> SUCESSO Then Error 21861
    
    sProdutoAnt = objKit.sProdutoRaiz

    EstruturaProduto.Nodes.Clear
    
    Set colComponentes = Nothing
    iProxChave = 1
    sProdutoSelAnterior = ""

    If Len(Versao.Text) > 0 Then

        lErro = Carrega_Arvore(objKit)
        If lErro <> SUCESSO And lErro <> 21834 Then Error 21821

        If lErro <> 21834 Then

            'leitura dos kits na tabela de kits
            lErro = CF("Kit_Le", objKit)
            If lErro <> SUCESSO And lErro <> 21826 Then Error 21822

            If lErro = 21826 Then Error 21827

            Observacao.Text = objKit.sObservacao
            Data.Text = Format(objKit.dtData, "dd/mm/yy")
            
            '??? leo
            If objKit.dPesoFator1 <> 0 Then
                MaoDeObra.PromptInclude = False
                MaoDeObra.Text = Format(objKit.dPesoFator1, FORMATO_KIT_PESO_FATOR)
                MaoDeObra.PromptInclude = True
            Else
                MaoDeObra.PromptInclude = False
                MaoDeObra.Text = ""
                MaoDeObra.PromptInclude = True
            End If
                
            If objKit.dPesoFator2 <> 0 Then
                BPF.PromptInclude = False
                BPF.Text = Format(objKit.dPesoFator2, FORMATO_KIT_PESO_FATOR)
                BPF.PromptInclude = True
            Else
                BPF.PromptInclude = False
                BPF.Text = ""
                BPF.PromptInclude = True
            End If
                
            If objKit.dPesoFator3 <> 0 Then
                Energia.PromptInclude = False
                Energia.Text = Format(objKit.dPesoFator3, FORMATO_KIT_PESO_FATOR)
                Energia.PromptInclude = True
            Else
                Energia.PromptInclude = False
                Energia.Text = ""
                Energia.PromptInclude = True
            End If
                
            If objKit.dPesoFator4 <> 0 Then
                Agua.PromptInclude = False
                Agua.Text = Format(objKit.dPesoFator4, FORMATO_KIT_PESO_FATOR)
                Agua.PromptInclude = True
            Else
                Agua.PromptInclude = False
                Agua.Text = ""
                Agua.PromptInclude = True
            End If
        
            If objKit.dPesoFator5 <> 0 Then
                PesoFator5.PromptInclude = False
                PesoFator5.Text = Format(objKit.dPesoFator5, FORMATO_KIT_PESO_FATOR)
                PesoFator5.PromptInclude = True
            Else
                PesoFator5.PromptInclude = False
                PesoFator5.Text = ""
                PesoFator5.PromptInclude = True
            End If
        
            If objKit.dPesoFator6 <> 0 Then
                PesoFator6.PromptInclude = False
                PesoFator6.Text = Format(objKit.dPesoFator6, FORMATO_KIT_PESO_FATOR)
                PesoFator6.PromptInclude = True
            Else
                PesoFator6.PromptInclude = False
                PesoFator6.Text = ""
                PesoFator6.PromptInclude = True
            End If
        
            '****************************************************
            'Alteracao Daniel em 11/07/2002
            'Seleciona na combo a situacao
            For iIndice = 0 To Situacao.ListCount - 1
            
                If Situacao.ItemData(iIndice) = objKit.iSituacao Then
                    
                    Situacao.ListIndex = iIndice
                    Exit For
                
                End If
            
            Next
            '****************************************************

            VersaoFormPreco.Value = objKit.iVersaoFormPreco
            
        End If

        'selecionar a raiz
        Set EstruturaProduto.SelectedItem = EstruturaProduto.Nodes.Item(1)
        EstruturaProduto.SelectedItem.Selected = True
        
    End If

    iAlterado = 0

    Traz_Tela_Produto_Kit = SUCESSO

    Exit Function

Erro_Traz_Tela_Produto_Kit:

    Traz_Tela_Produto_Kit = Err

    Select Case Err

        Case 21822, 21859, 21860, 21861

        Case 21827
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_INEXISTENTE", Err, objKit.sVersao, objKit.sProdutoRaiz)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162113)

    End Select

    Exit Function

End Function

'********************************
'Alteracao Daniel em 17/07/2002
Public Sub Situacao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Situacao_Click()

    iAlterado = REGISTRO_ALTERADO
    If Situacao.ListIndex >= 0 Then
        If Situacao.ItemData(Situacao.ListIndex) = KIT_SITUACAO_PADRAO Then VersaoFormPreco.Value = vbChecked
    End If

End Sub
'********************************

Public Sub Versao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 21862

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 21862

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162114)

    End Select

    Exit Sub

End Sub

Public Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 21863

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 21863

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162115)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objKit As New ClassKit

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Kit"

    'Lê os dados da Tela Kit
    lErro = Move_Tela_Memoria(objKit, True)
    If lErro <> SUCESSO Then Error 21864

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "ProdutoRaiz", objKit.sProdutoRaiz, STRING_PRODUTO, "ProdutoRaiz"
    colCampoValor.Add "Versao", objKit.sVersao, STRING_KIT_VERSAO, "Versao"
    colCampoValor.Add "Data", objKit.dtData, 0, "Data"
    colCampoValor.Add "Observacao", objKit.sObservacao, STRING_KIT_OBSERVACAO, "Observacao"
    colCampoValor.Add "Situacao", objKit.iSituacao, 0, "Situacao"
    colCampoValor.Add "VersaoFormPreco", objKit.iVersaoFormPreco, 0, "VersaoFormPreco"

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 21864

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162116)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objKit As ClassKit, Optional bSemMsgErro As Boolean = False) As Long

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto, objProdutoKit As ClassProdutoKit

On Error GoTo Erro_Move_Tela_Memoria
    
    If colComponentes.Count > 0 Then
        
        sProduto = String(STRING_PRODUTO, 0)

        lErro = Mascara_RetornaProdutoTela(colComponentes.Item(1).sProduto, sProduto)
        If lErro <> SUCESSO Then Error 55973
        
        objKit.sProdutoRaiz = colComponentes.Item(1).sProdutoRaiz
    Else
        sProduto = ""
    End If
    
    If Not bSemMsgErro Then
    
        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then Error 21865
    
        If lErro = 25041 Then Error 21866
        
    End If

    objKit.sVersao = Versao.Text
    

    If Len(Data.ClipText) > 0 Then objKit.dtData = CDate(Data.Text)

    objKit.sObservacao = Observacao.Text

'??? por leo
    objKit.dPesoFator1 = StrParaDbl(MaoDeObra.Text)
    objKit.dPesoFator2 = StrParaDbl(BPF.Text)
    objKit.dPesoFator3 = StrParaDbl(Energia.Text)
    objKit.dPesoFator4 = StrParaDbl(Agua.Text)
    objKit.dPesoFator5 = StrParaDbl(PesoFator5.Text)
    objKit.dPesoFator6 = StrParaDbl(PesoFator6.Text)

    Set objKit.colComponentes = colComponentes

    'atualizar a versao na colecao
    For Each objProdutoKit In colComponentes

        objProdutoKit.sVersao = objKit.sVersao

    Next
    
    '*****************************************************
    'Alteracao Daniel em 10/07/2002
    If Situacao.ListIndex >= 0 Then
        objKit.iSituacao = Situacao.ItemData(Situacao.ListIndex)
    End If
    '*****************************************************

    objKit.iVersaoFormPreco = VersaoFormPreco.Value
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 21865

        Case 21866
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 55973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, colComponentes.Item(1).sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162117)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objKit As New ClassKit

On Error GoTo Erro_Tela_Preenche

    objKit.sProdutoRaiz = colCampoValor.Item("ProdutoRaiz").vValor
    objKit.sVersao = colCampoValor.Item("Versao").vValor
    objKit.sObservacao = colCampoValor.Item("Observacao").vValor
    objKit.dtData = colCampoValor.Item("Data").vValor
    objKit.iSituacao = colCampoValor.Item("Situacao").vValor
    objKit.iVersaoFormPreco = colCampoValor.Item("VersaoFormPreco").vValor

    'Traz dados do Produto_Kit para a Tela
    lErro = Traz_Tela_Produto_Kit(objKit)
    If lErro <> SUCESSO And lErro <> 21861 Then Error 21867

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 21867

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162118)

    End Select

    Exit Sub

End Sub

Public Sub BotaoKits_Click()
'chamar browse de versoes do produto marcado

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoKits_Click

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 41302

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objKit.sProdutoRaiz = sProdutoFormatado
    objKit.sVersao = Versao.Text
    
    'Chama tela
    Call Chama_Tela("KitLista", colSelecao, objKit, objEventoKit)

    Exit Sub

Erro_BotaoKits_Click:

    Select Case Err
    
        Case 41302

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162119)

    End Select

    Exit Sub

End Sub

Function Cria_Raiz(objProduto As ClassProduto) As Long
'força a inclusao do proprio produto como o 1o elemento de uma nova "composicao"

Dim lErro As Long, objProdutoKit As New ClassProdutoKit

On Error GoTo Erro_Cria_Raiz

    objProdutoKit.dQuantidade = 1
    objProdutoKit.sProduto = objProduto.sCodigo
    objProdutoKit.iNivel = 0
    objProdutoKit.iSeq = 1
    objProdutoKit.iSeqPai = 0
    objProdutoKit.sProdutoRaiz = objProduto.sCodigo
    objProdutoKit.sUnidadeMed = objProduto.sSiglaUMEstoque
    objProdutoKit.iPosicaoArvore = 0
    
    objProdutoKit.iClasseUM = objProduto.iClasseUM
    objProdutoKit.sProdutoDesc = objProduto.sDescricao

    lErro = CF2(Me, "KT_Cria_Raiz", objProduto, objProdutoKit)
    If lErro <> SUCESSO Then Error 19396
    
    lErro = Insere_Raiz(objProdutoKit)
    If lErro <> SUCESSO Then Error 19396

    Cria_Raiz = SUCESSO

    Exit Function

Erro_Cria_Raiz:

    Cria_Raiz = Err

    Select Case Err

        Case 19396

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162120)

    End Select

    Exit Function

End Function

Function Insere_Raiz(objProdutoKit As ClassProdutoKit) As Long

Dim lErro As Long, sChave As String, sTexto As String, sProdutoMascarado As String, sProduto As String
Dim objNode As Node

On Error GoTo Erro_Insere_Raiz

    sProdutoMascarado = String(STRING_PRODUTO, 0)
    sProduto = objProdutoKit.sProduto

    'coloca o Produto no formato que é exibida na tela
    lErro = Mascara_MascararProduto(sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then Error 19395

    sChave = "X" & CStr(0) & String(2 - Len(CStr(0)), 0) & CStr(1) & String(3 - Len(CStr(1)), 0) & objProdutoKit.sProduto 'Alterado por Wagner
    sTexto = sProdutoMascarado & SEPARADOR & objProdutoKit.sProdutoDesc

    Set objNode = EstruturaProduto.Nodes.Add(, tvwFirst, sChave, sTexto)

    Call Calcula_Proxima_Chave(sChave)
    objNode.Tag = sChave
    colComponentes.Add objProdutoKit, sChave

    Set EstruturaProduto.SelectedItem = objNode

    Insere_Raiz = SUCESSO

    Exit Function

Erro_Insere_Raiz:

    Insere_Raiz = Err

    Select Case Err

        Case 19395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162121)

    End Select

    Exit Function

End Function

Function Limpa_Arvore_Kit() As Long
'Limpa a Arvore do Kit

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Arvore_Kit

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then Error 55194

    Nivel.Caption = ""
    Sequencial.Caption = ""

    EstruturaProduto.Nodes.Clear
    Set colComponentes = Nothing
    sProdutoSelAnterior = ""
    
    iProxChave = 1

    Limpa_Arvore_Kit = SUCESSO

    Exit Function

Erro_Limpa_Arvore_Kit:

    Limpa_Arvore_Kit = Err
    
    Select Case Err

        Case 55194

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162122)

    End Select

    Exit Function

End Function

Function Limpa_Componente_Kit() As Long
'Limpa a parte de Componente da Tela Kit

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Componente_Kit

    DescProdutoSel.Caption = ""
    ProdutoSel.PromptInclude = False
    ProdutoSel.Text = ""
    ProdutoSel.PromptInclude = True
    Nivel.Caption = ""
    Sequencial.Caption = ""
    NomeUM.Caption = ""
    SiglaUM.Clear
    Quantidade.Text = ""
    CustoStandard.Text = ""
    iClasseAtual = 0
    sProdutoSelAnterior = ""
    Composicao.ListIndex = -1
    PercentualPerda.Text = ""
    
    '##############################
    'Inserido por Wagner
    VersaoKitComp.Clear
    VersaoKitComp.Visible = True
    LabelVersaoComp.Visible = True
    '##############################
    
    Call CF2(Me, "KT_Limpa_Componente_Kit")
    
    Limpa_Componente_Kit = SUCESSO

    Exit Function

Erro_Limpa_Componente_Kit:

    Limpa_Componente_Kit = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162123)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_KIT
    Set Form_Load_Ocx = Me
    Caption = "Kit"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Kit"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call ProdutoLbl_Click
        ElseIf Me.ActiveControl Is Versao Then
            Call LabelVersao_Click
        ElseIf Me.ActiveControl Is ProdutoSel Then
            Call ComponenteLabel_Click
        End If
    End If

End Sub





























'Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label2, Source, X, Y)
'End Sub
'
'Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
'End Sub







'Private Sub LabelComponentes_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelComponentes, Source, X, Y)
'End Sub
'
'Private Sub LabelComponentes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelComponentes, Button, Shift, X, Y)
'End Sub
'
'Private Sub LabelProdutos_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelProdutos, Source, X, Y)
'End Sub
'
'Private Sub LabelProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelProdutos, Button, Shift, X, Y)
'End Sub

'*****************************************************
'Alteracao Daniel em 10/07/2002
Private Sub Preenche_Combo_Situacao()
'Preencha a Combo Situacao com: Padrão, Ativo e Inativo (guardando no item data o valor correspondente)

    Situacao.AddItem (KIT_SITUACAO_STRING_PADRAO)
    Situacao.ItemData(Situacao.NewIndex) = KIT_SITUACAO_PADRAO
    
    Situacao.AddItem (KIT_SITUACAO_STRING_ATIVO)
    Situacao.ItemData(Situacao.NewIndex) = KIT_SITUACAO_ATIVO
    
    Situacao.AddItem (KIT_SITUACAO_STRING_INATIVO)
    Situacao.ItemData(Situacao.NewIndex) = KIT_SITUACAO_INATIVO

End Sub
'*****************************************************

'######################################################
'Inserido por Wagner - 23/11/04
Private Sub Carrega_ComboVersoes(ByVal sProdutoRaiz As String)
    
Dim lErro As Long
Dim objKit As New ClassKit
Dim colKits As New Collection
    
On Error GoTo Erro_Carrega_ComboVersoes
    
    VersaoKitComp.Enabled = True
    
    'Limpa a Combo
    VersaoKitComp.Clear
    
    'Armazena o Produto Raiz do kit
    objKit.sProdutoRaiz = sProdutoRaiz
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
    If lErro <> SUCESSO And lErro <> 106333 Then gError 131302
    
    VersaoKitComp.AddItem ""
    
    'Carrega a Combo com os Dados da Colecao
    For Each objKit In colKits
    
        VersaoKitComp.AddItem (objKit.sVersao)
               
    Next
    
    Exit Sub
    
Erro_Carrega_ComboVersoes:

    Select Case gErr
    
        Case 131302
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162124)
    
    End Select
    
End Sub

Private Sub objEventoVersaoKitComp_evSelecao(obj1 As Object)

Dim objKit As ClassKit
Dim lErro As Long

On Error GoTo Erro_objEventoVersaoKitComp_evSelecao

    Set objKit = obj1

    Call Carrega_ComboVersoes(objKit.sProdutoRaiz)

    Call VersaoKitComp_Seleciona(objKit.sVersao)
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoVersaoKitComp_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162125)

    End Select

    Exit Sub
    
End Sub

Public Sub LabelVersaoComp_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_LabelVersaoComp_Click

    lErro = CF("Produto_Formata", ProdutoSel.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 131304

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        objKit.sProdutoRaiz = sProdutoFormatado
        If Len(Trim(VersaoKitComp.Text)) > 0 Then objKit.sVersao = VersaoKitComp.Text
            
        colSelecao.Add sProdutoFormatado
        
        Call Chama_Tela("KitVersaoLista", colSelecao, objKit, objEventoVersaoKitComp)
    
    Else
         gError 131305
         
    End If
    
    Exit Sub

Erro_LabelVersaoComp_Click:

    Select Case gErr

        Case 131304
        
        Case 131305
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162126)

    End Select

    Exit Sub

End Sub

Public Sub VersaoKitComp_Seleciona(sVersao As String)
Dim iIndice As Integer

    VersaoKitComp.ListIndex = -1
    For iIndice = 0 To VersaoKitComp.ListCount - 1
        If VersaoKitComp.List(iIndice) = sVersao Then
            VersaoKitComp.ListIndex = iIndice
            Exit For
        End If
    Next

End Sub

Private Function VersaoKitComp_SelecionaPadrao(sProduto As String)

Dim lErro As Long
Dim objKit As New ClassKit
    
On Error GoTo Erro_VersaoKitComp_SelecionaPadrao
    
    'Armazena o Produto Raiz do kit
    objKit.sProdutoRaiz = sProduto
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Padrao", objKit)
    If lErro <> SUCESSO And lErro <> 106304 Then gError 131304
    
    If lErro <> SUCESSO Then gError 131307
    
    Call VersaoKitComp_Seleciona(objKit.sVersao)
    
    VersaoKitComp_SelecionaPadrao = SUCESSO
    
    Exit Function

Erro_VersaoKitComp_SelecionaPadrao:

    VersaoKitComp_SelecionaPadrao = gErr
    
    Select Case gErr

        Case 131304
        
        Case 131307
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162127)

    End Select

    Exit Function
    
End Function
'##########################################################
Public Sub VersaoFormPreco_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub


Public Sub BotaoRoteiros_Click()

Dim lErro As Long
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoRoteiros_Click

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141966

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""

    objRoteirosDeFabricacao.sProdutoRaiz = sProdutoFormatado
    
    If Len(Trim(Versao.Text)) <> 0 Then
        objRoteirosDeFabricacao.sVersao = Versao.Text
    End If

    Call Chama_Tela("RoteirosDeFabricacao", objRoteirosDeFabricacao)

    Exit Sub

Erro_BotaoRoteiros_Click:

    Select Case gErr

        Case 141966

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 141967)

    End Select

    Exit Sub

End Sub
Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

Public Function KT_Extrai_ProdutoKit(ByVal objCT As Object, ByVal objProdutoKit As ClassProdutoKit) As Long
    KT_Extrai_ProdutoKit = SUCESSO
End Function

Public Function KT_Preenche_ProdutoKit(ByVal objCT As Object, ByVal objProdutoKit As ClassProdutoKit) As Long
    KT_Preenche_ProdutoKit = SUCESSO
End Function

Public Function KT_CampoSel_Limpa(ByVal objCT As Object) As Long
    KT_CampoSel_Limpa = SUCESSO
End Function

Public Function KT_Limpa_Tela_Kit(ByVal objCT As Object) As Long
    KT_Limpa_Tela_Kit = SUCESSO
End Function

Public Function KT_BotaoAlterar_Click(ByVal objCT As Object, ByVal objProdutoKitNovo As ClassProdutoKit, ByVal objProdutoKit As ClassProdutoKit) As Long
    KT_BotaoAlterar_Click = SUCESSO
End Function

Public Function KT_GravarRegistro(ByVal objCT As Object, ByVal objKit As ClassKit) As Long
    KT_GravarRegistro = SUCESSO
End Function

Public Function KT_Form_Load(ByVal objCT As Object) As Long
    KT_Form_Load = SUCESSO
End Function

Public Function KT_Limpa_Componente_Kit(ByVal objCT As Object) As Long
    KT_Limpa_Componente_Kit = SUCESSO
End Function

Public Function KT_Cria_Raiz(ByVal objCT As Object, ByVal objProduto As ClassProduto, ByVal objProdutoKit As ClassProdutoKit) As Long
    KT_Cria_Raiz = SUCESSO
End Function

Public Sub Agua_Validate(Cancel As Boolean)
    
Dim lErro As Long, dPeso As Double
    
On Error GoTo Erro_Agua_Validate

    'Se estiver preenchido
    If Len(Trim(Agua.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Agua.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dPeso = CDbl(Agua.Text)

        'Coloca o valor formatado na tela
        Agua.Text = Format(dPeso, FORMATO_KIT_PESO_FATOR)
        
    End If

    Exit Sub
    
Erro_Agua_Validate:

    Cancel = True

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162103)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BPF_Validate(Cancel As Boolean)
    
Dim lErro As Long, dPeso As Double
    
On Error GoTo Erro_BPF_Validate

    'Se estiver preenchido
    If Len(Trim(BPF.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(BPF.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dPeso = CDbl(BPF.Text)

        'Coloca o valor formatado na tela
        BPF.Text = Format(dPeso, FORMATO_KIT_PESO_FATOR)
        
    End If

    Exit Sub
    
Erro_BPF_Validate:

    Cancel = True

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162103)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Energia_Validate(Cancel As Boolean)
    
Dim lErro As Long, dPeso As Double
    
On Error GoTo Erro_Energia_Validate

    'Se estiver preenchido
    If Len(Trim(Energia.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Energia.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dPeso = CDbl(Energia.Text)

        'Coloca o valor formatado na tela
        Energia.Text = Format(dPeso, FORMATO_KIT_PESO_FATOR)
        
    End If

    Exit Sub
    
Erro_Energia_Validate:

    Cancel = True

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162103)
    
    End Select
    
    Exit Sub

End Sub

Public Sub MaoDeObra_Validate(Cancel As Boolean)
    
Dim lErro As Long, dPeso As Double
    
On Error GoTo Erro_MaoDeObra_Validate

    'Se estiver preenchido
    If Len(Trim(MaoDeObra.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(MaoDeObra.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dPeso = CDbl(MaoDeObra.Text)

        'Coloca o valor formatado na tela
        MaoDeObra.Text = Format(dPeso, FORMATO_KIT_PESO_FATOR)
        
    End If

    Exit Sub
    
Erro_MaoDeObra_Validate:

    Cancel = True

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162103)
    
    End Select
    
    Exit Sub

End Sub

Public Sub PesoFator5_Validate(Cancel As Boolean)
    
Dim lErro As Long, dPeso As Double
    
On Error GoTo Erro_PesoFator5_Validate

    'Se estiver preenchido
    If Len(Trim(PesoFator5.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(PesoFator5.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dPeso = CDbl(PesoFator5.Text)

        'Coloca o valor formatado na tela
        PesoFator5.Text = Format(dPeso, FORMATO_KIT_PESO_FATOR)
        
    End If

    Exit Sub
    
Erro_PesoFator5_Validate:

    Cancel = True

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162103)
    
    End Select
    
    Exit Sub

End Sub

Public Sub PesoFator6_Validate(Cancel As Boolean)
    
Dim lErro As Long, dPeso As Double
    
On Error GoTo Erro_PesoFator6_Validate

    'Se estiver preenchido
    If Len(Trim(PesoFator6.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(PesoFator6.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dPeso = CDbl(PesoFator6.Text)

        'Coloca o valor formatado na tela
        PesoFator6.Text = Format(dPeso, FORMATO_KIT_PESO_FATOR)
        
    End If

    Exit Sub
    
Erro_PesoFator6_Validate:

    Cancel = True

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162103)
    
    End Select
    
    Exit Sub

End Sub

