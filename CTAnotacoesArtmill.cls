VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAnotacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'No trata_parametros, terá que receber objAnotacao com ID e Origem preenchidos senao deverá dar erro.
'Se o numintdoc <> 0 entao traze-lo para a tela, senao trazer o registro com a ID e Origem passados com o maior NumIntDoc.
'Guardar globalmente ao form a identificacao de quem é o objeto carregado na tela.
'Fazer browse e permitir o uso de setas para navegar pelas anotacoes referentes ao mesmo ID/Origem,
'guardando sempre globalmente a identificacao de quem é o objeto carregado na tela.
'Para cada registro na tabela OrigensAnotacoes, que será preenchida manualmente, deverá haver uma constante ANOTACAO_ORIGEM_*.

'Property Variables:
Dim m_Caption As String
Event Unload()

'controle de altaracao
Dim iAlterado As Integer

'guarda o Id de controle interno
Dim gsIdInterno As String
Dim giOrigem As Integer

'evento do browser
Private WithEvents objEventoAnotacao As AdmEvento
Attribute objEventoAnotacao.VB_VarHelpID = -1

'--- inicio dos properties get dos controles da tela

Public Property Get FrameID() As Object
     Set FrameID = objUserControl.Controls("FrameID")
End Property

Public Property Get IdOrigem() As Object
     Set IdOrigem = objUserControl.Controls("IdOrigem")
End Property

Public Property Get LabelID() As Object
     Set LabelID = objUserControl.Controls("LabelID")
End Property

Public Property Get Origem() As Object
     Set Origem = objUserControl.Controls("Origem")
End Property

Public Property Get LabelOrigem() As Object
     Set LabelOrigem = objUserControl.Controls("LabelOrigem")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get Anotacao() As Object
     Set Anotacao = objUserControl.Controls("Anotacao")
End Property

Public Property Get Titulo() As Object
     Set Titulo = objUserControl.Controls("Titulo")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

Public Property Get LabelData() As Object
     Set LabelData = objUserControl.Controls("LabelData")
End Property

Public Property Get LabelTexto() As Object
     Set LabelTexto = objUserControl.Controls("LabelTexto")
End Property

Public Property Get LabelTitulo() As Object
     Set LabelTitulo = objUserControl.Controls("LabelTitulo")
End Property

'--- fim dos properties get dos controles da tela

Public Function Trata_Parametros(ByVal objAnotacoes As ClassAnotacoes) As Long
'espera o Id e a Origem de ques está chamando a tela
    
Dim lErro As Long
    
On Error GoTo Erro_Trata_Parametros
        
    'se não foi setado o obj recebido por parametro, erro
    If objAnotacoes Is Nothing Then gError 119871
        
    'verifica se foi passado o Id de controle interno
    If objAnotacoes.sID = "" Then gError 119785
        
    'verifica se foi passado a origem
    If objAnotacoes.iTipoDocOrigem = 0 Then gError 119786
                            
    'verifica se o Id externo (oq vai ficar visivel) foi passado
'???    If objAnotacoes.sIdExterno = "" Then gError 119925
                            
    'carrega a variavel global c/ o ID de controle interno
    gsIdInterno = objAnotacoes.sID
    giOrigem = objAnotacoes.iTipoDocOrigem
                            
    'verifica todos os dados do obj recebido e traz a anotação referente p/ a tela
    lErro = Anotacoes_Valida_Preenchimento(objAnotacoes)
    If lErro <> SUCESSO Then gError 119872
                            
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 119872
    
        Case 119871
            Call Rotina_Erro(vbOKOnly, "ERRO_OBJANOTACOES_NAO_RECEBIDO", gErr)
    
        Case 119785
            Call Rotina_Erro(vbOKOnly, "ERRO_ID_NAO_PASSADO", gErr)
    
        Case 119786
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PASSADA", gErr)
        
        Case 119926 '??? cadastrar string
            Call Rotina_Erro(vbOKOnly, "ERRO_ID_EXTERNO_NAO_PASSADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PLEO_VB", gErr, Error)
    
    End Select
    
    Exit Function
    
End Function

Private Function Anotacoes_Valida_Preenchimento(ByVal objAnotacoes As ClassAnotacoes) As Long
'valida e traz os dados p/ a tela de acordo c/ o obj recebido

Dim lErro As Long

On Error GoTo Erro_Anotacoes_Valida_Preenchimento

    'verifica se o Id passado existe na tabela referente a origem
    'e preenche o frame de ID/origem c/ o id/origem passados
    lErro = Anotacoes_Preenche_Origem(objAnotacoes)
    If lErro <> SUCESSO Then gError 119789
    
    'se foi passado o titulo e a data (só pelo botao edita do browser)
    If objAnotacoes.sTitulo <> "" And objAnotacoes.dtDataAlteracao <> DATA_NULA Then
    
        'verifica se ainda existe o registro passado pelo brouser
        lErro = CF("Anotacoes_Le_NumIntDoc", objAnotacoes)
        If lErro <> SUCESSO And lErro <> 119857 Then gError 119818
    
        'se não encontrou => ERRO
        If lErro = 119857 Then gError 119870
    
    'senão, foi chamado de alguma tela
    Else
        
        'Verifica no bd se ja existe alguma anotação referente ao Id/ Origem
        lErro = CF("Anotacoes_Le_OrigemID", objAnotacoes)
        If lErro <> SUCESSO And lErro <> 119844 Then gError 119790
        
    End If
        
    'se encontrou
    If lErro = SUCESSO Then
    
        'traz a anotação mais recente p/ a tela
        lErro = Traz_Anotacoes_Tela(objAnotacoes)
        If lErro <> SUCESSO Then gError 119791
    
    Else
    
        'limpa a tela toda p/ evitar o que antes foi escrito continue aparecendo
        Call Limpa_Tela_Anotacoes
    
    End If

    Anotacoes_Valida_Preenchimento = SUCESSO
    
    Exit Function

Erro_Anotacoes_Valida_Preenchimento:

    Anotacoes_Valida_Preenchimento = gErr

    Select Case gErr

        Case 119789, 119790, 119791, 119818
    
        Case 119870
            Call Rotina_Erro(vbOKOnly, "ERRO_ANOTACAO_NAO_ENCONTRADA", gErr, objAnotacoes.sID)

    End Select

    Exit Function
    
End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objAnotacoes As New ClassAnotacoes

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Anotacoes"

    'move os dados essenciais da tela p/ o obj
    lErro = Move_Tela_Memoria(objAnotacoes)
    If lErro <> SUCESSO Then gError 119792
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Origem", objAnotacoes.iTipoDocOrigem, 0, "Origem"
    colCampoValor.Add "ID", objAnotacoes.sID, STRING_ANOTACOES_ID, "ID"
    colCampoValor.Add "DataAlteracao", objAnotacoes.dtDataAlteracao, 0, "DataAlteracao"
    colCampoValor.Add "Titulo", objAnotacoes.sTitulo, STRING_ANOTACOES_TITULO, "Titulo"

    'Filtros para o Sistema de Setas (tem que ter o mesmo ID/Origem)
    colSelecao.Add "Origem", OP_IGUAL, objAnotacoes.iTipoDocOrigem
    colSelecao.Add "ID", OP_IGUAL, objAnotacoes.sID

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 119792
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175970)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objAnotacoes As New ClassAnotacoes

On Error GoTo Erro_Tela_Preenche

    'preenche o obj c/ os dados lidos p/ ser passado como parametro
    objAnotacoes.iTipoDocOrigem = colCampoValor.Item("Origem").vValor
    objAnotacoes.sID = colCampoValor.Item("ID").vValor
    objAnotacoes.dtDataAlteracao = colCampoValor.Item("DataAlteracao").vValor
    objAnotacoes.sTitulo = colCampoValor.Item("Titulo").vValor

    'busca o numintdoc e traz o registro p/ a tela
    lErro = Traz_Anotacoes_Tela1(objAnotacoes)
    If lErro <> SUCESSO Then gError 119793

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 119793

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175971)

    End Select

    Exit Sub

End Sub

Private Function Traz_Anotacoes_Tela1(ByVal objAnotacoes As ClassAnotacoes) As Long
'busca o numintdoc referente ao titulo/data/id/origem e traz a anotacao p/ a tela
    
Dim lErro As Long

On Error GoTo Erro_Traz_Anotacoes_Tela1
    
    'busca o numintdoc p/ obter o texto referente a origem/id
    lErro = CF("Anotacoes_Le_NumIntDoc", objAnotacoes)
    If lErro <> SUCESSO And lErro <> 119857 Then gError 119794
    
    'não tem anotação(numintdoc) referente a origem/id/data/titulo
    If lErro = 119857 Then gError 119795
    
    'Traz a anotação para a Tela
    lErro = Traz_Anotacoes_Tela(objAnotacoes)
    If lErro <> SUCESSO Then gError 119796

    Traz_Anotacoes_Tela1 = SUCESSO

    Exit Function

Erro_Traz_Anotacoes_Tela1:

    Traz_Anotacoes_Tela1 = gErr

    Select Case gErr

        Case 119794, 119796

        Case 119795
            Call Rotina_Erro(vbOKOnly, "ERRO_ANOTACAO_NAO_ENCONTRADA", gErr, objAnotacoes.sID)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175972)

    End Select

    Exit Function

End Function

Private Function Traz_Anotacoes_Tela(ByVal objAnotacoes As ClassAnotacoes) As Long
'traz a anotação p/ a tela referente a origem e ao id

Dim lErro As Long
Dim sTextoMontado As String

On Error GoTo Erro_Traz_Anotacoes_Tela

    'limpa a tela
    Call Limpa_Tela_Anotacoes

    'coloca a data na tela c/ a que estava no bd
    Data.Caption = Format(objAnotacoes.dtDataAlteracao, "dd/mm/yyyy")

    'preenche o titulo da anotação
    Titulo.Text = objAnotacoes.sTitulo

    'busca o texto da anotação que foi "montado" a partir do numintdoc
    lErro = CF("AnotacoesLinha_Le_Texto1", sTextoMontado, objAnotacoes.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 119875 Then gError 119797

    'não foi encontrado o texto da anotação
    If lErro = 119875 Then gError 119858

    'coloca o texto completo na tela
    Anotacao.Text = sTextoMontado
    
    'zera a variavel de alteração
    iAlterado = 0

    Traz_Anotacoes_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Anotacoes_Tela:

    Traz_Anotacoes_Tela = gErr

    Select Case gErr

        Case 119797

        Case 119858
            Call Rotina_Erro(vbOKOnly, "ERRO_ANOTACOESLINHA_NAO_ENCONTRADA", gErr, objAnotacoes.sID)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175973)

    End Select

    Exit Function

End Function

Private Function Anotacoes_Preenche_Origem(ByVal objAnotacoes As ClassAnotacoes) As Long
'verifica se o Id existe na tabela referente a origem
'preenche o frame id/origem c/ oq foi passado como parametro

Dim lErro As Long
Dim sDescricao As String
Dim objOrigemAnotacoes As New ClassOrigemAnotacoes

On Error GoTo Erro_Anotacoes_Preenche_Origem

    'preenche o obj c/ o código
    objOrigemAnotacoes.iCodigo = objAnotacoes.iTipoDocOrigem

    'busca as informações da origem a partir do cód.
    lErro = CF("OrigemAnotacoes_Le", objOrigemAnotacoes)
    If lErro <> SUCESSO And lErro <> 119853 Then gError 119799

    'não encontrou nada referente ao cód.
    If lErro = 119853 Then gError 119800

    'coloca a origem na tela
    Origem.Caption = objOrigemAnotacoes.sDescricao

'    'verifica se o Id passado existe na tabela referente a origem
'    lErro = CF("OrigemAnotacoes_Verifica_ID", objOrigemAnotacoes.sNomeTabela, objAnotacoes.sID)
'    If lErro <> SUCESSO And lErro <> 119866 Then gError 119867
'
'    'não existe o ID (interno)
'    If lErro = 119866 Then gError 119868

    '??? criar funcao p/obter id externo a partir de sId
    
'    'coloca a ID (externo) na tela
'    IdOrigem.Caption = objAnotacoes.sIdExterno

    Exit Function

    Anotacoes_Preenche_Origem = SUCESSO

Erro_Anotacoes_Preenche_Origem:

    Anotacoes_Preenche_Origem = gErr

    Select Case gErr

        Case 119799, 119867

        Case 119800
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_DESCRICAO_NAO_ENCONTRADA", gErr, objOrigemAnotacoes.iCodigo)

        Case 119868
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEMID_NAO_EXISTENTE", gErr, objAnotacoes.sID, objOrigemAnotacoes.sNomeTabela, objOrigemAnotacoes.sDescricao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175974)
        
    End Select

    Exit Function

End Function

Public Sub Anotacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BotaoExcluir_Click()
'inicia a etapa de exclusão dos registros

Dim lErro As Long
Dim objAnotacoes As New ClassAnotacoes
Dim MsgBoxResp As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se o titulo foi preenchido
    If Len(Trim(Titulo.Text)) = 0 Then gError 119801
    
    'pergunta se deseja excluir a anotação
    MsgBoxResp = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_ANOTACAO", Origem.Caption, IdOrigem.Caption)

    'se não, sai da rotina
    If MsgBoxResp = vbNo Then Exit Sub

    'ponteiro p/ ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'move p/ o obj os dados necessário p/ a exclusão
    lErro = Move_Tela_Memoria(objAnotacoes)
    If lErro <> SUCESSO Then gError 119802
    
    'exclui o registro
    lErro = CF("Anotacoes_Exclui", objAnotacoes)
    If lErro <> SUCESSO Then gError 119803

    'limpa a tela
    Call Limpa_Tela_Anotacoes

    'volta o ponteiro ao normal
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    'volta o ponteiro ao normal
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 119802, 119803
    
        Case 119801
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_PREENCHIDO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175975)
            
    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()
'fecha a tela

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'se houve alteração, pergunta se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 119804

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case 119804
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175976)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()
'limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'se houve alteração, pergunta se quer salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 119805

    'limpa a tela
    Call Limpa_Tela_Anotacoes

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 119805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175977)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()
'carrega as configurações iniciais da tela

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'instancia o obj do browser de anotações
    Set objEventoAnotacao = New AdmEvento

    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175978)
    
    End Select
    
    Exit Sub

End Sub

Private Function Move_Tela_Memoria(ByVal objAnotacoes As ClassAnotacoes) As Long
'move todos os dados da tela p/ o obj que foi recebido como parametro

On Error GoTo Erro_Move_Tela_Memoria

    'carrega c/ a data que está na tela
    objAnotacoes.dtDataAlteracao = StrParaDate(Data.Caption)
    
    'carrega o obj c/ o titulo
    objAnotacoes.sTitulo = Trim(Titulo.Text)

    'carrega o obj c/ o Id de controle interno que foi passado
    objAnotacoes.sID = gsIdInterno
    
    'carrega o obj c/ o cód. da origem que está no frame de origem
    objAnotacoes.iTipoDocOrigem = giOrigem
    
    'adiciona o texto colmpleto no obj
    objAnotacoes.sTextoCompleto = Anotacao.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175979)

    End Select
    
    Exit Function
    
End Function

Public Sub BotaoGravar_Click()
'inicia a etapa de gravação
    
Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'chama a função que da continuidade a gavação
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 119808
    
    'limpa a tela
    Call Limpa_Tela_Anotacoes
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 119808
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175980)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Limpa_Tela_Anotacoes()
'limpa toda a tela, exeto o frame de id/origem

    'limpa toda a tela
    Call Limpa_Tela(Me)
    
    'coloca a data de hj na caption da label data
    Data.Caption = Format(gdtDataHoje, "dd/mm/yyyy")

    'zera a variavel de alteração
    iAlterado = 0

End Sub

Public Function Gravar_Registro() As Long
'continua a etapa de gravação

Dim objAnotacoes As New ClassAnotacoes
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    'verifica se o titulo foi preenchido
    If Len(Trim(Titulo.Text)) = 0 Then Titulo.Text = "X" 'gError 119809 'Alterado por Wagner

    'verifica se alguma coisa foi escrita na anotação
    If Len(Trim(Anotacao.Text)) = 0 Then gError 119810

    'coloca o mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'carrega o obj c/ os dados da tela
    lErro = Move_Tela_Memoria(objAnotacoes)
    If lErro <> SUCESSO Then gError 119811
    
    'verifica se ja existe algum registro, se existir, pergunta se deseja atualizar o registro existente
    lErro = Trata_Alteracao(objAnotacoes, objAnotacoes.sID, objAnotacoes.iTipoDocOrigem, objAnotacoes.sTitulo, objAnotacoes.dtDataAlteracao)
    If lErro <> SUCESSO Then gError 119869

    'grava a anotação na tabela anotacoes e anotacoeslinha
    lErro = CF("Anotacoes_Grava", objAnotacoes)
    If lErro <> SUCESSO Then gError 119812

    'volta o ponteiro ao normal
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    'volta o ponteiro ao normal
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = gErr

    Select Case gErr

        Case 119811, 119812, 119869
    
        Case 119809
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_PREENCHIDO", gErr)

        Case 119810
            Call Rotina_Erro(vbOKOnly, "ERRO_ANOTACAO_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175981)
            
    End Select
    
    Exit Function

End Function

Public Sub LabelTitulo_Click()
'sub que traz o browser de anotacoes p/ a tela

Dim objAnotacoes As New ClassAnotacoes
Dim colSelecao As New Collection
Dim lErro As Long
Dim sSelecao As String

On Error GoTo Erro_LabelTitulo_Click

    'preenche o obj c/ os dados da tela
    lErro = Move_Tela_Memoria(objAnotacoes)
    If lErro <> SUCESSO Then gError 119861
        
    'adiciona os valores p/ seleção dinamica
    colSelecao.Add CStr(giOrigem)
    colSelecao.Add CStr(gsIdInterno)
        
    'adiciona os parametros de seleção dinamica
    sSelecao = "Origem = ? AND ID = ?"
        
    'chama a tela de anotações cadastradas
    Call Chama_Tela("AnotacoesLista", colSelecao, objAnotacoes, objEventoAnotacao, sSelecao)

    Exit Sub

Erro_LabelTitulo_Click:

    Select Case gErr

        Case 119861
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175982)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoAnotacao_evSelecao(obj1 As Object)
'traz p/ a tela a anotação selecionada

Dim objAnotacoes As ClassAnotacoes
Dim lErro As Long

On Error GoTo Erro_objEventoAnotacao_evSelecao

    Set objAnotacoes = obj1

    'traz p/ a tela a anotacao de acordo c/ a escolha
    lErro = Traz_Anotacoes_Tela(objAnotacoes)
    If lErro <> SUCESSO Then gError 119862

    Me.Show

    Exit Sub

Erro_objEventoAnotacao_evSelecao:

    Select Case gErr

        Case 119862

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175983)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub Titulo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Titulo Then
            Call LabelTitulo_Click
        End If
        
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    Set objEventoAnotacao = Nothing
    
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Anotações"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Anotacoes"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Public Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******

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

