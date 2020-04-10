VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTSerieNFiscal"
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

'DECLARACAO DE VARIAVEIS GLOBAIS
Public iAlterado As Integer
Public iSerieAlterada As Integer

'--- inicio dos properties get dos controles da tela

Public Property Get Eletronica() As Object
     Set Eletronica = objUserControl.Controls("Eletronica")
End Property

Public Property Get NomeTsk() As Object
     Set NomeTsk = objUserControl.Controls("NomeTsk")
End Property

Public Property Get Padrao() As Object
     Set Padrao = objUserControl.Controls("Padrao")
End Property

Public Property Get TipoFormulario() As Object
     Set TipoFormulario = objUserControl.Controls("TipoFormulario")
End Property

Public Property Get Serie() As Object
     Set Serie = objUserControl.Controls("Serie")
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

Public Property Get ProxNumNFiscal() As Object
     Set ProxNumNFiscal = objUserControl.Controls("ProxNumNFiscal")
End Property

Public Property Get MaxLinhasNF() As Object
     Set MaxLinhasNF = objUserControl.Controls("MaxLinhasNF")
End Property

Public Property Get LabelMaxLinhas() As Object
     Set LabelMaxLinhas = objUserControl.Controls("LabelMaxLinhas")
End Property

Public Property Get ModDocFis() As Object
     Set ModDocFis = objUserControl.Controls("ModDocFis")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Form_Load()

Dim objSerie As New ClassSerie
Dim colSerie As New colSerie
Dim lErro As Long

On Error GoTo Erro_Form_Load
        
    'Lê FilialEmpresa e Serie da tabela Serie e devolve na coleção
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then Error 22171

    'Preenche a ListBox Series com os objetos da coleção
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next

    TipoFormulario.ListIndex = 0
    
    'Eletronica.Value = vbunChecked
    
    lErro = CF("Carrega_Combo_ModDocFis", ModDocFis)
    If lErro <> SUCESSO Then Error 22171

    Call Combo_Seleciona_ItemData(ModDocFis, 31)
    
    Call CF2(Me, "SNF_Form_Load")
    
    iSerieAlterada = 0
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 22171

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174429)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objSerie As New ClassSerie

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Serie"

    'Le os dados da Tela SerieNFiscal
    lErro = Move_Tela_Memoria(objSerie)
    If lErro <> SUCESSO Then Error 22172

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objSerie.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Serie", objSerie.sSerie, STRING_SERIE, "Serie"
    colCampoValor.Add "ProxNumNFiscal", objSerie.lProxNumNFiscal, 0, "ProxNumNFiscal"
    colCampoValor.Add "TipoFormulario", objSerie.iTipoFormulario, 0, "TipoFormulario"
        
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 22172

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174430)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objSerie As New ClassSerie

On Error GoTo Erro_Tela_Preenche

    'Carrega objSerie com os dados passados em colCampoValor
    objSerie.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objSerie.sSerie = colCampoValor.Item("Serie").vValor
    objSerie.lProxNumNFiscal = colCampoValor.Item("ProxNumNFiscal").vValor
    objSerie.iTipoFormulario = colCampoValor.Item("TipoFormulario").vValor
    
    lErro = CF("Serie_Le", objSerie)
    If lErro <> SUCESSO And lErro <> 22202 Then Error 22173
    
    'Traz dados da Serie para a Tela
    lErro = Traz_Serie_Tela(objSerie)
    If lErro <> SUCESSO Then Error 22173

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 22173

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174431)

    End Select

    Exit Sub

End Sub

Function Traz_Serie_Tela(objSerie As ClassSerie) As Long
'Mostra os dados na tela

Dim iIndice As Integer
Dim lErro As Long
Dim sSeriePadrao As String

On Error GoTo Erro_Traz_Serie_Tela

    lErro = CF2(Me, "SNF_Traz_Serie_Tela1", objSerie)
    If lErro <> SUCESSO Then gError 64118
                    
    Serie.Text = Desconverte_Serie_Eletronica(objSerie.sSerie)
    ProxNumNFiscal.Text = CStr(objSerie.lProxNumNFiscal)
    NomeTsk.Text = objSerie.sNomeTsk
        
    For iIndice = 0 To TipoFormulario.ListCount - 1
        If TipoFormulario.ItemData(iIndice) = objSerie.iTipoFormulario Then
            TipoFormulario.ListIndex = iIndice
            Exit For
        End If
    Next
                                    
    'Lê a Serie Padrão
    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
    If lErro <> SUCESSO Then Error 64118

    'Se a Serie Padrão for igual a Série
    If sSeriePadrao = objSerie.sSerie Then
        'Então marca o Padrao
        Padrao.Value = vbChecked
    Else
        Padrao.Value = vbUnchecked
    End If
    
    If objSerie.iEletronica = MARCADO Then
        Eletronica.Value = vbChecked
    Else
        Eletronica.Value = vbUnchecked
    End If
    
    Call Combo_Seleciona_ItemData(ModDocFis, objSerie.iModDocFis)
        
    MaxLinhasNF.Text = CStr(objSerie.iMaxLinhasNF)
    
    lErro = CF2(Me, "SNF_Traz_Serie_Tela2", objSerie)
    If lErro <> SUCESSO Then gError 64118
    
    iSerieAlterada = 0
        
    iAlterado = 0

    Traz_Serie_Tela = SUCESSO

    Exit Function
    
Erro_Traz_Serie_Tela:

    Traz_Serie_Tela = Err
    
    Select Case Err
        
        Case 64118
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174432)

    End Select
    
    iAlterado = 0

    Exit Function
    
End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Function Trata_Parametros(Optional objSerie As ClassSerie) As Long

Dim lErro As Long
Dim sListBoxItem As String
Dim iIndice As Integer
Dim iCodigo As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há uma Serie selecionada
    If Not (objSerie Is Nothing) Then

        'Verifica se a Serie existe, lendo no BD a partir de FilialEmpresa e Serie
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then Error 22174

        'Se a Serie existe
        If lErro = SUCESSO Then

           lErro = Traz_Serie_Tela(objSerie)

        'Se a Serie não existe
        Else

            'Exibe apenas a Serie
            Serie.Text = objSerie.sSerie

        End If

    End If

    'Zerar iAlterado
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 22174

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174433)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Padrao_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ProxNumNFiscal_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ProxNumNFiscal_GotFocus()
        
    Call MaskEdBox_TrataGotFocus(ProxNumNFiscal, iAlterado)

End Sub

Public Sub ProxNumNFiscal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProxNumNFiscal_Validate

    'Verifica se foi preenchido o campo ProxNumNFiscal
    If Len(Trim(ProxNumNFiscal.Text)) = 0 Then Exit Sub

    lErro = Long_Critica(ProxNumNFiscal.Text)
    If lErro <> SUCESSO Then Error 22178

    Exit Sub

Erro_ProxNumNFiscal_Validate:

    Cancel = True


    Select Case Err

        Case 22178

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174434)

    End Select

    Exit Sub
    
End Sub

Public Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = 1

End Sub

Public Sub Serie_Click()

Dim lErro As Long
Dim sListBoxItem As String
Dim objSerie As New ClassSerie

On Error GoTo Erro_Serie_Click
    
    If Len(Trim(Serie.Text)) = 0 Then
        ProxNumNFiscal.Text = ""
        MaxLinhasNF.Text = ""
        Call CF2(Me, "SNF_Serie_Click")
        iSerieAlterada = 0
        iAlterado = 0
        Exit Sub
    End If
    
    'Guarda o valor da Serie selecionada na ListBox Series
    objSerie.sSerie = Serie.Text

    'Lê a Serie no BD
    lErro = CF("Serie_Le", objSerie)
    If lErro <> SUCESSO And lErro <> 22202 Then Error 22175

    'Se Serie não está cadastrada, erro
    If lErro = 12348 Then Error 22176

    'Exibe os dados da Serie
    lErro = Traz_Serie_Tela(objSerie)
    If lErro <> SUCESSO Then Error 22177

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Serie_Click:

    Select Case Err

    Case 22175, 22177

    Case 22176
        lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, objSerie.sSerie)
        Serie.RemoveItem (Serie.ListIndex)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174435)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava a Serie
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 22179

    'Limpa a Tela
    Call Limpa_Tela_Serie

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 22179

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174436)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Grava Serie no BD

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim bPadrao As Boolean

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Serie foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 22180

    
    'Verifica se a Serie contem apenas numeros
    'If IsNumeric(Serie.Text) = False Then gError 76242
    'Verifica se a Serie ultrapassou o tamanho maximo de 2
    'If Len(Trim(Serie.Text)) > 2 Then gError 76243
    
    'Verifica se ProxNumNFiscal foi preenchida
    If Len(Trim(ProxNumNFiscal.Text)) = 0 Then gError 33625
    
    lErro = CF2(Me, "SNF_Gravar_Registro")
    If lErro <> SUCESSO Then gError 22182

    'Preenche os objetos com os dados da tela
    lErro = Move_Tela_Memoria(objSerie)
    If lErro <> SUCESSO Then gError 22182
    
    lErro = Trata_Alteracao(objSerie, objSerie.iFilialEmpresa, objSerie.sSerie)
    If lErro <> SUCESSO Then gError 32317
        
    If Padrao.Value = vbChecked Then
        bPadrao = True
    Else
        bPadrao = False
    End If
    
    'Grava a Serie no BD
    lErro = CF("Serie_Grava", objSerie, bPadrao)
    If lErro <> SUCESSO Then gError 22183

    'Atualiza ListBox Serie
    Call Serie_Remove(objSerie)
    Call Serie_Adiciona(objSerie)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 32317

        Case 22180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 22181, 33625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROXNUMNFISCAL_NAO_PREENCHIDA", gErr)

        Case 22182, 22183

    '    Case 76242
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NUMERICA", gErr)
            
     '   Case 76243
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_MAIOR", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174437)

    End Select

    Exit Function

End Function

Public Sub Serie_Remove(objSerie As ClassSerie)
'Percorre a ListBox Series para remover a Serie caso ela exista

Dim iIndice As Integer

    For iIndice = 0 To Serie.ListCount - 1
    
        If Serie.List(iIndice) = objSerie.sSerie Then
    
            Serie.RemoveItem iIndice
            Exit For
    
        End If
    
    Next

End Sub

Public Sub Serie_Adiciona(objSerie As ClassSerie)
'Inclui Serie na List
Dim iIndice As Integer

    Serie.AddItem objSerie.sSerie

End Sub

Private Function Move_Tela_Memoria(objSerie As ClassSerie) As Long
'Lê os dados que estão na tela SerieNFiscal e coloca em objSerie

Dim lErro As Long

    'IDENTIFICACAO :
    If Len(Trim(Serie.Text)) > 0 Then objSerie.sSerie = Converte_Serie_Eletronica(Trim(Serie.Text), Eletronica.Value)

    If Len(Trim(ProxNumNFiscal.Text)) > 0 Then objSerie.lProxNumNFiscal = CLng(ProxNumNFiscal.Text)
    
    objSerie.iFilialEmpresa = giFilialEmpresa
    
    objSerie.iTipoFormulario = TipoFormulario.ItemData(TipoFormulario.ListIndex)

    If Len(Trim(NomeTsk.Text)) > 0 Then objSerie.sNomeTsk = NomeTsk.Text
    
    If Eletronica.Value = vbChecked Then
        objSerie.iEletronica = MARCADO
    Else
        objSerie.iEletronica = DESMARCADO
    End If
    
    objSerie.iModDocFis = ModDocFis.ItemData(ModDocFis.ListIndex)

    objSerie.iMaxLinhasNF = StrParaInt(MaxLinhasNF.Text)
    
    Call CF2(Me, "SNF_Move_Tela_Memoria", objSerie)

    Move_Tela_Memoria = SUCESSO

    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim colCodNomeFiliais As New AdmColCodigoNome
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a Serie foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Error 22185
    
    objSerie.sSerie = Converte_Serie_Eletronica(Trim(Serie.Text), Eletronica.Value) 'Serie.Text
        
    'Lê os dados da Serie a ser excluida
    lErro = CF("Serie_Le", objSerie)
    If lErro <> SUCESSO And lErro <> 22202 Then Error 22186

    'Verifica se Serie está cadastrada
    If lErro <> SUCESSO Then Error 22187

    'Envia aviso perguntando se realmente deseja excluir Serie
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_SERIE", objSerie.sSerie)

    If vbMsgRes = vbYes Then

        'Exclui Serie
        lErro = CF("Serie_Exclui", objSerie)
        If lErro <> SUCESSO Then Error 22188

        'Exclui da ListBox
        Call Serie_Remove(objSerie)

        'Limpa a Tela
        Call Limpa_Tela_Serie

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 22185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)

        Case 22186, 22188

        Case 22187   'Serie não está cadastrada
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, objSerie.sSerie)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174438)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 22184

    'Limpa a Tela
    Call Limpa_Tela_Serie

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 22184

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174439)

    End Select

End Sub

Sub Limpa_Tela_Serie()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa TextBox e MaskedEditBox
    Call Limpa_Tela(Me)
    Serie.Text = ""
    Serie.ListIndex = -1
    TipoFormulario.ListIndex = 0
    Padrao.Value = Unchecked
    iAlterado = 0
    'Eletronica.Value = vbnChecked
    Call Combo_Seleciona_ItemData(ModDocFis, 31)

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_SERIE_NOTA_FISCAL
    Set Form_Load_Ocx = Me
    Caption = "Séries de Notas Fiscais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "SerieNFiscal"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub



Public Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sListBoxItem As String
Dim objSerie As New ClassSerie

On Error GoTo Erro_Serie_Validate

    If iSerieAlterada = 0 Then Exit Sub
    
    If Len(Trim(Serie.Text)) = 0 Then
        ProxNumNFiscal.Text = ""
        MaxLinhasNF.Text = ""
        Call CF2(Me, "SNF_Serie_Validate1")
        iAlterado = 0
        Exit Sub
    End If
    
    If Len(Trim(Serie.Text)) > STRING_SERIE Then gError 64456
        
    'Guarda o valor da Serie selecionada na ListBox Series
    objSerie.sSerie = Serie.Text

    'Lê a Serie no BD
    lErro = CF("Serie_Le", objSerie)
    If lErro <> SUCESSO And lErro <> 22202 Then gError 60362

    'Se Serie não está cadastrada, erro
    If lErro = SUCESSO Then
        'Exibe os dados da Serie
        lErro = Traz_Serie_Tela(objSerie)
        If lErro <> SUCESSO Then gError 60364
    Else
        ProxNumNFiscal.Text = ""
        MaxLinhasNF.Text = ""
        Serie.Text = Desconverte_Serie_Eletronica(objSerie.sSerie)
        If ISSerieEletronica(objSerie.sSerie) Then
            Eletronica.Value = vbChecked
        Else
            Eletronica.Value = vbUnchecked
        End If
        lErro = CF2(Me, "SNF_Serie_Validate2", objSerie)
        If lErro <> SUCESSO Then gError 60364
        
        Call Trata_Modelo_Fiscal
    End If
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iSerieAlterada = 0

    Exit Sub

Erro_Serie_Validate:

    Cancel = True
    
    Select Case gErr
    
    Case 60362, 60364
        
    Case 64456
        lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_MAIOR", gErr)
        
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174440)

    End Select

    Exit Sub

End Sub

Public Sub TipoFormulario_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub TipoFormulario_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Unload(objme As Object)
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








Public Sub MaxLinhasNF_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub MaxLinhasNF_GotFocus()
        
    Call MaskEdBox_TrataGotFocus(MaxLinhasNF, iAlterado)

End Sub

Public Sub MaxLinhasNF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaxLinhasNF_Validate

    Call CF2(Me, "SNF_MaxLinhasNF_Validate")

    'Verifica se foi preenchido o campo MaxLinhasNF
    If Len(Trim(MaxLinhasNF.Text)) = 0 Then Exit Sub

    lErro = Valor_NaoNegativo_Critica(MaxLinhasNF.Text)
    If lErro <> SUCESSO Then Error 22178

    Exit Sub

Erro_MaxLinhasNF_Validate:

    Cancel = True

    Select Case Err

        Case 22178

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174428)

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

Public Function SNF_Form_Load(ByVal objCT As Object) As Long
    SNF_Form_Load = SUCESSO
End Function
Public Function SNF_Gravar_Registro(ByVal objCT As Object) As Long
    SNF_Gravar_Registro = SUCESSO
End Function
Public Function SNF_MaxLinhasNF_Validate(ByVal objCT As Object) As Long
    SNF_MaxLinhasNF_Validate = SUCESSO
End Function
Public Function SNF_Move_Tela_Memoria(ByVal objCT As Object, ByVal objSerie As ClassSerie) As Long
    SNF_Move_Tela_Memoria = SUCESSO
End Function
Public Function SNF_Serie_Click(ByVal objCT As Object) As Long
    SNF_Serie_Click = SUCESSO
End Function
Public Function SNF_Serie_Validate1(ByVal objCT As Object) As Long
    SNF_Serie_Validate1 = SUCESSO
End Function
Public Function SNF_Serie_Validate2(ByVal objCT As Object, ByVal objSerie As ClassSerie) As Long
    SNF_Serie_Validate2 = SUCESSO
End Function
Public Function SNF_Traz_Serie_Tela1(ByVal objCT As Object, objSerie As ClassSerie) As Long
    SNF_Traz_Serie_Tela1 = SUCESSO
End Function
Public Function SNF_Traz_Serie_Tela2(ByVal objCT As Object, objSerie As ClassSerie) As Long
    SNF_Traz_Serie_Tela2 = SUCESSO
End Function

Public Sub ModDocFis_Change()
     iAlterado = REGISTRO_ALTERADO
     Call Trata_Modelo_Fiscal
End Sub

Public Sub ModDocFis_Click()
     Call Trata_Modelo_Fiscal
End Sub

Private Sub Trata_Modelo_Fiscal()

Dim lErro As Long, iIndice As Integer
Dim objMod As New ClassModelosDocFiscais
Dim iTipoFormulario As Integer

On Error GoTo Erro_Trata_Modelo_Fiscal

    If ModDocFis.ListIndex <> -1 Then
    
        objMod.iCodigo = ModDocFis.ItemData(ModDocFis.ListIndex)
        
        lErro = CF("ModelosDocFiscais_Le", objMod)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objMod.iEletronica = MARCADO And objMod.iTipo = 0 Then
            Eletronica.Value = vbChecked
        Else
            Eletronica.Value = vbUnchecked
        End If
        
        '1-N.Fiscal
        '2-N.FiscalFatura
        '3-N.FiscalServiço
        '4-N.FiscalFaturaServiço
        '5-Conhecimento Transporte
        '6-Conhecimento Transporte Fatura
        Select Case objMod.iCodigo
            
            Case 9, 32
                iTipoFormulario = 6 'CT/CT-e
            
            Case 900 To 999
                iTipoFormulario = 4 'Serviço
                
            Case Else
                iTipoFormulario = 2
            
        End Select
        
        For iIndice = 0 To TipoFormulario.ListCount - 1
            If TipoFormulario.ItemData(iIndice) = iTipoFormulario Then
                TipoFormulario.ListIndex = iIndice
                Exit For
            End If
        Next
    
    End If

    Exit Sub

Erro_Trata_Modelo_Fiscal:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174428)

    End Select

    Exit Sub
    
End Sub
