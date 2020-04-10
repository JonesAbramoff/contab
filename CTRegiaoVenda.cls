VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTRegiaoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjInfoUsu As Object
Private WithEvents objEventoRegiao As AdmEvento
Attribute objEventoRegiao.VB_VarHelpID = -1

'DECLARACAO DE VARIAVEIS GLOBAIS
Public iAlterado As Integer

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get Descricao() As Object
     Set Descricao = objUserControl.Controls("Descricao")
End Property

Public Property Get Regioes() As Object
     Set Regioes = objUserControl.Controls("Regioes")
End Property

Public Property Get comboCobrador() As Object
     Set comboCobrador = objUserControl.Controls("ComboCobrador")
End Property

Public Property Get UsuRespCallCenter() As Object
     Set UsuRespCallCenter = objUserControl.Controls("UsuRespCallCenter")
End Property

Public Property Get Pais() As Object
     Set Pais = objUserControl.Controls("Pais")
End Property

Public Property Get Gerente() As Object
     Set Gerente = objUserControl.Controls("Gerente")
End Property

Public Property Get Label2() As Object
     Set Label2 = objUserControl.Controls("Label2")
End Property

Public Property Get label1() As Object
     Set label1 = objUserControl.Controls("Label1")
End Property

Public Property Get Label13() As Object
     Set Label13 = objUserControl.Controls("Label13")
End Property

Public Property Get Label63() As Object
     Set Label63 = objUserControl.Controls("Label63")
End Property

Public Property Get Label70() As Object
     Set Label70 = objUserControl.Controls("Label70")
End Property

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera codigo automatico da proxima Regiao de Venda
    lErro = CF("RegiaoVenda_Automatico", iCodigo)
    If lErro <> SUCESSO Then Error 57554

    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57554
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 166556)
    
    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer
Dim sCodigo As String
Dim iEncontrou As Integer
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 16190

    objRegiaoVenda.iCodigo = CInt(Codigo.Text)
    
    'Le a Região de Venda para Verificar se esta no BD
    lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
    If lErro = 16137 Then Error 16191
    
    'Se não estiver ---> ERRO
    If lErro <> SUCESSO Then Error 16192
    
    'Pede confirmação de exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_REGIOESVENDAS", objRegiaoVenda.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui Região de Venda
        lErro = CF("RegiaoVenda_Exclui", objRegiaoVenda)
        If lErro <> SUCESSO Then Error 16199

        Call Regioes_Remove(objRegiaoVenda)
        
        'Limpa a Tela
        Call Limpa_Tela_RegiaoVenda

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", Err, objRegiaoVenda.iCodigo)
        
        Case 16192, 16199

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166557)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long
Dim sCodigo As String

On Error GoTo Erro_BotaoGravar_Click

    'Grava a Região de Venda
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 16208
    
    'Limpa Tela
    Call Limpa_Tela_RegiaoVenda

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 16208
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166558)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate
    
    'verifica se codigo está preenchido
    If Len(Trim(Codigo.Text)) > 0 Then
    
        'verifica se codigo é numérico
        If Not IsNumeric(Codigo.Text) Then Error 16388
        
        'verifica se codigo é menor que um
        If CInt(Codigo.Text) < 1 Then Error 16387
        
    End If
        
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True

    Select Case Err

        Case 16387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", Err, Codigo.Text)

        Case 16388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", Err, Codigo.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166559)
    
    End Select

    Exit Sub
End Sub

Public Sub Form_Activate()
    
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    Set objEventoRegiao = Nothing
    
     'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub

Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub Limpa_Tela_RegiaoVenda()
'limpa todos os campos de input da tela RegiaoVenda

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_RegiaoVenda

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
        
    Call Limpa_Tela(Me)
    
    'Desmarca ListBox de Regioes de Vendas
    Regioes.ListIndex = -1
    
    comboCobrador.Text = ""
    UsuRespCallCenter.Text = ""
    
    Codigo.Text = ""

    Call CF2(Me, "RG_Limpa_Tela_RegiaoVenda")

    iAlterado = 0

    Exit Sub
    
Erro_Limpa_Tela_RegiaoVenda:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166560)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long
Dim sCodigo As String

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 16201

    'Limpa a Tela
    Call Limpa_Tela_RegiaoVenda
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 16201

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166561)

    End Select

    Exit Sub
    
End Sub

Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome

On Error GoTo Erro_Form_Load

    iAlterado = 0
    
    Set objEventoRegiao = New AdmEvento

    'Lê códigos e descrições das Regiões de Venda no BD
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 16132

    'Preenche listbox com descricao das regioes de venda
    For Each objCodigoDescricao In colCodigoDescricao
        Regioes.AddItem objCodigoDescricao.sNome
        Regioes.ItemData(Regioes.NewIndex) = objCodigoDescricao.iCodigo
    Next

    'reinicializa a colecao colCodigoDescricao
    Set colCodigoDescricao = New AdmColCodigoNome

    'leitura dos codigos e nome dos paises no BD
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 16133

    'preenche ComboBox com código e nome dos Paises
    For Each objCodigoDescricao In colCodigoDescricao
        Pais.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Pais.ItemData(Pais.NewIndex) = objCodigoDescricao.iCodigo
    Next

    'Seleciona Brasil no ComboBox
    For iIndice = 0 To Pais.ListCount - 1
        If right(Pais.List(iIndice), 6) = "Brasil" Then
            Pais.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Call Carrega_Usuarios
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 16132, 16133

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166562)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objRegiaoVenda As ClassRegiaoVenda) As Long
'Trata os parametros que podem ser passados quando ocorre a chamada da tela RegiaoVenda

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Verifica se houve passagem de parametro
    If Not (objRegiaoVenda Is Nothing) Then

        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then Error 16148

        If lErro = SUCESSO Then
            
            'Traz a Região para Tela
            Call Traz_Regiao_Tela(objRegiaoVenda)

        Else

            'mostra apenas o codigo da Regiao de Venda
            Codigo.Text = CStr(objRegiaoVenda.iCodigo)

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 16148

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166563)

    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim iPosicao As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 16166

    'verifica preenchimento da descricao
    If Len(Trim(Descricao.Text)) = 0 Then Error 16167

    'verifica preenchimento do pais
    If Len(Trim(Pais.Text)) = 0 Then Error 16207

    'preenche objRegiaoVenda
    objRegiaoVenda.iCodigo = CInt(Codigo.Text)
    objRegiaoVenda.sDescricao = Descricao.Text
    objRegiaoVenda.sUsuarioCobrador = comboCobrador.Text
    objRegiaoVenda.sUsuRespCallCenter = UsuRespCallCenter.Text
    
    If Len(Trim(Pais.Text)) > 0 Then
        objRegiaoVenda.iCodigoPais = Codigo_Extrai(Pais.Text)
    Else
        objRegiaoVenda.iCodigoPais = 0
    End If
    
    objRegiaoVenda.sGerente = Gerente.Text

    lErro = Trata_Alteracao(objRegiaoVenda, objRegiaoVenda.iCodigo)
    If lErro <> SUCESSO Then Error 32286

    lErro = CF2(Me, "RG_Grava_Registro", objRegiaoVenda)
    If lErro <> SUCESSO Then gError 195869

    lErro = CF("RegiaoVenda_Grava", objRegiaoVenda)
    If lErro <> SUCESSO Then Error 16168

    'Remove e adiciona na ListBox
    Call Regioes_Remove(objRegiaoVenda)
    Call Regioes_Adiciona(objRegiaoVenda)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", Err)

        Case 16168

        Case 16207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_PREENCHIDO", Err)

        Case 32286

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166564)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub Gerente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Pais_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Pais_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Pais_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Pais_Validate

    'Verifica se foi preenchido o Pais
    If Len(Trim(Pais.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o item selecionado na ComboBox Pais
    If Pais.Text = Pais.List(Pais.ListIndex) Then Exit Sub

    'Verifica se existe o item Pais na List da Combo, se existir seleciona o item
    lErro = Combo_Seleciona(Pais, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 16150

    'Nao existe o item com o CODIGO na List da ComboBox
    If lErro = 6730 Then Error 16151

    'Nao existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then Error 16152

    Exit Sub

Erro_Pais_Validate:

    Cancel = True
   
    Select Case Err

        Case 16150

        Case 16151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", Err, iCodigo)
            
        Case 16152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", Err, Trim(Pais.Text))
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166565)

    End Select

    Exit Sub

End Sub

Public Sub Regioes_DblClick()

Dim lErro As Long
Dim iIndice As Integer
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_Regioes_DblClick
        
    objRegiaoVenda.iCodigo = Regioes.ItemData(Regioes.ListIndex)
    
    'Lê a Regiao e seus dados
    lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
    If lErro <> SUCESSO And lErro <> 16137 Then Error 16148
    If lErro <> SUCESSO Then Error 16149  'Não encontrou Região
    
    'Traz a Região para Tela
    Call Traz_Regiao_Tela(objRegiaoVenda)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_Regioes_DblClick:

    Select Case Err

        Case 16148 'Tratado na Rotina Chamada

        Case 16149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA1", Err, Regioes.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 166566)

    End Select

    Exit Sub
End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim iIndice As Integer, objRegiaoVenda As New ClassRegiaoVenda
Dim iCodigoPais As String, lErro As Long

On Error GoTo Erro_Tela_Preenche

    objRegiaoVenda.iCodigo = colCampoValor.Item("Codigo").vValor
    objRegiaoVenda.sDescricao = colCampoValor.Item("Descricao").vValor
    objRegiaoVenda.sGerente = colCampoValor.Item("Gerente").vValor
    objRegiaoVenda.iCodigoPais = colCampoValor.Item("CodigoPais").vValor
    
    lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
    If lErro <> SUCESSO And lErro <> 16137 Then gError ERRO_SEM_MENSAGEM
    
    lErro = Traz_Regiao_Tela(objRegiaoVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166566)

    End Select

    Exit Sub
    
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim iCodigo As Integer
Dim iCodigoPais As Integer
Dim iPosicao As Integer
    
    'Informa tabela associada à Tela
    sTabela = "RegioesVendas"
    
    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    If Len(Trim(Codigo.Text)) > 0 Then iCodigo = CInt(Codigo.Text)
 
    If Len(Trim(Pais.Text)) > 0 Then
        iCodigoPais = Codigo_Extrai(Pais.Text)
    Else
        iCodigoPais = 0
    End If
       
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", Descricao.Text, STRING_DESCRICAO, "Descricao"
    colCampoValor.Add "CodigoPais", iCodigoPais, 0, "CodigoPais"
    colCampoValor.Add "Gerente", Gerente.Text, STRING_GERENTE, "Gerente"
    
End Sub

Private Sub Regioes_Adiciona(objRegiaoVenda As ClassRegiaoVenda)
'Inclui na List

    'Insere Região de venda na ListBox
    Regioes.AddItem objRegiaoVenda.sDescricao
    Regioes.ItemData(Regioes.NewIndex) = objRegiaoVenda.iCodigo

End Sub

Private Sub Regioes_Remove(objRegiaoVenda As ClassRegiaoVenda)
'Percorre a ListBox Regiões para remover o tipo caso ele exista

Dim iIndice As Integer

    For iIndice = 0 To Regioes.ListCount - 1
    
        If Regioes.ItemData(iIndice) = objRegiaoVenda.iCodigo Then
    
            Regioes.RemoveItem iIndice
            Exit For
    
        End If
    
    Next

End Sub

Private Function Traz_Regiao_Tela(objRegiaoVenda As ClassRegiaoVenda) As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Traz_Regiao_Tela

    'Mostra dados da Região de Venda na Tela
    Codigo.Text = objRegiaoVenda.iCodigo
    Descricao.Text = objRegiaoVenda.sDescricao
    Gerente.Text = objRegiaoVenda.sGerente

    For iIndice = 0 To Pais.ListCount - 1
        
        If Pais.ItemData(iIndice) = objRegiaoVenda.iCodigoPais Then
            Pais.ListIndex = iIndice
            Exit For
        End If
        
    Next
    
    comboCobrador.Text = objRegiaoVenda.sUsuarioCobrador
    Call comboCobrador_Validate(bSGECancelDummy)
    
    UsuRespCallCenter.Text = objRegiaoVenda.sUsuRespCallCenter
    Call UsuRespCallCenter_Validate(bSGECancelDummy)
    
    lErro = CF2(Me, "RG_Traz_Regiao_Tela", objRegiaoVenda)
    If lErro <> SUCESSO Then gError 197195
    
    iAlterado = 0
    
    Traz_Regiao_Tela = SUCESSO

    Exit Function

Erro_Traz_Regiao_Tela:

    Traz_Regiao_Tela = Err

    Select Case Err

        Case 197195

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 197196)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_REGIOES_VENDA
    Set Form_Load_Ocx = Me
    Caption = "Regiões de Venda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RegiaoVenda"
    
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

'***** fim do trecho a ser copiado ******

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
End Sub

Public Sub comboCobrador_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub comboCobrador_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_ComboCobrador_Validate
    
    'Verifica se algum codigo está selecionado
    If comboCobrador.ListIndex = -1 Then Exit Sub
    
    'Coloca o código selecionado nos obj's
    objUsuarios.sCodUsuario = comboCobrador.Text

    'Le o nome do Usário
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError 182172
    
    If lErro <> SUCESSO Then gError 182173
    
    Exit Sub
    
Erro_ComboCobrador_Validate:

    Cancel = True

    Select Case gErr
            
        Case 182172
        
        Case 182173 'O usuário não está na tabela
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182174)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub UsuRespCallCenter_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UsuRespCallCenter_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_UsuRespCallCenter_Validate
    
    'Verifica se algum codigo está selecionado
    If UsuRespCallCenter.ListIndex = -1 Then Exit Sub
    
    'Coloca o código selecionado nos obj's
    objUsuarios.sCodUsuario = UsuRespCallCenter.Text

    'Le o nome do Usário
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError 182172
    
    If lErro <> SUCESSO Then gError 182173
    
    Exit Sub
    
Erro_UsuRespCallCenter_Validate:

    Cancel = True

    Select Case gErr
            
        Case 182172
        
        Case 182173 'O usuário não está na tabela
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182174)
    
    End Select
    
    Exit Sub
    
End Sub

Private Function Carrega_Usuarios() As Long
'Carrega a Combo CodUsuarios com todos os usuários do BD

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 182135

    For Each objUsuarios In colUsuarios
        comboCobrador.AddItem objUsuarios.sCodUsuario
        UsuRespCallCenter.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = gErr

    Select Case gErr

        Case 182135

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182136)

    End Select

    Exit Function

End Function

Public Function RG_Grava_Registro(ByVal objCT As Object, ByVal objRegiaoVenda As ClassRegiaoVenda) As Long
    RG_Grava_Registro = SUCESSO
End Function

Public Function RG_Limpa_Tela_RegiaoVenda(ByVal objCT As Object) As Long
    RG_Limpa_Tela_RegiaoVenda = SUCESSO
End Function

Public Function RG_Traz_Regiao_Tela(ByVal objCT As Object, ByVal objRegiaoVenda As ClassRegiaoVenda) As Long
    RG_Traz_Regiao_Tela = SUCESSO
End Function

Public Sub LabelRegiao_Click()

Dim objRegiao As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Codigo.Text)) > 0 Then objRegiao.iCodigo = StrParaInt(Codigo.Text)
        
    'Chama a tela que lista os vendedores
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiao, objEventoRegiao)

End Sub

Private Sub objEventoRegiao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRegiao As ClassRegiaoVenda
Dim iIndice As Integer

On Error GoTo Erro_objEventoRegiao_evSelecao

    Set objRegiao = obj1

    lErro = CF("RegiaoVenda_Le", objRegiao)
    If lErro <> SUCESSO And lErro <> 16137 Then gError ERRO_SEM_MENSAGEM
    
    lErro = Traz_Regiao_Tela(objRegiao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Me.Show

    Exit Sub

Erro_objEventoRegiao_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197176)

    End Select

    Exit Sub

End Sub
