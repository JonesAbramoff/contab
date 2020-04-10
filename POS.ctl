VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl POS 
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5100
   ScaleHeight     =   2985
   ScaleWidth      =   5100
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   2865
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "POS.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "POS.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "POS.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "POS.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Terminais 
      Height          =   1230
      Left            =   300
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1620
      Width           =   4635
   End
   Begin VB.ComboBox Rede 
      Height          =   315
      ItemData        =   "POS.ctx":0994
      Left            =   855
      List            =   "POS.ctx":0996
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   2715
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   855
      TabIndex        =   0
      Top             =   240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   10
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Terminais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   10
      Top             =   1350
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   285
      Width           =   660
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Rede:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   8
      Top             =   840
      Width           =   525
   End
End
Attribute VB_Name = "POS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'
''Property Variables:
'Dim m_Caption As String
'Event Unload()
'
''Variável Global
'Public iAlterado As Integer
'
'Public Sub Form_Load()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Form_Load
'
'    'Carrega a combo de terminais
'    lErro = Carrega_Terminais()
'    If lErro <> SUCESSO Then gError 80678
'
'    'Carrega a combo de redes
'    lErro = Carrega_Redes()
'    If lErro <> SUCESSO Then gError 80679
'
'    lErro_Chama_Tela = SUCESSO
'
'    Exit Sub
'
'Erro_Form_Load:
'
'    lErro_Chama_Tela = gErr
'
'    Select Case gErr
'
'        Case 80678, 80679
'            'Erro tratado na rotina chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165077)
'
'    End Select
'
'    iAlterado = 0
'
'    Exit Sub
'
'End Sub
'Private Function Carrega_Terminais() As Long
''Carrega a lista de terminais com o código em questão
'
'Dim lErro As Long
'Dim objPOS As ClassPOS
'Dim colPOS As New Collection
'
'On Error GoTo Erro_Carrega_Terminais
'
'    'Lê todos os terminas da FilialEmpresa
'    lErro = CF("POS_Le_Todos", colPOS)
'    If lErro <> SUCESSO Then gError 80674
'
'    'Carrega a listbox de terminais com código
'    For Each objPOS In colPOS
'        Terminais.AddItem objPOS.sCodigo
'    Next
'
'    Carrega_Terminais = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Terminais:
'
'    Carrega_Terminais = gErr
'
'    Select Case gErr
'
'        Case 80674
'         'Erro tratado na rotina chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165078)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Carrega_Redes() As Long
'
'Dim lErro As Long
'Dim objCodigoNome As AdmCodigoNome
'Dim colCodigoNome As New AdmColCodigoNome
'
'On Error GoTo Erro_Carrega_Redes
'
'    'Lê o Código e o Nome de Todas as Redes do BD
'    lErro = CF("Cod_Nomes_Le", "Redes", "Codigo", "Nome", STRING_REDE_NOME, colCodigoNome)
'    If lErro <> SUCESSO Then gError 80575
'
'    'Carrega a combo de Redes
'    For Each objCodigoNome In colCodigoNome
'        Rede.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
'        Rede.ItemData(Rede.NewIndex) = objCodigoNome.iCodigo
'    Next
'
'    Carrega_Redes = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Redes:
'
'    Carrega_Redes = gErr
'
'    Select Case gErr
'
'        Case 80575
'         'Erro tratado na rotina chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165079)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Trata_Parametros(Optional objPOS As ClassPOS) As Long
'
'Dim lErro As Long
'Dim bEncontrou As Boolean
'
'On Error GoTo Erro_Trata_Parametros
'
'    'Se houver POS passado como parâmetro, exibe seus dados
'    If Not (objPOS Is Nothing) Then
'
'        If Len(Trim(objPOS.sCodigo)) > 0 Then
'
'            objPOS.iFilialEmpresa = giFilialEmpresa
'
'            'Lê POS no BD a partir do código
'            lErro = CF("POS_Le", objPOS)
'            If lErro <> SUCESSO And lErro <> 79590 Then gError 80680
'            If lErro = SUCESSO And objPOS.iFilialEmpresa = giFilialEmpresa Then
'
'                'Exibe os dados do POS
'                lErro = Traz_POS_Tela(objPOS)
'                If lErro <> SUCESSO Then gError 80681
'
'            Else
'                Codigo.Text = objPOS.sCodigo
'
'            End If
'
'        End If
'
'    End If
'
'    iAlterado = 0
'
'    Exit Function
'
'    Trata_Parametros = SUCESSO
'
'Erro_Trata_Parametros:
'
'    Trata_Parametros = gErr
'
'    Select Case gErr
'
'        Case 80680, 80681
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165080)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
''Extrai os campos da tela que correspondem aos campos no BD
'
'Dim lErro As Long
'Dim objPOS As New ClassPOS
'
'On Error GoTo Erro_Tela_Extrai
'
'    'Informa tabela associada à Tela
'    sTabela = "POS"
'
'    'Le os dados da Tela AdmMeioPagto
'    lErro = Move_Tela_Memoria(objPOS)
'    If lErro <> SUCESSO Then Error 80676
'
'    'Preenche a coleção colCampoValor, com nome do campo,
'    'valor atual (com a tipagem do BD), tamanho do campo
'    'no BD no caso de STRING e Key igual ao nome do campo
'    colCampoValor.Add "Codigo", objPOS.sCodigo, STRING_POS_CODIGO, "Codigo"
'    colCampoValor.Add "Rede", objPOS.iRede, 0, "Rede"
'
'    'Filtros para o Sistema de Setas
'    colSelecao.Add "FilialEmpresa", OP_IGUAL, objPOS.iFilialEmpresa
'
'    Exit Sub
'
'Erro_Tela_Extrai:
'
'    Select Case gErr
'
'        Case 80676
'        'Erro tratado na rotina chamadora
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165081)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
''Preenche os campos da tela com os correspondentes do BD
'
'Dim lErro As Long
'Dim objPOS As New ClassPOS
'
'On Error GoTo Erro_Tela_Preenche
'
'    objPOS.sCodigo = colCampoValor.Item("Codigo").vValor
'
'    If Len(Trim(objPOS.sCodigo)) > 0 Then
'
'        'Carrega objPOS com os dados passados em colCampoValor
'        objPOS.iRede = colCampoValor.Item("Rede").vValor
'
'        'Traz dados do POS para a Tela
'        lErro = Traz_POS_Tela(objPOS)
'        If lErro <> SUCESSO Then Error 80677
'
'    End If
'
'    iAlterado = 0
'
'    Exit Sub
'
'Erro_Tela_Preenche:
'
'    Select Case gErr
'
'        Case 80677
'        'Erro tratado na rotina chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165082)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'Function Traz_POS_Tela(objPOS As ClassPOS) As Long
'
'Dim iIndice As Integer
'
'    Call Limpa_Tela_POS
'
'    'Traz os dados para tela
'    Codigo.Text = objPOS.sCodigo
'
'    'Procedimento usado quando a combox não é editavél.
'    For iIndice = 0 To Rede.ListCount - 1
'        If Rede.ItemData(iIndice) = objPOS.iRede Then
'            Rede.ListIndex = iIndice
'            Exit For
'        End If
'    Next
'
'    iAlterado = 0
'
'    Traz_POS_Tela = SUCESSO
'
'    Exit Function
'
'End Function
'
'Function Move_Tela_Memoria(objPOS As ClassPOS) As Long
' 'Move os dados da tela para o POS
'
'    objPOS.sCodigo = Codigo.Text
'    objPOS.iRede = Codigo_Extrai(Rede.Text)
'    objPOS.iFilialEmpresa = giFilialEmpresa
'
'    Exit Function
'
'End Function
'
'Private Sub Terminais_Inclui(objPOS As ClassPOS)
''Adiciona na ListBox informações dos terminais
'
'    Terminais.AddItem objPOS.sCodigo
'
'End Sub
'
'
'Private Sub Codigo_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Rede_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Rede_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Private Sub Terminais_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Terminais_DblClick()
''Carrega para a tela POS selecionado através de um duplo-clique
'
'Dim objPOS As New ClassPOS
'Dim lErro As Long
'
'On Error GoTo Erro_Terminais_DblClick
'
'    If Terminais.ListIndex >= 0 Then
'
'        objPOS.sCodigo = Terminais.List(Terminais.ListIndex)
'        objPOS.iFilialEmpresa = giFilialEmpresa
'
'        'Procura o POS no BD através do código
'        lErro = CF("POS_Le", objPOS)
'        If lErro <> SUCESSO And lErro <> 79590 Then Error 80690
'
'        'Se não encontrou
'        If lErro = 79590 Then gError 80691
'
'        'Traz para a tela os dados do POS selecionado
'        lErro = Traz_POS_Tela(objPOS)
'        If lErro <> SUCESSO Then gError 80692
'
'        'Fecha o comando das setas se estiver aberto
'        lErro = ComandoSeta_Fechar(Me.Name)
'
'    End If
'
'    Exit Sub
'
'Erro_Terminais_DblClick:
'
'    Select Case gErr
'
'        Case 80690, 80692
'            'Erro tratado na rotina chamadora
'
'        Case 80691
'            Call Rotina_Erro(vbOKOnly, "ERRO_POS_NAO_CADASTRADO", gErr, objPOS.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165083)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Rede_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim vbMsgRes As VbMsgBoxResult
'Dim objRede As New ClassRede
'Dim iCodigo As Integer
'
'On Error GoTo Erro_Rede_Validate
'
'    'Verifica se foi preenchida a ComboBox Rede
'    If Len(Trim(Rede.Text)) = 0 Then Exit Sub
'
'    'Verifica se está preenchida com o item selecionado na ComboBox Rede
'    If Rede.Text = Rede.List(Rede.ListIndex) Then Exit Sub
'
'    'Verifica se existe o item na List da Combo. Se existir seleciona.
'    lErro = Combo_Seleciona(Rede, iCodigo)
'    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 80699
'
'    'Nao existe o item com o CÓDIGO na List da ComboBox
'    If lErro = 6730 Then
'
'        objRede.iCodigo = iCodigo
'
'        'Tenta ler Rede com esse código no BD
'        lErro = CF("Rede_Le", objRede)
'        If lErro <> SUCESSO And lErro <> 80591 Then gError 80700
'
'        'Não encontrou Rede no BD
'        If lErro = 80591 Then gError 80701
'
'        'Encontrou Rede no BD, coloca no Text da Combo
'        Rede.Text = CStr(objRede.iCodigo) & SEPARADOR & objRede.sNome
'
'    End If
'
'    'Não existe o item com a STRING na List da ComboBox
'    If lErro = 6731 Then gError 80702
'
'    Exit Sub
'
'Erro_Rede_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 80699, 80700, 80591
'        'Erro tratado na rotina chamadora
'
'        Case 80701, 80702
'            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_ENCONTRADA", gErr, Rede.Text)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165084)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Terminais_Exclui(objPOS As ClassPOS)
''Percorre a ListBox de Terminais para remover a informação em questão
'
'Dim iIndice As Integer
'
'    For iIndice = 0 To Terminais.ListCount - 1
'
'        If SCodigo_Extrai(Terminais.List(iIndice)) = objPOS.sCodigo Then
'            Terminais.RemoveItem (iIndice)
'            Exit For
'        End If
'     Next
'
'End Sub
'
'Private Sub BotaoLimpar_Click()
''chamada de Limpa_Tela_POS
'Dim lErro As Long
'
'On Error GoTo Erro_Botaolimpar_Click
'
'    lErro = Teste_Salva(Me, iAlterado)
'    If lErro <> SUCESSO Then gError 80698
'
'    'Limpa Tela
'    Call Limpa_Tela_POS
'
'    'Fecha o comando das setas se estiver aberto
'    lErro = ComandoSeta_Fechar(Me.Name)
'
'    iAlterado = 0
'
'    Exit Sub
'
'Erro_Botaolimpar_Click:
'
'    Select Case gErr
'
'        Case 80698
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165085)
'
'    End Select
'
'End Sub
'Sub Limpa_Tela_POS()
'
'Dim lErro As Long
'
'    'Limpa Tela
'    Call Limpa_Tela(Me)
'
'    Rede.ListIndex = -1
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoGravar_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoGravar_Click
'
'    lErro = Gravar_Registro()
'    If lErro <> SUCESSO Then gError 80693
'
'    Call Limpa_Tela_POS
'
'    iAlterado = 0
'
'    Exit Sub
'
'Erro_BotaoGravar_Click:
'
'    Select Case gErr
'
'        Case 80693
'            'Erro tratado na rotina chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165086)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Function Gravar_Registro() As Long
'
'Dim lErro As Long
'Dim objPOS As New ClassPOS
'
'On Error GoTo Erro_Gravar_Registro
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'verifica preenchimento do codigo
'    If Len(Trim(Codigo.Text)) = 0 Then gError 80694
'
'    'verifica se rede foi informada
'    If Len(Trim(Rede.Text)) = 0 Then gError 80695
'
'    'preenche objPOS
'    lErro = Move_Tela_Memoria(objPOS)
'    If lErro <> AD_SQL_SUCESSO Then gError 80696
'
'    lErro = Trata_Alteracao(objPOS, objPOS.sCodigo)
'    If lErro <> SUCESSO Then Error 32329
'
'    lErro = CF("POS_Grava", objPOS)
'    If lErro <> SUCESSO Then gError 80697
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    'Exclui da ListBox
'    Call Terminais_Exclui(objPOS)
'
'    'Inclui na ListBox
'    Call Terminais_Inclui(objPOS)
'
'    Gravar_Registro = SUCESSO
'
'    Exit Function
'
'Erro_Gravar_Registro:
'
'    Gravar_Registro = gErr
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case gErr
'
'        Case 32329
'
'        Case 80694
'            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
'
'        Case 80695
'            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_PREENCHIDA", gErr)
'
'        Case 80696, 80697
'            'Erro tratado na rotina chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165087)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub BotaoExcluir_Click()
'
'Dim lErro As Long
'Dim objPOS As New ClassPOS
'Dim vbMsgRes As VbMsgBoxResult
'
' On Error GoTo Erro_BotaoExcluir_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'Verifica se o codigo foi preenchido
'    If Len(Trim(Codigo.ClipText)) = 0 Then gError 80713
'
'     objPOS.sCodigo = Codigo.Text
'     objPOS.iFilialEmpresa = giFilialEmpresa
'
'    'le o POS com codigo
'    lErro = CF("POS_Le", objPOS)
'    If lErro <> SUCESSO And lErro <> 79590 Then gError 80714
'
'    'POS não está cadastrado
'    If lErro = 79590 Then gError 80715
'    If objPOS.iFilialEmpresa <> giFilialEmpresa Then Error 80686
'
'    'Envia aviso perguntando se realmente deseja excluir pos
'    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_POS", objPOS.sCodigo)
'
'    If vbMsgRes = vbYes Then
'
'        'Exclui POS
'        lErro = CF("POS_Exclui", objPOS)
'        If lErro <> SUCESSO Then gError 80716
'
'        'Exclui da List
'        Call Terminais_Exclui(objPOS)
'
'        Call Limpa_Tela_POS
'
'        'Fecha o comando das setas se estiver aberto
'        lErro = ComandoSeta_Fechar(Me.Name)
'
'        iAlterado = 0
'
'    End If
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoExcluir_Click:
'
'    Select Case gErr
'
'        Case 80686
'            Call Rotina_Erro(vbOKOnly, "ERRO_POS_OUTRA_FILIALEMPRESA", Err, objPOS.sCodigo, objPOS.iFilialEmpresa)
'
'        Case 80713
'            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
'
'        Case 80716
'            'Erro tratado na rotina chamadora
'
'        Case 80714, 80715
'            Call Rotina_Erro(vbOKOnly, "ERRO_POS_NAO_CADASTRADO", gErr, objPOS.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165088)
'
'    End Select
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoFechar_Click()
'
'    Unload Me
'
'End Sub
'Public Sub Form_Activate()
'
'    Call TelaIndice_Preenche(Me)
'
'End Sub
'Public Sub Form_Deactivate()
'
'    gi_ST_SetaIgnoraClick = 1
'
'End Sub
'Public Sub Form_Unload(Cancel As Integer)
'
'    'Libera a referência da tela e fecha o comando das setas se estiver aberto
'    Call ComandoSeta_Liberar(Me.Name)
'
'End Sub
'Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
'
'    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
'
'End Sub
'
''**** inicio do trecho a ser copiado *****
'Public Function Form_Load_Ocx() As Object
'
'    '??? Parent.HelpContextID = IDH_
'    Set Form_Load_Ocx = Me
'    Caption = "POS"
'    Call Form_Load
'
'End Function
'
'Public Function Name() As String
'
'    Name = "POS"
'
'End Function
'
'Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
'End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Controls
'Public Property Get Controls() As Object
'    Set Controls = UserControl.Controls
'End Property
'
'Public Property Get hWnd() As Long
'    hWnd = UserControl.hWnd
'End Property
'
'Public Property Get Height() As Long
'    Height = UserControl.Height
'End Property
'
'Public Property Get Width() As Long
'    Width = UserControl.Width
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ActiveControl
'Public Property Get ActiveControl() As Object
'    Set ActiveControl = UserControl.ActiveControl
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = UserControl.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    UserControl.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'End Property
'
''Load property values from storage
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'
'    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
'End Sub
'
''Write property values to storage
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
'    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
'End Sub
'
'Private Sub Unload(objme As Object)
''Parent.UnloadDoFilho
'   RaiseEvent Unload
'End Sub
'
'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    Parent.Caption = New_Caption
'    m_Caption = New_Caption
'End Property
'
''***** fim do trecho a ser copiado ******
'
'
