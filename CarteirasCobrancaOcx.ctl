VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CarteirasCobrancaOcx 
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   5565
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1845
      Picture         =   "CarteirasCobrancaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   360
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3255
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CarteirasCobrancaOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CarteirasCobrancaOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CarteirasCobrancaOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CarteirasCobrancaOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox Descricao 
      Height          =   300
      Left            =   1155
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   802
      Width           =   4245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Válida para"
      Height          =   660
      Left            =   135
      TabIndex        =   13
      Top             =   1230
      Width           =   5280
      Begin VB.OptionButton OpcaoEmpresa 
         Caption         =   "Própria Empresa"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Top             =   270
         Width           =   1710
      End
      Begin VB.OptionButton OpcaoOutros 
         Caption         =   "Outros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1335
         TabIndex        =   4
         Top             =   292
         Width           =   915
      End
      Begin VB.OptionButton OpcaoBanco 
         Caption         =   "Bancos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton OpcaoAmbos 
         Caption         =   "Ambos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2415
         TabIndex        =   5
         Top             =   285
         Width           =   915
      End
   End
   Begin VB.ListBox CarteirasCobranca 
      Height          =   1620
      ItemData        =   "CarteirasCobrancaOcx.ctx":0A7E
      Left            =   135
      List            =   "CarteirasCobrancaOcx.ctx":0A80
      TabIndex        =   7
      Top             =   2295
      Width           =   5265
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1155
      TabIndex        =   0
      Top             =   345
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
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
      Left            =   405
      TabIndex        =   14
      Top             =   405
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Left            =   135
      TabIndex        =   15
      Top             =   855
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Carteiras de Cobrança"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   135
      TabIndex        =   16
      Top             =   2070
      Width           =   1935
   End
End
Attribute VB_Name = "CarteirasCobrancaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera Código de Carteira automático.
    lErro = CF("CarteirasCobranca_Automatico",iCodigo)
    If lErro <> SUCESSO Then Error 57544

    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57544 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144239)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim vbMsgRet As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Código da Carteira foi informado
    If Len(Codigo.Text) = 0 Then Error 23430

    objCarteiraCobranca.iCodigo = CInt(Codigo.Text)

    'Verifica se o Código da carteira existe
    lErro = CF("CarteiraDeCobranca_Le",objCarteiraCobranca)
    If lErro <> SUCESSO And lErro <> 23413 Then Error 23431

    'Carteira não está cadastrado
    If lErro = 23413 Then Error 23432

    'Confirma a exclusão
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CARTEIRACOBRANCA")

    If vbMsgRet = vbYes Then

        'Exclui a Carteira
        lErro = CF("CarteiraDeCobranca_Exclui",objCarteiraCobranca.iCodigo)
        If lErro <> SUCESSO Then Error 23433

        'Exclui a Carteira da ListBox
        Call ListaCarteiraDeCobranca_Exclui(Codigo.Text)
        
        'Limpa a Tela
        Call Limpa_Tela_CarteirasCobranca

    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 23430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGOCARTCOBR_NAO_PREENCHIDO", Err, Error$)
            Codigo.SetFocus

        Case 23432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", Err, objCarteiraCobranca.iCodigo)

        Case 23431, 23433 'Tratados nas Rotinas chamadaa

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 144240)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Grava CarteirasCobranca
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 23426
    
    'Limpa a Tela
    Call Limpa_Tela_CarteirasCobranca
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 23426 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144241)

    End Select

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se houve alteração e confirma se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 23428

    'Limpa a Tela
    Call Limpa_Tela_CarteirasCobranca

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 23428 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144242)

     End Select

     Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 57997
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True
    
    Select Case Err
        
        Case 57997 'Erro tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144243)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCarteirasCobranca As New Collection
Dim sListBoxItem As String
Dim objCarteiraCobranca As ClassCarteiraCobranca

On Error GoTo Erro_CarteirasCobranca_Form_Load

    'Coloca todos as carteiras na coleção
    lErro = CF("CarteirasDeCobranca_Le_Todas",colCarteirasCobranca)
    If lErro <> SUCESSO Then Error 23402

    'Preenche a ListBox com as Carteiras existentes na coleção
    For Each objCarteiraCobranca In colCarteirasCobranca

        'Concatena Código e Descrição da carteira
        sListBoxItem = CStr(objCarteiraCobranca.iCodigo)
        sListBoxItem = sListBoxItem & SEPARADOR & Trim(objCarteiraCobranca.sDescricao)

        CarteirasCobranca.AddItem sListBoxItem
        CarteirasCobranca.ItemData(CarteirasCobranca.NewIndex) = objCarteiraCobranca.iCodigo
        
    Next
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_CarteirasCobranca_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 23402 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144244)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objCarteiraCobranca As ClassCarteiraCobranca) As Long

Dim lErro As Long
Dim sListBoxItem As String
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há uma carteira selecionada
    If Not (objCarteiraCobranca Is Nothing) Then

        'Verifica se a Carteira existe
        lErro = CF("CarteiraDeCobranca_Le",objCarteiraCobranca)
        If lErro <> 15011 And lErro <> SUCESSO Then Error 23408

        'Se carteira está cadastrada
        If lErro = SUCESSO Then

            Call Traz_Carteira_Tela(objCarteiraCobranca)
            
        'Se a carteira não está cadastrada
        Else

            'Mantém o Código da carteira na tela
            Codigo.Text = CStr(objCarteiraCobranca.iCodigo)

        End If

    End If

    'Zerar iAlterado
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 23408 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144245)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
 
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub CarteirasCobranca_DblClick()

Dim lErro As Long
Dim iValidaPara As Integer
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim sListBoxItem As String
Dim lSeparadorPosicao As Long

On Error GoTo Erro_CarteirasCobranca_DblClick
    
    'Verifica se há algum ítem da listbox selecionado
    If CarteirasCobranca.ListIndex = -1 Then Exit Sub
      
    'Pega a String do ítem selecionado
    sListBoxItem = CarteirasCobranca.List(CarteirasCobranca.ListIndex)
         
    objCarteiraCobranca.iCodigo = Codigo_Extrai(sListBoxItem)
            
    'Busca no BD as informações sobre a Carteira
    lErro = CF("CarteiraDeCobranca_Le",objCarteiraCobranca)
    If lErro <> SUCESSO Then Error 23456
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Traz_Carteira_Tela(objCarteiraCobranca)
    
    Exit Sub
    
Erro_CarteirasCobranca_DblClick:

    Select Case Err
    
        Case 23456 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144246)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub CarteirasCobranca_KeyPress(KeyAscii As Integer)

    'Se não tiver nenhum ítem selecionado na lista
    If CarteirasCobranca.ListIndex = -1 Then Exit Sub

    'Se a tecla pressionada for Enter
    If KeyAscii = ENTER_KEY Then

        'Executa o mesmo procedimento que o duplo click
        Call CarteirasCobranca_DblClick

    End If

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Function Gravar_Registro() As Long
'Grava Carteiras

Dim lErro As Long
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim iCodigo As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se dados da carteira foram informados
    If Len(Codigo.Text) = 0 Then Error 23392
    If Len(Trim(Descricao.Text)) = 0 Then Error 23393

    If OpcaoEmpresa.Value = True Then Error 59253
        
    'Preenche objeto objCarteiraCobranca
    lErro = Move_Tela_Memoria(objCarteiraCobranca)
    If lErro <> SUCESSO Then Error 19412
    
    lErro = Trata_Alteracao(objCarteiraCobranca, objCarteiraCobranca.iCodigo)
    If lErro <> SUCESSO Then Error 32287
    
    'Grava a Carteira no Banco de Dados
    lErro = CF("CarteiraDeCobranca_Grava",objCarteiraCobranca)
    If lErro <> SUCESSO Then Error 23407

    'Remove o ítem da lista de carteiras, se já existir
    Call ListaCarteiraDeCobranca_Exclui(objCarteiraCobranca.iCodigo)

    'Insere o ítem na lista de carteiras
    Call ListaCarteirasCobranca_Adiciona(objCarteiraCobranca)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 23392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBRANCA_NAO_INFORMADA", Err, Error$)
            Codigo.SetFocus

        Case 23393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", Err)
            Descricao.SetFocus

        Case 23407 'Tratado na Rotina chamada

        Case 32287

        Case 59253
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_CARTEIRA_EMPRESA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144247)

     End Select

     Exit Function

End Function

Private Sub ListaCarteirasCobranca_Adiciona(objCarteiraCobranca As ClassCarteiraCobranca)
'Adiciona ítem na ListBox carteirascobranca

Dim iIndice As Integer
    
    For iIndice = 0 To CarteirasCobranca.ListCount - 1

        If CarteirasCobranca.ItemData(iIndice) > objCarteiraCobranca.iCodigo Then Exit For
        
    Next

    CarteirasCobranca.AddItem objCarteiraCobranca.iCodigo & SEPARADOR & objCarteiraCobranca.sDescricao, iIndice
    CarteirasCobranca.ItemData(iIndice) = objCarteiraCobranca.iCodigo
    
End Sub

Private Sub ListaCarteiraDeCobranca_Exclui(iCodigo As Integer)
'Exclui ítem da ListBox carteirascobranca

Dim iIndice As Integer

    'Percorre todos os itens da ListBox
    For iIndice = 0 To CarteirasCobranca.ListCount - 1

        'Se o ItemData do ítem for igual ao Código passado em iCodigo
        If CarteirasCobranca.ItemData(iIndice) = iCodigo Then

            'Remove o ítem
            CarteirasCobranca.RemoveItem (iIndice)
            Exit For

        End If

    Next

End Sub

Private Sub OpcaoAmbos_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OpcaoBANCO_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OpcaoOutros_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Function Traz_Carteira_Tela(objCarteiraCobranca As ClassCarteiraCobranca) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Carteira_Tela

    'Exibe os dados de objCarteira na tela
    If objCarteiraCobranca.iCodigo = 0 Then
        Codigo.Text = ""
    Else
        Codigo.Text = CStr(objCarteiraCobranca.iCodigo)
    End If
    
    Descricao.Text = objCarteiraCobranca.sDescricao
        
    If objCarteiraCobranca.iValidaPara = CARTCOBR_PARA_BANCOS Then
        OpcaoBanco.Value = True
        
    ElseIf objCarteiraCobranca.iValidaPara = CARTCOBR_PARA_OUTROS Then
        OpcaoOutros.Value = True
        
    ElseIf objCarteiraCobranca.iValidaPara = CARTCOBR_PARA_AMBOS Then
        OpcaoAmbos.Value = True
        
    Else
        OpcaoEmpresa.Value = True
    End If
    
    iAlterado = 0

    Traz_Carteira_Tela = SUCESSO

    Exit Function

Erro_Traz_Carteira_Tela:

    Traz_Carteira_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144248)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objCarteiraCobranca As ClassCarteiraCobranca) As Long
'Lê os dados que estão na tela CarteirasCobranca e coloca em objCarteiraCobranca

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'IDENTIFICACAO :
    If Len(Trim(Codigo.Text)) > 0 Then objCarteiraCobranca.iCodigo = CInt(Codigo.Text)
    objCarteiraCobranca.sDescricao = Trim(Descricao.Text)
    
    If OpcaoBanco.Value Then
        objCarteiraCobranca.iValidaPara = CARTCOBR_PARA_BANCOS
        
    ElseIf OpcaoOutros.Value Then
        objCarteiraCobranca.iValidaPara = CARTCOBR_PARA_OUTROS
        
    Else
        objCarteiraCobranca.iValidaPara = CARTCOBR_PARA_AMBOS

    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:
    
    Move_Tela_Memoria = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144249)
            
    End Select
    
    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objCarteiraCobranca As New ClassCarteiraCobranca

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CarteirasCobranca"

    'Le os dados da Tela
    lErro = Move_Tela_Memoria(objCarteiraCobranca)
    If lErro <> SUCESSO Then Error 23457

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objCarteiraCobranca.iCodigo, 0, "Codigo"
    colCampoValor.Add "ValidaPara", objCarteiraCobranca.iValidaPara, 0, "ValidaPara"
    colCampoValor.Add "Descricao", objCarteiraCobranca.sDescricao, STRING_DESCRICAO_CARTCOBR, "Descricao"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 23457 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144250)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCarteiraCobranca As New ClassCarteiraCobranca

On Error GoTo Erro_Tela_Preenche

    objCarteiraCobranca.iCodigo = colCampoValor.Item("Codigo").vValor

    If objCarteiraCobranca.iCodigo > 0 Then

        'Carrega objCarteiraCobranca com os dados passados em colCampoValor
        objCarteiraCobranca.sDescricao = colCampoValor.Item("Descricao").vValor
        objCarteiraCobranca.iValidaPara = colCampoValor.Item("ValidaPara").vValor
        
        'Traz dados da Carteira para a Tela
        lErro = Traz_Carteira_Tela(objCarteiraCobranca)
        If lErro <> SUCESSO Then Error 23458

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 23458 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144251)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_CarteirasCobranca()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CarteirasCobranca

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa a Tela
    Call Limpa_Tela(Me)

    Codigo.Text = ""
    
    OpcaoBanco.Value = True
    
    iAlterado = 0

    Exit Sub
    
Erro_Limpa_Tela_CarteirasCobranca:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144252)

    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CARTEIRAS_COBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Carteiras de Cobrança"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CarteirasCobranca"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
End Sub







Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

