VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl TipoEmbalagem 
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   KeyPreview      =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   6330
   Begin VB.ListBox ListTipos 
      Height          =   1815
      ItemData        =   "TipoEmbalagemGR.ctx":0000
      Left            =   195
      List            =   "TipoEmbalagemGR.ctx":0002
      TabIndex        =   3
      Top             =   1650
      Width           =   6000
   End
   Begin VB.TextBox TextDescricao 
      Height          =   345
      Left            =   1230
      MaxLength       =   50
      TabIndex        =   2
      Top             =   930
      Width           =   4995
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4005
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoEmbalagemGR.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "TipoEmbalagemGR.ctx":0182
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoEmbalagemGR.ctx":06B4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoEmbalagemGR.ctx":083E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   300
      Left            =   1770
      Picture         =   "TipoEmbalagemGR.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   300
      Width           =   300
   End
   Begin MSMask.MaskEdBox MaskCodigo 
      Height          =   315
      Left            =   1215
      TabIndex        =   0
      Top             =   285
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipos"
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
      Left            =   225
      TabIndex        =   11
      Top             =   1425
      Width           =   480
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
      Left            =   165
      TabIndex        =   10
      Top             =   960
      Width           =   945
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
      Left            =   435
      TabIndex        =   9
      Top             =   315
      Width           =   660
   End
End
Attribute VB_Name = "TipoEmbalagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private Sub MaskCodigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskCodigo, iAlterado)

End Sub

Private Sub MaskCodigo_Validate(Cancel As Boolean)
'Verifica se o código é válido

Dim lErro As Long

On Error GoTo Erro_MaskCodigo_Validate
    
    'Verifica se código foi informado
    If Len(MaskCodigo.Text) > 0 Then
    
        'Verifica se o código é um valor positivo
        lErro = Valor_Positivo_Critica(MaskCodigo.Text)
        If lErro <> AD_SQL_SUCESSO Then gError 96500
       
    End If

    Exit Sub

Erro_MaskCodigo_Validate:

    Cancel = True

    Select Case gErr

        Case 96500

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub MaskCodigo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TextDescricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colTipoEmbalagem As New Collection
Dim objTipoEmbalagem As ClassTipoEmbalagem

On Error GoTo Erro_TipoEmbalagem_Form_Load

    'Lê cada tipo e descrição da tabela TipoEmbalagem e poe na coleção colTipoEmbalagem
    lErro = CF("TipoEmbalagem_Le_Todos", colTipoEmbalagem)
    If lErro <> AD_SQL_SUCESSO Then gError 96501
    
    'Carrega na listTipos os Tipos de Embalagem existentes
    For Each objTipoEmbalagem In colTipoEmbalagem
                
        'Adiciona novo item na Listbox ListTipos
        ListTipos.AddItem objTipoEmbalagem.iTipo & SEPARADOR & objTipoEmbalagem.sDescricao
        ListTipos.ItemData(ListTipos.NewIndex) = objTipoEmbalagem.iTipo
    Next
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_TipoEmbalagem_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 96501

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTipoEmbalagem As ClassTipoEmbalagem) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Tipo de Embalagem selecionado, exibir seus dados
    If Not (objTipoEmbalagem Is Nothing) Then

        'Verifica se o TipoEmbalagem existe
        lErro = Traz_TipoEmbalagem_Tela(objTipoEmbalagem)
        If lErro <> AD_SQL_SUCESSO And lErro <> 96548 Then gError 96508
        
        'Se não existe
        If lErro = 96548 Then
        
            'Limpa a Tela
            Call Limpa_Tela(Me)

            'Joga o código na tela
            MaskCodigo.Text = CStr(objTipoEmbalagem.iTipo)

        End If
    
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 96508

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Function Move_Tela_Memoria(objTipoEmbalagem As ClassTipoEmbalagem) As Long
'Move os campos da tela para o objTipoEmbalagem
      
    objTipoEmbalagem.iTipo = StrParaInt(MaskCodigo.Text)
    objTipoEmbalagem.sDescricao = TextDescricao.Text
    
   Move_Tela_Memoria = SUCESSO
   
End Function

Private Function Traz_TipoEmbalagem_Tela(objTipoEmbalagem As ClassTipoEmbalagem) As Long
'Coloca os dados do código passado como parâmetro na tela
Dim lErro As Long

On Error GoTo Erro_Traz_TipoEmbalagem_Tela

    lErro = CF("TipoEmbalagem_Le", objTipoEmbalagem)
    If lErro <> AD_SQL_SUCESSO And lErro <> 96507 Then gError 96509
    
    If lErro = 96507 Then gError 96548
    
    'TipoEmbalagem está cadastrado
    MaskCodigo.Text = CStr(objTipoEmbalagem.iTipo)
    TextDescricao.Text = objTipoEmbalagem.sDescricao
            
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Traz_TipoEmbalagem_Tela = SUCESSO
    
    Exit Function

Erro_Traz_TipoEmbalagem_Tela:

    Traz_TipoEmbalagem_Tela = gErr

    Select Case gErr
        
        Case 96509, 96548

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Private Sub BotaoProxNum_Click()
'Coloca o próximo número a ser gerado na tela

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático.
    lErro = TipoEmbalagem_Codigo_Automatico(iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96510

    'Joga o código na tela
    MaskCodigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 96510

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
        
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Function TipoEmbalagem_Codigo_Automatico(iCodigo As Integer) As Long
'Retorna o proximo número disponivel

Dim lErro As Long

On Error GoTo Erro_TipoEmbalagem_Codigo_Automatico

    'Gera número automático.
    lErro = CF("Config_Obter_Inteiro_Automatico", "FatConfig", "NUM_PROX_TIPO_EMBALAGEM", "TipoEmbalagem", "Tipo", iCodigo)
    If lErro <> SUCESSO Then gError 96511
    
    TipoEmbalagem_Codigo_Automatico = SUCESSO

    Exit Function

Erro_TipoEmbalagem_Codigo_Automatico:

    TipoEmbalagem_Codigo_Automatico = gErr

    Select Case gErr

        Case 96511
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click
    
    'Verifica se existe algo para ser salvo antes de sair
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> AD_SQL_SUCESSO Then gError 96512
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoFechar_Click:
    
    Select Case gErr
        
        Case 96512
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
        
    Exit Sub
        
End Sub

Private Sub ListTipos_DblClick()
'Traz para a tela os dados do listTipos selecionado

Dim lErro As Long
Dim objTipoEmbalagem As New ClassTipoEmbalagem

On Error GoTo Erro_ListTipos_DblClick

    objTipoEmbalagem.iTipo = ListTipos.ItemData(ListTipos.ListIndex)

    'Verifica se o TipoEmbalagem existe
    lErro = Traz_TipoEmbalagem_Tela(objTipoEmbalagem)
    If lErro <> AD_SQL_SUCESSO And lErro <> 96548 Then gError 96513
    
    'Se não encontrou --> erro
    If lErro = 96548 Then
    
        'Se Tipo Embalagem não está cadastrado, exclui da ListTipos
        ListTipos.RemoveItem (ListTipos.ListIndex)
        gError 92816
    
    End If
    
    iAlterado = 0

    Exit Sub

Erro_ListTipos_DblClick:

    Select Case gErr

        Case 92816
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOEMBALAGEM_NAO_CADASTRADO", gErr, objTipoEmbalagem.iTipo)

        Case 96513

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ListTipos_KeyPress(KeyAscii As Integer)

    'Se há um Tipo de Embalagem selecionado na ListTipos
    If ListTipos.ListIndex <> -1 Then

        If KeyAscii = ENTER_KEY Then

            Call ListTipos_DblClick

        End If

    End If

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Controla toda a rotina de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 96514

    'Limpa a Tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 96514

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Controla toda a rotina de gravação

Dim lErro As Long
Dim objTipoEmbalagem As New ClassTipoEmbalagem
Dim iCodigo As Integer

On Error GoTo Erro_Gravar_Registro
    
    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios foram informados
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 96515
    
    If Len(Trim(TextDescricao.Text)) = 0 Then gError 96516
    
    'Move os campos da tela para o objTipoEmbalagem
    lErro = Move_Tela_Memoria(objTipoEmbalagem)
    If lErro <> AD_SQL_SUCESSO Then gError 96517
    
    'Verifica se o Tipo de Embalagem já existe, se existir manda uma mensagem
    lErro = Trata_Alteracao(objTipoEmbalagem, objTipoEmbalagem.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 96518

    'Grava o Tipo de Embalagem no banco de dados
    lErro = CF("TipoEmbalagem_Grava", objTipoEmbalagem)
    If lErro <> AD_SQL_SUCESSO Then gError 96519

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iCodigo = objTipoEmbalagem.iTipo
    
    'Remove o item do listTipos
    Call ListTipos_Exclui(iCodigo)

    'Insere o item no listTipos
    Call ListTipos_Adicionar(objTipoEmbalagem)
    
    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 96515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)
            
        Case 96516
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", gErr)
                    
        Case 96517, 96518, 96519

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click
    
    'Verifica se existe algo para ser salvo antes de limpar a tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> AD_SQL_SUCESSO Then gError 96526

    'Limpa a Tela
    Call Limpa_Tela(Me)

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 96526

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Sub

End Sub

Private Sub ListTipos_Adicionar(objTipoEmbalagem As ClassTipoEmbalagem)

Dim sListBoxItem As String
Dim iPos As Integer
Dim iIndice As Integer

    'Concatena o código com a descrição do Tipo Embalagem
    sListBoxItem = CStr(objTipoEmbalagem.iTipo) & SEPARADOR & objTipoEmbalagem.sDescricao
    
    For iIndice = 0 To ListTipos.ListCount - 1
        
        'Se o campo selecionado do listtipos for menor que o do código passado...
        If ListTipos.ItemData(iIndice) < objTipoEmbalagem.iTipo Then
        
            'quarda a posição do próximo a ser incluido
            iPos = iIndice + 1
            
        End If
        
    Next
    
    'Caso o campo a ser incluido na listtipos seja o primeiro
    If ListTipos.ListCount = 0 Then iPos = 0
    
    'Adiciona um item na listTipos
    ListTipos.AddItem sListBoxItem, iPos
    ListTipos.ItemData(ListTipos.NewIndex) = objTipoEmbalagem.iTipo

End Sub

Private Sub ListTipos_Exclui(ByVal iCodigo As Integer)

Dim iIndice As Integer

    For iIndice = 0 To ListTipos.ListCount - 1

        If ListTipos.ItemData(iIndice) = iCodigo Then

            ListTipos.RemoveItem (iIndice)
            Exit For

        End If

    Next

End Sub

Private Sub BotaoExcluir_Click()
'Exclui o Tipo de Embalagem do código passado

Dim lErro As Long
Dim objTipoEmbalagem As New ClassTipoEmbalagem
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click
    
    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Código foi informado
    If Len(Trim(MaskCodigo.ClipText)) = 0 Then gError 96527

    objTipoEmbalagem.iTipo = CInt(MaskCodigo.Text)

    'Verifica se o TipoEmbalagem existe
    lErro = CF("TipoEmbalagem_Le", objTipoEmbalagem)
    If lErro <> SUCESSO And lErro <> 96507 Then gError 96528

    'TipoEmbalagem não está cadastrado
    If lErro = 96507 Then gError 96529

    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TIPOEMBALAGEM", objTipoEmbalagem.iTipo)

    If vbMsgRet = vbYes Then

        'exclui o Tipo Embalagem
        lErro = CF("TipoEmbalagem_Exclui", objTipoEmbalagem)
        If lErro <> SUCESSO Then gError 96530

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
        
        'Exclui o Tipo Embalagem da ListBox
        Call ListTipos_Exclui(objTipoEmbalagem.iTipo)

        'Limpa a Tela
        Call Limpa_Tela(Me)

        iAlterado = 0

    End If
    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:
    
    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 96527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 96528, 96530

        Case 96529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objTipoEmbalagem.iTipo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objTipoEmbalagem As New ClassTipoEmbalagem
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche
    
    'Preenche com o tipo retornado
    objTipoEmbalagem.iTipo = colCampoValor.Item("Tipo").vValor
    
    'Se tiver tipo informado
    If objTipoEmbalagem.iTipo <> 0 Then

        'Traz dados da Administradora para a Tela
        lErro = Traz_TipoEmbalagem_Tela(objTipoEmbalagem)
        If lErro <> SUCESSO And lErro <> 96548 Then gError 96543
        
        'se não tiver cadastrado --> erro.
        If lErro = 96548 Then gError 92817

        iAlterado = 0

    End If
      
    Exit Sub
    
Erro_Tela_Preenche:
       
    Select Case gErr

        Case 92817
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOEMBALAGEM_NAO_CADASTRADO", gErr, objTipoEmbalagem.iTipo)

        Case 96543

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Sub
    
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objTipoEmbalagem As New ClassTipoEmbalagem
Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TipoEmbalagem"
    
    'Le os dados da Tela TipoEmbalagem
    lErro = Move_Tela_Memoria(objTipoEmbalagem)
    If lErro <> SUCESSO Then gError 96544
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Tipo", objTipoEmbalagem.iTipo, 0, "Tipo"
    colCampoValor.Add "Descricao", objTipoEmbalagem.sDescricao, STRING_TIPOEMBALAGEM_DESCRICAO, "Descricao"
          
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr

        Case 96544

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

''**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_HISTORICO_PADRAO
    Set Form_Load_Ocx = Me
    Caption = "Tipo de Embalagem"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TipoEmbalagem"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
End Sub

'***** fim do trecho a ser copiado ******

