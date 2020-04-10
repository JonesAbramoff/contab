VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TipoMovEstoque 
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   LockControls    =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   8550
   Begin VB.CheckBox Ativo 
      Caption         =   "Ativo"
      Height          =   420
      Left            =   2130
      TabIndex        =   1
      Top             =   510
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6210
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoMovEstoque.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoMovEstoque.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TipoMovEstoque.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoMovEstoque.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CheckBox AtualizaConsumo 
      Caption         =   "Atualiza Totalizadores de Consumo/Vendas"
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
      Left            =   165
      TabIndex        =   4
      Top             =   2370
      Width           =   4065
   End
   Begin VB.ListBox ListaTipos 
      Height          =   3180
      ItemData        =   "TipoMovEstoque.ctx":0994
      Left            =   5340
      List            =   "TipoMovEstoque.ctx":0996
      TabIndex        =   7
      Top             =   1125
      Width           =   3045
   End
   Begin VB.ComboBox EntradaOuSaida 
      Height          =   315
      ItemData        =   "TipoMovEstoque.ctx":0998
      Left            =   1095
      List            =   "TipoMovEstoque.ctx":09A2
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1740
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Apropriação"
      Height          =   1635
      Left            =   150
      TabIndex        =   13
      Top             =   2865
      Width           =   4695
      Begin VB.ComboBox ApropriacaoComprado 
         Height          =   315
         ItemData        =   "TipoMovEstoque.ctx":09B6
         Left            =   1920
         List            =   "TipoMovEstoque.ctx":09C3
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   2445
      End
      Begin VB.ComboBox ApropriacaoProduzido 
         Height          =   315
         ItemData        =   "TipoMovEstoque.ctx":0A02
         Left            =   1905
         List            =   "TipoMovEstoque.ctx":0A0F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   945
         Width           =   2445
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Material Comprado:"
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
         Left            =   210
         TabIndex        =   15
         Top             =   390
         Width           =   1650
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Material Produzido:"
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
         Left            =   195
         TabIndex        =   14
         Top             =   1005
         Width           =   1650
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1095
      TabIndex        =   0
      Top             =   555
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1110
      TabIndex        =   2
      Top             =   1155
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin VB.Label TipoLabel 
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
      Left            =   390
      TabIndex        =   19
      Top             =   600
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipos de Movimento de Estoque"
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
      Left            =   5310
      TabIndex        =   18
      Top             =   900
      Width           =   2745
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
      Left            =   120
      TabIndex        =   17
      Top             =   1215
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
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
      Left            =   600
      TabIndex        =   16
      Top             =   1785
      Width           =   450
   End
End
Attribute VB_Name = "TipoMovEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''Falta revisao Jones
''
''Option Explicit
''
'''Property Variables:
''Dim m_Caption As String
''Event Unload()
''
''Const TIPO_ENTRADA = 0
''Const TIPO_SAIDA = 1
''
''Dim iAlterado As Integer
''
''
''Private Sub ApropriacaoPadrao_Change()
''
''
''End Sub
''
''Private Sub Ativo_Click()
''
''    iAlterado = 1
''
''End Sub
''
''Private Sub AtualizaConsumo_Click()
''
''    iAlterado = 1
''
''End Sub
''
''Private Sub BotaoExcluir_Click()
''
''Dim lErro As Long
''Dim iIndice As Integer
''Dim objTipoMovEstoque As New ClassTipoMovEst
''
''On Error GoTo Erro_BotaoExcluir_Click
''
''    'Se código não estiver preenchido
''    If Len(Codigo.Text) = 0 Then Error 21868
''
''    'Se não for um campo editável pelo usuário
''''''    If NaoEditavel.Value = True Then Error 21869
''
''    objTipoMovEstoque.iCodigo = CInt(Codigo.Text)
''
''    'Exclui o código
''    lErro = TiposMovEst_Exclui(objTipoMovEstoque)
''    If lErro <> SUCESSO And lErro <> 21873 And lErro <> 21876 Then Error 21878
''
''    'Se o tipo não pode ser excluido
''    If lErro = 21873 Then Error 21879
''
''    'Se não encontrou o tipo selecionado
''    If lErro = 21876 Then Error 21880
''
''    'Exclui da ListBox
''    For iIndice = 0 To ListaTipos.ListCount
''
''        'Se encontrou
''        If ListaTipos.ItemData(iIndice) = objTipoMovEstoque.iCodigo Then
''
''            ListaTipos.RemoveItem (iIndice)
''            Exit For
''
''        End If
''
''    Next
''
''    lErro = Limpa_Tela_TipoMovEstoque
''    If lErro <> SUCESSO Then Error 21921
''
''    Exit Sub
''
''Erro_BotaoExcluir_Click:
''
''    Select Case Err
''
''        Case 21868
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)
''
''        Case 21869
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOSMOVEST_NAOEDITAVEL", Err)
''
''        Case 21878, 21921
''
''        Case 21879
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOSMOVEST1", CInt(Codigo.Text))
''
''        Case 21880
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOSMOVEST_INEXISTENTE", Err, CInt(Codigo.Text))
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174885)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Private Sub BotaoFechar_Click()
''
''    Unload Me
''
''End Sub
''
''Private Sub BotaoGravar_Click()
''Dim lErro As Long
''
''On Error GoTo Erro_BotaoGravar_Click
''
''    lErro = Gravar_Registro
''    If lErro <> SUCESSO Then Error 21881
''
''    Exit Sub
''
''Erro_BotaoGravar_Click:
''
''    Select Case Err
''
''        Case 21881
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174886)
''
''    End Select
''
''    Exit Sub
''
''
''End Sub
''
''Private Sub BotaoLimpar_Click()
''
''Dim lErro As Long
''
''On Error GoTo Erro_BotaoLimpar_Click
''
''    lErro = Teste_Salva(Me, iAlterado)
''    If lErro <> SUCESSO Then Error 21894
''
''    'Limpa a tela
''    lErro = Limpa_Tela_TipoMovEstoque
''    If lErro <> SUCESSO Then Error 21901
''
''    Exit Sub
''
''Erro_BotaoLimpar_Click:
''
''    Select Case Err
''
''        Case 21894, 21901
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174887)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Private Sub Codigo_Validate(Cancel As Boolean)
''
''Dim lErro As Long
''Dim objTipoMovEstoque As New ClassTipoMovEst
''
''On Error GoTo Erro_Codigo_Validate
''
''    'Critica se é valor válido
''    lErro = Valor_NaoNegativo_Critica(Codigo.Text)
''    If lErro <> SUCESSO Then Error 21902
''
''    objTipoMovEstoque.iCodigo = CInt(Codigo.Text)
''
''    'Traz os dados para a tela, se encontrar
''    lErro = Traz_Tela_TipoMovEstoque(objTipoMovEstoque)
''    If lErro <> SUCESSO And lErro <> 21908 Then Error 21909
''
''    iAlterado = 0
''
''    Exit Sub
''
''Erro_Codigo_Validate:

''    Cancel = True

''
''    Select Case Err
''
''        Case 21902, 21908, 21909
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174888)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Private Sub Descricao_Change()
''
''    iAlterado = 1
''
''End Sub
''
''Private Sub EntradaOuSaida_Click()
''
''    'Uma saida tem por default a apropriacao por custo medio e uma entrada por custo informado
''    If EntradaOuSaida.ListIndex = TIPO_ENTRADA Then
''''''        ApropriacaoPadrao.ListIndex = TIPO_INFORMADO_USUARIO
''    Else
''        If EntradaOuSaida.ListIndex = TIPO_SAIDA Then
''''''            ApropriacaoPadrao.ListIndex = TIPO_CUSTO_MEDIO
''        End If
''    End If
''
''    iAlterado = 1
''
''End Sub
''
''Public Sub Form_Activate()
''
''    Call TelaIndice_Preenche(Me)
''
''End Sub
''
''Public Sub Form_Load()
''
''Dim colCodigoDescicao As New AdmColCodigoNome
''Dim objCodigoNome As New AdmCodigoNome
''Dim sItem As String
''Dim lErro As Long
''Dim iCodigo As Integer
''
''On Error GoTo Erro_Form_Load
''
''    'Le a lista de tipos já existentes
''    lErro = CF("Cod_Nomes_Le","TiposMovimentoEstoque", "Codigo", "Descricao", STRING_TIPOMOV_EST_DESCRICAO, colCodigoDescicao)
''    If lErro <> SUCESSO Then Error 21910
''
''    'Preenche a ListBox com os tipos existentes no BD
''    For Each objCodigoNome In colCodigoDescicao
''
''        sItem = CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
''        ListaTipos.AddItem sItem
''        ListaTipos.ItemData(ListaTipos.NewIndex) = objCodigoNome.iCodigo
''
''    Next
''
''    EntradaOuSaida.ListIndex = TIPO_ENTRADA
''
''
''''''    NaoEditavel.Value = False
''
''    'Gera o próximo código a ser utilizado
''    lErro = TiposMovEst_Automatico(iCodigo)
''    If lErro <> SUCESSO Then Error 21911
''
''    Codigo.Text = CStr(iCodigo)
''
''    iAlterado = 0
''
''    lErro_Chama_Tela = SUCESSO
''
''    Exit Sub
''
''Erro_Form_Load:
''
''    lErro_Chama_Tela = Err
''
''    Select Case Err
''
''        Case 21910, 21911
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174889)
''
''    End Select
''
''    iAlterado = 0
''
''    Exit Sub
''
''End Sub
''
''Private Sub ListaTipos_Click()
''
''Dim lErro As Long
''Dim objTipoMovEstoque As New ClassTipoMovEst
''
''On Error GoTo Erro_ListaTipos_Click
''
''    'Recebe o codigo a ser pesquisado
''    objTipoMovEstoque.iCodigo = ListaTipos.ItemData(ListaTipos.ListIndex)
''
''    'Traz os dados para a tela
''    lErro = Traz_Tela_TipoMovEstoque(objTipoMovEstoque)
''    If lErro <> SUCESSO And lErro <> 21908 Then Error 21912
''
''    'Não encontrou o codigo pesquisado
''    If lErro = 21908 Then Error 21913
''
''    iAlterado = 0
''
''    Exit Sub
''
''Erro_ListaTipos_Click:
''
''    Select Case Err
''
''        Case 21912
''
''        Case 21913
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOSMOVEST_INEXISTENTE", Err, objTipoMovEstoque.iCodigo)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174890)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Public Sub Form_Deactivate()
''
''    gi_ST_SetaIgnoraClick = 1
''
''End Sub
''
''Function Limpa_Tela_TipoMovEstoque() As Long
'''Limpa a Tela TipoMovEstoque
''
''Dim lErro As Long
''Dim iCodigo As Integer
''
''On Error GoTo Erro_Limpa_Tela_TipoMovEstoque
''
''    'Fecha o comando das setas se estiver aberto
''    lErro = ComandoSeta_Fechar(Me.Name)
''    If lErro <> SUCESSO Then Error 21895
''
''    Call Limpa_Tela(Me)
''
''    iAlterado = 0
''
''    EntradaOuSaida.ListIndex = TIPO_ENTRADA
''
''''''    ApropriacaoPadrao.ListIndex = TIPO_INFORMADO_USUARIO
''
''    AtualizaConsumo.Value = 0
''
''''''    NaoEditavel.Value = False
''
''    'Gera o proximo codigo a ser utilizado
''    lErro = TiposMovEst_Automatico(iCodigo)
''    If lErro <> SUCESSO Then Error 21899
''
''    Codigo.Text = CStr(iCodigo)
''
''    Limpa_Tela_TipoMovEstoque = SUCESSO
''
''    Exit Function
''
''Erro_Limpa_Tela_TipoMovEstoque:
''
''    Limpa_Tela_TipoMovEstoque = Err
''
''    Select Case Err
''
''        Case 21895, 21899
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174891)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Public Sub Form_Unload(Cancel As Integer)
''
''Dim lErro As Long
''
''On Error GoTo Erro_Form_Unload
''
''   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
''    lErro = ComandoSeta_Liberar(Me.Name)
''    If lErro <> SUCESSO Then Error 21900
''
''    Exit Sub
''
''Erro_Form_Unload:
''
''    Select Case Err
''
''        Case 21900
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174892)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''Public Function Traz_Tela_TipoMovEstoque(objTipoMovEstoque As ClassTipoMovEst) As Long
''
''Dim lErro As Long
''
''On Error GoTo Erro_Traz_Tela_TipoMovEstoque
''
''    'Le os dados de um determinado tipo
''    lErro = TiposMovEst_Le(objTipoMovEstoque)
''    If lErro <> SUCESSO And lErro <> 21906 Then Error 21907
''
''    'Não encontrou o tipo
''    If lErro = 21906 Then Error 21908
''
''    Codigo.Text = CStr(objTipoMovEstoque.iCodigo)
''    Descricao.Text = objTipoMovEstoque.sDescricao
''
''    If UCase(objTipoMovEstoque.sEntradaOuSaida) = TIPOMOV_EST_ENTRADA Then
''        EntradaOuSaida.ListIndex = TIPO_ENTRADA
''    Else
''        If UCase(objTipoMovEstoque.sEntradaOuSaida) = TIPOMOV_EST_SAIDA Then
''            EntradaOuSaida.ListIndex = TIPO_SAIDA
''        End If
''    End If
''
''''''    ApropriacaoPadrao.ListIndex = objTipoMovEstoque.iApropriacaoPadrao
''
''    If objTipoMovEstoque.iInativo = 0 Then
''        Ativo.Value = True
''    Else
''        Ativo.Value = False
''    End If
''
''    AtualizaConsumo.Value = objTipoMovEstoque.iAtualizaConsumo
''
''    If objTipoMovEstoque.iEditavel = TIPOMOV_EST_EDITAVEL Then
''
''''''        NaoEditavel.Value = False
''
''    Else
''
''''''        NaoEditavel.Value = True
''
''    End If
''
''    Traz_Tela_TipoMovEstoque = SUCESSO
''
''    Exit Function
''
''Erro_Traz_Tela_TipoMovEstoque:
''
''    Traz_Tela_TipoMovEstoque = Err
''
''    Select Case Err
''
''        Case 21907, 21908
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174893)
''
''    End Select
''
''    Exit Function
''
''End Function
''Function Trata_Parametros(Optional objTipoMovEstoque As ClassTipoMovEst) As Long
''
''Dim lErro As Long
''
''On Error GoTo Erro_Trata_Parametros
''
''    If Not (objTipoMovEstoque Is Nothing) Then
''
''        'Traz os dados para a tela
''        lErro = Traz_Tela_TipoMovEstoque(objTipoMovEstoque)
''        If lErro <> SUCESSO And lErro <> 21908 Then Error 21916
''
''        'Nao encontrou o tipo
''        If lErro = 21908 Then Error 21917
''
''    End If
''
''    iAlterado = 0
''
''    Trata_Parametros = SUCESSO
''
''    Exit Function
''
''Erro_Trata_Parametros:
''
''    Trata_Parametros = Err
''
''    Select Case Err
''
''        Case 21916
''
''        Case 21917
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOSMOVEST_INEXISTENTE", Err, objTipoMovEstoque.iCodigo)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174894)
''
''    End Select
''
''    iAlterado = 0
''
''    Exit Function
''
''End Function
''Public Function Gravar_Registro() As Long
''
''Dim lErro As Long
''Dim iIndice As Integer
''Dim sItem As String
''Dim objTipoMovEstoque As New ClassTipoMovEst
''
''On Error GoTo Erro_Gravar_Registro
''
''    'Transfere os dados da tela para objTipoMovEstoque
''    lErro = Move_Tela_Memoria(objTipoMovEstoque)
''    If lErro <> SUCESSO Then Error 21884
''
''    'Se nao é editavel -> erro
''''''    If objTipoMovEstoque.iEditavel = TIPOMOV_EST_NAOEDITAVEL Then Error 21885
''
''    'Grava ou atualiza os dados
''    lErro = TiposMovEst_Grava(objTipoMovEstoque)
''    If lErro <> SUCESSO Then Error 21893
''
''    'Procura na ListBox
''    For iIndice = 0 To ListaTipos.ListCount
''        If ListaTipos.ItemData(iIndice) = objTipoMovEstoque.iCodigo Then Exit For
''    Next
''
''    'Se nao esta na listbox, inclui
''    If iIndice = ListaTipos.ListCount And ListaTipos.ItemData(iIndice) <> objTipoMovEstoque.iCodigo Then
''
''        sItem = CStr(objTipoMovEstoque.iCodigo) & SEPARADOR & objTipoMovEstoque.sDescricao
''        ListaTipos.AddItem sItem
''        ListaTipos.ItemData(ListaTipos.NewIndex) = objTipoMovEstoque.iCodigo
''
''    End If
''
''    'Limpa a tela
''    lErro = Limpa_Tela_TipoMovEstoque
''    If lErro <> SUCESSO Then Error 21912
''
''    Gravar_Registro = SUCESSO
''
''    Exit Function
''
''Erro_Gravar_Registro:
''
''    Gravar_Registro = Err
''
''    Select Case Err
''
''        Case 21884, 21893, 21912
''
''        Case 21885
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOSMOVEST_NAOEDITAVEL", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174895)
''
''    End Select
''
''    Exit Function
''
''End Function
''
'''""""""""""""""""""""""""""""""""""""""""""""""
'''"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'''""""""""""""""""""""""""""""""""""""""""""""""
''
'''Extrai os campos da tela que correspondem aos campos no BD
''Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
''
''
''Dim lErro As Long
''Dim objTipoMovEstoque As New ClassTipoMovEst
''
''On Error GoTo Erro_Tela_Extrai
''
''    'Informa tabela associada à Tela
''    sTabela = "TiposMovimentoEstoque"
''
''    'Lê os dados da Tela Kit
''    lErro = Move_Tela_Memoria(objTipoMovEstoque)
''    If lErro <> SUCESSO Then Error 21918
''
''    'Preenche a coleção colCampoValor, com nome do campo,
''    'valor atual (com a tipagem do BD), tamanho do campo
''    'no BD no caso de STRING e Key igual ao nome do campo
''    colCampoValor.Add "Codigo", objTipoMovEstoque.iCodigo, 0, "Codigo"
''    colCampoValor.Add "Descricao", objTipoMovEstoque.sDescricao, STRING_TIPOMOV_EST_DESCRICAO, "Descricao"
''    colCampoValor.Add "EntradaOuSaida", objTipoMovEstoque.sEntradaOuSaida, 1, "EntradaOuSaida"
''    colCampoValor.Add "Inativo", objTipoMovEstoque.iInativo, 0, "Inativo"
''    colCampoValor.Add "AtualizaConsumo", objTipoMovEstoque.iAtualizaConsumo, 0, "AtualizaConsumo"
''    colCampoValor.Add "Editavel", objTipoMovEstoque.iEditavel, 0, "Editavel"
''    colCampoValor.Add "ValidoMovInt", objTipoMovEstoque.iValidoMovInt, 0, "ValidoMovInt"
''''''    colCampoValor.Add "ApropriacaoPadrao", objTipoMovEstoque.iApropriacaoPadrao, 0, "ApropriacaoPadrao"
''''''    colCampoValor.Add "PermiteCusto", objTipoMovEstoque.iPermiteCusto, 0, "PermiteCusto"
''
''    Exit Sub
''
''Erro_Tela_Extrai:
''
''    Select Case Err
''
''        'Erro já tratado
''        Case 21918
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174896)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Function Move_Tela_Memoria(objTipoMovEstoque As ClassTipoMovEst) As Long
''
''Dim lErro As Long
''Dim sProduto As String
''
''On Error GoTo Erro_Move_Tela_Memoria
''
''    If Len(Codigo.Text) > 0 Then
''
''        objTipoMovEstoque.iCodigo = CInt(Codigo.Text)
''
''        If Len(Descricao.Text) = 0 Then Error 21882
''
''        objTipoMovEstoque.sDescricao = Descricao.Text
''
'''        If NaoEditavel.Value = True Then
'''            objTipoMovEstoque.iEditavel = TIPOMOV_EST_NAOEDITAVEL
'''        Else
'''            objTipoMovEstoque.iEditavel = TIPOMOV_EST_EDITAVEL
'''        End If
''
''        If Ativo.Value = True Then
''            objTipoMovEstoque.iInativo = 0
''        Else
''            objTipoMovEstoque.iInativo = 1
''        End If
''
''        If Len(EntradaOuSaida.Text) = 0 Then Error 21883
''
''        If EntradaOuSaida.ListIndex = 0 Then
''            objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_ENTRADA
''        Else
''            objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_SAIDA
''        End If
''
''''''        objTipoMovEstoque.iApropriacaoPadrao = ApropriacaoPadrao.ItemData(ApropriacaoPadrao.ListIndex)
''
''        objTipoMovEstoque.iAtualizaConsumo = AtualizaConsumo.Value
''
''''''        objTipoMovEstoque.iPermiteCusto = 1
''        objTipoMovEstoque.iValidoMovInt = 1
''
''    End If
''
''    Move_Tela_Memoria = SUCESSO
''
''    Exit Function
''
''Erro_Move_Tela_Memoria:
''
''    Move_Tela_Memoria = Err
''
''    Select Case Err
''
''        Case 21882
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", Err)
''
''        Case 21883
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENTRADAOUSAIDA_NAO_PREENCHIDA", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174897)
''
''    End Select
''
''    Exit Function
''
''End Function
''
'''Preenche os campos da tela com os correspondentes do BD
''Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
''
''Dim lErro As Long
''Dim objTiposMovEstoque As New ClassTipoMovEst
''
''On Error GoTo Erro_Tela_Preenche
''
''    objTiposMovEstoque.iCodigo = colCampoValor.Item("Codigo").vValor
''
''    'Traz dados do Tipo de Movimentacao de estoque para a Tela
''    lErro = Traz_Tela_TipoMovEstoque(objTiposMovEstoque)
''    If lErro <> SUCESSO And lErro <> 21908 Then Error 21919
''
''    If lErro = 21908 Then Error 21920
''
''    Exit Sub
''
''Erro_Tela_Preenche:
''
''    Select Case Err
''
''        'Erro já tratado
''        Case 21919
''
''        Case 21920
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOSMOVEST_INEXISTENTE", Err, objTiposMovEstoque.iCodigo)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174898)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
'''funcoes auxiliares p/I/O
''
'''Mover para ClassMatGrava
''Function TiposMovEst_Grava(objTiposMovEstoque As ClassTipoMovEst) As Long
'''recebe uma objTipoMovEstoque e realiza a inserção/atualização(se possível) no bd
''
''Dim lErro As Long
''Dim iIndice As Integer
''Dim iEditavel As Integer
''Dim alComando(1) As Long
''Dim lTransacao As Long
''
''On Error GoTo Erro_TiposMovEst_Grava
''
''    For iIndice = LBound(alComando) To UBound(alComando)
''
''        'Abertura comando
''        alComando(iIndice) = 0
''        alComando(iIndice) = Comando_Abrir()
''        If alComando(iIndice) = 0 Then Error 21886
''
''    Next
''
''    'Abertura transação
''    lTransacao = Transacao_Abrir()
''    If lTransacao = 0 Then Error 21887
''
''    'Le arquivo de tipos de movimentos de estoque
''    lErro = Comando_ExecutarPos(alComando(0), "SELECT Editavel FROM TiposMovimentoEstoque WHERE Codigo = ? ", 0, iEditavel, objTiposMovEstoque.iCodigo)
''    If lErro <> AD_SQL_SUCESSO Then Error 21888
''
''    lErro = Comando_BuscarPrimeiro(alComando(0))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21889
''
''    If lErro = AD_SQL_SUCESSO Then
''
''        'Se encontrou atualiza
''''''        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TiposMovimentoEstoque SET Descricao = ?, EntradaOuSaida = ?, Inativo = ?, AtualizaConsumo = ?, Editavel = ?, ValidoMovInt = ?, ApropriacaoPadrao = ?, PermiteCusto = ?", alComando(0), objTiposMovEstoque.sDescricao, objTiposMovEstoque.sEntradaOuSaida, objTiposMovEstoque.iInativo, objTiposMovEstoque.iAtualizaConsumo, objTiposMovEstoque.iEditavel, objTiposMovEstoque.iValidoMovInt, objTiposMovEstoque.iApropriacaoPadrao, objTiposMovEstoque.iPermiteCusto)
''        If lErro <> AD_SQL_SUCESSO Then Error 21890
''
''    Else
''
''        'Se nao encontrou, insere
''''''        lErro = Comando_Executar(alComando(1), "INSERT INTO TiposMovimentoEstoque (Codigo, Descricao, EntradaOuSaida, Inativo, AtualizaConsumo, Editavel, ValidoMovInt, ApropriacaoPadrao, PermiteCusto) VALUES (?,?,?,?,?,?,?,?,?)", objTiposMovEstoque.iCodigo, objTiposMovEstoque.sDescricao, objTiposMovEstoque.sEntradaOuSaida, objTiposMovEstoque.iInativo, objTiposMovEstoque.iAtualizaConsumo, objTiposMovEstoque.iEditavel, objTiposMovEstoque.iValidoMovInt, objTiposMovEstoque.iApropriacaoPadrao, objTiposMovEstoque.iPermiteCusto)
''        If lErro <> AD_SQL_SUCESSO Then Error 21891
''
''    End If
''
''    'Confirma transação
''    lErro = Transacao_Commit()
''    If lErro <> AD_SQL_SUCESSO Then Error 21892
''
''    'Fechamento comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    TiposMovEst_Grava = SUCESSO
''
''    Exit Function
''
''Erro_TiposMovEst_Grava:
''
''    TiposMovEst_Grava = Err
''
''    Select Case Err
''
''        Case 21886
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 21887
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
''
''        Case 21888, 21889
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSMOVEST", Err)
''
''        Case 21890
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TIPOSMOVEST", Err, objTiposMovEstoque.iCodigo)
''
''        Case 21891
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_TIPOSMOVEST", Err, objTiposMovEstoque.iCodigo)
''
''        Case 21892
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174899)
''
''    End Select
''
''    'Fechamento transação
''    Call Transacao_Rollback
''
''    'Fechamento comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    Exit Function
''
''End Function
''
'''Mover para ClassMatGrava
''Function TiposMovEst_Exclui(objTiposMovEstoque As ClassTipoMovEst) As Long
'''recebe um objTiposMovEstoque e exclui do bd a partir do codigo passado
''
''Dim lErro As Long
''Dim alComando(2) As Long
''Dim iIndice As Integer
''Dim iEditavel As Integer
''
''On Error GoTo Erro_TiposMovEst_Exclui
''
''    For iIndice = LBound(alComando) To UBound(alComando)
''
''        'Abertura comando
''        alComando(iIndice) = 0
''        alComando(iIndice) = Comando_Abrir()
''        If alComando(iIndice) = 0 Then Error 21870
''
''    Next
''
''    'Le tabela de movimentos de estoque
''    lErro = Comando_ExecutarPos(alComando(2), "SELECT Item FROM MovimentoEstoque WHERE TipoMov = ? ", 0, iEditavel, objTiposMovEstoque.iCodigo)
''    If lErro <> AD_SQL_SUCESSO Then Error 21871
''
''    lErro = Comando_BuscarPrimeiro(alComando(2))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21872
''
''    'Se encontrou, nao pode excluir tipo
''    If lErro = AD_SQL_SUCESSO Then Error 21873
''
''    'Le tabela de tipos de movimentos de estoque
''    lErro = Comando_ExecutarPos(alComando(0), "SELECT Editavel FROM TiposMovimentoEstoque WHERE Codigo = ? ", 0, iEditavel, objTiposMovEstoque.iCodigo)
''    If lErro <> AD_SQL_SUCESSO Then Error 21874
''
''    lErro = Comando_BuscarPrimeiro(alComando(0))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21875
''
''    'Se nao encontrou ---> erro
''    If lErro = AD_SQL_SEM_DADOS Then Error 21876
''
''    'Exclui da tabela
''    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TiposMovimentoEstoque", alComando(0))
''    If lErro <> AD_SQL_SUCESSO Then Error 21877
''
''   'Fechamento comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    TiposMovEst_Exclui = SUCESSO
''
''    Exit Function
''
''Erro_TiposMovEst_Exclui:
''
''    TiposMovEst_Exclui = Err
''
''    Select Case Err
''
''        Case 21870
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 21873, 21876
''
''        Case 21874, 21875
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSMOVEST", Err)
''
''        Case 21877
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOSMOVEST", Err, objTiposMovEstoque.iCodigo)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174900)
''
''    End Select
''
''    'Fechamento transação
''    Call Transacao_Rollback
''
''   'Fechamento comando
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    Exit Function
''
''End Function
'''Mover para ClassMatSelect
''Function TiposMovEst_Le(objTipoMovEstoque As ClassTipoMovEst) As Long
'''Le a tabela de tipos de movimentos de estoque a partir do codigo fornecido em objTipoMovEstoque e devolve os dados neste mesmo obj.
''
''Dim lErro As Long
''Dim lComando As Long
''Dim sDescricao As String
''Dim sEntradaOuSaida As String
''Dim iInativo As Integer
''Dim iAtualizaConsumo As Integer
''Dim iEditavel As Integer
''Dim iApropriacaoPadrao As Integer
''
''On Error GoTo Erro_TiposMovEst_Le
''
''    'Abertura comando
''    lComando = 0
''
''    lComando = Comando_Abrir()
''    If lComando = 0 Then Error 21903
''
''    sDescricao = String(STRING_TIPOMOV_EST_DESCRICAO, 0)
''''''    sEntradaOuSaida = String(Len(STRING_TIPOMOV_EST_ENTRADAOUSAIDA), 0)
''
''    'Le a tabela de tipos de movimento de estoque
''    lErro = Comando_Executar(lComando, "SELECT Descricao, EntradaOuSaida, Inativo, AtualizaConsumo, Editavel, ApropriacaoPadrao FROM TiposMovimentoEstoque WHERE Codigo = ? ", sDescricao, sEntradaOuSaida, iInativo, iAtualizaConsumo, iEditavel, iApropriacaoPadrao, objTipoMovEstoque.iCodigo)
''    If lErro <> AD_SQL_SUCESSO Then Error 21904
''
''    lErro = Comando_BuscarPrimeiro(lComando)
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21905
''
''    If lErro = AD_SQL_SEM_DADOS Then Error 21906
''
''    objTipoMovEstoque.sDescricao = sDescricao
''    objTipoMovEstoque.sEntradaOuSaida = sEntradaOuSaida
''    objTipoMovEstoque.iInativo = iInativo
''    objTipoMovEstoque.iAtualizaConsumo = iAtualizaConsumo
''    objTipoMovEstoque.iEditavel = iEditavel
''''''    objTipoMovEstoque.iApropriacaoPadrao = iApropriacaoPadrao
''
''    'Fechamento comando
''    Call Comando_Fechar(lComando)
''
''    TiposMovEst_Le = SUCESSO
''
''    Exit Function
''
''Erro_TiposMovEst_Le:
''
''    TiposMovEst_Le = Err
''
''    Select Case Err
''
''        Case 21903
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 21904, 21905
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSMOVEST", Err)
''
''        Case 21906
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174901)
''
''    End Select
''
''   'Fechamento comando
''    Call Comando_Fechar(lComando)
''
''    Exit Function
''
''End Function
''
'''Mover para ClassMat
''Function TiposMovEst_Automatico(iCodigo As Integer) As Long
'''Le a tabela de tipos de movimentos de estoque e retorna o proximo codigo disponivel.
''
''Dim lErro As Long
''Dim lComando As Long
''
''On Error GoTo Erro_TiposMovEst_Automatico
''
''    'Abertura comando
''    lComando = 0
''
''    lComando = Comando_Abrir()
''    If lComando = 0 Then Error 21896
''
''    'Le o ultimo codigo de usuario utilizado na tabela de tipos de movimentos de estoque.
''    lErro = Comando_Executar(lComando, "SELECT Codigo FROM TiposMovimentoEstoque WHERE Codigo > ? ORDER BY Codigo DESC", iCodigo, TIPOMOV_EST_ULTIMOCODIGORESERVADO)
''    If lErro <> AD_SQL_SUCESSO Then Error 21897
''
''    lErro = Comando_BuscarPrimeiro(lComando)
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21898
''
''    If lErro = AD_SQL_SEM_DADOS Then
''        'Se nao encontrou, atribui o primeiro codigo permitido ao usuario
''        iCodigo = TIPOMOV_EST_ULTIMOCODIGORESERVADO + 1
''    Else
''        'Se encontrou, devolve um maior
''        iCodigo = iCodigo + 1
''    End If
''
''    'Fechamento comando
''    Call Comando_Fechar(lComando)
''
''    TiposMovEst_Automatico = SUCESSO
''
''    Exit Function
''
''Erro_TiposMovEst_Automatico:
''
''    TiposMovEst_Automatico = Err
''
''    Select Case Err
''
''        Case 21896
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 21897, 21898
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSMOVEST", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174902)
''
''    End Select
''
''   'Fechamento comando
''    Call Comando_Fechar(lComando)
''
''    Exit Function
''
''End Function
''
'''**** inicio do trecho a ser copiado *****
''
''Public Function Form_Load_Ocx() As Object
''
''    Parent.HelpContextID = IDH_TIPOS_MOV_ESTOQUE
''    Set Form_Load_Ocx = Me
''    Caption = "Tipo de Movimento de Estoque"
''    Call Form_Load
''
''End Function
''
''Public Function Name() As String
''
''    Name = "TipoMovEstoque"
''
''End Function
''
''Public Sub Show()
''    Parent.Show
''End Sub
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=UserControl,UserControl,-1,Controls
''Public Property Get Controls() As Object
''    Set Controls = UserControl.Controls
''End Property
''
''Public Property Get hWnd() As Long
''    hWnd = UserControl.hWnd
''End Property
''
''Public Property Get Height() As Long
''    Height = UserControl.Height
''End Property
''
''Public Property Get Width() As Long
''    Width = UserControl.Width
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=UserControl,UserControl,-1,ActiveControl
''Public Property Get ActiveControl() As Object
''    Set ActiveControl = UserControl.ActiveControl
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=UserControl,UserControl,-1,Enabled
''Public Property Get Enabled() As Boolean
''    Enabled = UserControl.Enabled
''End Property
''
''Public Property Let Enabled(ByVal New_Enabled As Boolean)
''    UserControl.Enabled() = New_Enabled
''    PropertyChanged "Enabled"
''End Property
''
'''Load property values from storage
''Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
''
''    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
''End Sub
''
'''Write property values to storage
''Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
''
''    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
''End Sub
''
''Private Sub Unload(objme As Object)
''
''   RaiseEvent Unload
''
''End Sub
''
''Public Property Get Caption() As String
''    Caption = m_Caption
''End Property
''
''Public Property Let Caption(ByVal New_Caption As String)
''    Parent.Caption = New_Caption
''    m_Caption = New_Caption
''End Property
''
'''**** fim do trecho a ser copiado *****
''
''

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub TipoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoLabel, Source, X, Y)
End Sub

Private Sub TipoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoLabel, Button, Shift, X, Y)
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

