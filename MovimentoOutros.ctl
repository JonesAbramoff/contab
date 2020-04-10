VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl MovimentoOutros 
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   KeyPreview      =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   7050
   Begin VB.Frame FrameOutrosDetalhados 
      Caption         =   "Outros Detalhados"
      Height          =   4440
      Left            =   240
      TabIndex        =   11
      Top             =   1050
      Width           =   6720
      Begin VB.ComboBox Parcelamento 
         Height          =   315
         Left            =   2115
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   255
         Width           =   1590
      End
      Begin MSMask.MaskEdBox ValorSangria 
         Height          =   300
         Left            =   4965
         TabIndex        =   24
         Top             =   255
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoValoresEmCaixa 
         Height          =   585
         Left            =   165
         Picture         =   "MovimentoOutros.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "F9 - Traz para a tela os boletos em caixa que podem ser ""sangrados""."
         Top             =   3735
         Width           =   1680
      End
      Begin VB.ComboBox Administradora 
         Height          =   315
         Left            =   540
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   270
         Width           =   1575
      End
      Begin MSMask.MaskEdBox ValorTotal 
         Height          =   300
         Left            =   3705
         TabIndex        =   13
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridOutros 
         Height          =   3315
         Left            =   150
         TabIndex        =   4
         Top             =   285
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   5847
         _Version        =   393216
         Rows            =   5
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LabelSangriaValor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5040
         TabIndex        =   23
         Top             =   3690
         Width           =   1230
      End
      Begin VB.Label LabelEmCaixaValor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3795
         TabIndex        =   22
         Top             =   3690
         Width           =   1230
      End
      Begin VB.Label LabelEmCaixaDetalhados 
         AutoSize        =   -1  'True
         Caption         =   "Totais:"
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
         Left            =   3150
         TabIndex        =   21
         Top             =   3750
         Width           =   600
      End
   End
   Begin VB.Frame FrameIdentificacao 
      Caption         =   "Identificação"
      Height          =   870
      Left            =   240
      TabIndex        =   9
      Top             =   45
      Width           =   3645
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1695
         Picture         =   "MovimentoOutros.ctx":2D5A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   360
         Width           =   300
      End
      Begin VB.CommandButton BotaoTrazer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2085
         Picture         =   "MovimentoOutros.ctx":2E44
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "F4 - Exibe na tela o movimento com o código informado."
         Top             =   210
         Width           =   1440
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   345
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Mask            =   "#######"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCodigo 
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
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   390
         Width           =   660
      End
   End
   Begin VB.CommandButton DesmembraMovto 
      Caption         =   "Desmembra"
      Height          =   465
      Left            =   150
      TabIndex        =   19
      Top             =   135
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame FrameOutrosNaoDetalhados 
      Caption         =   "Outros Não Detalhados"
      Height          =   825
      Left            =   240
      TabIndex        =   14
      Top             =   5640
      Width           =   6720
      Begin VB.TextBox ValorSangriaNaoDetalhado 
         Height          =   315
         Left            =   3960
         TabIndex        =   15
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label LabelEmCaixaNaoDetalhados 
         AutoSize        =   -1  'True
         Caption         =   "Em caixa:"
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
         TabIndex        =   18
         Top             =   405
         Width           =   840
      End
      Begin VB.Label LabelEmCaixaNaoDetalhadosValor 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1575
         TabIndex        =   17
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label LabelSangriaNaoDetalhado 
         AutoSize        =   -1  'True
         Caption         =   "Sangria:"
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
         Left            =   3120
         TabIndex        =   16
         Top             =   405
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4830
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "MovimentoOutros.ctx":5B0E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "F8 - Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "MovimentoOutros.ctx":5C8C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "F7 - Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "MovimentoOutros.ctx":61BE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "F6 - Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "MovimentoOutros.ctx":6348
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "F5 - Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   525
      Left            =   195
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   926
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"MovimentoOutros.ctx":64A2
   End
End
Attribute VB_Name = "MovimentoOutros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variáveis Globais
Dim objGridOutros As AdmGrid
Dim iAlterado As Integer
Dim iAdmnistradora_Alterado As Integer
Dim iValorSangria_Alterado As Integer
Dim iGrid_ValorTotal_Col As Integer
Dim iGrid_ValorSangria_Col As Integer
Dim iGrid_Administradora_Col As Integer
Dim iGrid_Parcelamento_Col As Integer
Dim glProxNumAuto As Long
 
'So para uso da função que desmebra o Log
Dim colMovimentosCaixaOutros As New Collection
Dim gcolImfCompl As New Collection

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo)

End Sub


'**** inicio do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Sangria Outros Meios de Pagamento"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "MovimentoGridOutros"

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
    'Parent.UnloadDoFilho

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String

    Caption = m_Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)

    Parent.Caption = New_Caption
    m_Caption = New_Caption

End Property

Public Sub Form_Load()
'Função inicialização da Tela
Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Instanciar o objGridOutros para apontar para uma posição de memória
    Set objGridOutros = New AdmGrid

    'Inicialização de GridOutros
    lErro = Inicializa_GridOutros(objGridOutros)
    If lErro <> SUCESSO Then gError 107987

    'Função que Carrega os Meios de Pagto relacionadas a Outros
    lErro = Carrega_Combo_Administradora()
    If lErro <> SUCESSO Then gError 107989

    Call BotaoValoresEmCaixa_Click

    'Indica que o carregamento da tela aconteceu com sucesso
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        'Erros tratados na rotina chamada
        Case 107987 To 107989

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163023)

    End Select

    Exit Sub

End Sub

Function Inicializa_GridOutros(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridOutros

   'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Administradora")
    objGridInt.colColuna.Add ("Parcelamento")
    objGridInt.colColuna.Add ("Valor em Caixa")
    objGridInt.colColuna.Add ("Valor Sangria")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Administradora.Name)
    objGridInt.colCampo.Add (Parcelamento.Name)
    objGridInt.colCampo.Add (ValorTotal.Name)
    objGridInt.colCampo.Add (ValorSangria.Name)

    'Colunas do Grid
    iGrid_Administradora_Col = 1
    iGrid_Parcelamento_Col = 2
    iGrid_ValorTotal_Col = 3
    iGrid_ValorSangria_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridOutros

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_LINHAS_GRID

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Largura da primeira coluna
    GridOutros.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

     
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridOutros = SUCESSO

    Exit Function

Erro_Inicializa_GridOutros:

    Inicializa_GridOutros = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163024)

    End Select

    Exit Function

End Function

Function Recalcula_Totais() As Long
'Calcula o Valor Total de Pagamentos de outras formas

Dim lErro As Long
Dim iIndice As Integer
Dim dValorTotalOutros As Double
Dim dValorSangria As Double
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Recalcula_Totais

    'Para todos os Cheque do Grid
    For iIndice = 1 To objGridOutros.iLinhasExistentes

        'Acumula o Valor Total dos meios de Pagto
        dValorTotalOutros = dValorTotalOutros + StrParaDbl(GridOutros.TextMatrix(iIndice, iGrid_ValorTotal_Col))

        'Acumula o Valor Total da Sangria
        dValorSangria = dValorSangria + StrParaDbl(GridOutros.TextMatrix(iIndice, iGrid_ValorSangria_Col))

    Next

    
    'Exibe o valor total tando de Sangria como o valor total de meio de pagto Outros que existem no caixa
    LabelEmCaixaValor.Caption = Format(dValorTotalOutros, "Standard")
    LabelSangriaValor.Caption = Format(dValorSangria, "Standard")


    Recalcula_Totais = SUCESSO

    Exit Function

Erro_Recalcula_Totais:

    Recalcula_Totais = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163025)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    If iCaminho = ROTINA_GRID_TRATA_TECLA_CAMPO2 Or iCaminho = ROTINA_GRID_ENTRADA_CELULA Or iCaminho = ROTINA_GRID_CLICK Then

        'Pesquisa controle da coluna em questão
        Select Case objControl.Name
    
            Case Parcelamento.Name
    
                If Len(Trim(GridOutros.TextMatrix(iLinha, iGrid_Administradora_Col))) = 0 Then
    
                    objControl.Enabled = False
    
                Else
    
                    objControl.Enabled = True
    
                    'Carrega a Combo de Parcelamentos
                    lErro = Carrega_Combo_Parcelamento(iLinha)
                    If lErro <> SUCESSO Then gError 105935
    
                End If
    
    
            'Se o campo for valor sangria
            Case ValorSangria.Name
    
                'Verifca se o Campo valorTotal está Preenchido
                 If Len(Trim(GridOutros.TextMatrix(iLinha, iGrid_ValorTotal_Col))) <> 0 Then
    
                    objControl.Enabled = True
    
                Else
    
                    'Desabilita o Campo Valor Total de Sangria
                    objControl.Enabled = False
    
                End If
    
        End Select

    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 105935

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163026)

    End Select

    Exit Sub

End Sub

Function Carrega_Combo_Parcelamento(iLinha As Integer) As Long
'Função que Carrega a Combo de Parcelamentos, Verificando se o parcelamento selecionado é igual a um dos parcelamento que é referente a admnistradora selecinada

Dim lErro As Long
Dim sNomeParcelamento As String
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim sNomeAdmnistradora As String
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Combo_Parcelamento

    'Atribui a Variável o nome do parcelamento
    sNomeParcelamento = GridOutros.TextMatrix(iLinha, iGrid_Parcelamento_Col)

    'Atribui o Nome da Admnistradora Selecionada
    sNomeAdmnistradora = GridOutros.TextMatrix(iLinha, iGrid_Administradora_Col)
    
    'Limpa a Combo
     Parcelamento.Clear
    
    'Incluir todos os Parcelamentos cadastratrados na Coleção Glogal de Parcelamentos Referenciando a Admnistradora Referenciada
    For Each objAdmMeioPagtoCondPagto In gcolOutros

         If sNomeAdmnistradora = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto And objAdmMeioPagtoCondPagto.dSaldo > 0 Then

            Parcelamento.AddItem objAdmMeioPagtoCondPagto.sNomeParcelamento
            Parcelamento.ItemData(Parcelamento.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento

        End If
    Next


    For iIndice = 0 To Parcelamento.ListCount - 1

        If sNomeParcelamento = Parcelamento.List(iIndice) Then
            Parcelamento.ListIndex = iIndice
            Exit For
        End If

    Next
    
    Carrega_Combo_Parcelamento = SUCESSO

    Exit Function

Erro_Carrega_Combo_Parcelamento:

    Carrega_Combo_Parcelamento = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163027)

    End Select

    Exit Function

End Function

Function Carrega_Combo_Administradora() As Long
'Função que Carrega a Combo de Admnistradoras

Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Carrega_Combo_Administradora


    For Each objAdmMeioPagtoCondPagto In gcolOutros

        'se tiver saldo e for especificado
        If objAdmMeioPagtoCondPagto.dSaldo > 0 And objAdmMeioPagtoCondPagto.iAdmMeioPagto <> 0 And objAdmMeioPagtoCondPagto.iAdmMeioPagto <> MEIO_PAGAMENTO_CONTRAVALE Then

                'Senão for nenhum dois dos acima carrega na combo de Administradora
                Administradora.AddItem objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
                Administradora.ItemData(Administradora.NewIndex) = objAdmMeioPagtoCondPagto.iAdmMeioPagto

        End If
        
    Next

    Carrega_Combo_Administradora = SUCESSO

    Exit Function

Erro_Carrega_Combo_Administradora:

    Carrega_Combo_Administradora = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163028)

    End Select

    Exit Function

End Function

Private Sub BotaoProxNum_Click()
'Botão que Gera um Próximo Numero para Movto

Dim lErro As Long
Dim lNumero As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Função que Gera o Próximo Código para a Tela de Sangria Relaciona ao meio de pgto Outros
    lErro = CF_ECF("Caixa_Obtem_NumAutomatico", lNumero)
    If lErro <> SUCESSO Then gError 111048

    'Exibir o Numero na Tela
    Codigo.Text = lNumero

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 111048

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163029)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazer_Click()
'Função que chama a função que preenche o grid

Dim lErro As Long
Dim lCodigo As String

On Error GoTo Erro_botaoTrazer_click

    'Verifica se o código não está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 111051

    lCodigo = StrParaLong(Codigo.Text)
    
    Call Limpa_Tela_MovimentoOutros
    
    Codigo.Text = lCodigo
    
    'Chama a função que preenche o grid
    lErro = Traz_MovimentoOutros_Tela(StrParaLong(Codigo.Text))
    If lErro <> SUCESSO Then gError 111052

    'Anula a Alteração
    iAlterado = 0
    
    Exit Sub

Erro_botaoTrazer_click:

    Select Case gErr

        Case 111051
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 111052
            'Erro tradado Dentro da Função que Foi Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163030)

    End Select

    Exit Sub

End Sub

Function Traz_MovimentoOutros_Tela(lNumero As Long) As Long
'Função que Trar o Movimento Outros para a tela

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim objMovimentoCaixa As ClassMovimentoCaixa
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iLinha As Integer

On Error GoTo Erro_Traz_MovimentoOutros_Tela

    Call Limpa_Tela_MovimentoOutros
    
    Codigo.Text = lNumero
    
    'Função que Lê os Movimentos de Caixa
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, lNumero)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 111053
    
    'Não existe Movimento com este Código
    If lErro = 107850 Then gError 111054

    For Each objMovimentoCaixa In colMovimentosCaixa
    
        'se o movimento não for de sangria de outros ==> erro
        If objMovimentoCaixa.iTipo <> MOVIMENTOCAIXA_SANGRIA_OUTROS Then gError 105741
            
        For Each objAdmMeioPagtoCondPagto In gcolOutros
        
            'Verifica se o Código do meio de Pagto é igual ao Código do Meio de Pagto do Movto
            If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovimentoCaixa.iAdmMeioPagto Then

                'Verifica se o codigo do Meio de Pagto é igual a zero
                If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        
                    'Exibe o Saldo + Valor do Movto
                    LabelEmCaixaNaoDetalhadosValor.Caption = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo + objMovimentoCaixa.dValor, "standard"))
                    'Coloca o Valor do Movto
                    ValorSangriaNaoDetalhado.Text = Format(objMovimentoCaixa.dValor, "Standard")
                    
                    Exit For

                'Senão Preenche o Grid, se o codigo do Meio for igual ao Código do meio do Movto
                ElseIf objAdmMeioPagtoCondPagto.iParcelamento = objMovimentoCaixa.iParcelamento Then

                    iLinha = iLinha + 1

                    GridOutros.TextMatrix(iLinha, iGrid_Administradora_Col) = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
                    GridOutros.TextMatrix(iLinha, iGrid_Parcelamento_Col) = objAdmMeioPagtoCondPagto.sNomeParcelamento
                    GridOutros.TextMatrix(iLinha, iGrid_ValorTotal_Col) = Format(objAdmMeioPagtoCondPagto.dSaldo + objMovimentoCaixa.dValor, "standard")
                    GridOutros.TextMatrix(iLinha, iGrid_ValorSangria_Col) = Format(objMovimentoCaixa.dValor, "standard")
                    Exit For
                    
                End If

            End If

        Next
    
    Next
    
    'atualizar Linhas Existentas
    objGridOutros.iLinhasExistentes = iLinha

    'Função que Atualiza os Totais
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 111055
    
    Traz_MovimentoOutros_Tela = SUCESSO

    Exit Function

Erro_Traz_MovimentoOutros_Tela:

    Traz_MovimentoOutros_Tela = gErr

    Select Case gErr

        Case 105741
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_SANGRIA_OUTROS, gErr, lNumero)

        Case 111054
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_INEXISTENTE, gErr, lNumero)
        
        Case 111053, 111055
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163031)

    End Select

    Exit Function

End Function

Private Sub GridOutros_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridOutros, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridOutros, iAlterado)
    End If

End Sub

Private Sub GridOutros_EnterCell()

    'Parametro não opcional
    Call Grid_Entrada_Celula(objGridOutros, iAlterado)

End Sub

Private Sub GridOutros_GotFocus()

    Call Grid_Recebe_Foco(objGridOutros)

End Sub

Private Sub GridOutros_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridOutros_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridOutros)

    If KeyCode = vbKeyDelete Then
        
        'Função que recalcula os totais no grid
        lErro = Recalcula_Totais()
        If lErro <> SUCESSO Then gError 105745
        
    End If

    Exit Sub

Erro_GridOutros_KeyDown:

    Select Case gErr

        Case 105745

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163032)

    End Select

    Exit Sub

End Sub

Private Sub GridOutros_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridOutros, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOutros, iAlterado)
    End If

End Sub

Private Sub GridOutros_LeaveCell()

    Call Saida_Celula(objGridOutros)

End Sub

Private Sub GridOutros_LostFocus()

    Call Grid_Libera_Foco(objGridOutros)

End Sub
Private Sub GridOutros_RowColChange()

    Call Grid_RowColChange(objGridOutros)

End Sub

Private Sub GridOutros_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridOutros)

End Sub

Private Sub GridOutros_Scroll()

    Call Grid_Scroll(objGridOutros)

End Sub

Private Sub Administradora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Administradora_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOutros)


End Sub

Private Sub Administradora_KeyPress(KeyAscii As Integer)


    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOutros)

End Sub

Private Sub Administradora_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOutros.objControle = Administradora
    lErro = Grid_Campo_Libera_Foco(objGridOutros)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Parcelamento_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Parcelamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOutros)

End Sub

Private Sub Parcelamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOutros)

End Sub

Private Sub Parcelamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOutros.objControle = Parcelamento
    lErro = Grid_Campo_Libera_Foco(objGridOutros)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorTotal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorTotal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOutros)

End Sub

Private Sub ValorTotal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOutros)

End Sub

Private Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOutros.objControle = ValorTotal
    lErro = Grid_Campo_Libera_Foco(objGridOutros)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorSangria_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorSangria_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorSangria_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOutros)

End Sub

Private Sub ValorSangria_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOutros)

End Sub

Private Sub ValorSangria_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOutros.objControle = ValorSangria
    lErro = Grid_Campo_Libera_Foco(objGridOutros)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

    'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col

            'Meio de Pagto
            Case iGrid_Administradora_Col
                lErro = Saida_Celula_Administradora(objGridInt)
                If lErro <> SUCESSO Then gError 111056

            'Parcelamento
            Case iGrid_Parcelamento_Col
                lErro = Saida_Celula_Parcelamento(objGridInt)
                If lErro <> SUCESSO Then gError 105936
            
            'ValorSangria
            Case iGrid_ValorSangria_Col
                lErro = Saida_Celula_ValorSangria(objGridInt)
                If lErro <> SUCESSO Then gError 111001

        End Select

    'Função que Finaliza a Saida de Celula
    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 111002

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 105936, 111001 To 111002, 111056
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163033)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Administradora(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iAchou As Integer

On Error GoTo Erro_Saida_Celula_Administradora

    Set objGridInt.objControle = Administradora
    
    If Administradora.Text <> GridOutros.TextMatrix(GridOutros.Row, iGrid_Administradora_Col) Then

        'Limpa o Grid na Linha Relacionada a Administradora em Questão na Coluna Relacionada ao Parcelamento
        GridOutros.TextMatrix(GridOutros.Row, iGrid_Parcelamento_Col) = ""
        
        'Limpa o Grid na Linha Relacionada a Administradora em Questão na Coluna Relacionada ao ValorTotal
        GridOutros.TextMatrix(GridOutros.Row, iGrid_ValorTotal_Col) = ""

        'Limpa o Grid na Linha Relacionada a Administradora em Questão na Coluna Relacionada ao Valor da Sangria
        GridOutros.TextMatrix(GridOutros.Row, iGrid_ValorSangria_Col) = ""

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 105937

    'Verifica se célula do Grid que identifica o campo Admnistradora esta preenchido
    If Len(Trim(GridOutros.TextMatrix(GridOutros.Row, iGrid_Administradora_Col))) <> 0 Then

        lErro = Recalcula_Totais()
        If lErro <> SUCESSO Then gError 105938

        'Acrescenta uma linha no Grid se for o caso
        If GridOutros.Row - GridOutros.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    Saida_Celula_Administradora = SUCESSO

    Exit Function

Erro_Saida_Celula_Administradora:

    Saida_Celula_Administradora = gErr

    Select Case gErr

        Case 105937, 105938

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163034)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Parcelamento(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim iLinha As Integer
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim iAchou As Integer

On Error GoTo Erro_Saida_Celula_Parcelamento

    Set objGridInt.objControle = Parcelamento

    If Parcelamento.Text <> GridOutros.TextMatrix(GridOutros.Row, iGrid_Parcelamento_Col) Then

        'Verifica se o Parcelamento esta selecionada ou foi Alterada
        If Len(Trim(Parcelamento.Text)) = 0 Then
    
            'Limpa o Campo Valor
            GridOutros.TextMatrix(GridOutros.Row, iGrid_ValorTotal_Col) = ""
    
            'Limpa o Campo Relacionado a Sangria
            GridOutros.TextMatrix(GridOutros.Row, iGrid_ValorSangria_Col) = ""
    
        'Se o campo Parcelamento foi Alterado e se o Campo Relacionado a Terminal Já Estivar Preenchido
        Else
    
            'Para ver se existe duplicidade no Grig
            For iLinha = 1 To objGridOutros.iLinhasExistentes
    
                If iLinha <> GridOutros.Row Then
    
                    If GridOutros.TextMatrix(iLinha, iGrid_Administradora_Col) = GridOutros.TextMatrix(GridOutros.Row, iGrid_Administradora_Col) And _
                    GridOutros.TextMatrix(iLinha, iGrid_Parcelamento_Col) = Parcelamento.Text Then gError 105939
    
                End If
    
            Next
            
    
            'Procura na Coleção de Outros a Tupla correspondente e preenche o Grid
            For Each objAdmMeioPagtoCondPagto In gcolOutros
    
                'Verifica se o Nome da Admnistradora é Igual e o parcelamento Selecionada é Igual ao da Tupla Admnistradora + parcelmento
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = GridOutros.TextMatrix(GridOutros.Row, iGrid_Administradora_Col) And objAdmMeioPagtoCondPagto.sNomeParcelamento = Parcelamento.Text Then
    
                        'Preenche o Grid com o Valor Total refente aos Outros em Questão
    
                        GridOutros.TextMatrix(GridOutros.Row, iGrid_ValorTotal_Col) = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
                        GridOutros.TextMatrix(GridOutros.Row, iGrid_ValorSangria_Col) = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
                        iAchou = 1
                        Exit For
    
                    End If
    
            Next
    
            If iAchou = 0 Then gError 105940
    
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 105941

    'Acrescenta uma linha no Grid se for o caso
    If GridOutros.Row - GridOutros.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If

    'Função que recalcula os totais no grid
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 105942

    Saida_Celula_Parcelamento = SUCESSO

    Exit Function

Erro_Saida_Celula_Parcelamento:

    Saida_Celula_Parcelamento = gErr

    Select Case gErr

        Case 105939
            Call Rotina_ErroECF(vbOKOnly, ERRO_LINHA_REPETIDA, gErr)

        Case 105940
            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_EXISTENTE1, gErr, Parcelamento.Text)

        Case 105941, 105942

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163035)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorSangria(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Saida_Celula_ValorSangria

    Set objGridInt.objControle = ValorSangria

    'Verifica se o Valor da Sangria esta Selecionado ou foi Alterada
    If Len(Trim(ValorSangria.Text)) <> 0 Then

        'Verifica se o Valor Digitado é Valido
        lErro = Valor_NaoNegativo_Critica(ValorSangria.Text)
        If lErro <> SUCESSO Then gError 111005

        'Verifica se o valor da sangria é maior que o valor total se for Erro
        If StrParaDbl(ValorSangria.Text) > StrParaDbl(GridOutros.TextMatrix(GridOutros.Row, iGrid_ValorTotal_Col)) Then gError 111006


    End If

    'Adcionar ao Grid
    GridOutros.TextMatrix(GridOutros.Row, iGrid_ValorSangria_Col) = ValorSangria.Text

    'Função que recalcula os totais no grid
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 111007

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 111008

    'Acrescenta uma linha no Grid se for o caso
    If GridOutros.Row - GridOutros.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If


    Saida_Celula_ValorSangria = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorSangria:

    Saida_Celula_ValorSangria = gErr

    Select Case gErr

        Case 111005

        Case 111006
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORSANGRIA_NAO_DISPONIVEL_GRID, gErr, ValorSangria.Text, GridOutros.TextMatrix(GridOutros.Row, iGrid_ValorTotal_Col), GridOutros.Row)

        Case 111007

        Case 111008
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163036)

    End Select

    Exit Function

End Function

Private Sub ValorSangriaNaoDetalhado_Validate(Cancel As Boolean)
'Função que valida os dados no campo valor de sangria de meios de Pagto não detalhados

Dim lErro As Long

On Error GoTo Erro_ValorSangriaNaoDetalhado_Validate

    'Verifica se o campo esta preenchido
    If Len(Trim(ValorSangriaNaoDetalhado.Text)) <> 0 Then

        'se esta preenchido então verificar o valor
        lErro = Valor_NaoNegativo_Critica(ValorSangriaNaoDetalhado.Text)
        If lErro <> SUCESSO Then gError 111009

        'Verifica se o valor da sangria é maior do que o valor do meio de pagto outros não especificados
        If StrParaDbl(ValorSangriaNaoDetalhado.Text) > StrParaDbl(LabelEmCaixaNaoDetalhadosValor.Caption) Then gError 111010

        ValorSangriaNaoDetalhado.Text = Format(ValorSangriaNaoDetalhado.Text, "standard")

    End If

    Exit Sub

Erro_ValorSangriaNaoDetalhado_Validate:

    Cancel = True

    Select Case gErr

        Case 111009

        Case 111010
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORSANGRIA_NAO_DISPONIVEL, gErr, ValorSangriaNaoDetalhado.Text, giCodCaixa)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163037)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_MovimentoOutros() As Long
'Função que Limpa a Tela

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_MovimentoOutros

    'Limpa os Controles básico da Tela
    Call Limpa_Tela(Me)

    'Limpa Grid
    Call Grid_Limpa(objGridOutros)
    
    'Limpa os Labes da Tela
    LabelEmCaixaValor.Caption = Format(0, "standard")
    LabelEmCaixaNaoDetalhadosValor.Caption = Format(0, "standard")
    LabelSangriaValor.Caption = Format(0, "standard")

    iAlterado = 0
    
    Limpa_Tela_MovimentoOutros = SUCESSO

    Exit Function

Erro_Limpa_Tela_MovimentoOutros:

    Limpa_Tela_MovimentoOutros = gErr

    Select Case gErr

        Case 111011
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163038)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Função que Realiza a Gravação

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim vbMsgRes As VbMsgBoxResult
Dim iTipoMovimento As Integer
Dim lNumMovto As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207977

    'Função que efeuara a gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 111012

    'Função que Limpa a Tela
    lErro = Limpa_Tela_MovimentoOutros
    If lErro <> SUCESSO Then gError 111013
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 111012, 111013, 207977

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163039)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim lNumMovto As Long
Dim iTipoMovimento As Integer
Dim colMovimentosCaixa As New Collection
Dim vbMsgRes As VbMsgBoxResult
Dim objMovCC As ClassMovimentoCaixa

On Error GoTo Erro_Gravar_Registro

    'Função que Valida a Gravação
    lErro = MovimentoOutros_Valida_Gravacao()
    If lErro <> SUCESSO Then gError 111014
    
    'Transforma o Codigo Texto para Long
    lNumMovto = StrParaLong(Codigo.Text)

    'Lê os Movimentos de Caixa da coleção global e carrega no coleção local a tela para o numero do Movto
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, lNumMovto)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 111015
    
    'Verifica se já existe um movimento para o código referido, Verifica se é Alteração
    If colMovimentosCaixa.Count > 0 Then

        Set objMovCC = colMovimentosCaixa(1)
    
        If objMovCC.iTipo <> MOVIMENTOCAIXA_SANGRIA_OUTROS Then gError 86291
        
         iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_OUTROS

        'Envia aviso perguntando se deseja atualizar o movimemtos
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ALTERACAO_MOVIMENTOCAIXA, Codigo.Text)

        'Se a Reposta for Negativa
        If vbMsgRes = vbNo Then gError 111016

        'Função que Faz a Alteração na Sangria do meio de pagto relaciona a Outros Previamente Executada, adciona o iTipoMovimento
        lErro = MovimentoCaixa_Prepara_Exclusao(colMovimentosCaixa, iTipoMovimento)
        If lErro <> SUCESSO Then gError 111017

    End If

    'Move os Dados da Sangria para a Memoria
    lErro = Move_Dados_Memoria(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 111018
    
    'Função que Grava os movimentos em Arquivos
    lErro = Caixa_Grava_Movimento(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 111019

    'Atualiza os Dados da Memoria
    lErro = MovimentoOutros_Atualiza_Memoria(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 111020

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

   Select Case gErr

        Case 86291
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_OUTROS, gErr, StrParaDbl(Codigo.Text))
            
        Case 86292
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
        
        Case 111014 To 111020

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163040)

    End Select

End Function

Function MovimentoOutros_Valida_Gravacao() As Long
'Função que Valida a Gravação

Dim lErro As Long
Dim iIndice As Integer
Dim dValor As Double
Dim dValorAtual As Double
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim objMovCaixa As ClassMovimentoCaixa
Dim iCont As Integer

On Error GoTo Erro_MovimentoOutros_Valida_Gravacao

    'Verifica se o código Foi Preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 111021

    'Verifica se não existem Linhas no Grid e as Label's não estejam preenchidas
    If objGridOutros.iLinhasExistentes = 0 And Len(Trim(ValorSangriaNaoDetalhado.Text)) = 0 And Len(Trim(LabelSangriaValor.Caption)) = 0 Then gError 111022

    'Verifica se no Grid o Campo Referente a Sangria Esta Preenchido
    For iIndice = 1 To objGridOutros.iLinhasExistentes

        If Len(GridOutros.TextMatrix(iIndice, iGrid_Administradora_Col)) = 0 Then gError 105943
        
        If Len(GridOutros.TextMatrix(iIndice, iGrid_Parcelamento_Col)) = 0 Then gError 105944
        
        If StrParaDbl(GridOutros.TextMatrix(iIndice, iGrid_ValorSangria_Col)) = 0 Then gError 105945

        'Para ver se existe duplicidade no Grig
        For iCont = 1 To objGridOutros.iLinhasExistentes

            If iCont <> iIndice Then

                If GridOutros.TextMatrix(iCont, iGrid_Administradora_Col) = GridOutros.TextMatrix(iIndice, iGrid_Administradora_Col) And _
                GridOutros.TextMatrix(iCont, iGrid_Parcelamento_Col) = GridOutros.TextMatrix(iIndice, iGrid_Parcelamento_Col) Then gError 105946
            
            End If

        Next


'        For Each objAdmMeioPagtoCondPagto In gcolOutros
'
'            'Verifica se o Código do meio de Pagto é igual ao Código do Meio de Pagto do Movto
'            If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = GridOutros.TextMatrix(iIndice, iGrid_Administradora_Col) Then
'
'                dValorAtual = 0
'
'                'se for uma alteração guarda o valor sangrado para ser somando ao valor disponivel para ser sangrado
'                For Each objMovCaixa In colMovimentosCaixa
'
'                    If objMovCaixa.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto Then
'
'                        dValorAtual = objMovCaixa.dValor
'                        Exit For
'
'                    End If
'
'                Next
'
'                If dValor > objAdmMeioPagtoCondPagto.dSaldo + dValorAtual Then gError 105746
'
'            End If
'
'        Next

    Next

    MovimentoOutros_Valida_Gravacao = SUCESSO

    Exit Function

Erro_MovimentoOutros_Valida_Gravacao:

    MovimentoOutros_Valida_Gravacao = gErr

    Select Case gErr

        Case 105943
            Call Rotina_ErroECF(vbOKOnly, ERRO_ADMINISTRADORA_NAO_PREENCHIDO_GRID, gErr, iIndice)

        Case 105944
            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_PREENCHIDO_GRID, gErr, iIndice)

        Case 105945
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORSANGRIA_NAO_INFORMADO_GRID, gErr, iIndice)

        Case 105946
            Call Rotina_ErroECF(vbOKOnly, ERRO_LINHA_REPETIDA, gErr)

        Case 111021
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 111022
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORSANGRIA_NAO_INFORMADO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163041)

    End Select

    Exit Function

End Function

Function MovimentoCaixa_Prepara_Exclusao(colMovimentosCaixa As Collection, Optional iTipoMovimento As Integer) As Long
'Função que Para Cada obj da Coleção adciona a esse movimento dizando q foi alterado o valor da Sangria

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_MovimentoCaixa_Prepara_Exclusao

    For Each objMovimentoCaixa In colMovimentosCaixa

        'Adciona o Tipo de Movimento a Coleção de Movimentos
        objMovimentoCaixa.iTipo = iTipoMovimento

    Next

    MovimentoCaixa_Prepara_Exclusao = SUCESSO

    Exit Function

Erro_MovimentoCaixa_Prepara_Exclusao:

    MovimentoCaixa_Prepara_Exclusao = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163042)

    End Select

    Exit Function

End Function

Function Move_Dados_Memoria(colMovimentosCaixa As Collection) As Long
'Função que Move os dados para a memoria

Dim lErro As Long
Dim iIndice As Integer
Dim objMovimentosCaixa As ClassMovimentoCaixa
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim dValor As Double

On Error GoTo Erro_Move_Dados_Memoria

    'verifica para cada linha do grid
    For iIndice = 1 To objGridOutros.iLinhasExistentes

        'Instancia um novo obj
        Set objMovimentosCaixa = New ClassMovimentoCaixa
    
        'Guarda em qual filial empresa que esta Trabalhando
        objMovimentosCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Guarda o valor da Sangria
        objMovimentosCaixa.dValor = StrParaDbl(GridOutros.TextMatrix(iIndice, iGrid_ValorSangria_Col))

        'Guardo o codigo do movimento
        objMovimentosCaixa.lNumMovto = StrParaLong(Codigo.Text)
        
        'Guarda o Tipo de Movimento
        objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_OUTROS
        
        For Each objAdmMeioPagtoCondPagto In gcolOutros

            If GridOutros.TextMatrix(iIndice, iGrid_Administradora_Col) = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto And GridOutros.TextMatrix(iIndice, iGrid_Parcelamento_Col) = objAdmMeioPagtoCondPagto.sNomeParcelamento Then

                'Guardo o Codigo da Admnistradora no Movimento Caixa
                objMovimentosCaixa.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto
                'Guarda o Código do Parcelamento da Linha
                objMovimentosCaixa.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento
                
                Exit For

            End If

        Next
        
        'Adciona a Coleção de ColMovimentosCaixa
        colMovimentosCaixa.Add objMovimentosCaixa

    Next
    
    dValor = StrParaDbl(ValorSangriaNaoDetalhado.Text)
    
    If dValor > 0 Then

        'Instancia novo Obj
        Set objMovimentosCaixa = New ClassMovimentoCaixa

        'Guarda Zero no Código do meio de Pagto não especificada
        objMovimentosCaixa.lNumMovto = StrParaLong(Codigo.Text)

        'Guarda em qual filial empresa que esta Trabalhando
        objMovimentosCaixa.iFilialEmpresa = giFilialEmpresa

        'Guarda o valor da Sangria
        objMovimentosCaixa.dValor = StrParaDbl(ValorSangriaNaoDetalhado.Text)
        
        'Guardo o Codigo do meio de Pagto no Movimento Caixa
        objMovimentosCaixa.iAdmMeioPagto = 0

        objMovimentosCaixa.iParcelamento = PARCELAMENTO_AVISTA

        'Guarda o Tipo de Movimento
        objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_OUTROS

        'Adciona a Coleção de ColMovimentosCaixa
        colMovimentosCaixa.Add objMovimentosCaixa

    End If

    Move_Dados_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_Memoria:

    Move_Dados_Memoria = gErr

    Select Case gErr

        Case 111025
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163043)

    End Select

    Exit Function

End Function

Function MovimentoOutros_Atualiza_Memoria(colMovimentosCaixa As Collection) As Long
'Função que Limpa a Coleção Global a Tela apos a Função de Gravação

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_MovimentoOutros_Atualiza_Memoria

    For Each objMovimentoCaixa In colMovimentosCaixa

        'Função que Atualiza os meios de Pagto Excluidos
        lErro = CF_ECF("MovimentoOutros_Atualiza_Memoria1", objMovimentoCaixa)
        If lErro <> SUCESSO Then gError 111027

        If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_OUTROS Then

            'Função que Retira de memória os Movimentos Excluidos
            lErro = MovimentoCaixa_Exclui_Memoria(objMovimentoCaixa)
            If lErro <> SUCESSO Then gError 111026

        Else
        
            'Adcionar a Coleção Global o objMovimento Caixa
            gcolMovimentosCaixa.Add objMovimentoCaixa
    
        End If

    Next

    MovimentoOutros_Atualiza_Memoria = SUCESSO

    Exit Function

Erro_MovimentoOutros_Atualiza_Memoria:

    MovimentoOutros_Atualiza_Memoria = gErr

    Select Case gErr

        Case 111026, 111027

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163044)

    End Select

    Exit Function

End Function

Function MovimentoCaixa_Exclui_Memoria(objMovimentoCaixa As ClassMovimentoCaixa) As Long
'Função que Exclui da Memória os Movimentos que Foram Alterados

Dim lErro As Long
Dim objMovimentoCaixaAux As New ClassMovimentoCaixa
Dim iIndice As Integer

On Error GoTo Erro_MovimentoCaixa_Exclui_Memoria

    For iIndice = gcolMovimentosCaixa.Count To 1 Step -1

        Set objMovimentoCaixaAux = gcolMovimentosCaixa.Item(iIndice)

        'Verifica se o movimento é o mesmo
        If objMovimentoCaixa.lNumMovto = objMovimentoCaixaAux.lNumMovto And objMovimentoCaixa.lSequencial = objMovimentoCaixaAux.lSequencial Then

            'Exclui o movimento da Coleção Global de MovimentosCaixa
            gcolMovimentosCaixa.Remove (iIndice)

            'Sai do Loop
            Exit For

        End If

    Next

    MovimentoCaixa_Exclui_Memoria = SUCESSO

    Exit Function

Erro_MovimentoCaixa_Exclui_Memoria:

    MovimentoCaixa_Exclui_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163045)

    End Select

    Exit Function

End Function

Function Caixa_Grava_Movimento(colMovimentosCaixa As Collection, Optional colcolInfoComplementar As Collection) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lSequencial As Long
Dim colRegistro As New Collection
Dim objOperador As New ClassOperador
Dim iIndice As Integer
Dim colInfoComplementar As New Collection
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim sNomeTerminal As String
Dim iCont As Integer
Dim objMovimentoCaixaAux As New ClassMovimentoCaixa
Dim bAchou As Boolean
Dim sMensagem As String
Dim lTamanho As Long
Dim sRetorno As String
Dim sArquivo As String

On Error GoTo Erro_Caixa_Grava_Movimento

    If giStatusSessao = SESSAO_ENCERRADA Then

        'Envia aviso perguntando se de seja Abrir sessão
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_DESEJA_ABRIR_SESSAO, giCodCaixa)

        If vbMsgRes = vbNo Then gError 111028

        'Função que Executa Abertura na Sessão
        lErro = CF_ECF("Operacoes_Executa_Abertura")
        If lErro <> SUCESSO Then gError 111029

    End If

    'Se for Necessário a Autorização do Gerente para abertura do Caixa
    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then

        'Chama a Tela de Senha
        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)

        'Sai de Função se a Tela de Login não Retornar ok
        If giRetornoTela <> vbOK Then gError 111030


    End If

    lTamanho = 255
    sRetorno = String(lTamanho, 0)

    'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
    Call GetPrivateProfileString(APLICACAO_DADOS, "DirDadosCC", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)
    
    'Se não encontrou
    If Len(Trim(sRetorno)) = 0 Or sRetorno = CStr(CONSTANTE_ERRO) Then gError 127097
    
    If right(sRetorno, 1) <> "\" Then sRetorno = sRetorno & "\"
    
    sArquivo = sRetorno & giCodEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOCC
    
    'Abre o arquivo de retorno
    Open sArquivo For Input Lock Read Write As #10

    'Função que Abre a Transação de Caixa, Identificador dentro do Caixa para um determinado MOVTO
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 111031

    lTamanho = 255
    sRetorno = String(lTamanho, 0)
        
    'Obtém a ultima transacao transferida
    Call GetPrivateProfileString(APLICACAO_DADOS, "UltimaTransacaoTransf", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
        
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)

    For Each objMovimentoCaixa In colMovimentosCaixa

        'se o numero da ultima transacao transferida ultrapassar o numero da transacao do movimento de caixa
        If objMovimentoCaixa.lSequencial <> 0 And StrParaLong(sRetorno) > objMovimentoCaixa.lSequencial Then gError 133849

        'Se Operador for Gerente
        objMovimentoCaixa.iGerente = objOperador.iCodigo

        lErro = Caixa_Grava_MovCx(objMovimentoCaixa, lSequencial)
        If lErro <> SUCESSO Then gError 105715

    Next

    lSequencial = lSequencial - 1
    
    'Fecha a Transação
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 111034

    Close #10

    Caixa_Grava_Movimento = SUCESSO

    Exit Function

Erro_Caixa_Grava_Movimento:

    Close #10

    Caixa_Grava_Movimento = gErr

    Select Case gErr
        
        Case 111028 To 111030
        
        Case 105715, 111031, 111034
            
        Case 133849
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163046)

    End Select

    Call CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)
    
    Exit Function

End Function

Function Caixa_Grava_MovCx(objMovimentoCaixa As ClassMovimentoCaixa, lSequencial As Long) As Long
'grava cada movimento de caixa passado como parametro

Dim lErro As Long
Dim colRegistro As New Collection
Dim sMensagem As String
Dim objMovCx As ClassMovimentoCaixa
Dim objTela As Object

On Error GoTo Erro_Caixa_Grava_MovCx

    'se for um movimento de exclusao de sangria ==> cria um novo movimento de caixa, grava no arquivão e deixa o que foi passado
    'como parametro sem alterar para permitir posteriormente retirar o movimento da memoria pelo sequencial original
    If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_OUTROS Then

        Set objMovCx = New ClassMovimentoCaixa
            
        lErro = CF_ECF("MovimentoCaixa_Copia", objMovimentoCaixa, objMovCx)
        If lErro <> SUCESSO Then gError 105712
    
    Else
    
        Set objMovCx = objMovimentoCaixa
    
    End If

    'Guarda o Sequencial no objmovimentoCaixa
    objMovCx.lSequencial = lSequencial

    lSequencial = lSequencial + 1

    'Guarda no objMovimentoCaixa os Dados que Serão Usados para a Geração do Movimento de Caixa
    lErro = CF_ECF("Move_DadosGlobais_Memoria", objMovCx)
    If lErro <> SUCESSO Then gError 111032

    'Funçao que Gera o Arquivo preparando para a gravação
    Call CF_ECF("MovimentoOutros_Gera_Log", colRegistro, objMovCx)

    'Função que Vai Gravar as Informações no Arquivo de Caixa
    lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
    If lErro <> SUCESSO Then gError 111033
    
    Set objTela = Me
    
    'para não ficar 3 movimentos com o mesmo Código(Numero de Movto) na Coleção gcolMovto
    If objMovCx.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_OUTROS Then
        
'        'Faz a sangria
'        lErro = CF_ECF("Sangria_AFRAC", objMovCx.dValor, sMensagem, -1)
'        If lErro <> SUCESSO Then gError 105713
'
'    Else
'        'Faz a sangria
'        lErro = CF_ECF("Sangria_AFRAC", objMovCx.dValor, sMensagem, 0)
'        If lErro <> SUCESSO Then gError 105714
'    End If

        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_OUTROS, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Exclusão Sangria Outros - Valor: " & Format(objMovCx.dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Exclusao Sangria - Outros")
        If lErro <> SUCESSO Then gError 117679

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Exclusao Sangria - Outros")
        If lErro <> SUCESSO Then gError 117680

        
    Else
        
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_OUTROS, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Sangria Outros - Valor: " & Format(objMovCx.dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Outros")
        If lErro <> SUCESSO Then gError 117681

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Outros")
        If lErro <> SUCESSO Then gError 117682

    End If

    Caixa_Grava_MovCx = SUCESSO
    
    Exit Function

Erro_Caixa_Grava_MovCx:

    Caixa_Grava_MovCx = gErr

    Select Case gErr

        Case 105712, 105713, 105714, 111032, 111033, 117679 To 117682

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163047)

    End Select
    
    Exit Function

End Function

'Function Move_DadosGlobais_Memoria(objMovimentoCaixa As ClassMovimentoCaixa) As Long
''Função que Move os Dados para a Memória
'
'Dim lErro As Long
'
'On Error GoTo Erro_Move_DadosGlobais_Memoria
'
'    'Guarda as Informações Globais no objMovimentosCaixa
'
'    objMovimentoCaixa.iFilialEmpresa = giFilialEmpresa
'
'    objMovimentoCaixa.iCaixa = giCodCaixa
'
'    objMovimentoCaixa.iCodOperador = giCodOperador
'
'    objMovimentoCaixa.dtDataMovimento = Date
'
'    objMovimentoCaixa.dHora = CDbl(Time)
'
'    Move_DadosGlobais_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_DadosGlobais_Memoria:
'
'    Move_DadosGlobais_Memoria = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163048)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Sub BotaoExcluir_Click()
'Botão que Exclui um Movimento de Sangria

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim vbMsgRes As VbMsgBoxResult
Dim lNumero As Long
Dim iTipoMovimento As Integer
Dim objMovCC As ClassMovimentoCaixa

On Error GoTo Erro_BotaoExcluir_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207978

    'Verifica se já foi executa a redução z para a data de hoje
    If gdtUltimaReducao = Date Then gError 111314
    
    'Verifica se o Codigo não foi Preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 111035

    lNumero = StrParaLong(Codigo.Text)

    'Verifica os Movimentos de Caixa para o Código em Questão
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, lNumero)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 111036
    
    'Senão encontrou o Movimento
    If lErro = 107850 Then gError 111037

    Set objMovCC = colMovimentosCaixa(1)

    If objMovCC.iTipo <> MOVIMENTOCAIXA_SANGRIA_OUTROS Then gError 86291
    
    'Pergunta se deseja Realmente Excluir o Movimento
    vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_EXCLUSAO_MOVIMENTOCAIXA, Codigo.Text)

    If vbMsgRes = vbNo Then gError 111038
    iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_OUTROS

    'Prepara os Movimentos para a Exclusão
    lErro = MovimentoCaixa_Prepara_Exclusao(colMovimentosCaixa, iTipoMovimento)
    If lErro <> SUCESSO Then gError 111039

    'Função que Grava a Exclusão de Outros
    lErro = Caixa_Grava_Movimento(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 111040
    
    'Atualiza os Dados na Memória
    lErro = MovimentoOutros_Atualiza_Memoria(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 111041

    'Função Que Limpa a Tela
    Call Limpa_Tela_MovimentoOutros

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 86291
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_OUTROS, gErr, StrParaDbl(Codigo.Text))
            
        Case 86292
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
        
        Case 111035
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 111037
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_INEXISTENTE, gErr, lNumero)

        Case 111036, 111038, 111039, 111040, 111041, 207978

        Case 111314
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163049)

    End Select

    Exit Sub

End Sub

Private Sub BotaoValoresEmCaixa_Click()
'Função que Verifica os Valores em Caixa

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim lNumero As Long

On Error GoTo Erro_BotaoValoresEmCaixa_Click


    'Limpa a Tela
    Call Limpa_Tela_MovimentoOutros
    
    'Traz os Valores em Caixa para a Tela
    lErro = Traz_ValoresEmCaixa_Tela()
    If lErro <> SUCESSO Then gError 111043

    'Função que Recalcula os Totáis
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 111045

    Exit Sub

Erro_BotaoValoresEmCaixa_Click:

    Select Case gErr

        Case 111042 To 111045

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163050)

    End Select

End Sub

Function Traz_ValoresEmCaixa_Tela() As Long
'Função que Traz Todos os valores de todos os movimentos relacionados com o meio de Pagto outros para a Tela independente de código

Dim lErro As Long
Dim iLinha As Integer
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Traz_ValoresEmCaixa_Tela

    'Verifica na Coleção Global
    For Each objAdmMeioPagtoCondPagto In gcolOutros

        'Verifica se o GridOutros não foi Especificado
        If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then

            LabelEmCaixaNaoDetalhadosValor.Caption = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
            ValorSangriaNaoDetalhado.Text = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))

        Else

            If objAdmMeioPagtoCondPagto.dSaldo <> 0 And objAdmMeioPagtoCondPagto.iAdmMeioPagto <> MEIO_PAGAMENTO_CONTRAVALE Then

                iLinha = iLinha + 1
                
                'Joga no Grid o nome do meio de Pagto
                GridOutros.TextMatrix(iLinha, iGrid_Administradora_Col) = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
                GridOutros.TextMatrix(iLinha, iGrid_Parcelamento_Col) = objAdmMeioPagtoCondPagto.sNomeParcelamento
                GridOutros.TextMatrix(iLinha, iGrid_ValorTotal_Col) = Format(objAdmMeioPagtoCondPagto.dSaldo, "Standard")
                GridOutros.TextMatrix(iLinha, iGrid_ValorSangria_Col) = Format(objAdmMeioPagtoCondPagto.dSaldo, "Standard")
                            
            End If
                
        End If

    Next

    objGridOutros.iLinhasExistentes = iLinha

    Traz_ValoresEmCaixa_Tela = SUCESSO

    Exit Function

Erro_Traz_ValoresEmCaixa_Tela:

    Traz_ValoresEmCaixa_Tela = gErr

    Select Case gErr

        Case 111045

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163051)

    End Select

    Exit Function

End Function

'Function Traz_ValoresEmCaixa_ValoresMovto_Tela(colMovimentosCaixa As Collection) As Long
''Função que Traz a tela os movimentos encontrados
'
'Dim lErro As Long
'Dim colGridOutros As New Collection
'Dim iIndice As Integer
'Dim objMovimentoCaixa As New ClassMovimentoCaixa
'Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
'Dim objAdmMeioPagtoCondPagtoAux As ClassAdmMeioPagtoCondPagto
'Dim objAdmMeioPagtoParcelasAux As ClassAdmMeioPagtoParcelas
'Dim objAdmMeioPagtoParcelas As ClassAdmMeioPagtoParcelas
'Dim iCont As Integer
'Dim iProcuraCombo As Integer
'Dim sNomeAdm As String
'Dim bAchou As Boolean
'
'On Error GoTo Erro_Traz_ValoresEmCaixa_ValoresMovto_Tela
'
'    'Copia Item a Item
'    For Each objAdmMeioPagtoCondPagto In gcolOutros
'
'        Set objAdmMeioPagtoCondPagtoAux = New ClassAdmMeioPagtoCondPagto
'
'        objAdmMeioPagtoCondPagtoAux.dDesconto = objAdmMeioPagtoCondPagto.dDesconto
'        objAdmMeioPagtoCondPagtoAux.dJuros = objAdmMeioPagtoCondPagto.dJuros
'        objAdmMeioPagtoCondPagtoAux.dSaldo = objAdmMeioPagtoCondPagto.dSaldo
'        objAdmMeioPagtoCondPagtoAux.dTaxa = objAdmMeioPagtoCondPagto.dTaxa
'        objAdmMeioPagtoCondPagtoAux.dValorMinimo = objAdmMeioPagtoCondPagto.dValorMinimo
'        objAdmMeioPagtoCondPagtoAux.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto
'        objAdmMeioPagtoCondPagtoAux.iFilialEmpresa = objAdmMeioPagtoCondPagto.iFilialEmpresa
'        objAdmMeioPagtoCondPagtoAux.iJurosParcelamento = objAdmMeioPagtoCondPagto.iJurosParcelamento
'        objAdmMeioPagtoCondPagtoAux.iNumParcelas = objAdmMeioPagtoCondPagto.iNumParcelas
'        objAdmMeioPagtoCondPagtoAux.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento
'        objAdmMeioPagtoCondPagtoAux.iParcelasRecebto = objAdmMeioPagtoCondPagto.iParcelasRecebto
'        objAdmMeioPagtoCondPagtoAux.iTipoCartao = objAdmMeioPagtoCondPagto.iTipoCartao
'        objAdmMeioPagtoCondPagtoAux.sNomeAdmMeioPagto = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
'        objAdmMeioPagtoCondPagtoAux.sNomeParcelamento = objAdmMeioPagtoCondPagto.sNomeParcelamento
'
'        For Each objAdmMeioPagtoParcelas In objAdmMeioPagtoCondPagto.colParcelas
'
'            Set objAdmMeioPagtoParcelasAux = New ClassAdmMeioPagtoParcelas
'
'            objAdmMeioPagtoParcelasAux.dPercRecebimento = objAdmMeioPagtoParcelas.dPercRecebimento
'            objAdmMeioPagtoParcelasAux.iAdmMeioPagto = objAdmMeioPagtoParcelas.iAdmMeioPagto
'            objAdmMeioPagtoParcelasAux.iFilialEmpresa = objAdmMeioPagtoParcelas.iFilialEmpresa
'            objAdmMeioPagtoParcelasAux.iIntervaloRecebimento = objAdmMeioPagtoParcelas.iIntervaloRecebimento
'            objAdmMeioPagtoParcelasAux.iParcela = objAdmMeioPagtoParcelas.iParcela
'            objAdmMeioPagtoParcelasAux.iParcelamento = objAdmMeioPagtoParcelas.iParcelamento
'
'            objAdmMeioPagtoCondPagtoAux.colParcelas.Add objAdmMeioPagtoParcelasAux
'
'        Next
'
'    colGridOutros.Add objAdmMeioPagtoCondPagtoAux
'
'    Next
'    bAchou = False
'    'Verifica para Cada Indice da Coleção
'    For iIndice = colGridOutros.Count To 1 Step -1
'
'        'Verifica para da Coleção de Moviementos de Caixa
'        For Each objMovimentoCaixa In colMovimentosCaixa
'            'Verifica se o Código do meio de Pagto é Igual a Zero
'            If objMovimentoCaixa.iAdmMeioPagto = 0 And objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_OUTROS And bAchou = False Then
'
'                'Instancia o objAdmMeioPagtoCondPagto para a apontar para Coleção de GridOutros
'                Set objAdmMeioPagtoCondPagto = colGridOutros.Item(iIndice)
'
'                'Preenche com o saldo do meio de pagto outros não especicado + Valor do Movto
'                LabelEmCaixaNaoDetalhadosValor.Caption = Format(objMovimentoCaixa.dValor + objAdmMeioPagtoCondPagto.dSaldo, "Standard")
'                ValorSangriaNaoDetalhado.Text = Format(objMovimentoCaixa.dValor, "standard")
'                bAchou = True
'
'            Else
'
'                'Instancia o objAdmMeioPagtoCondPagto para a apontar para Coleção de meios de pagto's outros
'                Set objAdmMeioPagtoCondPagto = colGridOutros.Item(iIndice)
'
'                'Verfica se o Movimento de Caixa está Relacionada com o  meio Outros na coleção Global de Outros
'                If objMovimentoCaixa.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto Then
'
'                    iCont = iCont + 1
'
'                    'Escreve no Grid o Nome do meio de Pagto
'                    GridOutros.TextMatrix(iCont, iGrid_Administradora_Col) = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
'
'
'                    'Escreve no Grid o Valor Total de saldo daquela meio de Pagto GridOutros
'                    GridOutros.TextMatrix(iCont, iGrid_ValorTotal_Col) = Format(objMovimentoCaixa.dValor + objAdmMeioPagtoCondPagto.dSaldo, "standard")
'
'                    'Escreve no Grid o Valor da Sangria do meio de pagto outros para o meio de Pagto
'                    GridOutros.TextMatrix(iCont, iGrid_ValorSangria_Col) = Format(objMovimentoCaixa.dValor, "standard")
'
'                End If
'
'            End If
'
'
'
'        Next
'
'        'Remove o GridOutros da Coleção de meios de Pagto Outros
'        colGridOutros.Remove (iIndice)
'
'    Next
'
'    'Atualiza o Numero de Linhas existentes no Grid
'    objGridOutros.iLinhasExistentes = iCont
'
'    'Função que Serve para Recalcular os Totais
'    lErro = Recalcula_Totais()
'    If lErro <> SUCESSO Then gError 111046
'
'    Traz_ValoresEmCaixa_ValoresMovto_Tela = SUCESSO
'
'    Exit Function
'
'Erro_Traz_ValoresEmCaixa_ValoresMovto_Tela:
'
'    Traz_ValoresEmCaixa_ValoresMovto_Tela = gErr
'
'    Select Case gErr
'
'        Case 111046
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163052)
'
'    End Select
'
'    Exit Function
'
'End Function


Private Sub BotaoLimpar_Click()
'Função Limpa a Tela

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Botaolimpar_Click

   If iAlterado = REGISTRO_ALTERADO Then

        'Envia aviso perguntando se os pagamentos devem ser aproveitados
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ALTERACAO_MOVIMENTOCAIXA1, Codigo.Text)

        If vbMsgRes = vbYes Then

            lErro = Gravar_Registro()
            If lErro <> SUCESSO Then gError 111047

        End If

    End If

    Call Limpa_Tela_MovimentoOutros
    
    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 111047

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163053)

    End Select

    Exit Sub

End Sub

Function Carrega_MovimentoOutros_NaoEsp() As Long
'Função que Carrega os Controles Relacionados aos Meios de Pagto Relacionado a Outros

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Carrega_MovimentoOutros_NaoEsp

    For Each objAdmMeioPagtoCondPagto In gcolOutros

        If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO Then

            LabelEmCaixaNaoDetalhadosValor.Caption = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
            ValorSangriaNaoDetalhado.Text = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))

        End If

    Next

    Carrega_MovimentoOutros_NaoEsp = SUCESSO

    Exit Function

Erro_Carrega_MovimentoOutros_NaoEsp:

    Carrega_MovimentoOutros_NaoEsp = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163054)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
'Botão que Executa o Fechamento da Tela

    Unload Me


End Sub

Private Sub LabelCodigo_Click()

Dim objMovimentoCaixa As New ClassMovimentoCaixa
    
    'Chama tela de MovimentoOutrosLosta
    Call Chama_TelaECF_Modal("MovimentoOutrosLista", objMovimentoCaixa)
    
    If Not (objMovimentoCaixa Is Nothing) Then
        'Verifica se o CodMovtoCaixa está preenchido e joga na coleção
        If objMovimentoCaixa.lNumMovto <> 0 Then
            Codigo.Text = objMovimentoCaixa.lNumMovto
            Call CodMovimentoOutros_Validate(False)
            
        End If
    End If
    
    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Função que Incrementa o Código Atravez da Tecla F2
Dim lErro As Long

On Error GoTo Erro_UserControl_KeyDown

    Select Case KeyCode

        Case KEYCODE_PROXIMO_NUMERO
            
            'Função que Incrementa o Código( Ultimo Movto + 1)
            If Not TrocaFoco(Me, BotaoProxNum) Then Exit Sub
            Call BotaoProxNum_Click

        Case KEYCODE_BROWSER

            Call LabelCodigo_Click

        Case vbKeyF4
            If Not TrocaFoco(Me, BotaoTrazer) Then Exit Sub
            Call BotaoTrazer_Click

        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoGravar) Then Exit Sub
            Call BotaoGravar_Click
            
        Case vbKeyF6
            If Not TrocaFoco(Me, BotaoExcluir) Then Exit Sub
            Call BotaoExcluir_Click
            
        Case vbKeyF7
            If Not TrocaFoco(Me, BotaoLimpar) Then Exit Sub
            Call BotaoLimpar_Click
            
        Case vbKeyF8
            If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
            Call BotaoFechar_Click

        Case vbKeyF9
            If Not TrocaFoco(Me, BotaoValoresEmCaixa) Then Exit Sub
            Call BotaoValoresEmCaixa_Click

    End Select

    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163055)

    End Select

    Exit Sub

End Sub

Sub CodMovimentoOutros_Validate(Cancel As Boolean)
'Função que Verifica se o Código Passado como parâmetro Existe na Coleção Globa de MOVTOCAIXA

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_CodMovimentoOutros_Validate

    'Verifica se existe movimento com o código passado
    For Each objMovimentoCaixa In gcolMovimentosCaixa
    
        If objMovimentoCaixa.lNumMovto = StrParaLong(Codigo.Text) Then
            
            'Função que traz o MovimentoOutros para a Tela
            lErro = Traz_MovimentoOutros_Tela(StrParaLong(Codigo.Text))
            If lErro <> SUCESSO Then gError 111048
            
            Exit For
        
        End If
    
    Next
    
    'Anula a Alteração
    iAlterado = 0
    
    Exit Sub
    
Erro_CodMovimentoOutros_Validate:
    
    Select Case gErr

        Case 111048

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163056)

    End Select

    Exit Sub

End Sub

Private Sub DesmembraMovto_Click()
'Função que Desmembra Log

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim colImfCompl As New Collection
Dim objMovimentosCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_DesmembraMovto_Click

    lErro = CF_ECF("Desmembra_MovimentosCaixa", colMovimentosCaixa, colImfCompl, TIPOREGISTROECF_MOVIMENTOCAIXA_OUTROS)
    If lErro <> SUCESSO Then gError 111057
    
    For Each objMovimentosCaixa In colMovimentosCaixa
        If objMovimentosCaixa.iTipo <> MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_OUTROS Then
            gcolMovimentosCaixa.Add colMovimentosCaixa
            gcolImfCompl.Add colImfCompl
        End If
    Next
    
    Exit Sub

Erro_DesmembraMovto_Click:
    
    Select Case gErr

        Case 111057

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163057)

    End Select

    Exit Sub


End Sub

