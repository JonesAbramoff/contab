VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl MovimentoTicket 
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   KeyPreview      =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   7035
   Begin VB.Frame FrameTicketsNaoDetalhados 
      Caption         =   "Ticket's Não Detalhados"
      Height          =   795
      Left            =   255
      TabIndex        =   15
      Top             =   5640
      Width           =   6630
      Begin VB.TextBox ValorSangriaNaoDetalhado 
         Height          =   315
         Left            =   3975
         TabIndex        =   16
         Top             =   315
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
         TabIndex        =   19
         Top             =   375
         Width           =   840
      End
      Begin VB.Label LabelEmCaixaNaoDetalhadosValor 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   315
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
         TabIndex        =   17
         Top             =   375
         Width           =   720
      End
   End
   Begin VB.Frame FrameTicketsDetalhados 
      Caption         =   "Ticket's Detalhados"
      Height          =   4425
      Left            =   225
      TabIndex        =   11
      Top             =   1080
      Width           =   6630
      Begin VB.ComboBox Parcelamento 
         Height          =   315
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   1590
      End
      Begin VB.CommandButton BotaoValoresEmCaixa 
         Height          =   585
         Left            =   90
         Picture         =   "MovimentoTicket.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "F9 - Traz para a tela os boletos em caixa que podem ser ""sangrados""."
         Top             =   3765
         Width           =   1680
      End
      Begin VB.ComboBox Administradora 
         Height          =   315
         Left            =   525
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   1545
      End
      Begin MSMask.MaskEdBox ValorSangria 
         Height          =   300
         Left            =   4905
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorTotal 
         Height          =   300
         Left            =   3675
         TabIndex        =   14
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridTickets 
         Height          =   3300
         Left            =   150
         TabIndex        =   4
         Top             =   315
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   5821
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
         Left            =   4995
         TabIndex        =   23
         Top             =   3705
         Width           =   1215
      End
      Begin VB.Label LabelEmCaixaValor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3765
         TabIndex        =   22
         Top             =   3705
         Width           =   1215
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
         Left            =   3105
         TabIndex        =   21
         Top             =   3765
         Width           =   600
      End
   End
   Begin VB.Frame FrameIdentificacao 
      Caption         =   "Identificação"
      Height          =   870
      Left            =   240
      TabIndex        =   9
      Top             =   60
      Width           =   3645
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1635
         Picture         =   "MovimentoTicket.ctx":2D5A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   345
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
         Picture         =   "MovimentoTicket.ctx":2E44
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "F4 - Exibe na tela o movimento com o código informado."
         Top             =   210
         Width           =   1440
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   795
         TabIndex        =   1
         Top             =   330
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
         Left            =   105
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   390
         Width           =   660
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4695
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "MovimentoTicket.ctx":5B0E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "F8 - Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1080
         Picture         =   "MovimentoTicket.ctx":5C8C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "F7 - Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "MovimentoTicket.ctx":61BE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "F6 - Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "MovimentoTicket.ctx":6348
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "F5 - Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   1710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   525
      Left            =   195
      TabIndex        =   25
      Top             =   795
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   926
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"MovimentoTicket.ctx":64A2
   End
End
Attribute VB_Name = "MovimentoTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Declarações Globais

Dim objGridTickets As AdmGrid
Dim iAlterado As Integer
Dim iAdmnistradora_Alterado As Integer
Dim iValorSangria_Alterado As Integer
Dim iGrid_ValorTotal_Col As Integer
Dim iGrid_ValorSangria_Col As Integer
Dim iGrid_Administradora_Col As Integer
Dim iGrid_Parcelamento_Col As Integer
Dim glProxNumAuto As Long
Dim gcolImfCompl As New Collection


'Property Variables:
Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Sangria Ticket"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Movimento Tickets"

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

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo)

End Sub

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

    'Instanciar o objGridTickets para apontar para uma posição de memória
    Set objGridTickets = New AdmGrid

    'Inicialização de Grid Ticketss
    lErro = Inicializa_GridTickets(objGridTickets)
    If lErro <> SUCESSO Then gError 107910

    'Função que Carrega as Admnistradoras de Meio de Pagto
    lErro = Carrega_Combo_Administradora()
    If lErro <> SUCESSO Then gError 107912

    Call BotaoValoresEmCaixa_Click

    'Indica que o carregamento da tela aconteceu com sucesso
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        'Erros tratados na rotina chamada
        Case 107910 To 107912

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163061)

    End Select

    Exit Sub

End Sub

Function Inicializa_GridTickets(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridTickets

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
    objGridInt.objGrid = GridTickets

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_LINHAS_GRID

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Largura da primeira coluna
    GridTickets.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
   
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridTickets = SUCESSO

    Exit Function

Erro_Inicializa_GridTickets:

    Inicializa_GridTickets = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163062)

    End Select

    Exit Function

End Function

Function Recalcula_Totais() As Long
'Calcula o Valor Total de Tickets especificados e não especificados

Dim lErro As Long
Dim iIndice As Integer
Dim dValorTotalTickets As Double
Dim dValorSangria As Double
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Recalcula_Totais

    'Para todos os Cheque do Grid
    For iIndice = 1 To objGridTickets.iLinhasExistentes

        'Acumula o Valor Total dos Ticktes
        dValorTotalTickets = dValorTotalTickets + StrParaDbl(GridTickets.TextMatrix(iIndice, iGrid_ValorTotal_Col))

        'Acumula o Valor Total da Sangria
        dValorSangria = dValorSangria + StrParaDbl(GridTickets.TextMatrix(iIndice, iGrid_ValorSangria_Col))

    Next

    
    'Exibe o valor total tando de Sangria como o valor total de Ticketss que existem no caixa
    LabelEmCaixaValor.Caption = Format(dValorTotalTickets, "Standard")
    LabelSangriaValor.Caption = Format(dValorSangria, "Standard")


    Recalcula_Totais = SUCESSO

    Exit Function

Erro_Recalcula_Totais:

    Recalcula_Totais = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163063)

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
    
                If Len(Trim(GridTickets.TextMatrix(iLinha, iGrid_Administradora_Col))) = 0 Then

    
                    objControl.Enabled = False
    
                Else
    
                    objControl.Enabled = True
    
                    'Carrega a Combo de Parcelamentos
                    lErro = Carrega_Combo_Parcelamento(iLinha)
                    If lErro <> SUCESSO Then gError 105947
    
                End If

    
            'Se o campo for valor sangria
            Case ValorSangria.Name
    
                'Verifca se o Campo valorTotal está Preenchido
                 If Len(Trim(GridTickets.TextMatrix(iLinha, iGrid_ValorTotal_Col))) <> 0 Then
    
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

        Case 105947

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163064)

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
    sNomeParcelamento = GridTickets.TextMatrix(iLinha, iGrid_Parcelamento_Col)

    'Atribui o Nome da Admnistradora Selecionada
    sNomeAdmnistradora = GridTickets.TextMatrix(iLinha, iGrid_Administradora_Col)
    
    'Limpa a Combo
    Parcelamento.Clear
    
    'Incluir todos os Parcelamentos cadastratrados na Coleção Glogal de Parcelamentos Referenciando a Admnistradora Referenciada
    For Each objAdmMeioPagtoCondPagto In gcolTicket

         If sNomeAdmnistradora = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto Then

            Parcelamento.AddItem objAdmMeioPagtoCondPagto.sNomeParcelamento
            Parcelamento.ItemData(Parcelamento.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento

        End If
    Next

    If Parcelamento.ListCount > 0 Then Parcelamento.AddItem " "

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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163065)

    End Select

    Exit Function

End Function

Function Carrega_Combo_Administradora() As Long
'Função que Carrega a Combo de Admnistradoras

Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim iIndice As Integer
Dim bAchou As Boolean

On Error GoTo Erro_Carrega_Combo_Administradora

    'Adiciona todas as Admnistradaoras lidas na Coleção Global na Combo
    For Each objAdmMeioPagtoCondPagto In gcolTicket
        
        If objAdmMeioPagtoCondPagto.iAdmMeioPagto <> 0 Then
        
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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163066)

    End Select

    Exit Function

End Function

Private Sub BotaoProxNum_Click()
'Botão que Gera um Próximo Numero para Movto

Dim lErro As Long
Dim lNumero As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Função que Gera o Próximo Código para a Tela de Sangria de Ticketss
    lErro = CF_ECF("Caixa_Obtem_NumAutomatico", lNumero)
    If lErro <> SUCESSO Then gError 107913

    'Exibir o Numero na Tela
    Codigo.Text = lNumero

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 107913

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163067)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazer_Click()
'Função que chama a função que preenche o grid

Dim lErro As Long
Dim lCodigo As String

On Error GoTo Erro_botaoTrazer_click

    'Verifica se o código não está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107917

    lCodigo = StrParaLong(Codigo.Text)
    
    Call Limpa_Tela_MovimentoTickets
    
    Codigo.Text = lCodigo
    
    'Chama a função que preenche o grid
    lErro = Traz_MovimentoTickets_Tela(StrParaLong(Codigo.Text))
    If lErro <> SUCESSO Then gError 107918

    'Anula a Alteração
    iAlterado = 0
    
    Exit Sub

Erro_botaoTrazer_click:

    Select Case gErr

        Case 107917
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 107918
            'Erro tradado Dentro da Função que Foi Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163068)

    End Select

    Exit Sub

End Sub

Function Traz_MovimentoTickets_Tela(lNumero As Long) As Long
'Função que Trar o Movimento Tickets para a tela

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim objMovimentoCaixa As New ClassMovimentoCaixa
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim iLinha As Integer

On Error GoTo Erro_Traz_MovimentoTickets_Tela

    Call Limpa_Tela_MovimentoTickets
    
    Codigo.Text = lNumero
    
    'Função que Lê os Movimentos de Caixa
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, lNumero)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 107919
    
    'Não existe Movimento com este Código
    If lErro = 107850 Then gError 107921

    For Each objMovimentoCaixa In colMovimentosCaixa
        
        If objMovimentoCaixa.iTipo <> MOVIMENTOCAIXA_SANGRIA_TICKET Then gError 105742
            
        For Each objAdmMeioPagtoCondPagto In gcolTicket
            'Verifica se o Código da Administradora é igual ao Código da Adimistradora do Movto
            If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovimentoCaixa.iAdmMeioPagto Then

                'Verifica se o codigo da admnistradora é igual a zero
                If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        
                        'Exibe o Saldo + Valor do Movto
                        LabelEmCaixaNaoDetalhadosValor.Caption = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo + objMovimentoCaixa.dValor, "standard"))
                        'Coloca o Valor do Movto
                        ValorSangriaNaoDetalhado.Text = Format(objMovimentoCaixa.dValor, "Standard")
                        
                        Exit For
                        
                'Senão Preenche o Grid, se o codigo da Adm for igual ao Código da Adm do Movto
                ElseIf objAdmMeioPagtoCondPagto.iParcelamento = objMovimentoCaixa.iParcelamento Then

                    iLinha = iLinha + 1

                    GridTickets.TextMatrix(iLinha, iGrid_Administradora_Col) = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
                    GridTickets.TextMatrix(iLinha, iGrid_Parcelamento_Col) = objAdmMeioPagtoCondPagto.sNomeParcelamento
                    GridTickets.TextMatrix(iLinha, iGrid_ValorTotal_Col) = Format(objAdmMeioPagtoCondPagto.dSaldo + objMovimentoCaixa.dValor, "standard")
                    GridTickets.TextMatrix(iLinha, iGrid_ValorSangria_Col) = Format(objMovimentoCaixa.dValor, "standard")
                    Exit For
                    
                End If

            End If

        Next
    
    Next
    
    'atualizar Linhas Existentas
    objGridTickets.iLinhasExistentes = iLinha

    'Função que Atualiza os Totais
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107920
    
    Traz_MovimentoTickets_Tela = SUCESSO

    Exit Function

Erro_Traz_MovimentoTickets_Tela:

    Traz_MovimentoTickets_Tela = gErr

    Select Case gErr

        Case 105742
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_SANGRIA_TICKET, gErr, lNumero)

        Case 107919, 107920

        Case 107921
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_INEXISTENTE, gErr, lNumero)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163069)

    End Select

    Exit Function

End Function

Private Sub GridTickets_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridTickets, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridTickets, iAlterado)
    End If

End Sub

Private Sub GridTickets_EnterCell()

    'Parametro não opcional
    Call Grid_Entrada_Celula(objGridTickets, iAlterado)

End Sub

Private Sub GridTickets_GotFocus()

    Call Grid_Recebe_Foco(objGridTickets)

End Sub

Private Sub GridTickets_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridTickets_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridTickets)

    If KeyCode = vbKeyDelete Then
        
        'Função que recalcula os totais no grid
        lErro = Recalcula_Totais()
        If lErro <> SUCESSO Then gError 105746
        
    End If

    Exit Sub

Erro_GridTickets_KeyDown:

    Select Case gErr

        Case 105746

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163070)

    End Select

    Exit Sub

End Sub

Private Sub GridTickets_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridTickets, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridTickets, iAlterado)
    End If

End Sub

Private Sub GridTickets_LeaveCell()

    Call Saida_Celula(objGridTickets)

End Sub

Private Sub GridTickets_LostFocus()

    Call Grid_Libera_Foco(objGridTickets)

End Sub
Private Sub GridTickets_RowColChange()

    Call Grid_RowColChange(objGridTickets)

End Sub

Private Sub GridTickets_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridTickets)

End Sub

Private Sub GridTickets_Scroll()

    Call Grid_Scroll(objGridTickets)

End Sub

Private Sub Administradora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Administradora_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTickets)


End Sub

Private Sub Administradora_KeyPress(KeyAscii As Integer)


    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTickets)

End Sub

Private Sub Administradora_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTickets.objControle = Administradora
    lErro = Grid_Campo_Libera_Foco(objGridTickets)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Parcelamento_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Parcelamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTickets)

End Sub

Private Sub Parcelamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTickets)

End Sub

Private Sub Parcelamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTickets.objControle = Parcelamento
    lErro = Grid_Campo_Libera_Foco(objGridTickets)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorTotal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorTotal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTickets)

End Sub

Private Sub ValorTotal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTickets)

End Sub

Private Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTickets.objControle = ValorTotal
    lErro = Grid_Campo_Libera_Foco(objGridTickets)

    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorSangria_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorSangria_Alterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorSangria_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTickets)

End Sub

Private Sub ValorSangria_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTickets)

End Sub

Private Sub ValorSangria_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTickets.objControle = ValorSangria
    lErro = Grid_Campo_Libera_Foco(objGridTickets)

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

            'Administradora
            Case iGrid_Administradora_Col
                lErro = Saida_Celula_Administradora(objGridInt)
                If lErro <> SUCESSO Then gError 107922

            'Parcelamento
            Case iGrid_Parcelamento_Col
                lErro = Saida_Celula_Parcelamento(objGridInt)
                If lErro <> SUCESSO Then gError 105948
            
            'ValorSangria
            Case iGrid_ValorSangria_Col
                lErro = Saida_Celula_ValorSangria(objGridInt)
                If lErro <> SUCESSO Then gError 107923

        End Select

    'Função que Finaliza a Saida de Celula
    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107924

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 105948, 107922 To 107924
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163071)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Administradora(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim iAchou As Integer

On Error GoTo Erro_Saida_Celula_Administradora

    Set objGridInt.objControle = Administradora
    
    If Administradora.Text <> GridTickets.TextMatrix(GridTickets.Row, iGrid_Administradora_Col) Then

        'Limpa o Grid na Linha Relacionada a Administradora em Questão na Coluna Relacionada ao Parcelamento
        GridTickets.TextMatrix(GridTickets.Row, iGrid_Parcelamento_Col) = ""
        
        'Limpa o Grid na Linha Relacionada a Administradora em Questão na Coluna Relacionada ao ValorTotal
        GridTickets.TextMatrix(GridTickets.Row, iGrid_ValorTotal_Col) = ""

        'Limpa o Grid na Linha Relacionada a Administradora em Questão na Coluna Relacionada ao Valor da Sangria
        GridTickets.TextMatrix(GridTickets.Row, iGrid_ValorSangria_Col) = ""

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107925

    'Verifica se célula do Grid que identifica o campo Admnistradora esta preenchido
    If Len(Trim(GridTickets.TextMatrix(GridTickets.Row, iGrid_Administradora_Col))) <> 0 Then

        lErro = Recalcula_Totais()
        If lErro <> SUCESSO Then gError 107926

        'Acrescenta uma linha no Grid se for o caso
        If GridTickets.Row - GridTickets.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    Saida_Celula_Administradora = SUCESSO

    Exit Function

Erro_Saida_Celula_Administradora:

    Saida_Celula_Administradora = gErr

    Select Case gErr

        Case 107925, 107926

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163072)

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

    If Parcelamento.Text <> GridTickets.TextMatrix(GridTickets.Row, iGrid_Parcelamento_Col) Then

        'Verifica se o Parcelamento esta selecionada ou foi Alterada
        If Len(Trim(Parcelamento.Text)) = 0 Then
    
            'Limpa o Campo Valor
            GridTickets.TextMatrix(GridTickets.Row, iGrid_ValorTotal_Col) = ""
    
            'Limpa o Campo Relacionado a Sangria
            GridTickets.TextMatrix(GridTickets.Row, iGrid_ValorSangria_Col) = ""
    
        'Se o campo Parcelamento foi Alterado e se o Campo Relacionado a Terminal Já Estivar Preenchido
        Else
    
            'Para ver se existe duplicidade no Grig
            For iLinha = 1 To objGridTickets.iLinhasExistentes
    
                If iLinha <> GridTickets.Row Then
    
                    If GridTickets.TextMatrix(iLinha, iGrid_Administradora_Col) = GridTickets.TextMatrix(GridTickets.Row, iGrid_Administradora_Col) And _
                    GridTickets.TextMatrix(iLinha, iGrid_Parcelamento_Col) = Parcelamento.Text Then gError 105949
    
                End If
    
            Next
            
    
            'Procura na Coleção de Tickets a Tupla correspondente e preenche o Grid
            For Each objAdmMeioPagtoCondPagto In gcolTicket
    
                'Verifica se o Nome da Admnistradora é Igual e o parcelamento Selecionada é Igual ao da Tupla Admnistradora + parcelmento
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = GridTickets.TextMatrix(GridTickets.Row, iGrid_Administradora_Col) And objAdmMeioPagtoCondPagto.sNomeParcelamento = Parcelamento.Text Then
    
                        'Preenche o Grid com o Valor Total refente aos Tickets em Questão
                        GridTickets.TextMatrix(GridTickets.Row, iGrid_ValorTotal_Col) = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
                        GridTickets.TextMatrix(GridTickets.Row, iGrid_ValorSangria_Col) = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
                        iAchou = 1
                        Exit For
    
                    End If
    
            Next
    
            If iAchou = 0 Then gError 105950
    
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 105951

    'Acrescenta uma linha no Grid se for o caso
    If GridTickets.Row - GridTickets.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If

    'Função que recalcula os totais no grid
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 105952

    Saida_Celula_Parcelamento = SUCESSO

    Exit Function

Erro_Saida_Celula_Parcelamento:

    Saida_Celula_Parcelamento = gErr

    Select Case gErr

        Case 105949
            Call Rotina_ErroECF(vbOKOnly, ERRO_LINHA_REPETIDA, gErr)

        Case 105950
            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_EXISTENTE1, gErr, Parcelamento.Text)

        Case 105951, 105952

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163073)

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

    'Verifica se oValor da Sangria esta Selecionado ou foi Alterada
    If Len(Trim(ValorSangria.Text)) <> 0 Then

        'Verifica se o Valor Digitado é Valido
        lErro = Valor_NaoNegativo_Critica(ValorSangria.Text)
        If lErro <> SUCESSO Then gError 107927

        'Verifica se o valor da sangria é maior que o valor total se for Erro
        If StrParaDbl(ValorSangria.Text) > StrParaDbl(GridTickets.TextMatrix(GridTickets.Row, iGrid_ValorTotal_Col)) Then gError 107928


    End If

    'Adcionar ao Grid
    GridTickets.TextMatrix(GridTickets.Row, iGrid_ValorSangria_Col) = ValorSangria.Text

    'Função que recalcula os totais no grid
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107929

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 107930

    'Acrescenta uma linha no Grid se for o caso
    If GridTickets.Row - GridTickets.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If


    Saida_Celula_ValorSangria = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorSangria:

    Saida_Celula_ValorSangria = gErr

    Select Case gErr

        Case 107927
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 107928
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORSANGRIA_NAO_DISPONIVEL_GRID, gErr, ValorSangria.Text, GridTickets.TextMatrix(GridTickets.Row, iGrid_ValorTotal_Col), GridTickets.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 107929
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 120083
            Call Rotina_ErroECF(vbOKOnly, ERRO_LINHA_REPETIDA, gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 107930
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163074)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub ValorSangriaNaoDetalhado_Validate(Cancel As Boolean)
'Função que valida os dados no campo valor de sangria em cartões de débito na conta

Dim lErro As Long

On Error GoTo Erro_ValorSangriaNaoDetalhado_Validate

    'Verifica se o campo esta preenchido
    If Len(Trim(ValorSangriaNaoDetalhado.Text)) <> 0 Then

        'se esta preenchido então verificar o valor
        lErro = Valor_NaoNegativo_Critica(ValorSangriaNaoDetalhado.Text)
        If lErro <> SUCESSO Then gError 107931

        'Verifica se o valor da sangria é maior do que o valor de Tickets não especificados
        If StrParaDbl(ValorSangriaNaoDetalhado.Text) > StrParaDbl(LabelEmCaixaNaoDetalhadosValor.Caption) Then gError 107932

        ValorSangriaNaoDetalhado.Text = Format(ValorSangriaNaoDetalhado.Text, "standard")

    End If

    Exit Sub

Erro_ValorSangriaNaoDetalhado_Validate:

    Cancel = True

    Select Case gErr

        Case 107931

        Case 107932
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORSANGRIA_NAO_DISPONIVEL, gErr, ValorSangriaNaoDetalhado.Text, giCodCaixa)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163075)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_MovimentoTickets() As Long
'Função que Limpa a Tela

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_MovimentoTickets

    'Limpa os Controles básico da Tela
    Call Limpa_Tela(Me)

    'Limpa Grid
    Call Grid_Limpa(objGridTickets)
    
    'Limpa os Labes da Tela
    LabelEmCaixaValor.Caption = Format(0, "standard")
    LabelEmCaixaNaoDetalhadosValor.Caption = Format(0, "standard")
    LabelSangriaValor.Caption = Format(0, "standard")

    iAlterado = 0
    
    Limpa_Tela_MovimentoTickets = SUCESSO

    Exit Function

Erro_Limpa_Tela_MovimentoTickets:

    Limpa_Tela_MovimentoTickets = gErr

    Select Case gErr

        Case 107981
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163076)

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
    If lErro <> SUCESSO Then gError 207979

    'Função que efeuara a gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 107933

    'Função que Limpa a Tela
    lErro = Limpa_Tela_MovimentoTickets
    If lErro <> SUCESSO Then gError 107934
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 107933, 107934, 207979

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163077)

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

    'Função que Valida a Gravação, se existem linhas repetidas no Grid
    lErro = MovimentoTicket_Valida_Gravacao(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107935

    'Transforma o Codigo Texto para Long
    lNumMovto = StrParaLong(Codigo.Text)

    'Lê os Movimentos de Caixa da coleção global e carrega no coleção local a tela para o numero do Movto
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, lNumMovto)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 107936

    'Verifica se já existe um movimento para o código referido, Verifica se é Alteração
    If colMovimentosCaixa.Count > 0 Then
        
        Set objMovCC = colMovimentosCaixa(1)
    
        If objMovCC.iTipo <> MOVIMENTOCAIXA_SANGRIA_TICKET Then gError 86291
        
        iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_TICKET

        'Envia aviso perguntando se deseja atualizar o movimemtos
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ALTERACAO_MOVIMENTOCAIXA, Codigo.Text)

        'Se a Reposta for Negativa
        If vbMsgRes = vbNo Then gError 107937

        'Função que Faz a Alteração na Sangria de Ticket Previamente Executada, adciona o iTipoMovimento
        lErro = MovimentoCaixa_Prepara_Exclusao(colMovimentosCaixa, iTipoMovimento)
        If lErro <> SUCESSO Then gError 107938

    End If

    'Move os Dados da Sangria para a Memoria
    lErro = Move_Dados_Memoria(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107939
    
    'Função que Grava os movimentos em Arquivos
    lErro = Caixa_Grava_Movimento(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107940

    'Atualiza os Dados da Memoria
    lErro = MovimentoTicket_Atualiza_Memoria(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107941

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

   Select Case gErr
        
        Case 86291
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_TICKET, gErr, StrParaDbl(Codigo.Text))
            
        Case 86292
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)

        Case 107935 To 107941

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163078)

    End Select

End Function

Function MovimentoTicket_Valida_Gravacao(colMovimentosCaixa As Collection) As Long
'Função que Valida a Gravação

Dim lErro As Long
Dim iIndice As Integer
Dim dValor As Double
Dim dValorAtual As Double
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim objMovCaixa As ClassMovimentoCaixa
Dim iCont As Integer

On Error GoTo Erro_MovimentoTicket_Valida_Gravacao

    'Verifica se o código Foi Preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107942

    'Verifica se não existem Linhas no Grid e as Label's não estejam preenchidas
    If objGridTickets.iLinhasExistentes = 0 And Len(Trim(LabelEmCaixaValor.Caption)) = 0 And Len(Trim(LabelSangriaValor.Caption)) = 0 And Len(Trim(ValorSangriaNaoDetalhado.Text)) Then gError 107943

    'Verifica se no Grid o Campo Referente a Sangria Esta Preenchido
    For iIndice = 1 To objGridTickets.iLinhasExistentes

        If Len(GridTickets.TextMatrix(iIndice, iGrid_Administradora_Col)) = 0 Then gError 105953
        
        If Len(GridTickets.TextMatrix(iIndice, iGrid_Parcelamento_Col)) = 0 Then gError 105954
        
        If StrParaDbl(GridTickets.TextMatrix(iIndice, iGrid_ValorSangria_Col)) = 0 Then gError 105955

        'Para ver se existe duplicidade no Grig
        For iCont = 1 To objGridTickets.iLinhasExistentes

            If iCont <> iIndice Then

                If GridTickets.TextMatrix(iCont, iGrid_Administradora_Col) = GridTickets.TextMatrix(iIndice, iGrid_Administradora_Col) And _
                GridTickets.TextMatrix(iCont, iGrid_Parcelamento_Col) = GridTickets.TextMatrix(iIndice, iGrid_Parcelamento_Col) Then gError 105956
            
            End If

        Next

    Next

    MovimentoTicket_Valida_Gravacao = SUCESSO

    Exit Function

Erro_MovimentoTicket_Valida_Gravacao:

    MovimentoTicket_Valida_Gravacao = gErr

    Select Case gErr

        Case 105953
            Call Rotina_ErroECF(vbOKOnly, ERRO_ADMINISTRADORA_NAO_PREENCHIDO_GRID, gErr, iIndice)

        Case 105954
            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_PREENCHIDO_GRID, gErr, iIndice)

        Case 105955
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORSANGRIA_NAO_INFORMADO_GRID, gErr, iIndice)

        Case 105956
            Call Rotina_ErroECF(vbOKOnly, ERRO_LINHA_REPETIDA, gErr)

        Case 107943
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORSANGRIA_NAO_INFORMADO, gErr)
        
        Case 107942
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163079)

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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163080)

    End Select

    Exit Function

End Function

Function Move_Dados_Memoria(colMovimentosCaixa As Collection) As Long
'Função que Move os dados para a memoria

Dim lErro As Long
Dim iIndice As Integer
Dim objMovimentosCaixa As ClassMovimentoCaixa
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim dValor As Double

On Error GoTo Erro_Move_Dados_Memoria

    'verifica para cada linha do grid
    For iIndice = 1 To objGridTickets.iLinhasExistentes

        'Instancia um novo obj
        Set objMovimentosCaixa = New ClassMovimentoCaixa
    
        'Guarda em qual filial empresa que esta Trabalhando
        objMovimentosCaixa.iFilialEmpresa = giFilialEmpresa
        
        'Guarda o valor da Sangria
        objMovimentosCaixa.dValor = StrParaDbl(GridTickets.TextMatrix(iIndice, iGrid_ValorSangria_Col))

        'Guardo o codigo do movimento
        objMovimentosCaixa.lNumMovto = StrParaLong(Codigo.Text)
        
        For Each objAdmMeioPagtoCondPagto In gcolTicket

            If GridTickets.TextMatrix(iIndice, iGrid_Administradora_Col) = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto And GridTickets.TextMatrix(iIndice, iGrid_Parcelamento_Col) = objAdmMeioPagtoCondPagto.sNomeParcelamento Then

                'Guardo o Codigo da Admnistradora no Movimento Caixa
                objMovimentosCaixa.iAdmMeioPagto = objAdmMeioPagtoCondPagto.iAdmMeioPagto
                'Guarda o Código do Parcelamento da Linha
                objMovimentosCaixa.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento
                

                Exit For

            End If

        Next
        
        'Guarda o Tipo de Movimento
        objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_TICKET
        
        'Adciona a Coleção de ColMovimentosCaixa
        colMovimentosCaixa.Add objMovimentosCaixa

    Next
        
    dValor = StrParaDbl(ValorSangriaNaoDetalhado.Text)
    
    If dValor > 0 Then

        'Instancia novo Obj
        Set objMovimentosCaixa = New ClassMovimentoCaixa

        'Guarda Zero no Código do Admnistradora não especificada
        objMovimentosCaixa.lNumMovto = StrParaLong(Codigo.Text)

        'Guarda em qual filial empresa que esta Trabalhando
        objMovimentosCaixa.iFilialEmpresa = giFilialEmpresa

        'Guarda o valor da Sangria
        objMovimentosCaixa.dValor = StrParaDbl(ValorSangriaNaoDetalhado.Text)

        'Guardo o Codigo da Admnistradora no Movimento Caixa
        objMovimentosCaixa.iAdmMeioPagto = 0
        
        objMovimentosCaixa.iParcelamento = PARCELAMENTO_AVISTA

        'Guarda o Tipo de Movimento
        objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_SANGRIA_TICKET
        
        'Adciona a Coleção de ColMovimentosCaixa
        colMovimentosCaixa.Add objMovimentosCaixa

    End If

    Move_Dados_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_Memoria:

    Move_Dados_Memoria = gErr

    Select Case gErr

        Case 107946
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163081)

    End Select

    Exit Function

End Function

Function MovimentoTicket_Atualiza_Memoria(colMovimentosCaixa As Collection) As Long
'Função que Limpa a Coleção Global a Tela apos a Função de Gravação

Dim lErro As Long
Dim objMovimentosCaixa As New ClassMovimentoCaixa
Dim iIndice As Integer
On Error GoTo Erro_MovimentoTicket_Atualiza_Memoria

    For Each objMovimentosCaixa In colMovimentosCaixa

        'Função que Atualiza os Boletos Excluidos
        lErro = CF_ECF("MovimentoTicket_Atualiza_Memoria1", objMovimentosCaixa)
        If lErro <> SUCESSO Then gError 107948

        If objMovimentosCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_TICKET Then


            'Função que Retira de memória os Movimentos Excluidos
            lErro = MovimentoCaixa_Exclui_Memoria(objMovimentosCaixa)
            If lErro <> SUCESSO Then gError 107947

        Else
        
            'Adcionar a Coleção Global o objMovimento Caixa
            gcolMovimentosCaixa.Add objMovimentosCaixa
    
        End If

    Next

    MovimentoTicket_Atualiza_Memoria = SUCESSO

    Exit Function

Erro_MovimentoTicket_Atualiza_Memoria:

    MovimentoTicket_Atualiza_Memoria = gErr

    Select Case gErr

        Case 107947, 107948

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163082)

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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163083)

    End Select

    Exit Function

End Function

Function Caixa_Grava_Movimento(colMovimentosCaixa As Collection) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lSequencial As Long
Dim colRegistro As New Collection
Dim objOperador As New ClassOperador
Dim iIndice As Integer
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

        If vbMsgRes = vbNo Then gError 107957

        'Função que Executa Abertura na Sessão
        lErro = CF_ECF("Operacoes_Executa_Abertura")
        If lErro <> SUCESSO Then gError 107958

    End If

    'Se for Necessário a Autorização do Gerente para abertura do Caixa
    If gobjLojaECF.iGerenteAutoriza = AUTORIZACAO_GERENTE Then

        'Chama a Tela de Senha
        Call Chama_TelaECF_Modal("OperadorLogin", objOperador, LOGIN_APENAS_GERENTE)

        'Sai de Função se a Tela de Login não Retornar ok
        If giRetornoTela <> vbOK Then gError 107959

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
    If lErro <> SUCESSO Then gError 107960

    lTamanho = 255
    sRetorno = String(lTamanho, 0)
        
    'Obtém a ultima transacao transferida
    Call GetPrivateProfileString(APLICACAO_DADOS, "UltimaTransacaoTransf", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
        
    'Retira os espaços no final da string
    sRetorno = StringZ(sRetorno)

    For Each objMovimentoCaixa In colMovimentosCaixa

        'se o numero da ultima transacao transferida ultrapassar o numero da transacao do movimento de caixa
        If objMovimentoCaixa.lSequencial <> 0 And StrParaLong(sRetorno) > objMovimentoCaixa.lSequencial Then gError 133850

        'Se Operador for Gerente
        objMovimentoCaixa.iGerente = objOperador.iCodigo

        lErro = Caixa_Grava_MovCx(objMovimentoCaixa, lSequencial)
        If lErro <> SUCESSO Then gError 105710

    Next

    lSequencial = lSequencial - 1
    
    'Fecha a Transação
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 107963

    Close #10

    Caixa_Grava_Movimento = SUCESSO

    Exit Function

Erro_Caixa_Grava_Movimento:

    Close #10

    Caixa_Grava_Movimento = gErr

    Select Case gErr
        
        Case 107957 To 107959
        
        Case 105710, 107960, 107963
            'Erro Tratado Dentro da Função Chamada
            
        Case 133850
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163084)

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
    If objMovimentoCaixa.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_TICKET Then

        Set objMovCx = New ClassMovimentoCaixa
            
        lErro = CF_ECF("MovimentoCaixa_Copia", objMovimentoCaixa, objMovCx)
        If lErro <> SUCESSO Then gError 105711
    
    Else
    
        Set objMovCx = objMovimentoCaixa
    
    End If

    'Guarda o Sequencial no objmovimentoCaixa
    objMovCx.lSequencial = lSequencial

    lSequencial = lSequencial + 1

    'Guarda no objMovimentoCaixa os Dados que Serão Usados para a Geração do Movimento de Caixa
    lErro = CF_ECF("Move_DadosGlobais_Memoria", objMovCx)
    If lErro <> SUCESSO Then gError 107961

    'Funçao que Gera o Arquivo preparando para a gravação
    Call CF_ECF("MovimentoTicket_Gera_Log", colRegistro, objMovCx)

    'Função que Vai Gravar as Informações no Arquivo de Caixa
    lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
    If lErro <> SUCESSO Then gError 107962
    
    Set objTela = Me
    
    'para não ficar 3 movimentos com o mesmo Código(Numero de Movto) na Coleção gcolMovto
    If objMovCx.iTipo = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_TICKET Then
        
'        'Faz a sangria
'        lErro = CF_ECF("Sangria_AFRAC", objMovCx.dValor, sMensagem, -1)
'        If lErro <> SUCESSO Then gError 109810
'
'    Else
'        'Faz a sangria
'        lErro = CF_ECF("Sangria_AFRAC", objMovCx.dValor, sMensagem, 0)
'        If lErro <> SUCESSO Then gError 109816
'    End If
    
    
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_TICKET, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Exclusão Sangria Ticket - Valor: " & Format(objMovCx.dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Exclusao Sangria - Ticket")
        If lErro <> SUCESSO Then gError 117675

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Exclusao Sangria - Ticket")
        If lErro <> SUCESSO Then gError 117676

        
    Else
        
        lErro = AFRAC_AbrirRelatorioGerencial(RELGER_SANGRIA_TICKET, objTela)

        lErro = AFRAC_ImprimirRelatorioGerencial("Sangria Ticket - Valor: " & Format(objMovCx.dValor, "standard"), objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Ticket")
        If lErro <> SUCESSO Then gError 117677

        lErro = AFRAC_FecharRelatorioGerencial(objTela)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Sangria - Ticket")
        If lErro <> SUCESSO Then gError 117678

    End If
    
    Caixa_Grava_MovCx = SUCESSO
    
    Exit Function

Erro_Caixa_Grava_MovCx:

    Caixa_Grava_MovCx = gErr

    Select Case gErr

        Case 105711, 107961, 107962, 109810, 109816, 117675 To 117678

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163085)

    End Select
    
    Exit Function

End Function

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
    If lErro <> SUCESSO Then gError 207980

    'Verifica se já foi executa a redução z para a data de hoje
    If gdtUltimaReducao = Date Then gError 111316
    
    'Verifica se o Codigo não foi Preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 107967

    lNumero = StrParaLong(Codigo.Text)

    'Verifica os Movimentos de Caixa para o Código em Questão
    lErro = CF_ECF("Caixa_MovimentoCaixa_Le_NumMovto", colMovimentosCaixa, lNumero)
    If lErro <> SUCESSO And lErro <> 107850 Then gError 107968
    
    'Senão encontrou o Movimento
    If lErro = 107850 Then gError 107969

    Set objMovCC = colMovimentosCaixa(1)

    If objMovCC.iTipo <> MOVIMENTOCAIXA_SANGRIA_TICKET Then gError 86291
    
    'Pergunta se deseja Realmente Excluir o Movimento
    vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_EXCLUSAO_MOVIMENTOCAIXA, Codigo.Text)

    If vbMsgRes = vbNo Then gError 107970
    iTipoMovimento = MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_TICKET

    'Prepara os Movimentos para a Exclusão
    lErro = MovimentoCaixa_Prepara_Exclusao(colMovimentosCaixa, iTipoMovimento)
    If lErro <> SUCESSO Then gError 107971

    'Função que Grava a Exclusão de Boletos
    lErro = Caixa_Grava_Movimento(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107972
    
    'Atualiza os Dados na Memória
    lErro = MovimentoTicket_Atualiza_Memoria(colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 107973

    'Função Que Limpa a Tela
    Call Limpa_Tela_MovimentoTickets

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr
        
        Case 86291
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_NAO_TICKET, gErr, StrParaDbl(Codigo.Text))
            
        Case 86292
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_JA_TRANSMITIDO, gErr)

        Case 107967
            Call Rotina_ErroECF(vbOKOnly, ERRO_CODIGO_NAO_PREENCHIDO1, gErr)

        Case 107969
            Call Rotina_ErroECF(vbOKOnly, ERRO_MOVIMENTO_INEXISTENTE, gErr, lNumero)

        Case 107968, 107970, 107971, 107972, 107973, 207980

        Case 111316
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163086)

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
    Call Limpa_Tela_MovimentoTickets

    'Traz os Valores em Caixa para a Tela
    lErro = Traz_ValoresEmCaixa_Tela()
    If lErro <> SUCESSO Then gError 107976

    'Função que Recalcula os Totáis
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError 107978

    Exit Sub

Erro_BotaoValoresEmCaixa_Click:

    Select Case gErr

        Case 107974, 107976 To 107978

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163087)

    End Select

End Sub

Function Traz_ValoresEmCaixa_Tela() As Long
'Função que Traz Todos os valores de teodos os movimentos relacionados com boletos para a Tela independente de código

Dim lErro As Long
Dim iLinha As Integer
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Traz_ValoresEmCaixa_Tela

    'Verifica na Coleção Global
    For Each objAdmMeioPagtoCondPagto In gcolTicket

        'Verifica se o Ticket não foi Especificado
        If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then

            LabelEmCaixaNaoDetalhadosValor.Caption = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
            ValorSangriaNaoDetalhado.Text = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))

        Else
        
            If objAdmMeioPagtoCondPagto.dSaldo <> 0 Then
        
                iLinha = iLinha + 1
        
                'Joga no Grid o nome da Adm
                GridTickets.TextMatrix(iLinha, iGrid_Administradora_Col) = objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto
                GridTickets.TextMatrix(iLinha, iGrid_Parcelamento_Col) = objAdmMeioPagtoCondPagto.sNomeParcelamento
                GridTickets.TextMatrix(iLinha, iGrid_ValorTotal_Col) = Format(objAdmMeioPagtoCondPagto.dSaldo, "Standard")
                GridTickets.TextMatrix(iLinha, iGrid_ValorSangria_Col) = Format(objAdmMeioPagtoCondPagto.dSaldo, "Standard")
                            
            End If
            
        End If
                            
    Next

    objGridTickets.iLinhasExistentes = iLinha

    Traz_ValoresEmCaixa_Tela = SUCESSO

    Exit Function

Erro_Traz_ValoresEmCaixa_Tela:

    Traz_ValoresEmCaixa_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163088)

    End Select

    Exit Function

End Function

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
            If lErro <> SUCESSO Then gError 107980

        End If

    End If

    Call Limpa_Tela_MovimentoTickets
    
    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 107980

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163089)

    End Select

    Exit Sub

End Sub

Function Carrega_MovimentoTicket_NaoEsp() As Long
'Função que Carrega os Controles Relacionados aos Cartões de Débito e Credito não especificados

Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Carrega_MovimentoTicket_NaoEsp

    For Each objAdmMeioPagtoCondPagto In gcolTicket

        If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO Then

            LabelEmCaixaNaoDetalhadosValor.Caption = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))
            ValorSangriaNaoDetalhado.Text = CStr(Format(objAdmMeioPagtoCondPagto.dSaldo, "standard"))

        End If

    Next

    Carrega_MovimentoTicket_NaoEsp = SUCESSO

    Exit Function

Erro_Carrega_MovimentoTicket_NaoEsp:

    Carrega_MovimentoTicket_NaoEsp = gErr

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163090)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
'Botão que Executa o Fechamento da Tela

    Unload Me


End Sub

Private Sub LabelCodigo_Click()

Dim objMovimentoCaixa As New ClassMovimentoCaixa
    
    'Chama tela de MovimentoBoletoLista
    Call Chama_TelaECF_Modal("MovimentoTicketLista", objMovimentoCaixa)
    
    If Not (objMovimentoCaixa Is Nothing) Then
        'Verifica se o Codvendedor está preenchido e joga na coleção
        If objMovimentoCaixa.lNumMovto <> 0 Then
            Codigo.Text = objMovimentoCaixa.lNumMovto
            Call CodMovimentoTicket_Validate(False)
            
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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163091)

    End Select

    Exit Sub

End Sub

Sub CodMovimentoTicket_Validate(Cancel As Boolean)
'Função que Verifica se o Código Passado como parâmetro Existe na Coleção Globa de MOVTOCAIXA

Dim lErro As Long
Dim objMovimentoCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_CodMovimentoTicket_Validate

    'Verifica se existe movimento com o código passado
    For Each objMovimentoCaixa In gcolMovimentosCaixa
    
        If objMovimentoCaixa.lNumMovto = StrParaLong(Codigo.Text) Then
            
            'Função que traz o MovimentoBoleto para a Tela
            lErro = Traz_MovimentoTickets_Tela(StrParaLong(Codigo.Text))
            If lErro <> SUCESSO Then gError 107982
            
            Exit For
        
        End If
    
    Next
    
    'Anula a Alteração
    iAlterado = 0
    
    Exit Sub
    
Erro_CodMovimentoTicket_Validate:
    
    Select Case gErr

        Case 107982

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163092)

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

    lErro = CF_ECF("Desmembra_MovimentosCaixa", colMovimentosCaixa, colImfCompl, TIPOREGISTROECF_MOVIMENTOCAIXA_TICKETS)
    If lErro <> SUCESSO Then gError 111065
    
   For Each objMovimentosCaixa In colMovimentosCaixa
        If objMovimentosCaixa.iTipo <> MOVIMENTOCAIXA_EXCLUSAO_SANGRIA_TICKET Then
            gcolMovimentosCaixa.Add colMovimentosCaixa
            gcolImfCompl.Add colImfCompl
        End If
    Next
    
    Exit Sub
    
Erro_DesmembraMovto_Click:
    
    Select Case gErr

        Case 111065

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163093)

    End Select

    Exit Sub


End Sub


