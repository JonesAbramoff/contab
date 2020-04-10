VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpDemApInvOcx 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8340
   KeyPreview      =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   8340
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDemApInvOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDemApInvOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDemApInvOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDemApInvOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4080
      Picture         =   "RelOpDemApInvOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDemApInvOcx.ctx":0A96
      Left            =   945
      List            =   "RelOpDemApInvOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2790
   End
   Begin VB.Frame Frame3 
      Caption         =   "Categoria de Produtos"
      Height          =   1785
      Left            =   120
      TabIndex        =   28
      Top             =   3765
      Width           =   5670
      Begin VB.ComboBox Categoria 
         Height          =   315
         Left            =   1650
         TabIndex        =   11
         Top             =   660
         Width           =   2745
      End
      Begin VB.ComboBox ValorInicial 
         Height          =   315
         Left            =   720
         TabIndex        =   12
         Top             =   1230
         Width           =   1950
      End
      Begin VB.CheckBox TodasCategorias 
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   285
         TabIndex        =   10
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox ValorFinal 
         Height          =   315
         Left            =   3420
         TabIndex        =   13
         Top             =   1215
         Width           =   2100
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   30
      End
      Begin VB.Label Label6 
         Caption         =   "Ate:"
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
         Height          =   255
         Left            =   2955
         TabIndex        =   31
         Top             =   1275
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "De:"
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
         Height          =   255
         Left            =   300
         TabIndex        =   30
         Top             =   1275
         Width           =   420
      End
      Begin VB.Label Label7 
         Caption         =   "Categoria:"
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
         Left            =   690
         TabIndex        =   29
         Top             =   705
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Almoxarifados"
      Height          =   1440
      Left            =   3120
      TabIndex        =   25
      Top             =   2160
      Width           =   2655
      Begin MSMask.MaskEdBox AlmoxarifadoInicial 
         Height          =   315
         Left            =   705
         TabIndex        =   8
         Top             =   315
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AlmoxarifadoFinal 
         Height          =   315
         Left            =   705
         TabIndex        =   9
         Top             =   915
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label labelAlmoxarifadoFinal 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   285
         TabIndex        =   27
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   330
         TabIndex        =   26
         Top             =   375
         Width           =   315
      End
   End
   Begin VB.ListBox Almoxarifados 
      Height          =   4350
      ItemData        =   "RelOpDemApInvOcx.ctx":0A9A
      Left            =   6000
      List            =   "RelOpDemApInvOcx.ctx":0A9C
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame FrameStatus 
      Caption         =   "Status"
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   5640
      Begin VB.OptionButton OpStatus 
         Caption         =   "Pendente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   210
         Width           =   1230
      End
      Begin VB.OptionButton OpStatus 
         Caption         =   "Processado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   4
         Top             =   210
         Width           =   1335
      End
      Begin VB.OptionButton OpStatus 
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
         Height          =   360
         Index           =   2
         Left            =   3720
         TabIndex        =   5
         Top             =   210
         Width           =   990
      End
   End
   Begin VB.ComboBox ComboTotaliza 
      Height          =   315
      ItemData        =   "RelOpDemApInvOcx.ctx":0A9E
      Left            =   3450
      List            =   "RelOpDemApInvOcx.ctx":0AAB
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   2280
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   1425
      Left            =   120
      TabIndex        =   21
      Top             =   2175
      Width           =   2865
      Begin MSMask.MaskEdBox TipoInicial 
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   345
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TipoFinal 
         Height          =   315
         Left            =   720
         TabIndex        =   7
         Top             =   945
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelTipoFinal 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   1005
         Width           =   360
      End
      Begin VB.Label LabelTipoInicial 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   405
         Width           =   315
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   1920
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   840
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInv 
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label9 
      Caption         =   "Ordena por:"
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
      Height          =   255
      Left            =   2340
      TabIndex        =   37
      Top             =   870
      Width           =   1080
   End
   Begin VB.Label Data 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
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
      Left            =   375
      TabIndex        =   36
      Top             =   900
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   315
      Width           =   615
   End
   Begin VB.Label LabelAlmoxarifado 
      AutoSize        =   -1  'True
      Caption         =   "Almoxarifados"
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
      Left            =   6000
      TabIndex        =   34
      Top             =   960
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "RelOpDemApInvOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giAlmoxInicial As Integer
Private WithEvents objEventoTipoInicial As AdmEvento
Attribute objEventoTipoInicial.VB_VarHelpID = -1
Private WithEvents objEventoTipoFinal As AdmEvento
Attribute objEventoTipoFinal.VB_VarHelpID = -1

Private Sub DataInv_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInv)

End Sub

Private Sub TipoInicial_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoInicial_Validate
    
    If Len(Trim(TipoInicial.Text)) <> 0 Then
    
        'Critica o valor
        lErro = Inteiro_Critica(Codigo_Extrai(TipoInicial.Text))
        If lErro <> SUCESSO Then Error 59555
    
        objTipoProduto.iTipo = CInt(Codigo_Extrai(TipoInicial.Text))
    
        'Lê o tipo
        lErro = CF("TipoDeProduto_Le",objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then Error 59556
        
        'Se não encontrar --> Erro
        If lErro = 22531 Then Error 59557
        
        TipoInicial.Text = objTipoProduto.iTipo & SEPARADOR & objTipoProduto.sDescricao
    
    End If
    
    Exit Sub

Erro_TipoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 59555, 59556

        Case 59557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", Err, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168071)

    End Select

    Exit Sub

End Sub

Private Sub TipoFinal_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoFinal_Validate
    
    If Len(Trim(TipoFinal.Text)) <> 0 Then
    
        'Critica o valor
        lErro = Inteiro_Critica(Codigo_Extrai(TipoFinal.Text))
        If lErro <> SUCESSO Then Error 59558
    
        objTipoProduto.iTipo = CInt(Codigo_Extrai(TipoFinal.Text))
    
        'Lê o tipo
        lErro = CF("TipoDeProduto_Le",objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then Error 59559
        
        'Se não encontrar --> Erro
        If lErro = 22531 Then Error 59560
        
        TipoFinal.Text = objTipoProduto.iTipo & SEPARADOR & objTipoProduto.sDescricao
    
    End If
    
    Exit Sub

Erro_TipoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 59558, 59559

        Case 59560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", Err, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168072)

    End Select

    Exit Sub

End Sub

Private Sub DataInv_Validate(Cancel As Boolean)

Dim sDataInv As String
Dim lErro As Long

On Error GoTo Erro_DataInv_Validate

    If Len(DataInv.ClipText) > 0 Then

        sDataInv = DataInv.Text
        
        lErro = Data_Critica(sDataInv)
        If lErro <> SUCESSO Then Error 59552

    End If

    Exit Sub

Erro_DataInv_Validate:

    Cancel = True


    Select Case Err

        Case 59552

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168073)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInv, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 59553

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 59553
            DataInv.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168074)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInv, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 59554

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 59554
            DataInv.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168075)

    End Select

    Exit Sub

End Sub

Private Sub AlmoxarifadoInicial_GotFocus()
'Mostra a lista de almoxarifado

Dim lErro As Long

On Error GoTo Erro_AlmoxarifadoInicial_GotFocus

    giAlmoxInicial = 1

    lErro = Mostra_Lista_Almoxarifado
    If lErro <> SUCESSO Then Error 54589

    Exit Sub

Erro_AlmoxarifadoInicial_GotFocus:

    Select Case Err

        Case 54589

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168076)

    End Select

    Exit Sub

End Sub

Private Sub AlmoxarifadoFinal_GotFocus()
'mostra a lista de almoxarifado

Dim lErro As Long

On Error GoTo Erro_AlmoxarifadoFinal_GotFocus

    giAlmoxInicial = 0

    lErro = Mostra_Lista_Almoxarifado
    If lErro <> SUCESSO Then Error 54590

    Exit Sub

Erro_AlmoxarifadoFinal_GotFocus:

    Select Case Err

        Case 54590

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168077)

    End Select

    Exit Sub

End Sub

Private Function Mostra_Lista_Almoxarifado() As Long
'esconde a treeview de produto e mostra a lista de almoxarifados

Dim lErro As Long

On Error GoTo Erro_Mostra_Lista_Almoxarifado

    'mostra a ListBox de almoxarifados
    Almoxarifados.Visible = True
    LabelAlmoxarifado.Visible = True

    Mostra_Lista_Almoxarifado = SUCESSO

    Exit Function

Erro_Mostra_Lista_Almoxarifado:

    Mostra_Lista_Almoxarifado = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168078)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load
        
    Set objEventoTipoInicial = New AdmEvento
    Set objEventoTipoFinal = New AdmEvento
    
    'carrega a ListBox Almoxarifados
    lErro = Carrega_Lista_Almoxarifado()
    If lErro <> SUCESSO Then Error 54591
    
    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas",colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 54592

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        Categoria.AddItem objCategoriaProduto.sCategoria

    Next
        
    Call Define_Padrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 54591, 54592

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168079)

    End Select

    Exit Sub

End Sub

Private Sub Define_Padrao()
'Preenche a tela com as opções padrão

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    giAlmoxInicial = 1
    
    TodasCategorias_Click
    TodasCategorias = 1
    
    OpStatus(1).Value = True
    
    ComboTotaliza.ListIndex = 0
        
    Call Mostra_Lista_Almoxarifado
    
    Exit Sub

Erro_Define_Padrao:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168080)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iTotaliza As Integer

On Error GoTo Erro_PreencherParametrosNaTela

 Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 54593

    'Pega status e exibe
    lErro = objRelOpcoes.ObterParametro("NSTATUS", sParam)
    If lErro <> SUCESSO Then Error 54594

    OpStatus(CInt(sParam)) = True
        
    'pega data e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINV", sParam)
    If lErro <> SUCESSO Then Error 54595

    Call DateParaMasked(DataInv, CDate(sParam))

   'pega parâmetro Almoxarifado Inicial e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NALMOXINIC", sParam)
    If lErro Then Error 54596
    
    AlmoxarifadoInicial.Text = sParam
    Call AlmoxarifadoInicial_Validate(bSGECancelDummy)
    
    'pega parâmetro Almoxarifado Final e exibe
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NALMOXFIM", sParam)
    If lErro Then Error 54597
    
    AlmoxarifadoFinal.Text = sParam
    Call AlmoxarifadoFinal_Validate(bSGECancelDummy)
   
    'pega parâmetro TodasCategorias e exibe
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro Then Error 54598

    TodasCategorias.Value = CInt(sParam)

    'pega parâmetro categoria de produto e exibe
    lErro = objRelOpcoes.ObterParametro("TCATPROD", sParam)
    If lErro Then Error 54599
    
    Categoria.Text = sParam
                
    'pega parâmetro tipo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TTIPOPRODINI", sParam)
    If lErro Then Error 54600
    
    TipoInicial.Text = sParam
    
    'pega parâmetro tipo final e exibe
    lErro = objRelOpcoes.ObterParametro("TTIPOPRODFIM", sParam)
    If lErro Then Error 54601
    
    TipoFinal.Text = sParam
    
    'pega parâmetro valor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODINI", sParam)
    If lErro Then Error 54602
    
    ValorInicial.Text = sParam
    
    'pega parâmetro Valor Final e exibe
    lErro = objRelOpcoes.ObterParametro("TITEMCATPRODFIM", sParam)
    If lErro Then Error 54603
    
    ValorFinal.Text = sParam
    
    'pega parâmetro de totalização
    sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("NTOTALIZA", sParam)
    If lErro Then Error 54604

    'seleciona ítem no ComboTotaliza
    iTotaliza = CInt(sParam)
    ComboTotaliza.ListIndex = iTotaliza
              
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 54593 To 54604
        
        Case 59555, 59556

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168081)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoTipoInicial = Nothing
    Set objEventoTipoFinal = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 54606
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 54607
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 54607
        
        Case 54606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168082)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    
    Categoria.Text = ""
    ValorInicial.Text = ""
    ValorFinal.Text = ""
    
    AlmoxarifadoInicial.Text = ""
    AlmoxarifadoFinal.Text = ""
    
    TipoInicial.Text = ""
    TipoFinal.Text = ""
      
    ComboOpcoes.SetFocus

End Sub

Private Function Formata_E_Critica_Parametros(sAlmox_I, sAlmox_F) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
   'critica Almoxarifado Inicial e Final
    If AlmoxarifadoInicial.Text <> "" Then
        sAlmox_I = CStr(Codigo_Extrai(AlmoxarifadoInicial.Text))
        
    Else
        sAlmox_I = ""
        
    End If
        
    If AlmoxarifadoFinal.Text <> "" Then
        sAlmox_F = CStr(Codigo_Extrai(AlmoxarifadoFinal.Text))
    
    Else
        sAlmox_F = ""
        
    End If
       
    If sAlmox_I <> "" And sAlmox_F <> "" Then
          
        If sAlmox_I <> "" And sAlmox_F <> "" Then
        
            If CInt(sAlmox_I) > CInt(sAlmox_F) Then Error 54609
        
        End If
        
    End If
    
    'tipo inicial não pode ser maior que o tipo final
    If Trim(TipoInicial.Text) <> "" And Trim(TipoFinal.Text) <> "" Then
    
         If TipoInicial.Text > TipoFinal.Text Then Error 54610
         
    End If
    
    'valor inicial não pode ser maior que o valor final
    If Trim(ValorInicial.Text) <> "" And Trim(ValorFinal.Text) <> "" Then
    
         If ValorInicial.Text > ValorFinal.Text Then Error 54611
         
    Else
        
        If Trim(ValorInicial.Text) = "" And Trim(ValorFinal.Text) = "" And TodasCategorias.Value = 0 Then Error 54612
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
         
        Case 54609
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INICIAL_MAIOR", Err)
            AlmoxarifadoInicial.SetFocus
                   
        Case 54611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", Err)
            ValorInicial.SetFocus
            
        Case 54610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_INICIAL_MAIOR", Err)
            TipoInicial.SetFocus
            
        Case 54612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO", Err)
            ValorInicial.SetFocus
              
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168083)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela
    Call Define_Padrao

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sAlmox_I As String
Dim sAlmox_F As String
Dim iIndice As Integer
Dim sTotaliza As String
Dim sStatus As String
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_PreencherRelOp
    
    If Len(DataInv.ClipText) = 0 Then Error 54615
        
    lErro = Formata_E_Critica_Parametros(sAlmox_I, sAlmox_F)
    If lErro <> SUCESSO Then Error 54616

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 54617
    
    'verifica opção selecionada
    For iIndice = 0 To 2
        If OpStatus(iIndice).Value = True Then sStatus = CStr(iIndice)
    Next

    lErro = objRelOpcoes.IncluirParametro("NSTATUS", sStatus)
    If lErro <> AD_BOOL_TRUE Then Error 54618
    
    If DataInv.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINV", DataInv.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINV", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 54619
                   
    lErro = objRelOpcoes.IncluirParametro("NALMOXINIC", sAlmox_I)
    If lErro <> AD_BOOL_TRUE Then Error 54620
        
    lErro = objRelOpcoes.IncluirParametro("TALMOXINICIAL", AlmoxarifadoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54621
    
    lErro = objRelOpcoes.IncluirParametro("NALMOXFIM", sAlmox_F)
    If lErro <> AD_BOOL_TRUE Then Error 54622
        
    lErro = objRelOpcoes.IncluirParametro("TALMOXFINAL", AlmoxarifadoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54623
       
    objEstoqueMes.iFilialEmpresa = giFilialEmpresa
   
    'Ler o mês e o ano que está aberto
    lErro = CF("EstoqueMes_Le_Aberto",objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 40673 Then Error 54624

    If lErro = 40673 Then Error 54625
 
    lErro = objRelOpcoes.IncluirParametro("NANO", objEstoqueMes.iAno)
    If lErro <> AD_BOOL_TRUE Then Error 54626
 
    lErro = objRelOpcoes.IncluirParametro("NMES", objEstoqueMes.iMes)
    If lErro <> AD_BOOL_TRUE Then Error 54627
    
    'le o ultimo ano/mes apurado
    lErro = CF("EstoqueMes_Le_Apurado",objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 46225 Then Error 54628
    
    If lErro = 46225 Then
        objEstoqueMes.iAno = 0
        objEstoqueMes.iMes = 0
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NANOAPURADO", objEstoqueMes.iAno)
    If lErro <> AD_BOOL_TRUE Then Error 54629
 
    lErro = objRelOpcoes.IncluirParametro("NMESAPURADO", objEstoqueMes.iMes)
    If lErro <> AD_BOOL_TRUE Then Error 54630
    
    lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(TodasCategorias.Value))
    If lErro <> AD_BOOL_TRUE Then Error 54631
    
    lErro = objRelOpcoes.IncluirParametro("TCATPROD", Categoria.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54632
    
    lErro = objRelOpcoes.IncluirParametro("TTIPOPRODINI", TipoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54633
    
    lErro = objRelOpcoes.IncluirParametro("TTIPOPRODFIM", TipoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54634
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODINI", ValorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54635
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATPRODFIM", ValorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54636
        
    sTotaliza = CStr(ComboTotaliza.ListIndex)
    
    lErro = objRelOpcoes.IncluirParametro("NTOTALIZA", sTotaliza)
    If lErro <> AD_BOOL_TRUE Then Error 54637

    If TodasCategorias.Value = 0 Then
        If ComboTotaliza.ListIndex = 0 Then
            If sStatus = "0" Then gobjRelatorio.sNomeTsk = "daipeaca"
            If sStatus = "1" Then gobjRelatorio.sNomeTsk = "daipraca"
            If sStatus = "2" Then gobjRelatorio.sNomeTsk = "daitoaca"
        ElseIf ComboTotaliza.ListIndex = 1 Then
            If sStatus = "0" Then gobjRelatorio.sNomeTsk = "daipetca"
            If sStatus = "1" Then gobjRelatorio.sNomeTsk = "daiprtca"
            If sStatus = "2" Then gobjRelatorio.sNomeTsk = "daitotca"
        ElseIf ComboTotaliza.ListIndex = 2 Then
            If sStatus = "0" Then gobjRelatorio.sNomeTsk = "daipeatc"
            If sStatus = "1" Then gobjRelatorio.sNomeTsk = "daipratc"
            If sStatus = "2" Then gobjRelatorio.sNomeTsk = "daitoatc"
        End If
    ElseIf TodasCategorias.Value = 1 Then
        If ComboTotaliza.ListIndex = 0 Then
            If sStatus = "0" Then gobjRelatorio.sNomeTsk = "daipenda"
            If sStatus = "1" Then gobjRelatorio.sNomeTsk = "daiproca"
            If sStatus = "2" Then gobjRelatorio.sNomeTsk = "daitodoa"
        ElseIf ComboTotaliza.ListIndex = 1 Then
            If sStatus = "0" Then gobjRelatorio.sNomeTsk = "daipendt"
            If sStatus = "1" Then gobjRelatorio.sNomeTsk = "daiproct"
            If sStatus = "2" Then gobjRelatorio.sNomeTsk = "daitodot"
        ElseIf ComboTotaliza.ListIndex = 2 Then
            If sStatus = "0" Then gobjRelatorio.sNomeTsk = "daipenat"
            If sStatus = "1" Then gobjRelatorio.sNomeTsk = "daiproat"
            If sStatus = "2" Then gobjRelatorio.sNomeTsk = "daitodat"
        End If
    End If
        
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sAlmox_I, sAlmox_F)
    If lErro <> SUCESSO Then Error 54638

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 54616 To 54638
                               
        Case 54615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168084)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 54639

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 54640

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 54639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 54640

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168085)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 54641

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 54641

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168086)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 54642

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 54643

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 54644

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 54642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 54643, 54644

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168087)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sAlmox_I As String, sAlmox_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

     sExpressao = ""

    If sAlmox_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "Almoxarifado >= " & Forprint_ConvInt(CInt(sAlmox_I))

    End If

    If sAlmox_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Almoxarifado <= " & Forprint_ConvInt(CInt(sAlmox_F))

    End If
    
     If TodasCategorias.Value = 0 Then
           
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CategoriaProduto = " & Forprint_ConvTexto(Categoria.Text)
            
        If ValorInicial.Text <> "" Then

            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "ItemCategoriaProduto  >= " & Forprint_ConvTexto(ValorInicial.Text)

        End If
        
        If ValorFinal.Text <> "" Then

            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "ItemCategoriaProduto <= " & Forprint_ConvTexto(ValorFinal.Text)

        End If
        
    End If
     
    If TipoInicial.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoProduto  >= " & Forprint_ConvInt(CInt(Codigo_Extrai(TipoInicial.Text)))

    End If
        
    If TipoFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoProduto <= " & Forprint_ConvInt(CInt(Codigo_Extrai(TipoFinal.Text)))

    End If
        
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168088)

    End Select

    Exit Function

End Function

Private Function Carrega_Lista_Almoxarifado() As Long
'Carrega a ListBox Almoxarifados

Dim lErro As Long
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Carrega_Lista_Almoxarifado
    
    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Almoxarifados_Le_FilialEmpresa",giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then Error 54645

    'Preenche a ListBox AlmoxarifadoList com os objetos da coleção
    For Each objAlmoxarifado In colAlmoxarifados
        Almoxarifados.AddItem objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
        Almoxarifados.ItemData(Almoxarifados.NewIndex) = objAlmoxarifado.iCodigo
    Next

    Carrega_Lista_Almoxarifado = SUCESSO

    Exit Function

Erro_Carrega_Lista_Almoxarifado:

    Carrega_Lista_Almoxarifado = Err

    Select Case Err

        Case 54645

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168089)

    End Select

    Exit Function

End Function

Private Sub Almoxarifados_DblClick()
'Preenche Almoxarifado Final ou Inicial com o almoxarifado selecionado

Dim lErro As Long
Dim sListBoxItem As String
Dim objCodigoDescricao As New AdmCodigoNome
Dim objAlmoxarifado As ClassAlmoxarifado
Dim objAlmoxSelecionado As ClassAlmoxarifado

On Error GoTo Erro_Almoxarifados_DblClick

    'Guarda a string selecionada na ListBox Almoxarifados
    sListBoxItem = Almoxarifados.List(Almoxarifados.ListIndex)
 
    If giAlmoxInicial = 1 Then
    
        AlmoxarifadoInicial.Text = sListBoxItem
        
    Else
        AlmoxarifadoFinal.Text = sListBoxItem

    End If

    Exit Sub

Erro_Almoxarifados_DblClick:

    Select Case Err

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168090)

    End Select

    Exit Sub

End Sub

Private Sub AlmoxarifadoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoInicial_Validate

    If Len(Trim(AlmoxarifadoInicial.Text)) > 0 Then
   
        'Tenta ler o Almoxarifado (NomeReduzido ou Código)
        lErro = TP_Almoxarifado_Le_ComCodigo(AlmoxarifadoInicial, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 54646

    End If
    
    Exit Sub

Erro_AlmoxarifadoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 54646

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168091)

    End Select

End Sub

Private Sub AlmoxarifadoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoFinal_Validate

    If Len(Trim(AlmoxarifadoFinal.Text)) > 0 Then

        'Tenta ler o Almoxarifado (NomeReduzido ou Código)
        lErro = TP_Almoxarifado_Le_ComCodigo(AlmoxarifadoFinal, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 54647

    End If
 
    Exit Sub

Erro_AlmoxarifadoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 54647

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168092)

    End Select

End Sub

Private Sub Categoria_GotFocus()

    If TodasCategorias.Value = 1 Then TodasCategorias.Value = 0
        
End Sub

Private Sub Categoria_Validate(Cancel As Boolean)

    Categoria_Click
 
End Sub

Private Sub Categoria_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colCategoria As New Collection

On Error GoTo Erro_Categoria_Click

    If Len(Trim(Categoria.Text)) > 0 Then

        ValorInicial.Clear
        ValorFinal.Clear
        
        'Preenche o objeto com a Categoria
         objCategoriaProduto.sCategoria = Categoria.Text

        'Lê Categoria De Produto no BD
        lErro = CF("CategoriaProduto_Le",objCategoriaProduto)
        If lErro <> SUCESSO And lErro <> 22540 Then Error 54647

        If lErro <> SUCESSO Then Error 54648 'Categoria não está cadastrada

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens",objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then Error 54649

        'Preenche Valor Inicial e final
        For Each objCategoriaProdutoItem In colCategoria

            ValorInicial.AddItem (objCategoriaProdutoItem.sItem)
            ValorFinal.AddItem (objCategoriaProdutoItem.sItem)

        Next

    Else
    
        ValorInicial.Text = ""
        ValorFinal.Text = ""
        ValorInicial.Clear
        ValorFinal.Clear

    End If

    Exit Sub

Erro_Categoria_Click:

    Select Case Err

        Case 54647
            Categoria.SetFocus
            
        Case 54648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", Err)
            Categoria.SetFocus
            
        Case 54649

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168093)

    End Select

    Exit Sub

End Sub

Private Sub ValorInicial_GotFocus()

    If TodasCategorias.Value = 1 Then TodasCategorias.Value = 0
    
End Sub

Private Sub ValorFinal_GotFocus()

    If TodasCategorias.Value = 1 Then TodasCategorias.Value = 0
    
End Sub

Private Sub ValorInicial_Validate(Cancel As Boolean)

    ValorInicial_Click

End Sub

Private Sub ValorFinal_Validate(Cancel As Boolean)

    ValorFinal_Click

End Sub

Private Sub TodasCategorias_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_TodasCategorias_Click

    'Limpa campos
    Categoria.Text = ""
    ValorInicial.Text = ""
    ValorFinal.Text = ""
    
    Exit Sub

Erro_TodasCategorias_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168094)

    End Select

    Exit Sub

End Sub

Private Sub ValorInicial_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_ValorInicial_Click

    If Len(Trim(ValorInicial.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ValorInicial)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = Categoria.Text
            objCategoriaProdutoItem.sItem = ValorInicial.Text

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item",objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then Error 54650

            If lErro <> SUCESSO Then Error 54651 'Item da Categoria não está cadastrado

        End If

    End If

    Exit Sub

Erro_ValorInicial_Click:

    Select Case Err

        Case 54650
            ValorInicial.SetFocus

        Case 54651
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168095)

    End Select

    Exit Sub

End Sub

Private Sub ValorFinal_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_ValorFinal_Click

    If Len(Trim(ValorFinal.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ValorFinal)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = Categoria.Text
            objCategoriaProdutoItem.sItem = ValorFinal.Text

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item",objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then Error 54652

            If lErro <> SUCESSO Then Error 54653 'Item da Categoria não está cadastrado

        End If

    End If

    Exit Sub

Erro_ValorFinal_Click:

    Select Case Err

        Case 54652
            ValorFinal.SetFocus

        Case 54653
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ValorFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168096)

    End Select

    Exit Sub

End Sub

Private Sub LabelTipoInicial_Click()

Dim lErro As Long
Dim objTipoProduto As ClassTipoDeProduto
Dim colSelecao As Collection

On Error GoTo Erro_LabelTipoInicial_Click

    If Len(Trim(TipoInicial.Text)) <> 0 Then

        Set objTipoProduto = New ClassTipoDeProduto
        objTipoProduto.iTipo = TipoInicial.Text

    End If

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoInicial)

    Exit Sub

Erro_LabelTipoInicial_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168097)

    End Select

    Exit Sub

End Sub

Private Sub LabelTipoFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objTipoProduto As ClassTipoDeProduto

On Error GoTo Erro_LabelTipoFinal_Click

    If Len(Trim(TipoFinal.Text)) <> 0 Then

        Set objTipoProduto = New ClassTipoDeProduto
        objTipoProduto.iTipo = TipoFinal.Text

    End If

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoFinal)

    Exit Sub

Erro_LabelTipoFinal_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168098)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoInicial_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoTipoInicial_evSelecao

    Set objTipoProduto = obj1

    TipoInicial.Text = objTipoProduto.iTipo
    
    Me.Show
    
    Exit Sub

Erro_objEventoTipoInicial_evSelecao:

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168099)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoFinal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoTipoFinal_evSelecao

    Set objTipoProduto = obj1

    TipoFinal.Text = objTipoProduto.iTipo
    
    Me.Show
    
    Exit Sub

Erro_objEventoTipoFinal_evSelecao:

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168100)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_DEM_APURACAO_INVENTARIO
    Set Form_Load_Ocx = Me
    Caption = "Demonstrativo de Apuração de Inventário"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpDemApInv"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is TipoInicial Then
            Call LabelTipoInicial_Click
        ElseIf Me.ActiveControl Is TipoFinal Then
            Call LabelTipoFinal_Click
        End If
    
    End If

End Sub


Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub labelAlmoxarifadoFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(labelAlmoxarifadoFinal, Source, X, Y)
End Sub

Private Sub labelAlmoxarifadoFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(labelAlmoxarifadoFinal, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelTipoFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipoFinal, Source, X, Y)
End Sub

Private Sub LabelTipoFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipoFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelTipoInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipoInicial, Source, X, Y)
End Sub

Private Sub LabelTipoInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipoInicial, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Data_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Data, Source, X, Y)
End Sub

Private Sub Data_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Data, Button, Shift, X, Y)
End Sub

Private Sub LabelAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAlmoxarifado, Source, X, Y)
End Sub

Private Sub LabelAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAlmoxarifado, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

