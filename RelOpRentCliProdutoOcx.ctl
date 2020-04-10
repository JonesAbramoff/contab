VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpRentCliProduto 
   ClientHeight    =   6435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
   ScaleHeight     =   6435
   ScaleWidth      =   7140
   Begin MSMask.MaskEdBox GastosPropaganda 
      Height          =   315
      Left            =   2295
      TabIndex        =   37
      Top             =   780
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      _Version        =   393216
      Format          =   "#0.#0\%"
      PromptChar      =   "_"
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   135
      TabIndex        =   31
      Top             =   1680
      Width           =   5565
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1590
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   315
         Left            =   600
         TabIndex        =   3
         Top             =   255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   4245
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   330
         Left            =   3270
         TabIndex        =   4
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dFimPrev 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2850
         TabIndex        =   35
         Top             =   300
         Width           =   450
      End
      Begin VB.Label dIniPrev 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   195
         TabIndex        =   34
         Top             =   300
         Width           =   390
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   900
      Left            =   135
      TabIndex        =   28
      Top             =   5385
      Width           =   5565
      Begin MSMask.MaskEdBox ClienteDe 
         Height          =   300
         Left            =   630
         TabIndex        =   5
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteAte 
         Height          =   300
         Left            =   3255
         TabIndex        =   6
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteAte 
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
         Left            =   2835
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   30
         Top             =   420
         Width           =   360
      End
      Begin VB.Label LabelClienteDe 
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   29
         Top             =   405
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Região"
      Height          =   1245
      Left            =   135
      TabIndex        =   23
      Top             =   2505
      Width           =   5565
      Begin MSMask.MaskEdBox RegiaoDe 
         Height          =   315
         Left            =   585
         TabIndex        =   11
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RegiaoAte 
         Height          =   315
         Left            =   585
         TabIndex        =   1
         Top             =   765
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label RegiaoAteDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2025
         TabIndex        =   27
         Top             =   765
         Width           =   3120
      End
      Begin VB.Label RegiaoDeDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2025
         TabIndex        =   26
         Top             =   315
         Width           =   3120
      End
      Begin VB.Label LabelRegiaoDe 
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
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   360
         Width           =   360
      End
      Begin VB.Label LabelRegiaoAte 
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
         Height          =   255
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   810
         Width           =   435
      End
   End
   Begin VB.Frame FrameCategoriaCliente 
      Caption         =   "Categoria"
      Height          =   1470
      Left            =   135
      TabIndex        =   18
      Top             =   3825
      Width           =   5565
      Begin VB.ComboBox CategoriaCliente 
         Height          =   315
         Left            =   1650
         TabIndex        =   8
         Top             =   540
         Width           =   2745
      End
      Begin VB.ComboBox CategoriaClienteDe 
         Height          =   315
         Left            =   555
         TabIndex        =   9
         Top             =   1020
         Width           =   1740
      End
      Begin VB.CheckBox CategoriaClienteTodas 
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
         Left            =   300
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CategoriaClienteAte 
         Height          =   315
         Left            =   3225
         TabIndex        =   10
         Top             =   1005
         Width           =   1740
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   15
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   30
      End
      Begin VB.Label LabelCategoriaClienteAte 
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
         Left            =   2805
         TabIndex        =   21
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label LabelCategoriaClienteDe 
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
         Left            =   180
         TabIndex        =   20
         Top             =   1065
         Width           =   315
      End
      Begin VB.Label LabelCategoriaCliente 
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
         Left            =   735
         TabIndex        =   19
         Top             =   585
         Width           =   855
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRentCliProdutoOcx.ctx":0000
      Left            =   1920
      List            =   "RelOpRentCliProdutoOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   270
      Width           =   2730
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
      Left            =   4920
      Picture         =   "RelOpRentCliProdutoOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   870
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4815
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRentCliProdutoOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRentCliProdutoOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRentCliProdutoOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRentCliProdutoOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox GastosFinanceiros 
      Height          =   315
      Left            =   2295
      TabIndex        =   39
      Top             =   1260
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      _Version        =   393216
      Format          =   "#0.#0\%"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      Caption         =   "Gastos Financeiros:"
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
      Left            =   525
      TabIndex        =   38
      Top             =   1305
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Gastos de Propaganda:"
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
      Left            =   210
      TabIndex        =   36
      Top             =   810
      Width           =   2040
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
      Left            =   1215
      TabIndex        =   17
      Top             =   315
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRentCliProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Inserido por Wagner
'####################################
Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long
'###################################

'Property Variables:
Dim m_Caption As String
Event Unload()

'Browses
Private WithEvents objEventoRegiaoVendaDe As AdmEvento
Attribute objEventoRegiaoVendaDe.VB_VarHelpID = -1
Private WithEvents objEventoRegiaoVendaAte As AdmEvento
Attribute objEventoRegiaoVendaAte.VB_VarHelpID = -1
Private WithEvents objEventoClienteDe As AdmEvento
Attribute objEventoClienteDe.VB_VarHelpID = -1
Private WithEvents objEventoClienteAte As AdmEvento
Attribute objEventoClienteAte.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long
Dim colCategoriaCliente As New Collection
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Form_Load
    
    Set objEventoClienteDe = New AdmEvento
    Set objEventoClienteAte = New AdmEvento
    Set objEventoRegiaoVendaDe = New AdmEvento
    Set objEventoRegiaoVendaAte = New AdmEvento
    
    'Le as categorias de Cliente
    lErro = CF("CategoriaCliente_Le_Todos", colCategoriaCliente)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 125641

    'Preenche CategoriaCliente
    For Each objCategoriaCliente In colCategoriaCliente

        CategoriaCliente.AddItem objCategoriaCliente.sCategoria

    Next
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 125642

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 125641, 125642

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179562)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 125643

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche a Combo Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 125644

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 125643
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 125644

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179563)

    End Select

    Exit Function

End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** EVENTO GOTFOCUS DOS CONTROLES - INÍCIO ***

Private Sub CategoriaClienteDe_GotFocus()

    If CategoriaClienteTodas.Value = 1 Then CategoriaClienteTodas.Value = 0
    
End Sub

Private Sub CategoriaClienteAte_GotFocus()

    If CategoriaClienteTodas.Value = 1 Then CategoriaClienteTodas.Value = 0
    
End Sub

Private Sub DataAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAte)

End Sub

Private Sub DataDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataDe)

End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO***
Private Sub ClienteDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteDe_Validate

    'se está Preenchido
    If Len(Trim(ClienteDe.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteDe, objCliente, 0)
        If lErro <> SUCESSO Then gError 125645

    End If

    Exit Sub

Erro_ClienteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 125645

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179564)

    End Select

End Sub

Private Sub ClienteAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteAte_Validate

    'Se está Preenchido
    If Len(Trim(ClienteAte.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteAte, objCliente, 0)
        If lErro <> SUCESSO Then gError 125646

    End If

    Exit Sub

Erro_ClienteAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125646

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179565)

    End Select

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub CategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_CategoriaCliente_Validate

    If Len(CategoriaCliente) <> 0 And CategoriaCliente.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Seleciona(CategoriaCliente, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 125647
        
        If lErro <> SUCESSO Then gError 125648
        
    Else
        If Len(CategoriaCliente) = 0 Then
        
            'se a Categoria não estiver preenchida ----> limpa e disabilita os Valores Inicial e Final
            CategoriaClienteDe.ListIndex = -1
            CategoriaClienteDe.Enabled = False
            CategoriaClienteAte.ListIndex = -1
            CategoriaClienteAte.Enabled = False
        
        End If
    
    End If
    
    Exit Sub

Erro_CategoriaCliente_Validate:

    Cancel = True

    Select Case gErr

        Case 125647
        
        Case 125648
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, CategoriaCliente.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179566)

    End Select

    Exit Sub
  
End Sub

Private Sub CategoriaClienteDe_Validate(Cancel As Boolean)

    Call CategoriaClienteDe_Click

End Sub

Private Sub CategoriaClienteAte_Validate(Cancel As Boolean)

    Call CategoriaClienteAte_Click

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Se a DataDe estiver preenchida
    If Len(DataDe.ClipText) > 0 Then

        sDataInic = DataDe.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 125649

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 125649

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179567)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'se a DataAte estiver preenchida
    If Len(DataAte.ClipText) > 0 Then

        sDataFim = DataAte.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 125650

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125650

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179568)

    End Select

    Exit Sub

End Sub

Private Sub RegiaoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoAte_Validate
    
    lErro = RegiaoVenda_Perde_Foco(RegiaoAte, RegiaoAteDesc)
    If lErro <> SUCESSO And lErro <> 87199 Then gError 125651
       
    If lErro = 87199 Then gError 125652
        
    Exit Sub

Erro_RegiaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125651
        
        Case 125652
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179569)

    End Select

    Exit Sub

End Sub

Private Sub RegiaoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sRegiaoInicial As String
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoDe_Validate

    lErro = RegiaoVenda_Perde_Foco(RegiaoDe, RegiaoDeDesc)
    If lErro <> SUCESSO And lErro <> 87199 Then gError 125653
       
    If lErro = 87199 Then gError 125654
    
    Exit Sub

Erro_RegiaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 125653
        
        Case 125654
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179570)

    End Select

    Exit Sub

End Sub
'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125655

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 125655
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179571)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125656

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 125656
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179572)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125657

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 125657
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179573)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125658

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 125658
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179574)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
    
    If Len(Trim(ClienteAte.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteAte.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteAte)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
    
    If Len(Trim(ClienteDe.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteDe.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteDe)

End Sub

Private Sub LabelRegiaoAte_Click()
    
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    'Se o tipo está preenchido
    If Len(Trim(RegiaoAte.Text)) > 0 Then
        
        'Preenche com o tipo da tela
        objRegiaoVenda.iCodigo = CInt(RegiaoAte.Text)
    
    End If
    
    'Chama Tela RegiãoVendaLista
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiaoVenda, objEventoRegiaoVendaAte)
    
End Sub

Private Sub LabelRegiaoDe_Click()

Dim objRegiaoVenda As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    'Se o tipo está preenchido
    If Len(Trim(RegiaoDe.Text)) > 0 Then
        
        'Preenche com o tipo da tela
        objRegiaoVenda.iCodigo = CInt(RegiaoDe.Text)
        
    End If
    
    'Chama Tela RegiãoVendaLista
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiaoVenda, objEventoRegiaoVendaDe)

End Sub
Private Sub CategoriaCliente_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim objCategoriaClienteItem As New ClassCategoriaClienteItem
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaCliente_Click

    If CategoriaCliente.ListIndex <> -1 Then

        'Preenche o objeto com a Categoria
         objCategoriaCliente.sCategoria = CategoriaCliente.Text

        'Lê os dados de itens de categorias de Cliente
        lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colCategoria)
        If lErro <> SUCESSO Then gError 125659
        
        'Habilita e Limpa Cliente De
        CategoriaClienteDe.Enabled = True
        CategoriaClienteDe.Clear
        
        'Habilita e Limpa Cliente Até
        CategoriaClienteAte.Enabled = True
        CategoriaClienteAte.Clear

        'Preenche a Combo Cliente De e Até
        For Each objCategoriaClienteItem In colCategoria
            
            CategoriaClienteDe.AddItem (objCategoriaClienteItem.sItem & SEPARADOR & objCategoriaClienteItem.sDescricao)
            
            CategoriaClienteAte.AddItem (objCategoriaClienteItem.sItem & SEPARADOR & objCategoriaClienteItem.sDescricao)

        Next
            
        'desmarca todasCategorias
        CategoriaClienteTodas.Value = 0
    
    Else
    
        'se a Categoria não estiver preenchida ----> limpa e desabilita os Cliente De e Até
        CategoriaClienteDe.Clear
        CategoriaClienteDe.ListIndex = -1
        CategoriaClienteDe.Enabled = False
        CategoriaClienteAte.Clear
        CategoriaClienteAte.ListIndex = -1
        CategoriaClienteAte.Enabled = False
    
    End If

    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case gErr

        Case 125659

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179575)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteTodas_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_CategoriaClienteTodas_Click
    
    'Se a check CategoriaClienteTodas estiver setada então desabilita as Combo Cliente, ClienteDe e ClienteAté
    If CategoriaClienteTodas.Value = 1 Then
    
        CategoriaCliente.ListIndex = -1
        CategoriaCliente.Enabled = False
        
        CategoriaClienteDe.ListIndex = -1
        CategoriaClienteDe.Enabled = False
        
        CategoriaClienteAte.ListIndex = -1
        CategoriaClienteAte.Enabled = False
    
    Else
        'se não habilita a ComboCliente
        CategoriaCliente.ListIndex = -1
        CategoriaCliente.Enabled = True
        
    End If
    
    Exit Sub

Erro_CategoriaClienteTodas_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179576)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteDe_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim objCategoriaClienteItem As New ClassCategoriaClienteItem
Dim colItens As New Collection

On Error GoTo Erro_CategoriaClienteDe_Click

    If Len(Trim(CategoriaClienteDe.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(CategoriaClienteDe)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaClienteItem.sCategoria = CategoriaCliente.Text
            objCategoriaClienteItem.sItem = CategoriaClienteDe.Text

            'Lê Categoria De Cliente no BD
            lErro = CF("CategoriaClienteItem_Le", objCategoriaClienteItem)
            If lErro <> SUCESSO And lErro <> 22603 Then gError 125660

            If lErro <> SUCESSO Then gError 125661 'Item da Categoria não está cadastrado

        End If

    End If

    Exit Sub

Erro_CategoriaClienteDe_Click:

    Select Case gErr
    
        Case 125660
            CategoriaClienteDe.SetFocus

        Case 125661
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", gErr, objCategoriaClienteItem.sItem, objCategoriaClienteItem.sCategoria)
            CategoriaClienteDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179577)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteAte_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim objCategoriaClienteItem As New ClassCategoriaClienteItem
Dim colItens As New Collection

On Error GoTo Erro_CategoriaClienteAte_Click

    If Len(Trim(CategoriaClienteAte.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(CategoriaClienteAte)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaClienteItem.sCategoria = CategoriaCliente.Text
            objCategoriaClienteItem.sItem = CategoriaClienteAte.Text

            'Lê Categoria De Cliente no BD
            lErro = CF("CategoriaClienteItem_Le", objCategoriaClienteItem)
            If lErro <> SUCESSO And lErro <> 22603 Then gError 125662

            If lErro <> SUCESSO Then gError 125663 'Item da Categoria não está cadastrado

        End If

    End If

    Exit Sub

Erro_CategoriaClienteAte_Click:

    Select Case gErr
    
        Case 125662
            CategoriaClienteAte.SetFocus

        Case 125663
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", gErr, objCategoriaClienteItem.sItem, objCategoriaClienteItem.sCategoria)
            CategoriaClienteAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179578)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 125664

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 125664

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179579)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    'Limpa a tela
    lErro = LimpaRelatorioRentCliProduto()
    If lErro <> SUCESSO Then gError 125665
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 125665
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179580)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 125666

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        'Exclui o elemento do banco de dados
        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 125667

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa a tela
        lErro = LimpaRelatorioRentCliProduto()
        If lErro <> SUCESSO Then gError 125668
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 125666
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 125667, 125668

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179581)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 125669

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 125670

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Grava no banco de dados
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 125671
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 125672
    
    'Limpa a tela
    lErro = LimpaRelatorioRentCliProduto()
    If lErro <> SUCESSO Then gError 125673
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 125669
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 125670 To 125673
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179582)

    End Select

    Exit Sub

End Sub
'*** EVENTO CLICK DOS CONTROLES - FIM ***

'*** FUNÇÕES DE APOIO A TELA - INÍCIO
Private Function LimpaRelatorioRentCliProduto()
'Limpa a tela RelOpRentCliProduto

Dim lErro As Long

On Error GoTo Erro_LimpaRelatorioRentCliProduto

    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 125674
    
    ComboOpcoes.Text = ""
   
    CategoriaCliente.ListIndex = -1
    CategoriaClienteDe.Clear
    CategoriaClienteAte.Clear
   
    RegiaoDeDesc.Caption = ""
    RegiaoAteDesc.Caption = ""
   
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 125675
    
    LimpaRelatorioRentCliProduto = SUCESSO
    
    Exit Function
    
Erro_LimpaRelatorioRentCliProduto:

    LimpaRelatorioRentCliProduto = gErr
    
    Select Case gErr
    
        Case 125674, 125675
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179583)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long
'Preenche as datas e carrega as combos da tela

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
        
    DataDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataAte.Text = Format(gdtDataAtual, "dd/mm/yy")
        
    CategoriaClienteTodas.Value = 1
        
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179584)

    End Select

    Exit Function

End Function

Public Function RegiaoVenda_Perde_Foco(Regiao As Object, Desc As Object) As Long
'recebe MaskEdBox da Região de Venda e o label da descrição

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoVenda_Perde_Foco

        
    If Len(Trim(Regiao.Text)) > 0 Then
        
        objRegiaoVenda.iCodigo = StrParaInt(Regiao.Text)
    
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 125676
    
        If lErro = 16137 Then gError 125677

        Desc.Caption = objRegiaoVenda.sDescricao

    Else

        Desc.Caption = ""

    End If

    RegiaoVenda_Perde_Foco = SUCESSO

    Exit Function

Erro_RegiaoVenda_Perde_Foco:

    RegiaoVenda_Perde_Foco = gErr

    Select Case gErr

        Case 125676
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGIOESVENDAS", gErr, objRegiaoVenda.iCodigo)

        Case 125677
            'Erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 179585)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExec As Boolean = False) As Long
'preenche o objRelOp com os dados fornecidos pelo usuário

Dim lErro As Long, lNumIntRel As Long
Dim sCliente_De As String, dInic As Date
Dim sCliente_Ate As String, dCustoFretesGeral As Double, dCustoComissGeral As Double

On Error GoTo Erro_PreencherRelOp
   
    'Critica os valores preenchidos pelo usuário
    lErro = Formata_E_Critica_Parametros(sCliente_De, sCliente_Ate)
    If lErro <> SUCESSO Then gError 125678
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 125679
        
    lErro = objRelOpcoes.IncluirParametro("NPROPAGANDA", CStr(StrParaDbl(GastosPropaganda.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 125680
        
    lErro = objRelOpcoes.IncluirParametro("NPERC_GASTO_FINAN", CStr(StrParaDbl(GastosFinanceiros.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 125680
        
    'Inclui o cliente inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_De)
    If lErro <> AD_BOOL_TRUE Then gError 125680
    
    'Inclui o cliente final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 125681
    
    'Inclui a categoria
    lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(CategoriaClienteTodas.Value))
    If lErro <> AD_BOOL_TRUE Then gError 125682
    
    lErro = objRelOpcoes.IncluirParametro("TCATCLI", CategoriaCliente.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125683
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATCLIINI", CategoriaClienteDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125684
    
    lErro = objRelOpcoes.IncluirParametro("TITEMCATCLIFIM", CategoriaClienteAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125685

    'Inclui a região
    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAINIC", RegiaoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125686

    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAFIM", RegiaoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125687
    
    'Inclui a data
    If DataDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataDe.Text)
        dInic = StrParaDate(DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 125688
    
    If DataAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 125689
    
    If bExec Then
    
        lErro = RelRentabilidadeCli_Prepara(lNumIntRel, StrParaLong(sCliente_De), StrParaLong(sCliente_Ate), Month(dInic), Year(dInic), StrParaInt(RegiaoDe.Text), StrParaInt(RegiaoAte.Text), giFilialEmpresa, CategoriaClienteTodas.Value, CategoriaCliente.Text, CategoriaClienteDe.Text, CategoriaClienteAte.Text, dCustoFretesGeral, dCustoComissGeral)
        If lErro <> SUCESSO Then gError 125680
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 125680
    
        lErro = objRelOpcoes.IncluirParametro("NCUSTFRETEGERAL", CStr(dCustoFretesGeral))
        If lErro <> AD_BOOL_TRUE Then gError 125680
            
        lErro = objRelOpcoes.IncluirParametro("NCUSTCOMISSGERAL", CStr(dCustoComissGeral))
        If lErro <> AD_BOOL_TRUE Then gError 125680
        
    End If
    'Faz a chamada da função que irá montar a expressão
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCliente_De, sCliente_Ate)
    If lErro <> SUCESSO Then gError 125690
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 125678 To 125690
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179586)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_De As String, sCliente_Ate As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
   
   'Verifica se o Cliente inicial foi preenchido
    If ClienteDe.Text <> "" Then
        sCliente_De = CStr(LCodigo_Extrai(ClienteDe.Text))
    Else
        sCliente_De = ""
    End If
    
    'Verifica se o Cliente Final foi preenchido
    If ClienteAte.Text <> "" Then
        sCliente_Ate = CStr(LCodigo_Extrai(ClienteAte.Text))
    Else
        sCliente_Ate = ""
    End If
            
    'Verifica se o Cliente Inicial é menor que o final, se não for --> ERRO
    If sCliente_De <> "" And sCliente_Ate <> "" Then
        
        If CInt(sCliente_De) > CInt(sCliente_Ate) Then gError 125691
        
    End If
    
    'Cliente De não pode ser maior que o Cliente Até
    If Len(Trim(CategoriaClienteDe.Text)) <> 0 And Len(Trim(CategoriaClienteAte.Text)) <> 0 Then
    
         If CategoriaClienteDe.Text > CategoriaClienteAte.Text Then gError 125692
         
    Else
        
        If Len(Trim(CategoriaClienteDe.Text)) = 0 And Len(Trim(CategoriaClienteAte.Text)) = 0 And CategoriaClienteTodas.Value = 0 Then gError 125693
    
    End If
    
    'Se RegiãoInicial e RegiãoFinal estão preenchidos
    If Len(Trim(RegiaoDe.Text)) > 0 And Len(Trim(RegiaoAte.Text)) > 0 Then
    
        'Se Região inicial for maior que Região final, erro
        If CLng(RegiaoDe.Text) > CLng(RegiaoAte.Text) Then gError 125694
        
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 125695
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
        
        Case 125691
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteDe.SetFocus
        
        Case 125692
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)
                
        Case 125693
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_NAO_INFORMADA", gErr)
        
        Case 125694
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAOVENDA_INICIAL_MAIOR", gErr)
        
        Case 125695
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179587)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCliente_De As String, sCliente_Ate As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
      
'    'Verifica se o Cliente Inicial foi preenchido
'    If sCliente_De <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvInt(CInt(sCliente_De))
'
'    End If
'
'    'Verifica se o Cliente Final foi preenchido
'    If sCliente_Ate <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvInt(CInt(sCliente_Ate))
'
'    End If
'
'    'Se a check todas não estiver setada
'    If CategoriaClienteTodas.Value = 0 Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "CategoriaCliente = " & Forprint_ConvTexto(CategoriaCliente.Text)
'
'        If CategoriaClienteDe.Text <> "" Then
'
'            If sExpressao <> "" Then sExpressao = sExpressao & " E "
'            sExpressao = sExpressao & "ItemCategoriaCliente  >= " & Forprint_ConvTexto(CategoriaClienteDe.Text)
'
'        End If
'
'        If CategoriaClienteAte.Text <> "" Then
'
'            If sExpressao <> "" Then sExpressao = sExpressao & " E "
'            sExpressao = sExpressao & "ItemCategoriaCliente <= " & Forprint_ConvTexto(CategoriaClienteAte.Text)
'
'        End If
'
'    End If
'
'    'se a região estiver preenchida
'    If RegiaoDe.Text <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "RegiaoVenda >= " & Forprint_ConvInt(Codigo_Extrai(RegiaoDe.Text))
'
'    End If
'
'    If RegiaoAte.Text <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "RegiaoVenda <= " & Forprint_ConvInt(Codigo_Extrai(RegiaoAte.Text))
'
'    End If
'
'    'se a data estiver preenchida
'    If Trim(DataDe.ClipText) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataDe.Text))
'
'    End If
'
'    If Trim(DataAte.ClipText) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataAte.Text))
'
'    End If
'
'    If sExpressao <> "" Then
'
'        objRelOpcoes.sSelecao = sExpressao
'
'    End If
    
    Monta_Expressao_Selecao = SUCESSO
    
    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179588)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim iTipo As Integer
Dim lErro As Long, dValor As Double
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 125696
    
    lErro = objRelOpcoes.ObterParametro("NPROPAGANDA", sParam)
    If lErro <> SUCESSO Then gError 125697
    
    dValor = CDbl(sParam)
    If dValor <> 0 Then
        GastosPropaganda.Text = CStr(dValor)
    Else
        GastosPropaganda.Text = ""
    End If
    
    lErro = objRelOpcoes.ObterParametro("NPERC_GASTO_FINAN", sParam)
    If lErro <> SUCESSO Then gError 125697
    
    dValor = CDbl(sParam)
    If dValor <> 0 Then
        GastosFinanceiros.Text = CStr(dValor)
    Else
        GastosFinanceiros.Text = ""
    End If
    
    'Preenche Cliente inicial
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 125697
    
    ClienteDe.Text = sParam
    Call ClienteDe_Validate(bSGECancelDummy)
    
    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 125698
    
    ClienteAte.Text = sParam
    Call ClienteAte_Validate(bSGECancelDummy)
    
    'pega parâmetro TodasCategorias e exibe
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro <> SUCESSO Then gError 125699

    CategoriaClienteTodas.Value = CInt(sParam)

    'pega parâmetro categoria de Cliente e exibe
    lErro = objRelOpcoes.ObterParametro("TCATCLI", sParam)
    If lErro <> SUCESSO Then gError 125700
        
    If sParam <> "" Then
    
        CategoriaCliente.Text = sParam
    
        CategoriaCliente.Text = sParam
        Call CategoriaCliente_Validate(bSGECancelDummy)
    
        'pega parâmetro valor inicial e exibe
        lErro = objRelOpcoes.ObterParametro("TITEMCATCLIINI", sParam)
        If lErro <> SUCESSO Then gError 125701
        
        CategoriaClienteDe.Enabled = True
        CategoriaClienteDe.Text = sParam
        Call Combo_Seleciona(CategoriaClienteDe, iTipo)
        
        
        'pega parâmetro Valor Final e exibe
        lErro = objRelOpcoes.ObterParametro("TITEMCATCLIFIM", sParam)
        If lErro <> SUCESSO Then gError 125702
    
        CategoriaClienteAte.Text = sParam
        CategoriaClienteAte.Enabled = True
    Else
    
        CategoriaClienteTodas.Value = 1
    
    End If
    
    'pega Região de Venda Inicial e exibe
    'sParam = String(255, 0)
    lErro = objRelOpcoes.ObterParametro("TREGIAOVENDAINIC", sParam)
    If lErro <> SUCESSO Then gError 125703

    RegiaoDe.Text = sParam
    Call RegiaoDe_Validate(bSGECancelDummy)
    
    'pega Região de Venda Final e exibe
    lErro = objRelOpcoes.ObterParametro("TREGIAOVENDAFIM", sParam)
    If lErro <> SUCESSO Then gError 125704

    RegiaoAte.Text = sParam
    Call RegiaoAte_Validate(bSGECancelDummy)
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then gError 125705

    Call DateParaMasked(DataDe, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 125706

    Call DateParaMasked(DataAte, CDate(sParam))
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 125696 To 125706
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179589)

    End Select

    Exit Function

End Function
'*** FUNÇÕES DE APOIO À TELA - FIM ***

'*** FUNÇÕES DO BROWSER - INÍCIO ***
Private Sub objEventoClienteDe_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    ClienteDe.Text = CStr(objCliente.lCodigo)
    Call ClienteDe_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteAte_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    ClienteAte.Text = CStr(objCliente.lCodigo)
    Call ClienteAte_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoRegiaoVendaDe_evSelecao(obj1 As Object)

Dim objRegiaoVenda As New ClassRegiaoVenda

    Set objRegiaoVenda = obj1
    
    RegiaoDe.Text = objRegiaoVenda.iCodigo
    RegiaoDeDesc.Caption = objRegiaoVenda.sDescricao

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoRegiaoVendaAte_evSelecao(obj1 As Object)

Dim objRegiaoVenda As New ClassRegiaoVenda

    Set objRegiaoVenda = obj1
    
    RegiaoAte.Text = objRegiaoVenda.iCodigo
    RegiaoAteDesc.Caption = objRegiaoVenda.sDescricao

    Me.Show

    Exit Sub

End Sub
'*** FUNÇÕES DO BROWSER - FIM ***

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoClienteDe = Nothing
    Set objEventoClienteAte = Nothing
        
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Análise Rentabilidade de Cliente por Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRentCliProduto"
    
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

Public Sub GastosPropaganda_Validate(Cancel As Boolean)

Dim lErro As String
Dim sGastosPropaganda As String

On Error GoTo Erro_GastosPropaganda_Validate

    'Verifica se foi preenchido a GastosPropaganda
    If Len(Trim(GastosPropaganda.Text)) = 0 Then Exit Sub

    'testa para ver se é uma porcentagem valida
    lErro = Porcentagem_Critica(GastosPropaganda.Text)
    If lErro <> SUCESSO Then Error 33958

    sGastosPropaganda = GastosPropaganda.Text

    GastosPropaganda.Text = Format(sGastosPropaganda, "Fixed")

    Exit Sub

Erro_GastosPropaganda_Validate:

    Cancel = True

    Select Case Err

        Case 33958

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179590)

    End Select

    Exit Sub

End Sub

Public Sub GastosFinanceiros_Validate(Cancel As Boolean)

Dim lErro As String
Dim sGastosFinanceiros As String

On Error GoTo Erro_GastosFinanceiros_Validate

    'Verifica se foi preenchido a GastosFinanceiros
    If Len(Trim(GastosFinanceiros.Text)) = 0 Then Exit Sub

    'testa para ver se é uma porcentagem valida
    lErro = Porcentagem_Critica(GastosFinanceiros.Text)
    If lErro <> SUCESSO Then Error 33958

    sGastosFinanceiros = GastosFinanceiros.Text

    GastosFinanceiros.Text = Format(sGastosFinanceiros, "Fixed")

    Exit Sub

Erro_GastosFinanceiros_Validate:

    Cancel = True

    Select Case Err

        Case 33958

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179591)

    End Select

    Exit Sub

End Sub

Public Function RelRentabilidadeCli_Prepara(lNumIntRel As Long, ByVal lClienteDe As Long, ByVal lClienteAte As Long, ByVal iMes As Integer, ByVal iAno As Integer, ByVal iRegiaoDe As Integer, ByVal iRegiaoAte As Integer, ByVal iFilialEmpresa As Integer, ByVal bCategoriaTodas As Boolean, ByVal sCategoria As String, ByVal sCategoria_I As String, ByVal sCategoria_F As String, dCustoFretesGeral As Double, dCustoComissGeral As Double) As Long
'Realiza a leitura do GastoContrato, ValorComissão, ValoFaturado e do CustoReposição
'através das informações passadas como parâmetro
'Preenche a tabela RentabilidadeCli

Dim lErro As Long
Dim lTransacao As Long, alComando(1 To 8) As Long
Dim dValorComissao As Double, dValorFaturado As Double
Dim dCustoReposicao As Double
Dim dtDataInicio As Date, dtDataFim As Date
Dim lNumIntDoc As Long, dCustoFretes As Double, dCustoContratos As Double
Dim iIndice As Integer
Dim lCliente As Long
Dim sMes As String
Dim sProduto As String
Dim dPercentual As Double
Dim sSQL As String 'Inserido por Wagner

On Error GoTo Erro_RelRentabilidadeCli_Prepara
    
    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 128180

    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 128181
    
    dtDataInicio = CDate("01/" & iMes & "/" & iAno)
    'Alterado por Wagner
    dtDataFim = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & iMes & "/" & iAno)))
    
    'Modifica o tipo do mês
    sMes = CStr(iMes)
    
    'Obtêm o NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_RENTABILIDADECLI", lNumIntRel)
    If lErro <> SUCESSO Then gError 128182

    '??? completar filtro de clientes em funçao dos outros parametros, talvez já filtrar filiais

    'Alterado por Wagner
    '############################################
    'Faz a seleção do Código do Cliente
    
    Call RelRentabilidadeCliSQL_Prepara(lClienteDe, lClienteAte, iRegiaoDe, iRegiaoAte, bCategoriaTodas, sCategoria, sCategoria_I, sCategoria_F, sSQL)
    
    lErro = RelRentabilidadeCliInt_Prepara(alComando(1), lCliente, lClienteDe, lClienteAte, iRegiaoDe, iRegiaoAte, bCategoriaTodas, sCategoria, sCategoria_I, sCategoria_F, sSQL)
'    If lClienteAte = 0 Then lClienteAte = 9999999
'    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Clientes WHERE Codigo BETWEEN ? AND ? ORDER BY Codigo", lCliente, lClienteDe, lClienteAte)
    If lErro <> AD_SQL_SUCESSO Then gError 128183
    '############################################

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128184

    Do While lErro = AD_SQL_SUCESSO
            
        'Realiza o somatório da Comissão
        lErro = Comando_Executar(alComando(2), "SELECT SUM(ComissoesNF.Valor) FROM ComissoesNF, NFiscal WHERE ComissoesNF.NumIntDoc = NFiscal.NumIntDoc AND NFiscal.FilialEmpresa = ? AND NFiscal.Cliente = ? AND NFiscal.DataEmissao BETWEEN ? AND ? AND NFiscal.Status <> 7", dValorComissao, iFilialEmpresa, lCliente, dtDataInicio, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 128185
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128186
    
        dCustoComissGeral = dCustoComissGeral + dValorComissao
        
        'Busca o percentual na tabela ContratoPropaganda
        lErro = Comando_Executar(alComando(3), "SELECT Percentual FROM ContratoPropaganda WHERE FilialEmpresa = ? AND Cliente = ? AND PeriodoDe <= ? AND PeriodoAte >= ?", dPercentual, iFilialEmpresa, lCliente, dtDataInicio, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 128187
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128188
        
        lErro = Comando_Executar(alComando(8), "SELECT SUM(ConhecTransp.ValorTotal) FROM NFiscal NFVenda, NFiscal ConhecTransp WHERE NFVenda.Status <> ? AND ConhecTransp.Status <> ? AND ConhecTransp.NumIntNotaOriginal = NFVenda.NumIntDoc AND ConhecTransp.TipoNFiscal = 105 AND NFVenda.DataEmissao BETWEEN ? AND ? AND NFVenda.FilialEmpresa = ? AND NFVenda.Cliente = ?", _
            dCustoFretes, STATUS_CANCELADO, STATUS_CANCELADO, dtDataInicio, dtDataFim, iFilialEmpresa, lCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 128187
        
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128188
        
        dCustoFretesGeral = dCustoFretesGeral + dCustoFretes
        
        'Realiza a seleção de Produtos
        sProduto = String(STRING_PRODUTO, 0)
        lErro = Comando_Executar(alComando(4), "SELECT DISTINCT ItensNFiscal.Produto FROM NFiscal, ItensNFiscal WHERE NFiscal.Cliente = ? AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.FilialEmpresa = ? AND NFiscal.DataEmissao BETWEEN ? AND ? AND NFiscal.Status <> 7 ORDER BY ItensNFiscal.Produto", sProduto, lCliente, iFilialEmpresa, dtDataInicio, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 128189
        
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128190
        
        Do While lErro = AD_SQL_SUCESSO
        
            'Busca o ValorFaturado do mês passado
            lErro = Comando_Executar(alComando(5), "SELECT SUM(ValorFaturado" & sMes & ") FROM SldMesFatFilCli WHERE FilialEmpresa = ? AND Ano = ? AND Cliente = ? AND Produto = ?", dValorFaturado, iFilialEmpresa, iAno, lCliente, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 128191
            
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128192
            
            'Busca o CustoReposição do mês passado
            lErro = Comando_Executar(alComando(6), "SELECT CustoReposicao" & sMes & " FROM SldMesEst WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ?", dCustoReposicao, iFilialEmpresa, iAno, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 128193
            
            lErro = Comando_BuscarPrimeiro(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128194
            
            'Realiza a gravação na Tabela RentabilidadeCli
            lErro = Comando_Executar(alComando(7), "INSERT INTO RelAnaliseRentabilidade (NumIntRel,Cliente,Produto,CustoFretes,CustoContratos,CustoComissoes,CustoReposicao,ValorFaturado) VALUES (?,?,?,?,?,?,?,?)", _
                lNumIntRel, lCliente, sProduto, dCustoFretes, dValorFaturado * dPercentual, dValorComissao, dCustoReposicao, dValorFaturado)
            If lErro <> AD_SQL_SUCESSO Then gError 128196
        
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128195
            
        Loop
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128197

    Loop
    
    'Fecha a Transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 128198

    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelRentabilidadeCli_Prepara = SUCESSO

    Exit Function
    
Erro_RelRentabilidadeCli_Prepara:

    RelRentabilidadeCli_Prepara = gErr
    
    Select Case gErr
    
        Case 128180
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 128181
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 128182

        Case 128183, 128184, 128197
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
            
        Case 128185, 128186
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESNF", gErr)
            
        Case 128187, 128188
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTRATOPROPAGANDA", gErr)
            
        Case 128189, 128190, 128195
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 128191, 128192
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFAT", gErr)
            
        Case 128193, 128194
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOMES", gErr)
            
        Case 128196
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELRENTABILIDADECLI", gErr)
            
        Case 128198
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179592)

    End Select

    Call Transacao_Rollback

    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Sub RelRentabilidadeCliSQL_Prepara(ByVal vlClienteDe As Variant, ByVal vlClienteAte As Variant, ByVal viRegiaoDe As Variant, ByVal viRegiaoAte As Variant, ByVal vbCategoriaTodas As Variant, ByVal vsCategoria As Variant, ByVal vsCategoria_I As Variant, ByVal vsCategoria_F As Variant, sSQL As String)
'monta o comando SQL para obtencao das fretes dinamicamente e retorna.
Dim sSelect As String, sWhere As String, sFrom As String, sOrderBy As String, sGroupBy As String

On Error GoTo Erro_RelRentabilidadeCliSQL_Prepara

    sSelect = "SELECT  Cli.Codigo "
       
    sGroupBy = "GROUP BY Cli.Codigo "
     
    sOrderBy = "ORDER BY Cli.Codigo "
                             
    If Len(Trim(vsCategoria)) > 0 Or Len(Trim(vsCategoria_I)) > 0 Or Len(Trim(vsCategoria_F)) > 0 Then
        
        sFrom = "FROM Clientes AS Cli , " & _
                        "FiliaisClientes AS FilCli, " & _
                        "FilialClienteCategorias AS FilCliCat "
        sWhere = "WHERE FilCli.CodCliente = FilCliCat.Cliente AND " & _
                        " FilCli.CodFilial = FilCliCat.Filial AND " & _
                        "Cli.Codigo = FilCli.CodCliente "
                        
    ElseIf viRegiaoDe <> 0 Or viRegiaoAte <> 0 Then
    
        sFrom = " FROM Clientes AS Cli , " & _
                    "FiliaisClientes AS FilCli "
        sWhere = " WHERE Cli.Codigo = FilCli.CodCliente "

    Else
    
        sFrom = "FROM Clientes AS Cli "
                     
        sWhere = "WHERE 1 = 1 "
    
    End If
    
    
    If viRegiaoDe <> 0 Then
        sWhere = sWhere & "AND FilCli.Regiao >= ? "
    End If
    
    If viRegiaoAte <> 0 Then
        sWhere = sWhere & "AND FilCli.Regiao <= ? "
    End If
   
    If vlClienteDe <> 0 Then
        sWhere = sWhere & "AND Cli.Codigo >= ? "
    End If
    
    If vlClienteAte <> 0 Then
        sWhere = sWhere & "AND Cli.Codigo <= ? "
    End If
    
    If Not vbCategoriaTodas Then
        sWhere = sWhere & "AND FilCliCat.Categoria = ? "
    End If
    
    If Len(Trim(vsCategoria_I)) > 0 Then
        sWhere = sWhere & "AND FilCliCat.Item >= ? "
    End If
    
    If Len(Trim(vsCategoria_F)) > 0 Then
        sWhere = sWhere & "AND FilCliCat.Item <= ? "
    End If
        
    sSQL = sSelect & sFrom & sWhere & sGroupBy & sOrderBy

    Exit Sub

Erro_RelRentabilidadeCliSQL_Prepara:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179593)

    End Select

    Exit Sub

End Sub

Private Function RelRentabilidadeCliInt_Prepara(ByVal lComando As Long, vlCliente As Variant, ByVal vlClienteDe As Variant, ByVal vlClienteAte As Variant, ByVal viRegiaoDe As Variant, ByVal viRegiaoAte As Variant, ByVal vbCategoriaTodas As Variant, ByVal vsCategoria As Variant, ByVal vsCategoria_I As Variant, ByVal vsCategoria_F As Variant, ByVal sSQL As String) As Long

Dim lErro As Long

On Error GoTo Erro_RelRentabilidadeCliInt_Prepara

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129200
    
    lErro = Comando_BindVarInt(lComando, vlCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129201

    If viRegiaoDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, viRegiaoDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129202
    End If
    
    If viRegiaoAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, viRegiaoAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129203
    End If
   
    If vlClienteDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlClienteDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129204
    End If
    
    If vlClienteAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlClienteAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129205
    End If
    
    If Not vbCategoriaTodas Then
        lErro = Comando_BindVarInt(lComando, vsCategoria_I)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129206
    End If
    
    If Len(Trim(vsCategoria_I)) > 0 Then
        lErro = Comando_BindVarInt(lComando, vsCategoria_I)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129207
    End If
    
    If Len(Trim(vsCategoria_F)) > 0 Then
        lErro = Comando_BindVarInt(lComando, vsCategoria_F)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129208
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129209
    
    RelRentabilidadeCliInt_Prepara = SUCESSO

    Exit Function

Erro_RelRentabilidadeCliInt_Prepara:

    RelRentabilidadeCliInt_Prepara = gErr

    Select Case gErr
    
        Case 129200 To 129209

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179594)

    End Select

    Exit Function

End Function


