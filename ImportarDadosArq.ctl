VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ImportarDadosArqOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      Caption         =   "Campos"
      Height          =   5220
      Index           =   0
      Left            =   60
      TabIndex        =   16
      Top             =   750
      Width           =   9420
      Begin VB.Frame Frame2 
         Caption         =   "Dados Adicionais"
         Height          =   1065
         Index           =   10
         Left            =   45
         TabIndex        =   45
         Top             =   4065
         Visible         =   0   'False
         Width           =   8040
         Begin VB.TextBox CodPrevVenda 
            Height          =   315
            Left            =   915
            MaxLength       =   10
            TabIndex        =   46
            Top             =   330
            Width           =   1380
         End
         Begin MSMask.MaskEdBox AnoPrevVenda 
            Height          =   315
            Left            =   3225
            TabIndex        =   47
            Top             =   330
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ano:"
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
            Index           =   3
            Left            =   2760
            TabIndex        =   49
            Top             =   390
            Width           =   405
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
            Left            =   225
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   48
            Top             =   390
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados Adicionais"
         Height          =   1065
         Index           =   20
         Left            =   45
         TabIndex        =   39
         Top             =   4080
         Visible         =   0   'False
         Width           =   8040
         Begin VB.Frame Frame 
            Caption         =   "Data Emissão"
            Height          =   750
            Index           =   0
            Left            =   990
            TabIndex        =   40
            Top             =   210
            Width           =   5520
            Begin VB.ComboBox Mes 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "ImportarDadosArq.ctx":0000
               Left            =   900
               List            =   "ImportarDadosArq.ctx":002B
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   270
               Width           =   2310
            End
            Begin MSMask.MaskEdBox Ano 
               Height          =   315
               Left            =   4260
               TabIndex        =   42
               Top             =   255
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Mês:"
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
               TabIndex        =   44
               Top             =   315
               Width           =   405
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Ano:"
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
               Left            =   3735
               TabIndex        =   43
               Top             =   315
               Width           =   405
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados Adicionais"
         Height          =   1065
         Index           =   5
         Left            =   45
         TabIndex        =   30
         Top             =   4095
         Visible         =   0   'False
         Width           =   8040
         Begin VB.CommandButton BotaoAtualizarTabela 
            Height          =   330
            Left            =   7485
            Picture         =   "ImportarDadosArq.ctx":00D0
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Atualizar a Lista de Almoxarifados"
            Top             =   225
            Width           =   420
         End
         Begin VB.CommandButton BotaoCriarTabela 
            Caption         =   "Criar"
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
            Left            =   6195
            TabIndex        =   34
            Top             =   240
            Width           =   1200
         End
         Begin VB.ComboBox Tabela 
            Height          =   315
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   255
            Width           =   960
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   2040
            TabIndex        =   36
            Top             =   645
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   3135
            TabIndex        =   37
            Top             =   660
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data de Vigência:"
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
            Index           =   2
            Left            =   390
            TabIndex        =   38
            Top             =   705
            Width           =   1530
         End
         Begin VB.Label TabelaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tabela:"
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
            Height          =   210
            Left            =   300
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   33
            Top             =   285
            Width           =   660
         End
         Begin VB.Label DescricaoTabela 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2040
            TabIndex        =   32
            Top             =   255
            Width           =   4125
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados Adicionais"
         Height          =   1065
         Index           =   4
         Left            =   45
         TabIndex        =   25
         Top             =   4095
         Visible         =   0   'False
         Width           =   8040
         Begin VB.ComboBox Almoxarifado 
            Height          =   315
            ItemData        =   "ImportarDadosArq.ctx":0522
            Left            =   1365
            List            =   "ImportarDadosArq.ctx":052F
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   405
            Width           =   2745
         End
         Begin VB.CommandButton BotaoCriarAlmox 
            Caption         =   "Criar"
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
            Left            =   4095
            TabIndex        =   27
            Top             =   390
            Width           =   1200
         End
         Begin VB.CommandButton BotaoAtualizarAlmox 
            Height          =   330
            Left            =   5385
            Picture         =   "ImportarDadosArq.ctx":0560
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Atualizar a Lista de Almoxarifados"
            Top             =   375
            Width           =   420
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Almoxarifado:"
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
            Left            =   120
            TabIndex        =   29
            Top             =   450
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Dados Adicionais"
         Height          =   1065
         Index           =   0
         Left            =   45
         TabIndex        =   24
         Top             =   4095
         Visible         =   0   'False
         Width           =   8040
      End
      Begin VB.ComboBox Campo 
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   840
         Width           =   5000
      End
      Begin VB.CommandButton BotaoConfirmarCampos 
         Caption         =   "Confirmar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   8085
         TabIndex        =   5
         Top             =   4140
         Width           =   1260
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -23270
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1560
         Width           =   1725
      End
      Begin MSFlexGridLib.MSFlexGrid GridCampos 
         Height          =   2325
         Left            =   405
         TabIndex        =   4
         Top             =   210
         Width           =   8490
         _ExtentX        =   14975
         _ExtentY        =   4101
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   300
         Left            =   -10000
         TabIndex        =   18
         Top             =   1560
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados"
      Height          =   5235
      Index           =   1
      Left            =   60
      TabIndex        =   13
      Top             =   750
      Visible         =   0   'False
      Width           =   9390
      Begin VB.CommandButton BotaoGravar 
         Caption         =   "Importar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7875
         TabIndex        =   8
         Top             =   4800
         Width           =   1380
      End
      Begin VB.TextBox Dado 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   885
         TabIndex        =   23
         Top             =   1125
         Width           =   1185
      End
      Begin VB.CommandButton BotaoVoltar 
         Caption         =   "Voltar"
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
         Left            =   6870
         TabIndex        =   7
         Top             =   4815
         Width           =   915
      End
      Begin VB.ComboBox FilialEmpresa 
         Height          =   315
         Left            =   -23270
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1560
         Width           =   1725
      End
      Begin MSFlexGridLib.MSFlexGrid GridDados 
         Height          =   2325
         Left            =   45
         TabIndex        =   6
         Top             =   210
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   4101
         _Version        =   393216
         AllowUserResizing=   1
      End
      Begin MSMask.MaskEdBox NumEtiqueta 
         Height          =   300
         Left            =   -10000
         TabIndex        =   14
         Top             =   1560
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Status 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   750
         TabIndex        =   22
         Top             =   4830
         Width           =   6090
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   4875
         Width           =   615
      End
   End
   Begin VB.CheckBox optDescPriLinha 
      Caption         =   "A Primeira Linha Contém a Descrição"
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
      Left            =   4560
      TabIndex        =   1
      Top             =   135
      Value           =   1  'Checked
      Width           =   3645
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      Left            =   1155
      TabIndex        =   0
      Top             =   75
      Width           =   3315
   End
   Begin VB.CommandButton BotaoProcurar 
      Caption         =   "..."
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
      Left            =   7725
      TabIndex        =   3
      Top             =   375
      Width           =   555
   End
   Begin VB.TextBox NomeArquivo 
      Height          =   285
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   420
      Width           =   6570
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8415
      ScaleHeight     =   495
      ScaleWidth      =   1005
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   75
      Width           =   1065
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   45
         Picture         =   "ImportarDadosArq.ctx":09B2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   525
         Picture         =   "ImportarDadosArq.ctx":0EE4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   615
      TabIndex        =   19
      Top             =   135
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo:"
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
      Left            =   345
      TabIndex        =   15
      Top             =   450
      Width           =   720
   End
End
Attribute VB_Name = "ImportarDadosArqOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis globais
Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim iTipoArqAnt As Integer
Dim sNomeArqAnt As String
Dim iPriLinha As Integer

Dim objGridDados As AdmGrid
Dim iGrid_Coluna_Col As Integer
Dim iGrid_Campo_Col As Integer

Dim gobjTabela As ClassImportTabelas

Dim objGridCampos As AdmGrid

Const NUM_COLUNAS_DADOS = 400

Const DATA_APAGA = #1/1/1822#
Const STRING_APAGA = "-"

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'instancia as variáveis globais
    Set objGridCampos = New AdmGrid
    
    'Inicializa o Grid
    lErro = Inicializa_Grid_Campos(objGridCampos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("Carrega_Combo", Tipo, "ImportTabelas", "Codigo", TIPO_LONG, "Descricao", TIPO_STR)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
          
    sNomeArqAnt = NomeArquivo.Text
    iPriLinha = optDescPriLinha.Value
    iTipoArqAnt = Codigo_Extrai(Tipo.Text)
    iFrameAtual = iTipoArqAnt
    
    lErro = Carrega_ComboAlmox()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Carrega_TabelaPreco()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call DateParaMasked(Data, Date)
    
    Set gobjTabela = New ClassImportTabelas
       
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208920)

    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros(Optional ByVal objObjeto As Object) As Long

Dim lErro As Long
Dim sClasse As String, iTipoArq As Integer
Dim objTabelaPrecoItem As ClassTabelaPrecoItem

On Error GoTo Erro_Trata_Parametros

    sClasse = UCase(TypeName(objObjeto))
    iTipoArq = 0
    
    Select Case sClasse
    
        Case "CLASSTABELAPRECOITEM"
            Set objTabelaPrecoItem = objObjeto
            iTipoArq = 5
            If objTabelaPrecoItem.iCodTabela <> 0 Then
                Call Combo_Seleciona_ItemData(Tabela, objTabelaPrecoItem.iCodTabela)
                Call DateParaMasked(Data, objTabelaPrecoItem.dtDataVigencia)
            End If
            
        Case "CLASSRELACCLIENTES"
            iTipoArq = 25
    
    End Select
    
    If iTipoArq <> 0 Then
        Call Combo_Seleciona_ItemData(Tipo, iTipoArq)
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208962)

    End Select

    Exit Function
    
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Libera os objetos e coleções globais
    Set objGridDados = Nothing
    Set objGridCampos = Nothing
    Set gobjTabela = Nothing
    
End Sub

Private Sub BotaoLimpar_Click()
'Dispara a limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se algum campo foi alterado e confirma se o usuário deseja
    'salvar antes de limpar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'limpa a tela
    Call Limpa_Tela_ImportarDados

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208921)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

'*** FUNCIONAMENTO DO GridArquivos - INÍCIO ***

'***** EVENTOS DO GRID - INÍCIO *******
Private Sub GridDados_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDados, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDados, iAlterado)
    End If

End Sub

Private Sub GridDados_EnterCell()
    Call Grid_Entrada_Celula(objGridDados, iAlterado)
End Sub

Private Sub GridDados_GotFocus()
    Call Grid_Recebe_Foco(objGridDados)
End Sub

Private Sub GridDados_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridDados)
End Sub

Private Sub GridDados_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDados, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDados, iAlterado)
    End If

End Sub

Private Sub GridDados_LeaveCell()
    Call Saida_Celula(objGridDados)
End Sub

Private Sub GridDados_RowColChange()
    Call Grid_RowColChange(objGridDados)
End Sub

Private Sub GridDados_Scroll()
    Call Grid_Scroll(objGridDados)
End Sub

Private Sub GridDados_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridDados)
End Sub

Private Sub GridCampos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCampos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCampos, iAlterado)
    End If

End Sub

Private Sub GridCampos_EnterCell()
    Call Grid_Entrada_Celula(objGridCampos, iAlterado)
End Sub

Private Sub GridCampos_GotFocus()
    Call Grid_Recebe_Foco(objGridCampos)
End Sub

Private Sub GridCampos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridCampos)
End Sub

Private Sub GridCampos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCampos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCampos, iAlterado)
    End If

End Sub

Private Sub GridCampos_LeaveCell()
    Call Saida_Celula(objGridCampos)
End Sub

Private Sub GridCampos_RowColChange()
    Call Grid_RowColChange(objGridCampos)
End Sub

Private Sub GridCampos_Scroll()
    Call Grid_Scroll(objGridCampos)
End Sub

Private Sub GridCampos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridCampos)
End Sub

Private Sub Campo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Campo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCampos)
End Sub

Private Sub Campo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCampos)
End Sub

Private Sub Campo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCampos.objControle = Campo
    lErro = Grid_Campo_Libera_Foco(objGridCampos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGrid As AdmGrid) As Long
'faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGrid)
    If lErro = SUCESSO Then

        If objGrid.objGrid.Name = GridCampos.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGrid.objGrid.Col
    
                Case iGrid_Campo_Col
                    lErro = Saida_Celula_Padrao(objGrid, Campo)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    
            End Select
            
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGrid)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208922)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(sSQL As String) As Long

Dim lErro As Long
Dim sCampos As String
Dim sParam As String
Dim objTabelaCampo As ClassImportTabelasCampos
Dim colLinhas As New Collection
Dim colColunas As Collection
Dim iLinha As Integer, sValor As String
Dim avValor() As Variant, lIndice As Long

On Error GoTo Erro_Move_Tela_Memoria

    For Each objTabelaCampo In gobjTabela.colCampos
        If sCampos <> "" Then
            sCampos = sCampos & ","
            sParam = sParam & ","
        End If
        sCampos = sCampos & objTabelaCampo.sCampo
        sParam = sParam & "?"
    Next
    
    ReDim avValor(0 To objGridDados.iLinhasExistentes * gobjTabela.colCampos.Count) As Variant
    
    lIndice = 0
    For iLinha = 1 To objGridDados.iLinhasExistentes
        Set colColunas = New Collection
        colLinhas.Add colColunas
        For Each objTabelaCampo In gobjTabela.colCampos
            sValor = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
            lIndice = lIndice + 1
            Select Case objTabelaCampo.iTipo
                Case TIPO_TEXTO
                    avValor(lIndice) = sValor
                Case TIPO_DATA
                    If Trim(sValor) = STRING_APAGA Then
                        avValor(lIndice) = DATA_APAGA
                    Else
                        avValor(lIndice) = CDate(StrParaDate(sValor))
                    End If
                Case TIPO_DECIMAL
                    avValor(lIndice) = CDbl(StrParaDbl(sValor))
                Case TIPO_NUMERICO
                    avValor(lIndice) = CLng(StrParaLong(sValor))
                Case TIPO_INVALIDO
                    If sValor = "" Then
                        avValor(lIndice) = -1
                    Else
                        avValor(lIndice) = CDbl(StrParaDbl(sValor))
                    End If
            End Select
            colColunas.Add avValor(lIndice)
        Next
    Next
    
    Set gobjTabela.colcolDados = colLinhas
    
    sSQL = "INSERT INTO " & gobjTabela.sTabela & "(" & sCampos & ") VALUES (" & sParam & ") "

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208923)

    End Select

End Function

Private Sub Limpa_Tela_ImportarDados()
'Limpa a tela com exceção do campo 'Modelo'

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_ImportarDados

    'Limpa os controles básicos da tela
    Call Limpa_Tela(Me)
    
    If Not (objGridDados Is Nothing) Then Call Grid_Limpa(objGridDados)
    Call Grid_Limpa(objGridCampos)
    
    Mes.ListIndex = -1
    
    optDescPriLinha.Value = vbChecked
    
    Tipo.ListIndex = -1
    
    Campo.Clear
    Status.Caption = ""
       
    sNomeArqAnt = NomeArquivo.Text
    iPriLinha = optDescPriLinha.Value
    iTipoArqAnt = Codigo_Extrai(Tipo.Text)

    Frame2(iFrameAtual).Visible = False
    iFrameAtual = Frame_Tipo(iTipoArqAnt)
    
    Frame1(0).Visible = True
    Frame1(1).Visible = False
    
    For iIndice = 1 To GridCampos.Rows - 1
        GridCampos.TextMatrix(iIndice, iGrid_Coluna_Col) = "Coluna" & CStr(iIndice)
    Next
    
    Call DateParaMasked(Data, Date)
    
    Set gobjTabela = New ClassImportTabelas

    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_ImportarDados:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208924)

    End Select
    
    Exit Sub
    
End Sub

Private Function Inicializa_Grid_Campos(objGrid As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Campos

    Set objGrid.objForm = Me

    'entitula as colunas
    objGrid.colColuna.Add "Coluna"
    objGrid.colColuna.Add "Campo"

    'guarda os nomes dos campos
    objGrid.colCampo.Add Campo.Name

    'inicializa os índices das colunas
    iGrid_Coluna_Col = 0
    iGrid_Campo_Col = 1

    'configura os atributos
    GridCampos.ColWidth(0) = 3000
    
    GridCampos.Rows = 100 + 1

    'vincula o grid da tela propriamente dito ao controlador de grid
    objGrid.objGrid = GridCampos

    'configura sua visualização
    objGrid.iLinhasVisiveis = 10
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'inicializa o grid
    Call Grid_Inicializa(objGrid)
    
    For iIndice = 1 To GridCampos.Rows - 1
        GridCampos.TextMatrix(iIndice, iGrid_Coluna_Col) = "Coluna" & CStr(iIndice)
    Next

    Inicializa_Grid_Campos = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Campos:

    Inicializa_Grid_Campos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208925)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Dados(objGrid As AdmGrid, ByVal iNumLinhas As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Dados

    Set objGrid.objForm = Me

    'entitula as colunas
    objGrid.colColuna.Add ""
    
    For iIndice = 1 To gobjTabela.colCampos.Count
        objGrid.colColuna.Add "Campo" & CStr(iIndice)
    Next

    'guarda os nomes dos Dados
    For iIndice = 1 To gobjTabela.colCampos.Count
        objGrid.colCampo.Add Dado.Name
    Next

    'configura os atributos
    GridDados.ColWidth(0) = 400
    
    If iNumLinhas > 100 Then
        GridDados.Rows = iNumLinhas + 1
    Else
        GridDados.Rows = 100 + 1
    End If

    'vincula o grid da tela propriamente dito ao controlador de grid
    objGrid.objGrid = GridDados

    'configura sua visualização
    objGrid.iLinhasVisiveis = 16
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'inicializa o grid
    Call Grid_Inicializa(objGrid)

    Inicializa_Grid_Dados = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Dados:

    Inicializa_Grid_Dados = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208926)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Importação de arquivo"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ImportarDadosArq"

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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property
'**** fim do trecho a ser copiado *****

Private Sub BotaoProcurar_Click()
   
On Error GoTo Erro_BotaoProcurar_Click
    
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "Excel 2007(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls|Calc(*.ods)|*.ods"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
    NomeArquivo.Text = CommonDialog1.FileName
    Call Trata_Arquivo
    
    Exit Sub

Erro_BotaoProcurar_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Private Function Frame_Tipo(ByVal iTipo As Integer) As Integer
    If Frame_Existe(iTipo) Then
        Frame_Tipo = iTipo
    Else
        Frame_Tipo = 0
    End If
    If iTipo >= 19 And iTipo <= 23 Then
        Frame_Tipo = 20
    End If
End Function

Private Function Frame_Existe(ByVal iTipo As Integer) As Boolean

On Error GoTo Erro_Frame_Existe

    If Frame2(iTipo).Visible Then
        Frame2(iTipo).Visible = True
    End If

    Frame_Existe = True

    Exit Function
    
Erro_Frame_Existe:

    Frame_Existe = False

    Exit Function

End Function

Private Function Trata_Arquivo() As Long

Dim lErro As Long, bTemInfo As Boolean, sValorAux As String
Dim iNumColunas As Integer, iColuna As Integer, iLinha As Integer, sValor As String
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
Dim objTabela As New ClassImportTabelas
Dim objTabelaCampo As ClassImportTabelasCampos
Dim lCampo As Long, bAlterouTipo As Boolean
Dim objExcelApp As New ClassExcelApp

On Error GoTo Erro_Trata_Arquivo

    bAlterouTipo = False

    If iTipoArqAnt <> Codigo_Extrai(Tipo.Text) Then
    
        bAlterouTipo = True
    
        iTipoArqAnt = Codigo_Extrai(Tipo.Text)
           
        Campo.Clear
        
        Frame2(iFrameAtual).Visible = False
            
        Call Grid_Limpa(objGridCampos)
        
        For iLinha = 1 To GridCampos.Rows - 1
            GridCampos.TextMatrix(iLinha, iGrid_Coluna_Col) = "Coluna" & CStr(iLinha)
        Next
            
        If iTipoArqAnt <> 0 Then
           
            iFrameAtual = Frame_Tipo(iTipoArqAnt)
            Frame2(iFrameAtual).Visible = True
            
            Frame1(0).Visible = True
            Frame1(1).Visible = False
            
            objTabela.lCodigo = iTipoArqAnt
            
            lErro = CF("ImportTabelas_Le", objTabela)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 208927
            
            Set gobjTabela = objTabela
            
            Campo.AddItem "0-Ignorar"
            Campo.ItemData(Campo.NewIndex) = 0
            
            For Each objTabelaCampo In objTabela.colCampos
                Campo.AddItem CStr(objTabelaCampo.lCodigo) & SEPARADOR & objTabelaCampo.sNomeExibicao
                Campo.ItemData(Campo.NewIndex) = objTabelaCampo.lCodigo
            Next
            
        Else
        
            Set gobjTabela = New ClassImportTabelas
        End If
    
    End If
    
    'Se mudou o arquivo ou o tipo e tanto tipo como nome do arquivo estão preenchidos
    If (sNomeArqAnt <> NomeArquivo.Text Or iPriLinha <> optDescPriLinha.Value Or bAlterouTipo) And iTipoArqAnt <> 0 And Len(Trim(NomeArquivo.Text)) > 0 Then
    
        Call Grid_Limpa(objGridCampos)
        
        For iLinha = 1 To GridCampos.Rows - 1
            GridCampos.TextMatrix(iLinha, iGrid_Coluna_Col) = "Coluna" & CStr(iLinha)
        Next
    
        sNomeArqAnt = NomeArquivo.Text
        iPriLinha = optDescPriLinha.Value
    
'        'Abre o excel
'        lErro = CF("Excel_Abrir")
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'        'Desabilita as mensagens do Excel
'        GL_objExcelSist.DisplayAlerts = False
'
'        Call GL_objExcelSist.Workbooks.Open(sNomeArqAnt)
'
'        Set objPastaTrabalho = GL_objExcelSist.ActiveWorkBook
'
'        'Seleciona a planilha ativa na pasta de trabalho criada
'        Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet

        'GL_objExcelSist.Visible = True
        
        'Abre o excel
        lErro = objExcelApp.Abrir()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objExcelApp.Abrir_Planilha(sNomeArqAnt)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        For iColuna = 1 To NUM_COLUNAS_DADOS
            'sValor = objPlanilhaExcel.cells(1, iColuna)
            sValor = objExcelApp.Obtem_Valor_Celula(1, iColuna)
            
            If Len(Trim(sValor)) > 0 Then
                If optDescPriLinha.Value = vbChecked Then GridCampos.TextMatrix(iColuna, iGrid_Coluna_Col) = sValor
            Else
                'Se nada está preenchido até a décima linha é porque já não tem mais informação
                bTemInfo = False
                For iLinha = 2 To 10
                    'sValorAux = objPlanilhaExcel.cells(iLinha, iColuna)
                    sValorAux = objExcelApp.Obtem_Valor_Celula(iLinha, iColuna)
                    If Len(Trim(sValorAux)) > 0 Then
                        bTemInfo = True
                        Exit For
                    End If
                Next
                If Not bTemInfo Then Exit For
            End If
            
            lErro = Retorna_Cod_Campo(lCampo, iColuna, sValor)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Call Combo_Seleciona_ItemData(Campo, lCampo)
            GridCampos.TextMatrix(iColuna, iGrid_Campo_Col) = Campo.Text
            
        Next
        iNumColunas = iColuna - 1
        objGridCampos.iLinhasExistentes = iNumColunas
        
        For iLinha = iNumColunas + 1 To GridCampos.Rows - 1
            GridCampos.TextMatrix(iLinha, iGrid_Coluna_Col) = ""
        Next
    
'        'Fecha o Excel
'        Call CF("Excel_Fechar")

        lErro = objExcelApp.Fechar()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    Trata_Arquivo = SUCESSO

    Exit Function

Erro_Trata_Arquivo:

    Trata_Arquivo = gErr

    Select Case gErr
    
        Case 208927
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTTABELAS_NAO_CADASTRADO", gErr, iTipoArqAnt)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208928)

    End Select
    
    Call objExcelApp.Fechar

    Exit Function

End Function

Private Function Retorna_Cod_Campo(lCampo As Long, ByVal iColuna As Integer, ByVal sValor As String) As Long

Dim bAchou As Boolean
Dim objTabelaCampo As ClassImportTabelasCampos

On Error GoTo Erro_Retorna_Cod_Campo

    lCampo = 0
    bAchou = False
    If Len(Trim(sValor)) > 0 Then
        
        For Each objTabelaCampo In gobjTabela.colCampos
            If UCase(objTabelaCampo.sNomeIgual1) = UCase(sValor) Then
                bAchou = True
                lCampo = objTabelaCampo.lCodigo
            End If
        Next
        If Not bAchou Then
            For Each objTabelaCampo In gobjTabela.colCampos
                If UCase(objTabelaCampo.sNomeIgual2) = UCase(sValor) Then
                    bAchou = True
                    lCampo = objTabelaCampo.lCodigo
                End If
            Next
        End If
        If Not bAchou Then
            For Each objTabelaCampo In gobjTabela.colCampos
                If UCase(objTabelaCampo.sNomeIgual3) = UCase(sValor) Then
                    bAchou = True
                    lCampo = objTabelaCampo.lCodigo
                End If
            Next
        End If
        If Not bAchou Then
            For Each objTabelaCampo In gobjTabela.colCampos
                If objTabelaCampo.sNomeLike1 <> "" And InStr(1, UCase(sValor), UCase(objTabelaCampo.sNomeLike1)) <> 0 Then
                    bAchou = True
                    lCampo = objTabelaCampo.lCodigo
                End If
            Next
        End If
        If Not bAchou Then
            For Each objTabelaCampo In gobjTabela.colCampos
                If objTabelaCampo.sNomeLike2 <> "" And InStr(1, UCase(sValor), UCase(objTabelaCampo.sNomeLike2)) <> 0 Then
                    bAchou = True
                    lCampo = objTabelaCampo.lCodigo
                End If
            Next
        End If
        If Not bAchou Then
            For Each objTabelaCampo In gobjTabela.colCampos
                If objTabelaCampo.sNomeLike3 <> "" And InStr(1, UCase(sValor), UCase(objTabelaCampo.sNomeLike3)) <> 0 Then
                    bAchou = True
                    lCampo = objTabelaCampo.lCodigo
                End If
            Next
        End If
    End If

    Retorna_Cod_Campo = SUCESSO

    Exit Function

Erro_Retorna_Cod_Campo:

    Retorna_Cod_Campo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208929)

    End Select

    Exit Function
    
End Function

Private Sub BotaoConfirmarCampos_Click()
    GL_objMDIForm.MousePointer = vbHourglass
    If Trata_Dados = SUCESSO Then
        Frame1(0).Visible = False
        Frame1(1).Visible = True
    End If
    GL_objMDIForm.MousePointer = vbNormal
End Sub

Private Sub BotaoVoltar_Click()
    Frame1(0).Visible = True
    Frame1(1).Visible = False
End Sub

Private Sub Tipo_Change()
    Call Trata_Arquivo
End Sub

Private Sub Tipo_Click()
    Call Trata_Arquivo
End Sub

Private Sub optDescPriLinha_Click()
    Call Trata_Arquivo
End Sub

Private Function Trata_Dados() As Long

Dim lErro As Long, iColunaExcel As Integer, iLinhaAux As Integer
Dim iColuna As Integer, iLinha As Integer, sValor As String
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
Dim objTabelaCampo As ClassImportTabelasCampos
Dim iNumLinhas  As Integer, objTela As Object, sFormato As String
Dim objExcelApp As New ClassExcelApp

On Error GoTo Erro_Trata_Dados

'    'Abre o excel
'    lErro = CF("Excel_Abrir")
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    'Desabilita as mensagens do Excel
'    GL_objExcelSist.DisplayAlerts = False
'
'    Call GL_objExcelSist.Workbooks.Open(sNomeArqAnt)
'
'    Set objPastaTrabalho = GL_objExcelSist.ActiveWorkBook
'
'    'Seleciona a planilha ativa na pasta de trabalho criada
'    Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
        
    'GL_objExcelSist.Visible = True
    
    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objExcelApp.Abrir_Planilha(sNomeArqAnt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Para cada campo
    iNumLinhas = 0
    For Each objTabelaCampo In gobjTabela.colCampos
                
        If objTabelaCampo.iObrigatorio = MARCADO Then
        
            iColunaExcel = 0
            
            'para cada coluna
            For iColuna = 1 To objGridCampos.iLinhasExistentes
                'Se é a coluna em questão
                If Codigo_Extrai(GridCampos.TextMatrix(iColuna, iGrid_Campo_Col)) = objTabelaCampo.lCodigo Then
                    iColunaExcel = iColuna
                    Exit For
                End If
            Next
            
            If iColunaExcel = 0 Then gError 208930 ' Tem que importar a chave da tabela

            iLinha = -1
            sValor = "A"
            Do While Len(Trim(sValor)) > 0
                iLinha = iLinha + 1
                'sValor = objPlanilhaExcel.cells(iLinha + 1, iColunaExcel)
                sValor = objExcelApp.Obtem_Valor_Celula(iLinha + 1, iColunaExcel)
            Loop
            If iNumLinhas < iLinha Then iNumLinhas = iLinha
            
        End If
    
    Next
    
    Set objGridDados = New AdmGrid
    
    lErro = Inicializa_Grid_Dados(objGridDados, iNumLinhas)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Grid_Limpa(objGridDados)
    
    For Each objTabelaCampo In gobjTabela.colCampos
           
        GridDados.TextMatrix(0, objTabelaCampo.lCodigo) = objTabelaCampo.sNomeExibicao
           
        iColunaExcel = 0
        
        'para cada coluna
        For iColuna = 1 To objGridCampos.iLinhasExistentes
            'Se é a coluna em questão
            If Codigo_Extrai(GridCampos.TextMatrix(iColuna, iGrid_Campo_Col)) = objTabelaCampo.lCodigo Then
                iColunaExcel = iColuna
                Exit For
            End If
        Next
            
        If iColunaExcel <> 0 Then

            iLinhaAux = 0
            For iLinha = 1 To iNumLinhas
                If iLinha <> 1 Or optDescPriLinha.Value = vbUnchecked Then
                    iLinhaAux = iLinhaAux + 1
                    'sValor = objPlanilhaExcel.cells(iLinha, iColunaExcel)
                    sValor = objExcelApp.Obtem_Valor_Celula(iLinha, iColunaExcel)

                    'sFormato = objPlanilhaExcel.cells(iLinha, iColunaExcel).numberformat
                    sFormato = objExcelApp.Obtem_NumberFormat_Celula(iLinha, iColunaExcel)
                   
                    If (Len(sFormato) - 1 <> Len(Replace(sFormato, ",", "")) Or InStr(1, sValor, ".") <> 0) And objTabelaCampo.iTipo <> 1 Then sFormato = ""
                    
                    If Len(Trim(sFormato)) > 0 Then sValor = Format(sValor, sFormato)
                    GridDados.TextMatrix(iLinhaAux, objTabelaCampo.lCodigo) = Trim(sValor)
                End If
            Next
            
        Else
            iLinhaAux = 0
            For iLinha = 1 To iNumLinhas
                If iLinha <> 1 Or optDescPriLinha.Value = vbUnchecked Then
                    iLinhaAux = iLinhaAux + 1
                    GridDados.TextMatrix(iLinhaAux, objTabelaCampo.lCodigo) = objTabelaCampo.sValorPadrao
                End If
            Next
        End If
    
    Next
    
    If optDescPriLinha.Value = vbChecked Then iNumLinhas = iNumLinhas - 1
     
    objGridDados.iLinhasExistentes = iNumLinhas
 
    'Fecha o Excel
    'Call CF("Excel_Fechar")
    
    lErro = objExcelApp.Fechar()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Status.Caption = ""
    
    Set objTela = Me
    
    'Tem que chamar a rotina de validação para ajustar os campos não importáveis e as formatações
    'Além de validar o que vai ser importado -> cada tipo de arquivo vai ter uma validação específica
    If Len(Trim(gobjTabela.sFuncaoValida)) > 0 Then
        Call CallByName(objTela, gobjTabela.sFuncaoValida, VbMethod)
    End If
    
    lErro = Valida_Dados_Padrao()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Trata_Dados = SUCESSO

    Exit Function

Erro_Trata_Dados:

    Trata_Dados = gErr

    Select Case gErr
    
        Case 208930
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTTABELAS_CAMPO_OBRIGATORIO", gErr, objTabelaCampo.sCampo, gobjTabela.sTabela)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208931)

    End Select
    
    Call objExcelApp.Fechar

    Exit Function

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Public Function Valida_Dados_Padrao() As Long

Dim lErro As Long
Dim sValor As String, iLinha As Integer, iLinhaAux As Integer, sValorAux As String
Dim objTabelaCampo As ClassImportTabelasCampos, lCampoReduzTam As Long

On Error GoTo Erro_Valida_Dados_Padrao

    'Verifica se todos campos obrigatórios estão preenchidos e se os campos chave não repetem
    'Valida tabém o tipo de dado do campo e limita o tamanho máximo de strings
    For Each objTabelaCampo In gobjTabela.colCampos
    
        If objTabelaCampo.iObrigatorio = MARCADO Then
            For iLinha = 1 To objGridDados.iLinhasExistentes
                sValor = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
                If Len(Trim(sValor)) = 0 Then gError 208936
            Next
        End If
        
        If objTabelaCampo.iChave = MARCADO Then
            For iLinha = 1 To objGridDados.iLinhasExistentes
                sValor = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
                For iLinhaAux = 1 To objGridDados.iLinhasExistentes
                    If iLinha <> iLinhaAux Then
                        sValorAux = GridDados.TextMatrix(iLinhaAux, objTabelaCampo.lCodigo)
                        If UCase(sValor) = UCase(sValorAux) Then gError 208937
                    End If
                Next
            Next
        End If
        
        For iLinha = 1 To objGridDados.iLinhasExistentes
            sValor = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
        
            If Len(Trim(sValor)) > 0 Then
                Select Case objTabelaCampo.iTipo
                
                    Case TIPO_TEXTO
                        If objTabelaCampo.iTamMax > 0 And objTabelaCampo.iTamMax < Len(sValor) Then
                            If lCampoReduzTam <> objTabelaCampo.lCodigo Then Call Rotina_Aviso(vbOKOnly, "ERRO_CAMPO_ACIMA_TAM_PERMITIDO", objTabelaCampo.sCampo, objTabelaCampo.iTamMax)
                            lCampoReduzTam = objTabelaCampo.lCodigo
                            GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = left(sValor, objTabelaCampo.iTamMax)
                        End If
                
                    Case TIPO_DATA
                        If Not IsDate(sValor) And Trim(sValor) <> STRING_APAGA Then gError 208938
        
                    Case TIPO_DECIMAL
                        If Not IsNumeric(sValor) Then gError 208938
        
                    Case TIPO_NUMERICO
                        If Not IsNumeric(sValor) Then gError 208938
        
                End Select
            End If
            
        Next

    Next
    
    Valida_Dados_Padrao = SUCESSO

    Exit Function

Erro_Valida_Dados_Padrao:

    Valida_Dados_Padrao = gErr

    Select Case gErr
    
        Case 208936
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPO_OBRIGATORIO_NAO_PREENCHIDO", gErr, objTabelaCampo.sCampo, iLinha)

        Case 208937
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPO_CHAVE_REPETIDO", gErr, objTabelaCampo.sCampo, iLinha, iLinhaAux)

        Case 208938
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPO_FORMATO_ERRADO", gErr, objTabelaCampo.sCampo, iLinha, sValor)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208939)

    End Select
    
    Status.Caption = "A importação não poderá ser realizada."

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa Tela
    Call Limpa_Tela_ImportarDados

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208940)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim sSQL As String

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Status.Caption)) > 0 Then gError 208941
    'If Len(Trim(gobjTabela.sFuncaoGrava)) = 0 Then gError 208942

    lErro = Move_Tela_Memoria(sSQL)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("ImportTabelas_Grava", gobjTabela, sSQL)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 208941
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTTABELAS_DADOS_COM_ERRO", gErr)

        Case 208942
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTTABELAS_SEM_FUNC_GRAVA", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208943)

    End Select

    Exit Function

End Function

Function Produto_Ajusta_Formato(sProduto As String, Optional bComMask As Boolean) As Long

Dim lErro As Long
Dim sProdutoNovo As String
Dim iSeg As Integer, iNumSeg As Integer
Dim objSegmento As New ClassSegmento, sSeg As String
Dim colSegmento As New Collection
Dim iPos As Integer, iTamFalta As Integer
Dim sProdSeg As String
Dim objProd As ClassProduto, iTeste As Integer, bAchouProd As Boolean
Dim objProdAux As ClassProduto, sProdutoBD As String, iProdPreenchido As Integer

On Error GoTo Erro_Produto_Ajusta_Formato

    objSegmento.sCodigo = "produto"

    'preenche toda colecao(colSegmento) em relacao ao formato corrente
    lErro = CF("Segmento_Le_Codigo", objSegmento, colSegmento)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sProdSeg = ""
    For Each objSegmento In colSegmento
        If objSegmento.iTipo = SEGMENTO_NUMERICO Then
            Select Case objSegmento.iPreenchimento
                Case ZEROS_ESPACOS
                    sProdSeg = sProdSeg & String(objSegmento.iTamanho, "0")
                Case ESPACOS
                    sProdSeg = sProdSeg & String(objSegmento.iTamanho, " ")
            End Select
        Else
            sProdSeg = sProdSeg & String(objSegmento.iTamanho, " ")
        End If
    Next
    
    Set objSegmento = colSegmento.Item(1)
    
    iPos = InStr(1, sProduto, objSegmento.sDelimitador)
        
    'Se tiver ponto e a quantidade de pontos corresponde a quantidade de segmentos - 1
    If iPos <> 0 And (colSegmento.Count - 1) = (Len(sProduto) - Len(Replace(sProduto, ".", ""))) Then
    
        bComMask = True
        
        sSeg = Mid(sProduto, 1, iPos - 1)
        
        iTamFalta = objSegmento.iTamanho - Len(sSeg)
        
        If iTamFalta = 0 Or objSegmento.iPreenchimento = PREENCH_LIMPA_BRANCOS Then
            sProdutoNovo = sProduto
        Else
            If objSegmento.iTipo = SEGMENTO_NUMERICO Then
                Select Case objSegmento.iPreenchimento
                    Case ZEROS_ESPACOS
                        sSeg = String(iTamFalta, "0") & sSeg
                    Case ESPACOS
                        sSeg = String(iTamFalta, " ") & sSeg
                End Select
            Else
                sSeg = sSeg & String(iTamFalta, " ")
            End If
            sProdutoNovo = sSeg & Mid(sProduto, iPos)
        End If
    Else
        bComMask = False
        iTeste = 0
        bAchouProd = False
        Do While Not bAchouProd
        
            Set objProd = New ClassProduto
            Set objProdAux = New ClassProduto
            
            iTeste = iTeste + 1
            iTamFalta = Len(sProdSeg) - Len(sProduto)
            
            Select Case iTeste
                Case 1
                    objProd.sCodigo = sProduto
                Case 2
                    objProd.sCodigo = String(iTamFalta, " ") & sProduto
                Case 3
                    objProd.sCodigo = String(iTamFalta, "0") & sProduto
                Case 4
                    objProd.sCodigo = sProduto & String(iTamFalta, " ")
                Case 5
                    objProd.sCodigo = sProduto & String(iTamFalta, " ")
                Case 6
                    lErro = CF("Produto_Formata", sProduto, sProdutoBD, iProdPreenchido)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    objProd.sCodigo = sProdutoBD
                Case 7
                    objProd.sCodigo = sProduto
                    Exit Do
            End Select
            
            lErro = CF("Produto_Le", objProd)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            If lErro = SUCESSO Then
            
                objProdAux.sNomeReduzido = objProd.sNomeReduzido
            
                lErro = CF("Produto_Le_NomeReduzido", objProdAux)
                If lErro <> SUCESSO And lErro <> 26927 Then gError ERRO_SEM_MENSAGEM
                
                objProd.sCodigo = objProdAux.sCodigo
            
                If lErro = SUCESSO Then bAchouProd = True
            End If
            
        Loop
        
        sProdutoNovo = objProd.sCodigo
    End If
    
    sProduto = sProdutoNovo

    Produto_Ajusta_Formato = SUCESSO

    Exit Function

Erro_Produto_Ajusta_Formato:

    Produto_Ajusta_Formato = gErr

    Select Case gErr
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208961)

    End Select

    Exit Function

End Function

'===========================================================
'PLANO DE CONTAS
'===========================================================
Public Function Valida_Dados_PlanoContas() As Long

Dim lErro As Long, sContaPos As String, iIndice As Integer
Dim sConta As String, sContaBD As String, sContaTela As String
Dim iCategoria As Integer, iTipo As Integer, iNatureza As Integer
Dim objTabelaCampo As ClassImportTabelasCampos, sContaPai As String
Dim iLinha As Integer, colcolColecoes As New Collection
Dim lColunaConta As Long, colColunas As New Collection, lColunaDesc As Long
Dim iContaPreenchida As Integer, sDescricao As String, iNivel As Integer
Dim lColunaTipo As Long, lColunaCat As Long, lColunaNat As Long, lColunaUsaSimples As Long
Dim lColunaContaSimples As Long, sContaSimples As String
Dim sContaAux As String, bAchouPai As Boolean

On Error GoTo Erro_Valida_Dados_PlanoContas

    'pega as colunas de dados importantes e formata a conta
    For Each objTabelaCampo In gobjTabela.colCampos
        Select Case UCase(objTabelaCampo.sCampo)
            Case "CONTA"
                lColunaConta = objTabelaCampo.lCodigo
                For iLinha = 1 To objGridDados.iLinhasExistentes
                
                    sConta = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
                    
                    lErro = CF("Conta_Formata", sConta, sContaBD, iContaPreenchida)
                    If lErro <> SUCESSO Then gError 208933
        
                    lErro = Mascara_RetornaContaTela(sContaBD, sContaTela)
                    If lErro <> SUCESSO Then gError 208933
        
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = sContaTela
                Next
            Case "DESCCONTA"
                lColunaDesc = objTabelaCampo.lCodigo
            Case "TIPOCONTA"
                lColunaTipo = objTabelaCampo.lCodigo
            Case "NATUREZA"
                lColunaNat = objTabelaCampo.lCodigo
            Case "CATEGORIA"
                lColunaCat = objTabelaCampo.lCodigo
            Case "CONTASIMPLES"
                lColunaContaSimples = objTabelaCampo.lCodigo
            Case "USACONTASIMPLES"
                lColunaUsaSimples = objTabelaCampo.lCodigo
        End Select
    Next
    
    'Ordena pela conta
    colColunas.Add lColunaConta

    Call Ordena_Grid(objGridDados, colColunas, ORDEM_CRESCENTE, colcolColecoes)

    'Obtém dados como o tipo, natureza, categoria, etc
    For iLinha = 1 To objGridDados.iLinhasExistentes
        sConta = GridDados.TextMatrix(iLinha, lColunaConta)
        iTipo = CONTA_ANALITICA
        'Verifica se a próxima conta é um filho, se sim a conta é sintética
        If iLinha <> objGridDados.iLinhasExistentes Then
            
            sContaPos = GridDados.TextMatrix(iLinha + 1, lColunaConta)

            lErro = CF("Conta_Formata", sContaPos, sContaBD, iContaPreenchida)
            If lErro <> SUCESSO Then gError 208933

            lErro = Mascara_RetornaContaPai(sContaBD, sContaPai)
            If lErro <> SUCESSO Then gError 208933
            
            If sContaPai <> "" Then
                lErro = Mascara_RetornaContaTela(sContaPai, sContaTela)
                If lErro <> SUCESSO Then gError 208933
            End If
            
            If sContaTela = sConta And sContaPai <> "" Then
                iTipo = CONTA_SINTETICA
            End If
            
        End If
        
        'obtém o nível da conta atual
        lErro = CF("Conta_Formata", sConta, sContaBD, iContaPreenchida)
        If lErro <> SUCESSO Then gError 208933
               
        lErro = Mascara_Conta_ObterNivel(sContaBD, iNivel)
        If lErro <> SUCESSO Then gError 208933

        'Calcula a categoria e natureza
        If iNivel = 1 Then
            sDescricao = GridDados.TextMatrix(iLinha, lColunaDesc)
            Select Case UCase(left(sDescricao, 5))
            
                Case "ATIVO"
                    iCategoria = 1
                    iNatureza = 2
                
                Case "PASSI"
                    iCategoria = 2
                    iNatureza = 1
                
                Case "RECEI"
                    iCategoria = 3
                    iNatureza = 1
                
                Case "DESPE"
                    iCategoria = 4
                    iNatureza = 2
                
                Case "RESUL"
                    iCategoria = 5
                    iNatureza = 1
                
                Case Else
                
                    If left(sConta, 1) = "1" Then
                        iCategoria = 1
                        iNatureza = 2
                    ElseIf left(sConta, 1) = "2" Then
                        iCategoria = 2
                        iNatureza = 1
                    ElseIf left(sConta, 1) = "3" Then
                        iCategoria = 3
                        iNatureza = 1
                    ElseIf left(sConta, 1) = "4" Then
                        iCategoria = 4
                        iNatureza = 2
                    Else
                        iCategoria = 5
                        iNatureza = 1
                    End If
            
            End Select
            
        Else
            iCategoria = 0
                           
            'Testa a conta para ver se o pai está no Grid
            lErro = Mascara_RetornaContaPai(sContaBD, sContaPai)
            If lErro <> SUCESSO Then gError 208933
            
            bAchouPai = False
            For iIndice = iLinha To 1 Step -1
                sContaAux = GridDados.TextMatrix(iIndice, lColunaConta)
                lErro = CF("Conta_Formata", sContaAux, sContaBD, iContaPreenchida)
                If lErro <> SUCESSO Then gError 208933
            
                If sContaBD = sContaPai Then
                    bAchouPai = True
                    Exit For
                End If
            Next
            
            'A conta pai não cadastrada
            If Not bAchouPai Then
                gError 208934
            End If
                
        End If
        GridDados.TextMatrix(iLinha, lColunaNat) = CStr(iNatureza)
        GridDados.TextMatrix(iLinha, lColunaCat) = CStr(iCategoria)
        GridDados.TextMatrix(iLinha, lColunaTipo) = CStr(iTipo)
        
        sContaSimples = GridDados.TextMatrix(iLinha, lColunaContaSimples)
        
        If Len(Trim(sContaSimples)) > 0 Then
            GridDados.TextMatrix(iLinha, lColunaUsaSimples) = MARCADO
        Else
            GridDados.TextMatrix(iLinha, lColunaUsaSimples) = DESMARCADO
        End If
        
    Next

    Valida_Dados_PlanoContas = SUCESSO

    Exit Function

Erro_Valida_Dados_PlanoContas:

    Valida_Dados_PlanoContas = gErr

    Select Case gErr
    
        Case 208933
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_FORMATACAO", gErr, sConta)
        
        Case 208934
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_SEM_PAI", gErr, sConta)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208935)

    End Select
    
    Status.Caption = "A importação não poderá ser realizada."

    Exit Function

End Function

'===========================================================
'PLANO DE CONTAS
'===========================================================

'===========================================================
'ESTOQUE INICIAL
'===========================================================
Function Carrega_ComboAlmox() As Long

Dim lErro As Long
Dim objAlmoxarifado As ClassAlmoxarifado
Dim colAlmoxFilial As New Collection

On Error GoTo Erro_Carrega_ComboAlmox

    Almoxarifado.Clear

    'Lê todas as Grades de Produto
    lErro = CF("Almoxarifado_Le_Todos", colAlmoxFilial)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Adiciona as Grades lidas na List
    For Each objAlmoxarifado In colAlmoxFilial
        Almoxarifado.AddItem objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
    Next
    
    If Almoxarifado.ListCount = 1 Then
        Almoxarifado.ListIndex = 0
    End If

    Carrega_ComboAlmox = SUCESSO

    Exit Function

Erro_Carrega_ComboAlmox:

    Carrega_ComboAlmox = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208952)

    End Select

    Exit Function

End Function

Private Sub BotaoCriarAlmox_Click()
    Call Chama_Tela("Almoxarifado")
End Sub

Public Sub BotaoAtualizarAlmox_Click()

Dim sValorAnt As String

On Error GoTo Erro_BotaoAtualizarAlmox_Click

    sValorAnt = Almoxarifado.Text

    Call Carrega_ComboAlmox

    Call CF("SCombo_Seleciona2", Almoxarifado, sValorAnt)

    Exit Sub

Erro_BotaoAtualizarAlmox_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208953)

    End Select

    Exit Sub

End Sub

Public Function Valida_Dados_EstoqueInicial() As Long

Dim lErro As Long, iProdPreenchido As Integer
Dim objTabelaCampo As ClassImportTabelasCampos, iLinha As Integer
Dim iIndice As Integer, sProduto As String, sProdutoBD As String, sProdutoTela As String
Dim iAlmoxarifado As Integer, objAlmox As New ClassAlmoxarifado, sData As String
Dim objProduto As ClassProduto, bComMask As Boolean

On Error GoTo Erro_Valida_Dados_EstoqueInicial

    iAlmoxarifado = Codigo_Extrai(Almoxarifado.Text)
    If iAlmoxarifado = 0 Then gError 208954
    
    objAlmox.iCodigo = iAlmoxarifado
    
    lErro = CF("Almoxarifado_Le", objAlmox)
    If lErro <> SUCESSO And lErro <> 25056 Then gError ERRO_SEM_MENSAGEM

    'Formata o produto e vê se ele existe e pega o almoxarifado e filial empresa dele
    For Each objTabelaCampo In gobjTabela.colCampos
        Select Case UCase(objTabelaCampo.sCampo)
            Case "PRODUTO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                
                    Set objProduto = New ClassProduto
                
                    sProduto = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
                    
                    lErro = Produto_Ajusta_Formato(sProduto, bComMask)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    If bComMask Then
                        lErro = CF("Produto_Formata", sProduto, sProdutoBD, iProdPreenchido)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    Else
                        sProdutoBD = sProduto
                    End If
                    
                    objProduto.sCodigo = sProdutoBD
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
                    If lErro <> SUCESSO Then gError 208956
                    
                    If objProduto.iGerencial = MARCADO Then gError 208957
                    
                    If objProduto.iControleEstoque = PRODUTO_SEM_ESTOQUE Then gError 208958
        
                    lErro = Mascara_RetornaProdutoTela(sProdutoBD, sProdutoTela)
                    If lErro <> SUCESSO Then gError 208959
        
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = sProdutoTela
                Next
            Case "ALMOXARIFADO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = CStr(objAlmox.iCodigo)
                Next
            Case "FILIALEMPRESA"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = CStr(objAlmox.iFilialEmpresa)
                Next
            Case "DATAINICIAL"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    sData = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
                    If StrParaDate(sData) = DATA_NULA Then
                        GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = Format(Date, "dd/mm/yyyy")
                    End If
                Next
        End Select
    Next

    Valida_Dados_EstoqueInicial = SUCESSO

    Exit Function

Erro_Valida_Dados_EstoqueInicial:

    Valida_Dados_EstoqueInicial = gErr

    Select Case gErr
    
        Case 208954
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO1", gErr)
        
        Case 208955, 208959
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FORMATACAO", gErr, sProduto)
        
        Case 208956
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
    
        Case 208957
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
    
        Case 208958
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", gErr, objProduto.sCodigo)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208960)

    End Select
    
    Status.Caption = "A importação não poderá ser realizada."

    Exit Function

End Function
'===========================================================
'ESTOQUE INICIAL
'===========================================================

'===========================================================
'TABELA DE PREÇO
'===========================================================
Private Sub UpDownData_DownClick()
'Dimunui a data

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Diminui a data em 1 dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208963)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()
'aumenta a data

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Aumenta a data em 1 dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208964)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data)
End Sub

Private Sub Data_Validate(Cancel As Boolean)
'verifica se o campo Data está correto

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se o campo Data foi preenchida
    If Len(Data.ClipText) > 0 Then
        
        'Critica a Data
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Data_Validate:
    
    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208965)

    End Select

    Exit Sub
    
End Sub

Private Function Carrega_TabelaPreco() As Long
'Carrega a ComboBox Tabela

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigo As New Collection

On Error GoTo Erro_Carrega_TabelaPreco

    'Preenche a ComboBox com  os Tipos de Documentos existentes no BD
    lErro = CF("TabelasPreco_Le_Codigos", colCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For iIndice = 1 To colCodigo.Count
        'Preenche a ComboBox Tabela com os objetos da colecao colTabelaPreco
        Tabela.AddItem colCodigo(iIndice)
        Tabela.ItemData(Tabela.NewIndex) = colCodigo(iIndice)
    Next
    
    If Tabela.ListCount = 1 Then Tabela.ListIndex = 0

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208966)

    End Select

    Exit Function

End Function

Public Sub Tabela_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tabela_Click()

Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco

On Error GoTo Error_Tabela_Click

    'Verifica se foi preenchida a ComboBox Tabela
    If Tabela.ListIndex <> -1 Then

        objTabelaPreco.iCodigo = CInt(Tabela.Text)

        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError ERRO_SEM_MENSAGEM

        If lErro = 28004 Then gError 208967

        DescricaoTabela.Caption = objTabelaPreco.sDescricao

    End If

    iAlterado = 0

    Exit Sub

Error_Tabela_Click:

    Select Case gErr

        Case 208967
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", gErr, objTabelaPreco.iCodigo)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208968)

    End Select

    Exit Sub

End Sub

Private Sub BotaoCriarTabela_Click()
Dim objTabelaPreco As New ClassTabelaPreco
    Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
End Sub

Public Sub BotaoAtualizarTabela_Click()

Dim iValorAnt As Integer

On Error GoTo Erro_BotaoAtualizarTabela_Click

    iValorAnt = StrParaInt(Tabela.Text)

    Call Carrega_TabelaPreco

    Call Combo_Seleciona(Tabela, iValorAnt)

    Exit Sub

Erro_BotaoAtualizarTabela_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208975)

    End Select

    Exit Sub

End Sub

Public Function Valida_Dados_TabelaPrecoItem() As Long

Dim lErro As Long, iProdPreenchido As Integer
Dim objTabelaCampo As ClassImportTabelasCampos, iLinha As Integer
Dim iIndice As Integer, sProduto As String, sProdutoBD As String, sProdutoTela As String
Dim iTabela As Integer, sData As String, dtData As Date
Dim objProduto As ClassProduto, bComMask As Boolean

On Error GoTo Erro_Valida_Dados_TabelaPrecoItem

    iTabela = StrParaInt(Tabela.Text)
    If iTabela = 0 Then gError 208969
    
    dtData = StrParaDate(Data.Text)
    If dtData = DATA_NULA Then gError 208970

    'Formata o produto e vê se ele existe
    For Each objTabelaCampo In gobjTabela.colCampos
        Select Case UCase(objTabelaCampo.sCampo)
            Case "PRODUTO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                
                    Set objProduto = New ClassProduto
                
                    sProduto = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
                    
                    lErro = Produto_Ajusta_Formato(sProduto, bComMask)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    If bComMask Then
                        lErro = CF("Produto_Formata", sProduto, sProdutoBD, iProdPreenchido)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    Else
                        sProdutoBD = sProduto
                    End If
                    
                    objProduto.sCodigo = sProdutoBD
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
                    If lErro <> SUCESSO Then gError 208972
        
                    lErro = Mascara_RetornaProdutoTela(sProdutoBD, sProdutoTela)
                    If lErro <> SUCESSO Then gError 208973
        
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = sProdutoTela
                Next
            Case "CODTABELA"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = CStr(iTabela)
                Next
            Case "FILIALEMPRESA"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = CStr(giFilialEmpresa)
                Next
            Case "DATAVIGENCIA"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = Format(dtData, "dd/mm/yyyy")
                Next
        End Select
    Next

    Valida_Dados_TabelaPrecoItem = SUCESSO

    Exit Function

Erro_Valida_Dados_TabelaPrecoItem:

    Valida_Dados_TabelaPrecoItem = gErr

    Select Case gErr
    
        Case 208969
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
        
        Case 208970
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_NAO_PREENCHIDA", gErr)
        
        Case 208971, 208973
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FORMATACAO", gErr, sProduto)
        
        Case 208972
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208973)

    End Select
    
    Status.Caption = "A importação não poderá ser realizada."

    Exit Function

End Function
'===========================================================
'TABELA DE PREÇO
'===========================================================

'===========================================================
'CENTRO DE CUSTO E LUCRO
'===========================================================
Public Function Valida_Dados_Ccl() As Long

Dim lErro As Long, sCclPos As String, iIndice As Integer
Dim sCcl As String, sCclBD As String, sCclTela As String
Dim iTipo As Integer, lColunaTipo As Long
Dim objTabelaCampo As ClassImportTabelasCampos, sCclPai As String
Dim iLinha As Integer, colcolColecoes As New Collection
Dim lColunaCcl As Long, colColunas As New Collection
Dim iCclPreenchida As Integer, iNivel As Integer
Dim sCclAux As String, bAchouPai As Boolean

On Error GoTo Erro_Valida_Dados_Ccl

    'pega as colunas de dados importantes e formata a Ccl
    For Each objTabelaCampo In gobjTabela.colCampos
        Select Case UCase(objTabelaCampo.sCampo)
            Case "CCL"
                lColunaCcl = objTabelaCampo.lCodigo
                For iLinha = 1 To objGridDados.iLinhasExistentes
                
                    sCcl = GridDados.TextMatrix(iLinha, lColunaCcl)
                    
                    lErro = CF("Ccl_Formata", sCcl, sCclBD, iCclPreenchida)
                    If lErro <> SUCESSO Then gError 208976
        
                    sCclTela = String(STRING_CCL, 0)
        
                    lErro = Mascara_RetornaItemTela(SEGMENTO_CCL, sCclBD, sCclTela)
                    If lErro <> SUCESSO Then gError 208976
        
                    GridDados.TextMatrix(iLinha, lColunaCcl) = sCclTela
                Next
            Case "TIPOCCL"
                lColunaTipo = objTabelaCampo.lCodigo
        End Select
    Next
    
    'Ordena pela Ccl
    colColunas.Add lColunaCcl

    Call Ordena_Grid(objGridDados, colColunas, ORDEM_CRESCENTE, colcolColecoes)

    'Obtém dados como o tipo
    For iLinha = 1 To objGridDados.iLinhasExistentes
        sCcl = GridDados.TextMatrix(iLinha, lColunaCcl)
        iTipo = CCL_ANALITICA
        'Verifica se a próxima Ccl é um filho, se sim a Ccl é sintética
        If iLinha <> objGridDados.iLinhasExistentes Then
            
            sCclPos = GridDados.TextMatrix(iLinha + 1, lColunaCcl)

            lErro = CF("Ccl_Formata", sCclPos, sCclBD, iCclPreenchida)
            If lErro <> SUCESSO Then gError 208976

            lErro = Mascara_RetornaCclPai(sCclBD, sCclPai)
            If lErro <> SUCESSO Then gError 208976
            
            If sCclPai <> "" Then
                sCclTela = String(STRING_CCL, 0)
                lErro = Mascara_RetornaItemTela(SEGMENTO_CCL, sCclPai, sCclTela)
                If lErro <> SUCESSO Then gError 208976
            End If
            
            If sCclTela = sCcl And sCclPai <> "" Then
                iTipo = CCL_SINTETICA
            End If
            
        End If
        
        'obtém o nível da Ccl atual
        lErro = CF("Ccl_Formata", sCcl, sCclBD, iCclPreenchida)
        If lErro <> SUCESSO Then gError 208976
               
        lErro = Mascara_Ccl_ObterNivel(sCclBD, iNivel)
        If lErro <> SUCESSO Then gError 208976

        If iNivel <> 1 Then
                           
            'Testa a Ccl para ver se o pai está no Grid
            lErro = Mascara_RetornaCclPai(sCclBD, sCclPai)
            If lErro <> SUCESSO Then gError 208976
            
            bAchouPai = False
            For iIndice = iLinha To 1 Step -1
                sCclAux = GridDados.TextMatrix(iIndice, lColunaCcl)
                lErro = CF("Ccl_Formata", sCclAux, sCclBD, iCclPreenchida)
                If lErro <> SUCESSO Then gError 208976
            
                If sCclBD = sCclPai Then
                    bAchouPai = True
                    Exit For
                End If
            Next
            
            'A Ccl pai não cadastrada
            If Not bAchouPai Then
                gError 208977
            End If
                
        End If
        GridDados.TextMatrix(iLinha, lColunaTipo) = CStr(iTipo)
                
    Next

    Valida_Dados_Ccl = SUCESSO

    Exit Function

Erro_Valida_Dados_Ccl:

    Valida_Dados_Ccl = gErr

    Select Case gErr
    
        Case 208976
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_FORMATACAO", gErr, sCcl)
        
        Case 208977
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_SEM_PAI", gErr, sCcl)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208978)

    End Select
    
    Status.Caption = "A importação não poderá ser realizada."

    Exit Function

End Function

'===========================================================
'CENTRO DE CUSTO E LUCRO
'===========================================================

'===========================================================
'G110
'===========================================================
Public Function Valida_Dados_SpedFis() As Long

Dim lErro As Long
Dim objTabelaCampo As ClassImportTabelasCampos, iLinha As Integer
Dim iIndice As Integer, iMes As Integer, iAno As Integer
Dim objProduto As ClassProduto
Dim sProduto As String, sProdutoBD As String, sProdutoTela As String, iProdPreenchido As Integer
Dim lColTipo As Long, lColCodigo As Long, iTipo As Integer, objCli As ClassCliente, objForn As ClassFornecedor
Dim objFilCli As ClassFilialCliente, objFilForn As ClassFilialFornecedor, iFilial As Integer, lCodigo As Long
Dim sCcl As String, sCclBD As String, iCclPreenchida As Integer, objCcl As ClassCcl
Dim sConta As String, sContaBD As String, iContaPreenchida As Integer, bComMask As Boolean

On Error GoTo Erro_Valida_Dados_SpedFis

    iMes = Codigo_Extrai(Mes.Text)
    If iMes = 0 Then gError 209320
    
    iAno = StrParaInt(Ano.Text)
    If iAno = 0 Then gError 209321
    
    'Não vai excluir na gravação para manter o histórico por ano\mês\filial
    'Então tem que excluir aqui para sobrepor os dados
    lErro = CF("SpedFis_Tabela_Exclui", gobjTabela.sTabela, iAno, iMes, giFilialEmpresa)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Formata o produto e vê se ele existe e pega o almoxarifado e filial empresa dele
    For Each objTabelaCampo In gobjTabela.colCampos
        Select Case UCase(objTabelaCampo.sCampo)
            Case "MES"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = CStr(iMes)
                Next
            Case "ANO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = CStr(iAno)
                Next
            Case "SEQ", "CODIGO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    If StrParaLong(GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)) = 0 Then GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = CStr(iLinha)
                Next
            Case "FILIALEMPRESA"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = CStr(giFilialEmpresa)
                Next
            Case "PRODUTO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                
                    Set objProduto = New ClassProduto
                
                    sProduto = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
                    
                    lErro = Produto_Ajusta_Formato(sProduto, bComMask)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    If bComMask Then
                        lErro = CF("Produto_Formata", sProduto, sProdutoBD, iProdPreenchido)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    Else
                        sProdutoBD = sProduto
                    End If
                    
                    objProduto.sCodigo = sProdutoBD
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
                    If lErro <> SUCESSO Then gError 209322
                    
                    If objProduto.iGerencial = MARCADO Then gError 209323
                    
                    lErro = Mascara_RetornaProdutoTela(sProdutoBD, sProdutoTela)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = sProdutoBD 'sProdutoTela
                Next
            Case "TIPOMOVTO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    Select Case GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
                        Case "SI", "IM", "IA", "CI", "MC", "BA", "AT", "PE", "OT"
                        Case Else
                            gError 209327
                     End Select
                Next
            Case "CODIGOBEM"
            
            Case "EMITENTE"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    Select Case StrParaInt(GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo))
                        Case "0", "1"
                        Case Else
                            gError 209328 '0- Emissão própria; 1- Terceiros
                     End Select
                Next
            Case "TIPOCLIFORN"
                lColTipo = objTabelaCampo.lCodigo
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    Select Case StrParaInt(GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo))
                        Case "1", "2"
                        Case Else
                            gError 209329 '1- Cliente; 2- Fornecedor
                     End Select
                Next
            Case "CLIFORN"
                lColCodigo = objTabelaCampo.lCodigo
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    iTipo = StrParaInt(GridDados.TextMatrix(iLinha, lColTipo))
                    lCodigo = StrParaLong(GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo))
                    If iTipo = 1 Then
                        Set objCli = New ClassCliente
                        objCli.lCodigo = lCodigo
                        lErro = CF("Cliente_Le", objCli)
                        If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
                        If lErro <> SUCESSO Then gError 209330
                    Else
                        Set objForn = New ClassFornecedor
                        objForn.lCodigo = lCodigo
                        lErro = CF("Fornecedor_Le", objForn)
                        If lErro <> SUCESSO And lErro <> 12729 Then gError ERRO_SEM_MENSAGEM
                        If lErro <> SUCESSO Then gError 209331
                    End If
                Next
            Case "FILIAL"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    iTipo = StrParaInt(GridDados.TextMatrix(iLinha, lColTipo))
                    lCodigo = StrParaInt(GridDados.TextMatrix(iLinha, lColCodigo))
                    iFilial = StrParaInt(GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo))
                    If iTipo = 1 Then
                        Set objFilCli = New ClassFilialCliente
                        objFilCli.lCodCliente = lCodigo
                        objFilCli.iCodFilial = iFilial
                        lErro = CF("FilialCliente_Le", objFilCli)
                        If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
                        If lErro <> SUCESSO Then gError 209332
                    Else
                        Set objFilForn = New ClassFilialFornecedor
                        objFilForn.lCodFornecedor = lCodigo
                        objFilForn.iCodFilial = iFilial
                        lErro = CF("FilialFornecedor_Le", objFilForn)
                        If lErro <> SUCESSO And lErro <> 12929 Then gError ERRO_SEM_MENSAGEM
                        If lErro <> SUCESSO Then gError 209333
                    End If
                Next
            Case "TIPO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    Select Case GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
                        Case "1", "2"
                        Case Else
                            gError 209336
                     End Select
                Next
            Case "CCL"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    sCcl = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
                    Select Case sCcl
                        Case "1", "2", "3", "4", "5"
                        Case Else
                            lErro = CF("Ccl_Formata", sCcl, sCclBD, iCclPreenchida)
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                            Set objCcl = New ClassCcl
                            objCcl.sCcl = sCclBD
                            lErro = CF("Ccl_Le", objCcl)
                            If lErro <> SUCESSO And lErro <> 5599 Then gError ERRO_SEM_MENSAGEM
                            If lErro <> SUCESSO Then gError 209337
                        
                            GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = sCclBD
                    End Select
                Next
            Case "CONTACONTABIL"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    
                    sConta = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)

                    lErro = CF("Conta_Formata", sConta, sContaBD, iContaPreenchida)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                    lErro = CF("PlanoConta_Le_Conta", sContaBD)
                    If lErro <> SUCESSO And lErro <> 10051 Then gError ERRO_SEM_MENSAGEM
                    If lErro <> SUCESSO Then gError 209338
                
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = sContaBD
                    
                Next
        End Select
    Next
    
    Valida_Dados_SpedFis = SUCESSO

    Exit Function

Erro_Valida_Dados_SpedFis:

    Valida_Dados_SpedFis = gErr

    Select Case gErr
    
        Case 209320
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREENCHIDO", gErr)
        
        Case 209321
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
        
        Case 209322
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
    
        Case 209323
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            
        Case 209327
            Call Rotina_Erro(vbOKOnly, "ERRO_SPED_TIPOMOVTO_INVALIDO", gErr, GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo), iLinha)
        
        Case 209328
            Call Rotina_Erro(vbOKOnly, "ERRO_SPED_EMITENTE_INVALIDO", gErr, GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo), iLinha)
        
        Case 209329
            Call Rotina_Erro(vbOKOnly, "ERRO_SPED_TIPOCLIFORN_INVALIDO", gErr, GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo), iLinha)
        
        Case 209330
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_LINHA_NAO_CADASTRADO", gErr, lCodigo, iLinha)
        
        Case 209331
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_LINHA_NAO_CADASTRADO", gErr, lCodigo, iLinha)
        
        Case 209332
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_LINHA_NAO_CADASTRADO", gErr, lCodigo, iFilial, iLinha)
        
        Case 209333
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORN_LINHA_NAO_CADASTRADO", gErr, lCodigo, iFilial, iLinha)
        
        Case 209336
            Call Rotina_Erro(vbOKOnly, "ERRO_SPED_TIPO_INVALIDO", gErr, GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo), iLinha)
        
        Case 209337
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_LINHA_NAO_CADASTRADO", gErr, sCclBD, iLinha)
        
        Case 209338
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_LINHA_NAO_CADASTRADO", gErr, sContaBD, iLinha)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209324)

    End Select
    
    Status.Caption = "A importação não poderá ser realizada."

    Exit Function

End Function

Private Sub Ano_GotFocus()
    Call MaskEdBox_TrataGotFocus(Ano, iAlterado)
End Sub

Private Sub Ano_Validate(bCancel As Boolean)

On Error GoTo Erro_Ano_Validate

    If Len(Trim(Ano.Text)) > 0 Then

        If StrParaInt(Ano.Text) < 2000 Or StrParaInt(Ano.Text) > Year(Date) Then gError 209325
        
    End If
    
    Exit Sub
    
Erro_Ano_Validate:

    Select Case gErr
    
        Case 209325
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_INVALIDO2", gErr)
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209326)

    End Select
    
End Sub
'===========================================================
'G110
'===========================================================

'===========================================================
'PREVVENDA
'===========================================================
Private Sub AnoPrevVenda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AnoPrevVenda_GotFocus()
    Call MaskEdBox_TrataGotFocus(AnoPrevVenda, iAlterado)
End Sub

Private Sub AnoPrevVenda_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AnoPrevVenda_Validate

    If Len(Trim(AnoPrevVenda.ClipText)) = 0 Then Exit Sub
    
    If StrParaInt(AnoPrevVenda.ClipText) < 2000 Then gError 91075

    Exit Sub

Erro_AnoPrevVenda_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 91075
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_MENOR_2000", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165137)

    End Select

    Exit Sub

End Sub

Public Function Valida_Dados_PrevVenda() As Long

Dim lErro As Long, iProdPreenchido As Integer, bComMask As Boolean
Dim objTabelaCampo As ClassImportTabelasCampos, iLinha As Integer
Dim iIndice As Integer, sProduto As String, sProdutoBD As String, sProdutoTela As String
Dim objProduto As ClassProduto, sCodPrevVenda As String, iAno As Integer
Dim lColCliente As Long, objCli As ClassCliente
Dim objFilCli As ClassFilialCliente

On Error GoTo Erro_Valida_Dados_PrevVenda

    sCodPrevVenda = Trim(CodPrevVenda.Text)
    If Len(sCodPrevVenda) = 0 Then gError 213701
    
    iAno = StrParaInt(AnoPrevVenda.Text)
    If iAno < Year(Now) Then gError 213702

    'Formata o produto e vê se ele existe e pega o almoxarifado e filial empresa dele
    For Each objTabelaCampo In gobjTabela.colCampos
        Select Case UCase(objTabelaCampo.sCampo)
            Case "PRODUTO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                
                    Set objProduto = New ClassProduto
                
                    sProduto = GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo)
                    
                    lErro = Produto_Ajusta_Formato(sProduto, bComMask)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    If bComMask Then
                        lErro = CF("Produto_Formata", sProduto, sProdutoBD, iProdPreenchido)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    Else
                        sProdutoBD = sProduto
                    End If
                    
                    objProduto.sCodigo = sProdutoBD
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
                    If lErro <> SUCESSO Then gError 213703
                    
                    If objProduto.iGerencial = MARCADO Then gError 213704
                    
                    If objProduto.iControleEstoque = PRODUTO_SEM_ESTOQUE Then gError 213705
        
                    lErro = Mascara_RetornaProdutoTela(sProdutoBD, sProdutoTela)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = sProdutoTela
                Next
            Case "CODIGO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = sCodPrevVenda
                Next
            Case "ANO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = CStr(iAno)
                Next
            Case "FILIALEMPRESA"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = CStr(giFilialEmpresa)
                Next
            Case "DATAATUALIZACAO"
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo) = Format(Date, "dd/mm/yyyy")
                Next
        
            Case "CLIENTE" 'Verificar se o cliente existe
                lColCliente = objTabelaCampo.lCodigo
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    Set objCli = New ClassCliente
                    objCli.lCodigo = StrParaLong(GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo))
                    lErro = CF("Cliente_Le", objCli)
                    If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
                    If lErro <> SUCESSO Then gError 213706
                Next
                
            Case "FILIAL" ' Verificar se a filial cliente existe
                For iLinha = 1 To objGridDados.iLinhasExistentes
                    Set objFilCli = New ClassFilialCliente
                    objFilCli.lCodCliente = StrParaLong(GridDados.TextMatrix(iLinha, lColCliente))
                    objFilCli.iCodFilial = StrParaInt(GridDados.TextMatrix(iLinha, objTabelaCampo.lCodigo))
                    lErro = CF("FilialCliente_Le", objFilCli)
                    If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
                    If lErro <> SUCESSO Then gError 213707
                Next
        End Select
    Next

    Valida_Dados_PrevVenda = SUCESSO

    Exit Function

Erro_Valida_Dados_PrevVenda:

    Valida_Dados_PrevVenda = gErr

    Select Case gErr
           
        Case 213701
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PREVVENDAMENSAL_NAO_PREENCHIDO", gErr)

        Case 213702
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
            
        Case 213703
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
    
        Case 213704
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
    
        Case 213705
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", gErr, objProduto.sCodigo)
        
        Case 213706
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_LINHA_NAO_CADASTRADO", gErr, objCli.lCodigo, iLinha)
        
        Case 213707
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_LINHA_NAO_CADASTRADO", gErr, objFilCli.lCodCliente, objFilCli.iCodFilial, iLinha)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213706)

    End Select
    
    Status.Caption = "A importação não poderá ser realizada."

    Exit Function

End Function
'===========================================================
'PREVVENDA
'===========================================================

