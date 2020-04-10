VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpNFRequisicoesOcx 
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   ScaleHeight     =   4155
   ScaleWidth      =   8130
   Begin VB.Frame Frame1 
      Caption         =   "Notas Fiscais"
      Height          =   1800
      Left            =   240
      TabIndex        =   25
      Top             =   2265
      Width           =   7755
      Begin VB.Frame Frame4 
         Caption         =   "Fornecedores"
         Height          =   705
         Left            =   4290
         TabIndex        =   37
         Top             =   930
         Width           =   3270
         Begin MSMask.MaskEdBox FornecedorDe 
            Height          =   300
            Left            =   555
            TabIndex        =   12
            Top             =   240
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecedorAte 
            Height          =   300
            Left            =   2130
            TabIndex        =   13
            Top             =   255
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelFornecedorAte 
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
            Left            =   1725
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   39
            Top             =   315
            Width           =   360
         End
         Begin VB.Label LabelFornecedorDe 
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   38
            Top             =   315
            Width           =   315
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data de Entrada"
         Height          =   690
         Left            =   180
         TabIndex        =   32
         Top             =   945
         Width           =   3990
         Begin MSComCtl2.UpDown UpDownDataDe 
            Height          =   315
            Left            =   1665
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataDe 
            Height          =   315
            Left            =   480
            TabIndex        =   10
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataAte 
            Height          =   315
            Left            =   3630
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAte 
            Height          =   315
            Left            =   2445
            TabIndex        =   11
            Top             =   225
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
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
            Left            =   2070
            TabIndex        =   36
            Top             =   315
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
            Left            =   165
            TabIndex        =   35
            Top             =   315
            Width           =   315
         End
      End
      Begin VB.Frame FrameCodigo 
         Caption         =   "Número"
         Height          =   675
         Left            =   4305
         TabIndex        =   29
         Top             =   225
         Width           =   3270
         Begin MSMask.MaskEdBox NumeroDe 
            Height          =   300
            Left            =   540
            TabIndex        =   8
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumeroAte 
            Height          =   300
            Left            =   2115
            TabIndex        =   9
            Top             =   240
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelNumeroAte 
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
            Left            =   1695
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   31
            Top             =   300
            Width           =   360
         End
         Begin VB.Label LabelNumeroDe 
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   30
            Top             =   315
            Width           =   315
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Séries"
         Height          =   690
         Left            =   165
         TabIndex        =   26
         Top             =   225
         Width           =   3990
         Begin VB.ComboBox SerieDe 
            Height          =   315
            Left            =   495
            TabIndex        =   6
            Top             =   240
            Width           =   885
         End
         Begin VB.ComboBox SerieAte 
            Height          =   315
            Left            =   2460
            TabIndex        =   7
            Top             =   225
            Width           =   885
         End
         Begin VB.Label LabelSerieDe 
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
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   270
            Width           =   315
         End
         Begin VB.Label LabelSerieAte 
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
            Left            =   2085
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   27
            Top             =   270
            Width           =   360
         End
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpNFRequisicoesOcx.ctx":0000
      Left            =   1530
      List            =   "RelOpNFRequisicoesOcx.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   510
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filial Empresa"
      Height          =   1110
      Left            =   255
      TabIndex        =   20
      Top             =   1005
      Width           =   7725
      Begin MSMask.MaskEdBox CodFilialDe 
         Height          =   300
         Left            =   1200
         TabIndex        =   2
         Top             =   285
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodFilialAte 
         Height          =   300
         Left            =   4890
         TabIndex        =   3
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeAte 
         Height          =   300
         Left            =   4875
         TabIndex        =   5
         Top             =   660
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeDe 
         Height          =   300
         Left            =   1185
         TabIndex        =   4
         Top             =   690
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelCodFilialDe 
         AutoSize        =   -1  'True
         Caption         =   "Código De:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   330
         Width           =   960
      End
      Begin VB.Label LabelNomeDe 
         AutoSize        =   -1  'True
         Caption         =   "Nome De:"
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
         Left            =   315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   750
         Width           =   855
      End
      Begin VB.Label LabelCodFilialAte 
         AutoSize        =   -1  'True
         Caption         =   "Código Até:"
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
         Left            =   3870
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label LabelNomeAte 
         AutoSize        =   -1  'True
         Caption         =   "Nome Até:"
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
         Left            =   3960
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   720
         Width           =   900
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpNFRequisicoesOcx.ctx":003D
      Left            =   930
      List            =   "RelOpNFRequisicoesOcx.ctx":003F
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
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
      Left            =   3945
      Picture         =   "RelOpNFRequisicoesOcx.ctx":0041
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   135
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5820
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpNFRequisicoesOcx.ctx":0143
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpNFRequisicoesOcx.ctx":02C1
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpNFRequisicoesOcx.ctx":07F3
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpNFRequisicoesOcx.ctx":097D
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Ordenados Por:"
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
      Left            =   195
      TabIndex        =   41
      Top             =   585
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   195
      TabIndex        =   40
      Top             =   135
      Width           =   615
   End
End
Attribute VB_Name = "RelOpNFRequisicoesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpCompradoresPC
Const ORD_POR_NF = 0
Const ORD_POR_DATA = 1
Const ORD_POR_FORNECEDOR = 2


Private WithEvents objEventoCodPCDe As AdmEvento
Attribute objEventoCodPCDe.VB_VarHelpID = -1
Private WithEvents objEventoCodPCAte As AdmEvento
Attribute objEventoCodPCAte.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialDe As AdmEvento
Attribute objEventoCodFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialAte As AdmEvento
Attribute objEventoCodFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNumNFDe As AdmEvento
Attribute objEventoNumNFDe.VB_VarHelpID = -1
Private WithEvents objEventoNumNFAte As AdmEvento
Attribute objEventoNumNFAte.VB_VarHelpID = -1
Private WithEvents objEventoFornDe As AdmEvento
Attribute objEventoFornDe.VB_VarHelpID = -1
Private WithEvents objEventoFornAte As AdmEvento
Attribute objEventoFornAte.VB_VarHelpID = -1
Private WithEvents objEventoSerieDe As AdmEvento
Attribute objEventoSerieDe.VB_VarHelpID = -1
Private WithEvents objEventoSerieAte As AdmEvento
Attribute objEventoSerieAte.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 73876

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 73877

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 73876

        Case 73877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170280)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 73881

    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus

    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 73881

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170281)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodPCDe = New AdmEvento
    Set objEventoCodPCAte = New AdmEvento
    Set objEventoCodFilialDe = New AdmEvento
    Set objEventoCodFilialAte = New AdmEvento
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento
    Set objEventoFornDe = New AdmEvento
    Set objEventoFornAte = New AdmEvento
    Set objEventoSerieDe = New AdmEvento
    Set objEventoSerieAte = New AdmEvento
    Set objEventoNumNFDe = New AdmEvento
    Set objEventoNumNFAte = New AdmEvento
    
    ComboOrdenacao.ListIndex = 0

    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 73882
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 73882
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170282)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoCodPCDe = Nothing
    Set objEventoCodPCAte = Nothing
    Set objEventoCodFilialDe = Nothing
    Set objEventoCodFilialAte = Nothing
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing
    Set objEventoNumNFDe = Nothing
    Set objEventoNumNFAte = Nothing
    Set objEventoFornDe = Nothing
    Set objEventoFornAte = Nothing
    Set objEventoSerieDe = Nothing
    Set objEventoSerieAte = Nothing
    
End Sub

Private Sub CodFilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFilialAte, iAlterado)
    
End Sub

Private Sub CodFilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFilialDe, iAlterado)
    
End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    
End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    
End Sub

Private Sub FornecedorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorAte, iAlterado)
    
End Sub

Private Sub FornecedorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorDe, iAlterado)
    
End Sub

Private Sub LabelSerieDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objSerie As New ClassSerie

On Error GoTo Erro_LabelSerieDe_Click

    If Len(Trim(SerieDe.Text)) > 0 Then
    
        objSerie.sSerie = SerieDe.Text
    End If
    
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieDe)
    
    Exit Sub

Erro_LabelSerieDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170283)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelSerieAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objSerie As New ClassSerie

On Error GoTo Erro_LabelSerieAte_Click

    If Len(Trim(SerieAte.Text)) > 0 Then
    
        objSerie.sSerie = SerieAte.Text
    End If
    
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieAte)
    
    Exit Sub

Erro_LabelSerieAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170284)

    End Select

    Exit Sub
    
End Sub


Private Sub LabelFornecedorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorDe_Click

    If Len(Trim(FornecedorDe.Text)) > 0 Then
    
        objFornecedor.lCodigo = StrParaLong(FornecedorDe.Text)
        
    End If
    
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornDe)
    
    Exit Sub

Erro_LabelFornecedorDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170285)

    End Select

    Exit Sub
    
End Sub
Private Sub LabelFornecedorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorAte_Click

    If Len(Trim(FornecedorAte.Text)) > 0 Then
    
        objFornecedor.lCodigo = StrParaLong(FornecedorAte.Text)
        
    End If
    
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornAte)
    
    Exit Sub

Erro_LabelFornecedorAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170286)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelNumeroAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objNF As New ClassNFiscal

On Error GoTo Erro_LabelNumeroAte_Click

    If Len(Trim(NumeroAte.Text)) > 0 Then
        'Preenche com o numero da tela
        objNF.lNumNotaFiscal = StrParaLong(NumeroAte.Text)
    End If

    'Chama Tela NFiscalEntradaTodasLista
    Call Chama_Tela("NFiscalEntradaTodasLista", colSelecao, objNF, objEventoNumNFAte)

   Exit Sub

Erro_LabelNumeroAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170287)

    End Select

    Exit Sub

End Sub
Private Sub LabelNumeroDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objNF As New ClassNFiscal

On Error GoTo Erro_LabelNumeroDe_Click

    If Len(Trim(NumeroDe.Text)) > 0 Then
        'Preenche com o numero da tela
        objNF.lNumNotaFiscal = StrParaLong(NumeroDe.Text)
    End If

    'Chama Tela NFiscalEntradaTodasLista
    Call Chama_Tela("NFiscalEntradaTodasLista", colSelecao, objNF, objEventoNumNFDe)

   Exit Sub

Erro_LabelNumeroDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170288)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 73887

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73887
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170289)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 73888

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 73888
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170290)

    End Select

    Exit Sub

End Sub


Private Sub NumeroAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumeroAte, iAlterado)
    
End Sub

Private Sub NumeroDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumeroDe, iAlterado)
    
End Sub

Private Sub SerieDe_Validate(Cancel As Boolean)

Dim objSerie As New ClassSerie
Dim lErro As Long

On Error GoTo Erro_SerieDe_Validate

    If Len(Trim(SerieDe.Text)) > 0 Then
        
        objSerie.sSerie = SerieDe.Text
        
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO Then gError 73889
        
    End If
    
    Exit Sub
    
Erro_SerieDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 73889
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170291)
            
    End Select
    
    Exit Sub
    
End Sub
Private Sub SerieAte_Validate(Cancel As Boolean)

Dim objSerie As New ClassSerie
Dim lErro As Long

On Error GoTo Erro_SerieAte_Validate

    If Len(Trim(SerieAte.Text)) > 0 Then
        
        objSerie.sSerie = SerieAte.Text
        
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO Then gError 73891
        
    End If
    
    Exit Sub
    
Erro_SerieAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 73891
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170292)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73892

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 73892
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170293)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73897

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 73897
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170294)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73898

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 73898
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170295)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73899

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 73899
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170296)

    End Select

    Exit Sub

End Sub


Private Sub LabelCodFilialDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodFilialDe_Click

    If Len(Trim(CodFilialDe.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodFilialDe)

   Exit Sub

Erro_LabelCodFilialDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170297)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodFilialAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodFilialAte_Click

    If Len(Trim(CodFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodFilialAte)

   Exit Sub

Erro_LabelCodFilialAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170298)

    End Select

    Exit Sub

End Sub




Private Sub LabelNomeDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeDe_Click

    If Len(Trim(NomeDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objFilialEmpresa.sNome = NomeDe.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialDe)

   Exit Sub

Erro_LabelNomeDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170299)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeAte_Click

    If Len(Trim(NomeAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.sNome = NomeAte.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialAte)

   Exit Sub

Erro_LabelNomeAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170300)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodFilialAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeDe.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeAte.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodFilialDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFornDe_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorDe.Text = CStr(objFornecedor.lCodigo)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoFornAte_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorAte.Text = CStr(objFornecedor.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoSerieAte_evSelecao(obj1 As Object)

Dim objSerie As New ClassSerie

    Set objSerie = obj1

    SerieAte.Text = objSerie.sSerie

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoSerieDe_evSelecao(obj1 As Object)

Dim objSerie As New ClassSerie

    Set objSerie = obj1

    SerieDe.Text = objSerie.sSerie

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNumNFAte_evSelecao(obj1 As Object)

Dim objNF As New ClassNFiscal

    Set objNF = obj1

    NumeroAte.Text = CStr(objNF.lNumNotaFiscal)

    Me.Show

End Sub

Private Sub objEventoNumNFDe_evSelecao(obj1 As Object)

Dim objNF As New ClassNFiscal

    Set objNF = obj1

    NumeroDe.Text = CStr(objNF.lNumNotaFiscal)

    Me.Show

End Sub


Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 73850

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73851

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 73852

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 73853

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 73850
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 73851 To 73853

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170301)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 73854

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 73855

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 73854
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 73855

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170302)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73856

Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_NF
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "SerieNF", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "NumNF", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEntrada", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemNFiscal", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemNFItemRC", 1)

            Case ORD_POR_DATA

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEntrada", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "SerieNF", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "NumNF", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemNFItemRC", 1)
                
            Case ORD_POR_FORNECEDOR
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilFornCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "SerieNF", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "NumNF", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEntrada", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemNFItemRC", 1)
                
            Case Else
                gError 74950

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 73856, 74950

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170303)

    End Select

    Exit Sub

End Sub


Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCodFilial_I As String
Dim sCodFilial_F As String
Dim sNomeFilial_I As String
Dim sNomeFilial_F As String
Dim sNumero_I As String
Dim sNumero_F As String
Dim sSerie_I As String
Dim sSerie_F As String
Dim sForn_I As String
Dim sForn_F As String
Dim sOrdenacaoPor As String
Dim iOrdenacao As Long
Dim sOrd As String

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sSerie_I, sSerie_F, sNumero_I, sNumero_F, sForn_I, sForn_F)
    If lErro <> SUCESSO Then gError 73912

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 73913

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sCodFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 73914

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73915

    lErro = objRelOpcoes.IncluirParametro("NNOTAFISCALINIC", sNumero_I)
    If lErro <> AD_BOOL_TRUE Then gError 73916

    lErro = objRelOpcoes.IncluirParametro("NCODFORNINIC", sForn_I)
    If lErro <> AD_BOOL_TRUE Then gError 73917

    lErro = objRelOpcoes.IncluirParametro("TSERIEINIC", sSerie_I)
    If lErro <> AD_BOOL_TRUE Then gError 73918

    'Preenche data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73919

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sCodFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 73920

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 73921

    lErro = objRelOpcoes.IncluirParametro("NNOTAFISCALFIM", sNumero_F)
    If lErro <> AD_BOOL_TRUE Then gError 73922

    lErro = objRelOpcoes.IncluirParametro("NCODFORNFIM", sForn_F)
    If lErro <> AD_BOOL_TRUE Then gError 73923

    lErro = objRelOpcoes.IncluirParametro("TSERIEFIM", sSerie_F)
    If lErro <> AD_BOOL_TRUE Then gError 73924

    'Preenche data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73925

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_NF

                sOrdenacaoPor = "NotaFiscal"

            Case ORD_POR_DATA

                sOrdenacaoPor = "Data"

            Case ORD_POR_FORNECEDOR

                sOrdenacaoPor = "Fornecedor"


            Case Else
                gError 73926

    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 73927

    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 73928

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sNumero_I, sNumero_F, sForn_I, sForn_F, sSerie_I, sSerie_F, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 73929

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 73912 To 73929

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170304)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sSerie_I As String, sSerie_F As String, sNumero_I As String, sNumero_F As String, sForn_I As String, sForn_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'critica Codigo da Filial Inicial e Final
    If CodFilialDe.Text <> "" Then
        sCodFilial_I = CStr(CodFilialDe.Text)
    Else
        sCodFilial_I = ""
    End If

    If CodFilialAte.Text <> "" Then
        sCodFilial_F = CStr(CodFilialAte.Text)
    Else
        sCodFilial_F = ""
    End If

    If sCodFilial_I <> "" And sCodFilial_F <> "" Then

        If StrParaInt(sCodFilial_I) > StrParaInt(sCodFilial_F) Then gError 73833

    End If

    If NomeDe.Text <> "" Then
        sNomeFilial_I = NomeDe.Text
    Else
        sNomeFilial_I = ""
    End If

    If NomeAte.Text <> "" Then
        sNomeFilial_F = NomeAte.Text
    Else
        sNomeFilial_F = ""
    End If

    If sNomeFilial_I <> "" And sNomeFilial_F <> "" Then
        If sNomeFilial_I > sNomeFilial_F Then gError 73834
    End If

    'critica NumeroNF Inicial e Final
    If NumeroDe.Text <> "" Then
        sNumero_I = CStr(NumeroDe.Text)
    Else
        sNumero_I = ""
    End If

    If NumeroAte.Text <> "" Then
        sNumero_F = CStr(NumeroAte.Text)
    Else
        sNumero_F = ""
    End If

    If sNumero_I <> "" And sNumero_F <> "" Then

        If StrParaLong(sNumero_I) > StrParaLong(sNumero_F) Then gError 73835

    End If

    'data inicial não pode ser maior que a final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 73836
    
    End If
    
    'critica Fornecedor Inicial e Final
    If FornecedorDe.Text <> "" Then
        sForn_I = CStr(FornecedorDe.Text)
    Else
        sForn_I = ""
    End If

    If FornecedorAte.Text <> "" Then
        sForn_F = CStr(FornecedorAte.Text)
    Else
        sForn_F = ""
    End If

    If sForn_I <> "" And sForn_F <> "" Then

        If StrParaLong(sForn_I) > StrParaLong(sForn_F) Then gError 73844

    End If

    'critica Serie Inicial e Final
    If SerieDe.Text <> "" Then
        sSerie_I = CStr(SerieDe.Text)
    Else
        sSerie_I = ""
    End If

    If SerieAte.Text <> "" Then
        sSerie_F = CStr(SerieAte.Text)
    Else
        sSerie_F = ""
    End If

    If sSerie_I <> "" And sSerie_F <> "" Then

        If sSerie_I > sSerie_F Then gError 73845

    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 73833
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodFilialDe.SetFocus

        Case 73834
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeDe.SetFocus

        Case 73835
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMNF_INICIAL_MAIOR", gErr)
            NumeroDe.SetFocus

        Case 72836
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus

        Case 73844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            FornecedorDe.SetFocus

        Case 73845
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_INICIAL_MAIOR", gErr)
            SerieDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170305)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sNumero_I As String, sNumero_F As String, sForn_I As String, sForn_F As String, sSerie_I As String, sSerie_F As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao


   If sCodFilial_I <> "" Then sExpressao = "FilEmpCod >= " & Forprint_ConvInt(StrParaInt(sCodFilial_I))

   If sCodFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCod <= " & Forprint_ConvInt(StrParaInt(sCodFilial_F))

    End If

   If sNomeFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNome >= " & Forprint_ConvTexto(sNomeFilial_I)

    End If

    If sNomeFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNome <= " & Forprint_ConvTexto(sNomeFilial_F)

    End If

    If sNumero_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NF >= " & Forprint_ConvLong(StrParaLong(sNumero_I))

    End If

    If sNumero_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NF <= " & Forprint_ConvLong(StrParaLong(sNumero_F))

    End If
    
    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Entrada >= " & Forprint_ConvData(CDate(DataDe.Text))

    End If

    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Entrada <= " & Forprint_ConvData(CDate(DataAte.Text))

    End If

    If sSerie_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Serie >= " & Forprint_ConvTexto((sSerie_I))

    End If

    If sSerie_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Serie <= " & Forprint_ConvTexto((sSerie_F))

    End If

    If sForn_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FornCod >= " & Forprint_ConvLong(StrParaLong(sForn_I))

    End If

    If sForn_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FornCod <= " & Forprint_ConvLong(StrParaLong(sForn_F))

    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170306)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sOrdenacaoPor As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 73900

    'pega Codigo Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73908

    CodFilialDe.Text = sParam
    Call CodFilialDe_Validate(bSGECancelDummy)

    'pega  Codigo Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73909

    CodFilialAte.Text = sParam
    Call CodFilialAte_Validate(bSGECancelDummy)

    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 73812

    NomeDe.Text = sParam
    Call NomeDe_Validate(bSGECancelDummy)

    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 73813

    NomeAte.Text = sParam
    Call NomeAte_Validate(bSGECancelDummy)

    'pega  Numero inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NNOTAFISCALINIC", sParam)
    If lErro <> SUCESSO Then gError 73817

    NumeroDe.Text = sParam

    'pega numero final e exibe
    lErro = objRelOpcoes.ObterParametro("NNOTAFISCALFIM", sParam)
    If lErro <> SUCESSO Then gError 73818

    NumeroAte.Text = sParam

    'pega Fornecedor Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNINIC", sParam)
    If lErro <> SUCESSO Then gError 73823

    FornecedorDe.Text = sParam

    'pega Fornecedor Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFORNFIM", sParam)
    If lErro <> SUCESSO Then gError 73824

    FornecedorAte.Text = sParam

    'pega  SErie Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TSERIEINIC", sParam)
    If lErro <> SUCESSO Then gError 73825

    SerieDe.Text = sParam
    Call SerieDe_Validate(bSGECancelDummy)

    'pega serie Final e exibe
    lErro = objRelOpcoes.ObterParametro("TSERIEFIM", sParam)
    If lErro <> SUCESSO Then gError 73827

    SerieAte.Text = sParam
    Call SerieAte_Validate(bSGECancelDummy)

    'pega data  inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINIC", sParam)
    If lErro <> SUCESSO Then gError 73828

    Call DateParaMasked(DataDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 73930

    Call DateParaMasked(DataAte, CDate(sParam))

    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 73931

    Select Case sOrdenacaoPor

            Case "Data"

                ComboOrdenacao.ListIndex = ORD_POR_DATA

            Case "Fornecedor"
                
                ComboOrdenacao.ListIndex = ORD_POR_FORNECEDOR

            Case "NF"
                
                ComboOrdenacao.ListIndex = ORD_POR_NF

            Case Else
                gError 73932

    End Select

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 73900, 73908, 73909, 73812, 73813, 73817, 73818, 73823
        
        Case 73824, 73825, 73827, 73828, 73930, 73931, 73932

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170307)

    End Select

    Exit Function

End Function


Private Sub CodFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodFilialDe_Validate

    If Len(Trim(CodFilialDe.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialDe.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 73933

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 73934

    End If

    Exit Sub

Erro_CodFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 73933

        Case 73934
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170308)

    End Select

    Exit Sub

End Sub
Private Sub CodFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodFilialAte_Validate

    If Len(Trim(CodFilialAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialAte.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 73935

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 73936

    End If

    Exit Sub

Erro_CodFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 73935

        Case 73936
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170309)

    End Select

    Exit Sub

End Sub

Private Sub NomeDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeDe_Validate

    bAchou = False

    If Len(Trim(NomeDe.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 73937

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeDe.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 73938

        NomeDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73937

        Case 73938
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170310)

    End Select

Exit Sub

End Sub

Private Sub NomeAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeAte_Validate

    bAchou = False
    If Len(Trim(NomeAte.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 73939

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = NomeAte.Text Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 73940

        NomeAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeAte_Validate:

    Cancel = True


    Select Case gErr

        Case 73939

        Case 73940
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170311)

    End Select

Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

''    Parent.HelpContextID = IDH_RELOP_REQ
    Set Form_Load_Ocx = Me
    Caption = "Relação de Notas Fiscais de Requisições de Compra"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpNFRequisicoes"

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
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is NumeroDe Then
            Call LabelNumeroDe_Click

        ElseIf Me.ActiveControl Is NumeroAte Then
            Call LabelNumeroAte_Click

        ElseIf Me.ActiveControl Is CodFilialDe Then
            Call LabelCodFilialDe_Click

        ElseIf Me.ActiveControl Is CodFilialAte Then
            Call LabelCodFilialAte_Click

        ElseIf Me.ActiveControl Is NomeDe Then
            Call LabelNomeDe_Click

        ElseIf Me.ActiveControl Is NomeAte Then
            Call LabelNomeAte_Click

        ElseIf Me.ActiveControl Is FornecedorDe Then
            Call LabelFornecedorDe_Click

        ElseIf Me.ActiveControl Is FornecedorAte Then
            Call LabelFornecedorAte_Click

        ElseIf Me.ActiveControl Is SerieDe Then
            Call LabelSerieDe_Click

        ElseIf Me.ActiveControl Is SerieAte Then
            Call LabelSerieAte_Click

        End If

    End If

End Sub


Private Sub LabelCodFilialDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFilialDe, Source, X, Y)
End Sub

Private Sub LabelCodFilialDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFilialDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodFilialAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFilialAte, Source, X, Y)
End Sub

Private Sub LabelCodFilialAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFilialAte, Button, Shift, X, Y)
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

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 73941

    'Carrega na combo
    For Each objSerie In colSerie
        SerieDe.AddItem objSerie.sSerie
        SerieAte.AddItem objSerie.sSerie
    Next

    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 73941

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170312)

    End Select

    Exit Function

End Function



Private Sub LabelFornecedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorAte, Source, X, Y)
End Sub

Private Sub LabelFornecedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorDe, Source, X, Y)
End Sub

Private Sub LabelFornecedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorDe, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelNumeroAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumeroAte, Source, X, Y)
End Sub

Private Sub LabelNumeroAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumeroAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNumeroDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumeroDe, Source, X, Y)
End Sub

Private Sub LabelNumeroDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumeroDe, Button, Shift, X, Y)
End Sub

Private Sub LabelSerieDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSerieDe, Source, X, Y)
End Sub

Private Sub LabelSerieDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSerieDe, Button, Shift, X, Y)
End Sub

Private Sub LabelSerieAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSerieAte, Source, X, Y)
End Sub

Private Sub LabelSerieAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSerieAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeDe, Source, X, Y)
End Sub

Private Sub LabelNomeDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeAte, Source, X, Y)
End Sub

Private Sub LabelNomeAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeAte, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

