VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPedCotacaoGeracaoOcx 
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   ScaleHeight     =   5205
   ScaleWidth      =   8550
   Begin VB.Frame Frame1 
      Caption         =   "Filial Empresa"
      Height          =   1200
      Left            =   990
      TabIndex        =   33
      Top             =   945
      Width           =   6420
      Begin MSMask.MaskEdBox CodFilialDe 
         Height          =   300
         Left            =   1200
         TabIndex        =   2
         Top             =   345
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodFilialAte 
         Height          =   300
         Left            =   4365
         TabIndex        =   3
         Top             =   330
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeDe 
         Height          =   300
         Left            =   1170
         TabIndex        =   4
         Top             =   780
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeAte 
         Height          =   300
         Left            =   4365
         TabIndex        =   5
         Top             =   780
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
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
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   37
         Top             =   840
         Width           =   855
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
         Left            =   3405
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   36
         Top             =   825
         Width           =   900
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   35
         Top             =   390
         Width           =   960
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
         Left            =   3270
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   34
         Top             =   405
         Width           =   1005
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpPedCotacaoGeracaoOcx.ctx":0000
      Left            =   1560
      List            =   "RelOpPedCotacaoGeracaoOcx.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6285
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPedCotacaoGeracaoOcx.ctx":001D
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPedCotacaoGeracaoOcx.ctx":0177
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPedCotacaoGeracaoOcx.ctx":0301
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPedCotacaoGeracaoOcx.ctx":0833
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
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
      Left            =   4410
      Picture         =   "RelOpPedCotacaoGeracaoOcx.ctx":09B1
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   135
      Width           =   1635
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPedCotacaoGeracaoOcx.ctx":0AB3
      Left            =   1560
      List            =   "RelOpPedCotacaoGeracaoOcx.ctx":0AB5
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   105
      Width           =   2595
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pedidos de Cotação"
      Height          =   2805
      Left            =   990
      TabIndex        =   19
      Top             =   2205
      Width           =   6420
      Begin VB.CheckBox CheckInclui 
         Caption         =   "Inclui Pedidos de Cotação já emitidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   270
         TabIndex        =   12
         Top             =   2295
         Visible         =   0   'False
         Width           =   4110
      End
      Begin VB.Frame Frame7 
         Caption         =   "Número"
         Height          =   660
         Left            =   210
         TabIndex        =   28
         Top             =   255
         Width           =   5580
         Begin MSMask.MaskEdBox CodPCDe 
            Height          =   300
            Left            =   525
            TabIndex        =   6
            Top             =   240
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodPCAte 
            Height          =   300
            Left            =   3405
            TabIndex        =   7
            Top             =   255
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodPCDe 
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
         Begin VB.Label LabelCodPCAte 
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
            Left            =   2955
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   29
            Top             =   285
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data"
         Height          =   675
         Left            =   210
         TabIndex        =   23
         Top             =   1620
         Width           =   5580
         Begin MSComCtl2.UpDown UpDownDataDe 
            Height          =   315
            Left            =   1710
            TabIndex        =   24
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
            Left            =   525
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
            Left            =   4635
            TabIndex        =   25
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
            Left            =   3450
            TabIndex        =   11
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
         Begin VB.Label Label4 
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
            Left            =   210
            TabIndex        =   27
            Top             =   315
            Width           =   315
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
            Left            =   2985
            TabIndex        =   26
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Comprador"
         Height          =   690
         Left            =   210
         TabIndex        =   20
         Top             =   930
         Width           =   5580
         Begin MSMask.MaskEdBox CompradorDe 
            Height          =   300
            Left            =   525
            TabIndex        =   8
            Top             =   240
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CompradorAte 
            Height          =   300
            Left            =   3405
            TabIndex        =   9
            Top             =   240
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCompradorDe 
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
            TabIndex        =   22
            Top             =   300
            Width           =   315
         End
         Begin VB.Label LabelCompradorAte 
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
            Left            =   2985
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   21
            Top             =   300
            Width           =   360
         End
      End
   End
   Begin VB.Label Label2 
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
      TabIndex        =   32
      Top             =   570
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
      Left            =   210
      TabIndex        =   31
      Top             =   135
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPedCotacaoGeracaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpPedCotacaoGeracao
Const ORD_POR_CODIGO = 0
Const ORD_POR_DATA = 1

Private WithEvents objEventoCodPCDe As AdmEvento
Attribute objEventoCodPCDe.VB_VarHelpID = -1
Private WithEvents objEventoCodPCAte As AdmEvento
Attribute objEventoCodPCAte.VB_VarHelpID = -1
Private WithEvents objEventoCompradorDe As AdmEvento
Attribute objEventoCompradorDe.VB_VarHelpID = -1
Private WithEvents objEventoCompradorAte As AdmEvento
Attribute objEventoCompradorAte.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialDe As AdmEvento
Attribute objEventoCodFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialAte As AdmEvento
Attribute objEventoCodFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 73858
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 73859

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 73858
        
        Case 73859
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170944)

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
    If lErro <> SUCESSO Then gError 73860
    
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CheckInclui.Value = vbUnchecked
    
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 73860
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170945)

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
        
    Set objEventoCompradorDe = New AdmEvento
    Set objEventoCompradorAte = New AdmEvento
    
    Set objEventoCodFilialDe = New AdmEvento
    Set objEventoCodFilialAte = New AdmEvento
    
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento
    
    ComboOrdenacao.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170946)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoCodPCDe = Nothing
    Set objEventoCodPCAte = Nothing
        
    Set objEventoCompradorDe = Nothing
    Set objEventoCompradorAte = Nothing
    
    Set objEventoCodFilialDe = Nothing
    Set objEventoCodFilialAte = Nothing
    
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing
    
End Sub

Private Sub CodFilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFilialAte, iAlterado)
    
End Sub

Private Sub CodFilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFilialDe, iAlterado)
    
End Sub

Private Sub CodFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodFilialDe_Validate

    If Len(Trim(CodFilialDe.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialDe.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 74661

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 74662

    End If

    Exit Sub

Erro_CodFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 74661

        Case 74662
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170947)

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
        If lErro <> SUCESSO And lErro <> 27378 Then gError 74663

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 74664

    End If

    Exit Sub

Erro_CodFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 74663

        Case 74664
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170948)

    End Select

    Exit Sub

End Sub

Private Sub CodPCAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodPCAte, iAlterado)
    
End Sub

Private Sub CodPCDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodPCDe, iAlterado)
    
End Sub

Private Sub CompradorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CompradorAte, iAlterado)
    
End Sub

Private Sub CompradorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CompradorDe, iAlterado)
    
End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    
End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    
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
        If lErro <> SUCESSO Then gError 74665

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeDe.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 74666

        NomeDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74665

        Case 74666
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170949)

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
        If lErro <> SUCESSO Then gError 74667

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeAte.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 74668

        NomeAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeAte_Validate:

    Cancel = True


    Select Case gErr

        Case 74667

        Case 74668
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170950)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170951)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170952)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170953)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170954)

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

Private Sub LabelCodPCAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCotacao As New ClassPedidoCotacao

On Error GoTo Erro_LabelCodPCAte_Click

    If Len(Trim(CodPCAte.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objPedCotacao.lCodigo = StrParaLong(CodPCAte.Text)
    End If

    'Chama Tela PedidoCotacaoTodosLista
    Call Chama_Tela("PedidoCotacaoTodosLista", colSelecao, objPedCotacao, objEventoCodPCAte)

   Exit Sub

Erro_LabelCodPCAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170955)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodPCDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_LabelCodPCDe_Click

    If Len(Trim(CodPCDe.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objPedidoCotacao.lCodigo = StrParaLong(CodPCDe.Text)
    End If

    'Chama Tela PedidoCotacaoTodosLista
    Call Chama_Tela("PedidoCotacaoTodosLista", colSelecao, objPedidoCotacao, objEventoCodPCDe)

   Exit Sub
   
Erro_LabelCodPCDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170956)

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
    If lErro <> SUCESSO Then gError 73861

    Exit Sub
                   
Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 73861
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170957)

    End Select

    Exit Sub

End Sub


Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73862

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 73862
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170958)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73863

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 73863
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170959)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 73864

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 73864
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170960)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 73865

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 73865
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170961)

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
    If lErro <> SUCESSO Then gError 73866

    Exit Sub
                   
Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 73866
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170962)

    End Select

    Exit Sub

End Sub



Private Sub LabelCompradorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelCompradorDe_Click

    If Len(Trim(CompradorDe.Text)) > 0 Then
        'Preenche com o comprador da tela
        objComprador.iCodigo = StrParaInt(CompradorDe.Text)
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoCompradorDe)

   Exit Sub

Erro_LabelCompradorDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170963)

    End Select

    Exit Sub

End Sub
Private Sub LabelCompradorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelCompradorAte_Click

    If Len(Trim(CompradorAte.Text)) > 0 Then
        'Preenche com o comprador da tela
        objComprador.iCodigo = StrParaInt(CompradorAte.Text)
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoCompradorAte)

   Exit Sub

Erro_LabelCompradorAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170964)

    End Select

    Exit Sub

End Sub



Private Sub objEventoCodPCAte_evSelecao(obj1 As Object)

Dim objPedCotacao As New ClassPedidoCotacao

    Set objPedCotacao = obj1

    CodPCAte.Text = CStr(objPedCotacao.lCodigo)

    Me.Show

End Sub
Private Sub objEventoCodPCDe_evSelecao(obj1 As Object)

Dim objPedCotacao As New ClassPedidoCotacao

    Set objPedCotacao = obj1

    CodPCDe.Text = CStr(objPedCotacao.lCodigo)

    Me.Show

End Sub

Private Sub objEventoCompradorDe_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    CompradorDe.Text = CStr(objComprador.iCodigo)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoCompradorAte_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    CompradorAte.Text = CStr(objComprador.iCodigo)

    Me.Show

    Exit Sub

End Sub


Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 73867

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73868

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 73869
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 73870
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 73867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 73868 To 73870
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170965)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 73871

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 73872

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 73871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 73872

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170966)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 73873

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CotacaoCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PedCotCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCotacao", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CondPagto", 1)
            
            Case ORD_POR_DATA

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataCotacao", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PedCotCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCotacao", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CondPagto", 1)
                
            Case Else
                gError 74960

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 73873, 74960

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170967)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCodPC_I As String
Dim sCodPC_F As String
Dim sComprador_I As String
Dim sComprador_F As String
Dim sCodFilial_I As String
Dim sCodFilial_F As String
Dim sNomeFilial_I As String
Dim sNomeFilial_F As String
Dim sCheck As String
Dim sOrdenacaoPor As String
Dim sOrd As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodPC_I, sCodPC_F, sComprador_I, sComprador_F)
    If lErro <> SUCESSO Then gError 73874

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 73875
         
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sCodFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 74675

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74676
    
    lErro = objRelOpcoes.IncluirParametro("NCOTACAOINIC", sCodPC_I)
    If lErro <> AD_BOOL_TRUE Then gError 73878
         
    lErro = objRelOpcoes.IncluirParametro("NCODCOMPINIC", sComprador_I)
    If lErro <> AD_BOOL_TRUE Then gError 73879
    
    'Preenche data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATACOTAINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATACOTAINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73880
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sCodFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 74677

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74678

    lErro = objRelOpcoes.IncluirParametro("NCOTACAOFIM", sCodPC_F)
    If lErro <> AD_BOOL_TRUE Then gError 73883
         
    lErro = objRelOpcoes.IncluirParametro("NCODCOMPFIM", sComprador_F)
    If lErro <> AD_BOOL_TRUE Then gError 73884
    
    'Preenche data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATACOTAFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATACOTAFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 73885
    
    'Exibe Pedidos já emitidos
    If CheckInclui.Value Then
        sCheck = vbChecked
    Else
        sCheck = vbUnchecked
    End If

    lErro = objRelOpcoes.IncluirParametro("NPEDEMITIDO", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 73886
    
    Select Case ComboOrdenacao.ListIndex
        
            Case ORD_POR_CODIGO
            
                sOrdenacaoPor = "Codigo"
                
            Case ORD_POR_DATA
                sOrdenacaoPor = "Data"
            
            Case Else
                gError 74689
                  
    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 74690
   
    sOrd = ComboOrdenacao.ListIndex
    
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 74691
       
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodPC_I, sCodPC_F, sComprador_I, sComprador_F, sCheck)
    If lErro <> SUCESSO Then gError 73890

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 73874, 73875, 73878, 73879, 73880, 73883, 73884
        
        Case 73885, 73886, 73890, 74676, 74677, 74678, 74679
        
        Case 74689, 74690, 74691
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170968)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodPC_I As String, sCodPC_F As String, sComprador_I As String, sComprador_F As String) As Long
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

        If StrParaInt(sCodFilial_I) > StrParaInt(sCodFilial_F) Then gError 74669

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
        If sNomeFilial_I > sNomeFilial_F Then gError 74670
    End If

    'critica CodigoPC Inicial e Final
    If CodPCDe.Text <> "" Then
        sCodPC_I = CStr(CodPCDe.Text)
    Else
        sCodPC_I = ""
    End If

    If CodPCAte.Text <> "" Then
        sCodPC_F = CStr(CodPCAte.Text)
    Else
        sCodPC_F = ""
    End If

    If sCodPC_I <> "" And sCodPC_F <> "" Then

        If StrParaLong(sCodPC_I) > StrParaLong(sCodPC_F) Then gError 73893

    End If
    
    'critica Comprador Inicial e Final
    If CompradorDe.Text <> "" Then
        sComprador_I = CStr(CompradorDe.Text)
    Else
        sComprador_I = ""
    End If
    
    If CompradorAte.Text <> "" Then
        sComprador_F = CStr(CompradorAte.Text)
    Else
        sComprador_F = ""
    End If
            
    If sComprador_I <> "" And sComprador_F <> "" Then
        
        If StrParaInt(sComprador_I) > StrParaInt(sComprador_F) Then gError 73894
        
    End If
    
    'data  inicial não pode ser maior que a data  final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 73895
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                        
        Case 73893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PC_INICIAL_MAIOR", gErr)
            CodPCDe.SetFocus
        
        Case 73894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_INICIAL_MAIOR", gErr)
            CompradorDe.SetFocus
        
        Case 73895
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
            
        Case 74669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodFilialDe.SetFocus

        Case 74670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170969)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodPC_I As String, sCodPC_F As String, sComprador_I As String, sComprador_F As String, sCheck As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    If sCodFilial_I <> "" Then sExpressao = "FilEmpCodInic"

    If sCodFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCodFim"

    End If

    If sNomeFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeInic"

    End If

    If sNomeFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeFim"

    End If

    If sCodPC_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PCCod >= " & Forprint_ConvLong(StrParaLong(sCodPC_I))

    End If
   
    If sCodPC_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PCCod <= " & Forprint_ConvLong(StrParaLong(sCodPC_F))

    End If
   
    If sComprador_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CompCod >= " & Forprint_ConvInt(StrParaInt(sComprador_I))

    End If
   
    If sComprador_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CompCod <= " & Forprint_ConvInt(StrParaInt(sComprador_F))

    End If
    
    If Trim(DataDe.ClipText) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataCot >= " & Forprint_ConvData(CDate(DataDe.Text))

    End If
    
    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataCot <= " & Forprint_ConvData(CDate(DataAte.Text))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170970)

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
    If lErro <> SUCESSO Then gError 73896
   
    'pega Codigo Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 74671

    CodFilialDe.Text = sParam
    Call CodFilialDe_Validate(bSGECancelDummy)

    'pega  Codigo Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 74672

    CodFilialAte.Text = sParam
    Call CodFilialAte_Validate(bSGECancelDummy)

    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 74673

    NomeDe.Text = sParam
    Call NomeDe_Validate(bSGECancelDummy)

    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 74674

    NomeAte.Text = sParam
    Call NomeAte_Validate(bSGECancelDummy)
   
    'pega  Codigo PC inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCOTACAOINIC", sParam)
    If lErro <> SUCESSO Then gError 73901
                   
    CodPCDe.Text = sParam
                                        
    'pega  Codigo PC final e exibe
    lErro = objRelOpcoes.ObterParametro("NCOTACAOFIM", sParam)
    If lErro <> SUCESSO Then gError 73902
                   
    CodPCAte.Text = sParam
                        
    'pega  Comprador Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCOMPINIC", sParam)
    If lErro <> SUCESSO Then gError 73903
                   
    CompradorDe.Text = sParam
    Call CompradorDe_Validate(bSGECancelDummy)
    
    'pega  comprador Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCOMPFIM", sParam)
    If lErro <> SUCESSO Then gError 73904
                   
    CompradorAte.Text = sParam
    Call CompradorAte_Validate(bSGECancelDummy)
                                   
    'pega data  inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATACOTAINIC", sParam)
    If lErro <> SUCESSO Then gError 73905

    Call DateParaMasked(DataDe, CDate(sParam))
       
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATACOTAFIM", sParam)
    If lErro <> SUCESSO Then gError 73906
    
    Call DateParaMasked(DataAte, CDate(sParam))
       
    lErro = objRelOpcoes.ObterParametro("NPEDEMITIDO", sParam)
    If lErro <> SUCESSO Then gError 73907

    If sParam = "1" Then
        CheckInclui.Value = vbChecked
    Else
        CheckInclui.Value = vbUnchecked
    End If
    
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 74687
    
    Select Case sOrdenacaoPor
        
            Case "Codigo"
            
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
            Case "Data"
            
                ComboOrdenacao.ListIndex = ORD_POR_DATA
            
            Case Else
                gError 74688
                  
    End Select
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 73896, 73901 To 73907, 74671, 74672, 74673, 74674
        
        Case 74687, 74688
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170971)

    End Select

    Exit Function

End Function

Private Sub CompradorDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_CompradorDe_Validate

    If Len(Trim(CompradorDe.Text)) > 0 Then

        lErro = CF("TP_Comprador_Le", CompradorDe, objComprador, 0)
        If lErro <> SUCESSO Then gError 73910
        
        CompradorDe.Text = objComprador.iCodigo
        
    End If

    Exit Sub

Erro_CompradorDe_Validate:

    Cancel = True


    Select Case gErr

        Case 73910

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170972)

    End Select

    Exit Sub

End Sub

Private Sub CompradorAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_CompradorAte_Validate

    If Len(Trim(CompradorAte.Text)) > 0 Then

        'Lê o código informado
        lErro = CF("TP_Comprador_Le", CompradorAte, objComprador, 0)
        If lErro <> SUCESSO Then gError 73911
                
        CompradorAte.Text = objComprador.iCodigo
        
    End If

    Exit Sub

Erro_CompradorAte_Validate:

    Cancel = True

    Select Case gErr

        Case 73911

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170973)

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
    Caption = "Pedidos de Cotação Emitidos por Geração"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPedCotacaoGeracao"
    
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
        
        If Me.ActiveControl Is CodPCDe Then
            Call LabelCodPCDe_Click
            
        ElseIf Me.ActiveControl Is CodPCAte Then
            Call LabelCodPCAte_Click
           
        ElseIf Me.ActiveControl Is CompradorDe Then
            Call LabelCompradorDe_Click
        
        ElseIf Me.ActiveControl Is CompradorAte Then
            Call LabelCompradorAte_Click
        
        End If
    
    End If

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

Private Sub LabelCodPCDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodPCDe, Source, X, Y)
End Sub

Private Sub LabelCodPCDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodPCDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodPCAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodPCAte, Source, X, Y)
End Sub

Private Sub LabelCodPCAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodPCAte, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelCompradorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCompradorDe, Source, X, Y)
End Sub

Private Sub LabelCompradorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCompradorDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCompradorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCompradorAte, Source, X, Y)
End Sub

Private Sub LabelCompradorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCompradorAte, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

