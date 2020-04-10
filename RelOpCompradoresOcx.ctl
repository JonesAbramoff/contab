VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpCompradoresOcx 
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   ScaleHeight     =   4605
   ScaleWidth      =   8280
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpCompradoresOcx.ctx":0000
      Left            =   2025
      List            =   "RelOpCompradoresOcx.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4035
      Width           =   3270
   End
   Begin VB.Frame Frame1 
      Caption         =   "Compradores"
      Height          =   2865
      Left            =   570
      TabIndex        =   16
      Top             =   930
      Width           =   6660
      Begin VB.CheckBox CheckInativos 
         Caption         =   "Inclui Inativos"
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
         Left            =   4230
         TabIndex        =   3
         Top             =   480
         Width           =   1875
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filial Empresa"
         Height          =   1560
         Left            =   180
         TabIndex        =   17
         Top             =   1110
         Width           =   6360
         Begin MSMask.MaskEdBox CodigoFilialDe 
            Height          =   300
            Left            =   1125
            TabIndex        =   4
            Top             =   375
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoFilialAte 
            Height          =   300
            Left            =   4230
            TabIndex        =   5
            Top             =   405
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeFilialAte 
            Height          =   300
            Left            =   4230
            TabIndex        =   7
            Top             =   975
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeFilialDe 
            Height          =   300
            Left            =   1125
            TabIndex        =   6
            Top             =   1005
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodigoDe 
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
            Left            =   150
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   21
            Top             =   405
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
            Left            =   180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   20
            Top             =   1050
            Width           =   855
         End
         Begin VB.Label LabelCodigoAte 
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
            Left            =   3180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   19
            Top             =   435
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
            Left            =   3180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   18
            Top             =   1035
            Width           =   900
         End
      End
      Begin MSMask.MaskEdBox CompradorDe 
         Height          =   300
         Left            =   1635
         TabIndex        =   1
         Top             =   330
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CompradorAte 
         Height          =   300
         Left            =   1620
         TabIndex        =   2
         Top             =   765
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelCompradorAte 
         AutoSize        =   -1  'True
         Caption         =   "Comprador Até:"
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
         TabIndex        =   24
         Top             =   810
         Width           =   1320
      End
      Begin VB.Label LabelCompradorDe 
         AutoSize        =   -1  'True
         Caption         =   "Comprador De:"
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
         TabIndex        =   23
         Top             =   345
         Width           =   1275
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
      Left            =   3915
      Picture         =   "RelOpCompradoresOcx.ctx":001C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   180
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5985
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   210
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpCompradoresOcx.ctx":011E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpCompradoresOcx.ctx":029C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpCompradoresOcx.ctx":07CE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpCompradoresOcx.ctx":0958
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpCompradoresOcx.ctx":0AB2
      Left            =   885
      List            =   "RelOpCompradoresOcx.ctx":0AB4
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   285
      Width           =   2730
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
      Left            =   630
      TabIndex        =   22
      Top             =   4095
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
      TabIndex        =   15
      Top             =   330
      Width           =   615
   End
End
Attribute VB_Name = "RelOpCompradoresOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpCompradores
Const ORD_POR_CODIGO = 0
Const ORD_POR_NOME = 1


Private WithEvents objEventoCodigoDe As AdmEvento
Attribute objEventoCodigoDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoAte As AdmEvento
Attribute objEventoCodigoAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeDe As AdmEvento
Attribute objEventoNomeDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeAte As AdmEvento
Attribute objEventoNomeAte.VB_VarHelpID = -1
Private WithEvents objEventoCodigoCompradorDe As AdmEvento
Attribute objEventoCodigoCompradorDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoCompradorAte As AdmEvento
Attribute objEventoCodigoCompradorAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeCompradorDe As AdmEvento
Attribute objEventoNomeCompradorDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeCompradorAte As AdmEvento
Attribute objEventoNomeCompradorAte.VB_VarHelpID = -1
Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 68563

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 68564

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 68564

        Case 68563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167657)

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
    If lErro <> SUCESSO Then gError 68565

    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CheckInativos.Value = vbUnchecked
    
    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 68565

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167658)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub


Public Sub Form_Load()

Dim lErro As Long


On Error GoTo Erro_Form_Load

    Set objEventoCodigoDe = New AdmEvento
    Set objEventoCodigoAte = New AdmEvento

    Set objEventoNomeDe = New AdmEvento
    Set objEventoNomeAte = New AdmEvento

    Set objEventoCodigoCompradorDe = New AdmEvento
    Set objEventoCodigoCompradorAte = New AdmEvento

    ComboOrdenacao.ListIndex = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167659)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoCodigoDe = Nothing
    Set objEventoCodigoAte = Nothing

    Set objEventoNomeDe = Nothing
    Set objEventoNomeAte = Nothing

    Set objEventoCodigoCompradorDe = Nothing
    Set objEventoCodigoCompradorAte = Nothing

End Sub


Private Sub CodigoFilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialAte, iAlterado)
    
End Sub

Private Sub CodigoFilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialDe, iAlterado)
    
End Sub

Private Sub CompradorDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_CompradorDe_Validate

    If Len(Trim(CompradorDe.Text)) > 0 Then

        lErro = CF("TP_Comprador_Le", CompradorDe, objComprador, 0)
        If lErro <> SUCESSO Then gError 68608
        
        CompradorDe.Text = CStr(objComprador.iCodigo)
        
    End If

    Exit Sub

Erro_CompradorDe_Validate:

    Cancel = True


    Select Case gErr

        Case 68608

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167660)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodigoAte_Click

    If Len(Trim(CodigoFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodigoAte)

   Exit Sub

Erro_LabelCodigoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167661)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodigoDe_Click

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodigoDe)

   Exit Sub

Erro_LabelCodigoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167662)

    End Select

    Exit Sub

End Sub


Private Sub LabelCompradorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelCompradorAte_Click

    If Len(Trim(CompradorAte.Text)) > 0 Then
        'Preenche com o Comprador da tela
        objComprador.iCodigo = Codigo_Extrai(CompradorAte.Text)
        
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoCodigoCompradorAte)

   Exit Sub

Erro_LabelCompradorAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167663)

    End Select

    Exit Sub

End Sub

Private Sub LabelCompradorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelCompradorDe_Click

    If Len(Trim(CompradorDe.Text)) > 0 Then
        'Preenche com o Comprador da tela
        objComprador.iCodigo = Codigo_Extrai(CompradorDe.Text)
        
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoCodigoCompradorDe)

   Exit Sub

Erro_LabelCompradorDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167664)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeDe_Click

    If Len(Trim(NomeFilialDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objFilialEmpresa.sNome = NomeFilialDe.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeDe)

   Exit Sub

Erro_LabelNomeDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167665)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeAte_Click

    If Len(Trim(NomeFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.sNome = NomeFilialAte.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeAte)

   Exit Sub

Erro_LabelNomeAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167666)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigoDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoCompradorAte_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    CompradorAte.Text = objComprador.iCodigo & SEPARADOR & objComprador.sNomeReduzido

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoCodigoCompradorDe_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    CompradorDe.Text = objComprador.iCodigo & SEPARADOR & objComprador.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeFilialDe.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeFilialAte.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 68566

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 68567

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 68568

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 68569

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 68566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 68567, 68568, 68569

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167667)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 68570

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 68571

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 68570
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 68571

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167668)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 68572

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 68572

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167669)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCodFil_I As String
Dim sCodFil_F As String
Dim sCodComp_I As String
Dim sCodComp_F As String
Dim sNomeFil_I As String
Dim sNomeFil_F As String
Dim sOrdenacaoPor As String
Dim sOrd As String
Dim sCheck  As String

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sCodFil_I, sCodFil_F, sNomeFil_I, sNomeFil_F, sCodComp_I, sCodComp_F)
    If lErro <> SUCESSO Then gError 68573

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 68574

    lErro = objRelOpcoes.IncluirParametro("NCODIGOFILINIC", sCodFil_I)
    If lErro <> AD_BOOL_TRUE Then gError 68575

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILINIC", NomeFilialDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68576

    lErro = objRelOpcoes.IncluirParametro("NCODIGOCOMPINIC", sCodComp_I)
    If lErro <> AD_BOOL_TRUE Then gError 68577
    
    lErro = objRelOpcoes.IncluirParametro("NCODIGOFILFIM", sCodFil_F)
    If lErro <> AD_BOOL_TRUE Then gError 68578

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILFIM", NomeFilialAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68579

    lErro = objRelOpcoes.IncluirParametro("NCODIGOCOMPFIM", sCodComp_F)
    If lErro <> AD_BOOL_TRUE Then gError 68580

    'Exibe Inativos
    If CheckInativos.Value Then
        sCheck = "S"
    Else
        sCheck = "N"
    End If

    lErro = objRelOpcoes.IncluirParametro("TINATIVOS", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 68624

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO

                sOrdenacaoPor = "CodComp"

            Case ORD_POR_NOME

                sOrdenacaoPor = "NomeComp"

            Case Else
                gError 68581

    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 68582

    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 68583

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodFil_I, sCodFil_F, sNomeFil_I, sNomeFil_F, sCodComp_I, sCodComp_F, sCheck, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 68584

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 68573 To 68584, 68624

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167670)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodFil_I As String, sCodFil_F As String, sNomeFil_I As String, sNomeFil_F As String, sCodComp_I As String, sCodComp_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'critica CodigoFilial Inicial e Final
    If CodigoFilialDe.Text <> "" Then
        sCodFil_I = CStr(CodigoFilialDe.Text)
    Else
        sCodFil_I = ""
    End If

    If CodigoFilialAte.Text <> "" Then
        sCodFil_F = CStr(CodigoFilialAte.Text)
    Else
        sCodFil_F = ""
    End If

    If sCodFil_I <> "" And sCodFil_F <> "" Then

        If StrParaInt(sCodFil_I) > StrParaInt(sCodFil_F) Then gError 68587

    End If

    If NomeFilialDe.Text <> "" Then
        sNomeFil_I = NomeFilialDe.Text
    Else
        sNomeFil_I = ""
    End If

    If NomeFilialAte.Text <> "" Then
        sNomeFil_F = NomeFilialAte.Text
    Else
        sNomeFil_F = ""
    End If

    If sNomeFil_I <> "" And sNomeFil_F <> "" Then
        If sNomeFil_I > sNomeFil_F Then gError 68588
    End If

    'critica CodigoComprador Inicial e Final
    If CompradorDe.Text <> "" Then
        sCodComp_I = CStr(Codigo_Extrai(CompradorDe.Text))
    Else
        sCodComp_I = ""
    End If

    If CompradorAte.Text <> "" Then
        sCodComp_F = CStr(Codigo_Extrai(CompradorAte.Text))
    Else
        sCodComp_F = ""
    End If

    If sCodComp_I <> "" And sCodComp_F <> "" Then

        If StrParaInt(sCodComp_I) > StrParaInt(sCodComp_F) Then gError 68589

    End If

    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 68587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodigoFilialDe.SetFocus

        Case 68588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeFilialDe.SetFocus

        Case 68589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_INICIAL_MAIOR", gErr)
            CompradorDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167671)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodFil_I As String, sCodFil_F As String, sNomeFil_I As String, sNomeFil_F As String, sCodComp_I As String, sCodComp_F As String, sInativo As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sCodFil_I <> "" Then sExpressao = "FilialEmpresaCodigo >= " & Forprint_ConvInt(StrParaInt(sCodFil_I))

   If sCodFil_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresaCodigo <= " & Forprint_ConvInt(StrParaInt(sCodFil_F))

    End If

   If sNomeFil_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresaNome >= " & Forprint_ConvTexto(sNomeFil_I)

    End If

    If sNomeFil_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresaNome <= " & Forprint_ConvTexto(sNomeFil_F)

    End If

    If sCodComp_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Codigo >= " & Forprint_ConvLong(StrParaLong(sCodComp_I))

    End If

    If sCodComp_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Codigo <= " & Forprint_ConvLong(StrParaLong(sCodComp_F))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167672)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sOrdenacaoPor As String


On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 68591

    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODIGOFILINIC", sParam)
    If lErro <> SUCESSO Then gError 68592

    CodigoFilialDe.Text = sParam
    Call CodigoFilialDe_Validate(bSGECancelDummy)

    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODIGOFILFIM", sParam)
    If lErro <> SUCESSO Then gError 68593

    CodigoFilialAte.Text = sParam
    Call CodigoFilialAte_Validate(bSGECancelDummy)

    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILINIC", sParam)
    If lErro <> SUCESSO Then gError 68594

    NomeFilialDe.Text = sParam
    Call NomeFilialDe_Validate(bSGECancelDummy)

    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILFIM", sParam)
    If lErro <> SUCESSO Then gError 68595

    NomeFilialAte.Text = sParam
    Call NomeFilialAte_Validate(bSGECancelDummy)
    
    
    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODIGOCOMPINIC", sParam)
    If lErro <> SUCESSO Then gError 68596

    CompradorDe.Text = sParam
    Call CompradorDe_Validate(bSGECancelDummy)

    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODIGOCOMPFIM", sParam)
    If lErro <> SUCESSO Then gError 68597

    CompradorAte.Text = sParam
    Call CompradorAte_Validate(bSGECancelDummy)

    'pega 'Inclui Inativos' e exibe
    lErro = objRelOpcoes.ObterParametro("TINATIVOS", sParam)
    If lErro <> SUCESSO Then gError 68625

    If sParam = "S" Then
        CheckInativos.Value = vbChecked
    Else
        CheckInativos.Value = vbUnchecked
    End If

    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 68598


    Select Case sOrdenacaoPor

            Case "CodComp"

                ComboOrdenacao.ListIndex = ORD_POR_CODIGO

            Case "NomeComp"

                ComboOrdenacao.ListIndex = ORD_POR_NOME
            
            Case Else
                gError 68599

    End Select

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 68591 To 68599, 68625

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167673)

    End Select

    Exit Function

End Function

Private Sub CodigoFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialDe_Validate

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialDe.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 68602
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 68603

    End If

    Exit Sub

Erro_CodigoFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 68602

        Case 68603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167674)

    End Select

    Exit Sub

End Sub
Private Sub CodigoFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialAte_Validate

    If Len(Trim(CodigoFilialAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialAte.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 68604
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 68605

    End If

    Exit Sub

Erro_CodigoFilialAte_Validate:

    Cancel = True

    Select Case gErr

        Case 68604

        Case 68605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167675)

    End Select

    Exit Sub

End Sub


Private Sub NomeFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim colFiliais As New Collection
Dim bAchou As Boolean

On Error GoTo Erro_NomeFilialDe_Validate

    bAchou = False
    
    If Len(Trim(NomeFilialDe.Text)) > 0 Then
    
        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 68683

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeFilialDe.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 68684
        
        NomeFilialDe.Text = objFilialEmpresa.sNome
        bAchou = False
        
    End If

    Exit Sub

Erro_NomeFilialDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68683

        Case 68684
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167676)

    End Select

Exit Sub

End Sub

Private Sub NomeFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim colFiliais As New Collection
Dim bAchou As Boolean

On Error GoTo Erro_NomeFilialAte_Validate

    bAchou = False
    
    If Len(Trim(NomeFilialAte.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 68685

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeFilialAte.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 68686
        
        NomeFilialAte.Text = objFilialEmpresa.sNome
        bAchou = False
        
    End If

    Exit Sub

Erro_NomeFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 68685

        Case 68686
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167677)

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
        If lErro <> SUCESSO Then gError 68606
        
        CompradorAte.Text = CStr(objComprador.iCodigo)
      
    End If

    Exit Sub

Erro_CompradorAte_Validate:

    Cancel = True

    Select Case gErr

        Case 68606

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167678)

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
    Caption = "Relação de Compradores"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpCompradores"

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

        If Me.ActiveControl Is CompradorDe Then
            Call LabelCompradorDe_Click

        ElseIf Me.ActiveControl Is CompradorAte Then
            Call LabelCompradorAte_Click

        ElseIf Me.ActiveControl Is NomeFilialDe Then
            Call LabelNomeDe_Click

        ElseIf Me.ActiveControl Is NomeFilialAte Then
            Call LabelNomeAte_Click

        ElseIf Me.ActiveControl Is CodigoFilialDe Then
            Call LabelCodigoDe_Click

        ElseIf Me.ActiveControl Is CodigoFilialAte Then
            Call LabelCodigoAte_Click

        End If

    End If

End Sub


Private Sub LabelCodigoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoDe, Source, X, Y)
End Sub

Private Sub LabelCodigoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoAte, Source, X, Y)
End Sub

Private Sub LabelCodigoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoAte, Button, Shift, X, Y)
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

Private Sub LabelCompradorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCompradorAte, Source, X, Y)
End Sub

Private Sub LabelCompradorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCompradorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCompradorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCompradorDe, Source, X, Y)
End Sub

Private Sub LabelCompradorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCompradorDe, Button, Shift, X, Y)
End Sub

