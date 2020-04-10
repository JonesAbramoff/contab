VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpRequisitantesOcx 
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   KeyPreview      =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   7950
   Begin VB.Frame Frame1 
      Caption         =   "Requisitantes"
      Height          =   2850
      Left            =   255
      TabIndex        =   16
      Top             =   1020
      Width           =   5520
      Begin VB.Frame FrameCcl 
         Caption         =   "Centro de Custo"
         Height          =   675
         Left            =   180
         TabIndex        =   19
         Top             =   2040
         Width           =   5160
         Begin MSMask.MaskEdBox CclDe 
            Height          =   315
            Left            =   630
            TabIndex        =   6
            Top             =   270
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CclAte 
            Height          =   315
            Left            =   3105
            TabIndex        =   7
            Top             =   270
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin VB.Label LabelCclAte 
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
            Left            =   2640
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   25
            Top             =   330
            Width           =   360
         End
         Begin VB.Label LabelCclDe 
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
            TabIndex        =   24
            Top             =   330
            Width           =   315
         End
      End
      Begin VB.Frame FrameNome 
         Caption         =   "Nome"
         Height          =   765
         Left            =   180
         TabIndex        =   18
         Top             =   1155
         Width           =   5160
         Begin MSMask.MaskEdBox NomeDe 
            Height          =   300
            Left            =   525
            TabIndex        =   4
            Top             =   285
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeAte 
            Height          =   300
            Left            =   3060
            TabIndex        =   5
            Top             =   285
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelNomeDe 
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
            TabIndex        =   21
            Top             =   338
            Width           =   315
         End
         Begin VB.Label LabelNomeAte 
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
            Left            =   2625
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   20
            Top             =   338
            Width           =   360
         End
      End
      Begin VB.Frame FrameCodigo 
         Caption         =   "Código"
         Height          =   825
         Left            =   180
         TabIndex        =   17
         Top             =   285
         Width           =   5160
         Begin MSMask.MaskEdBox CodigoDe 
            Height          =   300
            Left            =   525
            TabIndex        =   2
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoAte 
            Height          =   300
            Left            =   3105
            TabIndex        =   3
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodigoAte 
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
            Left            =   2655
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   23
            Top             =   405
            Width           =   360
         End
         Begin VB.Label LabelCodigoDe 
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
            Left            =   150
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   22
            Top             =   405
            Width           =   315
         End
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpRequisitantesOcx.ctx":0000
      Left            =   1560
      List            =   "RelOpRequisitantesOcx.ctx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5640
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRequisitantesOcx.ctx":0024
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRequisitantesOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRequisitantesOcx.ctx":0308
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRequisitantesOcx.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   13
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
      Left            =   6060
      Picture         =   "RelOpRequisitantesOcx.ctx":09B8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1095
      Width           =   1590
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRequisitantesOcx.ctx":0ABA
      Left            =   1560
      List            =   "RelOpRequisitantesOcx.ctx":0ABC
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   3135
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
      TabIndex        =   15
      Top             =   540
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
      Left            =   225
      TabIndex        =   14
      Top             =   135
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRequisitantesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpRequisitantes
Const ORD_POR_CODIGO = 0
Const ORD_POR_NOME = 1
Const ORD_POR_CCL = 2

Private WithEvents objEventoCodigoDe As AdmEvento
Attribute objEventoCodigoDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoAte As AdmEvento
Attribute objEventoCodigoAte.VB_VarHelpID = -1
Private WithEvents objEventoCclDe As AdmEvento
Attribute objEventoCclDe.VB_VarHelpID = -1
Private WithEvents objEventoCclAte As AdmEvento
Attribute objEventoCclAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeDe As AdmEvento
Attribute objEventoNomeDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeAte As AdmEvento
Attribute objEventoNomeAte.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 68505
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 68506

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 68506
        
        Case 68505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172944)

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
    If lErro <> SUCESSO Then gError 68507
    
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 68507
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172945)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub


Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraCcl As String

On Error GoTo Erro_Form_Load
    
    Set objEventoCodigoDe = New AdmEvento
    Set objEventoCodigoAte = New AdmEvento
        
    Set objEventoNomeDe = New AdmEvento
    Set objEventoNomeAte = New AdmEvento
        
    Set objEventoCclDe = New AdmEvento
    Set objEventoCclAte = New AdmEvento
            
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 68558

    CclDe.Mask = sMascaraCcl
    CclAte.Mask = sMascaraCcl
    
    ComboOrdenacao.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 68558
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172946)

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
    
    Set objEventoCclDe = Nothing
    Set objEventoCclAte = Nothing
    
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub LabelCodigoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_LabelCodigoAte_Click

    If Len(Trim(CodigoAte.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objRequisitante.lCodigo = StrParaLong(CodigoAte.Text)
    End If

    'Chama Tela RequisitanteLista
    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoCodigoAte)

   Exit Sub

Erro_LabelCodigoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172947)

    End Select

    Exit Sub

End Sub


Private Sub LabelCodigoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_LabelCodigoDe_Click

    If Len(Trim(CodigoDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objRequisitante.lCodigo = StrParaLong(CodigoDe.Text)
    End If

    'Chama Tela RequisitanteLista
    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoCodigoDe)

   Exit Sub

Erro_LabelCodigoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172948)

    End Select

    Exit Sub

End Sub

Private Sub LabelCclAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_LabelCclAte_Click

    If Len(Trim(CclAte.Text)) > 0 Then
        
        lErro = CF("Ccl_Formata", CclAte.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 68552
        
        'Preenche com o Ccl
        objCcl.sCcl = sCclFormata
        
    End If

    'Chama Tela Cclista
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclAte)

   Exit Sub

Erro_LabelCclAte_Click:

    Select Case gErr

        Case 68552
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172949)

    End Select

    Exit Sub

End Sub

Private Sub LabelCclDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_LabelCclDe_Click

    If Len(Trim(CclDe.Text)) > 0 Then
        
        lErro = CF("Ccl_Formata", CclDe.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 68553

        'Preenche com o Ccl
        objCcl.sCcl = sCclFormata
        
    End If

    'Chama Tela Cclista
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclDe)

   Exit Sub

Erro_LabelCclDe_Click:

    Select Case gErr

        Case 68553
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172950)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_LabelNomeDe_Click

    If Len(Trim(NomeDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objRequisitante.sNomeReduzido = NomeDe.Text
    End If

    'Chama Tela RequisitanteLista
    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoNomeDe)

   Exit Sub

Erro_LabelNomeDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172951)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_LabelNomeAte_Click

    If Len(Trim(NomeAte.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objRequisitante.sNomeReduzido = NomeAte.Text
    End If

    'Chama Tela RequisitanteLista
    Call Chama_Tela("RequisitanteLista", colSelecao, objRequisitante, objEventoNomeAte)

   Exit Sub

Erro_LabelNomeAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172952)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigoDe_evSelecao(obj1 As Object)

Dim objRequisitante As New ClassRequisitante

    Set objRequisitante = obj1

    CodigoDe.Text = CStr(objRequisitante.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoAte_evSelecao(obj1 As Object)

Dim objRequisitante As New ClassRequisitante

    Set objRequisitante = obj1

    CodigoAte.Text = CStr(objRequisitante.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeDe_evSelecao(obj1 As Object)

Dim objRequisitante As New ClassRequisitante

    Set objRequisitante = obj1

    NomeDe.Text = objRequisitante.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeAte_evSelecao(obj1 As Object)

Dim objRequisitante As New ClassRequisitante

    Set objRequisitante = obj1

    NomeAte.Text = objRequisitante.sNomeReduzido

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoCclDe_evSelecao(obj1 As Object)
'traz o ccl selecionado para a tela

Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCclDe_evSelecao

    Set objCcl = obj1

    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 68554

    CclDe.PromptInclude = False
    CclDe.Text = sCclMascarado
    CclDe.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclDe_evSelecao:

    Select Case gErr

        Case 68554

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172953)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCclAte_evSelecao(obj1 As Object)
'traz o ccl selecionado para a tela

Dim lErro As Long
Dim objCcl As ClassCcl
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCclAte_evSelecao

    Set objCcl = obj1

    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 68555

    CclAte.PromptInclude = False
    CclAte.Text = sCclMascarado
    CclAte.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclAte_evSelecao:

    Select Case gErr

        Case 68555

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172954)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 68508

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 68509

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 68510
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 68511
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 68508
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 68509, 68510, 68511
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 172955)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 68512

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 68513

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 68512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 68513

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172956)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 68514

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "Codigo", 1)
                
            Case ORD_POR_NOME

                Call gobjRelOpcoes.IncluirOrdenacao(1, "Nome", 1)
                
            Case ORD_POR_CCL

                Call gobjRelOpcoes.IncluirOrdenacao(1, "Ccl", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "Nome", 1)
                
            Case Else
                gError 74971

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 68514, 74971

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172957)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCodigo_I As String
Dim sCodigo_F As String
Dim sNome_I As String
Dim sNome_F As String
Dim sCcl_I As String
Dim sCcl_F As String
Dim sOrdenacaoPor As String
Dim iOrdenacao As Long
Dim sOrd As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sCodigo_I, sCodigo_F, sNome_I, sNome_F, sCcl_I, sCcl_F)
    If lErro <> SUCESSO Then gError 68515

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 68516
         
    lErro = objRelOpcoes.IncluirParametro("NCODIGOINIC", sCodigo_I)
    If lErro <> AD_BOOL_TRUE Then gError 68517
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEINIC", NomeDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68518
    
    lErro = objRelOpcoes.IncluirParametro("TCCLINIC", sCcl_I)
    If lErro <> AD_BOOL_TRUE Then gError 68519
    
    lErro = objRelOpcoes.IncluirParametro("NCODIGOFIM", sCodigo_F)
    If lErro <> AD_BOOL_TRUE Then gError 68520
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEFIM", NomeAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68521
        
    lErro = objRelOpcoes.IncluirParametro("TCCLFIM", sCcl_F)
    If lErro <> AD_BOOL_TRUE Then gError 68522
        
    Select Case ComboOrdenacao.ListIndex
        
            Case ORD_POR_CODIGO
            
                sOrdenacaoPor = "CodReq"
                
            Case ORD_POR_NOME
                
                sOrdenacaoPor = "NomeReq"
                
            Case ORD_POR_CCL
                sOrdenacaoPor = "Ccl"
                
            Case Else
                gError 68523
                  
    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 68524
   
    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 68556
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodigo_I, sCodigo_F, sNome_I, sNome_F, sCcl_I, sCcl_F, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 68525

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 68515 To 68525, 68556
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172958)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodigo_I As String, sCodigo_F As String, sNome_I As String, sNome_F As String, sCcl_I As String, sCcl_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim sCclFormata As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'critica Codigo Inicial e Final
    If CodigoDe.Text <> "" Then
        sCodigo_I = CStr(CodigoDe.Text)
    Else
        sCodigo_I = ""
    End If
    
    If CodigoAte.Text <> "" Then
        sCodigo_F = CStr(CodigoAte.Text)
    Else
        sCodigo_F = ""
    End If
            
    If sCodigo_I <> "" And sCodigo_F <> "" Then
        
        If StrParaLong(sCodigo_I) > StrParaLong(sCodigo_F) Then gError 68526
        
    End If
    
    If NomeDe.Text <> "" Then
        sNome_I = NomeDe.Text
    Else
        sNome_I = ""
    End If
    
    If NomeAte.Text <> "" Then
        sNome_F = NomeAte.Text
    Else
        sNome_F = ""
    End If
    
    If sNome_I <> "" And sNome_F <> "" Then
        If sNome_I > sNome_F Then gError 68557
    End If
    
    
    'critica Ccl Inicial e Final
    If CclDe.ClipText <> "" Then
        lErro = CF("Ccl_Formata", CclDe.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 68559
        
        sCcl_I = sCclFormata
    Else
        sCcl_I = ""
    End If
    
    If CclAte.ClipText <> "" Then
        lErro = CF("Ccl_Formata", CclAte.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 68560
        
        sCcl_F = sCclFormata
    Else
        sCcl_F = ""
    End If
            
    If sCcl_I <> "" And sCcl_F <> "" Then
        
        If sCcl_I > sCcl_F Then gError 68527
        
    End If
    
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 68526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INICIAL_MAIOR", gErr)
            CodigoDe.SetFocus
                
        Case 68527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR", gErr)
            CclDe.SetFocus
            
        Case 68557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INICIAL_MAIOR", gErr)
            NomeDe.SetFocus
            
        Case 68559, 68560
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172959)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodigo_I As String, sCodigo_F As String, sNome_I As String, sNome_F As String, sCcl_I As String, sCcl_F As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sCodigo_I <> "" Then sExpressao = "Codigo >= " & Forprint_ConvLong(StrParaLong(sCodigo_I))

   If sCodigo_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Codigo <= " & Forprint_ConvLong(StrParaLong(sCodigo_F))

    End If

   If sNome_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Nome >= " & Forprint_ConvTexto(sNome_I)

    End If
    
    If sNome_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Nome <= " & Forprint_ConvTexto(sNome_F)

    End If
   
    If sCcl_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ccl >= " & Forprint_ConvTexto((sCcl_I))

    End If
   
    If sCcl_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ccl <= " & Forprint_ConvTexto((sCcl_F))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172960)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sTipoCliente As String, iTipo As Integer
Dim sOrdenacaoPor As String
Dim sCclMascarado As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 68528
   
    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODIGOINIC", sParam)
    If lErro <> SUCESSO Then gError 68529
    
    CodigoDe.Text = sParam
    Call CodigoDe_Validate(bSGECancelDummy)
    
    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODIGOFIM", sParam)
    If lErro <> SUCESSO Then gError 68530
    
    CodigoAte.Text = sParam
    Call CodigoAte_Validate(bSGECancelDummy)
                
    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEINIC", sParam)
    If lErro <> SUCESSO Then gError 68531
                   
    NomeDe.Text = sParam
    Call NomeDe_Validate(bSGECancelDummy)
    
    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFIM", sParam)
    If lErro <> SUCESSO Then gError 68532
                   
    NomeAte.Text = sParam
    Call NomeAte_Validate(bSGECancelDummy)
                        
    'pega  Ccl Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLINIC", sParam)
    If lErro <> SUCESSO Then gError 68533
                   
    If Len(Trim(sParam)) > 0 Then
        lErro = Mascara_MascararCcl(sParam, sCclMascarado)
        If lErro <> SUCESSO Then gError 68561
        CclDe.PromptInclude = False
        CclDe.Text = sCclMascarado
        CclDe.PromptInclude = True
        
    End If
    Call CclDe_Validate(bSGECancelDummy)
                          
                          
    'pega  Ccl Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLFIM", sParam)
    If lErro <> SUCESSO Then gError 68534
                   
    If Len(Trim(sParam)) > 0 Then
    
        lErro = Mascara_MascararCcl(sParam, sCclMascarado)
        If lErro <> SUCESSO Then gError 68562
        
        CclAte.PromptInclude = False
        CclAte.Text = sCclMascarado
        CclAte.PromptInclude = True
        
    End If
    Call CclAte_Validate(bSGECancelDummy)
                              
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 68535
    
    
    Select Case sOrdenacaoPor
        
            Case "CodReq"
            
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
            Case "NomeReq"
            
                ComboOrdenacao.ListIndex = ORD_POR_NOME
                
            Case "Ccl"
            
                ComboOrdenacao.ListIndex = ORD_POR_CCL
                
            Case Else
                gError 68536
                  
    End Select
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 68528 To 68536, 68561, 68562
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172961)

    End Select

    Exit Function

End Function

Private Sub CodigoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_CodigoDe_Validate

    If Len(Trim(CodigoDe.Text)) > 0 Then

        objRequisitante.lCodigo = StrParaLong(CodigoDe.Text)
        'Lê o código informado
        lErro = CF("Requisitante_Le", objRequisitante)
        If lErro <> SUCESSO And lErro <> 49084 Then gError 68538

        'Se não encontrou o Requisitante ==> erro
        If lErro = 49084 Then gError 68539
        
    End If

    Exit Sub

Erro_CodigoDe_Validate:

    Cancel = True


    Select Case gErr

        Case 68538

        Case 68539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172962)

    End Select

    Exit Sub
    
End Sub


Private Sub CodigoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_CodigoAte_Validate

    If Len(Trim(CodigoAte.Text)) > 0 Then

        objRequisitante.lCodigo = StrParaLong(CodigoAte.Text)
        'Lê o código informado
        lErro = CF("Requisitante_Le", objRequisitante)
        If lErro <> SUCESSO And lErro <> 49084 Then gError 68540

        'Se não encontrou o Requisitante ==> erro
        If lErro = 49084 Then gError 68541
        
    End If

    Exit Sub

Erro_CodigoAte_Validate:

    Cancel = True


    Select Case gErr

        Case 68540

        Case 68541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172963)

    End Select

Exit Sub

End Sub
Private Sub NomeDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_NomeDe_Validate

    If Len(Trim(NomeDe.Text)) > 0 Then

        objRequisitante.sNomeReduzido = NomeDe.Text
        'Lê o Requisitante informado
        lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
        If lErro <> SUCESSO And lErro <> 51152 Then gError 68541

        'Se não encontrou o Requisitante ==> erro
        If lErro = 51152 Then gError 68542
        
        NomeDe.Text = objRequisitante.sNomeReduzido
        
    End If

    Exit Sub

Erro_NomeDe_Validate:

    Cancel = True


    Select Case gErr

        Case 68541

        Case 68542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172964)

    End Select

Exit Sub

End Sub

Private Sub NomeAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRequisitante As New ClassRequisitante

On Error GoTo Erro_NomeAte_Validate

    If Len(Trim(NomeAte.Text)) > 0 Then

        objRequisitante.sNomeReduzido = NomeAte.Text
        'Lê o Requisitante informado
        lErro = CF("Requisitante_Le_NomeReduzido", objRequisitante)
        If lErro <> SUCESSO And lErro <> 51152 Then gError 68543

        'Se não encontrou o Requisitante ==> erro
        If lErro = 51152 Then gError 68544
        
        NomeAte.Text = objRequisitante.sNomeReduzido
        
    End If

    Exit Sub

Erro_NomeAte_Validate:

    Cancel = True


    Select Case gErr

        Case 68543

        Case 68544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172965)

    End Select

Exit Sub

End Sub

Private Sub CclDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_CclDe_Validate

    If Len(Trim(CclDe.ClipText)) > 0 Then

        'Coloca Ccl no formato do BD
        lErro = CF("Ccl_Formata", CclDe.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 68545
        
        objCcl.sCcl = sCclFormata
        
        'Lê o Ccl informado
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 68546

        'Se não encontrou o Ccl ==> erro
        If lErro = 5599 Then gError 68547
            
    End If

    Exit Sub

Erro_CclDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68545, 68546

        Case 68547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172966)

    End Select

Exit Sub

End Sub

Private Sub CclAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormata As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_CclAte_Validate

    If Len(Trim(CclAte.ClipText)) > 0 Then

        'Coloca Ccl no formato do BD
        lErro = CF("Ccl_Formata", CclAte.Text, sCclFormata, iCclPreenchida)
        If lErro <> SUCESSO Then gError 68548
        
        objCcl.sCcl = sCclFormata
        
        'Lê o Ccl informado
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 68549

        'Se não encontrou o Ccl ==> erro
        If lErro = 5599 Then gError 68550
        
    End If

    Exit Sub

Erro_CclAte_Validate:

    Cancel = True

    Select Case gErr

        Case 68548, 68549

        Case 68550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172967)

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
    Caption = "Requisitantes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRequisitantes"
    
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
        
        If Me.ActiveControl Is CodigoDe Then
            Call LabelCodigoDe_Click
            
        ElseIf Me.ActiveControl Is CodigoAte Then
            Call LabelCodigoAte_Click
            
        ElseIf Me.ActiveControl Is NomeDe Then
            Call LabelNomeDe_Click
            
        ElseIf Me.ActiveControl Is NomeAte Then
            Call LabelNomeAte_Click
            
        ElseIf Me.ActiveControl Is CclDe Then
            Call LabelCclDe_Click
            
        ElseIf Me.ActiveControl Is CclAte Then
            Call LabelCclAte_Click
            
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



Private Sub LabelCclAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCClAte, Source, X, Y)
End Sub

Private Sub LabelCclAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCClAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCclDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCclDE, Source, X, Y)
End Sub

Private Sub LabelCclDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCclDE, Button, Shift, X, Y)
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

