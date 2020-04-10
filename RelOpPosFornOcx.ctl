VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPosFornOcx 
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3990
   ScaleWidth      =   6645
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPosFornOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPosFornOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPosFornOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPosFornOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fornecedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   120
      TabIndex        =   23
      Top             =   2565
      Width           =   3990
      Begin MSMask.MaskEdBox FornecedorInicial 
         Height          =   300
         Left            =   750
         TabIndex        =   5
         Top             =   315
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox FornecedorFinal 
         Height          =   300
         Left            =   765
         TabIndex        =   6
         Top             =   780
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label LabelFornInic 
         AutoSize        =   -1  'True
         Caption         =   "Inicial:"
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
         Left            =   105
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   375
         Width           =   585
      End
      Begin VB.Label LabelFornFim 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   825
         Width           =   480
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPosFornOcx.ctx":0994
      Left            =   855
      List            =   "RelOpPosFornOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3060
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
      Left            =   4500
      Picture         =   "RelOpPosFornOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   810
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Emissão"
      Height          =   810
      Left            =   120
      TabIndex        =   18
      Top             =   795
      Width           =   3990
      Begin MSComCtl2.UpDown UpDownEmissaoDe 
         Height          =   315
         Left            =   1665
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoDe 
         Height          =   315
         Left            =   525
         TabIndex        =   1
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissaoAte 
         Height          =   315
         Left            =   3585
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoAte 
         Height          =   315
         Left            =   2445
         TabIndex        =   2
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
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
         TabIndex        =   22
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label2 
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
         Left            =   2055
         TabIndex        =   21
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vencimento"
      Height          =   810
      Left            =   120
      TabIndex        =   13
      Top             =   1695
      Width           =   3990
      Begin MSComCtl2.UpDown UpDownVenctoDe 
         Height          =   315
         Left            =   1665
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VenctoDe 
         Height          =   315
         Left            =   510
         TabIndex        =   3
         Top             =   300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownVenctoAte 
         Height          =   315
         Left            =   3585
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VenctoAte 
         Height          =   315
         Left            =   2445
         TabIndex        =   4
         Top             =   300
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
         Left            =   2055
         TabIndex        =   17
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label7 
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
         TabIndex        =   16
         Top             =   360
         Width           =   315
      End
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
      Height          =   255
      Left            =   135
      TabIndex        =   26
      Top             =   285
      Width           =   690
   End
End
Attribute VB_Name = "RelOpPosFornOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoFornecedorInic As AdmEvento
Attribute objEventoFornecedorInic.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorFim As AdmEvento
Attribute objEventoFornecedorFim.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 48765

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 48766

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 48767
    
        ComboOpcoes.Text = ""
                
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 48765
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 48766, 48767

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171288)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 48768

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 48769
    
    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 48770
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 48771
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 48768
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 48769, 48770, 48771, 48772
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171289)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoFornecedorInic = New AdmEvento
    Set objEventoFornecedorFim = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171290)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 48774
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 48773
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 48773
        
        Case 48774
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171291)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 48775
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 48775
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171292)

    End Select

    Exit Sub
   
End Sub

Private Sub EmissaoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(EmissaoAte)

End Sub

Private Sub EmissaoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(EmissaoDe)

End Sub

Private Sub FornecedorFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorFinal_Validate

    If Len(Trim(FornecedorFinal.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(FornecedorFinal, objFornecedor, 0)
        If lErro <> SUCESSO Then Error 48776
        
    End If
    
    Exit Sub

Erro_FornecedorFinal_Validate:

    Cancel = True


    Select Case Err

        Case 48776

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171293)

    End Select

End Sub

Private Sub FornecedorInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorInicial_Validate

    If Len(Trim(FornecedorInicial.Text)) > 0 Then
   
        'Tenta ler o Fornecedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(FornecedorInicial, objFornecedor, 0)
        If lErro <> SUCESSO Then Error 48777

    End If
        
    Exit Sub

Erro_FornecedorInicial_Validate:

    Cancel = True


    Select Case Err

        Case 48777

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171294)

    End Select

End Sub

Private Sub LabelFornInic_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornInic_Click
    
    If Len(Trim(FornecedorInicial.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = LCodigo_Extrai(FornecedorInicial.Text)
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorInic)

   Exit Sub

Erro_LabelFornInic_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171295)

    End Select

    Exit Sub

End Sub

Private Sub LabelFornFim_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornFim_Click
    
    If Len(Trim(FornecedorFinal.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = LCodigo_Extrai(FornecedorFinal.Text)
    End If
    
    'Chama Tela FornecedorsLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorFim)

   Exit Sub

Erro_LabelFornFim_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171296)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoFornecedorFim_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    FornecedorFinal.Text = CStr(objFornecedor.lCodigo)
    Call FornecedorFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFornecedorInic_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    FornecedorInicial.Text = CStr(objFornecedor.lCodigo)
    Call FornecedorInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoFornecedorFim = Nothing
    Set objEventoFornecedorInic = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Private Sub EmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoAte_Validate

    If Len(EmissaoAte.ClipText) > 0 Then
        
        lErro = Data_Critica(EmissaoAte.Text)
        If lErro <> SUCESSO Then Error 48778

    End If

    Exit Sub

Erro_EmissaoAte_Validate:

    Cancel = True


    Select Case Err

        Case 48778

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171297)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoDe_Validate

    If Len(EmissaoDe.ClipText) > 0 Then

        lErro = Data_Critica(EmissaoDe.Text)
        If lErro <> SUCESSO Then Error 48779

    End If

    Exit Sub

Erro_EmissaoDe_Validate:

    Cancel = True


    Select Case Err

        Case 48779

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171298)

    End Select

    Exit Sub

End Sub

Private Sub VenctoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VenctoAte)

End Sub

Private Sub VenctoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VenctoAte_Validate

    If Len(VenctoAte.ClipText) > 0 Then
        
        lErro = Data_Critica(VenctoAte.Text)
        If lErro <> SUCESSO Then Error 48780

    End If

    Exit Sub

Erro_VenctoAte_Validate:

    Cancel = True


    Select Case Err

        Case 48780

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171299)

    End Select

    Exit Sub

End Sub

Private Sub VenctoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VenctoDe)

End Sub

Private Sub VenctoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VenctoDe_Validate

    If Len(VenctoDe.ClipText) > 0 Then

        lErro = Data_Critica(VenctoDe.Text)
        If lErro <> SUCESSO Then Error 48781

    End If

    Exit Sub

Erro_VenctoDe_Validate:

    Cancel = True


    Select Case Err

        Case 48781

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171300)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVenctoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVenctoAte_DownClick

    lErro = Data_Up_Down_Click(VenctoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 48782

    Exit Sub

Erro_UpDownVenctoAte_DownClick:

    Select Case Err

        Case 48782
            VenctoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171301)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVenctoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVenctoAte_UpClick

    lErro = Data_Up_Down_Click(VenctoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 48783

    Exit Sub

Erro_UpDownVenctoAte_UpClick:

    Select Case Err

        Case 48783
            VenctoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171302)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    lErro = Data_Up_Down_Click(EmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 48784

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case Err

        Case 48784
            EmissaoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171303)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    lErro = Data_Up_Down_Click(EmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 48785

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case Err

        Case 48785
            EmissaoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171304)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 48786

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 48786

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171305)

    End Select

    Exit Sub

End Sub

Private Function Formata_E_Critica_Parametros(sFornecedor_I As String, sFornecedor_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'critica Fornecedor Inicial e Final
    If FornecedorInicial.Text <> "" Then
        sFornecedor_I = CStr(LCodigo_Extrai(FornecedorInicial.Text))
    Else
        sFornecedor_I = ""
    End If
    
    If FornecedorFinal.Text <> "" Then
        sFornecedor_F = CStr(LCodigo_Extrai(FornecedorFinal.Text))
    Else
        sFornecedor_F = ""
    End If
            
    If sFornecedor_I <> "" And sFornecedor_F <> "" Then
        
        If CLng(sFornecedor_I) > CLng(sFornecedor_F) Then Error 48787
        
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(EmissaoDe.ClipText) <> "" And Trim(EmissaoAte.ClipText) <> "" Then
    
        If CDate(EmissaoDe.Text) > CDate(EmissaoAte.Text) Then Error 48809
    
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(VenctoDe.ClipText) <> "" And Trim(VenctoAte.ClipText) <> "" Then
    
        If CDate(VenctoDe.Text) > CDate(VenctoAte.Text) Then Error 48810
    
    End If
            
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
                
        Case 48787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", Err)
            FornecedorInicial.SetFocus
                
        Case 48809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_INICIAL_MAIOR", Err)
            EmissaoDe.SetFocus
            
        Case 48810
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VENCTO_INICIAL_MAIOR", Err)
            VenctoDe.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171306)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sFornecedor_I As String, sFornecedor_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    If sFornecedor_I <> "" Then sExpressao = "Forn >= " & Forprint_ConvLong(CLng(sFornecedor_I))

    If sFornecedor_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Forn <= " & Forprint_ConvLong(CLng(sFornecedor_F))

    End If
           
    If Trim(EmissaoDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao >= " & Forprint_ConvData(CDate(EmissaoDe.Text))

    End If
    
    If Trim(EmissaoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao <= " & Forprint_ConvData(CDate(EmissaoAte.Text))

    End If
           
    If Trim(VenctoDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataVencimento >= " & Forprint_ConvData(CDate(VenctoDe.Text))

    End If
    
    If Trim(VenctoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataVencimento <= " & Forprint_ConvData(CDate(VenctoAte.Text))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171307)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 48788
   
    'pega Fornecedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NFORNINIC", sParam)
    If lErro <> SUCESSO Then Error 48789
    
    FornecedorInicial.Text = sParam
    Call FornecedorInicial_Validate(bSGECancelDummy)
    
    'pega  Fornecedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NFORNFIM", sParam)
    If lErro <> SUCESSO Then Error 48790
    
    FornecedorFinal.Text = sParam
    Call FornecedorFinal_Validate(bSGECancelDummy)
                
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DEMINIC", sParam)
    If lErro <> SUCESSO Then Error 48791

    Call DateParaMasked(EmissaoDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DEMFIM", sParam)
    If lErro <> SUCESSO Then Error 48792

    Call DateParaMasked(EmissaoAte, CDate(sParam))
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DVENCINIC", sParam)
    If lErro <> SUCESSO Then Error 48793

    Call DateParaMasked(VenctoDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DVENCFIM", sParam)
    If lErro <> SUCESSO Then Error 48794

    Call DateParaMasked(VenctoAte, CDate(sParam))
            
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 48788, 48789, 48790, 48791, 48792, 48793, 48794
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171308)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sFornecedor_I As String
Dim sFornecedor_F As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sFornecedor_I, sFornecedor_F)
    If lErro <> SUCESSO Then Error 48795

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 48796
         
    lErro = objRelOpcoes.IncluirParametro("NFORNINIC", sFornecedor_I)
    If lErro <> AD_BOOL_TRUE Then Error 48797
    
    lErro = objRelOpcoes.IncluirParametro("TFORNINIC", FornecedorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54744
    
    lErro = objRelOpcoes.IncluirParametro("NFORNFIM", sFornecedor_F)
    If lErro <> AD_BOOL_TRUE Then Error 48798
    
    lErro = objRelOpcoes.IncluirParametro("TFORNFIM", FornecedorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54745
                   
    If Trim(EmissaoDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMINIC", EmissaoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 48799
    
    If Trim(EmissaoAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMFIM", EmissaoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 48800
    
    If Trim(VenctoDe.ClipText) <> "" Then
        'Preenche vencimento Inicial
        lErro = objRelOpcoes.IncluirParametro("DVENCINIC", VenctoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVENCINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 48801
    
    If Trim(VenctoAte.ClipText) <> "" Then
        'Preenche Vencimento Final
        lErro = objRelOpcoes.IncluirParametro("DVENCFIM", VenctoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVENCFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 48802

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sFornecedor_I, sFornecedor_F)
    If lErro <> SUCESSO Then Error 48803

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 48795, 48796, 48797, 48798, 48799, 48800, 48801, 48802, 48803
        
        Case 54744, 54745
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171309)

    End Select

    Exit Function

End Function

Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    lErro = Data_Up_Down_Click(EmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 48850

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case Err

        Case 48850
            EmissaoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171310)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    lErro = Data_Up_Down_Click(EmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 48851

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case Err

        Case 48851
            EmissaoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171311)

    End Select

    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_POSFORN
    Set Form_Load_Ocx = Me
    Caption = "Posição de Fornecedores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPosForn"
    
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
        
        If Me.ActiveControl Is FornecedorInicial Then
            Call LabelFornInic_Click
        ElseIf Me.ActiveControl Is FornecedorFinal Then
            Call LabelFornFim_Click
        End If
    
    End If

End Sub


Private Sub LabelFornInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornInic, Source, X, Y)
End Sub

Private Sub LabelFornInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornInic, Button, Shift, X, Y)
End Sub

Private Sub LabelFornFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornFim, Source, X, Y)
End Sub

Private Sub LabelFornFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornFim, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

