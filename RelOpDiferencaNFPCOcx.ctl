VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpDiferencaNFPCOcx 
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   ScaleHeight     =   4155
   ScaleWidth      =   7095
   Begin VB.Frame Frame1 
      Caption         =   "Notas Fiscais"
      Height          =   1515
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   2460
      Width           =   4875
      Begin VB.Frame Frame4 
         Caption         =   "Data de Entrega"
         Height          =   1125
         Left            =   2280
         TabIndex        =   21
         Top             =   210
         Width           =   2295
         Begin MSComCtl2.UpDown UpDownDataEntregaAte 
            Height          =   315
            Left            =   1830
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   720
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEntregaAte 
            Height          =   315
            Left            =   660
            TabIndex        =   8
            Top             =   720
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataEntregaDe 
            Height          =   315
            Left            =   1830
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   220
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEntregaDe 
            Height          =   315
            Left            =   660
            TabIndex        =   7
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
         Begin VB.Label LabelDataLimiteDe 
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
            TabIndex        =   25
            Top             =   320
            Width           =   315
         End
         Begin VB.Label LabelDataLimiteAte 
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
            Left            =   150
            TabIndex        =   24
            Top             =   820
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Número"
         Height          =   1125
         Left            =   300
         TabIndex        =   18
         Top             =   210
         Width           =   1875
         Begin MSMask.MaskEdBox NumeroDe 
            Height          =   300
            Left            =   660
            TabIndex        =   5
            Top             =   225
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumeroAte 
            Height          =   300
            Left            =   660
            TabIndex        =   6
            Top             =   720
            Width           =   1155
            _ExtentX        =   2037
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
            Left            =   240
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   20
            Top             =   820
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
            Left            =   240
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   19
            Top             =   320
            Width           =   315
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4770
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDiferencaNFPCOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDiferencaNFPCOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDiferencaNFPCOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDiferencaNFPCOcx.ctx":0816
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
      Left            =   5310
      Picture         =   "RelOpDiferencaNFPCOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1020
      Width           =   1605
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDiferencaNFPCOcx.ctx":0A96
      Left            =   900
      List            =   "RelOpDiferencaNFPCOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   2730
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pedidos de Compra"
      Height          =   1455
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   4875
      Begin VB.Frame Frame6 
         Caption         =   "Data de Envio"
         Height          =   1125
         Left            =   2280
         TabIndex        =   29
         Top             =   210
         Width           =   2295
         Begin MSComCtl2.UpDown UpDownDataEnvioAte 
            Height          =   315
            Left            =   1830
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   720
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEnvioAte 
            Height          =   315
            Left            =   660
            TabIndex        =   4
            Top             =   720
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataEnvioDe 
            Height          =   315
            Left            =   1830
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   220
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEnvioDe 
            Height          =   315
            Left            =   660
            TabIndex        =   3
            Top             =   220
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDataEnvioAte 
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
            Left            =   240
            TabIndex        =   33
            Top             =   820
            Width           =   360
         End
         Begin VB.Label LabelDataEnvioDe 
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
            Left            =   240
            TabIndex        =   32
            Top             =   315
            Width           =   315
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Código"
         Height          =   1125
         Left            =   300
         TabIndex        =   26
         Top             =   210
         Width           =   1875
         Begin MSMask.MaskEdBox CodigoPCDe 
            Height          =   300
            Left            =   660
            TabIndex        =   1
            Top             =   220
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoPCAte 
            Height          =   300
            Left            =   660
            TabIndex        =   2
            Top             =   720
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodigoPCDe 
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
            Left            =   240
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   320
            Width           =   315
         End
         Begin VB.Label LabelCodigoPCAte 
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
            Left            =   240
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   27
            Top             =   820
            Width           =   360
         End
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   225
      TabIndex        =   17
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "RelOpDiferencaNFPCOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCodPCDe As AdmEvento
Attribute objEventoCodPCDe.VB_VarHelpID = -1
Private WithEvents objEventoCodPCAte As AdmEvento
Attribute objEventoCodPCAte.VB_VarHelpID = -1
Private WithEvents objEventoNumNFDe As AdmEvento
Attribute objEventoNumNFDe.VB_VarHelpID = -1
Private WithEvents objEventoNumNFAte As AdmEvento
Attribute objEventoNumNFAte.VB_VarHelpID = -1

Dim iAlterado As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 74495

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 74496

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 74495

        Case 74496
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168337)

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
    If lErro <> SUCESSO Then gError 74497

    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus

    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 74497

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168338)

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
    Set objEventoNumNFDe = New AdmEvento
    Set objEventoNumNFAte = New AdmEvento
    

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168339)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoCodPCDe = Nothing
    Set objEventoCodPCAte = Nothing
    Set objEventoNumNFDe = Nothing
    Set objEventoNumNFAte = Nothing
    
End Sub

Private Sub CodigoPCAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoPCAte, iAlterado)
    
End Sub

Private Sub CodigoPCDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoPCDe, iAlterado)
    
End Sub

Private Sub DataEntregaAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEntregaAte, iAlterado)
    
End Sub

Private Sub DataEntregaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntregaAte_Validate

    'Verifica se a DataEntregaAte está preenchida
    If Len(Trim(DataEntregaAte.Text)) = 0 Then Exit Sub

    'Critica a DataEntregaAte informada
    lErro = Data_Critica(DataEntregaAte.Text)
    If lErro <> SUCESSO Then gError 74533

    Exit Sub

Erro_DataEntregaAte_Validate:

    Cancel = True

    Select Case gErr

        Case 74533
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168340)

    End Select

    Exit Sub

End Sub


Private Sub DataEntregaDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEntregaDe, iAlterado)
    
End Sub

Private Sub DataEnvioAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvioAte, iAlterado)
    
End Sub

Private Sub DataEnvioDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvioDe, iAlterado)
    
End Sub

Private Sub NumeroAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumeroAte, iAlterado)
    
End Sub

Private Sub NumeroDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumeroDe, iAlterado)
    
End Sub

Private Sub UpDownDataEntregaAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntregaAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEntregaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 74534

    Exit Sub

Erro_UpDownDataEntregaAte_DownClick:

    Select Case gErr

        Case 74534
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168341)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEntregaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntregaAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEntregaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 74535

    Exit Sub

Erro_UpDownDataEntregaAte_UpClick:

    Select Case gErr

        Case 74535
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168342)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEntregaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntregaDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEntregaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 74536

    Exit Sub

Erro_UpDownDataEntregaDe_DownClick:

    Select Case gErr

        Case 74536
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168343)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEntregaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntregaDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEntregaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 74537

    Exit Sub

Erro_UpDownDataEntregaDe_UpClick:

    Select Case gErr

        Case 74537
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168344)

    End Select

    Exit Sub

End Sub

Private Sub DataEntregaDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntregaDe_Validate

    'Verifica se a DataEntregaDe está preenchida
    If Len(Trim(DataEntregaDe.Text)) = 0 Then Exit Sub

    'Critica a DataEntregaDe informada
    lErro = Data_Critica(DataEntregaDe.Text)
    If lErro <> SUCESSO Then gError 74532

    Exit Sub

Erro_DataEntregaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74532
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168345)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168346)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168347)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioDe_Validate

    'Verifica se a DataEnvioDe está preenchida
    If Len(Trim(DataEnvioDe.Text)) = 0 Then Exit Sub

    'Critica a DataEnvioDe informada
    lErro = Data_Critica(DataEnvioDe.Text)
    If lErro <> SUCESSO Then gError 74498

    Exit Sub

Erro_DataEnvioDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74498
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168348)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioAte_Validate

    'Verifica se a DataEnvioDe está preenchida
    If Len(Trim(DataEnvioAte.Text)) = 0 Then Exit Sub

    'Critica a DataEnvioDe informada
    lErro = Data_Critica(DataEnvioAte.Text)
    If lErro <> SUCESSO Then gError 74499

    Exit Sub

Erro_DataEnvioAte_Validate:

    Cancel = True

    Select Case gErr

        Case 74499
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168349)

    End Select

    Exit Sub

End Sub



Private Sub UpDownDataEnvioAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 74500

    Exit Sub

Erro_UpDownDataEnvioAte_DownClick:

    Select Case gErr

        Case 74500
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168350)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 74501

    Exit Sub

Erro_UpDownDataEnvioAte_UpClick:

    Select Case gErr

        Case 74501
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168351)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 74502

    Exit Sub

Erro_UpDownDataEnvioDe_DownClick:

    Select Case gErr

        Case 74502
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168352)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 74503

    Exit Sub

Erro_UpDownDataEnvioDe_UpClick:

    Select Case gErr

        Case 74503
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 168353)

    End Select

    Exit Sub

End Sub


Private Sub LabelCodigoPCDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompra As New ClassPedidoCompras

On Error GoTo Erro_LabelCodigoPCDe_Click

    If Len(Trim(CodigoPCDe.Text)) > 0 Then
        
        objPedCompra.lCodigo = StrParaLong(CodigoPCDe.Text)
    End If

    'Chama Tela PedComprasTodosLista
    Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedCompra, objEventoCodPCDe)

   Exit Sub

Erro_LabelCodigoPCDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168354)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodigoPCAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompra As New ClassPedidoCompras

On Error GoTo Erro_LabelCodigoPCAte_Click

    If Len(Trim(CodigoPCAte.Text)) > 0 Then
        
        objPedCompra.lCodigo = StrParaLong(CodigoPCAte.Text)
    End If

    'Chama Tela PedComprasTodosLista
    Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedCompra, objEventoCodPCAte)

   Exit Sub

Erro_LabelCodigoPCAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168355)

    End Select

    Exit Sub

End Sub


Private Sub objEventoCodPCAte_evSelecao(obj1 As Object)

Dim objPedCompra As New ClassPedidoCompras

    Set objPedCompra = obj1

    CodigoPCAte.Text = CStr(objPedCompra.lCodigo)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoCodPCDe_evSelecao(obj1 As Object)

Dim objPedCompra As New ClassPedidoCompras

    Set objPedCompra = obj1

    CodigoPCDe.Text = CStr(objPedCompra.lCodigo)

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
    If ComboOpcoes.Text = "" Then gError 74504

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 74505

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 74506

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 74507

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 74504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 74505 To 74507

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168356)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 74508

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 74509

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 74508
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 74509

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168357)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 74510

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 74510

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168358)

    End Select

    Exit Sub

End Sub


Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCodPC_I As String
Dim sCodPC_F As String
Dim sNumero_I As String
Dim sNumero_F As String
Dim sCheck As String

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sCodPC_I, sCodPC_F, sNumero_I, sNumero_F)
    If lErro <> SUCESSO Then gError 74511

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 74512

    lErro = objRelOpcoes.IncluirParametro("NCODPCINIC", sCodPC_I)
    If lErro <> AD_BOOL_TRUE Then gError 74513

    lErro = objRelOpcoes.IncluirParametro("NNOTAFISCALINIC", sNumero_I)
    If lErro <> AD_BOOL_TRUE Then gError 74514

    'Preenche data inicial
    If Trim(DataEnvioDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", DataEnvioDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 74515

    'Preenche data entrega inicial
    If Trim(DataEntregaDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", DataEntregaDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 74520
    lErro = objRelOpcoes.IncluirParametro("NCODPCFIM", sCodPC_F)
    If lErro <> AD_BOOL_TRUE Then gError 74516

    lErro = objRelOpcoes.IncluirParametro("NNOTAFISCALFIM", sNumero_F)
    If lErro <> AD_BOOL_TRUE Then gError 74517

    'Preenche data final
    If Trim(DataEnvioAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", DataEnvioAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 74518

    'Preenche data entrega final
    If Trim(DataEntregaAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", DataEntregaAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 74519

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodPC_I, sCodPC_F, sNumero_I, sNumero_F)
    If lErro <> SUCESSO Then gError 73559

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 74511 To 74520

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168359)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodPC_I As String, sCodPC_F As String, sNumero_I As String, sNumero_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'critica Codigo PC Inicial e Final
    If CodigoPCDe.Text <> "" Then
        sCodPC_I = CStr(CodigoPCDe.Text)
    Else
        sCodPC_I = ""
    End If

    If CodigoPCAte.Text <> "" Then
        sCodPC_F = CStr(CodigoPCAte.Text)
    Else
        sCodPC_F = ""
    End If

    If sCodPC_I <> "" And sCodPC_F <> "" Then

        If StrParaLong(sCodPC_I) > StrParaLong(sCodPC_F) Then gError 74521

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

        If StrParaLong(sNumero_I) > StrParaLong(sNumero_F) Then gError 74522

    End If

    'data entrega inicial não pode ser maior que a final
    If Trim(DataEntregaDe.ClipText) <> "" And Trim(DataEntregaAte.ClipText) <> "" Then
    
         If CDate(DataEntregaDe.Text) > CDate(DataEntregaAte.Text) Then gError 74523
    
    End If
    
    'data envio inicial não pode ser maior que a final
    If Trim(DataEnvioDe.ClipText) <> "" And Trim(DataEnvioAte.ClipText) <> "" Then
    
         If CDate(DataEnvioDe.Text) > CDate(DataEnvioAte.Text) Then gError 74524
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 74521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PC_INICIAL_MAIOR", gErr)
            CodigoPCDe.SetFocus

        Case 74522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMNF_INICIAL_MAIOR", gErr)
            NumeroDe.SetFocus

        Case 74523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataEntregaDe.SetFocus

        Case 74524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataEnvioDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168360)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodPC_I As String, sCodPC_F As String, sNumero_I As String, sNumero_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao


   If sCodPC_I <> "" Then sExpressao = "S01"

   If sCodPC_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S02"

    End If

    If sNumero_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S03"

    End If

    If sNumero_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S04"

    End If
    
    If Trim(DataEntregaDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S05"

    End If

    If Trim(DataEntregaAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S06"

    End If
    
    If Trim(DataEnvioDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S07"

    End If

    If Trim(DataEnvioAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "S08"

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168361)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 74523

    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCINIC", sParam)
    If lErro <> SUCESSO Then gError 74524

    CodigoPCDe.Text = sParam

    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCFIM", sParam)
    If lErro <> SUCESSO Then gError 74525

    CodigoPCAte.Text = sParam

    'pega  Numero inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NNOTAFISCALINIC", sParam)
    If lErro <> SUCESSO Then gError 74526

    NumeroDe.Text = sParam

    'pega numero final e exibe
    lErro = objRelOpcoes.ObterParametro("NNOTAFISCALFIM", sParam)
    If lErro <> SUCESSO Then gError 74527

    NumeroAte.Text = sParam

    'pega data entrega inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINIC", sParam)
    If lErro <> SUCESSO Then gError 74528

    Call DateParaMasked(DataEntregaDe, CDate(sParam))

    'pega data entrega final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 74529

    Call DateParaMasked(DataEntregaAte, CDate(sParam))

    'pega data envio inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DENVINIC", sParam)
    If lErro <> SUCESSO Then gError 74530

    Call DateParaMasked(DataEnvioDe, CDate(sParam))

    'pega data envio final e exibe
    lErro = objRelOpcoes.ObterParametro("DENVFIM", sParam)
    If lErro <> SUCESSO Then gError 74531

    Call DateParaMasked(DataEnvioAte, CDate(sParam))

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 74523 To 74531

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168362)

    End Select

    Exit Function

End Function


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
    Caption = "Relação de Diferença entre Pedidos de Compra"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpDiferencaNFPC"

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

        ElseIf Me.ActiveControl Is CodigoPCDe Then
            Call LabelCodigoPCDe_Click

        ElseIf Me.ActiveControl Is CodigoPCAte Then
            Call LabelCodigoPCAte_Click

        End If

    End If

End Sub




Private Sub LabelNumeroDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumeroDe, Source, X, Y)
End Sub

Private Sub LabelNumeroDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumeroDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNumeroAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumeroAte, Source, X, Y)
End Sub

Private Sub LabelNumeroAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumeroAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDataLimiteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataLimiteAte, Source, X, Y)
End Sub

Private Sub LabelDataLimiteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataLimiteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDataLimiteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataLimiteDe, Source, X, Y)
End Sub

Private Sub LabelDataLimiteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataLimiteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDataEnvioDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataEnvioDe, Source, X, Y)
End Sub

Private Sub LabelDataEnvioDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataEnvioDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDataEnvioAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataEnvioAte, Source, X, Y)
End Sub

Private Sub LabelDataEnvioAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataEnvioAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoPCAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoPCAte, Source, X, Y)
End Sub

Private Sub LabelCodigoPCAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoPCAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoPCDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoPCDe, Source, X, Y)
End Sub

Private Sub LabelCodigoPCDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoPCDe, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

