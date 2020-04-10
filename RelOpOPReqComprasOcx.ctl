VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpOPReqComprasOcx 
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   ScaleHeight     =   4260
   ScaleWidth      =   7950
   Begin VB.Frame Frame1 
      Caption         =   "Ordens de Produção"
      Height          =   3270
      Left            =   180
      TabIndex        =   17
      Top             =   855
      Width           =   5565
      Begin VB.CheckBox CheckOPAtendidas 
         Caption         =   "Inclui Ordens de Produção Atendidas"
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
         Left            =   225
         TabIndex        =   8
         Top             =   2850
         Width           =   4005
      End
      Begin VB.Frame FrameCodigo 
         Caption         =   "Filial Empresa"
         Height          =   825
         Left            =   180
         TabIndex        =   24
         Top             =   270
         Width           =   5160
         Begin MSMask.MaskEdBox FilialDe 
            Height          =   300
            Left            =   525
            TabIndex        =   2
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialAte 
            Height          =   300
            Left            =   3105
            TabIndex        =   3
            Top             =   345
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label LabelEmpresaDe 
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
            TabIndex        =   26
            Top             =   405
            Width           =   315
         End
         Begin VB.Label LabelEmpresaAte 
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
            TabIndex        =   25
            Top             =   405
            Width           =   360
         End
      End
      Begin VB.Frame FrameNome 
         Caption         =   "Código"
         Height          =   765
         Left            =   180
         TabIndex        =   21
         Top             =   1185
         Width           =   5160
         Begin MSMask.MaskEdBox CodigoDe 
            Height          =   300
            Left            =   525
            TabIndex        =   4
            Top             =   285
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoAte 
            Height          =   300
            Left            =   3045
            TabIndex        =   5
            Top             =   285
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
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
            Left            =   2625
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   23
            Top             =   338
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   22
            Top             =   338
            Width           =   315
         End
      End
      Begin VB.Frame FrameCcl 
         Caption         =   "Data de Emissão"
         Height          =   675
         Left            =   180
         TabIndex        =   18
         Top             =   2040
         Width           =   5160
         Begin MSComCtl2.UpDown UpDownDataEmissaoDe 
            Height          =   315
            Left            =   1695
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissaoDe 
            Height          =   315
            Left            =   525
            TabIndex        =   6
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
         Begin MSComCtl2.UpDown UpDownDataEmissaoAte 
            Height          =   315
            Left            =   4200
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissaoAte 
            Height          =   315
            Left            =   3015
            TabIndex        =   7
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
         Begin VB.Label LabelDataEmissaoDe 
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
            TabIndex        =   20
            Top             =   330
            Width           =   315
         End
         Begin VB.Label LabelDataEmissaoAte 
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
            TabIndex        =   19
            Top             =   330
            Width           =   360
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5655
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   158
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpOPReqComprasOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpOPReqComprasOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpOPReqComprasOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpOPReqComprasOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   14
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
      Left            =   3795
      Picture         =   "RelOpOPReqComprasOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   135
      Width           =   1635
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpOPReqComprasOcx.ctx":0A96
      Left            =   6015
      List            =   "RelOpOPReqComprasOcx.ctx":0AA0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1815
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpOPReqComprasOcx.ctx":0ABE
      Left            =   945
      List            =   "RelOpOPReqComprasOcx.ctx":0AC0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   278
      Width           =   2700
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
      Left            =   6075
      TabIndex        =   16
      Top             =   1515
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
      Left            =   225
      TabIndex        =   15
      Top             =   308
      Width           =   615
   End
End
Attribute VB_Name = "RelOpOPReqComprasOcx"
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
Const ORD_POR_EMISSAO = 1


Private WithEvents objEventoCodigoDe As AdmEvento
Attribute objEventoCodigoDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoAte As AdmEvento
Attribute objEventoCodigoAte.VB_VarHelpID = -1
Private WithEvents objEventoFilialDe As AdmEvento
Attribute objEventoFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoFilialAte As AdmEvento
Attribute objEventoFilialAte.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 68981
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 68982

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 68981
        
        Case 68982
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170342)

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
    If lErro <> SUCESSO Then gError 68983
    
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CheckOPAtendidas.Value = vbUnchecked
    
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 68983
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170343)

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
    Set objEventoFilialDe = New AdmEvento
    Set objEventoFilialAte = New AdmEvento
        
    ComboOrdenacao.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170344)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoCodigoDe = Nothing
    Set objEventoCodigoAte = Nothing
    Set objEventoFilialDe = Nothing
    Set objEventoFilialAte = Nothing
    
End Sub



Private Sub CodigoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoAte, iAlterado)
    
End Sub

Private Sub CodigoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoDe, iAlterado)
    
End Sub

Private Sub DataEmissaoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)
    
End Sub

Private Sub DataEmissaoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEmissaoDe, iAlterado)
    
End Sub

Private Sub FilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(FilialAte, iAlterado)
    
End Sub

Private Sub FilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FilialDe, iAlterado)
    
End Sub

Private Sub LabelCodigoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objOrdemProducao As New ClassOrdemDeProducao

On Error GoTo Erro_LabelCodigoAte_Click

    If Len(Trim(CodigoAte.Text)) > 0 Then
        'Preenche com o Pedido de Venda da tela
        objOrdemProducao.sCodigo = CodigoAte.Text
    End If

    'Chama Tela OrdProdTodasListaModal
    Call Chama_Tela("OrdProdTodasListaModal", colSelecao, objOrdemProducao, objEventoCodigoAte)

   Exit Sub

Erro_LabelCodigoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170345)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objOrdemProducao As New ClassOrdemDeProducao

On Error GoTo Erro_LabelCodigoDe_Click

    If Len(Trim(CodigoDe.Text)) > 0 Then
        'Preenche com o Pedido de Venda da tela
        objOrdemProducao.sCodigo = CodigoDe.Text
    End If

    'Chama Tela OrdProdTodasListaModal
    Call Chama_Tela("OrdProdTodasListaModal", colSelecao, objOrdemProducao, objEventoCodigoDe)

   Exit Sub

Erro_LabelCodigoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170346)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataEmissaoDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEmissaoDe.Text)
    If lErro <> SUCESSO Then gError 68984

    Exit Sub
                   
Erro_DataEmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68984
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170347)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataEmissaoAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEmissaoAte.Text)
    If lErro <> SUCESSO Then gError 68985

    Exit Sub
                   
Erro_DataEmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 68985
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170348)

    End Select

    Exit Sub

End Sub


Private Sub UpDownDataEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 68986

    Exit Sub

Erro_UpDownDataEmissaoDe_DownClick:

    Select Case gErr

        Case 68986
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170349)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68987

    Exit Sub

Erro_UpDownDataEmissaoDe_UpClick:

    Select Case gErr

        Case 68987
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170350)

    End Select

    Exit Sub

End Sub
Private Sub UpDownDataEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 68988

    Exit Sub

Erro_UpDownDataEmissaoAte_DownClick:

    Select Case gErr

        Case 68988
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170351)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68989

    Exit Sub

Erro_UpDownDataEmissaoAte_UpClick:

    Select Case gErr

        Case 68989
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170352)

    End Select

    Exit Sub

End Sub

Private Sub LabelEmpresaDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelEmpresaDe_Click

    If Len(Trim(FilialDe.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(FilialDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoFilialDe)

   Exit Sub

Erro_LabelEmpresaDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170353)

    End Select

    Exit Sub

End Sub
Private Sub LabelEmpresaAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelEmpresaAte_Click

    If Len(Trim(FilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(FilialAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoFilialAte)

   Exit Sub

Erro_LabelEmpresaAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170354)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    FilialAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    FilialDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodigoAte_evSelecao(obj1 As Object)

Dim objOrdemProducao As New ClassOrdemDeProducao

    Set objOrdemProducao = obj1

    CodigoAte.Text = objOrdemProducao.sCodigo

    Me.Show

End Sub

Private Sub objEventoCodigoDe_evSelecao(obj1 As Object)

Dim objOrdemProducao As New ClassOrdemDeProducao

    Set objOrdemProducao = obj1

    CodigoDe.Text = objOrdemProducao.sCodigo

    Me.Show

End Sub


Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 68990

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 68991

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 68992
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 68993
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 68990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 68991 To 68993
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170355)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 68994

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 68995

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr
    
        Case 68994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 68995

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170356)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 68996

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "OPCodigo", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RCCodigo", 1)

            Case ORD_POR_EMISSAO

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEmissao", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "OPCodigo", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "RCCodigo", 1)
                
            Case Else
                gError 74951

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 68996, 74951

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170357)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sFilialEmpresa_I As String
Dim sFilialEmpresa_F As String
Dim sCodigo_I As String
Dim sCodigo_F As String
Dim sOrdenacaoPor As String
Dim iOrdenacao As Long
Dim sOrd As String
Dim sCheck As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sFilialEmpresa_I, sFilialEmpresa_F, sCodigo_I, sCodigo_F)
    If lErro <> SUCESSO Then gError 68997

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 68998
         
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILEMPINIC", sFilialEmpresa_I)
    If lErro <> AD_BOOL_TRUE Then gError 68999
         
    lErro = objRelOpcoes.IncluirParametro("TCODOPINIC", sCodigo_I)
    If lErro <> AD_BOOL_TRUE Then gError 72000
    
    'Preenche dataemissao inicial
    If Trim(DataEmissaoDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMIINIC", DataEmissaoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMIINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 72001
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILEMPFIM", sFilialEmpresa_F)
    If lErro <> AD_BOOL_TRUE Then gError 72002
         
    lErro = objRelOpcoes.IncluirParametro("TCODOPFIM", sCodigo_F)
    If lErro <> AD_BOOL_TRUE Then gError 72003
    
    'Preenche data de emissao Final
    If Trim(DataEmissaoAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMIFIM", DataEmissaoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMIFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 72004
    
    
    'Inclui OP Atendidas
    If CheckOPAtendidas.Value Then
        sCheck = vbChecked
        
    Else
        sCheck = vbUnchecked
    End If

    lErro = objRelOpcoes.IncluirParametro("NOPATENDIDAS", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 72005
    
    Select Case ComboOrdenacao.ListIndex
        
            Case ORD_POR_CODIGO
            
                sOrdenacaoPor = "CodOP"
                
            Case ORD_POR_EMISSAO
                sOrdenacaoPor = "DataEmissao"
                
            Case Else
                gError 72006
                  
    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 72007
   
    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 72008
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sFilialEmpresa_I, sFilialEmpresa_F, sCodigo_I, sCodigo_F, sCheck, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 72009

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 68997, 68998, 68999
        
        Case 72000 To 72009
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170358)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sFilialEmpresa_I As String, sFilialEmpresa_F As String, sCodigo_I As String, sCodigo_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

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
        
        If StrParaLong(sCodigo_I) > StrParaLong(sCodigo_F) Then gError 72010
        
    End If
    
    'critica CodigoFilialEmpresa Inicial e Final
    If FilialDe.Text <> "" Then
        sFilialEmpresa_I = CStr(FilialDe.Text)
    Else
        sFilialEmpresa_I = ""
    End If

    If FilialAte.Text <> "" Then
        sFilialEmpresa_F = CStr(FilialAte.Text)
    Else
        sFilialEmpresa_F = ""
    End If

    If sFilialEmpresa_I <> "" And sFilialEmpresa_F <> "" Then

        If StrParaInt(sFilialEmpresa_I) > StrParaInt(sFilialEmpresa_F) Then gError 72011

    End If
    
    'data de Envio inicial não pode ser maior que a final
    If Trim(DataEmissaoDe.ClipText) <> "" And Trim(DataEmissaoAte.ClipText) <> "" Then
    
         If CDate(DataEmissaoDe.Text) > CDate(DataEmissaoAte.Text) Then gError 72012
    
    End If
    
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
        
        Case 72010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_OP_INICIAL_MAIOR", gErr)
            CodigoDe.SetFocus
                
        Case 72011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            FilialDe.SetFocus
                
        Case 72012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_INICIAL_MAIOR", gErr)
            DataEmissaoDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170359)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sFilialEmpresa_I As String, sFilialEmpresa_F As String, sCodigo_I As String, sCodigo_F As String, sCheck As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sCodigo_I <> "" Then sExpressao = "OPCodigo >= " & Forprint_ConvTexto(sCodigo_I)

   If sCodigo_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "OPCodigo <= " & Forprint_ConvTexto(sCodigo_F)

    End If

    If sFilialEmpresa_I <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresaCod >= " & Forprint_ConvInt(StrParaInt(sFilialEmpresa_I))
    End If
    
    If sFilialEmpresa_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresaCod <= " & Forprint_ConvInt(StrParaInt(sFilialEmpresa_F))

    End If

   If Trim(DataEmissaoDe.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataEmissao >= " & Forprint_ConvData(CDate(DataEmissaoDe.Text))
        
    End If
    
    If Trim(DataEmissaoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataEmissao <= " & Forprint_ConvData(CDate(DataEmissaoAte.Text))

    End If
        
'''    If sCheck <> "" Then
'''        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'''        sExpressao = sExpressao & "Faturado= " & Forprint_ConvInt(StrParaInt(sCheck))
'''    End If
'''
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170360)

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
    If lErro <> SUCESSO Then gError 72013
   
    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCODOPINIC", sParam)
    If lErro <> SUCESSO Then gError 72014
    
    CodigoDe.Text = sParam
    
    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("TCODOPFIM", sParam)
    If lErro <> SUCESSO Then gError 72015
    
    CodigoAte.Text = sParam
                
    'pega CodigoFilialEmpresa inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILEMPINIC", sParam)
    If lErro <> SUCESSO Then gError 72016
    
    FilialDe.Text = sParam
    Call FilialDe_Validate(bSGECancelDummy)
    
    'pega  CodigoFilialEmpresa final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILEMPFIM", sParam)
    If lErro <> SUCESSO Then gError 72017
    
    FilialAte.Text = sParam
    Call FilialAte_Validate(bSGECancelDummy)
                                                           
    'pega DataEmissao inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DEMIINIC", sParam)
    If lErro <> SUCESSO Then gError 72018
    
    Call DateParaMasked(DataEmissaoDe, CDate(sParam))
    
    'pega data de emissao final e exibe
    lErro = objRelOpcoes.ObterParametro("DEMIFIM", sParam)
    If lErro <> SUCESSO Then gError 72019

    Call DateParaMasked(DataEmissaoAte, CDate(sParam))
   
    'pega 'OP Atendidas' e exibe
    lErro = objRelOpcoes.ObterParametro("NOPATENDIDAS", sParam)
    If lErro <> SUCESSO Then gError 72521

    If sParam = "1" Then
        CheckOPAtendidas.Value = vbChecked
    Else
        CheckOPAtendidas.Value = vbUnchecked
    End If
   
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 72020
    
    Select Case sOrdenacaoPor
        
            Case "CodOP"
            
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
            Case "DataEmissao"
                
                ComboOrdenacao.ListIndex = ORD_POR_EMISSAO
                
            Case Else
                gError 72021
                  
    End Select
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 72013 To 72021, 72521
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170361)

    End Select

    Exit Function

End Function
Private Sub FilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialDe_Validate

    If Len(Trim(FilialDe.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(FilialDe.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 72022
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 72023

    End If

    Exit Sub

Erro_FilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 72022

        Case 72023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170362)

    End Select

    Exit Sub

End Sub
Private Sub FilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialAte_Validate

    If Len(Trim(FilialAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(FilialAte.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 72024
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 72025

    End If

    Exit Sub

Erro_FilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 72024

        Case 72025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170363)

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
    Caption = "Relação de Ordens de Produção X Requisições de Compra"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpOPReqCompras"
    
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
            
        ElseIf Me.ActiveControl Is FilialDe Then
            Call LabelEmpresaDe_Click
        
        ElseIf Me.ActiveControl Is FilialAte Then
            Call LabelEmpresaAte_Click
        
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






Private Sub LabelEmpresaDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelEmpresaDe, Source, X, Y)
End Sub

Private Sub LabelEmpresaDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelEmpresaDe, Button, Shift, X, Y)
End Sub

Private Sub LabelEmpresaAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelEmpresaAte, Source, X, Y)
End Sub

Private Sub LabelEmpresaAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelEmpresaAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDataEmissaoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataEmissaoDe, Source, X, Y)
End Sub

Private Sub LabelDataEmissaoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataEmissaoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDataEmissaoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataEmissaoAte, Source, X, Y)
End Sub

Private Sub LabelDataEmissaoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataEmissaoAte, Button, Shift, X, Y)
End Sub

