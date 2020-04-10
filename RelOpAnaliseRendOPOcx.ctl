VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpAnaliseRendOPOcx 
   Appearance      =   0  'Flat
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   KeyPreview      =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   5820
   Begin VB.Frame FrameOrdemProducao 
      Caption         =   "Ordem de Produção"
      Height          =   2175
      Left            =   60
      TabIndex        =   26
      Top             =   2430
      Width           =   5685
      Begin VB.Frame FrameOPCodigo 
         Caption         =   "Código"
         Height          =   800
         Left            =   200
         TabIndex        =   30
         Top             =   300
         Width           =   5295
         Begin VB.TextBox OpCodigoInicial 
            Height          =   300
            Left            =   780
            TabIndex        =   4
            Top             =   300
            Width           =   1515
         End
         Begin VB.TextBox OpCodigoFinal 
            Height          =   300
            Left            =   3390
            TabIndex        =   5
            Top             =   300
            Width           =   1515
         End
         Begin VB.Label LabelOpFinal 
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
            Left            =   2970
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   32
            Top             =   360
            Width           =   360
         End
         Begin VB.Label LabelOpInicial 
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
            Left            =   360
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   31
            Top             =   353
            Width           =   315
         End
      End
      Begin VB.Frame FrameOPData 
         Caption         =   "Data"
         Height          =   800
         Left            =   200
         TabIndex        =   27
         Top             =   1200
         Width           =   5295
         Begin MSComCtl2.UpDown UpDownOPDataInicial 
            Height          =   315
            Left            =   1725
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   293
            Width           =   180
            _ExtentX        =   397
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox OPDataInicial 
            Height          =   300
            Left            =   750
            TabIndex        =   6
            Top             =   300
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownOPDataFinal 
            Height          =   315
            Left            =   4350
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   293
            Width           =   180
            _ExtentX        =   397
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox OPDataFinal 
            Height          =   300
            Left            =   3390
            TabIndex        =   8
            Top             =   300
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelOPDataInicial 
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
            Height          =   240
            Left            =   360
            TabIndex        =   29
            Top             =   330
            Width           =   345
         End
         Begin VB.Label LabelOPDataFinal 
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
            Left            =   2970
            TabIndex        =   28
            Top             =   353
            Width           =   360
         End
      End
   End
   Begin VB.Frame FrameEstatisticas 
      Caption         =   "Estatísticas - Variação %"
      Height          =   855
      Left            =   60
      TabIndex        =   23
      Top             =   1500
      Width           =   5685
      Begin MSMask.MaskEdBox PercVariacaoMenor 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PercVariacaoMaior 
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   360
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label LabelPercVariacaoMaior 
         AutoSize        =   -1  'True
         Caption         =   "% Maior que:"
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
         Left            =   3000
         TabIndex        =   25
         Top             =   420
         Width           =   1110
      End
      Begin VB.Label LabelPercVariacaoMenor 
         AutoSize        =   -1  'True
         Caption         =   "% Menor que:"
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
         Left            =   480
         TabIndex        =   24
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.Frame FrameProduto 
      Caption         =   "Produtos"
      Height          =   1332
      Left            =   90
      TabIndex        =   18
      Top             =   4680
      Width           =   5655
      Begin MSMask.MaskEdBox ProdutoFinal 
         Height          =   315
         Left            =   750
         TabIndex        =   11
         Top             =   870
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoInicial 
         Height          =   315
         Left            =   750
         TabIndex        =   10
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescProdInic 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   22
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label DescProdFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2295
         TabIndex        =   21
         Top             =   885
         Width           =   3135
      End
      Begin VB.Label LabelProdutoDe 
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   375
         Width           =   315
      End
      Begin VB.Label LabelProdutoAte 
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   900
         Width           =   360
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
      Left            =   3885
      Picture         =   "RelOpAnaliseRendOPOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   780
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpAnaliseRendOPOcx.ctx":0102
      Left            =   735
      List            =   "RelOpAnaliseRendOPOcx.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   2640
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3600
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpAnaliseRendOPOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpAnaliseRendOPOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpAnaliseRendOPOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpAnaliseRendOPOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
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
      Left            =   60
      TabIndex        =   17
      Top             =   210
      Width           =   615
   End
End
Attribute VB_Name = "RelOpAnaliseRendOPOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoOpDe As AdmEvento
Attribute objEventoOpDe.VB_VarHelpID = -1
Private WithEvents objEventoOpAte As AdmEvento
Attribute objEventoOpAte.VB_VarHelpID = -1

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub LabelOpFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOp As ClassOrdemDeProducao

On Error GoTo Erro_LabelOpFinal_Click

    If Len(Trim(OpCodigoFinal.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpCodigoFinal.Text

    End If

    Call Chama_Tela("OrdemProdBaixadasLista", colSelecao, objOp, objEventoOpAte)

   Exit Sub

Erro_LabelOpFinal_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167012)

    End Select

End Sub

Private Sub LabelOpInicial_Click()

Dim lErro As Long
Dim objOp As ClassOrdemDeProducao
Dim colSelecao As Collection

On Error GoTo Erro_LabelOpInicial_Click

    If Len(Trim(OpCodigoInicial.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpCodigoInicial.Text

    End If

    Call Chama_Tela("OrdemProdBaixadasLista", colSelecao, objOp, objEventoOpDe)

    Exit Sub

Erro_LabelOpInicial_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167013)

    End Select

End Sub


Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoFinal.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 103070

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 103070

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167014)

    End Select

End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoInicial.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 103071

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 103071

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167015)

    End Select

End Sub

Private Sub OpCodigoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpCodigoFinal_Validate

    If Len(Trim(OpCodigoFinal.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpCodigoFinal.Text)
        If lErro <> SUCESSO Then gError 103082

    End If

    Exit Sub

Erro_OpCodigoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 103082

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167016)

    End Select

End Sub

Private Sub OpCodigoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpCodigoInicial_Validate

    If Len(Trim(OpCodigoInicial.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpCodigoInicial.Text)
        If lErro <> SUCESSO Then gError 103081

    End If

    Exit Sub

Erro_OpCodigoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 103081

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167017)

    End Select

End Sub

Private Sub PercVariacaoMaior_Validate(Cancel As Boolean)
'Critica o Campo

Dim lErro As Long

On Error GoTo Erro_PercVariacaoMaior_Validate

    If Len(Trim(PercVariacaoMaior.ClipText)) > 0 Then
        
        'Critica a percentagem indicada
        lErro = Porcentagem_Critica(PercVariacaoMaior.ClipText)
        If lErro <> SUCESSO Then gError 108580
        
    End If

    Exit Sub
    
Erro_PercVariacaoMaior_Validate:

    Cancel = True

    Select Case gErr
    
        Case 108580

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167018)

    End Select

End Sub

Private Sub PercVariacaoMenor_Validate(Cancel As Boolean)
'Critica o Campo

Dim lErro As Long

On Error GoTo Erro_PercVariacaoMenor_Validate

    If Len(Trim(PercVariacaoMenor.ClipText)) > 0 Then
        
        'Critica a percentagem indicada
        lErro = Porcentagem_Critica(PercVariacaoMenor.ClipText)
        If lErro <> SUCESSO Then gError 108580
        
    End If

    Exit Sub
    
Erro_PercVariacaoMenor_Validate:

    Cancel = True

    Select Case gErr
    
        Case 108580

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167019)

    End Select

End Sub

Private Sub UpDownOPDataInicial_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataInicial_DownClick

    lErro = Data_Up_Down_Click(OPDataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37977

    Exit Sub

Erro_UpDownOPDataInicial_DownClick:

    Select Case Err

        Case 37977
            OPDataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167020)

    End Select

End Sub

Private Sub UpDownOPDataInicial_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataInicial_UpClick

    lErro = Data_Up_Down_Click(OPDataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37978

    Exit Sub

Erro_UpDownOPDataInicial_UpClick:

    Select Case Err

        Case 37978
            OPDataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167021)

    End Select

End Sub

Private Sub UpDownOPDataFinal_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataFinal_DownClick

    lErro = Data_Up_Down_Click(OPDataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37977

    Exit Sub

Erro_UpDownOPDataFinal_DownClick:

    Select Case Err

        Case 37977
            OPDataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167022)

    End Select

End Sub

Private Sub UpDownOPDataFinal_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownOPDataFinal_UpClick

    lErro = Data_Up_Down_Click(OPDataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37978

    Exit Sub

Erro_UpDownOPDataFinal_UpClick:

    Select Case Err

        Case 37978
            OPDataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167023)

    End Select

End Sub

Private Sub ProdutoInicial_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoInicial)

End Sub

Private Sub ProdutoFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoFinal)

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoInicial_Validate

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 108511

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 108512

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 108513
        
'*************************
        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 108591
        
        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 108592
        
        'Se não controla estoque => Erro
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 108593
        
        'Se nao for um produto produzido => Erro
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then gError 108594
'*************************

        DescProdInic.Caption = objProduto.sDescricao
        
    End If
    
    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 108512, 108511

        Case 108513
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoInicial.Text)

'*************************
        Case 108591
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoInicial.Text)
            
        Case 108592
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoInicial.Text)
        
        Case 108593
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoInicial.Text)
            
        Case 108594
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, ProdutoInicial.Text)
'*************************

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167024)

    End Select

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoFinal_Validate

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 108511

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 108512

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 108513
        
'*************************
        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 108591
        
        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 108592
        
        'Se não controla estoque => Erro
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 108593
        
        'Se nao for um produto produzido => Erro
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then gError 108594
'*************************

        DescProdFim.Caption = objProduto.sDescricao
        
    End If
    
    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 108512, 108511

        Case 108513
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoFinal.Text)

'*************************
        Case 108591
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoFinal.Text)
            
        Case 108592
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoFinal.Text)
        
        Case 108593
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoFinal.Text)
            
        Case 108594
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, ProdutoFinal.Text)
'*************************

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167025)

    End Select

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoOpDe = New AdmEvento
    Set objEventoOpAte = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 103051

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 103052

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 103051, 103052, 103087

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167026)

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoOpDe = Nothing
    Set objEventoOpAte = Nothing
    Set gobjRelOpcoes = Nothing
    Set gobjRelatorio = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'Parent.HelpContextID =
    Set Form_Load_Ocx = Me
    Caption = "Análise de Rendimento por Ordem de Produção"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpAnaliseRendOP"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        ElseIf Me.ActiveControl Is OpCodigoInicial Then
            Call LabelOpInicial_Click
        ElseIf Me.ActiveControl Is OpCodigoFinal Then
            Call LabelOpFinal_Click
        End If

    End If

End Sub

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

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 106470

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 106471

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 106472

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 106473

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 106470
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 106471, 106472, 106473

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167027)

    End Select

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 106473

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 106474

    
    lErro = objRelOpcoes.IncluirParametro("TPRODINI", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 106475

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 106476

    
    lErro = objRelOpcoes.IncluirParametro("TOPINI", OpCodigoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 106475

    lErro = objRelOpcoes.IncluirParametro("TOPFIM", OpCodigoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 106476

    If OPDataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDTOPINI", OPDataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDTOPINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 106475
    
    If OPDataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDTOPFIM", OPDataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDTOPFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 106476

    lErro = objRelOpcoes.IncluirParametro("NPERCMAIOR", StrParaDbl(PercVariacaoMenor.Text))
    If lErro <> AD_BOOL_TRUE Then gError 106475

    lErro = objRelOpcoes.IncluirParametro("NPERCMENOR", StrParaDbl(PercVariacaoMaior.Text))
    If lErro <> AD_BOOL_TRUE Then gError 106476


    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 106483

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 106473 To 106483

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167028)

    End Select

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sProd_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "PRODUTO >= " & Forprint_ConvTexto(sProd_I)

    End If

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PRODUTO <= " & Forprint_ConvTexto(sProd_F)

    End If

    If Trim(OpCodigoInicial.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CODIGOOP >= " & Forprint_ConvTexto(OpCodigoInicial.Text)

    End If

    If Trim(OpCodigoFinal.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CODIGOOP <= " & Forprint_ConvTexto(OpCodigoFinal.Text)

    End If

    If Trim(OPDataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DATAOP >= " & Forprint_ConvData(StrParaDate(OPDataInicial.Text))

    End If

    If Trim(OPDataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DATAOP <= " & Forprint_ConvData(StrParaDate(OPDataFinal.Text))

    End If
    
'    If Trim(PercVariacaoMenor.Text) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "PERCMENOR <= " & Forprint_ConvTexto(-StrParaDbl(PercVariacaoMenor.Text))
'
'    End If
'
'    If Trim(PercVariacaoMaior.Text) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "PERCMAIOR => " & Forprint_ConvTexto(PercVariacaoMaior.Text)
'
'    End If
        
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167029)

    End Select

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sProdutoMascarado As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 106485

    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINI", sParam)
    If lErro <> SUCESSO Then gError 106486

    If (Len(Trim(sParam)) > 0) Then
        
        lErro = Mascara_MascararProduto(sParam, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 106487
    
        ProdutoInicial.Text = sProdutoMascarado
        Call ProdutoInicial_Validate(bSGECancelDummy)
        
    End If

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 106488

    If Len(Trim(sParam)) > 0 Then
        
        lErro = Mascara_MascararProduto(sParam, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 106487
    
        ProdutoFinal.Text = sProdutoMascarado
        Call ProdutoFinal_Validate(bSGECancelDummy)
        
    End If
    
    'pega a OP Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TOPINI", sParam)
    If lErro <> SUCESSO Then gError 106492
    OpCodigoInicial.Text = sParam
    Call OpCodigoInicial_Validate(bSGECancelDummy)

    'pega a OP Final e exibe
    lErro = objRelOpcoes.ObterParametro("TOPFIM", sParam)
    OpCodigoFinal.Text = sParam
    If lErro <> SUCESSO Then gError 106493
    Call OpCodigoFinal_Validate(bSGECancelDummy)

    'pega a Data Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDTOPINI", sParam)
    If lErro <> SUCESSO Then gError 106494
    Call DateParaMasked(OPDataInicial, StrParaDate(sParam))

    'pega a Data Final e exibe
    lErro = objRelOpcoes.ObterParametro("DDTOPFIM", sParam)
    If lErro <> SUCESSO Then gError 106495
    Call DateParaMasked(OPDataFinal, StrParaDate(sParam))
    
    'pega a Percentagem Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NPERCMAIOR", sParam)
    If lErro <> SUCESSO Then gError 106494
    PercVariacaoMenor.Text = sParam
    Call PercVariacaoMenor_Validate(bSGECancelDummy)
    
    'pega a Percentagem Final e exibe
    lErro = objRelOpcoes.ObterParametro("NPERCMENOR", sParam)
    If lErro <> SUCESSO Then gError 106495
    PercVariacaoMaior.Text = sParam
    Call PercVariacaoMenor_Validate(bSGECancelDummy)

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 106485 To 106495, 108540, 108541, 108542

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167030)

    End Select

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 106496

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REL_OP_ANALISE_REND_OP")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 106497

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 106498

        ComboOpcoes.Text = ""
        DescProdInic.Caption = ""
        DescProdFim.Caption = ""

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 106496
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 106497, 106498

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167031)

    End Select

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 108500

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 108500

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167032)

    End Select

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 103064

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 103065

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 103066

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 103064, 103066

        Case 103065
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167033)

    End Select

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 103067

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 103068

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 103069

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 103067, 103069

        Case 103068
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167034)

    End Select

End Sub

Private Sub objEventoOpDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOpDe_evSelecao

    Set objOp = obj1

    objOp.iFilialEmpresa = giFilialEmpresa

    'busca ordem de produção baixada
    lErro = CF("OPBaixada_Le_SemItens", objOp)
    If lErro <> SUCESSO And lErro <> 34459 Then gError 103079

    If lErro = 34459 Then gError 103080

    'Coloca na tela o Código da OP
    OpCodigoInicial.Text = objOp.sCodigo

    Me.Show

    Exit Sub

Erro_objEventoOpDe_evSelecao:

    Select Case gErr

        Case 106458, 103079

        Case 103080
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", gErr)

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167035)

    End Select

End Sub

Private Sub objEventoOpAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOpAte_evSelecao

    Set objOp = obj1

    objOp.iFilialEmpresa = giFilialEmpresa

    'busca ordem de produção baixada
    lErro = CF("OPBaixada_Le_SemItens", objOp)
    If lErro <> SUCESSO And lErro <> 34459 Then gError 103079

    If lErro = 34459 Then gError 103080

    'Coloca na tela o Código da OP
    OpCodigoFinal.Text = objOp.sCodigo

    Me.Show

    Exit Sub

Erro_objEventoOpAte_evSelecao:

    Select Case gErr

        Case 106460, 103079

        Case 103080
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", gErr)

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167036)

    End Select

End Sub

Private Function Valida_OrdProd(sCodigoOP As String) As Long

Dim objOp As New ClassOrdemDeProducao
Dim lErro As Long

On Error GoTo Erro_Valida_OrdProd

    objOp.iFilialEmpresa = giFilialEmpresa
    objOp.sCodigo = sCodigoOP

    'busca ordem de produção baixada
    lErro = CF("OPBaixada_Le_SemItens", objOp)
    If lErro <> SUCESSO And lErro <> 34459 Then gError 103079

    If lErro = 34459 Then gError 103080

    Valida_OrdProd = SUCESSO

    Exit Function

Erro_Valida_OrdProd:

    Valida_OrdProd = gErr

    Select Case gErr

        Case 103079

        Case 103080
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167037)

    End Select

End Function

Private Sub BotaoLimpar_Click()
'Faz a Limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 106462

    ComboOpcoes.Text = ""
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""

    ComboOpcoes.SetFocus

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 106462

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167038)

    End Select

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 106465

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 106466

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 106467

    End If

    'Se a OP Inicial for maior que a Final => Erro
    If OpCodigoInicial.Text > OpCodigoFinal.Text Then gError 108544

    'Se a Data Inicial da OP for maior que a Data Final da OP => Erro
    If StrParaDate(OPDataInicial.Text) > StrParaDate(OPDataFinal.Text) Then gError 108545
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
        
        Case 106465
            ProdutoInicial.SetFocus

        Case 106466
            ProdutoFinal.SetFocus

        Case 106467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoInicial.SetFocus
            
        Case 108544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INICIAL_MAIOR", gErr)
            OpCodigoInicial.SetFocus
        
        Case 108545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_OP_INICIAL_MAIOR", gErr)
            OPDataInicial.SetFocus
             '??? ERRO_DATA_OP_INICIAL_MAIOR
             
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167039)

    End Select

End Function

Public Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 103088

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 103089

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 103089

        Case 103088
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167040)

    End Select

End Function


Private Sub OPDataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_OPDataInicial_Validate

    If Len(OPDataInicial.ClipText) > 0 Then

        sDataInic = OPDataInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 103072

    End If

    Exit Sub

Erro_OPDataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 103072

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167041)

    End Select

End Sub

Private Sub OPDataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_OPDataFinal_Validate

    If Len(OPDataFinal.ClipText) > 0 Then

        sDataFim = OPDataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 103073

    End If

    Exit Sub

Erro_OPDataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 103073

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167042)

    End Select

End Sub
