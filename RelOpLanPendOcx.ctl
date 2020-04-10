VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpLanPendOcx 
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   6495
   Begin VB.Frame Frame4 
      Caption         =   "Lotes"
      Height          =   1185
      Left            =   3405
      TabIndex        =   27
      Top             =   2955
      Width           =   2880
      Begin VB.TextBox LoteFinal 
         Height          =   315
         Left            =   975
         MaxLength       =   4
         TabIndex        =   8
         Top             =   690
         Width           =   1695
      End
      Begin VB.TextBox LoteInicial 
         Height          =   315
         Left            =   990
         MaxLength       =   4
         TabIndex        =   7
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label LabelLoteFinal 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   29
         Top             =   780
         Width           =   480
      End
      Begin VB.Label LabelLoteInicial 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   28
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4200
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpLanPendOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpLanPendOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpLanPendOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpLanPendOcx.ctx":0816
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
      Left            =   4380
      Picture         =   "RelOpLanPendOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpLanPendOcx.ctx":0A96
      Left            =   840
      List            =   "RelOpLanPendOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   570
      Width           =   2430
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datas"
      Height          =   1260
      Left            =   180
      TabIndex        =   21
      Top             =   1575
      Width           =   2895
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   2025
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   885
         TabIndex        =   1
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   2025
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   750
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   885
         TabIndex        =   2
         Top             =   750
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   25
         Top             =   345
         Width           =   585
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   315
         TabIndex        =   24
         Top             =   795
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Documentos"
      Height          =   1185
      Left            =   180
      TabIndex        =   18
      Top             =   2940
      Width           =   2880
      Begin MSMask.MaskEdBox DocumentoInicial 
         Height          =   285
         Left            =   930
         TabIndex        =   3
         Top             =   315
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DocumentoFinal 
         Height          =   285
         Left            =   945
         TabIndex        =   4
         Top             =   765
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDocInicial 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   360
         Width           =   585
      End
      Begin VB.Label LabelDocFinal 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   810
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Origens"
      Height          =   1260
      Left            =   3435
      TabIndex        =   15
      Top             =   1575
      Width           =   2880
      Begin VB.ComboBox OrigemInicial 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   255
         Width           =   1695
      End
      Begin VB.ComboBox OrigemFinal 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   735
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   255
         TabIndex        =   17
         Top             =   315
         Width           =   585
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   780
         Width           =   480
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   165
      TabIndex        =   26
      Top             =   615
      Width           =   630
   End
End
Attribute VB_Name = "RelOpLanPendOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoLancamentoInic As AdmEvento
Attribute objEventoLancamentoInic.VB_VarHelpID = -1
Private WithEvents objEventoLancamentoFim As AdmEvento
Attribute objEventoLancamentoFim.VB_VarHelpID = -1

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 47014

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 47015

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
        
        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 47067
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 47014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 47015, 47067

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169761)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 40990

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 40990

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169762)

    End Select
    
    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp

    If Len(DataInicial.ClipText) <> 0 And Len(DataFinal.ClipText) <> 0 Then
        'data inicial não pode ser maior que a data final
        If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 40991
    End If
    
    If Len(Trim(DocumentoInicial.Text)) <> 0 And Len(Trim(DocumentoFinal.Text)) <> 0 Then
        'Documento Inicial não pode ser maior que o Documento Final
        If CLng(DocumentoInicial.Text) > CLng(DocumentoFinal.Text) Then gError 40992
    End If
    
    If Len(OrigemInicial.Text) <> 0 And Len(OrigemFinal.Text) <> 0 Then
        'Origem Inicial não pode ser maior que a Origem Final
        If (gobjColOrigem.Origem(OrigemInicial.Text)) > (gobjColOrigem.Origem(OrigemFinal.Text)) Then gError 40993
    End If
    
    'lote inicial não pode ser maior que o lote final
    If LoteInicial.Text <> "" And LoteFinal.Text <> "" Then
        If CInt(LoteInicial.Text) > CInt(LoteFinal.Text) Then gError 90623
    End If
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 40994
    
    'Preenche data Inicial
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 40995
    
    'Preenche  data final
    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 40996
    
    lErro = objRelOpcoes.IncluirParametro("NDOCINIC", DocumentoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 40997
    
    lErro = objRelOpcoes.IncluirParametro("NDOCFIM", DocumentoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 40998
    
    lErro = objRelOpcoes.IncluirParametro("TORIGINIC", gobjColOrigem.Origem(OrigemInicial.Text))
    If lErro <> AD_BOOL_TRUE Then gError 40999
    
    lErro = objRelOpcoes.IncluirParametro("TORIGFIM", gobjColOrigem.Origem(OrigemFinal.Text))
    If lErro <> AD_BOOL_TRUE Then gError 47000

    lErro = objRelOpcoes.IncluirParametro("NLOTEINIC", LoteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90619

    lErro = objRelOpcoes.IncluirParametro("NLOTEFIM", LoteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90620
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 47001

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr
        
        Case 90623
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_INICIAL_MAIOR", gErr)
        
        Case 40991
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
        
        Case 40992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_INICIAL_MAIOR", gErr)
        
        Case 40993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_INICIAL_MAIOR", gErr)
        
        Case 40994, 40995, 40996, 40997, 40998, 40999, 47000, 47001, 90619, 90620

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169763)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 47011

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47012

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 47013

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47068
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 47011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 47012, 47013, 47068

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169764)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()

    Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47066
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47066
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169765)

    End Select

    Exit Sub
    
End Sub

Private Sub ComboOpcoes_Click()

    DataInicial.PromptInclude = False
    DataInicial.Text = ""
    DataInicial.PromptInclude = True
    
    DataFinal.PromptInclude = False
    DataFinal.Text = ""
    DataFinal.PromptInclude = True
       
    DocumentoInicial.PromptInclude = False
    DocumentoInicial.Text = ""
    DocumentoInicial.PromptInclude = True
    
    DocumentoFinal.PromptInclude = False
    DocumentoFinal.Text = ""
    DocumentoFinal.PromptInclude = True
    
    OrigemInicial.ListIndex = -1
    OrigemFinal.ListIndex = -1
    
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 40988

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 40988

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169766)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 40989

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 40989

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169767)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
  
    Set objEventoLancamentoInic = New AdmEvento
    Set objEventoLancamentoFim = New AdmEvento
    
    OrigemInicial.Clear
    
    'Inicializar a Origem Inicial
    For iIndice = 1 To gobjColOrigem.Count
        OrigemInicial.AddItem gobjColOrigem.Item(iIndice).sDescricao
    Next
    
    OrigemFinal.Clear
    
    'Inicializar a Origem Final
    For iIndice = 1 To gobjColOrigem.Count
        OrigemFinal.AddItem gobjColOrigem.Item(iIndice).sDescricao
    Next
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169768)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoLancamentoInic = Nothing
    Set objEventoLancamentoFim = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29562
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 40985
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 40985
        
        Case 29562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169769)

    End Select

    Exit Function

End Function

Private Sub DocumentoFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DocumentoFinal)

End Sub

Private Sub DocumentoInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DocumentoInicial)

End Sub

Private Sub LabelDocFinal_Click()

Dim lErro As Long
Dim objLancamento_Detalhe As New ClassLancamento_Detalhe
Dim colSelecao As New Collection
    
On Error GoTo Erro_LabelDocFinal_Click
    
    Call Chama_Tela("LancamentoLista", colSelecao, objLancamento_Detalhe, objEventoLancamentoFim)
  
    Exit Sub
    
Erro_LabelDocFinal_Click:
    
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169770)

    End Select

    Exit Sub

End Sub

Private Sub LabelDocInicial_Click()

Dim lErro As Long
Dim objLancamento_Detalhe As New ClassLancamento_Detalhe
Dim colSelecao As New Collection

On Error GoTo Erro_LabelDocInicial_Click

    Call Chama_Tela("LancamentoLista", colSelecao, objLancamento_Detalhe, objEventoLancamentoInic)

    Exit Sub
    
Erro_LabelDocInicial_Click:
      
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169771)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção

Dim sExpressao As String
Dim lErro As Long
    
On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""
    
    If Trim(DataInicial.ClipText) <> "" Then sExpressao = sExpressao & " LanPendData >= " & Forprint_ConvData(CDate(DataInicial.Text))

    If Trim(DataFinal.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & " LanPendData <= " & Forprint_ConvData(CDate(DataFinal.Text))
    End If
        
    If Trim(DocumentoInicial.Text) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "LanPendDoc >= " & Forprint_ConvLong(CLng(DocumentoInicial.Text))
    End If
    
    If Trim(DocumentoFinal.Text) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "LanPendDoc <= " & Forprint_ConvLong(CLng(DocumentoFinal.Text))
    End If
    
    If Trim(OrigemInicial.Text) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "LanPendOrigem >= " & Forprint_ConvTexto(gobjColOrigem.Origem(OrigemInicial.Text))
    End If
        
    If Trim(OrigemFinal.Text) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "LanPendOrigem <= " & Forprint_ConvTexto(gobjColOrigem.Origem(OrigemFinal.Text))
    End If
    
    If Trim(LoteInicial.Text) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Lote >= " & Forprint_ConvInt(CInt(LoteInicial.Text))
    End If
    
    If Trim(LoteFinal.Text) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Lote <= " & Forprint_ConvInt(CInt(LoteFinal.Text))
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169772)

    End Select

    Exit Function

End Function


Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sOrigemInicial As String
Dim sOrigemFinal As String
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar()
    If lErro <> SUCESSO Then gError 47002
    
    'Data Inicial
    lErro = objRelOpcoes.ObterParametro("DDINIC", sParam)
    If lErro <> SUCESSO Then gError 47003
    
    If sParam <> "07/09/1822" Then
    
        'coloca a data Inicial na tela
        DataInicial.PromptInclude = False
        DataInicial.Text = sParam
        DataInicial.PromptInclude = True
    
    End If
    
    'Data Final
    lErro = objRelOpcoes.ObterParametro("DDFIM", sParam)
    If lErro <> SUCESSO Then gError 47004
    
    If sParam <> "07/09/1822" Then
    
        'coloca a data Final na tela
        DataFinal.PromptInclude = False
        DataFinal.Text = sParam
        DataFinal.PromptInclude = True
        
    End If
    
    'Documento Inicial
    lErro = objRelOpcoes.ObterParametro("NDOCINIC", sParam)
    If lErro <> SUCESSO Then gError 47005
    
    'coloca a data Inicial na tela
    DocumentoInicial.PromptInclude = False
    DocumentoInicial.Text = sParam
    DocumentoInicial.PromptInclude = True
    
    'Documento Final
    lErro = objRelOpcoes.ObterParametro("NDOCFIM", sParam)
    If lErro <> SUCESSO Then gError 47006
    
    'coloca a documento Final na tela
    DocumentoFinal.PromptInclude = False
    DocumentoFinal.Text = sParam
    DocumentoFinal.PromptInclude = True
    
    'Documento Final
    lErro = objRelOpcoes.ObterParametro("TORIGINIC", sParam)
    If lErro <> SUCESSO Then gError 47007
    
    sOrigemInicial = sParam
    
    'Documento Final
    lErro = objRelOpcoes.ObterParametro("TORIGFIM", sParam)
    If lErro <> SUCESSO Then gError 47008
    
    sOrigemFinal = sParam
        
    lErro = Traz_Origem_Tela(sOrigemInicial, sOrigemFinal)
    If lErro <> SUCESSO Then gError 47009
        
     'pega Lote Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NLOTEINIC", sParam)
    If lErro <> SUCESSO Then gError 90621

    LoteInicial.Text = sParam

    'pega Lote Final e exibe
    lErro = objRelOpcoes.ObterParametro("NLOTEFIM", sParam)
    If lErro <> SUCESSO Then gError 90622

    LoteFinal.Text = sParam
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 47002, 47003, 47004, 47005, 47006, 47007, 47008, 47009, 90621, 90622

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169773)

    End Select

    Exit Function

End Function

Function Traz_Origem_Tela(sOrigemInicial As String, sOrigemFinal As String) As Long
'Coloca as Origens na tela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_Origem_Tela

    For iIndice = 0 To OrigemInicial.ListCount - 1
        If OrigemInicial.List(iIndice) = gobjColOrigem.Descricao(sOrigemInicial) Then
            OrigemInicial.ListIndex = iIndice
            Exit For
        End If
    Next
            
    For iIndice = 0 To OrigemFinal.ListCount - 1
        If OrigemFinal.List(iIndice) = gobjColOrigem.Descricao(sOrigemFinal) Then
            OrigemFinal.ListIndex = iIndice
            Exit For
        End If
    Next
 
    Traz_Origem_Tela = SUCESSO

    Exit Function

Erro_Traz_Origem_Tela:

    Traz_Origem_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169774)

    End Select

    Exit Function

End Function

Private Sub objEventoLancamentoFim_evSelecao(obj1 As Object)

'Traz o lançamento selecionado para a tela

Dim lErro As Long
Dim objLancamento_Detalhe As ClassLancamento_Detalhe

On Error GoTo Erro_objEventoLancamento_evSelecao

    Set objLancamento_Detalhe = obj1

    DocumentoFinal.PromptInclude = False
    DocumentoFinal.Text = CStr(objLancamento_Detalhe.lDoc)
    DocumentoFinal.PromptInclude = True
    
    Me.Show

    Exit Sub

Erro_objEventoLancamento_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169775)

    End Select

    Exit Sub

End Sub

Private Sub objEventoLancamentoInic_evSelecao(obj1 As Object)

'Traz o lançamento selecionado para a tela

Dim lErro As Long
Dim objLancamento_Detalhe As ClassLancamento_Detalhe

On Error GoTo Erro_objEventoLancamento_evSelecao

    Set objLancamento_Detalhe = obj1
    
    DocumentoInicial.PromptInclude = False
    DocumentoInicial.Text = CStr(objLancamento_Detalhe.lDoc)
    DocumentoInicial.PromptInclude = True
 
    Me.Show

    Exit Sub

Erro_objEventoLancamento_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169776)

    End Select

    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_LANCAMENTO_PENDENTE
    Set Form_Load_Ocx = Me
    Caption = "Lançamentos Pendentes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpLanPend"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is DocumentoInicial Then
            Call LabelDocInicial_Click
        ElseIf Me.ActiveControl Is DocumentoFinal Then
            Call LabelDocFinal_Click
        End If
    
    End If

End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelDocInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDocInicial, Source, X, Y)
End Sub

Private Sub LabelDocInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDocInicial, Button, Shift, X, Y)
End Sub

Private Sub LabelDocFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDocFinal, Source, X, Y)
End Sub

Private Sub LabelDocFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDocFinal, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
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

Private Sub LabelLoteInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelLoteInicial, Source, X, Y)
End Sub

Private Sub LabelLoteInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelLoteInicial, Button, Shift, X, Y)
End Sub

Private Sub LabelLoteFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelLoteFinal, Source, X, Y)
End Sub

Private Sub LabelLoteFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelLoteFinal, Button, Shift, X, Y)
End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 13280

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 13280
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169777)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 13281

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 13281
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169778)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 13282

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 13282
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169779)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 13284

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 13284
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169780)

    End Select

    Exit Sub

End Sub


