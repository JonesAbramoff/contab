VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOPLogWFWOcx 
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8010
   LockControls    =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   8010
   Begin VB.Frame Frame3 
      Caption         =   "Usuários"
      Height          =   870
      Left            =   150
      TabIndex        =   24
      Top             =   3495
      Width           =   7725
      Begin VB.ComboBox UsuarioAte 
         Height          =   315
         ItemData        =   "RelOPLogWFW.ctx":0000
         Left            =   5100
         List            =   "RelOPLogWFW.ctx":0002
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   315
         Width           =   2070
      End
      Begin VB.ComboBox UsuarioDe 
         Height          =   315
         ItemData        =   "RelOPLogWFW.ctx":0004
         Left            =   1575
         List            =   "RelOPLogWFW.ctx":0006
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   315
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário Até:"
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
         Index           =   7
         Left            =   3990
         TabIndex        =   26
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário De:"
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
         Index           =   6
         Left            =   525
         TabIndex        =   25
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transações"
      Height          =   1725
      Left            =   150
      TabIndex        =   20
      Top             =   1680
      Width           =   7725
      Begin VB.ComboBox TransacaoAte 
         Height          =   315
         Left            =   1605
         TabIndex        =   7
         Top             =   1230
         Width           =   5580
      End
      Begin VB.ComboBox Modulo 
         Height          =   315
         Left            =   1605
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   3630
      End
      Begin VB.ComboBox TransacaoDe 
         Height          =   315
         Left            =   1605
         TabIndex        =   6
         Top             =   750
         Width           =   5580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transação Até:"
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
         Index           =   3
         Left            =   255
         TabIndex        =   23
         Top             =   1290
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Módulo:"
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
         Index           =   0
         Left            =   885
         TabIndex        =   22
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transação De:"
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
         Index           =   5
         Left            =   300
         TabIndex        =   21
         Top             =   810
         Width           =   1275
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOPLogWFW.ctx":0008
      Left            =   810
      List            =   "RelOPLogWFW.ctx":000A
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   2916
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datas"
      Height          =   750
      Left            =   150
      TabIndex        =   16
      Top             =   810
      Width           =   7725
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   2775
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   397
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   285
         Left            =   1605
         TabIndex        =   1
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   6900
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   397
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   285
         Left            =   5730
         TabIndex        =   3
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   1200
         TabIndex        =   18
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   5280
         TabIndex        =   17
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5715
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOPLogWFW.ctx":000C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOPLogWFW.ctx":018A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOPLogWFW.ctx":06BC
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOPLogWFW.ctx":0846
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
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
      Left            =   3870
      Picture         =   "RelOPLogWFW.ctx":09A0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   1575
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
      Index           =   8
      Left            =   105
      TabIndex        =   19
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "RelOPLogWFWOcx"
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

Dim sModuloAnt As String

Private Sub Modulo_Change()

    If sModuloAnt <> Modulo.Text Then
    
        Call Carrega_Combo_Transacao(TransacaoDe, Modulo.Text)
        Call Carrega_Combo_Transacao(TransacaoAte, Modulo.Text)
    
        sModuloAnt = Modulo.Text
    End If

End Sub

Private Sub Modulo_Click()

    If sModuloAnt <> Modulo.Text Then
    
        Call Carrega_Combo_Transacao(TransacaoDe, Modulo.Text)
        Call Carrega_Combo_Transacao(TransacaoAte, Modulo.Text)
    
        sModuloAnt = Modulo.Text
    End If
    
End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 180114

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 180114
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180115)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 180116

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 180116
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180117)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 180118

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 180118
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180119)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 180120

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 180120
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180121)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 180122
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Atribui o nome do Relatório, fazendo desta tela um modelo genérico.
    Me.Caption = objRelatorio.sCodRel
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 180123
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 180122
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 180123
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180124)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 180125
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 180126
         
    lErro = objRelOpcoes.IncluirParametro("TMODULO", Modulo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 180127
         
    lErro = objRelOpcoes.IncluirParametro("TTRANSINIC", TransacaoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 180128

    lErro = objRelOpcoes.IncluirParametro("TTRANSFIM", TransacaoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 180129

    lErro = objRelOpcoes.IncluirParametro("DTINI", DataInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 180130

    lErro = objRelOpcoes.IncluirParametro("DTFIM", DataFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 180131
    
    lErro = objRelOpcoes.IncluirParametro("TUSUINIC", UsuarioDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 180132

    lErro = objRelOpcoes.IncluirParametro("TUSUFIM", UsuarioAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 180133
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 180134

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr
    
    Select Case gErr

        Case 180125 To 180134

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180135)
            
    End Select
    
End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 180136

    Call gobjRelatorio.Executar_Prossegue2(Me)
    
    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 180136

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 180137)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 180138
    
    ComboOpcoes.Text = ""
    
    UsuarioDe.ListIndex = -1
    UsuarioAte.ListIndex = -1
    Modulo.ListIndex = -1
    
    Call Carrega_Combo_Transacao(TransacaoDe, Modulo.Text)
    Call Carrega_Combo_Transacao(TransacaoAte, Modulo.Text)
    
    Call Padrao_Tela
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 180138
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180139)

    End Select

    Exit Sub
   
End Sub

Private Sub ComboOpcoes_Click()

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
        If lErro <> SUCESSO Then gError 180140

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 180140

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180141)

    End Select

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
        If lErro <> SUCESSO Then gError 180142

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 180142

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180143)

    End Select

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Call Carrega_Combo_Usuario(UsuarioDe)
    Call Carrega_Combo_Usuario(UsuarioAte)
    
    Call Carrega_Combo_Modulo(Modulo)
    
    Call Carrega_Combo_Transacao(TransacaoDe, Modulo.Text)
    Call Carrega_Combo_Transacao(TransacaoAte, Modulo.Text)
    
    Call Padrao_Tela

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180144)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção

Dim lErro As Long, sExpressao As String

On Error GoTo Erro_Monta_Expressao_Selecao

    If Len(Trim(Modulo.Text)) <> 0 Then sExpressao = "Modulo = " & Forprint_ConvTexto(Modulo.Text)

    If StrParaDate(DataInicial.Text) <> DATA_NULA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(DataInicial.Text)

    End If
    
    If StrParaDate(DataFinal.Text) <> DATA_NULA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(DataFinal.Text)

    End If
    
    If UsuarioDe.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Usuario >= " & Forprint_ConvTexto(UsuarioDe.Text)

    End If
    
    If UsuarioAte.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Usuario <= " & Forprint_ConvTexto(UsuarioAte.Text)

    End If
    
    If TransacaoDe.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Transacao >= " & Forprint_ConvTexto(TransacaoDe.Text)

    End If
    
    If TransacaoAte.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Transacao <= " & Forprint_ConvTexto(TransacaoAte.Text)

    End If
    
    objRelOpcoes.sSelecao = sExpressao

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 180145)

    End Select

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 180146
   
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TMODULO", sParam)
    If lErro <> SUCESSO Then gError 180147
    
    Call CF("SCombo_Seleciona2", Modulo, sParam)

    Call Carrega_Combo_Transacao(TransacaoDe, sParam)
    Call Carrega_Combo_Transacao(TransacaoAte, sParam)
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TTRANSINIC", sParam)
    If lErro <> SUCESSO Then gError 180148
    
    TransacaoDe.Text = sParam

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TTRANSFIM", sParam)
    If lErro <> SUCESSO Then gError 180149

    TransacaoAte.Text = sParam

    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TUSUINIC", sParam)
    If lErro <> SUCESSO Then gError 180150

    UsuarioDe.Text = sParam

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TUSUFIM", sParam)
    If lErro <> SUCESSO Then gError 180151
    
    UsuarioAte.Text = sParam
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DTINI", sParam)
    If lErro <> SUCESSO Then gError 180152
    
    If StrParaDate(sParam) <> DATA_NULA Then
        DataInicial.PromptInclude = False
        DataInicial.Text = Format(StrParaDate(sParam), "dd/mm/yy")
        DataInicial.PromptInclude = True
    End If
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DTFIM", sParam)
    If lErro <> SUCESSO Then gError 180153
        
    If StrParaDate(sParam) <> DATA_NULA Then
        DataFinal.PromptInclude = False
        DataFinal.Text = Format(StrParaDate(sParam), "dd/mm/yy")
        DataFinal.PromptInclude = True
    End If
        
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 180146 To 180153

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180154)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros() As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 180155
    
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(TransacaoDe.Text) <> "" And Trim(TransacaoAte.Text) <> "" Then
    
         If TransacaoDe.Text > TransacaoAte.Text Then gError 180156
    
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(UsuarioDe.Text) <> "" And Trim(UsuarioAte.Text) <> "" Then
    
         If UsuarioDe.Text > UsuarioAte.Text Then gError 180157
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 180155
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
            
        Case 180156
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_INICIAL_MAIOR", gErr)
            TransacaoDe.SetFocus
            
        Case 180157
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_INICIAL_MAIOR", gErr)
            UsuarioDe.SetFocus
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180158)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 180159

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPLOGWFW")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 180160

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 180161
    
        ComboOpcoes.Text = ""
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 180159
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 180160, 180161

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180162)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 180163

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 180164

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 180165
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 180166
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 180163
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 180164 To 180166

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180167)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_DIARIO
    Set Form_Load_Ocx = Me
    Caption = "Log"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpLogWFW"
    
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

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Function Carrega_Combo_Usuario(objCombo As ComboBox) As Long

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As Object

On Error GoTo Erro_Carrega_Combo_Usuario

    'Le todos os usuarios da tabela usuarios e coloca na colecao
    lErro = CF("Usuarios_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 180107

    'Coloca todos os Usuarios na ComboUsuario
    For Each objUsuarios In colUsuarios
        objCombo.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Combo_Usuario = SUCESSO
    
    Exit Function
    
Erro_Carrega_Combo_Usuario:

    Carrega_Combo_Usuario = gErr
    
    Select Case gErr
    
        Case 180107
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 180108)
    
    End Select
    
    Exit Function

End Function

Private Function Carrega_Combo_Modulo(objCombo As ComboBox) As Long
'carrega a combobox com  os módulos disponiveis para o sistema

Dim lErro As Long
Dim iIndice As Integer
    
On Error GoTo Erro_Carrega_Combo_Modulo

    objCombo.AddItem ""
        
    For iIndice = 1 To gcolModulo.Count
        If gcolModulo.Item(iIndice).iAtivo = MODULO_ATIVO Then
            objCombo.AddItem gcolModulo.Item(iIndice).sNome
        End If
    Next
    
    Carrega_Combo_Modulo = SUCESSO

    Exit Function

Erro_Carrega_Combo_Modulo:

    Carrega_Combo_Modulo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180109)

    End Select
    
    Exit Function

End Function

Private Function Carrega_Combo_Transacao(objCombo As ComboBox, ByVal sModulo As String) As Long
'carrega a combobox que contem as transacoes disponiveis para o modulo selecionado.

Dim lErro As Long
Dim colTransacao As New Collection
Dim objTransacao As ClassTransacaoWFW
    
On Error GoTo Erro_Carrega_Combo_Transacao
        
    objCombo.Clear
        
    'leitura das contas no BD
    lErro = CF("TransacaoWFW_Le_Todos2", gcolModulo.Sigla(sModulo), colTransacao)
    If lErro <> SUCESSO Then gError 180110
    
    For Each objTransacao In colTransacao
        objCombo.AddItem objTransacao.sTransacaoTela
        objCombo.ItemData(objCombo.NewIndex) = objTransacao.iCodigo
    Next
    
    Carrega_Combo_Transacao = SUCESSO

    Exit Function

Erro_Carrega_Combo_Transacao:

    Carrega_Combo_Transacao = gErr

    Select Case gErr

        Case 180110
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180111)

    End Select
    
    Exit Function

End Function

Private Function Padrao_Tela() As Long

    UsuarioDe.Text = gsUsuario
    UsuarioAte.Text = gsUsuario

End Function
