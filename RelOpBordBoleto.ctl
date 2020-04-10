VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpBordBoleto 
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   ScaleHeight     =   2595
   ScaleWidth      =   6840
   Begin VB.Frame FrameBordero 
      Caption         =   "Borderô"
      Height          =   735
      Left            =   210
      TabIndex        =   14
      Top             =   1695
      Width           =   4215
      Begin MSMask.MaskEdBox BorderoDe 
         Height          =   315
         Left            =   840
         TabIndex        =   15
         Top             =   285
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox BorderoAte 
         Height          =   315
         Left            =   2805
         TabIndex        =   16
         Top             =   285
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelBorderoAte 
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
         Height          =   195
         Left            =   2280
         TabIndex        =   18
         Top             =   345
         Width           =   360
      End
      Begin VB.Label LabelBorderoDe 
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
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpBordBoleto.ctx":0000
      Left            =   1305
      List            =   "RelOpBordBoleto.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   285
      Width           =   2670
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   735
      Left            =   210
      TabIndex        =   6
      Top             =   855
      Width           =   4215
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   1650
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   315
         Left            =   690
         TabIndex        =   8
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   300
         Left            =   3645
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   315
         Left            =   2700
         TabIndex        =   10
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataDe 
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
         Left            =   300
         TabIndex        =   12
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelDataAte 
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
         Left            =   2265
         TabIndex        =   11
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4560
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   135
      Width           =   2190
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpBordBoleto.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpBordBoleto.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpBordBoleto.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpBordBoleto.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   2
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
      Left            =   4845
      Picture         =   "RelOpBordBoleto.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1605
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
      Left            =   555
      TabIndex        =   19
      Top             =   315
      Width           =   645
   End
End
Attribute VB_Name = "RelOpBordBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giBorderoInicial As Integer

'Obj utilizado para o browser de Borderos
Private WithEvents objEventoBordero As AdmEvento
Attribute objEventoBordero.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'***** INICIALIZAÇÃO DA TELA - INÍCIO *****
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa objeto usado pelo Browser
    Set objEventoBordero = New AdmEvento
       
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 125140
       
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 125140

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167341)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    'Limpa Objetos da memoria
    Set objEventoBordero = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 125141
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 125142

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 125141
        
        Case 125142
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167342)

    End Select

    Exit Function

End Function
'***** INICIALIZAÇÃO DA TELA - FIM *****

'***** EVENTO GOTFOCUS DOS CONTROLES - INÍCIO *****
Private Sub BorderoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(BorderoAte)
End Sub

Private Sub BorderoDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(BorderoDe)
End Sub

Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe)
End Sub

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte)
End Sub
'***** EVENTO GOTFOCUS DOS CONTROLES - FIM *****

'***** EVENTO VALIDATE DOS CONTROLES - INÍCIO *****
Private Sub BorderoAte_Validate(Cancel As Boolean)
'verifica validade do campo BorderoAte

Dim lErro As Long

On Error GoTo Erro_BorderoAte_Validate

    'Se o campo BorderoAte foi preenchido
    If Len(Trim(BorderoAte.ClipText)) > 0 Then
        
        'verifica validade de BorderoAte
        lErro = Long_Critica(BorderoAte.Text)
        If lErro <> SUCESSO Then gError 125143
        
    End If

    giBorderoInicial = 1

    Exit Sub

Erro_BorderoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125143

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167343)

    End Select

End Sub

Private Sub BorderoDe_Validate(Cancel As Boolean)
'verifica validade do campo BorderoDe
Dim lErro As Long

On Error GoTo Erro_Borderode_Validate

    'Se o campo BorderoDe foi preenchido
    If Len(Trim(BorderoDe.ClipText)) > 0 Then

        'verifica validade de BorderoDe
        lErro = Long_Critica(BorderoDe.Text)
        If lErro <> SUCESSO Then gError 125144
        
    End If

    giBorderoInicial = 1

    Exit Sub

Erro_Borderode_Validate:

    Cancel = True

    Select Case gErr

        Case 125144

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167344)

    End Select
    
    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se DataDe foi preenchida
    If Len(Trim(DataDe.ClipText)) > 0 Then

        'Verifica Validade da DataDe
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 125145

    End If

    Exit Sub

Erro_DataDe_Validate:
    
    Cancel = True

    Select Case gErr

        Case 125145

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167345)

    End Select

    Exit Sub
    
End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'Verifica validade de DataAte

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se DataAte foi preenchida
    If Len(Trim(DataAte.ClipText)) > 0 Then

        'Verifica Validade da DataAte
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 125146

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125146
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167346)

    End Select

    Exit Sub
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)
End Sub
'***** EVENTO VALIDATE DOS CONTROLES - FIM *****

'***** EVENTO CLICK DOS CONTROLES - INÍCIO *****
Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 125147

    'Chama rotina que checa as opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 125148

    'Seta o nome da opção que será gravado como o nome que esta na comboOpções
    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Aciona rotina que grava opções do relatório
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 125149

    'Testa se nome no combo esta igual ao nome em gobjRelOpçoes.sNome
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 125150
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 125147
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 125148 To 125150

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167347)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExcluir_Click()
'Aciona a Rotina de exclusão das opções de relatório

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 125151

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 125152

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Aciona Rotinas para Limpar Tela
        Call BotaoLimpar_Click
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 125151
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 125152

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167348)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()
'Aciona rotinas que que checam as opções do relatório e ativam impressão do mesmo

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click
    
    'aciona rotina que checa opções do relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 125153

    gobjRelatorio.sNomeTsk = "BordBole"

    'Chama rotina que excuta a impressão do relatório
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr
        
        Case 125153
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167349)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoLimpar_Click()
'Aciona Rotinas de Limpeza de tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Chama função que limpa Relatório
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 125154
          
    'Limpa o campo ComboOpcoes
    ComboOpcoes.Text = ""
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 125155
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 125154
        
        Case 125155
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167350)

    End Select

    Exit Sub
    
End Sub

Private Sub ComboOpcoes_Click()
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
End Sub

Private Sub UpDownDataAte_DownClick()
'Diminui DataAte em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125156

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 125156
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167351)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataAte_UpClick()
'Aumenta DataAte em UM dia
Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125157

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 125157
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167352)

    End Select

    Exit Sub
End Sub

Private Sub UpDownDataDe_DownClick()
'Diminui DataDe em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125158

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 125158
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167353)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataDe_UpClick()
'Aumenta DataDe em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125159

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 125160
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167354)

    End Select

    Exit Sub

End Sub
'***** EVENTO CLICK DOS CONTROLES - FIM *****

'***** FUNÇÕES DE APOIO A TELA - INÍCIO *****
Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    DataDe.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167355)
    
    End Select
    
    Exit Function
    
End Function

Private Function Formata_E_Critica_Parametros(sDataInic As String, sDataFim As String, sBorderoIni As String, sBorderoFim As String) As Long
'Formata e verifica validade das opções passadas para gerar o relatório

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    '**** DATA *******
    'formata datas e verifica se DataDe é maior que  DataAte
    If Trim(DataDe.ClipText) <> "" Then
        sDataInic = DataDe.Text
    Else
        sDataInic = CStr(DATA_NULA)
    End If
    
    'verifica se DataAte foi preenchida
    If Trim(DataAte.ClipText) <> "" Then
        sDataFim = DataAte.Text
    Else
        sDataFim = CStr(DATA_NULA)
    End If
    
    'data Data inicial nao pode ser maior que a final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
        If CDate(sDataInic) > CDate(sDataFim) Then gError 125161
        
    End If
    '********* DATA *********
             
    '********* BORDERÔ ***********
    'verifica se o BorderoDe é maior que o BorderoAte
    If Trim(BorderoDe.Text) <> "" Then
        sBorderoIni = CStr(Trim(BorderoDe.Text))
    Else
        sBorderoIni = ""
    End If
    
    If Trim(BorderoAte.Text) <> "" Then
        sBorderoFim = CStr(Trim(BorderoAte.ClipText))
    Else
        sBorderoFim = ""
    End If
    
    If Trim(BorderoDe.ClipText) <> "" And Trim(BorderoAte.ClipText) <> "" Then
    
         If CLng(sBorderoIni) > CLng(sBorderoFim) Then gError 125162
    
    End If
    '********* BORDERÔ ***********
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
               
        Case 125161
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
               
        Case 125162
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDERO_INICIAL_MAIOR", gErr)
            BorderoDe.SetFocus
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167356)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sDestino As String
Dim sDataInic As String
Dim sDataFim As String
Dim sBorderoIni As String
Dim sBorderoFim As String
Dim sContaCorrente As String

On Error GoTo Erro_PreencherRelOp

    'Verifica Parametros , e formata os mesmos
    lErro = Formata_E_Critica_Parametros(sDataInic, sDataFim, sBorderoIni, sBorderoFim)
    If lErro <> SUCESSO Then gError 125163
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 125164
   
    'Inclui parametro de DataDe
    lErro = objRelOpcoes.IncluirParametro("DINI", sDataInic)
    If lErro <> AD_BOOL_TRUE Then gError 125165

    'Inclui parametro de DataAte
    lErro = objRelOpcoes.IncluirParametro("DFIM", sDataFim)
    If lErro <> AD_BOOL_TRUE Then gError 125166
       
    'Inclui parametro de BorderoDe
    lErro = objRelOpcoes.IncluirParametro("NBORDEROINIC", BorderoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125167
    
    'Inclui parametro de BorderoAte
    lErro = objRelOpcoes.IncluirParametro("NBORDEROFIM", BorderoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125168
    
    'Aciona Rotina que monta_expressão que será usada para gerar relatório
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sDataInic, sDataFim, sBorderoIni, sBorderoFim)
    If lErro <> SUCESSO Then gError 125169
        
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 125163 To 125169
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167357)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sDataInic As String, sDataFim As String, sBorderoIni As String, sBorderoFim As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    'Verifica se campo DataDe foi preenchido
    If Trim(DataDe.ClipText) <> "" Then sExpressao = "Data >= " & Forprint_ConvData(CDate(sDataInic))

    'Verifica se campo DataAte foi preenchido
    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de DataAte
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(sDataFim))

    End If
        
    'Verifica se campo BorderoDe foi preenchido
    If Trim(BorderoDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de BorderoDe
        sExpressao = sExpressao & "Bordero >= " & Forprint_ConvLong(CLng(sBorderoIni))

    End If
    
    'Verifica se campo BorderoAte foi preenchido
    If Trim(BorderoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        'Inclui na expressaõ o valor de BorderoAte
        sExpressao = sExpressao & "Bordero <= " & Forprint_ConvLong(CLng(sBorderoFim))

    End If
    
    'passa a expressão completa para o obj
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If


    Monta_Expressao_Selecao = SUCESSO
    
    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167358)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim Cancel As Boolean

On Error GoTo Erro_PreencherParametrosNaTela

    'Carrega parametros do relatorio gravado
    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 125170
            
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINI", sParam)
    If lErro <> SUCESSO Then gError 125171

    'Preenche campo DataDe
    Call DateParaMasked(DataDe, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 125172

    'Preenche campo DataAte
    Call DateParaMasked(DataAte, CDate(sParam))
        
    'Pega parametro BorderoDe e o Exibe
    lErro = objRelOpcoes.ObterParametro("NBORDEROINIC", sParam)
    If lErro <> SUCESSO Then gError 125173

    'Preenche campo BorderoDe
    BorderoDe.PromptInclude = False
    BorderoDe.Text = sParam
    BorderoDe.PromptInclude = True
    Call BorderoDe_Validate(bSGECancelDummy)
    
    'Pega parametro BorderoAte e o Exibe
    lErro = objRelOpcoes.ObterParametro("NBORDEROFIM", sParam)
    If lErro <> SUCESSO Then gError 125174

    'Preenche campo BorderoAte
    BorderoAte.PromptInclude = False
    BorderoAte.Text = sParam
    BorderoAte.PromptInclude = True
    Call BorderoAte_Validate(bSGECancelDummy)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 125170 To 125174
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167359)

    End Select

    Exit Function

End Function
'***** FUNÇÕES DE APOIO A TELA - FIM *****

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Borderô Boleto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpBordBoleto"

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


