VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GVExporta 
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
   ScaleHeight     =   2685
   ScaleWidth      =   6300
   Begin VB.Frame Frame2 
      Caption         =   "Localização do arquivo"
      Height          =   1230
      Left            =   135
      TabIndex        =   19
      Top             =   1380
      Width           =   6045
      Begin VB.CheckBox optNomeAuto 
         Caption         =   "Nome do Arquivo Automático"
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
         Left            =   915
         TabIndex        =   8
         Top             =   930
         Value           =   1  'Checked
         Width           =   3195
      End
      Begin VB.TextBox NomeArquivo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   930
         TabIndex        =   7
         Top             =   570
         Width           =   4440
      End
      Begin VB.CommandButton BotaoProcurar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5370
         TabIndex        =   6
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox NomeDiretorio 
         Height          =   285
         Left            =   930
         TabIndex        =   5
         Top             =   225
         Width           =   4440
      End
      Begin VB.Label Label1 
         Caption         =   "Arquivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   570
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   ".txt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   5
         Left            =   5400
         TabIndex        =   21
         Top             =   600
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Diretório:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   20
         Top             =   255
         Width           =   795
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Período"
      Height          =   630
      Left            =   135
      TabIndex        =   16
      Top             =   705
      Width           =   6045
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   2070
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   315
         Left            =   930
         TabIndex        =   1
         Top             =   225
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   4920
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   195
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   330
         Left            =   3795
         TabIndex        =   3
         Top             =   210
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelPeriodoDe 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   525
         TabIndex        =   18
         Top             =   255
         Width           =   390
      End
      Begin VB.Label LabelPeriodoAte 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   3375
         TabIndex        =   17
         Top             =   255
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3870
      ScaleHeight     =   495
      ScaleWidth      =   2220
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   45
      Width           =   2280
      Begin VB.CommandButton BotaoExecutar 
         Height          =   360
         Left            =   150
         Picture         =   "GVExporta.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exportar Arquivos"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1170
         Picture         =   "GVExporta.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1665
         Picture         =   "GVExporta.ctx":05CC
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   660
         Picture         =   "GVExporta.ctx":074A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1620
         Picture         =   "GVExporta.ctx":08A4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   75
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "GVExporta.ctx":0DD6
      Left            =   1050
      List            =   "GVExporta.ctx":0DD8
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   2685
   End
   Begin VB.Label Label4 
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
      Left            =   345
      TabIndex        =   15
      Top             =   225
      Width           =   615
   End
End
Attribute VB_Name = "GVExporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209216)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError ERRO_SEM_MENSAGEM
   
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINI", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataInicial, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataFinal, CDate(sParam))
            
    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("TNOMEDIR", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    NomeDiretorio.Text = sParam
    
    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("TNOMEARQ", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    NomeArquivo.Text = sParam

    lErro = objRelOpcoes.ObterParametro("NNOMEAUTO", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sParam) = MARCADO Then
        optNomeAuto.Value = vbChecked
        Call Calcula_NomeArq
    Else
        optNomeAuto.Value = vbUnchecked
    End If
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209217)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
       
End Sub

Function Trata_Parametros(Optional objRelatorio As AdmRelatorio, Optional objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 209218
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Calcula_NomeArq
    
    If ComboOpcoes.ListCount <> 0 Then
        ComboOpcoes.ListIndex = 0
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case 209218
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209219)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Critica_Parametros() As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim sCliente_De As String, sCliente_Ate As String

On Error GoTo Erro_Critica_Parametros
    
    If StrParaDate(DataInicial.Text) <> DATA_NULA And StrParaDate(DataFinal.Text) <> DATA_NULA And StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 209222
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 209248
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 209249
        
    Critica_Parametros = SUCESSO

    Exit Function

Erro_Critica_Parametros:

    Critica_Parametros = gErr

    Select Case gErr
        
        Case 209222
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus

        Case 209248
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
       
        Case 209249
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209224)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209225)

    End Select

    Exit Sub
    
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iNomeAuto As Integer
Dim sNomeArq As String

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = Critica_Parametros()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    If optNomeAuto.Value = vbChecked Then
        iNomeAuto = MARCADO
    Else
        iNomeAuto = DESMARCADO
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NNOMEAUTO", CStr(iNomeAuto))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINI", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEDIR", NomeDiretorio.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("TNOMEARQ", NomeArquivo.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
        
    If bExecutando Then
    
        sNomeArq = NomeDiretorio.Text & NomeArquivo.Text & ".txt"
                  
        lErro = CF("GV_Exporta", sNomeArq, StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209226)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 209227

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 209227
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209228)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_ARQUIVO_GERADO", NomeDiretorio.Text & NomeArquivo.Text & ".txt")

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209229)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 209230

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError ERRO_SEM_MENSAGEM

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 209230
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209231)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
     
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209232)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CONTAS_CORRENTES
    Set Form_Load_Ocx = Me
    Caption = "Exportação do arquivo para o gerenciador de Vendas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GVExprota"
    
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
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Localização física do arquivo"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        NomeDiretorio.Text = sBuffer
        Call NomeDiretorio_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209236)

    End Select

    Exit Sub
  
End Sub

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPos As Integer

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub
    
    If right(NomeDiretorio.Text, 1) <> "\" And right(NomeDiretorio.Text, 1) <> "/" Then
        iPos = InStr(1, NomeDiretorio.Text, "/")
        If iPos = 0 Then
            NomeDiretorio.Text = NomeDiretorio.Text & "\"
        Else
            NomeDiretorio.Text = NomeDiretorio.Text & "/"
        End If
    End If

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 209237

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 209237, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209238)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Call Calcula_NomeArq

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209243)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Call Calcula_NomeArq

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209244)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Calcula_NomeArq

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209245)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Calcula_NomeArq

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataInicial.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209246)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Calcula_NomeArq

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209247)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Calcula_NomeArq

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataFinal.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209248)

    End Select

    Exit Sub

End Sub

Private Sub optNomeAuto_Click()
    If optNomeAuto.Value = vbChecked Then
        Call Calcula_NomeArq
        NomeArquivo.Enabled = False
    Else
        NomeArquivo.Enabled = True
    End If
End Sub

Private Sub Calcula_NomeArq()
Dim sNomeArq As String
    If optNomeAuto.Value = vbChecked Then
    
        sNomeArq = "GV" & Format(giFilialEmpresa, "00") & "_DTE" & Format(Date, "YYYYMMDD") & Format(Now, "HHMMSS")
        
        If StrParaDate(DataInicial.Text) <> DATA_NULA Then
            sNomeArq = sNomeArq & "_DTI" & Format(StrParaDate(DataInicial.Text), "YYYYMMDD")
        End If
        If StrParaDate(DataFinal.Text) <> DATA_NULA Then
            sNomeArq = sNomeArq & "_DTF" & Format(StrParaDate(DataFinal.Text), "YYYYMMDD")
        End If
        
        NomeArquivo.Text = sNomeArq
    End If
End Sub
