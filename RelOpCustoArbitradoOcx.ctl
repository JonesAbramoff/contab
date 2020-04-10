VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpCustoArbitradoOcx 
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   LockControls    =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   5805
   Begin VB.Frame Frame2 
      Caption         =   "Custo Arbitrado"
      Height          =   585
      Left            =   180
      TabIndex        =   17
      Top             =   2835
      Width           =   5460
      Begin MSMask.MaskEdBox PercCustoArbitrado 
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Top             =   225
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Caption         =   "do maior preço de venda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1065
         TabIndex        =   19
         Top             =   255
         Width           =   4155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Outras Opções"
      Height          =   1200
      Left            =   180
      TabIndex        =   12
      Top             =   1575
      Width           =   5460
      Begin VB.OptionButton OptExibicao 
         Caption         =   "Exibir todos os produtos produzíveis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   225
         TabIndex        =   15
         Top             =   810
         Width           =   4665
      End
      Begin VB.OptionButton OptExibicao 
         Caption         =   "Exibir somente os produtos SEM estoque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   225
         TabIndex        =   14
         Top             =   555
         Width           =   4665
      End
      Begin VB.OptionButton OptExibicao 
         Caption         =   "Exibir somente os produtos COM estoque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   225
         TabIndex        =   13
         Top             =   300
         Value           =   -1  'True
         Width           =   4665
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mês / Ano"
      Height          =   825
      Left            =   180
      TabIndex        =   8
      Top             =   660
      Width           =   4215
      Begin VB.ComboBox Mes 
         Height          =   315
         ItemData        =   "RelOpCustoArbitradoOcx.ctx":0000
         Left            =   585
         List            =   "RelOpCustoArbitradoOcx.ctx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   315
         Width           =   1440
      End
      Begin MSMask.MaskEdBox Ano 
         Height          =   315
         Left            =   2760
         TabIndex        =   16
         Top             =   315
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelAno 
         Caption         =   "Ano:"
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
         Height          =   240
         Left            =   2325
         TabIndex        =   11
         Top             =   360
         Width           =   420
      End
      Begin VB.Label labelMes 
         Caption         =   "Mês:"
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
         Height          =   285
         Left            =   135
         TabIndex        =   10
         Top             =   360
         Width           =   465
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
      Height          =   570
      Left            =   4470
      Picture         =   "RelOpCustoArbitradoOcx.ctx":0094
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   735
      Width           =   1155
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpCustoArbitradoOcx.ctx":0196
      Left            =   840
      List            =   "RelOpCustoArbitradoOcx.ctx":0198
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   195
      Width           =   2460
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3495
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpCustoArbitradoOcx.ctx":019A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpCustoArbitradoOcx.ctx":0318
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpCustoArbitradoOcx.ctx":084A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpCustoArbitradoOcx.ctx":09D4
         Style           =   1  'Graphical
         TabIndex        =   4
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
      Left            =   165
      TabIndex        =   9
      Top             =   255
      Width           =   615
   End
End
Attribute VB_Name = "RelOpCustoArbitradoOcx"
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

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Call Carrega_Mes_Ano
    
    PercCustoArbitrado.Text = CStr(70)

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169106)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sMes As String
Dim sAno As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 87218

    lErro = objRelOpcoes.ObterParametro("NEXIBICAO", sParam)
    If lErro <> SUCESSO Then gError 87221
    
    OptExibicao(StrParaInt(sParam)).Value = True
    
    'pega o mês
    lErro = objRelOpcoes.ObterParametro("NMES", sParam)
    If lErro <> SUCESSO Then gError 87221
        
    'Atribui mês
    sMes = sParam
            
    'pega o ano
    lErro = objRelOpcoes.ObterParametro("NANO", sParam)
    If lErro <> SUCESSO Then gError 87221
    
    'Atribui o ano
    sAno = sParam
    
    'Com valores atribuídos de mês e ano, carrega as combos
    Call Carrega_Mes_Ano(sMes, sAno)
    
    'pega o ano
    lErro = objRelOpcoes.ObterParametro("NPERCCUSTO", sParam)
    If lErro <> SUCESSO Then gError 87221
    
    'Atribui o ano
    PercCustoArbitrado.Text = sParam
    Call PercCustoArbitrado_Validate(bSGECancelDummy)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 87218 To 87221

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169107)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 87222

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 87223

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 87222

        Case 87223
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169108)

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
    If lErro <> SUCESSO Then gError 87224

    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    OptExibicao(0).Value = True

    Call Carrega_Mes_Ano

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 87224

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169109)

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
Dim lNumIntRel As Long
Dim sAno As String
Dim sMes As String
Dim sTipo As String
Dim dPercCusto As Double

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sAno, sMes, sTipo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
                   
    lErro = objRelOpcoes.IncluirParametro("NMES", sMes)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NANO", sAno)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM

    lErro = objRelOpcoes.IncluirParametro("NEXIBICAO", sTipo)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    dPercCusto = StrParaDbl(Val(PercCustoArbitrado.Text) / 100)
    
    lErro = objRelOpcoes.IncluirParametro("NPERCCUSTO", PercCustoArbitrado.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If bExecutando Then
    
        lErro = CF("RelCustoArbitrado_Prepara", lNumIntRel, StrParaInt(sAno), StrParaInt(sMes), StrParaInt(sTipo), dPercCusto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    End If

    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 87245

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169110)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 87233

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 87234

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 87235

        ComboOpcoes.Text = ""

        Call Carrega_Mes_Ano
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 87233
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 87234, 87235

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169111)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169112)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 87237

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

        Case 87237
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169113)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169114)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sAno As String, sMes As String, sTipo As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    If Len(Ano.Text) = 0 Then gError 204114
    If Len(Mes.Text) = 0 Then gError 204115
    
    sMes = CStr(Mes.ItemData(Mes.ListIndex))
    sAno = Ano.Text
    
    'verifica opção selecionada
    If OptExibicao(0).Value Then sTipo = CStr(0)
    If OptExibicao(1).Value Then sTipo = CStr(1)
    If OptExibicao(2).Value Then sTipo = CStr(2)
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 204114
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREECHIDO", gErr)
        
        Case 204115
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREECHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169115)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_FAT_VENDEDOR
    Set Form_Load_Ocx = Me
    Caption = "Custo Arbitrado"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "RelOpCustoArbitrado"
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

Private Sub LabelMes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(labelMes, Source, X, Y)
End Sub

Private Sub LabelMes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(labelMes, Button, Shift, X, Y)
End Sub

Private Sub LabelAno_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAno, Source, X, Y)
End Sub

Private Sub LabelAno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAno, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Carrega_Mes_Ano(Optional sMes As String, Optional sAno As String)
'Função responsável pelo carregamento das combos ano e mês que não são editaveis

Dim iMes As Integer
Dim iAno As Integer
Dim iIndice As Integer
Dim iMax As Integer

    'Se a função for chamada de Define_Padrao
    If Len(Trim(sMes)) = 0 Then
        iMes = Month(Date)
    'Se a função for chamada de PreencheParametros na tela
    Else
        iMes = CInt(sMes)
    End If
    
    'Se a função for chamada de Define_Padrao
    If Len(Trim(sAno)) = 0 Then
        iAno = Year(Date)
    'Se a função for chamada de PreencheParametros na tela
    Else
        iAno = CInt(sAno)
    End If
    
    Mes.ListIndex = iMes - 1
       
    Ano.PromptInclude = False
    Ano.Text = CStr(iAno)
    Ano.PromptInclude = True

End Sub

Private Sub PercCustoArbitrado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercCustoArbitrado_Validate

    'Veifica se CargaMax está preenchida
    If Len(Trim(PercCustoArbitrado.Text)) <> 0 Then

       'Critica a CargaMax
       lErro = Porcentagem_Critica(PercCustoArbitrado.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_PercCustoArbitrado_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144357)

    End Select

    Exit Sub

End Sub

Private Sub PercCustoArbitrado_GotFocus()
Dim iAlterado As Integer
    Call MaskEdBox_TrataGotFocus(PercCustoArbitrado, iAlterado)
End Sub

Private Sub Ano_GotFocus()
Dim iAlterado As Integer
    Call MaskEdBox_TrataGotFocus(Ano, iAlterado)

End Sub

Private Sub Ano_Validate(Cancel As Boolean)

On Error GoTo Erro_Ano_Validate

    If Len(Trim(Ano.Text)) > 0 Then

        If Ano.Text < 1900 Then gError 204110
        
    End If
       
    Exit Sub
    
Erro_Ano_Validate:

    Cancel = True

    Select Case gErr
    
        Case 204110
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_INVALIDO", gErr)
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204111)

    End Select
    
End Sub

