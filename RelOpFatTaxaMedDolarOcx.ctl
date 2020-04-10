VERSION 5.00
Begin VB.UserControl RelOpFatTaxaMedDolar 
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   ScaleHeight     =   2745
   ScaleWidth      =   7935
   Begin VB.Frame Frame1 
      Caption         =   "Referente à"
      Height          =   870
      Left            =   210
      TabIndex        =   13
      Top             =   1650
      Width           =   5310
      Begin VB.OptionButton Recebimentos 
         Caption         =   "Recebimentos"
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
         Left            =   2715
         TabIndex        =   15
         Top             =   315
         Width           =   1560
      End
      Begin VB.OptionButton Pagamentos 
         Caption         =   "Pagamentos"
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
         Left            =   315
         TabIndex        =   14
         Top             =   330
         Value           =   -1  'True
         Width           =   1560
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mês / Ano"
      Height          =   825
      Left            =   180
      TabIndex        =   9
      Top             =   690
      Width           =   5325
      Begin VB.ComboBox Mes 
         Height          =   315
         ItemData        =   "RelOpFatTaxaMedDolarOcx.ctx":0000
         Left            =   585
         List            =   "RelOpFatTaxaMedDolarOcx.ctx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   315
         Width           =   1050
      End
      Begin VB.ComboBox Ano 
         Height          =   315
         ItemData        =   "RelOpFatTaxaMedDolarOcx.ctx":0094
         Left            =   3240
         List            =   "RelOpFatTaxaMedDolarOcx.ctx":00BA
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   315
         Width           =   1095
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   360
         Width           =   465
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2790
         TabIndex        =   12
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpFatTaxaMedDolarOcx.ctx":0100
      Left            =   1695
      List            =   "RelOpFatTaxaMedDolarOcx.ctx":0102
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   225
      Width           =   2730
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
      Left            =   5745
      Picture         =   "RelOpFatTaxaMedDolarOcx.ctx":0104
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   780
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5595
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpFatTaxaMedDolarOcx.ctx":0206
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpFatTaxaMedDolarOcx.ctx":0384
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpFatTaxaMedDolarOcx.ctx":08B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpFatTaxaMedDolarOcx.ctx":0A40
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label3 
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
      Left            =   990
      TabIndex        =   7
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "RelOpFatTaxaMedDolar"
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

'***** CARREGAMENTO DA TELA - INÍCIO *****
Public Sub Form_Load()

Dim lErro As Long
Dim sMes As String
Dim sAno As String

On Error GoTo Erro_Form_Load

    'Carrega o ano e o mês
    Call Carrega_Mes_Ano(sMes, sAno)

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 125807

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179508)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 125808
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 125809
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 125808
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 125809
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179509)

    End Select

    Exit Function

End Function
'***** CARREGAMENTO DA TELA - FIM *****

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub
'***** EVENTO VALIDATE DOS CONTROLES - FIM *****

'***** EVENTO CLICK DOS CONTROLES - INÍCIO *****
Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 125810

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 125811

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = LimpaRelatorioDemClienteProd()
        If lErro <> SUCESSO Then gError 125812
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 125810
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 125811, 125812

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179510)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'Preenche o Relatório
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 125813

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 125813

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179511)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 125814

    'Preenche o Relatório com os dados da tela
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 125815

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 125816

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 125814
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 125815, 125816

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179512)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar

    'Limpa a tela
    lErro = LimpaRelatorioDemClienteProd()
    If lErro <> SUCESSO Then gError 125817
    
    Exit Sub
    
Erro_BotaoLimpar:

    Select Case gErr

        Case 125817
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179513)

    End Select

    Exit Sub

End Sub
'***** EVENTO CLICK DOS CONTROLES - FIM *****

'***** FUNÇÕES DE APOIO À TELA *****
Private Function LimpaRelatorioDemClienteProd() As Long
'Limpa a tela
    
Dim lErro As Long
Dim sMes As String
Dim sAno As String

On Error GoTo Erro_LimpaRelatorioDemClienteProd
    
    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 125818
    
    'limpa a combo opções
    ComboOpcoes.Text = ""
    
    Call Carrega_Mes_Ano(sMes, sAno)
    
    LimpaRelatorioDemClienteProd = SUCESSO
    
    Exit Function
    
Erro_LimpaRelatorioDemClienteProd:
    
    LimpaRelatorioDemClienteProd = gErr
    
    Select Case gErr

        Case 125818
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179514)

    End Select

    Exit Function

End Function

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
    
    iMax = Ano.ListCount
       
    For iIndice = 0 To iMax
        If Ano.List(iIndice) = CStr(iAno) Then
            Ano.ListIndex = iIndice
            Exit For
        End If
    Next

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o objRelOp com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp
   
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 125819
        
    'Inclui o Mês
    lErro = objRelOpcoes.IncluirParametro("NMES", CStr(Mes.ItemData(Mes.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then gError 125820

    'Inclui o Ano
    lErro = objRelOpcoes.IncluirParametro("NANO", Ano.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125821
    
    'Faz a chamada da função que irá montar a expressão
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 125822
    
    lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", "1")
    If lErro <> AD_BOOL_TRUE Then gError 125788
    
    If Pagamentos.Value = True Then
        gobjRelatorio.sNomeTsk = "DolarPag"
    Else
        gobjRelatorio.sNomeTsk = "DolarRec"
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 125819 To 125822
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179515)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
    
'    'Verifica se o Mes está preenchido
'    If Trim(Mes.Text) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Mês = " & Forprint_ConvTexto(Mes.Text)
'
'    End If
'
'    'Verifica se o ano está preenchido
'    If Trim(Ano.Text) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Ano = " & Forprint_ConvTexto(Ano.Text)
'
'    End If
'
'    If sExpressao <> "" Then
'
'        objRelOpcoes.sSelecao = sExpressao
'
'    End If

    Monta_Expressao_Selecao = SUCESSO
    
    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179516)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sMes As String
Dim sAno As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 125823
    
    'pega o mês
    lErro = objRelOpcoes.ObterParametro("NMES", sParam)
    If lErro <> SUCESSO Then gError 125824
        
    'Aribui mês
    sMes = sParam
            
    'pega o ano
    lErro = objRelOpcoes.ObterParametro("NANO", sParam)
    If lErro <> SUCESSO Then gError 125825
    
    'Atribui o ano
    sAno = sParam
    
    'Com valores atribuídos de mês e ano, carrega as combos
    Call Carrega_Mes_Ano(sMes, sAno)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 125823 To 125825
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179517)

    End Select

    Exit Function

End Function
'***** FUNÇÕES DE APOIO À TELA - FIM

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Taxa Média do Dólar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpFatTaxaMedDolar"
    
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


