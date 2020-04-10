VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpComisCro 
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   6390
   Begin VB.CheckBox CheckPrevisao 
      Caption         =   "Incluir Previsões"
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
      Left            =   240
      TabIndex        =   21
      Top             =   900
      Value           =   1  'Checked
      Width           =   2010
   End
   Begin VB.Frame FrameData 
      Caption         =   "Período"
      Height          =   780
      Left            =   105
      TabIndex        =   14
      Top             =   1245
      Width           =   3840
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1485
         TabIndex        =   15
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
         Left            =   525
         TabIndex        =   16
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   3390
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   2415
         TabIndex        =   18
         Top             =   300
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dFim 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2010
         TabIndex        =   20
         Top             =   360
         Width           =   360
      End
      Begin VB.Label dIni 
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   135
         TabIndex        =   19
         Top             =   315
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpComisCro.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpComisCro.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpComisCro.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpComisCro.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   10
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
      Left            =   4425
      Picture         =   "RelOpComisCro.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   810
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpComisCro.ctx":0A96
      Left            =   870
      List            =   "RelOpComisCro.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2985
   End
   Begin VB.CheckBox CheckPulaPag 
      Caption         =   "Pular página a cada novo vendedor"
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
      Left            =   105
      TabIndex        =   4
      Top             =   3555
      Width           =   3645
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vendedores"
      Height          =   1290
      Left            =   105
      TabIndex        =   11
      Top             =   2130
      Width           =   3840
      Begin VB.OptionButton OptionTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         TabIndex        =   1
         Top             =   255
         Value           =   -1  'True
         Width           =   2490
      End
      Begin VB.OptionButton OptionApenasUm 
         Caption         =   "Apenas um"
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
         Left            =   165
         TabIndex        =   2
         Top             =   585
         Width           =   2550
      End
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   300
         Left            =   1410
         TabIndex        =   3
         Top             =   825
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label LabelVendedor 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Enabled         =   0   'False
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
         Left            =   465
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   885
         Width           =   885
      End
   End
   Begin VB.Label Label2 
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
      Left            =   195
      TabIndex        =   13
      Top             =   285
      Width           =   615
   End
End
Attribute VB_Name = "RelOpComisCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros de uma opcao salva anteriormente e exibe na tela

Dim lErro As Long
Dim sParam As String, sParamdat As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 23159

    'pega vendedor e exibe
    lErro = objRelOpcoes.ObterParametro("TOPVEND", sParam)
    If lErro <> SUCESSO Then Error 23160
    
    'Se a Opção todos Vendedores estiver preenchida
    
    If sParam = "UM" Then
        
        lErro = objRelOpcoes.ObterParametro("TVENDEDOR", sParam)
        If lErro <> SUCESSO Then Error 23161

        OptionApenasUm.Value = True
        Vendedor.Text = sParam
        Call Vendedor_Validate(bSGECancelDummy)

    ElseIf sParam = "Todos" Then
        
        OptionTodos.Value = True
        Vendedor.Enabled = False
            
    End If
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 37768

    Call DateParaMasked(DataInicial, CDate(sParam))
    
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 37769

    Call DateParaMasked(DataFinal, CDate(sParam))
    
    'pega 'Pula página a cada novo vendedor' e exibe
    lErro = objRelOpcoes.ObterParametro("TPULAPAGQBR0", sParam)
    If lErro <> SUCESSO Then Error 23164

    If sParam = "S" Then CheckPulaPag.Value = 1

    'pega 'IncluirPrevisao' e exibe
    lErro = objRelOpcoes.ObterParametro("TINCLUIRPREV", sParam)
    If lErro <> SUCESSO Then Error 23164

    If sParam = "S" Then CheckPrevisao.Value = 1

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 23159, 23160, 23161, 2362, 23163, 23164, 37768, 37769

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutar As Boolean = False) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iPer_I As Integer, iPer_F As Integer
Dim iExercicio As Integer
Dim sCheck As String, sCheck2 As String
Dim sCheckVend As String, sCheckDat As String
Dim sVendedor As String
Dim dtDataInicio As Date
Dim dtDataFim As Date, lNumIntRel As Long

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 23152

    'Pegar parametros da tela
    
    'Se a opção para todos os vendedores estiver selecionada
    If OptionTodos.Value = True Then
        
        sCheckVend = "TODOS"

    Else
        
        'Se a opção para apenas um vendedor estiver selecionada
        If Len(Trim(Vendedor.Text)) = 0 Then Error 23153
        sVendedor = Vendedor.Text
        sCheckVend = "UM"
    
    End If

    lErro = objRelOpcoes.IncluirParametro("TVENDEDOR", sVendedor)
    If lErro <> AD_BOOL_TRUE Then Error 23154

    lErro = objRelOpcoes.IncluirParametro("TOPVEND", sCheckVend)
    If lErro <> AD_BOOL_TRUE Then Error 23155

    'Pula Página a Cada Novo Vendedor
    If CheckPulaPag.Value Then
        sCheck = "S"
    Else
        sCheck = "N"
    End If

    If CheckPrevisao.Value Then
        sCheck2 = "S"
    Else
        sCheck2 = "N"
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 37791
    
    End If
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 37779
    
    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 37780
    
    lErro = objRelOpcoes.IncluirParametro("TPULAPAGQBR0", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 23158

    lErro = objRelOpcoes.IncluirParametro("TINCLUIRPREV", sCheck2)
    If lErro <> AD_BOOL_TRUE Then Error 23158

    If bExecutar Then
    
        If Trim(DataInicial.ClipText) = "" Or Trim(DataFinal.ClipText) = "" Then Error 37791
    
        GL_objMDIForm.MousePointer = vbHourglass
        lErro = CF("RelComisCro_Prepara", lNumIntRel, giFilialEmpresa, CDate(DataInicial.Text), CDate(DataFinal.Text), Codigo_Extrai(sVendedor), CheckPrevisao.Value)
        GL_objMDIForm.MousePointer = vbDefault
        If lErro <> SUCESSO Then Error 23158

        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then Error 23158
    
        lErro = Monta_Expressao_Selecao(objRelOpcoes)
        If lErro <> SUCESSO Then Error 23181
        
    End If
    '
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

         Case 37791
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataInicial.SetFocus
       
        Case 23151, 23152

        Case 23153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_FORNECIDO", Err)

        Case 23154, 23155, 23156, 23157, 23158, 23181, 37779, 37780

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção que será incluida dinamicamente para a execucao do relatorio

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 24976

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche combo com as opções de relatório
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 23174

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 23174

        Case 24976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 23167

    vbMsgRes = Rotina_Aviso(vbYesNo, "EXCLUSAO_RELOPCOMISVEND")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 23168

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 23167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 23168

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then Error 23180
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 23180

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'grava os parametros informados no preenchimento da tela associando-os a um "nome de opção"

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 23169

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23170

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 23171

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 57694

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 23169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 23170, 23171, 57694

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)

    CheckPulaPag.Value = 0
    OptionTodos.Value = True

    Vendedor.Text = ""
    Vendedor.Enabled = False

    ComboOpcoes.SetFocus

End Sub

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_OpcoesRel_Form_Load
    
    Set objEventoVendedor = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Unload Me

    Exit Sub

End Sub

Private Sub LabelVendedor_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_LabelVendedor_Click
    
    If Len(Trim(Vendedor.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

   Exit Sub

Erro_LabelVendedor_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    Vendedor.Text = CStr(objVendedor.iCodigo)
    Call Vendedor_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub OptionApenasUm_Click()

    Vendedor.Enabled = True
    LabelVendedor.Enabled = True

End Sub

Private Sub OptionTodos_Click()

    Vendedor.Text = ""
    Vendedor.Enabled = False
    LabelVendedor.Enabled = False
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoVendedor = Nothing
    
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor, 0)
        If lErro <> SUCESSO Then Error 52918

    End If
    
    OptionTodos.Value = False
    OptionApenasUm.Value = True
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True


    Select Case Err

        Case 52918
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)

    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_COMIS_VEND
    Set Form_Load_Ocx = Me
    Caption = "Cálculo de Comissões"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpComisCro"
    
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
        
        If Me.ActiveControl Is Vendedor Then
            Call LabelVendedor_Click
        End If
    
    End If

End Sub


Private Sub LabelVendedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedor, Source, X, Y)
End Sub

Private Sub LabelVendedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedor, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
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
        If lErro <> SUCESSO Then Error 37795

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 37795

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

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
        If lErro <> SUCESSO Then Error 37796

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 37796

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37797

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 37797
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37798

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 37798
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 37799

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 37799
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 37800

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 37800
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Sub

End Sub

