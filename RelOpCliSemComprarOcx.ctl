VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpCliSemComprarOcx 
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   6150
   Begin VB.Frame AlmoxarifadoInicial 
      Caption         =   "Vendedores"
      Height          =   1155
      Left            =   165
      TabIndex        =   20
      Top             =   1590
      Width           =   5865
      Begin VB.CheckBox QuebraPorVendedor 
         Caption         =   "Quebra por vendedor"
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
         Left            =   120
         TabIndex        =   5
         Top             =   765
         Value           =   1  'Checked
         Width           =   2610
      End
      Begin MSMask.MaskEdBox VendedorInicial 
         Height          =   300
         Left            =   690
         TabIndex        =   3
         Top             =   315
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendedorFinal 
         Height          =   300
         Left            =   3630
         TabIndex        =   4
         Top             =   300
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedorInicial 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   375
         Width           =   570
      End
      Begin VB.Label LabelVendedorFinal 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.Frame FrameTipoCliente 
      Caption         =   "Tipo de Cliente"
      Height          =   1095
      Left            =   150
      TabIndex        =   19
      Top             =   2820
      Width           =   5850
      Begin VB.ComboBox TipoCliente 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   585
         Width           =   2550
      End
      Begin VB.OptionButton TipoClienteApenas 
         Caption         =   "Apenas"
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
         Left            =   180
         TabIndex        =   7
         Top             =   615
         Width           =   1050
      End
      Begin VB.OptionButton TipoClienteTodos 
         Caption         =   "Todos os tipos"
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
         Left            =   180
         TabIndex        =   6
         Top             =   285
         Width           =   1620
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3870
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpCliSemComprarOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpCliSemComprarOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpCliSemComprarOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpCliSemComprarOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Número de dias sem comprar"
      Height          =   720
      Left            =   180
      TabIndex        =   15
      Top             =   765
      Width           =   3360
      Begin MSMask.MaskEdBox DiasDe 
         Height          =   300
         Left            =   660
         TabIndex        =   1
         Top             =   270
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DiasAte 
         Height          =   300
         Left            =   2145
         TabIndex        =   2
         Top             =   285
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
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
         Left            =   1785
         TabIndex        =   17
         Top             =   315
         Width           =   360
      End
      Begin VB.Label Label3 
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
         Left            =   315
         TabIndex        =   16
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpCliSemComprarOcx.ctx":0994
      Left            =   810
      List            =   "RelOpCliSemComprarOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
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
      Left            =   4200
      Picture         =   "RelOpCliSemComprarOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   870
      Width           =   1815
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
      Left            =   120
      TabIndex        =   18
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpCliSemComprarOcx"
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

Dim giVendedorInicial As Integer

Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
       
    Set objEventoVendedor = New AdmEvento
    
    'Carrega a combo Tipo Cliente
    Call Carrega_ComboTipoCliente(TipoCliente)
    
    TipoClienteTodos.Value = True
    TipoCliente.Enabled = False
    QuebraPorVendedor.Value = vbChecked
        
    giVendedorInicial = 1
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141878)

    End Select

    Exit Sub

End Sub


Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 141879
            
    'pega parâmetro Pedido Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NDIASDE", sParam)
    If lErro Then gError 141880
    
    If StrParaInt(sParam) <> 0 Then
        DiasDe.PromptInclude = False
        DiasDe.Text = sParam
        DiasDe.PromptInclude = True
    End If
    
    'pega parâmetro Pedido Final e exibe
    lErro = objRelOpcoes.ObterParametro("NDIASATE", sParam)
    If lErro Then gError 141881
    
    If StrParaInt(sParam) <> 0 Then
        DiasAte.PromptInclude = False
        DiasAte.Text = sParam
        DiasAte.PromptInclude = False
    End If
    
    'pega vendedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDINIC", sParam)
    If lErro Then gError 187971
    
    VendedorInicial.Text = sParam
    Call VendedorInicial_Validate(bSGECancelDummy)
    
    'pega  vendedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDFIM", sParam)
    If lErro Then gError 187972
    
    VendedorFinal.Text = sParam
    Call VendedorFinal_Validate(bSGECancelDummy)
    
    lErro = objRelOpcoes.ObterParametro("TTIPOCLIENTE", sParam)
    If lErro <> SUCESSO Then gError 187977

    'Preenche o tipo
    If sParam = "" Then
        TipoCliente.ListIndex = -1
        TipoCliente.Enabled = False
        TipoClienteTodos.Value = True
    Else
        TipoClienteApenas.Value = True
        TipoCliente.Enabled = True
        Call Combo_Seleciona_ItemData(TipoCliente, Codigo_Extrai((sParam)))
    End If
           
    'pega  vendedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NQUEBRA", sParam)
    If lErro Then gError 187989
    
    If StrParaInt(sParam) = MARCADO Then
        QuebraPorVendedor.Value = vbChecked
    Else
        QuebraPorVendedor.Value = vbUnchecked
    End If
           
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 141878 To 141881, 187971, 187972, 187977, 187989

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141882)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoVendedor = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 141883
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 141884

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 141883
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
            
        Case 141884
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141885)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Function Formata_E_Critica_Parametros(sVend_I As String, sVend_F As String, sTipoCliente As String, sQuebra As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'Dia inicial não pode ser maior que o dia final
    If StrParaInt(DiasDe.Text) <> 0 And StrParaInt(DiasAte.Text) <> 0 Then
    
         If StrParaInt(DiasDe.Text) > StrParaInt(DiasAte.Text) Then gError 141886
         
    End If
    
    If VendedorInicial.Text <> "" Then
        sVend_I = CStr(Codigo_Extrai(VendedorInicial.Text))
    Else
        sVend_I = ""
    End If
    
    If VendedorFinal.Text <> "" Then
        sVend_F = CStr(Codigo_Extrai(VendedorFinal.Text))
    Else
        sVend_F = ""
    End If
            
    If sVend_I <> "" And sVend_F <> "" Then
        
        If CInt(sVend_I) > CInt(sVend_F) Then gError 187970
        
    End If
    
    'Se a opção para todos os tipos estiver selecionada
    If TipoClienteTodos.Value = True Then
        sTipoCliente = ""
    Else
        If TipoCliente.Text = "" Then gError 187976
        sTipoCliente = TipoCliente.Text
    End If
    
    If QuebraPorVendedor.Value = vbChecked Then
        sQuebra = CStr(MARCADO)
    Else
        sQuebra = CStr(DESMARCADO)
    End If
       
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
  
        Case 141886
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DiasDe.SetFocus
            
        Case 187970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", gErr)
            VendedorInicial.SetFocus
            
        Case 187976
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO1", gErr)
            TipoCliente.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141887)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 141887
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus

    TipoClienteTodos.Value = True
    TipoCliente.Enabled = False
    
    QuebraPorVendedor.Value = vbChecked
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 141887
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141888)

    End Select

    Exit Sub
    
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sVend_I As String
Dim sVend_F As String
Dim sTipoCliente As String
Dim sQuebra As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sVend_I, sVend_F, sTipoCliente, sQuebra)
    If lErro <> SUCESSO Then gError 141889
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 141890
                
    lErro = objRelOpcoes.IncluirParametro("NDIASDE", DiasDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 141891
    
    lErro = objRelOpcoes.IncluirParametro("NDIASATE", DiasAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 141892
    
    lErro = objRelOpcoes.IncluirParametro("NVENDINIC", sVend_I)
    If lErro <> AD_BOOL_TRUE Then gError 187973

    lErro = objRelOpcoes.IncluirParametro("NVENDFIM", sVend_F)
    If lErro <> AD_BOOL_TRUE Then gError 187974
    
    lErro = objRelOpcoes.IncluirParametro("TTIPOCLIENTE", sTipoCliente)
    If lErro <> AD_BOOL_TRUE Then gError 187975
    
    lErro = objRelOpcoes.IncluirParametro("TVENDEDORDE", VendedorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 187978
    
    lErro = objRelOpcoes.IncluirParametro("TVENDEDORATE", VendedorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 187979
    
    lErro = objRelOpcoes.IncluirParametro("NQUEBRA", sQuebra)
    If lErro <> AD_BOOL_TRUE Then gError 187988
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sVend_I, sVend_F, sTipoCliente)
    If lErro <> SUCESSO Then gError 141893

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 141889 To 141893, 187973, 187974, 187975, 187978, 187979, 187988

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141894)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 141895

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 141896

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 141897
    
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 141895
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 141896, 141897

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141898)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 141899
    
    If QuebraPorVendedor.Value = vbChecked Then
        gobjRelatorio.sNomeTsk = "CLIFATQ"
    Else
        gobjRelatorio.sNomeTsk = "CLIFAT"
    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 141899

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141900)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 141901

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 141902

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 141903

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 141904
    
    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 141901
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 141902 To 141904

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141905)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, ByVal sVend_I As String, ByVal sVend_F As String, ByVal sTipoCliente As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sVend_I <> "" Then sExpressao = "Vendedor >= " & Forprint_ConvInt(CInt(sVend_I))

   If sVend_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vendedor <= " & Forprint_ConvInt(CInt(sVend_F))

    End If
    
    'Se a opção para apenas um tipo estiver selecionada
    If sTipoCliente <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoCliente = " & Forprint_ConvInt(Codigo_Extrai(sTipoCliente))

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141905)

    End Select

    Exit Function

End Function

Private Sub DiasDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DiasDe_Validate

    If Len(Trim(DiasDe.Text)) > 0 Then
        
        lErro = Valor_Positivo_Critica(DiasDe.Text)
        If lErro <> SUCESSO Then gError 141906
    
    End If
       
    Exit Sub

Erro_DiasDe_Validate:

    Cancel = True

    Select Case gErr
    
        Case 141906
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141907)

    End Select

    Exit Sub

End Sub

Private Sub DiasAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DiasAte_Validate

    If Len(Trim(DiasAte.Text)) > 0 Then
        
        lErro = Valor_Positivo_Critica(DiasAte.Text)
        If lErro <> SUCESSO Then gError 141908
    
    End If
       
    Exit Sub

Erro_DiasAte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 141908
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141909)

    End Select

    Exit Sub

End Sub
'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_DESPACHO
    Set Form_Load_Ocx = Me
    Caption = "Clientes sem Comprar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpCliSemComprar"
    
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
    
    End If

End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorFinal_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 2
    
    If Len(Trim(VendedorFinal.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorFinal.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub LabelVendedorInicial_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 1
    
    If Len(Trim(VendedorInicial.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorInicial.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim bCancel As Boolean

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    If giVendedorInicial = 1 Then
        VendedorInicial.Text = CStr(objVendedor.iCodigo)
        VendedorInicial_Validate (bCancel)
    Else
        VendedorFinal.Text = CStr(objVendedor.iCodigo)
        VendedorFinal_Validate (bCancel)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub VendedorInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorInicial_Validate

    If Len(Trim(VendedorInicial.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorInicial, objVendedor, 0)
        If lErro <> SUCESSO Then gError 187980

    End If
    
    giVendedorInicial = 1
    
    Exit Sub

Erro_VendedorInicial_Validate:

    Cancel = True
    
    Select Case gErr

        Case 187980
             Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187981)

    End Select

End Sub

Private Sub VendedorFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorFinal_Validate

    If Len(Trim(VendedorFinal.Text)) > 0 Then

        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorFinal, objVendedor, 0)
        If lErro <> SUCESSO Then gError 187982

    End If
    
    giVendedorInicial = 0
 
    Exit Sub

Erro_VendedorFinal_Validate:

    Cancel = True
    
    Select Case gErr

        Case 187982
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187983)

    End Select

End Sub

Private Sub LabelVendedorFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorFinal, Source, X, Y)
End Sub

Private Sub LabelVendedorFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorInicial, Source, X, Y)
End Sub

Private Sub LabelVendedorInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorInicial, Button, Shift, X, Y)
End Sub

Private Sub TipoClienteTodos_Click()

Dim lErro As Long

On Error GoTo Erro_TipoClienteTodos_Click

    'Desabilita o combotipo
    TipoCliente.ListIndex = -1
    TipoCliente.Enabled = False

    Exit Sub

Erro_TipoClienteTodos_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187984)

    End Select

    Exit Sub

End Sub

Private Sub TipoClienteApenas_Click()

Dim lErro As Long

On Error GoTo Erro_TipoClienteApenas_Click

    'Habilita a ComboTipo
    TipoCliente.Enabled = True

    Exit Sub

Erro_TipoClienteApenas_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187985)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_ComboTipoCliente(ByVal objComboBox As ComboBox)

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_ComboTipoCliente

    'Lê cada código e descrição da tabela TiposDeCliente
    lErro = CF("Cod_Nomes_Le", "TiposDeCliente", "Codigo", "Descricao", STRING_TIPO_CLIENTE_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 187986

    'Preenche a ComboBox Tipo com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        objComboBox.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoDescricao.iCodigo
    Next

    Exit Sub

Erro_Carrega_ComboTipoCliente:

    Select Case gErr

        Case 187986

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187987)

    End Select

    Exit Sub

End Sub
