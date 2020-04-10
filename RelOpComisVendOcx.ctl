VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpComisVendOcx 
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   6390
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpComisVendOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpComisVendOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpComisVendOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpComisVendOcx.ctx":0816
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
      Left            =   4200
      Picture         =   "RelOpComisVendOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   810
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpComisVendOcx.ctx":0A96
      Left            =   870
      List            =   "RelOpComisVendOcx.ctx":0A98
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
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   3225
      Width           =   3645
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vendedores"
      Height          =   1290
      Left            =   120
      TabIndex        =   18
      Top             =   645
      Width           =   3720
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
         TabIndex        =   19
         Top             =   885
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comissões"
      Height          =   1155
      Left            =   120
      TabIndex        =   15
      Top             =   1965
      Width           =   3720
      Begin VB.OptionButton OptionGerada 
         Caption         =   "geradas até:"
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
         Left            =   255
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.OptionButton OptionBaixa 
         Caption         =   "Baixadas em:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   255
         TabIndex        =   6
         Top             =   765
         Width           =   1545
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   2955
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   735
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
      End
      Begin MSMask.MaskEdBox DataGeradas 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
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
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   2955
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataBaixadas 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
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
      TabIndex        =   20
      Top             =   285
      Width           =   615
   End
End
Attribute VB_Name = "RelOpComisVendOcx"
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

Function Critica_Datas_ComisVend() As Long
'Faz a crítica da data geração ou da data de baixa

Dim lErro As Long

On Error GoTo Erro_Critica_Datas_ComisVend

    'Se data de geração estiver selecionada
    If OptionGerada.Value = True Then

        'data de geração não pode ser vazia
        If Len(Trim(DataGeradas.ClipText)) = 0 Then Error 23148

    Else
        'data de baixa não pode ser vazia
        If Len(Trim(DataBaixadas.ClipText)) = 0 Then Error 23149

    End If

    Critica_Datas_ComisVend = SUCESSO

    Exit Function

Erro_Critica_Datas_ComisVend:

    Critica_Datas_ComisVend = Err

    Select Case Err

        Case 23148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err, Error$)
            DataGeradas.SetFocus

        Case 23149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)
            DataBaixadas.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167640)

    End Select

    Exit Function

End Function

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
    
    'Pega data e exibe
    lErro = objRelOpcoes.ObterParametro("TOPDAT", sParam)
    If lErro <> SUCESSO Then Error 23162
    
    lErro = objRelOpcoes.ObterParametro("TDATA", sParamdat)
    If lErro <> SUCESSO Then Error 23163
    
    'Se a opção for gerada
    If sParam = "GERADA" Then
        
        OptionGerada.Value = True
        DataGeradas.Text = sParamdat
        
    ElseIf sParam = "BAIXADA" Then
        
        'Se a opção for Baixada
        OptionBaixa.Value = True
        DataBaixadas.Text = sParamdat

    End If

    'pega 'Pula página a cada novo vendedor' e exibe
    lErro = objRelOpcoes.ObterParametro("TPULAPAGQBR0", sParam)
    If lErro <> SUCESSO Then Error 23164

    If sParam = "S" Then CheckPulaPag.Value = 1

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 23159, 23160, 23161, 2362, 23163, 23164

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167641)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iPer_I As Integer, iPer_F As Integer
Dim iExercicio As Integer
Dim sCheck As String
Dim sCheckVend As String, sCheckDat As String
Dim sData As String
Dim sVendedor As String

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Datas_ComisVend
    If lErro <> SUCESSO Then Error 23151

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

    'Se a opção geradas até estiver selecionada
    If OptionGerada.Value = True Then
            
            sData = DataGeradas.Text
            sCheckDat = "GERADA"
    Else
        
        'Se a opção para baixadas estiver selecionada
        sData = DataBaixadas.Text
        sCheckDat = "BAIXADA"
    
    End If

    lErro = objRelOpcoes.IncluirParametro("TDATA", sData)
    If lErro <> AD_BOOL_TRUE Then Error 23156

    lErro = objRelOpcoes.IncluirParametro("TOPDAT", sCheckDat)
    If lErro <> AD_BOOL_TRUE Then Error 23157

    'Pula Página a Cada Novo Vendedor
    If CheckPulaPag.Value Then
        sCheck = "S"
    Else
        sCheck = "N"
    End If

    lErro = objRelOpcoes.IncluirParametro("TPULAPAGQBR0", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 23158

    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 23181

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 23151, 23152

        Case 23153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_FORNECIDO", Err)

        Case 23154, 23155, 23156, 23157, 23158, 23181

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 167642)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção que será incluida dinamicamente para a execucao do relatorio

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    'Se a opção para apenas um vendedor estiver selecionada
    If OptionApenasUm.Value = True Then

        If Len(Trim(Vendedor.Text)) > 0 Then sExpressao = "Vendedor = " & Forprint_ConvInt(Codigo_Extrai(Vendedor.Text))

    End If

    'Se a opção data de geracao está selecionada
    If OptionGerada.Value = True Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataGeracao <= " & Forprint_ConvData(CDate(DataGeradas.Text))

    'Se a opção data de baixa está selecionada
    Else
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataBaixa = " & Forprint_ConvData(CDate(DataBaixadas.Text))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167643)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167644)

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

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167645)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23180
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 23180

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167646)

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

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167647)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)

    CheckPulaPag.Value = 0
    OptionTodos.Value = True
    OptionGerada.Value = True
    Vendedor.Text = ""
    Vendedor.Enabled = False
    DataBaixadas.Enabled = False

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

Private Sub DataBaixadas_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataBaixadas)

End Sub

Private Sub DataGeradas_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataGeradas)

End Sub

Private Sub DataGeradas_Validate(Cancel As Boolean)

Dim sDataGeradas As String
Dim lErro As Long

On Error GoTo Erro_DataGeradas_Validate

    If Len(DataGeradas.ClipText) > 0 Then

        sDataGeradas = DataGeradas.Text
        
        lErro = Data_Critica(sDataGeradas)
        If lErro <> SUCESSO Then Error 23165

    End If

    Exit Sub

Erro_DataGeradas_Validate:

    Cancel = True


    Select Case Err

        Case 23165

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167648)

    End Select

    Exit Sub

End Sub

Private Sub DataBaixadas_Validate(Cancel As Boolean)

Dim sDataBaixa As String
Dim lErro As Long

On Error GoTo Erro_DataBaixadas_Validate

    If Len(DataBaixadas.ClipText) > 0 Then

        sDataBaixa = DataBaixadas.Text
        
        lErro = Data_Critica(sDataBaixa)
        If lErro <> SUCESSO Then Error 23166

    End If

    Exit Sub

Erro_DataBaixadas_Validate:

    Cancel = True


    Select Case Err

        Case 23166

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167649)

    End Select

    Exit Sub

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167650)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167651)

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

Private Sub OptionBaixa_Click()

    DataGeradas.Enabled = False
    DataBaixadas.Enabled = True
    UpDown1.Enabled = False
    UpDown2.Enabled = True

End Sub

Private Sub OptionGerada_Click()

    DataBaixadas.Enabled = False
    DataGeradas.Enabled = True
    UpDown1.Enabled = True
    UpDown2.Enabled = False

End Sub

Private Sub OptionTodos_Click()

    Vendedor.Text = ""
    Vendedor.Enabled = False
    LabelVendedor.Enabled = False
    
End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataGeradas, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23176

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 23176
            DataGeradas.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167652)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataGeradas, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23177

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 23177
            DataGeradas.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167653)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataBaixadas, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23178

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 23178
            DataBaixadas.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167654)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataBaixadas, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23179

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 23179
            DataBaixadas.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167655)

    End Select

    Exit Sub

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 167656)

    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_COMIS_VEND
    Set Form_Load_Ocx = Me
    Caption = "Comissões de Vendedores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpComisVend"
    
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

