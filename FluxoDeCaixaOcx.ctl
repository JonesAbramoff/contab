VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl FluxoDeCaixaOcx 
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   LockControls    =   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   7335
   Begin VB.CommandButton BotaoEditarDados 
      Caption         =   "Exibir Fluxo de Caixa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3907
      TabIndex        =   4
      Top             =   2325
      Width           =   2160
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4965
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FluxoDeCaixaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "FluxoDeCaixaOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "FluxoDeCaixaOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "FluxoDeCaixaOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Identificacao 
      Height          =   315
      Left            =   1425
      TabIndex        =   0
      Top             =   390
      Width           =   2790
   End
   Begin VB.TextBox Descricao 
      Height          =   300
      Left            =   1425
      TabIndex        =   1
      Top             =   870
      Width           =   3420
   End
   Begin VB.CommandButton Botao_Atualizar 
      Caption         =   "Atualizar Dados Reais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1267
      TabIndex        =   3
      Top             =   2325
      Width           =   2160
   End
   Begin MSMask.MaskEdBox DataFinal 
      Height          =   300
      Left            =   5670
      TabIndex        =   2
      Top             =   1320
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   300
      Left            =   6825
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1320
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label DataDadosReais 
      Height          =   240
      Left            =   2790
      TabIndex        =   11
      Top             =   1890
      Width           =   885
   End
   Begin VB.Label DataInicial 
      Height          =   240
      Left            =   1425
      TabIndex        =   12
      Top             =   1380
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data Base:"
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
      Left            =   375
      TabIndex        =   13
      Top             =   1365
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Identificação:"
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
      Left            =   150
      TabIndex        =   14
      Top             =   420
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Left            =   405
      TabIndex        =   15
      Top             =   900
      Width           =   930
   End
   Begin VB.Label Label3 
      Caption         =   "Mostrar previsão até:"
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
      Left            =   3735
      TabIndex        =   16
      Top             =   1380
      Width           =   1860
   End
   Begin VB.Label Label7 
      Caption         =   "Dados Reais Atualizados até:"
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
      Left            =   120
      TabIndex        =   17
      Top             =   1890
      Width           =   2550
   End
End
Attribute VB_Name = "FluxoDeCaixaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim sFluxo As String
Dim lFluxo1 As Long

Public Sub Aplic_Aplicacao_Click()

Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Aplic_Aplicacao_Click

    lErro = Carrega_Tela("FluxoAplic")
    If lErro <> SUCESSO Then Error 21226

    Exit Sub

Erro_Aplic_Aplicacao_Click:

    Select Case Err

        Case 21226

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160365)

    End Select

    Exit Sub

End Sub

Public Sub Aplic_TipoAplicacao_Click()

Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Aplic_TipoAplicacao_Click

    lErro = Carrega_Tela("FluxoTipoAplic")
    If lErro <> SUCESSO Then Error 21235

    Exit Sub

Erro_Aplic_TipoAplicacao_Click:

    Select Case Err

        Case 21235

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160366)

    End Select

    Exit Sub

End Sub

Private Sub Botao_Atualizar_Click()

Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Botao_Atualizar_Click


    GL_objMDIForm.MousePointer = vbHourglass

    lErro = MoveDadosTela_Variaveis(objFluxo)
    If lErro <> SUCESSO Then Error 21238

    lErro = CF("Fluxo_Le", objFluxo)
    If lErro <> SUCESSO And lErro <> 20104 Then Error 21242

    If lErro = 20104 Then Error 21243

    objFluxo.dtDataDadosReais = gdtDataAtual

    lErro = CF("Fluxo_ObterDadosReais", objFluxo)
    If lErro <> SUCESSO Then Error 21241

    DataDadosReais.Caption = Format(gdtDataAtual, "dd/mm/yyyy")

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_Botao_Atualizar_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 21238, 21241, 21242

       Case 21243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_NAO_CADASTRADO", Err, objFluxo.sFluxo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160367)

    End Select

    Exit Sub

End Sub

Private Sub BotaoEditarDados_Click()

        Set PopUpMenuFluxo.objTela = Me
        PopupMenu PopUpMenuFluxo.FluxoDeCaixa
        Set PopUpMenuFluxo.objTela = Nothing
        
End Sub

Private Sub BotaoExcluir_Click()

Dim lErro  As Integer
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica se a identificacao de um fluxo foi fornecida
    If Len(Identificacao.Text) = 0 Then Error 20100

    'Pede a confirmacao da exclusao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_FLUXO", Identificacao.Text)

    If vbMsgRes = vbYes Then

        objFluxo.sFluxo = Identificacao.Text

        'Chama a rotina de exclusao
        lErro = CF("Fluxo_Exclui", objFluxo)
        If lErro <> SUCESSO Then Error 20107

        For iIndice = 0 To Identificacao.ListCount - 1
            If Identificacao.Text = Identificacao.List(iIndice) Then
                Identificacao.RemoveItem (iIndice)
                Exit For
            End If
        Next

        Call Limpa_Tela_FluxoDeCaixa

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 20100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_NAO_PREENCHIDO", Err)

        Case 20107

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160368)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a rotina de gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 20096

    Call Limpa_Tela_FluxoDeCaixa

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 20096

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160369)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Confirma o pedido de limpeza da tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 20097

    'Limpa a tela
    Call Limpa_Tela_FluxoDeCaixa

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 20097

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160370)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_FluxoDeCaixa()

    Call Limpa_Tela(Me)

    Identificacao.Text = ""
    DataDadosReais.Caption = ""
    DataInicial.Caption = Format(gdtDataAtual, "dd/mm/yyyy")

End Sub

Private Sub DataFinal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal, iAlterado)

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colFluxo As New Collection
Dim objFluxo As ClassFluxo

On Error GoTo Erro_Form_Load

    sFluxo = ""

    lErro = CF("Fluxo_Le_Todos", colFluxo)
    If lErro <> SUCESSO Then Error 10914

    For Each objFluxo In colFluxo

        Identificacao.AddItem objFluxo.sFluxo
        Identificacao.ItemData(Identificacao.NewIndex) = objFluxo.lFluxoId

    Next

    DataInicial.Caption = Format(gdtDataAtual, "dd/mm/yy")
    DataFinal.Text = Format(gdtDataAtual, "dd/mm/yy")

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 10914

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160371)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    'verifica se a data está preenchida
    If Len(Trim(DataFinal.ClipText)) > 0 Then

        'verifica se a data final é válida
        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then Error 10920

        'verifica se a data final é menor que a data inicial. Se for ==> erro.
        If CDate(DataFinal.Text) < CDate(DataInicial.Caption) Then Error 10921

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 10920

        Case 10921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_DATAINI_MAIOR_DATAFIM", Err, DataFinal.Text, DataInicial.Caption)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160372)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub Identificacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Identificacao_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Identificacao_Click

    If Identificacao.ListIndex = -1 Then Exit Sub

    'Verifica se existe a necessidade de salvar os dados do fluxo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 10924

    'Pega o nome do fluxo atual
    sFluxo = Identificacao.Text

    'Exibe na tela os dados do fluxo
    lErro = Traz_Fluxo_Tela(sFluxo)
    If lErro <> SUCESSO Then Error 10925

    iAlterado = 0

    Exit Sub

Erro_Identificacao_Click:

    Select Case Err

        Case 10924
            Identificacao.Text = sFluxo

        Case 10925

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160373)

    End Select

    Exit Sub

End Sub

Private Sub Identificacao_Validate(Cancel As Boolean)

    If Identificacao.ListIndex = -1 Then DataInicial.Caption = Format(gdtDataAtual, "dd/mm/yyyy")

End Sub

Public Sub Pag_Fornecedor_Click()
Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Pag_Fornecedor_Click

    lErro = Carrega_Tela("FluxoPagForn")
    If lErro <> SUCESSO Then Error 21229

    Exit Sub

Erro_Pag_Fornecedor_Click:

    Select Case Err

        Case 21229

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160374)

    End Select

    Exit Sub

End Sub

Public Sub Pag_TipoFornecedor_Click()
Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Pag_TipoFornecedor_Click

    lErro = Carrega_Tela("FluxoTipoForn")
    If lErro <> SUCESSO Then Error 21227

    Exit Sub

Erro_Pag_TipoFornecedor_Click:

    Select Case Err

        Case 21227

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160375)

    End Select

    Exit Sub

End Sub

Public Sub Pag_Titulo_Click()
Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Pag_Titulo_Click

    lErro = Carrega_Tela("FluxoPag")
    If lErro <> SUCESSO Then Error 21230

    Exit Sub

Erro_Pag_Titulo_Click:

    Select Case Err

        Case 21230

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160376)

    End Select

    Exit Sub

End Sub

Public Sub Rec_Cliente_Click()
Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Rec_Cliente_Click

    lErro = Carrega_Tela("FluxoRecebCli")
    If lErro <> SUCESSO Then Error 21231

    Exit Sub

Erro_Rec_Cliente_Click:

    Select Case Err

        Case 21231

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160377)

    End Select

    Exit Sub

End Sub

Public Sub Rec_TipoCliente_Click()

Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_TipoCliente_Click

    lErro = Carrega_Tela("FluxoTipoCli")
    If lErro <> SUCESSO Then Error 21232

    Exit Sub

Erro_TipoCliente_Click:

    Select Case Err

        Case 21232

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160378)

    End Select

    Exit Sub

End Sub

Public Sub Rec_Titulo_Click()
Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Rec_Titulo_Click

    lErro = Carrega_Tela("FluxoReceb")
    If lErro <> SUCESSO Then Error 21233

    Exit Sub

Erro_Rec_Titulo_Click:

    Select Case Err

        Case 21233

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160379)

    End Select

    Exit Sub

End Sub

Public Sub Saldos_Iniciais_Click()
Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Saldos_Iniciais_Click

    lErro = Carrega_Tela("FluxoSaldoInicial")
    If lErro <> SUCESSO Then Error 21234

    Exit Sub

Erro_Saldos_Iniciais_Click:

    Select Case Err

        Case 21234

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160380)

    End Select

    Exit Sub

End Sub

Public Sub Sint_Projecao_Click()
Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Sint_Projecao_Click

    lErro = Carrega_Tela("FluxoSintProj")
    If lErro <> SUCESSO Then Error 21236

    Exit Sub

Erro_Sint_Projecao_Click:

    Select Case Err

        Case 21236

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160381)

    End Select

    Exit Sub

End Sub

Public Sub Sint_Revisao_Click()
Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Sint_Revisao_Click

    lErro = Carrega_Tela("FluxoSintRev")
    If lErro <> SUCESSO Then Error 21237

    Exit Sub

Erro_Sint_Revisao_Click:

    Select Case Err

        Case 21237

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160382)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        If CDate(DataFinal.Text) <= CDate(DataInicial.Caption) Then
            DataFinal.Text = Format(DataInicial.Caption, "dd/mm/yy")
        Else

            sData = DataFinal.Text

            lErro = Data_Diminui(sData)
            If lErro <> SUCESSO Then Error 10923

            DataFinal.Text = sData

        End If

    End If

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 10923

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160383)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 10922

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 10922

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160384)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'grava os dados do fluxo em questão

Dim lErro As Long
Dim objFluxo As New ClassFluxo
Dim iIndice As Integer
Dim iAchou As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'move os dados da tela para as variáveis
    lErro = MoveDadosTela_Variaveis(objFluxo)
    If lErro <> SUCESSO Then Error 10929

    If Len(Descricao.Text) = 0 Then Error 21274

    'grava o fluxo de caixa
    lErro = CF("Fluxo_Grava", objFluxo)
    If lErro <> SUCESSO Then Error 10930

    iAchou = 0

    'se o fluxo ainda não estava cadastrado na combo, insere.
    For iIndice = 0 To Identificacao.ListCount - 1
        If Identificacao.List(iIndice) = Identificacao.Text Then
            iAchou = 1
            Exit For
        End If
    Next

    'insere o fluxo na combo
    If iAchou = 0 Then
        Identificacao.AddItem Identificacao.Text
        Identificacao.ItemData(Identificacao.NewIndex) = objFluxo.lFluxoId
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 10929, 10930

        Case 21274
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160385)

    End Select

    Exit Function

End Function

Function Carrega_Tela(sNomeTela As String) As Long

Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Carrega_Tela

    lErro = MoveDadosTela_Variaveis(objFluxo)
    If lErro <> SUCESSO Then Error 21228

    lErro = CF("Fluxo_Le", objFluxo)
    If lErro <> SUCESSO And lErro <> 20104 Then Error 21239

    If lErro = 20104 Then Error 21240

    objFluxo.dtData = objFluxo.dtDataBase

    Call Chama_Tela(sNomeTela, objFluxo)

    Exit Function

Erro_Carrega_Tela:

    Select Case Err

        Case 21228, 21239

        Case 21240
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_NAO_CADASTRADO", Err, objFluxo.sFluxo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160386)

    End Select

    Exit Function

End Function

Function MoveDadosTela_Variaveis(objFluxo As ClassFluxo) As Long
'Move os dados do fluxo da tela para objFluxo

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_MoveDadosTela_Variaveis

    If Len(Identificacao.Text) = 0 Then Error 10927

    If Len(DataFinal.ClipText) = 0 Then Error 10928

    objFluxo.lFluxoId = lFluxo1
    objFluxo.sFluxo = Identificacao.Text
    objFluxo.sDescricao = Descricao.Text
    objFluxo.dtDataBase = CDate(DataInicial.Caption)
    objFluxo.dtDataFinal = CDate(DataFinal.Text)
    objFluxo.iFilialEmpresa = giFilialEmpresa

    MoveDadosTela_Variaveis = SUCESSO

    Exit Function

Erro_MoveDadosTela_Variaveis:

    MoveDadosTela_Variaveis = Err

    Select Case Err

        Case 10927
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_FLUXO_VAZIO", Err)

        Case 10928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_FLUXO_VAZIO", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160387)

    End Select

    Exit Function

End Function

Private Function Traz_Fluxo_Tela(sFluxo As String) As Long
'Coloca na Tela os dados do Fluxo passado como parametro

Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Traz_Fluxo_Tela

    objFluxo.sFluxo = String(STRING_FLUXO_NOME, 0)
    objFluxo.sFluxo = sFluxo
    objFluxo.iFilialEmpresa = giFilialEmpresa

    'Le o fluxo passado como parametro
    lErro = CF("Fluxo_Le", objFluxo)
    If lErro <> SUCESSO And lErro <> 20104 Then Error 20105

    If lErro = 20104 Then Error 20106

    'passa os dados para a Tela
    Descricao.Text = objFluxo.sDescricao
    DataInicial.Caption = Format(objFluxo.dtDataBase, "dd/mm/yyyy")
    DataFinal.Text = Format(objFluxo.dtDataFinal, "dd/mm/yy")
    If objFluxo.dtDataDadosReais = DATA_NULA Then
        DataDadosReais.Caption = ""
    Else
        DataDadosReais.Caption = Format(objFluxo.dtDataDadosReais, "dd/mm/yyyy")
    End If
    lFluxo1 = objFluxo.lFluxoId

    Traz_Fluxo_Tela = SUCESSO

    Exit Function

Erro_Traz_Fluxo_Tela:

    Traz_Fluxo_Tela = Err

    Select Case Err

        Case 20105

        Case 20106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FLUXO_NAO_CADASTRADO", Err, objFluxo.sFluxo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 160388)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FLUXO_CAIXA
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FluxoDeCaixa"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

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

Private Sub Unload(objme As Object)
    
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



Private Sub DataDadosReais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataDadosReais, Source, X, Y)
End Sub

Private Sub DataDadosReais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataDadosReais, Button, Shift, X, Y)
End Sub

Private Sub DataInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataInicial, Source, X, Y)
End Sub

Private Sub DataInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataInicial, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

