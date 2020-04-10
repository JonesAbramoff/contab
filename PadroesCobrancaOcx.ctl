VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl PadroesCobrancaOcx 
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   7065
   Begin VB.Frame Frame3 
      Caption         =   "Juros"
      Height          =   645
      Left            =   180
      TabIndex        =   15
      Top             =   3840
      Width           =   4215
      Begin MSMask.MaskEdBox JurosMensais 
         Height          =   300
         Left            =   1140
         TabIndex        =   8
         Top             =   240
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "#0.#0\%"
         PromptChar      =   "_"
      End
      Begin VB.Label JurosDiarios 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3030
         TabIndex        =   18
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Mensais:"
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
         Left            =   240
         TabIndex        =   19
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label10 
         Caption         =   "Diários:"
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
         Left            =   2160
         TabIndex        =   20
         Top             =   270
         Width           =   765
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1710
      Picture         =   "PadroesCobrancaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   270
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4725
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PadroesCobrancaOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "PadroesCobrancaOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "PadroesCobrancaOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "PadroesCobrancaOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CheckBox Inativo 
      Caption         =   "Inativo"
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
      Left            =   2535
      TabIndex        =   2
      Top             =   292
      Width           =   930
   End
   Begin VB.Frame Frame1 
      Caption         =   "Instrução Primária"
      Height          =   1245
      Left            =   180
      TabIndex        =   17
      Top             =   1245
      Width           =   4215
      Begin VB.ComboBox Instrucao1 
         Height          =   315
         Left            =   465
         TabIndex        =   4
         Top             =   375
         Width           =   3345
      End
      Begin MSMask.MaskEdBox DiasDeProtesto1 
         Height          =   300
         Left            =   3315
         TabIndex        =   5
         Top             =   817
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   4
         Mask            =   "9999"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDiasProtesto1 
         Caption         =   "Dias para Devolução / Protesto:"
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
         Height          =   255
         Left            =   450
         TabIndex        =   21
         Top             =   840
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Instrução Secundária"
      Height          =   1245
      Left            =   180
      TabIndex        =   16
      Top             =   2550
      Width           =   4215
      Begin VB.ComboBox Instrucao2 
         Height          =   315
         Left            =   465
         TabIndex        =   6
         Top             =   390
         Width           =   3360
      End
      Begin MSMask.MaskEdBox DiasDeProtesto2 
         Height          =   300
         Left            =   3330
         TabIndex        =   7
         Top             =   832
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   4
         Mask            =   "9999"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDiasProtesto2 
         Caption         =   "Dias para Devolução / Protesto:"
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
         Height          =   255
         Left            =   465
         TabIndex        =   22
         Top             =   855
         Width           =   2775
      End
   End
   Begin VB.ListBox Padroes 
      Height          =   3375
      Left            =   4620
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1110
      Width           =   2265
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   255
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1185
      TabIndex        =   3
      Top             =   795
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
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
      Left            =   180
      TabIndex        =   23
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Left            =   435
      TabIndex        =   24
      Top             =   315
      Width           =   675
   End
   Begin VB.Label Label13 
      Caption         =   "Padrões de Cobrança"
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
      Left            =   4635
      TabIndex        =   25
      Top             =   885
      Width           =   1905
   End
End
Attribute VB_Name = "PadroesCobrancaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iAlterado As Integer

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim sCodigo As String

On Error GoTo Erro_BotaoProxNum_Click

    'gera codigo automatico do proximo Padrao Cobranca
    lErro = CF("PadraoCobranca_Automatico",sCodigo)
    If lErro <> SUCESSO Then Error 57551

    Codigo.Text = sCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57551 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164104)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer
Dim iEncontrou As Integer
Dim objPadraoCobranca As New ClassPadraoCobranca

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 16585

    objPadraoCobranca.iCodigo = CInt(Codigo.Text)
    
    'Verifica se está no BD o Padrão de Cobrança que será excluido
    lErro = CF("PadraoCobranca_Le",objPadraoCobranca)
    If lErro <> SUCESSO And lErro <> 19298 Then Error 16587
    
    'Se não encontrou --> ERRO
    If lErro = 19298 Then Error 16586
    
    'pedido de confirmacao de exclusao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PADRAO_COBRANCA", objPadraoCobranca.iCodigo)

    If vbMsgRes = vbYes Then

        'exclui Padrao Cobranca
        lErro = CF("PadraoCobranca_Exclui",objPadraoCobranca)
        If lErro <> SUCESSO Then Error 16588
        
        'funcao que exclui o padrao da lista
        Call ListaPadraoCobranca_Exclui(objPadraoCobranca.iCodigo)
        
        'Limpa a Tela
        Call Limpa_Tela_PadraoCobranca

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_NAO_CADASTRADO", Err, objPadraoCobranca.sDescricao)
        
        Case 16587, 16588, 16589 'Tratados na Rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164105)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Padrão de Cobrança
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 16562
    
    Call Limpa_Tela_PadraoCobranca
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 16562 'Tratado na Rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164106)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 16605

    Call Limpa_Tela_PadraoCobranca

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 16605 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164107)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) > 0 Then
        
        'Testa se é inteiro
        lErro = Inteiro_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 16531

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 16531

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164108)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Private Sub DiasDeProtesto1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DiasDeProtesto1_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DiasDeProtesto1, iAlterado)

End Sub

Private Sub DiasDeProtesto1_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DiasDeProtesto1_Validate

    'Verifica preenchimento de DiasDeProtesto1
    If Len(Trim(DiasDeProtesto1.Text)) > 0 Then
        
        'Testa se é Inteiro
        lErro = Inteiro_Critica(DiasDeProtesto1.Text)
        If lErro <> SUCESSO Then Error 16559

    End If

    Exit Sub

Erro_DiasDeProtesto1_Validate:

    Cancel = True


    Select Case Err

        Case 16559

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164109)

    End Select

    Exit Sub

End Sub

Private Sub DiasDeProtesto2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DiasDeProtesto2_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DiasDeProtesto2, iAlterado)

End Sub

Private Sub DiasDeProtesto2_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DiasDeProtesto2_Validate

    'Verifica preenchimento de DiasDeProtesto2
    If Len(Trim(DiasDeProtesto2.Text)) > 0 Then
    
        'Testa se é Inteiro
        lErro = Inteiro_Critica(DiasDeProtesto2.Text)
        If lErro <> SUCESSO Then Error 16560

    End If

    Exit Sub

Erro_DiasDeProtesto2_Validate:

    Cancel = True


    Select Case Err

        Case 16560

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164110)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long, sTexto As String
Dim colCodigoDescricao As AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome

On Error GoTo Erro_Form_Load

    Set colCodigoDescricao = New AdmColCodigoNome

    'leitura dos codigos e descricoes de padrões de cobrança
    lErro = CF("Cod_Nomes_Le","PadroesCobranca", "Codigo", "Descricao", STRING_PADRAO_COBRANCA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 16524

    'preenche ListBox Padroes com descrição de padrões de cobrança
    For Each objCodigoDescricao In colCodigoDescricao

        Padroes.AddItem objCodigoDescricao.sNome
        Padroes.ItemData(Padroes.NewIndex) = objCodigoDescricao.iCodigo

    Next

    Set colCodigoDescricao = New AdmColCodigoNome

    'leitura dos codigos e descricoes de Tipos Instrução de Cobrança
    lErro = CF("Cod_Nomes_Le","TiposInstrCobranca", "Codigo", "Descricao", STRING_TIPO_INSTR_COBRANCA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 16525

    'preenche as ComboBox Instrucao1 e Instrucao2 com codigo e descricao de tipos instrucao de cobranca
    For Each objCodigoDescricao In colCodigoDescricao

        sTexto = CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Instrucao1.AddItem sTexto
        Instrucao1.ItemData(Instrucao1.NewIndex) = objCodigoDescricao.iCodigo
        Instrucao2.AddItem sTexto
        Instrucao2.ItemData(Instrucao2.NewIndex) = objCodigoDescricao.iCodigo

    Next
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 16524, 16525 'Tratados nas Rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164111)

    End Select
    
    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objPadraoCobranca As ClassPadraoCobranca) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se existir um Padrao Cobranca passado como parametro, exibir seus dados
    If Not (objPadraoCobranca Is Nothing) Then

        lErro = CF("PadraoCobranca_Le",objPadraoCobranca)
        If lErro <> SUCESSO And lErro <> 19298 Then Error 16526

        If lErro = SUCESSO Then

            'exibe dados do Padrao Cobranca na tela
            lErro = Traz_PadraoCobranca_Tela(objPadraoCobranca)
            If lErro <> SUCESSO Then Error 16527

        Else

            'exibe apenas o codigo
            Codigo.Text = CStr(objPadraoCobranca.iCodigo)

        End If


    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 16526, 16527 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164112)

    End Select

    Exit Function

End Function

Function Traz_PadraoCobranca_Tela(objPadraoCobranca As ClassPadraoCobranca) As Long
'Exibe os dados do Padrao Cobranca especificada em objPadraoCobranca

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_PadraoCobranca_Tela

    Codigo.Text = CStr(objPadraoCobranca.iCodigo)
    Descricao.Text = objPadraoCobranca.sDescricao
    
    If objPadraoCobranca.iDiasDeProtesto1 = 0 Then
        DiasDeProtesto1.Text = ""
    Else
        DiasDeProtesto1.Text = CStr(objPadraoCobranca.iDiasDeProtesto1)
    End If
    
    If objPadraoCobranca.iDiasDeProtesto2 = 0 Then
        DiasDeProtesto2.Text = ""
    Else
        DiasDeProtesto2.Text = CStr(objPadraoCobranca.iDiasDeProtesto2)
    End If
    
    Inativo.Value = objPadraoCobranca.iInativo
    JurosDiarios.Caption = Format(objPadraoCobranca.dJuros, "##0.00##%")
    JurosMensais.Text = Round((objPadraoCobranca.dJuros * 3000), 2)
    
    'Seleciona Tipo Instrucao de Cobranca na ComboBox Instrucao1
    If objPadraoCobranca.iInstrucao1 = 0 Then
        Instrucao1.Text = ""
    Else
        Instrucao1.Text = CStr(objPadraoCobranca.iInstrucao1)
        lErro = Combo_Item_Seleciona(Instrucao1)
        If lErro <> SUCESSO And lErro <> 12250 Then Error 16529
    End If

    'Seleciona Tipo Instrucao de Cobranca na ComboBox Instrucao2
    If objPadraoCobranca.iInstrucao2 = 0 Then
        Instrucao2.Text = ""
    Else
        Instrucao2.Text = CStr(objPadraoCobranca.iInstrucao2)
        lErro = Combo_Item_Seleciona(Instrucao2)
        If lErro <> SUCESSO And lErro <> 12250 Then Error 16530
    End If
    
    iAlterado = 0

    Traz_PadraoCobranca_Tela = SUCESSO

    Exit Function

Erro_Traz_PadraoCobranca_Tela:

    Traz_PadraoCobranca_Tela = Err

    Select Case Err

        Case 16529, 16530 'Tratados nas Rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164113)

    End Select

    Exit Function

End Function

Sub Limpa_Tela_PadraoCobranca()
'limpa todos os campos de input da tela PadroesCobranca

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_PadraoCobranca

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call Limpa_Tela(Me)

    Codigo.Text = ""
    
    JurosDiarios.Caption = ""
    
    'Desmarca instrucoes
    Instrucao1.ListIndex = -1
    Instrucao2.ListIndex = -1

    'Desmarca checkbox
    Inativo.Value = 0

    iAlterado = 0

    Exit Sub
    
Erro_Limpa_Tela_PadraoCobranca:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164114)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub Inativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Instrucao1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Instrucao1_Click()

    iAlterado = REGISTRO_ALTERADO
    
    Call Trata_Troca_Instrucao(Instrucao1, DiasDeProtesto1, LabelDiasProtesto1)

End Sub

Private Sub Instrucao1_Validate(bCancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoInstrCobr As New ClassTipoInstrCobr

On Error GoTo Erro_Instrucao1_Validate

    'Verifica se foi preenchida a ComboBox Instrucao1
    If Len(Trim(Instrucao1.Text)) <> 0 Then
        
        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Instrucao1, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 16545

        'Nao existe o ítem com o CÓDIGO na List da ComboBox
        If lErro = 6730 Then Error 16551

        'Não existe o ítem com a STRING na List da ComboBox
        If lErro = 6731 Then Error 16552
            
        If lErro = SUCESSO Then
            
            'Coloca o codigo no objeto para fazer a consulta no BD
            objTipoInstrCobr.iCodigo = Instrucao1.ItemData(Instrucao1.ListIndex)
                
            'Verificar se possui dias para Devolucao/protesto
            lErro = CF("TipoInstrCobranca_Le",objTipoInstrCobr)
            If lErro <> SUCESSO And lErro <> 16549 Then Error 40618
                
            'se não encontrou esse tipo ----> Erro
            If lErro = 16549 Then Error 40619
                
            Call HabilitaDiasInstrucao(objTipoInstrCobr.iRequerDias = 1, DiasDeProtesto1, LabelDiasProtesto1)
            
        End If
    
    Else
    
        Call HabilitaDiasInstrucao(False, DiasDeProtesto1, LabelDiasProtesto1)
    
    End If
        
    Exit Sub

Erro_Instrucao1_Validate:

    bCancel = True
    
    Select Case Err

        Case 16545, 40618

        Case 16551, 16552, 40619 'Não encontrou Tipo Instrucao Cobranca no BD
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_INSTRUCAO_COBRANCA_NAO_ENCONTRADO", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164115)

    End Select

    Exit Sub

End Sub

Private Sub Instrucao2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Instrucao2_Click()

    iAlterado = REGISTRO_ALTERADO
    
    Call Trata_Troca_Instrucao(Instrucao2, DiasDeProtesto2, LabelDiasProtesto2)
    
End Sub

Private Sub Instrucao2_Validate(bCancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoInstrCobr As New ClassTipoInstrCobr

On Error GoTo Erro_Instrucao2_Validate

    'Verifica se foi preenchida a ComboBox instrucao2
    If Len(Trim(Instrucao2.Text)) <> 0 Then
        
        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Instrucao2, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 16554

        'Nao existe o ítem com o CÓDIGO na List da ComboBox
        If lErro = 6730 Then Error 16555

        'Não existe o ítem com a STRING na List da ComboBox
        If lErro = 6731 Then Error 16556
            
        If lErro = SUCESSO Then
            
            'Coloca o codigo no objeto para fazer a consulta no BD
            objTipoInstrCobr.iCodigo = Instrucao2.ItemData(Instrucao2.ListIndex)
                
            'Verificar se possui dias para Devolucao/protesto
            lErro = CF("TipoInstrCobranca_Le",objTipoInstrCobr)
            If lErro <> SUCESSO And lErro <> 16549 Then Error 40620
                
            'se não encontrou esse tipo ----> Erro
            If lErro = 16549 Then Error 40621
                
            Call HabilitaDiasInstrucao(objTipoInstrCobr.iRequerDias = 1, DiasDeProtesto2, LabelDiasProtesto2)
            
        End If
        
    Else
    
        Call HabilitaDiasInstrucao(False, DiasDeProtesto2, LabelDiasProtesto2)
    
    End If
    
    Exit Sub

Erro_Instrucao2_Validate:

    bCancel = True
    
    Select Case Err

        Case 16554, 40620

        Case 16555, 16556, 40621  'Não encontrou Tipo Instrucao Cobranca no BD
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_INSTRUCAO_COBRANCA_NAO_ENCONTRADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164116)

    End Select

    Exit Sub

End Sub

Private Sub JurosMensais_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objPadraoCobranca As New ClassPadraoCobranca
        
    'Informa tabela associada à Tela
    sTabela = "PadroesCobranca"
        
    'Le os dados da tela PadroesCobranca
    Call Move_Tela_Memoria(objPadraoCobranca)
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
      
    colCampoValor.Add "Codigo", objPadraoCobranca.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objPadraoCobranca.sDescricao, STRING_PADRAO_COBRANCA_DESCRICAO, "Descricao"
    colCampoValor.Add "Inativo", objPadraoCobranca.iInativo, 0, "Inativo"
    colCampoValor.Add "Juros", objPadraoCobranca.dJuros, 0, "Juros"
    colCampoValor.Add "Instrucao1", objPadraoCobranca.iInstrucao1, 0, "Instrucao1"
    colCampoValor.Add "DiasDeProtesto1", objPadraoCobranca.iDiasDeProtesto1, 0, "DiasDeProtesto1"
    colCampoValor.Add "Instrucao2", objPadraoCobranca.iInstrucao2, 0, "Instrucao2"
    colCampoValor.Add "DiasDeProtesto2", objPadraoCobranca.iDiasDeProtesto2, 0, "DiasDeProtesto2"
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPadraoCobranca As New ClassPadraoCobranca

On Error GoTo Erro_Tela_Preenche

    objPadraoCobranca.iCodigo = colCampoValor.Item("Codigo").vValor

    If objPadraoCobranca.iCodigo <> 0 Then

        objPadraoCobranca.sDescricao = colCampoValor.Item("Descricao").vValor
        objPadraoCobranca.iInativo = colCampoValor.Item("Inativo").vValor
        objPadraoCobranca.dJuros = colCampoValor.Item("Juros").vValor
        objPadraoCobranca.iInstrucao1 = colCampoValor.Item("Instrucao1").vValor
        objPadraoCobranca.iDiasDeProtesto1 = colCampoValor.Item("DiasDeProtesto1").vValor
        objPadraoCobranca.iInstrucao2 = colCampoValor.Item("Instrucao2").vValor
        objPadraoCobranca.iDiasDeProtesto2 = colCampoValor.Item("DiasDeProtesto2").vValor
            
        'Joga o Padrão de Cobrança na Tela
        lErro = Traz_PadraoCobranca_Tela(objPadraoCobranca)
        If lErro <> SUCESSO Then Error 16609
        
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 16609 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164117)

    End Select

    Exit Sub
        
End Sub

Private Sub JurosMensais_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dJuros As Double

On Error GoTo Erro_JurosMensais_Validate
    
    ' Verifica se o Juros foi Preencido
    If Len(Trim(JurosMensais.Text)) <> 0 Then
    
        'Critica se 'e Percentagem
        lErro = Porcentagem_Critica(JurosMensais.Text)
        If lErro <> SUCESSO Then Error 16561
        
        JurosMensais.Text = Format(JurosMensais.Text, "Fixed")
        
        'Calcula o juros Diarios
        dJuros = CDbl(JurosMensais.Text)
        dJuros = dJuros / 30
        
        'Formata com 4 casas está correto
        JurosDiarios.Caption = Format(dJuros / 100, "##0.00##%")
        
    Else
        
        JurosDiarios.Caption = ""
    
    End If
        
    Exit Sub
    
Erro_JurosMensais_Validate:

    Cancel = True

    
    Select Case Err

        Case 16561
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164118)
    
    End Select

    Exit Sub

End Sub

Private Sub Padroes_DblClick()

Dim lErro As Long
Dim iIndice As Integer
Dim objPadraoCobranca As New ClassPadraoCobranca

On Error GoTo Erro_Padroes_DblClick

    objPadraoCobranca.iCodigo = Padroes.ItemData(Padroes.ListIndex)
    
    'Lê o Padrão de Cobrança selecionado
    lErro = CF("PadraoCobranca_Le",objPadraoCobranca)
    If lErro <> SUCESSO And lErro <> 19298 Then Error 16542
    
    'PadraoCobranca não está cadastrado --> ERRO
    If lErro = 19298 Then Error 16544
    
    'Preenche a Tela com os dados do Padrão de Cobrança selecionado
    lErro = Traz_PadraoCobranca_Tela(objPadraoCobranca)
    If lErro <> SUCESSO Then Error 16543

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Padroes_DblClick:

    Select Case Err

        Case 16542, 16543 'Tratado na Rotina chamada

        Case 16544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_NAO_CADASTRADO", Err, Padroes.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164119)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objPadraoCobranca As New ClassPadraoCobranca
Dim objTipoInstrCobr As New ClassTipoInstrCobr
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 16564

    'verifica preenchimento da descricao
    If Len(Trim(Descricao.Text)) = 0 Then Error 16565

    'verifica preenchimento da instrucao1
    If Len(Trim(Instrucao1.Text)) > 0 Then
    
        objTipoInstrCobr.iCodigo = Instrucao1.ItemData(Instrucao1.ListIndex)
        
        lErro = CF("TipoInstrCobranca_Le",objTipoInstrCobr)
        If lErro <> SUCESSO And lErro <> 16549 Then Error 16566
        
        'Se não encontrou --> ERRO
        If lErro = 16549 Then Error 16567
        
        If objTipoInstrCobr.iRequerDias = 1 Then
        
            'verifica preenchimento de DiasDeProtesto1
            If Len(Trim(DiasDeProtesto1.Text)) = 0 Then Error 16568
            
        Else 'objTipoIntrCobr.iRequerDias=0
        
            'Verifica se Dias de Dev/Protesto está preenchido
            If Len(Trim(DiasDeProtesto1.Text)) > 0 Then
            
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DIAS_PROTESTO", "Primária")

                If vbMsgRes = vbYes Then
                
                    DiasDeProtesto1.Text = "0"
                    
                Else 'vbMsgRes=vbNo
                
                    Error 16610
                    
                End If
            
            End If
        
        End If
        
    Else
        Error 40617
    
    End If
        
    'verifica preenchimento da instrucao2
    If Len(Trim(Instrucao2.Text)) > 0 Then
    
        objTipoInstrCobr.iCodigo = Instrucao2.ItemData(Instrucao2.ListIndex)
        
        'Lê o TipoInstrCobranca no BD
        lErro = CF("TipoInstrCobranca_Le",objTipoInstrCobr)
        If lErro <> SUCESSO And lErro <> 16549 Then Error 16569
        
        'Se não encontrou ---> ERRO
        If lErro = 16549 Then Error 16570
        
        If objTipoInstrCobr.iRequerDias = 1 Then
        
            'Verifica preenchimento de DiasDeProtesto2
            If Len(Trim(DiasDeProtesto2.Text)) = 0 Then Error 16571
            
        Else 'objTipoIntrCobr.iRequerDias=0
        
            'Verifica se Dias de Dev/Protesto está preenchido
            If Len(Trim(DiasDeProtesto2.Text)) > 0 Then
            
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DIAS_PROTESTO", "Secundária")

                If vbMsgRes = vbYes Then
                
                    DiasDeProtesto2.Text = "0"
                    
                Else 'vbMsgRes=vbNo
                
                    Error 16611
                    
                End If
            
            End If
            
        End If
        
    End If
    
    'preenche objPadraoCobranca
    Call Move_Tela_Memoria(objPadraoCobranca)
    
    lErro = Trata_Alteracao(objPadraoCobranca, objPadraoCobranca.iCodigo)
    If lErro <> SUCESSO Then Error 32287
        
    'chama função de gravação
    lErro = CF("PadraoCobranca_Grava",objPadraoCobranca)
    If lErro <> SUCESSO Then Error 32288

    'Remove o PadraoCobranca da ListBox Padroes
    Call Padroes_Exclui(objPadraoCobranca.iCodigo)
    
    'Insere o PadraoCobranca na ListBox Padroes
    Call Padroes_Adiciona(objPadraoCobranca)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", Err)

        Case 16566, 16569, 16572, 16610, 16611 'Tratado na Rotina chamada
        
        Case 16567, 16570
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_INSTRUCAO_COBRANCA_NAO_CADASTRADA", Err, objTipoInstrCobr.iCodigo)
        
        Case 16568
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DIAS_DE_PROTESTO1_NAO_PREENCHIDO", Err)
        
        Case 16571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DIAS_DE_PROTESTO2_NAO_PREENCHIDO", Err)
        
        Case 32287
        
        Case 40617
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_PRIMARIA_NAO_PREENCHIDA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164120)

    End Select
    
    Exit Function
        
End Function

Private Sub Padroes_Adiciona(objPadraoCobranca As ClassPadraoCobranca)

    Padroes.AddItem objPadraoCobranca.sDescricao
    Padroes.ItemData(Padroes.NewIndex) = objPadraoCobranca.iCodigo

End Sub

Private Sub Padroes_Exclui(iCodigo As Integer)

Dim iIndice As Integer

    For iIndice = 0 To Padroes.ListCount - 1

        If Padroes.ItemData(iIndice) = iCodigo Then

            Padroes.RemoveItem iIndice
            Exit For

        End If

    Next

End Sub

Private Sub Move_Tela_Memoria(objPadraoCobranca As ClassPadraoCobranca)
'Le os dados que estao na tela PadroesCobranca e coloca em objPadraoCobranca

Dim sJuros As String

    If Len(Trim(Codigo.Text)) > 0 Then objPadraoCobranca.iCodigo = CInt(Codigo.Text)
    objPadraoCobranca.sDescricao = Descricao.Text
    objPadraoCobranca.iInativo = Inativo.Value
    
    If Len(Trim(DiasDeProtesto1.Text)) = 0 Then
        objPadraoCobranca.iDiasDeProtesto1 = 0
    Else
        objPadraoCobranca.iDiasDeProtesto1 = CInt(DiasDeProtesto1.Text)
    End If

    If Len(Trim(DiasDeProtesto2.Text)) = 0 Then
        objPadraoCobranca.iDiasDeProtesto2 = 0
    Else
        objPadraoCobranca.iDiasDeProtesto2 = CInt(DiasDeProtesto2.Text)
    End If
    
    If Len(Trim(JurosDiarios.Caption)) = 0 Then
        objPadraoCobranca.dJuros = 0
    Else
        'Retira o "%" do juros
        sJuros = Mid(JurosDiarios.Caption, 1, Len(Trim(JurosDiarios.Caption)) - 1)
        objPadraoCobranca.dJuros = CDbl(sJuros) / 100
    End If
    
    If Len(Trim(Instrucao1.Text)) = 0 Then
        objPadraoCobranca.iInstrucao1 = 0
    Else
        objPadraoCobranca.iInstrucao1 = Codigo_Extrai(Instrucao1.Text)
    End If
    
    If Len(Trim(Instrucao2.Text)) = 0 Then
        objPadraoCobranca.iInstrucao2 = 0
    Else
        objPadraoCobranca.iInstrucao2 = Codigo_Extrai(Instrucao2.Text)
    End If
    
End Sub

Private Sub ListaPadraoCobranca_Exclui(iCodigo As Integer)
'Exclui ítem da ListBox Padroes Cobranca

Dim iIndice As Integer

    'Percorre todos os itens da ListBox
    For iIndice = 0 To Padroes.ListCount - 1

        'Se o ItemData do ítem for igual ao Código passado em iCodigo
        If Padroes.ItemData(iIndice) = iCodigo Then

            'Remove o ítem
            Padroes.RemoveItem (iIndice)
            Exit For

        End If

    Next

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PADROES_COBRANCAS
    Set Form_Load_Ocx = Me
    Caption = "Padrões de Cobrança"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PadroesCobranca"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
End Sub

Sub Trata_Troca_Instrucao(ComboInstrucao As Object, DiasDeProtesto As Object, LabelDiasProtesto As Object)

Dim lErro As Long
Dim objTipoInstrCobr As New ClassTipoInstrCobr

On Error GoTo Erro_Trata_Troca_Instrucao

    If ComboInstrucao.ListIndex <> -1 Then
    
        objTipoInstrCobr.iCodigo = ComboInstrucao.ItemData(ComboInstrucao.ListIndex)
        
        'Verificar se possui dias para Devolucao/protesto
        lErro = CF("TipoInstrCobranca_Le",objTipoInstrCobr)
        If lErro <> SUCESSO And lErro <> 16549 Then Error 59258
            
        'se não encontrou esse tipo ----> Erro
        If lErro = 16549 Then Error 59259
            
        Call HabilitaDiasInstrucao(objTipoInstrCobr.iRequerDias = 1, DiasDeProtesto, LabelDiasProtesto)

    Else
    
        Call HabilitaDiasInstrucao(False, DiasDeProtesto, LabelDiasProtesto)

    End If
    
    Exit Sub
     
Erro_Trata_Troca_Instrucao:

    Select Case Err
          
        Case 59258
        
        Case 59259 'Não encontrou Tipo Instrucao Cobranca no BD
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_INSTRUCAO_COBRANCA_NAO_ENCONTRADO", Err)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164121)
     
    End Select
     
    Exit Sub

End Sub

Sub HabilitaDiasInstrucao(bHabilita As Boolean, DiasDeProtesto As Object, LabelDiasProtesto As Object)

    If bHabilita = False Then
        'se não zera e desabilita
        DiasDeProtesto.Text = ""
        DiasDeProtesto.Enabled = False
        LabelDiasProtesto.Enabled = False
        
    Else
        LabelDiasProtesto.Enabled = True
        DiasDeProtesto.Enabled = True
    End If

End Sub

Private Sub JurosDiarios_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(JurosDiarios, Source, X, Y)
End Sub

Private Sub JurosDiarios_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(JurosDiarios, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub LabelDiasProtesto1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDiasProtesto1, Source, X, Y)
End Sub

Private Sub LabelDiasProtesto1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDiasProtesto1, Button, Shift, X, Y)
End Sub

Private Sub LabelDiasProtesto2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDiasProtesto2, Source, X, Y)
End Sub

Private Sub LabelDiasProtesto2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDiasProtesto2, Button, Shift, X, Y)
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

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

