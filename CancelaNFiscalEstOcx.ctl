VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CancelaNFiscalEst 
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   KeyPreview      =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   6720
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4845
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   165
      Width           =   1680
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "CancelaNFiscalEstOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "CancelaNFiscalEstOcx.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CancelaNFiscalEstOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informações Adicionais"
      Height          =   735
      Left            =   135
      TabIndex        =   16
      Top             =   2415
      Width           =   6420
      Begin MSMask.MaskEdBox MotivoCancel 
         Height          =   300
         Left            =   990
         TabIndex        =   2
         Top             =   285
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo:"
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
         Left            =   225
         TabIndex        =   17
         Top             =   345
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   1605
      Left            =   135
      TabIndex        =   3
      Top             =   750
      Width           =   6420
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   300
         Width           =   765
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   300
         Left            =   3240
         TabIndex        =   1
         Top             =   285
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label LblDataEmissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3270
         TabIndex        =   15
         Top             =   1170
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
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
         Left            =   2475
         TabIndex        =   14
         Top             =   1230
         Width           =   765
      End
      Begin VB.Label LblFilial 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   960
         TabIndex        =   13
         Top             =   1185
         Width           =   1335
      End
      Begin VB.Label LabelFilial 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
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
         Left            =   450
         TabIndex        =   12
         Top             =   1230
         Width           =   465
      End
      Begin VB.Label Fornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Emitente:"
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
         Left            =   105
         TabIndex        =   11
         Top             =   795
         Width           =   810
      End
      Begin VB.Label LblSerie 
         AutoSize        =   -1  'True
         Caption         =   "Série:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   360
         Width           =   510
      End
      Begin VB.Label LblNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Left            =   2490
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Left            =   4395
         TabIndex        =   8
         Top             =   1230
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Left            =   4710
         TabIndex        =   7
         Top             =   330
         Width           =   450
      End
      Begin VB.Label LblEmitente 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   765
         Width           =   5280
      End
      Begin VB.Label LblValor 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4965
         TabIndex        =   5
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label LblTipoNF 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5220
         TabIndex        =   4
         Top             =   315
         Width           =   1020
      End
   End
End
Attribute VB_Name = "CancelaNFiscalEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Dim WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFEntrada As New ClassNFiscal

On Error GoTo Erro_BotaoGravar_Click

    'verifica se todos os campos estao preenchidos ,se nao estiverem => erro
    If Len(Trim(Serie.Text)) = 0 Then Error 34624
    If Len(Trim(Numero.ClipText)) = 0 Then Error 34625

    'Move os dados da NF de entrada para objNFEntrada
    lErro = Move_Dados_NFiscal_Memoria(objNFEntrada)
    If lErro <> SUCESSO Then Error 34659

    'Lê a nota fiscal de entrada
    lErro = CF("NFiscalInternaEntrada_Le_Numero",objNFEntrada)
    If lErro <> SUCESSO And lErro <> 62144 Then Error 62148
    If lErro <> SUCESSO Then Error 62149
    'Verifica se a nota já está cancelada
    If objNFEntrada.iStatus = STATUS_CANCELADO Then Error 62142
    If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then Error 62188

    'pede confirmacao
     vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_NFISCALENTRADA", Numero.Text)

        If vbMsg = vbYes Then
            'Lê os itens da nota fiscal
            lErro = CF("NFiscalItens_Le",objNFEntrada)
            If lErro <> SUCESSO Then Error 62150

            objNFEntrada.sMotivoCancel = MotivoCancel.Text

            'chama NotaFiscal_Cancelar()
            lErro = CF("NotaFiscalEntrada_Cancelar",objNFEntrada)
            If lErro <> SUCESSO Then Error 34660

            Call Limpa_Tela_NFEntrada

            iAlterado = 0

        End If

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 34624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)

        Case 34625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", Err)

        Case 34659, 34660, 62148, 62150
        
        Case 62142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", Err, Serie.Text, Numero.Text)
        
        Case 62148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, Numero.Text)
        
        Case 62188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144166)
        
    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoLimpar_Click

    If iAlterado = REGISTRO_ALTERADO Then

        'Testa se deseja salvar as alterações
        vbMsgRes = Rotina_Aviso(vbYesNoCancel, "AVISO_DESEJA_SALVAR_ALTERACOES")

        If vbMsgRes = vbYes Then

            Call BotaoGravar_Click

        ElseIf vbMsgRes = vbNo Then

            Call Limpa_Tela_NFEntrada

            iAlterado = 0

        Else
            Error 34661
        End If

    End If

Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 34661

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144167)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objNFEntrada As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFEntrada Is Nothing) Then

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscalInternaEntrada_Le_Numero",objNFEntrada)
        If lErro <> SUCESSO And lErro <> 62144 Then Error 34617
        If lErro = 62144 Then Error 34618
        
        If objNFEntrada.iStatus = STATUS_CANCELADO Then Error 62143
        If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then Error 62190

        'Traz a nota para a tela
        lErro = Traz_NFEntrada_Tela(objNFEntrada)
        If lErro <> SUCESSO Then Error 34619

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 34617, 34619

        Case 34618
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFEntrada.lNumNotaFiscal)
            Call Limpa_Tela_NFEntrada
            iAlterado = 0

        Case 62143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", Err, Serie.Text, Numero.Text)

        Case 62190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144168)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub Limpa_Tela_NFEntrada()
'Limpa a Tela NFiscalEntrada

    Serie.Text = ""
    Numero.PromptInclude = False
    Numero.Text = ""
    Numero.PromptInclude = True

    Call Limpa_Tela_NFEntrada1

End Sub

Public Function Traz_EmitFilial_Tela(iEmitente As Integer, objNFEntrada As ClassNFiscal) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Traz_EmitFilial_Tela

    'Se Emitente for empresa
    If iEmitente = EMITENTE_EMPRESA Then
        'EMPRESA
        objFilialEmpresa.iCodFilial = objNFEntrada.iFilialEmpresa

        lErro = CF("FilialEmpresa_Le",objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 58073

        If lErro = 27378 Then Error 58074
        
        If giTipoVersao = VERSAO_LIGHT Then
            LblFilial.Visible = False
            LabelFilial.Visible = False
        End If

        LblEmitente.Caption = gsNomeEmpresa
        LblFilial.Caption = giFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
        
    'Se Emitente for Cliente
    ElseIf iEmitente = EMITENTE_CLIENTE Then

        objCliente.lCodigo = objNFEntrada.lCliente

        'Procura se o CLiente existe
        lErro = CF("Cliente_Le",objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then Error 34650
        If lErro = 12293 Then Error 34651

        objFilialCliente.lCodCliente = objNFEntrada.lCliente
        objFilialCliente.iCodFilial = objNFEntrada.iFilialCli

        'Procura se a Filial existe
        lErro = CF("FilialCliente_Le",objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then Error 34652
        If lErro = 12567 Then Error 34653

        If giTipoVersao = VERSAO_LIGHT Then
            LblFilial.Visible = False
            LabelFilial.Visible = False
        End If

        LblEmitente.Caption = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
        LblFilial.Caption = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome

    'Se Emitente for Fornecedor
    ElseIf iEmitente = EMITENTE_FORNECEDOR Then

        objFornecedor.lCodigo = objNFEntrada.lFornecedor

        'procura pelo Fornecedor
        lErro = CF("Fornecedor_Le",objFornecedor)
        If lErro <> SUCESSO And lErro <> 12732 Then Error 34654
        If lErro = 12732 Then Error 34655

        objFilialFornecedor.lCodFornecedor = objNFEntrada.lFornecedor
        objFilialFornecedor.iCodFilial = objNFEntrada.iFilialForn

        'procura pela filial do fornecedor
        lErro = CF("FilialFornecedor_Le",objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 12929 Then Error 34656
        If lErro = 12929 Then Error 34657

        If giTipoVersao = VERSAO_LIGHT Then
            LblFilial.Visible = True
            LabelFilial.Visible = True
        End If

        LblEmitente.Caption = objFornecedor.lCodigo & SEPARADOR & objFornecedor.sNomeReduzido
        LblFilial.Caption = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

    End If

    Traz_EmitFilial_Tela = SUCESSO

    Exit Function

Erro_Traz_EmitFilial_Tela:

    Traz_EmitFilial_Tela = Err

    Select Case Err

        Case 34650, 34652, 34654, 34656, 58073
            Numero.SetFocus

        Case 34651
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objNFEntrada.lCliente)
            Numero.SetFocus

        Case 34653
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", Err, objNFEntrada.lCliente)
            Numero.SetFocus

        Case 34655
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objNFEntrada.lFornecedor)
            Numero.SetFocus

        Case 34657
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_SEM_FILIAL", Err, objNFEntrada.lFornecedor)
            Numero.SetFocus

        Case 58074
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
            Numero.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144169)

    End Select

    Exit Function

End Function

Public Function Traz_NFEntrada_Tela(objNFEntrada As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal passada em objNFEntrada

Dim lErro As Long
Dim iIndice As Integer
Dim sTipoNF As String
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iEmitente As Integer

On Error GoTo Erro_Traz_NFEntrada_Tela

    'Limpa a tela NFicalEntrada
    Call Limpa_Tela_NFEntrada

    'Preenche o número da NF
    If objNFEntrada.lNumNotaFiscal > 0 Then
        Numero.PromptInclude = False
        Numero.Text = CStr(objNFEntrada.lNumNotaFiscal)
        Numero.PromptInclude = True
    End If

    'preenche a serie da NF
    Serie.Text = objNFEntrada.sSerie

    objTipoDocInfo.iCodigo = objNFEntrada.iTipoNFiscal

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo",objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 34648
    If lErro = 31415 Then Error 34649

    'preenche a Sigla da NF
    LblTipoNF.Caption = objTipoDocInfo.sSigla

    iEmitente = objTipoDocInfo.iEmitente

    'Traz a tela os dados de Emitente e Filial
    lErro = Traz_EmitFilial_Tela(iEmitente, objNFEntrada)
    If lErro <> SUCESSO Then Error 34658

    'Se a data não for nula coloca na Tela
    If objNFEntrada.dtDataEmissao <> DATA_NULA Then
        LblDataEmissao.Caption = Format(objNFEntrada.dtDataEmissao, "dd/mm/yy")
    Else
        LblDataEmissao.Caption = Format("", "dd/mm/yy")
    End If

    'Preenche o valor total da NF
    If objNFEntrada.dValorTotal > 0 Then
        LblValor.Caption = Format(objNFEntrada.dValorTotal, "Fixed")
    Else
        LblValor.Caption = Format(0, "Fixed")
    End If

    Traz_NFEntrada_Tela = SUCESSO

    Exit Function

Erro_Traz_NFEntrada_Tela:

    Traz_NFEntrada_Tela = Err

    Select Case Err

        Case 34648, 34658

        Case 34649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144170)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objSerie As ClassSerie
Dim colSerie As New colSerie

On Error GoTo Erro_Form_Load

    Set objEventoSerie = New AdmEvento
    Set objEventoNumero = New AdmEvento

    'nao pode entrar como EMPRESA_TODA
    If giFilialEmpresa = EMPRESA_TODA Then Error 34615

    'obtem a colecao de series
    lErro = CF("Series_Le",colSerie)
    If lErro <> SUCESSO Then Error 34616

    'preenche as duas combos de serie
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case Err

        Case 34615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_INVALIDA", Err)

        Case 34616

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144171)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Move_Dados_NFiscal_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os dados da NotaFiscalOriginal para a memória

Dim lErro As Long

On Error GoTo Erro_Move_Dados_NFiscal_Memoria

    'verifica se a Serie e o Número da NF de entrada estão preenchidos
    If Len(Trim(Numero.ClipText)) > 0 Then objNFiscal.lNumNotaFiscal = CLng(Numero.Text)
    If Len(Trim(Serie.Text)) > 0 Then objNFiscal.sSerie = Serie.Text

    Move_Dados_NFiscal_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_NFiscal_Memoria:

    Move_Dados_NFiscal_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144172)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoSerie = Nothing
    Set objEventoNumero = Nothing

End Sub

Private Sub LblNumero_Click()

Dim lErro As Long
Dim objNFEntrada As New ClassNFiscal
Dim colSelecao As Collection

On Error GoTo Erro_LblNumero_Click

    'Preenche objNFEntrada com o numero
    lErro = Move_Dados_NFiscal_Memoria(objNFEntrada)
    If lErro <> SUCESSO Then Error 34620

    Call Chama_Tela("NFiscalInternaEntradaLista", colSelecao, objNFEntrada, objEventoNumero)

    Exit Sub

Erro_LblNumero_Click:

    Select Case Err

        Case 34620

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144173)

    End Select

    Exit Sub

End Sub

Private Sub LblSerie_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objSerie As New ClassSerie
Dim colSelecao As Collection

On Error GoTo Erro_LblSerie_Click

    'transfere a série da tela p\ o objSerie
    objSerie.sSerie = Serie.Text

    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

    Exit Sub

Erro_LblSerie_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144174)

    End Select

    Exit Sub

End Sub

Private Sub MotivoCancel_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()

    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNFEntrada As New ClassNFiscal

On Error GoTo Erro_Numero_Validate
    
    Cancel = False
    
    'Se a série não estiver preenchida, sai.
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
    'Se o número estiver preenchido
    If Len(Trim(Numero.ClipText)) > 0 Then
        'Recolhe a série e o número
        objNFEntrada.lNumNotaFiscal = Numero.Text
        objNFEntrada.sSerie = Serie.Text

        'procura pela nota no BD
        lErro = CF("NFiscalInternaEntrada_Le_Numero",objNFEntrada)
        If lErro <> SUCESSO And lErro <> 62144 Then Error 34637
        If lErro = 62144 Then Error 34638 'Não encontrou
        'verifica se a nota já está cancelada
        If objNFEntrada.iStatus = STATUS_CANCELADO Then Error 62144
        If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then Error 62191

        'Traz a NotaFiscal de Entrada para a a tela
        lErro = Traz_NFEntrada_Tela(objNFEntrada)
        If lErro <> SUCESSO Then Error 34639

    End If

    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case Err

        Case 34637, 34639

        Case 34638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFEntrada.lNumNotaFiscal)
            Call Limpa_Tela_NFEntrada1
            iAlterado = 0
        
        Case 62144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", Err, Serie.Text, Numero.Text)

        Case 62191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144175)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFEntrada As ClassNFiscal

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objNFEntrada = obj1

    lErro = CF("NFiscalInternaEntrada_Le_Numero",objNFEntrada)
    If lErro <> SUCESSO And lErro <> 62144 Then Error 34632
    If lErro = 62144 Then Error 34633
    
    If objNFEntrada.iStatus = STATUS_CANCELADO Then Error 62145
    If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then Error 62192

    'Traz a NotaFiscal de Entrada para a a tela
    lErro = Traz_NFEntrada_Tela(objNFEntrada)
    If lErro <> SUCESSO Then Error 34623

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case Err

        Case 34623, 34632

        Case 34633
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFEntrada.lNumNotaFiscal)
            Call Limpa_Tela_NFEntrada
            iAlterado = 0

        Case 62145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", Err, Serie.Text, Numero.Text)
        
        Case 62192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144176)

    End Select

    Exit Sub

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objSerie As ClassSerie
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_objEventoSerie_evSelecao

    Set objSerie = obj1

    Serie.Text = objSerie.sSerie
    Call Serie_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoSerie_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144177)

    End Select

    Exit Sub

End Sub

Private Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNFEntrada As New ClassNFiscal
Dim objSerie As New ClassSerie

On Error GoTo Erro_Serie_Validate

    Cancel = False
    
    'Verifica se a série está preenchida
    If Len(Trim(Serie.Text)) > 0 Then
       'Verifica se o número está preenchido
       If Len(Trim(Numero.ClipText)) > 0 Then

            objNFEntrada.lNumNotaFiscal = Numero.Text
            objNFEntrada.sSerie = Serie.Text

            'procura pela nota no BD
            lErro = CF("NFiscalInternaEntrada_Le_Numero",objNFEntrada)
            If lErro <> SUCESSO And lErro <> 62144 Then Error 34634
            If lErro = 62144 Then Error 34635

            If objNFEntrada.iStatus = STATUS_CANCELADO Then Error 62146
            If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then Error 62193

            'Traz a NotaFiscal de Entrada para a a tela
            lErro = Traz_NFEntrada_Tela(objNFEntrada)
            If lErro <> SUCESSO Then Error 34636

        Else

            objSerie.sSerie = Serie.Text

            lErro = CF("Serie_Le",objSerie)
            If lErro <> SUCESSO And lErro <> 22202 Then Error 34646
            If lErro = 22202 Then Error 34647

        End If

    End If

    Exit Sub

Erro_Serie_Validate:

    Cancel = True
    
    Select Case Err

        Case 34634, 34636, 34646

        Case 34635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFEntrada.lNumNotaFiscal)
            Call Limpa_Tela_NFEntrada1
            iAlterado = 0

        Case 34646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, objSerie.sSerie)

        Case 62146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", Err, Serie.Text, Numero.Text)

        Case 62193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", Err)

        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144178)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_CANCELA_NFISCALEST
    Set Form_Load_Ocx = Me
    Caption = "Cancelamento de Nota Fiscal de Entrada"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CancelaNFiscal"

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

        If Me.ActiveControl Is Serie Then
            Call LblSerie_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call LblNumero_Click
        End If

    End If

End Sub

Public Sub Limpa_Tela_NFEntrada1()
'Limpa a Tela NFiscalEntrada

    LblEmitente.Caption = ""
    LblFilial.Caption = ""
    LblValor.Caption = ""
    LblTipoNF.Caption = ""
    LblDataEmissao.Caption = ""
    MotivoCancel.Text = ""
    
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub LblDataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblDataEmissao, Source, X, Y)
End Sub

Private Sub LblDataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblDataEmissao, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LblFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblFilial, Source, X, Y)
End Sub

Private Sub LblFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblFilial, Button, Shift, X, Y)
End Sub

Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub

Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub

Private Sub Fornecedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Fornecedor, Source, X, Y)
End Sub

Private Sub Fornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Fornecedor, Button, Shift, X, Y)
End Sub

Private Sub LblSerie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblSerie, Source, X, Y)
End Sub

Private Sub LblSerie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblSerie, Button, Shift, X, Y)
End Sub

Private Sub LblNumero_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNumero, Source, X, Y)
End Sub

Private Sub LblNumero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNumero, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LblEmitente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblEmitente, Source, X, Y)
End Sub

Private Sub LblEmitente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblEmitente, Button, Shift, X, Y)
End Sub

Private Sub LblValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblValor, Source, X, Y)
End Sub

Private Sub LblValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblValor, Button, Shift, X, Y)
End Sub

Private Sub LblTipoNF_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblTipoNF, Source, X, Y)
End Sub

Private Sub LblTipoNF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblTipoNF, Button, Shift, X, Y)
End Sub

