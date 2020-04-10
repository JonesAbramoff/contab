VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl DetPagOcx 
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   8160
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   5280
      ScaleHeight     =   495
      ScaleWidth      =   2685
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   2745
      Begin VB.CommandButton BotaoDocOriginal 
         Height          =   390
         Left            =   60
         Picture         =   "DetPagOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Consulta de título a pagar"
         Top             =   60
         Width           =   1020
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   390
         Left            =   2205
         Picture         =   "DetPagOcx.ctx":0F0A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   390
         Left            =   1690
         Picture         =   "DetPagOcx.ctx":1088
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   390
         Left            =   1175
         Picture         =   "DetPagOcx.ctx":15BA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Atributos"
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   2070
      Width           =   7905
      Begin VB.Label Label7 
         Caption         =   "Valor Original:"
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
         Left            =   360
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label ValorOriginal 
         Height          =   180
         Left            =   1620
         TabIndex        =   36
         Top             =   390
         Width           =   1320
      End
      Begin VB.Label Label16 
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
         Height          =   255
         Left            =   2955
         TabIndex        =   19
         Top             =   375
         Width           =   765
      End
      Begin VB.Label LabelTipoDoc 
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4815
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Tipo 
         Height          =   240
         Left            =   5310
         TabIndex        =   21
         Top             =   360
         Width           =   2190
      End
      Begin VB.Label DataEmissao 
         Height          =   270
         Left            =   3795
         TabIndex        =   22
         Top             =   375
         Width           =   870
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Confirmar"
      Height          =   2340
      Left            =   120
      TabIndex        =   17
      Top             =   2985
      Width           =   7905
      Begin VB.ComboBox MotivoDiferenca 
         Height          =   315
         Left            =   5325
         TabIndex        =   6
         Top             =   255
         Width           =   2370
      End
      Begin VB.ComboBox ComboCobrador 
         Height          =   315
         Left            =   1605
         TabIndex        =   8
         Top             =   1830
         Width           =   2205
      End
      Begin VB.ComboBox ComboTipoCobranca 
         Height          =   315
         Left            =   1635
         TabIndex        =   7
         Top             =   810
         Width           =   2175
      End
      Begin VB.ComboBox ComboPortador 
         Height          =   315
         Left            =   5295
         TabIndex        =   9
         Top             =   1830
         Width           =   2445
      End
      Begin MSComCtl2.UpDown UpDownVencimento 
         Height          =   300
         Left            =   6420
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   825
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataVencimento 
         Height          =   300
         Left            =   5325
         TabIndex        =   10
         Top             =   810
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodigodeBarras 
         Height          =   315
         Left            =   1620
         TabIndex        =   32
         Top             =   1320
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   57
         Mask            =   "#####.#####.#####.######.#####.######.#.#################"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   300
         Left            =   1620
         TabIndex        =   5
         Top             =   240
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label LabelMotivoDiferenca 
         AutoSize        =   -1  'True
         Caption         =   "Motivo Diferença:"
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
         Left            =   3765
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   35
         Top             =   300
         Width           =   1530
      End
      Begin VB.Label LabelValor 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1065
         TabIndex        =   34
         Top             =   270
         Width           =   510
      End
      Begin VB.Label CodigoBarras 
         AutoSize        =   -1  'True
         Caption         =   "Código de Barras:"
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
         Left            =   45
         TabIndex        =   33
         Top             =   1365
         Width           =   1530
      End
      Begin VB.Label Label4 
         Caption         =   "Vencimento:"
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
         Left            =   4215
         TabIndex        =   23
         Top             =   870
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Banco Cobrador:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   1890
         Width           =   1455
      End
      Begin VB.Label Cobranca 
         Caption         =   "Cobrança:"
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
         Height          =   270
         Left            =   690
         TabIndex        =   25
         Top             =   855
         Width           =   915
      End
      Begin VB.Label Portador 
         Caption         =   "Portador:"
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
         Left            =   4440
         TabIndex        =   26
         Top             =   1875
         Width           =   855
      End
   End
   Begin VB.Frame SSFrame2 
      Caption         =   "Título a Pagar"
      Height          =   1350
      Left            =   120
      TabIndex        =   15
      Top             =   690
      Width           =   7905
      Begin VB.CommandButton TrazerParcela 
         Caption         =   "Trazer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   4770
         TabIndex        =   4
         Top             =   825
         Width           =   1500
      End
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   5370
         TabIndex        =   1
         Top             =   345
         Width           =   2310
      End
      Begin MSMask.MaskEdBox Fornecedor 
         Height          =   300
         Left            =   1590
         TabIndex        =   0
         Top             =   345
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NumTitulo 
         Height          =   300
         Left            =   1590
         TabIndex        =   2
         Top             =   825
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "999999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Parcela 
         Height          =   300
         Left            =   3765
         TabIndex        =   3
         Top             =   840
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "999"
         PromptChar      =   " "
      End
      Begin VB.Label LabelNumParcela 
         AutoSize        =   -1  'True
         Caption         =   "Parcela:"
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
         Left            =   2940
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   885
         Width           =   720
      End
      Begin VB.Label NumeroLabel 
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
         Left            =   780
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   885
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   " Filial:"
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
         Left            =   4770
         TabIndex        =   29
         Top             =   405
         Width           =   525
      End
      Begin VB.Label FornecedorLabel 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor:"
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
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   30
         Top             =   375
         Width           =   1035
      End
   End
End
Attribute VB_Name = "DetPagOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim glNumIntParc As Long '# interno da parcela que está sendo alterada
Dim iAlterado As Integer
Dim iFornecedorAlterado As Integer
Dim iParcelaAlterada As Integer

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoTipoDoc As AdmEvento
Attribute objEventoTipoDoc.VB_VarHelpID = -1
Private WithEvents objEventoParcela As AdmEvento
Attribute objEventoParcela.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoFornecedor = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoTipoDoc = New AdmEvento
    Set objEventoParcela = New AdmEvento
    
    glNumIntParc = 0

    'Carrega os tipos de cobranca
    lErro = Carrega_TipoCobranca()
    If lErro <> SUCESSO Then gError 18951

    'Carrega os bancos
    lErro = Carrega_Bancos()
    If lErro <> SUCESSO Then gError 18953

    'Carrega os portadores
    lErro = Carrega_Portadores()
    If lErro <> SUCESSO Then gError 18955
    
    'Carrega Motivos de diferença
    lErro = Carrega_MotivoDiferenca()
    If lErro <> SUCESSO Then gError 500065

    'para desabilitar controles de acordo com o tipo de cobranca
    Call ComboTipoCobranca_Click
    
    'Zera os Flags de Alteração
    iParcelaAlterada = 0
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 18951, 18953, 18955, 500065 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158927)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_TipoCobranca() As Long
'Carrega a combo de Tipo de Cobrança

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoCobranca

    'Lê o código e a descrição de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 18952

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na List da Combo TipoCobranca
        ComboTipoCobranca.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        ComboTipoCobranca.ItemData(ComboTipoCobranca.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TipoCobranca = SUCESSO

    Exit Function

Erro_Carrega_TipoCobranca:

    Carrega_TipoCobranca = gErr

    Select Case gErr

        Case 18952 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158928)

    End Select

    Exit Function

End Function

Private Function Carrega_Bancos() As Long
'Carrega a combo de Cobrador

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Bancos

    'Leitura dos códigos e descrições dos Bancos BD
    lErro = CF("Cod_Nomes_Le", "Bancos", "CodBanco", "NomeReduzido", STRING_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then gError 18954

   'Preenche ComboBox com código e nome dos Bancos
    For iIndice = 1 To colCodigoNome.Count
        Set objCodigoNome = colCodigoNome(iIndice)
        ComboCobrador.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        ComboCobrador.ItemData(ComboCobrador.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_Bancos = SUCESSO

    Exit Function

Erro_Carrega_Bancos:

    Carrega_Bancos = gErr

    Select Case gErr

        Case 18954 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158929)

    End Select

    Exit Function

End Function

Private Function Carrega_Portadores() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_Portadores

    'Lê o código e a descrição de todos os Portadores Ativos
    lErro = CF("Portadores_Le_CodigosNomesRed", colCodigoDescricao)
    If lErro <> SUCESSO Then gError 18956

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na Combo de Portadores
        ComboPortador.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        ComboPortador.ItemData(ComboPortador.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_Portadores = SUCESSO

    Exit Function

Erro_Carrega_Portadores:

    Carrega_Portadores = gErr

    Select Case gErr

        Case 18956

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158930)

    End Select

    Exit Function

End Function

Private Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim objFornecedor As New ClassFornecedor
Dim objParcelaPagar As New ClassParcelaPagar

On Error GoTo Erro_BotaoDocOriginal_Click

        If (Len(Trim(NumTitulo.Text))) = 0 Then gError 57463

        objTituloPagar.lNumTitulo = StrParaLong(NumTitulo.Text)

        If Len(Trim(Fornecedor.Text)) = 0 Then gError 57464

        objFornecedor.sNomeReduzido = Fornecedor.Text

        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 57465

        If lErro = 6681 Then gError 57466

        objTituloPagar.lFornecedor = CLng(objFornecedor.lCodigo)
        objTituloPagar.iFilial = CInt(Filial.ItemData(Filial.ListIndex))
        objTituloPagar.dtDataEmissao = StrParaDate(DataEmissao.Caption)
        
        If Len(Trim(Tipo.Caption)) > 0 Then
            objTituloPagar.sSiglaDocumento = Tipo.Caption
        Else
            lErro = CF("TituloPagar_Le_Titulo", objTituloPagar)
            If lErro <> SUCESSO And lErro <> 43451 Then gError 57485
            'Se nao encontrar ==> Erro
            If lErro = 43451 Then gError 57486
        End If
        
        Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)

    Exit Sub

Erro_BotaoDocOriginal_Click:

    Select Case gErr

        Case 57463
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_PREENCHIDO", gErr)

        Case 57464
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 57465, 57485

        Case 57466
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)
        
        Case 57486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158931)

    End Select

    Exit Sub

End Sub

'Campos Referêntes ao Código de barras

Private Sub CodigodeBarras_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodigodeBarras_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigodeBarras, iAlterado)

End Sub

Private Sub CodigodeBarras_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodBanco As Integer
Dim dNumDias As Double
Dim dtDataVenc As Date
Dim sLinDig As String

On Error GoTo Erro_CodigodeBarras_Validate

    lErro = CB_Converte_LinDig(CodigodeBarras.ClipText, sLinDig)
    If lErro <> SUCESSO Then gError 188321
       
    CodigodeBarras.PromptInclude = False
    CodigodeBarras.Text = sLinDig
    CodigodeBarras.PromptInclude = True
        
    'se o cobrador estiver vazio e o codigo de barras nao, entao sugerir cobrador
    '??? melhoria: criar flag de alterado p/campo codigo de barras e sair direto caso o codigo nao tenha sido alterado
    If Len(Trim(CodigodeBarras.ClipText)) <> 0 And Len(Trim(ComboCobrador.Text)) = 0 Then
        
        'Preenche a Combo de Banco Cobrador com os 3 primeiros dígitos do Código de barras
        iCodBanco = StrParaInt(left(CodigodeBarras.Text, 3))
        ComboCobrador.Text = CStr(iCodBanco)
        
        'Chama o Validate da ComboCobrador
        Call ComboCobrador_Validate(bSGECancelDummy)

    End If
    
    If Len(Trim(CodigodeBarras.ClipText)) <> 0 And Len(Trim(DataVencimento.ClipText)) = 0 Then
'
'        If Len(Trim(CodigodeBarras.ClipText)) = 44 Then
'
'            dNumDias = StrParaDbl(Mid(CodigodeBarras.ClipText, 6, 4))
'
'            If dNumDias <> 0 Then
'
'                dtDataVenc = DateAdd("d", dNumDias, CDate("7/10/1997"))
'                DataVencimento.Text = Format(dtDataVenc, "dd/mm/yy")
'
'            End If
'
'        Else
        
            dNumDias = StrParaDbl(Mid(CodigodeBarras.ClipText, 34, 4))
        
            If dNumDias <> 0 Then
        
                dtDataVenc = DateAdd("d", dNumDias, CDate("7/10/1997"))
                DataVencimento.Text = Format(dtDataVenc, "dd/mm/yy")
                
            End If
        
'        End If
        
    End If
    
    Call Combo_Seleciona_ItemData(ComboTipoCobranca, TIPO_COBRANCA_BANCARIA)
    
    Exit Sub

Erro_CodigodeBarras_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 188321
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158932)
    
    End Select
    
    Exit Sub

End Sub

'Fim Campos Referêntes ao Código de barras

Private Sub ComboCobrador_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCobrador As New ClassCobrador
Dim vbMsgRes As VbMsgBoxResult
Dim objBanco As New ClassBanco

On Error GoTo Erro_ComboCobrador_Validate

    'Verifica se o Cobrador foi preenchido
    If Len(Trim(ComboCobrador.Text)) = 0 Then Exit Sub
    
    'Verifica se é um Cobrador selecionado
    If ComboCobrador.Text = ComboCobrador.List(ComboCobrador.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(ComboCobrador, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 18966
    
    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then
        
        objBanco.iCodBanco = iCodigo
        
        'Le o Banco no BD
        lErro = CF("Banco_Le", objBanco)
        If lErro <> SUCESSO And lErro <> 16091 Then gError 58658
        
        'Se não encontrou mesmo assim
        If lErro = 16091 Then gError 58659
                  
        ComboCobrador.Text = objBanco.iCodBanco & SEPARADOR & objBanco.sNomeReduzido
    
    End If
    
    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 18970
    
    Exit Sub
    
Erro_ComboCobrador_Validate:
    
    Cancel = True
    
    Select Case gErr
    
       Case 18966, 58658 'Tratados nas rotinas chamadas
        
        Case 18970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", gErr, ComboCobrador.Text)
        
        Case 58659
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_BANCO", objBanco.iCodBanco)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("Bancos", objBanco)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158933)
    
    End Select
    
    Exit Sub

End Sub

Private Sub ComboPortador_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objPortador As New ClassPortador
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ComboPortador_Validate

    'Verifica se o Portador foi preenchido
    If Len(Trim(ComboPortador.Text)) = 0 Then Exit Sub
    
    'Verifica se é um Portador selecionado
    If ComboPortador.Text = ComboPortador.List(ComboPortador.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(ComboPortador, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 18971
    
    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then
            
        objPortador.iCodigo = iCodigo
        
        'Le o Portador no BD
        lErro = CF("Portador_Le", objPortador)
        If lErro <> SUCESSO And lErro <> 15971 Then gError 58656
        
        'Se não encontrou --> ERRO
        If lErro = 15971 Then gError 58657
        
        'Se encontrou põe na Tela
        ComboPortador.Text = objPortador.iCodigo & SEPARADOR & objPortador.sNomeReduzido
        
    End If
    
    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 18974
    
    Exit Sub
    
Erro_ComboPortador_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 18971, 58656 'Tratados nas rotinas chamadas
        
        Case 18974
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_ENCONTRADO", gErr, ComboPortador.Text)
        
        Case 58657
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PORTADOR", objPortador.iCodigo)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("Portadores", objPortador)
            End If
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158934)
    
    End Select
    
    Exit Sub

End Sub

Private Sub DataVencimento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataVencimento, iAlterado)

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then
        Call Limpa_Parcela
        Exit Sub
    End If
    
    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 18961
    
    'Se não encontra valor que era CÓDIGO
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 18962

        sFornecedor = Fornecedor.Text
        objFilialFornecedor.iCodFilial = iCodigo
        
        'Pesquisa se existe Filial do Fornecedor
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 18963
        
        'Se não encontrou a Filial do Fornecedor --> erro
        If lErro = 18272 Then
        
            objFornecedor.sNomeReduzido = sFornecedor
            
            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 58661
            
            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            
            'Sugere cadastrar nova Filial
            gError 18964
        
        End If
        
        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome
        
    End If
    
    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 18965
    
    Exit Sub
    
Erro_Filial_Validate:

    Cancel = True
    
    Select Case gErr
    
       Case 18961, 18963, 58661 'Tratados nas rotinas chamadas
       
       Case 18962
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
       Case 18964
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If
        
        Case 18965
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Filial.Text)
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158935)
    
    End Select
    
    Exit Sub

End Sub

Private Sub LabelNumParcela_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objParcelaPagar As New ClassParcelaPagar
Dim iFilial As Integer
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelNumParcela_Click

    If Len(Trim(NumTitulo.Text)) = 0 Then gError 57462
        
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 57481
    
    If Len(Trim(Filial.Text)) = 0 Then gError 57482
    
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 57483

    'Se não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 57484

    iFilial = Codigo_Extrai(Filial.Text)

    colSelecao.Add objFornecedor.lCodigo
    colSelecao.Add iFilial
    colSelecao.Add StrParaLong(NumTitulo.Text)
    
    'Chama a tela
    Call Chama_Tela("ParcelasPagDetPagLista", colSelecao, objParcelaPagar, objEventoParcela)
    
    Exit Sub
    
Erro_LabelNumParcela_Click:

    Select Case gErr
    
        Case 57462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
         
        Case 57481
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 57482
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 57483
            'Tratado na rotina chamada
            
        Case 57484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158936)
            
    End Select
    
    Exit Sub


End Sub

Private Sub LabelTipoDoc_Click()
   
Dim objTipoDocumento As New ClassTipoDocumento
Dim colSelecao As Collection

    objTipoDocumento.sSigla = Tipo.Caption

    'Chama a Tela de browse
    Call Chama_Tela("TiposDocBaixaPagCancLista", colSelecao, objTipoDocumento, objEventoTipoDoc)

    Exit Sub
    
End Sub

Private Sub NumeroLabel_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim objParcelaPagar As New ClassParcelaPagar
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_NumeroLabel_Click

    'Se Forncedor estiver vazio, erro
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 26014
    
    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then gError 26015
    
    'Move os dados da Tela para objTituloPagar e objParcelaPagar
    lErro = Move_Tela_Memoria(objTituloPagar, objParcelaPagar)
    If lErro <> SUCESSO Then gError 26017
    
    'Adiciona filtros: lFornecedor e iFilial
    colSelecao.Add objTituloPagar.lFornecedor
    colSelecao.Add objTituloPagar.iFilial
        
    'Adciona Seleção Dinâmica
    sSelecao = " Fornecedor = ? AND Filial=? "

    'Chama Tela OutrosPagLista
    Call Chama_Tela("TitPagTodosLista", colSelecao, objTituloPagar, objEventoNumero, sSelecao)
    
    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 26014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 26015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
    
        Case 26016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
        
        Case 26017 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158937)

    End Select

    Exit Sub

End Sub

Private Sub NumTitulo_GotFocus()
    
Dim glNumAux As Long
    
    glNumAux = glNumIntParc
    Call MaskEdBox_TrataGotFocus(NumTitulo, iAlterado)
    glNumIntParc = glNumAux

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloPagar As ClassTituloPagar, Cancel As Boolean
Dim bCancel As Boolean

    Set objTituloPagar = obj1

    Call Limpa_Tela_DetPag
    
    Fornecedor.Text = objTituloPagar.lFornecedor
    Call Fornecedor_Validate(Cancel)
    
    Filial.Text = objTituloPagar.iFilial
    Call Filial_Validate(bCancel)
    
    Tipo.Caption = objTituloPagar.sSiglaDocumento
    
    If objTituloPagar.dtDataEmissao <> DATA_NULA Then
        DataEmissao.Caption = Format(objTituloPagar.dtDataEmissao, "dd/mm/yy")
    Else
        DataEmissao.Caption = ""
    End If
    
    NumTitulo.Text = CStr(objTituloPagar.lNumTitulo)
    
    iAlterado = 0
    
    Me.Show

    Exit Sub

End Sub

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche nomeReduzido com o fornecedor da tela
    If Len(Trim(Fornecedor.Text)) > 0 Then objFornecedor.sNomeReduzido = Fornecedor.Text

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor, Cancel As Boolean

    Set objFornecedor = obj1

    'Preenche campo Fornecedor
    Fornecedor.Text = objFornecedor.sNomeReduzido
        
    Call Fornecedor_Validate(Cancel)
    
    Me.Show

    Exit Sub

End Sub

Private Sub DataVencimento_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)

Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim lErro As Long

On Error GoTo Erro_DataVencimento_Validate

    'Verifica se a data de vencimento foi digitada
    If Len(Trim(DataVencimento.ClipText)) = 0 Then Exit Sub

    'Faz a critica da data digitada
    lErro = Data_Critica(DataVencimento.Text)
    If lErro <> SUCESSO Then gError 18978

    'Verifica se a data de emissao foi digitada
    If Not IsDate(DataEmissao.Caption) Then Exit Sub

    dtDataEmissao = CDate(DataEmissao.Caption)
    dtDataVencimento = CDate(DataVencimento.Text)

    'Verifica se a data de vencimento é menor que a data de emissão
    If dtDataVencimento < dtDataEmissao Then gError 18979

    Exit Sub

Erro_DataVencimento_Validate:

    Cancel = True

    Select Case gErr

        Case 18978

        Case 18979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_MENOR", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158938)

    End Select

    Exit Sub

End Sub

Private Sub objEventoParcela_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objParcelaPagar As ClassParcelaPagar, bTrazerDadosParcela As Boolean

On Error GoTo Erro_objEventoParcela_evSelecao

    Set objParcelaPagar = obj1

    bTrazerDadosParcela = False
    
    If Not (objParcelaPagar Is Nothing) Then
        Parcela.PromptInclude = False
        Parcela.Text = CStr(objParcelaPagar.iNumParcela)
        Parcela.PromptInclude = True
        Call Parcela_Validate(bSGECancelDummy)
        If bSGECancelDummy = False Then bTrazerDadosParcela = True
    End If
    
    Me.Show

    If bTrazerDadosParcela Then Call TrazerParcela_Click
    
    Exit Sub

Erro_objEventoParcela_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoDoc_evSelecao(obj1 As Object)

Dim objTipoDocumento As ClassTipoDocumento

    Set objTipoDocumento = obj1

    'Preenche o campo Tipo
    Tipo.Caption = objTipoDocumento.sSigla
        
    Me.Show

    Exit Sub

End Sub

Private Sub Parcela_GotFocus()

Dim glNumAux As Long
Dim iParcelaAux As Integer
    
    iParcelaAux = iParcelaAlterada
    glNumAux = glNumIntParc
    
    Call MaskEdBox_TrataGotFocus(Parcela, iAlterado)
    glNumIntParc = glNumAux
    iParcelaAlterada = iParcelaAux
    
End Sub

Private Sub TrazerParcela_Click()

Dim lErro As Long

On Error GoTo Erro_TrazerParcela_Click

    'Verifica se o Fornecedor, Filial e Parcela foram preenchidos
    If Len(Trim(Fornecedor.Text)) = 0 Or Len(Trim(Filial.Text)) = 0 Or Len(Trim(NumTitulo.Text)) = 0 Or Len(Trim(Parcela.Text)) = 0 Then gError 56583
    
    lErro = Carrega_Parcela()
    If lErro <> SUCESSO Then gError 18994

    Exit Sub
    
Erro_TrazerParcela_Click:

    Select Case gErr

        Case 18994 'tratado na rotina chamada

        Case 56583
            Call Rotina_Erro(vbOKOnly, "ERRO_DETPAG_CPO_NAO_PREENCHIDO", gErr)
            
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158940)

    End Select
    
    Exit Sub
    
End Sub

Private Sub UpDownVencimento_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencimento_DownClick
    
    'Diminui a data
    lErro = Data_Up_Down_Click(DataVencimento, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 26018

    Exit Sub

Erro_UpDownVencimento_DownClick:

    Select Case gErr

        Case 26018 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158941)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimento_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencimento_UpClick
    
    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataVencimento, AUMENTA_DATA)
    If lErro Then gError 43054

    Exit Sub

Erro_UpDownVencimento_UpClick:

    Select Case gErr

        Case 43054 'Tratadao na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158942)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Parcela() As Long
'Traz para a tela os dados da Parcela

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim objParcelaPagar As New ClassParcelaPagar
Dim bCancel As Boolean, bBaixada As Boolean

On Error GoTo Erro_Carrega_Parcela
        
    bBaixada = False
    
    'Recolhe os dados da Tela
    lErro = Move_Tela_Memoria(objTituloPagar, objParcelaPagar)
    If lErro <> SUCESSO Then gError 18981
    
    If objTituloPagar.sSiglaDocumento = "" Then
    
        'Lê o Título
        lErro = CF("TituloPagar_Le_Titulo", objTituloPagar)
        If lErro <> SUCESSO And lErro <> 43451 Then gError 18982
        
        If lErro <> SUCESSO Then
        
            'Lê o Titulo Baixado
            lErro = CF("TituloPagarBaixado_Le_Titulo", objTituloPagar)
            If lErro <> SUCESSO And lErro <> 113482 Then gError 113483
        
            bBaixada = True
            
        End If
        
    Else
    
        'Lê o Título levando em conta a sigla do docto
        lErro = CF("TituloPagar_Le_Numero", objTituloPagar)
        If lErro <> SUCESSO And lErro <> 18551 Then gError 59294
        
        If lErro <> SUCESSO Then
            
            'Lê o Título Baixado levando em conta a sigla do docto
            lErro = CF("TituloPagarBaixado_Le_Numero", objTituloPagar)
            If lErro <> SUCESSO And lErro <> 18556 Then gError 113485
        
            bBaixada = True
            
        End If
        
    End If
    
    'Se não encontrar --> Erro
    If lErro <> SUCESSO Then gError 18988
    
    objParcelaPagar.lNumIntTitulo = objTituloPagar.lNumIntDoc
    
    If bBaixada Then
    
        'Lê a Parcela referente ao Título lido
        lErro = CF("ParcelaPagarBaixada_Le_Numero", objParcelaPagar)
        If lErro <> SUCESSO And lErro <> 43625 Then gError 18983
    
    Else
    
        'Lê a Parcela referente ao Título lido
        lErro = CF("ParcelaPagar_Le_Numero", objParcelaPagar)
        If lErro <> SUCESSO And lErro <> 18987 Then gError 18983
    
    End If
    
    If lErro <> SUCESSO Then gError 69045
    
    'Se a parcela nao estiver aberta --> Erro
    If objParcelaPagar.iStatus <> STATUS_ABERTO And objParcelaPagar.iStatus <> STATUS_BAIXADO Then gError 26013
    
    'Põe na Tela os Dados lidos da Parcela
    glNumIntParc = objParcelaPagar.lNumIntDoc
    
    If objTituloPagar.dtDataEmissao <> DATA_NULA Then
        DataEmissao.Caption = Format(objTituloPagar.dtDataEmissao, "dd/mm/yy")
    Else
        DataEmissao.Caption = ""
    End If
    
    If objTituloPagar.sSiglaDocumento <> "" Then
        Tipo.Caption = objTituloPagar.sSiglaDocumento
    Else
        Tipo.Caption = ""
    End If
    
    Valor.Text = Format(objParcelaPagar.dValor, "Standard")
    
    ValorOriginal.Caption = Format(objParcelaPagar.dValorOriginal, "Standard")
    If objParcelaPagar.iMotivoDiferenca <> 0 Then
        MotivoDiferenca.Text = objParcelaPagar.iMotivoDiferenca
        Call MotivoDiferenca_Validate(bSGECancelDummy)
    End If
    
    ComboTipoCobranca.Text = objParcelaPagar.iTipoCobranca
    Call ComboTipoCobranca_Validate(bSGECancelDummy)
    
    If objParcelaPagar.iBancoCobrador <> 0 Then
        ComboCobrador.Text = objParcelaPagar.iBancoCobrador
        Call ComboCobrador_Validate(bCancel)
    Else
        ComboCobrador.Text = ""
    End If
    
    If objParcelaPagar.iPortador <> 0 Then
        ComboPortador.Text = objParcelaPagar.iPortador
        Call ComboPortador_Validate(bCancel)
    Else
        ComboPortador.Text = ""
    End If
    
    DataVencimento.PromptInclude = False
    DataVencimento.Text = Format(objParcelaPagar.dtDataVencimento, "dd/mm/yy")
    DataVencimento.PromptInclude = True
    
    'Carrega o Controle relacionado com o Código de barras
    CodigodeBarras.PromptInclude = False
    CodigodeBarras.Text = objParcelaPagar.sCodigoDeBarras
    CodigodeBarras.PromptInclude = True
    
    Call Desabilita_Campos_ParcBaixada(bBaixada)
    
    iAlterado = 0
    
    Carrega_Parcela = SUCESSO
    
    Exit Function
    
Erro_Carrega_Parcela:

    Carrega_Parcela = gErr
    
    Select Case gErr
    
        Case 18981, 18982, 59294, 113483, 113485 'Tratado na rotina chamada

        Case 18983
            Parcela.Text = ""
            Call Limpa_Parcela
    
        Case 18988
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", gErr)
        
        Case 26013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_PAG_NAO_ABERTA2", gErr, objParcelaPagar.iNumParcela, objTituloPagar.lNumTitulo)
        
        Case 69045
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_PAGAR_NAO_CADASTRADA", gErr, objParcelaPagar.iNumParcela)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158943)
            
    End Select
    
    Exit Function
    
End Function

Private Sub Filial_Change()

    glNumIntParc = 0
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

    glNumIntParc = 0
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ComboCobrador_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ComboCobrador_Click()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ComboPortador_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ComboPortador_Click()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub


Private Sub ComboTipoCobranca_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ComboTipoCobranca_Click()
    
Dim iCodigoAux As Integer
    
    iAlterado = REGISTRO_ALTERADO
    
    'Retira o Código da List da Combo
    iCodigoAux = Codigo_Extrai(ComboTipoCobranca.Text)
    'Verifica se é o código referênte a Cobrança bancária.
    'Se for habilita o Campo referênte ao código de barras
    If iCodigoAux = TIPO_COBRANCA_BANCARIA Then
        
        CodigodeBarras.Enabled = True
        CodigoBarras.Enabled = True
        Label1.Enabled = True
        ComboCobrador.Enabled = True
        
    Else
    
        CodigodeBarras.PromptInclude = False
        CodigodeBarras.Text = ""
        CodigodeBarras.PromptInclude = True
        CodigodeBarras.Enabled = False
        CodigoBarras.Enabled = False
        ComboCobrador.ListIndex = -1
        ComboCobrador.Enabled = False
        Label1.Enabled = False

    End If
        
End Sub

Private Sub ComboTipoCobranca_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoCobranca

On Error GoTo Erro_ComboTipoCobranca_Validate

    'Verifica se o Tipo de Cobrança foi preenchido
    If Len(Trim(ComboTipoCobranca.Text)) = 0 Then Exit Sub

    'Verifica se ele foi selecionado
    If ComboTipoCobranca.Text = ComboTipoCobranca.List(ComboTipoCobranca.ListIndex) Then Exit Sub
    
    'Seleciona o Tipo de Cobrança
    lErro = Combo_Seleciona(ComboTipoCobranca, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 18975
    
    'Não encontrou o CODIGO
    If lErro = 6730 Then gError 18976
    
    'Não encontrou a STRING
    If lErro = 6731 Then gError 18977
           
    Exit Sub
    
Erro_ComboTipoCobranca_Validate:

    Cancel = True


    Select Case gErr
            
        Case 18975

        Case 18976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_CADASTRADO", gErr, iCodigo)
                
        Case 18977
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_ENCONTRADO", gErr, ComboTipoCobranca.Text)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158944)
            
    End Select
    
    Exit Sub

End Sub

Private Sub NumTitulo_Change()
    
    glNumIntParc = 0
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NumTitulo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumTitulo_Validate

    'Verifica se o Numero foi preenchido
    If Len(Trim(NumTitulo.ClipText)) = 0 Then
        Call Limpa_Parcela
        Exit Sub
    End If

    'Critica se é Long positivo
    lErro = Long_Critica(NumTitulo.ClipText)
    If lErro <> SUCESSO Then gError 18995
            
    Exit Sub

Erro_NumTitulo_Validate:

    Cancel = True


    Select Case gErr

        Case 18995

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158945)

    End Select

    Exit Sub

End Sub

Private Sub Parcela_Change()
    
    glNumIntParc = 0
    iAlterado = REGISTRO_ALTERADO
    iParcelaAlterada = 1
    
End Sub

Private Sub Parcela_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Parcela_Validate

    'Se a Parcela não foi alterada
    If iParcelaAlterada = 0 Then Exit Sub
    
    'Se a parcela não foi preenchida
    If Len(Trim(Parcela.Text)) = 0 Then
        'Limpa
        Call Limpa_Parcela
        Exit Sub
    End If
    
    'Critica se é um numero inteiro
    lErro = Inteiro_Critica(Parcela.Text)
    If lErro <> SUCESSO Then gError 18997
    
    'Zera a Flag
    iParcelaAlterada = 0
    
    Exit Sub

Erro_Parcela_Validate:

    Cancel = True


    Select Case gErr
    
        Case 18997
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158946)
    
    End Select

    Exit Sub

End Sub

Private Sub Fornecedor_Change()

    glNumIntParc = 0
    iFornecedorAlterado = 1
    iAlterado = REGISTRO_ALTERADO
    
    Call Fornecedor_Preenche
    
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate
    
    'Se o Fornecedor foi alterado
    If iFornecedorAlterado = 1 Then

        'Verifica se o Fornecedor foi preenchido
        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 18959

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 18960

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then
            
            'Limpa Combo de Filial
            Filial.Clear

        End If

        'Zera a Flag
        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 18959, 18960 'Tratados nas rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158947)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 26012
    
    'Limpa a Tela
    Call Limpa_Tela_DetPag
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr
    
        Case 26012 'Tratado na rotina chamada
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158948)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Limpa_Tela_DetPag()

    'Chama função que limpa TextBoxes e MaskedEdits da Tela
    Call Limpa_Tela(Me)
    
    Call Desabilita_Campos_ParcBaixada(False)
    
    'Limpa os campos não são limpos pela função acima
    Filial.Clear
    ComboTipoCobranca.Text = ""
    ComboPortador.Text = ""
    ComboCobrador.Text = ""
    DataEmissao.Caption = ""
    Tipo.Caption = ""
    iParcelaAlterada = 0
    
    MotivoDiferenca.ListIndex = -1
    ValorOriginal.Caption = ""
    
    iAlterado = 0
    
End Sub

Private Function Move_Tela_Memoria(objTituloPagar As ClassTituloPagar, objParcelaPagar As ClassParcelaPagar) As Long
'Move os dados da Tela para objtituloPagar e objParcelaPagar

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim dtdataAux As Date

On Error GoTo Erro_Move_Tela_Memoria

    objFornecedor.sNomeReduzido = Fornecedor.Text
    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 18989
    
    'Não encontrou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 18990
        
    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    
    objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    
    If Len(Trim(NumTitulo.ClipText)) > 0 Then objTituloPagar.lNumTitulo = CLng(NumTitulo.ClipText)
    
    If Len(Trim(DataEmissao.Caption)) > 0 Then
        objTituloPagar.dtDataEmissao = CDate(DataEmissao.Caption)
    Else
        objTituloPagar.dtDataEmissao = DATA_NULA
    End If
    
    objTituloPagar.sSiglaDocumento = Tipo.Caption
    objTituloPagar.iFilialEmpresa = giFilialEmpresa
    
    If Len(Trim(Valor.Text)) > 0 Then objParcelaPagar.dValor = StrParaDbl(Valor.Text)

    objParcelaPagar.dValorOriginal = StrParaDbl(ValorOriginal.Caption)
    objParcelaPagar.iMotivoDiferenca = Codigo_Extrai(MotivoDiferenca.Text)

    If Len(Trim(Parcela.Text)) > 0 Then objParcelaPagar.iNumParcela = CInt(Parcela.Text)
    If Len(Trim(ComboTipoCobranca.Text)) > 0 Then objParcelaPagar.iTipoCobranca = ComboTipoCobranca.ItemData(ComboTipoCobranca.ListIndex)
    If Len(Trim(ComboCobrador.Text)) > 0 Then objParcelaPagar.iBancoCobrador = Codigo_Extrai(ComboCobrador.Text)
    If Len(Trim(ComboPortador.Text)) > 0 Then objParcelaPagar.iPortador = Codigo_Extrai(ComboPortador.Text)
    
    If Len(Trim(DataVencimento.ClipText)) > 0 Then
        
        objParcelaPagar.dtDataVencimento = CDate(DataVencimento.Text)
        
        'verifica a data para vencimento real
        lErro = CF("DataVencto_Real", objParcelaPagar.dtDataVencimento, dtdataAux)
        If lErro <> SUCESSO Then gError 19433
        
        objParcelaPagar.dtDataVencimentoReal = dtdataAux
    End If
     
    objParcelaPagar.sCodigoDeBarras = Trim(CodigodeBarras.ClipText)

    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 18989, 19433 'Tratados nas rotinas chamadas
        
        Case 18990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158949)
    
    End Select
    
    Exit Function

End Function

Private Sub Limpa_Parcela()
'Limpa na tela os campos que se referem a uma parcela

    DataEmissao.Caption = ""
    Tipo.Caption = ""
    ComboTipoCobranca.Text = ""
    ComboCobrador.Text = ""
    ComboPortador.Text = ""
    DataVencimento.PromptInclude = False
    DataVencimento.Text = ""
    DataVencimento.PromptInclude = True
    
     MotivoDiferenca.ListIndex = -1
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)
        
    Set objEventoNumero = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoTipoDoc = Nothing
    Set objEventoParcela = Nothing
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 26009
    
    Call Limpa_Tela_DetPag
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 26009
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158950)
            
    End Select
    
    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objParcelaPagar As New ClassParcelaPagar
Dim objTituloPagar As New ClassTituloPagar

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios da tela estao preenchidos
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 26001
    If Len(Trim(Filial.Text)) = 0 Then gError 26002
    If Len(Trim(NumTitulo.Text)) = 0 Then gError 26003
    If Len(Trim(Parcela.Text)) = 0 Then gError 26005
    If Len(Trim(ComboTipoCobranca.Text)) = 0 Then gError 26007
    If Len(Trim(DataVencimento.ClipText)) = 0 Then gError 26006

    If Len(Trim(MotivoDiferenca.Text)) = 0 And StrParaDbl(Valor.Text) <> StrParaDbl(ValorOriginal.Caption) Then gError 500069
    If Len(Trim(MotivoDiferenca.Text)) > 0 And StrParaDbl(ValorOriginal.Caption) = StrParaDbl(Valor.Text) Then gError 500070

    If glNumIntParc = 0 Then gError 49527
   
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objTituloPagar, objParcelaPagar)
    If lErro <> SUCESSO Then gError 26008
    
    'Verifica se Data de Vencimento é Menor que a Data de Emissão
    If objTituloPagar.dtDataEmissao <> DATA_NULA Then
        If CDate(objParcelaPagar.dtDataVencimento) < CDate(objTituloPagar.dtDataEmissao) Then gError 26008
    End If
    
    If objParcelaPagar.iTipoCobranca = TIPO_COBRANCA_DEP_CONTA Or objParcelaPagar.iTipoCobranca = TIPO_COBRANCA_DOC Or objParcelaPagar.iTipoCobranca = TIPO_COBRANCA_OP Then
    
        lErro = CF("Fornecedor_Verifica_DadosCta", objTituloPagar.lFornecedor, objTituloPagar.iFilial)
        If lErro <> SUCESSO Then gError 106576
        
    End If
    
    objParcelaPagar.lNumIntDoc = glNumIntParc
       
    'Faz as modificações nas Parcelas no BD
    lErro = CF("ParcelasPag_Modificar_DetPag", objParcelaPagar)
    If lErro <> SUCESSO And lErro <> 7799 Then gError 7802
    
    If lErro = 7799 Then

        'Função que Lê uma Altera uma Parcela de Um Título já baixado
        lErro = CF("ParcelasPagBaixadas_Modificar_DetPag", objParcelaPagar)
        If lErro <> SUCESSO And lErro <> 113483 Then gError 113486
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 7802, 26008, 106576, 113486 'Tratados nas rotinas chamadas
        
        Case 26001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 26002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 26003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
            
        Case 26005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_PREENCHIDA", gErr)
            
        Case 26006
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_NAO_INFORMADA", gErr, Parcela.Text)
        
        Case 26007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_PREENCHIDO", gErr)
            
        Case 26009
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", gErr, DataVencimento.Text, DataEmissao.Caption, objParcelaPagar.iNumParcela)
            
        Case 49527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TECLAR_BOTAO_TRAZER", gErr)
        
        Case 500069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOTIVODIFERENCAPARC_NAO_INFORMADO", gErr, Parcela.Text)

        Case 500070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOTIVODIFERENCA_INFORMADO_ERRADO", gErr, Parcela.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158951)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONFIRMACAO_COBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Confirmação de Cobrança"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "DetPag"
    
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

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is NumTitulo Then
            Call NumeroLabel_Click
        End If
    
    End If
    
End Sub
Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub LabelTipoDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipoDoc, Source, X, Y)
End Sub

Private Sub LabelTipoDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipoDoc, Button, Shift, X, Y)
End Sub

Private Sub Tipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Tipo, Source, X, Y)
End Sub

Private Sub Tipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Tipo, Button, Shift, X, Y)
End Sub

Private Sub DataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissao, Source, X, Y)
End Sub

Private Sub DataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissao, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Cobranca_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cobranca, Source, X, Y)
End Sub

Private Sub Cobranca_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cobranca, Button, Shift, X, Y)
End Sub

Private Sub Portador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Portador, Source, X, Y)
End Sub

Private Sub Portador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Portador, Button, Shift, X, Y)
End Sub

Private Sub LabelNumParcela_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumParcela, Source, X, Y)
End Sub

Private Sub LabelNumParcela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumParcela, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Sub Desabilita_Campos_ParcBaixada(ByVal bBaixada As Boolean)
'habilita ou desabilita os campos nao podem ser editados qdo uma parcela baixada for carregada

Dim lErro As Long, bHabilita As Boolean

On Error GoTo Erro_Desabilita_Campos

    bHabilita = (bBaixada = False)
    
    Label4.Enabled = bHabilita
    DataVencimento.Enabled = bHabilita
    UpDownVencimento.Enabled = bHabilita
    Portador.Enabled = bHabilita
    ComboPortador.Enabled = bHabilita

    LabelValor.Enabled = bHabilita
    LabelMotivoDiferenca.Enabled = bHabilita
    Valor.Enabled = bHabilita
    MotivoDiferenca.Enabled = bHabilita
    
    Exit Sub

Erro_Desabilita_Campos:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158952)
            
    End Select
    
    Exit Sub
    
End Sub



Private Sub Fornecedor_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134047

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134047

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158953)

    End Select
    
    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Se o Valor foi preenchido
    If Len(Trim(Valor.Text)) > 0 Then
        
        'Critica o valor
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 500069
    
    End If
    
    Exit Sub

Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 500069
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Exit Sub

End Sub

Private Sub MotivoDiferenca_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_MotivoDiferenca_Validate
    
    'Verifica se o MotivoDiferenca está preenchido
    If Len(Trim(MotivoDiferenca.Text)) > 0 Then
    
        If MotivoDiferenca.Text <> MotivoDiferenca.List(MotivoDiferenca.ListIndex) Then

            lErro = Combo_Item_Seleciona(MotivoDiferenca)
            If lErro <> SUCESSO And lErro <> 12250 Then gError 500067
    
            'Se não encontrou o Motivo de Diferença nem com o Código nem Descrição ==> erro
            If lErro <> SUCESSO Then gError 500068
            
        End If

    End If
    
    Exit Sub
    
Erro_MotivoDiferenca_Validate:

    Cancel = True
    
    Select Case gErr
            
        Case 500067
            
        Case 500068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOTIVODIFERENCA_NAO_ENCONTRADO", gErr, MotivoDiferenca.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Exit Sub
    
End Sub

Private Function Carrega_MotivoDiferenca() As Long
'Carrega a combo de Cobrador

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_MotivoDiferenca

    'Leitura dos códigos e descrições dos Bancos BD
    lErro = CF("Cod_Nomes_Le", "MotivoDiferenca", "Codigo", "Descricao", STRING_DESCRICAO_CAMPO, colCodigoNome)
    If lErro <> SUCESSO Then gError 500066

   'Preenche ComboBox com código e nome dos Bancos
    For iIndice = 1 To colCodigoNome.Count
        Set objCodigoNome = colCodigoNome(iIndice)
        MotivoDiferenca.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        MotivoDiferenca.ItemData(MotivoDiferenca.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_MotivoDiferenca = SUCESSO

    Exit Function

Erro_Carrega_MotivoDiferenca:

    Carrega_MotivoDiferenca = gErr

    Select Case gErr

        Case 500066 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function
