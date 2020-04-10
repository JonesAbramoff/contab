VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BorderoPag1Ocx 
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5490
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   5490
   Begin VB.Frame Frame3 
      Caption         =   "Geração"
      Height          =   1110
      Left            =   120
      TabIndex        =   24
      Top             =   75
      Width           =   5250
      Begin VB.ComboBox ContaCorrente 
         Height          =   315
         Left            =   915
         TabIndex        =   0
         Top             =   225
         Width           =   2820
      End
      Begin MSComCtl2.UpDown UpDownDataEmissao 
         Height          =   300
         Left            =   2055
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   930
         TabIndex        =   1
         Top             =   630
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataContabil 
         Height          =   300
         Left            =   4920
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   630
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataContabil 
         Height          =   300
         Left            =   3795
         TabIndex        =   3
         Top             =   630
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "Data Contábil:"
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
         Left            =   2490
         TabIndex        =   27
         Top             =   675
         Width           =   1395
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   90
         TabIndex        =   26
         Top             =   675
         Width           =   855
      End
      Begin VB.Label LabelContaCorrente 
         Caption         =   "Conta:"
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
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   255
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   1905
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4620
      Width           =   1680
      Begin VB.CommandButton BotaoSeguir 
         Height          =   330
         Left            =   90
         Picture         =   "BorderoPag1Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   90
         Width           =   930
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1125
         Picture         =   "BorderoPag1Ocx.ctx":0792
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro de Títulos por"
      Height          =   3270
      Left            =   120
      TabIndex        =   13
      Top             =   1230
      Width           =   5235
      Begin VB.TextBox ValorMaxBordero 
         Height          =   345
         Left            =   1905
         TabIndex        =   6
         Top             =   855
         Width           =   1485
      End
      Begin VB.ComboBox TipoCobranca 
         Height          =   315
         Left            =   1905
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1440
         Width           =   2865
      End
      Begin VB.CheckBox CheckDepositoOutroBanco 
         Caption         =   "Permitir depósito em conta de outro banco"
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
         Height          =   270
         Left            =   435
         TabIndex        =   10
         Top             =   2715
         Width           =   3930
      End
      Begin VB.Frame Frame2 
         Caption         =   "Liquidação de títulos"
         Height          =   645
         Left            =   180
         TabIndex        =   16
         Top             =   1950
         Width           =   4890
         Begin VB.OptionButton OptionAmbos 
            Caption         =   "ambos"
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
            Height          =   225
            Left            =   3870
            TabIndex        =   20
            Top             =   300
            Width           =   885
         End
         Begin VB.OptionButton OptionProprioBanco 
            Caption         =   "no próprio banco"
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
            Height          =   225
            Left            =   120
            TabIndex        =   8
            Top             =   300
            Width           =   1800
         End
         Begin VB.OptionButton OptionOutroBanco 
            Caption         =   "em outro banco"
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
            Height          =   225
            Left            =   1995
            TabIndex        =   9
            Top             =   300
            Width           =   1644
         End
      End
      Begin MSComCtl2.UpDown UpDownDataVencimento 
         Height          =   300
         Left            =   4920
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataVencimento 
         Height          =   300
         Left            =   3795
         TabIndex        =   5
         Top             =   345
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataVencimentoDe 
         Height          =   300
         Left            =   3000
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   345
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataVencimentoDe 
         Height          =   300
         Left            =   1905
         TabIndex        =   22
         Top             =   345
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento De:"
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
         TabIndex        =   23
         Top             =   390
         Width           =   1365
      End
      Begin VB.Label Label1 
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
         Left            =   3405
         TabIndex        =   17
         Top             =   390
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valor Máx. Bordero:"
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
         Left            =   150
         TabIndex        =   18
         Top             =   945
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cobrança:"
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
         Left            =   255
         TabIndex        =   19
         Top             =   1470
         Width           =   1590
      End
   End
End
Attribute VB_Name = "BorderoPag1Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variaveis globais
Dim iBorderoAlterado As Integer
Dim gobjBorderoPagEmissao As ClassBorderoPagEmissao

'Browsers
Private WithEvents objEventoContaCorrente As AdmEvento
Attribute objEventoContaCorrente.VB_VarHelpID = -1

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoSeguir_Click()

Dim lErro As Long, dtDataDe As Date
Dim iCodConta As Integer
Dim iTipoCobranca As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_BotaoSeguir_Click

    'Verifica se a ContaCorrente está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then Error 15741

    'Extrai o código da Conta
    iCodConta = Codigo_Extrai(ContaCorrente.Text)

    objContaCorrenteInt.iCodigo = iCodConta

    'Lê os dados da Conta
    lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 15742

    'Se a Conta não estiver cadastrada
    If lErro = 11807 Then Error 15743

    'Se a Conta não é Bancária
    If objContaCorrenteInt.iCodBanco = 0 Then Error 15744

    'Se alguma Filial tiver sido selecionada
    If giFilialEmpresa <> EMPRESA_TODA Then

        'Se a Conta não é da Filial selecionada
        If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 15745

    End If

    'Verifica se a DataEmissao está preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Error 15747

    'Verifica se a DataVencimento está preenchida
    If Len(Trim(DataVencimento.ClipText)) = 0 Then Error 15749

    dtDataDe = MaskedParaDate(DataVencimentoDe)
    
    If dtDataDe <> DATA_NULA And dtDataDe > MaskedParaDate(DataVencimento) Then Error 32301
    
    'Verifica se o TipoCobrança está preenchido
    If Len(Trim(TipoCobranca.Text)) = 0 Then Error 15751
    
    'Extrai o Código do Tipo de Cobrança que está na tela
    iTipoCobranca = Codigo_Extrai(TipoCobranca.Text)

    If gobjBorderoPagEmissao.colInfoParcPag.Count = 0 Or iBorderoAlterado = REGISTRO_ALTERADO Then
        Set gobjBorderoPagEmissao = New ClassBorderoPagEmissao
    End If

    'Mover os dados da tela p/gobjBorderoPagEmissao
    gobjBorderoPagEmissao.iPodeDOCOutroEstado = 1
    gobjBorderoPagEmissao.dtEmissao = CDate(DataEmissao.Text)
    gobjBorderoPagEmissao.dtVencto = CDate(DataVencimento.Text)
    gobjBorderoPagEmissao.dtVenctoDe = dtDataDe
    
    If StrParaDate(DataContabil.Text) = DATA_NULA Then
        gobjBorderoPagEmissao.dtContabil = CDate(DataEmissao.Text)
    Else
        gobjBorderoPagEmissao.dtContabil = CDate(DataContabil.Text)
    End If
    
    'gobjBorderoPagEmissao.dtContabil = gobjBorderoPagEmissao.dtVencto
    
    If Len(Trim(ValorMaxBordero.Text)) <> 0 Then

        gobjBorderoPagEmissao.dValorMaximo = CDbl(ValorMaxBordero.Text)

    Else

        gobjBorderoPagEmissao.dValorMaximo = 0

    End If
    
    gobjBorderoPagEmissao.iCta = iCodConta
    
    If OptionOutroBanco.Value = True Then
        gobjBorderoPagEmissao.iLiqTitOutroBco = LIQUID_TITULO_OUTRO_BANCO
    ElseIf OptionAmbos.Value = True Then
        gobjBorderoPagEmissao.iLiqTitOutroBco = LIQUID_TITULO_AMBOS_BANCO
    Else
        gobjBorderoPagEmissao.iLiqTitOutroBco = NAO_LIQUID_TITULO_OUTRO_BANCO
    End If
    
    If CheckDepositoOutroBanco.Value = 1 Then
        gobjBorderoPagEmissao.iPodeDepCtaOutroBco = PERMITE_DEP_OUTRO_BANCO
    Else
        gobjBorderoPagEmissao.iPodeDepCtaOutroBco = NAO_PERMITE_DEP_OUTRO_BANCO
    End If
        
    gobjBorderoPagEmissao.iTipoCobranca = iTipoCobranca
    gobjBorderoPagEmissao.lNumero = 0

    If gobjBorderoPagEmissao.colInfoParcPag.Count = 0 Or iBorderoAlterado = REGISTRO_ALTERADO Then

        'Preencher coleção de Parcelas à pagar
        lErro = CF("ParcelasPagar_Le_BorderoPag", gobjBorderoPagEmissao)
        If lErro <> SUCESSO Then Error 15779
    
    End If
    
    'Chama a tela do passo seguinte
    Call Chama_Tela("BorderoPag2", gobjBorderoPagEmissao)

    'Fecha a tela
    Unload Me
    
    Exit Sub

Erro_BotaoSeguir_Click:

    Select Case Err

        Case 15741
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)

        Case 15742, 15779

        Case 15743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, iCodConta)

        Case 15744
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", Err, ContaCorrente.Text)

        Case 15745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaCorrente.Text, giFilialEmpresa)

        Case 15747, 15749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 32301
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR_FINAL", Err)
        
        Case 15751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_COBRANCA_NAO_INFORMADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143780)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objBorderoPagEmissao As ClassBorderoPagEmissao) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If (objBorderoPagEmissao Is Nothing) Then Error 7310

    Set gobjBorderoPagEmissao = objBorderoPagEmissao
    
    If objBorderoPagEmissao.colInfoParcPag.Count > 0 Then
        
        If objBorderoPagEmissao.iCta <> 0 Then
        
            ContaCorrente.Text = objBorderoPagEmissao.iCta
            Call ContaCorrente_Validate(bSGECancelDummy)
        
        End If
        
        DataEmissao.Text = Format(objBorderoPagEmissao.dtEmissao, "dd/mm/yy")
        DataVencimento.Text = Format(objBorderoPagEmissao.dtVencto, "dd/mm/yy")
        Call DateParaMasked(DataVencimentoDe, objBorderoPagEmissao.dtVenctoDe)
        
        If objBorderoPagEmissao.dtEmissao <> objBorderoPagEmissao.dtContabil Then
            DataContabil.Text = Format(objBorderoPagEmissao.dtContabil, "dd/mm/yy")
        End If
        
        ValorMaxBordero.Text = objBorderoPagEmissao.dValorMaximo
        
        For iIndice = 0 To TipoCobranca.ListCount - 1
            
            If TipoCobranca.ItemData(iIndice) = objBorderoPagEmissao.iTipoCobranca Then
                TipoCobranca.ListIndex = iIndice
                Exit For
            End If
        Next
        
        If gobjBorderoPagEmissao.iLiqTitOutroBco = LIQUID_TITULO_OUTRO_BANCO Then
            OptionOutroBanco.Value = True
        ElseIf gobjBorderoPagEmissao.iLiqTitOutroBco = LIQUID_TITULO_AMBOS_BANCO Then
            OptionAmbos.Value = True
        ElseIf gobjBorderoPagEmissao.iLiqTitOutroBco = NAO_LIQUID_TITULO_OUTRO_BANCO Then
            OptionProprioBanco.Value = True
        End If
    
        If gobjBorderoPagEmissao.iPodeDepCtaOutroBco = PERMITE_DEP_OUTRO_BANCO Then
            CheckDepositoOutroBanco.Value = vbChecked
        ElseIf gobjBorderoPagEmissao.iPodeDepCtaOutroBco = NAO_PERMITE_DEP_OUTRO_BANCO Then
            CheckDepositoOutroBanco.Value = vbUnchecked
        End If
    
    End If
    
    iBorderoAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        Case 7310 'o parametro é obrigatorio
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_PARCELAS_PAG_SEL", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143781)

    End Select

    Exit Function

End Function

Private Sub CheckDepositoOutroBanco_Click()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaCorrente_Change()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaCorrente_Click()

    iBorderoAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrente_Validate

    'Verifica se a Conta está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o ítem selecionado na ComboBox CodConta
    If ContaCorrente.Text = ContaCorrente.List(ContaCorrente.ListIndex) Then Exit Sub

    'Verifica se o a Conta existe na Combo, e , se existir, seleciona
    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 15735

    'Se a Conta(CODIGO) não existe na Combo
    If lErro = 6730 Then

        objContaCorrenteInt.iCodigo = iCodigo

        'Lê os dados da Conta
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 15736

        'Se a Conta não estiver cadastrada
        If lErro = 11807 Then Error 15737

        'Se a Conta não é Bancária
        If objContaCorrenteInt.iCodBanco = 0 Then Error 15738

        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            'Se a Conta não é da Filial selecionada
            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 15739

        End If

        'Passa o código da Conta para a tela
        ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Se a Conta(STRING) não existe na Combo
    If lErro = 6731 Then Error 15740

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 15735, 15736

        Case 15737
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If

        Case 15738
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", Err, ContaCorrente.Text)

        Case 15739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaCorrente.Text, giFilialEmpresa)

        Case 15740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, ContaCorrente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143782)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a DataEmissao está vazia
    If Len(DataEmissao.ClipText) > 0 Then

        'Verifica se a DataEmissao é válida
        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then Error 15759

    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case Err

        Case 15759

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143783)

    End Select

    Exit Sub

End Sub

Private Sub DataVencimento_Change()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataVencimento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataVencimento)

End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataVencimento_Validate

    'Verifica se a DataVencimento está vazia
    If Len(DataVencimento.ClipText) > 0 Then

        'Verifica se a DataVencimento é válida
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then Error 15760

    End If

    Exit Sub

Erro_DataVencimento_Validate:

    Cancel = True


    Select Case Err

        Case 15760

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143784)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoNomeCobranca As New AdmColCodigoNome
Dim colCodigoNomeConta As New AdmColCodigoNome
Dim objCodigoNomeCobranca As New AdmCodigoNome
Dim objCodigoNomeConta As New AdmCodigoNome
Dim iTipoCobranca As Integer

On Error GoTo Erro_Form_Load

    Set objEventoContaCorrente = New AdmEvento

    'Carrega a Coleção de Tipos de Cobrança
    lErro = CF("TiposDeCobranca_Bordero_Le_CodigoDescricao", colCodigoNomeCobranca)
    If lErro <> SUCESSO Then Error 15777
    
    TipoCobranca.AddItem CStr(TIPO_COBRANCA_TODAS) & SEPARADOR & STRING_TIPO_COBRANCA_TODAS
    TipoCobranca.ItemData(TipoCobranca.NewIndex) = TIPO_COBRANCA_TODAS
    
    'Preenche a ComboBox TipoCobranca com os objetos da coleção de Tipos de Cobrança
    For Each objCodigoNomeCobranca In colCodigoNomeCobranca

        TipoCobranca.AddItem CStr(objCodigoNomeCobranca.iCodigo) & SEPARADOR & objCodigoNomeCobranca.sNome
        TipoCobranca.ItemData(TipoCobranca.NewIndex) = objCodigoNomeCobranca.iCodigo

    Next
    
    'Seleciona Cobrança Bancária como Todas
    TipoCobranca.ListIndex = 0
    OptionProprioBanco.Enabled = True
    OptionOutroBanco.Enabled = True
    OptionAmbos.Enabled = True
    OptionAmbos.Value = True
            
    'Carrega a Coleção de Contas
    lErro = CF("ContasCorrentes_Bancarias_Le_CodigosNomesRed", colCodigoNomeConta)
    If lErro <> SUCESSO Then Error 15757

    'Preenche a ComboBox CodConta com os objetos da coleção de Contas
    For Each objCodigoNomeConta In colCodigoNomeConta

        ContaCorrente.AddItem CStr(objCodigoNomeConta.iCodigo) & SEPARADOR & objCodigoNomeConta.sNome
        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNomeConta.iCodigo

    Next

    'Seleciona uma das Contas
    If ContaCorrente.ListCount <> 0 Then ContaCorrente.Text = ContaCorrente.List(0)

    'Preenche as Datas com a data corrente do sistema
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataVencimento.Text = Format(gdtDataAtual, "dd/mm/yy")
    Call DateParaMasked(DataVencimentoDe, DATA_NULA)
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 15757, 15777

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143785)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoContaCorrente = Nothing
    
    Set gobjBorderoPagEmissao = Nothing
    
End Sub

Private Sub LabelContaCorrente_Click()
'Chamada do Browse de Contas

Dim colSelecao As Collection
Dim objConta As New ClassContasCorrentesInternas

    'Se a Conta não está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then

        'Inicializa a Conta no Obj
        objConta.iCodigo = 0

    'Se a Conta está preenchida
    Else

        'Passa o Código da Conta que está na tela para o Obj
        objConta.iCodigo = Codigo_Extrai(ContaCorrente.Text)

    End If

    'Chama a tela com a lista de Contas
    Call Chama_Tela("CtaCorrBancariaLista", colSelecao, objConta, objEventoContaCorrente)

    Exit Sub

End Sub

Private Sub objEventoContaCorrente_evSelecao(obj1 As Object)

Dim objConta As ClassContasCorrentesInternas

    Set objConta = obj1
    
    ContaCorrente.Text = CStr(objConta.iCodigo)
    Call ContaCorrente_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub Option1_Click()

End Sub

Private Sub OptionAmbos_Click()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OptionOutroBanco_Click()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OptionProprioBanco_Click()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoCobranca_Change()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoCobranca_Click()

Dim iTipoCobranca As Integer

    iBorderoAlterado = REGISTRO_ALTERADO

    'Verifica se TipoCobrança está preenchido
    If Len(Trim(TipoCobranca.Text)) = 0 Then Exit Sub

    'Extrai o Código do Tipo de Cobrança
    iTipoCobranca = Codigo_Extrai(TipoCobranca.Text)
    
    'Atualiza na tela a habilitacao dos campos de acordo com o Tipo de cobranca selecionado
    Call Trata_Troca_TipoCobranca(iTipoCobranca)
        
End Sub

Private Sub Trata_Troca_TipoCobranca(ByVal iTipoCobranca As Integer)
'Atualiza na tela a habilitacao dos campos de acordo com o Tipo de cobranca selecionado

'    'Se Tipo de Cobrança for Bancária
'    If iTipoCobranca = TIPO_COBRANCA_BANCARIA Then
'
'        'Habilita Liquidação de Títulos
'        OptionProprioBanco.Enabled = True
'        OptionOutroBanco.Enabled = True
'        OptionAmbos.Enabled = True
'        OptionAmbos.Value = True
'
'    'Se Tipo de Cobrança não for Bancária
'    Else
'
'        'Desabilita Liquidação de Títulos
'        OptionProprioBanco.Value = False
'        OptionOutroBanco.Value = False
'        OptionAmbos.Value = False
'        OptionProprioBanco.Enabled = False
'        OptionOutroBanco.Enabled = False
'        OptionAmbos.Enabled = False
'
'    End If

    'Se Tipo de Cobrança for Depósito em Conta
    If iTipoCobranca = TIPO_COBRANCA_DEP_CONTA Then
        
        'Habilita CheckDepositoOutroBanco
        CheckDepositoOutroBanco.Enabled = True
        CheckDepositoOutroBanco.Value = 1
        
    'Se Tipo de Cobrança não for Depósito em Conta
    Else

        'Desabilita CheckDepositoOutroBanco
        CheckDepositoOutroBanco.Value = 0
        CheckDepositoOutroBanco.Enabled = False

    End If

End Sub

Private Sub UpDownDataEmissao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_DownClick

    'Diminui a data em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15763

    Exit Sub

Erro_UpDownDataEmissao_DownClick:

    Select Case Err

        Case 15763

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143786)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_UpClick

    'Aumenta a data em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15764

    Exit Sub

Erro_UpDownDataEmissao_UpClick:

    Select Case Err

        Case 15764

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143787)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimento_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_DownClick

    'Diminui a data em 1 dia
    lErro = Data_Up_Down_Click(DataVencimento, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15765

    Exit Sub

Erro_UpDownDataVencimento_DownClick:

    Select Case Err

        Case 15765

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143788)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimento_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_UpClick

    'Aumenta a data em 1 dia
    lErro = Data_Up_Down_Click(DataVencimento, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15766

    Exit Sub

Erro_UpDownDataVencimento_UpClick:

    Select Case Err

        Case 15766

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143789)

    End Select

    Exit Sub

End Sub

Private Sub ValorMaxBordero_Change()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorMaxBordero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorMaxBordero_Validate

    'Se ValorMaxBordero está preenchido
    If Len(Trim(ValorMaxBordero.Text)) > 0 Then

        'Verifica se ValorMaxBordero é válido
        lErro = Valor_Positivo_Critica(ValorMaxBordero.Text)
        If lErro <> SUCESSO Then Error 15778

        'Formata o texto
        ValorMaxBordero.Text = Format(ValorMaxBordero.Text, "Fixed")

    End If

    Exit Sub

Erro_ValorMaxBordero_Validate:

    Cancel = True


    Select Case Err

        Case 15778

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143790)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BORDERO_PAGT_P1
    Set Form_Load_Ocx = Me
    Caption = "Bordero de Pagamento - Passo 1"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoPag1"
    
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
        If Me.ActiveControl Is ContaCorrente Then
            Call LabelContaCorrente_Click
        End If
    
    End If
    
End Sub



Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub LabelContaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaCorrente, Source, X, Y)
End Sub

Private Sub LabelContaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaCorrente, Button, Shift, X, Y)
End Sub

Private Sub DataVencimentoDe_Change()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataVencimentoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataVencimentoDe)

End Sub

Private Sub DataVencimentoDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataVencimentoDe_Validate

    'Verifica se a DataVencimento está vazia
    If Len(DataVencimentoDe.ClipText) > 0 Then

        'Verifica se a DataVencimento é válida
        lErro = Data_Critica(DataVencimentoDe.Text)
        If lErro <> SUCESSO Then Error 15760

    End If

    Exit Sub

Erro_DataVencimentoDe_Validate:

    Cancel = True


    Select Case Err

        Case 15760

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143791)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimentoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimentoDe_DownClick

    'Diminui a data em 1 dia
    lErro = Data_Up_Down_Click(DataVencimentoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15765

    Exit Sub

Erro_UpDownDataVencimentoDe_DownClick:

    Select Case Err

        Case 15765

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143792)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimentoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimentoDe_UpClick

    'Aumenta a data em 1 dia
    lErro = Data_Up_Down_Click(DataVencimentoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15766

    Exit Sub

Erro_UpDownDataVencimentoDe_UpClick:

    Select Case Err

        Case 15766

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143793)

    End Select

    Exit Sub

End Sub

Private Sub DataContabil_Change()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataContabil_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataContabil)

End Sub

Private Sub DataContabil_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataContabil_Validate

    'Verifica se a DataContabil está vazia
    If Len(DataContabil.ClipText) > 0 Then

        'Verifica se a DataContabil é válida
        lErro = Data_Critica(DataContabil.Text)
        If lErro <> SUCESSO Then Error 15759

    End If

    Exit Sub

Erro_DataContabil_Validate:

    Cancel = True


    Select Case Err

        Case 15759

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143783)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataContabil_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataContabil_DownClick

    'Diminui a data em 1 dia
    lErro = Data_Up_Down_Click(DataContabil, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15763

    Exit Sub

Erro_UpDownDataContabil_DownClick:

    Select Case Err

        Case 15763

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143786)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataContabil_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataContabil_UpClick

    'Aumenta a data em 1 dia
    lErro = Data_Up_Down_Click(DataContabil, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15764

    Exit Sub

Erro_UpDownDataContabil_UpClick:

    Select Case Err

        Case 15764

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143787)

    End Select

    Exit Sub

End Sub
