VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ChequesPagOcx 
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   KeyPreview      =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   7755
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   3037
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3045
      Width           =   1680
      Begin VB.CommandButton BotaoSeguir 
         Height          =   345
         Left            =   90
         Picture         =   "ChequesPagOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   90
         Width           =   930
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1125
         Picture         =   "ChequesPagOcx.ctx":0792
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecionar os Títulos que Atendem às Seguintes Condições"
      Height          =   1410
      Left            =   120
      TabIndex        =   11
      Top             =   1590
      Width           =   7485
      Begin VB.Frame Frame4 
         Caption         =   "Vencimento"
         Height          =   1050
         Left            =   180
         TabIndex        =   19
         Top             =   225
         Width           =   3315
         Begin MSComCtl2.UpDown UpDownDataVencimento 
            Height          =   300
            Left            =   2550
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   645
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   300
            Left            =   1485
            TabIndex        =   21
            Top             =   645
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataVencimentoDe 
            Height          =   300
            Left            =   2580
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataVencimentoDe 
            Height          =   300
            Left            =   1485
            TabIndex        =   23
            Top             =   225
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
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
            Left            =   1095
            TabIndex        =   25
            Top             =   690
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
            Height          =   195
            Left            =   1140
            TabIndex        =   24
            Top             =   240
            Width           =   315
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Portador"
         Height          =   1050
         Left            =   3660
         TabIndex        =   12
         Top             =   225
         Width           =   3615
         Begin VB.ComboBox Portador 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   585
            Width           =   2130
         End
         Begin VB.OptionButton OptionQualquer 
            Caption         =   "Qualquer"
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
            Left            =   240
            TabIndex        =   4
            Top             =   285
            Width           =   1185
         End
         Begin VB.OptionButton OptionApenas 
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
            Height          =   225
            Left            =   240
            TabIndex        =   5
            Top             =   660
            Width           =   1005
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados Principais"
      Height          =   1350
      Left            =   150
      TabIndex        =   10
      Top             =   150
      Width           =   7455
      Begin VB.ComboBox ContaCorrente 
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Top             =   390
         Width           =   1965
      End
      Begin MSComCtl2.UpDown UpDownDataEmissao 
         Height          =   300
         Left            =   2790
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   855
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   1650
         TabIndex        =   2
         Top             =   855
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
         Left            =   6390
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   870
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataContabil 
         Height          =   300
         Left            =   5265
         TabIndex        =   3
         Top             =   870
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProxCheque 
         Height          =   300
         Left            =   5250
         TabIndex        =   1
         Top             =   405
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Próx. Cheque:"
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
         Left            =   3990
         TabIndex        =   15
         Top             =   435
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3975
         TabIndex        =   16
         Top             =   915
         Width           =   1230
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   810
         TabIndex        =   17
         Top             =   900
         Width           =   765
      End
      Begin VB.Label LabelContaCorrente 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   1005
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   435
         Width           =   570
      End
   End
End
Attribute VB_Name = "ChequesPagOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private gobjChequesPag As ClassChequesPag
Private WithEvents objEventoContaCorrente As AdmEvento
Attribute objEventoContaCorrente.VB_VarHelpID = -1

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoSeguir_Click()

Dim lErro As Long, dtDataDe As Date
Dim iCodConta As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_BotaoSeguir_Click

    'Verifica se a ContaCorrente está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then Error 15641

    'Verifica se o ProxCheque está preenchido
    If Len(Trim(ProxCheque.Text)) = 0 Then Error 15665
    
    'Verifica se a DataEmissao está preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Error 15644

    'Verifica se a DataVencimento está preenchida
    If Len(Trim(DataVencimento.ClipText)) = 0 Then Error 15646

    dtDataDe = MaskedParaDate(DataVencimentoDe)
    
    If dtDataDe <> DATA_NULA And dtDataDe > MaskedParaDate(DataVencimento) Then Error 32301
    
    'Extrai o Código da Conta que está na tela
    iCodConta = Codigo_Extrai(ContaCorrente.Text)

    'Passa o Código da Conta para o Obj
    objContaCorrenteInt.iCodigo = iCodConta

    'Lê os dados da Conta
    lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 15654

    'Se a Conta não estiver cadastrada
    If lErro = 11807 Then Error 15655

    'Se a Conta não é Bancária
    If objContaCorrenteInt.iCodBanco = 0 Then Error 15656
    
    'Se alguma Filial tiver sido selecionada
    If giFilialEmpresa <> EMPRESA_TODA Then
        
        'Se a Conta não é da Filial selecionada
        If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 15733
        
    End If

    'Se o Portador não for Qualquer
    If OptionQualquer.Value = False Then

        'Verifica se o Portador está preenchido
        If Len(Trim(Portador.Text)) = 0 Then Error 15642

    End If

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        'Verifica se a DataContabil está preenchida
        If Len(Trim(DataContabil.ClipText)) = 0 Then Error 15645
    
        'Verifica se a DataContabil é maior ou igual que a DataEmissao
        If CDate(DataContabil.Text) < CDate(DataEmissao.Text) Then Error 15643
    
        gobjChequesPag.dtContabil = CDate(DataContabil.Text)
    
    Else
    
        gobjChequesPag.dtContabil = gdtDataAtual
        
    End If
    
    gobjChequesPag.dtEmissao = CDate(DataEmissao.Text)
    gobjChequesPag.dtVencto = CDate(DataVencimento.Text)
    gobjChequesPag.dtVenctoDe = dtDataDe
    gobjChequesPag.iCta = iCodConta
    
    If OptionApenas.Value = True Then
        gobjChequesPag.iQualquerPortador = 0
        gobjChequesPag.iPortador = Codigo_Extrai(Portador.Text)
    Else
        gobjChequesPag.iQualquerPortador = QUALQUER_PORTADOR
    End If
    
    gobjChequesPag.lNumCheque = CLng(ProxCheque.Text)
     
    'preencher colecao de parcelas a pagar
    lErro = CF("ParcelasPagar_Le_ChequesPag", gobjChequesPag)
    If lErro <> SUCESSO Then Error 15861

    'Chama a tela do passo seguinte
    Call Chama_Tela("ChequesPag2", gobjChequesPag)

    'Fecha a tela
    Unload Me
    
    Exit Sub

Erro_BotaoSeguir_Click:

    Select Case Err

        Case 32301
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR_FINAL", Err)
        
        Case 15641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)

        Case 15642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_INFORMADO", Err)

        Case 15643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATACONTABIL_MENOR_DATAEMISSAO", Err)

        Case 15644, 15645, 15646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 15654, 15861
        
        Case 15655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, iCodConta)
            
        Case 15656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", Err, ContaCorrente.Text)
            
        Case 15665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROXCHEQUE_NAO_INFORMADO", Err)
        
        Case 15733
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaCorrente.Text, giFilialEmpresa)
            ContaCorrente.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144579)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objChequesPag As ClassChequesPag) As Long

Dim lErro As Long
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_Trata_Parametros

    Set gobjChequesPag = objChequesPag

    'se a cta já estiver preenchida significa que o usuario esta voltando do passo 2
    If (gobjChequesPag.iCta <> 0) Then
        
        ContaCorrente.Text = CStr(gobjChequesPag.iCta)
        Call ContaCorrente_Validate(bSGECancelDummy)
        ProxCheque.Text = CStr(gobjChequesPag.lNumCheque)
        DataEmissao.Text = Format(gobjChequesPag.dtEmissao, "dd/mm/yy")
        DataContabil.Text = Format(gobjChequesPag.dtContabil, "dd/mm/yy")
        DataVencimento.Text = Format(gobjChequesPag.dtVencto, "dd/mm/yy")
        Call DateParaMasked(DataVencimentoDe, gobjChequesPag.dtVenctoDe)
        
        'definindo o portador
        If gobjChequesPag.iQualquerPortador = QUALQUER_PORTADOR Then
            OptionQualquer.Value = True
            OptionApenas.Value = False
        Else
            OptionQualquer.Value = False
            OptionApenas.Value = True
            Portador.Text = CStr(gobjChequesPag.iPortador)
            Call Combo_Item_Seleciona(Portador)
        End If
        
        'Percorre todas as parelas da Coleção passada por parâmetro
        Do While gobjChequesPag.colInfoParcPag.Count <> 0
            
            gobjChequesPag.colInfoParcPag.Remove (1)
            
        Loop
   
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144580)

    End Select

    Exit Function

End Function

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

    'Verifica se a Conta existe na Combo, e , se existir, seleciona
    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 15637

    'Se a Conta(CODIGO) não existe na Combo
    If lErro = 6730 Then

        objContaCorrenteInt.iCodigo = iCodigo

        'Lê os dados da Conta
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 15638

        'Se a Conta não estiver cadastrada
        If lErro = 11807 Then Error 15639

        'Se a Conta não é Bancária
        If objContaCorrenteInt.iCodBanco = 0 Then Error 15664
        
        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then
        
            'Se a Conta não é da Filial selecionada
            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 15734
            
        End If
        
        'Passa o código da Conta para a tela
        ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Se a Conta(STRING) não existe na Combo
    If lErro = 6731 Then Error 15640

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 15637, 15638

        Case 15639
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If

        Case 15640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, ContaCorrente.Text)

        Case 15664
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", Err, ContaCorrente.Text)
            
        Case 15734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaCorrente.Text, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144581)

    End Select

    Exit Sub

End Sub

Private Sub DataContabil_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataContabil)

End Sub

Private Sub DataContabil_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataContabil_Validate

    'Se a DataContabil está preenchida
    If Len(DataContabil.ClipText) > 0 Then

        'Verifica se a DataContabil é válida
        lErro = Data_Critica(DataContabil.Text)
        If lErro <> SUCESSO Then Error 15647

    End If

    Exit Sub

Erro_DataContabil_Validate:

    Cancel = True


    Select Case Err

        Case 15647

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144582)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEmissao)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Se a DataEmissao está preenchida
    If Len(DataEmissao.ClipText) > 0 Then

        'Verifica se a DataEmissao é válida
        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then Error 15648

    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case Err

        Case 15648

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144583)

    End Select

    Exit Sub

End Sub

Private Sub DataVencimento_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataVencimento)
    
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataVencimento_Validate

    'Se a DataVencimento está preenchida
    If Len(DataVencimento.ClipText) > 0 Then

        'Verifica se a DataVencimento é válida
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then Error 15649

    End If

    Exit Sub

Erro_DataVencimento_Validate:

    Cancel = True


    Select Case Err

        Case 15649

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144584)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoNomeConta As New AdmColCodigoNome
Dim objCodigoNomeConta As New AdmCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoContaCorrente = New AdmEvento
    
    'Carrega a Coleção de Contas
    lErro = CF("ContasCorrentes_Bancarias_Le_CodigosNomesRed", colCodigoNomeConta)
    If lErro <> SUCESSO Then Error 15633

    'Preenche a ComboBox CodConta com os objetos da coleção de Contas
    For Each objCodigoNomeConta In colCodigoNomeConta

        ContaCorrente.AddItem CStr(objCodigoNomeConta.iCodigo) & SEPARADOR & objCodigoNomeConta.sNome
        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNomeConta.iCodigo

    Next

    'Seleciona uma das Contas
    If ContaCorrente.ListCount <> 0 Then ContaCorrente.ListIndex = 0

    'Carrega a ComboBox Portador
    lErro = Portador_Carrega(Portador)
    If lErro <> SUCESSO Then Error 15634

    'Preenche as Datas com a data corrente do sistema
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataVencimento.Text = Format(gdtDataAtual, "dd/mm/yy")
    Call DateParaMasked(DataVencimentoDe, DATA_NULA)
    
    'Mostra qualquer Portador como default
    OptionQualquer.Value = True
    OptionApenas.Value = False
    Portador.Enabled = False

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) <> MODULO_ATIVO Then
        Label10.Enabled = False
        DataContabil.Enabled = False
        UpDownDataContabil.Enabled = False
    Else
        DataContabil.Text = Format(gdtDataAtual, "dd/mm/yy")
    End If

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 15633, 15634

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144585)

    End Select

    Exit Sub

End Sub

Private Function Portador_Carrega(objComboBox As ComboBox) As Long
'Carrega a ComboBox de Portadores

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNomeRed As New AdmCodigoNome

On Error GoTo Erro_Portador_Carrega

    'Lê Codigos e NomesReduzidos de Portadores
    lErro = CF("Portadores_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then Error 15632

    'Preenche a ComboBox Portadores
    For Each objCodigoNomeRed In colCodigoNomeRed
        objComboBox.AddItem CStr(objCodigoNomeRed.iCodigo) & SEPARADOR & objCodigoNomeRed.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoNomeRed.iCodigo
    Next

    Portador_Carrega = SUCESSO

    Exit Function

Erro_Portador_Carrega:

    Portador_Carrega = Err

    Select Case Err

        Case 15632

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144586)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoContaCorrente = Nothing
    
    Set gobjChequesPag = Nothing
    
End Sub

Private Sub LabelContaCorrente_Click()
'Chamada do Browse de Contas

Dim colSelecao As Collection
Dim objConta As New ClassContasCorrentesInternas
Dim objChequesPag As New ClassChequesPag

    'Se a ContaCorrente não está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then

        'Inicializa a Conta no Obj
        objConta.iCodigo = 0

    'Se a ContaCorrente está preenchida
    Else

        'Passa o Código da Conta para o Obj
        objConta.iCodigo = Codigo_Extrai(ContaCorrente.Text)

    End If
    
    'Passa o Código da Filial selecionada para o Obj
    objConta.iFilialEmpresa = giFilialEmpresa

    'Se Empresa Toda está selecionada
    If giFilialEmpresa = EMPRESA_TODA Then
    
        'Chama a tela com a lista de Contas de toda a empresa
        Call Chama_Tela("CtaCorrBancariaTodasLista", colSelecao, objConta, objEventoContaCorrente)
        
    'Se alguma Filial está selecionada
    Else
    
        'Chama a tela com a lista de Contas da Filial selecionada
        Call Chama_Tela("CtaCorrBancariaLista", colSelecao, objConta, objEventoContaCorrente)
            
    End If
        
    Exit Sub

End Sub

Private Sub OptionApenas_Click()

    'Se a Opção Apenas um portador estiver selecionada, habilita a Combo de Portador
    If OptionApenas.Value = True Then Portador.Enabled = True

End Sub

Private Sub OptionQualquer_Click()
    
    'Se a Opção Qualquer portador estiver selecionada
    If OptionQualquer.Value = True Then
        
        'Desabilita a Combo de Portador
        Portador.ListIndex = -1
        Portador.Enabled = False
    
    End If
  
End Sub

Private Sub objEventoContaCorrente_evSelecao(obj1 As Object)

Dim objConta As ClassContasCorrentesInternas

    'Retorna os dados do Obj para a tela
    Set objConta = obj1
    ContaCorrente.Text = CStr(objConta.iCodigo)
    ContaCorrente.SetFocus
    
    'Exibe a tela
    Me.Show

    Exit Sub

End Sub

Private Sub ProxCheque_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProxCheque)

End Sub

Private Sub UpDownDataContabil_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataContabil_DownClick
    
    'Diminui a DataContabil em 1 dia
    lErro = Data_Up_Down_Click(DataContabil, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15658
    
    Exit Sub

Erro_UpDownDataContabil_DownClick:

    Select Case Err

        Case 15658

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144587)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataContabil_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataContabil_UpClick

    'Aumenta a DataContabil em 1 dia
    lErro = Data_Up_Down_Click(DataContabil, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15659
    
    Exit Sub
    
Erro_UpDownDataContabil_UpClick:

    Select Case Err

        Case 15659

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144588)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_DownClick

    'Diminui a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15661
    
    Exit Sub
    
Erro_UpDownDataEmissao_DownClick:

    Select Case Err

        Case 15661

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144589)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_UpClick

    'Aumenta a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15660
    
    Exit Sub
    
Erro_UpDownDataEmissao_UpClick:

    Select Case Err

        Case 15660

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144590)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimento_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_DownClick

    'Diminui a DataVencimento em 1 dia
    lErro = Data_Up_Down_Click(DataVencimento, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15662
    
    Exit Sub

Erro_UpDownDataVencimento_DownClick:

    Select Case Err

        Case 15662

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144591)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimento_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_UpClick

    'Aumenta a DataVencimento em 1 dia
    lErro = Data_Up_Down_Click(DataVencimento, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15663
    
    Exit Sub

Erro_UpDownDataVencimento_UpClick:

    Select Case Err

        Case 15663

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144592)

    End Select

    Exit Sub

End Sub

Private Sub Portador_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objPortador As New ClassPortador

On Error GoTo Erro_Portador_Validate

    'Verifica se o Portador foi preenchido
    If Len(Trim(Portador.Text)) = 0 Then Exit Sub
    
    'Verifica se é um Portador selecionado
    If Portador.Text = Portador.List(Portador.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Portador, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 48899
    
    'Se não encontra valor que contém CÓDIGO e não encontrou o valor que era STRING
    If lErro = 6730 Or lErro = 6731 Then Error 48900
        
    Exit Sub
    
Erro_Portador_Validate:

    Cancel = True


    Select Case Err
    
        Case 48899
        
        Case 48900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_ENCONTRADO", Err, Portador.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144593)
    
    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CHEQUES_PAGAR_P1
    Set Form_Load_Ocx = Me
    Caption = "Impressão de Cheques - Passo 1"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ChequesPag"
    
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



Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelContaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaCorrente, Source, X, Y)
End Sub

Private Sub LabelContaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaCorrente, Button, Shift, X, Y)
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
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144594)

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
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144595)

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
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144596)

    End Select

    Exit Sub

End Sub


