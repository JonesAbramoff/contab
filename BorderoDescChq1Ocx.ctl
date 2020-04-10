VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BorderoDescChq1Ocx 
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
   ScaleHeight     =   3195
   ScaleWidth      =   7170
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   5385
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   105
      Width           =   1680
      Begin VB.CommandButton BotaoSeguir 
         Height          =   330
         Left            =   90
         Picture         =   "BorderoDescChq1Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   90
         Width           =   930
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1125
         Picture         =   "BorderoDescChq1Ocx.ctx":0792
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecionar os Cheques que Atendem às Seguintes Condições"
      Height          =   900
      Left            =   60
      TabIndex        =   18
      Top             =   2205
      Width           =   7020
      Begin MSComCtl2.UpDown UpDownDataBomParaAte 
         Height          =   300
         Left            =   2565
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataBomParaAte 
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Agencia 
         Height          =   300
         Left            =   5850
         TabIndex        =   8
         Top             =   360
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Banco 
         Height          =   300
         Left            =   3795
         TabIndex        =   7
         Top             =   360
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agência:"
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
         Height          =   195
         Index           =   6
         Left            =   5025
         TabIndex        =   22
         Top             =   405
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
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
         Height          =   195
         Index           =   7
         Left            =   3105
         TabIndex        =   21
         Top             =   405
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bom Para Até:"
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
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   20
         Top             =   405
         Width           =   1230
      End
   End
   Begin VB.ComboBox ContaCorrente 
      Height          =   315
      Left            =   5355
      TabIndex        =   4
      Top             =   765
      Width           =   1695
   End
   Begin VB.ComboBox CarteiraCobranca 
      Height          =   315
      Left            =   1710
      TabIndex        =   1
      Top             =   765
      Width           =   1920
   End
   Begin VB.ComboBox Cobrador 
      Height          =   315
      Left            =   1710
      TabIndex        =   0
      Top             =   300
      Width           =   1920
   End
   Begin MSComCtl2.UpDown UpDownDataEmissao 
      Height          =   300
      Left            =   2865
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1215
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataEmissao 
      Height          =   300
      Left            =   1710
      TabIndex        =   2
      Top             =   1230
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
      Left            =   6510
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1215
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataContabil 
      Height          =   300
      Left            =   5355
      TabIndex        =   5
      Top             =   1215
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDataCredito 
      Height          =   300
      Left            =   2880
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1665
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataCredito 
      Height          =   300
      Left            =   1710
      TabIndex        =   3
      Top             =   1665
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data p/Crédito:"
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
      Index           =   8
      Left            =   330
      TabIndex        =   25
      Top             =   1695
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conta Corrente:"
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
      Index           =   1
      Left            =   3960
      TabIndex        =   17
      Top             =   810
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Carteira:"
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
      Index           =   0
      Left            =   930
      TabIndex        =   16
      Top             =   810
      Width           =   735
   End
   Begin VB.Label Label1 
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
      Index           =   2
      Left            =   4080
      TabIndex        =   15
      Top             =   1245
      Width           =   1230
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   900
      TabIndex        =   14
      Top             =   1245
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cobrador:"
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
      Index           =   3
      Left            =   825
      TabIndex        =   13
      Top             =   345
      Width           =   840
   End
End
Attribute VB_Name = "BorderoDescChq1Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Dim gobjBorderoDescChq As ClassBorderoDescChq
Public iAlterado As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'carrega a combo de cobradores
    lErro = Carrega_Cobradores()
    If lErro <> SUCESSO Then gError 109135
    
    'carrega a combo de contas correntes
    lErro = Carrega_ContasCorrentes()
    If lErro <> SUCESSO Then gError 109136
    
    'preecher a data emissão com a data atual
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
'    'preecher a data crédito com a data atual
'    DataCredito.PromptInclude = False
'    DataCredito.Text = Format(gdtDataAtual, "dd/mm/yy")
'    DataCredito.PromptInclude = True
    
    'preencher a data bom para até com a data atual
    DataBomParaAte.PromptInclude = False
    DataBomParaAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataBomParaAte.PromptInclude = True
    
    'se o módulo de contabilidade estiver ativo, preenche a data de contabilidade com a atual
    If gcolModulo.ATIVO(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        
        DataContabil.PromptInclude = False
        DataContabil.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataContabil.PromptInclude = True
        
    'caso contrário, esconde os campos
    Else
    
        DataContabil.Visible = False
        UpDownDataContabil.Visible = False
        
    End If
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 109135
        
        Case 109136
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143704)
            
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objBorderoDescChq As ClassBorderoDescChq) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'se o borderô está instanciado, o traz para a tela
    If Not objBorderoDescChq Is Nothing Then
    
        'faz o global dessa tela apontar para o recebido por parâmetro
        Set gobjBorderoDescChq = objBorderoDescChq
        
        lErro = Traz_BorderoDescChq_Tela(objBorderoDescChq)
        If lErro <> SUCESSO Then gError 109140
    
    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 109140
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143705)
            
    End Select
    
    Exit Function

End Function

Private Sub Agencia_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Banco_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CarteiraCobranca_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cobrador_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cobrador_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCobrador As New ClassCobrador

On Error GoTo Erro_Cobrador_Validate

    'limpa a combo de carteiracobranca
    CarteiraCobranca.Clear
    
    'limpa o conteúdo do Text da combo de conta corrente
    ContaCorrente.Text = ""
    
    'se cobrador estiver preenchido
    If Len(Trim(Cobrador.Text)) <> 0 Then
    
        'se não foi selecionado por clique na combo
        If Cobrador.ListIndex = -1 Then
            
            'seleciona o cobrador na combo
            lErro = Combo_Seleciona(Cobrador, iCodigo)
            
            objCobrador.iCodigo = iCodigo
            
            'se não achou pelo código
            If lErro = 6730 Then
                
                'tenta ler no BD.
                lErro = CF("Cobrador_Le", objCobrador)
                If lErro <> SUCESSO And lErro <> 19294 Then gError 109139
                
                'se não encontrou-> erro
                If lErro = 19294 Then gError 109140
            
            'não havia código, tenta buscar pela string
            Else
            
                'se não encontrou pela string-> erro
                If lErro = 6731 Then gError 109143
                
                'tenta ler no BD.
                lErro = CF("Cobrador_Le", objCobrador)
                If lErro <> SUCESSO And lErro <> 19294 Then gError 109144
                
                'se não encontrou-> erro
                If lErro = 19294 Then gError 109145
                
            End If
            
            'se o cobrador lido foi o de código 1-> erro
            If objCobrador.iCodigo = COBRADOR_PROPRIA_EMPRESA Then gError 109188
            
            'carrega a combo de carteiras de acordo com o cobrador selecionado.
            lErro = Carrega_CarteiraCobranca(objCobrador)
            If lErro <> SUCESSO And lErro <> 109142 Then gError 109138
            
            'se carregou a combo com sucesso, escolhe o primeiro item como default
            If lErro = SUCESSO Then CarteiraCobranca.ListIndex = 0
            
        'se foi selecionado com click
        Else
            
            'preenche o código do cobrador
            objCobrador.iCodigo = Codigo_Extrai(Cobrador.Text)
            
            'tenta ler no BD.
            lErro = CF("Cobrador_Le", objCobrador)
            If lErro <> SUCESSO And lErro <> 19294 Then gError 109187
                
            'se não encontrou-> erro
            If lErro = 19294 Then gError 109185
            
            'carrega a combo de carteiras de acordo com o cobrador selecionado.
            lErro = Carrega_CarteiraCobranca(objCobrador)
            If lErro <> SUCESSO And lErro <> 109142 Then gError 109186
        
            'se carregou a combo com sucesso, escolhe o primeiro item como default
            If lErro = SUCESSO Then CarteiraCobranca.ListIndex = 0
        
        End If
            
        'se a conta corrente não estiver preenchida
        If Len(Trim(ContaCorrente.Text)) = 0 Then
        
            'seleciona a conta associada ao cobrador
            ContaCorrente.Text = objCobrador.iCodCCI
            Call ContaCorrente_Validate(bSGECancelDummy)
        
        End If

    End If

    Exit Sub
    
Erro_Cobrador_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 109138, 109139, 109187, 109186
        
        Case 109140, 109145, 109185
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", gErr, objCobrador.iCodigo)
            
        Case 109143
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO1", gErr, Cobrador.Text)
            
        Case 109188
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_PROPRIA_EMPRESA_NAO_PERMITIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143706)
    
    End Select
    
    Exit Sub

End Sub

Private Sub CarteiraCobranca_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objCarteiraCobranca As New ClassCarteiraCobranca

On Error GoTo Erro_CarteiraCobranca_Validate

    'se carteiracobrança estiver preenchida
    If Len(Trim(CarteiraCobranca.Text)) <> 0 Then
    
        'se ela foi digitada e não selecionada por clique na combo
        If CarteiraCobranca.ListIndex = -1 Then
        
            'tenta selecionar na combo
            lErro = Combo_Seleciona(CarteiraCobranca, iCodigo)
            
            'se não encontrar pelo código
            If lErro = 6730 Then
            
                'se o cobrador não estiver preenchido-> erro
                If Len(Trim(Cobrador.Text)) = 0 Then gError 109152
                
                'preenche o cobrador e a carteira cobranca do objcarteiracobrador
                objCarteiraCobrador.iCobrador = Codigo_Extrai(Cobrador.Text)
                objCarteiraCobrador.iCodCarteiraCobranca = iCodigo
                
                '... para verificar se eles possuem associação na tabela intermediária
                lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
                If lErro <> SUCESSO And lErro <> 23551 Then gError 109153
                                
                'se não encontrou-> erro
                If lErro = 23551 Then gError 109154
                
                'preenche o código da carteira de cobrança...
                objCarteiraCobranca.iCodigo = iCodigo
                
                '... para buscar a carteira de cobrança
                lErro = CF("CarteiraCobranca_Le", objCarteiraCobranca)
                If lErro <> SUCESSO And lErro <> 23413 Then gError 109155
                
                'se não encontrou-> erro
                If lErro = 23413 Then gError 109156
                
                'preenche a combo
                CarteiraCobranca.Text = objCarteiraCobranca.iCodigo & SEPARADOR & objCarteiraCobranca.sDescricao
                
            Else
            
                'se não encontrou pela string-> erro
                If lErro = 6731 Then gError 109157
                
            End If
        
        End If
    
    End If
    
    Exit Sub

Erro_CarteiraCobranca_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 109152
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)
            
        Case 109153, 109155
        
        Case 109154
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRASCOBRADOR_NAO_CADASTRADO", gErr, objCarteiraCobrador.iCobrador, objCarteiraCobrador.iCodCarteiraCobranca)
            
        Case 109156
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADO", gErr, objCarteiraCobranca.iCodigo)
        
        Case 109157
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADO", gErr, CarteiraCobranca.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143707)
    
    End Select
    

End Sub

Private Sub ContaCorrente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaCorrente_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrente_Validate

    'Verifica se a Conta está preenchida
    If Len(Trim(ContaCorrente.Text)) <> 0 Then
    
        'Verifica se esta preenchida com o ítem selecionado na ComboBox CodConta
        If ContaCorrente.ListIndex <> -1 Then Exit Sub
    
        'Verifica se o a Conta existe na Combo, e , se existir, seleciona
        lErro = Combo_Seleciona(ContaCorrente, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 109146
    
        'Se a Conta(CODIGO) não existe na Combo
        If lErro = 6730 Then
    
            objContaCorrenteInt.iCodigo = iCodigo
    
            'Lê os dados da Conta
            lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
            If lErro <> SUCESSO And lErro <> 11807 Then gError 109147
    
            'Se a Conta não estiver cadastrada
            If lErro = 11807 Then gError 109148
    
            'Se a Conta não é Bancária
            If objContaCorrenteInt.iCodBanco = 0 Then gError 109149
    
            'Se alguma Filial tiver sido selecionada
            If giFilialEmpresa <> EMPRESA_TODA Then
    
                'Se a Conta não é da Filial selecionada
                If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then gError 109150
    
            End If
    
            'Passa o código da Conta para a tela
            ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido
    
        End If
    
        'Se a Conta(STRING) não existe na Combo
        If lErro = 6731 Then gError 109151
        
    End If
    
    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True
    
    Select Case gErr

        Case 109146, 109147

        Case 109148
            
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)
            
            If vbMsgRes = vbYes Then Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)

        Case 109149
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA1", gErr, ContaCorrente.Text)

        Case 109150
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", gErr, ContaCorrente.Text, giFilialEmpresa)

        Case 109151
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, ContaCorrente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143708)

    End Select

    Exit Sub

End Sub

Private Sub DataBomParaAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataContabil_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataCredito_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataEmissao_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'se a data estiver preenchida
    If Len(Trim(DataEmissao.ClipText)) <> 0 Then

        'critica a data
        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then gError 109158
    
    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 109158
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143709)
    
    End Select
    
    Exit Sub

End Sub

Private Sub UpDownDataEmissao_DownClick()
    Call Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
End Sub

Private Sub UpDownDataEmissao_UpClick()
    Call Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
End Sub

Private Sub DataContabil_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataContabil, iAlterado)
End Sub

Private Sub DataContabil_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataContabil_Validate

    'se a data estiver preenchida
    If Len(Trim(DataContabil.ClipText)) <> 0 Then

        'critica a data
        lErro = Data_Critica(DataContabil.Text)
        If lErro <> SUCESSO Then gError 109159
    
    End If

    Exit Sub

Erro_DataContabil_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 109159
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143710)
    
    End Select
    
    Exit Sub

End Sub

Private Sub UpDownDataContabil_DownClick()
    Call Data_Up_Down_Click(DataContabil, DIMINUI_DATA)
End Sub

Private Sub UpDownDataContabil_UpClick()
    Call Data_Up_Down_Click(DataContabil, AUMENTA_DATA)
End Sub

Private Sub DataCredito_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataCredito, iAlterado)
End Sub

Private Sub DataCredito_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCredito_Validate

    'se a data estiver preenchida
    If Len(Trim(DataCredito.ClipText)) <> 0 Then

        'critica a data
        lErro = Data_Critica(DataCredito.Text)
        If lErro <> SUCESSO Then gError 109160
    
    End If

    Exit Sub

Erro_DataCredito_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 109160
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143711)
    
    End Select
    
    Exit Sub

End Sub

Private Sub UpDownDataCredito_DownClick()
    Call Data_Up_Down_Click(DataCredito, DIMINUI_DATA)
End Sub

Private Sub UpDownDataCredito_UpClick()
    Call Data_Up_Down_Click(DataCredito, AUMENTA_DATA)
End Sub

Private Sub DataBomParaAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataBomParaAte, iAlterado)
End Sub

Private Sub DataBomParaAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataBomParaAte_Validate

    'se a data estiver preenchida
    If Len(Trim(DataBomParaAte.ClipText)) <> 0 Then

        'critica a data
        lErro = Data_Critica(DataBomParaAte.Text)
        If lErro <> SUCESSO Then gError 109161
    
    End If

    Exit Sub

Erro_DataBomParaAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 109161
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143712)
    
    End Select
    
    Exit Sub

End Sub

Private Sub UpDownDataBomParaAte_DownClick()
    Call Data_Up_Down_Click(DataBomParaAte, DIMINUI_DATA)
End Sub

Private Sub UpDownDataBomParaAte_UpClick()
    Call Data_Up_Down_Click(DataBomParaAte, AUMENTA_DATA)
End Sub

Private Sub Banco_GotFocus()
    Call MaskEdBox_TrataGotFocus(Banco, iAlterado)
End Sub

Private Sub Banco_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Banco_Validate

    'se o banco estiver preenchido
    If Len(Trim(Banco.Text)) <> 0 Then
        
        'critica
        lErro = Inteiro_Critica(Banco.Text)
        If lErro <> SUCESSO Then gError 109162
        
    End If
    
    Exit Sub

Erro_Banco_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 109162
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143713)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Agencia_GotFocus()
    Call MaskEdBox_TrataGotFocus(Agencia, iAlterado)
End Sub

Private Sub BotaoSeguir_Click()

Dim lErro As Long
Dim dtDataBomParaAte As Date
Dim iBanco As Integer
Dim sAgencia As String

On Error GoTo Erro_BotaoSeguir_Click

    'se a agência estiver preenchida, mas o banco não estiver-> erro
    If Len(Trim(Agencia.Text)) <> 0 And Len(Trim(Banco.Text)) = 0 Then gError 109164
    
    'se o cobrador não estiver preenchido->erro
    If Len(Trim(Cobrador.Text)) = 0 Then gError 109165
    
    'se a carteira não estiver preenchida-> erro
    If Len(Trim(CarteiraCobranca.Text)) = 0 Then gError 109166
    
    'se a CCI não estiver prrenchida-> erro
    If Len(Trim(ContaCorrente.Text)) = 0 Then gError 109167
    
    'se a data emissão nao estiver preenchida-> erro
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 109168
    
    'se o móduo de contabilidade está ativo
    If gcolModulo.ATIVO(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        
        'se a data de contabilidade não estiver preenchida-> erro
        If Len(Trim(DataContabil.ClipText)) = 0 Then gError 109169
        
        'se a data de contabilidade for menor que a de emissão-> erro
        If StrParaDate(DataContabil.Text) < StrParaDate(DataEmissao.Text) Then gError 109170
    
    End If
    
    'se a data de crédito não estiver preenchida-> erro
    If Len(Trim(DataCredito.ClipText)) = 0 Then gError 109171
    
    'Verifica se houve alterações dos dados passados para o filtro do cheque
    If iAlterado = REGISTRO_ALTERADO Then
        
        Set gobjBorderoDescChq = New ClassBorderoDescChq
        
        Call Move_Tela_Memoria(gobjBorderoDescChq)
            
        'lê os cheques que podem participar desse borderô
        lErro = CF("BorderoDescChq_Le_ChequesPre_Disp", gobjBorderoDescChq)
        If lErro <> SUCESSO And lErro <> 109228 Then gError 109183
        
        'se não encontrou nenhum cheque->erro
        If lErro = 109228 Then gError 109184
    
    End If
    
    Call Chama_Tela("BorderoDescChq2", gobjBorderoDescChq)
        
    Unload Me

Exit Sub

Erro_BotaoSeguir_Click:

    Select Case gErr
    
        Case 109164
            Call Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_INFORMADO", gErr)
            
        Case 109165
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)
            
        Case 109166
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBANCA_NAO_INFORMADA", gErr)
            
        Case 109167
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)
        
        Case 109168
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_INFORMADA", gErr)
            
        Case 109169
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_CONTABIL_NAO_PREENCHIDA", gErr)
            
        Case 109170
            Call Rotina_Erro(vbOKOnly, "ERRO_DATACONTABIL_MENOR_DATAEMISSAO", gErr)
            
        Case 109171
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_CREDITO_NAO_PREENCHIDA", gErr)
            
        Case 109183
        
        Case 109184
            Call Rotina_Erro(vbOKOnly, "ERRO_BORDERODESCCHQ_SEM_CHEQUESPRE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143714)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'libera as variáveis globais
    Set gobjBorderoDescChq = Nothing

End Sub

Private Sub Limpa_Tela_BorderoDescChq()

    Call Limpa_Tela(Me)
    
    'limpa as combos de cobrador e conta corrente
    Cobrador.ListIndex = -1
    ContaCorrente.ListIndex = -1
    
    'esvazia a combo de carteiracobranca
    CarteiraCobranca.ListIndex = -1
    CarteiraCobranca.Clear
    
    'preecher a data emissão com a data atual
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    'preencher a data bom para até com a data atual
    DataBomParaAte.PromptInclude = False
    DataBomParaAte.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataBomParaAte.PromptInclude = True
    
    'se o módulo de contabilidade estiver ativo, preenche a data de contabilidade com a atual
    If gcolModulo.ATIVO(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        
        DataContabil.PromptInclude = False
        DataContabil.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataContabil.PromptInclude = True
        
    End If
    
    iAlterado = 0

End Sub

Private Function Traz_BorderoDescChq_Tela(ByVal objBorderoDescChq As ClassBorderoDescChq) As Long

On Error GoTo Erro_Traz_BorderoDescChq_Tela

    'limpa a tela
    Call Limpa_Tela_BorderoDescChq
    
    'se o cobrador estiver preenchido
    If objBorderoDescChq.iCobrador <> 0 Then
    
        Cobrador.Text = objBorderoDescChq.iCobrador
        Call Cobrador_Validate(bSGECancelDummy)
    
    End If
    
    'se a carteira estiver preenchida
    If objBorderoDescChq.iCarteiraCobranca <> 0 Then
    
        CarteiraCobranca.Text = objBorderoDescChq.iCarteiraCobranca
        Call CarteiraCobranca_Validate(bSGECancelDummy)
    
    End If
    
    'se a CC estiver preenchida
    If objBorderoDescChq.iContaCorrente <> 0 Then
    
        ContaCorrente.Text = objBorderoDescChq.iContaCorrente
        Call ContaCorrente_Validate(bSGECancelDummy)
        
    End If
    
    'se a data de emissão estiver preenchida
    If objBorderoDescChq.dtDataEmissao <> DATA_NULA Then
    
        DataEmissao.PromptInclude = False
        DataEmissao.Text = Format(objBorderoDescChq.dtDataEmissao, "dd/mm/yy")
        DataEmissao.PromptInclude = True
        
    End If
    
    'se a data contábil estiver preenchida
    If objBorderoDescChq.dtDataContabil <> DATA_NULA Then
    
        DataContabil.PromptInclude = False
        DataContabil.Text = Format(objBorderoDescChq.dtDataContabil, "dd/mm/yy")
        DataContabil.PromptInclude = True
        
    End If

    'se a data de crédito estiver preenchida
    If objBorderoDescChq.dtDataDeposito <> DATA_NULA Then
    
        DataCredito.PromptInclude = False
        DataCredito.Text = Format(objBorderoDescChq.dtDataDeposito, "dd/mm/yy")
        DataCredito.PromptInclude = True
        
    End If
    
    'se a data de Bom para até estiver preenchida
    If objBorderoDescChq.dtDataBomParaAte <> DATA_NULA Then
    
        DataBomParaAte.PromptInclude = False
        DataBomParaAte.Text = Format(objBorderoDescChq.dtDataBomParaAte, "dd/mm/yy")
        DataBomParaAte.PromptInclude = True
        
    Else
        
        DataBomParaAte.PromptInclude = False
        DataBomParaAte.Text = ""
        DataBomParaAte.PromptInclude = True
    
    End If
    
    'se o banco estiver preenchido
    If objBorderoDescChq.iBanco <> 0 Then Banco.Text = objBorderoDescChq.iBanco
    
    'preenche a agência
    If Len(Trim(objBorderoDescChq.sAgencia)) <> 0 Then Agencia.Text = objBorderoDescChq.sAgencia
    
    iAlterado = 0

    Traz_BorderoDescChq_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_BorderoDescChq_Tela:
    
    Traz_BorderoDescChq_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143715)
            
    End Select
    
    Exit Function

End Function

Private Sub Move_Tela_Memoria(ByVal objBorderoDescChq As ClassBorderoDescChq)

On Error GoTo Erro_Move_Tela_Memoria

    objBorderoDescChq.iCobrador = Codigo_Extrai(Cobrador.Text)
    objBorderoDescChq.iCarteiraCobranca = Codigo_Extrai(CarteiraCobranca.Text)
    objBorderoDescChq.iContaCorrente = Codigo_Extrai(ContaCorrente.Text)
    objBorderoDescChq.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objBorderoDescChq.dtDataDeposito = StrParaDate(DataCredito.Text)
    objBorderoDescChq.dtDataContabil = StrParaDate(DataContabil.Text)
    objBorderoDescChq.dtDataBomParaAte = StrParaDate(DataBomParaAte.Text)
    objBorderoDescChq.iBanco = StrParaInt(Banco.Text)
    objBorderoDescChq.sAgencia = Trim(Agencia.Text)

    Exit Sub
    
Erro_Move_Tela_Memoria:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143716)
    
    End Select
        
    Exit Sub

End Sub

Private Function Carrega_CarteiraCobranca(ByVal objCobrador As ClassCobrador) As Long

Dim lErro As Long
Dim colCarteirasCobranca As New Collection
Dim objCarteiraCobranca As ClassCarteiraCobranca

On Error GoTo Erro_Carrega_CarteiraCobranca

    'lê as carteiras de um determinado cobrador.
    lErro = CF("Cobrador_Le_Carteiras2", objCobrador.iCodigo, colCarteirasCobranca)
    If lErro <> SUCESSO And lErro <> 109175 Then gError 109141
    
    'se não encontrar nenhuma carteira -> erro
    If lErro = 109175 Then gError 109142
    
    'preenche a combo de carteiras
    For Each objCarteiraCobranca In colCarteirasCobranca
    
        CarteiraCobranca.AddItem (objCarteiraCobranca.iCodigo & SEPARADOR & objCarteiraCobranca.sDescricao)
        CarteiraCobranca.ItemData(CarteiraCobranca.NewIndex) = objCarteiraCobranca.iCodigo
    
    Next

    Carrega_CarteiraCobranca = SUCESSO
    
    Exit Function
    
Erro_Carrega_CarteiraCobranca:

    Carrega_CarteiraCobranca = gErr
    
    Select Case gErr
    
        Case 109141, 109142
        'o erro 109142 não mostrará msg a exemplo da tela borderocobranca.
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143717)
    
    End Select
    
    Exit Function

End Function

Public Function Carrega_Cobradores() As Long

Dim colCobradores As New Collection
Dim objCobrador As ClassCobrador
Dim lErro As Long

On Error GoTo Erro_Carrega_Cobradores

    'lê todos os cobradores da filial e o carrega na coleção
    lErro = CF("Cobradores_Le_Todos_Filial", colCobradores)
    If lErro <> SUCESSO Then gError 109134
    
    'preenche a combo de cobradores com exceção da própria empresa
    For Each objCobrador In colCobradores
    
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA Then
            
            Cobrador.AddItem (objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido)
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo
        
        End If
    
    Next
    
    Carrega_Cobradores = SUCESSO
    
    Exit Function
    
Erro_Carrega_Cobradores:

    Carrega_Cobradores = gErr
    
    Select Case gErr
    
        Case 109134
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143718)
            
    End Select
    
    Exit Function

End Function

Private Function Carrega_ContasCorrentes() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colContasCorrentesInternas As New Collection
Dim objContaCorrenteInterna As New ClassContasCorrentesInternas

On Error GoTo Erro_Carrega_ContasCorrentes

    'Leitura dos códigos e descrições das Contas no BD
    lErro = CF("ContasCorrentesInternas_Le_Todas", colContasCorrentesInternas)
    If lErro <> SUCESSO Then gError 109137

    'Preenche listbox com descrição das contas
    For Each objContaCorrenteInterna In colContasCorrentesInternas
    
        'se não for uma conta de cheque pre
        If objContaCorrenteInterna.iChequePre <> CONTA_CHEQUE_PRE Then
        
            ContaCorrente.AddItem objContaCorrenteInterna.iCodigo & SEPARADOR & objContaCorrenteInterna.sNomeReduzido
            ContaCorrente.ItemData(ContaCorrente.NewIndex) = objContaCorrenteInterna.iCodigo
        
        End If
    
    Next
    
    Carrega_ContasCorrentes = SUCESSO

    Exit Function

Erro_Carrega_ContasCorrentes:

    Carrega_ContasCorrentes = gErr

    Select Case gErr

        Case 109137

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143719)

    End Select

    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    ' ???Parent.HelpContextID = IDH_BORDERO_DESCCHQ1
    Set Form_Load_Ocx = Me
    Caption = "Borderô de Desconto de Cheques - Passo 1"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoDescChq1"
    
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
