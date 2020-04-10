VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BorderoChequesPreOcx 
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ScaleHeight     =   2550
   ScaleWidth      =   6000
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   4140
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   1680
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1125
         Picture         =   "BorderoChequesPreOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoSeguir 
         Height          =   330
         Left            =   90
         Picture         =   "BorderoChequesPreOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   90
         Width           =   930
      End
   End
   Begin VB.ComboBox ContaCorrente 
      Height          =   315
      Left            =   1605
      TabIndex        =   0
      Text            =   "ContaCorrente"
      Top             =   120
      Width           =   2115
   End
   Begin MSComCtl2.UpDown UpDownEmissao 
      Height          =   300
      Left            =   2760
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   750
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Emissao 
      Height          =   300
      Left            =   1590
      TabIndex        =   1
      Top             =   750
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDeposito 
      Height          =   300
      Left            =   2745
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1425
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataDeposito 
      Height          =   300
      Left            =   1590
      TabIndex        =   2
      Top             =   1425
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
      Left            =   2760
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2070
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataContabil 
      Height          =   300
      Left            =   1590
      TabIndex        =   3
      Top             =   2055
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label9 
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
      Left            =   930
      TabIndex        =   7
      Top             =   210
      Width           =   570
   End
   Begin VB.Label Label8 
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
      Left            =   765
      TabIndex        =   8
      Top             =   810
      Width           =   765
   End
   Begin VB.Label LabelDeposito 
      AutoSize        =   -1  'True
      Caption         =   "Cheques p/depósito até:"
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
      Height          =   450
      Left            =   180
      TabIndex        =   9
      Top             =   1320
      Width           =   1365
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelContabil 
      AutoSize        =   -1  'True
      Caption         =   "Data Contábil do Borderô:"
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
      Height          =   390
      Left            =   300
      TabIndex        =   10
      Top             =   2010
      Width           =   1260
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "BorderoChequesPreOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoContaCorrente As AdmEvento
Attribute objEventoContaCorrente.VB_VarHelpID = -1
Dim gobjBorderoChequePre As ClassBorderoChequePre
Dim iBorderoAlterado As Integer



Private Sub BotaoFechar_Click()
        
        Unload Me
        
End Sub

Private Sub BotaoSeguir_Click()
 
Dim lErro As Long
Dim iCodConta As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_BotaoSeguir_Click

    'Verifica se a ContaCorrente está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then gError 22008
    
    'Verifica se a DataEmissao está preenchida
    If Len(Trim(Emissao.ClipText)) = 0 Then gError 22012
    
    'Verifica se a DataDeposito está preenchida
    If Len(Trim(DataDeposito.ClipText)) = 0 Then gError 22013

    'Verifica se a DataContabil está preenchida
    If Len(Trim(DataContabil.ClipText)) = 0 Then gError 22014

    'Extrai o Código da Conta que está na tela
    iCodConta = Codigo_Extrai(ContaCorrente.Text)

    'Passa o Código da Conta para o Obj
    objContaCorrenteInt.iCodigo = iCodConta

    'Lê os dados da Conta
    lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 22009

    'Se a Conta não estiver cadastrada
    If lErro = 11807 Then gError 22010

    'Se a Conta não é Bancária
    If objContaCorrenteInt.iCodBanco = 0 Then gError 22011
    
    'Se alguma Filial tiver sido selecionada
    If giFilialEmpresa <> EMPRESA_TODA Then
        
        'Se a Conta não é da Filial selecionada
        If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then gError 22017
        
    End If
    
    'Verifica se a DataDeposito é maior ou igual que a DataEmissao
    If CDate(DataDeposito.Text) < CDate(Emissao.Text) Then gError 22015
    
    'Verifica se a DataContabil é maior ou igual que a DataEmissao
    If CDate(DataContabil.Text) < CDate(Emissao.Text) Then gError 22016
    
    'Se não leu os cheques ainda ou alterou a data de depósito inicializa o obj
    If gobjBorderoChequePre.colchequepre.Count = 0 Or gobjBorderoChequePre.dtDataDeposito <> StrParaDate(DataDeposito.Text) Then 'iBorderoAlterado = REGISTRO_ALTERADO Then
        Set gobjBorderoChequePre = New ClassBorderoChequePre
    End If
    gobjBorderoChequePre.dtDataEmissao = StrParaDate(Emissao.Text)
    gobjBorderoChequePre.dtDataDeposito = StrParaDate(DataDeposito.Text)
    gobjBorderoChequePre.dtDataContabil = StrParaDate(DataContabil.Text)
    gobjBorderoChequePre.iCodNossaConta = iCodConta

    'Verifica se houve alterações dos dados passados para o filtro do cheque
    If gobjBorderoChequePre.colchequepre.Count = 0 Then  'iBorderoAlterado = REGISTRO_ALTERADO Then
        
        'Set gobjBorderoChequePre = New ClassBorderoChequePre
        
        'gobjBorderoChequePre.dtDataEmissao = CDate(Emissao.Text)
        'gobjBorderoChequePre.dtDataDeposito = CDate(DataDeposito.Text)
        'gobjBorderoChequePre.dtDataContabil = CDate(DataContabil.Text)
        'gobjBorderoChequePre.iCodNossaConta = iCodConta
                                                                                                                                                                                                               
        lErro = CF("BorderoChequePre_Le_Cheques", gobjBorderoChequePre.colchequepre, gobjBorderoChequePre.dtDataDeposito)
        If lErro <> SUCESSO Then gError 80337
    
    End If
        
    'Chama a tela do passo seguinte
    Call Chama_Tela("BorderoChequesPre1", gobjBorderoChequePre)
    
    'Fecha tela
    Unload Me
    
    Exit Sub

Erro_BotaoSeguir_Click:

    Select Case gErr
    
        Case 22008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)
    
        Case 22009, 7743
        
        Case 22010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, iCodConta)
            
        Case 22011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", gErr, ContaCorrente.Text)
            
        Case 22012, 22013, 22014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
            
        Case 22015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATADEPOSITO_MENOR_DATAEMISSAO", gErr)
        
        Case 22016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATACONTABIL_MENOR_DATAEMISSAO", gErr)
    
        Case 22017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", gErr, ContaCorrente.Text, giFilialEmpresa)
            ContaCorrente.SetFocus
        
        Case 80337
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143649)

    End Select

    Exit Sub

End Sub

Private Sub ContaCorrente_Change()

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 43081

    'Se a Conta(CODIGO) não existe na Combo
    If lErro = 6730 Then

        objContaCorrenteInt.iCodigo = iCodigo

        'Lê os dados da Conta
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 43082

        'Se a Conta não estiver cadastrada
        If lErro = 11807 Then Error 43083

        'Se a Conta não é Bancária
        If objContaCorrenteInt.iCodBanco = 0 Then Error 43084

        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            'Se a Conta não é da Filial selecionada
            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43085

        End If

        'Passa o código da Conta para a tela
        ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Se a Conta(STRING) não existe na Combo
    If lErro = 6731 Then Error 43086

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 43081, 43082

        Case 43083
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If

        Case 43084
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", Err, ContaCorrente.Text)

        Case 43085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaCorrente.Text, giFilialEmpresa)

        Case 43086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, ContaCorrente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143650)

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

Dim lErro As Long

On Error GoTo Erro_DataContabil_Validate

    'Verifica se a data de deposito está preenchida
    If Len(Trim(DataContabil.ClipText)) = 0 Then Exit Sub

    'Verifica se a data final é válida
    lErro = Data_Critica(DataContabil.Text)
    If lErro <> SUCESSO Then Error 43080

    Exit Sub

Erro_DataContabil_Validate:

    Cancel = True


    Select Case Err

        Case 43080

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143651)

    End Select

    Exit Sub

End Sub

Private Sub DataDeposito_Change()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDeposito_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataDeposito)

End Sub

Private Sub DataDeposito_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDeposito_Validate

    'Verifica se a data de depósito está preenchida
    If Len(Trim(DataDeposito.ClipText)) = 0 Then Exit Sub

    'Verifica se a data final é válida
    lErro = Data_Critica(DataDeposito.Text)
    If lErro <> SUCESSO Then Error 43079

    Exit Sub

Erro_DataDeposito_Validate:

    Cancel = True


    Select Case Err

        Case 43079

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143652)

    End Select

    Exit Sub

End Sub

Private Sub Emissao_Change()

    iBorderoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Emissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Emissao)

End Sub

Private Sub Emissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Emissao_Validate

    'Verifica se a data de emissão está preenchida
    If Len(Trim(Emissao.ClipText)) = 0 Then Exit Sub

    'Verifica se a data final é válida
    lErro = Data_Critica(Emissao.Text)
    If lErro <> SUCESSO Then Error 43087

    Exit Sub

Erro_Emissao_Validate:

    Cancel = True


    Select Case Err

        Case 43087

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143653)

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
    If lErro <> SUCESSO Then Error 22001

    'Preenche a ComboBox ContaConta com os objetos da coleção de Contas
    For Each objCodigoNomeConta In colCodigoNomeConta

        ContaCorrente.AddItem CStr(objCodigoNomeConta.iCodigo) & SEPARADOR & objCodigoNomeConta.sNome
        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNomeConta.iCodigo

    Next

    'Seleciona uma das Contas
    ContaCorrente.Text = ContaCorrente.List(PRIMEIRA_CONTA)

    'Preenche as Datas com a data corrente do sistema
    Emissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataDeposito.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataContabil.Text = Format(gdtDataAtual, "dd/mm/yy")

    'Verifica se o módulo de Contabilidade está ativo
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        LabelContabil.Visible = True
        DataContabil.Visible = True
        UpDownDataContabil.Visible = True
    Else
        LabelContabil.Visible = False
        DataContabil.Visible = False
        UpDownDataContabil.Visible = False
    End If
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 22001

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143654)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoContaCorrente = Nothing
    Set gobjBorderoChequePre = Nothing

End Sub

Private Sub UpDownEmissao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(Emissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 22002

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 22002

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143655)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(Emissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 22003

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 22003

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143656)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDeposito_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDeposito_DownClick

    'Diminui a DataVencimento em 1 dia
    lErro = Data_Up_Down_Click(DataDeposito, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 22004

    Exit Sub

Erro_UpDownDeposito_DownClick:

    Select Case Err

        Case 22004

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143657)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDeposito_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDeposito_UpClick

    'Aumenta a DataVencimento em 1 dia
    lErro = Data_Up_Down_Click(DataDeposito, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 22005

    Exit Sub

Erro_UpDownDeposito_UpClick:

    Select Case Err

        Case 22005

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143658)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataContabil_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataContabil_DownClick

    'Diminui a DataContabil em 1 dia
    lErro = Data_Up_Down_Click(DataContabil, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 22006

    Exit Sub

Erro_UpDownDataContabil_DownClick:

    Select Case Err

        Case 22006

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143659)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataContabil_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataContabil_UpClick

    'Aumenta a DataContabil em 1 dia
    lErro = Data_Up_Down_Click(DataContabil, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 22007

    Exit Sub

Erro_UpDownDataContabil_UpClick:

    Select Case Err

        Case 22007

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143660)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BORDERO_DEPOSITO_CHEQUES_PRE_DATADOS
    Set Form_Load_Ocx = Me
    Caption = "Borderô de Depósito de Cheques Pré-Datados"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoChequesPre"
    
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

Function Trata_Parametros(Optional objBorderoChequePre As ClassBorderoChequePre) As Long


    'Se a tela é aberta a partir do menu então deve setar o objBorderoChequePre
    If (objBorderoChequePre Is Nothing) Then
        Set gobjBorderoChequePre = New ClassBorderoChequePre
        
        gobjBorderoChequePre.dtDataEmissao = CDate(Emissao.Text)
        gobjBorderoChequePre.dtDataDeposito = CDate(DataDeposito.Text)
        gobjBorderoChequePre.dtDataContabil = CDate(DataContabil.Text)
        
    'Se a tela e chamada a partir de BorderoChequePre1 então preencher os dados
    'da tela BorderoChequePre com os dados do objBorderoChequePre
    Else
        
        Set gobjBorderoChequePre = objBorderoChequePre
    
        Emissao.Text = Format(objBorderoChequePre.dtDataEmissao, "dd/mm/yy")
        DataDeposito.Text = Format(objBorderoChequePre.dtDataDeposito, "dd/mm/yy")
        DataContabil.Text = Format(objBorderoChequePre.dtDataContabil, "dd/mm/yy")
        
        If objBorderoChequePre.colchequepre.Count > 0 Then
            
            If objBorderoChequePre.iCodNossaConta <> 0 Then
            
                ContaCorrente.Text = objBorderoChequePre.iCodNossaConta
                Call ContaCorrente_Validate(bSGECancelDummy)
            
            End If
                  
        End If
    
    End If
        
    iBorderoAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
End Function



Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub LabelDeposito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDeposito, Source, X, Y)
End Sub

Private Sub LabelDeposito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDeposito, Button, Shift, X, Y)
End Sub

Private Sub LabelContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContabil, Source, X, Y)
End Sub

Private Sub LabelContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContabil, Button, Shift, X, Y)
End Sub
