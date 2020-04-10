VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ConciliarExtratoBancarioOcx 
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   KeyPreview      =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6570
   Begin VB.ComboBox ContaCorrente 
      Height          =   315
      Left            =   2205
      TabIndex        =   7
      Top             =   270
      Width           =   2280
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critérios para Conciliação"
      Height          =   675
      Left            =   435
      TabIndex        =   6
      Top             =   1230
      Width           =   4065
      Begin VB.CheckBox Valor 
         Caption         =   "Valor"
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
         Left            =   1492
         TabIndex        =   12
         Top             =   315
         Value           =   1  'Checked
         Width           =   885
      End
      Begin VB.CheckBox Historico 
         Caption         =   "Histórico"
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
         Left            =   2640
         TabIndex        =   11
         Top             =   315
         Width           =   1140
      End
      Begin VB.CheckBox Data 
         Caption         =   "Data"
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
         Left            =   345
         TabIndex        =   10
         Top             =   315
         Value           =   1  'Checked
         Width           =   885
      End
   End
   Begin VB.CheckBox CheckLctosNaoConciliados 
      Caption         =   "Tentar conciliar lançamentos não conciliados de extrato já conciliado"
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
      Left            =   135
      TabIndex        =   1
      Top             =   2070
      Width           =   6240
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   570
      Left            =   5370
      Picture         =   "ConciliarExtratoBancarioOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   825
      Width           =   990
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   570
      Left            =   5385
      Picture         =   "ConciliarExtratoBancarioOcx.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   165
      Width           =   990
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   375
      TabIndex        =   5
      Top             =   2760
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   300
      Left            =   2985
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   780
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox NumExtrato 
      Height          =   300
      Left            =   2205
      TabIndex        =   0
      Top             =   780
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Processamento"
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
      Top             =   2535
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Extratos à partir do:"
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
      Left            =   450
      TabIndex        =   8
      Top             =   825
      Width           =   1695
   End
   Begin VB.Label LabelCtaCorrente 
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
      Left            =   795
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   9
      Top             =   315
      Width           =   1350
   End
End
Attribute VB_Name = "ConciliarExtratoBancarioOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Falta revisao Jones

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoContaCorrenteInt As AdmEvento
Attribute objEventoContaCorrenteInt.VB_VarHelpID = -1

Private Sub BotaoCancela_Click()

    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim iCodConta As Integer
Dim bConciliaExtrato As Boolean
Dim iNumExtratoIni As Long
Dim colExtratos As New Collection
Dim iIndice As Integer
Dim iData As Integer
Dim iValor As Integer
Dim iHistorico As Integer
Dim iNumExtrato As Integer

On Error GoTo Erro_BotaoOK_Click

    'Verifica se a Conta Corrente esta preenchida.
    If ContaCorrente.Text = "" Then Error 22028
    'Verifica se o Extrato esta preenchido. Se nao foi, erro
    If NumExtrato.Text = "" Then Error 22029
    If Data.Value = vbUnchecked And Valor.Value = vbUnchecked And Historico.Value = vbUnchecked Then Error 62306
    
    iData = Data.Value
    iValor = Valor.Value
    iHistorico = Historico.Value

    'Recolhe os dados da tela
    iCodConta = Codigo_Extrai(ContaCorrente.Text)
    iNumExtratoIni = StrParaLong(NumExtrato)
    If CheckLctosNaoConciliados.Value = vbChecked Then
        bConciliaExtrato = True
    Else
        bConciliaExtrato = False
    End If

    BotaoOK.Enabled = False
            
    'Busca no BD os extratos no intervalo pedido
    lErro = CF("ConciliacaoBancaria_Obter_Extratos", iCodConta, iNumExtratoIni, bConciliaExtrato, colExtratos)
    If lErro <> SUCESSO Then Error 62307
    'Se não achou nenhum --> erro
    If colExtratos.Count = 0 Then Error 62309
    
    ProgressBar1.Value = 0
    'Para cada extrato encontrado
    For iIndice = 1 To colExtratos.Count
        'Pega o numero do extrato
        iNumExtrato = colExtratos(iIndice)
        'Faz a conciliação de seus lançamentos
        lErro = CF("ConciliacaoBancaria_Automatica_Grava", iCodConta, iNumExtrato, iData, iValor, Historico)
        If lErro <> SUCESSO Then Error 62309
        'Atualiza o número de extratos conciliados
        ProgressBar1.Value = CLng(iIndice) * 100 / colExtratos.Count
        
    Next
    
    BotaoOK.Enabled = True
    
    Exit Sub

Erro_BotaoOK_Click:

    BotaoOK.Enabled = True

    Select Case Err

        Case 22028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)


        Case 22029
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXTRATO_NAO_INFORMADO", Err)

        Case 62306
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CRITERIOS_NAO_SELECIONADOS", Err)
        
        Case 62307, 62309

        Case 62308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXTRATOS_NAO_ENCONTRADOS", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154540)

    End Select

    Exit Sub

End Sub

Private Sub ContaCorrente_Click()

Dim lErro As Long
Dim iCodConta As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_ContaCorrente_Click

    If ContaCorrente.ListIndex = -1 Then Exit Sub
    
    'Extrai o Código da Conta que está na tela
    iCodConta = Codigo_Extrai(ContaCorrente.Text)

    'Passa o Código da Conta para o Obj
    objContaCorrenteInt.iCodigo = iCodConta

    'Lê os dados da Conta
    lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 22025

    'Se a Conta não estiver cadastrada
    If lErro = 11807 Then Error 22030

    'Se a Conta não é Bancária
    If objContaCorrenteInt.iCodBanco = 0 Then Error 22031

    'Preenche um default para Extrato
    NumExtrato.Text = objContaCorrenteInt.iNumMenorExtratoNaoConciliado

    Exit Sub

Erro_ContaCorrente_Click:

    Select Case Err

        Case 22025

        Case 22030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, iCodConta)

        Case 22031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", Err, ContaCorrente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154541)

    End Select

    Exit Sub

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 43527

    'Se a Conta(CODIGO) não existe na Combo
    If lErro = 6730 Then

        objContaCorrenteInt.iCodigo = iCodigo

        'Lê os dados da Conta
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 43528

        'Se a Conta não estiver cadastrada
        If lErro = 11807 Then Error 43529

        'Se a Conta não é Bancária
        If objContaCorrenteInt.iCodBanco = 0 Then Error 43530

        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            'Se a Conta não é da Filial selecionada
            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43531

        End If

        'Passa o código da Conta para a tela
        ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Se a Conta(STRING) não existe na Combo
    If lErro = 6731 Then Error 43532

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 43527, 43528

        Case 43529
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If

        Case 43530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", Err, ContaCorrente.Text)

        Case 43531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaCorrente.Text, giFilialEmpresa)

        Case 43532
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, ContaCorrente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154542)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoContaCorrenteInt = Nothing

End Sub

Private Sub LabelCtaCorrente_Click()

Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim colSelecao As Collection

    If Len(ContaCorrente.Text) = 0 Then
        objContaCorrenteInt.iCodigo = 0
    Else
        objContaCorrenteInt.iCodigo = Codigo_Extrai(ContaCorrente.Text)
    End If

    'Se alguma Filial tiver sido selecionada
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        'Chama a tela com a lista das contas correntes bancarias
        Call Chama_Tela("CtaCorrBancariaLista", colSelecao, objContaCorrenteInt, objEventoContaCorrenteInt)

        
    Else
        'Chama a tela com a lista de todas as contas correntes bancarias
        Call Chama_Tela("CtaCorrBancariaTodasLista", colSelecao, objContaCorrenteInt, objEventoContaCorrenteInt)

    End If

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoNomeConta As New AdmColCodigoNome
Dim objCodigoNomeConta As New AdmCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoContaCorrenteInt = New AdmEvento

    'Carrega a Coleção de Contas
    lErro = CF("ContasCorrentes_Bancarias_Le_CodigosNomesRed", colCodigoNomeConta)
    If lErro <> SUCESSO Then Error 22023

    'Preenche a ComboBox Conta com os objetos da coleção de Contas
    For Each objCodigoNomeConta In colCodigoNomeConta

        ContaCorrente.AddItem CStr(objCodigoNomeConta.iCodigo) & SEPARADOR & objCodigoNomeConta.sNome
        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNomeConta.iCodigo

    Next

    'Seleciona uma das Contas
    If ContaCorrente.ListCount > 0 Then ContaCorrente.ListIndex = 0

    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    ProgressBar1.Value = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 22023

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154543)

    End Select

    Exit Sub

End Sub

Private Sub NumExtrato_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumExtrato)

End Sub

Private Sub objEventoContaCorrenteInt_evSelecao(obj1 As Object)
Dim objContaCorrenteInt As ClassContasCorrentesInternas

Dim lErro As Long

On Error GoTo Erro_objEventoContaCorrenteInt_evSelecao

    Set objContaCorrenteInt = obj1

    'Traz para tela os dados da conta selecionada
    ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido
    Call ContaCorrente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoContaCorrenteInt_evSelecao:

    Select Case Err

        Case 22032

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154544)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim iExtrato As Integer

On Error GoTo Erro_UpDown1_DownClick

    'Diminui o numero do extrato
    If Len(Trim(NumExtrato.Text)) = 0 Then
        iExtrato = 0
    Else
        iExtrato = CInt(NumExtrato.Text)
    End If

    'número do extrato não pode ser menor que 0
    If iExtrato = 0 Then Error 22026

    NumExtrato.Text = CStr(iExtrato - 1)

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 22026

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154545)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim iExtrato As Integer

On Error GoTo Erro_UpDown1_UpClick

    'Aumenta o numero do extrato
    If Len(Trim(NumExtrato.Text)) = 0 Then
        iExtrato = 0
    Else
        iExtrato = CInt(NumExtrato.Text)
    End If

    'número do Extrato não pode ser maior que 9999
    If iExtrato = 9999 Then Error 22027

    NumExtrato.Text = CStr(iExtrato + 1)

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 22027

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154546)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONCILIAR_EXTRATO_BANCARIO
    Set Form_Load_Ocx = Me
    Caption = "Conciliar Extrato Bancário"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConciliarExtratoBancario"
    
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
            Call LabelCtaCorrente_Click
        End If
    
    End If
    
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(label1, Button, Shift, X, Y)
End Sub

Private Sub LabelCtaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCtaCorrente, Source, X, Y)
End Sub

Private Sub LabelCtaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCtaCorrente, Button, Shift, X, Y)
End Sub


Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

