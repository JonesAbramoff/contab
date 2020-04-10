VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl OperacaoArqCCBack 
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   ScaleHeight     =   3135
   ScaleWidth      =   4455
   Begin VB.Frame Frame2 
      Caption         =   "Intervalo de Trasmissão de Log"
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   2655
      Begin VB.TextBox LogFinal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Final:"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.Label LabelLogInicial 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Inicial:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operação"
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   2625
      Begin VB.OptionButton OptionTransmitir 
         Caption         =   "Transmitir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.OptionButton OptionRetransmitir 
         Caption         =   "Retransmitir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3165
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1140
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "OperacaoArqCCBack.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "OperacaoArqCCBack.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.ProgressBar BarraProgresso 
      Height          =   345
      Left            =   375
      TabIndex        =   11
      Top             =   2580
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "OperacaoArqCCBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Dim glNumIntDocFinal As Long

Event Unload()

Function Retransmitir_Arquivo() As Long

Dim sRetorno As String
Dim lTamanho As Long
Dim iSeqArq As Integer
Dim iCodEmpresa As Integer
Dim iIndice As Integer
Dim sNomeArq As String
Dim sNomeArq1 As String
Dim sRegistro As String
Dim bAchou As Boolean
Dim bEntrou As Boolean
Dim sNomeArq2 As String
Dim iPos As Integer
Dim sTipo As String
Dim sSeq As String
Dim colSeq As New Collection
Dim iIndiceInicio As Integer
Dim sRet As String
Dim sNomeArq3 As String
Dim iIndiceSeq As Integer
Dim iIndiceFim As Integer
Dim iIndiceAnt As Integer
Dim lErro As Long
Dim objBarraProgresso As Object
Dim colNomeArq As New Collection
Dim vNomeArq As Variant
Dim lIntervaloTrans As Long
Dim sNomeArqParam As String
Dim objObject As Object

On Error GoTo Erro_Retransmitir_Arquivo
    
    Call Chama_Tela_Modal("ExibirSequenciaisCCBack", colSeq)
            
    Set objBarraProgresso = BarraProgresso
            
    lErro = CF("Rotina_Gravacao_CC_Back", colNomeArq, glNumIntDocFinal, objBarraProgresso, 1, giFilialEmpresa, colSeq)
    If lErro <> SUCESSO Then gError 118912
    
    For Each vNomeArq In colNomeArq
        sNomeArq = sNomeArq & vNomeArq & " "
    Next
    
    If Len(sNomeArq) > 0 Then sNomeArq = Left(sNomeArq, Len(sNomeArq) - 1)
    
     If Len(Trim(gobjLoja.sFTPURL)) > 0 Then
            
        'Prepara para chamar rotina batch
        lErro = Sistema_Preparar_Batch(sNomeArqParam)
        If lErro <> SUCESSO Then gError 133593
            
        gobjLoja.sNomeArqParam = sNomeArqParam
            
        Set objObject = gobjLoja
            
        lErro = CF("Rotina_FTP_Recepcao_CC", objObject, 5)
        If lErro <> SUCESSO And lErro <> 133628 Then gError 133594
            
        If lErro <> SUCESSO Then gError 133634
            
        'Pergunta se Deseja Efetuar a Sangria
        Call Rotina_Aviso(vbOK, "AVISO_ARQUIVOS_TRANSMITIDOS", sNomeArq)
        
    Else
    
        'Pergunta se Deseja Efetuar a Sangria
        Call Rotina_Aviso(vbOK, "AVISO_ARQUIVOS_GERADOS", sNomeArq)
    
    End If
    
    Retransmitir_Arquivo = SUCESSO
    
    Exit Function
    
Erro_Retransmitir_Arquivo:
    
    Retransmitir_Arquivo = gErr

    Select Case gErr
                        
        Case 118912
        
        Case 133593, 133594
        
        Case 133634
            Call Rotina_Erro(vbOKOnly, "AVISO_NAO_CARREGOU_ROTINA_RECEPCAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 163641)
        
    End Select
    
    Exit Function
    
End Function

Private Function Transmitir_Arquivo() As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objBarraProgresso As Object
Dim colNomeArq As New Collection
Dim sNomeArq As String
Dim vNomeArq As Variant
Dim lIntervaloTrans As Long
Dim sNomeArqParam As String
Dim objObject As Object

On Error GoTo Erro_Transmitir_Arquivo

    'se o log final não está entre os limites --> erro.
    If StrParaLong(LogFinal.Text) > glNumIntDocFinal Or StrParaLong(LogFinal.Text) < StrParaLong(LabelLogInicial.Caption) Then gError 118911
    
    lNumIntDoc = StrParaLong(LogFinal.Text)
    
    Set objBarraProgresso = BarraProgresso
    
    lErro = CF("Rotina_Gravacao_CC_Back", colNomeArq, lNumIntDoc, objBarraProgresso, 0, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 118912
    
     If Len(Trim(gobjLoja.sFTPURL)) > 0 Then
            
        'Prepara para chamar rotina batch
        lErro = Sistema_Preparar_Batch(sNomeArqParam)
        If lErro <> SUCESSO Then gError 133592
            
        gobjLoja.sNomeArqParam = sNomeArqParam
            
        Set objObject = gobjLoja
            
        lErro = CF("Rotina_FTP_Recepcao_CC", objObject, 5)
        If lErro <> SUCESSO And lErro <> 133628 Then gError 133591
            
        If lErro <> SUCESSO Then gError 133635
            
        'Pergunta se Deseja Efetuar a Sangria
        Call Rotina_Aviso(vbOK, "AVISO_TRANSMISSAO_CONCLUIDA")
        
    Else
    
        'Pergunta se Deseja Efetuar a Sangria
        Call Rotina_Aviso(vbOK, "AVISO_ARQUIVO_GERADO", colNomeArq.Item(1))
    
    End If
    
    Transmitir_Arquivo = SUCESSO
    
    Exit Function
    
Erro_Transmitir_Arquivo:

    Transmitir_Arquivo = gErr
    
    Select Case gErr
        
        Case 118911
            Call Rotina_Erro(vbOKOnly, "ERRO_LOG_FORA_LIMITE", gErr, StrParaInt(LabelLogInicial.Caption), glNumIntDocFinal)
            
        Case 118912, 133591, 133592
                        
        Case 133635
            Call Rotina_Erro(vbOKOnly, "AVISO_NAO_CARREGOU_ROTINA_RECEPCAO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163642)
    
    End Select

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    If OptionTransmitir Then
        lErro = Transmitir_Arquivo
        If lErro <> SUCESSO Then gError 118900
    Else
        lErro = Retransmitir_Arquivo
        If lErro <> SUCESSO Then gError 118901
    End If
    
    'arquivo inexistente
    If lErro = 53 Then gError 118902
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
        
        Case 118900, 118901
        
        Case 118902
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_INEXISTENTE", gErr, NOME_ARQUIVOBACK)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163643)
    
    End Select

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a gravar registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 118903
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
    
        Case 118903
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 163644)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

Dim lNumIntDocInicial As Long
Dim lNumIntDocFinal As Long
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    lErro = CF("ControleLogCCBack_Le_Limites", giFilialEmpresa, lNumIntDocInicial, lNumIntDocFinal)
    If lErro <> SUCESSO And lErro <> 118910 Then gError 118904
    
    If lErro <> 118910 Then
        LogFinal.Enabled = True
        LabelLogInicial.Caption = lNumIntDocInicial
        glNumIntDocFinal = lNumIntDocFinal
        LogFinal.Text = lNumIntDocFinal
    Else
        LabelLogInicial.Caption = 0
        glNumIntDocFinal = 0
        LogFinal.Text = 0
        OptionRetransmitir.Value = True
        OptionTransmitir.Enabled = False
    End If
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
        
        Case 118904
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 163645)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 163646)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Operação de Arquivo de Caixa Central para backoffice"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "OperacaoArqCCBack"

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

Private Sub OptionRetransmitir_Click()
    
    LogFinal.Enabled = False
    
End Sub

Private Sub OptionTransmitir_Click()
    
    LogFinal.Enabled = True
    
End Sub

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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'**** fim do trecho a ser copiado *****


