VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpEmissaoFaturaOcx 
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   ScaleHeight     =   1905
   ScaleWidth      =   3945
   Begin VB.Timer Timer1 
      Left            =   30
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   2730
      ScaleHeight     =   465
      ScaleWidth      =   1035
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   1095
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   45
         Picture         =   "RelOpEmissaoFaturaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   555
         Picture         =   "RelOpEmissaoFaturaOcx.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fatura"
      Height          =   900
      Left            =   120
      TabIndex        =   6
      Top             =   855
      Width           =   3720
      Begin MSMask.MaskEdBox FaturaInicial 
         Height          =   300
         Left            =   675
         TabIndex        =   0
         Top             =   360
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FaturaFinal 
         Height          =   300
         Left            =   2415
         TabIndex        =   1
         Top             =   360
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
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
         Left            =   2010
         TabIndex        =   8
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label14 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   285
         TabIndex        =   7
         Top             =   420
         Width           =   315
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   420
      Picture         =   "RelOpEmissaoFaturaOcx.ctx":06B0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Width           =   1815
   End
End
Attribute VB_Name = "RelOpEmissaoFaturaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim glNumAcessosTimer As Long

Public Sub Form_Load()

Dim lErro As Long
Dim sFaturaInicial As String
Dim sFaturaFinal As String

On Error GoTo Erro_Form_Load
    
    'Preenche os default na Tela
    'Para Fatura de
    lErro = CF("CRFatConfig_Le",CRFATCFG_FATURA_NUM_PROX_IMPRESSAO, EMPRESA_TODA, sFaturaInicial)
    If lErro <> SUCESSO Then Error 61466
    
    'Para Fatura Até
    lErro = CF("CRFatConfig_Le",CRFATCFG_FATURA_NUM_PROX, EMPRESA_TODA, sFaturaFinal)
    If lErro <> SUCESSO Then Error 61467
    
    FaturaInicial.Text = sFaturaInicial
    FaturaFinal.Text = CStr(CLng(sFaturaFinal) - 1)
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 38213, 61466, 61467

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168448)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 22886
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 38211
        
        Case 22886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168449)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Critica_Parametros() As Long
'Critica os parâmetros que serão passados para o relatório

Dim lErro As Long

On Error GoTo Erro_Critica_Parametros
          
    If Len(Trim(FaturaInicial.Text)) = 0 Then gError 64467
    
    If Len(Trim(FaturaFinal.Text)) = 0 Then gError 64468
    
    'Verifica se o numero da Fatura inicial é maior que o da final
    If Len(Trim(FaturaInicial.ClipText)) > 0 And Len(Trim(FaturaFinal.ClipText)) > 0 Then
    
        If CLng(FaturaInicial.Text) > CLng(FaturaFinal.Text) Then gError 38219
    
    End If
        
    Critica_Parametros = SUCESSO

    Exit Function

Erro_Critica_Parametros:

    Critica_Parametros = gErr

    Select Case gErr
        
        Case 64467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURADE_NAO_PREENCHIDA", gErr)
        
        Case 64468
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURAATE_NAO_PREENCHIDA", gErr)
        
        Case 38219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)
            FaturaInicial.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168450)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47118
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47118
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168451)

    End Select

    Exit Sub
    
End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Parametros()
    If lErro <> SUCESSO Then Error 38222
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 38223
    
    lErro = objRelOpcoes.IncluirParametro("NFATURAINIC", FaturaInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38224

    lErro = objRelOpcoes.IncluirParametro("NFATURAFIM", FaturaFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 38225
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 38227
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 38222 To 38227

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168452)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim sNumProxFaturaImpressao As String
Dim lFaixaFinal As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 38230
    
    'Lock da Impressao da Fatura
    lErro = Fatura_Lock_Impressao
    If lErro <> SUCESSO Then Error 61468
    
    'Le o próximo número da Impressao da Fatura
    lErro = CF("CRFatConfig_Le",CRFATCFG_FATURA_NUM_PROX_IMPRESSAO, EMPRESA_TODA, sNumProxFaturaImpressao)
    If lErro <> SUCESSO Then Error 61469
    
    'Dá Mensagem ao usuário caso seja Reimpressão
    If CLng(FaturaInicial.Text) < CLng(sNumProxFaturaImpressao) Then
        
        'Verifica se a Faixa Final tambem não é menor que a que está no BD
        If CLng(FaturaFinal.Text) < CLng(sNumProxFaturaImpressao) Then
            lFaixaFinal = CLng(FaturaFinal.Text)
        Else
            lFaixaFinal = CLng(sNumProxFaturaImpressao) - 1
        End If
        
        'Avisa Reimpressao da Fatura e Pede Confirmação
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_FATURA_REIMPRESSA", CLng(FaturaInicial.Text), lFaixaFinal)
        If vbMsgRes = vbNo Then Error 61470
    
    End If
    
    'Altera a Flag para Imprimindo
    lErro = CF("CRFATConfig_Grava",CRFATCFG_FATURA_IMPRIMINDO, EMPRESA_TODA, RELATORIO_FATURA_IMPRIMINDO)
    If lErro <> SUCESSO Then Error 61471
        
    Me.Enabled = False
        
    lErro = gobjRelatorio.Executar_Prossegue
    If lErro <> SUCESSO And lErro <> 7072 Then Error 61472
    
    'Cancelou o relatório
    If lErro = 7072 Then Error 61473
    
    Timer1.Interval = INTERVALO_MONITORAMENTO_IMPRESSAO_FATURA

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 38230, 61468, 61469, 61470, 61471, 61472
        
        Case 61473
            Unload Me

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168453)

    End Select
    
    'Faz unlock da Tabela
    lErro = CF("CRFATConfig_Grava",CRFATCFG_FATURA_LOCKIMPRESSAO, EMPRESA_TODA, RELATORIO_FATURA_NAO_LOCKADO)
    
    Exit Sub

End Sub

Function Fatura_Lock_Impressao() As Long
    
Dim lErro As Long
Dim sLockImpressao As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Fatura_Lock_Impressao
    
    'Le para ver se esta lockado
    lErro = CF("CRFatConfig_Le",CRFATCFG_FATURA_LOCKIMPRESSAO, EMPRESA_TODA, sLockImpressao)
    If lErro <> SUCESSO Then Error 61474
    
    'Se está lockado Avisa
    If CLng(sLockImpressao) = RELATORIO_FATURA_LOCKADO Then
        
        'Avisa que a Impressão está bloqueada
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_FATURA_LOCKADA")
        If vbMsgRes = vbNo Then Error 61475
    
    End If
    
    sLockImpressao = CStr(RELATORIO_FATURA_LOCKADO)
    
    lErro = CF("CRFATConfig_Grava",CRFATCFG_FATURA_LOCKIMPRESSAO, EMPRESA_TODA, sLockImpressao)
    If lErro <> SUCESSO Then Error 61476
        
    Fatura_Lock_Impressao = SUCESSO
    
    Exit Function
    
Erro_Fatura_Lock_Impressao:

    Fatura_Lock_Impressao = Err
    
    Select Case Err
        
        Case 61474, 61475, 61476

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168454)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long


On Error GoTo Erro_Monta_Expressao_Selecao

   If Trim(FaturaInicial.Text) <> "" Then sExpressao = "Fatura >= " & Forprint_ConvLong(FaturaInicial.Text)

   If Trim(FaturaFinal.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Fatura <= " & Forprint_ConvLong(FaturaFinal.Text)

    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168455)

    End Select

    Exit Function

End Function

Private Sub FaturaInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FaturaInicial_Validate
    
    If Len(Trim(FaturaInicial.Text)) > 0 Then
        
        lErro = Long_Critica(FaturaInicial.Text)
        If lErro <> SUCESSO Then Error 38234
    
    End If
              
    Exit Sub

Erro_FaturaInicial_Validate:

    Cancel = True


    Select Case Err
    
        Case 38234
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168456)
            
    End Select
    
    Exit Sub

End Sub

Private Sub FaturaFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FaturaFinal_Validate
     
    If Len(Trim(FaturaFinal.Text)) > 0 Then
        
        lErro = Long_Critica(FaturaFinal.Text)
        If lErro <> SUCESSO Then Error 38235
        
    End If
        
    Exit Sub

Erro_FaturaFinal_Validate:

    Cancel = True


    Select Case Err
    
        Case 38235
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168457)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Timer1_Timer()

Dim lErro As Long
Dim lNumeroFaturasImpressas As Long
Dim sImprimindo As String

On Error GoTo Erro_Timer1_Timer

    glNumAcessosTimer = glNumAcessosTimer + 1
        
    lNumeroFaturasImpressas = CLng(FaturaFinal.Text) - CLng(FaturaInicial.Text)
    
    'Se não ultrapassou o tempo máximo de impressão
    If (glNumAcessosTimer * INTERVALO_MONITORAMENTO_IMPRESSAO_FATURA) <= (TEMPO_MAX_IMPRESSAO_UMA_FATURA * lNumeroFaturasImpressas) Then
    
        'Verifica se já Terminou a Impressão
        lErro = CF("CRFatConfig_Le",CRFATCFG_FATURA_IMPRIMINDO, EMPRESA_TODA, sImprimindo)
        If lErro <> SUCESSO Then Error 61477
       
        'Se terminou a Impressão
        If CInt(sImprimindo) = RELATORIO_FATURA_NAO_IMPRIMINDO Then
                   
           Timer1.Interval = 0
           
            'Chama a Tela de Controle de Impressão de Faturas
            Call Chama_Tela("RelOpControleImprFat", CLng(FaturaInicial.Text), CLng(FaturaFinal.Text))
          
            Unload Me
    
        End If
    
    Else
    
        'zera o timer
        Timer1.Interval = 0
        
        'Coloca iImprimindo = 0
        lErro = CF("CRFATConfig_Grava",CRFATCFG_FATURA_IMPRIMINDO, EMPRESA_TODA, 0)
        If lErro <> SUCESSO Then Error 61478
                
        'Chama a Tela de Controle de Impressão das Faturas
        Call Chama_Tela("RelOpControleImprFat", CLng(FaturaInicial.Text), CLng(FaturaFinal.Text))
        
        Unload Me
    
    End If
    
    Exit Sub
    
Erro_Timer1_Timer:
      
    Select Case Err
        
        Case 61477, 61478
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168458)
    
    End Select
    
    'Chama a Tela de Controle de Impressão das Faturas
    Call Chama_Tela("RelOpControleImprFat", CLng(FaturaInicial.Text), CLng(FaturaFinal.Text))
    
    Unload Me

    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_EMISSAO_FATURAS
    Set Form_Load_Ocx = Me
    Caption = "Emissão de Faturas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpEmissaoFatura"
    
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

Public Sub Unload(objme As Object)
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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

