VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpControleImprFatOcx 
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   ScaleHeight     =   2250
   ScaleWidth      =   4590
   Begin VB.Frame FrameNF 
      Caption         =   "Boa Impressão - Faturas a Receber"
      Height          =   1305
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   4215
      Begin VB.CheckBox TodaImpressaoRuim 
         Caption         =   "Não imprimiu bem todas as Faturas"
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
         TabIndex        =   2
         Top             =   810
         Width           =   3825
      End
      Begin MSMask.MaskEdBox FaturaFinal 
         Height          =   300
         Left            =   2385
         TabIndex        =   3
         Top             =   397
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1950
         TabIndex        =   6
         Top             =   450
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   450
         Width           =   300
      End
      Begin VB.Label FaturaInicial 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   570
         TabIndex        =   4
         Top             =   397
         Width           =   960
      End
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1605
      TabIndex        =   0
      Top             =   1665
      Width           =   1365
   End
End
Attribute VB_Name = "RelOpControleImprFatOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim giFinalizando As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()
Dim glFaturaAte As Long

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167896)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
  
    If giFinalizando = 0 Then
        Call BotaoOK_Click
    End If
    
    Unload Me
    
End Sub


Function Trata_Parametros(lFaturaDe As Long, lFaturaAte As Long) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    'Guarda número original Até (vindo da outra tela)
    glFaturaAte = lFaturaAte
        
    FaturaInicial.Caption = lFaturaDe
    FaturaFinal.Text = lFaturaAte
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167897)

    End Select

    Exit Function

End Function

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim lProxNumNFiscalImpressa As Long
Dim colConfigs As New ColCRFATConfig

On Error GoTo Erro_BotaoOK_Click
    
    'se a Checkbox TodaImpressaoRuim está desmarcada
    If TodaImpressaoRuim.Value = False Then
        
        'Então a Fatura Final tem que estar preenchida
        If Len(Trim(FaturaFinal.Text)) = 0 Then Error 61479
        
        lProxNumNFiscalImpressa = CLng(FaturaFinal.Text) + 1
        
        Call colConfigs.Add(CRFATCFG_FATURA_LOCKIMPRESSAO, EMPRESA_TODA, "", 0, CStr(RELATORIO_FATURA_LOCKADO))
        Call colConfigs.Add(CRFATCFG_FATURA_NUM_PROX_IMPRESSAO, EMPRESA_TODA, "", 0, CStr(lProxNumNFiscalImpressa))
        
        'UnLock e Atualiza a Tabela CrfatConfig
        lErro = CF("CRFATConfig_Grava_Configs",colConfigs)
        If lErro <> SUCESSO Then Error 61480
        
    Else
        
        'Faz Unlock da Tabela crfat Config
        lErro = CF("CRFATConfig_Grava",CRFATCFG_FATURA_LOCKIMPRESSAO, EMPRESA_TODA, RELATORIO_FATURA_NAO_LOCKADO)
        If lErro <> SUCESSO Then Error 61481
    
    End If
    
    giFinalizando = 1
    
    Unload Me
    
    Exit Sub
        
Erro_BotaoOK_Click:

    Select Case Err
        
        Case 61479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURA_ATE_IMPRESSAO_NAO_PREENCHIDO", Err)
        
        Case 61480, 61481
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167898)

    End Select

    Exit Sub

End Sub

Private Sub TodaImpressaoRuim_Click()

    If TodaImpressaoRuim.Value = vbChecked Then
        FaturaFinal.Enabled = False
    Else
        FaturaFinal.Enabled = True
    End If
    
End Sub

Private Sub FaturaFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FaturaFinal_Validate
     
    If Len(Trim(FaturaFinal.Text)) > 0 Then
        
        'Critica se é um long
        lErro = Long_Critica(FaturaFinal.Text)
        If lErro <> SUCESSO Then Error 61482
        
    End If
    
    'Verifica se é maior que o Numero da Tela Mãe se for --> ERRO
    If Len(Trim(FaturaFinal.Text)) > 0 Then
        If CLng(FaturaFinal.Text) > glFaturaAte Then Error 61483
    End If
            
    'Verifica se é menor que o Numero De. Se for --> ERRO
    If Len(Trim(FaturaFinal.Text)) > 0 Then
        If CLng(FaturaFinal.Text) < CLng(FaturaInicial.Caption) Then Error 61484
    End If
    
    Exit Sub

Erro_FaturaFinal_Validate:

    Cancel = True


    Select Case Err
    
        Case 61482
            
        Case 61483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURA_ATE_MAIOR_ANTERIOR", Err, Error$)
        
        Case 61484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURA_ATE_MENOR_NUMERO_DE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167899)
            
    End Select
    
    Exit Sub

End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CONTROLE_IMPRESSAO_FATURAS
    Set Form_Load_Ocx = Me
    Caption = "Controle de Impressão das Faturas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpControleImprFat"
    
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

Private Sub FaturaInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FaturaInicial, Source, X, Y)
End Sub

Private Sub FaturaInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FaturaInicial, Button, Shift, X, Y)
End Sub


Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
   Height = UserControl.Height
End Property

Public Property Get Width() As Long
   Width = UserControl.Width
End Property
