VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpControleImprNFOcx 
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   LockControls    =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   4590
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
      Left            =   1613
      TabIndex        =   1
      Top             =   1650
      Width           =   1365
   End
   Begin VB.Frame FrameNF 
      Caption         =   "Boa Impressão - Notas Fiscais"
      Height          =   1305
      Left            =   188
      TabIndex        =   2
      Top             =   165
      Width           =   4215
      Begin VB.CheckBox TodaImpressaoRuim 
         Caption         =   "Não imprimiu bem todas as Notas Fiscais"
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
         TabIndex        =   6
         Top             =   810
         Width           =   3825
      End
      Begin MSMask.MaskEdBox NFiscalFinal 
         Height          =   300
         Left            =   2385
         TabIndex        =   0
         Top             =   397
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label NFiscalInicial 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   570
         TabIndex        =   5
         Top             =   397
         Width           =   960
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
         TabIndex        =   4
         Top             =   450
         Width           =   300
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
         TabIndex        =   3
         Top             =   450
         Width           =   360
      End
   End
End
Attribute VB_Name = "RelOpControleImprNFOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim giFinalizando As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()
Dim gobjSerie As ClassSerie
Dim glNumeroNFAte As Long

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giFinalizando = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167900)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    If giFinalizando = 0 Then
        Call BotaoOK_Click
    End If
    
    Set gobjSerie = Nothing
    
    Unload Me
    
End Sub


Function Trata_Parametros(objSerie As ClassSerie, lNumeroNFDe As Long, lNumeroNFAte As Long) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    Set gobjSerie = objSerie
    'Guarda número original Até (vindo da outra tela)
    glNumeroNFAte = lNumeroNFAte
        
    NFiscalInicial.Caption = lNumeroNFDe
    NFiscalFinal.Text = lNumeroNFAte
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167901)

    End Select

    Exit Function

End Function

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim objSerie As New ClassSerie

On Error GoTo Erro_BotaoOK_Click
    
    'se a Checkbox TodaImpressaoRuim está desmarcada
    If TodaImpressaoRuim.Value = False Then
        
        'Então a Nota Fiscal Final tem que estar preenchida
        If Len(Trim(NFiscalFinal.Text)) = 0 Then Error 60395
        
        objSerie.sSerie = gobjSerie.sSerie
        objSerie.lProxNumNFiscalImpressa = CLng(NFiscalFinal.Text) + 1
        
        'UnLock e Atualiza a Tabela de Série
        lErro = CF("Serie_Unlock_Atualiza_ImpressaoNF",objSerie)
        If lErro <> SUCESSO And lErro <> 60409 Then Error 61013
        
        'Não encontrou a Série
        If lErro = 60409 Then Error 61036
        
    Else
        
        objSerie.sSerie = gobjSerie.sSerie
        
        'Faz Unlock da Tabela
        lErro = CF("Serie_Unlock_ImpressaoNF",objSerie)
        If lErro <> SUCESSO And lErro <> 60402 Then Error 61014
        
        If lErro = 60402 Then Error 61037
    
    End If
    
    giFinalizando = 1
    
    Unload Me
    
    Exit Sub
        
Erro_BotaoOK_Click:

    Select Case Err
        
        Case 61013, 61014
        
        Case 60395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_IMPRESSAO_NAO_PREENCHIDO", Err)
        
        Case 61036, 61037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, objSerie.sSerie)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167902)

    End Select

    Exit Sub

End Sub

Private Sub TodaImpressaoRuim_Click()

    If TodaImpressaoRuim.Value = vbChecked Then
        NFiscalFinal.Enabled = False
    Else
        NFiscalFinal.Enabled = True
    End If
    
End Sub

Private Sub NFiscalFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NFiscalFinal_Validate
     
    If Len(Trim(NFiscalFinal.Text)) > 0 Then
        
        'Critica se é um long
        lErro = Long_Critica(NFiscalFinal.Text)
        If lErro <> SUCESSO Then Error 60396
        
    End If
    
    'Verifica se é maior que o Numero da Tela Mãe se for --> ERRO
    If Len(Trim(NFiscalFinal.Text)) > 0 Then
        If CLng(NFiscalFinal.Text) > glNumeroNFAte Then Error 60397
    End If
            
    'Verifica se é menor que o Numero De. Se for --> ERRO
    If Len(Trim(NFiscalFinal.Text)) > 0 Then
        If CLng(NFiscalFinal.Text) < CLng(NFiscalInicial.Caption) Then Error 61035
    End If
    
    Exit Sub

Erro_NFiscalFinal_Validate:

    Cancel = True


    Select Case Err
    
        Case 60396
            
        Case 60397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_MAIOR_ANTERIOR", Err, Error$)
        
        Case 61035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_MENOR_NUMERO_DE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167903)
            
    End Select
    
    Exit Sub

End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CONTROLE_IMPRESSAO_NF
    Set Form_Load_Ocx = Me
    Caption = "Controle de Impressão de Notas Fiscais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpControleNF"
    
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

Private Sub NFiscalInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NFiscalInicial, Source, X, Y)
End Sub

Private Sub NFiscalInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NFiscalInicial, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

