VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl LMFSPorReducaoZ 
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ScaleHeight     =   2355
   ScaleWidth      =   4680
   Begin VB.CommandButton BotaoRelGer 
      Caption         =   "Rel. Gerencial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   150
      Picture         =   "LMFSPorReducaoZ.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1545
      Width           =   1935
   End
   Begin VB.CommandButton BotaoArquivo 
      Caption         =   "Arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Picture         =   "LMFSPorReducaoZ.ctx":3642
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1530
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame 
      Caption         =   "Intervalo de Reduções"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   4380
      Begin MSMask.MaskEdBox ReducaoDe 
         Height          =   420
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ReducaoAte 
         Height          =   420
         Left            =   3000
         TabIndex        =   2
         Top             =   480
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelReducaoDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   285
         TabIndex        =   4
         Top             =   540
         Width           =   435
      End
      Begin VB.Label LabelReducaoAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2400
         TabIndex        =   3
         Top             =   540
         Width           =   510
      End
   End
End
Attribute VB_Name = "LMFSPorReducaoZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub ReducaoDe_Validate(Cancel As Boolean)
'Valida os Dados no Intervalo de Redução

Dim lErro As Long

On Error GoTo Erro_ReducaoDe_Validate

    'Verifica se o Intervalo de redução não está preenchido sai do validate
    If Len(Trim(ReducaoDe.Text)) = 0 Then Exit Sub

    'Função que valida se Intervalode Redução é Positivo
    lErro = Valor_Positivo_Critica(ReducaoDe.Text)
    If lErro <> SUCESSO Then gError 204428

    Exit Sub

Erro_ReducaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 204428
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204429)

    End Select

    Exit Sub

End Sub

Private Sub ReducaoAte_Validate(Cancel As Boolean)
'Valida os Dados no Intervalo de Redução

Dim lErro As Long

On Error GoTo Erro_ReducaoAte_Validate

    'Verifica se o Intervalo de redução não está preenchido sai do validate
    If Len(Trim(ReducaoAte.Text)) = 0 Then Exit Sub

    'Função que valida se Intervalode Redução é Positivo
    lErro = Valor_Positivo_Critica(ReducaoAte.Text)
    If lErro <> SUCESSO Then gError 204430

    Exit Sub

Erro_ReducaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 204430
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204431)

    End Select

    Exit Sub

End Sub

Private Sub BotaoRelGer_Click()

Dim lErro As Long
Dim iTipoLeitura As Integer
Dim sDe As String
Dim sAte As String
Dim iTipo As Integer
Dim iArquivo As Integer

On Error GoTo Erro_BotaoRelGer_Click

    iTipo = LEITURA_SIMPLES
    iArquivo = 0
    iTipoLeitura = LEITURA_REDUCOES
    
    sDe = ReducaoDe.Text
    sAte = ReducaoAte.Text

    'Função que Vai Chamar Função da Afrac que Vai Executar a Leitura da Memoria Fiscal
    lErro = CF_ECF("MemoriaFiscal_Executa_Leitura", iTipoLeitura, sDe, sAte, iTipo, iArquivo)
    If lErro <> SUCESSO Then gError 204432
    
    'Limpa a tela
    Call Limpa_Tela(Me)
    
    Exit Sub
    
Erro_BotaoRelGer_Click:

    Select Case gErr

        Case 204432
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204433)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoArquivo_Click()

Dim lErro As Long
Dim iTipoLeitura As Integer
Dim sDe As String
Dim sAte As String
Dim iTipo As Integer
Dim iArquivo As Integer

On Error GoTo Erro_BotaoArquivo_Click

    iTipo = LEITURA_SIMPLES
    iArquivo = 1
    iTipoLeitura = LEITURA_REDUCOES
    
    sDe = ReducaoDe.Text
    sAte = ReducaoAte.Text

    'Função que Vai Chamar Função da Afrac que Vai Executar a Leitura da Memoria Fiscal
    lErro = CF_ECF("MemoriaFiscal_Executa_Leitura", iTipoLeitura, sDe, sAte, iTipo, iArquivo)
    If lErro <> SUCESSO Then gError 204434
    
    'Limpa a tela
    Call Limpa_Tela(Me)
    
    Exit Sub
    
Erro_BotaoArquivo_Click:

    Select Case gErr

        Case 204434
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204435)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Load()

    lErro_Chama_Tela = SUCESSO
    
    giRetornoTela = vbCancel

End Sub

Public Sub Form_Unload(Cancel As Integer)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Leitura MF Simples por Redução Z"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "LMFSPorReducaoZ"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim lErro As Long
    
On Error GoTo Erro_UserControl_KeyDown
    
    Select Case KeyCode
    
        Case vbKeyF8
'            Call BotaoFechar_Click
    
    End Select
    
    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210071)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub



