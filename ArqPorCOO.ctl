VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ArqPorCOO 
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   ScaleHeight     =   2265
   ScaleWidth      =   4740
   Begin VB.Frame Frame 
      Caption         =   "Intervalo de COO"
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
      Left            =   150
      TabIndex        =   1
      Top             =   75
      Width           =   4380
      Begin MSMask.MaskEdBox COODe 
         Height          =   420
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox COOAte 
         Height          =   420
         Left            =   3000
         TabIndex        =   3
         Top             =   480
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "######"
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   540
         Width           =   510
      End
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
      Left            =   1335
      Picture         =   "ArqPorCOO.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1530
      Width           =   1935
   End
End
Attribute VB_Name = "ArqPorCOO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub COODe_Validate(Cancel As Boolean)
'Valida os Dados no Intervalo de COO

Dim lErro As Long

On Error GoTo Erro_COODe_Validate

    'Verifica se o Intervalo de COO não está preenchido sai do validate
    If Len(Trim(COODe.Text)) = 0 Then Exit Sub

    'Função que valida se Intervalode Redução é Positivo
    lErro = Valor_Positivo_Critica(COODe.Text)
    If lErro <> SUCESSO Then gError 204490

    Exit Sub

Erro_COODe_Validate:

    Cancel = True

    Select Case gErr

        Case 204490
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204491)

    End Select

    Exit Sub

End Sub

Private Sub COOAte_Validate(Cancel As Boolean)
'Valida os Dados no Intervalo de Redução

Dim lErro As Long

On Error GoTo Erro_COOAte_Validate

    'Verifica se o Intervalo de redução não está preenchido sai do validate
    If Len(Trim(COOAte.Text)) = 0 Then Exit Sub

    'Função que valida se Intervalode Redução é Positivo
    lErro = Valor_Positivo_Critica(COOAte.Text)
    If lErro <> SUCESSO Then gError 204492

    Exit Sub

Erro_COOAte_Validate:

    Cancel = True

    Select Case gErr

        Case 204492
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204493)

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

    iTipoLeitura = LEITURA_COO
    
    sDe = COODe.Text
    sAte = COOAte.Text

    'Função que Vai Chamar Função da Afrac que Vai Executar a Leitura da Memoria Fiscal
    lErro = CF_ECF("ArqMFD_Executa", iTipoLeitura, sDe, sAte)
    If lErro <> SUCESSO Then gError 204494
    
    'Limpa a tela
    Call Limpa_Tela(Me)
    
    Exit Sub
    
Erro_BotaoArquivo_Click:

    Select Case gErr

        Case 204494
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204495)

    End Select

    Exit Sub
    
End Sub

Private Sub COOAte_GotFocus()
'Trata A entrada em algum Campo

Dim lErro As Long

On Error GoTo Erro_COOAte_GotFocus
    
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(COOAte)

    Exit Sub

Erro_COOAte_GotFocus:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204496)

    End Select

    Exit Sub

End Sub

Private Sub COODe_GotFocus()
'Trata A entrada em algum Campo

Dim lErro As Long

On Error GoTo Erro_COODe_GotFocus
    
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(COODe)

    Exit Sub

Erro_COODe_GotFocus:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204497)

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
    Caption = "Arquivo MFD por COO"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ArqPorCOO"

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
    
'        Case vbKeyF8
'            Call BotaoFechar_Click
    
    End Select
    
    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210062)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

