VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl DAVEmitidosArquivo 
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   ScaleHeight     =   2310
   ScaleWidth      =   5025
   Begin VB.CommandButton BotaoDAVArquivo 
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
      Picture         =   "DAVEmitidosArquivo.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1545
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Caption         =   "Intervalo de Datas"
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
      Left            =   285
      TabIndex        =   0
      Top             =   135
      Width           =   4380
      Begin MSMask.MaskEdBox DataDe 
         Height          =   420
         Left            =   585
         TabIndex        =   1
         Top             =   495
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   435
         Left            =   4080
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   180
         _ExtentX        =   450
         _ExtentY        =   767
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   420
         Left            =   2790
         TabIndex        =   3
         Top             =   480
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   741
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   435
         Left            =   1920
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   495
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   767
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   6
         Top             =   525
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "At�:"
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
         Left            =   2265
         TabIndex        =   5
         Top             =   540
         Width           =   510
      End
   End
End
Attribute VB_Name = "DAVEmitidosArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub DataDe_GotFocus()
    'Trata A entrada em algum Campo

Dim lErro As Long

On Error GoTo Erro_DataDe_GotFocus
    
    'Fun��o que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataDe)

    Exit Sub

Erro_DataDe_GotFocus:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204842)

    End Select

    Exit Sub


End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataDe_Validate

    'Verifica se Data De esta Preenchida se n�o sai do Validate
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Fun��o que Serve para Verificar se a Data � Valida
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 204843

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 204843
            'Erro Tratado Dentro da Fun��o Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204844)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_GotFocus()
'Trata A entrada em algum Campo

Dim lErro As Long

On Error GoTo Erro_DataAte_GotFocus
    
    'Fun��o que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataAte)

    Exit Sub

Erro_DataAte_GotFocus:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204845)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataAte_Validate

    'Verifica se Data At� esta Preenchida se n�o sai do Validate
    If Len(Trim(DataAte.ClipText)) = 0 Then Exit Sub

    'Fun��o que Serve para Verificar se a Data � Valida
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 204846

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 204846
            'Erro Tratado Dentro da Fun��o Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204847)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDAVArquivo_Click()

Dim lErro As Long
Dim iTipoLeitura As Integer
Dim dtDataDe As Date
Dim dtDataAte As Date

On Error GoTo Erro_BotaoDAVArquivo_Click

    'Verificar se as Datas Est�o Preenchidas se Erro
    If Len(Trim(DataDe.ClipText)) = 0 Or Len(Trim(DataAte.ClipText)) = 0 Then gError 204848
    
    dtDataDe = DataDe.Text
    dtDataAte = DataAte.Text

    If dtDataDe > dtDataAte Then gError 204849

    
    'Fun��o que Vai Chamar Fun��o da Afrac que Vai Executar a Leitura da Memoria Fiscal
    lErro = CF_ECF("DAVEmitidos_Grava", dtDataDe, dtDataAte)
    If lErro <> SUCESSO Then gError 204850
    
    'Limpa a tela
    Call Limpa_Tela(Me)
    
    Exit Sub
    
Erro_BotaoDAVArquivo_Click:

    Select Case gErr

        Case 204848
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATAS_NAO_PREENCHIDAS, gErr)

        Case 204849
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_INICIAL_MAIOR1, gErr)

        Case 204850
            'Erro Tratado Dentro da Fun��o Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204851)

    End Select

    Exit Sub

End Sub


Private Sub UpDownDataDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataDe_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 204852
    
    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 204852

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204853)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 204854

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 204854

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204855)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAte_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 204856
    
    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 204856

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204857)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 204858

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 204858

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 204859)

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
    Caption = "DAV Emitidos Arquivo"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "DAVEmitidosArquivo"

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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210064)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

