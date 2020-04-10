VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl ImportarNFeOcx 
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   ScaleHeight     =   2340
   ScaleWidth      =   6975
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   2595
      Picture         =   "ImportarNFeOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1590
      Width           =   855
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   3525
      Picture         =   "ImportarNFeOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1590
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação do Arquivo"
      Height          =   1260
      Left            =   105
      TabIndex        =   0
      Top             =   195
      Width           =   6765
      Begin VB.TextBox NomeArquivo 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Top             =   510
         Width           =   4425
      End
      Begin VB.CommandButton BotaoProcurar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6180
         TabIndex        =   1
         Top             =   465
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Arquivo:"
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
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   540
         Width           =   1560
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   1590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ImportarNFeOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_PROCESSA_ARQRETCOBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Importar Notas Fiscais Eletrônicas Municipais"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ImportarNFe"

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

Private Sub BotaoCancela_Click()
    Unload Me
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim sNomeArq As String
Dim sNomeDir As String
Dim iPos As Integer
Dim iPosAnt As Integer

On Error GoTo Erro_BotaoOK_Click

    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 192919
    
    'NomeArquivo.Text
    iPos = 1
    Do While iPos <> 0
        iPosAnt = iPos
        iPos = InStr(iPosAnt + 1, NomeArquivo.Text, "\")
    Loop
    
    sNomeDir = left(NomeArquivo.Text, iPosAnt)
    sNomeArq = Mid(NomeArquivo.Text, iPosAnt + 1)
    
    lErro = CF("NFe_Importa", sNomeDir, sNomeArq)
    If lErro <> SUCESSO Then gError 192920

    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr
        
        Case 192919
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", gErr)
            
        Case 192920

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192921)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProcurar_Click()

    On Error GoTo Erro_BotaoProcurar_Click

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen

    ' Display name of selected file

    NomeArquivo.Text = CommonDialog1.FileName
    Exit Sub

Erro_BotaoProcurar_Click:
    'User pressed the Cancel button
    Exit Sub

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
'***** fim do trecho a ser copiado ******


Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 192922)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

