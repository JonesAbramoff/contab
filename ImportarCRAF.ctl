VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ImportarCRAF 
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   LockControls    =   -1  'True
   ScaleHeight     =   2835
   ScaleWidth      =   5385
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   1680
      Picture         =   "ImportarCRAF.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2130
      Width           =   855
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   2610
      Picture         =   "ImportarCRAF.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2130
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação do Arquivo"
      Height          =   1815
      Left            =   105
      TabIndex        =   7
      Top             =   195
      Width           =   5175
      Begin VB.ComboBox Cobrador 
         Height          =   315
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   405
         Width           =   1920
      End
      Begin VB.TextBox NomeArquivo 
         Height          =   315
         Left            =   1695
         TabIndex        =   1
         Top             =   855
         Width           =   2955
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
         Left            =   4650
         TabIndex        =   2
         Top             =   840
         Width           =   360
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   2850
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1335
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Left            =   1695
         TabIndex        =   3
         Top             =   1320
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Movto Cta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   1380
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cobrador:"
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
         Left            =   825
         TabIndex        =   8
         Top             =   435
         Width           =   840
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
         Left            =   105
         TabIndex        =   9
         Top             =   900
         Width           =   1560
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4230
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ImportarCRAF"
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

    Parent.HelpContextID = IDH_PROCESSA_ARQRETCOBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Importação de Contas a Receber"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ImportarCRAF"

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
Dim objCobrador As New ClassCobrador
Dim sNomeArq As String
Dim sNomeDir As String
Dim iPOS As Integer
Dim iPosAnt As Integer
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_BotaoOK_Click

    'Verifica se o Cobrador foi selecionado
    If Cobrador.ListIndex = -1 Then gError 194087

    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 194088

    objCobrador.iCodigo = Codigo_Extrai(Cobrador.Text)

    'Lê os dados do cobrador
    lErro = CF("Cobrador_Le", objCobrador)
    If lErro <> SUCESSO And lErro <> 19294 Then gError 194089
    If lErro <> SUCESSO Then gError 194090
    
    If StrParaDate(Data.Text) = DATA_NULA Then
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_DATA_MOVCTA_NAO_PREENCHIDA_AF")
        If vbResult = vbNo Then gError 194089
    End If

    'NomeArquivo.Text
    iPOS = 1
    Do While iPOS <> 0
        iPosAnt = iPOS
        iPOS = InStr(iPosAnt + 1, NomeArquivo.Text, "\")
    Loop
    
    sNomeDir = Left(NomeArquivo.Text, iPosAnt)
    sNomeArq = Mid(NomeArquivo.Text, iPosAnt + 1)

    lErro = CF("Importa_CR_AF", giFilialEmpresa, objCobrador.iCodigo, sNomeDir, sNomeArq, StrParaDate(Data.Text))
    If lErro <> SUCESSO Then gError 194091

    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case 194087
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)
        
        Case 194088
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", gErr)
        
        Case 194089, 194091

        Case 194090
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", gErr, objCobrador.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194092)

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
    CommonDialog1.FilterIndex = 3
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
Dim ColCobrador As New Collection
Dim objCobrador As ClassCobrador

On Error GoTo Erro_Form_Load
    'Carrega a Coleção de Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then gError 194093

    For Each objCobrador In ColCobrador

        'Seleciona os cobradores ativos que utilizem cobrança eletrônica
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA And objCobrador.iInativo <> Inativo And objCobrador.iCobrancaEletronica = vbChecked Then
            Cobrador.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            Cobrador.ItemData(Cobrador.NewIndex) = objCobrador.iCodigo
        End If

    Next

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 194093

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194094)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function


Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'verifica se a data final é válida
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 194114

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 194114

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194115)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()
'diminui a data inicial

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 194116

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 194116

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194117)

    End Select

    Exit Sub


End Sub

Private Sub UpDownData_UpClick()
'aumenta a data inicial

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 194118

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 194118

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194119)

    End Select

    Exit Sub

End Sub
