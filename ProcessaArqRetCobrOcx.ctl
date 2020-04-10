VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl ProcessaArqRetCobrOcx 
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   ScaleHeight     =   2430
   ScaleWidth      =   5385
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   1680
      Picture         =   "ProcessaArqRetCobrOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1770
      Width           =   855
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   2610
      Picture         =   "ProcessaArqRetCobrOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1770
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação do Arquivo"
      Height          =   1395
      Left            =   105
      TabIndex        =   0
      Top             =   195
      Width           =   5175
      Begin VB.ComboBox Cobrador 
         Height          =   315
         Left            =   1695
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   405
         Width           =   1920
      End
      Begin VB.TextBox NomeArquivo 
         Height          =   315
         Left            =   1695
         TabIndex        =   2
         Top             =   870
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
         TabIndex        =   1
         Top             =   840
         Width           =   360
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
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   900
         Width           =   1560
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4230
      Top             =   1755
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ProcessaArqRetCobrOcx"
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
    Caption = "Processamento de Arquivo de Retorno"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ProcessaArqRetCobr"

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
Dim sNomeArqParam As String
Dim objCobrancaEletronica As New ClassCobrancaEletronica

On Error GoTo Erro_BotaoOK_Click

    'Verifica se o Cobrador foi selecionado
    If Cobrador.ListIndex = -1 Then Error 51641

    If Len(Trim(NomeArquivo.Text)) = 0 Then Error 62295

    objCobrancaEletronica.objCobrador.iCodigo = Codigo_Extrai(Cobrador.Text)

    'Lê os dados do cobrador
    lErro = CF("Cobrador_Le", objCobrancaEletronica.objCobrador)
    If lErro <> SUCESSO And lErro <> 19294 Then Error 51642
    If lErro <> SUCESSO Then Error 51643

    objCobrancaEletronica.sNomeArquivoRetorno = NomeArquivo.Text
    objCobrancaEletronica.objCobradorCNABInfo.iCodCobrador = objCobrancaEletronica.objCobrador.iCodigo
    lErro = CF("CobradorInfo_Le", objCobrancaEletronica.objCobrador.iCodigo, objCobrancaEletronica.objCobradorCNABInfo.colInformacoes)
    If lErro <> SUCESSO Then Error 32272
    
''    lErro = Sistema_Preparar_Batch(sNomeArqParam)
''    If lErro <> SUCESSO Then Error 62287

    lErro = CF("CobrancaEletronica_Abre_TelaRetornoArq", sNomeArqParam, objCobrancaEletronica)
    If lErro <> SUCESSO Then Error 62286

    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case Err

        Case 51641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", Err)

        Case 51642, 51644, 62286, 62287, 62293, 62294, 32272

        Case 51643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", Err, objCobrancaEletronica.objCobrador.iCodigo)

        Case 62295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165298)

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
    "(*.txt)|*.txt|Ret Files (*.ret)|*.ret"
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
    If lErro <> SUCESSO Then Error 51649

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

    lErro_Chama_Tela = Err

    Select Case Err

        Case 51649

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165299)

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

