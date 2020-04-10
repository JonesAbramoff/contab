VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ConsultaLoteNFeOcx 
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   KeyPreview      =   -1  'True
   ScaleHeight     =   1365
   ScaleWidth      =   3990
   Begin VB.CheckBox Scan 
      Caption         =   "Em Contingência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1095
      TabIndex        =   4
      Top             =   1035
      Width           =   2805
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   735
      Left            =   3330
      Picture         =   "ConsultaLoteNFeOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Fechar"
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton BotaoConsulta 
      Caption         =   "Consultar"
      Height          =   735
      Left            =   2355
      Picture         =   "ConsultaLoteNFeOcx.ctx":017E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Consultar"
      Top             =   120
      Width           =   825
   End
   Begin MSMask.MaskEdBox Lote 
      Height          =   315
      Left            =   1095
      TabIndex        =   0
      Top             =   585
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin VB.Label LoteLbl 
      AutoSize        =   -1  'True
      Caption         =   "Lote:"
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
      Left            =   585
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   630
      Width           =   450
   End
End
Attribute VB_Name = "ConsultaLoteNFeOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoLotes As AdmEvento
Attribute objEventoLotes.VB_VarHelpID = -1


Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)


End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iScan As Integer

On Error GoTo Erro_Form_Load
    
    Set objEventoLotes = New AdmEvento
    
    lErro = CF("NFeFedScan_Verifica_Contingencia", giFilialEmpresa, Date, iScan)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If iScan = MARCADO Then
        Scan.Value = vbChecked
        Scan.Caption = "Em Contingência - " & gobjCRFAT.sNFeSistemaContingencia
    Else
        Scan.Value = vbUnchecked
    End If
    
    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 199065)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoConsulta_Click()

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sDiretorio As String
Dim lRetorno As Long
Dim iScan As Integer
Dim iFilialEmpresa As Integer
Dim objVersao As New ClassVersaoNFe

On Error GoTo Erro_BotaoConsulta_Click

    iFilialEmpresa = giFilialEmpresa
    If giFilialEmpresa > 50 Then giFilialEmpresa = giFilialEmpresa - 50
    
    'verifica se o codigo foi preenchido
    If Len(Lote.Text) = 0 Then gError 203056

    objVersao.iCodigo = gobjCRFAT.iVersaoNFE
    
    lErro = CF("VersaoNFe_Le", objVersao)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207391

    sDiretorio = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
    sDiretorio = left(sDiretorio, lRetorno)

    iScan = IIf(Scan.Value = MARCADO, 1, -1)

    lErro = WinExec(sDiretorio & objVersao.sProgramaEnvio & " Consulta " & CStr(glEmpresa) & " " & CStr(giFilialEmpresa) & " " & Lote.Text & " " & CStr(iScan) & " " & IIf(iScan = MARCADO, gobjCRFAT.sNFeSistemaContingencia, ""), SW_NORMAL)

    Call Rotina_Aviso(vbOK, "AVISO_INICIO_CONSULTA_LOTE_NFE", Lote.Text)
    
    lErro = CF("NFE_Trata_Nota_Denegada")
    If lErro <> SUCESSO Then gError 207391

    Lote.Text = ""

    giFilialEmpresa = iFilialEmpresa
    
    Exit Sub
    
Erro_BotaoConsulta_Click:

    giFilialEmpresa = iFilialEmpresa
    
    Select Case gErr

        Case 203056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_PREENCHIDO", gErr)

        Case 207391

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203057)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Consulta de Lote de Envio de NFe"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConsultaLoteNFe"
    
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


Private Sub LoteLbl_Click()

Dim objNFeFedLoteView As New ClassNFeFedLoteView
Dim colSelecao As Collection

    If Len(Trim(Lote.Text)) > 0 Then
        objNFeFedLoteView.lLote = Lote.Text
    End If

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("NFeFedLoteViewLista", colSelecao, objNFeFedLoteView, objEventoLotes)

End Sub

Private Sub objEventoLotes_evSelecao(obj1 As Object)

Dim objNFeFedLoteView As ClassNFeFedLoteView
Dim bCancel As Boolean

    Set objNFeFedLoteView = obj1

    'Preenche o Cliente com o Cliente selecionado
    Lote.Text = objNFeFedLoteView.lLote

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
   ' Parent.UnloadDoFilho
    Set objEventoLotes = Nothing
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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Lote Then
            Call LoteLbl_Click
        End If
          
    End If

End Sub


