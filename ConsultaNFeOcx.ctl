VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ConsultaNFeOcx 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   ScaleHeight     =   1200
   ScaleWidth      =   7740
   Begin VB.CommandButton BotaoConsulta 
      Caption         =   "Consultar"
      Height          =   735
      Left            =   6015
      Picture         =   "ConsultaNFeOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Consultar"
      Top             =   180
      Width           =   825
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   735
      Left            =   6975
      Picture         =   "ConsultaNFeOcx.ctx":0ACE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Fechar"
      Top             =   180
      Width           =   480
   End
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
      Left            =   1530
      TabIndex        =   0
      Top             =   765
      Width           =   3840
   End
   Begin MSMask.MaskEdBox ChaveNFe 
      Height          =   315
      Left            =   1530
      TabIndex        =   3
      Top             =   315
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      MaxLength       =   44
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "############################################"
      PromptChar      =   " "
   End
   Begin VB.Label LabelChaveNFe 
      AutoSize        =   -1  'True
      Caption         =   "Chave da NFe:"
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
      Left            =   180
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   4
      Top             =   360
      Width           =   1290
   End
End
Attribute VB_Name = "ConsultaNFeOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoNFe As AdmEvento
Attribute objEventoNFe.VB_VarHelpID = -1


Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
'
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set objEventoNFe = Nothing
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iScan As Integer

On Error GoTo Erro_Form_Load

    Set objEventoNFe = New AdmEvento
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 207372)
    
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
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_BotaoConsulta_Click

    iFilialEmpresa = giFilialEmpresa
    If giFilialEmpresa > 50 Then giFilialEmpresa = giFilialEmpresa - 50
    
    'verifica se o codigo foi preenchido
    If Len(ChaveNFe.Text) = 0 Then gError 207373

    objFilialEmpresa.iCodFilial = giFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 201242
    
    If objFilialEmpresa.sCgc <> Mid(ChaveNFe.Text, 7, 14) Then gError 201242
    
    objVersao.iCodigo = gobjCRFAT.iVersaoNFE
    
    lErro = CF("VersaoNFe_Le", objVersao)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 207392
    
    sDiretorio = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
    sDiretorio = left(sDiretorio, lRetorno)

    iScan = IIf(Scan.Value = MARCADO, 1, -1)

    lErro = WinExec(sDiretorio & objVersao.sProgramaEnvio & " ConsultaNF " & CStr(glEmpresa) & " " & CStr(giFilialEmpresa) & " " & ChaveNFe.Text & " " & CStr(iScan) & " " & IIf(iScan = MARCADO, gobjCRFAT.sNFeSistemaContingencia, ""), SW_NORMAL)

    Call Rotina_Aviso(vbOK, "AVISO_INICIO_CONSULTA_NFE", ChaveNFe.Text)
    
    lErro = CF("NFE_Trata_Nota_Denegada")
    If lErro <> SUCESSO Then gError 207392
    
    
    ChaveNFe.PromptInclude = False
    ChaveNFe.Text = ""
    ChaveNFe.PromptInclude = True

    giFilialEmpresa = iFilialEmpresa
    
    Exit Sub
    
Erro_BotaoConsulta_Click:

    giFilialEmpresa = iFilialEmpresa
    
    Select Case gErr

        Case 207373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CHAVENFE_NAO_PREENCHIDA", gErr)

        Case 201242
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_VERIFIQUE_FILIALEMPRESA", gErr)
        
        Case 207392

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207374)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Consulta de Nota Fiscal Eletrônica"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConsultaNFe"
    
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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'
End Sub

Private Sub LabelChaveNFe_Click()

Dim colSelecao As New Collection
Dim sSelecao As String
Dim objNF As New ClassNFiscal

    objNF.sCodVerificacaoNFe = ChaveNFe.Text
    
    'chama o browser
    Call Chama_Tela("NFeChaveLista", colSelecao, objNF, objEventoNFe)
     
End Sub

Private Sub objEventoNFe_evSelecao(obj1 As Object)
'preenche a tela c/ os dados selecionados no browser

Dim objNF As ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_objEventoNFe_evSelecao

    Set objNF = obj1

    ChaveNFe.PromptInclude = False
    ChaveNFe.Text = objNF.sCodVerificacaoNFe
    ChaveNFe.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoNFe_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158649)
            
    End Select

    Exit Sub

End Sub
