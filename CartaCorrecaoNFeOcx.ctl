VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CartaCorrecaoNFeOcx 
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   KeyPreview      =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   8070
   Begin VB.CheckBox Scan 
      Caption         =   "Scan"
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
      Left            =   6450
      TabIndex        =   14
      Top             =   735
      Width           =   885
   End
   Begin VB.CommandButton BotaoCartas 
      Caption         =   "Cartas Enviadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   210
      TabIndex        =   13
      Top             =   5010
      Width           =   1740
   End
   Begin VB.TextBox Correcao 
      Height          =   3285
      Left            =   240
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1560
      Width           =   7680
   End
   Begin VB.CommandButton BotaoLimparSeq 
      Height          =   300
      Left            =   1905
      Picture         =   "CartaCorrecaoNFeOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Limpar o Número"
      Top             =   690
      Width           =   345
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   6120
      ScaleHeight     =   465
      ScaleWidth      =   1635
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   15
      Width           =   1695
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   90
         Picture         =   "CartaCorrecaoNFeOcx.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   615
         Picture         =   "CartaCorrecaoNFeOcx.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1095
         Picture         =   "CartaCorrecaoNFeOcx.ctx":0BBE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox ChaveNFe 
      Height          =   315
      Left            =   1425
      TabIndex        =   0
      Top             =   210
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
   Begin VB.Label Sequencial 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1425
      TabIndex        =   12
      Top             =   705
      Width           =   465
   End
   Begin VB.Label Label4 
      Caption         =   "Correção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   285
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   1260
      Width           =   885
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3840
      TabIndex        =   9
      Top             =   675
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
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
      Left            =   3180
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label SequencialLabel 
      Caption         =   "Sequencial:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   7
      Top             =   750
      Width           =   1095
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
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      Top             =   255
      Width           =   1290
   End
End
Attribute VB_Name = "CartaCorrecaoNFeOcx"
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
Private WithEvents objEventoCCE As AdmEvento
Attribute objEventoCCE.VB_VarHelpID = -1

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
'
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set objEventoNFe = Nothing
    Set objEventoCCE = Nothing
End Sub

Private Sub BotaoCartas_Click()

Dim colSelecao As New Collection
Dim sSelecao As String
Dim objCCE As New ClassCartaCorrecao

    objCCE.schNFe = Trim(ChaveNFe.Text)
    
    'chama o browser
    Call Chama_Tela("NFeCceLista", colSelecao, objCCE, objEventoCCE)

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()
    
On Error GoTo Erro_Form_Load

    Set objEventoNFe = New AdmEvento
    Set objEventoCCE = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 207372)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros(Optional objCCE As ClassCartaCorrecao) As Long

    If Not (objCCE Is Nothing) Then
    
        Call Traz_CCE(objCCE)
        
    End If

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoGravar_Click()

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sDiretorio As String
Dim lRetorno As Long
Dim iScan As Integer
Dim iFilialEmpresa As Integer
Dim objVersao As New ClassVersaoNFe
Dim objCCE As New ClassCartaCorrecao
Dim objFilialEmpresa As New AdmFiliais


On Error GoTo Erro_BotaoGravar_Click

    'verifica se o codigo foi preenchido
    If Len(ChaveNFe.Text) = 0 Then gError 207373

    If Len(Trim(Correcao.Text)) < 15 Then gError 201241

    iFilialEmpresa = giFilialEmpresa
    If iFilialEmpresa > 50 Then iFilialEmpresa = iFilialEmpresa - 50
    
    objFilialEmpresa.iCodFilial = iFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 201242
    
    If objFilialEmpresa.sCgc <> Mid(ChaveNFe.Text, 7, 14) Then gError 201242
    
    With objCCE
        .iFilialEmpresa = iFilialEmpresa
        .schNFe = ChaveNFe.Text
        .inSeqEvento = StrParaInt(Sequencial.Caption)
        .sCorrecao = Trim(Correcao.Text)
    End With
    
    lErro = CF("CCE_Enviar", objCCE)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objVersao.iCodigo = gobjCRFAT.iVersaoNFE
    
    lErro = CF("VersaoNFe_Le", objVersao)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    sDiretorio = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
    sDiretorio = left(sDiretorio, lRetorno)

    iScan = IIf(Scan.Value = MARCADO, 1, -1)

    lErro = WinExec(sDiretorio & objVersao.sProgramaEnvio & " CartaCorrecao " & CStr(glEmpresa) & " " & CStr(iFilialEmpresa) & " " & CStr(objCCE.lidLote) & " " & CStr(iScan), SW_NORMAL)

    Call BotaoLimpar_Click
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr

        Case 207373
            Call Rotina_Erro(vbOKOnly, "ERRO_CHAVENFE_NAO_PREENCHIDA", gErr)

        Case 201241
            Call Rotina_Erro(vbOKOnly, "ERRO_CCE_TAM_MIN_CORRECAO", gErr)

        Case 201242
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_VERIFIQUE_FILIALEMPRESA", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207374)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Carta de Correção Eletrônica"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CartaCorrecaoNFe"
    
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

Private Sub BotaoLimpar_Click()
    
    ChaveNFe.PromptInclude = False
    ChaveNFe.Text = ""
    ChaveNFe.PromptInclude = True

    Sequencial.Caption = ""
    Status.Caption = ""
    Correcao.Text = ""

End Sub

Private Sub BotaoLimparSeq_Click()
    Sequencial.Caption = ""
    Status.Caption = ""
End Sub

Private Sub Traz_CCE(objCCE As ClassCartaCorrecao)

Dim lErro As Long

On Error GoTo Erro_Traz_CCE

    lErro = CF("CCE_Le_RetEnv", objCCE)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    If lErro = ERRO_LEITURA_SEM_DADOS Then
        Status.Caption = "NÂO PROCESSADA"
    Else
        Select Case objCCE.scStat
        
            Case "135", "136"
                Status.Caption = "HOMOLOGADA"
            
            Case Else
                Status.Caption = "REJEITADA"
            
        End Select
        
    End If
    
    ChaveNFe.PromptInclude = False
    ChaveNFe.Text = objCCE.schNFe
    ChaveNFe.PromptInclude = True

    Sequencial.Caption = CStr(objCCE.inSeqEvento)

    Correcao.Text = objCCE.sCorrecao

    Exit Sub
    
Erro_Traz_CCE:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158649)
            
    End Select

    Exit Sub

End Sub

Private Sub objEventoCCE_evSelecao(obj1 As Object)
'preenche a tela c/ os dados selecionados no browser

Dim objCCE As ClassCartaCorrecao
Dim lErro As Long

On Error GoTo Erro_objEventoCCE_evSelecao

    Set objCCE = obj1

    Call Traz_CCE(objCCE)
    
    Me.Show

    Exit Sub

Erro_objEventoCCE_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158649)
            
    End Select

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

