VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ConfiguraFISOcx 
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   ScaleHeight     =   1815
   ScaleWidth      =   5070
   Begin VB.Frame Frame3 
      Caption         =   "Não permite incluir, alterar ou excluir lançamentos anteriores a"
      Height          =   660
      Left            =   225
      TabIndex        =   5
      Top             =   915
      Width           =   4635
      Begin MSMask.MaskEdBox DataBloqLimite 
         Height          =   315
         Left            =   1230
         TabIndex        =   6
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataBloqLimite 
         Height          =   300
         Left            =   2370
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Data Limite:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   285
         Width           =   1065
      End
   End
   Begin MSMask.MaskEdBox CodFiscalServico 
      Height          =   345
      Left            =   2550
      TabIndex        =   0
      Top             =   270
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   609
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   5
      Mask            =   "#####"
      PromptChar      =   " "
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3570
      ScaleHeight     =   495
      ScaleWidth      =   1185
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   1245
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "configurafis.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   645
         Picture         =   "configurafis.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código Fiscal de Serviços:"
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
      Left            =   180
      TabIndex        =   4
      Top             =   330
      Width           =   2280
   End
End
Attribute VB_Name = "ConfiguraFISOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Usar constantes públicas que existem para as redes

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Configuração Geral"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConfiguraFIS"
    
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

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a funcao Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 80163

    'Envia mensagem de configuração gravada
    Call Rotina_Aviso(vbOKOnly, "AVISO_CONFIGURACAO_GRAVADA")

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 80163

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154694)

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

Public Function Trata_Parametros()

    Trata_Parametros = SUCESSO

    Exit Function

End Function

Function Traz_FisConfig_Tela(objFIS As ClassFIS) As Long
'Traz os dados de objFisConfig para tela

Dim lErro As Long

On Error GoTo Erro_Traz_FisConfig_Tela

    'Limpa o campo CodFiscalServico
    CodFiscalServico.Text = ""
    
    'Se o objFis estiver preenchido então escrever no campo
    If objFIS.iCodFiscalServico > 0 Then CodFiscalServico.Text = objFIS.iCodFiscalServico
    
    If gobjFIS.dtFisBloqDataLimite <> DATA_NULA Then Call DateParaMasked(DataBloqLimite, gobjFIS.dtFisBloqDataLimite)
    
    Exit Function

Erro_Traz_FisConfig_Tela:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154695)

    End Select

    Exit Function

End Function

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iAlterado = 0
        
    'Chama a a função responsável por trazer os dados na tela
    lErro = Traz_FisConfig_Tela(gobjFIS)
    If lErro <> SUCESSO Then gError 80095

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154696)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function ConfiguraFIS_Gravar(objFIS As ClassFIS) As Long

'Rotina que abre a transação para chamda a função que irá _
 Gravar os dados

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ConfiguraFIS_Gravar

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 80159

    'Chamada a função responsável pela gravação dos dados
    lErro = CF("ConfiguraFIS_GravarTrans", objFIS)
    If lErro <> SUCESSO Then gError 80162

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 80160
    
    ConfiguraFIS_Gravar = SUCESSO

    Exit Function

Erro_ConfiguraFIS_Gravar:

    ConfiguraFIS_Gravar = gErr

    Select Case gErr

        Case 80159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 80160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 80162

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154697)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function


Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFIS As New ClassFIS

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Carrega objFis com os dados da tela
    objFIS.iCodFiscalServico = StrParaInt(CodFiscalServico.Text)
    objFIS.dtFisBloqDataLimite = StrParaDate(DataBloqLimite.Text)
    
    'Chama a função responsável em validar a gravação na tabela
    lErro = ConfiguraFIS_Gravar(objFIS)
    If lErro <> SUCESSO And lErro <> 80165 Then gError 80166
    
    Call gobjFIS.Inicializa

    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 80165, 80166

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154698)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Sub UpDownDataBloqLimite_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataBloqLimite_DownClick

    DataBloqLimite.SetFocus

    If Len(DataBloqLimite.ClipText) > 0 Then

        sData = DataBloqLimite.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataBloqLimite.Text = sData

    End If

    Exit Sub

Erro_UpDownDataBloqLimite_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155647)

    End Select

    Exit Sub

End Sub

Sub UpDownDataBloqLimite_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataBloqLimite_UpClick

    DataBloqLimite.SetFocus

    If Len(Trim(DataBloqLimite.ClipText)) > 0 Then

        sData = DataBloqLimite.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataBloqLimite.Text = sData

    End If

    Exit Sub

Erro_UpDownDataBloqLimite_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155648)

    End Select

    Exit Sub

End Sub

Sub DataBloqLimite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataBloqLimite_Validate

    If Len(Trim(DataBloqLimite.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataBloqLimite.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Exit Sub

Erro_DataBloqLimite_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            'erro tratado na rotina chamada
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155649)

    End Select

    Exit Sub
    
End Sub

