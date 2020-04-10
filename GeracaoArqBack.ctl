VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl GeracaoArqBack 
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   ScaleHeight     =   3330
   ScaleWidth      =   3810
   Begin VB.ListBox Msgs 
      Height          =   840
      ItemData        =   "GeracaoArqBack.ctx":0000
      Left            =   150
      List            =   "GeracaoArqBack.ctx":0002
      TabIndex        =   5
      Top             =   2280
      Width           =   3420
   End
   Begin VB.Frame Frame2 
      Caption         =   "Arquivo"
      Height          =   1185
      Left            =   120
      TabIndex        =   0
      Top             =   465
      Width           =   3525
      Begin VB.CheckBox RecalcularTribProd 
         Caption         =   "Recalcular Tributação dos Produtos"
         Height          =   240
         Left            =   165
         TabIndex        =   6
         Top             =   840
         Width           =   2865
      End
      Begin VB.CommandButton BotaoGerar 
         Caption         =   "Gravação - Caixas"
         Height          =   345
         Left            =   150
         TabIndex        =   1
         Top             =   375
         Width           =   1635
      End
      Begin VB.CommandButton BotaoLer 
         Caption         =   "Leitura - Central "
         Height          =   345
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   315
      Left            =   2880
      Picture         =   "GeracaoArqBack.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Fechar"
      Top             =   120
      Width           =   780
   End
   Begin MSComctlLib.ProgressBar BarraProgresso 
      Height          =   345
      Left            =   135
      TabIndex        =   4
      Top             =   1815
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "GeracaoArqBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declarações Globais
Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Se o Caixa Central for integrado ao BackOffice
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL_BACKOFFICE Then
        'Desabilita a Leitura de arquivo do Caixa Central
        BotaoLer.Enabled = False
    End If
    
    BarraProgresso.Min = 0
    BarraProgresso.Max = 100
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160710)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long
Dim objBarraProgresso As Object
Dim objMsgs As Object


On Error GoTo Erro_BotaoGerar_Click
       
    Set objBarraProgresso = BarraProgresso
       
    Set objMsgs = Msgs
       
    lErro = CF("GeracaoArqBack_Grava", objBarraProgresso, objMsgs, RecalcularTribProd.Value = vbChecked)
    If lErro <> SUCESSO Then gError 126394
       
    Exit Sub
    
Erro_BotaoGerar_Click:

    Select Case gErr

        Case 126394

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160711)

    End Select

    Exit Sub

End Sub


Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Public Sub Trata_Parametros()
    
End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Geração Arquivos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GeracaoArqBack"
    
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


Private Sub BotaoLer_Click()
'vai ler os dados do caixa central para back
Dim sNomeArq As String
Dim lErro As Long
Dim objArq As New AdmCodigoNome
Dim objBarraProgresso As Object
Dim lIntervaloTrans As Long
Dim objObject As Object
Dim sNomeArqParam As String

On Error GoTo Erro_BotaoLer_Click

    Call Chama_Tela_Modal("ExibirArquivosCCBack", objArq)
    
    If giRetornoTela = vbOK Then
        
        sNomeArq = objArq.sNome
            
        Set objBarraProgresso = BarraProgresso
        
        lErro = CF("Rotina_Carga_CC_Back", sNomeArq, objBarraProgresso)
        If lErro <> SUCESSO Then gError 118925
        
        'avisa que a gravacao foi  concluida
        Call Rotina_Aviso(vbOKOnly, "AVISO_LEITURA_CONCLUIDA_COM_SUCESSO")
        
    End If
    
    Exit Sub
        
Erro_BotaoLer_Click:
    
   Select Case gErr

        Case 118925, 133597, 133598
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160712)

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
