VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl FlCxCtb1Ocx 
   ClientHeight    =   1272
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3804
   ScaleHeight     =   1272
   ScaleWidth      =   3804
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   396
      Picture         =   "FlCxCtb1Ocx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   624
      Width           =   855
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   2292
      Picture         =   "FlCxCtb1Ocx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   648
      Width           =   855
   End
   Begin MSMask.MaskEdBox MesInicial 
      Height          =   312
      Left            =   864
      TabIndex        =   0
      Top             =   168
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   550
      _Version        =   393216
      MaxLength       =   2
      Mask            =   "##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox AnoInicial 
      Height          =   312
      Left            =   2544
      TabIndex        =   1
      Top             =   168
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   550
      _Version        =   393216
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label PIni 
      Caption         =   "Mês:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   216
      Width           =   492
   End
   Begin VB.Label PFim 
      Caption         =   "Ano:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   252
      Left            =   2064
      TabIndex        =   2
      Top             =   228
      Width           =   456
   End
End
Attribute VB_Name = "FlCxCtb1Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAlterado As Integer


'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Form_Load()
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
End Sub

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoCancela_Click()

    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim dtDataIni As Date
Dim dtDataFim As Date
Dim objPlanilha As New ClassPlanilhaExcel
Dim lNumIntRel As Long

On Error GoTo Erro_BotaoOK_Click

    'Exibe o ponteiro ampulheta
    MousePointer = vbHourglass
    
    'Se o mes não foi preenchida => erro
    If Len(Trim(MesInicial.ClipText)) = 0 Then gError 199209
        
    'Se o ano não foi preenchida => erro
    If Len(Trim(AnoInicial.ClipText)) = 0 Then gError 199210
      
    'Guarda as datas de início e fim do período que servirá de base para o gráfico
    dtDataIni = StrParaDate("01/" & MesInicial.Text & "/" & AnoInicial.Text)
    dtDataFim = DateAdd("m", 1, dtDataIni) - 1
    
    lNumIntRel = 2
    
    'Obtém os dados que serão utilizados para gerar a planilha que servirá de base ao gráfico
    'Chama a função que montará o gráfico no excel
    lErro = CF("RelFlCxCtb_Prepara", giFilialEmpresa, dtDataIni, dtDataFim, lNumIntRel)
    If lErro <> SUCESSO Then gError 199211
    
    lErro = CF("RelFlCxCtb_Move_Dados_Excel", objPlanilha, 2, StrParaInt(MesInicial.Text), StrParaInt(AnoInicial.Text))
    If lErro <> SUCESSO Then gError 199212
    
    lErro = CF("Excel_Gera_Planilha_Fluxo", objPlanilha)
    If lErro <> SUCESSO Then gError 199216
    
    'Exibe o ponteiro padrão
    MousePointer = vbDefault
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:
    
    Select Case gErr

        Case 199209
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREENCHIDO", gErr)
        
        Case 199210
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
        
        Case 199211, 199212, 199216
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199213)
            
    End Select
    
    'Exibe o ponteiro padrão
    MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub MesInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MesInicial_Validate

    If Len(MesInicial.ClipText) > 0 Then
        
        If StrParaInt(MesInicial.Text) > 12 Then gError 199205
        
    End If

    Exit Sub

Erro_MesInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 199205
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_INVALIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199206)

    End Select

    Exit Sub

End Sub

Private Sub AnoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AnoInicial_Validate

    If Len(AnoInicial.ClipText) > 0 Then
        
        If Len(AnoInicial.Text) < 4 Then gError 199207
        
    End If

    Exit Sub

Erro_AnoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 199207
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_ANO_INVALIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199208)

    End Select

    Exit Sub

End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa Contabil I"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FlCxCtb1"
    
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

Private Sub PIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PIni, Source, X, Y)
End Sub

Private Sub PIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PIni, Button, Shift, X, Y)
End Sub

Private Sub PFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PFim, Source, X, Y)
End Sub

Private Sub PFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PFim, Button, Shift, X, Y)
End Sub


