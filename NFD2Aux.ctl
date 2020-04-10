VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl NFD2Aux 
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8910
   ScaleHeight     =   1785
   ScaleWidth      =   8910
   Begin VB.Frame Frame1 
      Caption         =   "Destinatário"
      Height          =   885
      Left            =   0
      TabIndex        =   12
      Top             =   795
      Width           =   8760
      Begin VB.TextBox Destinatario 
         Height          =   540
         Left            =   120
         MaxLength       =   250
         TabIndex        =   13
         Top             =   255
         Width           =   8505
      End
   End
   Begin VB.Frame FrameIdentificacao 
      Caption         =   "Identificação"
      Height          =   705
      Left            =   195
      TabIndex        =   4
      Top             =   60
      Width           =   6420
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Left            =   2385
         TabIndex        =   5
         Top             =   255
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   4935
         TabIndex        =   6
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown SpinData 
         Height          =   315
         Left            =   6120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Serie 
         Height          =   315
         Left            =   750
         TabIndex        =   8
         Top             =   255
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         PromptChar      =   " "
      End
      Begin VB.Label LabelNum 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Left            =   1620
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   285
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
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
         Index           =   2
         Left            =   4020
         TabIndex        =   10
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Série:"
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
         Index           =   3
         Left            =   210
         TabIndex        =   9
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6810
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   165
      Width           =   1695
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "NFD2Aux.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "F5- Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   585
         Picture         =   "NFD2Aux.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "F7 - Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1095
         Picture         =   "NFD2Aux.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "F8 - Fechar"
         Top             =   75
         Width           =   420
      End
   End
End
Attribute VB_Name = "NFD2Aux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjVenda As ClassVenda

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Nota Fiscal - Modelo d2"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "NFD2Aux"
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

Public Sub Form_Load()
'Função inicialização da Tela
Dim lErro As Long

On Error GoTo Erro_Form_Load

    Call DateParaMasked(DataEmissao, Date)
    
    'Indica que o carregamento da tela aconteceu com sucesso
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        'Erros tratados na rotina chamada
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213430)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

End Sub

Private Sub BotaoFechar_Click()
    
    'Fechar a Tela
    Unload Me

End Sub


Function Trata_Parametros(ByVal objVenda As ClassVenda) As Long

    Set gobjVenda = objVenda

    If gobjVenda.objCupomFiscal.objNF.lNumNotaFiscal <> 0 Then
    
        Numero.Text = gobjVenda.objCupomFiscal.objNF.lNumNotaFiscal
        Serie.Text = gobjVenda.objCupomFiscal.objNF.sSerie
        
        DataEmissao.PromptInclude = False
        DataEmissao.Text = Format(gobjVenda.objCupomFiscal.objNF.dtDataEmissao, "dd/mm/yy")
        DataEmissao.PromptInclude = True
        
        Destinatario.Text = gobjVenda.objCupomFiscal.objNF.sDestino

    End If

    Trata_Parametros = SUCESSO

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lErro As Long

On Error GoTo Erro_UserControl_KeyDown

    Select Case KeyCode

        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoGravar) Then Exit Sub
            Call BotaoGravar_Click
            
        Case vbKeyF7
            If Not TrocaFoco(Me, BotaoLimpar) Then Exit Sub
            Call BotaoLimpar_Click
            
        Case vbKeyF8
            If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
            Call BotaoFechar_Click
            
    End Select

    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213434)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Função que efeuara a gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'    Call Limpa_Tela_NFD2
    
    'Fechar a Tela
    Unload Me
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213438)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objNFBD As New ClassNFiscal

On Error GoTo Erro_Gravar_Registro

    If Len(Trim(Numero.Text)) <> 0 Or Len(Trim(Serie.Text)) <> 0 Or Len(Trim(DataEmissao.ClipText)) <> 0 Then
    
        'Verifica os dados obrigatórios foram Preenchidos
        If Len(Trim(Numero.Text)) = 0 Then gError 213443
        If Len(Trim(Serie.Text)) = 0 Then gError 213444
        If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 213445
    
    End If
    
    lErro = Move_NFD2_Memoria(objNFBD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF_ECF("NFD2_Le", objNFBD)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If lErro <> ERRO_LEITURA_SEM_DADOS Then gError 213446
    
    'Guarda os dados na memoria que serão inseridos
    lErro = Move_NFD2_Memoria(gobjVenda.objCupomFiscal.objNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 213443
            Call Rotina_ErroECF(vbOKOnly, ERRO_NUMERO_NAO_PREENCHIDO, gErr)

        Case 213444
            Call Rotina_ErroECF(vbOKOnly, ERRO_SERIE_NAO_PREENCHIDA, gErr)
        
        Case 213445
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_NAO_PREENCHIDA1, gErr)
               
        Case 213446
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFD2_JA_EXISTENTE, gErr)
               
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213447)

    End Select

    Exit Function

End Function

Sub Limpa_Tela_NFD2()
'Função que Limpa a Tela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_NFD2

    'Limpa os Controles básico da Tela
    Call Limpa_Tela(Me)
    
    Exit Sub
    
Erro_Limpa_Tela_NFD2:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213448)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo BotaoLimpar_Click

    'Função que Lima a Tela
    Call Limpa_Tela_NFD2

    Exit Sub

BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213449)

    End Select

    Exit Sub

End Sub

Function Move_NFD2_Memoria(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long, iIndice As Integer
Dim objItem As ClassItemNF

On Error GoTo Erro_Move_NFD2_Memoria
    
    objNF.lNumNotaFiscal = StrParaLong(Numero.Text)
    objNF.sSerie = Trim(Serie.Text)
    objNF.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNF.sDestino = Destinatario.Text
    
    Move_NFD2_Memoria = SUCESSO
    
    Exit Function

Erro_Move_NFD2_Memoria:

    Move_NFD2_Memoria = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213450)

    End Select

    Exit Function

End Function

Private Sub SpinData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_DownClick

    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_SpinData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213456)

    End Select

    Exit Sub

End Sub

Private Sub SpinData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_UpClick

    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_SpinData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213457)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataEmissao_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_DataEmissao_Validate
    
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
        
    Exit Sub
    
Erro_DataEmissao_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213458)

    End Select

    Exit Sub
    
End Sub

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_GotFocus()
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)
End Sub

Private Sub Numero_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Numero_Validate
    
    If Len(Trim(Numero.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Numero.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
        
    Exit Sub
    
Erro_Numero_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213459)

    End Select

    Exit Sub
    
End Sub

Private Sub Serie_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Destinatario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

