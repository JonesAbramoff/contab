VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl CaixaECFConfig 
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ScaleHeight     =   2850
   ScaleWidth      =   4500
   Begin VB.Frame Frame1 
      Caption         =   "Balança"
      Height          =   1170
      Left            =   375
      TabIndex        =   7
      Top             =   1440
      Width           =   3885
      Begin VB.ComboBox BalancaModelo 
         Height          =   315
         ItemData        =   "CaixaECFConfig.ctx":0000
         Left            =   1665
         List            =   "CaixaECFConfig.ctx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   735
         Width           =   945
      End
      Begin MSMask.MaskEdBox BalancaPorta 
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   255
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   960
         TabIndex        =   9
         Top             =   795
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Porta:"
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
         Left            =   1110
         TabIndex        =   8
         Top             =   300
         Width           =   525
      End
   End
   Begin VB.CommandButton BotaoTEF 
      Caption         =   "TEF - Funções Administrativas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   405
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Gravar"
      Top             =   210
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3285
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1140
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "CaixaECFConfig.ctx":001F
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "CaixaECFConfig.ctx":019D
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   525
      Left            =   225
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   926
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"CaixaECFConfig.ctx":02F7
   End
End
Attribute VB_Name = "CaixaECFConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub Move_Tela_Memoria()

'    If OptionRemoverOrc.Value = True Then
'        giRemoveOrc = REMOVER_ORC
'    Else
'        giRemoveOrc = NAO_REMOVER_ORC
'    End If
    
    If Len(Trim(BalancaPorta.Text)) > 0 Then
        gsBalancaPorta = BalancaPorta.Text
    Else
        gsBalancaPorta = ""
    End If
    
    If BalancaModelo.ListIndex >= 0 Then
        giBalancaModelo = BalancaModelo.ItemData(BalancaModelo.ListIndex)
        gsBalancaNome = BalancaModelo.Text
    Else
        giBalancaModelo = 0
        gsBalancaNome = ""
    End If
    
End Sub

Public Function CaixaECFConfig_Grava() As Long

On Error GoTo Erro_CaixaECFConfig_Grava

    'grava na memória
    Call Move_Tela_Memoria

    'grava no arquivo
'    Call WritePrivateProfileString(APLICACAO_CAIXA, "RemoveOrcamento", CStr(giRemoveOrc), NOME_ARQUIVO_CAIXA)
    
    Call WritePrivateProfileString(APLICACAO_CAIXA, "BalancaPorta", gsBalancaPorta, NOME_ARQUIVO_CAIXA)
    
    Call WritePrivateProfileString(APLICACAO_CAIXA, "BalancaModelo", CStr(giBalancaModelo), NOME_ARQUIVO_CAIXA)
    Call WritePrivateProfileString(APLICACAO_CAIXA, "BalancaNome", CStr(gsBalancaNome), NOME_ARQUIVO_CAIXA)
    
    CaixaECFConfig_Grava = SUCESSO
    
    Exit Function
    
Erro_CaixaECFConfig_Grava:
    
    CaixaECFConfig_Grava = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144089)
    
    End Select
    
    Exit Function

End Function

Private Sub Inicializa_Terminal()

Dim lTamanho As Long
Dim sRetorno As String
Dim iPos As Integer
Dim sTel As String
Dim sRet As String
Dim sReg As String
Dim vbMsg As VbMsgBoxResult
Dim bOk As Boolean
Dim bOkVisa As Boolean
Dim lErro As Long
Dim sRetornoInicial As String
Dim lTamanhoInicial As Integer
Dim objMsg As Object
Dim objTela As Object

On Error GoTo Erro_Inicializa_Terminal
    
'    lTamanhoInicial = 70
'    sRetornoInicial = String(lTamanhoInicial, 0)
'
'    Call GetPrivateProfileString("DATA_TAB", "0202", CONSTANTE_ERRO, sRetornoInicial, lTamanhoInicial, "c:\visanet\Tabpos.cfg")
'
'    '****Inicialização do Terminal**********
'    bOk = False
'    bOkVisa = False
'
'    'Adiciona na variável global
'    glNumProxIdentificacao = glNumProxIdentificacao + 1
'
'    'Atualiza o arquivo
'    Call WritePrivateProfileString(APLICACAO_ECF, "NumProxIdent", CStr(glNumProxIdentificacao), NOME_ARQUIVO_CAIXA)
'
'    'Abre o arquivo Temporário
'    Open ARQUIVO_TEF_TEMP For Append As #1
'
'    'Informa dados de pagamento
'    Print #1, "000-000 = ADM"
'    Print #1, "001-000 = " & CStr(glNumProxIdentificacao)
'    Print #1, "999-999 = 0"
'
'    Close #1
'
'    'renomeando os arquivos
'    FileCopy ARQUIVO_TEF_TEMP, ARQUIVO_TEF_REQ
'
'    'Verifica o arquivo Ativo
'    sRet = Dir(ARQUIVO_TEF_ATIVO, vbNormal)
'    'TEF não está ativo
'    If sRet = "" Then gError 99760
'
'    Do
'        'Verifica o arquivo intpos.001
'        sRet = Dir(ARQUIVO_TEF_RESP2, vbNormal)
'        If sRet <> "" Then
'
'            'Abre o arquivo de resposta do gerenciador
'            Open ARQUIVO_TEF_RESP2 For Input As #1
'
'            'Até chegar ao fim do arquivo
'            Do While Not EOF(1)
'
'                'Busca o próximo registro do arquivo
'                Line Input #1, sReg
'
'                If Mid(sReg, 1, 7) = "001-000" Then
'                    If Right(sReg, Len(sReg) - 10) <> glNumProxIdentificacao Then
'                        sRet = ""
'                        Close #1
'                        Exit Do
'                    End If
'                End If
'
'                If Mid(sReg, 1, 7) = "010-000" Then
'                    If Mid(sReg, 11, Len(sReg) - 10) = "VISANET" Then bOkVisa = True
'                End If
'
'                If Mid(sReg, 1, 7) = "030-000" Then
'                    'mensagem para operador
'                    If Mid(sReg, 11, Len(sReg) - 10) <> "" Then vbMsg = Rotina_AvisoECF(vbOK, Mid(sReg, 11, Len(sReg) - 10))
'                    If bOkVisa Then
'                        If Mid(sReg, 11, Len(sReg) - 10) = "INICIACAO DO TERMINAL CONCLUIDA.." Then bOk = True
'                    End If
'                End If
'            Loop
'        End If
'    Loop Until sRet <> ""
'
'    Close #1
'
'    '******encerra inicialização*******
'
'    If bOkVisa Then
'        If bOk Then
'            lTamanho = 15
'            sRetorno = String(lTamanho, 0)
'
'            Call GetPrivateProfileString("DATA_TAB", "0202", CONSTANTE_ERRO, sRetorno, lTamanho, "c:\visanet\Tabpos.cfg")
'            iPos = InStr(1, sRetorno, " ")
'            sTel = Mid(sRetorno, 1, iPos - 1)
'
'            'Alterar o arquivo tabpos.cfg
'            Call WritePrivateProfileString("DATA_TAB", "0202", CStr(Operadora.Text & DDD.Text & sTel & "          " & Operadora.Text & DDD.Text & sTel & "          " & Operadora.Text & DDD.Text & sTel), "c:\visanet\Tabpos.cfg")
'
'        Else
'            vbMsg = Rotina_AvisoECF(vbOK, AVISO_NAO_INICIALIZADO_TERMINAL)
'
'            'Alterar o arquivo tabpos.cfg
'            Call WritePrivateProfileString("DATA_TAB", "0202", sRetornoInicial, "c:\visanet\Tabpos.cfg")
'        End If
'    End If
'
'    Exit Sub
    
Erro_Inicializa_Terminal:

    Select Case gErr
    
        Case 99760
            Call Rotina_ErroECF(vbOKOnly, ERRO_TEF_NAO_ATIVO, gErr)
            
        Case 133736
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144090)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    'grava as configurações no arquivo e na memória
    lErro = CaixaECFConfig_Grava()
    If lErro <> SUCESSO Then gError 109486

    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 109486
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144091)
    
    End Select

End Function

Private Sub BotaoTEF_Click()

Dim lErro As Long
Dim objMsg As Object
Dim objTela As Object
Dim objMsg1 As Object


On Error GoTo Erro_BotaoTEF_Click

'    If Len(Trim(Operadora.Text)) <> 0 Or Len(Trim(DDD.Text)) <> 0 Then
'        If Len(Trim(Operadora.Text)) = 0 Then gError 112397
'        If Len(Trim(DDD.Text)) = 0 Then gError 112398
'    End If
    
    Set objTela = Me
    Set objMsg = MsgTEF
    
    'Atualiza o arquivo(aberto e com Multiplo TEF)
    Call WritePrivateProfileString(APLICACAO_ECF, "COO", "0", NOME_ARQUIVO_CAIXA)
    
    lErro = CF_ECF("TEF_ADM_PAYGO", objMsg, objTela)
    If lErro <> SUCESSO Then gError 133736
    
'    Call Inicializa_Terminal
    
    Exit Sub
    
Erro_BotaoTEF_Click:
    
    Select Case gErr
    
        Case 112397
            Call Rotina_ErroECF(vbOKOnly, ERRO_OPERADORA_NAO_PREENCHIDA, gErr, giCodCaixa)
                    
        Case 112398
            Call Rotina_ErroECF(vbOKOnly, ERRO_DDD_NAO_PREENCHIDO, gErr, giCodCaixa)
                    
        Case 133736
                    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144092)
    
    End Select

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a gravar registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 109485
    
    'fecha a tela
    Unload Me

    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
    
        Case 109485
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144093)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    If giTEF = NAO_TEM_TEF Then BotaoTEF.Enabled = False
    
    BalancaPorta.Text = gsBalancaPorta
    
    For iIndice = 0 To BalancaModelo.ListCount - 1
        If BalancaModelo.ItemData(iIndice) = giBalancaModelo Then
            BalancaModelo.ListIndex = iIndice
            Exit For
        End If
    Next
        
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144094)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144095)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Configurações do Caixa"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CaixaECFConfig"

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

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'**** fim do trecho a ser copiado *****

Public Function objParent() As Object

    Set objParent = Parent
    
End Function

