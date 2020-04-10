VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl OCUsuArtlux 
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LockControls    =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   6000
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      Height          =   810
      Left            =   60
      TabIndex        =   18
      Top             =   0
      Width           =   5850
      Begin VB.TextBox Senha 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3870
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   285
         Width           =   1530
      End
      Begin VB.ComboBox Usuario 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   285
         Width           =   1650
      End
      Begin VB.Label Label1 
         Caption         =   "Usuário:"
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
         Height          =   360
         Index           =   0
         Left            =   75
         TabIndex        =   20
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Senha:"
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
         Height          =   360
         Left            =   3210
         TabIndex        =   19
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Início"
      Height          =   1380
      Left            =   60
      TabIndex        =   15
      Top             =   855
      Width           =   2805
      Begin VB.CommandButton BotaoIniciar 
         Caption         =   "Iniciar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         TabIndex        =   5
         Top             =   1005
         Width           =   2505
      End
      Begin MSComCtl2.UpDown UpDownDataIni 
         Height          =   300
         Left            =   1920
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataIni 
         Height          =   300
         Left            =   840
         TabIndex        =   2
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox HoraIni 
         Height          =   300
         Left            =   840
         TabIndex        =   4
         Top             =   600
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   17
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Index           =   905
         Left            =   330
         TabIndex        =   16
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Término"
      Height          =   1380
      Left            =   3105
      TabIndex        =   12
      Top             =   855
      Width           =   2805
      Begin VB.CommandButton BotaoFinalizar 
         Caption         =   "Finalizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         TabIndex        =   9
         Top             =   1005
         Width           =   2505
      End
      Begin MSComCtl2.UpDown UpDownDataFim 
         Height          =   300
         Left            =   1920
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   210
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFim 
         Height          =   300
         Left            =   840
         TabIndex        =   6
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox HoraFim 
         Height          =   300
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Index           =   1
         Left            =   330
         TabIndex        =   14
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   330
         TabIndex        =   13
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   2055
      Picture         =   "OCUsuArtlux.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   2985
      Picture         =   "OCUsuArtlux.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "OCUsuArtlux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()
 
Dim iAlterado As Integer
Dim gobjOC As ClassOCArtlux
Dim giEtapa As Integer

Private Sub BotaoCancela_Click()
    giRetornoTela = vbCancel
    Unload Me
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
   
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
        
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206691)
            
    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera as variaveis globais
    Set gobjOC = Nothing
    
End Sub

Public Function Trata_Parametros(ByVal objOC As ClassOCArtlux, iEtapa As Integer) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOCProd As ClassOCProdArtlux

On Error GoTo Erro_Trata_Parametros

    Set gobjOC = objOC
    giEtapa = iEtapa
        
    lErro = Usuarios_Carrega()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Traz_Tela(objOC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206692)
            
    End Select
    
    iAlterado = 0
    
    Exit Function
    
End Function

Public Function Traz_Tela(ByVal objOC As ClassOCArtlux) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim sUsu As String

On Error GoTo Erro_Traz_Tela
       
    If giEtapa = ETAPA_CORTE Then
        sUsu = objOC.sUsuCorte
        If objOC.dtDataIniCorte <> DATA_NULA Then
            DataIni.PromptInclude = False
            DataIni.Text = Format(objOC.dtDataIniCorte, "dd/mm/yy")
            DataIni.PromptInclude = True
            HoraIni.PromptInclude = False
            HoraIni.Text = Format(objOC.dHoraIniCorte, "hh:mm:ss")
            HoraIni.PromptInclude = True
        End If
        If objOC.dtDataFimCorte <> DATA_NULA Then
            DataFim.PromptInclude = False
            DataFim.Text = Format(objOC.dtDataFimCorte, "dd/mm/yy")
            DataFim.PromptInclude = True
            HoraFim.PromptInclude = False
            HoraFim.Text = Format(objOC.dHoraFimCorte, "hh:mm:ss")
            HoraFim.PromptInclude = True
        End If
    Else
        sUsu = objOC.sUsuForro
        If objOC.dtDataIniForro <> DATA_NULA Then
            DataIni.PromptInclude = False
            DataIni.Text = Format(objOC.dtDataIniForro, "dd/mm/yy")
            DataIni.PromptInclude = True
            HoraIni.PromptInclude = False
            HoraIni.Text = Format(objOC.dHoraIniForro, "hh:mm:ss")
            HoraIni.PromptInclude = True
        End If
        If objOC.dtDataFimForro <> DATA_NULA Then
            DataFim.PromptInclude = False
            DataFim.Text = Format(objOC.dtDataFimForro, "dd/mm/yy")
            DataFim.PromptInclude = True
            HoraFim.PromptInclude = False
            HoraFim.Text = Format(objOC.dHoraFimForro, "hh:mm:ss")
            HoraFim.PromptInclude = True
        End If
    End If
    
    For iIndice = 0 To Usuario.ListCount - 1
        If Usuario.List(iIndice) = sUsu Then
            Usuario.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Traz_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Tela:

    Traz_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206693)
            
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoFinalizar_Click()
    DataFim.PromptInclude = False
    DataFim.Text = Format(Date, "dd/mm/yy")
    DataFim.PromptInclude = True
    HoraFim.PromptInclude = False
    HoraFim.Text = Format(Time, "hh:mm:ss")
    HoraFim.PromptInclude = True
End Sub

Private Sub BotaoIniciar_Click()
    DataIni.PromptInclude = False
    DataIni.Text = Format(Date, "dd/mm/yy")
    DataIni.PromptInclude = True
    HoraIni.PromptInclude = False
    HoraIni.Text = Format(Time, "hh:mm:ss")
    HoraIni.PromptInclude = True
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Unload Me

    iAlterado = 0

    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206694)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objOC As New ClassOCArtlux
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Usuario.Text)) = 0 Then gError 206695
    If StrParaDate(DataIni.Text) = DATA_NULA Then gError 206696
    If StrParaDate(DataIni.Text) > StrParaDate(DataFim.Text) And StrParaDate(DataFim.Text) <> DATA_NULA Then gError 206697
    
    objUsuarios.sCodUsuario = Usuario.Text

    'Le o Usuario na tabela
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
    If lErro <> SUCESSO Then gError 206698

    If objUsuarios.sSenha <> Senha.Text Then gError 206699

    Call objOC.Copiar(gobjOC)
       
    lErro = Move_Tela_Memoria(objOC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
   
    If StrParaDate(DataIni.Text) = StrParaDate(DataFim.Text) Then
        If giEtapa = ETAPA_CORTE Then
            If objOC.dHoraIniCorte > objOC.dHoraFimCorte Then gError 206700
        Else
            If objOC.dHoraIniForro > objOC.dHoraIniForro Then gError 206700
        End If
    End If
    
    lErro = CF("OrdensDeCorteArtlux_Grava", objOC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    giRetornoTela = vbOK
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 206695
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", gErr)
    
        Case 206696
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INIC_NAO_PREENCHIDA", gErr)
            
        Case 206697
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
    
        Case 206698
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_ENCONTRADO", gErr)
    
        Case 206699
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_INVALIDA", gErr)
    
        Case 206700
            Call Rotina_Erro(vbOKOnly, "ERRO_HORAINI_MAIOR_HORAFIM", gErr)
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206701)
            
    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal objOC As ClassOCArtlux) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Move_Tela_Memoria
    
    If giEtapa = ETAPA_CORTE Then
        objOC.sUsuCorte = Usuario.Text
        objOC.dtDataIniCorte = StrParaDate(DataIni.Text)
        objOC.dtDataFimCorte = StrParaDate(DataFim.Text)
       
        If Len(Trim(HoraIni.ClipText)) > 0 Then
            objOC.dHoraIniCorte = CDate(HoraIni.Text)
        Else
            objOC.dHoraIniCorte = 0
        End If
        If Len(Trim(HoraFim.ClipText)) > 0 Then
            objOC.dHoraFimCorte = CDate(HoraFim.Text)
        Else
            objOC.dHoraFimCorte = 0
        End If
    Else
        objOC.sUsuForro = Usuario.Text
        objOC.dtDataIniForro = StrParaDate(DataIni.Text)
        objOC.dtDataFimForro = StrParaDate(DataFim.Text)
       
        If Len(Trim(HoraIni.ClipText)) > 0 Then
            objOC.dHoraIniForro = CDate(HoraIni.Text)
        Else
            objOC.dHoraIniForro = 0
        End If
        If Len(Trim(HoraFim.ClipText)) > 0 Then
            objOC.dHoraFimForro = CDate(HoraFim.Text)
        Else
            objOC.dHoraFimForro = 0
        End If
    End If
            
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206702)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LOCALIZACAO_PRODUTO
    Set Form_Load_Ocx = Me
    Caption = "Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OCUsuArtlux"
    
End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
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

Private Sub UpDownDataIni_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataIni_DownClick

    DataIni.SetFocus

    If Len(DataIni.ClipText) > 0 Then

        sData = DataIni.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataIni.Text = sData

    End If

    Exit Sub

Erro_UpDownDataIni_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206703)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataIni_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataIni_UpClick

    DataIni.SetFocus

    If Len(Trim(DataIni.ClipText)) > 0 Then

        sData = DataIni.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataIni.Text = sData

    End If

    Exit Sub

Erro_UpDownDataIni_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206704)

    End Select

    Exit Sub

End Sub

Private Sub DataIni_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataIni, iAlterado)
    
End Sub

Private Sub DataIni_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataIni_Validate

    If Len(Trim(DataIni.ClipText)) <> 0 Then

        lErro = Data_Critica(DataIni.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataIni_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206705)

    End Select

    Exit Sub

End Sub

Private Sub DataIni_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub


Private Sub UpDownDataFim_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFim_DownClick

    DataFim.SetFocus

    If Len(DataFim.ClipText) > 0 Then

        sData = DataFim.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataFim.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFim_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206706)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFim_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFim_UpClick

    DataFim.SetFocus

    If Len(Trim(DataFim.ClipText)) > 0 Then

        sData = DataFim.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataFim.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFim_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206707)

    End Select

    Exit Sub

End Sub

Private Sub DataFim_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFim, iAlterado)
    
End Sub

Private Sub DataFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFim_Validate

    If Len(Trim(DataFim.ClipText)) <> 0 Then

        lErro = Data_Critica(DataFim.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataFim_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206708)

    End Select

    Exit Sub

End Sub

Private Sub DataFim_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Usuarios_Carrega() As Long
'Carrega a ListBox

Dim lErro As Long
Dim objUsu As New ClassUsuProdArtlux
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios
Dim colUsu As New Collection

On Error GoTo Erro_Usuarios_Carrega

    'Le todos os Usuarios da Colecao
    lErro = CF("Usuarios_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Le todos os Compradores da Filial Empresa
    lErro = CF("UsuProdArtlux_Le_Todos", colUsu)
    If lErro <> SUCESSO And lErro <> 50126 Then gError ERRO_SEM_MENSAGEM

    For Each objUsu In colUsu
        For Each objUsuarios In colUsuarios
            If objUsu.sCodUsuario = objUsuarios.sCodUsuario Then
                If (objUsu.iAcessoCorte = MARCADO And giEtapa = ETAPA_CORTE) Or (objUsu.iAcessoForro = MARCADO And giEtapa = ETAPA_FORRO) Then
                    Usuario.AddItem objUsu.sCodUsuario
                End If
            End If
        Next
    Next

    Usuarios_Carrega = SUCESSO

    Exit Function

Erro_Usuarios_Carrega:

    Usuarios_Carrega = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206709)

    End Select

    Exit Function

End Function
