VERSION 5.00
Begin VB.UserControl ExibirArquivos 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   ScaleHeight     =   5790
   ScaleWidth      =   7230
   Begin VB.CommandButton BotaoFTP 
      Caption         =   "Download Arquivos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3855
      TabIndex        =   6
      Top             =   180
      Width           =   1950
   End
   Begin VB.FileListBox File1 
      Height          =   4380
      Left            =   3720
      Pattern         =   "*.ccc*"
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      Height          =   4365
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5955
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1140
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "ExibirArquivos.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "ExibirArquivos.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label DirName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   6855
   End
End
Attribute VB_Name = "ExibirArquivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAlterado As Integer
Dim gobjArq As New AdmCodigoNome

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub BotaoFTP_Click()

Dim sNomeArqParam As String
Dim lIntervaloTrans As Long
Dim lErro As Long
Dim objObject As Object

On Error GoTo Erro_BotaoFTP_Click

     If Len(Trim(gobjLoja.sFTPURL)) > 0 Then
            
        'Prepara para chamar rotina batch
        lErro = Sistema_Preparar_Batch(sNomeArqParam)
        If lErro <> SUCESSO Then gError 133512
            
        gobjLoja.sNomeArqParam = sNomeArqParam
            
        Set objObject = gobjLoja
            
        lErro = CF("Rotina_FTP_Recepcao_CC", objObject, 6)
        If lErro <> SUCESSO And lErro <> 133628 Then gError 133513
            
        If lErro <> SUCESSO Then gError 133631
            
        File1.Refresh
            
        Call Rotina_Aviso(vbOK, "AVISO_DOWNLOAD_ARQUIVOS_SUCESSO")
            
    Else
    
        gError 133519
            
    End If

    Exit Sub
    
Erro_BotaoFTP_Click:

    Select Case gErr
        
        Case 133512, 133513
        
        Case 133519
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXA_ARQUIVOS_NAO_REALIZADA", gErr)
        
        Case 133631
            Call Rotina_Erro(vbOKOnly, "AVISO_NAO_CARREGOU_ROTINA_RECEPCAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159808)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Dir1_Change()
    
'    File1.Path = Dir1.List(Dir1.ListIndex)
'    DirName.Caption = ""
    
        If File1.Path <> Dir1.Path Then
            File1.Path = Dir1.Path
            DirName.Caption = ""
        End If
    
    
End Sub

Private Sub File1_Click()
    
    DirName.Caption = Dir1.List(Dir1.ListIndex) & "\" & File1.List(File1.ListIndex)
        
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim sRegistro As String
Dim sNomeArq As String

On Error GoTo Erro_Gravar_Registro
    
    If Len(DirName.Caption) = 0 Then gError 112709
    
    'pesquisa a existencia do arquivo
    sNomeArq = Dir(DirName.Caption)

    'se o arquivo não foi encontrado ==> erro
    If Len(sNomeArq) = 0 Then gError 118091
    
    Open DirName.Caption For Input Lock Write As #1

        Line Input #1, sRegistro
        
    Close #1
    
    lErro = CF("Verifica_Nome_Arquivo1", sRegistro)
    If lErro <> SUCESSO Then gError 112711
    
    gobjArq.sNome = DirName.Caption
    
    giRetornoTela = vbOK
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
        
        Case 112709
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEARQ_NAO_PREENCHIDO1", gErr)
            
        Case 112711
                
        Case 118091
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_INEXISTENTE", gErr, sNomeArq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159809)
    
    End Select

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a gravar registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 112723
    
    'fecha a tela
    Unload Me

    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
    
        Case 112723
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159810)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    
    giRetornoTela = vbCancel
    
    Unload Me
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim lTamanho As Long
Dim sRetorno As String
Dim objLojaConfig As New ClassLojaConfig
Dim sNomeArqParam As String
Dim lIntervaloTrans As Long

On Error GoTo Erro_Form_Load
    
    Set gobjArq = New AdmCodigoNome
    
    objLojaConfig.iFilialEmpresa = EMPRESA_TODA
    objLojaConfig.sCodigo = DIRETORIO_TELA_EXIBIRARQUIVOS
    
    lErro = CF("LojaConfig_Le1", objLojaConfig)
    If lErro <> SUCESSO And lErro <> 126361 Then gError 126362
    
    'se nao encontrou o registro q armazena o ultimo diretorio acessado para esta tela
    If lErro = 126361 Then objLojaConfig.sConteudo = "."
    
    Dir1.Path = objLojaConfig.sConteudo
    
    File1.Path = objLojaConfig.sConteudo
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 53, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCAL_ARQUIVO_INVALIDO", gErr, sRetorno)
            
        Case 127099
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCAL_ARQUIVO_NAO_CONFIGURADO", gErr)
            
        Case 126362, 133510, 133511
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159811)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objArq As AdmCodigoNome) As Long

On Error GoTo Erro_Trata_Parametros
    
    Set gobjArq = objArq
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159812)
    
    End Select
    
    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim objLojaConfig As New ClassLojaConfig
Dim lErro As Long

On Error GoTo Erro_Form_Unload
    
    objLojaConfig.iFilialEmpresa = EMPRESA_TODA
    objLojaConfig.sCodigo = DIRETORIO_TELA_EXIBIRARQUIVOS
    objLojaConfig.sConteudo = Dir1.Path
    
    'grava o ultimo diretorio
    lErro = CF("LojaConfig_Grava", objLojaConfig)
    If lErro <> SUCESSO Then gError 126363
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr
    
        Case 126363
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159813)
    
    End Select
    
    Exit Sub

End Sub



'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Arquivos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ExibirArquivos"

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

Private Sub File1_DblClick()
    
    Call BotaoGravar_Click
    
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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

'**** fim do trecho a ser copiado *****



