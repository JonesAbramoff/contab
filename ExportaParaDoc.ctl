VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.UserControl ExportaParaDocOcx 
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9270
   ScaleHeight     =   2895
   ScaleWidth      =   9270
   Begin VB.Frame Frame3 
      Caption         =   "Identificação"
      Height          =   645
      Left            =   45
      TabIndex        =   16
      Top             =   660
      Width           =   9090
      Begin VB.CheckBox OpcaoPadrao 
         Caption         =   "Padrão"
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
         Left            =   6915
         TabIndex        =   22
         Top             =   285
         Width           =   975
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   8070
         Picture         =   "ExportaParaDoc.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Gravar"
         Top             =   165
         Width           =   390
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   330
         Left            =   8520
         Picture         =   "ExportaParaDoc.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Excluir"
         Top             =   165
         Width           =   390
      End
      Begin VB.ComboBox OpcoesTela 
         Height          =   315
         Left            =   1620
         TabIndex        =   18
         Top             =   210
         Width           =   5160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opções:"
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
         Index           =   3
         Left            =   855
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   255
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Documento"
      Height          =   600
      Left            =   60
      TabIndex        =   8
      Top             =   75
      Width           =   9075
      Begin VB.Label ChaveDoc 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   4620
         TabIndex        =   12
         Top             =   165
         Width           =   4290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Chave:"
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
         Left            =   3960
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label TipoDoc 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1605
         TabIndex        =   10
         Top             =   180
         Width           =   2190
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc:"
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
         Index           =   0
         Left            =   705
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Geração"
      Height          =   1440
      Index           =   6
      Left            =   60
      TabIndex        =   0
      Top             =   1335
      Width           =   9075
      Begin VB.CheckBox NomeArqAuto 
         Caption         =   "Auto"
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
         Left            =   5355
         TabIndex        =   17
         Top             =   1050
         Value           =   1  'Checked
         Width           =   720
      End
      Begin VB.TextBox NomeDiretorio 
         Height          =   285
         Left            =   1605
         TabIndex        =   14
         Top             =   675
         Width           =   3690
      End
      Begin VB.CommandButton BotaoProcurarDir 
         Caption         =   "..."
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
         Left            =   5325
         TabIndex        =   13
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox NomeArquivo 
         Height          =   315
         Left            =   1605
         MaxLength       =   80
         TabIndex        =   5
         ToolTipText     =   "Nome do arquivo de proposta a ser gerado"
         Top             =   1005
         Width           =   3675
      End
      Begin VB.CommandButton BotaoGerarArq 
         Caption         =   "Gerar Arquivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7530
         Picture         =   "ExportaParaDoc.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gera um arquivo de proposta com base no modelo escolhido"
         Top             =   645
         Width           =   1335
      End
      Begin VB.CommandButton BotaoProcurarModelo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8370
         TabIndex        =   3
         Top             =   285
         Width           =   495
      End
      Begin VB.TextBox Modelo 
         Height          =   315
         Left            =   1605
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   2
         ToolTipText     =   "Modelo base para geração da proposta (.doc)"
         Top             =   300
         Width           =   6705
      End
      Begin VB.CommandButton BotaoMnemonicos 
         Caption         =   "Mnemônicos Válidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6150
         TabIndex        =   1
         ToolTipText     =   "Mnemônicos válidos para utilização em modelo do word"
         Top             =   645
         Width           =   1305
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8550
         Top             =   930
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Localização:"
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
         Index           =   2
         Left            =   480
         TabIndex        =   15
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Arquivo Modelo:"
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
         Index           =   9
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   7
         Top             =   345
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Arquivo:"
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
         Index           =   16
         Left            =   60
         TabIndex        =   6
         Top             =   1065
         Width           =   1725
      End
   End
End
Attribute VB_Name = "ExportaParaDocOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAtualizaTela As Integer

Dim iAlterado As Integer
Dim giTipoDoc As Integer
Dim gsOpcaoAnt As String
Dim gsNomeTelaDoc As String
Dim gsNomeDocPadrao As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

''constantes do word - ini
'Private Const wdCharacter = 1
'Private Const wdGoToField = 7
'Private Const wdWord9TableBehavior = 1
'Private Const wdAutoFitContent = 1
'Private Const wdGoToLine = 3
'Private Const wdCell = 12
''constantes do word - fim

'Botão ProcurarDir - ini
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
'Botão ProcurarDir - fim

'Auxiliar - ini -> Objs globais para cálculo dos mnemônicos
Dim gobjCli As ClassCliente
Dim gobjFilCli As ClassFilialCliente
Dim gobjEndCli As ClassEndereco
Dim gobjOV As ClassOrcamentoVenda
Dim gobjProjeto As ClassProjetos
Dim gobjTribDoc As ClassTributacaoDoc
Dim gobjInfoAdic As ClassInfoAdic
Dim gobjCondPagto As ClassCondicaoPagto
'Auxiliar - fim

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Exportação para o Word"
    Call Form_Load
    
End Function

Public Function Name() As String
    Name = gsNomeTelaDoc
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
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property
'**** fim do trecho a ser copiado *****

Private Sub Form_Unload()
    Set gobjCli = Nothing
    Set gobjFilCli = Nothing
    Set gobjEndCli = Nothing
    Set gobjOV = Nothing
    Set gobjProjeto = Nothing
End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
 
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211247)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros(ByVal iTipoDoc As Integer, ByVal objDoc As Object) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long
Dim objTela As Object
Dim objProjetoInfo As New ClassProjetoInfo
Dim iTipoOrigemPRJ As Integer, lNumIntDocOrigem As Long
Dim sCodOP As String, iFilialEmpresa As Integer

On Error GoTo Erro_Trata_Parametros

    gsNomeTelaDoc = "ExportaParaDoc"
    giTipoDoc = iTipoDoc
    
    Select Case iTipoDoc
    
        Case MNEMONICO_MALADIRETA_TIPO_OV
        
            gsNomeTelaDoc = "ExportaParaDocOV"
        
            Set gobjOV = objDoc
            Set gobjTribDoc = gobjOV.objTributacao
            Set gobjInfoAdic = gobjOV.objInfoAdic
            
            Set gobjCli = New ClassCliente
            Set gobjFilCli = New ClassFilialCliente
            
            gobjCli.lCodigo = gobjOV.lCliente
            gobjFilCli.lCodCliente = gobjOV.lCliente
            gobjFilCli.iCodFilial = gobjOV.iFilial
            
            If gobjOV.iHistorico <> MARCADO Then
                gobjInfoAdic.iTipoDoc = TIPODOC_INFOADIC_OV
                iTipoOrigemPRJ = PRJ_CR_TIPO_OV
            Else
                gobjInfoAdic.iTipoDoc = TIPODOC_INFOADIC_OVHIST
                iTipoOrigemPRJ = PRJ_CR_TIPO_OVHIST
            End If
            gobjInfoAdic.lNumIntDoc = gobjOV.lNumIntDoc
            
            If gobjOV.iCondicaoPagto <> 0 Then
                Set gobjCondPagto = New ClassCondicaoPagto
                gobjCondPagto.iCodigo = gobjOV.iCondicaoPagto
            End If
    
            iFilialEmpresa = gobjOV.iFilialEmpresa
            lNumIntDocOrigem = gobjOV.lNumIntDoc
            sCodOP = ""
            
            TipoDoc.Caption = "Orçamento de Venda"
            ChaveDoc.Caption = "Código: " & CStr(gobjOV.lCodigo) & " Fil.Emp.: " & CStr(gobjOV.iFilialEmpresa)
        
            gsNomeDocPadrao = "OrcVenda_" & Format(gobjOV.iFilialEmpresa, "00") & "_" & Format(gobjOV.lCodigo, "000000000") & gobjCRFAT.sExtensaoGerRelExp
        
    End Select
    
    If Not (gobjCli Is Nothing) Then
    
        lErro = CF("Cliente_Le", gobjCli)
        If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    If Not (gobjFilCli Is Nothing) Then

        lErro = CF("FilialCliente_Le", gobjFilCli)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Set gobjEndCli = New ClassEndereco
        
        gobjEndCli.lCodigo = gobjFilCli.lEndereco
    
        lErro = CF("Endereco_Le", gobjEndCli)
        If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If Not (gobjInfoAdic Is Nothing) Then
    
        lErro = CF("InfoAdicionais_Le", gobjInfoAdic)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    If Not (gobjCondPagto Is Nothing) Then
    
        lErro = CF("CondicaoPagto_Le", gobjCondPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    objProjetoInfo.iTipoOrigem = iTipoOrigemPRJ
    objProjetoInfo.lNumIntDocOrigem = lNumIntDocOrigem
    objProjetoInfo.sCodigoOP = sCodOP
    objProjetoInfo.iFilialEmpresa = iFilialEmpresa

    'Le as associação gravadas no BD para esse tipo de documento
    lErro = CF("ProjetoInfo_Le", objProjetoInfo)
    If lErro = SUCESSO Then
    
        Set gobjProjeto = New ClassProjetos
        
        If objProjetoInfo.lNumIntDocPRJ > 0 Then
        
            gobjProjeto.lNumIntDoc = objProjetoInfo.lNumIntDocPRJ
            
            lErro = CF("Projetos_Le_NumIntDoc", gobjProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181679
            
        End If
    
    End If
    
    'Guarda em objTela os dados dessa tela
    Set objTela = Me
    
    lErro = CF("Carrega_OpcoesTela", objTela, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If NomeArqAuto.Value = vbChecked Then Call Trata_NomeArq
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211248)
    
    End Select
    
    iAlterado = 0
    
    Exit Function
    
End Function

Private Sub BotaoGerarArq_Click()

Dim lErro As Long
Dim objTela As Object
Dim sDiretorio As String

On Error GoTo Erro_BotaoGerarArq_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 189005
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 189006
    If Len(Trim(Modelo.Text)) = 0 Then gError 189007
    
    If InStr(1, NomeArquivo.Text, ".") = 0 Then
        NomeArquivo.Text = NomeArquivo.Text & ".doc"
    End If
    
    If right(NomeDiretorio.Text, 1) = "\" Or right(NomeDiretorio.Text, 1) = "/" Then
        sDiretorio = NomeDiretorio.Text & NomeArquivo.Text
    Else
        sDiretorio = NomeDiretorio.Text & "\" & NomeArquivo.Text
    End If
    
    lErro = Gera_Arquivo_Doc(sDiretorio)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGerarArq_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 189005
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_INFORMADO", gErr)
            NomeDiretorio.SetFocus
        
        Case 189006
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_INFORMADO", gErr)
            NomeArquivo.SetFocus
        
        Case 189007
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_INFORMADO", gErr)
            Modelo.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187949)

    End Select

    Exit Sub

End Sub

Private Function Gera_Arquivo_Doc(ByVal sDiretorio As String)

Dim lErro As Long
'Dim objWord As Object 'Word.Application
'Dim objDoc As Object 'Word.Document
'Dim objCampoForm As Object 'Word.FormField
Dim objMnemonicoMala As ClassMnemonicoMalaDireta
Dim vValor As Variant
Dim sMesNome As String
Dim objFSO As New FileSystemObject
'Dim dVersaoWord As Double
Dim sProdMask As String, iIndice As String

Dim objItemOV As ClassItemOV
Dim objParcOV As ClassParcelaOV
Dim objWordApp As New ClassWordApp, lNumFormFields As Long, lIndiceFF As Long, sNomeFF As String
Dim dVlrAcum As Double, sTextoAux As String

On Error GoTo Erro_Gera_Arquivo_Doc

    'Set objWord = CreateObject("Word.Application")
    
    lErro = objWordApp.Abrir
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'    dVersaoWord = 0
'    If IsNumeric(objWord.Version) Then
'        dVersaoWord = StrParaDbl(Replace(objWord.Version, ".", ","))
'    End If
'
'    If dVersaoWord < 15 Then
'        Set objDoc = objWord.Documents.Open(Modelo.Text, , True)
'    Else
'        Set objDoc = objWord.Documents.Open(Modelo.Text)
'    End If

    lErro = objWordApp.Abrir_Doc(Modelo.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'For Each objCampoForm In objDoc.FormFields
    
    lNumFormFields = objWordApp.Qtde_FormFields()
    
    For lIndiceFF = lNumFormFields To 1 Step -1
    
        'Call objCampoForm.Select
    
        'lErro = objWordApp.FormField_Seleciona(lIndiceFF)
        'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        vValor = ""
        Set objMnemonicoMala = New ClassMnemonicoMalaDireta
        
        lErro = objWordApp.FormField_Obtem_Nome(lIndiceFF, sNomeFF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        objMnemonicoMala.sMnemonico = sNomeFF 'objCampoForm.Name
        objMnemonicoMala.iTipo = giTipoDoc
        
        lErro = CF("MnemonicoMalaDireta_Le", objMnemonicoMala)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 187956
        
        If lErro <> SUCESSO Then gError 187957
    
        Select Case objMnemonicoMala.iTipoObj
        
            Case MNEMONICO_MALADIRETA_TIPOOBJ_CLIENTE
            
                lErro = Critica_ObjetoAtributo(gobjCli, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case MNEMONICO_MALADIRETA_TIPOOBJ_FILIALCLIENTE
            
                lErro = Critica_ObjetoAtributo(gobjFilCli, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                If objMnemonicoMala.sMnemonico = "CGC_Cliente" Then
                    Select Case Len(Trim(vValor))
                        Case STRING_CPF 'CPF
                            vValor = Format(vValor, "000\.000\.000-00; ; ; ")
                        Case STRING_CGC 'CGC
                            vValor = Format(vValor, "00\.000\.000\/0000-00; ; ; ")
                    End Select
                End If

            Case MNEMONICO_MALADIRETA_TIPOOBJ_PROJETO

                lErro = Critica_ObjetoAtributo(gobjProjeto, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            Case MNEMONICO_MALADIRETA_TIPOOBJ_ESCOPO

                lErro = Critica_ObjetoAtributo(gobjProjeto.objEscopo, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            Case MNEMONICO_MALADIRETA_TIPOOBJ_ENDERECO_CLIENTE

                lErro = Critica_ObjetoAtributo(gobjEndCli, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case MNEMONICO_MALADIRETA_TIPOOBJ_OV

                lErro = Critica_ObjetoAtributo(gobjOV, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case MNEMONICO_MALADIRETA_TIPOOBJ_TRIBUTACAODOC
            
                lErro = Critica_ObjetoAtributo(gobjTribDoc, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            Case MNEMONICO_MALADIRETA_TIPOOBJ_INFOADIC

                lErro = Critica_ObjetoAtributo(gobjInfoAdic, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case MNEMONICO_MALADIRETA_TIPOOBJ_INFOADIC_COMPRA

                lErro = Critica_ObjetoAtributo(gobjInfoAdic.objCompra, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case MNEMONICO_MALADIRETA_TIPOOBJ_INFOADIC_EXPORT
            
                lErro = Critica_ObjetoAtributo(gobjInfoAdic.objExportacao, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case MNEMONICO_MALADIRETA_TIPOOBJ_INFOADIC_ENDENT

                lErro = Critica_ObjetoAtributo(gobjInfoAdic.objRetEnt.objEnderecoEnt, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case MNEMONICO_MALADIRETA_TIPOOBJ_INFOADIC_ENDRET

                lErro = Critica_ObjetoAtributo(gobjInfoAdic.objRetEnt.objEnderecoRet, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case MNEMONICO_MALADIRETA_TIPOOBJ_CONDPAGTO

                lErro = Critica_ObjetoAtributo(gobjCondPagto, objMnemonicoMala.sNomeCampoObj, vValor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case MNEMONICO_MALADIRETA_TIPOOBJ_OUTROS
            
                Select Case objMnemonicoMala.sMnemonico
                
                    Case "Dia_Agora"
                        vValor = Format(Day(Now), "00")
                        
                    Case "Mes_Agora"
                        vValor = Format(Month(Now), "00")
                
                    Case "Ano_Agora"
                        vValor = Format(Year(Now), "0000")
                
                    Case "Hora_Agora"
                        vValor = Format(Now, "HH:MM:SS")
                
                    Case "Data_Agora"
                        vValor = Format(Now, "DD/MM/YYYY")
                
                    Case "Mes_Agora_Nome"
                        Call MesNome(Month(Now), sMesNome)
                        vValor = sMesNome
                
                    Case "Lista_ItensOV"
                                                            
'                        Call DOC_Cria_Tabela(objDoc, objWord, gobjOV.colItens.Count + 1, 8)
'                        Call DOC_Insere_Cabec_Tabela(objWord, "Item", "Produto", "Descrição", "Qtde", "UM", "Preço Unitário", "Desconto", "Preço Total")
                
                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, gobjOV.colItens.Count + 1, 8)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Item", "Produto", "Descrição", "Qtde", "UM", "Preço Unitário", "Desconto", "Preço Total")
                
                        iIndice = 0
                        For Each objItemOV In gobjOV.colItens
                        
                            iIndice = iIndice + 1
                            Call Mascara_RetornaProdutoTela(objItemOV.sProduto, sProdMask)
                            sProdMask = Trim(sProdMask)
                        
                            'Call DOC_Insere_Valores_Tabela(objWord, CStr(iIndice), sProdMask, objItemOV.sDescricao, Formata_Estoque(objItemOV.dQuantidade), objItemOV.sUnidadeMed, Format(objItemOV.dPrecoUnitario, "STANDARD"), Format(objItemOV.dValorDesconto, "STANDARD"), Format(objItemOV.dPrecoTotal, "STANDARD"))
                            Call objWordApp.DOC_Insere_Valores_Tabela(CStr(iIndice), sProdMask, objItemOV.sDescricao, Formata_Estoque(objItemOV.dQuantidade), objItemOV.sUnidadeMed, Format(objItemOV.dPrecoUnitario, "STANDARD"), Format(objItemOV.dValorDesconto, "STANDARD"), Format(objItemOV.dPrecoTotal, "STANDARD"))
                        Next
                        
                    Case "Lista_ItensOV_DtEnt"
                                                            
                        'Call DOC_Cria_Tabela(objDoc, objWord, gobjOV.colItens.Count + 1, 9)
                        'Call DOC_Insere_Cabec_Tabela(objWord, "Item", "Produto", "Descrição", "Qtde", "UM", "Preço Unitário", "Desconto", "Preço Total", "Entrega")
                
                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, gobjOV.colItens.Count + 1, 9)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Item", "Produto", "Descrição", "Qtde", "UM", "Preço Unitário", "Desconto", "Preço Total", "Entrega")
                
                        iIndice = 0
                        For Each objItemOV In gobjOV.colItens
                        
                            iIndice = iIndice + 1
                            Call Mascara_RetornaProdutoTela(objItemOV.sProduto, sProdMask)
                            sProdMask = Trim(sProdMask)
                        
                            'Call DOC_Insere_Valores_Tabela(objWord, CStr(iIndice), sProdMask, objItemOV.sDescricao, Formata_Estoque(objItemOV.dQuantidade), objItemOV.sUnidadeMed, Format(objItemOV.dPrecoUnitario, "STANDARD"), Format(objItemOV.dValorDesconto, "STANDARD"), Format(objItemOV.dPrecoTotal, "STANDARD"), Formata_Data(objItemOV.dtDataEntrega))
                            Call objWordApp.DOC_Insere_Valores_Tabela(CStr(iIndice), sProdMask, objItemOV.sDescricao, Formata_Estoque(objItemOV.dQuantidade), objItemOV.sUnidadeMed, Format(objItemOV.dPrecoUnitario, "STANDARD"), Format(objItemOV.dValorDesconto, "STANDARD"), Format(objItemOV.dPrecoTotal, "STANDARD"), Formata_Data(objItemOV.dtDataEntrega))
                        Next
                        
                    Case "Lista_Dados_Cobranca"
                    
                        'Call DOC_Cria_Tabela(objDoc, objWord, gobjOV.colParcela.Count + 1, 3)
                        'Call DOC_Insere_Cabec_Tabela(objWord, "Parcela", "Vencimento", "Valor")
                
                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, gobjOV.colParcela.Count + 1, 3)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Parcela", "Vencimento", "Valor")
                
                        iIndice = 0
                        For Each objParcOV In gobjOV.colParcela
                            iIndice = iIndice + 1
                            'Call DOC_Insere_Valores_Tabela(objWord, CStr(iIndice), Format(objParcOV.dtDataVencimento, "dd/mm/yyyy"), Format(objParcOV.dValor, "STANDARD"))
                        
                            Call objWordApp.DOC_Insere_Valores_Tabela(CStr(iIndice), Format(objParcOV.dtDataVencimento, "dd/mm/yyyy"), Format(objParcOV.dValor, "STANDARD"))
                        Next
                        
                    Case "Lista_ItensOV_Obs"
                                                            
                        'Call DOC_Cria_Tabela(objDoc, objWord, gobjOV.colItens.Count + 1, 3)
                        'Call DOC_Insere_Cabec_Tabela(objWord, "Item", "Produto", "Observação")
                
                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, gobjOV.colItens.Count + 1, 3)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Item", "Produto", "Observação")
                
                        iIndice = 0
                        For Each objItemOV In gobjOV.colItens
                            iIndice = iIndice + 1
                        
                            Call Mascara_RetornaProdutoTela(objItemOV.sProduto, sProdMask)
                            sProdMask = Trim(sProdMask)
                        
                            'Call DOC_Insere_Valores_Tabela(objWord, CStr(iIndice), sProdMask, objItemOV.sObservacao)
                        
                            Call objWordApp.DOC_Insere_Valores_Tabela(CStr(iIndice), sProdMask, objItemOV.sObservacao)
                        Next
                        
                    Case "Lista_ItensOV_Msg"
                                                            
                        'Call DOC_Cria_Tabela(objDoc, objWord, gobjOV.colItens.Count + 1, 3)
                        'Call DOC_Insere_Cabec_Tabela(objWord, "Item", "Produto", "Mensagem")
                
                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, gobjOV.colItens.Count + 1, 3)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Item", "Produto", "Mensagem")
                
                        iIndice = 0
                        For Each objItemOV In gobjOV.colItens
                            iIndice = iIndice + 1
                        
                            Call Mascara_RetornaProdutoTela(objItemOV.sProduto, sProdMask)
                            sProdMask = Trim(sProdMask)
                        
                            'Call DOC_Insere_Valores_Tabela(objWord, CStr(iIndice), sProdMask, objItemOV.objInfoAdicDocItem.sMsg)
                        
                            Call objWordApp.DOC_Insere_Valores_Tabela(CStr(iIndice), sProdMask, objItemOV.objInfoAdicDocItem.sMsg)
                        Next
                                                
                    Case "Data_Agora_WS"
                        vValor = Format(Now, "DD-MMM-YY")
                        
                    Case "Lst_ItensOVV_WS"
                        dVlrAcum = 0
                        Call objWordApp.DOC_Cria_Tabela(lIndiceFF, gobjOV.colItens.Count + 2, 5)
                        Call objWordApp.DOC_Insere_Cabec_Tabela("Item", "Quant." & Space(10), "Description" & Space(130), "Unit. Price" & Space(15), "Total Price" & Space(15))
                        iIndice = 0
                        For Each objItemOV In gobjOV.colItens
                            iIndice = iIndice + 1
                            Call objWordApp.DOC_Insere_Valores_Tabela(CStr(iIndice), Formata_Estoque(objItemOV.dQuantidade), objItemOV.sDescricao, "R$ " & Format(objItemOV.dPrecoUnitario, "STANDARD"), "R$ " & Format(objItemOV.dPrecoTotal, "STANDARD"))
                            dVlrAcum = dVlrAcum + objItemOV.dPrecoTotal
                        Next
                        Call objWordApp.DOC_Insere_Valores_Tabela("", "", "", "TOTAL", "R$ " & Format(dVlrAcum, "STANDARD"))
                
                    Case "Lst_ItensOVI_WS"
                        iIndice = 0
                        For Each objItemOV In gobjOV.colItens
                            If Not (objItemOV.objInfoUsu Is Nothing) Then
                                If objItemOV.objInfoUsu.dprecounitimp > DELTA_VALORMONETARIO Then
                                    iIndice = iIndice + 1
                                End If
                            End If
                        Next
                        If iIndice > 0 Then
                            dVlrAcum = 0
                            Call objWordApp.DOC_Cria_Tabela(lIndiceFF, iIndice + 2, 5)
                            Call objWordApp.DOC_Insere_Cabec_Tabela("Item", "Quant." & Space(10), "Description" & Space(130), "Unit. Price" & Space(15), "Total Price" & Space(15))
                    
                            iIndice = 0
                            For Each objItemOV In gobjOV.colItens
                                If Not (objItemOV.objInfoUsu Is Nothing) Then
                                    If objItemOV.objInfoUsu.dprecounitimp > DELTA_VALORMONETARIO Then
                                        iIndice = iIndice + 1
                                        Call objWordApp.DOC_Insere_Valores_Tabela(CStr(iIndice), Formata_Estoque(objItemOV.dQuantidade), objItemOV.sDescricao, "£ " & Format(objItemOV.objInfoUsu.dprecounitimp, "STANDARD"), "£ " & Format(objItemOV.objInfoUsu.dprecounitimp * objItemOV.dQuantidade, "STANDARD"))
                                        dVlrAcum = dVlrAcum + Arredonda_Moeda(objItemOV.objInfoUsu.dprecounitimp * objItemOV.dQuantidade)
                                    End If
                                End If
                            Next
                            Call objWordApp.DOC_Insere_Valores_Tabela("", "", "", "TOTAL", "£ " & Format(dVlrAcum, "STANDARD"))
                        End If
                        
                    Case "Entrega", "Entrega_EN"
                        If gobjOV.iDataEnt = OV_DATA_ENTREGA_DATA Then
                            vValor = Format(gobjOV.dtDataEntrega, "DD/MM/YYYY")
                        ElseIf gobjOV.iDataEnt = OV_DATA_ENTREGA_TEXTO Then
                            vValor = gobjOV.sPrazoTexto
                        Else
                            Select Case gobjOV.iDataEnt
                                Case OV_DATA_ENTREGA_PRAZO_DIAS_UTEIS
                                    If objMnemonicoMala.sMnemonico = "Entrega" Then
                                        sTextoAux = "Dias úteis"
                                    Else
                                        sTextoAux = "Working days"
                                    End If
                                Case OV_DATA_ENTREGA_PRAZO_DIAS_CORRIDOS
                                    If objMnemonicoMala.sMnemonico = "Entrega" Then
                                        sTextoAux = "Dias corridos"
                                    Else
                                        sTextoAux = "Days"
                                    End If
                                Case OV_DATA_ENTREGA_PRAZO_SEMANAS
                                    If objMnemonicoMala.sMnemonico = "Entrega" Then
                                        sTextoAux = "Semanas"
                                    Else
                                        sTextoAux = "Weeks"
                                    End If
                                Case OV_DATA_ENTREGA_PRAZO_MESES
                                    If objMnemonicoMala.sMnemonico = "Entrega" Then
                                        sTextoAux = "Meses"
                                    Else
                                        sTextoAux = "Months"
                                    End If
                            End Select
                       
                            vValor = CStr(gobjOV.iPrazoEntrega) & " " & sTextoAux
                        End If

                                                                                           
                    Case Else
                        gError 187964
                        
                End Select

            Case Else
                gError 187957
                
        End Select
        
        Select Case UCase(TypeName(vValor))
        
            Case "DATE"
                vValor = Formata_Data(vValor)
        
            Case "DOUBLE"
                vValor = Format(vValor, "STANDARD")
        
        End Select
        
        'objCampoForm.Range.Text = vValor

        lErro = objWordApp.FormField_Preenche_Valor(lIndiceFF, vValor)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Next
    
    lErro = objWordApp.Salvar(sDiretorio)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If UCase(right(sDiretorio, 3)) = "PDF" Or UCase(right(sDiretorio, 4)) = "PDF""" Then
        'Salva o documento
'        objDoc.ExportAsFixedFormat OutputFileName:=Replace(sDiretorio, """", ""), ExportFormat:=17
'        Call objDoc.Close(False)

        lErro = objWordApp.Fechar()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Call ShellExecute(hWnd, "open", sDiretorio, vbNullString, vbNullString, 1)
    Else
        'objDoc.SaveAs FileName:=sDiretorio, FileFormat:=0
        'objWord.Visible = True
    
        lErro = objWordApp.Mudar_Visibilidade(True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    Gera_Arquivo_Doc = SUCESSO

    Exit Function

Erro_Gera_Arquivo_Doc:

    Gera_Arquivo_Doc = False

    'Call objDoc.Close(False)
    Call objWordApp.Fechar

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 187956, 187993

        Case 187957
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOMALADIRETA_NAO_CADASTRADO", gErr, objMnemonicoMala.sMnemonico, objMnemonicoMala.iTipo)
        
        Case 187958 To 187963
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOMALA_ATRIBUTO_INVALIDO", gErr, objMnemonicoMala.sNomeCampoObj, objMnemonicoMala.iTipo)

        Case 187964
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOMALA_NAO_TRATADO", gErr, objMnemonicoMala.sNomeCampoObj)

        Case 187965
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOMALA_TIPO_INVALIDO", gErr, objMnemonicoMala.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187955)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoMnemonicos_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objMnemonicoMala As New ClassMnemonicoMalaDireta

On Error GoTo Erro_BotaoMnemonicos_Click

    colSelecao.Add giTipoDoc

    Call Chama_Tela_Modal("MnemonicoMalaDiretaLista", colSelecao, objMnemonicoMala, Nothing, "Tipo = ?")

    Exit Sub

Erro_BotaoMnemonicos_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187955)

    End Select

    Exit Sub
    
End Sub

Private Function Critica_ObjetoAtributo(ByVal objObj As Object, ByVal sAtributo As String, vValor As Variant) As Long

On Error GoTo Erro_Critica_ObjetoAtributo
    
    vValor = CallByName(objObj, sAtributo, VbGet)

    Critica_ObjetoAtributo = SUCESSO

    Exit Function

Erro_Critica_ObjetoAtributo:

    Critica_ObjetoAtributo = gErr

    Exit Function

End Function

Private Function Formata_Data(ByVal dtData As Date) As String
    
    If dtData = DATA_NULA Then
        Formata_Data = ""
    Else
        Formata_Data = Format(dtData, "dd/mm/yyyy")
    End If

    Exit Function

End Function

'Private Sub DOC_Cria_Tabela(ByVal objDoc As Object, ByVal objWord As Object, ByVal iNumLinhas As Integer, ByVal iNumColunas As Integer)
'
'    'objWord.selection.GoToNext wdGoToField
'    objWord.selection.MoveRight wdCharacter, 1
'
'    objDoc.Tables.Add objWord.selection.Range, iNumLinhas, iNumColunas, wdWord9TableBehavior, wdAutoFitContent
'    objWord.selection.Tables(1).ApplyStyleHeadingRows = True
'    objWord.selection.Tables(1).ApplyStyleLastRow = True
'    objWord.selection.Tables(1).ApplyStyleFirstColumn = True
'    objWord.selection.Tables(1).ApplyStyleLastColumn = True
'
'    Exit Sub
'
'End Sub
'
'Private Sub DOC_Insere_Cabec_Tabela(ByVal objWord As Object, ParamArray avParams())
'
'Dim iIndice As Integer
'Dim bBold As Boolean
'
'    bBold = objWord.selection.Font.Bold
'    For iIndice = 0 To UBound(avParams)
'
'        objWord.selection.Font.Bold = True
'        objWord.selection.TypeText avParams(iIndice)
'
'        'If iIndice <> UBound(avParams) Then
'            objWord.selection.MoveRight wdCharacter, 1
'        'End If
'
'    Next
'    objWord.selection.Font.Bold = bBold
'
'    Exit Sub
'
'End Sub
'
'Private Sub DOC_Insere_Valores_Tabela(ByVal objWord As Object, ParamArray avParams())
'
'Dim iIndice As Integer
'
'    objWord.selection.MoveRight wdCell
'    'objWord.selection.GoToNext wdGoToLine
'
'    For iIndice = 0 To UBound(avParams)
'
'        objWord.selection.TypeText avParams(iIndice)
'
'        'If iIndice <> UBound(avParams) Then
'            objWord.selection.MoveRight wdCharacter, 1
'        'End If
'
'    Next
'
'    Exit Sub
'
'End Sub
'
'Private Sub DOC_Insere_Valores_Tabela2(ByVal objWord As Object, ParamArray avParams())
'
'Dim iIndice As Integer
'
'    For iIndice = 0 To UBound(avParams)
'
'        objWord.selection.TypeText avParams(iIndice)
'
'        'If iIndice <> UBound(avParams) Then
'            objWord.selection.MoveRight wdCharacter, 1
'        'End If
'
'    Next
'
'    Exit Sub
'
'End Sub

Private Sub BotaoProcurarModelo_Click()

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    
On Error GoTo Erro_BotaoProcurarModelo_Click
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Word Files" & _
    "(*.doc)|*.doc"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    Modelo.Text = CommonDialog1.FileName
    
    Exit Sub

Erro_BotaoProcurarModelo_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Private Sub BotaoProcurarDir_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurarDir_Click

    szTitle = "Localização física dos arquivos .html"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        NomeDiretorio.Text = sBuffer
        Call NomeDiretorio_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurarDir_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192326)

    End Select

    Exit Sub
  
End Sub

Public Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPos As Integer

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub
    
    If right(NomeDiretorio.Text, 1) <> "\" And right(NomeDiretorio.Text, 1) <> "/" Then
        iPos = InStr(1, NomeDiretorio.Text, "/")
        If iPos = 0 Then
            NomeDiretorio.Text = NomeDiretorio.Text & "\"
        Else
            NomeDiretorio.Text = NomeDiretorio.Text & "/"
        End If
    End If

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 192327

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 192327, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192328)

    End Select

    Exit Sub

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_OPCAO_BROWSER_GRAVADA_SUCESSO", OpcoesTela.Text)
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr

        Case 200108
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200109)

    End Select
    
End Sub

Public Sub OpcoesTela_Validate(Cancel As Boolean)

Dim vbResult As VbMsgBoxResult

    vbResult = vbYes
    If Len(Trim(OpcoesTela.Text)) > 20 Then
        vbResult = Rotina_Aviso(vbYesNo, "ERRO_CAMPO_ACIMA_TAM_PERMITIDO", "Opção", "20")
        If vbResult = vbYes Then OpcoesTela.Text = left(OpcoesTela.Text, 20)
    End If
    
    If vbResult = vbNo Then
        Cancel = True
    Else
        'Se a opção não foi selecionada na combo => chama a função OpcoesTela_Click
        If OpcoesTela.ListIndex = -1 Then Call OpcoesTela_Click
    End If
End Sub

Public Function Gravar_Registro() As Long

Dim objTela As Object
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    Set objTela = Me
    
    lErro = CF("OpcoesTelas_Grava", objTela)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200101)
        
    End Select
    
End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTela As Object

On Error GoTo Erro_BotaoExcluir_Click

    Set objTela = Me
    
    lErro = CF("OpcoesTelas_Exclui", objTela)
    If lErro <> SUCESSO Then gError 200102

    Exit Sub
    
Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 200102
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200103)

    End Select

End Sub

Public Sub OpcoesTela_Click()
    
Dim lErro As Long
Dim objTela As Object
Dim iCancel As Integer

On Error GoTo Erro_OpcoesTela_Click

    Set objTela = Me
    
    If gsOpcaoAnt <> OpcoesTela.Text Then
    
        iAtualizaTela = MARCADO
        
        'Trata o evento click da combo opções
        lErro = CF("OpcoesTela_Click", objTela)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        gsOpcaoAnt = OpcoesTela.Text
        
    End If
    
    Exit Sub

Erro_OpcoesTela_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200105)

    End Select

End Sub

Private Sub Trata_NomeArq()

    If NomeArqAuto.Value = vbChecked Then
        NomeArquivo.Text = gsNomeDocPadrao
        NomeArquivo.Enabled = False
    Else
        NomeArquivo.Enabled = True
    End If
    
End Sub

Private Sub NomeArqAuto_Click()
    Call Trata_NomeArq
End Sub
