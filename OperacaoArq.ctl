VERSION 5.00
Begin VB.UserControl OperacaoArq 
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   KeyPreview      =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   4515
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3165
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1140
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "OperacaoArq.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "F5 - Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "OperacaoArq.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "F8 - Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operação"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2505
      Begin VB.OptionButton OptionTransmitir 
         Caption         =   "Transmitir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   300
         TabIndex        =   2
         Top             =   210
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.OptionButton OptionRetransmitir 
         Caption         =   "Retransmitir"
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
         Left            =   300
         TabIndex        =   1
         Top             =   660
         Width           =   1455
      End
   End
End
Attribute VB_Name = "OperacaoArq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Function Retransmitir_Arquivo() As Long

Dim sRetorno As String
Dim lTamanho As Long
Dim lSeqArq As Long
Dim iCodEmpresa As Integer
Dim iIndice As Integer
Dim sNomeArq As String
Dim sNomeArq1 As String
Dim sRegistro As String
Dim bAchou As Boolean
Dim bEntrou As Boolean
Dim sNomeArq2 As String
Dim iPos As Integer
Dim sTipo As String
Dim sSeq As String
Dim colSeq As New Collection
Dim iIndiceInicio As Integer
Dim sRet As String
Dim sNomeArq3 As String
Dim iIndiceSeq As Integer
Dim iIndiceFim As Integer
Dim colSeqNaoEncontrado As New Collection
Dim sNomeArqErro As String
Dim lErro As Long
Dim colRegistroTemp As New Collection
Dim sSeqAbert As String
Dim sSeqFech As String
Dim lSeq As Long
Dim iPosInicio As Long
Dim iIndice1 As Integer
Dim sArquivosGerados As String
Dim sDirDadosECF As String
Dim lIntervaloTrans As Long
Dim objObject As Object

On Error GoTo Erro_Retransmitir_Arquivo
    
    Call Chama_TelaECF_Modal("ExibirSequenciais", colSeq)
        
    sArquivosGerados = Chr(vbKeyReturn)
        
    lTamanho = 255
    sDirDadosECF = String(lTamanho, 0)
    
    'Obtém o diretório onde deve ser armazenado o arquivo com dados do backoffice
    Call GetPrivateProfileString(APLICACAO_DADOS, "DirDadosECF", CONSTANTE_ERRO, sDirDadosECF, lTamanho, NOME_ARQUIVO_CAIXA)
    
    'Retira os espaços no final da string
    sDirDadosECF = StringZ(sDirDadosECF)
    
    'Se não encontrou
    If Len(Trim(sDirDadosECF)) = 0 Or sDirDadosECF = CStr(CONSTANTE_ERRO) Then gError 127100
    
    'se o diretorio nao for terminado por \  ===> acrescentar
    If right(sDirDadosECF, 1) <> "\" Then sDirDadosECF = sDirDadosECF & "\"
    
    For iIndiceSeq = 1 To colSeq.Count
                
        lSeqArq = colSeq.Item(iIndiceSeq)
        sNomeArq1 = giCodEmpresa & "_" & giFilialEmpresa & "_" & giCodCaixa & "_" & lSeqArq
        sNomeArq2 = sDirDadosECF & sNomeArq1 & ".tmp"
        sNomeArq3 = sDirDadosECF & sNomeArq1 & ".ccc"
        
        sRet = Dir(sNomeArq2, vbNormal)
        
        'se encontrou um arquivo com mesmo nome --> remove e continua
        If sRet <> "" Then Kill (sNomeArq2)
        
        sRet = Dir(sNomeArq3, vbNormal)
        
        'se encontrou um arquivo com mesmo nome --> remove e continua
        If sRet <> "" Then Kill (sNomeArq3)
                        
        'grava o arquivo sNomeArq2 com o conteudo de MovimentoCaixa referente a lSeqArq
        lErro = CF_ECF("Gravar_Arquivo_Retransmissao", sNomeArq2, sNomeArq1 & ".ccc", lSeqArq)
        If lErro <> SUCESSO Then gError 204756
                        
        'renomeando os arquivos
        Name sNomeArq2 As sNomeArq3
                        
        sArquivosGerados = sArquivosGerados & sNomeArq3 & ", " & Chr(vbKeyReturn)
                        
    Next
    
    If Len(gobjLojaECF.sFTPURL) > 0 Then
           
        Set objObject = gobjLojaECF
           
        lErro = CF_ECF("Rotina_FTP_Envio_Caixa", objObject, 2)
        If lErro <> SUCESSO Then gError 133403
           
    End If
    
    If Len(sArquivosGerados) > 2 Then
        sArquivosGerados = left(sArquivosGerados, Len(sArquivosGerados) - 3) & Chr(vbKeyReturn)
    End If
    
    'Pergunta se Deseja Efetuar a Sangria
    Call Rotina_AvisoECF(vbOK, AVISO_ARQUIVOS_RETRANSMITIDOS, sArquivosGerados)
    
    Retransmitir_Arquivo = SUCESSO
    
    Exit Function
    
Erro_Retransmitir_Arquivo:
    
    Retransmitir_Arquivo = gErr

    Select Case gErr
    
        Case 53
                        
        Case 119012
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_RETRANSMITIDO, gErr)
        
        Case 126080, 133403, 204756
        
        Case 133402
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "DirDadosECF", APLICACAO_DADOS, NOME_ARQUIVO_CAIXA)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163635)
        
    End Select
    
    Exit Function
    
End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim lSequencial As Long
Dim lIntervaloTrans As Long
Dim sRetorno As String
Dim lTamanho As Long
Dim objObject As Object

On Error GoTo Erro_Gravar_Registro

    If OptionTransmitir Then
        
   '**** MARIO colocado para testar o envio de boletos POSTEF
        
            'Função que Abre a Transação de Caixa, Identifica o Movimento dentro do Caixa
            lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
            If lErro <> SUCESSO Then gError 107559
        
'                'executa a sangria de todos boletos PosTef
'                lErro = CF_ECF("Caixa_Executa_Sangria_Boleto_PosTef", lSequencial)
'                If lErro <> SUCESSO Then gError 105430
        
        
            'Fechar a Transação
            lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
            If lErro <> SUCESSO Then gError 107541
        
        
        lErro = CF_ECF("Transmitir_Arquivo")
        If lErro <> SUCESSO And lErro <> 53 Then gError 112616
        
        'arquivo aberto
        If lErro = 53 Then gError 112619
        
        If Len(Trim(gobjLojaECF.sFTPURL)) > 0 Then
            
            Set objObject = gobjLojaECF
                
            lErro = CF_ECF("Rotina_FTP_Envio_Caixa", objObject, 2)
            If lErro <> SUCESSO Then gError 133401
            
        End If
        
    Else
        lErro = Retransmitir_Arquivo
        If lErro <> SUCESSO And lErro <> 53 Then gError 112617
        
        'arquivo aberto
        If lErro = 53 Then gError 133400
        
    End If
    
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
        
        Case 105688, 105689, 112616, 112617, 126152, 133391, 133401
        
        Case 112619, 133400
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_ABERTO, gErr)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163636)
    
    End Select
        
    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then

        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 207981

    End If

    'chama a gravar registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 109485
    
    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
    
        Case 109485, 207981
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163637)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163638)
    
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
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163639)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Operação de Arquivo"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "OperacaoArq"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Função que Incrementa o Código Atravez da Tecla F2
Dim lErro As Long

On Error GoTo Erro_UserControl_KeyDown

    Select Case KeyCode

        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoGravar) Then Exit Sub
            Call BotaoGravar_Click
            
        Case vbKeyF8
            If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
            Call BotaoFechar_Click

    End Select

    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163640)

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

