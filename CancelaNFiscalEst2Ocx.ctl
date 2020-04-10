VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl CancelaNFiscalEst 
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   7095
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5220
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   165
      Width           =   1680
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "CancelaNFiscalEst2Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "CancelaNFiscalEst2Ocx.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CancelaNFiscalEst2Ocx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informações Adicionais"
      Height          =   1260
      Left            =   135
      TabIndex        =   18
      Top             =   2415
      Width           =   6795
      Begin MSMask.MaskEdBox MotivoCancel 
         Height          =   300
         Left            =   1365
         TabIndex        =   2
         Top             =   285
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDownCancelamento 
         Height          =   300
         Left            =   2445
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   720
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataCancelamento 
         Height          =   300
         Left            =   1365
         TabIndex        =   3
         Top             =   735
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cancelamento:"
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
         TabIndex        =   24
         Top             =   780
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo:"
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
         Left            =   600
         TabIndex        =   19
         Top             =   345
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   1605
      Left            =   135
      TabIndex        =   5
      Top             =   750
      Width           =   6795
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   1395
         TabIndex        =   0
         Top             =   300
         Width           =   765
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   300
         Left            =   3675
         TabIndex        =   1
         Top             =   285
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label LblDataEmissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3705
         TabIndex        =   17
         Top             =   1170
         Width           =   990
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
         Height          =   195
         Left            =   2910
         TabIndex        =   16
         Top             =   1230
         Width           =   765
      End
      Begin VB.Label LblFilial 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1395
         TabIndex        =   15
         Top             =   1185
         Width           =   1335
      End
      Begin VB.Label LabelFilial 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
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
         Left            =   885
         TabIndex        =   14
         Top             =   1230
         Width           =   465
      End
      Begin VB.Label Fornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Emitente:"
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
         Left            =   540
         TabIndex        =   13
         Top             =   795
         Width           =   810
      End
      Begin VB.Label LblSerie 
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
         Left            =   840
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   360
         Width           =   510
      End
      Begin VB.Label LblNumero 
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
         Left            =   2925
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Left            =   4830
         TabIndex        =   10
         Top             =   1230
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Left            =   5145
         TabIndex        =   9
         Top             =   330
         Width           =   450
      End
      Begin VB.Label LblEmitente 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1395
         TabIndex        =   8
         Top             =   765
         Width           =   5280
      End
      Begin VB.Label LblValor 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5400
         TabIndex        =   7
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label LblTipoNF 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5685
         TabIndex        =   6
         Top             =   315
         Width           =   1020
      End
   End
End
Attribute VB_Name = "CancelaNFiscalEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Dim WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFEntrada As New ClassNFiscal
Dim sTela As String
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim sMotivo As String
Dim sDiretorio As String
Dim lRetorno As Long
Dim iFilialEmpresa As Integer
Dim lNumIntNF As Long
Dim objVersao As New ClassVersaoNFe
Dim sStat As String
Dim dtData As Date, bNFJaCanceladaNoCorporator As Boolean
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoGravar_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    bNFJaCanceladaNoCorporator = False
    objNFEntrada.iRollBack = 0
            
    iFilialEmpresa = giFilialEmpresa
    
    'verifica se todos os campos estao preenchidos ,se nao estiverem => erro
    If Len(Trim(Serie.Text)) = 0 Then gError 34624
    If Len(Trim(Numero.ClipText)) = 0 Then gError 34625
    If Len(Trim(DataCancelamento.ClipText)) = 0 Then gError 182847
    If Len(Trim(MotivoCancel.Text)) < 15 Then gError 201490

    'Move os dados da NF de entrada para objNFEntrada
    lErro = Move_Dados_NFiscal_Memoria(objNFEntrada)
    If lErro <> SUCESSO Then gError 34659

    'Lê a nota fiscal de entrada
    lErro = CF("NFiscalInternaEntrada_Le_Numero", objNFEntrada)
    If lErro <> SUCESSO And lErro <> 62144 Then gError 62148
    If lErro <> SUCESSO Then gError 62149
    
    If gobjCRFAT.iUsaNFSE = DESMARCADO And gobjCRFAT.iUsaNFe = DESMARCADO And objNFEntrada.iStatus = STATUS_CANCELADO Then gError 62142
    
    If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then Error 62188


    If objNFEntrada.iStatus <> STATUS_CANCELADO Then

        'pede confirmacao
         vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_NFISCALENTRADA", Numero.Text)
    
        If vbMsg = vbYes Then
        
'            lErro = CF("NFeFedProtNFE_Le1", objNFEntrada.lNumIntDoc, sStat, dtData)
'            If lErro <> SUCESSO Then gError 207913
'
'            If sStat = "100" And DateDiff("d", dtData, Date) > 7 Then gError 207914

            lErro = CF("NFiscal_Valida_Canc", objNFEntrada)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            'Lê os itens da nota fiscal
            lErro = CF("NFiscalItens_Le", objNFEntrada)
            If lErro <> SUCESSO Then gError 62150
    
            objNFEntrada.sMotivoCancel = MotivoCancel.Text
    
            'Lê o nome da Tela onde Nota Fiscal foi gravada
            objTipoDocInfo.sSigla = LblTipoNF.Caption
            
            'Lê o nome da Tela onde Nota Fiscal foi gravada
            lErro = CF("TipoDocInfo_Le", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 27263 Then gError 67544
            If lErro = 27263 Then gError 67545
            
            If gobjCRFAT.iUsaNFe = MARCADO And ISSerieEletronica(objNFEntrada.sSerie) Then objNFEntrada.iRollBack = 1 'indica que deve fazer rollback do cancelamento pois só deve testar nesta fase
            
            'chama NotaFiscalEntrada_Cancelar()
            lErro = CF("NotaFiscalEntrada_Cancelar", objNFEntrada, StrParaDate(DataCancelamento.Text))
            If lErro <> SUCESSO Then gError 34660
            
            objNFEntrada.iRollBack = 0
            
        Else
        
            gError ERRO_SEM_MENSAGEM
        
        End If

    Else
    
        bNFJaCanceladaNoCorporator = True

    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    If giFilialEmpresa > 50 Then giFilialEmpresa = giFilialEmpresa - 50
    
    If gobjCRFAT.iUsaNFe = MARCADO And ISSerieEletronica(objNFEntrada.sSerie) Then
    
        'pede confirmacao
        'vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_CANCELAR_NFE", Numero.Text)
        vbMsg = vbYes
        
        sMotivo = MotivoCancel.Text
        
        If Len(Trim(sMotivo)) = 0 Then sMotivo = "*"
         
        sMotivo = Replace(sMotivo, " ", "_")
         
        If vbMsg = vbYes Then
    
            objVersao.iCodigo = gobjCRFAT.iVersaoNFE
            
            lErro = CF("VersaoNFe_Le", objVersao)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 62148
    
            sDiretorio = String(255, 0)
            lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
            sDiretorio = left(sDiretorio, lRetorno)
    
            If iFilialEmpresa <> giFilialEmpresa Then
                lNumIntNF = objNFEntrada.lNumIntDoc - 1
            Else
                lNumIntNF = objNFEntrada.lNumIntDoc
            End If
    
            lErro = WinExec(sDiretorio & objVersao.sProgramaEnvio & " Cancela " & CStr(glEmpresa) & " " & CStr(giFilialEmpresa) & " " & CStr(lNumIntNF) & " " & sMotivo, SW_NORMAL)
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_INICIO_CANCELANFE")
            
            If vbMsgRes = vbYes And bNFJaCanceladaNoCorporator = False Then

                lErro = CF("NFiscal_Valida_Homolog_Canc", objNFEntrada)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                'chama NotaFiscalEntrada_Cancelar()
                lErro = CF("NotaFiscalEntrada_Cancelar", objNFEntrada, StrParaDate(DataCancelamento.Text))
                If lErro <> SUCESSO Then gError 210910
            
            ElseIf bNFJaCanceladaNoCorporator = False Then
            
                lErro = CF("NFiscal_Valida_Homolog_Canc", objNFEntrada, True)
                If lErro <> SUCESSO And lErro <> ERRO_SEM_MENSAGEM Then gError ERRO_SEM_MENSAGEM
    
                If lErro = SUCESSO Then
                
                    'chama NotaFiscalEntrada_Cancelar()
                    lErro = CF("NotaFiscalEntrada_Cancelar", objNFEntrada, StrParaDate(DataCancelamento.Text))
                    If lErro <> SUCESSO Then gError 210910
            
                End If
            
            End If
            
        End If

    End If

    Call Limpa_Tela_NFEntrada

    iAlterado = 0

    giFilialEmpresa = iFilialEmpresa

    Exit Sub

Erro_BotaoGravar_Click:

    giFilialEmpresa = iFilialEmpresa

    Select Case gErr

        Case 34624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 34625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 34659, 34660, 62148, 62150, 67460, 67544, 207913, 210910, ERRO_SEM_MENSAGEM
        
        Case 62142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", gErr, Serie.Text, Numero.Text)
        
        Case 62149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, Numero.Text)
        
        Case 62188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", gErr)
        
        Case 67545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO", gErr, objTipoDocInfo.sSigla)
        
        Case 182847
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case 207914
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_NAO_PODE_SER_CANCELADA", gErr)
            
        Case 201490
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_MINIMO_15_CARACTERES", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144153)
        
    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoLimpar_Click

    If iAlterado = REGISTRO_ALTERADO Then

        'Testa se deseja salvar as alterações
        vbMsgRes = Rotina_Aviso(vbYesNoCancel, "AVISO_DESEJA_SALVAR_ALTERACOES")

        If vbMsgRes = vbYes Then

            Call BotaoGravar_Click

        ElseIf vbMsgRes = vbNo Then

            Call Limpa_Tela_NFEntrada

            iAlterado = 0

        Else
            Error 34661
        End If

    End If

Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 34661

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144154)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objNFEntrada As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFEntrada Is Nothing) Then

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscalInternaEntrada_Le_Numero", objNFEntrada)
        If lErro <> SUCESSO And lErro <> 62144 Then Error 34617
        If lErro = 62144 Then Error 34618
        
        If objNFEntrada.iStatus = STATUS_CANCELADO Then Error 62143
        If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then Error 62190

        'Traz a nota para a tela
        lErro = Traz_NFEntrada_Tela(objNFEntrada)
        If lErro <> SUCESSO Then Error 34619

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 34617, 34619

        Case 34618
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFEntrada.lNumNotaFiscal)
            Call Limpa_Tela_NFEntrada
            iAlterado = 0

        Case 62143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", Err, Serie.Text, Numero.Text)

        Case 62190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144155)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub Limpa_Tela_NFEntrada()
'Limpa a Tela NFiscalEntrada

    Serie.Text = ""
    Numero.PromptInclude = False
    Numero.Text = ""
    Numero.PromptInclude = True

    Call Limpa_Tela_NFEntrada1

End Sub

Public Function Traz_EmitFilial_Tela(iEmitente As Integer, objNFEntrada As ClassNFiscal) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Traz_EmitFilial_Tela

    'Se Emitente for empresa
    If iEmitente = EMITENTE_EMPRESA Then
        'EMPRESA
        objFilialEmpresa.iCodFilial = objNFEntrada.iFilialEmpresa

        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 58073

        If lErro = 27378 Then Error 58074
        
        If giTipoVersao = VERSAO_LIGHT Then
            LblFilial.Visible = False
            LabelFilial.Visible = False
        End If

        LblEmitente.Caption = gsNomeEmpresa
        LblFilial.Caption = giFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
        
    'Se Emitente for Cliente
    ElseIf iEmitente = EMITENTE_CLIENTE Then

        objCliente.lCodigo = objNFEntrada.lCliente

        'Procura se o CLiente existe
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then Error 34650
        If lErro = 12293 Then Error 34651

        objFilialCliente.lCodCliente = objNFEntrada.lCliente
        objFilialCliente.iCodFilial = objNFEntrada.iFilialCli

        'Procura se a Filial existe
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then Error 34652
        If lErro = 12567 Then Error 34653

        If giTipoVersao = VERSAO_LIGHT Then
            LblFilial.Visible = False
            LabelFilial.Visible = False
        End If

        LblEmitente.Caption = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
        LblFilial.Caption = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome

    'Se Emitente for Fornecedor
    ElseIf iEmitente = EMITENTE_FORNECEDOR Then

        objFornecedor.lCodigo = objNFEntrada.lFornecedor

        'procura pelo Fornecedor
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12732 Then Error 34654
        If lErro = 12732 Then Error 34655

        objFilialFornecedor.lCodFornecedor = objNFEntrada.lFornecedor
        objFilialFornecedor.iCodFilial = objNFEntrada.iFilialForn

        'procura pela filial do fornecedor
        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 12929 Then Error 34656
        If lErro = 12929 Then Error 34657

        If giTipoVersao = VERSAO_LIGHT Then
            LblFilial.Visible = True
            LabelFilial.Visible = True
        End If

        LblEmitente.Caption = objFornecedor.lCodigo & SEPARADOR & objFornecedor.sNomeReduzido
        LblFilial.Caption = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

    End If

    Traz_EmitFilial_Tela = SUCESSO

    Exit Function

Erro_Traz_EmitFilial_Tela:

    Traz_EmitFilial_Tela = Err

    Select Case Err

        Case 34650, 34652, 34654, 34656, 58073
            Numero.SetFocus

        Case 34651
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objNFEntrada.lCliente)
            Numero.SetFocus

        Case 34653
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", Err, objNFEntrada.lCliente)
            Numero.SetFocus

        Case 34655
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", Err, objNFEntrada.lFornecedor)
            Numero.SetFocus

        Case 34657
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_SEM_FILIAL", Err, objNFEntrada.lFornecedor)
            Numero.SetFocus

        Case 58074
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
            Numero.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144156)

    End Select

    Exit Function

End Function

Public Function Traz_NFEntrada_Tela(objNFEntrada As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal passada em objNFEntrada

Dim lErro As Long
Dim iIndice As Integer
Dim sTipoNF As String
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iEmitente As Integer

On Error GoTo Erro_Traz_NFEntrada_Tela

    'Limpa a tela NFicalEntrada
    Call Limpa_Tela_NFEntrada

    'Preenche o número da NF
    If objNFEntrada.lNumNotaFiscal > 0 Then
        Numero.PromptInclude = False
        Numero.Text = CStr(objNFEntrada.lNumNotaFiscal)
        Numero.PromptInclude = True
    End If

    'preenche a serie da NF
    Serie.Text = objNFEntrada.sSerie

    objTipoDocInfo.iCodigo = objNFEntrada.iTipoNFiscal

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 34648
    If lErro = 31415 Then Error 34649

    'preenche a Sigla da NF
    LblTipoNF.Caption = objTipoDocInfo.sSigla

    iEmitente = objTipoDocInfo.iEmitente

    'Traz a tela os dados de Emitente e Filial
    lErro = Traz_EmitFilial_Tela(iEmitente, objNFEntrada)
    If lErro <> SUCESSO Then Error 34658

    'Se a data não for nula coloca na Tela
    If objNFEntrada.dtDataEmissao <> DATA_NULA Then
        LblDataEmissao.Caption = Format(objNFEntrada.dtDataEmissao, "dd/mm/yy")
        DataCancelamento.PromptInclude = False
        DataCancelamento.Text = Format(objNFEntrada.dtDataEmissao, "dd/mm/yy")
        DataCancelamento.PromptInclude = True
    Else
        LblDataEmissao.Caption = Format("", "dd/mm/yy")
        DataCancelamento.PromptInclude = False
        DataCancelamento.Text = ""
        DataCancelamento.PromptInclude = True
    End If

    'Preenche o valor total da NF
    If objNFEntrada.dValorTotal > 0 Then
        LblValor.Caption = Format(objNFEntrada.dValorTotal, "Fixed")
    Else
        LblValor.Caption = Format(0, "Fixed")
    End If

    Traz_NFEntrada_Tela = SUCESSO

    Exit Function

Erro_Traz_NFEntrada_Tela:

    Traz_NFEntrada_Tela = Err

    Select Case Err

        Case 34648, 34658

        Case 34649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144157)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objSerie As ClassSerie
Dim colSerie As New colSerie

On Error GoTo Erro_Form_Load

    Set objEventoSerie = New AdmEvento
    Set objEventoNumero = New AdmEvento

    'nao pode entrar como EMPRESA_TODA
    If giFilialEmpresa = EMPRESA_TODA Then Error 34615

    'obtem a colecao de series
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then Error 34616

    'preenche as duas combos de serie
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case Err

        Case 34615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_INVALIDA", Err)

        Case 34616

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144158)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Move_Dados_NFiscal_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os dados da NotaFiscalOriginal para a memória

Dim lErro As Long

On Error GoTo Erro_Move_Dados_NFiscal_Memoria

    'verifica se a Serie e o Número da NF de entrada estão preenchidos
    If Len(Trim(Numero.ClipText)) > 0 Then objNFiscal.lNumNotaFiscal = CLng(Numero.Text)
    If Len(Trim(Serie.Text)) > 0 Then objNFiscal.sSerie = Serie.Text

    objNFiscal.iFilialEmpresa = giFilialEmpresa
    
    Move_Dados_NFiscal_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_NFiscal_Memoria:

    Move_Dados_NFiscal_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144159)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoSerie = Nothing
    Set objEventoNumero = Nothing

End Sub

Private Sub LblNumero_Click()

Dim lErro As Long
Dim objNFEntrada As New ClassNFiscal
Dim colSelecao As Collection

On Error GoTo Erro_LblNumero_Click

    'Preenche objNFEntrada com o numero
    lErro = Move_Dados_NFiscal_Memoria(objNFEntrada)
    If lErro <> SUCESSO Then Error 34620

    Call Chama_Tela("NFiscalInternaEntradaLista", colSelecao, objNFEntrada, objEventoNumero)

    Exit Sub

Erro_LblNumero_Click:

    Select Case Err

        Case 34620

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144160)

    End Select

    Exit Sub

End Sub

Private Sub LblSerie_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objSerie As New ClassSerie
Dim colSelecao As Collection

On Error GoTo Erro_LblSerie_Click

    'transfere a série da tela p\ o objSerie
    objSerie.sSerie = Serie.Text

    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

    Exit Sub

Erro_LblSerie_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144161)

    End Select

    Exit Sub

End Sub

Private Sub MotivoCancel_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()

    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNFEntrada As New ClassNFiscal

On Error GoTo Erro_Numero_Validate
    
    Cancel = False
    
    'Se a série não estiver preenchida, sai.
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
    'Se o número estiver preenchido
    If Len(Trim(Numero.ClipText)) > 0 Then
        'Recolhe a série e o número
        objNFEntrada.lNumNotaFiscal = Numero.Text
        objNFEntrada.sSerie = Serie.Text
        objNFEntrada.iFilialEmpresa = giFilialEmpresa
        
        'procura pela nota no BD
        lErro = CF("NFiscalInternaEntrada_Le_Numero", objNFEntrada)
        If lErro <> SUCESSO And lErro <> 62144 Then Error 34637
        If lErro = 62144 Then Error 34638 'Não encontrou
        'verifica se a nota já está cancelada
        'If objNFEntrada.iStatus = STATUS_CANCELADO Then Error 62144
        'If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then Error 62191

        'Traz a NotaFiscal de Entrada para a a tela
        lErro = Traz_NFEntrada_Tela(objNFEntrada)
        If lErro <> SUCESSO Then Error 34639

    End If

    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case Err

        Case 34637, 34639

        Case 34638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFEntrada.lNumNotaFiscal)
            Call Limpa_Tela_NFEntrada1
            iAlterado = 0
        
        Case 62144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", Err, Serie.Text, Numero.Text)

        Case 62191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144162)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFEntrada As ClassNFiscal

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objNFEntrada = obj1

    lErro = CF("NFiscalInternaEntrada_Le_Numero", objNFEntrada)
    If lErro <> SUCESSO And lErro <> 62144 Then Error 34632
    If lErro = 62144 Then Error 34633
    
    If objNFEntrada.iStatus = STATUS_CANCELADO Then Error 62145
    If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then Error 62192

    'Traz a NotaFiscal de Entrada para a a tela
    lErro = Traz_NFEntrada_Tela(objNFEntrada)
    If lErro <> SUCESSO Then Error 34623

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case Err

        Case 34623, 34632

        Case 34633
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFEntrada.lNumNotaFiscal)
            Call Limpa_Tela_NFEntrada
            iAlterado = 0

        Case 62145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", Err, Serie.Text, Numero.Text)
        
        Case 62192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144163)

    End Select

    Exit Sub

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objSerie As ClassSerie
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_objEventoSerie_evSelecao

    Set objSerie = obj1

    Serie.Text = objSerie.sSerie
    Call Serie_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoSerie_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144164)

    End Select

    Exit Sub

End Sub

Private Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNFEntrada As New ClassNFiscal
Dim objSerie As New ClassSerie

On Error GoTo Erro_Serie_Validate

    Cancel = False
    
    'Verifica se a série está preenchida
    If Len(Trim(Serie.Text)) > 0 Then
       'Verifica se o número está preenchido
       If Len(Trim(Numero.ClipText)) > 0 Then

            objNFEntrada.lNumNotaFiscal = Numero.Text
            objNFEntrada.sSerie = Serie.Text
            objNFEntrada.iFilialEmpresa = giFilialEmpresa
            
            'procura pela nota no BD
            lErro = CF("NFiscalInternaEntrada_Le_Numero", objNFEntrada)
            If lErro <> SUCESSO And lErro <> 62144 Then Error 34634
            If lErro = 62144 Then Error 34635

            If objNFEntrada.iStatus = STATUS_CANCELADO Then Error 62146
            'If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then Error 62193

            'Traz a NotaFiscal de Entrada para a a tela
            lErro = Traz_NFEntrada_Tela(objNFEntrada)
            If lErro <> SUCESSO Then Error 34636

        Else

            objSerie.sSerie = Serie.Text

            lErro = CF("Serie_Le", objSerie)
            If lErro <> SUCESSO And lErro <> 22202 Then Error 34646
            If lErro = 22202 Then Error 34647

        End If

    End If

    Exit Sub

Erro_Serie_Validate:

    Cancel = True
    
    Select Case Err

        Case 34634, 34636, 34646

        Case 34635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFEntrada.lNumNotaFiscal)
            Call Limpa_Tela_NFEntrada1
            iAlterado = 0

        Case 34647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, objSerie.sSerie)

        Case 62146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", Err, Serie.Text, Numero.Text)

        Case 62193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", Err)

        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144165)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_CANCELA_NFISCALEST
    Set Form_Load_Ocx = Me
    Caption = "Cancelamento de Nota Fiscal de Entrada"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CancelaNFiscal"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Serie Then
            Call LblSerie_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call LblNumero_Click
        End If

    End If

End Sub

Public Sub Limpa_Tela_NFEntrada1()
'Limpa a Tela NFiscalEntrada

    LblEmitente.Caption = ""
    LblFilial.Caption = ""
    LblValor.Caption = ""
    LblTipoNF.Caption = ""
    LblDataEmissao.Caption = ""
    MotivoCancel.Text = ""
    
    DataCancelamento.PromptInclude = False
    DataCancelamento.Text = ""
    DataCancelamento.PromptInclude = True
    
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub LblDataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblDataEmissao, Source, X, Y)
End Sub

Private Sub LblDataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblDataEmissao, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LblFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblFilial, Source, X, Y)
End Sub

Private Sub LblFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblFilial, Button, Shift, X, Y)
End Sub

Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub

Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub

Private Sub Fornecedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Fornecedor, Source, X, Y)
End Sub

Private Sub Fornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Fornecedor, Button, Shift, X, Y)
End Sub

Private Sub LblSerie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblSerie, Source, X, Y)
End Sub

Private Sub LblSerie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblSerie, Button, Shift, X, Y)
End Sub

Private Sub LblNumero_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNumero, Source, X, Y)
End Sub

Private Sub LblNumero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNumero, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LblEmitente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblEmitente, Source, X, Y)
End Sub

Private Sub LblEmitente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblEmitente, Button, Shift, X, Y)
End Sub

Private Sub LblValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblValor, Source, X, Y)
End Sub

Private Sub LblValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblValor, Button, Shift, X, Y)
End Sub

Private Sub LblTipoNF_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblTipoNF, Source, X, Y)
End Sub

Private Sub LblTipoNF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblTipoNF, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub DataCancelamento_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataCancelamento_Validate

    'Se a DataCancelamento está preenchida
    If Len(DataCancelamento.ClipText) <> 0 Then

        'Verifica se a DataCancelamento é válida
        lErro = Data_Critica(DataCancelamento.Text)
        If lErro <> SUCESSO Then gError 182848
        
    End If

    Exit Sub

Erro_DataCancelamento_Validate:

    Cancel = True

    Select Case gErr

        Case 182848

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182849)

    End Select

    Exit Sub

End Sub

Private Sub UpDownCancelamento_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownCancelamento_DownClick

    'Diminui a DataCancelamento em 1 dia
    lErro = Data_Up_Down_Click(DataCancelamento, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 182850

    Exit Sub

Erro_UpDownCancelamento_DownClick:

    Select Case gErr

        Case 182850

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182851)

    End Select

    Exit Sub

End Sub

Private Sub UpDownCancelamento_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownCancelamento_UpClick

    'Aumenta a DataCancelamento em 1 dia
    lErro = Data_Up_Down_Click(DataCancelamento, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 182852

    Exit Sub

Erro_UpDownCancelamento_UpClick:

    Select Case gErr

        Case 182852

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182853)

    End Select

    Exit Sub

End Sub
