VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRPCancelarFatura 
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   7470
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   1710
      Left            =   105
      TabIndex        =   10
      Top             =   750
      Width           =   7320
      Begin MSMask.MaskEdBox Numero 
         Height          =   300
         Left            =   1095
         TabIndex        =   0
         Top             =   330
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label Status 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4080
         TabIndex        =   23
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Left            =   3420
         TabIndex        =   22
         Top             =   360
         Width           =   615
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
         Index           =   0
         Left            =   3300
         TabIndex        =   21
         Top             =   1275
         Width           =   765
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
         Left            =   585
         TabIndex        =   20
         Top             =   1275
         Width           =   465
      End
      Begin VB.Label Fornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   390
         TabIndex        =   19
         Top             =   825
         Width           =   660
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
         Left            =   330
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   360
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
         Left            =   5340
         TabIndex        =   17
         Top             =   1275
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
         Left            =   5655
         TabIndex        =   16
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Cliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1095
         TabIndex        =   15
         Top             =   810
         Width           =   6105
      End
      Begin VB.Label TipoDoc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6165
         TabIndex        =   14
         Top             =   330
         Width           =   1020
      End
      Begin VB.Label Valor 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5910
         TabIndex        =   13
         Top             =   1230
         Width           =   1275
      End
      Begin VB.Label FilialEmpresa 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1095
         TabIndex        =   12
         Top             =   1230
         Width           =   2160
      End
      Begin VB.Label Emissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4095
         TabIndex        =   11
         Top             =   1230
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informações Adicionais"
      Height          =   1170
      Left            =   105
      TabIndex        =   8
      Top             =   2535
      Width           =   7320
      Begin VB.ComboBox Motivo 
         Height          =   315
         Left            =   1095
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   285
         Width           =   6075
      End
      Begin MSComCtl2.UpDown UpDownCanc 
         Height          =   300
         Left            =   2220
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataCanc 
         Height          =   315
         Left            =   1095
         TabIndex        =   2
         Top             =   705
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Canc. CTB:"
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
         Left            =   75
         TabIndex        =   24
         Top             =   750
         Width           =   990
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
         Left            =   420
         TabIndex        =   9
         Top             =   345
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5655
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1680
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TRPCancelarFatura.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "TRPCancelarFatura.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "TRPCancelarFatura.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "TRPCancelarFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giTipoDocDestino As Integer
Dim gobjDestino As Object

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1

'Variáveis globais
Dim iAlterado As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iAlterado = 0
    
    Set objEventoNumero = New AdmEvento
    
    Motivo.Clear
    lErro = CF("Carrega_Combo_Historico", Motivo, "TRPTitulosExp", STRING_TRP_OCR_HISTORICO, "Motivo")
    If lErro <> SUCESSO Then gError 190165

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr
    
        Case 190165

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192673)

    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros(Optional ByVal objFat As ClassFaturaTRP) As Long
'A tela não espera recebimento de parâmetros, portanto, essa função sempre retorna sucesso
    
    If Not (objFat Is Nothing) Then
        Numero.PromptInclude = False
        Numero.Text = objFat.lNumFat
        Numero.PromptInclude = True
        Call Numero_Validate(bSGECancelDummy)
    End If

    Trata_Parametros = SUCESSO
End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Libera os objetos e coleções globais

End Sub
'*** FECHAMENTO DA TELA - FIM ***

Private Sub BotaoLimpar_Click()
'Dispara a limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa a tela
    Call Limpa_Tela_CancelarFatura

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192674)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub Limpa_Tela_CancelarFatura()
'Limpa a tela com exceção do campo 'Modelo'

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CancelarFatura

    'Limpa os controles básicos da tela
    Call Limpa_Tela(Me)
    
    TipoDoc.Caption = ""
    Status.Caption = ""
    Cliente.Caption = ""
    Emissao.Caption = ""
    Valor.Caption = ""
    FilialEmpresa.Caption = ""
    
    giTipoDocDestino = 0
    Set gobjDestino = Nothing
    Set objEventoNumero = Nothing
    
    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_CancelarFatura:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192688)

    End Select
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cancelamento de Faturas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRPCancelarFatura"

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

Private Sub Motivo_Change()
    iAlterado = REGISTRO_ALTERADO
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

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objDestino As Object
Dim iTipoDocDestino As Integer
Dim objTitRec As ClassTituloReceber
Dim objTitPag As ClassTituloPagar
Dim objNFsPag As ClassNFsPag
Dim objCliente As New ClassCliente
Dim objForn As New ClassFornecedor
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Numero_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Numero.ClipText)) <> 0 Then

        'Critica a Codigo
        lErro = Long_Critica(Numero.Text)
        If lErro <> SUCESSO Then gError 196501
       
        lErro = CF("TRPFaturas_Le", StrParaLong(Numero.Text), objDestino, iTipoDocDestino, True, False)
        If lErro <> SUCESSO Then gError 196502
       
        giTipoDocDestino = iTipoDocDestino
        Set gobjDestino = objDestino
        
        Select Case iTipoDocDestino
        
            Case TRP_TIPO_DOC_DESTINO_TITREC
            
                If objDestino.iStatus = STATUS_EXCLUIDO Then gError 196778
                
                Set objTitRec = objDestino
                objCliente.lCodigo = objTitRec.lCliente
                
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO And lErro <> 12293 Then gError 196504
                
                objFilialEmpresa.iCodFilial = objTitRec.iFilialEmpresa
                
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO Then gError 196505
                            
                TipoDoc.Caption = objTitRec.sSiglaDocumento
                If objTitRec.iStatus = STATUS_BAIXADO Then
                    Status.Caption = "Baixado"
                Else
                    Status.Caption = "Aberto"
                End If
                Cliente.Caption = CStr(objCliente.lCodigo) & SEPARADOR & objCliente.sNomeReduzido
                Emissao.Caption = Format(objTitRec.dtDataEmissao, "dd/mm/yyyy")
                Valor.Caption = Format(objTitRec.dValor, "STANDARD")
                FilialEmpresa.Caption = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
            
            Case TRP_TIPO_DOC_DESTINO_TITPAG
            
                If objDestino.iStatus = STATUS_EXCLUIDO Then gError 196778
            
                Set objTitPag = objDestino
                objForn.lCodigo = objTitPag.lFornecedor
                
                lErro = CF("Fornecedor_Le", objForn)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 196506
                
                objFilialEmpresa.iCodFilial = objTitPag.iFilialEmpresa
                
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO Then gError 196507
                            
                TipoDoc.Caption = objTitPag.sSiglaDocumento
                If objTitPag.iStatus = STATUS_BAIXADO Then
                    Status.Caption = "Baixado"
                Else
                    Status.Caption = "Aberto"
                End If
                Cliente.Caption = CStr(objForn.lCodigo) & SEPARADOR & objForn.sNomeReduzido
                Emissao.Caption = Format(objTitPag.dtDataEmissao, "dd/mm/yyyy")
                Valor.Caption = Format(objTitPag.dValorTotal, "STANDARD")
                FilialEmpresa.Caption = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
    
            Case TRP_TIPO_DOC_DESTINO_NFSPAG
     
                If objDestino.iStatus = STATUS_EXCLUIDO Then gError 196778
     
                Set objNFsPag = objDestino
                objForn.lCodigo = objNFsPag.lFornecedor
                
                lErro = CF("Fornecedor_Le", objForn)
                If lErro <> SUCESSO And lErro <> 12729 Then gError 196508
                
                objFilialEmpresa.iCodFilial = objNFsPag.iFilialEmpresa
                
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO Then gError 196509
                            
                TipoDoc.Caption = ""
                Status.Caption = ""

                Cliente.Caption = CStr(objForn.lCodigo) & SEPARADOR & objForn.sNomeReduzido
                Emissao.Caption = Format(objNFsPag.dtDataEmissao, "dd/mm/yyyy")
                Valor.Caption = Format(objNFsPag.dValorTotal, "STANDARD")
                FilialEmpresa.Caption = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
            
            Case Else
                gError 196503
            
        End Select
               
        DataCanc.PromptInclude = False
        DataCanc.Text = Format(StrParaDate(Emissao.Caption), "dd/mm/yy")
        DataCanc.PromptInclude = True
            
    End If
    
    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case gErr

        Case 196501, 196502, 196504, 196505, 196506, 196507, 196508, 196509
        
        Case 196503
            Call Rotina_Erro(vbOKOnly, "ERRO_FATURA_NAO_CADASTRADA", gErr)

        Case 196778
            Call Rotina_Erro(vbOKOnly, "ERRO_TRP_FATURA_JA_CANCELADA", gErr, Numero.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196510)

    End Select

    Exit Sub

End Sub

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 196511

    'Limpa Tela
    Call Limpa_Tela_CancelarFatura

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 196511

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196512)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTitRec As ClassTituloReceber
Dim objTitPag As ClassTituloPagar
Dim objNFsPag As ClassNFsPag
Dim objContabil As ClassContabil
Dim vbMsgRes As VbMsgBoxResult
Dim objFaturaTRP As New ClassFaturaTRP
Dim iFilialAux As Integer

On Error GoTo Erro_Gravar_Registro

    iFilialAux = giFilialEmpresa
    giFilialEmpresa = Codigo_Extrai(FilialEmpresa.Caption)

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Motivo.Text)) = 0 Then gError 196513
        
    'Verifica se é maior que o tamanho maximo
    If Len(Trim(Motivo.Text)) > STRING_MOTIVOCANCEL Then gError 196777
    
    objFaturaTRP.sMotivo = Motivo.Text
    gobjDestino.dtDataEstorno = StrParaDate(DataCanc.Text)
    
    Select Case giTipoDocDestino
    
        Case TRP_TIPO_DOC_DESTINO_TITREC
        
            Set objTitRec = gobjDestino
            
             Set objContabil = New ClassContabil
             Call objContabil.Contabil_Inicializa_Contabilidade4(23, MODULO_CONTASARECEBER)
             
            'Pede confirmação da exclusão
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TITULORECEBER", objTitRec.lNumTitulo)
            If vbMsgRes = vbNo Then gError 196514
            
            Set objTitRec.objInfoUsu = objFaturaTRP
    
            'Exclui o Titulo (inclusive a sua parte contábil)
            lErro = CF("TituloReceber_Exclui", objTitRec, objContabil)
            If lErro <> SUCESSO Then gError 196515
        
        Case TRP_TIPO_DOC_DESTINO_TITPAG
        
            Set objTitPag = gobjDestino
        
            Set objContabil = New ClassContabil
            Call objContabil.Contabil_Inicializa_Contabilidade4(1, MODULO_CONTASAPAGAR)
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_NFFATPAG", objTitPag.lNumTitulo)
            If vbMsgRes = vbNo Then gError 196516
            
            Set objTitPag.objInfoUsu = objFaturaTRP
            
            'Exclui Nota Fiscal Fatura (incluindo dados contabeis (contabilidade))
            lErro = CF("NFFatPag_Exclui", objTitPag, objContabil)
            If lErro <> SUCESSO Then gError 196517

        Case TRP_TIPO_DOC_DESTINO_NFSPAG
        
            Set objNFsPag = gobjDestino
        
            Set objContabil = New ClassContabil
            Call objContabil.Contabil_Inicializa_Contabilidade4(14, MODULO_CONTASAPAGAR)
        
            'Pede a confirmação de exclusão
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_NFPAG", objNFsPag.lNumNotaFiscal)
            If vbMsgRes = vbNo Then gError 196518
    
            Set objNFsPag.objInfoUsu = objFaturaTRP
    
            'Faz a exclusão da Nota Fiscal (inclusive dados contábeis)
            lErro = CF("NFPag_Exclui", objNFsPag, objContabil)
            If lErro <> SUCESSO Then gError 196519

        Case Else
            gError 196520
        
    End Select

    Motivo.Clear
    lErro = CF("Carrega_Combo_Historico", Motivo, "TRPTitulosExp", STRING_TRP_OCR_HISTORICO, "Motivo")
    If lErro <> SUCESSO Then gError 190165
    
    GL_objMDIForm.MousePointer = vbDefault
    
    giFilialEmpresa = iFilialAux
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    giFilialEmpresa = iFilialAux

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 196513
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_NAO_INFORMADO", gErr)
            
        Case 196514 To 196520, 190165
        
        Case 196777
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_MAIOR_QUE_MAXIMO", gErr, Len(Trim(Motivo.Text)), STRING_MOTIVOCANCEL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196521)

    End Select

    Exit Function

End Function

Private Sub DataCanc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataCanc_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataCanc, iAlterado)
    
End Sub

Private Sub DataCanc_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCanc_Validate

    If Len(Trim(DataCanc.ClipText)) <> 0 Then

        lErro = Data_Critica(DataCanc.Text)
        If lErro <> SUCESSO Then gError 197024

    End If
    
    Exit Sub

Erro_DataCanc_Validate:

    Cancel = True

    Select Case gErr

        Case 197024

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197025)

    End Select

    Exit Sub

End Sub

Private Sub UpDownCanc_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownCanc_DownClick

    DataCanc.SetFocus

    If Len(DataCanc.ClipText) > 0 Then

        sData = DataCanc.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 197026

        DataCanc.Text = sData

    End If

    Exit Sub

Erro_UpDownCanc_DownClick:

    Select Case gErr

        Case 197026

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197027)

    End Select

    Exit Sub

End Sub


Private Sub UpDownCanc_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownCanc_UpClick

    DataCanc.SetFocus

    If Len(Trim(DataCanc.ClipText)) > 0 Then

        sData = DataCanc.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 197028

        DataCanc.Text = sData

    End If

    Exit Sub

Erro_UpDownCanc_UpClick:

    Select Case gErr

        Case 197028

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197029)

    End Select

    Exit Sub

End Sub

Private Sub LblNumero_Click()

Dim lErro As Long
Dim objFat As New ClassFaturaTRP
Dim colSelecao As New Collection

On Error GoTo Erro_LblNumero_Click

    'Verifica se o Numero foi preenchido
    If Len(Trim(Numero.Text)) <> 0 Then

        objFat.lNumFat = Numero.Text

    End If

    Call Chama_Tela("TRPFaturasLista", colSelecao, objFat, objEventoNumero)

    Exit Sub

Erro_LblNumero_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197010)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFat As New ClassFaturaTRP

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objFat = obj1
    
    Numero.PromptInclude = False
    Numero.Text = CStr(objFat.lNumFat)
    Numero.PromptInclude = True

    Call Numero_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197009)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Numero Then Call LblNumero_Click
    
    End If
    
End Sub
