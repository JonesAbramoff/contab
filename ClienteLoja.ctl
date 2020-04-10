VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ClienteLoja 
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   8730
   Begin VB.Frame Frame2 
      Caption         =   "Endereço"
      Height          =   2355
      Left            =   60
      TabIndex        =   20
      Top             =   1740
      Width           =   8445
      Begin VB.ComboBox Pais 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3540
         TabIndex        =   23
         Top             =   630
         Width           =   1995
      End
      Begin VB.TextBox Endereco 
         Height          =   315
         Left            =   1245
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   210
         Width           =   7005
      End
      Begin VB.ComboBox Estado 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1245
         TabIndex        =   21
         Top             =   1050
         Width           =   630
      End
      Begin MSMask.MaskEdBox Cidade 
         Height          =   315
         Left            =   3540
         TabIndex        =   24
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CEP 
         Height          =   315
         Left            =   5910
         TabIndex        =   25
         Top             =   1065
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Email 
         Height          =   315
         Left            =   3540
         TabIndex        =   26
         Top             =   1890
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Contato 
         Height          =   315
         Left            =   5910
         TabIndex        =   27
         Top             =   1890
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fax 
         Height          =   315
         Left            =   3540
         TabIndex        =   28
         Top             =   1470
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Bairro 
         Height          =   315
         Left            =   1245
         TabIndex        =   29
         Top             =   630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Telefone1 
         Height          =   315
         Left            =   1245
         TabIndex        =   30
         Top             =   1470
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Telefone2 
         Height          =   315
         Left            =   1245
         TabIndex        =   31
         Top             =   1890
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
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
         Left            =   270
         TabIndex        =   42
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
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
         Left            =   2775
         TabIndex        =   41
         Top             =   1125
         Width           =   675
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
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
         Left            =   450
         TabIndex        =   40
         Top             =   1095
         Width           =   675
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
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
         Left            =   570
         TabIndex        =   39
         Top             =   675
         Width           =   585
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Telefone 1:"
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
         Left            =   195
         TabIndex        =   38
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Telefone 2:"
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
         Left            =   180
         TabIndex        =   37
         Top             =   1935
         Width           =   1005
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "e-mail:"
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
         Left            =   2850
         TabIndex        =   36
         Top             =   1935
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
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
         Left            =   3045
         TabIndex        =   35
         Top             =   1545
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
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
         Left            =   5430
         TabIndex        =   34
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
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
         TabIndex        =   33
         Top             =   1935
         Width           =   750
      End
      Begin VB.Label PaisLabel 
         AutoSize        =   -1  'True
         Caption         =   "País:"
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
         Index           =   1
         Left            =   2955
         TabIndex        =   32
         Top             =   660
         Width           =   495
      End
   End
   Begin VB.CommandButton Filiais 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6720
      Picture         =   "ClienteLoja.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   840
      Width           =   1620
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   690
      Left            =   90
      TabIndex        =   14
      Top             =   4230
      Width           =   8460
      Begin MSMask.MaskEdBox CGC 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   270
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "##############"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RG 
         Height          =   315
         Left            =   4950
         TabIndex        =   6
         Top             =   270
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Mask            =   "###############"
         PromptChar      =   " "
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF:"
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
         Left            =   765
         TabIndex        =   16
         Top             =   315
         Width           =   975
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "RG:"
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
         Left            =   4515
         TabIndex        =   15
         Top             =   330
         Width           =   345
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2565
      Picture         =   "ClienteLoja.ctx":0DA2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   255
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6405
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "ClienteLoja.ctx":0E8C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "ClienteLoja.ctx":100A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ClienteLoja.ctx":153C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "ClienteLoja.ctx":16C6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1725
      TabIndex        =   1
      Top             =   240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox RazaoSocial 
      Height          =   315
      Left            =   1710
      TabIndex        =   3
      Top             =   765
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   40
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Top             =   1290
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "BackOffice:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   3330
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   18
      Top             =   300
      Width           =   1020
   End
   Begin VB.Label CodigoBO 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4425
      TabIndex        =   17
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nome Reduzido:"
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
      Left            =   255
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   13
      Top             =   1335
      Width           =   1410
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
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
      Left            =   1110
      TabIndex        =   12
      Top             =   825
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Left            =   1005
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   300
      Width           =   660
   End
End
Attribute VB_Name = "ClienteLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControl As Object

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático do próximo cliente
    lErro = CF("ClienteLoja_Automatico", lCodigo)
    If lErro <> SUCESSO Then Error 57529

    'Exibe código na Tela
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57529
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154236)
    
    End Select

    Exit Sub

End Sub

Public Sub Bairro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim colCodNomeFiliais As New AdmColCodigoNome
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 And Len(CodigoBO.Caption) = 0 Then gError 12430

    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 And Len(CodigoBO.Caption) <> 0 Then gError 117605

    If Len(CodigoBO.Caption) <> 0 Then gError 117606

    objCliente.lCodigoLoja = StrParaLong(Codigo.Text)
    objCliente.lCodigo = StrParaLong(CodigoBO.Caption)
    objCliente.iFilialEmpresaLoja = giFilialEmpresa

    'Lê os dados do cliente a ser excluido
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 12431

    'Verifica se cliente não está cadastrado
    If lErro = 12293 Then gError 12432

    'Envia aviso perguntando se realmente deseja excluir cliente e suas filiais
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CLIENTE", objCliente.lCodigoLoja)

    If vbMsgRes = vbYes Then

        'Exclui Cliente
        lErro = CF("Cliente_Exclui", objCliente)
        If lErro <> SUCESSO Then gError 12467

        'Limpa a Tela
        lErro = Limpa_Tela_Clientes()
        If lErro <> SUCESSO Then gError 58587

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 12431, 12467, 58587
        
        Case 12430
            Call Rotina_Erro(vbOKOnly, "ERRO_CODCLIENTE_NAO_PREENCHIDO", gErr)

        Case 12432
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigoLoja)

        Case 117605
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_CODIGO_JA_PREENCHIDO", gErr)

        Case 117606
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CLIENTE_TRANSFERIDO", gErr, CodigoBO.Caption)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154237)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub CEP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CEP_GotFocus()
    
    If Me.ActiveControl Is CEP Then
        Call MaskEdBox_TrataGotFocus(CEP, iAlterado)
    End If

End Sub

'Private Sub Desmembrar_Click()
'
'Dim objLog As New ClassLog
'Dim objClientes As ClassCliente
'Dim colEnderecos As New Collection
'
'    Call Limpa_Tela_Clientes
'
'    Call Log_Le(objLog)
'
'    Call Cliente_Desmembra_Log(objClientes, colEnderecos, objLog)
'
'    Call Exibe_Dados_Cliente(objClientes)
'
'End Sub

Function Cliente_Desmembra_Log(objCliente As ClassCliente, colEnderecos As Collection, objLog As ClassLog) As Long
'Função que informações do banco de Dados e Carrega no Obj

Dim lErro As Long


Dim iPosicao1 As Integer
Dim iPosicao2 As Integer
Dim iPosicao3 As Integer
Dim iPosicao4 As Integer
Dim iPosicao5 As Integer
Dim iPosicaoCol As Integer
Dim sCliente As String
Dim iIndice As Integer
Dim bFim As Boolean
Dim objFilialCliCategoria As ClassFilialCliCategoria
Dim objEndereco As ClassEndereco

On Error GoTo Erro_Cliente_Desmembra_Log
    
    'iPosicao4 Guarda o Final da String
    iPosicao4 = InStr(1, objLog.sLog, Chr(vbKeyEnd))
    
    'inicia a col de endereços
    iPosicaoCol = InStr(1, objLog.sLog, Chr(vbKeySeparator))
        
    'iPosicao1 Guarda a posição do Primeiro Control
    iPosicao1 = InStr(1, objLog.sLog, Chr(vbKeyControl))
    If iPosicao1 = 0 Then iPosicao1 = iPosicaoCol
    
    'String que Guarda as Propriedades do Objtecladoproduto
    sCliente = Mid(objLog.sLog, 1, iPosicao1 - 1)
        
    'Variável booleana que funcionará como Flag
    bFim = True
    'Inicilalização do objcliente
    Set objCliente = New ClassCliente
     
    'Primeira Posição
    iPosicao3 = 1
    
    'Procura o Primeiro Escape dentro da String stecladoproduto e Armazena a Posição
    iPosicao2 = (InStr(iPosicao3, sCliente, Chr(vbKeyEscape)))
    
    iIndice = 0
    
    sCliente = sCliente & Chr(vbKeyEscape)
    
    Do While iPosicao2 <> 0
        
       iIndice = iIndice + 1
        'Recolhe os Dados do Banco de Dados e Coloca no objtecladoproduto
        Select Case iIndice
        
            Case 1: objCliente.dComissaoVendas = StrParaDbl(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 2: objCliente.dDesconto = StrParaDbl(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 3: objCliente.dLimiteCredito = StrParaDbl(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 4: objCliente.dMediaCompra = StrParaDbl(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 5: objCliente.dSaldoAtrasados = StrParaDbl(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 6: objCliente.dSaldoDuplicatas = StrParaDbl(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 7: objCliente.dSaldoPedidosLiberados = StrParaDbl(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 8: objCliente.dSaldoTitulos = StrParaDbl(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 9: objCliente.dtDataPrimeiraCompra = StrParaDate(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 10: objCliente.dtDataUltChequeDevolvido = StrParaDate(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 11: objCliente.dtDataUltimaCompra = StrParaDate(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 12: objCliente.dtDataUltimoProtesto = StrParaDate(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 13: objCliente.dtDataUltVisita = StrParaDate(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 14: objCliente.dValorAcumuladoCompras = StrParaDbl(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 15: objCliente.dValPagtosAtraso = StrParaDbl(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 16: objCliente.iCodCobrador = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 17: objCliente.iCodMensagem = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 18: objCliente.iCodPadraoCobranca = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 19: objCliente.iCodTransportadora = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 20: objCliente.iCondicaoPagto = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 21: objCliente.iFreqVisitas = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 22: objCliente.iNumChequesDevolvidos = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 23: objCliente.iProxCodFilial = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 24: objCliente.iRegiao = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 25: objCliente.iTabelaPreco = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 26: objCliente.iTipo = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 27: objCliente.iTipoFrete = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 28: objCliente.iVendedor = StrParaInt(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 29: objCliente.lCodigo = StrParaLong(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 30: objCliente.lEndereco = StrParaLong(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 31: objCliente.lEnderecoCobranca = StrParaLong(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 32: objCliente.lEnderecoEntrega = StrParaLong(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 33: objCliente.lMaiorAtraso = StrParaLong(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 34: objCliente.lMediaAtraso = StrParaLong(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 35: objCliente.lNumeroCompras = StrParaLong(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 36: objCliente.lNumPagamentos = StrParaLong(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 37: objCliente.lNumTitulosProtestados = StrParaLong(Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3))
            Case 38: objCliente.sCgc = Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3)
            Case 39: objCliente.sContaContabil = Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3)
            Case 40: objCliente.sInscricaoEstadual = Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3)
            Case 41: objCliente.sInscricaoMunicipal = Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3)
            Case 42: objCliente.sInscricaoSuframa = Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3)
            Case 43: objCliente.sNomeReduzido = Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3)
            Case 44: objCliente.sObservacao = Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3)
            Case 45: objCliente.sObservacao2 = Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3)
            Case 46: objCliente.sRazaoSocial = Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3)
            Case 47: objCliente.sRG = Mid(sCliente, iPosicao3, iPosicao2 - iPosicao3)
            Case 48: Exit Do
        
        End Select
        'Atualiza as Posições
        iPosicao3 = iPosicao2 + 1
        iPosicao2 = (InStr(iPosicao3, sCliente, Chr(vbKeyEscape)))
    
    Loop
        
    iPosicao3 = iPosicao1 + 1
    
    Do While bFim <> False
                  
        'iPosicao1 Guarda a posição do Control Ponto Inicial
        iPosicao1 = InStr(iPosicao3, objLog.sLog, Chr(vbKeyControl))
        
        If iPosicao1 = 0 Then Exit Do
     
        'Atualiza as Posições
        iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
                
        'Atualiza o valor de Indice
        iIndice = 0
        
        'inicia o objtecladoprodutoCondPagto para receber os dados do Banco de Dados
        Set objFilialCliCategoria = New ClassFilialCliCategoria
        
        Do While iPosicao2 > iPosicao3
        
            iIndice = iIndice + 1
            
            'Recolhe os Dados do Banco de Dados e Coloca no objtecladoprodutoitens
            Select Case iIndice
            
                Case 1: objFilialCliCategoria.iFilial = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - (iPosicao3)))
                Case 2: objFilialCliCategoria.lCliente = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
                Case 3: objFilialCliCategoria.sCategoria = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 4: objFilialCliCategoria.sItem = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 5: Exit Do
        
            End Select
            
            'Atualiza as Posições
            iPosicao3 = iPosicao2 + 1
            iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
                
            'Verfica se a Posição que é atualizada no decorrer da função é maior que a posição da Flag End
            If (iPosicao2 > iPosicao1) Or iPosicao2 = 0 Then
                'A flag Fim Recebe False
                iPosicao2 = iPosicao1
            End If
            
        Loop
        
        objCliente.colCategoriaItem.Add objFilialCliCategoria
        
        'Verfica se a Posição que é atualizada no decorrer da função é maior que a posição da Flag End
        If iPosicao3 > iPosicao4 Then
            'A flag Fim Recebe False
            bFim = False
        End If
        
    Loop
    
    iPosicao3 = iPosicaoCol + 1
    
    Do While bFim <> False
                  
        'iPosicao1 Guarda a posição dos outros separador
        iPosicao1 = InStr(iPosicao3, objLog.sLog, Chr(vbKeySeparator))
        
        If iPosicao1 = 0 Then iPosicao1 = iPosicao4
     
        'Atualiza as Posições
        iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
                
        'Atualiza o valor de Indice
        iIndice = 0
        
        'inicia o objtecladoprodutoCondPagto para receber os dados do Banco de Dados
        Set objEndereco = New ClassEndereco
        
        Do While iPosicao2 <= iPosicao1
        
            iIndice = iIndice + 1
            
            'Recolhe os Dados do Banco de Dados e Coloca no objtecladoprodutoitens
            Select Case iIndice
            
                Case 1: objEndereco.iCodigoPais = StrParaInt(Mid(objLog.sLog, iPosicao3, iPosicao2 - (iPosicao3)))
                Case 2: objEndereco.lCodigo = StrParaLong(Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3))
                Case 3: objEndereco.sBairro = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 4: objEndereco.sCEP = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 5: objEndereco.sCidade = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 6: objEndereco.sContato = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 7: objEndereco.sEmail = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 8: objEndereco.sEndereco = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 9: objEndereco.sFax = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 10: objEndereco.sSiglaEstado = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 11: objEndereco.sTelefone1 = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 12: objEndereco.sTelefone2 = Mid(objLog.sLog, iPosicao3, iPosicao2 - iPosicao3)
                Case 13: Exit Do
        
            End Select
            
            'Atualiza as Posições
            iPosicao3 = iPosicao2 + 1
            iPosicao2 = (InStr(iPosicao3, objLog.sLog, Chr(vbKeyEscape)))
                
            'Verfica se a Posição que é atualizada no decorrer da função é maior que a posição da Flag End
            If (iPosicao2 > iPosicao1) Or iPosicao2 = 0 Then
                'A flag Fim Recebe False
                iPosicao2 = iPosicao1
            End If
            
        Loop
        
        colEnderecos.Add objEndereco
                    
        If colEnderecos.Count = 1 Then
            objCliente.lEndereco = objEndereco.lCodigo
        Else
            If colEnderecos.Count = 2 Then
                objCliente.lEnderecoCobranca = objEndereco.lCodigo
            Else
                objCliente.lEnderecoEntrega = objEndereco.lCodigo
            End If
        End If
        
        'Verfica se a Posição que é atualizada no decorrer da função é maior que a posição da Flag End
        If iPosicao3 > iPosicao4 Then
            'A flag Fim Recebe False
            bFim = False
        End If
        
    Loop
        
    Cliente_Desmembra_Log = SUCESSO

    Exit Function

Erro_Cliente_Desmembra_Log:

    Cliente_Desmembra_Log = gErr

    Select Case gErr

        Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154238)

        End Select

    
    Exit Function

End Function

Function Log_Le(objLog As ClassLog) As Long

Dim lErro As Long
Dim tLog As typeLog
Dim lComando As Long

On Error GoTo Erro_Log_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104197

    'Inicializa o Buffer da Variáveis String
    tLog.sLog1 = String(STRING_CONCATENACAO, 0)
    tLog.sLog2 = String(STRING_CONCATENACAO, 0)
    tLog.sLog3 = String(STRING_CONCATENACAO, 0)
    tLog.sLog4 = String(STRING_CONCATENACAO, 0)

    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Operacao, Log1, Log2, Log3, Log4 , Data , Hora FROM Log ORDER BY  NumIntDoc DESC", tLog.lNumIntDoc, tLog.iOperacao, tLog.sLog1, tLog.sLog2, tLog.sLog3, tLog.sLog4, tLog.dtData, tLog.dHora)
    If lErro <> SUCESSO Then gError 104198

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104199


    If lErro = AD_SQL_SUCESSO Then

        'Carrega o objLog com as Infromações de bonco de dados
        objLog.lNumIntDoc = tLog.lNumIntDoc
        objLog.iOperacao = tLog.iOperacao
        objLog.sLog = tLog.sLog1 & tLog.sLog2 & tLog.sLog3 & tLog.sLog4
        objLog.dtData = tLog.dtData
        objLog.dHora = tLog.dHora

    End If

    If lErro = AD_SQL_SEM_DADOS Then gError 104202
    
    Log_Le = SUCESSO

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

Erro_Log_Le:

    Log_Le = gErr

   Select Case gErr

    Case gErr

        Case 104198, 104199
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOG", gErr)
    
        Case 104202
            Call Rotina_Erro(vbOKOnly, "ERRO_LOG_NAO_EXISTENTE", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154239)

        End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub Filiais_Click()

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filiais_Click

    'Verifica se foi preenchido o Codigo
    If Len(Trim(Codigo.Text)) = 0 Then gError 112635

    'Preenche objCliente
    objCliente.lCodigoLoja = CLng(Codigo.Text)
    
    'Lê o Cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 112636

    'Se não achou o Cliente
    If lErro <> SUCESSO Then

        'Envia aviso perguntando se deseja cadastrar novo Cliente
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE")

        If vbMsgRes = vbYes Then
            
            'Grava o novo cliente
            lErro = Gravar_Registro()
            If lErro <> SUCESSO Then gError 112637
            
            'Chama a Tela de Filiais de Cliente
            objFilialCliente.lCodClienteLoja = CLng(Codigo.Text)
            objFilialCliente.iCodFilialLoja = 1  'p/começar exibindo a matriz
            
            Call Chama_Tela("FiliaisClientesLoja", objFilialCliente)
        
        End If
    Else
    
        'Chama a Tela de Filiais de Cliente
        objFilialCliente.lCodClienteLoja = CLng(Codigo.Text)
        objFilialCliente.iCodFilialLoja = 1  'p/começar exibindo a matriz
        
        Call Chama_Tela("FiliaisClientesLoja", objFilialCliente)
    
    End If

    Exit Sub

Erro_Filiais_Click:

    Select Case gErr

        Case 112635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODCLIENTE_NAO_PREENCHIDO", Err)

        Case 112636, 112637

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154240)

    End Select

    Exit Sub
    
End Sub

Public Sub RG_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub RG_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(RG, iAlterado)

End Sub

Public Sub CGC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)

End Sub

Public Sub Cidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se foi preenchido o campo Codigo Cliente
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Critica se é um Long
    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 19299

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 19299

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154241)

    End Select

    Exit Sub

End Sub

Public Sub Contato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Email_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Endereco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Estado_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Fax_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o Cliente
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 12360

    'Limpa a Tela
    lErro = Limpa_Tela_Clientes()
    If lErro <> SUCESSO Then Error 58588
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 12360, 58588

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154242)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Verifica se dados de Cliente necessários foram preenchidos
'Grava Cliente no BD
'Atualiza ListBox de Clientes

Dim lErro As Long
Dim iIndice As Integer
Dim objCliente As New ClassCliente
Dim colEndereco As New Collection
Dim iLoja As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'somente um dos codigos precisa estar preenchido.
    
    'Verifica se foi preenchido o Código e o Codigo do Backoffice esta vazio
    If Len(Trim(Codigo.Text)) = 0 And Len(CodigoBO.Caption) = 0 Then gError 12361
    
    'Verifica se foi preenchido o Código e o Codigo do Backoffice simultaneamente
    If Len(Trim(Codigo.Text)) <> 0 And Len(CodigoBO.Caption) <> 0 Then gError 117597
    
    'Verifica se foi preenchida a Razao Social
    If Len(Trim(RazaoSocial.Text)) = 0 Then gError 12362

    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 12363
    
    'Verifica se foi preenchido o Estado dos Endereços
    If Len(Trim(Endereco.Text)) <> 0 Then
        If Len(Trim(Estado.Text)) = 0 Then gError 43290
    End If
    
    'Lê os dados dos Enderecos e coloca em colEndereco
    lErro = Le_Dados_Enderecos(colEndereco)
    If lErro <> SUCESSO Then gError 12507

    'Lê os dados da Tela relacionados ao Cliente
    lErro = Le_Dados_Cliente(objCliente)
    If lErro <> SUCESSO Then gError 43293
    
    'Se o CGC estiver Preenchido
    If Len(Trim(objCliente.sCgc)) > 0 Then
    
        If objCliente.lCodigo = 0 Then
    
            iLoja = 1
    
            'Verifica se tem outro Cliente com o mesmo CGC e dá aviso
            lErro = CF("FilialCliente_Testa_CGC", objCliente.lCodigoLoja, 0, objCliente.sCgc, iLoja)
            If lErro <> SUCESSO Then gError 126017
            
        Else
            
            'Verifica se tem outro Cliente com o mesmo CGC e dá aviso
            lErro = CF("FilialCliente_Testa_CGC", objCliente.lCodigoLoja, 0, objCliente.sCgc)
            If lErro <> SUCESSO Then gError 58615
            
        End If
            
            
    End If
    
    'Grava o Cliente no BD
    lErro = CF("Cliente_Grava", objCliente, colEndereco)
    If lErro <> SUCESSO Then gError 43294

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 12361
            Call Rotina_Erro(vbOKOnly, "ERRO_CODCLIENTE_NAO_PREENCHIDO", gErr)

        Case 12362
            Call Rotina_Erro(vbOKOnly, "ERRO_RAZ_SOC_NAO_PREENCHIDA", gErr)

        Case 12363
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)

        Case 12507, 43293, 43294, 58615, 126017

        Case 43290, 43291, 43292
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_PREENCHIDO", gErr)

        Case 117597
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_CODIGO_JA_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154243)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 12484

    'Limpa a Tela
    lErro = Limpa_Tela_Clientes()
    If lErro <> SUCESSO Then Error 58589
    
    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case Err

        Case 12484, 58589

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154244)

    End Select

End Sub

Public Sub CGC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC_Validate
    
    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGC.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(CGC.Text))

        Case STRING_CPF 'CPF
            
            'Critica Cpf
            lErro = Cpf_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 12316
            
            'Formata e coloca na Tela
            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text

        Case STRING_CGC 'CGC
            
            'Critica CGC
            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 12317
            
            'Formata e Coloca na Tela
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text

        Case Else
                
            Error 12318

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True


    Select Case Err

        Case 12316, 12317

        Case 12318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154245)

    End Select


    Exit Sub

End Sub

Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub RazaoSocial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Estado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Estado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Estado_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(Estado.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Estado
    If Estado.Text = Estado.List(Estado.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na Combo Estado, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(Estado)
    If lErro <> SUCESSO And lErro <> 58583 Then Error 12324

    'Não existe o ítem na ComboBox Estado
    If lErro = 58583 Then Error 12325

    Exit Sub

Erro_Estado_Validate:

    Cancel = True


    Select Case Err

        Case 12324

        Case 12325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", Err, Estado.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154246)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Inicializa variávies AdmEvento
    Set objEventoCliente = New AdmEvento
    
    RazaoSocial.MaxLength = STRING_CLIENTE_RAZAO_SOCIAL
    Endereco.MaxLength = STRING_ENDERECO
    Bairro.MaxLength = STRING_BAIRRO
    Cidade.MaxLength = STRING_CIDADE
    'Implementado pois agora é possível ter constantes cutomizadas em função de tamanhos de campos do BD. AdmLib.ClassConsCust
    Telefone1.MaxLength = STRING_TELEFONE
    Telefone2.MaxLength = STRING_TELEFONE
    Fax.MaxLength = STRING_FAX
    Email.MaxLength = STRING_EMAIL
    Contato.MaxLength = STRING_CONTATO
    
    
    'Prepara as Combos do Tab de Endereco
    lErro = Inicializa_Tab_Enderecos()
    If lErro <> SUCESSO Then Error 58086
    
'    'carrega a combo de tipos de clientes
'    lErro = Carrega_Tipo()
'    If lErro <> SUCESSO Then gError 113815
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
        
        Case 58084, 58085, 58086, 58087, 58505
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154247)

    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

'Public Sub Tipo_Change()
'    iAlterado = REGISTRO_ALTERADO
'End Sub
'
'Public Sub Tipo_Click()
'    iAlterado = REGISTRO_ALTERADO
'End Sub
'
'Public Sub Tipo_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objTipoCliente As New ClassTipoCliente
'Dim vbMsgRes As VbMsgBoxResult
'Dim iCodigo As Integer
'
'On Error GoTo Erro_Tipo_Validate
'
'    'Verifica se foi preenchida a ComboBox Tipo
'    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub
'
'    'Verifica se está preenchida com o ítem selecionado na ComboBox Tipo
'    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then Exit Sub
'
'    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
'    lErro = Combo_Seleciona(Tipo, iCodigo)
'    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 113817
'
'    'Não existe o ítem com o CÓDIGO na List da ComboBox
'    If lErro = 6730 Then
'
'        objTipoCliente.iCodigo = iCodigo
'
'        'Tenta ler TipoCliente com esse código no BD
'        lErro = CF("TipoDeCliente_Le", objTipoCliente)
'        If lErro <> SUCESSO And lErro <> 28943 Then gError 113818
'
'        'Não encontrou Tipo Cliente no BD
'        If lErro = 28943 Then gError 113819
'
'        'Exibe dados de TipoCliente na tela
'        Tipo.Text = CStr(iCodigo) & SEPARADOR & objTipoCliente.sDescricao
'
'    End If
'
'    'Não existe o ítem com a STRING na List da ComboBox
'    If lErro = 6731 Then gError 113821
'
'    Exit Sub
'
'Erro_Tipo_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 113820, 113817, 113818  'Já tratado na rotina chamada
'
'        Case 113819 'Não encontrou Tipo Cliente no BD
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOCLIENTE")
'
'            If vbMsgRes = vbYes Then
'
'                'Chama a tela de TiposDeClientes
'                Call Chama_Tela("TipoCliente", objTipoCliente)
'
'            End If
'
'        Case 113821
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCLIENTE_INEXISTENTE", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154248)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Function Carrega_Tipo() As Long
'
'Dim lErro As Long
'Dim colCodigoDescricao As New AdmColCodigoNome
'Dim objCodigoDescricao As AdmCodigoNome
'
'On Error GoTo Erro_Carrega_Tipo
'
'    'Lê cada código e descrição da tabela TiposDeCliente
'    lErro = CF("Cod_Nomes_Le", "TiposDeCliente", "Codigo", "Descricao", STRING_TIPO_CLIENTE_DESCRICAO, colCodigoDescricao)
'    If lErro <> SUCESSO Then Error 113816
'
'    'Preenche a ComboBox Tipo com os objetos da colecao colCodigoDescricao
'    For Each objCodigoDescricao In colCodigoDescricao
'        Tipo.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
'        Tipo.ItemData(Tipo.NewIndex) = objCodigoDescricao.iCodigo
'    Next
'
'    Carrega_Tipo = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Tipo:
'
'    Carrega_Tipo = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154249)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Inicializa_Tab_Enderecos() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Inicializa_Tab_Enderecos
    
    'Lê cada codigo da tabela Estados
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then Error 12268

    'Preenche as ComboBox Estados com os objetos da colecao colCodigo
    For Each vCodigo In colCodigo
            Estado.AddItem vCodigo
    Next

    'Preenche Combos de País
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela Paises
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 12269

    'Percorre as 3 Combos de País
    'Preenche cada ComboBox País com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Pais.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Pais.ItemData(Pais.NewIndex) = objCodigoDescricao.iCodigo
    Next

        'Seleciona Brasil se existir
    For iIndice2 = 0 To Pais.ListCount - 1
        If Right(Pais.List(iIndice2), 6) = "Brasil" Then
            Pais.ListIndex = iIndice2
            Exit For
        End If
    Next

    Inicializa_Tab_Enderecos = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Tab_Enderecos:

    Inicializa_Tab_Enderecos = Err
    
    Select Case Err
        
        Case 12268, 12269
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154250)
        
    End Select
        
    Exit Function
    
End Function

Public Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NomeReduzido_Validate(Cancel As Boolean)

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate

    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then Error 25001

    End If

    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True


    Select Case Err

        Case 25001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154251)

    End Select

End Sub

Public Sub Pais_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Pais_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Pais_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPais As New ClassPais

On Error GoTo Erro_Pais_Validate

    'Verifica se foi preenchida a Combo Pais
    If Len(Trim(Pais.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Pais
    If Pais.Text = Pais.List(Pais.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Pais, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 12326

    'Nao existe o item com o CODIGO na List da ComboBox
    If lErro = 6730 Then

        objPais.iCodigo = iCodigo
        
        'Tenta ler Pais com esse codigo no BD
        lErro = CF("Paises_Le", objPais)
        If lErro <> SUCESSO And lErro <> 47876 Then gError 76427
        If lErro <> SUCESSO Then gError 76428

        Pais.Text = CStr(iCodigo) & SEPARADOR & objPais.sNome
        
    End If

    'Nao existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 6900

    Exit Sub

Erro_Pais_Validate:

    Cancel = True


    Select Case gErr

        Case 6900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, Trim(Pais.Text))

        Case 12326, 76427

        Case 76428
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PAIS", objPais.iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("Paises", objPais)
            Else
                'Segura o foco
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154252)

    End Select

    Exit Sub

End Sub

Public Sub Telefone1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Telefone2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Function Limpa_Tela_Clientes() As Long

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Clientes

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa as TextBox e as MaskedEditBox
    Call Limpa_Tela(Me)

    'Seleciona Brasil nas Combos País
    For iIndice2 = 0 To Pais.ListCount - 1
        If Right(Pais.List(iIndice2), 6) = "Brasil" Then
            Pais.ListIndex = iIndice2
            Exit For
        End If
    Next

    'Limpa o código de identificação
    CodigoBO.Caption = ""
    Estado.ListIndex = -1
    
'    Tipo.ListIndex = -1
    
    Limpa_Tela_Clientes = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_Clientes:

    Limpa_Tela_Clientes = Err
    
    Select Case Err
        
        Case 58584
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154253)

    End Select
    
    Exit Function
        
End Function

Function Trata_Parametros(Optional objCliente As ClassCliente) As Long

Dim lErro As Long
Dim objClienteEstatistica As New ClassFilialClienteEst

On Error GoTo Erro_Trata_Parametros

    'Se houver Cliente passado como parâmetro, exibe seus dados
    If Not (objCliente Is Nothing) Then

        'Se Codigo é positivo
        If objCliente.lCodigoLoja > 0 Then

            'Lê Cliente no BD a partir do código
            lErro = CF("Cliente_Le_Estendida", objCliente, objClienteEstatistica)
            If lErro <> SUCESSO And lErro <> 52545 Then Error 12283

            'Se não encontrou o Cliente no BD
            If lErro <> SUCESSO Then

                'Limpa a Tela e exibe apenas o código
                lErro = Limpa_Tela_Clientes()
                If lErro <> SUCESSO Then Error 58585
                
                Codigo.Text = CStr(objCliente.lCodigoLoja)

            Else  'Encontrou Cliente no BD

                'Exibe os dados do Cliente
                lErro = Exibe_Dados_Cliente(objCliente)
                If lErro <> SUCESSO Then Error 12500

            End If

        'se Nome Reduzido está preenchido
        ElseIf Len(Trim(objCliente.sNomeReduzido)) > 0 Then

            'Lê Cliente no BD a partir do Nome Reduzido
            lErro = CF("Cliente_Le_NomeRed_Estendida", objCliente, objClienteEstatistica)
            If lErro <> SUCESSO And lErro <> 52693 Then Error 6923

            'Se não encontrou o Cliente no BD
            If lErro <> SUCESSO Then

                'Limpa a Tela e exibe apenas o NomeReduzido
                lErro = Limpa_Tela_Clientes()
                If lErro <> SUCESSO Then Error 58586
                
                NomeReduzido.Text = CStr(objCliente.sNomeReduzido)

            Else  'Encontrou Cliente no BD

                'Exibe os dados do Cliente
                lErro = Exibe_Dados_Cliente(objCliente)
                If lErro <> SUCESSO Then Error 6924

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 6923, 6924, 12283, 12500, 58585, 58586  'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154254)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Exibe_Dados_Cliente(objCliente As ClassCliente) As Long
'Exibe os dados de Cliente na tela

Dim lErro As Long
Dim iIndice As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Exibe_Dados_Cliente
    
    'TAB IDENTIFICACAO :
    lErro = Exibe_Dados_Cliente_Identificacao(objCliente)
    If lErro <> SUCESSO Then Error 58095
    
    'TAB INSCRICOES :
    lErro = Exibe_Dados_Cliente_Inscricoes(objCliente)
    If lErro <> SUCESSO Then Error 58095
    
    'TAB ENDERECOS :
    lErro = Exibe_Dados_Cliente_Enderecos(objCliente)
    If lErro <> SUCESSO Then Error 58096
    
    iAlterado = 0

    Exibe_Dados_Cliente = SUCESSO

    Exit Function

Erro_Exibe_Dados_Cliente:
    
    Exibe_Dados_Cliente = Err

    Select Case Err
        
        Case 58095, 58096, 58097
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154255)

    End Select

    Exit Function

End Function

Private Function Exibe_Dados_Cliente_Identificacao(objCliente As ClassCliente) As Long

On Error GoTo Erro_Exibe_Dados_Cliente_Identificacao

    Codigo.Text = IIf(objCliente.lCodigoLoja <> 0, CStr(objCliente.lCodigoLoja), "")
    CodigoBO.Caption = IIf(objCliente.lCodigo <> 0, CStr(objCliente.lCodigo), "")
    RazaoSocial.Text = objCliente.sRazaoSocial
    NomeReduzido.Text = objCliente.sNomeReduzido
'    Tipo.Text = objCliente.iTipo
'    Call Tipo_Validate(bSGECancelDummy)
    
    Exibe_Dados_Cliente_Identificacao = Err
    
    Exit Function
    
Erro_Exibe_Dados_Cliente_Identificacao:

    Exibe_Dados_Cliente_Identificacao = SUCESSO
    
    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154256)

    End Select

    Exit Function

End Function

Private Function Exibe_Dados_Cliente_Inscricoes(objCliente As ClassCliente) As Long
'Exibe as inscrições do Cliente
Dim objFilialCliente As New ClassFilialCliente
Dim lErro As Long

On Error GoTo Erro_Exibe_Dados_Cliente_Inscricoes
    
    'Inicializa objFilialCliente
    objFilialCliente.lCodClienteLoja = objCliente.lCodigoLoja
    objFilialCliente.iCodFilialLoja = FILIAL_MATRIZ
    objFilialCliente.iFilialEmpresaLoja = giFilialEmpresa
    
    'Lê o restante dos dados do Cliente na tabela de Filiais
    lErro = CF("FilialCliente_Le_Loja", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 112607 Then Error 19211
        
    RG.Text = objFilialCliente.sRG
    CGC.Text = objFilialCliente.sCgc
    Call CGC_Validate(bSGECancelDummy)
    
    Exibe_Dados_Cliente_Inscricoes = Err
    
    Exit Function
    
Erro_Exibe_Dados_Cliente_Inscricoes:

    Exibe_Dados_Cliente_Inscricoes = SUCESSO
    
    Select Case Err
        
        Case 19211
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154257)

    End Select

    Exit Function
    
End Function

Private Function Exibe_Dados_Cliente_Enderecos(objCliente As ClassCliente) As Long
'Exibe os Endereços do Cliente

Dim lErro As Long
Dim iIndice As Integer
Dim objEndereco As ClassEndereco
Dim colEnderecos As New colEndereco

On Error GoTo Erro_Exibe_Dados_Cliente_Enderecos

    'Lê os dados dos tres tipos de enderecos
    lErro = CF("Enderecos_Le_Cliente", colEnderecos, objCliente)
    If lErro <> SUCESSO Then Error 12304
    
    Set objEndereco = colEnderecos.Item(1)
    
    Endereco.Text = objEndereco.sEndereco
    Bairro.Text = objEndereco.sBairro
    Cidade.Text = objEndereco.sCidade
    CEP.Text = objEndereco.sCEP
    Estado.Text = objEndereco.sSiglaEstado
    If objEndereco.iCodigoPais = 0 Then
        Pais.Text = ""
    Else
        Pais.Text = objEndereco.iCodigoPais
        Call Pais_Validate(bSGECancelDummy)
    End If
    Telefone1.Text = objEndereco.sTelefone1
    Telefone2.Text = objEndereco.sTelefone2
    Fax.Text = objEndereco.sFax
    Email.Text = objEndereco.sEmail
    Contato.Text = objEndereco.sContato
    
    Exibe_Dados_Cliente_Enderecos = SUCESSO
    
    Exit Function
    
Erro_Exibe_Dados_Cliente_Enderecos:
    
    Exibe_Dados_Cliente_Enderecos = Err
    
    Select Case Err
            
        Case 12304
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154258)

    End Select

    Exit Function

End Function

Private Function Le_Dados_Cliente(objCliente As ClassCliente) As Long
'Lê os dados que estão na tela de Clientes e coloca em objCliente

Dim lErro As Long

On Error GoTo Erro_Le_Dados_Cliente

    'IDENTIFICACAO :

    objCliente.lCodigoLoja = StrParaLong(Codigo.Text)
    objCliente.lCodigo = StrParaLong(CodigoBO.Caption)

    objCliente.sRazaoSocial = Trim(RazaoSocial.Text)
    objCliente.sNomeReduzido = Trim(NomeReduzido.Text)
    objCliente.iFilialEmpresaLoja = giFilialEmpresa
    
    'INSCRICOES :
    
    objCliente.sRG = Trim(RG.Text)
    objCliente.sCgc = Trim(CGC.Text)
    
    Le_Dados_Cliente = SUCESSO

    Exit Function

Erro_Le_Dados_Cliente:

    Le_Dados_Cliente = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154259)

    End Select

    Exit Function

End Function

Private Function Le_Dados_Enderecos(colEndereco As Collection) As Long
'Lê os dados relativos ao endereco(0 = principal , 1 = entrega , 2 = cobranca) do cliente e coloca em colEndereco

Dim iIndice As Integer
Dim objEndereco As ClassEndereco
Dim iEstadoPreenchido As Integer

    'Verifica se tem algum estado Preenchido
    If Len(Trim(Estado.Text)) > 0 Then
        iEstadoPreenchido = iIndice
    End If
    
    Set objEndereco = New ClassEndereco

    objEndereco.sEndereco = Trim(Endereco.Text)
    objEndereco.sBairro = Trim(Bairro.Text)
    objEndereco.sCidade = Trim(Cidade.Text)
    objEndereco.sCEP = Trim(CEP.Text)
    
    'Se o Endereco não estiver Preenchido --> Seta o Estado que esta Preenchido em Algum dos Frames
    If Len(Trim(Endereco.Text)) > 0 Then
        objEndereco.sSiglaEstado = Trim(Estado.Text)
    Else
        objEndereco.sSiglaEstado = Trim(Estado.Text)
    End If
    
    If Len(Trim(Pais.Text)) = 0 Then
        objEndereco.iCodigoPais = 0
    Else
        objEndereco.iCodigoPais = Codigo_Extrai(Pais.Text)
    End If

    objEndereco.sTelefone1 = Trim(Telefone1.Text)
    objEndereco.sTelefone2 = Trim(Telefone2.Text)
    objEndereco.sFax = Trim(Fax.Text)
    objEndereco.sEmail = Trim(Email.Text)
    objEndereco.sContato = Trim(Contato.Text)

    colEndereco.Add objEndereco
    
    Le_Dados_Enderecos = SUCESSO

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Clientes"

    'Lê os dados da Tela Clientes
    lErro = Le_Dados_Cliente(objCliente)
    If lErro <> SUCESSO Then Error 12584

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodigoLoja", objCliente.lCodigoLoja, 0, "CodigoLoja"
    colCampoValor.Add "Codigo", objCliente.lCodigo, 0, "Codigo"
    colCampoValor.Add "RazaoSocial", objCliente.sRazaoSocial, STRING_CLIENTE_RAZAO_SOCIAL, "RazaoSocial"
    colCampoValor.Add "NomeReduzido", objCliente.sNomeReduzido, STRING_CLIENTE_NOME_REDUZIDO, "NomeReduzido"
   
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 12584

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154260)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objClienteEstatistica As New ClassFilialClienteEst

On Error GoTo Erro_Tela_Preenche

    objCliente.lCodigoLoja = colCampoValor.Item("CodigoLoja").vValor
    objCliente.lCodigo = colCampoValor.Item("Codigo").vValor
    objCliente.iFilialEmpresaLoja = giFilialEmpresa

    'Lê o Cliente no BD
    lErro = CF("Cliente_Le_Estendida", objCliente, objClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 52545 Then gError 71923

    'Se cliente não está cadastrado, erro
    If lErro = 12293 Then gError 71925
    
    'Exibe o Cliente na Tela
    lErro = Exibe_Dados_Cliente(objCliente)
    If lErro <> SUCESSO Then Error 19213

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 19211, 19213, 52691
        
        Case 71925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigoLoja)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154261)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Descarrega variáveis globais tipo AdmEvento
    Set objEventoCliente = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Label1_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    objCliente.lCodigoLoja = StrParaLong(Codigo.Text)
    
    objCliente.lCodigo = StrParaLong(CodigoBO.Caption)

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Public Sub Label3_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_Label3_Click

    objCliente.sNomeReduzido = NomeReduzido.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

    Exit Sub

Erro_Label3_Click:

    Select Case gErr

        Case 71926

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154262)

    End Select

    Exit Sub


End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Executa o Validate
    Call Cliente_Traz_Tela(objCliente.lCodigo, objCliente.lCodigoLoja)

    Me.Show

    Exit Sub

End Sub

Public Sub Cliente_Traz_Tela(ByVal lCodigo As Long, ByVal lCodigoLoja As Long)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objClienteEstatistica As New ClassFilialClienteEst

On Error GoTo Erro_Cliente_Traz_Tela

    'Guarda o valor do código do Cliente selecionado na ListBox ClientesList
    objCliente.lCodigoLoja = lCodigoLoja
    objCliente.iFilialEmpresaLoja = giFilialEmpresa
    objCliente.lCodigo = lCodigo

    'Lê o Cliente no BD
    lErro = CF("Cliente_Le_Estendida", objCliente, objClienteEstatistica)
    If lErro <> SUCESSO And lErro <> 52545 Then gError 71923

    'Se cliente não está cadastrado, erro
    If lErro = 12293 Then gError 71925

    'Exibe os dados do Cliente
    lErro = Exibe_Dados_Cliente(objCliente)
    If lErro <> SUCESSO Then gError 71924

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Cliente_Traz_Tela:

    Select Case gErr

        Case 71923, 71924

        Case 71925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigoLoja)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154263)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Clientes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ClienteLoja"
    
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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call Label1_Click
        ElseIf Me.ActiveControl Is NomeReduzido Then
            Call Label3_Click
        End If
    
    End If

End Sub

