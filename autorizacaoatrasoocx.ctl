VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl AutorizacaoAtrasoOcx 
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   ScaleHeight     =   5895
   ScaleWidth      =   7590
   Begin VB.Frame Frame5 
      Caption         =   "Total Utilizado"
      Height          =   765
      Left            =   165
      TabIndex        =   31
      Top             =   4200
      Width           =   2940
      Begin VB.Label Label12 
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
         Left            =   255
         TabIndex        =   33
         Top             =   375
         Width           =   510
      End
      Begin VB.Label TotalGeral 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   885
         TabIndex        =   32
         Top             =   330
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Limites em"
      Height          =   675
      Left            =   150
      TabIndex        =   26
      Top             =   3465
      Width           =   7200
      Begin VB.Label NFNaoFat 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5235
         TabIndex        =   30
         Top             =   225
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "NFs não faturadas:"
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
         Left            =   3600
         TabIndex        =   29
         Top             =   285
         Width           =   1635
      End
      Begin VB.Label PedVendas 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1545
         TabIndex        =   28
         Top             =   210
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pedidos Venda:"
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
         Left            =   165
         TabIndex        =   27
         Top             =   270
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Responsável pela autorização"
      Height          =   1605
      Left            =   3285
      TabIndex        =   21
      Top             =   4215
      Width           =   3060
      Begin VB.ComboBox ComboUsuario 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   465
         Width           =   2010
      End
      Begin VB.TextBox TextSenha 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   810
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   945
         Width           =   2010
      End
      Begin VB.Label UsuariosLabel 
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
         Left            =   195
         TabIndex        =   23
         Top             =   510
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1005
         Width           =   615
      End
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
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
      Left            =   6480
      Picture         =   "autorizacaoatrasoocx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4365
      Width           =   975
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
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
      Left            =   6480
      Picture         =   "autorizacaoatrasoocx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5130
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Títulos em Aberto"
      Height          =   2955
      Left            =   165
      TabIndex        =   7
      Top             =   480
      Width           =   7185
      Begin MSMask.MaskEdBox Status 
         Height          =   225
         Left            =   5400
         TabIndex        =   25
         Top             =   870
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Saldo 
         Height          =   225
         Left            =   4290
         TabIndex        =   24
         Top             =   660
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataVencto 
         Height          =   225
         Left            =   4980
         TabIndex        =   9
         Top             =   1200
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   225
         Left            =   3270
         TabIndex        =   10
         Top             =   1110
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Parcela 
         Height          =   225
         Left            =   2220
         TabIndex        =   11
         Top             =   1050
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NumTitulo 
         Height          =   225
         Left            =   1200
         TabIndex        =   12
         Top             =   1080
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridParcelasAbertas 
         Height          =   1590
         Left            =   270
         TabIndex        =   8
         Top             =   330
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   2805
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColor       =   16777215
         ForeColorFixed  =   0
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Total em Aberto:"
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
         Left            =   3660
         TabIndex        =   18
         Top             =   2460
         Width           =   1965
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Total em Atraso:"
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
         Left            =   135
         TabIndex        =   17
         Top             =   2460
         Width           =   1950
      End
      Begin VB.Label TotalAberto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5625
         TabIndex        =   16
         Top             =   2415
         Width           =   1410
      End
      Begin VB.Label TotalAtraso 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2085
         TabIndex        =   15
         Top             =   2415
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Crédito a ser concedido"
      Height          =   765
      Left            =   165
      TabIndex        =   0
      Top             =   5040
      Width           =   2940
      Begin VB.Label LabelValor 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   885
         TabIndex        =   6
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         Left            =   255
         TabIndex        =   5
         Top             =   375
         Width           =   510
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Limite de Crédito:"
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
      Left            =   3900
      TabIndex        =   20
      Top             =   150
      Width           =   1500
   End
   Begin VB.Label LimiteCredito 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5430
      TabIndex        =   19
      Top             =   90
      Width           =   1770
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   14
      Top             =   180
      Width           =   660
   End
   Begin VB.Label LabelCliente 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   960
      TabIndex        =   13
      Top             =   135
      Width           =   2580
   End
End
Attribute VB_Name = "AutorizacaoAtrasoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Trecho inserido por Leo em 28/02/02

'Variáveis relacionadas ao Grid de Parcelas em aberto
Dim objGridParcelasAbertas As AdmGrid
Dim iGrid_NumTitulo_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Saldo_Col As Integer
Dim iGrid_DataVencto_Col As Integer
Dim iGrid_Status_Col As Integer

'Leo até aqui

Dim objAutorizacaoCredito As ClassAutorizacaoCredito

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objUsuarios As New ClassUsuarios
Dim objLiberacaoCredito As New ClassLiberacaoCredito
Dim dValorCreditoSolicitado As Double
Dim objValorLiberadoCredito As New ClassValorLiberadoCredito

On Error GoTo Erro_BotaoOK_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Usuario foi Preenchido
    If Len(ComboUsuario.Text) = 0 Then Error 44434
    
    'Verifica se digitou a senha
    If Len(TextSenha.Text) = 0 Then Error 44435
    
    objUsuarios.sCodUsuario = ComboUsuario.Text

    'le os dados do usuário
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 44433 Then Error 44436
    
    'se o usuário não está cadastrado ==> erro.
    If lErro = 44433 Then Error 44437
    
    'se a senha não for a que está cadastada ==> erro
    If TextSenha.Text <> objUsuarios.sSenha Then Error 44438
    
    If giTipoVersao = VERSAO_FULL Then
    
        objLiberacaoCredito.sCodUsuario = objUsuarios.sCodUsuario
        
        'Lê a liberacao de credito a partir do código do usuario.
        lErro = CF("LiberacaoCredito_Le", objLiberacaoCredito)
        If lErro <> SUCESSO And lErro <> 36968 Then Error 44440
        
        'se não foi encontrado autorização para o usuario liberar credito
        If lErro = 36968 Then Error 44441
            
        dValorCreditoSolicitado = CDbl(LabelValor.Caption)
            
        'se o valor do crédito solicitado ultrapassar o limite de credito que o usuario pode conceder por operacao
        If dValorCreditoSolicitado > objLiberacaoCredito.dLimiteOperacao Then Error 44442
            
        objValorLiberadoCredito.sCodUsuario = objUsuarios.sCodUsuario
        objValorLiberadoCredito.iAno = Year(gdtDataAtual)
            
        'Lê a estatistica de liberação de credito de um usuario em um determinado ano
        lErro = CF("ValorLiberadoCredito_Le", objValorLiberadoCredito)
        If lErro <> SUCESSO And lErro <> 36973 Then Error 44443
            
        'se o valor do pedido ultrapassar o valor mensal que o usuario tem capacidade de liberar
        If dValorCreditoSolicitado > objLiberacaoCredito.dLimiteMensal - objValorLiberadoCredito.adValorLiberado(Month(gdtDataAtual)) Then Error 44445
    
    End If
    
    objAutorizacaoCredito.iCreditoAutorizado = CREDITO_APROVADO
    objAutorizacaoCredito.sCodUsuario = ComboUsuario.Text
        
    GL_objMDIForm.MousePointer = vbDefault
     
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 44434
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", Err)
    
        Case 44435
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_PREENCHIDA", Err)
    
        Case 44436, 44440, 44443
    
        Case 44437
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", Err, objUsuarios.sCodUsuario)
    
        Case 44438
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_INVALIDA", Err)
    
        Case 44441
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_INEXISTENTE", Err, objLiberacaoCredito.sCodUsuario)
    
        Case 44442
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_LIMITEOPERACAO", Err, objLiberacaoCredito.sCodUsuario)
    
        Case 44445
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_LIMITEMENSAL", Err, objLiberacaoCredito.sCodUsuario)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 143177)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoCancela_Click()

    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objGridParcelasAbertas = New AdmGrid 'Por leo em 28/02/02
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 143178)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objAutorizacaoCredito = Nothing
    Set objGridParcelasAbertas = Nothing 'Por leo em 28/02/02

End Sub

Function Trata_Parametros(Optional objAutorizacaoCredito1 As ClassAutorizacaoCredito) As Long

Dim lErro As Long
Dim iIndice As Integer, objCliente As New ClassCliente
Dim colUsuarios As New Collection
Dim objUsuarios As ClassUsuarios
Dim colUsuariosComLiberacao As New Collection
Dim objUsuariosComLiberacao As ClassUsuarios
Dim colParcRec As New Collection

'****Incluido poe Shirley em 04/06/2002***********
Dim objClienteEstatistica As New ClassFilialClienteEst
'*************************************************

On Error GoTo Erro_Trata_Parametros

    If (objAutorizacaoCredito1 Is Nothing) Then Error 44429
    If (objAutorizacaoCredito1 Is Nothing) Then Set objAutorizacaoCredito1 = New ClassAutorizacaoCredito
    'Preenche Cliente e VAlor
    objCliente.lCodigo = objAutorizacaoCredito1.lCliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 59357
    If lErro <> SUCESSO Then gError 59358
    
    LabelCliente.Caption = objAutorizacaoCredito1.lCliente & SEPARADOR & objCliente.sNomeReduzido
    LabelValor.Caption = Format(objAutorizacaoCredito1.dValor, "Standard")
    LimiteCredito.Caption = Format(objCliente.dLimiteCredito, "STANDARD") 'Por Leo em 01/03/02
    
    'Preenche a Combo com Usuarios que tem alçada Superior ao Valor da Operacao
    
    'Le todos os usuarios para esta Filial Empresa
    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 44446
    
    If giTipoVersao = VERSAO_FULL Then
        'Le todos os Usuarios que tem Liberacao de Crédito por Operacao e Mensal
        lErro = CF("Usuarios_Com_LiberacaoCredito_Le", colUsuarios, objAutorizacaoCredito1.dValor, colUsuariosComLiberacao)
        If lErro <> SUCESSO Then gError 58581
    
        'Preenche a combo de Usuarios
        For Each objUsuariosComLiberacao In colUsuariosComLiberacao
            ComboUsuario.AddItem objUsuariosComLiberacao.sCodUsuario
        Next
        
    ElseIf giTipoVersao = VERSAO_LIGHT Then
        
        'Preenche a combo de Usuarios
        For Each objUsuarios In colUsuarios
            ComboUsuario.AddItem objUsuarios.sCodUsuario
        Next
    
    End If
    
    'Lê as parcelas em aberto de um cliente
    lErro = CF("ParcelasRec_Cli_Abertas_Le", colParcRec, objCliente.lCodigo)
    If lErro <> SUCESSO And lErro <> 94399 Then gError 94396
    
    lErro = Inicializa_GridParcelasAbertas(objGridParcelasAbertas, colParcRec.Count)
    If lErro <> SUCESSO Then gError 94391

    If lErro = SUCESSO Then
        
        'Preenche o GridParcelasAbertas
        lErro = Preenche_GridParcelasAbertas(colParcRec)
        If lErro <> SUCESSO Then gError 94397
    
    End If
    
    '****Incluido poe Shirley em 04/06/2002***********
    objClienteEstatistica.lCodCliente = objAutorizacaoCredito1.lCliente
    lErro = CF("Cliente_Le_Estatistica_Credito", objClienteEstatistica)
    If lErro <> SUCESSO Then gError 52954
    '*************************************************
    
    PedVendas.Caption = Format(objClienteEstatistica.dSaldoPedidosLiberados, "Standard")
    NFNaoFat.Caption = Format(objClienteEstatistica.dValorNFsNaoFaturadas, "Standard")
    
    TotalGeral.Caption = Format(objClienteEstatistica.dSaldoPedidosLiberados + objClienteEstatistica.dValorNFsNaoFaturadas + StrParaDbl(TotalAberto.Caption) + StrParaDbl(TotalAtraso.Caption), "Standard")
    
    'Seta o objeto global a Tela (objAutorizacaoCredito)
    Set objAutorizacaoCredito = objAutorizacaoCredito1
    
    objAutorizacaoCredito.iCreditoAutorizado = CREDITO_RECUSADO
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 94391
        
        Case 44446, 59357, 58581 'Tratados nas Rotinas Chamadas
    
        Case 59358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
        
        Case 44429
            lErro = Rotina_Erro(vbOKOnly, "TELA_AUTCRED_CHAMADA_SEM_PARAMETRO", gErr)
    
        Case 94396, 94397
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143179)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_AUTORIZACAO_CREDITO
    Set Form_Load_Ocx = Me
    Caption = "Autorização de Crédito"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AutorizacaoCredito"
    
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



Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValor, Source, X, Y)
End Sub

Private Sub LabelValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValor, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub UsuariosLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UsuariosLabel, Source, X, Y)
End Sub

Private Sub UsuariosLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UsuariosLabel, Button, Shift, X, Y)
End Sub

'Leo daqui para baixo - 28/02/02

Private Function Inicializa_GridParcelasAbertas(objGridInt As AdmGrid, ByVal iLinhasExistentes As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridParcelasAbertas
    
    'Tela em questão
    Set objGridInt.objForm = Me
    
    'Títulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Titulo")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Saldo")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Status")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (NumTitulo.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (Saldo.Name)
    objGridInt.colCampo.Add (DataVencto.Name)
    objGridInt.colCampo.Add (Status.Name)
    
    iGrid_NumTitulo_Col = 1
    iGrid_Parcela_Col = 2
    iGrid_Valor_Col = 3
    iGrid_Saldo_Col = 4
    iGrid_DataVencto_Col = 5
    iGrid_Status_Col = 6
        
    objGridInt.objGrid = GridParcelasAbertas
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = iLinhasExistentes + 2
                
    If iLinhasExistentes < 5 Then
                
        'linhas visiveis do grid
        objGridInt.iLinhasVisiveis = iLinhasExistentes
                
    Else
                    
        'linhas visiveis do grid
        objGridInt.iLinhasVisiveis = 5
                       
    End If
                       
    GridParcelasAbertas.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR

    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridParcelasAbertas = SUCESSO
    
    Exit Function
    
Erro_Inicializa_GridParcelasAbertas:

    Inicializa_GridParcelasAbertas = gErr
    
    Select Case gErr
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143180)
        
    End Select

    Exit Function
        
End Function


    
'Por Leo em 01/03/02
Private Function Preenche_GridParcelasAbertas(ByVal colParcRec As Collection) As Long
    
Dim iIndice As Integer
Dim objParcRec As ClassInfoParcRec
Dim lErro As Long
'Dim dtDataLimite As Date
Dim dTotalAtraso As Double
Dim dTotalAberto As Double

On Error GoTo Erro_Preenche_GridParcelasAbertas

    'Limpa o grid de Parcelas Abertas
    Call Grid_Limpa(objGridParcelasAbertas)
    
'    'Verifica a data limite p/ atraso de uma parcela. Retorna DATA_NULA caso a empresa não use o Bloqueio por atraso.
'    lErro = CF("Verifica_DataLimite_Bloqueio_Atraso", dtDataLimite)
'    If lErro <> SUCESSO Then gError 94398
    
    For Each objParcRec In colParcRec
                                        
        iIndice = iIndice + 1
                
        'Coloca os dados das Parcelas na tela
        GridParcelasAbertas.TextMatrix(iIndice, iGrid_NumTitulo_Col) = objParcRec.lNumTitulo
        GridParcelasAbertas.TextMatrix(iIndice, iGrid_Parcela_Col) = objParcRec.iNumParcela
        GridParcelasAbertas.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objParcRec.dValor, "STANDARD")
        GridParcelasAbertas.TextMatrix(iIndice, iGrid_Saldo_Col) = Format(objParcRec.dSaldoParcela, "STANDARD")
        GridParcelasAbertas.TextMatrix(iIndice, iGrid_DataVencto_Col) = Format(objParcRec.dtDataVencimentoReal, "dd/mm/yy")
    
'        'Se a parcela tiver data de vencimento real menor ou igual que a data limite p/ atraso.
'        If dtDataLimite <> DATA_NULA And objParcRec.dtDataVencimentoReal <= dtDataLimite Then
    
        If objParcRec.dtDataVencimentoReal < gdtDataHoje Then
    
            GridParcelasAbertas.TextMatrix(iIndice, iGrid_Status_Col) = "ATRASO"
            
            dTotalAtraso = dTotalAtraso + objParcRec.dSaldoParcela
        
        End If
        
        dTotalAberto = dTotalAberto + objParcRec.dSaldoParcela
        
    Next
    
    TotalAtraso.Caption = Format(dTotalAtraso, "STANDARD")
    
    TotalAberto.Caption = Format(dTotalAberto, "STANDARD")
    
    'Inicializa o número de linhas existentes no grid
    objGridParcelasAbertas.iLinhasExistentes = iIndice
    
    Preenche_GridParcelasAbertas = SUCESSO
    
    Exit Function
    
Erro_Preenche_GridParcelasAbertas:
    
    Preenche_GridParcelasAbertas = gErr
    
    Select Case gErr
    
        Case 94398
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143181)
    
    End Select

    Exit Function
    
End Function

