VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpTitPagRateio 
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   ScaleHeight     =   4155
   ScaleWidth      =   7965
   Begin VB.Frame Frame1 
      Caption         =   "NDs"
      Height          =   825
      Left            =   285
      TabIndex        =   17
      Top             =   2895
      Width           =   5355
      Begin MSMask.MaskEdBox ND_I 
         Height          =   300
         Left            =   570
         TabIndex        =   18
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ND_F 
         Height          =   300
         Left            =   3240
         TabIndex        =   19
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   210
         TabIndex        =   21
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   2805
         TabIndex        =   20
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.OptionButton OptReimprimir 
      Caption         =   "Reimprimir NDs"
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
      Left            =   195
      TabIndex        =   16
      Top             =   2520
      Width           =   2415
   End
   Begin VB.OptionButton OptGerar 
      Caption         =   "Gerar NDs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   15
      Top             =   930
      Width           =   1545
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   825
      Left            =   285
      TabIndex        =   7
      Top             =   1575
      Width           =   5355
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   570
         TabIndex        =   8
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   9
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   2805
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   360
         Width           =   360
      End
      Begin VB.Label LabelClienteDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpTitPagRateioOcx.ctx":0000
      Left            =   1440
      List            =   "RelOpTitPagRateioOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   255
      Width           =   2730
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5775
      Picture         =   "RelOpTitPagRateioOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   825
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5670
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpTitPagRateioOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpTitPagRateioOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpTitPagRateioOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpTitPagRateioOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   2265
      TabIndex        =   13
      Top             =   1170
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Despesas acima de:"
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
      Left            =   465
      TabIndex        =   14
      Top             =   1230
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
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
      Left            =   735
      TabIndex        =   12
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpTitPagRateio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long
Private Declare Function Comando_PrepararPosInt Lib "ADSQLMN.DLL" Alias "AD_Comando_PrepararPos" (ByVal lComando As Long, ByVal lpSQLStmt As String, ByVal lSelect As Long) As Long

Private WithEvents objEventoClienteInic As AdmEvento
Attribute objEventoClienteInic.VB_VarHelpID = -1
Private WithEvents objEventoClienteFim As AdmEvento
Attribute objEventoClienteFim.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 129064
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 129065
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 129064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 129065
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 129066
    
    ComboOpcoes.Text = ""
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 129067
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 129066, 129067
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
   
End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate
    
    'Se está Preenchido
    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 129068

    End If
    
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 129068
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate
    
    'se está Preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 129069

    End If
        
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 129069
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoClienteInic = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
         
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 129070
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 129070
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteAte_Click
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteFim)

   Exit Sub

Erro_LabelClienteAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteDe_Click
    
    If Len(Trim(ClienteInicial.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteInic)

   Exit Sub

Erro_LabelClienteDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub ND_F_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ND_F_Validate

    If Len(Trim(ND_F.Text)) > 0 Then
        lErro = Valor_Positivo_Critica(ND_F.Text)
        If lErro <> SUCESSO Then gError 129086
    End If
    
    Exit Sub

Erro_ND_F_Validate:

    Cancel = True

    Select Case gErr

        Case 129086

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ND_I_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ND_I_Validate

    If Len(Trim(ND_I.Text)) > 0 Then
        lErro = Valor_Positivo_Critica(ND_I.Text)
        If lErro <> SUCESSO Then gError 129087
    End If
    
    Exit Sub

Erro_ND_I_Validate:

    Cancel = True

    Select Case gErr

        Case 129087

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Controle_Ativos(bFlag As Boolean)

    Valor.Enabled = bFlag
    ClienteInicial.Enabled = bFlag
    ClienteFinal.Enabled = bFlag
    ND_I.Enabled = Not bFlag
    ND_F.Enabled = Not bFlag

End Sub

Private Sub objEventoClienteFim_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche o Cliente Final com o Codigo selecionado
    ClienteFinal.Text = CStr(objCliente.lCodigo)
    'Preenche o Cliente Final com Codigo - Descricao
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteInic_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche o Cliente Inical com o codigo
    ClienteInicial.Text = CStr(objCliente.lCodigo)
    
    'Preenche o Cliente Inicial com codigo - Descricao
    Call ClienteInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 129071

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 129072

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 129073
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 129074
    
    Call BotaoLimpar_Click
               
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 129071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 129072 To 129074
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 129075

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 129076

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 129075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 129076

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 129077
    
    lErro = ND_Gera()
    If lErro <> SUCESSO Then gError 129114
        
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 129077, 129114

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sValor As String
Dim sND_I As String
Dim sND_F As String
Dim sOPT As String

On Error GoTo Erro_PreencherRelOp
            
    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sCliente_I, sCliente_F, sValor, sND_I, sND_F)
    If lErro <> SUCESSO Then gError 129078
        
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 129079
         
    'Preenche o Cliente Inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError 129080
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 129081
    
    'Preenche o Cliente Final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError 129082
     
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 129083
           
    'Preenche o Valor
    lErro = objRelOpcoes.IncluirParametro("NVLRRATEIO", sValor)
    If lErro <> AD_BOOL_TRUE Then gError 129084

    'Preenche ND inicial
    lErro = objRelOpcoes.IncluirParametro("NNDINIC", sND_I)
    If lErro <> AD_BOOL_TRUE Then gError 129088
    
    'Preenche ND Final
    lErro = objRelOpcoes.IncluirParametro("NNDFIM", sND_F)
    If lErro <> AD_BOOL_TRUE Then gError 129089
    
    If OptGerar.Value = True Then
        sOPT = "GERAR"
    Else
        sOPT = "REIMPRIMIR"
    End If
    
    lErro = objRelOpcoes.IncluirParametro("TOPT", sOPT)
    If lErro <> AD_BOOL_TRUE Then gError 129112

    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCliente_I, sCliente_F, sValor, sND_I, sND_F)
    If lErro <> SUCESSO Then gError 129085

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 129078 To 129085
        
        Case 129088, 129089, 129112
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function ND_Gera() As Long
'preenche objRelOpcoes com os dados da tela

Dim sCliente_I As String
Dim sCliente_F As String
Dim sValor As String
Dim sND_I As String
Dim sND_F As String
Dim lErro As Long

On Error GoTo Erro_ND_Gera

    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sCliente_I, sCliente_F, sValor, sND_I, sND_F)
    If lErro <> SUCESSO Then gError 129115
            
    If OptGerar.Value = True Then
        lErro = TitulosPagRateioND_Grava(sCliente_I, sCliente_F, sValor, sND_I, sND_F)
        If lErro <> SUCESSO Then gError 129111
    End If

    ND_Gera = SUCESSO

    Exit Function

Erro_ND_Gera:

    ND_Gera = gErr

    Select Case gErr
        
        Case 129111, 129115
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_I As String, sCliente_F As String, sValor As String, sND_I As String, sND_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
        
    If OptGerar.Value = True Then
    
        'critica Cliente Inicial e Final
        If ClienteInicial.Text <> "" Then
            sCliente_I = CStr(LCodigo_Extrai(ClienteInicial.Text))
        Else
            sCliente_I = ""
        End If
        
        If ClienteFinal.Text <> "" Then
            sCliente_F = CStr(LCodigo_Extrai(ClienteFinal.Text))
        Else
            sCliente_F = ""
        End If
                               
        If Valor.Text <> "" Then
            sValor = Format(Valor.Text, "Standard")
        Else
            sValor = ""
        End If

        If sCliente_I <> "" And sCliente_F <> "" Then
            If StrParaLong(sCliente_I) > StrParaLong(sCliente_F) Then gError 129057
        End If
    
    Else
    
        'critica ND Inicial e Final
        sND_I = ND_I.Text
    
        sND_F = ND_F.Text
    
        If sND_I <> "" And sND_F <> "" Then
            If StrParaLong(sND_I) > StrParaLong(sND_F) Then gError 129090
        End If
   
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 129057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
                
        Case 129058
        
        Case 129090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ND_INICIAL_MAIOR", gErr)
            ND_I.SetFocus
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCliente_I As String, sCliente_F As String, sValor As String, sND_I As String, sND_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sCliente_I <> "" Then sExpressao = "Cliente >= " & Forprint_ConvLong(StrParaLong(sCliente_I))

   If sCliente_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(StrParaLong(sCliente_F))

    End If
    
    If sValor <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Valor >= " & Forprint_ConvDouble(StrParaDbl(sValor))
    End If
    
   If sND_I <> "" Then sExpressao = "ND >= " & Forprint_ConvLong(StrParaLong(sND_I))

   If sND_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "ND <= " & Forprint_ConvLong(StrParaLong(sND_F))

    End If
                 
    If giFilialEmpresa <> gobjCR.iFilialCentralizadora Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(CInt(giFilialEmpresa))
    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 129059
   
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 129060
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 129061
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
                    
    'pega valor
    lErro = objRelOpcoes.ObterParametro("NVLRRATEIO", sParam)
    If lErro <> SUCESSO Then gError 129062
    
    Valor.Text = sParam
    Call Valor_Validate(bSGECancelDummy)
   
    'pega ND inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NNDINIC", sParam)
    If lErro <> SUCESSO Then gError 129091
    
    ND_I.Text = sParam
    Call ND_I_Validate(bSGECancelDummy)

    'pega ND  final e exibe
    lErro = objRelOpcoes.ObterParametro("NNDFIM", sParam)
    If lErro <> SUCESSO Then gError 129092
    
    ND_F.Text = sParam
    Call ND_F_Validate(bSGECancelDummy)
    
    'pega opção
    lErro = objRelOpcoes.ObterParametro("TOPT", sParam)
    If lErro <> SUCESSO Then gError 129113
    
    If sParam = "GERAR" Then
        OptGerar.Value = True
        Controle_Ativos (True)
    Else
        OptReimprimir.Value = True
        Controle_Ativos (False)
    End If

          
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 129059 To 129062
        
        Case 129091, 129092
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    OptGerar.Value = True
    
    Call Controle_Ativos(True)
       
    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    Exit Function
    
End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub OptGerar_Click()

    Controle_Ativos (True)

End Sub

Private Sub OptReimprimir_Click()

    Controle_Ativos (False)

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    If Len(Trim(Valor.Text)) <> 0 Then
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 129063
    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 129063

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub
Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing
    
End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITREC_L
    Set Form_Load_Ocx = Me
    Caption = "Títulos a Pagar Rateio"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpTitPagRateio"
    
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

Public Sub Unload(objme As Object)
    
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
        
        If Me.ActiveControl Is ClienteInicial Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteFinal Then
            Call LabelClienteAte_Click
        End If
    
    End If

End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

'#################################
'Inserido por Wagner
'Acesso ao BD

Public Function TitulosPagRateioND_Grava(sCliente_I As String, sCliente_F As String, sValor As String, sND_I As String, sND_F As String) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iSeqAux As Integer
Dim lTransacao As Long
Dim sSQL As String
Dim lNumProxND As Long

Dim vlClienteAux As Variant
Dim vlCliente As Variant
Dim vsCliente_I, vsCliente_F, vsValor, vsND_I, vsND_F As Variant

On Error GoTo Erro_TitulosPagRateioND_Grava

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 129100
       
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 129093
    Next
    
    vsCliente_I = StrParaLong(sCliente_I)
    vsCliente_F = StrParaLong(sCliente_F)
    vsValor = StrParaDbl(sValor)
    vsND_I = StrParaLong(sND_I)
    vsND_F = StrParaLong(sND_F)
    vlCliente = 0
    
    lErro = TitulosPagRateioND_Prepara(vlCliente, vsCliente_I, vsCliente_F, vsValor, vsND_I, vsND_F, sSQL, alComando(1))
    If lErro <> SUCESSO Then gError 129102
                  
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129096
    
    While lErro <> AD_SQL_SEM_DADOS
    
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 129097
    
        If vlClienteAux <> vlCliente Then
            lErro = CF("Config_ObterNumInt", "CPRConfig", "NUM_PROX_ND", lNumProxND)
            If lErro <> SUCESSO Then gError 129095
            vlClienteAux = vlCliente
        End If

        'Exclui Rateio
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE TitulosPagRateio Set ND = ?, Data_Ger_ND = ? ", alComando(1), lNumProxND, Date)
        If lErro <> AD_SQL_SUCESSO Then gError 129098

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129099

    Wend
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Finaliza transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 129101
        
    TitulosPagRateioND_Grava = SUCESSO
    
    Exit Function
    
Erro_TitulosPagRateioND_Grava:

    TitulosPagRateioND_Grava = gErr

    Select Case gErr
    
        Case 129093
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
              
        Case 129095
        
        Case 129096 To 129099
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAGRATEIO", gErr)
        
        Case 129100
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 129101
                Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function TitulosPagRateioND_Prepara(vlCliente As Variant, ByVal vsCliente_I As Variant, ByVal vsCliente_F As Variant, ByVal vsValor As Variant, ByVal vsND_I As Variant, ByVal vsND_F As Variant, sSQL As String, lComando As Long) As Long

Dim lErro As Long
Dim sSelect As String, sWhere As String, sFrom As String, sOrderBy As String
Dim viCobrar As Variant

On Error GoTo Erro_TitulosPagRateioND_Prepara

    sSelect = "SELECT  Cliente "
    
    sFrom = "FROM  TitulosPagRateio "
                     
    sWhere = "WHERE Cobrar = ? "
                         
    If vsCliente_I <> 0 Then
        sWhere = sWhere & "AND Cliente >= ? "
    End If
    
    If vsCliente_F <> 0 Then
        sWhere = sWhere & "AND Cliente <= ? "
    End If
    
    If vsValor <> 0 Then
        sWhere = sWhere & "AND Valor >= ? "
    End If

    If vsND_I <> 0 Then
            sWhere = sWhere & "AND ND >= ? "
    End If
    
    If vsND_F <> 0 Then
            sWhere = sWhere & "AND ND <= ? "
    End If
    
    sOrderBy = " ORDER BY Cliente"
    
    sSQL = sSelect & sFrom & sWhere & sOrderBy
    
    viCobrar = StrParaInt("1")
    
    lErro = Comando_PrepararPosInt(lComando, sSQL, 0)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129103
       
    lErro = Comando_BindVarInt(lComando, vlCliente)
    If (lErro <> AD_SQL_SUCESSO) Then Error 129104
    
    lErro = Comando_BindVarInt(lComando, viCobrar)
    If (lErro <> AD_SQL_SUCESSO) Then Error 129136

    If vsCliente_I <> 0 Then
        lErro = Comando_BindVarInt(lComando, vsCliente_I)
        If (lErro <> AD_SQL_SUCESSO) Then Error 129105
    End If
    
    If vsCliente_F <> 0 Then
        lErro = Comando_BindVarInt(lComando, vsCliente_F)
        If (lErro <> AD_SQL_SUCESSO) Then Error 129106
    End If
    
    If vsValor <> 0 Then
        lErro = Comando_BindVarInt(lComando, vsValor)
        If (lErro <> AD_SQL_SUCESSO) Then Error 129107
    End If

    If vsND_I <> 0 Then
        lErro = Comando_BindVarInt(lComando, vsND_I)
        If (lErro <> AD_SQL_SUCESSO) Then Error 129108
    End If
    
    If vsND_F <> 0 Then
        lErro = Comando_BindVarInt(lComando, vsND_F)
        If (lErro <> AD_SQL_SUCESSO) Then Error 129109
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then Error 129110
    
    TitulosPagRateioND_Prepara = SUCESSO

    Exit Function

Erro_TitulosPagRateioND_Prepara:

    TitulosPagRateioND_Prepara = gErr

    Select Case gErr
            
        Case 129103 To 129110
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAGRATEIO", gErr)

        Case 129136
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAGRATEIO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function
