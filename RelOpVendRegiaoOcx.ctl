VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpVendRegiaoOcx 
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   LockControls    =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   6270
   Begin VB.Frame Frame1 
      Caption         =   "Região"
      Height          =   1245
      Left            =   120
      TabIndex        =   15
      Top             =   2250
      Width           =   5955
      Begin MSMask.MaskEdBox RegiaoInicial 
         Height          =   315
         Left            =   945
         TabIndex        =   3
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RegiaoFinal 
         Height          =   315
         Left            =   945
         TabIndex        =   4
         Top             =   765
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelRegiaoAte 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   510
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   810
         Width           =   435
      End
      Begin VB.Label LabelRegiaoDe 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   555
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   360
         Width           =   360
      End
      Begin VB.Label RegiaoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2385
         TabIndex        =   17
         Top             =   315
         Width           =   3120
      End
      Begin VB.Label RegiaoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2385
         TabIndex        =   16
         Top             =   765
         Width           =   3120
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpVendRegiaoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpVendRegiaoOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpVendRegiaoOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpVendRegiaoOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame AlmoxarifadoInicial 
      Caption         =   "Vendedores"
      Height          =   1320
      Left            =   120
      TabIndex        =   10
      Top             =   870
      Width           =   3540
      Begin MSMask.MaskEdBox VendedorInicial 
         Height          =   300
         Left            =   975
         TabIndex        =   1
         Top             =   315
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendedorFinal 
         Height          =   300
         Left            =   975
         TabIndex        =   2
         Top             =   780
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedorFinal 
         Alignment       =   1  'Right Justify
         Caption         =   "Final:"
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
         Height          =   255
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   855
         Width           =   525
      End
      Begin VB.Label LabelVendedorInicial 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicial:"
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
         Height          =   255
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   375
         Width           =   570
      End
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
      Left            =   4005
      Picture         =   "RelOpVendRegiaoOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   870
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpVendRegiaoOcx.ctx":0A96
      Left            =   825
      List            =   "RelOpVendRegiaoOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2730
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpVendRegiaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Dim giVendedorInicial As Integer
Dim giRegiaoVenda As Integer

'Browses
Private WithEvents objEventoRegiaoVenda As AdmEvento
Attribute objEventoRegiaoVenda.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoVendedor = New AdmEvento
    Set objEventoRegiaoVenda = New AdmEvento
        
    giVendedorInicial = 1
    giRegiaoVenda = 1
           
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173702)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 138980
   
    lErro = objRelOpcoes.ObterParametro("NREGIAOVENDAINIC", sParam)
    If lErro <> SUCESSO Then gError 138981

    RegiaoInicial.Text = sParam
    Call RegiaoInicial_Validate(bSGECancelDummy)
    
    'pega Região de Venda Final e exibe
    lErro = objRelOpcoes.ObterParametro("NREGIAOVENDAFIM", sParam)
    If lErro <> SUCESSO Then gError 138982

    RegiaoFinal.Text = sParam
    Call RegiaoFinal_Validate(bSGECancelDummy)
    
    'pega vendedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDINIC", sParam)
    If lErro <> SUCESSO Then gError 138983
    
    VendedorInicial.Text = sParam
    Call VendedorInicial_Validate(bSGECancelDummy)
    
    'pega  vendedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDFIM", sParam)
    If lErro <> SUCESSO Then gError 138984
    
    VendedorFinal.Text = sParam
    Call VendedorFinal_Validate(bSGECancelDummy)
          
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 138980 To 138984

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173703)

    End Select

    Exit Function

End Function

Private Function PreencheComboOpcoes(sCodRel As String) As Long

Dim colRelParametros As New Collection
Dim lErro As Long
Dim objRelOpcoes As AdmRelOpcoes

On Error GoTo Erro_PreencheComboOpcoes

    'le os nomes das opcoes do relatório existentes no BD
    lErro = CF("RelOpcoes_Le_Todos", sCodRel, colRelParametros)
    If lErro <> SUCESSO Then gError 138985

    'preenche o ComboBox com os nomes das opções do relatório
    For Each objRelOpcoes In colRelParametros
        ComboOpcoes.AddItem objRelOpcoes.sNome
    Next

    PreencheComboOpcoes = SUCESSO

    Exit Function

Erro_PreencheComboOpcoes:

    PreencheComboOpcoes = gErr

    Select Case gErr

        Case 138985

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173704)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long
Dim iOpcao As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 138986
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = PreencheComboOpcoes(gobjRelatorio.sCodRel)
    If lErro <> SUCESSO Then gError 138987

    'verifica se o nome da opção passada está no ComboBox
    For iOpcao = 0 To ComboOpcoes.ListCount - 1

        If ComboOpcoes.List(iOpcao) = gobjRelOpcoes.sNome Then

            ComboOpcoes.Text = ComboOpcoes.List(iOpcao)

            lErro = PreencherParametrosNaTela(gobjRelOpcoes)
            If lErro <> SUCESSO Then gError 138988

            Exit For

        End If

    Next
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 138986
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 138987, 138988
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173705)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
  
    Set objEventoVendedor = Nothing
    Set objEventoRegiaoVenda = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 138988
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    RegiaoAte.Caption = ""
    RegiaoDe.Caption = ""
    
    giRegiaoVenda = 1
    giVendedorInicial = 1
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 138988
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173706)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

Dim lErro As Long

On Error GoTo Erro_ComboOpcoes_Click
    
    If ComboOpcoes.ListIndex = -1 Then Exit Sub

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Le", gobjRelOpcoes)
    If (lErro <> SUCESSO) Then gError 138989

    lErro = PreencherParametrosNaTela(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 138990

    Exit Sub

Erro_ComboOpcoes_Click:

    Select Case gErr

        Case 138989, 138990

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173707)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sVend_I As String
Dim sVend_F As String

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros(sVend_I, sVend_F)
    If lErro <> SUCESSO Then gError 138991
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 138992
         
    lErro = objRelOpcoes.IncluirParametro("NVENDINIC", sVend_I)
    If lErro <> AD_BOOL_TRUE Then gError 138993

    lErro = objRelOpcoes.IncluirParametro("NVENDFIM", sVend_F)
    If lErro <> AD_BOOL_TRUE Then gError 138994
    
    lErro = objRelOpcoes.IncluirParametro("TVENDINIC", VendedorInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 140153

    lErro = objRelOpcoes.IncluirParametro("TVENDFIM", VendedorFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 140154
       
    lErro = objRelOpcoes.IncluirParametro("NREGIAOVENDAINIC", RegiaoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 138995

    lErro = objRelOpcoes.IncluirParametro("NREGIAOVENDAFIM", RegiaoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 138996
    
    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAINIC", RegiaoDe.Caption)
    If lErro <> AD_BOOL_TRUE Then gError 140155

    lErro = objRelOpcoes.IncluirParametro("TREGIAOVENDAFIM", RegiaoAte.Caption)
    If lErro <> AD_BOOL_TRUE Then gError 140156
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sVend_I, sVend_F)
    If lErro <> SUCESSO Then gError 138997
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 138991 To 138997, 140153 To 140156

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173708)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 138998

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELVENDREG")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 138999

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        Call BotaoLimpar_Click
    
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 138998
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 138999, 140000

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173709)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 140001

    Me.Enabled = False
    Call gobjRelatorio.Executar_Prossegue

    Unload Me

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 140001

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173710)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 140002

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 140003

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 140004

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 140005
    
    Call BotaoLimpar_Click
    
    ComboOpcoes.Text = ""

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 140002
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 140003 To 140006

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173711)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sVend_I As String, sVend_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iRegVendaInc As Integer
Dim iRegVendaFin As Integer

On Error GoTo Erro_Monta_Expressao_Selecao

    iRegVendaInc = Codigo_Extrai(RegiaoInicial.Text)
    iRegVendaFin = Codigo_Extrai(RegiaoFinal.Text)
    
   If sVend_I <> "" Then sExpressao = "Vendedor >= " & Forprint_ConvInt(StrParaInt(sVend_I))

   If sVend_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vendedor <= " & Forprint_ConvInt(StrParaInt(sVend_F))

    End If
     
    If RegiaoInicial.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "RegiaoVenda >= " & Forprint_ConvInt(iRegVendaInc)
    
    End If
    
    If RegiaoFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "RegiaoVenda <= " & Forprint_ConvInt(iRegVendaFin)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173712)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sVend_I As String, sVend_F As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'Se RegiãoInicial e RegiãoFinal estão preenchidos
    If Len(Trim(RegiaoInicial.Text)) > 0 And Len(Trim(RegiaoFinal.Text)) > 0 Then
    
        'Se Região inicial for maior que Região final, erro
        If CLng(RegiaoInicial.Text) > CLng(RegiaoFinal.Text) Then gError 140007
        
    End If
    
    'critica vendedor Inicial e Final
    If VendedorInicial.Text <> "" Then
        sVend_I = CStr(Codigo_Extrai(VendedorInicial.Text))
    Else
        sVend_I = ""
    End If
    
    If VendedorFinal.Text <> "" Then
        sVend_F = CStr(Codigo_Extrai(VendedorFinal.Text))
    Else
        sVend_F = ""
    End If
            
    If sVend_I <> "" And sVend_F <> "" Then
        
        If CInt(sVend_I) > CInt(sVend_F) Then gError 140008
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 140007
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAOVENDA_INICIAL_MAIOR", gErr)
            RegiaoInicial.SetFocus
       
        Case 140008
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", gErr)
            VendedorInicial.SetFocus
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173713)

    End Select

    Exit Function

End Function

Private Sub LabelVendedorFinal_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 0
    
    If Len(Trim(VendedorFinal.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorFinal.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub LabelVendedorInicial_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 1
    
    If Len(Trim(VendedorInicial.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorInicial.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    If giVendedorInicial = 1 Then
        VendedorInicial.Text = CStr(objVendedor.iCodigo)
        VendedorInicial_Validate (bSGECancelDummy)
    Else
        VendedorFinal.Text = CStr(objVendedor.iCodigo)
        VendedorFinal_Validate (bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is RegiaoInicial Then
            Call LabelRegiaoDe_Click
        ElseIf Me.ActiveControl Is RegiaoFinal Then
            Call LabelRegiaoAte_Click
        ElseIf Me.ActiveControl Is VendedorInicial Then
            Call LabelVendedorInicial_Click
        ElseIf Me.ActiveControl Is VendedorFinal Then
            Call LabelVendedorFinal_Click
        End If
    
    End If
    
End Sub

Private Sub VendedorInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorInicial_Validate

    If Len(Trim(VendedorInicial.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorInicial, objVendedor, 0)
        If lErro <> SUCESSO Then gError 140009

    End If
    
    giVendedorInicial = 1
    
    Exit Sub

Erro_VendedorInicial_Validate:

    Cancel = True
    
    Select Case gErr

        Case 140009
             Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173714)

    End Select

End Sub

Private Sub VendedorFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorFinal_Validate

    If Len(Trim(VendedorFinal.Text)) > 0 Then

        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorFinal, objVendedor, 0)
        If lErro <> SUCESSO Then gError 140010

    End If
    
    giVendedorInicial = 0
 
    Exit Sub

Erro_VendedorFinal_Validate:

    Cancel = True
    
    Select Case gErr

        Case 140010
         Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173715)

    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Vendedor x Região x Cliente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpVendRegiao"
    
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

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
   Height = UserControl.Height
End Property

Public Property Get Width() As Long
   Width = UserControl.Width
End Property

Private Sub LabelVendedorFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorFinal, Source, X, Y)
End Sub

Private Sub LabelVendedorFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorInicial, Source, X, Y)
End Sub

Private Sub LabelVendedorInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorInicial, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelRegiaoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRegiaoDe, Source, X, Y)
End Sub

Private Sub LabelRegiaoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRegiaoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelRegiaoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRegiaoAte, Source, X, Y)
End Sub

Private Sub LabelRegiaoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRegiaoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelRegiaoAte_Click()
    
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    giRegiaoVenda = 0
    
    'Se o tipo está preenchido
    If Len(Trim(RegiaoFinal.Text)) > 0 Then
        
        'Preenche com o tipo da tela
        objRegiaoVenda.iCodigo = StrParaInt(RegiaoFinal.Text)
    
    End If
    
    'Chama Tela RegiãoVendaLista
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiaoVenda, objEventoRegiaoVenda)
    
End Sub

Private Sub LabelRegiaoDe_Click()

Dim objRegiaoVenda As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    giRegiaoVenda = 1
    
    'Se o tipo está preenchido
    If Len(Trim(RegiaoInicial.Text)) > 0 Then
        
        'Preenche com o tipo da tela
        objRegiaoVenda.iCodigo = StrParaInt(RegiaoInicial.Text)
        
    End If
    
    'Chama Tela RegiãoVendaLista
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiaoVenda, objEventoRegiaoVenda)

End Sub

Private Sub objEventoRegiaoVenda_evSelecao(obj1 As Object)

Dim objRegiaoVenda As New ClassRegiaoVenda

    Set objRegiaoVenda = obj1
    
    'Preenche campo Tipo de produto
    If giRegiaoVenda = 1 Then
        RegiaoInicial.PromptInclude = False
        RegiaoInicial.Text = objRegiaoVenda.iCodigo
        RegiaoInicial.PromptInclude = True
        RegiaoDe.Caption = objRegiaoVenda.sDescricao
    Else
        RegiaoFinal.PromptInclude = False
        RegiaoFinal.Text = objRegiaoVenda.iCodigo
        RegiaoFinal.PromptInclude = True
        RegiaoAte.Caption = objRegiaoVenda.sDescricao
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub RegiaoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_RegiaoFinal_Validate

    giRegiaoVenda = 0
                                
    lErro = RegiaoVenda_Perde_Foco(RegiaoFinal, RegiaoAte)
    If lErro <> SUCESSO And lErro <> 140016 Then gError 140011
       
    If lErro = 140016 Then gError 140012
        
    Exit Sub

Erro_RegiaoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 140011
        
        Case 140012
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, RegiaoFinal.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173716)

    End Select

    Exit Sub

End Sub

Private Sub RegiaoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_RegiaoInicial_Validate

    giRegiaoVenda = 1
                
    lErro = RegiaoVenda_Perde_Foco(RegiaoInicial, RegiaoDe)
    If lErro <> SUCESSO And lErro <> 140016 Then gError 140013
       
    If lErro = 140016 Then gError 140014
    
    Exit Sub

Erro_RegiaoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 140013
        
        Case 140014
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, RegiaoInicial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173717)

    End Select

    Exit Sub

End Sub

Public Function RegiaoVenda_Perde_Foco(Regiao As Object, Desc As Object) As Long
'recebe MaskEdBox da Região de Venda e o label da descrição

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_RegiaoVenda_Perde_Foco
        
    If Len(Trim(Regiao.Text)) > 0 Then
        
        lErro = Inteiro_Critica(Regiao.Text)
        If lErro <> SUCESSO Then gError 140102
        
        objRegiaoVenda.iCodigo = StrParaInt(Regiao.Text)
    
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 140015
    
        If lErro = 16137 Then gError 140016

        Desc.Caption = objRegiaoVenda.sDescricao

    Else

        Desc.Caption = ""

    End If

    RegiaoVenda_Perde_Foco = SUCESSO

    Exit Function

Erro_RegiaoVenda_Perde_Foco:

    RegiaoVenda_Perde_Foco = gErr

    Select Case gErr

        Case 140015
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGIOESVENDAS", gErr, objRegiaoVenda.iCodigo)

        Case 140016, 140102
            'Erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173718)

    End Select

    Exit Function

End Function

