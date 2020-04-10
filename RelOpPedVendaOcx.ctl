VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPedVendaOcx 
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7155
   ScaleHeight     =   2760
   ScaleWidth      =   7155
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
      Left            =   5145
      Picture         =   "RelOpPedVendaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   900
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4665
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPedVendaOcx.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPedVendaOcx.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPedVendaOcx.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPedVendaOcx.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPedVendaOcx.ctx":0A96
      Left            =   1035
      List            =   "RelOpPedVendaOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   2670
   End
   Begin VB.Frame FramePedido 
      Caption         =   "Pedido"
      Height          =   675
      Left            =   285
      TabIndex        =   15
      Top             =   840
      Width           =   4305
      Begin MSMask.MaskEdBox PedidoInicial 
         Height          =   300
         Left            =   765
         TabIndex        =   1
         Top             =   270
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PedidoFinal 
         Height          =   300
         Left            =   2865
         TabIndex        =   2
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelPedFinal 
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
         Left            =   2415
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   17
         Top             =   285
         Width           =   360
      End
      Begin VB.Label LabelPedInicial 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   390
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   810
      Left            =   285
      TabIndex        =   10
      Top             =   1800
      Width           =   4275
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1620
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   315
         Left            =   660
         TabIndex        =   3
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   3780
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   315
         Left            =   2835
         TabIndex        =   4
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label dIni 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   270
         TabIndex        =   14
         Top             =   330
         Width           =   315
      End
      Begin VB.Label dFim 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2400
         TabIndex        =   13
         Top             =   315
         Width           =   360
      End
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
      Left            =   345
      TabIndex        =   18
      Top             =   255
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPedVendaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoOp As AdmEvento
Attribute objEventoOp.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Dim giPedidoInicial As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoOp = New AdmEvento
     
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171049)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoOp = Nothing
    
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 90769
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 90770

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 90770
        
        Case 90769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171050)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 90771

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 90772

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 90773

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 90774
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 90771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 90772 To 90774

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171051)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 90775

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 90776

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError 90777
    
        lErro = Define_Padrao()
        If lErro <> SUCESSO Then gError 90778
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 90775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 90776 To 90778

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171052)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 90779
    
    Call Limpa_Tela(Me)
    
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 90780
    
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 90779, 90780
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171053)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 90781
    
    gobjRelatorio.sNomeTsk = "PedVenda"

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 90781

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171054)

    End Select

    Exit Sub

End Sub

Private Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
        
    giPedidoInicial = 1
            
    Call Limpa_Tela(Me)
    
    ComboOpcoes.Text = ""
        
    Define_Padrao = SUCESSO

    Exit Function

Erro_Define_Padrao:

    Define_Padrao = gErr

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171055)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 90782
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 90783
   
    lErro = objRelOpcoes.IncluirParametro("NPEDIDOINIC", PedidoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90784
    
    lErro = objRelOpcoes.IncluirParametro("NPEDIDOFIM", PedidoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90785
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINI", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 90786

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 90787
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError 90788
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr


    Select Case gErr

        Case 90782 To 90788
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171056)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 90789
            
    'pega parâmetro Pedido Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NPEDIDOINIC", sParam)
    If lErro Then gError 90790
    
    PedidoInicial.Text = sParam
    
    'pega parâmetro Pedido Final e exibe
    lErro = objRelOpcoes.ObterParametro("NPEDIDOFIM", sParam)
    If lErro Then gError 90809
    
    PedidoFinal.Text = sParam
            
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINI", sParam)
    If lErro <> SUCESSO Then gError 90791

    Call DateParaMasked(DataInicial, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 90792

    Call DateParaMasked(DataFinal, CDate(sParam))
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 90789 To 90792
        
        Case 90809
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171057)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

       
    'Pedido inicial não pode ser maior que o Pedido final
    If Trim(PedidoInicial.Text) <> "" And Trim(PedidoFinal.Text) <> "" Then
    
         If CLng(PedidoInicial.Text) > CLng(PedidoFinal.Text) Then gError 90793
         
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then gError 90794
    
    End If
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
            
        Case 90793
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", gErr)
            PedidoInicial.SetFocus
        
        Case 90794
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataInicial.SetFocus
               
         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171058)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iRegVendaInc As Integer
Dim iRegVendaFin As Integer

On Error GoTo Erro_Monta_Expressao_Selecao

    If Trim(PedidoInicial.Text) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PedidoVenda >= " & Forprint_ConvLong(CLng(PedidoInicial.Text))
        
    End If

    If PedidoFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PedidoVenda <= " & Forprint_ConvLong(CLng(PedidoFinal.Text))

    End If
    
    If Trim(DataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(CDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataFinal.Text))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171059)

    End Select

    Exit Function

End Function

Private Sub LabelPedInicial_Click()

Dim lErro As Long
Dim objOp As ClassPedidoDeVenda
Dim colSelecao As Collection

On Error GoTo Erro_LabelPedInicial_Click

    giPedidoInicial = 1

    If Len(Trim(PedidoInicial.Text)) <> 0 Then
    
        lErro = Long_Critica(PedidoInicial.Text)
        If lErro <> SUCESSO Then gError 90795
        
        Set objOp = New ClassPedidoDeVenda
        objOp.lCodigo = CLng(PedidoInicial.Text)

    End If

    Call Chama_Tela("PedidoVendaListaModal", colSelecao, objOp, objEventoOp)
    
    Exit Sub

Erro_LabelPedInicial_Click:

    Select Case gErr
    
        Case 90795

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171060)

    End Select

    Exit Sub

End Sub

Private Sub LabelPedFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOp As ClassPedidoDeVenda

On Error GoTo Erro_LabelPedFinal_Click

    giPedidoInicial = 0

    If Len(Trim(PedidoFinal.Text)) <> 0 Then
    
        lErro = Long_Critica(PedidoFinal.Text)
        If lErro <> SUCESSO Then gError 90796

        Set objOp = New ClassPedidoDeVenda
        objOp.lCodigo = CLng(PedidoFinal.Text)

    End If

    Call Chama_Tela("PedidoVendaListaModal", colSelecao, objOp, objEventoOp)
   
   Exit Sub

Erro_LabelPedFinal_Click:

    Select Case gErr
    
        Case 90796

        Case Else
           lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171061)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOp_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassPedidoDeVenda

On Error GoTo Erro_objEventoOp_evSelecao

    Set objOp = obj1

    If giPedidoInicial = 1 Then
        PedidoInicial.Text = CStr(objOp.lCodigo)
    Else
        PedidoFinal.Text = CStr(objOp.lCodigo)
    End If

    Exit Sub

Erro_objEventoOp_evSelecao:

    Select Case gErr

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171062)

    End Select

    Exit Sub

End Sub

Private Sub PedidoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim ObjPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoInicial_Validate

    giPedidoInicial = 1
    
    If Len(Trim(PedidoInicial.Text)) > 0 Then
        
        lErro = Long_Critica(PedidoInicial.Text)
        If lErro <> SUCESSO Then gError 90797
    
        ObjPedidoVenda.lCodigo = CLng(PedidoInicial.Text)
        ObjPedidoVenda.iFilialEmpresa = giFilialEmpresa
       
        lErro = CF("PedidoDeVenda_Le",ObjPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then gError 90798
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 90799
        
    End If
       
    Exit Sub

Erro_PedidoInicial_Validate:

    Cancel = True

    Select Case gErr
    
        Case 90797, 90798
        
        Case 90799
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, ObjPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171063)

    End Select

    Exit Sub

End Sub

Private Sub PedidoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim ObjPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoFinal_Validate

    giPedidoInicial = 0
    
    If Len(Trim(PedidoFinal.Text)) > 0 Then
    
        lErro = Long_Critica(PedidoFinal.Text)
        If lErro <> SUCESSO Then gError 90800
    
        ObjPedidoVenda.lCodigo = CLng(PedidoFinal.Text)
        ObjPedidoVenda.iFilialEmpresa = giFilialEmpresa
       
        lErro = CF("PedidoDeVenda_Le",ObjPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then gError 90801
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 90802
        
    End If
       
    Exit Sub

Erro_PedidoFinal_Validate:

    Cancel = True

    Select Case gErr
    
        Case 90800, 90801
                
        Case 90802
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, ObjPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171064)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then gError 90803

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 90803

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171065)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then gError 90804

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 90804

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171066)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 90805

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 90805
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171067)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 90806

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 90806
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171068)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 90807

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 90807
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171069)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 90808

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 90808
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171070)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Relação dos Pedidos de Venda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPedVenda"
    
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
        
        If Me.ActiveControl Is PedidoInicial Then
            Call LabelPedInicial_Click
        ElseIf Me.ActiveControl Is PedidoFinal Then
            Call LabelPedFinal_Click
        End If
   
    End If

End Sub
Private Sub LabelPedInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedInicial, Source, X, Y)
End Sub

Private Sub LabelPedInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedInicial, Button, Shift, X, Y)
End Sub

Private Sub LabelPedFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedFinal, Source, X, Y)
End Sub

Private Sub LabelPedFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedFinal, Button, Shift, X, Y)
End Sub

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub
