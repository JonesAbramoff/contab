VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpCadVendOcx 
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3390
   ScaleWidth      =   7845
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5550
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpCadVendOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpCadVendOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpCadVendOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpCadVendOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vendedores"
      Height          =   825
      Left            =   120
      TabIndex        =   14
      Top             =   1935
      Width           =   5355
      Begin MSMask.MaskEdBox VendInicial 
         Height          =   300
         Left            =   615
         TabIndex        =   4
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VendFinal 
         Height          =   300
         Left            =   3240
         TabIndex        =   5
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedorDe 
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   353
         Width           =   315
      End
      Begin VB.Label LabelVendedorAte 
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
         Left            =   2805
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Top             =   353
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpCadVendOcx.ctx":0994
      Left            =   1320
      List            =   "RelOpCadVendOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
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
      Left            =   5655
      Picture         =   "RelOpCadVendOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   825
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Vendedor"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   5355
      Begin VB.ComboBox ComboTipo 
         Height          =   315
         Left            =   1890
         TabIndex        =   3
         Top             =   600
         Width           =   3225
      End
      Begin VB.OptionButton OptionUmTipo 
         Caption         =   "Apenas do Tipo"
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
         Left            =   150
         TabIndex        =   2
         Top             =   630
         Width           =   1755
      End
      Begin VB.OptionButton OptionTodosTipos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   1
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpCadVendOcx.ctx":0A9A
      Left            =   1515
      List            =   "RelOpCadVendOcx.ctx":0AA4
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2925
      Width           =   3270
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
      Left            =   615
      TabIndex        =   18
      Top             =   300
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ordenados Por:"
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
      Left            =   120
      TabIndex        =   17
      Top             =   2985
      Width           =   1335
   End
End
Attribute VB_Name = "RelOpCadVendOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoVendInic As AdmEvento
Attribute objEventoVendInic.VB_VarHelpID = -1
Private WithEvents objEventoVendFim As AdmEvento
Attribute objEventoVendFim.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 47687
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 47691
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 47691
        
        Case 47687
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167473)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 56728
    
    Call OptionTodosTipos_Click
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case Err
    
        Case 56728
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167474)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub

Private Sub ComboTipo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_ComboTipo_Validate

    'Verifica se foi preenchida a ComboBox Tipo
    If Len(Trim(ComboTipo.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Tipo
    If ComboTipo.Text = ComboTipo.List(ComboTipo.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ComboTipo, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 56722

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro <> SUCESSO Then Error 56723

    Exit Sub

Erro_ComboTipo_Validate:

    Cancel = True


    Select Case Err

        Case 56722 'Tratado na rotina chamada
    
        Case 56723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_VENDEDOR_NAO_ENCONTRADO2", Err, ComboTipo.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 167475)

    End Select

    Exit Sub

End Sub

Private Sub VendFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendFinal_Validate

    If Len(Trim(VendFinal.Text)) > 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendFinal, objVendedor, 0)
        If lErro <> SUCESSO Then Error 47689

    End If
    
    Exit Sub

Erro_VendFinal_Validate:

    Cancel = True


    Select Case Err

        Case 47689
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 167476)

    End Select

End Sub

Private Sub VendInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendInicial_Validate

    If Len(Trim(VendInicial.Text)) > 0 Then
   
        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendInicial, objVendedor, 0)
        If lErro <> SUCESSO Then Error 47690

    End If
        
    Exit Sub

Erro_VendInicial_Validate:

    Cancel = True


    Select Case Err

        Case 47690
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 167477)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoVendInic = New AdmEvento
    Set objEventoVendFim = New AdmEvento
       
    lErro = PreencheComboTipo()
    If lErro <> SUCESSO Then Error 47692
        
    ComboOrdenacao.ListIndex = 0
    
    OptionTodosTipos_Click
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 47692
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167478)

    End Select

    Exit Sub

End Sub

Function PreencheComboTipo() As Long

Dim lErro As Long
Dim colCodigoDescricaoVendedor As New AdmColCodigoNome
Dim objCodigoDescricaoVendedor As New AdmCodigoNome

On Error GoTo Erro_PreencheComboTipo

    lErro = CF("Cod_Nomes_Le","TiposdeVendedor", "Codigo", "Descricao", STRING_TIPO_DE_VENDEDOR_DESCRICAO, colCodigoDescricaoVendedor)
    If lErro <> SUCESSO Then Error 47694
    
   'preenche a ListBox ComboTipo com os objetos da colecao
    For Each objCodigoDescricaoVendedor In colCodigoDescricaoVendedor
        ComboTipo.AddItem objCodigoDescricaoVendedor.iCodigo & SEPARADOR & objCodigoDescricaoVendedor.sNome
        ComboTipo.ItemData(ComboTipo.NewIndex) = objCodigoDescricaoVendedor.iCodigo
    Next
        
    PreencheComboTipo = SUCESSO

    Exit Function
    
Erro_PreencheComboTipo:

    PreencheComboTipo = Err

    Select Case Err

    Case 47694
    
    Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167479)

    End Select

    Exit Function

End Function

Private Sub LabelVendedorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_LabelVendedorAte_Click
    
    If Len(Trim(VendFinal.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendFinal.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendFim)

   Exit Sub

Erro_LabelVendedorAte_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167480)

    End Select

    Exit Sub

End Sub

Private Sub LabelVendedorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_LabelVendedorDe_Click
    
    If Len(Trim(VendInicial.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendInicial.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendInic)

   Exit Sub

Erro_LabelVendedorDe_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167481)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoVendFim_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    VendFinal.Text = CStr(objVendedor.iCodigo)
    Call VendFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoVendInic_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    VendInicial.Text = CStr(objVendedor.iCodigo)
    Call VendInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub OptionTodosTipos_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_OptionTodosTipos_Click

    ComboTipo.ListIndex = -1
    ComboTipo.Enabled = False
    OptionTodosTipos.Value = True
    
            
    Exit Sub

Erro_OptionTodosTipos_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167482)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 47695

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47696

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 47697
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47698
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 47695
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 47696, 47697, 47698, 47897
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167483)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 47699

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 47700

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 47699
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 47700, 47701

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167484)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47702

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 47702

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167485)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sVendedor_I As String
Dim sVendedor_F As String
Dim sOrdenacaoPor As String
Dim sCheckTipo As String
Dim sVendedorTipo As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sVendedor_I, sVendedor_F, sCheckTipo, sVendedorTipo)
    If lErro <> SUCESSO Then Error 47707

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 47708
         
    lErro = objRelOpcoes.IncluirParametro("NVENDINIC", sVendedor_I)
    If lErro <> AD_BOOL_TRUE Then Error 47709
         
    lErro = objRelOpcoes.IncluirParametro("TVENDINIC", VendInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54758
    
    lErro = objRelOpcoes.IncluirParametro("NVENDFIM", sVendedor_F)
    If lErro <> AD_BOOL_TRUE Then Error 47710
    
    lErro = objRelOpcoes.IncluirParametro("TVENDFIM", VendFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54759
        
    Select Case ComboOrdenacao.ListIndex
        
        Case ORD_POR_CODIGO
        
            sOrdenacaoPor = "CodVend"
            
        Case ORD_POR_NOME
            
            sOrdenacaoPor = "NomeVend"
        
        Case Else
            Error 47711
                  
    End Select
    
    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then Error 47712
        
    lErro = objRelOpcoes.IncluirParametro("TTIPOVEND", sVendedorTipo)
    If lErro <> AD_BOOL_TRUE Then Error 47713
    
    lErro = objRelOpcoes.IncluirParametro("TTVENDEDOR", ComboTipo.Text)
    If lErro <> AD_BOOL_TRUE Then Error 54757

    lErro = objRelOpcoes.IncluirParametro("TOPTIPO", sCheckTipo)
    If lErro <> AD_BOOL_TRUE Then Error 47714
       
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sVendedor_I, sVendedor_F, sVendedorTipo, sCheckTipo, sOrdenacaoPor)
    If lErro <> SUCESSO Then Error 47715

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 47707, 47708, 47709, 47710, 47711, 47712, 47713, 47714, 47715
        
        Case 54757, 54758, 54759
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167486)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sVendedor_I As String, sVendedor_F As String, sCheckTipo As String, sVendedorTipo As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'critica Vendedor Inicial e Final
    If VendInicial.Text <> "" Then
        sVendedor_I = CStr(LCodigo_Extrai(VendInicial.Text))
    Else
        sVendedor_I = ""
    End If
    
    If VendFinal.Text <> "" Then
        sVendedor_F = CStr(LCodigo_Extrai(VendFinal.Text))
    Else
        sVendedor_F = ""
    End If
            
    If sVendedor_I <> "" And sVendedor_F <> "" Then
        
        If CInt(sVendedor_I) > CInt(sVendedor_F) Then Error 47716
        
    End If
            
    'Se a opção para todos os Vendedores estiver selecionada
    If OptionTodosTipos.Value = True Then
        sCheckTipo = "Todos"
        sVendedorTipo = ""
    
    'Se a opção para apenas um vendedor estiver selecionada
    Else
        If ComboTipo.Text = "" Then Error 47717
        sCheckTipo = "Um"
        sVendedorTipo = CStr(Codigo_Extrai(ComboTipo.Text))
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
                
        Case 47716
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", Err)
            VendInicial.SetFocus
                
        Case 47717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_VENDEDOR_NAO_PREENCHIDO", Err)
            ComboTipo.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167487)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sVendedor_I As String, sVendedor_F As String, sVendedorTipo As String, sCheckTipo As String, sOrdenacaoPor As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sVendedor_I <> "" Then sExpressao = "Vendedor >= " & Forprint_ConvInt(CInt(sVendedor_I))

   If sVendedor_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vendedor <= " & Forprint_ConvInt(CInt(sVendedor_F))

    End If
           
    'Se a opção para apenas um Vendedor estiver selecionada
    If sCheckTipo = "Um" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoVendedor = " & Forprint_ConvInt(CInt(sVendedorTipo))

    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167488)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sTipoVendedor As String, iTipo As Integer
Dim sOrdenacaoPor As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 47718
   
    'pega Vendedor inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDINIC", sParam)
    If lErro <> SUCESSO Then Error 47719
    
    VendInicial.Text = sParam
    Call VendInicial_Validate(bSGECancelDummy)
    
    'pega  Vendedor final e exibe
    lErro = objRelOpcoes.ObterParametro("NVENDFIM", sParam)
    If lErro <> SUCESSO Then Error 47720
    
    VendFinal.Text = sParam
    Call VendFinal_Validate(bSGECancelDummy)
                
    'pega  Vendedor final e exibe
    lErro = objRelOpcoes.ObterParametro("TOPTIPO", sParam)
    If lErro <> SUCESSO Then Error 47721
    
    If sParam = "Todos" Then
    
        Call OptionTodosTipos_Click
        
    Else
    
        'pega tipo de Vendedor
        lErro = objRelOpcoes.ObterParametro("TTIPOVEND", sTipoVendedor)
        If lErro <> SUCESSO Then Error 47722
                       
        OptionUmTipo.Value = True
        ComboTipo.Enabled = True
        ComboTipo.Text = sTipoVendedor
        Call Combo_Seleciona(ComboTipo, iTipo)
     
    End If
                        
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then Error 47723
        
    Select Case sOrdenacaoPor
        
        Case "CodVend"
        
            ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
        Case "NomeVend"
            
            ComboOrdenacao.ListIndex = ORD_POR_NOME
                            
        Case Else
            Error 47724
                  
    End Select
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 47718, 47719, 47720, 47721, 47722, 47723, 47724
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167489)

    End Select

    Exit Function

End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub OptionUmTipo_Click()

Dim lErro As Long

On Error GoTo Erro_OptionUmTipo_Click

    ComboTipo.ListIndex = -1
    ComboTipo.Enabled = True
    ComboTipo.SetFocus
    
    Exit Sub

Erro_OptionUmTipo_Click:

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167490)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoVendInic = Nothing
    Set objEventoVendFim = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_RELOP_VEND
    Set Form_Load_Ocx = Me
    Caption = "Relação de Vendedores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpCadVend"
    
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
        
        If Me.ActiveControl Is VendInicial Then
            Call LabelVendedorDe_Click
        ElseIf Me.ActiveControl Is VendFinal Then
            Call LabelVendedorAte_Click
        End If
    
    End If

End Sub


Private Sub LabelVendedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorDe, Source, X, Y)
End Sub

Private Sub LabelVendedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorDe, Button, Shift, X, Y)
End Sub

Private Sub LabelVendedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVendedorAte, Source, X, Y)
End Sub

Private Sub LabelVendedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVendedorAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

