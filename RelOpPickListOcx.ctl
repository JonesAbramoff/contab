VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPickListOcx 
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   8700
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6360
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPickListOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPickListOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPickListOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPickListOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
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
      Left            =   4530
      Picture         =   "RelOpPickListOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPickListOcx.ctx":0A96
      Left            =   1380
      List            =   "RelOpPickListOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2916
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ordem de Produção"
      Height          =   705
      Left            =   150
      TabIndex        =   22
      Top             =   1035
      Width           =   5640
      Begin VB.TextBox OpFinal 
         Height          =   300
         Left            =   3405
         TabIndex        =   2
         Top             =   285
         Width           =   1680
      End
      Begin VB.TextBox OpInicial 
         Height          =   300
         Left            =   825
         TabIndex        =   1
         Top             =   315
         Width           =   1680
      End
      Begin VB.Label LabelOpFinal 
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
         Left            =   2955
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   330
         Width           =   360
      End
      Begin VB.Label LabelOpInicial 
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
         Left            =   465
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   2130
      Width           =   5670
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1845
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   840
         TabIndex        =   3
         Top             =   285
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   4425
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   3420
         TabIndex        =   4
         Top             =   285
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
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
         Left            =   3015
         TabIndex        =   21
         Top             =   345
         Width           =   360
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
         Height          =   240
         Left            =   435
         TabIndex        =   20
         Top             =   315
         Width           =   345
      End
   End
   Begin VB.ListBox Almoxarifados 
      Height          =   2985
      ItemData        =   "RelOpPickListOcx.ctx":0A9A
      Left            =   6000
      List            =   "RelOpPickListOcx.ctx":0A9C
      TabIndex        =   7
      Top             =   1095
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Almoxarifados"
      Height          =   840
      Left            =   150
      TabIndex        =   14
      Top             =   3225
      Width           =   5655
      Begin MSMask.MaskEdBox AlmoxarifadoInicial 
         Height          =   315
         Left            =   690
         TabIndex        =   5
         Top             =   315
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AlmoxarifadoFinal 
         Height          =   315
         Left            =   3345
         TabIndex        =   6
         Top             =   315
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label labelAlmoxarifadoFinal 
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
         Left            =   2925
         TabIndex        =   16
         Top             =   375
         Width           =   360
      End
      Begin VB.Label Label2 
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
         Left            =   315
         TabIndex        =   15
         Top             =   375
         Width           =   315
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
      Left            =   675
      TabIndex        =   26
      Top             =   315
      Width           =   615
   End
   Begin VB.Label LabelAlmoxarifado 
      AutoSize        =   -1  'True
      Caption         =   "Almoxarifados"
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
      Left            =   5985
      TabIndex        =   25
      Top             =   855
      Width           =   1185
   End
End
Attribute VB_Name = "RelOpPickListOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoOpInic As AdmEvento
Attribute objEventoOpInic.VB_VarHelpID = -1
Private WithEvents objEventoOpFim As AdmEvento
Attribute objEventoOpFim.VB_VarHelpID = -1

Dim giOp_Inicial As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 47608
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 47572
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 47572
        
        Case 47608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171121)

    End Select

    Exit Function

End Function
Private Sub AlmoxarifadoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoFinal_Validate

    If Len(Trim(AlmoxarifadoFinal.Text)) > 0 Then
       
        'Tenta ler o Almoxarifado (NomeReduzido ou Código)
        lErro = TP_Almoxarifado_Le_ComCodigo(AlmoxarifadoFinal, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 47566
      
    End If
    
    Almoxarifados.Tag = "Amoxarifado_Final"
    
    Exit Sub

Erro_AlmoxarifadoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 47566
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171122)

    End Select
    
End Sub

Private Sub AlmoxarifadoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoInicial_Validate

    If Len(Trim(AlmoxarifadoInicial.Text)) > 0 Then
       
        'Tenta ler o Almoxarifado (NomeReduzido ou Código)
        lErro = TP_Almoxarifado_Le_ComCodigo(AlmoxarifadoInicial, objAlmoxarifado)
        If lErro <> SUCESSO Then Error 47567
            
    End If
    
    Almoxarifados.Tag = "Amoxarifado_Inicial"
    
    Exit Sub

Erro_AlmoxarifadoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 47567
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171123)

    End Select

End Sub
Private Sub Almoxarifados_DblClick()
'Preenche Almoxarifado Final ou Inicial com o almoxarifado selecionado

Dim lErro As Long

On Error GoTo Erro_Almoxarifados_DblClick

    If Almoxarifados.Tag = "Amoxarifado_Inicial" Then
  
        AlmoxarifadoInicial.Text = Almoxarifados.List(Almoxarifados.ListIndex)
    
    Else
        If Almoxarifados.Tag = "Amoxarifado_Final" Then
    
            AlmoxarifadoFinal.Text = Almoxarifados.List(Almoxarifados.ListIndex)
        
        End If
    
    End If
    
    Exit Sub

Erro_Almoxarifados_DblClick:

    Select Case Err

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171124)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 47568

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 47569

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47570
        
''        lErro = Define_Padrao()
''        If lErro <> SUCESSO Then Error 47945
    
        ComboOpcoes.Text = ""
           
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 47568
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 47569, 47570, 47945

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171125)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47571

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 47571

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171126)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 47572

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47573

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 47574

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47575
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 47572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 47573, 47574, 47575, 47576, 47944

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171127)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47577
    
    ComboOpcoes.Text = ""
           
''    lErro = Define_Padrao()
''    If lErro <> SUCESSO Then Error 47943
           
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47943
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171128)

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
        If lErro <> SUCESSO Then Error 47570

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 47570

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171129)

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
        If lErro <> SUCESSO Then Error 47571

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 47571

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171130)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load
    
    Set objEventoOpInic = New AdmEvento
    Set objEventoOpFim = New AdmEvento
    
''    'Preenche as combos de filial Empresa guardando no itemData o codigo
''    lErro = Carrega_FilialEmpresa()
''    If lErro <> SUCESSO Then Error 47573

    'carrega a ListBox Almoxarifados
    lErro = Carrega_Lista_Almoxarifado()
    If lErro <> SUCESSO Then Error 47574
    
''    lErro = Define_Padrao()
''    If lErro <> SUCESSO Then Error 47942
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 47573, 47574, 47942

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171131)

    End Select

    Exit Sub

End Sub
''
''Private Function Carrega_FilialEmpresa() As Long
'''Carrega as Combos FilialEmpresaInicial e FilialEmpresaFinal
''
''Dim lErro As Long
''Dim objCodigoNome As New AdmCodigoNome
''Dim iIndice As Integer
''Dim colCodigoDescricao As New AdmColCodigoNome
''
''On Error GoTo Erro_Carrega_FilialEmpresa
''
''    'Lê Códigos e NomesReduzidos da tabela FilialEmpresa e devolve na coleção
''    lErro = CF("Cod_Nomes_Le","FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAL_NOME, colCodigoDescricao)
''    If lErro <> SUCESSO Then Error 47575
''
''    'preenche as combos iniciais e finais
''    For Each objCodigoNome In colCodigoDescricao
''
''        If objCodigoNome.iCodigo <> 0 Then
''            FilialEmpresaInicial.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
''            FilialEmpresaInicial.ItemData(FilialEmpresaInicial.NewIndex) = objCodigoNome.iCodigo
''
''            FilialEmpresaFinal.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
''            FilialEmpresaFinal.ItemData(FilialEmpresaFinal.NewIndex) = objCodigoNome.iCodigo
''        End If
''
''    Next
''
''    Carrega_FilialEmpresa = SUCESSO
''
''    Exit Function
''
''Erro_Carrega_FilialEmpresa:
''
''    Carrega_FilialEmpresa = Err
''
''    Select Case Err
''
''        Case 47575
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171132)
''
''    End Select
''
''    Exit Function
''
''End Function

Private Function Carrega_Lista_Almoxarifado() As Long
'Carrega a ListBox Almoxarifados

Dim lErro As Long
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Carrega_Lista_Almoxarifado
    
    'Lê Códigos e NomesReduzidos da tabela Almoxarifado e devolve na coleção
    lErro = CF("Almoxarifados_Le_FilialEmpresa",giFilialEmpresa, colAlmoxarifados)
    If lErro <> SUCESSO Then Error 47576

    'Preenche a ListBox AlmoxarifadoList com os objetos da coleção
    For Each objAlmoxarifado In colAlmoxarifados
        Almoxarifados.AddItem objAlmoxarifado.iCodigo & SEPARADOR & objAlmoxarifado.sNomeReduzido
        Almoxarifados.ItemData(Almoxarifados.NewIndex) = objAlmoxarifado.iCodigo
    Next

    Carrega_Lista_Almoxarifado = SUCESSO

    Exit Function

Erro_Carrega_Lista_Almoxarifado:

    Carrega_Lista_Almoxarifado = Err

    Select Case Err

        Case 47576

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171133)

    End Select

    Exit Function

End Function
''
''Private Sub FilialEmpresaInicial_Validate(Cancel As Boolean)
'''Busca a filial com código digitado na lista FilialEmpresa
''
''Dim lErro As Long
''Dim iCodigo As Integer
''
''On Error GoTo Erro_FilialEmpresaInicial_Validate
''
''    'se uma opcao da lista estiver selecionada, OK
''    If FilialEmpresaInicial.ListIndex <> -1 Then Exit Sub
''
''    If Len(Trim(FilialEmpresaInicial.Text)) = 0 Then Exit Sub
''
''    lErro = Combo_Seleciona(FilialEmpresaInicial, iCodigo)
''    If lErro <> SUCESSO Then Error 47577
''
''    Exit Sub
''
''Erro_FilialEmpresaInicial_Validate:

''    Cancel = True

''
''    Select Case Err
''
''        Case 47577
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171134)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Private Sub FilialEmpresaFinal_Validate(Cancel As Boolean)
'''Busca a filial com código digitado na lista FilialEmpresa
''
''Dim lErro As Long
''Dim iCodigo As Integer
''
''On Error GoTo Erro_FilialEmpresaFinal_Validate
''
''    'se uma opcao da lista estiver selecionada, OK
''    If FilialEmpresaFinal.ListIndex <> -1 Then Exit Sub
''
''    If Len(Trim(FilialEmpresaFinal.Text)) = 0 Then Exit Sub
''
''    lErro = Combo_Seleciona(FilialEmpresaFinal, iCodigo)
''    If lErro <> SUCESSO Then Error 47578
''
''    Exit Sub
''
''Erro_FilialEmpresaFinal_Validate:

''    Cancel = True

''
''    Select Case Err
''
''        Case 47578
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171135)
''
''    End Select
''
''    Exit Sub
''
''End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoOpInic = Nothing
    Set objEventoOpFim = Nothing
    
End Sub

Private Sub LabelOpFinal_Click()

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao
Dim colSelecao As Collection

On Error GoTo Erro_LabelOpFinal_Click
  
    If Len(Trim(OpFinal.Text)) <> 0 Then

        objOp.sCodigo = OpFinal.Text

    End If

    Call Chama_Tela("OrdProdTodasListaModal", colSelecao, objOp, objEventoOpFim)
    
    Exit Sub

Erro_LabelOpFinal_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171136)

    End Select

    Exit Sub

End Sub

Private Sub LabelOpInicial_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_LabelOpInicial_Click
    
    If Len(Trim(OpInicial.Text)) <> 0 Then
        
        objOp.sCodigo = OpInicial.Text

    End If

    Call Chama_Tela("OrdProdTodasListaModal", colSelecao, objOp, objEventoOpInic)
   
   Exit Sub

Erro_LabelOpInicial_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171137)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOpFim_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOpFim_evSelecao

    Set objOp = obj1
    
    OpFinal.Text = objOp.sCodigo
        
    Me.Show
    
    Exit Sub

Erro_objEventoOpFim_evSelecao:

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171138)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOpInic_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOpInic_evSelecao

    Set objOp = obj1
    
    OpInicial.Text = objOp.sCodigo
    
    Me.Show
    
    Exit Sub

Erro_objEventoOpInic_evSelecao:

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171139)

    End Select

    Exit Sub

End Sub

Private Function Valida_OrdProd(sCodigoOP As String) As Long

Dim objOp As New ClassOrdemDeProducao
Dim lErro As Long

On Error GoTo Erro_Valida_OrdProd

    objOp.iFilialEmpresa = giFilialEmpresa
    objOp.sCodigo = sCodigoOP

    lErro = CF("OrdemDeProducao_Le_SemItens",objOp)
    If lErro <> SUCESSO And lErro <> 34455 Then Error 59545

    If lErro = 34455 Then Error 59546

    Valida_OrdProd = SUCESSO

    Exit Function

Erro_Valida_OrdProd:

    Valida_OrdProd = Err

    Select Case Err

        Case 59545
        
        Case 59546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_ABERTA_INEXISTENTE", Err)
   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171140)

    End Select

    Exit Function

End Function

Private Sub OpInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpInicial_Validate

    giOp_Inicial = 1

    If Len(Trim(OpInicial.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpInicial.Text)
        If lErro <> SUCESSO Then Error 59547
        
    End If

    Exit Sub

Erro_OpInicial_Validate:

    Cancel = True


    Select Case Err

        Case 59547
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171141)

    End Select

    Exit Sub

End Sub

Private Sub OpFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpFinal_Validate

    giOp_Inicial = 0

    If Len(Trim(OpFinal.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpFinal.Text)
        If lErro <> SUCESSO Then Error 59548
    
    End If

    Exit Sub

Erro_OpFinal_Validate:

    Cancel = True


    Select Case Err

        Case 59548
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171142)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47579

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 47579
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171143)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47580

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 47580
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171144)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47581

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 47581
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171145)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47582

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 47582
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171146)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sAlmox_I As String
Dim sAlmox_F As String
''Dim sFilial_I As String
''Dim sFilial_F As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sAlmox_I, sAlmox_F)
    If lErro <> SUCESSO Then Error 47583

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 47584

    lErro = objRelOpcoes.IncluirParametro("TOPINIC", OpInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47585

    lErro = objRelOpcoes.IncluirParametro("TOPFIM", OpFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47586
    
    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 47587

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 47588
        
''    lErro = objRelOpcoes.IncluirParametro("NFILIALINIC", sFilial_I)
''    If lErro <> AD_BOOL_TRUE Then Error
''
''    lErro = objRelOpcoes.IncluirParametro("NFILIALFIM", sFilial_F)
''    If lErro <> AD_BOOL_TRUE Then Error
    
    lErro = objRelOpcoes.IncluirParametro("NALMOXINIC", sAlmox_I)
    If lErro <> AD_BOOL_TRUE Then Error 47591
    
    lErro = objRelOpcoes.IncluirParametro("TALMOXINICIAL", AlmoxarifadoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47589
        
    lErro = objRelOpcoes.IncluirParametro("NALMOXFIM", sAlmox_F)
    If lErro <> AD_BOOL_TRUE Then Error 47592
    
    lErro = objRelOpcoes.IncluirParametro("TALMOXFINAL", AlmoxarifadoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 47590
       
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sAlmox_I, sAlmox_F)
    If lErro <> SUCESSO Then Error 47593

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 47583, 47584, 47585, 47586, 47587, 47588, 47589, 47590, 47591, 47592, 47593

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171147)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sAlmox_I As String, sAlmox_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long


On Error GoTo Erro_Monta_Expressao_Selecao

   If OpInicial.Text <> "" Then sExpressao = "OrdemProducao >= " & Forprint_ConvTexto(OpInicial.Text)

    If OpFinal.Text <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "OrdemProducao <= " & Forprint_ConvTexto(OpFinal.Text)

    End If


     If sAlmox_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Almoxarifado >= " & Forprint_ConvInt(CInt(sAlmox_I))

    End If

    If sAlmox_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Almoxarifado <= " & Forprint_ConvInt(CInt(sAlmox_F))

    End If
    
''     If sFilial_I <> "" Then
''
''        If sExpressao <> "" Then sExpressao = sExpressao & " E "
''        sExpressao = sExpressao & "FilialEmpresa >= " & Forprint_ConvInt(CInt(sFilial_I))
''
''    End If
''
''    If sFilial_F <> "" Then
''
''        If sExpressao <> "" Then sExpressao = sExpressao & " E "
''        sExpressao = sExpressao & "FilialEmpresa <= " & Forprint_ConvInt(CInt(sFilial_F))
''
''    End If
        
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

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171148)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sAlmox_I As String, sAlmox_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
   'data inicial não pode ser maior que a data final
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 47594
    
    End If
      
    If AlmoxarifadoInicial.Text <> "" And AlmoxarifadoFinal.Text <> "" Then
        
        'critica Almoxarifado Inicial e Final
        sAlmox_I = CStr(Codigo_Extrai(AlmoxarifadoInicial.Text))
        sAlmox_F = CStr(Codigo_Extrai(AlmoxarifadoFinal.Text))

        If CInt(sAlmox_I) > CInt(sAlmox_F) Then Error 47595
        
    End If
    
    'Ordem de Produção inicial não pode ser maior que a final
    If Trim(OpInicial.Text) <> "" And Trim(OpFinal.Text) <> "" Then

        If OpInicial.Text > OpFinal.Text Then Error 47596

    End If
    
''    If giFilialEmpresa <> EMPRESA_TODA Then
''
''        sFilial_I = CStr(giFilialEmpresa)
''        sFilial_F = CStr(giFilialEmpresa)
''
''    Else
''
''        'critica FilialEmpresa Inicial e Final
''        If FilialEmpresaInicial.ListIndex <> -1 Then
''            sFilial_I = CStr(FilialEmpresaInicial.ItemData(FilialEmpresaInicial.ListIndex))
''        Else
''            sFilial_I = ""
''        End If
''
''        If FilialEmpresaFinal.ListIndex <> -1 Then
''            sFilial_F = CStr(FilialEmpresaFinal.ItemData(FilialEmpresaFinal.ListIndex))
''        Else
''            sFilial_F = ""
''        End If
''
''        If sFilial_I <> "" And sFilial_F <> "" Then
''
''            If CInt(sFilial_I) > CInt(sFilial_F) Then Error 47597
''
''        End If
''
''    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err
        
        Case 47594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataInicial.SetFocus
        
        Case 47595
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INICIAL_MAIOR", Err)
            AlmoxarifadoInicial.SetFocus
        
        Case 47596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INICIAL_MAIOR", Err)
            OpInicial.SetFocus
            
''        Case 47597
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_INICIAL_MAIOR", Err)
''            FilialEmpresaInicial.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171149)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iFilialInic, iFilialFim As Integer
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 47599
    
    'pega Ordem de Producao Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TOPINIC", sParam)
    If lErro <> SUCESSO Then Error 47600

    OpInicial.Text = sParam

    'pega Ordem de Producao Final e exibe
    lErro = objRelOpcoes.ObterParametro("TOPFIM", sParam)
    If lErro <> SUCESSO Then Error 47601

    OpFinal.Text = sParam
    
    'pega parâmetro Almoxarifado Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NALMOXINIC", sParam)
    If lErro <> SUCESSO Then Error 47602
    
    AlmoxarifadoInicial.Text = sParam
    Call AlmoxarifadoInicial_Validate(bSGECancelDummy)
    
    'pega parâmetro Almoxarifado Final e exibe
    lErro = objRelOpcoes.ObterParametro("NALMOXFIM", sParam)
    If lErro <> SUCESSO Then Error 47603
    
    AlmoxarifadoFinal.Text = sParam
    Call AlmoxarifadoFinal_Validate(bSGECancelDummy)
    
''    If giFilialEmpresa <> EMPRESA_TODA Then
''
''        'Preenche em Branco
''        FilialEmpresaInicial.ListIndex = -1
''        FilialEmpresaFinal.ListIndex = -1
''
''        'desabilita a combo
''        FilialEmpresaInicial.Enabled = False
''        FilialEmpresaFinal.Enabled = False
''
''    Else
''
''        'pega parâmetro FilialEmpresa Inicial
''        lErro = objRelOpcoes.ObterParametro("NFILIALINIC", sParam)
''        If lErro <> SUCESSO Then Error 47604
''
''        FilialEmpresaInicial.Text = sParam
''        Call FilialEmpresaInicial_Validate(bSGECancelDummy)
''
''        'pega parâmetro FilialEmpresa Final
''        lErro = objRelOpcoes.ObterParametro("NFILIALFIM", sParam)
''        If lErro <> SUCESSO Then Error 47605
''
''        FilialEmpresaFinal.Text = sParam
''        Call FilialEmpresaFinal_Validate(bSGECancelDummy)
''
''    End If
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 47606

    If sParam <> "07/09/1822" Then Call DateParaMasked(DataInicial, CDate(sParam))
    'DataInicial.PromptInclude = False
    'DataInicial.Text = sParam
    'DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 47607

    If sParam <> "07/09/1822" Then Call DateParaMasked(DataFinal, CDate(sParam))
    'DataFinal.PromptInclude = False
    'DataFinal.Text = sParam
    'DataFinal.PromptInclude = True
    
    PreencherParametrosNaTela = SUCESSO
    
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 47599, 47600, 47601, 47602, 47603, 47604, 47605, 47606, 47607
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171150)

    End Select

    Exit Function

End Function

''Function Define_Padrao() As Long
'''Preenche a tela com as opções padrão de FilialEmpresa
''
''Dim iIndice As Integer
''Dim lErro As Long
''
''On Error GoTo Erro_Define_Padrao
''
''    If giFilialEmpresa <> EMPRESA_TODA Then
''
''        FilialEmpresaInicial.ListIndex = -1
''        FilialEmpresaFinal.ListIndex = -1
''
''        FilialEmpresaInicial.Enabled = False
''        FilialEmpresaFinal.Enabled = False
''
''    Else
''
''        FilialEmpresaInicial.ListIndex = -1
''        FilialEmpresaFinal.ListIndex = -1
''
''    End If
''
''    Define_Padrao = SUCESSO
''
''    Exit Function
''
''Erro_Define_Padrao:
''
''    Define_Padrao = Err
''
''    Select Case Err
''
''         Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171151)
''
''    End Select
''
''    Exit Function
''
''End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PICK_LIST
    Set Form_Load_Ocx = Me
    Caption = "Pick-List"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPickList"
    
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
        
        If Me.ActiveControl Is OpInicial Then
            Call LabelOpInicial_Click
        ElseIf Me.ActiveControl Is OpFinal Then
            Call LabelOpFinal_Click
        End If
    
    End If

End Sub


Private Sub LabelOpFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOpFinal, Source, X, Y)
End Sub

Private Sub LabelOpFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOpFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelOpInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOpInicial, Source, X, Y)
End Sub

Private Sub LabelOpInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOpInicial, Button, Shift, X, Y)
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

Private Sub labelAlmoxarifadoFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(labelAlmoxarifadoFinal, Source, X, Y)
End Sub

Private Sub labelAlmoxarifadoFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(labelAlmoxarifadoFinal, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAlmoxarifado, Source, X, Y)
End Sub

Private Sub LabelAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAlmoxarifado, Button, Shift, X, Y)
End Sub

