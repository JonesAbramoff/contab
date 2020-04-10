VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpParcRecDifOcx 
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   KeyPreview      =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   7830
   Begin VB.Frame Frame5 
      Caption         =   "Registro da Diferença na Parcela"
      Height          =   675
      Left            =   165
      TabIndex        =   30
      Top             =   1335
      Width           =   5355
      Begin MSComCtl2.UpDown UpDownRegistroDe 
         Height          =   315
         Left            =   2385
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox RegistroDe 
         Height          =   285
         Left            =   1230
         TabIndex        =   3
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownRegistroAte 
         Height          =   315
         Left            =   4485
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox RegistroAte 
         Height          =   285
         Left            =   3330
         TabIndex        =   4
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
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
         Left            =   2940
         TabIndex        =   34
         Top             =   315
         Width           =   360
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
         Left            =   870
         TabIndex        =   33
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5550
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpParcRecDifOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpParcRecDifOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpParcRecDifOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpParcRecDifOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Cliente"
      Height          =   1035
      Left            =   165
      TabIndex        =   28
      Top             =   3120
      Width           =   5355
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
         TabIndex        =   8
         Top             =   315
         Width           =   960
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
         TabIndex        =   9
         Top             =   630
         Width           =   1755
      End
      Begin VB.ComboBox ComboTipo 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   3225
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
      Left            =   5670
      Picture         =   "RelOpParcRecDifOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   825
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpParcRecDifOcx.ctx":0A96
      Left            =   1320
      List            =   "RelOpParcRecDifOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2730
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   720
      Left            =   165
      TabIndex        =   25
      Top             =   4200
      Width           =   5355
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   570
         TabIndex        =   11
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
         TabIndex        =   12
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2805
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cobrador"
      Height          =   1050
      Left            =   150
      TabIndex        =   24
      Top             =   2010
      Width           =   5355
      Begin VB.ComboBox ComboCobrador 
         Height          =   315
         Left            =   2265
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   2925
      End
      Begin VB.OptionButton OptionApenasCobrador 
         Caption         =   "Apenas do Cobrador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   165
         TabIndex        =   6
         Top             =   630
         Width           =   2070
      End
      Begin VB.OptionButton OptionTodosCobradores 
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
         TabIndex        =   5
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Emissão do Título"
      Height          =   705
      Left            =   150
      TabIndex        =   19
      Top             =   630
      Width           =   5355
      Begin MSComCtl2.UpDown UpDownEmissaoDe 
         Height          =   315
         Left            =   2385
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoDe 
         Height          =   285
         Left            =   1230
         TabIndex        =   1
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissaoAte 
         Height          =   315
         Left            =   4485
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoAte 
         Height          =   285
         Left            =   3330
         TabIndex        =   2
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
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
         Height          =   195
         Left            =   870
         TabIndex        =   23
         Top             =   315
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
         Left            =   2940
         TabIndex        =   22
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
      Left            =   615
      TabIndex        =   29
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpParcRecDifOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoClienteInic As AdmEvento
Attribute objEventoClienteInic.VB_VarHelpID = -1
Private WithEvents objEventoClienteFim As AdmEvento
Attribute objEventoClienteFim.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 177977
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 177978
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 177977
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
            
        Case 177978
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177979)

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
    If lErro <> SUCESSO Then gError 177980
    
    ComboOpcoes.Text = ""
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 177981
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 177980 To 177981
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177982)

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
        lErro = TP_cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 177983

    End If
    
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 177983
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177984)

    End Select

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate
    
    'se está Preenchido
    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 177985

    End If
        
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 177985
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177986)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoClienteInic = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
    
    'Preenche com os Tipos de Clientes
    lErro = PreencheComboTipo()
    If lErro <> SUCESSO Then gError 177987
    
    'Preenche com os Cobradores
    lErro = PreencheComboCobrador()
    If lErro <> SUCESSO Then gError 177988
     
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 177989
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 177987 To 177989
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177990)

    End Select

    Exit Sub

End Sub

Function PreencheComboCobrador() As Long

Dim lErro As Long
Dim colCobrador As New Collection
Dim objCobrador As New ClassCobrador

On Error GoTo Erro_PreencheComboCobrador

    'Le cada codigo e nome da tabela Cobradores
    lErro = CF("Cobradores_Le_Todos_Filial", colCobrador)
    If lErro <> SUCESSO Then gError 177991

    'preenche a ComboBox Cobrador com os objetos da colecao colCobrador
    For Each objCobrador In colCobrador
    
        'Verifica se o cobrador é inativo
        If objCobrador.iInativo <> INATIVO Then
        
            ComboCobrador.AddItem CStr(objCobrador.iCodigo) & SEPARADOR & objCobrador.sNomeReduzido
            ComboCobrador.ItemData(ComboCobrador.NewIndex) = objCobrador.iCodigo

        End If

    Next

    PreencheComboCobrador = SUCESSO
    
    Exit Function
    
Erro_PreencheComboCobrador:

    PreencheComboCobrador = gErr

    Select Case gErr

    Case 177991
    
    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177992)

    End Select

    Exit Function

End Function

Function PreencheComboTipo() As Long

Dim lErro As Long
Dim colCodigoDescricaoCliente As New AdmColCodigoNome
Dim objCodigoDescricaoCliente As New AdmCodigoNome

On Error GoTo Erro_PreencheComboTipo
    
    'Preenche a Colecao com os Tipos de clientes
    lErro = CF("Cod_Nomes_Le", "TiposdeCliente", "Codigo", "Descricao", STRING_TIPO_CLIENTE_DESCRICAO, colCodigoDescricaoCliente)
    If lErro <> SUCESSO Then gError 177993
    
   'preenche a ListBox ComboTipo com os objetos da colecao
    For Each objCodigoDescricaoCliente In colCodigoDescricaoCliente
        ComboTipo.AddItem objCodigoDescricaoCliente.iCodigo & SEPARADOR & objCodigoDescricaoCliente.sNome
        ComboTipo.ItemData(ComboTipo.NewIndex) = objCodigoDescricaoCliente.iCodigo
    Next
        
    PreencheComboTipo = SUCESSO

    Exit Function
    
Erro_PreencheComboTipo:

    PreencheComboTipo = gErr

    Select Case gErr

    Case 177993
    
    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177994)

    End Select

    Exit Function

End Function

Private Sub EmissaoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(EmissaoAte)

End Sub

Private Sub EmissaoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(EmissaoDe)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177995)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177996)

    End Select

    Exit Sub
    
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

Private Sub OptionApenasCobrador_Click()

Dim lErro As Long

On Error GoTo Erro_OptionApenasCobrador_Click
    
    'Limpa a Combo Cobrador e abilita
    ComboCobrador.ListIndex = -1
    ComboCobrador.Enabled = True
    ComboCobrador.SetFocus
    
    Exit Sub

Erro_OptionApenasCobrador_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177997)

    End Select

    Exit Sub
    
End Sub

Private Sub OptionTodosCobradores_Click()

Dim lErro As Long

On Error GoTo Erro_OptionTodosCobradores_Click
    
    'Limpa e Desabilita a ComboCobrador
    ComboCobrador.ListIndex = -1
    ComboCobrador.Enabled = False
    OptionTodosCobradores.Value = True
    
    Exit Sub

Erro_OptionTodosCobradores_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177998)

    End Select

End Sub

Private Sub OptionTodosTipos_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_OptionTodosTipos_Click
    
    'Limpa e desabilita a ComboTipo
    ComboTipo.ListIndex = -1
    ComboTipo.Enabled = False
    OptionTodosTipos.Value = True
    
    Exit Sub

Erro_OptionTodosTipos_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177999)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 180000

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 180001

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 180002
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 180003
    
    Call BotaoLimpar_Click
               
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 180000
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 180001 To 180003
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180004)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 180005

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 180006

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 180005
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 180006

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180007)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 180008

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 180008

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180009)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCliente_I As String
Dim sCliente_F As String
Dim sCheckTipo As String
Dim sClienteTipo As String
Dim sCheckCobrador As String
Dim sCobrador As String

On Error GoTo Erro_PreencherRelOp

    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sCliente_I, sCliente_F, sCheckTipo, sClienteTipo, sCheckCobrador, sCobrador)
    If lErro <> SUCESSO Then gError 180010

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 180011
         
    'Preenche o Cliente Inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError 180012
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 180013
    
    'Preenche o Cliente Final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError 180014
     
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 180015
                   
    'Preenche o tipo do Cliente
    lErro = objRelOpcoes.IncluirParametro("TTIPOCLIENTE", sClienteTipo)
    If lErro <> AD_BOOL_TRUE Then gError 180016
    
    'Preenche com a Opcao Tipocliente(TodosClientes ou um Cliente)
    lErro = objRelOpcoes.IncluirParametro("TOPTIPO", sCheckTipo)
    If lErro <> AD_BOOL_TRUE Then gError 180017
           
    'Preenche o Cobrador
    lErro = objRelOpcoes.IncluirParametro("TCOBRADOR", sCobrador)
    If lErro <> AD_BOOL_TRUE Then gError 180018
    
    'Preenche a Opcao do Cobrador (todos ou um cobrador)
    lErro = objRelOpcoes.IncluirParametro("TOPCOBRADOR", sCheckCobrador)
    If lErro <> AD_BOOL_TRUE Then gError 180019

    If EmissaoDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMINIC", EmissaoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 180020

    If EmissaoAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMFIM", EmissaoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 180021
    
    If RegistroDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DREGINIC", RegistroDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DREGINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 180022

    If RegistroAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DREGFIM", RegistroAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DREGFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 180023

    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCliente_I, sCliente_F, sClienteTipo, sCheckTipo, sCobrador, sCheckCobrador)
    If lErro <> SUCESSO Then gError 180024

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 180010 To 180024
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180025)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_I As String, sCliente_F As String, sCheckTipo As String, sClienteTipo As String, sCheckCobrador As String, sCobrador As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais
'E critica o Tipocliente e Cobrador

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
       
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
            
    If sCliente_I <> "" And sCliente_F <> "" Then
        
        If CLng(sCliente_I) > CLng(sCliente_F) Then gError 180026
        
    End If
            
    'Se a opção para todos os Clientes estiver selecionada
    If OptionTodosTipos.Value = True Then
        sCheckTipo = "Todos"
        sClienteTipo = ""
    
    'Se a opção para apenas um Cliente estiver selecionada
    Else
        'TEm que indicar o tipo do Cliente
        If ComboTipo.Text = "" Then gError 180027
        sCheckTipo = "Um"
        sClienteTipo = ComboTipo.Text
    
    End If
    
    If OptionTodosCobradores.Value = True Then
        sCheckCobrador = "Todos"
        sCobrador = ""
    
    'Se a opção para apenas um Cliente estiver selecionada
    Else
        'tem que indicar o Cobrador
        If ComboCobrador.Text = "" Then gError 180028
        sCheckCobrador = "Um"
        sCobrador = ComboCobrador.Text
    
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(EmissaoDe.ClipText) <> "" And Trim(EmissaoAte.ClipText) <> "" Then
    
         If CDate(EmissaoDe.Text) > CDate(EmissaoAte.Text) Then gError 180029
    
    End If
    
    'data Registro inicial nao pode ser maior que a final
    If Trim(RegistroDe.ClipText) <> "" And Trim(RegistroAte.ClipText) <> "" Then
    
        If CDate(RegistroDe.Text) > CDate(RegistroAte.Text) Then gError 180030
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 180026
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
                
        Case 180027
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_PREENCHIDO", gErr)
            ComboTipo.SetFocus
            
        Case 180028
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)
            ComboCobrador.SetFocus
        
        Case 180029
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_INICIAL_MAIOR", gErr)
            EmissaoDe.SetFocus
            
        Case 180030
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_VENCTO_INICIAL_MAIOR", gErr)
            RegistroDe.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180031)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCliente_I As String, sCliente_F As String, sClienteTipo As String, sCheckTipo As String, sCobrador As String, sCheckCobrador As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sCliente_I <> "" Then sExpressao = "Cliente >= " & Forprint_ConvLong(CLng(sCliente_I))

   If sCliente_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sCliente_F))

    End If
           
    'Se a opção para apenas um cliente estiver selecionada
    If sCheckTipo = "Um" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoCliente = " & Forprint_ConvInt(CInt(Codigo_Extrai(sClienteTipo)))

    End If
    
    'Se a opção para apenas um cobrador estiver selecionada
    If sCheckCobrador = "Um" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cobrador = " & Forprint_ConvInt(CInt(Codigo_Extrai(sCobrador)))

    End If
    
    If Trim(EmissaoDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao >= " & Forprint_ConvData(CDate(EmissaoDe.Text))

    End If
    
    If Trim(EmissaoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao <= " & Forprint_ConvData(CDate(EmissaoAte.Text))

    End If
    
    If Trim(RegistroDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Registro >= " & Forprint_ConvData(CDate(RegistroDe.Text))

    End If
    
    If Trim(RegistroAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Registro <= " & Forprint_ConvData(CDate(RegistroAte.Text))

    End If
    
    If giFilialEmpresa <> EMPRESA_TODA And giFilialEmpresa <> gobjCR.iFilialCentralizadora Then
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180031)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoCliente As String
Dim sCobrador As String
'Catharine
Dim iCobrador As Integer
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 180032
   
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 180033
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 180034
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
                
    'pega  Tipo cliente e Exibe
    lErro = objRelOpcoes.ObterParametro("TOPTIPO", sParam)
    If lErro <> SUCESSO Then gError 180035
                   
    If sParam = "Todos" Then
    
        Call OptionTodosTipos_Click
    
    Else
        'se é "um tipo só" então exibe o tipo
        lErro = objRelOpcoes.ObterParametro("TTIPOCLIENTE", sTipoCliente)
        If lErro <> SUCESSO Then gError 180036
                            
        OptionUmTipo.Value = True
        ComboTipo.Enabled = True
        
        If sTipoCliente = "" Then
            ComboTipo.ListIndex = -1
        Else
            ComboTipo.Text = sTipoCliente
        End If
    End If
           
    'Pega o TipoCobrador e Exibe
    lErro = objRelOpcoes.ObterParametro("TOPCOBRADOR", sParam)
    If lErro <> SUCESSO Then gError 180037
                   
    If sParam = "Todos" Then
    
        Call OptionTodosCobradores_Click
    
    Else
    
        'se existe um só cobrador entao exibe
        lErro = objRelOpcoes.ObterParametro("TCOBRADOR", sCobrador)
        If lErro <> SUCESSO Then gError 180038
                            
        OptionApenasCobrador.Value = True
        ComboCobrador.Enabled = True
        
        If sCobrador = "" Then
            ComboCobrador.ListIndex = -1
        Else
            
            'Obtem o código do cobrador
            iCobrador = CInt(Codigo_Extrai(sCobrador))
            
            For iIndice = 0 To ComboCobrador.ListCount - 1
                'Verifica se existe o código obtido do Cobrador na combo de Cobrador
                If ComboCobrador.ItemData(iIndice) = iCobrador Then
                    
                    ComboCobrador.ListIndex = -1
                    'Preenche a combo do Cobrador com código e nome do cobrador
                    ComboCobrador.Text = sCobrador
                    
                End If
                
            Next
        
        End If
        
    End If
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DEMINIC", sParam)
    If lErro <> SUCESSO Then gError 180040

    Call DateParaMasked(EmissaoDe, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DEMFIM", sParam)
    If lErro <> SUCESSO Then gError 180041

    Call DateParaMasked(EmissaoAte, CDate(sParam))
       
    'pega data Registro inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DREGINIC", sParam)
    If lErro <> SUCESSO Then gError 180042

    Call DateParaMasked(RegistroDe, CDate(sParam))
    
    'pega data Registro final e exibe
    lErro = objRelOpcoes.ObterParametro("DREGFIM", sParam)
    If lErro <> SUCESSO Then gError 180043

    Call DateParaMasked(RegistroAte, CDate(sParam))
              
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 180032 To 180043
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180044)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    'defina todos os tipos
    Call OptionTodosTipos_Click
    
    'define todos os cobradores
    Call OptionTodosCobradores_Click
    
    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180045)
    
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
    
    'Limpa Combo Tipo e Abilita
    ComboTipo.ListIndex = -1
    ComboTipo.Enabled = True
    ComboTipo.SetFocus
    
    Exit Sub

Erro_OptionUmTipo_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180046)

    End Select

    Exit Sub
    
End Sub

Private Sub EmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoAte_Validate

    If Len(EmissaoAte.ClipText) > 0 Then
        
        lErro = Data_Critica(EmissaoAte.Text)
        If lErro <> SUCESSO Then gError 180047

    End If

    Exit Sub

Erro_EmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 180047

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180048)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoDe_Validate

    If Len(EmissaoDe.ClipText) > 0 Then

        lErro = Data_Critica(EmissaoDe.Text)
        If lErro <> SUCESSO Then gError 180049

    End If

    Exit Sub

Erro_EmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 180049

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180050)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing
    
End Sub
    
Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    lErro = Data_Up_Down_Click(EmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 180051

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 180051
            EmissaoDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180052)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    lErro = Data_Up_Down_Click(EmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 180053

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 180053
            EmissaoDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180054)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    lErro = Data_Up_Down_Click(EmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 180055

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 180055
            EmissaoAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180056)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    lErro = Data_Up_Down_Click(EmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 180057

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 180057
            EmissaoAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180058)

    End Select

    Exit Sub

End Sub

Private Sub UpDownRegistroDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownRegistroDe_DownClick

    lErro = Data_Up_Down_Click(RegistroDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 180059

    Exit Sub

Erro_UpDownRegistroDe_DownClick:

    Select Case gErr

        Case 180059
            RegistroDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180060)

    End Select

    Exit Sub

End Sub

Private Sub UpDownRegistroDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownRegistroDe_UpClick

    lErro = Data_Up_Down_Click(RegistroDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 180061

    Exit Sub

Erro_UpDownRegistroDe_UpClick:

    Select Case gErr

        Case 180061
            RegistroDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180062)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownRegistroAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownRegistroAte_DownClick

    lErro = Data_Up_Down_Click(RegistroAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 180063

    Exit Sub

Erro_UpDownRegistroAte_DownClick:

    Select Case gErr

        Case 180063
            RegistroAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180064)

    End Select

    Exit Sub

End Sub

Private Sub UpDownRegistroAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownRegistroAte_UpClick

    lErro = Data_Up_Down_Click(RegistroAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 180065

    Exit Sub

Erro_UpDownRegistroAte_UpClick:

    Select Case gErr

        Case 180065
            RegistroAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180066)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TIT_REC
    Set Form_Load_Ocx = Me
    Caption = "Diferenças nas Parcelas a Receber"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpParcRecDif"
    
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub RegistroAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(RegistroAte)
    
End Sub

Private Sub RegistroAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_RegistroAte_Validate

    If Len(RegistroAte.ClipText) > 0 Then

        lErro = Data_Critica(RegistroAte.Text)
        If lErro <> SUCESSO Then gError 180067

    End If

    Exit Sub

Erro_RegistroAte_Validate:

    Cancel = True

    Select Case gErr

        Case 180067

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180068)

    End Select

    Exit Sub

End Sub


Private Sub RegistroDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(RegistroDe)
    
End Sub

Private Sub RegistroDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_RegistroDe_Validate

    If Len(RegistroDe.ClipText) > 0 Then

        lErro = Data_Critica(RegistroDe.Text)
        If lErro <> SUCESSO Then gError 180069

    End If

    Exit Sub

Erro_RegistroDe_Validate:

    Cancel = True

    Select Case gErr

        Case 180069

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180070)

    End Select

    Exit Sub

End Sub
