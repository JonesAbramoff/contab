VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpTitRecOcx 
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   7830
   Begin VB.Frame Frame7 
      Caption         =   "Vendedor"
      Height          =   720
      Left            =   165
      TabIndex        =   37
      Top             =   4890
      Width           =   5355
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label VendedorLabel 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         TabIndex        =   38
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Vencimento"
      Height          =   675
      Left            =   165
      TabIndex        =   32
      Top             =   1260
      Width           =   5355
      Begin MSComCtl2.UpDown UpDownVencimentoDe 
         Height          =   315
         Left            =   2385
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VencimentoDe 
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
      Begin MSComCtl2.UpDown UpDownVencimentoAte 
         Height          =   315
         Left            =   4485
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox VencimentoAte 
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5550
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpTitRecOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpTitRecOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpTitRecOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpTitRecOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Cliente"
      Height          =   1035
      Left            =   165
      TabIndex        =   30
      Top             =   3045
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
      Picture         =   "RelOpTitRecOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   825
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpTitRecOcx.ctx":0A96
      Left            =   1320
      List            =   "RelOpTitRecOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2730
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   720
      Left            =   165
      TabIndex        =   27
      Top             =   4125
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cobrador"
      Height          =   1050
      Left            =   150
      TabIndex        =   26
      Top             =   1935
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
   Begin VB.CheckBox CheckAnalitico 
      Caption         =   "Exibe Título a Título"
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
      Left            =   165
      TabIndex        =   14
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Emissão"
      Height          =   705
      Left            =   150
      TabIndex        =   21
      Top             =   555
      Width           =   5355
      Begin MSComCtl2.UpDown UpDownEmissaoDe 
         Height          =   315
         Left            =   2385
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   25
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
         TabIndex        =   24
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
      TabIndex        =   31
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpTitRecOcx"
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
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 47732
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 47736
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 47736
        
        Case 47732
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173497)

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
    If lErro <> SUCESSO Then Error 47733
    
    ComboOpcoes.Text = ""
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 47785
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47733
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173498)

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
        If lErro <> SUCESSO Then Error 47734

    End If
    
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True


    Select Case Err

        Case 47734
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173499)

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
        If lErro <> SUCESSO Then Error 47735

    End If
        
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True


    Select Case Err

        Case 47735
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173500)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoClienteInic = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    
    'Preenche com os Tipos de Clientes
    lErro = PreencheComboTipo()
    If lErro <> SUCESSO Then Error 47737
    
    'Preenche com os Cobradores
    lErro = PreencheComboCobrador()
    If lErro <> SUCESSO Then Error 47738
     
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then Error 47739
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 47737, 47338, 47739
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173501)

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
    If lErro <> SUCESSO Then Error 57338

    'preenche a ComboBox Cobrador com os objetos da colecao colCobrador
    For Each objCobrador In colCobrador
    
        'Verifica se o cobrador é inativo
        If objCobrador.iInativo <> Inativo Then
        
            ComboCobrador.AddItem CStr(objCobrador.iCodigo) & SEPARADOR & objCobrador.sNomeReduzido
            ComboCobrador.ItemData(ComboCobrador.NewIndex) = objCobrador.iCodigo

        End If

    Next

    PreencheComboCobrador = SUCESSO
    
    Exit Function
    
Erro_PreencheComboCobrador:

    PreencheComboCobrador = Err

    Select Case Err

    Case 47741, 57338
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173502)

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
    If lErro <> SUCESSO Then Error 47742
    
   'preenche a ListBox ComboTipo com os objetos da colecao
    For Each objCodigoDescricaoCliente In colCodigoDescricaoCliente
        ComboTipo.AddItem objCodigoDescricaoCliente.iCodigo & SEPARADOR & objCodigoDescricaoCliente.sNome
        ComboTipo.ItemData(ComboTipo.NewIndex) = objCodigoDescricaoCliente.iCodigo
    Next
        
    PreencheComboTipo = SUCESSO

    Exit Function
    
Erro_PreencheComboTipo:

    PreencheComboTipo = Err

    Select Case Err

    Case 47742
    
    Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173503)

    End Select

    Exit Function

End Function

''Private Sub DataRef_GotFocus()
''
''    Call MaskEdBox_TrataGotFocus(DataRef)
''
''End Sub

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

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173504)

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

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173505)

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

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173506)

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

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173507)

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

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173508)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 47743

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47744

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 47745
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47746
    
    Call BotaoLimpar_Click
               
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 47743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 47744, 47745, 47746
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173509)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 47748

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 47749

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 47748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 47749

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173510)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 47752
    
    If CheckAnalitico.Value = vbChecked Then
        gobjRelatorio.sNomeTsk = "titrec"
    Else
        gobjRelatorio.sNomeTsk = "titrec2"
    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 47752

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173511)

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
Dim iVendedor As Integer

On Error GoTo Erro_PreencherRelOp
            
''    'data de Referência não pode ser vazia
''    If Len(DataRef.ClipText) = 0 Then Error 59631

    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sCliente_I, sCliente_F, sCheckTipo, sClienteTipo, sCheckCobrador, sCobrador, iVendedor)
    If lErro <> SUCESSO Then gError 47757

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 47758
         
    'Preenche o Cliente Inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError 47759
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54772
    
    'Preenche o Cliente Final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError 47760
     
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 54773
                   
    'Preenche o tipo do Cliente
    lErro = objRelOpcoes.IncluirParametro("TTIPOCLIENTE", sClienteTipo)
    If lErro <> AD_BOOL_TRUE Then gError 47761
    
    'Preenche com a Opcao Tipocliente(TodosClientes ou um Cliente)
    lErro = objRelOpcoes.IncluirParametro("TOPTIPO", sCheckTipo)
    If lErro <> AD_BOOL_TRUE Then gError 47762
           
    'Preenche o Cobrador
    lErro = objRelOpcoes.IncluirParametro("TCOBRADOR", sCobrador)
    If lErro <> AD_BOOL_TRUE Then gError 47763
    
    'Preenche a Opcao do Cobrador (todos ou um cobrador)
    lErro = objRelOpcoes.IncluirParametro("TOPCOBRADOR", sCheckCobrador)
    If lErro <> AD_BOOL_TRUE Then gError 47764
       
    'Preenche com o Exibir Titulo a Titulo
    lErro = objRelOpcoes.IncluirParametro("NEXIBTIT", CStr(CheckAnalitico.Value))
    If lErro <> AD_BOOL_TRUE Then gError 47765

    If EmissaoDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMINIC", EmissaoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 47783

    If EmissaoAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMFIM", EmissaoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 47784
    
    If VencimentoDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DVENINIC", VencimentoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVENINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 76238

    If VencimentoAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DVENFIM", VencimentoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DVENFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 76239
    
''    lErro = objRelOpcoes.IncluirParametro("DREF", DataRef.Text)
''    If lErro <> AD_BOOL_TRUE Then Error 47787

    'Preenche com o Exibir Titulo a Titulo
    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", CStr(iVendedor))
    If lErro <> AD_BOOL_TRUE Then gError 47765

    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCliente_I, sCliente_F, sClienteTipo, sCheckTipo, sCobrador, sCheckCobrador)
    If lErro <> SUCESSO Then gError 47766

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

''        Case 59631
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err, Error$)
''            DataRef.SetFocus

        Case 47757, 47758, 47759, 47760, 47761, 47762, 47763, 47764, 47765, 47766
        
        Case 47783, 47784, 47787, 54772, 54773, 76238, 76239
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173512)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_I As String, sCliente_F As String, sCheckTipo As String, sClienteTipo As String, sCheckCobrador As String, sCobrador As String, iVendedor As Integer) As Long
'Verifica se os parâmetros iniciais são maiores que os finais
'E critica o Tipocliente e Cobrador

Dim lErro As Long
Dim objVendedor As New ClassVendedor

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
        
        If CLng(sCliente_I) > CLng(sCliente_F) Then gError 47767
        
    End If
            
    'Se a opção para todos os Clientes estiver selecionada
    If OptionTodosTipos.Value = True Then
        sCheckTipo = "Todos"
        sClienteTipo = ""
    
    'Se a opção para apenas um Cliente estiver selecionada
    Else
        'TEm que indicar o tipo do Cliente
        If ComboTipo.Text = "" Then gError 47768
        sCheckTipo = "Um"
        sClienteTipo = ComboTipo.Text
    
    End If
    
    If OptionTodosCobradores.Value = True Then
        sCheckCobrador = "Todos"
        sCobrador = ""
    
    'Se a opção para apenas um Cliente estiver selecionada
    Else
        'tem que indicar o Cobrador
        If ComboCobrador.Text = "" Then gError 47769
        sCheckCobrador = "Um"
        sCobrador = ComboCobrador.Text
    
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(EmissaoDe.ClipText) <> "" And Trim(EmissaoAte.ClipText) <> "" Then
    
         If CDate(EmissaoDe.Text) > CDate(EmissaoAte.Text) Then gError 47780
    
    End If
    
    'data vencimento inicial nao pode ser maior que a final
    If Trim(VencimentoDe.ClipText) <> "" And Trim(VencimentoAte.ClipText) <> "" Then
    
        If CDate(VencimentoDe.Text) > CDate(VencimentoAte.Text) Then gError 76237
        
    End If
    
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Verifica se vendedor existe
    If objVendedor.sNomeReduzido <> "" Then
    
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError ERRO_SEM_MENSAGEM

        iVendedor = objVendedor.iCodigo

    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 47767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteInicial.SetFocus
                
        Case 47768
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_PREENCHIDO", gErr)
            ComboTipo.SetFocus
            
        Case 47769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", gErr)
            ComboCobrador.SetFocus
        
        Case 47780
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_INICIAL_MAIOR", gErr)
            EmissaoDe.SetFocus
            
        Case 76237
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VENCTO_INICIAL_MAIOR", gErr)
            VencimentoDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173513)

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
    
    If Trim(VencimentoDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vencimento >= " & Forprint_ConvData(CDate(VencimentoDe.Text))

    End If
    
    If Trim(VencimentoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Vencimento <= " & Forprint_ConvData(CDate(VencimentoAte.Text))

    End If
    
''???    If sExpressao <> "" Then sExpressao = sExpressao & " E "
''???    sExpressao = sExpressao & "ExibeTitulo = " & Forprint_ConvInt(CheckAnalitico.Value)
    
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

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173514)

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
    If lErro <> SUCESSO Then gError 47770
   
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 47771
    
    ClienteInicial.Text = sParam
    Call ClienteInicial_Validate(bSGECancelDummy)
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 47772
    
    ClienteFinal.Text = sParam
    Call ClienteFinal_Validate(bSGECancelDummy)
                
    'pega  Tipo cliente e Exibe
    lErro = objRelOpcoes.ObterParametro("TOPTIPO", sParam)
    If lErro <> SUCESSO Then gError 47773
                   
    If sParam = "Todos" Then
    
        Call OptionTodosTipos_Click
    
    Else
        'se é "um tipo só" então exibe o tipo
        lErro = objRelOpcoes.ObterParametro("TTIPOCLIENTE", sTipoCliente)
        If lErro <> SUCESSO Then gError 47774
                            
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
    If lErro <> SUCESSO Then gError 47775
                   
    If sParam = "Todos" Then
    
        Call OptionTodosCobradores_Click
    
    Else
    
        'se existe um só cobrador entao exibe
        lErro = objRelOpcoes.ObterParametro("TCOBRADOR", sCobrador)
        If lErro <> SUCESSO Then gError 47776
                            
        OptionApenasCobrador.Value = True
        ComboCobrador.Enabled = True
        
        If sCobrador = "" Then
            ComboCobrador.ListIndex = -1
        Else
'Catharine
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
'Catharine
          ' Apagar
          '  ComboCobrador.Text = sCobrador
        End If
        
    End If
    
    lErro = objRelOpcoes.ObterParametro("NEXIBTIT", sParam)
    If lErro <> SUCESSO Then gError 47777
    
    CheckAnalitico.Value = CInt(sParam)
   
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DEMINIC", sParam)
    If lErro <> SUCESSO Then gError 47781

    Call DateParaMasked(EmissaoDe, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DEMFIM", sParam)
    If lErro <> SUCESSO Then gError 47782

    Call DateParaMasked(EmissaoAte, CDate(sParam))
       
    'pega data vencimento inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DVENINIC", sParam)
    If lErro <> SUCESSO Then gError 76240

    Call DateParaMasked(VencimentoDe, CDate(sParam))
    
    'pega data vencimento final e exibe
    lErro = objRelOpcoes.ObterParametro("DVENFIM", sParam)
    If lErro <> SUCESSO Then gError 76241

    Call DateParaMasked(VencimentoAte, CDate(sParam))
       
''    'pega data final e exibe
''    lErro = objRelOpcoes.ObterParametro("DREF", sParam)
''    If lErro <> SUCESSO Then Error 47788
''
''    Call DateParaMasked(DataRef, CDate(sParam))

    lErro = objRelOpcoes.ObterParametro("NVENDEDOR", sParam)
    If lErro <> SUCESSO Then gError 47773
    
    If StrParaInt(sParam) <> 0 Then
        Vendedor.Text = sParam
        Call Vendedor_Validate(bSGECancelDummy)
    End If
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 47770, 47771, 47772, 47773, 47774, 47775, 47776
        
        Case 47777, 47781, 47782, 47788, 76240, 76241
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173515)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
''    'Define Data de Referencia como data atual
''    DataRef.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'defina todos os tipos
    Call OptionTodosTipos_Click
    
    'define todos os cobradores
    Call OptionTodosCobradores_Click
    
    'define Exibir Titulo a Titulo como Padrao
    CheckAnalitico.Value = 1
    
    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173516)
    
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

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173517)

    End Select

    Exit Sub
    
End Sub

Private Sub EmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoAte_Validate

    If Len(EmissaoAte.ClipText) > 0 Then
        
        lErro = Data_Critica(EmissaoAte.Text)
        If lErro <> SUCESSO Then Error 47789

    End If

    Exit Sub

Erro_EmissaoAte_Validate:

    Cancel = True


    Select Case Err

        Case 47789

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173518)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoDe_Validate

    If Len(EmissaoDe.ClipText) > 0 Then

        lErro = Data_Critica(EmissaoDe.Text)
        If lErro <> SUCESSO Then Error 47790

    End If

    Exit Sub

Erro_EmissaoDe_Validate:

    Cancel = True


    Select Case Err

        Case 47790

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173519)

    End Select

    Exit Sub

End Sub

''Private Sub DataRef_Validate(Cancel As Boolean)
''
''Dim lErro As Long
''
''On Error GoTo Erro_DataRef_Validate
''
''    If Len(DataRef.ClipText) > 0 Then
''
''        lErro = Data_Critica(DataRef.Text)
''        If lErro <> SUCESSO Then Error 47791
''
''    End If
''
''    Exit Sub
''
''Erro_DataRef_Validate:
''
''    Cancel = True
''
''
''    Select Case Err
''
''        Case 47791
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173520)
''
''    End Select
''
''    Exit Sub
''
''End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing
    Set objEventoVendedor = Nothing
    
End Sub


''Private Sub UpDownDataRef_DownClick()
''
''Dim lErro As Long
''
''On Error GoTo Erro_UpDownDataRef_DownClick
''
''    lErro = Data_Up_Down_Click(DataRef, DIMINUI_DATA)
''    If lErro <> SUCESSO Then Error 47850
''
''    Exit Sub
''
''Erro_UpDownDataRef_DownClick:
''
''    Select Case Err
''
''        Case 47850
''            DataRef.SetFocus
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173521)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
''Private Sub UpDownDataRef_UpClick()
''
''Dim lErro As Long
''
''On Error GoTo Erro_UpDownDataRef_UpClick
''
''    lErro = Data_Up_Down_Click(DataRef, AUMENTA_DATA)
''    If lErro <> SUCESSO Then Error 47851
''
''    Exit Sub
''
''Erro_UpDownDataRef_UpClick:
''
''    Select Case Err
''
''        Case 47851
''            DataRef.SetFocus
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173522)
''
''    End Select
''
''    Exit Sub
''
''End Sub
    
Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    lErro = Data_Up_Down_Click(EmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47852

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case Err

        Case 47852
            EmissaoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173523)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    lErro = Data_Up_Down_Click(EmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47853

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case Err

        Case 47853
            EmissaoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173524)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    lErro = Data_Up_Down_Click(EmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 47854

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case Err

        Case 47854
            EmissaoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173525)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    lErro = Data_Up_Down_Click(EmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 47855

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case Err

        Case 47855
            EmissaoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173526)

    End Select

    Exit Sub

End Sub
Private Sub UpDownVencimentoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoDe_DownClick

    lErro = Data_Up_Down_Click(VencimentoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 76231

    Exit Sub

Erro_UpDownVencimentoDe_DownClick:

    Select Case gErr

        Case 76231
            VencimentoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173527)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimentoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoDe_UpClick

    lErro = Data_Up_Down_Click(VencimentoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 76232

    Exit Sub

Erro_UpDownVencimentoDe_UpClick:

    Select Case gErr

        Case 76232
            VencimentoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173528)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownVencimentoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoAte_DownClick

    lErro = Data_Up_Down_Click(VencimentoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 76233

    Exit Sub

Erro_UpDownVencimentoAte_DownClick:

    Select Case gErr

        Case 76233
            VencimentoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173529)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimentoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoAte_UpClick

    lErro = Data_Up_Down_Click(VencimentoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 76234

    Exit Sub

Erro_UpDownVencimentoAte_UpClick:

    Select Case gErr

        Case 76234
            VencimentoAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173530)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TIT_REC
    Set Form_Load_Ocx = Me
    Caption = "Títulos a Receber"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpTitRec"
    
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

Private Sub VencimentoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(VencimentoAte)
    
End Sub

Private Sub VencimentoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VencimentoAte_Validate

    If Len(VencimentoAte.ClipText) > 0 Then

        lErro = Data_Critica(VencimentoAte.Text)
        If lErro <> SUCESSO Then gError 76236

    End If

    Exit Sub

Erro_VencimentoAte_Validate:

    Cancel = True


    Select Case gErr

        Case 76236

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173531)

    End Select

    Exit Sub

End Sub


Private Sub VencimentoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(VencimentoDe)
    
End Sub

Private Sub VencimentoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VencimentoDe_Validate

    If Len(VencimentoDe.ClipText) > 0 Then

        lErro = Data_Critica(VencimentoDe.Text)
        If lErro <> SUCESSO Then gError 76235

    End If

    Exit Sub

Erro_VencimentoDe_Validate:

    Cancel = True


    Select Case gErr

        Case 76235

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173532)

    End Select

    Exit Sub

End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim dPercComissao As Double

On Error GoTo Erro_Vendedor_Validate

    'Se Vendedor foi alterado,
    If Len(Trim(Vendedor.Text)) <> 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le(Vendedor, objVendedor)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209207)
    
    End Select

End Sub

Public Sub VendedorLabel_Click()

'BROWSE VENDEDOR :

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor.Text = objVendedor.sNomeReduzido

    Me.Show

    Vendedor.SetFocus 'Inserido por Wagner
    
    Exit Sub

End Sub
