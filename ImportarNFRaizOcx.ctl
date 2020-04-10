VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ImportarNFRaiz 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame FrameSel 
      Caption         =   "Selecionados:"
      Height          =   930
      Left            =   6600
      TabIndex        =   30
      Top             =   4980
      Width           =   2715
      Begin VB.Label Label5 
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
         Left            =   555
         TabIndex        =   34
         Top             =   615
         Width           =   510
      End
      Begin VB.Label ValorSel 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1185
         TabIndex        =   33
         Top             =   555
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Toneladas:"
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
         TabIndex        =   32
         Top             =   300
         Width           =   960
      End
      Begin VB.Label ToneladasSel 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1185
         TabIndex        =   31
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame FrameTotal 
      Caption         =   "Total:"
      Height          =   930
      Left            =   3720
      TabIndex        =   25
      Top             =   4980
      Width           =   2715
      Begin VB.Label ValorTotal 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1185
         TabIndex        =   29
         Top             =   540
         Width           =   1365
      End
      Begin VB.Label Label6 
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
         Left            =   600
         TabIndex        =   28
         Top             =   615
         Width           =   510
      End
      Begin VB.Label ToneladasTotal 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1185
         TabIndex        =   27
         Top             =   195
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Toneladas:"
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
         TabIndex        =   26
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.CommandButton BotaoDesmarcar 
      Caption         =   "Desmarcar Todas"
      Height          =   555
      Left            =   1830
      Picture         =   "ImportarNFRaizOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5340
      Width           =   1530
   End
   Begin VB.CommandButton BotaoMarcar 
      Caption         =   "Marcar Todas"
      Height          =   555
      Left            =   120
      Picture         =   "ImportarNFRaizOcx.ctx":11E2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5340
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      Caption         =   "Importação dos Dados"
      Height          =   690
      Left            =   120
      TabIndex        =   13
      Top             =   45
      Width           =   7230
      Begin VB.CommandButton BotaoTrazer 
         Caption         =   "Trazer Dados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4980
         TabIndex        =   4
         Top             =   210
         Width           =   1710
      End
      Begin MSComCtl2.UpDown UpDownDataInicial 
         Height          =   300
         Left            =   1890
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   735
         TabIndex        =   0
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataFinal 
         Height          =   300
         Left            =   4350
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   3195
         TabIndex        =   2
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   15
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   2730
         TabIndex        =   14
         Top             =   285
         Width           =   360
      End
   End
   Begin VB.Frame FrameNF 
      Caption         =   "Notas Fiscais"
      Height          =   4170
      Left            =   120
      TabIndex        =   12
      Top             =   765
      Width           =   9210
      Begin VB.CheckBox Selecionado 
         Height          =   255
         Left            =   4605
         TabIndex        =   24
         Top             =   1095
         Width           =   465
      End
      Begin MSMask.MaskEdBox Filial 
         Height          =   300
         Left            =   6330
         TabIndex        =   23
         Top             =   3345
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fornecedor 
         Height          =   300
         Left            =   5235
         TabIndex        =   22
         Top             =   2580
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   4680
         TabIndex        =   21
         Top             =   3210
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorNF 
         Height          =   300
         Left            =   5805
         TabIndex        =   20
         Top             =   1500
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox ValorFrete 
         Height          =   300
         Left            =   5595
         TabIndex        =   19
         Top             =   2205
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox PrecoUnitario 
         Height          =   300
         Left            =   3690
         TabIndex        =   16
         Top             =   1650
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   300
         Left            =   2190
         TabIndex        =   17
         Top             =   1620
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   300
         Left            =   390
         TabIndex        =   18
         Top             =   1605
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   3555
         Left            =   90
         TabIndex        =   5
         Top             =   225
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   6271
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
   End
   Begin VB.PictureBox Picture9 
      Height          =   555
      Left            =   7680
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   90
      Width           =   1680
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ImportarNFRaizOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1110
         Picture         =   "ImportarNFRaizOcx.ctx":2356
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   607
         Picture         =   "ImportarNFRaizOcx.ctx":24D4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
   End
End
Attribute VB_Name = "ImportarNFRaiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim gcolNF As Collection

Dim objGridItens As AdmGrid
Dim iGrid_Selecionado_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_DataEmissao_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_PrecoUnitario_Col As Integer
'Dim iGrid_ValorProdutos_Col As Integer
Dim iGrid_ValorFrete_Col As Integer
Dim iGrid_ValorNF_Col As Integer

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_PROCESSA_ARQRETCOBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Importar Notas Fiscais de Raiz"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "ImportarNFRaiz"
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

Public Sub Form_Load()

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_Form_Load

    Set objGridItens = New AdmGrid
    Set gcolNF = New Collection

    Call Inicializa_Grid_Itens(objGridItens)
    
    DoEvents
    
    lErro = CF("ImportacaoNFRaiz_Le_Data", dtData)
    If lErro <> SUCESSO Then gError 181967
    
    If dtData <> DATA_NULA Then
    
        DataInicial.PromptInclude = False
        DataInicial.Text = Format(dtData, "dd/mm/yy")
        DataInicial.PromptInclude = True
    
        DataFinal.PromptInclude = False
        DataFinal.Text = Format(DateAdd("d", 1, dtData), "dd/mm/yy")
        DataFinal.PromptInclude = True
        
        Call BotaoTrazer_Click
    
    End If
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub
    
Public Sub Form_Activate()

    'Carrega os índices da tela
    'Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    'gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridItens = Nothing
    Set gcolNF = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long
    Trata_Parametros = SUCESSO
End Function

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Private Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Private Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Private Sub GridItens_LostFocus()
    Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer

    'Guarda iLinhasExistentes
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes

    'Verifica se a Tecla apertada foi Del
    If KeyCode = vbKeyDelete Then
        'Guarda o índice da Linha a ser Excluída
        iLinhaAnterior = GridItens.Row
    End If

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Itens
    
    'Tela em questão
    Set objGridInt.objForm = Me
    
    'Títulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Data Emissão")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Toneladas")
    objGridInt.colColuna.Add ("Valor Tonelada")
    'objGridInt.colColuna.Add ("Valor Produto")
    objGridInt.colColuna.Add ("Valor Frete")
    objGridInt.colColuna.Add ("Valor Nota")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Selecionado.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    'objGridInt.colCampo.Add (ValorProdutos.Name)
    objGridInt.colCampo.Add (ValorFrete.Name)
    objGridInt.colCampo.Add (ValorNF.Name)
    
    iGrid_Selecionado_Col = 1
    iGrid_Fornecedor_Col = 2
    iGrid_Filial_Col = 3
    iGrid_DataEmissao_Col = 4
    iGrid_Produto_Col = 5
    iGrid_Quantidade_Col = 6
    iGrid_PrecoUnitario_Col = 7
    'iGrid_ValorProdutos_Col = 8
    iGrid_ValorFrete_Col = 8
    iGrid_ValorNF_Col = 9
        
    objGridInt.objGrid = GridItens
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10
        
    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1
    
    GridItens.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Itens:

    Inicializa_Grid_Itens = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Function
        
End Function

Private Sub Selecionado_Click()

    iAlterado = REGISTRO_ALTERADO
    
    Call Calcula_Totais
        
End Sub

Private Sub Selecionado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)
        
End Sub

Private Sub Selecionado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
        
End Sub

Private Sub Selecionado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Selecionado
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Inicializa variáveis para saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
        
        'Finaliza variáveis para saída de célula
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 181910

    End If
       
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr
                
        Case 181910
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Function
    
End Function

Private Sub BotaoTrazer_Click()

Dim lErro As Long
Dim colRegTick As New Collection
Dim sProduto As String
Dim sProdutoMasc As String
Dim objFornecedor As ClassFornecedor
Dim objFilialFornecedor As ClassFilialFornecedor
Dim objRegTick As ClassRegTick
Dim iLinha As Integer
Dim iFilial As Integer
Dim objNFiscal As ClassNFiscal
Dim objItemNF As ClassItemNF

On Error GoTo Erro_BotaoTrazer_Click

    Call Grid_Limpa(objGridItens)
    
    Set gcolNF = New Collection

    sProduto = "0000001"
    sProdutoMasc = "000.0001"

    If StrParaDate(DataInicial.Text) = DATA_NULA Then gError 181911
    If StrParaDate(DataFinal.Text) = DATA_NULA Then gError 181912
    If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 181913
   
    lErro = CF("RegTick_Le", StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text), colRegTick)
    If lErro <> SUCESSO Then gError 181914
    
    iLinha = 0
    For Each objRegTick In colRegTick
    
        Set objFornecedor = New ClassFornecedor
        Set objFilialFornecedor = New ClassFilialFornecedor
        Set objNFiscal = New ClassNFiscal
        Set objItemNF = New ClassItemNF
    
        iLinha = iLinha + 1
    
        GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = sProdutoMasc
        
        objFornecedor.sCgc = Replace(Replace(Replace(Replace(objRegTick.sCgc, ".", ""), "-", ""), "/", ""), "\", "")
        
        lErro = CF("Fornecedor_Le_Cgc", objFornecedor, iFilial)
        If lErro <> SUCESSO And lErro <> 6694 And lErro <> 6697 Then gError 181915
        
        If lErro <> SUCESSO Then gError 181943
        
        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
        objFilialFornecedor.iCodFilial = iFilial
        
        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 12929 Then gError 181916
        
        Set objNFiscal.objInfoUsu = objRegTick
        
        objNFiscal.lFornecedor = objFornecedor.lCodigo
        objNFiscal.iFilialForn = objFilialFornecedor.iCodFilial
        objNFiscal.dtDataEmissao = StrParaDate(Format(objRegTick.dtTick_DtHrPesoFinal, "dd/mm/yyyy"))
        objNFiscal.dtDataEntrada = StrParaDate(Format(objRegTick.dtTick_DtHrPesoFinal, "dd/mm/yyyy"))
        
        objItemNF.dPrecoUnitario = objRegTick.dTick_PesoLiqCorrUsu
        If objItemNF.dPrecoUnitario > 1000 Then objItemNF.dPrecoUnitario = objItemNF.dPrecoUnitario / 1000
        
        objItemNF.dQuantidade = objRegTick.dTick_LiquidoCorrigido / 1000
        objItemNF.sProduto = sProduto
        objItemNF.sUnidadeMed = "TON"
        
        objNFiscal.dValorFrete = Arredonda_Moeda(objItemNF.dQuantidade * StrParaDbl(objRegTick.sTick_CampoUsu2))
        objNFiscal.dValorProdutos = Arredonda_Moeda(objItemNF.dQuantidade * objItemNF.dPrecoUnitario)
        objNFiscal.dValorTotal = objNFiscal.dValorProdutos
        
        objNFiscal.ColItensNF.Add1 objItemNF
        
        gcolNF.Add objNFiscal
    
        GridItens.TextMatrix(iLinha, iGrid_Selecionado_Col) = S_MARCADO
        GridItens.TextMatrix(iLinha, iGrid_Fornecedor_Col) = objFornecedor.lCodigo & SEPARADOR & objFornecedor.sNomeReduzido
        GridItens.TextMatrix(iLinha, iGrid_Filial_Col) = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
        GridItens.TextMatrix(iLinha, iGrid_DataEmissao_Col) = Format(objNFiscal.dtDataEmissao, "dd/mm/yyyy")
        GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(objItemNF.dQuantidade)
        GridItens.TextMatrix(iLinha, iGrid_ValorFrete_Col) = Format(objNFiscal.dValorFrete, "STANDARD")
        GridItens.TextMatrix(iLinha, iGrid_ValorNF_Col) = Format(objNFiscal.dValorTotal, "STANDARD")
        'GridItens.TextMatrix(iLinha, iGrid_ValorProdutos_Col) = Format(objNFiscal.dValorProdutos, "STANDARD")
        GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(objItemNF.dPrecoUnitario, "STANDARD")
    
        objNFiscal.dValorFrete = 0 'Limpa para não afetar a tributação
    
    Next
    
    objGridItens.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridItens)
    
    Call Calcula_Totais

    Exit Sub

Erro_BotaoTrazer_Click:

    Select Case gErr
    
        Case 181911
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INIC_NAO_PREENCHIDA", gErr)
            DataInicial.SetFocus
            
        Case 181912
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_NAO_PREENCHIDA", gErr)
            DataFinal.SetFocus
            
        Case 181913
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
            DataInicial.SetFocus
            
        Case 181914 To 181916
        
        Case 181943
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO2", gErr, objFornecedor.sCgc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub
    
End Sub

Function Limpa_Tela_NFRaiz() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_NFRaiz
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    ToneladasSel.Caption = ""
    ToneladasTotal.Caption = ""
    ValorSel.Caption = ""
    ValorTotal.Caption = ""

    Call Grid_Limpa(objGridItens)
    
    Set gcolNF = New Collection
    
    iAlterado = 0

    Limpa_Tela_NFRaiz = SUCESSO

    Exit Function

Erro_Limpa_Tela_NFRaiz:

    Limpa_Tela_NFRaiz = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Grava a Máquina
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 181917

    'Limpa Tela
    Call Limpa_Tela_NFRaiz
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 181917

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 181918

    Call Limpa_Tela_NFRaiz
    
    'Fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 181918

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim colNFs As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Move_Tela_Memoria(colNFs)
    If lErro <> SUCESSO Then gError 181919
    
    If colNFs.Count = 0 Then gError 181942

    'Gera e Gravas as notas fiscais automaticamente
    lErro = CF("NFiscaisRaiz_Importar", colNFs)
    If lErro <> SUCESSO Then gError 181920
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 181919, 181920
        
        Case 181942
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_NAO_SELECIONADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal colNFs As Collection) As Long

Dim lErro As Long
Dim objNF As ClassNFiscal
Dim iIndice As Integer

On Error GoTo Erro_Move_Tela_Memoria

    iIndice = 0
    For Each objNF In gcolNF
    
        iIndice = iIndice + 1
    
        'Se a nota estiver marcada
        If StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Selecionado_Col)) = MARCADO Then
            colNFs.Add objNF
        End If
    
    Next

    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    Exit Function

End Function

Private Function Marca_Desmarca(ByVal bFlag As Boolean) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Marca_Desmarca

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        If bFlag Then
            GridItens.TextMatrix(iIndice, iGrid_Selecionado_Col) = S_MARCADO
        Else
            GridItens.TextMatrix(iIndice, iGrid_Selecionado_Col) = S_DESMARCADO
        End If
    
    Next
    
    Call Grid_Refresh_Checkbox(objGridItens)
    
    Call Calcula_Totais

    Marca_Desmarca = SUCESSO
    
    Exit Function
    
Erro_Marca_Desmarca:

    Marca_Desmarca = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    Exit Function

End Function

Private Sub BotaoDesmarcar_Click()
    Call Marca_Desmarca(False)
End Sub

Private Sub BotaoMarcar_Click()
    Call Marca_Desmarca(True)
End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    'verifica se a data está preenchida
    If Len(Trim(DataFinal.ClipText)) > 0 Then

        'verifica se a data final é válida
        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then gError 181944

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 181944

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    'verifica se a data está preenchida
    If Len(Trim(DataInicial.ClipText)) > 0 Then

        'verifica se a data final é válida
        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then gError 181945

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 181945

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFinal_DownClick()
'diminui a data final

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_DownClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 181946

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_DownClick:

    Select Case gErr

        Case 181946

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFinal_UpClick()
'aumenta a data final

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_UpClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 181947

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_UpClick:

    Select Case gErr

        Case 181947

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicial_DownClick()
'diminui a data inicial

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicial_DownClick

    DataInicial.SetFocus

    If Len(DataInicial.ClipText) > 0 Then

        sData = DataInicial.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 181948

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_DownClick:

    Select Case gErr

        Case 181948

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub


End Sub

Private Sub UpDownDataInicial_UpClick()
'aumenta a data inicial

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicial_UpClick

    DataInicial.SetFocus

    If Len(DataInicial.ClipText) > 0 Then

        sData = DataInicial.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 181949

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_UpClick:

    Select Case gErr

        Case 181949

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Function Calcula_Totais() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dTonSel As Double
Dim dVlrSel As Double
Dim iSel As Double
Dim dTonTot As Double
Dim dVlrTot As Double
Dim iTot As Double

On Error GoTo Erro_Calcula_Totais

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        dTonTot = dTonTot + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        dVlrTot = dVlrTot + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorNF_Col))
        iTot = iTot + 1
    
        'Se a nota estiver marcada
        If StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Selecionado_Col)) = MARCADO Then
        
            dTonSel = dTonSel + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
            dVlrSel = dVlrSel + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorNF_Col))
            iSel = iSel + 1
        
        End If
    
    Next
    
    If iTot > 1 Then
        FrameTotal.Caption = "Total: " & CStr(iTot) & " itens"
    Else
        FrameTotal.Caption = "Total: " & CStr(iTot) & " item"
    End If
    
    If iSel > 1 Then
        FrameSel.Caption = "Selecionados: " & CStr(iSel) & " itens"
    Else
        FrameSel.Caption = "Selecionados: " & CStr(iSel) & " item"
    End If
    
    ToneladasSel.Caption = Formata_Estoque(dTonSel)
    ValorSel.Caption = Format(dVlrSel, "STANDARD")
    
    ToneladasTotal.Caption = Formata_Estoque(dTonTot)
    ValorTotal.Caption = Format(dVlrTot, "STANDARD")
    
    Calcula_Totais = SUCESSO
    
    Exit Function
    
Erro_Calcula_Totais:

    Calcula_Totais = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    Exit Function

End Function
