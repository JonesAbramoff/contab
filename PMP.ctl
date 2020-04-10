VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl PMP 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame2 
      Caption         =   "Itens"
      Height          =   5280
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   9285
      Begin VB.CommandButton BotaoGrafico 
         Caption         =   "Cronograma Gráfico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1920
         TabIndex        =   27
         ToolTipText     =   "Abre Tela do Cronograma Gráfico das Etapas da Produção"
         Top             =   4650
         Width           =   1440
      End
      Begin VB.CommandButton BotaoCriticas 
         Caption         =   "Relatório de Críticas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3465
         TabIndex        =   24
         ToolTipText     =   "Abre o Relatório de Críticas"
         Top             =   4650
         Width           =   1665
      End
      Begin VB.CommandButton BotaoApontamento 
         Caption         =   "Apontamento das Etapas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5250
         TabIndex        =   23
         ToolTipText     =   "Abre o Relatório com todas as Etapas e seus Apontamentos"
         Top             =   4650
         Width           =   1665
      End
      Begin VB.CommandButton BotaoPreviaCarga 
         Caption         =   "Carga nos Centros de Trabalho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7035
         TabIndex        =   6
         ToolTipText     =   "Abre o Relatório de Carga nos CTs"
         Top             =   4650
         Width           =   2010
      End
      Begin VB.TextBox Descricao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   3180
         MaxLength       =   250
         TabIndex        =   5
         Top             =   2070
         Width           =   2010
      End
      Begin VB.CommandButton BotaoOP 
         Caption         =   "Ordem de Produção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   165
         TabIndex        =   4
         ToolTipText     =   "Abre a tela de Ordem de Produção"
         Top             =   4650
         Width           =   1665
      End
      Begin MSMask.MaskEdBox DataProducao 
         Height          =   315
         Left            =   5070
         TabIndex        =   7
         Top             =   885
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodOP 
         Height          =   315
         Left            =   990
         TabIndex        =   8
         Top             =   1965
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   6
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Prioridade 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   2130
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataNecess 
         Height          =   315
         Left            =   7095
         TabIndex        =   10
         Top             =   2220
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox UM 
         Height          =   315
         Left            =   4980
         TabIndex        =   11
         Top             =   2205
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Versao 
         Height          =   315
         Left            =   2535
         TabIndex        =   12
         Top             =   1785
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   2535
         TabIndex        =   13
         Top             =   1365
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   315
         Left            =   5640
         TabIndex        =   14
         Top             =   2235
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox Cliente 
         Height          =   315
         Left            =   990
         TabIndex        =   15
         Top             =   2835
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   510
         Left            =   90
         TabIndex        =   16
         Top             =   255
         Width           =   9060
         _ExtentX        =   15981
         _ExtentY        =   900
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7215
      ScaleHeight     =   495
      ScaleWidth      =   2115
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   45
      Width           =   2175
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   590
         Picture         =   "PMP.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1135
         Picture         =   "PMP.ctx":018A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   45
         Picture         =   "PMP.ctx":06BC
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1680
         Picture         =   "PMP.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Versão:"
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
      Left            =   2295
      TabIndex        =   26
      Top             =   210
      Width           =   660
   End
   Begin VB.Label VersaoPMP 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2985
      TabIndex        =   25
      Top             =   180
      Width           =   1965
   End
   Begin VB.Label Data 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   5700
      TabIndex        =   20
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Codigo 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   795
      TabIndex        =   19
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
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
      Left            =   5160
      TabIndex        =   18
      Top             =   210
      Width           =   480
   End
   Begin VB.Label LabelCodigo 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   17
      Top             =   210
      Width           =   690
   End
End
Attribute VB_Name = "PMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim iFrameAtual As Integer

Dim gobjPMP As ClassPMP

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

'Grid de Itens
Dim objGridItens As AdmGrid
Dim iGrid_Cliente_Col As Integer
Dim iGrid_CodOP_Col As Integer
Dim iGrid_Prioridade_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Versao_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_DataNecess_Col As Integer
Dim iGrid_DataProducao_Col As Integer

Private Sub BotaoApontamento_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoApontamento_Click

    'Se não tiver linha selecionada => Erro
    If GridItens.Row = 0 Then gError 138058

    lErro = objRelatorio.ExecutarDireto("Apontamento de Produção", "", 0, "", "TOP", gobjPMP.colItens.Item(GridItens.Row).sCodOPOrigem, "NNUMINTDOCPMP", CStr(gobjPMP.colItens.Item(GridItens.Row).lNumIntDoc))
    If lErro <> SUCESSO Then gError 138059

    Exit Sub

Erro_BotaoApontamento_Click:

    Select Case gErr

        Case 138058
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 138059

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165022)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGrafico_Click()
'Abre a tela de Cronograma Gráfico

Dim lErro As Long
Dim objTelaGrafico As New ClassTelaGrafico

On Error GoTo Erro_BotaoGrafico_Click:

    lErro = Atualiza_Cronograma(objTelaGrafico)
    If lErro <> SUCESSO Then gError 139107
    
    Call Chama_Tela_Nova_Instancia("TelaGrafico", objTelaGrafico)

    Exit Sub

Erro_BotaoGrafico_Click:

    Select Case gErr
    
        Case 139107

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165023)

    End Select
    
    Exit Sub
    
End Sub

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Grava a a simulação
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 138000

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 138000

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165024)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165025)

    End Select

    Exit Sub

End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Plano Mestre de Produção"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "PMP"

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
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iAlterado = 0
    
    Set gobjPMP = New ClassPMP
    
    Set objEventoCodigo = New AdmEvento
       
    'Grid Itens
    Set objGridItens = New AdmGrid
    
    'tela em questão
    Set objGridItens.objForm = Me
    
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 138001
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 138001
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165026)

    End Select

    Exit Sub

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridItens = Nothing
    Set objEventoCodigo = Nothing
    
    Set gobjPMP = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165027)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objPMP As ClassPMP) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objPMP Is Nothing) Then
    
        'Mostra os dados do Maquinas na tela
        lErro = Traz_PMP_Tela(objPMP)
        If lErro <> SUCESSO Then gError 138012
    
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 138012

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165028)

    End Select
    
    Exit Function

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If
    
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhaAnterior As Integer
Dim iLinhasExistentesAnterior As Integer

On Error GoTo Erro_GridItens_KeyDown

    'guarda as linhas do grid antes de apagar
    iLinhaAnterior = GridItens.Row
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
    
    'se apagou a linha realmente ...
    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then
        'apaga o item
        gobjPMP.colItens.Remove iLinhaAnterior
    End If
    
    Exit Sub
    
Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165029)
    
    End Select

    Exit Sub
        
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

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case Else
                    'Não há tratamento específico na saída de célula,
                    'uma vez que os campos não serão editados
        
            End Select
                        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 138002

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 138002
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165030)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("O.P.")
    objGrid.colColuna.Add ("Prior.")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("UM")
    objGrid.colColuna.Add ("Quant")
    objGrid.colColuna.Add ("Dt Necess")
    objGrid.colColuna.Add ("Dt Fim Prod")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Cliente.Name)
    objGrid.colCampo.Add (CodOP.Name)
    objGrid.colCampo.Add (Prioridade.Name)
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (Versao.Name)
    objGrid.colCampo.Add (Descricao.Name)
    objGrid.colCampo.Add (UM.Name)
    objGrid.colCampo.Add (Quantidade.Name)
    objGrid.colCampo.Add (DataNecess.Name)
    objGrid.colCampo.Add (DataProducao.Name)

    'Colunas do Grid
    iGrid_Cliente_Col = 1
    iGrid_CodOP_Col = 2
    iGrid_Prioridade_Col = 3
    iGrid_Produto_Col = 4
    iGrid_Versao_Col = 5
    iGrid_Descricao_Col = 6
    iGrid_UM_Col = 7
    iGrid_Quantidade_Col = 8
    iGrid_DataNecess_Col = 9
    iGrid_DataProducao_Col = 10

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
                
    objControl.Enabled = False
                
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165031)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cliente_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Cliente
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Versao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Versao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Versao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Versao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Versao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Descricao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Descricao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UM_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub UM_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub UM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UM
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataNecess_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataNecess_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DataNecess_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DataNecess_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataNecess
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CodOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub CodOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub CodOp_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CodOP
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Prioridade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Prioridade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Prioridade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Prioridade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Prioridade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Traz_PMP_Tela(ByVal objPMP As ClassPMP) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_PMP_Tela
       
    Set gobjPMP = objPMP
    
    Codigo.Caption = objPMP.lCodGeracao
    VersaoPMP.Caption = objPMP.sVersao
    Data.Caption = Format(objPMP.dtDataGeracao, "dd/mm/yyyy")
    
    lErro = CF("PlanoMestreProducaoItens_Le", objPMP)
    If lErro <> SUCESSO And lErro <> 136303 Then gError 138003
       
    lErro = Preenche_Grid_Itens(objPMP)
    If lErro <> SUCESSO Then gError 138004
    
    Traz_PMP_Tela = SUCESSO

    Exit Function

Erro_Traz_PMP_Tela:

    Traz_PMP_Tela = gErr

    Select Case gErr
    
        Case 138003 To 138004

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165032)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Itens(objPMP As ClassPMP) As Long

Dim lErro As Long
Dim objPMPItens As New ClassPMPItens
Dim iLinha As Integer
Dim objProduto As ClassProduto
Dim sProdutoMascarado  As String
Dim objCliente As ClassCliente
Dim sStatus As String

On Error GoTo Erro_Preenche_Grid_Itens

    Call Grid_Limpa(objGridItens)

    For Each objPMPItens In objPMP.colItens

'        If objPMPItens.objItemOP.iSituacao <> ITEMOP_SITUACAO_BAIXADA Then

            iLinha = iLinha + 1
                    
            Set objProduto = New ClassProduto
            Set objCliente = New ClassCliente
            
            objProduto.sCodigo = objPMPItens.sProduto
            
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 138005
        
            'Se não achou o Produto --> erro
            If lErro = 28030 Then gError 138006
            
            sProdutoMascarado = String(STRING_PRODUTO, 0)
    
            'Coloca a máscara no produto
            lErro = Mascara_RetornaProdutoTela(objPMPItens.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 138007
    
            GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = sProdutoMascarado
            GridItens.TextMatrix(iLinha, iGrid_Descricao_Col) = objProduto.sDescricao
    
            objCliente.lCodigo = objPMPItens.lCliente
                
            'le o nome reduzido do cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 138008
            
            If lErro <> SUCESSO Then
                GridItens.TextMatrix(iLinha, iGrid_Cliente_Col) = "INTERNO"
            Else
                'preenche com o nome reduzido do cliente
                GridItens.TextMatrix(iLinha, iGrid_Cliente_Col) = objCliente.sNomeReduzido
            End If
            
            GridItens.TextMatrix(iLinha, iGrid_CodOP_Col) = objPMPItens.sCodOPOrigem
            GridItens.TextMatrix(iLinha, iGrid_DataNecess_Col) = Format(objPMPItens.dtDataNecessidade, "dd/mm/yyyy")
            GridItens.TextMatrix(iLinha, iGrid_DataProducao_Col) = Format(objPMPItens.objItemOP.dtDataFimProd, "dd/mm/yyyy")
            GridItens.TextMatrix(iLinha, iGrid_Prioridade_Col) = objPMPItens.iPrioridade
            GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(objPMPItens.dQuantidade)
            GridItens.TextMatrix(iLinha, iGrid_UM_Col) = objPMPItens.sUM
            GridItens.TextMatrix(iLinha, iGrid_Versao_Col) = objPMPItens.sVersao

'        End If

    Next

    Call Grid_Refresh_Checkbox(objGridItens)

    objGridItens.iLinhasExistentes = iLinha
    
    Preenche_Grid_Itens = SUCESSO
    
    Exit Function

Erro_Preenche_Grid_Itens:

    Preenche_Grid_Itens = gErr
    
    Select Case gErr
    
        Case 138005, 138008
        
        Case 138006
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 138007
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165033)

    End Select
    
    Exit Function

End Function

Private Sub BotaoOP_Click()

Dim lErro As Long
Dim objOP As New ClassOrdemDeProducao
   
On Error GoTo Erro_BotaoVerEtapas_Click
    
    'Se não tiver linha selecionada => Erro
    If GridItens.Row = 0 Then gError 138009

    objOP.iFilialEmpresa = gobjPMP.colItens.Item(GridItens.Row).iFilialEmpresa
    objOP.sCodigo = gobjPMP.colItens.Item(GridItens.Row).sCodOPOrigem

    Call Chama_Tela("OrdemProducao", objOP)

    Exit Sub

Erro_BotaoVerEtapas_Click:

    Select Case gErr

        Case 138009
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165034)

    End Select
    
    Exit Sub
    
End Sub

Function Limpa_Tela_PMP() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_PMP

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Codigo.Caption = ""
    VersaoPMP.Caption = ""
    Data.Caption = ""
    
    Call Grid_Limpa(objGridItens)
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0

    Limpa_Tela_PMP = SUCESSO

    Exit Function

Erro_Limpa_Tela_PMP:

    Limpa_Tela_PMP = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165035)

    End Select

    Exit Function

End Function

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 138010

    Call Limpa_Tela_PMP
    
    'Fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 138010

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165036)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Grava o PMP Simulado no Banco de Dados
    lErro = CF("PMP_Grava", gobjPMP)
    If lErro <> SUCESSO Then gError 138051
    
    Call Limpa_Tela_PMP
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 138051

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165037)

    End Select

    Exit Function

End Function

Private Sub BotaoPreviaCarga_Click()

Dim lErro As Long
Dim objPMP As New ClassPMP
    
On Error GoTo Erro_BotaoPreviaCarga_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("RelPreviaCargaCT_Prepara", objPMP)
    If lErro <> SUCESSO Then gError 138011
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoPreviaCarga_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 138011

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165038)

    End Select
    
    Exit Sub
    
End Sub

Private Sub DataProducao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataProducao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DataProducao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DataProducao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataProducao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objPMP As New ClassPMP
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Caption)) <> 0 Then

        objPMP.lCodGeracao = StrParaDbl(Codigo.Caption)
        objPMP.dtDataGeracao = StrParaDate(Data.Caption)

    End If

    Call Chama_Tela("PMPLista", colSelecao, objPMP, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165039)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPMP As ClassPMP

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPMP = obj1

    'Mostra os dados do Maquinas na tela
    lErro = Traz_PMP_Tela(objPMP)
    If lErro <> SUCESSO Then gError 138012

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 138012

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165040)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas
Dim vbMsgRes As VbMsgBoxResult
Dim objPMP As New ClassPMP

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Caption)) = 0 Then gError 138013

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PMP", gobjPMP.lCodGeracao)

    If vbMsgRes = vbYes Then

        'Exclui a a simulação
        'lErro = CF("PMP_Exclui", gobjPMP)
        
        '##########################################
        'Não pode excluir porque pode haver itens que estão ligados a OPs
        'baixadas, a solução é excluir só o que está ativo, ou seja, gravar sem
        'itens
        '##########################################
        
        objPMP.dtDataGeracao = gobjPMP.dtDataGeracao
        objPMP.lCodGeracao = gobjPMP.lCodGeracao
        objPMP.sVersao = gobjPMP.sVersao
        
        lErro = CF("PMP_Grava", objPMP)
        If lErro <> SUCESSO Then gError 138014
    
        'Limpa Tela
        Call Limpa_Tela_PMP
    
        iAlterado = 0
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 138013
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PMP_NAO_PREENCHIDO", gErr)
        
        Case 138014

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165041)

    End Select

    Exit Sub

End Sub

Private Sub BotaoCriticas_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoCriticas_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("RelCriticasPMP_Prepara", gobjPMP)
    If lErro <> SUCESSO Then gError 138072
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoCriticas_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 138072

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165042)

    End Select
    
    Exit Sub
    
End Sub

Public Function Atualiza_Cronograma(objTelaGrafico As ClassTelaGrafico, Optional bMarcaComErros As Boolean = False) As Long
'Acerta a tela de cronograma gráfico após o retorno da tela de etapas
'Para isso remonta o objTelaGrafico com os novos dados

Dim objTelaGraficoItem As New ClassTelaGraficoItens
Dim objPMPItem As ClassPMPItens
Dim objPO As ClassPlanoOperacional
Dim objPOAux As ClassPlanoOperacional
Dim iIndice As Integer
Dim iCont As Integer
Dim sNL As String
Dim objCT As ClassCentrodeTrabalho
Dim lErro As Long
Dim bPrimeira As Boolean
Dim objOP As ClassOrdemDeProducao
Dim objBotao As ClassTelaGraficoBotao

On Error GoTo Erro_Atualiza_Cronograma

    Set objTelaGrafico.colBotoes = New Collection
    Set objTelaGrafico.colItens = New Collection
    Set objTelaGrafico.colParametros = New Collection

    Set objTelaGrafico.objTela = Me
    
    objTelaGrafico.sNomeTela = "Cronograma dos Itens do Plano Mestre de Produção"
    objTelaGrafico.iTamanhoDia = 540
    objTelaGrafico.iModal = DESMARCADO
    objTelaGrafico.iAtualizaRetornoClick = MARCADO
    objTelaGrafico.sNomeFuncAtualiza = "Atualiza_Cronograma"

    objTelaGrafico.colParametros.Add bMarcaComErros

'    Set objBotao = New ClassTelaGraficoBotao
'
'    objBotao.colParametros.Add objTelaGrafico
'    objBotao.colParametros.Add True
'    objBotao.sNome = "Marcar Etapas com Erro"
'    objBotao.sNomeFuncao = "Atualiza_Cronograma"
'    objBotao.sTextoExibicao = "Exibir Etapas com Erro"
'    objBotao.iAtualizaRetornoClick = MARCADO
'
'    objTelaGrafico.colBotoes.Add objBotao
'
'    Set objBotao = New ClassTelaGraficoBotao
'
'    objBotao.colParametros.Add objTelaGrafico
'    objBotao.colParametros.Add False
'    objBotao.sNome = "Desmarcar Etapas com Erro"
'    objBotao.sNomeFuncao = "Atualiza_Cronograma"
'    objBotao.sTextoExibicao = "Exibir Etapas com Erro"
'    objBotao.iAtualizaRetornoClick = MARCADO
'
'    objTelaGrafico.colBotoes.Add objBotao

    iIndice = 0
    
    sNL = Chr(13) & Chr(10)

    'Para cada item do Plano Mestre
    For Each objPMPItem In gobjPMP.colItens
    
        If objPMPItem.objItemOP.iSituacao <> ITEMOP_SITUACAO_BAIXADA Then
    
            iIndice = iIndice + 1
        
            iCont = 0
        
            'Para cada etapa
            For Each objPO In objPMPItem.ColPO
            
                iCont = iCont + 1
            
                Set objTelaGraficoItem = New ClassTelaGraficoItens
                Set objCT = New ClassCentrodeTrabalho
                
                Set objOP = New ClassOrdemDeProducao
                
                objOP.sCodigo = objPO.sCodOPOrigem
                objOP.iFilialEmpresa = objPO.iFilialEmpresa
            
                objTelaGraficoItem.colobj.Add objOP
                
                bPrimeira = True
                
                For Each objPOAux In objPMPItem.ColPO
                    'Se existe outra com data de início menor ou igual desde que
                    'o nó esteja depois na estrutura de árvore então o PO corrente
                    'não é o primeiro
                    If (objPOAux.dtDataInicio < objPO.dtDataInicio) Or (objPOAux.dtDataInicio = objPO.dtDataInicio And ((objPOAux.iNivel > objPO.iNivel) Or (objPOAux.iNivel = objPO.iNivel And objPOAux.iSeq > objPO.iSeq))) Then
                        bPrimeira = False
                    End If
                Next
                
                'Se for a primeira é a final em termos de data
                If iCont = 1 Then
                    objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_FIM
                End If
                           
                If bPrimeira Then
                    
                    If iCont = 1 Then
                        objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_INICIO_E_FIM
                    Else
                        objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_INICIO
                    End If
                    
                End If
                
                objCT.lNumIntDoc = objPO.lNumIntDocCT
                
                'Le o centro de trabalho
                lErro = CF("CentrodeTrabalho_Le_NumIntDoc", objCT)
                If lErro <> SUCESSO And lErro <> 134590 Then gError 139108
                        
                objTelaGraficoItem.dtDataFim = objPO.dtDataFim
                objTelaGraficoItem.dtDataInicio = objPO.dtDataInicio
                objTelaGraficoItem.sNomeTela = "OrdemProducao"
                objTelaGraficoItem.sTextoExibicao = "OP Origem: " & objPMPItem.sCodOPOrigem & sNL & "OP: " & objPO.sCodOPOrigem & sNL & "Prioridade: " & objPMPItem.iPrioridade & sNL & "Data Início: " & Format(objPO.dtDataInicio, "dd/mm/yyyy") & sNL & "Data Fim: " & Format(objPO.dtDataFim, "dd/mm/yyyy") & sNL & "CT: " & objCT.sNomeReduzido
                objTelaGraficoItem.sNome = objPO.sCodOPOrigem
                objTelaGraficoItem.iIndiceCor = iIndice
                
                If bMarcaComErros Then
                    If objPO.iStatus <> PO_STATUS_OK Then
                        objTelaGraficoItem.lCor = vbRed
                    Else
                        objTelaGraficoItem.lCor = vbGreen
                    End If
                End If
            
                objTelaGrafico.colItens.Add objTelaGraficoItem
            
            Next
            
        End If
        
    Next
    
    Atualiza_Cronograma = SUCESSO

    Exit Function

Erro_Atualiza_Cronograma:

    Atualiza_Cronograma = gErr

    Select Case gErr
    
        Case 139108

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165043)

    End Select
    
    Exit Function
    
End Function


