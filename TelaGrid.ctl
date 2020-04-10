VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl TelaGrid 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   8700
   Begin VB.ListBox Grids 
      Height          =   1620
      ItemData        =   "TelaGrid.ctx":0000
      Left            =   4410
      List            =   "TelaGrid.ctx":0002
      TabIndex        =   9
      Top             =   315
      Width           =   4035
   End
   Begin VB.ComboBox NomeArq 
      Height          =   315
      ItemData        =   "TelaGrid.ctx":0004
      Left            =   795
      List            =   "TelaGrid.ctx":0006
      TabIndex        =   7
      ToolTipText     =   "Digite ou Escolha o Nome da View ou Tabela do Banco de Dados."
      Top             =   885
      Width           =   3480
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   3690
      Picture         =   "TelaGrid.ctx":0008
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5325
      Width           =   1005
   End
   Begin VB.CommandButton BotaoIncluir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15
      Picture         =   "TelaGrid.ctx":0162
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Inclui a Operação na Árvore do Roteiro"
      Top             =   1425
      Width           =   1335
   End
   Begin VB.CommandButton BotaoRemover 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1485
      Picture         =   "TelaGrid.ctx":19B0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exclui a Operação da Árvore do Roteiro"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame Campos 
      Caption         =   "Campos"
      Height          =   3225
      Left            =   60
      TabIndex        =   0
      Top             =   2055
      Width           =   8535
      Begin MSMask.MaskEdBox Ordem 
         Height          =   270
         Left            =   3585
         TabIndex        =   12
         Top             =   1455
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Controle 
         Height          =   270
         Left            =   1350
         TabIndex        =   11
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   2850
         Left            =   150
         TabIndex        =   1
         Top             =   255
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   5027
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin MSMask.MaskEdBox Nome 
      Height          =   315
      Left            =   825
      TabIndex        =   5
      Top             =   345
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Label13 
      Caption         =   "Grids"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4410
      TabIndex        =   10
      Top             =   30
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Arquivo:"
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
      Left            =   30
      TabIndex        =   8
      Top             =   930
      Width           =   765
   End
   Begin VB.Label ProdutoLabel 
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
      Left            =   210
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   6
      Top             =   375
      Width           =   555
   End
End
Attribute VB_Name = "TelaGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjTela As ClassCriaTela

Public objGridItens As AdmGrid
Dim iGrid_Controle_Col As Integer
Dim iGrid_Ordem_Col As Integer

Public iAlterado As Integer

Const TIPO_GRID = 1
Const TIPO_FRAME = 2
Const TIPO_OUTRO = 3

Const ARQUIVO_TABELA = "U"
Const ARQUIVO_VIEW = "V"
Const TECLA_TAB = "    "
Const STRING_STRING_MAX = 255

'**** inicio do trecho a ser copiado *****
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Tela Grid"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Tela Grid"

End Function

Public Sub Show()
'    Me.Show
'    Parent.SetFocus
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

Public Sub Form_Load()

    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort
    
    Set objGridItens = New AdmGrid
    
    Call Inicializa_Grid_Itens(objGridItens)
   
    'Sinaliza que o Form_Loas ocorreu com sucesso
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

End Sub

Function Trata_Parametros(ByVal objTela As ClassCriaTela, ByVal colCombo As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Faz a variável global a tela apontar para a variável passada
    Set gobjTela = objTela
    
    For iIndice = 1 To colCombo.Count
    
        NomeArq.AddItem colCombo.Item(iIndice).sNome
        NomeArq.ItemData(NomeArq.NewIndex) = iIndice
        
    Next
        
    lErro = Traz_TelaGrid_Tela(objTela)
    If lErro <> SUCESSO Then gError 136202
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 136202
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174622)
    
    End Select
    
    Exit Function
        
End Function

Function Saida_Celula(objGridItens As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridItens)

    If lErro = SUCESSO Then

        Select Case objGridItens.objGrid.Col
        
            Case iGrid_Controle_Col

                lErro = Saida_Celula_Controle(objGridItens)
                If lErro <> SUCESSO Then gError 123221

            Case iGrid_Ordem_Col

                lErro = Saida_Celula_Ordem(objGridItens)
                If lErro <> SUCESSO Then gError 123221
                
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridItens)
        If lErro <> SUCESSO Then gError 123222

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 123221, 123222

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174623)

    End Select

    Exit Function

End Function

Private Sub BotaoOK_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoOK_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 126559
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    iAlterado = 0
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case 126559
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174624)

    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
    
        Case 136212
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174625)

    End Select

    Exit Function

End Function

Private Sub Controle_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Controle_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Controle_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Controle_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Controle()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Ordem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ordem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Ordem_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Ordem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Ordem()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set objGridItens = Nothing
    
End Sub

Private Function Saida_Celula_Ordem(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Ordem

    Set objGridInt.objControle = Ordem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132961

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridItens.Row - GridItens.FixedRows) = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    Saida_Celula_Ordem = SUCESSO

    Exit Function

Erro_Saida_Celula_Ordem:

    Saida_Celula_Ordem = gErr

    Select Case gErr

        Case 132961
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174626)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Controle(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objControle As ClassCriaControles
Dim objControleFilho As ClassCriaControles

On Error GoTo Erro_Saida_Celula_Controle

    Set objGridInt.objControle = Controle

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132961
    
    For Each objControle In gobjTela.colControles
        If UCase(objControle.sNome) = UCase(Controle.Text) Then gError 132962
        For Each objControleFilho In objControle.colControles
            If UCase(objControleFilho.sNome) = UCase(Controle.Text) Then gError 132962
        Next
    Next
    
    For iIndice = 1 To objGridInt.iLinhasExistentes
        If iIndice <> GridItens.Row Then
            If GridItens.TextMatrix(iIndice, iGrid_Controle_Col) = Controle.Text Then gError 132962
        End If
    Next

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridItens.Row - GridItens.FixedRows) = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    Saida_Celula_Controle = SUCESSO

    Exit Function

Erro_Saida_Celula_Controle:

    Saida_Celula_Controle = gErr

    Select Case gErr

        Case 132961
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                        
        Case 132962
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTROLE_JA_EXISTENTE", gErr, objControle.sNome)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174627)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Traz_TelaGrid_Tela(ByVal objTela As ClassCriaTela) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objControle As ClassCriaControles

On Error GoTo Erro_Traz_TelaGrid_Tela
   
    For Each objControle In objTela.colControles
            
        If objControle.iTipo = TIPO_GRID Then
        
            Grids.AddItem objControle.sNome
            Grids.ItemData(Grids.NewIndex) = objControle.iOrdem
        
        End If
    
    Next

    objGridItens.iLinhasExistentes = iIndice
    
    Traz_TelaGrid_Tela = SUCESSO

    Exit Function

Erro_Traz_TelaGrid_Tela:

    Traz_TelaGrid_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174628)
    
    End Select
    
    Exit Function
    
End Function

Function Move_TelaGrid_Memoria(ByVal objControle As ClassCriaControles) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objControleFilho As ClassCriaControles

On Error GoTo Erro_Move_TelaGrid_Memoria

    Set objControle.colControles = New Collection

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        Set objControleFilho = New ClassCriaControles
        
        objControleFilho.sNome = GridItens.TextMatrix(iIndice, iGrid_Controle_Col)
        objControleFilho.iOrdem = GridItens.TextMatrix(iIndice, iGrid_Ordem_Col)
        objControleFilho.sGrid = objControle.sNome
        objControleFilho.iTipo = TIPO_OUTRO
        
        objControle.colControles.Add objControleFilho
    
    Next
      
    Move_TelaGrid_Memoria = SUCESSO

    Exit Function

Erro_Move_TelaGrid_Memoria:

    Move_TelaGrid_Memoria = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174629)
    
    End Select
    
    Exit Function
    
End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Controle")
    objGridInt.colColuna.Add ("Ordem")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Controle.Name)
    objGridInt.colCampo.Add (Ordem.Name)

    iGrid_Controle_Col = 1
    iGrid_Ordem_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Public Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Public Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

End Sub

Private Sub BotaoIncluir_Click()

Dim objControle As New ClassCriaControles
Dim colSaida As New Collection
Dim colCampos As New Collection

    objControle.sNome = Nome.Text
    objControle.iTipo = TIPO_GRID
    objControle.sTipo = "MSFlexGrid"
    
    Call Move_TelaGrid_Memoria(objControle)
    
    colCampos.Add "iOrdem"
    
    Call Ordena_Colecao(objControle.colControles, colSaida, colCampos)
    
    Set objControle.colControles = colSaida
    
    gobjTela.colControles.Add objControle

    Grids.AddItem objControle.sNome
    
    Call Limpa_Tela_TelaGrid

End Sub

Private Sub NomeArq_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colColunasTabelas As New Collection

On Error GoTo Erro_NomeArq_Validate

    If Len(Trim(NomeArq.Text)) <> 0 Then
        
        lErro = ColunasTabelas_Le(NomeArq.Text, colColunasTabelas)
        If lErro <> SUCESSO Then gError 131721
        
        lErro = Preenche_GridColunas(colColunasTabelas)
        If lErro <> SUCESSO Then gError 131722
            
    End If

    Exit Sub

Erro_NomeArq_Validate:

    Cancel = True

    Select Case gErr
    
        Case 131721 To 131722

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174630)

    End Select

    Exit Sub

End Sub

Private Function Preenche_GridColunas(ByVal colColunasTabelas As Collection) As Long

Dim lErro As Long
Dim objColunasTabelas As ClassColunasTabelas
Dim iIndice As Integer

On Error GoTo Erro_Preenche_GridColunas
    
    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)
    
    iIndice = 0
    
    For Each objColunasTabelas In colColunasTabelas
    
        iIndice = iIndice + 1
    
        GridItens.TextMatrix(iIndice, iGrid_Controle_Col) = CStr(objColunasTabelas.sColuna)
        GridItens.TextMatrix(iIndice, iGrid_Ordem_Col) = CStr(iIndice)
    
    Next
    
    Call Grid_Refresh_Checkbox(objGridItens)
            
    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice
            
    Preenche_GridColunas = SUCESSO
    
    Exit Function
    
Erro_Preenche_GridColunas:

    Preenche_GridColunas = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174631)
    
    End Select
    
    Exit Function
    
End Function

Private Function Traz_Grids_Tela(ByVal objControle As ClassCriaControles) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objControleFilho As ClassCriaControles

On Error GoTo Erro_Traz_Grids_Tela
    
    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    Nome.Text = objControle.sNome
    
    For Each objControleFilho In objControle.colControles
        
        iIndice = iIndice + 1
        
        GridItens.TextMatrix(iIndice, iGrid_Controle_Col) = objControleFilho.sNome
        GridItens.TextMatrix(iIndice, iGrid_Ordem_Col) = objControleFilho.iOrdem
    
    Next
    
    Call Grid_Refresh_Checkbox(objGridItens)
            
    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice
            
    Traz_Grids_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Grids_Tela:

    Traz_Grids_Tela = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174632)
    
    End Select
    
    Exit Function
    
End Function

Private Sub Grids_DblClick()

Dim lErro As Long
Dim objControle As New ClassCriaControles
Dim objControleAux As New ClassCriaControles

On Error GoTo Erro_Grids_DblClick

    'Guarda o valor do codigo do Tipo da Mão-de-Obra selecionado na ListBox Tipos
    objControle.sNome = Grids.Text
    
    For Each objControleAux In gobjTela.colControles
        If objControleAux.sNome = objControle.sNome Then Exit For
    Next

    'Mostra os dados do TiposDeMaodeObra na tela
    lErro = Traz_Grids_Tela(objControleAux)
    If lErro <> SUCESSO Then gError 137557

    Me.Show
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_Grids_DblClick:

    Grids.SetFocus

    Select Case gErr

    Case 137557
        'erro tratado na rotina chamada
    
    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174633)

    End Select

    Exit Sub

End Sub

'#############################################
'INSERIDO POR WAGNER
'ROTINAS DE BD
Public Function ColunasTabelas_Le(ByVal sNomeArq As String, ByVal colColunasTabelas As Collection) As Long
'Lê syscolumns e sysobjects

Dim lErro As Long
Dim lComando As Long
Dim tColunasTabelas As typeColunasTabelas
Dim objColunasTabelas As ClassColunasTabelas

On Error GoTo Erro_ColunasTabelas_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 131700
    
    With tColunasTabelas
    
        'Aloca espaço no buffer
        .sArquivo = String(STRING_STRING_MAX, 0)
        .sArquivoTipo = String(STRING_STRING_MAX, 0)
        .sColuna = String(STRING_STRING_MAX, 0)
        .sColunaTipo = String(STRING_STRING_MAX, 0)
    
        'Le o syscolumns e sysobjects
        lErro = Comando_Executar(lComando, "SELECT O.Name,O.xtype, C.Name, T.Name, C.length, CONVERT(int,C.xprec) " & _
                                            "FROM syscolumns AS C, sysobjects AS O, systypes AS T " & _
                                            "WHERE O.id = C.id AND C.xtype = T.xtype AND O.name = ? ORDER BY C.colorder ", _
                                            .sArquivo, .sArquivoTipo, .sColuna, .sColunaTipo, .lColunaTamanho, .lColunaPrecisao, sNomeArq)
        If lErro <> AD_SQL_SUCESSO Then gError 131701
    
    End With

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131702

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objColunasTabelas = New ClassColunasTabelas

        With objColunasTabelas
        
            .sArquivo = Trim(tColunasTabelas.sArquivo)
            .sArquivoTipo = Trim(tColunasTabelas.sArquivoTipo)
            .sColuna = Trim(tColunasTabelas.sColuna)
            .sColunaTipo = Trim(tColunasTabelas.sColunaTipo)
            If .sArquivoTipo = ARQUIVO_VIEW And .sColunaTipo = "string" Then
                .lColunaTamanho = 255
                .lColunaPrecisao = 255
            Else
                .lColunaTamanho = tColunasTabelas.lColunaTamanho
                .lColunaPrecisao = tColunasTabelas.lColunaPrecisao
            End If
            .lTamanhoTela = 1100 + (.lColunaTamanho * 10)
        
        End With

        colColunasTabelas.Add objColunasTabelas

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131703

    Loop

    Call Comando_Fechar(lComando)

    ColunasTabelas_Le = SUCESSO

    Exit Function

Erro_ColunasTabelas_Le:

    ColunasTabelas_Le = gErr

    Select Case gErr

        Case 131700
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131701, 131702, 131703
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SYSCOLUMNS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174634)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Limpa_Tela_TelaGrid() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TelaGrid

    'Funcção genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridItens)
    
    NomeArq.ListIndex = -1
    
    Limpa_Tela_TelaGrid = SUCESSO

    Exit Function

Erro_Limpa_Tela_TelaGrid:

    Limpa_Tela_TelaGrid = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174635)

    End Select

    Exit Function

End Function

Private Sub BotaoAlterar_Click()

End Sub

Private Sub BotaoRemover_Click()

Dim objControle As New ClassCriaControles
Dim iIndice As Integer
Dim bAchou As Boolean

On Error GoTo Erro_BotaoRemover_Click

    If bAchou >= 0 Then
        Grids.RemoveItem (Grids.ListIndex)
    End If
    
    iIndice = 0
    bAchou = False
    
    For Each objControle In gobjTela.colControles
        
        iIndice = iIndice + 1
        
        If objControle.sNome = Nome.Text And objControle.iTipo = TIPO_GRID Then
            bAchou = True
            Exit For
        End If
    
    Next
    
    If bAchou Then
        gobjTela.colControles.Remove (iIndice)
    End If
    
    Call Limpa_Tela_TelaGrid
    
    Exit Sub
    
Erro_BotaoRemover_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174636)

    End Select

    Exit Sub
    
End Sub
'####################################################################################
