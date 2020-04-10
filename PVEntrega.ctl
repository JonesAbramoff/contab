VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl PVEntrega 
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   KeyPreview      =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   6255
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   3045
      Picture         =   "PVEntrega.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3675
      Width           =   1005
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   1320
      Picture         =   "PVEntrega.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3660
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   "Produto "
      Height          =   660
      Left            =   165
      TabIndex        =   1
      Top             =   75
      Width           =   5805
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2445
         TabIndex        =   4
         Top             =   225
         Width           =   3210
      End
      Begin VB.Label LabelProduto 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Left            =   135
         TabIndex        =   3
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Produto 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   930
         TabIndex        =   2
         Top             =   225
         Width           =   1485
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Entrega"
      Height          =   2730
      Left            =   165
      TabIndex        =   5
      Top             =   810
      Width           =   5805
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   225
         Left            =   2325
         TabIndex        =   12
         Top             =   735
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   397
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PedidoCliente 
         Height          =   225
         Left            =   3870
         TabIndex        =   11
         Top             =   735
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataEntrega 
         Height          =   225
         Left            =   705
         TabIndex        =   10
         Top             =   735
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   1860
         Left            =   120
         TabIndex        =   0
         Top             =   225
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label QuantidadeTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2970
         TabIndex        =   9
         Top             =   2220
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Total:"
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
         Index           =   2
         Left            =   1335
         TabIndex        =   8
         Top             =   2280
         Width           =   1545
      End
   End
End
Attribute VB_Name = "PVEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gcolDatas As Collection

Public objGridItens As AdmGrid
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_DataEntrega_Col As Integer

Public iAlterado As Integer

'**** inicio do trecho a ser copiado *****
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Data de Entrega"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "PVEntrega"

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


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
    End If

End Sub

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

Function Trata_Parametros(ByVal sProdutoTela As String, Optional ByVal colDatas As Collection) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Faz a variável global a tela apontar para a variável passada
    Set gcolDatas = colDatas
    
    Produto.Caption = sProdutoTela
        
    lErro = Traz_Datas_Tela(sProdutoTela, colDatas)
    If lErro <> SUCESSO Then gError 182557
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 182557
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182558)
    
    End Select
    
    Exit Function
        
End Function

Function Saida_Celula(objGridItens As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridItens)

    If lErro = SUCESSO Then

        Select Case GridItens.Col

            Case iGrid_DataEntrega_Col

                lErro = Saida_Celula_DataEntrega(objGridItens)
                If lErro <> SUCESSO Then gError 182559
        
            Case iGrid_Quantidade_Col
        
                lErro = Saida_Celula_Quantidade(objGridItens)
                If lErro <> SUCESSO Then gError 182560

        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridItens)
        If lErro <> SUCESSO Then gError 182561

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 182559 To 182561

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182562)

    End Select

    Exit Function

End Function

Private Sub BotaoCancela_Click()
    
    'Nao mexer no obj da tela
    giRetornoTela = vbOK
    
    Unload Me
    
    Exit Sub

End Sub

Private Sub BotaoOK_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoOK_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 182563
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    iAlterado = 0
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case 182563
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182564)

    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro
    
    lErro = Move_Datas_Memoria()
    If lErro <> SUCESSO Then gError 182565
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
    
        Case 182565
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182566)

    End Select

    Exit Function

End Function

Private Sub DataEntrega_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntrega_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub DataEntrega_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub DataEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataEntrega()
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

    Set objGridItens.objControle = Quantidade()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridItens = Nothing
    
End Sub

Private Function Saida_Celula_DataEntrega(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataEntrega

    Set objGridInt.objControle = DataEntrega

    'Verifica se valor está preenchido
    If Len(DataEntrega.ClipText) > 0 Then

        'Critica se valor é positivo
        lErro = Data_Critica(DataEntrega.Text)
        If lErro <> SUCESSO Then gError 182567
 
        If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 182568

    Saida_Celula_DataEntrega = SUCESSO

    Exit Function

Erro_Saida_Celula_DataEntrega:

    Saida_Celula_DataEntrega = gErr

    Select Case gErr
    
        Case 182567, 182568
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182569)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidadeque está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Quantidade.Text) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 182571

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

        If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        End If
        
    End If

    GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Quantidade.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 182572
    
    Call Calcula_Totais
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr
    
        Case 182571, 182572
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182573)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Traz_Datas_Tela(ByVal sProdutoTela As String, colDatas As Collection) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objPVEntrega As ClassPVEntrega
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Traz_Datas_Tela

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", sProdutoTela, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 182574

    'Armazena produto
    objProduto.sCodigo = sProdutoFormatado
        
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 182575
    
    Descricao.Caption = objProduto.sDescricao
        
    For Each objPVEntrega In colDatas
    
        iIndice = iIndice + 1
         
        GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col) = Format(objPVEntrega.dtDataEntrega, "dd/mm/yyyy")
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objPVEntrega.dQuantidade)
        
    Next
    
    objGridItens.iLinhasExistentes = iIndice
    
    Call Calcula_Totais
           
    Traz_Datas_Tela = SUCESSO

    Exit Function

Erro_Traz_Datas_Tela:

    Traz_Datas_Tela = gErr
    
    Select Case gErr
    
        Case 182574, 182575
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182576)
    
    End Select
    
    Exit Function
    
End Function

Function Move_Datas_Memoria() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objPVEntrega As ClassPVEntrega
 
On Error GoTo Erro_Move_Datas_Memoria

    For iIndice = gcolDatas.Count To 1 Step -1
        gcolDatas.Remove iIndice
    Next

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        Set objPVEntrega = New ClassPVEntrega
    
        objPVEntrega.dtDataEntrega = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col))
        objPVEntrega.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
        gcolDatas.Add objPVEntrega
        
    Next
      
    Move_Datas_Memoria = SUCESSO

    Exit Function

Erro_Move_Datas_Memoria:

    Move_Datas_Memoria = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182577)
    
    End Select
    
    Exit Function
    
End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Data de Entrega")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)

    iGrid_Quantidade_Col = 1
    iGrid_DataEntrega_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

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

Private Sub Calcula_Totais()

Dim iIndice As Integer
Dim dQuantidade As Double

    For iIndice = 1 To objGridItens.iLinhasExistentes
        dQuantidade = dQuantidade + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
    Next
    
    QuantidadeTotal.Caption = Formata_Estoque(dQuantidade)

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name

        Case objControl.Name
            objControl.Enabled = True
            
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182578)

    End Select

    Exit Sub

End Sub
