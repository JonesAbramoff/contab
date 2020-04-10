VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ProdSolicOCx 
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   9480
   Begin MSMask.MaskEdBox ProdSolicSRV 
      Height          =   225
      Left            =   4350
      TabIndex        =   14
      Top             =   2385
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Servico 
      Height          =   225
      Left            =   870
      TabIndex        =   13
      Top             =   2850
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.CommandButton BotaoServicos 
      Caption         =   "Serviços"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   150
      TabIndex        =   12
      Top             =   4320
      Width           =   1365
   End
   Begin VB.CommandButton BotaoProdutos 
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1755
      TabIndex        =   11
      Top             =   4320
      Width           =   1365
   End
   Begin VB.TextBox DescricaoProdSolicSRV 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   5610
      MaxLength       =   250
      TabIndex        =   10
      Top             =   3735
      Width           =   2490
   End
   Begin VB.TextBox DescricaoServico 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   2280
      MaxLength       =   250
      TabIndex        =   9
      Top             =   2895
      Width           =   2490
   End
   Begin VB.CommandButton BotaoPecaServico 
      Caption         =   "Peças x Serviços"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7875
      TabIndex        =   8
      Top             =   135
      Width           =   1365
   End
   Begin MSMask.MaskEdBox Contrato 
      Height          =   225
      Left            =   6900
      TabIndex        =   6
      Top             =   2670
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   10
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Garantia 
      Height          =   225
      Left            =   5535
      TabIndex        =   7
      Top             =   2655
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin VB.ComboBox FilialOP 
      Height          =   315
      Left            =   7005
      TabIndex        =   4
      Top             =   3225
      Width           =   1905
   End
   Begin MSMask.MaskEdBox Lote 
      Height          =   270
      Left            =   5640
      TabIndex        =   5
      Top             =   3255
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   476
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   225
      Left            =   2370
      TabIndex        =   0
      Top             =   3315
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
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   570
      Left            =   3090
      Picture         =   "ProdSolicOCx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1365
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   570
      Left            =   4500
      Picture         =   "ProdSolicOCx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid GridServicos 
      Height          =   3405
      Left            =   195
      TabIndex        =   1
      Top             =   765
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   6006
      _Version        =   393216
      Rows            =   12
      Cols            =   3
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
   End
End
Attribute VB_Name = "ProdSolicOCx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iTipoAlterado As Integer
Dim iAlterado As Integer


Dim gcolProdSolicSRVOrigem As Collection
Dim gcolProdSolicSRVDestino As Collection
Dim objGridServico As AdmGrid
Dim gobjTela As Object

Dim iGrid_Servico_Col As Integer
Dim iGrid_DescricaoServico_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_ProdSolicSRV_Col As Integer
Dim iGrid_DescricaoProdSolicSRV_Col As Integer
Dim iGrid_Lote_Col As Integer
Dim iGrid_FilialOP_Col As Integer
Dim iGrid_Garantia_Col As Integer
Dim iGrid_Contrato_Col As Integer

Dim giFrameAtual As Integer

Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Public Function Trata_Parametros(ByVal colProdSolicSRVOrigem As Collection, ByVal lSolicitacao As Long, colProdSolicSRVDestino As Collection, objTela As Object) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long
Dim objSolicSRV As New ClassSolicSRV
Dim objItensSolicSRV As ClassItensSolicSRV
Dim sProduto As String
Dim objProdSolicSRVDestino As ClassProdSolicSRV
Dim sServico As String
Dim objProdSolicSRVOrigem As ClassProdSolicSRV
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim objProduto1 As New ClassProduto

On Error GoTo Erro_Trata_Parametros

    objTela.Enabled = False
    Set gobjTela = objTela
    
    Set gcolProdSolicSRVOrigem = colProdSolicSRVOrigem
    Set gcolProdSolicSRVDestino = colProdSolicSRVDestino
    
    For iIndice = 1 To colProdSolicSRVDestino.Count
    
        Set objProdSolicSRVDestino = colProdSolicSRVDestino(iIndice)
            
        lErro = Mascara_RetornaProdutoTela(objProdSolicSRVDestino.sServicoOrcSRV, sServico)
        If lErro <> SUCESSO Then gError 186932
            
        lErro = Mascara_RetornaProdutoTela(objProdSolicSRVDestino.sProduto, sProduto)
        If lErro <> SUCESSO Then gError 186933
            
        objProduto.sCodigo = objProdSolicSRVDestino.sServicoOrcSRV
            
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 188136
            
        If lErro = 28030 Then gError 188137
            
        objProduto1.sCodigo = objProdSolicSRVDestino.sProduto
            
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto1)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 188138
            
        If lErro = 28030 Then gError 188139
            
        GridServicos.TextMatrix(iIndice, iGrid_Servico_Col) = sServico
        GridServicos.TextMatrix(iIndice, iGrid_DescricaoServico_Col) = objProduto.sDescricao
        GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objProdSolicSRVDestino.dQuantidade)
        GridServicos.TextMatrix(iIndice, iGrid_ProdSolicSRV_Col) = sProduto
        GridServicos.TextMatrix(iIndice, iGrid_DescricaoProdSolicSRV_Col) = objProduto1.sDescricao
        GridServicos.TextMatrix(iIndice, iGrid_Lote_Col) = objProdSolicSRVDestino.sLote
        If objProdSolicSRVDestino.iFilialOP > 0 Then
            GridServicos.TextMatrix(iIndice, iGrid_FilialOP_Col) = objProdSolicSRVDestino.iFilialOP
        End If
        
        If objProdSolicSRVDestino.lGarantia > 0 Then
            GridServicos.TextMatrix(iIndice, iGrid_Garantia_Col) = objProdSolicSRVDestino.lGarantia
        End If
        
        GridServicos.TextMatrix(iIndice, iGrid_Contrato_Col) = objProdSolicSRVDestino.sContrato
    
    Next
    
    'Atualiza o número de linhas existentes
    objGridServico.iLinhasExistentes = colProdSolicSRVDestino.Count
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 186916, 186930 To 186933, 188136, 188138
        
        Case 186917
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
        
        Case 188137
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 188139
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto1.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186918)
    
    End Select
    
    Exit Function
    
End Function

Private Function Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Servico)
    If lErro <> SUCESSO Then gError 188237

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdSolicSRV)
    If lErro <> SUCESSO Then gError 188238
    
    Set objGridServico = New AdmGrid
    
    Call Inicializa_Grid_Servico(objGridServico)
    
    Set objEventoServico = New AdmEvento
    Set objEventoProduto = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Function
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
        
        Case 188237, 188238
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186919)
    
    End Select
    
    Exit Function
    
End Function

Private Function Inicializa_Grid_Servico(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Desc. Serviço")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Desc. Produto")
    objGridInt.colColuna.Add ("Lote/Num.Série")
    objGridInt.colColuna.Add ("FilialOP")
    objGridInt.colColuna.Add ("Garantia")
    objGridInt.colColuna.Add ("Contrato")
    

    objGridInt.colCampo.Add (Servico.Name)
    objGridInt.colCampo.Add (DescricaoServico.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (ProdSolicSRV.Name)
    objGridInt.colCampo.Add (DescricaoProdSolicSRV.Name)
    objGridInt.colCampo.Add (Lote.Name)
    objGridInt.colCampo.Add (FilialOP.Name)
    objGridInt.colCampo.Add (Garantia.Name)
    objGridInt.colCampo.Add (Contrato.Name)

    'Controles que participam do Grid
    iGrid_Servico_Col = 1
    iGrid_DescricaoServico_Col = 2
    iGrid_Quantidade_Col = 3
    iGrid_ProdSolicSRV_Col = 4
    iGrid_DescricaoProdSolicSRV_Col = 5
    iGrid_Lote_Col = 6
    iGrid_FilialOP_Col = 7
    iGrid_Garantia_Col = 8
    iGrid_Contrato_Col = 9

    objGridInt.objGrid = GridServicos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PRODSOLICSRV + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridServicos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Servico = SUCESSO

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoProduto = Nothing
    Set objEventoServico = Nothing
        
    gobjTela.Enabled = True
    
End Sub

Private Sub BotaoCancela_Click()
    Unload Me
End Sub

Public Sub Form_Activate()
'    Call TelaIndice_Preenche(Me)
End Sub

'***************************************************
'Trecho de codigo comum as telas
'***************************************************

Public Function Form_Load_Ocx() As Object
'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Serviço x Produto Solicitado"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "ProdSolic"
End Function

Public Sub Show()
'    Parent.Show
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

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
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

Private Sub GridServicos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServico, iAlterado)

    End If

End Sub

Private Sub GridServicos_EnterCell()

    Call Grid_Entrada_Celula(objGridServico, iAlterado)

End Sub

Private Sub GridServicos_GotFocus()

    Call Grid_Recebe_Foco(objGridServico)

End Sub

Private Sub GridServicos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServico, iAlterado)
    End If


End Sub

Private Sub GridServicos_LeaveCell()

    Call Saida_Celula(objGridServico)

End Sub

Private Sub GridServicos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridServico)

End Sub

Private Sub GridServicos_Scroll()

    Call Grid_Scroll(objGridServico)

End Sub

Private Sub GridServicos_RowColChange()

    Call Grid_RowColChange(objGridServico)

End Sub

Private Sub GridServicos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridServico.iLinhasExistentes
    iItemAtual = GridServicos.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridServico)

    If objGridServico.iLinhasExistentes < iLinhasExistentesAnterior Then

         gcolProdSolicSRVDestino.Remove (iItemAtual)

    End If

End Sub

Public Sub Servico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Servico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub Servico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub Servico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = Servico
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ProdSolicSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdSolicSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub ProdSolicSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub ProdSolicSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = ProdSolicSRV
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Lote_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub Lote_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub Lote_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = Lote
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub FilialOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub FilialOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub FilialOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = FilialOP
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Garantia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Garantia_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub Garantia_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub Garantia_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = Garantia
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Contrato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Contrato_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub Contrato_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub Contrato_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = Contrato
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then

        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col

            'Se for a de Servico
            Case iGrid_Servico_Col
                lErro = Saida_Celula_Servico(objGridInt)
                If lErro <> SUCESSO Then gError 186920
    
            'Se for a de Quantidade
            Case iGrid_Quantidade_Col
                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then gError 186921
        
            'Se for a de ProdSolicSRV
            Case iGrid_ProdSolicSRV_Col
                lErro = Saida_Celula_ProdSolicSRV(objGridInt)
                If lErro <> SUCESSO Then gError 186922
        
            Case iGrid_Lote_Col
                lErro = Saida_Celula_Lote(objGridInt)
                If lErro <> SUCESSO Then gError 188094
        
            Case iGrid_FilialOP_Col
                lErro = Saida_Celula_FilialOP(objGridInt)
                If lErro <> SUCESSO Then gError 188095
        
            Case iGrid_Garantia_Col
                lErro = Saida_Celula_Garantia(objGridInt)
                If lErro <> SUCESSO Then gError 188096
        
            Case iGrid_Contrato_Col
                lErro = Saida_Celula_Contrato(objGridInt)
                If lErro <> SUCESSO Then gError 188097
        
        End Select


    End If
        

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186923
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 186920 To 186923, 188094 To 188097

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186924)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Servico(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto
Dim objProdSolicSRV As New ClassProdSolicSRV

On Error GoTo Erro_Saida_Celula_Servico

    Set objGridInt.objControle = Servico

    lErro = Servico_Saida_Celula()
    If lErro <> SUCESSO Then gError 188199

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186925

    If Len(Trim(Servico.ClipText)) > 0 Then

        If GridServicos.Row - GridServicos.FixedRows = objGridServico.iLinhasExistentes Then
            
            objGridServico.iLinhasExistentes = objGridServico.iLinhasExistentes + 1
            
            Set objProdSolicSRV = New ClassProdSolicSRV
            
            gcolProdSolicSRVDestino.Add objProdSolicSRV
            
        End If
    
    End If
    
    Saida_Celula_Servico = SUCESSO

    Exit Function

Erro_Saida_Celula_Servico:

    Saida_Celula_Servico = gErr

    Select Case gErr

        Case 186925, 188199
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 186926)

    End Select

    Exit Function

End Function

Private Function Servico_Saida_Celula() As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Servico_Saida_Celula

    If Len(Trim(Servico.ClipText)) > 0 Then

    'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", Servico.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 188171
    
        If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then
            gError 188172
        End If
    
        If objProduto.iNatureza <> NATUREZA_PROD_SERVICO Then gError 188173
    
        'Se o produto não foi encontrado ==> Pergunta se deseja criar
        If lErro = 51381 Then gError 188174
    
        GridServicos.TextMatrix(GridServicos.Row, iGrid_DescricaoServico_Col) = objProduto.sDescricao

    End If

    Servico_Saida_Celula = SUCESSO

    Exit Function

Erro_Servico_Saida_Celula:

    Servico_Saida_Celula = gErr

    Select Case gErr

        Case 188171, 188172

        Case 188173
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_NATUREZA_SERVICO", gErr, objProduto.sCodigo)

        Case 188174
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 188200)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidadeque está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Quantidade.Text) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 186927

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186928
    
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 186927, 186928
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186929)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdSolicSRV(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim sProdutoFormatado As String
Dim objProdSolicSRV As New ClassProdSolicSRV

On Error GoTo Erro_Saida_Celula_ProdSolicSRV

    Set objGridInt.objControle = ProdSolicSRV

    lErro = ProdSolicSRV_Saida_Celula()
    If lErro <> SUCESSO Then gError 188191

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186925

    If Len(Trim(ProdSolicSRV.ClipText)) > 0 Then

        If GridServicos.Row - GridServicos.FixedRows = objGridServico.iLinhasExistentes Then
            
            objGridServico.iLinhasExistentes = objGridServico.iLinhasExistentes + 1
    
            gcolProdSolicSRVDestino.Add objProdSolicSRV
        
        End If
    
    End If
    
    Saida_Celula_ProdSolicSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdSolicSRV:

    Saida_Celula_ProdSolicSRV = gErr

    Select Case gErr

        Case 186925, 188191
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 186926)

    End Select

    Exit Function

End Function

Private Function ProdSolicSRV_Saida_Celula() As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String

On Error GoTo Erro_ProdSolicSRV_Saida_Celula

    If Len(Trim(ProdSolicSRV.ClipText)) > 0 Then

        'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", ProdSolicSRV.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 188177
    
        If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then
            gError 188178
        End If
    
        'Se o produto não foi encontrado ==> Pergunta se deseja criar
        If lErro = 51381 Then gError 188179
    
        GridServicos.TextMatrix(GridServicos.Row, iGrid_DescricaoProdSolicSRV_Col) = objProduto.sDescricao
    
    End If
    
    ProdSolicSRV_Saida_Celula = SUCESSO

    Exit Function

Erro_ProdSolicSRV_Saida_Celula:

    ProdSolicSRV_Saida_Celula = gErr

    Select Case gErr

        Case 188177, 188178

        Case 188179
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 188192)

    End Select

    Exit Function

End Function



Private Function Saida_Celula_Lote(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim dQuantidade As Double
Dim lGarantia As Long
Dim sContrato As String

On Error GoTo Erro_Saida_Celula_Lote

    Set objGridInt.objControle = Lote
    
    If Len(Trim(Lote.Text)) > 0 Then
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridServicos.TextMatrix(GridServicos.Row, iGrid_ProdSolicSRV_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 188095
            
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            objProduto.sCodigo = sProdutoFormatado
                    
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 188096
                
            If lErro = 28030 Then gError 188097
                
            If gobjSRV.iVerificaLote = VERIFICA_LOTE Then
                
                'Se for rastro por lote
                If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                    
                    objRastroLote.sCodigo = Lote.Text
                    objRastroLote.sProduto = sProdutoFormatado
                    
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 188098
                    
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 188099
                    
                'Se for rastro por OP
                ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                    
                    If Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_FilialOP_Col))) > 0 Then
                        
                        objRastroLote.sCodigo = Lote.Text
                        objRastroLote.sProduto = sProdutoFormatado
                        objRastroLote.iFilialOP = Codigo_Extrai(GridServicos.TextMatrix(GridServicos.Row, iGrid_FilialOP_Col))
                        
                        'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                        lErro = CF("RastreamentoLote_Le", objRastroLote)
                        If lErro <> SUCESSO And lErro <> 75710 Then gError 188100
                        
                        'Se não encontrou --> Erro
                        If lErro = 75710 Then gError 188101
                        
                    End If
                    
                End If
            
            End If
        
        End If
    
    End If
            
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188102

    If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then

        'se tiver garantia associada, traz para a tela
        lErro = CF("Pesquisa_Garantia", GridServicos.TextMatrix(GridServicos.Row, iGrid_ProdSolicSRV_Col), GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col), GridServicos.TextMatrix(GridServicos.Row, iGrid_Lote_Col), StrParaInt(GridServicos.TextMatrix(GridServicos.Row, iGrid_FilialOP_Col)), lGarantia)
        If lErro <> SUCESSO Then gError 188103

        If lGarantia <> 0 And Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_Garantia_Col))) = 0 Then GridServicos.TextMatrix(GridServicos.Row, iGrid_Garantia_Col) = lGarantia

    End If

    If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then

        'se tiver contrato associada, traz para a tela
        lErro = CF("Pesquisa_Contrato", GridServicos.TextMatrix(GridServicos.Row, iGrid_ProdSolicSRV_Col), GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col), GridServicos.TextMatrix(GridServicos.Row, iGrid_Lote_Col), StrParaInt(GridServicos.TextMatrix(GridServicos.Row, iGrid_FilialOP_Col)), sContrato)
        If lErro <> SUCESSO Then gError 188104

        If Len(Trim(sContrato)) > 0 And Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_Contrato_Col))) = 0 Then GridServicos.TextMatrix(GridServicos.Row, iGrid_Contrato_Col) = sContrato

    End If

    Saida_Celula_Lote = SUCESSO

    Exit Function

Erro_Saida_Celula_Lote:

    Saida_Celula_Lote = gErr

    Select Case gErr

        Case 188095, 188096, 188098, 188100, 188102 To 188104
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 188097
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 188099, 188101
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_NUMSERIE_NAO_CADASTRADO", gErr, objRastroLote.sCodigo, objRastroLote.sProduto, objRastroLote.iFilialOP)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188105)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialOP(objGridInt As AdmGrid) As Long
'Faz a saida de celula da Filial da Ordem de Produção

Dim lErro As Long
Dim objFilialOP As New AdmFiliais
Dim iCodigo As Integer
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objRastroLote As New ClassRastreamentoLote
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim dQuantidade As Double
Dim lGarantia As Long
Dim sContrato As String

On Error GoTo Erro_Saida_Celula_FilialOP

    Set objGridInt.objControle = FilialOP

    If Len(Trim(FilialOP.Text)) <> 0 Then
            
        'Verifica se é uma FilialOP selecionada
        If FilialOP.Text <> FilialOP.List(FilialOP.ListIndex) Then
        
            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialOP, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 188106
    
            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then
    
                objFilialOP.iCodFilial = iCodigo
    
                'Pesquisa se existe FilialOP com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 188107
        
                'Se não encontrou a FilialOP
                If lErro = 27378 Then gError 188108
        
                'coloca na tela
                FilialOP.Text = iCodigo & SEPARADOR & objFilialOP.sNome
            
            
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 188109
                    
        End If
        
        If gobjSRV.iVerificaLote = VERIFICA_LOTE Then
        
            If Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_Lote_Col))) > 0 Then
                
                lErro = CF("Produto_Formata", GridServicos.TextMatrix(GridServicos.Row, iGrid_ProdSolicSRV_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 188110
                                    
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                
                    objRastroLote.sCodigo = GridServicos.TextMatrix(GridServicos.Row, iGrid_Lote_Col)
                    objRastroLote.sProduto = sProdutoFormatado
                    objRastroLote.iFilialOP = Codigo_Extrai(FilialOP.Text)
                
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 188111
                    
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 188112
                                
                End If
                
            End If
        
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188113

    If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then

        'se tiver garantia associada, traz para a tela
        lErro = CF("Pesquisa_Garantia", GridServicos.TextMatrix(GridServicos.Row, iGrid_ProdSolicSRV_Col), GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col), GridServicos.TextMatrix(GridServicos.Row, iGrid_Lote_Col), StrParaInt(GridServicos.TextMatrix(GridServicos.Row, iGrid_FilialOP_Col)), lGarantia)
        If lErro <> SUCESSO Then gError 188114

        If lGarantia <> 0 And Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_Garantia_Col))) = 0 Then GridServicos.TextMatrix(GridServicos.Row, iGrid_Garantia_Col) = lGarantia

    End If

    If gobjSRV.iGarantiaAutoSolic = GARANTIA_AUTOMATICA_SOLICITACAO Then

        'se tiver contrato associada, traz para a tela
        lErro = CF("Pesquisa_Contrato", GridServicos.TextMatrix(GridServicos.Row, iGrid_ProdSolicSRV_Col), GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col), GridServicos.TextMatrix(GridServicos.Row, iGrid_Lote_Col), StrParaInt(GridServicos.TextMatrix(GridServicos.Row, iGrid_FilialOP_Col)), sContrato)
        If lErro <> SUCESSO Then gError 188115

        If Len(Trim(sContrato)) > 0 And Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_Contrato_Col))) = 0 Then GridServicos.TextMatrix(GridServicos.Row, iGrid_Contrato_Col) = sContrato

    End If

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case 188106, 188107, 188110, 188111, 188113, 188114, 188115
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 188108
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 188109
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 188112
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 188116)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Garantia(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Garantia está deixando de ser a corrente

Dim lErro As Long
Dim objGarantia As New ClassGarantia

On Error GoTo Erro_Saida_Celula_Garantia

    Set objGridInt.objControle = Garantia

    If Len(Trim(Garantia.Text)) > 0 Then

        lErro = Long_Critica(Garantia.Text)
        If lErro <> SUCESSO Then gError 188117

        objGarantia.iFilialEmpresa = giFilialEmpresa
        objGarantia.lCodigo = StrParaLong(Garantia.Text)
        objGarantia.sProduto = GridServicos.TextMatrix(GridServicos.Row, iGrid_ProdSolicSRV_Col)
        objGarantia.sServico = GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col)
        objGarantia.sLote = GridServicos.TextMatrix(GridServicos.Row, iGrid_Lote_Col)
        objGarantia.iFilialOP = StrParaInt(GridServicos.TextMatrix(GridServicos.Row, iGrid_FilialOP_Col))

        lErro = CF("Testa_Garantia", objGarantia)
        If lErro <> SUCESSO Then gError 188118
        
    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188119
    
    Saida_Celula_Garantia = SUCESSO

    Exit Function

Erro_Saida_Celula_Garantia:

    Saida_Celula_Garantia = gErr

    Select Case gErr

        Case 188117 To 188119
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188120)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Contrato(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Contrato está deixando de ser a corrente

Dim lErro As Long
Dim objItensDeContratoSrv As New ClassItensDeContratoSrv

On Error GoTo Erro_Saida_Celula_Contrato

    Set objGridInt.objControle = Contrato

    If Len(Trim(Contrato.Text)) > 0 Then

        lErro = Long_Critica(Contrato.Text)
        If lErro <> SUCESSO Then gError 188128

        objItensDeContratoSrv.iFilialEmpresa = giFilialEmpresa
        objItensDeContratoSrv.lCodigo = StrParaLong(Contrato.Text)
        objItensDeContratoSrv.sProduto = GridServicos.TextMatrix(GridServicos.Row, iGrid_ProdSolicSRV_Col)
        objItensDeContratoSrv.sServico = GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col)
        objItensDeContratoSrv.sLote = GridServicos.TextMatrix(GridServicos.Row, iGrid_Lote_Col)
        objItensDeContratoSrv.iFilialOP = StrParaInt(GridServicos.TextMatrix(GridServicos.Row, iGrid_FilialOP_Col))
        
        lErro = CF("Testa_Contrato", objItensDeContratoSrv)
        If lErro <> SUCESSO Then gError 188129

    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188130
    
    Saida_Celula_Contrato = SUCESSO

    Exit Function

Erro_Saida_Celula_Contrato:

    Saida_Celula_Contrato = gErr

    Select Case gErr

        Case 188128 To 188130
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188131)

    End Select

    Exit Function

End Function

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim objGarantia As New ClassGarantia

On Error GoTo Erro_BotaoOK_Click

    lErro = Valida_Dados_Tela()
    If lErro <> SUCESSO Then gError 186942

    'Move os dados da tela para o objRelacionamentoClie
    lErro = Move_ProdSolicSRV_Memoria()
    If lErro <> SUCESSO Then gError 186943

    lErro = gobjTela.Atualiza_Tela()
    If lErro <> SUCESSO Then gError 188207
    
    iAlterado = 0

    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case 186942 To 186943, 188207
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186944)

    End Select

    Exit Sub

End Sub

Private Function Valida_Dados_Tela() As Long
'Verifica se os dados da tela são válidos

Dim lErro As Long
Dim iIndice As Integer
Dim dQuantidade As Double
Dim objProdSolicSRV As ClassProdSolicSRV
Dim colProdSolicSRV As New Collection
Dim iAchou As Integer
Dim objProdSolicSRVOrigem As ClassProdSolicSRV
Dim sServicoFormatado As String
Dim iServicoPreenchido As Integer
Dim sProdSolicSRV As String
Dim iProdutoPreenchido As Integer
Dim iIndice1 As Integer
Dim sServicoFormatado1 As String
Dim iServicoPreenchido1 As Integer
Dim sProdSolicSRV1 As String
Dim iProdutoPreenchido1 As Integer

On Error GoTo Erro_Valida_Dados_Tela

    For iIndice = 1 To objGridServico.iLinhasExistentes

        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Servico_Col))) = 0 Then gError 186935
        
        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 186936
        
        If StrParaDbl(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col)) <= 0 Then gError 186938
        
        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_ProdSolicSRV_Col))) = 0 Then gError 186937

        iAchou = 0

        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_Servico_Col), sServicoFormatado, iServicoPreenchido)
        If lErro <> SUCESSO Then gError 186940

        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_ProdSolicSRV_Col), sProdSolicSRV, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 186980


        For iIndice1 = iIndice + 1 To objGridServico.iLinhasExistentes

            lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice1, iGrid_Servico_Col), sServicoFormatado1, iServicoPreenchido1)
            If lErro <> SUCESSO Then gError 186982

            lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice1, iGrid_ProdSolicSRV_Col), sProdSolicSRV1, iProdutoPreenchido1)
            If lErro <> SUCESSO Then gError 186983

            If sServicoFormatado1 = sServicoFormatado And sProdSolicSRV1 = sProdSolicSRV And _
            GridServicos.TextMatrix(iIndice, iGrid_Lote_Col) = GridServicos.TextMatrix(iIndice1, iGrid_Lote_Col) And _
            GridServicos.TextMatrix(iIndice, iGrid_FilialOP_Col) = GridServicos.TextMatrix(iIndice1, iGrid_FilialOP_Col) Then gError 186984

        Next

        For Each objProdSolicSRV In colProdSolicSRV
            If objProdSolicSRV.sServicoOrcSRV = sServicoFormatado Then
                objProdSolicSRV.dQuantidade = objProdSolicSRV.dQuantidade + StrParaDbl(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))
                iAchou = 1
                Exit For
            End If
        Next
        
        If iAchou = 0 Then
        
            Set objProdSolicSRV = New ClassProdSolicSRV
            
            objProdSolicSRV.sServicoOrcSRV = sServicoFormatado
            objProdSolicSRV.dQuantidade = StrParaDbl(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))
            
            colProdSolicSRV.Add objProdSolicSRV
            
        End If

    Next
    
'    For Each objProdSolicSRV In colProdSolicSRV
'
'        For Each objProdSolicSRVOrigem In gcolProdSolicSRVOrigem
'            If objProdSolicSRV.sServicoOrcSRV = objProdSolicSRVOrigem.sServicoOrcSRV Then
'                If objProdSolicSRV.dQuantidade > objProdSolicSRVOrigem.dQuantidade Then gError 186939
'                Exit For
'            End If
'        Next
'
'    Next
    
    Valida_Dados_Tela = SUCESSO

    Exit Function

Erro_Valida_Dados_Tela:

    Valida_Dados_Tela = gErr
    
    Select Case gErr
    
        Case 186935
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case 186936
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA_GRID1", gErr, iIndice)
        
        Case 186937
            Call Rotina_Erro(vbOKOnly, "ERRO_PROD_SOLICITADO_NAO_PREENCHIDO_GRID", gErr, iIndice)
            
        Case 186938
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_POSITIVA_GRID", gErr, iIndice)
            
        Case 186939
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_MAIOR_ORCADA", gErr, objProdSolicSRVOrigem.dQuantidade, objProdSolicSRV.sServicoOrcSRV, objProdSolicSRV.dQuantidade)
            
        Case 186940, 186980, 186982, 186983
            
        Case 186984
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_DUPLICADO_GRID", gErr, iIndice, iIndice1)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186941)

    End Select

End Function

Private Function Move_ProdSolicSRV_Memoria() As Long
'Move os dados da tela para objGarantia

Dim lErro As Long
Dim objProdSolicSRV As ClassProdSolicSRV
Dim sServicoFormatado As String
Dim iServicoPreenchido As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer

On Error GoTo Erro_Move_ProdSolicSRV_Memoria

    For iIndice = 1 To objGridServico.iLinhasExistentes
    
        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_Servico_Col), sServicoFormatado, iServicoPreenchido)
        If lErro <> SUCESSO Then gError 186945

        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_ProdSolicSRV_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 186946

        If iIndice > gcolProdSolicSRVDestino.Count Then gError 188144
        
        Set objProdSolicSRV = gcolProdSolicSRVDestino.Item(iIndice)
        
        objProdSolicSRV.sServicoOrcSRV = sServicoFormatado
        objProdSolicSRV.dQuantidade = StrParaDbl(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objProdSolicSRV.sProduto = sProdutoFormatado
        objProdSolicSRV.lGarantia = StrParaLong(GridServicos.TextMatrix(iIndice, iGrid_Garantia_Col))
        objProdSolicSRV.sContrato = GridServicos.TextMatrix(iIndice, iGrid_Contrato_Col)
        
        lErro = Move_RastroEstoque_Memoria(iIndice, objProdSolicSRV)
        If lErro <> SUCESSO Then gError 188127
        
    Next

    Move_ProdSolicSRV_Memoria = SUCESSO

    Exit Function

Erro_Move_ProdSolicSRV_Memoria:

    Move_ProdSolicSRV_Memoria = gErr

    Select Case gErr

        Case 186945, 186946, 188127

        Case 188144

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186947)

    End Select

    Exit Function

End Function

Private Function Move_RastroEstoque_Memoria(iLinha As Integer, objProdSolicSRV As ClassProdSolicSRV) As Long
'Move o Rastro dos Itens de Movimento

Dim objProduto As New ClassProduto, lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_RastroEstoque_Memoria
    
    lErro = CF("Produto_Formata", GridServicos.TextMatrix(iLinha, iGrid_ProdSolicSRV_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 188121
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 188122
    
        If lErro = 28030 Then gError 188123
        
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
            
            'Se colocou o Número do Lote
            If Len(Trim(GridServicos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
                objProdSolicSRV.sLote = GridServicos.TextMatrix(iLinha, iGrid_Lote_Col)
            End If
            
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
            
            'se o lote está preenchido e a filial não ==> erro
            If Len(Trim(GridServicos.TextMatrix(iLinha, iGrid_Lote_Col))) <> 0 Then
               
                If Len(Trim(GridServicos.TextMatrix(iLinha, iGrid_FilialOP_Col))) = 0 Then gError 188124
                
                objProdSolicSRV.sLote = GridServicos.TextMatrix(iLinha, iGrid_Lote_Col)
                objProdSolicSRV.iFilialOP = Codigo_Extrai(GridServicos.TextMatrix(iLinha, iGrid_FilialOP_Col))
                
            End If
                
            'se a filial está preenchida e o lote não ==> erro
            If Len(Trim(GridServicos.TextMatrix(iLinha, iGrid_FilialOP_Col))) <> 0 And _
               Len(Trim(GridServicos.TextMatrix(iLinha, iGrid_Lote_Col))) = 0 Then gError 188125
                
        End If
    
    End If
    
    Move_RastroEstoque_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_RastroEstoque_Memoria:

    Move_RastroEstoque_Memoria = gErr
    
    Select Case gErr
        
        Case 188121, 188122
        
        Case 188123
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 188124
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iLinha)
        
        Case 188125
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTO_NAO_PREENCHIDO", gErr, iLinha)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 188126)
    
    End Select
    
    Exit Function
    
End Function

Private Sub BotaoPecaServico_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim colProdutoSRV As Collection
Dim vServico As Integer
Dim objProdutoSRV As ClassProdutoSRV
Dim sServico As String
Dim iPreenchido As Integer
Dim sProduto As String
Dim objProdSolicSRV As New ClassProdSolicSRV

On Error GoTo Erro_BotaoPecaServico_Click
    
    If GridServicos.Row < 1 Or GridServicos.Row > objGridServico.iLinhasExistentes Then gError 188144
    
    lErro = Move_ProdSolicSRV_Memoria()
    If lErro <> SUCESSO Then gError 188145
        
    
'    lErro = CF("Produto_Formata", GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col), sServico, iPreenchido)
'    If lErro <> SUCESSO Then gError 188145
'
'    If iPreenchido = PRODUTO_VAZIO Then gError 188146
'
'    gcolProdSolicSRVDestino.Item(GridServicos.Row).sServicoOrcSRV = sServico
'
'    lErro = CF("Produto_Formata", GridServicos.TextMatrix(GridServicos.Row, iGrid_ProdSolicSRV_Col), sProduto, iPreenchido)
'    If lErro <> SUCESSO Then gError 188147
'
'    If iPreenchido = PRODUTO_VAZIO Then gError 188148
'
'    gcolProdSolicSRVDestino.Item(GridServicos.Row).sProduto = sProduto
'    gcolProdSolicSRVDestino.Item(GridServicos.Row).sLote = GridServicos.TextMatrix(GridServicos.Row, iGrid_Lote_Col)
'    gcolProdSolicSRVDestino.Item(GridServicos.Row).iFilialOP = StrParaInt(GridServicos.TextMatrix(GridServicos.Row, iGrid_FilialOP_Col))
    
    Call Chama_Tela("ProdutoSRV", GridServicos.Row, gcolProdSolicSRVDestino, Me)
    
    Exit Sub
    
Erro_BotaoPecaServico_Click:

    Select Case gErr

        Case 188144
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 188145, 188147

        Case 188148
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO_GRID1", gErr, GridServicos.Row)

        Case 188148
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr, GridServicos.Row)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188054)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name

        Case Servico.Name
            If gcolProdSolicSRVDestino.Count >= iLinha Then
                If gcolProdSolicSRVDestino.Item(iLinha).colProdutoSRV.Count > 0 Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
                End If
            Else
                objControl.Enabled = True
            End If
            
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188184)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is ProdSolicSRV Then

        sProduto1 = ProdSolicSRV.Text

    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridServicos.Row = 0 Then gError 188185

        sProduto1 = GridServicos.TextMatrix(GridServicos.Row, iGrid_ProdSolicSRV_Col)

    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 188186

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 188185
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 188186

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188187)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridServicos.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 188188

    ProdSolicSRV.Text = sProduto

    GridServicos.TextMatrix(GridServicos, iGrid_ProdSolicSRV_Col) = ProdSolicSRV.Text

    lErro = ProdSolicSRV_Saida_Celula()
    If lErro <> SUCESSO Then

        If Not (Me.ActiveControl Is ProdSolicSRV) Then
    
            GridServicos.TextMatrix(GridServicos.Row, iGrid_ProdSolicSRV_Col) = ""
    
        End If

        gError 188189
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 188188
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 188189

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188190)

    End Select

    Exit Sub

End Sub

Public Sub BotaoServicos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String
Dim sSelecaoSQL As String

On Error GoTo Erro_BotaoServicos_Click

    If Me.ActiveControl Is Servico Then

        sProduto1 = Servico.Text

    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridServicos.Row = 0 Then gError 188193

        sProduto1 = GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col)

    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 188194

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto

    Set colSelecao = New Collection

    colSelecao.Add NATUREZA_PROD_SERVICO

    sSelecaoSQL = "Natureza=?"

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoServico, sSelecaoSQL)

    Exit Sub

Erro_BotaoServicos_Click:

    Select Case gErr

        Case 188193
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 188194

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188195)

    End Select

    Exit Sub

End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridServicos.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 188196

    Servico.Text = sProduto

    GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col) = Servico.Text

    lErro = Servico_Saida_Celula()
    If lErro <> SUCESSO Then

        If Not (Me.ActiveControl Is Servico) Then

            GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col) = ""

        End If

        gError 188197
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 188196
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 188197

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188198)

    End Select

    Exit Sub

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Servico Then
            Call BotaoServicos_Click
        ElseIf Me.ActiveControl Is ProdSolicSRV Then
            Call BotaoProdutos_Click
        End If

    End If

End Sub

