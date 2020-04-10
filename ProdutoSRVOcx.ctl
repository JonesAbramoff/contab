VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ProdutoSRVOcx 
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   KeyPreview      =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   8085
   Begin VB.CommandButton BotaoPecas 
      Caption         =   "Peças"
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
      Left            =   210
      TabIndex        =   20
      Top             =   4800
      Width           =   1365
   End
   Begin VB.ComboBox Item 
      Height          =   315
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   285
      Width           =   780
   End
   Begin VB.TextBox DescricaoPeca 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   1650
      MaxLength       =   250
      TabIndex        =   16
      Top             =   3990
      Width           =   2490
   End
   Begin MSMask.MaskEdBox Peca 
      Height          =   225
      Left            =   450
      TabIndex        =   17
      Top             =   4005
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Contrato 
      Height          =   225
      Left            =   6450
      TabIndex        =   8
      Top             =   4035
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
      Left            =   5475
      TabIndex        =   9
      Top             =   4035
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   6030
      Picture         =   "ProdutoSRVOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   7050
      Picture         =   "ProdutoSRVOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   225
      Left            =   4200
      TabIndex        =   1
      Top             =   4005
      Width           =   1170
      _ExtentX        =   2064
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
   Begin MSFlexGridLib.MSFlexGrid GridServicos 
      Height          =   2520
      Left            =   180
      TabIndex        =   0
      Top             =   2160
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   4445
      _Version        =   393216
      Rows            =   6
      Cols            =   3
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Item:"
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
      Height          =   195
      Left            =   720
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   19
      Top             =   360
      Width           =   435
   End
   Begin VB.Label LabelDescProduto 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2550
      TabIndex        =   15
      Top             =   1170
      Width           =   3375
   End
   Begin VB.Label LabelDescServico 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2550
      TabIndex        =   14
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label LabelFilialOP 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4035
      TabIndex        =   13
      Top             =   1635
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "FilialOP:"
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
      Height          =   195
      Left            =   3210
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   12
      Top             =   1695
      Width           =   720
   End
   Begin VB.Label LabelLote 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1275
      TabIndex        =   11
      Top             =   1635
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Lote:"
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
      Height          =   195
      Left            =   705
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   10
      Top             =   1695
      Width           =   450
   End
   Begin VB.Label LabelProduto 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1275
      TabIndex        =   7
      Top             =   1170
      Width           =   1245
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   420
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   6
      Top             =   1230
      Width           =   735
   End
   Begin VB.Label LabelServico 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1275
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Serviço:"
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
      Height          =   195
      Left            =   435
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   4
      Top             =   780
      Width           =   720
   End
End
Attribute VB_Name = "ProdutoSRVOcx"
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


Dim gcolProdSolicSRV As Collection
Dim objGridServico As AdmGrid
Dim giItem As Integer
Dim gobjTela As Object


Dim iGrid_Peca_Col As Integer
Dim iGrid_DescPeca_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Garantia_Col As Integer
Dim iGrid_Contrato_Col As Integer

Dim giFrameAtual As Integer

Private WithEvents objEventoPeca As AdmEvento
Attribute objEventoPeca.VB_VarHelpID = -1

Public Function Trata_Parametros(ByVal iItem As Integer, colProdSolicSRV As Collection, objTela As Object) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    objTela.Enabled = False
    Set gobjTela = objTela


    Set gcolProdSolicSRV = colProdSolicSRV

    For iIndice = 1 To colProdSolicSRV.Count
        Item.AddItem (iIndice)
    Next

    Item.ListIndex = iItem - 1

    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188179)
    
    End Select
    
    Exit Function
    
End Function

Private Function Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Peca)
    If lErro <> SUCESSO Then gError 188239
    
    Set objGridServico = New AdmGrid
    
    Call Inicializa_Grid_Servico(objGridServico)
    
    Set objEventoPeca = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Function
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
        
        Case 188239
        
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
    objGridInt.colColuna.Add ("Peça")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Garantia")
    objGridInt.colColuna.Add ("Contrato")

    objGridInt.colCampo.Add (Peca.Name)
    objGridInt.colCampo.Add (DescricaoPeca.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Garantia.Name)
    objGridInt.colCampo.Add (Contrato.Name)

    'Controles que participam do Grid
    iGrid_Peca_Col = 1
    iGrid_DescPeca_Col = 2
    iGrid_Quantidade_Col = 3
    iGrid_Garantia_Col = 4
    iGrid_Contrato_Col = 5

    objGridInt.objGrid = GridServicos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PRODUTOSRV + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Largura da primeira coluna
    GridServicos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Servico = SUCESSO

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoPeca = Nothing
    
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
    Caption = "Peça x Serviço"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "ProdutoSRV"
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

Private Sub Item_Click()

Dim lErro As Long
Dim iItem As Integer
Dim iIndice As Integer

On Error GoTo Erro_Item_Click

    iItem = StrParaInt(Item.Text)

    If giItem <> iItem Then
        
        lErro = Valida_Dados_Tela()
        If lErro <> SUCESSO Then gError 188180
    
        'Move os dados da tela para o objRelacionamentoClie
        lErro = Move_ProdutoSRV_Memoria()
        If lErro <> SUCESSO Then gError 188181
    
        lErro = Traz_ProdutoSRV_Tela(iItem)
        If lErro <> SUCESSO Then gError 188182
    
    End If
    
    Exit Sub
    
Erro_Item_Click:

    Select Case gErr
    
        Case 188180, 188181
            For iIndice = 0 To Item.ListCount - 1
                If Item.List(iIndice) = giItem Then
                    Item.ListIndex = iIndice
                    Exit For
                End If
            Next
            
        Case 188182
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 188183)

    End Select
        
    Exit Sub

End Sub

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

    Call Grid_Trata_Tecla1(KeyCode, objGridServico)

End Sub

Public Sub Peca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Peca_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub Peca_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub Peca_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = Peca
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

            'Se for a de ProdutoSRV
            Case iGrid_Peca_Col
                lErro = Saida_Celula_Peca(objGridInt)
                If lErro <> SUCESSO Then gError 188014

            'Se for a de Quantidade
            Case iGrid_Quantidade_Col
                lErro = Saida_Celula_Quantidade(objGridInt)
                If lErro <> SUCESSO Then gError 188015
        
            'Se for a de Garantia
            Case iGrid_Garantia_Col
                lErro = Saida_Celula_Garantia(objGridInt)
                If lErro <> SUCESSO Then gError 188016
    
            'Se for a de Contrato
            Case iGrid_Contrato_Col
                lErro = Saida_Celula_Contrato(objGridInt)
                If lErro <> SUCESSO Then gError 188154
        
        End Select


    End If
        

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188017
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 188014 To 188017, 188154

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188018)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Peca(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Peca

    Set objGridInt.objControle = Peca

    lErro = Peca_Saida_Celula()
    If lErro <> SUCESSO Then gError 188236

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188019

    If Len(Trim(Peca.ClipText)) <> 0 Then

        If GridServicos.Row - GridServicos.FixedRows = objGridServico.iLinhasExistentes Then
            
            objGridServico.iLinhasExistentes = objGridServico.iLinhasExistentes + 1
    
        End If
    
    End If

    Saida_Celula_Peca = SUCESSO

    Exit Function

Erro_Saida_Celula_Peca:

    Saida_Celula_Peca = gErr

    Select Case gErr

        Case 188019, 188236
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 188158)

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
        If lErro <> SUCESSO Then gError 188021

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188022
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 188021, 188022
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188023)

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
        If lErro <> SUCESSO Then gError 188159

        objGarantia.iFilialEmpresa = giFilialEmpresa
        objGarantia.lCodigo = StrParaLong(Garantia.Text)
        objGarantia.sProduto = LabelProduto.Caption
        objGarantia.sServico = GridServicos.TextMatrix(GridServicos.Row, iGrid_Peca_Col)
        objGarantia.sLote = LabelLote.Caption
        objGarantia.iFilialOP = LabelFilialOP.Caption

        lErro = CF("Testa_Garantia", objGarantia)
        If lErro <> SUCESSO Then gError 188160
        
    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188161
    
    Saida_Celula_Garantia = SUCESSO

    Exit Function

Erro_Saida_Celula_Garantia:

    Saida_Celula_Garantia = gErr

    Select Case gErr

        Case 188159 To 188161
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188162)

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

        objItensDeContratoSrv.iFilialEmpresa = giFilialEmpresa
        objItensDeContratoSrv.sCodigoContrato = Contrato.Text
        objItensDeContratoSrv.sProduto = LabelProduto.Caption
        objItensDeContratoSrv.sServico = GridServicos.TextMatrix(GridServicos.Row, iGrid_Peca_Col)
        objItensDeContratoSrv.sLote = LabelLote.Caption
        objItensDeContratoSrv.iFilialOP = LabelFilialOP.Caption
        
        lErro = CF("Testa_Contrato", objItensDeContratoSrv)
        If lErro <> SUCESSO Then gError 188164

    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188165
    
    Saida_Celula_Contrato = SUCESSO

    Exit Function

Erro_Saida_Celula_Contrato:

    Saida_Celula_Contrato = gErr

    Select Case gErr

        Case 188163 To 188165
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188166)

    End Select

    Exit Function

End Function

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim objGarantia As New ClassGarantia

On Error GoTo Erro_BotaoOK_Click

    lErro = Valida_Dados_Tela()
    If lErro <> SUCESSO Then gError 188026

    'Move os dados da tela para o objRelacionamentoClie
    lErro = Move_ProdutoSRV_Memoria()
    If lErro <> SUCESSO Then gError 188027

    iAlterado = 0

    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case 188026, 188027
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188028)

    End Select

    Exit Sub

End Sub

Private Function Valida_Dados_Tela() As Long
'Verifica se os dados da tela são válidos

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sPecaFormatada As String
Dim iPecaPreenchida As Integer
Dim sPecaFormatada1 As String
Dim iPecaPreenchida1 As Integer

On Error GoTo Erro_Valida_Dados_Tela

    For iIndice = 1 To objGridServico.iLinhasExistentes

        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Peca_Col))) = 0 Then gError 188029
        
        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 188030
        
        If StrParaDbl(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col)) <= 0 Then gError 188031
        
        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_Peca_Col), sPecaFormatada, iPecaPreenchida)
        If lErro <> SUCESSO Then gError 188033


        For iIndice1 = iIndice + 1 To objGridServico.iLinhasExistentes

            lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice1, iGrid_Peca_Col), sPecaFormatada1, iPecaPreenchida1)
            If lErro <> SUCESSO Then gError 188035

            If sPecaFormatada1 = sPecaFormatada Then gError 188037

        Next

    Next
    
    Valida_Dados_Tela = SUCESSO

    Exit Function

Erro_Valida_Dados_Tela:

    Valida_Dados_Tela = gErr
    
    Select Case gErr
    
        Case 188029
            Call Rotina_Erro(vbOKOnly, "ERRO_PECA_NAO_PREENCHIDA_GRID", gErr, iIndice)

        Case 188030
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA_GRID1", gErr, iIndice)
        
        Case 188031
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_POSITIVA_GRID", gErr, iIndice)
            
        Case 188033, 188035
            
        Case 188037
            Call Rotina_Erro(vbOKOnly, "ERRO_PECA_SERVICO_DUPLICADO_GRID", gErr, iIndice, iIndice1)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188039)

    End Select

End Function

Private Function Move_ProdutoSRV_Memoria() As Long
'Move os dados da tela para objGarantia

Dim lErro As Long
Dim objProdutoSRV As ClassProdutoSRV
Dim sServicoFormatado As String
Dim iServicoPreenchido As Integer
Dim sPecaFormatada As String
Dim iPecaPreenchida As Integer
Dim iIndice As Integer
Dim colProdutoSRV As Collection

On Error GoTo Erro_Move_ProdutoSRV_Memoria

    If giItem <> 0 Then

        Set colProdutoSRV = gcolProdSolicSRV(giItem).colProdutoSRV
    
        For iIndice = colProdutoSRV.Count To 1 Step -1
            colProdutoSRV.Remove (iIndice)
        Next
        
        For iIndice = 1 To objGridServico.iLinhasExistentes
        
            lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_Peca_Col), sPecaFormatada, iPecaPreenchida)
            If lErro <> SUCESSO Then gError 188040
    
            Set objProdutoSRV = New ClassProdutoSRV
            
            objProdutoSRV.sProduto = sPecaFormatada
            objProdutoSRV.dQuantidade = StrParaDbl(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))
            objProdutoSRV.lGarantia = StrParaLong(GridServicos.TextMatrix(iIndice, iGrid_Garantia_Col))
            objProdutoSRV.sContrato = GridServicos.TextMatrix(iIndice, iGrid_Contrato_Col)
            
            colProdutoSRV.Add objProdutoSRV
    
        Next

    End If

    Move_ProdutoSRV_Memoria = SUCESSO

    Exit Function

Erro_Move_ProdutoSRV_Memoria:

    Move_ProdutoSRV_Memoria = gErr

    Select Case gErr

        Case 188040

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188042)

    End Select

    Exit Function

End Function

Private Function Traz_ProdutoSRV_Tela(ByVal iItem As Integer) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long
Dim sProduto As String
Dim objProdutoSRV As ClassProdutoSRV
Dim iIndice As Integer
Dim sServico As String
Dim objProduto As New ClassProduto
Dim sPeca As String
Dim objProdSolicSRV As ClassProdSolicSRV
Dim colProdutoSRV As Collection

On Error GoTo Erro_Traz_ProdutoSRV_Tela

    Call Grid_Limpa(objGridServico)

    Set objProdSolicSRV = gcolProdSolicSRV(iItem)
    Set colProdutoSRV = objProdSolicSRV.colProdutoSRV
    giItem = iItem

    objProduto.sCodigo = objProdSolicSRV.sServicoOrcSRV
        
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 188145
        
    If lErro = 28030 Then gError 188146

    lErro = Mascara_RetornaProdutoTela(objProdSolicSRV.sServicoOrcSRV, sServico)
    If lErro <> SUCESSO Then gError 188147

    LabelServico.Caption = sServico
    LabelDescServico.Caption = objProduto.sDescricao
    
    objProduto.sCodigo = objProdSolicSRV.sProduto
        
    'Lê os demais atributos do Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 188148
        
    If lErro = 28030 Then gError 188149
    
    lErro = Mascara_RetornaProdutoTela(objProdSolicSRV.sProduto, sProduto)
    If lErro <> SUCESSO Then gError 188150
    
    LabelProduto.Caption = sProduto
    LabelDescProduto.Caption = objProduto.sDescricao
    
    LabelLote.Caption = objProdSolicSRV.sLote
    
    If objProdSolicSRV.iFilialOP > 0 Then
        LabelFilialOP.Caption = objProdSolicSRV.iFilialOP
    End If
    
    For iIndice = 1 To colProdutoSRV.Count
    
        Set objProdutoSRV = colProdutoSRV(iIndice)
            
        objProduto.sCodigo = objProdutoSRV.sProduto
            
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 188151
            
        If lErro = 28030 Then gError 188152
            
        lErro = Mascara_RetornaProdutoTela(objProdutoSRV.sProduto, sPeca)
        If lErro <> SUCESSO Then gError 188153
            
        GridServicos.TextMatrix(iIndice, iGrid_Peca_Col) = sPeca
        GridServicos.TextMatrix(iIndice, iGrid_DescPeca_Col) = objProduto.sDescricao
        GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objProdutoSRV.dQuantidade)
        If objProdutoSRV.lGarantia > 0 Then
            GridServicos.TextMatrix(iIndice, iGrid_Garantia_Col) = objProdutoSRV.lGarantia
        End If
        
        GridServicos.TextMatrix(iIndice, iGrid_Contrato_Col) = objProdutoSRV.sContrato
    
    Next
    
    'Atualiza o número de linhas existentes
    objGridServico.iLinhasExistentes = colProdutoSRV.Count
    
    iAlterado = 0
    
    Traz_ProdutoSRV_Tela = SUCESSO
    
    Exit Function

Erro_Traz_ProdutoSRV_Tela:

    Traz_ProdutoSRV_Tela = gErr
    
    Select Case gErr
    
        Case 188145, 188147, 188148, 188150, 188151, 188153
        
        Case 188146, 188149, 188152
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188013)
    
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoPecas_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String
Dim sSelecaoSQL As String

On Error GoTo Erro_BotaoPecas_Click

    If Me.ActiveControl Is Peca Then

        sProduto1 = Peca.Text

    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridServicos.Row = 0 Then gError 188201

        sProduto1 = GridServicos.TextMatrix(GridServicos.Row, iGrid_Peca_Col)

    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 188202

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto

    Set colSelecao = New Collection

    colSelecao.Add NATUREZA_PROD_SERVICO

    sSelecaoSQL = "Natureza<>?"

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoPeca, sSelecaoSQL)

    Exit Sub

Erro_BotaoPecas_Click:

    Select Case gErr

        Case 188201
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 188202

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188203)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPeca_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoPeca_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridServicos.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 188204

    Peca.PromptInclude = False
    Peca.Text = sProduto
    Peca.PromptInclude = True

    GridServicos.TextMatrix(GridServicos.Row, iGrid_Peca_Col) = Peca.Text

    lErro = Peca_Saida_Celula()
    If lErro <> SUCESSO Then

        If Not (Me.ActiveControl Is Peca) Then

            GridServicos.TextMatrix(GridServicos.Row, iGrid_Peca_Col) = ""

        End If

        gError 188205
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoPeca_evSelecao:

    Select Case gErr

        Case 188204
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 188205

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188206)

    End Select

    Exit Sub

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Peca Then
            Call BotaoPecas_Click
        End If

    End If

End Sub

Private Function Peca_Saida_Celula() As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String

On Error GoTo Erro_Peca_Saida_Celula

    If Len(Trim(Peca.ClipText)) > 0 Then

        'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", Peca.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 188167
    
        If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then
            gError 188168
        End If
    
        If objProduto.iNatureza = NATUREZA_PROD_SERVICO Then gError 188169
    
        'Se o produto não foi encontrado ==> Pergunta se deseja criar
        If lErro = 51381 Then gError 188170
    
        GridServicos.TextMatrix(GridServicos.Row, iGrid_DescPeca_Col) = objProduto.sDescricao
    
    End If
    
    Peca_Saida_Celula = SUCESSO

    Exit Function

Erro_Peca_Saida_Celula:

    Peca_Saida_Celula = gErr

    Select Case gErr

        Case 188167, 188168

        Case 188169
            Call Rotina_Erro(vbOKOnly, "ERRO_PECA_NAO_PODE_TER_NATUREZA_SERVICO", gErr, objProduto.sCodigo)

        Case 188170
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 188177, 188178

        Case 188179
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 188192)

    End Select

    Exit Function

End Function

