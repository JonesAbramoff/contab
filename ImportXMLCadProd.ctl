VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ImportXMLCadProd 
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   KeyPreview      =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   10095
   Begin VB.CommandButton BotaoOK 
      Caption         =   "Gravar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4080
      Picture         =   "ImportXMLCadProd.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6210
      Width           =   885
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5070
      Picture         =   "ImportXMLCadProd.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6210
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   6105
      Index           =   2
      Left            =   60
      TabIndex        =   3
      Top             =   75
      Width           =   9975
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1920
         Picture         =   "ImportXMLCadProd.ctx":025C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5505
         Width           =   1710
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   75
         Picture         =   "ImportXMLCadProd.ctx":143E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5505
         Width           =   1710
      End
      Begin VB.CheckBox Selecionado 
         Height          =   210
         Left            =   645
         TabIndex        =   15
         Top             =   4875
         Width           =   675
      End
      Begin VB.TextBox NomeRedProd 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   6810
         MaxLength       =   50
         TabIndex        =   14
         Top             =   4200
         Width           =   1845
      End
      Begin VB.ComboBox TipoProd 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   510
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4125
         Width           =   2370
      End
      Begin VB.ComboBox ClasseUM 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8025
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4515
         Width           =   1185
      End
      Begin MSMask.MaskEdBox OrigemMerc 
         Height          =   225
         Left            =   6090
         TabIndex        =   11
         Top             =   4155
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.ComboBox UnidadeMed 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9150
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   4500
         Width           =   660
      End
      Begin VB.TextBox NCMProd 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   9
         Top             =   4200
         Width           =   1260
      End
      Begin VB.TextBox EANProd 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2895
         MaxLength       =   50
         TabIndex        =   8
         Top             =   4200
         Width           =   1845
      End
      Begin VB.TextBox DescProd 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   3735
         MaxLength       =   50
         TabIndex        =   6
         Top             =   4560
         Width           =   3510
      End
      Begin MSMask.MaskEdBox ProdutoXml 
         Height          =   225
         Left            =   510
         TabIndex        =   7
         Top             =   4575
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.TextBox UnidadeMedXml 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   7110
         MaxLength       =   50
         TabIndex        =   5
         Top             =   4560
         Width           =   945
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   225
         Left            =   2115
         TabIndex        =   4
         Top             =   4560
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   5295
         Left            =   60
         TabIndex        =   0
         Top             =   195
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   9340
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "ImportXMLCadProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim objGridItens As AdmGrid
Dim iGrid_Sel_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_ProdutoXml_Col As Integer
Dim iGrid_DescrProd_Col As Integer
Dim iGrid_TipoProd_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_UnidadeMedXml_Col As Integer
Dim iGrid_ClasseUM_Col As Integer
Dim iGrid_EANProd_Col As Integer
Dim iGrid_NCMProd_Col As Integer
Dim iGrid_NomeRedProd_Col As Integer
Dim iGrid_OrigMerc_Col As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cadastro de Produtos por Importação de XML de NFe/CTe"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ImportXMLCadProd"

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

Private Sub BotaoCancela_Click()
    Unload Me
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

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridItens = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213800)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Inicializa a Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Carrega_TiposProd()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Carrega_ClassesUM()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    DescProd.MaxLength = STRING_PRODUTO_DESCRICAO_TELA
    
    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213801)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long
Dim colProd As New Collection

On Error GoTo Erro_Trata_Parametros

    lErro = CF("ImportXMLCadProd_Le", colProd)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If colProd.Count = 0 Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_SEM_PRODUTOS_NAO_CADASTRADOS")
        'Call BotaoCancela_Click
    Else
        lErro = Traz_Produtos_Tela(colProd)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel
    
    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213802)

    End Select

    Exit Function

End Function

Private Function Traz_Produtos_Tela(ByVal colProdutos As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto, bComMask As Boolean
Dim sProdutoEnxuto As String, sProduto As String
Dim sProdutoFormatado As String, iProdutoPreenchido As Integer

On Error GoTo Erro_Traz_Produtos_Tela

    iIndice = 0
    
    If colProdutos.Count >= objGridItens.objGrid.Rows Then
        Call Refaz_Grid(objGridItens, colProdutos.Count)
    End If

    'Para cada ítem
    For Each objProduto In colProdutos

        iIndice = iIndice + 1
        
        sProduto = objProduto.sCodigo
        
        If sProduto <> "" Then
        
            lErro = Produto_Ajusta_Formato(sProduto, bComMask)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If bComMask Then
                lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido, True)
                If lErro <> SUCESSO Then sProdutoEnxuto = ""
            Else
                lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
        Else
            sProdutoEnxuto = ""
        End If

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
               
        GridItens.TextMatrix(iIndice, iGrid_Sel_Col) = S_MARCADO
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_ProdutoXml_Col) = objProduto.sReferencia
        GridItens.TextMatrix(iIndice, iGrid_DescrProd_Col) = objProduto.sDescricao
        GridItens.TextMatrix(iIndice, iGrid_NomeRedProd_Col) = objProduto.sNomeReduzido
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMedXml_Col) = objProduto.sSiglaUMTrib
        GridItens.TextMatrix(iIndice, iGrid_NCMProd_Col) = objProduto.sIPICodigo
        GridItens.TextMatrix(iIndice, iGrid_EANProd_Col) = objProduto.sCodigoBarras
        GridItens.TextMatrix(iIndice, iGrid_OrigMerc_Col) = objProduto.iOrigemMercadoria
        
        Call Combo_Seleciona_ItemData(ClasseUM, objProduto.iClasseUM)
        GridItens.TextMatrix(iIndice, iGrid_ClasseUM_Col) = ClasseUM.Text
        
        Call Combo_Seleciona_ItemData(TipoProd, objProduto.iTipo)
        GridItens.TextMatrix(iIndice, iGrid_TipoProd_Col) = TipoProd.Text
        
    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice
    
    Call Grid_Refresh_Checkbox(objGridItens)

    Traz_Produtos_Tela = SUCESSO

    Exit Function

Erro_Traz_Produtos_Tela:
   
    Traz_Produtos_Tela = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213803)

    End Select

    Exit Function
    
End Function

Private Function Move_Tela_Memoria(ByVal colProdutos As Collection) As Long

Dim lErro As Long, iIndice As Integer
Dim sProdutoFormatado As String, iProdutoPreenchido As Integer
Dim objProduto As ClassProduto
Dim objTipoProd As ClassTipoDeProduto
Dim objProdutoBD As ClassProduto
Dim objProdutoAux As ClassProduto, iLinhaAux As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        If GridItens.TextMatrix(iIndice, iGrid_Sel_Col) = S_MARCADO Then
            
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If iProdutoPreenchido = PRODUTO_VAZIO Then gError 213804
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 213805
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_NomeRedProd_Col))) = 0 Then gError 213806
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_DescrProd_Col))) = 0 Then gError 213807
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_TipoProd_Col))) = 0 Then gError 213808
            
            iLinhaAux = 0
            For Each objProdutoAux In colProdutos
                iLinhaAux = iLinhaAux + 1
                If objProdutoAux.sCodigo = sProdutoFormatado Then gError 213812
            Next
            
            Set objProduto = New ClassProduto
            Set objTipoProd = New ClassTipoDeProduto
          
            objProduto.iAtivo = PRODUTO_ATIVO
            objProduto.sCodigo = sProdutoFormatado
            objProduto.sReferencia = Trim(objProduto.sCodigo)
            objProduto.sDescricao = GridItens.TextMatrix(iIndice, iGrid_DescrProd_Col)
            objProduto.sNomeReduzido = GridItens.TextMatrix(iIndice, iGrid_NomeRedProd_Col)
            objProduto.sCodigoBarras = GridItens.TextMatrix(iIndice, iGrid_EANProd_Col)
            objProduto.iTipo = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_TipoProd_Col))
            objProduto.iClasseUM = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_ClasseUM_Col))
            objProduto.sSiglaUMVenda = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
            objProduto.sIPICodigo = GridItens.TextMatrix(iIndice, iGrid_NCMProd_Col)
            objProduto.iOrigemMercadoria = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_OrigMerc_Col))
            
            objProduto.colCodBarras.Add objProduto.sCodigoBarras
            
            objTipoProd.iTipo = objProduto.iTipo
        
            lErro = CF("TipoDeProduto_Le", objTipoProd)
            If lErro <> SUCESSO And lErro <> 22531 Then gError ERRO_SEM_MENSAGEM
            
            objProduto.iControleEstoque = objTipoProd.iControleEstoque
            objProduto.iNatureza = objTipoProd.iNatureza
            objProduto.iFaturamento = objTipoProd.iFaturamento
            objProduto.iCompras = objTipoProd.iCompras
            objProduto.iApropriacaoCusto = objTipoProd.iApropriacaoCusto
            objProduto.sSiglaUMCompra = objProduto.sSiglaUMVenda
            objProduto.sSiglaUMEstoque = objProduto.sSiglaUMVenda
            objProduto.sSiglaUMTrib = objProduto.sSiglaUMVenda
            objProduto.iConsideraQuantCotAnt = objTipoProd.iConsideraQuantCotAnt
            objProduto.dPercentMenosQuantCotAnt = objTipoProd.dPercentMenosQuantCotAnt
            objProduto.dPercentMaisQuantCotAnt = objTipoProd.dPercentMaisQuantCotAnt
            objProduto.iTemFaixaReceb = objTipoProd.iTemFaixaReceb
            objProduto.dPercentMaisReceb = objTipoProd.dPercentMaisReceb
            objProduto.dPercentMenosReceb = objTipoProd.dPercentMenosReceb
            objProduto.iRecebForaFaixa = objTipoProd.iRecebForaFaixa
            
            Set objProdutoBD = New ClassProduto
            
            objProdutoBD.sCodigo = objProduto.sCodigo
            
            lErro = CF("Produto_Le", objProdutoBD)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            If lErro = SUCESSO Then gError 213809
            
            Set objProdutoBD = New ClassProduto
            
            objProdutoBD.sNomeReduzido = objProduto.sNomeReduzido
            
            lErro = CF("Produto_Le_NomeReduzido", objProdutoBD)
            If lErro <> SUCESSO And lErro <> 26927 Then gError ERRO_SEM_MENSAGEM
        
            If lErro = SUCESSO Then gError 213810
        
            colProdutos.Add objProduto
        End If

    Next
    
    If colProdutos.Count = 0 Then gError 213826
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:
   
    Move_Tela_Memoria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 213804
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_LINHA", gErr, iIndice)

        Case 213805
             Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA_LINHA", gErr, iIndice)

        Case 213806
             Call Rotina_Erro(vbOKOnly, "ERRO_NOMEREDPROD_NAO_PREENCHIDO_LINHA", gErr, iIndice)

        Case 213807
             Call Rotina_Erro(vbOKOnly, "ERRO_DESCPROD_NAO_PREENCHIDA_LINHA", gErr, iIndice)

        Case 213808
             Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPROD_NAO_PREENCHIDO_LINHA", gErr, iIndice)

        Case 213809
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_CODIGO_JA_EXISTE", gErr, objProduto.sCodigo, iIndice)

        Case 213810
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NOMERED_JA_EXISTE", gErr, objProduto.sCodigo, iIndice, objProdutoBD.sCodigo)

        Case 213812
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_REPETIDO_NO_GRID", gErr, objProduto.sCodigo, iIndice)

        Case 213826
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_SELECIONADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213811)

    End Select

    Exit Function
    
End Function

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add (" ")
    objGrid.colColuna.Add ("Produto XML")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Nome Red.")
    objGrid.colColuna.Add ("Tipo")
    objGrid.colColuna.Add ("U.M. XML")
    objGrid.colColuna.Add ("Classe UM")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("EAN")
    objGrid.colColuna.Add ("NCM")
    objGrid.colColuna.Add ("Origem")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Selecionado.Name)
    objGrid.colCampo.Add (ProdutoXml.Name)
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (DescProd.Name)
    objGrid.colCampo.Add (NomeRedProd.Name)
    objGrid.colCampo.Add (TipoProd.Name)
    objGrid.colCampo.Add (UnidadeMedXml.Name)
    objGrid.colCampo.Add (ClasseUM.Name)
    objGrid.colCampo.Add (UnidadeMed.Name)
    objGrid.colCampo.Add (EANProd.Name)
    objGrid.colCampo.Add (NCMProd.Name)
    objGrid.colCampo.Add (OrigemMerc.Name)

    'Colunas do Grid
    iGrid_Sel_Col = 1
    iGrid_ProdutoXml_Col = 2
    iGrid_Produto_Col = 3
    iGrid_DescrProd_Col = 4
    iGrid_NomeRedProd_Col = 5
    iGrid_TipoProd_Col = 6
    iGrid_UnidadeMedXml_Col = 7
    iGrid_ClasseUM_Col = 8
    iGrid_UnidadeMed_Col = 9
    iGrid_EANProd_Col = 10
    iGrid_NCMProd_Col = 11
    iGrid_OrigMerc_Col = 12

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_NF + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 14

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

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

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Private Sub GridItens_LostFocus()
    Call Grid_Libera_Foco(objGridItens)
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

Private Sub UnidadeMed_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UnidadeMed_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescProd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DescProd_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub DescProd_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub DescProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescProd
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoProd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoProd_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub TipoProd_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub TipoProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TipoProd
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NomeRedProd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NomeRedProd_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub NomeRedProd_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub NomeRedProd_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = NomeRedProd
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ClasseUM_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ClasseUM_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub ClasseUM_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub ClasseUM_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ClasseUM
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'GridItensNF
        If objGridInt.objGrid.Name = GridItens.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_Produto_Col
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                Case iGrid_TipoProd_Col
                    lErro = Saida_Celula_Padrao(objGridInt, TipoProd)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Case iGrid_ClasseUM_Col
                    lErro = Saida_Celula_ClasseUM(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                Case iGrid_DescrProd_Col
                    lErro = Saida_Celula_Padrao(objGridInt, DescProd)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                Case iGrid_NomeRedProd_Col
                    lErro = Saida_Celula_NomeRedProd(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            End Select
                    
        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 213813

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 213813
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213814)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim iTipo As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name
        
        Case Produto.Name
            objControl.Enabled = True
            
        'Unidade de Medida
        Case UnidadeMed.Name

            UnidadeMed.Clear

            'Guarda a UM que está no Grid
            sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

            If Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_ClasseUM_Col)) = 0 Then
                UnidadeMed.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objClasseUM.iClasse = Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_ClasseUM_Col))
                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next
                'Seleciona na UM que está preenchida
                If sUM = "" Then
                    UnidadeMed.ListIndex = -1
                Else
                    UnidadeMed.Text = sUM
                    If Len(Trim(sUM)) > 0 Then
                        lErro = Combo_Item_Igual(UnidadeMed)
                        If lErro <> SUCESSO And lErro <> 12253 Then gError ERRO_SEM_MENSAGEM
                    End If
                End If
            End If
            
        Case DescProd.Name, NomeRedProd.Name, ClasseUM.Name, TipoProd.Name, Selecionado.Name
            objControl.Enabled = True
        
        Case Else
            objControl.Enabled = False

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213815)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

Dim lErro As Long
Dim objProdutoBD As New ClassProduto
Dim sProduto As String, sProdutoFormatado As String, iPreenchido As Integer

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) > 0 Then
    
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objProdutoBD.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProdutoBD)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then gError 213809

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 213809
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_CODIGO_JA_EXISTE", gErr, objProdutoBD.sCodigo, GridItens.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213816)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213817)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_ClasseUM(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iClasseUMAnt As Integer
Dim objClasseUM As New ClassClasseUM
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_ClasseUM

    Set objGridInt.objControle = ClasseUM
    
    objClasseUM.iClasse = Codigo_Extrai(ClasseUM)
    
    iClasseUMAnt = Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_ClasseUM_Col))

    'Verificar se é uma classe cadastrada em ClasseUM
    lErro = CF("ClasseUM_Le", objClasseUM)
    If lErro <> SUCESSO And lErro <> 22537 Then gError ERRO_SEM_MENSAGEM

    If lErro = 22537 Then gError 213818
    
    If iClasseUMAnt <> objClasseUM.iClasse Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = ""
    
        lErro = Carrega_CombosUM(objClasseUM)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_ClasseUM = SUCESSO

    Exit Function

Erro_Saida_Celula_ClasseUM:

    Saida_Celula_ClasseUM = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 213818
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLASSEUM", objClasseUM.iClasse)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("ClasseUM", objClasseUM)
            End If
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213819)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_NomeRedProd(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objProdutoBD As New ClassProduto

On Error GoTo Erro_Saida_Celula_NomeRedProd

    Set objGridInt.objControle = NomeRedProd
    
    objProdutoBD.sNomeReduzido = NomeRedProd.Text
    
    lErro = CF("Produto_Le_NomeReduzido", objProdutoBD)
    If lErro <> SUCESSO And lErro <> 26927 Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then gError 213810
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_NomeRedProd = SUCESSO

    Exit Function

Erro_Saida_Celula_NomeRedProd:

    Saida_Celula_NomeRedProd = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 213810
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NOMERED_JA_EXISTE", gErr, GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, objProdutoBD.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213820)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim colProdutos As New Collection

On Error GoTo Erro_BotaoOK_Click

    lErro = Move_Tela_Memoria(colProdutos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("ImportXMLCadProd_Grava", colProdutos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213821)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Selecionado_Change()
    iAlterado = REGISTRO_ALTERADO
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

Private Function Carrega_CombosUM(objClasseUM As ClassClasseUM) As Long
'Carrega as combos de Unidades de Medida de acordo com a ClasseUM passada

Dim lErro As Long
Dim colSiglas As New Collection
Dim iIndice As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_Carrega_CombosUM

    'Lê as U.M. da Classe passada
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
    If lErro <> SUCESSO And lErro <> 22539 Then gError ERRO_SEM_MENSAGEM
    
    'Carrega as combos
    If lErro = SUCESSO Then

        For iIndice = 1 To colSiglas.Count
            UnidadeMed.AddItem colSiglas.Item(iIndice).sSigla
            If colSiglas.Item(iIndice).sSigla = objClasseUM.sSiglaUMBase Then iIndice2 = iIndice
        Next

        UnidadeMed.ListIndex = iIndice2 - 1

    End If

    Carrega_CombosUM = SUCESSO

    Exit Function

Erro_Carrega_CombosUM:

    Carrega_CombosUM = Err

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213822)

    End Select

    Exit Function

End Function

Private Function Carrega_TiposProd() As Long

Dim lErro As Long
Dim objTipo As ClassTipoDeProduto
Dim colTipos As New Collection

On Error GoTo Erro_Carrega_TiposProd

    'Lê as U.M. da Classe passada
    lErro = CF("TipoDeProduto_Le_Todos", colTipos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objTipo In colTipos
        TipoProd.AddItem CStr(objTipo.iTipo) & SEPARADOR & objTipo.sDescricao
        TipoProd.ItemData(TipoProd.NewIndex) = objTipo.iTipo
    Next

    Carrega_TiposProd = SUCESSO

    Exit Function

Erro_Carrega_TiposProd:

    Carrega_TiposProd = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213823)

    End Select

    Exit Function

End Function

Private Function Carrega_ClassesUM() As Long

Dim lErro As Long
Dim objClasseUM As ClassClasseUM
Dim colClasses As New Collection

On Error GoTo Erro_Carrega_ClassesUM

    'Lê as U.M. da Classe passada
    lErro = CF("ClasseUM_Le_Todas", colClasses)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objClasseUM In colClasses
        ClasseUM.AddItem CStr(objClasseUM.iClasse) & SEPARADOR & objClasseUM.sDescricao
        ClasseUM.ItemData(ClasseUM.NewIndex) = objClasseUM.iClasse
    Next

    Carrega_ClassesUM = SUCESSO

    Exit Function

Erro_Carrega_ClassesUM:

    Carrega_ClassesUM = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213824)

    End Select

    Exit Function

End Function

Function Produto_Ajusta_Formato(sProduto As String, Optional bComMask As Boolean) As Long

Dim lErro As Long
Dim sProdutoNovo As String
Dim iSeg As Integer, iNumSeg As Integer
Dim objsegmento As New ClassSegmento, sSeg As String
Dim colSegmento As New Collection
Dim iPos As Integer, iTamFalta As Integer
Dim sProdSeg As String
Dim objProd As ClassProduto, iTeste As Integer, bAchouProd As Boolean
Dim objProdAux As ClassProduto, sProdutoBD As String, iProdPreenchido As Integer

On Error GoTo Erro_Produto_Ajusta_Formato

    objsegmento.sCodigo = "produto"

    'preenche toda colecao(colSegmento) em relacao ao formato corrente
    lErro = CF("Segmento_Le_Codigo", objsegmento, colSegmento)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sProdSeg = ""
    For Each objsegmento In colSegmento
        If objsegmento.iTipo = SEGMENTO_NUMERICO Then
            Select Case objsegmento.iPreenchimento
                Case ZEROS_ESPACOS
                    sProdSeg = sProdSeg & String(objsegmento.iTamanho, "0")
                Case ESPACOS
                    sProdSeg = sProdSeg & String(objsegmento.iTamanho, " ")
            End Select
        Else
            sProdSeg = sProdSeg & String(objsegmento.iTamanho, " ")
        End If
    Next
    
    Set objsegmento = colSegmento.Item(1)
    
    iPos = InStr(1, sProduto, objsegmento.sDelimitador)
        
    'Se tiver ponto e a quantidade de pontos corresponde a quantidade de segmentos - 1
    If iPos <> 0 And (colSegmento.Count - 1) = (Len(sProduto) - Len(Replace(sProduto, ".", ""))) Then
    
        bComMask = True
        
        sSeg = Mid(sProduto, 1, iPos - 1)
        
        iTamFalta = objsegmento.iTamanho - Len(sSeg)
        
        If iTamFalta = 0 Or objsegmento.iPreenchimento = PREENCH_LIMPA_BRANCOS Then
            sProdutoNovo = sProduto
        Else
            If objsegmento.iTipo = SEGMENTO_NUMERICO Then
                Select Case objsegmento.iPreenchimento
                    Case ZEROS_ESPACOS
                        sSeg = String(iTamFalta, "0") & sSeg
                    Case ESPACOS
                        sSeg = String(iTamFalta, " ") & sSeg
                End Select
            Else
                sSeg = sSeg & String(iTamFalta, " ")
            End If
            sProdutoNovo = sSeg & Mid(sProduto, iPos)
        End If
    Else
        bComMask = False
        
        iTamFalta = Len(sProdSeg) - Len(sProduto)
        
        If iTamFalta > 0 Then
            sProdutoNovo = sProduto & right(sProdSeg, iTamFalta)
        Else
            sProdutoNovo = left(sProduto, Len(sProdSeg))
        End If

    End If
    
    sProduto = sProdutoNovo

    Produto_Ajusta_Formato = SUCESSO

    Exit Function

Erro_Produto_Ajusta_Formato:

    Produto_Ajusta_Formato = gErr

    Select Case gErr
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213825)

    End Select

    Exit Function

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    Call Ordenacao_Limpa(objGridItens)

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Private Sub BotaoDesmarcarTodos_Click()
    Call Grid_Marca_Desmarca(objGridItens, iGrid_Sel_Col, DESMARCADO)
End Sub

Private Sub BotaoMarcarTodos_Click()
    Call Grid_Marca_Desmarca(objGridItens, iGrid_Sel_Col, MARCADO)
End Sub
