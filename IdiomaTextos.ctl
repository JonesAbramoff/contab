VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl IdiomaTextos 
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   KeyPreview      =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   10095
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
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
      Left            =   4125
      Picture         =   "IdiomaTextos.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5325
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
      Left            =   5115
      Picture         =   "IdiomaTextos.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5325
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   1350
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   9885
      Begin VB.Label Valor 
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Left            =   855
         TabIndex        =   12
         Top             =   600
         Width           =   8925
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   315
         TabIndex        =   9
         Top             =   675
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tela:"
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
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   285
         Width           =   450
      End
      Begin VB.Label Tela 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   855
         TabIndex        =   7
         Top             =   210
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Controle:"
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
         Index           =   0
         Left            =   4245
         TabIndex        =   6
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Controle 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5055
         TabIndex        =   5
         Top             =   210
         Width           =   1785
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Textos"
      Height          =   3930
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   1365
      Width           =   9975
      Begin VB.TextBox DetalheM 
         Height          =   885
         Left            =   105
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   3000
         Width           =   9765
      End
      Begin VB.TextBox IdiomaDet 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2655
         Width           =   1785
      End
      Begin VB.TextBox IdiomaCod 
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   3240
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   810
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Idioma 
         BorderStyle     =   0  'None
         Height          =   150
         Left            =   3045
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   810
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.TextBox Detalhe 
         Height          =   885
         Left            =   105
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3000
         Width           =   9765
      End
      Begin VB.TextBox Texto 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1185
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1380
         Width           =   8205
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   2265
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   3995
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Idioma:"
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
         Left            =   225
         TabIndex        =   13
         Top             =   2685
         Width           =   630
      End
   End
End
Attribute VB_Name = "IdiomaTextos"
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
Dim iGrid_Sigla_Col As Integer
Dim iGrid_Texto_Col As Integer
Dim iGrid_Idioma_Col As Integer
Dim iGrid_IdiomaCod_Col As Integer

Private gobjIdiomaTela As ClassIdiomaTela
Private gcolTextos As Collection

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Textos em outros Idiomas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "IdiomaTextos"

End Function

Public Sub Show()
'???? comentei para nao dar erro nesta tela pq é modal
'    Parent.Show
 '   Parent.SetFocus
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

    Set gobjIdiomaTela = Nothing
    Set gcolTextos = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211650)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211651)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(ByVal objIdiomaTela As ClassIdiomaTela) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjIdiomaTela = objIdiomaTela
    Set gcolTextos = New Collection
    
    lErro = Traz_Dados_Tela(objIdiomaTela)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel
    
    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211652)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Tela(ByVal objIdiomaTela As ClassIdiomaTela) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objIdioma As ClassIdioma
Dim objControle As ClassIdiomaTelaControls
Dim objTabela As ClassIdiomaTab
Dim objCampo As ClassIdiomaTabCampo
Dim objTexto As ClassIdiomaTabCampoTexto

On Error GoTo Erro_Traz_Dados_Tela

    Valor.Caption = objIdiomaTela.objControleAtivo.Text

    iIndice = 0
    For Each objIdioma In objIdiomaTela.colIdiomas
        If objIdioma.iPadrao = DESMARCADO Then
            iIndice = iIndice + 1
            GridItens.TextMatrix(iIndice, iGrid_Sigla_Col) = objIdioma.sSigla
            GridItens.TextMatrix(iIndice, iGrid_Idioma_Col) = objIdioma.sDescricao
            
            For Each objControle In objIdiomaTela.colControles
                If UCase(objControle.sNomeControle) = UCase(objIdiomaTela.objControleAtivo.Name) Then
                    Tela.Caption = objControle.sNomeTelaExibicao
                    Controle.Caption = objControle.sNomeControleExibicao
                    If objControle.iComMultiLine = MARCADO Then
                        DetalheM.Visible = True
                        Detalhe.Visible = False
                    Else
                        DetalheM.Visible = False
                        Detalhe.Visible = True
                    End If
                    If objControle.iComMaxLen = MARCADO Then
                        Detalhe.MaxLength = objIdiomaTela.objControleAtivo.MaxLength
                    End If
                    For Each objTabela In objIdiomaTela.colTabelas
                        If UCase(objTabela.sNomeTabela) = UCase(objControle.sNomeTabela) Then
                            For Each objCampo In objTabela.colCampos
                                If UCase(objCampo.sNomeCampo) = UCase(objControle.sNomeCampo) Then
                                    For Each objTexto In objCampo.colTextos
                                        If objTexto.iIdioma = objIdioma.iCodigo Then
                                             GridItens.TextMatrix(iIndice, iGrid_Texto_Col) = objTexto.sTexto
                                             gcolTextos.Add objTexto
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                    Exit For
                End If
            Next
        End If
    Next
    objGridItens.iLinhasExistentes = iIndice
    
    For iIndice = objGridItens.iLinhasExistentes + 1 To GridItens.Rows - 1
        GridItens.TextMatrix(iIndice, iGrid_Sigla_Col) = ""
    Next

    Traz_Dados_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_Tela:
   
    Traz_Dados_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211653)

    End Select

    Exit Function
    
End Function

Private Function Move_Tela_Memoria()

Dim lErro As Long, iIndice As Integer
Dim objTexto As ClassIdiomaTabCampoTexto

On Error GoTo Erro_Move_Tela_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
        Set objTexto = gcolTextos.Item(iIndice)
        objTexto.sTexto = GridItens.TextMatrix(iIndice, iGrid_Texto_Col)
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:
   
    Move_Tela_Memoria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211655)

    End Select

    Exit Function
    
End Function

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Idioma.Width = 0
    IdiomaCod.Width = 0
    
    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("Idioma")
    objGrid.colColuna.Add ("Texto")
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Texto.Name)
    objGrid.colCampo.Add (Idioma.Name)
    objGrid.colCampo.Add (IdiomaCod.Name)

    'Colunas do Grid
    iGrid_Sigla_Col = 0
    iGrid_Texto_Col = 1
    iGrid_Idioma_Col = 2
    iGrid_IdiomaCod_Col = 3

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 101

    objGrid.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 900

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGrid)

    GridItens.ColWidth(iGrid_Idioma_Col) = 0
    GridItens.ColWidth(iGrid_IdiomaCod_Col) = 0

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
    Call Exibe_CampoDet_Grid(objGridItens, iGrid_Texto_Col, Detalhe)
    Call Exibe_CampoDet_Grid(objGridItens, iGrid_Texto_Col, DetalheM)
    Call Exibe_CampoDet_Grid(objGridItens, iGrid_Idioma_Col, IdiomaDet)
    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
        Detalhe.Locked = False
    Else
        Detalhe.Locked = True
    End If
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

Private Sub Texto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Texto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Texto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Texto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Texto
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

                Case iGrid_Texto_Col

                    lErro = Saida_Celula_Padrao(objGridInt, Texto)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            End Select
                    
        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 211656

    End If
    
    Call Exibe_CampoDet_Grid(objGridItens, iGrid_Texto_Col, Detalhe)
    Call Exibe_CampoDet_Grid(objGridItens, iGrid_Texto_Col, DetalheM)
    Call Exibe_CampoDet_Grid(objGridItens, iGrid_Idioma_Col, IdiomaDet)

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 211656
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211657)

    End Select

    Exit Function

End Function

Private Sub BotaoOK_Click()
Dim lErro As Long
    lErro = Move_Tela_Memoria
    If lErro = SUCESSO Then
        'Nao mexer no obj da tela
        giRetornoTela = vbOK
    
        Unload Me
    End If
    
End Sub

Private Sub Exibe_CampoDet_Grid(ByVal objGridInt As AdmGrid, ByVal iColunaExibir As Integer, ByVal objControle As Object)

Dim iLinha As Integer

On Error GoTo Erro_Exibe_CampoDet_Grid

    iLinha = objGridInt.objGrid.Row
    
    If iLinha > 0 And iLinha <= objGridInt.iLinhasExistentes Then
        objControle.Text = objGridInt.objGrid.TextMatrix(iLinha, iColunaExibir)
    Else
        objControle.Text = ""
    End If

    Exit Sub

Erro_Exibe_CampoDet_Grid:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208641)

    End Select

    Exit Sub
    
End Sub

Private Sub Detalhe_Validate(Cancel As Boolean)
    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
        GridItens.TextMatrix(GridItens.Row, iGrid_Texto_Col) = Detalhe.Text
    End If
End Sub

Private Sub DetalheM_Validate(Cancel As Boolean)
    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
        GridItens.TextMatrix(GridItens.Row, iGrid_Texto_Col) = DetalheM.Text
    End If
End Sub
