VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl TelaTab 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   8700
   Begin VB.ListBox Frames 
      Height          =   1620
      ItemData        =   "TelaTab.ctx":0000
      Left            =   4380
      List            =   "TelaTab.ctx":0002
      TabIndex        =   11
      Top             =   345
      Width           =   4035
   End
   Begin VB.Frame Controles 
      Caption         =   "Controles x Frames"
      Height          =   3300
      Left            =   90
      TabIndex        =   8
      Top             =   2010
      Width           =   8535
      Begin VB.ComboBox Frame 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "TelaTab.ctx":0004
         Left            =   3585
         List            =   "TelaTab.ctx":000B
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1125
         Width           =   2745
      End
      Begin MSMask.MaskEdBox Tipo 
         Height          =   270
         Left            =   6390
         TabIndex        =   13
         Top             =   1185
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Controle 
         Height          =   270
         Left            =   1155
         TabIndex        =   10
         Top             =   1125
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   2880
         Left            =   135
         TabIndex        =   9
         Top             =   270
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   5080
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.CommandButton BotaoAlterar 
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
      Left            =   2835
      Picture         =   "TelaTab.ctx":0012
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Altera a Operação da Árvore do Roteiro"
      Top             =   1455
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
      Left            =   1500
      Picture         =   "TelaTab.ctx":1938
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exclui a Operação da Árvore do Roteiro"
      Top             =   1440
      Width           =   1335
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
      Left            =   90
      Picture         =   "TelaTab.ctx":325E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Inclui a Operação na Árvore do Roteiro"
      Top             =   1425
      Width           =   1335
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   3705
      Picture         =   "TelaTab.ctx":4AAC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5370
      Width           =   1005
   End
   Begin MSMask.MaskEdBox Nome 
      Height          =   315
      Left            =   1020
      TabIndex        =   4
      Top             =   975
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Ordem 
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Top             =   375
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   1
      Mask            =   "#"
      PromptChar      =   " "
   End
   Begin VB.Label Label13 
      Caption         =   "Frames"
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
      Left            =   4365
      TabIndex        =   12
      Top             =   60
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ordem:"
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
      Left            =   135
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   7
      Top             =   420
      Width           =   615
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
      Left            =   195
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   5
      Top             =   1035
      Width           =   555
   End
End
Attribute VB_Name = "TelaTab"
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
Dim iGrid_Frame_Col As Integer
Dim iGrid_Tipo_Col As Integer

Public iAlterado As Integer

Const TIPO_GRID = 1
Const TIPO_FRAME = 2
Const TIPO_OUTRO = 3

'**** inicio do trecho a ser copiado *****
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Tela Tab"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Tela Tab"

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
    
    Ordem.PromptInclude = False
    Ordem.Text = "1"
    Ordem.PromptInclude = True
    
    Set objGridItens = New AdmGrid
    
    Call Inicializa_Grid_Itens(objGridItens)
   
    'Sinaliza que o Form_Loas ocorreu com sucesso
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

End Sub

Function Trata_Parametros(ByVal objTela As ClassCriaTela) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Faz a variável global a tela apontar para a variável passada
    Set gobjTela = objTela
        
    lErro = Traz_TelaTab_Tela(objTela)
    If lErro <> SUCESSO Then gError 136202
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 136202
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174644)
    
    End Select
    
    Exit Function
        
End Function

Function Saida_Celula(objGridItens As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridItens)

    If lErro = SUCESSO Then

        lErro = Saida_Celula_Frame(objGridItens)
        If lErro <> SUCESSO Then gError 123221

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174645)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174646)

    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro
    
    lErro = Move_TelaTab_Memoria(gobjTela)
    If lErro <> SUCESSO Then gError 136212
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
    
        Case 136212
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174647)

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

Private Sub Frame_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Frame_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Frame_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Frame_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Frame()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Tipo()
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set objGridItens = Nothing
    
End Sub

Private Function Saida_Celula_Frame(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Frame

    Set objGridInt.objControle = Frame

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132961

    Saida_Celula_Frame = SUCESSO

    Exit Function

Erro_Saida_Celula_Frame:

    Saida_Celula_Frame = gErr

    Select Case gErr

        Case 132961
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174648)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Traz_TelaTab_Tela(ByVal objTela As ClassCriaTela) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objControle As ClassCriaControles

On Error GoTo Erro_Traz_TelaTab_Tela
   
    For Each objControle In objTela.colControles
            
        If objControle.sGrid = "" And objControle.iTipo <> TIPO_FRAME Then
        
            iIndice = iIndice + 1
    
            GridItens.TextMatrix(iIndice, iGrid_Controle_Col) = objControle.sNome
            GridItens.TextMatrix(iIndice, iGrid_Frame_Col) = objControle.sFrame
            GridItens.TextMatrix(iIndice, iGrid_Tipo_Col) = objControle.sTipo
    
        End If
        
        If objControle.iTipo = TIPO_FRAME Then
            Frames.AddItem objControle.sNome
            Frames.ItemData(Frames.NewIndex) = objControle.iOrdem
            
            Frame.AddItem objControle.sNome
            Frame.ItemData(Frame.NewIndex) = objControle.iOrdem
        End If
    
    Next

    objGridItens.iLinhasExistentes = iIndice
    
    Traz_TelaTab_Tela = SUCESSO

    Exit Function

Erro_Traz_TelaTab_Tela:

    Traz_TelaTab_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174649)
    
    End Select
    
    Exit Function
    
End Function

Function Move_TelaTab_Memoria(ByVal objTela As ClassCriaTela) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objControle As ClassCriaControles

On Error GoTo Erro_Move_TelaTab_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        For Each objControle In objTela.colControles
                
            If GridItens.TextMatrix(iIndice, iGrid_Controle_Col) = objControle.sNome Then
        
                objControle.sFrame = GridItens.TextMatrix(iIndice, iGrid_Frame_Col)
        
            End If
            
            If GridItens.TextMatrix(iIndice, iGrid_Controle_Col) = objControle.sGrid Then
            
                objControle.sFrame = GridItens.TextMatrix(iIndice, iGrid_Frame_Col)
            
            End If
        
        Next
    
    Next
      
    Move_TelaTab_Memoria = SUCESSO

    Exit Function

Erro_Move_TelaTab_Memoria:

    Move_TelaTab_Memoria = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174650)
    
    End Select
    
    Exit Function
    
End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Controle")
    objGridInt.colColuna.Add ("Frame")
    objGridInt.colColuna.Add ("Tipo")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Controle.Name)
    objGridInt.colCampo.Add (Frame.Name)
    objGridInt.colCampo.Add (Tipo.Name)

    iGrid_Controle_Col = 1
    iGrid_Frame_Col = 2
    iGrid_Tipo_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR

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

    objControle.sNome = Nome.Text
    objControle.iOrdem = StrParaInt(Ordem.Text)
    objControle.iTipo = TIPO_FRAME
    objControle.sTipo = "Frame"
    
    gobjTela.colControles.Add objControle

    Frames.AddItem objControle.sNome
    Frames.ItemData(Frames.NewIndex) = StrParaInt(Ordem.Text)
    
    Frame.AddItem objControle.sNome
    
    Ordem.PromptInclude = False
    Ordem.Text = StrParaInt(Ordem.Text) + 1
    Ordem.PromptInclude = True
    
    Nome.Text = ""

End Sub

Private Sub BotaoAlterar_Click()

Dim objControle As New ClassCriaControles

On Error GoTo Erro_BotaoAlterar_Click
    
    For Each objControle In gobjTela.colControles
        
        If objControle.sNome = Nome.Text And objControle.iTipo = TIPO_FRAME Then
            objControle.iOrdem = StrParaInt(Ordem.Text)
        End If
    
    Next
    
    Exit Sub
    
Erro_BotaoAlterar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174651)

    End Select

    Exit Sub

End Sub

Private Sub BotaoRemover_Click()

Dim objControle As New ClassCriaControles
Dim iIndice As Integer
Dim bAchou As Boolean

On Error GoTo Erro_BotaoRemover_Click

    If bAchou >= 0 Then
        Frames.RemoveItem (Frames.ListIndex)
    End If
    
    iIndice = 0
    bAchou = False
    
    For Each objControle In gobjTela.colControles
        
        iIndice = iIndice + 1
        
        If objControle.sNome = Nome.Text And objControle.iTipo = TIPO_FRAME Then
            bAchou = True
            Exit For
        End If
    
    Next
    
    If bAchou Then
        gobjTela.colControles.Remove (iIndice)
    End If
    
    Nome.Text = ""
    Ordem.Text = ""
    
    Exit Sub
    
Erro_BotaoRemover_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174652)

    End Select

    Exit Sub
    
End Sub
