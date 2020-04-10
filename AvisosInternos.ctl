VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl AvisosInternosOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   4920
      Picture         =   "AvisosInternos.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   3990
      Picture         =   "AvisosInternos.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Detalhe 
      Height          =   870
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4515
      Width           =   9420
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":025C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Leia mais"
      Top             =   405
      Width           =   570
   End
   Begin VB.CheckBox Excluir 
      Height          =   210
      Left            =   1155
      TabIndex        =   13
      Top             =   1395
      Width           =   615
   End
   Begin VB.TextBox Assunto 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   690
      TabIndex        =   12
      Top             =   2220
      Width           =   6255
   End
   Begin VB.TextBox Link 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1365
      TabIndex        =   11
      Top             =   3105
      Width           =   585
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":0729
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Leia mais"
      Top             =   735
      Width           =   570
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":0BF6
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Leia mais"
      Top             =   1065
      Width           =   570
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":10C3
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Leia mais"
      Top             =   1395
      Width           =   570
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":1590
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Leia mais"
      Top             =   1725
      Width           =   570
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":1A5D
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Leia mais"
      Top             =   2055
      Width           =   570
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":1F2A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Leia mais"
      Top             =   2385
      Width           =   570
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":23F7
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Leia mais"
      Top             =   2715
      Width           =   570
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":28C4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Leia mais"
      Top             =   3060
      Width           =   570
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":2D91
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Leia mais"
      Top             =   3390
      Width           =   570
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":325E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Leia mais"
      Top             =   3720
      Width           =   570
   End
   Begin VB.CommandButton BotaoLink 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   8535
      Picture         =   "AvisosInternos.ctx":372B
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Leia mais"
      Top             =   4050
      Width           =   570
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   1200
      Width           =   1160
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid GridItens 
      Height          =   4485
      Left            =   15
      TabIndex        =   16
      Top             =   30
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   7911
      _Version        =   393216
   End
End
Attribute VB_Name = "AvisosInternosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim iAlterado As Integer

Dim gcolAvisos As New Collection
Dim gbDesativaFuncGrid As Boolean

Dim objGridItens As AdmGrid
Dim iGrid_Data_Col As Integer
Dim iGrid_Assunto_Col As Integer
Dim iGrid_Excluir_Col As Integer

Function Trata_Parametros() As Long
    gbDesativaFuncGrid = False
End Function

Private Sub BotaoCancela_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoLink_Click(Index As Integer)
Dim iLinha As Integer
Dim objAviso As ClassAvisos
    If Not (gcolAvisos Is Nothing) Then
        iLinha = GridItens.TopRow + Index
        If iLinha <= gcolAvisos.Count Then
            Set objAviso = gcolAvisos.Item(iLinha)
            If Len(Trim(objAviso.sLink)) > 0 Then
                Call ShellExecute(hWnd, "open", objAviso.sLink, vbNullString, vbNullString, 1)
                objAviso.iLido = MARCADO
            End If
        End If
    End If
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objAviso As ClassAvisos

On Error GoTo Erro_BotaoOK_Click

    For iIndice = 1 To objGridItens.iLinhasExistentes
        Set objAviso = gcolAvisos(iIndice)
        objAviso.iExcluido = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Excluir_Col))
    Next

    lErro = CF("Avisos_Grava", gcolAvisos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208983)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim objAviso As ClassAvisos, iIndice As Integer

On Error GoTo Erro_Form_Load

    gbDesativaFuncGrid = False
    
    lErro = CF("Avisos_Le", gcolAvisos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Set objGridItens = New AdmGrid
    
    'Executa a Inicialização do grid Bloqueio
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iIndice = 0
    For Each objAviso In gcolAvisos
        iIndice = iIndice + 1
        
        GridItens.TextMatrix(iIndice, iGrid_Data_Col) = Format(objAviso.dtData)
        GridItens.TextMatrix(iIndice, iGrid_Assunto_Col) = objAviso.sAssunto
        
        If objAviso.iNovo = MARCADO Then
            gbDesativaFuncGrid = True
            GridItens.Row = iIndice
            GridItens.Col = iGrid_Assunto_Col
            GridItens.CellFontBold = True
            
            GridItens.Col = iGrid_Data_Col
            GridItens.Row = iIndice
            GridItens.CellFontBold = True
            gbDesativaFuncGrid = False
        End If
        
    Next
    objGridItens.iLinhasExistentes = gcolAvisos.Count
    
    Call Trata_BotaoLink

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    gbDesativaFuncGrid = False
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208984)
    
    End Select
    
    Exit Sub

End Sub

Private Sub GridItens_Click()
Dim iExecutaEntradaCelula As Integer

    If Not gbDesativaFuncGrid Then
        Call Grid_Click(objGridItens, iExecutaEntradaCelula)
    
        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridItens, iAlterado)
        End If
    End If
End Sub

Private Sub GridItens_GotFocus()
    If Not gbDesativaFuncGrid Then Call Grid_Recebe_Foco(objGridItens)
End Sub

Private Sub GridItens_EnterCell()
    If Not gbDesativaFuncGrid Then Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Private Sub GridItens_LeaveCell()
    If Not gbDesativaFuncGrid Then Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not gbDesativaFuncGrid Then Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer
    
    If Not gbDesativaFuncGrid Then
        Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)
    
        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridItens, iAlterado)
        End If
    End If

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
    If Not gbDesativaFuncGrid Then Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub GridItens_RowColChange()
Dim objAviso As ClassAvisos

    If Not gbDesativaFuncGrid Then
        Call Grid_RowColChange(objGridItens)
        If Not (gcolAvisos Is Nothing) Then
            If GridItens.Row <= gcolAvisos.Count And GridItens.Row > 0 Then
                Set objAviso = gcolAvisos.Item(GridItens.Row)
                Detalhe.Text = objAviso.sAssunto
            Else
                Detalhe.Text = ""
            End If
        Else
            Detalhe.Text = ""
        End If
    End If
End Sub

Private Sub GridItens_Scroll()
    If Not gbDesativaFuncGrid Then
        Call Grid_Scroll(objGridItens)
        Call Trata_BotaoLink
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Set objGridItens = Nothing
     Set gcolAvisos = Nothing
End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Bloqueio
Dim objControle As Object

On Error GoTo Erro_Inicializa_Grid_Itens

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Excluir")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Assunto")
    objGridInt.colColuna.Add ("Leia+")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Excluir.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (Assunto.Name)
    objGridInt.colCampo.Add (Link.Name)
    
    iGrid_Excluir_Col = 1
    iGrid_Data_Col = 2
    iGrid_Assunto_Col = 3
    
    objGridInt.objGrid = GridItens

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Itens:

    Inicializa_Grid_Itens = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208984)
    
    End Select
    
    Exit Function

End Function


Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_Excluir_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, Excluir)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
            End Select
            
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 208737

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 208737
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208738)

    End Select

    Exit Function

End Function

Private Sub Trata_BotaoLink()

Dim lErro As Long
Dim iIndice As Integer
Dim iIndiceBotao As Integer
Dim objAviso As ClassAvisos

On Error GoTo Erro_Trata_BotaoLink

    iIndiceBotao = -1
    For iIndice = GridItens.TopRow To GridItens.TopRow + 11
        iIndiceBotao = iIndiceBotao + 1
        If Not (gcolAvisos Is Nothing) Then
            If iIndice <= gcolAvisos.Count Then
                Set objAviso = gcolAvisos.Item(iIndice)
                
                If Len(Trim(objAviso.sLink)) > 0 Then
                    BotaoLink(iIndiceBotao).Enabled = True
                Else
                    BotaoLink(iIndiceBotao).Enabled = False
                End If
            Else
                BotaoLink(iIndiceBotao).Enabled = False
            End If
        Else
            BotaoLink(iIndiceBotao).Enabled = False
        End If
    Next
       
    Exit Sub
    
Erro_Trata_BotaoLink:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208984)
    
    End Select
    
    Exit Sub

End Sub



'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Avisos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AvisosInternos"
    
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

