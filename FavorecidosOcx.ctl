VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Begin VB.UserControl FavorecidosOcx 
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   5550
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1485
      Picture         =   "FavorecidosOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   255
      Width           =   300
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   3630
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "FavorecidosOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "FavorecidosOcx.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FavorecidosOcx.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CheckBox Inativo 
      Caption         =   "Inativo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2250
      TabIndex        =   2
      Top             =   285
      Width           =   975
   End
   Begin VB.ListBox Lista_Favorecidos 
      Height          =   1620
      Left            =   195
      TabIndex        =   4
      Top             =   1695
      Width           =   5160
   End
   Begin MSMask.MaskEdBox Nome 
      Height          =   300
      Left            =   870
      TabIndex        =   3
      Top             =   870
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   529
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   870
      TabIndex        =   0
      Top             =   255
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Favorecidos"
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
      Left            =   210
      TabIndex        =   9
      Top             =   1455
      Width           =   1050
   End
   Begin VB.Label Label2 
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
      Left            =   225
      TabIndex        =   10
      Top             =   930
      Width           =   555
   End
   Begin VB.Label LabelFavorecido 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   315
      Width           =   660
   End
End
Attribute VB_Name = "FavorecidosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoFavorecido As AdmEvento
Attribute objEventoFavorecido.VB_VarHelpID = -1

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Nao há Favorecido selecionado. Gera número automático.
    lErro = CF("Favorecidos_Automatico", iCodigo)
    If lErro <> SUCESSO Then Error 57548
    
    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57548
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160177)
    
    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim sEspacos As String
Dim sListBoxItem As String
Dim iIndice As Integer

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    Codigo.Text = CStr(colCampoValor.Item("Codigo").vValor)
    Nome.Text = colCampoValor.Item("Nome").vValor
    Inativo.Value = colCampoValor.Item("Inativo").vValor

    iAlterado = 0
    
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim iCodigo As Integer
    
    'Informa tabela associada à Tela
    sTabela = "Favorecidos"
    
    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    If Len(Codigo.Text) > 0 Then
        iCodigo = CInt(Codigo.Text)
    Else
        iCodigo = 0
    End If
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", iCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", Nome.Text, STRING_FAVORECIDO, "Nome"
    colCampoValor.Add "Inativo", Inativo.Value, 0, "Inativo"
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then Error 19410

    'Limpa Tela
    Call Limpa_Tela_Favorecidos

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 19410

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 160178)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFavorecidos As New ClassFavorecidos
Dim iCodigo As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se dados de Favorecidos foram informados
    If Len(Codigo.Text) = 0 Then Error 17031
    If Len(Trim(Nome.Text)) = 0 Then Error 17032
        
    'Preenche objeto Favorecidos
    objFavorecidos.iCodigo = CInt(Codigo.Text)
    objFavorecidos.sNome = Trim(Nome.Text)
    objFavorecidos.iInativo = Inativo.Value
        
    'Grava o Favorecido no banco de dados
    lErro = CF("Favorecidos_Grava", objFavorecidos)
    If lErro <> SUCESSO Then Error 17034
    
    'Remove e adiciona na ListBox
    Call Lista_Favorecidos_Remove(objFavorecidos)
    Call Lista_Favorecidos_Adiciona(objFavorecidos)
            
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 17031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_FAVORECIDO_NAO_INFORMADO", Err)
            
        Case 17032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_FAVORECIDO_NAO_INFORMADO", Err)
    
        Case 17034
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160179)

     End Select
        
     Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 17061
 
    Call Limpa_Tela_Favorecidos
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err
    
        Case 17061
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160180)

     End Select
        
     Exit Sub

End Sub

Public Sub Form_Activate()
    
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colFavorecidos As New Collection
Dim objFavorecidos As ClassFavorecidos
Dim sCodigo As String
Dim sListBoxItem As String

On Error GoTo Erro_Favorecidos_Form_Load

    Set objEventoFavorecido = New AdmEvento
    
    'Preenche a ListBox com Favorecidos existentes no BD
    lErro = CF("Favorecidos_Le_Todos", colFavorecidos)
    If lErro <> SUCESSO Then Error 17001
    
    For Each objFavorecidos In colFavorecidos
    
        Call Lista_Favorecidos_Adiciona(objFavorecidos)
        
    Next
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Favorecidos_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
            
        Case 17001
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160181)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objFavorecidos As ClassFavorecidos) As Long

Dim lErro As Long
Dim sListBoxItem As String
Dim sEspacos As String
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há um Favorecido selecionado, exibir seus dados
    If Not (objFavorecidos Is Nothing) Then
        
        'Verifica se o Favorecido existe
        lErro = CF("Favorecido_Le", objFavorecidos)
        If lErro <> 17015 And lErro <> SUCESSO Then Error 17003
        
        'se Favorecido está cadastrado
        If lErro = SUCESSO Then
        
            Call Traz_Favorecido_Tela(objFavorecidos)
            
        Else
        
            'Favorecido não está cadastrado
            Codigo.Text = CStr(objFavorecidos.iCodigo)
            
        End If
                
    End If

    iAlterado = 0
        
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 17003
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160182)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoFavorecido = New AdmEvento
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

On Error GoTo Erro_Codigo_Validate

    'Verifica preenchimento do sequencial
    If Len(Trim(Codigo.Text)) > 0 Then

        'Verifica se o sequencial é numérico
        If Not IsNumeric(Codigo.Text) Then Error 55966

        'Verifica se codigo é menor que um
        If CInt(Codigo.Text) < 1 Then Error 55967

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case Err

        Case 55966, 55967
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INVALIDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160183)

    End Select

    Exit Sub


End Sub

Private Sub Inativo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Lista_Favorecidos_DblClick()

Dim lErro As Long
Dim sListBoxItem As String
Dim objFavorecidos As New ClassFavorecidos

On Error GoTo Erro_Favorecidos_DblClick
    
    'Se não há Favorecido selecionado sai da rotina
    If Lista_Favorecidos.ListIndex = -1 Then Exit Sub
    
    objFavorecidos.iCodigo = Lista_Favorecidos.ItemData(Lista_Favorecidos.ListIndex)
       
    'Verifica se o Favorecido existe
    lErro = CF("Favorecido_Le", objFavorecidos)
    If lErro <> 17015 And lErro <> SUCESSO Then Error 17028
    
    If lErro = SUCESSO Then 'Favorecido está cadastrado
            
        Call Traz_Favorecido_Tela(objFavorecidos)

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
        
    Else 'Favorecido não existe
    
        'Exclui da ListBox
        Lista_Favorecidos.RemoveItem (Lista_Favorecidos.ListIndex)
        
    End If
 
    Exit Sub
    
Erro_Favorecidos_DblClick:

    Select Case Err
            
        Case 17028
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160184)

    End Select
    
    Exit Sub
    
End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Private Sub Lista_Favorecidos_Adiciona(objFavorecidos As ClassFavorecidos)
'Inclui na Lista mantendo a ordenacao por codigo

Dim iIndice As Integer

    For iIndice = 0 To Lista_Favorecidos.ListCount - 1

        If Lista_Favorecidos.ItemData(iIndice) > objFavorecidos.iCodigo Then Exit For
        
    Next

    Lista_Favorecidos.AddItem objFavorecidos.iCodigo & SEPARADOR & objFavorecidos.sNome, iIndice
    Lista_Favorecidos.ItemData(iIndice) = objFavorecidos.iCodigo

End Sub

Private Sub Lista_Favorecidos_Remove(objFavorecidos As ClassFavorecidos)
'Percorre a ListBox Lista_Favorecidos para remover o tipo caso ele exista

Dim iIndice As Integer

    For iIndice = 0 To Lista_Favorecidos.ListCount - 1
    
        If Lista_Favorecidos.ItemData(iIndice) = objFavorecidos.iCodigo Then
    
            Lista_Favorecidos.RemoveItem iIndice
            Exit For
    
        End If
    
    Next

End Sub

Function Limpa_Tela_Favorecidos() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Favorecidos

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    'Limpa a Tela
    Call Limpa_Tela(Me)
    
    Codigo.Text = ""
    
    Inativo.Value = 0
    
    iAlterado = 0
    
    Limpa_Tela_Favorecidos = SUCESSO
        
    Exit Function
    
Erro_Limpa_Tela_Favorecidos:

    Limpa_Tela_Favorecidos = Err
    
    Select Case Err
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160185)

     End Select
     
     Exit Function

End Function

Sub Traz_Favorecido_Tela(objFavorecidos As ClassFavorecidos)

    'Favorecido está cadastrado
    Codigo.Text = CStr(objFavorecidos.iCodigo)
    Nome.Text = objFavorecidos.sNome
    Inativo.Value = objFavorecidos.iInativo
            
    iAlterado = 0
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FAVORECIDOS
    Set Form_Load_Ocx = Me
    Caption = "Favorecidos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Favorecidos"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelFavorecido_Click
        End If
    
    End If

End Sub


Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Public Sub LabelFavorecido_Click()

Dim objFavorecido As New ClassFavorecidos
Dim colSelecao As New Collection

    If Len(Trim(Codigo.Text)) > 0 Then
        objFavorecido.iCodigo = StrParaInt(Codigo.Text)
        objFavorecido.sNome = Nome.Text
    End If

    Call Chama_Tela("FavorecidosLista", colSelecao, objFavorecido, objEventoFavorecido)
    
End Sub

Private Sub objEventoFavorecido_evSelecao(obj1 As Object)

Dim objFavorecido As ClassFavorecidos

    Set objFavorecido = obj1

    If Not (objFavorecido Is Nothing) Then Call Traz_Favorecido_Tela(objFavorecido)

    Me.Show

    Exit Sub

End Sub


