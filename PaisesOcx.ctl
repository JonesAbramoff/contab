VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl PaisesOcx 
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   5685
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1620
      Picture         =   "PaisesOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   345
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3420
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "PaisesOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "PaisesOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "PaisesOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "PaisesOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Paises 
      Height          =   2010
      ItemData        =   "PaisesOcx.ctx":0A7E
      Left            =   120
      List            =   "PaisesOcx.ctx":0A80
      TabIndex        =   4
      Top             =   2265
      Width           =   5430
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   330
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Nome 
      Height          =   315
      Left            =   1170
      TabIndex        =   2
      Top             =   870
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CodBacen 
      Height          =   315
      Left            =   1170
      TabIndex        =   3
      Top             =   1395
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label DescBacen 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1980
      TabIndex        =   14
      Top             =   1395
      Width           =   3555
   End
   Begin VB.Label LabelPaisesBC 
      AutoSize        =   -1  'True
      Caption         =   "Código BC:"
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
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   13
      Top             =   1455
      Width           =   945
   End
   Begin VB.Label Label1 
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
      Left            =   420
      TabIndex        =   10
      Top             =   390
      Width           =   675
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
      Left            =   525
      TabIndex        =   11
      Top             =   915
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Países"
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
      Left            =   150
      TabIndex        =   12
      Top             =   2040
      Width           =   615
   End
End
Attribute VB_Name = "PaisesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoPaisesBC As AdmEvento
Attribute objEventoPaisesBC.VB_VarHelpID = -1

Dim iAlterado As Integer

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lNumAuto As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera Código automático
    lErro = CF("Config_ObterAutomatico", "CPRConfig", "NUM_PROX_PAIS", "Paises", "Codigo", lNumAuto)
    If lErro <> SUCESSO Then Error 57552

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lNumAuto)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57552
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164293)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objPais As New ClassPais
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Pais foi informado
    If Len(Codigo.Text) = 0 Then Error 47879
    
    objPais.iCodigo = CInt(Codigo.Text)

    'Verifica se o Pais existe
    lErro = CF("Paises_Le", objPais)
    If lErro <> SUCESSO And lErro <> 47876 Then Error 47880
    
    'Pais não está cadastrado - --- -> Erro
    If lErro = 47876 Then Error 47881
    
    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_PAIS")
            
    If vbMsgRet = vbYes Then
        
        'exclui o Pais
        lErro = CF("Paises_Exclui", objPais.iCodigo)
        If lErro Then Error 47882
        
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
        
        'Exclui o Pais da ListBox
        Call Paises_Remove(objPais.iCodigo)
    
        'Limpa a Tela
        Call Limpa_Tela(Me)
        
        DescBacen.Caption = ""
            
        iAlterado = 0
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
                    
        Case 47879
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_PREENCHIDO", Err)
            Codigo.SetFocus
            
        Case 47880, 47882
        
        Case 47881
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", Err, objPais.iCodigo)
            Codigo.SetFocus
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 164294)
        
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava Pais
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 47884

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    DescBacen.Caption = ""

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 47884

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164295)

    End Select


End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()
    
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoPaisesBC = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)
    
End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim objPais As New ClassPais
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    objPais.iCodigo = colCampoValor.Item("Codigo").vValor

    If objPais.iCodigo <> 0 Then

        lErro = CF("Paises_Le", objPais)
        If lErro <> SUCESSO And lErro <> 47876 Then gError 76423

        'Coloca colCampoValor na Tela
        'Conversão de tipagem para a tipagem da tela se necessário
        Codigo.PromptInclude = False
        Codigo.Text = CStr(colCampoValor.Item("Codigo").vValor)
        Codigo.PromptInclude = True
        Nome.Text = colCampoValor.Item("Nome").vValor
    
    End If
    
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
    
        Case 76423
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164296)

     End Select
        
     Exit Sub
        
End Sub

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim iCodigo As Integer
Dim objPais As New ClassPais
    
    'Informa tabela associada à Tela
    sTabela = "Paises"
    
    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    objPais.iCodigo = StrParaLong(Codigo.Text)
    
    objPais.sNome = Nome.Text
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objPais.iCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", objPais.sNome, STRING_PAISES_NOME, "Nome"

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
 
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 47886
    
    DescBacen.Caption = ""

    'Limpa a Tela
    Call Limpa_Tela(Me)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err
    
        Case 47886
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164296)

     End Select
        
     Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoPaisesBC = New AdmEvento

    'Lê cada código e descrição da tabela Países
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 47888

    'Preenche cada ComboBox País com os objetos da coleção colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        Paises.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Paises.ItemData(Paises.NewIndex) = objCodigoDescricao.iCodigo

    Next
    
    lErro_Chama_Tela = SUCESSO

    iAlterado = 0

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 47888

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164297)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub
    
Function Trata_Parametros(Optional objPais As ClassPais) As Long

Dim lErro As Long
Dim sLista As String

On Error GoTo Erro_Trata_Parametros

    'Se há um Pais selecionado, exibir seus dados
    If Not (objPais Is Nothing) Then
    
        lErro = CF("Paises_Le", objPais)
        If lErro <> SUCESSO And lErro <> 47876 Then gError 76423
        
        If lErro = SUCESSO Then
        
            sLista = objPais.iCodigo & SEPARADOR & objPais.sNome
        
            lErro = List_Item_Igual(Paises, sLista)
            If lErro <> SUCESSO And lErro <> 12253 Then gError 47890
        
            'Se não encontrou o pais em questao -----> Erro
            If lErro = 12253 Then gError 47891
        
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objPais.iCodigo)
            Codigo.PromptInclude = True
            
            Nome.Text = objPais.sNome
            
            CodBacen.PromptInclude = False
            If objPais.iCodBacen <> 0 Then
                CodBacen.Text = CStr(objPais.iCodBacen)
            Else
                CodBacen.Text = ""
            End If
            CodBacen.PromptInclude = True
            Call CodBacen_Validate(bSGECancelDummy)
            
        Else
        
            'Limpa a tela
            Call Limpa_Tela(Me)
            DescBacen.Caption = ""
                
            'Exibe apenas o código
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objPais.iCodigo)
            Codigo.PromptInclude = True
            
        End If
        
    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 47890

        Case 47891
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", gErr, objPais.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164298)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 57975
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True
    
    Select Case Err
        
        Case 57975 'Erro tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164299)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Paises_DblClick()

Dim lErro As Long
Dim sListBoxItem As String
Dim lSeparadorPosicao As Long
Dim objPais As New ClassPais

On Error GoTo Erro_Paises_DblClick

    'Pega a String do ítem selecionado
    sListBoxItem = Paises.List(Paises.ListIndex)

    'Acha a posição do separador (-)
    lSeparadorPosicao = InStr(sListBoxItem, SEPARADOR)

    'Preenche Código e Descrição do Histórico na Tela
    Codigo.PromptInclude = False
    Codigo.Text = Trim(Left(sListBoxItem, lSeparadorPosicao - 1))
    Codigo.PromptInclude = True
    Nome.Text = Mid(sListBoxItem, lSeparadorPosicao + 1)
    
    objPais.iCodigo = StrParaInt(Codigo.Text)
    
    lErro = CF("Paises_Le", objPais)
    If lErro <> SUCESSO And lErro <> 47876 Then gError 200976
    
    CodBacen.PromptInclude = False
    If objPais.iCodBacen <> 0 Then
        CodBacen.Text = CStr(objPais.iCodBacen)
    Else
        CodBacen.Text = ""
    End If
    CodBacen.PromptInclude = True
    Call CodBacen_Validate(bSGECancelDummy)
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

Erro_Paises_DblClick:

    Select Case gErr
    
        Case 200976

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200975)

    End Select

    Exit Sub
    
End Sub

Private Sub Paises_Adiciona(objPais As ClassPais)
        
Dim sEspacos As String
Dim sListBoxItem As String
Dim iIndice As Integer

    For iIndice = 0 To Paises.ListCount - 1
        
        If Paises.ItemData(iIndice) > objPais.iCodigo Then Exit For
        
    Next
    
    'Concatena o código com a descrição da Mensagem
    sListBoxItem = CStr(objPais.iCodigo) & SEPARADOR & objPais.sNome
    Paises.AddItem (sListBoxItem), iIndice
    Paises.ItemData(iIndice) = objPais.iCodigo
        
End Sub

Private Sub Paises_Remove(iCodigo As Integer)
'Remove da ListBox

Dim iIndice As Integer

    For iIndice = 0 To Paises.ListCount - 1
    
        If Paises.ItemData(iIndice) = iCodigo Then
            
            Paises.RemoveItem (iIndice)
            Exit For
            
        End If
    
    Next

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objPais As New ClassPais

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then Error 47892

    'verifica preenchimento do nome
    If Len(Trim(Nome.Text)) = 0 Then Error 47893
    
    'Preenche objeto
    objPais.iCodigo = CInt(Codigo.Text)
    objPais.sNome = Nome.Text
    objPais.iCodBacen = StrParaInt(CodBacen.Text)
    
    'grava na Tabela o Pais a que esta sendo cadastrado
    lErro = CF("Paises_Grava", objPais)
    If lErro <> SUCESSO Then Error 47894

    'Remove e adiciona na ListBox
    Call Paises_Remove(objPais.iCodigo)
    Call Paises_Adiciona(objPais)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 47892
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 47893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", Err)
      
        Case 47894

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164300)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PAISES
    Set Form_Load_Ocx = Me
    Caption = "Cadastro de Países"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Paises"
    
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
End Sub




Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelPaisesBC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPaisesBC, Source, X, Y)
End Sub

Private Sub LabelPaisesBC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPaisesBC, Button, Shift, X, Y)
End Sub

Private Sub LabelPaisesBC_Click()

Dim colSelecao As New Collection
Dim objPaisBC As New ClassPaisesBC

On Error GoTo Erro_LabelPaisesBC_Click

    objPaisBC.sPais = Nome.Text

    'Chama Tela TituloReceberLista
    Call Chama_Tela("PaisesBCLista", colSelecao, objPaisBC, objEventoPaisesBC)

    Exit Sub

Erro_LabelPaisesBC_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200973)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPaisesBC_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPaisBC As ClassPaisesBC

On Error GoTo Erro_objEventoPaisesBC_evSelecao

    Set objPaisBC = obj1
    
    CodBacen.PromptInclude = False
    CodBacen.Text = CStr(objPaisBC.iCodBacen)
    CodBacen.PromptInclude = True
    Call CodBacen_Validate(bSGECancelDummy)
    
    Nome.Text = Left(objPaisBC.sPais, STRING_PAISES_NOME)
    
    Exit Sub

Erro_objEventoPaisesBC_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200974)

    End Select

    Exit Sub

End Sub

Private Sub CodBacen_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodBacen_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodBacen, iAlterado)
End Sub

Private Sub CodBacen_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sDescBacen As String

On Error GoTo Erro_CodBacen_Validate

    If Len(Trim(CodBacen.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(CodBacen.Text)
        If lErro <> SUCESSO Then gError 57975
        
        lErro = CF("Le_Campo_Tabela", "PaisesBC", "Pais", TIPO_STR, "CodBacen", StrParaInt(CodBacen.Text), sDescBacen, DescBacen)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200981
        
        If lErro <> SUCESSO Then gError 200982
    Else
    
        DescBacen.Caption = ""
        
    End If
    
    Exit Sub
    
Erro_CodBacen_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 57975, 200981 'Erro tratado na rotina chamada
        
        Case 200982
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_CODBACEN_NAO_CADASTRADO", gErr, StrParaInt(CodBacen.Text))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164299)
    
    End Select
    
    Exit Sub

End Sub
