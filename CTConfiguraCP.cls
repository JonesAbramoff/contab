VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTConfiguraCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

'Não transferir.
Const CPCONFIG_AGLUTINA_LANCAM_POR_DIA = 0
Const CPCONFIG_GERA_LOTE_AUTOMATICO = 1

Dim iAlterado As Integer

'--- inicio dos properties get dos controles da tela

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get ListaConfigura() As Object
     Set ListaConfigura = objUserControl.Controls("ListaConfigura")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155621)

    End Select

    Exit Function

End Function

Public Sub Form_Load()
           
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Checa Aglutina lançamentos por dia
    If gobjCP.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA Then
        ListaConfigura.Selected(CPCONFIG_AGLUTINA_LANCAM_POR_DIA) = True
    Else
        ListaConfigura.Selected(CPCONFIG_AGLUTINA_LANCAM_POR_DIA) = False
    End If
    
'    If giTipoVersao <> VERSAO_LIGHT Then
    
        'Checa "Gera Lote Automatico"
        If gobjCP.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO Then
            ListaConfigura.Selected(CPCONFIG_GERA_LOTE_AUTOMATICO) = True
        Else
            ListaConfigura.Selected(CPCONFIG_GERA_LOTE_AUTOMATICO) = False
        End If
            
'    End If
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155622)

    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 41878
    
    iAlterado = 0
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err

        Case 41878 'Tratado na Rotina Chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155623)

    End Select

    Exit Sub
    
End Sub
    
Public Function Gravar_Registro() As Long

Dim lErro As Long
    
On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Move os Dados para Memoria
    If ListaConfigura.Selected(CPCONFIG_AGLUTINA_LANCAM_POR_DIA) = True Then
        gobjCP.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA
    Else
        gobjCP.iAglutinaLancamPorDia = NAO_AGLUTINA_LANCAM_POR_DIA
    End If
    
'    If giTipoVersao <> VERSAO_LIGHT Then
    
        If ListaConfigura.Selected(CPCONFIG_GERA_LOTE_AUTOMATICO) = True Then
            gobjCP.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO
        Else
            gobjCP.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
        End If
    
'    End If
    
    'Grava na Tabela  "CPConfig" as Configurações
    lErro = gobjCP.Gravar()
    If lErro <> SUCESSO Then Error 41879
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 41879 'Tratado na Rotina Chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155624)
    End Select

    Exit Function
    
End Function

Public Sub ListaConfigura_ItemCheck(Item As Integer)
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONFIGURA_CP
    Set Form_Load_Ocx = Me
    Caption = "Configuração do Módulo de Contas a Pagar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConfiguraCP"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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



Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

