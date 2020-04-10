VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl ExtratoBancarioCNABOcx 
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   ScaleHeight     =   1635
   ScaleWidth      =   5895
   Begin VB.ComboBox Banco 
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Text            =   "Banco"
      Top             =   330
      Width           =   2340
   End
   Begin VB.CommandButton BotaoProcurar 
      Caption         =   "Procurar..."
      Height          =   540
      Left            =   4725
      Picture         =   "ExtratoBancarioCNABOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   885
      Width           =   990
   End
   Begin VB.TextBox Arquivo 
      Height          =   300
      Left            =   1050
      TabIndex        =   1
      Top             =   1110
      Width           =   2940
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   4725
      Picture         =   "ExtratoBancarioCNABOcx.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   3585
      Picture         =   "ExtratoBancarioCNABOcx.ctx":0204
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   990
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4110
      Top             =   915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LabelBanco 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
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
      Left            =   345
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Arquivo:"
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
      Height          =   210
      Left            =   270
      TabIndex        =   6
      Top             =   1140
      Width           =   690
   End
End
Attribute VB_Name = "ExtratoBancarioCNABOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoBancos As AdmEvento
Attribute objEventoBancos.VB_VarHelpID = -1

Private Sub BotaoOK_Click()
Dim lErro As Long
Dim iBanco As Integer
Dim sTeste As String
Dim iTam As Integer, sNomeArq As String, sRegistro As String

On Error GoTo Erro_BotaoOK_Click

    'Verifica se a Combo Banco foi preenchida. Se nao foi, erro
    If Banco.Text = "" Then Error 22021
    'Verifica se TextBox Arquivo foi preenchido. Se nao foi, erro
    If Arquivo.Text = "" Then Error 22022

    If Banco.Text <> "" Then
       iBanco = Codigo_Extrai(Banco.Text)
    End If

    iTam = 0
    sNomeArq = Arquivo.Text
    
    'Abre o arquivo de retorno
    Open sNomeArq For Input As #2
    
    'Até chegar ao fim do arquivo
    Do While Not EOF(2)
    
        'Busca o próximo registro do arquivo (na 1a vez vai ser o de header)
        Line Input #2, sRegistro
        
        iTam = Len(sRegistro)
        Exit Do
    
    Loop
    
    Close #2
    
    If iTam = 240 Then
    
        lErro = CF("ExtratoCNAB_Importar", giFilialEmpresa, iBanco, sNomeArq)
        If lErro <> SUCESSO Then Error 22023
    
        Call MsgBox("Arquivo Processado!", , "Corporator")
    
    Else 'layout antigo de 200 posições
    
        Call Chama_Tela("ExtratoBancarioCNAB2", iBanco, sNomeArq)
        
    End If

    Unload Me
    
    Exit Sub

Erro_BotaoOK_Click:

    Select Case Err

        Case 22021
             lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_PREENCHIDO", Err)

        Case 22022
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", Err)

        Case 22023

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159897)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoBancos = New AdmEvento

    'Carrega  a combo de bancos
    lErro = Preenche_Combo_Bancos()
    If lErro <> SUCESSO Then Error 22019

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 22019

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159898)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoBancos = Nothing
    
End Sub

Private Sub LabelBanco_Click()

Dim objBanco As New ClassBanco
Dim colSelecao As Collection

    If Len(Banco.Text) = 0 Then
        objBanco.iCodBanco = 0
    Else
        objBanco.iCodBanco = Codigo_Extrai(Banco.Text)
    End If

    Call Chama_Tela("BancosLista", colSelecao, objBanco, objEventoBancos)

End Sub

Private Function Preenche_Combo_Bancos() As Long

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Combo_Bancos

    'leitura dos codigos e descricoes das ListaCodConta de venda no BD
    lErro = CF("Cod_Nomes_Le", "Bancos", "CodBanco", "NomeReduzido", STRING_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then Error 22020

   'preenche ComboBox com código e nome dos CodBancos
    For iIndice = 1 To colCodigoNome.Count
        Set objCodigoNome = colCodigoNome(iIndice)
        Banco.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Banco.ItemData(Banco.NewIndex) = objCodigoNome.iCodigo
    Next

    'Seleciona uma dos Bancos
    Banco.Text = Banco.List(PRIMEIRA_CONTA)

    Preenche_Combo_Bancos = SUCESSO

    Exit Function

Erro_Preenche_Combo_Bancos:

    Preenche_Combo_Bancos = Err

    Select Case Err

        Case 22020

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159899)

    End Select

    Exit Function

End Function

Private Sub BotaoCancelar_Click()

    Unload Me

End Sub

Private Sub BotaoProcurar_Click()

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo Erro_BotaoProcurar_Click
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    Arquivo.Text = CommonDialog1.FileName
    Exit Sub

Erro_BotaoProcurar_Click:
    'User pressed the Cancel button
    Exit Sub
End Sub

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EXTRATO_BANCARIO_CNAB
    Set Form_Load_Ocx = Me
    Caption = "Recepção de Extrato Bancário"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ExtratoBancarioCNAB"
    
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

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Banco Then
            Call LabelBanco_Click
        End If
    
    End If
    
End Sub


Private Sub LabelBanco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelBanco, Source, X, Y)
End Sub

Private Sub LabelBanco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelBanco, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

