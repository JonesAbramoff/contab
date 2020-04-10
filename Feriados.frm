VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{454464FA-6BBA-4224-B6CD-4A4CA1778A0F}#1.0#0"; "ADMCALENDAR.OCX"
Begin VB.Form Feriados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Feriados"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "Feriados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   555
      Left            =   5640
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   195
      Width           =   1170
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "Feriados.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Feriados.frx":02C8
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin AdmCalendar.Calendar Calendar1 
      Height          =   3075
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   5424
      Day             =   1
      Month           =   1
      Year            =   1999
      BeginProperty DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificação"
      Height          =   690
      Left            =   3465
      TabIndex        =   7
      Top             =   765
      Width           =   3360
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   675
         TabIndex        =   1
         Top             =   270
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoTrazer 
         Caption         =   "Trazer"
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
         Left            =   2175
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   8
         Top             =   330
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Atributos"
      Height          =   1590
      Left            =   3465
      TabIndex        =   6
      Top             =   1530
      Width           =   3360
      Begin VB.TextBox DescrFeriado 
         Height          =   345
         Left            =   405
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1020
         Width           =   2865
      End
      Begin VB.OptionButton OpcaoFeriado 
         Caption         =   "Feriado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         TabIndex        =   4
         Top             =   675
         Width           =   2250
      End
      Begin VB.OptionButton OpcaoDiaNormal 
         Caption         =   "Dia Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   3
         Top             =   285
         Width           =   1845
      End
   End
End
Attribute VB_Name = "Feriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'esta tela faz a manutencao da tabela de feriados
'todo feriado tem que ter uma descricao

'para setar ou obter a data corrente do calendario: Calendar1.Value
'para repintar: Calendar1.Refresh
'Calendar1.DayBold: para colocar ou retirar negrito

Dim iAlterado As Integer

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 43380
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 43380

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 160221)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazer_Click()

    If Len(Trim(Data.Text)) > 0 Then
        Calendar1.Value = DateValue(Data.Text)
        Call Traz_DataTela(Calendar1.Value, Calendar1.Value)
    End If
    
End Sub

Sub Traz_DataTela(ByVal OldDate As Date, ByVal NewDate As Date)

Dim lErro As Long
Dim objFeriado As New ClassFeriado

On Error GoTo Erro_Traz_DataTela

    'se trocou mes ou ano acertar os dias em negrito (feriados)
    If Month(OldDate) <> Month(NewDate) Or Year(OldDate) <> Year(NewDate) Then
        Call Traz_Todos(NewDate)
    End If
    
    objFeriado.iFilialEmpresa = giFilialEmpresa
    objFeriado.dtData = CDate(NewDate)

    'Lê a Data p/ Verificar se é Feriado
    lErro = CF("Feriado_Le",objFeriado)
    If lErro <> SUCESSO And lErro <> 43379 Then Error 43408

    If lErro = SUCESSO Then
        OpcaoFeriado.Value = True
        DescrFeriado.Text = objFeriado.sDescricao
    Else
        OpcaoDiaNormal.Value = True
        DescrFeriado.Text = ""
    End If
    
    iAlterado = 0

    Exit Sub
    
Erro_Traz_DataTela:

    Select Case Err
    
        Case 43408
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160222)

    End Select
    
    Exit Sub

End Sub

Sub Traz_Todos(dtData As Date)

Dim lErro As Long
Dim iIndice As Integer
Dim dtDataAux As Date
Dim iMaxDias As Integer
Dim sData As String
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim objFeriado As ClassFeriado
Dim colFeriados As New Collection

On Error GoTo Erro_Traz_Todos

    iMaxDias = MaxDayInMonth(Month(Data.Text), Year(Data.Text))
    
    sData = 1 & "/" & Month(Data.Text) & "/" & Year(Data.Text)
    
    dtDataInicial = CDate(sData)
    
    sData = iMaxDias & "/" & Month(Data.Text) & "/" & Year(Data.Text)
    
    dtDataFinal = CDate(sData)
    
    'Lê todos feriados do mês
    lErro = CF("Feriado_Le_Todos",dtDataInicial, dtDataFinal, colFeriados)
    If lErro <> SUCESSO Then Error 43413
  
    'Limpa os Feriados marcados anteriormente
    For iIndice = 0 To iMaxDias - 1
        Calendar1.DayBold(Day(dtDataInicial + iIndice)) = False
    Next
    
    If colFeriados.Count <> 0 Then
        For Each objFeriado In colFeriados
            Calendar1.DayBold(Day(objFeriado.dtData)) = True
        Next
    End If
    
    Calendar1.Refresh
    
    iAlterado = 0

    Exit Sub
    
Erro_Traz_Todos:

    Select Case Err
    
        Case 43413
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160223)

    End Select
    
    Exit Sub

End Sub

Private Sub Calendar1_DateChange(ByVal OldDate As Date, ByVal NewDate As Date)
    
    Data.Text = Format(NewDate, "dd/mm/yy")

    Call Traz_DataTela(OldDate, NewDate)
    
End Sub

Private Sub Calendar1_WillChangeDate(ByVal NewDate As Date, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Calendar1_WillChangeDate

    If iAlterado <> 0 Then
    
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then Error 43409

        Calendar1.Refresh
        
        iAlterado = 0
            
    End If
    
    Exit Sub
    
Erro_Calendar1_WillChangeDate:
    
    Select Case Err
    
        Case 43409
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160224)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub Data_LostFocus()
    
Dim lErro As Long

On Error GoTo Erro_Data_LostFocus

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then Error 43406
    
    Exit Sub

Erro_Data_LostFocus:
    
    Select Case Err
    
        Case 43406
            Data.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160225)
        
    End Select
    
    Exit Sub

End Sub

Private Sub DescrFeriado_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub DescrFeriado_LostFocus()

    If Len(Trim(DescrFeriado.Text)) <> 0 Then OpcaoFeriado.Value = True
    
End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Calendar1.Value = Date

    Data.Text = Format(Calendar1.Value, "dd/mm/yy")
    
    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160226)

    End Select
    
    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFeriado As New ClassFeriado

On Error GoTo Erro_Gravar_Registro

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then Error 43381
    
    objFeriado.iFilialEmpresa = giFilialEmpresa
    objFeriado.dtData = CDate(Data.Text)
    
    'Lê o Feriado para verificar se é uma exclusão ou gravação
    lErro = CF("Feriado_Le",objFeriado)
    If lErro <> SUCESSO And lErro <> 43379 Then Error 43410
    
    'Se achou o Feriado e a Opção for Dia Normal
    If lErro = SUCESSO And OpcaoDiaNormal.Value = True Then
    
        'Exclui o Feriado
        lErro = CF("Feriado_Exclui",objFeriado)
        If lErro <> SUCESSO Then Error 43411
        
        'Desmarca o Feriado que acabou de ser excluído
        Calendar1.DayBold(Day(Calendar1.Value)) = False
        
        'Limpa a Descrição do Feriado
        DescrFeriado.Text = ""
        
    Else
        'Verifica se a Descrição do Feriado foi preenchida
        If Len(Trim(DescrFeriado.Text)) = 0 Then Error 43382
        
        objFeriado.sDescricao = DescrFeriado.Text
        
        'Chama Feriado_Grava
        lErro = CF("Feriado_Grava",objFeriado)
        If lErro <> SUCESSO Then Error 43383
        
        'Marca o dia que acabou de ser gravado como Feriado
        Calendar1.DayBold(Day(Calendar1.Value)) = True
    
    End If
    
    iAlterado = 0

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    Select Case Err

        Case 43381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_FERIADO_NAO_PREENCHIDA", Err)
            Data.SetFocus
        
        Case 43382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_FERIADO_NAO_PREENCHIDA", Err)

        Case 43383, 43410, 43411
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 160227)

    End Select

    Exit Function

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode)
      
End Sub

Private Sub OpcaoDiaNormal_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OpcaoFeriado_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Function MaxDayInMonth(nMonth As Long, Optional nYear As Long = 0) As Long
    Select Case nMonth
        Case 9, 4, 6, 11    '30 days hath September,
                            'April, June, and November
            MaxDayInMonth = 30
        
        Case 2              'February -- check for leapyear
            'The full rule for leap years is that they occur in
            'every year divisible by four, except that they don't
            'occur in years divisible by 100, except that they
            '*do* in years divisible by 400.
            If (nYear Mod 4) = 0 Then
                If nYear Mod 100 = 0 Then
                    If nYear Mod 400 = 0 Then
                        MaxDayInMonth = 29
                    Else
                        MaxDayInMonth = 28
                    End If 'divisible by 400
                Else
                    MaxDayInMonth = 29
                End If 'divisible by 100
            Else
                MaxDayInMonth = 28
            End If 'divisible by 4
        
        Case Else           'All the rest have 31
            MaxDayInMonth = 31
    
    End Select
End Function 'MaxDayInMonth()

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

