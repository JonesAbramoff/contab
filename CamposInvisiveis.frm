VERSION 5.00
Begin VB.Form CamposInvisiveis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Campos Não Visiveis"
   ClientHeight    =   1665
   ClientLeft      =   5820
   ClientTop       =   5490
   ClientWidth     =   2670
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2670
   Begin VB.ListBox ListaCamposInvisiveis 
      Height          =   1425
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   2400
   End
End
Attribute VB_Name = "CamposInvisiveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Dim Formato As RECT

    Call GetWindowRect(Me.hWnd, Formato)
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, Formato.left, Formato.top, Formato.right - Formato.left, Formato.bottom - Formato.top, SWP_SHOWWINDOW)
    Set gobjCamposInvisiveis = Me


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set gobjCamposInvisiveis = Nothing
End Sub

Private Sub ListaCamposInvisiveis_DblClick()

Dim objControle As Object
Dim sNome As String
Dim iIndice As Integer

On Error GoTo Erro_ListaCamposInvisiveis_DblClick

    For Each objControle In gobjTelaAtiva.Controls
        If Not (TypeName(objControle) = "Menu") And Not (TypeName(objControle) = "Timer") And Not (TypeName(objControle) = "Line") And Not (TypeName(objControle) = "CommonDialog") And Not (TypeName(objControle) = "Image") Then
            
            sNome = objControle.Name
            If objControle.Index > -1 Then sNome = sNome & "(" & objControle.Index & ")"
                                    
            If sNome = ListaCamposInvisiveis.Text Then
                Call Propriedades.ComboCampos_Seleciona(objControle)
                Exit For
            End If
        End If
    Next
        
    Exit Sub
    
Erro_ListaCamposInvisiveis_DblClick:

    Select Case Err
    
        Case 343
            Resume Next
            
        Case Else
        
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 144138)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub Carrega_Campos_Invisiveis()

Dim iIndice As Integer
Dim objControle As Object
Dim sNome As String

On Error GoTo Erro_Carrega_Campos_Invisiveis

    ListaCamposInvisiveis.Clear

    For Each objControle In gobjTelaAtiva.Controls
        If Not (TypeName(objControle) = "Timer") And Not (TypeName(objControle) = "Line") And Not (TypeName(objControle) = "CommonDialog") And Not (TypeName(objControle) = "Image") Then
            If Not (TypeName(objControle) = "Menu") And objControle.Visible = False Then
                iIndice = -1
                iIndice = objControle.Index
                sNome = objControle.Name
                If iIndice > -1 Then sNome = sNome & "(" & CStr(iIndice) & ")"
                
                ListaCamposInvisiveis.AddItem sNome
                ListaCamposInvisiveis.ItemData(ListaCamposInvisiveis.NewIndex) = iIndice
                
            End If
        End If
    Next

    Exit Sub

Erro_Carrega_Campos_Invisiveis:

    Select Case Err
    
        Case 343
            Resume Next
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 144139)
            
    End Select
    
    Exit Sub

End Sub
