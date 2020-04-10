VERSION 5.00
Begin VB.Form bematech_teste 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "bematech_teste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Integer, sModelo As String


    i = Bematech_FI_AbrePortaSerial()
    MsgBox (CStr(i))
    i = Bematech_FI_VerificaImpressoraLigada
    MsgBox (CStr(i))
    i = Bematech_FI_FechaPortaSerial()
    MsgBox (CStr(i))
End Sub
