VERSION 5.00
Begin VB.Form frmSniffer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password sniffer - Fernando Robles de Juan"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "frmSniffer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1260
      Left            =   15
      TabIndex        =   2
      Top             =   -60
      Width           =   4305
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1845
         TabIndex        =   0
         Top             =   255
         Width           =   2340
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   345
         Left            =   1680
         TabIndex        =   1
         Top             =   750
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   570
         Picture         =   "frmSniffer.frx":030A
         Top             =   645
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre de la ventana:"
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   315
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmSniffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Fernando Robles de Juan - Sevilla, España.
'
' Este programa sólo funciona bajo Win95/98 y Millenium
' This program only run under Win95/98 and Millenium
'

Private Sub cmdBuscar_Click()
    ' Busco la ventana principal
    ' Find principal window
    retHwnd = FindWindow(vbNullString, Text1.Text)
    ' Y a enumerar ventanitas....
    ' And start to enumerate child windows of principal window
    Call EnumChildWindows(retHwnd, AddressOf EnumWinProc, 0&)
End Sub

Private Sub Form_Load()
    Me.Icon = Image1.Picture
End Sub
