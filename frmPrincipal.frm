VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPrincipal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cria CNPJ Válido"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3030
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Copiar para área de transferência"
      Height          =   465
      Left            =   840
      TabIndex        =   2
      Top             =   1170
      Width           =   2175
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   345
      Left            =   300
      TabIndex        =   0
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   15
      Mask            =   "##.###.###/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblDigito 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   1
      Top             =   390
      Width           =   585
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Clipboard.SetText (MaskEdBox1.Text & "-" & lblDigito.Caption)
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
    Call ValidarCNPJ_CPF(Replace(Replace(MaskEdBox1.Text, ".", ""), "/", "") & "99")
End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ValidarCNPJ_CPF(Replace(Replace(MaskEdBox1.Text, ".", ""), "/", "") & "99")
End Sub
