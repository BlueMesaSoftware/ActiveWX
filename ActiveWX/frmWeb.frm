VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmWeb 
   Caption         =   "World Wide Web Weather"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6000
      TabIndex        =   3
      Top             =   5250
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7140
      TabIndex        =   2
      Top             =   5250
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "World Wide Web Site HTML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4995
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   7965
      Begin RichTextLib.RichTextBox rtxHTML 
         Height          =   4365
         Left            =   270
         TabIndex        =   1
         Top             =   390
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7699
         _Version        =   327681
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmWeb.frx":0000
      End
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdStop_Click()
  
End Sub

Public Sub sbLoad_Form(vHTML As Variant)
   rtxHTML.Text = vHTML
   Me.Show vbModal
End Sub

Private Sub cmdSave_Click()
   mdiMain.ActiveForm.strHTML = rtxHTML.Text
End Sub
