VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ME!"
   ClientHeight    =   3270
   ClientLeft      =   3375
   ClientTop       =   4275
   ClientWidth     =   5460
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4425
      Top             =   2190
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   465
      Left            =   4170
      TabIndex        =   2
      Top             =   585
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   645
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   45
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   0
      Width           =   3315
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   $"Form2.frx":206A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   30
         TabIndex        =   1
         Top             =   1650
         Width           =   3195
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Picture1.Height = Picture1.Height + 1
If Picture1.Height = 183 Then Timer1.Enabled = False
End Sub
