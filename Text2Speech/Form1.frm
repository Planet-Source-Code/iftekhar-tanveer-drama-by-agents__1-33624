VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chayan's Drama Creater By Agents"
   ClientHeight    =   4110
   ClientLeft      =   2625
   ClientTop       =   2775
   ClientWidth     =   7005
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7005
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form1.frx":044A
      Left            =   4650
      List            =   "Form1.frx":0457
      TabIndex        =   11
      Text            =   "Combo3"
      Top             =   2220
      Width           =   2280
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   405
      Left            =   4620
      TabIndex        =   10
      Top             =   1785
      Width           =   2310
   End
   Begin VB.Frame Frame1 
      Caption         =   "Charecters"
      Height          =   1545
      Left            =   90
      TabIndex        =   4
      Top             =   135
      Width           =   6345
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0476
         Left            =   885
         List            =   "Form1.frx":04FE
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   420
         Width           =   2355
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Start/Play"
         Height          =   495
         Left            =   3300
         TabIndex        =   8
         Top             =   345
         Width           =   2355
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show"
         Height          =   465
         Left            =   3300
         TabIndex        =   7
         Top             =   855
         Width           =   1185
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":06B9
         Left            =   915
         List            =   "Form1.frx":06C6
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   930
         Width           =   2280
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Hide"
         Height          =   465
         Left            =   4500
         TabIndex        =   5
         Top             =   855
         Width           =   1155
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   465
      Left            =   4620
      TabIndex        =   2
      Top             =   3105
      Width           =   2355
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Start Drama !"
      Default         =   -1  'True
      Height          =   495
      Left            =   4635
      TabIndex        =   1
      Top             =   2595
      Width           =   2355
   End
   Begin VB.TextBox txtText 
      Height          =   1950
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":06E0
      Top             =   2115
      Width           =   4470
   End
   Begin VB.Label Label5 
      Caption         =   "Script:"
      Height          =   240
      Left            =   105
      TabIndex        =   3
      Top             =   1830
      Width           =   1095
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   1560
      Top             =   3300
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public ActorVisible As Boolean

Private Sub cmbVoiceType_Click()
On Error Resume Next
 TextToSpeech1.CurrentMode = cmbVoiceType.ListIndex + 1
 
HScroll1.Max = TextToSpeech1.MaxSpeed
HScroll1.Min = TextToSpeech1.MinSpeed
HScroll2.Max = TextToSpeech1.MaxPitch
HScroll2.Min = TextToSpeech1.MinPitch
HScroll1.Value = TextToSpeech1.Speed
HScroll2.Value = TextToSpeech1.Pitch
End Sub

Private Sub cmdSpeek_Click()
TextToSpeech1.Speak txtText.Text
End Sub

Private Sub Combo3_Click()
If Combo3.Text = "Default" Then
txtText.Text = "greet | Hello! My name is <actorname>.| explain | I am a poet. And i've written a new poem. | write | writing writing; trying to write ... trying to write about a great fight.| pleased |Thank you for hearing my poem. I am going now | wave |Good bye. | hide|"
ElseIf Combo3.Text = "Drama 1" Then
txtText.Text = "announce|I like to play|move to 400,400||move to 00,400||pleased| Oh. I love Playing !!! | move to 40,40 ||move to 300,560||move to 100,20||move to 600,40||congratulate"
ElseIf Combo3.Text = "Drama 2" Then
txtText.Text = "pleased|i can speek bengali. AAmar Shhonaar Baangla, AAmi Tomaay Bhaalobashhee. Chheerodeen tomaar aakashh tomaar baataashh; aamaar praanee baajaay bnashhee.| pleased | This means, `My golden bangla, I love you. Your sky, Your air;  plays the flute in my heart'. This song was Written and Composed by Tagor. He got the noble prize in 1913. This is the national song of Bangladesh."
End If
End Sub

Private Sub Command1_Click()
Agent1.Characters(Combo2.Text).Show
ActorVisible = True
End Sub


Private Sub Command2_Click()
Agent1.Characters(Combo2.Text).Hide
ActorVisible = False
End Sub

Private Sub Command3_Click()
Agent1.Characters(Combo2.Text).Show
ActorVisible = True
Form2.Show
'Agent1.Characters(Combo2.Text).Play "Announce"
Agent1.Characters(Combo2.Text).MoveTo Form2.ScaleLeft, Form2.ScaleTop + 230
Agent1.Characters(Combo2.Text).GestureAt Form2.ScaleLeft + 230, Form2.ScaleTop + 200
Agent1.Characters(Combo2.Text).Speak "My boss - Iftekhar Tanveer made this program."
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim k As Long, kTmp As Long, sText As String
Dim p As Long, pt As String
Dim Actions() As String
sText = LCase(txtText.Text)
sText = Replace(sText, "<actorname>", Agent1.Characters(Combo2.Text).Name)
sText = Replace(sText, vbLf, "")
sText = Replace(sText, Chr(13), " ")
Actions() = Split(sText, "|")
For i = 0 To UBound(Actions)
    If (i And 1) = 1 Then
        If Actions(i) <> "" Then
            Agent1.Characters(Combo2.Text).Speak Actions(i)
        End If
    Else
        If Trim(Actions(i)) = "hide" Then Agent1.Characters(Combo2.Text).Hide: ActorVisible = False: GoTo 1
        If Trim(Actions(i)) = "show" Then Agent1.Characters(Combo2.Text).Show: ActorVisible = True: GoTo 1
        If Left(Trim(Actions(i)), 4) = "wait" Then
            k = Val(Mid(Trim(Actions(i)), 5))
            kTmp = Timer
            Do: DoEvents: Loop Until Timer >= kTmp + k
        End If
        
        If Left(Trim(Actions(i)), 7) = "move to" Then
            pt = Mid(Actions(i), 8)
            p = InStr(pt, ",")
            Agent1.Characters(Combo2.Text).MoveTo Left(pt, p), Mid(pt, p + 1)
        End If

        Agent1.Characters(Combo2.Text).Play Trim(Actions(i))
    End If
1: Next


End Sub

Private Sub Command5_Click()
On Error Resume Next
Agent1.Characters(Combo2.Text).Stop
End Sub

Private Sub Command6_Click()
On Error Resume Next
Agent1.Characters(Combo2.Text).Play (Combo1.List(Combo1.ListIndex))

End Sub

Private Sub Form_Load()
   Agent1.Characters.Load "peedy", "C:\WINDOWS\MSAGENT\CHARS\peedy.acs"
   Agent1.Characters.Load "genie", "C:\WINDOWS\MSAGENT\CHARS\genie.acs"
   Agent1.Characters.Load "merlin", "C:\WINDOWS\MSAGENT\CHARS\merlin.acs"
    ' This syntax will generate RequestStart and RequestComplete events.
    
Agent1.Characters("peedy").Show
ActorVisible = True
Combo2.ListIndex = 0
Combo3.ListIndex = 0
End Sub

Private Sub HScroll1_Change()
Label1.Caption = "Speed:"
TextToSpeech1.Speed = HScroll1.Value
Label4 = HScroll1.Value
End Sub




Private Sub HScroll2_Change()
TextToSpeech1.Pitch = HScroll2.Value
Label3 = HScroll2.Value
End Sub

Private Sub TextToSpeech1_AudioStart(ByVal hi As Long, ByVal lo As Long)
Command2.Enabled = True
End Sub

Private Sub TextToSpeech1_TextDataDone(ByVal hi As Long, ByVal lo As Long, ByVal Flags As Long)
Command2.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
If ActorVisible Then
Agent1.Characters("genie").Hide
Agent1.Characters("peedy").Hide
Agent1.Characters("merlin").Hide
l = Timer
Do: DoEvents:  Loop Until Timer >= l + 2
End If
End Sub
