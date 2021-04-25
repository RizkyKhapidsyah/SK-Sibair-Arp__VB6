VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SibairArp"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "SibairArp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Height          =   195
      Left            =   3120
      TabIndex        =   11
      ToolTipText     =   "Cocher pour enregistrer les résultats dans un fichier"
      Top             =   360
      Width           =   255
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   4455
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   3765
      ItemData        =   "SibairArp.frx":1CFA
      Left            =   0
      List            =   "SibairArp.frx":1CFC
      TabIndex        =   9
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox ipa2 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Text            =   "254"
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox ipa1 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   "52"
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox ipd4 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Text            =   "1"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox ipd3 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "52"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox ipd2 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "6"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox ipd1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "152"
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Ip fin"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Ip départ"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim ip1 As Integer
Dim ip2 As Integer
Dim ECHO As ICMP_ECHO_REPLY
Dim cpt As Integer
Dim num_fichier As Integer
Dim ligne As String

List1.Clear
Form1.MousePointer = vbHourglass
cpt = 0
    
If Check1.Value = vbChecked Then

    On Error GoTo Err_Fichier
    
    num_fichier = FreeFile
    Open "arp.log" For Output As #num_fichier
End If
    
For ip1 = Val(ipd3) To Val(ipa1)
    For ip2 = Val(ipd4) To Val(ipa2)
        ip_ping = ipd1 & "." & ipd2 & "." & Trim(Str(ip1)) & "." & Trim(Str(ip2))
        StatusBar1.SimpleText = ip_ping
        DoEvents
        Call Ping(ip_ping, ECHO, 30)
        DoEvents
        If ECHO.status = 0 Then
            mac_adresse = Recherche_Mac(ip_ping)
            List1.AddItem (ip_ping & ";" & ECHO.RoundTripTime & "ms;" & mac_adresse)
                If Check1.Value = vbChecked Then
                    Print #num_fichier, ip_ping & ";" & mac_adresse
                End If
            cpt = cpt + 1
            DoEvents
            List1.Refresh
        End If
    Next ip2
Next ip1

Form1.MousePointer = vbDefault
StatusBar1.SimpleText = "Terminé avec " & Str(cpt) & " trouvé(s)"

If Check1.Value = vbChecked Then
    Close #num_fichier
End If
Err_Fichier:
End Sub

Private Sub ipd1_GotFocus()
    ipd1.SelStart = 0
    ipd1.SelLength = Len(ipd1)
End Sub
Private Sub ipd2_GotFocus()
    ipd2.SelStart = 0
    ipd2.SelLength = Len(ipd2)
End Sub
Private Sub ipd3_GotFocus()
    ipd3.SelStart = 0
    ipd3.SelLength = Len(ipd3)
End Sub
Private Sub ipd4_GotFocus()
    ipd4.SelStart = 0
    ipd4.SelLength = Len(ipd4)
End Sub
Private Sub ipa1_GotFocus()
    ipa1.SelStart = 0
    ipa1.SelLength = Len(ipa1)
End Sub
Private Sub ipa2_GotFocus()
    ipa2.SelStart = 0
    ipa2.SelLength = Len(ipa2)
End Sub
