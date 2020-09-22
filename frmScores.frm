VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Best scores"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   9
         Top             =   3360
         Width           =   3255
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Are you Nuts???"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Expert"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intermediate"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Beginner"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1455
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

Private Sub Form_Load()
varPath = App.Path & "\"
ReDim arrayScores(100)
If Dir(varPath & "scores.dat", vbNormal) <> "" Then
    f = FreeFile
    Open varPath & "scores.dat" For Input As #f
        varCount = 0
        Do Until EOF(f)
            varCount = varCount + 1
            If varCount > 100 Then Exit Do
            Line Input #f, arrayScores(varCount)
        Loop
    Close #f
    For a = 1 To 100
        If InStr(arrayScores(a), "BEGINNERNAME=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varBeginnerName = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
        If InStr(arrayScores(a), "BEGINNERSCORE=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varBeginnerScore = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
        
        If InStr(arrayScores(a), "INTERMEDIATENAME=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varIntermediateName = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
        If InStr(arrayScores(a), "INTERMEDIATESCORE=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varIntermediateScore = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
                
        If InStr(arrayScores(a), "EXPERTNAME=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varExpertName = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
        If InStr(arrayScores(a), "EXPERTSCORE=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varExpertScore = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
                
        If InStr(arrayScores(a), "NUTSNAME=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varNutsName = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
        If InStr(arrayScores(a), "NUTSSCORE=") <> 0 Then
            varEq = InStr(arrayScores(a), "=")
            varNutsScore = Right(arrayScores(a), Len(arrayScores(a)) - varEq)
        End If
    Next a
    If varBeginnerName <> "" Then
        Label2(0) = varBeginnerName & " : " & varBeginnerScore & " secondes"
    End If
    If varIntermediateName <> "" Then
        Label2(1) = varIntermediateName & " : " & varIntermediateScore & " secondes"
    End If
    If varExpertName <> "" Then
        Label2(2) = varExpertName & " : " & varExpertScore & " secondes"
    End If
    If varNutsName <> "" Then
        Label2(3) = varNutsName & " : " & varNutsScore & " secondes"
    End If
End If
End Sub
