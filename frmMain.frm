VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "URL and Email Stripper"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   5880
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.htm, *.html"
      DialogTitle     =   "Open HTML File"
      Filter          =   "*.htm, *.html"
   End
   Begin VB.CommandButton cmdYURL 
      Caption         =   "Get URL's from HTML file's"
      Height          =   855
      Left            =   2520
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdEmail 
      Caption         =   "Get email address's from HTML file's"
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox lstURL 
      Height          =   3180
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   4095
   End
   Begin VB.ListBox lstEmail 
      Height          =   3180
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblEmail 
      Caption         =   "Number of Email's: 0"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblURL 
      Caption         =   "Number of URL's: 0"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      Caption         =   "These two routines will strip all email address's and URL's from any html document and add them to a List Box."
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, St1, St2, tmpY As Integer
Public Sub StripURL(FilePath As String)
Dim tmpURL1, tmpURL2 As String

Open FilePath For Input As #1
Do Until EOF(1)
    Input #1, tmpURL1
    For X = 1 To Len(tmpURL1)
        tmpURL2 = Mid(tmpURL1, X, 7)
        If tmpURL2 = "http://" Then
            St1 = X
            tmpY = X
                For Y = 1 To Len(tmpURL1)
                    tmpURL2 = Mid(tmpURL1, tmpY, 1)
                        If tmpURL2 = Chr(34) Then
                            St2 = tmpY
                            lstURL.AddItem Mid(tmpURL1, St1, ((St2 - St1)))
                            Exit For
                        Else
                            tmpY = tmpY + 1
                        End If
                Next Y
        End If
    Next X
Loop

Close #1
    
End Sub


Private Sub cmdEmail_Click()
Dialog.ShowOpen
StripEmail (Dialog.FileName)
lblEmail = "Number of Email's: " & lstEmail.ListCount
End Sub

Private Sub cmdYURL_Click()
Dialog.ShowOpen
StripURL (Dialog.FileName)
lblURL = "Number of URL's: " & lstURL.ListCount
End Sub

Public Sub StripEmail(FilePath As String)
Dim tmpEmail1, tmpEmail2 As String

Open FilePath For Input As #1
Do Until EOF(1)
Input #1, tmpEmail1
For X = 1 To Len(tmpEmail1)
    tmpEmail2 = Mid(tmpEmail1, X, 7)
    If tmpEmail2 = "mailto:" Then
        St1 = X
        tmpY = X + 1
        For Y = 1 To Len(tmpEmail1)
            tmpEmail2 = Mid(tmpEmail1, tmpY, 1)
            If tmpEmail2 = Chr(34) Then
                St2 = tmpY
                tmpEmail2 = Mid(tmpEmail1, St1 + 7, ((St2 - St1) - 7))
                If (Left(tmpEmail2, 2) <> "//") And (Left(tmpEmail2, 1) <> " ") Then
                    lstEmail.AddItem tmpEmail2
                    Exit For
                End If
            End If
            tmpY = tmpY + 1
        Next Y
    End If
Next X
Loop
Close #1
End Sub
