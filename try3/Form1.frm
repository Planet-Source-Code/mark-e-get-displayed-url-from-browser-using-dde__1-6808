VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "The Url Armed Robber"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Internet Explorer Information"
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   4455
      Begin VB.TextBox TxtIETitle 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox TxtIEUrl 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label LabIEWinTitle 
         Caption         =   "Window Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LabIEURL 
         Caption         =   "URL : "
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.OptionButton OptIE 
      Caption         =   "Internet Explorer"
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.OptionButton OptNetscape 
      Caption         =   "Netscape"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Timer TimerCheckBrowsers 
      Interval        =   500
      Left            =   3360
      Top             =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Netscape Information"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4335
      Begin VB.TextBox TxtNSUrl 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox TxtNSTitle 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label LabNSURL 
         Caption         =   "URL : "
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.Label LabNSWinTitle 
         Caption         =   "Window Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox TxtDDE 
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub GetURLFromNetscape(url As String, title As String, _
    framename As String)

    On Error GoTo GUBErrHandler
    TxtDDE.LinkTopic = "NETSCAPE|WWW_GetWindowInfo"
    
    ' tell Netscape to send us
    ' name and title of the last active
    ' window or frame
    
    TxtDDE.LinkItem = &HFFFFFFFF
    TxtDDE.LinkMode = 2
    TxtDDE.LinkRequest
    
    ' parse out info given to us by Netscape in
    ' txtDDE.Text; should be in the form
    '        "URL","Page title","FrameName"
    
    Dim cc As Long, parms(3) As String, quoting As Boolean
    Dim thisParm As Integer, p As Long, c As Byte
    thisParm = 1
    quoting = False
    For i = 1 To Len(TxtDDE)
        c = Asc(Mid(TxtDDE, i, 1))
        Select Case c
            Case 34     ' quotation mark
                quoting = Not quoting
            Case 44     ' comma
                If Not quoting Then
                    thisParm = thisParm + 1
                    If thisParm > 3 Then Exit For
                End If
            Case Else
                If quoting Then
                    parms(thisParm) = parms(thisParm) & Chr(c)
                End If
        End Select
    Next i
    
    url = parms(1)
    title = parms(2)
    framename = parms(3)
    Exit Sub
    
GUBErrHandler:
    ' skip process if any errors occur, i.e., Netscape
    ' did not respond to DDE initiate event
    MsgBox "Browser not loaded."
    On Error GoTo 0
End Sub
         
Sub GetURLfromIE(url As String, title As String)

 On Error GoTo GUBErrHandler
    TxtDDE.LinkTopic = "iexplore|WWW_GetWindowInfo"
    
    ' tell ie to send us
    ' name and title of the last active
    ' window or frame
    
    TxtDDE.LinkItem = &HFFFFFFFF
    TxtDDE.LinkMode = 2
    TxtDDE.LinkRequest
    
    ' parse out info given to us by ie in
    ' txtDDE.Text; should be in the form
    '        "URL","Page title"
    
    Dim cc As Long, parms(2) As String, quoting As Boolean
    Dim thisParm As Integer, p As Long, c As Byte
    thisParm = 1
    quoting = False
    For i = 1 To Len(TxtDDE)
        c = Asc(Mid(TxtDDE, i, 1))
        Select Case c
            Case 34     ' quotation mark
                quoting = Not quoting
            Case 44     ' comma
                If Not quoting Then
                    thisParm = thisParm + 1
                    If thisParm > 2 Then Exit For
                End If
            Case Else
                If quoting Then
                    parms(thisParm) = parms(thisParm) & Chr(c)
                End If
        End Select
    Next i
    
    url = parms(1)
    title = parms(2)
    Exit Sub
    
GUBErrHandler:
    ' skip process if any errors occur, i.e., Netscape
    ' did not respond to DDE initiate event
    MsgBox "Browser not loaded."
    On Error GoTo 0
End Sub

Private Sub TimerCheckBrowsers_Timer()
Dim TheUrl As String, TheTitle As String, TheFrame As String
'TxtDDE.Text = ""
If OptIE.Value = True Then
Call GetURLfromIE(TheUrl, TheTitle)
TxtIETitle.Text = TheTitle
TxtIEUrl.Text = TheUrl
ElseIf OptNetscape.Value = True Then
Call GetURLFromNetscape(TheUrl, TheTitle, TheFrame)
TxtNSTitle.Text = TheTitle
TxtNSUrl.Text = TheUrl
End If

End Sub
