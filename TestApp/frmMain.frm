VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQryGet 
      Caption         =   "GET Query"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoogleHTML 
      Caption         =   "Get Google HTML"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoogleLogo 
      Caption         =   "Download Google Logo"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   720
      Width           =   9375
   End
   Begin VB.CommandButton cmdRSS 
      Caption         =   "Get RSS Feed"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oWeb As WebComm.HTTP
Attribute oWeb.VB_VarHelpID = -1

Private Sub cmdGoogleHTML_Click()
    Dim lRes As Long
    Dim bRes As Boolean
    Const cTestURL = "http://www.google.com"
    
    ClearStatus                     'Reset status textbox
    oWeb.ClearError                 'Reset errors
    lRes = oWeb.OpenURL(cTestURL)   'Open the URL
    If lRes = 200 And oWeb.LastErrorNumber = 0 Then 'All OK?
        'Show HTTP Status
        AddStatus lRes & ": " & cTestURL
        
        'Show response headers
        ShowHeaders
        
        'Demonstrate use of .Text property
        AddStatus oWeb.Text
    End If
End Sub

Private Sub cmdGoogleLogo_Click()
    Dim lRes As Long
    Dim bRes As Boolean
    Dim strOutfile As String
    Const cTestURL = "http://www.google.com/images/logo.gif"
    
    ClearStatus                     'Reset status textbox
    oWeb.ClearError                 'Reset errors
    lRes = oWeb.OpenURL(cTestURL)   'Open the URL
    If lRes = 200 And oWeb.LastErrorNumber = 0 Then 'All OK?
        'Show HTTP Status
        AddStatus lRes & ": " & cTestURL
        
        'Show response headers
        ShowHeaders
        
        'Demonstrate use of .SaveToFile Method\
        strOutfile = App.Path & "\logo.gif"
        bRes = oWeb.SaveToFile(strOutfile)
        If bRes Then
            AddStatus "SaveToFile succeeded (" & strOutfile & ")"
        Else
            AddStatus "SaveToFile failed"
        End If
    End If
End Sub

Private Sub cmdQryGet_Click()
    Dim lRes As Long
    Const cTestURL = "http://www.google.com/search?hl=en&q=google+rocks"
    
    ClearStatus                     'Reset status textbox
    oWeb.ClearError                 'Reset errors
    lRes = oWeb.OpenURL(cTestURL)   'Open the URL
    If lRes = 200 And oWeb.LastErrorNumber = 0 Then 'All OK?
        'Show HTTP Status
        AddStatus lRes & ": " & cTestURL
        
        'Show response headers
        ShowHeaders
        
        'Show plain text in status box
        AddStatus oWeb.Text
    End If
End Sub

Private Sub cmdRSS_Click()
    Dim lRes As Long
    Const cTestURL = "http://msdn.microsoft.com/xml/rss.xml"
    
    ClearStatus                     'Reset status textbox
    oWeb.ClearError                 'Reset errors
    lRes = oWeb.OpenURL(cTestURL)   'Open the URL
    If lRes = 200 And oWeb.LastErrorNumber = 0 Then 'All OK?
        'Show HTTP Status
        AddStatus lRes & ": " & cTestURL
        
        'Show response headers
        ShowHeaders
        
        'Show plain text in status box
        AddStatus oWeb.Text
        
        'Demonstrate use of .XMLDoc property
        MsgBox oWeb.XMLDoc.SelectNodes("rss/channel/item").length & " items in RSS feed.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Form_Load()
    'Initialise our object
    Set oWeb = New WebComm.HTTP
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Free our object
    Set oWeb = Nothing
End Sub

Private Sub oWeb_ErrorOccured(lngErrCode As Long, strErrDesc As String)
    'Notify user of errors
    MsgBox "Error: " & lngErrCode & "-" & strErrDesc, vbCritical + vbApplicationModal
End Sub

Private Sub ShowHeaders()
    'Display response headers
    Dim T As Long
    
    AddStatus String(50, "=")
    For T = 0 To oWeb.ResponseHeaderCount - 1
        AddStatus oWeb.ResponseHeader(T).HeaderName & " : " & oWeb.ResponseHeader(T).Value
    Next
    AddStatus String(50, "=")
End Sub

Private Sub ClearStatus()
    'Clear status textbox
    txtStatus.Text = ""
End Sub

Private Sub AddStatus(strStatus As String)
    'Add status text to textbox
    txtStatus.Text = txtStatus.Text & strStatus & vbCrLf
End Sub
