VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ErrorOccured(lngErrCode As Long, strErrDesc As String)

Private m_oHTTP As Object
Private m_lngLastErrNumber As Long
Private m_strLastErrDesc As String

Private m_colRequestHeaders As Collection
Private m_colResponseHeaders As Collection

Private Enum enHTTPMethod
    mGet
    mPOST
End Enum

Private Sub Class_Initialize()
    'Initialise this object
    ClearError
    Set m_colRequestHeaders = New Collection
    Set m_colResponseHeaders = New Collection
    Set m_oHTTP = fCreateObject("MSXML2.XMLHTTP.4.0")
    If m_oHTTP Is Nothing Then
        Set m_oHTTP = fCreateObject("MSXML2.XMLHTTP.3.0")
        If m_oHTTP Is Nothing Then
            Set m_oHTTP = fCreateObject("Microsoft.XMLHTTP")
        End If
    End If
End Sub

Private Sub Class_Terminate()
    'Clean up this object
    Set m_oHTTP = Nothing
    ClearRequestHeaders
    ClearResponseHeaders
    Set m_colResponseHeaders = Nothing
    Set m_colRequestHeaders = Nothing
End Sub

Public Function AddRequestHeader(strHeaderName As String, strValue As String) As Header
    'Add a request header to the collection of requestheaders
    Dim oHeader As Header
    Set oHeader = New Header
    oHeader.HeaderName = strHeaderName
    oHeader.Value = strValue
    m_colRequestHeaders.Add oHeader
    Set AddRequestHeader = oHeader
End Function

Public Sub ClearRequestHeaders()
    'Clear all requestheaders
    While m_colRequestHeaders.Count > 0
        m_colRequestHeaders.Remove 1
    Wend
End Sub

Public Sub ClearResponseHeaders()
    'Clear all responseheaders
    While m_colResponseHeaders.Count > 0
        m_colResponseHeaders.Remove 1
    Wend
End Sub

Public Property Get RequestHeader(lIndex As Long) As Header
    'Return a requestheader by index
    On Error GoTo errHeader
    Set RequestHeader = m_colRequestHeaders(lIndex + 1)
    Exit Property
errHeader:
    Set RequestHeader = Nothing
End Property

Public Property Get RequestHeaderByName(strName) As Header
    'Return a requestheader by headername
    Dim oRes As Header
    Dim T As Long
    
    On Error GoTo errHeader
    Set oRes = Nothing
    For T = 1 To m_colRequestHeaders.Count
        If StrComp(m_colRequestHeaders(T).HeaderName, strName, vbTextCompare) = 0 Then
            Set oRes = m_colRequestHeaders(T)
            Exit For
        End If
    Next
    Set RequestHeaderByName = oRes
    
    Exit Property
errHeader:
    Set RequestHeaderByName = Nothing
End Property

Public Property Get RequestHeaderCount() As Long
    'Return requestheadercount
    RequestHeaderCount = m_colRequestHeaders.Count
End Property

Public Sub RemoveRequestHeader(lIndex As Long)
    'Remove a requestheader by index
    On Error GoTo errRemoveHeader
    m_colRequestHeaders.Remove lIndex + 1
    Exit Sub
errRemoveHeader:
End Sub

Public Property Get ResponseHeader(lIndex As Long) As Header
    'Return a responseheader by index
    On Error GoTo errHeader
    Set ResponseHeader = m_colResponseHeaders(lIndex + 1)
    Exit Property
errHeader:
    Set ResponseHeader = Nothing
End Property

Public Property Get ResponseHeaderByName(strName) As Header
    'Return a responseheader by name
    Dim oRes As Header
    Dim T As Long
    
    On Error GoTo errHeader
    Set oRes = Nothing
    For T = 1 To m_colResponseHeaders.Count
        If StrComp(m_colResponseHeaders(T).HeaderName, strName, vbTextCompare) = 0 Then
            Set oRes = m_colResponseHeaders(T)
            Exit For
        End If
    Next
    Set ResponseHeaderByName = oRes
    
    Exit Property
errHeader:
    Set ResponseHeaderByName = Nothing
End Property

Public Property Get ResponseHeaderCount() As Long
    'Return the responseheader count
    ResponseHeaderCount = m_colResponseHeaders.Count
End Property

Public Property Get XMLDoc() As Object
    'Return the response XML Document (if any)
    On Error GoTo errXML
    If Not m_oHTTP Is Nothing Then
        Set XMLDoc = m_oHTTP.ResponseXML
    Else
        Set XMLDoc = Nothing
    End If
    Exit Property
errXML:
    RaiseError Err.Number, Err.Description
End Property

Public Property Get Text() As String
    'Return the response Text (if any)
    On Error GoTo errText
    If Not m_oHTTP Is Nothing Then
        Text = m_oHTTP.ResponseText
    Else
        Text = ""
    End If
    Exit Property
errText:
    RaiseError Err.Number, Err.Description
End Property

Public Property Get Raw() As Variant
    'Return the raw respons
    On Error GoTo errRaw
    If Not m_oHTTP Is Nothing Then
        Raw = m_oHTTP.ResponseBody
    Else
        Set Raw = Nothing
    End If
    Exit Property
errRaw:
    RaiseError Err.Number, Err.Description
End Property

Public Function OpenURL(strURL As String, Optional strUser As String = vbNullString, Optional strPass As String = vbNullString, Optional vPostBody As Variant) As Long
    'Open any url and store results in this object
    'Parameters:
    '   strURL    : URL to open (examples: http://www.mysite.com
    '                                      http://www.mysite.com/myxmldoc.xml
    '                                      http://mysite.com/mylogo.jpg
    '                                      http://www.mysite.com/mypage.asp?qry=test&id=15
    '                           )
    '   strUser   : OPTIONAL. Pass this parameter to authenticate
    '   strPass   : OPTIONAL. Pass this parameter to authenticate
    '   vPostBody : OPTIONAL. If specified, this data will be posted
    'Returns:
    '   HTTP Status as long. 200 = OK. See http://www.w3.org/Protocols/rfc2616/rfc2616-sec10.html#sec10
    '   for more information
    
    Dim oHeader As Header
    Dim arrTMP() As String
    Dim arrVal() As String
    Dim T As Long
    
    On Error GoTo errOpenURL
    If Not m_oHTTP Is Nothing Then
        m_oHTTP.Open GetMethodString(IIf(IsMissing(vPostBody), mGet, mPOST)), strURL, False, strUser, strPass
        
        'Add requestheaders
        For Each oHeader In m_colRequestHeaders
            m_oHTTP.setRequestHeader oHeader.HeaderName, oHeader.Value
        Next
        
        'Post data?
        If IsMissing(vPostBody) Then
            'No, just open the url
            m_oHTTP.Send
        Else
            'Post data according to vartype
            m_oHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            Select Case VarType(vPostBody)
                Case vbObject:
                    m_oHTTP.Send vPostBody.xml
                Case Else:
                    m_oHTTP.Send vPostBody
            End Select
        End If
        
        'Get response headers
        ClearResponseHeaders
        arrTMP = Split(m_oHTTP.getallresponseheaders, vbCrLf)
        For T = 0 To UBound(arrTMP)
            Set oHeader = New Header
            arrVal = Split(arrTMP(T), ":", 2)
            If UBound(arrVal) = 1 Then
                oHeader.HeaderName = Trim(arrVal(0))
                oHeader.Value = Trim(arrVal(1))
                m_colResponseHeaders.Add oHeader
            End If
        Next
        
        'Return HTTP status
        OpenURL = m_oHTTP.Status
    End If
    Exit Function
errOpenURL:
    RaiseError Err.Number, Err.Description
End Function

Public Function SaveToFile(strFileName As String) As Boolean
    'Save received data to file
    Dim oStream As Object
    
    On Error GoTo errSaveToFile
    If Not m_oHTTP Is Nothing Then
        Set oStream = CreateObject("ADODB.Stream")
        With oStream
            .Charset = "iso-8859-1"
            .Type = 1 ' adTypeBinary
            .Open
            .Write Me.Raw
            .SaveToFile strFileName, 2 'adSaveCreateOverwrite
            .Close
        End With
        Set oStream = Nothing
    End If
    SaveToFile = True
    Exit Function

errSaveToFile:
    RaiseError Err.Number, Err.Description
    SaveToFile = False
End Function

Private Function GetMethodString(eMethod As enHTTPMethod) As String
    'Map enum to string
    Select Case eMethod
        Case mPOST: GetMethodString = "POST"
        Case Else: GetMethodString = "GET"
    End Select
End Function

Private Function fCreateObject(strObject As String, Optional strServer As String = "") As Object
    'Try to create an object. When this fails, raise an error
    On Error GoTo errCreateObject
    Set fCreateObject = CreateObject(strObject, strServer)
    Exit Function
errCreateObject:
    RaiseError Err.Number, "Error creating object '" & strObject & "'"
    Set fCreateObject = Nothing
End Function

Public Property Get LastErrorDescription() As String
    'Return last error description
    LastErrorDescription = m_strLastErrDesc
End Property

Public Property Get LastErrorNumber() As Long
    'Return last error number
    LastErrorNumber = m_lngLastErrNumber
End Property

Public Sub ClearError()
    'Clear errors
    m_lngLastErrNumber = 0
    m_strLastErrDesc = ""
End Sub

Private Sub RaiseError(lngNumber As Long, strDesc As String)
    'Store last error and raise the ErrorOccured event
    m_lngLastErrNumber = lngNumber
    m_strLastErrDesc = Replace(strDesc, "send method", "OpenURL method", , , vbTextCompare)
    RaiseEvent ErrorOccured(m_lngLastErrNumber, m_strLastErrDesc)
End Sub
