Attribute VB_Name = "MdlGetSys"
Option Explicit

'--> Avaco Keyboard Diagnostic
'--> version 1.00
'--> Version Language : English
'--> By Agus Ramadhani
'--> avaco software
'--> http://avaco-software.tripod.com
'--> avaco@9cy.Com
'--> 2002-2003
'--> Don't forget to Vote :)

Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub

Public Sub MeOnTop(Form As Form)
    SetWindowPos Form.hWnd, -1, 0, 0, 0, 0, 1 Or 2
End Sub


Public Sub MeDown(Form As Form)
    SetWindowPos Form.hWnd, -2, 0, 0, 0, 0, 1 Or 2
End Sub

Public Function Get_KeyboardType() As String
    Select Case GetKeyboardType(0)
        Case 0: Get_KeyboardType = "Unknown / Not Specified"
        Case 1: Get_KeyboardType = "IBM PC/XT ( ) or compatible (83-key)"
        Case 2: Get_KeyboardType = "Olivetti ICO (102-key) keyboard"
        Case 3: Get_KeyboardType = "IBM PC/AT (84-key) or similar"
        Case 4: Get_KeyboardType = "IBM enhanced (101- or 102-key)"
        Case 5: Get_KeyboardType = "Nokia 1050 and similar"
        Case 6: Get_KeyboardType = "Nokia 9140 and similar"
        Case 7: Get_KeyboardType = "Japanese"
        Case Else: Get_KeyboardType = "Unknown / Not Specified"
    End Select
End Function

Public Function Get_KeyboardFuncKeys() As String
    Select Case GetKeyboardType(2)
        Case 0: Failed "GetKeyboardType"
        Case 1: Get_KeyboardFuncKeys = "10"
        Case 2: Get_KeyboardFuncKeys = "12/18"
        Case 3: Get_KeyboardFuncKeys = "10"
        Case 4: Get_KeyboardFuncKeys = "12"
        Case 5: Get_KeyboardFuncKeys = "10"
        Case 6: Get_KeyboardFuncKeys = "24"
        Case 7: Get_KeyboardFuncKeys = "10"
        Case Else: Get_KeyboardFuncKeys = "Hardware dependent and specified by the OEM"
    End Select
End Function

Public Function Get_KeyboardLayout() As String
    Dim KeyboardLayout As String * 9
    
    If GetKeyboardLayoutName(KeyboardLayout) = 0 Then
        Failed "GetKeyboardLayoutName"
    Else
        Get_KeyboardLayout = Fix_NullTermStr(KeyboardLayout)
    End If
End Function
Public Function LangIdent(strCode As String) As String
    Select Case strCode
        Case "00000000": LangIdent = "Language Neutral"
        Case "00000400": LangIdent = "Process Default Language"
        Case "00000436": LangIdent = "Afrikaans"
        Case "0000041c": LangIdent = "Albanian"
        Case "00000401": LangIdent = "Arabic (Saudi Arabia)"
        Case "00000801": LangIdent = "Arabic (Iraq)"
        Case "00000c01": LangIdent = "Arabic (Egypt)"
        Case "00001001": LangIdent = "Arabic (Libya)"
        Case "00001401": LangIdent = "Arabic (Algeria)"
        Case "00001801": LangIdent = "Arabic (Morocco)"
        Case "00001c01": LangIdent = "Arabic (Tunisia)"
        Case "00002001": LangIdent = "Arabic (Oman)"
        Case "00002401": LangIdent = "Arabic (Yemen)"
        Case "00002801": LangIdent = "Arabic (Syria)"
        Case "00002c01": LangIdent = "Arabic (Jordan)"
        Case "00003001": LangIdent = "Arabic (Lebanon)"
        Case "00003401": LangIdent = "Arabic (Kuwait)"
        Case "00003801": LangIdent = "Arabic (U.A.E.)"
        Case "00003c01": LangIdent = "Arabic (Bahrain)"
        Case "00004001": LangIdent = "Arabic (Qatar)"
        Case "0000042b": LangIdent = "Armenian"
        Case "0000044d": LangIdent = "Assamese"
        Case "0000042c": LangIdent = "Azeri (Latin)"
        Case "0000082c": LangIdent = "Azeri (Cyrillic)"
        Case "0000042d": LangIdent = "Basque"
        Case "00000423": LangIdent = "Belarussian"
        Case "00000445": LangIdent = "Bengali"
        Case "00000402": LangIdent = "Bulgarian"
        Case "00000455": LangIdent = "Burmese"
        Case "00000403": LangIdent = "Catalan"
        Case "00000404": LangIdent = "Chinese (Taiwan)"
        Case "00000804": LangIdent = "Chinese (PRC)"
        Case "00000c04": LangIdent = "Chinese (Hong Kong SAR, PRC)"
        Case "00001004": LangIdent = "Chinese (Singapore)"
        Case "00001404": LangIdent = "Chinese (Macau SAR)"
        Case "0000041a": LangIdent = "Croatian"
        Case "00000405": LangIdent = "Czech"
        Case "00000406": LangIdent = "Danish"
        Case "00000413": LangIdent = "Dutch (Netherlands)"
        Case "00000813": LangIdent = "Dutch (Belgium)"
        Case "00000409": LangIdent = "English (United States)"
        Case "00000809": LangIdent = "English (United Kingdom)"
        Case "00000c09": LangIdent = "English (Australian)"
        Case "00001009": LangIdent = "English (Canadian)"
        Case "00001409": LangIdent = "English (New Zealand)"
        Case "00001809": LangIdent = "English (Ireland)"
        Case "00001c09": LangIdent = "English (South Africa)"
        Case "00002009": LangIdent = "English (Jamaica)"
        Case "00002409": LangIdent = "English (Caribbean)"
        Case "00002809": LangIdent = "English (Belize)"
        Case "00002c09": LangIdent = "English (Trinidad)"
        Case "00003009": LangIdent = "English (Zimbabwe)"
        Case "00003409": LangIdent = "English (Philippines)"
        Case "00000425": LangIdent = "Estonian"
        Case "00000438": LangIdent = "Faeroese"
        Case "00000429": LangIdent = "Farsi"
        Case "0000040b": LangIdent = "Finnish"
        Case "0000040c": LangIdent = "French (Standard)"
        Case "0000080c": LangIdent = "French (Belgian)"
        Case "00000c0c": LangIdent = "French (Canadian)"
        Case "0000100c": LangIdent = "French (Switzerland)"
        Case "0000140c": LangIdent = "French (Luxembourg)"
        Case "0000180c": LangIdent = "French (Monaco)"
        Case "0000043c": LangIdent = "Gaelic - Scotland"
        Case "00000437": LangIdent = "Georgian"
        Case "00000407": LangIdent = "German (Standard)"
        Case "00000807": LangIdent = "German (Switzerland)"
        Case "00000c07": LangIdent = "German (Austria)"
        Case "00001007": LangIdent = "German (Luxembourg)"
        Case "00001407": LangIdent = "German (Liechtenstein)"
        Case "00000408": LangIdent = "Greek"
        Case "00000447": LangIdent = "Gujarati"
        Case "0000040d": LangIdent = "Hebrew"
        Case "00000439": LangIdent = "Hindi"
        Case "0000040e": LangIdent = "Hungarian"
        Case "0000040f": LangIdent = "Icelandic"
        Case "00000421": LangIdent = "Indonesian"
        Case "00000410": LangIdent = "Italian (Standard)"
        Case "00000810": LangIdent = "Italian (Switzerland)"
        Case "00000411": LangIdent = "Japanese"
        Case "0000044b": LangIdent = "Kannada"
        Case "00000860": LangIdent = "Kashmiri (India)"
        Case "0000043f": LangIdent = "Kazakh"
        Case "00000457": LangIdent = "Konkani"
        Case "00000412": LangIdent = "Korean"
        Case "00000812": LangIdent = "Korean (Johab)"
        Case "00000426": LangIdent = "Latvian"
        Case "00000427": LangIdent = "Lithuanian"
        Case "00000827": LangIdent = "Lithuanian (Classic)"
        Case "0000042f": LangIdent = "Macedonian"
        Case "0000043e": LangIdent = "Malay (Malaysian)"
        Case "0000083e": LangIdent = "Malay (Brunei Darussalam)"
        Case "0000044c": LangIdent = "Malayalam"
        Case "0000043a": LangIdent = "Maltese"
        Case "00000458": LangIdent = "Manipuri"
        Case "0000044e": LangIdent = "Marathi"
        Case "00000861": LangIdent = "Nepali (India)"
        Case "00000414": LangIdent = "Norwegian (Bokmal)"
        Case "00000814": LangIdent = "Norwegian (Nynorsk)"
        Case "00000448": LangIdent = "Oriya"
        Case "00000415": LangIdent = "Polish"
        Case "00000416": LangIdent = "Portuguese (Brazil)"
        Case "00000816": LangIdent = "Portuguese (Standard)"
        Case "00000446": LangIdent = "Punjabi"
        Case "00000417": LangIdent = "Raeto-Romance"
        Case "00000418": LangIdent = "Romanian"
        Case "00000818": LangIdent = "Romanian - Moldova"
        Case "00000419": LangIdent = "Russian"
        Case "00000819": LangIdent = "Russian - Moldova"
        Case "0000044f": LangIdent = "Sanskrit"
        Case "00000c1a": LangIdent = "Serbian (Cyrillic)"
        Case "0000081a": LangIdent = "Serbian (Latin)"
        Case "00000459": LangIdent = "Sindhi"
        Case "0000041b": LangIdent = "Slovak"
        Case "00000424": LangIdent = "Slovenian"
        Case "0000042e": LangIdent = "Sorbian"
        Case "0000040a": LangIdent = "Spanish (Traditional Sort)"
        Case "0000080a": LangIdent = "Spanish (Mexican)"
        Case "00000c0a": LangIdent = "Spanish (Modern Sort)"
        Case "0000100a": LangIdent = "Spanish (Guatemala)"
        Case "0000140a": LangIdent = "Spanish (Costa Rica)"
        Case "0000180a": LangIdent = "Spanish (Panama)"
        Case "00001c0a": LangIdent = "Spanish (Dominican Republic)"
        Case "0000200a": LangIdent = "Spanish (Venezuela)"
        Case "0000240a": LangIdent = "Spanish (Colombia)"
        Case "0000280a": LangIdent = "Spanish (Peru)"
        Case "00002c0a": LangIdent = "Spanish (Argentina)"
        Case "0000300a": LangIdent = "Spanish (Ecuador)"
        Case "0000340a": LangIdent = "Spanish (Chile)"
        Case "0000380a": LangIdent = "Spanish (Uruguay)"
        Case "00003c0a": LangIdent = "Spanish (Paraguay)"
        Case "0000400a": LangIdent = "Spanish (Bolivia)"
        Case "0000440a": LangIdent = "Spanish (El Salvador)"
        Case "0000480a": LangIdent = "Spanish (Honduras)"
        Case "00004c0a": LangIdent = "Spanish (Nicaragua)"
        Case "0000500a": LangIdent = "Spanish (Puerto Rico)"
        Case "00000430": LangIdent = "Sutu"
        Case "00000441": LangIdent = "Swahili (Kenya)"
        Case "0000041d": LangIdent = "Swedish"
        Case "0000081d": LangIdent = "Swedish (Finland)"
        Case "00000449": LangIdent = "Tamil"
        Case "00000444": LangIdent = "Tatar (Tatarstan)"
        Case "0000044a": LangIdent = "Telugu"
        Case "0000041e": LangIdent = "Thai"
        Case "00000431": LangIdent = "Tsonga"
        Case "0000041f": LangIdent = "Turkish"
        Case "00000422": LangIdent = "Ukrainian"
        Case "00000420": LangIdent = "Urdu (Pakistan)"
        Case "00000820": LangIdent = "Urdu (India)"
        Case "00000443": LangIdent = "Uzbek (Latin)"
        Case "00000843": LangIdent = "Uzbek (Cyrillic)"
        Case "0000042a": LangIdent = "Vietnamese"
        Case "00000434": LangIdent = "Xhosa"
        Case "0000043d": LangIdent = "Yiddish"
        Case "00000435": LangIdent = "Zulu"
    End Select
End Function

Public Sub Errors(lngError As Long, apiFunction As String, Optional errDescription As String, Optional NoMsgBox As Boolean)
    errDescription = space$(2048)
    Dim errmsg
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, 0&, lngError, 0, errDescription, 2048, 0&
    
    errDescription = Trim$(errDescription)
    If errDescription = "" Then errDescription = "No description available."
    
    If NoMsgBox = False Then
        If errmsg = True Then
            MsgBox apiFunction & vbCrLf & vbCrLf & errDescription, vbExclamation, "Error"
        End If
    End If
End Sub

Public Sub Failed(strAPI As String)
   Dim errmsg
    If errmsg = True Then
        If Err.LastDllError = 0 Then
            MsgBox strAPI & vbCrLf & vbCrLf & "Failed", vbExclamation, "Error"
        Else
            Errors Err.LastDllError, strAPI
        End If
    End If
End Sub

Public Function Fix_NullTermStr(strData As String) As String
    If strData = "" Then Exit Function
    If InStr(1, strData, Chr$(0)) = 0 Then
        Exit Function
    Else
        Fix_NullTermStr = Left$(strData, InStr(1, strData, Chr$(0)) - 1)
    End If
End Function

