Attribute VB_Name = "ZoteroCiteLinker"
' An MS Word macro that links Zotero citations to their bibliography.
' Uses custom bookmark naming based on ID, Title, Year, and Author.
' Example: Cite_id9323_A_deep_lea_2021_Jiang
' UnlinkCitations function updated to only delete bookmarks starting with "Cite_".
' Added helper functions to automatically detect and link URLs and email addresses.
' Added code for coloring the citation links.
'
' Author: Serdar Berat AYDIN (s.b.a@msn.com)
'
' --- Acknowledgments ---
' This compilation is primarily based on modifications to the ZoteroLinkCitation
' code by user 'altairwei' and '8gengen8'.
'
' The base versions can be accessed at the following links:
' https://github.com/altairwei/ZoteroLinkCitation
' https://github.com/8gengen8/ZoteroCitationLink-lgg

'-------------------------------------------------------------------
' GLOBAL DECLARATIONS & TYPES
'-------------------------------------------------------------------
Option Explicit

' Defines the structure for holding parsed citation data
Type Citation
    BibTitle As String      ' Title for searching in the bibliography
    BibId As String         ' Zotero ID for naming the bookmark
    BibAuthor As String     ' First author's last name
    BibYear As String       ' Publication year
    Start As Long           ' Start character position of the citation
    End As Long             ' End character position of the citation
End Type

Public myRibbon As IRibbonUI
Public gUserTextStyle As String
Public Const ZCL_VERSION As String = "2.5.0"

Private Const ZCL_TEMP_PREFIX As String = "ZoteroLinkCitationTempBookmark"

#If Mac Then
    Private Const ZCL_PATH_SEP As String = "/"
#Else
    Private Const ZCL_PATH_SEP As String = "\"
#End If
Private Const ZCL_REPO_URL As String = "https://github.com/sBaydin/ZoteroCiteLinker"
Private Const ZCL_RELEASES_URL As String = "https://github.com/sBaydin/ZoteroCiteLinker/releases/latest"
Private Const ZCL_LATEST_API As String = "https://api.github.com/repos/sBaydin/ZoteroCiteLinker/releases/latest"

Private p&, token, dic

Public Sub onRibbonLoad(ribbon As IRibbonUI)
    Set myRibbon = ribbon
End Sub

' Returns a Dictionary-compatible object. macOS has no Scripting.Dictionary,
' so this uses the bundled CustomDictionary class module there.
Private Function CreateDict() As Object
    #If Mac Then
        Set CreateDict = New CustomDictionary
    #Else
        Set CreateDict = CreateObject("Scripting.Dictionary")
    #End If
End Function

'-------------------------------------------------------------------
' VBA JSON PARSER
' Source: https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a
'-------------------------------------------------------------------

' Main entry point for the JSON parser.
' Initializes the parser state and starts parsing.
Private Function ParseJSON(json$, Optional key$ = "obj") As Object
    p = 1
    token = Tokenize(json)
    Set dic = CreateDict()
    If token(p) = "{" Then ParseObj key Else ParseArr key
    Set ParseJSON = dic
End Function

' Recursively parses a JSON object (elements within {}).
Private Function ParseObj(key$)
    Do: p = p + 1
        Select Case token(p)
            Case "]"
            Case "[":   ParseArr key
            Case "{"
                    If token(p + 1) = "}" Then
                        p = p + 1
                        dic.Add key, "null"
                    Else
                        ParseObj key
                    End If
            
            Case "}":   key = ReducePath(key): Exit Do
            Case ":":   key = key & "." & token(p - 1)
            Case ",":   key = ReducePath(key)
            Case Else: If token(p + 1) <> ":" Then dic.Add key, token(p)
        End Select
    Loop
End Function

' Recursively parses a JSON array (elements within []).
Private Function ParseArr(key$)
    Dim e&
    Do: p = p + 1
        Select Case token(p)
            Case "}"
            Case "{":   ParseObj key & ArrayID(e)
            Case "[":   ParseArr key & ArrayID(e)
            Case "]":   Exit Do
            Case ":":   key = key & ArrayID(e)
            Case ",":   e = e + 1
            Case Else: dic.Add key & ArrayID(e), token(p)
        End Select
    Loop
End Function

' Splits the raw JSON string into a VBA array of tokens (keys, values, operators).
Private Function Tokenize(s$)
    #If Mac Then
        Tokenize = TokenizeVBA(s)
    #Else
        Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
        Tokenize = RExtract(s, Pattern, True)
    #End If
End Function

' Pure VBA JSON tokenizer for macOS (no VBScript.RegExp dependency).
Private Function TokenizeVBA(s$)
    Dim tokens() As String
    Dim tCount As Long
    Dim i As Long
    Dim sLen As Long
    Dim ch As String
    Dim strVal As String
    Dim numStr As String
    Dim word As String

    sLen = Len(s)
    tCount = 0
    ReDim tokens(1 To sLen + 1)

    i = 1
    Do While i <= sLen
        ch = Mid(s, i, 1)

        If ch = " " Or ch = vbCr Or ch = vbLf Or ch = vbTab Then
            i = i + 1

        ElseIf ch = """" Then
            strVal = ""
            i = i + 1
            Do While i <= sLen
                ch = Mid(s, i, 1)
                If ch = "\" Then
                    strVal = strVal & ch
                    i = i + 1
                    If i <= sLen Then strVal = strVal & Mid(s, i, 1)
                ElseIf ch = """" Then
                    Exit Do
                Else
                    strVal = strVal & ch
                End If
                i = i + 1
            Loop
            i = i + 1
            tCount = tCount + 1
            tokens(tCount) = strVal

        ElseIf ch Like "[0-9]" Or _
               (ch = "-" And i + 1 <= sLen And Mid(s, i + 1, 1) Like "[0-9]") Or _
               (ch = "+" And i + 1 <= sLen And Mid(s, i + 1, 1) Like "[0-9]") Then
            numStr = ch
            i = i + 1
            Do While i <= sLen
                ch = Mid(s, i, 1)
                If ch Like "[0-9]" Or ch = "." Then
                    numStr = numStr & ch
                    i = i + 1
                ElseIf ch = "e" Or ch = "E" Then
                    numStr = numStr & ch
                    i = i + 1
                    If i <= sLen Then
                        ch = Mid(s, i, 1)
                        If ch = "+" Or ch = "-" Then
                            numStr = numStr & ch
                            i = i + 1
                        End If
                    End If
                Else
                    Exit Do
                End If
            Loop
            tCount = tCount + 1
            tokens(tCount) = numStr

        ElseIf ch Like "[A-Za-z_]" Then
            word = ch
            i = i + 1
            Do While i <= sLen
                ch = Mid(s, i, 1)
                If ch Like "[A-Za-z0-9_]" Then
                    word = word & ch
                    i = i + 1
                Else
                    Exit Do
                End If
            Loop
            tCount = tCount + 1
            tokens(tCount) = word

        Else
            tCount = tCount + 1
            tokens(tCount) = ch
            i = i + 1
        End If
    Loop

    If tCount > 0 Then
        ReDim Preserve tokens(1 To tCount)
    Else
        ReDim tokens(1 To 1)
        tokens(1) = ""
    End If

    TokenizeVBA = tokens
End Function

' Uses VBScript.RegExp to extract all matches based on a pattern.
Private Function RExtract(s$, Pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n
  Dim v()
  With CreateObject("VBScript.RegExp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .Test(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.Value
        If bGroup1Bias Then If Len(n.submatches(0)) Or n.Value = """""" Then v(c) = n.submatches(0)
      Next
    End If
  End With
  RExtract = v
End Function

' Formats an array index for the dictionary key (e.g., "(0)").
Private Function ArrayID$(e)
    ArrayID = "(" & e & ")"
End Function

' Moves one level up in the object path (e.g., "obj.data.item" -> "obj.data").
Private Function ReducePath$(key$)
    If InStr(key, ".") Then ReducePath = Left(key, InStrRev(key, ".") - 1) Else ReducePath = key
End Function

' Extracts all values from the dictionary whose keys match a pattern.
Function GetFilteredValues(dic, match)
    Dim c&, i&, v, w
    v = dic.Keys
    ReDim w(1 To dic.Count)
    For i = 0 To UBound(v)
        If v(i) Like match Then
            c = c + 1
            w(c) = dic.Item(v(i))
        End If
    Next
    ReDim Preserve w(1 To c)
    GetFilteredValues = w
End Function

' Extracts data into a 2D array based on a list of column patterns.
Function GetFilteredTable(dic, cols)
    Dim c&, i&, j&, v, w, z
    v = dic.Keys
    z = GetFilteredValues(dic, cols(0))
    ReDim w(1 To UBound(z), 1 To UBound(cols) + 1)
    For j = 1 To UBound(cols) + 1
         z = GetFilteredValues(dic, cols(j - 1))
         For i = 1 To UBound(z)
            w(i, j) = z(i)
         Next
    Next
    GetFilteredTable = w
End Function

'-------------------------------------------------------------------
' ZOTERO LINKER UTILITIES
'-------------------------------------------------------------------

' Checks if the zotero.exe process is running.
Private Function IsZoteroRunning() As Boolean
    #If Mac Then
        ' The macro only reads the document fields; no process check is needed.
        IsZoteroRunning = True
        Exit Function
    #End If

    Dim colProcess As Object
    Dim objProcess As Object
    Dim strQuery As String

    IsZoteroRunning = False

    strQuery = "SELECT * FROM Win32_Process WHERE Name = 'zotero.exe'"

    On Error Resume Next
    Set colProcess = GetObject("winmgmts:\\.\root\cimv2").ExecQuery(strQuery)

    If colProcess.Count > 0 Then
        IsZoteroRunning = True
    End If

    Set objProcess = Nothing
    Set colProcess = Nothing
    On Error GoTo 0
End Function

' Sanitizes a string to be a valid Word bookmark name.
Private Function CleanForBookmarkName(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    #If Mac Then
        For i = 1 To Len(s)
            ch = Mid(s, i, 1)
            If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") _
               Or (ch >= "0" And ch <= "9") Or ch = "_" Then
                result = result & ch
            Else
                result = result & "_"
            End If
        Next i
    #Else
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        With re
            .Global = True
            .Pattern = "[^A-Za-z0-9_]"
        End With
        result = re.Replace(s, "_")
    #End If

    CleanForBookmarkName = result

    If CleanForBookmarkName Like "[0-9]*" Then
        CleanForBookmarkName = "_" & CleanForBookmarkName
    End If
End Function

' Scrolls the active window to a range on Windows. Word for Mac lacks
' ScrollIntoView, so the caller's following Select is enough there.
Private Sub ScrollToRange(ByVal target As Range)
    #If Mac = False Then
        ActiveWindow.ScrollIntoView target, True
    #End If
End Sub

' Sorts a 1D variant array in place using the QuickSort algorithm.
Private Sub QuickSort(arr As Variant, inLow As Long, inHigh As Long)
    '--- Variable Declarations ---
    Dim pivot As String
    Dim tmpSwap As Variant
    Dim low As Long
    Dim high As Long
    
    '--- Initializations ---
    low = inLow
    high = inHigh
    pivot = arr((low + high) \ 2)
    
    While (low <= high)
        While (arr(low) < pivot And low < inHigh)
            low = low + 1
        Wend
        
        While (pivot < arr(high) And high > inLow)
            high = high - 1
        Wend
        
        If (low <= high) Then
            tmpSwap = arr(low)
            arr(low) = arr(high)
            arr(high) = tmpSwap
            low = low + 1
            high = high - 1
        End If
    Wend
    
    If (inLow < high) Then QuickSort arr, inLow, high
    If (low < inHigh) Then QuickSort arr, low, inHigh
End Sub

' Extracts and concatenates all 'ZOTERO_PREF' custom document properties.
Private Function ExtractZoteroPrefData() As String
    '--- Variable Declarations ---
    Dim prop As Variant
    Dim dict As Object
    Dim sortedKeys As Variant
    Dim concatenatedValues As String
    Dim key As Variant

    '--- Initializations ---
    Set dict = CreateDict()

    ' Find all Zotero preference properties
    For Each prop In ActiveDocument.CustomDocumentProperties
        If Left(prop.name, 11) = "ZOTERO_PREF" Then
            dict.Item(prop.name) = prop.Value
        End If
    Next prop
    
    sortedKeys = dict.Keys
    If dict.Count > 0 Then
        Call QuickSort(sortedKeys, LBound(sortedKeys), UBound(sortedKeys))
    End If

    For Each key In sortedKeys
        concatenatedValues = concatenatedValues & dict.Item(key)
    Next key

    ExtractZoteroPrefData = concatenatedValues
End Function

' Parses Zotero preferences from an XML data string.
Private Function GetZoteroPrefsFromXml(zoteroData As String) As Object
    Dim dict As Object
    Dim segments() As String

    #If Mac Then
        Dim styleIdFull As String
    #Else
        Dim xmlDoc As Object
        Dim dataElem As Object
        Dim sessionElem As Object
        Dim styleElem As Object
        Dim prefElem As Object
    #End If

    Set dict = CreateDict()

    #If Mac Then
        dict.Item("data-version") = GetXmlTagAttribute(zoteroData, "data", "data-version")
        dict.Item("zotero-version") = GetXmlTagAttribute(zoteroData, "data", "zotero-version")
        dict.Item("session-id") = GetXmlTagAttribute(zoteroData, "session", "id")

        styleIdFull = GetXmlTagAttribute(zoteroData, "style", "id")
        If Len(styleIdFull) > 0 Then
            segments = Split(styleIdFull, "/")
            dict.Item("style-id") = segments(UBound(segments))
            dict.Item("hasBibliography") = GetXmlTagAttribute(zoteroData, "style", "hasBibliography")
            dict.Item("bibliographyStyleHasBeenSet") = GetXmlTagAttribute(zoteroData, "style", "bibliographyStyleHasBeenSet")
        End If

        dict.Item("pref-fieldType") = GetXmlPrefValue(zoteroData, "fieldType")
    #Else
        Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
        xmlDoc.Async = False
        xmlDoc.LoadXML zoteroData

        If xmlDoc.ParseError.ErrorCode <> 0 Then
            MsgBox "XML Parse Error: " & xmlDoc.ParseError.Reason
            Set GetZoteroPrefsFromXml = dict
            Exit Function
        End If

        Set dataElem = xmlDoc.SelectSingleNode("//data")
        If Not dataElem Is Nothing Then
            dict.Item("data-version") = dataElem.getAttribute("data-version")
            dict.Item("zotero-version") = dataElem.getAttribute("zotero-version")
        End If

        Set sessionElem = xmlDoc.SelectSingleNode("//session")
        If Not sessionElem Is Nothing Then
            dict.Item("session-id") = sessionElem.getAttribute("id")
        End If

        Set styleElem = xmlDoc.SelectSingleNode("//style")
        If Not styleElem Is Nothing Then
            segments = Split(styleElem.getAttribute("id"), "/")
            dict.Item("style-id") = segments(UBound(segments))
            dict.Item("hasBibliography") = styleElem.getAttribute("hasBibliography")
            dict.Item("bibliographyStyleHasBeenSet") = styleElem.getAttribute("bibliographyStyleHasBeenSet")
        End If

        Set prefElem = xmlDoc.SelectSingleNode("//prefs/pref[@name='fieldType']")
        If Not prefElem Is Nothing Then
            dict.Item("pref-fieldType") = prefElem.getAttribute("value")
        End If
    #End If

    Set GetZoteroPrefsFromXml = dict
End Function

' Finds the value of an attribute in the first matching XML tag (macOS path).
Private Function GetXmlTagAttribute(ByVal xml As String, ByVal tagName As String, ByVal attrName As String) As String
    Dim pos As Long
    pos = InStr(1, xml, "<" & tagName, vbTextCompare)
    If pos = 0 Then Exit Function

    Dim tagEnd As Long
    tagEnd = InStr(pos, xml, ">")
    If tagEnd = 0 Then Exit Function

    Dim tagContent As String
    tagContent = Mid(xml, pos, tagEnd - pos + 1)

    Dim attrPos As Long
    attrPos = InStr(1, tagContent, attrName & "=""", vbTextCompare)
    If attrPos = 0 Then Exit Function

    Dim valueStart As Long
    valueStart = attrPos + Len(attrName) + 2
    Dim valueEnd As Long
    valueEnd = InStr(valueStart, tagContent, """")
    If valueEnd = 0 Then Exit Function

    GetXmlTagAttribute = Mid(tagContent, valueStart, valueEnd - valueStart)
End Function

' Finds the value attribute of a named <pref> tag (macOS path).
Private Function GetXmlPrefValue(ByVal xml As String, ByVal prefName As String) As String
    Dim searchPattern As String
    searchPattern = "name=""" & prefName & """"

    Dim pos As Long
    pos = 1
    Do
        pos = InStr(pos, xml, "<pref", vbTextCompare)
        If pos = 0 Then Exit Function

        Dim tagEnd As Long
        tagEnd = InStr(pos, xml, ">")
        If tagEnd = 0 Then Exit Function

        Dim tagContent As String
        tagContent = Mid(xml, pos, tagEnd - pos + 1)

        If InStr(1, tagContent, searchPattern, vbTextCompare) > 0 Then
            Dim attrPos As Long
            attrPos = InStr(1, tagContent, "value=""", vbTextCompare)
            If attrPos = 0 Then Exit Function

            Dim valueStart As Long
            valueStart = attrPos + 7
            Dim valueEnd As Long
            valueEnd = InStr(valueStart, tagContent, """")
            If valueEnd = 0 Then Exit Function

            GetXmlPrefValue = Mid(tagContent, valueStart, valueEnd - valueStart)
            Exit Function
        End If

        pos = tagEnd + 1
    Loop
End Function

' Parses Zotero preferences from a JSON data string.
Private Function GetZoteroPrefsFromJson(zoteroData As String) As Object
    Dim jsonObj As Object
    Dim dict As Object
    Dim segments() As String

    Set jsonObj = ParseJSON(Trim(zoteroData), "prefs")
    Set dict = CreateDict()

    dict.Item("data-version") = jsonObj.Item("prefs.dataVersion")
    dict.Item("zotero-version") = jsonObj.Item("prefs.zoteroVersion")
    dict.Item("session-id") = jsonObj.Item("prefs.sessionID")

    segments = Split(jsonObj.Item("prefs.style.styleID"), "/")
    dict.Item("style-id") = segments(UBound(segments))
    dict.Item("hasBibliography") = jsonObj.Item("prefs.style.hasBibliography")
    dict.Item("bibliographyStyleHasBeenSet") = jsonObj.Item("prefs.style.bibliographyStyleHasBeenSet")

    dict.Item("pref-fieldType") = jsonObj.Item("prefs.prefs.fieldType")

    Set GetZoteroPrefsFromJson = dict
End Function

' Detects Zotero pref format (XML/JSON) and calls the appropriate parser.
Private Function GetZoteroPrefs() As Object
    Dim zoteroData As String
    Dim firstChar As String

    zoteroData = ExtractZoteroPrefData()
    firstChar = Left(Trim(zoteroData), 1)

    Select Case firstChar
        Case "<"
            Set GetZoteroPrefs = GetZoteroPrefsFromXml(zoteroData)
        Case "{", "["
            Set GetZoteroPrefs = GetZoteroPrefsFromJson(zoteroData)
        Case Else
            Err.Raise vbObjectError + 1, "GetZoteroPrefs", "Can not find Zotero Preference in this document."
    End Select

End Function

' Removes a specified list of HTML tags from a string.
Private Function RemoveSpecifiedHtmlTags(inputString As String, tagsToRemove As Variant) As String
    Dim tag As Variant

    #If Mac Then
        Dim pos As Long, endPos As Long
        For Each tag In tagsToRemove
            inputString = Replace(inputString, "</" & tag & ">", "", Compare:=vbTextCompare)
            inputString = Replace(inputString, "<" & tag & ">", "", Compare:=vbTextCompare)
            Do
                pos = InStr(1, inputString, "<" & tag & " ", vbTextCompare)
                If pos = 0 Then Exit Do
                endPos = InStr(pos, inputString, ">")
                If endPos = 0 Then Exit Do
                inputString = Left(inputString, pos - 1) & Mid(inputString, endPos + 1)
            Loop
        Next tag
    #Else
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")

        For Each tag In tagsToRemove
            With regex
                .Global = True
                .IgnoreCase = True
                .Pattern = "</?" & tag & ".*?>"
                inputString = .Replace(inputString, "")
            End With
        Next tag
    #End If

    RemoveSpecifiedHtmlTags = inputString
End Function

' Removes common formatting HTML tags (i, sub, sup) from a string.
Private Function RemoveHtmlTags(inputString As String) As String
    Dim tagsToRemove() As Variant
    tagsToRemove = Array("i", "sub", "sup")
    RemoveHtmlTags = RemoveSpecifiedHtmlTags(inputString, tagsToRemove)
End Function

' Parses the JSON data from a Zotero CSL_CITATION field code.
Private Function ParseCSLCitationJson(ByVal code As String) As Object
    Dim jsonObj As Object
    Set jsonObj = ParseJSON(Trim(Replace(code, "ADDIN ZOTERO_ITEM CSL_CITATION", "")), "CSL")
    Set ParseCSLCitationJson = jsonObj
End Function

' Checks if a named style exists in the document.
Function StyleExists(ByVal styleToTest As String, ByVal docToTest As Word.Document) As Boolean
    Dim testStyle As Word.style
    On Error Resume Next
    Set testStyle = docToTest.styles(styleToTest)
    On Error GoTo 0
    StyleExists = Not testStyle Is Nothing
End Function

'-------------------------------------------------------------------
' CITATION STYLE HANDLERS
'-------------------------------------------------------------------

' Extracts citations for author-year styles, e.g., (Dweba et al., 2017)
Private Sub ExtractAuthorYearCitations(field As field, ByRef citations() As Citation, _
        Optional onlyYear As Boolean = False, Optional multiRefCommaSep As Boolean = True)
    
    '--- Variable Declarations ---
    Dim targetRange As Range, charRange As Range
    Dim rangeIndex As Long
    Dim inCitation As Boolean, nComma As Integer, beginYear As Boolean
    Dim json As Object
    Dim startChar As Long, endChar As Long
    Dim i As Long

    '--- Initializations ---
    Set targetRange = field.result
    Set charRange = targetRange.Duplicate
    charRange.Collapse wdCollapseStart
    Set json = ParseCSLCitationJson(field.code)

    ReDim citations(0)
    rangeIndex = -1
    
    inCitation = False
    nComma = 0
    beginYear = False
    
    For i = 1 To targetRange.Characters.Count
        charRange.Start = targetRange.Start + i - 1
        charRange.End = targetRange.Start + i

        ' Start of full author-year citation
        If charRange.Text = "(" And Not onlyYear Then
            inCitation = True
            startChar = charRange.Start + 1

        ' Start of year citation
        ElseIf charRange.Text Like "[0-9]" Then
            beginYear = True

            If onlyYear And Not inCitation Then
                inCitation = True
                startChar = charRange.Start
            End If

        ' Check multiple citations of same author
        ElseIf multiRefCommaSep And charRange.Text = "," Then
            nComma = nComma + 1
            If nComma > 1 And beginYear Then
                GoTo CreateCitationObject
            End If

        ' End of citation
        ElseIf charRange.Text = ";" Or charRange.Text = ")" Then
            beginYear = False
            If multiRefCommaSep Then nComma = 0

CreateCitationObject:
            If inCitation Then
                endChar = charRange.Start

                rangeIndex = rangeIndex + 1
                If rangeIndex > UBound(citations) Then
                    ReDim Preserve citations(0 To rangeIndex)
                End If

                citations(rangeIndex).Start = startChar
                citations(rangeIndex).End = endChar
                citations(rangeIndex).BibTitle = RemoveHtmlTags( _
                    json.Item("CSL.citationItems(" & rangeIndex & ").itemData.title"))
                citations(rangeIndex).BibId = json.Item("CSL.citationItems(" & rangeIndex & ").itemData.id")
                
                Dim yearVal As String
                On Error Resume Next
                citations(rangeIndex).BibAuthor = json.Item("CSL.citationItems(" & rangeIndex & ").itemData.author(0).family")
                
                yearVal = json.Item("CSL.citationItems(" & rangeIndex & ").itemData.issued.dateparts(0)(0)")
                If yearVal = "" Then
                    yearVal = json.Item("CSL.citationItems(" & rangeIndex & ").itemData.issued.date-parts(0)(0)")
                End If
                citations(rangeIndex).BibYear = yearVal
                
                On Error GoTo 0

                inCitation = False
            End If

            If (charRange.Text = ";" Or charRange.Text = ",") And Not onlyYear Then
                i = i + 1
                startChar = endChar + 2
                inCitation = True
            End If

        End If
    Next i

    ' Handle case where no citations were found
    If rangeIndex = -1 Then
        ReDim citations(0 To -1)
    Else
        ReDim Preserve citations(0 To rangeIndex)
    End If
End Sub

' Extracts citations for simple numeric styles, e.g., [1], [2], [3]
Private Sub ExtractNumberInBrackets(field As field, ByRef citations() As Citation, Optional bracket As String = "[]")
    
    '--- Variable Declarations ---
    Dim targetRange As Range, charRange As Range
    Dim rangeIndex As Long
    Dim startBracket As String, endBracket As String
    Dim inBrackets As Boolean
    Dim json As Object
    Dim startChar As Long, endChar As Long
    Dim i As Long
    
    '--- Initializations ---
    Set targetRange = field.result
    Set charRange = targetRange.Duplicate
    charRange.Collapse wdCollapseStart
    Set json = ParseCSLCitationJson(field.code)
    
    ReDim citations(0)
    rangeIndex = -1
    
    startBracket = Left(bracket, 1)
    endBracket = Right(bracket, 1)
    
    inBrackets = False
    
    For i = 1 To targetRange.Characters.Count
        charRange.Start = targetRange.Start + i - 1
        charRange.End = targetRange.Start + i

        If charRange.Text = startBracket Then
            inBrackets = True
            startChar = charRange.Start + 1
        ElseIf charRange.Text = endBracket And inBrackets Then
            If startChar < endChar Then
                rangeIndex = rangeIndex + 1
                If rangeIndex > UBound(citations) Then
                    ReDim Preserve citations(0 To rangeIndex)
                End If
                With citations(rangeIndex)
                    .Start = startChar
                    .End = endChar
                    .BibTitle = RemoveHtmlTags( _
                        json.Item("CSL.citationItems(" & rangeIndex & ").itemData.title"))
                    .BibId = json.Item("CSL.citationItems(" & rangeIndex & ").itemData.id")
                    
                    Dim yearVal As String
                    On Error Resume Next
                    .BibAuthor = json.Item("CSL.citationItems(" & rangeIndex & ").itemData.author(0).family")
                    yearVal = json.Item("CSL.citationItems(" & rangeIndex & ").itemData.issued.dateparts(0)(0)")
                    If yearVal = "" Then
                        yearVal = json.Item("CSL.citationItems(" & rangeIndex & ").itemData.issued.date-parts(0)(0)")
                    End If
                    .BibYear = yearVal
                    
                    On Error GoTo 0
                End With
            End If
            inBrackets = False
        ElseIf inBrackets And IsNumeric(charRange.Text) Then
            endChar = charRange.End
        End If
    Next i

    ' Handle case where no citations were found
    If rangeIndex = -1 Then
        ReDim citations(0 To -1)
    Else
        ReDim Preserve citations(0 To rangeIndex)
    End If
End Sub

' Extracts citations for complex numeric styles, e.g., [47, 98, 100-102]
Private Sub ExtractSerialNumberCitations(field As field, ByRef citations() As Citation, Optional border = "")
    
    '--- Variable Declarations ---
    Dim targetRange As Range, charRange As Range
    Dim rangeIndex As Long, citOrder As Long
    Dim startBorder As String, endBorder As String
    Dim inCitation As Boolean
    Dim lastNum As Long
    Dim json As Object
    Dim startChar As Long, endChar As Long
    Dim currentChar As String, citationText As String
    Dim i As Long, RL As Long
    
    '--- Initializations ---
    Set targetRange = field.result
    Set charRange = targetRange.Duplicate
    charRange.Collapse wdCollapseStart
    Set json = ParseCSLCitationJson(field.code)
    
    ReDim citations(0)
    rangeIndex = -1
    citOrder = -1
    
    startBorder = Left(border, 1)
    endBorder = Right(border, 1)
    
    RL = targetRange.Characters.Count
    inCitation = False
    lastNum = 0

    If Len(endBorder) = 0 Then
        RL = RL + 1
        endBorder = "]"
    End If

    For i = 1 To RL
        charRange.Start = targetRange.Start + i - 1
        charRange.End = targetRange.Start + i

        If i <= targetRange.Characters.Count Then
            currentChar = charRange.Text
        Else
            currentChar = endBorder
        End If

        If currentChar Like "[0-9]" And Not inCitation Then
            inCitation = True
            startChar = charRange.Start
            citationText = currentChar

        ElseIf currentChar = "," Or currentChar = endBorder Or currentChar = ChrW(8211) Then

            If currentChar = ChrW(8211) Then
                lastNum = CLng(citationText)
            End If

            If inCitation Then
                endChar = charRange.Start

                rangeIndex = rangeIndex + 1
                If rangeIndex > UBound(citations) Then
                    ReDim Preserve citations(0 To rangeIndex)
                End If

                If (currentChar = "," Or currentChar = endBorder) And lastNum > 0 Then
                    citOrder = citOrder + CLng(citationText) - lastNum
                Else
                    citOrder = citOrder + 1
                End If

                citations(rangeIndex).Start = startChar
                citations(rangeIndex).End = endChar
                citations(rangeIndex).BibTitle = RemoveHtmlTags( _
                    json.Item("CSL.citationItems(" & citOrder & ").itemData.title"))
                citations(rangeIndex).BibId = json.Item("CSL.citationItems(" & citOrder & ").itemData.id")
                
                Dim yearVal As String
                On Error Resume Next
                citations(rangeIndex).BibAuthor = json.Item("CSL.citationItems(" & citOrder & ").itemData.author(0).family")
                
                yearVal = json.Item("CSL.citationItems(" & citOrder & ").itemData.issued.dateparts(0)(0)")
                If yearVal = "" Then
                    yearVal = json.Item("CSL.citationItems(" & citOrder & ").itemData.issued.date-parts(0)(0)")
                End If
                citations(rangeIndex).BibYear = yearVal
                
                On Error GoTo 0

                If Len(citations(rangeIndex).BibTitle) = 0 Then
                    Err.Raise vbObjectError + 1, "ExtractCitations", "Can not find citation CSL data"
                End If

                inCitation = False
            End If

            If currentChar = "," Or currentChar = endBorder Then
                lastNum = 0
            End If

        ElseIf inCitation Then
            citationText = citationText & currentChar
        End If
    Next i

    ' Handle case where no citations were found
    If rangeIndex = -1 Then
        ReDim citations(0 To -1)
    Else
        ReDim Preserve citations(0 To rangeIndex)
    End If
End Sub

'-------------------------------------------------------------------
' SUPPORTED CITATION STYLES
'-------------------------------------------------------------------

' Returns the supported Zotero styles as (ID, display name, group) arrays.
Private Function ZCL_Styles() As Collection
    Dim styles As New Collection
    
    ' Author-year styles
    styles.Add Array("apa", "APA", "AUTHOR-YEAR")
    styles.Add Array("american-political-science-association", "APSA", "AUTHOR-YEAR")
    styles.Add Array("american-sociological-association", "ASA", "AUTHOR-YEAR")
    styles.Add Array("chicago-author-date", "Chicago (Author-Date)", "AUTHOR-YEAR")
    styles.Add Array("harvard-cite-them-right", "Cite Them Right Harvard", "AUTHOR-YEAR")
    styles.Add Array("elsevier-harvard", "Elsevier Harvard", "AUTHOR-YEAR")
    styles.Add Array("molecular-plant", "Molecular Plant", "AUTHOR-YEAR")
    styles.Add Array("china-national-standard-gb-t-7714-2015-author-date", "China NS (Author-Date)", "AUTHOR-YEAR")
    
    ' Author-only styles
    styles.Add Array("modern-language-association", "MLA", "AUTHOR-ONLY")
    
    ' Numeric styles
    styles.Add Array("ieee", "IEEE", "NUMERIC")
    styles.Add Array("vancouver", "Vancouver", "NUMERIC")
    styles.Add Array("physical-review-letters", "Physical Review Letters", "NUMERIC")
    styles.Add Array("american-physics-society", "APS", "NUMERIC")
    styles.Add Array("american-chemical-society", "ACS", "NUMERIC")
    styles.Add Array("american-medical-association", "AMA", "NUMERIC")
    styles.Add Array("nature", "Nature", "NUMERIC")
    styles.Add Array("bmc-medicine", "BMC Medicine", "NUMERIC")
    styles.Add Array("china-national-standard-gb-t-7714-2015-numeric", "China NS (Numeric)", "NUMERIC")
    styles.Add Array("archives-of-computational-methods-in-engineering", "Archives of Comp. Methods", "NUMERIC")
    styles.Add Array("elsevier-vancouver", "Elsevier NLM/Vancouver", "NUMERIC")
    styles.Add Array("sage-vancouver-brackets", "SAGE NLM/Vancouver (brackets)", "NUMERIC")
    
    Set ZCL_Styles = styles
End Function

' Checks if the given Zotero style ID is in the supported list.
Private Function isSupportedStyle(ByVal style As String) As Boolean
    Dim styles As Collection
    Dim item As Variant
    Set styles = ZCL_Styles()
    For Each item In styles
        If StrComp(item(0), style, vbTextCompare) = 0 Then
            isSupportedStyle = True
            Exit Function
        End If
    Next item
End Function

' Routes the citation field to the correct parsing function based on the style ID.
Private Sub ExtractCitations(field As field, ByRef citations() As Citation, style As String)
    Select Case style
        Case "molecular-plant", "chicago-author-date", "modern-language-association"
            Call ExtractAuthorYearCitations(field, citations, onlyYear:=False, multiRefCommaSep:=False)

        Case "apa", "china-national-standard-gb-t-7714-2015-author-date", _
             "american-political-science-association", "american-sociological-association", _
             "harvard-cite-them-right"
            Call ExtractAuthorYearCitations(field, citations, onlyYear:=True, multiRefCommaSep:=True)

        Case "elsevier-harvard"
            Call ExtractAuthorYearCitations(field, citations, onlyYear:=False, multiRefCommaSep:=True)

        Case "vancouver"
            Call ExtractSerialNumberCitations(field, citations, "()")
        
        Case "physical-review-letters", "american-physics-society", "china-national-standard-gb-t-7714-2015-numeric", "bmc-medicine", "ieee", "archives-of-computational-methods-in-engineering", "elsevier-vancouver", "sage-vancouver-brackets"
            Call ExtractSerialNumberCitations(field, citations, "[]")

        Case "american-chemical-society", "american-medical-association", "nature"
            Call ExtractSerialNumberCitations(field, citations, "")

        Case Else
            Err.Raise vbObjectError + 1, "ExtractCitations", "Citation style not recognized: " & style
    End Select
End Sub

' Displays the About dialog.
Public Sub ZCL_About(control As IRibbonControl)
    Dim msg As String
    
    msg = "ZoteroCiteLinker " & ZCL_VERSION & vbCrLf & vbCrLf & _
          "An MS Word add-in that links Zotero citations" & vbCrLf & _
          "to their matching bibliography entries." & vbCrLf & vbCrLf & _
          "Windows and macOS supported." & vbCrLf & vbCrLf & _
          "Author:  Serdar Berat AYDIN" & vbCrLf & _
          "GitHub:  https://github.com/sBaydin/ZoteroCiteLinker" & vbCrLf & _
          "License: MIT"
    
    MsgBox msg, vbInformation, "About ZoteroCiteLinker"
End Sub

' Kept for templates whose Ribbon still calls the old About callback.
Public Sub ZCL_Information(control As IRibbonControl)
    ZCL_About control
End Sub

' Displays the supported Zotero styles, grouped by citation format.
Public Sub ZCL_SupportedStyles(control As IRibbonControl)
    Dim styles As Collection
    Dim item As Variant
    Dim msg As String
    Dim currentGroup As String
    Dim i As Long
    
    Set styles = ZCL_Styles()
    
    msg = "ZoteroCiteLinker " & ZCL_VERSION & vbCrLf & vbCrLf & _
          styles.Count & " Zotero styles supported"
    currentGroup = ""
    
    For i = 1 To styles.Count
        item = styles(i)
        If item(2) <> currentGroup Then
            currentGroup = item(2)
            msg = msg & vbCrLf & vbCrLf & currentGroup
        End If
        msg = msg & vbCrLf & "  " & item(1) & " - " & item(0)
    Next i
    
    MsgBox msg, vbInformation, "Supported Citation Styles"
End Sub

'-------------------------------------------------------------------
' GITHUB & UPDATES
'-------------------------------------------------------------------

' Opens the GitHub repository in the default browser.
Public Sub ZCL_OpenGitHub(control As IRibbonControl)
    ZCL_OpenURL ZCL_REPO_URL
End Sub

' Compares the installed version with the newest GitHub release.
Public Sub ZCL_CheckUpdates(control As IRibbonControl)
    Dim body As String
    Dim latest As String
    Dim cmp As Long
    Dim msg As String
    
    body = ZCL_FetchText(ZCL_LATEST_API)
    
    If Len(body) = 0 Then
        msg = "Could not reach GitHub to check the version." & vbCrLf & vbCrLf & _
              "The automatic check needs Word for Windows and an internet connection." & vbCrLf & vbCrLf & _
              "Open the releases page in your browser instead?"
        If MsgBox(msg, vbYesNo + vbQuestion, "Check for Updates") = vbYes Then
            ZCL_OpenURL ZCL_RELEASES_URL
        End If
        Exit Sub
    End If
    
    latest = ZCL_LatestTag(body)
    
    If Len(latest) = 0 Then
        msg = "Could not read the latest version number from GitHub." & vbCrLf & vbCrLf & _
              "Open the releases page in your browser instead?"
        If MsgBox(msg, vbYesNo + vbQuestion, "Check for Updates") = vbYes Then
            ZCL_OpenURL ZCL_RELEASES_URL
        End If
        Exit Sub
    End If
    
    cmp = ZCL_CompareVersions(latest, ZCL_VERSION)
    
    If cmp > 0 Then
        msg = "A newer version is available." & vbCrLf & vbCrLf & _
              "Installed: " & ZCL_VERSION & vbCrLf & _
              "Latest: " & latest & vbCrLf & vbCrLf & _
              "Open the download page in your browser?"
        If MsgBox(msg, vbYesNo + vbQuestion, "Update Available") = vbYes Then
            ZCL_OpenURL ZCL_RELEASES_URL
        End If
    Else
        MsgBox "You are running the latest version." & vbCrLf & vbCrLf & _
               "Installed: " & ZCL_VERSION & vbCrLf & _
               "Latest: " & latest, vbInformation, "No Updates"
    End If
End Sub

' Opens a web address in the default browser.
Private Sub ZCL_OpenURL(ByVal url As String)
    On Error Resume Next
    ThisDocument.FollowHyperlink url
    On Error GoTo 0
End Sub

' Downloads a URL as plain text. Returns "" when there is no HTTP support,
' which is the case on macOS, or when the request fails.
Private Function ZCL_FetchText(ByVal url As String) As String
    Dim progIds As Variant
    Dim i As Long
    Dim http As Object
    
    progIds = Array("MSXML2.XMLHTTP.6.0", "MSXML2.XMLHTTP")
    
    For i = 0 To UBound(progIds)
        Set http = Nothing
        On Error Resume Next
        Err.Clear
        Set http = CreateObject(CStr(progIds(i)))
        If Err.Number = 0 Then
            Err.Clear
            http.Open "GET", url, False
            http.setRequestHeader "User-Agent", "ZoteroCiteLinker"
            http.Send
            If Err.Number = 0 Then
                If http.Status = 200 Then
                    ZCL_FetchText = http.responseText
                End If
            End If
        End If
        On Error GoTo 0
        If Len(ZCL_FetchText) > 0 Then Exit Function
    Next i
End Function

' Reads the tag_name value out of a GitHub release JSON response.
Private Function ZCL_LatestTag(ByVal body As String) As String
    Const tagKey As String = """tag_name"":"
    Dim startPos As Long
    Dim endPos As Long
    Dim raw As String
    
    startPos = InStr(1, body, tagKey, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + Len(tagKey)
    
    startPos = InStr(startPos, body, """")
    If startPos = 0 Then Exit Function
    startPos = startPos + 1
    
    endPos = InStr(startPos, body, """")
    If endPos = 0 Then Exit Function
    
    raw = Mid$(body, startPos, endPos - startPos)
    If Left$(raw, 1) = "v" Or Left$(raw, 1) = "V" Then
        raw = Mid$(raw, 2)
    End If
    ZCL_LatestTag = Trim$(raw)
End Function

' Compares dotted version numbers: 1 when a is newer, -1 when b is newer, 0 when equal.
Private Function ZCL_CompareVersions(ByVal a As String, ByVal b As String) As Long
    Dim aParts As Variant
    Dim bParts As Variant
    Dim parts As Long
    Dim i As Long
    Dim aNum As Long
    Dim bNum As Long
    
    aParts = Split(a, ".")
    bParts = Split(b, ".")
    
    parts = UBound(aParts)
    If UBound(bParts) > parts Then parts = UBound(bParts)
    
    For i = 0 To parts
        aNum = 0
        bNum = 0
        If i <= UBound(aParts) Then aNum = Val(aParts(i))
        If i <= UBound(bParts) Then bNum = Val(bParts(i))
        
        If aNum > bNum Then
            ZCL_CompareVersions = 1
            Exit Function
        ElseIf aNum < bNum Then
            ZCL_CompareVersions = -1
            Exit Function
        End If
    Next i
End Function

'-------------------------------------------------------------------
' LINK HEALTH CHECK
'-------------------------------------------------------------------

' Counts the Zotero citation fields so a progress figure can be shown.
Private Function ZCL_CountCitationFields(targetFields) As Long
    Dim aField As field
    
    For Each aField In targetFields
        If aField.Type = wdFieldAddin Then
            If InStr(aField.code, "ADDIN ZOTERO_ITEM") > 0 Then
                ZCL_CountCitationFields = ZCL_CountCitationFields + 1
            End If
        End If
    Next aField
End Function

' Counts leftover temporary bookmarks from an interrupted linking run.
Private Function ZCL_TempBookmarkCount(doc As Document) As Long
    Dim bk As Bookmark
    
    For Each bk In doc.Bookmarks
        If Left$(bk.name, Len(ZCL_TEMP_PREFIX)) = ZCL_TEMP_PREFIX Then
            ZCL_TempBookmarkCount = ZCL_TempBookmarkCount + 1
        End If
    Next bk
End Function

' Deletes leftover temporary bookmarks from an interrupted linking run.
Private Sub ZCL_DeleteTempBookmarks(doc As Document)
    Dim names As Collection
    Dim bk As Bookmark
    Dim target As String
    Dim i As Long
    
    Set names = New Collection
    For Each bk In doc.Bookmarks
        If Left$(bk.name, Len(ZCL_TEMP_PREFIX)) = ZCL_TEMP_PREFIX Then
            names.Add bk.name
        End If
    Next bk
    
    For i = 1 To names.Count
        target = CStr(names(i))
        On Error Resume Next
        If doc.Bookmarks.Exists(target) Then doc.Bookmarks(target).Delete
        On Error GoTo 0
    Next i
End Sub

' Restores application state after a linking run, including after a cancel.
Private Sub ZCL_EndLinkRun()
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableCancelKey = wdCancelDisabled
    
    On Error Resume Next
    If ActiveDocument Is Nothing Then Exit Sub
    ZCL_DeleteTempBookmarks ActiveDocument
    On Error GoTo 0
End Sub

' Reports whether the citation links in the active document are still healthy.
Public Sub ZCL_CheckLinks(control As IRibbonControl)
    Dim doc As Document
    Dim aField As field
    Dim hp As Hyperlink
    Dim citFields As Long
    Dim linked As Long
    Dim unlinked As Long
    Dim broken As Long
    Dim stale As Long
    Dim brokenList As String
    Dim target As String
    Dim msg As String
    
    Set doc = ActiveDocument
    
    ' 1. Citation fields, and how many of them carry a hyperlink
    For Each aField In doc.Fields
        If aField.Type = wdFieldAddin Then
            If InStr(aField.code, "ADDIN ZOTERO_ITEM") > 0 Then
                citFields = citFields + 1
                On Error Resume Next
                If aField.result.Hyperlinks.Count = 0 Then unlinked = unlinked + 1
                On Error GoTo 0
            End If
        End If
    Next aField
    
    ' 2. Citation links, and whether their target bookmark still exists
    For Each hp In doc.Hyperlinks
        target = ZCL_LinkTarget(hp)
        
        If ZCL_IsCiteName(target) Then
            linked = linked + 1
            If Not doc.Bookmarks.Exists(target) Then
                broken = broken + 1
                If broken <= 3 Then
                    brokenList = brokenList & vbCrLf & "  " & target
                End If
            End If
        End If
    Next hp
    
    ' 3. Leftover temporary bookmarks from an interrupted run
    stale = ZCL_TempBookmarkCount(doc)
    If stale > 0 Then ZCL_DeleteTempBookmarks doc
    
    msg = "ZoteroCiteLinker Link Check" & vbCrLf & vbCrLf & _
          "Citation fields:  " & citFields & vbCrLf & _
          "Citation links:   " & linked & vbCrLf & _
          "Not linked:       " & unlinked & vbCrLf & _
          "Broken links:     " & broken
    
    If broken > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Broken targets:" & brokenList
        If broken > 3 Then msg = msg & vbCrLf & "  ..."
    End If
    
    If stale > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Removed " & stale & " leftover temporary bookmark(s)."
    End If
    
    If citFields = 0 Then
        MsgBox msg & vbCrLf & vbCrLf & "No Zotero citation fields were found in this document.", _
               vbInformation, "Link Check"
        Exit Sub
    End If
    
    If unlinked = 0 And broken = 0 Then
        MsgBox msg & vbCrLf & vbCrLf & "All citations are linked.", _
               vbInformation, "Link Check"
        Exit Sub
    End If
    
    If MsgBox(msg & vbCrLf & vbCrLf & "Run Link All now to rebuild the links?", _
              vbYesNo + vbQuestion, "Link Check") = vbYes Then
        ZCL_LinkCitationAll Nothing
    End If
End Sub

'-------------------------------------------------------------------
' CITATION ANALYSIS
'-------------------------------------------------------------------

' Returns the bookmark a citation link points to, without the "#" prefix.
Private Function ZCL_LinkTarget(ByVal hp As Hyperlink) As String
    Dim s As String
    
    On Error Resume Next
    s = hp.SubAddress
    On Error GoTo 0
    
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)
    ZCL_LinkTarget = s
End Function

' True when the name looks like a bookmark created by this add-in.
Private Function ZCL_IsCiteName(ByVal name As String) As Boolean
    ZCL_IsCiteName = (UCase$(Left$(name, 5)) = "CITE_")
End Function

' Collapses whitespace and shortens a text to a readable label.
Private Function ZCL_ShortText(ByVal s As String, ByVal maxLen As Long) As String
    Dim cut As String
    Dim lastSpace As Long
    
    s = Replace(s, vbCr, " ")
    s = Replace(s, Chr$(10), " ")
    s = Replace(s, Chr$(7), " ")
    s = Replace(s, Chr$(9), " ")
    s = Replace(s, Chr$(160), " ")
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    
    If Len(s) > maxLen Then
        cut = Left$(s, maxLen)
        lastSpace = InStrRev(cut, " ")
        If lastSpace > 0 Then cut = Left$(cut, lastSpace - 1)
        s = Trim$(cut) & "..."
    End If
    
    ZCL_ShortText = s
End Function

' Escapes a string for HTML, keeping the output pure ASCII so the file
' opens correctly on any platform.
Private Function ZCL_HtmlEscape(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "&"
                out = out & "&amp;"
            Case "<"
                out = out & "&lt;"
            Case ">"
                out = out & "&gt;"
            Case """"
                out = out & "&quot;"
            Case Else
                code = AscW(ch)
                If code < 0 Then code = code + 65536
                If code < 32 Or code > 126 Then
                    out = out & "&#" & code & ";"
                Else
                    out = out & ch
                End If
        End Select
    Next i
    
    ZCL_HtmlEscape = out
End Function

' Returns the page number of a range, or 0 when Word cannot report it.
Private Function ZCL_PageOf(rng As Range) As Long
    On Error Resume Next
    ZCL_PageOf = rng.Information(wdActiveEndPageNumber)
    On Error GoTo 0
End Function

' Page number as display text, with a placeholder when it is unknown.
Private Function ZCL_PageLabel(rng As Range) As String
    Dim p As Long
    p = ZCL_PageOf(rng)
    If p > 0 Then
        ZCL_PageLabel = CStr(p)
    Else
        ZCL_PageLabel = "?"
    End If
End Function

' The title stored in a citation hyperlink's tooltip, or "" when there is none.
Private Function ZCL_LinkTitle(ByVal hp As Hyperlink) As String
    Dim s As String
    
    On Error Resume Next
    s = hp.ScreenTip
    On Error GoTo 0
    
    ZCL_LinkTitle = Trim$(s)
End Function

' Returns the Cite_ bookmark of the bibliography entry under the cursor.
' The Zotero bibliography is a single field, so entry boundaries cannot be
' trusted from the bookmarks alone. Strategies, most precise first:
'   1. the entry title (stored as a citation tooltip) appears in the current
'      paragraph or line of the bibliography,
'   2. the closest title before the cursor in the bibliography,
'   3. the smallest citation bookmark that contains the selection.
Private Function ZCL_BookmarkAtSelection(doc As Document) As String
    Dim bk As Bookmark
    Dim hp As Hyperlink
    Dim aField As field
    Dim bibRange As Range
    Dim paraStart As Long
    Dim paraEnd As Long
    Dim segStart As Long
    Dim segEnd As Long
    Dim segText As String
    Dim prefixText As String
    Dim suffixText As String
    Dim title As String
    Dim target As String
    Dim pos As Long
    Dim bestPos As Long
    Dim bestName As String
    Dim bestLen As Long
    Dim selStart As Long
    Dim selEnd As Long
    Dim p1 As Long
    Dim p2 As Long
    Dim p3 As Long

    selStart = Selection.Start
    selEnd = Selection.End

    ' The Zotero bibliography field, when the document has one
    Set bibRange = Nothing
    For Each aField In doc.Fields
        If aField.Type = wdFieldAddin Then
            If InStr(aField.code, "ADDIN ZOTERO_BIBL") > 0 Then
                Set bibRange = aField.result
                Exit For
            End If
        End If
    Next aField

    ' 1. Title match inside the paragraph or line at the cursor
    If Not bibRange Is Nothing Then
        If selStart >= bibRange.Start And selStart <= bibRange.End Then
            segStart = bibRange.Start
            segEnd = bibRange.End

            If selStart > bibRange.Start Then
                prefixText = doc.Range(bibRange.Start, selStart).Text
                p1 = InStrRev(prefixText, vbCr)
                p2 = InStrRev(prefixText, Chr$(11))
                p3 = InStrRev(prefixText, Chr$(7))
                If p2 > p1 Then p1 = p2
                If p3 > p1 Then p1 = p3
                If p1 > 0 Then segStart = bibRange.Start + p1
            End If

            If selStart < bibRange.End Then
                suffixText = doc.Range(selStart, bibRange.End).Text
                p1 = InStr(suffixText, vbCr)
                p2 = InStr(suffixText, Chr$(11))
                p3 = InStr(suffixText, Chr$(7))
                If p2 > 0 And (p1 = 0 Or p2 < p1) Then p1 = p2
                If p3 > 0 And (p1 = 0 Or p3 < p1) Then p1 = p3
                If p1 > 0 Then segEnd = selStart + p1 - 1
            End If

            If segEnd - segStart <= 1200 Then
                segText = doc.Range(segStart, segEnd).Text
                For Each hp In doc.Hyperlinks
                    title = ZCL_LinkTitle(hp)
                    If Len(title) >= 12 Then
                        If InStr(1, segText, title, vbTextCompare) > 0 Then
                            target = ZCL_LinkTarget(hp)
                            If ZCL_IsCiteName(target) Then
                                ZCL_BookmarkAtSelection = target
                                Exit Function
                            End If
                        End If
                    End If
                Next hp
            End If

            ' 2. Closest entry title before the cursor
            If selStart > bibRange.Start Then
                prefixText = doc.Range(bibRange.Start, selStart).Text
            Else
                prefixText = ""
            End If
        Else
            prefixText = ""
        End If
    Else
        On Error Resume Next
        paraStart = Selection.Paragraphs(1).Range.Start
        paraEnd = Selection.Paragraphs(1).Range.End
        On Error GoTo 0

        prefixText = ""
        If paraStart > 0 And paraStart < selStart Then
            prefixText = doc.Range(paraStart, selStart).Text
        End If

        If paraEnd > paraStart And paraEnd - paraStart <= 1200 Then
            segText = doc.Range(paraStart, paraEnd).Text
            For Each hp In doc.Hyperlinks
                title = ZCL_LinkTitle(hp)
                If Len(title) >= 12 Then
                    If InStr(1, segText, title, vbTextCompare) > 0 Then
                        target = ZCL_LinkTarget(hp)
                        If ZCL_IsCiteName(target) Then
                            ZCL_BookmarkAtSelection = target
                            Exit Function
                        End If
                    End If
                End If
            Next hp
        End If
    End If

    bestPos = 0
    bestName = ""

    If Len(prefixText) > 0 Then
        For Each hp In doc.Hyperlinks
            title = ZCL_LinkTitle(hp)
            If Len(title) >= 12 Then
                pos = InStrRev(prefixText, title, -1, vbTextCompare)
                If pos > bestPos Then
                    target = ZCL_LinkTarget(hp)
                    If ZCL_IsCiteName(target) Then
                        bestPos = pos
                        bestName = target
                    End If
                End If
            End If
        Next hp
    End If

    If Len(bestName) > 0 Then
        ZCL_BookmarkAtSelection = bestName
        Exit Function
    End If

    ' 3. Smallest bookmark that contains the selection
    bestLen = -1
    For Each bk In doc.Bookmarks
        If ZCL_IsCiteName(bk.name) Then
            If bk.Range.Start <= selStart And bk.Range.End >= selEnd Then
                If bestLen < 0 Or (bk.Range.End - bk.Range.Start) < bestLen Then
                    bestLen = bk.Range.End - bk.Range.Start
                    bestName = bk.name
                End If
            End If
        End If
    Next bk

    ZCL_BookmarkAtSelection = bestName
End Function

' Asks where to save the audit report. Returns "" when the user cancels.
Private Function ZCL_AskSavePath(ByVal suggested As String) As String
    Dim dlg As Object
    Dim path As String
    Dim dialogWorks As Boolean
    
    dialogWorks = False
    
    On Error Resume Next
    Set dlg = Application.FileDialog(msoFileDialogSaveAs)
    If Err.Number = 0 Then
        dialogWorks = True
        dlg.Title = "Save citation audit report"
        dlg.InitialFileName = suggested
        If dlg.Show = -1 Then
            path = dlg.SelectedItems(1)
        Else
            path = ""
        End If
        If Err.Number <> 0 Then
            dialogWorks = False
            path = ""
        End If
    End If
    Err.Clear
    On Error GoTo 0
    
    If Not dialogWorks Then
        path = InputBox("Full path for the HTML report:", "Save citation audit", suggested)
    End If
    
    If Len(path) = 0 Then
        ZCL_AskSavePath = ""
        Exit Function
    End If
    
    ZCL_AskSavePath = ZCL_EnsureHtmlExt(path)
End Function

' Makes sure the chosen file name ends in .html, stripping a Word extension
' that the Save As dialog may have added.
Private Function ZCL_EnsureHtmlExt(ByVal path As String) As String
    Dim dotPos As Long
    Dim ext As String
    
    dotPos = InStrRev(path, ".")
    If dotPos > 0 Then
        If InStr(dotPos, path, "\") = 0 And InStr(dotPos, path, "/") = 0 Then
            ext = LCase$(Mid$(path, dotPos + 1))
        End If
    End If
    
    Select Case ext
        Case "html", "htm"
            ' already a report name
        Case "docx", "doc", "dotx", "dotm", "txt", "rtf"
            path = Left$(path, dotPos - 1) & ".html"
        Case Else
            path = path & ".html"
    End Select
    
    ZCL_EnsureHtmlExt = path
End Function

' Opens a local file with Word's own handler (no external process is
' started, which keeps virus scanners calm).
Private Sub ZCL_OpenFile(ByVal filePath As String)
    On Error Resume Next
    ThisDocument.FollowHyperlink Address:=filePath
    On Error GoTo 0
End Sub

' Suggests <document>-citation-audit.html next to the document.
Private Function ZCL_DefaultAuditPath(doc As Document) As String
    Dim folder As String
    Dim baseName As String
    Dim dotPos As Long
    
    folder = doc.Path
    If Len(folder) = 0 Then
        folder = Environ("USERPROFILE")
        If Len(folder) = 0 Then folder = Environ("HOME")
    End If
    
    baseName = doc.name
    dotPos = InStrRev(baseName, ".")
    If dotPos > 1 Then baseName = Left$(baseName, dotPos - 1)
    If Len(baseName) = 0 Then baseName = "document"
    
    ZCL_DefaultAuditPath = folder & ZCL_PATH_SEP & baseName & "-citation-audit.html"
End Function

' Lists the citations that point to the bibliography entry under the cursor.
Public Sub ZCL_CitedBy(control As IRibbonControl)
    Dim doc As Document
    Dim bk As Bookmark
    Dim hp As Hyperlink
    Dim targetName As String
    Dim starts() As Long
    Dim ends() As Long
    Dim pages() As Long
    Dim texts() As String
    Dim n As Long
    Dim i As Long
    Dim prompt As String
    Dim answer As String
    Dim pick As Long
    Dim rng As Range
    Dim label As String
    Dim pageText As String
    Dim firstTitle As String
    Dim bkText As String
    
    Set doc = ActiveDocument
    targetName = ZCL_BookmarkAtSelection(doc)
    
    If Len(targetName) = 0 Then
        MsgBox "Place the cursor inside a bibliography entry that is already linked." & vbCrLf & vbCrLf & _
               "If the bibliography has no links yet, run Link All first.", _
               vbExclamation, "Cited By"
        Exit Sub
    End If
    
    Set bk = doc.Bookmarks(targetName)
    bkText = bk.Range.Text
    label = ZCL_ShortText(bkText, 110)
    
    ReDim starts(0 To 0)
    ReDim ends(0 To 0)
    ReDim pages(0 To 0)
    ReDim texts(0 To 0)
    n = 0
    
    For Each hp In doc.Hyperlinks
        If StrComp(ZCL_LinkTarget(hp), targetName, vbTextCompare) = 0 Then
            n = n + 1
            ReDim Preserve starts(0 To n - 1)
            ReDim Preserve ends(0 To n - 1)
            ReDim Preserve pages(0 To n - 1)
            ReDim Preserve texts(0 To n - 1)
            
            If Len(firstTitle) = 0 Then firstTitle = ZCL_LinkTitle(hp)
            starts(n - 1) = hp.Range.Start
            ends(n - 1) = hp.Range.End
            pages(n - 1) = ZCL_PageOf(hp.Range)
            texts(n - 1) = ZCL_ShortText(hp.Range.Text, 24)
        End If
    Next hp
    
    If n = 0 Then
        MsgBox label & vbCrLf & vbCrLf & "No citations point to this entry.", _
               vbInformation, "Cited By"
        Exit Sub
    End If
    
    ' A bookmark that covers the whole bibliography gives a useless label
    If Len(bkText) > 1000 And Len(firstTitle) > 0 Then
        label = ZCL_ShortText(firstTitle, 110)
    End If
    
    prompt = label & vbCrLf & vbCrLf & "Cited " & n & " time"
    If n > 1 Then prompt = prompt & "s"
    prompt = prompt & ":" & vbCrLf & vbCrLf
    
    For i = 1 To n
        If i > 12 Then
            prompt = prompt & "  ... and " & (n - 12) & " more"
            Exit For
        End If
        pageText = CStr(pages(i - 1))
        If pages(i - 1) = 0 Then pageText = "?"
        prompt = prompt & " " & i & ". page " & pageText & "   " & texts(i - 1) & vbCrLf
    Next i
    
    MsgBox prompt, vbInformation, "Cited By"
    
    answer = InputBox("Type a number between 1 and " & n & " to jump to that citation." & vbCrLf & vbCrLf & _
                      "Leave it empty, or press Cancel, to close this dialog.", _
                      "Cited By", "1")
    
    answer = Trim$(answer)
    If Len(answer) = 0 Then Exit Sub
    If Not IsNumeric(answer) Then Exit Sub
    pick = CLng(answer)
    If pick < 1 Or pick > n Then Exit Sub
    
    Set rng = doc.Range(Start:=starts(pick - 1), End:=ends(pick - 1))
    
    On Error Resume Next
    rng.Select
    ScrollToRange rng
    On Error GoTo 0
End Sub

' Writes an HTML report of references, their citations and uncited entries.
Public Sub ZCL_CitationAudit(control As IRibbonControl)
    Dim doc As Document
    Dim bk As Bookmark
    Dim hp As Hyperlink
    Dim aField As field
    Dim labels As Object
    Dim counts As Object
    Dim pagesText As Object
    Dim titles As Object
    Dim names As Collection
    Dim bkName As String
    Dim target As String
    Dim label As String
    Dim citFields As Long
    Dim unlinked As Long
    Dim orphans As Long
    Dim uncited As Long
    Dim links As Long
    Dim pageList As String
    Dim html As String
    Dim lines As Variant
    Dim path As String
    Dim folder As String
    Dim pageCount As Long
    Dim f As Integer
    Dim i As Long
    Dim errNum As Long
    Dim errDesc As String
    Dim reportTitle As String

    Set doc = ActiveDocument
    Set labels = CreateDict()
    Set counts = CreateDict()
    Set pagesText = CreateDict()
    Set titles = CreateDict()
    Set names = New Collection

    ' 1. Titles from the citation tooltips, used when a bookmark covers the
    '    whole bibliography instead of a single entry
    For Each hp In doc.Hyperlinks
        target = ZCL_LinkTarget(hp)
        If ZCL_IsCiteName(target) Then
            If Not titles.Exists(target) Then
                titles.Item(target) = ZCL_LinkTitle(hp)
            End If
        End If
    Next hp

    ' 2. Bibliography entries that carry a bookmark from this add-in
    For Each bk In doc.Bookmarks
        If ZCL_IsCiteName(bk.name) Then
            bkName = bk.name
            If Not labels.Exists(bkName) Then
                label = bk.Range.Text
                If Len(label) > 1000 Then
                    label = ""
                    If titles.Exists(bkName) Then label = CStr(titles.Item(bkName))
                    If Len(label) = 0 Then label = bkName
                End If
                labels.Item(bkName) = ZCL_ShortText(label, 150)
                counts.Item(bkName) = 0
                pagesText.Item(bkName) = ""
                names.Add bkName
            End If
        End If
    Next bk

    ' 3. Citation links in the document
    For Each hp In doc.Hyperlinks
        target = ZCL_LinkTarget(hp)
        If ZCL_IsCiteName(target) Then
            links = links + 1
            If counts.Exists(target) Then
                counts.Item(target) = CLng(counts.Item(target)) + 1
                pagesText.Item(target) = pagesText.Item(target) & "|" & ZCL_PageLabel(hp.Range)
            Else
                orphans = orphans + 1
            End If
        End If
    Next hp

    ' 4. Citation fields that carry no link at all
    For Each aField In doc.Fields
        If aField.Type = wdFieldAddin Then
            If InStr(aField.code, "ADDIN ZOTERO_ITEM") > 0 Then
                citFields = citFields + 1
                On Error Resume Next
                If aField.result.Hyperlinks.Count = 0 Then unlinked = unlinked + 1
                On Error GoTo 0
            End If
        End If
    Next aField

    For i = 1 To names.Count
        If CLng(counts.Item(names(i))) = 0 Then uncited = uncited + 1
    Next i

    If names.Count = 0 Then
        MsgBox "No linked references were found in this document." & vbCrLf & vbCrLf & _
               "Run Link All first, then run the audit again.", vbExclamation, "Citation Audit"
        Exit Sub
    End If

    folder = doc.Path
    pageCount = 0
    On Error Resume Next
    pageCount = doc.ComputeStatistics(wdStatisticPages)
    On Error GoTo 0

    ' 5. Build the report
    reportTitle = "ZoteroCiteLinker " & ZCL_VERSION & " Audit Report"

    html = "<!DOCTYPE html>" & vbCrLf
    html = html & "<html><head><meta charset=""utf-8"">" & vbCrLf
    html = html & "<title>" & ZCL_HtmlEscape(reportTitle) & "</title>" & vbCrLf
    html = html & "<style>" & vbCrLf
    html = html & "body{font-family:-apple-system,'Segoe UI',Arial,sans-serif;margin:32px;color:#222}" & vbCrLf
    html = html & "h1{font-size:26px;text-transform:uppercase;letter-spacing:.6px;margin:0 0 8px}" & vbCrLf
    html = html & ".meta{color:#555;font-size:13px;margin:0 0 24px;line-height:1.7}" & vbCrLf
    html = html & ".meta b{color:#222}" & vbCrLf
    html = html & "h2{font-size:16px;margin-top:28px}" & vbCrLf
    html = html & "table{border-collapse:collapse;font-size:13px;width:100%}" & vbCrLf
    html = html & "th,td{border:1px solid #ddd;padding:6px 8px;text-align:left;vertical-align:top}" & vbCrLf
    html = html & "th{background:#f4f4f4}tr:nth-child(even) td{background:#fafafa}" & vbCrLf
    html = html & "th.sortable{cursor:pointer;user-select:none}th.sortable:hover{background:#e9e9e9}" & vbCrLf
    html = html & "th.sortable[data-dir=asc]:after{content:' [A-Z]'}th.sortable[data-dir=desc]:after{content:' [Z-A]'}" & vbCrLf
    html = html & "th.num[data-dir=asc]:after{content:' [0-9]'}th.num[data-dir=desc]:after{content:' [9-0]'}" & vbCrLf
    html = html & ".num{text-align:right;white-space:nowrap}.zero{color:#b00;font-weight:600}" & vbCrLf
    html = html & ".toolbar{display:flex;gap:14px;align-items:center;margin:12px 0 10px;font-size:13px}" & vbCrLf
    html = html & ".toolbar input[type=search]{width:280px;padding:6px 8px;border:1px solid #ccc;border-radius:4px}" & vbCrLf
    html = html & ".hint{color:#777;font-size:12px;margin:8px 0 0}" & vbCrLf
    html = html & "</style></head><body>" & vbCrLf
    html = html & "<h1>" & ZCL_HtmlEscape(reportTitle) & "</h1>" & vbCrLf
    html = html & "<div class=""meta"">" & vbCrLf
    html = html & "<div><b>Document:</b> " & ZCL_HtmlEscape(doc.name) & "</div>" & vbCrLf
    If Len(folder) > 0 Then
        html = html & "<div><b>Folder:</b> " & ZCL_HtmlEscape(folder) & "</div>" & vbCrLf
    End If
    html = html & "<div><b>Generated:</b> " & Format$(Now, "yyyy-mm-dd") & " " & Format$(Now, "hh:nn") & "</div>" & vbCrLf
    If pageCount > 0 Then
        html = html & "<div><b>Pages:</b> " & pageCount & "</div>" & vbCrLf
    End If
    html = html & "</div>" & vbCrLf

    html = html & "<h2>Summary</h2>" & vbCrLf & "<table>" & vbCrLf
    html = html & "<tr><th>References with a bookmark</th><td class=""num"">" & names.Count & "</td></tr>" & vbCrLf
    html = html & "<tr><th>Citation links</th><td class=""num"">" & links & "</td></tr>" & vbCrLf
    html = html & "<tr><th>Uncited references</th><td class=""num"">" & uncited & "</td></tr>" & vbCrLf
    html = html & "<tr><th>Citation fields</th><td class=""num"">" & citFields & "</td></tr>" & vbCrLf
    html = html & "<tr><th>Fields without a link</th><td class=""num"">" & unlinked & "</td></tr>" & vbCrLf
    html = html & "<tr><th>Links to a missing entry</th><td class=""num"">" & orphans & "</td></tr>" & vbCrLf
    html = html & "</table>" & vbCrLf

    html = html & "<h2>Uncited references (" & uncited & ")</h2>" & vbCrLf
    If uncited = 0 Then
        html = html & "<p>Every reference is cited at least once.</p>" & vbCrLf
    Else
        html = html & "<table><tr><th>Reference</th></tr>" & vbCrLf
        For i = 1 To names.Count
            bkName = names(i)
            If CLng(counts.Item(bkName)) = 0 Then
                html = html & "<tr><td>" & ZCL_HtmlEscape(labels.Item(bkName)) & "</td></tr>" & vbCrLf
            End If
        Next i
        html = html & "</table>" & vbCrLf
    End If

    html = html & "<h2>References and their citations (" & names.Count & ")</h2>" & vbCrLf
    html = html & "<div class=""toolbar"">" & vbCrLf
    html = html & "<input id=""q"" type=""search"" placeholder=""Filter references..."">" & vbCrLf
    html = html & "<label><input id=""onlyUncited"" type=""checkbox""> Only uncited</label>" & vbCrLf
    html = html & "</div>" & vbCrLf
    html = html & "<table id=""refs"">" & vbCrLf
    html = html & "<thead><tr>" & _
                  "<th class=""sortable"" data-sort=""text"">Reference</th>" & _
                  "<th class=""sortable num"" data-sort=""num"">Citations</th>" & _
                  "<th class=""sortable"" data-sort=""firstnum"">Pages</th>" & _
                  "</tr></thead>" & vbCrLf
    html = html & "<tbody>" & vbCrLf
    For i = 1 To names.Count
        bkName = names(i)
        pageList = ZCL_UniquePages(CStr(pagesText.Item(bkName)))

        html = html & "<tr data-uncited=""" & IIf(CLng(counts.Item(bkName)) = 0, "1", "0") & """>"
        html = html & "<td>" & ZCL_HtmlEscape(labels.Item(bkName)) & "</td>"
        html = html & "<td class=""num"
        If CLng(counts.Item(bkName)) = 0 Then html = html & " zero"
        html = html & """>" & counts.Item(bkName) & "</td>"
        html = html & "<td>" & pageList & "</td></tr>" & vbCrLf
    Next i
    html = html & "</tbody></table>" & vbCrLf
    html = html & "<p class=""hint"">Click a column header to sort the table, or type in the box to filter it.</p>" & vbCrLf

    html = html & "<script>" & vbCrLf
    html = html & "(function(){" & vbCrLf
    html = html & "var t=document.getElementById('refs'),q=document.getElementById('q'),u=document.getElementById('onlyUncited');" & vbCrLf
    html = html & "if(!t||!t.tHead||!t.tBodies[0])return;" & vbCrLf
    html = html & "var body=t.tBodies[0],heads=t.tHead.rows[0].cells,dir={};" & vbCrLf
    html = html & "function rows(){return Array.prototype.slice.call(body.rows);}" & vbCrLf
    html = html & "function apply(){" & vbCrLf
    html = html & "var s=q&&q.value?q.value.toLowerCase():'',only=u&&u.checked;" & vbCrLf
    html = html & "rows().forEach(function(r){" & vbCrLf
    html = html & "var hit=r.cells[0].textContent.toLowerCase().indexOf(s)>=0;" & vbCrLf
    html = html & "var unc=r.getAttribute('data-uncited')==='1';" & vbCrLf
    html = html & "r.style.display=(hit&&(!only||unc))?'':'none';});}" & vbCrLf
    html = html & "function key(t,m){if(m==='num')return parseFloat(t)||0;" & vbCrLf
    html = html & "if(m==='firstnum')return parseInt(t,10)||0;return t;}" & vbCrLf
    html = html & "function sortBy(i,numeric){" & vbCrLf
    html = html & "var d=dir[i]==='asc'?'desc':'asc';dir[i]=d;" & vbCrLf
    html = html & "for(var k=0;k<heads.length;k++){heads[k].removeAttribute('data-dir');}" & vbCrLf
    html = html & "heads[i].setAttribute('data-dir',d);" & vbCrLf
    html = html & "rows().sort(function(a,b){" & vbCrLf
    html = html & "var x=a.cells[i].textContent.trim(),y=b.cells[i].textContent.trim();" & vbCrLf
    html = html & "var v=key(x,numeric).toString().localeCompare(key(y,numeric).toString(),'en',{numeric:true});" & vbCrLf
    html = html & "return d==='asc'?v:-v;}).forEach(function(r){body.appendChild(r);});" & vbCrLf
    html = html & "apply();}" & vbCrLf
    html = html & "for(var j=0;j<heads.length;j++){(function(j){var h=heads[j];" & vbCrLf
    html = html & "if(!h.getAttribute('data-sort'))return;" & vbCrLf
    html = html & "h.addEventListener('click',function(){sortBy(j,h.getAttribute('data-sort'));});})(j);}" & vbCrLf
    html = html & "if(q)q.addEventListener('input',apply);" & vbCrLf
    html = html & "if(u)u.addEventListener('change',apply);" & vbCrLf
    html = html & "})();" & vbCrLf
    html = html & "</script>" & vbCrLf

    html = html & "</body></html>" & vbCrLf

    ' 6. Ask where to save, then write the file
    path = ZCL_AskSavePath(ZCL_DefaultAuditPath(doc))
    If Len(path) = 0 Then Exit Sub

    On Error GoTo WriteFailed
    f = FreeFile
    Open path For Output As #f
    lines = Split(html, vbCrLf)
    For i = 0 To UBound(lines)
        Print #f, lines(i)
    Next i
    Close #f
    On Error GoTo 0

    If MsgBox("Report saved to:" & vbCrLf & path & vbCrLf & vbCrLf & "Open it now?", _
              vbYesNo + vbQuestion, "Citation Audit") = vbYes Then
        ZCL_OpenFile path
    End If
    Exit Sub

WriteFailed:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Close #f
    On Error GoTo 0
    MsgBox "The report could not be written:" & vbCrLf & path & vbCrLf & vbCrLf & _
           "Error " & errNum & ": " & errDesc, vbCritical, "Citation Audit"
End Sub

' Turns "|3|12|12|31" into "3, 12, 31".
Private Function ZCL_UniquePages(ByVal raw As String) As String
    Dim parts() As String
    Dim seen As Object
    Dim i As Long
    Dim p As String
    Dim out As String

    If Len(raw) = 0 Then
        ZCL_UniquePages = "-"
        Exit Function
    End If

    Set seen = CreateDict()
    parts = Split(Mid$(raw, 2), "|")

    For i = 0 To UBound(parts)
        p = Trim$(parts(i))
        If Len(p) > 0 Then
            If Not seen.Exists(p) Then
                seen.Item(p) = 1
                If Len(out) > 0 Then out = out & ", "
                out = out & p
            End If
        End If
    Next i

    If Len(out) = 0 Then out = "-"
    ZCL_UniquePages = out
End Function

'-------------------------------------------------------------------
' ZOTERO LINK CITATION
'-------------------------------------------------------------------

' The main engine that links all Zotero citation fields to their bibliography entries.
Private Sub ZoteroLinkCitation(targetFields, Optional debugging As Boolean = False, Optional notify As Boolean = False)
    '--- Variable Declarations ---
    ' Function-level settings and objects
    Dim prefs As Object
    Dim styleId As String
    Dim userTextStyle As String
    Dim bibField As field
    Dim i As Long, iCount As Integer
    Dim resp As String
    
    ' Main field loop
    Dim aField As field
    Dim json As Object
    Dim cits() As Citation, cit As Citation
    Dim hasOneItem As Boolean, hasManyItems As Boolean
    Dim response As VbMsgBoxResult
    
    ' Inner citation loop
    Dim rng As Range
    Dim hp As Hyperlink
    Dim tempBookmarkName As String
    Dim rngBibliography As Range
    Dim rngFound As Range
    
    ' Data parsing/handling
    Dim yearVal As String
    Dim titleToSearch As String, authorToSearch As String, yearToSearch As String, zoteroId As String
    Dim authorForName As String, yearForName As String, titleSnippet As String
    Dim titleAnchor As String
    Dim searchTitle As String
    
    ' Progress and cancellation
    Dim totalFields As Long
    Dim processed As Long
    Dim errNum As Long
    Dim errDesc As String

    '--- Initializations ---
    Set prefs = GetZoteroPrefs()
    Set bibField = Nothing
    iCount = 0
    
    ' Do not support Bookmark-type citations
    If Not prefs.Item("pref-fieldType") = "Field" Then
        MsgBox "Only support 'Fields' type citations" & vbCrLf & vbCrLf & _
            "Click on Document Preferences in Word, then expand the Advanced Options section. " & _
            "This will allow you to verify whether the 'Store citation as bookmarks' option was accidentally enabled.", _
            vbCritical, "Error"
        Exit Sub
    End If

    styleId = prefs.Item("style-id")
    If Not isSupportedStyle(styleId) Then
        MsgBox "The current citation style is not yet supported: " & styleId, vbCritical, "Error"
        Exit Sub
    End If

    If notify Then
        resp = InputBox(title:="Set an MS Word style for hyperlinks?", _
                        prompt:="If you want to set a certain style for hyperlinks," & _
                                " enter the name of that style below." & vbCrLf & _
                                "(Note: Must be a 'Character Style', not a 'Paragraph Style'.)")
        
        If resp <> "" And StyleExists(resp, ActiveDocument) Then
            ' CRITICAL CHECK: Verify the style TYPE.
            ' We ONLY want to apply Character styles (like "Emphasis" or "Strong").
            
            Dim styleType As WdStyleType
            styleType = ActiveDocument.styles(resp).Type
            
            If styleType = wdStyleTypeCharacter Then
                ' This is a Character Style. It's safe to use.
                userTextStyle = resp
            Else
                ' This is a Paragraph or Linked style.
                ' Applying it would format the entire paragraph.
                MsgBox "'" & resp & "' is a Paragraph or Linked style." & vbCrLf & vbCrLf & _
                       "Only 'Character' styles (e.g., 'Emphasis') can be applied to citations." & vbCrLf & _
                       "No style will be applied.", vbExclamation
            End If
        End If
    End If

    ' Find the Zotero bibliography field
    For i = ActiveDocument.Fields.Count To 1 Step -1
        If ActiveDocument.Fields(i).Type = wdFieldAddin Then
            If InStr(ActiveDocument.Fields(i).code, "ADDIN ZOTERO_BIBL") > 0 Then
                Set bibField = ActiveDocument.Fields(i)
                Exit For
            End If
        End If
    Next i

    If bibField Is Nothing Then
        Err.Raise vbObjectError + 513, , "Can not find Zotero bibliography field."
    End If

    '--- Progress and cancellation ---
    totalFields = ZCL_CountCitationFields(targetFields)
    processed = 0
    Application.EnableCancelKey = wdCancelInterrupt
    On Error GoTo ZCL_LinkAborted

    ' Iterate through all fields in the document
    For Each aField In targetFields
        ' Check if the field is a Zotero citation
        If aField.Type = wdFieldAddin Then
            If InStr(aField.code, "ADDIN ZOTERO_ITEM") > 0 Then

                processed = processed + 1
                Application.StatusBar = "ZoteroCiteLinker: citation " & processed & " of " & totalFields & "..."
                If processed Mod 10 = 0 Then DoEvents

                If debugging Then
                    ' Focus to next field
                    Application.ScreenUpdating = True
                    ScrollToRange aField.result
                    aField.result.Select
            
                    ' Update the document
                    DoEvents
                    
                    ' Debug logic
                    response = MsgBox("Processed " & iCount & " citations, found next:" & vbCrLf & _
                        aField.result.Text & vbCrLf & vbCrLf & _
                        "Yes = Process this group" & vbCrLf & _
                        "No = Skip this group" & vbCrLf & _
                        "Cancel = Exit macro", _
                        vbYesNoCancel + vbQuestion, "Debug: Continue?")
            
                    Select Case response
                        Case vbYes
                        Case vbNo
                            GoTo SkipToNextField
                        Case vbCancel
                            GoTo ExitTheMacro
                    End Select
            
                    Application.ScreenUpdating = False
                End If

                Set json = ParseCSLCitationJson(aField.code)
                
                On Error Resume Next
                hasOneItem = (json.Item("CSL.citationItems(0).id") <> "")
                hasManyItems = (json.Item("CSL.citationItems(1).id") <> "")
                On Error GoTo ZCL_LinkAborted
                
                If hasOneItem And Not hasManyItems Then
                    ' LOGIC: SINGLE CITATION (including custom text)
                    ReDim cits(0 To 0)
                    
                    cits(0).Start = aField.result.Start
                    cits(0).End = aField.result.End
                    cits(0).BibTitle = RemoveHtmlTags(json.Item("CSL.citationItems(0).itemData.title"))
                    cits(0).BibId = json.Item("CSL.citationItems(0).itemData.id")
                    
                    On Error Resume Next
                    cits(0).BibAuthor = json.Item("CSL.citationItems(0).itemData.author(0).family")
                    yearVal = json.Item("CSL.citationItems(0).itemData.issued.dateparts(0)(0)")
                    If yearVal = "" Then
                         yearVal = json.Item("CSL.citationItems(0).itemData.issued.date-parts(0)(0)")
                    End If
                    cits(0).BibYear = yearVal
                    On Error GoTo ZCL_LinkAborted
                
                Else
                    ' LOGIC: GROUPED CITATION OR EMPTY FIELD (This is a grouped citation (e.g., [1, 5-7]) or an empty/broken field.)
                    Call ExtractCitations(aField, cits, styleId)
                
                    If UBound(cits) = -1 Then GoTo SkipToNextField
                End If

                For i = 0 To UBound(cits)
                    cit = cits(i)
                    Set rng = aField.result.Document.Range(Start:=cit.Start, End:=cit.End)
                    tempBookmarkName = ZCL_TEMP_PREFIX & i
                    ActiveDocument.Bookmarks.Add name:=tempBookmarkName, Range:=rng
                Next i

                For i = 0 To UBound(cits)
                    cit = cits(i)
                    
                    titleToSearch = cit.BibTitle
                    authorToSearch = cit.BibAuthor
                    yearToSearch = cit.BibYear
                    zoteroId = cit.BibId
                    
                    ' Cleaned data for BOOKMARK NAME
                    ' 1. Sanitize data (robust cleaning for bookmarks)
                    authorForName = CleanForBookmarkName(authorToSearch)
                    yearForName = CleanForBookmarkName(yearToSearch)
                    titleSnippet = CleanForBookmarkName(titleToSearch)
                    
                    ' 2. Assign defaults if data is missing
                    If Len(authorForName) = 0 Then authorForName = "NoAuthor"
                    If Len(yearForName) = 0 Then yearForName = "NoYear"
                    If Len(titleSnippet) = 0 Then titleSnippet = "NoTitle"
                    
                    ' 3. Truncate to fit Word's 40-character bookmark limit (Format: Cite_id(5)_Title(10)_Year(4)_Author(10) = 39 char)
                    If Len(authorForName) > 10 Then authorForName = Left(authorForName, 10)
                    If Len(titleSnippet) > 10 Then titleSnippet = Left(titleSnippet, 10)

                    ' 4. Desired NEW format: Cite_id[ZoteroID]_[TitleFirst10]_[Year]_[AuthorLastname]
                    titleAnchor = "Cite_id" & zoteroId & "_" & titleSnippet & yearForName & "_" & authorForName
                    'titleAnchor = "Cite_id" & zoteroId & "_" & titleSnippet & "_" & yearForName & "_" & authorForName
                    ' Get the range of Zotero bibliography
                    Set rngBibliography = bibField.result
                    
                    ' SAFE SEARCH LOGIC (Title first, then Author)
                    ' Get the first 100 chars of the title for searching
                    If Len(titleToSearch) > 100 Then
                        searchTitle = Left(titleToSearch, 100)
                    Else
                        searchTitle = titleToSearch
                    End If
                    
                    With rngBibliography.Find
                        .ClearFormatting
                        .Text = searchTitle '<-- 1. ARAMA: Sadece Baslik ile
                        .Forward = True
                        .Wrap = wdFindStop
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False ' <-- Joker karakter YOK
                        .Execute
                    End With
                    
                    ' If not found by title and author exists, try again with author
                    If Not rngBibliography.Find.found And Len(authorToSearch) > 0 Then
                        ' Reset the search range (important)
                        Set rngBibliography = bibField.result
                        
                        With rngBibliography.Find
                            .ClearFormatting
                            .Text = authorToSearch ' <-- 2nd SEARCH: With Author only
                            .Forward = True
                            .Wrap = wdFindStop
                            .Format = False
                            .MatchCase = False
                            .MatchWholeWord = False
                            .MatchWildcards = False
                            .Execute
                        End With
                    End If

                    ' Check if the text was found
                    If rngBibliography.Find.found Then
                        Set rngFound = rngBibliography.Duplicate
                        rngFound.Expand Unit:=wdParagraph

                        ' Zotero's bibliography can collapse into one large
                        ' paragraph. In that case only bookmark the line that
                        ' holds the entry, so one entry never covers all of them.
                        If Len(rngFound.Text) > 1200 Then
                            Set rngFound = rngBibliography.Duplicate
                            rngFound.Expand Unit:=wdLine
                        End If

                        If Right(rngFound.Text, 1) = vbCr Then
                            rngFound.End = rngFound.End - 1
                        End If

                        ActiveDocument.Bookmarks.Add Range:=rngFound, name:=titleAnchor
                    Else
                        If MsgBox("Not found in bibliography:" & vbCrLf & titleToSearch & vbCrLf & vbCrLf & _
                                "Do you want to continue with the next Zotero citation?", _
                                vbYesNo + vbCritical, "Error") = vbNo Then
                            GoTo ExitTheMacro
                        Else
                            GoTo SkipToNextCitation
                        End If
                    End If

                    ' Create hyperlink according to temporary bookmark
                    Set hp = ActiveDocument.Hyperlinks.Add( _
                        Anchor:=ActiveDocument.Bookmarks(ZCL_TEMP_PREFIX & i).Range, _
                        SubAddress:=titleAnchor, ScreenTip:=titleToSearch) 'ScreenTip set displayed information of hyperlink
                    hp.Range.Font.Underline = wdUnderlineNone 'Remove hyperlink underline
                    
                    ' Apply text style to the hyperlink
                    If userTextStyle <> "" Then
                        hp.Range.style = ActiveDocument.styles(userTextStyle)
                    End If

                    iCount = iCount + 1

SkipToNextCitation:
                    ActiveDocument.Bookmarks(ZCL_TEMP_PREFIX & i).Delete

                Next i

            End If
        End If
SkipToNextField: ' Label for skipping a field (e.g., in debug mode or for empty fields)
    Next aField
    
    Call SetZoteroCitationColor

ExitTheMacro:

    ZCL_EndLinkRun

    If notify Then MsgBox "Linked " & iCount & " Zotero citations.", vbInformation, "Finish"

    Exit Sub

ZCL_LinkAborted:
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear

    ZCL_EndLinkRun

    If errNum = 18 Then
        MsgBox "Linking cancelled." & vbCrLf & vbCrLf & _
               "Linked " & iCount & " citations before stopping.", _
               vbInformation, "Cancelled"
        Exit Sub
    End If

    Err.Raise errNum, "ZoteroCiteLinker", errDesc

End Sub

' Public macro: Unlinks all, then relinks all citations in "debug" mode (one by one).
Public Sub ZCL_LinkCitationSelect(control As IRibbonControl)
    
    If Not IsZoteroRunning() Then
        MsgBox "Zotero application does not appear to be running." & vbCrLf & vbCrLf & _
               "Please open Zotero and then run this command again.", _
               vbExclamation, "Zotero Not Found"
        Exit Sub
    End If
    
    '--- Variable Declarations ---
    Dim originalRng As Range
    Dim debugging As Boolean
    
    '--- Initializations ---
    Set originalRng = Selection.Range
    debugging = True ' True = Manual/Debug mode (step-by-step confirmation)

    ' 1. Delete all existing Zotero citation hyperlinks
    Call FuncZoteroUnlinkCitations
    
    ' 2. Add Zotero citation hyperlinks (in debug mode)
    Call ZoteroLinkCitation(ActiveDocument.Fields, debugging, True)

    ' 3. Restore the original selection
    ScrollToRange originalRng
    originalRng.Select

    ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub

' Public macro: Unlinks all, then relinks all citations in the document automatically.
Public Sub ZCL_LinkCitationAll(control As IRibbonControl)
    
    If Not IsZoteroRunning() Then
        MsgBox "Zotero application does not appear to be running." & vbCrLf & vbCrLf & _
               "Please open Zotero and then run this command again.", _
               vbExclamation, "Zotero Not Found"
        Exit Sub
    End If

    '--- Variable Declarations ---
    Dim originalRng As Range
    Dim debugging As Boolean
    
    '--- Initializations ---
    Set originalRng = Selection.Range
    debugging = False

    ' 1. Delete all existing Zotero citation hyperlinks
    Call FuncZoteroUnlinkCitations
    
    ' 2. Add Zotero citation hyperlinks (in automatic mode)
    Call ZoteroLinkCitation(ActiveDocument.Fields, debugging, True)

    ' 3. Restore the original selection
    ScrollToRange originalRng
    originalRng.Select

    ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub

'-------------------------------------------------------------------
' ZOTERO UNLINK FUNCTIONS
'-------------------------------------------------------------------

' Core engine: Safely removes all "CITE_" hyperlinks and bookmarks created by this macro.
Private Sub FuncZoteroUnlinkCitations()
    '--- Variable Declarations ---
    Dim doc As Document
    Dim hp As Hyperlink
    Dim i As Long
    Dim subAddr As String
    Dim bk As Bookmark
    
    '--- Initializations ---
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    On Error GoTo 0 ' Clear any previous errors
    
    ' 1. Delete our hyperlinks (prefixed "CITE_")
    For i = doc.Hyperlinks.Count To 1 Step -1
        subAddr = ""
        
        ' Enable error handling (for broken/external links)
        On Error Resume Next
        Set hp = doc.Hyperlinks(i)
        
        If Err.Number = 0 Then
            subAddr = hp.SubAddress
        End If
        
        Err.Clear
        On Error GoTo 0
        
        ' Case-insensitive check, allowing for the "#" prefix
        subAddr = UCase(VBA.CStr(subAddr))
        
        If Left(subAddr, 5) = "CITE_" Or Left(subAddr, 6) = "#CITE_" Then
            On Error Resume Next
            hp.Range.Fields(1).Unlink
            On Error GoTo 0
        End If
    Next i
    
    ' 2. Delete our bookmarks (prefixed "CITE_")
    For Each bk In doc.Bookmarks
        On Error Resume Next
        If Left(UCase(bk.name), 5) = "CITE_" Then
            bk.Delete
        End If
        On Error GoTo 0
    Next bk
    
    Application.ScreenUpdating = True
End Sub

' Core engine: Removes all WEB hyperlinks (http, https, mailto) from the document.
Private Sub FuncZoteroUnlinkWebLinks()
    '--- Variable Declarations ---
    Dim doc As Document
    Dim hp As Hyperlink
    Dim i As Long
    Dim subAddr As String
    Dim hpAddress As String
    
    '--- Initializations ---
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    For i = doc.Hyperlinks.Count To 1 Step -1
        subAddr = ""
        hpAddress = ""
        
        On Error Resume Next
        Set hp = doc.Hyperlinks(i)
        
        If Err.Number = 0 Then
            subAddr = UCase(hp.SubAddress)
            hpAddress = UCase(hp.Address)
        End If
        Err.Clear
        On Error GoTo 0

        ' 1. SAFETY CHECK: If it's a Zotero Citation link, SKIP IT.
        If Left(subAddr, 5) = "CITE_" Or Left(subAddr, 6) = "#CITE_" Then
        Else
            If Left(hpAddress, 4) = "HTTP" Or Left(hpAddress, 7) = "MAILTO:" Then
                On Error Resume Next
                hp.Range.Fields(1).Unlink
                On Error GoTo 0
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub

' Public macro to unlink all Zotero citations and reset their color.
Public Sub ZCL_UnlinkCitations(control As IRibbonControl)
    ' 1. Call the core engine to remove links and bookmarks
    Call FuncZoteroUnlinkCitations
    ' 2. Call the color engine to reset text color to Auto (Black)
    Call SetZoteroCitationColor(wdAuto)
    MsgBox "Removed Zotero citation links successfully!", vbInformation
End Sub

' Public macro to remove all web (http) and email (mailto) links from the document.
' This will NOT touch Zotero citation links.
Public Sub ZCL_UnlinkURLsMail(control As IRibbonControl)
    Call FuncZoteroUnlinkWebLinks
    MsgBox "All web and email links have been removed (Zotero citations were not affected).", vbInformation
End Sub

'-------------------------------------------------------------------
' CITATION COLOR UTILITIES
'-------------------------------------------------------------------

' Engine: Applies a specified color to all Zotero citation fields.
Private Sub SetZoteroCitationColor(Optional color As WdColorIndex = wdAuto)
    '--- Variable Declarations ---
    Dim fld As field

    Application.ScreenUpdating = False
    
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldAddin Then
            If InStr(1, fld.code.Text, "ADDIN ZOTERO_ITEM", vbTextCompare) > 0 Then
                fld.result.Font.ColorIndex = color
            End If
        End If
    Next fld
    Application.ScreenUpdating = True
End Sub

' Public macro to manually change the color of all Zotero citations using an InputBox menu.
Public Sub ZCL_ChangeColor(control As IRibbonControl)
    '--- Variable Declarations ---
    Dim userInput As String
    Dim chosenColor As WdColorIndex ' Defaults to 0 (wdAuto)
    Dim prompt As String
    
    '--- Initializations ---
    ' 1. Build a clear prompt message showing all options
    prompt = "Please enter a color NAME or its NUMBER:" & vbCrLf & vbCrLf & _
             "1 = Blue" & vbCrLf & _
             "2 = Red" & vbCrLf & _
             "3 = Green" & vbCrLf & _
             "4 = Pink" & vbCrLf & _
             "5 = Teal" & vbCrLf & _
             "6 = Auto (Black) (Default)"

    ' 2. Show the InputBox with "6" (Auto/Black) as the default choice
    userInput = InputBox(prompt, "Change Citation Color", "6")
    
    ' 3. Exit if the user pressed Cancel
    If userInput = "" Then Exit Sub
    
    ' 4. Convert user input (number OR name) to a WdColorIndex
    Select Case LCase(Trim(userInput))
        Case "1", "blue"
            chosenColor = wdBlue
        Case "2", "red"
            chosenColor = wdRed
        Case "3", "green"
            chosenColor = wdGreen
        Case "4", "pink"
            chosenColor = wdPink
        Case "5", "teal"
            chosenColor = wdTeal
        Case "6", "auto", "black", "default"
            chosenColor = wdAuto
        Case Else
            chosenColor = wdAuto
    End Select
    
    ' 5. Call the engine to apply the chosen color
    Call SetZoteroCitationColor(chosenColor)
    MsgBox "Zotero citation colors have been updated.", vbInformation
End Sub

'-------------------------------------------------------------------
' CITATION STYLE UTILITIES
'-------------------------------------------------------------------

' Engine: Applies a specified STYLE to all Zotero citation fields.
Private Sub SetZoteroCitationStyle(ByVal styleName As String)
    '--- Variable Declarations ---
    Dim fld As field
    Dim styleToApply As String
    
    Application.ScreenUpdating = False
    
    ' 1. Determine which style to apply
    If styleName = "" Then
        ' If the styleName is empty, we reset to the default character style.
        ' "Default Paragraph Font" is a built-in Word style.
        styleToApply = ActiveDocument.styles(wdStyleDefaultParagraphFont).NameLocal
    Else
        ' If a style name is given, check if it exists first
        If Not StyleExists(styleName, ActiveDocument) Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        styleToApply = styleName
    End If
    
    ' 2. Loop through fields and apply the style
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldAddin Then
            If InStr(1, fld.code.Text, "ADDIN ZOTERO_ITEM", vbTextCompare) > 0 Then
                fld.result.style = styleToApply
            End If
        End If
    Next fld
    
    Application.ScreenUpdating = True
End Sub

' User set and save a style name to a global variable.
Public Sub ZCL_SetCitationStyle(control As IRibbonControl)
    '--- Variable Declarations ---
    Dim resp As String
    Dim styleType As WdStyleType
    Dim prompt As String
    
    ' 1. Promts
    prompt = "!! WARNING: This action modifies citation styles !!" & vbCrLf & vbCrLf
    prompt = prompt & "Please use a 'Character' style ONLY (e.g., 'Emphasis')." & vbCrLf & vbCrLf
    prompt = prompt & "Using a 'Paragraph' style will reformat the entire paragraph for every citation." & vbCrLf & vbCrLf
    prompt = prompt & "This action CANNOT be easily undone with a single Ctrl+Z." & vbCrLf & vbCrLf
    prompt = prompt & "It is STRONGLY recommended to SAVE your document first." & vbCrLf & vbCrLf
    prompt = prompt & "Current style: " & IIf(gUserTextStyle = "", "Normal (Default)", gUserTextStyle) & vbCrLf
    prompt = prompt & "Enter style name (or 'normal' to reset):"
         
    resp = InputBox(prompt, "Set Citation Style")
    
    ' 2. Exit if the user pressed Cancel
    If StrPtr(resp) = 0 Then Exit Sub
    
    ' 3. User wants to clear the style
    If LCase(resp) = "Normal" Or LCase(resp) = "default" Or resp = "" Then
        gUserTextStyle = ""
        MsgBox "Citation style has been cleared. Default Word formatting will be used.", vbInformation, "Style Cleared"
        Exit Sub
    End If

    Call SetZoteroCitationStyle(resp)

End Sub

'-------------------------------------------------------------------
' URL & HYPERLINK UTILITIES
'-------------------------------------------------------------------

' Engine: Finds all "http://" or "https://" text in a range and converts them to active hyperlinks.
Private Sub ConvertTextToHyperlinks(targetRange As Range)
    Dim p As Long
    Dim paraRange As Range
    Dim linkRange As Range
    Dim urlText As String

    #If Mac Then
        Dim ptxt As String
        Dim urlStart As Long
        Dim urlEnd As Long
        Dim searchPos As Long
        Dim ch As String
    #Else
        Dim re As Object
        Dim matches As Object
        Dim match As Object
        Dim i As Long
        Set re = CreateObject("VBScript.RegExp")
        With re
            .Global = True
            .IgnoreCase = True
            .Pattern = "https?://[^\s]+"
        End With
    #End If

    For p = 1 To targetRange.Paragraphs.Count
        Set paraRange = targetRange.Paragraphs(p).Range

        ' BUGFIX: the first paragraph of a field result can start at the field
        ' code, which shifts regex indices. Clip it to the visible result.
        If paraRange.Start < targetRange.Start Then paraRange.Start = targetRange.Start
        If paraRange.End > targetRange.End Then paraRange.End = targetRange.End

        #If Mac Then
            ptxt = paraRange.Text
            searchPos = 1
            Do
                urlStart = InStr(searchPos, ptxt, "http", vbTextCompare)
                If urlStart = 0 Then Exit Do

                urlEnd = 0
                If LCase(Mid(ptxt, urlStart, 7)) = "http://" Then
                    urlEnd = urlStart + 7
                ElseIf LCase(Mid(ptxt, urlStart, 8)) = "https://" Then
                    urlEnd = urlStart + 8
                End If

                If urlEnd = 0 Then
                    searchPos = urlStart + 1
                Else
                    Do While urlEnd <= Len(ptxt)
                        ch = Mid(ptxt, urlEnd, 1)
                        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then Exit Do
                        urlEnd = urlEnd + 1
                    Loop

                    urlText = Mid(ptxt, urlStart, urlEnd - urlStart)
                    Do While Len(urlText) > 0
                        ch = Right(urlText, 1)
                        If ch = "." Or ch = "," Or ch = ")" Then
                            urlText = Left(urlText, Len(urlText) - 1)
                        Else
                            Exit Do
                        End If
                    Loop

                    If Len(urlText) > 0 Then
                        Set linkRange = paraRange.Duplicate
                        linkRange.Start = paraRange.Start + urlStart - 1
                        linkRange.End = linkRange.Start + Len(urlText)

                        If linkRange.Hyperlinks.Count = 0 Then
                            On Error Resume Next
                            targetRange.Document.Hyperlinks.Add Anchor:=linkRange, Address:=urlText
                            On Error GoTo 0
                        End If
                        searchPos = urlStart + Len(urlText)
                    Else
                        searchPos = urlStart + 1
                    End If
                End If
            Loop
        #Else
            On Error Resume Next
            Set matches = re.Execute(paraRange.Text)
            If Err.Number <> 0 Then GoTo NextParagraph
            On Error GoTo 0

            For i = matches.Count - 1 To 0 Step -1
                Set match = matches(i)

                Set linkRange = paraRange.Duplicate
                linkRange.Start = paraRange.Start + match.FirstIndex
                linkRange.End = linkRange.Start + match.Length

                urlText = match.Value

                Do While Right(urlText, 1) = "." Or Right(urlText, 1) = "," Or Right(urlText, 1) = ")"
                    urlText = Left(urlText, Len(urlText) - 1)
                    linkRange.End = linkRange.End - 1
                Loop

                If linkRange.Hyperlinks.Count = 0 Then
                    On Error Resume Next
                    targetRange.Document.Hyperlinks.Add Anchor:=linkRange, Address:=urlText
                    On Error GoTo 0
                End If
            Next i
NextParagraph:
        #End If
    Next p
End Sub

' Public macro: Finds the Zotero bibliography and activates all URLs within it.
Public Sub ZCL_LinkBibliographyURLs(control As IRibbonControl)
    '--- Variable Declarations ---
    Dim i As Long
    Dim bibField As field
    
    '--- Initializations ---
    Application.ScreenUpdating = False
    Set bibField = Nothing

    ' 1. Find the Zotero bibliography field
    For i = ActiveDocument.Fields.Count To 1 Step -1
        If ActiveDocument.Fields(i).Type = wdFieldAddin Then
            If InStr(ActiveDocument.Fields(i).code, "ADDIN ZOTERO_BIBL") > 0 Then
                Set bibField = ActiveDocument.Fields(i)
                Exit For
            End If
        End If
    Next i

    ' 2. If the field is found, proceed
    If bibField Is Nothing Then
        ' ----- BURAYI DEGISTIRDIK -----
        Application.ScreenUpdating = True
        MsgBox "A Zotero bibliography could not be found in this document." & vbCrLf & vbCrLf & _
               "Please use the Zotero tab to 'Add/Edit Bibliography' first," & _
               "and then run this command again.", _
               vbExclamation, "Bibliography Not Found"
        Exit Sub
        ' ----- DEGISIKLIK SONU -----
    End If
    
    ' 3. Update the field (to get the latest version)
    bibField.Update
    
    ' BUGFIX: Give Word time to finish the update
    DoEvents
    
    ' 4. Call the engine on the *result* of the field
    Call ConvertTextToHyperlinks(bibField.result)
    
    Application.ScreenUpdating = True
    MsgBox "URLs in the bibliography have been activated.", vbInformation
End Sub

' Engine: Finds all email addresses in a range and converts them to "mailto:" hyperlinks.
Private Sub ConvertEmailsToHyperlinks(targetRange As Range)
    Dim linkRange As Range
    Dim baseStart As Long
    Dim docText As String

    On Error Resume Next
    docText = targetRange.Text
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0

    baseStart = targetRange.Start

    #If Mac Then
        Dim atPos As Long
        Dim startPos As Long
        Dim endPos As Long
        Dim ch As String
        Dim addr As String
        Dim dotPos As Long
        Dim tld As String

        atPos = 1
        Do
            atPos = InStr(atPos, docText, "@")
            If atPos = 0 Then Exit Do

            startPos = atPos - 1
            Do While startPos >= 1
                ch = Mid(docText, startPos, 1)
                If ch Like "[A-Za-z0-9._%+-]" Then
                    startPos = startPos - 1
                Else
                    Exit Do
                End If
            Loop
            startPos = startPos + 1

            endPos = atPos + 1
            Do While endPos <= Len(docText)
                ch = Mid(docText, endPos, 1)
                If ch Like "[A-Za-z0-9.-]" Then
                    endPos = endPos + 1
                Else
                    Exit Do
                End If
            Loop
            endPos = endPos - 1

            Do While endPos > atPos And Mid(docText, endPos, 1) = "."
                endPos = endPos - 1
            Loop

            If startPos < atPos And endPos > atPos Then
                addr = Mid(docText, startPos, endPos - startPos + 1)
                dotPos = InStrRev(addr, ".")
                If dotPos > InStr(addr, "@") And dotPos < Len(addr) Then
                    tld = Mid(addr, dotPos + 1)
                    If Len(tld) >= 2 And IsAlphaOnly(tld) Then
                        Set linkRange = targetRange.Document.Range( _
                            Start:=baseStart + startPos - 1, _
                            End:=baseStart + endPos)
                        If linkRange.Hyperlinks.Count = 0 Then
                            On Error Resume Next
                            targetRange.Document.Hyperlinks.Add Anchor:=linkRange, Address:="mailto:" & addr
                            On Error GoTo 0
                        End If
                    End If
                End If
            End If

            atPos = atPos + 1
        Loop
    #Else
        Dim re As Object
        Dim matches As Object
        Dim match As Object
        Dim i As Long

        Set re = CreateObject("VBScript.RegExp")
        With re
            .Global = True
            .IgnoreCase = True
            .Pattern = "([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})"
        End With

        Set matches = re.Execute(docText)
        If Err.Number <> 0 Then
            Set re = Nothing
            Exit Sub
        End If
        On Error GoTo 0

        For i = matches.Count - 1 To 0 Step -1
            Set match = matches(i)

            Set linkRange = targetRange.Document.Range( _
                Start:=baseStart + match.FirstIndex, _
                End:=baseStart + match.FirstIndex + match.Length)

            If linkRange.Hyperlinks.Count = 0 Then
                On Error Resume Next
                targetRange.Document.Hyperlinks.Add Anchor:=linkRange, Address:="mailto:" & match.Value
                On Error GoTo 0
            End If
        Next i

        Set match = Nothing
        Set matches = Nothing
        Set re = Nothing
    #End If

    Set linkRange = Nothing
End Sub

Private Function IsAlphaOnly(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If Not (Mid(s, i, 1) Like "[A-Za-z]") Then Exit Function
    Next i
    IsAlphaOnly = True
End Function

' Public macro: Finds and activates all email addresses in the entire document.
Public Sub ZCL_LinkEmails(control As IRibbonControl)
    
    Application.ScreenUpdating = False
    
    Call ConvertEmailsToHyperlinks(ActiveDocument.Content)
    
    Application.ScreenUpdating = True
    MsgBox "All email addresses in the document have been activated.", vbInformation
End Sub
