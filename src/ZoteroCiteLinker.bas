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
Public Const ZCL_VERSION As String = "2.1.0-mac-beta"

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

' Checks if the given Zotero style ID is in the hardcoded supported list.
Private Function isSupportedStyle(ByVal style As String) As Boolean
    Dim predefinedList As String
    predefinedList = "|" & _
        "physical-review-letters|american-physics-society|" & _
        "molecular-plant|ieee|apa|vancouver|american-chemical-society|" & _
        "american-medical-association|nature|american-political-science-association|" & _
        "american-sociological-association|chicago-author-date|bmc-medicine|" & _
        "china-national-standard-gb-t-7714-2015-numeric|" & _
        "china-national-standard-gb-t-7714-2015-author-date|" & _
        "harvard-cite-them-right|elsevier-harvard|modern-language-association|" & _
        "archives-of-computational-methods-in-engineering|" & _
        "elsevier-vancouver|"
    style = "|" & style & "|"
    isSupportedStyle = InStr(1, predefinedList, style, vbTextCompare) > 0
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
        
        Case "physical-review-letters", "american-physics-society", "china-national-standard-gb-t-7714-2015-numeric", "bmc-medicine", "ieee", "archives-of-computational-methods-in-engineering", "elsevier-vancouver"
            Call ExtractSerialNumberCitations(field, citations, "[]")

        Case "american-chemical-society", "american-medical-association", "nature"
            Call ExtractSerialNumberCitations(field, citations, "")

        Case Else
            Err.Raise vbObjectError + 1, "ExtractCitations", "Citation style not recognized: " & style
    End Select
End Sub

' Displays a MsgBox listing all hardcoded supported citation styles.
Public Sub ZCL_Information(control As IRibbonControl)
    '--- Variable Declarations ---
    Dim styles As Variant
    Dim msg As String
    Dim i As Long
    
    '--- Initializations ---
    styles = Split("physical-review-letters|american-physics-society|molecular-plant|ieee|apa|vancouver|american-chemical-society|american-medical-association|nature|" & _
                   "american-political-science-association|american-sociological-association|chicago-author-date|bmc-medicine|" & _
                   "china-national-standard-gb-t-7714-2015-numeric|china-national-standard-gb-t-7714-2015-author-date|" & _
                   "harvard-cite-them-right|elsevier-harvard|modern-language-association|" & _
                   "archives-of-computational-methods-in-engineering|elsevier-vancouver", "|")
    
    msg = "ZoteroCiteLinker " & ZCL_VERSION & vbCrLf & vbCrLf & "Zotero styles supported by this macro:" & vbCrLf & vbCrLf
    
    For i = 0 To UBound(styles)
        msg = msg & i + 1 & ". " & styles(i) & vbCrLf
    Next
    
    MsgBox msg, vbInformation, "Supported Reference Styles"
End Sub

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

    '--- Initializations ---
    Set prefs = GetZoteroPrefs()
    Set bibField = Nothing
    iCount = 0
    
    ' Do not support Bookmark-type citations
    If Not prefs("pref-fieldType") = "Field" Then
        MsgBox "Only support 'Fields' type citations" & vbCrLf & vbCrLf & _
            "Click on Document Preferences in Word, then expand the Advanced Options section. " & _
            "This will allow you to verify whether the 'Store citation as bookmarks' option was accidentally enabled.", _
            vbCritical, "Error"
        Exit Sub
    End If

    styleId = prefs("style-id")
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

    ' Iterate through all fields in the document
    For Each aField In targetFields
        ' Check if the field is a Zotero citation
        If aField.Type = wdFieldAddin Then
            If InStr(aField.code, "ADDIN ZOTERO_ITEM") > 0 Then

                If debugging Then
                    ' Focus to next field
                    Application.ScreenUpdating = True
                    ActiveWindow.ScrollIntoView aField.result, True
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
                On Error GoTo 0
                
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
                    On Error GoTo 0
                
                Else
                    ' LOGIC: GROUPED CITATION OR EMPTY FIELD (This is a grouped citation (e.g., [1, 5-7]) or an empty/broken field.)
                    Call ExtractCitations(aField, cits, styleId)
                
                    If UBound(cits) = -1 Then GoTo SkipToNextField
                End If

                For i = 0 To UBound(cits)
                    cit = cits(i)
                    Set rng = aField.result.Document.Range(Start:=cit.Start, End:=cit.End)
                    tempBookmarkName = "ZoteroLinkCitationTempBookmark" & i
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
                        Anchor:=ActiveDocument.Bookmarks("ZoteroLinkCitationTempBookmark" & i).Range, _
                        SubAddress:=titleAnchor, ScreenTip:=titleToSearch) 'ScreenTip set displayed information of hyperlink
                    hp.Range.Font.Underline = wdUnderlineNone 'Remove hyperlink underline
                    
                    ' Apply text style to the hyperlink
                    If userTextStyle <> "" Then
                        hp.Range.style = ActiveDocument.styles(userTextStyle)
                    End If

                    iCount = iCount + 1

SkipToNextCitation:
                    ActiveDocument.Bookmarks("ZoteroLinkCitationTempBookmark" & i).Delete

                Next i

            End If
        End If
SkipToNextField: ' Label for skipping a field (e.g., in debug mode or for empty fields)
    Next aField
    
    Call SetZoteroCitationColor

ExitTheMacro:

    If notify Then MsgBox "Linked " & iCount & " Zotero citations.", vbInformation, "Finish"

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
    ActiveWindow.ScrollIntoView originalRng, True
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
    ActiveWindow.ScrollIntoView originalRng, True
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
