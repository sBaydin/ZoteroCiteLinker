VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CustomDictionary - Dictionary replacement for macOS.
' macOS Word cannot create Scripting.Dictionary, so the main module uses this
' class there. It exposes the same members the caller needs:
' Add, Item, Keys, Count and Exists.

Option Explicit

Private mKeys() As Variant
Private mValues() As Variant
Private mCount As Long
Private mCapacity As Long

Private Sub Class_Initialize()
    mCount = 0
    mCapacity = 16
    ReDim mKeys(1 To mCapacity)
    ReDim mValues(1 To mCapacity)
End Sub

Private Sub EnsureCapacity()
    If mCount >= mCapacity Then
        mCapacity = mCapacity * 2
        ReDim Preserve mKeys(1 To mCapacity)
        ReDim Preserve mValues(1 To mCapacity)
    End If
End Sub

Private Function FindKey(ByVal key As Variant) As Long
    Dim i As Long
    For i = 1 To mCount
        If mKeys(i) = key Then
            FindKey = i
            Exit Function
        End If
    Next i
    FindKey = 0
End Function

Public Sub Add(ByVal key As Variant, ByVal value As Variant)
    EnsureCapacity
    mCount = mCount + 1
    mKeys(mCount) = key
    mValues(mCount) = value
End Sub

Public Property Get Item(ByVal key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim idx As Long
    idx = FindKey(key)
    If idx > 0 Then
        Item = mValues(idx)
    Else
        Item = Empty
    End If
End Property

Public Property Let Item(ByVal key As Variant, ByVal value As Variant)
    Dim idx As Long
    idx = FindKey(key)
    If idx > 0 Then
        mValues(idx) = value
    Else
        Add key, value
    End If
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Public Function Keys() As Variant
    If mCount = 0 Then
        Keys = Array()
        Exit Function
    End If
    Dim result() As Variant
    ReDim result(0 To mCount - 1)
    Dim i As Long
    For i = 1 To mCount
        result(i - 1) = mKeys(i)
    Next i
    Keys = result
End Function

Public Function Exists(ByVal key As Variant) As Boolean
    Exists = (FindKey(key) > 0)
End Function
