VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZCLAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Watches the two moments where stale links become visible: a document is
' opened, and a document window becomes active. Both handlers only compare
' a cheap fingerprint; the work itself stays in ZoteroCiteLinker.
Option Explicit

Public WithEvents App As Word.Application

Private Sub App_DocumentOpen(ByVal Doc As Document)
    On Error Resume Next
    ZCL_OnDocumentEvent Doc, False
End Sub

Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
    On Error Resume Next
    ZCL_OnDocumentEvent Doc, True
End Sub
