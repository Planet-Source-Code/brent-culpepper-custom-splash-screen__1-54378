Attribute VB_Name = "modStartUp"
Option Explicit

Public Sub Main()
    ' Set the properties for the custom splash screen

    With frmSplash
        .Duration = 8           ' Can use either seconds or milliseconds here
        .FadeSpeed = 80         ' This is a pretty slow fade speed
        .DialogAction fadeInOut ' Fade in/out. This line also results in the form being loaded in memory
        .Show
    End With
    
    ' You can load your main form either from here or
    ' from the splash screen's Unload event. Since the
    ' splash form is on top, it won't hurt anything
    ' having your other form(s) loading in the background.
    frmMain.Show
        
End Sub
