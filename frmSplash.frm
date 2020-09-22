VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Splash Screen"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4830
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer tmrFade 
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblSplash 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label array for captions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   2850
   End
   Begin VB.Label lblSplash 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label array for captions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   2850
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'   Programmer:     Brent Culpepper a.k.a. IDontKnow (See Credits)
'   Project:        Custom splash screen with fade in/out option
'   Date:           May 22, 2004
'*****************************************************************************************
'   Credits:
' ------------------------------------------------------------------------------
'
' The custom region code in this project comes directly from a demo created
' by "The Hand", Garrett Sever (garrett@elitevb.com), and posted at
' http://www.EliteVB.com. The name of the demo is "Custom complex window region
' demo". I would urge everybody to visit EliteVB, it is a fantastic resource!
'
' ------------------------------------------------------------------------------
'
' The custom region code was modified by LazyJay from http://www.vbCity.com.
' (my home away from home!) It is now much faster than the original version.
'
' ------------------------------------------------------------------------------
'
' Credit for the method of using SetLayeredWindowAttributes to fade the form
' in and out goes to the http://www.vbnet.mgps.org website. Another essential site!
'
' ------------------------------------------------------------------------------
'
' My sincere thanks to all these gentlemen! Basically all that I have done
' is to combine their work into a form that will be easy to reuse between
' projects. All credit goes to them, any mistakes are mine alone ;)
'
'*****************************************************************************************
'
' My advice to people who are happy with this splashform and want to use it in
' their own projects is this: Remove the picture I used as a demo from the form
' (Unless you happen to need a skull as your program's "first impression"). Then
' save the form to your template folder with a distinctive name like "Custom Splash".
' (Form templates are found at "Microsoft Visual Studio\VB98\Template\Forms").
' Now all you need to do to add it to a project is to select 'Add Form' from the
' VB menu and then select "Custom Splash" from the dialog box that appears. Be
' sure and rename/resave the form in your project so you don't accidentally overwrite
' your template. Add your choice of bitmap to the form and rock-n-roll!
' ------------------------------------------------------------------------------
'
' One more thing to be aware of. I didn't want to slow down the code by normalizing
' the transparent color. The picture that I used in the demo was edited using Paint
' to make sure the magenta mask color I used was consistent. I figured it was better
' to do a little manual editing than slow down the code ;)
' ------------------------------------------------------------------------------
' Usage Example:
'
'   With frmSplash
'       .Duration = 8           ' Can use either seconds or milliseconds here
'       .FadeSpeed = 75         ' This is a pretty slow fade speed
'       .DialogAction fadeInOut ' Fade in/out. This line also results in the form being loaded in memory
'       .Fixed = False          ' The form is moveable. If true the position is fixed.
'       .Show
'   End With
' ------------------------------------------------------------------------------
' The options for DialogAction (fading) are as follows:
'
'   1. fadeNone : no fade
'   2. fadeIn : The form will fade open and close normally
'   3. FadeOut : The form opens normal and fades closed
'   4. FadeInOut: The form fades both when opening and closing
' ------------------------------------------------------------------------------
' Note about the Duration property:
'
' If the duration is > 0 then the timer that closes the form is enabled, and the
' form will be closed when the timer expires. If you don't want to use the timer
' to close the splash form, just assign a value of zero to the Duration property.
' The form will remain open until either dblclicked or unloaded by code.
' ------------------------------------------------------------------------------
' Supported Platforms:
'
' Unfortunately SetLayeredWindowAttributes is not supported prior to W2K. The
' DialogAction sub will check the version to see if fade is supported; if not,
' the mode will be changed to fadeNone. You will end up with a custom shaped
' form, but it will only open and close normally. One 'feature'(spelled b-u-g)
' I have found with using fadeNone is this--If you move the form you will see
' ghost-image trails from the graphics. If somebody finds a fix for this, great!
' If not, I would suggest setting the Fixed property to True for pre-W2K users.
'*****************************************************************************************
Option Explicit

' Declares for FormOnTop
'*****************************************
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_CAPTIONTEXT = 9
Private Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Declares for manipulation of the regions
'*****************************************
Private Const RGN_AND = 1   'Creates the intersection of the two combined regions.
Private Const RGN_COPY = 5  'Creates a copy of the region identified by hrgnSrc1.
Private Const RGN_OR = 2    'Creates the union of two combined regions.
Private Const RGN_XOR = 3   'Creates the union of two combined regions except for any overlapping areas.
Private Const RGN_DIFF = 4  'Combines the parts of hrgnSrc1 that are not part of hrgnSrc2.
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

' Declares for color retrieval
'*****************************************
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

' Declares to get system settings
'*****************************************
Const SM_CYCAPTION = 4      'Height of windows caption
Const SM_CXBORDER = 5       'Width of no-sizable borders
Const SM_CYBORDER = 6       'Height of non-sizable borders
Const SM_CXDLGFRAME = 7     'Width of dialog box borders
Const SM_CYDLGFRAME = 8     'Height of dialog box borders
Private Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long

' Declares for cleaning up GDI32 objects
'*****************************************
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

' Declares for layered window/transparency
'*****************************************
Public Enum dlgShowActions
   fadeNone = 0
   fadeIn = 1
   fadeOut = 2
   fadeInOut = 3
End Enum

Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_RIGHT As Long = &H1000
Private Const WS_EX_LEFTSCROLLBAR As Long = &H4000
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_COLORKEY As Long = &H1
Private Const LWA_ALPHA As Long = &H2
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long

' Version decection to see if layered windows are supported
Private Declare Function GetVersion Lib "kernel32" () As Long

' Private member constants & variables
'*******************************************
' True for 2K and above
Private m_bSupportsLayeredWindows As Boolean

' Speed of the fade loop. The ideal speed is between 20 - 100
Const DEFAULT_FADESPEED As Integer = 50
Private m_FadeSpeed     As Integer

' Time in milliseconds before unloading the form. If the
' value is zero the form will not be unloaded by the timer.
Private m_Duration      As Integer

' Stores the handle of the window region.
Private oldRgn          As Long

' X-Y positions for moving the form
Private XPos            As Long
Private YPos            As Long
Private m_Fixed         As Boolean

' Stores the mode used for loading and unloading
Private unloadAction    As dlgShowActions
Private fadeMode        As dlgShowActions

' General Info variables used in most splash screens
Private m_strTitle      As String
Private m_strVersion    As String



'*******************************************
'       Custom Properties
'*******************************************
Public Property Get FadeSpeed() As Integer
    FadeSpeed = m_FadeSpeed
End Property

Public Property Let FadeSpeed(ByVal NewFadeSpeed As Integer)
' FadeSpeed is the millisecond interval time used by the fade
' timer. Don't get silly with it! A good speed is somewhere in
' the range of 20 to 100 or so. Experiment until you get
' the results that you are after.
    m_FadeSpeed = NewFadeSpeed
End Property

Public Property Get Duration() As Integer
    Duration = m_Duration
End Property

Public Property Let Duration(ByVal NewDuration As Integer)
' This property is the duration to show the form before unloading.
' If the value is zero then the timer is not used to unload the
' form. Otherwise we check the value and see if the user assigned
' the value in seconds or milliseconds. If the value is in seconds,
' convert it before assigning to the variable.
    If NewDuration > 1000 Then
        ' Assume milliseconds
        m_Duration = NewDuration
    Else
        m_Duration = NewDuration * 1000
    End If
End Property

Public Property Get Fixed() As Boolean
    Fixed = m_Fixed
End Property

Public Property Let Fixed(ByVal NewFixed As Boolean)
' If the property is True then don't allow the form to be moved
    m_Fixed = NewFixed
End Property

Private Sub Form_Initialize()
    
    Debug.Print m_FadeSpeed
    If m_FadeSpeed = 0 Then m_FadeSpeed = DEFAULT_FADESPEED
    Debug.Print m_FadeSpeed
End Sub

Private Sub Form_Load()
    
    Dim X As Long 'Used to loop thru the pixels
    Dim Y As Long 'Used to loop thru the pixels
    Dim wid As Long 'Width of picture - Used to loop thru the pixels
    Dim hgt As Long 'Height of picture - Used to loop thru the pixels
    Dim ttlHeight As Long 'Height of the form's titlebar & top border in pixels
    Dim xBorder As Long 'Width of the form's side borders in pixels
    Dim rgnPic As Long 'Region of the picture
    Dim rgnPixel As Long 'Region of a pixel - used to subtract out tiny areas
    Dim colPixel As Long 'Color of a pixel in the picture
    Dim picDC As Long 'Temporary device context used to get pixel color info
    Dim oldBmp As Long '1x1 bitmap created when picDC is created.
    Dim transColor As Long
    Dim rgnStartX As Long
    
    ' Center form and place on top of other forms:
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    FormOnTop True
    
    ' The following are just general-use strings that can be used with labels.
    ' Placed here as a convenience so modify or delete as you see fit.
    m_strVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    m_strTitle = App.Title
    
    ' Calculate the size of the picture which we will fit the form to
    wid = Me.ScaleX(Me.Picture.Width, vbHimetric, vbPixels)
    hgt = Me.ScaleX(Me.Picture.Height, vbHimetric, vbPixels)

    ' Create a region the same size as our picture dimensions
    rgnPic = CreateRectRgn(0, 0, wid, hgt)

    ' Select our picture into a temporary device context so we can
    ' read the color information.
    picDC = CreateCompatibleDC(Me.hdc)
    oldBmp = SelectObject(picDC, Me.Picture.Handle)
     
    transColor = GetPixel(picDC, 0, 0)
     
    ' Loop thru all pixels in the picture
    For Y = 0 To hgt
        rgnStartX = -1
        For X = 0 To wid
            ' check the color of each pixel
            colPixel = GetPixel(picDC, X, Y)
            If colPixel = transColor Then
                If rgnStartX = -1 Then
                    rgnStartX = X
                End If
            Else
                If rgnStartX > -1 Then
                    ' If the color is our mask color then create a tiny
                    ' region for it and remove it from the picture
                    rgnPixel = CreateRectRgn(rgnStartX, Y, X, Y + 1)
                    CombineRgn rgnPic, rgnPic, rgnPixel, RGN_XOR
                    ' Clean up our graphics resource
                    DeleteObject rgnPixel
                    rgnStartX = -1
                End If
            End If
        Next X
        If rgnStartX > -1 Then
            ' If the color is our mask color then create a tiny
            ' region for it and remove it from the picture
            rgnPixel = CreateRectRgn(rgnStartX, Y, X, Y + 1)
            CombineRgn rgnPic, rgnPic, rgnPixel, RGN_XOR
            ' Clean up our graphics resource
            DeleteObject rgnPixel
        End If
    Next Y
             
    ' Clean up our temporary picture resources
    SelectObject picDC, oldBmp
    DeleteDC picDC
    DeleteObject oldBmp
     
    ' Calculate how much we need to offset the region so it lays directly ontop
    ' of the form's picture (client x,y instead of form x,y)
    ttlHeight = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYDLGFRAME)
    xBorder = GetSystemMetrics(SM_CXDLGFRAME)
     
    ' Offset our region by the calculated amount
    OffsetRgn rgnPic, xBorder, ttlHeight
     
    ' Fit the window to our new region
    ' If its the first time, store the original handle.
    ' If its more than the first time, just delete the previous custom handle.
    If oldRgn = 0 Then
        oldRgn = SetWindowRgn(Me.hwnd, rgnPic, True)
    Else
        DeleteObject SetWindowRgn(Me.hwnd, rgnPic, True)
    End If
    
    ' Vanity Plate, just to demonstrate using labels
    ' to pass events to the form. The label is a control
    ' array, so if you are using multiple labels for name,
    ' version, copyright, etc, this way the form will
    ' receive mouse events from any/all label controls.
    ' This allows the form to be moved even if the mouse
    ' event occurs on a label. Label dblclicks will also
    ' unload the form just like form dblclicks.
    lblSplash(0).Caption = "  " & m_strTitle & vbNewLine & "by IDontKnow"
    lblSplash(1).Caption = m_strVersion
    
    ' Check if we are using the timer to close the form.
    ' If we are then set the interval and enable the timer.
    If m_Duration <> 0 Then
        tmrShow.Interval = m_Duration
        tmrShow.Enabled = True
    End If
    
End Sub

Private Sub FormOnTop(bOnTop As Boolean)
    Select Case bOnTop
        Case True
            SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
        Case False
            SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
    End Select
End Sub

Private Sub lblSplash_DblClick(Index As Integer)
' Pass any dblclicks from the label array to the form.
    Call Form_DblClick
End Sub

Private Sub tmrShow_Timer()
    Unload Me
End Sub

Private Sub Form_DblClick()
    ' Allow the user to close the form in case
    ' they don't want to wait for it to time-out.
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' NOTE: We use this for moving the form. Since most splash screens
' will include labels with name, version information, etc, we also
' duplicate this code in the lblSplash label array. This way no
' matter where the user clicks we can still respond.
    XPos = X
    YPos = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Fixed Then Exit Sub
    If Button = vbLeftButton Then
        Me.Left = Me.Left - (XPos - X)
        Me.Top = Me.Top - (YPos - Y)
    End If
End Sub

Private Sub lblSplash_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    XPos = X
    YPos = Y
End Sub

Private Sub lblSplash_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Fixed Then Exit Sub
    If Button = vbLeftButton Then
        Me.Left = Me.Left - (XPos - X)
        Me.Top = Me.Top - (YPos - Y)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Set the old window region back to the original
    ' (and delete our custom region.)
    DeleteObject SetWindowRgn(Me.hwnd, oldRgn, True)
End Sub

Private Sub GetWindowsVersion(Optional ByRef lMajor = 0, _
                              Optional ByRef lMinor = 0, _
                              Optional ByRef lRevision = 0, _
                              Optional ByRef lBuildNumber = 0, _
                              Optional ByRef bIsNt = False)
    Dim lR As Long
    lR = GetVersion()
    lBuildNumber = (lR And &H7F000000) \ &H1000000
    If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
    lRevision = (lR And &HFF0000) \ &H10000
    lMinor = (lR And &HFF00&) \ &H100
    lMajor = (lR And &HFF)
    bIsNt = ((lR And &H80000000) = 0)
End Sub

'****************************************************************
' NOTE: Thanks to http://www.vbnet.mgps.org for the excellent
' SetLayeredWindowAttributes demo used in this project!
'****************************************************************

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' If the user presses the close button
    ' and the 'out' mode is fade, cancel
    ' the close and instead invoke the
    ' timer to cause the fade.
    '
    ' The timer code changes the unloadAction
    ' value to prevent this check from executing
    ' again when the timer code issues the
    ' Unload command.
    If ((UnloadMode = vbFormControlMenu) Or _
        (UnloadMode = vbFormCode)) And _
        (unloadAction = fadeOut) Or _
        (unloadAction = fadeInOut) Then
        Cancel = True
        fadeMode = fadeOut
        tmrFade.Interval = m_FadeSpeed
        tmrFade.Enabled = True
    End If
End Sub

Private Sub tmrFade_Timer()
' Change the alpha value to fade the form in/out
    Static fadeValue As Long
    Dim alpha As Long
   
    Select Case fadeMode
        Case fadeOut:
            ' Prevents the form's QueryUnload sub
            ' from stopping the unloading of the
            ' form via code here
            unloadAction = 0
            If (fadeValue + (256 * 0.05)) >= 256 Then
                ' Done, so reset the fadeValue to
                ' allow for fading out if required
                tmrFade.Enabled = False
                fadeValue = 0
                Unload Me
                Exit Sub
            End If
            fadeValue = fadeValue + (256 * 0.05)
            alpha = (256 - fadeValue)
      
        Case fadeIn:
            If (fadeValue + (256 * 0.05)) >= 256 Then
                ' Done, but one more call to
                ' SetLayeredWindowAttributes is
                ' required to set the final opacity to 255
                tmrFade.Enabled = False
                fadeValue = 0
                alpha = 255
            Else
                fadeValue = fadeValue + (256 * 0.05)
                alpha = fadeValue
            End If
      
        Case Else
    End Select

    SetLayeredWindowAttributes Me.hwnd, 0&, alpha, LWA_ALPHA
End Sub

Private Function AdjustWindowStyle()
    Dim style As Long
    ' In order to have transparent windows, the
    ' WS_EX_LAYERED window style must be applied
    ' to the form
    style = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
   
    If Not (style And WS_EX_LAYERED = WS_EX_LAYERED) Then
        style = style Or WS_EX_LAYERED
        SetWindowLong Me.hwnd, GWL_EXSTYLE, style
    End If
End Function

Public Sub DialogAction(dlgEffectsMethod As dlgShowActions)
    Dim alpha As Long
    ' Alpha=0: window transparent
    ' Alpha=255: window opaque
    
    ' Version detection:
    Dim lMajor As Long
    
    GetWindowsVersion lMajor
    If lMajor >= 5 Then
        m_bSupportsLayeredWindows = True
    Else
        dlgEffectsMethod = fadeNone
    End If
    
    If m_FadeSpeed = 0 Then m_FadeSpeed = DEFAULT_FADESPEED
    
    Select Case dlgEffectsMethod
        Case fadeNone  ' Show 'normally'
            ' Nothing to do, so exit and let
            ' the calling routine's Show command
            ' control the display
        
        Case fadeOut ' Show normally but prepare for a fade out
            ' This requires changing the window style
            ' and calling SetLayeredWindowAttributes once
            ' specifying a value of opaque (255). To
            ' cause the form to fade out, an 'unloadAction'
            ' flag is set
            unloadAction = dlgEffectsMethod
            Call AdjustWindowStyle
            alpha = 255
            SetLayeredWindowAttributes Me.hwnd, 0&, alpha, LWA_ALPHA
      
        Case fadeIn, fadeInOut ' Show form by fading in
            ' Just adjust the window style and
            ' use a timer to fade the window in
            Call AdjustWindowStyle
            fadeMode = fadeIn
            tmrFade.Interval = m_FadeSpeed
            tmrFade.Enabled = True
            ' But ... if the effect mode is to
            ' fade in/out, set the unloadAction flag
            If dlgEffectsMethod = fadeInOut Then unloadAction = fadeOut
    End Select
End Sub
