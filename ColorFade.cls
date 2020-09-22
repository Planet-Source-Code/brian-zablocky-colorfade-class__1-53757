VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorFade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ====================================================================
'
'   For those who suck at math, you might want to read this!
'
'   ColorFade produces what I like to call a "perfect interpolation"
'   between two colors.  Individual reds, greens, and blues are faded
'   as separate parameters and then recombined to form a color that
'   is somewhere in between.
'
'   Lets say you have two colors, red and blue.  School taught us the
'   result of that is purple.  Now, what if you want to mix a 60/40,
'   or a 35/65, or 92/8.  These ratio get harder to calculate.  This
'   class will do that work for you.
'
'   PLEASE NOTE:  I did program this for speed.  On a P3-450 with 96mb
'   of ram you will not notice the execution speed, however, several
'   optimizations could be made to speed things up.
'
'
'   PS.  There was a radial gradient, but no matter what I did, it
'   only smoothed when the form was resized.  This was unacceptable
'   to me as a good outcome so I scrapped it.
'
' ====================================================================
'
'   USAGE:  Step    1:  Declare It:
'                               Dim Fade As New ColorFade
'
'                   2:  Set Start Color:
'                               Fade.StartColor = RGB(128, 128, 0)
'
'                   3:  Set End Color:
'                               Fade.EndColor = RGB(0, 128, 128)
'
'                   4:  Get a Color!
'                               c = Fade.GetFadeColor(45)
'                                                     ||
'                                                     ||
'           NOTE:  This is a percentage! -------------++
'
'
'
'
' ====================================================================
'
'  NOTE:  Commenting this class is bottom priority for me.
'
' ====================================================================

Option Explicit



Private mvarFadeStartColor As Long  'Start Color
Private mvarFadeEndColor As Long    'End Color

Private fcsr As Long    'Red Start
Private fcsg As Long    'Green Start
Private fcsb As Long    'Blue Start

Private fcer As Long    'Red End
Private fceg As Long    'Green End
Private fceb As Long    'Blue End




Enum ColorFadeGradientConstants
    ' These are the currently implemented gradients
    ' for the PaintObj methods.  If you add a gradient
    ' here you must allow program it in to the other
    ' methods below.  Good luck...
    
    HorizontalGradient = 0
    VerticalGradient = 1
    DiagUpperLeftGradient = 2
    DiagUpperRightGradient = 3
    DiagLowerLeftGradient = 4
    DiagLowerRightGradient = 5
    '
    ' Other Gradients, if you have a day or two...
    '
End Enum







Public Function GetFadeColor(Optional fcFadePercent As Long = 50) As Long
    ' This is the original purpose of the class.  I wanted to specify a
    ' start and end color, then have the computer tell me what was in
    ' between.  I decided on a percentile scale for this.
    '
    ' For all intensive purposes, 0% is the start color and 100% is the
    ' end color.  You could use this to fade text in a rich text box.
    '
    ' This method is now a wrapper for the DetermineFade function.  I
    ' did this so a programmer could make more methods using the
    ' functionality of this class.
    '
    GetFadeColor = DetermineFade(fcsr, fcsg, fcsb, fcer, fceg, fceb, fcFadePercent)
    
End Function




Public Function GetFadeColor2(Optional fcFadePercent As Long = 50, Optional fcSegments As Long = 2) As Long
    '
    ' Returns a segmented linear fade.  By segmented, I mean that you start
    ' and end with the start color.  The end color is in between somewhere.
    '
    ' This is a novel concept. 0% and 100% is the startcolor.  For every
    ' "segment" you request, one peak is made.  The actual width of the peaks
    ' is fixed based on your monitor resolution, so 1 peak could mean two
    ' or three on your form.  This makes it look more uniform when the user
    ' resizes their form.
    '
    ' The end color is 50% on 1 segment, 25% and 75% on two segments, and so
    ' on.  You can still use "For i = 0 to 100" to make your gradients.
    '
    '
    Dim fcs As Long, fcp As Long, s As Single, p As Single
    Dim np As Long

    fcs = fcSegments                'Maybe I should of ByVal'd, but who cares?
    If fcs > 100 Then fcs = 100     'More than 12 segments looks like crap
    If fcs < 1 Then fcs = 1         '0 segments makes no fade here

    fcp = fcFadePercent             'Another byval waiting to happen
    If fcp > 100 Then fcp = 100     'More than 100% could cause an error
    If fcp < 1 Then fcp = 1         'Same with less than 1% for the simple
                                    'fact that the segmented counter was
                                    'programmed for speed

    s = CSng(100 / fcs)             'Get a fraction for percentage perposes.
    p = CSng(Int(fcp / s) + 1)      'Determine the segment based on the percent
                                    'With 2 segments, 60 percent falls in the 2nd
                                    'segment.  We need to know which one we are in.
    
    ' This next chunk is very hard to explain.  The program needs to pull a
    ' real percentage based on the segment count and the current segment in
    ' order to use the DetermineColor function.  So these formulas will take
    ' care of that problem.
    If p = 1 Then
        'If we are in the first segment, we can just make a real percentage
        'based on the Segment Count and the requested percentage.  This will
        'save some processing time.
        np = fcp * fcs
    Else
        'Since we are in a larger segment, we need to find out what the real
        'percentage is.  It's a complex and hard to explain formula.
        np = (fcp - ((p - 1) * s)) * fcs    'Don't Ask.
    End If
    
    
    ' Now we get the color.  If the segment is an even number then we are on
    ' our way down the mountain, so we need a color somewhere between End and
    ' start.  Odd numbers are going up so we get it from Start to End.
    If (p / 2) = Int(p / 2) Then
            ' Even Segment - End-Start
            GetFadeColor2 = DetermineFade(fcer, fceg, fceb, fcsr, fcsg, fcsb, np)
    Else
            ' Odd Segment - Start-End
            GetFadeColor2 = DetermineFade(fcsr, fcsg, fcsb, fcer, fceg, fceb, np)
    End If
        
End Function









Public Sub PaintObj(Obj As Object, Gradient As ColorFadeGradientConstants)
    '
    ' Wrapper function for the private PaintObject function.  You can
    ' specify other types of paintable objects here.  The only actual
    ' requirement is that the object MUST support the Line method using
    ' VBs syntax.
    '
    If (TypeOf Obj Is Form) Or (TypeOf Obj Is PictureBox) Then
        ' We only support forms and picture boxes for now.
        PaintObject Obj, Gradient
    End If

End Sub



Public Sub PaintObj2(Obj As Object, Gradient As ColorFadeGradientConstants, Optional CycleCount As Long = 1)
    '
    ' Wrapper function for the Segmented PaintObject function.  Same comments
    ' as the method above.
    '
    If (TypeOf Obj Is Form) Or (TypeOf Obj Is PictureBox) Then
        ' We only support forms and picture boxes for now.
        PaintObject2 Obj, Gradient, CycleCount
    End If
        
End Sub









Private Sub PaintObject(frm As Object, Gradient As ColorFadeGradientConstants)
    ' This is the gradient method for objects.  Use this to automate the
    ' drawing of gradients using VBs Line method.  This function does almost
    ' no error checking so make sure that you keep it private and wrap it
    ' with a public function instead.
    
    Dim ScreenX As Long     'Screen dimension
    Dim ScreenY As Long     'Screen dimension
    Dim h As Long           'Object Height
    Dim w As Long           'Object Width
    Dim y As Long           'Current Y-Position
    Dim x As Long           'Current X-Position
    Dim a As Long           'Current Angle Offset Multiplier
    Dim ax As Long          'Current Angle Offset
        
    Dim x1 As Long, x2 As Long  'Left and Right of the object
    Dim y1 As Long, y2 As Long  'Top and Bottom of the object
    
    Dim i As Long   'Loop Counter
    Dim j As Long   'Loop Counter
    Dim c As Long   'Cycles or other value
    
    
    ScreenX = Screen.TwipsPerPixelX 'Determine the horizontal screen dimension
    ScreenY = Screen.TwipsPerPixelY 'Determine the vertical screen dimension

    x1 = 0: x2 = (frm.Width / ScreenX)  'Get the pixel width of the object
    y1 = 0: y2 = (frm.Height / ScreenY) 'Get the pixel height of the object
    
    ax = (y2 / 2)   'Since we are dealing with 45 degree angles, the horizontal
                    'offset of the object will be half of the vertical size.
                    'This sounds tough to believe but trust me on this.


    ' This long chunk of code is where the gradients are performed.  Based on
    ' the data above, you could make your own gradients.  For example, with
    ' very little work you could add a 20 degree angled gradient.
    Select Case Gradient
        Case HorizontalGradient
            For i = y1 To y2
                y = (i / y2) * 100
                frm.Line (x1 * ScreenX, i * ScreenY)-(x2 * ScreenX, i * ScreenY), Me.GetFadeColor(y)
            Next i
            
        
        Case VerticalGradient
            For i = x1 To x2
                x = (i / x2) * 100
                frm.Line (i * ScreenX, y1 * ScreenY)-(i * ScreenX, y2 * ScreenY), Me.GetFadeColor(x)
            Next i
            
            
        Case DiagLowerLeftGradient
            For i = x1 - ax To x2 + ax
                x = (i / x2) * 100
                frm.Line ((i - ax) * ScreenX, y1 * ScreenY)-((i + ax) * ScreenX, y2 * ScreenY), Me.GetFadeColor(x)
            Next i
        
        Case DiagUpperLeftGradient
            For i = x2 + ax To x1 - ax Step -1
                x = (i / x2) * 100
                frm.Line ((i + ax) * ScreenX, y1 * ScreenY)-((i - ax) * ScreenX, y2 * ScreenY), Me.GetFadeColor(x)
            Next i
            
        Case DiagUpperRightGradient
            For i = x1 - ax To x2 + ax
                x = 100 - ((i / x2) * 100)
                frm.Line ((i - ax) * ScreenX, y1 * ScreenY)-((i + ax) * ScreenX, y2 * ScreenY), Me.GetFadeColor(x)
            Next i
            
        Case DiagLowerRightGradient
            For i = x2 + ax To x1 - ax Step -1
                x = 100 - ((i / x2) * 100)
                frm.Line ((i + ax) * ScreenX, y1 * ScreenY)-((i - ax) * ScreenX, y2 * ScreenY), Me.GetFadeColor(x)
            Next i
            
        Case Else
            ' The gradient is not found!  For simplicity's sake we
            ' will ignore this error and just do nothing.
    End Select
    


End Sub








Private Sub PaintObject2(frm As Object, Gradient As ColorFadeGradientConstants, Optional CycleCount As Long = 1)
    ' This is the gradient method for objects.  Use this to automate the
    ' drawing of gradients using VBs Line method.  This function does almost
    ' no error checking so make sure that you keep it private and wrap it
    ' with a public function instead.
    
    Dim ScreenX As Long     'Screen dimension
    Dim ScreenY As Long     'Screen dimension
    Dim h As Long           'Object Height
    Dim w As Long           'Object Width
    Dim y As Long           'Current Y-Position
    Dim x As Long           'Current X-Position
    Dim a As Long           'Current Angle Offset Multiplier
    Dim ax As Long          'Current Angle Offset
        
    Dim x1 As Long, x2 As Long  'Left and Right of the object
    Dim y1 As Long, y2 As Long  'Top and Bottom of the object
    
    Dim i As Long   'Loop Counter
    Dim j As Long   'Loop Counter
    Dim c As Long   'Cycles or other value
    
    
    ScreenX = Screen.TwipsPerPixelX 'Determine the horizontal screen dimension
    ScreenY = Screen.TwipsPerPixelY 'Determine the vertical screen dimension

    x1 = 0: x2 = (frm.Width / ScreenX)  'Get the pixel width of the object
    y1 = 0: y2 = (frm.Height / ScreenY) 'Get the pixel height of the object
              
              
    c = CycleCount
    If c < 1 Then c = 1     'We are capping the cycle count for the sake
    If c > 12 Then c = 12   'of a slower computers resources.
    c = c * 2               'Double it for the rounded effect.
    
    
    ' This long chunk of code is where the gradients are performed.  This is
    ' much more complex than the previous gradient.  What makes things easier
    ' is that diagonal gradients look the same from opposing corners.
    Select Case Gradient
        Case HorizontalGradient
            x1 = 0: x2 = frm.Width
            y1 = 1: y2 = Int(frm.Height / (ScreenY * 100)) + 1
    
            For j = y1 To y2
                For i = 1 To 100
                    frm.Line (x1, y)-(x2, y), Me.GetFadeColor2(i, c)
                    y = (((ScreenY * 100) * (j - 1)) + (i * ScreenY))
                Next i
            Next j
        
        
        Case VerticalGradient
            x1 = 1: x2 = Int(frm.Width / (ScreenX * 100)) + 1
            y1 = 0: y2 = frm.Height
    
            For j = x1 To x2
                For i = 1 To 100
                    frm.Line (x, y1)-(x, y2), Me.GetFadeColor2(i, c)
                    x = (((ScreenX * 100) * (j - 1)) + (i * ScreenX))
                Next i
            Next j
        
        Case DiagUpperRightGradient, DiagLowerLeftGradient
            x1 = 1: x2 = Int(frm.Width / (ScreenX * 100)) + 1
            y1 = 0: y2 = frm.Height
            
            ax = Int(((y2 / (ScreenY * 100)) + 1) / 2) + 1
           'ax = Int(((frm.Height / (ScreenY * 100)) + 1) / 2) 'Fixed
            a = (y2 / 2)
            
            For j = x1 - ax To x2 + ax
                For i = 1 To 100
                    frm.Line (x - a, y1)-(x + a, y2), Me.GetFadeColor2(i, c)
                    x = (((ScreenX * 100) * (j - 1)) + (i * ScreenX))
                Next i
            Next j
        
        Case DiagUpperLeftGradient, DiagLowerRightGradient
            x1 = 1: x2 = Int(frm.Width / (ScreenX * 100)) + 1
            y1 = 0: y2 = frm.Height
            
            ax = Int(((y2 / (ScreenY * 100)) + 1) / 2) + 1
            a = (y2 / 2)
            
            For j = x1 - ax To x2 + ax
                For i = 1 To 100
                    frm.Line (x + a, y1)-(x - a, y2), Me.GetFadeColor2(i, c)
                    x = (((ScreenX * 100) * (j - 1)) + (i * ScreenX))
                Next i
            Next j
    End Select
    
End Sub



















Public Property Let FadeEndColor(ByVal vData As Long)
    'Sets the fade end color and the individual RGB values
    mvarFadeEndColor = vData
    
    ' You MUST "call" this procedure.  If you just branch, there
    ' is a small chance that VB will not update the RGB values.
    Call DetermineRGB(vData, fcer, fceg, fceb)
End Property


Public Property Get FadeEndColor() As Long
    'Returns the fade end color
    FadeEndColor = mvarFadeEndColor
End Property



Public Property Let FadeStartColor(ByVal vData As Long)
    'Sets the fade end color and the individual RGB values
    mvarFadeStartColor = vData
    
    ' You MUST "call" this procedure.  If you just branch, there
    ' is a small chance that VB will not update the RGB values.
    Call DetermineRGB(vData, fcsr, fcsg, fcsb)
End Property


Public Property Get FadeStartColor() As Long
    'Returns the fade start color.
    FadeStartColor = mvarFadeStartColor
End Property








Private Sub DetermineRGB(varRGB As Long, varRed As Long, varGreen As Long, varBlue As Long)
    ' Determines the individual Red, Green, and Blue values from a long
    ' number.  Please note that there is probably a faster mathematical
    ' way of doing this, but I don't know it, and since this procedure is
    ' only called when you set the start or end color to something new, it
    ' really isn't important.
    '
    ' NOTICE:  You must "CALL" this procedure, not just branch.
    '
            
    Dim rs$ 'String to hold the hex of the color
    rs$ = Hex(varRGB)
    
    While Len(rs$) < 6
        ' We need a six character string in the BBGGRR so we can
        ' extract the individual colors.  0s are prepended for this.
        rs$ = "0" & rs$
    Wend
    
    varRed = Val("&H" & Mid(rs$, 5, 2))     'Duh
    varGreen = Val("&H" & Mid(rs$, 3, 2))   'Duh again
    varBlue = Val("&H" & Mid(rs$, 1, 2))    'Duh one more time.
  
End Sub




Private Function DetermineFade(R1 As Long, G1 As Long, B1 As Long, R2 As Long, G2 As Long, B2 As Long, Percent As Long) As Long
        ' Returns a color that sits somewhere between a start and end color.
        ' Basically its a three-part wrapper for the simple MidPoint function.
        '
        ' If you want to improve speed this is the place to do it.
        '

        Dim fcr As Long     'Red
        Dim fcg As Long     'Green
        Dim fcb As Long     'Blue
        
        Dim fcp As Long     'Percentage
        
        fcp = Percent       ' I know, I know.  ByVal it...
        If fcp > 100 Then fcp = 100
        If fcp < 0 Then fcp = 0
                        
        fcr = MidPoint(R1, R2, fcp)     'Interpolate
        fcg = MidPoint(G1, G2, fcp)     'Interpolate
        fcb = MidPoint(B1, B2, fcp)     'Interpolate
        
        DetermineFade = RGB(fcr, fcg, fcb)  'Return a long number.

End Function




Private Function MidPoint(x1 As Long, x2 As Long, p As Long) As Long
        ' This is a simple interpolation function.  It will find a midpoint
        ' number based on the percentage.  This function is the single most
        ' called function in the entire class, so keep it short and sweet.

        If x1 > x2 Then
            MidPoint = x1 - ((x1 - x2) * (p / 100))
        Else
            MidPoint = ((x2 - x1) * (p / 100)) + x1
        End If

End Function
