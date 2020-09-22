Attribute VB_Name = "modStartup"

Public Fade As New ColorFade

Public FadeDirection As ColorFadeGradientConstants
Public FadeCycles As Long


Sub main()

    Fade.FadeStartColor = RGB(0, 0, 0)
    Fade.FadeEndColor = RGB(255, 255, 255)
        
    frmMain.Show

End Sub
