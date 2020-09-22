Attribute VB_Name = "TitleAnimations"
Public Sub LoadAnimation()
Select Case pTitleType
Case 1
    frmMain.tmrAnimateTitle.Interval = 100
    frmMain.tmrAnimateTitle.Enabled = True
    C = frmMain.Caption
    CO = Len(C) + 1
    frmMain.Caption = ""

    If frmMain.BorderStyle <> 2 Then
        FS = frmMain.ScaleWidth + 250
    Else
        FS = frmMain.ScaleWidth + 500
    End If
Case 2


End Select
End Sub


Public Sub ResizeWindow()
    If frmMain.WindowState = 1 Then
        FS = 3500
    Else
        FS = frmMain.ScaleWidth
    End If
End Sub

Public Sub SLIDER()
    On Error GoTo ATH
    Static C01 As Integer ' Counter 1
    Static CO2 As Integer ' Counter 2
    Static A As String 'To move Caption
    Dim R As String 'Restore Caption
    Dim T As String 'Restore Caption
    
XX:


    If CO > 0 Then
        C01 = CO
        T = Mid(C, C01, 1)
        CO = CO - 1
        R = " "
        Mid(R, 1) = T
        frmMain.Caption = R & frmMain.Caption
    Else
        A = A & " "
        R = " "
        Mid(R, 1) = A
        frmMain.Caption = R & frmMain.Caption
    End If


    If CO2 >= FS Then
        CO2 = 0
        CO = Len(C)
        frmMain.Caption = ""
        GoTo XX
    Else
        CO2 = CO2 + 50
    End If
    Exit Sub
ATH:
End Sub
