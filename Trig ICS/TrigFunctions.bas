Attribute VB_Name = "Module1"
Option Explicit

Global Const PI = 3.14159

Public Sub GetValues(ByRef A As Single, _
                     ByRef K As Single, _
                     ByRef C As Single, _
                     ByRef D As Single, _
                     ByRef XMin As Integer, _
                     ByRef XMax As Integer, _
                     ByRef YMin As Integer, _
                     ByRef YMax As Integer)
    
    Dim StA As String
    Dim StK As String
    
    Dim StC As String
    Dim StD As String
    
    Dim StXMin As String
    Dim StXMax As String
    
    Dim StYMin As String
    Dim StYMax As String
    
    StA = frmMain.txtVStretch.Text
    StK = frmMain.txtHStretch.Text
    
    StC = frmMain.txtVShift.Text
    StD = frmMain.txtHShift.Text
    
    StXMin = frmMain.txtXMin.Text
    StXMax = frmMain.txtXMax.Text
    
    StYMin = frmMain.txtYMin.Text
    StYMax = frmMain.txtYMax.Text
    
    'Checks if the user inputed anything
    'and if what they inputed was valid for
    
    'Vertical Stretch
    If StA = "" Then
        A = 1
    ElseIf IsNumeric(StA) And Left$(StA, 1) <> "$" Then
        A = Val(StA)
    Else
        MsgBox "Please enter valid numbers for each field.", vbExclamation + vbOKOnly, "Error: Invalid Values"
    End If
    
    'Horizontal Stretch
    If StK = "" Then
        K = 1
    ElseIf IsNumeric(StK) And Left$(StK, 1) <> "$" Then
        K = Val(StK)
    Else
        MsgBox "Please enter valid numbers for each field.", vbExclamation + vbOKOnly, "Error: Invalid Values"
    End If
    
    'Vertical Shift
    If StC = "" Then
        C = 0
    ElseIf IsNumeric(StC) And Left$(StC, 1) <> "$" Then
        C = Val(StC)
    Else
        MsgBox "Please enter valid numbers for each field.", vbExclamation + vbOKOnly, "Error: Invalid Values"
    End If
    
    'Horizontal Shift
    If StD = "" Then
        D = 0
    ElseIf IsNumeric(StD) And Left$(StD, 1) <> "$" Then
        D = Val(StD)
    Else
        MsgBox "Please enter valid numbers for each field.", vbExclamation + vbOKOnly, "Error: Invalid Values"
    End If
    
    'Domain
    If StXMin = "" Then
        XMin = 0
    ElseIf IsNumeric(StXMin) And Left$(StXMin, 1) <> "$" And InStr(1, StXMin, ".") = 0 Then
        XMin = Val(StXMin)
    Else
        MsgBox "Please enter valid numbers for each field.", vbExclamation + vbOKOnly, "Error: Invalid Values"
    End If
    If StXMax = "" Then
        XMax = 360
    ElseIf IsNumeric(StXMax) And Left$(StXMax, 1) <> "$" And InStr(1, StXMax, ".") = 0 Then
        XMax = Val(StXMax)
    Else
        MsgBox "Please enter valid numbers for each field.", vbExclamation + vbOKOnly, "Error: Invalid Values"
    End If
    
    'Range
    If StYMin = "" Then
        YMin = -1
    ElseIf IsNumeric(StYMin) And Left$(StYMin, 1) <> "$" And InStr(1, StYMin, ".") = 0 Then
        YMin = Val(StYMin)
    Else
        MsgBox "Please enter valid numbers for each field.", vbExclamation + vbOKOnly, "Error: Invalid Values"
    End If
    If StYMax = "" Then
        YMax = 1
    ElseIf IsNumeric(StYMax) And Left$(StYMax, 1) <> "$" And InStr(1, StYMax, ".") = 0 Then
        YMax = Val(StYMax)
    Else
        MsgBox "Please enter valid numbers for each field.", vbExclamation + vbOKOnly, "Error: Invalid Values"
    End If
    
End Sub
                     
Public Sub Graph(ByVal XAXIS As Integer, ByVal YAXIS As Integer, _
                 ByVal XSTRETCH As Single, ByVal YSTRETCH As Single, _
                 ByVal A As Single, _
                 ByVal K As Single, _
                 ByVal C As Single, _
                 ByVal D As Single, _
                 ByVal XMin As Integer, _
                 ByVal XMax As Integer, _
                 ByVal YMin As Integer, _
                 ByVal YMax As Integer)
            
    Dim Z As Single
    Dim XVal As Single
    Dim Radians As Single
    Dim YVal As Single
    
    Dim CRed As Integer
    Dim CGreen As Integer
    Dim CBlue As Integer
    
    'Changes colour depending on sin, cos, or tan function
    If frmMain.optSin.Value Then
        CRed = 255
        CGreen = 0
        CBlue = 0
        
    ElseIf frmMain.optCos.Value Then
        CRed = 0
        CGreen = 0
        CBlue = 255
    ElseIf frmMain.optTan.Value Then
        CRed = 0
        CGreen = 150
        CBlue = 150
    End If
        
    For Z = -6000 To 8000 Step 0.01
    
        XVal = Z
        Radians = (Z / 180) * PI
        
        If frmMain.optSin.Value Then
            YVal = Sin(Radians)
        ElseIf frmMain.optCos.Value Then
            YVal = Cos(Radians)
        ElseIf frmMain.optTan.Value Then
            YVal = Tan(Radians)
        End If
        
        'Updates the X and Y values given the appropriate mapping
        '(X, Y) --> (X / K + D, Y * A + C)
        XVal = (XVal * XSTRETCH) / K + (D * XSTRETCH)
        YVal = (YVal * YSTRETCH) * A + (C * YSTRETCH)
        
        'Checks to see if the new x and y values fit within the domain/range
        If (YVal >= (YMin * YSTRETCH) And YVal <= (YMax * YSTRETCH)) Then
            If (XVal >= (XMin * XSTRETCH) And XVal <= (XMax * XSTRETCH)) Then
                frmMain.PSet (XVal + YAXIS, XAXIS - YVal), RGB(CRed, CGreen, CBlue)
            End If
        End If
    Next Z
    
End Sub

Public Sub DisplayValues(ByVal A As Single, _
                         ByVal K As Single, _
                         ByVal C As Single, _
                         ByVal D As Single, _
                         ByVal XMin As Integer, _
                         ByVal XMax As Integer, _
                         ByVal YMin As Integer, _
                         ByVal YMax As Integer)
    
    With frmMain
        .txtVShift = Str$(A)
        .txtHShift = Str$(K)
        
        .txtVShift = Str$(C)
        .txtHShift = Str$(D)
        
        .txtXMin = Str$(XMin)
        .txtXMax = Str$(XMax)
        
        .txtYMin = Str$(YMin)
        .txtYMax = Str$(YMax)
    End With

End Sub

Public Sub Reset()
    
    With frmMain
        'Initially sets values of option buttons
        'and textboxes
        .optSin.Value = True
        
        .optXLessEqual1.Value = True
        .optXLessEqual2.Value = True
        
        .optYLessEqual1.Value = True
        .optYLessEqual2.Value = True
        
        .txtXMin.Text = "0"
        .txtXMax.Text = "360"
        
        .txtYMin.Text = "-1"
        .txtYMax.Text = "1"
        
        .txtVStretch.Text = "1"
        .txtHStretch.Text = "1"
        
        .txtVShift.Text = "0"
        .txtHShift.Text = "0"
    End With
    
End Sub
