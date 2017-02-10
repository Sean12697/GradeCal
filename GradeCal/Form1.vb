Public Class Form1
    Dim SayGrades As String

    Private Sub TextToSpeach(ByRef Words As String)
        Dim SAPI
        SAPI = CreateObject("sapi.spvoice")
        SAPI.Speak(Words)
    End Sub

    Function Year1PointsToGrade(ByRef Number As Integer)
        Dim Grade As String
        Select Case Number
            Case 0 To 659
                Grade = "Fail"
            Case 660 To 689
                Grade = "MP"
            Case 690 To 719
                Grade = "MM"
            Case 720 To 749
                Grade = "DM"
            Case 750 To 769
                Grade = "DD"
            Case 770 To 789
                Grade = "D*D"
            Case 790 To 1299
                Grade = "D*D*"
            Case Else
                Grade = "Error"
        End Select
        Return Grade
    End Function

    Function Year2PointsToGrade(ByRef Number As Integer)
        Dim Grade As String
        Select Case Number
            Case 0 To 1299
                Grade = "Fail"
                UCAS = 0
            Case 1300 To 1339
                Grade = "MPP"
                UCAS = 160
            Case 1340 To 1379
                Grade = "MMP"
                UCAS = 200
            Case 1380 To 1419
                Grade = "MMM"
                UCAS = 240
            Case 1420 To 1459
                Grade = "DMM"
                UCAS = 280
            Case 1460 To 1499
                Grade = "DDM"
                UCAS = 320
            Case 1500 To 1529
                Grade = "DDD"
                UCAS = 360
            Case 1530 To 1559
                Grade = "D*DD"
                UCAS = 400
            Case 1560 To 1589
                Grade = "D*D*D"
                UCAS = 440
            Case 1590 To 9999
                Grade = "D*D*D*"
                UCAS = 480
            Case Else
                Grade = "Error"
                UCAS = 0
        End Select
        Return Grade
    End Function

    Private Sub UpdateThings()
        Dim Year1Grades = Year1PointsToGrade(Year1points)
        Dim Year2Grades = Year2PointsToGrade(Year2points)
        Dim Year1GradesName As String
        Dim Year2GradesName As String
        Select Case Year1Grades
            Case "MM"
                Year1GradesName = "Double Merit"
            Case "DD"
                Year1GradesName = "Double Distiction"
            Case "D*D*"
                Year1GradesName = "Double Distiction Star"
            Case Else
                Year1GradesName = Year1Grades
        End Select
        Select Case Year2Grades
            Case "MMM"
                Year2GradesName = "Triple Merit"
            Case "DDD"
                Year2GradesName = "Triple Distiction"
            Case "D*D*D*"
                Year2GradesName = "Triple Distiction Star"
            Case Else
                Year2GradesName = Year2Grades
        End Select
        Year1points = U1points + U2points + U3points + U6points + U25points + U29points + U31points + U40points + U42points
        Year2points = Year1points + U11points + U14points + U17points + U18points + U22points + U23points + U28points + U30points + U37points
        lblYear1.Text = "Year 1: " & Year1Grades & " (" & Year1points & " Points)"
        lblYear2.Text = "Year 2: " & Year2Grades & " (" & Year2points & " Points)"
        lblUCAS.Text = "UCAS: " & UCAS
        If Year1Grades = "Fail" And Year2Grades = "Fail" Then
            SayGrades = "You are currently failing the course with " & Year2points & " points."
        End If
        If Year2Grades = "Fail" = False And Year2Grades = "Error" = False Then
            SayGrades = "You have " & Year2GradesName & " with " & Year2points & " points and a UCAS score of " & UCAS
        End If
        If Year1Grades = "Fail" = False And Year1Grades = "Error" = False And Year2Grades = "Fail" Then
            SayGrades = "You have " & Year2GradesName & " with " & Year2points & " points and a UCAS score of " & UCAS
        End If
    End Sub

    Private Sub year1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbU1.SelectedIndexChanged, cmbU2.SelectedIndexChanged, cmbU3.SelectedIndexChanged, cmbU6.SelectedIndexChanged, cmbU25.SelectedIndexChanged, cmbU29.SelectedIndexChanged, cmbU31.SelectedIndexChanged, cmbU40.SelectedIndexChanged, cmbU42.SelectedIndexChanged
        Select Case DirectCast(sender, ComboBox).Name
            Case cmbU1.Name
                Select Case cmbU1.Text
                    Case "P"
                        U1points = 70
                    Case "M"
                        U1points = 80
                    Case "D"
                        U1points = 90
                    Case Else
                        U1points = 0
                End Select
            Case cmbU2.Name
                Select Case cmbU2.Text
                    Case "P"
                        U2points = 70
                    Case "M"
                        U2points = 80
                    Case "D"
                        U2points = 90
                    Case Else
                        U2points = 0
                End Select
            Case cmbU3.Name
                Select Case cmbU3.Text
                    Case "P"
                        U3points = 70
                    Case "M"
                        U3points = 80
                    Case "D"
                        U3points = 90
                    Case Else
                        U3points = 0
                End Select
            Case cmbU6.Name
                Select Case cmbU6.Text
                    Case "P"
                        U6points = 70
                    Case "M"
                        U6points = 80
                    Case "D"
                        U6points = 90
                    Case Else
                        U6points = 0
                End Select
            Case cmbU25.Name
                Select Case cmbU25.Text
                    Case "P"
                        U25points = 70
                    Case "M"
                        U25points = 80
                    Case "D"
                        U25points = 90
                    Case Else
                        U25points = 0
                End Select
            Case cmbU29.Name
                Select Case cmbU29.Text
                    Case "P"
                        U29points = 70
                    Case "M"
                        U29points = 80
                    Case "D"
                        U29points = 90
                    Case Else
                        U29points = 0
                End Select
            Case cmbU31.Name
                Select Case cmbU31.Text
                    Case "P"
                        U31points = 70
                    Case "M"
                        U31points = 80
                    Case "D"
                        U31points = 90
                    Case Else
                        U31points = 0
                End Select
            Case cmbU40.Name
                Select Case cmbU40.Text
                    Case "P"
                        U40points = 70
                    Case "M"
                        U40points = 80
                    Case "D"
                        U40points = 90
                    Case Else
                        U40points = 0
                End Select
            Case cmbU42.Name
                Select Case cmbU42.Text
                    Case "P"
                        U42points = 70
                    Case "M"
                        U42points = 80
                    Case "D"
                        U42points = 90
                    Case Else
                        U42points = 0
                End Select
        End Select
        UpdateThings()
    End Sub

    Private Sub year2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbU37.SelectedIndexChanged, cmbU30.SelectedIndexChanged, cmbU28.SelectedIndexChanged, cmbU11.SelectedIndexChanged, cmbU22.SelectedIndexChanged, cmbU18.SelectedIndexChanged, cmbU17.SelectedIndexChanged, cmbU14.SelectedIndexChanged, cmbU23.SelectedIndexChanged
        Select Case DirectCast(sender, ComboBox).Name
            Case cmbU11.Name
                Select Case cmbU11.Text
                    Case "P"
                        U11points = 70
                    Case "M"
                        U11points = 80
                    Case "D"
                        U11points = 90
                    Case Else
                        U11points = 0
                End Select
            Case cmbU14.Name
                Select Case cmbU14.Text
                    Case "P"
                        U14points = 70
                    Case "M"
                        U14points = 80
                    Case "D"
                        U14points = 90
                    Case Else
                        U14points = 0
                End Select
            Case cmbU17.Name
                Select Case cmbU17.Text
                    Case "P"
                        U17points = 70
                    Case "M"
                        U17points = 80
                    Case "D"
                        U17points = 90
                    Case Else
                        U17points = 0
                End Select
            Case cmbU18.Name
                Select Case cmbU18.Text
                    Case "P"
                        U18points = 70
                    Case "M"
                        U18points = 80
                    Case "D"
                        U18points = 90
                    Case Else
                        U18points = 0
                End Select
            Case cmbU22.Name
                Select Case cmbU22.Text
                    Case "P"
                        U22points = 70
                    Case "M"
                        U22points = 80
                    Case "D"
                        U22points = 90
                    Case Else
                        U22points = 0
                End Select
            Case cmbU23.Name
                Select Case cmbU23.Text
                    Case "P"
                        U23points = 70
                    Case "M"
                        U23points = 80
                    Case "D"
                        U23points = 90
                    Case Else
                        U23points = 0
                End Select
            Case cmbU28.Name
                Select Case cmbU28.Text
                    Case "P"
                        U28points = 70
                    Case "M"
                        U28points = 80
                    Case "D"
                        U28points = 90
                    Case Else
                        U28points = 0
                End Select
            Case cmbU30.Name
                Select Case cmbU30.Text
                    Case "P"
                        U30points = 70
                    Case "M"
                        U30points = 80
                    Case "D"
                        U30points = 90
                    Case Else
                        U30points = 0
                End Select
            Case cmbU37.Name
                Select Case cmbU37.Text
                    Case "P"
                        U37points = 70
                    Case "M"
                        U37points = 80
                    Case "D"
                        U37points = 90
                    Case Else
                        U37points = 0
                End Select
        End Select
        UpdateThings()
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        SetComboBoxes("-")
    End Sub

    Private Sub btnPass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPass.Click
        SetComboBoxes("P")
    End Sub

    Private Sub btnMerit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMerit.Click
        SetComboBoxes("M")
    End Sub

    Private Sub btnDist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDist.Click
        SetComboBoxes("D")
    End Sub

    Private Sub SetComboBoxes(ByRef Text As String)
        cmbU1.Text = Text
        cmbU11.Text = Text
        cmbU14.Text = Text
        cmbU17.Text = Text
        cmbU18.Text = Text
        cmbU2.Text = Text
        cmbU22.Text = Text
        cmbU23.Text = Text
        cmbU25.Text = Text
        cmbU28.Text = Text
        cmbU29.Text = Text
        cmbU3.Text = Text
        cmbU30.Text = Text
        cmbU31.Text = Text
        cmbU37.Text = Text
        cmbU40.Text = Text
        cmbU42.Text = Text
        cmbU6.Text = Text

        UpdateThings()
    End Sub

    Private Sub btnSpeak_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSpeak.Click
        TextToSpeach(SayGrades)
    End Sub
End Class
