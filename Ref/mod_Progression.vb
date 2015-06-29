Module mod_Progression
    Private arrNote(7) As String
    Private arrNum(7) As Integer

    Private Sub iniArr()
        For index = 0 To 6
            arrNum(index) = 0
            arrNote(index) = ""
        Next
    End Sub

    'Show all label of process major (Cm.....)
    Sub lblShowProgressionMaj()
        frm_Main.lbl_Progression1.Text = arrNote(0)
        frm_Main.lbl_Progression2.Text = arrNote(1) & "m"
        frm_Main.lbl_Progression3.Text = arrNote(2) & "m"
        frm_Main.lbl_Progression4.Text = arrNote(3)
        frm_Main.lbl_Progression5.Text = arrNote(4)
        frm_Main.lbl_Progression6.Text = arrNote(5) & "m"
        frm_Main.lbl_Progression7.Text = arrNote(6) & "◦"
        'frm_Main.lbl_Dg.Text = frm_Main.cBx_Dg.Text
    End Sub

    'Show all label of process minor (Cm.....)
    Sub lblShowProgressionMin()
        frm_Main.lbl_Progression1.Text = arrNote(0) & "m"
        frm_Main.lbl_Progression2.Text = arrNote(1) & "◦"
        frm_Main.lbl_Progression3.Text = arrNote(2)
        frm_Main.lbl_Progression4.Text = arrNote(3) & "m"
        frm_Main.lbl_Progression5.Text = arrNote(4) & "m"
        frm_Main.lbl_Progression6.Text = arrNote(5)
        frm_Main.lbl_Progression7.Text = arrNote(6)
    End Sub

    'Execute all process (C,D,E.....)
    Sub exeProgression()
        'Major 
        If frm_Main.cobx_ProgressionX.SelectedIndex = 0 Then
            getArrayNumMaj(1, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            lblShowProgressionMaj()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 1 Then
            getArrayNumMaj(2, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            lblShowProgressionMaj()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 2 Then
            getArrayNumMaj(3, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            lblShowProgressionMaj()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 3 Then
            getArrayNumMaj(4, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            lblShowProgressionMaj()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 4 Then
            getArrayNumMaj(5, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            lblShowProgressionMaj()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 5 Then
            getArrayNumMaj(6, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            lblShowProgressionMaj()

            'a exception
        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 6 Then
            ' getArrayNumMaj(7, arrNum)
            ' getArrayNotes(arrNum, arrNote, True)
            arrNote = {"F#", "G#", "A#", "B", "C#", "D#", "E#"}
            lblShowProgressionMaj()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 7 Then
            getArrayNumMaj(8, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            lblShowProgressionMaj()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 8 Then
            getArrayNumMaj(9, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            lblShowProgressionMaj()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 9 Then
            getArrayNumMaj(10, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            lblShowProgressionMaj()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 10 Then
            getArrayNumMaj(11, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            lblShowProgressionMaj()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 11 Then
            getArrayNumMaj(12, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            lblShowProgressionMaj()

            'a exception
        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 12 Then
            arrNote = {"", "", "", "", "", "", ""}
            frm_Main.lbl_Dg.Text = ""
            'lblShowProgressionMaj()
        End If

        'Minor
        If frm_Main.cobx_ProgressionX.SelectedIndex = 13 Then
            getArrayNumMin(1, arrNum)
            getArrayNotes(arrNum, arrNote, False)

            lblShowProgressionMin()


        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 14 Then
            getArrayNumMin(2, arrNum)
            getArrayNotes(arrNum, arrNote, True)

            lblShowProgressionMin()


        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 15 Then
            getArrayNumMin(3, arrNum)
            getArrayNotes(arrNum, arrNote, False)

            lblShowProgressionMin()

            'a exception
        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 16 Then
            'getArrayNumMin(4, arrNum)
            'getArrayNotes(arrNum, arrNote, True)
            arrNote = {"D#", "E#", "F#", "G#", "A#", "B", "C#"}

            lblShowProgressionMin()


        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 17 Then
            getArrayNumMin(5, arrNum)
            getArrayNotes(arrNum, arrNote, True)

            lblShowProgressionMin()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 18 Then
            getArrayNumMin(6, arrNum)
            getArrayNotes(arrNum, arrNote, False)

            lblShowProgressionMin()


        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 19 Then
            getArrayNumMin(7, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            lblShowProgressionMin()


        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 20 Then
            getArrayNumMin(8, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            lblShowProgressionMin()


        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 21 Then
            getArrayNumMin(9, arrNum)
            getArrayNotes(arrNum, arrNote, True)

            lblShowProgressionMin()


        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 22 Then
            getArrayNumMin(10, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            lblShowProgressionMin()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 23 Then
            getArrayNumMin(11, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            lblShowProgressionMin()

        ElseIf frm_Main.cobx_ProgressionX.SelectedIndex = 24 Then
            getArrayNumMin(12, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            lblShowProgressionMin()

        End If


    End Sub
End Module
