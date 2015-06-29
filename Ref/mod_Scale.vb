Module mod_Scale

    Private arrNote(7) As String
    Private arrNum(7) As Integer

    Private Sub iniArr()
        For index = 0 To 6
            arrNum(index) = 0
            arrNote(index) = ""
        Next
    End Sub

    Sub lblShowScale()
        frm_Main.lbl_Scale1.Text = arrNote(0)
        frm_Main.lbl_Scale2.Text = arrNote(1)
        frm_Main.lbl_Scale3.Text = arrNote(2)
        frm_Main.lbl_Scale4.Text = arrNote(3)
        frm_Main.lbl_Scale5.Text = arrNote(4)
        frm_Main.lbl_Scale6.Text = arrNote(5)
        frm_Main.lbl_Scale7.Text = arrNote(6)
        'frm_Main.lbl_Eh.Text = frm_Main.cBx_Eh.Text
    End Sub

    'Execute all scales (C,D,E.....)
    Sub exeScale()
        'Major scales
        If frm_Main.cobx_ScaleX.SelectedIndex = 0 Then
            getArrayNumMaj(1, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            CmajScale()
        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 1 Then
            getArrayNumMaj(2, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            DbmajScale()
        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 2 Then
            getArrayNumMaj(3, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            DmajScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 3 Then
            getArrayNumMaj(4, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            EbmajScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 4 Then
            getArrayNumMaj(5, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            EmajScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 5 Then
            getArrayNumMaj(6, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            FmajScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 6 Then
            ' getArrayNumMaj(7, arrNum)
            ' getArrayNotes(arrNum, arrNote, True)
            arrNote = {"F#", "G#", "A#", "B", "C#", "D#", "E#"}
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            FsmajScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 7 Then
            getArrayNumMaj(8, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            GmajScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 8 Then
            getArrayNumMaj(9, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            AbmajScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 9 Then
            getArrayNumMaj(10, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            AmajScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 10 Then
            getArrayNumMaj(11, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            BbmajScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 11 Then
            getArrayNumMaj(12, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "Majeur"
            lblShowScale()
            setAllReset()
            BmajScale()

            'a exception
        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 12 Then
            'getArrayNumMaj(0, arrNum)
            'getArrayNotes(arrNum, arrNote, True)
            arrNote = {"", "", "", "", "", "", ""}
            frm_Main.lbl_Eh.Text = ""
            lblShowScale()
            setAllReset()

        End If

        'Minor scales
        If frm_Main.cobx_ScaleX.SelectedIndex = 13 Then
            getArrayNumMin(1, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            CminScale()
        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 14 Then
            getArrayNumMin(2, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            CsminScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 15 Then
            getArrayNumMin(3, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            DminScale()

            'a exception
        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 16 Then
            'getArrayNumMin(4, arrNum)
            'getArrayNotes(arrNum, arrNote, True)
            arrNote = {"D#", "E#", "F#", "G#", "A#", "B", "C#"}
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            DsminScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 17 Then
            getArrayNumMin(5, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            EminScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 18 Then
            getArrayNumMin(6, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            FminScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 19 Then
            getArrayNumMin(7, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            FsminScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 20 Then
            getArrayNumMin(8, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            GminScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 21 Then
            getArrayNumMin(9, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            GsminScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 22 Then
            getArrayNumMin(10, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            AminScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 23 Then
            getArrayNumMin(11, arrNum)
            getArrayNotes(arrNum, arrNote, False)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            BbminScale()

        ElseIf frm_Main.cobx_ScaleX.SelectedIndex = 24 Then
            getArrayNumMin(12, arrNum)
            getArrayNotes(arrNum, arrNote, True)
            frm_Main.lbl_Eh.Text = arrNote(0) & " " & "mineur"
            lblShowScale()
            setAllReset()
            BminScale()
        End If

    End Sub
End Module
