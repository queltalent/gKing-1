Module mod_Chord
    Private arrNoteChord(4) As String
    Private arrNumChord(4) As Integer

    Private Sub iniArr()
        For index = 0 To 3
            arrNumChord(index) = 0
            arrNoteChord(index) = ""
        Next
    End Sub

    'Show all label of Chord3 (C,D,E.....)
    Sub lblShowChord()
        frm_Main.lbl_Chord1.Text = arrNoteChord(0)
        frm_Main.lbl_Chord3.Text = arrNoteChord(1)
        frm_Main.lbl_Chord5.Text = arrNoteChord(2)
    End Sub

    'Show all label Chord7 (C7,D7,E7.....)
    Sub lblShowChord7()
        frm_Main.lbl_Chord1.Text = arrNoteChord(0)
        frm_Main.lbl_Chord3.Text = arrNoteChord(1)
        frm_Main.lbl_Chord5.Text = arrNoteChord(2)
        frm_Main.lbl_Chord7.Text = arrNoteChord(3)
    End Sub

    'Execute all Chord (C,D,E.....)
    Sub exeChord()
        '********************** Major Chord *******************************
        If frm_Main.cobx_ChordX.SelectedIndex = 0 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(1, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 1 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(2, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 2 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(3, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 3 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(4, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 4 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(5, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 5 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(6, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 6 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(7, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 7 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(8, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 8 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(9, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 9 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(10, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 10 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(11, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 11 And frm_Main.cobx_ChordY.SelectedIndex = 0 Then
            getArrNumChordMaj(12, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord()
        End If

        'Minor Chord
        If frm_Main.cobx_ChordX.SelectedIndex = 0 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            getArrNumChordMin(1, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

            'a exception
        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 1 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            'getArrNumChordMin(2, arrNumChord)
            'getArrNoteChord(arrNumChord, arrNoteChord, False)
            arrNoteChord = {"Db", "Fb", "Ab", ""}
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 2 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            getArrNumChordMin(3, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 3 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            getArrNumChordMin(4, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 4 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            getArrNumChordMin(5, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 5 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            getArrNumChordMin(6, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 6 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            getArrNumChordMin(7, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 7 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            getArrNumChordMin(8, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

            'a exception
        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 8 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            'getArrNumChordMin(9, arrNumChord)
            'getArrNoteChord(arrNumChord, arrNoteChord, False)
            arrNoteChord = {"Ab", "Cb", "Eb", ""}
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 9 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            getArrNumChordMin(10, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 10 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            getArrNumChordMin(11, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 11 And frm_Main.cobx_ChordY.SelectedIndex = 1 Then
            getArrNumChordMin(12, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord()
        End If


        'Major 7 Chord
        If frm_Main.cobx_ChordX.SelectedIndex = 0 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(1, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 1 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(2, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 2 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(3, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 3 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(4, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 4 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(5, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 5 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(6, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 6 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(7, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 7 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(8, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 8 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(9, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 9 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(10, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 10 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(11, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 11 And frm_Main.cobx_ChordY.SelectedIndex = 3 Then
            getArrNumChordMaj(12, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()
        End If

        'Minor 7 Chord
        If frm_Main.cobx_ChordX.SelectedIndex = 0 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            getArrNumChordMin(1, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

            'a exception
        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 1 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            'getArrNumChordMin(2, arrNumChord)
            'getArrNoteChord(arrNumChord, arrNoteChord, False)
            arrNoteChord = {"Db", "Fb", "Ab", ""}
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 2 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            getArrNumChordMin(3, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 3 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            getArrNumChordMin(4, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 4 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            getArrNumChordMin(5, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 5 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            getArrNumChordMin(6, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 6 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            getArrNumChordMin(7, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 7 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            getArrNumChordMin(8, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

            'a exception
        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 8 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            'getArrNumChordMin(9, arrNumChord)
            'getArrNoteChord(arrNumChord, arrNoteChord, False)
            arrNoteChord = {"Ab", "Cb", "Eb", ""}
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 9 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            getArrNumChordMin(10, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 10 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            getArrNumChordMin(11, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 11 And frm_Main.cobx_ChordY.SelectedIndex = 4 Then
            getArrNumChordMin(12, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()
        End If



        'Diminished 7 Chord
        If frm_Main.cobx_ChordX.SelectedIndex = 0 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            getArrNumChordDim(1, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

            'a exception
        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 1 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            getArrNumChordDim(2, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            'arrNoteChord = {"Db", "Fb", "Ab", ""}
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 2 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            getArrNumChordDim(3, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 3 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            getArrNumChordDim(4, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 4 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            getArrNumChordDim(5, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 5 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            getArrNumChordDim(6, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 6 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            getArrNumChordDim(7, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 7 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            getArrNumChordDim(8, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

            'a exception
        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 8 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            'getArrNumChordDim(9, arrNumChord)
            'getArrNoteChord(arrNumChord, arrNoteChord, False)
            'arrNoteChord = {"Ab", "Cb", "Eb"}
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 9 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            getArrNumChordDim(10, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 10 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            getArrNumChordDim(11, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 11 And frm_Main.cobx_ChordY.SelectedIndex = 5 Then
            getArrNumChordDim(12, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()
        End If


        'Dominant 7 Chord
        If frm_Main.cobx_ChordX.SelectedIndex = 0 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            getArrNumChordDom(1, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

            'a exception
        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 1 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            getArrNumChordDom(2, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            'arrNoteChord = {"Db", "Fb", "Ab", ""}
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 2 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            getArrNumChordDom(3, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 3 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            getArrNumChordDom(4, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 4 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            getArrNumChordDom(5, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 5 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            getArrNumChordDom(6, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 6 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            getArrNumChordDom(7, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, True)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 7 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            getArrNumChordDom(8, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

            'a exception
        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 8 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            'getArrNumChordDom(9, arrNumChord)
            'getArrNoteChord(arrNumChord, arrNoteChord, False)
            'arrNoteChord = {"Ab", "Cb", "Eb"}
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 9 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            getArrNumChordDom(10, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 10 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            getArrNumChordDom(11, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()

        ElseIf frm_Main.cobx_ChordX.SelectedIndex = 11 And frm_Main.cobx_ChordY.SelectedIndex = 6 Then
            getArrNumChordDom(12, arrNumChord)
            getArrNoteChord(arrNumChord, arrNoteChord, False)
            lblShowChord7()
        End If
    End Sub
End Module
