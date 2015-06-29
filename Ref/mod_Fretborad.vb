Module mod_Fretborad

#Region "All Notes Sets here "
    'Show every notes you need on fretboard, dim all sets of note first.
    Dim setC As New List(Of Label)
    Dim setCs As New List(Of Label)
    Dim setDb As New List(Of Label)
    Dim setD As New List(Of Label)
    Dim setDs As New List(Of Label)
    Dim setEb As New List(Of Label)
    Dim setE As New List(Of Label)
    Dim setF As New List(Of Label)
    Dim setFs As New List(Of Label)
    Dim setGb As New List(Of Label)
    Dim setG As New List(Of Label)
    Dim setGs As New List(Of Label)
    Dim setAb As New List(Of Label)
    Dim setA As New List(Of Label)
    Dim setAss As New List(Of Label)
    Dim setBb As New List(Of Label)
    Dim setB As New List(Of Label)

    Sub setCFretboard()
        setC.Add(frm_Main.lbl_1_2)
        setC.Add(frm_Main.lbl_3_5)
        setC.Add(frm_Main.lbl_5_3)
        setC.Add(frm_Main.lbl_8_1)
        setC.Add(frm_Main.lbl_8_6)
        setC.Add(frm_Main.lbl_10_4)
        setC.Add(frm_Main.lbl_13_2)
        setC.Add(frm_Main.lbl_15_5)
    End Sub

    Sub setCsFretboard()
        setCs.Add(frm_Main.lbl_2_2)
        setCs.Add(frm_Main.lbl_4_5)
        setCs.Add(frm_Main.lbl_6_3)
        setCs.Add(frm_Main.lbl_9_1)
        setCs.Add(frm_Main.lbl_9_6)
        setCs.Add(frm_Main.lbl_11_4)
        setCs.Add(frm_Main.lbl_14_2)
    End Sub

    Sub setDbFretboard()
        setDb.Add(frm_Main.lbl_2_2)
        setDb.Add(frm_Main.lbl_4_5)
        setDb.Add(frm_Main.lbl_6_3)
        setDb.Add(frm_Main.lbl_9_1)
        setDb.Add(frm_Main.lbl_9_6)
        setDb.Add(frm_Main.lbl_11_4)
        setDb.Add(frm_Main.lbl_14_2)
    End Sub

    Sub setDFretboard()
        setD.Add(frm_Main.lbl_0_4)
        setD.Add(frm_Main.lbl_3_2)
        setD.Add(frm_Main.lbl_5_5)
        setD.Add(frm_Main.lbl_7_3)
        setD.Add(frm_Main.lbl_10_1)
        setD.Add(frm_Main.lbl_10_6)
        setD.Add(frm_Main.lbl_12_4)
        setD.Add(frm_Main.lbl_15_2)
    End Sub

    Sub setDsFretboard()
        setDs.Add(frm_Main.lbl_1_4)
        setDs.Add(frm_Main.lbl_4_2)
        setDs.Add(frm_Main.lbl_6_5)
        setDs.Add(frm_Main.lbl_8_3)
        setDs.Add(frm_Main.lbl_11_1)
        setDs.Add(frm_Main.lbl_11_6)
        setDs.Add(frm_Main.lbl_13_4)
    End Sub

    Sub setEbFretboard()
        setEb.Add(frm_Main.lbl_1_4)
        setEb.Add(frm_Main.lbl_4_2)
        setEb.Add(frm_Main.lbl_6_5)
        setEb.Add(frm_Main.lbl_8_3)
        setEb.Add(frm_Main.lbl_11_1)
        setEb.Add(frm_Main.lbl_11_6)
        setEb.Add(frm_Main.lbl_13_4)
    End Sub

    Sub setEFretboard()
        setE.Add(frm_Main.lbl_0_1)
        setE.Add(frm_Main.lbl_0_6)
        setE.Add(frm_Main.lbl_2_4)
        setE.Add(frm_Main.lbl_5_2)
        setE.Add(frm_Main.lbl_7_5)
        setE.Add(frm_Main.lbl_9_3)
        setE.Add(frm_Main.lbl_12_1)
        setE.Add(frm_Main.lbl_12_6)
        setE.Add(frm_Main.lbl_14_4)
    End Sub

    Sub setFFretboard()
        setF.Add(frm_Main.lbl_1_1)
        setF.Add(frm_Main.lbl_1_6)
        setF.Add(frm_Main.lbl_3_4)
        setF.Add(frm_Main.lbl_6_2)
        setF.Add(frm_Main.lbl_8_5)
        setF.Add(frm_Main.lbl_10_3)
        setF.Add(frm_Main.lbl_13_1)
        setF.Add(frm_Main.lbl_13_6)
        setF.Add(frm_Main.lbl_15_4)
    End Sub

    Sub setFsFretboard()
        setFs.Add(frm_Main.lbl_2_1)
        setFs.Add(frm_Main.lbl_2_6)
        setFs.Add(frm_Main.lbl_4_4)
        setFs.Add(frm_Main.lbl_7_2)
        setFs.Add(frm_Main.lbl_9_5)
        setFs.Add(frm_Main.lbl_11_3)
        setFs.Add(frm_Main.lbl_14_1)
        setFs.Add(frm_Main.lbl_14_6)
    End Sub

    Sub setGbFretboard()
        setGb.Add(frm_Main.lbl_2_1)
        setGb.Add(frm_Main.lbl_2_6)
        setGb.Add(frm_Main.lbl_4_4)
        setGb.Add(frm_Main.lbl_7_2)
        setGb.Add(frm_Main.lbl_9_5)
        setGb.Add(frm_Main.lbl_11_3)
        setGb.Add(frm_Main.lbl_14_1)
        setGb.Add(frm_Main.lbl_14_6)
    End Sub

    Sub setGFretboard()
        setG.Add(frm_Main.lbl_0_3)
        setG.Add(frm_Main.lbl_3_1)
        setG.Add(frm_Main.lbl_3_6)
        setG.Add(frm_Main.lbl_5_4)
        setG.Add(frm_Main.lbl_8_2)
        setG.Add(frm_Main.lbl_10_5)
        setG.Add(frm_Main.lbl_12_3)
        setG.Add(frm_Main.lbl_15_1)
        setG.Add(frm_Main.lbl_15_6)
    End Sub

    Sub setGsFretboard()
        setGs.Add(frm_Main.lbl_1_3)
        setGs.Add(frm_Main.lbl_4_1)
        setGs.Add(frm_Main.lbl_4_6)
        setGs.Add(frm_Main.lbl_6_4)
        setGs.Add(frm_Main.lbl_9_2)
        setGs.Add(frm_Main.lbl_11_5)
        setGs.Add(frm_Main.lbl_13_3)
    End Sub

    Sub setAbFretboard()
        setAb.Add(frm_Main.lbl_1_3)
        setAb.Add(frm_Main.lbl_4_1)
        setAb.Add(frm_Main.lbl_4_6)
        setAb.Add(frm_Main.lbl_6_4)
        setAb.Add(frm_Main.lbl_9_2)
        setAb.Add(frm_Main.lbl_11_5)
        setAb.Add(frm_Main.lbl_13_3)

    End Sub

    Sub setAFretboard()
        setA.Add(frm_Main.lbl_0_5)
        setA.Add(frm_Main.lbl_2_3)
        setA.Add(frm_Main.lbl_5_1)
        setA.Add(frm_Main.lbl_5_6)
        setA.Add(frm_Main.lbl_7_4)
        setA.Add(frm_Main.lbl_10_2)
        setA.Add(frm_Main.lbl_12_5)
        setA.Add(frm_Main.lbl_14_3)
    End Sub

    Sub setAssFretboard()
        setAss.Add(frm_Main.lbl_1_5)
        setAss.Add(frm_Main.lbl_3_3)
        setAss.Add(frm_Main.lbl_6_1)
        setAss.Add(frm_Main.lbl_6_6)
        setAss.Add(frm_Main.lbl_8_4)
        setAss.Add(frm_Main.lbl_11_2)
        setAss.Add(frm_Main.lbl_13_5)
        setAss.Add(frm_Main.lbl_15_3)
    End Sub

    Sub setBbFretboard()
        setBb.Add(frm_Main.lbl_1_5)
        setBb.Add(frm_Main.lbl_3_3)
        setBb.Add(frm_Main.lbl_6_1)
        setBb.Add(frm_Main.lbl_6_6)
        setBb.Add(frm_Main.lbl_8_4)
        setBb.Add(frm_Main.lbl_11_2)
        setBb.Add(frm_Main.lbl_13_5)
        setBb.Add(frm_Main.lbl_15_3)
    End Sub

    Sub setBFretboard()
        setB.Add(frm_Main.lbl_0_2)
        setB.Add(frm_Main.lbl_2_5)
        setB.Add(frm_Main.lbl_4_3)
        setB.Add(frm_Main.lbl_7_1)
        setB.Add(frm_Main.lbl_7_6)
        setB.Add(frm_Main.lbl_9_4)
        setB.Add(frm_Main.lbl_12_2)
        setB.Add(frm_Main.lbl_14_5)
    End Sub

#End Region
#Region "Reset all Notes when combox changed"
    Sub setAllReset()
        For Each C As Label In setC
            C.Visible = False
            C.BackColor = Color.SeaShell
        Next

        For Each Cs As Label In setCs
            Cs.Visible = False
            Cs.BackColor = Color.SeaShell
        Next

        For Each Db As Label In setDb
            Db.Visible = False
            Db.BackColor = Color.SeaShell
        Next

        For Each D As Label In setD
            D.Visible = False
            D.BackColor = Color.SeaShell
        Next

        For Each Ds As Label In setDs
            Ds.Visible = False
            Ds.BackColor = Color.SeaShell
        Next

        For Each Eb As Label In setEb
            Eb.Visible = False
            Eb.BackColor = Color.SeaShell
        Next

        For Each E As Label In setE
            E.Visible = False
            E.BackColor = Color.SeaShell
        Next

        For Each F As Label In setF
            F.Visible = False
            F.BackColor = Color.SeaShell
        Next

        For Each Fs As Label In setFs
            Fs.Visible = False
            Fs.BackColor = Color.SeaShell
        Next

        For Each Gb As Label In setGb
            Gb.Visible = False
            Gb.BackColor = Color.SeaShell
        Next

        For Each G As Label In setG
            G.Visible = False
            G.BackColor = Color.SeaShell
        Next

        For Each Gs As Label In setGs
            Gs.Visible = False
            Gs.BackColor = Color.SeaShell
        Next

        For Each Ab As Label In setAb
            Ab.Visible = False
            Ab.BackColor = Color.SeaShell
        Next

        For Each A As Label In setA
            A.Visible = False
            A.BackColor = Color.SeaShell
        Next

        For Each Ass As Label In setAss
            Ass.Visible = False
            Ass.BackColor = Color.SeaShell
        Next

        For Each Bb As Label In setBb
            Bb.Visible = False
            Bb.BackColor = Color.SeaShell
        Next

        For Each B As Label In setB
            B.Visible = False
            B.BackColor = Color.SeaShell
        Next
    End Sub
#End Region
#Region "Realize and display all SCALES on fretboard"

    'C Major scale
    Sub CmajScale()
        For Each C As Label In setC
            C.Visible = True
            C.BackColor = Color.LightSalmon
        Next

        For Each D As Label In setD
            D.Visible = True
        Next

        For Each E As Label In setE
            E.Visible = True
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
        Next

        For Each G As Label In setG
            G.Visible = True
        Next

        For Each A As Label In setA
            A.Visible = True
        Next

        For Each B As Label In setB
            B.Visible = True
        Next
    End Sub

    'Db Major scale
    Sub DbmajScale()
        For Each Db As Label In setDb
            Db.Visible = True
            Db.Text = "Db"
            Db.BackColor = Color.LightSalmon
        Next

        For Each Eb As Label In setEb
            Eb.Visible = True
            Eb.Text = "Eb"
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
        Next

        For Each Gb As Label In setGb
            Gb.Visible = True
            Gb.Text = "Gb"
        Next

        For Each Ab As Label In setAb
            Ab.Visible = True
            Ab.Text = "Ab"
        Next

        For Each Bb As Label In setBb
            Bb.Visible = True
            Bb.Text = "Bb"
        Next

        For Each C As Label In setC
            C.Visible = True
        Next
    End Sub

    'D Major scale
    Sub DmajScale()
        For Each D As Label In setD
            D.Visible = True
            D.BackColor = Color.LightSalmon
        Next

        For Each E As Label In setE
            E.Visible = True
        Next

        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
        Next

        For Each G As Label In setG
            G.Visible = True
        Next

        For Each A As Label In setA
            A.Visible = True
        Next

        For Each B As Label In setB
            B.Visible = True
        Next

        For Each Cs As Label In setCs
            Cs.Visible = True
            Cs.Text = "C#"
        Next
    End Sub

    'Eb Major scale
    Sub EbmajScale()
        For Each Eb As Label In setEb
            Eb.Visible = True
            Eb.Text = "Eb"
            Eb.BackColor = Color.LightSalmon
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
        Next

        For Each G As Label In setG
            G.Visible = True
        Next

        For Each Ab As Label In setAb
            Ab.Visible = True
            Ab.Text = "Ab"
        Next

        For Each Bb As Label In setBb
            Bb.Visible = True
            Bb.Text = "Bb"
        Next

        For Each C As Label In setC
            C.Visible = True
        Next

        For Each D As Label In setD
            D.Visible = True
        Next

    End Sub

    'E Major scale
    Sub EmajScale()
        For Each E As Label In setE
            E.Visible = True
            E.BackColor = Color.LightSalmon
        Next

        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
        Next

        For Each Gs As Label In setGs
            Gs.Visible = True
            Gs.Text = "G#"
        Next

        For Each A As Label In setA
            A.Visible = True
        Next

        For Each B As Label In setB
            B.Visible = True
        Next

        For Each Cs As Label In setCs
            Cs.Visible = True
            Cs.Text = "C#"
        Next

        For Each Ds As Label In setDs
            Ds.Visible = True
            Ds.Text = "D#"
        Next

    End Sub

    'F Major scale
    Sub FmajScale()
        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
            F.BackColor = Color.LightSalmon
        Next

        For Each G As Label In setG
            G.Visible = True
        Next

        For Each A As Label In setA
            A.Visible = True
        Next

        For Each Bb As Label In setBb
            Bb.Visible = True
            Bb.Text = "Bb"
        Next

        For Each C As Label In setC
            C.Visible = True
        Next

        For Each D As Label In setD
            D.Visible = True
        Next

        For Each E As Label In setE
            E.Visible = True
        Next

    End Sub

    'Fs Major scale
    Sub FsmajScale()
        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
            Fs.BackColor = Color.LightSalmon
        Next

        For Each Gs As Label In setGs
            Gs.Visible = True
            Gs.Text = "G#"
        Next

        For Each Ass As Label In setAss
            Ass.Visible = True
            Ass.Text = "A#"
        Next

        For Each B As Label In setB
            B.Visible = True
        Next

        For Each Cs As Label In setCs
            Cs.Visible = True
            Cs.Text = "C#"
        Next

        For Each Ds As Label In setDs
            Ds.Visible = True
            Ds.Text = "D#"
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "E#"
        Next

    End Sub

    'G Major scale
    Sub GmajScale()
        For Each G As Label In setG
            G.Visible = True
            G.BackColor = Color.LightSalmon
        Next

        For Each A As Label In setA
            A.Visible = True
        Next

        For Each B As Label In setB
            B.Visible = True
        Next

        For Each C As Label In setC
            C.Visible = True
        Next

        For Each D As Label In setD
            D.Visible = True
        Next

        For Each E As Label In setE
            E.Visible = True
        Next

        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
        Next
    End Sub

    'Ab Major scale
    Sub AbmajScale()
        For Each Ab As Label In setAb
            Ab.Visible = True
            Ab.Text = "Ab"
            Ab.BackColor = Color.LightSalmon
        Next

        For Each Bb As Label In setBb
            Bb.Visible = True
            Bb.Text = "Bb"
        Next

        For Each C As Label In setC
            C.Visible = True
        Next

        For Each Db As Label In setDb
            Db.Visible = True
            Db.Text = "Db"
        Next

        For Each Eb As Label In setEb
            Eb.Visible = True
            Eb.Text = "Eb"
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
        Next

        For Each G As Label In setG
            G.Visible = True

        Next
    End Sub

    'A Major scale
    Sub AmajScale()
        For Each A As Label In setA
            A.Visible = True
            A.BackColor = Color.LightSalmon
        Next

        For Each B As Label In setB
            B.Visible = True
        Next

        For Each Cs As Label In setCs
            Cs.Visible = True
            Cs.Text = "C#"
        Next

        For Each D As Label In setD
            D.Visible = True
        Next

        For Each E As Label In setE
            E.Visible = True
        Next

        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
        Next

        For Each Gs As Label In setGs
            Gs.Visible = True
            Gs.Text = "G#"

        Next
    End Sub

    'Bb Major scale
    Sub BbmajScale()
        For Each Bb As Label In setBb
            Bb.Visible = True
            Bb.Text = "Bb"
            Bb.BackColor = Color.LightSalmon
        Next

        For Each C As Label In setC
            C.Visible = True
        Next

        For Each D As Label In setD
            D.Visible = True
        Next

        For Each Eb As Label In setEb
            Eb.Visible = True
            Eb.Text = "Eb"
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
        Next

        For Each G As Label In setG
            G.Visible = True
        Next

        For Each A As Label In setA
            A.Visible = True
        Next
    End Sub

    'B Major scale
    Sub BmajScale()
        For Each B As Label In setB
            B.Visible = True
            B.BackColor = Color.LightSalmon
        Next

        For Each Cs As Label In setCs
            Cs.Visible = True
            Cs.Text = "C#"
        Next

        For Each Ds As Label In setDs
            Ds.Visible = True
            Ds.Text = "D#"
        Next

        For Each E As Label In setE
            E.Visible = True
        Next

        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
        Next

        For Each Gs As Label In setGs
            Gs.Visible = True
            Gs.Text = "G#"

        Next

        For Each Ass As Label In setAss
            Ass.Visible = True
            Ass.Text = "A#"

        Next
    End Sub

    'C minor scale
    Sub CminScale()
        For Each C As Label In setC
            C.Visible = True
            C.BackColor = Color.LightSalmon
        Next

        For Each D As Label In setD
            D.Visible = True
        Next

        For Each Eb As Label In setEb
            Eb.Visible = True
            Eb.Text = "Eb"
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
        Next

        For Each G As Label In setG
            G.Visible = True
        Next

        For Each Ab As Label In setAb
            Ab.Visible = True
            Ab.Text = "Ab"
        Next

        For Each Bb As Label In setBb
            Bb.Visible = True
            Bb.Text = "Bb"
        Next
    End Sub

    'C# minor scale
    Sub CsminScale()
        For Each Cs As Label In setCs
            Cs.Visible = True
            Cs.Text = "C#"
            Cs.BackColor = Color.LightSalmon
        Next

        For Each Ds As Label In setDs
            Ds.Visible = True
            Ds.Text = "D#"
        Next

        For Each E As Label In setE
            E.Visible = True
        Next

        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
        Next

        For Each Gs As Label In setGs
            Gs.Visible = True
            Gs.Text = "G#"
        Next

        For Each A As Label In setA
            A.Visible = True
        Next

        For Each B As Label In setB
            B.Visible = True
        Next
    End Sub

    'D minor scale
    Sub DminScale()
        For Each D As Label In setD
            D.Visible = True
            D.BackColor = Color.LightSalmon
        Next

        For Each E As Label In setE
            E.Visible = True
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
        Next

        For Each G As Label In setG
            G.Visible = True
        Next

        For Each A As Label In setA
            A.Visible = True
        Next

        For Each Bb As Label In setBb
            Bb.Visible = True
            Bb.Text = "Bb"
        Next

        For Each C As Label In setC
            C.Visible = True
        Next
    End Sub

    'Ds minor scale
    Sub DsminScale()
        For Each Ds As Label In setDs
            Ds.Visible = True
            Ds.Text = "D#"
            Ds.BackColor = Color.LightSalmon
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "E#"
        Next

        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
        Next

        For Each Gs As Label In setGs
            Gs.Visible = True
            Gs.Text = "G#"
        Next

        For Each Ass As Label In setAss
            Ass.Visible = True
            Ass.Text = "A#"
        Next

        For Each B As Label In setB
            B.Visible = True
        Next

        For Each Cs As Label In setCs
            Cs.Visible = True
            Cs.Text = "C#"
        Next
    End Sub

    'E minor scale
    Sub EminScale()
        For Each E As Label In setE
            E.Visible = True
            E.BackColor = Color.LightSalmon
        Next

        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
        Next

        For Each G As Label In setG
            G.Visible = True
        Next

        For Each A As Label In setA
            A.Visible = True
        Next

        For Each B As Label In setB
            B.Visible = True
        Next

        For Each C As Label In setC
            C.Visible = True
        Next

        For Each D As Label In setD
            D.Visible = True
        Next
    End Sub

    'F minor scale
    Sub FminScale()
        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
            F.BackColor = Color.LightSalmon
        Next

        For Each G As Label In setG
            G.Visible = True
        Next

        For Each Ab As Label In setAb
            Ab.Visible = True
            Ab.Text = "Ab"

        Next

        For Each Bb As Label In setBb
            Bb.Visible = True
            Bb.Text = "Bb"
        Next

        For Each C As Label In setC
            C.Visible = True
        Next

        For Each Db As Label In setDb
            Db.Visible = True
            Db.Text = "Db"
        Next

        For Each Eb As Label In setEb
            Eb.Visible = True
            Eb.Text = "Eb"
        Next
    End Sub

    'F# minor scale
    Sub FsminScale()
        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
            Fs.BackColor = Color.LightSalmon
        Next

        For Each Gs As Label In setGs
            Gs.Visible = True
            Gs.Text = "G#"
        Next

        For Each A As Label In setA
            A.Visible = True
        Next

        For Each B As Label In setB
            B.Visible = True
        Next

        For Each Cs As Label In setCs
            Cs.Visible = True
            Cs.Text = "C#"
        Next

        For Each D As Label In setD
            D.Visible = True
        Next

        For Each E As Label In setE
            E.Visible = True
        Next
    End Sub

    'G minor scale
    Sub GminScale()
        For Each G As Label In setG
            G.Visible = True
            G.BackColor = Color.LightSalmon
        Next

        For Each A As Label In setA
            A.Visible = True
        Next

        For Each Bb As Label In setBb
            Bb.Visible = True
            Bb.Text = "Bb"
        Next

        For Each C As Label In setC
            C.Visible = True
        Next

        For Each D As Label In setD
            D.Visible = True
        Next

        For Each Eb As Label In setEb
            Eb.Visible = True
            Eb.Text = "Eb"
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
        Next

    End Sub

    'G# minor scale
    Sub GsminScale()
        For Each Gs As Label In setGs
            Gs.Visible = True
            Gs.Text = "G#"
            Gs.BackColor = Color.LightSalmon
        Next

        For Each Ass As Label In setAss
            Ass.Visible = True
            Ass.Text = "A#"
        Next

        For Each B As Label In setB
            B.Visible = True
        Next

        For Each Cs As Label In setCs
            Cs.Visible = True
            Cs.Text = "C#"
        Next

        For Each Ds As Label In setDs
            Ds.Visible = True
            Ds.Text = "D#"
        Next

        For Each E As Label In setE
            E.Visible = True
        Next

        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
        Next

    End Sub

    'A minor scale
    Sub AminScale()
        For Each A As Label In setA
            A.Visible = True
            A.BackColor = Color.LightSalmon
        Next

        For Each B As Label In setB
            B.Visible = True
        Next

        For Each C As Label In setC
            C.Visible = True
        Next

        For Each D As Label In setD
            D.Visible = True
        Next

        For Each E As Label In setE
            E.Visible = True
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
        Next

        For Each G As Label In setG
            G.Visible = True
        Next

    End Sub

    'Bb minor scale
    Sub BbminScale()
        For Each Bb As Label In setBb
            Bb.Visible = True
            Bb.Text = "Bb"
            Bb.BackColor = Color.LightSalmon
        Next

        For Each C As Label In setC
            C.Visible = True
        Next

        For Each Db As Label In setDb
            Db.Visible = True
            Db.Text = "Db"
        Next

        For Each Eb As Label In setEb
            Eb.Visible = True
            Eb.Text = "Eb"
        Next

        For Each F As Label In setF
            F.Visible = True
            F.Text = "F"
        Next

        For Each Gb As Label In setGb
            Gb.Visible = True
            Gb.Text = "Gb"
        Next

        For Each Ab As Label In setAb
            Ab.Visible = True
            Ab.Text = "Ab"
        Next
    End Sub

    'B minor scale
    Sub BminScale()
        For Each B As Label In setB
            B.Visible = True
            B.BackColor = Color.LightSalmon
        Next

        For Each Cs As Label In setCs
            Cs.Visible = True
            Cs.Text = "C#"
        Next

        For Each D As Label In setD
            D.Visible = True

        Next

        For Each E As Label In setE
            E.Visible = True
        Next

        For Each Fs As Label In setFs
            Fs.Visible = True
            Fs.Text = "F#"
        Next

        For Each G As Label In setG
            G.Visible = True
        Next

        For Each A As Label In setA
            A.Visible = True
        Next

    End Sub

#End Region
#Region "Realize and display all CHORDS on fretboard"

#End Region

End Module
