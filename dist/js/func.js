

Skip to content
GitHub
This repository

    Pull requests
    Issues
    Gist

    @leanham

1
0

    1

BoisLAB/gKing

gKing/func.js
@leanham leanham 6 days ago Create func.js

1 contributor
37 lines (31 sloc) 1.087 kB
/* ============================================================
 * Main function v0.1.0
 * ============================================================ */

 var arrBaseNoteSharp = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]
 var arrBaseNoteFlat = ["C", "Db", "D", "Eb", "E", "F", "Gb", "G", "Ab", "A", "Bb", "B"];
 
@@ -20,33 +12,29 @@
     for (i = 0; i < 6; i++) { 
         arr[i] = (arrNumMaj[i] + offset) % 12;
     }
 
     return arr;
 }
 
 function getArrayNotes(arrayNum, isSharp) {
         var arr = ["", "", "", "", "", ""];
         for (i = 0; i < 6; i++) { 
             if (arrayNum[i] == 0 ){
                 arr [i] = "B"
             }
             else{
                 if(isSharp == 1)
                     arr [0] = arrBaseNoteSharp[arrayNum[i] - 1]
                 else
                     arr [i] = arrBaseNoteFlat[arrayNum[i] - 1]
             }
         }
        return arr ;
 }
 
 function myFunction() {
    var arr = getArrayNumMaj(3);
    return getArrayNotes(arr,0);
 }
-document.getElementById("demo").innerHTML = myFunction();

    Status API Training Shop Blog About Help 

    © 2015 GitHub, Inc. Terms Privacy Security Contact 

