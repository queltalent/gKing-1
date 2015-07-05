var arrBaseNoteSharp = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"];
var arrBaseNoteFlat = ["C", "Db", "D", "Eb", "E", "F", "Gb", "G", "Ab", "A", "Bb", "B"];

var arrNumMaj = [1, 3, 5, 6, 8, 10, 12];
var arrNumMin = [1, 3, 4, 6, 8, 9, 11];

//Major Scale
function getArrNumMaj(targetNum,isSharp){
    var arrNum = new Array(1,1,1,1,1,1,1);/*ini Array*/
    var arrResult = ["", "", "","", "", ""];/*ini Array*/
    var offset =  targetNum - 1;
    /*calculat the arrNum by offset and put the result in array*/
    for (i = 0; i < 7; i++) { 
        arrNum[i] = (arrNumMaj[i] + offset) % 12;
		if (arrNum[i] == 0 ){
            arrResult [i] = "B";
        }
        else{
            if(isSharp == 1)
                arrResult [0] = arrBaseNoteSharp[arrNum[i] - 1];
            else
                arrResult [i] = arrBaseNoteFlat[arrNum[i] - 1];
        }
    }
    return arrResult ;
}
//Minor
function getArrNumMin(targetNum,isSharp){ /*Minor*/
    var arrNum = new Array(1,1,1,1,1,1,1);
    var arrResult = ["", "", "","", "", ""];
    var offset =  targetNum - 1;

    for (i = 0; i < 7; i++) {
        arrNum[i] = (arrNumMin[i] + offset) % 12;
        if (arrNum[i] == 0 ){
            arrResult [i] = "B";
        }
        else{
            if(isSharp == 1)
                arrResult [0] = arrBaseNoteSharp[arrNum[i] - 1];
            else
                arrResult [i] = arrBaseNoteFlat[arrNum[i] - 1];
        }
    }
    return arrResult ;
}
