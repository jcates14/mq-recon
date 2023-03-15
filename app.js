var XLSX = require("xlsx");

var workbook = XLSX.readFile('NewestTest copy.xlsx');
var sheet_name_list = workbook.SheetNames; // list of sheet names
console.log(sheet_name_list);

let events_ws = workbook.Sheets['events']; // Get the sheet "events"
let files_ws = workbook.Sheets['files']; // Get the sheet "files"

var events_range = XLSX.utils.decode_range(events_ws['!ref']); // Get range of "events" sheet
var events_num_rows = events_range.e.r - events_range.s.r + 1; // number of rows on "events" sheet
var events_num_cols = events_range.e.c - events_range.s.c + 1; // number of columns on "events" sheet
console.log(events_range);
console.log(events_num_rows, events_num_cols);
 
var files_range = XLSX.utils.decode_range(files_ws['!ref']); // Get range of "files" sheet
var files_num_rows = files_range.e.r - files_range.s.r + 1; // number of rows on "files" sheet
var files_num_cols = files_range.e.c - files_range.s.c + 1; //  number of columns on "events" sheet
console.log(files_range);
console.log(files_num_rows, files_num_cols);



for (let i = 2; i < events_num_rows; i++) {
    console.log("Reading row " + i + " from events sheet");              // Loop through "events" sheet
    //console.log(events_ws["I" + i].v);
    if (events_ws["I" + i].v === "null") {
        console.log("continuing because I" + i + " is null");
        continue;
    } else {
        console.log("nri_id is not null...");
        let nri_id = events_ws["J" + i].v;
        let mq_id = events_ws["A" + i].v;
        console.log("nri_id is " + nri_id + " and mq_id is " + mq_id);
        for (j = 2; j < files_num_rows; j++) {
            console.log(files_ws["Reading row " + j + " from files sheet"]);
            console.log(files_ws["A" + j].v);
            if ((nri_id === files_ws["A" + j].v && files_ws["T" + j].v === "null") || (nri_id === files_ws["B" + j].v && files_ws["T" + j].v === "null")) {
                console.log(files_ws["A" + j].v + " files column A");
                console.log("Matched " + nri_id + " on files row " + j);
                let files_id = files_ws["C" + j].v;
                files_ws["T" + j].v = mq_id;
                events_ws["AB" + i].v = files_id;
                break; // adding this

            }
        }

    };
}

for (let a = 2; a < events_num_rows; a++) {
    console.log("Reading row " + a + " from events sheet. Last loop!"); 
    if (events_ws["AB" + a].v === "null") {
        let lastfour_id = events_ws["Z" + a].v;
        let mq_id = events_ws["A" + a].v;
        console.log("lastfour_id is " + lastfour_id + " and mq_id is " + mq_id);
        for (b = 2; b < files_num_rows; b++) {
            if (lastfour_id === files_ws["U" + b].v && files_ws["T" + b].v === "null") {
                let files_id = files_ws["C" + b].v;
                files_ws["T" + b].v = mq_id;
                events_ws["AB" + a].v = files_id;
                break; //adding this
            }
        }
    }
}





//console.log(XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]))
//var test = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);


//worksheet['B2'].v = 'TEST';
//console.log(worksheet['B2'].v); // Get cell value of B2


XLSX.writeFile(workbook, 'Newest new new.xlsx'); // write/save new file