const XLSX = require("xlsx");

// convert file xlsx to array
const wb = XLSX.readFile('master.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const result = [];
let item = {};
for (let cell in ws) {
    const cellAsString = cell.toString();
    if (cellAsString[1] !== 'r' && cellAsString !== 'm' && cellAsString[1] > 1) {
        if (cellAsString[0] === 'A') {
            item.id = ws[cell].v;
        }
        if (cellAsString[0] === 'B') {
            item.no = ws[cell].v;
        }
        if (cellAsString[0] === 'P') {
            item.a = ws[cell].v;
        }
        if (cellAsString[0] === 'Q') {
            item.b = ws[cell].v;
        }
        if (cellAsString[0] === 'R') {
            item.c = ws[cell].v;
        }
        if (cellAsString[0] === 'S') {
            item.d = ws[cell].v;
            result.push(item);
            item = {};
        }
    }
}
const newResult = result.map(item => {
    return Object.values(item);
})
console.log(newResult);


// // convert file json to xlsx
// const newWB = XLSX.utils.book_new();
// const newWS = XLSX.utils.json_to_sheet([
//     {
//         "id": "#2905",
//         "name": "Chinh",
//         "age": 21
//     },
//     {
//         "id": "#2807",
//         "name": "Nga",
//         "age": 21
//     },
// ]);
// XLSX.utils.book_append_sheet(newWB, newWS, "information");
// XLSX.writeFile(newWB, 'information.xlsx');

