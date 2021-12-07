// const XLSX = require("xlsx");

// // convert file xlsx to array
// const wb = XLSX.readFile('master.xlsx');
// const ws = wb.Sheets[wb.SheetNames[0]];
// console.log(ws);
// const result = [];
// let item = {};
// for (let cell in ws) {
//     const cellAsString = cell.toString();
//     console.log(cellAsString);
//     if (cellAsString[1] !== 'r' && cellAsString !== 'm' && cellAsString[1] > 1) {
//         if (cellAsString[0] === 'A') {
//             item.id = ws[cell].v;
//         }
//         if (cellAsString[0] === 'B') {
//             item.no = ws[cell].v;
//         }
//         if (cellAsString[0] === 'P') {
//             item.a = ws[cell].v;
//         }
//         if (cellAsString[0] === 'Q') {
//             item.b = ws[cell].v;
//         }
//         if (cellAsString[0] === 'R') {
//             item.c = ws[cell].v;
//         }
//         if (cellAsString[0] === 'S') {
//             item.d = ws[cell].v;
//             result.push(item);
//             item = {};
//         }
//     }
// }
// const newResult = result.map(item => {
//     return Object.values(item);
// })
// console.log(newResult);



const XLSX = require('xlsx');

const wb = XLSX.readFile('./master.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const cols = ["A", "B", "P", "Q", "R", "S"];
const maxRow = XLSX.utils.decode_range(ws["!ref"]).e.r;

const result = [];
for (let i = 2; i <= maxRow + 1; i++) {
    result.push(
        cols.map(col => {
            return +ws[col + i].w;
        })
    )
}
console.log(result);

