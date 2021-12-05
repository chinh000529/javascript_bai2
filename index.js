const XLSX = require("xlsx");

// convert file xlsx to array
const wb = XLSX.readFile('master.xlsx');
const ws = wb.SheetNames[0];
const data = XLSX.utils.sheet_to_json(wb.Sheets[ws]);
const result = data.map(item => {
    return Object.values(item);
})
console.log(result);


// convert file json to xlsx
const newWB = XLSX.utils.book_new();
const newWS = XLSX.utils.json_to_sheet([
    {
        "id": "#2905",
        "name": "Chinh",
        "age": 21
    },
    {
        "id": "#2807",
        "name": "Nga",
        "age": 21
    },
]);
XLSX.utils.book_append_sheet(newWB, newWS, "information");
XLSX.writeFile(newWB, 'information.xlsx');
