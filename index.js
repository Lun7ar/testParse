let XLSX = require("xlsx");

const parse = (filename) => {
    const excelData = XLSX.readFile(filename);
    const worksheet = excelData.Sheets[excelData.SheetNames[0]];
    
    return Object.keys(worksheet).filter(x => /^D\d+/.test(x)).map((x) =>  ({ 
        x,
        column: worksheet[x].v,
    }));
};

parse("./example.xlsx").forEach((element) => {
    console.log(element.column);
});
