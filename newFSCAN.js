
var fs = require("fs");
const { resolve } = require("path");
var path = require("path");
var xlsx = require("xlsx");




async function stepOne() {
var files = fs.readdirSync('V:/FSCAN_ENV/History/');

var targetFiles1 = files.filter(function (file) {
    return file.startsWith('FSCAN');
})

var targetFiles = targetFiles1.filter(function (file) {
    return path.extname(file).toLowerCase() === ".xls" || path.extname(file).toLowerCase() === ".xlsx";
})


    var sorFile = targetFiles.sort((a, b) => b.localeCompare(a, undefined, { numeric: true, sensitivity: 'base' }))[0];
    console.log(sorFile);
    var origFile =  fs.createReadStream('V:/FSCAN_ENV/History/' + sorFile);

    var copiedFile = fs.createWriteStream('Z:/FSCAN_Automation/2' + sorFile);

    const promise = new Promise(resolve => {
      origFile.pipe(copiedFile).on('finish', () => {
        resolve();
      })
    })

    await promise;

      var addToname = sorFile.substring(6,18);

      var filesOnD = fs.readdirSync('Z:/FSCAN_Automation/');

      var targetfilesOnD = filesOnD.filter(function(file) {
          return path.extname(file).toLowerCase() === ".xls" || path.extname(file).toLowerCase() === ".xlsx";
      })

      var newSorfile = targetfilesOnD.sort((a, b) => b.localeCompare(a, undefined, { numeric: true, sensitivity: 'base' }))[0];

      return {
        file: newSorfile,
        addToname,
        sorFile
      };
}

    // console.log(contargetfilesOnD);

  async function genFscan(file, addToname){


        var wb = xlsx.readFile(`${file}`);

        

        var ws = wb.Sheets["Summary"];


        var data = xlsx.utils.sheet_to_json(ws);

        var newData = data.map(function(record){
            delete record.Interrupted;
            delete record.Committed;
            delete record.Purging;
            delete record.Crashed;
            delete record.StartTimeRetrieval;
            delete record.EndTimeRetrieval;
            delete record.DeviceLogCount;
            delete record.FileLogCount;
            // delete record.Connected;
            return record;
            })

    // Get disconnected  
        var disconnected = newData.filter(row => {
            return (row.Disconnected === "True");
        });
        const fileDatadis = newData.filter(row => row.Disconnected === "True");

    // Get connected    var connected = newData.filter(row => row.Connected === "True")
        const fileDatacon = newData.filter(row => row.Connected === "True" );
        //company
        const companyCellWidth = disconnected.sort((a, b) => {
            return b.Company.length - a.Company.length;
            })[0].Company.length;

        // branch
            const branchCellWidth = disconnected.sort((a, b) => {
            return b.Branch.length - a.Branch.length;
            })[0].Branch.length;

        // name 
            const nameCellWidth = disconnected.sort((a, b) => {
            return b.Name.length - a.Name.length;
            })[0].Name.length;
        // ip 

            const ipCellWidth = disconnected.sort((a, b) => {
            return b.IP.length - a.IP.length;
            })[0].IP.length;

    // console.log(Object.keys(companyCellWidth).length);

        var newWB = xlsx.utils.book_new();


        // var newWSdis = xlsx.utils.json_to_sheet(fileDatadis);

        // var sheet1 = xlsx.utils.sheet_to_json(newWSdis);

            var delcon = fileDatadis.map(function(record){
                delete record.Connected;
                return record;
            })

            var deldis = fileDatacon.map(function(record){
                delete record.Disconnected;
                return record;
            })

        var newWSdis = xlsx.utils.json_to_sheet(delcon);

        var newWScon = xlsx.utils.json_to_sheet(deldis);

        // var sheet2 = xlsx.utils.sheet_to_json(newWScon)

    // // Set cell width 
        const colsConfig = [
            {wch: companyCellWidth + 2},
            {wch: branchCellWidth + 2}, 
            {wch: nameCellWidth + 2}, 
            {wch: ipCellWidth + 2},
            {wch: 10}
        ];


        newWSdis["!cols"] = colsConfig;
        newWScon["!cols"] = colsConfig;



        xlsx.utils.book_append_sheet(newWB, newWSdis, "Disconnected");


        xlsx.utils.book_append_sheet(newWB, newWScon, "Connected");


        xlsx.writeFile(newWB, "2FSCAN Disconnected G7 Devices_" + addToname + ".xlsx");

        const promise = new Promise((resolve) => {
            fs.rename("Z:/FSCAN_Automation/2FSCAN Disconnected G7 Devices_" + addToname + ".xlsx", "Z:/FSCAN_Automation/Generated Files/FSCAN Disconnected G7 Devices_" + addToname + ".xlsx", (err) => {
                if (err) {
                    throw err;
                    console.log('Rename complete!');
                }
                else{
                    resolve();
                }
        } )

      });

      await promise;

    }
    
   async function genFscanFW(file, addToname){ 


        var xlsx = require("xlsx");

        var wb2 = xlsx.readFile(file);

        var ws2 = wb2.Sheets["Summary"];

        var data2 = xlsx.utils.sheet_to_json(ws2);

        var newData2 = data2.map(function(record){
            delete record.Interrupted;
            delete record.Committed;
            delete record.Purging;
            delete record.Crashed;
            delete record.StartTimeRetrieval;
            delete record.EndTimeRetrieval;
            delete record.DeviceLogCount;
            delete record.FileLogCount;
            return record;
            })

    // Get disconnected  
        var company = newData2.filter(row => row.Company === "FAMILYHEALTH & BEAUTY CORP." || row.Company === "WATSONS PERSONAL CARE STORES(PHILS)");
        const fileDatadis2 = newData2.filter(row => row.Company === "FAMILYHEALTH & BEAUTY CORP." || row.Company ==="WATSONS PERSONAL CARE STORES(PHILS)");

        //company
        const companyCellWidth = company.sort((a, b) => {
            return b.Company.length - a.Company.length;
            })[0].Company.length;

        // branch
            const branchCellWidth = company.sort((a, b) => {
            return b.Branch.length - a.Branch.length;
            })[0].Branch.length;

        // name 
            const nameCellWidth = company.sort((a, b) => {
            return b.Name.length - a.Name.length;
            })[0].Name.length;
        // ip 

            const ipCellWidth = company.sort((a, b) => {
            return b.IP.length - a.IP.length;
            })[0].IP.length;

    // console.log(Object.keys(companyCellWidth).length);

        var newWB2 = xlsx.utils.book_new();

        var newWSdis2 = xlsx.utils.json_to_sheet(fileDatadis2);

    // // Set cell width 
        const colsConfig = [
            {wch: companyCellWidth + 2},
            {wch: branchCellWidth + 2}, 
            {wch: nameCellWidth + 2}, 
            {wch: ipCellWidth + 2},
            {wch: 10}
        ];


        newWSdis2["!cols"] = colsConfig;

        xlsx.utils.book_append_sheet(newWB2, newWSdis2, "Disconnected");

        xlsx.writeFile(newWB2, "2Disconnected G7 Devices (Watsons and Family Health)_" + addToname + ".xlsx");

        const promise = new Promise(() => {
            fs.rename("Z:/FSCAN_Automation/2Disconnected G7 Devices (Watsons and Family Health)_" + addToname + ".xlsx", "Z:/FSCAN_Automation/Generated Files/Disconnected G7 Devices (Watsons and Family Health)_" + addToname + ".xlsx", (err) => {
            err ? console.log(err): resolve();
              });
        })

        await promise;
    }

    async function delFile(file) {
      const promise = new Promise((resolve) => {
        fs.unlinkSync('Z:/FSCAN_Automation/2' + file)
        resolve();
      })

      await promise;
    };

    

    // new Promise(stepOne)
    // .then(genFscan)
    // .then(genFscanFW)
    // .then(delFile)
    // .catch(console.log('Error'));

    // new Promise((resolve) => {
    //   const file = stepOne();
    //   resolve(file);
    // })

    // new Promise((resolve) => {
    //   const file = stepOne(); // returns a file 
    //   resolve(file)
    // })
    //   .then((file) => {
    //     if (file) {
    //       console.log(file);
    //       genFscan(file);
    //     }
    //   })
    // .catch(err => console.log('Error: ', err));

    async function getData() {
      const {file, addToname, sorFile } = await stepOne();

      const promises = [
         genFscan(file, addToname),
         genFscanFW(file, addToname),
         delFile(sorFile)
      ]

      await Promise.all(promises);
      
    };
    
    getData();
