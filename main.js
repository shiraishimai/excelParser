const OUTPUT_JSON = "list.json",
    SOURCE_FILE = "source.html",
    OUTPUT_FILE = "index.html";
let fs = require("fs"),
    cwd = process.cwd(),
    XLSX = require("xlsx"),
    path = require("path"),
    spawn = require("child_process").spawn,
    exe = (cmd, callback) => {
        console.log("cmd>>", cmd);
        let arg = Array.isArray(cmd) ? cmd : cmd.split(" "),
            command = arg.shift(),
            subProcess = spawn(command, arg, {
                stdio: "inherit"
            });
        subProcess.on("exit", code => {});
        subProcess.on("close", code => {
            // @TODO: NEED BACKSLASH!!!!!!!!!!!
            console.log("Close with code:", code);
            if (code === 0) return callback();
            return callback(code);
        });
        subProcess.on("error", code => {
            // @TODO: NEED BACKSLASH!!!!!!!!!!!
            console.log("Error with code:", code);
            return callback(code);
        });
    };
fs.readdir(cwd, (err, fileList) => {
    fileList = fileList.filter(file => path.extname(file) === ".xlsx");
    fileList.map(file => {
            return {
                "name": file,
                "time": fs.statSync(file).mtime.getTime()
            };
        })
        .sort((a, b) => a.time - b.time) // Sort with ascending order
        .map(file => file.name);
    let excelFile = fileList.pop(); // Get the latest modified file
    console.log("Excel file:", excelFile);
    fs.readFile(path.resolve(cwd, excelFile), (err, buffer) => {
        if (err) return console.log("Error:", err);
        let data = new Uint8Array(buffer),
            arr = new Array();
        for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
        let bitString = arr.join(""),
            workbook = XLSX.read(bitString, {type: 'binary'}),
            worksheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Parser
        Promise.all([
            prepareJSON(parseToJSON(worksheet, detectRange(worksheet))),
            prepareHTML(excelFile)
        ]).then(() => {
            // Upload to server
            console.log("Uploading to server...");
        });
    });
});

function prepareJSON(listArray) {
    return new Promise((resolve, reject) => {
        fs.writeFile(OUTPUT_JSON, JSON.stringify(listArray), err => {
            if (err) throw err;
            console.log("Completed:", OUTPUT_JSON);
            return resolve(OUTPUT_JSON);
        });
    });
}
function prepareHTML(excelFile) {
    return new Promise((resolve, reject) => {
        fs.readFile(SOURCE_FILE, "utf8", (err, data) => {
            if (err) throw err;
            // @TODO: NEED BACKSLASH!!!!!!!!!!!
            data = data.replace(/CSV_VERSION/g, excelFile);
            fs.writeFile(OUTPUT_FILE, data, "utf8", err => {
                if (err) throw err;
                console.log("Completed:", OUTPUT_FILE);
                return resolve(OUTPUT_FILE);
            });
        });
    });
}

function detectRange(worksheet) {
    // Calculating number of records
    let range = XLSX.utils.decode_range(worksheet["!ref"]),
        isCounting = false,
        numOfRecords = 0,
        row = (range.s.r+1),    // Cell starts from 1 not 0
        startingRow, endingRow;
    for (; row <= (range.e.r+1); row++) {   // Need to check the last row
        if (worksheet['A'+row] && worksheet['A'+row].t === 'n') {
            if (!isCounting) {
                console.log("Starting record: A"+(startingRow = row));
                isCounting = true;
            }
            numOfRecords++;
        } else if (isCounting) {
            row = row-1;
            break;
        }
    }
    console.log("Ending record: A"+(endingRow = row));
    return {numOfRecords, startingRow, endingRow};
}

function parseToJSON(worksheet, range) {
    let list = [];
    for (let _row = range.startingRow; _row <= range.endingRow; _row++) {
        list.push({
            "fileId": worksheet['A'+_row].w,
            "url": worksheet['C'+_row].w
        });
    }
    return list;
}