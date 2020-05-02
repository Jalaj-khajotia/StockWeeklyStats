console.log('Application loaded');

var XLSX = require('xlsx');
var fileName1 = '27.csv';
var fileName2 = '30.csv';
var minPrice = 10;
console.log('File used is ' + fileName1);
var workbook1 = XLSX.readFile(fileName1);
console.log('File used is ' + fileName2);
var workbook2 = XLSX.readFile(fileName2);
var sheet_name_list1 = workbook1.SheetNames;
var sheet_name_list2 = workbook2.SheetNames;

const readline = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout
});

var excel2 = [];
var excel1 = [];
sheet_name_list1.forEach(function (y) {
    var worksheet = workbook1.Sheets[y];
    var headers = {};

    for (z in worksheet) {
        if (z[0] === '!')
            continue;
        //parse out the column, row, and value
        var tt = 0;
        for (var i = 0; i < z.length; i++) {
            if (!isNaN(z[i])) {
                tt = i;
                break;
            }
        };
        var col = z.substring(0, tt);
        var row = parseInt(z.substring(tt));
        var value = worksheet[z].v;

        //store header names
        if (row == 1 && value) {
            headers[col] = value;
            continue;
        }

        if (!excel1[row])
            excel1[row] = {};
        excel1[row][headers[col]] = value;
    }
    //drop those first two rows which are empty
    excel1.shift();
    excel1.shift();

});

sheet_name_list2.forEach(function (y) {
    var worksheet = workbook2.Sheets[y];
    var headers = {};

    for (z in worksheet) {
        if (z[0] === '!')
            continue;
        //parse out the column, row, and value
        var tt = 0;
        for (var i = 0; i < z.length; i++) {
            if (!isNaN(z[i])) {
                tt = i;
                break;
            }
        };
        var col = z.substring(0, tt);
        var row = parseInt(z.substring(tt));
        var value = worksheet[z].v;

        //store header names
        if (row == 1 && value) {
            headers[col] = value;
            continue;
        }

        if (!excel2[row])
            excel2[row] = {};
        excel2[row][headers[col]] = value;
    }
    //drop those first two rows which are empty
    excel2.shift();
    excel2.shift();

    mydata = [];
    var winners = [];
    var i = 0,
    limit = 0;
});

var stockList1 = [];
var stockList2 = [];
var no = 0, j = 0;
var mainlist = [];
excel1.forEach(function (stock) {
    stockList1[no++] = {
        SYMBOL: stock.SYMBOL,
        SERIES: stock.SERIES,
        OPEN1: stock.OPEN,
        HIGH1: stock.HIGH,
        LOW1: stock.LOW,
        CLOSE1: stock.CLOSE,
        LAST1: stock.LAST,
        TOTTRDQTY1: stock.TOTTRDQTY,
        TOTTRDVAL1: stock.TOTTRDVAL
    }
})

no = 0;
excel2.forEach(function (stock) {
    stockList2[no++] = {
        SYMBOL: stock.SYMBOL,
        SERIES: stock.SERIES,
        OPEN2: stock.OPEN,
        HIGH2: stock.HIGH,
        LOW2: stock.LOW,
        CLOSE2: stock.CLOSE,
        LAST2: stock.LAST,
        TOTTRDQTY2: stock.TOTTRDQTY,
        TOTTRDVAL2: stock.TOTTRDVAL
    }
});

//console.log(excel1.length);
stockList1.forEach(function (item) {
    stockList2.forEach(function (item2) {
        if (item.SYMBOL == item2.SYMBOL && item.SERIES == item2.SERIES) {
            mainlist[j] = {
                SYMBOL: item.SYMBOL,
                SERIES: item.SERIES,
                OPEN1: item.OPEN1,
                HIGH1: item.HIGH1,
                LOW1: item.LOW1,
                CLOSE1: item.CLOSE1,
                LAST1: item.LAST1,
                TOTTRDQTY1: item.TOTTRDQTY1,
                TOTTRDVAL1: item.TOTTRDVAL1,
                OPEN2: item2.OPEN2,
                HIGH2: item2.HIGH2,
                LOW2: item2.LOW2,
                CLOSE2: item2.CLOSE2,
                LAST2: item2.LAST2,
                TOTTRDQTY2: item2.TOTTRDQTY2,
                TOTTRDVAL2: item2.TOTTRDVAL2
            }
            j++;
        }
    });
});

//console.log(mainlist.length);
readline.question(`Enter the 1. Weekly Gainers 2. Weekly Loosers `, (job) => {
    switch (job) {
    case '1':
        WeeklyGainers();
        break;
    case '2':
        WeeklyLoosers();
        break;

    default:
        console.log('kaboom');
    }
});

var weekGainerList = [], weekLooserList = [], i = 0;
function WeeklyGainers() {
    readline.question(`Enter the lower limit? `, (lower) => {
        readline.question('Enter Upper Limit, Default:100% ', (upper) => {
            mainlist.forEach(function (cell) {
                var riseFall = (cell.CLOSE2 - cell.OPEN1) * 100 / cell.OPEN1;
                var roundGainer = Math.round(riseFall * 100) / 100;
                limit = upper == 0 ? 100 : upper;
                if (roundGainer >= lower && roundGainer < upper && cell.TOTTRDQTY1 > 10000 && cell.OPEN1 >= minPrice) {

                    weekGainerList[i++] = {
                        Symbol: cell.SYMBOL,
                        Percentage: roundGainer,
                        CMP: cell.CLOSE2
                    }
                };
            });
            weekGainerList.sort(function (a, b) {
                return a.Percentage - b.Percentage;
            });

            console.log('');
            console.log('Total no of stocks found are ' + weekGainerList.length);
            console.log('');
            console.log('Listing stocks which rose > ' + lower + '% but are lower than  < ' + upper + '%');
            console.log('');
            console.log('  ' + 'Stock Name' + '\t ' + '% Increase' + '   CMP');
            weekGainerList.forEach(function (stock) {
                console.log('  ' + stock.Symbol + '\t ' + stock.Percentage + '\t     ' + stock.CMP);
            })
            // console.log(weekGainerList);
            readline.close()
        });
    });
}

function WeeklyLoosers() {
    i = 0;
    readline.question(`Enter the lower limit? `, (lower) => {
        readline.question('Enter Upper Limit, Default:100% ', (upper) => {
            console.log((mainlist[0].CLOSE2 - mainlist[0].CLOSE1) * 100 / mainlist[0].CLOSE1);
            mainlist.forEach(function (cell) {
                var riseFall = (cell.CLOSE2 - cell.OPEN1) * 100 / cell.OPEN1;
                var roundGainer = Math.round(riseFall * 100) / 100;
                limit = upper == 0 ? 100 : upper;
                if (roundGainer <= -lower && roundGainer > -upper && cell.TOTTRDQTY1 > 10000) {

                    weekLooserList[i++] = {
                        Symbol: cell.SYMBOL,
                        Percentage: roundGainer,
                        CMP: cell.CLOSE2
                    }
                };
            });
            weekLooserList.sort(function (a, b) {
                return a.Percentage - b.Percentage;
            });

            console.log('');
            console.log('Total no of stocks found are ' + weekLooserList.length);
            console.log('');
            console.log('Listing stocks which fell > ' + lower + '% but are lower than  < ' + upper + '%');
            console.log('');

            console.log('  ' + 'Stock Name' + '\t ' + '% Decrease' + '   CMP');
            weekLooserList.forEach(function (stock) {
                console.log('  ' + stock.Symbol + '\t ' + stock.Percentage + '\t     ' + stock.CMP);
            })
            // console.log(weekLooserList);
            readline.close()
        });
    });
}