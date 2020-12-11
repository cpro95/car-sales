const axios = require('axios');
const cheerio = require('cheerio');
const XLSX = require('xlsx');

const getHtml = async (url) => {
    try {
        return await axios.get(url);
    } catch (err) {
        console.error(error);
    }
};

// global variable for excel output
var ws_data = [];

// main execution
for (i = 1; i < 12; i++) {
    let month = '';
    if (i < 10) {
        month = '2020-0' + i;
    } else {
        month = '2020-' + i;
    }

    let car_url = 'http://auto.danawa.com/auto/?Work=record&Tab=Model&Brand=303,304,307&Month=' + month + '-00';

    getHtml(car_url)
        .then(html => {
            // console.log(html);
            const $ = cheerio.load(html.data);
            var car_result = {};

            var nameArr = [];
            $('table.recordTable')
                .find('tbody tr')
                .find('td.title')
                .find('a')
                .each((i, el) => {
                    nameArr.push(
                        $(el)
                            .text()
                            .replace(/\n/g, '')
                            .replace(/ /g, '')
                    );
                });
            car_result.nameArr = nameArr;

            var numberArr = [];
            $('table.recordTable')
                .find('tbody tr')
                .find('td.right')
                .each((i, el) => {
                    if ((i + 1) % 3 == 2) {
                        var prev = $(el).text().replace(/,/g, '');
                        var chil = $(el).children().text().replace(/,/g, '');
                        var prev2 = prev.replace(chil, '');
                        var chil2 = chil.slice(0, -1);
                        if ($(el).text().includes("▼")) {
                            var final = Number(prev2) - Number(chil2);
                            numberArr.push(
                                final
                            );
                        } else if ($(el).text().includes("▲")) {
                            var final = Number(prev2) + Number(chil2);
                            numberArr.push(
                                final
                            )
                        } else {
                            numberArr.push(
                                Number(prev2)
                            )
                        }
                    }
                });

            car_result.numberArr = numberArr;
            return car_result;
        })
        .then(res => {
            for (var i = 0; i < res.nameArr.length; i++) {
                var imsiArr = [];
                imsiArr.push(res.nameArr[i]);
                imsiArr.push(res.numberArr[i]);
                imsiArr.push(month);
                ws_data.push(imsiArr);
            }

            // save to excel file
            var wb = XLSX.utils.book_new();
            wb.SheetNames.push("CarSales");
            // var ws_data = [['hello' , 'world']];  //a row with 2 columns
            var ws = XLSX.utils.aoa_to_sheet(ws_data);
            wb.Sheets["CarSales"] = ws;
            XLSX.writeFile(wb, 'out.xlsx');

        })
        .catch(error => console.log(error));
}