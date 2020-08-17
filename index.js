const fetch = require("node-fetch");
const XLSX = require("xlsx");
const fse = require("fs-extra");

// Scrape data from this URL
const url = "http://www.provinz.bz.it/news/de/news.asp?news_action=300&news_image_id=1077270";

function main() {
    fetch(url)
        .then(function (res) {
            if (!res.ok) throw new Error("fetch failed");
            return res.arrayBuffer();
        })
        .then(function (ab) {
            // get Excel Sheet informations
            const data = new Uint8Array(ab);
            const workbook = XLSX.read(data, { type: "array" });
            const sheetContent = workbook.Sheets[workbook.SheetNames[0]];
            const range = XLSX.utils.decode_range(sheetContent["!ref"]);
            const row_start = range.s.r;
            const row_end = range.e.r;

            const covid_data = [];

            // Newest date in sheet (get it it from cell E3)
            let dt = sheetContent.E3.v.replace("Totali al ", "");

            for (let i = row_start; i < row_end; i++) {
                let cellMunicipality = "B" + i,
                    cellTotalPositivesToday = "E" + i,
                    cellTotalPositivesYesterday = "D" + i,
                    cellTotalCuredToday = "H" + i,
                    cellTotalCuredYesterday = "G" + i,
                    cellDeceased = "J" + i,
                    cellActivePositives = "K" + i;
                if (sheetContent[cellMunicipality] !== undefined) {
                    // Get rows which contain the string "Totale"
                    if (sheetContent[cellMunicipality].v.includes("Totale")) {
                        if (
                            // don't push last 2 rows ("Comune sconosciuto Totale" and "Totale complessivo")
                            sheetContent[cellMunicipality].v !== "Totale complessivo" &&
                            sheetContent[cellMunicipality].v !== "Comune sconosciuto Totale"
                        ) {
                            covid_data.push({
                                municipality: sheetContent[cellMunicipality].v.replace(" Totale", ""),
                                totalToday: sheetContent[cellTotalPositivesToday].v,
                                totalYesterday: sheetContent[cellTotalPositivesYesterday].v,
                                increaseSinceDayBefore:
                                    sheetContent[cellTotalPositivesToday].v -
                                    sheetContent[cellTotalPositivesYesterday].v,
                                totalCuredToday: sheetContent[cellTotalCuredToday].v,
                                totalCuredYesterday: sheetContent[cellTotalCuredYesterday].v,
                                deceased: sheetContent[cellDeceased].v,
                                activePositives: sheetContent[cellActivePositives].v,
                            });
                        }
                    }
                }
            }

            const parts = dt.split("-");
            const newestDateInSheet = new Date(
                parseInt(parts[2], 10),
                parseInt(parts[1], 10) - 1,
                parseInt(parts[0], 10)
            );

            const newestDateInSheet_formatted = formatDate(newestDateInSheet),
                todaysDate = new Date(),
                todaysDate_formatted = formatDate(todaysDate);

            // check if data is from today
            if (newestDateInSheet_formatted === todaysDate_formatted) {
                //save in file
                fse.outputFile(`output/${newestDateInSheet_formatted}.json`, JSON.stringify(covid_data), (err) => {
                    if (err) {
                        console.log(err);
                    } else {
                        console.log("The file was saved!");
                    }
                });
            } else {
                console.log("Data is not from today!");
            }
        })
        .catch((error) => {
            console.log(error);
        });
}

function formatDate(date) {
    const day = date.getDate(),
        month = date.getMonth() + 1,
        year = date.getFullYear(),
        date_formatted = day + "-" + month + "-" + year;

    return date_formatted;
}

main();
