const fetch = require("node-fetch");
const XLSX = require("xlsx");
const fse = require("fs-extra");

// Scrape data from this URL
// URL has to be changed manually every day
const url = "https://www.sabes.it/de/news.asp?aktuelles_action=300&aktuelles_image_id=1081139";

const listOfMunicipalities = [
    "ALDINO",
    "ANDRIANO",
    "ANTERIVO",
    "APPIANO SULLA STRADA DEL VINO",
    "AVELENGO",
    "BADIA",
    "BARBIANO",
    "BOLZANO",
    "BRAIES",
    "BRENNERO",
    "BRESSANONE",
    "BRONZOLO",
    "BRUNICO",
    "CAINES",
    "CALDARO SULLA STRADA DEL VINO",
    "CAMPO TURES",
    "CAMPO DI TRENS",
    "CASTELBELLO-CIARDES",
    "CASTELROTTO",
    "CERMES",
    "CHIENES",
    "CHIUSA",
    "CORNEDO ALL'ISARCO",
    "CORTACCIA SULLA STRADA DEL VINO",
    "CORTINA SULLA STRADA DEL VINO",
    "CORVARA IN BADIA",
    "CURON VENOSTA",
    "DOBBIACO",
    "EGNA",
    "FALZES",
    "FIE' ALLO SCILIAR",
    "FORTEZZA",
    "FUNES",
    "GAIS",
    "GARGAZZONE",
    "GLORENZA",
    "LA VALLE",
    "LACES",
    "LAGUNDO",
    "LAION",
    "LAIVES",
    "LANA",
    "LASA",
    "LAUREGNO",
    "LUSON",
    "MAGRE' SULLA STRADA DEL VINO",
    "MALLES VENOSTA",
    "MAREBBE",
    "MARLENGO",
    "MARTELLO",
    "MELTINA",
    "MERANO",
    "MONGUELFO - TESIDO",
    "MONTAGNA",
    "MOSO IN PASSIRIA",
    "NALLES",
    "NATURNO",
    "NAZ-SCIAVES",
    "NOVA LEVANTE",
    "NOVA PONENTE",
    "ORA",
    "ORTISEI",
    "PARCINES",
    "PERCA",
    "PLAUS",
    "PONTE GARDENA",
    "POSTAL",
    "PRATO ALLO STELVIO",
    "PREDOI",
    "PROVES",
    "RACINES",
    "RASUN ANTERSELVA",
    "RENON",
    "RIFIANO",
    "RIO DI PUSTERIA",
    "RODENGO",
    "SALORNO SULLA STRADA DEL VINO",
    "SAN CANDIDO",
    "SAN GENESIO ATESINO",
    "SAN LEONARDO IN PASSIRIA",
    "SAN LORENZO DI SEBATO",
    "SAN MARTINO IN BADIA",
    "SAN MARTINO IN PASSIRIA",
    "SAN PANCRAZIO",
    "SANTA CRISTINA VALGARDENA",
    "SARENTINO",
    "SCENA",
    "SELVA DEI MOLINI",
    "SELVA DI VAL GARDENA",
    "SENALE-SAN FELICE",
    "SENALES",
    "SESTO",
    "SILANDRO",
    "SLUDERNO",
    "STELVIO",
    "TERENTO",
    "TERLANO",
    "TERMENO SULLA STRADA DEL VINO",
    "TESIMO",
    "TIRES",
    "TIROLO",
    "TRODENA NEL PARCO NATURALE",
    "TUBRE",
    "ULTIMO",
    "VADENA",
    "VAL DI VIZZE",
    "VALDAORA",
    "VALLE AURINA",
    "VALLE DI CASIES",
    "VANDOIES",
    "VARNA",
    "VELTURNO",
    "VERANO",
    "VILLABASSA",
    "VILLANDRO",
    "VIPITENO",
];

function main() {
    fetch(url)
        .then(function (res) {
            if (!res.ok) throw new Error("fetch failed");
            return res.arrayBuffer();
        })
        .then(function (ab) {
            // get Excel Sheet informations
            const data = new Uint8Array(ab);
            const workbook = XLSX.read(data, {
                type: "array",
            });
            const sheetContent = workbook.Sheets[workbook.SheetNames[0]];
            const range = XLSX.utils.decode_range(sheetContent["!ref"]);
            // const row_start = range.s.r;
            // const row_end = range.e.r;

            const covid_data = [];

            // Newest date in sheet (get it it from cell E3)
            let dt = sheetContent.E3.v.replace("Totali al ", "");

            // Loop through 1000 rows
            for (let i = 0; i < 1000; i++) {
                let cellMunicipality = "B" + i,
                    cellTotalPositivesToday = "E" + i,
                    cellTotalPositivesYesterday = "D" + i,
                    cellTotalCuredToday = "H" + i,
                    cellTotalCuredYesterday = "G" + i,
                    cellDeceased = "J" + i,
                    cellActivePositives = "K" + i;
                cellMunicipalityUnknownToday = "F" + i;
                cellTotalPositivesOfAllMunicipalitiesToday = "F" + i;
                cellTotalPositivesOfAllMunicipalitiesUntilToday = "E" + i;
                cellTotalCuredUntilToday = "H" + i;
                cellTotalDeceasedUntilToday = "J" + i;
                cellTotalActivePositivesUntilToday = "K" + i;

                if (sheetContent[cellMunicipality] !== undefined) {
                    // Get rows which contain the string "Comune sconosciuto Totale"
                    if (sheetContent[cellMunicipality].v.includes("Comune sconosciuto Totale")) {
                        covid_data.push({
                            municipalityUnknownToday: sheetContent[cellMunicipalityUnknownToday].v,
                        });
                    }
                    // Get rows which contain the string "Totale complessivo"
                    if (sheetContent[cellMunicipality].v.includes("Totale complessivo")) {
                        covid_data.push({
                            totalSum: {
                                positivesUntilToday: sheetContent[cellTotalPositivesOfAllMunicipalitiesUntilToday].v,
                                positivesToday: sheetContent[cellTotalPositivesOfAllMunicipalitiesToday].v,
                                curedUntilToday: sheetContent[cellTotalCuredUntilToday].v,
                                deceasedUntilToday: sheetContent[cellTotalDeceasedUntilToday].v,
                                activePostitivesUntilToday: sheetContent[cellTotalActivePositivesUntilToday].v,
                            },
                        });
                    }
                    // Get rows which contain the string "Totale"
                    if (sheetContent[cellMunicipality].v.includes("Totale")) {
                        if (
                            //only add municipalities of South Tyrol
                            listOfMunicipalities.includes(sheetContent[cellMunicipality].v.replace(" Totale", ""))
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
                                deceased: sheetContent[cellDeceased] !== undefined && sheetContent[cellDeceased].v,
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
