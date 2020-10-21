const fetch = require("node-fetch");
const XLSX = require("xlsx");
const fse = require("fs-extra");
const puppeteer = require("puppeteer");

// Scrape data from this URL
// URL has to be changed manually every day
const pressPostUrl = "https://www.sabes.it/de/news.asp?aktuelles_action=4&aktuelles_article_id=644669";

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

async function main() {
    let xlsxUrl = "";
    let hospitalNumbers = {};

    async function scrapePressPost(pressPostUrl) {
        const browser = await puppeteer.launch();
        const page = await browser.newPage();
        await page.goto(pressPostUrl);

        async function getXlsxUrl() {
            const [el] = await page.$x('//*[@id="content"]/div[2]/div/div[1]/ol/li[2]/a');
            const href = await el.getProperty("href");
            const hrefText = await href.jsonValue();
            return hrefText;
        }

        async function getInHospitalNumbers() {
            const paragraphs = await page.$$("p");

            const inHospital = { hospitalNumbers };
            for (var i = 0; i < paragraphs.length; i++) {
                let valueHandle = await paragraphs[i].getProperty("innerText");
                let paragraphText = await valueHandle.jsonValue();

                if (paragraphText.includes("Auf Normalstationen im Krankenhaus")) {
                    inHospital.hospitalNumbers.normalBed = Number(paragraphText.split(":").pop());
                }

                if (paragraphText.includes("in Intensivbetreuung")) {
                    inHospital.hospitalNumbers.intensiveCare = Number(paragraphText.split(":").pop());
                }

                if (paragraphText.includes("In Gossensaß")) {
                    inHospital.hospitalNumbers.gossensass = Number(paragraphText.split(":").pop());
                }
            }
            return inHospital;
        }

        hospitalNumbers = await getInHospitalNumbers();

        xlsxUrl = await getXlsxUrl();

        browser.close();
    }

    await scrapePressPost(pressPostUrl);

    fetch(xlsxUrl)
        .then(async function (res) {
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
            let dt = sheetContent.F3.v.replace("Gesamt - Totale", "");
            
            // Loop through 1300 rows
            for (let i = 0; i < 1300; i++) {
                let cellMunicipality = "C" + i,
                    cellTotalPositivesToday = "F" + i,
                    cellTotalPositivesYesterday = "E" + i,
                    cellActivePositives = "H" + i;
                cellMunicipalityUnknownToday = "G" + i;
                cellTotalPositivesOfAllMunicipalitiesToday = "G" + i;
                cellTotalPositivesOfAllMunicipalitiesUntilToday = "F" + i;
                cellTotalActivePositivesUntilToday = "H" + i;

                if (sheetContent[cellMunicipality] !== undefined) {
                    // Get rows which contain the string "Comune sconosciuto Totale"
                    if (sheetContent[cellMunicipality].v.includes("Comune sconosciuto Totale")) {
                        covid_data.push({
                            municipalityUnknownToday:
                                sheetContent[cellMunicipalityUnknownToday] === undefined
                                    ? 0
                                    : sheetContent[cellMunicipalityUnknownToday].v,
                        });
                    }
                    // Get rows which contain the string "Totale complessivo"
                    if (sheetContent[cellMunicipality].v.includes("Totale complessivo")) {
                        covid_data.push({
                            totalSum: {
                                positivesUntilToday:
                                    sheetContent[cellTotalPositivesOfAllMunicipalitiesUntilToday] !== undefined &&
                                    sheetContent[cellTotalPositivesOfAllMunicipalitiesUntilToday].v,
                                positivesToday:
                                    sheetContent[cellTotalPositivesOfAllMunicipalitiesToday] !== undefined &&
                                    sheetContent[cellTotalPositivesOfAllMunicipalitiesToday].v,
                                // curedUntilToday:
                                //     // sheetContent[cellTotalCuredUntilToday] !== undefined &&
                                //     sheetContent[cellTotalCuredUntilToday].v,
                                // deceasedUntilToday:
                                //     sheetContent[cellTotalDeceasedUntilToday] !== undefined &&
                                //     sheetContent[cellTotalDeceasedUntilToday].v,
                                activePostitivesUntilToday:
                                    sheetContent[cellTotalActivePositivesUntilToday] !== undefined &&
                                    sheetContent[cellTotalActivePositivesUntilToday].v,
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
                                // totalCuredToday:
                                //     sheetContent[cellTotalCuredToday] !== undefined &&
                                //     sheetContent[cellTotalCuredToday].v,
                                // totalCuredYesterday: sheetContent[cellTotalCuredYesterday].v,
                                // deceased: sheetContent[cellDeceased] !== undefined && sheetContent[cellDeceased].v,
                                activePositives:
                                    sheetContent[cellActivePositives] !== undefined &&
                                    sheetContent[cellActivePositives].v,
                            });
                        }
                    }
                }
            }

            covid_data.push(hospitalNumbers);

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
