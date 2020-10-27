import fs from "fs/promises";

const outputFolder = "./output/";
const outputFileExtension = ".json";
const fileWithIstatCodes = JSON.parse(await fs.readFile(`${outputFolder}27-10-2020${outputFileExtension}`, "utf-8"));

const outFileNames = [
  "25-10-2020",
  "24-10-2020",
  "23-10-2020",
  "22-10-2020",
  "21-10-2020",
  "20-10-2020",
  "19-10-2020",
  "18-10-2020",
  "17-10-2020",
  "16-10-2020",
  "15-10-2020",
  "14-10-2020",
  "13-10-2020",
  "12-10-2020",
  "11-10-2020",
  "10-10-2020",
  "9-10-2020",
  "8-10-2020",
  "7-10-2020",
  "6-10-2020",
  "5-10-2020",
  "4-10-2020",
  "3-10-2020",
  "2-10-2020",
  "1-10-2020",
  "30-9-2020",
  "29-9-2020",
  "28-9-2020",
  "27-9-2020",
  "26-9-2020",
  "25-9-2020",
  "24-9-2020",
  "23-9-2020",
  "22-9-2020",
  "21-9-2020",
  "20-9-2020",
  "19-9-2020",
  "18-9-2020",
  "17-9-2020",
  "16-9-2020",
  "15-9-2020",
  "14-9-2020",
  "13-9-2020",
  "12-9-2020",
  "11-9-2020",
  "10-9-2020",
  "9-9-2020",
  "8-9-2020",
  "7-9-2020",
  "6-9-2020",
  "5-9-2020",
  "4-9-2020",
  "3-9-2020",
  "2-9-2020",
  "1-9-2020",
  "31-8-2020",
  "30-8-2020",
  "29-8-2020",
  "28-8-2020",
  "27-8-2020",
  "26-8-2020",
  "25-8-2020",
  "24-8-2020",
  "23-8-2020",
  "22-8-2020",
  "21-8-2020",
  "20-8-2020",
  "19-8-2020",
  "18-8-2020",
  "17-8-2020",
  "16-8-2020",
  "15-8-2020",
  "14-8-2020",
  "13-8-2020",
  "12-8-2020",
  "11-8-2020"
];

for (const fileName of outFileNames) {
  const filePath = `${outputFolder}${fileName}${outputFileExtension}`;
  const file = JSON.parse(await fs.readFile(filePath, "utf-8"))

  for (const municipality of file) {
    if (municipality.municipality) {
      const istatCode = fileWithIstatCodes.find((data) => data.municipality === municipality.municipality).municipalityIstatCode;

      municipality.municipalityIstatCode = istatCode;
    }
  }

  await fs.writeFile(filePath, JSON.stringify(file), "utf-8");
}