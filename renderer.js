const docx = require("docx");
const fs = require("fs");
const { AlignmentType, Document, Packer, Paragraph, TextRun, ImageRun } = docx;
const { saveAs } = require("file-saver");
const XLSX = require("xlsx");
const Chart = require("chart.js");
const JSZip = require("jszip");
const $ = (selector) => document.querySelector(selector);

const $button = $("#wordButton");

//Eventos

document.getElementById("fileInput").addEventListener("change", function () {
  const fileName = document.getElementById("fileInput").files[0].name;
  document.getElementById("fileName").textContent = fileName;
});

$button.addEventListener("click", async () => {
  const file = document.getElementById("fileInput").files[0];
  if (!file) {
    alert("Please select an Excel file.");
    return;
  }

  const documents = await handleFile(file);
  if (documents) {
    const zipBlob = await createZipWithDocuments(documents);
    saveAs(zipBlob, "evaluaciones.zip");
  }
});

//Funciones generales

const handleFile = async (file) => {
  try {
    const data = await readExcelFile(file);
    const groupedData = groupByLeaderEmail(data);

    const documents = {};
    for (const [leaderEmail, leaderData] of Object.entries(groupedData)) {
      const leaderName = leaderData[0]["Nombre líder evaluado"];
      const fileName = `Evaluación_${leaderEmail}.docx`;
      const docBlob = await generateWordForLeader(leaderData, leaderName, leaderEmail);
      documents[fileName] = docBlob;
    }

    return documents;
  } catch (error) {
    console.error("Error processing the file:", error);
    alert("Error processing the file. Please check the console for more details.");
  }
};

const readExcelFile = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        let jsonData = XLSX.utils.sheet_to_json(worksheet);

        jsonData = jsonData.map((entry) => {
          const cleanedEntry = {};
          for (const key in entry) {
            const cleanedKey = key.replace(/\s*↵$/, "").trim();
            cleanedEntry[cleanedKey] = entry[key];
          }
          return cleanedEntry;
        });

        resolve(jsonData);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

// Generar varios documentos de Word

const groupByLeaderEmail = (data) => {
  return data.reduce((acc, item) => {
    const leaderEmail = item["Correo líder evaluado"];
    if (!acc[leaderEmail]) {
      acc[leaderEmail] = [];
    }
    acc[leaderEmail].push(item);
    return acc;
  }, {});
};

const processExcelDataMultiWord = (xlsxData) => {
  const question1Responses = {
    Diaria: 0,
    Semanal: 0,
    Quincenal: 0,
    Mensual: 0,
    Semestral: 0,
    Anual: 0,
  };

  const question2Responses = {
    "Más de 4 veces al año": 0,
    "2 a 3 veces al año": 0,
    "1 vez al año": 0,
    "Nunca he recibido feedback": 0,
  };

  const question3Responses = {
    "Totalmente insatisfecho": 0,
    Insatisfecho: 0,
    "Ni satisfecho ni insatisfecho": 0,
    Satisfecho: 0,
    "Totalmente satisfecho": 0,
  };

  xlsxData.forEach((response) => {
    question1Responses[
      response[
        "¿Con qué frecuencia mantienes comunicación con tu líder directo?"
      ]
    ]++;
    question2Responses[
      response[
        "¿Con qué frecuencia recibes feedback por parte de tu líder directo?"
      ]
    ]++;
    question3Responses[
      response[
        "Del 1 al 5, ¿Qué tan satisfecho(a) te encuentras con la gestión de tu líder directo?"
      ]
    ]++;
  });

  return { question1Responses, question2Responses, question3Responses };
};

const generateChartMultiWord = (data, question, labels) => {
  const ctx = document.createElement("canvas");
  document.body.appendChild(ctx);

  return new Promise((resolve, reject) => {
    const myChart = new Chart(ctx, {
      type: "bar",
      data: {
        labels: labels,
        datasets: [
          {
            label: question,
            data: Object.values(data),
            backgroundColor: [
              "rgb(255, 86, 77, 0.2)",
              "rgb(250, 190, 0, 0.2)",
              "rgb(64, 193, 239, 0.2)",
              "rgb(68, 25, 126, 0.2)",
              "rgb(143, 199, 69, 0.2)",
            ],
            borderColor: [
              "rgb(255, 86, 77)",
              "rgb(250, 190, 0)",
              "rgb(64, 193, 239)",
              "rgb(68, 25, 126)",
              "rgb(143, 199, 69)",
            ],
            borderWidth: 1,
          },
        ],
      },
      options: {
        scales: {
          y: {
            beginAtZero: true,
            ticks: {
              stepSize: 1,
              callback: function (value) {
                return Number.isInteger(value) ? value : null;
              },
            },
          },
        },
      },
    });
    myChart.update();

    setTimeout(() => {
      ctx.toBlob((blob) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          resolve(reader.result);
          document.body.removeChild(ctx);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(blob);
      }, "image/png");
    }, 1000);
  });
};

const generatePieChartMultiWord = (data, question, labels) => {
  const ctx = document.createElement("canvas");
  document.body.appendChild(ctx);

  return new Promise((resolve, reject) => {
    const myChart = new Chart(ctx, {
      type: "pie",
      data: {
        labels: labels,
        datasets: [
          {
            label: question,
            data: Object.values(data),
            backgroundColor: [
              "rgb(255, 86, 77)",
              "rgb(250, 190, 0)",
              "rgb(64, 193, 239)",
              "rgb(68, 25, 126)",
              "rgb(143, 199, 69)",
            ],
            borderWidth: 1,
          },
        ],
      },
      options: {
        scales: {
          y: {
            beginAtZero: true,
          },
        },
      },
    });
    myChart.update();

    setTimeout(() => {
      ctx.toBlob((blob) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          resolve(reader.result);
          document.body.removeChild(ctx);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(blob);
      }, "image/png");
    }, 1000);
  });
};

const generateWordForLeader = async (leaderData, leaderName, leaderEmail) => {
  try {
    const { question1Responses, question2Responses, question3Responses } =
      processExcelDataMultiWord(leaderData);
    const additionalComments = leaderData
      .filter((entry) => entry["Comentarios adicionales:"])
      .map((entry) => entry["Comentarios adicionales:"]);
    const imageBuffer1 = await generateChartMultiWord(
      question1Responses,
      "¿Con qué frecuencia mantienes comunicación con tu líder directo?",
      ["Diaria", "Semanal", "Quincenal", "Mensual", "Semestral", "Anual"]
    );
    const imageBuffer2 = await generateChartMultiWord(
      question2Responses,
      "¿Con qué frecuencia recibes feedback por parte de tu líder directo?",
      [
        "1 vez al año",
        "2 a 3 veces al año",
        "Más de 4 veces al año",
        "Nunca he recibido feedback",
      ]
    );
    const imageBuffer3 = await generatePieChartMultiWord(
      question3Responses,
      "Del 1 al 5, ¿Qué tan satisfecho(a) te encuentras con la gestión de tu líder directo?",
      [
        "Totalmente insatisfecho",
        "Insatisfecho",
        "Ni satisfecho ni insatisfecho",
        "Satisfecho",
        "Totalmente satisfecho",
      ]
    );
    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "Resultados de Evaluación de Liderazgo",
                  bold: true,
                  size: 44,
                  font: "Montserrat",
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new ImageRun({
                  data: fs.readFileSync("./encora-logo.jpg"),
                  transformation: {
                    width: 300,
                    height: 150,
                  },
                }),
              ],
              spacing: {
                after: 5000,
              },
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "Líder Evaluado:",
                  size: 38,
                  bold: true,
                  font: "Montserrat",
                  color: "ff564d",
                }),
              ],
              spacing: {
                after: 200,
              },
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: leaderName,
                  size: 38,
                  font: "Montserrat",
                }),
              ],
              spacing: {
                after: 2000,
              },
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "Correo:",
                  size: 38,
                  bold: true,
                  font: "Montserrat",
                  color: "8fc745",
                }),
              ],
              spacing: {
                after: 200,
              },
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: leaderEmail,
                  size: 38,
                  font: "Montserrat",
                }),
              ],
              spacing: {
                after: 2000,
              },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "¿Con qué frecuencia mantienes comunicación con tu líder directo?",
                  size: 25,
                  bold: true,
                  font: "Montserrat",
                }),
              ],
              spacing: {
                after: 300,
              },
            }),
            new Paragraph({
              children: [
                new ImageRun({
                  data: imageBuffer1,
                  transformation: {
                    width: 600,
                    height: 300,
                  },
                }),
              ],
              spacing: {
                after: 1000,
              },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "¿Con qué frecuencia recibes feedback por parte de tu líder directo?",
                  size: 25,
                  bold: true,
                  font: "Montserrat",
                }),
              ],
              spacing: {
                after: 300,
              },
            }),
            new Paragraph({
              children: [
                new ImageRun({
                  data: imageBuffer2,
                  transformation: {
                    width: 600,
                    height: 300,
                  },
                }),
              ],
              spacing: {
                after: 2500,
              },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Del 1 al 5, ¿Qué tan satisfecho(a) te encuentras con la gestión de tu líder directo?",
                  size: 25,
                  bold: true,
                  font: "Montserrat",
                }),
              ],
              spacing: {
                after: 300,
              },
            }),
            new Paragraph({
              children: [
                new ImageRun({
                  data: imageBuffer3,
                  transformation: {
                    width: 600,
                    height: 600,
                  },
                }),
              ],
              spacing: {
                after: 2000,
              },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Comentarios adicionales:",
                  size: 25,
                  bold: true,
                  font: "Montserrat",
                }),
              ],
              spacing: {
                after: 300,
              },
            }),
            ...additionalComments.map(
              (comment) =>
                new Paragraph({
                  children: [
                    new TextRun({
                      text: comment,
                      size: 20,
                      font: "Montserrat",
                    }),
                  ],
                  spacing: {
                    after: 200,
                  },
                })
            ),
          ],
        },
      ],
    });

    return Packer.toBlob(doc);
  } catch (error) {
    console.error("Error generating the Word document:", error);
  }
};

// Guardar y generar archivo .Zip

const createZipWithDocuments = async (documents) => {
  const zip = new JSZip();
  for (const [fileName, fileBlob] of Object.entries(documents)) {
    zip.file(fileName, fileBlob);
  }
  return zip.generateAsync({ type: "blob" });
};

