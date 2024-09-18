const fs = require("fs");
//const pdfPrinter = require("pdf-to-printer");
//const doxcConvert = require("docx-pdf");
const docx = require("docx");
const { exec } = require("child_process");
const path = require("path");
const { createCanvas } = require("canvas");
const jsBarcode = require("jsbarcode");

const { spawn } = require("child_process");

let command = `Start-Process "C:\\Users\\HP\\OneDrive\\Desktop\\Zebra-Print-Test-Beta\\output.docx" -Verb print`;

spawn("powershell.exe", [command]);

const readline = require('readline');

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

const canvas = createCanvas(113, 30);

let img = new docx.ImageRun({
    type: "png",
    data: fs.readFileSync("./images/example-2.png"),
    transformation: {
        width: 113,
        height: 43
    }
})


let skuNumbers = 1000;

let inputting = true;

const generateBarcode = async () => {
    let value = "";
    for (let i = 0; i < 12; i++) {
        value += Math.floor(Math.random() * 9)
    } 

    console.log(value);
    return value;
}

// A function to wrap rl.question in a Promise
const askQuestion = (question) => {
    return new Promise((resolve) => {
        rl.question(question, resolve);
    });
};

(async () => {
    let SKU1 = await askQuestion('Enter SKU1: ');
    let SKU2 = await askQuestion('Enter SKU2: ');
    let SKU3 = await askQuestion('Enter SKU3: ');

    if (SKU1 != "" && SKU2 != "" && SKU3 != "") {
        let barcodeValue = await generateBarcode();

        jsBarcode(canvas, barcodeValue, {
            format: 'CODE39', // Barcode format
            width: 2, // Width of each bar
            height: 30, // Height of the barcode
            displayValue: true, // Show the value below the barcode
            fontSize: 15,
        });

        let barcodeImg = canvas.toDataURL("image/png");

        let barCodeImg = new docx.ImageRun({
            type: "png",
            data: barcodeImg,
            transformation: {
                width: 113,
                height: 30
            }
        })
        

        let doc = new docx.Document({
            sections: [
                {
                    properties: {page: {size: {width: "3cm", height: "2cm"}, margin: {left: "0cm", right: "0cm", top: "0cm", bottom: "0cm"}}},
                    children: [
                        new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            border: {
                                bottom: {
                                    style: docx.BorderStyle.SINGLE,
                                    size: 3,
                                    color: "000000" 
                                }
                            },
                            children: [
                                new docx.TextRun({
                                    text: "S K U",
                                    size: 20,
                                    bold: true,
                                    font: "Calibri (Body)"
                                }),
                            ]
                        }),
                        new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                                new docx.TextRun({
                                    text: "1                           2                           3",
                                    size: 10,
                                    bold: true,
                                    font: "Calibri (Body)",
                                })
                            ]
                        }),
                        new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            border: {
                                top: {
                                    style: docx.BorderStyle.SINGLE,
                                    size: 3,
                                    color: "000000"
                                },
                                bottom: {
                                    style: docx.BorderStyle.SINGLE,
                                    size: 3,
                                    color: "000000" 
                                }
                            },
                            children: [
                                new docx.TextRun({
                                    text: `${SKU1}  |  ${SKU2}  |  ${SKU3}`,
                                    size: 17,
                                    bold: true,
                                    font: "Calibri (Body)",
                                })
                            ]
                        }),
                        new docx.Paragraph({
                            children: [
                                barCodeImg
                            ]
                        })
                    ],
                },
            ],
        });

        docx.Packer.toBuffer(doc).then((buffer) => {
            fs.writeFileSync("output2.docx", buffer);
            console.log("DOCX file created successfully!");
        });

        inputting = false;

        if (inputting == false) {
            spawn("powershell.exe", [command]);
        }
    }
    rl.close();
})();


/*
new docx.Paragraph({
                    
                    frame: {
                        position: {
                            x: 100,
                            y: 200,

                        },
                        width: 200,
                        height: 20,
                        anchor: {
                            horizontal: docx.FrameAnchorType.MARGIN,
                            vertical: docx.FrameAnchorType.MARGIN,
                        },
                        alignment: {
                            x: docx.HorizontalPositionAlign.CENTER,
                            y: docx.VerticalPositionAlign.TOP,
                        },
                    },
                    children: [
                        new docx.TextRun({
                            text: "1000",
                            size: 10
                        }),
                        new docx.InsertedTextRun({
                            text: "1000",
                            size: 10,
                            position: 100
                        })
                    ]
                })
*/

/*
new Paragraph({
                    children: [
                        new TextRun("Hello, this is a docx file created using Node.js!"),
                        new TextRun({
                            text: " This is another sentence.",
                            bold: true,
                            size: "3px",
                        })
                    ],
                }),
*/



let currentFile = "output.docx";
let outputFile = "zebest.pdf";
/*
setTimeout(() => {
    let command = `"C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE" /mFilePrintDefault /q "${currentFile}"`;

    exec(command, (err) => {
        if (err) console.log(err);
        else {
            console.log("File sent to printer!");
        }
    })
}, 1000);*/

/*
doxcConvert(currentFile, outputFile, (err) => {
    if (err) console.log(err);
    else {
        console.log("Successfull converted");

        pdfPrinter.print(outputFile).then((err, jobID) => {
            if (err) console.log(err);
            console.log(jobID);
        })
    }
})*/
/*
fs.writeFile("zebest.docx", "ARTIKLI", (err) => {
    if (err) console.error(err);
    else {
        console.log("File made");

        fs.close(0, err => {
            if (err) console.error(err);
            else console.log("FILE CLOSED");
        })      
        
        fs.print();
    }
})*/
