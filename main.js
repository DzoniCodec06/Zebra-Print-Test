const fs = require("fs");
//const pdfPrinter = require("pdf-to-printer");
//const doxcConvert = require("docx-pdf");
const docx = require("docx");
const { exec } = require("child_process");
const path = require("path");
const { createCanvas } = require("canvas");
const jsBarcode = require("jsbarcode");


const canvas = createCanvas(113, 30);

let barcodeValue = "1234567890";

jsBarcode(canvas, barcodeValue, {
    format: 'CODE128', // Barcode format
    width: 2, // Width of each bar
    height: 30, // Height of the barcode
    displayValue: true, // Show the value below the barcode
    fontSize: 15,
});

let barcodeImg = canvas.toDataURL("image/png");

let img = new docx.ImageRun({
    type: "png",
    data: fs.readFileSync("./images/example-2.png"),
    transformation: {
        width: 113,
        height: 43
    }
})

let barCodeImg = new docx.ImageRun({
    type: "png",
    data: barcodeImg,
    transformation: {
        width: 113,
        height: 30
    }
})

let skuNumbers = 1000;

let doc = new docx.Document({
    sections: [
        {
            properties: {page: {size: {width: "3cm", height: "2cm"}, margin: {left: "0cm", right: "0cm", top: "0cm", bottom: "0cm"}}},
            children: [
                new docx.Paragraph({
                    children: [
                        img,
                        barCodeImg,
                    ]
                }),
            ],
        },
    ],
});

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

docx.Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("output2.docx", buffer);
    console.log("DOCX file created successfully!");
});


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
