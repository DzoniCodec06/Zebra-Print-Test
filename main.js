const fs = require("fs");
const pdfPrinter = require("pdf-to-printer");
const doxcConvert = require("docx-pdf");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const { exec } = require("child_process");

let doc = new Document({
    sections: [
        {
            properties: {page: {size: {width: "3cm", height: "2cm"}, margin: {left: "0.2cm", right: "0cm", top: "0cm", bottom: "0cm"}}},
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello, this is a docx file created using Node.js!"),
                        new TextRun({
                            text: " This is another sentence.",
                            bold: true,
                            size: "3px"
                        })
                    ],
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("output.docx", buffer);
    console.log("DOCX file created successfully!");
});


let currentFile = "output.docx";
let outputFile = "zebest.pdf";

setTimeout(() => {
    let command = `"C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE" /mFilePrintDefault /q "${currentFile}"`;

    exec(command, (err) => {
        if (err) console.log(err);
        else {
            console.log("File sent to printer!");
        }
    })
}, 1000);

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
