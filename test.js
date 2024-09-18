//let command = `"C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE" /mFilePrintDefault /q "${currentFile}"`;
const { spawn } = require("child_process");

let command = `Start-Process "G:\\Programiranje\\JavaScript\\Zebra-Print-Test-main\\output.docx" -Verb print`;

spawn("powershell.exe", [command]);
