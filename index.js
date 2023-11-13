/* 
GET DATA LIKE:
- Computer Name
- Operating System
- System Model
- Processor
- Main Board
- Device Name
- RAM
- IP
- MAC


FROM Belarc Advisor HTML FILE TO EXCEL FILE

create a folder called files and put all the html files in it
RUN: npm install
EXECUTE: node index.js
*/
const fs = require('fs');
const { JSDOM } = require('jsdom');
const ExcelJS = require('exceljs');
const path = require('path');
const folderPath = 'files';

const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Inventario');

fs.readdir(folderPath, (err, files) => {
  files.forEach(file => {
    const filePath = path.join(folderPath, file);
    console.log(filePath);

    let html = fs.readFileSync(filePath, 'utf-8');
    // Read the HTML file
    /*const html = fs.readFileSync('files/GestiondeRiezgos-AuditoriaInterna-IsraelValdiviesoSalinas-AuditorInterno.html', 'utf-8');*/
    
    // Load the HTML into JSDOM
    const dom = new JSDOM(html);
    
    // Get the title element and extract the text
    const title = dom.window.document.querySelector('h1').textContent;
    
    // Log the title to the console
    console.log(title);
    
    
    /* Computer Name */
    const rows = dom.window.document.querySelectorAll('.reportHeader tbody tr');
    
    // Loop through the rows to find the one with "Computer Name"
    let computerNameOnly = null;
    rows.forEach(function(row) {
        const thText = row.querySelector('th').textContent.trim();
        if (thText === 'Computer Name:') {
            // Extract the computer name from the corresponding td
            const computerName = row.querySelector('td').textContent.trim();
    
            // Extract only the computer name without additional information
            computerNameOnly = computerName.split(' ')[0];
    
            // Log the computer name to the console
            console.log('Computer Name:', computerNameOnly);
        }
    });
    
    
    /* Operating System */
    
    const reportSectionBody = dom.window.document.querySelector('.reportSection .reportSectionBody');
    const osInfo = reportSectionBody.textContent.trim();
    console.log('Operating System Information:', osInfo.substring(0, osInfo.indexOf('Version')));
    
    /* System Model */
    
    const systemDiv = dom.window.document.querySelectorAll('.reportSection .reportSectionBody')[1]
    const PCModel = systemDiv.innerHTML.split('<br>')[0].replace(/[\n\t]/g, '');
    const PCSerialNumber = systemDiv.innerHTML.split('<br>')[1].split(':')[1].replace(/[\n\t]/g, '');
    console.log('PC Model:', PCModel)
    console.log('PC Serial Number:', PCSerialNumber)
    
    /* get the processor specification from html */
    const processorDiv = dom.window.document.querySelectorAll('.reportSection .reportSectionBody')[2];
    const processorInfo = processorDiv.textContent.trim().substring(0, 37);/* 37 could be changed */
    console.log(processorInfo);
    
    
    /* get the main board specification from html */
    
    const mainBoardDiv = dom.window.document.querySelectorAll('.reportSection .reportSectionBody')[3]
    const mainBoardInfo = mainBoardDiv.textContent.trim().substring(0, mainBoardDiv.textContent.trim().indexOf('Serial Number')).split(':')[1];
    console.log(mainBoardInfo);
    
    /*  */
    // Select the table row that contains the device information
    const deviceRow = dom.window.document.querySelectorAll('tr .hasInfo');
    
    // Select the table cell that contains the device name
    const deviceNameCell1 = deviceRow[0].parentNode.textContent.trim();
    let deviceNameCell2 = null;
    
    const StorageTable = dom.window.document.querySelectorAll('table')[2];
    if(StorageTable.rows.length > 4){
      deviceNameCell2 = deviceRow[1].textContent.trim();
    }
    
    // Log the device name to the console
    console.log(deviceNameCell1);
    console.log(deviceNameCell2);
    
    /* get RAM from PC */
    const RAM = dom.window.document.querySelectorAll('.reportSection.rsRight .reportSectionBody')[2].textContent.trim();
    console.log(RAM);
    
    /* get IP and MAC */
    
    const spanList = Array.from(dom.window.document.querySelectorAll('span'));
    const primary = spanList.find(span => {
      if(span.textContent == 'primary' && span.parentNode.parentNode.textContent.includes('IPv4'))
        return true;
      else
        return false;
    });
    
    const IP = primary.parentNode.parentNode.textContent
    
    console.log(IP.split(':')[1]);

    let MAC = null;
    if(primary.parentNode.parentNode.nextElementSibling.nextElementSibling.textContent.includes('Physical')){
      MAC = primary.parentNode.parentNode.nextElementSibling.nextElementSibling.textContent
    } else {
      MAC = primary.parentNode.parentNode.nextElementSibling.nextElementSibling.nextElementSibling.textContent
    }

    console.log(MAC.split('Address:')[1]);

    /* add row to worksheet */
    worksheet.addRow([
      file,computerNameOnly, osInfo.substring(0, osInfo.indexOf('Version')),
      PCModel, PCSerialNumber, processorInfo, mainBoardInfo, deviceNameCell1, deviceNameCell2,
      RAM, IP.split(':')[1], MAC.split('Address:')[1]]);
  });
});

setTimeout(() =>{
  workbook.xlsx.writeFile('Inventario.xlsx')
  .then(function() {
    console.log('File saved.');
  })
  .catch(function(error) {
    console.log(error);
  });

}, 5000)