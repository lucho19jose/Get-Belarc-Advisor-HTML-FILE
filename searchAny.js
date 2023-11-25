/* const fs = require('fs');

// Function to search for a specific word in an HTML file
function searchWordInHTMLFile(filePath, word) {
    try {
        // Read the HTML file content
        const content = fs.readFileSync(filePath, 'utf8');

        // Check if the word exists in the content
        if (content.includes(word)) {
            console.log(`The word "${word}" was found in the HTML file.`);
        } else {
            console.log(`The word "${word}" was not found in the HTML file.`);
        }
    } catch (error) {
        console.error('An error occurred while searching for the word:', error);
    }
}
 */
// Usage example
/* const filePath = '/path/to/html/file.html'; // Replace with the path to your HTML file
const word = 'example'; // Replace with the word you want to search for
searchWordInHTMLFile(filePath, word);
 */
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

worksheet.addRow([
  'Usuario',
  'Cargo',
  'Estado',
  'Hostname',
  'Visor'
]);
const word = 'Adobe Systems - Acrobat';//TrueView 
const secondWord = 'Adobe Acrobat Reader';
fs.readdir(folderPath, (err, files) => {
  files.forEach(file => {
    const filePath = path.join(folderPath, file);
    console.log(filePath);

    let html = fs.readFileSync(filePath, 'utf-8');
    // Read the HTML file
    /*const html = fs.readFileSync('files/GestiondeRiezgos-AuditoriaInterna-IsraelValdiviesoSalinas-AuditorInterno.html', 'utf-8');*/
    if (html.includes(word)) {
        console.log(`The word "${word}" was found in the HTML file.`);
        const dom = new JSDOM(html);

        /* Computer Name */
        const rows = dom.window.document.querySelectorAll('.reportHeader tbody tr');
        
        // Loop through the rows to find the one with "Computer Name"
        let computerNameOnly = null;
        rows.forEach(function(row) {
            const thText = row.querySelector('th').textContent.trim();
            if (thText === 'Computer Name:' || thText === 'Nombre de la computadora:') {
                // Extract the computer name from the corresponding td
                const computerName = row.querySelector('td').textContent.trim();
        
                // Extract only the computer name without additional information
                computerNameOnly = computerName.split(' ')[0];
        
                // Log the computer name to the console
                console.log('Computer Name:', computerNameOnly);
            }
        });
        let viewer = false;
        if(html.includes(secondWord)){
          viewer = true;
        }
        
        worksheet.addRow([
            file.split('-').length > 3 ? file.split('-')[2].trim() : file.split('-')[1].trim(),
            file.split('-').length > 3 ? (file.split('-')[3]).split('.')[0].trim() : (file.split('-')[2]).split('.')[0].trim(),
            'activo',
            computerNameOnly,
            viewer ? 'Si' : 'No'
            ]);
    } else {
        //console.log(`The word "${word}" was not found in the HTML file.`);
    }

    /* add row to worksheet */
   
  });
});

setTimeout(() =>{
  workbook.xlsx.writeFile(`${word}.xlsx`)
  .then(function() {
    console.log('File saved.');
  })
  .catch(function(error) {
    console.log(error);
  });

}, 1000)
