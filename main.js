const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const { Cluster } = require('puppeteer-cluster');

(async () => {
    const cluster = await Cluster.launch({
        concurrency: Cluster.CONCURRENCY_PAGE,
        maxConcurrency: 4, // Adjust the concurrency as per your requirements
        puppeteerOptions: {
          headless: false,
          executablePath: 'C:\\Users\\hyips\\AppData\\Local\\Chromium\\Application\\chrome.exe',
        },
      });

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Data');

  try {
    const data = await fs.promises.readFile('urls.txt', 'utf-8');
    const urls = data.split('\n');

    await cluster.task(async ({ page, data: data2 }) => {
        const { urlIndex, url } = data2;
        // console.log(url);
        await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
        await page.waitForSelector('.content_table');
  
        const data = await page.evaluate(() => {
          const table = document.querySelector('.content_table');
          const rows = Array.from(table.querySelectorAll('tr'));
          return rows.map((row) => {
            const cells = Array.from(row.querySelectorAll('td'));
            return cells.map((cell) => {
              const text = cell.innerHTML.trim();
              const scriptStart = text.indexOf('<script>');
              const scriptEnd = text.indexOf('</script>');
              const noscriptStart = text.indexOf('<noscript>');
  
              if (scriptStart !== -1 && scriptEnd !== -1 && noscriptStart !== -1) {
                const start = scriptEnd + '</script>'.length;
                const end = noscriptStart;
                return text.substring(start, end);
              }
  
              const bgColor = cell.getAttribute('bgcolor');
              const cellText = cell.textContent.trim();
  
              if (bgColor) {
                return `${cellText} (${bgColor})`;
              }
  
              return cellText;
            });
          });
        });
        
        data.forEach((row, index) => {
            // console.log(row);
            if(index === 0) {
                if(urlIndex > 0) {
                    return
                }
                row.splice(3,1)
                row.splice(4,1)
                row.splice(5,1)
                row[4] = "FT RESULTS"
                // row.splice(3, 0, "")
                console.log(row);
            }
            if(index === 1) {
                if(urlIndex > 0) {
                    return
                }
                row.unshift("")
                row.splice(2,1)
                row.splice(7,1)
            }
            if(index > 1) {
                row.splice(2,1)
                row.splice(2,1)
                row.splice(2,1)
                row.splice(2,1)
                row.splice(7,1)
            }
            const excelRow = worksheet.addRow(row);
          });
        
            if(urlIndex === 0) {
                worksheet.mergeCells("C1:G1")
                worksheet.mergeCells("H1:J1")
                worksheet.mergeCells("A1:A2")
                worksheet.mergeCells("K1:K2")
            }
        
            const resultCell = worksheet.getCell("K1")
            resultCell.value = "FT RESULTS"
            resultCell.alignment = { vertical: 'middle', horizontal: 'center' };
            worksheet.getColumn(11).width = 20
        
            const firstRow = worksheet.getRow(1);
        
            // Set the background color of the first row
            firstRow.eachCell((cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '778899' },
            };
            cell.font = {
                color: { argb: 'FFFFFF' },
                bold: true,
              };
            });
        
            const secondRow = worksheet.getRow(2);
        
            // Set the background color of the second row
            secondRow.eachCell((cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '778899' },
            };
            cell.font = {
                color: { argb: 'FFFFFF' },
                bold: true,
              };
            });
        
            const outcomeCell = worksheet.getCell('C1');
        
            // Center the value in the merged cell
            outcomeCell.alignment = { vertical: 'middle', horizontal: 'center' };
        
            const dateCell = worksheet.getCell('A1');
        
            // Center the value in the merged cell
            dateCell.alignment = { vertical: 'middle', horizontal: 'center' };
        
            const teamCell = worksheet.getCell('B2');
        
            // Center the value in the merged cell
            teamCell.alignment = { vertical: 'middle', horizontal: 'center' };
        
            const matchCell = worksheet.getCell('B1');
        
            // Center the value in the merged cell
            matchCell.alignment = { vertical: 'middle', horizontal: 'center' };
        
            const emptyCell = worksheet.getCell('H1');
        
            // Center the value in the merged cell
            emptyCell.value = "";
        
            // Set the width of column 2 (B)
            const columnB = worksheet.getColumn(2);
            columnB.width = 45; // Set the width to 15
        
        
            // Iterate through each row and cell
            worksheet.eachRow((row, rowNumber) => {
                row.eachCell((cell, colNumber) => {
                // Set border for each cell
                cell.border = {
                    top: { style: 'medium' },
                    left: { style: 'medium' },
                    bottom: { style: 'medium' },
                    right: { style: 'medium' },
                };
                
                // Regular expression pattern to match the color value inside parentheses
                const colorPattern = /\((.*?)\)/;
        
                // Function to extract the color value and remove it from the original string
                const extractColor = (string) => {
                const match = string.match(colorPattern);
                if (match) {
                    let color = match[1]; // Extract the color value
                    const modifiedString = string.replace(colorPattern, ""); // Remove the color value from the original string
                    return { color, modifiedString };
                }
                return { color: null, modifiedString: string };
                };
        
                let { color: bgColor, modifiedString: textValue } = extractColor(cell.value);
                cell.value = textValue
                if(bgColor != null) {
                    bgColor = bgColor == "AliceBlue" ? "F0F8FF" : bgColor
                    bgColor = bgColor == "Lime" ? "00D600" : bgColor
                    bgColor = bgColor == "Yellow" ? "D6D600" : bgColor
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: bgColor.replace("#", "") },
                    };
                }
        
                });
            });
        
            worksheet.eachRow((row, rowNumber) => {
                const cell = row.getCell(2); // Assuming column 2 is the second column (index 1)
                const cellValue = cell.value;
              
                if (!cellValue) {
                  worksheet.spliceRows(rowNumber, 1); // Remove the row if column 2 is empty
                }
            });
        
            worksheet.eachRow((row, rowNumber) => {
                const cell = row.getCell(2); // Assuming column 2 is the second column (index 1)
                const cellValue = cell.value;
              
                if (!cellValue) {
                  worksheet.spliceRows(rowNumber, 1); // Remove the row if column 2 is empty
                }
            });
        });

        for (const [urlIndex, url] of urls.entries()) {
            await cluster.queue({ urlIndex, url });
          }
      
          await cluster.idle();
          await cluster.close();

    await workbook.xlsx.writeFile('data.xlsx');
    console.log('Excel file created successfully');
  } catch (error) {
    console.error('Error reading file:', error);
  } finally {
    // await browser.close();
  }
})();
