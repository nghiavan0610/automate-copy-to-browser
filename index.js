const xlsx = require('xlsx');
const puppeteer = require('puppeteer');
require('dotenv').config();
const fs = require('fs').promises;
const path = require('path');

// output
const date = new Date();

let dateTime = `${date.getFullYear()}-${date.getMonth()+1}-${date.getDate()}_`; 
dateTime += `${date.getHours()}h${date.getMinutes()}`;

const outFolder = path.join(__dirname, '../zalo-order');
const outPath = path.join(outFolder, `invalid-orders_${dateTime}.txt`);

// excel file path
const excelFolder = process.argv[2];

const workbook = xlsx.readFile(excelFolder); 
const workbook_sheet = workbook.SheetNames;

const workbook_response = xlsx.utils.sheet_to_json(workbook.Sheets[workbook_sheet[0]]);

const orderIds = workbook_response.map(order => order['Order Id']);

(async () => {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    
    await page.goto(process.env.LOGIN_URL);

    await page.evaluate(() => {
        const div = document.querySelector('div'); 
        if(div.innerText.includes("NHÂN VIÊN")) {
          div.click();
        } 
    });

    await page.waitForSelector('#username');

    await page.type('#username', process.env.USER);
    await page.type('#password', process.env.PASSWORD);
    await page.click('input[type="submit"]');

    // await page.waitForNavigation();

    // navigate to translog
    await page.waitForSelector('#transaction-menu');
    await page.click('a[href="/translog"]');

    const validOrderIds = [];
  
    for(const orderId of orderIds) {
        console.log('orderId: ', orderId);
        try {
            await page.waitForSelector('#txtTransID');
            await page.type('#txtTransID', orderId);

            await page.waitForSelector('#btnSearch');
            await page.click('button[id="btnSearch"]');

            await page.waitForTimeout(5000);

            const zaloPayId = await page.evaluate(() => {
                return document.querySelector('tbody tr td:nth-child(4)').innerText;
            });

            console.log(`zaloPayId: ${zaloPayId} - orderId: ${orderId}`);
              
            if(zaloPayId === orderId) {
                const status = await page.evaluate(() => {
                    return document.querySelector('tbody tr td:nth-child(10)').innerText; 
                });
            
                if (status.includes("Thành công")) {
                    validOrderIds.push(orderId);
                }
            } else {
                console.log(`Order ID ${orderId} not found`);
                continue;
            }
        } catch (error) {
            console.log(`Order ID ${orderId} invalid`);
            console.log('error: ' + error);
            validOrderIds.push(orderId);
            continue;
        }
    }
    await fs.writeFile(outPath, validOrderIds.join('\n'));

    await browser.close();
  
  })();