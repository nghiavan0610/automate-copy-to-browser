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
    
    await page.goto(process.env.URL);

    await page.type('#Email', process.env.EMAIL);
    await page.type('#Password', process.env.PASSWORD);
    await page.click('button[type="submit"]');

    await page.waitForNavigation();

    const invalidOrderIds = [];
  
    for(const orderId of orderIds) {

        try {
            await page.waitForSelector('#OrderId');
            await page.type('#OrderId', orderId);

            await page.waitForSelector('#OrderDetailsRequestViewModel_GetOrderDetails');
            await page.click('button[id="OrderDetailsRequestViewModel_GetOrderDetails"]');
    
            // await page.waitForSelector('#alohaOrderStatus', {visible: true});
            await page.waitForFunction(() => {
                return document.getElementById('alohaOrderStatus').innerText; 
            }, {timeout: 10000});

            const statusText = await page.evaluate(() => {
                return document.getElementById('alohaOrderStatus').innerText; 
            });

            console.log(statusText);
        } catch (error) {
            console.log(`Order ID ${orderId} invalid`);
            console.log('error: ' + error);
            invalidOrderIds.push(orderId);
            continue;
        }
    }
    await fs.writeFile(outPath, invalidOrderIds.join('\n'));

    await browser.close();
  
  })();