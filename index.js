const XLSX = require('xlsx');
const puppeteer = require('puppeteer');

// Load the workbook
const workbook = XLSX.readFile('input.xlsx');

// Get the first sheet name
const sheetName = workbook.SheetNames[0];

// Get the data from the first sheet
const sheet = workbook.Sheets[sheetName];

// Convert the sheet to JSON
const data = XLSX.utils.sheet_to_json(sheet);

// Prepare an array to hold the results
const results = [];

// Function to check for the <a> tag and log the rel attribute
async function checkLink(backlinkUrl, targetUrl) {
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();

    try {
        await page.goto(backlinkUrl, { waitUntil: 'networkidle2' });

        // Get the page content
        const content = await page.content();

        // Use Cheerio to parse the rendered HTML
        const $ = require('cheerio').load(content);

        // Find the <a> tag with the href matching the target URL
        const aTag = $('a[href="' + targetUrl + '"]');

        if (aTag.length > 0) {
          const relValue = aTag.attr('rel');
          let FollowStatus = relValue == undefined ? 'Dofollow' : "Nofollow"
          console.log(`Found <a> tag for URL: ${targetUrl} in ${backlinkUrl}. rel: ${FollowStatus}`);
          return FollowStatus; // If rel is undefined, return 'Dofollow'
      } else {
          console.log(`No <a> tag found for URL: ${targetUrl} in ${backlinkUrl}.`);
          return 'Not Found'; // If no <a> tag is found
      }
    } catch (error) {
        console.error(`Error fetching ${backlinkUrl}:`, error.message);
        return 'Website Not Found'; // Return 'Not Found' on error
    } finally {
        await browser.close();
    }
}

// Loop over the data and check each Backlink URL
async function processLinks() {
    for (const row of data) {
        if (row['Backlink URL'] && row['URL']) {
            const relValue = await checkLink(row['Backlink URL'], row['URL']);
            results.push({
                Sr: row['Sr#'],
                URL: row['URL'],
                Anchor: row['Anchor'],
                BacklinkURL: row['Backlink URL'],
                DA: row['DA'],
                SS: row['SS'],
                Type: relValue
            });
        }
    }

    // Write results to a new XLSX file
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.json_to_sheet(results);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Results');
    XLSX.writeFile(newWorkbook, 'output_file.xlsx');
}

// Start processing the links
processLinks();


