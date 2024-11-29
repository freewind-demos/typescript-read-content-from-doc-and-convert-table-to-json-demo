import * as mammoth from 'mammoth';
import * as cheerio from 'cheerio';
import * as path from 'path';

async function extractTableAfterTitle(docPath: string, targetTitle: string): Promise<Record<string, string> | null> {
    try {
        // Convert the .docx file to HTML
        const html = (await mammoth.convertToHtml({path: docPath})).value;

        // Load HTML content into cheerio
        const $ = cheerio.load(html);

        // Find the target title (h1, h2, p, etc.)
        const titleElement = $('*').filter((_, elem) => $(elem).text().trim() === targetTitle);
        
        if (titleElement.length === 0) {
            console.error(`Title "${targetTitle}" not found in document`);
            return null;
        }

        // Find the next table after the title
        const table = titleElement.nextAll('table').first();
        
        if (table.length === 0) {
            console.error('No table found after the specified title');
            return null;
        }

        const result: Record<string, string> = {};
        
        // Process table rows
        table.find('tr').each((index, row) => {
            const cells = $(row).find('td');
            
            // Skip first row as it's a header spanning all columns
            if (index === 0) return;
            
            // For other rows, expect key-value pairs
            if (cells.length === 2) {
                const key = $(cells[0]).text().trim();
                const value = $(cells[1]).text().trim();
                result[key] = value;
            }
        });

        return result;
    } catch (error) {
        console.error('Error processing document:', error);
        return null;
    }
}

// Example usage
if (process.argv.length < 4) {
    console.log('Usage: npm start <docx-file-path> <title-to-find>');
    process.exit(1);
}

const docPath = path.resolve(process.argv[2]);
const titleToFind = process.argv[3];

extractTableAfterTitle(docPath, titleToFind)
    .then(result => {
        if (result) {
            console.log('Extracted JSON:');
            console.log(JSON.stringify(result, null, 2));
        }
    })
    .catch(error => {
        console.error('Error:', error);
    });
