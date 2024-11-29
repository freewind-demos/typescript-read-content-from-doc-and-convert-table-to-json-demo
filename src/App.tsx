import React, { useState } from 'react';
import * as mammoth from 'mammoth';
import * as cheerio from 'cheerio';

function extractTableAfterTitle(html: string, targetTitle: string): Record<string, string> | null {
    try {
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

function App() {
    const [file, setFile] = useState<File | null>(null);
    const [title, setTitle] = useState('');
    const [result, setResult] = useState<Record<string, string> | null>(null);
    const [error, setError] = useState<string | null>(null);

    const handleSubmit = async (e: React.FormEvent) => {
        e.preventDefault();
        if (!file || !title) {
            setError('Please select a file and enter a title');
            return;
        }

        try {
            // Convert file to ArrayBuffer
            const arrayBuffer = await file.arrayBuffer();
            
            // Convert to HTML using mammoth
            const result = await mammoth.convertToHtml({arrayBuffer});
            const html = result.value;

            // Extract table data
            const data = extractTableAfterTitle(html, title);
            
            if (data) {
                setResult(data);
                setError(null);
            } else {
                setError('Could not find table after the specified title');
                setResult(null);
            }
        } catch (err) {
            setError(err instanceof Error ? err.message : 'An error occurred');
            setResult(null);
        }
    };

    return (
        <div style={{ maxWidth: '800px', margin: '0 auto', padding: '20px' }}>
            <h1>Word Table to JSON Converter</h1>
            
            <form onSubmit={handleSubmit} style={{ marginBottom: '20px' }}>
                <div style={{ marginBottom: '15px' }}>
                    <label style={{ display: 'block', marginBottom: '5px' }}>
                        Word Document:
                    </label>
                    <input
                        type="file"
                        accept=".docx"
                        onChange={(e) => setFile(e.target.files?.[0] || null)}
                        style={{ marginBottom: '10px' }}
                    />
                </div>

                <div style={{ marginBottom: '15px' }}>
                    <label style={{ display: 'block', marginBottom: '5px' }}>
                        Title to Find:
                    </label>
                    <input
                        type="text"
                        value={title}
                        onChange={(e) => setTitle(e.target.value)}
                        style={{
                            width: '100%',
                            padding: '8px',
                            borderRadius: '4px',
                            border: '1px solid #ccc'
                        }}
                    />
                </div>

                <button
                    type="submit"
                    style={{
                        padding: '10px 20px',
                        backgroundColor: '#007bff',
                        color: 'white',
                        border: 'none',
                        borderRadius: '4px',
                        cursor: 'pointer'
                    }}
                >
                    Convert
                </button>
            </form>

            {error && (
                <div style={{ color: 'red', marginBottom: '20px' }}>
                    {error}
                </div>
            )}

            {result && (
                <div>
                    <h2>Result:</h2>
                    <pre style={{
                        backgroundColor: '#f5f5f5',
                        padding: '15px',
                        borderRadius: '4px',
                        overflow: 'auto'
                    }}>
                        {JSON.stringify(result, null, 2)}
                    </pre>
                </div>
            )}
        </div>
    );
}

export default App;
