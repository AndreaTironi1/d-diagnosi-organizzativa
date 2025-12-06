import express from 'express';
import Anthropic from '@anthropic-ai/sdk';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import multer from 'multer';
import xlsx from 'xlsx';
import jwt from 'jsonwebtoken';
import JSZip from 'jszip';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// JWT secret key
const JWT_SECRET = process.env.JWT_SECRET || 'prompt-executor-jwt-secret-key-2025';

// Hardcoded credentials
const CREDENTIALS = {
  username: 'dasein',
  password: 'Donatella2025!@'
};

// Configure multer for file uploads
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());

// JWT Authentication middleware
function authenticateToken(req, res, next) {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1]; // Bearer TOKEN

  if (!token) {
    return res.status(401).json({ error: 'Authentication required' });
  }

  jwt.verify(token, JWT_SECRET, (err, user) => {
    if (err) {
      return res.status(403).json({ error: 'Invalid or expired token' });
    }
    req.user = user;
    next();
  });
}

app.use(express.static('public'));

// Serve index.html for root path
app.get('/', (req, res) => {
  const filePath = join(process.cwd(), 'public', 'index.html');
  res.sendFile(filePath);
});

// Serve login.html
app.get('/login.html', (req, res) => {
  const filePath = join(process.cwd(), 'public', 'login.html');
  res.sendFile(filePath);
});

const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
});

// Replace variables in prompt template using square bracket syntax [column_name]
function replaceVariables(template, variables) {
  let result = template;
  for (const [key, value] of Object.entries(variables)) {
    // Escape special regex characters in the key
    const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(`\\[${escapedKey}\\]`, 'g');
    // Convert value to string and handle null/undefined
    const stringValue = value != null ? String(value) : '';
    result = result.replace(regex, stringValue);
  }
  return result;
}

// Extract variable names from template (now looking for [variable] syntax)
function extractVariableNames(template) {
  const regex = /\[([^\]]+)\]/g;
  const variables = new Set();
  let match;
  while ((match = regex.exec(template)) !== null) {
    variables.add(match[1]);
  }
  return Array.from(variables);
}

// API endpoint to execute prompt
app.post('/api/execute', authenticateToken, async (req, res) => {
  try {
    const { prompt, variables } = req.body;

    if (!prompt) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    if (!process.env.ANTHROPIC_API_KEY) {
      return res.status(500).json({
        error: 'ANTHROPIC_API_KEY not configured. Please set it in .env file'
      });
    }

    // Replace variables in the prompt
    const processedPrompt = replaceVariables(prompt, variables || {});

    // Call Claude API
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-5-20250929',
      max_tokens: 8192,
      messages: [
        {
          role: 'user',
          content: processedPrompt
        }
      ]
    });

    res.json({
      success: true,
      processedPrompt,
      response: message.content[0].text,
      usage: {
        inputTokens: message.usage.input_tokens,
        outputTokens: message.usage.output_tokens
      }
    });

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({
      error: error.message || 'Failed to process request'
    });
  }
});

// Login endpoint
app.post('/api/login', (req, res) => {
  const { username, password } = req.body;

  if (username === CREDENTIALS.username && password === CREDENTIALS.password) {
    // Generate JWT token
    const token = jwt.sign(
      { username: username },
      JWT_SECRET,
      { expiresIn: '24h' }
    );

    res.json({
      success: true,
      message: 'Login successful',
      token: token
    });
  } else {
    res.status(401).json({ success: false, error: 'Invalid username or password' });
  }
});

// Logout endpoint (client-side will remove token)
app.post('/api/logout', (req, res) => {
  res.json({ success: true, message: 'Logged out successfully' });
});

// API endpoint to parse variables from prompt
app.post('/api/parse-variables', authenticateToken, (req, res) => {
  try {
    const { prompt } = req.body;

    if (!prompt) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    const variables = extractVariableNames(prompt);
    res.json({ variables });

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: 'Failed to parse variables' });
  }
});

// API endpoint to upload and parse Excel file
app.post('/api/upload-excel', authenticateToken, upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    // Parse Excel file
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert to JSON
    const data = xlsx.utils.sheet_to_json(sheet);

    if (data.length === 0) {
      return res.status(400).json({ error: 'Excel file is empty' });
    }

    // Get column names from first row
    const columns = Object.keys(data[0]);

    res.json({
      success: true,
      columns,
      rowCount: data.length,
      data
    });

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: 'Failed to parse Excel file' });
  }
});

// API endpoint to execute prompt for all rows in Excel
app.post('/api/execute-batch', authenticateToken, async (req, res) => {
  try {
    const { prompt, excelData, model } = req.body;

    if (!prompt) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    if (!excelData || !Array.isArray(excelData) || excelData.length === 0) {
      return res.status(400).json({ error: 'Excel data is required' });
    }

    if (!process.env.ANTHROPIC_API_KEY) {
      return res.status(500).json({
        error: 'ANTHROPIC_API_KEY not configured. Please set it in .env file'
      });
    }

    // Use provided model or default to Sonnet 4.5
    const selectedModel = model || 'claude-sonnet-4-5-20250929';

    const results = [];

    // Process each row
    for (let i = 0; i < excelData.length; i++) {
      const row = excelData[i];

      // Replace variables with values from current row
      const processedPrompt = replaceVariables(prompt, row);

      try {
        // Call Claude API
        const message = await anthropic.messages.create({
          model: selectedModel,
          max_tokens: 8192,
          messages: [
            {
              role: 'user',
              content: processedPrompt
            }
          ]
        });

        results.push({
          rowIndex: i,
          rowData: row,
          processedPrompt,
          response: message.content[0].text,
          usage: {
            inputTokens: message.usage.input_tokens,
            outputTokens: message.usage.output_tokens
          },
          success: true
        });

      } catch (error) {
        results.push({
          rowIndex: i,
          rowData: row,
          processedPrompt,
          error: error.message,
          success: false
        });
      }
    }

    // Calculate total usage
    const totalUsage = results.reduce((acc, result) => {
      if (result.success) {
        acc.inputTokens += result.usage.inputTokens;
        acc.outputTokens += result.usage.outputTokens;
      }
      return acc;
    }, { inputTokens: 0, outputTokens: 0 });

    res.json({
      success: true,
      results,
      totalUsage,
      successCount: results.filter(r => r.success).length,
      errorCount: results.filter(r => !r.success).length
    });

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({
      error: error.message || 'Failed to process batch request'
    });
  }
});

// Helper function to create a single Excel file from a result
function createExcelFromResult(result, rowIndex) {
  const workbook = xlsx.utils.book_new();

  // Try to parse JSON or CSV from Claude response
  let parsedData = null;
  let jsonFound = false;
  let csvFound = false;

  console.log(`\n=== Processing Excel for row ${rowIndex} ===`);
  console.log(`Response length: ${result.response ? result.response.length : 0} characters`);

  if (result.success && result.response) {
    const response = result.response.trim();
    console.log(`First 200 chars: ${response.substring(0, 200)}`);

    try {
      // FIRST: Try CSV parsing (new approach)
      // Look for CSV with semicolon separator and header
      if (response.includes(';') && response.includes('Tabella;Area_Contrattuale')) {
        console.log('🔍 Detected CSV format with semicolon separator');
        try {
          // Extract CSV (remove any markdown code blocks)
          let csvContent = response;
          const csvBlockMatch = response.match(/```(?:csv)?\s*([\s\S]*?)\s*```/i);
          if (csvBlockMatch) {
            csvContent = csvBlockMatch[1].trim();
            console.log('Extracted CSV from code block');
          }

          // Parse CSV into array of objects
          const lines = csvContent.split('\n').map(line => line.trim()).filter(line => line);
          if (lines.length > 1) {
            const headers = lines[0].split(';').map(h => h.trim());
            console.log(`CSV headers: ${headers.join(', ')}`);

            const csvData = [];
            for (let i = 1; i < lines.length; i++) {
              const values = lines[i].split(';').map(v => v.trim().replace(/^"|"$/g, ''));
              if (values.length === headers.length) {
                const obj = {};
                headers.forEach((header, index) => {
                  obj[header] = values[index];
                });
                csvData.push(obj);
              }
            }

            if (csvData.length > 0) {
              // Group CSV data by table
              const groupedByTable = {};
              csvData.forEach(row => {
                const tableName = row.Tabella || row.tabella || 'UNKNOWN';
                if (!groupedByTable[tableName]) {
                  groupedByTable[tableName] = [];
                }
                groupedByTable[tableName].push(row);
              });

              // Convert to PA competenze format
              parsedData = {
                tabella_1_normativa_generale: groupedByTable['TABELLA_1'] || [],
                tabella_2_normativa_nazionale_regionale: groupedByTable['TABELLA_2'] || [],
                tabella_3_normativa_specifica_profilo: groupedByTable['TABELLA_3'] || [],
                tabella_4_competenze_tecnico_specialistiche: groupedByTable['TABELLA_4'] || [],
                tabella_5_competenze_gestionali_procedurali: groupedByTable['TABELLA_5'] || [],
                tabella_6_competenze_trasversali: groupedByTable['TABELLA_6'] || [],
                tabella_7_competenze_informatiche: groupedByTable['TABELLA_7'] || [],
                tabella_8_competenze_linguistiche: groupedByTable['TABELLA_8'] || []
              };

              csvFound = true;
              console.log(`✅ Successfully parsed CSV: ${csvData.length} total rows`);
              console.log(`📊 Tables found: ${Object.keys(groupedByTable).join(', ')}`);
            }
          }
        } catch (e) {
          console.log('Failed to parse CSV:', e.message);
        }
      }

      // FALLBACK: Try multiple JSON extraction strategies (if CSV not found)
      if (!csvFound) {
        // Strategy 1: Look for ```json code blocks (case-insensitive, flexible whitespace)
        const jsonBlockMatch = response.match(/```\s*json\s*([\s\S]*?)\s*```/i);
        if (jsonBlockMatch) {
        console.log('Found JSON in code block (```json)');
        try {
          parsedData = JSON.parse(jsonBlockMatch[1].trim());
          jsonFound = true;
          console.log('✅ Successfully parsed JSON from code block');
        } catch (e) {
          console.log('Failed to parse JSON from code block:', e.message);
        }
      }

      // Strategy 2: Look for any code block ``` (might contain JSON)
      if (!jsonFound) {
        const anyBlockMatch = response.match(/```([\s\S]*?)```/);
        if (anyBlockMatch) {
          console.log('Found generic code block (```)');
          try {
            parsedData = JSON.parse(anyBlockMatch[1].trim());
            jsonFound = true;
            console.log('✅ Successfully parsed JSON from generic code block');
          } catch (e) {
            console.log('Generic code block does not contain valid JSON');
          }
        }
      }

      // Strategy 3: Try to find JSON object/array in the response (look for { or [)
      if (!jsonFound) {
        const jsonObjectMatch = response.match(/\{[\s\S]*\}/);
        const jsonArrayMatch = response.match(/\[[\s\S]*\]/);

        if (jsonObjectMatch || jsonArrayMatch) {
          const jsonString = jsonObjectMatch ? jsonObjectMatch[0] : jsonArrayMatch[0];
          console.log('Found JSON-like structure in response');
          try {
            parsedData = JSON.parse(jsonString);
            jsonFound = true;
            console.log('✅ Successfully parsed JSON structure');
          } catch (e) {
            console.log('JSON-like structure is not valid JSON:', e.message);
          }
        }
      }

      // Strategy 4: Try to parse the entire response as JSON
      if (!jsonFound) {
        console.log('Trying to parse entire response as JSON');
        try {
          parsedData = JSON.parse(response);
          jsonFound = true;
          console.log('✅ Successfully parsed entire response as JSON');
        } catch (e) {
          console.log('Entire response is not valid JSON');
        }
      }

      } // End of if (!csvFound) for JSON parsing

      // Log what we found
      if (csvFound) {
        console.log('✅ CSV parsing successful');
      } else if (jsonFound && parsedData) {
        console.log('📊 Parsed data type:', Array.isArray(parsedData) ? `Array[${parsedData.length}]` : typeof parsedData);
        if (typeof parsedData === 'object' && !Array.isArray(parsedData)) {
          console.log('📊 Object keys:', Object.keys(parsedData).join(', '));
        }
      } else {
        console.log('⚠️ No CSV or JSON found in response');
      }

    } catch (e) {
      console.log('❌ Unexpected error in JSON parsing:', e.message);
    }
  }

  // If we have PA competenze format, create sheets
  if (parsedData && parsedData.tabella_1_normativa_generale) {
    const tables = {
      'Normativa Generale': parsedData.tabella_1_normativa_generale,
      'Normativa Naz-Reg': parsedData.tabella_2_normativa_nazionale_regionale,
      'Normativa Specifica': parsedData.tabella_3_normativa_specifica_profilo,
      'Competenze Tecnico-Spec': parsedData.tabella_4_competenze_tecnico_specialistiche,
      'Competenze Gestionali': parsedData.tabella_5_competenze_gestionali_procedurali,
      'Competenze Trasversali': parsedData.tabella_6_competenze_trasversali,
      'Competenze Informatiche': parsedData.tabella_7_competenze_informatiche,
      'Competenze Linguistiche': parsedData.tabella_8_competenze_linguistiche
    };

    for (const [sheetName, tableData] of Object.entries(tables)) {
      if (tableData && Array.isArray(tableData) && tableData.length > 0) {
        const sheet = xlsx.utils.json_to_sheet(tableData);
        xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
      }
    }

    // Add summary sheets
    if (parsedData.sintesi_esecutiva) {
      const sintesiData = [{
        'Testo': parsedData.sintesi_esecutiva.testo || '',
        'Top 3 Competenze': (parsedData.sintesi_esecutiva.top_3_competenze_critiche || []).join('; '),
        'Priorità Formative': (parsedData.sintesi_esecutiva.priorita_formative || []).join('; '),
        'Gap Tipici': (parsedData.sintesi_esecutiva.gap_tipici || []).join('; '),
        'Normative Regionali': parsedData.sintesi_esecutiva.normative_regionali || ''
      }];
      const sintesiSheet = xlsx.utils.json_to_sheet(sintesiData);
      xlsx.utils.book_append_sheet(workbook, sintesiSheet, 'Sintesi Esecutiva');
    }

    if (parsedData.raccomandazioni_operative) {
      const raccData = [{
        'Percorsi Formativi': (parsedData.raccomandazioni_operative.percorsi_formativi || []).join('; '),
        'Certificazioni Utili': (parsedData.raccomandazioni_operative.certificazioni_utili || []).join('; '),
        'Modalità di Verifica': (parsedData.raccomandazioni_operative.modalita_verifica || []).join('; ')
      }];
      const raccSheet = xlsx.utils.json_to_sheet(raccData);
      xlsx.utils.book_append_sheet(workbook, raccSheet, 'Raccomandazioni');
    }
  }

  // ALWAYS add a Debug sheet with the raw response for troubleshooting
  if (result.success && result.response) {
    const debugData = [{
      'Row Index': rowIndex + 1,
      'Response Length': result.response.length,
      'Format Detected': csvFound ? 'CSV' : (jsonFound ? 'JSON' : 'Unknown'),
      'CSV Found': csvFound ? 'YES' : 'NO',
      'JSON Found': jsonFound ? 'YES' : 'NO',
      'Parsed Type': parsedData ? (Array.isArray(parsedData) ? `Array[${parsedData.length}]` : typeof parsedData) : 'null',
      'First 500 chars': result.response.substring(0, 500),
      'Last 500 chars': result.response.substring(Math.max(0, result.response.length - 500))
    }];
    const debugSheet = xlsx.utils.json_to_sheet(debugData);
    xlsx.utils.book_append_sheet(workbook, debugSheet, 'DEBUG');
  }

  if (!parsedData || (!parsedData.tabella_1_normativa_generale && !Array.isArray(parsedData) && typeof parsedData !== 'object')) {
    // Fallback: check if parsedData is an array, if so explode it into rows
    if (parsedData && Array.isArray(parsedData) && parsedData.length > 0) {
      console.log(`Creating Data sheet with ${parsedData.length} rows from array`);
      const sheet = xlsx.utils.json_to_sheet(parsedData);
      xlsx.utils.book_append_sheet(workbook, sheet, 'Data');
      console.log('✅ Array exploded into Excel rows successfully');
    } else if (parsedData && typeof parsedData === 'object' && !Array.isArray(parsedData)) {
      console.log('parsedData is an object (not array), looking for arrays inside...');
      // If parsedData is an object, try to find arrays within it and create sheets
      let hasSheets = false;
      for (const [key, value] of Object.entries(parsedData)) {
        if (Array.isArray(value) && value.length > 0) {
          const sheetName = key.substring(0, 31); // Excel sheet name limit
          console.log(`Creating sheet '${sheetName}' with ${value.length} rows from object property '${key}'`);
          const sheet = xlsx.utils.json_to_sheet(value);
          xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
          hasSheets = true;
          console.log(`✅ Array '${key}' exploded into Excel rows successfully`);
        }
      }

      // If no arrays found, create a sheet with the object properties
      if (!hasSheets) {
        console.log('No arrays found in object, creating single row with object properties');
        const sheet = xlsx.utils.json_to_sheet([parsedData]);
        xlsx.utils.book_append_sheet(workbook, sheet, 'Data');
      }
    } else {
      console.log('⚠️ Using fallback - no parseable JSON data found');
      // Ultimate fallback: create a simple info sheet without raw JSON/text in cells
      const fallbackData = [{
        'Row #': rowIndex + 1,
        ...result.rowData,
        'Status': result.success ? 'Success - No structured data found' : 'Error',
        'Note': result.success ? 'Response did not contain parseable JSON data' : (result.error || 'Request failed')
      }];
      if (result.usage) {
        fallbackData[0]['Input Tokens'] = result.usage.inputTokens;
        fallbackData[0]['Output Tokens'] = result.usage.outputTokens;
      }
      const sheet = xlsx.utils.json_to_sheet(fallbackData);
      xlsx.utils.book_append_sheet(workbook, sheet, 'Info');

      // If there's a response but it's not JSON, create a separate text sheet with chunks
      if (result.success && result.response && !jsonFound) {
        // Split long text into manageable chunks (Excel cell limit is 32767 chars)
        const maxChunkSize = 30000;
        const responseText = result.response;
        const chunks = [];

        for (let i = 0; i < responseText.length; i += maxChunkSize) {
          chunks.push({
            'Part': Math.floor(i / maxChunkSize) + 1,
            'Content': responseText.substring(i, i + maxChunkSize)
          });
        }

        const textSheet = xlsx.utils.json_to_sheet(chunks);
        xlsx.utils.book_append_sheet(workbook, textSheet, 'Response Text');
      }
    }
  }

  return workbook;
}

// Helper function to generate filename from row data
function generateFilename(result, rowIndex) {
  const rowData = result.rowData || {};
  const profilo = rowData.PROFILO || rowData.Profilo || rowData.profilo || '';
  const settore = rowData.SETTORE || rowData.Settore || rowData.settore || '';

  let filename = 'Analisi';
  if (profilo) filename += `_${profilo.replace(/[^a-zA-Z0-9]/g, '_')}`;
  if (settore) filename += `_${settore.replace(/[^a-zA-Z0-9]/g, '_')}`;
  filename += `_row${rowIndex + 1}`;

  return filename + '.xlsx';
}

// API endpoint to download single row as Excel
app.post('/api/download-single-excel', authenticateToken, (req, res) => {
  try {
    const { result, rowIndex } = req.body;

    if (!result) {
      return res.status(400).json({ error: 'Result data is required' });
    }

    const workbook = createExcelFromResult(result, rowIndex || 0);
    const excelBuffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    const filename = generateFilename(result, rowIndex || 0);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
    res.send(excelBuffer);

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: 'Failed to generate Excel file' });
  }
});

// API endpoint to download all results as ZIP
app.post('/api/download-excel-zip', authenticateToken, async (req, res) => {
  try {
    const { results } = req.body;

    if (!results || !Array.isArray(results)) {
      return res.status(400).json({ error: 'Results data is required' });
    }

    const zip = new JSZip();

    // Create an Excel file for each result
    results.forEach((result, index) => {
      const workbook = createExcelFromResult(result, index);
      const excelBuffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
      const filename = generateFilename(result, index);
      zip.file(filename, excelBuffer);
    });

    // Generate ZIP file
    const zipBuffer = await zip.generateAsync({ type: 'nodebuffer' });

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename=analisi_competenze_${Date.now()}.zip`);
    res.send(zipBuffer);

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: 'Failed to generate ZIP file' });
  }
});

// API endpoint to download results as Excel
app.post('/api/download-excel', authenticateToken, (req, res) => {
  try {
    const { results } = req.body;

    if (!results || !Array.isArray(results)) {
      return res.status(400).json({ error: 'Results data is required' });
    }

    // Create a new workbook
    const workbook = xlsx.utils.book_new();

    // Prepare data for Excel - add original data + Claude response
    const excelData = results.map((result, index) => {
      const row = {
        'Row #': index + 1,
        ...result.rowData,
        'Claude Response': result.response || result.error || 'N/A',
        'Status': result.success ? 'Success' : 'Error'
      };

      if (result.usage) {
        row['Input Tokens'] = result.usage.inputTokens;
        row['Output Tokens'] = result.usage.outputTokens;
      }

      return row;
    });

    // Try to parse JSON responses from Claude
    results.forEach((result, index) => {
      if (result.success && result.response) {
        try {
          // Try to extract JSON from the response
          const jsonMatch = result.response.match(/```json\s*([\s\S]*?)\s*```/);
          if (!jsonMatch) {
            // Try without code blocks
            const parsed = JSON.parse(result.response);
            processStructuredData(workbook, parsed, index);
          } else {
            const parsed = JSON.parse(jsonMatch[1]);
            processStructuredData(workbook, parsed, index);
          }
        } catch (e) {
          // If parsing fails, skip
          console.log('Failed to parse JSON for row', index, e.message);
        }
      }
    });

    // Create main sheet with all responses
    const mainSheet = xlsx.utils.json_to_sheet(excelData);
    xlsx.utils.book_append_sheet(workbook, mainSheet, 'Results');

    // Helper function to process structured data
    function processStructuredData(workbook, data, sourceRow) {
      // Check if this is the PA competenze format
      if (data.tabella_1_normativa_generale) {
        // Process each table as a separate sheet
        const tables = {
          'Normativa Generale': data.tabella_1_normativa_generale,
          'Normativa Naz-Reg': data.tabella_2_normativa_nazionale_regionale,
          'Normativa Specifica': data.tabella_3_normativa_specifica_profilo,
          'Competenze Tecnico-Spec': data.tabella_4_competenze_tecnico_specialistiche,
          'Competenze Gestionali': data.tabella_5_competenze_gestionali_procedurali,
          'Competenze Trasversali': data.tabella_6_competenze_trasversali,
          'Competenze Informatiche': data.tabella_7_competenze_informatiche,
          'Competenze Linguistiche': data.tabella_8_competenze_linguistiche
        };

        for (const [sheetName, tableData] of Object.entries(tables)) {
          if (tableData && Array.isArray(tableData) && tableData.length > 0) {
            const sheet = xlsx.utils.json_to_sheet(tableData);
            xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
          }
        }

        // Add summary sheets if present
        if (data.sintesi_esecutiva) {
          const sintesiData = [{
            'Testo': data.sintesi_esecutiva.testo || '',
            'Top 3 Competenze': (data.sintesi_esecutiva.top_3_competenze_critiche || []).join('; '),
            'Priorità Formative': (data.sintesi_esecutiva.priorita_formative || []).join('; '),
            'Gap Tipici': (data.sintesi_esecutiva.gap_tipici || []).join('; '),
            'Normative Regionali': data.sintesi_esecutiva.normative_regionali || ''
          }];
          const sintesiSheet = xlsx.utils.json_to_sheet(sintesiData);
          xlsx.utils.book_append_sheet(workbook, sintesiSheet, 'Sintesi Esecutiva');
        }

        if (data.raccomandazioni_operative) {
          const raccData = [{
            'Percorsi Formativi': (data.raccomandazioni_operative.percorsi_formativi || []).join('; '),
            'Certificazioni Utili': (data.raccomandazioni_operative.certificazioni_utili || []).join('; '),
            'Modalità di Verifica': (data.raccomandazioni_operative.modalita_verifica || []).join('; ')
          }];
          const raccSheet = xlsx.utils.json_to_sheet(raccData);
          xlsx.utils.book_append_sheet(workbook, raccSheet, 'Raccomandazioni');
        }
      } else if (Array.isArray(data)) {
        // Generic array data
        const sheet = xlsx.utils.json_to_sheet(data);
        xlsx.utils.book_append_sheet(workbook, sheet, `Data Row ${sourceRow + 1}`);
      } else if (typeof data === 'object') {
        // Generic object - convert to single row
        const sheet = xlsx.utils.json_to_sheet([data]);
        xlsx.utils.book_append_sheet(workbook, sheet, `Data Row ${sourceRow + 1}`);
      }
    }

    // Generate Excel file buffer
    const excelBuffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    // Set headers for file download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=claude_results_${Date.now()}.xlsx`);

    res.send(excelBuffer);

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: 'Failed to generate Excel file' });
  }
});

// Only listen on port in local development
if (process.env.NODE_ENV !== 'production') {
  app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
    if (!process.env.ANTHROPIC_API_KEY) {
      console.warn('⚠️  Warning: ANTHROPIC_API_KEY not set in .env file');
    }
  });
}

// Export for Vercel
export default app;
