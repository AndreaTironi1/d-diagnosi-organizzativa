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

app.use(express.json({ limit: '4mb' }));

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

// API endpoint to get account balance/credits
app.get('/api/account-balance', authenticateToken, async (req, res) => {
  try {
    if (!process.env.ANTHROPIC_API_KEY) {
      return res.status(500).json({
        error: 'ANTHROPIC_API_KEY not configured'
      });
    }

    // Get account balance using Anthropic SDK
    // Note: This is a placeholder - Anthropic API doesn't expose balance endpoint directly
    // We'll return usage info from recent request instead
    res.json({
      success: true,
      message: 'Balance information not available via API',
      note: 'Monitor usage at https://console.anthropic.com'
    });

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({
      error: error.message || 'Failed to get balance'
    });
  }
});

// API endpoint to execute prompt
app.post('/api/execute', authenticateToken, async (req, res) => {
  try {
    const { prompt, variables, rowIndex, model } = req.body;

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

    // Use provided model or default to Sonnet 4.5
    const selectedModel = model || 'claude-sonnet-4-5-20250929';

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

    res.json({
      success: true,
      rowIndex: rowIndex !== undefined ? rowIndex : 0,
      rowData: variables || {},
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
      // Look for CSV with semicolon separator and header (optimized format with Nome_Tabella)
      if (response.includes(';') && (response.includes('Nome_Tabella;Area_Contrattuale') || response.includes('Tabella;Area_Contrattuale'))) {
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
              // Check if this is the new flat format with Nome_Tabella column
              if (csvData[0].Nome_Tabella) {
                // New optimized format - use data directly as flat array
                console.log(`✅ Detected optimized flat CSV format with ${csvData.length} rows`);
                parsedData = { competenze_flat: csvData };
                csvFound = true;
              } else {
                // Old format - group by Tabella column
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
                  tabella_1_normativa_generale: groupedByTable['T1'] || [],
                  tabella_2_normativa_nazionale_regionale: groupedByTable['T2'] || [],
                  tabella_3_normativa_specifica_profilo: groupedByTable['T3'] || [],
                  tabella_4_competenze_tecnico_specialistiche: groupedByTable['T4'] || [],
                  tabella_5_competenze_gestionali_procedurali: groupedByTable['T5'] || [],
                  tabella_6_competenze_trasversali: groupedByTable['T6'] || [],
                  tabella_7_competenze_informatiche: groupedByTable['T7'] || [],
                  tabella_8_competenze_linguistiche: groupedByTable['T8'] || []
                };

                csvFound = true;
                console.log(`✅ Successfully parsed CSV: ${csvData.length} total rows`);
                console.log(`📊 Tables found: ${Object.keys(groupedByTable).join(', ')}`);
              }
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

  // Track if RISULTATI sheet was created
  let risultatiCreated = false;

  // If we have PA competenze format, create only RISULTATI sheet
  if (parsedData && (parsedData.competenze_flat || parsedData.tabella_1_normativa_generale)) {
    // New optimized flat format - data already combined
    if (parsedData.competenze_flat && Array.isArray(parsedData.competenze_flat)) {
      console.log(`Creating RISULTATI sheet with ${parsedData.competenze_flat.length} rows (optimized flat format)`);
      const tutteSheet = xlsx.utils.json_to_sheet(parsedData.competenze_flat);
      xlsx.utils.book_append_sheet(workbook, tutteSheet, 'RISULTATI');
      console.log('✅ RISULTATI sheet created successfully');
      risultatiCreated = true;
    } else {
      // Old format with 8 separate tables - combine them
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

      const tutteLeRighe = [];
      for (const [sheetName, tableData] of Object.entries(tables)) {
        if (tableData && Array.isArray(tableData) && tableData.length > 0) {
          tutteLeRighe.push(...tableData);
        }
      }

      if (tutteLeRighe.length > 0) {
        console.log(`Creating RISULTATI sheet with ${tutteLeRighe.length} total rows`);
        const tutteSheet = xlsx.utils.json_to_sheet(tutteLeRighe);
        xlsx.utils.book_append_sheet(workbook, tutteSheet, 'RISULTATI');
        console.log('✅ RISULTATI sheet created successfully');
        risultatiCreated = true;
      }
    }
  }


  // Only create fallback sheets if RISULTATI was not created
  if (!risultatiCreated) {
    if (!parsedData || (!parsedData.tabella_1_normativa_generale && !Array.isArray(parsedData) && typeof parsedData !== 'object')) {
      // Fallback: check if parsedData is an array, if so explode it into rows
      if (parsedData && Array.isArray(parsedData) && parsedData.length > 0) {
        console.log(`Creating RISULTATO sheet with ${parsedData.length} rows from array`);
        const sheet = xlsx.utils.json_to_sheet(parsedData);
        xlsx.utils.book_append_sheet(workbook, sheet, 'RISULTATO');
        console.log('✅ Array exploded into Excel rows successfully');
      } else if (parsedData && typeof parsedData === 'object' && !Array.isArray(parsedData)) {
        console.log('parsedData is an object (not array), looking for arrays inside...');
        // If parsedData is an object, try to find arrays within it and create RISULTATO sheet
        let hasSheets = false;
        for (const [key, value] of Object.entries(parsedData)) {
          if (Array.isArray(value) && value.length > 0) {
            console.log(`Creating RISULTATO sheet with ${value.length} rows from object property '${key}'`);
            const sheet = xlsx.utils.json_to_sheet(value);
            xlsx.utils.book_append_sheet(workbook, sheet, 'RISULTATO');
            hasSheets = true;
            console.log(`✅ Array '${key}' exploded into Excel rows successfully`);
            break; // Only create one RISULTATO sheet from first array found
          }
        }

        // If no arrays found, try to parse response as CSV
        if (!hasSheets && result.success && result.response) {
          console.log('⚠️ No arrays in object, attempting to parse response as CSV');
          const response = result.response.trim();

          // Try to parse as CSV with semicolon separator
          if (response.includes(';')) {
            try {
              const lines = response.split('\n').map(line => line.trim()).filter(line => line);
              if (lines.length > 1) {
                const headers = lines[0].split(';').map(h => h.trim());
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
                  console.log(`✅ Parsed CSV with ${csvData.length} rows, creating RISULTATO sheet`);
                  const sheet = xlsx.utils.json_to_sheet(csvData);
                  xlsx.utils.book_append_sheet(workbook, sheet, 'RISULTATO');
                  hasSheets = true;
                }
              }
            } catch (e) {
              console.log('Failed to parse response as CSV:', e.message);
            }
          }
        }

        // If still no sheets, create a sheet with the object properties
        if (!hasSheets) {
          console.log('No data found, creating single row with object properties');
          const sheet = xlsx.utils.json_to_sheet([parsedData]);
          xlsx.utils.book_append_sheet(workbook, sheet, 'RISULTATO');
        }
      } else {
        console.log('⚠️ Using ultimate fallback - attempting to parse raw response as CSV');

        // Ultimate fallback: try to parse raw response as CSV
        if (result.success && result.response) {
          const response = result.response.trim();
          let csvParsed = false;

          if (response.includes(';')) {
            try {
              const lines = response.split('\n').map(line => line.trim()).filter(line => line);
              if (lines.length > 1) {
                const headers = lines[0].split(';').map(h => h.trim());
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
                  console.log(`✅ Parsed raw response CSV with ${csvData.length} rows, creating RISULTATO sheet`);
                  const sheet = xlsx.utils.json_to_sheet(csvData);
                  xlsx.utils.book_append_sheet(workbook, sheet, 'RISULTATO');
                  csvParsed = true;
                }
              }
            } catch (e) {
              console.log('Failed to parse raw response as CSV:', e.message);
            }
          }

          // If CSV parsing failed, create minimal info sheet
          if (!csvParsed) {
            console.log('⚠️ All parsing attempts failed, creating minimal info sheet');
            const fallbackData = [{
              'Row #': rowIndex + 1,
              ...result.rowData,
              'Status': 'No parseable data',
              'Note': 'Response could not be converted to table format'
            }];
            const sheet = xlsx.utils.json_to_sheet(fallbackData);
            xlsx.utils.book_append_sheet(workbook, sheet, 'RISULTATO');
          }
        }
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

// Helper function to create consolidated Excel from multiple Excel files
function createConsolidatedExcel(excelBuffers) {
  console.log('\n=== Starting Consolidated Excel Creation ===');
  console.log(`Processing ${excelBuffers.length} Excel files`);

  const consolidatedRows = [];

  // Process each Excel buffer
  excelBuffers.forEach((buffer, fileIndex) => {
    try {
      // Read the Excel file
      const workbook = xlsx.read(buffer, { type: 'buffer' });

      // Look for RISULTATI or RISULTATO sheet
      let sheetName = null;
      if (workbook.SheetNames.includes('RISULTATI')) {
        sheetName = 'RISULTATI';
      } else if (workbook.SheetNames.includes('RISULTATO')) {
        sheetName = 'RISULTATO';
      } else if (workbook.SheetNames.length > 0) {
        // Fallback to first sheet
        sheetName = workbook.SheetNames[0];
      }

      if (sheetName) {
        const sheet = workbook.Sheets[sheetName];
        const rows = xlsx.utils.sheet_to_json(sheet);

        console.log(`File ${fileIndex + 1}: Found ${rows.length} rows in sheet "${sheetName}"`);

        // Add all rows to consolidated array
        consolidatedRows.push(...rows);
      } else {
        console.log(`File ${fileIndex + 1}: No sheets found, skipping`);
      }
    } catch (error) {
      console.error(`Error processing file ${fileIndex + 1}:`, error.message);
    }
  });

  console.log(`Total consolidated rows: ${consolidatedRows.length}`);

  // Create new workbook with consolidated data
  const consolidatedWorkbook = xlsx.utils.book_new();
  const consolidatedSheet = xlsx.utils.json_to_sheet(consolidatedRows);
  xlsx.utils.book_append_sheet(consolidatedWorkbook, consolidatedSheet, 'TUTTI_I_RISULTATI');

  console.log('✅ Consolidated Excel created successfully');

  return consolidatedWorkbook;
}

// API endpoint to download all results as ZIP
app.post('/api/download-excel-zip', authenticateToken, async (req, res) => {
  try {
    const { results } = req.body;

    if (!results || !Array.isArray(results)) {
      return res.status(400).json({ error: 'Results data is required' });
    }

    const zip = new JSZip();
    const excelBuffers = [];

    // Create an Excel file for each result
    results.forEach((result, index) => {
      const workbook = createExcelFromResult(result, index);
      const excelBuffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
      const filename = generateFilename(result, index);
      zip.file(filename, excelBuffer);

      // Store buffer for consolidation
      excelBuffers.push(excelBuffer);
    });

    // Create consolidated Excel file
    const timestamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\..+/, '').replace('T', '_');
    const consolidatedWorkbook = createConsolidatedExcel(excelBuffers);
    const consolidatedBuffer = xlsx.write(consolidatedWorkbook, { type: 'buffer', bookType: 'xlsx' });
    const consolidatedFilename = `CONSOLIDATO_RISULTATI_${timestamp}.xlsx`;
    zip.file(consolidatedFilename, consolidatedBuffer);

    console.log(`📦 Added consolidated file: ${consolidatedFilename}`);

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
      // Check if this is the PA competenze format (optimized or old)
      if (data.competenze_flat || data.tabella_1_normativa_generale) {
        // New optimized flat format - create only RISULTATI
        if (data.competenze_flat && Array.isArray(data.competenze_flat)) {
          const sheet = xlsx.utils.json_to_sheet(data.competenze_flat);
          xlsx.utils.book_append_sheet(workbook, sheet, 'RISULTATI');
        } else {
          // Old format - combine all tables into RISULTATI
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

          const tutteLeRighe = [];
          for (const [sheetName, tableData] of Object.entries(tables)) {
            if (tableData && Array.isArray(tableData) && tableData.length > 0) {
              tutteLeRighe.push(...tableData);
            }
          }

          if (tutteLeRighe.length > 0) {
            const sheet = xlsx.utils.json_to_sheet(tutteLeRighe);
            xlsx.utils.book_append_sheet(workbook, sheet, 'RISULTATI');
          }
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
