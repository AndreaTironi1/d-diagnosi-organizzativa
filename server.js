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

  // Try to parse JSON from Claude response
  let parsedData = null;
  let jsonFound = false;

  if (result.success && result.response) {
    try {
      const jsonMatch = result.response.match(/```json\s*([\s\S]*?)\s*```/);
      if (jsonMatch) {
        parsedData = JSON.parse(jsonMatch[1]);
        jsonFound = true;
      } else {
        // Try to parse the entire response as JSON
        try {
          parsedData = JSON.parse(result.response);
          jsonFound = true;
        } catch (e) {
          // Not valid JSON, that's ok - will use text fallback
        }
      }
    } catch (e) {
      console.log('Failed to parse JSON for row', rowIndex, e.message);
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
  } else {
    // Fallback: check if parsedData is an array, if so explode it into rows
    if (parsedData && Array.isArray(parsedData) && parsedData.length > 0) {
      const sheet = xlsx.utils.json_to_sheet(parsedData);
      xlsx.utils.book_append_sheet(workbook, sheet, 'Data');
    } else if (parsedData && typeof parsedData === 'object' && !Array.isArray(parsedData)) {
      // If parsedData is an object, try to find arrays within it and create sheets
      let hasSheets = false;
      for (const [key, value] of Object.entries(parsedData)) {
        if (Array.isArray(value) && value.length > 0) {
          const sheetName = key.substring(0, 31); // Excel sheet name limit
          const sheet = xlsx.utils.json_to_sheet(value);
          xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
          hasSheets = true;
        }
      }

      // If no arrays found, create a sheet with the object properties
      if (!hasSheets) {
        const sheet = xlsx.utils.json_to_sheet([parsedData]);
        xlsx.utils.book_append_sheet(workbook, sheet, 'Data');
      }
    } else {
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
