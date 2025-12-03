import express from 'express';
import Anthropic from '@anthropic-ai/sdk';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import multer from 'multer';
import xlsx from 'xlsx';
import session from 'express-session';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// Hardcoded credentials
const CREDENTIALS = {
  username: 'dasein',
  password: 'Donatella2025!@'
};

// Configure multer for file uploads
const upload = multer({ storage: multer.memoryStorage() });

// Session configuration
app.use(session({
  secret: 'prompt-executor-secret-key-2025',
  resave: false,
  saveUninitialized: false,
  cookie: {
    maxAge: 24 * 60 * 60 * 1000 // 24 hours
  }
}));

app.use(express.json());

// Authentication middleware
function requireAuth(req, res, next) {
  if (req.session && req.session.authenticated) {
    return next();
  }
  res.status(401).json({ error: 'Authentication required' });
}

// Serve login page for unauthenticated users
app.use((req, res, next) => {
  // Allow access to login page and login API
  if (req.path === '/login.html' || req.path === '/api/login' || req.path === '/api/logout') {
    return next();
  }

  // If not authenticated and trying to access protected resource, redirect to login
  if (!req.session || !req.session.authenticated) {
    if (req.path.startsWith('/api/')) {
      return res.status(401).json({ error: 'Authentication required' });
    }
    return res.redirect('/login.html');
  }

  next();
});

app.use(express.static('public'));

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
app.post('/api/execute', async (req, res) => {
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
      max_tokens: 4096,
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
    req.session.authenticated = true;
    req.session.username = username;
    res.json({ success: true, message: 'Login successful' });
  } else {
    res.status(401).json({ success: false, error: 'Invalid username or password' });
  }
});

// Logout endpoint
app.post('/api/logout', (req, res) => {
  req.session.destroy((err) => {
    if (err) {
      return res.status(500).json({ error: 'Failed to logout' });
    }
    res.json({ success: true, message: 'Logged out successfully' });
  });
});

// API endpoint to parse variables from prompt
app.post('/api/parse-variables', (req, res) => {
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
app.post('/api/upload-excel', upload.single('file'), (req, res) => {
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
app.post('/api/execute-batch', async (req, res) => {
  try {
    const { prompt, excelData } = req.body;

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

    const results = [];

    // Process each row
    for (let i = 0; i < excelData.length; i++) {
      const row = excelData[i];

      // Replace variables with values from current row
      const processedPrompt = replaceVariables(prompt, row);

      try {
        // Call Claude API
        const message = await anthropic.messages.create({
          model: 'claude-sonnet-4-5-20250929',
          max_tokens: 4096,
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

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  if (!process.env.ANTHROPIC_API_KEY) {
    console.warn('⚠️  Warning: ANTHROPIC_API_KEY not set in .env file');
  }
});
