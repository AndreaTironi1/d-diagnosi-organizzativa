# Prompt Executor Web App

A web application that executes prompts with multiple variables using the Claude API, with Excel batch processing support.

## Features

- **User Authentication**: Secure login system with session management
- **Excel Integration**: Upload Excel files and process multiple rows in batch
- **Column Mapping**: Map Excel columns to prompt variables using `[column_name]` syntax
- **Batch Processing**: Execute prompts for all rows automatically
- **Real-time Progress**: Track processing status for each row
- **Token Usage Tracking**: Monitor API usage with detailed statistics
- **Results Export**: Download batch results as JSON
- **Clean, Responsive UI**: Professional interface with step-by-step workflow

## Prerequisites

- Node.js (v18 or higher)
- Anthropic API key ([Get one here](https://console.anthropic.com/))

## Installation

1. Install dependencies:
```bash
npm install
```

2. Create a `.env` file from the example:
```bash
copy .env.example .env
```

3. Edit `.env` and add your Anthropic API key:
```
ANTHROPIC_API_KEY=your_actual_api_key_here
```

## Usage

1. Start the server:
```bash
npm start
```

2. Open your browser and navigate to:
```
http://localhost:3000
```

3. Login with credentials:
   - **Username**: `dasein`
   - **Password**: `Donatella2025!@`

4. Follow the 3-step workflow:

   **Step 1: Upload Excel File**
   - Click the upload area and select your Excel file (.xlsx or .xls)
   - The app will display detected columns and row count
   - Column names will be shown in `[column_name]` format

   **Step 2: Enter Prompt Template**
   - Write your prompt using `[column_name]` syntax to reference Excel columns
   - Example: `Write a summary about [Product_Name] which costs [Price]`

   **Step 3: Execute Batch Processing**
   - Click "Execute for All Rows"
   - The app processes each row sequentially
   - View results for each row with success/error status
   - Download complete results as JSON

## Excel Format Example

Your Excel file should have column headers in the first row:

| Product_Name | Price | Category | Description |
|-------------|-------|----------|-------------|
| Laptop | $999 | Electronics | High-performance laptop |
| Mouse | $25 | Accessories | Wireless mouse |
| Keyboard | $75 | Accessories | Mechanical keyboard |

## Example Prompts

**Example 1: Product Descriptions**
```
Write a compelling product description for [Product_Name] in the [Category] category.
The product costs [Price] and its key feature is: [Description]
Make it engaging and highlight the value proposition.
```

**Example 2: Email Personalization**
```
Write a personalized email to [Customer_Name] about their recent purchase of [Product].
Thank them and suggest a related product from our [Category] category.
Keep the tone [Tone].
```

**Example 3: Content Generation**
```
Create a social media post about [Topic] for [Platform].
The target audience is [Audience] and the tone should be [Style].
Include relevant hashtags.
```

**Example 4: Data Analysis Summary**
```
Analyze this data point: [Metric_Name] has a value of [Value] for [Date].
Compare it to the previous period [Previous_Value] and provide insights.
```

## API Endpoints

### POST /api/login
Authenticates a user and creates a session.

**Request:**
```json
{
  "username": "dasein",
  "password": "Donatella2025!@"
}
```

**Response:**
```json
{
  "success": true,
  "message": "Login successful"
}
```

### POST /api/logout
Destroys the current session and logs out the user.

**Response:**
```json
{
  "success": true,
  "message": "Logged out successfully"
}
```

### POST /api/upload-excel
Uploads and parses an Excel file.

**Request:** Multipart form data with file field

**Response:**
```json
{
  "success": true,
  "columns": ["Product_Name", "Price", "Category"],
  "rowCount": 10,
  "data": [
    {"Product_Name": "Laptop", "Price": "$999", "Category": "Electronics"},
    ...
  ]
}
```

### POST /api/execute-batch
Executes a prompt for all rows in the Excel data.

**Request:**
```json
{
  "prompt": "Write about [Product_Name] which costs [Price]",
  "excelData": [
    {"Product_Name": "Laptop", "Price": "$999"},
    {"Product_Name": "Mouse", "Price": "$25"}
  ]
}
```

**Response:**
```json
{
  "success": true,
  "results": [
    {
      "rowIndex": 0,
      "rowData": {"Product_Name": "Laptop", "Price": "$999"},
      "processedPrompt": "Write about Laptop which costs $999",
      "response": "Claude's response...",
      "usage": {"inputTokens": 50, "outputTokens": 200},
      "success": true
    }
  ],
  "totalUsage": {"inputTokens": 100, "outputTokens": 400},
  "successCount": 2,
  "errorCount": 0
}
```

### POST /api/execute
Executes a single prompt with variables (legacy endpoint).

**Request:**
```json
{
  "prompt": "Write about [topic]",
  "variables": {"topic": "nature"}
}
```

### POST /api/parse-variables
Extracts variable names from a prompt template.

**Request:**
```json
{
  "prompt": "Hello [name], you are [age] years old"
}
```

**Response:**
```json
{
  "variables": ["name", "age"]
}
```

## Configuration

Edit `.env` file to configure:

- `ANTHROPIC_API_KEY`: Your Anthropic API key (required)
- `PORT`: Server port (optional, defaults to 3000)

## Project Structure

```
.
├── server.js           # Express server with Claude API integration and authentication
├── public/
│   ├── index.html      # Main application interface
│   └── login.html      # Login page
├── package.json        # Dependencies and scripts
├── .env.example        # Environment variables template
├── .env                # Your actual environment variables (create this)
└── README.md           # This file
```

## Notes

- **Authentication**: Login credentials are hardcoded in server.js for simplicity
- **Session Management**: Sessions last 24 hours by default
- The app uses Claude Sonnet 4.5 model by default
- Maximum token output is set to 4096 per request
- API key is never exposed to the frontend
- All API calls are made server-side for security
- Batch processing is sequential (one row at a time) to avoid rate limits
- Variable syntax changed from `{variable}` to `[column_name]` for Excel compatibility
- Results can be downloaded as JSON for further processing

## Troubleshooting

**"ANTHROPIC_API_KEY not configured" error:**
- Make sure you created a `.env` file
- Verify your API key is correctly set in `.env`
- Restart the server after changing `.env`

**Port already in use:**
- Change the `PORT` in `.env` to a different value (e.g., 3001)
- Or stop the application using that port

**Excel file not parsing correctly:**
- Ensure your Excel file has headers in the first row
- Use simple column names without special characters
- Save the file as .xlsx format for best compatibility

**Batch processing errors:**
- Check that all column names in your prompt match Excel columns exactly
- Use square brackets `[Column_Name]` not curly braces
- Large batches may take time - be patient and don't refresh the page

## License

ISC
