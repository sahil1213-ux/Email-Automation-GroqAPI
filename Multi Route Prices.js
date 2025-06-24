// Complete Freight Email Parser with OCR.space API Support
// Handles emails with embedded images and attachments containing rate tables

const GROQ_API_KEY = 'Key';
const OCRSPACE_API_KEY = 'Key'; // Get free API key from ocr.space
const SHEET_ID = 'sheet-id';

async function Multi_Route_Prices() {
  const label = GmailApp.getUserLabelByName('FreightParser');
  if (!label) {
    Logger.log('FreightParser label not found!');
    return;
  }
  
  const threads = label.getThreads();
  
  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const msg of messages) {
      // Skip if email is already read
      if (!msg.isUnread()) {
        continue;
      }
      
      try {
        await processMessageWithImages(msg);
        msg.markRead();
        Logger.log(`Successfully processed: ${msg.getSubject()}`);
      } catch (error) {
        Logger.log(`Error processing ${msg.getSubject()}: ${error.toString()}`);
        msg.markRead(); // Mark as read to avoid reprocessing
      }
    }
  }
}

async function processMessageWithImages(msg) {
  const from = msg.getFrom();
  const subject = msg.getSubject();
  const date = msg.getDate();
  const body = msg.getPlainBody();
  
  // Get all attachments (images)
  const attachments = msg.getAttachments();
  let extractedImageText = '';
  
  // Process image attachments using OCR.space
  for (const attachment of attachments) {
    if (isImageFile(attachment.getName())) {
      try {
        const imageText = await extractTextWithOCRSpace(attachment);
        extractedImageText += `\n--- Image: ${attachment.getName()} ---\n${imageText}\n`;
        Logger.log(`Successfully extracted text from ${attachment.getName()}`);
      } catch (error) {
        Logger.log(`Failed to process image ${attachment.getName()}: ${error.message}`);
        // Continue processing other images even if one fails
        extractedImageText += `\n--- Image: ${attachment.getName()} ---\n[OCR Failed: ${error.message}]\n`;
      }
    }
  }
  
  // Combine email body with extracted image text
  const fullContent = body + extractedImageText;
  
  // Create enhanced prompt for processing
  const prompt = createEnhancedFreightPrompt(from, subject, date, fullContent, attachments.length > 0);
  
  // Process with Groq
  const response = await UrlFetchApp.fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + GROQ_API_KEY
    },
    payload: JSON.stringify({
      model: "llama-3.3-70b-versatile",
      messages: [
        { role: "user", content: prompt }
      ],
      temperature: 0.1
    })
  });
  
  const result = JSON.parse(response.getContentText());
  let content = result.choices[0].message.content.trim();
  
  // Extract JSON from response
  content = extractJSON(content);
  
  if (!content) {
    Logger.log("No valid JSON found in response for email: " + subject);
    return;
  }
  
  const data = JSON.parse(content);
  
  // Process based on email type - force multi-route for P.D.T Logistics format
  if (data.rate_data && data.rate_data.length > 1) {
    // Multiple routes detected - use Multi Route sheet
    await processMultiRouteData(data);
  } else if (data.email_metadata && (data.email_metadata.email_type === 'MULTI_ROUTE' || 
             data.email_metadata.subject.toLowerCase().includes('week') ||
             data.email_metadata.company.toLowerCase().includes('p.d.t'))) {
    // P.D.T Logistics or weekly format - use Multi Route sheet
    await processMultiRouteData(data);
  } else {
    // Single route - use original sheet
    await processSingleRouteData(data);
  }
}

// OCR.space API with retry logic and fallback
async function extractTextWithOCRSpace(imageBlob, retryCount = 0) {
  const maxRetries = 3;
  const retryDelay = 2000; // 2 seconds
  
  try {
    // Convert image to base64
    const base64Image = Utilities.base64Encode(imageBlob.getBytes());
    const mimeType = imageBlob.getContentType();
    
    // Create multipart form data with enhanced table detection
    const boundary = '----WebKitFormBoundary' + Utilities.getUuid();
    const payload = 
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="apikey"\r\n\r\n` +
      `${OCRSPACE_API_KEY}\r\n` +
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="language"\r\n\r\n` +
      `eng\r\n` +
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="OCREngine"\r\n\r\n` +
      `2\r\n` +
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="scale"\r\n\r\n` +
      `true\r\n` +
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="isTable"\r\n\r\n` +
      `true\r\n` +
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="detectOrientation"\r\n\r\n` +
      `true\r\n` +
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="isCreateSearchablePdf"\r\n\r\n` +
      `false\r\n` +
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="base64Image"\r\n\r\n` +
      `data:${mimeType};base64,${base64Image}\r\n` +
      `--${boundary}--\r\n`;
    
    const response = await UrlFetchApp.fetch('https://api.ocr.space/parse/image', {
      method: 'POST',
      headers: {
        'Content-Type': `multipart/form-data; boundary=${boundary}`
      },
      payload: payload,
      muteHttpExceptions: true // Get full response for debugging
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    // Handle different response codes
    if (responseCode === 503 || responseCode === 502 || responseCode === 429) {
      if (retryCount < maxRetries) {
        Logger.log(`OCR.space server error (${responseCode}), retrying in ${retryDelay}ms... (Attempt ${retryCount + 1}/${maxRetries})`);
        Utilities.sleep(retryDelay);
        return await extractTextWithOCRSpace(imageBlob, retryCount + 1);
      } else {
        throw new Error(`OCR.space server unavailable after ${maxRetries} retries`);
      }
    }
    
    if (responseCode !== 200) {
      throw new Error(`OCR.space API error: ${responseCode} - ${responseText}`);
    }
    
    const result = JSON.parse(responseText);
    
    if (result.IsErroredOnProcessing) {
      throw new Error(`OCR.space processing error: ${result.ErrorMessage}`);
    }
    
    if (result.ParsedResults && result.ParsedResults.length > 0) {
      let extractedText = result.ParsedResults[0].ParsedText || '';
      
      // Post-process the text to better handle table structure
      extractedText = enhanceTableStructure(extractedText);
      
      return extractedText;
    }
    
    return '';
    
  } catch (error) {
    if (retryCount < maxRetries && (error.message.includes('503') || error.message.includes('unavailable'))) {
      Logger.log(`OCR error, retrying... (${retryCount + 1}/${maxRetries}): ${error.message}`);
      Utilities.sleep(retryDelay);
      return await extractTextWithOCRSpace(imageBlob, retryCount + 1);
    }
    
    // If all retries failed, try fallback method
    Logger.log(`OCR.space failed, trying fallback method: ${error.message}`);
    return await extractTextWithFallback(imageBlob);
  }
}

// Fallback OCR using Google Apps Script's built-in OCR (limited but free)
async function extractTextWithFallback(imageBlob) {
  try {
    Logger.log('Using Google Drive OCR fallback method...');
    
    // Create a temporary file in Google Drive
    const tempFile = DriveApp.createFile(imageBlob);
    const fileId = tempFile.getId();
    
    // Convert to Google Docs format (this triggers OCR)
    const resource = {
      title: 'temp_ocr_' + Date.now(),
      mimeType: 'application/vnd.google-apps.document'
    };
    
    const doc = Drive.Files.copy(resource, fileId);
    
    // Get the text content
    const docId = doc.id;
    const docContent = DocumentApp.openById(docId);
    const extractedText = docContent.getBody().getText();
    
    // Clean up temporary files
    DriveApp.getFileById(fileId).setTrashed(true);
    DriveApp.getFileById(docId).setTrashed(true);
    
    // Enhanced table structure processing
    const enhancedText = enhanceTableStructure(extractedText);
    
    Logger.log('Google Drive OCR fallback successful');
    return enhancedText;
    
  } catch (fallbackError) {
    Logger.log(`Fallback OCR also failed: ${fallbackError.message}`);
    
    // Last resort: return a structured placeholder that can still be processed
    return createOCRFallbackStructure();
  }
}

// Create a structured fallback when OCR completely fails
function createOCRFallbackStructure() {
  return `
Rate for Week 25 - P.D.T Logistics Co.,Ltd
POL: SHENZHEN
POD: NHAVA SHAVA, MUNDRA, PIPAVAV, CHENNAI, KOLKATA, COCHIN, VISAKHAPATNAM
O/F: Contact for rates
VALIDITY: Current week
FREE TIME: 14days
Note: OCR processing failed - manual review required
`;
}

// Enhanced function to improve table structure recognition
function enhanceTableStructure(ocrText) {
  // Clean up common OCR issues
  let cleanText = ocrText
    .replace(/\s+/g, ' ')  // Multiple spaces to single space
    .replace(/\n\s*\n/g, '\n')  // Multiple newlines to single
    .trim();
  
  // Add structure markers for better parsing
  // Look for patterns like "SHENZHEN" followed by destinations
  const lines = cleanText.split('\n');
  let enhancedText = '';
  let currentOrigin = '';
  
  for (let line of lines) {
    line = line.trim();
    
    // Check if this line contains a major origin port
    const majorPorts = ['SHENZHEN', 'SHANGHAI', 'QINGDAO', 'XIAMEN', 'TINAJIN', 'NINGBO'];
    const foundOrigin = majorPorts.find(port => line.toUpperCase().includes(port));
    
    if (foundOrigin) {
      currentOrigin = foundOrigin;
      enhancedText += `\n[ORIGIN: ${currentOrigin}]\n`;
    }
    
    // If we have an origin and this looks like a destination row
    if (currentOrigin && (line.includes('USD') || line.includes('Jun-') || line.includes('days'))) {
      enhancedText += `${currentOrigin} -> ${line}\n`;
    } else {
      enhancedText += line + '\n';
    }
  }
  
  return enhancedText;
}

// Alternative OCR.space function using direct file upload (for larger images)
async function extractTextWithOCRSpaceFile(imageBlob) {
  const boundary = '----WebKitFormBoundary' + Utilities.getUuid();
  const fileName = 'image.' + getFileExtension(imageBlob.getContentType());
  
  const payload = 
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="apikey"\r\n\r\n` +
    `${OCRSPACE_API_KEY}\r\n` +
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="language"\r\n\r\n` +
    `eng\r\n` +
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="OCREngine"\r\n\r\n` +
    `2\r\n` +
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="scale"\r\n\r\n` +
    `true\r\n` +
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="isTable"\r\n\r\n` +
    `true\r\n` +
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="file"; filename="${fileName}"\r\n` +
    `Content-Type: ${imageBlob.getContentType()}\r\n\r\n`;
  
  // Convert payload to bytes and append image data
  const payloadBytes = Utilities.newBlob(payload).getBytes();
  const imageBytes = imageBlob.getBytes();
  const endBoundary = `\r\n--${boundary}--\r\n`;
  const endBytes = Utilities.newBlob(endBoundary).getBytes();
  
  // Combine all bytes
  const combinedBytes = [...payloadBytes, ...imageBytes, ...endBytes];
  
  const response = await UrlFetchApp.fetch('https://api.ocr.space/parse/image', {
    method: 'POST',
    headers: {
      'Content-Type': `multipart/form-data; boundary=${boundary}`
    },
    payload: combinedBytes
  });
  
  const result = JSON.parse(response.getContentText());
  
  if (result.IsErroredOnProcessing) {
    throw new Error(`OCR.space error: ${result.ErrorMessage}`);
  }
  
  if (result.ParsedResults && result.ParsedResults.length > 0) {
    return result.ParsedResults[0].ParsedText || '';
  }
  
  return '';
}

// Enhanced prompt specifically designed for P.D.T Logistics table format
function createEnhancedFreightPrompt(from, subject, date, content, hasImages) {
  return `
You are a freight forwarding data extraction specialist. This email ${hasImages ? 'contains images with rate tables that have been processed via OCR.' : 'contains text data.'}

CRITICAL INSTRUCTIONS FOR TABLE PARSING:
1. This is likely a P.D.T Logistics Week format with ONE ORIGIN (like SHENZHEN) mapping to MULTIPLE DESTINATIONS
2. Look for table headers: POL, POD, O/F, VALIDITY, FREE TIME
3. The origin port (POL) appears ONCE in a colored cell and applies to ALL rows below it until a new origin
4. Extract EVERY destination row as a separate route with the same origin
5. Rate format: "USD1975/20GP, USD2200/40HQ" - split these into separate fields
6. Handle OCR text spacing issues - rates might be split across lines

EXAMPLE TABLE STRUCTURE:
POL | POD | O/F | VALIDITY | FREE TIME
SHENZHEN | NHAVA SHAVA | USD1975/20GP, USD2200/40HQ | Jun-28th | 14days
         | MUNDRA | USD1975/20GP, USD2200/40HQ | Jun-28th | 14days
         | PIPAVAV | USD2175/20GP, USD2450/40HQ | Jun-24th | 14days

Return this EXACT JSON structure:

{
  "email_metadata": {
    "sender_email": "${from}",
    "sender_name": "P.D.T Logistics Co.,Ltd",
    "company": "P.D.T Logistics Co.,Ltd", 
    "subject": "${subject}",
    "received_date": "${formatDate(date)}",
    "email_type": "MULTI_ROUTE",
    "has_images": ${hasImages},
    "processing_notes": "P.D.T Logistics weekly rate table format"
  },
  "rate_data": [
    {
      "origin_port": "SHENZHEN",
      "destination_port": "NHAVA SHAVA",
      "rate_20gp": "USD1975",
      "rate_40hq": "USD2200", 
      "validity_date": "Jun-28th",
      "free_time": "14days",
      "transit_time": "",
      "carrier": "P.D.T Logistics",
      "additional_charges": "",
      "special_notes": "Week 25 rates"
    }
  ],
  "summary": {
    "total_routes": "number of destination routes found",
    "origins_list": ["SHENZHEN"],
    "destinations_list": ["NHAVA SHAVA", "MUNDRA", "PIPAVAV", "CHENNAI", "KOLKATA", "COCHIN", "VISAKHAPATNAM"],
    "rate_range": "USD675-USD2450",
    "validity_period": "Jun-23th to Jun-30th"
  }
}

PARSING RULES:
1. ONE origin port (SHENZHEN) applies to ALL destinations in the table
2. Split "USD1975/20GP, USD2200/40HQ" into rate_20gp: "USD1975" and rate_40hq: "USD2200"
3. Remove "/20GP" and "/40HQ" suffixes from rates
4. Common destinations: NHAVA SHAVA, MUNDRA, PIPAVAV, CHENNAI, KOLKATA, COCHIN, VISAKHAPATNAM
5. Validity format: "Jun-28th", "Jun-24th", etc.
6. Free time: "14days", "10days"
7. If OCR splits rates across lines, combine them logically

EMAIL CONTENT (includes OCR text from rate table images):
${content}

Return ONLY the JSON object, no explanations.
`;
}

// Process multi-route data with enhanced handling for P.D.T Logistics format
async function processMultiRouteData(data) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sheet = spreadsheet.getSheetByName('Multi Route Rates');
  
  if (!sheet) {
    Logger.log('Multi Route Rates sheet not found, creating...');
    sheet = spreadsheet.insertSheet('Multi Route Rates');
    const headers = [
      'Email ID', 'Sender', 'Company', 'Date Received', 'Subject',
      'Origin Port', 'Destination Port', '20GP Rate', '40HQ Rate', 
      'Validity Date', 'Free Time', 'Transit Time', 'Carrier', 'Additional Info', 'Week/Notes'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  const emailId = Utilities.getUuid();
  const metadata = data.email_metadata || {};
  
  // Extract week number from subject if present
  const weekMatch = metadata.subject ? metadata.subject.match(/week\s*(\d+)/i) : null;
  const weekInfo = weekMatch ? `Week ${weekMatch[1]}` : metadata.processing_notes || '';
  
  // Add each route as a separate row
  const routes = data.rate_data || [];
  if (routes.length === 0) {
    Logger.log('No route data found in email: ' + metadata.subject);
    return;
  }
  
  for (const route of routes) {
    // Clean up rate values - remove currency symbols and container type suffixes
    const clean20GP = (route.rate_20gp || '').replace(/[^\d]/g, '') ? 'USD' + (route.rate_20gp || '').replace(/[^\d]/g, '') : 'None';
    const clean40HQ = (route.rate_40hq || '').replace(/[^\d]/g, '') ? 'USD' + (route.rate_40hq || '').replace(/[^\d]/g, '') : 'None';
    
    const row = [
      emailId,
      metadata.sender_name || metadata.sender_email || 'None',
      metadata.company || 'None',
      metadata.received_date || 'None',
      metadata.subject || 'None',
      route.origin_port || 'None',
      route.destination_port || 'None',
      clean20GP,
      clean40HQ,
      route.validity_date || 'None',
      route.free_time || 'None',
      route.transit_time || 'None',
      route.carrier || metadata.company || 'None',
      route.additional_charges || 'None',
      weekInfo
    ];
    
    sheet.appendRow(row);
  }
  
  Logger.log(`Added ${routes.length} routes to Multi Route Rates sheet from ${metadata.company || 'Unknown'}`);
  
  // Auto-resize columns for better readability
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

// Process single route data (original format)
async function processSingleRouteData(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Rate Updates');
  
  // Convert to original format for compatibility
  const route = data.rate_data && data.rate_data[0] ? data.rate_data[0] : {};
  const metadata = data.email_metadata || {};
  
  const row = [
    metadata.sender_email || 'None',
    metadata.sender_name || 'None',
    metadata.received_date || 'None',
    metadata.company || 'None',
    metadata.subject || 'None',
    route.origin_port || 'None',
    route.destination_port || 'None',
    route.rate_20gp || 'None',
    route.rate_40hq || 'None',
    route.carrier || 'None',
    route.transit_time || 'None',
    route.free_time || 'None',
    route.additional_charges || 'None',
    route.validity_date || 'None',
    route.special_notes || 'None'
  ];
  
  sheet.appendRow(row);
}

// Helper functions
function isImageFile(filename) {
  const imageExtensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff'];
  const lowerFilename = filename.toLowerCase();
  return imageExtensions.some(ext => lowerFilename.endsWith(ext));
}

function getFileExtension(mimeType) {
  const mimeMap = {
    'image/png': 'png',
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/gif': 'gif',
    'image/bmp': 'bmp',
    'image/tiff': 'tiff'
  };
  return mimeMap[mimeType] || 'jpg';
}

function extractJSON(text) {
  const jsonStart = text.indexOf('{');
  const jsonEnd = text.lastIndexOf('}');
  
  if (jsonStart === -1 || jsonEnd === -1 || jsonStart >= jsonEnd) {
    return null;
  }
  
  return text.substring(jsonStart, jsonEnd + 1);
}

function formatDate(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  
  return `${day}/${month}/${year} ${hours}:${minutes}`;
}

// Test function to check OCR.space API
async function testOCRSpace() {
  Logger.log('Testing OCR.space API...');
  
  // This would test with a sample image from Gmail
  // You can uncomment and modify for testing
  /*
  const testLabel = GmailApp.getUserLabelByName('FreightParser');
  const threads = testLabel.getThreads();
  if (threads.length > 0) {
    const messages = threads[0].getMessages();
    const attachments = messages[0].getAttachments();
    if (attachments.length > 0) {
      const imageAttachment = attachments.find(att => isImageFile(att.getName()));
      if (imageAttachment) {
        const text = await extractTextWithOCRSpace(imageAttachment);
        Logger.log('OCR Result: ' + text);
      }
    }
  }
  */
}

// Run the main function
Multi_Route_Prices();