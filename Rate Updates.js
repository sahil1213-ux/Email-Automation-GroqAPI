async function Rate_Updates() {
    const label = GmailApp.getUserLabelByName('FreightParser');
    const threads = label.getThreads();
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const msg of messages) {
        // Skip if email is already read
        if (!msg.isUnread()) {
          continue;
        }
        
        const from = msg.getFrom();
        const subject = msg.getSubject();
        const date = msg.getDate();
        const body = msg.getPlainBody();
        
        const prompt = `
  Extract freight forwarding data from the email below and return ONLY a valid JSON object with no additional text or explanation.
  
  Required JSON format (use exactly these keys):
  {
    "Sender Email": "<email address>",
    "Sender Name": "<sender's full name>",
    "Received Date & Time": "<dd/mm/yyyy hh:mm>",
    "Company": "<company name>",
    "Subject": "<email subject line>",
    "Origin": "<Port of Loading (POL)>",
    "Destination": "<Port of Discharge (POD)>",
    "20 ft Rate($)": "<USD rate for 20ft container>",
    "40 ft Rate($)": "<USD rate for 40ft container>",
    "Carrier": "<carrier name>",
    "Transit Time": "<number of days or range>",
    "Free Time": "<free time in days>",
    "Additional Charges": "<describe extra charges or say None>",
    "Validity / Urgency": "<any urgency like first come first serve>",
    "Full Body / Notes": "<summary of email body or notes or None>"
  }
  
  If any field is not present, use "None" as the value.
  Return ONLY the JSON object, no other text.
  
  Email details:
  From: ${from}
  Subject: ${subject}
  Date: ${date}
  Body: ${body}
  `;
  
        try {
          const response = UrlFetchApp.fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "post",
            contentType: "application/json",
            headers: {
              Authorization: "Bearer " + 'GROK_API'
            },
            payload: JSON.stringify({
              model: "llama-3.3-70b-versatile",
              messages: [
                { role: "user", content: prompt }
              ],
              temperature: 0.1 // Lower temperature for more consistent output
            })
          });
          
          const result = JSON.parse(response.getContentText());
          let content = result.choices[0].message.content.trim();
          
          // Clean up the response to extract JSON
          content = extractJSON(content);
          
          if (!content) {
            Logger.log("No valid JSON found in response for email: " + subject);
            continue;
          }
          
          const data = JSON.parse(content);
          
          // Validate that we have the expected structure
          const requiredKeys = [
            "Sender Email", "Sender Name", "Received Date & Time", "Company", 
            "Subject", "Origin", "Destination", "20 ft Rate($)", "40 ft Rate($)", 
            "Carrier", "Transit Time", "Free Time", "Additional Charges", 
            "Validity / Urgency", "Full Body / Notes"
          ];
          
          // Fill in missing keys with "None"
          requiredKeys.forEach(key => {
            if (!data.hasOwnProperty(key)) {
              data[key] = "None";
            }
          });
          
          const row_content = [
            data["Sender Email"] || from, // Fallback to actual email if not extracted
            data["Sender Name"],
            data["Received Date & Time"] || formatDate(date), // Fallback to actual date
            data["Company"],
            data["Subject"] || subject, // Fallback to actual subject
            data["Origin"],
            data["Destination"],
            data["20 ft Rate($)"],
            data["40 ft Rate($)"],
            data["Carrier"],
            data["Transit Time"],
            data["Free Time"],
            data["Additional Charges"],
            data["Validity / Urgency"],
            data["Full Body / Notes"]
          ];
          
          const sheet = SpreadsheetApp.openById('1ZGzg1M8oQT5AFT7Vkpq5fL9vbay-IyQHNui8eG5cRX4')
                                      .getSheetByName('Rate Updates');
          sheet.appendRow(row_content);
          
          // Mark email as read after successful processing
          msg.markRead();
          
          Logger.log("Successfully processed email: " + subject);
          
        } catch (e) {
          Logger.log("Error processing email '" + subject + "': " + e.toString());
          Logger.log("Raw API response: " + (result ? result.choices[0].message.content : 'No response'));
          
          // Add a fallback row with basic info
          const fallbackRow = [
            from, "None", formatDate(date), "None", subject, 
            "None", "None", "None", "None", "None", "None", 
            "None", "None", "None", "Error parsing: " + e.toString()
          ];
          
          try {
            const sheet = SpreadsheetApp.openById('1ZGzg1M8oQT5AFT7Vkpq5fL9vbay-IyQHNui8eG5cRX4')
                                        .getSheetByName('Rate Updates');
            sheet.appendRow(fallbackRow);
            
            // Mark as read even if parsing failed (to avoid reprocessing)
            msg.markRead();
          } catch (sheetError) {
            Logger.log("Could not add fallback row: " + sheetError.toString());
          }
        }
      }
    }
  }
  
  // Helper function to extract JSON from mixed content
  function extractJSON(text) {
    // Look for JSON object boundaries
    const jsonStart = text.indexOf('{');
    const jsonEnd = text.lastIndexOf('}');
    
    if (jsonStart === -1 || jsonEnd === -1 || jsonStart >= jsonEnd) {
      return null;
    }
    
    return text.substring(jsonStart, jsonEnd + 1);
  }
  
  // Helper function to format date
  function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    
    return `${day}/${month}/${year} ${hours}:${minutes}`;
  }
  
  // Run the function
  Rate_Updates();