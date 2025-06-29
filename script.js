/**
 * Google Apps Script for Excel Data Viewer
 * This script fetches data from Google Sheets and returns it as JSON
 */

// Main function that handles GET requests
function doGet() {
  try {
    // Get the active spreadsheet and specific sheet named "graph"
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("graph");
    
    // Check if the "graph" sheet exists
    if (!sheet) {
      return createErrorResponse('Sheet "graph" not found. Please create a sheet named "graph" with the required columns.');
    }
    
    // Get all data from the sheet
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    
    // If no data, return empty array
    if (data.length <= 1) {
      return createJsonResponse([]);
    }
    
    // Get headers (first row)
    const headers = data[0];
    
    // Get data rows (skip header)
    const rows = data.slice(1);
    
    // Convert rows to JSON objects
    const jsonData = rows.map((row, index) => {
      return {
        orderId: row[0] || `ORD-${String(index + 1).padStart(3, '0')}`,
        projectId: row[1] || `PRJ-${String(index + 1).padStart(3, '0')}`,
        projectName: row[2] || 'Unnamed Project',
        projectType: row[3] || 'Other',
        region: row[4] || 'Unknown',
        price: parseFloat(row[5]) || 0,
        date: formatDate(row[6]) || new Date().toISOString().split('T')[0]
      };
    });
    
    // Return JSON response
    return createJsonResponse(jsonData);
    
  } catch (error) {
    console.error('Error in doGet:', error);
    return createErrorResponse('Failed to fetch data: ' + error.message);
  }
}

// Function to handle POST requests (for adding new data)
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    // Validate required fields
    if (!data.projectName || !data.projectType || !data.region || !data.price) {
      return createErrorResponse('Missing required fields');
    }
    
    // Get the "graph" sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("graph");
    
    // Check if the "graph" sheet exists
    if (!sheet) {
      return createErrorResponse('Sheet "graph" not found. Please create a sheet named "graph" first.');
    }
    
    // Generate new Order ID and Project ID
    const lastRow = sheet.getLastRow();
    const orderId = `ORD-${String(lastRow).padStart(3, '0')}`;
    const projectId = `PRJ-${String(lastRow).padStart(3, '0')}`;
    
    // Prepare row data
    const rowData = [
      orderId,
      projectId,
      data.projectName,
      data.projectType,
      data.region,
      data.price,
      new Date().toISOString().split('T')[0] // Current date
    ];
    
    // Add new row
    sheet.appendRow(rowData);
    
    return createJsonResponse({ 
      success: true, 
      message: 'Data added successfully',
      orderId: orderId,
      projectId: projectId
    });
    
  } catch (error) {
    console.error('Error in doPost:', error);
    return createErrorResponse('Failed to add data: ' + error.message);
  }
}

// Function to create JSON response
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Function to create error response
function createErrorResponse(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ error: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Function to format date
function formatDate(dateValue) {
  if (!dateValue) return null;
  
  try {
    const date = new Date(dateValue);
    return date.toISOString().split('T')[0];
  } catch (error) {
    return null;
  }
}

// Function to get data by specific criteria
function getDataByFilter(criteria) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("graph");
  
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.filter(row => {
    // Filter by project type
    if (criteria.projectType && row[3] !== criteria.projectType) {
      return false;
    }
    
    // Filter by region
    if (criteria.region && row[4] !== criteria.region) {
      return false;
    }
    
    // Filter by date range
    if (criteria.startDate && criteria.endDate) {
      const rowDate = new Date(row[6]);
      const startDate = new Date(criteria.startDate);
      const endDate = new Date(criteria.endDate);
      
      if (rowDate < startDate || rowDate > endDate) {
        return false;
      }
    }
    
    return true;
  });
}

// Function to get statistics
function getStatistics() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("graph");
  
  if (!sheet) {
    return {
      totalOrders: 0,
      totalRevenue: 0,
      avgPrice: 0,
      projectTypeCounts: {},
      regionCounts: {}
    };
  }
  
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  
  const totalOrders = rows.length;
  const totalRevenue = rows.reduce((sum, row) => sum + (parseFloat(row[5]) || 0), 0);
  const avgPrice = totalOrders > 0 ? totalRevenue / totalOrders : 0;
  
  // Count by project type
  const projectTypeCounts = {};
  rows.forEach(row => {
    const type = row[3] || 'Other';
    projectTypeCounts[type] = (projectTypeCounts[type] || 0) + 1;
  });
  
  // Count by region
  const regionCounts = {};
  rows.forEach(row => {
    const region = row[4] || 'Unknown';
    regionCounts[region] = (regionCounts[region] || 0) + 1;
  });
  
  return {
    totalOrders,
    totalRevenue,
    avgPrice,
    projectTypeCounts,
    regionCounts
  };
}

// Function to handle OPTIONS requests (for CORS)
function doOptions() {
  return ContentService.createTextOutput('');
}

// Test function to verify the script is working
function testScript() {
  try {
    const result = doGet();
    console.log('Script test successful');
    console.log('Response:', result.getContent());
    return true;
  } catch (error) {
    console.error('Script test failed:', error);
    return false;
  }
}

// Function to set up the "graph" sheet with headers if it's empty
function setupSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("graph");
  
  // Create the "graph" sheet if it doesn't exist
  if (!sheet) {
    sheet = spreadsheet.insertSheet("graph");
    console.log('Created new sheet named "graph"');
  }
  
  // Check if sheet is empty
  if (sheet.getLastRow() === 0) {
    // Add headers
    const headers = [
      'Order ID',
      'Project ID', 
      'Project Name',
      'Project Type',
      'Region',
      'Price',
      'Date'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Add some sample data
    const sampleData = [
      ['ORD-001', 'PRJ-001', 'E-commerce Website', 'Development', 'North America', 15000, '2024-01-15'],
      ['ORD-002', 'PRJ-002', 'Brand Identity Design', 'Design', 'Europe', 8000, '2024-01-20'],
      ['ORD-003', 'PRJ-003', 'Digital Marketing Campaign', 'Marketing', 'Asia Pacific', 12000, '2024-01-25'],
      ['ORD-004', 'PRJ-004', 'Mobile App Development', 'Development', 'North America', 25000, '2024-02-01'],
      ['ORD-005', 'PRJ-005', 'UI/UX Redesign', 'Design', 'Europe', 9500, '2024-02-05']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    console.log('Sheet "graph" setup completed with sample data');
  } else {
    console.log('Sheet "graph" already exists with data');
  }
} 