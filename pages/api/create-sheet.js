// pages/api/create-sheet.js
import { google } from 'googleapis';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ message: 'Method not allowed' });
  }

  try {
    const { data, title } = req.body;

    if (!data || !Array.isArray(data)) {
      return res.status(400).json({ message: 'Invalid data format' });
    }

    // Initialize Google Sheets API
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.VITE_GOOGLE_SERVICE_ACCOUNT_EMAIL,
        private_key: process.env.VITE_GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
      },
      scopes: [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
      ],
    });

    const sheets = google.sheets({ version: 'v4', auth });
    const drive = google.drive({ version: 'v3', auth });

    // Create new spreadsheet
    const spreadsheet = await sheets.spreadsheets.create({
      requestBody: {
        properties: {
          title: title || 'RVCE Counsellor Data'
        }
      }
    });

    const spreadsheetId = spreadsheet.data.spreadsheetId;

    // Add data to the sheet
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: 'A1',
      valueInputOption: 'RAW',
      requestBody: {
        values: data
      }
    });

    // Advanced formatting requests
    const requests = [
      // Logo space formatting (rows 1-8 merged)
      {
        mergeCells: {
          range: {
            sheetId: 0,
            startRowIndex: 0,
            endRowIndex: 8,
            startColumnIndex: 0,
            endColumnIndex: 10
          },
          mergeType: 'MERGE_ALL'
        }
      },
      // Logo cell styling
      {
        repeatCell: {
          range: {
            sheetId: 0,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: 0,
            endColumnIndex: 1
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: {
                red: 0.97,
                green: 0.98,
                blue: 0.98
              },
              textFormat: {
                italic: true,
                fontSize: 14,
                foregroundColor: {
                  red: 0.4,
                  green: 0.4,
                  blue: 0.4
                }
              },
              horizontalAlignment: 'CENTER',
              verticalAlignment: 'MIDDLE'
            }
          },
          fields: 'userEnteredFormat'
        }
      },
      // Header row formatting (row 9 - blue background)
      {
        repeatCell: {
          range: {
            sheetId: 0,
            startRowIndex: 8,
            endRowIndex: 9,
            startColumnIndex: 0,
            endColumnIndex: 10
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: {
                red: 0.21,
                green: 0.38,
                blue: 0.57
              },
              textFormat: {
                bold: true,
                fontSize: 12,
                foregroundColor: {
                  red: 1,
                  green: 1,
                  blue: 1
                }
              },
              horizontalAlignment: 'CENTER',
              verticalAlignment: 'MIDDLE'
            }
          },
          fields: 'userEnteredFormat'
        }
      },
      // Set column widths
      {
        updateDimensionProperties: {
          range: {
            sheetId: 0,
            dimension: 'COLUMNS',
            startIndex: 0,
            endIndex: 10
          },
          properties: {
            pixelSize: 120
          },
          fields: 'pixelSize'
        }
      }
    ];

    // Add formatting for batch separators (dark blue) and branch separators (green)
    // We'll identify these by checking row content
    for (let i = 9; i < data.length + 1; i++) {
      const rowData = data[i - 1];
      if (rowData && rowData[0]) {
        const cellValue = String(rowData[0]).toLowerCase();
        
        // Batch separator formatting (dark blue)
        if (cellValue.includes('batch') && cellValue.includes('year')) {
          requests.push({
            mergeCells: {
              range: {
                sheetId: 0,
                startRowIndex: i - 1,
                endRowIndex: i,
                startColumnIndex: 0,
                endColumnIndex: 10
              },
              mergeType: 'MERGE_ALL'
            }
          });
          
          requests.push({
            repeatCell: {
              range: {
                sheetId: 0,
                startRowIndex: i - 1,
                endRowIndex: i,
                startColumnIndex: 0,
                endColumnIndex: 1
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: {
                    red: 0.04,
                    green: 0.33,
                    blue: 0.58
                  },
                  textFormat: {
                    bold: true,
                    fontSize: 16,
                    foregroundColor: {
                      red: 1,
                      green: 1,
                      blue: 1
                    }
                  },
                  horizontalAlignment: 'CENTER',
                  verticalAlignment: 'MIDDLE'
                }
              },
              fields: 'userEnteredFormat'
            }
          });
          
          // Set row height
          requests.push({
            updateDimensionProperties: {
              range: {
                sheetId: 0,
                dimension: 'ROWS',
                startIndex: i - 1,
                endIndex: i
              },
              properties: {
                pixelSize: 40
              },
              fields: 'pixelSize'
            }
          });
        }
        
        // Branch separator formatting (green)
        else if (cellValue.includes('engineering') || cellValue.includes('science') || 
                 cellValue.includes('technology') || cellValue.includes('management')) {
          requests.push({
            mergeCells: {
              range: {
                sheetId: 0,
                startRowIndex: i - 1,
                endRowIndex: i,
                startColumnIndex: 0,
                endColumnIndex: 10
              },
              mergeType: 'MERGE_ALL'
            }
          });
          
          requests.push({
            repeatCell: {
              range: {
                sheetId: 0,
                startRowIndex: i - 1,
                endRowIndex: i,
                startColumnIndex: 0,
                endColumnIndex: 1
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: {
                    red: 0.42,
                    green: 0.66,
                    blue: 0.31
                  },
                  textFormat: {
                    bold: true,
                    fontSize: 14,
                    foregroundColor: {
                      red: 1,
                      green: 1,
                      blue: 1
                    }
                  },
                  horizontalAlignment: 'CENTER',
                  verticalAlignment: 'MIDDLE'
                }
              },
              fields: 'userEnteredFormat'
            }
          });
          
          // Set row height
          requests.push({
            updateDimensionProperties: {
              range: {
                sheetId: 0,
                dimension: 'ROWS',
                startIndex: i - 1,
                endIndex: i
              },
              properties: {
                pixelSize: 30
              },
              fields: 'pixelSize'
            }
          });
        }
      }
    }

    // Apply all formatting
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests
      }
    });

    // Add alternating row colors for data rows
    const alternatingRowsRequest = {
      addBanding: {
        bandedRange: {
          range: {
            sheetId: 0,
            startRowIndex: 9,
            endRowIndex: data.length,
            startColumnIndex: 0,
            endColumnIndex: 10
          },
          rowProperties: {
            firstBandColor: {
              red: 1,
              green: 1,
              blue: 1
            },
            secondBandColor: {
              red: 0.97,
              green: 0.98,
              blue: 0.98
            }
          }
        }
      }
    };

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [alternatingRowsRequest]
      }
    });

    // Make sheet publicly viewable (read-only)
    await drive.permissions.create({
      fileId: spreadsheetId,
      requestBody: {
        role: 'reader',
        type: 'anyone'
      }
    });

    // Return the sheet URL
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
    
    res.status(200).json({ 
      url: sheetUrl,
      spreadsheetId,
      message: 'Google Sheet created successfully with professional formatting'
    });

  } catch (error) {
    console.error('Error creating Google Sheet:', error);
    res.status(500).json({ 
      message: 'Failed to create Google Sheet',
      error: error.message 
    });
  }
}

// Optional: Add request size limit for large datasets
export const config = {
  api: {
    bodyParser: {
      sizeLimit: '10mb',
    },
  },
};