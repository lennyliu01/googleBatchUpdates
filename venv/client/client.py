import gspread
import gspread_formatting as f
from oauth2client.service_account import ServiceAccountCredentials
class google_client(object):
    def __init__(self,googlesheet):
        scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
        doc_creds = ServiceAccountCredentials.from_json_keyfile_name("/Users/lennyliu/Documents/project/gsheet_client/testcred.json", scope)
        gc = gspread.authorize(doc_creds)
        filename = googlesheet
        self.sh = gc.open(filename)
    def open_tab(self,sheetname):
        working_sheet = self.sh.worksheet(sheetname)
        return working_sheet
    def new_tab(self,new_sheet_name):
        self.sh.add_worksheet(new_sheet_name,rows="1000",cols="30")
        print('added a new sheet: ',new_sheet_name)
    def delete_tab(self,sheetname):
        worksheet = self.sh.worksheet(sheetname)
        self.sh.del_worksheet(worksheet)
        print('removed a sheet :',sheetname)
    def clear_all_values(self,sheetname):
        worksheet = self.sh.worksheet(sheetname)
        self.sh.values_clear(sheetname)
    def sheet_formatter(self,sheetname):
        working_sheet = self.sh.worksheet(sheetname)
        sheetId = working_sheet.id
        body = {
            "requests": [
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheetId,
                        "dimension": 'COLUMNS',
                        "startIndex": 1, 
                        "endIndex": 5  
                    }
                }
            }
            ]
        }
        self.sh.batch_update(body)
        header_fmt = f.CellFormat(
        backgroundColor=f.Color(0.043137256,0.3254902,0.5803922),
        textFormat=f.TextFormat(bold=True, foregroundColor=f.Color(1,1,1)),
        horizontalAlignment='CENTER',verticalAlignment='MIDDLE'
        )
        f.set_frozen(working_sheet,rows=1,cols=2) # freeze the headers and hotel_id and status column
        f.format_cell_range(working_sheet,'1',header_fmt)


#test connection
name = 'test'
working_file = google_client('batchupdate')
sheet = working_file.open_tab(name)
working_file.sheet_formatter(name)