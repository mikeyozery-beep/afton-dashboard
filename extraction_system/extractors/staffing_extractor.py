"""Extract Staffing Vacancies from Excel file"""
import pandas as pd
from datetime import datetime
from pathlib import Path

class StaffingExtractor:
    def __init__(self, config):
        self.config = config
        self.results = {}
    
    def extract_staffing_vacancies(self):
        try:
            dashboard_path = Path(r"C:/Users/MichaelOzery/OneDrive - Afton Properties/Old Dropbox/My PC (DESKTOP-5D77V89)/Mike Ozery/Dashboard Data")
            file_path = dashboard_path / "Open_Positions.xlsx"
            
            if not file_path.exists():
                self.results['staffing_vacancies'] = {
                    'value': None,
                    'source': str(file_path),
                    'status': 'FILE_NOT_FOUND',
                    'message': 'Please download Open_Positions.xlsx from SharePoint',
                    'timestamp': datetime.now().isoformat()
                }
                return False
            
            df = pd.read_excel(file_path)
            total = len(df)
            
            self.results['staffing_vacancies'] = {
                'value': total,
                'source': 'Open_Positions.xlsx',
                'timestamp': datetime.now().isoformat(),
                'status': 'SUCCESS'
            }
            return True
        except Exception as e:
            self.results['staffing_vacancies'] = {
                'value': None,
                'error': str(e),
                'timestamp': datetime.now().isoformat()
            }
            return False
    
    def get_results(self):
        return self.results
