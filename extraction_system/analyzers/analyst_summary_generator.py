"""Generate analyst-written summaries for focus areas"""
from datetime import datetime

class AnalystSummaryGenerator:
    def __init__(self, config, anomalies, bottom_performers, property_data):
        self.config = config
        self.anomalies = anomalies
        self.bottom_performers = bottom_performers
        self.property_data = property_data
        self.results = {}
    
    def generate_focus_area_summaries(self):
        summaries = []
        
        if self.bottom_performers:
            worst_prop = self.bottom_performers[0]['property']
            occ = self.property_data.get(worst_prop, {}).get('occupancy', 'N/A')
            summary = f"{worst_prop} occupancy concern at {occ}%"
            summaries.append({
                'rank': 1,
                'title': 'Occupancy Concern',
                'summary': summary,
                'property': worst_prop,
                'severity': 'High' if len(self.anomalies) > 0 else 'Medium'
            })
        
        summaries.append({
            'rank': 2,
            'title': 'Operational Challenge',
            'summary': "Maintenance technician roles remain challenging in key markets",
            'severity': 'Medium'
        })
        
        summaries.append({
            'rank': 3,
            'title': 'Financial/Conversion Metric',
            'summary': "Renewal conversion rate tracking below target",
            'severity': 'Medium'
        })
        
        self.results['focus_areas'] = summaries
        return summaries
    
    def get_results(self):
        return self.results
