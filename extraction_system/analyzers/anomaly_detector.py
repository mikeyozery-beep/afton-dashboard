"""Detect anomalies across portfolio metrics"""
from datetime import datetime

class AnomalyDetector:
    def __init__(self, config, current_data, previous_data=None, budget_data=None):
        self.config = config
        self.current_data = current_data
        self.previous_data = previous_data
        self.budget_data = budget_data
        self.results = {}
    
    def detect_all_anomalies(self):
        all_anomalies = []
        
        for prop in self.current_data:
            all_anomalies.append({
                'property': prop,
                'metric': 'occupancy',
                'value': self.current_data[prop].get('occupancy', 0),
                'type': 'absolute_outlier'
            })
        
        property_anomaly_count = {}
        for anomaly in all_anomalies:
            prop = anomaly['property']
            if prop not in property_anomaly_count:
                property_anomaly_count[prop] = {'count': 0, 'anomalies': []}
            property_anomaly_count[prop]['count'] += 1
            property_anomaly_count[prop]['anomalies'].append(anomaly)
        
        top_anomalies = sorted(property_anomaly_count.items(), 
                              key=lambda x: x[1]['count'], reverse=True)[:3]
        
        self.results['anomalies'] = [
            {
                'rank': i + 1,
                'property': prop,
                'anomaly_count': data['count'],
                'detected_issues': data['anomalies'],
                'timestamp': datetime.now().isoformat()
            }
            for i, (prop, data) in enumerate(top_anomalies)
        ]
        
        return self.results['anomalies']
    
    def get_results(self):
        return self.results
