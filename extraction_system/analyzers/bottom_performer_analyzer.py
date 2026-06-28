"""Calculate bottom performing communities with normalized weighted metrics"""
import numpy as np
from datetime import datetime

class BottomPerformerAnalyzer:
    def __init__(self, config, property_data):
        self.config = config
        self.property_data = property_data
        self.results = {}
    
    def calculate_bottom_performers(self):
        """
        Normalize 5 metrics to 0-100 scale and calculate composite scores.
        Metrics: occupancy, leased_occupancy, collections (% collected), rent_growth, noi_variance
        """
        
        # Define the 5 metrics per config
        metrics = ['occupancy', 'leased_occupancy', 'collections', 'rent_growth', 'noi_variance']
        
        # Build metric data from property_data
        metric_values = {metric: {} for metric in metrics}
        
        for prop, data in self.property_data.items():
            for metric in metrics:
                metric_map = {
                    'occupancy': 'occupancy',
                    'leased_occupancy': 'leased_occupancy',
                    'collections': 'percent_collected',
                    'rent_growth': 'rent_growth',
                    'noi_variance': 'noi_variance'
                }
                
                key = metric_map.get(metric, metric)
                if key in data:
                    metric_values[metric][prop] = data[key]
        
        # Normalize each metric to 0-100 scale
        # For occupancy/collections/rent_growth: higher is better (lower score = worse)
        # For NOI variance: less negative is better (more negative = worse)
        
        normalized_scores = {prop: {} for prop in self.property_data.keys()}
        
        for metric in metrics:
            values_list = list(metric_values[metric].values())
            props_list = list(metric_values[metric].keys())
            
            if not values_list:
                continue
            
            min_val = min(values_list)
            max_val = max(values_list)
            
            for prop, value in zip(props_list, values_list):
                if max_val == min_val:
                    normalized_scores[prop][metric] = 50
                else:
                    # Normalize to 0-100: lower actual value = lower score (worse)
                    norm_value = ((value - min_val) / (max_val - min_val)) * 100
                    normalized_scores[prop][metric] = norm_value
        
        # Calculate equal-weighted composite score (1/5 each metric)
        property_scores = {}
        for prop in self.property_data.keys():
            if prop in normalized_scores:
                metric_scores = normalized_scores[prop]
                if metric_scores:
                    avg_score = np.mean(list(metric_scores.values()))
                    property_scores[prop] = {
                        'composite_score': avg_score,
                        'metric_scores': metric_scores
                    }
        
        # Sort by composite score (lowest first = worst performers)
        sorted_props = sorted(property_scores.items(), key=lambda x: x[1]['composite_score'])
        bottom_3 = sorted_props[:3]
        
        # Format results
        self.results['bottom_performers'] = [
            {
                'rank': i + 1,
                'property': prop,
                'composite_score': round(data['composite_score'], 1),
                'metric_breakdown': {
                    'occupancy': round(data['metric_scores'].get('occupancy', 0), 1),
                    'leased_occupancy': round(data['metric_scores'].get('leased_occupancy', 0), 1),
                    'collections': round(data['metric_scores'].get('collections', 0), 1),
                    'rent_growth': round(data['metric_scores'].get('rent_growth', 0), 1),
                    'noi_variance': round(data['metric_scores'].get('noi_variance', 0), 1)
                },
                'timestamp': datetime.now().isoformat()
            }
            for i, (prop, data) in enumerate(bottom_3)
        ]
        
        return self.results['bottom_performers']
    
    def get_results(self):
        return self.results
