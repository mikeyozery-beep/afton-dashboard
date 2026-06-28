"""Master extraction and analysis orchestrator for Afton Dashboard"""
import json
import sys
from pathlib import Path
from datetime import datetime
import logging

# Setup logging with proper folder creation
log_dir = Path('logs')
log_dir.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_dir / 'extraction.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class DashboardExtractor:
    def __init__(self, config_path='config/data_sources.json'):
        config_file = Path(config_path)
        if not config_file.exists():
            logger.error(f"Config file not found: {config_file.absolute()}")
            raise FileNotFoundError(f"Config not found at {config_file.absolute()}")
        
        with open(config_file) as f:
            self.config = json.load(f)
        self.metrics = {}
        self.analysis_results = {}
    
    def extract_all_daily(self):
        """Run all daily extractions"""
        logger.info("Starting daily extraction cycle")
        
        # Excel extractions
        try:
            from extractors.excel_extractor import ExcelExtractor
            excel_ext = ExcelExtractor(self.config)
            excel_ext.extract_occupancy_metrics()
            excel_ext.extract_percent_collected()
            self.metrics.update(excel_ext.get_results())
            logger.info("✓ Excel extractions complete")
        except Exception as e:
            logger.error(f"Excel extraction failed: {e}")
        
        # Box Score Summary
        try:
            from extractors.box_score_extractor import BoxScoreExtractor
            box_ext = BoxScoreExtractor(self.config)
            box_ext.extract_summary_metrics()
            self.metrics.update(box_ext.get_results())
            logger.info("✓ Box Score extractions complete")
        except Exception as e:
            logger.error(f"Box Score extraction failed: {e}")
        
        # Staffing Vacancies from Excel
        try:
            from extractors.staffing_extractor import StaffingExtractor
            staff_ext = StaffingExtractor(self.config)
            staff_ext.extract_staffing_vacancies()
            self.metrics.update(staff_ext.get_results())
            logger.info("✓ Staffing extraction complete")
        except Exception as e:
            logger.error(f"Staffing extraction failed: {e}")

        # Budget data (occupancy budget weighted average)
        try:
            from extractors.budget_extractor import BudgetExtractor
            budget_ext = BudgetExtractor(self.config)
            budget_config = self.config.get('data_sources', {}).get('budgets', {})
            budget_ext.set_paths(
                budget_config.get('portfolio_index_path'),
                budget_config.get('budgets_folder_path')
            )
            budget_ext.calculate_occupancy_budget_weighted_avg()
            self.metrics.update(budget_ext.get_results())
            logger.info("✓ Budget extraction complete")
        except Exception as e:
            logger.error(f"Budget extraction failed: {e}")

        return self.metrics
    
    def extract_all_monthly(self):
        """Run all monthly extractions"""
        logger.info("Starting monthly extraction cycle")
        
        try:
            from extractors.excel_extractor import ExcelExtractor
            excel_ext = ExcelExtractor(self.config)
            excel_ext.extract_rent_growth()
            self.metrics.update(excel_ext.get_results())
            logger.info("✓ Monthly Excel extractions complete")
        except Exception as e:
            logger.error(f"Monthly extraction failed: {e}")
        
        try:
            from extractors.financial_extractor import FinancialExtractor
            fin_ext = FinancialExtractor(self.config)
            fin_ext.extract_all_financial()
            self.metrics.update(fin_ext.get_results())
            logger.info("✓ Financial extractions complete")
        except Exception as e:
            logger.error(f"Financial extraction failed: {e}")
        
        return self.metrics
    
    def run_analysis(self):
        """Run all analysis modules"""
        logger.info("Starting analysis phase")
        
        sample_property_data = {
            'Sienna Heights': {'occupancy': 90.4, 'leased_occupancy': 92.1, 'collections': 96.7, 'rent_growth': 1.6, 'noi_variance': -3.5},
            'Cassia Apartments': {'occupancy': 88.6, 'leased_occupancy': 90.3, 'collections': 96.2, 'rent_growth': 0.9, 'noi_variance': -2.9},
            'Hills of Corona': {'occupancy': 89.1, 'leased_occupancy': 89.6, 'collections': 97.1, 'rent_growth': 1.2, 'noi_variance': -2.3},
        }
        
        try:
            from analyzers.bottom_performer_analyzer import BottomPerformerAnalyzer
            bottom_analyzer = BottomPerformerAnalyzer(self.config, sample_property_data)
            bottom_performers = bottom_analyzer.calculate_bottom_performers()
            self.analysis_results['bottom_performers'] = bottom_performers
            logger.info("✓ Bottom performer analysis complete")
        except Exception as e:
            logger.error(f"Bottom performer analysis failed: {e}")
        
        try:
            from analyzers.anomaly_detector import AnomalyDetector
            anomaly_detector = AnomalyDetector(self.config, sample_property_data)
            anomalies = anomaly_detector.detect_all_anomalies()
            self.analysis_results['anomalies'] = anomalies
            logger.info("✓ Anomaly detection complete")
        except Exception as e:
            logger.error(f"Anomaly detection failed: {e}")
        
        try:
            from analyzers.analyst_summary_generator import AnalystSummaryGenerator
            summary_gen = AnalystSummaryGenerator(self.config, anomalies if 'anomalies' in self.analysis_results else [], 
                                                 bottom_performers if 'bottom_performers' in self.analysis_results else [], 
                                                 sample_property_data)
            focus_areas = summary_gen.generate_focus_area_summaries()
            self.analysis_results['focus_areas'] = focus_areas
            logger.info("✓ Analyst summary generation complete")
        except Exception as e:
            logger.error(f"Analyst summary generation failed: {e}")
        
        return self.analysis_results
    
    def output_dashboard_json(self, output_path='outputs/dashboard_metrics.json'):
        """Format and output all metrics and analysis to JSON"""
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)
        
        output = {
            'timestamp': datetime.now().isoformat(),
            'metrics': self.metrics,
            'analysis': self.analysis_results,
            'status': 'complete'
        }
        
        with open(output_path, 'w') as f:
            json.dump(output, f, indent=2)
        
        logger.info(f"✓ Dashboard JSON output to {output_path}")
        return output_path
    
    def run_full_cycle(self, monthly=False):
        """Execute full extraction and analysis cycle"""
        logger.info("=" * 60)
        logger.info("AFTON DASHBOARD EXTRACTION CYCLE STARTED")
        logger.info("=" * 60)
        
        self.extract_all_daily()
        
        if monthly:
            self.extract_all_monthly()
        
        self.run_analysis()
        self.output_dashboard_json()
        
        logger.info("=" * 60)
        logger.info("EXTRACTION CYCLE COMPLETE")
        logger.info("=" * 60)
        
        return {
            'metrics': self.metrics,
            'analysis': self.analysis_results
        }


if __name__ == '__main__':
    try:
        extractor = DashboardExtractor()
        monthly = '--monthly' in sys.argv
        results = extractor.run_full_cycle(monthly=monthly)
        print("\n✓ Extraction complete. Check logs/extraction.log for details.\n")
    except Exception as e:
        logger.error(f"FATAL ERROR: {e}", exc_info=True)
        sys.exit(1)
