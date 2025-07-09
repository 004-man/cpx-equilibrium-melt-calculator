import pandas as pd
import numpy as np
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

class GeochemicalMeltCalculator:
    """
    Intelligent, flexible geochemical modeling tool for calculating equilibrium 
    melt compositions from clinopyroxene data with automatic element detection
    and multi-study capability.
    """
    
    def __init__(self, excel_file_path):
        """
        Initialize with Excel file path.
        
        Parameters:
        -----------
        excel_file_path : str
            Path to the Excel file containing geochemical data
        """
        self.excel_file = excel_file_path
        self.all_sheets = None
        self.study_data = {}
        self.kd_data = {}
        self.normalizing_data = {}
        self.results = {}
        
        # Load and analyze the Excel file
        self._load_excel_data()
        
    def _load_excel_data(self):
        """Load and intelligently parse Excel data"""
        try:
            # Read all sheets
            self.all_sheets = pd.read_excel(self.excel_file, sheet_name=None)
            print(f"📊 Loaded Excel file with {len(self.all_sheets)} sheets")
            print(f"📋 Available sheets: {list(self.all_sheets.keys())}")
            
            # Identify different types of sheets
            self._identify_sheet_types()
            
        except Exception as e:
            print(f"❌ Error loading Excel file: {e}")
            raise
    
    def _identify_sheet_types(self):
        """Intelligently identify sheet types and content"""
        
        # Keywords to identify different sheet types
        kd_keywords = ['kd', 'partition', 'coefficient']
        norm_keywords = ['normalizing', 'primitive', 'mantle', 'chondrite', 'pm']
        ref_keywords = ['reference', 'citation', 'ref']
        
        for sheet_name, df in self.all_sheets.items():
            sheet_lower = sheet_name.lower()
            
            # Check if it's a Kd values sheet
            if any(keyword in sheet_lower for keyword in kd_keywords):
                self._parse_kd_sheet(sheet_name, df)
                
            # Check if it's a normalizing values sheet
            elif any(keyword in sheet_lower for keyword in norm_keywords):
                self._parse_normalizing_sheet(sheet_name, df)
                
            # Check if it's a references sheet
            elif any(keyword in sheet_lower for keyword in ref_keywords):
                print(f"📚 Found references sheet: {sheet_name}")
                
            # Otherwise, treat as study data
            else:
                self._parse_study_sheet(sheet_name, df)
    
    def _parse_study_sheet(self, sheet_name, df):
        """Parse study data sheet with clinopyroxene compositions"""
        try:
            # Set first column as index (elements)
            df.set_index(df.columns[0], inplace=True)
            
            # Remove any empty rows/columns
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            # Convert to numeric, replacing non-numeric with NaN
            df = df.apply(pd.to_numeric, errors='coerce')
            
            # Store study data
            self.study_data[sheet_name] = df
            
            elements = df.index.tolist()
            samples = df.columns.tolist()
            
            print(f"🔬 Study '{sheet_name}': {len(elements)} elements, {len(samples)} samples")
            print(f"   Elements: {elements[:5]}{'...' if len(elements) > 5 else ''}")
            print(f"   Samples: {samples[:3]}{'...' if len(samples) > 3 else ''}")
            
        except Exception as e:
            print(f"⚠️ Error parsing study sheet '{sheet_name}': {e}")
    
    def _parse_kd_sheet(self, sheet_name, df):
        """Parse Kd values sheet"""
        try:
            # Assume first column is elements, second is Kd values
            df.set_index(df.columns[0], inplace=True)
            kd_column = df.columns[0]  # First data column
            
            # Convert to numeric
            kd_series = pd.to_numeric(df[kd_column], errors='coerce')
            
            # Store Kd data
            self.kd_data = kd_series.dropna()
            
            print(f"⚖️ Loaded Kd values: {len(self.kd_data)} elements")
            print(f"   Source: {kd_column}")
            print(f"   Elements: {list(self.kd_data.index[:5])}{'...' if len(self.kd_data) > 5 else ''}")
            
        except Exception as e:
            print(f"⚠️ Error parsing Kd sheet '{sheet_name}': {e}")
    
    def _parse_normalizing_sheet(self, sheet_name, df):
        """Parse normalizing values sheet"""
        try:
            # Set first column as index (elements)
            df.set_index(df.columns[0], inplace=True)
            norm_column = df.columns[0]  # First data column
            
            # Convert to numeric
            norm_series = pd.to_numeric(df[norm_column], errors='coerce')
            
            # Store normalizing data
            self.normalizing_data = norm_series.dropna()
            
            print(f"📏 Loaded normalizing values: {len(self.normalizing_data)} elements")
            print(f"   Standard: {norm_column}")
            print(f"   Elements: {list(self.normalizing_data.index[:5])}{'...' if len(self.normalizing_data) > 5 else ''}")
            
        except Exception as e:
            print(f"⚠️ Error parsing normalizing sheet '{sheet_name}': {e}")
    
    def calculate_equilibrium_melts(self):
        """Calculate equilibrium melt compositions for all studies"""
        
        if not self.kd_data.any() or not self.normalizing_data.any():
            print("❌ Missing Kd values or normalizing data!")
            return
        
        print("\n🧮 Calculating equilibrium melt compositions...")
        
        for study_name, cpx_data in self.study_data.items():
            print(f"\n📊 Processing study: {study_name}")
            
            # Find common elements across all datasets while preserving original order
            common_elements_set = set(cpx_data.index) & set(self.kd_data.index) & set(self.normalizing_data.index)
            # Preserve the original order from the Excel file
            common_elements = [elem for elem in cpx_data.index if elem in common_elements_set]
            
            if len(common_elements) == 0:
                print(f"⚠️ No common elements found for study '{study_name}'")
                continue
            
            print(f"   ✅ Processing {len(common_elements)} common elements: {common_elements}")
            
            # Calculate melt compositions for each sample
            study_results = {}
            
            for sample in cpx_data.columns:
                sample_results = self._calculate_sample_melt(cpx_data[sample], common_elements, common_elements)
                if sample_results is not None:
                    study_results[sample] = sample_results
            
            if study_results:
                self.results[study_name] = study_results
                print(f"   ✅ Successfully calculated melts for {len(study_results)} samples")
            else:
                print(f"   ⚠️ No valid results for study '{study_name}'")
    
    def _calculate_sample_melt(self, cpx_sample, common_elements, original_element_order):
        """Calculate melt composition for a single sample"""
        try:
            # Filter data for common elements only
            cpx_conc = cpx_sample.loc[common_elements].dropna()
            
            if len(cpx_conc) == 0:
                return None
            
            # Get corresponding Kd and normalizing values
            kd_vals = self.kd_data.loc[cpx_conc.index]
            norm_vals = self.normalizing_data.loc[cpx_conc.index]
            
            # Calculate melt concentration: Melt = Cpx / Kd
            melt_conc = cpx_conc / kd_vals
            
            # Calculate normalized values: Normalized = Melt / PM
            normalized_vals = melt_conc / norm_vals
            
            # Preserve original element order - only include elements that have data
            available_elements = [elem for elem in original_element_order if elem in cpx_conc.index]
            
            return {
                'cpx_concentration': cpx_conc.reindex(available_elements),
                'melt_concentration': melt_conc.reindex(available_elements),
                'normalized_values': normalized_vals.reindex(available_elements),
                'elements': available_elements
            }
            
        except Exception as e:
            print(f"      ⚠️ Error calculating melt for sample: {e}")
            return None
    
    def generate_results_table(self, study_name=None):
        """Generate comprehensive results table"""
        
        if not self.results:
            print("❌ No results available. Run calculate_equilibrium_melts() first.")
            return None
        
        # If specific study requested, process only that study
        studies_to_process = [study_name] if study_name else list(self.results.keys())
        
        all_results = []
        
        for study in studies_to_process:
            if study not in self.results:
                print(f"⚠️ Study '{study}' not found in results")
                continue
                
            study_data = self.results[study]
            
            for sample_name, sample_data in study_data.items():
                elements = sample_data['elements']
                
                for element in elements:
                    result_row = {
                        'Study': study,
                        'Sample': sample_name,
                        'Element': element,
                        'Cpx_Concentration': sample_data['cpx_concentration'][element],
                        'Kd_Value': self.kd_data[element],
                        'Melt_Concentration': sample_data['melt_concentration'][element],
                        'PM_Normalizing_Value': self.normalizing_data[element],
                        'PM_Normalized': sample_data['normalized_values'][element]
                    }
                    all_results.append(result_row)
        
        results_df = pd.DataFrame(all_results)
        
        # Sort results to maintain original element order
        if not results_df.empty:
            # Create a mapping of element to its original position
            first_study = list(self.results.keys())[0]
            first_sample = list(self.results[first_study].keys())[0]
            original_order = self.results[first_study][first_sample]['elements']
            element_order_map = {elem: i for i, elem in enumerate(original_order)}
            
            # Sort DataFrame by original element order
            results_df['_sort_key'] = results_df['Element'].map(element_order_map)
            results_df = results_df.sort_values(['Study', 'Sample', '_sort_key']).drop('_sort_key', axis=1)
        
        if not results_df.empty:
            print(f"📋 Generated results table with {len(results_df)} rows")
            return results_df
        else:
            print("⚠️ No results to display")
            return None
    
    def export_results(self, filename='Equilibrium_Melt_Results.xlsx'):
        """Export results to Excel file with improved structure"""
        
        if not self.results:
            print("❌ No results available. Run calculate_equilibrium_melts() first.")
            return
        
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            
            # Export individual study results with improved structure
            for study_name, study_data in self.results.items():
                
                # Get original element order from Excel input
                original_excel_order = list(self.study_data[study_name].index)
                
                # Find elements that have results (preserving original order)
                elements_with_results = set()
                for sample_data in study_data.values():
                    elements_with_results.update(sample_data['elements'])
                
                ordered_elements = [elem for elem in original_excel_order if elem in elements_with_results]
                
                # Create structured output: Melt compositions first, then normalized values
                output_data = []
                
                # Section 1: Calculated Melt Compositions
                output_data.append(['CALCULATED MELT COMPOSITION'] + [''] * (len(study_data)))
                output_data.append(['=== MELT COMPOSITIONS (ppm) ==='] + [''] * (len(study_data)))
                output_data.append(['Element'] + list(study_data.keys()))
                
                for element in ordered_elements:
                    row = [element]
                    for sample_name, sample_data in study_data.items():
                        if element in sample_data['melt_concentration'].index:
                            row.append(sample_data['melt_concentration'][element])
                        else:
                            row.append('')
                    output_data.append(row)
                
                # Add spacing
                output_data.append([''] * (len(study_data) + 1))
                output_data.append([''] * (len(study_data) + 1))
                
                # Section 2: Normalized Values
                output_data.append(['NORMALIZED VALUES'] + [''] * (len(study_data)))
                output_data.append(['=== PM NORMALIZED VALUES ==='] + [''] * (len(study_data)))
                output_data.append(['Element'] + list(study_data.keys()))
                
                for element in ordered_elements:
                    row = [element]
                    for sample_name, sample_data in study_data.items():
                        if element in sample_data['normalized_values'].index:
                            row.append(sample_data['normalized_values'][element])
                        else:
                            row.append('')
                    output_data.append(row)
                
                # Convert to DataFrame and export
                max_cols = max(len(row) for row in output_data)
                for row in output_data:
                    while len(row) < max_cols:
                        row.append('')
                
                # Create column names
                col_names = [f'Col_{i}' for i in range(max_cols)]
                output_df = pd.DataFrame(output_data, columns=col_names)
                output_df.to_excel(writer, sheet_name=f'{study_name[:30]}', index=False, header=False)
            
            # Export input data summary
            summary_data = {
                'Parameter': ['Number of Studies', 'Total Samples', 'Common Elements', 'Kd Source', 'Normalizing Standard'],
                'Value': [
                    len(self.results),
                    sum(len(study_data) for study_data in self.results.values()),
                    len(set().union(*[set().union(*[sample_data['elements'] for sample_data in study_data.values()]) 
                                    for study_data in self.results.values()])),
                    'Grassi et al. (2012)' if 'grassi' in str(self.kd_data.name).lower() else 'Unknown',
                    'McDonough & Sun (1995)' if 'mcdonough' in str(self.normalizing_data.name).lower() else 'Unknown'
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        print(f"📊 Exported results to '{filename}'")
    
    def print_summary(self):
        """Print comprehensive summary of loaded data and results"""
        
        print("\n" + "="*60)
        print("🧪 GEOCHEMICAL MELT CALCULATOR SUMMARY")
        print("="*60)
        
        # Input data summary
        print(f"\n📥 INPUT DATA:")
        print(f"   📁 Excel file: {self.excel_file}")
        print(f"   📊 Studies loaded: {len(self.study_data)}")
        
        for study_name, data in self.study_data.items():
            print(f"      • {study_name}: {len(data.columns)} samples, {len(data.index)} elements")
        
        print(f"   ⚖️ Kd values: {len(self.kd_data)} elements")
        print(f"   📏 Normalizing values: {len(self.normalizing_data)} elements")
        
        # Results summary
        if self.results:
            print(f"\n📊 CALCULATED RESULTS:")
            total_samples = sum(len(study_data) for study_data in self.results.values())
            print(f"   ✅ Successfully processed: {len(self.results)} studies, {total_samples} samples")
            
            for study_name, study_data in self.results.items():
                print(f"      • {study_name}: {len(study_data)} samples")
        else:
            print(f"\n❌ No results calculated yet. Run calculate_equilibrium_melts()")
        
        print("\n" + "="*60)
        """Print comprehensive summary of loaded data and results"""
        
        print("\n" + "="*60)
        print("🧪 GEOCHEMICAL MELT CALCULATOR SUMMARY")
        print("="*60)
        
        # Input data summary
        print(f"\n📥 INPUT DATA:")
        print(f"   📁 Excel file: {self.excel_file}")
        print(f"   📊 Studies loaded: {len(self.study_data)}")
        
        for study_name, data in self.study_data.items():
            print(f"      • {study_name}: {len(data.columns)} samples, {len(data.index)} elements")
        
        print(f"   ⚖️ Kd values: {len(self.kd_data)} elements")
        print(f"   📏 Normalizing values: {len(self.normalizing_data)} elements")
        
        # Results summary
        if self.results:
            print(f"\n📊 CALCULATED RESULTS:")
            total_samples = sum(len(study_data) for study_data in self.results.values())
            print(f"   ✅ Successfully processed: {len(self.results)} studies, {total_samples} samples")
            
            for study_name, study_data in self.results.items():
                print(f"      • {study_name}: {len(study_data)} samples")
        else:
            print(f"\n❌ No results calculated yet. Run calculate_equilibrium_melts()")
        
        print("\n" + "="*60)
    
    def print_summary(self):
        """Print comprehensive summary of loaded data and results"""
        
        print("\n" + "="*60)
        print("🧪 GEOCHEMICAL MELT CALCULATOR SUMMARY")
        print("="*60)
        
        # Input data summary
        print(f"\n📥 INPUT DATA:")
        print(f"   📁 Excel file: {self.excel_file}")
        print(f"   📊 Studies loaded: {len(self.study_data)}")
        
        for study_name, data in self.study_data.items():
            print(f"      • {study_name}: {len(data.columns)} samples, {len(data.index)} elements")
        
        print(f"   ⚖️ Kd values: {len(self.kd_data)} elements")
        print(f"   📏 Normalizing values: {len(self.normalizing_data)} elements")
        
        # Results summary
        if self.results:
            print(f"\n📊 CALCULATED RESULTS:")
            total_samples = sum(len(study_data) for study_data in self.results.values())
            print(f"   ✅ Successfully processed: {len(self.results)} studies, {total_samples} samples")
            
            for study_name, study_data in self.results.items():
                print(f"      • {study_name}: {len(study_data)} samples")
        else:
            print(f"\n❌ No results calculated yet. Run calculate_equilibrium_melts()")
        
        print("\n" + "="*60)


def main():
    """
    Main function demonstrating the usage of GeochemicalMeltCalculator
    """
    
    # Example usage
    print("🧪 INTELLIGENT GEOCHEMICAL EQUILIBRIUM MELT CALCULATOR")
    print("="*55)
    
    # Initialize calculator
    excel_file = "Equilibrium melt calculation Input.xlsx"  # Update path as needed
    
    try:
        # Create calculator instance
        calculator = GeochemicalMeltCalculator(excel_file)
        
        # Calculate equilibrium melts
        calculator.calculate_equilibrium_melts()
        
        # Export results (no plot generation)
        calculator.export_results()
        
        # Print summary
        calculator.print_summary()
        
        print(f"\n✅ Analysis complete! Check 'Equilibrium_Melt_Results.xlsx' for detailed results.")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False
    
    return True


if __name__ == "__main__":
    main()