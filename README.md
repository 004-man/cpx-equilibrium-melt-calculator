📖 Overview
A flexible Python tool for calculating equilibrium melt compositions from clinopyroxene geochemical data. This tool automatically detects elements, handles multiple studies, and generates primitive mantle-normalized results with proper element ordering.
🎯 Key Features

🧠 Intelligent Data Processing: Automatically detects sheet types and element suites
🔄 Multi-Study Capability: Processes multiple studies from different Excel sheets
📊 Flexible Input: Adapts to any element combination and normalizing standards
⚖️ Equilibrium Calculations: Calculates melt compositions using partition coefficients (Kd)
📏 Normalization: Primitive mantle normalization with any reference standard
📋 Clean Output: Well-structured Excel results with proper element ordering
🎨 Preservation of Order: Maintains your exact element sequence from input

🧪 Geochemical Application
This tool is designed for:

Clinopyroxene-melt equilibrium modeling
Trace element geochemistry

📋 Requirements
pandas>=1.3.0
openpyxl>=3.0.7
numpy>=1.20.0
🚀 Installation

📊 Input Data Structure
Your Excel file should contain the following sheets:
1. Study Data Sheets (e.g., "Present Study", "Literature Data")
Element    Sample1    Sample2    Sample3
Th         0.062      0.177      0.041
U          0.023      0.044      0.024
La         2.12       7.85       1.528
...
2. "Kd Values" Sheet
Element                    Kd Cpx/Melt
Th                         0.051
U                          0.076
La                         0.0044
...
3. "Normalizing values" Sheet
Element    McDonough and Sun (1995)
Th         0.0795
U          0.0203
La         0.648
...
💻 Usage
Basic Usage
pythonfrom geochem_melt_calculator import GeochemicalMeltCalculator

# Initialize calculator
calculator = GeochemicalMeltCalculator("your_data.xlsx")

# Calculate equilibrium melts
calculator.calculate_equilibrium_melts()

# Export results to Excel
calculator.export_results("results.xlsx")

# Print summary
calculator.print_summary()
Advanced Usage
python# Process specific study
results_table = calculator.generate_results_table("Present Study")

# Custom output filename
calculator.export_results("My_Melt_Calculations_2024.xlsx")
📊 Output Structure
The tool generates an Excel file with:
Study-Specific Sheets
CALCULATED MELT COMPOSITION
=== MELT COMPOSITIONS (ppm) ===
Element    Sample1    Sample2    Sample3
Th         1.25       3.47       0.80
...

NORMALIZED VALUES  
=== PM NORMALIZED VALUES ===
Element    Sample1    Sample2    Sample3
Th         15.7       43.6       10.1
...
Summary Sheet

Number of studies processed
Total samples analyzed
Elements included
Kd source reference
Normalizing standard reference

🔬 Scientific Background
The tool calculates melt compositions using the fundamental relationship:
Melt Concentration = Cpx Concentration / Kd
Where Kd is the partition coefficient (Cpx/Melt) from experimental studies.
Normalization
Primitive mantle normalization follows:
Normalized Value = Melt Concentration / Primitive Mantle Value
📚 Example Workflow

Prepare your data in the required Excel format
Run the calculator with your input file
Review results in the generated Excel output
Copy normalized values for spider diagrams
Use melt compositions for further geochemical modeling

📖 Citation
If you use this tool in your research, please cite:
Sarbajit Dash (2025). cpx equilibrium melt calculator. 
GitHub repository: https://github.com/004-man/cpx-equilibrium-melt-calculator

📄 License
This project is licensed under the MIT License - see the LICENSE file for details.

🙏 Acknowledgments
Kd Values: Default values from Grassi et al. (2012)
Normalizing Values: Default PM values from McDonough & Sun (1995)
Geochemical Community: For partition coefficient databases

📚 References

Grassi, D., Schmidt, M.W. and Günther, D., 2012. Element partitioning during carbonated pelite melting at 8, 13 and 22 GPa and the sediment signature in the EM mantle components. Earth and Planetary Science Letters, 327, pp.84-96.
McDonough, W.F. and Sun, S.S., 1995. The composition of the Earth. Chemical geology, 120(3-4), pp.223-253.
