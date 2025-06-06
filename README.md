# Sequence Alignment Tool

A comprehensive tool for analyzing DNA sequence alignments, identifying mutations, and generating detailed reports.

## Overview

This tool processes DNA sequence files (AB1 format), aligns them against a reference sequence, identifies mutations, and generates detailed reports with visualizations. It's designed for researchers and lab technicians working with DNA sequencing data.

## Features

- Sequence alignment against a reference FASTA file
- Mutation detection and classification (silent vs. missense)
- Codon and amino acid analysis
- Variant identification across samples
- Comprehensive Excel reports with:
  - Detailed mutation listings
  - Summary statistics
  - Codon-level analysis
  - Variant grouping
- Visualization plots:
  - Mutation distribution by sample
  - Silent vs. missense mutation ratios
  - Mutation positions along the sequence

## Requirements

- Python 3.6+
- Required Python packages:
  - pandas
  - matplotlib
  - seaborn
  - openpyxl
  - biopython (for sequence processing)

## Usage

Run the analysis using the provided batch file:

run_analysis.bat [reference.fasta] [ab1_directory]

Where:
- `[reference.fasta]` is your reference sequence file
- `[ab1_directory]` is the directory containing your AB1 sequence files

## Output

The tool generates:
1. `mutations.xlsx` - Raw mutation data
2. `mutation_report.xlsx` - Formatted report with multiple analysis sheets
3. `mutation_plots/` - Directory containing visualization plots

## Example

run_analysis.bat reference.fasta samples/

## Report Contents

The generated Excel report contains multiple sheets:
- **Mutation Summary**: Detailed listing of all mutations by sample
- **Summary Statistics**: Overview of mutation counts and types
- **Codon Analysis**: Analysis of mutations at the codon level
- **Variant Analysis**: Grouping of samples by mutation patterns

## License

[Your license information here]

## Contact

[Your contact information here]
