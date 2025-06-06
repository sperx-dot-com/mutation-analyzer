@echo off
echo Multiple Sequence Alignment Tool
echo ==============================

if "%1"=="" (
    echo Usage: run_analysis.bat [reference.fasta] [ab1_directory]
    exit /b 1
)

if "%2"=="" (
    echo Usage: run_analysis.bat [reference.fasta] [ab1_directory]
    exit /b 1
)

echo Running sequence alignment...
python sequence_aligner.py --ref %1 --input %2 --output mutations.xlsx

echo Generating reports...
python report_generator.py --input mutations.xlsx --output mutation_report.xlsx --plots mutation_plots

echo Analysis complete!
echo Results saved to mutation_report.xlsx
echo Plots saved to mutation_plots directory