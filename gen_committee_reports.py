import pandas as pd
import os
from datetime import datetime
import argparse
import pdflatex as ptex 
from pathlib import Path


def tex_to_pdf(tex_file, filename):
    print(f"Converting {tex_file} to PDF {filename}")
    pdfl = ptex.PDFLaTeX.from_texfile(tex_file)
    pdf, log, completed_process = pdfl.create_pdf()
    with open("test.pdf", 'wb') as pdfout:
        pdfout.write(pdf)

def load_data(file_path):
    """Load and prepare data from CSV or Excel file"""
    _, file_extension = os.path.splitext(file_path)
    if file_extension.lower() == '.csv':
        df = pd.read_csv(file_path)
    elif file_extension.lower() in ['.xls', '.xlsx']:
        df = pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format. Please provide a CSV or Excel file.")
    return df

def generate_latex_reports(df, output_dir):
    
    # # Print column names to debug
    # print("Available columns in CSV:")
    # for i, col in enumerate(df.columns):
    #     print(f"{i}: {col}")
    
    # Generate a report for each row (committee)
    for index, row in df.iterrows():
        # Skip row if committee name is missing
        if 'Committee, Group, or Task Force' not in row or pd.isna(row['Committee, Group, or Task Force']):
            print(f"Skipping row {index} - missing committee name")
            continue
            
        committee_name = row['Committee, Group, or Task Force']
        # Clean the committee name for use in filenames
        clean_name = ''.join(c if c.isalnum() else '_' for c in committee_name)
        filename = f"{output_dir}/{clean_name}_report.tex"
        
        # Safely get values with fallbacks
        def safe_get(row, key, default="N/A"):
            if key in row and pd.notna(row[key]):
                return str(row[key])
            return default
        
        with open(filename, 'w', encoding='utf-8') as f:
            # Write LaTeX document header
            f.write(r'''\documentclass{article}
\usepackage[margin=1in]{geometry}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{xcolor}
\usepackage{hyperref}
\usepackage{enumitem}

\hypersetup{
    colorlinks=true,
    linkcolor=blue,
    filecolor=magenta,
    urlcolor=blue,
}

\title{Committee Report: ''' + committee_name + r'''}
\author{Human Factors and Ergonomics Society}
\date{''' + datetime.now().strftime("%B %d, %Y") + r'''}

\begin{document}

\maketitle

\section{Committee Information}
\begin{itemize}[leftmargin=*]
    \item \textbf{Reporting Period:} ''' + safe_get(row, 'Report period') + r'''
    \item \textbf{Year:} ''' + safe_get(row, 'Year') + r'''
    \item \textbf{Chair:} ''' + safe_get(row, 'Name of Chair') + r'''
    \item \textbf{Members:} ''' + safe_get(row, 'Members separated by commas (or N/A if no other members than the chair).') + r'''
    \item \textbf{Division:} ''' + safe_get(row, 'Division') + r'''
\end{itemize}

\section{Committee Assessment}
\begin{itemize}[leftmargin=*]
''')
            
            # Look for the alignment score column (flexible matching)
            alignment_score_col = [col for col in df.columns if 'aligned' in col.lower() and 'operating rules' in col.lower()]
            if alignment_score_col:
                f.write(r'    \item \textbf{Alignment with Operating Rules:} ' + safe_get(row, alignment_score_col[0]) + r' / 5' + '\n')
            
            # Look for the alignment improvement column
            alignment_improve_col = [col for col in df.columns if 'score of 3 or lower' in col.lower() and 'align' in col.lower()]
            if alignment_improve_col and pd.notna(row.get(alignment_improve_col[0], None)):
                f.write(r'    \item \textbf{Alignment Improvement Needed:} ' + safe_get(row, alignment_improve_col[0]) + '\n')
            
            # Look for the functioning score column
            functioning_col = [col for col in df.columns if 'functioning' in col.lower() and 'score' in col.lower()]
            if functioning_col:
                f.write(r'    \item \textbf{Overall Functioning:} ' + safe_get(row, functioning_col[0]) + r' / 5' + '\n')
            
            # Look for the functioning improvement column
            functioning_improve_col = [col for col in df.columns if 'reported a value of 3 or less' in col.lower() and 'functioning' in col.lower()]
            if functioning_improve_col and pd.notna(row.get(functioning_improve_col[0], None)):
                f.write(r'    \item \textbf{Functioning Improvement Needed:} ' + safe_get(row, functioning_improve_col[0]) + '\n')
            
            f.write(r'''\end{itemize}

\section{Resource Requirements}
\begin{itemize}[leftmargin=*]
''')
            # Non-budgetary resources
            non_budget_col = [col for col in df.columns if 'non-budgetary' in col.lower()]
            if non_budget_col and pd.notna(row.get(non_budget_col[0], None)):
                f.write(r'    \item \textbf{Non-budgetary Resources:} ' + safe_get(row, non_budget_col[0]) + '\n')
            
            # Budgetary requests
            budget_col = [col for col in df.columns if 'budgetary request' in col.lower()]
            if budget_col and pd.notna(row.get(budget_col[0], None)):
                f.write(r'    \item \textbf{Budgetary Requests:} ' + safe_get(row, budget_col[0]) + '\n')
            
            # Executive Council requests
            exec_col = [col for col in df.columns if 'executive council' in col.lower() or 'actions to executive' in col.lower()]
            if exec_col and pd.notna(row.get(exec_col[0], None)):
                f.write(r'    \item \textbf{Executive Council Requests:} ' + safe_get(row, exec_col[0]) + '\n')
            
            # Time commitments
            chair_time_col = [col for col in df.columns if 'hours' in col.lower() and 'chair' in col.lower()]
            member_time_col = [col for col in df.columns if 'hours' in col.lower() and 'member' in col.lower()]
            
            if chair_time_col:
                f.write(r'    \item \textbf{Chair Time Commitment:} ' + safe_get(row, chair_time_col[0]) + r' hours per month' + '\n')
            
            if member_time_col:
                f.write(r'    \item \textbf{Member Time Commitment:} ' + safe_get(row, member_time_col[0]) + r' hours per month' + '\n')
            
            f.write(r'''\end{itemize}

\section{Objectives}
''')

            # Function to write objectives
            def write_objective(f, row, obj_num):
                obj_desc_col = [col for col in df.columns if f'Objective {obj_num} - Q1' in col]
                if not obj_desc_col or pd.isna(row.get(obj_desc_col[0], None)):
                    return False
                
                f.write(f'''\\subsection{{Objective {obj_num}}}
\\begin{{itemize}}[leftmargin=*]
''')
                
                # Description
                f.write(r'    \item \textbf{Description:} ' + safe_get(row, obj_desc_col[0]) + '\n')
                
                # Operational Strategic Goal
                op_goal_col = [col for col in df.columns if f'Objective {obj_num} - Q2' in col]
                if op_goal_col:
                    f.write(r'    \item \textbf{Operational Strategic Goal:} ' + safe_get(row, op_goal_col[0]) + '\n')
                
                # Transformation Strategic Goal
                trans_goal_col = [col for col in df.columns if f'Objective {obj_num} - Q3' in col]
                if trans_goal_col:
                    f.write(r'    \item \textbf{Transformation Strategic Goal:} ' + safe_get(row, trans_goal_col[0]) + '\n')
                
                # Progress
                progress_col = [col for col in df.columns if f'Objective {obj_num} -  Q4' in col]
                if progress_col:
                    f.write(r'    \item \textbf{Progress in Past 6 Months:} ' + safe_get(row, progress_col[0]) + '\n')
                
                # Challenges
                challenges_col = [col for col in df.columns if f'Objective {obj_num} -  Q5' in col]
                if challenges_col and pd.notna(row.get(challenges_col[0], None)):
                    f.write(r'    \item \textbf{Challenges/Barriers:} ' + safe_get(row, challenges_col[0]) + '\n')
                
                # Metrics
                metrics_col = [col for col in df.columns if f'Objective {obj_num} -  Q6' in col]
                if metrics_col and pd.notna(row.get(metrics_col[0], None)):
                    f.write(r'    \item \textbf{Metrics:} ' + safe_get(row, metrics_col[0]) + '\n')
                
                # Achieved
                achieved_col = [col for col in df.columns if f'Objective {obj_num} - Q7' in col]
                if achieved_col:
                    f.write(r'    \item \textbf{Achieved:} ' + safe_get(row, achieved_col[0]) + '\n')
                
                # Target Date
                target_col = [col for col in df.columns if f'Objective {obj_num} - Q8' in col]
                if target_col and pd.notna(row.get(target_col[0], None)):
                    f.write(r'    \item \textbf{Target Date:} ' + safe_get(row, target_col[0]) + '\n')
                
                # Next 6 Months Plan
                plan_col = [col for col in df.columns if f'Objective {obj_num} - Q9' in col]
                if plan_col:
                    f.write(r'    \item \textbf{Next 6 Months Plan:} ' + safe_get(row, plan_col[0]) + '\n')
                
                # Future Challenges
                future_col = [col for col in df.columns if f'Objective {obj_num} - Q10' in col]
                if future_col and pd.notna(row.get(future_col[0], None)):
                    f.write(r'    \item \textbf{Future Challenges:} ' + safe_get(row, future_col[0]) + '\n')
                
                f.write(r'''\end{itemize}
''')
                return True
            
            # Write objectives
            write_objective(f, row, 1)
            
            # Check if additional objectives exist
            obj2_col = [col for col in df.columns if 'Do you have 2nd objective' in col]
            if obj2_col and row.get(obj2_col[0]) == 'Yes':
                write_objective(f, row, 2)
            
            obj3_col = [col for col in df.columns if 'Do you have 3rd objective' in col]
            if obj3_col and row.get(obj3_col[0]) == 'Yes':
                write_objective(f, row, 3)
            
            obj4_col = [col for col in df.columns if 'Do you have 4th objective' in col]
            if obj4_col and row.get(obj4_col[0]) == 'Yes':
                write_objective(f, row, 4)
            
            # Add report feedback section
            f.write(r'''\section{Report Feedback}
\begin{itemize}[leftmargin=*]
''')
            
            # Difficulty rating
            difficulty_col = [col for col in df.columns if 'difficulty of completing the report' in col.lower()]
            if difficulty_col:
                f.write(r'    \item \textbf{Difficulty of Completing Report:} ' + safe_get(row, difficulty_col[0]) + r' / 5' + '\n')
            
            # Comments
            comments_col = [col for col in df.columns if 'comments about the reporting process' in col.lower()]
            if comments_col and pd.notna(row.get(comments_col[0], None)):
                f.write(r'    \item \textbf{Comments on Reporting Process:} ' + safe_get(row, comments_col[0]) + '\n')
            
            f.write(r'''\end{itemize}

\end{document}
''')
        
        print(f"Created report for {committee_name}: {filename}")
    
    print(f"Generated {len(df)} committee reports in the '{output_dir}' directory.")

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Generate HFES Strategic Goals Report")
    parser.add_argument("datafile", help="Path to the input data file (CSV or Excel)")
    parser.add_argument("-o", "--output", default="committee_reports", help="Dir to the output LaTeX files")
    args = parser.parse_args()
    
    # Load data
    df = load_data(args.datafile)
    
    # Generate report
    generate_latex_reports(df, args.output)

    # Convert LaTeX to PDF
    tex_files = Path(args.output).rglob("*.tex")
    for tex_file in tex_files:
        print(f"Converting {tex_file} to PDF")
        tex_to_pdf(tex_file, tex_file.with_suffix(".pdf"))


if __name__ == "__main__":
    main()