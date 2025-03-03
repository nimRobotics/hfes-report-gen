#!/usr/bin/env python3
import pandas as pd
import os
from datetime import datetime
import argparse
import pdflatex as ptex

OBJECTIVES = 4

def tex_to_pdf(tex_file, filename):
    pdfl = ptex.PDFLaTeX.from_texfile(tex_file)
    pdf, log, completed_process = pdfl.create_pdf()
    with open(filename, 'wb') as pdfout:
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

def get_other_goals(df):
    """Extract uncategorized goals from the data"""
    other_goals = set()
    for i in range(1, OBJECTIVES+1):  # For objectives 1-4
        # Objective 1 - Which HFES strategic goal(s) does this support? [Select all that apply]
        col_name = f"Objective {i} - Which HFES strategic goal(s) does this support? [Select all that apply]"
        values = df[col_name].dropna().unique()
        for value in values:
            if isinstance(value, str) and value.strip() != "":
                other_goals.add(value)
    return sorted(list(other_goals))

def get_operational_goals(df):
    """Extract unique operational goals from the data"""
    op_goals = set()
    for i in range(1, OBJECTIVES+1):  # For objectives 1-4
        col_name = f"Objective {i} - Q2 What is the HFES operational strategic goal does that this objective primarily supports?"
        values = df[col_name].dropna().unique()
        for value in values:
            if isinstance(value, str) and value.strip() != "":
                op_goals.add(value)
    return sorted(list(op_goals))

def get_transformation_goals(df):
    """Extract unique transformation goals from the data"""
    tr_goals = set()
    for i in range(1, OBJECTIVES+1):  # For objectives 1-4
        col_name = f"Objective {i} - Q3 What is the HFES transformation strategic goal does that this objective primarily supports?"
        values = df[col_name].dropna().unique()
        for value in values:
            if isinstance(value, str) and value.strip() != "" and value.lower() != "none":
                tr_goals.add(value)
    return sorted(list(tr_goals))

def get_committees_by_goal(df, goal_type, goal):
    """Get all committees that have objectives supporting a specific goal"""
    committees = set()
    for i in range(1, OBJECTIVES+1):  # For objectives 1-4
        if goal_type == "operational":
            goal_col = f"Objective {i} - Q2 What is the HFES operational strategic goal does that this objective primarily supports?"
        else:  # transformation
            goal_col = f"Objective {i} - Q3 What is the HFES transformation strategic goal does that this objective primarily supports?"
        
        # Filter rows where the goal matches
        matches = df[df[goal_col] == goal]
        
        for _, row in matches.iterrows():
            committee = row["Committee, Group, or Task Force"]
            if isinstance(committee, str) and committee.strip() != "":
                committees.add(committee)
    
    return sorted(list(committees))

def get_objectives_for_committee_and_goal(df, committee, goal_type, goal):
    """Get all objectives for a specific committee supporting a specific goal"""
    objectives = []
    
    # Filter rows for this committee
    committee_rows = df[df["Committee, Group, or Task Force"] == committee]
    
    for _, row in committee_rows.iterrows():
        for i in range(1, 5):  # For objectives 1-4
            if goal_type == "operational":
                goal_col = f"Objective {i} - Q2 What is the HFES operational strategic goal does that this objective primarily supports?"
            else:  # transformation
                goal_col = f"Objective {i} - Q3 What is the HFES transformation strategic goal does that this objective primarily supports?"
            
            # Check if this objective supports the specified goal
            if row.get(goal_col) == goal:
                objective = {
                    "description": row.get(f"Objective {i} - Q1 Describe the objective and what impact has/will achieving this objective have on HFES. Please use as much text as needed."),
                    "work_done": row.get(f"Objective {i} -  Q4 What steps have you taken over the past 6 months to move you towards achieving this objective? Please be as descriptive as necessary."),
                    "work_planned": row.get(f"Objective {i} - Q9 What steps do you plan to take in the next 6 months towards achieving this objective?"),
                    "achieved": row.get(f"Objective {i} - Q7 Have you achieved this objective?"),
                    "target_date": row.get(f"Objective {i} - Q8 What is your target date for achieving this goal [DATE or N/A if it is an ongoing goal]"),
                    "challenges": row.get(f"Objective {i} - Q10 What  challenges/barriers are there for completing your goal by the target date set? (N/A for none)")
                }
                
                # Only add if we have actual content
                if objective["description"] and isinstance(objective["description"], str) and objective["description"].strip() != "":
                    objectives.append(objective)
    
    return objectives

def escape_tex(text):
    """Escape special LaTeX characters in text"""
    if not isinstance(text, str):
        return ""
    
    # Characters to escape: # $ % & _ { } ~ ^ \ 
    escapes = {
        '&': '\\&',
        '%': '\\%',
        '$': '\\$',
        '#': '\\#',
        '_': '\\_',
        '{': '\\{',
        '}': '\\}',
        '~': '\\textasciitilde{}',
        '^': '\\textasciicircum{}',
        '\\': '\\textbackslash{}',
        '---': '--',  # en-dash
    }
    
    for char, replacement in escapes.items():
        text = text.replace(char, replacement)
    
    return text

def is_text_na(text):
    """
    Check if text is 'N/A' or empty
    check na, N/A, n/a, empty string, or whitespace, case insensitive

    return True if text is 'N/A' or empty, False otherwise
    """
    text = str(text)
    return text.strip().lower() in ["n/a", "na", "nan", ""]


def generate_report(df, output_file):
    """Generate the LaTeX report from the dataframe"""
    # Get the year and report period from the first row
    # year = escape_tex(str(df.iloc[0]["Year"]))
    # report_period = escape_tex(str(df.iloc[0]["Report period"]))
    year = "2021"
    report_period = "Q1-Q2"
    
    # Start building the LaTeX document
    tex = r"""
    \documentclass{article}
    \usepackage[english]{babel}
    \usepackage[letterpaper,margin=1in]{geometry}
    \usepackage{amsmath}
    \usepackage{graphicx}
    \usepackage[colorlinks=true, allcolors=blue]{hyperref}

    \usepackage[utf8]{inputenc}
    \usepackage{xcolor}
    \usepackage{titlesec}
    \usepackage{enumitem}
    \hypersetup{
        colorlinks=true,
        linkcolor=blue,
        filecolor=magenta,
        urlcolor=blue,
    }

    \title{Human Factors and Ergonomics Society}
    \author{Report on Strategic Goals}
    \date{""" + year + " " + report_period + r"""}

    \begin{document}
    \maketitle
    """
    
    # Get operational goals
    op_goals = get_operational_goals(df)
    tex += r"\section{OPERATIONAL STRATEGIC GOALS}" + "\n\n"
    
    # Process each operational goal
    for goal_idx, goal in enumerate(op_goals, 1):
        tex += r"\subsection{" + escape_tex(goal) + "}\n\n"
        
        # Get committees for this goal
        committees = get_committees_by_goal(df, "operational", goal)
        
        # Process each committee for this goal
        for committee_idx, committee in enumerate(committees, ord('a')):
            tex += r"\subsubsection{" + escape_tex(committee) + "}\n\n"
            
            # Get objectives for this committee and goal
            objectives = get_objectives_for_committee_and_goal(df, committee, "operational", goal)
            
            # Process each objective
            for obj_idx, objective in enumerate(objectives, 1):
                tex += r"\paragraph{Objective " + str(obj_idx) + r"} " + escape_tex(objective['description']) + "\n\n"
                
                if objective['work_done']:
                    tex += r"\textbf{Work Done:} " + escape_tex(str(objective['work_done'])) + "\n\n"
                
                if objective['work_planned']:
                    tex += r"\textbf{Work Planned:} " + escape_tex(str(objective['work_planned'])) + "\n\n"
                
                if objective['achieved']:
                    tex += r"\textbf{Objective achieved:} " + escape_tex(str(objective['achieved']))
                    if objective['target_date'] and not is_text_na(objective['target_date']):
                        tex += " (Target: " + escape_tex(str(objective['target_date'])) + ")"
                    tex += "\n\n"
                
                if objective['challenges'] and not is_text_na(objective['challenges']):
                    tex += r"\textbf{Challenges:} " + escape_tex(str(objective['challenges'])) + "\n\n"
                
                tex += "\n\n"
    
    # Get transformation goals
    tr_goals = get_transformation_goals(df)
    if tr_goals:
        tex += r"\section{TRANSFORMATION STRATEGIC GOALS}" + "\n\n"
        
        # Process each transformation goal
        for goal_idx, goal in enumerate(tr_goals, 1):
            tex += r"\subsection{" + escape_tex(goal) + "}\n\n"
            
            # Get committees for this goal
            committees = get_committees_by_goal(df, "transformation", goal)
            
            # Process each committee for this goal
            for committee_idx, committee in enumerate(committees, ord('a')):
                tex += r"\subsubsection{" + escape_tex(committee) + "}\n\n"
                
                # Get objectives for this committee and goal
                objectives = get_objectives_for_committee_and_goal(df, committee, "transformation", goal)
                
                # Process each objective
                for obj_idx, objective in enumerate(objectives, 1):
                    tex += r"\paragraph{Objective " + str(obj_idx) + r"} " + escape_tex(objective['description']) + "\n\n"
                    
                    if objective['work_done']:
                        tex += r"\textbf{Work Done:} " + escape_tex(str(objective['work_done'])) + "\n\n"
                    
                    if objective['work_planned']:
                        tex += r"\textbf{Work Planned:} " + escape_tex(str(objective['work_planned'])) + "\n\n"
                    
                    if objective['achieved']:
                        tex += r"\textbf{Objective achieved:} " + escape_tex(str(objective['achieved']))
                        if objective['target_date'] and not is_text_na(objective['target_date']):
                            tex += " (Target: " + escape_tex(str(objective['target_date'])) + ")"
                        tex += "\n\n"
                    
                    if objective['challenges'] and not is_text_na(objective['challenges']):
                        tex += r"\textbf{Challenges:} " + escape_tex(str(objective['challenges'])) + "\n\n"
                    
                    tex +=  "\n\n"
    
    # Generate a date stamp at the end
    tex += r"\vspace{1cm}" + "\n"
    tex += r"\noindent\textit{Report generated on " + datetime.now().strftime('%Y-%m-%d') + "}" + "\n"
    
    # Close the document
    tex += r"\end{document}"
    
    # Write the LaTeX to file
    with open(output_file, 'w') as f:
        f.write(tex)
    
    print(f"LaTeX report generated: {output_file}")
    return tex

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Generate HFES Strategic Goals Report")
    parser.add_argument("datafile", help="Path to the input data file (CSV or Excel)")
    parser.add_argument("-o", "--output", default="HFES_Strategic_Goals_Report.tex", help="Path to the output LaTeX file")
    args = parser.parse_args()
    
    # Load data
    df = load_data(args.datafile)
    
    # Generate report
    generate_report(df, args.output)

    # Convert LaTeX to PDF
    pdf_file = args.output.replace(".tex", ".pdf")
    print(f"Converting LaTeX to PDF: {pdf_file}")
    print(f"PDF report generated: {pdf_file}")
    
    # Suggest pdflatex command
    print("\nTo convert the LaTeX file to PDF, use the following command:")
    print(f"pdflatex {args.output}")

if __name__ == "__main__":
    main()
