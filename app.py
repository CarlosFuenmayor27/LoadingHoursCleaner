import os
import pandas as pd
from flask import Flask, request, render_template, send_file
from datetime import datetime
import re
from functools import reduce

# Initialize Flask app
app = Flask(__name__)

# Define employees to exclude during processing
EXCLUDED_EMPLOYEES = {
    'Total', 'Consultants', 'Associates', 'Consultant', 'Associate', '0',
    'Total Worked/Forecasted', 'Hours/mo', 'Available', '% Loading Goal'
}

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        uploaded_file = request.files['file']
        if uploaded_file.filename.endswith('.xlsx'):
            file_path = uploaded_file.filename
            uploaded_file.save(file_path)
            output_file = process_excel(file_path)
            return render_template('upload.html', success=True)
        else:
            return 'Only .xlsx files are supported.'
    return render_template('upload.html', success=False)

@app.route('/download')
def download():
    return send_file("Loading_Hours.xlsx", as_attachment=True)

def process_excel(file_path):
    xl = pd.ExcelFile(file_path)
    final_df = pd.DataFrame()
    match = re.search(r'20\d{2}', os.path.basename(file_path))
    start_year = int(match.group()) if match else 2024

    def is_project_sheet(raw_df, sheet_name):
        name = sheet_name.lower()
        if any(kw in name for kw in ['summary-hours', 'future', 'chart']):
            return False
        for i, row in raw_df.iterrows():
            text_cells = sum(1 for v in row if isinstance(v, str) and v.strip())
            num_cells = sum(1 for v in row if isinstance(v, (int, float)) and v != 0)
            if text_cells >= 1 and num_cells >= 1:
                return True
        return False

    for sheet in xl.sheet_names:
        try:
            raw = xl.parse(sheet, header=None)
            if not is_project_sheet(raw, sheet):
                continue

            month_keywords = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
                              'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
            month_row_idx = None
            for i, row in raw.iterrows():
                values = [str(v).lower()[:3] for v in row if pd.notna(v)]
                if any(m in values for m in month_keywords):
                    month_row_idx = i
                    break
            if month_row_idx is None:
                continue

            month_row = raw.iloc[month_row_idx]
            headers = []
            current_year = start_year
            previous_month_index = -1
            for mo in month_row:
                if pd.isna(mo):
                    headers.append(None)
                    continue
                m = str(mo).strip().replace('.', '')[:3].lower()
                if m not in month_keywords:
                    headers.append(None)
                    continue
                month_index = month_keywords.index(m)
                if previous_month_index != -1 and month_index < previous_month_index:
                    current_year += 1
                previous_month_index = month_index
                headers.append(f"{m.capitalize()} {current_year}")
            if len(headers) < 2:
                continue
            headers[1] = 'Employee'

            df_data = raw.iloc[month_row_idx + 1:].copy()
            df_data.columns = headers

            JOB_TYPE_LABELS = {
                'consultant': 'Consultant',
                'consultants': 'Consultant',
                'associate': 'Associate',
                'associates': 'Associate'
            }
            job_type_col = []
            current_job_type = 'Unknown'
            for idx, row in df_data.iterrows():
                val_b = str(row.get('Employee', '')).strip().lower()
                val_a = str(row.iloc[0]).strip().lower()
                if val_b in JOB_TYPE_LABELS:
                    current_job_type = JOB_TYPE_LABELS[val_b]
                elif val_a in JOB_TYPE_LABELS:
                    current_job_type = JOB_TYPE_LABELS[val_a]
                job_type_col.append(current_job_type)
            df_data['Job Type'] = job_type_col

            df_data = df_data[df_data['Employee'].notna()]
            df_data = df_data[~df_data['Employee'].astype(str).str.strip().isin(EXCLUDED_EMPLOYEES)]

            date_cols = [col for col in df_data.columns if col not in [None, 'Employee', 'Job Type']]
            df_melted = df_data.melt(id_vars=['Employee', 'Job Type'],
                                     value_vars=date_cols,
                                     var_name='Date', value_name='Hours')
            df_melted['Hours'] = pd.to_numeric(df_melted['Hours'], errors='coerce').fillna(0)

            def get_work_type(date_str):
                try:
                    dt = datetime.strptime(date_str, "%b %Y")
                    return 'Predicted' if dt > datetime.today().replace(day=1) else 'Actual'
                except:
                    return 'Unknown'

            df_melted['Work Type'] = df_melted['Date'].apply(get_work_type)
            df_melted['Project'] = sheet
            df_melted['Job Type'] = df_melted['Job Type'].replace('Unknown', 'Consultant')

            final_df = pd.concat([final_df, df_melted], ignore_index=True)
        except Exception as e:
            print(f"Error in sheet '{sheet}': {e}")
            continue

    goal_table = pd.DataFrame()
    try:
        summary_df = xl.parse("Summary-hours", header=None)
        goal_row_idx, goal_col_idx = None, None
        for i, row in summary_df.iterrows():
            for j, cell in enumerate(row):
                if isinstance(cell, str) and "% loading goal" in cell.strip().lower():
                    goal_row_idx, goal_col_idx = i, j
                    break
            if goal_row_idx is not None:
                break
        if goal_row_idx is None:
            raise ValueError("Loading Goal row not found.")
        goal_row = summary_df.iloc[goal_row_idx]

        month_row_idx = None
        month_keywords = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
                          'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
        for i, row in summary_df.iterrows():
            if sum(1 for cell in row if isinstance(cell, str) and cell.strip()[:3].lower() in month_keywords) >= 6:
                month_row_idx = i
                break
        if month_row_idx is None:
            raise ValueError("Month label row not found.")
        month_row = summary_df.iloc[month_row_idx]

        month_labels = []
        values = []
        current_year = start_year
        previous_month_index = -1

        for col_idx in range(len(goal_row)):
            raw_month = str(month_row[col_idx]).strip()
            if not raw_month:
                continue
            abbr = raw_month[:3].capitalize()
            if abbr.lower() not in month_keywords:
                continue
            month_index = month_keywords.index(abbr.lower())
            if previous_month_index != -1 and month_index < previous_month_index:
                current_year += 1
            previous_month_index = month_index
            val = goal_row[col_idx]
            if isinstance(val, (int, float)):
                month_labels.append(f"{abbr} {current_year}")
                values.append(val)

        goal_table = pd.DataFrame({
            "Month": month_labels,
            "Loading Goal": values
        })

    except Exception as e:
        print(f"Error generating % Goal table: {e}")

    loading_sheets = {}
    for sheet in xl.sheet_names:
        year_match = re.search(r'20\d{2}', sheet.lower())
        if 'chart' not in sheet.lower() or not year_match:
                    continue
        year = year_match.group(0)
        try:
            df = xl.parse(sheet, header=None, skiprows=7).dropna(how='all')
            df = df[[5, 6, 7, 8]].copy()
            df.columns = [
                'Employee',
                f'Projected {year}',
                f'Desired {year}',
                f'% Loading {year}'
            ]
            df['Employee'] = df['Employee'].astype(str).str.strip().str.title()
            df.dropna(subset=['Employee'], inplace=True)
            df.drop(df[df['Employee'].str.lower().str.contains(
                "avg|average|target|estimate|difference|^0$|^nan$", na=False)].index, inplace=True)
            loading_sheets[year] = df
        except Exception as e:
            print(f"Could not process loading sheet '{sheet}': {e}")

    loading_table = pd.DataFrame()
    sorted_years = sorted(loading_sheets.keys())
    if sorted_years:
        loading_table = reduce(lambda left, right: pd.merge(left, right, on='Employee', how='outer'),
                               loading_sheets.values())
        column_map = {}
        if len(sorted_years) >= 1:
            column_map.update({
                f'Projected {sorted_years[0]}': 'Projected First Year',
                f'Desired {sorted_years[0]}': 'Desired First Year',
                f'% Loading {sorted_years[0]}': 'Loading First Year',
            })
        if len(sorted_years) >= 2:
            column_map.update({
                f'Projected {sorted_years[1]}': 'Projected Second Year',
                f'Desired {sorted_years[1]}': 'Desired Second Year',
                f'% Loading {sorted_years[1]}': 'Loading Second Year',
            })
        loading_table.rename(columns=column_map, inplace=True)
        for col in ['Projected Second Year', 'Desired Second Year', 'Loading Second Year']:
            if col not in loading_table.columns:
                loading_table[col] = 0
        loading_table = loading_table[loading_table['Employee'].astype(str).str.strip() != '']

    output_path = "Loading_Hours.xlsx"
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, sheet_name="Project Hours", index=False)
        loading_table.to_excel(writer, sheet_name="Loading Comparison", index=False)
        goal_table.to_excel(writer, sheet_name="Loading Goal", index=False)

    return output_path

if __name__ == '__main__':
    app.run(debug=True)
