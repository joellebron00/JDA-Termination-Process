
import pandas as pd
from io import BytesIO

TERMINATIONS_COLUMNS = [
    'Reason For Leaving', 'Unique Id', 'HR Request: Created By',
    'Team Member', 'One Up Manager', 'Priority State',
    'Worked State', 'Eligible to Return to Program',
]

def _validate_columns(df, required, file_label):
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(
            f"El archivo de {file_label} no tiene estas columnas: {', '.join(missing)}. "
            "Verifica que subiste el archivo correcto."
        )

def load_terminations(file):
    try:
        df = pd.read_excel(file)
    except Exception as e:
        raise ValueError("No se pudo leer el archivo de Sage. Verifica que sea un .xlsx válido.") from e

    df.columns = df.columns.str.strip()
    _validate_columns(df, TERMINATIONS_COLUMNS, "Sage")
    return df
###################

RESPONSES_COLUMNS = [
    'Completion time',
    "Expert's / Lead's Partner JDA Employee Number",
    'Would you recommend this expert/lead for rehire back to TTL?',
]

def load_responses(file):
    try:
        df = pd.read_excel(file)
    except Exception as e:
        raise ValueError("No se pudo leer el archivo del Form. Verifica que sea un .xlsx válido.") from e

    df.columns = df.columns.str.strip()
    _validate_columns(df, RESPONSES_COLUMNS, "respuestas del Form")
    return df

def normalize_employee_id(series):
    return (
        series
        .astype('string')
        .str.strip()
        .str.replace(r'^JDA[\s-]*', '', regex=True, case=False)
        .str.replace(r'\.0$', '', regex=True)
        .str.strip()
        .replace('', pd.NA)
    )
################

TERMINATIONS_OUTPUT_COLUMNS = [
    'HR Request: Created By', 'Unique Id', 'Team Member',
    'One Up Manager', 'Priority State', 'Worked State',
    'Eligible to Return to Program',
]

ELIGIBLE_MAP = {
    'true': 'Yes', '1': 'Yes', '1.0': 'Yes',
    'false': 'No', '0': 'No', '0.0': 'No',
}

def clean_terminations(df):
    is_end_of_assignment = (
        df['Reason For Leaving'].astype(str).str.strip() == 'End of Assignment'
    )
    df = df[is_end_of_assignment].copy()

    df['Unique Id'] = normalize_employee_id(df['Unique Id'])
    df = df.drop_duplicates(subset='Unique Id')

   
    df['Eligible to Return to Program'] = (
        df['Eligible to Return to Program']
        .astype(str).str.strip().str.lower()
        .map(ELIGIBLE_MAP)
    )

    return df[TERMINATIONS_OUTPUT_COLUMNS]

############

RESPONSES_OUTPUT_COLUMNS = [
    'Completion time', 'employee_number',
    'Would you recommend this expert/lead for rehire back to TTL?',
]

def clean_responses(df):
    df = df.copy()

    df['Completion time'] = pd.to_datetime(df['Completion time'], errors='coerce')
    df['employee_number'] = normalize_employee_id(
        df["Expert's / Lead's Partner JDA Employee Number"]
    )
    df = df.dropna(subset=['employee_number'])

    df = (
        df.sort_values('Completion time', na_position='first')
          .drop_duplicates(subset='employee_number', keep='last')
    )

    return df[RESPONSES_OUTPUT_COLUMNS]

####################

COMPLETED_COLUMNS = [
    'One Up Manager', 'HR Request: Created By', 'Unique Id', 'Team Member',
    'Priority State', 'Worked State',
    'Would you recommend this expert/lead for rehire back to TTL?',
    'Eligible to Return to Program',
]

INCOMPLETE_COLUMNS = [
    'One Up Manager', 'HR Request: Created By', 'Unique Id',
    'Team Member', 'Priority State', 'Worked State',
]

def merge_data(terminations, responses):
    return pd.merge(
        terminations, responses,
        left_on='Unique Id', right_on='employee_number',
        how='left', validate='one_to_one',
    )

def split_forms(merged):
    has_form = merged['employee_number'].notna()
    completed = merged.loc[has_form, COMPLETED_COLUMNS].reset_index(drop=True)
    incomplete = merged.loc[~has_form, INCOMPLETE_COLUMNS].reset_index(drop=True)
    return completed, incomplete

#################

MANAGER_COLUMN = 'HR Request: Created By'

def summary_by_manager(completed, incomplete):
    completed_count = completed.groupby(MANAGER_COLUMN, dropna=False).size().rename('completed_forms')
    incomplete_count = incomplete.groupby(MANAGER_COLUMN, dropna=False).size().rename('incomplete_forms')
    priority_incomplete = (
        incomplete[incomplete['Priority State'] == 'Yes']
        .groupby(MANAGER_COLUMN, dropna=False).size()
        .rename('incomplete_priority')
    )

    summary = (
        pd.concat([completed_count, incomplete_count, priority_incomplete], axis=1)
        .fillna(0)
        .astype(int)
    )
    summary['total_forms'] = summary['completed_forms'] + summary['incomplete_forms']
    summary['completion_rate'] = (summary['completed_forms'] / summary['total_forms'] * 100).round(1)

    return summary.sort_values('incomplete_forms', ascending=False).reset_index()


##########

def compute_kpis(completed, incomplete):
    total = len(completed) + len(incomplete)
    completion_rate = round(len(completed) / total * 100, 1) if total else 0.0

    return {
        'total_terminations': total,
        'completed_forms': len(completed),
        'incomplete_forms': len(incomplete),
        'completion_rate': completion_rate,
        'incomplete_priority': int((incomplete['Priority State'] == 'Yes').sum()),
    }


def to_excel_bytes(sheets):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return buffer.getvalue()


def run_pipeline(terminations_file, responses_file):
    terminations = clean_terminations(load_terminations(terminations_file))
    responses = clean_responses(load_responses(responses_file))

    merged = merge_data(terminations, responses)
    completed, incomplete = split_forms(merged)

    return {
        'completed': completed,
        'incomplete': incomplete,
        'summary': summary_by_manager(completed, incomplete),
        'kpis': compute_kpis(completed, incomplete),
    }



# terminationss = (r"C:\Users\User\Downloads\All Pending HR Requests-2026-10-17-22-00-40.xlsx")
# responsess = (r"C:\Users\User\JDA TSG\Human Resources - Documents\New HR Files\HR Data Analyst\Termination-form\Responses File\TTL Expert _ Lead End-of-Season Review(1-1255).xlsx")

# results = run_pipeline(terminationss, responsess)



# results['summary'].to_excel(r"C:\Users\User\Downloads\summary.xlsx", index=False)

# results['completed'].to_excel(r"C:\Users\User\Downloads\completed.xlsx", index=False)

# results['incomplete'].to_excel(r"C:\Users\User\Downloads\incomplete_completed.xlsx", index=False)