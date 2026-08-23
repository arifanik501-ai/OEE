import json

days_of_week = ['Sat', 'Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri'] # August 1, 2026 is Saturday
machines = ['Rotor (CNC)', 'Bottom Cover (CNC)', 'Top Cover (CNC)']

rows = []

for r in range(6, 99):
    row_idx = r - 6
    day_num = (row_idx // 3) + 1  # 1 to 31
    m_idx = row_idx % 3           # 0, 1, 2
    
    date_str = f"{day_num}-Aug-26"
    day_str = days_of_week[(day_num - 1) % 7]
    
    row_dict = {
        'row': r,
        'A': {'val': date_str, 'formula': '=AM3' if r == 6 else (f'=A{r-3}+1' if m_idx == 0 else f'=A{r-1}')},
        'B': {'val': day_str, 'formula': f'=A{r}'},
        'C': {'val': 'Morning', 'formula': None},
        'D': {'val': machines[m_idx], 'formula': None},
        'E': {'val': None, 'formula': None},
        'F': {'val': None, 'formula': None},
        'G': {'val': None, 'formula': None},
        'H': {'val': None, 'formula': None},
        'I': {'val': 0, 'formula': f'=IF(E{r}<>"",30,0)'},
        'J': {'val': '-', 'formula': f'=H{r}-AH{r}'}
    }
    
    cols_dt = ['K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']
    for c in cols_dt:
        row_dict[c] = {'val': None, 'formula': None}
        
    row_dict['AH'] = {'val': 0, 'formula': f'=SUM(K{r}:AG{r})'}
    row_dict['AI'] = {'val': 0, 'formula': f'=IFERROR(J{r}/H{r},"0")'}
    row_dict['AJ'] = {'val': 0, 'formula': f'=IFERROR(F{r}/E{r},"0")'}
    row_dict['AK'] = {'val': 0, 'formula': f'=IFERROR(F{r}/(F{r}+G{r}),"0")'}
    row_dict['AL'] = {'val': 0, 'formula': f'=IFERROR(AK{r}*AJ{r}*AI{r},"0")'}
    row_dict['AM'] = {'val': '', 'formula': None}
    
    rows.append(row_dict)

with open('initial_sheet_data.js', 'w', encoding='utf-8') as f:
    f.write('const INITIAL_EXCEL_ROWS = ' + json.dumps(rows, ensure_ascii=False, indent=2) + ';')

print(f"Generated clean dataset for August with {len(rows)} rows!")
