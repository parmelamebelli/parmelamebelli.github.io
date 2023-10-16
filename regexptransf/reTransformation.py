import regex as re
import openpyxl

def build_metrics_dict(workbook):
    sheet_metrics = workbook["Metrics"]
    metrics_dict = {}

    for row in sheet_metrics.iter_rows():
        if row[0].value is None or row[4].value is None:
            continue  # continue to the next row instead of breaking
        key = (row[0].value or '', row[2].value or '', row[3].value or '')  # convert None to empty strings
        metrics_dict[key] = row[4].value
    return metrics_dict

def get_metric_from_memory(metrics_dict, t_value, r_value, c_value):
    if r_value == ['']:
        metric = metrics_dict.get((t_value, '', c_value))
    elif c_value == ['']:
        metric = metrics_dict.get((t_value, r_value, ''))
    else:

        metric = metrics_dict.get((t_value, r_value, c_value))
    return metric

def get_c_values_from_structure_c(structure_c):
    if not structure_c or not isinstance(structure_c, str):
        return ['']
    match = re.search(r'scope\({t: .*?, c:(.*?), f: .*?, fv: .*?}\)', structure_c)
    return match.group(1).strip().split(";") if match else ['']

def get_r_values_from_structure_c(structure_c):
    if not structure_c or not isinstance(structure_c, str):
        return ['']
    match = re.search(r'scope\({t: .*?, r:(.*?), f: .*?, fv: .*?}\)', structure_c)
    return match.group(1).strip().split(";") if match else ['']

def compute_ps(t_value, r_value, c_value, input_str, metrics_dict):
    last_placeholder = ""
    metric_type = get_metric_from_memory(metrics_dict, t_value, r_value.strip(), c_value)
    suffix = ""  # default suffix
    if metric_type in ["Metric: Monetary", "Metric: Pure", "Metric: Decimal", "Metric: Integer"]:
        prefix = "VALUE"  # default prefix for these metric types
        if "not(isNull(" in input_str:
            prefix = "( EXISTS" if input_str.startswith("(") else "EXISTS"
            suffix = ") " if input_str.endswith(")))") else ""
        elif "isNull(" in input_str:
            prefix = "( NOT( EXISTS" if input_str.startswith("(") else "NOT( EXISTS"
            suffix = ") )" if input_str.endswith("))") else ") "
    else:
        prefix = "TEXT"  # default prefix for other metric types
        if "not(isNull(" in input_str:
            prefix = "( EXISTS_TEXT" if input_str.startswith("(") else "EXISTS_TEXT"
            suffix = ") " if input_str.endswith(")))") else ""
        elif "isNull(" in input_str:
            prefix = "( NOT ( EXISTS_TEXT" if input_str.startswith("(") else "NOT ( EXISTS_TEXT"
            suffix = ") )" if input_str.endswith("))") else ") "
        elif "not(matches" in input_str:
            prefix = "NOT ( LIKE( TEXT"
            suffix = ")"
        elif "matches" in input_str:
            prefix = "LIKE( TEXT"
    match = re.search(r'\[s2c_([A-Za-z0-9]+:[A-Za-z0-9]+)\]', input_str)

    if match:
        last_placeholder = ",' TXTMD'"
    elif not match and "TEXT" in prefix:
        last_placeholder = ",' ELEM'"
    else:
        ""
    return {
        't_value': t_value,
        'r_value': r_value,
        'prefix': prefix,
        'suffix': suffix,
        'last_placeholder': last_placeholder
    }

def convert_structure(input_str, c_values_from_structure_c, r_values_from_structure_c, workbook, metrics_dict):
    results = []
    starts_with_sum = input_str.lower().startswith('sum')
    match = re.search(r'({t: (.*?)(, r: (.*?))?(, c: ((?:[^,]+?)(?=(?:, |}))))?(, z: (.*?))?})', input_str)
    if not match:
        return input_str

    t_value = match.group(2)
    r_values = match.group(4).split(';') if match.group(4) else r_values_from_structure_c
    c_values = match.group(6).split(';') if match.group(6) else c_values_from_structure_c

    if t_value.startswith("SE"):
        segments = t_value.split(".")
        segments[0] = "S"
        if len(segments[2]) == 1:
            segments[2] = "0" + segments[2]
        if len(segments) > 3:
            if segments[3] == "16":
                segments[3] = "01"
            elif segments[3] == "17":
                segments[3] = "02"
        t_value = ".".join(segments)[:10]
    for r_value in r_values:
     if t_value == "S.01.01.01" and r_value and r_value.startswith("ER"):
        r_value = r_value.replace("ER", "R", 1)

    t_value = t_value[:10]

    if not starts_with_sum:
        r_value = r_values[0] if match.group(4) else r_values_from_structure_c
        c_value = c_values[0] if match.group(6) else c_values_from_structure_c
        process_structure = compute_ps(t_value, r_value, c_value, input_str, metrics_dict)

        return f"{process_structure['prefix']}( #REPPER, '#COMPANY', '{t_value}', '{c_value}', '{r_value}', '#FISCVARNT', '', '', '', '', '', ''{process_structure['last_placeholder']} ) {process_structure['suffix']}"

    if starts_with_sum and match.group(4):
        if len(r_values) > 1:
            c_value = c_values[0] if match.group(6) else c_values_from_structure_c
            results = []
            for r_value in r_values:
                process_structure = compute_ps(t_value, r_value.strip(), c_value.strip(), input_str, metrics_dict)
                structure_B = f"{process_structure['prefix']}( #REPPER, '#COMPANY', '{t_value}', '{c_value}', '{r_value}', '#FISCVARNT', '', '', '', '', '', ''{process_structure['last_placeholder']} ) {process_structure['suffix']}"
                results.append(structure_B)
            return ' + '.join(results)

        elif len(r_values) == 1:
            r_value = r_values[0]

        elif len(r_values) < 1:  # This handles the case when there's no r_value
            r_value = ''  # Or any default value
        c_value = c_values[0] if match.group(6) else c_values_from_structure_c
        t_value = t_value[:10]
        process_structure = compute_ps(t_value, r_value, c_value, input_str, metrics_dict)
        condition = "1 == 1"
        if t_value.startswith('S.06.02'):
            condition = f"TEXT( #REPPER, '#COMPANY', '{t_value}', 'SCRC0150', '', '#FISCVARNT', '', '', '', '', '', '', 'ELEM' ) <> 'AGG_ONLY'"
        return f"SUM( '{t_value}', {condition}, {process_structure['prefix']}( #REPPER, '#COMPANY', '{t_value}', '{c_value}', '{r_value}', '#FISCVARNT', '', '', '', '', '', ''{process_structure['last_placeholder']} ) {process_structure['suffix']} )"

    if starts_with_sum and match.group(6):
        if len(c_values) > 1:
            r_value = r_values[0] if match.group(4) else r_values_from_structure_c
            results = []

            for c_value in c_values:
                process_structure = compute_ps(t_value, r_value.strip(), c_value.strip(), input_str, metrics_dict)
                structure_B = f"{process_structure['prefix']}( #REPPER, '#COMPANY', '{t_value}', '{c_value}', '{r_value}', '#FISCVARNT', '', '', '', '', '', ''{process_structure['last_placeholder']} ) {process_structure['suffix']}"
                results.append(structure_B)
            return ' + '.join(results)

        elif len(c_values) == 1:
            c_value = c_values[0]
        elif len(c_values) < 1:  # This handles the case when there's no r_value
            c_value = ''  # Or any default value

        r_value = r_values[0] if match.group(4) else r_values_from_structure_c
        t_value = t_value[:10]
        process_structure = compute_ps(t_value, r_value, c_value, input_str, metrics_dict)
        condition = "1 == 1"

        if t_value.startswith('S.06.02'):
            condition = f"TEXT( #REPPER, '#COMPANY', '{t_value}', 'SCRC0150', '', '#FISCVARNT', '', '', '', '', '', '', 'ELEM' ) <> 'AGG_ONLY'"
        return f"SUM( '{t_value}', {condition}, {process_structure['prefix']}( #REPPER, '#COMPANY', '{t_value}', '{c_value}', '{r_value}', '#FISCVARNT', '', '', '', '', '', ''{process_structure['last_placeholder']} ) {process_structure['suffix']} )"

def add_brackets(expression):
    # List of operators to match
    operators = ['AND', 'OR', '<=', '>=', '=', '<', '>', '+', '-', '*', '/']

    # If the expression contains "EXISTS", return it as is

    if "EXISTS" in expression and len(re.findall(r'\b(AND|OR)\b', expression)) > 1:
        return expression

    # For each operator

    for operator in operators:

        if operator in expression:
            # Split the expression at the operator
            parts = expression.split(operator, 1)  # Split at the first occurrence of operator
            # Check if the part ends with '1' before and starts with '1' after '==' and handle it specially
            if (parts[0].strip().endswith('1') or parts[1].strip().startswith('1')) and operator == '=':
                return expression  # if it's '1 == 1', return it as is
            # Check if the part is a number and wrap it with brackets
            left = ' ( ' + parts[0].strip() + ' ) ' if not parts[0].strip().replace('.', '', 1).isdigit() else ' ( ' + parts[0].strip() + ' ) '

            right = ' ( ' + parts[1].strip() + ' ) ' if not parts[1].strip().replace('.', '', 1).isdigit() else ' ( ' + parts[1].strip() + ' ) '

            # Return the bracketed parts joined by the operator
            return left + operator + right
    # If no operator matched, return the original expression
    return expression

def convert_equation(equation, c_values_from_structure_c, r_values_from_structure_c, workbook, metrics_dict):
    # Adjusted regular expression for splitting

    parts = re.split(
        r' *(<=|>=|==|\+|-|=(?! \[s2c)|, (?! ?([a-zA-Z]))|and|or|(?<=imax\([^)]+)\)|(?<=imin\([^)]+)\)|(?<=imax)\(|(?<=imin)\(|(?<![a-zA-Z])\(|>|<|\*) *',
        equation)


    for i in range(len(parts)):
        if parts[i] is None:
            parts[i] = ''  # convert None to empty string
        match = re.search(r'(\s*(=|!=)\s*)\[s2c_([A-Za-z0-9]+:[A-Za-z0-9]+)\]', parts[i])

        if match:
            # Determine the operator based on the captured value
            operator = "=" if match.group(2) == "=" else "<>"
            parts[i] = convert_structure(parts[i], c_values_from_structure_c, r_values_from_structure_c, workbook,
                                         metrics_dict) + f" {operator} '" + match.group(3) + "'"

        elif '{t:' in parts[i]:
            parts[i] = convert_structure(parts[i], c_values_from_structure_c, r_values_from_structure_c, workbook,
                                         metrics_dict)
        elif parts[i].lower() in ['+', '-', '=', '*', 'and', 'or', '>', '<', '<=', '>=', 'imin', 'imax', '\(', '\)',
                                  ' , ']:

            parts[i] = ' ' + parts[i].upper() + ' '
        elif parts[i].replace('.', '', 1).isdigit() or (
                parts[i].startswith('-') and parts[i][1:].replace('.', '', 1).isdigit()):
            if float(parts[i]) == 0:
                parts[i] = "0.00"

            # Otherwise, leave parts[i] unchanged.

    converted_equation = ''.join([str(x) if x is not None else '' for x in parts]).strip().replace('IMIN ',
                                                                                                   'MIN').replace(
        'IMAX ', 'MAX')
    if "EXISTS" not in converted_equation and len(re.findall(r'\b(AND|OR)\b', converted_equation)) > 1 and not match:
        segments = re.split(r'(\bAND\b|\bOR\b)', converted_equation)
        new_segments = []
        for i in range(0, len(segments) - 1, 2):
            new_segments.append('(' + segments[i].strip() + ')')
            new_segments.append(segments[i + 1].strip())
        new_segments.append('(' + segments[-1].strip() + ')')
        converted_equation = ' '.join(new_segments)
        return converted_equation
    elif match:
        return converted_equation
    else:
        return add_brackets(converted_equation)
    return add_brackets(converted_equation)

# Load the workbook and worksheet
file_path1 = f"C:/Users/User/filepath"
workbook = openpyxl.load_workbook(file_path1)
sheet = workbook["Sheet1"]
metrics_dict = build_metrics_dict(workbook)

row = 1

while sheet.cell(row=row, column=3).value:

    cell_a_value = sheet.cell(row=row, column=1).value
    cell_b_value = sheet.cell(row=row, column=2).value
    cell_c_value = sheet.cell(row=row, column=3).value
    c_values_from_structure_c = get_c_values_from_structure_c(cell_a_value)
    r_values_from_structure_c = get_r_values_from_structure_c(cell_a_value)
    # Store the converted values temporarily
    converted_values = []

    for c_value in c_values_from_structure_c:

        for r_value in r_values_from_structure_c:
            converted_value = convert_equation(cell_c_value, c_value, r_value, workbook, metrics_dict)
            converted_values.append(converted_value)

    # Add the converted values to the worksheet and insert necessary rows

    for idx_offset, converted_value in enumerate(converted_values):
        sheet.cell(row=row + idx_offset, column=5).value = converted_value
        if not cell_b_value:
            sheet.cell(row=row + idx_offset, column=4).value = ""
        else:
            sheet.cell(row=row + idx_offset, column=4).value = convert_equation(cell_b_value, c_value, r_value,
                                                                                workbook, metrics_dict)
        if idx_offset < len(converted_values) - 1:
            sheet.insert_rows(row + idx_offset + 1)
    # Adjust the row pointer
    row += len(converted_values)
# Save the modified workbook
workbook.save(file_path1)