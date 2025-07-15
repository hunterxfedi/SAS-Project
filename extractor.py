import os
import re
import pandas as pd

CONTROL_KEYWORDS = {
    "if", "then", "else", "do", "end", "put", "goto", "abort", "return",
    "symdel", "until", "while", "scan", "substr", "eval", "upcase", "lowcase",
    "index", "find", "length", "sysfunc", "sysget"
}

def read_sas_file(filepath):
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()
    content = re.sub(r'/\*.*?\*/', '', content, flags=re.DOTALL)
    content = re.sub(r'^\s*\*.*?;', '', content, flags=re.MULTILINE)
    return content

def extract_all_blocks(code, filepath):
    rows = []

    # %INCLUDE
    for inc in re.findall(r'%include\s+["\'](.+?)["\']\s*;', code, flags=re.IGNORECASE):
        rows.append({
            "statement": f"%include \"{inc}\";",
            "INCLUDE_PATH": inc,
            "DEPENDENCY_EXISTS": "Yes" if os.path.isfile(os.path.join("SAS Files", os.path.basename(inc))) else "No",
            "file_path": filepath
        })

    # %LET
    for var, val in re.findall(r'%let\s+(\w+)\s*=\s*([^;]+);', code, flags=re.IGNORECASE):
        rows.append({
            "statement": f"%let {var}={val};",
            "LET_STATEMENT": var,
            "file_path": filepath
        })

    # %MACRO CALLS
    for name, args in re.findall(r'%(\w+)\s*\((.*?)\)\s*;', code):
        if name.lower() not in CONTROL_KEYWORDS:
            rows.append({
                "statement": f"%{name}({args});",
                "MACRO_CALL": name,
                "file_path": filepath
            })

    # MERGE
    for block in re.findall(r'data\s+.*?run\s*;', code, flags=re.DOTALL | re.IGNORECASE):
        match = re.search(r'merge\s+(.*?);', block, flags=re.IGNORECASE)
        if match:
            merge_line = match.group(1)
            raw_tables = re.findall(r'([a-zA-Z_][\w.&]*)\s*(?=\(|\s|$)', merge_line)
            ignore_keywords = {"by", "in", "keep", "rename", "=", "then", "else", "do", "and", "or", "where", "into", "the", "to"}
            cleaned = [t for t in raw_tables if t.lower() not in ignore_keywords]
            rows.append({
                "statement": f"merge {merge_line};",
                "tables_sourcejoin": ", ".join(cleaned),
                "file_path": filepath
            })

    # IMPROVED DATA step WRITE-BACK
    for match in re.finditer(r'data\s+((?:[\w&]+\.)?[\w&]+)(?:\s*\(.*?\))?\s*;', code, flags=re.IGNORECASE):
        dataset_name = match.group(1).strip()
        
        # Skip DATA _NULL_ as it doesn't create datasets
        if dataset_name.lower() != "_null_":
            rows.append({
                "statement": f"data {dataset_name};",
                "output_table": dataset_name,
                "WRITE_BACK": "Yes",
                "write_back_type": "DATA_STEP",
                "file_path": filepath
            })

    # IMPROVED PROC SQL
    for sql_block in re.findall(r'proc\s+sql.*?quit;', code, flags=re.IGNORECASE | re.DOTALL):
        # First line for SQL procedure detection
        first_line = sql_block.strip().splitlines()[0] if sql_block.strip() else "proc sql;"
        rows.append({
            "statement": first_line,
            "WRITE_BACK": "Yes",
            "file_path": filepath
        })
        
        # Look for CREATE TABLE statements
        for match in re.finditer(r'create\s+table\s+((?:[\w&]+\.)?[\w&]+)', sql_block, flags=re.IGNORECASE):
            table_name = match.group(1)
            rows.append({
                "statement": f"create table {table_name}",
                "output_table": table_name,
                "WRITE_BACK": "Yes",
                "write_back_type": "PROC_SQL_CREATE",
                "file_path": filepath
            })
        
        # Look for INSERT INTO statements
        for match in re.finditer(r'insert\s+into\s+((?:[\w&]+\.)?[\w&]+)', sql_block, flags=re.IGNORECASE):
            table_name = match.group(1)
            rows.append({
                "statement": f"insert into {table_name}",
                "output_table": table_name,
                "WRITE_BACK": "Yes",
                "write_back_type": "PROC_SQL_INSERT",
                "file_path": filepath
            })

        # FROM / JOIN inputs (keep existing - these are NOT write-backs)
        for keyword, libref, table in re.findall(r'(from|join)\s+(\w+)\.(\w+)', sql_block, flags=re.IGNORECASE):
            rows.append({
                "statement": f"{keyword.lower()} {libref}.{table}",
                "Input tables": f"{libref}.{table}",
                "tables_sourcejoin": table,
                "file_path": filepath
            })

    # PROC SORT with OUT= (NEW)
    for match in re.finditer(r'proc\s+sort\s+data\s*=\s*([\w&\.]+).*?out\s*=\s*([\w&\.]+)', code, flags=re.IGNORECASE | re.DOTALL):
        output_table = match.group(2)
        rows.append({
            "statement": f"proc sort out={output_table}",
            "output_table": output_table,
            "WRITE_BACK": "Yes",
            "write_back_type": "PROC_SORT",
            "file_path": filepath
        })

    # PROC MEANS/SUMMARY with OUT= (IMPROVED)
    for match in re.finditer(r'proc\s+(means|summary)\s+.*?out\s*=\s*([\w&\.]+)', code, flags=re.IGNORECASE | re.DOTALL):
        proc_name = match.group(1)
        output_table = match.group(2)
        rows.append({
            "statement": f"proc {proc_name} out={output_table}",
            "output_table": output_table,
            "WRITE_BACK": "Yes",
            "write_back_type": f"PROC_{proc_name.upper()}",
            "file_path": filepath
        })

    # PROC FREQ with OUT= (NEW)
    for match in re.finditer(r'proc\s+freq\s+.*?out\s*=\s*([\w&\.]+)', code, flags=re.IGNORECASE | re.DOTALL):
        output_table = match.group(1)
        rows.append({
            "statement": f"proc freq out={output_table}",
            "output_table": output_table,
            "WRITE_BACK": "Yes",
            "write_back_type": "PROC_FREQ",
            "file_path": filepath
        })

    # PROC TRANSPOSE with OUT= (NEW)
    for match in re.finditer(r'proc\s+transpose\s+.*?out\s*=\s*([\w&\.]+)', code, flags=re.IGNORECASE | re.DOTALL):
        output_table = match.group(1)
        rows.append({
            "statement": f"proc transpose out={output_table}",
            "output_table": output_table,
            "WRITE_BACK": "Yes",
            "write_back_type": "PROC_TRANSPOSE",
            "file_path": filepath
        })

    # PROC APPEND (NEW)
    for match in re.finditer(r'proc\s+append\s+.*?base\s*=\s*([\w&\.]+)', code, flags=re.IGNORECASE | re.DOTALL):
        base_table = match.group(1)
        rows.append({
            "statement": f"proc append base={base_table}",
            "output_table": base_table,
            "WRITE_BACK": "Yes",
            "write_back_type": "PROC_APPEND",
            "file_path": filepath
        })

    # PROC DATASETS MODIFY (NEW)
    for dataset_block in re.findall(r'proc\s+datasets.*?quit;', code, flags=re.IGNORECASE | re.DOTALL):
        for match in re.finditer(r'modify\s+([\w&\.]+)', dataset_block, flags=re.IGNORECASE):
            table_name = match.group(1)
            rows.append({
                "statement": f"proc datasets modify {table_name}",
                "output_table": table_name,
                "WRITE_BACK": "Yes",
                "write_back_type": "PROC_DATASETS_MODIFY",
                "file_path": filepath
            })

    # Generic PROC with OUT= (IMPROVED - catch other statistical procedures)
    procs_with_out = ["univariate", "corr", "reg", "logistic", "glm", "mixed", "genmod", 
                      "ttest", "npar1way", "anova", "glimmix", "lifereg", "phreg", 
                      "surveyfreq", "surveymeans", "surveylogistic"]
    
    for proc_name in procs_with_out:
        pattern = rf'proc\s+{proc_name}\s+.*?out\s*=\s*([\w&\.]+)'
        for match in re.finditer(pattern, code, flags=re.IGNORECASE | re.DOTALL):
            output_table = match.group(1)
            rows.append({
                "statement": f"proc {proc_name} out={output_table}",
                "output_table": output_table,
                "WRITE_BACK": "Yes",
                "write_back_type": f"PROC_{proc_name.upper()}",
                "file_path": filepath
            })

    # PROC IMPORT (keep existing - NOT marked as write-back as requested)
    for match in re.finditer(
        r'proc\s+import.*?(?:out\s*=\s*([\w&\.]+)).*?(?:datafile\s*=\s*["\'](.+?)["\'])|'
        r'proc\s+import.*?(?:datafile\s*=\s*["\'](.+?)["\']).*?(?:out\s*=\s*([\w&\.]+))',
        code,
        flags=re.IGNORECASE | re.DOTALL
    ):
        out_table = match.group(1) or match.group(4)
        infile = match.group(2) or match.group(3)
        rows.append({
            "statement": f"proc import out={out_table}",
            "Input tables": infile,
            "output_table": out_table,
            "import proc": "Yes",
            "file_path": filepath
        })

    # PROC EXPORT (keep existing - NOT marked as write-back as requested)
    for match in re.findall(r'proc\s+export\s+data\s*=\s*([\w&\.]+).*?outfile\s*=\s*["\'](.+?)["\']', code, flags=re.DOTALL | re.IGNORECASE):
        rows.append({
            "statement": f"proc export data={match[0]}",
            "output_table": match[1],
            "export proc": "Yes",
            "file_path": filepath
        })

    return rows

def main():
    base_dir = "SAS Files"
    all_results = []

    if not os.path.isdir(base_dir):
        print("❌ 'SAS Files' folder not found.")
        return

    for filename in os.listdir(base_dir):
        if filename.endswith(".sas"):
            path = os.path.join(base_dir, filename)
            print(f"📄 Processing: {filename}")
            code = read_sas_file(path)
            all_results.extend(extract_all_blocks(code, path))

    df = pd.DataFrame(all_results)
    
    df.to_excel("final_analysis.xlsx", index=False)
    print(f"\n✅ Done! Extracted {len(all_results)} rows into 'final_analysis.xlsx'")

if __name__ == "__main__":
    main()