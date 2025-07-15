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

    # DATA step WRITE-BACK
    for match in re.findall(r'data\s+([\w&]+)\.([\w&]+).*?;', code, flags=re.IGNORECASE):
        full = f"{match[0]}.{match[1]}"
        rows.append({
            "statement": f"data {full}",
            "output_table": full,
            "WRITE_BACK": "Yes",
            
            "file_path": filepath
        })

    # ✅ PROC SQL - always WRITE_BACK
    for sql_block in re.findall(r'proc\s+sql.*?quit;', code, flags=re.IGNORECASE | re.DOTALL):
        first_line = sql_block.strip().splitlines()[0] if sql_block.strip() else "proc sql;"
        output_table = None

        # Optional CREATE TABLE
        create_match = re.search(r'create\s+table\s+([\w&]+\.[\w&]+)', sql_block, flags=re.IGNORECASE)
        if create_match:
            output_table = create_match.group(1)

        rows.append({
            "statement": first_line,
            "output_table": output_table,
            "WRITE_BACK": "Yes",  # force write-back
            "file_path": filepath
        })

        # FROM / JOIN inputs
        for keyword, libref, table in re.findall(r'(from|join)\s+(\w+)\.(\w+)', sql_block, flags=re.IGNORECASE):
            rows.append({
                "statement": f"{keyword.lower()} {libref}.{table}",
                "Input tables": f"{libref}.{table}",
                "tables_sourcejoin": table,
                "file_path": filepath
            })

    # Other PROCs with OUT= or CREATE TABLE
    for proc_block in re.findall(r'(proc\s+([\w]+).*?)(run;|quit;)', code, flags=re.IGNORECASE | re.DOTALL):
        full_block, proc_name, _ = proc_block
        proc_name = proc_name.lower()

        if proc_name in {"import", "export"}:
            continue

        has_out = re.search(r'\bout\s*=\s*([\w&\.]+)', full_block, re.IGNORECASE)
        has_create = re.search(r'create\s+table\s+([\w&]+\.[\w&]+)', full_block, re.IGNORECASE)
        output_table = has_out.group(1) if has_out else (has_create.group(1) if has_create else None)

        if output_table:
            first_line = full_block.strip().split("\n")[0].strip()
            rows.append({
                "statement": first_line,
                "output_table": output_table,
                "WRITE_BACK": "Yes",

                "file_path": filepath
            })

    
    # PROC IMPORT
    for match in re.finditer(
        r'proc\s+import.*?(?:out\s*=\s*(\w+)).*?(?:datafile\s*=\s*["\'](.+?)["\'])|'
        r'proc\s+import.*?(?:datafile\s*=\s*["\'](.+?)["\']).*?(?:out\s*=\s*(\w+))',
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

   # PROC EXPORT
    for match in re.findall(r'proc\s+export\s+data\s*=\s*([\w&\.]+).*?outfile\s*=\s*["\'](.+?)["\']', code, flags=re.DOTALL | re.IGNORECASE):
        rows.append({
             "statement": f"proc export data={match[0]}",
             "output_table": match[1],
            "WRITE_BACK": "No",
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
