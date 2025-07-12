import os
import re
import pandas as pd

declared_libs = set()

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
    return content.lower()

def extract_all_blocks(code, filepath):
    rows = []

    # %INCLUDE
    for inc in re.findall(r'%include\s+["\'](.+?)["\']\s*;', code):
        rows.append({
            "statement": f"%include \"{inc}\";",
            "INCLUDE_PATH": inc,
            "DEPENDENCY_EXISTS": "Yes" if os.path.isfile(os.path.join("SAS Files", os.path.basename(inc))) else "No",
            "file_path": filepath
        })

    # %LET
    for var, val in re.findall(r'%let\s+(\w+)\s*=\s*([^;]+);', code):
        rows.append({
            "statement": f"%let {var}={val};",
            "LET_STATEMENT": var,
            "file_path": filepath
        })

    # %MACRO CALLS
    for name, args in re.findall(r'%(\w+)\s*\((.*?)\)\s*;', code):
        if name not in CONTROL_KEYWORDS:
            rows.append({
                "statement": f"%{name}({args});",
                "MACRO_CALL": name,
                "file_path": filepath
            })

    # LIBNAME
    for libref, path in re.findall(r'libname\s+(\w+)\s+["\']([^"\']+)["\'];', code):
        declared_libs.add(libref)

    # Input Tables: SET, FROM, JOIN
    for m in re.findall(r'set\s+(\w+)\.(\w+)', code):
        rows.append({
            "statement": f"set {m[0]}.{m[1]}",
            "Input tables": f"{m[0]}.{m[1]}",
            "tables_sourcejoin": m[1],
            "file_path": filepath
        })

    for m in re.findall(r'(?:from|join)\s+(\w+)\.(\w+)', code):
        rows.append({
            "statement": f"read {m[0]}.{m[1]}",
            "Input tables": f"{m[0]}.{m[1]}",
            "tables_sourcejoin": m[1],
            "file_path": filepath
        })

    # Output Tables: DATA or CREATE TABLE
    for m in re.findall(r'(data|create\s+table)\s+(\w+)\.(\w+)', code):
        rows.append({
            "statement": f"{m[1]}.{m[2]}",
            "output_table": f"{m[1]}.{m[2]}",
            "WRITE_BACK": "Yes",
            "file_path": filepath
        })

    # MERGE
    for block in re.findall(r'data\s+.*?run\s*;', code, flags=re.DOTALL | re.IGNORECASE):
        match = re.search(r'merge\s+(.*?);', block, flags=re.IGNORECASE)
        if match:
            merge_line = match.group(1)
            raw_tables = re.findall(r'([a-zA-Z_][\w.]*)\s*(?=\(|\s|$)', merge_line)
            ignore_keywords = {"by", "in", "keep", "rename", "=", "then", "else", "do", "and", "or", "where", "into", "the", "to"}
            cleaned = [t for t in raw_tables if t.lower() not in ignore_keywords]
            rows.append({
                "statement": f"merge {merge_line};",
                "tables_sourcejoin": ", ".join(cleaned),
                "file_path": filepath
            })

    # PROC SQL: CREATE TABLE
    for sql_block in re.findall(r'proc\s+sql.*?quit;', code, flags=re.DOTALL):
        if "create table" in sql_block:
            match = re.search(r'create\s+table\s+(\w+\.\w+)', sql_block)
            if match:
                out_table = match.group(1)
                rows.append({
                    "statement": "proc sql ...",
                    "output_table": out_table,
                    "WRITE_BACK": "Yes",
                    "file_path": filepath
                })

    # PROC SORT/TRANSPOSE/MEANS/SUMMARY with OUT=
    for proc in ["sort", "transpose", "means", "summary"]:
        pattern = rf'proc\s+{proc}\s+.*?out\s*=\s*(\w+\.\w+|\w+)'
        for out in re.findall(pattern, code, flags=re.DOTALL):
            rows.append({
                "statement": f"proc {proc} out={out}",
                "output_table": out,
                "WRITE_BACK": "Yes",
                "file_path": filepath
            })

    # PROC IMPORT (external input = write-back to dataset)
    for match in re.findall(r'proc\s+import\s+datafile\s*=\s*["\'](.+?)["\'].*?out\s*=\s*(\w+)', code, flags=re.DOTALL):
        rows.append({
            "statement": f"proc import out={match[1]}",
            "Input tables": match[0],
            "import proc": "Yes",
            "WRITE_BACK": "Yes",
            "file_path": filepath
        })

    # PROC EXPORT (write-back to file)
    for match in re.findall(r'proc\s+export\s+data\s*=\s*(\w+).*?outfile\s*=\s*["\'](.+?)["\']', code, flags=re.DOTALL):
        rows.append({
            "statement": f"proc export data={match[0]}",
            "output_table": match[1],
            "export proc": "Yes",
            "WRITE_BACK": "Yes",
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
