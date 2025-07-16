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
            "PROC_SQL": "Yes",
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

    # DATABASE CONNECTION ANALYSIS
    db_connections = detect_database_connections(code, filepath)
    rows.extend(db_connections)

    return rows

def detect_database_connections(code, filepath):
    """Detect database connections and potential missing connections"""
    rows = []
    
    # Track found connections
    found_connections = {
        'libname': [],
        'proc_sql_connect': [],
        'proc_datasets_engine': []
    }
    
    # 1. LIBNAME statements with database engines
    db_engines = ['oracle', 'teradata', 'mysql', 'postgres', 'sqlserver', 'db2', 'netezza', 'sybase', 'odbc', 'oledb']
    
    for match in re.finditer(r'libname\s+(\w+)\s+(\w+)(?:\s+.*?)?;', code, flags=re.IGNORECASE | re.DOTALL):
        libref = match.group(1)
        engine = match.group(2).lower()
        
        if engine in db_engines:
            found_connections['libname'].append(libref)
            rows.append({
                "statement": f"libname {libref} {engine}",
                "DB_CONNECTION": "Yes",
                "connection_type": f"LIBNAME_{engine.upper()}",
                "libref": libref,
                "engine": engine,
                "file_path": filepath
            })
    
    # 2. PROC SQL CONNECT statements
    for match in re.finditer(r'connect\s+to\s+(\w+)(?:\s+.*?)?;', code, flags=re.IGNORECASE | re.DOTALL):
        engine = match.group(1).lower()
        found_connections['proc_sql_connect'].append(engine)
        rows.append({
            "statement": f"connect to {engine}",
            "DB_CONNECTION": "Yes",
            "connection_type": f"PROC_SQL_CONNECT_{engine.upper()}",
            "engine": engine,
            "file_path": filepath
        })
    
    # 3. Look for database library references without connections
    db_table_refs = []
    
    # Find all library.table references
    for match in re.finditer(r'(?:from|join|data|set|merge|update|modify|into)\s+(\w+)\.(\w+)', code, flags=re.IGNORECASE):
        libref = match.group(1).lower()
        table = match.group(2)
        
        # Skip common SAS libraries
        if libref not in ['work', 'sashelp', 'sasuser', 'webwork']:
            db_table_refs.append((libref, table, match.group(0)))
    
    # 4. Check for missing connections
    for libref, table, statement in db_table_refs:
        # Check if this library has a connection defined
        has_connection = (
            libref in [lib.lower() for lib in found_connections['libname']] or
            any(conn in statement.lower() for conn in found_connections['proc_sql_connect'])
        )
        
        if not has_connection:
            rows.append({
                "statement": statement,
                "referenced_table": f"{libref}.{table}",
                "libref": libref,
                "MISSING_CONNECTION": "Yes",
                "connection_issue": "Library referenced but no connection found",
                "file_path": filepath
            })
    
    # 5. Look for PROC SQL pass-through without connections
    passthrough_patterns = [
        r'select\s+.*?\s+from\s+connection\s+to\s+(\w+)',
        r'execute\s+\(.*?\)\s+by\s+(\w+)',
        r'disconnect\s+from\s+(\w+)'
    ]
    
    for pattern in passthrough_patterns:
        for match in re.finditer(pattern, code, flags=re.IGNORECASE | re.DOTALL):
            connection_name = match.group(1).lower()
            
            # Check if this connection was established
            if connection_name not in found_connections['proc_sql_connect']:
                rows.append({
                    "statement": match.group(0),
                    "connection_name": connection_name,
                    "MISSING_CONNECTION": "Yes",
                    "connection_issue": "Pass-through query without connection",
                    "file_path": filepath
                })
    
    # 6. Look for database-specific syntax without connections
    db_syntax_patterns = [
        (r'bulk\s+insert', 'SQL Server bulk insert without connection'),
        (r'exec\s+sp_', 'SQL Server stored procedure without connection'),
        (r'exec\s+dbms_', 'Oracle DBMS package without connection'),
        (r'select\s+.*?\s+from\s+dual', 'Oracle DUAL table without connection')
    ]
    
    for pattern, description in db_syntax_patterns:
        for match in re.finditer(pattern, code, flags=re.IGNORECASE):
            rows.append({
                "statement": match.group(0),
                "MISSING_CONNECTION": "Yes",
                "connection_issue": description,
                "file_path": filepath
            })
    
    # 7. Summary of connections found
    if found_connections['libname'] or found_connections['proc_sql_connect']:
        rows.append({
            "statement": "Connection Summary",
            "DB_CONNECTION": "Yes",
            "connection_type": "SUMMARY",
            "libname_connections": ', '.join(found_connections['libname']),
            "sql_connections": ', '.join(found_connections['proc_sql_connect']),
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