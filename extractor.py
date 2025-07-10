import os
import re
import pandas as pd

declared_libs = set()  # Will store all libnames defined

def extract_from_file(file_path):
    results = []

    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
        for line in lines:
            clean_line = line.strip().lower()
            if clean_line == "" or clean_line.startswith("*"):
                continue

            row = {
                "statement": clean_line,
                "file_path": file_path,
                "INCLUDE": "",
                "INCLUDE_PATH": "",
                "DEPENDENCY_EXISTS": "",
                "LIBNAME": "",
                "MACRO_DEF": "",
                "MACRO_CALL": "",
                "tables_sourcejoin": "",
                "tables_table": "",
                "source": "",
                "input_table": "",
                "output_table": "",
                "WRITE_BACK": "No",
                "MISSING_LIBNAME": ""
            }

            # %include "file.sas";
            include_path = re.search(r'%include\s+["\'](.+?)["\']', clean_line)
            if include_path:
                row["INCLUDE"] = "yes"
                include_file = include_path.group(1)
                row["INCLUDE_PATH"] = include_file
                full_include_path = os.path.join("SAS Files", os.path.basename(include_file))
                row["DEPENDENCY_EXISTS"] = "Yes" if os.path.isfile(full_include_path) else "No"

            # libname libref 'path';
            lib = re.search(r'libname\s+(\w+)\s+["\']([^"\']+)["\']', clean_line)
            if lib:
                libref = lib.group(1)
                declared_libs.add(libref.lower())
                row["LIBNAME"] = libref
                row["source"] = lib.group(2)

            # %macro macroname(...)
            macro_def = re.search(r'%macro\s+(\w+)', clean_line)
            if macro_def:
                row["MACRO_DEF"] = macro_def.group(1)

            # %macro_call(...)
            macro_call = re.search(r'%(\w+)\((.*?)\)', clean_line)
            if macro_call:
                row["MACRO_CALL"] = macro_call.group(1)

            # data output.table;
            data_ = re.search(r'data\s+(\w+)\.(\w+)', clean_line)
            if data_:
                libref = data_.group(1)
                tablename = data_.group(2)
                row["output_table"] = f"{libref}.{tablename}"
                row["tables_table"] = tablename

                # Write-back check (lib ≠ work)
                if libref != "work":
                    row["WRITE_BACK"] = "Yes"

                # Check for missing libname
                if libref not in declared_libs:
                    row["MISSING_LIBNAME"] = "Yes"
                else:
                    row["MISSING_LIBNAME"] = "No"

            # set lib.table;
            set_ = re.search(r'set\s+(\w+)\.(\w+)', clean_line)
            if set_:
                libref = set_.group(1)
                tablename = set_.group(2)
                row["input_table"] = f"{libref}.{tablename}"
                row["tables_sourcejoin"] = tablename
                if libref not in declared_libs:
                    row["MISSING_LIBNAME"] = "Yes"
                else:
                    row["MISSING_LIBNAME"] = "No"

            # merge table1 table2;
            merge_ = re.search(r'merge\s+(.+);', clean_line)
            if merge_:
                merged_tables = merge_.group(1).split()
                row["tables_sourcejoin"] = ", ".join(merged_tables)

            # proc sql create table lib.table as ...
            sql_ = re.search(r'create\s+table\s+(\w+)\.(\w+)', clean_line)
            if sql_:
                libref = sql_.group(1)
                tablename = sql_.group(2)
                row["output_table"] = f"{libref}.{tablename}"
                row["tables_table"] = tablename

                if libref != "work":
                    row["WRITE_BACK"] = "Yes"

                if libref not in declared_libs:
                    row["MISSING_LIBNAME"] = "Yes"
                else:
                    row["MISSING_LIBNAME"] = "No"

            results.append(row)

    return results


def main():
    base_dir = "SAS Files"
    all_results = []

    for filename in os.listdir(base_dir):
        if filename.endswith(".sas"):
            file_path = os.path.join(base_dir, filename)
            print(f"📄 Processing: {filename}")
            extracted = extract_from_file(file_path)
            all_results.extend(extracted)

    df = pd.DataFrame(all_results)
    df.to_csv("pgm_analysis.csv", index=False)
    print(f"\n✅ Done! Extracted {len(all_results)} rows into 'pgm_analysis.csv'")


if __name__ == "__main__":
    main()

