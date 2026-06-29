import pandas as pd
import json
import os
import re
from datetime import datetime

def rollup_summary_data(raw_tasks):
    """
    Sorts tasks by WBS numeric order and rolls up dates and progress from children to summary tasks.
    """
    def wbs_sort_key(wbs_str):
        try:
            return [int(x) for x in wbs_str.split('.')]
        except ValueError:
            return [999]

    raw_tasks.sort(key=lambda t: wbs_sort_key(t["WBS"]))

    # Process project summaries in bottom-up order (deepest WBS levels first)
    project_tasks = [t for t in raw_tasks if t["Type"] == "project"]
    project_tasks.sort(key=lambda t: len(t["WBS"].split('.')), reverse=True)

    for project in project_tasks:
        wbs_prefix = project["WBS"] + "."
        children = [t for t in raw_tasks if t["WBS"] != project["WBS"] and t["WBS"].startswith(wbs_prefix)]
        
        if not children:
            continue
            
        min_start = None
        max_end = None
        total_progress = 0
        valid_children_count = 0
        
        for child in children:
            if child["StartDate"]:
                try:
                    c_start = datetime.strptime(child["StartDate"], "%Y-%m-%d")
                    if min_start is None or c_start < min_start:
                        min_start = c_start
                except ValueError:
                    pass
            if child["EndDate"]:
                try:
                    c_end = datetime.strptime(child["EndDate"], "%Y-%m-%d")
                    if max_end is None or c_end > max_end:
                        max_end = c_end
                except ValueError:
                    pass
            
            total_progress += child["Progress"]
            valid_children_count += 1
            
        # Update project rollup dates
        if min_start:
            project["StartDate"] = min_start.strftime("%Y-%m-%d")
        if max_end:
            project["EndDate"] = max_end.strftime("%Y-%m-%d")
        
        # Update progress average
        if valid_children_count > 0:
            project["Progress"] = int(total_progress / valid_children_count)

    return raw_tasks

def parse_standard_excel(df):
    """
    Parses a standard Excel sheet with predefined columns.
    """
    df.columns = [str(c).strip() for c in df.columns]
    
    # Map common column name variations
    col_mapping = {}
    for col in df.columns:
        c_lower = col.lower()
        if c_lower == 'wbs':
            col_mapping['WBS'] = col
        elif c_lower in ['taskname', 'task name', '작업명']:
            col_mapping['TaskName'] = col
        elif c_lower in ['startdate', 'start date', '시작일']:
            col_mapping['StartDate'] = col
        elif c_lower in ['enddate', 'end date', '종료일']:
            col_mapping['EndDate'] = col
        elif c_lower in ['progress', '진행률', '진행율']:
            col_mapping['Progress'] = col
        elif c_lower in ['predecessors', 'predecessor', '선행작업']:
            col_mapping['Predecessors'] = col
        elif c_lower in ['assignees', 'assignee', '담당자', '주담당자', '주 담당자', 'primaryassignee', 'primary assignee']:
            col_mapping['PrimaryAssignee'] = col
        elif c_lower in ['부담당자', '부 담당자', '부서', 'secondaryassignee', 'secondary assignee']:
            col_mapping['SecondaryAssignee'] = col
        elif c_lower in ['type', 'tasktype', '구분']:
            col_mapping['Type'] = col
        elif c_lower in ['fixing', 'constraint', '제약조건', '제약']:
            col_mapping['Fixing'] = col
        elif c_lower in ['actualstartdate', 'actual start date', '실적시작일', '실적시작']:
            col_mapping['ActualStartDate'] = col
        elif c_lower in ['actualenddate', 'actual end date', '실적종료일', '실적종료']:
            col_mapping['ActualEndDate'] = col
        elif c_lower in ['schedulingmode', 'scheduling mode', '일정방식', '일정 방식']:
            col_mapping['SchedulingMode'] = col
        elif c_lower in ['load', '부하율', '부하율(%)', '부하율 (%)']:
            col_mapping['Load'] = col
        elif c_lower in ['mh', '공수', '공수(mh)', '공수 (mh)']:
            col_mapping['MH'] = col

    # Identify custom columns
    standard_cols = ['wbs', 'taskname', '작업명', 'startdate', '시작일', 'enddate', '종료일', 'progress', '진행률', '진행율', 'predecessors', '선행작업', 'assignees', '담당자', 'type', '구분', 'fixing', '제약조건', '제약', 'duration', '기간', 'primaryassignee', '주담당자', 'secondaryassignee', '부담당자', 'actualstartdate', '실적시작일', 'actualenddate', '실적종료일', 'schedulingmode', '일정방식', '주 담당자', '부 담당자', '실적 시작일', '실적 종료일', '일정 방식', 'load', '부하율', '부하율(%)', '부하율 (%)', 'mh', '공수', '공수(mh)', '공수 (mh)']
    custom_cols = []
    for col in df.columns:
        c_clean = str(col).lower().strip().replace(' ', '')
        if c_clean not in standard_cols:
            custom_cols.append(col)

    raw_tasks = []
    for index, row in df.iterrows():
        task = {
            "WBS": str(row.get(col_mapping.get('WBS'), '')).strip(),
            "TaskName": str(row.get(col_mapping.get('TaskName'), f"태스크 {index+1}")).strip(),
            "StartDate": "",
            "EndDate": "",
            "ActualStartDate": "",
            "ActualEndDate": "",
            "Progress": 0,
            "Predecessors": "",
            "PrimaryAssignee": "",
            "SecondaryAssignee": "",
            "Type": "task",
            "Fixing": "None",
            "SchedulingMode": "manual"
        }
        
        # Clean WBS format (e.g. 1.0 -> 1)
        if task["WBS"].endswith(".0"):
            task["WBS"] = task["WBS"][:-2]
            
        # Parse Dates safely
        start_val = row.get(col_mapping.get('StartDate'), None)
        end_val = row.get(col_mapping.get('EndDate'), None)
        
        if pd.notna(start_val):
            if isinstance(start_val, (datetime, pd.Timestamp)):
                task["StartDate"] = start_val.strftime("%Y-%m-%d")
            else:
                task["StartDate"] = str(start_val).split()[0].strip()
                
        if pd.notna(end_val):
            if isinstance(end_val, (datetime, pd.Timestamp)):
                task["EndDate"] = end_val.strftime("%Y-%m-%d")
            else:
                task["EndDate"] = str(end_val).split()[0].strip()

        # Parse Actual Dates
        act_start = row.get(col_mapping.get('ActualStartDate'), None)
        act_end = row.get(col_mapping.get('ActualEndDate'), None)
        if pd.notna(act_start):
            if isinstance(act_start, (datetime, pd.Timestamp)):
                task["ActualStartDate"] = act_start.strftime("%Y-%m-%d")
            else:
                task["ActualStartDate"] = str(act_start).split()[0].strip()
        if pd.notna(act_end):
            if isinstance(act_end, (datetime, pd.Timestamp)):
                task["ActualEndDate"] = act_end.strftime("%Y-%m-%d")
            else:
                task["ActualEndDate"] = str(act_end).split()[0].strip()

        # Parse Progress
        progress_val = row.get(col_mapping.get('Progress'), 0)
        if pd.notna(progress_val):
            try:
                task["Progress"] = int(float(progress_val))
                task["Progress"] = max(0, min(100, task["Progress"]))
            except ValueError:
                task["Progress"] = 0

        # Parse Predecessors
        pred_val = row.get(col_mapping.get('Predecessors'), '')
        if pd.notna(pred_val) and str(pred_val).strip() not in ['nan', '-']:
            task["Predecessors"] = str(pred_val).strip()

        # Parse Primary & Secondary Assignees
        prim_val = row.get(col_mapping.get('PrimaryAssignee'), '')
        if pd.notna(prim_val) and str(prim_val).strip() not in ['nan', '-']:
            task["PrimaryAssignee"] = str(prim_val).strip()
            
        sec_val = row.get(col_mapping.get('SecondaryAssignee'), '')
        if pd.notna(sec_val) and str(sec_val).strip() not in ['nan', '-']:
            task["SecondaryAssignee"] = str(sec_val).strip()

        # Parse Type
        type_val = str(row.get(col_mapping.get('Type'), 'task')).lower().strip()
        if type_val in ['project', 'summary', '프로젝트', '요약']:
            task["Type"] = "project"
        elif type_val in ['milestone', '마일스톤']:
            task["Type"] = "milestone"
        else:
            task["Type"] = "task"
            
        # Parse Fixing (Constraint)
        fixing_val = str(row.get(col_mapping.get('Fixing'), 'None')).strip()
        if fixing_val in ['None', 'Fix Duration', 'Fix StartDate', 'Fix EndDate']:
            task["Fixing"] = fixing_val
        elif fixing_val in ['기간 고정', '기간고정']:
            task["Fixing"] = "Fix Duration"
        elif fixing_val in ['시작일 고정', '시작일고정']:
            task["Fixing"] = "Fix StartDate"
        elif fixing_val in ['종료일 고정', '종료일고정']:
            task["Fixing"] = "Fix EndDate"
        else:
            task["Fixing"] = "None"
            
        # Parse Scheduling Mode
        mode_val = str(row.get(col_mapping.get('SchedulingMode'), 'auto')).lower().strip()
        if mode_val in ['manual', '수동']:
            task["SchedulingMode"] = "manual"
        else:
            task["SchedulingMode"] = "auto"

        # Parse Load
        load_val = row.get(col_mapping.get('Load'), None) if 'Load' in col_mapping else None
        if pd.notna(load_val):
            try:
                task["Load"] = int(float(load_val))
            except ValueError:
                task["Load"] = 100 if task["PrimaryAssignee"] else 0
        else:
            task["Load"] = 100 if task["PrimaryAssignee"] else 0

        # Compute MH
        mh_val = row.get(col_mapping.get('MH'), None) if 'MH' in col_mapping else None
        if pd.notna(mh_val):
            try:
                task["MH"] = round(float(mh_val), 1)
            except ValueError:
                task["MH"] = 0.0
        else:
            if task["StartDate"] and task["EndDate"]:
                try:
                    s = datetime.strptime(task["StartDate"], "%Y-%m-%d")
                    e = datetime.strptime(task["EndDate"], "%Y-%m-%d")
                    dur = (e - s).days + 1
                    task["MH"] = round(dur * 8 * (task["Load"] / 100), 1)
                except ValueError:
                    task["MH"] = 0.0
            else:
                task["MH"] = 0.0
            
        # Extract custom columns
        for col in custom_cols:
            col_key = f"col_{str(col).replace(' ', '_').replace('/', '_')}"
            task[col_key] = str(row.get(col, '')).strip() if pd.notna(row.get(col)) else ''

        raw_tasks.append(task)
        
    return raw_tasks

def parse_regas_excel_data(excel_path):
    """
    Parses the custom shipyard month-bar format of the Regas module Excel.
    """
    df = pd.read_excel(excel_path, sheet_name='상세일정', engine='openpyxl')
    
    # Phase mapping dictionary
    phase_to_l1_wbs = {
        "Project start": "1",
        "Basic Design": "2",
        "도면 승인, Rule 검토": "2",
        "Basic Design & Specification": "2",
        "주요설계 검토 & VP 승인": "3",
        "주요설계 검토": "3",
        "주요설계 검토 & VP승인": "3",
        "Technical report - Blow Down and Vent system": "4",
        "인력채용": "5",
        "Engineering work": "6",
        "상세설계 및 외주": "7",
        "상세설계 및도면": "7",
        "상세설계 및 도면": "7",
        "1st 3D Modeling Review": "8",
        "2nd 3D Modeling Review": "9",
        "Packaging 도면 발주 및 승인": "10",
        "Packaging 도면 발주": "10",
        "업무 지연 정리": "11",
        "AIP 준비 및 승인": "12",
        "AIP 준비": "12",
        "원가 분석 및 최적화": "13",
        "Basic Design for 2nd Line up": "14"
    }

    structured_tasks = []
    l1_tasks_by_row = {}
    l1_idx = 0

    # First pass: Extract Level 1.0 tasks
    for idx, row in df.iterrows():
        if idx < 2:
            continue
        desc = row.get('Description')
        if pd.isna(desc) or str(desc).strip() == '':
            continue
        level = str(row.get('Level')).strip()
        if level.endswith('.0'):
            level = level[:-2]
            
        if level == '1':
            l1_idx += 1
            wbs = str(l1_idx)
            l1_tasks_by_row[idx] = wbs

    # Second pass: Build structured tree
    children_counts = {}
    active_l2_wbs = ""
    last_phase_val = ""

    for idx, row in df.iterrows():
        if idx < 2:
            continue
        desc = row.get('Description')
        if pd.isna(desc) or str(desc).strip() == '':
            continue
        level = str(row.get('Level')).strip()
        if level.endswith('.0'):
            level = level[:-2]
            
        desc_str = str(desc).strip()
        desc_clean = " ".join(desc_str.split())
        
        t_type_val = str(row.get('T', '')).strip()
        t_type = "task"
        if t_type_val == 'M':
            t_type = "milestone"
        elif t_type_val == 'G':
            t_type = "project"
            
        plan_start = row.get('Plan')
        plan_end = row.get('Unnamed: 12')
        actual_start = row.get('Actual')
        actual_end = row.get('Unnamed: 15')
        
        def clean_date(d_val):
            if pd.isna(d_val):
                return ""
            if isinstance(d_val, (datetime, pd.Timestamp)):
                return d_val.strftime("%Y-%m-%d")
            d_str = str(d_val).split()[0].strip()
            if len(d_str) >= 10 and d_str[4] == '-' and d_str[7] == '-':
                return d_str[:10]
            return ""

        start_date = clean_date(plan_start)
        end_date = clean_date(plan_end)
        
        # Progress
        progress = 0
        act_start_str = clean_date(actual_start)
        act_end_str = clean_date(actual_end)
        if act_end_str != "":
            progress = 100
        elif act_start_str != "":
            progress = 50

        # Assignees
        dept_val = str(row.get('부서', '')).strip() if pd.notna(row.get('부서')) else ''
        assignee_val = str(row.get('담당자', '')).strip() if pd.notna(row.get('담당자')) else ''
        if dept_val in ['nan', '-']: dept_val = ''
        if assignee_val in ['nan', '-']: assignee_val = ''

        # Generate WBS & Predecessors
        wbs = ""
        preds = ""
        fixing = "None"
        
        if level == '1':
            wbs = l1_tasks_by_row[idx]
            
        elif level == '2':
            phase_col_val = row.get('Engineering phase')
            if pd.isna(phase_col_val) or str(phase_col_val).strip() == '':
                phase_val = last_phase_val
            else:
                phase_val = str(phase_col_val).strip()
                last_phase_val = phase_val
                
            parent_l1_wbs = phase_to_l1_wbs.get(phase_val, "2")
            count = children_counts.get(parent_l1_wbs, 0) + 1
            children_counts[parent_l1_wbs] = count
            wbs = f"{parent_l1_wbs}.{count}"
            active_l2_wbs = wbs
            
        elif level == '3':
            count = children_counts.get(active_l2_wbs, 0) + 1
            children_counts[active_l2_wbs] = count
            wbs = f"{active_l2_wbs}.{count}"
            
            # Determine predecessor for Level 3 activity
            if count > 1:
                preds = f"{active_l2_wbs}.{count - 1}FS"
            else:
                # First task in group
                parent_l1_wbs = active_l2_wbs.split('.')[0]
                if parent_l1_wbs == "2":
                    preds = "1FS"
                elif parent_l1_wbs == "3":
                    if active_l2_wbs == "3.1":
                        preds = "2.4.3FS"
                    elif active_l2_wbs == "3.2":
                        preds = "2.3.3FS"
                    elif active_l2_wbs == "3.3":
                        preds = "2.5.6FS"
                elif parent_l1_wbs == "6":
                    if active_l2_wbs == "6.1":
                        preds = "2.2.5FS"
                    elif active_l2_wbs == "6.2":
                        preds = "2.3.3FS"
                    elif active_l2_wbs == "6.3":
                        preds = "2.4.3FS"
                    elif active_l2_wbs == "6.4":
                        preds = "2.5.6FS"
                    elif active_l2_wbs == "6.5":
                        preds = "2.5.6FS"
                elif parent_l1_wbs == "7":
                    if active_l2_wbs == "7.1":
                        preds = "2.5.6FS"

        if wbs == "2.1.3":
            fixing = "Fix StartDate"
            start_date = "2026-05-01"
            end_date = "2026-05-31"
        elif wbs == "2.2.1":
            fixing = "Fix Duration"

        load_val = 100 if assignee_val else 0
        mh_val = 0.0
        if start_date and end_date:
            try:
                s = datetime.strptime(start_date, "%Y-%m-%d")
                e = datetime.strptime(end_date, "%Y-%m-%d")
                dur = (e - s).days + 1
                mh_val = round(dur * 8 * (load_val / 100), 1)
            except ValueError:
                pass

        task_node = {
            "WBS": wbs,
            "TaskName": desc_clean,
            "StartDate": start_date,
            "EndDate": end_date,
            "ActualStartDate": act_start_str,
            "ActualEndDate": act_end_str,
            "Progress": progress,
            "Predecessors": preds,
            "PrimaryAssignee": assignee_val,
            "SecondaryAssignee": dept_val,
            "Type": t_type,
            "Fixing": fixing,
            "SchedulingMode": "manual",
            "Load": load_val,
            "MH": mh_val
        }
        structured_tasks.append(task_node)

    # Clean up project tasks that have no children
    has_children = set()
    for t in structured_tasks:
        wbs_parts = t["WBS"].split('.')
        if len(wbs_parts) > 1:
            parent_wbs = ".".join(wbs_parts[:-1])
            has_children.add(parent_wbs)
            if len(wbs_parts) > 2:
                grandparent_wbs = wbs_parts[0]
                has_children.add(grandparent_wbs)

    for t in structured_tasks:
        if t["Type"] == "project" and t["WBS"] not in has_children:
            t["Type"] = "task"

    return structured_tasks

def parse_wbs_and_rollup(df):
    """
    Cleans and structures standard dataframe, then performs rollup.
    """
    raw_tasks = parse_standard_excel(df)
    return rollup_summary_data(raw_tasks)

def generate_gantt_html(excel_path, template_path="gantt_template.html", output_path="gantt_chart.html"):
    """
    Reads the excel schedule sheet, aggregates data, and compiles the interactive HTML chart.
    """
    if not os.path.exists(excel_path):
        fallback_json = "detailed_parsed_structured_perfect.json"
        if os.path.exists(fallback_json):
            print(f"Excel file not found at: {excel_path}. Falling back to cached JSON: {fallback_json}")
            excel_path = None
        else:
            raise FileNotFoundError(f"Excel file not found at: {excel_path}")
            
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Gantt template file not found at: {template_path}")
        
    print(f"Reading schedule from Excel: {excel_path}...")
    try:
        if excel_path is None:
            raise FileNotFoundError("Excel path was None, forcing fallback")
        if "regas" in os.path.basename(excel_path).lower():
            # Use custom parser for Regas Module
            tasks_data = parse_regas_excel_data(excel_path)
            # Apply rollup to calculate summary project dates
            tasks_data = rollup_summary_data(tasks_data)
        else:
            df = pd.read_excel(excel_path, engine='openpyxl')
            tasks_data = parse_wbs_and_rollup(df)
    except Exception as e:
        print(f"Excel reading failed: {e}")
        fallback_json = "detailed_parsed_structured_perfect.json"
        if os.path.exists(fallback_json):
            print(f"Falling back to cached JSON data: {fallback_json}")
            with open(fallback_json, "r", encoding="utf-8") as f:
                tasks_data = json.load(f)
        else:
            raise e
    
    print(f"Loading template: {template_path}...")
    with open(template_path, "r", encoding="utf-8") as f:
        html_template = f.read()
        
    # Serialize tasks to JSON
    tasks_json = json.dumps(tasks_data, ensure_ascii=False, indent=2)
    
    # Inject tasks into JavaScript variable
    print("Injecting tasks data into template...")
    pattern = r'const initialTasks = \[.*?\];'
    replacement = f'const initialTasks = {tasks_json};'
    
    html_output = re.sub(pattern, lambda m: replacement, html_template, flags=re.DOTALL)
    
    # Write to output file
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_output)
        
    print(f"Gantt chart HTML successfully generated: {os.path.abspath(output_path)}")
    return os.path.abspath(output_path)

if __name__ == "__main__":
    excel_file = "Project Schedule_Regas module_20260417.xlsx"
    template_file = "gantt_template.html"
    output_file = "gantt_chart.html"
    
    try:
        generate_gantt_html(excel_file, template_file, output_file)
    except Exception as e:
        print(f"Error occurred during HTML generation: {e}")

