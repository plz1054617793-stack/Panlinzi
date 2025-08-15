import streamlit as st
import pandas as pd
from pathlib import Path
import re
from io import BytesIO

# 设置标题和侧边栏
st.title("A320新飞机引进机身项目自动化处理平台")
sidebar = st.sidebar
sidebar.title("设置")
sidebar.markdown("**作者：周福来**")
plane_model = sidebar.selectbox("飞机型号", ["A318", "A319", "A320", "A321"], key="plane_model_select")
plane_submodel = sidebar.selectbox("飞机子型号", ["A321-201", "None"], key="plane_submodel_select")
engine_model = sidebar.selectbox("发动机型号", ["CFM56-5", "IAE", "PW1100G", "LEAP-1A"], key="engine_model_select")
engine_submodel = sidebar.selectbox("发动机子型号", ["CFM56-5A", "CFM56-5B", "PW1100G-JM", "PW6122A", "PW6124A", "V2500-A1", "V2500-A5", "None"], key="engine_submodel_select")
plane_registration = sidebar.text_input("飞机注册号/MSN号", key="plane_registration_input")

# 检查是否所有字段都已填写
if not all([plane_model, plane_submodel, engine_model, engine_submodel, plane_registration]):
    st.error("请填写所有必填项")
    st.stop()

# 文件上传和处理的函数
def upload_and_process_file(file_uploader_label, file_types):  
    uploaded_file = st.file_uploader(file_uploader_label, type=file_types)  
    if uploaded_file:  
        file_extension = uploaded_file.name.split('.')[-1]  # 获取文件扩展名  
        if file_extension in ['xlsx', 'xls']:  
            excel_file = pd.ExcelFile(uploaded_file)  
            selected_sheet = st.selectbox("选择 Sheet 表单", excel_file.sheet_names, key=f"{file_uploader_label.replace(' ', '_')}_sheet_select")  
            df = pd.read_excel(excel_file, sheet_name=selected_sheet)  
        elif file_extension == 'csv':  
            df = pd.read_csv(uploaded_file)  # 直接读取 CSV 文件  
        else:  
            st.error("不支持的文件类型")  
            return None  
          
        st.write(f"{file_uploader_label.split('文件上传')[0]} 内容预览:")  
        st.write(df)  
        return df  
    return None 

# 上传各种文件
config_df = upload_and_process_file("构型差异文件上传", ["xlsx", "xls","csv"])
mod_df = upload_and_process_file("MOD 文件上传", ["xlsx", "xls","csv"])
sb_df = upload_and_process_file("SB 文件上传", ["xlsx", "xls","csv"])
mpd_df = upload_and_process_file("MPD 文件上传", ["xlsx", "xls","csv"])
maintenance_df = upload_and_process_file("维修方案飞机明细", ["xlsx", "xls","csv"])

# 判断构型差异的函数（通用函数，可用于多种情况，此处也用于MPD的判断）
def evaluate_configuration_formula(row, plane_model, plane_submodel, engine_model, engine_submodel, mod_df, sb_df):
    configuration_formula = row["CONFIGURATION FORMULA"]
    # 获取原文件除"CONFIGURATION FORMULA"外的其他列内容
    other_columns = {col: row[col] for col in row if col!= "CONFIGURATION FORMULA"}
    lines = re.split('\n', configuration_formula.strip())

    # 初始化 results 列表，长度与 lines 相同
    results = [""] * len(lines)

    # 第一次判断（MOD 匹配）
    mod_numbers_with_lines = [(line.split()[0], line.split()[1]) for line in lines if line.startswith(("POST", "PRE"))]
    for i, original_line in enumerate(lines):
        if "OR" in original_line:
            results[i] = original_line
            continue
        new_line = ""
        if original_line.startswith("POST") or original_line.startswith("PRE"):
            match = re.match(r'^(POST|PRE)\s+(.*)$', original_line)
            if match:
                prefix = match.group(1)
                mod_number = match.group(2).replace(" ", "")
            match_found = any(mod_number.strip() in str(mod) for mod in mod_df["MOD"])
            if prefix == "POST" and match_found:
                new_line = f"{original_line}成立"
                results[i] = new_line
            elif prefix == "POST" and not match_found:
                new_line = f"{original_line}不符合"
                results[i] = new_line
            elif prefix == "PRE" and not match_found:
                new_line = f"{original_line}成立"
                results[i] = new_line
            elif prefix == "PRE" and match_found:
                new_line = f"{original_line}不符合"
                results[i] = new_line
            else:
                new_line = original_line
                results[i] = new_line
        elif re.match(r'^\s*OR\s*$', original_line):
            new_line = original_line
            results[i] = new_line
        else:
            new_line = original_line
            results[i] = new_line
    # 第一次判断后的结果作为第二次判断（MOD匹配）的输入
    second_judgment_lines = results.copy()

    # 第二次判断（SB 匹配）
    new_results_sb = []
    for i, line in enumerate(second_judgment_lines):
        if "OR" in line or line.startswith("A") or line.startswith("LEAP") or line.startswith("PW") or line.startswith("IAE") or line.startswith("CFM") or line.startswith("V25"):
            new_results_sb.append(line)
            continue
        number = None
        if line.startswith("("):
            match = re.search(r'\((.*?)\)', line)
            if match:
                number = match.group(1)
            else:
                new_results_sb.append(line)
                continue
        current_line_index = lines.index(line) if line in lines else -1
        if current_line_index == -1:
            new_results_sb.append(line)
            continue
        post_found = False
        pre_found = False
        first_judgment_status = None
        for j in range(current_line_index - 1, -1, -1):
            reversed_line = lines[j]
            if reversed_line.startswith("POST"):
                post_found = True
                first_judgment_status = "成立" if "成立" in reversed_line else "不符合"
                break
            elif reversed_line.startswith("PRE"):
                pre_found = True
                break
        sb_match_found = False
        if number is not None:  # 只有当number被赋值后才进行下面的匹配操作
            for sb_num in sb_df["SB号"]:
                if re.match(r'^' + re.escape(number) + r'$', str(sb_num).strip()):
                    sb_match_found = True
                    break
        if post_found:
            if sb_match_found:
                new_line = f"{line}成立"
            else:
                new_line = f"{line}不符合"
        elif pre_found:
            if sb_match_found:
                new_line = f"{line}不符合"
            else:
                new_line = f"{line}成立"
        else:
            new_line = line
        new_results_sb.append(new_line)
    results = new_results_sb

    # 第三次判断后的结果作为第二次判断的输入
    third_judgment_lines = results.copy()

    # 第三次判断
    third_judgment_results = []  # 创建新列表存储第二次判断的结果
    for i, line in enumerate(third_judgment_lines):
        if line == "OR" or line.startswith(("POST", "PRE", "(")):
            third_judgment_results.append(line)  # 直接添加特殊行到新列表
        else:
            plane_match = (plane_model in line) or (line == "None" and plane_model not in ["A318", "A319", "A20", "A321"])
            plane_sub_match = (plane_submodel in line) or (line == "None" and plane_submodel == "None")
            engine_match = (engine_model in line) or (line == "None" and engine_model not in ["CFM56-5", "IAE", "PW1100G", "LEAP-1A"])
            engine_sub_match = (engine_submodel in line) or (line == "None" and engine_submodel == "None")
            result = (plane_match or plane_sub_match or engine_match or engine_sub_match)
            line_result = "成立" if result else "不符合"
            third_judgment_results.append(f"{line}{line_result}")  # 添加处理后的行到新列表
    results.clear()
    results.extend(third_judgment_results)  # 将新列表的内容添加到results

    # 第四次判断
    fourth_judgment_lines = third_judgment_results.copy()
    for i, line in enumerate(fourth_judgment_lines):
        if "OR" in line:
            fourth_judgment_lines[i] = line
            continue
        if line.startswith("(") and "成立" in line:
            post_found = False
            for j in range(i - 1, -1, -1):
                reversed_line = fourth_judgment_lines[j]
                if reversed_line.startswith("POST"):
                    post_found = True
                    if "成立" in reversed_line:
                        break
                    elif "不符合" in reversed_line:
                        updated_line = re.sub(r'不符合', 'SB替代成立', reversed_line)
                        fourth_judgment_lines[j] = updated_line
                        break
                else:
                    new_line = line
                if not post_found:
                    break
    results = fourth_judgment_lines

    # 第五次判断
    fifth_judgment_lines = results.copy()
    new_fifth_judgment_lines = []
    processing_post_block = False  # 布尔变量，用于跟踪是否正在处理POST块
    processed_parentheses_lines = set()  # 用于记录已经处理过的以"("开头的行及其索引

    for i, line in enumerate(fifth_judgment_lines):
        if processing_post_block:
            # 如果正在处理POST块
            if line.startswith("(") and (i, line) not in processed_parentheses_lines:
                parentheses_content = line.strip("成立").strip("不符合")  # 去除前后的括号和空格
                if parentheses_content:
                    # 对以"("开头的行进行处理，并添加到结果中
                    new_temp_line = f"{parentheses_content} MOD替代成立"
                    new_fifth_judgment_lines.append(new_temp_line)
                    processed_parentheses_lines.add((i, line))
            elif line == "OR" or line.startswith("PRE") or line.startswith("POST"):  # 遇到OR或PRE行，退出POST块处理
                new_fifth_judgment_lines.append(line)
                processing_post_block = False
        else:
            # 如果不在处理POST块
            if "OR" in line or line.startswith(("A", "LEAP", "PW", "IAE", "CFM", "V25")):
                # 直接添加这些行到结果中
                new_fifth_judgment_lines.append(line)
            elif line.startswith("POST"):
                # 添加POST行到结果中
                new_fifth_judgment_lines.append(line)
                # 检查是否包含“成立”
                if "成立" in line:
                    # 设置标志，开始处理POST块
                    processing_post_block = True
            else:
                # 对于其他行，如果不在POST块中，则添加到结果中
                new_fifth_judgment_lines.append(line)
    # 更新结果
    results = new_fifth_judgment_lines
    
    # 合并结果并输出
    config_diff_details = "\n".join(results)
    # 优化后的确定final_result逻辑（这里假设整体的判断逻辑与原代码类似，可根据实际情况调整）
    if "OR" not in configuration_formula:
        all_established = all("不符合" not in res for res in results if res!= "OR")
        final_result = "适用于此构型差异" if all_established else "不适用于此构型差异"
    else:
        or_lines_indices = [i for i, line in enumerate(lines) if line == "OR"]
        first_or_index = or_lines_indices[0] if or_lines_indices else len(lines)
        last_or_index = or_lines_indices[-1] if or_lines_indices else -1

        before_or_established = all("不符合" not in res for res in results[:first_or_index] if res!= "OR")
        after_or_established = all("不符合" not in res for res in results[last_or_index + 1:] if res!= "OR")

        or_groups = []
        start_index = first_or_index
        for end_index in or_lines_indices[1:]:
            or_groups.append(results[start_index + 1:end_index])
            start_index = end_index

        or_groups_established = any(all("不符合" not in res for res in group if res!= "OR") for group in or_groups)

        final_result = "适用于此构型差异" if (before_or_established or after_or_established or or_groups_established) else "不适用于此构型差异"

    return config_diff_details, final_result

# 判断MPD函数
def evaluate_mpd_configuration(row, plane_model, plane_submodel, engine_model, engine_submodel, mod_df, sb_df):
    applicability_text = row["APPLICABILITY"]
    if not isinstance(applicability_text, str):
        applicability_text = str(applicability_text)
    lines = re.split('\n', applicability_text.strip())

    # 初始化 results 列表，长度与 lines 相同
    results = [""] * len(lines)

    # 第一次判断（MOD匹配）
    def perform_mod_match(line, mod_df):
        if line.startswith(("POST", "PRE")):
            match = re.match(r'^(POST|PRE)\s+(.*)$', line)
            if match:
                prefix = match.group(1)
                mod_number = match.group(2).replace(" ", "")
                match_found = any(mod_number.strip() in str(mod) for mod in mod_df["MOD"])
                if prefix == "POST" and match_found:
                    return f"{line}成立"
                elif prefix == "POST" and not match_found:
                    return f"{line}不符合"
                elif prefix == "PRE" and not match_found:
                    return f"{line}成立"
                elif prefix == "PRE" and match_found:
                    return f"{line}不符合"
        return line
    for i, original_line in enumerate(lines):
        if "OR" in original_line:
            results[i] = original_line
            continue
        results[i] = perform_mod_match(original_line, mod_df)
    

    # 第二次判断（SB匹配） 
    def perform_sb_match(line, sb_df):
        if "OR" in line or re.match(r'^[A-Za-z]', line): 
            return line
        number = None
        if line.startswith("("):
            match = re.search(r'\((.*?)\)', line)
            if match:
                number = match.group(1)
            else:
                return line
        post_found = False
        pre_found = False
        first_judgment_status = None
        for j in range(len(lines) - 1, -1, -1):
            reversed_line = lines[j]
            if reversed_line.startswith("POST"):
                post_found = True
                first_judgment_status = "成立" if "成立" in reversed_line else "不符合"
                break
            elif reversed_line.startswith("PRE"):
                pre_found = True
                break
        sb_match_found = False
        if number is not None:
            for sb_num in sb_df["SB号"]:
                if re.match(r'^' + re.escape(number) + r'$', str(sb_num).strip()):
                    sb_match_found = True
                    break
        if post_found:
            if sb_match_found:
                return f"{line}成立"
            else:
                return f"{line}不符合"
        elif pre_found:
            if sb_match_found:
                return f"{line}不符合"
            else:
                return f"{line}成立"
        return line
    second_judgment_lines = results.copy()
    for i, line in enumerate(second_judgment_lines):
        results[i] = perform_sb_match(line, sb_df)
    
    # 第三次判断（机型匹配） 
    def is_exact_match(target, line):  #检查去除空格后的行内容是否与目标字符串完全一致"""  
        return target == line.strip() if line else False  
    # 第三次判断  
    third_judgment_results = []  
    for i, line in enumerate(results):  
        stripped_line = line.strip()  
        if stripped_line == "OR" or stripped_line.startswith(("POST", "PRE", "(")):  
            third_judgment_results.append(line)  
        else:  
            plane_match = is_exact_match(plane_model, stripped_line)  
            plane_sub_match = is_exact_match(plane_submodel, stripped_line)  
            engine_match = is_exact_match(engine_model, stripped_line)  
            engine_sub_match = is_exact_match(engine_submodel, stripped_line)  
          
            result = (plane_match or plane_sub_match or engine_match or engine_sub_match)  
            line_result = "成立" if result else "不符合"  
            third_judgment_results.append(f"{line}{line_result}")  
    results = third_judgment_results

    # 第四次判断  
    fourth_judgment_results = []  
    group_item_set = set(config_df["GROUP ITEM"])  # 将 GROUP ITEM 列的值存储在一个集合中  
    for i, line in enumerate(results):  
        stripped_line = line.strip()  
        # 处理 "ALL" 的情况  
        if "ALL" in stripped_line:  
            fourth_judgment_results.append(f"{stripped_line.strip("成立").strip("不符合")}成立")  
            continue  
        # 特殊处理 "PRE-SPR-CURVE" 和 "POST-SPR-CURVE"  
        if stripped_line in ["PRE-SPR-CURVE", "POST-SPR-CURVE"]:  
            if stripped_line in group_item_set:  
                group_item_match = config_df[config_df["GROUP ITEM"] == stripped_line]  
                if not group_item_match.empty:  
                    corresponding_result = group_item_match["构型差异判断结果"].iloc[0]  
                    if corresponding_result == "适用于此构型差异":  
                        fourth_judgment_results.append(f"{stripped_line}成立")  
                    else:  
                        fourth_judgment_results.append(f"{stripped_line}不符合")  
                else:  
                    # 如果理论上不应该发生，但为了健壮性还是加上  
                    fourth_judgment_results.append(f"{stripped_line}不符合")  # 假设未找到匹配时默认为不符合  
            else:  
                # 如果不在 GROUP ITEM 集合中（理论上不应该发生），直接添加不符合  
                fourth_judgment_results.append(f"{stripped_line}不符合")  
            continue
        # 处理 "OR" 和特定前缀的情况  
        if stripped_line == "OR" or stripped_line.startswith(("POST", "PRE", "(")):
            fourth_judgment_results.append(line)  
            continue  
        # 检查 GROUP ITEM  
        if stripped_line in group_item_set:  
            group_item_match = config_df[config_df["GROUP ITEM"] == stripped_line]  
            if not group_item_match.empty:  
                corresponding_result = group_item_match["构型差异判断结果"].iloc[0]  
                if corresponding_result == "适用于此构型差异":  
                    fourth_judgment_results.append(f"{stripped_line}成立")  
                else:  
                    fourth_judgment_results.append(line)  # 或者可以添加 "不符合" 后缀，根据需要  
            else:  
                # 如果理论上不应该发生，但为了健壮性还是加上  
                fourth_judgment_results.append(line)  
        else:  
            # 如果不在 GROUP ITEM 集合中，直接添加原行（或者可以添加 "不符合" 后缀）  
            fourth_judgment_results.append(line)  
    results = fourth_judgment_results 

    # 第五次判断
    fifth_judgment_lines = fourth_judgment_results.copy()
    for i, line in enumerate(fifth_judgment_lines):
        if "OR" in line:
            fifth_judgment_lines[i] = line
            continue
        if line.startswith("(") and "成立" in line:
            post_found = False
            for j in range(i - 1, -1, -1):
                reversed_line = fifth_judgment_lines[j]
                if reversed_line.startswith("POST"):
                    post_found = True
                    if "成立" in reversed_line:
                        break
                    elif "不符合" in reversed_line:
                        updated_line = re.sub(r'不符合', 'SB替代成立', reversed_line)
                        fifth_judgment_lines[j] = updated_line
                        break
                else:
                    new_line = line
                if not post_found:
                    break
    results = fifth_judgment_lines

    # 第六次判断
    sixth_judgment_lines = results.copy()
    new_sixth_judgment_lines = []
    processing_post_block = False  # 布尔变量，用于跟踪是否正在处理POST块
    processed_parentheses_lines = set()  # 用于记录已经处理过的以"("开头的行及其索引

    for i, line in enumerate(sixth_judgment_lines):
        if processing_post_block:
            # 如果正在处理POST块
            if line.startswith("(") and (i, line) not in processed_parentheses_lines:
                parentheses_content = line.strip("成立").strip("不符合")  # 去除前后的括号和空格
                if parentheses_content:
                    # 对以"("开头的行进行处理，并添加到结果中
                    new_temp_line = f"{parentheses_content} MOD替代成立"
                    new_sixth_judgment_lines.append(new_temp_line)
                    processed_parentheses_lines.add((i, line))
            elif line == "OR" or line.startswith("PRE") or line.startswith("POST"):  # 遇到OR或PRE行，退出POST块处理
                new_sixth_judgment_lines.append(line)
                processing_post_block = False
        else:
            # 如果不在处理POST块
            if "OR" in line or line.startswith(("A", "LEAP", "PW", "IAE", "CFM", "V25")):
                # 直接添加这些行到结果中
                new_sixth_judgment_lines.append(line)
            elif line.startswith("POST"):
                # 添加POST行到结果中
                new_sixth_judgment_lines.append(line)
                # 检查是否包含“成立”
                if "成立" in line:
                    # 设置标志，开始处理POST块
                    processing_post_block = True
            else:
                # 对于其他行，如果不在POST块中，则添加到结果中
                new_sixth_judgment_lines.append(line)
    # 更新结果
    results = new_sixth_judgment_lines

    # 合并结果并输出
    config_diff_details = "\n".join(results)

    # 优化后的确定final_result逻辑（这里假设整体的判断逻辑与原代码类似，可根据实际情况调整）
    if "OR" not in applicability_text:
        all_established = all("不符合" not in res for res in results if res!= "OR")
        final_result = "适用于此项目" if all_established else "不适用于此项目"
    else:
        or_lines_indices = [i for i, line in enumerate(lines) if line == "OR"]
        first_or_index = or_lines_indices[0] if or_lines_indices else len(lines)
        last_or_index = or_lines_indices[-1] if or_lines_indices else -1

        before_or_established = all("不符合" not in res for res in results[:first_or_index] if res!= "OR")
        after_or_established = all("不符合" not in res for res in results[last_or_index + 1:] if res!= "OR")

        or_groups = []
        start_index = first_or_index
        for end_index in or_lines_indices[1:]:
            or_groups.append(results[start_index + 1:end_index])
            start_index = end_index

        or_groups_established = any(all("不符合" not in res for res in group if res!= "OR") for group in or_groups)

        final_result = "适用于此项目" if (before_or_established or after_or_established or or_groups_established) else "不适用于此项目"

    return config_diff_details, final_result

# 第七次判断逻辑
def evaluate_maintenance_statement(mpd_df, maintenance_df, plane_registration):
    # 添加新列以存储MP判断结果和相关状态
    mpd_df['飞机明细是否包含'] = "否"
    mpd_df['营运人是否有此条目'] = "否"
    mpd_df['主MP是否需要改版'] = "否"
    mpd_df['营运人MP是否需要改版'] = "否"
    mpd_df['MP判断结果'] = ""

    # 遍历MPD DataFrame
    for index, row in mpd_df.iterrows():
        # 确保TASK NUMBER是字符串类型
        task_number = str(row['TASK NUMBER']).strip() if pd.notnull(row['TASK NUMBER']) else ""
        mpd_result = row['MPD判断明细结果']  # 获取MPD判断明细结果

        # 检查项目号一致性
        task_number_stripped = task_number  # 已经去除空格
        project_match = maintenance_df[maintenance_df['项目号'].apply(lambda x: str(x).strip()) == task_number_stripped]

        # 检查营运人是否有此项目
        operator_has_project = not project_match[project_match['飞机明细'].notna()].empty

        # 检查飞机明细是否包含此项目
        # 将飞机明细列中的字符串分割为列表
        plane_details = [p.strip() for p in str(project_match['飞机明细'].iloc[0]).split(',')] if not project_match.empty else []
        plane_included = plane_registration.strip() in plane_details

        # 更新相关状态列
        mpd_df.at[index, '飞机明细是否包含'] = "是" if plane_included else "否"
        mpd_df.at[index, '营运人是否有此条目'] = "是" if operator_has_project else "否"

        # 根据逻辑确定MP判断结果
        if project_match.empty:
            if operator_has_project and mpd_result == "适用于此项目" and not plane_included:
                mpd_df.at[index, 'MP判断结果'] = "此项目适用于此架飞机，现行MP无此项目，飞机明细不包含此架飞机，需要新增"
            if not operator_has_project and mpd_result == "不适用于此项目" and not plane_included:
                mpd_df.at[index, 'MP判断结果'] = "此项目不适用于此架飞机，现行MP无此项目，飞机明细不包含此架飞机，需要新增"
        else:
            if operator_has_project and mpd_result == "适用于此项目" and plane_included:
                mpd_df.at[index, 'MP判断结果'] = "此项目适用于此架飞机，现行MP有此项目，飞机明细包含此架飞机，无需改版"
            elif operator_has_project and mpd_result == "适用于此项目" and not plane_included:
                mpd_df.at[index, 'MP判断结果'] = "此项目适用于此架飞机，现行MP有此项目，飞机明细不包含此架飞机，需要改版-手动复核飞机明细"
            elif operator_has_project and mpd_result != "适用于此项目" and plane_included:
                mpd_df.at[index, 'MP判断结果'] = "此项目不适用于此架飞机，现行MP有此项目，飞机明细包含此架飞机，需要改版"
            elif operator_has_project and mpd_result != "适用于此项目" and not plane_included:
                mpd_df.at[index, 'MP判断结果'] = "此项目不适用于此架飞机，现行MP无此项目，飞机明细不包含此架飞机，无需改版"

        # 根据MP判断结果更新是否需要改版的列
        if "需要改版" in mpd_df.at[index, 'MP判断结果']:
            mpd_df.at[index, '主MP是否需要改版'] = "是"
            mpd_df.at[index, '营运人MP是否需要改版'] = "是"

    return mpd_df
    
# 合并后的评估函数
def evaluate_all(mpd_df, maintenance_df, plane_registration, plane_model, plane_submodel, engine_model, engine_submodel, mod_df, sb_df):
    # 首先进行MPD配置评估
    mpd_df[["MPD判断明细", "MPD判断明细结果"]] = mpd_df.apply(
        lambda row: pd.Series(evaluate_mpd_configuration(row, plane_model, plane_submodel, engine_model, engine_submodel, mod_df, sb_df)),
        axis=1
    )
    
    # 然后进行维修方案评估
    mpd_df = evaluate_maintenance_statement(mpd_df, maintenance_df, plane_registration)
    
    return mpd_df

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    # 使用writer.close()替代writer.save()
    writer.close()
    processed_data = output.getvalue()
    return processed_data

# “构型差异评估”按钮逻辑（原代码部分）
if st.button("构型差异评估", key="execute_config_diff_button"):
    if config_df is not None and "CONFIGURATION FORMULA" in config_df.columns:
        config_df[["构型差异明细", "构型差异判断结果"]] = config_df.apply(
            lambda row: pd.Series(evaluate_configuration_formula(row, plane_model, plane_submodel, engine_model, engine_submodel, mod_df, sb_df)),
            axis=1
        )
        result_df = config_df  # 更新 result_df
        st.write("构型差异判断结果:")
        st.write(result_df)

        # 添加下载按钮及逻辑
        excel_data = to_excel(result_df)
        st.download_button(
            label="下载构型差异评估结果",
            data=excel_data,
            file_name="构型差异评估结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("没有可处理的构型差异文件或缺少必要的列")


# 按钮逻辑
if st.button("维修方案评估", key="execute_all_evaluation_button"):
    if (mpd_df is not None and "APPLICABILITY" in mpd_df.columns and
        maintenance_df is not None and plane_registration):
        # 执行全部评估
        result_df = evaluate_all(mpd_df, maintenance_df, plane_registration, plane_model, plane_submodel, engine_model, engine_submodel, mod_df, sb_df)
        st.write("飞机适用性评估:")
        st.write(result_df)

        # 添加下载按钮及逻辑
        excel_data = to_excel(result_df)
        st.download_button(
            label="飞机适用性评估",
            data=excel_data,
            file_name="飞机适用性评估.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("请确保所有必要的文件和字段都已正确填写和上传。")

# 新增按钮逻辑
if st.button("新飞机引进评估", key="execute_full_evaluation_button"):
    # 检查所有必要的文件是否已上传
    if not all([df is not None for df in [config_df, mod_df, sb_df, mpd_df, maintenance_df]]):
        st.error("请确保所有必要的文件都已正确上传。")
    else:
        # 检查所有必要的列是否存在
        if not all(["CONFIGURATION FORMULA" in config_df.columns, "APPLICABILITY" in mpd_df.columns, "项目号" in maintenance_df.columns, "飞机明细" in maintenance_df.columns]):
            st.error("请确保所有必要的列都存在于文件中。")
        elif not plane_registration:
            st.error("飞机注册号/MSN号不能为空，请填写。")
        else:
            # 进行构型差异评估
            if "CONFIGURATION FORMULA" in config_df.columns:
                config_df[["构型差异明细", "构型差异判断结果"]] = config_df.apply(
                    lambda row: pd.Series(evaluate_configuration_formula(row, plane_model, plane_submodel, engine_model, engine_submodel, mod_df, sb_df)),
                    axis=1
                )
                st.write("构型差异评估结果:")
                st.write(config_df[["构型差异明细", "构型差异判断结果"]])
            else:
                st.error("构型差异文件中缺少必要的'CONFIGURATION FORMULA'列。")

            # 进行MPD配置评估
            if "APPLICABILITY" in mpd_df.columns:
                mpd_df[["MPD判断明细", "MPD判断明细结果"]] = mpd_df.apply(
                    lambda row: pd.Series(evaluate_mpd_configuration(row, plane_model, plane_submodel, engine_model, engine_submodel, mod_df, sb_df)),
                    axis=1
                )
            else:
                st.error("MPD文件中缺少必要的'APPLICABILITY'列。")

            # 进行维修方案评估
            if "项目号" in maintenance_df.columns and "飞机明细" in maintenance_df.columns:
                result_df = evaluate_maintenance_statement(mpd_df, maintenance_df, plane_registration)
                st.write("维修方案评估结果:")
                st.write(result_df)
            else:
                st.error("维修方案飞机明细文件中缺少必要的'项目号'或'飞机明细'列。")

            # 添加下载按钮及逻辑
            if not config_df.empty and not result_df.empty:
                excel_data_config = to_excel(config_df[["构型差异明细", "构型差异判断结果"]])
                excel_data_maintenance = to_excel(result_df)
                
                st.download_button(
                    label="下载构型差异评估结果",
                    data=excel_data_config,
                    file_name="构型差异评估结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.download_button(
                    label="下载维修方案评估结果",
                    data=excel_data_maintenance,
                    file_name="维修方案评估结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
