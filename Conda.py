import streamlit as st
import pandas as pd
import os
from pathlib import Path
import numpy as np
from scipy.interpolate import interp1d
from sklearn.preprocessing import LabelEncoder
import itertools
import datetime
import plotly.graph_objects as go
from pywt import wavedec


# 设置标题和作者
st.title("DAR处理平台")
# 创建侧边栏
sidebar = st.sidebar.radio("设置", ["数据处理", "数据应用"])
st.sidebar.markdown("作者: 周福来")


def convert_elements_type(elements, target_type, file_path, param_name):
    
    try:
        if isinstance(elements, list):
            converted_elements = [target_type(x) for x in elements]
        else:
            converted_elements = target_type(elements)
     
        if not isinstance(converted_elements, list):
            converted_elements = [converted_elements]
        return converted_elements
    except ValueError:
        st.error(f"无法将输入的限定参数元素转换为 {target_type.__name__} 类型，请检查输入的数据类型，文件路径：{file_path}，限定参数名称：{param_name}")
        return None

def calculate_with_limited_param(input_folder_path, param_name, limited_param_name, limited_param_elements):
    results = []
    input_folder = Path(input_folder_path)
    limited_param_elements = [limited_param_elements]  # 确保是列表

    # 检查文件夹路径是否存在
    if not input_folder.exists():
        st.error(f"输入的文件夹路径 {input_folder_path} 不存在，请重新输入。")
        return results

    for file_path in input_folder.glob('*.*'):
        if file_path.suffix.lower() in ['.csv', '.xlsx']:
            try:
                if file_path.suffix.lower() == '.csv':
                    data = pd.read_csv(file_path, skiprows=[0, 1, 2, 3, 4, 5, 6, 8, 9, 10], encoding='gbk', low_memory=False)
                else:
                    data = pd.read_excel(file_path)

                # 检查限定参数名称和查找参数名称对应的列是否存在
                if limited_param_name not in data.columns:
                    st.error(f"在文件 {file_path.name} 中未找到限定参数列 {limited_param_name}，请检查输入的限定参数名称是否正确，文件路径：{file_path}")
                    continue
                if param_name not in data.columns:
                    st.error(f"在文件 {file_path.name} 中未找到查找参数名称列 {param_name}，请检查输入的查找参数名称是否正确，文件路径：{file_path}")
                    continue

                limited_param_col = data[limited_param_name]
                # 获取限定参数列的第一个元素的数据类型作为目标类型
                target_type = type(limited_param_col.iloc[0])
                # 转换限定参数元素的数据类型
                converted_elements = convert_elements_type(limited_param_elements, target_type, file_path, limited_param_name)
                if converted_elements is None:
                    continue

                filtered_data = data[data[limited_param_name].isin(converted_elements)]
                if not filtered_data.empty:
                    element_stats = {
                        "文件名": file_path.name,
                        "平均值": filtered_data[param_name].mean(),
                        "最大值": filtered_data[param_name].max(),
                        "最小值": filtered_data[param_name].min(),
                        "方差": filtered_data[param_name].var()
                    }
                    results.append(element_stats)
            except Exception as e:
                print(f"处理文件 {file_path} 时出错: {e}")
                st.error(f"处理文件 {file_path} 时出现异常，请检查文件格式或数据内容是否正确，具体错误信息: {e}，文件路径：{file_path}")
    return results


# 定义draw_graph函数，使用Plotly
def draw_graph(folder_path, selected_column, filter_column=None, filter_min=None, filter_max=None):
    color_cycle = itertools.cycle(['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'])
    fig = go.Figure()
    if not os.path.exists(folder_path):
        st.error(f"输入的文件夹路径 {folder_path} 不存在，请重新输入。")
        return
    for filename in os.listdir(folder_path):
        if filename.endswith('.csv') or filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            try:
                df = pd.read_csv(file_path) if filename.endswith('.csv') else pd.read_excel(file_path)
                if selected_column in df.columns:
                    filtered_df = df
                    if filter_column in df.columns and filter_min is not None and filter_max is not None:
                        filtered_df = df[(df[filter_column] >= filter_min) & (df[filter_column] <= filter_max)]
                    fig.add_trace(go.Scatter(x=filtered_df.index, y=filtered_df[selected_column], mode='lines', name=os.path.basename(filename)))
            except Exception as e:
                print(f"Error reading file {file_path}: {e}")
                st.error(f"读取文件 {file_path} 时出错，请检查文件格式或权限等问题，具体错误信息: {e}")
    if fig.data:  # 检查是否有数据轨迹添加到图形中
        fig.update_layout(title=f'Distribution Plot of {selected_column} for Different Files', xaxis_title='Data Point Index', yaxis_title=selected_column)
        st.plotly_chart(fig)
    else:
        st.error("未找到可绘制的数据，请检查输入的参数以及文件内容。")


# 定义快速傅里叶变换函数
def fast_fourier_transform(folder_path, param_name):
    all_data = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.csv') or filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            df = pd.read_csv(file_path) if filename.endswith('.csv') else pd.read_excel(file_path)
            if param_name in df.columns:
                data = df[param_name].values
                all_data.extend(data)
    if all_data:
        fft_result = np.fft.fft(all_data)
        freq = np.fft.fftfreq(len(all_data))
        fig = go.Figure(data=[go.Scatter(x=freq, y=np.abs(fft_result), mode='lines')])
        fig.update_layout(title=f'Fast Fourier Transform of {param_name}', xaxis_title='Frequency', yaxis_title='Amplitude')
        st.plotly_chart(fig)
    else:
        st.error("参数列数据为空，请检查参数名称是否正确或文件夹中的文件内容。")


# 定义小波变换函数
def wavelet_transform(folder_path, param_name):
    all_data = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.csv') or filename.endslant('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            df = pd.read_csv(file_path) if filename.endswith('.csv') else pd.read_excel(file_path)
            if param_name in df.columns:
                data = df[param_name].values
                all_data.extend(data)
    if all_data:
        coeffs = wavedec(all_data, 'db1', level=1)
        cA, cD = coeffs[0], coeffs[-1]
        fig = go.Figure(data=[go.Scatter(x=np.arange(len(cA)), y=cA, mode='lines', name='Approximation'), go.Scatter(x=np.arange(len(cD)), y=cD, mode='lines', name='Detail')])
        fig.update_layout(title=f'Wavelet Transform of {param_name}', xaxis_title='Index', yaxis_title='Coefficients')
        st.plotly_chart(fig)
    else:
        st.error("参数列数据为空，请检查参数名称是否正确或文件夹中的文件内容。")


# 数据应用功能
if sidebar == "数据应用":
    uploaded_file = st.file_uploader("上传文件 (csv 或 excel)", type=["csv", "xlsx"])
    if uploaded_file is not None:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file)
        st.write("文件内容:")
        st.dataframe(df)
        col_names = df.columns.tolist()
        st.write("参数名称:")
        st.write(col_names)
        selected_col = st.selectbox("选择参数", col_names)
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=df.index, y=df[selected_col], mode='lines'))
        st.plotly_chart(fig)
    folder_path = st.text_input("输入文件夹地址")
    param_name = st.text_input("查找参数名称")
    use_filters = st.checkbox("应用限定参数")
    if use_filters:
        filter_col = st.text_input("限定参数名称")
        filter_min = st.number_input("限定参数最小值")
        filter_max = st.number_input("限定参数最大值")
        if st.button("绘制参数图"):
            if folder_path and param_name:
                try:
                    draw_graph(folder_path, param_name, filter_col, filter_min, filter_max)
                    st.success("参数图绘制完成")
                except Exception as e:
                    st.error(f"绘制参数图时出错: {e}")
            else:
                st.error("请输入有效的文件夹路径和参数名称")
    else:
        if st.button("绘制参数图"):
            if folder_path and param_name:
                try:
                    draw_graph(folder_path, param_name)
                    st.success("参数图绘制完成")
                except Exception as e:
                    st.error(f"绘制参数图时出错: {e}")
            else:
                st.error("请输入有效的文件夹路径和参数名称")
    transformation_type = st.selectbox("选择变换类型", ["快速傅里叶变换", "小波变换"])
    if st.button("函数执行"):
        if folder_path and param_name:
            if transformation_type == "快速傅里叶变换":
                fast_fourier_transform(folder_path, param_name)
            elif transformation_type == "小波变换":
                wavelet_transform(folder_path, param_name)
        else:
            st.error("请输入有效的文件夹路径和参数名称")

# 数据处理功能部分
if sidebar == "数据处理":
    input_folder_path = st.text_input("输入文件夹地址")
    norm_output_folder = st.text_input("归一化输出文件夹")
    param_name = st.text_input("查找参数名称")
    target_rows = st.number_input("归一化后的横坐标个数", min_value=1, value=10000)
    use_limited_param = st.checkbox("限定参数计算")
    use_compare_param = st.checkbox("对比参数功能")
    compare_param_name = st.text_input("对比参数名称") if use_compare_param else None
    
    def process_and_normalize(input_file_path, target_rows, output_file_path):
        # 根据文件扩展名读取文件
        if input_file_path.endswith('.csv'):
            data = pd.read_csv(input_file_path, skiprows=[0, 1, 2, 3, 4, 5, 6, 8, 9, 10], encoding='gbk', low_memory=False)
        elif input_file_path.endswith(('.xlsx', '.xls')):
            data = pd.read_excel(input_file_path)
        else:
            raise ValueError("Unsupported file format. Please use CSV or Excel files.")

        # 区分数值列和离散文本列
        numeric_columns = data.select_dtypes(include=['number']).columns
        discrete_columns = data.select_dtypes(include=['object']).columns

        # 处理数值列，填充缺失值（先前向填充，再后向填充）
        numeric_data = data[numeric_columns].ffill().bfill()

        # 处理离散数据（编码），每一列都重新开始
        discrete_data = data[discrete_columns]
        for col in discrete_columns:
            le = LabelEncoder()
            discrete_data[col] = discrete_data[col].ffill().bfill()
            discrete_data[col] = le.fit_transform(discrete_data[col])

        # 合并处理后的数据
        processed_data = pd.concat([discrete_data, numeric_data], axis=1)

        # 转换为NumPy数组以便进行“归一化”（实际上是插值）
        x = processed_data.values
        idx = np.arange(x.shape[0])

        # 使用一维插值进行“归一化”（实际上是重新采样）
        f = interp1d(idx, x, axis=0, fill_value='extrapolate')
        idx_new = np.linspace(0, idx.max(), target_rows)
        x_new = f(idx_new)

        # 将处理后的数据转换回DataFrame
        df_resampled = pd.DataFrame(x_new, columns=processed_data.columns)

        # 保存到输出文件
        if output_file_path.endswith('.csv'):
            df_resampled.to_csv(output_file_path, index=False)
        elif output_file_path.endswith(('.xlsx', '.xls')):
            df_resampled.to_excel(output_file_path, index=False)
        else:
            raise ValueError("Unsupported file format for output. Please use CSV or Excel files.")

    def calculate_compared_params(file_path, param_name, compare_param_name, limited_param_name=None, limited_param_elements=None):
        try:
            if file_path.suffix.lower() == '.csv':
                data = pd.read_csv(file_path, skiprows=[0, 1, 2, 3, 4, 5, 6, 8, 9, 10], encoding='gbk', low_memory=False)
            else:
                data = pd.read_excel(file_path)
            if param_name in data.columns and compare_param_name in data.columns:
                # 先根据限定参数进行筛选（如果有限定参数相关配置）
                if limited_param_name and limited_param_elements:
                    target_type = type(data[limited_param_name].iloc[0])
                    converted_elements = convert_elements_type(limited_param_elements, target_type, file_path, limited_param_name)
                    if converted_elements is None:
                        return None
                    filtered_data = data[data[limited_param_name].isin(converted_elements)]
                else:
                    filtered_data = data
                if not filtered_data.empty:
                    compared_data = filtered_data[[param_name, compare_param_name]].dropna()
                    compared_data['差值'] = compared_data[param_name] - compared_data[compare_param_name]
                    return compared_data
        except Exception as e:
            print(f"处理文件 {file_path} 时出错: {e}")
            st.error(f"处理文件 {file_path} 时出现异常，请检查文件格式或数据内容是否正确，具体错误信息: {e}，文件路径：{file_path}")
        return None

    if st.button("归一化处理"):
        if input_folder_path and norm_output_folder:
            input_folder = Path(input_folder_path)
            norm_folder = Path(norm_output_folder)
            norm_folder.mkdir(parents=True, exist_ok=True)

            for filename in input_folder.glob('*.csv'):
                input_file_path = filename
                norm_file_path = norm_folder / f"{filename.stem}_norm.csv"

                try:
                    process_and_normalize(str(input_file_path), int(target_rows), str(norm_file_path))
                    print(f"文件 {input_file_path} 归一化处理完成，保存到 {norm_file_path}")
                except Exception as e:
                    print(f"文件 {input_file_path} 归一化处理出错: {e}")

            for filename in input_folder.glob('*.xlsx'):
                input_file_path = filename
                norm_file_path = norm_folder / f"{filename.stem}_norm.xlsx"

                try:
                    data = pd.read_excel(input_file_path)
                    process_and_normalize(str(input_file_path), int(target_rows), str(norm_file_path))
                    print(f"文件 {input_file_path} 归一化处理完成，保存到 {norm_file_path}")
                except Exception as e:
                    print(f"文件 {input_file_path} 归一化处理出错: {e}")

            st.success("文件夹中的所有文件归一化处理完成")
        else:
            st.error("请输入有效的文件夹路径")

    limited_param_name = st.text_input("限定参数名称（既定横坐标参数名称）") if use_limited_param else None
    limited_param_elements = st.text_input("限定参数元素（既定横坐标参数下的元素）") if use_limited_param else None

    if st.button("绘制参数图"):
        if norm_output_folder:
            norm_folder = Path(norm_output_folder)
            try:
                draw_graph(norm_output_folder, param_name, target_rows)
                st.success("参数图绘制完成")
            except Exception as e:
                st.error(f"绘制参数图时出错: {e}")
        else:
            st.error("请输入有效的归一化输出文件夹路径")

    if st.button("对比参数计算"):
        if input_folder_path and param_name and compare_param_name:
            input_folder = Path(input_folder_path)
            results = []
            for file_path in input_folder.glob('*.*'):
                if file_path.suffix.lower() in ['.csv', '.xlsx']:
                    result = calculate_compared_params(file_path, param_name, compare_param_name, limited_param_name, limited_param_elements)
                    if result is not None:
                        results.append(result)
                        st.subheader(f"文件：{file_path.name} 的对比参数计算结果")
                        st.dataframe(result)
                        fig = go.Figure(data=[go.Scatter(x=result.index, y=result['差值'], mode='lines', name='差值')])
                        fig.update_layout(title=f'文件：{file_path.name} 对比参数差值图', xaxis_title='Data Point Index', yaxis_title='差值')
                        st.plotly_chart(fig)
            if not results:
                st.error("未找到符合条件的数据，请检查输入参数是否正确")
        else:
            st.error("请输入有效的文件夹路径、查找参数名称以及对比参数名称")

# 只有当处于数据处理页面（sidebar == "数据处理"）时，才显示限定参数计算按钮
if sidebar == "数据处理":
    if st.button("限定参数计算"):
        if not input_folder_path:
            st.error("请输入有效的文件夹路径，当前路径为空，请重新输入。")
        elif not param_name:
            st.error("请输入有效的查找参数名称，当前名称为空，请重新输入。")
        elif not use_limited_param:
            st.error("请勾选限定参数计算选项，以启用此功能。")
        elif not limited_param_name:
            st.error("请输入有效的限定参数名称（既定横坐标参数名称），当前名称为空，请重新输入。")
        elif not limited_param_elements:
            st.error("请输入有效的限定参数元素（既定横坐标参数下的元素），当前元素为空，请重新输入。")
        else:
            calculation_results = calculate_with_limited_param(input_folder_path, param_name, limited_param_name, limited_param_elements)
            if calculation_results:
                result_df = pd.DataFrame(calculation_results)
                st.dataframe(result_df)
            else:
                st.error("未找到符合条件的数据，请检查输入参数是否正确，以下是一些排查建议：")
                st.write("1. 确认输入的文件夹路径下包含正确格式（CSV 或 Excel）的文件，且应用有读取权限。")
                st.write("2. 仔细核对限定参数名称、查找参数名称是否与文件中的列名准确匹配，注意大小写及拼写。")
                st.write("3. 检查限定参数元素的数据类型与对应列的数据类型是否一致，以及元素是否确实存在于限定参数列中。")
