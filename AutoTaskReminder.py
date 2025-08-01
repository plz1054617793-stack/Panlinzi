import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime, time, timedelta
import threading
import pandas as pd
import time as tm
import uuid
from copy import deepcopy
from queue import Queue

# 全局队列用于线程间通信
update_queue = Queue()

# 线程安全的会话状态访问工具类
class ThreadSafeState:
    """用于线程安全地获取和同步session_state数据"""
    @staticmethod
    def get_safe_copy():
        """获取session_state关键数据的深拷贝，避免线程上下文问题"""
        try:
            # 确保所有必要的键都已初始化
            required_keys = ["tasks", "scheduled_times", "smtp_config", "logged_in"]
            for key in required_keys:
                if key not in st.session_state:
                    st.session_state[key] = [] if key in ["tasks", "scheduled_times"] else {} if key == "smtp_config" else False
            
            state = {
                "tasks": deepcopy(st.session_state.get('tasks', [])),
                "scheduled_times": deepcopy(st.session_state.get('scheduled_times', [])),
                "smtp_config": deepcopy(st.session_state.get('smtp_config', {})),
                "logged_in": st.session_state.get('logged_in', False)
            }
            return state
        except Exception as e:
            st.error(f"获取会话状态失败: {str(e)}")
            return None

# 检查Streamlit版本，使用兼容的页面刷新方法
def refresh_page():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

# 初始化session_state
def init_session_state():
    if 'initialized' not in st.session_state:
        st.session_state.initialized = True
        st.session_state.logged_in = False
        st.session_state.tasks = []  # 存储任务列表
        st.session_state.smtp_config = {}  # 邮箱配置
        st.session_state.scheduled_times = []  # 存储定时时间（time对象）
        st.session_state.scheduler_running = False  # 定时任务运行状态
        st.session_state.file_path = "tasks.csv"  # 任务保存路径
        st.session_state.file_path_input = st.session_state.file_path
        st.session_state.scheduler_status = "未启动"  # 定时任务状态文本

init_session_state()

# 处理队列中的更新任务（在主线程中执行）
def process_update_queue():
    while not update_queue.empty():
        task_id = update_queue.get()
        # 在主线程中更新任务状态
        for task in st.session_state.tasks:
            if task["事项ID"] == task_id:
                task["reminded"] = True
                break
        update_queue.task_done()

# 页面布局
st.title("任务管控系统")

# 登录模块
with st.expander("邮箱登录", expanded=not st.session_state.logged_in):
    col1, col2 = st.columns(2)
    with col1:
        email = st.text_input("网易邮箱地址", key="login_email")
    with col2:
        password = st.text_input("授权码", type="password", key="login_password")
    
    if st.button("登录"):
        if not email or not password:
            st.error("请输入邮箱地址和授权码")
        else:
            try:
                # 测试SMTP连接
                with smtplib.SMTP_SSL('smtp.163.com', 465) as server:
                    server.login(email, password)
                st.session_state.smtp_config = {
                    'username': email,
                    'password': password
                }
                st.session_state.logged_in = True
                st.success("登录成功！")
                refresh_page()
            except Exception as e:
                st.error(f"登录失败: {str(e)}")

# 文件保存路径设置
st.text_input("本地文件保存地址", key="file_path_input", value=st.session_state.file_path)
if st.session_state.file_path_input != st.session_state.file_path:
    st.session_state.file_path = st.session_state.file_path_input

# 显示定时任务状态
st.markdown(f"### 定时任务状态: {st.session_state.scheduler_status}")

# 处理队列更新（在主线程中）
process_update_queue()

# 任务表格（带状态列）
if st.session_state.tasks:
    df = pd.DataFrame(st.session_state.tasks)
    df['到期状态'] = df['到期日期'].apply(
        lambda x: '已到期' if x < datetime.now().date() else '未到期'
    )
    df['提醒状态'] = df.apply(
        lambda row: '已提醒' if row.get('reminded', False) else '待提醒',
        axis=1
    )
    st.table(df)
else:
    st.info("暂无任务，请添加新任务")

# 任务数据本地存储功能
st.markdown("### 任务管理")
col_save, col_load, col_refresh = st.columns(3)
with col_save:
    if st.button("保存所有任务到本地"):
        try:
            if st.session_state.tasks:
                df = pd.DataFrame(st.session_state.tasks)
                df['到期日期'] = df['到期日期'].astype(str)  # 日期转为字符串存储
                df.to_csv(st.session_state.file_path, index=False)
                st.success(f"任务已保存到 {st.session_state.file_path}")
            else:
                st.warning("没有任务可保存")
        except Exception as e:
            st.error(f"保存失败: {str(e)}")

with col_load:
    if st.button("加载历史任务"):
        try:
            df = pd.read_csv(st.session_state.file_path)
            df['到期日期'] = pd.to_datetime(df['到期日期']).dt.date  # 还原日期格式
            st.session_state.tasks = df.to_dict('records')
            st.success(f"已加载 {len(df)} 条历史任务")
        except Exception as e:
            st.error(f"加载失败: {str(e)}")

with col_refresh:
    if st.button("刷新任务列表"):
        refresh_page()

# 添加任务模块
st.markdown("### 添加新任务")
with st.form("添加任务"):
    col1, col2 = st.columns(2)
    with col1:
        title = st.text_input("邮件标题")
    with col2:
        recipients = st.text_area("收件人列表（分号分隔）")
    
    content = st.text_area("邮件内容")
    due_date = st.date_input("到期日期")
    
    if st.form_submit_button("添加任务"):
        if not title or not recipients or not content or due_date is None:
            st.error("请填写所有必填字段")
        else:
            new_task = {
                "事项ID": str(uuid.uuid4()),
                "邮件标题": title,
                "邮件内容": content,
                "自定义收件人列表": recipients,
                "到期日期": due_date,
                "reminded": False  # 仅记录是否提醒过，不限制发送次数
            }
            st.session_state.tasks.append(new_task)
            st.success("任务添加成功！")
            refresh_page()

# 邮件发送核心功能（线程安全版）
def send_email(task, smtp_config):
    """独立于session_state的邮件发送函数，通过参数传入配置，自动添加到期日期"""
    try:
        if not smtp_config or not smtp_config.get('username') or not smtp_config.get('password'):
            print("SMTP配置缺失，无法发送邮件")
            return False
        
        # 在原始内容后添加到期日期信息
        due_date_str = task['到期日期'].strftime("%Y年%m月%d日")
        full_content = f"{task['邮件内容']}\n\n【任务到期日期：{due_date_str}】"
        
        server = smtplib.SMTP_SSL('smtp.163.com', 465)
        server.login(smtp_config['username'], smtp_config['password'])
        
        msg = MIMEText(full_content, 'plain', 'utf-8')
        msg['Subject'] = Header(task['邮件标题'], 'utf-8')
        msg['From'] = smtp_config['username']
        msg['To'] = ";".join(task['自定义收件人列表'].split(';'))
        
        server.sendmail(
            smtp_config['username'],
            task['自定义收件人列表'].split(';'),
            msg.as_string()
        )
        server.quit()
        print(f"任务 {task['事项ID']} 邮件发送成功！")
        return True
    except Exception as e:
        print(f"邮件发送失败: {str(e)}")
        return False

# 立即发送未到期任务（支持多次发送）
st.markdown("### 邮件发送")
if st.button("立即发送未到期任务") and st.session_state.logged_in:
    if not st.session_state.tasks:
        st.warning("没有任务可发送")
    else:
        success_count = 0
        current_date = datetime.now().date()
        # 筛选未到期任务（到期日期 >= 当前日期）
        for task in st.session_state.tasks:
            if task['到期日期'] >= current_date:
                if send_email(task, st.session_state.smtp_config):
                    task['reminded'] = True  # 更新为已提醒，但仍可再次发送
                    success_count += 1
        st.success(f"已向未到期任务发送 {success_count} 封邮件")
        refresh_page()

# 多定时发送设置（1分钟间隔滚轮选择）
st.markdown("### 定时发送设置")
st.write("添加定时发送时间（时:分，秒固定为00，支持滚轮选择）：")

# 时间选择（小时和分钟独立选择，秒固定为00）
col_hour, col_minute = st.columns(2)
with col_hour:
    selected_hour = st.slider("选择小时", 0, 23, 8, step=1)
with col_minute:
    selected_minute = st.slider("选择分钟", 0, 59, 0, step=1)  # 分钟间隔1分钟

# 生成完整时间对象（秒固定为00）
new_schedule_time = time(selected_hour, selected_minute, 0)
time_str = new_schedule_time.strftime("%H:%M:%S")

# 添加定时时间按钮
col_add, _ = st.columns([1, 3])
with col_add:
    if st.button("添加定时时间"):
        # 检查是否已存在相同时间
        exists = any(t.strftime("%H:%M:%S") == time_str for t in st.session_state.scheduled_times)
        if exists:
            st.warning("该时间已存在！")
        else:
            st.session_state.scheduled_times.append(new_schedule_time)
            st.success(f"已添加定时时间：{time_str}")
            refresh_page()

# 显示已添加的定时时间并支持删除
st.write("已设置的定时时间：")
if st.session_state.scheduled_times:
    # 按时间排序显示
    sorted_times = sorted(st.session_state.scheduled_times, key=lambda t: (t.hour, t.minute))
    for idx, t in enumerate(sorted_times):
        time_str = t.strftime("%H:%M:%S")
        col_time, col_del = st.columns([3, 1])
        with col_time:
            st.write(f"- {time_str}")
        with col_del:
            if st.button("删除", key=f"del_time_{idx}"):
                # 从原始列表中删除
                st.session_state.scheduled_times = [
                    time for time in st.session_state.scheduled_times 
                    if time.strftime("%H:%M:%S") != time_str
                ]
                st.success(f"已删除定时时间：{time_str}")
                refresh_page()
else:
    st.info("暂无定时时间，请添加")

# 定时发送核心线程（解决上下文缺失问题）
def scheduled_send_handler(initial_schedules):
    """定时发送线程，通过队列与主线程通信，只发送未到期任务"""
    # 使用传入的初始定时数据，避免直接访问session_state
    st.session_state.scheduler_status = "运行中"
    last_sync_time = datetime.now()
    local_tasks = []  # 线程内缓存的任务
    local_schedules = initial_schedules  # 使用传入的初始定时数据
    local_smtp_config = {}  # 线程内缓存的SMTP配置
    
    # 记录上次执行时间，避免重复执行
    last_executed_times = {t.strftime("%H:%M:%S"): None for t in local_schedules}
    
    while st.session_state.get('scheduler_running', False):
        try:
            # 每30秒同步一次最新数据（避免频繁访问session_state）
            if (datetime.now() - last_sync_time) > timedelta(seconds=30):
                safe_state = ThreadSafeState.get_safe_copy()
                if safe_state:
                    local_tasks = safe_state["tasks"]
                    local_schedules = safe_state["scheduled_times"]
                    local_smtp_config = safe_state["smtp_config"]
                    last_sync_time = datetime.now()
                    
                    # 更新上次执行时间记录
                    for t in local_schedules:
                        time_str = t.strftime("%H:%M:%S")
                        if time_str not in last_executed_times:
                            last_executed_times[time_str] = None
            
            # 检查是否需要发送（当前时间匹配任何定时时间）
            now = datetime.now()
            current_time = now.time()  # 当前时间（时分秒）
            current_date = now.date()  # 当前日期
            
            # 遍历所有定时时间，检查是否匹配
            for schedule_time in local_schedules:
                time_str = schedule_time.strftime("%H:%M:%S")
                last_executed = last_executed_times.get(time_str)
                
                # 如果当前时间等于定时时间且今天尚未执行
                if (current_time.hour == schedule_time.hour and 
                    current_time.minute == schedule_time.minute and 
                    (last_executed is None or last_executed.date() < current_date)):
                    
                    print(f"[{now}] 执行定时任务: {time_str}")
                    st.session_state.scheduler_status = f"正在执行 {time_str} 的定时任务"
                    
                    # 只发送未到期任务
                    if local_tasks and local_smtp_config:
                        for task in local_tasks:
                            if task['到期日期'] >= current_date:
                                # 发送成功后通过队列通知主线程更新状态
                                if send_email(task, local_smtp_config):
                                    update_queue.put(task["事项ID"])
                    
                    # 更新最后执行时间
                    last_executed_times[time_str] = now
                    st.session_state.scheduler_status = "运行中"
            
            # 每分钟检查一次
            tm.sleep(60)
            
        except Exception as e:
            print(f"定时任务执行错误: {str(e)}")
            tm.sleep(60)  # 出错后等待1分钟再重试
    
    st.session_state.scheduler_status = "已停止"
    print("定时任务线程已终止")

# 启动/停止定时任务线程
col_start, col_stop = st.columns(2)
with col_start:
    if not st.session_state.scheduler_running:
        if st.button("启动所有定时任务") and st.session_state.logged_in:
            if not st.session_state.scheduled_times:
                st.error("请先添加至少一个定时时间！")
            else:
                # 启动线程时传递初始的定时数据，避免线程直接访问session_state
                initial_schedules = deepcopy(st.session_state.scheduled_times)
                st.session_state.scheduler_running = True
                threading.Thread(target=scheduled_send_handler, args=(initial_schedules,), daemon=True).start()
                st.success("定时任务已启动，将在设置的时间点发送提醒")
with col_stop:
    if st.session_state.scheduler_running:
        if st.button("停止所有定时任务"):
            st.session_state.scheduler_running = False
            st.success("定时任务已停止（当前线程将在60秒内终止）")
