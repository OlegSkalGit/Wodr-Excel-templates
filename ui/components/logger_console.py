import streamlit as st
import subprocess
import os
import sys

def run_subprocess_and_stream(args):
    """Runs the python automation script and streams console outputs to Streamlit."""
    python_exe = sys.executable
    script_path = "_templates_machine_.py"
    
    cmd = [python_exe, script_path] + args
    
    log_area = st.empty()
    progress_bar = st.progress(0.0)
    
    st.info(f"🚀 Запущено команду: `python _templates_machine_.py {' '.join(f'\"{a}\"' for a in args)}`")
    
    st.session_state["last_operation_logs"] = []
    st.session_state["last_operation_status"] = "running"
    st.session_state["last_operation_cmd"] = f"python _templates_machine_.py {' '.join(f'\"{a}\"' for a in args)}"
    
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    
    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=False,
        env=env
    )
    
    logs = []
    
    while True:
        line_bytes = process.stdout.readline()
        if not line_bytes:
            break
            
        line = ""
        try:
            line = line_bytes.decode('utf-8')
        except UnicodeDecodeError:
            try:
                line = line_bytes.decode('cp1251')
            except Exception:
                line = line_bytes.decode('utf-8', errors='replace')
                
        cleaned_line = line.strip()
        if cleaned_line:
            logs.append(cleaned_line)
            st.session_state["last_operation_logs"] = logs
            log_area.code("\n".join(logs[-15:]))
            
            if "Створено" in cleaned_line or "Обробка" in cleaned_line:
                progress_bar.progress(0.5)
            elif "Готово!" in cleaned_line or "Успішно!" in cleaned_line:
                progress_bar.progress(1.0)
                
    process.wait()
    
    if process.returncode == 0:
        progress_bar.progress(1.0)
        st.session_state["last_operation_status"] = "success"
        st.session_state["trigger_balloons"] = True
        st.session_state["gen_completion_status"] = "success"
    else:
        st.session_state["last_operation_status"] = "error"
        st.session_state["trigger_balloons"] = False
        st.session_state["gen_completion_status"] = "error"
        
    return process.returncode, logs

def show_last_operation_logs():
    """Displays the persistent logs of the last executed operation from session state."""
    if "last_operation_logs" in st.session_state and st.session_state["last_operation_logs"]:
        status = st.session_state.get("last_operation_status", "")
        cmd = st.session_state.get("last_operation_cmd", "")
        
        if st.session_state.get("trigger_balloons"):
            st.balloons()
            st.session_state["trigger_balloons"] = False
            
        st.markdown("---")
        if status == "success":
            st.success("🎉 Завдання успішно виконано!")
            st.markdown("### 🟢 Результат останньої операції: Успішно")
        elif status == "error":
            st.error(f"❌ Помилка виконання завдання!")
            st.markdown("### 🔴 Результат останньої операції: Помилка")
        else:
            st.markdown("### 🟡 Результат останньої операції: Виконується")
            
        if cmd:
            st.caption(f"Команда: `{cmd}`")
            
        with st.expander("📋 Показати повний лог консолі", expanded=False):
            st.code("\n".join(st.session_state["last_operation_logs"]))
