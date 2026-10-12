from datetime import date, time, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from services.excel_storage import (
    ensure_workbook,
    read_table,
    add_project,
    add_task,
    add_session,
    list_projects,
    list_tasks,
)

st.set_page_config(page_title="Тайм-менеджмент", page_icon="⏱️", layout="wide")

st.title("Личный тайм-менеджмент")
st.caption("Локальное приложение. Данные хранятся в Excel на этом компьютере.")

ensure_workbook()

tab_home, tab_projects, tab_tasks, tab_sessions, tab_history = st.tabs(
    ["Главная", "Проекты", "Задачи", "Добавить сессию", "История"]
)

with tab_home:
    projects = read_table("Projects")
    tasks = read_table("Tasks")
    sessions = read_table("Sessions")

    c1, c2, c3 = st.columns(3)
    c1.metric("Проекты", len(projects))
    c2.metric("Задачи", len(tasks))
    c3.metric("Рабочие сессии", len(sessions))

    if not sessions.empty and "Длительность (ч)" in sessions.columns:
        total = pd.to_numeric(sessions["Длительность (ч)"], errors="coerce").fillna(0).sum()
    else:
        total = 0
    st.metric("Учтено времени", f"{total:.2f} ч")
    st.info("Это первая версия. Пока доступны проекты, задачи, ручной ввод сессий и просмотр истории. Таймер добавим следующим этапом.")

with tab_projects:
    st.subheader("Создать проект")
    with st.form("project_form", clear_on_submit=True):
        project_name = st.text_input("Название проекта")
        categories = ["Обучение", "Работа", "Личное", "Другое"]
        category = st.selectbox("Категория", categories)
        status = st.selectbox("Статус", ["Не начат", "В процессе", "На паузе", "Завершён"])
        target_hours = st.number_input("Планируемое время (ч)", min_value=0.0, step=1.0)
        description = st.text_area("Описание (необязательно)")
        submitted = st.form_submit_button("Создать проект", type="primary")
        if submitted:
            if not project_name.strip():
                st.error("Введи название проекта.")
            else:
                try:
                    project_id = add_project(project_name.strip(), category, status, target_hours, description.strip())
                    st.success(f"Проект создан: {project_id}")
                    st.rerun()
                except ValueError as exc:
                    st.error(str(exc))

    st.subheader("Существующие проекты")
    projects = read_table("Projects")
    if projects.empty:
        st.caption("Пока проектов нет.")
    else:
        st.dataframe(projects, use_container_width=True, hide_index=True)

with tab_tasks:
    st.subheader("Создать задачу")
    projects = list_projects()
    if not projects:
        st.warning("Сначала создай хотя бы один проект.")
    else:
        project_labels = {f'{p["Название"]} ({p["ID"]})': p["ID"] for p in projects}
        with st.form("task_form", clear_on_submit=True):
            selected_project_label = st.selectbox("Проект", list(project_labels.keys()))
            task_name = st.text_input("Название задачи")
            task_status = st.selectbox("Статус задачи", ["Не начата", "В процессе", "Готова", "Отложена"])
            task_description = st.text_area("Описание (необязательно)")
            submitted = st.form_submit_button("Создать задачу", type="primary")
            if submitted:
                if not task_name.strip():
                    st.error("Введи название задачи.")
                else:
                    try:
                        task_id = add_task(project_labels[selected_project_label], task_name.strip(), task_status, task_description.strip())
                        st.success(f"Задача создана: {task_id}")
                        st.rerun()
                    except ValueError as exc:
                        st.error(str(exc))

    st.subheader("Существующие задачи")
    tasks = read_table("Tasks")
    if tasks.empty:
        st.caption("Пока задач нет.")
    else:
        st.dataframe(tasks, use_container_width=True, hide_index=True)

with tab_sessions:
    st.subheader("Добавить рабочую сессию вручную")
    projects = list_projects()
    if not projects:
        st.warning("Сначала создай проект и задачу.")
    else:
        project_labels = {f'{p["Название"]} ({p["ID"]})': p["ID"] for p in projects}
        selected_project_label = st.selectbox("Проект для сессии", list(project_labels.keys()), key="session_project")
        selected_project_id = project_labels[selected_project_label]
        tasks = list_tasks(selected_project_id)
        if not tasks:
            st.warning("Для выбранного проекта пока нет задач. Создай задачу на вкладке «Задачи».")
        else:
            task_labels = {f'{t["Название"]} ({t["ID"]})': t["ID"] for t in tasks}
            with st.form("session_form", clear_on_submit=True):
                session_date = st.date_input("Дата", value=date.today())
                task_label = st.selectbox("Задача", list(task_labels.keys()))
                start_time = st.time_input("Время начала", value=time(9, 0))
                end_time = st.time_input("Время окончания", value=time(10, 0))
                notes = st.text_area("Что было сделано")
                submitted = st.form_submit_button("Сохранить сессию", type="primary")
                if submitted:
                    start_dt = datetime.combine(session_date, start_time)
                    end_dt = datetime.combine(session_date, end_time)
                    if end_dt <= start_dt:
                        st.error("Время окончания должно быть позже времени начала. Сессии через полночь пока не поддерживаются.")
                    else:
                        try:
                            session_id = add_session(
                                session_date.isoformat(),
                                selected_project_id,
                                task_labels[task_label],
                                start_dt.isoformat(timespec="minutes"),
                                end_dt.isoformat(timespec="minutes"),
                                notes.strip(),
                            )
                            st.success(f"Сессия сохранена: {session_id}")
                            st.rerun()
                        except ValueError as exc:
                            st.error(str(exc))

with tab_history:
    st.subheader("История рабочих сессий")
    sessions = read_table("Sessions")
    if sessions.empty:
        st.caption("Пока нет сохранённых сессий.")
    else:
        projects_df = read_table("Projects")
        tasks_df = read_table("Tasks")
        if not projects_df.empty:
            project_names = dict(zip(projects_df["ID"], projects_df["Название"]))
            sessions["Проект"] = sessions["ID проекта"].map(project_names).fillna(sessions["ID проекта"])
        if not tasks_df.empty:
            task_names = dict(zip(tasks_df["ID"], tasks_df["Название"]))
            sessions["Задача"] = sessions["ID задачи"].map(task_names).fillna(sessions["ID задачи"])
        st.dataframe(sessions, use_container_width=True, hide_index=True)
        st.download_button(
            "Скачать копию истории в CSV",
            sessions.to_csv(index=False).encode("utf-8-sig"),
            file_name="time_sessions.csv",
            mime="text/csv",
        )
