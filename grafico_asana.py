import matplotlib
matplotlib.use("Agg")  # Backend seguro para gerar imagens sem janela gráfica

import asana
from asana.rest import ApiException
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.gridspec import GridSpec
import json
import math
from dotenv import load_dotenv
import os

load_dotenv()
PERSONAL_ACCESS_TOKEN = os.getenv("ASANA_PERSONAL_ACCESS_TOKEN")
WORKSPACE_ID = os.getenv("ASANA_WORKSPACE_ID")

def get_workspace_data(client):
    workspace_gid = WORKSPACE_ID
    opts = {'limit': 100}

    try:
        api_response = client.get_projects_for_workspace(workspace_gid, opts)
    except Exception as e:
        print(f"Erro ao consultar projetos, ERRO: {e}")
        api_response = []

    return api_response


def get_task_data(client, project_gid):
    opts = {
        'limit': 100,
        'opt_fields': "name, assignee.name, completed, assignee_status, due_on, created_at"
    }
    print(f"🔍 Buscando tarefas do Projeto ID: {project_gid}...")

    try:
        tasks = client.get_tasks_for_project(project_gid, opts)
    except Exception as e:
        print(f"Erro ao consultar tarefas, ERRO: {e}")
        tasks = []

    return tasks


def concat_information(project_id, project_name, tasks):
    data = []

    for task in tasks:
        assignee = task.get('assignee')
        assignee_name = assignee['name'] if assignee else 'Não Atribuído'
        status = "Concluída" if task.get('completed', False) else "Pendente"

        data.append({
            'Tarefa': task.get('name', 'Sem nome'),
            'Aluno': assignee_name,
            'Concluída': task.get('completed', False),
            'Status': status,
            'Due_Date': task.get('due_on'),
            'Created_At': task.get('created_at'),
        })

    print(f"✅ {len(data)} tarefas encontradas")
    return pd.DataFrame(data)


def analyze_progress(df):
    if df.empty:
        return pd.DataFrame()

    status = df.groupby('Aluno')['Concluída'].agg(
        total_tarefas='count',
        tarefas_concluidas='sum'
    ).reset_index()

    status['Porcentagem_Conclusao'] = (
        status['tarefas_concluidas'] / status['total_tarefas'] * 100
    ).round(1)

    status['tarefas_pendentes'] = (
        status['total_tarefas'] - status['tarefas_concluidas']
    )

    return status.sort_values('tarefas_concluidas', ascending=False)


__DASHBOARD_IMAGES = []
__DASHBOARD_PROJECT_NAMES = []


def __fig_to_rgb_array(fig):
    fig.canvas.draw()
    rgba = np.asarray(fig.canvas.buffer_rgba())
    rgb = rgba[:, :, :3].copy()
    return rgb


def __assemble_and_save_grid(images, titles, cols=2,
                             out_file="dashboard_unico_todos_projetos.png",
                             dpi=180):

    if not images:
        print("❌ Nenhuma imagem para montar o dashboard final.")
        return

    n = len(images)
    rows = math.ceil(n / cols)

    first_h, first_w = images[0].shape[0], images[0].shape[1]
    scale = 0.02

    fig = plt.figure(figsize=(max(12, cols * first_w * scale),
                              max(8, rows * first_h * scale)))

    gs = GridSpec(rows, cols, figure=fig, wspace=0.25, hspace=0.35)

    idx = 0
    for r in range(rows):
        for c in range(cols):
            ax = fig.add_subplot(gs[r, c])
            if idx < n:
                ax.imshow(images[idx])
                ax.set_title(titles[idx], fontsize=10, fontweight="bold")
            ax.axis("off")
            idx += 1

    plt.subplots_adjust(left=0.02, right=0.98, top=0.94, bottom=0.04)

    fig.savefig(out_file, dpi=dpi, bbox_inches="tight")
    plt.close(fig)

    print(f"\n📊 Dashboard unificado gerado apenas uma vez: {out_file}")


def create_powerbi_dashboard(df_tasks, df_status, project_name):

    if df_tasks.empty:
        print("Nenhum dado para gerar dashboard.")
        return

    fig = plt.figure(figsize=(20, 12))
    fig.suptitle(f'DASHBOARD DE PRODUTIVIDADE - {project_name}',
                 fontsize=16, fontweight='bold', y=0.98)

    gs = GridSpec(3, 4, figure=fig, hspace=0.4, wspace=0.3)

    # -------- GRÁFICO 1 --------
    ax1 = fig.add_subplot(gs[0, :2])

    df_plot = df_status[df_status['Aluno'] != 'Não Atribuído'] if not df_status.empty else pd.DataFrame()

    if not df_plot.empty:
        df_plot = df_plot.sort_values("tarefas_concluidas", ascending=True)
        colors = plt.cm.Blues(np.linspace(0.6, 0.9, len(df_plot)))
        bars = ax1.barh(df_plot['Aluno'], df_plot['tarefas_concluidas'], color=colors)

        for bar, row in zip(bars, df_plot.itertuples()):
            ax1.text(bar.get_width() + 0.2, bar.get_y() + bar.get_height()/2,
                     f'{int(row.tarefas_concluidas)}', va='center', ha='left')

        ax1.set_title("Tarefas Concluídas por Aluno")
        ax1.grid(axis="x", linestyle="--", alpha=0.3)
    else:
        ax1.text(0.5, 0.5, "Sem dados", ha="center", va="center")

    # -------- GRÁFICO 2 --------
    ax2 = fig.add_subplot(gs[0, 2:])
    status_counts = df_tasks['Status'].value_counts() if not df_tasks.empty else pd.Series([])

    if not status_counts.empty:
        ax2.pie(status_counts.values, labels=status_counts.index, autopct="%1.0f%%")
    else:
        ax2.text(0.5, 0.5, "Sem dados", ha="center", va="center")
    ax2.set_title("Distribuição de Status")

    # -------- GRÁFICO 3 --------
    ax3 = fig.add_subplot(gs[1, :2])
    if not df_plot.empty:
        top = df_plot.tail(5)
        bars = ax3.bar(top['Aluno'], top['tarefas_concluidas'])
        for bar in bars:
            ax3.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
                     str(int(bar.get_height())), ha='center')
        ax3.set_title("Top 5 Alunos")
    else:
        ax3.text(0.5, 0.5, "Sem dados", ha="center", va="center")

    # -------- GRÁFICO 4 --------
    ax4 = fig.add_subplot(gs[1, 2:])
    ax4.axis("off")

    total_tarefas = len(df_tasks)
    total_concluidas = df_tasks["Concluída"].sum()
    total_pendentes = total_tarefas - total_concluidas

    text = f"""
    MÉTRICAS DO PROJETO

    Total de tarefas: {total_tarefas}
    Concluídas: {total_concluidas}
    Pendentes: {total_pendentes}
    """
    ax4.text(0.05, 0.95, text, va="top", fontfamily="monospace", fontsize=12)

    # -------- GRÁFICO 5 --------
    ax5 = fig.add_subplot(gs[2, :])
    if not df_plot.empty:
        df_comp = df_plot.sort_values("tarefas_concluidas", ascending=False)
        ax5.bar(df_comp['Aluno'], df_comp['tarefas_concluidas'], label='Concluídas')
        ax5.bar(df_comp['Aluno'], df_comp['tarefas_pendentes'],
                bottom=df_comp['tarefas_concluidas'], label='Pendentes')
        ax5.legend()
        ax5.set_title("Concluídas vs Pendentes")
    else:
        ax5.text(0.5, 0.5, "Sem dados", ha="center", va="center")

    plt.subplots_adjust(left=0.03, right=0.97, top=0.95, bottom=0.05)

    # ---- CONVERTE FIGURA EM IMAGEM ----
    img = __fig_to_rgb_array(fig)
    __DASHBOARD_IMAGES.append(img)
    __DASHBOARD_PROJECT_NAMES.append(project_name)

    plt.close(fig)


# ==========================================
#        MAIN
# ==========================================

if __name__ == "__main__":
    try:
        configuration = asana.Configuration()
        configuration.access_token = PERSONAL_ACCESS_TOKEN
        api_client = asana.ApiClient(configuration)

        projects_api = asana.ProjectsApi(api_client)
        tasks_api = asana.TasksApi(api_client)

        print("🚀 INICIANDO ANÁLISE DE PRODUTIVIDADE")

        projects = get_workspace_data(projects_api)

        for proj in projects:
            project_id = proj.get('gid')
            project_name = proj.get('name')

            tasks = get_task_data(tasks_api, project_id)
            df_tasks = concat_information(project_id, project_name, tasks)

            if df_tasks.empty:
                print(f"❌ Nenhuma tarefa encontrada para {project_name}")
                continue

            df_status = analyze_progress(df_tasks)

            print("\n🎨 Gerando dashboard individual...")
            create_powerbi_dashboard(df_tasks, df_status, project_name)

        # 🔥 AGORA SIM: gerar a imagem final UMA ÚNICA VEZ
        print("\n🖼 Montando dashboard final consolidado...")
        __assemble_and_save_grid(__DASHBOARD_IMAGES, __DASHBOARD_PROJECT_NAMES, cols=2)

    except Exception as e:
        print(f"❌ ERRO: {e}")
        import traceback
        traceback.print_exc()
