import datetime
import os
import pandas as pd
import streamlit as st

# ----------------- 1. 文件与数据初始化 -----------------
DATA_FILE = "system_data.xlsx"


def init_data():
  if not os.path.exists(DATA_FILE):
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
      # 员工表
      pd.DataFrame(
          columns=["员工姓名", "员工工号", "备注"],
          data=[["张三", "E001", "班长"], ["李四", "E002", "操作工"]],
      ).to_excel(writer, sheet_name="员工名单", index=False)
      # 机台项目表
      pd.DataFrame(
          columns=["机台号", "项目名称", "目标产能（件/台/班）", "备注"],
          data=[["M01", "产品A-冲压", 1000], ["M02", "产品B-组装", 800]],
      ).to_excel(writer, sheet_name="机台项目", index=False)
      # 生产记录表
      pd.DataFrame(columns=[
          "生产日期",
          "录入时间",
          "班次",
          "员工姓名",
          "员工工号",
          "机台号",
          "项目名称",
          "良品数量",
          "不良品数量",
          "总产量",
          "目标产能",
          "目标达成率",
          "良品率",
          "备注",
      ]).to_excel(writer, sheet_name="生产录入", index=False)
      # 账号密码表
      pd.DataFrame(
          columns=["账号", "密码", "角色", "姓名"],
          data=[
              ["admin", "123456", "管理员", "系统管理员"],
              ["user1", "123456", "普通用户", "张三"],
          ],
      ).to_excel(writer, sheet_name="账号密码", index=False)


init_data()


# 加载数据函数
def load_data():
  return {
      "员工名单": pd.read_excel(DATA_FILE, sheet_name="员工名单"),
      "机台项目": pd.read_excel(DATA_FILE, sheet_name="机台项目"),
      "生产录入": pd.read_excel(DATA_FILE, sheet_name="生产录入"),
      "账号密码": pd.read_excel(DATA_FILE, sheet_name="账号密码"),
  }


# 保存数据函数
def save_data(dfs):
  with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
    for sheet_name, df in dfs.items():
      df.to_excel(writer, sheet_name=sheet_name, index=False)


# ----------------- 2. 界面布局与登录系统 -----------------
st.set_page_config(
    page_title="生产报表填报系统", page_icon="📊", layout="wide"
)

# 初始化 Session 状态
if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
  st.session_state.user = None
  st.session_state.role = None

dfs = load_data()

# 登录侧边栏
st.sidebar.title("🔐 系统登录")
if not st.session_state.logged_in:
  username = st.sidebar.text_input("账号")
  password = st.sidebar.text_input("密码", type="password")
  if st.sidebar.button("登录", type="primary"):
    users_df = dfs["账号密码"]
    match = users_df[
        (users_df["账号"] == username) & (users_df["密码"] == password)
    ]
    if not match.empty:
      st.session_state.logged_in = True
      st.session_state.user = username
      st.session_state.role = match.iloc[0]["角色"]
      st.session_state.real_name = match.iloc[0]["姓名"]
      st.sidebar.success("登录成功！")
      st.rerun()
    else:
      st.sidebar.error("账号或密码错误！")
  st.stop()
else:
  st.sidebar.write(
      f"当前用户: **{st.session_state.real_name}** ({st.session_state.role})"
  )
  if st.sidebar.button("退出登录"):
    st.session_state.logged_in = False
    st.session_state.user = None
    st.session_state.role = None
    st.rerun()

# ----------------- 3. 主页面逻辑 -----------------
st.title("📊 现代化智能生产报表填报系统")

menu = ["数据填报", "生产看板与日报"]
if st.session_state.role == "管理员":
  menu.append("后台数据维护")

choice = st.sidebar.selectbox("系统功能导航", menu)

# --- 功能模块 1：数据填报 ---
if choice == "数据填报":
  st.subheader("📝 生产数据在线填报")

  with st.form("production_form"):
    col1, col2, col3 = st.columns(3)

    with col1:
      prod_date = st.date_input(
          "生产日期", value=datetime.date.today()
      )  # 默认今天

    with col2:
      # 自动抓取本地/当前时间
      now_time = datetime.datetime.now()
      time_str = st.text_input(
          "录入时间 (自动抓取)", value=now_time.strftime("%H:%M:%S")
      )

    with col3:
      # 班次自动识别逻辑：08:00 - 20:00 为白班，其余为夜班
      current_hour = now_time.hour
      auto_shift = "白班" if 8 <= current_hour < 20 else "夜班"
      shift = st.selectbox("班次 (自动识别)", ["白班", "夜班"], index=0 if auto_shift=="白班" else 1)

    st.markdown("---")

    col4, col5 = st.columns(2)
    with col4:
      # 工号联动姓名：根据工号带出姓名
      emp_df = dfs["员工名单"]
      emp_ids = emp_df["员工工号"].tolist()
      selected_id = st.selectbox("员工工号", emp_ids)

      # 自动查找姓名
      matched_name = (
          emp_df[emp_df["员工工号"] == selected_id]["员工姓名"].values[0]
          if selected_id in emp_ids
          else ""
      )
      st.text_input("员工姓名 (根据工号自动带出)", value=matched_name, disabled=True)

    with col5:
      mach_df = dfs["机台项目"]
      mach_list = mach_df["机台号"].tolist() if not mach_df.empty else []
      selected_mach = st.selectbox("机台号", mach_list)

      # 自动匹配机台对应的项目名称与目标产能
      matched_project = ""
      target_capacity = 0
      if selected_mach and not mach_df.empty:
        row = mach_df[mach_df["机台号"] == selected_mach]
        if not row.empty:
          matched_project = row.iloc[0]["项目名称"]
          target_capacity = row.iloc[0]["目标产能（件/台/班）"]

      st.text_input(
          "项目名称 (自动匹配)", value=matched_project, disabled=True
      )
      st.number_input(
          "目标产能（件/班）", value=int(target_capacity), disabled=True
      )

    st.markdown("---")
    col6, col7, col8 = st.columns(3)
    with col6:
      good_count = st.number_input("良品数量", min_value=0, value=0, step=1)
    with col7:
      bad_count = st.number_input("不良品数量", min_value=0, value=0, step=1)
    with col8:
      remark = st.text_input("备注说明")

    submitted = st.form_submit_button("🚀 提交报表", type="primary")

    if submitted:
      total_count = good_count + bad_count
      # 核心要求1：目标产能达成率计算基于良品数
      # 达成率 = 良品数 / 目标产能
      达成率 = (
          (good_count / target_capacity) * 100 if target_capacity > 0 else 0
      )
      良品率 = (
          (good_count / total_count) * 100 if total_count > 0 else 0
      )

      new_row = {
          "生产日期": str(prod_date),
          "录入时间": time_str,
          "班次": shift,
          "员工姓名": matched_name,
          "员工工号": selected_id,
          "机台号": selected_mach,
          "项目名称": matched_project,
          "良品数量": good_count,
          "不良品数量": bad_count,
          "总产量": total_count,
          "目标产能": target_capacity,
          "目标达成率": f"{达成率:.2f}%",
          "良品率": f"{良品率:.2f}%",
          "备注": remark,
      }

      dfs["生产录入"] = pd.concat(
          [dfs["生产录入"], pd.DataFrame([new_row])], ignore_index=True
      )
      save_data(dfs)
      st.success(
          f"🎉 提交成功！本次良品达成率为: **{达成率:.2f}%**（基于良品数计算）"
      )

# --- 功能模块 2：生产看板与日报 ---
elif choice == "生产看板与日报":
  st.subheader("📈 生产看板与数据汇总")
  df_records = dfs["生产录入"]

  if df_records.empty:
    st.info("暂无生产填报数据，请先在填报页面录入数据。")
  else:
    col_f1, col_f2 = st.columns(2)
    with col_f1:
      selected_date = st.selectbox(
          "筛选日期", df_records["生产日期"].unique()
      )
    with col_f2:
      selected_shift = st.selectbox(
          "筛选班次", ["全部"] + list(df_records["班次"].unique())
      )

    # 筛选逻辑
    filtered_df = df_records[df_records["生产日期"] == selected_date]
    if selected_shift != "全部":
      filtered_df = filtered_df[filtered_df["班次"] == selected_shift]

    st.dataframe(filtered_df, use_container_width=True)

    # 下载报表
    csv = filtered_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "📥 导出当前筛选报表 (CSV)",
        csv,
        f"生产报表_{selected_date}.csv",
        "text/csv",
    )

# --- 功能模块 3：后台数据维护（仅管理员） ---
elif choice == "后台数据维护" and st.session_state.role == "管理员":
  st.subheader("⚙️ 后台数据维护中心")

  tab1, tab2, tab3 = st.tabs(["员工名单管理", "机台项目产能管理", "账号密码管理"])

  with tab1:
    st.write("维护员工信息（工号与姓名）")
    edited_emp = st.data_editor(dfs["员工名单"], num_rows="dynamic")
    if st.button("保存员工名单"):
      dfs["员工名单"] = edited_emp
      save_data(dfs)
      st.success("员工名单保存成功！")

  with tab2:
    st.write("维护机台、对应项目及标准目标产能")
    edited_mach = st.data_editor(dfs["机台项目"], num_rows="dynamic")
    if st.button("保存机台项目"):
      dfs["机台项目"] = edited_mach
      save_data(dfs)
      st.success("机台项目保存成功！")

  with tab3:
    st.write("管理系统后台账号与普通用户权限")
    edited_user = st.data_editor(dfs["账号密码"], num_rows="dynamic")
    if st.button("保存账号密码"):
      dfs["账号密码"] = edited_user
      save_data(dfs)
      st.success("账号密码保存成功！")
