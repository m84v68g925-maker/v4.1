def init_data():
  if not os.path.exists(DATA_FILE):
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
      pd.DataFrame(
          columns=["员工姓名", "员工工号", "备注"],
          data=[["张三", "E001", "班长"], ["李四", "E002", "操作工"]],
      ).to_excel(writer, sheet_name="员工名单", index=False)
      pd.DataFrame(
          columns=["机台号", "项目名称", "目标产能（件/台/班）", "备注"],
          data=[
              ["M01", "产品A-冲压", 1000, ""],
              ["M02", "产品B-组装", 800, ""],
          ],
      ).to_excel(writer, sheet_name="机台项目", index=False)
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
      pd.DataFrame(
          columns=["账号", "密码", "角色", "姓名"],
          data=[
              ["admin", "123456", "管理员", "系统管理员"],
              ["user1", "123456", "普通用户", "张三"],
          ],
      ).to_excel(writer, sheet_name="账号密码", index=False)
