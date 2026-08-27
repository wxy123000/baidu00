import json
import shutil
from pathlib import Path

from streamlit.testing.v1 import AppTest


ROOT = Path(__file__).resolve().parent
SANDBOX = ROOT / ".mvp_test_runtime"


def fresh_app():
    SANDBOX.mkdir(exist_ok=True)
    shutil.copy2(ROOT / "app.py", SANDBOX / "app.py")
    db = SANDBOX / "mvp_metrics.db"
    if db.exists():
        db.unlink()
    at = AppTest.from_file(str(SANDBOX / "app.py"), default_timeout=30)
    at.run()
    return at


def labels(items):
    return [getattr(x, "label", "") for x in items]


def no_exception(at):
    return len(at.exception) == 0


results = []


def record(case_id, module, action, passed, detail):
    results.append({
        "case_id": case_id,
        "module": module,
        "action": action,
        "passed": bool(passed),
        "detail": detail,
    })


at = fresh_app()
record("FT-001", "启动与首页", "首次启动并显示首页", no_exception(at) and "快速输入" in labels(at.text_area),
       f"exceptions={len(at.exception)}; textareas={labels(at.text_area)}")

# Empty-input validation.
btn = next(x for x in at.button if x.label == "📋 使用文本")
btn.click().run()
record("FT-002", "首页", "未输入文本时点击使用文本", no_exception(at) and any("请先输入文本" in x.value for x in at.warning),
       f"warnings={[x.value for x in at.warning]}")

# Text-based core flow.
at = fresh_app()
ta = next(x for x in at.text_area if x.label == "快速输入")
ta.set_value("项目周会讨论了测试计划、负责人和本周交付时间，请整理会议纪要。")
next(x for x in at.button if x.label == "📋 使用文本").click().run()
record("FT-003", "资料输入", "首页粘贴文本进入功能选择", no_exception(at) and at.session_state["page"] == "选择功能",
       f"page={at.session_state['page']}; source_len={len(at.session_state['source_text'])}")

# Recommendation and four content-generation functions in local mode.
features = ["会议纪要", "文档问答", "任务拆解", "周报生成"]
for idx, feature in enumerate(features, start=4):
    at.session_state["selected_feature"] = feature
    at.session_state["page"] = "选择功能"
    at.run()
    if feature == "文档问答" and at.text_input:
        q = next((x for x in at.text_input if "资料询问" in x.label), None)
        if q:
            q.set_value("这份资料的主要任务是什么？")
    start = next(x for x in at.button if x.label == "开始生成 ➜")
    start.click().run()
    ok = no_exception(at) and at.session_state["page"] == "生成结果" and bool(at.session_state["result"].strip())
    record(f"FT-{idx:03d}", feature, "本地演示模式生成结果", ok,
           f"page={at.session_state['page']}; provider={at.session_state['ai_provider']}; result_len={len(at.session_state['result'])}")
    at.session_state["page"] = "选择功能"
    at.run()

# Editing.
at.session_state["page"] = "生成结果"
at.run()
next(x for x in at.button if x.label == "✎ 编辑确认").click().run()
edit = next(x for x in at.text_area if "AI 生成结果" in x.label)
edit.set_value(at.session_state["result"] + "\n\n人工确认：内容已核对。")
next(x for x in at.button if x.label == "✓ 保存修改").click().run()
record("FT-008", "编辑确认", "编辑并保存AI生成结果", no_exception(at) and "人工确认" in at.session_state["result"],
       f"saved={'人工确认' in at.session_state['result']}; success={[x.value for x in at.success]}")

# Export page creates all three formats without exceptions.
at.session_state["page"] = "导出保存"
at.run()
format_results = []
for fmt in ["Word (.docx)", "PDF (.pdf)", "TXT (.txt)"]:
    radio = next(x for x in at.radio if x.label == "导出格式")
    radio.set_value(fmt).run()
    format_results.append((fmt, no_exception(at), len(at.get("download_button"))))
record("FT-009", "导出保存", "生成Word、PDF、TXT下载内容", all(x[1] and x[2] > 0 for x in format_results),
       str(format_results))

# Save to My Files and then open list.
next(x for x in at.button if x.label == "▣ 保存到我的文件").click().run()
saved_ok = len(at.session_state["saved_files"]) >= 1
at.session_state["page"] = "我的文件"
at.run()
record("FT-010", "我的文件", "保存结果并在我的文件中显示", no_exception(at) and saved_ok and len(at.get("download_button")) >= 1,
       f"saved_files={len(at.session_state['saved_files'])}; downloads={len(at.get('download_button'))}")

# History created automatically and supports filtering view.
at.session_state["page"] = "历史记录"
at.run()
record("FT-011", "历史记录", "显示已生成历史记录", no_exception(at) and len(at.session_state["history"]) >= 4,
       f"history_count={len(at.session_state['history'])}")

# AI chat local fallback.
at.session_state["page"] = "AI 咨询"
at.run()
record("FT-012", "AI咨询", "咨询页在本地模式正常加载", no_exception(at) and len(at.chat_input) == 1,
       f"chat_inputs={len(at.chat_input)}; messages={len(at.session_state['chat_messages'])}")

# Automation local mode.
at.session_state["page"] = "任务自动化"
at.run()
goal = next(x for x in at.text_area if x.label == "项目目标或工作要求")
goal.set_value("两周内完成MVP测试、修复问题并形成交付报告。")
next(x for x in at.button if x.label == "⚙️ 自动生成任务").click().run()
record("FT-013", "任务自动化", "生成结构化任务清单", no_exception(at) and len(at.session_state["automation_tasks"]) >= 3,
       f"task_count={len(at.session_state['automation_tasks'])}")

# Metrics page and feedback page.
at.session_state["page"] = "数据指标"
at.run()
record("FT-014", "数据指标", "事件指标页面加载", no_exception(at) and len(at.metric) >= 8,
       f"metrics={[(x.label, x.value) for x in at.metric]}")

at.session_state["page"] = "帮助与反馈"
at.run()
submit_buttons = [x for x in at.button if x.label == "提交反馈"]
if at.slider:
    at.slider[0].set_value(5)
if submit_buttons:
    submit_buttons[0].click().run()
feedback_saved = any("感谢反馈" in x.value for x in at.success)
record("FT-015", "帮助与反馈", "提交评分与反馈", no_exception(at) and feedback_saved,
       f"saved={feedback_saved}; success={[x.value for x in at.success]}")

# Static/data checks.
import sqlite3
con = sqlite3.connect(SANDBOX / "mvp_metrics.db")
tables = [x[0] for x in con.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")]
event_count = con.execute("SELECT COUNT(*) FROM events").fetchone()[0]
feedback_count = con.execute("SELECT COUNT(*) FROM feedback").fetchone()[0]
record("FT-016", "数据存储", "SQLite表结构与事件、反馈写入", tables == ["events", "feedback", "sqlite_sequence"] and event_count > 0 and feedback_count > 0,
       f"tables={tables}; event_count={event_count}; feedback_count={feedback_count}")

summary = {
    "total": len(results),
    "passed": sum(x["passed"] for x in results),
    "failed": sum(not x["passed"] for x in results),
    "results": results,
}
(ROOT / "mvp_function_test_results.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(summary, ensure_ascii=False, indent=2))
