import json
from pathlib import Path

from streamlit.testing.v1 import AppTest


ROOT = Path(__file__).resolve().parent
APP = ROOT / "app.py"


def fresh_app():
    at = AppTest.from_file(str(APP), default_timeout=90)
    at.run()
    return at


def generate(at, feature, source, question=""):
    at.session_state["source_text"] = source
    at.session_state["source_names"] = ["测试文本"]
    at.session_state["selected_feature"] = feature
    at.session_state["page"] = "选择功能"
    at.run()
    if feature == "文档问答":
        box = next(x for x in at.text_input if "资料询问" in x.label)
        box.set_value(question)
    next(x for x in at.button if x.label == "开始生成 ➜").click().run(timeout=90)
    return at.session_state["result"]


def record(case_id, title, passed, evidence):
    return {"case_id": case_id, "title": title, "passed": bool(passed), "evidence": evidence}


results = []

# AC-011/012/013: one meeting without owner or deadline.
at = fresh_app()
meeting_source = "会议决定继续开展用户调研工作，并在后续会议中讨论具体执行安排。"
meeting = generate(at, "会议纪要", meeting_source)
ac011 = all(term in meeting for term in ("会议摘要", "关键结论", "待办事项"))
results.append(record("AC-011", "会议纪要基础结构", ac011,
                      "摘要/关键结论/待办事项=" + str(ac011)))

unclear_marked = any(term in meeting for term in ("待确认", "未明确", "资料中未提及"))
invented_generic = "项目负责人" in meeting and "本周内" in meeting
results.append(record("AC-012", "未明确负责人和截止时间", unclear_marked and not invented_generic,
                      f"unclear_marked={unclear_marked}, invented_generic={invented_generic}"))

# Edit, export and save the meeting result.
at.session_state["page"] = "生成结果"
at.run()
next(x for x in at.button if x.label == "✎ 编辑确认").click().run()
edit = next(x for x in at.text_area if "AI 生成结果" in x.label)
edit.set_value(meeting + "\n\n人工确认：会议纪要已核对。")
next(x for x in at.button if x.label == "✓ 保存修改").click().run()
edited_ok = "人工确认" in at.session_state["result"]
at.session_state["page"] = "导出保存"
at.run()
exports = []
for fmt in ("Word (.docx)", "PDF (.pdf)", "TXT (.txt)"):
    next(x for x in at.radio if x.label == "导出格式").set_value(fmt).run()
    exports.append(len(at.get("download_button")) > 0 and len(at.exception) == 0)
next(x for x in at.button if x.label == "▣ 保存到我的文件").click().run()
saved_ok = len(at.session_state["saved_files"]) > 0
results.append(record("AC-013", "会议纪要编辑导出保存", edited_ok and all(exports) and saved_ok,
                      f"edited={edited_ok}, exports={exports}, saved={saved_ok}"))

# AC-014/015/016/017: document QA.
qa_source = "本项目测试周期为4周，预算为5万元，第一批测试用户为20人。资料未说明项目负责人。"
qa_known = generate(at, "文档问答", qa_source, "项目测试周期和预算分别是多少？")
ac014 = "4周" in qa_known and "5万元" in qa_known
results.append(record("AC-014", "围绕文档获得相关答案", ac014,
                      f"contains_4周={'4周' in qa_known}, contains_5万元={'5万元' in qa_known}"))

qa_unknown_source = "本项目测试周期为4周，预算为5万元，第一批测试用户为20人。"
qa_unknown = generate(at, "文档问答", qa_unknown_source, "项目负责人是谁？")
no_answer = any(term in qa_unknown for term in ("无法确定", "未提及", "未明确", "没有答案"))
results.append(record("AC-015", "文档无答案时不编造", no_answer,
                      f"no_answer_marker={no_answer}; preview={qa_unknown[:180]}"))

source_evidence = "关键依据" in qa_known and any(
    term in qa_known for term in ("本项目测试周期为4周", "预算为5万元")
)
results.append(record("AC-016", "回答标注来源片段", source_evidence,
                      f"关键依据={'关键依据' in qa_known}, source_quote={source_evidence}"))

at.session_state["page"] = "生成结果"
at.run()
followup_control = len(at.chat_input) > 0 or any(
    "继续提问" in x.label or "追问" in x.label for x in at.text_input
)
results.append(record("AC-017", "文档问答多轮追问", followup_control,
                      f"chat_inputs={len(at.chat_input)}, followup_control={followup_control}"))

# AC-018/019/020: task breakdown without an explicit owner/deadline.
task_source = "请完成上传功能问题修复，包括问题复现、原因定位、开发修复和回归测试。"
tasks = generate(at, "任务拆解", task_source)
task_fields = (
    "任务" in tasks
    and "负责人" in tasks
    and "优先级" in tasks
    and any(term in tasks for term in ("截止", "时间建议", "时间安排"))
)
results.append(record("AC-018", "任务拆解规定字段", task_fields,
                      f"task_fields={task_fields}"))

task_unclear = any(term in tasks for term in ("待确认", "未明确"))
generic_assignments = all(term in tasks for term in ("项目负责人", "执行成员", "第 1 天"))
results.append(record("AC-019", "任务未明确字段处理", task_unclear and not generic_assignments,
                      f"unclear_marked={task_unclear}, generic_assignments={generic_assignments}"))

at.session_state["page"] = "生成结果"
at.run()
next(x for x in at.button if x.label == "✎ 编辑确认").click().run()
task_edit = next(x for x in at.text_area if "AI 生成结果" in x.label)
task_edit.set_value(tasks + "\n\n人工补充：任务结果已确认。")
next(x for x in at.button if x.label == "✓ 保存修改").click().run()
task_edited = "人工补充" in at.session_state["result"]
at.session_state["page"] = "导出保存"
at.run()
task_export = len(at.get("download_button")) > 0 and len(at.exception) == 0
next(x for x in at.button if x.label == "▣ 保存到我的文件").click().run()
task_saved = len(at.session_state["saved_files"]) >= 2
results.append(record("AC-020", "任务结果编辑导出保存", task_edited and task_export and task_saved,
                      f"edited={task_edited}, export={task_export}, saved={task_saved}"))

summary = {
    "total": len(results),
    "passed": sum(x["passed"] for x in results),
    "failed": sum(not x["passed"] for x in results),
    "results": results,
}
(ROOT / "prd_ac011_ac020_results.json").write_text(
    json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
)
print(json.dumps(summary, ensure_ascii=False, indent=2))
