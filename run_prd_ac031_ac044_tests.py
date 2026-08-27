import json
import sqlite3
import time
from pathlib import Path

from streamlit.testing.v1 import AppTest


ROOT = Path(__file__).resolve().parent
APP = ROOT / "app.py"
DB = ROOT / "mvp_metrics.db"


def fresh_app():
    at = AppTest.from_file(str(APP), default_timeout=90)
    at.run()
    return at


def row(case_id, title, status, evidence):
    return {"case_id": case_id, "title": title, "status": status, "evidence": evidence}


def event_count(name):
    with sqlite3.connect(DB) as connection:
        return connection.execute("SELECT COUNT(*) FROM events WHERE event_name=?", (name,)).fetchone()[0]


results = []
source = APP.read_text(encoding="utf-8")

# AC-031/032: recommendation, adoption, manual change and events.
at = fresh_app()
at.session_state["source_text"] = "会议讨论上传功能修复，张明负责，下周一回归测试。"
at.session_state["source_names"] = ["推荐测试文本"]
at.session_state["selected_feature"] = "会议纪要"
at.session_state["recommended_feature"] = ""
at.session_state["page"] = "选择功能"
shown_before = event_count("recommendation_shown")
adopt_before = event_count("recommendation_adopted")
change_before = event_count("feature_selected")
at.run(timeout=90)
recommendation = at.session_state["recommended_feature"]
reason = at.session_state["recommendation_reason"]
adopt_buttons = [x for x in at.button if x.label.startswith("采用推荐：")]
if adopt_buttons:
    adopt_buttons[0].click().run()
different = next((name for name in ("文档问答", "任务拆解", "周报生成") if name != recommendation), "任务拆解")
change_button = next((x for x in at.button if x.label == f"选择{different}"), None)
if change_button:
    change_button.click().run()
ac031 = recommendation in ("会议纪要", "文档问答", "任务拆解", "周报生成") and bool(reason) and bool(adopt_buttons) and change_button is not None
results.append(row("AC-031", "推荐功能、理由、采用或改选", "通过" if ac031 else "不通过",
                   f"recommendation={recommendation}, reason={bool(reason)}, adopt={bool(adopt_buttons)}, change={change_button is not None}"))

shown_after = event_count("recommendation_shown")
adopt_after = event_count("recommendation_adopted")
change_after = event_count("feature_selected")
ac032 = shown_after > shown_before and adopt_after > adopt_before and change_after > change_before
results.append(row("AC-032", "推荐展示、采用和修改埋点", "通过" if ac032 else "不通过",
                   f"shown_delta={shown_after-shown_before}, adopted_delta={adopt_after-adopt_before}, changed_delta={change_after-change_before}"))

# AC-033/034/035/036: task automation end-to-end.
auto = fresh_app()
auto.session_state["page"] = "任务自动化"
auto.run()
goal = next(x for x in auto.text_area if x.label == "项目目标或工作要求")
goal.set_value("两周内完成MVP测试、问题修复和验收报告。")
next(x for x in auto.button if x.label == "⚙️ 自动生成任务").click().run(timeout=90)
tasks = auto.session_state["automation_tasks"]
required = all(all(key in task for key in ("title", "owner", "priority", "deadline", "action", "done")) for task in tasks) if tasks else False
results.append(row("AC-033", "生成字段完整的结构化任务", "通过" if required else "不通过",
                   f"task_count={len(tasks)}, fields_complete={required}, errors={[x.value for x in auto.error]}"))

fallback_event = event_count("automation_fallback_used")
fallback_ok = bool(tasks) and fallback_event > 0
results.append(row("AC-034", "接口失败时模拟任务降级", "通过" if fallback_ok else "不通过",
                   f"task_count={len(tasks)}, fallback_events={fallback_event}, errors={[x.value for x in auto.error]}"))

if tasks:
    checkboxes = list(auto.checkbox)
    if checkboxes:
        checkboxes[0].set_value(True).run()
    progress_ok = len(auto.progress) > 0
    export_ok = len(auto.get("download_button")) > 0
    status_ok = bool(auto.session_state["automation_tasks"][0].get("done"))
    ac035_status = "通过" if progress_ok and export_ok and status_ok else "不通过"
    ac035_evidence = f"progress={progress_ok}, status={status_ok}, export={export_ok}"
else:
    ac035_status = "阻塞"
    ac035_evidence = "任务生成失败，无法继续验证状态、进度和导出"
results.append(row("AC-035", "状态更新、进度和任务导出", ac035_status, ac035_evidence))

auto_text = "\n".join([x.value for x in auto.caption] + [x.value for x in auto.info] + [x.value for x in auto.warning])
disclaimer = any(term in auto_text for term in ("不会真实执行", "仅用于", "辅助规划", "不连接外部办公平台"))
results.append(row("AC-036", "说明不会真实执行外部任务", "通过" if disclaimer else "不通过",
                   f"disclaimer={disclaimer}"))

# AC-037: feedback fields and persistence.
feedback_before = 0
with sqlite3.connect(DB) as connection:
    feedback_before = connection.execute("SELECT COUNT(*) FROM feedback").fetchone()[0]
feedback = fresh_app()
feedback.session_state["page"] = "帮助与反馈"
feedback.run()
if feedback.slider:
    feedback.slider[0].set_value(4)
radios = list(feedback.radio)
for radio in radios:
    if radio.label == "结果是否有帮助？":
        radio.set_value("有帮助")
    elif radio.label == "是否愿意再次使用？":
        radio.set_value("愿意")
comment = next((x for x in feedback.text_area if "意见" in x.label), None)
if comment:
    comment.set_value("AC037反馈字段验证。")
next(x for x in feedback.button if x.label == "提交反馈").click().run()
with sqlite3.connect(DB) as connection:
    feedback_after = connection.execute("SELECT COUNT(*) FROM feedback").fetchone()[0]
ac037 = feedback_after == feedback_before + 1
results.append(row("AC-037", "提交完整用户反馈", "通过" if ac037 else "不通过",
                   f"feedback_delta={feedback_after-feedback_before}"))

# AC-038: provider/fallback/result events must be distinct.
generation = fresh_app()
generation.session_state["source_text"] = "本周完成数据看板测试。"
generation.session_state["source_names"] = ["指标测试文本"]
generation.session_state["selected_feature"] = "周报生成"
generation.session_state["page"] = "选择功能"
event_names = ("qianfan_succeeded", "qianfan_failed", "local_fallback_used", "result_displayed")
before_events = {name: event_count(name) for name in event_names}
generation.run()
started = time.perf_counter()
next(x for x in generation.button if x.label == "开始生成 ➜").click().run(timeout=90)
elapsed = time.perf_counter() - started
after_events = {name: event_count(name) for name in event_names}
deltas = {name: after_events[name] - before_events[name] for name in event_names}
provider_failed = bool(generation.session_state["ai_error"])
if provider_failed:
    ac038 = deltas["qianfan_failed"] > 0 and deltas["local_fallback_used"] > 0 and deltas["result_displayed"] > 0
else:
    ac038 = deltas["qianfan_succeeded"] > 0 and deltas["result_displayed"] > 0
results.append(row("AC-038", "区分模型成功、降级、结果返回和API失败", "通过" if ac038 else "不通过",
                   f"provider_failed={provider_failed}, deltas={deltas}"))

# AC-039/040: dashboard metric labels and calculation coverage.
metrics = fresh_app()
metrics.session_state["page"] = "数据指标"
metrics.run()
metric_labels = [x.label for x in metrics.metric]
save_rate = any("保存率" in label for label in metric_labels)
results.append(row("AC-039", "主动保存率", "通过" if save_rate else "部分通过",
                   f"metric_labels={metric_labels}, result_saved_event={'result_saved' in source}"))

special_metrics = {
    "多轮对话率": any("多轮" in label for label in metric_labels),
    "推荐采用率": any("推荐采用率" in label for label in metric_labels),
    "Function Calling成功率": any("Function" in label or "任务自动化成功率" in label for label in metric_labels),
}
results.append(row("AC-040", "咨询、推荐和自动化专项指标", "通过" if all(special_metrics.values()) else "不通过",
                   special_metrics))

# AC-041: warning and metrics schema exclude uploaded document/chat bodies.
upload = fresh_app()
upload.session_state["page"] = "上传资料"
upload.run()
privacy_warning = any("敏感信息" in x.value and "不会保存文档正文" in x.value for x in upload.warning)
with sqlite3.connect(DB) as connection:
    event_columns = [row[1] for row in connection.execute("PRAGMA table_info(events)")]
    feedback_columns = [row[1] for row in connection.execute("PRAGMA table_info(feedback)")]
forbidden = {"document_body", "source_text", "chat_content", "conversation"}
schema_safe = not forbidden.intersection(event_columns + feedback_columns)
results.append(row("AC-041", "敏感信息提示及指标库不保存正文", "通过" if privacy_warning and schema_safe else "不通过",
                   f"privacy_warning={privacy_warning}, metric_schema_safe={schema_safe}"))

# AC-042: loading state and failure explanation exist in the flow.
spinner_present = "AI 正在生成结果，请稍候" in source
failure_explanation = "千帆调用失败，已自动使用本地演示结果" in source and "请检查 API Key" in source
results.append(row("AC-042", "加载状态和失败原因", "通过" if spinner_present and failure_explanation else "不通过",
                   f"spinner={spinner_present}, failure_explanation={failure_explanation}"))

# AC-043: real provider target cannot be accepted when fallback was used.
if generation.session_state["ai_error"]:
    status_043 = "未执行"
    evidence_043 = f"千帆调用失败，当前仅测得含本地降级的端到端耗时{elapsed:.2f}秒，不能作为真实模型响应时间"
else:
    status_043 = "通过" if elapsed <= 10 else "不通过"
    evidence_043 = f"真实千帆端到端耗时={elapsed:.2f}秒，目标<=10秒"
results.append(row("AC-043", "响应时间达到目标", status_043, evidence_043))

# AC-044: deprecated Streamlit argument remains.
deprecated_count = source.count("use_container_width")
results.append(row("AC-044", "Streamlit组件兼容性", "通过" if deprecated_count == 0 else "部分通过",
                   f"use_container_width_count={deprecated_count}"))

counts = {}
for item in results:
    counts[item["status"]] = counts.get(item["status"], 0) + 1
summary = {"total": len(results), "status_counts": counts, "results": results}
(ROOT / "prd_ac031_ac044_results.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(summary, ensure_ascii=False, indent=2))
