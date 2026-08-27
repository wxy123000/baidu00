import ast
import re
from pathlib import Path

from streamlit.testing.v1 import AppTest


ROOT = Path(__file__).resolve().parent
APP = ROOT / "app.py"
SOURCE = APP.read_text(encoding="utf-8")


def check(condition, message):
    if not condition:
        raise AssertionError(message)
    print(f"PASS: {message}")


# 1、3：任务自动化在千帆不可用时仍生成任务，且字段完整。
auto = AppTest.from_file(str(APP), default_timeout=90)
auto.run()
auto.session_state["page"] = "任务自动化"
auto.run()
goal = next(item for item in auto.text_area if item.label == "项目目标或工作要求")
goal.set_value("5个工作日内完成Bug整理、修复、回归测试和测试报告。")
next(item for item in auto.button if item.label == "⚙️ 自动生成任务").click().run(timeout=90)
tasks = auto.session_state["automation_tasks"]
required = {
    "title", "owner", "priority", "deadline", "action",
    "dependency", "acceptance", "risk", "done",
}
check(bool(tasks), "任务自动化失败时仍返回本地任务清单")
check(all(required.issubset(task) for task in tasks), "每项自动化任务包含负责人、优先级、时间、依赖、验收和风险字段")
check(auto.session_state["automation_provider"] != "待生成", "任务来源被明确记录")


# 3：普通任务拆解的本地结果也必须具有完整字段。
task_block = SOURCE[SOURCE.index('if feature == "任务拆解":'):SOURCE.index('    return f"""本周工作周报')]
for label in ("工作说明：", "负责人建议：", "优先级：", "时间安排：", "前置依赖：", "验收标准：", "主要风险："):
    check(label in task_block, f"普通任务拆解包含“{label}”字段")


# 2：数据看板明确区分真实调用、降级、结果返回与接口失败。
dashboard = AppTest.from_file(str(APP), default_timeout=90)
dashboard.run()
dashboard.session_state["page"] = "数据指标"
dashboard.run()
metric_labels = {item.label for item in dashboard.metric}
for label in ("总体结果返回率", "千帆真实成功率", "本地降级率", "API调用失败率"):
    check(label in metric_labels, f"数据看板展示“{label}”")


# 4：独立执行检索函数，验证答案依据能够保留文件名和页码。
tree = ast.parse(SOURCE)
wanted = {
    "source_excerpt", "split_long_text", "question_terms",
    "select_long_document_context", "retrieve_relevant_passages",
}
selected_nodes = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in wanted]
module = ast.Module(body=selected_nodes, type_ignores=[])
namespace = {"re": re}
exec(compile(module, str(APP), "exec"), namespace)
sample = "【项目说明.pdf】\n【第 1 页】\n项目预算为5万元。\n【第 2 页】\n项目负责人为张明，测试周期为4周。"
passages = namespace["retrieve_relevant_passages"](sample, "项目负责人是谁？")
joined = "\n".join(passages)
check("项目说明.pdf" in joined, "文档问答依据保留来源文件名")
check("第 2 页" in joined, "文档问答依据保留页码")

print("RESULT: 4/4 optimization groups passed")
