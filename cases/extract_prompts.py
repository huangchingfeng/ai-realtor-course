"""從 5 個 generate_ppt_module*.py 提取案例資料，輸出 prompts.json"""
import json
import re
import ast
import os

CASES_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(os.path.dirname(CASES_DIR), "prompt-library", "data")

# 模組元資料
MODULES = [
    {"id": "M1", "name": "AI 行情分析師", "description": "讓 AI 幫你做完行情功課", "icon": "📊"},
    {"id": "M2", "name": "AI 簡報製作", "description": "用 AI 做出專業簡報", "icon": "📑"},
    {"id": "M3", "name": "AI 物件行銷工廠", "description": "一間物件，14 件跨平台內容", "icon": "📢"},
    {"id": "M4", "name": "AI 追蹤成交系統", "description": "從帶看到成交，AI 幫你追到底", "icon": "🤝"},
    {"id": "M5", "name": "AI 公域獲客飛輪", "description": "社群內容批量生產，永不斷更", "icon": "🚀"},
]


def extract_cases_from_file(filepath):
    """從 Python 檔案中提取 cases 列表"""
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()

    # 找到 cases = [ 的起始位置
    match = re.search(r"^cases\s*=\s*\[", content, re.MULTILINE)
    if not match:
        print(f"  找不到 cases 陣列：{filepath}")
        return []

    # 從 cases = [ 開始，找到對應的 ]
    start = match.start()
    bracket_count = 0
    end = start
    in_string = False
    string_char = None
    escape_next = False

    for i in range(match.end() - 1, len(content)):
        c = content[i]

        if escape_next:
            escape_next = False
            continue

        if c == "\\":
            escape_next = True
            continue

        if in_string:
            if c == string_char:
                in_string = False
            continue

        if c in ('"', "'"):
            in_string = True
            string_char = c
            continue

        if c == "[":
            bracket_count += 1
        elif c == "]":
            bracket_count -= 1
            if bracket_count == 0:
                end = i + 1
                break

    cases_str = content[start:end]

    # 用 ast.literal_eval 解析
    list_str = cases_str[cases_str.index("["):]
    try:
        cases = ast.literal_eval(list_str)
        return cases
    except (SyntaxError, ValueError) as e:
        print(f"  解析失敗：{filepath} - {e}")
        return []


def normalize_level(level_str):
    """標準化難度標籤"""
    if "新手" in level_str:
        return {"level": "新手", "levelTag": "green"}
    elif "中級" in level_str:
        return {"level": "中級", "levelTag": "yellow"}
    elif "進階" in level_str:
        return {"level": "進階", "levelTag": "red"}
    return {"level": "未分類", "levelTag": "gray"}


def normalize_tool(tool_str):
    """提取工具分類"""
    tool_lower = tool_str.lower()
    categories = []
    if "chatgpt" in tool_lower:
        categories.append("chatgpt")
    if "gemini" in tool_lower:
        categories.append("gemini")
    if "perplexity" in tool_lower:
        categories.append("perplexity")
    if "gamma" in tool_lower:
        categories.append("gamma")
    if "notebooklm" in tool_lower:
        categories.append("notebooklm")
    if "canva" in tool_lower:
        categories.append("canva")
    if not categories:
        categories.append("other")
    return categories


def clean_system_note(note):
    """清理 system_note 中的 emoji 前綴"""
    if not note:
        return ""
    return note.replace("🔗 系統化：", "").replace("🔗 ", "").strip()


def main():
    all_prompts = []

    for i, module in enumerate(MODULES):
        module_num = i + 1
        filepath = os.path.join(CASES_DIR, f"generate_ppt_module{module_num}.py")

        print(f"處理 {module['id']}：{module['name']}...")

        if not os.path.exists(filepath):
            print(f"  檔案不存在：{filepath}")
            continue

        cases = extract_cases_from_file(filepath)
        print(f"  提取到 {len(cases)} 個案例")

        for case in cases:
            level_info = normalize_level(case.get("level", ""))
            system_note_raw = case.get("system_note", "")

            prompt_entry = {
                "id": f"{module['id']}-{case['num']}",
                "module": module["id"],
                "num": case["num"],
                "title": case["title"],
                "level": level_info["level"],
                "levelTag": level_info["levelTag"],
                "tool": case["tool"],
                "toolCategories": normalize_tool(case["tool"]),
                "scenario": case["scenario"],
                "task": case["task"],
                "uploadFormat": case.get("upload_format", ""),
                "uploadFile": case.get("upload_file", ""),
                "uploadNote": case.get("upload_note", ""),
                "systemNote": clean_system_note(system_note_raw),
                "prompt": case["prompt"],
            }
            all_prompts.append(prompt_entry)

    # 統計
    total = len(all_prompts)
    level_dist = {}
    tool_dist = {}
    for p in all_prompts:
        level_dist[p["level"]] = level_dist.get(p["level"], 0) + 1
        for tc in p["toolCategories"]:
            tool_dist[tc] = tool_dist.get(tc, 0) + 1

    # 組裝輸出
    output = {
        "meta": {
            "total": total,
            "lastUpdated": "2026-03-02",
            "version": "1.0",
            "courseName": "AI 超級房仲實戰班",
            "instructor": "阿峰老師（黃敬峰）",
        },
        "stats": {
            "levelDistribution": level_dist,
            "toolDistribution": tool_dist,
        },
        "modules": MODULES,
        "prompts": all_prompts,
    }

    # 確保輸出目錄存在
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, "prompts.json")

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n完成！共 {total} 個案例")
    print(f"輸出：{output_path}")
    print(f"難度分布：{level_dist}")
    print(f"工具分布：{tool_dist}")


if __name__ == "__main__":
    main()
