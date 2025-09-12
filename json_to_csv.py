#!/usr/bin/env python3
"""
将上一步生成的嵌套 JSON 扁平化为适合 Dify 导入的 CSV。

功能要点：
- 使用 sys.argv 接收输入 .json 文件路径。
- 使用内置 json/csv 库读取与写出。
- 递归遍历 JSON 的 sections 树，生成知识块(Chunk)行。
- 若节点的 content 非空，则将其合并为一个文本块写出一行。
- 在递归时维护父级标题路径，生成以 ' > ' 连接的 section_path。
- CSV 使用 utf-8-sig 编码，Excel 打开更友好。
"""

from __future__ import annotations

import csv
import json
import os
import sys
from typing import Any, Dict, Iterable, List, Sequence


HEADERS: Sequence[str] = (
    "text",
    "sop_id",
    "sop_name",
    "section_path",
    "image_filenames",
)


def load_json(input_path: str) -> Dict[str, Any]:
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"文件不存在: {input_path}")
    if not input_path.lower().endswith(".json"):
        raise ValueError("输入文件必须为 .json 格式")
    try:
        with open(input_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as exc:
        raise ValueError(f"JSON 解析失败: {exc}")


def iter_chunks(root: Dict[str, Any]) -> Iterable[Dict[str, str]]:
    """遍历 JSON 树，生成 CSV 行所需的数据字典。

    约定输入 JSON 根结构：
    {
      "sop_id": str,
      "sop_name": str,
      "sections": [
        {"title": str, "level": int, "content": [str, ...], "subsections": [...]},
        ...
      ]
    }

    递归规则：
    - 对每个节点(node)，如果其 content 非空，则将所有 content 合并为一个字符串，
      并输出一行：text, sop_id, sop_name, section_path, image_filename(空字符串)。
    - section_path 为从根到当前节点(含当前)的标题名，使用 ' > ' 连接。
    - 然后继续递归其 subsections。
    """

    sop_id = str(root.get("sop_id", "")).strip()
    sop_name = str(root.get("sop_name", "")).strip()
    sections = root.get("sections") or []

    def _walk(node: Dict[str, Any], path_titles: List[str]) -> Iterable[Dict[str, str]]:
        # 当前节点的标题加入路径
        title = str(node.get("title", "")).strip()
        current_path = path_titles + ([title] if title else [])

        # 若 content 非空，合并为一个文本块，作为一条知识行
        content_list = node.get("content") or []
        if isinstance(content_list, list):
            merged_text = "\n".join(
                [str(item).strip() for item in content_list if str(item).strip()]
            ).strip()
        else:
            merged_text = str(content_list).strip()

        # 获取图片列表
        images_list = node.get("images") or []
        if isinstance(images_list, list):
            image_filenames = ",".join([str(img) for img in images_list if str(img).strip()])
        else:
            image_filenames = ""

        if merged_text:
            yield {
                "text": merged_text,
                "sop_id": sop_id,
                "sop_name": sop_name,
                "section_path": " > ".join([p for p in current_path if p]),
                "image_filenames": image_filenames,
            }

        # 递归子节点
        for child in node.get("subsections") or []:
            if isinstance(child, dict):
                yield from _walk(child, current_path)

    for top in sections:
        if isinstance(top, dict):
            yield from _walk(top, [])


def write_csv(rows: Iterable[Dict[str, str]], output_path: str) -> None:
    # 使用 utf-8-sig 以便 Excel 识别 BOM，避免中文乱码
    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(HEADERS))
        writer.writeheader()
        for row in rows:
            # 仅保留已定义的头字段，避免多余键
            writer.writerow({k: row.get(k, "") for k in HEADERS})


def main(argv: List[str]) -> int:
    if len(argv) < 2:
        sys.stderr.write("用法: python json_to_csv.py <输入文件.json>\n")
        return 1

    input_path = argv[1]
    try:
        data = load_json(input_path)
    except (FileNotFoundError, ValueError) as exc:
        sys.stderr.write(f"[ERROR] {exc}\n")
        return 1
    except Exception as exc:  # pragma: no cover
        sys.stderr.write(f"[ERROR] 未预期的异常: {exc}\n")
        return 1

    rows = list(iter_chunks(data))

    # 输出同名 .csv
    output_path = os.path.splitext(input_path)[0] + ".csv"
    try:
        write_csv(rows, output_path)
    except Exception as exc:  # pragma: no cover
        sys.stderr.write(f"[ERROR] 写入 CSV 失败: {exc}\n")
        return 1

    print(f"已生成: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))


