"""Streamlit app — upload Excel, parse, visualize graph, Q&A."""

import json
import hashlib
from pathlib import Path

import streamlit as st

from core.parser import parse_workbook_with_values
from core.graph_builder import build_graph
from core.recalc_engine import RecalcEngine
from core.section_detector import detect_business_items
from storage.json_io import save_graph, load_graph
from storage.sqlite_db import init_db, add_upload, update_upload, get_uploads
from llm.query_resolver import QueryResolver

DATA_DIR = Path(__file__).parent.parent / "data"


def file_hash(path: str) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def main():
    st.set_page_config(page_title="财务模型知识图谱", layout="wide", page_icon="")
    st.title("财务模型知识图谱系统")

    init_db()

    tab_upload, tab_graph, tab_query, tab_recalc, tab_compare, tab_history = st.tabs(
        ["上传解析", "图谱可视化", "知识问答", "重算分析", "版本对比", "上传历史"]
    )

    with tab_upload:
        st.header("上传财务模型 Excel")
        uploaded = st.file_uploader("选择 .xlsx 文件", type=["xlsx"])

        if uploaded:
            save_path = DATA_DIR / "uploaded" / uploaded.name
            save_path.parent.mkdir(parents=True, exist_ok=True)
            with open(save_path, "wb") as f:
                f.write(uploaded.getvalue())

            with st.spinner("正在解析 Excel..."):
                try:
                    fg = build_graph(str(save_path))
                    items = detect_business_items(fg)

                    graph_path = DATA_DIR / "graphs" / f"{save_path.stem}.json"
                    save_graph(fg, graph_path)

                    upload_id = add_upload(uploaded.name, file_hash(str(save_path)))
                    update_upload(
                        upload_id,
                        status="success",
                        sheet_count=len({c.sheet for c in fg.cells.values()}),
                        cell_count=len(fg.cells),
                        formula_count=len(fg.get_cell_ids_with_formulas()),
                        graph_path=str(graph_path),
                    )

                    st.success(f"解析完成！{len(fg.cells)} 个单元格，{len(fg.get_cell_ids_with_formulas())} 个公式")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("业务项", len(items))
                    c2.metric("图谱节点", fg.graph.number_of_nodes())
                    c3.metric("关系边", fg.graph.number_of_edges())

                    # Store in session state
                    st.session_state["fg"] = fg
                    st.session_state["items"] = items
                    st.session_state["resolver"] = QueryResolver(fg)
                    st.session_state["recalc"] = RecalcEngine(fg)

                except Exception as e:
                    st.error(f"解析失败: {e}")

    with tab_graph:
        if "fg" not in st.session_state:
            st.warning("请先上传 Excel 文件")
        else:
            fg = st.session_state["fg"]
            items = st.session_state.get("items", [])

            c1, c2, c3 = st.columns(3)
            c1.metric("单元格", len(fg.cells))
            c2.metric("业务项", len(items))
            c3.metric("循环引用", fg.has_circular())

            # Sheet selector
            sheets = sorted({c.sheet for c in fg.cells.values()})
            selected_sheet = st.selectbox("选择 Sheet", sheets)

            sheet_cells = [c for c in fg.cells.values() if c.sheet == selected_sheet]
            formula_cells = [c for c in sheet_cells if c.formula_raw]

            st.write(f"**{selected_sheet}**: {len(sheet_cells)} 个单元格, {len(formula_cells)} 个公式")

            # Show business items from this sheet — paginated
            sheet_items = [i for i in items if i.sheet == selected_sheet]
            if sheet_items:
                st.subheader(f"业务项 ({len(sheet_items)} 个)")
                # Search filter
                search = st.text_input("搜索业务项", key=f"search_{selected_sheet}", placeholder="输入关键词...")
                filtered = sheet_items
                if search:
                    filtered = [i for i in sheet_items if search in i.name]
                    st.write(f"找到 {len(filtered)} 个匹配项")

                # Pagination
                page_size = 20
                total_pages = (len(filtered) + page_size - 1) // page_size
                page = st.slider("页码", 1, max(total_pages, 1), 1, key=f"page_{selected_sheet}")
                start = (page - 1) * page_size
                page_items = filtered[start:start + page_size]

                for item in page_items:
                    val_cell = fg.cells.get(item.value_cell) if item.value_cell else None
                    val = val_cell.value if val_cell else "N/A"
                    ts = "TS" if item.has_time_series else ""
                    unit = item.unit or ""
                    st.write(f"  - **{item.name}** = {val} {unit} {ts}")
                if len(filtered) > page_size:
                    st.caption(f"第 {page}/{total_pages} 页，共 {len(filtered)} 项")

                # Business item detail expander
                if sheet_items:
                    detail_name = st.selectbox("查看业务项详情", [i.name for i in sheet_items], key=f"detail_{selected_sheet}")
                    if detail_name:
                        detail_item = next((i for i in sheet_items if i.name == detail_name), None)
                        if detail_item:
                            with st.expander(f"📋 {detail_item.name}", expanded=True):
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.write(f"**Sheet**: {detail_item.sheet}")
                                    st.write(f"**单位**: {detail_item.unit or '-'}")
                                    st.write(f"**时间序列**: {'是' if detail_item.has_time_series else '否'}")
                                    if detail_item.section:
                                        st.write(f"**板块**: {detail_item.section}")
                                with col2:
                                    val_cell = fg.cells.get(detail_item.value_cell) if detail_item.value_cell else None
                                    st.write(f"**值单元格**: {detail_item.value_cell or '-'}")
                                    st.write(f"**值**: {val_cell.value if val_cell else 'N/A'}")
                                    st.write(f"**关联单元格**: {len(detail_item.cell_ids)} 个")

                                # Time series chart
                                if detail_item.has_time_series and detail_item.columns.time_series_start:
                                    from models.cell_node import col_to_index
                                    start_ci = col_to_index(detail_item.columns.time_series_start)
                                    end_ci = col_to_index(detail_item.columns.time_series_end) if detail_item.columns.time_series_end else start_ci
                                    ts_values = []
                                    ts_labels = []
                                    for cid in detail_item.cell_ids:
                                        cell = fg.cells.get(cid)
                                        if cell and start_ci <= cell.col_index <= end_ci and isinstance(cell.value, (int, float)):
                                            ts_labels.append(f"{cell.col}{cell.row}")
                                            ts_values.append(cell.value)
                                    if ts_values:
                                        import pandas as pd
                                        df = pd.DataFrame({"单元格": ts_labels, "值": ts_values})
                                        st.line_chart(df.set_index("单元格"))

                                # Dependency chain
                                deps = fg.get_dependencies(detail_item.value_cell) if detail_item.value_cell else []
                                if deps:
                                    st.write(f"**依赖链** (前10): {' → '.join(deps[:10])}")

            # Formula dependencies
            if formula_cells:
                with st.expander("公式依赖关系 (前20)"):
                    for cell in formula_cells[:20]:
                        deps = fg.get_dependencies(cell.id)
                        if deps:
                            st.write(f"`{cell.id}` --> {', '.join(deps[:5])}")

            # pyvis visualization — filtered to selected sheet
            viz_mode = st.radio("可视化范围", ["当前 Sheet", "全局摘要"], horizontal=True)
            try:
                from pyvis.network import Network
                net = Network(height="600px", width="100%", directed=True, notebook=True)

                if viz_mode == "当前 Sheet":
                    # Show only selected sheet's formula cells + 1-hop neighbors
                    sheet_formula = [c for c in sheet_cells if c.formula_raw][:80]
                    node_ids = {c.id for c in sheet_formula}
                    # Add 1-hop dependencies (max 70 more)
                    dep_count = 0
                    for cell in sheet_formula:
                        for tgt in fg.graph.successors(cell.id):
                            if tgt not in node_ids and dep_count < 70:
                                node_ids.add(tgt)
                                dep_count += 1
                else:
                    # Global summary: business items + key relationships
                    bi_ids = {bi.id for bi in items[:50]}
                    node_ids = bi_ids
                    for bi_id in bi_ids:
                        for dep in fg.get_dependencies(bi_id)[:3]:
                            node_ids.add(dep)

                # Build nodes
                for cid in node_ids:
                    cell = fg.cells.get(cid)
                    if cell:
                        label = f"{cell.col}{cell.row}"
                        if cell.formula_raw:
                            label = f"={cell.formula_raw[:15]}..."
                        color = "#FF6B6B" if cell.formula_raw else "#4ECDC4"
                        net.add_node(cid, label=label, color=color, title=str(cell.value or ""))

                # Build edges (only within selected nodes)
                edge_count = 0
                for src in node_ids:
                    for tgt in fg.graph.successors(src):
                        if tgt in node_ids and edge_count < 500:
                            net.add_edge(src, tgt)
                            edge_count += 1

                net.save_graph(str(DATA_DIR / "graphs" / "graph_viz.html"))
                st.components.v1.html(
                    open(DATA_DIR / "graphs" / "graph_viz.html", encoding="utf-8").read(),
                    height=650,
                )
            except Exception as e:
                st.info(f"可视化加载中... ({e})")

    with tab_query:
        st.header("知识问答")
        if "fg" not in st.session_state:
            st.warning("请先上传 Excel 文件")
        else:
            items = st.session_state.get("items", [])
            resolver = st.session_state.get("resolver")

            # Show available financial indicators
            with st.expander("可查询的财务指标"):
                by_sheet = {}
                for item in items[:100]:
                    by_sheet.setdefault(item.sheet, []).append(item)
                for sheet, s_items in sorted(by_sheet.items()):
                    st.write(f"**{sheet}**: {', '.join(i.name for i in s_items[:5])}...")

            # Query input
            query = st.text_input("输入查询问题", placeholder="如: 建设期是多少？2030年营业收入是多少？总投资和资本金对比")
            if query and resolver:
                result = resolver.resolve(query)
                st.subheader("查询结果")
                st.write(result.explanation)

                # Comparison mode
                if result.compare_entity:
                    col_a, col_b = st.columns(2)
                    item_a, item_b = result.entity, result.compare_entity
                    with col_a:
                        st.metric("A: " + item_a.name, result.value or "N/A", item_a.unit or "")
                    with col_b:
                        st.metric("B: " + item_b.name, result.compare_value or "N/A", item_b.unit or "")

                    # Side-by-side details
                    col_a, col_b = st.columns(2)
                    with col_a:
                        st.write(f"**Sheet**: {item_a.sheet}")
                        st.write(f"**时间序列**: {'是' if item_a.has_time_series else '否'}")
                        if item_a.section:
                            st.write(f"**所属板块**: {item_a.section}")
                    with col_b:
                        st.write(f"**Sheet**: {item_b.sheet}")
                        st.write(f"**时间序列**: {'是' if item_b.has_time_series else '否'}")
                        if item_b.section:
                            st.write(f"**所属板块**: {item_b.section}")

                    # Value comparison
                    if result.value is not None and result.compare_value is not None:
                        try:
                            diff = float(result.value) - float(result.compare_value)
                            pct = (diff / abs(float(result.compare_value)) * 100) if result.compare_value != 0 else 0
                            st.metric("差值 (A - B)", f"{diff:.2f}", f"{pct:+.1f}%")
                        except (ValueError, TypeError):
                            pass

                elif result.entity:
                    item = result.entity
                    val_cell = resolver.fg.cells.get(item.value_cell) if item.value_cell else None
                    cols = st.columns(3)
                    cols[0].metric("指标", item.name)
                    cols[1].metric("值", result.value or val_cell.value if val_cell else "N/A")
                    cols[2].metric("单位", item.unit or "-")

                    if item.has_time_series:
                        st.write("时间序列: 是")
                    if result.related_items:
                        st.write(f"相关指标: {', '.join(result.related_items)}")

    with tab_recalc:
        st.header("重算分析")
        if "fg" not in st.session_state:
            st.warning("请先上传 Excel 文件")
        else:
            fg = st.session_state["fg"]
            recalc = st.session_state.get("recalc")
            items = st.session_state.get("items", [])

            st.write("修改输入参数，查看对整个财务模型的影响")

            # Filter by sheet first
            all_sheets = sorted({c.sheet for c in fg.cells.values() if not c.formula_raw and c.data_type == "number"})
            sel_sheet = st.selectbox("选择 Sheet", all_sheets, key="recalc_sheet")

            # Filter input cells by sheet, sorted by value magnitude
            input_cells = [c for c in fg.cells.values()
                          if not c.formula_raw and c.data_type == "number" and c.sheet == sel_sheet]
            input_cells = sorted(input_cells, key=lambda c: abs(c.value) if isinstance(c.value, (int, float)) else 0, reverse=True)

            # Search filter
            search = st.text_input("搜索参数", placeholder="输入参数名或单元格...", key="recalc_search")
            if search:
                input_cells = [c for c in input_cells if search in c.id or search in str(c.value)]

            # Paginate
            page_size = 30
            total_pages = max(1, (len(input_cells) + page_size - 1) // page_size)
            page = st.slider("页码", 1, total_pages, 1, key="recalc_page")
            page_cells = input_cells[(page - 1) * page_size:page * page_size]

            cell_id = st.selectbox(
                "选择要修改的单元格",
                [f"{c.id} = {c.value:.4f}" if isinstance(c.value, (int, float)) else f"{c.id} = {c.value}" for c in page_cells],
            )

            new_value = st.number_input("新值", value=0.0, step=1.0)

            if st.button("执行重算"):
                if cell_id and recalc:
                    cid = cell_id.split(" = ")[0]
                    cell = fg.cells.get(cid)
                    if cell:
                        with st.spinner("重算中..."):
                            old_val = cell.value
                            cell.value = new_value
                            result = recalc.recalculate(cid)
                            cell.value = old_val  # Restore input cell only — engine doesn't mutate dependents

                        st.success(f"重算完成，{result.total_changed} 个单元格发生变化")

                        if result.changed_cells:
                            # Summary metrics
                            numeric_changes = [d for d in result.changed_cells
                                              if isinstance(d.old_value, (int, float)) and isinstance(d.new_value, (int, float))]
                            if numeric_changes:
                                top_5 = sorted(numeric_changes, key=lambda d: abs(d.new_value - d.old_value), reverse=True)[:5]
                                st.subheader("影响最大的 5 个单元格")
                                for d in top_5:
                                    diff = d.new_value - d.old_value
                                    pct = (diff / abs(d.old_value) * 100) if d.old_value != 0 else 0
                                    st.metric(d.cell_id, f"{d.new_value:,.2f}", f"{diff:+,.2f} ({pct:+.1f}%)")

                            # Bar chart of top changes
                            if numeric_changes and len(numeric_changes) > 1:
                                import pandas as pd
                                chart_data = []
                                for d in sorted(numeric_changes, key=lambda d: abs(d.new_value - d.old_value), reverse=True)[:15]:
                                    diff = d.new_value - d.old_value
                                    short_id = d.cell_id.split("_")[-2] + d.cell_id.split("_")[-1]  # e.g. "96_I"
                                    chart_data.append({"单元格": short_id, "变化量": diff})
                                df = pd.DataFrame(chart_data)
                                st.bar_chart(df.set_index("单元格"), horizontal=True)

                            # Full detail expander
                            with st.expander("完整变化列表"):
                                for i, delta in enumerate(result.changed_cells):
                                    if isinstance(delta.old_value, (int, float)) and isinstance(delta.new_value, (int, float)):
                                        st.write(f"`{delta.cell_id}`: {delta.old_value:,.4f} → {delta.new_value:,.4f}")
                                    else:
                                        st.write(f"`{delta.cell_id}`: {delta.old_value} → {delta.new_value}")

    with tab_compare:
        st.header("版本对比")
        st.write("上传两个不同的 Excel 文件进行对比")

        col_a, col_b = st.columns(2)
        with col_a:
            uploaded_a = st.file_uploader("版本 A", type=["xlsx"], key="file_a")
        with col_b:
            uploaded_b = st.file_uploader("版本 B", type=["xlsx"], key="file_b")

        if uploaded_a and uploaded_b:
            if st.button("开始对比"):
                save_a = DATA_DIR / "uploaded" / f"_cmp_a_{uploaded_a.name}"
                save_b = DATA_DIR / "uploaded" / f"_cmp_b_{uploaded_b.name}"

                with open(save_a, "wb") as f:
                    f.write(uploaded_a.getvalue())
                with open(save_b, "wb") as f:
                    f.write(uploaded_b.getvalue())

                with st.spinner("解析并对比中..."):
                    try:
                        from storage.version_diff import compare_versions_by_excel
                        result = compare_versions_by_excel(
                            str(save_a), str(save_b),
                            uploaded_a.name, uploaded_b.name,
                        )

                        st.success(result.summary)

                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("新增", result.added_count)
                        c2.metric("删除", result.removed_count)
                        c3.metric("修改", result.modified_count)
                        c4.metric("不变", result.unchanged_count)

                        # Visual: diff distribution by sheet
                        if result.node_diffs:
                            modified_diffs = [d for d in result.node_diffs if d.status == "modified"]
                            if modified_diffs and len(modified_diffs) > 1:
                                import pandas as pd
                                from collections import Counter
                                sheet_counts = Counter()
                                for d in modified_diffs:
                                    sheet = d.cell_id.split("_")[0] if "_" in d.cell_id else "unknown"
                                    sheet_counts[sheet] += 1
                                df = pd.DataFrame({"Sheet": list(sheet_counts.keys()), "变化数": list(sheet_counts.values())})
                                st.bar_chart(df.set_index("Sheet"), horizontal=True)

                        # Show top diffs
                        if result.node_diffs:
                            st.subheader("单元格变化 (前20)")
                            for diff in result.node_diffs[:20]:
                                status_icon = {"added": "+", "removed": "-", "modified": "~"}.get(diff.status, "?")
                                old_s = diff.old_value if diff.old_value is not None else "-"
                                new_s = diff.new_value if diff.new_value is not None else "-"
                                st.write(f"  {status_icon} `{diff.cell_id}`: {old_s} --> {new_s}")

                        if result.item_diffs:
                            with st.expander(f"业务项变化 ({len(result.item_diffs)} 个)"):
                                added = [i for i in result.item_diffs if i.status == "added"]
                                removed = [i for i in result.item_diffs if i.status == "removed"]
                                modified = [i for i in result.item_diffs if i.status == "modified"]
                                if added:
                                    st.write(f"+ 新增 {len(added)} 个: {', '.join(i.item_name for i in added[:5])}")
                                if removed:
                                    st.write(f"- 删除 {len(removed)} 个: {', '.join(i.item_name for i in removed[:5])}")
                                if modified:
                                    st.write(f"~ 修改 {len(modified)} 个:")
                                    for item in modified[:10]:
                                        st.write(f"  {item.item_name}: {item.old_value} --> {item.new_value} {item.unit}")

                    except Exception as e:
                        st.error(f"对比失败: {e}")

    with tab_history:
        st.header("上传历史")
        uploads = get_uploads()
        if uploads:
            for u in uploads[:10]:
                status_color = "green" if u["status"] == "success" else "red"
                with st.expander(f"[{u['filename']}] — {u['status']} — {u['cell_count']} cells, {u['formula_count']} formulas — {u['upload_time']}"):
                    st.write(f"**文件**: {u['filename']}")
                    st.write(f"**状态**: {u['status']}")
                    st.write(f"**Sheet 数**: {u.get('sheet_count', 'N/A')}")
                    st.write(f"**单元格**: {u['cell_count']}")
                    st.write(f"**公式**: {u['formula_count']}")
                    if u.get('graph_path'):
                        st.write(f"**图谱路径**: {u['graph_path']}")
        else:
            st.info("暂无上传记录")


if __name__ == "__main__":
    main()
