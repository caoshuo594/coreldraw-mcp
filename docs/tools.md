# 工具入口与高级功能

## 18 个直接工具

| 工具 | 用途 |
| --- | --- |
| `corel_status` | 读取连接状态与打开文档 |
| `list_open_documents` | 列出路径、修改状态、页数及尺寸 |
| `attach_active_document` | 绑定明确指定的当前文档 |
| `attach_open_document` | 按完整路径绑定已打开文件 |
| `create_new_document` | 新建指定毫米宽高的文档 |
| `inspect_document` | 读取指定会话的页面与对象 |
| `apply_operations` | 执行带修订检查的短批次编辑 |
| `render_preview` | 导出 PNG 预览 |
| `undo_last_edit` | 撤销文档最近的命令组 |
| `redo_last_edit` | 重做最近撤销的命令组 |
| `save_current_document` | 保存已有路径的文档 |
| `save_document_as` | 另存为绝对 CDR 路径 |
| `export_document` | 导出 PDF、DXF、SVG、PNG、JPEG |
| `detach_document` | 结束会话，不关闭文档 |
| `open_document` | 打开绝对路径文件，已打开则绑定 |
| `list_feature_tools` | 查询高级名称、分类和 JSON 参数结构 |
| `run_feature_tool` | 调用一个高级功能 |
| `verify_saved_document` | 重开磁盘副本，报告文字、字体、尺寸及溢出 |

精确参数以客户端取得的工具结构为准。实现见 [live.py](../server/tools/live.py)、[server.py](../server/server.py)、[unified.py](../server/tools/unified.py)。

## 查询与调用

```text
list_feature_tools()
list_feature_tools(category="shapes")
list_feature_tools(name="create_rectangle")
```

指定 `name` 会返回输入 JSON Schema。文档的高级修改通常需要 `session_id` 和 `expected_revision`；读表、模板列表、打开/新建文件与隔离模板合并使用各自路径。

```json
{
  "name": "create_rectangle",
  "arguments": {"x": 10, "y": 20, "width": 80, "height": 40, "top_left": true},
  "session_id": "从绑定结果读取",
  "expected_revision": 0
}
```

常规编辑优先使用 `apply_operations`。部分高级工具沿用原生坐标，先用查询结果和已知对象确认。

## 单页工作区与导出范围

页面框不限制可编辑内容的摆放范围。`render_preview` 与 PNG/JPEG 的 `export_document` 支持以下 `scope`；保存验收的预览使用对应的 `preview_scope`。

| 范围 | 包含内容 | 执行方式 |
| --- | --- | --- |
| `page`（默认） | 页面矩形范围 | 保持原来的页面输出行为 |
| `artwork` | 当前页可见、可打印对象，包含符合条件的框外对象 | 按对象整体包围范围输出 |
| `workspace` | 可见可打印对象，以及可见桌面层对象 | 逐个复制根对象到临时文档输出，随后关闭临时文档 |

`padding_mm` 指定图上毫米留边。`workspace` 会使用剪贴板，不修改源页面尺寸、对象位置或桌面层打印开关，也不包含隐藏层和普通不可打印层。

```json
{
  "session_id": "从绑定结果读取",
  "destination_path": "D:\\Projects\\preview.png",
  "width_px": 1600,
  "scope": "workspace",
  "padding_mm": 4
}
```

X8 可能在保存或重开后把框外对象转入不可打印的共享桌面层；对象计数正确不代表页面范围预览包含全部内容。含桌面内容的设计应使用 `workspace` 并检查保存重开的完整预览。单独把对象移回本地图层后，不能仅凭内存状态认定已解决。

PDF/SVG/DXF 的核心导出入口仍使用页面范围；传入非页面范围会明确报错。需要框外内容的矢量输出时，在临时导出文档中设置实际输出范围，不为导出而改变源工作区。升级代码后需重新连接 MCP，才能加载新增参数与实现。

## X8 素材导入

`import_svg` 和 `import_image` 使用 X8 支持的图层 `ImportEx` / `Finish` 接口。新增对象按导入前后的 StaticID 差集识别，避免把旧对象当成导入结果。多个新对象按明确范围组合；只有一个对象时保留该对象，不调用单对象分组。

## 73 项高级功能

下面的分类清单由 [feature_catalog.json](../server/feature_catalog.json) 生成，函数实现位于 `server/tools/` 对应模块。

### document（11 项）

[阅读实现](../server/tools/document.py)

- `open_template`
- `create_document`
- `save_document`
- `close_document`
- `add_page`
- `set_page_size`
- `get_document_info`
- `list_all_text_shapes`
- `add_guideline`
- `switch_page`
- `delete_page`

### shapes（20 项）

[阅读实现](../server/tools/shapes.py)

- `create_rectangle`
- `create_ellipse`
- `create_line`
- `import_svg`
- `import_image`
- `set_shape_size`
- `set_shape_position`
- `boolean_operation`
- `convert_to_curves`
- `group_shapes`
- `find_shape_by_name`
- `align_shapes`
- `distribute_shapes`
- `set_z_order`
- `delete_shape`
- `rotate_shape`
- `ungroup_shapes`
- `scale_shape`
- `select_shapes`
- `powerclip`

### text（6 项）

[阅读实现](../server/tools/text.py)

- `set_text_content`
- `set_text_style`
- `fit_text_to_frame`
- `check_text_overflow`
- `convert_text_to_curves`
- `create_text_frame`

### colors（9 项）

[阅读实现](../server/tools/colors.py)

- `set_fill_cmyk`
- `set_fill_rgb`
- `set_fill_pantone`
- `set_outline`
- `set_no_fill`
- `set_no_outline`
- `check_rgb_colors`
- `set_fountain_fill`
- `set_transparency`

### layers（5 项）

[阅读实现](../server/tools/layers.py)

- `create_layer`
- `get_layers`
- `assign_to_layer`
- `set_layer_visible`
- `lock_layer`

### export（8 项）

[阅读实现](../server/tools/export.py)

- `export_pdf`
- `export_dxf`
- `export_ai`
- `export_svg`
- `export_png`
- `export_jpeg`
- `export_preview_png`
- `batch_export`

### preflight（4 项）

[阅读实现](../server/tools/preflight.py)

- `check_dimensions`
- `check_text_overflow_all`
- `check_missing_fonts`
- `get_color_report`

### templates（3 项）

[阅读实现](../server/tools/templates.py)

- `list_templates`
- `get_template_info`
- `get_size_variant`

### engineering（2 项）

[阅读实现](../server/tools/engineering.py)

- `populate_title_block`
- `assemble_engineering_drawing`

### data_merge（5 项）

[阅读实现](../server/tools/data_merge.py)

- `read_excel_data`
- `merge_record`
- `batch_merge`
- `generate_barcode`
- `generate_qrcode`
