# CalcBook V3.7.4 内测版说明

> 本版本为**内测版**：不上传 Gitee / GitHub，仅本机试用。源码本地提交，未推送任何远端。

## 一、六项需求落地

1. **表格默认显示行列区**：新建工程 / 新建分部时自动预置 15 行空计算式（「选项配置 → 分部 → 默认计算式个数」可调）；老设置里存的默认值 1 会一次性迁移为 15，用户自己改过个数的不动。
2. **分部打断**：新增 `split_section_at_row`——表格右键「在此行打断分部（本行起归新分部）」、行号右键、Ribbon「编辑」组「分部打断」三个入口；本行起以下各行移入紧随其后的新分部（名称对话框，重名自动加序号 `_unique_child_name/_unique_book_name`），打断后自动切到新分部；支持撤销/恢复。
3. **项目结构树行距收紧**：`QTreeWidget::item { height: fm.height()+6px }`，按名称字体实际高度显示。
4. **全选角右键菜单修复并加强**：
   - **真 bug**：旧代码只在 build_ui 时 `findChild(QAbstractButton)` 接线一次；角标按钮由 Qt 延迟创建/可能被重建，运行时连接丢失 → 右键毫无反应（探针实测复现：emit 信号无槽响应）。
   - 修法：`_ensure_corner_menu()` 在每次 `refresh_table` 时惰性接线，按按钮对象（`_corner_btn`）防重连。
   - 菜单补上复制 / 剪切 / 粘贴 / 清除内容 / 删除选中行 / 填充颜色等单元格操作。
5. **冻结窗格修复（真根因）**：
   - **bug**：`closeEditor` 信号在 `setItemDelegate(LiveSyncDelegate)` **之前**接在旧默认代理上，代理被替换后连接失效 → 编辑过一次单元格后 `hide_for_edit()` 隐藏的冻结镜像**永不恢复**，滚动时冻结区消失。
   - 修法：连接移到代理设置之后；另加 `sync_one_height`（行高单行同步，接 `verticalHeader().sectionResized`）、`_keep_visible`（滚动兜底恢复，编辑让位期间除外）、冻结镜像右键转发主表菜单（`_MirrorTable.contextMenuEvent`）。
6. **函数面板调整**：面板幂运算 power 取消（公式里 `2^3` 仍可用），改为立方 `cube`（引擎注册 `'cube': lambda x: x*x*x`）；常量 π 组挪到面板最上方。

## 二、测试

- 新增 `smoke_v52.py`：27/27 通过（版本、默认空行、分部打断+撤销恢复、冻结修复回归、树行距、角标菜单、函数面板、cube/power/^ 求值）。
- 全量回归 v23~v52 全绿（v36 因默认空行改为整体赋值造数据；v42 函数断言 power→cube）。

## 三、打包

- PyInstaller onefile：`python -m PyInstaller --noconfirm --distpath dist --workpath build "build/工程量计算书.spec"`。
- 产物：`CalcBook-v3.7.4.exe`，启动 16s 存活、错误日志零新增。

## 四、Git

- 本地提交，**未推送任何远端**（Gitee/GitHub 零写入）。
- 另：V3.7.2 的正式发布（等用户删 Gitee 镜像后合 main + tag + Release）仍挂起。
