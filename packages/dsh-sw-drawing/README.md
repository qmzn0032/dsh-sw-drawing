> 说明：这份文档是从开发工作区搬过来的，里面的 `_plugins\dsh-sw-drawing\` 对应发布包里的
> `packages\dsh-sw-drawing\`，`py\` 下的文件名一致。真机校准表保留原样，方便对照。
# @local/dsh-sw-drawing

SOLIDWORKS 的 DSH 出图插件（M1）。内核是纯 Python 脚本，外壳是可安装 bundle。

## 是什么关系

```
@local/dsh-sw-drawing          ← bundle：plugin_manager 装/卸
├── package.json               dsh.bundle.patch
├── cordis.patch.yml           插入 sw_bridge 这个 MCP 服务
└── py/                        ← 内核，同一份代码能脱离 DSH 单跑
    ├── sw_com.py              建模层（会话/单位/审批/质量属性/草图拉伸切除）
    ├── sw_features.py         特征层（选面/面草图/改名/线性+圆周阵列/抽壳/圆角/倒角/异型孔）
    ├── sw_drawing.py          GB 出图：建图/图幅/三视图/选比例/导出
    ├── gb_rules.json          GB 规则表（图幅/比例/字体/公差/粗糙度/标题栏）
    └── mcp_sw_server.py       MCP 工具面（17 个工具，写操作全过审批）
```

`dsh-sw-bridge/` 是历史目录（含浮窗版 app.py），**不再是被 DSH 调用的那份**，改动请落在本目录。

## 工具

| 工具 | 读写 | 说明 |
|---|---|---|
| `sw_status` / `sw_part_summary` / `sw_mass_properties` / `sw_rebuild` | 读 | 会话、特征树、质量属性、重建 |
| `sw_set_material` / `sw_new` / `sw_build` | 写 | 材料、新建零件、草图→拉伸/切除 |
| **`sw_drawing_build`** | 写 | 零件 → GB 图框 + 第一角三视图 + 自动选比例 → `.SLDDRW` + PDF；`font` 参数按**文档属性**设字体（默认汉仪长仿宋体） |
| **`sw_export`** | 写 | 活动文档 → `.pdf` / `.dwg` / `.dxf` / `.bmp` |
| **`sw_face_sketch`** | 写 | 在实体面上画草图（P0） |
| **`sw_rename`** | 写 | 特征改名（P0） |
| **`sw_pattern`** | 写 | 线性阵列（P0） |
| **`sw_fillet`** | 写 | 圆角 `kind=fillet` / 等距倒角 `kind=chamfer`（P1） |
| **`sw_shell`** | 写 | 抽壳：`faces` 给开口面，留空=闭合空腔（P1） |
| **`sw_circular_pattern`** | 写 | 圆周阵列：`angleDeg` 默认 360 整圈，`axis`=x/y/z 自动建参考轴（P1） |
| **`sw_hole`** | 写 | 打孔：通孔 / 盲孔 / 柱形沉头孔（面草图+切除的几何等价实现，P1） |
| **`sw_hole_wizard`** | 写 | **异型孔向导**：柱形沉头 / 锥形沉头 / 螺栓通孔 / 螺纹孔，尺寸取自 SW 的 ISO 标准表（P1，2026-11-21 打通） |

## 命令行单跑（不经 DSH）

```powershell
cd _plugins\dsh-sw-drawing\py
$env:PYTHONIOENCODING='utf-8'
python sw_drawing.py "C:\path\part.SLDPRT" "C:\out"
```

## 真机校准（SW2024 SP0.1 / interop 32.0.1.4，2026-11-20）

| # | 事实 |
|---|---|
| 1 | `sw.NewDocument(gb_a3.drwdot)` 出来的就是 **GB 图框**（`a3 - gb.slddrt`）。用户现有 `功能牌1.SLDDRW` 用的是 ISO 图框。 |
| 2 | `Create1stAngleViews(path)` 0.06 s 建出 3 视图，但**按 1:1 摆**，长零件会重叠。解法：先建一张临时图纸量尺寸 → 定比例 → 关掉 → 再建正式图纸、**先设比例后插视图**，让 SW 自己按最终尺寸排。 |
| 3 | `Extension.SaveAs(..., None, ...)` 存 `.slddrw` 报「类型不匹配」，必须传 `VARIANT(VT_DISPATCH, None)`；导 PDF 第四参要真 `IExportPdfData`。 |
| 4 | `GetViews()` 返回 tuple 套一层，元素是 `PyIDispatch`，要再 `Dispatch` 一次。 |
| 5 | `sheet.GetProperties()` 7 个值：下标 2/3 是比例分子分母，**4 是第一角标志（GB=1，ISO=0）**。 |
| 6 | `IView.GetOutline` / `GetDisplayDimensionCount` 在本机是**属性**不是方法。 |
| 7 | 抽壳是 `IModelDoc2.InsertFeatureShell(Thickness, Outward)`，**不是** FeatureManager 的方法；返回值成功也是 `None`，要用特征树差集按类型名 `Shell` 找。 |
| 8 | 圆周阵列 `FeatureCircularPattern5` 的 `Spacing` 是**弧度**，语义随 `EqualSpacing` 变：`False`=每份间隔角，`True`=总角度（360° 按 360/N，其它按 总角/(N-1)）。 |
| 9 | `FlipDirection=False` 是顺时针。判定角度**只能看面心**——4 个不重叠的孔，间隔 90° 和 30° 体积一模一样。 |
| 10 | 参考轴用两个基准面的交线：`Select2` 两个面 → `InsertAxis2(True)`。x=上视∩前视、y=右视∩前视、z=右视∩上视，三种都实测可用。 |
| 11 | 沙箱 workspace-write 下 `GetActiveObject` 报 `-2147221021`，连 COM 必须提权；`connect(launch=True)` 用 `Popen` 拉起再轮询 ROT，约 2.8 s。 |
| 12 | ~~异型孔 `HoleWizard5` 一律返回 `None`~~ —— **2026-11-21 已打通，根因两条**：① 选的是**草图特征**，而异型孔要的是**草图点对象**（`ISketchPoint.Select2(Append, 0)`，且不要给选择标记，官方 Remarks 原话是 *preselect the sketch points and do not specify any selection marks*）；选错对象 SW 静默返回 `None`。② `FastenerTypeIndex` 传了 0，而它必须是该标准下的具体序号：ISO 内六角圆柱头 139 / 沉头螺钉 140 / 螺栓通孔(ISO 273) 144 / 内螺纹 147。 |
| 13 | 孔改用「面草图圆 + `FeatureCut4`」拼：Ø10 通孔、Ø10+Ø18×6 沉头、Ø10 盲孔深 4 体积全对。沉头 = 先打穿再盲切一层，几何等价。 |
| 14 | **`IDimension.Value` 返回文档单位（mm）不是米**；`ISldWorks` 上的设置接口叫 `GetUserPreferenceStringValue` / `SetUserPreferenceStringValue`，读出来自带 `\lang\<语言>`。 |
| 15 | **副本阶梯法**量各特征体积：拷贝原件 → 打开副本 → 从树底往上 `feat.Select2(False,0)` + `doc.EditSuppress` → 每步重建读体积。`SetSuppression2` 在晚期绑定下不认。 |
| 16 | 本版 interop 的 `ISketchManager` **没有 `InsertSketchText`**，`IModelDoc2` 也没有 `EditPaste`（`EditCopy` 有），草图文字只能读不能建，所以文字雕刻类特征无法用 API 复刻。 |
| 17 | **二维线性阵列**第二方向的选择标记是 `Select2(Append, 2)`（第一方向 1、种子 4）。实测 3×20mm + 2×15mm → 6 孔，体积和 6 个孔位全对。 |
| 18 | 异型孔 `HoleWizard5` 27 参签名（反射本机 interop）：`(GenericHoleType, StandardIndex, FastenerTypeIndex, SSize, EndType, Diameter, Depth, Length, Value1..12, ThreadClass, RevDir, FeatureScope, AutoSelect, AssemblyFeatureScope, AutoSelectComponents, PropagateFeatureToParts)`。单位米、角度弧度；**Value1..12 一律传 0**（传 -1 时 clearance 会静默返回 None）。 |
| 19 | 异型孔实测体积（100×60×10 板，ISO 标准表）：M8 柱形沉头 **1597.3302**（Ø8.4 通 + Ø15×8.6 沉头）、M8 锥形沉头 **966.0774**（Ø17.92 / 90°，回读 1.5708 rad）、M10×2 螺栓通孔 **1731.8030**（Ø10.5，ISO 273 普通配合）、M8 螺纹孔盲 12 底孔 Ø6.8 **363.1681**。 |
| 20 | **螺纹孔的坑**：`GenericHoleType=4`(Tap) + `FastenerTypeIndex=147` 时，`Diameter` 给 6.8mm 实际切出来是**标称 Ø8.0**（挖 502.6548 = π·4²×10），`SSize` 换成 `"M8x1.25"` 无效，`Diameter=0` 更会退化成 Ø50.8（=2 英寸 API 默认值）。所以 `sw_hole_wizard` 的 `tap` 改成「按 ISO 2306 底孔切除 + 试加装饰螺纹线」，体积才和界面里建的一致（363.1681）。 |
| 21 | `InsertCosmeticThread3(Standard, StandardType, Size, Diameter, EndType, Depth, Note)` 在晚期绑定下**返回 None**（圆柱面已选中），`SelectionManager.CreateSelectData()` 报 `DISP_E_MEMBERNOTFOUND`，`SelectByID2` 的 Callout 传 `None` 报 `DISP_E_TYPEMISMATCH`（要 `VARIANT(VT_DISPATCH, None)`）。装饰螺纹线暂未加上，孔本身正确。 |
| 22 | **改字体走「文件属性」最省事**：`IModelDocExtension.SetUserPreferenceTextFormat(pref, swDetailingNoOptionSpecified=0, tf)`，pref 取 `swDetailingNoteTextFormat=0` / `DimensionTextFormat=1` / `AnnotationTextFormat=19`。先从 `GetUserPreferenceTextFormat(...)` 拿到原对象、只改 `TypeFaceName` 再写回，字号/粗斜体不会被带跑（标题栏各格字高不同，这点必须保住）。 |
| 23 | 图纸格式里那 80 多个标题栏文字 `INote.GetUseDocTextFormat()` 全是 **True**，即跟着文档属性走——改完文档属性它们一起变，不用逐条改。而 `SelectAll` 拿回来的这些对象在本机 pywin32 上 `GetAnnotation` / `GetTextFormat` 都报「找不到成员」，逐条改这条路走不通（78 条 `swSelSKETCHSEGS` + 82 条 `swSelNOTES`）。 |
| 24 | 只有视图里的尺寸/孔标注带**本地格式**，走 `IAnnotation.GetTextFormat(i)` → `ITextFormat.TypeFaceName` → `SetTextFormat(i, UseDoc=False, tf)`；`GetUseDocTextFormat` 要带 `Index` 参数，不带报「非选择性的参数」。 |
| 25 | 「汉仪长仿宋体」在字体里的族名是 `HYSWLongFangSong`（`C:\Windows\Fonts\hyswlongfangsong.ttf` 与用户目录 `汉仪长仿宋体.ttf` 同名同族、不同版本）。SW 只认族名，装两份同族文件时用哪份由 Windows 决定（2026-11-21 实测本机解析到 **2023 版**，见第 28 条）；PDF 里嵌的 `XXXXXX+HYSWLongFangSong` 就是它，可以用 `/BaseFont` 反查一张图到底用的哪个字体。 |
| 26 | **图纸格式 / 工程图模板怎么存**：`ISheet.SaveFormat(路径.slddrt)` 把当前图纸格式导出（之前先 `ActivateSheet`），存出来 52 130 字节，图框+标题栏全在；换格式用 `IDrawingDoc.SetupSheet6(Name, PaperSize, **TemplateIn=12**(swDwgTemplateCustom), Scale1, Scale2, FirstAngle, **TemplateName=路径**, Width, Height, PropertyViewName, RemoveModifiedNotes, ZoneL, ZoneR, ZoneT, ZoneB, ZoneRow, ZoneCol)`，回读 `ISheet.GetTemplateName()` 确认换上了；再 `Extension.SaveAs(xxx.drwdot, 0, 1, null, ...)` 就能存成工程图模板（本机 67 KB → 83 KB）。 |
| 27 | **存盘要用 Save2，不是 Save3**：`IModelDoc2.Save3(1)` 在本机抛 `(-2147352561, '非选择性的参数')`，`Save2(1)` 正常，但**返回值是 0**，不能当成功标志——判定口径是「文件 mtime 变了 + `GetSaveFlag` 归零 + 大小 > 1 KB」。另外 `Extension.SaveAs` 不能覆盖 SW 自己正打开的那个文件（它先 `os.remove`，而 SW 占着句柄 → `WinError 32`），覆盖已开文档只能 `Save2`。 |
| 28 | **两份汉仪长仿宋体到底差在哪**（解析 ttf 实测）：字符集完全相同（各 8 633 字），但 426 个字形轮廓不同，集中在 ASCII 数字/字母；同一串 `0123456789 M8 S15 Ø10 2.5`，2017 版墨宽 795 px、2023 版 771 px，**2017 版宽约 3%**。判断「Windows 实际喂给程序的是哪份」的黑盒办法：GDI+ 里用 `System.Drawing.FontFamily("汉仪长仿宋体")` 渲一遍，再用 `PrivateFontCollection` 只加单份文件各渲一遍比像素——本机系统渲染与 2023 文件**逐像素 0 差异**、与 2017 文件差 14 922 像素，即**当前用的是 2023 版（HKLM 系统级那份）**。 |
| 30 | **审批超时必须短于 DSH 的 60 s**：`dsh-mcp-client` 的 `toolCallTimeoutMs` 默认 **60 000 ms**。确认框原来等 120 s，于是出现「DSH 先超时报失败、用户后点允许 → SW 真改了但记录是失败」的错位。现在默认 45 s、硬上限 55 s（`approval_dialog.resolve_timeout`）。 |
| 31 | **MCP 主循环原来是单线程 `for line in sys.stdin`**，确认框一弹整个服务端失联（`ping` / `tools/list` 全无响应）。改成：主线程只读 stdin 并处理轻量请求，`tools/call` 丢给一条工作线程串行执行；实测写操作进行中 `ping` **20 ms** 返回。 |
| 32 | **确认框改由子进程弹**（Tk 不进 MCP 进程）：父进程 `subprocess.run(..., stdout=DEVNULL)`，靠退出码传结果（0 允许 / 1 拒绝 / 2 允许并本次启动放行）。子进程 stdout 必须 DEVNULL，否则 Tk 的任何输出都会污染 JSON-RPC 通道；超时用 `timeout=` 硬杀，不留幽灵窗口。 |
| 33 | **入口改成 `py\launch_server.cmd`**：launcher 按 `SW_BRIDGE_PYTHON` → hermes venv → `py -3` → `python` 的顺序找解释器，找不到就往 stderr 打原因并退出码 2。这样 hermes 升级/搬家不会再让整个插件失联；patch 里只剩 bundle 自身路径一条本机事实。 |
| 34 | **一次误判记录**：调试时看到确认框在 1 s 内返回「允许」，一度判定为幽灵鼠标事件。实际是**用户自己点的**——弹窗总出现在指针正下方，用户看到就点了。教训：自动化测试里有人机交互弹窗时，rc=0 不能当成「程序自己批的」。顺手加固：确认框改成 Label 自绘按钮，按下+抬起都要在按钮内、0.35 s 冷却吞双击余波、窗口不再压指针、不绑键盘事件。 || 29 | 落地位置：模板在 `模板\A3-GB-汉仪长仿宋体.slddrt` + `.drwdot`（A3/第一角/比例 1:1/文档属性字体已设好，`drwdot` 里按绝对路径引用同目录那份 `slddrt`）；一键换字体在 `工具\`（`.cmd` 拖放 + `font_switch.py`，`--check/--dry-run/--no-save/--no-backup/--font`）。 |

## 已知缺口

- **视图排布松散**：`IView.Position` 赋值在本机行为异常（赋 `(0.15,0.20)` 回读 `(0.0,0.15)`，且会把三个视图压到同一点），所以手动紧凑排布关掉了（`relayout=False`），只靠自动选比例避免重叠。图能用，但不像手排的那么紧凑。
- **没有尺寸、标注、标题栏数据**：M2/M3 才做。现在标题栏只有图框自带的空壳（材料/重量/比例会跟着模型变）。
- **只有三视图**：轴测、剖视、局部放大、断裂都没接。
- **比例按视图块估算**：没有考虑标题栏占位，极端长宽比零件可能选得偏保守。

## 建模侧缺口（三维）

- ~~**异型孔向导用不了**~~ —— **2026-11-21 打通**，见真机校准第 12/18/19 条，工具是 `sw_hole_wizard`（柱形沉头 / 锥形沉头 / 螺栓通孔 / 螺纹孔）。剩下的缺口只有一条：**装饰螺纹线加不上**（`InsertCosmeticThread3` 在晚期绑定下返回 `None`），所以螺纹孔的树里是「切除-拉伸」而不是 HoleWzd，工程图不会自动出螺纹标注；孔径/深度按 ISO 2306 底孔，体积正确。
- **文字雕刻复刻不了**：本版 interop 既没有 `ISketchManager.InsertSketchText`，`IModelDoc2` 上也没有 `EditPaste`（`EditCopy` 有，复制得出来但粘不进去）。草图文字只能读（`ISketch.GetSketchTextSegments()` / `ISketchText.Text`）。要做得把字形轮廓转 DXF 再 `InsertDwgOrDxfFile` 导入。
- **镜像、扫描、放样、旋转、拔模** 未接（`InsertMirrorFeature2` / `FeatureSweep2` / `FeatureLoft2` / `FeatureRevolve2` 的签名已反射到）。
- **变壁厚抽壳**（`InsertFeatureShellAddThickness`）未接，只支持等壁厚。
- **边线选择只能靠射线打**：`body.GetEdges()` 返回的边对象在本机取不到任何成员，所以圆角/倒角要用户给 point+dir。
- **同时只支持一个实体特征**：没有「按面选体」的概念，多实体零件会选到哪个面取决于 find_face 的最近点逻辑。


## 暂缓项

### 2026-11-21 更新：异型孔向导已打通，从暂缓项里移出

用户 2026-11-20 让先搁置**草图文字**和**异型孔向导**两项；2026-11-21 用户要求「异型孔既然有接口就应该能用」，
查官方 Remarks（Wayback 上的 2018 API Help）后定位到两条根因，已实现为 `sw_hole_wizard` 并真机对账通过。
仍然搁置的只剩 **草图文字**。

### 草图文字（仍搁置）

- 现状：本版 interop 无 `ISketchManager.InsertSketchText`，`IModelDoc2` 无 `EditPaste`（`EditCopy` 有），
  手写的合规 R12 DXF 也能被 `GetImportFileData` 解析但 `InsertDwgOrDxfFile` 不建特征。
- 影响：原件 `功能牌1.SLDPRT` 里那 4.1132 mm³ 的文字雕刻（草稿 41.132 mm² × 0.1 深）复刻不了。
- 重启条件（任一即可）：① 升级到确有 `InsertSketchText` 的 SW/interop；② 用户给一份 SW 自己认可的 DXF；
  ③ 用户提供原件里文字的字体/字号，改用 `CreateLine` 画字形轮廓（只能近似，不保证 4.1132 mm³ 体积一致）。
- 不重启时的验收口径：**可复现特征体积一致**（复刻件 4463.0341 mm³ vs 原件同阶段 4463.0307 mm³，差 0.0034 mm³），
  文字雕刻不复刻。
