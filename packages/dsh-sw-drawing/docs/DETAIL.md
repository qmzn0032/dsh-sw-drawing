# SOLIDWORKS 桥接（sw_bridge）

把 SOLIDWORKS 2024 接进 DSH 的 MCP 服务。Agent 通过 17 个 `sw_*` 工具直接操作 SW：读模型、建模、
改材料、出 GB 工程图、导 PDF。

## 先决条件

| 项 | 要求 |
|---|---|
| Windows | 已验证 Win11 |
| SOLIDWORKS | 2024 SP0.1（interop 32.0.1.4），**必须已经在运行**，不自动启动 |
| Python | 3.10+，且有 `pywin32` |
| DSH | 安装插件需要相应权限 |

工具只用 `GetActiveObject` 连接已经跑着的 SW，所以「SW 没开」是一切报错的第一个怀疑对象。

## 工具

| 工具 | 读/写 | 作用 |
|---|---|---|
| `sw_status` | 读 | SW 是否在跑、当前文档 |
| `sw_part_summary` | 读 | 名称/路径/配置/特征树/材料 |
| `sw_mass_properties` | 读 | 体积、质量、表面积、重心、包围盒 |
| `sw_rebuild` | 写 | 重建并返回特征树 |
| `sw_set_material` | 写 | 改当前配置的材料 |
| `sw_new` | 写 | 按模板新建零件/装配体 |
| `sw_build` | 写 | 草图 + 拉伸/切除的批量建模 |
| `sw_face_sketch` | 写 | 在实体面上画草图并切除/加料 |
| `sw_fillet` | 写 | 圆角/倒角（射线选边） |
| `sw_shell` | 写 | 抽壳 |
| `sw_hole` | 写 | 通孔/盲孔/沉头孔（几何实现） |
| `sw_hole_wizard` | 写 | 异型孔向导：柱形沉头/锥形沉头/螺栓通孔/螺纹孔，尺寸取 ISO 标准表 |
| `sw_pattern` | 写 | 线性阵列（一维/二维） |
| `sw_circular_pattern` | 写 | 圆周阵列 |
| `sw_rename` | 写 | 给特征起语义名 |
| `sw_drawing_build` | 写 | 从零件出 GB 工程图（A4–A0、自动选比例、尺寸/公差/孔标注、存 .SLDDRW + PDF） |
| `sw_export` | 写 | 导出 PDF/DWG/DXF/BMP |

## 写操作的确认

13 个写工具在真正下刀之前会弹一个 Windows 确认框（显示工具名和参数），45 s 不点按拒绝。
勾选「本次启动内全部允许」后，本次 DSH 进程内不再逐次询问（重启后恢复询问）。

之所以是 45 s：DSH 的 `toolCallTimeoutMs` 默认 60 s，弹窗等得比它久会出现「DSH 记失败、SW 真改了」的错位。

## 已知限制

- 只在 SW 2024 SP0.1 上校准过；其它小版本的接口序号/枚举可能对不上。
- 少数字面接口在本机不可用（`Save3`、`InsertCosmeticThread3`、`ISketchManager.InsertSketchText`），
  已用替代路子绕开，代价是螺纹孔是「切除」而非真正的 HoleWzd。
- 工程图字体依赖本机装了同名族字体（默认「汉仪长仿宋体」），没装会静默 fallback。
- 手动视图排布关闭（`IView.Position` 在本机行为异常），靠自动选比例避免重叠。
- 没有自动化测试，全部靠真机对账（体积/包围盒回读）。
