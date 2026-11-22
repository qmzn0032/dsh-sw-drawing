# dsh-sw-drawing

把 **SOLIDWORKS 2024** 接进 **DeepSeek Harness (DSH)**。仓库里是一个 DSH 插件包（bundle），
装完 agent 就能通过 17 个 `sw_*` 工具直接读模型、建模、改材料、出 GB 工程图、导 PDF。

> 只在一台机器上真机校准过：Windows 11 + SOLIDWORKS 2024 SP0.1（interop 32.0.1.4）+ Python 3.13。
> 换 SW 小版本可能个别接口对不上，见「已知限制」。

## 能做什么

- **建模**：草图（矩形/圆/线/多段线/弧）、拉伸、切除、圆角/倒角、抽壳、异型孔向导（ISO 标准表）
- **阵列**：线性（一维/二维）、圆周
- **查询**：模型概况、质量属性、特征树
- **出图**：从零件一键出 GB 三视图，自动选比例，加尺寸/公差/孔标注，存 `.SLDDRW` 并导 PDF
- **导出**：PDF / DWG / DXF / BMP
- **字体**：工程图字体按「文件属性」批量换（标题栏那几十上百条文字自动跟着变）

## 前置条件

| 项 | 说明 |
|---|---|
| Windows | 已验证 Win11 |
| SOLIDWORKS 2024 | **必须先打开**。工具只用 `GetActiveObject` 连已经运行的实例，不会自己拉起来 |
| Python 3 + pywin32 | `pip install pywin32`（launcher 会自己找解释器，见下） |
| DSH | 安装本地插件需要相应权限（会弹确认） |
| 汉仪长仿宋体（可选） | 工程图模板默认用 GB 长仿宋。这是**商业字体**，仓库不附带，请自行取得授权 |

## 安装

```
1. 下载/克隆本仓库到任意目录
2. 双击 install.cmd          ← 找 Python、查 pywin32、把本机路径写进 patch
3. DSH: 设置 → 插件 → 添加插件 → 选 packages\dsh-sw-drawing
4. 重启 DSH
```

第 3 步如果没有图形入口，也可以直接把这句话发给 DSH：

```
安装这个本地插件目录：<仓库绝对路径>\packages\dsh-sw-drawing
```

install.cmd 会告诉你确切的路径。它只写一个文件：`packages\dsh-sw-drawing\cordis.patch.yml`
（这个文件在 `.gitignore` 里，因为里面是本机绝对路径）。

## 验证

1. 打开 SOLIDWORKS，随便开一个零件
2. 在 DSH 里问：「用 sw_status 看看 SW 状态」
3. 返回 `running: true` 且 `document` 是你的零件名 → 通了

## 工具

| 工具 | 读写 | 作用 |
|---|---|---|
| `sw_status` | 读 | SW 是否在跑、当前文档 |
| `sw_part_summary` | 读 | 名称/路径/配置/特征树/材料 |
| `sw_mass_properties` | 读 | 体积、质量、表面积、重心、包围盒 |
| `sw_rebuild` | 写 | 重建并返回特征树 |
| `sw_set_material` | 写 | 改当前配置材料 |
| `sw_new` | 写 | 按模板新建零件/装配体 |
| `sw_build` | 写 | 草图 + 拉伸/切除批量建模 |
| `sw_face_sketch` | 写 | 实体面上画草图并切除/加料 |
| `sw_fillet` | 写 | 圆角/倒角（射线选边） |
| `sw_shell` | 写 | 抽壳 |
| `sw_hole` | 写 | 通孔/盲孔/沉头孔 |
| `sw_hole_wizard` | 写 | 异型孔向导：柱形/锥形沉头、螺栓通孔、螺纹孔（尺寸取 ISO 表） |
| `sw_pattern` | 写 | 线性阵列 |
| `sw_circular_pattern` | 写 | 圆周阵列 |
| `sw_rename` | 写 | 特征语义化命名 |
| `sw_drawing_build` | 写 | 出 GB 工程图（A4–A0、自动比例、尺寸/公差/孔标注、PDF） |
| `sw_export` | 写 | 导出 PDF/DWG/DXF/BMP |

## 写操作的确认框

13 个写工具在动手前会弹一个 Windows 确认框（工具名 + 参数），**45 秒不点按拒绝**。
勾「本次启动内全部允许」后，本 DSH 进程内不再逐次询问，重启恢复。

45 秒不是随手定的：DSH 单次工具调用默认 60 秒超时，确认框等得比它久就会出现
「DSH 记失败、SW 却真改了」的错位。

## 目录

```
install.cmd / install.ps1      安装脚本（生成本机 patch + 打印下一步）
README.md                      这份说明
CHANGELOG.md                   版本变更记录
LICENSE                        MIT 许可证
packages/dsh-sw-drawing/       插件本体（DSH bundle）
  py/                          MCP 服务 + SOLIDWORKS COM 封装
  docs/DETAIL.md               插件详情页
templates/                     A3 GB 图纸格式（.slddrt）
tools/                         独立小工具：一键换字体
```

## 已知限制

- **SW 必须已经在运行**，且与 DSH 同一用户、同一完整性级别（UAC 权限不同会连不上）。
- 只在 SW 2024 SP0.1 校准过；其它版本个别接口序号可能不同。
- 本机不可用的字面接口（`Save3`、`InsertCosmeticThread3`、`ISketchManager.InsertSketchText`）
  已用替代方案绕开，代价是螺纹孔表现为「切除」而不是真正的 HoleWzd。
- 工程图字体依赖本机装了同名族字体；没装时 SW 会**静默 fallback**，字形直接变。
- 视图手动排布关闭（`IView.Position` 在本机行为异常），靠自动选比例避免重叠。
- 插件是 **link 安装**：移动仓库目录后要在 DSH 里重新添加一次。
- `templates\A3-GB-汉仪长仿宋体.slddrt` 需要在 SW 的「图纸属性 → 图纸格式」里手工指一次。

## 许可

代码以 **MIT 许可**发布，见 [LICENSE](LICENSE)：可以自由使用、修改、再分发、商用，
保留版权声明即可。

第三方资产**不**在本许可范围内：

- SOLIDWORKS 及其 API 归 Dassault Systèmes，使用前请自备合法授权；仓库不含任何 Dassault 的二进制文件。
- GB 长仿宋体等字体是商业字体，**仓库不附带**，请自行取得授权。
- `templates/` 下的图纸格式模板由本仓库作者制作，随 MIT 一并提供。

变更记录见 [CHANGELOG.md](CHANGELOG.md)。
