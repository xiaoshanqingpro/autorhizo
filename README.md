# WinRHIZO 根系扫描半自动批量分析工具

用于 WinRHIZO Pro 2017a（32 位）根系扫描图像的半自动批量分析辅助工具。

## 功能
- **ROI 自动检测**：OpenCV 自动检测扫描图片中的根系包围盒，生成预览图（`_ROI预览/`）
- **自动画框**：把检测到的 ROI 框绘制到 WinRHIZO 中（画框即触发分析）
- **自动缩放**：控制 WinRHIZO 缩放到目标倍数（如 1/8.0，可看到完整根系）
- **数据跟踪**：监测 WinRHIZO 分析结果自动保存到数据文件，检测到新行即推进
- **进度管理**：队列显示、已处理计数、断点续跑（已有数据的图片自动跳过）
- **汇总**：把数据文件解析为按 `F/G + 样地 + 重复` 分组的 CSV/XLSX 汇总表（多次分析全部保留，含分析序号）

## 环境依赖
- Python 3.10+
- numpy, opencv-python, pillow, openpyxl, uiautomation（见 `requirements.txt`）

## 使用
1. 在 `config.json` 设置日期、图片目录、数据文件路径
2. 双击 `启动.bat`（或 `python batch_analyze.py`）打开工具窗口
3. 点"开始"，按队列提示逐张处理：
   - 在 WinRHIZO 中用"获取图片"打开当前图（**此步需手动**，见下方限制）
   - 点工具的"完成"按钮（或按注册的热键 F12/F11/F10）
   - 工具自动缩放、绘制 ROI 框、等待分析结果
   - 如弹出"数据未保存"对话框，手动选 Open one → 选择数据文件 → 确定
4. 全部完成后生成 `_汇总.csv` / `_汇总.xlsx`

## 已知限制（重要）
- **自动打开图片不可靠**：WinRHIZO 2017a 的现代文件对话框拒绝程序化输入（UIA/键盘均被封锁），需要手动获取图片
- **ROI 屏幕映射需校准**：画框的屏幕坐标由 `config.json` 的 `view_ox_off` / `view_oy_off` / `target_zoom` 校准，不同窗口尺寸/布局需重新校准
- **数据保存依赖数据文件打开**：分析结果自动保存到 WinRHIZO 中打开的数据文件；未打开时会出现"数据未保存"对话框，需手动处理

## 文件
- `batch_analyze.py` — 主程序
- `config.json` — 配置（日期/路径/ROI 检测参数/映射校准）
- `requirements.txt` — Python 依赖
- `启动.bat` — Windows 双击启动
