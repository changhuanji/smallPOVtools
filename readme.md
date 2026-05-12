# 多功能工具箱 (v2.7)

基于 Python `tkinter` 开发的桌面 GUI 应用程序，集成了 PPT 批量处理、图片处理、透明视频生成与视频剪辑辅助工具。采用模块化设计，共 6 大功能模块。

## 目录结构

```
main.py              # 主程序入口
ppt_tools.py         # 功能1 - PPT 批量生成
ppt_color_tools.py   # 功能2 - PPT 改色与清理
ppt_img_tools.py     # 功能3 - PPT 逐页导出与图片处理
video_tools.py       # 功能4 - 透明动画视频生成
fps_tool.py          # 功能5 - 视频变速计算器
web_tools.py         # 功能6 - 地图与 API 工具
ffmpeg.exe           # [必须] 视频编码核心组件（需自行下载放入目录）
```

## 运行环境

**系统要求：**
- Windows 10/11（依赖 COM 组件调用 PowerPoint，仅支持 Windows）
- 必须安装 **Microsoft PowerPoint**
- 必须下载 **ffmpeg.exe** 放置在脚本同级目录（用于功能 4）

**Python 依赖：**

```bash
pip install python-pptx pywin32 opencv-python numpy
```

## 功能模块

### 1. PPT 批量生成器 (`ppt_tools.py`)

读取 PPT 模板的第一页，根据 TXT 数据文件的行数复制页面，替换指定占位符。

- 选择 `.pptx` 模板（在文本框内预留占位符，如 `{name}`）
- 选择 `.txt` 数据文件（每一行对应生成一页 PPT）
- 输入占位符文本，点击生成

### 2. PPT 改色与清理工具 (`ppt_color_tools.py`)

对 PPT 进行全局样式修改和内容清洗。

- **字体改色**：输入 RGB 值，将所有文本修改为指定颜色
- **背景改色**：修改幻灯片背景为纯色或保留透明
- **去除空格**：一键删除所有文本中的空格和换行符
- **删除空框**：自动检测并删除无可见文本的空白文本框

### 3. PPT 导出与图片处理 (`ppt_img_tools.py`)

包含三个子功能面板（Tab）：

**Tab 1 - PPT 逐页透明导出**
- 调用 PowerPoint 原生接口将幻灯片导出为高精度 PNG
- 支持透明背景、DPI 设置（72-600，默认 288）、强制宽高比（16:9 / 4:3）
- 实验性功能：使用 OpenCV 强制裁切/填充到指定像素（如 3840×2160）

**Tab 2 - 批量裁切/扩展**
- 正数向内裁切，负数向外扩展画布（填充透明像素）
- JPG 等无透明通道格式自动转换为 PNG

**Tab 3 - 批量图片去底**
- 色度键抠图：根据目标 RGB 颜色和容差值将背景转为透明

### 4. 透明动画视频生成器 (`video_tools.py`)

将静态图片转换为含 Alpha 通道的运动视频。

- 支持 **ProRes 4444** (CPU) 和 **HEVC** (NVIDIA GPU) 两种编码
- 生成 `.mov` 含透明通道，可直接拖入 DaVinci Resolve 或 Premiere Pro
- 可配置运动角度、距离、速度、起始位置（9 宫格预设 + 自定义坐标）
- 支持 1080P、4K 及自定义分辨率

### 5. 视频变速计算器 (`fps_tool.py`)

视频剪辑辅助工具，计算素材变速比率。

- 仿非编软件的 H:M:S:F 时间码输入框，输入两位数字自动跳转
- 输出变速百分比并自动复制到剪贴板
- 支持独立设置时间线与素材帧率
- 支持窗口置顶，方便在剪辑软件上方悬浮使用
- 支持反向变速（负值百分比）

### 6. 地图与 API 工具 (`web_tools.py`)

- **地图导航**：高德、百度、腾讯地图入口，强制调用 Edge 打开
- **API 预留**：预留了天气、汇率、快递查询等 API 接口 UI 框架

## 常见问题

1. **功能 4 报错"未检测到 ffmpeg"**
   - 下载 `ffmpeg.exe`（通常位于 FFmpeg 的 bin 文件夹内），复制到 `.py` 文件同级目录

2. **PPT 导出分辨率高度不对（如 3840×2199）**
   - 在功能 3 Tab 1 中启用实验性功能，设置目标宽高并选择居中锚点

3. **达芬奇无法识别 MOV 透明通道**
   - 使用 CPU (ProRes) 模式导出。在达芬奇媒体池中右键素材 → Clip Attributes → Alpha Mode 设为 Straight 或 Premultiplied

4. **程序无响应**
   - 视频渲染和大量 PPT 导出为耗时操作，程序使用多线程避免界面卡死，请耐心等待进度条走完

## 技术栈

- **语言**: Python 3.11+
- **GUI**: Tkinter
- **核心库**: python-pptx, opencv-python, pywin32, numpy, subprocess
