# Windows Test MCP 服务

用于Windows桌面程序自动化测试的MCP(Model Context Protocol)服务，提供截图、键盘和鼠标模拟操作功能。

## 功能特性

### 截图工具
- `screenshot_capture` - 截取整个屏幕
- `screenshot_region` - 截取指定区域

### 键盘工具
- `keyboard_press` - 按键点击（按下+释放）
- `keyboard_down` - 按下按键不放
- `keyboard_up` - 释放按键
- `keyboard_type` - 输入文本字符串

支持的按键包括：
- 字母数字：`a-z`, `0-9`
- 功能键：`f1` - `f12`
- 特殊键：`enter`, `space`, `tab`, `esc`, `backspace`, `delete`
- 方向键：`up`, `down`, `left`, `right`
- 修饰键：`ctrl`, `shift`, `alt`, `win`

### 鼠标工具
- `mouse_move` - 移动鼠标到指定位置
- `mouse_click` - 点击鼠标按钮
- `mouse_down` - 按下鼠标按钮
- `mouse_up` - 释放鼠标按钮
- `mouse_scroll` - 滚动鼠标滚轮
- `mouse_get_position` - 获取当前鼠标位置
- `get_screen_size` - 获取屏幕尺寸

支持的鼠标按钮：`left`, `right`, `middle`, `mouse4`, `mouse5`

## 安装

### 环境要求
- Python >= 3.10
- Windows 操作系统

### 安装步骤

```bash
# 克隆或下载项目
cd windows-test-mcp

# 创建虚拟环境
python -m venv venv

# 激活虚拟环境
venv\Scripts\activate

# 安装依赖
pip install -e ".[dev]"
```

## 使用

### 作为MCP服务器运行

#### 使用 stdio 传输（默认）
```bash
python -m windows_test_mcp.server
```

#### 使用 SSE 传输
```bash
python -m windows_test_mcp.server --transport sse --port 8000
```

### Claude Desktop 配置

在 `claude_desktop_config.json` 中添加：

```json
{
  "mcpServers": {
    "windows-test": {
      "command": "python",
      "args": ["-m", "windows_test_mcp.server"],
      "cwd": "D:\\path\\to\\windows-test-mcp",
      "env": {
        "PYTHONPATH": "D:\\path\\to\\windows-test-mcp\\src"
      }
    }
  }
}
```

## 工具使用示例

### 截图
```python
# 截取整个屏幕
screenshot_capture()

# 截取指定区域 (x=100, y=100, width=800, height=600)
screenshot_region(x=100, y=100, width=800, height=600)
```

### 键盘操作
```python
# 按 Enter 键
keyboard_press(key="enter")

# 输入文本
keyboard_type(text="Hello World", interval=0.01)

# 组合键: Ctrl+C
keyboard_down(key="ctrl")
keyboard_press(key="c")
keyboard_up(key="ctrl")
```

### 鼠标操作
```python
# 移动鼠标
mouse_move(x=500, y=300, duration=0.5)

# 左键点击
mouse_click(button="left")

# 右键点击指定位置
mouse_click(x=500, y=300, button="right")

# 双击
mouse_click(clicks=2)

# 滚动鼠标
mouse_scroll(clicks=5)
```

## 安全提示

1. **FAILSAFE**: pyautogui启用了故障保护功能，将鼠标快速移动到屏幕左上角(0,0)会触发异常停止程序
2. **权限**: 某些应用程序可能需要管理员权限才能进行键盘鼠标操作
3. **谨慎使用**: 自动化操作可能影响当前工作，建议在测试环境中使用

## 项目结构

```
windows-test-mcp/
├── src/
│   └── windows_test_mcp/
│       ├── __init__.py
│       └── server.py
├── pyproject.toml
└── README.md
```

## 依赖

- `fastmcp` - MCP框架
- `Pillow` - 图像处理
- `pyautogui` - 键盘鼠标控制
- `pydantic` - 数据验证

## 许可证

MIT License
