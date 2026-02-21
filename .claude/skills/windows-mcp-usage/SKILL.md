---
name: windows-mcp-usage
description: Windows MCP工具正确使用指南。当用户使用windows-app-test-mcp服务器提供的工具操作Windows应用程序时使用。包含窗口管理、应用启动、截图、键盘鼠标操作的最佳实践。特别用于纠正Claude在使用这些工具时的常见错误，如未激活窗口就执行键盘/鼠标操作。
---

# Windows MCP 使用指南

## 核心原则

使用Windows MCP工具操作应用程序时，**必须先激活窗口到前台，再执行键盘或鼠标操作**。这是最常见的错误。

## 工具分类

### 1. 应用程序管理工具

| 工具 | 用途 | 是否需要激活 |
|------|------|--------------|
| `app_launch` | 启动应用程序 | 后台启动，无需激活 |
| `app_terminate` | 终止应用程序 | 直接终止，无需激活 |
| `app_list_running` | 列出运行中的应用 | 只读查询，无需激活 |

**使用示例：**
```python
# 启动记事本
app_launch({"app_path": "notepad.exe", "wait_time": 1.0})

# 终止记事本
app_terminate({"process_name": "notepad.exe"})
```

### 2. 窗口管理工具

| 工具 | 用途 | 是否需要激活 |
|------|------|--------------|
| `window_activate` | 激活窗口到前台 | 这就是激活操作本身 |
| `window_capture_background` | 后台截图窗口 | 无需激活，后台截图 |
| `window_capture_foreground` | 前台截图窗口 | 自动激活并截图 |
| `window_get_rect` | 获取窗口位置大小 | 只读查询，无需激活 |

**使用示例：**
```python
# 查看窗口状态（后台截图，无需激活）
window_capture_background({"process_name": "notepad", "filename": "check_state"})

# 激活窗口到前台
window_activate({"process_name": "notepad", "wait_time": 0.5})
```

### 3. 截图工具

| 工具 | 用途 | 是否需要激活 |
|------|------|--------------|
| `screenshot_capture` | 截取整个屏幕 | 无需激活 |
| `screenshot_region` | 截取指定区域 | 无需激活 |

### 4. 键盘鼠标工具（关键：必须先激活）

| 工具 | 用途 | **必须先激活窗口** |
|------|------|-------------------|
| `keyboard_press` | 按键操作 | **是** |
| `keyboard_type` | 输入文本 | **是** |
| `keyboard_down` | 按下不放 | **是** |
| `keyboard_up` | 释放按键 | **是** |
| `mouse_click` | 鼠标点击 | **是** |
| `mouse_move` | 移动鼠标 | **是** |
| `mouse_down` | 鼠标按下 | **是** |
| `mouse_up` | 鼠标释放 | **是** |
| `mouse_scroll` | 滚轮滚动 | **是** |

## 标准操作流程

### 场景1：启动应用并输入内容

**错误做法：**
```python
# ❌ 错误：直接启动后就输入，窗口可能未激活
app_launch({"app_path": "notepad.exe"})
keyboard_type({"text": "Hello World"})  # 可能输入到错误的地方
```

**正确做法：**
```python
# 1. 启动应用
app_launch({"app_path": "notepad.exe", "wait_time": 1.0})

# 2. 激活窗口到前台
window_activate({"process_name": "notepad", "wait_time": 0.5})

# 3. 执行键盘/鼠标操作
keyboard_type({"text": "Hello World"})
```

### 场景2：查看应用状态

```python
# 只需后台截图，无需激活
window_capture_background({"process_name": "chrome", "filename": "browser_state"})
```

### 场景3：点击特定按钮

```python
# 1. 激活窗口
window_activate({"process_name": "notepad", "wait_time": 0.5})

# 2. 移动鼠标到按钮位置
mouse_move({"x": 100, "y": 200, "duration": 0.2})

# 3. 点击
mouse_click({"button": "left"})
```

## 进程名说明

- `process_name` 支持部分匹配，如 `"notepad"` 可以匹配 `"notepad.exe"`
- 常见进程名：
  - 记事本: `notepad.exe`
  - 计算器: `calc.exe` 或 `Calculator`
  - Chrome: `chrome.exe`
  - Edge: `msedge.exe`
  - VS Code: `Code.exe`

## 注意事项

1. **键盘/鼠标操作前必须激活窗口** - 这是最常见的错误
2. **激活后等待0.3-0.5秒** - 给窗口切换留出时间
3. **坐标是屏幕绝对坐标** - 不是窗口相对坐标
4. **后台截图不会干扰当前工作** - 适合监控状态
5. **前台截图会自动激活窗口** - 如果需要看最终效果可用此工具

## 错误排查

| 问题 | 可能原因 | 解决方案 |
|------|----------|----------|
| 输入没反应 | 窗口未激活 | 先调用 window_activate |
| 点击位置错误 | 坐标是窗口相对坐标 | 使用屏幕绝对坐标 |
| 找不到窗口 | 进程名不对 | 用 app_list_running 查看正确进程名 |
| 截图失败 | 窗口已关闭 | 检查应用是否还在运行 |

## 快速参考

**只读操作（无需激活）：**
- 查看窗口状态 → `window_capture_background`
- 获取窗口位置 → `window_get_rect`
- 屏幕截图 → `screenshot_capture`

**写入操作（必须先激活）：**
- 输入文字 → 先 `window_activate` 再 `keyboard_type`
- 点击按钮 → 先 `window_activate` 再 `mouse_click`
- 组合键 → 先 `window_activate` 再 `keyboard_down/up`
