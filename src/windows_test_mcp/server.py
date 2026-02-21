"""Windows测试MCP服务器.

提供截图、键盘鼠标模拟操作功能。
"""

import base64
import io
import logging
import platform
import subprocess
from typing import Literal, Optional

from fastmcp import FastMCP
from PIL import ImageGrab
from pydantic import BaseModel, Field

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 检查Windows平台
if platform.system() != "Windows":
    logger.warning("此MCP服务器专为Windows设计，当前平台可能不完全兼容")

# 导入pyautogui
try:
    import pyautogui

    # 配置pyautogui安全设置
    pyautogui.FAILSAFE = True  # 鼠标移动到屏幕角落会触发异常
    pyautogui.PAUSE = 0.05  # 每个操作后的默认暂停时间
except ImportError:
    logger.error("pyautogui未安装，键盘鼠标功能将不可用")
    pyautogui = None

# 导入pywin32用于窗口管理
# pywin32相关导入放在函数内部以避免启动时失败
pywin32_available = False
try:
    import win32gui
    import win32con
    import win32ui
    import win32api
    import win32process
    from PIL import Image

    pywin32_available = True
except ImportError:
    logger.warning("pywin32未安装，窗口管理功能将不可用")

# 创建MCP服务器
mcp = FastMCP("windows-test-mcp")


# ==================== 截图工具 ====================


class ScreenshotCaptureInput(BaseModel):
    """截图输入参数."""

    filename: Optional[str] = Field(
        None, description="保存截图的文件名(可选，不包含扩展名)"
    )


class ScreenshotRegionInput(BaseModel):
    """区域截图输入参数."""

    x: int = Field(..., description="截图区域左上角X坐标", ge=0)
    y: int = Field(..., description="截图区域左上角Y坐标", ge=0)
    width: int = Field(..., description="截图区域宽度", gt=0)
    height: int = Field(..., description="截图区域高度", gt=0)
    filename: Optional[str] = Field(
        None, description="保存截图的文件名(可选，不包含扩展名)"
    )


class ScreenshotOutput(BaseModel):
    """截图输出结果."""

    success: bool = Field(..., description="操作是否成功")
    message: str = Field(..., description="操作结果消息")
    base64_image: Optional[str] = Field(None, description="Base64编码的PNG图片")
    filename: Optional[str] = Field(None, description="保存的文件名")


@mcp.tool(
    description="截取整个屏幕",
    annotations={"readOnlyHint": True, "destructiveHint": False},
)
def screenshot_capture(input_data: ScreenshotCaptureInput) -> ScreenshotOutput:
    """截取整个屏幕并返回base64编码的图片.

    Args:
        input_data: 截图参数

    Returns:
        ScreenshotOutput: 包含base64编码图片的结果
    """
    try:
        # 截取屏幕
        screenshot = ImageGrab.grab()

        # 转换为base64
        buffer = io.BytesIO()
        screenshot.save(buffer, format="PNG")
        base64_image = base64.b64encode(buffer.getvalue()).decode("utf-8")

        # 如果指定了文件名，保存到文件
        saved_filename = None
        if input_data.filename:
            saved_filename = f"{input_data.filename}.png"
            screenshot.save(saved_filename)

        return ScreenshotOutput(
            success=True,
            message="截图成功",
            base64_image=base64_image,
            filename=saved_filename,
        )
    except Exception as e:
        logger.error(f"截图失败: {e}")
        return ScreenshotOutput(
            success=False, message=f"截图失败: {str(e)}", base64_image=None, filename=None
        )


@mcp.tool(
    description="截取屏幕指定区域",
    annotations={"readOnlyHint": True, "destructiveHint": False},
)
def screenshot_region(input_data: ScreenshotRegionInput) -> ScreenshotOutput:
    """截取屏幕指定区域并返回base64编码的图片.

    Args:
        input_data: 区域截图参数，包含x, y, width, height

    Returns:
        ScreenshotOutput: 包含base64编码图片的结果
    """
    try:
        # 计算截图区域
        bbox = (input_data.x, input_data.y, input_data.x + input_data.width, input_data.y + input_data.height)

        # 截取指定区域
        screenshot = ImageGrab.grab(bbox=bbox)

        # 转换为base64
        buffer = io.BytesIO()
        screenshot.save(buffer, format="PNG")
        base64_image = base64.b64encode(buffer.getvalue()).decode("utf-8")

        # 如果指定了文件名，保存到文件
        saved_filename = None
        if input_data.filename:
            saved_filename = f"{input_data.filename}.png"
            screenshot.save(saved_filename)

        return ScreenshotOutput(
            success=True,
            message="区域截图成功",
            base64_image=base64_image,
            filename=saved_filename,
        )
    except Exception as e:
        logger.error(f"区域截图失败: {e}")
        return ScreenshotOutput(
            success=False, message=f"区域截图失败: {str(e)}", base64_image=None, filename=None
        )


# ==================== 键盘工具 ====================


class KeyboardPressInput(BaseModel):
    """按键输入参数."""

    key: str = Field(
        ...,
        description="要按下的键，支持单个字符或特殊键名称如'enter','space','ctrl','shift','alt','tab','esc','backspace','delete','up','down','left','right'等",
    )
    presses: int = Field(1, description="按键次数", ge=1)
    interval: float = Field(0.0, description="多次按键之间的间隔时间(秒)", ge=0)


class KeyboardTypeInput(BaseModel):
    """键盘输入字符串参数."""

    text: str = Field(..., description="要输入的文本")
    interval: float = Field(0.01, description="字符间间隔时间(秒)", ge=0)


class KeyboardOperationOutput(BaseModel):
    """键盘操作输出结果."""

    success: bool = Field(..., description="操作是否成功")
    message: str = Field(..., description="操作结果消息")


@mcp.tool(
    description="按下并释放指定按键",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def keyboard_press(input_data: KeyboardPressInput) -> KeyboardOperationOutput:
    """按下并释放指定按键(click操作).

    Args:
        input_data: 按键参数，包含key、presses、interval

    Returns:
        KeyboardOperationOutput: 操作结果
    """
    if pyautogui is None:
        return KeyboardOperationOutput(success=False, message="pyautogui未安装")

    try:
        pyautogui.press(input_data.key, presses=input_data.presses, interval=input_data.interval)
        return KeyboardOperationOutput(
            success=True, message=f"按键 '{input_data.key}' 操作成功"
        )
    except Exception as e:
        logger.error(f"按键操作失败: {e}")
        return KeyboardOperationOutput(success=False, message=f"按键操作失败: {str(e)}")


@mcp.tool(
    description="按下指定按键不放",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def keyboard_down(input_data: KeyboardPressInput) -> KeyboardOperationOutput:
    """按下指定按键不释放.

    Args:
        input_data: 按键参数，包含key

    Returns:
        KeyboardOperationOutput: 操作结果
    """
    if pyautogui is None:
        return KeyboardOperationOutput(success=False, message="pyautogui未安装")

    try:
        pyautogui.keyDown(input_data.key)
        return KeyboardOperationOutput(
            success=True, message=f"按键 '{input_data.key}' 按下成功"
        )
    except Exception as e:
        logger.error(f"按键按下失败: {e}")
        return KeyboardOperationOutput(success=False, message=f"按键按下失败: {str(e)}")


@mcp.tool(
    description="释放指定按键",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def keyboard_up(input_data: KeyboardPressInput) -> KeyboardOperationOutput:
    """释放指定按键.

    Args:
        input_data: 按键参数，包含key

    Returns:
        KeyboardOperationOutput: 操作结果
    """
    if pyautogui is None:
        return KeyboardOperationOutput(success=False, message="pyautogui未安装")

    try:
        pyautogui.keyUp(input_data.key)
        return KeyboardOperationOutput(
            success=True, message=f"按键 '{input_data.key}' 释放成功"
        )
    except Exception as e:
        logger.error(f"按键释放失败: {e}")
        return KeyboardOperationOutput(success=False, message=f"按键释放失败: {str(e)}")


@mcp.tool(
    description="输入文本字符串",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def keyboard_type(input_data: KeyboardTypeInput) -> KeyboardOperationOutput:
    """模拟键盘输入文本字符串.

    Args:
        input_data: 输入参数，包含text和interval

    Returns:
        KeyboardOperationOutput: 操作结果
    """
    if pyautogui is None:
        return KeyboardOperationOutput(success=False, message="pyautogui未安装")

    try:
        pyautogui.typewrite(input_data.text, interval=input_data.interval)
        return KeyboardOperationOutput(success=True, message=f"文本输入成功: {input_data.text}")
    except Exception as e:
        logger.error(f"文本输入失败: {e}")
        return KeyboardOperationOutput(success=False, message=f"文本输入失败: {str(e)}")


# ==================== 鼠标工具 ====================

MouseButton = Literal["left", "right", "middle", "mouse4", "mouse5"]


class MouseMoveInput(BaseModel):
    """鼠标移动输入参数."""

    x: int = Field(..., description="目标X坐标", ge=0)
    y: int = Field(..., description="目标Y坐标", ge=0)
    duration: float = Field(0.0, description="移动动画持续时间(秒)", ge=0)


class MouseClickInput(BaseModel):
    """鼠标点击输入参数."""

    x: Optional[int] = Field(None, description="点击位置X坐标(不指定则使用当前位置)", ge=0)
    y: Optional[int] = Field(None, description="点击位置Y坐标(不指定则使用当前位置)", ge=0)
    button: MouseButton = Field("left", description="鼠标按钮: left, right, middle, mouse4, mouse5")
    clicks: int = Field(1, description="点击次数", ge=1)
    interval: float = Field(0.0, description="多次点击间隔时间(秒)", ge=0)


class MouseDownInput(BaseModel):
    """鼠标按下输入参数."""

    x: Optional[int] = Field(None, description="按下位置X坐标(不指定则使用当前位置)", ge=0)
    y: Optional[int] = Field(None, description="按下位置Y坐标(不指定则使用当前位置)", ge=0)
    button: MouseButton = Field("left", description="鼠标按钮: left, right, middle, mouse4, mouse5")


class MouseUpInput(BaseModel):
    """鼠标释放输入参数."""

    x: Optional[int] = Field(None, description="释放位置X坐标(不指定则使用当前位置)", ge=0)
    y: Optional[int] = Field(None, description="释放位置Y坐标(不指定则使用当前位置)", ge=0)
    button: MouseButton = Field("left", description="鼠标按钮: left, right, middle, mouse4, mouse5")


class MouseScrollInput(BaseModel):
    """鼠标滚轮输入参数."""

    clicks: int = Field(..., description="滚动距离(正数向上，负数向下)")
    x: Optional[int] = Field(None, description="滚动时鼠标X坐标", ge=0)
    y: Optional[int] = Field(None, description="滚动时鼠标Y坐标", ge=0)


class MouseOperationOutput(BaseModel):
    """鼠标操作输出结果."""

    success: bool = Field(..., description="操作是否成功")
    message: str = Field(..., description="操作结果消息")
    current_position: Optional[tuple[int, int]] = Field(None, description="当前鼠标位置(x, y)")


@mcp.tool(
    description="移动鼠标到指定位置",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def mouse_move(input_data: MouseMoveInput) -> MouseOperationOutput:
    """移动鼠标到指定坐标位置.

    Args:
        input_data: 移动参数，包含x, y, duration

    Returns:
        MouseOperationOutput: 操作结果和当前位置
    """
    if pyautogui is None:
        return MouseOperationOutput(success=False, message="pyautogui未安装", current_position=None)

    try:
        pyautogui.moveTo(input_data.x, input_data.y, duration=input_data.duration)
        current_pos = pyautogui.position()
        return MouseOperationOutput(
            success=True,
            message=f"鼠标移动到 ({input_data.x}, {input_data.y})",
            current_position=(current_pos.x, current_pos.y),
        )
    except Exception as e:
        logger.error(f"鼠标移动失败: {e}")
        return MouseOperationOutput(
            success=False, message=f"鼠标移动失败: {str(e)}", current_position=None
        )


@mcp.tool(
    description="点击鼠标指定按钮",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def mouse_click(input_data: MouseClickInput) -> MouseOperationOutput:
    """在指定位置点击鼠标按钮.

    Args:
        input_data: 点击参数，包含x, y, button, clicks, interval

    Returns:
        MouseOperationOutput: 操作结果和当前位置
    """
    if pyautogui is None:
        return MouseOperationOutput(success=False, message="pyautogui未安装", current_position=None)

    try:
        # 映射按钮名称
        button_map = {
            "left": "left",
            "right": "right",
            "middle": "middle",
            "mouse4": "x1",  # Windows侧键1
            "mouse5": "x2",  # Windows侧键2
        }
        pyautogui_button = button_map.get(input_data.button, "left")

        # 执行点击
        pyautogui.click(
            x=input_data.x,
            y=input_data.y,
            button=pyautogui_button,
            clicks=input_data.clicks,
            interval=input_data.interval,
        )

        current_pos = pyautogui.position()
        return MouseOperationOutput(
            success=True,
            message=f"鼠标 {input_data.button} 按钮点击 {input_data.clicks} 次",
            current_position=(current_pos.x, current_pos.y),
        )
    except Exception as e:
        logger.error(f"鼠标点击失败: {e}")
        return MouseOperationOutput(
            success=False, message=f"鼠标点击失败: {str(e)}", current_position=None
        )


@mcp.tool(
    description="按下鼠标按钮不放",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def mouse_down(input_data: MouseDownInput) -> MouseOperationOutput:
    """在指定位置按下鼠标按钮不释放.

    Args:
        input_data: 按下参数，包含x, y, button

    Returns:
        MouseOperationOutput: 操作结果和当前位置
    """
    if pyautogui is None:
        return MouseOperationOutput(success=False, message="pyautogui未安装", current_position=None)

    try:
        # 如果指定了位置，先移动
        if input_data.x is not None and input_data.y is not None:
            pyautogui.moveTo(input_data.x, input_data.y)

        # 映射按钮名称
        button_map = {
            "left": "left",
            "right": "right",
            "middle": "middle",
            "mouse4": "x1",
            "mouse5": "x2",
        }
        pyautogui_button = button_map.get(input_data.button, "left")

        pyautogui.mouseDown(button=pyautogui_button)

        current_pos = pyautogui.position()
        return MouseOperationOutput(
            success=True,
            message=f"鼠标 {input_data.button} 按钮按下",
            current_position=(current_pos.x, current_pos.y),
        )
    except Exception as e:
        logger.error(f"鼠标按下失败: {e}")
        return MouseOperationOutput(
            success=False, message=f"鼠标按下失败: {str(e)}", current_position=None
        )


@mcp.tool(
    description="释放鼠标按钮",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def mouse_up(input_data: MouseUpInput) -> MouseOperationOutput:
    """在指定位置释放鼠标按钮.

    Args:
        input_data: 释放参数，包含x, y, button

    Returns:
        MouseOperationOutput: 操作结果和当前位置
    """
    if pyautogui is None:
        return MouseOperationOutput(success=False, message="pyautogui未安装", current_position=None)

    try:
        # 如果指定了位置，先移动
        if input_data.x is not None and input_data.y is not None:
            pyautogui.moveTo(input_data.x, input_data.y)

        # 映射按钮名称
        button_map = {
            "left": "left",
            "right": "right",
            "middle": "middle",
            "mouse4": "x1",
            "mouse5": "x2",
        }
        pyautogui_button = button_map.get(input_data.button, "left")

        pyautogui.mouseUp(button=pyautogui_button)

        current_pos = pyautogui.position()
        return MouseOperationOutput(
            success=True,
            message=f"鼠标 {input_data.button} 按钮释放",
            current_position=(current_pos.x, current_pos.y),
        )
    except Exception as e:
        logger.error(f"鼠标释放失败: {e}")
        return MouseOperationOutput(
            success=False, message=f"鼠标释放失败: {str(e)}", current_position=None
        )


@mcp.tool(
    description="滚动鼠标滚轮",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def mouse_scroll(input_data: MouseScrollInput) -> MouseOperationOutput:
    """滚动鼠标滚轮.

    Args:
        input_data: 滚动参数，包含clicks, x, y

    Returns:
        MouseOperationOutput: 操作结果和当前位置
    """
    if pyautogui is None:
        return MouseOperationOutput(success=False, message="pyautogui未安装", current_position=None)

    try:
        # 如果指定了位置，先移动鼠标
        if input_data.x is not None and input_data.y is not None:
            pyautogui.moveTo(input_data.x, input_data.y)

        pyautogui.scroll(input_data.clicks)

        current_pos = pyautogui.position()
        direction = "向上" if input_data.clicks > 0 else "向下"
        return MouseOperationOutput(
            success=True,
            message=f"鼠标滚轮{direction}滚动 {abs(input_data.clicks)} 单位",
            current_position=(current_pos.x, current_pos.y),
        )
    except Exception as e:
        logger.error(f"鼠标滚动失败: {e}")
        return MouseOperationOutput(
            success=False, message=f"鼠标滚动失败: {str(e)}", current_position=None
        )


@mcp.tool(
    description="获取当前鼠标位置",
    annotations={"readOnlyHint": True, "destructiveHint": False},
)
def mouse_get_position() -> MouseOperationOutput:
    """获取当前鼠标光标位置.

    Returns:
        MouseOperationOutput: 当前鼠标位置
    """
    if pyautogui is None:
        return MouseOperationOutput(success=False, message="pyautogui未安装", current_position=None)

    try:
        current_pos = pyautogui.position()
        return MouseOperationOutput(
            success=True,
            message=f"当前鼠标位置: ({current_pos.x}, {current_pos.y})",
            current_position=(current_pos.x, current_pos.y),
        )
    except Exception as e:
        logger.error(f"获取鼠标位置失败: {e}")
        return MouseOperationOutput(
            success=False, message=f"获取鼠标位置失败: {str(e)}", current_position=None
        )


@mcp.tool(
    description="获取屏幕尺寸",
    annotations={"readOnlyHint": True, "destructiveHint": False},
)
def get_screen_size() -> dict:
    """获取屏幕尺寸.

    Returns:
        dict: 包含屏幕宽度和高度
    """
    if pyautogui is None:
        return {"success": False, "message": "pyautogui未安装"}

    try:
        size = pyautogui.size()
        return {
            "success": True,
            "width": size.width,
            "height": size.height,
            "message": f"屏幕尺寸: {size.width}x{size.height}",
        }
    except Exception as e:
        logger.error(f"获取屏幕尺寸失败: {e}")
        return {"success": False, "message": f"获取屏幕尺寸失败: {str(e)}"}


# ==================== 窗口管理工具 ====================


class WindowActivateInput(BaseModel):
    """窗口激活输入参数."""

    process_name: str = Field(..., description="进程名，如 'notepad.exe' 或部分名称 'notepad'")
    wait_time: float = Field(0.5, description="激活后等待时间(秒)", ge=0)


class WindowCaptureInput(BaseModel):
    """窗口截图输入参数."""

    process_name: str = Field(..., description="进程名")
    filename: Optional[str] = Field(None, description="保存截图的文件名(可选)")


class WindowGetRectInput(BaseModel):
    """获取窗口矩形输入参数."""

    process_name: str = Field(..., description="进程名")


class WindowOperationOutput(BaseModel):
    """窗口操作输出结果."""

    success: bool = Field(..., description="操作是否成功")
    message: str = Field(..., description="操作结果消息")
    hwnd: Optional[int] = Field(None, description="窗口句柄")


class WindowRectOutput(BaseModel):
    """窗口矩形信息输出."""

    success: bool = Field(..., description="操作是否成功")
    message: str = Field(..., description="操作结果消息")
    x: Optional[int] = Field(None, description="窗口左上角X坐标")
    y: Optional[int] = Field(None, description="窗口左上角Y坐标")
    width: Optional[int] = Field(None, description="窗口宽度")
    height: Optional[int] = Field(None, description="窗口高度")


class WindowScreenshotOutput(BaseModel):
    """窗口截图输出结果."""

    success: bool = Field(..., description="操作是否成功")
    message: str = Field(..., description="操作结果消息")
    base64_image: Optional[str] = Field(None, description="Base64编码的PNG图片")
    filename: Optional[str] = Field(None, description="保存的文件名")
    hwnd: Optional[int] = Field(None, description="窗口句柄")


def _find_window_by_process_name(process_name: str) -> list[int]:
    """通过进程名查找所有匹配的窗口句柄."""
    if not pywin32_available:
        return []

    hwnds = []
    process_name_lower = process_name.lower()

    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd):
            try:
                # 获取窗口所属进程ID
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                # 打开进程获取名称
                handle = win32api.OpenProcess(
                    win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid
                )
                try:
                    module_name = win32process.GetModuleFileNameEx(handle, 0)
                    # 支持部分匹配进程名
                    if process_name_lower in module_name.lower():
                        extra.append(hwnd)
                finally:
                    win32api.CloseHandle(handle)
            except Exception:
                pass
        return True

    win32gui.EnumWindows(callback, hwnds)
    return hwnds


def _capture_window_to_image(hwnd: int) -> Image.Image:
    """截图窗口并返回PIL Image对象."""
    # 获取窗口尺寸
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top

    # 创建设备上下文
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    # 创建位图
    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    saveDC.SelectObject(saveBitMap)

    # 使用PrintWindow截图（支持后台窗口）
    # PW_RENDERFULLCONTENT = 2 (Win8+ 支持渲染完整内容)
    win32gui.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)

    # 转换为PIL Image
    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    im = Image.frombuffer(
        "RGB",
        (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
        bmpstr,
        "raw",
        "BGRX",
        0,
        1,
    )

    # 清理资源
    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)

    return im


@mcp.tool(
    description="通过进程名激活窗口到前台",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def window_activate(input_data: WindowActivateInput) -> WindowOperationOutput:
    """通过进程名找到窗口并激活到前台.

    Args:
        input_data: 激活参数，包含process_name和wait_time

    Returns:
        WindowOperationOutput: 操作结果
    """
    if not pywin32_available:
        return WindowOperationOutput(
            success=False, message="pywin32未安装，窗口管理功能不可用", hwnd=None
        )

    try:
        hwnds = _find_window_by_process_name(input_data.process_name)

        if not hwnds:
            return WindowOperationOutput(
                success=False,
                message=f"未找到进程名包含 '{input_data.process_name}' 的窗口",
                hwnd=None,
            )

        hwnd = hwnds[0]

        # 如果窗口最小化，先恢复
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        # 强制激活
        win32gui.SetForegroundWindow(hwnd)

        # 等待指定时间
        import time

        time.sleep(input_data.wait_time)

        return WindowOperationOutput(
            success=True,
            message=f"窗口已激活，句柄: {hwnd}",
            hwnd=hwnd,
        )
    except Exception as e:
        logger.error(f"窗口激活失败: {e}")
        return WindowOperationOutput(
            success=False, message=f"窗口激活失败: {str(e)}", hwnd=None
        )


@mcp.tool(
    description="通过进程名后台截图窗口（无需调出前台）",
    annotations={"readOnlyHint": True, "destructiveHint": False},
)
def window_capture_background(input_data: WindowCaptureInput) -> WindowScreenshotOutput:
    """通过进程名找到窗口并在后台截图（不激活窗口）.

    Args:
        input_data: 截图参数，包含process_name和可选的filename

    Returns:
        WindowScreenshotOutput: 包含base64编码图片的结果
    """
    if not pywin32_available:
        return WindowScreenshotOutput(
            success=False,
            message="pywin32未安装，窗口管理功能不可用",
            base64_image=None,
            filename=None,
            hwnd=None,
        )

    try:
        hwnds = _find_window_by_process_name(input_data.process_name)

        if not hwnds:
            return WindowScreenshotOutput(
                success=False,
                message=f"未找到进程名包含 '{input_data.process_name}' 的窗口",
                base64_image=None,
                filename=None,
                hwnd=None,
            )

        hwnd = hwnds[0]

        # 截图窗口
        screenshot = _capture_window_to_image(hwnd)

        # 转换为base64
        buffer = io.BytesIO()
        screenshot.save(buffer, format="PNG")
        base64_image = base64.b64encode(buffer.getvalue()).decode("utf-8")

        # 如果指定了文件名，保存到文件
        saved_filename = None
        if input_data.filename:
            saved_filename = f"{input_data.filename}.png"
            screenshot.save(saved_filename)

        return WindowScreenshotOutput(
            success=True,
            message="后台截图成功",
            base64_image=base64_image,
            filename=saved_filename,
            hwnd=hwnd,
        )
    except Exception as e:
        logger.error(f"后台截图失败: {e}")
        return WindowScreenshotOutput(
            success=False,
            message=f"后台截图失败: {str(e)}",
            base64_image=None,
            filename=None,
            hwnd=None,
        )


@mcp.tool(
    description="通过进程名激活窗口并截图",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def window_capture_foreground(input_data: WindowCaptureInput) -> WindowScreenshotOutput:
    """通过进程名找到窗口，激活到前台后截图.

    Args:
        input_data: 截图参数，包含process_name和可选的filename

    Returns:
        WindowScreenshotOutput: 包含base64编码图片的结果
    """
    if not pywin32_available:
        return WindowScreenshotOutput(
            success=False,
            message="pywin32未安装，窗口管理功能不可用",
            base64_image=None,
            filename=None,
            hwnd=None,
        )

    try:
        hwnds = _find_window_by_process_name(input_data.process_name)

        if not hwnds:
            return WindowScreenshotOutput(
                success=False,
                message=f"未找到进程名包含 '{input_data.process_name}' 的窗口",
                base64_image=None,
                filename=None,
                hwnd=None,
            )

        hwnd = hwnds[0]

        # 如果窗口最小化，先恢复
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        # 强制激活
        win32gui.SetForegroundWindow(hwnd)

        # 等待窗口完全激活
        import time

        time.sleep(0.5)

        # 截图窗口
        screenshot = _capture_window_to_image(hwnd)

        # 转换为base64
        buffer = io.BytesIO()
        screenshot.save(buffer, format="PNG")
        base64_image = base64.b64encode(buffer.getvalue()).decode("utf-8")

        # 如果指定了文件名，保存到文件
        saved_filename = None
        if input_data.filename:
            saved_filename = f"{input_data.filename}.png"
            screenshot.save(saved_filename)

        return WindowScreenshotOutput(
            success=True,
            message="前台截图成功",
            base64_image=base64_image,
            filename=saved_filename,
            hwnd=hwnd,
        )
    except Exception as e:
        logger.error(f"前台截图失败: {e}")
        return WindowScreenshotOutput(
            success=False,
            message=f"前台截图失败: {str(e)}",
            base64_image=None,
            filename=None,
            hwnd=None,
        )


@mcp.tool(
    description="通过进程名获取窗口大小和位置",
    annotations={"readOnlyHint": True, "destructiveHint": False},
)
def window_get_rect(input_data: WindowGetRectInput) -> WindowRectOutput:
    """通过进程名找到窗口并获取其位置和大小.

    Args:
        input_data: 参数，包含process_name

    Returns:
        WindowRectOutput: 窗口位置和大小信息
    """
    if not pywin32_available:
        return WindowRectOutput(
            success=False,
            message="pywin32未安装，窗口管理功能不可用",
            x=None,
            y=None,
            width=None,
            height=None,
        )

    try:
        hwnds = _find_window_by_process_name(input_data.process_name)

        if not hwnds:
            return WindowRectOutput(
                success=False,
                message=f"未找到进程名包含 '{input_data.process_name}' 的窗口",
                x=None,
                y=None,
                width=None,
                height=None,
            )

        hwnd = hwnds[0]

        # 获取窗口矩形
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)

        return WindowRectOutput(
            success=True,
            message=f"窗口位置: ({left}, {top}), 大小: {right - left}x{bottom - top}",
            x=left,
            y=top,
            width=right - left,
            height=bottom - top,
        )
    except Exception as e:
        logger.error(f"获取窗口矩形失败: {e}")
        return WindowRectOutput(
            success=False,
            message=f"获取窗口矩形失败: {str(e)}",
            x=None,
            y=None,
            width=None,
            height=None,
        )


# ==================== 应用程序管理工具 ====================

import psutil


class AppLaunchInput(BaseModel):
    """启动应用程序输入参数."""

    app_path: str = Field(..., description="应用程序路径，如 'notepad.exe' 或完整路径 'C:\\Windows\\notepad.exe'")
    args: Optional[str] = Field(None, description="启动参数（可选）")
    cwd: Optional[str] = Field(None, description="工作目录（可选），指定程序启动时的工作文件夹路径")
    wait_time: float = Field(0.5, description="启动后等待时间(秒)", ge=0)


class AppTerminateInput(BaseModel):
    """终止应用程序输入参数."""

    process_name: str = Field(..., description="进程名，如 'notepad.exe' 或部分名称 'notepad'")


class AppLaunchOutput(BaseModel):
    """应用程序启动输出结果."""

    success: bool = Field(..., description="操作是否成功")
    message: str = Field(..., description="操作结果消息")
    pid: Optional[int] = Field(None, description="进程ID")
    process_name: Optional[str] = Field(None, description="进程名称")


class AppTerminateOutput(BaseModel):
    """应用程序终止输出结果."""

    success: bool = Field(..., description="操作是否成功")
    message: str = Field(..., description="操作结果消息")
    terminated_pids: list[int] = Field(default_factory=list, description="被终止的进程ID列表")


class AppListOutput(BaseModel):
    """应用程序列表输出结果."""

    success: bool = Field(..., description="操作是否成功")
    message: str = Field(..., description="操作结果消息")
    processes: list[dict] = Field(default_factory=list, description="正在运行的进程列表")


@mcp.tool(
    description="后台启动应用程序",
    annotations={"readOnlyHint": False, "destructiveHint": False},
)
def app_launch(input_data: AppLaunchInput) -> AppLaunchOutput:
    """后台启动应用程序并返回进程信息.

    Args:
        input_data: 启动参数，包含app_path、args和wait_time

    Returns:
        AppLaunchOutput: 包含进程ID和名称的结果
    """
    try:
        # 构建命令
        cmd = [input_data.app_path]
        if input_data.args:
            cmd.extend(input_data.args.split())

        # 准备启动参数
        popen_kwargs = {
            "stdout": subprocess.DEVNULL,
            "stderr": subprocess.DEVNULL,
            "stdin": subprocess.DEVNULL,
        }

        # 如果指定了工作目录，添加到参数
        if input_data.cwd:
            popen_kwargs["cwd"] = input_data.cwd

        # 后台启动程序（不阻塞）
        if platform.system() == "Windows":
            # Windows: 使用CREATE_NEW_CONSOLE创建新窗口，不阻塞
            creation_flags = subprocess.CREATE_NEW_CONSOLE | subprocess.CREATE_NEW_PROCESS_GROUP
            process = subprocess.Popen(
                cmd,
                creationflags=creation_flags,
                **popen_kwargs,
            )
        else:
            # 其他平台
            process = subprocess.Popen(
                cmd,
                start_new_session=True,
                **popen_kwargs,
            )

        # 等待指定时间让程序初始化
        import time
        time.sleep(input_data.wait_time)

        # 获取进程信息
        try:
            proc = psutil.Process(process.pid)
            process_name = proc.name()
        except psutil.NoSuchProcess:
            process_name = input_data.app_path.split("\\")[-1].split("/")[-1]

        return AppLaunchOutput(
            success=True,
            message=f"应用程序 '{input_data.app_path}' 启动成功",
            pid=process.pid,
            process_name=process_name,
        )
    except FileNotFoundError:
        return AppLaunchOutput(
            success=False,
            message=f"找不到应用程序: '{input_data.app_path}'",
            pid=None,
            process_name=None,
        )
    except Exception as e:
        logger.error(f"启动应用程序失败: {e}")
        return AppLaunchOutput(
            success=False,
            message=f"启动应用程序失败: {str(e)}",
            pid=None,
            process_name=None,
        )


@mcp.tool(
    description="终止指定名称的应用程序进程",
    annotations={"readOnlyHint": False, "destructiveHint": True},
)
def app_terminate(input_data: AppTerminateInput) -> AppTerminateOutput:
    """通过进程名终止所有匹配的应用程序.

    Args:
        input_data: 终止参数，包含process_name

    Returns:
        AppTerminateOutput: 包含被终止的进程ID列表
    """
    try:
        terminated_pids = []
        process_name_lower = input_data.process_name.lower()

        # 遍历所有进程查找匹配的
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            try:
                proc_name = proc.info['name'] or ""
                proc_exe = proc.info['exe'] or ""

                # 匹配进程名或路径
                if (process_name_lower in proc_name.lower() or
                    process_name_lower in proc_exe.lower()):
                    pid = proc.info['pid']
                    proc.terminate()  # 先尝试优雅终止
                    terminated_pids.append(pid)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        if terminated_pids:
            return AppTerminateOutput(
                success=True,
                message=f"已终止 {len(terminated_pids)} 个进程: {input_data.process_name}",
                terminated_pids=terminated_pids,
            )
        else:
            return AppTerminateOutput(
                success=False,
                message=f"未找到匹配的进程: {input_data.process_name}",
                terminated_pids=[],
            )
    except Exception as e:
        logger.error(f"终止进程失败: {e}")
        return AppTerminateOutput(
            success=False,
            message=f"终止进程失败: {str(e)}",
            terminated_pids=[],
        )


@mcp.tool(
    description="列出当前正在运行的应用程序进程",
    annotations={"readOnlyHint": True, "destructiveHint": False},
)
def app_list_running() -> AppListOutput:
    """获取当前系统中正在运行的应用程序列表.

    Returns:
        AppListOutput: 包含进程列表的结果
    """
    try:
        processes = []

        for proc in psutil.process_iter(['pid', 'name', 'exe', 'memory_info', 'cpu_percent']):
            try:
                proc_info = proc.info
                # 只显示有窗口的进程（GUI应用）或常见的应用程序
                process_name = proc_info['name'] or ""
                exe_path = proc_info['exe'] or ""

                # 过滤掉系统进程，只保留可能的应用程序
                if exe_path and not exe_path.lower().startswith(
                    (r'c:\windows\system32', r'c:\windows\syswow64')
                ):
                    processes.append({
                        'pid': proc_info['pid'],
                        'name': process_name,
                        'exe': exe_path,
                        'memory_mb': round(proc_info['memory_info'].rss / 1024 / 1024, 2) if proc_info['memory_info'] else 0,
                    })
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        # 按内存使用排序，取前50个
        processes.sort(key=lambda x: x['memory_mb'], reverse=True)
        processes = processes[:50]

        return AppListOutput(
            success=True,
            message=f"当前有 {len(processes)} 个应用程序在运行",
            processes=processes,
        )
    except Exception as e:
        logger.error(f"获取进程列表失败: {e}")
        return AppListOutput(
            success=False,
            message=f"获取进程列表失败: {str(e)}",
            processes=[],
        )


# 启动服务器入口点
def main():
    """运行MCP服务器."""
    mcp.run()


if __name__ == "__main__":
    main()
