"""Windows测试MCP服务器.

提供截图、键盘鼠标模拟操作功能。
"""

import base64
import io
import logging
import platform
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


# 启动服务器入口点
def main():
    """运行MCP服务器."""
    mcp.run()


if __name__ == "__main__":
    main()
