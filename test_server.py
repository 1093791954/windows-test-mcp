"""Windows Test MCP 服务器测试脚本.

此脚本用于测试服务器是否正确加载和工具是否正常工作。
"""

import json
import sys
from pathlib import Path

# 添加src到路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

from windows_test_mcp.server import (
    ScreenshotCaptureInput,
    ScreenshotRegionInput,
    KeyboardPressInput,
    KeyboardTypeInput,
    MouseMoveInput,
    MouseClickInput,
    MouseDownInput,
    MouseUpInput,
    MouseScrollInput,
    screenshot_capture,
    screenshot_region,
    keyboard_press,
    keyboard_down,
    keyboard_up,
    keyboard_type,
    mouse_move,
    mouse_click,
    mouse_down,
    mouse_up,
    mouse_scroll,
    mouse_get_position,
    get_screen_size,
)


def test_screenshot():
    """测试截图工具."""
    print("\n=== 测试截图工具 ===")

    # 测试全屏截图
    result = screenshot_capture(ScreenshotCaptureInput())
    print(f"全屏截图: {result.success} - {result.message}")
    if result.success:
        print(f"  图片大小: {len(result.base64_image) if result.base64_image else 0} 字符(base64)")

    # 测试区域截图
    result = screenshot_region(ScreenshotRegionInput(x=0, y=0, width=100, height=100))
    print(f"区域截图: {result.success} - {result.message}")


def test_mouse_info():
    """测试鼠标信息获取."""
    print("\n=== 测试鼠标信息工具 ===")

    # 获取屏幕尺寸
    result = get_screen_size()
    print(f"屏幕尺寸: {result}")

    # 获取鼠标位置
    result = mouse_get_position()
    print(f"鼠标位置: {result}")


def test_mouse_move():
    """测试鼠标移动."""
    print("\n=== 测试鼠标移动 ===")
    print("注意: 这将会移动您的鼠标!")
    input("按 Enter 继续...")

    # 移动鼠标
    result = mouse_move(MouseMoveInput(x=100, y=100, duration=0.5))
    print(f"移动结果: {result}")


def test_mouse_click():
    """测试鼠标点击."""
    print("\n=== 测试鼠标点击 ===")
    print("注意: 这将会点击您的鼠标!")
    input("按 Enter 继续...")

    # 左键点击当前位置
    result = mouse_click(MouseClickInput(button="left", clicks=1))
    print(f"左键点击: {result}")


def test_keyboard():
    """测试键盘操作."""
    print("\n=== 测试键盘操作 ===")
    print("注意: 这将会模拟键盘输入!")
    input("按 Enter 继续...")

    # 输入文本
    print("将输入 'Hello'")
    result = keyboard_type(KeyboardTypeInput(text="Hello", interval=0.01))
    print(f"输入结果: {result}")

    # 按Enter
    print("将按 Enter 键")
    result = keyboard_press(KeyboardPressInput(key="enter"))
    print(f"按键结果: {result}")


def list_all_tools():
    """列出所有可用工具."""
    print("\n=== 可用工具列表 ===")

    tools = [
        "1. 截图工具:",
        "   - screenshot_capture: 截取整个屏幕",
        "   - screenshot_region: 截取指定区域",
        "",
        "2. 键盘工具:",
        "   - keyboard_press: 按键点击",
        "   - keyboard_down: 按下按键",
        "   - keyboard_up: 释放按键",
        "   - keyboard_type: 输入文本",
        "",
        "3. 鼠标工具:",
        "   - mouse_move: 移动鼠标",
        "   - mouse_click: 点击鼠标",
        "   - mouse_down: 按下鼠标按钮",
        "   - mouse_up: 释放鼠标按钮",
        "   - mouse_scroll: 滚动鼠标滚轮",
        "   - mouse_get_position: 获取鼠标位置",
        "   - get_screen_size: 获取屏幕尺寸",
    ]

    for tool in tools:
        print(tool)


def main():
    """主函数."""
    print("Windows Test MCP 服务器测试")
    print("=" * 40)

    # 列出所有工具
    list_all_tools()

    # 运行测试
    try:
        test_screenshot()
        test_mouse_info()

        # 交互式测试
        print("\n" + "=" * 40)
        print("是否继续测试鼠标和键盘操作?")
        print("注意: 这些测试将实际控制您的鼠标和键盘!")
        choice = input("输入 'yes' 继续, 或其他键跳过: ")

        if choice.lower() == "yes":
            test_mouse_move()
            test_mouse_click()
            test_keyboard()
        else:
            print("跳过鼠标和键盘测试")

    except KeyboardInterrupt:
        print("\n测试被用户中断")
    except Exception as e:
        print(f"\n测试出错: {e}")
        import traceback
        traceback.print_exc()

    print("\n" + "=" * 40)
    print("测试完成!")


if __name__ == "__main__":
    main()
