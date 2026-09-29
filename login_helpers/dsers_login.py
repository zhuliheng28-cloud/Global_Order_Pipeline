import asyncio
from typing import Optional, Callable, Any

async def wait_or_auto_login_dsers(
    page, 
    log_func: Optional[Callable[[str], None]] = print, 
    timeout_seconds: int = 30, 
    task_info: Optional[Any] = None,
    prefix: str = ""
):
    """
    等待 DSers 登录：
    若指定时间内用户手动点击登录并跳转，则立即进入后续流程；
    若超时（默认 30 秒）后仍停留在登录页面，则自动定位并点击【LOG IN】按钮提交登录。
    """
    def _log(msg: str):
        full_msg = f"{prefix} {msg}".strip() if prefix else msg
        if log_func:
            log_func(full_msg)

    try:
        # 若未处于登录页，直接放行
        if "login" not in page.url.lower() and "accounts.dsers.com" not in page.url.lower():
            return

        _log(f"[*] 检测到 DSers 登录页面，等待登录中（若 {timeout_seconds} 秒内未手动点击，将自动点击【LOG IN】）...")
        
        elapsed = 0
        auto_clicked = False
        click_attempts = 0
        
        while "login" in page.url.lower() or "accounts.dsers.com" in page.url.lower():
            if task_info:
                if hasattr(task_info, "async_check_pause"):
                    await task_info.async_check_pause()
                elif hasattr(task_info, "check_pause"):
                    task_info.check_pause()

            # 达到超时时间且尚未自动点击，或已点击但在 10 秒后仍未跳转（最多重试 2 次）
            should_click = False
            if elapsed >= timeout_seconds and not auto_clicked:
                should_click = True
            elif auto_clicked and click_attempts < 2 and (elapsed - timeout_seconds) >= 10:
                should_click = True

            if should_click:
                click_attempts += 1
                _log(f"[*] {timeout_seconds} 秒内未检测到手动登录，正在自动点击【LOG IN】按钮 (尝试 {click_attempts}/2)...")
                try:
                    # 1. 尝试 Playwright 原生定位器点击
                    btn_loc = page.locator("button:has-text('LOG IN'), button:has-text('Log in'), button:has-text('Log In'), button[type='submit'], .ant-btn-primary").filter(visible=True).first
                    if await btn_loc.count() > 0:
                        await btn_loc.click(force=True)
                        auto_clicked = True
                        _log("[*] 已成功点击【LOG IN】按钮，等待登录完成...")
                    else:
                        # 2. 备用 evaluate 检索点击
                        clicked = await page.evaluate('''() => {
                            const buttons = Array.from(document.querySelectorAll('button, input[type="submit"], a.btn'));
                            for (const b of buttons) {
                                const text = (b.innerText || b.value || '').trim().toUpperCase();
                                if (text.includes('LOG IN') || text === 'LOGIN') {
                                    b.click();
                                    return true;
                                }
                            }
                            const sub = document.querySelector('button[type="submit"], input[type="submit"]');
                            if (sub) {
                                sub.click();
                                return true;
                            }
                            return false;
                        }''')
                        if clicked:
                            auto_clicked = True
                            _log("[*] 已通过备用选择器点击【LOG IN】按钮，等待登录完成...")
                        else:
                            _log("[!] 未找到【LOG IN】按钮，请手动点击登录。")
                except Exception as e:
                    _log(f"[!] 自动点击登录异常: {e}")

            # 每 5 秒输出一次剩余倒计时提示
            if not auto_clicked and elapsed > 0 and elapsed % 5 == 0 and elapsed < timeout_seconds:
                _log(f"[*] 等待手动登录中... (剩余 {timeout_seconds - elapsed} 秒后将自动点击【LOG IN】)")

            await asyncio.sleep(1)
            elapsed += 1
    except Exception as e:
        _log(f"[*] 登录等待循环结束: {e}")
