#!/usr/bin/env python3
"""easyMoney.swift 的 Windows Python 迁移版入口。

这里保留兼容入口和公共导出层；具体实现按功能拆分到
easy_money_win_* 模块中。
"""

from __future__ import annotations

from easy_money_win_core import *
from easy_money_win_input import *
from easy_money_win_capture import *
from easy_money_win_uia import *
from easy_money_win_llm import *
from easy_money_win_commands import *
from easy_money_win_llm import _DOTENV_CACHE


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print()
        print_ts("已停止")
        raise SystemExit(130)
    except EasyMoneyError as exc:
        print_ts(f"错误: {exc}")
        raise SystemExit(1)
