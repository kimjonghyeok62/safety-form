# -*- coding: utf-8 -*-
"""Reliable way to get the live workbook COM object via the Running Object Table,
binding directly to the file-path moniker instead of GetActiveObject (which can
return a stale/zombie Excel.Application instance when multiple processes exist)."""
import os

import pythoncom
import win32com.client

# 이 스크립트 폴더의 한 단계 위(프로젝트 폴더)에 있는 통합문서
TARGET_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                           "중대재해예방 안전관련 자동화서식.xlsm")


def get_workbook(path=TARGET_PATH):
    rot = pythoncom.GetRunningObjectTable()
    enum = rot.EnumRunning()
    while True:
        monikers = enum.Next(1)
        if not monikers:
            break
        moniker = monikers[0]
        ctx = pythoncom.CreateBindCtx(0)
        try:
            name = moniker.GetDisplayName(ctx, None)
        except Exception:
            continue
        if name == path:
            obj = rot.GetObject(moniker)
            idisp = obj.QueryInterface(pythoncom.IID_IDispatch)
            return win32com.client.Dispatch(idisp)
    raise RuntimeError(f"Workbook not found in ROT: {path}")
