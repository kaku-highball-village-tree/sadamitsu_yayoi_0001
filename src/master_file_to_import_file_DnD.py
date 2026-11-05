# -*- coding: utf-8 -*-
"""商品マスタCSVをTSVへ変換するWindowsドラッグ＆ドロップ版です。"""

import csv
import os
from typing import List

import win32api
import win32con
import win32gui

WINDOW_TITLE = "商品マスタ CSV to TSV (Drag & Drop)"
PARENTHESIS_TRANSLATION_TABLE = str.maketrans("（）", "()")


class ConversionError(Exception):
    """CSVをTSVへ変換できない場合の例外です。"""


def validate_csv_path(pszInputPath: str) -> str:
    pszInputPath = os.path.abspath(pszInputPath)
    if not os.path.isfile(pszInputPath):
        raise ConversionError("入力ファイルが見つかりません: " + pszInputPath)
    if os.path.splitext(pszInputPath)[1].lower() != ".csv":
        raise ConversionError("CSVファイルをドラッグ＆ドロップしてください。")
    return pszInputPath


def create_output_path(pszInputPath: str) -> str:
    pszBasePath, _ = os.path.splitext(pszInputPath)
    return pszBasePath + ".tsv"


def convert_dropped_csv_to_tsv(pszInputPath: str) -> str:
    """CP932の43列CSVをBOMなしUTF-8のTSVへ変換します。"""
    pszInputPath = validate_csv_path(pszInputPath)
    pszOutputPath = create_output_path(pszInputPath)

    try:
        with open(pszInputPath, "r", encoding="cp932", newline="") as objInput:
            objReader = csv.reader(objInput)
            with open(pszOutputPath, "w", encoding="utf-8", newline="") as objOutput:
                objWriter = csv.writer(
                    objOutput,
                    delimiter="\t",
                    lineterminator="\r\n",
                    quoting=csv.QUOTE_MINIMAL,
                )
                for arrRow in objReader:
                    objWriter.writerow(arrRow)
    except (OSError, UnicodeError, csv.Error) as objException:
        raise ConversionError(
            "CSVの読み込みまたはTSVの書き込みに失敗しました: "
            + str(objException)
        ) from objException

    return pszOutputPath


def create_step0001_tsv(pszInputTsvPath: str) -> str:
    """TSV内の全角括弧と空白を置換・除去したstep0001ファイルを作成します。"""
    pszInputTsvPath = os.path.abspath(pszInputTsvPath)
    pszDirectory, pszFileName = os.path.split(pszInputTsvPath)
    pszOutputPath = os.path.join(pszDirectory, "step0001_" + pszFileName)

    try:
        with open(pszInputTsvPath, "r", encoding="utf-8", newline="") as objInput:
            objReader = csv.reader(objInput, delimiter="\t")
            with open(pszOutputPath, "w", encoding="utf-8", newline="") as objOutput:
                objWriter = csv.writer(
                    objOutput,
                    delimiter="\t",
                    lineterminator="\r\n",
                    quoting=csv.QUOTE_MINIMAL,
                )
                for arrRow in objReader:
                    arrNormalizedRow = [
                        pszValue.translate(PARENTHESIS_TRANSLATION_TABLE)
                        .replace(" ", "")
                        .replace("　", "")
                        for pszValue in arrRow
                    ]
                    objWriter.writerow(arrNormalizedRow)
    except (OSError, UnicodeError, csv.Error) as objException:
        raise ConversionError(
            "TSVの読み込みまたはstep0001 TSVの書き込みに失敗しました: "
            + str(objException)
        ) from objException

    return pszOutputPath


def show_message(pszMessage: str, bIsError: bool = False) -> None:
    iIcon = win32con.MB_ICONERROR if bIsError else win32con.MB_ICONINFORMATION
    win32gui.MessageBox(0, pszMessage, WINDOW_TITLE, win32con.MB_OK | iIcon)


def draw_instruction_text(iWindowHandle: int) -> None:
    objPaintStruct = win32gui.BeginPaint(iWindowHandle)
    iDeviceContextHandle, objPaint = objPaintStruct
    try:
        objClientRect = win32gui.GetClientRect(iWindowHandle)
        pszText = (
            "CSVファイルをここにドラッグ＆ドロップしてください。\n\n"
            "CP932のCSVを、同じフォルダーにBOMなしUTF-8のTSVとして作成します。\n"
            "同名のTSVがある場合は上書きします。\n\n"
            "複数ファイルをまとめて処理できます。"
        )
        win32gui.DrawText(
            iDeviceContextHandle,
            pszText,
            -1,
            objClientRect,
            win32con.DT_LEFT | win32con.DT_VCENTER | win32con.DT_WORDBREAK,
        )
    finally:
        win32gui.EndPaint(iWindowHandle, objPaint)


def window_proc(iWindowHandle: int, iMessage: int, iWparam: int, iLparam: int) -> int:
    if iMessage == win32con.WM_CREATE:
        win32gui.DragAcceptFiles(iWindowHandle, True)
        return 0
    if iMessage == win32con.WM_DROPFILES:
        iDropHandle = iWparam
        try:
            iFileCount = win32api.DragQueryFile(iDropHandle, -1)
            arrInputPaths: List[str] = [
                os.path.abspath(win32api.DragQueryFile(iDropHandle, iIndex))
                for iIndex in range(iFileCount)
            ]
        finally:
            win32api.DragFinish(iDropHandle)

        if not arrInputPaths:
            show_message("ドラッグされたファイルがありません。", True)
            return 0

        arrSummaryLines: List[str] = []
        iSuccessCount = 0
        iFailureCount = 0
        for pszInputPath in arrInputPaths:
            try:
                pszOutputPath = convert_dropped_csv_to_tsv(pszInputPath)
                pszStep0001Path = create_step0001_tsv(pszOutputPath)
            except ConversionError as objException:
                iFailureCount += 1
                arrSummaryLines.append(
                    "失敗: "
                    + os.path.basename(pszInputPath)
                    + "\n  "
                    + str(objException)
                )
            else:
                iSuccessCount += 1
                arrSummaryLines.append(
                    "成功: "
                    + os.path.basename(pszInputPath)
                    + "\n  作成: "
                    + os.path.basename(pszOutputPath)
                    + "、"
                    + os.path.basename(pszStep0001Path)
                )

        pszSummary = (
            "処理結果: "
            + str(len(arrInputPaths))
            + "件（成功: "
            + str(iSuccessCount)
            + "件、失敗: "
            + str(iFailureCount)
            + "件）\n\n"
            + "\n\n".join(arrSummaryLines)
        )
        show_message(pszSummary, iFailureCount > 0)
        return 0
    if iMessage == win32con.WM_PAINT:
        draw_instruction_text(iWindowHandle)
        return 0
    if iMessage == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    return win32gui.DefWindowProc(iWindowHandle, iMessage, iWparam, iLparam)


def main() -> None:
    pszWindowClassName = "MasterFileToImportFileDndWindowClass"
    iInstanceHandle = win32api.GetModuleHandle(None)
    objWndClass = win32gui.WNDCLASS()
    objWndClass.hInstance = iInstanceHandle
    objWndClass.lpszClassName = pszWindowClassName
    objWndClass.lpfnWndProc = window_proc
    objWndClass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    objWndClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    objWndClass.hbrBackground = win32con.COLOR_WINDOW + 1
    win32gui.RegisterClass(objWndClass)

    iWindowHandle = win32gui.CreateWindowEx(
        win32con.WS_EX_ACCEPTFILES,
        pszWindowClassName,
        WINDOW_TITLE,
        win32con.WS_OVERLAPPED
        | win32con.WS_CAPTION
        | win32con.WS_SYSMENU
        | win32con.WS_MINIMIZEBOX,
        win32con.CW_USEDEFAULT,
        win32con.CW_USEDEFAULT,
        560,
        300,
        0,
        0,
        iInstanceHandle,
        None,
    )
    win32gui.ShowWindow(iWindowHandle, win32con.SW_SHOWNORMAL)
    win32gui.UpdateWindow(iWindowHandle)
    win32gui.PumpMessages()


if __name__ == "__main__":
    main()
