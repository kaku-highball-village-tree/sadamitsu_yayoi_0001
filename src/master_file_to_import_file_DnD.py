# -*- coding: utf-8 -*-
"""CSVをTSVへ変換するWindowsドラッグ＆ドロップ版です。"""

import csv
import os
import textwrap
import tempfile
import unicodedata
import zipfile
from decimal import Decimal, InvalidOperation
from xml.etree import ElementTree
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.utils.exceptions import IllegalCharacterError
from openpyxl.utils import get_column_letter

import win32api
import win32con
import win32gui

WINDOW_TITLE = "商品マスタ CSV to TSV (Drag & Drop)"
PARENTHESIS_TRANSLATION_TABLE = str.maketrans("（）", "()")
PRICE_MASTER_STEP0001_PREFIX = "step0001_得意先別商品売上単価マスタ"
CUSTOMER_MASTER_STEP0001_PREFIX = "step0001_得意先マスタ"
PRODUCT_MASTER_STEP0001_PREFIX = "step0001_商品マスタ"
INVENTORY_MASTER_STEP0001_PREFIX = "step0001_在庫マスタ"
EXCEL_NUMERIC_HEADERS = {
    "売上単価(ｋｇ)",
    "売上単価(個数)",
    "社内単価(ｋｇ)",
    "社内単価(個数)",
    "売上単価(その他)",
    "期首在庫数量",
    "仕入数量",
    "入庫出来高数量",
    "入庫移動数量",
    "入庫その他数量１",
    "売上数量",
    "出庫出来高数量",
    "出庫移動数量",
    "出庫その他数量２",
    "実棚卸数量",
}
EXCEL_MAX_COLUMN_WIDTH = 25


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
    """CP932のCSVをBOMなしUTF-8のTSVへ変換します。"""
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


def create_step0001_excel_file(pszStep0001TsvPath: str) -> str:
    """step0001 TSVを同名の1シートExcelへ、値を文字列のまま出力します。"""
    pszStep0001TsvPath = os.path.abspath(pszStep0001TsvPath)
    pszExcelPath = os.path.splitext(pszStep0001TsvPath)[0] + ".xlsx"
    iLastRow = 0
    iLastColumn = 0

    try:
        objWorkbook = Workbook()
        objWorksheet = objWorkbook.active
        with open(pszStep0001TsvPath, "r", encoding="utf-8", newline="") as objInput:
            for iRow, arrValues in enumerate(
                csv.reader(objInput, delimiter="\t"), start=1
            ):
                iLastRow = iRow
                iLastColumn = max(iLastColumn, len(arrValues))
                for iColumn, pszValue in enumerate(arrValues, start=1):
                    objCell = objWorksheet.cell(iRow, iColumn, pszValue)
                    objCell.data_type = "s"
                    objCell.number_format = "@"
        objWorkbook.save(pszExcelPath)
        if iLastRow > 1 and iLastColumn > 0:
            pszLastColumn = get_column_letter(iLastColumn)
            ignore_excel_text_number_warnings(
                pszExcelPath,
                ["A2:" + pszLastColumn + str(iLastRow)],
            )
    except (OSError, UnicodeError, csv.Error, ValueError, IllegalCharacterError) as objException:
        raise ConversionError(
            "step0001 TSVの読み込みまたはExcelファイルの作成に失敗しました: "
            + str(objException)
        ) from objException

    return pszExcelPath


def find_tsv_header_index(
    arrHeader: List[str], arrCandidateNames: List[str], pszInputPath: str
) -> int:
    for pszCandidateName in arrCandidateNames:
        if pszCandidateName in arrHeader:
            return arrHeader.index(pszCandidateName)
    raise ConversionError(
        "必要な見出しがTSVにありません: "
        + " / ".join(arrCandidateNames)
        + " ("
        + os.path.basename(pszInputPath)
        + ")"
    )


def read_master_lookup(
    pszInputPath: str,
    arrKeyHeaderNames: List[str],
    arrValueHeaderNames: List[str],
) -> Dict[str, Tuple[str, ...]]:
    with open(pszInputPath, "r", encoding="utf-8", newline="") as objInput:
        objReader = csv.reader(objInput, delimiter="\t")
        arrHeader = next(objReader, None)
        if arrHeader is None:
            raise ConversionError("マスタTSVに見出し行がありません: " + pszInputPath)
        iKeyIndex = find_tsv_header_index(
            arrHeader, arrKeyHeaderNames, pszInputPath
        )
        arrValueIndexes = [
            find_tsv_header_index(arrHeader, [pszName], pszInputPath)
            for pszName in arrValueHeaderNames
        ]

        objLookup: Dict[str, Tuple[str, ...]] = {}
        for arrRow in objReader:
            pszKey = arrRow[iKeyIndex] if iKeyIndex < len(arrRow) else ""
            if not pszKey or pszKey in objLookup:
                continue
            objLookup[pszKey] = tuple(
                arrRow[iIndex] if iIndex < len(arrRow) else ""
                for iIndex in arrValueIndexes
            )
    return objLookup


def _create_step0002_tsv(
    pszPriceTsvPath: str,
    pszCustomerTsvPath: str,
    pszProductTsvPath: str,
    bNameInsertedOnly: bool,
) -> str:
    """得意先・商品名をコード照合で追加したstep0002 TSVを作成します。"""
    pszPriceTsvPath = os.path.abspath(pszPriceTsvPath)
    pszCustomerTsvPath = os.path.abspath(pszCustomerTsvPath)
    pszProductTsvPath = os.path.abspath(pszProductTsvPath)
    pszDirectory, pszFileName = os.path.split(pszPriceTsvPath)
    if pszFileName.startswith("step0001_"):
        pszFileName = pszFileName[len("step0001_") :]
    if bNameInsertedOnly:
        pszPriceMasterPrefix = "得意先別商品売上単価マスタ"
        if pszFileName.startswith(pszPriceMasterPrefix):
            pszFileName = (
                pszPriceMasterPrefix
                + "_名挿入"
                + pszFileName[len(pszPriceMasterPrefix) :]
            )
    pszOutputPath = os.path.join(pszDirectory, "step0002_" + pszFileName)

    try:
        objCustomerLookup = read_master_lookup(
            pszCustomerTsvPath,
            ["得意先コード"],
            ["得意先名(漢字１)", "得意先名(漢字２)", "得意先名(カナ)"],
        )
        objProductLookup = read_master_lookup(
            pszProductTsvPath,
            ["商品コ－ド", "商品コード"],
            ["商品名(漢字)", "商品名(カナ)"],
        )
        with open(pszPriceTsvPath, "r", encoding="utf-8", newline="") as objInput:
            objReader = csv.reader(objInput, delimiter="\t")
            arrHeader = next(objReader, None)
            if arrHeader is None:
                raise ConversionError("得意先別商品売上単価TSVに見出し行がありません。")
            iCustomerCodeIndex = find_tsv_header_index(
                arrHeader, ["得意先コード"], pszPriceTsvPath
            )
            iProductCodeIndex = find_tsv_header_index(
                arrHeader, ["商品コード", "商品コ－ド"], pszPriceTsvPath
            )
            arrOutputHeader: List[str] = []
            arrCustomerNameHeaders = (
                ["得意先名(漢字１)"]
                if bNameInsertedOnly
                else ["得意先名(漢字１)", "得意先名(漢字２)", "得意先名(カナ)"]
            )
            arrProductNameHeaders = (
                ["商品名(漢字)"]
                if bNameInsertedOnly
                else ["商品名(漢字)", "商品名(カナ)"]
            )
            for iIndex, pszValue in enumerate(arrHeader):
                arrOutputHeader.append(pszValue)
                if iIndex == iCustomerCodeIndex:
                    arrOutputHeader.extend(arrCustomerNameHeaders)
                if iIndex == iProductCodeIndex:
                    arrOutputHeader.extend(arrProductNameHeaders)

            with open(pszOutputPath, "w", encoding="utf-8", newline="") as objOutput:
                objWriter = csv.writer(
                    objOutput,
                    delimiter="\t",
                    lineterminator="\r\n",
                    quoting=csv.QUOTE_MINIMAL,
                )
                objWriter.writerow(arrOutputHeader)
                for arrRow in objReader:
                    iRequiredLength = max(
                        len(arrHeader), iCustomerCodeIndex + 1, iProductCodeIndex + 1
                    )
                    arrPaddedRow = arrRow + [""] * max(0, iRequiredLength - len(arrRow))
                    pszCustomerCode = arrPaddedRow[iCustomerCodeIndex]
                    pszProductCode = arrPaddedRow[iProductCodeIndex]
                    arrOutputRow: List[str] = []
                    for iIndex, pszValue in enumerate(arrPaddedRow):
                        arrOutputRow.append(pszValue)
                        if iIndex == iCustomerCodeIndex:
                            arrCustomerNames = objCustomerLookup.get(
                                pszCustomerCode, ("", "", "")
                            )
                            arrOutputRow.extend(
                                arrCustomerNames[: len(arrCustomerNameHeaders)]
                            )
                        if iIndex == iProductCodeIndex:
                            arrProductNames = objProductLookup.get(
                                pszProductCode, ("", "")
                            )
                            arrOutputRow.extend(
                                arrProductNames[: len(arrProductNameHeaders)]
                            )
                    objWriter.writerow(arrOutputRow)
    except ConversionError:
        raise
    except (OSError, UnicodeError, csv.Error) as objException:
        raise ConversionError(
            "step0001 TSVの読み込みまたはstep0002 TSVの書き込みに失敗しました: "
            + str(objException)
        ) from objException

    return pszOutputPath


def create_step0002_tsv(
    pszPriceTsvPath: str,
    pszCustomerTsvPath: str,
    pszProductTsvPath: str,
) -> str:
    """得意先名3列・商品名2列を追加したstep0002 TSVを作成します。"""
    return _create_step0002_tsv(
        pszPriceTsvPath, pszCustomerTsvPath, pszProductTsvPath, False
    )


def create_step0002_name_inserted_tsv(
    pszPriceTsvPath: str,
    pszCustomerTsvPath: str,
    pszProductTsvPath: str,
) -> str:
    """得意先名1列・商品名1列を追加した名挿入step0002 TSVを作成します。"""
    return _create_step0002_tsv(
        pszPriceTsvPath, pszCustomerTsvPath, pszProductTsvPath, True
    )


def create_inventory_step0002_tsv(
    pszInventoryTsvPath: str, pszProductTsvPath: str
) -> str:
    """商品名2列を商品コードの直後に加えた在庫step0002 TSVを作成します。"""
    return _create_inventory_step0002_tsv(
        pszInventoryTsvPath, pszProductTsvPath, False
    )


def create_inventory_step0002_name_inserted_tsv(
    pszInventoryTsvPath: str, pszProductTsvPath: str
) -> str:
    """商品名(漢字)を商品コードの直後に加えた在庫TSVを作成します。"""
    return _create_inventory_step0002_tsv(
        pszInventoryTsvPath, pszProductTsvPath, True
    )


def _create_inventory_step0002_tsv(
    pszInventoryTsvPath: str, pszProductTsvPath: str, bNameInsertedOnly: bool
) -> str:
    pszInventoryTsvPath = os.path.abspath(pszInventoryTsvPath)
    pszProductTsvPath = os.path.abspath(pszProductTsvPath)
    pszDirectory, pszFileName = os.path.split(pszInventoryTsvPath)
    if pszFileName.startswith("step0001_"):
        pszFileName = pszFileName[len("step0001_") :]
    if bNameInsertedOnly and pszFileName.startswith("在庫マスタ"):
        pszFileName = "在庫マスタ_名挿入" + pszFileName[len("在庫マスタ") :]
    pszOutputPath = os.path.join(pszDirectory, "step0002_" + pszFileName)

    try:
        objProductLookup = read_master_lookup(
            pszProductTsvPath,
            ["商品コ－ド", "商品コード"],
            ["商品名(漢字)", "商品名(カナ)"],
        )
        with open(pszInventoryTsvPath, "r", encoding="utf-8", newline="") as objInput:
            objReader = csv.reader(objInput, delimiter="\t")
            arrHeader = next(objReader, None)
            if arrHeader is None:
                raise ConversionError("在庫マスタTSVに見出し行がありません。")
            iProductCodeIndex = find_tsv_header_index(
                arrHeader, ["商品コード", "商品コ－ド"], pszInventoryTsvPath
            )
            arrNameHeaders = ["商品名(漢字)"] if bNameInsertedOnly else [
                "商品名(漢字)", "商品名(カナ)"
            ]
            arrOutputHeader = list(arrHeader)
            arrOutputHeader[iProductCodeIndex + 1 : iProductCodeIndex + 1] = arrNameHeaders

            with open(pszOutputPath, "w", encoding="utf-8", newline="") as objOutput:
                objWriter = csv.writer(
                    objOutput,
                    delimiter="\t",
                    lineterminator="\r\n",
                    quoting=csv.QUOTE_MINIMAL,
                )
                objWriter.writerow(arrOutputHeader)
                for arrRow in objReader:
                    arrPaddedRow = arrRow + [""] * max(0, len(arrHeader) - len(arrRow))
                    pszProductCode = arrPaddedRow[iProductCodeIndex]
                    arrProductNames = objProductLookup.get(pszProductCode, ("", ""))
                    arrOutputRow = list(arrPaddedRow)
                    arrOutputRow[iProductCodeIndex + 1 : iProductCodeIndex + 1] = (
                        list(arrProductNames[: len(arrNameHeaders)])
                    )
                    objWriter.writerow(arrOutputRow)
    except ConversionError:
        raise
    except (OSError, UnicodeError, csv.Error) as objException:
        raise ConversionError(
            "在庫または商品マスタTSVの読み込み・step0002 TSVの書き込みに失敗しました: "
            + str(objException)
        ) from objException
    return pszOutputPath


def find_step0001_master_paths(
    pszStep0001Path: str,
) -> Optional[Tuple[str, str, str]]:
    pszDirectory, pszFileName = os.path.split(os.path.abspath(pszStep0001Path))
    for pszPrefix in (
        PRICE_MASTER_STEP0001_PREFIX,
        CUSTOMER_MASTER_STEP0001_PREFIX,
        PRODUCT_MASTER_STEP0001_PREFIX,
    ):
        if pszFileName.startswith(pszPrefix):
            pszSuffix = pszFileName[len(pszPrefix) :]
            if not pszSuffix.lower().endswith(".tsv"):
                return None
            return (
                os.path.join(pszDirectory, PRICE_MASTER_STEP0001_PREFIX + pszSuffix),
                os.path.join(pszDirectory, CUSTOMER_MASTER_STEP0001_PREFIX + pszSuffix),
                os.path.join(pszDirectory, PRODUCT_MASTER_STEP0001_PREFIX + pszSuffix),
            )
    return None


def create_step0002_tsv_if_ready(
    pszStep0001Path: str,
) -> Tuple[Optional[str], Optional[str], Optional[str], List[str]]:
    arrMasterPaths = find_step0001_master_paths(pszStep0001Path)
    if arrMasterPaths is None:
        return None, None, None, []
    pszPriceTsvPath, pszCustomerTsvPath, pszProductTsvPath = arrMasterPaths
    arrMissingPaths = [pszPath for pszPath in arrMasterPaths if not os.path.isfile(pszPath)]
    if arrMissingPaths:
        return None, None, None, arrMissingPaths
    pszStep0002Path = create_step0002_tsv(
        pszPriceTsvPath, pszCustomerTsvPath, pszProductTsvPath
    )
    pszStep0002NameInsertedPath = create_step0002_name_inserted_tsv(
        pszPriceTsvPath, pszCustomerTsvPath, pszProductTsvPath
    )
    pszExcelPath = create_step0002_excel_file(
        pszStep0002Path, pszStep0002NameInsertedPath
    )
    return pszStep0002Path, pszStep0002NameInsertedPath, pszExcelPath, []


def find_step0001_inventory_master_paths(
    pszStep0001Path: str,
) -> Optional[Tuple[str, str]]:
    pszDirectory, pszFileName = os.path.split(os.path.abspath(pszStep0001Path))
    for pszPrefix in (INVENTORY_MASTER_STEP0001_PREFIX, PRODUCT_MASTER_STEP0001_PREFIX):
        if pszFileName.startswith(pszPrefix):
            pszSuffix = pszFileName[len(pszPrefix) :]
            if not pszSuffix.lower().endswith(".tsv"):
                return None
            return (
                os.path.join(pszDirectory, INVENTORY_MASTER_STEP0001_PREFIX + pszSuffix),
                os.path.join(pszDirectory, PRODUCT_MASTER_STEP0001_PREFIX + pszSuffix),
            )
    return None


def create_inventory_step0002_tsv_if_ready(
    pszStep0001Path: str,
) -> Tuple[Optional[str], Optional[str], Optional[str], List[str]]:
    arrMasterPaths = find_step0001_inventory_master_paths(pszStep0001Path)
    if arrMasterPaths is None:
        return None, None, None, []
    pszInventoryTsvPath, pszProductTsvPath = arrMasterPaths
    pszInputFileName = os.path.basename(os.path.abspath(pszStep0001Path))
    if (
        not pszInputFileName.startswith(INVENTORY_MASTER_STEP0001_PREFIX)
        and not os.path.isfile(pszInventoryTsvPath)
    ):
        return None, None, None, []
    arrMissingPaths = [pszPath for pszPath in arrMasterPaths if not os.path.isfile(pszPath)]
    if arrMissingPaths:
        return None, None, None, arrMissingPaths
    pszStep0002Path = create_inventory_step0002_tsv(
        pszInventoryTsvPath, pszProductTsvPath
    )
    pszStep0002NameInsertedPath = create_inventory_step0002_name_inserted_tsv(
        pszInventoryTsvPath, pszProductTsvPath
    )
    pszExcelPath = create_step0002_excel_file(
        pszStep0002Path, pszStep0002NameInsertedPath
    )
    return pszStep0002Path, pszStep0002NameInsertedPath, pszExcelPath, []


def create_step0002_excel_file(
    pszStep0002Path: str, pszStep0002NameInsertedPath: str
) -> str:
    """2つのstep0002 TSVを指定シート名のExcelブックにまとめます。"""
    pszDirectory, pszFileName = os.path.split(pszStep0002Path)
    if not pszFileName.startswith("step0002_"):
        raise ConversionError("step0002 TSVのファイル名が不正です: " + pszFileName)
    pszExcelPath = os.path.join(
        pszDirectory, os.path.splitext(pszFileName[len("step0002_") :])[0] + ".xlsx"
    )

    arrSheetTextRanges: List[str] = []
    try:
        objWorkbook = Workbook()
        for iIndex, (pszTsvPath, pszSheetName) in enumerate(
            (
                (pszStep0002Path, "カナあり"),
                (pszStep0002NameInsertedPath, "カナなし"),
            )
        ):
            objSheet = (
                objWorkbook.active if iIndex == 0 else objWorkbook.create_sheet()
            )
            objSheet.title = pszSheetName
            arrHeaders: List[str] = []
            arrIgnoredCells: List[str] = []
            iLastRow = 0
            with open(pszTsvPath, "r", encoding="utf-8", newline="") as objInput:
                for iRow, arrValues in enumerate(
                    csv.reader(objInput, delimiter="\t"), start=1
                ):
                    iLastRow = iRow
                    if iRow == 1:
                        arrHeaders = arrValues
                    for iColumn, pszValue in enumerate(arrValues, start=1):
                        pszHeader = (
                            arrHeaders[iColumn - 1]
                            if iColumn <= len(arrHeaders)
                            else ""
                        )
                        if (
                            iRow > 1
                            and pszHeader in EXCEL_NUMERIC_HEADERS
                            and pszValue
                        ):
                            try:
                                objValue = Decimal(pszValue)
                                if not objValue.is_finite():
                                    raise InvalidOperation
                            except InvalidOperation:
                                objValue = pszValue
                                arrIgnoredCells.append(
                                    get_column_letter(iColumn) + str(iRow)
                                )
                            else:
                                if len(objValue.as_tuple().digits) > 15:
                                    objValue = pszValue
                                    arrIgnoredCells.append(
                                        get_column_letter(iColumn) + str(iRow)
                                    )
                                elif objValue == objValue.to_integral_value():
                                    objValue = int(objValue)
                            objCell = objSheet.cell(iRow, iColumn, objValue)
                            if isinstance(objValue, str):
                                objCell.data_type = "s"
                                objCell.number_format = "@"
                        else:
                            objCell = objSheet.cell(iRow, iColumn, pszValue)
                            objCell.data_type = "s"
                            objCell.number_format = "@"
            if iLastRow > 1:
                arrTextColumnRanges = [
                    get_column_letter(iColumn)
                    + "2:"
                    + get_column_letter(iColumn)
                    + str(iLastRow)
                    for iColumn, pszHeader in enumerate(arrHeaders, start=1)
                    if pszHeader not in EXCEL_NUMERIC_HEADERS
                ]
                arrSheetTextRanges.append(" ".join(arrTextColumnRanges))
            else:
                arrSheetTextRanges.append("")
            if arrIgnoredCells:
                arrSheetTextRanges[-1] = " ".join(
                    filter(None, [arrSheetTextRanges[-1], " ".join(arrIgnoredCells)])
                )
        adjust_excel_column_widths(objWorkbook)
        objWorkbook.save(pszExcelPath)
        ignore_excel_text_number_warnings(pszExcelPath, arrSheetTextRanges)
    except (
        OSError,
        UnicodeError,
        csv.Error,
        ValueError,
        IllegalCharacterError,
    ) as objException:
        raise ConversionError(
            "step0002 TSVの読み込みまたはExcelファイルの作成に失敗しました: "
            + str(objException)
        ) from objException
    return pszExcelPath


def adjust_excel_column_widths(objWorkbook: Workbook) -> None:
    """各シートの内容に合わせて幅を調整し、25を上限にします。"""
    for objSheet in objWorkbook.worksheets:
        for arrColumnCells in objSheet.iter_cols():
            iMaximumWidth = 0
            for objCell in arrColumnCells:
                if objCell.value is None:
                    continue
                pszValue = str(objCell.value)
                iCellWidth = sum(
                    2 if unicodedata.east_asian_width(pszCharacter) in ("F", "W") else 1
                    for pszCharacter in pszValue
                )
                iMaximumWidth = max(iMaximumWidth, iCellWidth)
            iColumn = arrColumnCells[0].column
            objSheet.column_dimensions[get_column_letter(iColumn)].width = min(
                max(1, iMaximumWidth), EXCEL_MAX_COLUMN_WIDTH
            )


def ignore_excel_text_number_warnings(
    pszExcelPath: str, arrSheetTextRanges: List[str]
) -> None:
    """文字列で保持する列に「数値が文字列」警告の無視範囲を設定します。"""
    pszTemporaryPath = ""
    try:
        with tempfile.NamedTemporaryFile(
            prefix="excel_ignored_errors_",
            suffix=".xlsx",
            dir=os.path.dirname(pszExcelPath),
            delete=False,
        ) as objTemporaryFile:
            pszTemporaryPath = objTemporaryFile.name

        pszSpreadsheetNamespace = (
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        )
        ElementTree.register_namespace("", pszSpreadsheetNamespace)
        with zipfile.ZipFile(pszExcelPath, "r") as objSourceZip:
            with zipfile.ZipFile(
                pszTemporaryPath, "w", compression=zipfile.ZIP_DEFLATED
            ) as objOutputZip:
                for objZipInfo in objSourceZip.infolist():
                    objData = objSourceZip.read(objZipInfo.filename)
                    pszName = objZipInfo.filename
                    if pszName.startswith("xl/worksheets/sheet") and pszName.endswith(
                        ".xml"
                    ):
                        pszSheetNumber = os.path.splitext(
                            os.path.basename(pszName)
                        )[0][len("sheet") :]
                        if pszSheetNumber.isdigit():
                            iSheetIndex = int(pszSheetNumber) - 1
                            if (
                                iSheetIndex < len(arrSheetTextRanges)
                                and arrSheetTextRanges[iSheetIndex]
                            ):
                                objRoot = ElementTree.fromstring(objData)
                                pszIgnoredErrorsTag = (
                                    "{" + pszSpreadsheetNamespace + "}ignoredErrors"
                                )
                                objIgnoredErrors = ElementTree.Element(
                                    pszIgnoredErrorsTag
                                )
                                objIgnoredError = ElementTree.SubElement(
                                    objIgnoredErrors,
                                    "{" + pszSpreadsheetNamespace + "}ignoredError",
                                )
                                objIgnoredError.set(
                                    "sqref", arrSheetTextRanges[iSheetIndex]
                                )
                                objIgnoredError.set("numberStoredAsText", "1")
                                arrTrailingTags = {
                                    "drawing",
                                    "legacyDrawing",
                                    "legacyDrawingHF",
                                    "picture",
                                    "oleObjects",
                                    "controls",
                                    "webPublishItems",
                                    "tableParts",
                                    "extLst",
                                }
                                iInsertIndex = len(objRoot)
                                for iIndex, objChild in enumerate(objRoot):
                                    pszChildTag = objChild.tag.rsplit("}", 1)[-1]
                                    if pszChildTag in arrTrailingTags:
                                        iInsertIndex = iIndex
                                        break
                                objRoot.insert(iInsertIndex, objIgnoredErrors)
                                objData = ElementTree.tostring(
                                    objRoot,
                                    encoding="utf-8",
                                    xml_declaration=True,
                                )
                    objOutputZip.writestr(objZipInfo, objData)
        os.replace(pszTemporaryPath, pszExcelPath)
    except (OSError, ValueError, zipfile.BadZipFile, ElementTree.ParseError) as objException:
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        raise ConversionError(
            "Excelの文字列セル警告設定に失敗しました: " + str(objException)
        ) from objException


def show_message(pszMessage: str, bIsError: bool = False) -> None:
    iScreenWidth = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    iScreenHeight = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    iMaxCharactersPerLine = max(24, min(80, (iScreenWidth - 160) // 17))
    iMaxLines = max(8, min(32, (iScreenHeight - 180) // 24))
    arrDisplayLines: List[str] = []
    bTruncated = False

    for pszLine in pszMessage.splitlines():
        arrWrappedLines = textwrap.wrap(
            pszLine,
            width=iMaxCharactersPerLine,
            break_long_words=True,
            break_on_hyphens=False,
        ) or [""]
        if len(arrDisplayLines) + len(arrWrappedLines) > iMaxLines:
            bTruncated = True
            break
        arrDisplayLines.extend(arrWrappedLines)

    if bTruncated:
        arrDisplayLines = arrDisplayLines[: iMaxLines - 1]
        arrDisplayLines.append("（以降の結果を省略しました）")
    pszMessage = "\n".join(arrDisplayLines)
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
            "各step0001 TSVから、同じフォルダーに同名のExcelも作成します。\n\n"
            "得意先別商品売上単価・得意先・商品マスタのstep0001 TSVが揃うと、"
            "通常版と名挿入版のstep0002 TSV、および2シートのExcelファイルも作成します。\n\n"
            "在庫マスタと商品マスタのstep0001 TSVが揃った場合も同様に作成します。\n\n"
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
        arrProcessedStep0001Paths: List[str] = []
        iSuccessCount = 0
        iFailureCount = 0
        for pszInputPath in arrInputPaths:
            try:
                pszOutputPath = convert_dropped_csv_to_tsv(pszInputPath)
                pszStep0001Path = create_step0001_tsv(pszOutputPath)
                pszStep0001ExcelPath = create_step0001_excel_file(pszStep0001Path)
                arrProcessedStep0001Paths.append(pszStep0001Path)
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
                    + "、"
                    + os.path.basename(pszStep0001ExcelPath)
                )

        arrStep0002Lines: List[str] = []
        arrCheckedPricePaths = set()
        bStep0002Error = False
        for pszStep0001Path in arrProcessedStep0001Paths:
            arrMasterPaths = find_step0001_master_paths(pszStep0001Path)
            if arrMasterPaths is None:
                continue
            pszPriceTsvPath = arrMasterPaths[0]
            if pszPriceTsvPath in arrCheckedPricePaths:
                continue
            arrCheckedPricePaths.add(pszPriceTsvPath)
            try:
                (
                    pszStep0002Path,
                    pszStep0002NameInsertedPath,
                    pszExcelPath,
                    arrMissingPaths,
                ) = create_step0002_tsv_if_ready(
                    pszStep0001Path
                )
            except ConversionError as objException:
                bStep0002Error = True
                arrStep0002Lines.append(
                    "step0002作成失敗: "
                    + os.path.basename(pszPriceTsvPath)
                    + "\n  "
                    + str(objException)
                )
            else:
                if pszStep0002Path:
                    arrStep0002Lines.append(
                        "step0002作成: " + os.path.basename(pszStep0002Path)
                    )
                if pszStep0002NameInsertedPath:
                    arrStep0002Lines.append(
                        "名挿入step0002作成: "
                        + os.path.basename(pszStep0002NameInsertedPath)
                    )
                if pszExcelPath:
                    arrStep0002Lines.append(
                        "Excel作成: " + os.path.basename(pszExcelPath)
                    )
                elif arrMissingPaths:
                    arrStep0002Lines.append(
                        "step0002未作成（必要なstep0001ファイルが不足）:\n  "
                        + "\n  ".join(os.path.basename(pszPath) for pszPath in arrMissingPaths)
                    )

        arrCheckedInventoryPaths = set()
        for pszStep0001Path in arrProcessedStep0001Paths:
            arrMasterPaths = find_step0001_inventory_master_paths(pszStep0001Path)
            if arrMasterPaths is None:
                continue
            pszInventoryTsvPath = arrMasterPaths[0]
            if pszInventoryTsvPath in arrCheckedInventoryPaths:
                continue
            arrCheckedInventoryPaths.add(pszInventoryTsvPath)
            try:
                (
                    pszInventoryStep0002Path,
                    pszInventoryNameInsertedPath,
                    pszInventoryExcelPath,
                    arrInventoryMissingPaths,
                ) = create_inventory_step0002_tsv_if_ready(pszStep0001Path)
            except ConversionError as objException:
                bStep0002Error = True
                arrStep0002Lines.append(
                    "在庫step0002作成失敗: "
                    + os.path.basename(pszInventoryTsvPath)
                    + "\n  "
                    + str(objException)
                )
            else:
                if pszInventoryStep0002Path:
                    arrStep0002Lines.append(
                        "在庫step0002作成: "
                        + os.path.basename(pszInventoryStep0002Path)
                    )
                if pszInventoryNameInsertedPath:
                    arrStep0002Lines.append(
                        "在庫名挿入step0002作成: "
                        + os.path.basename(pszInventoryNameInsertedPath)
                    )
                if pszInventoryExcelPath:
                    arrStep0002Lines.append(
                        "在庫Excel作成: " + os.path.basename(pszInventoryExcelPath)
                    )
                elif arrInventoryMissingPaths:
                    arrStep0002Lines.append(
                        "在庫step0002未作成（必要なstep0001ファイルが不足）:\n  "
                        + "\n  ".join(
                            os.path.basename(pszPath)
                            for pszPath in arrInventoryMissingPaths
                        )
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
        if arrStep0002Lines:
            pszSummary += "\n\n" + "\n\n".join(arrStep0002Lines)
        show_message(pszSummary, iFailureCount > 0 or bStep0002Error)
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
