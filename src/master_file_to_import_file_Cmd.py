# -*- coding: utf-8 -*-
"""CSVを取込用TSVへ変換するコマンドライン版です。"""

import argparse
import csv
import os
import sys
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.utils.exceptions import IllegalCharacterError


PARENTHESIS_TRANSLATION_TABLE = str.maketrans("（）", "()")
PRICE_MASTER_STEP0001_PREFIX = "step0001_得意先別商品売上単価マスタ"
CUSTOMER_MASTER_STEP0001_PREFIX = "step0001_得意先マスタ"
PRODUCT_MASTER_STEP0001_PREFIX = "step0001_商品マスタ"


class ConversionError(Exception):
    """CSVやTSVを処理できない場合の例外です。"""


def validate_input_path(pszInputPath: str) -> str:
    pszInputPath = os.path.abspath(pszInputPath)
    if not os.path.isfile(pszInputPath):
        raise ConversionError("入力ファイルが見つかりません: " + pszInputPath)
    if os.path.splitext(pszInputPath)[1].lower() != ".csv":
        raise ConversionError("拡張子がCSVのファイルを指定してください。")
    return pszInputPath


def create_output_path(pszInputPath: str) -> str:
    pszBasePath, _ = os.path.splitext(pszInputPath)
    return pszBasePath + ".tsv"


def convert_csv_to_tsv(pszInputPath: str) -> str:
    """CP932のCSVをBOMなしUTF-8、タブ区切りで上書き出力します。"""
    pszInputPath = validate_input_path(pszInputPath)
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
            with open(pszTsvPath, "r", encoding="utf-8", newline="") as objInput:
                for iRow, arrValues in enumerate(
                    csv.reader(objInput, delimiter="\t"), start=1
                ):
                    for iColumn, pszValue in enumerate(arrValues, start=1):
                        objCell = objSheet.cell(iRow, iColumn, pszValue)
                        objCell.data_type = "s"
                        objCell.number_format = "@"
        objWorkbook.save(pszExcelPath)
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


def main() -> int:
    objParser = argparse.ArgumentParser(
        description="CSVをBOMなしUTF-8のTSVへ変換します。"
    )
    objParser.add_argument("input_csv_path", help="変換するCSVファイルのパス")
    objArgs = objParser.parse_args()

    try:
        pszOutputPath = convert_csv_to_tsv(objArgs.input_csv_path)
        pszStep0001Path = create_step0001_tsv(pszOutputPath)
        (
            pszStep0002Path,
            pszStep0002NameInsertedPath,
            pszExcelPath,
            arrMissingPaths,
        ) = create_step0002_tsv_if_ready(
            pszStep0001Path
        )
    except ConversionError as objException:
        print("エラー: " + str(objException), file=sys.stderr)
        return 1

    print("TSVを作成しました: " + pszOutputPath)
    print("step0001 TSVを作成しました: " + pszStep0001Path)
    if pszStep0002Path:
        print("step0002 TSVを作成しました: " + pszStep0002Path)
    if pszStep0002NameInsertedPath:
        print("名挿入step0002 TSVを作成しました: " + pszStep0002NameInsertedPath)
    if pszExcelPath:
        print("Excelファイルを作成しました: " + pszExcelPath)
    elif arrMissingPaths:
        print("step0002 TSVは未作成です。不足ファイル:")
        for pszMissingPath in arrMissingPaths:
            print("  " + pszMissingPath)
    return 0


if __name__ == "__main__":
    sys.exit(main())
