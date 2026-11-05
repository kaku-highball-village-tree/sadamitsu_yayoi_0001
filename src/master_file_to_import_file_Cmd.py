# -*- coding: utf-8 -*-
"""商品マスタCSVを取込用TSVへ変換するコマンドライン版です。"""

import argparse
import csv
import os
import sys


PRODUCT_MASTER_NAME = "商品マスタ"
PARENTHESIS_TRANSLATION_TABLE = str.maketrans("（）", "()")


class ConversionError(Exception):
    """CSVを商品マスタTSVへ変換できない場合の例外です。"""


def validate_input_path(pszInputPath: str) -> str:
    pszInputPath = os.path.abspath(pszInputPath)
    if not os.path.isfile(pszInputPath):
        raise ConversionError("入力ファイルが見つかりません: " + pszInputPath)
    if os.path.splitext(pszInputPath)[1].lower() != ".csv":
        raise ConversionError("拡張子がCSVのファイルを指定してください。")
    if PRODUCT_MASTER_NAME not in os.path.basename(pszInputPath):
        raise ConversionError("ファイル名に「商品マスタ」を含むCSVを指定してください。")
    return pszInputPath


def create_output_path(pszInputPath: str) -> str:
    pszBasePath, _ = os.path.splitext(pszInputPath)
    return pszBasePath + ".tsv"


def convert_csv_to_tsv(pszInputPath: str) -> str:
    """CP932の43列CSVをBOMなしUTF-8、タブ区切りで上書き出力します。"""
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


def main() -> int:
    objParser = argparse.ArgumentParser(
        description="商品マスタCSVをBOMなしUTF-8のTSVへ変換します。"
    )
    objParser.add_argument("input_csv_path", help="変換する商品マスタCSVのパス")
    objArgs = objParser.parse_args()

    try:
        pszOutputPath = convert_csv_to_tsv(objArgs.input_csv_path)
        pszStep0001Path = create_step0001_tsv(pszOutputPath)
    except ConversionError as objException:
        print("エラー: " + str(objException), file=sys.stderr)
        return 1

    print("TSVを作成しました: " + pszOutputPath)
    print("step0001 TSVを作成しました: " + pszStep0001Path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
