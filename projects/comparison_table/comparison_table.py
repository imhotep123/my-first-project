"""新旧対照表作成プログラム

3つのWord文書から新旧対照表（A4横）を作成します。
- 左列: 現行規約（元の書式を保持）
- 中列: 改正案（元の書式を保持）
- 右列: 備考（挿入後 MS Pゴシック 8pt に統一）
"""

import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox

import win32com.client

WD_ORIENT_LANDSCAPE = 1
WD_FORMAT_XML_DOCUMENT = 12
WD_COLLAPSE_START = 1


def select_file(title):
    """ファイル選択ダイアログを表示し、選択されたファイルパスを返す。"""
    path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Word文書", "*.docx *.doc"), ("すべてのファイル", "*.*")],
    )
    return path


def select_save_path():
    """保存先ファイルパスをダイアログで指定する。"""
    path = filedialog.asksaveasfilename(
        title="新旧対照表の保存先を指定してください",
        defaultextension=".docx",
        filetypes=[("Word文書", "*.docx"), ("すべてのファイル", "*.*")],
    )
    return path


def insert_file_into_cell(word, cell, file_path):
    """ソース文書の内容をクリップボード経由でテーブルセルに書式保持して挿入する。

    Range.InsertFile は COM バインディングによっては E_FAIL を返す場合があるため、
    ソース文書を開いて Content.Copy() → Range.Paste() のクリップボード方式を採用する。
    """
    # ソース文書を読み取り専用で開き、全コンテンツをクリップボードにコピー
    doc_src = word.Documents.Open(file_path, ReadOnly=True)
    doc_src.Content.Copy()
    doc_src.Close(SaveChanges=0)

    # セル先頭に折り畳んでクリップボードから貼り付け
    rng = cell.Range
    rng.Collapse(WD_COLLAPSE_START)
    rng.Paste()


def create_comparison_table(old_path, new_path, comment_path, output_path):
    """3つのWord文書から新旧対照表を作成して保存する。"""
    word = None
    doc_out = None

    try:
        print("Word起動中...")
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        old_abs = os.path.abspath(old_path)
        new_abs = os.path.abspath(new_path)
        comment_abs = os.path.abspath(comment_path)
        output_abs = os.path.abspath(output_path)

        print("新規文書作成中...")
        doc_out = word.Documents.Add()

        print("ページ設定中（A4横向き）...")
        # 1mm = 2.83465pt (CentimetersToPoints の COM 呼び出しを避け手動計算)
        MM = 2.83465

        ps = doc_out.PageSetup
        ps.PaperSize = 9          # wdPaperA4
        ps.Orientation = WD_ORIENT_LANDSCAPE
        ps.TopMargin = 12.7 * MM  # 12.7mm
        ps.BottomMargin = 12.7 * MM
        ps.LeftMargin = 12.7 * MM
        ps.RightMargin = 12.7 * MM

        print("タイトル行挿入中...")
        # 文書先頭にタイトルを挿入（参照ファイル準拠: MS ゴシック 16pt 太字 中央揃え）
        doc_out.Range(0, 0).InsertAfter("管理規約　新旧対照表")
        title_para_rng = doc_out.Paragraphs(1).Range
        title_para_rng.Font.Name = "MS ゴシック"
        title_para_rng.Font.Size = 16
        title_para_rng.Font.Bold = True
        title_para_rng.ParagraphFormat.Alignment = 1    # wdAlignParagraphCenter
        title_para_rng.ParagraphFormat.SpaceAfter = 8   # 参照ファイル値
        title_para_rng.ParagraphFormat.SpaceBefore = 0
        title_para_rng.ParagraphFormat.LineSpacingRule = 0  # wdLineSpaceSingle

        # 列幅の計算（参照ファイル実測値に合わせ右列=92pt）
        # 利用可能幅 = 297mm - 2×12.7mm = 271.6mm
        available_width_pt = (297 - 2 * 12.7) * MM
        right_col_width = 92                                       # 参照ファイル: 92.15pt
        left_mid_col_width = (available_width_pt - right_col_width) / 2

        print("表作成中（2行 × 3列）...")
        # タイトル段落の末尾位置に表を追加
        doc_end = doc_out.Range()
        doc_end.Collapse(0)  # wdCollapseEnd = 0
        table = doc_out.Tables.Add(doc_end, 2, 3)
        table.AllowAutoFit = False
        # 表全体の幅をページ内有効幅に固定（はみ出し防止）
        table.PreferredWidthType = 3  # wdPreferredWidthPoints
        table.PreferredWidth = available_width_pt

        print("列幅設定中...")
        table.Columns(1).Width = left_mid_col_width
        table.Columns(2).Width = left_mid_col_width
        table.Columns(3).Width = right_col_width

        print("ヘッダー行設定中...")
        # テキスト
        table.Cell(1, 1).Range.Text = "現行規約"
        table.Cell(1, 2).Range.Text = "改正案"
        table.Cell(1, 3).Range.Text = "備考"
        # 書式（参照ファイル準拠: ＭＳ 明朝 10pt・グレー背景・中央/中央/左揃え）
        HEADER_GRAY = 15132390   # #E6E6E6 (RGB 230,230,230)
        for col in range(1, 4):
            cell = table.Cell(1, col)
            cell.Range.Font.Name = "ＭＳ 明朝"
            cell.Range.Font.Size = 10
            cell.Range.Font.Bold = False
            cell.Shading.BackgroundPatternColor = HEADER_GRAY
            cell.Range.ParagraphFormat.SpaceBefore = 0
            cell.Range.ParagraphFormat.SpaceAfter = 0
        table.Cell(1, 1).Range.ParagraphFormat.Alignment = 1  # Center
        table.Cell(1, 2).Range.ParagraphFormat.Alignment = 1  # Center
        table.Cell(1, 3).Range.ParagraphFormat.Alignment = 0  # Left

        print(f"旧規約を挿入中: {old_abs}")
        insert_file_into_cell(word, table.Cell(2, 1), old_abs)
        print("  → 完了")

        print(f"改正案を挿入中: {new_abs}")
        insert_file_into_cell(word, table.Cell(2, 2), new_abs)
        print("  → 完了")

        print(f"コメントを挿入中: {comment_abs}")
        insert_file_into_cell(word, table.Cell(2, 3), comment_abs)
        print("  → 完了")

        print("コンテンツ行の書式設定中...")
        # 参照ファイル準拠
        # 左列・中列: ＭＳ 明朝 10pt・行間12pt固定・段落後0pt
        WD_LINE_SPACE_EXACTLY = 4
        for col in (1, 2):
            rng = table.Cell(2, col).Range
            rng.Font.Name = "ＭＳ 明朝"
            rng.Font.Size = 10
            rng.ParagraphFormat.LineSpacingRule = WD_LINE_SPACE_EXACTLY
            rng.ParagraphFormat.LineSpacing = 12
            rng.ParagraphFormat.SpaceBefore = 0
            rng.ParagraphFormat.SpaceAfter = 0
        # 右列: ＭＳ Ｐゴシック 8pt・行間10pt固定・段落後0pt
        rng3 = table.Cell(2, 3).Range
        rng3.Font.Name = "ＭＳ Ｐゴシック"
        rng3.Font.Size = 8
        rng3.ParagraphFormat.LineSpacingRule = WD_LINE_SPACE_EXACTLY
        rng3.ParagraphFormat.LineSpacing = 10
        rng3.ParagraphFormat.SpaceBefore = 0
        rng3.ParagraphFormat.SpaceAfter = 0

        print(f"保存中: {output_abs}")
        doc_out.SaveAs2(output_abs, FileFormat=WD_FORMAT_XML_DOCUMENT)
        print("新旧対照表を保存しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}", file=sys.stderr)
        raise

    finally:
        if doc_out:
            try:
                doc_out.Close(SaveChanges=0)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass


def main():
    root = tk.Tk()
    root.withdraw()

    old_path = select_file("旧規約のファイル名を指定してください")
    if not old_path:
        messagebox.showinfo("キャンセル", "旧規約のファイルが選択されませんでした。")
        return

    new_path = select_file("改正案のファイル名を指定してください")
    if not new_path:
        messagebox.showinfo("キャンセル", "改正案のファイルが選択されませんでした。")
        return

    comment_path = select_file("コメントのファイル名を指定してください")
    if not comment_path:
        messagebox.showinfo("キャンセル", "コメントのファイルが選択されませんでした。")
        return

    output_path = select_save_path()
    if not output_path:
        messagebox.showinfo("キャンセル", "保存先が指定されませんでした。")
        return

    try:
        create_comparison_table(old_path, new_path, comment_path, output_path)
        messagebox.showinfo(
            "完了",
            f"新旧対照表を作成しました:\n{output_path}",
        )
    except Exception as e:
        messagebox.showerror("エラー", f"作成中にエラーが発生しました:\n{e}")


if __name__ == "__main__":
    main()
