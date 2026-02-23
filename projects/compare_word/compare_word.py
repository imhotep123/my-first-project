"""Word文書比較プログラム

2つのWord文書をWordの組み込み「比較」機能で比較し、
結果と旧文書(消し線)・新文書(アンダー)の3ファイルを保存します。

- 比較結果: 変更履歴（消し線・下線）付きの文書
- 消し線: 挿入（下線）を削除し、削除箇所を消し線書式（赤字）に変換して固定
- アンダー: 削除（消し線）を削除し、挿入箇所を下線書式（青字）に変換して固定
"""

import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox

import win32com.client

WD_REVISION_INSERT = 1
WD_REVISION_DELETE = 2
WD_UNDERLINE_SINGLE = 1
WD_COLOR_RED = 6
WD_COLOR_BLUE = 2


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
        title="比較結果の保存先を指定",
        defaultextension=".docx",
        filetypes=[("Word文書", "*.docx"), ("すべてのファイル", "*.*")],
    )
    return path


def process_old_document(doc):
    """旧文書処理: 挿入履歴を削除し、削除履歴を消し線書式（赤字）に変換して固定する。

    VBAマクロ準拠:
    - wdRevisionInsert: rev.Range.Delete() で挿入テキストを削除
    - wdRevisionDelete: StrikeThrough + 赤字を適用してから Reject() で書式として固定
    - その他: Accept()
    """
    doc.TrackRevisions = False
    while doc.Revisions.Count > 0:
        rev = doc.Revisions(1)
        try:
            if rev.Type == WD_REVISION_INSERT:
                rev.Range.Delete()
            elif rev.Type == WD_REVISION_DELETE:
                rev.Range.Font.StrikeThrough = True
                rev.Range.Font.ColorIndex = WD_COLOR_RED
                rev.Reject()
            else:
                rev.Accept()
        except Exception:
            try:
                rev.Accept()
            except Exception:
                break


def process_new_document(doc):
    """新文書処理: 削除履歴を削除し、挿入履歴を下線書式（青字）に変換して固定する。

    VBAマクロ準拠:
    - wdRevisionDelete: rev.Range.Delete() で削除テキストを削除
    - wdRevisionInsert: Underline + 青字を適用してから Accept() で書式として固定
    - その他: Accept()
    """
    doc.TrackRevisions = False
    while doc.Revisions.Count > 0:
        rev = doc.Revisions(1)
        try:
            if rev.Type == WD_REVISION_DELETE:
                rev.Range.Delete()
            elif rev.Type == WD_REVISION_INSERT:
                rev.Range.Font.Underline = WD_UNDERLINE_SINGLE
                rev.Range.Font.ColorIndex = WD_COLOR_BLUE
                rev.Accept()
            else:
                rev.Accept()
        except Exception:
            try:
                rev.Accept()
            except Exception:
                break


def extract_old_new_documents(word, output_path):
    """比較結果ファイルから消し線・アンダー文書を生成して保存する。

    呼び出し元で comp_doc を先に閉じてから本関数を呼ぶこと。
    同一ファイルを Documents.Open すると既存オブジェクトが返され
    Close 後に comp_doc が無効になる（RPC_E_DISCONNECTED）のを防ぐ。

    VBAマクロ準拠:
    - 消し線: 挿入履歴を削除し、削除履歴を消し線書式（赤字）に変換して固定
    - アンダー: 削除履歴を削除し、挿入履歴を下線書式（青字）に変換して固定
    """
    base, ext = os.path.splitext(output_path)
    old_path = base + "_消し線" + ext
    new_path = base + "_アンダ" + ext

    # --- 消し線文書: 挿入を削除し、削除を消し線書式（赤字）に変換 ---
    doc_old = word.Documents.Open(output_path)
    process_old_document(doc_old)
    doc_old.SaveAs2(old_path, FileFormat=12)
    doc_old.Close(SaveChanges=0)
    print(f"消し線文書を保存しました: {old_path}")

    # --- アンダー文書: 削除を削除し、挿入を下線書式（青字）に変換 ---
    doc_new = word.Documents.Open(output_path)
    process_new_document(doc_new)
    doc_new.SaveAs2(new_path, FileFormat=12)
    doc_new.Close(SaveChanges=0)
    print(f"アンダー文書を保存しました: {new_path}")

    return old_path, new_path


def compare_documents(original_path, revised_path, output_path):
    """Wordの比較機能を使って2つの文書を比較し、3ファイルを保存する。"""
    word = None
    original_doc = None
    revised_doc = None
    comp_doc = None

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        original_abs = os.path.abspath(original_path)
        revised_abs = os.path.abspath(revised_path)
        output_abs = os.path.abspath(output_path)

        original_doc = word.Documents.Open(original_abs, ReadOnly=True)
        revised_doc = word.Documents.Open(revised_abs, ReadOnly=True)

        # CompareDocuments: Wordの「比較」機能を呼び出す
        # wdCompareDestinationNew = 2 (新しい文書に結果を出力)
        comp_doc = word.CompareDocuments(
            OriginalDocument=original_doc,
            RevisedDocument=revised_doc,
            Destination=2,
            CompareFormatting=True,
            CompareCaseChanges=True,
            CompareWhitespace=True,
            CompareTables=True,
            CompareHeaders=True,
            CompareFootnotes=True,
            CompareTextboxes=True,
            CompareFields=True,
            CompareComments=True,
        )

        # wdFormatXMLDocument = 12
        comp_doc.SaveAs2(output_abs, FileFormat=12)
        print(f"比較結果を保存しました: {output_abs}")

        # 元文書・改訂文書を閉じる
        revised_doc.Close(SaveChanges=0)
        revised_doc = None
        original_doc.Close(SaveChanges=0)
        original_doc = None

        # comp_doc を先に閉じる
        # ※ extract_old_new_documents 内で同一ファイルを Documents.Open すると
        #   Word は既存の comp_doc オブジェクトをそのまま返すため、
        #   doc_old.Close() が comp_doc を無効にし RPC_E_DISCONNECTED が発生する。
        #   先に閉じることで新規オブジェクトとして開き直せる。
        comp_doc.Close(SaveChanges=0)
        comp_doc = None

        old_path, new_path = extract_old_new_documents(word, output_abs)

    except Exception as e:
        print(f"エラーが発生しました: {e}", file=sys.stderr)
        raise

    finally:
        # 最後の文書が閉じられた時点でWordが自動終了している場合があるため、
        # 各クリーンアップ操作を個別に try-except で保護する
        # (RPC_E_DISCONNECTED: -2147417848 を無視)
        if comp_doc:
            try:
                comp_doc.Close(SaveChanges=0)
            except Exception:
                pass
        if revised_doc:
            try:
                revised_doc.Close(SaveChanges=0)
            except Exception:
                pass
        if original_doc:
            try:
                original_doc.Close(SaveChanges=0)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass

    return old_path, new_path


def main():
    root = tk.Tk()
    root.withdraw()

    original_path = select_file("元の文書（旧規約）を選択してください")
    if not original_path:
        messagebox.showinfo("キャンセル", "元の文書が選択されませんでした。")
        return

    revised_path = select_file("比較対象の文書（改正案）を選択してください")
    if not revised_path:
        messagebox.showinfo("キャンセル", "比較対象の文書が選択されませんでした。")
        return

    output_path = select_save_path()
    if not output_path:
        messagebox.showinfo("キャンセル", "保存先が指定されませんでした。")
        return

    try:
        old_path, new_path = compare_documents(
            original_path, revised_path, output_path
        )
        messagebox.showinfo(
            "完了",
            f"比較結果（変更履歴付き）:\n{output_path}\n\n"
            f"消し線（旧規約）:\n{old_path}\n\n"
            f"アンダー（改正案）:\n{new_path}",
        )
    except Exception as e:
        messagebox.showerror("エラー", f"比較中にエラーが発生しました:\n{e}")


if __name__ == "__main__":
    main()
