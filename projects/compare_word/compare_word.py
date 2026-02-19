"""Word文書比較プログラム

2つのWord文書をWordの組み込み「比較」機能で比較し、
結果と旧文書(Old)・新文書(New)の3ファイルを保存します。

- 比較結果: 変更履歴（消し線・下線）付きの文書
- Old: 挿入（下線）を除去 → 消し線テキストが残る
- New: 削除（消し線）を除去 → 下線テキストが残る
"""

import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox

import win32com.client


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


def reject_insertions(doc):
    """挿入履歴(wdRevisionInsert=1)を全て拒否する。

    1件処理するたびにループを再開する（コレクション変更への対策）。
    VBAの On Error Resume Next 相当の try/except で個別エラーをスキップ。
    """
    WD_REVISION_INSERT = 1
    while True:
        found = False
        for rev in doc.Revisions:
            try:
                if rev.Type == WD_REVISION_INSERT:
                    rev.Reject()
                    found = True
                    break  # コレクション変更後はループを再開
            except Exception:
                pass
        if not found:
            break


def accept_deletions(doc):
    """削除履歴(wdRevisionDelete=2)を全て承諾する。

    1件処理するたびにループを再開する（コレクション変更への対策）。
    VBAの On Error Resume Next 相当の try/except で個別エラーをスキップ。
    """
    WD_REVISION_DELETE = 2
    while True:
        found = False
        for rev in doc.Revisions:
            try:
                if rev.Type == WD_REVISION_DELETE:
                    rev.Accept()
                    found = True
                    break  # コレクション変更後はループを再開
            except Exception:
                pass
        if not found:
            break


def extract_old_new_documents(word, output_path):
    """比較結果ファイルからOld・New文書を生成して保存する。

    VBAマクロ②相当の処理:
    - Old: 挿入履歴(下線)を個別に拒否 → 消し線テキストと書式が残る
    - New: 削除履歴(消し線)を個別に承諾 → 下線テキストと書式が残る
    """
    base, ext = os.path.splitext(output_path)
    old_path = base + "_Old" + ext
    new_path = base + "_New" + ext

    # --- Old文書: 挿入(下線)を拒否して除去 → 消し線テキストが残る ---
    doc_old = word.Documents.Open(output_path)
    reject_insertions(doc_old)
    doc_old.SaveAs2(old_path, FileFormat=12)
    doc_old.Close(SaveChanges=0)
    print(f"Old文書を保存しました: {old_path}")

    # --- New文書: 削除(消し線)を承諾して除去 → 下線テキストが残る ---
    doc_new = word.Documents.Open(output_path)
    accept_deletions(doc_new)
    doc_new.SaveAs2(new_path, FileFormat=12)
    doc_new.Close(SaveChanges=0)
    print(f"New文書を保存しました: {new_path}")

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

        comp_doc.Close(SaveChanges=0)
        comp_doc = None
        revised_doc.Close(SaveChanges=0)
        revised_doc = None
        original_doc.Close(SaveChanges=0)
        original_doc = None

        # 比較結果ファイルからOld・New文書を生成
        old_path, new_path = extract_old_new_documents(word, output_abs)

    except Exception as e:
        print(f"エラーが発生しました: {e}", file=sys.stderr)
        raise

    finally:
        if comp_doc:
            comp_doc.Close(SaveChanges=0)
        if revised_doc:
            revised_doc.Close(SaveChanges=0)
        if original_doc:
            original_doc.Close(SaveChanges=0)
        if word:
            word.Quit()

    return old_path, new_path


def main():
    root = tk.Tk()
    root.withdraw()

    original_path = select_file("元の文書を選択してください")
    if not original_path:
        messagebox.showinfo("キャンセル", "元の文書が選択されませんでした。")
        return

    revised_path = select_file("比較対象の文書を選択してください")
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
            f"Old（消し線テキスト）:\n{old_path}\n\n"
            f"New（下線テキスト）:\n{new_path}",
        )
    except Exception as e:
        messagebox.showerror("エラー", f"比較中にエラーが発生しました:\n{e}")


if __name__ == "__main__":
    main()
