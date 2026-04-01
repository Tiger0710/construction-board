"""VBA付きテンプレート生成 (初回のみ実行)

template/gantt_template.xlsm を生成する。
create_input_system.py が openpyxl + keep_vba=True で再利用。

必要条件:
  - Windows + Excel
  - Excelオプション → トラストセンター → 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」有効
"""
import os
import sys
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "template")
TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, "gantt_template.xlsm")

# === VBA: ThisWorkbook ===
VBA_THISWORKBOOK = """
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    SyncGanttToDaily
End Sub
"""

# === VBA: SyncModule ===
VBA_MODULE = """
Public Sub SyncGanttToDaily()
    Dim wsG As Worksheet, wsD As Worksheet
    On Error Resume Next
    Set wsG = ThisWorkbook.Worksheets("ガントチャート")
    Set wsD = ThisWorkbook.Worksheets("日次入力")
    On Error GoTo 0
    If wsG Is Nothing Or wsD Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanUp

    ' === 1. Read existing daily data ===
    Dim existing As Object
    Set existing = CreateObject("Scripting.Dictionary")
    Dim lastRowD As Long
    lastRowD = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRowD
        If Not IsEmpty(wsD.Cells(r, 1).Value) Then
            Dim cv As String, tv As String
            cv = Trim(wsD.Cells(r, 2).Value & "")
            tv = Trim(wsD.Cells(r, 3).Value & "")
            If cv <> "" And tv <> "" Then
                Dim ek As String
                ek = Format(wsD.Cells(r, 1).Value, "yyyy-mm-dd") & "|" & cv & "|" & tv
                existing(ek) = Trim(wsD.Cells(r, 4).Value & "") & Chr(0) & _
                               Trim(wsD.Cells(r, 5).Value & "") & Chr(0) & _
                               Trim(wsD.Cells(r, 6).Value & "")
            End If
        End If
    Next r

    ' === 2. Generate entries from Gantt ===
    Dim entries As New Collection
    Dim lastRowG As Long
    lastRowG = wsG.Cells(wsG.Rows.Count, 1).End(xlUp).Row

    For r = 3 To lastRowG
        Dim gc As String, gt As String
        gc = Trim(wsG.Cells(r, 1).Value & "")
        gt = Trim(wsG.Cells(r, 2).Value & "")
        If gc = "" Or gt = "" Then GoTo NextRow

        Dim sd As Variant, ed As Variant
        sd = wsG.Cells(r, 7).Value
        ed = wsG.Cells(r, 8).Value
        If Not IsDate(sd) Or Not IsDate(ed) Then GoTo NextRow
        If CDate(ed) < CDate(sd) Then GoTo NextRow

        Dim d As Date
        For d = CDate(sd) To CDate(ed)
            Dim lk As String
            lk = Format(d, "yyyy-mm-dd") & "|" & gc & "|" & gt
            Dim dn As String, wc As String, pr As String
            dn = "": wc = "": pr = ""
            If existing.Exists(lk) Then
                Dim parts() As String
                parts = Split(existing(lk), Chr(0))
                If UBound(parts) >= 0 Then dn = parts(0)
                If UBound(parts) >= 1 Then wc = parts(1)
                If UBound(parts) >= 2 Then pr = parts(2)
            End If
            entries.Add Array(d, gc, gt, dn, wc, pr)
        Next d
NextRow:
    Next r

    ' === 3. Clear and batch write ===
    If lastRowD >= 2 Then wsD.Range("A2:F" & lastRowD).ClearContents

    Dim cnt As Long
    cnt = entries.Count
    If cnt > 0 Then
        Dim arr() As Variant
        ReDim arr(1 To cnt, 1 To 6)
        Dim i As Long
        For i = 1 To cnt
            Dim e As Variant
            e = entries(i)
            arr(i, 1) = e(0)
            arr(i, 2) = e(1)
            arr(i, 3) = e(2)
            arr(i, 4) = e(3)
            arr(i, 5) = e(4)
            arr(i, 6) = e(5)
        Next i
        wsD.Range("A2:F" & (cnt + 1)).Value = arr
        wsD.Range("A2:A" & (cnt + 1)).NumberFormat = "M/D"
    End If

CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
"""


def main():
    try:
        import win32com.client as win32
    except ImportError:
        print("ERROR: pywin32 が必要: pip install pywin32")
        sys.exit(1)

    os.makedirs(TEMPLATE_DIR, exist_ok=True)

    if os.path.exists(TEMPLATE_PATH):
        os.remove(TEMPLATE_PATH)

    print("Excel起動...")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Add()

        # --- ガントチャート sheet ---
        ws = wb.Worksheets(1)
        ws.Name = "ガントチャート"

        labels = ["客先", "工事件名", "現場担当者", "安品担当者",
                  "協力会社名", "協力会社担当者", "開始", "終了"]
        widths = [12, 38, 10, 10, 16, 10, 10, 10]
        for i, (label, w) in enumerate(zip(labels, widths), 1):
            ws.Cells(2, i).Value = label
            ws.Columns(i).ColumnWidth = w

        ws.Rows(1).RowHeight = 28

        # --- 日次入力 sheet ---
        ws_d = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws_d.Name = "日次入力"

        d_headers = ["日付", "客先", "工事件名", "昼/夜", "工事内容", "重点工事"]
        d_widths = [12, 12, 28, 8, 40, 10]
        for i, (h, w) in enumerate(zip(d_headers, d_widths), 1):
            ws_d.Cells(1, i).Value = h
            ws_d.Columns(i).ColumnWidth = w

        # Delete extra sheets
        while wb.Worksheets.Count > 2:
            for i in range(wb.Worksheets.Count, 0, -1):
                name = wb.Worksheets(i).Name
                if name not in ("ガントチャート", "日次入力"):
                    wb.Worksheets(i).Delete()
                    break

        # --- VBA injection ---
        print("VBAコード挿入...")
        vb_project = wb.VBProject

        # ThisWorkbook
        tb = vb_project.VBComponents("ThisWorkbook")
        tb.CodeModule.AddFromString(VBA_THISWORKBOOK.strip())

        # Standard module
        mod = vb_project.VBComponents.Add(1)  # vbext_ct_StdModule
        mod.Name = "SyncModule"
        mod.CodeModule.AddFromString(VBA_MODULE.strip())

        # Activate ガントチャート
        wb.Worksheets("ガントチャート").Activate()

        # Save as xlsm (FileFormat=52)
        abs_path = os.path.abspath(TEMPLATE_PATH)
        print(f"保存: {abs_path}")
        wb.SaveAs(abs_path, FileFormat=52)
        wb.Close(False)

        print("テンプレート作成完了!")

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        excel.Quit()
        time.sleep(1)


if __name__ == "__main__":
    main()
