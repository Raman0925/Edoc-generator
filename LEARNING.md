# Learn Python by understanding this project

Your real task has four steps: read values, put them in the approval template, export PDFs, merge the PDFs. Each step has one owner in the code.

## 1. A mapping is data

```python
Mapping(
    label="Serial number",
    source_sheet="Report",
    source_cell="B4",
    target_sheet="Approval",
    target_cell="F6",
    mode="text",
)
```

A Python `dataclass` groups related values into one object. Instead of five hard-coded assignments, the app stores a list of these mappings. The GUI edits that list, and JSON saves it between runs. That is why you can change which values are extracted without changing Python code.

Read `Mapping` in `edoc/config.py`. Its `validate()` method catches an invalid address before Excel tries to use it. The same validation works whether the mapping came from a screen or a saved file.

**Small exercise:** read `cell_address()`. Work out why `$b$4` becomes `B4`, and why `B4:C4` is rejected.

## 2. Excel is controlled by an adapter

```python
cell = workbook.Worksheets.Item("Report").Range("B4")
value = cell.Value2
```

`Value2` retrieves the underlying value. `Text` retrieves the displayed value, so a serial shown as `00012` can stay `00012` instead of becoming the number `12`. The adapter also handles dates, merged cells, protected destinations and workbook closing.

An adapter is a small boundary between your program and another system. Keeping Excel code in `edoc/excel.py` lets tests substitute a fake adapter while testing the rest of the pipeline.

**Small exercise:** explain which mode you would choose for a sensitivity value and for a serial number. What information does Text mode lose when a number is displayed with only two decimal places?

## 3. A loop handles any number of reports

The important shape of `run_batch()` is:

```python
for report in reports:
    try:
        # Copy files, read mapped values, fill approval, export, merge.
        result = Result(source=str(report), status="success")
    except Exception as error:
        result = Result(source=str(report), status="failed", error=str(error))
```

One report's failure becomes a result rather than stopping the whole batch. Temporary folders hold unfinished work. A completed folder is moved into place only after all steps succeed. That prevents an unfinished PDF from looking like a completed result.

**Small exercise:** find where the code checks `cancel.is_set()`. Explain why stopping finishes the current report first.

## 4. The GUI and Excel use separate threads

Tkinter manages the screen on the main thread. Excel processing runs on a worker thread. A `queue.Queue` passes progress messages back to the screen. The worker never changes Tk widgets directly.

That keeps the app responsive while Excel exports a PDF. `CoInitialize()` and `CoUninitialize()` are Windows COM setup/cleanup calls for the thread controlling Excel.

## 5. Tests describe promises the app must keep

Open `tests/test_engine.py`, especially `test_five_reports_produce_five_complete_documents`. It checks that five reports produce five final PDFs, that approval pages come first, and that original files have the same hashes afterwards.

A hash is a fingerprint of file bytes. Comparing before/after hashes can reveal an unintended change, though it does not tell you whether the original report's calibration data was correct.

Run the tests after any behavior change. Then test an actual workbook on Windows: fake adapters cannot prove Excel layout or Office compatibility.
