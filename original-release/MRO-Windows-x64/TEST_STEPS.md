# MRO Test Steps

## 1. Build Check

Run:

```powershell
dotnet build .\Mro.sln
dotnet run --project .\tests\Mro.Core.Specs --
```

Expected result:

- build succeeds
- all specs pass

## 2. Create Windows Package

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\package-win.ps1
```

Expected output package:

- `artifacts\packages\MRO-Windows-x64.zip`

## 3. Install On Windows

Prerequisite:

- target machine has `.NET 8 Desktop Runtime` installed

1. Extract `artifacts\packages\MRO-Windows-x64.zip`
2. Open the extracted folder
3. Run `Install-MRO.cmd`

Expected result:

- apps are copied to `%LOCALAPPDATA%\MRO`
- desktop shortcuts are created:
  - `马里奥产品抽检分类系统`
  - `马里奥产品抽检辅助系统`

## 4. Main App Smoke Test

1. Launch `马里奥产品抽检分类系统`
2. Confirm the window opens normally
3. Import the standard workbook using:
   - `检测标准表.xlsx`
4. Confirm status shows shared standard catalog updated
5. Select weighing workbook:
   - `W250905.xlsx` or a helper-generated table2 file
6. Run analyze

Expected result:

- no crash
- product groups or issue summary appears
- logs panel updates

## 5. Helper App Smoke Test

1. Launch `马里奥产品抽检辅助系统`
2. Confirm the window opens normally
3. Reload shared standard store
4. Add one entry:
   - barcode: `6978606880071`
   - batch size: `40`
5. Run analyze
6. Run generate workbook

Expected result:

- no crash
- required sample count is shown
- exported table2 workbook is generated
- logs panel updates

## 6. Round-Trip Smoke Test

1. Take the helper-generated workbook from step 5
2. Open `马里奥产品抽检分类系统`
3. Use that workbook as the weighing input
4. Provide the same batch size for the matching barcode
5. Run export

Expected result:

- export succeeds
- one output Excel file is created
- output contains detail worksheet(s) and one summary worksheet

## 7. Error Path Checks

Check these cases:

1. Helper app with empty barcode and empty batch size
2. Helper app with valid barcode and batch size `0`
3. Main app with no shared standard catalog file
4. Main app with unmatched barcodes in the weighing workbook

Expected result:

- app stays responsive
- issue summary is shown in Chinese
- status message is readable
- log entries record success or failure
