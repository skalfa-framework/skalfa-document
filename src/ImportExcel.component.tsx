"use client"
import { Icon, type IconName } from "@skalfa/skalfa-icon";
;

import { useMemo, useState } from "react";
import ExcelJS from "exceljs";
import { api, FetchControlType, registry } from "@skalfa/skalfa-app-core";

// Dynamically resolve UI components and context from registry to avoid compile-time dependencies
const TableComponent = (props: any) => {
  const Comp = registry.get("TableComponent");
  return Comp ? <Comp {...props} /> : null;
};
const ButtonComponent = (props: any) => {
  const Comp = registry.get("ButtonComponent");
  return Comp ? <Comp {...props} /> : null;
};
const IconButtonComponent = (props: any) => {
  const Comp = registry.get("IconButtonComponent");
  return Comp ? <Comp {...props} /> : null;
};
const SelectComponent = (props: any) => {
  const Comp = registry.get("SelectComponent");
  return Comp ? <Comp {...props} /> : null;
};
const ModalComponent = (props: any) => {
  const Comp = registry.get("ModalComponent");
  return Comp ? <Comp {...props} /> : null;
};

const useToggleContext = () => {
  const hook = registry.get("useToggleContext");
  return hook ? hook() : { toggle: {} as Record<string, any>, setToggle: () => {} };
};


export type ImportExcelColumnControlType = {
  label: string;
  selector: string;
};

type ImportColumn = {
  selector: string;
  label: string;
  source: string | null;
};

type ImportExcelProps = {
  columnControl: ImportExcelColumnControlType[];
  onSubmit?: (rows: any[]) => void;
  submitControl?: FetchControlType;
  fetchControl?: FetchControlType;
};



function numberToExcelColumn(index: number): string
{
  let column = "";
  let n = index;

  while (n >= 0) {
    column = String.fromCharCode((n % 26) + 65) + column;
    n = Math.floor(n / 26) - 1;
  }

  return column;
}



export function ImportExcel({ columnControl, onSubmit, submitControl, fetchControl }: ImportExcelProps) {
  const { toggle, setToggle }  =  useToggleContext()

  const [columns, setColumns]  =  useState<ImportColumn[]>([]);
  const [rows, setRows]        =  useState<Record<string, any>[]>([]);
  const [loaded, setLoaded]    =  useState(false);

  const [processing, setProcessing]  =  useState(false);
  const [progress, setProgress]      =  useState({ success: 0, failed: 0, total: 0 });
  const [errors, setErrors]          =  useState<Record<number, string>>({});


  const handleImportFile = async (file: File) => {
    const workbook  =  new ExcelJS.Workbook();
    const buffer    =  await file.arrayBuffer();

    await workbook.xlsx.load(buffer);
    const sheet = workbook.worksheets[0];

    const excelColumns: ImportColumn[] = [];
    sheet.getRow(1).eachCell((_, colIndex) => {
      const label = numberToExcelColumn(colIndex - 1);

      excelColumns.push({
        selector: label,
        label: label,
        source: null,
      });
    });

    const excelRows: Record<string, any>[] = [];
    sheet.eachRow((row, rowIndex) => {
      if (rowIndex === 1) return;

      const item: Record<string, any> = {};
      excelColumns.forEach((col, i) => {
        item[col.selector] = row.getCell(i + 1).value;
      });

      excelRows.push(item);
    });

    setColumns(excelColumns);
    setRows(excelRows);
    setLoaded(true);
  };

  const getColumnLabel = (source: string | null) => {
    if (!source) return "";

    const found = columnControl?.find(c => c.selector === source);

    return found?.label ?? source;
  };


  const tableColumns = useMemo(() => {
    return columns?.map((c => ({
      ...c,
      label: <div className="w-full text-center">{c.label}</div>
    })));
  }, [columns]);


  const tableData = useMemo(() => {
    if (!loaded) return [];

    const mappingRow: Record<string, any> = {};

    columns.forEach(col => {
      mappingRow[col.selector] = (
        <>
          <div className="flex justify-between">
            <p className="font-semibold">{getColumnLabel(col.source) || <p className="text-light-foreground">-- PILIH KOLOM --</p>}</p>

            <IconButtonComponent
              icon="solid/edit"
              size="xs"
              paint="warning"
              variant="outline"
              disabled={columns.length <= 1}
              onClick={() => setToggle("MODAL_FIELD_IMPORT", { selector: col.selector, value: col.source })}
            />
          </div>
        </>
      );
    });

    return [mappingRow, ...rows];
  }, [columns, rows, loaded, columnControl]);


  const generatePayload = () => {
    return rows.map(row => {
      const payload: Record<string, any> = {};

      columns.forEach(col => {
        if (col.source) {
          payload[col.source] = row[col.selector];
        }
      });

      return payload;
    });
  };

  const handleSubmit = async () => {
    const payload = generatePayload();

    if (onSubmit) {
      onSubmit(payload);
      return;
    }

    setProcessing(true);
    setProgress({ success: 0, failed: 0, total: payload.length });
    setErrors({});

    if (submitControl?.path) {
      try {
        await api({
          path     :  submitControl?.path,
          method   :  "POST",
          payload  :  { data: payload },
        });

        setProgress({ success: payload.length, failed: 0, total: payload.length });
        setToggle("MODAL_IMPORT_SUCCESS", true);

        setTimeout(() => {
          setToggle(Object.keys(toggle).find(k => k.startsWith("MODAL_IMPORT_")) || "", false);
        }, 1500);

      } catch (e: any) {
        setProgress({ success: 0, failed: payload.length, total: payload.length });
        setErrors({ 0: e.message || "Bulk import failed" });
      }
    } else if (fetchControl) {
      let successCount  =  0;
      let failedCount   =  0;

      for (let i = 0; i < payload.length; i++) {
        const item = payload[i];

        try {
          await api({
            ...fetchControl,
            method: "POST",
            payload: item,
          });
          successCount++;
        } catch (e: any) {
          failedCount++;
          setErrors(prev => ({ ...prev, [i]: e.message || "Failed" }));
        }

        setProgress({ success: successCount, failed: failedCount, total: payload.length });
      }
    }

    setProcessing(false);
  };

  return (
    <>

      {!loaded && (
        <div className="p-4 relative">
          <input
            type="file"
            accept=".xlsx"
            onChange={e =>
            {
              const file = e.target.files?.[0];
              if (file) handleImportFile(file);
            }}
            className="text-transparent bg-background w-full aspect-video border border-dashed relative file:hidden placeholder:hidden rounded-md cursor-pointer"
          />

          <div className="absolute top-1/2 left-1/2 -translate-1/2 text-light-foreground">
            Pilih atau tarik file excel di sini
          </div>
        </div>
      )}

      {loaded && (
        <TableComponent
          controlBar={false}
          columns={tableColumns}
          data={tableData}
          pagination={false}
          noIndex
          className="p-4 bg-background row::bg-background row::border-0 row::gap-0 row::!hover:bg-background column::p-2 column::border head-column::p-2 head-column::border"
        />
      )}

      {loaded && !processing && (
        <div className="px-4 mt-8">
          <ButtonComponent
            label="Import Data"
            block
            onClick={handleSubmit}
          />
        </div>
      )}

      {processing && (
        <div className="px-4 mt-8 flex flex-col gap-2">
          <div className="flex justify-between text-sm">
            <span>Processing: {progress.success + progress.failed} / {progress.total}</span>
            {progress.failed > 0 && <span className="text-danger">{progress.failed} Failed</span>}
          </div>
          <div className="w-full h-2 bg-secondary rounded-full overflow-hidden">
            <div
              className="h-full bg-primary transition-all duration-300"
              style={{ width: `${((progress.success + progress.failed) / progress.total) * 100}%` }}
            />
          </div>
        </div>
      )}

      {Object.keys(errors).length > 0 && (
        <div className="px-4 mt-4 max-h-40 overflow-y-auto">
          <p className="text-danger font-semibold mb-2">Errors Details:</p>
          {Object.entries(errors).map(([idx, msg]) => (
            <div key={idx} className="text-xs text-danger">
              Row {Number(idx) + 1}: {msg}
            </div>
          ))}
        </div>
      )}


      <ModalComponent
        show={!!toggle["MODAL_FIELD_IMPORT"]}
        onClose={() => setToggle("MODAL_FIELD_IMPORT", false)}
        title="Pilih Kolom"
        footer={
          <div className="flex justify-end">
            <ButtonComponent
              label="Terapkan"
              onClick={() => {
                if (!!(toggle["MODAL_FIELD_IMPORT"] as { value: string })?.value) {
                  setColumns(prev =>
                    prev.map(c => c.selector === (toggle["MODAL_FIELD_IMPORT"] as { selector: string })?.selector ? { ...c, source: String((toggle["MODAL_FIELD_IMPORT"] as { value: string })?.value) } : c)
                  )
                }
                setToggle("MODAL_FIELD_IMPORT", false)
              }}
            />
          </div>
        }
      >
        <div className="p-4">
          <SelectComponent
            name={`column_${(toggle["MODAL_FIELD_IMPORT"] as { selector: string })?.selector}`}
            placeholder="Pilih kolom data..."
            value={(toggle["MODAL_FIELD_IMPORT"] as { value: string })?.value ?? ""}
            onChange={(e: any) => setToggle("MODAL_FIELD_IMPORT", { ...(toggle["MODAL_FIELD_IMPORT"] as object), value: e })}
            options={columnControl.map(c => ({
              label: c.label,
              value: c.selector,
            }))}
          />
        </div>
      </ModalComponent>
    </>
  );
}
