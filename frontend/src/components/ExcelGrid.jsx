import React, { useMemo, useCallback, useRef, useState } from 'react';
import { AgGridReact } from 'ag-grid-react';
import { AllCommunityModule, ModuleRegistry } from 'ag-grid-community';
import ExcelSetFilter from './ExcelSetFilter';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-alpine.css';

ModuleRegistry.registerModules([AllCommunityModule]);

export default function ExcelGrid({ rowData, columnDefs, onGridReady, onCellValueChanged, gridRef, searchText }) {
  const defaultColDef = useMemo(() => {
    return {
      flex: 1,
      minWidth: 100,
      editable: true,
      resizable: true,
      sortable: true,
      filter: ExcelSetFilter,
      columnMenu: 'new',
    };
  }, []);

  const getRowClass = (params) => {
    if (params.data && params.data._isEditedRow) {
      return 'edited-row';
    }
    return '';
  };

  const cellClassRules = {
    'edited-cell': (params) => {
      // We can use a custom property added during edit to highlight
      return params.data && params.data._editedCells && params.data._editedCells[params.colDef.field];
    }
  };

  // Inject cellClassRules to all columnDefs
  const updatedColumnDefs = useMemo(() => {
    return columnDefs.map(col => ({
      ...col,
      cellClassRules,
    }));
  }, [columnDefs]);

  return (
    <div className="ag-theme-alpine grid-container" style={{ height: 600, width: '100%' }}>
      <AgGridReact
        ref={gridRef}
        rowData={rowData}
        columnDefs={updatedColumnDefs}
        defaultColDef={defaultColDef}
        onGridReady={onGridReady}
        onCellValueChanged={onCellValueChanged}
        quickFilterText={searchText}
        theme="legacy"
        rowSelection={{ mode: 'multiRow', enableClickSelection: false }}
        getRowClass={getRowClass}
        undoRedoCellEditing={true}
        undoRedoCellEditingLimit={20}
        enableCellChangeFlash={true}
      />
    </div>
  );
}
