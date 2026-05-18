import React, { useState, useRef, useCallback } from 'react';
import axios from 'axios';
import { Upload, Download, Scissors, Trash2, Search, FileSpreadsheet, Eraser } from 'lucide-react';
import ExcelGrid from './components/ExcelGrid';

function App() {
  const [file, setFile] = useState(null);
  const [rowData, setRowData] = useState([]);
  const [columnDefs, setColumnDefs] = useState([]);
  const [sheetName, setSheetName] = useState('');
  const [searchText, setSearchText] = useState('');
  const [loading, setLoading] = useState(false);
  const [gridApi, setGridApi] = useState(null);
  const gridRef = useRef();
  const fileInputRef = useRef(null);

  const handleGridReady = (params) => {
    setGridApi(params.api);
  };

  const handleFileUpload = async (e) => {
    const selectedFile = e.target.files[0];
    if (!selectedFile) return;

    const validExtensions = ['.xlsx', '.xls', '.csv', '.ods', '.xlsb', '.xlsm', '.txt'];
    const extension = '.' + selectedFile.name.split('.').pop().toLowerCase();
    
    if (!validExtensions.includes(extension)) {
      alert('Please upload a valid spreadsheet file (.xlsx, .xls, .csv, .ods, etc).');
      return;
    }

    setFile(selectedFile);
    setLoading(true);

    const formData = new FormData();
    formData.append('file', selectedFile);

    try {
      const response = await axios.post('http://localhost:3000/api/upload', formData, {
        headers: { 'Content-Type': 'multipart/form-data' },
      });

      const { headers, rows, sheetName: sName } = response.data;
      
      const rowNumberCol = {
        headerName: '',
        valueGetter: 'node.rowIndex + 1',
        width: 60,
        minWidth: 60,
        maxWidth: 60,
        pinned: 'left',
        suppressMenu: true,
        filter: false,
        sortable: false,
        editable: false,
        resizable: false,
        cellStyle: { background: '#f3f2f1', color: '#605e5c', textAlign: 'center', fontWeight: '500', borderRight: '1px solid #e1dfdd' }
      };

      // Create AG Grid column definitions
      const newColDefs = [
        rowNumberCol,
        ...headers.map(header => ({
          field: header,
          headerName: header,
        }))
      ];

      // Add internal tracking properties to rows
      const initialRows = rows.map(r => ({ ...r, _editedCells: {} }));

      setColumnDefs(newColDefs);
      setRowData(initialRows);
      setSheetName(sName);
    } catch (error) {
      console.error('Upload error:', error);
      alert('Failed to process the file.');
    } finally {
      setLoading(false);
    }
  };

  const handleClearFile = () => {
    setFile(null);
    setRowData([]);
    setColumnDefs([]);
    setSheetName('');
    setSearchText('');
    setGridApi(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleCellValueChanged = useCallback((params) => {
    const { colDef, data, oldValue, newValue } = params;
    if (oldValue !== newValue) {
      data._editedCells = { ...data._editedCells, [colDef.field]: true };
      data._isEditedRow = true;
      params.api.applyTransaction({ update: [data] });
    }
  }, []);

  const trimText = () => {
    if (!gridApi) return;
    const rowNodes = [];
    gridApi.forEachNode(node => rowNodes.push(node));

    const updatedRows = [];
    rowNodes.forEach(node => {
      let changed = false;
      const data = node.data;
      
      Object.keys(data).forEach(key => {
        if (key.startsWith('_')) return; // Skip internal keys
        if (typeof data[key] === 'string') {
          const trimmed = data[key].trim();
          if (trimmed !== data[key]) {
            data[key] = trimmed;
            data._editedCells = { ...data._editedCells, [key]: true };
            data._isEditedRow = true;
            changed = true;
          }
        }
      });
      
      if (changed) {
        updatedRows.push(data);
      }
    });

    if (updatedRows.length > 0) {
      gridApi.applyTransaction({ update: updatedRows });
      alert(`Trimmed text in ${updatedRows.length} rows.`);
    } else {
      alert('No extra spaces found.');
    }
  };

  const removeEmptyRowsColumns = () => {
    if (!gridApi) return;
    
    // Remove empty rows
    const rowNodes = [];
    gridApi.forEachNode(node => rowNodes.push(node));
    
    const rowsToRemove = [];
    rowNodes.forEach(node => {
      const data = node.data;
      const isEmpty = Object.keys(data).every(key => {
        if (key.startsWith('_')) return true; // Ignore internal keys
        return data[key] === null || data[key] === undefined || data[key] === '';
      });
      if (isEmpty) rowsToRemove.push(data);
    });

    if (rowsToRemove.length > 0) {
      gridApi.applyTransaction({ remove: rowsToRemove });
    }

    // We can also implement removing empty columns, but standard Excel 
    // files usually have fixed headers. For simplicity, we just remove empty rows.
    alert(`Removed ${rowsToRemove.length} empty rows.`);
  };

  const deleteSelectedRows = () => {
    if (!gridApi) return;
    const selectedRows = gridApi.getSelectedRows();
    if (selectedRows.length === 0) {
      alert('Please select rows to delete first.');
      return;
    }
    
    gridApi.applyTransaction({ remove: selectedRows });
  };

  const handleSaveAndDownload = async () => {
    if (!gridApi) {
      alert('Grid is not ready. Please try again.');
      return;
    }
    setLoading(true);

    const rows = [];
    gridApi.forEachNodeAfterFilterAndSort(node => {
      // Clean internal tracking variables before sending to backend
      const cleanData = { ...node.data };
      delete cleanData._editedCells;
      delete cleanData._isEditedRow;
      rows.push(cleanData);
    });

    try {
      const extStr = file ? file.name.split('.').pop().toLowerCase() : 'xlsx';
      const validExportTypes = ['xlsx', 'xls', 'xlsb', 'xlsm', 'csv', 'ods', 'txt'];
      const bookType = validExportTypes.includes(extStr) ? extStr : 'xlsx';
      const filename = file ? `edited_${file.name}` : `edited_data.${bookType}`;

      const response = await axios.post('http://localhost:3000/api/download', {
        rows,
        sheetName: sheetName || 'Sheet1',
        bookType,
        filename
      }, {
        responseType: 'blob'
      });

      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', filename);
      document.body.appendChild(link);
      link.click();
      link.parentNode.removeChild(link);
    } catch (error) {
      console.error('Download error:', error);
      alert('Failed to generate Excel file.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="app-container">
      <header className="header">
        <h1>Excel Web Editor</h1>
        <p>Upload, Edit, Clean, and Download Spreadsheets</p>
      </header>

      {loading && (
        <div className="loading-overlay">
          <div className="spinner"></div>
          <p>Processing...</p>
        </div>
      )}

      {!rowData.length ? (
        <div className="glass-panel" style={{ flexGrow: 1, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <label className="upload-zone">
            <input 
              type="file" 
              accept=".xlsx, .xls, .csv, .ods, .xlsb, .xlsm, .txt" 
              style={{ display: 'none' }} 
              onChange={handleFileUpload}
              ref={fileInputRef}
            />
            <FileSpreadsheet className="upload-icon" />
            <span className="upload-text">Click or Drag & Drop Spreadsheet File Here</span>
            <span style={{ marginTop: '10px', color: 'var(--text-muted)' }}>Supports .xlsx, .xls, .csv, .ods, and more</span>
          </label>
        </div>
      ) : (
        <>
          <div className="glass-panel">
            <div className="toolbar">
              <div className="toolbar-group">
                <div style={{ position: 'relative' }}>
                  <Search style={{ position: 'absolute', left: '10px', top: '10px', color: 'var(--text-muted)', width: '18px', height: '18px' }} />
                  <input
                    type="text"
                    className="search-input"
                    placeholder="Search in grid..."
                    style={{ paddingLeft: '35px' }}
                    value={searchText}
                    onChange={(e) => setSearchText(e.target.value)}
                  />
                </div>
              </div>

              <div className="toolbar-group">
                <button className="btn btn-secondary" onClick={trimText}>
                  <Scissors size={18} /> Trim Text
                </button>
                <button className="btn btn-secondary" onClick={removeEmptyRowsColumns}>
                  <Eraser size={18} /> Remove Empty Rows
                </button>
                <button className="btn btn-secondary" onClick={handleClearFile}>
                  <Trash2 size={18} /> Clear File
                </button>
                <button className="btn btn-danger" onClick={deleteSelectedRows}>
                  <Trash2 size={18} /> Delete Selected
                </button>
                <button className="btn btn-success" onClick={handleSaveAndDownload}>
                  <Download size={18} /> Save & Download
                </button>
              </div>
            </div>
          </div>

          <ExcelGrid 
            gridRef={gridRef}
            rowData={rowData}
            columnDefs={columnDefs}
            onGridReady={handleGridReady}
            onCellValueChanged={handleCellValueChanged}
            searchText={searchText}
          />
        </>
      )}
    </div>
  );
}

export default App;
