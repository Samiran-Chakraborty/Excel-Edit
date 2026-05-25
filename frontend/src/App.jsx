import React, { useState, useRef, useCallback, useEffect } from 'react';
import axios from 'axios';
import { Upload, Download, Scissors, Trash2, Search, FileSpreadsheet, Eraser, Undo2, Redo2, Type } from 'lucide-react';
import { Toaster, toast } from 'react-hot-toast';
import ExcelGrid from './components/ExcelGrid';

function App() {
  const [file, setFile] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [rowData, setRowData] = useState([]);
  const [columnDefs, setColumnDefs] = useState([]);
  const [sheetName, setSheetName] = useState('');
  const [searchText, setSearchText] = useState('');
  const [loading, setLoading] = useState(false);
  const [gridApi, setGridApi] = useState(null);
  const gridRef = useRef();
  const fileInputRef = useRef(null);

  const [isDragging, setIsDragging] = useState(false);
  const [history, setHistory] = useState([]);
  const [future, setFuture] = useState([]);

  // Add Text panel state
  const [showAddText, setShowAddText] = useState(false);
  const [addTextCol, setAddTextCol] = useState('');
  const [addTextMode, setAddTextMode] = useState('prefix'); // 'prefix' | 'suffix' | 'custom'
  const [addTextValue, setAddTextValue] = useState('');
  // Condition state (only apply where...)
  const [useCondition, setUseCondition] = useState(false);
  const [condCol, setCondCol] = useState('');
  const [condOp, setCondOp] = useState('equals'); // equals | contains | startswith | endswith
  const [condVal, setCondVal] = useState('');
  
  // Highlight Duplicates state
  const [showHighlightDupes, setShowHighlightDupes] = useState(false);
  const [dupeCol, setDupeCol] = useState('ALL');

  const saveHistoryBeforeAction = () => {
    if (!gridApi) return;
    const currentData = [];
    gridApi.forEachNode(node => {
      currentData.push(JSON.parse(JSON.stringify(node.data)));
    });
    setHistory(prev => [...prev, currentData].slice(-20));
    setFuture([]);
  };

  const handleUndo = useCallback(() => {
    if (!gridApi || history.length === 0) return;
    
    const currentData = [];
    gridApi.forEachNode(node => currentData.push(JSON.parse(JSON.stringify(node.data))));
    setFuture(prev => [currentData, ...prev].slice(-20));

    const previousState = history[history.length - 1];
    setHistory(prev => prev.slice(0, -1));

    setRowData(previousState);
  }, [gridApi, history]);

  const handleRedo = useCallback(() => {
    if (!gridApi || future.length === 0) return;

    const currentData = [];
    gridApi.forEachNode(node => currentData.push(JSON.parse(JSON.stringify(node.data))));
    setHistory(prev => [...prev, currentData].slice(-20));

    const nextState = future[0];
    setFuture(prev => prev.slice(1));

    setRowData(nextState);
  }, [gridApi, future]);

  useEffect(() => {
    const handleKeyDown = (e) => {
      if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
      
      if (e.ctrlKey && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        handleUndo();
      } else if (e.ctrlKey && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        handleRedo();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleUndo, handleRedo]);

  const handleGridReady = (params) => {
    setGridApi(params.api);
  };

  const processFile = async (selectedFile) => {
    if (!selectedFile) return;

    const validExtensions = ['.xlsx', '.xls', '.csv', '.ods', '.xlsb', '.xlsm', '.txt'];
    const extension = '.' + selectedFile.name.split('.').pop().toLowerCase();
    
    if (!validExtensions.includes(extension)) {
      toast.error('Please upload a valid spreadsheet file (.xlsx, .xls, .csv, .ods, etc).');
      return;
    }

    setFile(selectedFile);
    setLoading(true);

    const formData = new FormData();
    formData.append('file', selectedFile);

    try {
      const response = await axios.post('/api/upload', formData, {
        headers: { 'Content-Type': 'multipart/form-data' },
      });

      const fetchedSheets = response.data.sheets;
      
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

      const processedSheets = fetchedSheets.map(s => ({
        sheetName: s.sheetName,
        columnDefs: [
          rowNumberCol,
          ...s.headers.map(header => ({ field: header, headerName: header }))
        ],
        rowData: s.rows.map(r => ({ ...r, _editedCells: {} }))
      }));

      setSheets(processedSheets);
      setActiveSheetIndex(0);
      setColumnDefs(processedSheets[0].columnDefs);
      setRowData(processedSheets[0].rowData);
      setSheetName(processedSheets[0].sheetName);
    } catch (error) {
      console.error('Upload error:', error);
      toast.error('Failed to process the file.');
    } finally {
      setLoading(false);
    }
  };

  const handleFileUpload = (e) => {
    processFile(e.target.files[0]);
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  const handleClearFile = () => {
    setFile(null);
    setSheets([]);
    setActiveSheetIndex(0);
    setRowData([]);
    setColumnDefs([]);
    setSheetName('');
    setSearchText('');
    setHistory([]);
    setFuture([]);
    setGridApi(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const getCurrentGridData = () => {
    if (!gridApi) return [];
    gridApi.stopEditing(false);
    const rows = [];
    gridApi.forEachNodeAfterFilterAndSort(node => {
      rows.push(node.data);
    });
    return rows;
  };

  const handleTabSwitch = (index) => {
    if (index === activeSheetIndex || !gridApi) return;
    
    // Save current sheet data
    const currentData = getCurrentGridData();
    const updatedSheets = [...sheets];
    updatedSheets[activeSheetIndex] = {
      ...updatedSheets[activeSheetIndex],
      rowData: currentData
    };
    
    setSheets(updatedSheets);
    setActiveSheetIndex(index);
    
    // Load new sheet data
    const nextSheet = updatedSheets[index];
    setRowData(nextSheet.rowData);
    setColumnDefs(nextSheet.columnDefs);
    setSheetName(nextSheet.sheetName);
    
    setHistory([]);
    setFuture([]);
  };

  const applyHighlightDuplicates = () => {
    if (!gridApi) return;
    
    saveHistoryBeforeAction();
    const rowNodes = [];
    gridApi.forEachNode(node => rowNodes.push(node));
    
    const seen = new Set();
    const updatedRows = [];
    let dupeCount = 0;

    rowNodes.forEach(node => {
      const data = node.data;
      let key = '';
      if (dupeCol === 'ALL') {
        const cleanData = {};
        Object.keys(data).forEach(k => {
          if (!k.startsWith('_')) cleanData[k] = data[k];
        });
        key = JSON.stringify(cleanData);
      } else {
        key = data[dupeCol] !== undefined ? String(data[dupeCol]) : '';
      }

      if (seen.has(key)) {
        if (!data._isDuplicate) {
          data._isDuplicate = true;
          updatedRows.push(data);
          dupeCount++;
        }
      } else {
        seen.add(key);
        if (data._isDuplicate) {
          data._isDuplicate = false;
          updatedRows.push(data);
        }
      }
    });

    if (updatedRows.length > 0) {
      gridApi.applyTransaction({ update: updatedRows });
      gridApi.redrawRows();
    }
    
    if (dupeCount > 0) {
      toast.error(`Found ${dupeCount} duplicate rows.`);
    } else {
      toast.success('No duplicates found.');
    }
    setShowHighlightDupes(false);
  };

  const clearDuplicatesHighlight = () => {
    if (!gridApi) return;
    const rowNodes = [];
    gridApi.forEachNode(node => rowNodes.push(node));
    const updatedRows = [];
    rowNodes.forEach(node => {
      if (node.data._isDuplicate) {
        node.data._isDuplicate = false;
        updatedRows.push(node.data);
      }
    });
    if (updatedRows.length > 0) {
      gridApi.applyTransaction({ update: updatedRows });
      gridApi.redrawRows();
      toast.success('Cleared duplicate highlights.');
    }
    setShowHighlightDupes(false);
  };

  const handleCellValueChanged = useCallback((params) => {
    const { colDef, data, oldValue, newValue } = params;
    if (oldValue !== newValue) {
      // Save history BEFORE applying the change
      const previousData = [];
      params.api.forEachNode(node => {
        const rowDataCopy = JSON.parse(JSON.stringify(node.data));
        if (node.data === data) {
           rowDataCopy[colDef.field] = oldValue;
        }
        previousData.push(rowDataCopy);
      });
      setHistory(prev => [...prev, previousData].slice(-20));
      setFuture([]);

      data._editedCells = { ...data._editedCells, [colDef.field]: true };
      data._isEditedRow = true;
      params.api.applyTransaction({ update: [data] });
    }
  }, []);

  // Check if a row matches the condition
  const rowMatchesCondition = (data) => {
    if (!useCondition || !condCol || !condVal) return true;
    const cell = data[condCol] !== undefined ? String(data[condCol]) : '';
    const val  = condVal;
    switch (condOp) {
      case 'equals':     return cell === val;
      case 'contains':   return cell.includes(val);
      case 'startswith': return cell.startsWith(val);
      case 'endswith':   return cell.endsWith(val);
      default:           return true;
    }
  };

  const applyAddText = () => {
    if (!gridApi || !addTextCol || !addTextValue) return;

    saveHistoryBeforeAction();

    const rowNodes = [];
    gridApi.forEachNode(node => rowNodes.push(node));
    const updatedRows = [];

    rowNodes.forEach(node => {
      const data = node.data;

      // Skip rows that don't match the condition
      if (!rowMatchesCondition(data)) return;

      const original = data[addTextCol] !== undefined ? String(data[addTextCol]) : '';
      let newVal = original;

      if (addTextMode === 'prefix') {
        newVal = addTextValue + original;
      } else if (addTextMode === 'suffix') {
        newVal = original + addTextValue;
      } else if (addTextMode === 'custom') {
        newVal = addTextValue.replace(/\{\{value\}\}/g, original);
      }

      if (newVal !== original) {
        data[addTextCol] = newVal;
        data._editedCells = { ...data._editedCells, [addTextCol]: true };
        data._isEditedRow = true;
        updatedRows.push(data);
      }
    });

    if (updatedRows.length > 0) {
      gridApi.applyTransaction({ update: updatedRows });
      toast.success(`Updated ${updatedRows.length} cells in "${addTextCol}".`);
    } else {
      setHistory(prev => prev.slice(0, -1));
      toast('No changes made — check your condition or text.');
    }
    setShowAddText(false);
    setAddTextValue('');
  };

  // Live preview — pick first row that matches condition
  const getAddTextPreview = () => {
    if (!gridApi || !addTextCol || !addTextValue) return null;
    let sample = '';
    gridApi.forEachNode(node => {
      if (!sample && rowMatchesCondition(node.data) &&
          node.data[addTextCol] !== undefined && node.data[addTextCol] !== '') {
        sample = String(node.data[addTextCol]);
      }
    });
    if (!sample) return null;
    if (addTextMode === 'prefix') return addTextValue + sample;
    if (addTextMode === 'suffix') return sample + addTextValue;
    if (addTextMode === 'custom') return addTextValue.replace(/\{\{value\}\}/g, sample);
    return null;
  };

  const trimText = () => {
    if (!gridApi) return;
    
    saveHistoryBeforeAction();
    
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
      toast.success(`Trimmed text in ${updatedRows.length} rows.`);
    } else {
      // Remove the useless history state if nothing changed
      setHistory(prev => prev.slice(0, -1));
      toast('No extra spaces found.');
    }
  };

  const removeEmptyRowsColumns = () => {
    if (!gridApi) return;
    
    saveHistoryBeforeAction();
    
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
      toast.success(`Removed ${rowsToRemove.length} empty rows.`);
    } else {
      setHistory(prev => prev.slice(0, -1));
      toast(`No empty rows found.`);
    }
  };

  const deleteSelectedRows = () => {
    if (!gridApi) return;
    const selectedRows = gridApi.getSelectedRows();
    if (selectedRows.length === 0) {
      toast.error('Please select rows to delete first.');
      return;
    }
    
    saveHistoryBeforeAction();
    gridApi.applyTransaction({ remove: selectedRows });
  };

  const handleSaveAndDownload = async () => {
    if (!gridApi) {
      toast.error('Grid is not ready. Please try again.');
      return;
    }
    
    // Stop any active cell editing to ensure the latest changes are saved
    gridApi.stopEditing(false);
    
    setLoading(true);

    const rows = [];
    gridApi.forEachNodeAfterFilterAndSort(node => {
      // Clean internal tracking variables before sending to backend
      const cleanData = { ...node.data };
      delete cleanData._editedCells;
      delete cleanData._isEditedRow;
      rows.push(cleanData);
    });

    // Prepare all sheets
    const payloadSheets = sheets.map((s, idx) => {
      if (idx === activeSheetIndex) {
        return { sheetName: s.sheetName, rows: rows };
      }
      return {
        sheetName: s.sheetName,
        rows: s.rowData.map(r => {
          const cleanData = { ...r };
          delete cleanData._editedCells;
          delete cleanData._isEditedRow;
          return cleanData;
        })
      };
    });

    try {
      const extStr = file ? file.name.split('.').pop().toLowerCase() : 'xlsx';
      const validExportTypes = ['xlsx', 'xls', 'xlsb', 'xlsm', 'csv', 'ods', 'txt'];
      const bookType = validExportTypes.includes(extStr) ? extStr : 'xlsx';
      const filename = file ? `edited_${file.name}` : `edited_data.${bookType}`;

      const response = await axios.post('/api/download', {
        sheets: payloadSheets,
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
      toast.error('Failed to generate Excel file.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="app-container">
      <Toaster position="bottom-right" />
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
          <label 
            className={`upload-zone ${isDragging ? 'is-dragging' : ''}`}
            onDragOver={handleDragOver}
            onDragEnter={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
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
                <button 
                  className="btn btn-secondary" 
                  onClick={handleUndo}
                  disabled={history.length === 0}
                  title="Undo (Ctrl+Z)"
                  style={{ padding: '0.5rem', opacity: history.length === 0 ? 0.5 : 1, cursor: history.length === 0 ? 'not-allowed' : 'pointer' }}
                >
                  <Undo2 size={18} />
                </button>
                <button 
                  className="btn btn-secondary" 
                  onClick={handleRedo}
                  disabled={future.length === 0}
                  title="Redo (Ctrl+Y)"
                  style={{ padding: '0.5rem', opacity: future.length === 0 ? 0.5 : 1, cursor: future.length === 0 ? 'not-allowed' : 'pointer' }}
                >
                  <Redo2 size={18} />
                </button>
                <div style={{ position: 'relative', marginLeft: '0.5rem' }}>
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
                <button
                  className={`btn btn-secondary ${showAddText ? 'btn-active' : ''}`}
                  onClick={() => { setShowAddText(v => !v); setShowHighlightDupes(false); setAddTextValue(''); setAddTextCol(columnDefs.filter(c => c.field).map(c => c.field)[0] || ''); }}
                >
                  <Type size={18} /> Add Text
                </button>
                <button
                  className={`btn btn-secondary ${showHighlightDupes ? 'btn-active' : ''}`}
                  onClick={() => { setShowHighlightDupes(v => !v); setShowAddText(false); }}
                >
                  <Search size={18} /> Highlight Duplicates
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

          {/* Add Text Panel */}
          {showAddText && (
            <div className="add-text-panel">
              <div className="add-text-panel-header">
                <Type size={16} />
                <span>Add Text to Column</span>
              </div>
              <div className="add-text-panel-body">
                {/* Column selector */}
                <div className="add-text-field">
                  <label>Apply to Column</label>
                  <select
                    value={addTextCol}
                    onChange={e => setAddTextCol(e.target.value)}
                    className="add-text-select"
                  >
                    {columnDefs.filter(c => c.field).map(c => (
                      <option key={c.field} value={c.field}>{c.headerName || c.field}</option>
                    ))}
                  </select>
                </div>

                {/* Condition toggle */}
                <div className="add-text-field" style={{ justifyContent: 'flex-end' }}>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '0.4rem', cursor: 'pointer', textTransform: 'none', letterSpacing: 0, fontSize: '0.85rem' }}>
                    <input
                      type="checkbox"
                      checked={useCondition}
                      onChange={e => { setUseCondition(e.target.checked); setCondVal(''); setCondCol(columnDefs.filter(c => c.field)[0]?.field || ''); }}
                      style={{ accentColor: 'var(--excel-green)', width: 14, height: 14 }}
                    />
                    Only where…
                  </label>
                </div>

                {/* Condition row */}
                {useCondition && (
                  <>
                    <div className="add-text-field">
                      <label>If Column</label>
                      <select value={condCol} onChange={e => setCondCol(e.target.value)} className="add-text-select">
                        {columnDefs.filter(c => c.field).map(c => (
                          <option key={c.field} value={c.field}>{c.headerName || c.field}</option>
                        ))}
                      </select>
                    </div>
                    <div className="add-text-field">
                      <label>Operator</label>
                      <select value={condOp} onChange={e => setCondOp(e.target.value)} className="add-text-select" style={{ minWidth: 130 }}>
                        <option value="equals">= Equals</option>
                        <option value="contains">Contains</option>
                        <option value="startswith">Starts with</option>
                        <option value="endswith">Ends with</option>
                      </select>
                    </div>
                    <div className="add-text-field">
                      <label>Value</label>
                      <input
                        type="text"
                        className="add-text-input"
                        style={{ minWidth: 140 }}
                        placeholder="e.g.  T_1"
                        value={condVal}
                        onChange={e => setCondVal(e.target.value)}
                      />
                    </div>
                  </>
                )}

                {/* Mode tabs */}
                <div className="add-text-field">
                  <label>Position</label>
                  <div className="add-text-tabs">
                    {['prefix', 'suffix', 'custom'].map(mode => (
                      <button
                        key={mode}
                        className={`add-text-tab ${addTextMode === mode ? 'active' : ''}`}
                        onClick={() => { setAddTextMode(mode); setAddTextValue(''); }}
                      >
                        {mode === 'prefix' ? '← Prefix' : mode === 'suffix' ? 'Suffix →' : '✦ Custom'}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Text input */}
                <div className="add-text-field">
                  <label>
                    {addTextMode === 'prefix' && 'Text to add at the front'}
                    {addTextMode === 'suffix' && 'Text to add at the end'}
                    {addTextMode === 'custom' && <>Pattern — use <code>{'{{}}'}</code> for original value</>}
                  </label>
                  <input
                    type="text"
                    className="add-text-input"
                    placeholder={
                      addTextMode === 'prefix' ? 'e.g.  ID_' :
                      addTextMode === 'suffix' ? 'e.g.  _done' :
                      'e.g.  prefix_{{value}}_suffix'
                    }
                    value={addTextValue}
                    onChange={e => setAddTextValue(e.target.value)}
                  />
                </div>

                {/* Live preview */}
                {getAddTextPreview() && (
                  <div className="add-text-preview">
                    <span className="preview-label">Preview:</span>
                    <span className="preview-value">{getAddTextPreview()}</span>
                  </div>
                )}

                {/* Actions */}
                <div className="add-text-actions">
                  <button className="btn btn-secondary" onClick={() => setShowAddText(false)}>Cancel</button>
                  <button
                    className="btn btn-success"
                    disabled={!addTextCol || !addTextValue}
                    onClick={applyAddText}
                  >
                    Apply to All Rows
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Highlight Duplicates Panel */}
          {showHighlightDupes && (
            <div className="add-text-panel">
              <div className="add-text-panel-header" style={{ background: '#d13438' }}>
                <Search size={16} />
                <span>Highlight Duplicates</span>
              </div>
              <div className="add-text-panel-body">
                <div className="add-text-field">
                  <label>Column to check</label>
                  <select
                    value={dupeCol}
                    onChange={e => setDupeCol(e.target.value)}
                    className="add-text-select"
                  >
                    <option value="ALL">All Columns (Entire Row)</option>
                    {columnDefs.filter(c => c.field).map(c => (
                      <option key={c.field} value={c.field}>{c.headerName || c.field}</option>
                    ))}
                  </select>
                </div>
                <div className="add-text-actions">
                  <button className="btn btn-secondary" onClick={() => setShowHighlightDupes(false)}>Cancel</button>
                  <button className="btn btn-danger" onClick={applyHighlightDuplicates}>Find Duplicates</button>
                  <button className="btn btn-secondary" onClick={clearDuplicatesHighlight}>Clear</button>
                </div>
              </div>
            </div>
          )}
        </div>

          <ExcelGrid 
            gridRef={gridRef}
            rowData={rowData}
            columnDefs={columnDefs}
            onGridReady={handleGridReady}
            onCellValueChanged={handleCellValueChanged}
            searchText={searchText}
          />
          
          {/* Sheet Tabs */}
          {sheets.length > 1 && (
            <div className="sheet-tabs">
              {sheets.map((sheet, idx) => (
                <button
                  key={idx}
                  className={`sheet-tab ${idx === activeSheetIndex ? 'active' : ''}`}
                  onClick={() => handleTabSwitch(idx)}
                >
                  {sheet.sheetName}
                </button>
              ))}
            </div>
          )}
        </>
      )}
    </div>
  );
}

export default App;
