import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { useGridFilter } from 'ag-grid-react';

export default function ExcelSetFilter(props) {
  const { model, onModelChange, getValue, colDef, api } = props;
  
  const [uniqueValues, setUniqueValues] = useState([]);
  const [selectedValues, setSelectedValues] = useState(new Set());
  const [searchTerm, setSearchTerm] = useState('');

  // Extract unique values from the grid data
  const updateValues = useCallback(() => {
    if (!api) return;
    
    const values = new Set();
    api.forEachNode(node => {
      if (node.data) {
        const val = getValue ? getValue(node) : node.data[colDef.field];
        if (val !== undefined) {
          values.add(val);
        }
      }
    });
    
    const sortedValues = Array.from(values).sort((a, b) => {
      if (a === b) return 0;
      if (a === null || a === '') return -1;
      if (b === null || b === '') return 1;
      return String(a).localeCompare(String(b));
    });

    setUniqueValues(sortedValues);
  }, [api, colDef.field, getValue]);

  // Listen to grid changes to update unique values
  useEffect(() => {
    updateValues();

    if (!api) return;

    const onDataChanged = () => {
      updateValues();
    };

    api.addEventListener('cellValueChanged', onDataChanged);
    api.addEventListener('rowDataUpdated', onDataChanged);

    return () => {
      api.removeEventListener('cellValueChanged', onDataChanged);
      api.removeEventListener('rowDataUpdated', onDataChanged);
    };
  }, [api, updateValues]);

  // Sync selected values in UI with the active model
  useEffect(() => {
    if (model && model.values) {
      setSelectedValues(new Set(model.values));
    } else {
      setSelectedValues(new Set(uniqueValues));
    }
  }, [model, uniqueValues]);

  // Define doesFilterPass logic for AG Grid
  const doesFilterPass = useCallback((params) => {
    const { node } = params;
    const val = getValue ? getValue(node) : (node.data ? node.data[colDef.field] : undefined);
    
    if (!model || !model.values) return true;
    
    const selectedSet = new Set(model.values);
    return selectedSet.has(val);
  }, [model, colDef.field, getValue]);

  // Register with AG Grid using the useGridFilter hook
  useGridFilter({
    doesFilterPass
  });

  const onSearchChange = (e) => {
    setSearchTerm(e.target.value.toLowerCase());
  };

  const toggleSelectAll = (e) => {
    if (e.target.checked) {
      setSelectedValues(new Set(uniqueValues));
    } else {
      setSelectedValues(new Set());
    }
  };

  const toggleValue = (val) => {
    const newSelected = new Set(selectedValues);
    if (newSelected.has(val)) {
      newSelected.delete(val);
    } else {
      newSelected.add(val);
    }
    setSelectedValues(newSelected);
  };

  const closePopup = () => {
    if (typeof props.hidePopup === 'function') {
      props.hidePopup();
    } else if (api && typeof api.hidePopupMenu === 'function') {
      api.hidePopupMenu();
    } else {
      document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
    }
  };

  const onOk = () => {
    const allSelected = selectedValues.size === uniqueValues.length;
    const newModel = allSelected || uniqueValues.length === 0
      ? null
      : { values: Array.from(selectedValues) };
      
    onModelChange(newModel);
    closePopup();
  };

  const onCancel = () => {
    // Revert UI to the current model state
    if (model && model.values) {
      setSelectedValues(new Set(model.values));
    } else {
      setSelectedValues(new Set(uniqueValues));
    }
    closePopup();
  };

  const filteredValues = useMemo(() => {
    return uniqueValues.filter(v => 
      String(v).toLowerCase().includes(searchTerm)
    );
  }, [uniqueValues, searchTerm]);

  const isAllSelected = selectedValues.size === uniqueValues.length;
  const isIndeterminate = selectedValues.size > 0 && selectedValues.size < uniqueValues.length;

  return (
    <div style={{ 
      padding: '12px', 
      minWidth: '240px', 
      background: '#ffffff', 
      display: 'flex', 
      flexDirection: 'column', 
      gap: '10px',
      color: '#201f1e' 
    }}>
      
      {/* Search Input */}
      <div style={{ position: 'relative' }}>
        <input 
          type="text" 
          placeholder="Search" 
          value={searchTerm}
          onChange={onSearchChange}
          style={{
            width: '100%',
            padding: '6px 8px 6px 28px',
            border: '1px solid #8a8886',
            borderRadius: '2px',
            outline: 'none',
            fontSize: '13px',
            boxSizing: 'border-box'
          }}
        />
        <span style={{ position: 'absolute', left: '8px', top: '7px', fontSize: '12px', color: '#605e5c' }}>🔍</span>
      </div>

      {/* Select All Checkbox */}
      <div style={{ paddingBottom: '8px', borderBottom: '1px solid #e1dfdd' }}>
        <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '13px', fontWeight: '500' }}>
          <input 
            type="checkbox" 
            checked={isAllSelected}
            ref={input => {
              if (input) input.indeterminate = isIndeterminate;
            }}
            onChange={toggleSelectAll} 
            style={{ 
              cursor: 'pointer', 
              width: '16px', 
              height: '16px',
              accentColor: '#217346'
            }}
          />
          (Select All)
        </label>
      </div>

      {/* Values List */}
      <div style={{ maxHeight: '250px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '6px', paddingRight: '4px' }}>
        {filteredValues.map((val, i) => (
          <label key={i} style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '13px' }}>
            <input 
              type="checkbox" 
              checked={selectedValues.has(val)} 
              onChange={() => toggleValue(val)} 
              style={{ 
                cursor: 'pointer', 
                width: '16px', 
                height: '16px',
                accentColor: '#217346'
              }}
            />
            {val === '' || val === null ? '(Blanks)' : val}
          </label>
        ))}
        {filteredValues.length === 0 && (
          <div style={{ padding: '12px', color: '#605e5c', fontStyle: 'italic', fontSize: '13px', textAlign: 'center' }}>
            No matches found
          </div>
        )}
      </div>

      {/* Footer Buttons */}
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '4px', paddingTop: '12px', borderTop: '1px solid #e1dfdd' }}>
        <button 
          onClick={onCancel} 
          style={{ padding: '6px 16px', background: 'white', color: '#201f1e', border: '1px solid #8a8886', borderRadius: '2px', cursor: 'pointer', fontSize: '12px', fontWeight: '500' }}>
          Cancel
        </button>
        <button 
          onClick={onOk} 
          style={{ padding: '6px 20px', background: '#217346', color: 'white', border: '1px solid #217346', borderRadius: '2px', cursor: 'pointer', fontSize: '12px', fontWeight: '500' }}>
          OK
        </button>
      </div>

    </div>
  );
}
