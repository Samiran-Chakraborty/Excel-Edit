import React, { useState, useEffect, useRef, useCallback, forwardRef, useImperativeHandle } from 'react';

const ExcelSetFilter = forwardRef((props, ref) => {
  const [uniqueValues, setUniqueValues] = useState([]);
  const uniqueValuesRef = useRef([]);
  const [selectedValues, setSelectedValues] = useState(new Set());
  const appliedValuesRef = useRef(new Set());
  const [searchTerm, setSearchTerm] = useState('');

  // Extract unique values
  const updateValues = useCallback(() => {
    if (!props.api) return;
    
    const values = new Set();
    props.api.forEachNode(node => {
      if (node.data && node.data[props.colDef.field] !== undefined) {
        values.add(node.data[props.colDef.field]);
      }
    });
    
    const sortedValues = Array.from(values).sort((a, b) => {
      if (a === b) return 0;
      if (a === null || a === '') return -1;
      if (b === null || b === '') return 1;
      return String(a).localeCompare(String(b));
    });

    const wasAllSelected = appliedValuesRef.current.size === uniqueValuesRef.current.length || uniqueValuesRef.current.length === 0;
    
    setUniqueValues(sortedValues);
    uniqueValuesRef.current = sortedValues;
    
    if (wasAllSelected) {
      setSelectedValues(new Set(sortedValues));
      appliedValuesRef.current = new Set(sortedValues);
    } else {
      const validSelected = new Set(Array.from(appliedValuesRef.current).filter(v => values.has(v)));
      setSelectedValues(validSelected);
      appliedValuesRef.current = validSelected;
    }
  }, [props.api, props.colDef.field]);

  useEffect(() => {
    updateValues();

    if (!props.api) return;

    const onDataChanged = () => {
      updateValues();
    };

    props.api.addEventListener('cellValueChanged', onDataChanged);
    props.api.addEventListener('rowDataUpdated', onDataChanged);

    return () => {
      props.api.removeEventListener('cellValueChanged', onDataChanged);
      props.api.removeEventListener('rowDataUpdated', onDataChanged);
    };
  }, [props.api, updateValues]);

  // Expose AG Grid filter lifecycle methods using standard useImperativeHandle
  useImperativeHandle(ref, () => {
    return {
      afterGuiAttached: () => {
        updateValues();
      },
      isFilterActive: () => {
        return appliedValuesRef.current.size !== uniqueValuesRef.current.length && uniqueValuesRef.current.length > 0;
      },
      doesFilterPass: (params) => {
        const val = params.data[props.colDef.field];
        return appliedValuesRef.current.has(val);
      },
      getModel: () => {
        if (appliedValuesRef.current.size === uniqueValuesRef.current.length || uniqueValuesRef.current.length === 0) return null;
        return { values: Array.from(appliedValuesRef.current) };
      },
      setModel: (model) => {
        if (model && model.values) {
          const newSet = new Set(model.values);
          appliedValuesRef.current = newSet;
          setSelectedValues(newSet);
        } else {
          appliedValuesRef.current = new Set(uniqueValuesRef.current);
          setSelectedValues(new Set(uniqueValuesRef.current));
        }
      }
    };
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

  const onOk = () => {
    appliedValuesRef.current = new Set(selectedValues);
    
    // Build the new model
    const model = appliedValuesRef.current.size === uniqueValuesRef.current.length || uniqueValuesRef.current.length === 0
      ? null
      : { values: Array.from(appliedValuesRef.current) };
      
    // Notify AG Grid of the filter change
    if (typeof props.onModelChange === "function") {
      props.onModelChange(model);
    } else if (typeof props.filterChangedCallback === "function") {
      props.filterChangedCallback();
    } else {
      props.api.onFilterChanged();
    }
    
    // Close the filter menu
    if (typeof props.hidePopup === 'function') {
      props.hidePopup();
    } else if (props.api && typeof props.api.hidePopupMenu === 'function') {
      props.api.hidePopupMenu();
    } else {
      // Dispatch Escape to close popups
      document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
    }
  };

  const onCancel = () => {
    setSelectedValues(new Set(appliedValuesRef.current)); // Revert UI to applied state
    if (typeof props.hidePopup === 'function') {
      props.hidePopup();
    } else if (props.api && typeof props.api.hidePopupMenu === 'function') {
      props.api.hidePopupMenu();
    } else {
      document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
    }
  };

  const filteredValues = uniqueValues.filter(v => 
    String(v).toLowerCase().includes(searchTerm)
  );

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
});

export default ExcelSetFilter;
