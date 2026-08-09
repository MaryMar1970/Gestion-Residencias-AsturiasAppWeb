import { useState, useRef, useCallback } from 'react';

export type ColumnDef = {
  key: string;
  label: string;
  defaultWidth: number;
  minWidth?: number;
  visible?: boolean;
  isFixed?: boolean;
};

export type ColumnState = {
  key: string;
  label: string;
  width: number;
  visible: boolean;
  isFixed?: boolean;
};

export type SavedConfig = {
  order: string[];
  widths: Record<string, number>;
  hidden: string[];
};

export function useColumnConfig(
  storageKey: string,
  initialColumns: ColumnDef[]
) {
  const [columns, setColumns] = useState<ColumnState[]>(() => {
    const savedJson = localStorage.getItem(storageKey);
    let saved: SavedConfig | null = null;
    if (savedJson) {
      try { saved = JSON.parse(savedJson); } catch { }
    }

    const initMap = new Map(initialColumns.map(c => [c.key, c]));

    let orderKeys: string[] = [];
    if (saved && Array.isArray(saved.order)) {
      orderKeys = saved.order.filter(k => initMap.has(k));
      initialColumns.forEach(c => {
        if (!orderKeys.includes(c.key)) orderKeys.push(c.key);
      });
    } else {
      orderKeys = initialColumns.map(c => c.key);
    }

    const nonFixed = orderKeys.filter(k => !initMap.get(k)?.isFixed);
    const fixed = orderKeys.filter(k => initMap.get(k)?.isFixed);
    const finalOrder = [...nonFixed, ...fixed];

    return finalOrder.map(key => {
      const def = initMap.get(key)!;
      const width = saved?.widths?.[key] ?? def.defaultWidth;
      const hidden = saved?.hidden?.includes(key);
      const visible = hidden !== undefined ? !hidden : (def.visible !== false);
      return {
        key: def.key,
        label: def.label,
        width,
        visible,
        isFixed: def.isFixed,
      };
    });
  });

  const [draggedKey, setDraggedKey] = useState<string | null>(null);
  const [dragOverKey, setDragOverKey] = useState<string | null>(null);
  const [showPicker, setShowPicker] = useState(false);

  const resizingRef = useRef<{ key: string; startX: number; startWidth: number } | null>(null);

  const saveConfig = useCallback((cols: ColumnState[]) => {
    const order = cols.map(c => c.key);
    const widths: Record<string, number> = {};
    const hidden: string[] = [];

    cols.forEach(c => {
      widths[c.key] = c.width;
      if (!c.visible) hidden.push(c.key);
    });

    const config: SavedConfig = { order, widths, hidden };
    localStorage.setItem(storageKey, JSON.stringify(config));
  }, [storageKey]);

  const updateColumns = (updater: (prev: ColumnState[]) => ColumnState[]) => {
    setColumns(prev => {
      const next = updater(prev);
      saveConfig(next);
      return next;
    });
  };

  const toggleVisibility = (key: string) => {
    updateColumns(prev =>
      prev.map(c => c.key === key ? { ...c, visible: !c.visible } : c)
    );
  };

  const handleDragStart = (key: string, e: React.DragEvent) => {
    const col = columns.find(c => c.key === key);
    if (col?.isFixed) return;
    setDraggedKey(key);
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', key);
  };

  const handleDragOver = (key: string, e: React.DragEvent) => {
    const col = columns.find(c => c.key === key);
    if (!draggedKey || draggedKey === key || col?.isFixed) return;
    e.preventDefault();
    setDragOverKey(key);
  };

  const handleDrop = (targetKey: string, e: React.DragEvent) => {
    e.preventDefault();
    if (!draggedKey || draggedKey === targetKey) {
      setDraggedKey(null);
      setDragOverKey(null);
      return;
    }

    updateColumns(prev => {
      const fromIdx = prev.findIndex(c => c.key === draggedKey);
      const toIdx = prev.findIndex(c => c.key === targetKey);
      if (fromIdx === -1 || toIdx === -1) return prev;

      const next = [...prev];
      const [moved] = next.splice(fromIdx, 1);
      next.splice(toIdx, 0, moved);
      return next;
    });

    setDraggedKey(null);
    setDragOverKey(null);
  };

  const handleResizeStart = (key: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const col = columns.find(c => c.key === key);
    if (!col) return;

    resizingRef.current = {
      key,
      startX: e.clientX,
      startWidth: col.width,
    };

    const handleMouseMove = (moveEv: MouseEvent) => {
      if (!resizingRef.current) return;
      const { key: targetKey, startX, startWidth } = resizingRef.current;
      const diff = moveEv.clientX - startX;
      const def = initialColumns.find(c => c.key === targetKey);
      const minW = def?.minWidth ?? 35;
      const newWidth = Math.max(minW, startWidth + diff);

      setColumns(prev =>
        prev.map(c => c.key === targetKey ? { ...c, width: newWidth } : c)
      );
    };

    const handleMouseUp = () => {
      if (resizingRef.current) {
        resizingRef.current = null;
        setColumns(latest => {
          saveConfig(latest);
          return latest;
        });
      }
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
  };

  const resetDefaults = () => {
    localStorage.removeItem(storageKey);
    const defaults = initialColumns.map(c => ({
      key: c.key,
      label: c.label,
      width: c.defaultWidth,
      visible: c.visible !== false,
      isFixed: c.isFixed,
    }));
    setColumns(defaults);
  };

  const visibleColumns = columns.filter(c => c.visible);

  return {
    columns,
    visibleColumns,
    draggedKey,
    dragOverKey,
    showPicker,
    setShowPicker,
    toggleVisibility,
    handleDragStart,
    handleDragOver,
    handleDrop,
    handleResizeStart,
    resetDefaults,
  };
}
