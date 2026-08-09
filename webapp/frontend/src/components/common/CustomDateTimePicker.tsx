import React, { useState, useEffect, useRef } from 'react';

interface CustomDateTimePickerProps {
  value: string; // ISO string or YYYY-MM-DDTHH:mm
  onChange: (value: string) => void;
  required?: boolean;
}

export const CustomDateTimePicker: React.FC<CustomDateTimePickerProps> = ({
  value,
  onChange
}) => {
  const [isOpen, setIsOpen] = useState(false);

  // Parse initial date (with localStorage fallback to last selected date/time)
  const parseValueDate = (valStr: string) => {
    let target = valStr;
    if (!target) {
      target = localStorage.getItem('lastSelectedFechaSolicitud') || '';
    }
    if (!target) return new Date();
    const d = new Date(target);
    return isNaN(d.getTime()) ? new Date() : d;
  };

  const initialD = parseValueDate(value);

  // View state
  const [viewDate, setViewDate] = useState(new Date(initialD.getFullYear(), initialD.getMonth(), 1));
  const [selYear, setSelYear] = useState(initialD.getFullYear());
  const [selMonth, setSelMonth] = useState(initialD.getMonth());
  const [selDay, setSelDay] = useState(initialD.getDate());
  const [selHour, setSelHour] = useState(initialD.getHours());
  const [selMin, setSelMin] = useState(initialD.getMinutes());

  // User explicit selection tracking
  const [hasSelDate, setHasSelDate] = useState(false);
  const [hasSelHour, setHasSelHour] = useState(false);
  const [hasSelMin, setHasSelMin] = useState(false);

  const popoverRef = useRef<HTMLDivElement>(null);
  const hourListRef = useRef<HTMLDivElement>(null);
  const minListRef = useRef<HTMLDivElement>(null);

  // Format display text for input
  const formatDisplay = (valStr: string) => {
    if (!valStr) return '';
    const d = parseValueDate(valStr);
    const pad = (n: number) => n.toString().padStart(2, '0');
    return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  };

  const [inputText, setInputText] = useState(() => formatDisplay(value));

  useEffect(() => {
    setInputText(formatDisplay(value));
  }, [value]);

  // Open popover & reset selection flags
  const handleOpen = () => {
    const d = parseValueDate(value);
    setViewDate(new Date(d.getFullYear(), d.getMonth(), 1));
    setSelYear(d.getFullYear());
    setSelMonth(d.getMonth());
    setSelDay(d.getDate());
    setSelHour(d.getHours());
    setSelMin(d.getMinutes());

    setHasSelDate(false);
    setHasSelHour(false);
    setHasSelMin(false);

    setIsOpen(true);
  };

  // Close popup and save value
  const commitAndClose = (y: number, m: number, d: number, h: number, min: number) => {
    const pad = (n: number) => n.toString().padStart(2, '0');
    const formatted = `${y}-${pad(m + 1)}-${pad(d)}T${pad(h)}:${pad(min)}`;
    localStorage.setItem('lastSelectedFechaSolicitud', formatted);
    onChange(formatted);
    setIsOpen(false);
  };

  // Masked manual input helpers
  const formatFromDigits = (digits: string) => {
    const d = digits.slice(0, 12);
    let res = '';
    if (d.length > 0) res += d.slice(0, 2);
    if (d.length >= 2) res += '/';
    if (d.length > 2) res += d.slice(2, 4);
    if (d.length >= 4) res += '/';
    if (d.length > 4) res += d.slice(4, 8);
    if (d.length >= 8) res += ' ';
    if (d.length > 8) res += d.slice(8, 10);
    if (d.length >= 10) res += ':';
    if (d.length > 10) res += d.slice(10, 12);
    return res;
  };

  const parseToIso = (formattedStr: string) => {
    const match = formattedStr.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})$/);
    if (!match) return null;
    const [, dd, mm, yyyy, hh, min] = match;
    const d = parseInt(dd, 10);
    const m = parseInt(mm, 10) - 1;
    const y = parseInt(yyyy, 10);
    const h = parseInt(hh, 10);
    const mi = parseInt(min, 10);
    const dateObj = new Date(y, m, d, h, mi);
    if (
      dateObj.getFullYear() === y &&
      dateObj.getMonth() === m &&
      dateObj.getDate() === d &&
      h >= 0 && h < 24 &&
      mi >= 0 && mi < 60
    ) {
      const pad = (n: number) => n.toString().padStart(2, '0');
      return `${yyyy}-${pad(m + 1)}-${pad(d)}T${pad(h)}:${pad(mi)}`;
    }
    return null;
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const rawValue = e.target.value;
    const digitsOnly = rawValue.replace(/\D/g, '');
    const formatted = formatFromDigits(digitsOnly);
    setInputText(formatted);

    if (digitsOnly.length === 12) {
      const iso = parseToIso(formatted);
      if (iso) {
        localStorage.setItem('lastSelectedFechaSolicitud', iso);
        onChange(iso);
      }
    }
  };

  const handleInputBlur = () => {
    const iso = parseToIso(inputText);
    if (iso) {
      localStorage.setItem('lastSelectedFechaSolicitud', iso);
      onChange(iso);
    } else {
      setInputText(formatDisplay(value));
    }
  };

  // Click outside listener
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (popoverRef.current && !popoverRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [isOpen]);

  // Scroll active hour/min into view on open
  useEffect(() => {
    if (isOpen) {
      setTimeout(() => {
        const hourEl = hourListRef.current?.querySelector('[data-selected="true"]');
        if (hourEl) hourEl.scrollIntoView({ block: 'center' });

        const minEl = minListRef.current?.querySelector('[data-selected="true"]');
        if (minEl) minEl.scrollIntoView({ block: 'center' });
      }, 50);
    }
  }, [isOpen]);

  // Selection handlers
  const onSelectDate = (d: number) => {
    const nextY = viewDate.getFullYear();
    const nextM = viewDate.getMonth();
    setSelYear(nextY);
    setSelMonth(nextM);
    setSelDay(d);
    setHasSelDate(true);

    if (hasSelHour && hasSelMin) {
      commitAndClose(nextY, nextM, d, selHour, selMin);
    }
  };

  const onSelectHour = (h: number) => {
    setSelHour(h);
    setHasSelHour(true);

    if (hasSelDate && hasSelMin) {
      commitAndClose(selYear, selMonth, selDay, h, selMin);
    }
  };

  const onSelectMin = (m: number) => {
    setSelMin(m);
    setHasSelMin(true);

    if (hasSelDate && hasSelHour) {
      commitAndClose(selYear, selMonth, selDay, selHour, m);
    }
  };

  // Calendar helpers
  const daysInMonth = new Date(viewDate.getFullYear(), viewDate.getMonth() + 1, 0).getDate();
  const firstDayIndex = (new Date(viewDate.getFullYear(), viewDate.getMonth(), 1).getDay() + 6) % 7; // Monday = 0

  const monthNames = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];

  const prevMonth = () => {
    setViewDate(new Date(viewDate.getFullYear(), viewDate.getMonth() - 1, 1));
  };
  const nextMonth = () => {
    setViewDate(new Date(viewDate.getFullYear(), viewDate.getMonth() + 1, 1));
  };

  return (
    <div style={{ position: 'relative', width: '100%' }} ref={popoverRef}>
      {/* Dual Zone Input Control */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          background: 'var(--surface, #ffffff)',
          border: '1px solid var(--border, #cbd5e1)',
          borderRadius: 6,
          fontSize: 12,
          height: 28,
          boxSizing: 'border-box',
          overflow: 'hidden'
        }}
      >
        {/* Left/Center: Text input for manual typing */}
        <input
          type="text"
          value={inputText}
          onChange={handleInputChange}
          onBlur={handleInputBlur}
          onFocus={() => setIsOpen(false)}
          placeholder="DD/MM/YYYY HH:mm"
          maxLength={16}
          style={{
            flex: 1,
            border: 'none',
            outline: 'none',
            background: 'transparent',
            padding: '4px 8px',
            fontSize: 12,
            fontWeight: 600,
            color: 'var(--text, #0f172a)',
            fontFamily: 'inherit',
            minWidth: 0
          }}
        />

        {/* Right Zone: Button to open mouse calendar picker */}
        <button
          type="button"
          onClick={() => {
            if (isOpen) {
              setIsOpen(false);
            } else {
              handleOpen();
            }
          }}
          title="Abrir selector de fecha/hora con el ratón"
          style={{
            border: 'none',
            borderLeft: '1px solid var(--border, #cbd5e1)',
            background: 'var(--surface-2, #f8fafc)',
            padding: '0 8px',
            height: '100%',
            display: 'flex',
            alignItems: 'center',
            gap: 4,
            cursor: 'pointer',
            color: 'var(--primary, #1e3a8a)',
            fontSize: 12
          }}
        >
          <span>📅</span>
          <span style={{ fontSize: 9, color: 'var(--text-muted)' }}>▼</span>
        </button>
      </div>

      {/* Popover Dropdown */}
      {isOpen && (
        <div
          style={{
            position: 'absolute',
            top: 'calc(100% + 4px)',
            left: 0,
            zIndex: 3000,
            background: '#ffffff',
            border: '1px solid #cbd5e1',
            borderRadius: 10,
            boxShadow: '0 12px 32px rgba(0,0,0,0.2)',
            padding: 12,
            width: 440,
            display: 'flex',
            flexDirection: 'column',
            gap: 10,
            color: '#0f172a',
            fontFamily: 'inherit'
          }}
        >
          {/* Header */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #e2e8f0', paddingBottom: 6 }}>
            <span style={{ fontSize: 12, fontWeight: 700, color: '#1e3a8a' }}>
              Seleccionar Fecha y Hora
            </span>
          </div>

          {/* Main Grid: Calendar (Left) | Hours (Middle) | Minutes (Right) */}
          <div style={{ display: 'grid', gridTemplateColumns: '230px 85px 85px', gap: 10 }}>
            {/* Left: Month Calendar */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              {/* Month / Year Nav */}
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <button
                  type="button"
                  onClick={prevMonth}
                  style={{ border: 'none', background: 'transparent', cursor: 'pointer', fontWeight: 700, padding: '2px 6px', fontSize: 13 }}
                >
                  ◀
                </button>
                <span style={{ fontSize: 12, fontWeight: 700, color: '#1e293b' }}>
                  {monthNames[viewDate.getMonth()]} {viewDate.getFullYear()}
                </span>
                <button
                  type="button"
                  onClick={nextMonth}
                  style={{ border: 'none', background: 'transparent', cursor: 'pointer', fontWeight: 700, padding: '2px 6px', fontSize: 13 }}
                >
                  ▶
                </button>
              </div>

              {/* Day Headers */}
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', textAlign: 'center', fontSize: 10.5, fontWeight: 700, color: '#64748b' }}>
                <span>L</span><span>M</span><span>X</span><span>J</span><span>V</span><span>S</span><span>D</span>
              </div>

              {/* Day Cells */}
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 2, textAlign: 'center' }}>
                {Array.from({ length: firstDayIndex }).map((_, idx) => (
                  <div key={`empty-${idx}`} />
                ))}
                {Array.from({ length: daysInMonth }).map((_, idx) => {
                  const dayNum = idx + 1;
                  const isMatchDay = viewDate.getFullYear() === selYear && viewDate.getMonth() === selMonth && dayNum === selDay;
                  const isBlueHighlight = isMatchDay && hasSelDate;
                  const isDiscreetDefault = isMatchDay && !hasSelDate;

                  return (
                    <div
                      key={dayNum}
                      onClick={() => onSelectDate(dayNum)}
                      style={{
                        padding: '4px 0',
                        fontSize: 11,
                        borderRadius: 4,
                        cursor: 'pointer',
                        fontWeight: isMatchDay ? 700 : 500,
                        background: isBlueHighlight ? '#2563eb' : isDiscreetDefault ? '#e2e8f0' : 'transparent',
                        color: isBlueHighlight ? '#ffffff' : isDiscreetDefault ? '#1e293b' : '#334155',
                        border: isDiscreetDefault ? '1px dashed #94a3b8' : '1px solid transparent',
                        transition: 'all 0.15s ease'
                      }}
                    >
                      {dayNum}
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Middle: Hours Column */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              <div style={{ fontSize: 10.5, fontWeight: 700, color: '#64748b', textAlign: 'center' }}>Hora</div>
              <div
                ref={hourListRef}
                style={{
                  maxHeight: 180,
                  overflowY: 'auto',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: 3,
                  paddingRight: 4
                }}
              >
                {Array.from({ length: 24 }).map((_, h) => {
                  const isMatch = h === selHour;
                  const isBlue = isMatch && hasSelHour;
                  const isDiscreet = isMatch && !hasSelHour;

                  return (
                    <div
                      key={h}
                      data-selected={isMatch}
                      onClick={() => onSelectHour(h)}
                      style={{
                        padding: '3px 0',
                        textAlign: 'center',
                        fontSize: 11.5,
                        borderRadius: 4,
                        cursor: 'pointer',
                        fontWeight: isMatch ? 700 : 500,
                        background: isBlue ? '#2563eb' : isDiscreet ? '#e2e8f0' : '#f8fafc',
                        color: isBlue ? '#ffffff' : isDiscreet ? '#1e293b' : '#334155',
                        border: isDiscreet ? '1px dashed #94a3b8' : '1px solid #e2e8f0'
                      }}
                    >
                      {h.toString().padStart(2, '0')}
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Right: Minutes Column */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              <div style={{ fontSize: 10.5, fontWeight: 700, color: '#64748b', textAlign: 'center' }}>Minuto</div>
              <div
                ref={minListRef}
                style={{
                  maxHeight: 180,
                  overflowY: 'auto',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: 3,
                  paddingRight: 4
                }}
              >
                {Array.from({ length: 60 }).map((_, m) => {
                  const isMatch = m === selMin;
                  const isBlue = isMatch && hasSelMin;
                  const isDiscreet = isMatch && !hasSelMin;

                  return (
                    <div
                      key={m}
                      data-selected={isMatch}
                      onClick={() => onSelectMin(m)}
                      style={{
                        padding: '3px 0',
                        textAlign: 'center',
                        fontSize: 11.5,
                        borderRadius: 4,
                        cursor: 'pointer',
                        fontWeight: isMatch ? 700 : 500,
                        background: isBlue ? '#2563eb' : isDiscreet ? '#e2e8f0' : '#f8fafc',
                        color: isBlue ? '#ffffff' : isDiscreet ? '#1e293b' : '#334155',
                        border: isDiscreet ? '1px dashed #94a3b8' : '1px solid #e2e8f0'
                      }}
                    >
                      :{m.toString().padStart(2, '0')}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
