import { AlertTriangle, CheckCircle2, Info, X, XCircle } from 'lucide-react';
import { useEffect, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { create } from 'zustand';
import { uid } from '@/lib/utils';

type Kind = 'info' | 'success' | 'error' | 'warning' | 'loading';

interface ToastItem {
  id: string;
  kind: Kind;
  message: ReactNode;
  detail?: ReactNode;
  action?: { label: string; onClick: () => void };
  duration: number;
  leaving?: boolean;
}

const useToasts = create<{ items: ToastItem[] }>(() => ({ items: [] }));

function push(kind: Kind, message: ReactNode, opts: Partial<Omit<ToastItem, 'id' | 'kind' | 'message'>> = {}): string {
  const id = uid('t');
  const duration = opts.duration ?? (kind === 'error' ? 7000 : kind === 'loading' ? 0 : 3600);
  useToasts.setState((s) => ({ items: [...s.items.slice(-4), { id, kind, message, duration, ...opts }] }));
  return id;
}

function dismiss(id: string) {
  useToasts.setState((s) => ({ items: s.items.map((t) => (t.id === id ? { ...t, leaving: true } : t)) }));
  setTimeout(() => useToasts.setState((s) => ({ items: s.items.filter((t) => t.id !== id) })), 220);
}

export const toast = {
  info: (m: ReactNode, o?: Partial<ToastItem>) => push('info', m, o),
  success: (m: ReactNode, o?: Partial<ToastItem>) => push('success', m, o),
  error: (m: ReactNode, o?: Partial<ToastItem>) => push('error', m, o),
  warning: (m: ReactNode, o?: Partial<ToastItem>) => push('warning', m, o),
  loading: (m: ReactNode, o?: Partial<ToastItem>) => push('loading', m, o),
  dismiss,
  update(id: string, kind: Kind, message: ReactNode, opts: Partial<ToastItem> = {}) {
    const duration = opts.duration ?? (kind === 'error' ? 7000 : 3200);
    useToasts.setState((s) => ({ items: s.items.map((t) => (t.id === id ? { ...t, kind, message, duration, ...opts } : t)) }));
  },
};

const ICONS: Record<Kind, ReactNode> = {
  info: <Info size={18} />,
  success: <CheckCircle2 size={18} />,
  error: <XCircle size={18} />,
  warning: <AlertTriangle size={18} />,
  loading: <span className="spinner" />,
};

function ToastView({ t }: { t: ToastItem }) {
  useEffect(() => {
    if (!t.duration) return;
    const timer = setTimeout(() => dismiss(t.id), t.duration);
    return () => clearTimeout(timer);
  }, [t.id, t.duration, t.kind]);
  return (
    <div className={`toast toast-${t.kind}${t.leaving ? ' leaving' : ''}`} role="status">
      <span className="toast-icon">{ICONS[t.kind]}</span>
      <div className="toast-body">
        <div className="toast-msg">{t.message}</div>
        {t.detail && <div className="toast-detail">{t.detail}</div>}
      </div>
      {t.action && (
        <button
          className="btn btn-sm toast-action"
          onClick={() => {
            t.action!.onClick();
            dismiss(t.id);
          }}
        >
          {t.action.label}
        </button>
      )}
      <button className="icon-btn toast-x" aria-label="Dismiss" onClick={() => dismiss(t.id)}>
        <X size={14} />
      </button>
    </div>
  );
}

export function ToastHost() {
  const items = useToasts((s) => s.items);
  return createPortal(
    <div className="toast-stack">
      {items.map((t) => (
        <ToastView key={t.id} t={t} />
      ))}
    </div>,
    document.body,
  );
}
