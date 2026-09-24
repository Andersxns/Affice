import { X } from 'lucide-react';
import { useEffect, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { create } from 'zustand';
import { uid } from '@/lib/utils';

export interface ModalProps {
  title?: ReactNode;
  icon?: ReactNode;
  onClose: () => void;
  children: ReactNode;
  footer?: ReactNode;
  width?: number | string;
  className?: string;
  /** Close when clicking the dimmed backdrop. */
  dismissable?: boolean;
  bodyClassName?: string;
}

export function Modal({ title, icon, onClose, children, footer, width = 460, className = '', dismissable = true, bodyClassName = '' }: ModalProps) {
  const ref = useRef<HTMLDivElement>(null);
  const [closing, setClosing] = useState(false);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;

  const requestClose = () => {
    setClosing(true);
    setTimeout(() => closeRef.current(), 130);
  };

  useEffect(() => {
    const prev = document.activeElement as HTMLElement | null;
    const el = ref.current;
    const first = el?.querySelector<HTMLElement>('[data-autofocus], input:not([type=hidden]), textarea, select, button.btn-primary');
    (first ?? el)?.focus({ preventScroll: true });
    if (first instanceof HTMLInputElement) first.select();
    return () => prev?.focus?.({ preventScroll: true });
  }, []);

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.stopPropagation();
      requestClose();
    }
    if (e.key === 'Tab' && ref.current) {
      // focus trap
      const f = Array.from(ref.current.querySelectorAll<HTMLElement>('button, input, select, textarea, [tabindex="0"], a[href]')).filter(
        (n) => !n.hasAttribute('disabled') && n.offsetParent !== null,
      );
      if (!f.length) return;
      const i = f.indexOf(document.activeElement as HTMLElement);
      if (e.shiftKey && (i <= 0)) {
        e.preventDefault();
        f[f.length - 1].focus();
      } else if (!e.shiftKey && i === f.length - 1) {
        e.preventDefault();
        f[0].focus();
      }
    }
  };

  return createPortal(
    <div
      className={`modal-backdrop${closing ? ' closing' : ''}`}
      onMouseDown={(e) => {
        if (dismissable && e.target === e.currentTarget) requestClose();
      }}
    >
      <div
        ref={ref}
        role="dialog"
        aria-modal="true"
        tabIndex={-1}
        className={`modal${closing ? ' closing' : ''} ${className}`}
        style={{ width }}
        onKeyDown={onKeyDown}
      >
        {title !== undefined && (
          <div className="modal-head">
            {icon && <span className="modal-icon">{icon}</span>}
            <h2 className="modal-title">{title}</h2>
            <button className="icon-btn modal-x" aria-label="Close" onClick={requestClose}>
              <X size={16} />
            </button>
          </div>
        )}
        <div className={`modal-body ${bodyClassName}`}>{children}</div>
        {footer && <div className="modal-foot">{footer}</div>}
      </div>
    </div>,
    document.body,
  );
}

/* ------------------------------------------------------------ dialog host */

interface HostedDialog {
  id: string;
  render: (close: (value?: unknown) => void) => ReactNode;
  resolve: (v: unknown) => void;
}

const useDialogStore = create<{ dialogs: HostedDialog[] }>(() => ({ dialogs: [] }));

/** Opens any modal UI imperatively; resolves with the value passed to close(). */
export function openDialog<T = unknown>(render: (close: (value?: T) => void) => ReactNode): Promise<T | undefined> {
  return new Promise((resolve) => {
    const id = uid('dlg');
    const entry: HostedDialog = {
      id,
      render: render as HostedDialog['render'],
      resolve: resolve as (v: unknown) => void,
    };
    useDialogStore.setState((s) => ({ dialogs: [...s.dialogs, entry] }));
  });
}

function closeDialog(id: string, value: unknown) {
  const d = useDialogStore.getState().dialogs.find((x) => x.id === id);
  useDialogStore.setState((s) => ({ dialogs: s.dialogs.filter((x) => x.id !== id) }));
  d?.resolve(value);
}

export function DialogHost() {
  const dialogs = useDialogStore((s) => s.dialogs);
  return (
    <>
      {dialogs.map((d) => (
        <DialogSlot key={d.id} d={d} />
      ))}
    </>
  );
}

function DialogSlot({ d }: { d: HostedDialog }) {
  return <>{d.render((v) => closeDialog(d.id, v))}</>;
}

/* --------------------------------------------------------------- helpers */

export interface ConfirmOptions {
  title: string;
  message?: ReactNode;
  confirmLabel?: string;
  cancelLabel?: string;
  danger?: boolean;
  icon?: ReactNode;
}

export function confirmDialog(o: ConfirmOptions): Promise<boolean> {
  return openDialog<boolean>((close) => (
    <Modal
      title={o.title}
      icon={o.icon}
      onClose={() => close(false)}
      width={420}
      footer={
        <>
          <button className="btn" onClick={() => close(false)}>
            {o.cancelLabel ?? 'Cancel'}
          </button>
          <button className={`btn ${o.danger ? 'btn-danger' : 'btn-primary'}`} data-autofocus onClick={() => close(true)}>
            {o.confirmLabel ?? 'OK'}
          </button>
        </>
      }
    >
      {typeof o.message === 'string' ? <p className="dialog-text">{o.message}</p> : o.message}
    </Modal>
  )).then((v) => Boolean(v));
}

export function alertDialog(title: string, message?: ReactNode): Promise<void> {
  return openDialog<void>((close) => (
    <Modal
      title={title}
      onClose={() => close()}
      width={420}
      footer={
        <button className="btn btn-primary" data-autofocus onClick={() => close()}>
          OK
        </button>
      }
    >
      {typeof message === 'string' ? <p className="dialog-text">{message}</p> : message}
    </Modal>
  )).then(() => undefined);
}

export type SaveChoice = 'save' | 'discard' | 'cancel';

export function saveChangesDialog(name: string): Promise<SaveChoice> {
  return openDialog<SaveChoice>((close) => (
    <Modal
      title="Save your changes?"
      onClose={() => close('cancel')}
      width={440}
      footer={
        <>
          <button className="btn btn-ghost" onClick={() => close('discard')}>
            Don’t save
          </button>
          <span className="spacer" />
          <button className="btn" onClick={() => close('cancel')}>
            Cancel
          </button>
          <button className="btn btn-primary" data-autofocus onClick={() => close('save')}>
            Save
          </button>
        </>
      }
    >
      <p className="dialog-text">
        Do you want to save the changes you made to <strong>{name}</strong>? Your changes will be lost if you don’t save them.
      </p>
    </Modal>
  )).then((v) => v ?? 'cancel');
}

export interface PromptOptions {
  title: string;
  label?: string;
  value?: string;
  placeholder?: string;
  confirmLabel?: string;
  validate?: (v: string) => string | null;
  multiline?: boolean;
}

export function promptDialog(o: PromptOptions): Promise<string | null> {
  return openDialog<string | null>((close) => <PromptBody o={o} close={close} />).then((v) => (v === undefined ? null : v));
}

function PromptBody({ o, close }: { o: PromptOptions; close: (v?: string | null) => void }) {
  const [value, setValue] = useState(o.value ?? '');
  const error = o.validate?.(value) ?? null;
  const submit = () => {
    if (!error) close(value);
  };
  return (
    <Modal
      title={o.title}
      onClose={() => close(null)}
      width={440}
      footer={
        <>
          <button className="btn" onClick={() => close(null)}>
            Cancel
          </button>
          <button className="btn btn-primary" disabled={Boolean(error)} onClick={submit}>
            {o.confirmLabel ?? 'OK'}
          </button>
        </>
      }
    >
      <label className="field">
        {o.label && <span className="field-label">{o.label}</span>}
        {o.multiline ? (
          <textarea className="input" rows={5} value={value} placeholder={o.placeholder} onChange={(e) => setValue(e.target.value)} />
        ) : (
          <input
            className="input"
            value={value}
            placeholder={o.placeholder}
            onChange={(e) => setValue(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') submit();
            }}
          />
        )}
        {error && value && <span className="field-error">{error}</span>}
      </label>
    </Modal>
  );
}
