import type { ContextMenuParams } from '@/shared/types';
import type { MenuItem } from '@/ui/menu';

/**
 * The active editor can provide extra context-menu items for the native (spell-checked)
 * context menu that Electron reports. Returning undefined falls back to Cut/Copy/Paste.
 */
export const contextMenuExtender: { current: ((p: ContextMenuParams) => MenuItem[] | undefined) | null } = { current: null };
