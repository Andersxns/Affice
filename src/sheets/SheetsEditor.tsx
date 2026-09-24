import type { Tab } from '@/app/workspace';

export default function SheetsEditor({ tab }: { tab: Tab; active: boolean }) {
  return <div className="editor-loading">{tab.title}</div>;
}
