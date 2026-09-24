import type { Tab } from '@/app/workspace';

export default function SlidesEditor({ tab }: { tab: Tab; active: boolean }) {
  return <div className="editor-loading">{tab.title}</div>;
}
