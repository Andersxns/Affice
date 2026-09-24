import type { AppKind } from '@/shared/types';

export interface TemplateMeta {
  id: string;
  kind: AppKind;
  name: string;
  description: string;
  /** Visual hint for the dashboard preview */
  preview: {
    accent: string;
    bg?: string;
    variant: string;
  };
}

export const TEMPLATES: TemplateMeta[] = [
  { id: 'doc-resume', kind: 'doc', name: 'Modern résumé', description: 'Clean one-page CV with sections for experience and skills.', preview: { accent: '#2f6dff', variant: 'resume' } },
  { id: 'doc-letter', kind: 'doc', name: 'Cover letter', description: 'Professional letter with your details and a friendly layout.', preview: { accent: '#0d9488', variant: 'letter' } },
  { id: 'doc-report', kind: 'doc', name: 'Project report', description: 'Title page, contents, headings and a results table.', preview: { accent: '#7c3aed', variant: 'report' } },
  { id: 'doc-meeting', kind: 'doc', name: 'Meeting notes', description: 'Agenda, decisions and action items with checkboxes.', preview: { accent: '#f26a26', variant: 'notes' } },
  { id: 'doc-essay', kind: 'doc', name: 'Essay', description: 'Double-spaced academic essay with heading and works cited.', preview: { accent: '#475569', variant: 'essay' } },
  { id: 'doc-newsletter', kind: 'doc', name: 'Newsletter', description: 'Bold header, highlights and a two-column table layout.', preview: { accent: '#db2777', variant: 'newsletter' } },

  { id: 'sheet-budget', kind: 'sheet', name: 'Monthly budget', description: 'Income, expenses, totals and a spending chart.', preview: { accent: '#17a35a', variant: 'budget' } },
  { id: 'sheet-invoice', kind: 'sheet', name: 'Invoice', description: 'Itemised invoice that calculates tax and totals.', preview: { accent: '#2f6dff', variant: 'invoice' } },
  { id: 'sheet-tracker', kind: 'sheet', name: 'Task tracker', description: 'Status, owners, due dates and progress bars.', preview: { accent: '#f26a26', variant: 'tracker' } },
  { id: 'sheet-grades', kind: 'sheet', name: 'Gradebook', description: 'Scores, averages and letter grades per student.', preview: { accent: '#7c3aed', variant: 'grades' } },
  { id: 'sheet-schedule', kind: 'sheet', name: 'Weekly schedule', description: 'Plan your week hour by hour.', preview: { accent: '#0891b2', variant: 'schedule' } },

  { id: 'slides-pitch', kind: 'slides', name: 'Pitch deck', description: 'Problem, solution, market and team — bold and modern.', preview: { accent: '#2f6dff', bg: '#0b1b4d', variant: 'pitch' } },
  { id: 'slides-lesson', kind: 'slides', name: 'Lesson', description: 'Friendly slides for teaching and workshops.', preview: { accent: '#17a35a', bg: '#f3faf5', variant: 'lesson' } },
  { id: 'slides-minimal', kind: 'slides', name: 'Minimal', description: 'Lots of white space and elegant type.', preview: { accent: '#111827', bg: '#ffffff', variant: 'minimal' } },
  { id: 'slides-midnight', kind: 'slides', name: 'Midnight', description: 'Dark theme with vivid gradient accents.', preview: { accent: '#a855f7', bg: '#0f0f1a', variant: 'midnight' } },
  { id: 'slides-sunset', kind: 'slides', name: 'Sunset', description: 'Warm gradients for creative stories.', preview: { accent: '#f26a26', bg: '#fff4ec', variant: 'sunset' } },
];
