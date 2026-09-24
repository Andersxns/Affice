import { defaultSettings, type DocSettings } from './model';
import type { LoadedDoc } from './formats';

const today = () => new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' });

function withSettings(patch: Partial<DocSettings>): DocSettings {
  return { ...defaultSettings(), ...patch };
}

const TEMPLATES: Record<string, () => LoadedDoc> = {
  'doc-resume': () => ({
    settings: withSettings({
      styleSet: 'modern',
      page: { ...defaultSettings().page, margins: { top: 0.7, right: 0.8, bottom: 0.7, left: 0.8 } },
      styles: { normal: { font: 'Arial', size: 10.5, color: '#1f2937', after: 4, lineHeight: 1.25 } },
    }),
    content: `
<p data-style="title">Alex Morgan</p>
<p data-style="subtitle">Product Designer · alex.morgan@email.com · +1 555 0100 · Portland, OR</p>
<hr>
<h2>Profile</h2>
<p>Designer with 7 years of experience turning complex problems into simple, delightful products. I love working closely with engineers and customers, and I care about accessibility, performance and craft.</p>
<h2>Experience</h2>
<h3>Senior Product Designer — Northwind Labs</h3>
<p><em>2021 – Present · Remote</em></p>
<ul>
<li><p>Led the redesign of the onboarding flow, increasing activation by <strong>32%</strong>.</p></li>
<li><p>Built and maintained a cross-platform design system used by 40+ engineers.</p></li>
<li><p>Mentored four junior designers and ran weekly design critiques.</p></li>
</ul>
<h3>Product Designer — Contoso</h3>
<p><em>2017 – 2021 · Seattle, WA</em></p>
<ul>
<li><p>Designed the mobile checkout used by 2M monthly customers.</p></li>
<li><p>Partnered with research to run 60+ usability sessions.</p></li>
</ul>
<h2>Education</h2>
<p><strong>B.A. Interaction Design</strong> — University of Washington, 2017</p>
<h2>Skills</h2>
<table data-style="plain"><tbody>
<tr><td><p><strong>Design</strong></p><p>UX research, prototyping, visual design, accessibility</p></td><td><p><strong>Tools</strong></p><p>Figma, Affice, HTML/CSS, user testing platforms</p></td></tr>
</tbody></table>
<p></p>`,
    comments: [],
  }),

  'doc-letter': () => ({
    settings: withSettings({ styleSet: 'office' }),
    content: `
<p><strong>Jordan Lee</strong><br>123 Maple Street<br>Springfield, ST 12345<br>jordan.lee@email.com</p>
<p>${today()}</p>
<p>Hiring Manager<br>Fabrikam, Inc.<br>500 Market Avenue<br>Springfield, ST 12345</p>
<p>Dear Hiring Manager,</p>
<p>I am excited to apply for the <strong>Marketing Coordinator</strong> position at Fabrikam. With three years of experience planning campaigns and a passion for clear, human communication, I would love to help your team tell its story.</p>
<p>In my current role at Litware, I coordinated product launches across email, social and events, growing our newsletter audience by 45% in one year. I enjoy turning data into ideas, and I thrive when working with designers, writers and sales teams.</p>
<p>I would welcome the chance to discuss how my experience can support Fabrikam’s goals. Thank you for your time and consideration.</p>
<p>Sincerely,</p>
<p></p>
<p>Jordan Lee</p>`,
    comments: [],
  }),

  'doc-report': () => ({
    settings: withSettings({
      styleSet: 'classic',
      footer: { left: '', center: 'Page {PAGE} of {PAGES}', right: '' },
      header: { left: 'Project Aurora', center: '', right: '{DATE}' },
      differentFirstPage: true,
    }),
    content: `
<p></p><p></p><p></p>
<p data-style="title" style="text-align:center">Project Aurora</p>
<p data-style="subtitle" style="text-align:center">Quarterly Progress Report · Q3</p>
<p style="text-align:center">Prepared by the Aurora team · ${today()}</p>
<div class="page-break" data-page-break="true"></div>
<div data-toc="" data-title="Contents"></div>
<h1>Executive summary</h1>
<p>This quarter the team delivered the first public beta, reduced infrastructure costs and grew our pilot program to twelve customers. This report summarises progress, key metrics and the plan for next quarter.</p>
<h1>Goals and results</h1>
<h2>Delivery</h2>
<p>We shipped three major releases on schedule. Customer-reported defects fell for the second quarter in a row.</p>
<h2>Metrics</h2>
<table data-style="accent"><tbody>
<tr><th><p>Metric</p></th><th><p>Target</p></th><th><p>Actual</p></th><th><p>Status</p></th></tr>
<tr><td><p>Active pilot customers</p></td><td><p>10</p></td><td><p>12</p></td><td><p>Ahead</p></td></tr>
<tr><td><p>Uptime</p></td><td><p>99.5%</p></td><td><p>99.8%</p></td><td><p>On track</p></td></tr>
<tr><td><p>Cloud cost / month</p></td><td><p>$18k</p></td><td><p>$15.2k</p></td><td><p>Ahead</p></td></tr>
</tbody></table>
<h1>Risks</h1>
<ul><li><p>Hiring for two senior roles is behind plan.</p></li><li><p>A key vendor contract renews in November.</p></li></ul>
<h1>Next quarter</h1>
<ol><li><p>General availability launch.</p></li><li><p>Self-serve billing.</p></li><li><p>Expand the pilot to 25 customers.</p></li></ol>`,
    comments: [],
  }),

  'doc-meeting': () => ({
    settings: withSettings({ styleSet: 'minimal' }),
    content: `
<h1>Weekly team meeting</h1>
<p><strong>Date:</strong> ${today()} · <strong>Time:</strong> 10:00 · <strong>Location:</strong> Room 4B / Video call</p>
<p><strong>Attendees:</strong> Sam, Priya, Diego, Mei, Olu</p>
<h2>Agenda</h2>
<ol><li><p>Review last week’s action items</p></li><li><p>Launch readiness</p></li><li><p>Open questions</p></li></ol>
<h2>Notes</h2>
<ul><li><p>Launch checklist is 80% complete; marketing assets arrive Thursday.</p></li><li><p>Support team needs updated FAQ by Friday.</p></li></ul>
<h2>Decisions</h2>
<ul><li><p>Launch date confirmed for the 14th.</p></li></ul>
<h2>Action items</h2>
<ul data-type="taskList">
<li data-type="taskItem" data-checked="false"><p>Finalise FAQ — <strong>Priya</strong> (Fri)</p></li>
<li data-type="taskItem" data-checked="false"><p>Book launch-day war room — <strong>Diego</strong> (Wed)</p></li>
<li data-type="taskItem" data-checked="true"><p>Share draft release notes — <strong>Mei</strong></p></li>
</ul>`,
    comments: [],
  }),

  'doc-essay': () => ({
    settings: withSettings({
      styleSet: 'academic',
      header: { left: '', center: '', right: 'Lastname {PAGE}' },
    }),
    content: `
<p>Your Name</p>
<p>Professor Name</p>
<p>Course Name</p>
<p>${today()}</p>
<p style="text-align:center">The Title of Your Essay</p>
<p style="text-indent:48px">Begin your essay here. The first paragraph introduces your topic and ends with a clear thesis statement that tells the reader what you will argue and why it matters.</p>
<p style="text-indent:48px">Each body paragraph should focus on one main idea that supports your thesis. Start with a topic sentence, provide evidence such as quotations or data, and explain how that evidence supports your argument.</p>
<p style="text-indent:48px">In your conclusion, restate your thesis in new words, summarise your key points and explain the broader significance of your argument.</p>
<div class="page-break" data-page-break="true"></div>
<p style="text-align:center">Works Cited</p>
<p style="margin-left:48px;text-indent:-48px">Lastname, Firstname. <em>Title of Book</em>. Publisher, Year.</p>
<p style="margin-left:48px;text-indent:-48px">Lastname, Firstname. “Title of Article.” <em>Journal Name</em>, vol. 1, no. 2, Year, pp. 10–20.</p>`,
    comments: [],
  }),

  'doc-newsletter': () => ({
    settings: withSettings({ styleSet: 'vibrant' }),
    content: `
<p data-style="title">The Monthly Spark</p>
<p data-style="subtitle">News, wins and ideas from our community · ${new Date().toLocaleDateString(undefined, { month: 'long', year: 'numeric' })}</p>
<hr>
<h1>Highlights</h1>
<table data-style="plain"><tbody><tr>
<td><h3>New community garden</h3><p>Thanks to 60 volunteers, the Riverside garden opened with 40 raised beds. Sign up for a plot at the front desk.</p></td>
<td><h3>Summer workshop series</h3><p>Join free Saturday sessions on budgeting, photography and first aid. Seats are limited — register early!</p></td>
</tr></tbody></table>
<h1>Member spotlight</h1>
<p><strong>Maria Gonzales</strong> has volunteered every week for five years. “I love seeing neighbours become friends,” she says. Thank you, Maria!</p>
<blockquote><p>“Alone we can do so little; together we can do so much.” — Helen Keller</p></blockquote>
<h1>Upcoming events</h1>
<table data-style="banded"><tbody>
<tr><th><p>Date</p></th><th><p>Event</p></th><th><p>Where</p></th></tr>
<tr><td><p>June 8</p></td><td><p>Neighbourhood clean-up</p></td><td><p>Main Park</p></td></tr>
<tr><td><p>June 15</p></td><td><p>Photography workshop</p></td><td><p>Library, Room 2</p></td></tr>
<tr><td><p>June 29</p></td><td><p>Summer picnic</p></td><td><p>Riverside Garden</p></td></tr>
</tbody></table>`,
    comments: [],
  }),
};

export function docTemplate(id: string): LoadedDoc {
  const t = TEMPLATES[id];
  if (!t) return { settings: defaultSettings(), content: { type: 'doc', content: [{ type: 'paragraph' }] }, comments: [] };
  return t();
}
