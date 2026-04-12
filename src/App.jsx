import React, { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Users,
  Gauge,
  Trophy,
  AlertTriangle,
  Upload,
  Save,
  BarChart3,
  Settings2,
  CalendarRange,
  UserCircle2,
  ClipboardList,
  Download,
  Search,
  BookOpen,
} from 'lucide-react';
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Radar,
  RadarChart,
  PolarGrid,
  PolarAngleAxis,
  PolarRadiusAxis,
  LineChart,
  Line,
  Legend,
} from 'recharts';

const CATEGORY_CONFIG = [
  { key: 'storiesValidated', short: 'Stories Validated', full: 'Number of Stories validated with highest quality', max: 20, below: 8, average: 12, above: 16, exceed: 20 },
  { key: 'bugsValidated', short: 'Bugs Validated', full: 'Number of bugs validated', max: 5, below: 2, average: 3, above: 4, exceed: 5 },
  { key: 'testCases', short: 'Test Cases', full: 'Number of Test cases executed during Release', max: 20, below: 8, average: 12, above: 16, exceed: 20 },
  { key: 'bugsRaised', short: 'Bugs Raised', full: 'Number of bugs raised', max: 5, below: 2, average: 3, above: 4, exceed: 5 },
  { key: 'automationScenarios', short: 'Automation Scenarios', full: 'Number of Automation Scenarios created', max: 20, below: 8, average: 12, above: 16, exceed: 20 },
  { key: 'prReviews', short: 'PR Reviews', full: 'Number of PR reviews completed', max: 10, below: 3, average: 5, above: 8, exceed: 10 },
  { key: 'initiativesTaken', short: 'Initiatives Taken', full: 'Number of initiatives taken', max: 5, below: 2, average: 3, above: 4, exceed: 5 },
  { key: 'trainingSessions', short: 'Training Sessions', full: 'Number of training sessions held', max: 5, below: 2, average: 3, above: 4, exceed: 5 },
  { key: 'collaboration', short: 'Collaboration', full: 'Team collaboration (helping hand, reviews, support)', max: 10, below: 5, average: 7, above: 8, exceed: 10 },
];

const RATING_RUBRIC = [
  { rating: 'Below Average', points: '40 pts', description: 'Below team average' },
  { rating: 'Average', points: '60 pts', description: 'At team average' },
  { rating: 'Above Average', points: '80 pts', description: 'Up to 20% above team average' },
  { rating: 'Exceeding', points: '100 pts', description: 'More than 20% above team average' },
];

const PERIODS = [
  'Mar 1-15, 2026','Mar 16-31, 2026','Apr 1-15, 2026','Apr 16-30, 2026','May 1-15, 2026','May 16-31, 2026',
  'Jun 1-15, 2026','Jun 16-30, 2026','Jul 1-15, 2026','Jul 16-31, 2026','Aug 1-15, 2026','Aug 16-31, 2026',
  'Sep 1-15, 2026','Sep 16-30, 2026','Oct 1-15, 2026','Oct 16-31, 2026','Nov 1-15, 2026','Nov 16-30, 2026',
  'Dec 1-15, 2026','Dec 16-31, 2026','Jan 1-15, 2027','Jan 16-31, 2027','Feb 1-15, 2027','Feb 16-28, 2027'
];

const DEFAULT_MEMBERS = ['Isabelle', 'Agjeliki', 'Teddy', 'Krish', 'Chris'];
const ALL_MEMBERS = 'All Members';

const HEADER_MAP = {
  storiesValidated: ['Number of Stories validate', 'Number of Stories validated', 'stories'],
  bugsValidated: ['Number of bugs validated', 'bugs validated'],
  testCases: ['Number of Test cases execu', 'Number of Test cases executed', 'test cases'],
  bugsRaised: ['Number of bugs raised', 'bugs raised'],
  automationScenarios: ['Number of Automation Scena', 'Automation Scenarios', 'automation'],
  prReviews: ['Number of PR reviews compl', 'PR reviews', 'pr reviews'],
  initiativesTaken: ['Number of initiatives taken', 'initiatives'],
  trainingSessions: ['Number of training session', 'training'],
  collaboration: ['Team collaboration', 'collaboration'],
};

function buildInitialData(members = DEFAULT_MEMBERS) {
  const data = {};
  members.forEach((member) => {
    data[member] = {};
    PERIODS.forEach((period) => {
      data[member][period] = Object.fromEntries(CATEGORY_CONFIG.map((c) => [c.key, '']));
    });
  });
  return data;
}

function toNumber(value) {
  if (value === '' || value === null || value === undefined) return '';
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : '';
}

function getBandLabel(score) {
  if (score === '') return '';
  if (score > 80) return 'Exceeding';
  if (score > 60) return 'Above Average';
  if (score > 40) return 'Average';
  return 'Below Average';
}

function getBadgeClass(label) {
  switch (label) {
    case 'Exceeding': return 'badge success';
    case 'Above Average': return 'badge info';
    case 'Average': return 'badge warning';
    case 'Below Average': return 'badge danger';
    default: return 'badge';
  }
}

function scoreCategory(rawValue, teamAverage, config) {
  const raw = toNumber(rawValue);
  if (raw === '') return '';
  if (teamAverage === 0) return raw > 0 ? config.exceed : config.average;
  if (raw > teamAverage * 1.2) return config.exceed;
  if (raw > teamAverage) return config.above;
  if (raw === teamAverage) return config.average;
  return config.below;
}

function findColumnIndex(headers, key) {
  const candidates = HEADER_MAP[key] || [];
  return headers.findIndex((header) => {
    const normalized = String(header || '').toLowerCase();
    return candidates.some((candidate) => normalized.includes(candidate.toLowerCase()));
  });
}

function importWorkbook(file, setMembers, setData, setSelectedMember, setSelectedPeriod, setFileName) {
  const reader = new FileReader();
  reader.onload = (evt) => {
    const wb = XLSX.read(evt.target.result, { type: 'binary' });
    const memberSheets = wb.SheetNames.filter((name) => !['Summary', 'Team Data'].includes(name));
    const nextMembers = memberSheets.length ? memberSheets : DEFAULT_MEMBERS;
    const imported = buildInitialData(nextMembers);

    nextMembers.forEach((member) => {
      const ws = wb.Sheets[member];
      if (!ws) return;
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const headerRowIndex = rows.findIndex((row) => row[0] === 'Biweekly Period' && String(row[1] || '').includes('Stories'));
      if (headerRowIndex === -1) return;
      const headers = rows[headerRowIndex];
      for (let i = headerRowIndex + 1; i < headerRowIndex + 1 + PERIODS.length; i += 1) {
        const row = rows[i] || [];
        const period = row[0];
        if (!PERIODS.includes(period)) continue;
        CATEGORY_CONFIG.forEach((category) => {
          const colIndex = findColumnIndex(headers, category.key);
          if (colIndex >= 0) {
            const value = row[colIndex];
            imported[member][period][category.key] = value === '' ? '' : toNumber(value);
          }
        });
      }
    });

    setMembers(nextMembers);
    setData(imported);
    setSelectedMember(ALL_MEMBERS);
    setSelectedPeriod(PERIODS[0]);
    setFileName(file.name);
  };
  reader.readAsBinaryString(file);
}

function exportWorkbook(members, data, computed) {
  const wb = XLSX.utils.book_new();

  const summaryRows = [
    ['QA Team Performance Tracker - Scoring Rubric'],
    [],
    ['Rating', 'Points Range', 'Description'],
    ...RATING_RUBRIC.map((r) => [r.rating, r.points, r.description]),
    [],
    ['Category', 'Max Pts', 'Below Avg', 'Average', 'Above Avg', 'Exceeding'],
    ...CATEGORY_CONFIG.map((c) => [c.full, c.max, c.below, c.average, c.above, c.exceed]),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryRows), 'Summary');

  const teamRows = [
    ['Biweekly Period', ...CATEGORY_CONFIG.map((c) => c.short)],
    ...PERIODS.map((period) => [
      period,
      ...CATEGORY_CONFIG.map((c) => computed.teamAverages[period][c.key]),
    ]),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(teamRows), 'Team Data');

  members.forEach((member) => {
    const rawRows = [
      [`${member} - QA Performance Tracker`],
      ['Enter raw counts below. Scores and ratings calculate automatically in the app.'],
      [],
      ['RAW DATA ENTRY'],
      ['Biweekly Period', ...CATEGORY_CONFIG.map((c) => c.short)],
      ...PERIODS.map((period) => [period, ...CATEGORY_CONFIG.map((c) => data[member][period][c.key])]),
      [],
      ['SCORED POINTS'],
      ['Biweekly Period', ...CATEGORY_CONFIG.map((c) => c.short), 'Score', 'Rating'],
      ...PERIODS.map((period) => {
        const record = computed.memberPeriods[member][period];
        return [period, ...CATEGORY_CONFIG.map((c) => record.categoryPoints[c.key]), record.totalScore, record.rating];
      }),
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rawRows), member);
  });

  XLSX.writeFile(wb, 'qa-team-performance-export.xlsx');
}

function StatCard({ icon: Icon, label, value, subtext, accent }) {
  return (
    <div className="stat-card">
      <div>
        <div className="stat-label">{label}</div>
        <div className="stat-value">{value}</div>
        {subtext ? <div className="stat-subtext">{subtext}</div> : null}
      </div>
      <div className={`stat-icon ${accent}`}><Icon size={22} /></div>
    </div>
  );
}

function FilterToolbar({ selectedPeriod, setSelectedPeriod, selectedMember, setSelectedMember, search, setSearch, members }) {
  return (
    <section className="toolbar card">
      <div className="toolbar-grid">
        <div>
          <label>Biweekly Period</label>
          <select value={selectedPeriod} onChange={(e) => setSelectedPeriod(e.target.value)}>
            {PERIODS.map((period) => <option key={period}>{period}</option>)}
          </select>
        </div>
        <div>
          <label>Selected Member</label>
          <select value={selectedMember} onChange={(e) => setSelectedMember(e.target.value)}>
            <option value={ALL_MEMBERS}>{ALL_MEMBERS}</option>
            {members.map((member) => <option key={member}>{member}</option>)}
          </select>
        </div>
        <div>
          <label>Search Team Member</label>
          <div className="search-wrap">
            <Search size={16} />
            <input type="text" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Filter by team member" />
          </div>
        </div>
      </div>
    </section>
  );
}

function App() {
  const [members, setMembers] = useState(DEFAULT_MEMBERS);
  const [data, setData] = useState(buildInitialData(DEFAULT_MEMBERS));
  const [selectedMember, setSelectedMember] = useState(ALL_MEMBERS);
  const [selectedPeriod, setSelectedPeriod] = useState(PERIODS[0]);
  const [selectedCategory, setSelectedCategory] = useState(CATEGORY_CONFIG[0].key);
  const [quickMember, setQuickMember] = useState(DEFAULT_MEMBERS[0]);
  const [quickPeriod, setQuickPeriod] = useState(PERIODS[0]);
  const [quickValue, setQuickValue] = useState('');
  const [search, setSearch] = useState('');
  const [tab, setTab] = useState('overview');
  const [fileName, setFileName] = useState('');

  const computed = useMemo(() => {
    const teamAverages = {};
    PERIODS.forEach((period) => {
      teamAverages[period] = {};
      CATEGORY_CONFIG.forEach((category) => {
        const values = members.map((member) => toNumber(data[member]?.[period]?.[category.key])).filter((value) => value !== '');
        teamAverages[period][category.key] = values.length ? Number((values.reduce((sum, value) => sum + value, 0) / values.length).toFixed(2)) : 0;
      });
    });

    const memberPeriods = {};
    members.forEach((member) => {
      memberPeriods[member] = {};
      PERIODS.forEach((period) => {
        const categoryPoints = {};
        let hasAnyValue = false;
        CATEGORY_CONFIG.forEach((category) => {
          const raw = data[member]?.[period]?.[category.key] ?? '';
          if (raw !== '') hasAnyValue = true;
          categoryPoints[category.key] = scoreCategory(raw, teamAverages[period][category.key], category);
        });
        const totalScore = hasAnyValue ? CATEGORY_CONFIG.reduce((sum, category) => sum + (categoryPoints[category.key] || 0), 0) : '';
        const rating = totalScore === '' ? '' : getBandLabel(totalScore);
        memberPeriods[member][period] = { categoryPoints, totalScore, rating };
      });
    });

    return { teamAverages, memberPeriods };
  }, [members, data]);

  const searchedMembers = useMemo(() => {
    const query = search.trim().toLowerCase();
    if (!query) return members;
    return members.filter((member) => member.toLowerCase().includes(query));
  }, [members, search]);

  const activeSelection = useMemo(() => {
    const query = search.trim().toLowerCase();
    if (!query) return selectedMember;
    const exact = members.find((member) => member.toLowerCase() === query);
    if (exact) return exact;
    const partialMatches = members.filter((member) => member.toLowerCase().includes(query));
    return partialMatches.length === 1 ? partialMatches[0] : ALL_MEMBERS;
  }, [members, search, selectedMember]);

  const visibleMembers = useMemo(() => {
    if (activeSelection === ALL_MEMBERS) return members;
    return members.filter((member) => member === activeSelection);
  }, [members, activeSelection]);

  const periodData = useMemo(() => {
    return visibleMembers.map((member) => {
      const scoreRecord = computed.memberPeriods[member][selectedPeriod];
      const raw = data[member][selectedPeriod];
      return {
        name: member,
        score: scoreRecord.totalScore || 0,
        rating: scoreRecord.rating || '—',
        points: scoreRecord.categoryPoints,
        raw,
      };
    }).sort((a, b) => b.score - a.score);
  }, [visibleMembers, computed, data, selectedPeriod]);

  const summary = useMemo(() => {
    const avgRows = members.map((member) => {
      const rows = PERIODS.map((period) => computed.memberPeriods[member][period].totalScore).filter((v) => v !== '');
      const averageScore = rows.length ? Number((rows.reduce((sum, v) => sum + v, 0) / rows.length).toFixed(1)) : 0;
      return { member, averageScore };
    }).filter((row) => row.averageScore > 0);

    const top = [...avgRows].sort((a, b) => b.averageScore - a.averageScore)[0];
    const support = [...avgRows].sort((a, b) => a.averageScore - b.averageScore)[0];
    const averageScore = avgRows.length ? Number((avgRows.reduce((sum, row) => sum + row.averageScore, 0) / avgRows.length).toFixed(1)) : 0;
    return { top, support, averageScore };
  }, [members, computed]);

  const effectiveMember = activeSelection === ALL_MEMBERS ? null : activeSelection;
  const selectedRecord = effectiveMember ? computed.memberPeriods[effectiveMember][selectedPeriod] : null;

  const overviewPeriodData = useMemo(() => {
    return visibleMembers.map((member) => {
      const scoreRecord = computed.memberPeriods[member][selectedPeriod];
      const raw = data[member][selectedPeriod];
      return {
        name: member,
        score: scoreRecord.totalScore || 0,
        rating: scoreRecord.rating || '—',
        points: scoreRecord.categoryPoints,
        raw,
      };
    }).sort((a, b) => b.score - a.score);
  }, [visibleMembers, computed, data, selectedPeriod]);

  const overviewTrendMember = useMemo(() => {
    if (activeSelection && activeSelection !== ALL_MEMBERS) return activeSelection;
    return members[0] || null;
  }, [activeSelection, members]);

  const overviewTrendData = useMemo(() => {
    if (!overviewTrendMember) return [];
    return PERIODS.map((period) => ({ period, total: computed.memberPeriods[overviewTrendMember][period].totalScore || 0 }));
  }, [overviewTrendMember, computed]);

  const individualRadar = useMemo(() => {
    if (!effectiveMember) return [];
    return CATEGORY_CONFIG.map((category) => ({
      category: category.short,
      score: computed.memberPeriods[effectiveMember][selectedPeriod].categoryPoints[category.key] || 0,
      raw: data[effectiveMember][selectedPeriod][category.key] || 0,
    }));
  }, [effectiveMember, selectedPeriod, computed, data]);

  const memberVsTeamData = useMemo(() => {
    if (!effectiveMember) return [];
    return CATEGORY_CONFIG.map((category) => ({
      category: category.short,
      member: Number(data[effectiveMember][selectedPeriod][category.key] || 0),
      team: Number(computed.teamAverages[selectedPeriod][category.key] || 0),
    }));
  }, [effectiveMember, selectedPeriod, computed, data]);

  const handleFieldChange = (member, period, category, value) => {
    setData((prev) => ({
      ...prev,
      [member]: {
        ...prev[member],
        [period]: {
          ...prev[member][period],
          [category]: value === '' ? '' : Number(value),
        },
      },
    }));
  };

  const applyQuickUpdate = () => {
    handleFieldChange(quickMember, quickPeriod, selectedCategory, quickValue === '' ? '' : Number(quickValue));
    setQuickValue('');
  };

  return (
    <div className="app-shell">
      <header className="hero">
        <div>
          <span className="eyebrow">QA Performance Manager</span>
          <h1>Biweekly QA performance analytics across delivery, quality, collaboration, and growth metrics.</h1>
          <p>Input raw values, calculate scores automatically, review rankings, and inspect individual performance from one consolidated dashboard.</p>
        </div>
        <div className="hero-actions">
          <label className="upload-btn">
            <Upload size={18} /> Import workbook
            <input type="file" accept=".xlsx,.xls" onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) importWorkbook(file, setMembers, setData, setSelectedMember, setSelectedPeriod, setFileName);
            }} />
          </label>
          <button className="secondary-btn" onClick={() => exportWorkbook(members, data, computed)}>
            <Download size={18} /> Export workbook
          </button>
          <div className="file-name">{fileName || 'No workbook uploaded yet.'}</div>
        </div>
      </header>

      <section className="stats-grid">
        <StatCard icon={Users} label="Team Members" value={members.length} accent="blue" />
        <StatCard icon={Gauge} label="Average Score" value={summary.averageScore || '—'} accent="violet" />
        <StatCard icon={Trophy} label="Top Performer" value={summary.top?.member || '—'} subtext={summary.top ? `${summary.top.averageScore} avg score` : ''} accent="green" />
        <StatCard icon={AlertTriangle} label="Needs Support" value={summary.support?.member || '—'} subtext={summary.support ? `${summary.support.averageScore} avg score` : ''} accent="orange" />
      </section>

      <section className="control-grid static-cards">
        <div className="card control-card">
          <div className="card-title"><UserCircle2 size={18} /> Quick entry</div>
          <div className="control-row three">
            <div>
              <label>Member</label>
              <select value={quickMember} onChange={(e) => setQuickMember(e.target.value)}>
                {members.map((member) => <option key={member}>{member}</option>)}
              </select>
            </div>
            <div>
              <label>Biweekly period</label>
              <select value={quickPeriod} onChange={(e) => setQuickPeriod(e.target.value)}>
                {PERIODS.map((period) => <option key={period}>{period}</option>)}
              </select>
            </div>
            <div>
              <label>Category</label>
              <select value={selectedCategory} onChange={(e) => setSelectedCategory(e.target.value)}>
                {CATEGORY_CONFIG.map((category) => <option value={category.key} key={category.key}>{category.short}</option>)}
              </select>
            </div>
          </div>
          <div className="control-row quick-update-row">
            <div className="quick-value-box">
              <label>Value</label>
              <input type="number" value={quickValue} onChange={(e) => setQuickValue(e.target.value)} placeholder="Enter raw count" />
            </div>
            <button onClick={applyQuickUpdate}><Save size={16} /> Save update</button>
          </div>
        </div>

        <div className="card rubric-card">
          <div className="card-title"><Settings2 size={18} /> Scoring logic</div>
          <p className="card-subtitle">Rating thresholds applied across the workbook scoring model.</p>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Rating</th>
                  <th>Points</th>
                  <th>Description</th>
                </tr>
              </thead>
              <tbody>
                {RATING_RUBRIC.map((item) => (
                  <tr key={item.rating}>
                    <td>{item.rating}</td>
                    <td>{item.points}</td>
                    <td>{item.description}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </section>

      <nav className="tabs">
        {['overview', 'members', 'individual', 'rubric'].map((item) => (
          <button key={item} className={tab === item ? 'tab active' : 'tab'} onClick={() => setTab(item)}>
            {item[0].toUpperCase() + item.slice(1)}
          </button>
        ))}
      </nav>

      {tab === 'overview' && (
        <>
          <FilterToolbar
            selectedPeriod={selectedPeriod}
            setSelectedPeriod={setSelectedPeriod}
            selectedMember={selectedMember}
            setSelectedMember={setSelectedMember}
            search={search}
            setSearch={setSearch}
            members={members}
          />
          <section className="chart-grid-two">
            <div className="card chart-card">
              <div className="card-title"><BarChart3 size={18} /> Period ranking</div>
              <p className="card-subtitle">Total scored points for the selected biweekly period.</p>
              <ResponsiveContainer width="100%" height={320}>
                <BarChart data={overviewPeriodData}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="name" hide />
                  <YAxis domain={[0, 100]} />
                  <Tooltip formatter={(value) => [value, 'Score']} />
                  <Bar dataKey="score" fill="#5b7fff" radius={[8, 8, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div className="card chart-card">
              <div className="card-title"><CalendarRange size={18} /> Selected member trend</div>
              <p className="card-subtitle">Biweekly score trend over the full year.</p>
              <ResponsiveContainer width="100%" height={320}>
                <LineChart data={overviewTrendData}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="period" hide />
                  <YAxis domain={[0, 100]} />
                  <Tooltip />
                  <Line type="monotone" dataKey="total" stroke="#10b981" strokeWidth={3} />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </section>
        </>
      )}

      {tab === 'members' && (
        <>
          <FilterToolbar
            selectedPeriod={selectedPeriod}
            setSelectedPeriod={setSelectedPeriod}
            selectedMember={selectedMember}
            setSelectedMember={setSelectedMember}
            search={search}
            setSearch={setSearch}
            members={members}
          />
          <section className="members-stack">
            <div className="card input-panel">
              <div className="card-title"><ClipboardList size={18} /> Member input form</div>
              {effectiveMember ? (
                <>
                  <div className="row-between compact-row">
                    <div className="period-pill">{effectiveMember}</div>
                    <div className="period-pill">{selectedPeriod}</div>
                  </div>
                  <div className="input-grid">
                    {CATEGORY_CONFIG.map((category) => (
                      <div key={category.key} className="input-item">
                        <label>{category.short}</label>
                        <input
                          type="number"
                          value={data[effectiveMember][selectedPeriod][category.key]}
                          onChange={(e) => handleFieldChange(effectiveMember, selectedPeriod, category.key, e.target.value)}
                          placeholder="0"
                        />
                        <div className="input-meta">
                          <span>Team avg: {computed.teamAverages[selectedPeriod][category.key]}</span>
                          <span>Pts: {selectedRecord?.categoryPoints[category.key] || '-'}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                  <div className="score-banner">
                    <div>
                      <div className="banner-label">Total Score</div>
                      <div className="banner-value">{selectedRecord?.totalScore || '—'}</div>
                    </div>
                    <div>
                      <div className="banner-label">Rating</div>
                      <div className={getBadgeClass(selectedRecord?.rating)}>{selectedRecord?.rating || '—'}</div>
                    </div>
                  </div>
                </>
              ) : (
                <div className="empty-inline">Select a specific member to edit input values for the selected period.</div>
              )}
            </div>

            <div className="card leaderboard-panel">
              <div className="panel-header">
                <div className="card-title"><Users size={18} /> Member leaderboard</div>
                <div className="period-pill">{selectedPeriod}</div>
              </div>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Name</th>
                      <th>Score</th>
                      <th>Stories Validated</th>
                      <th>Automation Scenarios</th>
                      <th>Test Cases</th>
                      <th>PR Reviews</th>
                      <th>Collaboration</th>
                      <th>Bugs Raised</th>
                      <th>Bugs Validated</th>
                      <th>Initiatives Taken</th>
                      <th>Training Sessions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {periodData.map((row) => (
                      <tr key={row.name} onClick={() => setSelectedMember(row.name)} className={row.name === effectiveMember ? 'active-row' : ''}>
                        <td>{row.name}</td>
                        <td>{row.score || '—'}</td>
                        <td>{row.raw.storiesValidated || 0}</td>
                        <td>{row.raw.automationScenarios || 0}</td>
                        <td>{row.raw.testCases || 0}</td>
                        <td>{row.raw.prReviews || 0}</td>
                        <td>{row.raw.collaboration || 0}</td>
                        <td>{row.raw.bugsRaised || 0}</td>
                        <td>{row.raw.bugsValidated || 0}</td>
                        <td>{row.raw.initiativesTaken || 0}</td>
                        <td>{row.raw.trainingSessions || 0}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        </>
      )}

      {tab === 'individual' && (
        <>
          <FilterToolbar
            selectedPeriod={selectedPeriod}
            setSelectedPeriod={setSelectedPeriod}
            selectedMember={selectedMember}
            setSelectedMember={setSelectedMember}
            search={search}
            setSearch={setSearch}
            members={members}
          />
          {effectiveMember ? (
            <section className="chart-grid-two">
              <div className="card chart-card">
                  <div className="row-between">
                    <div>
                      <div className="card-title"><UserCircle2 size={18} /> {effectiveMember}</div>
                      <p className="card-subtitle">Detailed view for {selectedPeriod}.</p>
                    </div>
                    <div className={getBadgeClass(selectedRecord?.rating)}>{selectedRecord?.rating || '—'}</div>
                  </div>
                  <ResponsiveContainer width="100%" height={320}>
                    <RadarChart data={individualRadar}>
                      <PolarGrid />
                      <PolarAngleAxis dataKey="category" tick={{ fontSize: 11 }} />
                      <PolarRadiusAxis domain={[0, 20]} />
                      <Tooltip />
                      <Radar dataKey="score" stroke="#5b7fff" fill="#5b7fff" fillOpacity={0.35} />
                    </RadarChart>
                  </ResponsiveContainer>
                  <div className="insight-box">
                    {(() => {
                      const entries = CATEGORY_CONFIG.map((category) => [category.short, selectedRecord?.categoryPoints[category.key] || 0]).sort((a, b) => b[1] - a[1]);
                      return (
                        <>
                          <p><strong>Strongest area:</strong> {entries[0][0]} ({entries[0][1]} pts)</p>
                          <p><strong>Focus area:</strong> {entries[entries.length - 1][0]} ({entries[entries.length - 1][1]} pts)</p>
                          <p><strong>Total score:</strong> {selectedRecord?.totalScore || 0} / 100</p>
                        </>
                      );
                    })()}
                  </div>
              </div>

              <div className="card single-card-section">
                <div className="card-title"><BarChart3 size={18} /> Member vs team averages</div>
                <p className="card-subtitle">Raw values for {effectiveMember} compared with the selected period team averages.</p>
                <ResponsiveContainer width="100%" height={320}>
                  <BarChart data={memberVsTeamData}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="category" hide />
                    <YAxis />
                    <Tooltip />
                    <Legend />
                    <Bar dataKey="member" fill="#2563eb" radius={[6, 6, 0, 0]} />
                    <Bar dataKey="team" fill="#94a3b8" radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
                <div className="metric-list">
                  {memberVsTeamData.map((metric) => (
                    <div key={metric.category} className="metric-row-compact">
                      <div>{metric.category}</div>
                      <div>{metric.member} vs team {metric.team}</div>
                    </div>
                  ))}
                </div>
              </div>
            </section>
          ) : (
            <section className="card single-card-section">
              <div className="empty-inline">Select a specific member to view individual analytics and team comparisons.</div>
            </section>
          )}
        </>
      )}

      {tab === 'rubric' && (
        <section className="chart-grid-two">
          <div className="card">
            <div className="card-title"><Settings2 size={18} /> Rating rubric</div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Category</th>
                    <th>Max</th>
                    <th>Below Avg</th>
                    <th>Average</th>
                    <th>Above Avg</th>
                    <th>Exceeding</th>
                  </tr>
                </thead>
                <tbody>
                  {CATEGORY_CONFIG.map((item) => (
                    <tr key={item.key}>
                      <td>{item.short}</td>
                      <td>{item.max}</td>
                      <td>{item.below}</td>
                      <td>{item.average}</td>
                      <td>{item.above}</td>
                      <td>{item.exceed}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
          <div className="card">
            <div className="card-title"><BookOpen size={18} /> How this app reads your workbook</div>
            <ul className="rule-list">
              <li>Uses the <strong>Summary</strong> logic to apply category scoring thresholds and final ratings.</li>
              <li>Uses <strong>Team Data</strong> averages logic by biweekly period for category comparisons.</li>
              <li>Reads one tab per team member and supports workbook import/export.</li>
              <li>Pulls <strong>raw data</strong> inputs and recalculates <strong>scored points</strong> automatically across the dashboard.</li>
              <li>Search and member filters stay aligned across Overview, Members, and Individual views.</li>
            </ul>
          </div>
        </section>
      )}
    </div>
  );
}

export default App;
