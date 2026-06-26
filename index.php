<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Missions ACOBI</title>
  <style>
    @font-face { font-family: 'Nunito'; font-style: normal; font-weight: 300 700; font-display: swap;
      src: url('fonts/nunito-latin-ext.woff2') format('woff2');
      unicode-range: U+0100-02BA, U+02BD-02C5, U+02C7-02CC, U+02CE-02D7, U+02DD-02FF, U+0304, U+0308, U+0329, U+1D00-1DBF, U+1E00-1E9F, U+1EF2-1EFF, U+2020, U+20A0-20AB, U+20AD-20C0, U+2113, U+2C60-2C7F, U+A720-A7FF; }
    @font-face { font-family: 'Nunito'; font-style: normal; font-weight: 300 700; font-display: swap;
      src: url('fonts/nunito-latin.woff2') format('woff2');
      unicode-range: U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, U+0304, U+0308, U+0329, U+2000-206F, U+20AC, U+2122, U+2191, U+2193, U+2212, U+2215, U+FEFF, U+FFFD; }
  </style>
  <!-- SheetJS pour lire les fichiers Excel -->
  <script src="xlsx.full.min.js"></script>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    :root {
      --bg:        #1a2340;
      --card-bg:   #212b4c;
      --card-alt:  #263155;
      --border:    rgba(255,255,255,0.08);
      --accent:    #5DD8B8;
      --accent2:   #4ac8ff;
      --text:      #ffffff;
      --text-muted: rgba(255,255,255,0.55);
      --danger:    #ff6b6b;
      --success:   #5DD8B8;
      --warning:   #ffd166;
      --radius:    14px;
    }

    body {
      font-family: 'Nunito', sans-serif;
      background-color: var(--bg);
      color: var(--text);
      min-height: 100vh;
    }

    /* ── HEADER ── */
    header {
      padding: 2rem 3rem 0;
    }
    .header-badge {
      font-size: 0.75rem;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: var(--accent);
      font-weight: 700;
      margin-bottom: 0.4rem;
    }
    header h1 {
      font-size: 2.2rem;
      font-weight: 300;
    }
    header h1 span { color: var(--accent); font-weight: 700; }

    /* ── TABS ── */
    .tabs {
      display: flex;
      gap: 0.5rem;
      padding: 2rem 3rem 0;
      border-bottom: 1px solid var(--border);
    }
    .tab-btn {
      padding: 0.7rem 1.6rem;
      border-radius: 10px 10px 0 0;
      border: none;
      background: transparent;
      color: var(--text-muted);
      font-family: 'Nunito', sans-serif;
      font-size: 0.95rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
    }
    .tab-btn:hover { color: var(--text); }
    .tab-btn.active {
      background: var(--card-bg);
      color: var(--accent);
      border: 1px solid var(--border);
      border-bottom: 1px solid var(--card-bg);
    }

    /* ── CONTENT ── */
    .tab-content { display: none; padding: 2.5rem 3rem; }
    .tab-content.active { display: block; }

    /* ── SECTION TITLES ── */
    .section-title {
      font-size: 0.75rem;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: var(--accent);
      font-weight: 700;
      margin-bottom: 1.2rem;
      margin-top: 2rem;
    }
    .section-title:first-child { margin-top: 0; }

    /* ── CARDS GRID ── */
    .grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
      gap: 1.2rem;
    }

    /* ── CARD ── */
    .card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 1.5rem;
      transition: transform 0.2s, border-color 0.2s;
      position: relative;
    }
    .card:hover { transform: translateY(-3px); border-color: rgba(93,216,184,0.3); }

    .card-logo {
      width: 56px;
      height: 56px;
      border-radius: 10px;
      background: var(--card-alt);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1.6rem;
      margin-bottom: 1rem;
      overflow: hidden;
    }
    .card-logo img { width: 100%; height: 100%; object-fit: contain; padding: 4px; }

    .card-title {
      font-size: 1rem;
      font-weight: 700;
      margin-bottom: 0.3rem;
    }
    .card-client {
      font-size: 0.85rem;
      color: var(--accent);
      font-weight: 600;
      margin-bottom: 0.5rem;
    }
    .card-collab {
      font-size: 0.82rem;
      color: var(--text-muted);
      margin-bottom: 0.8rem;
    }
    .card-dates {
      font-size: 0.8rem;
      color: var(--text-muted);
      margin-bottom: 1rem;
    }
    .card-arrow {
      color: var(--accent);
      font-size: 1.1rem;
      margin-top: 0.5rem;
    }

    /* ── BADGE STATUT ── */
    .badge {
      display: inline-block;
      padding: 0.25rem 0.75rem;
      border-radius: 20px;
      font-size: 0.75rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      margin-bottom: 0.8rem;
    }
    .badge-encours  { background: rgba(74,200,255,0.15); color: var(--accent2); }
    .badge-terminee { background: rgba(93,216,184,0.15); color: var(--success); }

    /* ── FILTRES ── */
    .filters {
      display: flex;
      gap: 1rem;
      flex-wrap: wrap;
      margin-bottom: 2rem;
      align-items: flex-end;
    }
    .filter-group { display: flex; flex-direction: column; gap: 0.3rem; }
    .filter-group label {
      font-size: 0.75rem;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--text-muted);
      font-weight: 600;
    }
    .filter-group select, .filter-group input[type=text] {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 8px;
      color: var(--text);
      padding: 0.5rem 0.9rem;
      font-family: 'Nunito', sans-serif;
      font-size: 0.9rem;
      cursor: pointer;
      min-width: 160px;
    }
    .filter-group select:focus, .filter-group input:focus {
      outline: none;
      border-color: var(--accent);
    }

    /* ── FORMULAIRES ── */
    .form-card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 1.8rem;
      margin-bottom: 1.5rem;
    }
    .form-card h3 {
      font-size: 1rem;
      font-weight: 700;
      margin-bottom: 1.2rem;
      color: var(--accent);
    }
    .form-row {
      display: flex;
      gap: 1rem;
      flex-wrap: wrap;
      margin-bottom: 1rem;
    }
    .form-group {
      display: flex;
      flex-direction: column;
      gap: 0.4rem;
      flex: 1;
      min-width: 180px;
    }
    .form-group label {
      font-size: 0.8rem;
      font-weight: 600;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.07em;
    }
    .form-group input,
    .form-group select,
    .form-group textarea {
      background: var(--card-alt);
      border: 1px solid var(--border);
      border-radius: 8px;
      color: var(--text);
      padding: 0.6rem 0.9rem;
      font-family: 'Nunito', sans-serif;
      font-size: 0.9rem;
    }
    .form-group input:focus,
    .form-group select:focus,
    .form-group textarea:focus {
      outline: none;
      border-color: var(--accent);
    }
    .form-group select[multiple] { height: 100px; }

    /* ── LISTE COLLABORATEUR AVEC FILTRE ── */
    .collab-dropdown {
      display: none;
      position: absolute;
      top: 100%;
      left: 0;
      right: 0;
      z-index: 200;
      background: var(--card-bg);
      border: 1px solid var(--accent);
      border-top: none;
      border-radius: 0 0 8px 8px;
      max-height: 200px;
      overflow-y: auto;
    }
    .collab-dd-item {
      padding: 0.5rem 0.9rem;
      cursor: pointer;
      font-size: 0.9rem;
    }
    .collab-dd-item:hover { background: var(--card-alt); }
    .search-clear-btn {
      position: absolute; right: 0.5rem; top: 50%; transform: translateY(-50%);
      background: none; border: none; cursor: pointer; color: var(--text-muted);
      font-size: 1rem; line-height: 1; padding: 0 0.2rem; display: none;
    }
    .search-clear-btn:hover { color: var(--danger); }

    /* ── SOUS-ONGLETS PARAMÉTRAGE ── */
    .subtabs {
      display: flex;
      gap: 0.4rem;
      margin-bottom: 2rem;
      border-bottom: 1px solid var(--border);
      padding-bottom: 0;
    }
    .subtab-btn {
      padding: 0.55rem 1.3rem;
      border-radius: 8px 8px 0 0;
      border: none;
      background: transparent;
      color: var(--text-muted);
      font-family: 'Nunito', sans-serif;
      font-size: 0.9rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
    }
    .subtab-btn:hover { color: var(--text); }
    .subtab-btn.active {
      background: var(--card-alt);
      color: var(--accent);
      border: 1px solid var(--border);
      border-bottom: 1px solid var(--card-alt);
    }
    .subtab-content { display: none; }
    .subtab-content.active { display: block; }

    /* ── BOUTONS ── */
    .btn {
      padding: 0.6rem 1.4rem;
      border-radius: 8px;
      border: none;
      font-family: 'Nunito', sans-serif;
      font-size: 0.9rem;
      font-weight: 700;
      cursor: pointer;
      transition: opacity 0.2s, transform 0.1s;
    }
    .btn:hover { opacity: 0.85; transform: translateY(-1px); }
    .btn-primary { background: var(--accent); color: var(--bg); }
    .btn-secondary { background: var(--card-alt); color: var(--text); border: 1px solid var(--border); }
    .btn-danger { background: rgba(255,107,107,0.15); color: var(--danger); border: 1px solid rgba(255,107,107,0.3); }
    .btn-sm { padding: 0.3rem 0.8rem; font-size: 0.8rem; }

    /* ── TABLES ── */
    .data-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.9rem;
    }
    .data-table th {
      text-align: left;
      padding: 0.7rem 1rem;
      background: var(--card-alt);
      color: var(--text-muted);
      font-size: 0.75rem;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      font-weight: 700;
    }
    .data-table th.sortable { cursor: pointer; user-select: none; }
    .data-table th.sortable:hover { color: var(--accent); }
    .data-table th.sortable::after { content: ' ⇅'; opacity: 0.4; font-size: 0.7rem; }
    .data-table th.sort-asc::after  { content: ' ▲'; opacity: 1; color: var(--accent); }
    .data-table th.sort-desc::after { content: ' ▼'; opacity: 1; color: var(--accent); }
    .data-table th:first-child { border-radius: 8px 0 0 8px; }
    .data-table th:last-child  { border-radius: 0 8px 8px 0; }
    .data-table td {
      padding: 0.7rem 1rem;
      border-bottom: 1px solid var(--border);
      vertical-align: middle;
    }
    .data-table tr:last-child td { border-bottom: none; }
    .data-table .logo-mini {
      width: 32px; height: 32px;
      border-radius: 6px;
      background: var(--card-alt);
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 1rem;
      overflow: hidden;
      vertical-align: middle;
      margin-right: 0.5rem;
    }
    .data-table .logo-mini img { width: 100%; height: 100%; object-fit: contain; }

    /* ── TOGGLE VUE ── */
    .view-toggle {
      display: flex;
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 3px;
      gap: 3px;
    }
    .view-btn {
      padding: 0.45rem 1.2rem;
      border-radius: 8px;
      border: none;
      background: transparent;
      color: var(--text-muted);
      font-family: 'Nunito', sans-serif;
      font-size: 0.88rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.18s;
    }
    .view-btn.active {
      background: var(--accent);
      color: var(--bg);
    }

    /* ── RADIO STATUT ── */
    .radio-group {
      display: flex;
      gap: 0.4rem;
      flex-wrap: wrap;
    }
    .radio-option {
      display: flex;
      align-items: center;
      gap: 0;
      cursor: pointer;
    }
    .radio-option input[type=radio] { display: none; }
    .radio-option span {
      display: inline-block;
      padding: 0.45rem 1rem;
      border-radius: 8px;
      border: 1px solid var(--border);
      background: var(--card-bg);
      color: var(--text-muted);
      font-size: 0.88rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.18s;
      white-space: nowrap;
    }
    .radio-option input[type=radio]:checked + span {
      background: rgba(93,216,184,0.15);
      border-color: var(--accent);
      color: var(--accent);
    }
    .radio-option:hover span { color: var(--text); }

    /* ── CARTE CLIENT (vue groupée) ── */
    .client-card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 1.4rem;
      transition: transform 0.2s, border-color 0.2s;
    }
    .client-card:hover { transform: translateY(-3px); border-color: rgba(93,216,184,0.3); }
    .client-card-logo-box {
      background: #fff;
      border-radius: 10px;
      display: flex;
      align-items: center;
      justify-content: center;
      width: 100%;
      height: 72px;
      margin-bottom: 1rem;
      overflow: hidden;
      padding: 8px;
    }
    .client-card-logo-box img {
      max-width: 100%;
      max-height: 100%;
      object-fit: contain;
    }
    .client-card-logo-initiales {
      font-size: 1.6rem;
      font-weight: 700;
      color: var(--bg);
      letter-spacing: 0.05em;
    }
    .client-collab-list {
      list-style: none;
      margin: 0;
      padding: 0;
      display: flex;
      flex-direction: column;
      gap: 0.35rem;
    }
    .client-collab-list li {
      font-size: 0.80rem;
      color: var(--text);
      line-height: 1.4;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .client-collab-list li .collab-date {
      color: var(--text-muted);
      font-size: 0.78rem;
    }

    /* ── CARTE MISSION (vue collab) ── */
    .mission-card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 1.4rem;
      display: flex;
      flex-direction: column;
      gap: 0.6rem;
      transition: transform 0.2s, border-color 0.2s;
    }
    .mission-card:hover { transform: translateY(-3px); border-color: rgba(93,216,184,0.3); }
    .mission-card-header {
      display: flex;
      align-items: center;
      gap: 0.9rem;
    }
    .mission-card-logo {
      width: 52px;
      height: 52px;
      flex-shrink: 0;
      background: #fff;
      border-radius: 8px;
      display: flex;
      align-items: center;
      justify-content: center;
      overflow: hidden;
      padding: 4px;
    }
    .mission-card-logo img { width: 100%; height: 100%; object-fit: contain; }
    .mission-card-title { font-size: 1rem; font-weight: 700; }
    .mission-card-client { font-size: 0.85rem; color: var(--accent); font-weight: 600; }
    .mission-card-field {
      font-size: 0.85rem;
      color: var(--text-muted);
      display: flex;
      gap: 0.4rem;
      align-items: baseline;
    }
    .mission-card-field strong { color: var(--text); font-weight: 600; }
    .mission-card-details {
      font-size: 0.83rem;
      color: var(--text-muted);
      line-height: 1.5;
      border-top: 1px solid var(--border);
      padding-top: 0.6rem;
      margin-top: 0.2rem;
    }

    /* ── ANALYSE ── */
    .analyse-section { margin-bottom: 2rem; }
    .analyse-section-title {
      font-size: 0.75rem;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: var(--accent);
      font-weight: 700;
      margin-bottom: 1rem;
    }
    .analyse-stat-row { display: flex; gap: 1rem; flex-wrap: wrap; margin-bottom: 2rem; }
    .analyse-stat {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 1rem 1.6rem;
      display: flex;
      flex-direction: column;
      gap: 0.2rem;
      min-width: 140px;
    }
    .analyse-stat .val { font-size: 2rem; font-weight: 700; color: var(--accent); line-height: 1; }
    .analyse-stat .val.blue { color: var(--accent2); }
    .analyse-stat .val.yellow { color: var(--warning); }
    .analyse-stat .lbl { font-size: 0.78rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.07em; }
    .analyse-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem; }
    @media (max-width: 900px) { .analyse-grid { grid-template-columns: 1fr; } }
    .analyse-card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 1.4rem;
    }
    .analyse-card-title {
      font-size: 0.8rem;
      text-transform: uppercase;
      letter-spacing: 0.1em;
      color: var(--accent);
      font-weight: 700;
      margin-bottom: 1rem;
    }
    .analyse-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 0.55rem 0;
      border-bottom: 1px solid var(--border);
      font-size: 0.88rem;
      gap: 0.5rem;
    }
    .analyse-row:last-child { border-bottom: none; }
    .analyse-row .name { font-weight: 600; }
    .analyse-row .meta { color: var(--text-muted); font-size: 0.82rem; text-align: right; }
    .analyse-bar-wrap { display: flex; align-items: center; gap: 0.6rem; min-width: 80px; }
    .analyse-bar-bg { flex: 1; height: 6px; background: var(--card-alt); border-radius: 3px; overflow: hidden; }
    .analyse-bar-fill { height: 100%; border-radius: 3px; background: var(--accent); transition: width 0.4s; }

    /* ── MULTI-SELECT CUSTOM ── */
    .ms-wrapper { position: relative; min-width: 200px; }
    .ms-btn {
      width: 100%;
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 8px;
      color: var(--text);
      padding: 0.5rem 0.9rem;
      font-family: 'Nunito', sans-serif;
      font-size: 0.9rem;
      cursor: pointer;
      text-align: left;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 0.5rem;
    }
    .ms-btn:hover, .ms-btn.open { border-color: var(--accent); }
    .ms-dropdown {
      display: none;
      position: absolute;
      top: calc(100% + 4px);
      left: 0;
      min-width: 100%;
      background: var(--card-bg);
      border: 1px solid var(--accent);
      border-radius: 8px;
      z-index: 200;
      box-shadow: 0 4px 16px rgba(0,0,0,0.35);
      overflow: hidden;
      max-height: 240px;
      overflow-y: auto;
    }
    .ms-dropdown.open { display: block; }
    .ms-item {
      display: flex;
      align-items: center;
      gap: 0.6rem;
      padding: 0.5rem 0.9rem;
      cursor: pointer;
      font-size: 0.88rem;
      transition: background 0.12s;
    }
    .ms-item:hover { background: var(--card-alt); }
    .ms-item.ms-all { border-bottom: 1px solid var(--border); font-weight: 700; color: var(--accent); }
    .ms-item input[type=checkbox] { accent-color: var(--accent); width: 14px; height: 14px; cursor: pointer; }

    /* ── INFOBULLE ── */
    .stat-chip { position: relative; }
    .stat-chip[data-tooltip]:hover::after {
      content: attr(data-tooltip);
      position: absolute;
      top: calc(100% + 8px);
      left: 0;
      transform: none;
      background: #0d1526;
      border: 1px solid var(--border);
      color: var(--text);
      font-size: 0.8rem;
      font-weight: 400;
      padding: 0.5rem 0.85rem;
      border-radius: 8px;
      white-space: pre-line;
      min-width: 320px;
      max-width: 480px;
      z-index: 300;
      box-shadow: 0 4px 16px rgba(0,0,0,0.4);
      line-height: 1.8;
      pointer-events: none;
    }
    .stat-chip[data-tooltip] { cursor: default; }
    .stat-chip[data-tooltip]:hover { border-color: var(--accent); }

    /* ── TAG PICKER ── */
    .tag-picker {
      display: flex;
      flex-wrap: wrap;
      gap: 0.4rem;
      min-height: 2rem;
      align-items: center;
    }
    .tag-pick {
      display: inline-block;
      padding: 0.28rem 0.85rem;
      border-radius: 20px;
      border: 1px solid var(--border);
      background: var(--card-alt);
      color: var(--text-muted);
      font-size: 0.82rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.15s;
      user-select: none;
    }
    .tag-pick:hover { border-color: var(--accent); color: var(--text); }
    .tag-pick.selected {
      background: rgba(93,216,184,0.18);
      border-color: var(--accent);
      color: var(--accent);
    }
    .tag-pick-empty { font-size: 0.82rem; color: var(--text-muted); font-style: italic; }
    .tag-pick-add {
      border-style: dashed;
      color: var(--accent);
      border-color: var(--accent);
      background: transparent;
      padding: 0.28rem 0.7rem;
      font-size: 1rem;
      line-height: 1;
    }
    .tag-pick-add:hover { background: rgba(93,216,184,0.1); }
    .tag-add-form {
      display: inline-flex;
      align-items: center;
      gap: 0.3rem;
    }
    .tag-add-input {
      background: var(--card-alt);
      border: 1px solid var(--accent);
      border-radius: 20px;
      color: var(--text);
      padding: 0.22rem 0.75rem;
      font-family: 'Nunito', sans-serif;
      font-size: 0.82rem;
      outline: none;
      width: 160px;
    }
    .tag-pick-confirm { border-color: var(--accent); color: var(--accent); background: rgba(93,216,184,0.15); }
    .tag-pick-cancel  { border-color: var(--danger);  color: var(--danger);  background: rgba(255,107,107,0.1); }

    /* ── IMPORT EXCEL ── */
    .import-zone {
      border: 2px dashed var(--border);
      border-radius: var(--radius);
      padding: 2rem;
      text-align: center;
      cursor: pointer;
      transition: border-color 0.2s;
      margin-bottom: 1rem;
    }
    .import-zone:hover, .import-zone.drag-over { border-color: var(--accent); }
    .import-zone .icon { font-size: 2rem; margin-bottom: 0.5rem; }
    .import-zone p { color: var(--text-muted); font-size: 0.9rem; }
    .import-zone strong { color: var(--accent); }

    /* ── TOAST ── */
    #toast {
      position: fixed;
      bottom: 2rem;
      right: 2rem;
      background: var(--accent);
      color: var(--bg);
      padding: 0.8rem 1.5rem;
      border-radius: 10px;
      font-weight: 700;
      font-size: 0.9rem;
      opacity: 0;
      transform: translateY(10px);
      transition: all 0.3s;
      z-index: 1000;
      pointer-events: none;
    }
    #toast.show { opacity: 1; transform: translateY(0); }
    #toast.error { background: var(--danger); color: #fff; }

    /* ── EMPTY STATE ── */
    .empty-state {
      text-align: center;
      padding: 4rem 2rem;
      color: var(--text-muted);
    }
    .empty-state .icon { font-size: 3rem; margin-bottom: 1rem; }

    /* ── MODAL ── */
    .modal-overlay {
      display: none;
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.6);
      z-index: 500;
      align-items: center;
      justify-content: center;
    }
    .modal-overlay.open { display: flex; }
    .modal {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 2rem;
      width: 90%;
      max-width: 520px;
      max-height: 90vh;
      overflow-y: auto;
    }
    .modal h3 { color: var(--accent); margin-bottom: 1.2rem; font-size: 1.1rem; }
    .modal-actions { display: flex; gap: 0.8rem; justify-content: flex-end; margin-top: 1.5rem; }

    .stat-row {
      display: flex;
      gap: 1rem;
      margin-bottom: 2rem;
      align-items: flex-start;
      overflow: visible;
    }
    .stat-chip {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 0.8rem 1.4rem;
      display: flex;
      flex-direction: column;
      gap: 0.2rem;
      position: relative;
    }
    .stat-chip .val { font-size: 1.6rem; font-weight: 700; color: var(--accent); }
    .stat-chip .lbl { font-size: 0.78rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.07em; }
  </style>
</head>
<body>

<header>
  <div class="header-badge">ACOBI</div>
  <h1>Suivi des <span>missions</span>.</h1>
</header>

<nav class="tabs">
  <button class="tab-btn active" onclick="showTab('visu')">🗂 Visualisation</button>
  <button class="tab-btn" onclick="showTab('collab-view')">👤 Vue Collab</button>
  <button class="tab-btn" onclick="showTab('analyse')">📊 Analyse</button>
  <button class="tab-btn" onclick="showTab('param')">⚙️ Paramétrage</button>
</nav>

<!-- ════════════════════════════════════════
     ONGLET VISUALISATION
════════════════════════════════════════ -->
<div id="tab-visu" class="tab-content active">

  <div class="stat-row" id="stats-row">
    <!-- rempli par JS -->
  </div>

  <div class="filters" style="align-items:flex-start">
    <div class="filter-group">
      <label>Afficher</label>
      <div class="view-toggle">
        <button class="view-btn active" id="view-btn-client" onclick="setView('client')">Par client</button>
        <button class="view-btn" id="view-btn-collab" onclick="setView('collab')">Par collaborateur</button>
      </div>
    </div>
    <div class="filter-group">
      <label>Statut</label>
      <div class="radio-group">
        <label class="radio-option">
          <input type="radio" name="f-statut" value="en_cours" checked onchange="renderCards()">
          <span>En cours</span>
        </label>
        <label class="radio-option">
          <input type="radio" name="f-statut" value="terminee" onchange="renderCards()">
          <span>Terminées</span>
        </label>
        <label class="radio-option">
          <input type="radio" name="f-statut" value="" onchange="renderCards()">
          <span>Toutes</span>
        </label>
      </div>
    </div>
    <div class="filter-group" id="f-client-group">
      <label>Client</label>
      <select id="f-client" onchange="renderCards()">
        <option value="">Tous</option>
      </select>
    </div>
    <div class="filter-group" id="f-collab-group">
      <label>Collaborateur</label>
      <select id="f-collab" onchange="renderCards()">
        <option value="">Tous</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Recherche</label>
      <input type="text" id="f-search" placeholder="Nom de mission…" oninput="renderCards()">
    </div>
  </div>

  <div class="grid" id="missions-grid">
    <!-- rempli par JS -->
  </div>
</div>

<!-- ════════════════════════════════════════
     ONGLET VUE COLLAB
════════════════════════════════════════ -->
<div id="tab-collab-view" class="tab-content">

  <div class="filters" style="align-items:flex-start">
    <div class="filter-group">
      <label>Collaborateur</label>
      <div style="position:relative;min-width:320px">
        <input type="text" id="cv-collab-search" placeholder="Rechercher un collaborateur…" autocomplete="off"
          oninput="filterCVCollabDD();toggleClearBtn('cv-collab')" onfocus="showCVCollabDD()" onblur="hideCVCollabDD()"
          style="background:var(--card-bg);border:1px solid var(--border);border-radius:8px;color:var(--text);padding:0.5rem 2rem 0.5rem 0.9rem;font-family:'Nunito',sans-serif;font-size:0.9rem;width:100%">
        <button class="search-clear-btn" id="cv-collab-clear" onclick="clearSearchFilter('cv-collab','renderCollabView')" title="Vider le filtre">✕</button>
        <input type="hidden" id="cv-collab-id">
        <div id="cv-collab-dd" class="collab-dropdown"></div>
      </div>
    </div>
    <div class="filter-group">
      <label>Statut</label>
      <div class="radio-group">
        <label class="radio-option">
          <input type="radio" name="cv-statut" value="en_cours" checked onchange="renderCollabView()">
          <span>En cours</span>
        </label>
        <label class="radio-option">
          <input type="radio" name="cv-statut" value="terminee" onchange="renderCollabView()">
          <span>Terminées</span>
        </label>
        <label class="radio-option">
          <input type="radio" name="cv-statut" value="" onchange="renderCollabView()">
          <span>Toutes</span>
        </label>
      </div>
    </div>
  </div>

  <div class="grid" id="cv-missions-grid" style="grid-template-columns:repeat(auto-fill,minmax(300px,1fr))">
    <!-- rempli par JS -->
  </div>

</div>

<!-- ════════════════════════════════════════
     ONGLET ANALYSE
════════════════════════════════════════ -->
<div id="tab-analyse" class="tab-content">

  <div class="filters" style="align-items:flex-start;margin-bottom:1.5rem">

    <!-- Toggle vue -->
    <div class="filter-group">
      <label>Afficher</label>
      <div class="view-toggle">
        <button class="view-btn active" id="an-btn-perimetre" onclick="setAnalyseView('perimetre')">Par Périmètre Métier</button>
        <button class="view-btn" id="an-btn-methode" onclick="setAnalyseView('methode')">Par Méthode / Outil</button>
      </div>
    </div>

    <!-- Filtre Périmètre Métier (vue 1) -->
    <div class="filter-group" id="an-filter-perimetre">
      <label>Périmètre Métier</label>
      <select id="an-perimetre" onchange="renderAnalyse()" style="background:var(--card-bg);border:1px solid var(--border);border-radius:8px;color:var(--text);padding:0.5rem 0.9rem;font-family:'Nunito',sans-serif;font-size:0.9rem;min-width:220px">
        <option value="">— Tous les périmètres —</option>
      </select>
    </div>

    <!-- Filtre Méthodes / Outils (vue 2) -->
    <div class="filter-group" id="an-filter-methode" style="display:none">
      <label>Méthodes / Outils clés <span style="color:var(--text-muted);font-weight:400;font-size:0.78rem">(sélectionner une ou plusieurs)</span></label>
      <div class="tag-picker" id="an-methodes-picker" style="max-width:700px"></div>
    </div>

    <!-- Filtre statut (commun) -->
    <div class="filter-group">
      <label>Statut</label>
      <div class="radio-group">
        <label class="radio-option"><input type="radio" name="an-statut" value="" checked onchange="renderAnalyse()"><span>Toutes</span></label>
        <label class="radio-option"><input type="radio" name="an-statut" value="en_cours" onchange="renderAnalyse()"><span>En cours</span></label>
        <label class="radio-option"><input type="radio" name="an-statut" value="terminee" onchange="renderAnalyse()"><span>Terminées</span></label>
      </div>
    </div>

    <!-- Filtre effectifs (commun) -->
    <div class="filter-group">
      <label>Effectifs</label>
      <div class="radio-group">
        <label class="radio-option"><input type="radio" name="an-effectifs" value="" checked onchange="renderAnalyse()"><span>Tous</span></label>
        <label class="radio-option"><input type="radio" name="an-effectifs" value="actif" onchange="renderAnalyse()"><span>Dans les effectifs à date</span></label>
      </div>
    </div>

  </div>

  <div id="analyse-content">
    <div class="empty-state"><div class="icon">📊</div><p>Sélectionnez un périmètre métier pour afficher l'analyse.</p></div>
  </div>

</div>

<!-- ════════════════════════════════════════
     ONGLET PARAMETRAGE
════════════════════════════════════════ -->
<div id="tab-param" class="tab-content">

  <!-- SOUS-ONGLETS -->
  <div class="subtabs">
    <button class="subtab-btn active" onclick="showSubTab('collab')">👤 Collaborateurs</button>
    <button class="subtab-btn" onclick="showSubTab('clients')">🏢 Clients</button>
    <button class="subtab-btn" onclick="showSubTab('missions')">📋 Missions</button>
    <button class="subtab-btn" onclick="showSubTab('perimetre')">🎯 Périmètre Missions</button>
  </div>

  <!-- SOUS-ONGLET COLLABORATEURS -->
  <div id="subtab-collab" class="subtab-content active">
  <div class="section-title">👤 Collaborateurs</div>
  <div class="form-card">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0.5rem;margin-bottom:0.25rem">
      <h3 id="c-form-title">Ajouter un collaborateur</h3>
      <button class="btn" style="background:var(--card-alt);color:var(--accent);border:1px solid var(--accent);font-size:0.85rem" onclick="document.getElementById('c-import-input').click()">📥 Import en masse</button>
      <input type="file" id="c-import-input" accept=".xlsx,.xls" style="display:none" onchange="importCollaborateurs(this)">
    </div>
    <input type="hidden" id="c-edit-id">
    <div class="form-row">
      <div class="form-group">
        <label>Nom</label>
        <input type="text" id="c-nom" placeholder="ex. Dupont">
      </div>
      <div class="form-group">
        <label>Prénom</label>
        <input type="text" id="c-prenom" placeholder="ex. Marie">
      </div>
      <div class="form-group">
        <label>Sexe</label>
        <select id="c-sexe">
          <option value="">-- Sélectionner --</option>
          <option value="Masculin">Masculin</option>
          <option value="Féminin">Féminin</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Date d'entrée</label>
        <input type="date" id="c-date-entree">
      </div>
      <div class="form-group">
        <label>Date de sortie</label>
        <input type="date" id="c-date-sortie">
      </div>
      <div class="form-group">
        <label>Co-pilote</label>
        <select id="c-copilote">
          <option value="">-- Aucun --</option>
        </select>
      </div>
    </div>
    <div style="display:flex;gap:0.75rem">
      <button class="btn btn-primary" id="c-submit-btn" onclick="saveCollaborateur()">Ajouter</button>
      <button class="btn" id="c-cancel-btn" style="display:none;background:var(--card-alt);color:var(--text)" onclick="cancelEditCollaborateur()">Annuler</button>
    </div>
  </div>

  <div class="filters" style="align-items:flex-start;margin-bottom:1.2rem">
    <div class="filter-group">
      <label>Statut</label>
      <div class="radio-group">
        <label class="radio-option"><input type="radio" name="fc-statut" value="" checked onchange="renderCollaborateurs()"><span>Tous</span></label>
        <label class="radio-option"><input type="radio" name="fc-statut" value="present" onchange="renderCollaborateurs()"><span>Présent à date</span></label>
      </div>
    </div>
    <div class="filter-group">
      <label>Co-pilote</label>
      <div style="position:relative;min-width:260px">
        <input type="text" id="fc-copilote-search" placeholder="Rechercher un co-pilote…" autocomplete="off"
          oninput="filterFCCopiloteDD();toggleClearBtn('fc-copilote')" onfocus="showFCCopiloteDD()" onblur="hideFCCopiloteDD()"
          style="background:var(--card-bg);border:1px solid var(--border);border-radius:8px;color:var(--text);padding:0.5rem 2rem 0.5rem 0.9rem;font-family:'Nunito',sans-serif;font-size:0.9rem;width:100%">
        <button class="search-clear-btn" id="fc-copilote-clear" onclick="clearSearchFilter('fc-copilote','renderCollaborateurs')" title="Vider le filtre">✕</button>
        <input type="hidden" id="fc-copilote-id">
        <div id="fc-copilote-dd" class="collab-dropdown"></div>
      </div>
    </div>
    <div class="filter-group">
      <label>Année d'arrivée</label>
      <div class="ms-wrapper" id="fc-annee-wrapper">
        <button class="ms-btn" id="fc-annee-btn" onclick="toggleMsDropdown('fc-annee')">
          <span id="fc-annee-label">Toutes les années</span>
          <span style="font-size:0.7rem;opacity:0.6">▼</span>
        </button>
        <div class="ms-dropdown" id="fc-annee-dd">
          <label class="ms-item ms-all">
            <input type="checkbox" id="fc-annee-all" onchange="toggleAllAnnees(this)">
            Sélectionner / Désélectionner tout
          </label>
          <div id="fc-annee-options"></div>
        </div>
      </div>
    </div>
  </div>

  <table class="data-table" id="table-collabs" style="width:100%">
    <thead><tr>
      <th class="sortable" onclick="sortTable('collabs','prenom')">Prénom Nom</th>
      <th class="sortable" onclick="sortTable('collabs','sexe')">Sexe</th>
      <th class="sortable" onclick="sortTable('collabs','dateEntree')">Date entrée</th>
      <th class="sortable" onclick="sortTable('collabs','dateSortie')">Date sortie</th>
      <th class="sortable" onclick="sortTable('collabs','copilote')">Co-pilote</th>
      <th>Actions</th>
    </tr></thead>
    <tbody id="tbody-collabs"></tbody>
  </table>

  </div><!-- /subtab-collab -->

  <!-- SOUS-ONGLET CLIENTS -->
  <div id="subtab-clients" class="subtab-content">
  <div class="section-title">🏢 Clients</div>
  <div class="form-card">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0.5rem;margin-bottom:0.25rem">
      <h3>Ajouter un client</h3>
      <button class="btn" style="background:var(--card-alt);color:var(--accent);border:1px solid var(--accent);font-size:0.85rem" onclick="document.getElementById('cl-import-input').click()">📥 Import en masse</button>
      <input type="file" id="cl-import-input" accept=".xlsx,.xls" style="display:none" onchange="importClients(this)">
    </div>
    <input type="hidden" id="cl-logo">
    <div class="form-row" style="align-items:center;gap:1.5rem">
      <div class="form-group" style="flex:1">
        <label>Nom du client</label>
        <input type="text" id="cl-nom" placeholder="ex. EDF">
      </div>
      <div class="form-group" style="flex:0;min-width:auto">
        <label>Logo</label>
        <div id="cl-logo-preview" style="width:56px;height:56px;display:flex;align-items:center;justify-content:center;background:var(--card-alt);border-radius:10px;overflow:hidden"></div>
      </div>
    </div>
    <div style="display:flex;align-items:center;gap:1rem;margin-top:0.5rem;margin-bottom:1rem">
      <button class="btn btn-secondary" type="button" onclick="document.getElementById('cl-logo-file').click()">📁 Charger un logo</button>
      <input type="file" id="cl-logo-file" accept="image/*" style="display:none" onchange="loadLogoFile(this, 'cl-logo', 'cl-logo-preview')">
      <span id="cl-file-label" style="display:none;color:var(--accent);font-size:0.85rem;font-weight:600">✓ Logo prêt</span>
    </div>
    <button class="btn btn-primary" onclick="addClient()">Ajouter</button>
  </div>

  <table class="data-table" id="table-clients">
    <thead><tr>
      <th class="sortable" onclick="sortTable('clients','nom')">Client</th>
      <th>Logo</th>
      <th>Actions</th>
    </tr></thead>
    <tbody></tbody>
  </table>

  </div><!-- /subtab-clients -->

  <!-- SOUS-ONGLET MISSIONS -->
  <div id="subtab-missions" class="subtab-content">
  <div class="section-title">📋 Missions</div>

  <!-- Formulaire mission -->
  <div class="form-card">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0.5rem;margin-bottom:0.25rem">
      <h3>Ajouter une mission</h3>
      <button class="btn" style="background:var(--card-alt);color:var(--accent);border:1px solid var(--accent);font-size:0.85rem" onclick="document.getElementById('m-import-input').click()">📥 Import en masse</button>
      <input type="file" id="m-import-input" accept=".xlsx,.xls" style="display:none" onchange="importMissions(this)">
    </div>
    <div id="import-missions-summary" style="display:none;margin-bottom:1rem"></div>
    <div class="form-row">
      <div class="form-group" style="flex:2">
        <label>Titre de la mission</label>
        <input type="text" id="m-titre" placeholder="ex. Transformation digitale">
      </div>
      <div class="form-group">
        <label>Client *</label>
        <select id="m-client"><option value="">-- Sélectionner --</option></select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="position:relative">
        <label>Collaborateur *</label>
        <input type="text" id="m-collab-search" placeholder="Rechercher..." autocomplete="off"
          oninput="filterCollabDropdown('m')" onfocus="showCollabDropdown('m')" onblur="hideCollabDropdown('m')">
        <input type="hidden" id="m-collabs">
        <div id="m-collab-dd" class="collab-dropdown"></div>
      </div>
      <div class="form-group">
        <label>Date de début *</label>
        <input type="date" id="m-debut">
      </div>
      <div class="form-group">
        <label>Date de fin</label>
        <input type="date" id="m-fin">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="flex:1">
        <label>Périmètre Métier</label>
        <div class="tag-picker" id="m-perimetres"></div>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="flex:1">
        <label>Méthodes / Outils clés</label>
        <div class="tag-picker" id="m-methodes"></div>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="flex:1">
        <label>Détails de la mission</label>
        <textarea id="m-details" rows="3" placeholder="Description, contexte, objectifs…" style="width:100%;background:var(--card-alt);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:0.6rem 0.9rem;font-family:inherit;font-size:0.95rem;resize:vertical"></textarea>
      </div>
    </div>
    <button class="btn btn-primary" onclick="addMission()">Ajouter la mission</button>
  </div>

  <div style="display:flex;gap:1rem;flex-wrap:wrap;margin-bottom:0.8rem;align-items:flex-end">
    <div style="display:flex;flex-direction:column;gap:0.3rem">
      <label style="font-size:0.75rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.07em">Filtrer par client</label>
      <select id="param-f-client" onchange="renderMissions()" style="background:var(--card-alt);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:0.5rem 0.9rem;font-family:inherit;font-size:0.9rem;min-width:180px">
        <option value="">Tous</option>
      </select>
    </div>
    <div style="display:flex;flex-direction:column;gap:0.3rem">
      <label style="font-size:0.75rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.07em">Filtrer par collaborateur</label>
      <select id="param-f-collab" onchange="renderMissions()" style="background:var(--card-alt);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:0.5rem 0.9rem;font-family:inherit;font-size:0.9rem;min-width:200px">
        <option value="">Tous</option>
      </select>
    </div>
    <div style="display:flex;flex-direction:column;gap:0.3rem">
      <label style="font-size:0.75rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.07em">Statut</label>
      <select id="param-f-statut" onchange="renderMissions()" style="background:var(--card-alt);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:0.5rem 0.9rem;font-family:inherit;font-size:0.9rem;min-width:150px">
        <option value="">Tous</option>
        <option value="en_cours">En cours</option>
        <option value="terminee">Terminée</option>
      </select>
    </div>
    <button class="btn" style="background:transparent;border:1px solid var(--border);color:var(--text-muted);padding:0.5rem 1rem;font-size:0.85rem" onclick="document.getElementById('param-f-client').value='';document.getElementById('param-f-collab').value='';document.getElementById('param-f-statut').value='';renderMissions()">✕ Effacer</button>
  </div>

  <table class="data-table" id="table-missions" style="width:100%">
    <thead><tr>
      <th class="sortable" onclick="sortTable('missions','titre')">Mission</th>
      <th class="sortable" onclick="sortTable('missions','client')">Client</th>
      <th class="sortable" onclick="sortTable('missions','collabs')">Collaborateurs</th>
      <th class="sortable" onclick="sortTable('missions','debut')">Période</th>
      <th class="sortable" onclick="sortTable('missions','statut')">Statut</th>
      <th style="min-width:140px">Actions</th>
    </tr></thead>
    <tbody id="tbody-missions"></tbody>
  </table>

  </div><!-- /subtab-missions -->

  <!-- SOUS-ONGLET PÉRIMÈTRE MISSIONS -->
  <div id="subtab-perimetre" class="subtab-content">

    <div style="display:grid;grid-template-columns:1fr 1fr;gap:2rem;align-items:start">

      <!-- PÉRIMÈTRE MÉTIER -->
      <div>
        <div class="section-title">🎯 Périmètre Métier</div>
        <div class="form-card">
          <h3 id="pm-form-title">Ajouter une valeur</h3>
          <input type="hidden" id="pm-edit-id">
          <div class="form-row">
            <div class="form-group" style="flex:1">
              <label>Valeur</label>
              <input type="text" id="pm-nom" placeholder="ex. Finance, RH, Supply Chain…">
            </div>
          </div>
          <div style="display:flex;gap:0.75rem">
            <button class="btn btn-primary" id="pm-submit-btn" onclick="savePerimetre()">Ajouter</button>
            <button class="btn" id="pm-cancel-btn" style="display:none;background:var(--card-alt);color:var(--text)" onclick="cancelEditPerimetre()">Annuler</button>
          </div>
        </div>
        <table class="data-table" style="width:100%">
          <thead><tr>
            <th>Valeur</th>
            <th style="width:140px">Actions</th>
          </tr></thead>
        </table>
        <div style="max-height:260px;overflow-y:auto;border:1px solid var(--border);border-top:none;border-radius:0 0 8px 8px">
          <table class="data-table" style="width:100%"><tbody id="tbody-perimetres"></tbody></table>
        </div>
      </div>

      <!-- MÉTHODES / OUTILS CLÉS -->
      <div>
        <div class="section-title">🔧 Méthodes / Outils clés</div>
        <div class="form-card">
          <h3 id="mo-form-title">Ajouter une valeur</h3>
          <input type="hidden" id="mo-edit-id">
          <div class="form-row">
            <div class="form-group" style="flex:1">
              <label>Valeur</label>
              <input type="text" id="mo-nom" placeholder="ex. Agile, Lean, SAP, Power BI…">
            </div>
          </div>
          <div style="display:flex;gap:0.75rem">
            <button class="btn btn-primary" id="mo-submit-btn" onclick="saveMethode()">Ajouter</button>
            <button class="btn" id="mo-cancel-btn" style="display:none;background:var(--card-alt);color:var(--text)" onclick="cancelEditMethode()">Annuler</button>
          </div>
        </div>
        <table class="data-table" style="width:100%">
          <thead><tr>
            <th>Valeur</th>
            <th style="width:140px">Actions</th>
          </tr></thead>
        </table>
        <div style="max-height:260px;overflow-y:auto;border:1px solid var(--border);border-top:none;border-radius:0 0 8px 8px">
          <table class="data-table" style="width:100%"><tbody id="tbody-methodes"></tbody></table>
        </div>
      </div>

    </div>

  </div><!-- /subtab-perimetre -->

</div>

<!-- TOAST -->
<div id="toast"></div>

<!-- MODAL CONFIRMATION SUPPRESSION -->
<div class="modal-overlay" id="modal-confirm">
  <div class="modal">
    <h3>Confirmer la suppression</h3>
    <p id="modal-msg" style="color:var(--text-muted);line-height:1.6"></p>
    <div class="modal-actions">
      <button class="btn btn-secondary" onclick="closeModal()">Annuler</button>
      <button class="btn btn-danger" id="modal-confirm-btn">Supprimer</button>
    </div>
  </div>
</div>

<!-- MODAL MODIFICATION CLIENT -->
<div class="modal-overlay" id="modal-edit-client">
  <div class="modal">
    <h3>✏️ Modifier le client</h3>
    <input type="hidden" id="edit-cl-id">
    <!-- Champ caché pour stocker la valeur du logo -->
    <input type="hidden" id="edit-cl-logo">
    <div class="form-row" style="align-items:center;gap:1.5rem">
      <div class="form-group" style="flex:1">
        <label>Nom du client</label>
        <input type="text" id="edit-cl-nom" placeholder="ex. EDF">
      </div>
      <div class="form-group" style="flex:0;min-width:auto">
        <label>Logo actuel</label>
        <div id="edit-cl-logo-preview" style="width:56px;height:56px;display:flex;align-items:center;justify-content:center;background:var(--card-alt);border-radius:10px;overflow:hidden"></div>
      </div>
    </div>
    <div style="display:flex;align-items:center;gap:1rem;margin-top:0.5rem">
      <button class="btn btn-secondary" type="button" onclick="document.getElementById('edit-cl-logo-file').click()">📁 Changer le logo</button>
      <input type="file" id="edit-cl-logo-file" accept="image/*" style="display:none" onchange="loadLogoFile(this, 'edit-cl-logo', 'edit-cl-logo-preview')">
      <span id="edit-cl-file-label" style="display:none;color:var(--accent);font-size:0.85rem;font-weight:600">✓ Nouveau logo prêt</span>
    </div>
    <div class="modal-actions">
      <button class="btn btn-secondary" onclick="closeEditClientModal()">Annuler</button>
      <button class="btn btn-primary" onclick="saveEditClient()">Enregistrer</button>
    </div>
  </div>
</div>

<!-- MODAL MODIFICATION MISSION -->
<!-- MODAL MODIFICATION COLLABORATEUR -->
<div class="modal-overlay" id="modal-edit-collab">
  <div class="modal">
    <h3 id="modal-collab-title">✏️ Modifier le collaborateur</h3>
    <input type="hidden" id="mc-id">
    <div class="form-row">
      <div class="form-group">
        <label>Nom</label>
        <input type="text" id="mc-nom">
      </div>
      <div class="form-group">
        <label>Prénom</label>
        <input type="text" id="mc-prenom">
      </div>
      <div class="form-group">
        <label>Sexe</label>
        <select id="mc-sexe">
          <option value="">-- Sélectionner --</option>
          <option value="Masculin">Masculin</option>
          <option value="Féminin">Féminin</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Date d'entrée</label>
        <input type="date" id="mc-date-entree">
      </div>
      <div class="form-group">
        <label>Date de sortie</label>
        <input type="date" id="mc-date-sortie">
      </div>
      <div class="form-group">
        <label>Co-pilote</label>
        <select id="mc-copilote"></select>
      </div>
    </div>
    <div class="modal-actions">
      <button class="btn btn-secondary" onclick="closeEditCollabModal()">Annuler</button>
      <button class="btn btn-primary" onclick="saveEditCollaborateur()">Enregistrer</button>
    </div>
  </div>
</div>

<div class="modal-overlay" id="modal-edit">
  <div class="modal">
    <h3>✏️ Modifier la mission</h3>
    <input type="hidden" id="edit-id">
    <div class="form-row">
      <div class="form-group" style="flex:2">
        <label>Titre de la mission</label>
        <input type="text" id="edit-titre" placeholder="ex. Transformation digitale">
      </div>
      <div class="form-group">
        <label>Client</label>
        <select id="edit-client"><option value="">-- Sélectionner --</option></select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="position:relative">
        <label>Collaborateur</label>
        <input type="text" id="edit-collab-search" placeholder="Rechercher..." autocomplete="off"
          oninput="filterCollabDropdown('edit')" onfocus="showCollabDropdown('edit')" onblur="hideCollabDropdown('edit')">
        <input type="hidden" id="edit-collabs">
        <div id="edit-collab-dd" class="collab-dropdown"></div>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label>Date de début</label>
        <input type="date" id="edit-debut">
      </div>
      <div class="form-group">
        <label>Date de fin</label>
        <input type="date" id="edit-fin">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="flex:1">
        <label>Périmètre Métier</label>
        <div class="tag-picker" id="edit-perimetres"></div>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="flex:1">
        <label>Méthodes / Outils clés</label>
        <div class="tag-picker" id="edit-methodes"></div>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="flex:1">
        <label>Description</label>
        <textarea id="edit-details" rows="3" placeholder="Description, contexte, objectifs…" style="width:100%;background:var(--card-alt);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:0.6rem 0.9rem;font-family:inherit;font-size:0.9rem;resize:vertical"></textarea>
      </div>
    </div>
    <div class="modal-actions">
      <button class="btn btn-secondary" onclick="closeEditModal()">Annuler</button>
      <button class="btn btn-primary" onclick="saveEditMission()">Enregistrer</button>
    </div>
  </div>
</div>

<script>
// ══════════════════════════════════════════
//  ÉTAT DE L'APPLICATION
// ══════════════════════════════════════════
let DB = {
  collaborateurs: [],
  clients: [],
  missions: [],
  perimetres: [],
  methodes: []
};
let deleteCallback = null;

// ── Persistance base de données ──
function save() {
  fetch('api.php', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(DB)
  }).catch(() => toast('Erreur de sauvegarde', 'error'));
}

async function load() {
  try {
    const res = await fetch('api.php');
    const data = await res.json();
    DB = { collaborateurs: data.collaborateurs || [], missions: data.missions || [], clients: data.clients || [], perimetres: data.perimetres || [], methodes: data.methodes || [] };
  } catch(e) {
    DB = { collaborateurs: [], missions: [], clients: [] };
  }
}

// ── Tri des tableaux ──
const sortState = {
  collabs:  { col: null, dir: 1 },
  clients:  { col: null, dir: 1 },
  missions: { col: null, dir: 1 }
};
function sortTable(table, col) {
  if (sortState[table].col === col) sortState[table].dir *= -1;
  else { sortState[table].col = col; sortState[table].dir = 1; }
  if (table === 'collabs')  renderCollaborateurs();
  if (table === 'clients')  renderClients();
  if (table === 'missions') renderMissions();
}
function updateSortHeaders(tableId, table) {
  document.querySelectorAll('#' + tableId + ' th.sortable').forEach(th => {
    th.classList.remove('sort-asc', 'sort-desc');
    if (th.getAttribute('onclick').includes("'" + sortState[table].col + "'")) {
      th.classList.add(sortState[table].dir === 1 ? 'sort-asc' : 'sort-desc');
    }
  });
}
function applySortCollabs(list) {
  const { col, dir } = sortState.collabs;
  if (!col) return list;
  return [...list].sort((a, b) => {
    let va = col === 'prenom' ? (a.prenom + ' ' + a.nom) : (a[col] || '');
    let vb = col === 'prenom' ? (b.prenom + ' ' + b.nom) : (b[col] || '');
    if (col === 'copilote') {
      const ca = DB.collaborateurs.find(x => x.id === a.copilote);
      const cb = DB.collaborateurs.find(x => x.id === b.copilote);
      va = ca ? ca.prenom + ' ' + ca.nom : '';
      vb = cb ? cb.prenom + ' ' + cb.nom : '';
    }
    return va.localeCompare(vb, 'fr') * dir;
  });
}
function applySortClients(list) {
  const { col, dir } = sortState.clients;
  if (!col) return list;
  return [...list].sort((a, b) => (a[col] || '').localeCompare(b[col] || '', 'fr') * dir);
}
function applySortMissions(list) {
  const { col, dir } = sortState.missions;
  if (!col) return list;
  return [...list].sort((a, b) => {
    let va = '', vb = '';
    if (col === 'titre')  { va = a.titre || ''; vb = b.titre || ''; }
    if (col === 'client') {
      const ca = DB.clients.find(x => x.id === a.clientId);
      const cb = DB.clients.find(x => x.id === b.clientId);
      va = ca ? ca.nom : ''; vb = cb ? cb.nom : '';
    }
    if (col === 'collabs') {
      va = (a.collabIds || []).map(id => { const c = DB.collaborateurs.find(x => x.id === id); return c ? c.prenom + ' ' + c.nom : ''; }).join(', ');
      vb = (b.collabIds || []).map(id => { const c = DB.collaborateurs.find(x => x.id === id); return c ? c.prenom + ' ' + c.nom : ''; }).join(', ');
    }
    if (col === 'debut')  { va = a.debut || ''; vb = b.debut || ''; }
    if (col === 'statut') { va = getStatut(a); vb = getStatut(b); }
    return va.localeCompare(vb, 'fr') * dir;
  });
}

// ── ID unique ──
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

// ══════════════════════════════════════════
//  NAVIGATION ONGLETS
// ══════════════════════════════════════════
function showSubTab(name) {
  document.querySelectorAll('.subtab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.subtab-btn').forEach(el => el.classList.remove('active'));
  document.getElementById('subtab-' + name).classList.add('active');
  document.querySelectorAll('.subtab-btn').forEach(btn => {
    if (btn.getAttribute('onclick').includes("'" + name + "'")) btn.classList.add('active');
  });
}

function showTab(name) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  event.target.classList.add('active');
  if (name === 'visu') renderAll();
  if (name === 'collab-view') { refreshSelects(); renderCollabView(); }
  if (name === 'analyse') { refreshAnalyseSelect(); if (analyseView === 'methode') refreshAnalyseMethodesPicker(); renderAnalyse(); }
}

// ══════════════════════════════════════════
//  TOAST
// ══════════════════════════════════════════
function toast(msg, type = 'ok') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'show' + (type === 'error' ? ' error' : '');
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.className = '', 2800);
}

// ══════════════════════════════════════════
//  MODAL CONFIRMATION
// ══════════════════════════════════════════
function openModal(msg, cb) {
  document.getElementById('modal-msg').textContent = msg;
  document.getElementById('modal-confirm').classList.add('open');
  deleteCallback = cb;
}
function closeModal() {
  document.getElementById('modal-confirm').classList.remove('open');
  deleteCallback = null;
}
document.getElementById('modal-confirm-btn').onclick = () => {
  if (deleteCallback) deleteCallback();
  closeModal();
};

// ══════════════════════════════════════════
//  COLLABORATEURS
// ══════════════════════════════════════════
function saveCollaborateur() {
  const editId = document.getElementById('c-edit-id').value;
  const nom         = document.getElementById('c-nom').value.trim();
  const prenom      = document.getElementById('c-prenom').value.trim();
  const sexe        = document.getElementById('c-sexe').value;
  const dateEntree  = document.getElementById('c-date-entree').value;
  const dateSortie  = document.getElementById('c-date-sortie').value;
  const copilote    = document.getElementById('c-copilote').value;
  if (!nom || !prenom) { toast('Nom et prénom requis', 'error'); return; }
  if (editId) {
    const c = DB.collaborateurs.find(x => x.id === editId);
    Object.assign(c, { nom, prenom, sexe, dateEntree, dateSortie, copilote });
    toast('Collaborateur modifié ✓');
  } else {
    DB.collaborateurs.push({ id: uid(), nom, prenom, sexe, dateEntree, dateSortie, copilote });
    toast('Collaborateur ajouté ✓');
  }
  save();
  resetCollaborateurForm();
  renderCollaborateurs();
  refreshSelects();
}

function resetCollaborateurForm() {
  document.getElementById('c-edit-id').value = '';
  document.getElementById('c-nom').value = '';
  document.getElementById('c-prenom').value = '';
  document.getElementById('c-sexe').value = '';
  document.getElementById('c-date-entree').value = '';
  document.getElementById('c-date-sortie').value = '';
  document.getElementById('c-copilote').value = '';
  document.getElementById('c-form-title').textContent = 'Ajouter un collaborateur';
  document.getElementById('c-submit-btn').textContent = 'Ajouter';
  document.getElementById('c-cancel-btn').style.display = 'none';
}

function editCollaborateur(id) {
  const c = DB.collaborateurs.find(x => x.id === id);
  if (!c) return;
  document.getElementById('mc-id').value          = id;
  document.getElementById('mc-nom').value         = c.nom;
  document.getElementById('mc-prenom').value      = c.prenom;
  document.getElementById('mc-sexe').value        = c.sexe || '';
  document.getElementById('mc-date-entree').value = c.dateEntree || '';
  document.getElementById('mc-date-sortie').value = c.dateSortie || '';
  document.getElementById('modal-collab-title').textContent = `✏️ Modifier ${c.prenom} ${c.nom}`;
  // Remplir le select co-pilote en excluant le collaborateur lui-même
  const today = new Date().toISOString().split('T')[0];
  const mcCopilote = document.getElementById('mc-copilote');
  mcCopilote.innerHTML = '<option value="">-- Aucun --</option>' +
    DB.collaborateurs
      .filter(x => x.id !== id)
      .filter(x => !x.dateSortie || x.dateSortie >= today)
      .sort((a, b) => a.prenom.localeCompare(b.prenom, 'fr'))
      .map(x => `<option value="${x.id}" ${x.id === c.copilote ? 'selected' : ''}>${x.prenom} ${x.nom}</option>`)
      .join('');
  document.getElementById('modal-edit-collab').classList.add('open');
}

function closeEditCollabModal() {
  document.getElementById('modal-edit-collab').classList.remove('open');
}

function saveEditCollaborateur() {
  const id         = document.getElementById('mc-id').value;
  const nom        = document.getElementById('mc-nom').value.trim();
  const prenom     = document.getElementById('mc-prenom').value.trim();
  const sexe       = document.getElementById('mc-sexe').value;
  const dateEntree = document.getElementById('mc-date-entree').value;
  const dateSortie = document.getElementById('mc-date-sortie').value;
  const copilote   = document.getElementById('mc-copilote').value;
  if (!nom || !prenom) { toast('Nom et prénom requis', 'error'); return; }
  const c = DB.collaborateurs.find(x => x.id === id);
  if (!c) return;
  Object.assign(c, { nom, prenom, sexe, dateEntree, dateSortie, copilote });
  save();
  closeEditCollabModal();
  renderCollaborateurs();
  refreshSelects();
  toast('Collaborateur modifié ✓');
}

function cancelEditCollaborateur() {
  resetCollaborateurForm();
  refreshCopiloteSelect(null);
}

function refreshCopiloteSelect(excludeId) {
  const sel = document.getElementById('c-copilote');
  const prev = sel.value;
  sel.innerHTML = '<option value="">-- Aucun --</option>' +
    DB.collaborateurs
      .filter(c => c.id !== excludeId)
      .filter(c => !c.dateSortie || c.dateSortie >= new Date().toISOString().split('T')[0])
      .sort((a, b) => a.prenom.localeCompare(b.prenom, 'fr'))
      .map(c => `<option value="${c.id}">${c.prenom} ${c.nom}</option>`)
      .join('');
  sel.value = prev;
}

function deleteCollaborateur(id) {
  const c = DB.collaborateurs.find(x => x.id === id);
  openModal(`Supprimer ${c.prenom} ${c.nom} ?`, () => {
    DB.collaborateurs = DB.collaborateurs.filter(x => x.id !== id);
    save(); renderCollaborateurs(); refreshSelects();
    toast('Collaborateur supprimé');
  });
}

// ── Multi-select années ──
function toggleMsDropdown(id) {
  const dd  = document.getElementById(id + '-dd');
  const btn = document.getElementById(id + '-btn');
  const isOpen = dd.classList.contains('open');
  document.querySelectorAll('.ms-dropdown.open').forEach(el => el.classList.remove('open'));
  document.querySelectorAll('.ms-btn.open').forEach(el => el.classList.remove('open'));
  if (!isOpen) { dd.classList.add('open'); btn.classList.add('open'); }
}
document.addEventListener('click', e => {
  if (!e.target.closest('.ms-wrapper')) {
    document.querySelectorAll('.ms-dropdown.open').forEach(el => el.classList.remove('open'));
    document.querySelectorAll('.ms-btn.open').forEach(el => el.classList.remove('open'));
  }
});
function toggleAllAnnees(el) {
  document.querySelectorAll('#fc-annee-options input[type=checkbox]').forEach(cb => cb.checked = el.checked);
  updateAnneesLabel();
  renderCollaborateurs();
}
function updateAnneesLabel() {
  const all  = [...document.querySelectorAll('#fc-annee-options input[type=checkbox]')];
  const sel  = all.filter(cb => cb.checked);
  const allEl = document.getElementById('fc-annee-all');
  allEl.checked = sel.length === all.length;
  allEl.indeterminate = sel.length > 0 && sel.length < all.length;
  document.getElementById('fc-annee-label').textContent = sel.length === 0 || sel.length === all.length
    ? 'Toutes les années'
    : sel.map(cb => cb.value).join(', ');
}
function refreshAnneesOptions() {
  const annees = [...new Set(DB.collaborateurs.map(c => (c.dateEntree||'').slice(0,4)).filter(Boolean))].sort();
  const container = document.getElementById('fc-annee-options');
  if (!container) return;
  const prev = [...document.querySelectorAll('#fc-annee-options input:checked')].map(cb => cb.value);
  container.innerHTML = annees.map(a => `
    <label class="ms-item">
      <input type="checkbox" value="${a}" ${prev.length === 0 || prev.includes(a) ? 'checked' : ''} onchange="updateAnneesLabel();renderCollaborateurs()">
      ${a}
    </label>`).join('');
  updateAnneesLabel();
}
function getSelectedAnnees() {
  const all = [...document.querySelectorAll('#fc-annee-options input[type=checkbox]')];
  const sel = all.filter(cb => cb.checked).map(cb => cb.value);
  return sel.length === all.length ? [] : sel; // [] = tout sélectionné = pas de filtre
}

function filterFCCopiloteDD() {
  const search = (document.getElementById('fc-copilote-search').value || '').toLowerCase();
  const dd = document.getElementById('fc-copilote-dd');
  const copiloteIds = new Set(DB.collaborateurs.map(c => c.copilote).filter(Boolean));
  const matches = DB.collaborateurs
    .filter(c => copiloteIds.has(c.id))
    .filter(c => !search || c.prenom.toLowerCase().includes(search) || c.nom.toLowerCase().includes(search) || (c.prenom + ' ' + c.nom).toLowerCase().includes(search))
    .sort((a, b) => a.prenom.localeCompare(b.prenom, 'fr'));
  dd.innerHTML = matches.length
    ? matches.map(c => `<div class="collab-dd-item" onmousedown="selectFCCopilote('${c.id}','${c.prenom} ${c.nom}')">${c.prenom} ${c.nom}</div>`).join('')
    : '<div style="padding:0.5rem 0.9rem;color:var(--text-muted);font-size:0.85rem">Aucun résultat</div>';
  dd.style.display = 'block';
}
function showFCCopiloteDD() { filterFCCopiloteDD(); }
function hideFCCopiloteDD() {
  setTimeout(() => { const dd = document.getElementById('fc-copilote-dd'); if (dd) dd.style.display = 'none'; }, 200);
}
function selectFCCopilote(id, label) {
  document.getElementById('fc-copilote-id').value = id;
  document.getElementById('fc-copilote-search').value = label;
  document.getElementById('fc-copilote-dd').style.display = 'none';
  toggleClearBtn('fc-copilote');
  renderCollaborateurs();
}

function renderCollaborateurs() {
  const tbody = document.getElementById('tbody-collabs');
  updateSortHeaders('table-collabs', 'collabs');
  const today = new Date().toISOString().split('T')[0];
  const fStatut = document.querySelector('input[name="fc-statut"]:checked');
  const statutVal = fStatut ? fStatut.value : '';
  const anneesFilter = getSelectedAnnees();
  const copiloteFilter = document.getElementById('fc-copilote-id').value;

  let allCollabs = DB.collaborateurs.filter(c => {
    if (statutVal === 'present' && c.dateSortie && c.dateSortie < today) return false;
    if (anneesFilter.length && !anneesFilter.includes((c.dateEntree||'').slice(0,4))) return false;
    if (copiloteFilter && c.copilote !== copiloteFilter) return false;
    return true;
  });
  const list = applySortCollabs(allCollabs);
  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="6" style="color:var(--text-muted);text-align:center;padding:1.5rem">Aucun collaborateur</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(c => {
    const copilote = c.copilote ? DB.collaborateurs.find(x => x.id === c.copilote) : null;
    const copiloteLabel = copilote ? `${copilote.prenom} ${copilote.nom}` : '—';
    const fmt = d => d ? d.split('-').reverse().join('/') : '—';
    return `
    <tr>
      <td>${c.prenom} ${c.nom}</td>
      <td style="color:var(--text-muted)">${c.sexe || '—'}</td>
      <td style="color:var(--text-muted)">${fmt(c.dateEntree)}</td>
      <td style="color:var(--text-muted)">${fmt(c.dateSortie)}</td>
      <td style="color:var(--text-muted)">${copiloteLabel}</td>
      <td style="display:flex;gap:0.5rem">
        <button class="btn btn-primary btn-sm" onclick="editCollaborateur('${c.id}')">Modifier</button>
        <button class="btn btn-danger btn-sm" onclick="deleteCollaborateur('${c.id}')">Supprimer</button>
      </td>
    </tr>`;
  }).join('');
}

function importCollaborateurs(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets['Feuil1'] || wb.Sheets[wb.SheetNames[0]];
      if (!ws) { toast('Onglet "Feuil1" introuvable dans le fichier', 'error'); return; }
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
      if (!rows.length) { toast('Fichier vide ou non reconnu', 'error'); return; }

      // Première passe : créer tous les collaborateurs sans co-pilote
      const newIds = {};
      let added = 0;
      rows.forEach(row => {
        const nom    = (row['Nom']    || '').toString().trim();
        const prenom = (row['Prénom'] || row['Prenom'] || '').toString().trim();
        if (!nom || !prenom) return;
        const sexe       = (row['Sexe'] || '').toString().trim();
        const dateEntree = excelDateToISO(row['Entrée'] || row['Entree'] || row['Date entrée'] || '');
        const dateSortie = excelDateToISO(row['Sortie'] || row['Date sortie'] || '');
        const key = (prenom + ' ' + nom).toLowerCase();
        const existing = DB.collaborateurs.find(c => (c.prenom + ' ' + c.nom).toLowerCase() === key);
        if (existing) {
          Object.assign(existing, { sexe, dateEntree, dateSortie });
          newIds[key] = existing.id;
        } else {
          const id = uid();
          DB.collaborateurs.push({ id, nom, prenom, sexe, dateEntree, dateSortie, copilote: '' });
          newIds[key] = id;
          added++;
        }
      });

      // Deuxième passe : associer les co-pilotes
      rows.forEach(row => {
        const nom    = (row['Nom']    || '').toString().trim();
        const prenom = (row['Prénom'] || row['Prenom'] || '').toString().trim();
        if (!nom || !prenom) return;
        const copiloteStr = (row['Co-pilote'] || row['Copilote'] || '').toString().trim();
        if (!copiloteStr || copiloteStr.toLowerCase() === 'n/a') return;
        const key = (prenom + ' ' + nom).toLowerCase();
        const collab = DB.collaborateurs.find(c => c.id === newIds[key]);
        if (!collab) return;
        const coKey = copiloteStr.toLowerCase();
        const co = DB.collaborateurs.find(c => (c.prenom + ' ' + c.nom).toLowerCase() === coKey);
        if (co) collab.copilote = co.id;
      });

      save();
      refreshAnneesOptions();
      renderCollaborateurs();
      refreshSelects();
      toast(`${added} collaborateur(s) importé(s) ✓`);
    } catch(err) {
      toast('Erreur : ' + err.message, 'error');
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

function excelDateToISO(val) {
  if (!val && val !== 0) return '';
  const s = val.toString().trim();
  if (!s) return '';
  // Format DD/MM/YYYY ou DD/MM/YY
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    const y = m[3].length === 2 ? '20' + m[3] : m[3];
    return `${y}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  }
  // Format YYYY-MM-DD déjà correct
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return '';
}

// ══════════════════════════════════════════
//  CLIENTS
// ══════════════════════════════════════════
function isUrl(s) { return s.startsWith('http') || s.startsWith('data:'); }

function logoHtml(logo, size = 40, clientNom = '') {
  if (!logo) return initialesHtml(clientNom, size);
  if (isUrl(logo)) return `<img src="${logo}" style="width:${size}px;height:${size}px;object-fit:contain;border-radius:6px" onerror="this.replaceWith(initialesNode('${clientNom.replace(/'/g,"\\'")}', ${size}))">`;
  return `<span style="font-size:${size * 0.5}px">${logo}</span>`;
}

function loadLogoFile(input, logoFieldId, previewId) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    const dataUrl = e.target.result;
    document.getElementById(logoFieldId).value = dataUrl;
    document.getElementById(previewId).innerHTML =
      `<img src="${dataUrl}" style="width:100%;height:100%;object-fit:contain;padding:4px">`;
    // Afficher le label de confirmation (cl- ou edit-cl-)
    const prefix = logoFieldId.startsWith('edit') ? 'edit-cl' : 'cl';
    const lbl = document.getElementById(prefix + '-file-label');
    if (lbl) lbl.style.display = 'inline';
  };
  reader.readAsDataURL(file);
}

function initialesHtml(nom, size = 40) {
  const initiales = (nom || '?').split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
  return `<span style="display:inline-flex;align-items:center;justify-content:center;width:${size}px;height:${size}px;border-radius:8px;background:var(--card-alt);color:var(--accent);font-weight:700;font-size:${size * 0.35}px">${initiales}</span>`;
}

function initialesNode(nom, size = 40) {
  const el = document.createElement('span');
  const initiales = (nom || '?').split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
  el.style.cssText = `display:inline-flex;align-items:center;justify-content:center;width:${size}px;height:${size}px;border-radius:8px;background:var(--card-alt);color:var(--accent);font-weight:700;font-size:${size * 0.35}px`;
  el.textContent = initiales;
  return el;
}

function addClient() {
  const nom  = document.getElementById('cl-nom').value.trim();
  const logo = document.getElementById('cl-logo').value.trim();
  if (!nom) { toast('Nom du client requis', 'error'); return; }
  DB.clients.push({ id: uid(), nom, logo });
  save();
  document.getElementById('cl-nom').value = '';
  document.getElementById('cl-logo').value = '';
  document.getElementById('cl-logo-file').value = '';
  document.getElementById('cl-file-label').style.display = 'none';
  document.getElementById('cl-logo-preview').innerHTML = '';
  renderClients();
  refreshSelects();
  toast('Client ajouté ✓');
}

function openEditClient(id) {
  const c = DB.clients.find(x => x.id === id);
  if (!c) return;
  document.getElementById('edit-cl-id').value   = c.id;
  document.getElementById('edit-cl-nom').value  = c.nom;
  document.getElementById('edit-cl-logo').value = c.logo || '';
  document.getElementById('edit-cl-file-label').style.display = 'none';
  document.getElementById('edit-cl-logo-file').value = '';
  // Aperçu logo existant
  const prev = document.getElementById('edit-cl-logo-preview');
  if (c.logo) {
    prev.innerHTML = `<img src="${c.logo}" style="width:100%;height:100%;object-fit:contain;padding:4px" onerror="this.parentElement.innerHTML=initialesHtml('${c.nom.replace(/'/g,"\\'")}',56)">`;
  } else {
    prev.innerHTML = initialesHtml(c.nom, 56);
  }
  document.getElementById('modal-edit-client').classList.add('open');
}

function closeEditClientModal() {
  document.getElementById('modal-edit-client').classList.remove('open');
}

function saveEditClient() {
  const id   = document.getElementById('edit-cl-id').value;
  const nom  = document.getElementById('edit-cl-nom').value.trim();
  const logo = document.getElementById('edit-cl-logo').value.trim();
  if (!nom) { toast('Nom du client requis', 'error'); return; }
  const idx = DB.clients.findIndex(x => x.id === id);
  if (idx === -1) return;
  DB.clients[idx] = { id, nom, logo };
  save();
  closeEditClientModal();
  renderClients();
  refreshSelects();
  renderCards();
  toast('Client modifié ✓');
}

function deleteClient(id) {
  const c = DB.clients.find(x => x.id === id);
  openModal(`Supprimer le client "${c.nom}" ?`, () => {
    DB.clients = DB.clients.filter(x => x.id !== id);
    save(); renderClients(); refreshSelects();
    toast('Client supprimé');
  });
}

function renderClients() {
  const tbody = document.querySelector('#table-clients tbody');
  updateSortHeaders('table-clients', 'clients');
  const list = applySortClients(DB.clients);
  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="3" style="color:var(--text-muted);text-align:center;padding:1.5rem">Aucun client</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(c => `
    <tr>
      <td>${c.nom}</td>
      <td><span class="logo-mini">${logoHtml(c.logo, 28, c.nom)}</span></td>
      <td style="display:flex;gap:0.5rem">
        <button class="btn btn-primary btn-sm" onclick="openEditClient('${c.id}')">Modifier</button>
        <button class="btn btn-danger btn-sm" onclick="deleteClient('${c.id}')">Supprimer</button>
      </td>
    </tr>
  `).join('');
}

// ══════════════════════════════════════════
//  SELECTS DYNAMIQUES
// ══════════════════════════════════════════
function refreshSelects() {
  // Select client (formulaire mission)
  const mClient = document.getElementById('m-client');
  const prev = mClient.value;
  mClient.innerHTML = '<option value="">-- Sélectionner --</option>' +
    DB.clients.map(c => `<option value="${c.id}">${c.nom}</option>`).join('');
  mClient.value = prev;

  // Select co-pilote (formulaire collaborateur)
  const editId = document.getElementById('c-edit-id').value;
  refreshCopiloteSelect(editId || null);

  // Filtres visualisation
  const fClient = document.getElementById('f-client');
  const fClient2 = fClient.value;
  fClient.innerHTML = '<option value="">Tous</option>' +
    DB.clients.map(c => `<option value="${c.id}">${c.nom}</option>`).join('');
  fClient.value = fClient2;

  const fCollab = document.getElementById('f-collab');
  const fCollab2 = fCollab.value;
  fCollab.innerHTML = '<option value="">Tous</option>' +
    DB.collaborateurs.map(c => `<option value="${c.id}">${c.prenom} ${c.nom}</option>`).join('');
  fCollab.value = fCollab2;

  // Tag pickers du formulaire d'ajout de mission
  renderTagPicker('m-perimetres', DB.perimetres, [], 'perimetre');
  renderTagPicker('m-methodes',   DB.methodes,   [], 'methode');
}

function importClients(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets['Feuil1'] || wb.Sheets[wb.SheetNames[0]];
      if (!ws) { toast('Onglet "Feuil1" introuvable', 'error'); return; }
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
      if (!rows.length) { toast('Fichier vide ou non reconnu', 'error'); return; }
      let added = 0;
      rows.forEach(row => {
        const nom = (row['Client'] || '').toString().trim();
        if (!nom) return;
        const exists = DB.clients.find(c => c.nom.toLowerCase() === nom.toLowerCase());
        if (!exists) {
          DB.clients.push({ id: uid(), nom, logo: '' });
          added++;
        }
      });
      save();
      renderClients();
      refreshSelects();
      toast(`${added} client(s) importé(s) ✓`);
    } catch(err) {
      toast('Erreur : ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ══════════════════════════════════════════
//  MISSIONS
// ══════════════════════════════════════════
function getStatut(m) {
  const today = new Date().toISOString().split('T')[0];
  return (m.fin && m.fin < today) ? 'terminee' : 'en_cours';
}

function addMission(data = null) {
  let titre, clientId, collabIds, debut, fin, details, perimetreIds, methodeIds;
  if (data) {
    ({ titre, clientId, collabIds, debut, fin, details } = data);
    perimetreIds = data.perimetreIds || [];
    methodeIds   = data.methodeIds   || [];
  } else {
    titre        = document.getElementById('m-titre').value.trim();
    clientId     = document.getElementById('m-client').value;
    const collabId = document.getElementById('m-collabs').value;
    collabIds    = collabId ? [collabId] : [];
    debut        = document.getElementById('m-debut').value;
    fin          = document.getElementById('m-fin').value;
    details      = document.getElementById('m-details').value.trim();
    perimetreIds = getSelectedTags('m-perimetres');
    methodeIds   = getSelectedTags('m-methodes');
    if (!clientId) { toast('Client requis', 'error'); return; }
    if (!collabId) { toast('Collaborateur requis', 'error'); return; }
    if (!debut)    { toast('Date de début requise', 'error'); return; }
  }

  DB.missions.push({ id: uid(), titre, clientId, collabIds, debut, fin: fin || '', details: details || '', perimetreIds: perimetreIds || [], methodeIds: methodeIds || [] });
  save();

  if (!data) {
    document.getElementById('m-titre').value = '';
    document.getElementById('m-client').value = '';
    document.getElementById('m-collab-search').value = '';
    document.getElementById('m-collabs').value = '';
    document.getElementById('m-debut').value = '';
    document.getElementById('m-fin').value = '';
    document.getElementById('m-details').value = '';
    renderTagPicker('m-perimetres', DB.perimetres, [], 'perimetre');
    renderTagPicker('m-methodes',   DB.methodes,   [], 'methode');
  }
  renderMissions();
  toast('Mission ajoutée ✓');
}

function openEditMission(id) {
  const m = DB.missions.find(x => x.id === id);
  if (!m) return;

  document.getElementById('edit-id').value      = m.id;
  document.getElementById('edit-titre').value   = m.titre;
  document.getElementById('edit-debut').value   = m.debut   || '';
  document.getElementById('edit-fin').value     = m.fin     || '';
  document.getElementById('edit-details').value = m.details || '';

  // Remplir le select client
  const editClient = document.getElementById('edit-client');
  editClient.innerHTML = '<option value="">-- Sélectionner --</option>' +
    DB.clients.map(c => `<option value="${c.id}" ${c.id === m.clientId ? 'selected' : ''}>${c.nom}</option>`).join('');

  // Pré-remplir le collaborateur
  const firstCollabId = (m.collabIds || [])[0] || '';
  const firstCollab = firstCollabId ? DB.collaborateurs.find(x => x.id === firstCollabId) : null;
  document.getElementById('edit-collabs').value = firstCollabId;
  document.getElementById('edit-collab-search').value = firstCollab ? firstCollab.prenom + ' ' + firstCollab.nom : '';

  // Pré-remplir les tags
  renderTagPicker('edit-perimetres', DB.perimetres, m.perimetreIds || [], 'perimetre');
  renderTagPicker('edit-methodes',   DB.methodes,   m.methodeIds   || [], 'methode');

  document.getElementById('modal-edit').classList.add('open');
}

function closeEditModal() {
  document.getElementById('modal-edit').classList.remove('open');
}

function saveEditMission() {
  const id       = document.getElementById('edit-id').value;
  const titre    = document.getElementById('edit-titre').value.trim();
  const clientId = document.getElementById('edit-client').value;
  const collabId = document.getElementById('edit-collabs').value;
  const collabIds = collabId ? [collabId] : [];
  const debut    = document.getElementById('edit-debut').value;
  const fin      = document.getElementById('edit-fin').value;
  const details  = document.getElementById('edit-details').value.trim();

  const perimetreIds = getSelectedTags('edit-perimetres');
  const methodeIds   = getSelectedTags('edit-methodes');

  const idx = DB.missions.findIndex(x => x.id === id);
  if (idx === -1) return;
  DB.missions[idx] = { id, titre, clientId, collabIds, debut, fin: fin || '', details, perimetreIds, methodeIds };
  save();
  closeEditModal();
  renderMissions();
  renderCards();
  renderCollabView();
  renderStats();
  toast('Mission modifiée ✓');
}

function deleteMission(id) {
  const m = DB.missions.find(x => x.id === id);
  openModal(`Supprimer la mission "${m.titre}" ?`, () => {
    DB.missions = DB.missions.filter(x => x.id !== id);
    save(); renderMissions(); renderCards();
    toast('Mission supprimée');
  });
}

function renderMissions() {
  const tbody = document.getElementById('tbody-missions');
  updateSortHeaders('table-missions', 'missions');

  // Mettre à jour les filtres
  const pfClient = document.getElementById('param-f-client');
  const pfCollab = document.getElementById('param-f-collab');
  if (pfClient && pfCollab) {
    const prevClient = pfClient.value, prevCollab = pfCollab.value;
    pfClient.innerHTML = '<option value="">Tous</option>' + DB.clients.map(c => `<option value="${c.id}">${c.nom}</option>`).join('');
    pfCollab.innerHTML = '<option value="">Tous</option>' + DB.collaborateurs.map(c => `<option value="${c.id}">${c.prenom} ${c.nom}</option>`).join('');
    pfClient.value = prevClient; pfCollab.value = prevCollab;
  }

  const pfStatut = document.getElementById('param-f-statut');
  const fClient = pfClient ? pfClient.value : '';
  const fCollab = pfCollab ? pfCollab.value : '';
  const fStatut = pfStatut ? pfStatut.value : '';
  let list = DB.missions.filter(m => {
    if (fClient && m.clientId !== fClient) return false;
    if (fCollab && !(m.collabIds || []).includes(fCollab)) return false;
    if (fStatut && getStatut(m) !== fStatut) return false;
    return true;
  });
  list = applySortMissions(list);

  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="6" style="color:var(--text-muted);text-align:center;padding:1.5rem">Aucune mission</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(m => {
    const client = DB.clients.find(c => c.id === m.clientId);
    const collabs = (m.collabIds || []).map(id => {
      const c = DB.collaborateurs.find(x => x.id === id);
      return c ? `${c.prenom} ${c.nom}` : '';
    }).filter(Boolean).join(', ');
    const periode = [m.debut, m.fin].filter(Boolean).map(d => formatDate(d)).join(' → ') || '—';
    const statut = getStatut(m);
    const badgeCls = statut === 'en_cours' ? 'badge-encours' : 'badge-terminee';
    const badgeLbl = statut === 'en_cours' ? 'En cours' : 'Terminée';
    return `
      <tr>
        <td>${m.titre || '—'}</td>
        <td>${client ? client.nom : '—'}</td>
        <td style="color:var(--text-muted)">${collabs || '—'}</td>
        <td style="color:var(--text-muted);font-size:0.82rem">${periode}</td>
        <td><span class="badge ${badgeCls}">${badgeLbl}</span></td>
        <td style="display:flex;gap:0.5rem">
          <button class="btn btn-primary btn-sm" onclick="openEditMission('${m.id}')">Modifier</button>
          <button class="btn btn-danger btn-sm" onclick="deleteMission('${m.id}')">Supprimer</button>
        </td>
      </tr>`;
  }).join('');
}

// ══════════════════════════════════════════
//  VISUALISATION — VUE & CARTES
// ══════════════════════════════════════════
let currentView = 'client';

function setView(view) {
  currentView = view;
  document.getElementById('view-btn-client').classList.toggle('active', view === 'client');
  document.getElementById('view-btn-collab').classList.toggle('active', view === 'collab');
  // Masquer le filtre client en vue "par client" (inutile de filtrer par client si on est groupé par client)
  document.getElementById('f-client-group').style.display = view === 'client' ? 'none' : '';
  document.getElementById('f-collab-group').style.display = view === 'collab' ? 'none' : '';
  renderCards();
}

function getFilteredMissions() {
  const fStatutEl = document.querySelector('input[name="f-statut"]:checked');
  const fStatut = fStatutEl ? fStatutEl.value : 'en_cours';
  const fClient = document.getElementById('f-client').value;
  const fCollab = document.getElementById('f-collab').value;
  const fSearch = document.getElementById('f-search').value.toLowerCase();

  return DB.missions.filter(m => {
    if (fStatut && getStatut(m) !== fStatut) return false;
    if (fClient && m.clientId !== fClient) return false;
    if (fCollab && !(m.collabIds || []).includes(fCollab)) return false;
    if (fSearch && !(m.titre || '').toLowerCase().includes(fSearch)) return false;
    return true;
  });
}

function renderCards() {
  if (currentView === 'client') renderByClient();
  else renderByCollab();
}

function renderByClient() {
  const missions = getFilteredMissions();
  const grid = document.getElementById('missions-grid');
  grid.style.gridTemplateColumns = 'repeat(4, 1fr)';

  if (!missions.length) {
    grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1">
      <div class="icon">🔍</div>
      <p>Aucune mission trouvée.</p>
    </div>`;
    return;
  }

  // Regrouper par client
  const byClient = {};
  missions.forEach(m => {
    const cid = m.clientId || '__sans__';
    if (!byClient[cid]) byClient[cid] = [];
    byClient[cid].push(m);
  });

  // Trier les clients par nom
  const clientIds = Object.keys(byClient).sort((a, b) => {
    const ca = DB.clients.find(x => x.id === a);
    const cb = DB.clients.find(x => x.id === b);
    return (ca ? ca.nom : '').localeCompare(cb ? cb.nom : '', 'fr');
  });

  grid.innerHTML = clientIds.map(cid => {
    const client = DB.clients.find(x => x.id === cid);
    const clientMissions = byClient[cid];

    // Construire la liste unique des collabs avec dates de début/fin et statut
    const collabMap = {};
    clientMissions.forEach(m => {
      const terminee = getStatut(m) === 'terminee';
      (m.collabIds || []).forEach(id => {
        if (!collabMap[id]) {
          collabMap[id] = { debut: m.debut, fin: m.fin, toutTerminee: terminee };
        } else {
          if (m.debut && (!collabMap[id].debut || m.debut < collabMap[id].debut)) collabMap[id].debut = m.debut;
          if (m.fin  && (!collabMap[id].fin  || m.fin  > collabMap[id].fin))  collabMap[id].fin  = m.fin;
          if (!terminee) collabMap[id].toutTerminee = false;
        }
      });
    });

    const collabItems = Object.entries(collabMap).map(([id, info]) => {
      const c = DB.collaborateurs.find(x => x.id === id);
      if (!c) return null;
      let dateStr = '';
      if (info.toutTerminee && info.debut && info.fin) {
        dateStr = `<span class="collab-date">${formatDateCourt(info.debut)} - ${formatDateCourt(info.fin)}</span>`;
      } else if (info.debut) {
        dateStr = `<span class="collab-date">${formatDateCourt(info.debut)}</span>`;
      }
      return `<li>${c.prenom} ${c.nom} ${dateStr}</li>`;
    }).filter(Boolean).sort().join('');

    // Logo
    let logoContent;
    if (client && client.logo && isUrl(client.logo)) {
      logoContent = `<img src="${client.logo}" onerror="this.parentElement.innerHTML='<span class=\\'client-card-logo-initiales\\'>${(client.nom||'?').split(' ').map(w=>w[0]).join('').toUpperCase().slice(0,2)}</span>'">`;
    } else if (client) {
      const initiales = (client.nom || '?').split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
      logoContent = `<span class="client-card-logo-initiales" style="color:var(--bg)">${initiales}</span>`;
    } else {
      logoContent = `<span class="client-card-logo-initiales" style="color:var(--bg)">?</span>`;
    }

    return `
    <div class="client-card">
      <div class="client-card-logo-box">${logoContent}</div>
      <ul class="client-collab-list">${collabItems}</ul>
    </div>`;
  }).join('');
}

function renderByCollab() {
  const missions = getFilteredMissions();
  const grid = document.getElementById('missions-grid');
  grid.style.gridTemplateColumns = 'repeat(auto-fill, minmax(260px, 1fr))';

  if (!missions.length) {
    grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1">
      <div class="icon">🔍</div>
      <p>Aucune mission trouvée.</p>
    </div>`;
    return;
  }

  // Regrouper par collaborateur
  const byCollab = {};
  missions.forEach(m => {
    (m.collabIds || []).forEach(id => {
      if (!byCollab[id]) byCollab[id] = [];
      byCollab[id].push(m);
    });
  });

  const collabIds = Object.keys(byCollab).sort((a, b) => {
    const ca = DB.collaborateurs.find(x => x.id === a);
    const cb = DB.collaborateurs.find(x => x.id === b);
    return (ca ? ca.prenom + ' ' + ca.nom : '').localeCompare(cb ? cb.prenom + ' ' + cb.nom : '', 'fr');
  });

  grid.innerHTML = collabIds.map(cid => {
    const collab = DB.collaborateurs.find(x => x.id === cid);
    if (!collab) return '';
    const collabMissions = byCollab[cid];
    const initiales = (collab.prenom[0] + collab.nom[0]).toUpperCase();

    const missionItems = collabMissions.map(m => {
      const client = DB.clients.find(x => x.id === m.clientId);
      const statut = getStatut(m);
      const badgeCls = statut === 'en_cours' ? 'badge-encours' : 'badge-terminee';
      const periode = [m.debut, m.fin].filter(Boolean).map(d => formatDate(d)).join(' → ');
      return `
        <div style="padding:0.6rem 0;border-bottom:1px solid var(--border)">
          <div style="display:flex;align-items:center;gap:0.5rem;margin-bottom:0.3rem">
            ${client ? `<span style="font-size:0.9rem;font-weight:600;color:var(--accent)">${client.nom}</span>` : ''}
            <span class="badge ${badgeCls}" style="margin:0;font-size:0.7rem">${statut === 'en_cours' ? 'En cours' : 'Terminée'}</span>
          </div>
          ${m.titre ? `<div style="font-size:0.85rem;color:var(--text-muted)">${m.titre}</div>` : ''}
          ${periode ? `<div style="font-size:0.8rem;color:var(--text-muted)">📅 ${periode}</div>` : ''}
        </div>`;
    }).join('');

    return `
    <div class="card">
      <div style="display:flex;align-items:center;gap:0.8rem;margin-bottom:1rem">
        <div style="width:44px;height:44px;border-radius:50%;background:var(--card-alt);display:flex;align-items:center;justify-content:center;font-weight:700;color:var(--accent);font-size:1rem;flex-shrink:0">${initiales}</div>
        <div>
          <div style="font-weight:700;font-size:1rem">${collab.prenom} ${collab.nom}</div>
          <div style="font-size:0.8rem;color:var(--text-muted)">${collabMissions.length} mission${collabMissions.length > 1 ? 's' : ''}</div>
        </div>
      </div>
      ${missionItems}
    </div>`;
  }).join('');
}

function renderStats() {
  const today  = new Date().toISOString().split('T')[0];
  const annee  = today.slice(0, 4);

  // Vision globale
  const totalMissions = DB.missions.length;
  const totalClients  = DB.clients.length;

  // Vision année en cours
  const debutAnnee = annee + '-01-01';
  const finAnnee   = annee + '-12-31';
  const missionsDemarrees = DB.missions.filter(m => (m.debut || '').startsWith(annee));
  const missionsTerminees = DB.missions.filter(m => (m.fin || '').startsWith(annee) && m.fin < today);
  const demarragesAnnee = missionsDemarrees.length;
  const termineesAnnee  = missionsTerminees.length;

  function missionTooltipLines(list) {
    return list.map(m => {
      const cl = DB.clients.find(x => x.id === m.clientId);
      const collabs = (m.collabIds || []).map(id => { const c = DB.collaborateurs.find(x => x.id === id); return c ? c.prenom + ' ' + c.nom : ''; }).filter(Boolean).join(', ');
      return (cl ? cl.nom : '—') + (collabs ? ' · ' + collabs : '');
    }).join('\n');
  }
  const tooltipDemarrages = missionsDemarrees.length ? missionTooltipLines(missionsDemarrees) : '';
  const tooltipTerminees  = missionsTerminees.length  ? missionTooltipLines(missionsTerminees)  : '';

  // Missions actives pendant l'année (démarrées avant fin d'année ET pas terminées avant début d'année)
  const missionsAnnee = DB.missions.filter(m =>
    m.debut && m.debut <= finAnnee && (!m.fin || m.fin >= debutAnnee)
  );
  const clientsAnneeIds = [...new Set(missionsAnnee.map(m => m.clientId).filter(Boolean))];
  const nbClientsAnnee  = clientsAnneeIds.length;

  // Nouveaux clients : mission démarrée en année ET aucune mission avant cette année pour ce client
  const nouveauxClients = clientsAnneeIds.filter(cid => {
    const aDebutEnAnnee = missionsAnnee.some(m => m.clientId === cid && (m.debut || '').startsWith(annee));
    if (!aDebutEnAnnee) return false;
    const avantAnnee = DB.missions.some(m => m.clientId === cid && m.debut && m.debut < debutAnnee);
    return !avantAnnee;
  });
  const nbNouveauxClients = nouveauxClients.length;
  const tooltipNouveaux = nouveauxClients.length
    ? nouveauxClients.map(id => { const cl = DB.clients.find(x => x.id === id); return cl ? '· ' + cl.nom : ''; }).filter(Boolean).join('\n')
    : '';

  // Vision à l'instant T
  const enCours       = DB.missions.filter(m => getStatut(m) === 'en_cours').length;
  const clientsActifs = new Set(DB.missions.filter(m => getStatut(m) === 'en_cours').map(m => m.clientId).filter(Boolean)).size;
  const collabsPresents = DB.collaborateurs.filter(c => c.dateEntree && c.dateEntree <= today && (!c.dateSortie || c.dateSortie > today)).length;

  document.getElementById('stats-row').innerHTML = `
    <div style="display:flex;flex-direction:column;gap:0.4rem">
      <div style="font-size:0.72rem;text-transform:uppercase;letter-spacing:0.1em;color:var(--text-muted);font-weight:700;padding-left:0.2rem">Vision globale</div>
      <div style="display:flex;gap:0.8rem">
        <div class="stat-chip"><span class="val">${totalMissions}</span><span class="lbl">Missions totales</span></div>
        <div class="stat-chip"><span class="val" style="color:var(--warning)">${totalClients}</span><span class="lbl">Clients total</span></div>
      </div>
    </div>
    <div style="width:1px;background:var(--border);align-self:stretch;margin:0 0.5rem"></div>
    <div style="display:flex;flex-direction:column;gap:0.4rem">
      <div style="font-size:0.72rem;text-transform:uppercase;letter-spacing:0.1em;color:var(--text-muted);font-weight:700;padding-left:0.2rem">Année ${annee}</div>
      <div style="display:flex;gap:0.8rem">
        <div class="stat-chip" ${tooltipDemarrages ? `data-tooltip="${tooltipDemarrages.replace(/"/g,'&quot;')}"` : ''}><span class="val" style="color:var(--accent2)">${demarragesAnnee}</span><span class="lbl">Démarrages</span></div>
        <div class="stat-chip" ${tooltipTerminees ? `data-tooltip="${tooltipTerminees.replace(/"/g,'&quot;')}"` : ''}><span class="val">${termineesAnnee}</span><span class="lbl">Terminées</span></div>
        <div class="stat-chip"><span class="val" style="color:var(--warning)">${nbClientsAnnee}</span><span class="lbl">Clients différents</span></div>
        <div class="stat-chip" ${tooltipNouveaux ? `data-tooltip="Nouveaux clients :\n${tooltipNouveaux}"` : ''}><span class="val" style="color:var(--accent)">${nbNouveauxClients}</span><span class="lbl">Nouveaux clients</span></div>
      </div>
    </div>
    <div style="width:1px;background:var(--border);align-self:stretch;margin:0 0.5rem"></div>
    <div style="display:flex;flex-direction:column;gap:0.4rem">
      <div style="font-size:0.72rem;text-transform:uppercase;letter-spacing:0.1em;color:var(--text-muted);font-weight:700;padding-left:0.2rem">À l'instant T</div>
      <div style="display:flex;gap:0.8rem">
        <div class="stat-chip"><span class="val" style="color:var(--accent)">${enCours}</span><span class="lbl">Missions en cours</span></div>
        <div class="stat-chip"><span class="val" style="color:var(--accent)">${clientsActifs}</span><span class="lbl">Clients actifs</span></div>
        <div class="stat-chip"><span class="val" style="color:var(--warning)">${collabsPresents}</span><span class="lbl">Collaborateurs</span></div>
      </div>
    </div>
  `;
}

function renderAll() {
  refreshAnneesOptions();
  renderCollaborateurs();
  renderClients();
  renderMissions();
  renderPerimetres();
  renderMethodes();
  refreshSelects();
  refreshAnalyseSelect();
  renderStats();
  // Appliquer l'état initial de la vue
  setView(currentView);
}

// ══════════════════════════════════════════
//  IMPORT EXCEL
// ══════════════════════════════════════════
// ══════════════════════════════════════════
//  DROPDOWN COLLABORATEUR AVEC FILTRE
// ══════════════════════════════════════════
function filterCollabDropdown(prefix) {
  const search = (document.getElementById(prefix + '-collab-search').value || '').toLowerCase();
  const today  = new Date().toISOString().split('T')[0];
  const dd     = document.getElementById(prefix + '-collab-dd');
  const matches = DB.collaborateurs
    .filter(c => !c.dateSortie || c.dateSortie >= today)
    .filter(c => {
      const label = (c.nom + ' ' + c.prenom).toLowerCase();
      return !search || label.includes(search) || c.prenom.toLowerCase().includes(search) || c.nom.toLowerCase().includes(search);
    })
    .sort((a, b) => a.prenom.localeCompare(b.prenom, 'fr'));
  dd.innerHTML = matches.length
    ? matches.map(c => `<div class="collab-dd-item" onmousedown="selectCollabDD('${prefix}','${c.id}','${c.prenom} ${c.nom}')">${c.prenom} ${c.nom}</div>`).join('')
    : '<div style="padding:0.5rem 0.9rem;color:var(--text-muted);font-size:0.85rem">Aucun résultat</div>';
  dd.style.display = 'block';
}
function showCollabDropdown(prefix) { filterCollabDropdown(prefix); }
function hideCollabDropdown(prefix) {
  setTimeout(() => { const dd = document.getElementById(prefix + '-collab-dd'); if (dd) dd.style.display = 'none'; }, 200);
}
function selectCollabDD(prefix, id, label) {
  document.getElementById(prefix + '-collabs').value = id;
  document.getElementById(prefix + '-collab-search').value = label;
  document.getElementById(prefix + '-collab-dd').style.display = 'none';
}

// ══════════════════════════════════════════
//  IMPORT EN MASSE — MISSIONS
// ══════════════════════════════════════════
function importMissions(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = '';
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets['Feuil1'] || wb.Sheets[wb.SheetNames[0]];
      if (!ws) { toast('Onglet "Feuil1" introuvable', 'error'); return; }
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
      if (!rows.length) { toast('Fichier vide', 'error'); return; }
      let created = 0, updated = 0;
      const skipped = [];

      rows.forEach((row, i) => {
        const titre      = (row['Titre de la mission'] || '').toString().trim();
        const clientNom  = (row['Client'] || '').toString().trim();
        const collabStr  = (row['Collaborateurs'] || '').toString().trim();
        const debut      = excelDateToISO(row['Date Début Mission'] || row['Date Début'] || row['Date debut'] || '');
        const fin        = excelDateToISO(row['Date Fin Mission']   || row['Date Fin']   || row['Date fin']  || '');
        const details    = (row['Détails Missions'] || row['Détails de la mission'] || '').toString().trim();
        const lineNum    = i + 2;

        if (!clientNom) { skipped.push(`Ligne ${lineNum} (${titre || '—'}) : client manquant`); return; }
        if (!collabStr) { skipped.push(`Ligne ${lineNum} (${titre || '—'}) : collaborateur manquant`); return; }
        if (!debut)     { skipped.push(`Ligne ${lineNum} (${titre || '—'}) : date de début manquante`); return; }

        // Collaborateur : doit exister
        const collab = DB.collaborateurs.find(c => (c.nom + ' ' + c.prenom).toLowerCase() === collabStr.toLowerCase()
          || (c.prenom + ' ' + c.nom).toLowerCase() === collabStr.toLowerCase());
        if (!collab) { skipped.push(`Ligne ${lineNum} (${titre}) : collaborateur "${collabStr}" introuvable`); return; }

        // Client : créer si inexistant
        let client = DB.clients.find(c => c.nom.toLowerCase() === clientNom.toLowerCase());
        if (!client) { client = { id: uid(), nom: clientNom, logo: '' }; DB.clients.push(client); }

        // Mission existante (même collaborateur + client + dates) → mise à jour
        const existing = DB.missions.find(m =>
          m.clientId === client.id &&
          m.debut === debut &&
          m.fin === (fin || '') &&
          (m.collabIds || []).includes(collab.id)
        );
        if (existing) {
          if (!existing.collabIds.includes(collab.id)) existing.collabIds.push(collab.id);
          if (debut) existing.debut = debut;
          if (fin)   existing.fin   = fin;
          if (details) existing.details = details;
          updated++;
        } else {
          DB.missions.push({ id: uid(), titre, clientId: client.id, collabIds: [collab.id], debut, fin: fin || '', details });
          created++;
        }
      });

      save();
      renderAll();

      // Afficher le résumé
      const summaryDiv = document.getElementById('import-missions-summary');
      let html = `<div style="background:var(--card-bg);border:1px solid var(--border);border-radius:var(--radius);padding:1.2rem">
        <div style="font-weight:700;margin-bottom:0.6rem;color:var(--accent)">📊 Résumé de l'import</div>
        <div>✅ ${created} mission(s) créée(s)</div>
        <div>🔄 ${updated} mission(s) mise(s) à jour</div>`;
      if (skipped.length) {
        html += `<div style="margin-top:0.6rem;color:var(--warning);font-weight:600">⚠️ ${skipped.length} ligne(s) ignorée(s) :</div>
          <ul style="margin:0.3rem 0 0 1.2rem;color:var(--text-muted);font-size:0.85rem">`;
        skipped.forEach(s => { html += `<li>${s}</li>`; });
        html += '</ul>';
      }
      html += '</div>';
      summaryDiv.innerHTML = html;
      summaryDiv.style.display = 'block';
    } catch(err) {
      toast('Erreur : ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function handleDrop(event) {
  event.preventDefault();
  document.getElementById('drop-zone').classList.remove('drag-over');
  const file = event.dataTransfer.files[0];
  if (file) handleExcelFile(file);
}

function handleExcelFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array' });
      const ws   = wb.Sheets['Feuil1'] || wb.Sheets[wb.SheetNames[0]];
      if (!ws) { toast('Onglet "Feuil1" introuvable', 'error'); return; }
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });

      // Regrouper les lignes par mission (Titre Mission + Client + Date Début + Date Fin)
      const missionsMap = {};
      rows.forEach(row => {
        const titre = (row['Titre Mission'] || '').toString().trim();
        if (!titre) return;
        const clientNom = (row['Client'] || '').toString().trim();
        const debut     = excelDateToISO(row['Date Début'] || row['Date Debut'] || '');
        const fin       = excelDateToISO(row['Date Fin']   || '');
        const details   = (row['Détails Missions'] || row['Details Missions'] || '').toString().trim();
        const nom       = (row['Nom']    || '').toString().trim();
        const prenom    = (row['Prénom'] || row['Prenom'] || '').toString().trim();

        const key = [titre, clientNom, debut, fin].join('|');
        if (!missionsMap[key]) {
          missionsMap[key] = { titre, clientNom, debut, fin, details, collabIds: [] };
        }

        if (nom || prenom) {
          const fullName = (prenom + ' ' + nom).trim().toLowerCase();
          let c = DB.collaborateurs.find(x => (x.prenom + ' ' + x.nom).toLowerCase() === fullName);
          if (!c) { c = { id: uid(), nom, prenom, sexe: '', dateEntree: '', dateSortie: '', copilote: '' }; DB.collaborateurs.push(c); }
          if (!missionsMap[key].collabIds.includes(c.id)) missionsMap[key].collabIds.push(c.id);
        }
      });

      let added = 0;
      Object.values(missionsMap).forEach(({ titre, clientNom, debut, fin, details, collabIds }) => {
        let clientId = '';
        if (clientNom) {
          let c = DB.clients.find(x => x.nom.toLowerCase() === clientNom.toLowerCase());
          if (!c) { c = { id: uid(), nom: clientNom, logo: '' }; DB.clients.push(c); }
          clientId = c.id;
        }
        const statut = (fin && fin < new Date().toISOString().split('T')[0]) ? 'terminee' : 'en_cours';
        DB.missions.push({ id: uid(), titre, clientId, collabIds, debut, fin, details, statut });
        added++;
      });

      save();
      renderAll();
      toast(`${added} mission(s) importée(s) ✓`);
    } catch(err) {
      toast('Erreur : ' + err.message, 'error');
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

function excelDateToString(val) {
  if (!val) return '';
  if (typeof val === 'number') {
    // numéro de série Excel → Date
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return d.toISOString().split('T')[0];
  }
  const s = val.toString().trim();
  // Essai formats FR (dd/mm/yyyy) ou ISO
  const fr = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (fr) return `${fr[3]}-${fr[2].padStart(2,'0')}-${fr[1].padStart(2,'0')}`;
  return s;
}

function formatDate(d) {
  if (!d) return '';
  const [y, m, day] = d.split('-');
  return `${day}/${m}/${y}`;
}
function formatDateCourt(d) {
  if (!d) return '';
  const [y, m, day] = d.split('-');
  return `${day}/${m}/${y.slice(2)}`;
}

// ══════════════════════════════════════════
//  ANALYSE
// ══════════════════════════════════════════
let analyseView = 'perimetre';

function filtreEffectifsActifs() {
  const el = document.querySelector('input[name="an-effectifs"]:checked');
  return el && el.value === 'actif';
}
function collabEstActif(id) {
  const today = new Date().toISOString().split('T')[0];
  const c = DB.collaborateurs.find(x => x.id === id);
  return c && (!c.dateSortie || c.dateSortie >= today);
}

function setAnalyseView(view) {
  analyseView = view;
  document.getElementById('an-btn-perimetre').classList.toggle('active', view === 'perimetre');
  document.getElementById('an-btn-methode').classList.toggle('active', view === 'methode');
  document.getElementById('an-filter-perimetre').style.display = view === 'perimetre' ? '' : 'none';
  document.getElementById('an-filter-methode').style.display   = view === 'methode'   ? '' : 'none';
  if (view === 'methode') refreshAnalyseMethodesPicker();
  renderAnalyse();
}

function refreshAnalyseSelect() {
  const sel = document.getElementById('an-perimetre');
  if (!sel) return;
  const prev = sel.value;
  sel.innerHTML = '<option value="">— Tous les périmètres —</option>' +
    [...(DB.perimetres || [])].sort((a,b) => a.nom.localeCompare(b.nom,'fr'))
      .map(p => `<option value="${p.id}">${p.nom}</option>`).join('');
  sel.value = prev;
}

function refreshAnalyseMethodesPicker() {
  const container = document.getElementById('an-methodes-picker');
  if (!container) return;
  const sel = getSelectedTags('an-methodes-picker');
  const items = [...(DB.methodes || [])].sort((a,b) => a.nom.localeCompare(b.nom,'fr'));
  if (!items.length) {
    container.innerHTML = '<span class="tag-pick-empty">Aucune méthode disponible — ajoutez-en dans Paramétrage.</span>';
    return;
  }
  container.innerHTML = items.map(item =>
    `<span class="tag-pick${sel.includes(item.id) ? ' selected' : ''}" data-id="${item.id}" onclick="this.classList.toggle('selected');renderAnalyse()">${item.nom}</span>`
  ).join('');
}

function renderAnalyse() {
  if (analyseView === 'methode') { renderAnalyseMethodes(); return; }

  const perimetreId = document.getElementById('an-perimetre').value;
  const statutEl   = document.querySelector('input[name="an-statut"]:checked');
  const fStatut    = statutEl ? statutEl.value : '';
  const content    = document.getElementById('analyse-content');

  // Filtrer les missions
  let missions = DB.missions.filter(m => {
    if (perimetreId && !(m.perimetreIds || []).includes(perimetreId)) return false;
    if (fStatut && getStatut(m) !== fStatut) return false;
    if (filtreEffectifsActifs()) {
      const hasActif = (m.collabIds || []).some(id => collabEstActif(id));
      if (!hasActif) return false;
    }
    return true;
  });

  if (!missions.length) {
    const msg = perimetreId
      ? 'Aucune mission trouvée pour ce périmètre avec ce filtre.'
      : 'Aucune mission trouvée.';
    content.innerHTML = `<div class="empty-state"><div class="icon">🔍</div><p>${msg}</p></div>`;
    return;
  }

  const perimetreNom = perimetreId
    ? (DB.perimetres.find(p => p.id === perimetreId) || {}).nom || ''
    : 'Tous les périmètres';

  // ── Chiffres clés ──
  const nbMissions   = missions.length;
  const nbEncours    = missions.filter(m => getStatut(m) === 'en_cours').length;
  const nbTerminees  = missions.filter(m => getStatut(m) === 'terminee').length;
  const clientIds    = [...new Set(missions.map(m => m.clientId).filter(Boolean))];
  const collabIds    = [...new Set(missions.flatMap(m => m.collabIds || []))];

  // ── Collaborateurs ──
  const collabStats = collabIds.map(id => {
    const c = DB.collaborateurs.find(x => x.id === id);
    if (!c) return null;
    const ms = missions.filter(m => (m.collabIds || []).includes(id));
    const dates = ms.map(m => m.debut).filter(Boolean).sort();
    const fins  = ms.map(m => m.fin).filter(Boolean).sort();
    const debut = dates[0] ? formatDate(dates[0]) : '—';
    const fin   = fins[fins.length-1] ? formatDate(fins[fins.length-1]) : 'en cours';
    return { nom: `${c.prenom} ${c.nom}`, nb: ms.length, debut, fin };
  }).filter(Boolean).sort((a,b) => b.nb - a.nb);
  const maxCollabNb = collabStats[0] ? collabStats[0].nb : 1;

  // ── Clients ──
  const clientStats = clientIds.map(id => {
    const cl = DB.clients.find(x => x.id === id);
    if (!cl) return null;
    const ms = missions.filter(m => m.clientId === id);
    return { nom: cl.nom, logo: cl.logo, nb: ms.length };
  }).filter(Boolean).sort((a,b) => b.nb - a.nb);
  const maxClientNb = clientStats[0] ? clientStats[0].nb : 1;

  // ── Méthodes/Outils les plus utilisés ──
  const methodeCount = {};
  missions.forEach(m => (m.methodeIds || []).forEach(id => { methodeCount[id] = (methodeCount[id] || 0) + 1; }));
  const methodeStats = Object.entries(methodeCount)
    .map(([id, nb]) => { const mo = DB.methodes.find(x => x.id === id); return mo ? { nom: mo.nom, nb } : null; })
    .filter(Boolean).sort((a,b) => b.nb - a.nb);
  const maxMethodeNb = methodeStats[0] ? methodeStats[0].nb : 1;

  // ── Périmètres associés (si vue "tous") ──
  let perimetresAssocHtml = '';
  if (!perimetreId) {
    const pCount = {};
    missions.forEach(m => (m.perimetreIds || []).forEach(id => { pCount[id] = (pCount[id] || 0) + 1; }));
    const pStats = Object.entries(pCount)
      .map(([id, nb]) => { const p = DB.perimetres.find(x => x.id === id); return p ? { nom: p.nom, nb } : null; })
      .filter(Boolean).sort((a,b) => b.nb - a.nb);
    const maxP = pStats[0] ? pStats[0].nb : 1;
    if (pStats.length) {
      perimetresAssocHtml = `
      <div class="analyse-card">
        <div class="analyse-card-title">🎯 Périmètres Métier</div>
        ${pStats.map(p => `
          <div class="analyse-row">
            <span class="name">${p.nom}</span>
            <div class="analyse-bar-wrap">
              <div class="analyse-bar-bg"><div class="analyse-bar-fill" style="width:${Math.round(p.nb/maxP*100)}%"></div></div>
              <span class="meta">${p.nb} mission${p.nb>1?'s':''}</span>
            </div>
          </div>`).join('')}
      </div>`;
    }
  }

  content.innerHTML = `
    <div style="margin-bottom:0.5rem;font-size:1.05rem;font-weight:700;color:var(--text)">${perimetreNom}</div>

    <div class="analyse-grid">

      <div class="analyse-card">
        <div class="analyse-card-title">👤 Collaborateurs</div>
        ${collabStats.length ? collabStats.map(c => `
          <div class="analyse-row">
            <div>
              <div class="name">${c.nom}</div>
              <div class="meta">${c.debut} → ${c.fin}</div>
            </div>
            <div class="analyse-bar-wrap">
              <div class="analyse-bar-bg"><div class="analyse-bar-fill" style="width:${Math.round(c.nb/maxCollabNb*100)}%"></div></div>
              <span class="meta">${c.nb} mission${c.nb>1?'s':''}</span>
            </div>
          </div>`).join('') : '<p style="color:var(--text-muted);font-size:0.85rem">Aucun collaborateur</p>'}
      </div>

      <div class="analyse-card">
        <div class="analyse-card-title">🏢 Clients</div>
        ${clientStats.length ? clientStats.map(cl => `
          <div class="analyse-row">
            <span class="name">${cl.nom}</span>
            <div class="analyse-bar-wrap">
              <div class="analyse-bar-bg"><div class="analyse-bar-fill" style="width:${Math.round(cl.nb/maxClientNb*100)}%"></div></div>
              <span class="meta">${cl.nb} mission${cl.nb>1?'s':''}</span>
            </div>
          </div>`).join('') : '<p style="color:var(--text-muted);font-size:0.85rem">Aucun client</p>'}
      </div>

      ${methodeStats.length ? `
      <div class="analyse-card">
        <div class="analyse-card-title">🔧 Méthodes / Outils clés</div>
        ${methodeStats.map(mo => `
          <div class="analyse-row">
            <span class="name">${mo.nom}</span>
            <div class="analyse-bar-wrap">
              <div class="analyse-bar-bg"><div class="analyse-bar-fill" style="width:${Math.round(mo.nb/maxMethodeNb*100)}%;background:var(--accent2)"></div></div>
              <span class="meta">${mo.nb} mission${mo.nb>1?'s':''}</span>
            </div>
          </div>`).join('')}
      </div>` : ''}

      ${perimetresAssocHtml}

    </div>`;
}

function buildAnalyseStatsHtml(missions) {
  const today = new Date().toISOString().split('T')[0];
  // Vision globale
  const nbTotal     = missions.length;
  const nbTerminees = missions.filter(m => getStatut(m) === 'terminee').length;
  const nbClients   = new Set(missions.map(m => m.clientId).filter(Boolean)).size;
  const nbCollabs   = new Set(missions.flatMap(m => m.collabIds || [])).size;
  // À l'instant T
  const nbEnCours   = missions.filter(m => getStatut(m) === 'en_cours').length;
  const collabsActifs = new Set(
    missions
      .filter(m => getStatut(m) === 'en_cours')
      .flatMap(m => m.collabIds || [])
      .filter(id => {
        const c = DB.collaborateurs.find(x => x.id === id);
        return c && (!c.dateSortie || c.dateSortie >= today);
      })
  ).size;

  const chip = (val, lbl, color) =>
    `<div class="stat-chip"><span class="val" style="color:${color||'var(--accent)'}">${val}</span><span class="lbl">${lbl}</span></div>`;
  const sep = `<div style="width:1px;background:var(--border);align-self:stretch;margin:0 0.5rem"></div>`;
  const grpLabel = txt => `<div style="font-size:0.72rem;text-transform:uppercase;letter-spacing:0.1em;color:var(--text-muted);font-weight:700;padding-left:0.2rem">${txt}</div>`;

  return `<div style="display:flex;gap:1rem;flex-wrap:wrap;align-items:flex-start;margin-bottom:2rem">
    <div style="display:flex;flex-direction:column;gap:0.4rem">
      ${grpLabel('Vision globale')}
      <div style="display:flex;gap:0.8rem;flex-wrap:wrap">
        ${chip(nbTotal,     'Missions totales')}
        ${chip(nbTerminees, 'Terminées')}
        ${chip(nbClients,   'Clients',       'var(--warning)')}
        ${chip(nbCollabs,   'Collaborateurs','var(--warning)')}
      </div>
    </div>
    ${sep}
    <div style="display:flex;flex-direction:column;gap:0.4rem">
      ${grpLabel('À l\'instant T')}
      <div style="display:flex;gap:0.8rem;flex-wrap:wrap">
        ${chip(nbEnCours,     'Missions en cours', 'var(--accent2)')}
        ${chip(collabsActifs, 'Collabs dans les effectifs')}
      </div>
    </div>
  </div>`;
}

// ──────────────────────────────────────────
function renderAnalyseMethodes() {
  const content  = document.getElementById('analyse-content');
  const statutEl = document.querySelector('input[name="an-statut"]:checked');
  const fStatut  = statutEl ? statutEl.value : '';
  const selIds   = getSelectedTags('an-methodes-picker');

  if (!selIds.length) {
    content.innerHTML = `<div class="empty-state"><div class="icon">🔧</div><p>Sélectionnez une ou plusieurs méthodes / outils pour afficher l'analyse.</p></div>`;
    return;
  }

  // Missions ayant au moins une des méthodes sélectionnées
  let missions = DB.missions.filter(m => {
    if (fStatut && getStatut(m) !== fStatut) return false;
    if (!(m.methodeIds || []).some(id => selIds.includes(id))) return false;
    if (filtreEffectifsActifs()) {
      const hasActif = (m.collabIds || []).some(id => collabEstActif(id));
      if (!hasActif) return false;
    }
    return true;
  });

  const selNoms = selIds.map(id => { const mo = DB.methodes.find(x => x.id === id); return mo ? mo.nom : ''; }).filter(Boolean);

  if (!missions.length) {
    content.innerHTML = `<div class="empty-state"><div class="icon">🔍</div><p>Aucune mission trouvée pour cette sélection.</p></div>`;
    return;
  }

  // Chiffres clés
  const nbEncours   = missions.filter(m => getStatut(m) === 'en_cours').length;
  const nbTerminees = missions.filter(m => getStatut(m) === 'terminee').length;
  const clientIds   = [...new Set(missions.map(m => m.clientId).filter(Boolean))];
  const collabIds   = [...new Set(missions.flatMap(m => m.collabIds || []))];

  // Collaborateurs
  const collabStats = collabIds.map(id => {
    const c = DB.collaborateurs.find(x => x.id === id);
    if (!c) return null;
    const ms = missions.filter(m => (m.collabIds || []).includes(id));
    const dates = ms.map(m => m.debut).filter(Boolean).sort();
    const fins  = ms.map(m => m.fin).filter(Boolean).sort();
    return { nom: `${c.prenom} ${c.nom}`, nb: ms.length,
      debut: dates[0] ? formatDate(dates[0]) : '—',
      fin: fins[fins.length-1] ? formatDate(fins[fins.length-1]) : 'en cours' };
  }).filter(Boolean).sort((a,b) => b.nb - a.nb);
  const maxC = collabStats[0] ? collabStats[0].nb : 1;

  // Clients
  const clientStats = clientIds.map(id => {
    const cl = DB.clients.find(x => x.id === id);
    if (!cl) return null;
    return { nom: cl.nom, nb: missions.filter(m => m.clientId === id).length };
  }).filter(Boolean).sort((a,b) => b.nb - a.nb);
  const maxCl = clientStats[0] ? clientStats[0].nb : 1;

  // Périmètres associés
  const pCount = {};
  missions.forEach(m => (m.perimetreIds || []).forEach(id => { pCount[id] = (pCount[id]||0)+1; }));
  const pStats = Object.entries(pCount)
    .map(([id,nb]) => { const p = DB.perimetres.find(x => x.id === id); return p ? {nom:p.nom,nb} : null; })
    .filter(Boolean).sort((a,b) => b.nb - a.nb);
  const maxP = pStats[0] ? pStats[0].nb : 1;

  // Liste des missions
  const missionsSorted = [...missions].sort((a,b) => {
    const sa = getStatut(a) === 'en_cours' ? 0 : 1;
    const sb = getStatut(b) === 'en_cours' ? 0 : 1;
    if (sa !== sb) return sa - sb;
    return (b.debut||'').localeCompare(a.debut||'');
  });

  content.innerHTML = `
    <div style="margin-bottom:0.5rem;font-size:1.05rem;font-weight:700;color:var(--text)">${selNoms.join(', ')}</div>

    <div class="analyse-grid">

      <div class="analyse-card">
        <div class="analyse-card-title">👤 Collaborateurs</div>
        ${collabStats.map(c => `
          <div class="analyse-row">
            <div><div class="name">${c.nom}</div><div class="meta">${c.debut} → ${c.fin}</div></div>
            <div class="analyse-bar-wrap">
              <div class="analyse-bar-bg"><div class="analyse-bar-fill" style="width:${Math.round(c.nb/maxC*100)}%"></div></div>
              <span class="meta">${c.nb} mission${c.nb>1?'s':''}</span>
            </div>
          </div>`).join('')}
      </div>

      <div class="analyse-card">
        <div class="analyse-card-title">🏢 Clients</div>
        ${clientStats.map(cl => `
          <div class="analyse-row">
            <span class="name">${cl.nom}</span>
            <div class="analyse-bar-wrap">
              <div class="analyse-bar-bg"><div class="analyse-bar-fill" style="width:${Math.round(cl.nb/maxCl*100)}%"></div></div>
              <span class="meta">${cl.nb} mission${cl.nb>1?'s':''}</span>
            </div>
          </div>`).join('')}
      </div>

      ${pStats.length ? `
      <div class="analyse-card">
        <div class="analyse-card-title">🎯 Périmètres Métier associés</div>
        ${pStats.map(p => `
          <div class="analyse-row">
            <span class="name">${p.nom}</span>
            <div class="analyse-bar-wrap">
              <div class="analyse-bar-bg"><div class="analyse-bar-fill" style="width:${Math.round(p.nb/maxP*100)}%;background:var(--accent2)"></div></div>
              <span class="meta">${p.nb} mission${p.nb>1?'s':''}</span>
            </div>
          </div>`).join('')}
      </div>` : ''}

      <div class="analyse-card" style="grid-column:1/-1">
        <div class="analyse-card-title">📋 Missions concernées</div>
        ${missionsSorted.map(m => {
          const cl = DB.clients.find(x => x.id === m.clientId);
          const collabs = (m.collabIds||[]).map(id => { const c = DB.collaborateurs.find(x=>x.id===id); return c ? c.prenom+' '+c.nom : ''; }).filter(Boolean).join(', ');
          const statut = getStatut(m);
          const badgeCls = statut === 'en_cours' ? 'badge-encours' : 'badge-terminee';
          const badgeLbl = statut === 'en_cours' ? 'En cours' : 'Terminée';
          const periode = [m.debut, m.fin].filter(Boolean).map(d => formatDate(d)).join(' → ');
          return `
          <div class="analyse-row" style="flex-wrap:wrap;gap:0.3rem 1rem">
            <div style="flex:1;min-width:200px">
              <div class="name">${m.titre || '—'} ${cl ? `<span style="color:var(--accent);font-weight:400">· ${cl.nom}</span>` : ''}</div>
              ${collabs ? `<div class="meta">👤 ${collabs}</div>` : ''}
            </div>
            <div style="display:flex;gap:0.5rem;align-items:center">
              ${periode ? `<span class="meta">📅 ${periode}</span>` : ''}
              <span class="badge ${badgeCls}" style="margin:0">${badgeLbl}</span>
            </div>
          </div>`;
        }).join('')}
      </div>

    </div>`;
}

// ══════════════════════════════════════════
//  TAG PICKER
// ══════════════════════════════════════════
function renderTagPicker(containerId, items, selectedIds, addType) {
  const container = document.getElementById(containerId);
  if (!container) return;
  const sel = selectedIds || [];
  let html = (items && items.length)
    ? items.sort((a, b) => a.nom.localeCompare(b.nom, 'fr'))
        .map(item => `<span class="tag-pick${sel.includes(item.id) ? ' selected' : ''}" data-id="${item.id}" onclick="this.classList.toggle('selected')">${item.nom}</span>`)
        .join('')
    : '<span class="tag-pick-empty">Aucune valeur —</span>';
  if (addType) {
    html += `<span class="tag-pick tag-pick-add" onclick="showTagAdd('${containerId}','${addType}')" title="Ajouter une valeur">＋</span>`;
  }
  container.innerHTML = html;
}
function getSelectedTags(containerId) {
  return [...document.querySelectorAll(`#${containerId} .tag-pick.selected`)].map(el => el.dataset.id);
}
function showTagAdd(containerId, type) {
  const container = document.getElementById(containerId);
  const existing = container.querySelector('.tag-add-form');
  if (existing) { existing.remove(); return; }
  const form = document.createElement('span');
  form.className = 'tag-add-form';
  form.innerHTML = `
    <input class="tag-add-input" type="text" placeholder="Nouvelle valeur…"
      onkeydown="if(event.key==='Enter'){event.preventDefault();confirmTagAdd('${containerId}','${type}',this)}
                 if(event.key==='Escape')this.closest('.tag-add-form').remove()">
    <span class="tag-pick tag-pick-confirm" onmousedown="confirmTagAdd('${containerId}','${type}',this.previousElementSibling)">✓</span>
    <span class="tag-pick tag-pick-cancel"  onmousedown="this.closest('.tag-add-form').remove()">✕</span>`;
  container.appendChild(form);
  form.querySelector('input').focus();
}
function confirmTagAdd(containerId, type, inputEl) {
  const nom = (inputEl.value || '').trim();
  if (!nom) { inputEl.focus(); return; }
  const list = type === 'perimetre' ? DB.perimetres : DB.methodes;
  const newItem = { id: uid(), nom };
  list.push(newItem);
  save();
  if (type === 'perimetre') renderPerimetres(); else renderMethodes();
  // Rafraîchir les deux tag pickers du même type en conservant la sélection
  ['m-', 'edit-'].forEach(prefix => {
    const cid = prefix + (type === 'perimetre' ? 'perimetres' : 'methodes');
    const sel = getSelectedTags(cid);
    if (cid === containerId) sel.push(newItem.id);
    renderTagPicker(cid, list, sel, type);
  });
  toast(`"${nom}" ajouté ✓`);
}

// ══════════════════════════════════════════
//  PÉRIMÈTRE MÉTIER
// ══════════════════════════════════════════
function savePerimetre() {
  const editId = document.getElementById('pm-edit-id').value;
  const nom = document.getElementById('pm-nom').value.trim();
  if (!nom) { toast('Valeur requise', 'error'); return; }
  if (editId) {
    const p = DB.perimetres.find(x => x.id === editId);
    if (p) p.nom = nom;
    toast('Valeur modifiée ✓');
  } else {
    DB.perimetres.push({ id: uid(), nom });
    toast('Valeur ajoutée ✓');
  }
  save();
  resetPerimetreForm();
  renderPerimetres();
}
function resetPerimetreForm() {
  document.getElementById('pm-edit-id').value = '';
  document.getElementById('pm-nom').value = '';
  document.getElementById('pm-form-title').textContent = 'Ajouter une valeur';
  document.getElementById('pm-submit-btn').textContent = 'Ajouter';
  document.getElementById('pm-cancel-btn').style.display = 'none';
}
function editPerimetre(id) {
  const p = DB.perimetres.find(x => x.id === id);
  document.getElementById('pm-edit-id').value = id;
  document.getElementById('pm-nom').value = p.nom;
  document.getElementById('pm-form-title').textContent = 'Modifier la valeur';
  document.getElementById('pm-submit-btn').textContent = 'Enregistrer';
  document.getElementById('pm-cancel-btn').style.display = 'inline-block';
}
function cancelEditPerimetre() { resetPerimetreForm(); }
function deletePerimetre(id) {
  const p = DB.perimetres.find(x => x.id === id);
  openModal(`Supprimer "${p.nom}" ?`, () => {
    DB.perimetres = DB.perimetres.filter(x => x.id !== id);
    save(); renderPerimetres(); toast('Valeur supprimée');
  });
}
function renderPerimetres() {
  const tbody = document.getElementById('tbody-perimetres');
  const list = [...(DB.perimetres || [])].sort((a, b) => a.nom.localeCompare(b.nom, 'fr'));
  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="2" style="color:var(--text-muted);text-align:center;padding:1.5rem">Aucune valeur</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(p => `
    <tr>
      <td>${p.nom}</td>
      <td style="display:flex;gap:0.5rem">
        <button class="btn btn-primary btn-sm" onclick="editPerimetre('${p.id}')">Modifier</button>
        <button class="btn btn-danger btn-sm" onclick="deletePerimetre('${p.id}')">Supprimer</button>
      </td>
    </tr>`).join('');
}

// ══════════════════════════════════════════
//  MÉTHODES / OUTILS CLÉS
// ══════════════════════════════════════════
function saveMethode() {
  const editId = document.getElementById('mo-edit-id').value;
  const nom = document.getElementById('mo-nom').value.trim();
  if (!nom) { toast('Valeur requise', 'error'); return; }
  if (editId) {
    const m = DB.methodes.find(x => x.id === editId);
    if (m) m.nom = nom;
    toast('Valeur modifiée ✓');
  } else {
    DB.methodes.push({ id: uid(), nom });
    toast('Valeur ajoutée ✓');
  }
  save();
  resetMethodeForm();
  renderMethodes();
}
function resetMethodeForm() {
  document.getElementById('mo-edit-id').value = '';
  document.getElementById('mo-nom').value = '';
  document.getElementById('mo-form-title').textContent = 'Ajouter une valeur';
  document.getElementById('mo-submit-btn').textContent = 'Ajouter';
  document.getElementById('mo-cancel-btn').style.display = 'none';
}
function editMethode(id) {
  const m = DB.methodes.find(x => x.id === id);
  document.getElementById('mo-edit-id').value = id;
  document.getElementById('mo-nom').value = m.nom;
  document.getElementById('mo-form-title').textContent = 'Modifier la valeur';
  document.getElementById('mo-submit-btn').textContent = 'Enregistrer';
  document.getElementById('mo-cancel-btn').style.display = 'inline-block';
}
function cancelEditMethode() { resetMethodeForm(); }
function deleteMethode(id) {
  const m = DB.methodes.find(x => x.id === id);
  openModal(`Supprimer "${m.nom}" ?`, () => {
    DB.methodes = DB.methodes.filter(x => x.id !== id);
    save(); renderMethodes(); toast('Valeur supprimée');
  });
}
function renderMethodes() {
  const tbody = document.getElementById('tbody-methodes');
  const list = [...(DB.methodes || [])].sort((a, b) => a.nom.localeCompare(b.nom, 'fr'));
  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="2" style="color:var(--text-muted);text-align:center;padding:1.5rem">Aucune valeur</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(m => `
    <tr>
      <td>${m.nom}</td>
      <td style="display:flex;gap:0.5rem">
        <button class="btn btn-primary btn-sm" onclick="editMethode('${m.id}')">Modifier</button>
        <button class="btn btn-danger btn-sm" onclick="deleteMethode('${m.id}')">Supprimer</button>
      </td>
    </tr>`).join('');
}

// ══════════════════════════════════════════
//  VUE COLLAB
// ══════════════════════════════════════════
function filterCVCollabDD() {
  const search = (document.getElementById('cv-collab-search').value || '').toLowerCase();
  const dd = document.getElementById('cv-collab-dd');
  const matches = DB.collaborateurs
    .filter(c => !search || c.prenom.toLowerCase().includes(search) || c.nom.toLowerCase().includes(search) || (c.prenom + ' ' + c.nom).toLowerCase().includes(search))
    .sort((a, b) => a.prenom.localeCompare(b.prenom, 'fr'));
  dd.innerHTML = matches.length
    ? matches.map(c => `<div class="collab-dd-item" onmousedown="selectCVCollab('${c.id}','${c.prenom} ${c.nom}')">${c.prenom} ${c.nom}</div>`).join('')
    : '<div style="padding:0.5rem 0.9rem;color:var(--text-muted);font-size:0.85rem">Aucun résultat</div>';
  dd.style.display = 'block';
}
function showCVCollabDD() { filterCVCollabDD(); }
function hideCVCollabDD() {
  setTimeout(() => { const dd = document.getElementById('cv-collab-dd'); if (dd) dd.style.display = 'none'; }, 200);
}
function toggleClearBtn(prefix) {
  const val = document.getElementById(prefix + '-search').value;
  const btn = document.getElementById(prefix + '-clear');
  if (btn) btn.style.display = val ? 'block' : 'none';
}
function clearSearchFilter(prefix, renderFn) {
  document.getElementById(prefix + '-search').value = '';
  document.getElementById(prefix + '-id').value = '';
  toggleClearBtn(prefix);
  window[renderFn]();
}

function selectCVCollab(id, label) {
  document.getElementById('cv-collab-id').value = id;
  document.getElementById('cv-collab-search').value = label;
  document.getElementById('cv-collab-dd').style.display = 'none';
  toggleClearBtn('cv-collab');
  renderCollabView();
}

function renderCollabView() {
  const collabId = document.getElementById('cv-collab-id').value;
  const copiloteId = '';
  const statutEl = document.querySelector('input[name="cv-statut"]:checked');
  const fStatut = statutEl ? statutEl.value : 'en_cours';
  const grid = document.getElementById('cv-missions-grid');

  let missions = DB.missions.filter(m => {
    if (collabId && !(m.collabIds || []).includes(collabId)) return false;

    if (fStatut && getStatut(m) !== fStatut) return false;
    return true;
  });

  // Tri champ 1 : En cours avant Terminées — champ 2 : date début décroissante
  missions = missions.sort((a, b) => {
    const aEncours = getStatut(a) === 'en_cours';
    const bEncours = getStatut(b) === 'en_cours';
    if (aEncours && !bEncours) return -1;   // a en cours, b terminé → a avant
    if (!aEncours && bEncours) return 1;    // a terminé, b en cours → b avant
    // même statut : date de début décroissante
    const da = a.debut || '';
    const db = b.debut || '';
    if (db > da) return 1;
    if (db < da) return -1;
    return 0;
  });

  if (!missions.length) {
    grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1">
      <div class="icon">🔍</div>
      <p>${collabId ? 'Aucune mission trouvée pour ce collaborateur.' : 'Sélectionnez un collaborateur ou choisissez un filtre de statut.'}</p>
    </div>`;
    return;
  }

  grid.innerHTML = missions.map(m => {
    const client = DB.clients.find(c => c.id === m.clientId);
    const collabs = (m.collabIds || []).map(id => {
      const c = DB.collaborateurs.find(x => x.id === id);
      return c ? `${c.prenom} ${c.nom}` : '';
    }).filter(Boolean).join(', ');
    const statut = getStatut(m);
    const badgeCls = statut === 'en_cours' ? 'badge-encours' : 'badge-terminee';
    const badgeLbl = statut === 'en_cours' ? 'En cours' : 'Terminée';
    const periode = [m.debut, m.fin].filter(Boolean).map(d => formatDate(d)).join(' → ');

    // Logo
    let logoContent;
    if (client && client.logo && isUrl(client.logo)) {
      const initFallback = (client.nom||'?').split(' ').map(w=>w[0]).join('').toUpperCase().slice(0,2);
      logoContent = `<img src="${client.logo}" onerror="this.style.display='none';this.parentElement.innerHTML='<span style=\\'font-weight:700;color:var(--accent);font-size:1rem\\'>${initFallback}</span>'">`;
    } else if (client) {
      const initiales = (client.nom || '?').split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
      logoContent = `<span style="font-weight:700;color:var(--accent);font-size:1rem">${initiales}</span>`;
    } else {
      logoContent = `<span style="color:var(--text-muted)">?</span>`;
    }

    return `
    <div class="mission-card" onclick="openEditMission('${m.id}')" style="cursor:pointer" title="Cliquer pour modifier">
      <div class="mission-card-header">
        <div class="mission-card-logo">${logoContent}</div>
        <div>
          <div class="mission-card-title">${m.titre || '—'}</div>
          ${client ? `<div class="mission-card-client">${client.nom}</div>` : ''}
        </div>
      </div>
      <span class="badge ${badgeCls}" style="align-self:flex-start">${badgeLbl}</span>
      ${collabs ? `<div class="mission-card-field">👤 <strong>${collabs}</strong></div>` : ''}
      ${periode ? `<div class="mission-card-field">📅 <strong>${periode}</strong></div>` : ''}
      ${(()=>{
        const tags = [
          ...(m.perimetreIds||[]).map(id => { const p = DB.perimetres.find(x=>x.id===id); return p ? `<span class="tag-pick selected" style="pointer-events:none">${p.nom}</span>` : ''; }),
          ...(m.methodeIds||[]).map(id => { const p = DB.methodes.find(x=>x.id===id); return p ? `<span class="tag-pick selected" style="pointer-events:none;background:rgba(74,200,255,0.15);border-color:var(--accent2);color:var(--accent2)">${p.nom}</span>` : ''; })
        ].filter(Boolean);
        return tags.length ? `<div class="tag-picker" style="margin-top:0.2rem">${tags.join('')}</div>` : '';
      })()}
      ${m.details ? `<div class="mission-card-details">${m.details}</div>` : ''}
    </div>`;
  }).join('');
}

// ══════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════
load().then(() => renderAll());
</script>
</body>
</html>
