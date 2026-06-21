// Leadership OS — Google Sheets sync endpoint
//
// SETUP:
// 1. Open your spreadsheet:
//    https://docs.google.com/spreadsheets/d/1jKGz295Gy5LuckY9GZSOXq9Xo4AHmA2bibKMssZWIvs/edit
// 2. Extensions > Apps Script
// 3. Delete everything in Code.gs and paste this file's contents in its place
// 4. Save (Cmd+S)
// 5. Deploy > Manage deployments > pick the existing deployment > pencil icon
//    Version: "New version" > Deploy
//    (the /exec URL stays the same — no need to update webAppUrl in the HTML)
//
// TABS MANAGED (created automatically with header rows if missing):
//   - "tasks"             Task, Quadrant, Owner, Done, DueDate, MilestoneID
//   - "ideas"             Idea, CaptureDate, Status, Notes
//   - "recipes"           Name, Protein, Calories, Icon
//   - "meal_plans"        PlanName, NumDays, Day, MealType, RecipeName
//   - "people"            PersonID, Name, Notes
//   - "areas"             AreaID, Name
//   - "aspirations"       AspirationID, Text, _LegacyRoleId, HorizonYears, StartDate, EndDate, LastUpdated, Area
//                         (col C is a legacy RoleId that is no longer written or returned;
//                          existing sheet data is left untouched — only cols A,B,D-H are read)
//   - "aspiration_roster" RosterID, AspirationID, PersonID, RoleType
//   - "milestones"        MilestoneId, AspId, Horizon, Text, MetricDef, MetricTarget,
//                         MetricCurrent, ProgressPct, Owner, DueDate, ParentMilestoneId, Status
//
// NOTE: The old "roles" tab is retired. Its Sheet data is left untouched but
//       the script no longer reads or writes it.

// ─── Tab names ───────────────────────────────────────────────────────────────

const TASKS_SHEET             = 'tasks';
const IDEAS_SHEET             = 'ideas';
const RECIPES_SHEET           = 'recipes';
const MEAL_PLANS_SHEET        = 'meal_plans';
const PEOPLE_SHEET            = 'people';
const AREAS_SHEET             = 'areas';
const ASPIRATIONS_SHEET       = 'aspirations';
const ASPIRATION_ROSTER_SHEET = 'aspiration_roster';
const MILESTONES_SHEET        = 'milestones';

// ─── Headers ─────────────────────────────────────────────────────────────────

const TASKS_HEADER             = ['Task', 'Quadrant', 'Owner', 'Done', 'DueDate', 'MilestoneID'];
const IDEAS_HEADER             = ['Idea', 'CaptureDate', 'Status', 'Notes'];
const RECIPES_HEADER           = ['Name', 'Protein', 'Calories', 'Icon'];
const MEAL_PLANS_HEADER        = ['PlanName', 'NumDays', 'Day', 'MealType', 'RecipeName'];
const PEOPLE_HEADER            = ['PersonID', 'Name', 'Notes'];
const AREAS_HEADER             = ['AreaID', 'Name'];
// Col C (_LegacyRoleId) is kept so existing data rows are not disturbed.
// New writes put '' there. Only cols A,B,D-H are consumed.
const ASPIRATIONS_HEADER       = ['AspirationID', 'Text', '_LegacyRoleId', 'HorizonYears', 'StartDate', 'EndDate', 'LastUpdated', 'Area'];
const ASPIRATION_ROSTER_HEADER = ['RosterID', 'AspirationID', 'PersonID', 'RoleType'];
const MILESTONES_HEADER        = ['MilestoneId', 'AspId', 'Horizon', 'Text', 'MetricDef', 'MetricTarget',
                                   'MetricCurrent', 'ProgressPct', 'Owner', 'DueDate', 'ParentMilestoneId', 'Status'];

// ─── Shared helpers ──────────────────────────────────────────────────────────

function getSheet(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
  return sheet;
}

function formatCell(val) {
  if (val != null && typeof val.getFullYear === 'function') {
    const y = val.getFullYear(),
          m = String(val.getMonth() + 1).padStart(2, '0'),
          d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return val == null ? '' : String(val);
}

function readRows(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Find row index (1-based, accounting for header) by matching a column value.
// Returns -1 if not found.
function findRowByCol(sheet, colIndex, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const vals = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (formatCell(vals[i][0]) === value) return i + 2;
  }
  return -1;
}

// ─── Row mappers ─────────────────────────────────────────────────────────────

function rowToTask(row) {
  return {
    task:        formatCell(row[0]),
    quadrant:    formatCell(row[1]),
    owner:       formatCell(row[2]),
    done:        formatCell(row[3]).toUpperCase() === 'TRUE',
    dueDate:     formatCell(row[4]),
    milestoneId: formatCell(row[5])   // blank for unlinked tasks; row[5] may be undefined for old rows
  };
}

function rowToIdea(row) {
  return {
    text:        formatCell(row[0]),
    captureDate: formatCell(row[1]),
    status:      formatCell(row[2]),
    notes:       formatCell(row[3])
  };
}

function rowToRecipe(row) {
  return {
    name:     formatCell(row[0]),
    protein:  formatCell(row[1]),
    calories: formatCell(row[2]),
    icon:     formatCell(row[3])
  };
}

function rowToMealPlan(row) {
  return {
    planName:   formatCell(row[0]),
    numDays:    formatCell(row[1]),
    day:        formatCell(row[2]),
    mealType:   formatCell(row[3]),
    recipeName: formatCell(row[4])
  };
}

function rowToPerson(row) {
  return {
    personId: formatCell(row[0]),
    name:     formatCell(row[1]),
    notes:    formatCell(row[2])
  };
}

function rowToArea(row) {
  return {
    areaId: formatCell(row[0]),
    name:   formatCell(row[1])
  };
}

// Aspirations: col 0=AspId, 1=Text, 2=_LegacyRoleId (skipped), 3=HorizonYears,
//              4=StartDate, 5=EndDate, 6=LastUpdated, 7=Area
function rowToAspiration(row) {
  return {
    aspirationId: formatCell(row[0]),
    text:         formatCell(row[1]),
    // col 2 is legacy RoleId — intentionally not returned
    horizonYears: formatCell(row[3]),
    startDate:    formatCell(row[4]),
    endDate:      formatCell(row[5]),
    lastUpdated:  formatCell(row[6]),
    area:         formatCell(row[7])   // row[7] may be undefined for old rows → ''
  };
}

function rowToRosterEntry(row) {
  return {
    rosterId:     formatCell(row[0]),
    aspirationId: formatCell(row[1]),
    personId:     formatCell(row[2]),
    roleType:     formatCell(row[3])
  };
}

function rowToMilestone(row) {
  return {
    milestoneId:       formatCell(row[0]),
    aspirationId:      formatCell(row[1]),
    horizon:           formatCell(row[2]),
    text:              formatCell(row[3]),
    metricDefinition:  formatCell(row[4]),
    metricTarget:      formatCell(row[5]),
    metricCurrent:     formatCell(row[6]),
    progressPct:       formatCell(row[7]),
    owner:             formatCell(row[8]),
    dueDate:           formatCell(row[9]),
    parentMilestoneId: formatCell(row[10]),
    status:            formatCell(row[11])
  };
}

// ─── Routing ─────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.tab === 'ideas')              return handleIdeasPost(body);
    if (body.tab === 'recipes')            return handleRecipesPost(body);
    if (body.tab === 'meal_plans')         return handleMealPlansPost(body);
    if (body.tab === 'people')             return handlePeoplePost(body);
    if (body.tab === 'areas')              return handleAreasPost(body);
    if (body.tab === 'aspirations')        return handleAspirationsPost(body);
    if (body.tab === 'aspiration_roster')  return handleRosterPost(body);
    if (body.tab === 'milestones')         return handleMilestonesPost(body);
    return handleTasksPost(body);
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message });
  }
}

// ─── Tasks ───────────────────────────────────────────────────────────────────

function handleTasksPost(body) {
  const sheet = getSheet(TASKS_SHEET, TASKS_HEADER);
  if (body.action === 'append') {
    const t = body.task || {};
    sheet.appendRow([
      t.task        || '',
      t.quadrant    || '',
      t.owner       || '',
      t.done ? 'TRUE' : 'FALSE',
      t.dueDate     || '',
      t.milestoneId || ''
    ]);
    return jsonOutput({ ok: true });
  }
  // 'sync': clear and rewrite the whole tab (also updates header to 6 cols)
  const tasks = body.tasks || [];
  sheet.clearContents();
  const rows = [TASKS_HEADER];
  tasks.forEach(t => rows.push([
    t.task        || '',
    t.quadrant    || '',
    t.owner       || '',
    t.done ? 'TRUE' : 'FALSE',
    t.dueDate     || '',
    t.milestoneId || ''
  ]));
  sheet.getRange(1, 1, rows.length, TASKS_HEADER.length).setValues(rows);
  return jsonOutput({ ok: true, count: tasks.length });
}

// ─── Ideas ───────────────────────────────────────────────────────────────────

function handleIdeasPost(body) {
  const sheet = getSheet(IDEAS_SHEET, IDEAS_HEADER);
  if (body.action === 'convert') {
    setIdeaStatus(sheet, body.match, 'Converted');
    if (body.task) handleTasksPost({ action: 'append', task: body.task });
    return jsonOutput({ ok: true });
  }
  if (body.action === 'release') {
    setIdeaStatus(sheet, body.match, 'Released');
    return jsonOutput({ ok: true });
  }
  const idea = body.idea || {};
  sheet.appendRow([
    idea.text        || '',
    idea.captureDate || '',
    idea.status      || 'Active',
    idea.notes       || ''
  ]);
  return jsonOutput({ ok: true });
}

function setIdeaStatus(sheet, match, status) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !match) return;
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (formatCell(rows[i][0]) === match.text &&
        formatCell(rows[i][1]) === match.captureDate) {
      sheet.getRange(i + 2, 3).setValue(status);
      return;
    }
  }
}

// ─── Recipes ─────────────────────────────────────────────────────────────────

function handleRecipesPost(body) {
  const sheet = getSheet(RECIPES_SHEET, RECIPES_HEADER);
  if (body.action === 'append') {
    const r = body.recipe || {};
    sheet.appendRow([r.name || '', r.protein || 0, r.calories || 0, r.icon || '']);
    return jsonOutput({ ok: true });
  }
  if (body.action === 'delete') {
    const rowIdx = findRowByCol(sheet, 1, body.name || '');
    if (rowIdx < 0) return jsonOutput({ ok: false, error: 'Recipe not found: ' + body.name });
    sheet.deleteRow(rowIdx);
    return jsonOutput({ ok: true });
  }
  return jsonOutput({ ok: false, error: 'Unknown action: ' + body.action });
}

// ─── Meal Plans ──────────────────────────────────────────────────────────────

function handleMealPlansPost(body) {
  const sheet = getSheet(MEAL_PLANS_SHEET, MEAL_PLANS_HEADER);
  const planRows = body.rows || [];
  sheet.clearContents();
  const rows = [MEAL_PLANS_HEADER];
  planRows.forEach(r => rows.push([
    r.planName || '', r.numDays || '', r.day || '', r.mealType || '', r.recipeName || ''
  ]));
  sheet.getRange(1, 1, rows.length, MEAL_PLANS_HEADER.length).setValues(rows);
  return jsonOutput({ ok: true, count: planRows.length });
}

// ─── People ──────────────────────────────────────────────────────────────────

function handlePeoplePost(body) {
  const sheet = getSheet(PEOPLE_SHEET, PEOPLE_HEADER);
  const p = body.person || {};

  if (body.action === 'append') {
    sheet.appendRow([p.personId || '', p.name || '', p.notes || '']);
    return jsonOutput({ ok: true });
  }

  if (body.action === 'update') {
    const rowIdx = findRowByCol(sheet, 1, p.personId || '');
    if (rowIdx < 0) {
      sheet.appendRow([p.personId || '', p.name || '', p.notes || '']);
    } else {
      sheet.getRange(rowIdx, 1, 1, PEOPLE_HEADER.length).setValues([[
        p.personId || '', p.name || '', p.notes || ''
      ]]);
    }
    return jsonOutput({ ok: true });
  }

  if (body.action === 'delete') {
    const rowIdx = findRowByCol(sheet, 1, body.personId || '');
    if (rowIdx >= 0) sheet.deleteRow(rowIdx);
    return jsonOutput({ ok: true });
  }

  return jsonOutput({ ok: false, error: 'Unknown action: ' + body.action });
}

// ─── Areas ───────────────────────────────────────────────────────────────────

function handleAreasPost(body) {
  const sheet = getSheet(AREAS_SHEET, AREAS_HEADER);
  const a = body.area || {};

  if (body.action === 'append') {
    sheet.appendRow([a.areaId || '', a.name || '']);
    return jsonOutput({ ok: true });
  }

  if (body.action === 'delete') {
    const rowIdx = findRowByCol(sheet, 1, body.areaId || '');
    if (rowIdx >= 0) sheet.deleteRow(rowIdx);
    return jsonOutput({ ok: true });
  }

  return jsonOutput({ ok: false, error: 'Unknown action: ' + body.action });
}

// ─── Aspiration Roster ───────────────────────────────────────────────────────

function handleRosterPost(body) {
  const sheet = getSheet(ASPIRATION_ROSTER_SHEET, ASPIRATION_ROSTER_HEADER);
  const r = body.rosterEntry || {};

  if (body.action === 'append') {
    sheet.appendRow([r.rosterId || '', r.aspirationId || '', r.personId || '', r.roleType || '']);
    return jsonOutput({ ok: true });
  }

  if (body.action === 'update') {
    const rowIdx = findRowByCol(sheet, 1, r.rosterId || '');
    if (rowIdx < 0) {
      sheet.appendRow([r.rosterId || '', r.aspirationId || '', r.personId || '', r.roleType || '']);
    } else {
      sheet.getRange(rowIdx, 1, 1, ASPIRATION_ROSTER_HEADER.length).setValues([[
        r.rosterId || '', r.aspirationId || '', r.personId || '', r.roleType || ''
      ]]);
    }
    return jsonOutput({ ok: true });
  }

  if (body.action === 'delete') {
    const rowIdx = findRowByCol(sheet, 1, body.rosterId || '');
    if (rowIdx >= 0) sheet.deleteRow(rowIdx);
    return jsonOutput({ ok: true });
  }

  return jsonOutput({ ok: false, error: 'Unknown action: ' + body.action });
}

// ─── Aspirations ─────────────────────────────────────────────────────────────

function handleAspirationsPost(body) {
  const sheet = getSheet(ASPIRATIONS_SHEET, ASPIRATIONS_HEADER);
  const a = body.aspiration || {};

  if (body.action === 'append') {
    sheet.appendRow([
      a.aspirationId || '', a.text || '', '',       // col C: legacy RoleId — always ''
      a.horizonYears || '', a.startDate || '', a.endDate || '', a.lastUpdated || '',
      a.area || ''
    ]);
    return jsonOutput({ ok: true });
  }

  if (body.action === 'update') {
    const rowIdx = findRowByCol(sheet, 1, a.aspirationId || '');
    if (rowIdx < 0) {
      sheet.appendRow([
        a.aspirationId || '', a.text || '', '',
        a.horizonYears || '', a.startDate || '', a.endDate || '', a.lastUpdated || '',
        a.area || ''
      ]);
    } else {
      sheet.getRange(rowIdx, 1, 1, ASPIRATIONS_HEADER.length).setValues([[
        a.aspirationId || '', a.text || '', '',
        a.horizonYears || '', a.startDate || '', a.endDate || '', a.lastUpdated || '',
        a.area || ''
      ]]);
    }
    return jsonOutput({ ok: true });
  }

  if (body.action === 'delete') {
    const rowIdx = findRowByCol(sheet, 1, body.aspirationId || '');
    if (rowIdx >= 0) sheet.deleteRow(rowIdx);
    return jsonOutput({ ok: true });
  }

  return jsonOutput({ ok: false, error: 'Unknown action: ' + body.action });
}

// ─── Milestones ──────────────────────────────────────────────────────────────

function handleMilestonesPost(body) {
  const sheet = getSheet(MILESTONES_SHEET, MILESTONES_HEADER);
  const m = body.milestone || {};

  if (body.action === 'append') {
    sheet.appendRow([
      m.milestoneId || '', m.aspirationId || '', m.horizon || '', m.text || '',
      m.metricDefinition || '', m.metricTarget || 0, m.metricCurrent || 0,
      m.progressPct || 0, m.owner || '', m.dueDate || '',
      m.parentMilestoneId || '', m.status || 'Active'
    ]);
    return jsonOutput({ ok: true });
  }

  if (body.action === 'update') {
    const rowIdx = findRowByCol(sheet, 1, m.milestoneId || '');
    if (rowIdx < 0) {
      sheet.appendRow([
        m.milestoneId || '', m.aspirationId || '', m.horizon || '', m.text || '',
        m.metricDefinition || '', m.metricTarget || 0, m.metricCurrent || 0,
        m.progressPct || 0, m.owner || '', m.dueDate || '',
        m.parentMilestoneId || '', m.status || 'Active'
      ]);
    } else {
      sheet.getRange(rowIdx, 1, 1, MILESTONES_HEADER.length).setValues([[
        m.milestoneId || '', m.aspirationId || '', m.horizon || '', m.text || '',
        m.metricDefinition || '', m.metricTarget || 0, m.metricCurrent || 0,
        m.progressPct || 0, m.owner || '', m.dueDate || '',
        m.parentMilestoneId || '', m.status || 'Active'
      ]]);
    }
    return jsonOutput({ ok: true });
  }

  if (body.action === 'delete') {
    const rowIdx = findRowByCol(sheet, 1, body.milestoneId || '');
    if (rowIdx >= 0) sheet.deleteRow(rowIdx);
    return jsonOutput({ ok: true });
  }

  return jsonOutput({ ok: false, error: 'Unknown action: ' + body.action });
}

// ─── doGet ───────────────────────────────────────────────────────────────────

function doGet(e) {
  const tasks            = readRows(getSheet(TASKS_SHEET,             TASKS_HEADER))
                             .filter(r => r[0]).map(rowToTask);
  const ideas            = readRows(getSheet(IDEAS_SHEET,             IDEAS_HEADER))
                             .filter(r => r[0]).map(rowToIdea);
  const recipes          = readRows(getSheet(RECIPES_SHEET,           RECIPES_HEADER))
                             .filter(r => r[0]).map(rowToRecipe);
  const mealPlans        = readRows(getSheet(MEAL_PLANS_SHEET,        MEAL_PLANS_HEADER))
                             .filter(r => r[0]).map(rowToMealPlan);
  const people           = readRows(getSheet(PEOPLE_SHEET,            PEOPLE_HEADER))
                             .filter(r => r[0]).map(rowToPerson);
  const areas            = readRows(getSheet(AREAS_SHEET,             AREAS_HEADER))
                             .filter(r => r[0]).map(rowToArea);
  const aspirations      = readRows(getSheet(ASPIRATIONS_SHEET,       ASPIRATIONS_HEADER))
                             .filter(r => r[0]).map(rowToAspiration);
  const aspirationRoster = readRows(getSheet(ASPIRATION_ROSTER_SHEET, ASPIRATION_ROSTER_HEADER))
                             .filter(r => r[0]).map(rowToRosterEntry);
  const milestones       = readRows(getSheet(MILESTONES_SHEET,        MILESTONES_HEADER))
                             .filter(r => r[0]).map(rowToMilestone);

  return jsonOutput({
    ok: true,
    tasks, ideas, recipes, mealPlans,
    people, areas, aspirations, aspirationRoster, milestones
  });
}
