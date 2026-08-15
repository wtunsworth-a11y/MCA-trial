// =============================================================================
// MCA Project Activity Tracker — Google Apps Script
// =============================================================================
// SETUP: Tools > Script editor, paste this file, run setupWorkbook() once.
// Set APPROVER_EMAIL below before running setup.
// =============================================================================

const CONFIG = {
  APPROVER_EMAIL:   'approver@example.com',   // ← change before setup
  SHEET_ACTIVITIES: 'Activities',
  SHEET_TASKS:      'Tasks',
  SHEET_PROGRESS:   'Progress Log',
  SHEET_BUDGET:     'Budget',
  SHEET_DASHBOARD:  'Dashboard',
};

// Exchange rate: PGK per 1 EUR (update here if rate changes)
const PGK_PER_EUR = 5;

// ---------------------------------------------------------------------------
// Master activity list (imported from Activities_v1.csv)
// Columns: [StrategicOutput, SubOutput, Code, Name, Scheduled, Details, Partner, CIStaff]
// ---------------------------------------------------------------------------
const ACTIVITIES = [
  ['SO1','1.1','1.1-1','Develop agreed data sets','Q4 2025','11 datasets identified (online and freely updated to minimise copyright and data ownership issues','CIFOR-ICRAF','Will'],
  ['SO1','1.1','1.1-2','Curate available data','Q4 2025','Capture all available data and reports. Digitise and store securely','CIFOR-ICRAF','Will'],
  ['SO1','1.1','1.1-3','Catalogue emerging datasets','Q2 2028','Capture emerging datasets and store them securely','CIFOR-ICRAF','Will'],
  ['SO1','1.1','1.1-4','Develop data sharing portals','Q3 2025','To agree storage locations and access rights for OPG and Community','CIFOR-ICRAF','Will'],
  ['SO1','1.1','1.1-5','Support development of Provincial Sustainable Land Use Plan','Q2 2026','DLPP policy action for implementation at Provincial level. Combines data from other sources','NPCB','Will'],
  ['SO1','1.1','1.1-6','Support strategic management of intellectual property (including traditional ecological knowledge)','Q4 2026','May require policy work, but needs to be in place to ensure safeguards for us and future development partners','CIFOR-ICRAF','Will'],
  ['SO1','1.1','1.1-7','Total Economic Valuation for Managalas Conservation Area','Q4 2025','Evaluate the capital asset of MCA. Enables engagement with private sector for future funding opportunities','RSD','Will'],
  ['SO1','1.1','1.1-8','Develop an early warning system for flooding at Pongani wet crossing','Q2 2026','Use rain gauges and water level to predict flooding at Pongani wet crossing and enable travellers to avoid danger','ALS','Will'],
  ['SO1','1.2','1.2-1','Initiate Oro FCCB PCMC Natural Resource Management Committee','Q3 2025','CEPA and CCDA legal and policy actions to implement at provincial level','NPCB','Will'],
  ['SO1','1.2','1.2-2','Re-establish Provincial Forest Management Committee','Q4 2025','PNGFA legal actions to implement at Provincial level','NPCB','Will'],
  ['SO1','1.2','1.2-3','Update the Provincial Forest Management Plan','Q2 2026','PNGFA legal actions to implement at Provincial level','CIFOR-ICRAF','Will'],
  ['SO1','1.2','1.2-4','METT Assessment for Managalas','Q2 2025','Understand the baseline of conservation in MCA','CIFOR-ICRAF','Will'],
  ['SO1','1.2','1.2-5','METT Assessment for other PAs in Oro (Hombareta and Lejo)','Q2 2026','Comparison for other PAs in Oro as a basis for OPG interventions','NPCB','Will'],
  ['SO1','1.2','1.2-6','Generate social inclusion safeguards for FCCB related issues','Q1 2026','Needs to be in place to ensure safeguards for us and future development partners','CIFOR-ICRAF','Linda'],
  ['SO1','1.2','1.2-7','Review existing (National and Provincial) policy for provincial implementation','Q1 2025','Workshops held with OPG and OPA to identify the 12 activities','Lester','Will'],
  ['SO1','1.2','1.2-8','NPCB Act requires a Biodiversity Conservation Policy','Q4 2025','Provincial legal activity to implement','NPCB','Will'],
  ['SO1','1.2','1.2-9','Develop Provincial Policy for PES within existing national policy frameworks','Q4 2025','Requirement for other work under SO3 in seeking sustainable financing','NPCB','Will'],
  ['SO1','1.2','1.2-10','Support law enforcement for FCCB related activities','Q2 2026','Improve environmental safeguards across the Province','JR','Will'],
  ['SO1','1.2','1.2-11','Develop a draft Provincial Alluvial Mining Policy for Oro','Q1 2026','Identified as a need during baseline surveys in MCA based on the unregulated alluvial mining activities','SAMS','Will'],
  ['SO1','1.3','1.3-1','Strengthen the institutional landscapes of NPCB, MCF, clan and Wards','Q2 2028','Key institutions for decision making to be strengthened','various','Will'],
  ['SO1','1.3','1.3-2','Assess key market features relevant to PES in the Oro Province','Q4 2025','To better understand the available market for sustainable financing of conservation','RSD','Will'],
  ['SO1','1.3','1.3-3','Provide awareness on PES among key stakeholders','Q4 2025','To ensure provincial and community understanding required for sustainable financing','NPCB','Will'],
  ['SO1','1.3','1.3-4','Strengthen capacity for community empowerment (legal support)','Q4 2025','Ensure that landowners are aware of their rights and the project team is not overstepping','Paula','Will'],
  ['SO1','1.4','1.4-1','Review of existing Ecosystem Service trading systems encountered in the region','Q3 2024','Learn from other sustainable financing mechanisms in the region','Himlal','Will'],
  ['SO1','1.4','1.4-2','Communicate the value proposition around Ecosystem Services and trading system','Q2 2026','Ensure realistic expectations for what sustainable financing can deliver','RSD','Will'],
  ['SO1','1.4','1.4-3','Document potential utilization of ES trading system by both the private and public sectors in the Oro Province','Q2 2026','Additional use cases for sustainable financing in Oro','CIFOR-ICRAF','Will'],
  ['SO1','1.4','1.4-4','Assessment of potential contributors to PES','Q2 2026','PES is not one thing and we expect a matrix of incomes supporting MCA and Oro','RSD','Will'],
  ['SO1','1.5','1.5-1','Develop CAMP (Conservation Area Management Plan)','Q3 2025','Key document for communicating Conservation in MCA to future partners','CIFOR-ICRAF','Will'],
  ['SO1','1.5','1.5-2','Undertake clan boundary mapping with clans to support clan level implementation of CAMP','Q2 2026','Requested in clan elder workshops, will be used to ensure that boundary is up to date and widely accepted','CIFOR-ICRAF','Mellie'],
  ['SO1','1.5','1.5-3','Support CAMC (Conservation Area Management Committee)','Q2 2028','CAMC is a legal requirement for PAs','Paula','Will'],
  ['SO1','1.5','1.5-4','Develop a monitoring system for MCA','Q4 2025','','CIFOR-ICRAF','Rhett'],
  ['SO1','1.5','1.5-5','Training for Clan Stewards (2 per clan) to support clans in implementing CAMP','Q3 2026','Clan Stewards embedded in each clan to ensure technical capacity to deliver','CIFOR-ICRAF','Mellie'],
  ['SO1','1.5','1.5-6','Develop performing team of Rangers','Q4 2026','Popular request on site for gazetted Rangers. To address their roles','CIFOR-ICRAF','Mellie'],
  ['SO1','1.5','1.5-7','Register MCA under the IUCN Green List of Protected Areas','Q2 2027','Opportunity for raising the profile of MCA as an internationally recognised performing PA','CIFOR-ICRAF','Rhett'],
  ['SO1','1.5','1.5-8','Review CAMP in line with monitoring results and agreed sustainable business practices','Q2 2028','Update CAMP before end of the project','CIFOR-ICRAF','Mellie'],
  ['SO1','1.5','1.5-9','Prepare Species Management Plans for target species','Q2 2027','More focused management of identified priority species','CIFOR-ICRAF','Rhett'],
  ['SO2','2.1','2.1-1','Raise awareness within MCA on the importance of Conservation','Q2 2028','Ongoing awareness programmes in MCA through community workshops and targeted activities','CIFOR-ICRAF','Satia'],
  ['SO2','2.1','2.1-2','Engage with media & support communications activities for MCA and FCCB','Q2 2028','Ensure positive news stories about Managalas are flowing to the National and Provincial media','CIFOR-ICRAF','Satia'],
  ['SO2','2.1','2.1-3','NFI assessment on three clusters in Managalas','Q4 2026','Part of the Biodiversity survey; National Forest Inventory clusters','CIFOR-ICRAF','Rhett'],
  ['SO2','2.1','2.1-4','Engage schools to include MCA within teaching','Q3 2026','Establish regular meeting with teachers to engage with schools','Kerry','Will'],
  ['SO2','2.1','2.1-5','Develop Managalas specific teaching materials in line with National Curriculum','Q2 2026','Develop teaching materials for schools with location specific messaging','Kerry','Will'],
  ['SO2','2.2','2.2-1','Develop network of tertiary institutions to utilise the plateau','Q2 2028','Part of developing Research as an economic opportunity','CIFOR-ICRAF','Will'],
  ['SO2','2.2','2.2-2','Identify and support research projects for students','Q2 2028','Create opportunities for Post-graduate and Under-graduate students to undertake research within Managalas','CIFOR-ICRAF','Will'],
  ['SO2','2.2','2.2-3','Support FCCB-related research works with a particular focus on PNG researchers','Q2 2028','Partnerships with PNG Universities. One Mphil student completed fieldwork. Seeking 9 in 2026','CIFOR-ICRAF','Will'],
  ['SO2','2.2','2.2-4','Develop methods for population monitoring of QABB in Managalas','Q2 2026','Part of developing QABB as a legal economic opportunity; requires consistent population monitoring method','Christine','Will'],
  ['SO2','2.2','2.2-5','Convene a colloquium on Queen Alexandras Birdwing Butterfly to identify a research agenda','Q4 2025','Update all available information on QABB and ensure copies are held in Oro. Develop a targeted research agenda','CIFOR-ICRAF','Will'],
  ['SO2','2.2','2.2-6','Biodiversity Assessment of Managalas Conservation Area','Q2 2027','Comprehensive Biodiversity assessment of MCA using multiple methods (CT, PAM, eDNA, TEK plus NFI and PSP)','CIFOR-ICRAF','Rhett'],
  ['SO2','2.2','2.2-7','Establish nine Permanent Sample Plots in Managalas as sites for future research work','Q4 2026','Through partnership with PNGFRI establish 9 Permanent Sample Plots for forest researchers to continue to use in the future','CIFOR-ICRAF','Rhett'],
  ['SO2','2.2','2.2-8','Assess and act on reintroduction of Goura Pigeon to MCA following extirpation through hunting','Q4 2027','Reintroduce a species lost to overhunting as part of awareness on hunting and impacts on the environment','CIFOR-ICRAF','Will'],
  ['SO2','2.3','2.3-1','Support Benson Garui Scholarship for MCA population','Q2 2028','21 students supported in 2024. System reviewed in 2024. 45 students supported in 2025','CIFOR-ICRAF','Will'],
  ['SO2','2.3','2.3-2','Develop teacher networks with schools serving the MCA community','Q3 2025','','Kerry','Will'],
  ['SO2','2.4','2.4-1','Perform training need assessment on FCCB among target groups and provide support for FCCB related training courses','Q2 2024','Identify target groups and enable access to training','RSD','Will'],
  ['SO2','2.4','2.4-2','Design and support FCCB training courses for public servants and communities','Q4 2025','Design training courses for communities and public servants to learn more about community based conservation','TBD','Will'],
  ['SO2','2.4','2.4-3','Assess training effectiveness and scalability','Q4 2026','Undertake assessment of training provided for effectiveness','TBD','Will'],
  ['SO2','2.4','2.4-4','Review existing decision-making processes relevant to FCCB in the Oro Province','Q4 2025','Identify decision making bodies within Oro and MCA that affect natural resources','RSD','Will'],
  ['SO2','2.4','2.4-5','Enhance social inclusion with conservation planning','Q2 2028','Research to understand and enhance social inclusion','CIFOR-ICRAF','Linda'],
  ['SO2','2.4','2.4-6','Provide the valuation of empowering women and youth + traditional landowners in decision-making processes','Q4 2026','Research into the value of increasing inclusion in decision making processes','TBD','TBD'],
  ['SO2','2.4','2.4-7','Assess new decision-making processes effectiveness, efficiency, impact and scalability','Q4 2027','Undertake an assessment of the impact of enhanced structures and processes for decision making','TBD','TBD'],
  ['SO3','3.1','3.1-1','Assess bottlenecks and opportunities of existing + innovative FCCB-related value chains','Q3 2025','Clarify the issues that value chains in CA are facing','TBD','TBD'],
  ['SO3','3.1','3.1-2','Assess initiatives leading to evidence-based decision making','Q2 2027','Assessment of how farmers make decisions and how uptake of new technology can be enhanced','TBD','TBD'],
  ['SO3','3.1','3.1-3','Develop extension materials, assess group dynamics and provide training for improved coffee quality','Q2 2026','Prepare manuals and guides for all livelihood activities to ensure that work can be replicated and scaled up','TBD','Will'],
  ['SO3','3.1','3.1-4','Co-develop activities to engage youth groups in coffee farming with DDA and LLG','Q2 2026','Seeking means to engage youth in agriculture','TBD','Will'],
  ['SO3','3.1','3.1-5','Prepare a business case for downstream processing of coffee in Oro Province','Q4 2025','Business plan for a coffee processing (green bean) industry in Oro sufficient to engage with potential investors','TBD','Will'],
  ['SO3','3.1','3.1-6','Employ extension workers in MCA to offer hands-on training for Vanilla','Q2 2028','Field officers trained and employed to work in their own communities to support vanilla production','KL','Will'],
  ['SO3','3.1','3.1-7','Explore opportunities for a single source market for Managalas produced vanilla','Q4 2025','Single source products from sites with conservation impacts can attract higher prices','KL','Will'],
  ['SO3','3.1','3.1-8','Create a cocoa nursery and provide cocoa extension services in Managalas and adjacent LLGs','Q3 2025','Agreement with Cocoa Board to engage and develop Provincial Cocoa Nursery','CB','Will'],
  ['SO3','3.1','3.1-9','Undertake commodity mapping for EUDR compliance assessment in Managalas','Q2 2026','Vanilla, Coffee and Cocoa plots mapped and process ongoing','KL','Will'],
  ['SO3','3.1','3.1-10','Determine sustainable levels of forest product extraction in key value chains','Q2 2026','Determine how much (Massoy, firewood, building materials and other forest products) can be taken from the forest without damaging the forest','CIFOR-ICRAF','Josh'],
  ['SO3','3.1','3.1-11','Establish model farms for fresh vegetable supply to Popondetta','Q3 2025','Model farms installed with the intention to provide vegetables to Popondetta town','KL','Will'],
  ['SO3','3.1','3.1-12','Explore and recommend Food and Tree Crop Calendars for enhanced nutrition and incomes','Q4 2025','Earlier household survey information to be analysed for nutrition and income distribution through the year','CIFOR-ICRAF','Will'],
  ['SO3','3.1','3.1-13','Market analysis and product development for Okari nuts','Q3 2027','Exploration of the value chain for Okari nuts (Terminalia kaernbacchii)','Griffith','Will'],
  ['SO3','3.1','3.1-14','Prepare a feasibility study for Balsa (cultivation and processing) in Oro Province','Q3 2025','Assessment of balsa as a potential crop for Oro Province','Nalau','Will'],
  ['SO3','3.1','3.1-15','Establish model farms for livestock related farming as an offset to hunting and a boost to nutrition','Q4 2026','Establish small producer groups to cooperatively produce eggs, chickens, ducks, fish or pigs','CIFOR-ICRAF','Will'],
  ['SO3','3.1','3.1-16','Undertake an assessment of alluvial gold mining technologies in MCA, propose steps to mitigate impacts','Q1 2026','New Activity developed based on Household Survey, Clan Elder Workshops and observed level of alluvial mining','SAMS','Will'],
  ['SO3','3.1','3.1-17','Assess options for high value ginger or turmeric production in MCA','Q3 2026','Potential economic opportunity. Requires lab assessment for access to high value markets','KL','Will'],
  ['SO3','3.1','3.1-18','Assess options for high value stingless bee (Meliponideae) honey production in MCA','Q3 2026','Developed based on the unexplored potential for stingless bee honey in PNG','Meli','Will'],
  ['SO3','3.1','3.1-19','Prepare a tourism management and plan for Managalas Conservation Area','Q4 2025','Will evolve into broader tourism assessment for Oro Province','Peters','Will'],
  ['SO3','3.1','3.1-20','Document and promote best practices and enhanced processes within key value chains','Q4 2025','Document best practice manuals for farmers in Managalas to achieve quantity and quality production','Howard','Will'],
  ['SO3','3.1','3.1-21','Support the Provincial Government to prepare a Green Growth Strategy for Oro Province','Q2 2026','Support OPG to prepare a green growth strategy for the Oro Province','RSD','Will'],
  ['SO3','3.1','3.1-22','Undertake a study on development pathways of Oro Province and their outcomes','Q4 2025','Multi-dimensional assessment of outcomes (environmental, economic, social) of various development pathways in Oro Province','CIFOR-ICRAF','Karl'],
  ['SO3','3.1','3.1-23','Improve access to quality planting materials','Q2 2028','Develop nurseries and seed supply chains','CIFOR-ICRAF','Will'],
  ['SO3','3.1','3.1-24','Assess and Act on introduction of fuel efficient cookstoves','Q3 2026','Alternative cooking tech to reduce firewood use from the forest','FGV','Will'],
  ['SO3','3.1','3.1-25','Develop Branding for Managalas goods and services','Q4 2026','A strong branding and narrative for goods and services from Managalas will support the value proposition in the open market','CIFOR-ICRAF','Will'],
  ['SO3','3.2','3.2-1','Assess market access and financial performances of existing MSMEs and new entrants','Q2 2026','Assessment of SME performance to gauge training needs','KL','Will'],
  ['SO3','3.2','3.2-2','Assess capital requirements and human resource development efforts for enhancing value chains','Q2 2026','Design of programs to enhance value chains','RSD','Will'],
  ['SO3','3.2','3.2-3','Support the development of sustainable business models and plans','Q2 2026','Prepare business plans for viable value chains to enable financing as required','CIFOR-ICRAF','Will'],
  ['SO3','3.2','3.2-4','Determine the needs for green entrepreneurship support within key value chains','Q2 2025','From assessment, identify support mechanisms for SMEs','RSD','Will'],
  ['SO3','3.2','3.2-5','Review existing host capacity for Incubator/Accelerator (I/A) services in the Oro Province','Q2 2025','Assessment undertaken but not very favourable; subsequent activities cancelled','RSD','Will'],
  ['SO3','3.2','3.2-6','Support the set-up of the Incubator/Accelerator (I/A) functions','','This activity pending further review and discussion','RSD','Will'],
  ['SO3','3.2','3.2-7','Document the making of the Incubator/Accelerator (I/A) capacity in the Oro province','','','RSD','Will'],
  ['SO3','3.2','3.2-8','Communicate and promote thinking & action opportunities provided by the Incubator/Accelerator (I/A) services','','','RSD','Will'],
  ['SO3','3.2','3.2-9','Support women and youth willing to engage in green entrepreneurship','Q2 2026','Identify specific programs to enhance engagement in green economy','CIFOR-ICRAF','Will'],
  ['SO3','3.2','3.2-10','Facilitate access to knowledge, capital and niche market outlets','Q2 2028','Using business plans, engage with financial institutions to mobilise capital','RSD','Will'],
  ['SO3','3.3','3.3-1','Assess potential Ecosystem Services of the MCA','Q4 2025','','RSD','Will'],
  ['SO3','3.3','3.3-2','Identify modalities for benefit sharing','Q4 2026','','CIFOR-ICRAF','Will'],
  ['SO3','3.3','3.3-3','Assess available and potential market for PES','Q4 2025','Market research into buyers interests and requirements ongoing','Howard','Will'],
  ['SO3','3.3','3.3-4','Document and promote best practices for community management of PES','Q4 2025','','CIFOR-ICRAF','Will'],
  ['SO3','3.4','3.4-1','Install internet nodes in MCA (including power)','Q4 2025','Wireless network across Managalas to be installed and linked to internet','Richard','Will'],
  ['SO3','3.4','3.4-3','Undertake an impact assessment on gravel extraction methods for road construction and maintenance within MCA','Q4 2025','Roads are important but can be damaging. Assessment on impacts and mitigation options','Narua','Will'],
  ['SO3','3.4','3.4-4','Assess options for power generation from renewable sources of energy','Q4 2025','Hydropower was originally intended, but transmission is difficult. Distributed solar is more likely now','TBD','Will'],
  ['SO3','3.4','3.4-5','Support the development of power generating capacity (solar? Hydropower?) at key sites','Q4 2026','Installation of decisions from above','TBD','Will'],
  ['SO3','3.4','3.4-6','Contribute to infrastructure development and equipment acquisition relevant to green economy activity','Q4 2026','Intention is for three secure storage locations across Managalas (including cold storage)','TBD','Will'],
  ['SO3','3.4','3.4-7','Investment in protein producing livelihoods','Q4 2026','Fishponds (20), Piggery (11), Poultry (eggs, ducks or chickens) (22), Vegetable farms (5)','CIFOR-ICRAF','Will'],
  ['SO3','3.4','3.4-8','Facilitate access to capital investment for improved value chains','Q1 2027','Monitor usage of infrastructure','RSD','Will'],
];

// =============================================================================
// SETUP FUNCTIONS
// =============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MCA Tracker')
    .addItem('Initial Setup (run once)', 'setupWorkbook')
    .addSeparator()
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('View Gantt Chart', 'showGanttChart')
    .addSeparator()
    .addItem('Add Tasks to an Activity', 'showAddTaskDialog')
    .addItem('Log Progress on a Task', 'showProgressDialog')
    .addToUi();
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupActivitiesSheet(ss);
  setupTasksSheet(ss);
  setupProgressSheet(ss);
  setupBudgetSheet(ss);
  setupDashboardSheet(ss);
  // Move sheets into logical order
  const order = [CONFIG.SHEET_DASHBOARD, CONFIG.SHEET_ACTIVITIES, CONFIG.SHEET_TASKS, CONFIG.SHEET_PROGRESS, CONFIG.SHEET_BUDGET];
  order.reverse().forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.moveActiveSheet(ss.setActiveSheet(sh).getIndex()), ss.moveActiveSheet(1);
  });
  SpreadsheetApp.getUi().alert(
    'Setup complete!\n\n' +
    'Next steps:\n' +
    '1. Share this sheet with your team (View/Comment for data entry via the form)\n' +
    '2. The approver email is set to: ' + CONFIG.APPROVER_EMAIL + '\n' +
    '3. Have activity owners add tasks via MCA Tracker > Add Tasks to an Activity\n' +
    '4. Team members can then log progress via MCA Tracker > Log Progress on a Task'
  );
}

// ---------------------------------------------------------------------------
// ACTIVITIES SHEET
// ---------------------------------------------------------------------------
function setupActivitiesSheet(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_ACTIVITIES);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_ACTIVITIES);
  sh.clearContents();
  sh.clearFormats();

  const headers = [
    'Strategic Output', 'Sub-Output', 'Activity Code', 'Activity Name',
    'Scheduled Completion', 'Details', 'Responsible Partner', 'CI Staff Oversight',
    'Tasks Defined', '% Complete (Approved)', 'Status'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sh, headers.length);

  // Write activity data
  sh.getRange(2, 1, ACTIVITIES.length, 8).setValues(ACTIVITIES);

  // Freeze header, set column widths
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 350);
  sh.setColumnWidth(5, 140);
  sh.setColumnWidth(6, 300);
  sh.setColumnWidth(7, 140);
  sh.setColumnWidth(8, 120);
  sh.setColumnWidth(9, 100);
  sh.setColumnWidth(10, 160);
  sh.setColumnWidth(11, 100);

  // No sheet-level protection — editors can add new activities directly.
  // The warning-only mode is intentionally omitted so rows can be appended freely.

  // Colour rows by Strategic Output
  colorActivitiesByOutput(sh);
}

function colorActivitiesByOutput(sh) {
  const lastRow = sh.getLastRow();
  const data = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  const colors = { 'SO1': '#d9ead3', 'SO2': '#dae8fc', 'SO3': '#fff2cc' };
  data.forEach((row, i) => {
    const color = colors[row[0]] || '#ffffff';
    sh.getRange(i + 2, 1, 1, 11).setBackground(color);
  });
}

// ---------------------------------------------------------------------------
// TASKS SHEET
// ---------------------------------------------------------------------------
function setupTasksSheet(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_TASKS);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_TASKS);
  sh.clearContents();
  sh.clearFormats();

  const headers = [
    'Task Code', 'Activity Code', 'Strategic Output', 'Sub-Output', 'Activity Name',
    'Task Name', 'Task Description', 'Assigned To',
    'Planned Start', 'Planned End', 'Weight (%)',
    'Created By', 'Created Date', 'Status', 'Depends On (Task Code)'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sh, headers.length);

  // Dropdown for Activity Code (col 2) — references the Activities sheet live
  // so new activities added to that sheet automatically appear here
  const actSh = ss.getSheetByName(CONFIG.SHEET_ACTIVITIES);
  const actRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(actSh.getRange('C2:C500'), true)
    .setAllowInvalid(true) // warn but allow manual entry for activities added before refresh
    .build();
  sh.getRange(2, 2, 500, 1).setDataValidation(actRule);

  // Dropdown for Status (col 14)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Not Started', 'In Progress', 'Completed', 'On Hold', 'Cancelled'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, 14, 500, 1).setDataValidation(statusRule);

  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 90);
  sh.setColumnWidth(5, 300);
  sh.setColumnWidth(6, 250);
  sh.setColumnWidth(7, 300);
  sh.setColumnWidth(8, 130);
  sh.setColumnWidth(9, 110);
  sh.setColumnWidth(10, 110);
  sh.setColumnWidth(11, 90);
  sh.setColumnWidth(12, 120);
  sh.setColumnWidth(13, 110);
  sh.setColumnWidth(14, 100);
}

// ---------------------------------------------------------------------------
// PROGRESS LOG SHEET
// ---------------------------------------------------------------------------
function setupProgressSheet(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_PROGRESS);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_PROGRESS);
  sh.clearContents();
  sh.clearFormats();

  const headers = [
    'Entry ID', 'Task Code', 'Activity Code', 'Activity Name',
    'Submitted By', 'Submission Date', '% Complete', 'Status Update',
    'Evidence / Notes', 'Supporting Link',
    'Approval Status', 'Approved By', 'Approval Date', 'Approver Notes'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sh, headers.length);

  // Dropdown for Approval Status (col 11) — only approver should use this
  const approvalRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Rejected'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, 11, 2000, 1).setDataValidation(approvalRule);

  // % Complete validation (col 7)
  const pctRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 100)
    .setHelpText('Enter a number between 0 and 100')
    .build();
  sh.getRange(2, 7, 2000, 1).setDataValidation(pctRule);

  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 80);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 280);
  sh.setColumnWidth(5, 140);
  sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 90);
  sh.setColumnWidth(8, 300);
  sh.setColumnWidth(9, 250);
  sh.setColumnWidth(10, 200);
  sh.setColumnWidth(11, 120);
  sh.setColumnWidth(12, 120);
  sh.setColumnWidth(13, 120);
  sh.setColumnWidth(14, 250);

  // Conditional formatting: colour by approval status
  const pendingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$K2="Pending"')
    .setBackground('#fff2cc')
    .setRanges([sh.getRange('A2:N2000')])
    .build();
  const approvedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$K2="Approved"')
    .setBackground('#d9ead3')
    .setRanges([sh.getRange('A2:N2000')])
    .build();
  const rejectedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$K2="Rejected"')
    .setBackground('#f4cccc')
    .setRanges([sh.getRange('A2:N2000')])
    .build();
  sh.setConditionalFormatRules([pendingRule, approvedRule, rejectedRule]);
}

// ---------------------------------------------------------------------------
// DASHBOARD SHEET
// ---------------------------------------------------------------------------
function setupDashboardSheet(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_DASHBOARD);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_DASHBOARD);
  sh.clearContents();
  sh.clearFormats();

  const sectionStyle = (range) => range.setFontWeight('bold').setFontColor('#ffffff').setBackground('#1c4587');

  // ---- TITLE (rows 1-2) ----
  sh.getRange('A1').setValue('MCA PROJECT TRACKER — DASHBOARD');
  sh.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('G1').setFormula('=NOW()');
  sh.getRange('G1').setNumberFormat('dd MMM yyyy HH:mm').setFontColor('#666666');
  sh.getRange('A1:G1').setBackground('#c9daf8');
  sh.getRange('A2').setValue('Only APPROVED progress entries are reflected in the statistics below.');
  sh.getRange('A2').setFontColor('#cc0000').setFontStyle('italic');

  // ---- SECTION 1: SUMMARY BY STRATEGIC OUTPUT (rows 4-8) ----
  // Uses COUNTIF/SUMPRODUCT/AVERAGEIF — fixed output, no overflow risk
  let row = 4;
  sectionStyle(sh.getRange(row, 1, 1, 5).merge());
  sh.getRange(row, 1).setValue('SUMMARY BY STRATEGIC OUTPUT');
  row++;
  sh.getRange(row, 1, 1, 5).setValues([['Strategic Output','Total Activities','With Tasks Defined','Avg % Complete','Status']]);
  sh.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  [['SO1','SO1 – Policy & Governance'],['SO2','SO2 – Awareness & Research'],['SO3','SO3 – Livelihoods & Economy']].forEach(([so, label]) => {
    sh.getRange(row, 1).setValue(label);
    sh.getRange(row, 2).setFormula(`=COUNTIF(Activities!A:A,"${so}")`);
    sh.getRange(row, 3).setFormula(`=SUMPRODUCT((Activities!A$2:A$500="${so}")*(Activities!I$2:I$500>0))`);
    sh.getRange(row, 4).setFormula(`=IFERROR(AVERAGEIF(Activities!A$2:A$500,"${so}",Activities!J$2:J$500),0)`);
    sh.getRange(row, 4).setNumberFormat('0%');
    sh.getRange(row, 5).setFormula(`=IF(D${row}=0,"Not Started",IF(D${row}<0.5,"In Progress",IF(D${row}<1,"Nearly Complete","Complete")))`);
    row++;
  });

  // ---- SECTION 2: SCHEDULE STATUS BY QUARTER (rows 11-30) ----
  // Uses SUMPRODUCT — fixed 16 rows, no overflow risk
  row = 11;
  sectionStyle(sh.getRange(row, 1, 1, 5).merge());
  sh.getRange(row, 1).setValue('SCHEDULE STATUS — Activities by Quarter');
  row++;
  sh.getRange(row, 1, 1, 5).setValues([['Quarter','Total','Complete','In Progress','Not Started']]);
  sh.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  ['Q1 2025','Q2 2025','Q3 2025','Q4 2025','Q1 2026','Q2 2026','Q3 2026','Q4 2026',
   'Q1 2027','Q2 2027','Q3 2027','Q4 2027','Q1 2028','Q2 2028','Q3 2028','Q4 2028'].forEach(q => {
    sh.getRange(row, 1).setValue(q);
    sh.getRange(row, 2).setFormula(`=COUNTIF(Activities!E:E,"${q}")`);
    sh.getRange(row, 3).setFormula(`=SUMPRODUCT((Activities!E$2:E$500="${q}")*(Activities!K$2:K$500="Complete"))`);
    sh.getRange(row, 4).setFormula(`=SUMPRODUCT((Activities!E$2:E$500="${q}")*(Activities!K$2:K$500="In Progress"))`);
    sh.getRange(row, 5).setFormula(`=SUMPRODUCT((Activities!E$2:E$500="${q}")*(Activities!K$2:K$500="Not Started"))`);
    row++;
  });

  // ---- SECTION 3: PENDING APPROVALS (row 33 onwards) ----
  // QUERY can expand freely — nothing placed below until row 200
  row = 33;
  sectionStyle(sh.getRange(row, 1, 1, 6).merge());
  sh.getRange(row, 1).setValue('PENDING APPROVALS');
  row++;
  sh.getRange(row, 1, 1, 6).setValues([['Entry ID','Task Code','Activity Code','Submitted By','Date','% Complete']]);
  sh.getRange(row, 1, 1, 6).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  sh.getRange(row, 1).setFormula(
    `=IFERROR(QUERY('Progress Log'!A:N,"SELECT A,B,C,E,F,G WHERE K='Pending' ORDER BY F DESC",0),"No pending entries")`
  );

  // ---- SECTION 4: ACTIVITY PROGRESS (row 200 onwards) ----
  // QUERY returns ~95 rows — nothing placed below until row 350
  row = 200;
  sectionStyle(sh.getRange(row, 1, 1, 7).merge());
  sh.getRange(row, 1).setValue('ACTIVITY PROGRESS (Approved entries only)');
  row++;
  sh.getRange(row, 1, 1, 7).setValues([['Code','Activity','Partner','CI Staff','Scheduled','Tasks','% Complete']]);
  sh.getRange(row, 1, 1, 7).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  sh.getRange(row, 1).setFormula(
    `=IFERROR(QUERY(Activities!A:K,"SELECT C,D,G,H,E,I,J WHERE C IS NOT NULL ORDER BY A,C",0),"Run Refresh Dashboard first")`
  );

  // ---- SECTION 5: BY RESPONSIBLE PARTNER (row 350 onwards) ----
  // QUERY expands freely — nothing below
  row = 350;
  sectionStyle(sh.getRange(row, 1, 1, 3).merge());
  sh.getRange(row, 1).setValue('PROGRESS BY RESPONSIBLE PARTNER');
  row++;
  sh.getRange(row, 1, 1, 3).setValues([['Partner','Activities','Avg % Complete']]);
  sh.getRange(row, 1, 1, 3).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  sh.getRange(row, 1).setFormula(
    `=IFERROR(QUERY(Activities!G$2:K$500,"SELECT G,COUNT(G),AVG(J) WHERE G IS NOT NULL GROUP BY G ORDER BY G",0),"")`
  );

  // ---- SECTION 6: BUDGET SUMMARY (row 420 onwards — well clear of partner QUERY) ----
  row = 420;
  sectionStyle(sh.getRange(row, 1, 1, 8).merge());
  sh.getRange(row, 1).setValue('BUDGET SUMMARY  |  Exchange rate: PGK ' + PGK_PER_EUR + ' = EUR 1  |  All totals shown in EUR equivalent');
  row++;

  // 6a: By Strategic Output × Year
  sh.getRange(row, 1, 1, 5).setValues([['Strategic Output','Yr 4 (Jun26–May27) EUR','Yr 5 (Jun27–May28) EUR','Yr 6 (Jun28–May29) EUR','Total EUR']]);
  sh.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  [['SO1','SO1 – Policy & Governance'],['SO2','SO2 – Awareness & Research'],['SO3','SO3 – Livelihoods & Economy']].forEach(so => {
    const [code, label] = so;
    sh.getRange(row, 1).setValue(label);
    // SUMIF on Budget col A (SO) × col N (EUR equiv) × col J/K/L (years)
    // EUR equivalent = col N; year breakdown needs separate calc per year in EUR equiv
    // We sum col N filtered by SO and by year column having a value
    // Simpler: sum Budget col N (EUR equiv total) — but we need year breakdown
    // Year columns J/K/L in Budget, col A = SO, col I = currency
    // EUR equiv per year = if EUR then year_col else year_col/PGK_PER_EUR
    sh.getRange(row, 2).setFormula(`=IFERROR(SUMPRODUCT((Budget!A$2:A$2000="${code}")*(Budget!I$2:I$2000="EUR")*Budget!J$2:J$2000)+SUMPRODUCT((Budget!A$2:A$2000="${code}")*(Budget!I$2:I$2000="PGK")*Budget!J$2:J$2000)/${PGK_PER_EUR},0)`);
    sh.getRange(row, 3).setFormula(`=IFERROR(SUMPRODUCT((Budget!A$2:A$2000="${code}")*(Budget!I$2:I$2000="EUR")*Budget!K$2:K$2000)+SUMPRODUCT((Budget!A$2:A$2000="${code}")*(Budget!I$2:I$2000="PGK")*Budget!K$2:K$2000)/${PGK_PER_EUR},0)`);
    sh.getRange(row, 4).setFormula(`=IFERROR(SUMPRODUCT((Budget!A$2:A$2000="${code}")*(Budget!I$2:I$2000="EUR")*Budget!L$2:L$2000)+SUMPRODUCT((Budget!A$2:A$2000="${code}")*(Budget!I$2:I$2000="PGK")*Budget!L$2:L$2000)/${PGK_PER_EUR},0)`);
    sh.getRange(row, 5).setFormula(`=B${row}+C${row}+D${row}`);
    sh.getRange(row, 2, 1, 4).setNumberFormat('#,##0');
    row++;
  });
  // Grand total row
  sh.getRange(row, 1).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(row, 2).setFormula(`=IFERROR(SUMPRODUCT((Budget!I$2:I$2000="EUR")*Budget!J$2:J$2000)+SUMPRODUCT((Budget!I$2:I$2000="PGK")*Budget!J$2:J$2000)/${PGK_PER_EUR},0)`);
  sh.getRange(row, 3).setFormula(`=IFERROR(SUMPRODUCT((Budget!I$2:I$2000="EUR")*Budget!K$2:K$2000)+SUMPRODUCT((Budget!I$2:I$2000="PGK")*Budget!K$2:K$2000)/${PGK_PER_EUR},0)`);
  sh.getRange(row, 4).setFormula(`=IFERROR(SUMPRODUCT((Budget!I$2:I$2000="EUR")*Budget!L$2:L$2000)+SUMPRODUCT((Budget!I$2:I$2000="PGK")*Budget!L$2:L$2000)/${PGK_PER_EUR},0)`);
  sh.getRange(row, 5).setFormula(`=B${row}+C${row}+D${row}`);
  sh.getRange(row, 1, 1, 5).setBackground('#c9daf8').setFontWeight('bold');
  sh.getRange(row, 2, 1, 4).setNumberFormat('#,##0');
  row += 2;

  // 6b: By Cost Category × Year
  sectionStyle(sh.getRange(row, 1, 1, 5).merge());
  sh.getRange(row, 1).setValue('BUDGET BY COST CATEGORY (EUR equivalent)');
  row++;
  sh.getRange(row, 1, 1, 5).setValues([['Category','Yr 4 EUR','Yr 5 EUR','Yr 6 EUR','Total EUR']]);
  sh.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  ['Equipment','Capital','Consultant','Partner'].forEach(cat => {
    sh.getRange(row, 1).setValue(cat);
    sh.getRange(row, 2).setFormula(`=IFERROR(SUMPRODUCT((Budget!H$2:H$2000="${cat}")*(Budget!I$2:I$2000="EUR")*Budget!J$2:J$2000)+SUMPRODUCT((Budget!H$2:H$2000="${cat}")*(Budget!I$2:I$2000="PGK")*Budget!J$2:J$2000)/${PGK_PER_EUR},0)`);
    sh.getRange(row, 3).setFormula(`=IFERROR(SUMPRODUCT((Budget!H$2:H$2000="${cat}")*(Budget!I$2:I$2000="EUR")*Budget!K$2:K$2000)+SUMPRODUCT((Budget!H$2:H$2000="${cat}")*(Budget!I$2:I$2000="PGK")*Budget!K$2:K$2000)/${PGK_PER_EUR},0)`);
    sh.getRange(row, 4).setFormula(`=IFERROR(SUMPRODUCT((Budget!H$2:H$2000="${cat}")*(Budget!I$2:I$2000="EUR")*Budget!L$2:L$2000)+SUMPRODUCT((Budget!H$2:H$2000="${cat}")*(Budget!I$2:I$2000="PGK")*Budget!L$2:L$2000)/${PGK_PER_EUR},0)`);
    sh.getRange(row, 5).setFormula(`=B${row}+C${row}+D${row}`);
    sh.getRange(row, 2, 1, 4).setNumberFormat('#,##0');
    row++;
  });
  row++;

  // 6c: EUR vs PGK split (original currencies, not converted)
  sectionStyle(sh.getRange(row, 1, 1, 5).merge());
  sh.getRange(row, 1).setValue('BUDGET BY CURRENCY (original amounts — not converted)');
  row++;
  sh.getRange(row, 1, 1, 5).setValues([['Currency','Yr 4','Yr 5','Yr 6','Total']]);
  sh.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  ['EUR','PGK'].forEach(cur => {
    sh.getRange(row, 1).setValue(cur);
    sh.getRange(row, 2).setFormula(`=IFERROR(SUMIF(Budget!I$2:I$2000,"${cur}",Budget!J$2:J$2000),0)`);
    sh.getRange(row, 3).setFormula(`=IFERROR(SUMIF(Budget!I$2:I$2000,"${cur}",Budget!K$2:K$2000),0)`);
    sh.getRange(row, 4).setFormula(`=IFERROR(SUMIF(Budget!I$2:I$2000,"${cur}",Budget!L$2:L$2000),0)`);
    sh.getRange(row, 5).setFormula(`=B${row}+C${row}+D${row}`);
    sh.getRange(row, 2, 1, 4).setNumberFormat('#,##0.00');
    row++;
  });

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(2, 320);
  sh.setColumnWidth(3, 140);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 100);
  sh.setColumnWidth(6, 80);
  sh.setColumnWidth(7, 120);
}

// =============================================================================
// ON EDIT TRIGGER — approval workflow
// =============================================================================

function onEdit(e) {
  const sh = e.range.getSheet();

  // Budget sheet auto-fills
  if (sh.getName() === CONFIG.SHEET_BUDGET) {
    const col = e.range.getColumn();
    const row = e.range.getRow();
    if (row < 2) return;
    if (col === 3) autoFillBudgetFromActivity(sh, row, e.value); // Activity Code entered
    if (col === 5) autoFillBudgetFromTask(sh, row, e.value);     // Task Code entered
    return;
  }

  if (sh.getName() !== CONFIG.SHEET_PROGRESS) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row < 2) return;

  // Column 11 = Approval Status
  if (col === 11) {
    const newValue = e.value;
    const progressSh = sh;
    const rowData = progressSh.getRange(row, 1, 1, 14).getValues()[0];

    if (newValue === 'Approved' || newValue === 'Rejected') {
      // Stamp who approved and when
      const user = Session.getActiveUser().getEmail() || 'Approver';
      progressSh.getRange(row, 12).setValue(user);      // Approved By
      progressSh.getRange(row, 13).setValue(new Date()); // Approval Date

      // If approved, trigger dashboard refresh
      if (newValue === 'Approved') {
        refreshDashboard();
        // Notify the submitter
        notifySubmitter(rowData, 'Approved');
      } else {
        notifySubmitter(rowData, 'Rejected');
      }
    }

    // Auto-clear approval stamp if reset to Pending
    if (newValue === 'Pending') {
      progressSh.getRange(row, 12).clearContent();
      progressSh.getRange(row, 13).clearContent();
    }
  }

  // Column 2 = Task Code entered in Progress Log — auto-fill related columns
  if (col === 2 && sh.getName() === CONFIG.SHEET_PROGRESS) {
    autoFillProgressRow(sh, row, e.value);
  }
}

function autoFillProgressRow(sh, row, taskCode) {
  if (!taskCode) return;
  const tasksSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_TASKS);
  const tasksData = tasksSh.getDataRange().getValues();
  for (let i = 1; i < tasksData.length; i++) {
    if (tasksData[i][0] === taskCode) {
      sh.getRange(row, 3).setValue(tasksData[i][1]); // Activity Code
      sh.getRange(row, 4).setValue(tasksData[i][4]); // Activity Name
      break;
    }
  }
}

// =============================================================================
// REFRESH DASHBOARD — recalculates % complete on Activities sheet
// =============================================================================

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const actSh  = ss.getSheetByName(CONFIG.SHEET_ACTIVITIES);
  const taskSh = ss.getSheetByName(CONFIG.SHEET_TASKS);
  const progSh = ss.getSheetByName(CONFIG.SHEET_PROGRESS);

  const actData  = actSh.getRange(2, 1, actSh.getLastRow() - 1, 11).getValues();
  const taskData = taskSh.getLastRow() > 1 ? taskSh.getRange(2, 1, taskSh.getLastRow() - 1, 14).getValues() : [];
  const progData = progSh.getLastRow() > 1 ? progSh.getRange(2, 1, progSh.getLastRow() - 1, 14).getValues() : [];

  // Only approved entries
  const approved = progData.filter(r => r[10] === 'Approved');

  actData.forEach((act, i) => {
    const actCode = act[2];

    // Count tasks for this activity
    const actTasks = taskData.filter(t => t[1] === actCode);
    actSh.getRange(i + 2, 9).setValue(actTasks.length);

    if (actTasks.length === 0) {
      actSh.getRange(i + 2, 10).setValue('');
      actSh.getRange(i + 2, 11).setValue('No Tasks Defined');
      return;
    }

    // For each task, find the latest approved % complete
    let totalWeight = 0;
    let weightedComplete = 0;

    actTasks.forEach(task => {
      const taskCode = task[0];
      const weight = Number(task[10]) || (100 / actTasks.length); // default equal weight
      const taskApproved = approved.filter(p => p[1] === taskCode);
      // Take highest approved % for this task (most recent meaningful entry)
      let pct = 0;
      if (taskApproved.length > 0) {
        pct = Math.max(...taskApproved.map(p => Number(p[6]) || 0));
      }
      totalWeight += weight;
      weightedComplete += (pct / 100) * weight;
    });

    const pctComplete = totalWeight > 0 ? weightedComplete / totalWeight : 0;
    actSh.getRange(i + 2, 10).setValue(pctComplete).setNumberFormat('0%');

    // Status
    let status = 'Not Started';
    if (pctComplete > 0 && pctComplete < 1) status = 'In Progress';
    if (pctComplete >= 1) status = 'Complete';
    actSh.getRange(i + 2, 11).setValue(status);
  });

  // Refresh timestamp on dashboard
  const dashSh = ss.getSheetByName(CONFIG.SHEET_DASHBOARD);
  if (dashSh) dashSh.getRange('B1').setValue(new Date());

  SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard refreshed with approved data only.', 'MCA Tracker', 3);
}

// =============================================================================
// DIALOGS FOR DATA ENTRY
// =============================================================================

function showAddTaskDialog() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(ADD_TASK_HTML)
    .setWidth(520)
    .setHeight(600)
    .setTitle('Add Tasks to Activity');
  ui.showModalDialog(html, 'Add Tasks to Activity');
}

function showProgressDialog() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(LOG_PROGRESS_HTML)
    .setWidth(520)
    .setHeight(580)
    .setTitle('Log Progress on Task');
  ui.showModalDialog(html, 'Log Progress on Task');
}

// Called from HTML dialog — saves a new task
function saveTask(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSh = ss.getSheetByName(CONFIG.SHEET_TASKS);
  const actSh  = ss.getSheetByName(CONFIG.SHEET_ACTIVITIES);

  // Find activity info
  const actData = actSh.getDataRange().getValues();
  let actRow = null;
  for (let i = 1; i < actData.length; i++) {
    if (actData[i][2] === data.activityCode) { actRow = actData[i]; break; }
  }
  if (!actRow) throw new Error('Activity code not found: ' + data.activityCode);

  // Generate task code
  const existing = taskSh.getLastRow() > 1 ? taskSh.getRange(2, 1, taskSh.getLastRow() - 1, 2).getValues() : [];
  const actTasks = existing.filter(r => r[1] === data.activityCode);
  const taskNum  = actTasks.length + 1;
  const taskCode = data.activityCode + '.T' + taskNum;

  const row = [
    taskCode,
    data.activityCode,
    actRow[0], // Strategic Output
    actRow[1], // Sub-Output
    actRow[3], // Activity Name
    data.taskName,
    data.taskDescription,
    data.assignedTo,
    data.plannedStart,
    data.plannedEnd,
    data.weight || '',
    Session.getActiveUser().getEmail(),
    new Date(),
    'Not Started',
    data.dependsOn || ''
  ];

  taskSh.appendRow(row);
  return taskCode;
}

// Called from HTML dialog — saves a progress entry
function saveProgress(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progSh = ss.getSheetByName(CONFIG.SHEET_PROGRESS);
  const taskSh = ss.getSheetByName(CONFIG.SHEET_TASKS);

  // Find task info
  const taskData = taskSh.getLastRow() > 1 ? taskSh.getRange(2, 1, taskSh.getLastRow() - 1, 14).getValues() : [];
  let taskRow = null;
  for (let t of taskData) { if (t[0] === data.taskCode) { taskRow = t; break; } }
  if (!taskRow) throw new Error('Task code not found: ' + data.taskCode);

  // Auto-increment entry ID
  const lastRow = progSh.getLastRow();
  const entryId = lastRow < 2 ? 'E001' : 'E' + String(lastRow).padStart(3, '0');

  const row = [
    entryId,
    data.taskCode,
    taskRow[1], // Activity Code
    taskRow[4], // Activity Name
    data.submittedBy || Session.getActiveUser().getEmail(),
    new Date(),
    Number(data.pctComplete),
    data.statusUpdate,
    data.evidenceNotes,
    data.supportingLink || '',
    'Pending', // Always starts as Pending
    '', // Approved By
    '', // Approval Date
    ''  // Approver Notes
  ];

  progSh.appendRow(row);

  // Email the approver
  notifyApprover(row, entryId);
  return entryId;
}

// Get task list for a given activity code (called from dialog)
function getTasksForActivity(activityCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSh = ss.getSheetByName(CONFIG.SHEET_TASKS);
  if (taskSh.getLastRow() < 2) return [];
  const data = taskSh.getRange(2, 1, taskSh.getLastRow() - 1, 14).getValues();
  return data.filter(r => r[1] === activityCode && r[0]).map(r => ({ code: r[0], name: r[5] }));
}

// Get all activity codes from the Activities sheet (called from dialog)
// Reads live so new rows added directly to the sheet are picked up immediately
function getActivityCodes() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ACTIVITIES);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const data = sh.getRange(2, 1, lastRow - 1, 8).getValues();
  return data
    .filter(r => r[2]) // skip blank rows
    .map(r => ({ code: r[2], name: r[3], partner: r[6] }));
}

// =============================================================================
// BUDGET SHEET
// =============================================================================

function setupBudgetSheet(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_BUDGET);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_BUDGET);
  sh.clearContents();
  sh.clearFormats();

  const headers = [
    'Strategic Output', 'Sub-Output', 'Activity Code', 'Activity Name',
    'Task Code', 'Task Name', 'Responsible Partner',
    'Cost Category', 'Currency',
    'Yr 4 (Jun 26–May 27)', 'Yr 5 (Jun 27–May 28)', 'Yr 6 (Jun 28–May 29)',
    'Total', 'EUR Equivalent', 'Notes'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sh, headers.length);

  // Exchange rate note in row 1 col P
  sh.getRange(1, 16).setValue('Rate: PGK ' + PGK_PER_EUR + ' = EUR 1');
  sh.getRange(1, 16).setFontStyle('italic').setFontColor('#666666').setBackground('#f3f3f3');

  // Data validation — Activity Code (col 3)
  const actSh = ss.getSheetByName(CONFIG.SHEET_ACTIVITIES);
  const actRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(actSh.getRange('C2:C500'), true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, 3, 1000, 1).setDataValidation(actRule);

  // Cost Category (col 8)
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Equipment', 'Capital', 'Consultant', 'Partner'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, 8, 1000, 1).setDataValidation(catRule);

  // Currency (col 9)
  const currRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['EUR', 'PGK'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, 9, 1000, 1).setDataValidation(currRule);
  sh.getRange(2, 9, 1000, 1).setValue('EUR');

  // Number format for amount columns (10-14)
  sh.getRange(2, 10, 1000, 5).setNumberFormat('#,##0.00');

  // Total formula (col 13) = sum of Yr4+Yr5+Yr6
  // EUR Equivalent (col 14) = if EUR then Total else Total/PGK_PER_EUR
  // These are set as formulas row by row via a helper, or we set them as array-style below
  // We'll set them per-row in onEdit, but also set a template note
  sh.getRange(2, 13).setFormula('=IF(J2+K2+L2=0,"",J2+K2+L2)');
  sh.getRange(2, 14).setFormula('=IF(M2="","",IF(I2="EUR",M2,M2/' + PGK_PER_EUR + '))');

  // Freeze header, set column widths
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 260);
  sh.setColumnWidth(5, 100);
  sh.setColumnWidth(6, 220);
  sh.setColumnWidth(7, 130);
  sh.setColumnWidth(8, 110);
  sh.setColumnWidth(9, 70);
  sh.setColumnWidth(10, 130);
  sh.setColumnWidth(11, 130);
  sh.setColumnWidth(12, 130);
  sh.setColumnWidth(13, 110);
  sh.setColumnWidth(14, 110);
  sh.setColumnWidth(15, 220);

  // Conditional formatting by SO
  const so1Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2="SO1"').setBackground('#d9ead3')
    .setRanges([sh.getRange('A2:O1000')]).build();
  const so2Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2="SO2"').setBackground('#dae8fc')
    .setRanges([sh.getRange('A2:O1000')]).build();
  const so3Rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2="SO3"').setBackground('#fff2cc')
    .setRanges([sh.getRange('A2:O1000')]).build();
  sh.setConditionalFormatRules([so1Rule, so2Rule, so3Rule]);

  // Helper note for users
  sh.getRange(2, 1).setNote(
    'How to use:\n' +
    '1. Type or select an Activity Code in column C — columns A, B, D, G fill automatically\n' +
    '2. Optionally enter a Task Code in column E — column F fills automatically\n' +
    '3. Select Cost Category and Currency\n' +
    '4. Enter amounts for each project year\n' +
    '5. Total and EUR Equivalent calculate automatically\n\n' +
    'Copy row 2 formulas (cols M & N) down as you add new rows.'
  );
}

function autoFillBudgetFromActivity(sh, row, actCode) {
  if (!actCode) return;
  const actSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ACTIVITIES);
  const data = actSh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === actCode) {
      sh.getRange(row, 1).setValue(data[i][0]); // Strategic Output
      sh.getRange(row, 2).setValue(data[i][1]); // Sub-Output
      sh.getRange(row, 4).setValue(data[i][3]); // Activity Name
      sh.getRange(row, 7).setValue(data[i][6]); // Partner
      // Set Total and EUR Equivalent formulas for this row
      sh.getRange(row, 13).setFormula('=IF(J' + row + '+K' + row + '+L' + row + '=0,"",J' + row + '+K' + row + '+L' + row + ')');
      sh.getRange(row, 14).setFormula('=IF(M' + row + '="","",IF(I' + row + '="EUR",M' + row + ',M' + row + '/' + PGK_PER_EUR + '))');
      sh.getRange(row, 13).setNumberFormat('#,##0.00');
      sh.getRange(row, 14).setNumberFormat('#,##0.00');
      break;
    }
  }
}

function autoFillBudgetFromTask(sh, row, taskCode) {
  if (!taskCode) return;
  const taskSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_TASKS);
  if (taskSh.getLastRow() < 2) return;
  const data = taskSh.getRange(2, 1, taskSh.getLastRow() - 1, 15).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === taskCode) {
      sh.getRange(row, 6).setValue(data[i][5]); // Task Name
      // Also fill activity if not already filled
      const existingAct = sh.getRange(row, 3).getValue();
      if (!existingAct) {
        sh.getRange(row, 3).setValue(data[i][1]);
        autoFillBudgetFromActivity(sh, row, data[i][1]);
      }
      break;
    }
  }
}

// =============================================================================
// GANTT CHART
// =============================================================================

function showGanttChart() {
  const html = HtmlService.createHtmlOutput(GANTT_HTML)
    .setWidth(1100)
    .setHeight(680)
    .setTitle('MCA Gantt Chart');
  SpreadsheetApp.getUi().showModalDialog(html, 'MCA Project Gantt Chart');
}

function getGanttData() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const actSh   = ss.getSheetByName(CONFIG.SHEET_ACTIVITIES);
  const taskSh  = ss.getSheetByName(CONFIG.SHEET_TASKS);
  const progSh  = ss.getSheetByName(CONFIG.SHEET_PROGRESS);
  const tz      = Session.getScriptTimeZone();

  const actData  = actSh.getLastRow()  > 1 ? actSh.getRange(2,  1, actSh.getLastRow()  - 1, 11).getValues() : [];
  const taskData = taskSh.getLastRow() > 1 ? taskSh.getRange(2, 1, taskSh.getLastRow() - 1, 15).getValues() : [];
  const progData = progSh.getLastRow() > 1 ? progSh.getRange(2, 1, progSh.getLastRow() - 1, 14).getValues() : [];

  // Latest approved % per task
  const taskPct = {};
  progData.filter(r => r[10] === 'Approved').forEach(p => {
    const code = p[1], pct = Number(p[6]) || 0;
    if (!taskPct[code] || pct > taskPct[code]) taskPct[code] = pct;
  });

  function fmtDate(v) {
    if (!v) return '';
    try {
      const d = new Date(v);
      return isNaN(d) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    } catch(e) { return ''; }
  }

  const activities = actData
    .filter(r => r[2])
    .map(r => {
      const actCode = r[2];
      const tasks = taskData
        .filter(t => t[1] === actCode && t[0])
        .map(t => ({
          code:       t[0],
          name:       t[5],
          assignedTo: t[7],
          start:      fmtDate(t[8]),
          end:        fmtDate(t[9]),
          dependsOn:  (t[14] || '').toString().trim(),
          status:     t[13] || 'Not Started',
          pctComplete: taskPct[t[0]] || 0,
        }));
      return {
        code:      actCode,
        name:      r[3],
        so:        r[0],
        subOutput: r[1],
        partner:   r[6],
        ciStaff:   r[7],
        scheduled: r[4],
        tasks,
      };
    });

  return { activities };
}

const GANTT_HTML = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;font-size:12px;background:#fff;overflow:hidden}
#toolbar{display:flex;align-items:center;gap:10px;padding:6px 12px;background:#1c4587;color:#fff;flex-wrap:wrap}
#toolbar strong{font-size:13px;white-space:nowrap}
#toolbar label{font-size:11px;white-space:nowrap}
#toolbar select,#toolbar input{padding:3px 6px;font-size:11px;border-radius:3px;border:none}
.legend{display:flex;gap:10px;margin-left:auto;align-items:center;font-size:11px}
.ld{width:12px;height:12px;border-radius:2px;display:inline-block;margin-right:3px;vertical-align:middle}
#wrap{display:flex;height:calc(100vh - 38px)}
#labels{width:280px;flex-shrink:0;border-right:2px solid #1c4587;overflow:hidden;display:flex;flex-direction:column}
#lhdr{height:52px;background:#e8f0fe;border-bottom:1px solid #ccc;display:flex;align-items:center;padding:0 8px;font-weight:bold;font-size:11px;color:#1c4587;flex-shrink:0}
#lbody{overflow-y:auto;flex:1}
#chart{flex:1;overflow:auto;position:relative}
#loading{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);font-size:14px;color:#666}
.grp{background:#1c4587;color:#fff;font-weight:bold;font-size:11px;padding:3px 8px;height:22px;display:flex;align-items:center}
.sub{background:#dce8fb;color:#1c4587;font-weight:bold;font-size:10px;padding:2px 14px;height:18px;display:flex;align-items:center}
.arow{height:36px;border-bottom:1px solid #eee;display:flex;align-items:center;padding:0 8px;cursor:default;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.arow:hover{background:#f0f4ff}
.acode{color:#666;font-weight:bold;margin-right:4px;font-size:10px}
.aname{font-size:10px;overflow:hidden;text-overflow:ellipsis}
#tooltip{position:fixed;background:#333;color:#fff;padding:6px 10px;border-radius:4px;font-size:11px;pointer-events:none;display:none;z-index:999;max-width:280px;line-height:1.5}
</style>
</head>
<body>
<div id="toolbar">
  <strong>MCA Gantt Chart</strong>
  <label>Output: <select id="fSO" onchange="applyFilter()">
    <option value="">All Outputs</option>
    <option value="SO1">SO1 – Policy &amp; Governance</option>
    <option value="SO2">SO2 – Awareness &amp; Research</option>
    <option value="SO3">SO3 – Livelihoods &amp; Economy</option>
  </select></label>
  <label>Sub-Output: <select id="fSub" onchange="applyFilter()"><option value="">All</option></select></label>
  <label>Show: <select id="fShow" onchange="applyFilter()">
    <option value="all">All activities</option>
    <option value="tasks">Activities with tasks only</option>
  </select></label>
  <div class="legend">
    <span><span class="ld" style="background:#4a7c59"></span>SO1</span>
    <span><span class="ld" style="background:#3d6b9e"></span>SO2</span>
    <span><span class="ld" style="background:#b8860b"></span>SO3</span>
    <span><span class="ld" style="background:#ccc"></span>No tasks yet</span>
    <span style="color:#e53935;font-size:11px">— Today</span>
    <span style="color:#e53935;font-size:11px">→ Dependency</span>
  </div>
</div>
<div id="wrap">
  <div id="labels"><div id="lhdr">Activity</div><div id="lbody"></div></div>
  <div id="chart"><div id="loading">Loading chart data…</div><canvas id="cv"></canvas></div>
</div>
<div id="tooltip"></div>

<script>
var RAW=null, FILTERED=null;
var ROW_H=36, GRP_H=22, SUB_H=18, HDR_H=52;
var SO_COLOR={SO1:'#4a7c59',SO2:'#3d6b9e',SO3:'#b8860b'};
var SO_LIGHT={SO1:'#d9ead3',SO2:'#dae8fc',SO3:'#fff2cc'};
var taskPos={};

function qStart(q){var m=q&&q.match(/Q([1-4]) (\\d{4})/);if(!m)return null;return new Date(+m[2],(+m[1]-1)*3,1);}
function qEnd(q){var m=q&&q.match(/Q([1-4]) (\\d{4})/);if(!m)return null;return new Date(+m[2],+m[1]*3,0);}
function pd(v){if(!v)return null;var d=new Date(v);return isNaN(d)?null:d;}

function applyFilter(){
  if(!RAW)return;
  var fSO=document.getElementById('fSO').value;
  var fSub=document.getElementById('fSub').value;
  var fShow=document.getElementById('fShow').value;
  // refresh sub-output options
  var subSel=document.getElementById('fSub');
  var curSub=subSel.value;
  var subs=[...new Set(RAW.activities.filter(function(a){return !fSO||a.so===fSO;}).map(function(a){return a.subOutput;}))].sort();
  subSel.innerHTML='<option value="">All</option>';
  subs.forEach(function(s){var o=document.createElement('option');o.value=s;o.textContent=s;if(s===curSub)o.selected=true;subSel.appendChild(o);});

  FILTERED=RAW.activities.filter(function(a){
    if(fSO&&a.so!==fSO)return false;
    if(fSub&&a.subOutput!==fSub)return false;
    if(fShow==='tasks'&&a.tasks.length===0)return false;
    return true;
  });
  drawLabels();
  drawChart();
}

function drawLabels(){
  var body=document.getElementById('lbody');
  body.innerHTML='';
  var curSO=null,curSub=null;
  FILTERED.forEach(function(act){
    if(act.so!==curSO){curSO=act.so;curSub=null;
      var h=document.createElement('div');h.className='grp';
      h.textContent=act.so+(act.so==='SO1'?' – Policy & Governance':act.so==='SO2'?' – Awareness & Research':' – Livelihoods & Economy');
      body.appendChild(h);}
    if(act.subOutput!==curSub){curSub=act.subOutput;
      var h=document.createElement('div');h.className='sub';h.textContent='Sub-Output '+act.subOutput;body.appendChild(h);}
    var r=document.createElement('div');r.className='arow';
    r.title=act.code+': '+act.name+' | Partner: '+act.partner;
    r.innerHTML='<span class="acode">'+act.code+'</span><span class="aname">'+act.name+'</span>';
    body.appendChild(r);
  });
}

function buildRows(){
  var rows=[],curSO=null,curSub=null;
  FILTERED.forEach(function(act){
    if(act.so!==curSO){curSO=act.so;curSub=null;rows.push({type:'so',so:act.so,h:GRP_H});}
    if(act.subOutput!==curSub){curSub=act.subOutput;rows.push({type:'sub',sub:act.subOutput,h:SUB_H});}
    rows.push({type:'act',act:act,h:ROW_H});
  });
  return rows;
}

function drawChart(){
  var cv=document.getElementById('cv');
  var panel=document.getElementById('chart');
  var rows=buildRows();

  // Date range
  var minD=new Date('2024-06-01'),maxD=new Date('2028-12-31');
  FILTERED.forEach(function(act){
    act.tasks.forEach(function(t){
      var s=pd(t.start),e=pd(t.end);
      if(s&&s<minD)minD=s; if(e&&e>maxD)maxD=e;
    });
    var qs=qStart(act.scheduled),qe=qEnd(act.scheduled);
    if(qs&&qs<minD)minD=qs; if(qe&&qe>maxD)maxD=qe;
  });

  var totalH=rows.reduce(function(s,r){return s+r.h;},0)+HDR_H+4;
  var totalW=Math.max(panel.clientWidth-4, 900);
  cv.width=totalW; cv.height=totalH;
  var ctx=cv.getContext('2d');
  ctx.clearRect(0,0,totalW,totalH);

  var DAY=86400000;
  var span=(maxD-minD)/DAY;
  var dayPx=totalW/span;
  function xOf(d){return((d-minD)/DAY)*dayPx;}

  // Header background
  ctx.fillStyle='#e8f0fe'; ctx.fillRect(0,0,totalW,HDR_H);

  // Quarter/year grid + labels
  var y0=minD.getFullYear(), y1=maxD.getFullYear()+1;
  for(var y=y0;y<=y1;y++){
    for(var q=0;q<4;q++){
      var qd=new Date(y,q*3,1);
      if(qd<minD||qd>maxD)continue;
      var x=xOf(qd);
      ctx.strokeStyle=q===0?'#aaa':'#e0e0e0'; ctx.lineWidth=q===0?1.5:1;
      ctx.beginPath();ctx.moveTo(x,0);ctx.lineTo(x,totalH);ctx.stroke();
      ctx.fillStyle='#1c4587';
      ctx.font=q===0?'bold 11px Arial':'10px Arial';
      ctx.fillText(q===0?String(y):'Q'+(q+1),x+3,q===0?14:28);
      ctx.fillStyle='#555'; ctx.font='10px Arial';
      if(q>0)ctx.fillText('Q'+(q+1)+' '+y,x+3,42);
    }
  }

  // Today line
  var today=new Date();
  if(today>=minD&&today<=maxD){
    var tx=xOf(today);
    ctx.save();ctx.strokeStyle='#e53935';ctx.lineWidth=2;ctx.setLineDash([4,3]);
    ctx.beginPath();ctx.moveTo(tx,0);ctx.lineTo(tx,totalH);ctx.stroke();
    ctx.setLineDash([]);ctx.fillStyle='#e53935';ctx.font='bold 10px Arial';
    ctx.fillText('Today',tx+3,HDR_H-4);ctx.restore();
  }

  // Draw rows
  taskPos={};
  var cy=HDR_H;
  rows.forEach(function(row){
    if(row.type==='so'){
      ctx.fillStyle='#1c4587';ctx.fillRect(0,cy,totalW,GRP_H);cy+=GRP_H;
    } else if(row.type==='sub'){
      ctx.fillStyle='#dce8fb';ctx.fillRect(0,cy,totalW,SUB_H);cy+=SUB_H;
    } else {
      var act=row.act;
      var color=SO_COLOR[act.so]||'#666';
      var light=SO_LIGHT[act.so]||'#eee';

      // Row background
      ctx.fillStyle=light+'44'; ctx.fillRect(0,cy,totalW,ROW_H);

      if(act.tasks.length===0){
        // Placeholder bar from scheduled quarter
        var qs=qStart(act.scheduled),qe=qEnd(act.scheduled);
        if(qs&&qe){
          var bx=xOf(qs),bw=Math.max(xOf(qe)-bx,4);
          var by=cy+10,bh=16;
          ctx.fillStyle='#ddd';ctx.fillRect(bx,by,bw,bh);
          ctx.strokeStyle='#bbb';ctx.lineWidth=1;ctx.strokeRect(bx,by,bw,bh);
          ctx.fillStyle='#999';ctx.font='9px Arial';ctx.fillText('No tasks – '+act.scheduled,bx+3,by+bh-4);
        }
      } else {
        // Assign tasks to lanes (detect parallelism)
        var sorted=act.tasks.slice().sort(function(a,b){
          var as=pd(a.start),bs=pd(b.start);
          if(!as&&!bs)return 0; if(!as)return 1; if(!bs)return -1; return as-bs;
        });
        var laneEnd=[];
        sorted.forEach(function(t){
          var ts=pd(t.start),te=pd(t.end);
          if(!ts||!te){t._lane=0;return;}
          var lane=0;
          while(laneEnd[lane]&&laneEnd[lane]>ts)lane++;
          laneEnd[lane]=te; t._lane=lane;
        });
        var nLanes=Math.max(1,laneEnd.length);
        var lh=Math.floor((ROW_H-6)/nLanes);
        var minBarH=Math.max(6,lh-2);

        sorted.forEach(function(t){
          var ts=pd(t.start),te=pd(t.end);
          if(!ts||!te)return;
          var bx=xOf(ts),bw=Math.max(xOf(te)-bx,4);
          var by=cy+3+(t._lane||0)*lh;
          var bh=minBarH;
          var pct=(t.pctComplete||0)/100;

          // Background track
          ctx.fillStyle='#e0e0e0'; ctx.fillRect(bx,by,bw,bh);
          // Progress fill
          if(pct>0){ctx.fillStyle=color;ctx.fillRect(bx,by,bw*pct,bh);}
          // Border
          ctx.strokeStyle=color;ctx.lineWidth=1;ctx.strokeRect(bx,by,bw,bh);

          // Store for dependency arrows and tooltips
          taskPos[t.code]={x1:bx,x2:bx+bw,y1:by,y2:by+bh,cy:by+bh/2,task:t,act:act};

          // Label if bar wide enough
          if(bw>35){
            ctx.save();ctx.beginPath();ctx.rect(bx+2,by,bw-4,bh);ctx.clip();
            ctx.fillStyle=pct>0.55?'#fff':'#333';ctx.font='8px Arial';
            ctx.fillText(t.name,bx+3,by+bh-3);ctx.restore();
          }
        });
      }
      cy+=ROW_H;
    }
  });

  // Dependency arrows (drawn after all bars so arrows appear on top)
  rows.filter(function(r){return r.type==='act';}).forEach(function(row){
    row.act.tasks.forEach(function(t){
      if(!t.dependsOn)return;
      var from=taskPos[t.dependsOn],to=taskPos[t.code];
      if(!from||!to)return;
      ctx.save();ctx.strokeStyle='#e53935';ctx.lineWidth=1.5;ctx.setLineDash([3,2]);
      ctx.beginPath();
      ctx.moveTo(from.x2,from.cy);
      var mx=(from.x2+to.x1)/2;
      ctx.bezierCurveTo(mx,from.cy,mx,to.cy,to.x1,to.cy);
      ctx.stroke();ctx.setLineDash([]);
      // Arrowhead
      ctx.fillStyle='#e53935';ctx.beginPath();
      ctx.moveTo(to.x1,to.cy);ctx.lineTo(to.x1-6,to.cy-4);ctx.lineTo(to.x1-6,to.cy+4);
      ctx.closePath();ctx.fill();ctx.restore();
    });
  });

  document.getElementById('loading').style.display='none';

  // Sync scroll
  var lb=document.getElementById('lbody');
  panel.onscroll=function(){lb.scrollTop=Math.max(0,panel.scrollTop-HDR_H);};

  // Tooltip on hover
  cv.onmousemove=function(e){
    var rect=cv.getBoundingClientRect();
    var mx=e.clientX-rect.left, my=e.clientY-rect.top;
    var tip=document.getElementById('tooltip');
    var hit=null;
    Object.keys(taskPos).forEach(function(k){
      var p=taskPos[k];
      if(mx>=p.x1&&mx<=p.x2&&my>=p.y1&&my<=p.y2)hit=p;
    });
    if(hit){
      var t=hit.task,a=hit.act;
      tip.innerHTML='<b>'+t.code+'</b><br>'+t.name+'<br>'+'Activity: '+a.code+'<br>'+'Assigned: '+t.assignedTo+'<br>'+'Dates: '+(t.start||'?')+' → '+(t.end||'?')+'<br>'+'Complete: '+(t.pctComplete||0)+'%'+(t.dependsOn?'<br>Depends on: '+t.dependsOn:'');
      tip.style.display='block';tip.style.left=(e.clientX+14)+'px';tip.style.top=(e.clientY-10)+'px';
    } else {tip.style.display='none';}
  };
  cv.onmouseleave=function(){document.getElementById('tooltip').style.display='none';};
}

google.script.run
  .withSuccessHandler(function(data){
    RAW=data;
    // populate sub-output dropdown
    var subs=[...new Set(data.activities.map(function(a){return a.subOutput;}))].sort();
    var subSel=document.getElementById('fSub');
    subs.forEach(function(s){var o=document.createElement('option');o.value=s;o.textContent=s;subSel.appendChild(o);});
    FILTERED=data.activities.slice();
    drawLabels();
    drawChart();
  })
  .withFailureHandler(function(e){
    document.getElementById('loading').textContent='Error: '+e.message;
  })
  .getGanttData();

window.onresize=function(){if(FILTERED)drawChart();};
</script>
</body>
</html>
`;

// =============================================================================
// EMAIL NOTIFICATIONS
// =============================================================================

function notifyApprover(row, entryId) {
  try {
    const subject = `[MCA Tracker] New progress entry awaiting approval — ${row[2]}`;
    const body = `A new progress entry has been submitted and requires your approval.\n\n` +
      `Entry ID:      ${entryId}\n` +
      `Task Code:     ${row[1]}\n` +
      `Activity:      ${row[3]}\n` +
      `Submitted by:  ${row[4]}\n` +
      `% Complete:    ${row[6]}%\n` +
      `Status Update: ${row[7]}\n\n` +
      `To approve or reject, open the MCA Tracker spreadsheet,\n` +
      `go to the "Progress Log" sheet and change the "Approval Status"\n` +
      `column for entry ${entryId} to Approved or Rejected.\n\n` +
      `You can add notes in the "Approver Notes" column.\n`;
    GmailApp.sendEmail(CONFIG.APPROVER_EMAIL, subject, body);
  } catch (e) {
    Logger.log('Email error: ' + e.message);
  }
}

function notifySubmitter(rowData, decision) {
  try {
    const email = rowData[4]; // Submitted By
    if (!email || !email.includes('@')) return;
    const subject = `[MCA Tracker] Your progress entry has been ${decision} — ${rowData[2]}`;
    const body = `Your progress entry for the following task has been ${decision}:\n\n` +
      `Entry ID:      ${rowData[0]}\n` +
      `Task Code:     ${rowData[1]}\n` +
      `Activity:      ${rowData[3]}\n` +
      `% Complete:    ${rowData[6]}%\n` +
      `Decision:      ${decision}\n` +
      `Approver:      ${rowData[11]}\n` +
      `Approver Notes: ${rowData[13] || 'None'}\n`;
    GmailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log('Submitter email error: ' + e.message);
  }
}

// =============================================================================
// UTILITY
// =============================================================================

function styleHeaderRow(sh, numCols) {
  const hdr = sh.getRange(1, 1, 1, numCols);
  hdr.setFontWeight('bold')
     .setFontColor('#ffffff')
     .setBackground('#1c4587')
     .setWrap(true);
  sh.setRowHeight(1, 40);
}

// =============================================================================
// HTML DIALOGS (embedded)
// =============================================================================

const ADD_TASK_HTML = `
<!DOCTYPE html>
<html>
<head>
<style>
  body { font-family: Arial, sans-serif; font-size: 13px; padding: 10px; }
  label { display: block; margin-top: 10px; font-weight: bold; }
  input, select, textarea { width: 100%; padding: 5px; margin-top: 3px; box-sizing: border-box; }
  button { margin-top: 15px; padding: 8px 18px; background: #1c4587; color: white; border: none; border-radius: 3px; cursor: pointer; }
  #msg { margin-top: 10px; padding: 8px; border-radius: 3px; display: none; }
  .success { background: #d9ead3; color: #1c4c1c; }
  .error   { background: #f4cccc; color: #7c0000; }
</style>
</head>
<body>
<h3 style="color:#1c4587">Add Tasks to Activity</h3>
<label>Activity Code *
  <select id="actCode" required>
    <option value="">— select —</option>
  </select>
</label>
<label>Task Name *<input id="taskName" type="text" required></label>
<label>Task Description<textarea id="taskDesc" rows="3"></textarea></label>
<label>Assigned To *<input id="assignedTo" type="text" required></label>
<label>Planned Start Date<input id="startDate" type="date"></label>
<label>Planned End Date<input id="endDate" type="date"></label>
<label>Weight (% of activity this task represents; all tasks should sum to 100)
  <input id="weight" type="number" min="0" max="100" placeholder="e.g. 25">
</label>
<label>Depends On (Task Code)
  <input id="dependsOn" type="text" placeholder="e.g. 1.1-1.T1 — leave blank if no dependency">
  <small style="color:#666;font-size:11px">This task cannot start until the task entered here is complete</small>
</label>
<button onclick="submit()">Add Task</button>
<button onclick="google.script.host.close()" style="background:#888">Cancel</button>
<div id="msg"></div>
<script>
  google.script.run.withSuccessHandler(function(acts) {
    var sel = document.getElementById('actCode');
    acts.forEach(function(a) {
      var opt = document.createElement('option');
      opt.value = a.code;
      opt.textContent = a.code + ' — ' + a.name.substring(0,55);
      sel.appendChild(opt);
    });
  }).getActivityCodes();

  function submit() {
    var data = {
      activityCode:    document.getElementById('actCode').value,
      taskName:        document.getElementById('taskName').value.trim(),
      taskDescription: document.getElementById('taskDesc').value.trim(),
      assignedTo:      document.getElementById('assignedTo').value.trim(),
      plannedStart:    document.getElementById('startDate').value,
      plannedEnd:      document.getElementById('endDate').value,
      weight:          document.getElementById('weight').value,
      dependsOn:       document.getElementById('dependsOn').value.trim()
    };
    if (!data.activityCode || !data.taskName || !data.assignedTo) {
      show('Please fill in all required fields (*).', 'error'); return;
    }
    google.script.run
      .withSuccessHandler(function(code) { show('Task saved! Code: ' + code + '. You can add another task or close.', 'success'); })
      .withFailureHandler(function(e)    { show('Error: ' + e.message, 'error'); })
      .saveTask(data);
  }
  function show(msg, type) {
    var el = document.getElementById('msg');
    el.textContent = msg; el.className = type; el.style.display = 'block';
  }
</script>
</body>
</html>
`;

const LOG_PROGRESS_HTML = `
<!DOCTYPE html>
<html>
<head>
<style>
  body { font-family: Arial, sans-serif; font-size: 13px; padding: 10px; }
  label { display: block; margin-top: 10px; font-weight: bold; }
  input, select, textarea { width: 100%; padding: 5px; margin-top: 3px; box-sizing: border-box; }
  button { margin-top: 15px; padding: 8px 18px; background: #1c4587; color: white; border: none; border-radius: 3px; cursor: pointer; }
  #msg { margin-top: 10px; padding: 8px; border-radius: 3px; display: none; }
  .success { background: #d9ead3; color: #1c4c1c; }
  .error   { background: #f4cccc; color: #7c0000; }
  #taskSel { display: none; }
</style>
</head>
<body>
<h3 style="color:#1c4587">Log Progress on Task</h3>
<label>Your Name / Email *<input id="submittedBy" type="text" required placeholder="name or email"></label>
<label>Activity Code *
  <select id="actCode" onchange="loadTasks()" required>
    <option value="">— select activity —</option>
  </select>
</label>
<div id="taskSel">
  <label>Task *
    <select id="taskCode" required>
      <option value="">— select task —</option>
    </select>
  </label>
</div>
<label>% Complete (0–100) *<input id="pct" type="number" min="0" max="100" required placeholder="e.g. 50"></label>
<label>Status Update *<textarea id="statusUpdate" rows="3" required placeholder="What was done? What is the current status?"></textarea></label>
<label>Evidence / Supporting Notes<textarea id="evidence" rows="2" placeholder="Documents, observations, meeting outcomes..."></textarea></label>
<label>Supporting Link (optional)<input id="link" type="url" placeholder="https://..."></label>
<button onclick="submit()">Submit for Approval</button>
<button onclick="google.script.host.close()" style="background:#888">Cancel</button>
<div id="msg"></div>
<script>
  google.script.run.withSuccessHandler(function(acts) {
    var sel = document.getElementById('actCode');
    acts.forEach(function(a) {
      var opt = document.createElement('option');
      opt.value = a.code;
      opt.textContent = a.code + ' — ' + a.name.substring(0,50);
      sel.appendChild(opt);
    });
  }).getActivityCodes();

  function loadTasks() {
    var actCode = document.getElementById('actCode').value;
    var taskSel = document.getElementById('taskCode');
    taskSel.innerHTML = '<option value="">— loading... —</option>';
    document.getElementById('taskSel').style.display = actCode ? 'block' : 'none';
    if (!actCode) return;
    google.script.run.withSuccessHandler(function(tasks) {
      taskSel.innerHTML = '<option value="">— select task —</option>';
      if (tasks.length === 0) {
        taskSel.innerHTML = '<option value="">No tasks defined yet for this activity</option>';
      } else {
        tasks.forEach(function(t) {
          var opt = document.createElement('option');
          opt.value = t.code;
          opt.textContent = t.code + ' — ' + t.name;
          taskSel.appendChild(opt);
        });
      }
    }).getTasksForActivity(actCode);
  }

  function submit() {
    var data = {
      submittedBy:  document.getElementById('submittedBy').value.trim(),
      taskCode:     document.getElementById('taskCode').value,
      pctComplete:  document.getElementById('pct').value,
      statusUpdate: document.getElementById('statusUpdate').value.trim(),
      evidenceNotes:document.getElementById('evidence').value.trim(),
      supportingLink:document.getElementById('link').value.trim()
    };
    if (!data.submittedBy || !data.taskCode || data.pctComplete === '' || !data.statusUpdate) {
      show('Please fill in all required fields (*).', 'error'); return;
    }
    google.script.run
      .withSuccessHandler(function(id) { show('Submitted! Entry ID: ' + id + '. The approver has been notified.', 'success'); })
      .withFailureHandler(function(e)  { show('Error: ' + e.message, 'error'); })
      .saveProgress(data);
  }
  function show(msg, type) {
    var el = document.getElementById('msg');
    el.textContent = msg; el.className = type; el.style.display = 'block';
  }
</script>
</body>
</html>
`;
