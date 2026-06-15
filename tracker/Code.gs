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
  SHEET_DASHBOARD:  'Dashboard',
};

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
  setupDashboardSheet(ss);
  // Move sheets into logical order
  const order = [CONFIG.SHEET_DASHBOARD, CONFIG.SHEET_ACTIVITIES, CONFIG.SHEET_TASKS, CONFIG.SHEET_PROGRESS];
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
    'Created By', 'Created Date', 'Status'
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
  sh.getRange(2, 14, 500, 1).setValue('Not Started');

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

  // Title
  sh.getRange('A1').setValue('MCA PROJECT TRACKER — DASHBOARD');
  sh.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1c4587');
  sh.getRange('B1').setFormula('=NOW()');
  sh.getRange('B1').setNumberFormat('dd MMM yyyy HH:mm').setFontColor('#666666');
  sh.getRange('A1:N1').setBackground('#c9daf8');

  sh.getRange('A2').setValue('⚠ Only APPROVED progress entries are reflected below.');
  sh.getRange('A2').setFontColor('#cc0000').setFontStyle('italic');

  // Section headers helper
  const sectionStyle = (range) => range.setFontWeight('bold').setFontColor('#ffffff').setBackground('#1c4587');

  // ---- PENDING APPROVALS ----
  let row = 4;
  sectionStyle(sh.getRange(row, 1, 1, 6).merge());
  sh.getRange(row, 1).setValue('PENDING APPROVALS');
  row++;
  const pendingHeaders = ['Entry ID', 'Task Code', 'Activity Code', 'Submitted By', 'Submitted Date', '% Complete'];
  sh.getRange(row, 1, 1, 6).setValues([pendingHeaders]).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  // QUERY formula pulling pending rows from Progress Log
  sh.getRange(row, 1).setFormula(
    `=IFERROR(QUERY('Progress Log'!A:N, "SELECT A,B,C,E,F,G WHERE K='Pending' ORDER BY F DESC LABEL A 'Entry ID', B 'Task Code', C 'Activity Code', E 'Submitted By', F 'Submitted Date', G '% Complete'", 0), "No pending entries")`
  );

  // ---- SUMMARY BY STRATEGIC OUTPUT ----
  row = 20;
  sectionStyle(sh.getRange(row, 1, 1, 5).merge());
  sh.getRange(row, 1).setValue('SUMMARY BY STRATEGIC OUTPUT (Approved entries only)');
  row++;
  const soHeaders = ['Strategic Output', 'Total Activities', 'Activities with Tasks', 'Avg % Complete', 'Status'];
  sh.getRange(row, 1, 1, 5).setValues([soHeaders]).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  ['SO1','SO2','SO3'].forEach((so, i) => {
    const r = row + i;
    const soName = so === 'SO1' ? 'SO1 – Policy & Governance' : so === 'SO2' ? 'SO2 – Awareness & Research' : 'SO3 – Livelihoods & Economy';
    sh.getRange(r, 1).setValue(soName);
    sh.getRange(r, 2).setFormula(`=COUNTIF(Activities!A:A,"${so}")`);
    sh.getRange(r, 3).setFormula(
      `=SUMPRODUCT((Activities!A$2:A$200="${so}")*(Activities!I$2:I$200>0))`
    );
    sh.getRange(r, 4).setFormula(
      `=IFERROR(AVERAGEIF(Activities!A$2:A$200,"${so}",Activities!J$2:J$200),0)`
    );
    sh.getRange(r, 4).setNumberFormat('0%');
    sh.getRange(r, 5).setFormula(
      `=IF(D${r}=0,"Not Started",IF(D${r}<0.5,"In Progress",IF(D${r}<1,"Nearly Complete","Complete")))`
    );
  });

  // ---- ACTIVITY PROGRESS TABLE ----
  row = 28;
  sectionStyle(sh.getRange(row, 1, 1, 7).merge());
  sh.getRange(row, 1).setValue('ACTIVITY PROGRESS (Approved entries only)');
  row++;
  const actHeaders = ['Code', 'Activity', 'Partner', 'CI Staff', 'Scheduled', 'Tasks', '% Complete'];
  sh.getRange(row, 1, 1, 7).setValues([actHeaders]).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  // Pull from Activities sheet with live % complete
  sh.getRange(row, 1).setFormula(
    `=IFERROR(QUERY(Activities!A:K, "SELECT C,D,G,H,E,I,J WHERE C IS NOT NULL ORDER BY A,C LABEL C 'Code', D 'Activity', G 'Partner', H 'CI Staff', E 'Scheduled', I 'Tasks', J '% Complete'", 0), "Run Refresh Dashboard first")`
  );

  // ---- BY RESPONSIBLE PARTNER ----
  row = 130;
  sectionStyle(sh.getRange(row, 1, 1, 3).merge());
  sh.getRange(row, 1).setValue('PROGRESS BY RESPONSIBLE PARTNER');
  row++;
  sh.getRange(row, 1, 1, 3).setValues([['Partner', 'Activities', 'Avg % Complete']]).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  sh.getRange(row, 1).setFormula(
    `=IFERROR(QUERY(Activities!G$2:K$200, "SELECT G, COUNT(G), AVG(J) WHERE G IS NOT NULL GROUP BY G ORDER BY G LABEL G 'Partner', COUNT(G) 'Activities', AVG(J) 'Avg % Complete'", 0), "")`
  );

  // ---- SCHEDULE STATUS ----
  row = 155;
  sectionStyle(sh.getRange(row, 1, 1, 5).merge());
  sh.getRange(row, 1).setValue('SCHEDULE STATUS — Activities by Quarter');
  row++;
  sh.getRange(row, 1, 1, 5).setValues([['Quarter', 'Count', 'Completed', 'In Progress', 'Not Started']]).setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  ['Q1 2025','Q2 2025','Q3 2025','Q4 2025','Q1 2026','Q2 2026','Q3 2026','Q4 2026','Q1 2027','Q2 2027','Q3 2027','Q4 2027','Q1 2028','Q2 2028','Q3 2028','Q4 2028'].forEach((q, i) => {
    const r = row + i;
    sh.getRange(r, 1).setValue(q);
    sh.getRange(r, 2).setFormula(`=COUNTIF(Activities!E:E,"${q}")`);
    sh.getRange(r, 3).setFormula(`=SUMPRODUCT((Activities!E$2:E$200="${q}")*(Activities!K$2:K$200="Complete"))`);
    sh.getRange(r, 4).setFormula(`=SUMPRODUCT((Activities!E$2:E$200="${q}")*(Activities!K$2:K$200="In Progress"))`);
    sh.getRange(r, 5).setFormula(`=SUMPRODUCT((Activities!E$2:E$200="${q}")*(Activities!K$2:K$200="Not Started"))`);
  });

  sh.setFrozenRows(3);
  sh.setColumnWidth(1, 120);
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
    'Not Started'
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
      weight:          document.getElementById('weight').value
    };
    if (!data.activityCode || !data.taskName || !data.assignedTo) {
      show('Please fill in all required fields (*).', 'error'); return;
    }
    google.script.run
      .withSuccessHandler(function(code) { show('Task saved! Code: ' + code, 'success'); })
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
