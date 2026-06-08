// State Management
const state = {
  propertyAddress: 'Skyview Towers, Unit 502', // House address from Layout Setup
  roomsList: ['Living Room / Hall', 'Kitchen', 'Master Bedroom', 'Guest Bathroom', 'Master Bathroom', 'Balcony'], // Selected rooms layout
  currentRoom: 'Living Room / Hall', // Room focus
  activeMode: 'field', // Redesign: Default to room-centric Field Mode
  activeCategory: 'all', // Active category toggle filter ('all' or specific CategoryName)
  activeCategoryIndex: 0, // Default focus Category: Electrical (index 0) to scroll into view
  trainingDataset: [], // Active learning dataset collection
  timerSeconds: 45 * 60,
  timerInterval: null,
  
  // Property Configuration (Toggles from setup wizard)
  config: {
    has_ac: true,
    has_geyser: true,
    has_window: true,
    has_wardrobe: true,
    has_purifier: true,
    has_softener: true,
    has_pump: true,
    has_fridge: true,
    has_chimney: true,
    has_exhaust: true,
    has_db: true,
    has_outdoor_lights: true,
    has_outdoor_taps: true,
    has_balcony_enclosure: true,
    has_poja_wood: true
  },
  
  // Checklist records
  // Key format: "RoomName::CategoryName::ComponentName"
  // Value format: { status: 'pending' | 'good' | 'issue', selectedIssues: [], photos: [] }
  checkpointStates: {},
  
  // Track open/collapsed state of category accordions
  expandedCategories: {
    "Electrical": true,
    "Plumbing": false,
    "Carpentry": false,
    "Flooring": false,
    "Painting": false,
    "Dampness & Seepage": false,
    "Structure & Civil": false,
    "Termite": false,
    "General Pest": false,
    "AC & Appliances": false,
    "Water & Fittings": false,
    "Cleaning & Hygiene": false
  },

  // Track open/collapsed state of component checklist drawers
  // Key: "RoomName::CategoryName::ComponentName" -> boolean
  expandedDrawers: {}
};

// All available categories mapped verbatim from Excel routingMatrix
const categoriesList = [
  "Electrical",
  "Plumbing",
  "Carpentry",
  "Flooring",
  "Painting",
  "Dampness & Seepage",
  "Structure & Civil",
  "Termite",
  "General Pest",
  "AC & Appliances",
  "Water & Fittings",
  "Cleaning & Hygiene"
];

// Helper to determine if a specific component is enabled based on Property Setup Configurator
function isComponentEnabled(compName) {
  const name = compName.toLowerCase();
  
  if (name.includes('air cond') || name.includes('ac /') || name.includes('ac point') || name.includes('ac & appliances')) {
    return !!state.config.has_ac;
  }
  if (name.includes('geyser') || name.includes('hot water')) {
    return !!state.config.has_geyser;
  }
  if (name.includes('window')) {
    return !!state.config.has_window;
  }
  if (name.includes('wardrobe') || name.includes('closet')) {
    return !!state.config.has_wardrobe;
  }
  if (name.includes('purifier') || name.includes('ro')) {
    return !!state.config.has_purifier;
  }
  if (name.includes('softener')) {
    return !!state.config.has_softener;
  }
  if (name.includes('pump') || name.includes('booster')) {
    return !!state.config.has_pump;
  }
  if (name.includes('fridge') || name.includes('refrigerator')) {
    return !!state.config.has_fridge;
  }
  if (name.includes('chimney') || name.includes('hood')) {
    return !!state.config.has_chimney;
  }
  if (name.includes('exhaust')) {
    return !!state.config.has_exhaust;
  }
  if (name.includes('db ') || name.includes('mcb panel') || name.includes('distribution board')) {
    return !!state.config.has_db;
  }
  if (name.includes('outdoor light') || name.includes('balcony light') || name.includes('external light')) {
    return !!state.config.has_outdoor_lights;
  }
  if (name.includes('outdoor tap') || name.includes('balcony tap') || name.includes('garden tap')) {
    return !!state.config.has_outdoor_taps;
  }
  if (name.includes('balcony glass') || name.includes('enclosure') || name.includes('glazing')) {
    return !!state.config.has_balcony_enclosure;
  }
  if (name.includes('pooja') || name.includes('mandir')) {
    return !!state.config.has_poja_wood;
  }
  
  return true;
}

// Database function providing custom checklist issues for all unique Excel items
function getIssuesForComponent(catName, compName) {
  const name = compName.toLowerCase();
  
  if (name.includes('switchboard')) {
    return [
      "Burnt or charred switch or socket face",
      "Switchboard loose or hanging from wall mount",
      "Cracked or broken front switch plate face",
      "Sparking observed when toggling switches",
      "No power / dead socket on board"
    ];
  }
  if (name.includes('socket') || name.includes('outlet') || name.includes('multi-plug') || name.includes('point')) {
    return [
      "Socket pin holes loose or worn out",
      "No power supply / dead outlet",
      "Socket casing cracked or loose",
      "Heavy load switch sparking or overheating"
    ];
  }
  if (name.includes('light fixture')) {
    return [
      "Bulb / LED tube fused or not glowing",
      "Fixture loose or hanging dangerously",
      "Choke / driver buzzing noise",
      "Diffuser cover missing or cracked"
    ];
  }
  if (name.includes('fan')) {
    return [
      "Ceiling fan wobbling excessively at high speed",
      "Noisy operation / bearing sound",
      "Regulator not controlling fan speed",
      "Fan blades accumulated grease or dirt"
    ];
  }
  if (name.includes('wiring')) {
    return [
      "Exposed wire hazard in open view",
      "Loose wire joints / temporary joints",
      "Casing or capping missing or broken",
      "Unanchored wire conduits hanging down"
    ];
  }
  if (name.includes('db ') || name.includes('mcb') || name.includes('distribution board')) {
    return [
      "DB panel door loose or broken",
      "Loose wiring connections inside panel",
      "Frequent MCB tripping under load",
      "MCB switches charred or damaged"
    ];
  }
  if (name.includes('tap') || name.includes('faucet')) {
    return [
      "Tap dripping / leaking from spout",
      "Tap base loose or leaking at body",
      "Thread worn out / handle slipping",
      "Aspirator missing or blocked with silt"
    ];
  }
  if (name.includes('sink') || name.includes('basin') || name.includes('vanity')) {
    return [
      "Sink chipped or cracked body",
      "Drain pipe leaking at bottom connection",
      "Slow drainage / water accumulation",
      "Countertop sealant peeled / water seeping"
    ];
  }
  if (name.includes('drainage') || name.includes('drain')) {
    return [
      "Floor drain slow or blocked",
      "Foul smell rising from drain",
      "Drain grating missing or broken",
      "Cockroach/pest activity inside drain"
    ];
  }
  if (name.includes('pipes') || name.includes('plumbing line')) {
    return [
      "Visible pipe leak or dampness around joints",
      "Visible pipe rust or corrosion",
      "Pipe brackets loose or missing",
      "Water hammer noises in pipes"
    ];
  }
  if (name.includes('wc ') || name.includes('toilet')) {
    return [
      "Water seepage at WC base / floor connection",
      "Continuous running / cistern leak",
      "WC cracked or physically damaged",
      "Flush button jammed or unresponsive"
    ];
  }
  if (name.includes('shower')) {
    return [
      "Shower dripping / leaking when off",
      "Shower diverter / mixer not functioning",
      "Shower head clogged / uneven spray",
      "Shower arm loose on wall fitting"
    ];
  }
  if (name.includes('door')) {
    return [
      "Door misaligned / won't close properly",
      "Lock or latch failure on frame",
      "Hinge damage / rust observed",
      "Door swollen / scraping floor bottom",
      "Door handle loose or missing"
    ];
  }
  if (name.includes('window')) {
    return [
      "Window panel scraping on frame",
      "Glass pane cracked or loose",
      "Latch or window lock broken",
      "Channel sliding track jammed or dirty"
    ];
  }
  if (name.includes('cabinet') || name.includes('wardrobe') || name.includes('shelf') || name.includes('drawer')) {
    return [
      "Cabinet swelling / wood warping",
      "Hinges rusted or loose on shutters",
      "Drawer slider jammed or hard to pull",
      "Laminate peeling off wooden panels",
      "Handles/knobs loose or missing"
    ];
  }
  if (name.includes('wood') || name.includes('paneling')) {
    return [
      "Paneling board loose / warping",
      "Fading wood polish or severe stains",
      "Hollow termite sound behind woodwork"
    ];
  }
  if ((name.includes('wall') && name.includes('paint')) || name.includes('interior wall')) {
    return [
      "Paint peeling / bubbling observed on wall",
      "Visible stains or peeling paint",
      "Wall color faded / uneven patches",
      "Scratches or cracks in plaster coat"
    ];
  }
  if (name.includes('ceiling') && name.includes('paint')) {
    return [
      "Ceiling paint peeling or bubbling",
      "Moisture stains on ceiling surface",
      "Fading or chalky paint texture"
    ];
  }
  if (name.includes('dampness') || name.includes('seepage') || name.includes('wall (internal)') || name.includes('ceiling') || name.includes('corner') || name.includes('skirting')) {
    return [
      "Wet seepage patch observed on surface",
      "Dampness seepage along floor skirting board",
      "Efflorescence/salt bubbling on plaster",
      "Black mold/fungus growth on surface",
      "Active water droplet condensation/dripping"
    ];
  }
  if (name.includes('crack')) {
    return [
      "Hairline plaster cracks on wall surface",
      "Deep structural crack in masonry",
      "Ceiling joint plaster cracking"
    ];
  }
  if (name.includes('termite')) {
    return [
      "Termite mud tubes on skirting or door frame",
      "Hollow wooden door frames detected",
      "Termite dust / frass observed under wood"
    ];
  }
  if (name.includes('cockroach') || name.includes('pest') || name.includes('rodent') || name.includes('ant') || name.includes('mosquito')) {
    return [
      "Live pest activity / sightings observed",
      "Pest droppings found in cabinets/ducts",
      "Bite marks on wiring/pipes detected",
      "Mosquito breeding in standing water"
    ];
  }
  if (name.includes('air cond') || name.includes('ac')) {
    return [
      "AC water dripping from indoor panel unit",
      "AC compressor fan noisy / vibrations",
      "Cooling insufficient / air flow weak",
      "AC remote control missing/unresponsive"
    ];
  }
  if (name.includes('geyser') || name.includes('hot water')) {
    return [
      "Geyser not heating water",
      "Water leakage from geyser safety valve",
      "Geyser indicator light failing"
    ];
  }
  if (name.includes('refrigerator') || name.includes('fridge')) {
    return [
      "Fridge not cooling properly",
      "Compressor making abnormal loud hum",
      "Gasket seal torn / door loose"
    ];
  }
  if (name.includes('chimney') || name.includes('hood')) {
    return [
      "Chimney suction power extremely weak",
      "Excessive noise/vibration from blower",
      "Baffle filters choked with heavy grease"
    ];
  }
  if (name.includes('flooring') || name.includes('tile') || name.includes('stone') || name.includes('paving')) {
    return [
      "Floor tile cracked or loose tile grout",
      "Hollow sounding tiles under tapping",
      "Stains, scratches or uneven joint levels",
      "Grout erosion between joints"
    ];
  }
  if (name.includes('cleaning') || name.includes('hygiene') || name.includes('dirt')) {
    return [
      "Accumulated dust or cobwebs in corners",
      "Debris or trash left un-cleared",
      "Severe oil stains or hard water scaling"
    ];
  }
  
  // Generic fallback issues
  return [
    "Material wear and tear / aging observed",
    "Loose fitting or misaligned mounting",
    "Functional degradation / unresponsive unit",
    "Visible rust, corrosion or surface stain"
  ];
}

// Toggles visual selection highlight on layout checkboxes
window.toggleCheckboxCardHighlight = function(input) {
  if (input) {
    input.parentElement.classList.toggle('selected', input.checked);
  }
};

// Adds a custom room option dynamically to the Layout Setup checkboxes list
window.addCustomRoomOption = function() {
  const customRoomInput = document.getElementById('input-custom-room');
  if (!customRoomInput) return;
  
  const roomName = customRoomInput.value.trim();
  if (roomName === "") {
    alert("Please enter a valid room name.");
    return;
  }
  
  const container = document.getElementById('rooms-setup-grid');
  if (!container) return;
  
  // Check if option already exists
  const existingCheckboxes = container.querySelectorAll('input[type="checkbox"]');
  let duplicate = false;
  existingCheckboxes.forEach(chk => {
    if (chk.value.toLowerCase() === roomName.toLowerCase()) {
      duplicate = true;
    }
  });
  
  if (duplicate) {
    alert("This room is already in the layout builder list.");
    return;
  }
  
  const label = document.createElement('label');
  label.className = "room-checkbox-card selected";
  label.innerHTML = `
    <input type="checkbox" value="${roomName}" checked onchange="toggleCheckboxCardHighlight(this)">
    <span>🚪 ${roomName}</span>
  `;
  container.appendChild(label);
  
  customRoomInput.value = "";
  alert(`"${roomName}" added to layout options and auto-checked!`);
};

// Transition from Setup Landing view into primary Workbench view
window.startInspectionWalkthrough = function() {
  const addressInput = document.getElementById('input-property-address');
  if (!addressInput) return;
  
  const address = addressInput.value.trim();
  if (address === "") {
    alert("Please enter the property address first.");
    return;
  }
  
  const selectedRooms = [];
  const container = document.getElementById('rooms-setup-grid');
  if (container) {
    const checked = container.querySelectorAll('input[type="checkbox"]:checked');
    checked.forEach(chk => {
      selectedRooms.push(chk.value);
    });
  }
  
  if (selectedRooms.length === 0) {
    alert("Please select at least one room for the inspection layout.");
    return;
  }
  
  // Save to state
  state.propertyAddress = address;
  state.roomsList = selectedRooms;
  
  // Re-verify active room is valid (or default to first room)
  if (!state.roomsList.includes(state.currentRoom)) {
    state.currentRoom = state.roomsList[0];
  }
  
  // Initialize checklist database records for selected rooms
  setupData();
  
  // Set default view titles
  const addressEl = document.querySelector('.logo-text p');
  if (addressEl) {
    addressEl.innerText = `Layout: ${address}`;
  }
  
  // Hide landing screen and open main container
  document.getElementById('setup-landing-screen').style.display = "none";
  document.getElementById('main-app-container').style.display = "flex";
  
  // Render
  renderTopRoomBar();
  renderSidebar();
  renderActiveRoom();
};

// Open the layout creator editor overlay
window.showSetupScreen = function() {
  // Pre-populate input
  const addressInput = document.getElementById('input-property-address');
  if (addressInput) {
    addressInput.value = state.propertyAddress || "";
  }
  
  // Sync checkboxes with state.roomsList
  const container = document.getElementById('rooms-setup-grid');
  if (container) {
    const labels = container.querySelectorAll('.room-checkbox-card');
    labels.forEach(lbl => {
      const checkbox = lbl.querySelector('input[type="checkbox"]');
      if (checkbox) {
        const isPresent = state.roomsList.includes(checkbox.value);
        checkbox.checked = isPresent;
        lbl.classList.toggle('selected', isPresent);
      }
    });
  }
  
  document.getElementById('setup-landing-screen').style.display = "flex";
  document.getElementById('main-app-container').style.display = "none";
};

// Dynamically extracts, filters, and formats components for a specific room and category
function getComponentsForRoom(roomName, catName) {
  // Find all rows in routingMatrix matching this room & category
  let rows = (inspectionData && inspectionData.routingMatrix) ? 
    inspectionData.routingMatrix.filter(row => row.room === roomName && row.category === catName) : [];
  
  // Custom Room Fallback: if no matching room rows exist in excel data, load typical templates
  if (rows.length === 0) {
    const fallbackCatalog = {
      "Electrical": ["Switchboard", "Socket / Outlet", "Light Fixture", "Fan (Ceiling)"],
      "Carpentry": ["Door", "Window"],
      "Flooring": ["Flooring Material", "Floor – Structural / Drainage"],
      "Painting": ["Interior Walls", "Ceiling Paint"],
      "Dampness & Seepage": ["Wall (internal)", "Ceiling", "Corners & Joints"],
      "Structure & Civil": ["Wall Cracks"],
      "Termite": ["Termite Inspection"],
      "Cleaning & Hygiene": ["Cleaning & Hygiene Item"]
    };
    
    const items = fallbackCatalog[catName] || [];
    rows = items.map(item => ({ room: roomName, category: catName, item: item }));
  }
  
  const comps = [];
  const uniqueItems = new Set();
  
  rows.forEach(row => {
    if (uniqueItems.has(row.item)) return;
    uniqueItems.add(row.item);
    
    // Ignore dummy placeholder rows
    if (row.item.includes('items TBD') || row.item.includes('(items TBD)')) return;
    
    // Filter based on Property Setup Configurator
    if (!isComponentEnabled(row.item)) return;
    
    comps.push({
      name: row.item,
      issues: getIssuesForComponent(catName, row.item)
    });
  });
  
  return comps;
}

// Build initial records data dynamically for selected rooms
function setupData() {
  state.roomsList.forEach(roomName => {
    initializeCheckpointStatesForRoom(roomName);
  });
}

function initializeCheckpointStatesForRoom(roomName) {
  categoriesList.forEach(cat => {
    const components = getComponentsForRoom(roomName, cat);
    components.forEach(comp => {
      const key = `${roomName}::${cat}::${comp.name}`;
      if (!state.checkpointStates[key]) {
        state.checkpointStates[key] = {
          status: 'pending',
          selectedIssues: [],
          photos: [],
          notes: ''
        };
        state.expandedDrawers[key] = false;
      }
    });
  });
}

// Initialize App (Shows Layout Setup Screen first!)
document.addEventListener('DOMContentLoaded', () => {
  setupTimer();
  showSetupScreen(); // Load the layout creation screen by default
  setupEventListeners();
});

// Get unique rooms in layout
function getUniqueRooms() {
  return state.roomsList;
}

// Render left Sidebar rooms navigation list (no back button required!)
// Render left Sidebar categories menu (Layer 2) with completion stats for the active room
function renderSidebar() {
  const container = document.getElementById('category-sidebar-list');
  if (!container) return;
  container.innerHTML = '';
  
  // Add an "All Categories" toggle item at the very top of the sidebar list
  const showAllActive = state.activeCategory === 'all';
  
  // Compute stats for "All Categories"
  let allTotal = 0;
  let allChecked = 0;
  let allIssues = 0;
  
  categoriesList.forEach(cat => {
    const components = getComponentsForRoom(state.currentRoom, cat);
    components.forEach(comp => {
      allTotal++;
      const key = `${state.currentRoom}::${cat}::${comp.name}`;
      const rec = state.checkpointStates[key] || { status: 'pending' };
      if (rec.status !== 'pending') allChecked++;
      if (rec.status === 'issue' || rec.status === 'minor' || rec.status === 'major') allIssues++;
    });
  });
  
  let allStatusClass = '';
  if (allChecked === allTotal && allTotal > 0) {
    allStatusClass = 'verified';
  } else if (allIssues > 0) {
    allStatusClass = 'has-issues';
  } else if (allChecked > 0) {
    allStatusClass = 'in-progress';
  }
  
  const allDiv = document.createElement('a');
  allDiv.className = `category-sidebar-item ${showAllActive ? 'active' : ''} ${allStatusClass}`;
  allDiv.href = '#';
  allDiv.innerHTML = `
    <div class="category-sidebar-left">
      <span class="category-item-icon">📂</span>
      <span class="category-item-name">All Categories</span>
    </div>
    <div class="category-sidebar-right">
      <span class="category-progress-text">${allChecked}/${allTotal}</span>
      <span class="category-status-dot"></span>
    </div>
  `;
  allDiv.addEventListener('click', (e) => {
    e.preventDefault();
    state.activeCategory = 'all';
    renderSidebar();
    renderActiveRoom();
  });
  container.appendChild(allDiv);
  
  // Render specific categories
  categoriesList.forEach(catName => {
    const isActive = state.activeCategory === catName;
    const components = getComponentsForRoom(state.currentRoom, catName);
    
    // Calculate stats
    let total = components.length;
    let checked = 0;
    let issues = 0;
    
    components.forEach(comp => {
      const key = `${state.currentRoom}::${catName}::${comp.name}`;
      const rec = state.checkpointStates[key] || { status: 'pending' };
      if (rec.status !== 'pending') checked++;
      if (rec.status === 'issue' || rec.status === 'minor' || rec.status === 'major') issues++;
    });
    
    let statusClass = '';
    if (checked === total && total > 0) {
      statusClass = 'verified';
    } else if (issues > 0) {
      statusClass = 'has-issues';
    } else if (checked > 0) {
      statusClass = 'in-progress';
    }
    
    const div = document.createElement('a');
    div.className = `category-sidebar-item ${isActive ? 'active' : ''} ${statusClass}`;
    div.href = '#';
    div.innerHTML = `
      <div class="category-sidebar-left">
        <span class="category-item-icon">${getCategorySidebarIcon(catName)}</span>
        <span class="category-item-name">${catName}</span>
      </div>
      <div class="category-sidebar-right">
        <span class="category-progress-text">${checked}/${total}</span>
        <span class="category-status-dot"></span>
      </div>
    `;
    
    div.addEventListener('click', (e) => {
      e.preventDefault();
      state.activeCategory = catName;
      renderSidebar();
      renderActiveRoom();
      
      // If in Audit Mode, automatically expand this category's accordion and scroll to it
      if (state.activeMode === 'audit') {
        state.expandedCategories[catName] = true;
        renderAccordions();
        setTimeout(() => {
          const el = document.getElementById(`accordion-${catName.replace(/[^a-zA-Z0-9]/g, '')}`);
          if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }, 100);
      }
    });
    
    container.appendChild(div);
  });
}

// Category sidebar icon helper
function getCategorySidebarIcon(cat) {
  const c = cat.toLowerCase();
  if (c.includes('electrical')) return '⚡';
  if (c.includes('plumbing')) return '🚰';
  if (c.includes('carpentry')) return '🚪';
  if (c.includes('flooring')) return '🧱';
  if (c.includes('painting')) return '🎨';
  if (c.includes('dampness')) return '💧';
  if (c.includes('civil') || c.includes('structure')) return '🏗️';
  if (c.includes('termite')) return '🪳';
  if (c.includes('pest')) return '🐜';
  if (c.includes('appliances') || c.includes('ac')) return '❄️';
  if (c.includes('water')) return '🚿';
  if (c.includes('cleaning')) return '🧹';
  return '📦';
}

// Get room icons
function getRoomIcon(room) {
  const r = room.toLowerCase();
  if (r.includes('living') || r.includes('hall')) return '🛋️';
  if (r.includes('dining')) return '🪑';
  if (r.includes('kitchen')) return '🍳';
  if (r.includes('bathroom')) return '🚿';
  if (r.includes('bedroom')) return '🛏️';
  if (r.includes('balcony')) return '🌅';
  if (r.includes('utility') || r.includes('service')) return '🧺';
  if (r.includes('pooja')) return '🙏';
  if (r.includes('staircase')) return '🪜';
  if (r.includes('terrace')) return '🪴';
  if (r.includes('garden') || r.includes('compound')) return '🌳';
  return '🚪';
}

// Active Room render
function renderActiveRoom() {
  if (typeof window.stopWebcamStream === 'function') {
    window.stopWebcamStream();
    window.clearAiSweepIntervals();
  }
  
  document.getElementById('banner-active-title').innerHTML = `${getRoomIcon(state.currentRoom)} ${state.currentRoom}`;
  
  let totalIssuesCount = 0;
  let totalRoomItems = 0;
  categoriesList.forEach(cat => {
    const components = getComponentsForRoom(state.currentRoom, cat);
    components.forEach(comp => {
      totalRoomItems++;
      const key = `${state.currentRoom}::${cat}::${comp.name}`;
      const rec = state.checkpointStates[key];
      if (rec && (rec.status === 'issue' || rec.status === 'minor' || rec.status === 'major')) {
        totalIssuesCount++;
      }
    });
  });
  
  const subtitleText = (state.activeMode === 'field' || state.activeMode === 'express') ?
    `${totalRoomItems} assets · ${totalIssuesCount} issues flagged` :
    `${categoriesList.length} sections · ${totalIssuesCount} issues flagged`;
  
  document.getElementById('banner-sections-count').innerText = subtitleText;
  
  renderTopRoomBar(); // Render Room selection pills at the top
  updateCircularProgressRing();
  renderAccordions();
  renderBottomNavBar();
}

// Render top horizontal Room tabs (Layer 1)
function renderTopRoomBar() {
  const container = document.getElementById('top-room-selector-bar');
  if (!container) return;
  container.innerHTML = '';
  
  state.roomsList.forEach(roomName => {
    const isActive = roomName === state.currentRoom;
    
    // Calculate stats for room
    let total = 0;
    let checked = 0;
    let issues = 0;
    
    categoriesList.forEach(cat => {
      const components = getComponentsForRoom(roomName, cat);
      components.forEach(comp => {
        total++;
        const key = `${roomName}::${cat}::${comp.name}`;
        const rec = state.checkpointStates[key] || { status: 'pending' };
        if (rec.status !== 'pending') checked++;
        if (rec.status === 'issue' || rec.status === 'minor' || rec.status === 'major') issues++;
      });
    });
    
    let statusClass = '';
    if (checked === total && total > 0) {
      statusClass = 'verified';
    } else if (issues > 0) {
      statusClass = 'has-issues';
    } else if (checked > 0) {
      statusClass = 'in-progress';
    }
    
    const btn = document.createElement('button');
    btn.className = `top-room-pill ${isActive ? 'active' : ''} ${statusClass}`;
    btn.innerHTML = `
      <span>${getRoomIcon(roomName)}</span>
      <strong>${roomName}</strong>
      <span class="room-progress-dot"></span>
    `;
    btn.addEventListener('click', () => {
      state.currentRoom = roomName;
      renderTopRoomBar();
      renderSidebar(); // Refresh category percentages for the new active room!
      renderActiveRoom();
    });
    
    container.appendChild(btn);
  });
}

// Update Banner Circle Progress
function updateCircularProgressRing() {
  let totalItems = 0;
  let checkedItems = 0;
  
  categoriesList.forEach(cat => {
    const components = getComponentsForRoom(state.currentRoom, cat);
    components.forEach(comp => {
      totalItems++;
      const key = `${state.currentRoom}::${cat}::${comp.name}`;
      const rec = state.checkpointStates[key];
      if (rec && rec.status !== 'pending') {
        checkedItems++;
      }
    });
  });
  
  const progressPct = totalItems > 0 ? Math.round((checkedItems / totalItems) * 100) : 0;
  document.getElementById('progress-ring-percentage').innerText = `${progressPct}%`;
  
  const strokeOffset = 175.9 - (progressPct / 100) * 175.9;
  document.getElementById('progress-ring-circle').style.strokeDashoffset = strokeOffset;
}

// Walkthrough stage categorization logic based on natural technician movement
function getWalkthroughStageForComponent(catName, compName) {
  const name = compName.toLowerCase();
  
  // Stage 1: Entry Check
  if (name.includes('door') && !name.includes('cabinet') && !name.includes('wardrobe') && !name.includes('shelf') && !name.includes('drawer')) {
    return 'Stage 1: Entry Check';
  }
  if (name.includes('switchboard') || name.includes('mcb') || name.includes('db ') || name.includes('distribution board') || name.includes('wiring')) {
    return 'Stage 1: Entry Check';
  }
  
  // Stage 3: Ceiling Inspection
  if (name.includes('light fixture') || (name.includes('fan') && !name.includes('exhaust')) || name.includes('ceiling') || name.includes('slab')) {
    return 'Stage 3: Ceiling Inspection';
  }
  
  // Stage 4: Floor Inspection
  if (name.includes('flooring') || name.includes('tile') || name.includes('stone') || name.includes('drain') || name.includes('skirting') || name.includes('termite') || catName.toLowerCase().includes('pest') || catName.toLowerCase().includes('cleaning')) {
    return 'Stage 4: Floor Inspection';
  }
  
  // Stage 2: Clockwise Room Sweep (default fallback)
  return 'Stage 2: Clockwise Room Sweep';
}

// Map component name to visual emojis
function getComponentEmoji(compName) {
  const name = compName.toLowerCase();
  if (name.includes('door')) return '🚪';
  if (name.includes('switchboard')) return '🎛️';
  if (name.includes('socket') || name.includes('outlet') || name.includes('plug') || name.includes('extension')) return '🔌';
  if (name.includes('light')) return '💡';
  if (name.includes('fan')) return '🌀';
  if (name.includes('window')) return '🪟';
  if (name.includes('wardrobe') || name.includes('cabinet') || name.includes('closet') || name.includes('shelf') || name.includes('drawer')) return '🗄️';
  if (name.includes('faucet') || name.includes('tap')) return '🚰';
  if (name.includes('sink') || name.includes('basin') || name.includes('vanity')) return '🦪';
  if (name.includes('wc') || name.includes('toilet')) return '🚽';
  if (name.includes('shower')) return '🚿';
  if (name.includes('drain')) return '🕳️';
  if (name.includes('flooring') || name.includes('tile') || name.includes('stone')) return '🧱';
  if (name.includes('ceiling') || name.includes('slab')) return '🏠';
  if (name.includes('paint') || name.includes('wall')) return '🎨';
  if (name.includes('dampness') || name.includes('seepage')) return '💧';
  if (name.includes('termite') || name.includes('pest')) return '🪳';
  return '📦';
}

// Redesign Mode Selector controller
window.setInspectionMode = function(mode) {
  state.activeMode = mode;
  
  document.getElementById('mode-btn-field').classList.toggle('active', mode === 'field');
  document.getElementById('mode-btn-express').classList.toggle('active', mode === 'express');
  document.getElementById('mode-btn-audit').classList.toggle('active', mode === 'audit');
  
  const titleEl = document.getElementById('workspace-view-title');
  if (titleEl) {
    if (mode === 'field') {
      titleEl.innerText = 'Walkpath Movement Checklist';
    } else if (mode === 'express') {
      titleEl.innerText = 'Express 1-Tap Grid';
    } else if (mode === 'audit') {
      titleEl.innerText = 'Full Audit Category Accordions';
    }
  }
  
  renderActiveRoom();
  renderSidebar();
};

// Render Category Accordion Panels or walkpath groups containing Component Cards
function renderAccordions() {
  const container = document.getElementById('accordions-grid-container');
  container.innerHTML = '';
  
  // ==========================================
  // 1. FIELD MODE
  // ==========================================
  if (state.activeMode === 'field') {
    const stages = {
      "Stage 1: Entry Check": [],
      "Stage 2: Clockwise Room Sweep": [],
      "Stage 3: Ceiling Inspection": [],
      "Stage 4: Floor Inspection": []
    };
    
    categoriesList.forEach(catName => {
      if (state.activeCategory !== 'all' && state.activeCategory !== catName) return;
      const components = getComponentsForRoom(state.currentRoom, catName);
      components.forEach(comp => {
        const stage = getWalkthroughStageForComponent(catName, comp.name);
        stages[stage].push({
          categoryName: catName,
          name: comp.name,
          issues: comp.issues
        });
      });
    });
    
    let stageIndex = 1;
    let totalRoomItems = 0;
    let checkedRoomItems = 0;
    
    Object.keys(stages).forEach(stageName => {
      const comps = stages[stageName];
      if (comps.length === 0) return;
      
      totalRoomItems += comps.length;
      
      const stageContainer = document.createElement('div');
      stageContainer.className = 'walkpath-stage-container';
      stageContainer.innerHTML = `
        <div class="walkpath-stage-title-row">
          <div class="walkpath-stage-num-badge">${stageIndex}</div>
          <div class="walkpath-stage-name">${stageName}</div>
        </div>
        <div class="stage-body" style="display:flex; flex-direction:column; gap:12px; margin-bottom: 20px;"></div>
      `;
      
      const body = stageContainer.querySelector('.stage-body');
      
      comps.forEach((comp, idx) => {
        const key = `${state.currentRoom}::${comp.categoryName}::${comp.name}`;
        const record = state.checkpointStates[key] || { status: 'pending', selectedIssues: [], photos: [], notes: '' };
        if (record.status !== 'pending') checkedRoomItems++;
        
        const isDrawerOpen = !!state.expandedDrawers[key];
        const cardBadgeNum = (idx + 1).toString().padStart(2, '0');
        
        const card = document.createElement('div');
        card.className = `checkpoint-card status-${record.status}`;
        card.id = `card-${key.replace(/[^a-zA-Z0-9]/g, '')}`;
        
        card.innerHTML = `
          <div style="width:100%; display:flex; flex-direction:column;">
            <div class="checkpoint-card-left" onclick="toggleCheckpointControls('${key}')" style="display:flex; justify-content:space-between; align-items:center; width:100%;">
              <div style="display:flex; align-items:center; gap:16px;">
                <div class="checkpoint-index">${cardBadgeNum}</div>
                <div style="display: flex; flex-direction: column; text-align: left;">
                  <span class="checkpoint-title">${getComponentEmoji(comp.name)} ${comp.name}</span>
                  <span style="font-size: 0.65rem; color: var(--text-muted); text-transform: uppercase;">${comp.categoryName}</span>
                </div>
              </div>
              <div class="checkpoint-card-right" style="display:flex; align-items:center; gap:8px;">
                ${record.status === 'pending' ? `<span class="badge-pill badge-inspect">🕒 Inspect</span>` : ''}
                ${record.status === 'good' ? `<span class="badge-pill badge-healthy">🟩 Healthy</span>` : ''}
                ${record.status === 'minor' ? `<span class="badge-pill badge-minor">⚠️ Minor</span>` : ''}
                ${record.status === 'major' ? `<span class="badge-pill badge-defect">🟥 Major</span>` : ''}
                <span style="font-size:0.7rem; color:hsl(220, 15%, 60%); margin-right:4px; transform: rotate(${isDrawerOpen ? '0' : '-90'}deg); transition: transform var(--transition-fast); display: inline-block;">▼</span>
              </div>
            </div>
            
            <div class="inline-assessment-row" style="margin-top: 10px; display: flex; gap: 8px;">
              <button class="btn-assess btn-assess-good ${record.status === 'good' ? 'active' : ''}" onclick="event.stopPropagation(); setCheckpointRedesignStatus('${comp.categoryName}', '${comp.name}', 'good')">
                ✓ OK
              </button>
              <button class="btn-assess btn-assess-minor ${record.status === 'minor' ? 'active' : ''}" onclick="event.stopPropagation(); setCheckpointRedesignStatus('${comp.categoryName}', '${comp.name}', 'minor')">
                ⚠ Minor
              </button>
              <button class="btn-assess btn-assess-major ${record.status === 'major' ? 'active' : ''}" onclick="event.stopPropagation(); setCheckpointRedesignStatus('${comp.categoryName}', '${comp.name}', 'major')">
                ❌ Major
              </button>
            </div>
            
            <div class="checkpoint-drawer" id="drawer-${key.replace(/[^a-zA-Z0-9]/g, '')}" style="display: ${isDrawerOpen ? 'flex' : 'none'}; margin-top:12px; border-top:1px dashed var(--bg-border); padding-top:10px;">
              ${(record.status === 'minor' || record.status === 'major') ? `
                <div style="margin-top:6px; display:flex; flex-direction:column; gap:10px; width: 100%;">
                  <span style="font-size:0.75rem; font-weight:700; color:var(--text-secondary); text-transform:uppercase; text-align: left;">Select Observed Issues:</span>
                  <div class="issue-items-list" style="display:flex; flex-direction:column; gap:6px;">
                    ${comp.issues.map(issText => {
                      const isChecked = record.selectedIssues.includes(issText);
                      return `
                        <label class="issue-checkbox-label ${isChecked ? 'selected' : ''}" style="padding: 8px 12px; border-radius:6px; border:1px solid var(--bg-border); display:flex; align-items:center; gap:10px; cursor:pointer;">
                          <input type="checkbox" style="width:16px; height:16px; accent-color:var(--danger);" ${isChecked ? 'checked' : ''} 
                            onchange="toggleSpecificIssueSelection('${comp.categoryName}', '${comp.name}', '${issText}', this.checked)">
                          <span style="font-size:0.82rem; color:var(--text-primary); font-weight:600; text-align: left;">${issText}</span>
                        </label>
                      `;
                    }).join('')}
                  </div>
                  
                  <div style="margin-top: 6px;">
                    <div class="photo-uploader" onclick="mockPhotoUpload('${comp.categoryName}', '${comp.name}')">
                      <span>📷 Capture Defect Proof Photo</span>
                    </div>
                    <div class="photo-thumbnail-container" id="photos-${key.replace(/[^a-zA-Z0-9]/g, '')}"></div>
                  </div>
                  
                  <div style="margin-top: 10px; display:flex; flex-direction:column; gap:6px;">
                    <span style="font-size:0.75rem; font-weight:700; color:var(--text-secondary); text-transform:uppercase; text-align: left;">Inspection Notes:</span>
                    <textarea class="note-textarea" style="width:100%; height:60px; padding:8px; border-radius:6px; border:1px solid var(--bg-border); font-size:0.8rem; font-family:var(--font-body); resize:vertical;" 
                      placeholder="Add descriptive notes..." oninput="saveCheckpointNote('${comp.categoryName}', '${comp.name}', this.value)">${record.notes || ''}</textarea>
                  </div>
                  
                  <div style="margin-top: 10px; display: flex; justify-content: flex-end;">
                    <button type="button" class="btn btn-primary" onclick="event.stopPropagation(); closeCheckpointDrawer('${key}')" 
                      style="padding: 8px 16px; font-size: 0.8rem; border-radius: 8px; background-color: var(--success); border-color: var(--success); color: #ffffff; font-weight: bold; width: 100%;">
                      ✔️ Done & Collapse Card
                    </button>
                  </div>
                </div>
              ` : ''}
            </div>
          </div>
        `;
        
        if ((record.status === 'minor' || record.status === 'major') && record.photos && record.photos.length > 0) {
          setTimeout(() => {
            const photoThumbContainer = card.querySelector(`#photos-${key.replace(/[^a-zA-Z0-9]/g, '')}`);
            if (photoThumbContainer) {
              photoThumbContainer.innerHTML = '';
              record.photos.forEach((ph, pIdx) => {
                const thumb = document.createElement('div');
                thumb.style.position = 'relative';
                thumb.innerHTML = `
                  <img src="${ph}" class="photo-thumbnail">
                  <div class="delete-photo" onclick="event.stopPropagation(); deleteMockPhoto('${comp.categoryName}', '${comp.name}', ${pIdx})" 
                    style="position:absolute; top:-6px; right:-6px; background-color:var(--danger); width:18px; height:18px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:12px; cursor:pointer;">×</div>
                `;
                photoThumbContainer.appendChild(thumb);
              });
            }
          }, 0);
        }
        
        body.appendChild(card);
      });
      
      container.appendChild(stageContainer);
      stageIndex++;
    });
    
    // Check if room is complete (Stage 5)
    if (totalRoomItems > 0 && checkedRoomItems === totalRoomItems) {
      const roomCompleteContainer = document.createElement('div');
      roomCompleteContainer.className = 'walkpath-stage-container';
      roomCompleteContainer.style.marginTop = '30px';
      roomCompleteContainer.innerHTML = `
        <div class="walkpath-stage-title-row">
          <div class="walkpath-stage-num-badge" style="background-color: var(--success);">🎉</div>
          <div class="walkpath-stage-name" style="color: var(--success);">Stage 5: Room Complete</div>
        </div>
        <div style="background-color: var(--success-bg); border: 1px solid var(--success); border-radius: 12px; padding: 20px; text-align: center; color: var(--success); font-weight: bold; margin-bottom: 20px;">
          <p style="margin: 0 0 12px 0; font-size: 1rem;">All assets in ${state.currentRoom} have been inspected successfully!</p>
          <button class="btn btn-primary" onclick="goToNextRoom()" style="background-color: var(--bg-sidebar); border-color: var(--bg-sidebar); padding: 8px 20px; border-radius: 20px; font-size: 0.8rem; height: auto;">
            Next Room ➡️
          </button>
        </div>
      `;
      container.appendChild(roomCompleteContainer);
    }
  }
  
  // ==========================================
  // 2. EXPRESS MODE
  // ==========================================
  else if (state.activeMode === 'express') {
    const gridContainer = document.createElement('div');
    gridContainer.className = 'express-grid-layout';
    
    let totalRoomItems = 0;
    let checkedRoomItems = 0;
    
    categoriesList.forEach(catName => {
      if (state.activeCategory !== 'all' && state.activeCategory !== catName) return;
      const components = getComponentsForRoom(state.currentRoom, catName);
      components.forEach(comp => {
        totalRoomItems++;
        const key = `${state.currentRoom}::${catName}::${comp.name}`;
        const record = state.checkpointStates[key] || { status: 'pending', selectedIssues: [], photos: [], notes: '' };
        if (record.status !== 'pending') checkedRoomItems++;
        
        const card = document.createElement('div');
        card.className = `express-asset-card express-${record.status}`;
        card.id = `card-express-${key.replace(/[^a-zA-Z0-9]/g, '')}`;
        
        card.innerHTML = `
          <div class="express-card-badge" style="color: #fff; background-color: ${record.status === 'good' ? 'var(--success)' : record.status === 'minor' ? 'hsl(35, 90%, 55%)' : record.status === 'major' ? 'var(--danger)' : 'var(--text-muted)'}">
            ${record.status === 'good' ? '✓' : record.status === 'minor' ? '⚠' : record.status === 'major' ? '❌' : '?'}
          </div>
          <div style="font-size: 2rem; margin-bottom: 4px;">${getComponentEmoji(comp.name)}</div>
          <div class="express-card-title">${comp.name}</div>
          <div class="express-card-stage-label">${getWalkthroughStageForComponent(catName, comp.name)}</div>
          
          <button class="express-gear-btn" onclick="event.stopPropagation(); openExpressIssueSheet('${catName}', '${comp.name}')">⚙️</button>
        `;
        
        // Setup mouse / touch interactions for Single Tap vs Long Press
        let isLongPress = false;
        let pressTimer;
        
        const startPress = (e) => {
          isLongPress = false;
          pressTimer = setTimeout(() => {
            isLongPress = true;
            openExpressIssueSheet(catName, comp.name);
          }, 600);
        };
        
        const endPress = (e) => {
          clearTimeout(pressTimer);
          if (!isLongPress) {
            setCheckpointStatusExpressOK(catName, comp.name);
          }
        };
        
        card.addEventListener('mousedown', startPress);
        card.addEventListener('mouseup', endPress);
        card.addEventListener('touchstart', startPress);
        card.addEventListener('touchend', endPress);
        
        gridContainer.appendChild(card);
      });
    });
    
    container.appendChild(gridContainer);
    
    // Check if room is complete (Stage 5)
    if (totalRoomItems > 0 && checkedRoomItems === totalRoomItems) {
      const roomCompleteContainer = document.createElement('div');
      roomCompleteContainer.className = 'walkpath-stage-container';
      roomCompleteContainer.style.marginTop = '30px';
      roomCompleteContainer.innerHTML = `
        <div class="walkpath-stage-title-row">
          <div class="walkpath-stage-num-badge" style="background-color: var(--success);">🎉</div>
          <div class="walkpath-stage-name" style="color: var(--success);">Stage 5: Room Complete</div>
        </div>
        <div style="background-color: var(--success-bg); border: 1px solid var(--success); border-radius: 12px; padding: 20px; text-align: center; color: var(--success); font-weight: bold; margin-bottom: 20px;">
          <p style="margin: 0 0 12px 0; font-size: 1rem;">All assets in ${state.currentRoom} have been inspected successfully!</p>
          <button class="btn btn-primary" onclick="goToNextRoom()" style="background-color: var(--bg-sidebar); border-color: var(--bg-sidebar); padding: 8px 20px; border-radius: 20px; font-size: 0.8rem; height: auto;">
            Next Room ➡️
          </button>
        </div>
      `;
      container.appendChild(roomCompleteContainer);
    }
  }
  
  // ==========================================
  // 3. AUDIT MODE
  // ==========================================
  else if (state.activeMode === 'audit') {
    categoriesList.forEach((catName, cIdx) => {
      if (state.activeCategory !== 'all' && state.activeCategory !== catName) return;
      const isExpanded = !!state.expandedCategories[catName] || state.activeCategory === catName;
      const components = getComponentsForRoom(state.currentRoom, catName);
      
      if (components.length === 0) return;
      
      const catBadgeNum = (cIdx + 1).toString().padStart(2, '0');
      
      let catTotal = components.length;
      let catChecked = 0;
      let catIssues = 0;
      
      components.forEach(comp => {
        const key = `${state.currentRoom}::${catName}::${comp.name}`;
        const rec = state.checkpointStates[key] || { status: 'pending' };
        if (rec.status !== 'pending') catChecked++;
        if (rec.status === 'issue' || rec.status === 'minor' || rec.status === 'major') catIssues++;
      });
      
      const catPct = catTotal > 0 ? Math.round((catChecked / catTotal) * 100) : 0;
      
      const accordion = document.createElement('div');
      accordion.className = `section-accordion ${isExpanded ? '' : 'collapsed'}`;
      accordion.id = `accordion-${catName.replace(/[^a-zA-Z0-9]/g, '')}`;
      
      accordion.innerHTML = `
        <div class="accordion-header" onclick="toggleCategoryAccordion('${catName}')">
          <div class="accordion-header-left">
            <div class="index-badge">${catBadgeNum}</div>
            <div class="accordion-title-block">
              <span class="accordion-title">${catName}</span>
              <span class="accordion-subtitle">${catChecked} of ${catTotal} checked</span>
            </div>
          </div>
          <div class="accordion-header-right">
            <span class="accordion-pct-badge">${catPct}%</span>
            <span class="accordion-chevron" style="font-size:0.7rem; color:hsl(220, 15%, 60%);">▼</span>
          </div>
        </div>
        <div class="accordion-body"></div>
      `;
      
      const body = accordion.querySelector('.accordion-body');
      
      components.forEach((comp, idx) => {
        const key = `${state.currentRoom}::${catName}::${comp.name}`;
        const record = state.checkpointStates[key] || { status: 'pending', selectedIssues: [], photos: [], notes: '' };
        const checkBadgeNum = (idx + 1).toString().padStart(2, '0');
        const isDrawerOpen = !!state.expandedDrawers[key];
        
        const card = document.createElement('div');
        card.className = `checkpoint-card status-${record.status}`;
        card.id = `card-${key.replace(/[^a-zA-Z0-9]/g, '')}`;
        
        const hasIssue = record.status === 'issue' || record.status === 'minor' || record.status === 'major';
        
        card.innerHTML = `
          <div style="width:100%; display:flex; flex-direction:column;">
            <div class="checkpoint-card-left" onclick="toggleCheckpointControls('${key}')" style="display:flex; justify-content:space-between; align-items:center; width:100%;">
              <div style="display:flex; align-items:center; gap:16px;">
                <div class="checkpoint-index">${checkBadgeNum}</div>
                <span class="checkpoint-title">${comp.name}</span>
              </div>
              <div class="checkpoint-card-right" style="display:flex; align-items:center; gap:8px;">
                ${record.status === 'pending' ? `<span class="badge-pill badge-inspect">🕒 Inspect</span>` : ''}
                ${record.status === 'good' ? `<span class="badge-pill badge-healthy">🟩 Healthy</span>` : ''}
                ${record.status === 'minor' ? `<span class="badge-pill badge-minor">⚠️ Minor</span>` : ''}
                ${record.status === 'major' ? `<span class="badge-pill badge-defect">🟥 Major</span>` : ''}
                ${record.status === 'issue' ? `<span class="badge-pill badge-defect">🟥 Defect</span>` : ''}
                <span style="font-size:0.7rem; color:hsl(220, 15%, 60%); margin-right:4px; transform: rotate(${isDrawerOpen ? '0' : '-90'}deg); transition: transform var(--transition-fast); display: inline-block;">▼</span>
              </div>
            </div>
            
            <div class="checkpoint-drawer" id="drawer-${key.replace(/[^a-zA-Z0-9]/g, '')}" style="display: ${isDrawerOpen ? 'flex' : 'none'}; margin-top:12px; border-top:1px dashed var(--bg-border); padding-top:10px;">
              <div class="inline-toggle-box">
                <button class="toggle-btn toggle-btn-good ${record.status === 'good' ? 'active' : ''}" onclick="setCheckpointStatus('${catName}', '${comp.name}', 'good')">
                  🟩 Good Condition
                </button>
                <button class="toggle-btn toggle-btn-issue ${hasIssue ? 'active' : ''}" onclick="setCheckpointStatus('${catName}', '${comp.name}', 'issue')">
                  🟥 Issue Found
                </button>
              </div>
              
              ${hasIssue ? `
                <div style="margin-top:12px; display:flex; flex-direction:column; gap:10px; width: 100%;">
                  <div style="display:flex; gap:10px; align-items:center;">
                    <span style="font-size:0.75rem; font-weight:700; color:var(--text-secondary); text-transform:uppercase;">Refine Severity:</span>
                    <button class="btn btn-secondary btn-sm ${record.status === 'minor' ? 'active' : ''}" onclick="setCheckpointStatus('${catName}', '${comp.name}', 'minor')" style="padding:4px 8px; font-size:0.7rem; font-weight:bold; height: auto;">⚠️ Minor</button>
                    <button class="btn btn-secondary btn-sm ${record.status === 'major' || record.status === 'issue' ? 'active' : ''}" onclick="setCheckpointStatus('${catName}', '${comp.name}', 'major')" style="padding:4px 8px; font-size:0.7rem; font-weight:bold; background-color:rgba(220,38,38,0.1); border-color:var(--danger); color:var(--danger); height: auto;">🟥 Major</button>
                  </div>
                  
                  <span style="font-size:0.75rem; font-weight:700; color:var(--text-secondary); text-transform:uppercase; text-align: left;">Select Observed Issues:</span>
                  <div class="issue-items-list" style="display:flex; flex-direction:column; gap:6px;">
                    ${comp.issues.map(issText => {
                      const isChecked = record.selectedIssues.includes(issText);
                      return `
                        <label class="issue-checkbox-label ${isChecked ? 'selected' : ''}" style="padding: 8px 12px; border-radius:6px; border:1px solid var(--bg-border); display:flex; align-items:center; gap:10px; cursor:pointer;">
                          <input type="checkbox" style="width:16px; height:16px; accent-color:var(--danger);" ${isChecked ? 'checked' : ''} 
                            onchange="toggleSpecificIssueSelection('${catName}', '${comp.name}', '${issText}', this.checked)">
                          <span style="font-size:0.82rem; color:var(--text-primary); font-weight:600; text-align: left;">${issText}</span>
                        </label>
                      `;
                    }).join('')}
                  </div>
                  
                  <div style="margin-top: 6px;">
                    <div class="photo-uploader" onclick="mockPhotoUpload('${catName}', '${comp.name}')">
                      <span>📷 Capture Defect Proof Photo</span>
                    </div>
                    <div class="photo-thumbnail-container" id="photos-${key.replace(/[^a-zA-Z0-9]/g, '')}"></div>
                  </div>
                  
                  <div style="margin-top: 10px; display:flex; flex-direction:column; gap:6px;">
                    <span style="font-size:0.75rem; font-weight:700; color:var(--text-secondary); text-transform:uppercase; text-align: left;">Inspection Notes:</span>
                    <textarea class="note-textarea" style="width:100%; height:60px; padding:8px; border-radius:6px; border:1px solid var(--bg-border); font-size:0.8rem; font-family:var(--font-body); resize:vertical;" 
                      placeholder="Add descriptive notes..." oninput="saveCheckpointNote('${catName}', '${comp.name}', this.value)">${record.notes || ''}</textarea>
                  </div>
                  
                  <div style="margin-top: 10px; display: flex; justify-content: flex-end;">
                    <button type="button" class="btn btn-primary" onclick="event.stopPropagation(); closeCheckpointDrawer('${key}')" 
                      style="padding: 8px 16px; font-size: 0.8rem; border-radius: 8px; background-color: var(--success); border-color: var(--success); color: #ffffff; font-weight: bold; width: 100%;">
                      ✔️ Done & Collapse Card
                    </button>
                  </div>
                </div>
              ` : ''}
            </div>
          </div>
        `;
        
        if (hasIssue && record.photos && record.photos.length > 0) {
          setTimeout(() => {
            const photoThumbContainer = card.querySelector(`#photos-${key.replace(/[^a-zA-Z0-9]/g, '')}`);
            if (photoThumbContainer) {
              photoThumbContainer.innerHTML = '';
              record.photos.forEach((ph, pIdx) => {
                const thumb = document.createElement('div');
                thumb.style.position = 'relative';
                thumb.innerHTML = `
                  <img src="${ph}" class="photo-thumbnail">
                  <div class="delete-photo" onclick="event.stopPropagation(); deleteMockPhoto('${catName}', '${comp.name}', ${pIdx})" 
                    style="position:absolute; top:-6px; right:-6px; background-color:var(--danger); width:18px; height:18px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:12px; cursor:pointer;">×</div>
                `;
                photoThumbContainer.appendChild(thumb);
              });
            }
          }, 0);
        }
        
        body.appendChild(card);
      });
      
      container.appendChild(accordion);
    });
  }
}

// Set status and expand details panel if Minor or Major issue detected
window.setCheckpointRedesignStatus = function(catName, compName, status) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  if (!record) return;
  
  record.status = status;
  
  if (status === 'good') {
    record.selectedIssues = [];
    record.photos = [];
    state.expandedDrawers[key] = false;
  } else {
    state.expandedDrawers[key] = true;
    
    // Auto-check first issue dynamically if none are selected
    const components = getComponentsForRoom(state.currentRoom, catName);
    const compData = components.find(x => x.name === compName);
    if (compData && compData.issues && compData.issues.length > 0 && record.selectedIssues.length === 0) {
      record.selectedIssues.push(compData.issues[0]);
    }
  }
  
  renderActiveRoom();
  renderSidebar();
};

// Realtime note input saving
window.saveCheckpointNote = function(catName, compName, value) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  if (record) {
    record.notes = value;
  }
};

// Express Mode Helper: Tap marks as Healthy (Good)
window.setCheckpointStatusExpressOK = function(catName, compName) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  if (!record) return;
  
  record.status = 'good';
  record.selectedIssues = [];
  record.photos = [];
  
  renderActiveRoom();
  renderSidebar();
};

// Express Mode Bottom Sheet Controls
window.openExpressIssueSheet = function(catName, compName) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  const issues = getIssuesForComponent(catName, compName);
  
  let sheet = document.getElementById('express-bottom-sheet-container');
  if (!sheet) {
    sheet = document.createElement('div');
    sheet.id = 'express-bottom-sheet-container';
    document.body.appendChild(sheet);
  }
  
  sheet.innerHTML = `
    <div class="express-sheet-overlay-backdrop" onclick="closeExpressIssueSheet()"></div>
    <div class="express-overlay-sheet">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <div style="text-align: left;">
          <h3 style="margin:0; font-family:var(--font-heading); font-size:1.1rem; color:var(--bg-sidebar);">${getComponentEmoji(compName)} ${compName}</h3>
          <span style="font-size:0.75rem; color:var(--text-muted); text-transform:uppercase;">${getWalkthroughStageForComponent(catName, compName)}</span>
        </div>
        <button onclick="closeExpressIssueSheet()" style="background:transparent; border:none; font-size:1.5rem; cursor:pointer; color:var(--text-secondary);">×</button>
      </div>
      
      <div style="display:flex; flex-direction:column; gap:16px; overflow-y:auto; max-height:60vh;">
        <div>
          <label style="font-size:0.75rem; font-weight:700; color:var(--text-secondary); text-transform:uppercase; display:block; margin-bottom:8px; text-align: left;">Assessment Status:</label>
          <div class="inline-assessment-row">
            <button class="btn-assess btn-assess-good ${record.status === 'good' ? 'active' : ''}" onclick="setExpressSheetStatus('${catName}', '${compName}', 'good')">
              ✓ OK
            </button>
            <button class="btn-assess btn-assess-minor ${record.status === 'minor' ? 'active' : ''}" onclick="setExpressSheetStatus('${catName}', '${compName}', 'minor')">
              ⚠ Minor Issue
            </button>
            <button class="btn-assess btn-assess-major ${record.status === 'major' ? 'active' : ''}" onclick="setExpressSheetStatus('${catName}', '${compName}', 'major')">
              ❌ Major Issue
            </button>
          </div>
        </div>
        
        <div id="express-defect-details-container" style="display: ${record.status === 'minor' || record.status === 'major' ? 'block' : 'none'};">
          <label style="font-size:0.75rem; font-weight:700; color:var(--text-secondary); text-transform:uppercase; display:block; margin-bottom:8px; text-align: left;">Observed Defects:</label>
          <div style="display:flex; flex-direction:column; gap:6px; margin-bottom:16px;">
            ${issues.map(issText => {
              const isChecked = record.selectedIssues.includes(issText);
              return `
                <label class="issue-checkbox-label ${isChecked ? 'selected' : ''}" style="padding: 8px 12px; border-radius:6px; border:1px solid var(--bg-border); display:flex; align-items:center; gap:10px; cursor:pointer; margin-bottom:4px;">
                  <input type="checkbox" style="width:16px; height:16px; accent-color:var(--danger);" ${isChecked ? 'checked' : ''} 
                    onchange="toggleExpressSheetIssue('${catName}', '${compName}', '${issText}', this.checked)">
                  <span style="font-size:0.82rem; color:var(--text-primary); font-weight:600; text-align: left;">${issText}</span>
                </label>
              `;
            }).join('')}
          </div>
          
          <label style="font-size:0.75rem; font-weight:700; color:var(--text-secondary); text-transform:uppercase; display:block; margin-bottom:8px; text-align: left;">Visual Proof Photo:</label>
          <div style="margin-bottom:16px;">
            <div class="photo-uploader" onclick="mockExpressPhotoUpload('${catName}', '${compName}')">
              <span>📷 Capture Defect Proof Photo</span>
            </div>
            <div class="photo-thumbnail-container" id="express-photos-container" style="margin-top:10px; display:flex; gap:8px;"></div>
          </div>
          
          <label style="font-size:0.75rem; font-weight:700; color:var(--text-secondary); text-transform:uppercase; display:block; margin-bottom:8px; text-align: left;">Notes:</label>
          <textarea style="width:100%; height:60px; padding:8px; border-radius:6px; border:1px solid var(--bg-border); font-size:0.8rem; font-family:var(--font-body); resize:vertical;" 
            placeholder="Add descriptive notes..." oninput="saveCheckpointNote('${catName}', '${compName}', this.value)">${record.notes || ''}</textarea>
        </div>
        
        <button class="btn btn-primary" onclick="closeExpressIssueSheet()" style="width:100%; padding:12px; font-weight:bold; border-radius:12px; background-color:var(--bg-sidebar); border-color:var(--bg-sidebar); color:#ffffff; font-size:0.9rem; margin-top:10px; height: auto;">
          ✔️ Save & Close Sheet
        </button>
      </div>
    </div>
  `;
  
  renderExpressSheetPhotos(catName, compName);
};

window.closeExpressIssueSheet = function() {
  const sheet = document.getElementById('express-bottom-sheet-container');
  if (sheet) {
    sheet.innerHTML = '';
  }
  renderActiveRoom();
  renderSidebar();
};

window.setExpressSheetStatus = function(catName, compName, status) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  if (!record) return;
  
  record.status = status;
  
  const container = document.getElementById('express-defect-details-container');
  if (container) {
    if (status === 'good') {
      container.style.display = 'none';
      record.selectedIssues = [];
      record.photos = [];
    } else {
      container.style.display = 'block';
      
      const components = getComponentsForRoom(state.currentRoom, catName);
      const compData = components.find(x => x.name === compName);
      if (compData && compData.issues && compData.issues.length > 0 && record.selectedIssues.length === 0) {
        record.selectedIssues.push(compData.issues[0]);
      }
    }
  }
  
  // Refresh bottom sheet buttons
  const buttons = document.querySelectorAll('.express-overlay-sheet .btn-assess');
  buttons.forEach((btn, idx) => {
    btn.classList.remove('active');
    if (idx === 0 && status === 'good') btn.classList.add('active');
    if (idx === 1 && status === 'minor') btn.classList.add('active');
    if (idx === 2 && status === 'major') btn.classList.add('active');
  });
  
  // Re-render sheet checklist and photos if status changed
  if (status !== 'good') {
    const issues = getIssuesForComponent(catName, compName);
    const listContainer = document.querySelector('.express-overlay-sheet .issue-items-list');
    if (listContainer) {
      listContainer.innerHTML = issues.map(issText => {
        const isChecked = record.selectedIssues.includes(issText);
        return `
          <label class="issue-checkbox-label ${isChecked ? 'selected' : ''}" style="padding: 8px 12px; border-radius:6px; border:1px solid var(--bg-border); display:flex; align-items:center; gap:10px; cursor:pointer; margin-bottom:4px;">
            <input type="checkbox" style="width:16px; height:16px; accent-color:var(--danger);" ${isChecked ? 'checked' : ''} 
              onchange="toggleExpressSheetIssue('${catName}', '${compName}', '${issText}', this.checked)">
            <span style="font-size:0.82rem; color:var(--text-primary); font-weight:600; text-align: left;">${issText}</span>
          </label>
        `;
      }).join('');
    }
    renderExpressSheetPhotos(catName, compName);
  }
};

window.toggleExpressSheetIssue = function(catName, compName, issueText, isChecked) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  if (!record) return;
  
  if (isChecked) {
    if (!record.selectedIssues.includes(issueText)) {
      record.selectedIssues.push(issueText);
    }
  } else {
    record.selectedIssues = record.selectedIssues.filter(x => x !== issueText);
  }
  
  // Refresh issue checkboxes styling
  const labels = document.querySelectorAll('.express-overlay-sheet .issue-checkbox-label');
  labels.forEach(lbl => {
    const txt = lbl.querySelector('span').innerText;
    if (txt === issueText) {
      lbl.classList.toggle('selected', isChecked);
    }
  });
};

window.mockExpressPhotoUpload = function(catName, compName) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  if (!record) return;
  
  const canvas = document.createElement('canvas');
  canvas.width = 300;
  canvas.height = 200;
  const ctx = canvas.getContext('2d');
  
  const gradient = ctx.createLinearGradient(0, 0, 300, 200);
  gradient.addColorStop(0, '#0e1b3e');
  gradient.addColorStop(1, '#dc2626');
  ctx.fillStyle = gradient;
  ctx.fillRect(0, 0, 300, 200);
  
  ctx.strokeStyle = '#ffffff';
  ctx.lineWidth = 1;
  ctx.strokeRect(60, 50, 180, 100);
  
  ctx.fillStyle = '#ffffff';
  ctx.font = '9px sans-serif';
  ctx.fillText(`DEFECT CAPTURED`, 70, 70);
  ctx.fillText(`EXPRESS STREAM OK`, 70, 130);
  
  const photoUrl = canvas.toDataURL('image/jpeg');
  record.photos.push(photoUrl);
  
  renderExpressSheetPhotos(catName, compName);
};

window.deleteExpressSheetPhoto = function(catName, compName, pIdx) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  if (record && record.photos) {
    record.photos.splice(pIdx, 1);
    renderExpressSheetPhotos(catName, compName);
  }
};

window.renderExpressSheetPhotos = function(catName, compName) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  const container = document.getElementById('express-photos-container');
  if (container && record && record.photos) {
    container.innerHTML = '';
    record.photos.forEach((ph, pIdx) => {
      const thumb = document.createElement('div');
      thumb.style.position = 'relative';
      thumb.innerHTML = `
        <img src="${ph}" class="photo-thumbnail">
        <div class="delete-photo" onclick="event.stopPropagation(); deleteExpressSheetPhoto('${catName}', '${compName}', ${pIdx})" 
          style="position:absolute; top:-6px; right:-6px; background-color:var(--danger); width:18px; height:18px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:12px; cursor:pointer;">×</div>
      `;
      container.appendChild(thumb);
    });
  }
};

// Stepper navigation for Field Mode Complete screen
window.goToNextRoom = function() {
  const rooms = getUniqueRooms();
  const currentRoomIdx = rooms.indexOf(state.currentRoom);
  const nextIdx = currentRoomIdx + 1;
  if (nextIdx < rooms.length) {
    state.currentRoom = rooms[nextIdx];
    alert(`Moving walkthrough to ${state.currentRoom}...`);
  } else {
    alert("You have inspected all property rooms! Click 'Save PDF Report' to download the final audit sheet.");
  }
  renderSidebar();
  renderActiveRoom();
};

// Collapsible category accordions
window.toggleCategoryAccordion = function(catName) {
  state.expandedCategories[catName] = !state.expandedCategories[catName];
  renderAccordions();
};

// Toggle items drawers in state
window.toggleCheckpointControls = function(key) {
  state.expandedDrawers[key] = !state.expandedDrawers[key];
  renderAccordions();
};

// Collapse active checkpoint drawer
window.closeCheckpointDrawer = function(key) {
  state.expandedDrawers[key] = false;
  renderActiveRoom();
};

// Set checkpoint status (Locks issues and collapses on good, expands and keeps open on issue)
window.setCheckpointStatus = function(catName, compName, status) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  if (!record) return;
  
  record.status = status;
  
  // Ensure the category accordion remains expanded!
  state.expandedCategories[catName] = true;
  
  if (status !== 'issue') {
    record.selectedIssues = [];
    record.photos = [];
    state.expandedDrawers[key] = false; // Collapse on good condition
  } else {
    state.expandedDrawers[key] = true; // FORCE expansion open on Issue Found!
    
    // Auto-check first issue dynamically
    const components = getComponentsForRoom(state.currentRoom, catName);
    const compData = components.find(x => x.name === compName);
    if (compData && compData.issues && compData.issues.length > 0 && record.selectedIssues.length === 0) {
      record.selectedIssues.push(compData.issues[0]);
    }
  }
  
  renderActiveRoom();
  renderSidebar();
};

// Toggle checklist issues
window.toggleSpecificIssueSelection = function(catName, compName, issueText, isChecked) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  if (!record) return;
  
  if (isChecked) {
    if (!record.selectedIssues.includes(issueText)) {
      record.selectedIssues.push(issueText);
    }
  } else {
    record.selectedIssues = record.selectedIssues.filter(x => x !== issueText);
  }
  
  renderActiveRoom();
  renderSidebar();
};

// Mock capture photo
window.mockPhotoUpload = function(catName, compName) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  if (!record) return;
  
  const canvas = document.createElement('canvas');
  canvas.width = 300;
  canvas.height = 200;
  const ctx = canvas.getContext('2d');
  
  const gradient = ctx.createLinearGradient(0, 0, 300, 200);
  gradient.addColorStop(0, '#0e1b3e');
  gradient.addColorStop(1, '#1d4ed8');
  ctx.fillStyle = gradient;
  ctx.fillRect(0, 0, 300, 200);
  
  ctx.strokeStyle = '#ffffff';
  ctx.lineWidth = 1;
  ctx.strokeRect(60, 50, 180, 100);
  
  ctx.fillStyle = '#ffffff';
  ctx.font = '9px sans-serif';
  ctx.fillText(`THERMAL DIAGNOSTIC OK`, 70, 70);
  ctx.fillText(`FIELD CAPTURE SUCCESS`, 70, 130);
  
  const photoUrl = canvas.toDataURL('image/jpeg');
  record.photos.push(photoUrl);
  
  renderActiveRoom();
};

window.deleteMockPhoto = function(catName, compName, pIdx) {
  const key = `${state.currentRoom}::${catName}::${compName}`;
  const record = state.checkpointStates[key];
  if (record && record.photos) {
    record.photos.splice(pIdx, 1);
    renderActiveRoom();
  }
};

// Render Bottom floating Category navigation Wizard bar (Safe checks for null elements)
function renderBottomNavBar() {
  const prevIdx = state.activeCategoryIndex - 1;
  const nextIdx = state.activeCategoryIndex + 1;
  
  const prevCat = prevIdx >= 0 ? categoriesList[prevIdx] : "Property Setup";
  const nextCat = nextIdx < categoriesList.length ? categoriesList[nextIdx] : "Finish & PDF";
  
  const navContainer = document.getElementById('wizard-bottom-navigation');
  if (!navContainer) return; // Prevent crashes if elements don't exist yet!
  
  navContainer.innerHTML = '';
  
  // Prev Button
  const prevBtn = document.createElement('button');
  prevBtn.className = 'prev-btn-pill';
  prevBtn.innerHTML = `
    <span>◀ PREV</span>
    <strong>${prevCat}</strong>
  `;
  prevBtn.addEventListener('click', () => {
    if (prevIdx >= 0) {
      state.activeCategoryIndex = prevIdx;
      expandAndScrollToCategory(categoriesList[prevIdx]);
    } else {
      toggleConfigModal(true);
    }
  });
  
  // Next Button
  const nextBtn = document.createElement('button');
  nextBtn.className = 'next-btn-pill';
  nextBtn.innerHTML = `
    <span>NEXT ▶</span>
    <strong>${nextCat}</strong>
  `;
  nextBtn.addEventListener('click', () => {
    if (nextIdx < categoriesList.length) {
      state.activeCategoryIndex = nextIdx;
      expandAndScrollToCategory(categoriesList[nextIdx]);
    } else {
      downloadPdfReport();
    }
  });
  
  // Center Pagination dots
  const centerPagination = document.createElement('div');
  centerPagination.className = 'pagination-center';
  
  centerPagination.innerHTML = `
    <div class="page-indicator-text">${state.activeCategoryIndex + 1} <span>/ ${categoriesList.length}</span></div>
  `;
  
  navContainer.appendChild(prevBtn);
  navContainer.appendChild(centerPagination);
  navContainer.appendChild(nextBtn);
}

// Expand selected category and smooth scroll viewport to it
function expandAndScrollToCategory(catName) {
  categoriesList.forEach(c => {
    state.expandedCategories[c] = (c === catName);
  });
  
  renderActiveRoom();
  
  setTimeout(() => {
    const el = document.getElementById(`accordion-${catName.replace(/[^a-zA-Z0-9]/g, '')}`);
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }, 100);
}

// 45-Minute Timer
function setupTimer() {
  state.timerSeconds = 45 * 60;
  const timerEl = document.getElementById('timer-clock');
  
  state.timerInterval = setInterval(() => {
    state.timerSeconds--;
    if (state.timerSeconds <= 0) {
      clearInterval(state.timerInterval);
      timerEl.innerHTML = "00:00";
      timerEl.parentElement.classList.add('timer-warning');
      alert("Inspection time budget of 45 minutes completed!");
      return;
    }
    
    const mins = Math.floor(state.timerSeconds / 60);
    const secs = state.timerSeconds % 60;
    
    if (timerEl) {
      timerEl.innerHTML = `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
    }
  }, 1000);
}

// Property Configuration
window.toggleConfigModal = function(show) {
  const modal = document.getElementById('config-modal');
  if (show) {
    document.getElementById('chk_ac').checked = !!state.config.has_ac;
    document.getElementById('chk_geyser').checked = !!state.config.has_geyser;
    document.getElementById('chk_window').checked = !!state.config.has_window;
    document.getElementById('chk_wardrobe').checked = !!state.config.has_wardrobe;
    document.getElementById('chk_purifier').checked = !!state.config.has_purifier;
    document.getElementById('chk_softener').checked = !!state.config.has_softener;
    document.getElementById('chk_pump').checked = !!state.config.has_pump;
    document.getElementById('chk_fridge').checked = !!state.config.has_fridge;
    document.getElementById('chk_chimney').checked = !!state.config.has_chimney;
    document.getElementById('chk_exhaust').checked = !!state.config.has_exhaust;
    document.getElementById('chk_db').checked = !!state.config.has_db;
    document.getElementById('chk_outdoor_lights').checked = !!state.config.has_outdoor_lights;
    document.getElementById('chk_outdoor_taps').checked = !!state.config.has_outdoor_taps;
    document.getElementById('chk_balcony_enclosure').checked = !!state.config.has_balcony_enclosure;
    document.getElementById('chk_pooja_wood').checked = !!state.config.has_poja_wood;
    
    modal.classList.add('active');
  } else {
    modal.classList.remove('active');
  }
};

// Save Property Config
window.savePropertyConfig = function() {
  state.config.has_ac = document.getElementById('chk_ac').checked;
  state.config.has_geyser = document.getElementById('chk_geyser').checked;
  state.config.has_window = document.getElementById('chk_window').checked;
  state.config.has_wardrobe = document.getElementById('chk_wardrobe').checked;
  state.config.has_purifier = document.getElementById('chk_purifier').checked;
  state.config.has_softener = document.getElementById('chk_softener').checked;
  state.config.has_pump = document.getElementById('chk_pump').checked;
  state.config.has_fridge = document.getElementById('chk_fridge').checked;
  state.config.has_chimney = document.getElementById('chk_chimney').checked;
  state.config.has_exhaust = document.getElementById('chk_exhaust').checked;
  state.config.has_db = document.getElementById('chk_db').checked;
  state.config.has_outdoor_lights = document.getElementById('chk_outdoor_lights').checked;
  state.config.has_outdoor_taps = document.getElementById('chk_outdoor_taps').checked;
  state.config.has_balcony_enclosure = document.getElementById('chk_balcony_enclosure').checked;
  state.config.has_poja_wood = document.getElementById('chk_pooja_wood').checked;
  
  toggleConfigModal(false);
  renderSidebar();
  renderActiveRoom();
  alert("Property configuration updated! Checklist rebuilt.");
};

// Print/PDF Generation report
window.downloadPdfReport = function() {
  // Validate checklist completeness (every active component must be marked Good or Issue)
  const pendingData = getPendingInspectionItems();
  if (pendingData.totalPendingCount > 0) {
    showIncompleteWarningModal(pendingData);
    return;
  }
  
  const printContent = document.getElementById('print-report-container');
  printContent.innerHTML = '';
  
  const header = document.createElement('div');
  header.className = 'print-header';
  
  const now = new Date();
  const dateStr = now.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' });
  
  header.innerHTML = `
    <div class="print-logo">🏢 Handy sQuad</div>
    <div class="print-title">Property Health Check Inspection Report</div>
    <div class="print-meta-grid">
      <div class="print-meta-item"><strong>Asset Unit ID:</strong> Skyview-Unit-502</div>
      <div class="print-meta-item"><strong>Date Executed:</strong> ${dateStr}</div>
      <div class="print-meta-item"><strong>Property Type:</strong> 3BHK Standard Apartment</div>
      <div class="print-meta-item"><strong>Inspector Name:</strong> Supervisor S. Kumar</div>
    </div>
  `;
  printContent.appendChild(header);
  
  const summaryTitle = document.createElement('h2');
  summaryTitle.className = 'print-section-title';
  summaryTitle.innerText = "Executive Summary";
  printContent.appendChild(summaryTitle);
  
  let totalChecked = 0;
  let totalPending = 0;
  let totalGood = 0;
  let issueItems = [];
  
  const rooms = getUniqueRooms();
  
  rooms.forEach(roomName => {
    categoriesList.forEach(catName => {
      const components = getComponentsForRoom(roomName, catName);
      components.forEach(comp => {
        const key = `${roomName}::${catName}::${comp.name}`;
        const rec = state.checkpointStates[key] || { status: 'pending', selectedIssues: [], photos: [] };
        if (rec.status === 'pending') {
          totalPending++;
        } else {
          totalChecked++;
          if (rec.status === 'good') {
            totalGood++;
          } else if (rec.status === 'issue' || rec.status === 'minor' || rec.status === 'major') {
            const severity = rec.status === 'major' ? 'HIGH' : 'MEDIUM';
            const urgency = rec.status === 'major' ? 'Immediate' : 'Soon';
            
            rec.selectedIssues.forEach(issText => {
              issueItems.push({
                room: roomName,
                category: catName,
                item: comp.name,
                label: issText,
                cause: "Material fatigue / loose contact / alignment wear",
                recommendation: `Repair or adjust ${comp.name} components: address ${issText}`,
                impact: "Functional degradation; structural safety risk",
                severity: severity,
                urgency: urgency
              });
            });
          }
        }
      });
    });
  });
  
  const statsTable = document.createElement('table');
  statsTable.className = 'print-table';
  statsTable.innerHTML = `
    <tr>
      <th style="width: 50%;">Metric</th>
      <th style="width: 50%;">Value</th>
    </tr>
    <tr>
      <td>Total Inspected Components</td>
      <td><strong>${totalChecked} Components</strong></td>
    </tr>
    <tr>
      <td>Uninspected / Pending Checklist Items</td>
      <td><strong>${totalPending} Items</strong></td>
    </tr>
    <tr>
      <td>Healthy Components (No Issue)</td>
      <td><span class="badge-print badge-print-good">${totalGood} Healthy</span></td>
    </tr>
    <tr>
      <td>Flagged Defective Components</td>
      <td><span class="badge-print badge-print-issue">${issueItems.length} Issues Flagged</span></td>
    </tr>
    <tr>
      <td>Overall Inspection Grade</td>
      <td><strong>${issueItems.length > 0 ? 'ATTENTION REQUIRED (REPAIRS MANDATED)' : totalPending > 0 ? 'INCOMPLETE WORKBENCH' : 'COMPLETELY CLEAR'}</strong></td>
    </tr>
  `;
  printContent.appendChild(statsTable);
  
  const findingsTitle = document.createElement('h2');
  findingsTitle.className = 'print-section-title';
  findingsTitle.innerText = "Detailed Findings (Auto-Filled Diagnostics)";
  printContent.appendChild(findingsTitle);
  
  if (issueItems.length === 0) {
    const p = document.createElement('p');
    p.style.fontSize = '10pt';
    p.style.color = '#374151';
    p.innerText = "All inspected components were evaluated as completely healthy. No repairs or specialist operations are required.";
    printContent.appendChild(p);
  } else {
    const findingsTable = document.createElement('table');
    findingsTable.className = 'print-table';
    
    let ths = `
      <tr>
        <th style="width: 18%;">Location & Component</th>
        <th style="width: 25%;">Flagged Issue & Impact</th>
        <th style="width: 20%;">Root Cause</th>
        <th style="width: 25%;">Recommended Action</th>
        <th style="width: 12%;">Risk & Urgency</th>
      </tr>
    `;
    
    issueItems.forEach(iss => {
      ths += `
        <tr>
          <td><strong>${iss.room}</strong><br><span style="font-size:8pt; color:#2563eb; font-weight:600;">${iss.item}</span><br><span style="font-size:7.5pt; color:#6b7280;">${iss.category}</span></td>
          <td><strong>${iss.label}</strong><br><span style="font-size:8pt; color:#4b5563; display:block; margin-top:4px;"><strong>Impact:</strong> ${iss.impact}</span></td>
          <td style="font-size:8.5pt; color:#374151;">${iss.cause}</td>
          <td><strong style="color:#1e3a8a;">${iss.recommendation}</strong></td>
          <td><span class="badge-print ${iss.severity === 'SAFETY' || iss.severity === 'HIGH' ? 'badge-print-critical' : 'badge-print-issue'}">${iss.severity}</span><br><span style="font-size:7.5pt; color:#4b5563; display:block; margin-top:4px;">⏱️ ${iss.urgency}</span></td>
        </tr>
      `;
    });
    
    findingsTable.innerHTML = ths;
    printContent.appendChild(findingsTable);
  }
  
  const footer = document.createElement('div');
  footer.style.marginTop = '60px';
  footer.style.display = 'grid';
  footer.style.gridTemplateColumns = '1fr 1fr';
  footer.style.gap = '80px';
  footer.style.pageBreakInside = 'avoid';
  
  footer.innerHTML = `
    <div>
      <div style="border-top:1px solid #333; margin-top:20px; padding-top:8px; font-size:10pt; text-align:center;">
        Supervisor Signature (S. Kumar)
      </div>
    </div>
    <div>
      <div style="border-top:1px solid #333; margin-top:20px; padding-top:8px; font-size:10pt; text-align:center;">
        Customer Sign-Off (Unit Owner)
      </div>
    </div>
  `;
  printContent.appendChild(footer);
  
  window.print();
};

// Scan checklist states to find components that have not been marked Good or Issue
function getPendingInspectionItems() {
  const pendingByRoom = {};
  let totalPendingCount = 0;
  
  const rooms = getUniqueRooms();
  rooms.forEach(roomName => {
    categoriesList.forEach(catName => {
      const components = getComponentsForRoom(roomName, catName);
      components.forEach(comp => {
        const key = `${roomName}::${catName}::${comp.name}`;
        const rec = state.checkpointStates[key] || { status: 'pending' };
        if (rec.status === 'pending') {
          if (!pendingByRoom[roomName]) {
            pendingByRoom[roomName] = [];
          }
          pendingByRoom[roomName].push({
            category: catName,
            component: comp.name
          });
          totalPendingCount++;
        }
      });
    });
  });
  
  return {
    pendingByRoom,
    totalPendingCount
  };
}

// Display high-fidelity overlay modal detailing incomplete room checklist items
function showIncompleteWarningModal(pendingData) {
  // Remove existing validation modal if any
  const existing = document.getElementById('validation-modal');
  if (existing) existing.remove();
  
  const modal = document.createElement('div');
  modal.id = 'validation-modal';
  modal.className = 'modal-overlay active';
  modal.style.zIndex = '1000';
  modal.style.display = 'flex';
  modal.style.alignItems = 'center';
  modal.style.justifyContent = 'center';
  
  let roomsHtml = '';
  Object.keys(pendingData.pendingByRoom).forEach(roomName => {
    const items = pendingData.pendingByRoom[roomName];
    roomsHtml += `
      <div class="pending-room-group" style="margin-bottom: 14px; border: 1px solid var(--bg-border); border-radius: var(--border-radius-sm); padding: 12px; background: hsl(214, 25%, 98%);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; border-bottom:1px solid var(--bg-border); padding-bottom:6px;">
          <strong style="color:var(--bg-sidebar); font-size:0.92rem;">${getRoomIcon(roomName)} ${roomName}</strong>
          <span style="font-size:0.75rem; font-weight:700; background:rgba(220, 38, 38, 0.1); color:var(--danger); padding: 2px 8px; border-radius:10px;">${items.length} left</span>
        </div>
        <div style="display:flex; flex-direction:column; gap:4px; max-height:100px; overflow-y:auto; margin-bottom:10px;">
          ${items.map(it => `
            <div style="font-size:0.8rem; color:var(--text-secondary); display:flex; justify-content:space-between;">
              <span>• ${it.component}</span>
              <span style="font-size:0.7rem; color:var(--text-muted); font-style:italic;">${it.category}</span>
            </div>
          `).join('')}
        </div>
        <button class="btn btn-secondary" onclick="navigateToRoomFromWarning('${roomName}')" 
          style="width: 100%; padding: 6px; font-size: 0.75rem; border-radius: 6px; height: auto; font-weight: bold; border-color:var(--blue-active); color:var(--blue-active);">
          🔍 Inspect ${roomName} Now
        </button>
      </div>
    `;
  });
  
  modal.innerHTML = `
    <div class="modal-card" style="max-width: 460px; max-height: 85vh; display: flex; flex-direction: column; box-shadow: var(--shadow-lg);">
      <div class="modal-header" style="border-bottom: 1px solid var(--bg-border); padding-bottom: 12px; display:flex; justify-content:space-between; align-items:center; width:100%;">
        <h3 class="modal-title" style="color:var(--danger); display:flex; align-items:center; gap:8px; margin:0; font-size:1.15rem;">⚠️ Checklist Incomplete</h3>
        <button class="modal-close" onclick="closeValidationModal()" style="border:none; background:transparent; font-size:1.5rem; cursor:pointer;">×</button>
      </div>
      <div class="modal-body" style="overflow-y: auto; flex: 1; padding: 16px 0;">
        <p style="font-size:0.85rem; color:var(--text-secondary); line-height: 1.45; margin-bottom: 16px;">
          You must evaluate either <strong>Good Condition</strong> or <strong>Issue Found</strong> for every active component. There are currently <strong>${pendingData.totalPendingCount} items pending</strong> across these rooms:
        </p>
        <div style="max-height: 340px; overflow-y: auto; padding-right: 4px;">
          ${roomsHtml}
        </div>
      </div>
      <div class="modal-footer" style="border-top: 1px solid var(--bg-border); padding-top: 12px; display:flex; justify-content:flex-end;">
        <button class="btn btn-primary" onclick="closeValidationModal()" style="background-color: var(--bg-sidebar); border-color:var(--bg-sidebar); font-weight: bold; width: 100%;">
          Close & Continue Inspecting
        </button>
      </div>
    </div>
  `;
  
  document.body.appendChild(modal);
}

window.navigateToRoomFromWarning = function(roomName) {
  state.currentRoom = roomName;
  closeValidationModal();
  renderSidebar();
  renderActiveRoom();
};

window.closeValidationModal = function() {
  const modal = document.getElementById('validation-modal');
  if (modal) modal.remove();
};

function setupEventListeners() {
  document.getElementById('btn-setup-trigger').addEventListener('click', () => {
    toggleConfigModal(true);
  });
  document.getElementById('btn-config-setup').addEventListener('click', () => {
    toggleConfigModal(true);
  });
  
  // Mobile Hamburger Toggle
  const burger = document.getElementById('mobile-menu-toggle');
  const overlay = document.getElementById('mobile-sidebar-overlay');
  const container = document.querySelector('.app-container');
  
  if (burger) {
    burger.addEventListener('click', () => {
      container.classList.toggle('sidebar-open');
    });
  }
  
  if (overlay) {
    overlay.addEventListener('click', () => {
      container.classList.remove('sidebar-open');
    });
  }

  // Redesign: Connect AI simulation buttons
  const simBtn = document.getElementById('btn-ai-simulate');
  if (simBtn) {
    simBtn.addEventListener('click', () => {
      simulateAiVisionSweep();
    });
  }

  const confirmBtn = document.getElementById('btn-ai-confirm');
  if (confirmBtn) {
    confirmBtn.addEventListener('click', () => {
      confirmAiDefect();
    });
  }

  const dismissBtn = document.getElementById('btn-ai-dismiss');
  if (dismissBtn) {
    dismissBtn.addEventListener('click', () => {
      dismissAiDefect();
    });
  }
}

// Global/window variables to track camera stream and timeouts
window.aiStream = null;
window.aiAnimationInterval = null;
window.aiSweepTimeouts = [];

window.logToAiConsole = function(type, msg) {
  const consoleEl = document.getElementById('ai-terminal-console');
  if (!consoleEl) return;
  const timeStr = new Date().toTimeString().split(' ')[0] + '.' + String(new Date().getMilliseconds()).padStart(3, '0');
  const lineEl = document.createElement('div');
  lineEl.className = `terminal-line text-${type}`;
  lineEl.innerText = `[${timeStr}] [${type.toUpperCase()}] ${msg}`;
  consoleEl.appendChild(lineEl);
  // Auto-scroll to bottom
  consoleEl.scrollTop = consoleEl.scrollHeight;
};

window.stopWebcamStream = function() {
  if (window.aiStream) {
    const tracks = window.aiStream.getTracks();
    tracks.forEach(track => track.stop());
    window.aiStream = null;
    window.logToAiConsole('system', 'Camera stream stopped and tracks released.');
  }
  const videoEl = document.getElementById('ai-video-stream');
  if (videoEl) {
    videoEl.srcObject = null;
    videoEl.style.display = 'none';
  }
};

window.clearAiSweepIntervals = function() {
  if (window.aiAnimationInterval) {
    clearInterval(window.aiAnimationInterval);
    window.aiAnimationInterval = null;
  }
  if (window.aiSweepTimeouts) {
    window.aiSweepTimeouts.forEach(t => clearTimeout(t));
    window.aiSweepTimeouts = [];
  }
};

window.animateOverlayBox = function(targetX, targetY, targetW, targetH, durationMs) {
  const overlayBox = document.getElementById('ai-overlay-box');
  if (!overlayBox) return;
  
  overlayBox.style.display = 'block';
  const startTime = performance.now();
  
  const startX = parseFloat(overlayBox.style.left) || 20;
  const startY = parseFloat(overlayBox.style.top) || 20;
  const startW = parseFloat(overlayBox.style.width) || 60;
  const startH = parseFloat(overlayBox.style.height) || 60;
  
  if (window.aiAnimationInterval) {
    clearInterval(window.aiAnimationInterval);
  }
  
  window.aiAnimationInterval = setInterval(() => {
    const elapsed = performance.now() - startTime;
    const progress = Math.min(elapsed / durationMs, 1);
    
    // easeOutQuad
    const ease = progress * (2 - progress);
    
    const currX = startX + (targetX - startX) * ease;
    const currY = startY + (targetY - startY) * ease;
    const currW = startW + (targetW - startW) * ease;
    const currH = startH + (targetH - startH) * ease;
    
    overlayBox.style.left = `${currX}px`;
    overlayBox.style.top = `${currY}px`;
    overlayBox.style.width = `${currW}px`;
    overlayBox.style.height = `${currH}px`;
    
    if (progress >= 1) {
      clearInterval(window.aiAnimationInterval);
    }
  }, 16);
};

// AI Vision Sweep Walkpath Simulation
window.simulateAiVisionSweep = function() {
  window.stopWebcamStream();
  window.clearAiSweepIntervals();
  
  const feedText = document.getElementById('ai-feed-text');
  const overlayBox = document.getElementById('ai-overlay-box');
  const overlayLabel = document.getElementById('ai-overlay-label');
  const alertBox = document.getElementById('ai-alert-box');
  const alertText = document.getElementById('ai-alert-text');
  const videoEl = document.getElementById('ai-video-stream');
  
  if (!feedText || !overlayBox || !overlayLabel || !alertBox || !alertText) return;
  
  alertBox.style.display = 'none';
  overlayBox.style.display = 'none';
  
  // Clear developer console
  const consoleEl = document.getElementById('ai-terminal-console');
  if (consoleEl) consoleEl.innerHTML = '';
  
  window.logToAiConsole('system', 'Requesting hardware camera access (facingMode: environment)...');
  feedText.innerText = "🎥 Initializing camera sweep walkthrough...";
  
  // Try to stream environment camera
  navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } })
    .then(stream => {
      startScannerStream(stream);
    })
    .catch(err => {
      window.logToAiConsole('warning', `Rear camera unavailable (${err.name}: ${err.message}). Trying standard camera...`);
      navigator.mediaDevices.getUserMedia({ video: true })
        .then(stream => {
          startScannerStream(stream);
        })
        .catch(err2 => {
          window.logToAiConsole('error', `Camera access blocked (${err2.name}: ${err2.message})`);
          window.logToAiConsole('system', 'Swapping to simulated walkthrough fallback.');
          runFallbackSimulation();
        });
    });
    
  function startScannerStream(stream) {
    window.aiStream = stream;
    if (videoEl) {
      videoEl.style.display = 'block';
      videoEl.setAttribute('playsinline', 'true');
      videoEl.setAttribute('autoplay', 'true');
      videoEl.setAttribute('muted', 'true');
      videoEl.srcObject = stream;
      
      // Explicitly trigger play to bypass mobile browser autoplay blocks
      videoEl.play().then(() => {
        window.logToAiConsole('system', 'Camera playback started successfully.');
      }).catch(playErr => {
        window.logToAiConsole('warning', `Autoplay pending user action: ${playErr.message}`);
      });
    }
    window.logToAiConsole('system', 'Camera stream active. Resolving video dimensions...');
    runDetectionSequence(true);
  }
  
  function runFallbackSimulation() {
    if (videoEl) videoEl.style.display = 'none';
    window.logToAiConsole('system', 'Initiating simulated scan pattern...');
    runDetectionSequence(false);
  }
  
  function runDetectionSequence(hasCamera) {
    // Stage 1: Calibration (0s to 1.5s)
    window.logToAiConsole('system', 'Visual calibration initiated. Running lighting auto-balance...');
    if (hasCamera) {
      window.logToAiConsole('system', 'RGB frame buffer bound. Hardware acceleration: ENABLED.');
    }
    
    // Find active components in the current room and category
    let targetComponent = null;
    let targetCategory = null;
    
    if (state.activeCategory === 'all') {
      for (let cat of categoriesList) {
        const comps = getComponentsForRoom(state.currentRoom, cat);
        if (comps && comps.length > 0) {
          targetComponent = comps[0];
          targetCategory = cat;
          break;
        }
      }
    } else {
      const comps = getComponentsForRoom(state.currentRoom, state.activeCategory);
      if (comps && comps.length > 0) {
        targetComponent = comps[0];
        targetCategory = state.activeCategory;
      }
    }
    
    // Abort if no inspectable assets are in the current room/category checklist
    if (!targetComponent || !targetCategory) {
      window.logToAiConsole('error', 'Scan aborted: No active components found in this category/room layout.');
      feedText.innerText = "❌ No active components in this view configuration.";
      window.stopWebcamStream();
      window.clearAiSweepIntervals();
      return;
    }
    
    const targetAsset = targetComponent.name;
    const issueText = (targetComponent.issues && targetComponent.issues.length > 0) ? 
      targetComponent.issues[0] : "General defect or wear identified";
      
    // Generate semi-random high confidence rate
    const confRate = 88 + Math.floor(Math.random() * 11); // 88% - 98%
    const confidence = `${confRate}%`;
    const defectDesc = `${issueText} [${confidence} confidence]`;
    
    // Store in global window state so confirmAiDefect accesses the exact scanned asset
    window.aiDetectedComponent = targetAsset;
    window.aiDetectedCategory = targetCategory;
    window.aiDetectedIssue = issueText;
    window.aiDetectedConfidence = confidence;
    
    window.logToAiConsole('system', `Target locked: Scanning checklist item '${targetAsset}'...`);
    
    // Animate box floating around finding things
    overlayBox.style.left = '20px';
    overlayBox.style.top = '20px';
    overlayBox.style.width = '60px';
    overlayBox.style.height = '60px';
    overlayBox.style.borderColor = '#3b82f6'; // blue scanning
    overlayBox.style.backgroundColor = 'rgba(59, 130, 246, 0.1)';
    overlayLabel.innerText = "SCANNING...";
    overlayLabel.style.backgroundColor = '#3b82f6';
    
    window.animateOverlayBox(80, 40, 120, 100, 1200);
    
    // Timeout 1: (1.5 seconds) - Detect fixture
    const t1 = setTimeout(() => {
      feedText.innerText = `🔍 Scanning room assets: ${targetAsset} identified.`;
      window.logToAiConsole('ai', `Detected bounding anchor: ${targetAsset} [${confidence}]`);
      window.logToAiConsole('system', `Querying database master templates for standard compliance...`);
      
      overlayBox.style.borderColor = '#10b981'; // green detection
      overlayBox.style.backgroundColor = 'rgba(16, 185, 129, 0.1)';
      overlayLabel.innerText = `${targetAsset} [${confidence}]`;
      overlayLabel.style.backgroundColor = '#10b981';
      
      // Keep animating box to focus closer
      window.animateOverlayBox(60, 30, 150, 110, 1200);
    }, 1500);
    window.aiSweepTimeouts.push(t1);
    
    // Timeout 2: (3.0 seconds) - Defect evaluation
    const t2 = setTimeout(() => {
      feedText.innerText = "⚠️ Defect identified! Evaluating anomaly details...";
      window.logToAiConsole('warning', `Anomaly matched against issue master database!`);
      window.logToAiConsole('ai', `Defect classification: ${defectDesc}`);
      window.logToAiConsole('system', `Action payload generated. Prompting inspector for approval...`);
      
      overlayBox.style.borderColor = '#ef4444'; // red warning
      overlayBox.style.backgroundColor = 'rgba(239, 68, 68, 0.15)';
      overlayLabel.innerText = `⚠️ DEFECT: ${confidence}`;
      overlayLabel.style.backgroundColor = '#ef4444';
      
      alertText.innerText = defectDesc;
      alertBox.style.display = 'block';
    }, 3000);
    window.aiSweepTimeouts.push(t2);
  }
};

window.confirmAiDefect = function() {
  const targetAsset = window.aiDetectedComponent;
  const targetCategory = window.aiDetectedCategory;
  const issueText = window.aiDetectedIssue;
  const confidence = window.aiDetectedConfidence || "94%";
  
  if (!targetAsset || !targetCategory || !issueText) {
    alert("No active AI detection to confirm.");
    return;
  }
  
  let matchKey = `${state.currentRoom}::${targetCategory}::${targetAsset}`;
  let matchCat = targetCategory;
  let matchCompName = targetAsset;
  
  // Verify component exists in checklist state
  if (matchKey && state.checkpointStates[matchKey]) {
    const record = state.checkpointStates[matchKey];
    record.status = 'major';
    
    if (!record.selectedIssues.includes(issueText)) {
      record.selectedIssues.push(issueText);
    }
    
    record.notes = `AI Vision Sweep Defect Auto-Detected: ${issueText}`;
    
    // Canvas capturing video frame or drawing mock diagnostic UI
    const videoEl = document.getElementById('ai-video-stream');
    const overlayBox = document.getElementById('ai-overlay-box');
    
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    
    let useCameraFrame = false;
    let drawX = 60, drawY = 50, drawW = 280, drawH = 180;
    let imgWidth = 400, imgHeight = 300;
    
    if (window.aiStream && videoEl && videoEl.readyState >= 2) {
      imgWidth = videoEl.videoWidth || 640;
      imgHeight = videoEl.videoHeight || 480;
      canvas.width = imgWidth;
      canvas.height = imgHeight;
      
      // Draw camera image
      ctx.drawImage(videoEl, 0, 0, imgWidth, imgHeight);
      useCameraFrame = true;
      
      const feedBox = document.getElementById('ai-feed-box');
      if (feedBox && overlayBox) {
        const feedRect = feedBox.getBoundingClientRect();
        const boxRect = overlayBox.getBoundingClientRect();
        
        const scaleX = imgWidth / feedRect.width;
        const scaleY = imgHeight / feedRect.height;
        
        drawX = (boxRect.left - feedRect.left) * scaleX;
        drawY = (boxRect.top - feedRect.top) * scaleY;
        drawW = boxRect.width * scaleX;
        drawH = boxRect.height * scaleY;
        
        // Red Bounding Box on Snapshot
        ctx.strokeStyle = '#ef4444';
        ctx.lineWidth = Math.max(3, Math.round(imgWidth / 150));
        ctx.strokeRect(drawX, drawY, drawW, drawH);
        
        // Defect banner
        ctx.fillStyle = '#ef4444';
        const labelHeight = Math.max(20, Math.round(imgHeight / 20));
        const labelY = Math.max(0, drawY - labelHeight);
        ctx.fillRect(drawX, labelY, drawW, labelHeight);
        
        ctx.fillStyle = '#ffffff';
        const fontSize = Math.max(10, Math.round(imgHeight / 25));
        ctx.font = `bold ${fontSize}px sans-serif`;
        ctx.fillText(`⚠️ AI ANOMALY: ${targetAsset}`, drawX + 6, labelY + fontSize * 0.8);
      }
    }
    
    if (!useCameraFrame) {
      canvas.width = 400;
      canvas.height = 300;
      
      ctx.fillStyle = '#0f172a';
      ctx.fillRect(0, 0, 400, 300);
      
      ctx.strokeStyle = 'rgba(56, 189, 248, 0.08)';
      ctx.lineWidth = 1;
      for (let x = 0; x < 400; x += 20) {
        ctx.beginPath();
        ctx.moveTo(x, 0);
        ctx.lineTo(x, 300);
        ctx.stroke();
      }
      for (let y = 0; y < 300; y += 20) {
        ctx.beginPath();
        ctx.moveTo(0, y);
        ctx.lineTo(400, y);
        ctx.stroke();
      }
      
      ctx.strokeStyle = 'rgba(56, 189, 248, 0.15)';
      ctx.beginPath();
      ctx.arc(200, 150, 100, 0, 2 * Math.PI);
      ctx.stroke();
      
      ctx.strokeStyle = '#ef4444';
      ctx.lineWidth = 3;
      ctx.strokeRect(60, 50, 280, 180);
      
      ctx.fillStyle = '#ef4444';
      ctx.fillRect(60, 20, 280, 30);
      ctx.fillStyle = '#ffffff';
      ctx.font = 'bold 12px sans-serif';
      ctx.fillText(`⚠️ AI DEFECT: ${targetAsset.toUpperCase()}`, 70, 40);
      
      ctx.fillStyle = '#e2e8f0';
      ctx.font = '11px monospace';
      ctx.fillText(`ROOM: ${state.currentRoom.toUpperCase()}`, 80, 90);
      ctx.fillText(`CATEGORY: ${targetCategory.toUpperCase()}`, 80, 115);
      ctx.fillText(`CLASSIFICATION: ${issueText}`, 80, 140);
      ctx.fillText(`CONFIDENCE RATE: ${confidence}`, 80, 165);
      ctx.fillText(`RESOLUTION STATUS: UNRESOLVED`, 80, 190);
      
      ctx.fillStyle = '#ef4444';
      ctx.font = 'bold 9px monospace';
      ctx.fillText(`[TRACKING SCAN ERROR]`, 80, 215);
    }
    
    const photoUrl = canvas.toDataURL('image/jpeg');
    record.photos.push(photoUrl);
    
    // Save to trainingDataset collector for Active Learning
    let yoloClass = 0;
    const lowerAsset = targetAsset.toLowerCase();
    if (lowerAsset.includes("toilet") || lowerAsset.includes("wc")) yoloClass = 1;
    else if (lowerAsset.includes("chimney")) yoloClass = 2;
    else if (lowerAsset.includes("wardrobe")) yoloClass = 3;
    else if (lowerAsset.includes("tile")) yoloClass = 4;
    else if (lowerAsset.includes("light") || lowerAsset.includes("fixture")) yoloClass = 5;
    
    const cx = (drawX + drawW / 2.0) / imgWidth;
    const cy = (drawY + drawH / 2.0) / imgHeight;
    const bw = drawW / imgWidth;
    const bh = drawH / imgHeight;
    
    state.trainingDataset.push({
      fileName: `img_inspect_${Date.now()}`,
      room: state.currentRoom,
      asset: targetAsset,
      defect: issueText,
      image: photoUrl,
      yolo: {
        classId: yoloClass,
        cx: parseFloat(cx.toFixed(6)),
        cy: parseFloat(cy.toFixed(6)),
        bw: parseFloat(bw.toFixed(6)),
        bh: parseFloat(bh.toFixed(6))
      }
    });
    
    state.activeCategory = matchCat;
    state.expandedCategories[matchCat] = true;
    state.expandedDrawers[matchKey] = true;
    
    window.stopWebcamStream();
    window.clearAiSweepIntervals();
    
    renderActiveRoom();
    renderSidebar();
    
    alert(`AI defect confirmed for ${matchCompName}! Asset status updated to Major Issue.`);
  }
};

window.dismissAiDefect = function() {
  const alertBox = document.getElementById('ai-alert-box');
  const overlayBox = document.getElementById('ai-overlay-box');
  const feedText = document.getElementById('ai-feed-text');
  
  if (alertBox) alertBox.style.display = 'none';
  if (overlayBox) overlayBox.style.display = 'none';
  if (feedText) feedText.innerText = "📹 Live Video Stream Walkthrough Simulation";
  
  window.stopWebcamStream();
  window.clearAiSweepIntervals();
  
  window.logToAiConsole('system', 'AI scan sweep halted/reset.');
};

window.exportAiDataset = function() {
  if (!state.trainingDataset || state.trainingDataset.length === 0) {
    alert("No training dataset frames collected yet! Run some camera sweeps and click 'Confirm & Flag' to log defects first.");
    return;
  }
  
  // Create a Python unpacker script
  const pythonScript = `import os
import json
import base64

# Dataset payload from Handy sQuad Inspections
DATASET = ${JSON.stringify(state.trainingDataset)}

print(f"Unpacking {len(DATASET)} training frames locally...")

# Create target directories
os.makedirs("dataset/images/train", exist_ok=True)
os.makedirs("dataset/labels/train", exist_ok=True)

for i, item in enumerate(DATASET):
    file_name = item.get("fileName", f"frame_{i}")
    img_data_url = item.get("image", "")
    yolo = item.get("yolo", {})
    
    if "data:image/jpeg;base64," in img_data_url:
        base64_str = img_data_url.split(",")[1]
    else:
        print(f"Skip frame {i}: Invalid image data url")
        continue
        
    # Write image
    img_path = f"dataset/images/train/{file_name}.jpg"
    with open(img_path, "wb") as img_file:
        img_file.write(base64.b64decode(base64_str))
        
    # Write YOLO label
    label_path = f"dataset/labels/train/{file_name}.txt"
    with open(label_path, "w") as lbl_file:
        class_id = yolo.get("classId", 0)
        cx = yolo.get("cx", 0.5)
        cy = yolo.get("cy", 0.5)
        bw = yolo.get("bw", 0.5)
        bh = yolo.get("bh", 0.5)
        lbl_file.write(f"{class_id} {cx:.6f} {cy:.6f} {bw:.6f} {bh:.6f}\\n")
        
    print(f" -> Unpacked {file_name}.jpg and corresponding YOLO annotation.")

print("\\nSuccess! Your local dataset is ready under dataset/images/train/ and dataset/labels/train/")
print("Now run 'python train.py' to train your custom AI!")
`;

  // Download the python script
  const blob = new Blob([pythonScript], { type: "text/plain" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "unpack_dataset.py";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  
  alert(`Successfully generated 'unpack_dataset.py' containing ${state.trainingDataset.length} annotated training frames! Save it in your project folder and run 'python unpack_dataset.py' to unpack.`);
};
