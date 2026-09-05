import { useState, useEffect, useCallback, useMemo, Fragment, Component } from "react";
import * as XLSX from "xlsx";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";
import { storage } from "./lib/storage";
import Papa from "papaparse";
import {
  ClipboardList,
  Grape,
  FlaskConical,
  Plus,
  Trash2,
  Download,
  Loader2,
  ChevronDown,
  ChevronUp,
  CheckCircle2,
  RotateCcw,
  ChevronLeft,
  ChevronRight,
  Circle,
  CheckSquare,
  Printer,
  Calendar,
  Calculator,
  Pencil,
  Check,
  X,
  Home,
  Sun,
  CloudRain,
  Wind,
  Droplets,
  Sunrise,
  Sunset,
  Thermometer,
  Package,
  Copy,
  BookmarkPlus,
  Library,
  ArrowUp,
  ArrowDown,
  ArrowUpDown,
  Wine,
  GitCompare,
  Star,
  Archive,
  ArchiveRestore,
  HardDrive,
  UploadCloud,
  Beaker,
  Cylinder,
  Package2,
  Microscope,
  CalendarPlus,
  FileText,
  Sprout,
  Clock,
  DollarSign,
  LogOut,
  Eye,
  Award,
  MapPin,
  BookOpen,
  Phone,
  Tag,
  Search,
  Shield,
  Truck,
  BookMarked,
  Car,
  Receipt,
} from "lucide-react";

// McMinnville Municipal Airport (KMMV), Oregon — used for the homepage weather widget
const HOME_COORDS = { lat: 45.1944, lon: -123.1364, tz: "America/Los_Angeles" };

// Maps Open-Meteo WMO weather codes to a short label + emoji
const WEATHER_CODES = {
  0: { label: "Clear sky", icon: "☀️" },
  1: { label: "Mostly clear", icon: "🌤️" },
  2: { label: "Partly cloudy", icon: "⛅" },
  3: { label: "Overcast", icon: "☁️" },
  45: { label: "Fog", icon: "🌫️" },
  48: { label: "Fog", icon: "🌫️" },
  51: { label: "Light drizzle", icon: "🌦️" },
  53: { label: "Drizzle", icon: "🌦️" },
  55: { label: "Heavy drizzle", icon: "🌧️" },
  61: { label: "Light rain", icon: "🌦️" },
  63: { label: "Rain", icon: "🌧️" },
  65: { label: "Heavy rain", icon: "🌧️" },
  71: { label: "Light snow", icon: "🌨️" },
  73: { label: "Snow", icon: "🌨️" },
  75: { label: "Heavy snow", icon: "❄️" },
  80: { label: "Rain showers", icon: "🌦️" },
  81: { label: "Rain showers", icon: "🌧️" },
  82: { label: "Violent showers", icon: "⛈️" },
  95: { label: "Thunderstorm", icon: "⛈️" },
  96: { label: "Thunderstorm w/ hail", icon: "⛈️" },
  99: { label: "Thunderstorm w/ hail", icon: "⛈️" },
};
const weatherInfo = (code) => WEATHER_CODES[code] || { label: "—", icon: "🌡️" };

// Crew members available for assignment across all tabs
const CREW_MEMBERS = ["Ryan Clifford", "Tom Fitzpatrick", "Daniel Lethin", "David Nemarnik"];
// Tasting House associates — starts empty since this is a different roster than the winery crew;
// build it via the "+ Add new" picker on Timesheets, or the Manage Lists panel under Backup.
const TASTING_ASSOCIATES = [];

// Cellar task categories for Work Orders
const WINERY_TASK_TYPES = ["Transfer", "Top Off", "Additions", "Rack", "Cold Stabilize", "Clean", "Inoculate", "Fine and Filter", "Pump Over", "Punch Down", "Fermentation"];
// Kept as an alias since existing code (cascades, saved data) checks against this name
const TASK_TYPES = WINERY_TASK_TYPES;
const VINEYARD_TASK_TYPES = ["Prune", "Canopy Management", "Spray / Pest Management", "Irrigate", "Mow / Cultivate", "Scout / Monitor", "Trellis / Wire Work", "Frost Protection", "Fruit Thinning", "Other"];
const WORKORDER_CATEGORIES = ["Vineyard", "Winery"];
function taskTypesForCategory(category) {
  return category === "Vineyard" ? VINEYARD_TASK_TYPES : WINERY_TASK_TYPES;
}
// Stages a Fermentation work order can mark — completing one of these, linked to a Lot,
// automatically advances that lot's position on the Fermentation Overview stage map.
const FERMENT_WORK_STAGES = ["Start Primary Fermentation", "Complete Primary Fermentation", "Start Malolactic Fermentation", "Complete Malolactic Fermentation"];
// Sub-types shown alongside Task Type when the task is an Addition
const ADDITION_TYPES = ["Acid", "Sugar", "Water", "Tannin", "Nutrients", "SO2"];
// Grape varieties available for Harvest Tonnage
const GRAPE_VARIETIES = ["Pinot Noir", "Chardonnay", "Riesling", "Muscat", "Nebbiolo", "Arneis", "Other"];
const US_STATES = [
  "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware",
  "District of Columbia", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa",
  "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota",
  "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey",
  "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon",
  "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah",
  "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming",
];
// Typical full-season GDD ranges (base 50°F, Winkler Index climate regions) for each variety —
// general viticultural reference, not a precise target. Actual optimal GDD varies by clone,
// site, and desired wine style; "Other" has no standard published range.
const VARIETY_GDD_REFERENCE = {
  "Pinot Noir": { min: 2000, max: 2500, region: "I" },
  "Chardonnay": { min: 2000, max: 2500, region: "I" },
  "Riesling": { min: 1700, max: 2300, region: "I" },
  "Muscat": { min: 2300, max: 2900, region: "I–II" },
  "Nebbiolo": { min: 3000, max: 3500, region: "III" },
  "Arneis": { min: 2500, max: 3000, region: "II" },
};
// Vineyard blocks available for Harvest Tonnage
const VINEYARD_BLOCKS = ["Antonina Block", "Solar Block", "Church Block", "Tasting House Block", "Winery Block"];
// Grape clones available for Harvest Tonnage
const GRAPE_CLONES = ["114", "115", "777", "Pommard", "Wadenswil", "76", "96", "Other"];
// Starter spray programs — rename, remove, or add more via the Vineyard work order picker or
// the Manage Lists panel under Backup; this is just a reasonable starting point.
const SPRAY_PROGRAMS = ["Powdery Mildew Program", "Botrytis Program", "Downy Mildew Program", "Organic IPM"];

// ---------- Simple sections (single-form, single-log) ----------
const SIMPLE_SECTIONS = [
  {
    key: "fruitAnalysis",
    label: "Fruit Analysis",
    icon: Beaker,
    sheetName: "Fruit Analysis",
    fields: [
      { name: "date", label: "Date", type: "date" },
      { name: "block", label: "Block / Vineyard", type: "block-picker" },
      { name: "brix", label: "Brix", type: "number" },
      { name: "ph", label: "pH", type: "number" },
      { name: "ta", label: "TA (g/L)", type: "number" },
      { name: "notes", label: "Notes", type: "textarea" },
    ],
  },
  {
    key: "harvest",
    label: "Harvest Tonnage",
    icon: Grape,
    sheetName: "Harvest Tonnage",
    fields: [
      { name: "date", label: "Date", type: "date" },
      { name: "block", label: "Block / Vineyard", type: "block-picker" },
      { name: "variety", label: "Variety", type: "select", options: GRAPE_VARIETIES },
      { name: "clone", label: "Clone", type: "clone-picker" },
      { name: "tons", label: "Gross Tons", type: "number" },
      { name: "lbs", label: "Gross Lbs", type: "number" },
      { name: "tareWeight", label: "Tare Weight (lbs)", type: "number", optional: true },
      { name: "netTons", label: "Net Tons", type: "number", optional: true },
      { name: "netLbs", label: "Net Lbs", type: "number", optional: true },
      { name: "weighMaster", label: "Weigh Master", type: "select", options: CREW_MEMBERS },
      { name: "notes", label: "Notes", type: "textarea" },
    ],
  },
  {
    key: "tanks",
    label: "Vessels",
    icon: Cylinder,
    sheetName: "Vessels",
    fields: [
      { name: "date", label: "Date", type: "date" },
      { name: "tankName", label: "Tank", type: "text" },
      { name: "capacityGal", label: "Capacity (gal)", type: "number" },
      { name: "status", label: "Status", type: "select", options: ["Empty", "In Use", "Cleaning", "Out of Service"] },
      { name: "contents", label: "Current Contents", type: "text" },
      { name: "notes", label: "Notes", type: "textarea" },
    ],
  },
  {
    key: "bottling",
    label: "Bottling",
    icon: Package2,
    sheetName: "Bottling",
    fields: [
      { name: "date", label: "Bottling Date", type: "date" },
      { name: "wineName", label: "Wine Name", type: "text" },
      { name: "vintage", label: "Vintage", type: "text" },
      { name: "blendReference", label: "Blend Reference", type: "text" },
      { name: "cases", label: "Cases", type: "number" },
      { name: "bottleSize", label: "Bottle Size", type: "select", options: ["750ml", "375ml", "500ml", "1.5L", "Other"] },
      { name: "lotCode", label: "Lot Code", type: "text" },
      { name: "notes", label: "Notes", type: "textarea" },
    ],
  },
  {
    key: "labResults",
    label: "Lab Results",
    icon: Microscope,
    sheetName: "Lab Results",
    fields: [
      { name: "date", label: "Analysis Date", type: "date" },
      { name: "labName", label: "Lab", type: "text" },
      { name: "sampleType", label: "Sample Type", type: "select", options: ["Juice", "Wine"] },
      { name: "targetLabel", label: "Lot / Vessel", type: "text" },
      { name: "brix", label: "Brix", type: "number" },
      { name: "ta", label: "TA (g/L)", type: "number" },
      { name: "ph", label: "pH", type: "number" },
      { name: "va", label: "VA (g/L)", type: "number" },
      { name: "malicAcid", label: "L-Malic Acid (g/L)", type: "number" },
      { name: "tartaricAcid", label: "Tartaric Acid (g/L)", type: "number" },
      { name: "glucoseFructose", label: "Glucose + Fructose (g/L)", type: "number" },
      { name: "ammonia", label: "Ammonia (mg/L)", type: "number" },
      { name: "alphaAminoN", label: "Alpha-Amino N (mg/L)", type: "number" },
      { name: "yan", label: "YAN (mg/L as N)", type: "number" },
      { name: "potassium", label: "Potassium (mg/L)", type: "number" },
      { name: "freeSO2", label: "Free SO2 (mg/L)", type: "number" },
      { name: "totalSO2", label: "Total SO2 (mg/L)", type: "number" },
      { name: "molecularSO2", label: "Molecular SO2 (mg/L)", type: "number" },
      { name: "ethanol", label: "Ethanol (% vol)", type: "number" },
      { name: "notes", label: "Notes", type: "textarea" },
    ],
  },
  {
    key: "vineHealth",
    label: "Vine Health",
    icon: Sprout,
    sheetName: "Vine Health",
    fields: [
      { name: "date", label: "Date", type: "date" },
      { name: "block", label: "Block / Vineyard", type: "block-picker" },
      { name: "observationType", label: "Observation Type", type: "select", options: ["Phenology Stage", "Pest Pressure", "Disease Alert"] },
      { name: "phenologyStage", label: "Phenology Stage (if Phenology)", type: "select", options: ["Bud Break", "Flowering", "Fruit Set", "Veraison"] },
      { name: "pestType", label: "Pest (if Pest Pressure)", type: "text" },
      { name: "diseaseType", label: "Disease (if Disease Alert)", type: "text" },
      { name: "severity", label: "Severity (if Pest or Disease)", type: "select", options: ["Low", "Medium", "High"] },
      { name: "notes", label: "Notes", type: "textarea" },
    ],
  },
  {
    key: "thoTimesheets",
    label: "Timesheets",
    icon: Clock,
    sheetName: "Timesheets",
    fields: [
      { name: "date", label: "Date", type: "date" },
      { name: "employeeName", label: "Associate", type: "associate-picker" },
      { name: "hoursWorked", label: "Hours Worked", type: "number" },
    ],
  },
  {
    key: "thoTips",
    label: "Tips",
    icon: DollarSign,
    sheetName: "Tips",
    fields: [
      { name: "date", label: "Date", type: "date" },
      { name: "totalTips", label: "Total Pooled Tips ($)", type: "number" },
    ],
  },
  {
    key: "thoMileage",
    label: "Mileage",
    icon: Car,
    sheetName: "Mileage",
    fields: [
      { name: "date", label: "Date", type: "date" },
      { name: "employeeName", label: "Associate", type: "associate-picker" },
      { name: "startLocation", label: "Starting Address", type: "text" },
      { name: "endLocation", label: "Destination Address", type: "text" },
      { name: "purpose", label: "Detailed Description", type: "text" },
      { name: "odometerStart", label: "Odometer Start", type: "number" },
      { name: "odometerEnd", label: "Odometer End", type: "number" },
      { name: "miles", label: "Mileage", type: "number" },
    ],
  },
  {
    key: "thoExpenses",
    label: "Expenses",
    icon: Receipt,
    sheetName: "Expenses",
    fields: [
      { name: "date", label: "Date", type: "date" },
      { name: "employeeName", label: "Associate", type: "associate-picker" },
      { name: "expenseName", label: "Expense Name", type: "text" },
      { name: "glCode", label: "GL Code", type: "text" },
      { name: "category", label: "Category", type: "select", options: ["Supplies", "Meals", "Equipment", "Other"] },
      { name: "amount", label: "Amount ($)", type: "number" },
      { name: "description", label: "Detailed Description", type: "text" },
      { name: "receipt", label: "Receipt Photo", type: "photo" },
    ],
  },
  {
    key: "accolades",
    label: "Accolades",
    icon: Award,
    sheetName: "Accolades",
    fields: [
      { name: "date", label: "Date", type: "date" },
      { name: "wineName", label: "Wine Name", type: "text" },
      { name: "vintage", label: "Vintage", type: "text" },
      { name: "publication", label: "Publication / Critic", type: "text" },
      { name: "score", label: "Score / Rating", type: "text" },
      { name: "notes", label: "Review Notes / Quote", type: "textarea" },
    ],
  },
  {
    key: "vineyardBlockDetails",
    label: "Vineyard Blocks",
    icon: MapPin,
    sheetName: "Vineyard Blocks",
    fields: [
      { name: "block", label: "Block / Vineyard", type: "block-picker" },
      { name: "variety", label: "Variety", type: "select", options: GRAPE_VARIETIES },
      { name: "acreage", label: "Acreage", type: "number" },
      { name: "soilType", label: "Soil Type", type: "text" },
    ],
  },
  {
    key: "contacts",
    label: "Contacts",
    icon: Phone,
    sheetName: "Contacts",
    fields: [
      { name: "name", label: "Name", type: "text" },
      { name: "type", label: "Type", type: "select", options: ["Employee", "Vendor", "Emergency Contact"] },
      { name: "roleOrCompany", label: "Role / Company", type: "text" },
      { name: "phone", label: "Phone", type: "text" },
      { name: "email", label: "Email", type: "text" },
      { name: "notes", label: "Notes", type: "textarea" },
    ],
  },
  {
    key: "winePricing",
    label: "Pricing",
    icon: Tag,
    sheetName: "Pricing",
    fields: [
      { name: "wineName", label: "Wine Name", type: "text" },
      { name: "vintage", label: "Vintage", type: "text" },
      { name: "wholesalePrice", label: "Wholesale Price ($)", type: "number" },
      { name: "retailPrice", label: "Retail Price ($)", type: "number" },
      { name: "notes", label: "Notes", type: "textarea" },
    ],
  },
  {
    key: "wineClubTiers",
    label: "Wine Club",
    icon: Wine,
    sheetName: "Wine Club",
    fields: [
      { name: "tierName", label: "Tier Name", type: "text" },
      { name: "price", label: "Price per Shipment ($)", type: "number" },
      { name: "shipmentFrequency", label: "Shipment Frequency", type: "text" },
      { name: "benefits", label: "Benefits", type: "textarea" },
      { name: "notes", label: "Notes", type: "textarea" },
    ],
  },
  {
    key: "statePermits",
    label: "My Permits",
    icon: Shield,
    sheetName: "State Permits",
    fields: [
      { name: "state", label: "State", type: "select", options: US_STATES },
      { name: "permitNumber", label: "Permit / License #", type: "text" },
      { name: "issuingAgency", label: "Issuing Agency", type: "text" },
      { name: "issueDate", label: "Issue Date", type: "date" },
      { name: "expirationDate", label: "Expiration / Renewal Date", type: "date" },
      { name: "reminderLeadDays", label: "Reminder Lead Time (days)", type: "number" },
      { name: "renewalCost", label: "Renewal Cost ($)", type: "number", optional: true },
      { name: "selfNotedCap", label: "Self-Noted Annual Case Cap (optional, informational only)", type: "number", optional: true },
      { name: "permitDocument", label: "Permit Document", type: "document", optional: true },
      { name: "notes", label: "Notes", type: "textarea" },
    ],
  },
  {
    key: "shipmentLog",
    label: "Shipment Log",
    icon: Truck,
    sheetName: "Shipment Log",
    fields: [
      { name: "date", label: "Date", type: "date" },
      { name: "destinationState", label: "Destination State", type: "select", options: US_STATES },
      { name: "cases", label: "Cases Shipped", type: "number" },
      { name: "volumeGallons", label: "Volume (gallons)", type: "number", optional: true },
      { name: "orderValue", label: "Order Value ($)", type: "number", optional: true },
      { name: "taxCollected", label: "Tax Collected ($, self-reported)", type: "number", optional: true },
      { name: "referenceOrderNumber", label: "Reference Invoice / Order #", type: "text", optional: true },
    ],
  },
  {
    key: "stateRuleReference",
    label: "State Reference",
    icon: BookMarked,
    sheetName: "State Reference",
    fields: [
      { name: "state", label: "State", type: "select", options: US_STATES },
      { name: "dtcStatus", label: "DTC Shipping Status", type: "select", options: ["Allowed", "Restricted", "Banned", "Not Yet Verified"] },
      { name: "restrictionNotes", label: "Restriction Notes (plain English)", type: "textarea" },
      { name: "annualCapPerHousehold", label: "Annual Cap per Household (if known)", type: "text", optional: true },
      { name: "reportFrequency", label: "Typical Report Frequency", type: "text", optional: true },
      { name: "lastVerifiedDate", label: "Last Verified Date", type: "date", optional: true },
      { name: "sourceLink", label: "Source Link", type: "text", optional: true },
    ],
  },
];

// Work order fields shown when adding a new to-do
const WORKORDER_FIELDS = [
  { name: "category", label: "Category", type: "select", options: WORKORDER_CATEGORIES },
  { name: "dateAssigned", label: "Date Assigned", type: "date" },
  { name: "date", label: "Due Date", type: "date" },
  { name: "assignedTo", label: "Assigned To", type: "select", options: CREW_MEMBERS },
  { name: "task", label: "Task", type: "text" },
  { name: "taskType", label: "Task Type", type: "select", options: TASK_TYPES },
  { name: "additionType", label: "Addition Type", type: "select", options: ADDITION_TYPES },
  { name: "fermentStage", label: "Fermentation Stage", type: "select", options: FERMENT_WORK_STAGES },
  { name: "sprayProgram", label: "Spray Program", type: "spray-program-picker" },
  { name: "lots", label: "Lots", type: "lots-picker" },
  { name: "barrels", label: "Barrels", type: "barrels-picker" },
  { name: "calculations", label: "Calculations", type: "textarea" },
  { name: "directions", label: "Directions", type: "textarea" },
  { name: "priority", label: "Priority", type: "select", options: ["Low", "Medium", "High"] },
  { name: "notes", label: "Notes", type: "textarea" },
];
// Same fields as a work order, minus the dates and the specific lots/barrels for this one
// instance — since a template gets fresh dates (and a fresh set of lots/barrels) each time it's used
const TEMPLATE_FIELDS = WORKORDER_FIELDS.filter((f) => !["date", "dateAssigned", "lots", "barrels"].includes(f.name));
// Returns WORKORDER_FIELDS with the Task Type options swapped to match the chosen category, so
// Vineyard work orders offer vineyard tasks and Winery work orders offer winery tasks
function workOrderFieldsForCategory(fields, category) {
  return fields.map((f) => (f.name === "taskType" ? { ...f, options: taskTypesForCategory(category) } : f));
}
// Hides fields that only make sense for a specific Task Type until that type is actually
// selected — e.g. "Addition Type" only shows once Task Type is "Additions" — instead of always
// showing every field with a "(if X)" hedge label regardless of what's actually relevant.
function visibleWorkOrderFields(fields, form) {
  return fields.filter((f) => {
    if (f.name === "lots" || f.name === "barrels") return !!form.taskType;
    if (f.name === "additionType") return form.taskType === "Additions";
    if (f.name === "fermentStage") return form.taskType === "Fermentation";
    if (f.name === "sprayProgram") return form.taskType === "Spray / Pest Management";
    return true;
  });
}
// Formats an auto-assigned work order number like 7 -> "WO-0007"
const formatOrderNumber = (n) => (n ? `WO-${String(n).padStart(4, "0")}` : "");

// Vessel types available when starting a new ferment
// Specific vessel IDs (e.g. "B10", "T5") — starts empty since these are specific to your cellar;
// build the list via the Vessel field's "+ Add new" or the Manage Lists panel under Backup.
const VESSEL_TYPES = [];
// Malolactic fermentation status options, tracked separately from primary fermentation
const ML_STATUSES = ["Not Started", "Inoculated", "In Progress", "Complete"];

// Fermentation lot static fields (asked once, when the ferment starts)
// Vintage year options — a reasonable spread around the current year
const VINTAGE_YEARS = Array.from({ length: 10 }, (_, i) => String(new Date().getFullYear() - 8 + i));
// Wine styles across your portfolio (reds, whites, rosé, dessert, sparkling)
const WINE_STYLES = ["Dry Red", "Dry White", "Rosé", "Dessert / Late Harvest", "Sparkling", "Other"];

const FERMENT_LOT_FIELDS = [
  { name: "tankId", label: "Lot", type: "text" },
  { name: "vessel", label: "Vessel", type: "vessel-picker" },
  { name: "variety", label: "Variety", type: "select", options: GRAPE_VARIETIES },
  { name: "vintage", label: "Vintage", type: "select", options: VINTAGE_YEARS },
  { name: "wineStyle", label: "Wine Style", type: "select", options: WINE_STYLES },
  { name: "startDate", label: "Start Date", type: "date" },
  { name: "startingBrix", label: "Starting Brix", type: "number" },
  { name: "startingTemp", label: "Starting Temp (°F)", type: "number" },
  { name: "notes", label: "Starting Notes", type: "textarea" },
];

// Matches the two check-in rounds you actually use
const FERMENT_SESSIONS = ["A.M.", "P.M."];
// Guesses which session it is right now, for auto-logging and for defaulting the Quick Log picker
function guessSession() {
  return new Date().getHours() < 12 ? "A.M." : "P.M.";
}

// Daily reading fields (asked every day, per ferment)
const FERMENT_READING_FIELDS = [
  { name: "date", label: "Date", type: "date" },
  { name: "session", label: "Session", type: "select", options: FERMENT_SESSIONS },
  { name: "workDone", label: "Work Done", type: "checkbox-group", options: ["Punch Down", "Pump Over", "Cold Soak"] },
  { name: "additions", label: "Additions", type: "checkbox-group", options: ADDITION_TYPES },
  { name: "brix", label: "Brix", type: "number" },
  { name: "temp", label: "Temp (°F)", type: "number" },
  { name: "ph", label: "pH", type: "number" },
  { name: "notes", label: "Notes", type: "textarea" },
];

// Builds the Fermentation lots + Barrel records (with fills linked by id) for the one-time 2025 import
// Real lab results from two sample reports the user shared, used for a one-time import so they
// can see actual data in the new Lab Results section rather than an empty tab.
const LAB_IMPORT_DATA = [
  // ETS Laboratories — juice/must panel, report #2240307O, 10/18/2025
  { date: "2025-10-18", labName: "ETS Laboratories", sampleType: "Juice", targetLabel: "Dessert",
    brix: 33.7, ta: 7.1, ph: 3.08, va: 0.05, malicAcid: 3.55, tartaricAcid: 5.0, glucoseFructose: 374,
    ammonia: 94, alphaAminoN: 128, yan: 205, potassium: 600,
    notes: "VA reported as <0.05 g/L (below detection limit)." },
  { date: "2025-10-18", labName: "ETS Laboratories", sampleType: "Juice", targetLabel: "Nebbiolo",
    brix: 23.7, ta: 8.6, ph: 2.92, va: 0.05, malicAcid: 3.32, tartaricAcid: 6.7, glucoseFructose: 243,
    ammonia: 25, alphaAminoN: 70, yan: 91, potassium: 850,
    notes: "VA reported as <0.05 g/L (below detection limit)." },

  // Core Enology Analytical Services — wine panel, Order #45634, 4/16/2025
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "114 B2",
    ta: 5.2, ph: 3.71, va: 0.52, freeSO2: 32, molecularSO2: 0.39, ethanol: 13.41 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "114 B4",
    ta: 5.1, ph: 3.73, va: 0.52, freeSO2: 28, molecularSO2: 0.33, ethanol: 13.67 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "11A MIX B3",
    ta: 5.0, ph: 3.76, va: 0.59, freeSO2: 28, molecularSO2: 0.30, ethanol: 13.34 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "114 T4",
    ta: 5.2, ph: 3.77, va: 0.52, freeSO2: 29, molecularSO2: 0.31, ethanol: 13.54 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "777 T1",
    ta: 5.0, ph: 3.64, va: 0.45, freeSO2: 31, molecularSO2: 0.44, ethanol: 13.00 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "777 T2",
    ta: 5.0, ph: 3.62, va: 0.45, freeSO2: 34, molecularSO2: 0.51, ethanol: 12.96 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "777 T3",
    ta: 5.1, ph: 3.64, va: 0.43, freeSO2: 32, molecularSO2: 0.46, ethanol: 13.00 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "P115 B16",
    ta: 5.3, ph: 3.65, va: 0.43, freeSO2: 31, molecularSO2: 0.44, ethanol: 14.45 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "P777 MIX B14",
    ta: 5.2, ph: 3.67, va: 0.44, freeSO2: 27, molecularSO2: 0.36, ethanol: 14.40 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "P777 WC B5",
    ta: 5.2, ph: 3.87, va: 0.98, freeSO2: 33, molecularSO2: 0.28, ethanol: 13.79 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "PMIX B15",
    ta: 5.2, ph: 3.69, va: 0.43, freeSO2: 31, molecularSO2: 0.40, ethanol: 14.78 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "POM B6",
    ta: 5.0, ph: 3.73, va: 0.49, freeSO2: 34, molecularSO2: 0.40, ethanol: 13.88 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "POM B7",
    ta: 5.1, ph: 3.71, va: 0.47, freeSO2: 36, molecularSO2: 0.44, ethanol: 13.99 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "POM B8",
    ta: 5.0, ph: 3.74, va: 0.48, freeSO2: 31, molecularSO2: 0.35, ethanol: 14.19 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "POM B11",
    ta: 5.1, ph: 3.75, va: 0.48, freeSO2: 26, molecularSO2: 0.29, ethanol: 14.57 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "POM MIX B12",
    ta: 5.2, ph: 3.69, va: 0.43, freeSO2: 29, molecularSO2: 0.37, ethanol: 14.43 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "POM T5",
    ta: 5.2, ph: 3.76, va: 0.46, freeSO2: 32, molecularSO2: 0.35, ethanol: 14.28 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "POM WC B9",
    ta: 4.9, ph: 3.89, va: 0.60, freeSO2: 24, molecularSO2: 0.19, ethanol: 13.58 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "POM WC B10",
    ta: 5.0, ph: 4.00, va: 0.93, freeSO2: 34, molecularSO2: 0.21, ethanol: 13.67 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "WAD B13",
    ta: 5.3, ph: 3.60, va: 0.36, freeSO2: 30, molecularSO2: 0.47, ethanol: 13.97 },
  { date: "2025-04-16", labName: "Core Enology Analytical Services", sampleType: "Wine", targetLabel: "YV 114 B1",
    ta: 5.1, ph: 3.70, va: 0.54, freeSO2: 31, molecularSO2: 0.39, ethanol: 14.05 },
];

function buildLabResultsImport() {
  return LAB_IMPORT_DATA.map((r) => ({ id: genId(), totalSO2: "", notes: "", ...r }));
}

function buildBarrelImport2025() {
  const lines = BARREL_IMPORT_2025_RAW.split("\n").map((l) => l.trim()).filter(Boolean);
  const lotIdByCode = {};
  const fermentLots = [];
  const barrels = [];

  lines.forEach((line) => {
    const [barrelNumber, ...rest] = line.split(/\s+/);
    const lotCode = rest.join(" ");
    if (!lotCode) return;

    if (!lotIdByCode[lotCode]) {
      const id = genId();
      lotIdByCode[lotCode] = id;
      fermentLots.push({
        id,
        tankId: lotCode,
        vessel: "",
        variety: LOT_VARIETY_GUESS[lotCode] || "Other",
        startDate: "",
        startingBrix: "",
        startingTemp: "",
        status: "Complete",
        notes: "Bulk-imported with the 2025 barrel program — please verify variety and start date.",
        readings: [],
      });
    }

    barrels.push({
      id: genId(),
      barrelNumber,
      size: "225",
      toast: "",
      notes: "",
      soldDate: "",
      soldTo: "",
      soldContact: "",
      soldPrice: "",
      saleNotes: "",
      fills: [{ id: genId(), lotId: lotIdByCode[lotCode], fillDate: "", emptyDate: "" }],
    });
  });

  return { fermentLots, barrels };
}

// One-time 2025 barrel program import: barrel # paired with its cellar lot/blend code.
// Each line is "BARREL# LOT CODE" (lot code may itself contain spaces).
const BARREL_IMPORT_2025_RAW = `
23-C17 114 T5
25-C07 114 T5
25-E05 114 T5
25-D05 114 T5
25-R04 114 T5
24-C12 114 T5
24-E09 114 T5
24-R02 114 T5
23-C18 114 T5
23-D15 114 T5
22-C25 114 T5
25-C01 a777WC B4
24-C14 a777WC B4
24-D07 a777WC B4
25-D06 WadWC B5
24-C06 WadWC B5
24-E02 WadWC B5
25-C02 115WC B6
25-E01 115WC B6
24-C10 115WC B6
25-R01 p777WC B7
24-D05 p777WC B7
23-C20 p777WC B7
25-C03 777 T1
25-E02 777 T1
25-D07 777 T1
25-F01 777 T1
24-C01 777 T1
24-D11 777 T1
23-C27 777 T1
23-R41 777 T1
22-C21 777 T1
25-M01 777 T2
25-C04 777 T2
25-R02 777 T2
25-B01 777 T2
24-C08 777 T2
24-E01 777 T2
23-C25 777 T2
23-E37 777 T2
23-D12 777 T2
25-C06 777YV114 T3
25-E03 777YV114 T3
25-R03 777YV114 T3
25-F02 777YV114 T3
24-C13 777YV114 T3
24-D06 777YV114 T3
24-R03 777YV114 T3
23-C28 777YV114 T3
23-E38 777YV114 T3
25-C05 NWPOM T4
25-D08 NWPOM T4
25-E04 NWPOM T4
25-M02 NWPOM T4
24-C03 NWPOM T4
24-D12 NWPOM T4
24-E10 NWPOM T4
23-D09 NWPOM T4
23-R44 NWPOM T4
23-C21 NWPOM T4
25-C08 ChPOM B1
25-E06 ChPOM B1
24-C11 ChPOM B1
24-R05 ChPOM B1
25-C09 114 B2
25L-R05 114 B2
24-C04 114 B2
23-C19 114 B2
25-C13 115WAD B8
25-E09 115WAD B8
24-D10 115WAD B8
23-C29 115WAD B8
25-C14 BRAD B9
24-R01 BRAD B9
23-D10 BRAD B9
25-C12 a777POM B3
25-D01 a777POM B3
24-C02 a777POM B3
23-D14 a777POM B3
25-D02 JAY WC B10
24-C05 JAY WC B10
23-D16 JAY WC B10
25-C10 ChPOM T6
25-D04 ChPOM T6
25-E07 ChPOM T6
25-M03 ChPOM T6
24-C07 ChPOM T6
24-D08 ChPOM T6
24-E08 ChPOM T6
23-D13 ChPOM T6
23-R43 ChPOM T6
23-C24 ChPOM T6
25-C11 777 T2
25-E08 777 T2
25-D03 777 T2
25L-R06 777 T2
24-C09 777 T2
24-E03 777 T2
24-R04 777 T2
23-C22 777 T2
23-D11 777 T2
23-C23 777 T2
`.trim();

// Best-guess variety per lot code, based on standard Pinot Noir clone numbers (114/115/777/Pommard/
// Wadenswil) vs. the "Ch" prefix for Chardonnay. Flagged to the user as an assumption to verify.
const LOT_VARIETY_GUESS = {
  "114 T5": "Pinot Noir",
  "a777WC B4": "Pinot Noir",
  "WadWC B5": "Pinot Noir",
  "115WC B6": "Pinot Noir",
  "p777WC B7": "Pinot Noir",
  "777 T1": "Pinot Noir",
  "777 T2": "Pinot Noir",
  "777YV114 T3": "Pinot Noir",
  "NWPOM T4": "Pinot Noir",
  "ChPOM B1": "Chardonnay",
  "114 B2": "Pinot Noir",
  "115WAD B8": "Pinot Noir",
  "BRAD B9": "Pinot Noir",
  "a777POM B3": "Pinot Noir",
  "JAY WC B10": "Pinot Noir",
  "ChPOM T6": "Chardonnay",
};

const COOPERAGES = { C: "Cadus", D: "Damy", E: "Ermitage", R: "Remond", M: "Mercury", F: "Francois Freres", B: "Billon" };
// Cooperage names offered in the Add a Barrel dropdown (same makers, plus a catch-all)
const COOPERAGE_NAMES = [...Object.values(COOPERAGES), "Other"];
const BARREL_TOAST_LEVELS = ["Light", "Medium", "Medium+", "Heavy"];
const BARREL_WINE_COLORS = ["Red", "White"];

// ---------- Blending & barrel tasting domain ----------
const RATING_DIMENSIONS = [
  { key: "tannin", label: "Tannin" },
  { key: "acid", label: "Acid" },
  { key: "body", label: "Body" },
  { key: "aroma", label: "Aroma Intensity" },
  { key: "flavorIntensity", label: "Flavor Intensity" },
  { key: "balance", label: "Overall Balance" },
];

// Most recent tasting on file for a given barrel, or null if it's never been tasted
function latestTastingForBarrel(barrelId, tastings) {
  const entries = (tastings || []).filter((t) => t.barrelId === barrelId).sort((a, b) => (a.date < b.date ? 1 : -1));
  return entries[0] || null;
}

// Percentage-weighted average rating profile for a blend, based on each component barrel's
// most recent tasting. This is a simplification (real blending isn't strictly linear) but gives
// a useful directional comparison between blend trials.
function computeBlendProfile(blend, tastings) {
  const dims = RATING_DIMENSIONS.map((d) => d.key);
  let totalPct = 0;
  const sums = Object.fromEntries(dims.map((k) => [k, 0]));
  let tastedCount = 0;
  (blend.components || []).forEach((c) => {
    const t = latestTastingForBarrel(c.barrelId, tastings);
    const pct = parseFloat(c.percentage) || 0;
    if (t && pct > 0) {
      tastedCount += 1;
      dims.forEach((k) => {
        sums[k] += (parseFloat(t[k]) || 0) * pct;
      });
      totalPct += pct;
    }
  });
  if (totalPct === 0) return null;
  const profile = {};
  dims.forEach((k) => {
    profile[k] = sums[k] / totalPct;
  });
  return { profile, tastedCount, totalComponents: (blend.components || []).length };
}

const BARREL_FOREST_ORIGINS = ["Allier", "Tronçais", "Vosges", "Nevers", "Center France (blend)", "American Oak", "Hungarian Oak", "Other"];

// Parses a barrel number like "25-C01" or "25L-R05" into { year, series, code, sequence, cooperage }
function parseBarrelNumber(num) {
  if (!num) return null;
  const match = /^(\d{2})([A-Za-z]?)-([A-Za-z])(\d{1,3})$/.exec(num.trim());
  if (!match) return null;
  const [, year, series, code, sequence] = match;
  const upperCode = code.toUpperCase();
  return { year, series: series || "", code: upperCode, sequence, cooperage: COOPERAGES[upperCode] || "Unknown" };
}

// A barrel's status is derived from its data, never stored directly, so it can never drift out of sync
function barrelStatus(barrel) {
  if (barrel.soldDate) return "Sold";
  if (barrel.retiredDate) return "Retired";
  return barrel.fills.some((f) => !f.emptyDate) ? "In Use" : "Empty";
}
function activeBarrelFill(barrel) {
  return barrel.fills.find((f) => !f.emptyDate) || null;
}

// Reads a fill's lot makeup — supports both the new multi-lot blend format (components: [{lotLabel, percentage}])
// and older fills that just had a single lotId pointing at a Fermentation lot. Always returns an array.
function getFillComponents(fill, fermentLots) {
  if (!fill) return [];
  if (Array.isArray(fill.components) && fill.components.length > 0) {
    return fill.components.map((c) => {
      const lot = c.lotId ? (fermentLots || []).find((l) => l.id === c.lotId) : null;
      return {
        label: c.lotLabel || lot?.tankId || "Unnamed lot",
        variety: lot?.variety || "",
        percentage: c.percentage,
      };
    });
  }
  if (fill.lotId) {
    const lot = (fermentLots || []).find((l) => l.id === fill.lotId);
    return [{ label: lot?.tankId || "Unnamed lot", variety: lot?.variety || "", percentage: 100 }];
  }
  return [];
}
// One-line summary like "114 T5 (60%) + 777 T2 (40%)"
function summarizeFillComponents(fill, fermentLots) {
  const parts = getFillComponents(fill, fermentLots);
  if (parts.length === 0) return "";
  if (parts.length === 1 && (parts[0].percentage == null || parts[0].percentage === 100)) return parts[0].label;
  return parts.map((p) => `${p.label}${p.percentage != null && p.percentage !== "" ? ` (${p.percentage}%)` : ""}`).join(" + ");
}

// Parses pasted bulk-import text into barrel records. One barrel per line:
// "25-C01" or "25-C01, Medium" (toast level is optional). All barrels are 225L.
function parseBulkBarrels(text) {
  const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
  const barrels = [];
  const errors = [];
  lines.forEach((line, i) => {
    const parts = line.split(",").map((p) => p.trim());
    const barrelNumber = parts[0];
    if (!parseBarrelNumber(barrelNumber)) {
      errors.push(`Line ${i + 1}: "${barrelNumber}" doesn't match the format YY-C## (e.g. 25-C01)`);
      return;
    }
    barrels.push({
      id: genId(),
      barrelNumber,
      size: "225",
      cooperage: "",
      wineColor: "",
      forest: "",
      toast: parts[1] || "",
      notes: "",
      soldDate: "",
      soldTo: "",
      soldContact: "",
      soldPrice: "",
      saleNotes: "",
      retiredDate: "",
      retiredReason: "",
      fills: [],
    });
  });
  return { barrels, errors };
}

// Describes which use this barrel is on (1st, 2nd, 3rd+/Neutral), derived from fill history
function fillCountLabel(barrel) {
  const total = barrel.fills.length;
  const ordinal = (n) => (n === 1 ? "1st" : n === 2 ? "2nd" : `${n}th`);
  const activeIndex = barrel.fills.findIndex((f) => !f.emptyDate);
  if (activeIndex >= 0) {
    const useNum = activeIndex + 1;
    return `${ordinal(useNum)} Use${useNum >= 3 ? " (Neutral)" : ""}`;
  }
  if (total === 0) return "New / Unused";
  const nextUse = total + 1;
  return `Next: ${ordinal(nextUse)} Use${nextUse >= 3 ? " (Neutral)" : ""}`;
}

const ALL_TABS = [
  { key: "home", label: "Dashboard", icon: Home },
  { key: "workorders", label: "Work Orders", icon: ClipboardList },
  SIMPLE_SECTIONS.find((s) => s.key === "fruitAnalysis"),
  SIMPLE_SECTIONS.find((s) => s.key === "harvest"),
  SIMPLE_SECTIONS.find((s) => s.key === "vineHealth"),
  { key: "ferment", label: "Fermentation", icon: FlaskConical },
  { key: "barrels", label: "Barrels", icon: Package },
  { key: "blending", label: "Blending", icon: Wine },
  SIMPLE_SECTIONS.find((s) => s.key === "tanks"),
  SIMPLE_SECTIONS.find((s) => s.key === "bottling"),
  { key: "techSheetBuilder", label: "Tech Sheet Builder", icon: FileText },
  SIMPLE_SECTIONS.find((s) => s.key === "labResults"),
  { key: "compliance", label: "Compliance", icon: Shield },
  { key: "aboutAlloro", label: "Team Resources", icon: BookOpen },
  { key: "thoPayroll", label: "Payroll", icon: DollarSign },
  SIMPLE_SECTIONS.find((s) => s.key === "thoMileage"),
  SIMPLE_SECTIONS.find((s) => s.key === "thoExpenses"),
  { key: "calendar", label: "Calendar", icon: Calendar },
  { key: "formulas", label: "Formulas", icon: Calculator },
  { key: "backup", label: "Backup", icon: HardDrive },
];

// Dashboard and Work Orders stay visible at all times, outside any category — Dashboard because
// it's the landing page and meant to be a cross-cutting overview, Work Orders because it's used
// constantly for both vineyard and cellar tasks and shouldn't cost an extra tap to reach.
const PERSISTENT_NAV_KEYS = ["home", "workorders"];

// Everything else groups into three categories. Clicking a category reveals its tabs as a
// second row; clicking a tab within it is remembered, so returning to that category later goes
// straight back to where you left off instead of resetting to the first tab every time.
const NAV_CATEGORIES = {
  vineyard: { label: "Vineyard", dotColor: "bg-lime-400", keys: ["fruitAnalysis", "harvest", "vineHealth"] },
  winery: { label: "Winery", dotColor: "bg-amber-400", keys: ["ferment", "barrels", "blending", "tanks", "bottling", "techSheetBuilder"] },
  tho: { label: "THO", dotColor: "bg-rose-400", keys: ["thoPayroll", "thoMileage", "thoExpenses", "aboutAlloro"] },
  data: { label: "Data", dotColor: "bg-sky-300", keys: ["labResults", "compliance", "calendar", "formulas", "backup"] },
};

function categoryOfKey(key) {
  for (const [catKey, cat] of Object.entries(NAV_CATEGORIES)) {
    if (cat.keys.includes(key)) return catKey;
  }
  return null;
}

// Flattens all app data into a common shape ({title, headers, rows}) shared by every export format
function buildExportSections(data) {
  const sections = [];

  const woHeaders = ["Work Order #", "Date Assigned", "Due Date", "Assigned To", "Task", "Task Type", "Addition Type", "Lots", "Barrels", "Calculations", "Directions", "Priority", "Status", "Date Completed", "Notes"];
  const woRows = data.workorders.map((o) => ({
    "Work Order #": formatOrderNumber(o.orderNumber), "Date Assigned": o.dateAssigned, "Due Date": o.date, "Assigned To": o.assignedTo, "Task": o.task,
    "Task Type": o.taskType, "Addition Type": o.taskType === "Additions" ? o.additionType : "",
    "Lots": Array.isArray(o.lots) ? o.lots.join(", ") : "",
    "Barrels": Array.isArray(o.barrels) ? o.barrels.map((id) => (data.barrels || []).find((b) => b.id === id)?.barrelNumber).filter(Boolean).join(", ") : "",
    "Calculations": o.calculations, "Directions": o.directions,
    "Priority": o.priority, "Status": o.status, "Date Completed": o.dateCompleted, "Notes": o.notes,
  }));
  sections.push({ title: "Work Orders", headers: woHeaders, rows: woRows });

  SIMPLE_SECTIONS.forEach((section) => {
    const headers = section.fields.map((f) => f.label);
    const rows = data[section.key].map((row) =>
      Object.fromEntries(section.fields.map((f) => [
        f.label,
        (f.type === "photo" || f.type === "document") ? (row[f.name] ? "Yes" : "No") : (row[f.name] ?? ""),
      ]))
    );
    sections.push({ title: section.sheetName, headers, rows });
  });

  const fermentHeaders = ["Tank / Lot ID", "Vessel", "Variety", "Start Date", "Starting Brix", "Starting Temp (°F)", "Status", "Reading Date", "Work Done", "Brix", "Temp (°F)", "pH", "Notes"];
  const fermentRows = [];
  data.ferment.forEach((lot) => {
    const lotBase = {
      "Tank / Lot ID": lot.tankId, "Vessel": lot.vessel, "Variety": lot.variety, "Start Date": lot.startDate,
      "Starting Brix": lot.startingBrix, "Starting Temp (°F)": lot.startingTemp, "Status": lot.status,
    };
    if (lot.readings.length === 0) {
      fermentRows.push({
        ...lotBase, "Reading Date": "", "Work Done": "", "Brix": "", "Temp (°F)": "", "pH": "", "Notes": lot.notes,
      });
    } else {
      [...lot.readings].sort((a, b) => (a.date > b.date ? 1 : -1)).forEach((r) => {
        fermentRows.push({
          ...lotBase, "Reading Date": r.date,
          "Work Done": Array.isArray(r.workDone) ? r.workDone.join(", ") : "",
          "Brix": r.brix, "Temp (°F)": r.temp, "pH": r.ph, "Notes": r.notes,
        });
      });
    }
  });
  sections.push({ title: "Fermentation", headers: fermentHeaders, rows: fermentRows });

  const barrelHeaders = [
    "Barrel #", "Cooperage", "Barrel Year", "Wine Color", "Forest Origin", "Toast", "Status", "Fill Count",
    "Fill Lot", "Fill Variety", "Fill Date", "Empty Date", "Sold Date", "Sold To", "Buyer Contact", "Sale Price",
    "Retired Date", "Retired Reason", "Notes",
  ];
  const barrelRows = [];
  (data.barrels || []).forEach((b) => {
    const parsed = parseBarrelNumber(b.barrelNumber);
    const base = {
      "Barrel #": b.barrelNumber, "Cooperage": b.cooperage || parsed?.cooperage || "", "Barrel Year": parsed ? `20${parsed.year}` : "",
      "Wine Color": b.wineColor || "", "Forest Origin": b.forest || "",
      "Toast": b.toast, "Status": barrelStatus(b), "Fill Count": b.fills.length,
      "Sold Date": b.soldDate, "Sold To": b.soldTo, "Buyer Contact": b.soldContact, "Sale Price": b.soldPrice,
      "Retired Date": b.retiredDate, "Retired Reason": b.retiredReason, "Notes": b.notes,
    };
    if (b.fills.length === 0) {
      barrelRows.push({ ...base, "Fill Lot": "", "Fill Variety": "", "Fill Date": "", "Empty Date": "" });
    } else {
      [...b.fills].sort((x, y) => (x.fillDate > y.fillDate ? 1 : -1)).forEach((f) => {
        const components = getFillComponents(f, data.ferment);
        barrelRows.push({
          ...base,
          "Fill Lot": summarizeFillComponents(f, data.ferment),
          "Fill Variety": [...new Set(components.map((c) => c.variety).filter(Boolean))].join(", "),
          "Fill Date": f.fillDate, "Empty Date": f.emptyDate,
        });
      });
    }
  });
  sections.push({ title: "Barrels", headers: barrelHeaders, rows: barrelRows });

  return sections;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function escapeHtml(val) {
  return String(val ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

const emptyForm = (fields) =>
  Object.fromEntries(
    fields.map((f) => [f.name, ["checkbox-group", "lots-picker", "barrels-picker"].includes(f.type) ? [] : ""])
  );
const todayISO = () => new Date().toISOString().split("T")[0];
// Like todayISO, but resolves "today" in a specific IANA timezone instead of UTC — needed
// wherever we're matching against Open-Meteo data (which is fetched with an explicit timezone),
// since UTC has usually already rolled over to the next calendar day by Pacific-time evening.
function todayInTimezone(tz) {
  try {
    return new Intl.DateTimeFormat("en-CA", { timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
  } catch {
    return todayISO();
  }
}
const genId = () => `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
const pad2 = (n) => String(n).padStart(2, "0");
const isoFor = (y, m, d) => `${y}-${pad2(m + 1)}-${pad2(d)}`;
const inRange = (dateStr, from, to) => {
  if (!dateStr) return false;
  if (from && dateStr < from) return false;
  if (to && dateStr > to) return false;
  return true;
};

// ---------- Sorting helpers, shared by on-screen tables and every export/print path ----------
const PRIORITY_RANK = { High: 3, Medium: 2, Low: 1 };
function compareValues(a, b, field) {
  if (field === "priority") return (PRIORITY_RANK[a] || 0) - (PRIORITY_RANK[b] || 0);
  const na = parseFloat(a), nb = parseFloat(b);
  const bothNumeric = a !== "" && a != null && b !== "" && b != null && !isNaN(na) && !isNaN(nb);
  if (bothNumeric) return na - nb;
  return String(a ?? "").localeCompare(String(b ?? ""));
}
function sortRows(rows, field, direction) {
  if (!field) return rows;
  const sorted = [...rows].sort((a, b) => compareValues(a[field], b[field], field));
  return direction === "desc" ? sorted.reverse() : sorted;
}

// Tom's method: mL of a 10% KMBS/SO2 stock solution needed to raise free SO2 from
// current to target, for a given volume in gallons. 172g KMBS per 1000mL stock solution,
// KMBS is 57.6% available SO2 by weight.
function calcSO2StockML(volumeGal, currentMgL, targetMgL) {
  const vol = parseFloat(volumeGal), cur = parseFloat(currentMgL), tgt = parseFloat(targetMgL);
  if (isNaN(vol) || isNaN(cur) || isNaN(tgt) || vol <= 0 || tgt <= cur) return null;
  const liters = vol * 3.78541;
  const mL = ((tgt - cur) * liters) / 99.072;
  return { liters, mL, vol, cur, tgt };
}

const WORKORDER_SORT_OPTIONS = [
  { value: "orderNumber", label: "Work Order #" },
  { value: "dateAssigned", label: "Date Assigned" },
  { value: "date", label: "Due Date" },
  { value: "priority", label: "Priority" },
  { value: "task", label: "Task" },
  { value: "taskType", label: "Task Type" },
  { value: "assignedTo", label: "Assigned To" },
  { value: "status", label: "Status" },
];
const FERMENT_SORT_OPTIONS = [
  { value: "tankId", label: "Lot" },
  { value: "variety", label: "Variety" },
  { value: "vessel", label: "Vessel" },
  { value: "startDate", label: "Start Date" },
  { value: "status", label: "Status" },
];

// Resize + compress a captured/uploaded photo down to a small JPEG data URL
function compressImage(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const maxDim = 1000;
        let { width, height } = img;
        if (width > maxDim || height > maxDim) {
          if (width > height) {
            height = Math.round((height * maxDim) / width);
            width = maxDim;
          } else {
            width = Math.round((width * maxDim) / height);
            height = maxDim;
          }
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", 0.6));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

const PRIORITY_STYLES = {
  High: "bg-rose-100 text-rose-800",
  Medium: "bg-amber-100 text-amber-800",
  Low: "bg-stone-200 text-stone-700",
};

// Color coding used on the master Calendar tab to tell sections apart
const CATEGORY_META = {
  "Work Orders (Open)": { word: "To-Do", pill: "bg-indigo-50 text-indigo-700" },
  "Work Orders (Closed)": { word: "Order Done", pill: "bg-ink-50 text-ink-800" },
  "Harvest": { word: "Harvest", pill: "bg-amber-50 text-amber-700" },
  "Fruit Analysis": { word: "Fruit Check", pill: "bg-lime-50 text-lime-700" },
  "Fermentation": { word: "Ferment", pill: "bg-rose-50 text-rose-700" },
  "Fermentation Complete": { word: "Ferm Done", pill: "bg-teal-50 text-teal-700" },
  "Malolactic": { word: "ML Done", pill: "bg-purple-50 text-purple-700" },
  "Barrel Filled": { word: "Filled", pill: "bg-sky-50 text-sky-700" },
  "Barrel Emptied": { word: "Emptied", pill: "bg-stone-100 text-stone-600" },
  "Barrel Retired": { word: "Retired", pill: "bg-orange-50 text-orange-700" },
  "Barrel Sold": { word: "Sold", pill: "bg-stone-100 text-stone-600" },
  "Lab Check": { word: "Lab Check", pill: "bg-cyan-50 text-cyan-700" },
  "Bottling": { word: "Bottled", pill: "bg-violet-50 text-violet-700" },
  "Tastings": { word: "Tasted", pill: "bg-pink-50 text-pink-700" },
  "Weather": { word: "Weather", pill: "bg-blue-50 text-blue-700" },
  "Phenology": { word: "Phenology", pill: "bg-lime-50 text-lime-700" },
  "Pest Alert": { word: "Pest Alert", pill: "bg-orange-50 text-orange-700" },
  "Disease Alert": { word: "Disease Alert", pill: "bg-red-50 text-red-700" },
};

// Flattens every section's dated activity into a single { isoDate: [events] } map
function buildCalendarEvents(data) {
  const map = {};
  const add = (date, section, title, detail) => {
    if (!date) return;
    if (!map[date]) map[date] = [];
    map[date].push({ section, title, detail });
  };

  data.workorders.forEach((o) => {
    const typeLabel = o.taskType ? (o.taskType === "Additions" && o.additionType ? `${o.taskType} (${o.additionType})` : o.taskType) : "";
    const title = `${o.orderNumber ? formatOrderNumber(o.orderNumber) + " — " : ""}${o.task || "Untitled task"}`;
    if (o.status === "Complete" && o.dateCompleted) {
      // Closed work orders live on the date they were completed, not the original due date
      add(o.dateCompleted, "Work Orders (Closed)", title, `Completed${typeLabel ? " · " + typeLabel : ""}${o.assignedTo ? " · " + o.assignedTo : ""}`);
    } else if (o.date) {
      // Open work orders show on the calendar on the date they were assigned
      add(o.date, "Work Orders (Open)", title, `Assigned${typeLabel ? " · " + typeLabel : ""}${o.assignedTo ? " · " + o.assignedTo : ""}${o.priority ? " · " + o.priority : ""}`);
    }
  });

  data.harvest.forEach((row) => {
    const weightLabel = row.netTons
      ? `${row.netTons} net tons`
      : row.tons
      ? `${row.tons} gross tons`
      : "";
    add(
      row.date,
      "Harvest",
      [row.variety, row.clone ? `Clone ${row.clone}` : "", row.block].filter(Boolean).join(" — ") || "Harvest entry",
      `${weightLabel}${row.weighMaster ? " · " + row.weighMaster : ""}`
    );
  });

  (data.fruitAnalysis || []).forEach((row) => {
    add(
      row.date,
      "Fruit Analysis",
      row.block || "Fruit sample",
      `${row.brix ? row.brix + "° Brix" : ""}${row.ph ? " · pH " + row.ph : ""}${row.ta ? " · TA " + row.ta : ""}`
    );
  });

  data.ferment.forEach((lot) => {
    if (lot.startDate) add(lot.startDate, "Fermentation", `${lot.tankId || "Tank"} started`, lot.variety || "");
    lot.readings.forEach((r) => {
      const workLabel = Array.isArray(r.workDone) && r.workDone.length > 0 ? r.workDone.join(", ") : "Reading logged";
      add(r.date, "Fermentation", `${lot.tankId || "Tank"} — ${workLabel}`, `${r.brix ? r.brix + "° Brix" : ""}${r.temp ? " · " + r.temp + "°F" : ""}`);
    });
    if (lot.dateCompleted) add(lot.dateCompleted, "Fermentation Complete", `${lot.tankId || "Tank"} — fermentation complete`, lot.variety || "");
    if (lot.mlCompleteDate) add(lot.mlCompleteDate, "Malolactic", `${lot.tankId || "Tank"} — ML complete`, lot.mlNotes || "");
  });

  (data.barrels || []).forEach((b) => {
    (b.fills || []).forEach((f) => {
      if (f.fillDate) add(f.fillDate, "Barrel Filled", `${b.barrelNumber} filled`, summarizeFillComponents(f, data.ferment));
      if (f.emptyDate) add(f.emptyDate, "Barrel Emptied", `${b.barrelNumber} emptied`, "");
    });
    (b.labChecks || []).forEach((c) => {
      add(c.date, "Lab Check", `${b.barrelNumber} — lab check`, `${c.freeSO2 ? "Free SO2 " + c.freeSO2 : ""}${c.totalSO2 ? " · Total SO2 " + c.totalSO2 : ""}${c.va ? " · VA " + c.va : ""}`);
    });
    if (b.retiredDate) add(b.retiredDate, "Barrel Retired", `${b.barrelNumber} retired`, b.retiredReason || "");
    if (b.soldDate) add(b.soldDate, "Barrel Sold", `${b.barrelNumber} sold`, b.soldTo ? `To ${b.soldTo}` : "");
  });

  (data.bottling || []).forEach((row) => {
    add(row.date, "Bottling", row.wineName || "Bottling run", `${row.cases ? row.cases + " cases" : ""}${row.bottleSize ? " · " + row.bottleSize : ""}${row.vintage ? " · " + row.vintage : ""}`);
  });

  (data.tastings || []).forEach((t) => {
    const barrel = (data.barrels || []).find((b) => b.id === t.barrelId);
    add(t.date, "Tastings", `${barrel?.barrelNumber || "Barrel"} tasted`, t.notes || "");
  });

  (data.weatherLogs || []).forEach((w) => {
    const high = w.high != null ? `${Math.round(w.high)}°` : "—";
    const low = w.low != null ? `${Math.round(w.low)}°` : "—";
    add(
      w.date,
      "Weather",
      `${high} / ${low}${w.conditionLabel ? " — " + w.conditionLabel : ""}`,
      `${w.humidity != null ? "Humidity " + w.humidity + "%" : ""}${w.windMph != null ? " · Wind " + Math.round(w.windMph) + " mph" : ""}${w.gddTotal != null ? " · GDD " + w.gddTotal : ""}`
    );
  });

  (data.vineHealth || []).forEach((v) => {
    if (v.observationType === "Phenology Stage" && v.phenologyStage) {
      add(v.date, "Phenology", `${v.phenologyStage}${v.block ? " — " + v.block : ""}`, v.notes || "");
    } else if (v.observationType === "Pest Pressure") {
      add(v.date, "Pest Alert", `${v.pestType || "Pest"}${v.block ? " — " + v.block : ""}${v.severity ? " (" + v.severity + ")" : ""}`, v.notes || "");
    } else if (v.observationType === "Disease Alert") {
      add(v.date, "Disease Alert", `${v.diseaseType || "Disease"}${v.block ? " — " + v.block : ""}${v.severity ? " (" + v.severity + ")" : ""}`, v.notes || "");
    }
  });

  return map;
}

// ---------- Multi-select lot picker with the ability to type in a brand-new lot ----------
// ---------- Multi-select lot picker: click existing lots to toggle them, or type a name that
// doesn't match anything yet to add a brand-new lot on the fly. ----------
function LotsPickerField({ value, onChange, fermentLots, lotNamesList, onRegisterLotName }) {
  const [search, setSearch] = useState("");
  const selected = Array.isArray(value) ? value : [];
  // Prefer the managed master list; fall back to Fermentation lot names if it's not passed in
  const suggestions = lotNamesList && lotNamesList.length > 0 ? lotNamesList : (fermentLots || []).map((l) => l.tankId).filter(Boolean);

  const toggleLot = (name) => {
    onChange(selected.includes(name) ? selected.filter((l) => l !== name) : [...selected, name]);
  };

  const filtered = search ? suggestions.filter((name) => name.toLowerCase().includes(search.toLowerCase())) : suggestions;
  const exactMatch = suggestions.some((name) => name.toLowerCase() === search.trim().toLowerCase());

  const addNewLot = () => {
    const trimmed = search.trim();
    if (!trimmed || selected.includes(trimmed)) return;
    onChange([...selected, trimmed]);
    if (onRegisterLotName) onRegisterLotName(trimmed);
    setSearch("");
  };

  return (
    <div>
      <input
        type="text"
        placeholder="Search lots, or type a new lot name"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && search.trim() && !exactMatch) {
            e.preventDefault();
            addNewLot();
          }
        }}
        className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm mb-2 focus:outline-none focus:ring-2 focus:ring-ink-800"
      />
      <div className="flex flex-wrap gap-1.5 max-h-32 overflow-y-auto border border-stone-200 rounded-md p-2">
        {filtered.length === 0 && !search ? (
          <span className="font-body text-xs text-stone-400">No lots yet — type a name above to add one.</span>
        ) : (
          filtered.slice(0, 60).map((name) => {
            const isSelected = selected.includes(name);
            return (
              <button
                key={name}
                type="button"
                onClick={() => toggleLot(name)}
                className={`font-body text-xs px-2 py-1 rounded-full border transition-colors ${
                  isSelected ? "bg-ink-800 text-white border-ink-800" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                }`}
              >
                {name}
              </button>
            );
          })
        )}
        {search.trim() && !exactMatch && (
          <button
            type="button"
            onClick={addNewLot}
            className="font-body flex items-center gap-1 text-xs px-2 py-1 rounded-full border border-dashed border-ink-400 text-ink-800 hover:bg-ink-50"
          >
            <Plus size={11} /> Add "{search.trim()}"
          </button>
        )}
      </div>
      {selected.length > 0 && (
        <div className="flex flex-wrap gap-1.5 mt-2">
          {selected.map((lot) => (
            <span key={lot} className="font-body flex items-center gap-1 text-xs bg-ink-50 text-ink-800 px-2 py-1 rounded-full">
              {lot}
              <button type="button" onClick={() => toggleLot(lot)} className="hover:text-red-700">
                <X size={12} />
              </button>
            </span>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- Multi-select barrel picker, searchable since there can be many barrels ----------
function BarrelsPickerField({ value, onChange, barrelsList }) {
  const [search, setSearch] = useState("");
  const selected = Array.isArray(value) ? value : [];

  const toggleBarrel = (id) => {
    onChange(selected.includes(id) ? selected.filter((b) => b !== id) : [...selected, id]);
  };

  const filtered = search
    ? barrelsList.filter((b) => b.barrelNumber.toLowerCase().includes(search.toLowerCase()))
    : barrelsList;
  const selectedBarrels = barrelsList.filter((b) => selected.includes(b.id));

  return (
    <div>
      <input
        type="text"
        placeholder="Search barrel #"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm mb-2 focus:outline-none focus:ring-2 focus:ring-ink-800"
      />
      <div className="flex flex-wrap gap-1.5 max-h-32 overflow-y-auto border border-stone-200 rounded-md p-2">
        {filtered.length === 0 ? (
          <span className="font-body text-xs text-stone-400">No barrels match.</span>
        ) : (
          filtered.slice(0, 60).map((b) => {
            const isSelected = selected.includes(b.id);
            return (
              <button
                key={b.id}
                type="button"
                onClick={() => toggleBarrel(b.id)}
                className={`font-body text-xs px-2 py-1 rounded-full border transition-colors ${
                  isSelected ? "bg-ink-800 text-white border-ink-800" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                }`}
              >
                {b.barrelNumber}
              </button>
            );
          })
        )}
      </div>
      {selectedBarrels.length > 0 && (
        <div className="flex flex-wrap gap-1.5 mt-2">
          {selectedBarrels.map((b) => (
            <span key={b.id} className="font-body flex items-center gap-1 text-xs bg-ink-50 text-ink-800 px-2 py-1 rounded-full">
              {b.barrelNumber}
              <button type="button" onClick={() => toggleBarrel(b.id)} className="hover:text-red-700">
                <X size={12} />
              </button>
            </span>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- Block/Vineyard picker with the ability to add and save a brand-new block ----------
// ---------- Generic single-select dropdown with the ability to add and save a brand-new option ----------
function AddableSelectField({ value, onChange, options, onAddOption, addLabel }) {
  const [adding, setAdding] = useState(false);
  const [newOption, setNewOption] = useState("");

  const handleSelectChange = (e) => {
    if (e.target.value === "__add_new__") {
      setAdding(true);
      setNewOption("");
    } else {
      onChange(e.target.value);
    }
  };

  const saveNewOption = () => {
    const trimmed = newOption.trim();
    if (!trimmed) return;
    if (onAddOption) onAddOption(trimmed);
    onChange(trimmed);
    setNewOption("");
    setAdding(false);
  };

  if (adding) {
    return (
      <div className="flex gap-2">
        <input
          type="text"
          autoFocus
          placeholder={addLabel}
          value={newOption}
          onChange={(e) => setNewOption(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              saveNewOption();
            }
          }}
          className="font-body flex-1 border border-stone-300 rounded-md px-2.5 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
        <button
          type="button"
          onClick={saveNewOption}
          className="font-body text-sm font-medium bg-ink-900 hover:bg-ink-800 text-white px-3 py-1.5 rounded-md shrink-0"
        >
          Save
        </button>
        <button
          type="button"
          onClick={() => { setAdding(false); setNewOption(""); }}
          className="font-body text-sm text-stone-500 hover:text-stone-700 px-2 shrink-0"
        >
          Cancel
        </button>
      </div>
    );
  }

  return (
    <select
      value={value}
      onChange={handleSelectChange}
      className="font-body w-full max-w-[200px] border border-stone-300 rounded-md px-2.5 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
    >
      <option value="">Select…</option>
      {options.map((o) => (
        <option key={o} value={o}>{o}</option>
      ))}
      <option value="__add_new__">+ Add new…</option>
    </select>
  );
}

function Field({ f, value, onChange, hideLabel, fermentLots, barrelsList, blocksList, onAddBlock, vesselTypesList, onAddVesselType, lotNamesList, onRegisterLotName, clonesList, onAddClone, sprayProgramsList, onAddSprayProgram, associatesList, onAddAssociate }) {
  return (
    <div className={f.type === "textarea" || f.type === "checkbox-group" || f.type === "lots-picker" || f.type === "barrels-picker" ? "sm:col-span-2" : ""}>
      {!hideLabel && <label className="font-body block text-xs font-medium text-stone-600 mb-1">{f.label}</label>}
      {f.type === "select" ? (
        <select
          value={value}
          onChange={(e) => onChange(e.target.value)}
          className="font-body w-full max-w-[200px] border border-stone-300 rounded-md px-2.5 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
        >
          <option value="">Select…</option>
          {f.options.map((opt) => (
            <option key={opt} value={opt}>{opt}</option>
          ))}
        </select>
      ) : f.type === "textarea" ? (
        <textarea
          value={value}
          onChange={(e) => onChange(e.target.value)}
          rows={2}
          className="font-body w-full border border-stone-300 rounded-md px-2.5 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
      ) : f.type === "checkbox-group" ? (
        <div className="flex flex-wrap gap-3 pt-1">
          {f.options.map((opt) => {
            const checked = Array.isArray(value) && value.includes(opt);
            return (
              <label key={opt} className="font-body flex items-center gap-1.5 text-sm text-stone-700">
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={() => {
                    const current = Array.isArray(value) ? value : [];
                    onChange(checked ? current.filter((v) => v !== opt) : [...current, opt]);
                  }}
                  className="rounded border-stone-300 text-ink-800 focus:ring-ink-800"
                />
                {opt}
              </label>
            );
          })}
        </div>
      ) : f.type === "lots-picker" ? (
        <LotsPickerField value={value} onChange={onChange} fermentLots={fermentLots || []} lotNamesList={lotNamesList} onRegisterLotName={onRegisterLotName} />
      ) : f.type === "barrels-picker" ? (
        <BarrelsPickerField value={value} onChange={onChange} barrelsList={barrelsList || []} />
      ) : f.type === "block-picker" ? (
        <AddableSelectField value={value} onChange={onChange} options={blocksList || []} onAddOption={onAddBlock} addLabel="New block / vineyard name" />
      ) : f.type === "vessel-picker" ? (
        <AddableSelectField value={value} onChange={onChange} options={vesselTypesList || []} onAddOption={onAddVesselType} addLabel="New vessel name" />
      ) : f.type === "clone-picker" ? (
        <AddableSelectField value={value} onChange={onChange} options={clonesList || []} onAddOption={onAddClone} addLabel="New clone" />
      ) : f.type === "spray-program-picker" ? (
        <AddableSelectField value={value} onChange={onChange} options={sprayProgramsList || []} onAddOption={onAddSprayProgram} addLabel="New spray program" />
      ) : f.type === "associate-picker" ? (
        <AddableSelectField value={value} onChange={onChange} options={associatesList || []} onAddOption={onAddAssociate} addLabel="New tasting associate" />
      ) : f.type === "photo" ? (
        <div>
          <input
            type="file"
            accept="image/*"
            capture="environment"
            onChange={async (e) => {
              const file = e.target.files?.[0];
              if (!file) return;
              try {
                const dataUrl = await compressImage(file);
                onChange(dataUrl);
              } catch {
                // ignore failed capture
              }
              e.target.value = "";
            }}
            className="font-body w-full text-xs text-stone-600 file:mr-2 file:py-2 file:px-3 file:rounded-md file:border-0 file:bg-ink-900 file:text-white file:text-xs file:font-medium hover:file:bg-ink-800"
          />
          {value && (
            <div className="mt-2 flex items-center gap-2">
              <a href={value} target="_blank" rel="noopener noreferrer">
                <img src={value} alt="Receipt preview" className="w-16 h-16 object-cover rounded border border-stone-300" />
              </a>
              <button type="button" onClick={() => onChange("")} className="font-body text-xs text-stone-400 hover:text-red-700">
                Remove
              </button>
            </div>
          )}
        </div>
      ) : f.type === "document" ? (
        <div>
          <input
            type="file"
            accept="image/*,application/pdf"
            onChange={async (e) => {
              const file = e.target.files?.[0];
              if (!file) return;
              try {
                if (file.type === "application/pdf") {
                  const dataUrl = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result);
                    reader.onerror = reject;
                    reader.readAsDataURL(file);
                  });
                  onChange(dataUrl);
                } else {
                  const dataUrl = await compressImage(file);
                  onChange(dataUrl);
                }
              } catch {
                // ignore failed upload
              }
              e.target.value = "";
            }}
            className="font-body w-full text-xs text-stone-600 file:mr-2 file:py-2 file:px-3 file:rounded-md file:border-0 file:bg-ink-900 file:text-white file:text-xs file:font-medium hover:file:bg-ink-800"
          />
          {value && (
            <div className="mt-2 flex items-center gap-2">
              {value.startsWith("data:application/pdf") ? (
                <a href={value} target="_blank" rel="noopener noreferrer" className="font-body text-xs text-ink-800 underline">
                  View uploaded PDF
                </a>
              ) : (
                <a href={value} target="_blank" rel="noopener noreferrer">
                  <img src={value} alt="Document preview" className="w-16 h-16 object-cover rounded border border-stone-300" />
                </a>
              )}
              <button type="button" onClick={() => onChange("")} className="font-body text-xs text-stone-400 hover:text-red-700">
                Remove
              </button>
            </div>
          )}
        </div>
      ) : (
        <input
          type={f.type}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          className={`font-body border border-stone-300 rounded-md px-2.5 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800 ${
            f.type === "number" || f.type === "date" ? "w-full max-w-[140px]" : "w-full"
          }`}
        />
      )}
    </div>
  );
}

// ---------- Work order checklist row ----------
function WorkOrderRow({ order, onToggle, onDelete, onDuplicate, onSaveAsTemplate, isEditing, editForm, onEditChange, onStartEdit, onSaveEdit, onCancelEdit, fermentLots, barrelsList, lotNamesList, onRegisterLotName, sprayProgramsList, onAddSprayProgram }) {
  const isComplete = order.status === "Complete";

  if (isEditing) {
    const visibleFields = visibleWorkOrderFields(workOrderFieldsForCategory(WORKORDER_FIELDS, editForm.category), editForm);
    return (
      <li className="px-4 py-3 bg-ink-50">
        {order.orderNumber && (
          <p className="font-body text-xs font-mono text-stone-400 mb-2">{formatOrderNumber(order.orderNumber)}</p>
        )}
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 mb-3">
          {visibleFields.map((f) => (
            <Field
              key={f.name}
              f={f}
              value={editForm[f.name]}
              onChange={(v) => onEditChange(f.name, v)}
              fermentLots={fermentLots}
              barrelsList={barrelsList}
              lotNamesList={lotNamesList}
              onRegisterLotName={onRegisterLotName}
              sprayProgramsList={sprayProgramsList}
              onAddSprayProgram={onAddSprayProgram}
            />
          ))}
        </div>
        <div className="flex items-center gap-3">
          <button onClick={onSaveEdit} className="font-body flex items-center gap-1 text-sm font-medium text-ink-800 hover:text-ink-900">
            <Check size={15} /> Save
          </button>
          <button onClick={onCancelEdit} className="font-body flex items-center gap-1 text-sm text-stone-500 hover:text-stone-700">
            <X size={15} /> Cancel
          </button>
        </div>
      </li>
    );
  }

  return (
    <li className="flex items-start gap-3 px-4 py-3 hover:bg-stone-50">
      <button
        onClick={() => onToggle(order.id)}
        className="mt-0.5 text-ink-900 hover:text-ink-700 shrink-0"
        title={isComplete ? "Mark as open" : "Mark complete"}
      >
        {isComplete ? <CheckSquare size={20} /> : <Circle size={20} />}
      </button>
      <div className="flex-1 min-w-0">
        <p className={`font-body text-sm ${isComplete ? "line-through text-stone-400" : "text-stone-900"}`}>
          {order.orderNumber && <span className="font-mono text-stone-400 mr-1.5">{formatOrderNumber(order.orderNumber)}</span>}
          {order.task || "Untitled task"}
        </p>
        <div className="flex flex-wrap items-center gap-2 mt-1">
          {order.taskType && (
            <span className="font-body text-xs px-1.5 py-0.5 rounded bg-ink-100 text-ink-800">
              {order.taskType}{order.taskType === "Additions" && order.additionType ? ` · ${order.additionType}` : ""}
            </span>
          )}
          {order.priority && (
            <span className={`font-body text-xs px-1.5 py-0.5 rounded ${PRIORITY_STYLES[order.priority] || "bg-stone-100 text-stone-600"}`}>
              {order.priority}
            </span>
          )}
          {order.assignedTo && <span className="font-body text-xs text-stone-500">{order.assignedTo}</span>}
          {order.dateAssigned && <span className="font-body text-xs text-stone-400">· assigned {order.dateAssigned}</span>}
          {order.date && <span className="font-body text-xs text-stone-400">· due {order.date}</span>}
          {isComplete && order.dateCompleted && (
            <span className="font-body text-xs text-green-700">· completed {order.dateCompleted}</span>
          )}
        </div>
        {order.calculations && (
          <p className="font-body text-xs text-stone-500 mt-1"><span className="font-medium">Calculations:</span> {order.calculations}</p>
        )}
        {order.directions && (
          <p className="font-body text-xs text-stone-500 mt-1"><span className="font-medium">Directions:</span> {order.directions}</p>
        )}
        {Array.isArray(order.lots) && order.lots.length > 0 && (
          <p className="font-body text-xs text-stone-500 mt-1">
            <span className="font-medium">Lots:</span> {order.lots.join(", ")}
          </p>
        )}
        {Array.isArray(order.barrels) && order.barrels.length > 0 && (
          <p className="font-body text-xs text-stone-500 mt-1">
            <span className="font-medium">Barrels:</span>{" "}
            {order.barrels
              .map((id) => (barrelsList || []).find((b) => b.id === id)?.barrelNumber)
              .filter(Boolean)
              .join(", ")}
          </p>
        )}
        {order.notes && <p className="font-body text-xs text-stone-400 mt-1">{order.notes}</p>}
      </div>
      {isComplete && (
        <button
          onClick={() => onToggle(order.id)}
          className="font-body flex items-center gap-1 text-xs font-medium text-ink-700 hover:text-ink-900 shrink-0"
          title="Reopen this work order"
        >
          <RotateCcw size={14} /> Reopen
        </button>
      )}
      <button onClick={() => onDuplicate(order)} className="text-stone-300 hover:text-ink-800 shrink-0" title="Duplicate for another day">
        <Copy size={14} />
      </button>
      <button onClick={() => onSaveAsTemplate(order)} className="text-stone-300 hover:text-ink-800 shrink-0" title="Save as template">
        <BookmarkPlus size={14} />
      </button>
      <button onClick={() => onStartEdit(order)} className="text-stone-300 hover:text-ink-800 shrink-0" title="Edit">
        <Pencil size={14} />
      </button>
      <button onClick={() => onDelete(order.id)} className="text-stone-300 hover:text-red-700 shrink-0" title="Delete">
        <Trash2 size={15} />
      </button>
    </li>
  );
}

// Groups completed work orders by the month they were closed, most recent first
function groupByMonth(orders) {
  const groups = {};
  orders.forEach((o) => {
    if (!o.dateCompleted) return;
    const key = o.dateCompleted.slice(0, 7); // YYYY-MM
    if (!groups[key]) groups[key] = [];
    groups[key].push(o);
  });
  return Object.entries(groups)
    .sort((a, b) => (a[0] < b[0] ? 1 : -1))
    .map(([key, groupOrders]) => {
      const [y, m] = key.split("-");
      const label = new Date(Number(y), Number(m) - 1, 1).toLocaleDateString(undefined, { month: "long", year: "numeric" });
      return { key, label, orders: [...groupOrders].sort((a, b) => (a.dateCompleted < b.dateCompleted ? 1 : -1)) };
    });
}

// Groups archived fermentation lots by vintage year (derived from start date), most recent first
function groupByVintage(lots) {
  const groups = {};
  lots.forEach((l) => {
    const key = l.startDate ? l.startDate.slice(0, 4) : "Unknown Vintage";
    if (!groups[key]) groups[key] = [];
    groups[key].push(l);
  });
  return Object.entries(groups)
    .sort((a, b) => (a[0] < b[0] ? 1 : -1))
    .map(([key, groupLots]) => ({
      key,
      label: key === "Unknown Vintage" ? key : `${key} Vintage`,
      lots: [...groupLots].sort((a, b) => (a.tankId || "").localeCompare(b.tankId || "")),
    }));
}

// ---------- One row in the Work Order Templates library ----------
function TemplateRow({ template, onUse, onDelete }) {
  return (
    <li className="flex items-start gap-3 px-4 py-3 hover:bg-stone-50">
      <Library size={16} className="text-stone-300 mt-0.5 shrink-0" />
      <div className="flex-1 min-w-0">
        <p className="font-body text-sm text-stone-900">{template.task || "Untitled template"}</p>
        <div className="flex flex-wrap items-center gap-2 mt-1">
          {template.taskType && (
            <span className="font-body text-xs px-1.5 py-0.5 rounded bg-ink-100 text-ink-800">
              {template.taskType}{template.taskType === "Additions" && template.additionType ? ` · ${template.additionType}` : ""}
            </span>
          )}
          {template.priority && (
            <span className={`font-body text-xs px-1.5 py-0.5 rounded ${PRIORITY_STYLES[template.priority] || "bg-stone-100 text-stone-600"}`}>
              {template.priority}
            </span>
          )}
          {template.assignedTo && <span className="font-body text-xs text-stone-500">{template.assignedTo}</span>}
        </div>
        {template.notes && <p className="font-body text-xs text-stone-400 mt-1">{template.notes}</p>}
      </div>
      <button
        onClick={() => onUse(template)}
        className="font-body flex items-center gap-1 text-xs font-medium text-ink-700 hover:text-ink-900 shrink-0"
        title="Create a work order from this template"
      >
        <Plus size={14} /> Use
      </button>
      <button onClick={() => onDelete(template.id)} className="text-stone-300 hover:text-red-700 shrink-0" title="Delete template">
        <Trash2 size={14} />
      </button>
    </li>
  );
}


function ArchiveGroup({ label, orders, onToggle, onDelete, onDuplicate, onSaveAsTemplate, editingWorkOrderId, editWorkOrderForm, onEditChange, onStartEdit, onSaveEdit, onCancelEdit, fermentLots, barrelsList, lotNamesList, onRegisterLotName, sprayProgramsList, onAddSprayProgram }) {
  const [expanded, setExpanded] = useState(false);
  return (
    <div className="border-b border-stone-100 last:border-b-0">
      <button onClick={() => setExpanded((v) => !v)} className="w-full flex items-center justify-between px-4 py-3 hover:bg-stone-50">
        <span className="font-body text-sm font-medium text-stone-700">{label}</span>
        <span className="flex items-center gap-2">
          <span className="font-body text-xs text-stone-400">{orders.length} closed</span>
          {expanded ? <ChevronUp size={16} className="text-stone-400" /> : <ChevronDown size={16} className="text-stone-400" />}
        </span>
      </button>
      {expanded && (
        <ul className="divide-y divide-stone-100">
          {orders.map((o) => (
            <WorkOrderRow
              key={o.id}
              order={o}
              onToggle={onToggle}
              onDelete={onDelete}
              onDuplicate={onDuplicate}
              onSaveAsTemplate={onSaveAsTemplate}
              isEditing={editingWorkOrderId === o.id}
              editForm={editWorkOrderForm}
              onEditChange={onEditChange}
              onStartEdit={onStartEdit}
              onSaveEdit={onSaveEdit}
              onCancelEdit={onCancelEdit}
              fermentLots={fermentLots}
              barrelsList={barrelsList}
              lotNamesList={lotNamesList}
              onRegisterLotName={onRegisterLotName}
              sprayProgramsList={sprayProgramsList}
              onAddSprayProgram={onAddSprayProgram}
            />
          ))}
        </ul>
      )}
    </div>
  );
}


function CompletedCalendar({ orders }) {
  const [cursor, setCursor] = useState(() => {
    const now = new Date();
    return { year: now.getFullYear(), month: now.getMonth() };
  });
  const [selectedDate, setSelectedDate] = useState(null);

  const byDate = useMemo(() => {
    const map = {};
    orders.forEach((o) => {
      if (!o.dateCompleted) return;
      map[o.dateCompleted] = map[o.dateCompleted] || [];
      map[o.dateCompleted].push(o);
    });
    return map;
  }, [orders]);

  const { year, month } = cursor;
  const firstWeekday = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < firstWeekday; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);
  while (cells.length % 7 !== 0) cells.push(null);

  const monthLabel = new Date(year, month, 1).toLocaleDateString(undefined, { month: "long", year: "numeric" });

  const changeMonth = (delta) => {
    setSelectedDate(null);
    setCursor((prev) => {
      let m = prev.month + delta;
      let y = prev.year;
      if (m < 0) { m = 11; y -= 1; }
      if (m > 11) { m = 0; y += 1; }
      return { year: y, month: m };
    });
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
      <div className="flex items-center justify-between mb-3">
        <h2 className="font-brand text-lg text-ink-950">Completed — by Date</h2>
        <div className="flex items-center gap-2">
          <button onClick={() => changeMonth(-1)} className="text-stone-400 hover:text-ink-800">
            <ChevronLeft size={18} />
          </button>
          <span className="font-body text-sm text-stone-600 w-32 text-center">{monthLabel}</span>
          <button onClick={() => changeMonth(1)} className="text-stone-400 hover:text-ink-800">
            <ChevronRight size={18} />
          </button>
        </div>
      </div>

      <div className="grid grid-cols-7 gap-1 text-center mb-1">
        {["S", "M", "T", "W", "T", "F", "S"].map((d, i) => (
          <div key={i} className="font-body text-xs text-stone-400 py-1">{d}</div>
        ))}
      </div>
      <div className="grid grid-cols-7 gap-1">
        {cells.map((day, idx) => {
          if (day === null) return <div key={idx} />;
          const iso = isoFor(year, month, day);
          const count = byDate[iso]?.length || 0;
          const isSelected = selectedDate === iso;
          return (
            <button
              key={idx}
              onClick={() => count > 0 && setSelectedDate(isSelected ? null : iso)}
              className={`font-body aspect-square rounded-md text-xs flex flex-col items-center justify-center gap-0.5 border transition-colors ${
                isSelected
                  ? "border-ink-800 bg-ink-50"
                  : count > 0
                  ? "border-stone-200 hover:border-ink-300 cursor-pointer"
                  : "border-transparent text-stone-300"
              }`}
            >
              <span>{day}</span>
              {count > 0 && <span className="w-1.5 h-1.5 rounded-full bg-ink-800" />}
            </button>
          );
        })}
      </div>

      {selectedDate && (
        <div className="mt-4 border-t border-stone-100 pt-3">
          <p className="font-body text-xs font-semibold text-stone-600 mb-2">Completed on {selectedDate}</p>
          <ul className="space-y-2">
            {byDate[selectedDate].map((o) => (
              <li key={o.id} className="font-body text-sm bg-stone-50 rounded-md px-3 py-2">
                <p className="text-stone-800">{o.task}</p>
                <p className="text-xs text-stone-500 mt-0.5">
                  {o.assignedTo || ""}
                </p>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}

// ---------- Master Calendar: every section's activity, in one place, by date ----------
function MasterCalendar({ data }) {
  const [cursor, setCursor] = useState(() => {
    const now = new Date();
    return { year: now.getFullYear(), month: now.getMonth() };
  });
  const [selectedDate, setSelectedDate] = useState(null);

  const eventsByDate = useMemo(() => buildCalendarEvents(data), [data]);

  const { year, month } = cursor;
  const firstWeekday = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < firstWeekday; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);
  while (cells.length % 7 !== 0) cells.push(null);

  const monthLabel = new Date(year, month, 1).toLocaleDateString(undefined, { month: "long", year: "numeric" });

  const changeMonth = (delta) => {
    setSelectedDate(null);
    setCursor((prev) => {
      let m = prev.month + delta;
      let y = prev.year;
      if (m < 0) { m = 11; y -= 1; }
      if (m > 11) { m = 0; y += 1; }
      return { year: y, month: m };
    });
  };

  const selectedEvents = selectedDate ? eventsByDate[selectedDate] || [] : [];

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
        <h2 className="font-brand text-lg text-ink-950">Activity Calendar</h2>
        <div className="flex items-center gap-2">
          <button onClick={() => changeMonth(-1)} className="text-stone-400 hover:text-ink-800">
            <ChevronLeft size={18} />
          </button>
          <span className="font-body text-sm text-stone-600 w-32 text-center">{monthLabel}</span>
          <button onClick={() => changeMonth(1)} className="text-stone-400 hover:text-ink-800">
            <ChevronRight size={18} />
          </button>
        </div>
      </div>

      <p className="font-body text-xs text-stone-500 mb-4">
        Every task, entry, and reading across every tab, shown on the day it was assigned, logged, or completed.
      </p>

      <div className="grid grid-cols-7 gap-1.5 text-center mb-1.5">
        {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((d, i) => (
          <div key={i} className="font-body text-xs font-medium text-stone-400 py-1">{d}</div>
        ))}
      </div>
      <div className="grid grid-cols-7 gap-1.5">
        {cells.map((day, idx) => {
          if (day === null) return <div key={idx} />;
          const iso = isoFor(year, month, day);
          const events = eventsByDate[iso] || [];
          const sections = [...new Set(events.map((e) => e.section))];
          const isSelected = selectedDate === iso;
          const isToday = iso === todayISO();
          const shown = sections.slice(0, 2);
          const extra = sections.length - shown.length;
          return (
            <button
              key={idx}
              onClick={() => events.length > 0 && setSelectedDate(isSelected ? null : iso)}
              className={`font-body min-h-[64px] rounded-md p-1 flex flex-col items-start gap-0.5 border transition-colors ${
                isSelected
                  ? "border-ink-800 bg-ink-50"
                  : events.length > 0
                  ? "border-stone-200 hover:border-ink-300 cursor-pointer bg-white"
                  : "border-transparent"
              }`}
            >
              <span className={`text-xs px-1 rounded-full ${isToday ? "bg-ink-900 text-white font-semibold" : "text-stone-500"}`}>
                {day}
              </span>
              {shown.length > 0 && (
                <div className="flex flex-col gap-0.5 w-full">
                  {shown.map((s) => (
                    <span
                      key={s}
                      className={`font-body text-[10px] leading-tight px-1 py-0.5 rounded truncate w-full text-left ${CATEGORY_META[s]?.pill || "bg-stone-100 text-stone-600"}`}
                    >
                      {CATEGORY_META[s]?.word || s}
                    </span>
                  ))}
                  {extra > 0 && (
                    <span className="font-body text-[10px] text-stone-400 px-1">+{extra} more</span>
                  )}
                </div>
              )}
            </button>
          );
        })}
      </div>

      {selectedDate && (
        <div className="mt-4 border-t border-stone-100 pt-3">
          <p className="font-body text-xs font-semibold text-stone-600 mb-2">
            {selectedEvents.length} item{selectedEvents.length === 1 ? "" : "s"} on {selectedDate}
          </p>
          <ul className="space-y-2">
            {selectedEvents.map((e, i) => (
              <li key={i} className="font-body text-sm bg-stone-50 rounded-md px-3 py-2">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className={`text-xs font-semibold px-1.5 py-0.5 rounded ${CATEGORY_META[e.section]?.pill || "bg-stone-100 text-stone-600"}`}>
                    {CATEGORY_META[e.section]?.word || e.section}
                  </span>
                  <span className="text-stone-800">{e.title}</span>
                </div>
                {e.detail && <p className="text-xs text-stone-500 mt-1">{e.detail}</p>}
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}

// ---------- One card per active/complete ferment, with its own daily-reading form ----------
// Flags a likely stalled fermentation: the two most recent readings show almost no Brix drop
// while the lot is still meaningfully above dry. This is the same "no change over 2+ days"
// rule of thumb professional winemakers use to catch stuck fermentations early.
function detectStall(readingsAsc) {
  const withBrix = readingsAsc.filter((r) => r.brix !== "" && r.brix != null && !isNaN(parseFloat(r.brix)));
  if (withBrix.length < 2) return null;
  const last = withBrix[withBrix.length - 1];
  const prev = withBrix[withBrix.length - 2];
  const lastBrix = parseFloat(last.brix);
  const prevBrix = parseFloat(prev.brix);
  const drop = prevBrix - lastBrix;
  if (lastBrix > 0.5 && drop <= 0.3) {
    return { prevDate: prev.date, lastDate: last.date, drop };
  }
  return null;
}

// ---------- Visual Brix + Temp fermentation curve for one lot ----------
function FermentationChart({ readings }) {
  const chartData = readings
    .filter((r) => (r.brix !== "" && r.brix != null) || (r.temp !== "" && r.temp != null))
    .map((r) => ({
      date: r.date ? r.date.slice(5) : "",
      fullDate: r.date,
      brix: r.brix !== "" && r.brix != null ? parseFloat(r.brix) : null,
      temp: r.temp !== "" && r.temp != null ? parseFloat(r.temp) : null,
    }));

  if (chartData.length < 2) return null;

  return (
    <div className="bg-white border border-stone-200 rounded-md p-3 mb-4">
      <p className="font-body text-xs font-semibold text-stone-600 mb-2">Fermentation Curve</p>
      <ResponsiveContainer width="100%" height={200}>
        <LineChart data={chartData} margin={{ top: 5, right: 8, left: -18, bottom: 0 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#e7e5e4" />
          <XAxis dataKey="date" tick={{ fontSize: 11, fill: "#78716c" }} />
          <YAxis yAxisId="brix" tick={{ fontSize: 11, fill: "#065f46" }} width={34} />
          <YAxis yAxisId="temp" orientation="right" tick={{ fontSize: 11, fill: "#b45309" }} width={34} />
          <Tooltip
            contentStyle={{ fontSize: 12, fontFamily: "Inter, sans-serif", borderRadius: 8, border: "1px solid #e7e5e4" }}
            labelFormatter={(label, payload) => payload?.[0]?.payload?.fullDate || label}
          />
          <Legend wrapperStyle={{ fontSize: 11 }} />
          <Line yAxisId="brix" type="monotone" dataKey="brix" name="Brix" stroke="#065f46" strokeWidth={2} dot={{ r: 3 }} connectNulls />
          <Line yAxisId="temp" type="monotone" dataKey="temp" name="Temp (°F)" stroke="#b45309" strokeWidth={2} dot={{ r: 3 }} connectNulls />
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
}

// A fixed color per active lot, cycled if there are more lots than colors
// A curated palette rooted in the winery's own brand tones (deep emerald, gold, wine-red,
// moss, terracotta) rather than generic dashboard colors — cycled if there are more active
// lots than colors.
const FERMENT_CHART_COLORS = ["#065f46", "#b45309", "#9f1239", "#0f766e", "#7c2d12", "#4d7c0f", "#1e3a5f", "#a16207"];

// Custom legend: one row per lot (color swatch + name), with a single caption explaining the
// solid/dashed convention — reads far cleaner than Recharts' default legend, which would show
// two separate, near-duplicate entries per lot ("Lot Brix" / "Lot Temp"). Each lot is clickable
// to show/hide it on the chart, since a handful of active lots can get busy on one timeline.
function FermentChartLegend({ lots, colorFor, hiddenLots, onToggleLot, onShowAll }) {
  const anyHidden = hiddenLots.size > 0;
  return (
    <div className="flex flex-wrap items-center gap-x-1 gap-y-1.5 mt-3 pt-3 border-t border-stone-100">
      {lots.map((lot) => {
        const isHidden = hiddenLots.has(lot.id);
        return (
          <button
            key={lot.id}
            type="button"
            onClick={() => onToggleLot(lot.id)}
            title={isHidden ? "Click to show on chart" : "Click to hide from chart"}
            className={`font-body flex items-center gap-1.5 text-xs px-1.5 py-0.5 rounded hover:bg-stone-50 transition-opacity ${isHidden ? "opacity-40" : ""}`}
          >
            <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: colorFor(lot) }} />
            <span className={isHidden ? "text-stone-400" : "text-stone-600"}>{lot.tankId || "Untitled"}</span>
          </button>
        );
      })}
      {anyHidden && (
        <button type="button" onClick={onShowAll} className="font-body text-xs text-ink-800 hover:text-ink-900 underline px-1.5">
          Show all
        </button>
      )}
      <span className="font-body flex items-center gap-3 text-xs text-stone-400 ml-auto">
        <span className="flex items-center gap-1">
          <span className="w-3 h-0.5 bg-stone-400 rounded-full" /> Brix
        </span>
        <span className="flex items-center gap-1">
          <svg width="12" height="2" className="shrink-0">
            <line x1="0" y1="1" x2="12" y2="1" stroke="#a8a29e" strokeWidth="2" strokeDasharray="3 2" />
          </svg>
          Temp
        </span>
      </span>
    </div>
  );
}

// ---------- Dashboard chart: every active ferment's Brix (solid) and Temp (dashed) on one
// timeline, so you can compare all active lots at a glance. Reads straight from live data, so it
// updates automatically whenever a reading is logged (Quick Log or Detailed View, either one). ----------
function MultiLotFermentChart({ lots }) {
  const [hiddenLots, setHiddenLots] = useState(new Set());
  const toggleLot = (id) =>
    setHiddenLots((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  const showAllLots = () => setHiddenLots(new Set());

  const colorFor = (lot) => {
    const i = lots.findIndex((l) => l.id === lot.id);
    return FERMENT_CHART_COLORS[i % FERMENT_CHART_COLORS.length];
  };

  if (lots.length === 0) {
    return (
      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h3 className="font-brand text-lg text-ink-950 mb-1">Active Fermentations</h3>
        <p className="font-body text-sm text-stone-400">No active ferments right now.</p>
      </div>
    );
  }

  // One row per distinct date across every active lot, each lot contributing its own
  // brix/temp columns (using that lot's latest reading that day, if there's more than one).
  const dateSet = new Set();
  const perLotByDate = {};
  lots.forEach((lot) => {
    const sorted = [...lot.readings].sort((a, b) => (a.date < b.date ? -1 : 1));
    const byDate = {};
    sorted.forEach((r) => {
      byDate[r.date] = {
        brix: r.brix !== "" && r.brix != null ? parseFloat(r.brix) : null,
        temp: r.temp !== "" && r.temp != null ? parseFloat(r.temp) : null,
      };
      dateSet.add(r.date);
    });
    perLotByDate[lot.id] = byDate;
  });
  const dates = [...dateSet].sort();
  const chartData = dates.map((d) => {
    const row = { date: d.slice(5), fullDate: d };
    lots.forEach((lot) => {
      const entry = perLotByDate[lot.id][d];
      row[`${lot.id}_brix`] = entry ? entry.brix : null;
      row[`${lot.id}_temp`] = entry ? entry.temp : null;
    });
    return row;
  });

  const visibleLots = lots.filter((lot) => !hiddenLots.has(lot.id));

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
      <div className="flex items-center justify-between mb-1">
        <h3 className="font-brand text-lg text-ink-950">Active Fermentations</h3>
        <span className="font-body text-xs bg-ink-50 text-ink-800 px-2 py-0.5 rounded-full">
          {lots.length} active
        </span>
      </div>
      <p className="font-body text-xs text-stone-500 mb-3">
        Updates automatically as readings are logged, from Quick Log or the detailed view. Click a lot below to show/hide it.
      </p>
      {chartData.length < 2 ? (
        <p className="font-body text-sm text-stone-400">Not enough readings logged yet to plot a curve.</p>
      ) : (
        <>
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={chartData} margin={{ top: 8, right: 8, left: -18, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f0ee" vertical={false} />
              <XAxis dataKey="date" tick={{ fontSize: 11, fill: "#a8a29e", fontFamily: "Inter, sans-serif" }} axisLine={{ stroke: "#e7e5e4" }} tickLine={false} />
              <YAxis
                yAxisId="brix"
                tick={{ fontSize: 11, fill: "#a8a29e", fontFamily: "Inter, sans-serif" }}
                width={34}
                axisLine={false}
                tickLine={false}
                label={{ value: "Brix", angle: -90, position: "insideLeft", fontSize: 11, fill: "#a8a29e" }}
              />
              <YAxis
                yAxisId="temp"
                orientation="right"
                tick={{ fontSize: 11, fill: "#a8a29e", fontFamily: "Inter, sans-serif" }}
                width={34}
                axisLine={false}
                tickLine={false}
                label={{ value: "°F", angle: 90, position: "insideRight", fontSize: 11, fill: "#a8a29e" }}
              />
              <Tooltip
                contentStyle={{ fontSize: 12, fontFamily: "Inter, sans-serif", borderRadius: 10, border: "1px solid #e7e5e4", boxShadow: "0 4px 12px rgba(0,0,0,0.06)" }}
                labelFormatter={(label, payload) => payload?.[0]?.payload?.fullDate || label}
              />
              {visibleLots.map((lot) => {
                const color = colorFor(lot);
                const label = lot.tankId || "Untitled";
                return (
                  <Fragment key={lot.id}>
                    <Line
                      yAxisId="brix"
                      type="monotone"
                      dataKey={`${lot.id}_brix`}
                      name={`${label} Brix`}
                      stroke={color}
                      strokeWidth={2.25}
                      dot={{ r: 2.5, strokeWidth: 0, fill: color }}
                      activeDot={{ r: 4 }}
                      connectNulls
                    />
                    <Line
                      yAxisId="temp"
                      type="monotone"
                      dataKey={`${lot.id}_temp`}
                      name={`${label} Temp`}
                      stroke={color}
                      strokeOpacity={0.55}
                      strokeWidth={1.5}
                      strokeDasharray="4 3"
                      dot={false}
                      activeDot={{ r: 3 }}
                      connectNulls
                    />
                  </Fragment>
                );
              })}
            </LineChart>
          </ResponsiveContainer>
          <FermentChartLegend lots={lots} colorFor={colorFor} hiddenLots={hiddenLots} onToggleLot={toggleLot} onShowAll={showAllLots} />
        </>
      )}
    </div>
  );
}

// ---------- Dashboard chart: total tonnage by grape variety, net tons where recorded, gross
// otherwise. Reads straight from live Harvest data, so it updates automatically as entries are
// logged, edited, or deleted. ----------
function TonnageByVarietyChart({ harvest }) {
  const totals = {};
  harvest.forEach((h) => {
    const variety = h.variety || "Unspecified";
    const net = h.netTons !== "" && h.netTons != null ? parseFloat(h.netTons) : null;
    const gross = h.tons !== "" && h.tons != null ? parseFloat(h.tons) : null;
    const tons = net != null && !isNaN(net) ? net : gross;
    if (tons == null || isNaN(tons)) return;
    totals[variety] = (totals[variety] || 0) + tons;
  });
  const chartData = Object.entries(totals)
    .map(([variety, tons]) => ({ variety, tons: Math.round(tons * 100) / 100 }))
    .sort((a, b) => b.tons - a.tons);
  const grandTotal = chartData.reduce((sum, d) => sum + d.tons, 0);

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
      <div className="flex items-center justify-between mb-1">
        <h3 className="font-brand text-lg text-ink-950">Tonnage by Variety</h3>
        {grandTotal > 0 && (
          <span className="font-body text-xs bg-amber-50 text-amber-700 px-2 py-0.5 rounded-full">
            {grandTotal.toFixed(1)} tons total
          </span>
        )}
      </div>
      <p className="font-body text-xs text-stone-500 mb-3">
        Net tons where recorded, gross otherwise — updates automatically as harvest entries are logged.
      </p>
      {chartData.length === 0 ? (
        <p className="font-body text-sm text-stone-400">No harvest tonnage logged yet.</p>
      ) : (
        <div className="flex flex-col sm:flex-row items-center gap-4">
          <div className="w-full sm:w-[220px] shrink-0">
            <ResponsiveContainer width="100%" height={220}>
              <PieChart>
                <Pie data={chartData} dataKey="tons" nameKey="variety" cx="50%" cy="50%" innerRadius={52} outerRadius={82} paddingAngle={2} stroke="none">
                  {chartData.map((entry, i) => (
                    <Cell key={entry.variety} fill={FERMENT_CHART_COLORS[i % FERMENT_CHART_COLORS.length]} />
                  ))}
                </Pie>
                <Tooltip
                  contentStyle={{ fontSize: 12, fontFamily: "Inter, sans-serif", borderRadius: 10, border: "1px solid #e7e5e4", boxShadow: "0 4px 12px rgba(0,0,0,0.06)" }}
                  formatter={(value) => [`${value} tons`, ""]}
                />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="flex-1 w-full space-y-1.5">
            {chartData.map((entry, i) => (
              <div key={entry.variety} className="flex items-center justify-between font-body text-sm">
                <span className="flex items-center gap-2 text-stone-700">
                  <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: FERMENT_CHART_COLORS[i % FERMENT_CHART_COLORS.length] }} />
                  {entry.variety}
                </span>
                <span className="text-stone-500">
                  {entry.tons.toFixed(1)}T{grandTotal > 0 ? ` · ${Math.round((entry.tons / grandTotal) * 100)}%` : ""}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function groupFruitAnalysisByBlock(entries) {
  const groups = {};
  entries.forEach((e) => {
    const key = e.block || "No Block Set";
    if (!groups[key]) groups[key] = [];
    groups[key].push(e);
  });
  return Object.entries(groups)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([block, blockEntries]) => ({
      block,
      entries: [...blockEntries].sort((a, b) => (a.date > b.date ? 1 : -1)),
    }));
}

// ---------- Visual Brix + pH ripening curve for one vineyard block ----------
function RipeningChart({ block, entries }) {
  const chartData = entries
    .filter((r) => (r.brix !== "" && r.brix != null) || (r.ph !== "" && r.ph != null) || (r.ta !== "" && r.ta != null))
    .map((r) => ({
      date: r.date ? r.date.slice(5) : "",
      fullDate: r.date,
      brix: r.brix !== "" && r.brix != null ? parseFloat(r.brix) : null,
      ph: r.ph !== "" && r.ph != null ? parseFloat(r.ph) : null,
      ta: r.ta !== "" && r.ta != null ? parseFloat(r.ta) : null,
    }));

  if (chartData.length < 2) return null;

  return (
    <div className="bg-white border border-stone-200 rounded-md p-3">
      <p className="font-body text-xs font-semibold text-stone-600 mb-2">{block}</p>
      <ResponsiveContainer width="100%" height={180}>
        <LineChart data={chartData} margin={{ top: 5, right: 8, left: -18, bottom: 0 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#e7e5e4" />
          <XAxis dataKey="date" tick={{ fontSize: 11, fill: "#78716c" }} />
          <YAxis yAxisId="brix" tick={{ fontSize: 11, fill: "#065f46" }} width={34} />
          <YAxis yAxisId="ph" orientation="right" tick={{ fontSize: 11, fill: "#b45309" }} width={34} domain={[2.8, 4.2]} />
          <YAxis yAxisId="ta" hide domain={["auto", "auto"]} />
          <Tooltip
            contentStyle={{ fontSize: 12, fontFamily: "Inter, sans-serif", borderRadius: 8, border: "1px solid #e7e5e4" }}
            labelFormatter={(label, payload) => payload?.[0]?.payload?.fullDate || label}
          />
          <Legend wrapperStyle={{ fontSize: 11 }} />
          <Line yAxisId="brix" type="monotone" dataKey="brix" name="Brix" stroke="#065f46" strokeWidth={2} dot={{ r: 3 }} connectNulls />
          <Line yAxisId="ph" type="monotone" dataKey="ph" name="pH" stroke="#b45309" strokeWidth={2} dot={{ r: 3 }} connectNulls />
          <Line yAxisId="ta" type="monotone" dataKey="ta" name="TA (g/L)" stroke="#7c3aed" strokeWidth={2} dot={{ r: 3 }} connectNulls />
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
}

// ---------- Generic paste-to-import panel: paste CSV/tab-separated text with a header row,
// matches columns to a section's fields by name or label (case/spacing insensitive), and shows
// a preview before actually adding anything. ----------
// ---------- Harvest Tonnage batch entry: shared block/variety/date/clone/weigh master once,
// then as many bin weights as needed — matches weighing multiple bins from the same pick. ----------
function HarvestBatchEntryForm({ fields, onSubmit, saving, vineyardBlocks, onAddBlock, clones, onAddClone, defaultTareWeight, onUpdateDefaultTareWeight }) {
  const headerFieldNames = ["date", "block", "variety", "clone", "weighMaster", "notes"];
  const headerFields = fields.filter((f) => headerFieldNames.includes(f.name));

  const [header, setHeader] = useState(() => {
    const initial = { date: todayISO() };
    headerFields.forEach((f) => {
      if (!(f.name in initial)) initial[f.name] = "";
    });
    return initial;
  });
  const emptyBin = () => ({ id: genId(), tons: "", lbs: "", tareWeight: defaultTareWeight || "", netTons: "", netLbs: "" });
  const [bins, setBins] = useState([emptyBin()]);
  const [error, setError] = useState("");

  const updateBin = (id, field, value) => {
    setBins((prev) =>
      prev.map((b) => {
        if (b.id !== id) return b;
        const next = { ...b, [field]: value };
        if (field === "tons" || field === "lbs") {
          if (value === "") {
            next[field === "tons" ? "lbs" : "tons"] = "";
          } else {
            const num = parseFloat(value);
            if (!isNaN(num)) {
              if (field === "tons") next.lbs = String(Math.round(num * 2000 * 100) / 100);
              else next.tons = String(Math.round((num / 2000) * 1000) / 1000);
            }
          }
        }
        const grossLbs = parseFloat(next.lbs);
        const tare = parseFloat(next.tareWeight);
        if (!isNaN(grossLbs) && !isNaN(tare) && grossLbs >= tare) {
          const netLbsVal = grossLbs - tare;
          next.netLbs = String(Math.round(netLbsVal * 100) / 100);
          next.netTons = String(Math.round((netLbsVal / 2000) * 1000) / 1000);
        } else {
          next.netLbs = "";
          next.netTons = "";
        }
        return next;
      })
    );
  };
  const addBinRow = () => setBins((prev) => [...prev, emptyBin()]);
  const removeBinRow = (id) => setBins((prev) => (prev.length > 1 ? prev.filter((b) => b.id !== id) : prev));

  const submit = (e) => {
    e.preventDefault();
    if (!header.block) {
      setError('Please fill in "Block / Vineyard"');
      return;
    }
    const filledBins = bins.filter((b) => b.tons || b.lbs);
    if (filledBins.length === 0) {
      setError("Add at least one bin's weight");
      return;
    }
    setError("");
    onSubmit(header, filledBins);
    setBins([emptyBin()]);
  };

  return (
    <form onSubmit={submit} className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5 mb-6">
      <h2 className="font-brand text-lg text-ink-950 mb-1">New Harvest Tonnage Entry</h2>
      <p className="font-body text-xs text-stone-500 mb-3">
        Fill in the block/variety/date once, then log each bin's weight below — handy when several bins come off the same pick.
      </p>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4 pb-4 border-b border-stone-100">
        {headerFields.map((f) => (
          <Field
            key={f.name}
            f={f}
            value={header[f.name]}
            onChange={(v) => setHeader((p) => ({ ...p, [f.name]: v }))}
            blocksList={vineyardBlocks}
            onAddBlock={onAddBlock}
            clonesList={clones}
            onAddClone={onAddClone}
          />
        ))}
      </div>

      <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
        <p className="font-body text-xs font-semibold text-stone-600">Bin Weights</p>
        <div className="flex items-center gap-2">
          <label className="font-body text-xs text-stone-500">Default Tare (lbs):</label>
          <input
            type="number"
            value={defaultTareWeight}
            onChange={(e) => onUpdateDefaultTareWeight(e.target.value)}
            placeholder="e.g. 25"
            className="font-body w-20 border border-stone-300 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
          />
        </div>
      </div>
      <div className="space-y-2 mb-3">
        {bins.map((bin, i) => (
          <div key={bin.id} className="grid grid-cols-2 sm:grid-cols-5 gap-2 items-end bg-stone-50 border border-stone-200 rounded-md p-2">
            <div>
              <label className="font-body block text-xs font-medium text-stone-600 mb-1">Gross Tons</label>
              <input
                type="number"
                value={bin.tons}
                onChange={(e) => updateBin(bin.id, "tons", e.target.value)}
                className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
              />
            </div>
            <div>
              <label className="font-body block text-xs font-medium text-stone-600 mb-1">Gross Lbs</label>
              <input
                type="number"
                value={bin.lbs}
                onChange={(e) => updateBin(bin.id, "lbs", e.target.value)}
                className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
              />
            </div>
            <div>
              <label className="font-body block text-xs font-medium text-stone-600 mb-1">Tare (lbs)</label>
              <input
                type="number"
                value={bin.tareWeight}
                onChange={(e) => updateBin(bin.id, "tareWeight", e.target.value)}
                className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
              />
            </div>
            <div>
              <label className="font-body block text-xs font-medium text-stone-600 mb-1">Net Tons</label>
              <input type="text" readOnly value={bin.netTons} className="font-body w-full border border-stone-200 bg-stone-100 rounded-md px-2 py-1.5 text-sm text-stone-600" />
            </div>
            <div className="flex items-end gap-1">
              <div className="flex-1">
                <label className="font-body block text-xs font-medium text-stone-600 mb-1">Net Lbs</label>
                <input type="text" readOnly value={bin.netLbs} className="font-body w-full border border-stone-200 bg-stone-100 rounded-md px-2 py-1.5 text-sm text-stone-600" />
              </div>
              {bins.length > 1 && (
                <button type="button" onClick={() => removeBinRow(bin.id)} className="text-stone-400 hover:text-red-700 pb-2" title="Remove bin">
                  <Trash2 size={16} />
                </button>
              )}
            </div>
            <p className="font-body text-xs text-stone-400 col-span-2 sm:col-span-5 -mt-1">Bin {i + 1}</p>
          </div>
        ))}
      </div>
      <button
        type="button"
        onClick={addBinRow}
        className="font-body flex items-center gap-1 text-xs font-medium text-ink-800 hover:text-ink-900 mb-3"
      >
        <Plus size={13} /> Add Another Bin
      </button>

      {error && <p className="font-body text-sm text-red-700 mb-2">{error}</p>}
      <button
        type="submit"
        disabled={saving}
        className="font-body flex items-center gap-2 bg-ink-900 hover:bg-ink-800 disabled:opacity-50 text-white text-sm font-medium px-4 py-2 rounded-md transition-colors"
      >
        {saving ? <Loader2 size={16} className="animate-spin" /> : <Plus size={16} />}
        Save {bins.length > 1 ? `${bins.length} Bins` : "Entry"}
      </button>
    </form>
  );
}

function BulkImportPanel({ fields, onImport, onClose }) {
  const [text, setText] = useState("");
  const [error, setError] = useState("");
  const [preview, setPreview] = useState(null);

  const normalize = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const fieldLookup = {};
  fields.forEach((f) => {
    fieldLookup[normalize(f.name)] = f.name;
    fieldLookup[normalize(f.label)] = f.name;
  });

  const parse = () => {
    if (!text.trim()) {
      setError("Paste some data first.");
      return;
    }
    const result = Papa.parse(text.trim(), { header: true, skipEmptyLines: true });
    if (result.errors && result.errors.length > 0 && result.data.length === 0) {
      setError("Couldn't read that as a table — make sure it's copied from Excel/CSV with a header row.");
      return;
    }
    const rows = result.data
      .map((row) => {
        const mapped = {};
        Object.entries(row).forEach(([key, val]) => {
          const fieldName = fieldLookup[normalize(key)];
          if (fieldName && val !== "") mapped[fieldName] = val;
        });
        return mapped;
      })
      .filter((r) => Object.keys(r).length > 0);
    if (rows.length === 0) {
      setError("No recognizable columns found — check that your header row matches the field names shown below.");
      return;
    }
    setError("");
    setPreview(rows);
  };

  const confirmImport = () => {
    onImport(preview);
    setText("");
    setPreview(null);
    onClose();
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5 mb-6">
      <div className="flex items-center justify-between mb-1">
        <h2 className="font-brand text-lg text-ink-950">Import from Paste</h2>
        <button onClick={onClose} className="text-stone-400 hover:text-stone-600">
          <X size={18} />
        </button>
      </div>
      <p className="font-body text-xs text-stone-500 mb-2">
        Copy rows from Excel or a CSV file — including the header row — and paste below. Recognized column headers
        (any of these work, spacing/case doesn't matter): {fields.map((f) => f.label).join(", ")}.
      </p>
      <textarea
        value={text}
        onChange={(e) => {
          setText(e.target.value);
          setPreview(null);
        }}
        rows={6}
        placeholder="Paste your data here, including the header row..."
        className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm font-mono mb-2 focus:outline-none focus:ring-2 focus:ring-ink-800"
      />
      {error && <p className="font-body text-xs text-red-700 mb-2">{error}</p>}
      {!preview ? (
        <button
          onClick={parse}
          className="font-body text-sm font-medium bg-ink-900 hover:bg-ink-800 text-white px-4 py-2 rounded-md"
        >
          Preview Import
        </button>
      ) : (
        <div>
          <p className="font-body text-xs text-stone-600 mb-2">
            Found {preview.length} row{preview.length === 1 ? "" : "s"} ready to import.
          </p>
          <div className="flex gap-3">
            <button
              onClick={confirmImport}
              className="font-body text-sm font-medium bg-ink-900 hover:bg-ink-800 text-white px-4 py-2 rounded-md"
            >
              Import {preview.length} Row{preview.length === 1 ? "" : "s"}
            </button>
            <button onClick={() => setPreview(null)} className="font-body text-sm text-stone-500 hover:text-stone-700">
              Back
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ---------- Picker linking a fermentation lot back to the harvest pick(s) it came from ----------
function HarvestRefsPicker({ value, onChange, harvestEntries }) {
  const [search, setSearch] = useState("");
  const selected = Array.isArray(value) ? value : [];
  const toggle = (id) => onChange(selected.includes(id) ? selected.filter((x) => x !== id) : [...selected, id]);
  const labelFor = (h) =>
    `${h.date || "?"} — ${h.block || "Unknown block"}${h.variety ? " · " + h.variety : ""}${h.netTons ? ` (${h.netTons}T net)` : h.tons ? ` (${h.tons}T)` : ""}`;
  const filtered = search ? harvestEntries.filter((h) => labelFor(h).toLowerCase().includes(search.toLowerCase())) : harvestEntries;
  const selectedEntries = harvestEntries.filter((h) => selected.includes(h.id));

  return (
    <div>
      <input
        type="text"
        placeholder="Search harvest picks by date, block, or variety"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm mb-2 focus:outline-none focus:ring-2 focus:ring-ink-800"
      />
      <div className="flex flex-wrap gap-1.5 max-h-28 overflow-y-auto border border-stone-200 rounded-md p-2">
        {filtered.length === 0 ? (
          <span className="font-body text-xs text-stone-400">No harvest entries match.</span>
        ) : (
          filtered.slice(0, 60).map((h) => {
            const isSelected = selected.includes(h.id);
            return (
              <button
                key={h.id}
                type="button"
                onClick={() => toggle(h.id)}
                className={`font-body text-xs px-2 py-1 rounded-full border transition-colors ${
                  isSelected ? "bg-ink-800 text-white border-ink-800" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                }`}
              >
                {labelFor(h)}
              </button>
            );
          })
        )}
      </div>
      {selectedEntries.length > 0 && (
        <div className="flex flex-wrap gap-1.5 mt-2">
          {selectedEntries.map((h) => (
            <span key={h.id} className="font-body flex items-center gap-1 text-xs bg-ink-50 text-ink-800 px-2 py-1 rounded-full">
              {labelFor(h)}
              <button type="button" onClick={() => toggle(h.id)} className="hover:text-red-700">
                <X size={12} />
              </button>
            </span>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- Tech Sheets: gathers everything needed for a wine + vintage into one shareable
// document — harvest date pulled from Harvest Tonnage, vineyard phenology, blend composition,
// elevage, and vintage notes, with a bottle shot image and clean print/PDF output. ----------
function harvestDatesFor(refs, harvestEntries) {
  const dates = [...new Set(harvestEntries.filter((h) => (refs || []).includes(h.id)).map((h) => h.date).filter(Boolean))].sort();
  return dates;
}

function emptyTechSheetForm() {
  return {
    id: "",
    wineName: "",
    vintage: "",
    harvestRefs: [],
    budbreakDate: "",
    veraisonDate: "",
    varietyCloneBlend: "",
    vintageNotes: "",
    newOakPercent: "",
    elevageDetails: "",
    bottleShotImage: "",
  };
}

function TechSheetForm({ initial, harvestEntries, onSave, onCancel }) {
  const [form, setForm] = useState(initial || emptyTechSheetForm());
  const [error, setError] = useState("");

  const submit = (e) => {
    e.preventDefault();
    if (!form.wineName.trim()) {
      setError('Please fill in "Wine Name"');
      return;
    }
    setError("");
    onSave({ ...form, id: form.id || genId() });
  };

  return (
    <form onSubmit={submit} className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5 mb-6">
      <h2 className="font-brand text-lg text-ink-950 mb-3">{initial ? "Edit Tech Sheet" : "New Tech Sheet"}</h2>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4">
        <div>
          <label className="font-body block text-xs font-medium text-stone-600 mb-1">Wine Name</label>
          <input
            type="text"
            value={form.wineName}
            onChange={(e) => setForm((p) => ({ ...p, wineName: e.target.value }))}
            className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
          />
        </div>
        <div>
          <label className="font-body block text-xs font-medium text-stone-600 mb-1">Vintage</label>
          <select
            value={form.vintage}
            onChange={(e) => setForm((p) => ({ ...p, vintage: e.target.value }))}
            className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
          >
            <option value="">Select…</option>
            {VINTAGE_YEARS.map((y) => (
              <option key={y} value={y}>{y}</option>
            ))}
          </select>
        </div>
      </div>

      <div className="mb-4">
        <label className="font-body block text-xs font-medium text-stone-600 mb-1">Harvest Pick(s)</label>
        <p className="font-body text-xs text-stone-400 mb-1">Link the Harvest Tonnage entries for this wine — the harvest date is pulled from these automatically.</p>
        <HarvestRefsPicker value={form.harvestRefs} onChange={(refs) => setForm((p) => ({ ...p, harvestRefs: refs }))} harvestEntries={harvestEntries} />
        {form.harvestRefs.length > 0 && (
          <p className="font-body text-xs text-ink-700 mt-1">
            Harvest date{harvestDatesFor(form.harvestRefs, harvestEntries).length > 1 ? "s" : ""}: {harvestDatesFor(form.harvestRefs, harvestEntries).join(", ")}
          </p>
        )}
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4">
        <div>
          <label className="font-body block text-xs font-medium text-stone-600 mb-1">Budbreak Date</label>
          <input
            type="date"
            value={form.budbreakDate}
            onChange={(e) => setForm((p) => ({ ...p, budbreakDate: e.target.value }))}
            className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
          />
        </div>
        <div>
          <label className="font-body block text-xs font-medium text-stone-600 mb-1">Veraison Date</label>
          <input
            type="date"
            value={form.veraisonDate}
            onChange={(e) => setForm((p) => ({ ...p, veraisonDate: e.target.value }))}
            className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
          />
        </div>
      </div>

      <div className="mb-4">
        <label className="font-body block text-xs font-medium text-stone-600 mb-1">Variety & Clone Blend</label>
        <input
          type="text"
          placeholder="e.g. 100% Pinot Noir (clones 114, 115, 777)"
          value={form.varietyCloneBlend}
          onChange={(e) => setForm((p) => ({ ...p, varietyCloneBlend: e.target.value }))}
          className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4">
        <div>
          <label className="font-body block text-xs font-medium text-stone-600 mb-1">New Oak %</label>
          <input
            type="number"
            value={form.newOakPercent}
            onChange={(e) => setForm((p) => ({ ...p, newOakPercent: e.target.value }))}
            className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
          />
        </div>
        <div>
          <label className="font-body block text-xs font-medium text-stone-600 mb-1">Bottle Shot</label>
          <input
            type="file"
            accept="image/*"
            onChange={async (e) => {
              const file = e.target.files?.[0];
              if (!file) return;
              try {
                const dataUrl = await compressImage(file);
                setForm((p) => ({ ...p, bottleShotImage: dataUrl }));
              } catch {
                // ignore failed upload
              }
              e.target.value = "";
            }}
            className="font-body w-full text-xs text-stone-600 file:mr-2 file:py-2 file:px-3 file:rounded-md file:border-0 file:bg-ink-900 file:text-white file:text-xs file:font-medium hover:file:bg-ink-800"
          />
          {form.bottleShotImage && (
            <div className="mt-2 flex items-center gap-2">
              <img src={form.bottleShotImage} alt="Bottle shot preview" className="w-14 h-14 object-cover rounded border border-stone-300" />
              <button type="button" onClick={() => setForm((p) => ({ ...p, bottleShotImage: "" }))} className="font-body text-xs text-stone-400 hover:text-red-700">
                Remove
              </button>
            </div>
          )}
        </div>
      </div>

      <div className="mb-4">
        <label className="font-body block text-xs font-medium text-stone-600 mb-1">Elevage Details</label>
        <textarea
          placeholder="e.g. 14 months in French oak barrels (30% new), remainder neutral; 4 months in stainless steel tank prior to blending"
          value={form.elevageDetails}
          onChange={(e) => setForm((p) => ({ ...p, elevageDetails: e.target.value }))}
          rows={2}
          className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
      </div>

      <div className="mb-4">
        <label className="font-body block text-xs font-medium text-stone-600 mb-1">Vintage Notes</label>
        <textarea
          value={form.vintageNotes}
          onChange={(e) => setForm((p) => ({ ...p, vintageNotes: e.target.value }))}
          rows={3}
          className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
      </div>

      {error && <p className="font-body text-sm text-red-700 mb-3">{error}</p>}
      <div className="flex items-center gap-3">
        <button type="submit" className="font-body flex items-center gap-2 bg-ink-900 hover:bg-ink-800 text-white text-sm font-medium px-4 py-2 rounded-md">
          <Check size={16} /> Save Tech Sheet
        </button>
        <button type="button" onClick={onCancel} className="font-body text-sm text-stone-500 hover:text-stone-700">
          Cancel
        </button>
      </div>
    </form>
  );
}

function TechSheetCard({ sheet, harvestEntries, onView, onEdit, onDelete }) {
  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 flex items-center gap-3">
      {sheet.bottleShotImage ? (
        <img src={sheet.bottleShotImage} alt={sheet.wineName} className="w-12 h-12 object-cover rounded border border-stone-200 shrink-0" />
      ) : (
        <div className="w-12 h-12 rounded bg-stone-100 flex items-center justify-center shrink-0">
          <Wine size={20} className="text-stone-300" />
        </div>
      )}
      <div className="flex-1 min-w-0">
        <p className="font-brand text-base text-ink-950 truncate">{sheet.wineName || "Untitled"}</p>
        <p className="font-body text-xs text-stone-500">
          {sheet.vintage || "No vintage"}
          {harvestDatesFor(sheet.harvestRefs, harvestEntries).length > 0 && ` · Harvest ${harvestDatesFor(sheet.harvestRefs, harvestEntries)[0]}`}
        </p>
      </div>
      <div className="flex items-center gap-2 shrink-0">
        <button onClick={onView} className="font-body text-xs font-medium text-ink-800 hover:text-ink-900 border border-ink-200 rounded-md px-2.5 py-1.5">
          View
        </button>
        {onEdit && (
          <button onClick={onEdit} className="text-stone-400 hover:text-ink-800" title="Edit">
            <Pencil size={15} />
          </button>
        )}
        {onDelete && (
          <button onClick={onDelete} className="text-stone-400 hover:text-red-700" title="Delete">
            <Trash2 size={15} />
          </button>
        )}
      </div>
    </div>
  );
}

function TechSheetDocument({ sheet, harvestEntries, onPrint, onClose }) {
  const dates = harvestDatesFor(sheet.harvestRefs, harvestEntries);
  const row = (label, value) =>
    value ? (
      <div className="flex justify-between py-2 border-b border-stone-100 last:border-b-0">
        <span className="font-body text-xs text-stone-500">{label}</span>
        <span className="font-body text-sm text-stone-800 text-right">{value}</span>
      </div>
    ) : null;

  return (
    <div className="bg-white border border-stone-200 rounded-lg overflow-hidden mb-6">
      <div className="flex items-center justify-between px-4 py-3 border-b border-stone-200">
        <h2 className="font-brand text-lg text-ink-950">Tech Sheet</h2>
        <div className="flex items-center gap-3">
          <button onClick={onPrint} className="font-body flex items-center gap-1.5 text-sm font-medium text-ink-900 hover:text-ink-700">
            <Printer size={15} /> Print / PDF
          </button>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-600">
            <X size={18} />
          </button>
        </div>
      </div>
      <div className="p-5 sm:p-6">
        {sheet.bottleShotImage && (
          <img src={sheet.bottleShotImage} alt={sheet.wineName} className="w-32 h-32 object-cover rounded-lg border border-stone-200 mb-4" />
        )}
        <h1 className="font-brand text-2xl text-ink-950 mb-1">{sheet.wineName || "Untitled Wine"}</h1>
        <p className="font-body text-stone-500 mb-4">{sheet.vintage ? `${sheet.vintage} Vintage` : "Vintage not set"}</p>

        <p className="font-body text-xs font-semibold text-stone-600 uppercase tracking-wide mb-1 mt-4">Vineyard</p>
        {row("Harvest Date(s)", dates.join(", "))}
        {row("Budbreak", sheet.budbreakDate)}
        {row("Veraison", sheet.veraisonDate)}
        {row("Variety & Clones", sheet.varietyCloneBlend)}

        <p className="font-body text-xs font-semibold text-stone-600 uppercase tracking-wide mb-1 mt-4">Winemaking</p>
        {row("New Oak", sheet.newOakPercent !== "" && sheet.newOakPercent != null ? `${sheet.newOakPercent}%` : "")}
        {row("Elevage", sheet.elevageDetails)}

        {sheet.vintageNotes && (
          <>
            <p className="font-body text-xs font-semibold text-stone-600 uppercase tracking-wide mb-1 mt-4">Vintage Notes</p>
            <p className="font-body text-sm text-stone-700 whitespace-pre-wrap">{sheet.vintageNotes}</p>
          </>
        )}
      </div>
    </div>
  );
}

// ---------- THO Payout Calculator: combines Timesheets + Tips into a per-associate payout,
// using either exact hours worked or half-day/full-day shares — computed per day/shift, then
// summed across the selected range, so a busy day's tips aren't blended evenly with a slow one. ----------
function computeTipPayout(timesheets, tips, startDate, endDate, method, halfShareWeight, fullShareWeight) {
  const inRange = (d) => (!startDate || d >= startDate) && (!endDate || d <= endDate);
  const timesheetsByDate = {};
  timesheets.filter((t) => inRange(t.date)).forEach((t) => {
    if (!timesheetsByDate[t.date]) timesheetsByDate[t.date] = [];
    timesheetsByDate[t.date].push(t);
  });
  const tipsByDate = {};
  tips.filter((t) => inRange(t.date)).forEach((t) => {
    tipsByDate[t.date] = (tipsByDate[t.date] || 0) + (parseFloat(t.totalTips) || 0);
  });

  const payoutByEmployee = {};
  const warnings = [];
  const dayBreakdown = [];
  const allDates = [...new Set([...Object.keys(timesheetsByDate), ...Object.keys(tipsByDate)])].sort();

  allDates.forEach((date) => {
    const dayEntries = timesheetsByDate[date] || [];
    const dayTips = tipsByDate[date] || 0;

    if (dayEntries.length === 0 && dayTips > 0) {
      warnings.push(`No timesheet data for ${date} — $${dayTips.toFixed(2)} in tips couldn't be distributed.`);
      return;
    }
    if (dayEntries.length > 0 && dayTips === 0) {
      warnings.push(`No tips recorded for ${date} — that day's hours won't generate a tip-out.`);
    }
    if (dayTips === 0 || dayEntries.length === 0) return;

    const dayResult = [];
    if (method === "shares") {
      const totalShares = dayEntries.reduce((sum, t) => sum + ((parseFloat(t.hoursWorked) || 0) >= 4 ? fullShareWeight : halfShareWeight), 0);
      if (totalShares > 0) {
        dayEntries.forEach((t) => {
          const hours = parseFloat(t.hoursWorked) || 0;
          const shareWeight = hours >= 4 ? fullShareWeight : halfShareWeight;
          const amount = (shareWeight / totalShares) * dayTips;
          dayResult.push({ employeeName: t.employeeName, hours, amount });
        });
      }
    } else {
      const totalHours = dayEntries.reduce((sum, t) => sum + (parseFloat(t.hoursWorked) || 0), 0);
      if (totalHours > 0) {
        dayEntries.forEach((t) => {
          const hours = parseFloat(t.hoursWorked) || 0;
          const amount = (hours / totalHours) * dayTips;
          dayResult.push({ employeeName: t.employeeName, hours, amount });
        });
      }
    }

    dayResult.forEach(({ employeeName, hours, amount }) => {
      if (!payoutByEmployee[employeeName]) payoutByEmployee[employeeName] = { totalHours: 0, totalPayout: 0 };
      payoutByEmployee[employeeName].totalHours += hours;
      payoutByEmployee[employeeName].totalPayout += amount;
    });
    dayBreakdown.push({ date, tips: dayTips, entries: dayResult });
  });

  const rows = Object.entries(payoutByEmployee)
    .map(([employeeName, v]) => ({ employeeName, totalHours: Math.round(v.totalHours * 100) / 100, totalPayout: Math.round(v.totalPayout * 100) / 100 }))
    .sort((a, b) => a.employeeName.localeCompare(b.employeeName));

  return { rows, warnings, dayBreakdown, totalTips: Object.values(tipsByDate).reduce((s, v) => s + v, 0) };
}

// ---------- Reusable add-form + editable table for one section, usable standalone on a page
// that shows several sections at once (rather than the single-activeKey generic table) ----------
function SimpleDataPanel({ title, fields, rows, onAdd, onUpdate, onDelete, confirmAction, associatesList, onAddAssociate, allowImport }) {
  const [form, setForm] = useState(() => emptyForm(fields));
  const [error, setError] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState(null);
  const [showImport, setShowImport] = useState(false);

  const submit = (e) => {
    e.preventDefault();
    const missing = fields.find((f) => f.type !== "textarea" && f.type !== "photo" && !f.optional && !String(form[f.name] ?? "").trim());
    if (missing) {
      setError(`Please fill in "${missing.label}"`);
      return;
    }
    setError("");
    onAdd({ id: genId(), ...form });
    setForm(emptyForm(fields));
  };

  const startEdit = (row) => {
    setEditingId(row.id);
    setEditForm({ ...row });
    setError("");
  };
  const saveEdit = () => {
    const missing = fields.find((f) => f.type !== "textarea" && f.type !== "photo" && !f.optional && !String(editForm[f.name] ?? "").trim());
    if (missing) {
      setError(`Please fill in "${missing.label}"`);
      return;
    }
    setError("");
    onUpdate(editingId, editForm);
    setEditingId(null);
    setEditForm(null);
  };

  const handleDeleteRow = (id) => {
    confirmAction("Delete this entry? This can't be undone.", () => onDelete(id));
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg overflow-hidden mb-6">
      <div className="px-4 py-3 border-b border-stone-200 flex items-center justify-between">
        <h2 className="font-brand text-lg text-ink-950">{title}</h2>
        {allowImport && (
          <button
            onClick={() => setShowImport((v) => !v)}
            className="font-body flex items-center gap-1.5 text-xs font-medium text-ink-900 hover:text-ink-700 border border-ink-200 rounded-md px-2.5 py-1.5"
          >
            <UploadCloud size={13} /> Import
          </button>
        )}
      </div>
      <div className="p-4">
        {showImport && (
          <BulkImportPanel
            fields={fields}
            onImport={(importedRows) => {
              importedRows.forEach((r) => onAdd({ id: genId(), ...emptyForm(fields), ...r }));
              setShowImport(false);
            }}
            onClose={() => setShowImport(false)}
          />
        )}
        <form onSubmit={submit} className="grid grid-cols-1 sm:grid-cols-3 gap-2 mb-3 items-end">
          {fields.map((f) => (
            <div key={f.name}>
              <Field f={f} value={form[f.name]} onChange={(v) => setForm((p) => ({ ...p, [f.name]: v }))} associatesList={associatesList} onAddAssociate={onAddAssociate} />
            </div>
          ))}
          <button type="submit" className="font-body flex items-center gap-1.5 bg-ink-900 hover:bg-ink-800 text-white text-sm font-medium px-3 py-2 rounded-md h-fit">
            <Plus size={15} /> Add
          </button>
        </form>
        {error && <p className="font-body text-xs text-red-700 mb-2">{error}</p>}
        {rows.length === 0 ? (
          <p className="font-body text-sm text-stone-400 text-center py-4">No entries yet.</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm font-body">
              <thead>
                <tr className="text-left text-stone-500 bg-stone-50">
                  {fields.map((f) => (
                    <th key={f.name} className="px-3 py-2 whitespace-nowrap">{f.label}</th>
                  ))}
                  <th className="px-3 py-2 w-16"></th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row) => {
                  const isEditing = editingId === row.id;
                  return (
                    <tr key={row.id} className={`border-t border-stone-100 ${isEditing ? "bg-ink-50" : ""}`}>
                      {fields.map((f) => (
                        <td key={f.name} className="px-2 py-2 align-top" style={{ minWidth: isEditing ? 120 : undefined }}>
                          {isEditing ? (
                            <Field f={f} value={editForm[f.name]} onChange={(v) => setEditForm((p) => ({ ...p, [f.name]: v }))} hideLabel associatesList={associatesList} onAddAssociate={onAddAssociate} />
                          ) : f.type === "photo" ? (
                            row[f.name] ? (
                              <a href={row[f.name]} target="_blank" rel="noopener noreferrer">
                                <img src={row[f.name]} alt="Receipt" className="w-9 h-9 object-cover rounded border border-stone-300" />
                              </a>
                            ) : (
                              "—"
                            )
                          ) : f.type === "document" ? (
                            row[f.name] ? (
                              row[f.name].startsWith("data:application/pdf") ? (
                                <a href={row[f.name]} target="_blank" rel="noopener noreferrer" className="font-body text-xs text-ink-800 underline">
                                  View PDF
                                </a>
                              ) : (
                                <a href={row[f.name]} target="_blank" rel="noopener noreferrer">
                                  <img src={row[f.name]} alt="Document" className="w-9 h-9 object-cover rounded border border-stone-300" />
                                </a>
                              )
                            ) : (
                              "—"
                            )
                          ) : (
                            <span className="whitespace-nowrap">{row[f.name] || "—"}</span>
                          )}
                        </td>
                      ))}
                      <td className="px-2 py-2">
                        {isEditing ? (
                          <div className="flex items-center gap-2">
                            <button onClick={saveEdit} className="text-ink-700 hover:text-ink-900" title="Save">
                              <Check size={15} />
                            </button>
                            <button onClick={() => { setEditingId(null); setEditForm(null); setError(""); }} className="text-stone-400 hover:text-stone-600" title="Cancel">
                              <X size={15} />
                            </button>
                          </div>
                        ) : (
                          <div className="flex items-center gap-2">
                            <button onClick={() => startEdit(row)} className="text-stone-400 hover:text-ink-800" title="Edit">
                              <Pencil size={14} />
                            </button>
                            <button onClick={() => handleDeleteRow(row.id)} className="text-stone-400 hover:text-red-700" title="Delete">
                              <Trash2 size={15} />
                            </button>
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

// ---------- Combined Payroll page: Timesheets, then Tips, then the Payout Calculator, all on
// one page — so the whole payroll workflow reads top to bottom without switching tabs. ----------
function THOPayrollPage({ data, onAddThoEntry, onUpdateThoEntry, onDeleteThoEntry, confirmAction, tastingAssociates, onAddAssociate }) {
  const timesheetFields = SIMPLE_SECTIONS.find((s) => s.key === "thoTimesheets").fields;
  const tipsFields = SIMPLE_SECTIONS.find((s) => s.key === "thoTips").fields;

  return (
    <div>
      <SimpleDataPanel
        title="Timesheets"
        fields={timesheetFields}
        rows={data.thoTimesheets}
        onAdd={(entry) => onAddThoEntry("thoTimesheets", entry)}
        onUpdate={(id, changes) => onUpdateThoEntry("thoTimesheets", id, changes)}
        onDelete={(id) => onDeleteThoEntry("thoTimesheets", id)}
        confirmAction={confirmAction}
        associatesList={tastingAssociates}
        onAddAssociate={onAddAssociate}
        allowImport
      />
      <SimpleDataPanel
        title="Tips"
        fields={tipsFields}
        rows={data.thoTips}
        onAdd={(entry) => onAddThoEntry("thoTips", entry)}
        onUpdate={(id, changes) => onUpdateThoEntry("thoTips", id, changes)}
        onDelete={(id) => onDeleteThoEntry("thoTips", id)}
        confirmAction={confirmAction}
        allowImport
      />
      <THOPayoutCalculator data={data} />
    </div>
  );
}

// ---------- Tech Sheets panel — list/form/view modes, extracted so it can be embedded inside
// the About Alloro tab instead of living as its own top-level page. ----------
function TechSheetsPanel({ data, techSheetMode, setTechSheetMode, onSave, onDelete, onPrint }) {
  if (techSheetMode.mode === "form") {
    return (
      <TechSheetForm
        initial={techSheetMode.sheetId ? data.techSheets.find((s) => s.id === techSheetMode.sheetId) : null}
        harvestEntries={data.harvest}
        onSave={async (sheet) => {
          await onSave(sheet);
          setTechSheetMode({ mode: "view", sheetId: sheet.id });
        }}
        onCancel={() => setTechSheetMode({ mode: "list", sheetId: null })}
      />
    );
  }
  if (techSheetMode.mode === "view" && techSheetMode.sheetId) {
    const sheet = data.techSheets.find((s) => s.id === techSheetMode.sheetId);
    if (!sheet) return null;
    return (
      <TechSheetDocument
        sheet={sheet}
        harvestEntries={data.harvest}
        onPrint={() => onPrint(sheet.id)}
        onClose={() => setTechSheetMode({ mode: "list", sheetId: null })}
      />
    );
  }
  return (
    <>
      <div className="flex items-center justify-between mb-4">
        <div>
          <h2 className="font-brand text-xl text-ink-950">Tech Sheets</h2>
          <p className="font-body text-xs text-stone-500">One document per wine and vintage — for your team and for distribution.</p>
        </div>
        <button
          onClick={() => setTechSheetMode({ mode: "form", sheetId: null })}
          className="font-body flex items-center gap-1.5 text-sm font-medium bg-ink-900 hover:bg-ink-800 text-white px-3 py-2 rounded-md"
        >
          <Plus size={15} /> New Tech Sheet
        </button>
      </div>
      {data.techSheets.length === 0 ? (
        <p className="font-body text-sm text-stone-500 bg-white border border-stone-200 rounded-lg p-6 text-center">
          No tech sheets yet — create one once a wine's ready.
        </p>
      ) : (
        <div className="space-y-2">
          {data.techSheets.map((sheet) => (
            <TechSheetCard
              key={sheet.id}
              sheet={sheet}
              harvestEntries={data.harvest}
              onView={() => setTechSheetMode({ mode: "view", sheetId: sheet.id })}
              onEdit={() => setTechSheetMode({ mode: "form", sheetId: sheet.id })}
              onDelete={() => onDelete(sheet.id)}
            />
          ))}
        </div>
      )}
    </>
  );
}

// ---------- Case Production summary — computed from Bottling records, grouped by the calendar
// year each run was bottled (not vintage, since a wine can be bottled the year after harvest). ----------
function CaseProductionSummary({ bottling }) {
  const byYear = {};
  bottling.forEach((b) => {
    if (!b.date) return;
    const year = b.date.slice(0, 4);
    const cases = parseFloat(b.cases) || 0;
    if (!byYear[year]) byYear[year] = { total: 0, byWine: {} };
    byYear[year].total += cases;
    const wineKey = `${b.wineName || "Untitled"}${b.vintage ? " (" + b.vintage + ")" : ""}`;
    byYear[year].byWine[wineKey] = (byYear[year].byWine[wineKey] || 0) + cases;
  });
  const years = Object.keys(byYear).sort((a, b) => b.localeCompare(a));

  return (
    <div className="space-y-3">
      <div>
        <h2 className="font-brand text-xl text-ink-950">Case Production by Year</h2>
        <p className="font-body text-xs text-stone-500">Computed automatically from Bottling records, grouped by the year each run was bottled.</p>
      </div>
      {years.length === 0 ? (
        <p className="font-body text-sm text-stone-500 bg-white border border-stone-200 rounded-lg p-6 text-center">No bottling records yet.</p>
      ) : (
        years.map((year) => (
          <div key={year} className="bg-white border border-stone-200 rounded-lg overflow-hidden">
            <div className="px-4 py-3 border-b border-stone-200 flex items-center justify-between">
              <h3 className="font-brand text-lg text-ink-950">{year}</h3>
              <span className="font-body text-sm text-stone-600">{byYear[year].total.toLocaleString()} cases total</span>
            </div>
            <div className="p-4 space-y-1.5">
              {Object.entries(byYear[year].byWine).map(([wine, cases]) => (
                <div key={wine} className="flex items-center justify-between font-body text-sm">
                  <span className="text-stone-700">{wine}</span>
                  <span className="text-stone-500">{cases.toLocaleString()} cases</span>
                </div>
              ))}
            </div>
          </div>
        ))
      )}
    </div>
  );
}

// ---------- History: a narrative "our story" text plus a dated timeline of milestones ----------
function HistoryPanel({ alloroStory, onUpdateStory, milestones, onAddThoEntry, onUpdateThoEntry, onDeleteThoEntry, confirmAction }) {
  const [storyDraft, setStoryDraft] = useState(alloroStory);
  const [storySaved, setStorySaved] = useState(false);

  const saveStory = () => {
    onUpdateStory(storyDraft);
    setStorySaved(true);
    setTimeout(() => setStorySaved(false), 1500);
  };

  const milestoneFields = [
    { name: "date", label: "Date", type: "date" },
    { name: "title", label: "Title", type: "text" },
    { name: "description", label: "Description", type: "textarea" },
  ];
  const sortedMilestones = [...milestones].sort((a, b) => (a.date < b.date ? -1 : 1));

  return (
    <div className="space-y-4">
      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-1">Our Story</h2>
        <p className="font-body text-xs text-stone-500 mb-3">A narrative about Alloro — the kind of thing you'd put on a website "About" page.</p>
        <textarea
          value={storyDraft}
          onChange={(e) => setStoryDraft(e.target.value)}
          rows={8}
          className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm mb-3 focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
        <button
          onClick={saveStory}
          className={`font-body text-sm font-medium px-4 py-2 rounded-md text-white ${storySaved ? "bg-ink-700" : "bg-ink-900 hover:bg-ink-800"}`}
        >
          {storySaved ? "✓ Saved" : "Save Story"}
        </button>
      </div>

      <SimpleDataPanel
        title="Timeline"
        fields={milestoneFields}
        rows={sortedMilestones}
        onAdd={(entry) => onAddThoEntry("historyMilestones", entry)}
        onUpdate={(id, changes) => onUpdateThoEntry("historyMilestones", id, changes)}
        onDelete={(id) => onDeleteThoEntry("historyMilestones", id)}
        confirmAction={confirmAction}
      />
    </div>
  );
}

// ---------- Vineyard: an uploaded reference map image plus a structured block/variety/acreage
// /soil table ----------
function VineyardPanel({ vineyardMapImage, onUpdateMapImage, blockDetails, onAddThoEntry, onUpdateThoEntry, onDeleteThoEntry, confirmAction, vineyardBlocks, onAddBlock }) {
  const blockFields = SIMPLE_SECTIONS.find((s) => s.key === "vineyardBlockDetails").fields;

  return (
    <div className="space-y-4">
      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-1">Vineyard Map</h2>
        <p className="font-body text-xs text-stone-500 mb-3">Upload a photo or scan of the actual vineyard map for reference.</p>
        {vineyardMapImage && (
          <img src={vineyardMapImage} alt="Vineyard map" className="w-full max-w-lg rounded-lg border border-stone-200 mb-3" />
        )}
        <input
          type="file"
          accept="image/*"
          onChange={async (e) => {
            const file = e.target.files?.[0];
            if (!file) return;
            try {
              const dataUrl = await compressImage(file);
              onUpdateMapImage(dataUrl);
            } catch {
              // ignore failed upload
            }
            e.target.value = "";
          }}
          className="font-body w-full text-xs text-stone-600 file:mr-2 file:py-2 file:px-3 file:rounded-md file:border-0 file:bg-ink-900 file:text-white file:text-xs file:font-medium hover:file:bg-ink-800"
        />
        {vineyardMapImage && (
          <button onClick={() => onUpdateMapImage("")} className="font-body text-xs text-stone-400 hover:text-red-700 mt-2">
            Remove map image
          </button>
        )}
      </div>

      <SimpleDataPanel
        title="Blocks, Varieties & Soil"
        fields={blockFields}
        rows={blockDetails}
        onAdd={(entry) => onAddThoEntry("vineyardBlockDetails", entry)}
        onUpdate={(id, changes) => onUpdateThoEntry("vineyardBlockDetails", id, changes)}
        onDelete={(id) => onDeleteThoEntry("vineyardBlockDetails", id)}
        confirmAction={confirmAction}
        blocksList={vineyardBlocks}
        onAddBlock={onAddBlock}
      />
    </div>
  );
}

// ---------- About Alloro: the shared brand/info hub — Overview (case production), History,
// Vineyard, Accolades, and Tech Sheets, as internal sub-tabs on one page. ----------
// ---------- Read-only Tech Sheets browsing for About Alloro — view and print only. Building
// and editing tech sheets happens under Winery instead, so there's exactly one place that can
// change them, and no risk of the "published" copy drifting from what's actually true. ----------
function TechSheetsReadOnlyPanel({ data, viewSheetId, setViewSheetId, onPrint }) {
  if (viewSheetId) {
    const sheet = data.techSheets.find((s) => s.id === viewSheetId);
    if (!sheet) return null;
    return (
      <TechSheetDocument
        sheet={sheet}
        harvestEntries={data.harvest}
        onPrint={() => onPrint(sheet.id)}
        onClose={() => setViewSheetId(null)}
      />
    );
  }
  return (
    <>
      <div className="mb-4">
        <h2 className="font-brand text-xl text-ink-950">Tech Sheets</h2>
        <p className="font-body text-xs text-stone-500">View and print — built and kept up to date under Winery.</p>
      </div>
      {data.techSheets.length === 0 ? (
        <p className="font-body text-sm text-stone-500 bg-white border border-stone-200 rounded-lg p-6 text-center">
          No tech sheets published yet.
        </p>
      ) : (
        <div className="space-y-2">
          {data.techSheets.map((sheet) => (
            <TechSheetCard key={sheet.id} sheet={sheet} harvestEntries={data.harvest} onView={() => setViewSheetId(sheet.id)} />
          ))}
        </div>
      )}
    </>
  );
}

// ---------- Search across all Team Resources content — History, Vineyard, Accolades, Tech
// Sheets, Contacts, Pricing, and Wine Club — so the team can find what they need from one box
// instead of hunting through sub-tabs. ----------
function searchTeamResources(query, data, alloroStory) {
  const q = query.trim().toLowerCase();
  if (!q) return [];
  const results = [];
  const hit = (haystack) => haystack.toLowerCase().includes(q);

  if (alloroStory && hit(alloroStory)) {
    results.push({ subTab: "history", title: "Our Story", snippet: alloroStory.slice(0, 140) });
  }
  data.historyMilestones.forEach((m) => {
    if (hit(`${m.title || ""} ${m.description || ""}`)) {
      results.push({ subTab: "history", title: m.title || "Milestone", snippet: m.description || "" });
    }
  });
  data.vineyardBlockDetails.forEach((b) => {
    if (hit(`${b.block || ""} ${b.variety || ""} ${b.soilType || ""}`)) {
      results.push({ subTab: "vineyard", title: b.block || "Block", snippet: [b.variety, b.soilType].filter(Boolean).join(" · ") });
    }
  });
  data.accolades.forEach((a) => {
    if (hit(`${a.wineName || ""} ${a.publication || ""} ${a.notes || ""}`)) {
      results.push({ subTab: "accolades", title: `${a.wineName || "Wine"}${a.publication ? " — " + a.publication : ""}`, snippet: a.notes || "" });
    }
  });
  data.techSheets.forEach((t) => {
    if (hit(`${t.wineName || ""} ${t.vintage || ""} ${t.vintageNotes || ""} ${t.varietyCloneBlend || ""} ${t.elevageDetails || ""}`)) {
      results.push({ subTab: "techSheets", title: `${t.wineName || "Wine"}${t.vintage ? " " + t.vintage : ""}`, snippet: t.vintageNotes || "", sheetId: t.id });
    }
  });
  data.contacts.forEach((c) => {
    if (hit(`${c.name || ""} ${c.roleOrCompany || ""} ${c.notes || ""}`)) {
      results.push({ subTab: "contacts", title: c.name || "Contact", snippet: c.roleOrCompany || "" });
    }
  });
  data.winePricing.forEach((p) => {
    if (hit(`${p.wineName || ""} ${p.notes || ""}`)) {
      results.push({ subTab: "pricing", title: p.wineName || "Wine", snippet: `Wholesale $${p.wholesalePrice || "—"} · Retail $${p.retailPrice || "—"}` });
    }
  });
  data.wineClubTiers.forEach((t) => {
    if (hit(`${t.tierName || ""} ${t.benefits || ""} ${t.notes || ""}`)) {
      results.push({ subTab: "wineClub", title: t.tierName || "Tier", snippet: t.benefits || "" });
    }
  });
  return results;
}

function AboutAlloroTab({ data, setPrintJob, alloroStory, onUpdateStory, vineyardMapImage, onUpdateMapImage, onAddThoEntry, onUpdateThoEntry, onDeleteThoEntry, confirmAction, vineyardBlocks, onAddBlock }) {
  const [subTab, setSubTab] = useState("overview");
  const [viewSheetId, setViewSheetId] = useState(null);
  const [searchQuery, setSearchQuery] = useState("");
  const accoladeFields = SIMPLE_SECTIONS.find((s) => s.key === "accolades").fields;
  const sortedAccolades = [...data.accolades].sort((a, b) => (a.date < b.date ? 1 : -1));
  const contactFields = SIMPLE_SECTIONS.find((s) => s.key === "contacts").fields;
  const pricingFields = SIMPLE_SECTIONS.find((s) => s.key === "winePricing").fields;
  const wineClubFields = SIMPLE_SECTIONS.find((s) => s.key === "wineClubTiers").fields;
  const searchResults = searchTeamResources(searchQuery, data, alloroStory);

  const SUB_TABS = [
    { key: "overview", label: "Overview" },
    { key: "history", label: "History" },
    { key: "vineyard", label: "Vineyard" },
    { key: "accolades", label: "Accolades" },
    { key: "techSheets", label: "Tech Sheets" },
    { key: "contacts", label: "Contacts" },
    { key: "pricing", label: "Pricing" },
    { key: "wineClub", label: "Wine Club" },
  ];

  const goToResult = (result) => {
    setSubTab(result.subTab);
    if (result.sheetId) setViewSheetId(result.sheetId);
    setSearchQuery("");
  };

  return (
    <div>
      <div className="relative mb-4">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
        <input
          type="text"
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          placeholder="Search Team Resources — history, vineyard, accolades, tech sheets, contacts, pricing, wine club…"
          className="font-body w-full border border-stone-300 rounded-md pl-9 pr-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
      </div>

      {searchQuery.trim() ? (
        <div className="bg-white border border-stone-200 rounded-lg overflow-hidden">
          <div className="px-4 py-2.5 border-b border-stone-200">
            <p className="font-body text-xs text-stone-500">{searchResults.length} result{searchResults.length === 1 ? "" : "s"}</p>
          </div>
          {searchResults.length === 0 ? (
            <p className="font-body text-sm text-stone-400 text-center py-8">No matches.</p>
          ) : (
            <div className="divide-y divide-stone-100">
              {searchResults.map((r, i) => (
                <button key={i} onClick={() => goToResult(r)} className="w-full text-left px-4 py-3 hover:bg-stone-50">
                  <div className="flex items-center justify-between">
                    <p className="font-body text-sm font-medium text-stone-800">{r.title}</p>
                    <span className="font-body text-xs text-stone-400">{SUB_TABS.find((t) => t.key === r.subTab)?.label}</span>
                  </div>
                  {r.snippet && <p className="font-body text-xs text-stone-500 mt-0.5 truncate">{r.snippet}</p>}
                </button>
              ))}
            </div>
          )}
        </div>
      ) : (
        <>
          <div className="flex gap-2 mb-6 flex-wrap">
            {SUB_TABS.map((t) => (
              <button
                key={t.key}
                onClick={() => setSubTab(t.key)}
                className={`font-body text-sm font-medium px-4 py-2 rounded-md border ${
                  subTab === t.key ? "bg-ink-900 text-white border-ink-900" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                }`}
              >
                {t.label}
              </button>
            ))}
          </div>

          {subTab === "overview" ? (
            <CaseProductionSummary bottling={data.bottling} />
          ) : subTab === "history" ? (
            <HistoryPanel
              alloroStory={alloroStory}
              onUpdateStory={onUpdateStory}
              milestones={data.historyMilestones}
              onAddThoEntry={onAddThoEntry}
              onUpdateThoEntry={onUpdateThoEntry}
              onDeleteThoEntry={onDeleteThoEntry}
              confirmAction={confirmAction}
            />
          ) : subTab === "vineyard" ? (
            <VineyardPanel
              vineyardMapImage={vineyardMapImage}
              onUpdateMapImage={onUpdateMapImage}
              blockDetails={data.vineyardBlockDetails}
              onAddThoEntry={onAddThoEntry}
              onUpdateThoEntry={onUpdateThoEntry}
              onDeleteThoEntry={onDeleteThoEntry}
              confirmAction={confirmAction}
              vineyardBlocks={vineyardBlocks}
              onAddBlock={onAddBlock}
            />
          ) : subTab === "accolades" ? (
            <SimpleDataPanel
              title="Wine Accolades"
              fields={accoladeFields}
              rows={sortedAccolades}
              onAdd={(entry) => onAddThoEntry("accolades", entry)}
              onUpdate={(id, changes) => onUpdateThoEntry("accolades", id, changes)}
              onDelete={(id) => onDeleteThoEntry("accolades", id)}
              confirmAction={confirmAction}
            />
          ) : subTab === "contacts" ? (
            <SimpleDataPanel
              title="Contacts"
              fields={contactFields}
              rows={data.contacts}
              onAdd={(entry) => onAddThoEntry("contacts", entry)}
              onUpdate={(id, changes) => onUpdateThoEntry("contacts", id, changes)}
              onDelete={(id) => onDeleteThoEntry("contacts", id)}
              confirmAction={confirmAction}
            />
          ) : subTab === "pricing" ? (
            <SimpleDataPanel
              title="Pricing"
              fields={pricingFields}
              rows={data.winePricing}
              onAdd={(entry) => onAddThoEntry("winePricing", entry)}
              onUpdate={(id, changes) => onUpdateThoEntry("winePricing", id, changes)}
              onDelete={(id) => onDeleteThoEntry("winePricing", id)}
              confirmAction={confirmAction}
            />
          ) : subTab === "wineClub" ? (
            <SimpleDataPanel
              title="Wine Club Tiers"
              fields={wineClubFields}
              rows={data.wineClubTiers}
              onAdd={(entry) => onAddThoEntry("wineClubTiers", entry)}
              onUpdate={(id, changes) => onUpdateThoEntry("wineClubTiers", id, changes)}
              onDelete={(id) => onDeleteThoEntry("wineClubTiers", id)}
              confirmAction={confirmAction}
            />
          ) : (
            <TechSheetsReadOnlyPanel
              data={data}
              viewSheetId={viewSheetId}
              setViewSheetId={setViewSheetId}
              onPrint={(sheetId) => setPrintJob({ type: "techSheet", sheetId })}
            />
          )}
        </>
      )}
    </div>
  );
}

// ---------- Compliance: a tracker, not an authority. Only tracks what the winery tells it —
// their permits, their self-noted caps, their shipments — and reminds them of dates. Never
// asserts legality, never calculates tax owed, never files anything. ----------
const COMPLIANCE_DISCLAIMER =
  "This is a tracking tool based on information you provide. It is not legal, tax, or compliance advice. Verify current rules with your state's Alcohol Beverage Control agency before shipping.";

// Furthest-along status a permit can be in, computed purely from dates the winery entered —
// never asserts a shipment is legal, just watches the dates.
function computePermitStatus(permit) {
  if (!permit.expirationDate) return "Not Licensed";
  const today = todayISO();
  if (today >= permit.expirationDate) return "Expired";
  const leadDays = parseInt(permit.reminderLeadDays, 10) || 45;
  const exp = new Date(permit.expirationDate + "T00:00:00");
  const reminder = new Date(exp);
  reminder.setDate(reminder.getDate() - leadDays);
  const reminderISO = reminder.toISOString().slice(0, 10);
  if (today >= reminderISO) return "Expiring Soon";
  return "Active";
}

// Best (most current) status on file per state, since a state can have more than one permit
// record over time (e.g. an old expired one plus a fresh renewal).
function bestPermitStatusByState(permits) {
  const rank = { Active: 3, "Expiring Soon": 2, Expired: 1, "Not Licensed": 0 };
  const map = {};
  permits.forEach((p) => {
    if (!p.state) return;
    const status = computePermitStatus(p);
    if (!map[p.state] || rank[status] > rank[map[p.state]]) map[p.state] = status;
  });
  return map;
}

// Flags states with shipments logged but no currently-valid permit on file — a flag to
// investigate, not proof of an actual violation.
function complianceConflicts(shipments, permits) {
  const statusByState = bestPermitStatusByState(permits);
  const shippedStates = [...new Set(shipments.map((s) => s.destinationState).filter(Boolean))];
  return shippedStates.filter((state) => {
    const status = statusByState[state];
    return !status || status === "Expired" || status === "Not Licensed";
  });
}

function ComplianceOverview({ permits, shipments, setSubTab }) {
  const statusByState = bestPermitStatusByState(permits);
  const licensed = Object.entries(statusByState).filter(([, s]) => s === "Active").map(([st]) => st).sort();
  const expiringSoon = Object.entries(statusByState).filter(([, s]) => s === "Expiring Soon").map(([st]) => st).sort();
  const expired = Object.entries(statusByState).filter(([, s]) => s === "Expired").map(([st]) => st).sort();
  const conflicts = complianceConflicts(shipments, permits).sort();

  const Chip = ({ text, color }) => (
    <span className={`font-body text-xs rounded px-2 py-0.5 ${color}`}>{text}</span>
  );

  return (
    <div className="space-y-4">
      <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
        <p className="font-body text-xs text-amber-800">{COMPLIANCE_DISCLAIMER}</p>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <div className="bg-white border border-stone-200 rounded-lg p-4">
          <p className="font-body text-xs text-stone-500 mb-2">Licensed states</p>
          <p className="font-brand text-2xl text-ink-950 mb-2">{licensed.length}</p>
          <div className="flex flex-wrap gap-1">
            {licensed.length === 0 ? <span className="font-body text-xs text-stone-400">None yet</span> : licensed.map((s) => <Chip key={s} text={s} color="bg-green-50 text-green-700" />)}
          </div>
        </div>
        <div className="bg-white border border-stone-200 rounded-lg p-4">
          <p className="font-body text-xs text-stone-500 mb-2">Expiring soon</p>
          <p className="font-brand text-2xl text-amber-700 mb-2">{expiringSoon.length}</p>
          <div className="flex flex-wrap gap-1">
            {expiringSoon.length === 0 ? <span className="font-body text-xs text-stone-400">None</span> : expiringSoon.map((s) => <Chip key={s} text={s} color="bg-amber-50 text-amber-700" />)}
          </div>
        </div>
        <div className="bg-white border border-stone-200 rounded-lg p-4">
          <p className="font-body text-xs text-stone-500 mb-2">Needs attention</p>
          <p className="font-brand text-2xl text-red-700 mb-2">{expired.length + conflicts.length}</p>
          <div className="flex flex-wrap gap-1">
            {expired.length === 0 && conflicts.length === 0 ? (
              <span className="font-body text-xs text-stone-400">None</span>
            ) : (
              <>
                {expired.map((s) => <Chip key={"exp-" + s} text={`${s} — expired`} color="bg-red-50 text-red-700" />)}
                {conflicts.map((s) => <Chip key={"conf-" + s} text={`${s} — no active permit on file`} color="bg-red-50 text-red-700" />)}
              </>
            )}
          </div>
        </div>
      </div>

      <p className="font-body text-xs text-stone-400">
        "Needs attention" flags are worth checking, not proof of an actual violation — a conflict just means shipments were logged to a state without a currently-valid permit on file here.
        <button onClick={() => setSubTab("reference")} className="text-ink-700 underline ml-1">Browse all 50 states →</button>
      </p>
    </div>
  );
}

function ComplianceTab({ data, onAddThoEntry, onUpdateThoEntry, onDeleteThoEntry, confirmAction, associatesList, onAddAssociate }) {
  const [subTab, setSubTab] = useState("overview");
  const permitFields = SIMPLE_SECTIONS.find((s) => s.key === "statePermits").fields;
  const shipmentFields = SIMPLE_SECTIONS.find((s) => s.key === "shipmentLog").fields;
  const referenceFields = SIMPLE_SECTIONS.find((s) => s.key === "stateRuleReference").fields;

  const sortedPermits = [...data.statePermits].sort((a, b) => (a.expirationDate || "9999").localeCompare(b.expirationDate || "9999"));
  const sortedShipments = [...data.shipmentLog].sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  const sortedReference = [...data.stateRuleReference].sort((a, b) => (a.state || "").localeCompare(b.state || ""));

  const SUB_TABS = [
    { key: "overview", label: "Overview" },
    { key: "permits", label: "My Permits" },
    { key: "shipments", label: "Shipment Log" },
    { key: "reference", label: "State Reference" },
  ];

  return (
    <div>
      <div className="flex gap-2 mb-6 flex-wrap">
        {SUB_TABS.map((t) => (
          <button
            key={t.key}
            onClick={() => setSubTab(t.key)}
            className={`font-body text-sm font-medium px-4 py-2 rounded-md border ${
              subTab === t.key ? "bg-ink-900 text-white border-ink-900" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
            }`}
          >
            {t.label}
          </button>
        ))}
      </div>

      {subTab === "overview" ? (
        <ComplianceOverview permits={data.statePermits} shipments={data.shipmentLog} setSubTab={setSubTab} />
      ) : subTab === "permits" ? (
        <div className="space-y-4">
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
            <p className="font-body text-xs text-amber-800">{COMPLIANCE_DISCLAIMER}</p>
          </div>
          <SimpleDataPanel
            title="My Permits"
            fields={permitFields}
            rows={sortedPermits}
            onAdd={(entry) => onAddThoEntry("statePermits", entry)}
            onUpdate={(id, changes) => onUpdateThoEntry("statePermits", id, changes)}
            onDelete={(id) => onDeleteThoEntry("statePermits", id)}
            confirmAction={confirmAction}
          />
        </div>
      ) : subTab === "shipments" ? (
        <div className="space-y-4">
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
            <p className="font-body text-xs text-amber-800">{COMPLIANCE_DISCLAIMER}</p>
          </div>
          <SimpleDataPanel
            title="Shipment Log"
            fields={shipmentFields}
            rows={sortedShipments}
            onAdd={(entry) => onAddThoEntry("shipmentLog", entry)}
            onUpdate={(id, changes) => onUpdateThoEntry("shipmentLog", id, changes)}
            onDelete={(id) => onDeleteThoEntry("shipmentLog", id)}
            confirmAction={confirmAction}
          />
        </div>
      ) : (
        <div className="space-y-4">
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
            <p className="font-body text-xs text-amber-800">{COMPLIANCE_DISCLAIMER}</p>
          </div>
          <p className="font-body text-xs text-stone-500">
            All 50 states + DC are listed below. Most start as "Not Yet Verified" — this is curated public information, refreshed on a quarterly cadence, not original legal analysis. Update a state's status once you've actually verified it against a current source.
          </p>
          <SimpleDataPanel
            title="State Reference"
            fields={referenceFields}
            rows={sortedReference}
            onAdd={(entry) => onAddThoEntry("stateRuleReference", entry)}
            onUpdate={(id, changes) => onUpdateThoEntry("stateRuleReference", id, changes)}
            onDelete={(id) => onDeleteThoEntry("stateRuleReference", id)}
            confirmAction={confirmAction}
          />
        </div>
      )}
    </div>
  );
}

function THOPayoutCalculator({ data }) {
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [method, setMethod] = useState("hours");
  const [halfShareWeight, setHalfShareWeight] = useState("0.5");
  const [fullShareWeight, setFullShareWeight] = useState("1");

  const result = computeTipPayout(
    data.thoTimesheets,
    data.thoTips,
    startDate,
    endDate,
    method,
    parseFloat(halfShareWeight) || 0,
    parseFloat(fullShareWeight) || 0
  );

  const exportPayout = () => {
    const wb = XLSX.utils.book_new();
    const methodLabel = method === "shares" ? `Half/Full Day Shares (${halfShareWeight} / ${fullShareWeight})` : "By Hours Worked";
    const summaryAoa = [
      ["Alloro Winery Tracker — Tasting House Tip Payout"],
      ["Period:", `${startDate || "all dates"} to ${endDate || "all dates"}`],
      ["Method:", methodLabel],
      ["Total Tips in Period:", `$${result.totalTips.toFixed(2)}`],
      [],
      ["Associate", "Total Hours", "Payout ($)"],
      ...result.rows.map((r) => [r.employeeName, r.totalHours, r.totalPayout]),
      [],
      ["Total", result.rows.reduce((s, r) => s + r.totalHours, 0), result.rows.reduce((s, r) => s + r.totalPayout, 0)],
    ];
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryAoa);
    XLSX.utils.book_append_sheet(wb, summarySheet, "Payout Summary");

    const timesheetSheet = XLSX.utils.aoa_to_sheet([
      ["Date", "Associate", "Hours Worked"],
      ...data.thoTimesheets.filter((t) => (!startDate || t.date >= startDate) && (!endDate || t.date <= endDate)).map((t) => [t.date, t.employeeName, t.hoursWorked]),
    ]);
    XLSX.utils.book_append_sheet(wb, timesheetSheet, "Timesheets");

    const tipsSheet = XLSX.utils.aoa_to_sheet([
      ["Date", "Total Pooled Tips"],
      ...data.thoTips.filter((t) => (!startDate || t.date >= startDate) && (!endDate || t.date <= endDate)).map((t) => [t.date, t.totalTips]),
    ]);
    XLSX.utils.book_append_sheet(wb, tipsSheet, "Tips");

    XLSX.writeFile(wb, `tho-tip-payout-${startDate || "all"}-to-${endDate || "all"}.xlsx`);
  };

  return (
    <div className="space-y-4">
      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-1">Payout Calculator</h2>
        <p className="font-body text-xs text-stone-500 mb-4">
          Combines Timesheets and Tips for the period below into a payout per associate, split per day so a busy shift's tips don't get evenly blended with a slow one.
        </p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4">
          <div>
            <label className="font-body block text-xs font-medium text-stone-600 mb-1">Start Date</label>
            <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800" />
          </div>
          <div>
            <label className="font-body block text-xs font-medium text-stone-600 mb-1">End Date</label>
            <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800" />
          </div>
        </div>

        <label className="font-body block text-xs font-medium text-stone-600 mb-1">Distribution Method</label>
        <div className="flex gap-2 mb-3">
          <button
            onClick={() => setMethod("hours")}
            className={`font-body text-sm font-medium px-3 py-2 rounded-md border ${method === "hours" ? "bg-ink-900 text-white border-ink-900" : "bg-white text-stone-600 border-stone-300"}`}
          >
            By Hours Worked
          </button>
          <button
            onClick={() => setMethod("shares")}
            className={`font-body text-sm font-medium px-3 py-2 rounded-md border ${method === "shares" ? "bg-ink-900 text-white border-ink-900" : "bg-white text-stone-600 border-stone-300"}`}
          >
            Half-Day / Full-Day Shares
          </button>
        </div>

        {method === "shares" && (
          <div className="bg-stone-50 border border-stone-200 rounded-md p-3 mb-3">
            <p className="font-body text-xs text-stone-500 mb-2">Under 4 hours worked in a day = Half Day. 4+ hours = Full Day.</p>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label className="font-body block text-xs font-medium text-stone-600 mb-1">Half Day Share</label>
                <input type="number" step="0.1" value={halfShareWeight} onChange={(e) => setHalfShareWeight(e.target.value)} className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm" />
              </div>
              <div>
                <label className="font-body block text-xs font-medium text-stone-600 mb-1">Full Day Share</label>
                <input type="number" step="0.1" value={fullShareWeight} onChange={(e) => setFullShareWeight(e.target.value)} className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm" />
              </div>
            </div>
          </div>
        )}

        {result.warnings.length > 0 && (
          <div className="bg-amber-50 border border-amber-200 rounded-md p-3 mb-3">
            {result.warnings.map((w, i) => (
              <p key={i} className="font-body text-xs text-amber-800">{w}</p>
            ))}
          </div>
        )}

        <button
          onClick={exportPayout}
          disabled={result.rows.length === 0}
          className="font-body flex items-center gap-2 bg-ink-900 hover:bg-ink-800 disabled:opacity-40 text-white text-sm font-medium px-4 py-2 rounded-md"
        >
          <Download size={16} /> Export for Accounting
        </button>
      </div>

      <div className="bg-white border border-stone-200 rounded-lg overflow-hidden">
        <div className="px-4 py-3 border-b border-stone-200 flex items-center justify-between">
          <h2 className="font-brand text-lg text-ink-950">Payout Summary</h2>
          <span className="font-body text-xs text-stone-500">Total tips: ${result.totalTips.toFixed(2)}</span>
        </div>
        {result.rows.length === 0 ? (
          <p className="font-body text-sm text-stone-500 px-4 py-8 text-center">
            No overlapping Timesheets and Tips data for this period yet.
          </p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm font-body">
            <thead>
              <tr className="text-left text-stone-500 bg-stone-50">
                <th className="px-4 py-2">Associate</th>
                <th className="px-4 py-2">Total Hours</th>
                <th className="px-4 py-2">Payout</th>
              </tr>
            </thead>
            <tbody>
              {result.rows.map((r) => (
                <tr key={r.employeeName} className="border-t border-stone-100">
                  <td className="px-4 py-2 text-stone-800">{r.employeeName}</td>
                  <td className="px-4 py-2 text-stone-600">{r.totalHours}</td>
                  <td className="px-4 py-2 text-stone-800 font-medium">${r.totalPayout.toFixed(2)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          </div>
        )}
      </div>
    </div>
  );
}

const FERMENT_STAGES = ["Started", "Primary Complete", "ML In Progress", "ML Complete"];

// Furthest milestone a lot has reached, 0-3 — matches FERMENT_STAGES indices
function computeStageIndex(lot) {
  if (lot.mlStatus === "Complete") return 3;
  if (lot.mlStatus === "Inoculated" || lot.mlStatus === "In Progress") return 2;
  if (lot.status === "Complete") return 1;
  return 0;
}

function FermentStageStepper({ stageIndex }) {
  return (
    <div className="flex items-center">
      {FERMENT_STAGES.map((label, i) => (
        <Fragment key={label}>
          <div className="flex flex-col items-center" style={{ width: 0 }}>
            <div
              className={`w-3 h-3 rounded-full shrink-0 ${
                i < stageIndex ? "bg-ink-700" : i === stageIndex ? "bg-ink-700 ring-4 ring-ink-100" : "bg-stone-200"
              }`}
            />
            <span className={`font-body text-[10px] mt-1.5 text-center whitespace-nowrap ${i <= stageIndex ? "text-ink-800 font-medium" : "text-stone-400"}`}>
              {label}
            </span>
          </div>
          {i < FERMENT_STAGES.length - 1 && <div className={`flex-1 h-0.5 mx-1 mb-4 ${i < stageIndex ? "bg-ink-700" : "bg-stone-200"}`} />}
        </Fragment>
      ))}
    </div>
  );
}

function FermentOverviewCard({ lot, onClick }) {
  const stageIndex = computeStageIndex(lot);
  const latestReading = [...lot.readings].sort((a, b) => (a.date < b.date ? 1 : -1))[0];
  return (
    <button type="button" onClick={onClick} className="bg-white border border-stone-200 rounded-lg p-4 w-full text-left hover:border-ink-300 transition-colors">
      <div className="flex items-start justify-between gap-3 mb-4">
        <div>
          <p className="font-brand text-base text-ink-950">{lot.tankId || "Untitled"}</p>
          <p className="font-body text-xs text-stone-500">
            {[lot.vessel, lot.variety, lot.vintage].filter(Boolean).join(" · ") || "No details yet"}
          </p>
        </div>
        {latestReading && (
          <div className="text-right shrink-0">
            <p className="font-body text-xs text-stone-400">As of {latestReading.date}</p>
            <p className="font-body text-sm text-stone-700">
              {latestReading.brix !== "" && latestReading.brix != null ? `${latestReading.brix}° Brix` : ""}
              {latestReading.temp !== "" && latestReading.temp != null ? ` · ${latestReading.temp}°F` : ""}
            </p>
          </div>
        )}
      </div>
      <FermentStageStepper stageIndex={stageIndex} />
      <div className="flex flex-wrap gap-x-3 gap-y-1 mt-4 font-body text-xs text-stone-500">
        {lot.startDate && <span>Started {lot.startDate}</span>}
        {lot.dateCompleted && <span>· Primary done {lot.dateCompleted}</span>}
        {lot.mlInoculationDate && <span>· ML started {lot.mlInoculationDate}</span>}
        {lot.mlCompleteDate && <span>· ML done {lot.mlCompleteDate}</span>}
      </div>
    </button>
  );
}

function FermentOverview({ lots, onSelectLot }) {
  const currentLots = lots.filter((lot) => computeStageIndex(lot) < 3);
  return (
    <div className="space-y-3 mb-6">
      <div>
        <h2 className="font-brand text-xl text-ink-950">Fermentation Overview</h2>
        <p className="font-body text-xs text-stone-500">
          Every lot still moving through primary or malolactic fermentation. Tap a card to open its details.
        </p>
      </div>
      {currentLots.length === 0 ? (
        <p className="font-body text-sm text-stone-500 bg-white border border-stone-200 rounded-lg p-6 text-center">
          No ferments currently in progress.
        </p>
      ) : (
        currentLots.map((lot) => <FermentOverviewCard key={lot.id} lot={lot} onClick={() => onSelectLot(lot.id)} />)
      )}
    </div>
  );
}

// ---------- Quick Log: fast, walk-down-the-list daily entry mirroring the paper log ----------
function QuickFermentLog({ lots, allLots, onSaveRow, onEndDay, onDeleteLot, onSwitchToDetailed, confirmAction }) {
  const [logDate, setLogDate] = useState(todayISO());
  const [session, setSession] = useState(guessSession());
  const [rowInputs, setRowInputs] = useState({});
  const [justSaved, setJustSaved] = useState({});

  const isToday = logDate === todayISO();
  // Today (and any future date) shows the working list of active ferments. Any other date shows
  // whichever lots actually have a reading logged that day — active or since completed — so
  // past days stay fully readable even after a lot finishes.
  const visibleLots = isToday || logDate > todayISO() ? lots : allLots.filter((lot) => lot.readings.some((r) => r.date === logDate));

  const existingFor = (lot) => lot.readings.find((r) => r.date === logDate && r.session === session);

  const rowValue = (lot) => {
    if (rowInputs[lot.id]) return rowInputs[lot.id];
    const existing = existingFor(lot);
    return {
      temp: existing?.temp ?? "",
      brix: existing?.brix ?? "",
      ph: existing?.ph ?? "",
      workDone: existing?.workDone || [],
      additions: existing?.additions || [],
      notes: existing?.notes ?? "",
    };
  };

  const updateField = (lot, field, val) => {
    setRowInputs((prev) => ({ ...prev, [lot.id]: { ...rowValue(lot), [field]: val } }));
  };
  const toggleWorkDone = (lot, item) => {
    const current = rowValue(lot);
    updateField(lot, "workDone", current.workDone.includes(item) ? current.workDone.filter((w) => w !== item) : [...current.workDone, item]);
  };
  const toggleAddition = (lot, item) => {
    const current = rowValue(lot);
    updateField(lot, "additions", current.additions.includes(item) ? current.additions.filter((a) => a !== item) : [...current.additions, item]);
  };

  const saveRow = (lot) => {
    onSaveRow(lot, logDate, session, rowValue(lot));
    setJustSaved((prev) => ({ ...prev, [lot.id]: true }));
    setTimeout(() => setJustSaved((prev) => ({ ...prev, [lot.id]: false })), 1500);
  };

  const endDay = () => {
    onEndDay(logDate);
    const next = new Date(logDate + "T00:00:00");
    next.setDate(next.getDate() + 1);
    setLogDate(next.toISOString().split("T")[0]);
    setSession("A.M.");
    setRowInputs({});
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg overflow-hidden mb-6">
      <div className="px-4 py-3 border-b border-stone-200 flex items-center justify-between flex-wrap gap-2">
        <div>
          <h2 className="font-brand text-lg text-ink-950">Quick Log</h2>
          <p className="font-body text-xs text-stone-500">
            {isToday ? "Walk down the list, log this round, move to the next." : "Viewing a past day — showing every lot with an entry that day."}
          </p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <input
            type="date"
            value={logDate}
            onChange={(e) => setLogDate(e.target.value)}
            className="font-body text-xs border border-stone-300 rounded px-2 py-1.5"
          />
          <div className="flex rounded-md border border-stone-300 overflow-hidden">
            {FERMENT_SESSIONS.map((s) => (
              <button
                key={s}
                onClick={() => setSession(s)}
                className={`font-body text-xs px-2.5 py-1.5 ${session === s ? "bg-ink-900 text-white" : "bg-white text-stone-600 hover:bg-stone-50"}`}
              >
                {s}
              </button>
            ))}
          </div>
          {isToday && (
            <button
              onClick={endDay}
              className="font-body flex items-center gap-1.5 text-xs font-medium bg-ink-900 hover:bg-ink-800 text-white px-3 py-1.5 rounded-md"
            >
              <ArchiveRestore size={13} /> End Day
            </button>
          )}
        </div>
      </div>
      {visibleLots.length === 0 ? (
        <p className="font-body text-sm text-stone-500 px-4 py-8 text-center">
          {isToday ? "No active ferments to log right now." : "No entries logged on this date."}
        </p>
      ) : (
        <div className="divide-y divide-stone-100">
          {visibleLots.map((lot) => {
            const v = rowValue(lot);
            const existing = existingFor(lot);
            return (
              <div key={lot.id} className="p-3">
                <div className="flex items-center justify-between mb-2">
                  <div>
                    <span className="font-body text-sm font-semibold text-stone-800">{lot.tankId || "Untitled"}</span>
                    {lot.vessel && <span className="font-body text-xs text-stone-400 ml-2">{lot.vessel}</span>}
                    {lot.variety && <span className="font-body text-xs text-stone-400 ml-2">· {lot.variety}</span>}
                  </div>
                  <div className="flex items-center gap-2">
                    {existing && <span className="font-body text-xs text-ink-700">Already logged this round</span>}
                    <button
                      onClick={() => onSwitchToDetailed(lot.id)}
                      className="text-stone-400 hover:text-ink-800"
                      title="Edit lot details (opens Detailed View)"
                    >
                      <Pencil size={14} />
                    </button>
                    <button
                      onClick={() =>
                        confirmAction("Delete this fermentation lot and all its readings? This can't be undone.", () => onDeleteLot(lot.id))
                      }
                      className="text-stone-400 hover:text-red-700"
                      title="Delete lot"
                    >
                      <Trash2 size={14} />
                    </button>
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 mb-2">
                  <input
                    type="number"
                    placeholder="Temp °F"
                    value={v.temp}
                    onChange={(e) => updateField(lot, "temp", e.target.value)}
                    className="font-body border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                  />
                  <input
                    type="number"
                    placeholder="Brix"
                    value={v.brix}
                    onChange={(e) => updateField(lot, "brix", e.target.value)}
                    className="font-body border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                  />
                  <input
                    type="number"
                    placeholder="pH"
                    value={v.ph}
                    onChange={(e) => updateField(lot, "ph", e.target.value)}
                    className="font-body border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                  />
                </div>
                <div className="flex flex-wrap gap-1.5 mb-2">
                  <button
                    type="button"
                    onClick={() => toggleWorkDone(lot, "Pump Over")}
                    className={`font-body text-xs px-2.5 py-1.5 rounded-md border ${
                      v.workDone.includes("Pump Over") ? "bg-ink-800 text-white border-ink-800" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                    }`}
                  >
                    O — Pump Over
                  </button>
                  <button
                    type="button"
                    onClick={() => toggleWorkDone(lot, "Punch Down")}
                    className={`font-body text-xs px-2.5 py-1.5 rounded-md border ${
                      v.workDone.includes("Punch Down") ? "bg-ink-800 text-white border-ink-800" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                    }`}
                  >
                    D — Punch Down
                  </button>
                  <button
                    type="button"
                    onClick={() => toggleWorkDone(lot, "Cold Soak")}
                    className={`font-body text-xs px-2.5 py-1.5 rounded-md border ${
                      v.workDone.includes("Cold Soak") ? "bg-ink-800 text-white border-ink-800" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                    }`}
                  >
                    Cold Soak
                  </button>
                </div>
                <div className="flex flex-wrap gap-1.5 mb-2">
                  {ADDITION_TYPES.map((a) => (
                    <button
                      key={a}
                      type="button"
                      onClick={() => toggleAddition(lot, a)}
                      className={`font-body text-xs px-2.5 py-1 rounded-full border transition-colors ${
                        v.additions.includes(a) ? "bg-amber-600 text-white border-amber-600" : "bg-white text-stone-600 border-stone-300 hover:border-amber-400"
                      }`}
                    >
                      + {a}
                    </button>
                  ))}
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="text"
                    placeholder="Notes"
                    value={v.notes}
                    onChange={(e) => updateField(lot, "notes", e.target.value)}
                    className="font-body flex-1 border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                  />
                  <button
                    onClick={() => saveRow(lot)}
                    className={`font-body text-sm font-medium px-3 py-1.5 rounded-md text-white shrink-0 ${justSaved[lot.id] ? "bg-ink-700" : "bg-ink-900 hover:bg-ink-800"}`}
                  >
                    {justSaved[lot.id] ? "✓ Saved" : "Log"}
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function FermentLotCard({ lot, onAddReading, onToggleComplete, onDeleteLot, onUpdateLot, onUpdateReading, onDeleteReading, confirmAction, harvestEntries, vesselTypesList, onAddVesselType }) {
  const [expanded, setExpanded] = useState(lot.status === "Active");
  const [readingForm, setReadingForm] = useState({ ...emptyForm(FERMENT_READING_FIELDS), date: todayISO() });
  const [error, setError] = useState("");
  const [editingLot, setEditingLot] = useState(false);
  const [lotEditForm, setLotEditForm] = useState(null);
  const [editingReadingId, setEditingReadingId] = useState(null);
  const [readingEditForm, setReadingEditForm] = useState(null);
  const [editingML, setEditingML] = useState(false);
  const [mlForm, setMlForm] = useState(null);

  const startEditML = () => {
    setMlForm({
      mlStatus: lot.mlStatus || "Not Started",
      mlInoculationDate: lot.mlInoculationDate || "",
      mlCompleteDate: lot.mlCompleteDate || "",
      mlNotes: lot.mlNotes || "",
    });
    setEditingML(true);
  };
  const saveEditML = () => {
    onUpdateLot(lot.id, mlForm);
    setEditingML(false);
  };

  const handleChange = (name, value) => setReadingForm((prev) => ({ ...prev, [name]: value }));

  const submitReading = (e) => {
    e.preventDefault();
    if (!readingForm.date) {
      setError("Date is required");
      return;
    }
    setError("");
    onAddReading(lot.id, { id: genId(), ...readingForm });
    setReadingForm({ ...emptyForm(FERMENT_READING_FIELDS), date: todayISO() });
  };

  const startEditLot = () => {
    setLotEditForm({
      tankId: lot.tankId,
      vessel: lot.vessel,
      variety: lot.variety,
      startDate: lot.startDate,
      startingBrix: lot.startingBrix,
      startingTemp: lot.startingTemp,
      notes: lot.notes,
    });
    setEditingLot(true);
  };
  const saveEditLot = () => {
    if (!lotEditForm.tankId.trim()) {
      setError('Please fill in "Tank / Lot ID"');
      return;
    }
    setError("");
    onUpdateLot(lot.id, lotEditForm);
    setEditingLot(false);
  };

  const startEditReading = (r) => {
    setEditingReadingId(r.id);
    setReadingEditForm({ ...r });
  };
  const saveEditReading = () => {
    onUpdateReading(lot.id, editingReadingId, readingEditForm);
    setEditingReadingId(null);
    setReadingEditForm(null);
  };
  const handleDeleteReading = (readingId) => {
    confirmAction("Delete this reading? This can't be undone.", () => onDeleteReading(lot.id, readingId));
  };

  const readings = [...lot.readings].sort((a, b) => (a.date < b.date ? 1 : -1));
  const readingsAsc = [...lot.readings].sort((a, b) => (a.date > b.date ? 1 : -1));
  const stall = lot.status === "Active" ? detectStall(readingsAsc) : null;

  return (
    <div className="bg-white border border-stone-200 rounded-lg overflow-hidden mb-4">
      <button
        onClick={() => setExpanded((v) => !v)}
        className="w-full flex items-center justify-between px-4 py-3 hover:bg-stone-50"
      >
        <div className="flex items-center gap-3 text-left">
          <span
            className={`text-xs font-body font-semibold px-2 py-0.5 rounded-full ${
              lot.status === "Active" ? "bg-green-100 text-green-800" : "bg-stone-200 text-stone-600"
            }`}
          >
            {lot.status}
          </span>
          <div>
            <p className="font-brand text-ink-950 leading-tight">
              {lot.tankId || "Untitled Tank"} {lot.variety && `— ${lot.variety}`}
            </p>
            <p className="font-body text-xs text-stone-500">
              {lot.vessel ? `${lot.vessel} · ` : ""}Started {lot.startDate || "—"} · {lot.readings.length} reading{lot.readings.length === 1 ? "" : "s"}
              {(lot.startingBrix || lot.startingTemp) && (
                <> · Start: {lot.startingBrix ? `${lot.startingBrix}° Brix` : ""}{lot.startingBrix && lot.startingTemp ? ", " : ""}{lot.startingTemp ? `${lot.startingTemp}°F` : ""}</>
              )}
            </p>
            {lot.mlStatus && lot.mlStatus !== "Not Started" && (
              <span
                className={`inline-block font-body text-xs px-1.5 py-0.5 rounded mt-1 ${
                  lot.mlStatus === "Complete" ? "bg-ink-100 text-ink-800" : "bg-purple-100 text-purple-800"
                }`}
              >
                ML: {lot.mlStatus}
              </span>
            )}
            {readings[0] && (
              <div className="flex flex-wrap items-center gap-1.5 mt-1">
                <span className="font-body text-xs text-stone-400">Last: {readings[0].date}</span>
                {Array.isArray(readings[0].workDone) &&
                  readings[0].workDone.map((w) => (
                    <span key={w} className="font-body text-xs bg-amber-100 text-amber-800 px-1.5 py-0.5 rounded">
                      {w}
                    </span>
                  ))}
                {readings[0].brix && (
                  <span className="font-body text-xs text-stone-500">{readings[0].brix}° Brix</span>
                )}
                {readings[0].temp && (
                  <span className="font-body text-xs text-stone-500">{readings[0].temp}°F</span>
                )}
              </div>
            )}
          </div>
        </div>
        {expanded ? <ChevronUp size={18} className="text-stone-400" /> : <ChevronDown size={18} className="text-stone-400" />}
      </button>

      {expanded && (
        <div className="border-t border-stone-200 px-4 py-4">
          {editingLot ? (
            <div className="bg-ink-50 border border-ink-100 rounded-md p-3 mb-4">
              <p className="font-body text-xs font-semibold text-stone-600 mb-2">Edit ferment details</p>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
                {FERMENT_LOT_FIELDS.map((f) => (
                  <Field key={f.name} f={f} value={lotEditForm[f.name]} onChange={(v) => setLotEditForm((p) => ({ ...p, [f.name]: v }))} vesselTypesList={vesselTypesList} onAddVesselType={onAddVesselType} />
                ))}
              </div>
              <div className="flex items-center gap-3">
                <button onClick={saveEditLot} className="font-body flex items-center gap-1 text-sm font-medium text-ink-800 hover:text-ink-900">
                  <Check size={15} /> Save
                </button>
                <button onClick={() => setEditingLot(false)} className="font-body flex items-center gap-1 text-sm text-stone-500 hover:text-stone-700">
                  <X size={15} /> Cancel
                </button>
              </div>
            </div>
          ) : (
            lot.notes && <p className="font-body text-xs text-stone-500 mb-3 italic">Starting notes: {lot.notes}</p>
          )}

          <div className="mb-4">
            <p className="font-body text-xs font-semibold text-stone-600 mb-2">Sourced From (Harvest Picks)</p>
            <HarvestRefsPicker
              value={lot.harvestRefs || []}
              onChange={(refs) => onUpdateLot(lot.id, { harvestRefs: refs })}
              harvestEntries={harvestEntries || []}
            />
          </div>

          <div className="bg-purple-50 border border-purple-100 rounded-md p-3 mb-4">
            <div className="flex items-center justify-between mb-2">
              <p className="font-body text-xs font-semibold text-stone-600">Malolactic Fermentation</p>
              {!editingML && (
                <button onClick={startEditML} className="font-body flex items-center gap-1 text-xs font-medium text-purple-800 hover:text-purple-900">
                  <Pencil size={12} /> {lot.mlStatus ? "Edit" : "Track ML"}
                </button>
              )}
            </div>
            {editingML ? (
              <>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 mb-2">
                  <div>
                    <label className="font-body block text-xs font-medium text-stone-600 mb-1">ML Status</label>
                    <select
                      value={mlForm.mlStatus}
                      onChange={(e) => setMlForm((p) => ({ ...p, mlStatus: e.target.value }))}
                      className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                    >
                      {ML_STATUSES.map((s) => (
                        <option key={s} value={s}>{s}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="font-body block text-xs font-medium text-stone-600 mb-1">Inoculation Date</label>
                    <input
                      type="date"
                      value={mlForm.mlInoculationDate}
                      onChange={(e) => setMlForm((p) => ({ ...p, mlInoculationDate: e.target.value }))}
                      className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                    />
                  </div>
                  <div>
                    <label className="font-body block text-xs font-medium text-stone-600 mb-1">Complete Date</label>
                    <input
                      type="date"
                      value={mlForm.mlCompleteDate}
                      onChange={(e) => setMlForm((p) => ({ ...p, mlCompleteDate: e.target.value }))}
                      className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                    />
                  </div>
                </div>
                <label className="font-body block text-xs font-medium text-stone-600 mb-1">
                  Notes (chromatography checks, ML culture used, etc.)
                </label>
                <textarea
                  value={mlForm.mlNotes}
                  onChange={(e) => setMlForm((p) => ({ ...p, mlNotes: e.target.value }))}
                  rows={2}
                  className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm mb-2 focus:outline-none focus:ring-2 focus:ring-ink-800"
                />
                <div className="flex items-center gap-3">
                  <button onClick={saveEditML} className="font-body flex items-center gap-1 text-sm font-medium text-purple-800 hover:text-purple-900">
                    <Check size={15} /> Save
                  </button>
                  <button onClick={() => setEditingML(false)} className="font-body flex items-center gap-1 text-sm text-stone-500 hover:text-stone-700">
                    <X size={15} /> Cancel
                  </button>
                </div>
              </>
            ) : lot.mlStatus ? (
              <p className="font-body text-xs text-stone-600">
                Status: <span className="font-medium">{lot.mlStatus}</span>
                {lot.mlInoculationDate && <> · Inoculated {lot.mlInoculationDate}</>}
                {lot.mlCompleteDate && <> · Complete {lot.mlCompleteDate}</>}
                {lot.mlNotes && <span className="block text-stone-500 mt-1">{lot.mlNotes}</span>}
              </p>
            ) : (
              <p className="font-body text-xs text-stone-400">Not tracked yet for this lot.</p>
            )}
          </div>

          <FermentationChart readings={readingsAsc} />

          {stall && (
            <div className="bg-amber-50 border border-amber-200 rounded-md p-3 mb-4 flex items-start gap-2">
              <span className="text-base leading-none">⚠️</span>
              <p className="font-body text-xs text-amber-800">
                Brix barely moved between {stall.prevDate} and {stall.lastDate} ({stall.drop.toFixed(1)}° drop) — worth checking for a stalled fermentation.
              </p>
            </div>
          )}

          <form onSubmit={submitReading} className="bg-stone-50 border border-stone-200 rounded-md p-3 mb-4">
            <p className="font-body text-xs font-semibold text-stone-600 mb-2">Log a reading</p>
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 items-end">
              {FERMENT_READING_FIELDS.map((f) => (
                <Field key={f.name} f={f} value={readingForm[f.name]} onChange={(v) => handleChange(f.name, v)} />
              ))}
              <button
                type="submit"
                className="font-body flex items-center justify-center gap-1 bg-ink-900 hover:bg-ink-800 text-white text-sm font-medium px-3 py-2 rounded-md h-fit"
              >
                <Plus size={15} /> Add
              </button>
            </div>
            {error && <p className="font-body text-xs text-red-700 mt-2">{error}</p>}
          </form>

          {readings.length === 0 ? (
            <p className="font-body text-sm text-stone-500 text-center py-4">No readings logged yet.</p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm font-body">
                <thead>
                  <tr className="bg-stone-100 text-stone-600 text-left">
                    <th className="px-3 py-2 font-medium whitespace-nowrap">Date</th>
                    <th className="px-3 py-2 font-medium whitespace-nowrap">Session</th>
                    <th className="px-3 py-2 font-medium whitespace-nowrap">Work Done</th>
                    <th className="px-3 py-2 font-medium whitespace-nowrap">Additions</th>
                    <th className="px-3 py-2 font-medium whitespace-nowrap">Brix</th>
                    <th className="px-3 py-2 font-medium whitespace-nowrap">Temp (°F)</th>
                    <th className="px-3 py-2 font-medium whitespace-nowrap">pH</th>
                    <th className="px-3 py-2 font-medium whitespace-nowrap">Notes</th>
                    <th className="px-3 py-2 w-16"></th>
                  </tr>
                </thead>
                <tbody>
                  {readings.map((r) =>
                    editingReadingId === r.id ? (
                      <tr key={r.id} className="border-t border-stone-100 bg-ink-50">
                        {FERMENT_READING_FIELDS.map((f) => (
                          <td key={f.name} className="px-2 py-2 align-top" style={{ minWidth: 110 }}>
                            <Field f={f} value={readingEditForm[f.name]} onChange={(v) => setReadingEditForm((p) => ({ ...p, [f.name]: v }))} hideLabel />
                          </td>
                        ))}
                        <td className="px-2 py-2">
                          <div className="flex items-center gap-2">
                            <button onClick={saveEditReading} className="text-ink-700 hover:text-ink-900" title="Save">
                              <Check size={15} />
                            </button>
                            <button onClick={() => setEditingReadingId(null)} className="text-stone-400 hover:text-stone-600" title="Cancel">
                              <X size={15} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ) : (
                      <tr key={r.id} className="border-t border-stone-100">
                        <td className="px-3 py-2 whitespace-nowrap">{r.date}</td>
                        <td className="px-3 py-2 whitespace-nowrap">{r.session || "—"}</td>
                        <td className="px-3 py-2 whitespace-nowrap">
                          {Array.isArray(r.workDone) && r.workDone.length > 0 ? r.workDone.join(", ") : "—"}
                        </td>
                        <td className="px-3 py-2 whitespace-nowrap">
                          {Array.isArray(r.additions) && r.additions.length > 0 ? r.additions.join(", ") : "—"}
                        </td>
                        <td className="px-3 py-2 whitespace-nowrap">{r.brix || "—"}</td>
                        <td className="px-3 py-2 whitespace-nowrap">{r.temp || "—"}</td>
                        <td className="px-3 py-2 whitespace-nowrap">{r.ph || "—"}</td>
                        <td className="px-3 py-2 max-w-xs truncate">{r.notes || "—"}</td>
                        <td className="px-3 py-2">
                          <div className="flex items-center gap-2">
                            <button onClick={() => startEditReading(r)} className="text-stone-400 hover:text-ink-800" title="Edit">
                              <Pencil size={14} />
                            </button>
                            <button onClick={() => handleDeleteReading(r.id)} className="text-stone-400 hover:text-red-700" title="Delete">
                              <Trash2 size={14} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    )
                  )}
                </tbody>
              </table>
            </div>
          )}

          <div className="flex items-center gap-3 mt-4 pt-3 border-t border-stone-100">
            <button
              onClick={() => onToggleComplete(lot.id)}
              className="font-body flex items-center gap-1.5 text-sm font-medium px-3 py-1.5 rounded-md border border-stone-300 hover:bg-stone-50 text-stone-700"
            >
              {lot.status === "Active" ? <CheckCircle2 size={15} /> : <RotateCcw size={15} />}
              {lot.status === "Active" ? "Mark Complete" : "Reopen Ferment"}
            </button>
            {!editingLot && (
              <button
                onClick={startEditLot}
                className="font-body flex items-center gap-1.5 text-sm font-medium px-3 py-1.5 rounded-md border border-stone-300 hover:bg-stone-50 text-stone-700"
              >
                <Pencil size={14} /> Edit Details
              </button>
            )}
            {lot.status === "Complete" && (
              <button
                onClick={() => onUpdateLot(lot.id, { archived: !lot.archived })}
                className="font-body flex items-center gap-1.5 text-sm font-medium px-3 py-1.5 rounded-md border border-stone-300 hover:bg-stone-50 text-stone-700"
              >
                {lot.archived ? <ArchiveRestore size={14} /> : <Archive size={14} />}
                {lot.archived ? "Unarchive" : "Archive"}
              </button>
            )}
            <button
              onClick={() => confirmAction("Delete this fermentation lot and all its readings? This can't be undone.", () => onDeleteLot(lot.id))}
              className="font-body flex items-center gap-1.5 text-sm text-stone-400 hover:text-red-700 ml-auto"
            >
              <Trash2 size={15} /> Delete Ferment
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ---------- One collapsible vintage year inside the Fermentation archive ----------
function FermentVintageGroup({ label, lots, onAddReading, onToggleComplete, onDeleteLot, onUpdateLot, onUpdateReading, onDeleteReading, confirmAction, harvestEntries, vesselTypesList, onAddVesselType }) {
  const [expanded, setExpanded] = useState(false);
  return (
    <div className="border-b border-stone-100 last:border-b-0">
      <button onClick={() => setExpanded((v) => !v)} className="w-full flex items-center justify-between px-4 py-3 hover:bg-stone-50">
        <span className="font-body text-sm font-medium text-stone-700">{label}</span>
        <span className="flex items-center gap-2">
          <span className="font-body text-xs text-stone-400">{lots.length} lot{lots.length === 1 ? "" : "s"}</span>
          {expanded ? <ChevronUp size={16} className="text-stone-400" /> : <ChevronDown size={16} className="text-stone-400" />}
        </span>
      </button>
      {expanded && (
        <div className="px-4 pb-4">
          {lots.map((lot) => (
            <FermentLotCard
              key={lot.id}
              lot={lot}
              onAddReading={onAddReading}
              onToggleComplete={onToggleComplete}
              onDeleteLot={onDeleteLot}
              onUpdateLot={onUpdateLot}
              onUpdateReading={onUpdateReading}
              onDeleteReading={onDeleteReading}
              confirmAction={confirmAction}
              harvestEntries={harvestEntries}
              vesselTypesList={vesselTypesList}
              onAddVesselType={onAddVesselType}
            />
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- Small labeled number input used by the calculators below ----------
function CalcInput({ label, value, onChange, unit }) {
  return (
    <div>
      <label className="font-body block text-xs font-medium text-stone-600 mb-1">{label}</label>
      <div className="relative">
        <input
          type="number"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
        {unit && <span className="font-body absolute right-3 top-1/2 -translate-y-1/2 text-xs text-stone-400">{unit}</span>}
      </div>
    </div>
  );
}

function CalcResult({ label, value, unit }) {
  return (
    <div className="bg-ink-50 border border-ink-100 rounded-md px-3 py-2">
      <p className="font-body text-xs text-ink-800">{label}</p>
      <p className="font-brand text-xl text-ink-950">{value !== "" ? value : "—"}{value !== "" && unit ? ` ${unit}` : ""}</p>
    </div>
  );
}

function CalcCard({ title, children }) {
  return (
    <div className="border border-stone-200 rounded-lg p-4">
      <h3 className="font-body text-sm font-semibold text-stone-700 mb-3">{title}</h3>
      {children}
    </div>
  );
}

// Reference formulas shown as plain text/read-only cards
const REFERENCE_FORMULAS = [
  {
    title: "Potential Alcohol",
    formula: "Potential ABV (%) = Brix × 0.55",
    note: "Rule-of-thumb conversion factor; actual yeast efficiency typically runs 0.55–0.64 depending on strain and fermentation conditions.",
  },
  {
    title: "Titratable Acidity (TA)",
    formula: "TA (g/L as tartaric) = (mL NaOH × Normality × 75) ÷ mL sample",
    note: "Standard titration formula using 0.1N NaOH and a 10 mL sample; multiply result accordingly if using different volumes.",
  },
  {
    title: "Bound SO₂",
    formula: "Bound SO₂ = Total SO₂ − Free SO₂",
    note: "Useful for tracking how much SO₂ has combined with wine compounds versus what remains protectively active.",
  },
  {
    title: "Racking Loss",
    formula: "Loss (%) = (Volume Before − Volume After) ÷ Volume Before × 100",
    note: "Typical racking loss runs 5–10% depending on lees volume and vessel geometry.",
  },
  {
    title: "Yield Estimate",
    formula: "1 ton grapes ≈ 150 gallons wine ≈ 60 cases (720 bottles)",
    note: "Approximate — actual yield depends on press efficiency, variety, and juice-to-skin ratio.",
  },
  {
    title: "Barrel Volume",
    formula: "1 standard barrel = 59.6 gal = 225 L ≈ 24 cases",
    note: "Based on a standard 225 L Bordeaux-style barrel; puncheons and other formats vary.",
  },
];

// Tom's own formulas — empty for now, filled in later
const TOMS_FORMULAS = [
  {
    title: "SO₂ Addition — 10% Liquid Solution",
    formula: "mL = [(Target − Current) mg/L × Volume(L)] ÷ (172 × 0.576)",
    note: "Tom's method, using a premixed 10% KMBS stock solution (172g KMBS per 1000mL, KMBS is 57.6% available SO2 by weight) — dosed in mL rather than grams of powder. Also available as an auto-calculator when building a Work Order: set Task Type to \"Additions\" and Addition Type to \"SO2\" and the Directions box will fill in for you.",
  },
];

const MOLECULAR_SO2_TABLE = [
  { ph: "3.0", freeSO2: "8 ppm" },
  { ph: "3.2", freeSO2: "13 ppm" },
  { ph: "3.4", freeSO2: "20 ppm" },
  { ph: "3.6", freeSO2: "32 ppm" },
  { ph: "3.8", freeSO2: "50 ppm" },
  { ph: "4.0", freeSO2: "80 ppm" },
];

const BRIX_SG_TABLE = [
  { brix: "18", sg: "1.074" },
  { brix: "20", sg: "1.083" },
  { brix: "22", sg: "1.091" },
  { brix: "24", sg: "1.100" },
  { brix: "26", sg: "1.109" },
];

const UNIT_CONVERSIONS = [
  { category: "Volume", rows: ["1 gallon = 3.785 liters", "1 barrel = 59.6 gal = 225 L", "1 case = 2.38 gal = 9 L"] },
  { category: "Weight", rows: ["1 ton = 2,000 lbs = 907.18 kg", "1 lb = 453.6 g", "1 kg = 2.2046 lbs"] },
  { category: "Temperature", rows: ["°F = (°C × 9/5) + 32", "°C = (°F − 32) × 5/9"] },
  { category: "Area", rows: ["1 acre = 0.4047 hectares", "1 hectare = 2.4711 acres"] },
  { category: "Concentration", rows: ["1 ppm = 1 mg/L (dilute aqueous solutions)"] },
];

function FormulasTab() {
  const [brix, setBrix] = useState("");
  const potentialAlcohol = brix !== "" && !isNaN(parseFloat(brix)) ? (parseFloat(brix) * 0.55).toFixed(2) : "";

  const [so2Current, setSo2Current] = useState("");
  const [so2Target, setSo2Target] = useState("");
  const [so2Volume, setSo2Volume] = useState("");
  const kmsGrams = (() => {
    const cur = parseFloat(so2Current), tgt = parseFloat(so2Target), vol = parseFloat(so2Volume);
    if (isNaN(cur) || isNaN(tgt) || isNaN(vol) || tgt <= cur || vol <= 0) return "";
    const liters = vol * 3.78541;
    const gramsSO2 = ((tgt - cur) * liters) / 1000;
    return (gramsSO2 / 0.57).toFixed(1);
  })();

  const [tomSo2Current, setTomSo2Current] = useState("");
  const [tomSo2Target, setTomSo2Target] = useState("");
  const [tomSo2Volume, setTomSo2Volume] = useState("");
  const tomSo2Result = calcSO2StockML(tomSo2Volume, tomSo2Current, tomSo2Target);
  const tomSo2ML = tomSo2Result ? tomSo2Result.mL.toFixed(1) : "";

  const [chaptBrix, setChaptBrix] = useState("");
  const [chaptVolume, setChaptVolume] = useState("");
  const sugarLbs = (() => {
    const inc = parseFloat(chaptBrix), vol = parseFloat(chaptVolume);
    if (isNaN(inc) || isNaN(vol) || inc <= 0 || vol <= 0) return "";
    return (((inc * vol) / 100) * 1.6).toFixed(1);
  })();

  const [acidCurrent, setAcidCurrent] = useState("");
  const [acidTarget, setAcidTarget] = useState("");
  const [acidVolume, setAcidVolume] = useState("");
  const tartaricGrams = (() => {
    const cur = parseFloat(acidCurrent), tgt = parseFloat(acidTarget), vol = parseFloat(acidVolume);
    if (isNaN(cur) || isNaN(tgt) || isNaN(vol) || tgt <= cur || vol <= 0) return "";
    const liters = vol * 3.78541;
    return ((tgt - cur) * liters).toFixed(1);
  })();

  const [waterCurrentBrix, setWaterCurrentBrix] = useState("");
  const [waterTargetBrix, setWaterTargetBrix] = useState("");
  const [waterVolume, setWaterVolume] = useState("");
  const waterToAdd = (() => {
    const cur = parseFloat(waterCurrentBrix), tgt = parseFloat(waterTargetBrix), vol = parseFloat(waterVolume);
    if (isNaN(cur) || isNaN(tgt) || isNaN(vol) || tgt <= 0 || tgt >= cur || vol <= 0) return "";
    const newVolume = vol * (cur / tgt);
    return (newVolume - vol).toFixed(1);
  })();

  const [fahrenheit, setFahrenheit] = useState("");
  const celsius = fahrenheit !== "" && !isNaN(parseFloat(fahrenheit)) ? (((parseFloat(fahrenheit) - 32) * 5) / 9).toFixed(1) : "";

  const [celsiusIn, setCelsiusIn] = useState("");
  const fahrenheitOut = celsiusIn !== "" && !isNaN(parseFloat(celsiusIn)) ? ((parseFloat(celsiusIn) * 9) / 5 + 32).toFixed(1) : "";

  const [gallons, setGallons] = useState("");
  const liters = gallons !== "" && !isNaN(parseFloat(gallons)) ? (parseFloat(gallons) * 3.78541).toFixed(2) : "";

  const [tons, setTons] = useState("");
  const lbsFromTons = tons !== "" && !isNaN(parseFloat(tons)) ? (parseFloat(tons) * 2000).toFixed(0) : "";
  const kgFromTons = tons !== "" && !isNaN(parseFloat(tons)) ? (parseFloat(tons) * 907.185).toFixed(1) : "";

  return (
    <div className="space-y-6">
      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-1">Quick Calculators</h2>
        <p className="font-body text-xs text-stone-500 mb-4">
          Enter values below for an instant estimate — these are working estimates, not lab-precise measurements.
        </p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <CalcCard title="Potential Alcohol from Brix">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 items-end">
              <CalcInput label="Brix" value={brix} onChange={setBrix} unit="°Bx" />
              <CalcResult label="Potential ABV" value={potentialAlcohol} unit="%" />
            </div>
          </CalcCard>

          <CalcCard title="SO₂ Addition (as KMS powder)">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
              <CalcInput label="Current Free SO₂" value={so2Current} onChange={setSo2Current} unit="ppm" />
              <CalcInput label="Target Free SO₂" value={so2Target} onChange={setSo2Target} unit="ppm" />
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 items-end">
              <CalcInput label="Volume" value={so2Volume} onChange={setSo2Volume} unit="gal" />
              <CalcResult label="KMS Needed" value={kmsGrams} unit="g" />
            </div>
          </CalcCard>

          <CalcCard title="SO₂ Addition — 10% Liquid Solution (Tom's Method)">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
              <CalcInput label="Current Free SO₂" value={tomSo2Current} onChange={setTomSo2Current} unit="mg/L" />
              <CalcInput label="Target Free SO₂" value={tomSo2Target} onChange={setTomSo2Target} unit="mg/L" />
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 items-end">
              <CalcInput label="Volume" value={tomSo2Volume} onChange={setTomSo2Volume} unit="gal" />
              <CalcResult label="10% Solution Needed" value={tomSo2ML} unit="mL" />
            </div>
            <p className="font-body text-xs text-stone-400 mt-2">
              Uses a premixed 10% KMBS stock solution — different method from the powder calculator above, dosed in mL instead of grams.
            </p>
          </CalcCard>

          <CalcCard title="Chaptalization (Sugar Addition)">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
              <CalcInput label="Desired Brix Increase" value={chaptBrix} onChange={setChaptBrix} unit="°Bx" />
              <CalcInput label="Volume" value={chaptVolume} onChange={setChaptVolume} unit="gal" />
            </div>
            <CalcResult label="Sugar Needed" value={sugarLbs} unit="lbs" />
            <p className="font-body text-xs text-stone-400 mt-2">
              Chaptalization is restricted or prohibited in some regions/appellations — confirm local regulations before adding sugar.
            </p>
          </CalcCard>

          <CalcCard title="Acid Addition (Tartaric)">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
              <CalcInput label="Current TA" value={acidCurrent} onChange={setAcidCurrent} unit="g/L" />
              <CalcInput label="Target TA" value={acidTarget} onChange={setAcidTarget} unit="g/L" />
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 items-end">
              <CalcInput label="Volume" value={acidVolume} onChange={setAcidVolume} unit="gal" />
              <CalcResult label="Tartaric Acid Needed" value={tartaricGrams} unit="g" />
            </div>
            <p className="font-body text-xs text-stone-400 mt-2">
              Assumes the standard 1:1 rule of thumb — 1 g/L of tartaric acid added raises measured TA by about 1 g/L. Add gradually and re-check TA/pH, since real must doesn't always respond perfectly linearly.
            </p>
          </CalcCard>

          <CalcCard title="Water Addition (Dilution)">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
              <CalcInput label="Current Brix" value={waterCurrentBrix} onChange={setWaterCurrentBrix} unit="°Bx" />
              <CalcInput label="Target Brix" value={waterTargetBrix} onChange={setWaterTargetBrix} unit="°Bx" />
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 items-end">
              <CalcInput label="Current Volume" value={waterVolume} onChange={setWaterVolume} unit="gal" />
              <CalcResult label="Water to Add" value={waterToAdd} unit="gal" />
            </div>
            <p className="font-body text-xs text-stone-400 mt-2">
              Simple dilution math — adding water lowers Brix (and proportionally, TA and other dissolved solids) by the same ratio across the whole batch.
            </p>
          </CalcCard>

          <CalcCard title="Temperature (°F ↔ °C)">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3 items-end">
              <CalcInput label="°F" value={fahrenheit} onChange={setFahrenheit} />
              <CalcResult label="Result" value={celsius} unit="°C" />
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 items-end">
              <CalcInput label="°C" value={celsiusIn} onChange={setCelsiusIn} />
              <CalcResult label="Result" value={fahrenheitOut} unit="°F" />
            </div>
          </CalcCard>

          <CalcCard title="Volume (Gallons → Liters)">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 items-end">
              <CalcInput label="Gallons" value={gallons} onChange={setGallons} unit="gal" />
              <CalcResult label="Liters" value={liters} unit="L" />
            </div>
          </CalcCard>

          <CalcCard title="Weight (Tons → Lbs / Kg)">
            <CalcInput label="Tons" value={tons} onChange={setTons} unit="tons" />
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-3">
              <CalcResult label="Pounds" value={lbsFromTons} unit="lbs" />
              <CalcResult label="Kilograms" value={kgFromTons} unit="kg" />
            </div>
          </CalcCard>
        </div>
      </div>

      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-3">Reference Formulas</h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {REFERENCE_FORMULAS.map((f) => (
            <div key={f.title} className="border border-stone-200 rounded-lg p-4">
              <h3 className="font-body text-sm font-semibold text-stone-700 mb-1">{f.title}</h3>
              <p className="font-body text-sm text-ink-800 mb-1">{f.formula}</p>
              <p className="font-body text-xs text-stone-400">{f.note}</p>
            </div>
          ))}
        </div>
      </div>

      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-3">Tom's Formulas</h2>
        {TOMS_FORMULAS.length === 0 ? (
          <p className="font-body text-sm text-stone-500 py-4 text-center">No formulas added yet — check back soon.</p>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            {TOMS_FORMULAS.map((f) => (
              <div key={f.title} className="border border-stone-200 rounded-lg p-4">
                <h3 className="font-body text-sm font-semibold text-stone-700 mb-1">{f.title}</h3>
                <p className="font-body text-sm text-ink-800 mb-1">{f.formula}</p>
                {f.note && <p className="font-body text-xs text-stone-400">{f.note}</p>}
              </div>
            ))}
          </div>
        )}
      </div>


      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-3">Molecular SO₂ Reference — Free SO₂ needed for 0.8 ppm molecular SO₂</h2>
        <div className="overflow-x-auto">
          <table className="w-full text-sm font-body">
            <thead>
              <tr className="bg-stone-100 text-stone-600 text-left">
                <th className="px-3 py-2 font-medium">pH</th>
                <th className="px-3 py-2 font-medium">Free SO₂ Needed</th>
              </tr>
            </thead>
            <tbody>
              {MOLECULAR_SO2_TABLE.map((row) => (
                <tr key={row.ph} className="border-t border-stone-100">
                  <td className="px-3 py-2">{row.ph}</td>
                  <td className="px-3 py-2">{row.freeSO2}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p className="font-body text-xs text-stone-400 mt-2">Approximate values — use a lab or calculator for precise molecular SO₂ management.</p>
      </div>

      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-3">Brix ↔ Specific Gravity Reference</h2>
        <div className="overflow-x-auto">
          <table className="w-full text-sm font-body">
            <thead>
              <tr className="bg-stone-100 text-stone-600 text-left">
                <th className="px-3 py-2 font-medium">Brix</th>
                <th className="px-3 py-2 font-medium">Specific Gravity</th>
              </tr>
            </thead>
            <tbody>
              {BRIX_SG_TABLE.map((row) => (
                <tr key={row.brix} className="border-t border-stone-100">
                  <td className="px-3 py-2">{row.brix}°</td>
                  <td className="px-3 py-2">{row.sg}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-3">Unit Conversions</h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {UNIT_CONVERSIONS.map((group) => (
            <div key={group.category} className="border border-stone-200 rounded-lg p-4">
              <h3 className="font-body text-sm font-semibold text-stone-700 mb-2">{group.category}</h3>
              <ul className="font-body text-sm text-stone-600 space-y-1">
                {group.rows.map((row) => (
                  <li key={row}>{row}</li>
                ))}
              </ul>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ---------- Small presentational helpers for the print-only sheet ----------
function PrintHeader({ subtitle }) {
  return (
    <div style={{ borderBottom: "2px solid #1D1915", paddingBottom: 12, marginBottom: 16 }}>
      <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 700, color: "#1D1915", margin: 0 }}>
        Alloro Winery Tracker
      </h1>
      <p style={{ fontFamily: "Arial, sans-serif", fontSize: 12, color: "#555", margin: "4px 0 0" }}>{subtitle}</p>
    </div>
  );
}

function PrintSectionTitle({ children }) {
  return (
    <h2 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 20, color: "#065f46", margin: "20px 0 8px" }}>
      {children}
    </h2>
  );
}

function PrintTable({ headers, rows }) {
  return (
    <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: "Arial, sans-serif", fontSize: 12 }}>
      <thead>
        <tr>
          {headers.map((h) => (
            <th key={h} style={{ border: "1px solid #ccc", padding: 6, textAlign: "left" }}>{h}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.map((row, i) => (
          <tr key={i}>
            {row.map((cell, j) => (
              <td key={j} style={{ border: "1px solid #ccc", padding: 6 }}>{cell || ""}</td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function PrintEmpty({ children }) {
  return <p style={{ fontFamily: "Arial, sans-serif", fontSize: 12, color: "#888" }}>{children}</p>;
}

// ---------- Homepage: weather, growing degree days, today's work, and the calendar ----------
// ---------- One barrel row: shows status, current contents, and fill/empty/sell actions ----------
function BarrelRow({ barrel, fermentLots, onFill, onEmpty, onUpdate, onDelete, confirmAction, selected, onToggleSelect }) {
  const [expanded, setExpanded] = useState(false);
  const [showFillForm, setShowFillForm] = useState(false);
  const [fillRows, setFillRows] = useState([{ id: genId(), lotLabel: "", percentage: "100" }]);
  const [fillDate, setFillDate] = useState(todayISO());
  const [showSellForm, setShowSellForm] = useState(false);
  const [sellForm, setSellForm] = useState({ soldDate: todayISO(), soldTo: "", soldContact: "", soldPrice: "", saleNotes: "" });
  const [showRetireForm, setShowRetireForm] = useState(false);
  const [retireForm, setRetireForm] = useState({ retiredDate: todayISO(), retiredReason: "" });
  const [editing, setEditing] = useState(false);
  const [editForm, setEditForm] = useState(null);
  const [formError, setFormError] = useState("");
  const [showLabForm, setShowLabForm] = useState(false);
  const [labForm, setLabForm] = useState({ date: todayISO(), freeSO2: "", totalSO2: "", va: "", notes: "" });
  const [editingLabId, setEditingLabId] = useState(null);
  const [labEditForm, setLabEditForm] = useState(null);

  const submitLabCheck = () => {
    if (!labForm.date) return;
    const entry = { id: genId(), ...labForm };
    onUpdate(barrel.id, { labChecks: [entry, ...(barrel.labChecks || [])] });
    setLabForm({ date: todayISO(), freeSO2: "", totalSO2: "", va: "", notes: "" });
    setShowLabForm(false);
  };
  const startEditLab = (check) => {
    setEditingLabId(check.id);
    setLabEditForm({ ...check });
  };
  const saveEditLab = () => {
    const updated = (barrel.labChecks || []).map((c) => (c.id === editingLabId ? labEditForm : c));
    onUpdate(barrel.id, { labChecks: updated });
    setEditingLabId(null);
    setLabEditForm(null);
  };
  const deleteLabCheck = (id) => {
    confirmAction("Delete this lab check? This can't be undone.", () => {
      onUpdate(barrel.id, { labChecks: (barrel.labChecks || []).filter((c) => c.id !== id) });
    });
  };

  const status = barrelStatus(barrel);
  const fill = activeBarrelFill(barrel);
  const fillSummary = summarizeFillComponents(fill, fermentLots);
  const workHistory = [...(barrel.workHistory || [])].sort((a, b) => (a.dateCompleted < b.dateCompleted ? 1 : -1));
  const lastWork = workHistory[0] || null;
  const parsed = parseBarrelNumber(barrel.barrelNumber);
  const statusStyle =
    status === "In Use" ? "bg-green-100 text-green-800"
    : status === "Sold" ? "bg-stone-200 text-stone-600"
    : status === "Retired" ? "bg-orange-100 text-orange-800"
    : "bg-sky-100 text-sky-800";

  const addFillRow = () =>
    setFillRows((rows) => [
      ...rows.map((r) => ({ ...r, percentage: rows.length === 1 && r.percentage === "100" ? "" : r.percentage })),
      { id: genId(), lotLabel: "", percentage: "" },
    ]);
  const removeFillRow = (id) =>
    setFillRows((rows) => {
      if (rows.length <= 1) return rows;
      const next = rows.filter((r) => r.id !== id);
      return next.length === 1 ? [{ ...next[0], percentage: next[0].percentage || "100" }] : next;
    });
  const updateFillRow = (id, changes) => setFillRows((rows) => rows.map((r) => (r.id === id ? { ...r, ...changes } : r)));
  const fillPercentTotal = fillRows.reduce((sum, r) => sum + (parseFloat(r.percentage) || 0), 0);

  const submitFill = () => {
    const validRows = fillRows.filter((r) => r.lotLabel.trim());
    if (validRows.length === 0) {
      setFormError("Enter at least one lot");
      return;
    }
    setFormError("");
    const components = validRows.map((r) => {
      const matchedLot = fermentLots.find((l) => (l.tankId || "").toLowerCase() === r.lotLabel.trim().toLowerCase());
      return {
        id: genId(),
        lotLabel: r.lotLabel.trim(),
        lotId: matchedLot ? matchedLot.id : "",
        percentage: r.percentage === "" ? "" : parseFloat(r.percentage),
      };
    });
    onFill(barrel.id, { id: genId(), fillDate, emptyDate: "", components });
    setShowFillForm(false);
    setFillRows([{ id: genId(), lotLabel: "", percentage: "100" }]);
    setFillDate(todayISO());
  };

  const submitSell = () => {
    if (!sellForm.soldDate) {
      setFormError("Sold date is required");
      return;
    }
    setFormError("");
    onUpdate(barrel.id, sellForm);
    setShowSellForm(false);
  };

  const submitRetire = () => {
    if (!retireForm.retiredDate) {
      setFormError("Retired date is required");
      return;
    }
    setFormError("");
    onUpdate(barrel.id, retireForm);
    setShowRetireForm(false);
  };

  const startEdit = () => {
    setEditForm({
      barrelNumber: barrel.barrelNumber,
      cooperage: barrel.cooperage || "",
      wineColor: barrel.wineColor || "",
      forest: barrel.forest || "",
      toast: barrel.toast,
      notes: barrel.notes,
    });
    setEditing(true);
  };
  const saveEdit = () => {
    if (!parseBarrelNumber(editForm.barrelNumber)) {
      setFormError('Barrel number must look like "25-C01"');
      return;
    }
    setFormError("");
    onUpdate(barrel.id, editForm);
    setEditing(false);
  };

  const history = [...barrel.fills].sort((a, b) => (a.fillDate < b.fillDate ? 1 : -1));

  return (
    <div className="border-b border-stone-100 last:border-b-0">
      <div className="flex items-center gap-3 px-4 py-3 hover:bg-stone-50">
        <input
          type="checkbox"
          checked={!!selected}
          onChange={onToggleSelect}
          className="rounded border-stone-300 shrink-0"
        />
        <button onClick={() => setExpanded((v) => !v)} className="text-stone-400 shrink-0">
          {expanded ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
        </button>
        <div className="flex-1 min-w-0">
          <div className="flex flex-wrap items-center gap-2">
            <span className="font-body text-sm font-semibold text-stone-800">{barrel.barrelNumber}</span>
            <span className={`font-body text-xs px-1.5 py-0.5 rounded ${statusStyle}`}>{status}</span>
            {(barrel.cooperage || parsed) && (
              <span className="font-body text-xs text-stone-400">
                {barrel.cooperage || parsed?.cooperage}{parsed ? ` · 20${parsed.year}` : ""}
              </span>
            )}
            {status !== "Sold" && status !== "Retired" && <span className="font-body text-xs text-stone-400">· {fillCountLabel(barrel)}</span>}
            {barrel.wineColor && (
              <span className={`font-body text-xs px-1.5 py-0.5 rounded ${barrel.wineColor === "Red" ? "bg-rose-100 text-rose-800" : "bg-amber-50 text-amber-700"}`}>
                {barrel.wineColor}
              </span>
            )}
            {barrel.forest && <span className="font-body text-xs text-stone-400">· {barrel.forest}</span>}
            {barrel.toast && <span className="font-body text-xs text-stone-400">· {barrel.toast} Toast</span>}
          </div>
          <p className="font-body text-xs text-stone-500 mt-0.5">
            {fillSummary
              ? fillSummary
              : status === "Sold"
              ? `Sold ${barrel.soldDate}${barrel.soldTo ? " to " + barrel.soldTo : ""}`
              : status === "Retired"
              ? `Retired ${barrel.retiredDate}${barrel.retiredReason ? " — " + barrel.retiredReason : ""}`
              : "Empty"}
          </p>
          {lastWork && (
            <p className="font-body text-xs text-stone-400 mt-0.5">
              Last work: {lastWork.task || lastWork.taskType || "Work order"} — completed {lastWork.dateCompleted}
              {lastWork.workOrderNumber ? ` (${formatOrderNumber(lastWork.workOrderNumber)})` : ""}
            </p>
          )}
        </div>
        <div className="flex items-center gap-2 shrink-0">
          {status === "Empty" && (
            <button
              onClick={() => { setShowFillForm((v) => !v); setShowSellForm(false); setShowRetireForm(false); setFormError(""); }}
              className="font-body text-xs font-medium text-ink-800 hover:text-ink-900 border border-ink-200 rounded px-2 py-1"
            >
              Fill
            </button>
          )}
          {status === "In Use" && (
            <button
              onClick={() => confirmAction("Mark this barrel empty?", () => onEmpty(barrel.id, fill.id), { confirmLabel: "Empty", tone: "neutral" })}
              className="font-body text-xs font-medium text-stone-600 hover:text-stone-800 border border-stone-300 rounded px-2 py-1"
            >
              Empty
            </button>
          )}
          {status !== "Sold" && status !== "Retired" && (
            <button
              onClick={() => { setShowSellForm((v) => !v); setShowFillForm(false); setShowRetireForm(false); setFormError(""); }}
              className="font-body text-xs font-medium text-stone-600 hover:text-stone-800 border border-stone-300 rounded px-2 py-1"
            >
              Sell
            </button>
          )}
          {status !== "Sold" && status !== "Retired" && (
            <button
              onClick={() => { setShowRetireForm((v) => !v); setShowFillForm(false); setShowSellForm(false); setFormError(""); }}
              className="font-body text-xs font-medium text-orange-700 hover:text-orange-900 border border-orange-200 rounded px-2 py-1"
            >
              Retire
            </button>
          )}
        </div>
      </div>

      {showFillForm && (
        <div className="px-4 pb-3 bg-ink-50">
          <div className="pt-2 space-y-2">
            <datalist id={`lots-${barrel.id}`}>
              {fermentLots.map((l) => (
                <option key={l.id} value={l.tankId || ""} />
              ))}
            </datalist>
            {fillRows.map((row, i) => (
              <div key={row.id} className="grid grid-cols-1 sm:grid-cols-[1fr_120px_auto] gap-2 items-end">
                <div>
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">
                    {i === 0 ? "Lot" : "Additional Lot"}
                  </label>
                  <input
                    type="text"
                    list={`lots-${barrel.id}`}
                    placeholder="e.g. 114 T5 — type a new lot or pick an existing one"
                    value={row.lotLabel}
                    onChange={(e) => updateFillRow(row.id, { lotLabel: e.target.value })}
                    className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                  />
                </div>
                <div>
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">%</label>
                  <input
                    type="number"
                    value={row.percentage}
                    onChange={(e) => updateFillRow(row.id, { percentage: e.target.value })}
                    className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                  />
                </div>
                {fillRows.length > 1 && (
                  <button onClick={() => removeFillRow(row.id)} className="text-stone-400 hover:text-red-700 pb-2" title="Remove lot">
                    <Trash2 size={16} />
                  </button>
                )}
              </div>
            ))}
            <button
              onClick={addFillRow}
              className="font-body flex items-center gap-1 text-xs font-medium text-ink-800 hover:text-ink-900"
            >
              <Plus size={13} /> Add another lot (blend)
            </button>
            {fillRows.length > 1 && (
              <p className={`font-body text-xs ${fillPercentTotal === 100 ? "text-stone-400" : "text-amber-700"}`}>
                Total: {fillPercentTotal}% {fillPercentTotal !== 100 && "(doesn't have to add up to 100%, just flagging it)"}
              </p>
            )}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 items-end">
              <div>
                <label className="font-body block text-xs font-medium text-stone-600 mb-1">Fill Date</label>
                <input
                  type="date"
                  value={fillDate}
                  onChange={(e) => setFillDate(e.target.value)}
                  className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                />
              </div>
              <div className="flex items-center gap-2 sm:col-span-2">
                <button onClick={submitFill} className="font-body text-sm font-medium bg-ink-900 hover:bg-ink-800 text-white px-3 py-2 rounded-md">
                  Fill Barrel
                </button>
                <button onClick={() => setShowFillForm(false)} className="font-body text-sm text-stone-500 hover:text-stone-700">
                  Cancel
                </button>
              </div>
            </div>
          </div>
          {formError && <p className="font-body text-xs text-red-700 mt-2">{formError}</p>}
        </div>
      )}

      {showSellForm && (
        <div className="px-4 pb-3 bg-stone-50">
          <div className="grid grid-cols-1 sm:grid-cols-4 gap-2 pt-2">
            <div>
              <label className="font-body block text-xs font-medium text-stone-600 mb-1">Sold Date</label>
              <input
                type="date"
                value={sellForm.soldDate}
                onChange={(e) => setSellForm((p) => ({ ...p, soldDate: e.target.value }))}
                className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
              />
            </div>
            <div>
              <label className="font-body block text-xs font-medium text-stone-600 mb-1">Buyer</label>
              <input
                type="text"
                value={sellForm.soldTo}
                onChange={(e) => setSellForm((p) => ({ ...p, soldTo: e.target.value }))}
                className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
              />
            </div>
            <div>
              <label className="font-body block text-xs font-medium text-stone-600 mb-1">Buyer Contact</label>
              <input
                type="text"
                placeholder="phone or email"
                value={sellForm.soldContact}
                onChange={(e) => setSellForm((p) => ({ ...p, soldContact: e.target.value }))}
                className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
              />
            </div>
            <div>
              <label className="font-body block text-xs font-medium text-stone-600 mb-1">Sale Price ($)</label>
              <input
                type="number"
                value={sellForm.soldPrice}
                onChange={(e) => setSellForm((p) => ({ ...p, soldPrice: e.target.value }))}
                className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
              />
            </div>
          </div>
          <div className="mt-2">
            <label className="font-body block text-xs font-medium text-stone-600 mb-1">Notes</label>
            <textarea
              value={sellForm.saleNotes}
              onChange={(e) => setSellForm((p) => ({ ...p, saleNotes: e.target.value }))}
              rows={2}
              className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
            />
          </div>
          <div className="flex items-center gap-2 mt-2">
            <button onClick={submitSell} className="font-body text-sm font-medium bg-stone-800 hover:bg-stone-700 text-white px-3 py-2 rounded-md">
              Mark Sold
            </button>
            <button onClick={() => setShowSellForm(false)} className="font-body text-sm text-stone-500 hover:text-stone-700">
              Cancel
            </button>
          </div>
          {formError && <p className="font-body text-xs text-red-700 mt-2">{formError}</p>}
        </div>
      )}

      {showRetireForm && (
        <div className="px-4 pb-3 bg-orange-50">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 pt-2">
            <div>
              <label className="font-body block text-xs font-medium text-stone-600 mb-1">Retired Date</label>
              <input
                type="date"
                value={retireForm.retiredDate}
                onChange={(e) => setRetireForm((p) => ({ ...p, retiredDate: e.target.value }))}
                className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
              />
            </div>
            <div>
              <label className="font-body block text-xs font-medium text-stone-600 mb-1">Reason</label>
              <input
                type="text"
                placeholder="e.g. cracked, too old, repurposed"
                value={retireForm.retiredReason}
                onChange={(e) => setRetireForm((p) => ({ ...p, retiredReason: e.target.value }))}
                className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
              />
            </div>
          </div>
          <div className="flex items-center gap-2 mt-2">
            <button onClick={submitRetire} className="font-body text-sm font-medium bg-orange-700 hover:bg-orange-800 text-white px-3 py-2 rounded-md">
              Mark Retired
            </button>
            <button onClick={() => setShowRetireForm(false)} className="font-body text-sm text-stone-500 hover:text-stone-700">
              Cancel
            </button>
          </div>
          {formError && <p className="font-body text-xs text-red-700 mt-2">{formError}</p>}
        </div>
      )}

      {expanded && (
        <div className="px-4 pb-4 bg-stone-50">
          {editing ? (
            <div className="border border-stone-200 rounded-md p-3 bg-white mb-3">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-2">
                <div>
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">Barrel Number</label>
                  <input
                    type="text"
                    value={editForm.barrelNumber}
                    onChange={(e) => setEditForm((p) => ({ ...p, barrelNumber: e.target.value }))}
                    className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm"
                  />
                </div>
                <div>
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">Cooperage</label>
                  <select
                    value={editForm.cooperage}
                    onChange={(e) => setEditForm((p) => ({ ...p, cooperage: e.target.value }))}
                    className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm"
                  >
                    <option value="">— (use barrel # letter code)</option>
                    {COOPERAGE_NAMES.map((c) => (
                      <option key={c} value={c}>{c}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">Toast</label>
                  <select
                    value={editForm.toast}
                    onChange={(e) => setEditForm((p) => ({ ...p, toast: e.target.value }))}
                    className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm"
                  >
                    <option value="">—</option>
                    {BARREL_TOAST_LEVELS.map((t) => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">Wine Color</label>
                  <select
                    value={editForm.wineColor}
                    onChange={(e) => setEditForm((p) => ({ ...p, wineColor: e.target.value }))}
                    className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm"
                  >
                    <option value="">—</option>
                    {BARREL_WINE_COLORS.map((c) => (
                      <option key={c} value={c}>{c}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">Forest Origin</label>
                  <select
                    value={editForm.forest}
                    onChange={(e) => setEditForm((p) => ({ ...p, forest: e.target.value }))}
                    className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm"
                  >
                    <option value="">—</option>
                    {BARREL_FOREST_ORIGINS.map((f) => (
                      <option key={f} value={f}>{f}</option>
                    ))}
                  </select>
                </div>
                <div className="sm:col-span-2">
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">Notes</label>
                  <textarea
                    value={editForm.notes}
                    onChange={(e) => setEditForm((p) => ({ ...p, notes: e.target.value }))}
                    rows={2}
                    className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm"
                  />
                </div>
              </div>
              {formError && <p className="font-body text-xs text-red-700 mb-2">{formError}</p>}
              <div className="flex items-center gap-3">
                <button onClick={saveEdit} className="font-body flex items-center gap-1 text-sm font-medium text-ink-800 hover:text-ink-900">
                  <Check size={15} /> Save
                </button>
                <button onClick={() => setEditing(false)} className="font-body flex items-center gap-1 text-sm text-stone-500 hover:text-stone-700">
                  <X size={15} /> Cancel
                </button>
              </div>
            </div>
          ) : (
            <div className="flex items-center gap-4 flex-wrap font-body text-xs text-stone-500 mb-3 pt-1">
              <span>Size: 225L</span>
              {barrel.toast && <span>Toast: {barrel.toast}</span>}
              {barrel.wineColor && <span>Wine: {barrel.wineColor}</span>}
              {barrel.forest && <span>Forest: {barrel.forest}</span>}
              <button onClick={startEdit} className="text-stone-400 hover:text-ink-800 flex items-center gap-1">
                <Pencil size={12} /> Edit
              </button>
              <button
                onClick={() => confirmAction("Delete this barrel and its fill history? This can't be undone.", () => onDelete(barrel.id))}
                className="text-stone-400 hover:text-red-700 flex items-center gap-1"
              >
                <Trash2 size={12} /> Delete
              </button>
            </div>
          )}

          {status === "Sold" && (
            <p className="font-body text-xs text-stone-600 mb-3">
              Sold {barrel.soldDate} to {barrel.soldTo || "—"}
              {barrel.soldContact ? ` (${barrel.soldContact})` : ""} for ${barrel.soldPrice || "—"}
              {barrel.saleNotes && <> · {barrel.saleNotes}</>}
            </p>
          )}

          {status === "Retired" && (
            <p className="font-body text-xs text-stone-600 mb-3">
              Retired {barrel.retiredDate}
              {barrel.retiredReason && <> · {barrel.retiredReason}</>}
            </p>
          )}

          <p className="font-body text-xs font-semibold text-stone-600 mb-2">Fill History</p>
          {history.length === 0 ? (
            <p className="font-body text-xs text-stone-400">No fills logged yet.</p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-xs font-body">
              <thead>
                <tr className="text-left text-stone-500">
                  <th className="py-1 pr-2">Lot(s)</th>
                  <th className="py-1 pr-2">Filled</th>
                  <th className="py-1 pr-2">Emptied</th>
                </tr>
              </thead>
              <tbody>
                {history.map((f) => (
                  <tr key={f.id} className="border-t border-stone-200">
                    <td className="py-1 pr-2">{summarizeFillComponents(f, fermentLots) || "—"}</td>
                    <td className="py-1 pr-2">{f.fillDate}</td>
                    <td className="py-1 pr-2">{f.emptyDate || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            </div>
          )}

          <p className="font-body text-xs font-semibold text-stone-600 mb-2 mt-4">Work History</p>
          {workHistory.length === 0 ? (
            <p className="font-body text-xs text-stone-400">No work orders completed on this barrel yet.</p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-xs font-body">
              <thead>
                <tr className="text-left text-stone-500">
                  <th className="py-1 pr-2">Work Order</th>
                  <th className="py-1 pr-2">Task</th>
                  <th className="py-1 pr-2">Completed</th>
                </tr>
              </thead>
              <tbody>
                {workHistory.map((h) => (
                  <tr key={h.id} className="border-t border-stone-200">
                    <td className="py-1 pr-2 font-mono">{h.workOrderNumber ? formatOrderNumber(h.workOrderNumber) : "—"}</td>
                    <td className="py-1 pr-2">
                      {h.task || "—"}
                      {h.taskType ? ` (${h.taskType}${h.taskType === "Additions" && h.additionType ? " — " + h.additionType : ""})` : ""}
                    </td>
                    <td className="py-1 pr-2">{h.dateCompleted}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            </div>
          )}

          <div className="flex items-center justify-between mt-4 mb-2">
            <p className="font-body text-xs font-semibold text-stone-600">Lab Checks</p>
            <button
              onClick={() => setShowLabForm((v) => !v)}
              className="font-body flex items-center gap-1 text-xs font-medium text-ink-800 hover:text-ink-900"
            >
              <Plus size={12} /> Log Check
            </button>
          </div>
          {showLabForm && (
            <div className="bg-stone-50 border border-stone-200 rounded-md p-3 mb-3">
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-2">
                <div>
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">Date</label>
                  <input
                    type="date"
                    value={labForm.date}
                    onChange={(e) => setLabForm((p) => ({ ...p, date: e.target.value }))}
                    className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm"
                  />
                </div>
                <div>
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">Free SO₂ (mg/L)</label>
                  <input
                    type="number"
                    value={labForm.freeSO2}
                    onChange={(e) => setLabForm((p) => ({ ...p, freeSO2: e.target.value }))}
                    className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm"
                  />
                </div>
                <div>
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">Total SO₂ (mg/L)</label>
                  <input
                    type="number"
                    value={labForm.totalSO2}
                    onChange={(e) => setLabForm((p) => ({ ...p, totalSO2: e.target.value }))}
                    className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm"
                  />
                </div>
                <div>
                  <label className="font-body block text-xs font-medium text-stone-600 mb-1">VA (g/L)</label>
                  <input
                    type="number"
                    value={labForm.va}
                    onChange={(e) => setLabForm((p) => ({ ...p, va: e.target.value }))}
                    className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm"
                  />
                </div>
              </div>
              <label className="font-body block text-xs font-medium text-stone-600 mb-1">Notes</label>
              <textarea
                value={labForm.notes}
                onChange={(e) => setLabForm((p) => ({ ...p, notes: e.target.value }))}
                rows={2}
                className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm mb-2"
              />
              <div className="flex items-center gap-3">
                <button onClick={submitLabCheck} className="font-body text-sm font-medium text-ink-800 hover:text-ink-900">
                  Save Check
                </button>
                <button onClick={() => setShowLabForm(false)} className="font-body text-sm text-stone-500 hover:text-stone-700">
                  Cancel
                </button>
              </div>
            </div>
          )}
          {(barrel.labChecks || []).length === 0 ? (
            <p className="font-body text-xs text-stone-400">No lab checks logged yet.</p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-xs font-body">
              <thead>
                <tr className="text-left text-stone-500">
                  <th className="py-1 pr-2">Date</th>
                  <th className="py-1 pr-2">Free SO₂</th>
                  <th className="py-1 pr-2">Total SO₂</th>
                  <th className="py-1 pr-2">VA</th>
                  <th className="py-1 pr-2">Notes</th>
                  <th className="py-1 pr-2 w-12"></th>
                </tr>
              </thead>
              <tbody>
                {[...(barrel.labChecks || [])].sort((a, b) => (a.date < b.date ? 1 : -1)).map((c) =>
                  editingLabId === c.id ? (
                    <tr key={c.id} className="border-t border-stone-200 bg-ink-50">
                      <td className="py-1 pr-2"><input type="date" value={labEditForm.date} onChange={(e) => setLabEditForm((p) => ({ ...p, date: e.target.value }))} className="w-28 border border-stone-300 rounded px-1 py-0.5" /></td>
                      <td className="py-1 pr-2"><input type="number" value={labEditForm.freeSO2} onChange={(e) => setLabEditForm((p) => ({ ...p, freeSO2: e.target.value }))} className="w-16 border border-stone-300 rounded px-1 py-0.5" /></td>
                      <td className="py-1 pr-2"><input type="number" value={labEditForm.totalSO2} onChange={(e) => setLabEditForm((p) => ({ ...p, totalSO2: e.target.value }))} className="w-16 border border-stone-300 rounded px-1 py-0.5" /></td>
                      <td className="py-1 pr-2"><input type="number" value={labEditForm.va} onChange={(e) => setLabEditForm((p) => ({ ...p, va: e.target.value }))} className="w-16 border border-stone-300 rounded px-1 py-0.5" /></td>
                      <td className="py-1 pr-2"><input type="text" value={labEditForm.notes} onChange={(e) => setLabEditForm((p) => ({ ...p, notes: e.target.value }))} className="w-full border border-stone-300 rounded px-1 py-0.5" /></td>
                      <td className="py-1 pr-2">
                        <div className="flex items-center gap-1.5">
                          <button onClick={saveEditLab} className="text-ink-700 hover:text-ink-900"><Check size={13} /></button>
                          <button onClick={() => setEditingLabId(null)} className="text-stone-400 hover:text-stone-600"><X size={13} /></button>
                        </div>
                      </td>
                    </tr>
                  ) : (
                    <tr key={c.id} className="border-t border-stone-200">
                      <td className="py-1 pr-2">{c.date}</td>
                      <td className="py-1 pr-2">{c.freeSO2 || "—"}</td>
                      <td className="py-1 pr-2">{c.totalSO2 || "—"}</td>
                      <td className="py-1 pr-2">{c.va || "—"}</td>
                      <td className="py-1 pr-2">{c.notes || "—"}</td>
                      <td className="py-1 pr-2">
                        <div className="flex items-center gap-1.5">
                          <button onClick={() => startEditLab(c)} className="text-stone-400 hover:text-ink-800"><Pencil size={12} /></button>
                          <button onClick={() => deleteLabCheck(c.id)} className="text-stone-400 hover:text-red-700"><Trash2 size={13} /></button>
                        </div>
                      </td>
                    </tr>
                  )
                )}
              </tbody>
            </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ---------- Barrels tab: bulk import, add, filter/search, and the barrel list ----------
function BarrelsTab({ data, onAddBarrel, onBulkAdd, onFillBarrel, onEmptyBarrel, onUpdateBarrel, onDeleteBarrel, onBulkEmpty, onBulkFill, onBulkSell, confirmAction }) {
  const [bulkText, setBulkText] = useState("");
  const [bulkErrors, setBulkErrors] = useState([]);
  const [showBulk, setShowBulk] = useState(data.barrels.length === 0);

  const [newBarrel, setNewBarrel] = useState({ barrelNumber: "", cooperage: "", wineColor: "", forest: "", toast: "", notes: "" });
  const [addError, setAddError] = useState("");

  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("All");
  const [cooperageFilter, setCooperageFilter] = useState("All");
  const [wineColorFilter, setWineColorFilter] = useState("All");
  const [lotFilter, setLotFilter] = useState("All");

  const [selectedIds, setSelectedIds] = useState(new Set());
  const [showBulkFill, setShowBulkFill] = useState(false);
  const [bulkFillRows, setBulkFillRows] = useState([{ id: genId(), lotLabel: "", percentage: "100" }]);
  const [bulkFillDate, setBulkFillDate] = useState(todayISO());
  const [showBulkSell, setShowBulkSell] = useState(false);
  const [bulkSellForm, setBulkSellForm] = useState({ soldDate: todayISO(), soldTo: "", soldContact: "", soldPrice: "", saleNotes: "" });

  const addBulkFillRow = () =>
    setBulkFillRows((rows) => [
      ...rows.map((r) => ({ ...r, percentage: rows.length === 1 && r.percentage === "100" ? "" : r.percentage })),
      { id: genId(), lotLabel: "", percentage: "" },
    ]);
  const removeBulkFillRow = (id) =>
    setBulkFillRows((rows) => {
      if (rows.length <= 1) return rows;
      const next = rows.filter((r) => r.id !== id);
      return next.length === 1 ? [{ ...next[0], percentage: next[0].percentage || "100" }] : next;
    });
  const updateBulkFillRow = (id, changes) => setBulkFillRows((rows) => rows.map((r) => (r.id === id ? { ...r, ...changes } : r)));

  const toggleSelect = (id) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const clearSelection = () => {
    setSelectedIds(new Set());
    setShowBulkFill(false);
    setShowBulkSell(false);
    setBulkFillRows([{ id: genId(), lotLabel: "", percentage: "100" }]);
  };

  const handleBulkImport = () => {
    const { barrels, errors } = parseBulkBarrels(bulkText);
    setBulkErrors(errors);
    if (barrels.length > 0) {
      onBulkAdd(barrels);
      setBulkText("");
    }
  };

  const handleAddBarrel = (e) => {
    e.preventDefault();
    if (!parseBarrelNumber(newBarrel.barrelNumber)) {
      setAddError('Barrel number must look like "25-C01"');
      return;
    }
    setAddError("");
    onAddBarrel({ id: genId(), ...newBarrel, size: "225", soldDate: "", soldTo: "", soldContact: "", soldPrice: "", saleNotes: "", retiredDate: "", retiredReason: "", fills: [] });
    setNewBarrel({ barrelNumber: "", cooperage: "", wineColor: "", forest: "", toast: "", notes: "" });
  };

  const filtered = data.barrels.filter((b) => {
    if (search && !b.barrelNumber.toLowerCase().includes(search.toLowerCase())) return false;
    if (statusFilter !== "All" && barrelStatus(b) !== statusFilter) return false;
    if (cooperageFilter !== "All") {
      const parsed = parseBarrelNumber(b.barrelNumber);
      const effectiveCooperage = b.cooperage || parsed?.cooperage;
      if (effectiveCooperage !== cooperageFilter) return false;
    }
    if (lotFilter !== "All" && !b.fills.some((f) => f.lotId === lotFilter || (f.components || []).some((c) => c.lotId === lotFilter))) return false;
    if (wineColorFilter !== "All" && b.wineColor !== wineColorFilter) return false;
    return true;
  });

  const lotsForFilter = [...(data.ferment || [])].sort((a, b) => (a.tankId || "").localeCompare(b.tankId || ""));

  const selectedBarrels = data.barrels.filter((b) => selectedIds.has(b.id));
  const emptyEligible = selectedBarrels.filter((b) => barrelStatus(b) === "In Use").length;
  const fillEligible = selectedBarrels.filter((b) => barrelStatus(b) === "Empty").length;
  const sellEligible = selectedBarrels.filter((b) => barrelStatus(b) !== "Sold").length;

  const allVisibleSelected = filtered.length > 0 && filtered.every((b) => selectedIds.has(b.id));
  const toggleSelectAll = () => {
    setSelectedIds(allVisibleSelected ? new Set() : new Set(filtered.map((b) => b.id)));
  };

  const handleBulkEmpty = () => {
    confirmAction(
      `Mark ${emptyEligible} barrel${emptyEligible === 1 ? "" : "s"} empty?`,
      () => {
        onBulkEmpty([...selectedIds]);
        clearSelection();
      },
      { confirmLabel: "Empty", tone: "neutral" }
    );
  };

  const submitBulkFill = () => {
    const validRows = bulkFillRows.filter((r) => r.lotLabel.trim());
    if (validRows.length === 0) return;
    const components = validRows.map((r) => {
      const matchedLot = (data.ferment || []).find((l) => (l.tankId || "").toLowerCase() === r.lotLabel.trim().toLowerCase());
      return { lotLabel: r.lotLabel.trim(), lotId: matchedLot ? matchedLot.id : "", percentage: r.percentage === "" ? "" : parseFloat(r.percentage) };
    });
    onBulkFill([...selectedIds], components, bulkFillDate);
    clearSelection();
  };

  const submitBulkSell = () => {
    if (!bulkSellForm.soldDate) return;
    onBulkSell([...selectedIds], bulkSellForm);
    clearSelection();
    setBulkSellForm({ soldDate: todayISO(), soldTo: "", soldContact: "", soldPrice: "", saleNotes: "" });
  };

  const counts = {
    total: data.barrels.length,
    empty: data.barrels.filter((b) => barrelStatus(b) === "Empty").length,
    inUse: data.barrels.filter((b) => barrelStatus(b) === "In Use").length,
    sold: data.barrels.filter((b) => barrelStatus(b) === "Sold").length,
    retired: data.barrels.filter((b) => barrelStatus(b) === "Retired").length,
  };

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
        <div className="bg-white border border-stone-200 rounded-lg p-3 text-center">
          <p className="font-brand text-2xl text-ink-950">{counts.total}</p>
          <p className="font-body text-xs text-stone-500">Total Barrels</p>
        </div>
        <div className="bg-white border border-stone-200 rounded-lg p-3 text-center">
          <p className="font-brand text-2xl text-sky-700">{counts.empty}</p>
          <p className="font-body text-xs text-stone-500">Empty</p>
        </div>
        <div className="bg-white border border-stone-200 rounded-lg p-3 text-center">
          <p className="font-brand text-2xl text-green-700">{counts.inUse}</p>
          <p className="font-body text-xs text-stone-500">In Use</p>
        </div>
        <div className="bg-white border border-stone-200 rounded-lg p-3 text-center">
          <p className="font-brand text-2xl text-stone-500">{counts.sold}</p>
          <p className="font-body text-xs text-stone-500">Sold</p>
        </div>
        <div className="bg-white border border-stone-200 rounded-lg p-3 text-center">
          <p className="font-brand text-2xl text-orange-700">{counts.retired}</p>
          <p className="font-body text-xs text-stone-500">Retired</p>
        </div>
      </div>

      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <button onClick={() => setShowBulk((v) => !v)} className="w-full flex items-center justify-between">
          <h2 className="font-brand text-lg text-ink-950">Bulk Import Barrels</h2>
          {showBulk ? <ChevronUp size={18} className="text-stone-400" /> : <ChevronDown size={18} className="text-stone-400" />}
        </button>
        {showBulk && (
          <>
            <p className="font-body text-xs text-stone-500 mt-2 mb-2">
              One barrel per line: just the barrel number (e.g. <code>25-C01</code>), or add a toast level separated by a comma
              (e.g. <code>25-C01, Medium+</code>). All barrels are 225L.
            </p>
            <textarea
              value={bulkText}
              onChange={(e) => setBulkText(e.target.value)}
              rows={6}
              placeholder={"25-C01\n25-C02\n25-D01, Medium+"}
              className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-ink-800"
            />
            {bulkErrors.length > 0 && (
              <div className="font-body text-xs text-red-700 mt-2 space-y-0.5">
                {bulkErrors.map((e, i) => (
                  <p key={i}>{e}</p>
                ))}
              </div>
            )}
            <button
              onClick={handleBulkImport}
              className="font-body mt-3 flex items-center gap-2 bg-ink-900 hover:bg-ink-800 text-white text-sm font-medium px-4 py-2 rounded-md"
            >
              <Plus size={16} /> Import Barrels
            </button>
          </>
        )}
      </div>

      <form onSubmit={handleAddBarrel} className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-3">Add a Barrel</h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label className="font-body block text-xs font-medium text-stone-600 mb-1">Barrel Number</label>
            <input
              type="text"
              placeholder="25-C01"
              value={newBarrel.barrelNumber}
              onChange={(e) => setNewBarrel((p) => ({ ...p, barrelNumber: e.target.value }))}
              className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
            />
          </div>
          <div>
            <label className="font-body block text-xs font-medium text-stone-600 mb-1">Cooperage</label>
            <select
              value={newBarrel.cooperage}
              onChange={(e) => setNewBarrel((p) => ({ ...p, cooperage: e.target.value }))}
              className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
            >
              <option value="">— (use barrel # letter code)</option>
              {COOPERAGE_NAMES.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="font-body block text-xs font-medium text-stone-600 mb-1">Toast</label>
            <select
              value={newBarrel.toast}
              onChange={(e) => setNewBarrel((p) => ({ ...p, toast: e.target.value }))}
              className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
            >
              <option value="">—</option>
              {BARREL_TOAST_LEVELS.map((t) => (
                <option key={t} value={t}>{t}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="font-body block text-xs font-medium text-stone-600 mb-1">Wine Color</label>
            <select
              value={newBarrel.wineColor}
              onChange={(e) => setNewBarrel((p) => ({ ...p, wineColor: e.target.value }))}
              className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
            >
              <option value="">—</option>
              {BARREL_WINE_COLORS.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="font-body block text-xs font-medium text-stone-600 mb-1">Forest Origin</label>
            <select
              value={newBarrel.forest}
              onChange={(e) => setNewBarrel((p) => ({ ...p, forest: e.target.value }))}
              className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
            >
              <option value="">—</option>
              {BARREL_FOREST_ORIGINS.map((f) => (
                <option key={f} value={f}>{f}</option>
              ))}
            </select>
          </div>
          <div className="sm:col-span-2">
            <label className="font-body block text-xs font-medium text-stone-600 mb-1">Notes</label>
            <textarea
              value={newBarrel.notes}
              onChange={(e) => setNewBarrel((p) => ({ ...p, notes: e.target.value }))}
              rows={2}
              className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
            />
          </div>
        </div>
        {addError && <p className="font-body text-sm text-red-700 mt-3">{addError}</p>}
        <button
          type="submit"
          className="font-body mt-4 flex items-center gap-2 bg-ink-900 hover:bg-ink-800 text-white text-sm font-medium px-4 py-2 rounded-md"
        >
          <Plus size={16} /> Add Barrel
        </button>
      </form>

      <div className="bg-white border border-stone-200 rounded-lg overflow-hidden">
        <div className="px-4 py-3 border-b border-stone-200 flex items-center gap-2 flex-wrap">
          <input
            type="checkbox"
            checked={allVisibleSelected}
            onChange={toggleSelectAll}
            title="Select all visible barrels"
            className="rounded border-stone-300"
          />
          <input
            type="text"
            placeholder="Search barrel #"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="font-body text-sm border border-stone-300 rounded-md px-3 py-1.5 flex-1 min-w-[140px]"
          />
          <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className="font-body text-sm border border-stone-300 rounded-md px-2 py-1.5">
            <option value="All">All Statuses</option>
            <option value="Empty">Empty</option>
            <option value="In Use">In Use</option>
            <option value="Sold">Sold</option>
            <option value="Retired">Retired</option>
          </select>
          <select value={cooperageFilter} onChange={(e) => setCooperageFilter(e.target.value)} className="font-body text-sm border border-stone-300 rounded-md px-2 py-1.5">
            <option value="All">All Cooperages</option>
            {COOPERAGE_NAMES.map((c) => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
          <select value={lotFilter} onChange={(e) => setLotFilter(e.target.value)} className="font-body text-sm border border-stone-300 rounded-md px-2 py-1.5">
            <option value="All">All Lots</option>
            {lotsForFilter.map((lot) => (
              <option key={lot.id} value={lot.id}>{lot.tankId || "Untitled Lot"}</option>
            ))}
          </select>
          <select value={wineColorFilter} onChange={(e) => setWineColorFilter(e.target.value)} className="font-body text-sm border border-stone-300 rounded-md px-2 py-1.5">
            <option value="All">Red & White</option>
            {BARREL_WINE_COLORS.map((c) => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
          <span className="font-body text-xs text-stone-400 ml-auto">
            {filtered.length} of {data.barrels.length}
          </span>
        </div>

        {selectedIds.size > 0 && (
          <div className="px-4 py-3 border-b border-stone-200 bg-ink-50 flex items-center gap-2 flex-wrap">
            <span className="font-body text-sm font-medium text-ink-900">{selectedIds.size} selected</span>
            <button
              onClick={handleBulkEmpty}
              disabled={emptyEligible === 0}
              className="font-body text-xs font-medium text-stone-700 hover:text-stone-900 border border-stone-300 rounded px-2 py-1 disabled:opacity-40 disabled:cursor-not-allowed bg-white"
            >
              Empty ({emptyEligible})
            </button>
            <button
              onClick={() => { setShowBulkFill((v) => !v); setShowBulkSell(false); }}
              disabled={fillEligible === 0}
              className="font-body text-xs font-medium text-ink-800 hover:text-ink-900 border border-ink-200 rounded px-2 py-1 disabled:opacity-40 disabled:cursor-not-allowed bg-white"
            >
              Fill ({fillEligible})
            </button>
            <button
              onClick={() => { setShowBulkSell((v) => !v); setShowBulkFill(false); }}
              disabled={sellEligible === 0}
              className="font-body text-xs font-medium text-stone-700 hover:text-stone-900 border border-stone-300 rounded px-2 py-1 disabled:opacity-40 disabled:cursor-not-allowed bg-white"
            >
              Sell ({sellEligible})
            </button>
            <button onClick={clearSelection} className="font-body text-xs text-stone-500 hover:text-stone-700 underline ml-auto">
              Clear selection
            </button>
          </div>
        )}

        {showBulkFill && (
          <div className="px-4 py-3 border-b border-stone-200 bg-ink-50/60">
            <p className="font-body text-xs font-semibold text-stone-700 mb-2">
              Fill {fillEligible} empty barrel{fillEligible === 1 ? "" : "s"} with the same lot(s)
            </p>
            <datalist id="bulk-fill-lots">
              {(data.ferment || []).map((l) => (
                <option key={l.id} value={l.tankId || ""} />
              ))}
            </datalist>
            <div className="space-y-2 mb-2">
              {bulkFillRows.map((row, i) => (
                <div key={row.id} className="grid grid-cols-1 sm:grid-cols-[1fr_120px_auto] gap-2 items-end">
                  <div>
                    <label className="font-body block text-xs font-medium text-stone-600 mb-1">
                      {i === 0 ? "Lot" : "Additional Lot"}
                    </label>
                    <input
                      type="text"
                      list="bulk-fill-lots"
                      placeholder="e.g. 114 T5 — type a new lot or pick an existing one"
                      value={row.lotLabel}
                      onChange={(e) => updateBulkFillRow(row.id, { lotLabel: e.target.value })}
                      className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                    />
                  </div>
                  <div>
                    <label className="font-body block text-xs font-medium text-stone-600 mb-1">%</label>
                    <input
                      type="number"
                      value={row.percentage}
                      onChange={(e) => updateBulkFillRow(row.id, { percentage: e.target.value })}
                      className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                    />
                  </div>
                  {bulkFillRows.length > 1 && (
                    <button onClick={() => removeBulkFillRow(row.id)} className="text-stone-400 hover:text-red-700 pb-2" title="Remove lot">
                      <Trash2 size={16} />
                    </button>
                  )}
                </div>
              ))}
              <button onClick={addBulkFillRow} className="font-body flex items-center gap-1 text-xs font-medium text-ink-800 hover:text-ink-900">
                <Plus size={13} /> Add another lot (blend)
              </button>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 items-end">
              <div>
                <label className="font-body block text-xs font-medium text-stone-600 mb-1">Fill Date</label>
                <input
                  type="date"
                  value={bulkFillDate}
                  onChange={(e) => setBulkFillDate(e.target.value)}
                  className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                />
              </div>
              <div className="flex items-center gap-2 sm:col-span-2">
                <button onClick={submitBulkFill} className="font-body text-sm font-medium bg-ink-900 hover:bg-ink-800 text-white px-3 py-2 rounded-md">
                  Fill Barrels
                </button>
                <button onClick={() => setShowBulkFill(false)} className="font-body text-sm text-stone-500 hover:text-stone-700">
                  Cancel
                </button>
              </div>
            </div>
          </div>
        )}

        {showBulkSell && (
          <div className="px-4 py-3 border-b border-stone-200 bg-stone-50">
            <p className="font-body text-xs font-semibold text-stone-700 mb-2">
              Mark {sellEligible} barrel{sellEligible === 1 ? "" : "s"} sold with the same sale details
            </p>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
              <div>
                <label className="font-body block text-xs font-medium text-stone-600 mb-1">Sold Date</label>
                <input
                  type="date"
                  value={bulkSellForm.soldDate}
                  onChange={(e) => setBulkSellForm((p) => ({ ...p, soldDate: e.target.value }))}
                  className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                />
              </div>
              <div>
                <label className="font-body block text-xs font-medium text-stone-600 mb-1">Sold To</label>
                <input
                  type="text"
                  value={bulkSellForm.soldTo}
                  onChange={(e) => setBulkSellForm((p) => ({ ...p, soldTo: e.target.value }))}
                  className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                />
              </div>
              <div>
                <label className="font-body block text-xs font-medium text-stone-600 mb-1">Buyer Contact</label>
                <input
                  type="text"
                  value={bulkSellForm.soldContact}
                  onChange={(e) => setBulkSellForm((p) => ({ ...p, soldContact: e.target.value }))}
                  className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                />
              </div>
              <div>
                <label className="font-body block text-xs font-medium text-stone-600 mb-1">Sale Price (each)</label>
                <input
                  type="number"
                  value={bulkSellForm.soldPrice}
                  onChange={(e) => setBulkSellForm((p) => ({ ...p, soldPrice: e.target.value }))}
                  className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                />
              </div>
              <div className="sm:col-span-2">
                <label className="font-body block text-xs font-medium text-stone-600 mb-1">Sale Notes</label>
                <textarea
                  value={bulkSellForm.saleNotes}
                  onChange={(e) => setBulkSellForm((p) => ({ ...p, saleNotes: e.target.value }))}
                  rows={2}
                  className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                />
              </div>
            </div>
            <div className="flex items-center gap-2 mt-2">
              <button onClick={submitBulkSell} className="font-body text-sm font-medium bg-ink-900 hover:bg-ink-800 text-white px-3 py-2 rounded-md">
                Mark Sold
              </button>
              <button onClick={() => setShowBulkSell(false)} className="font-body text-sm text-stone-500 hover:text-stone-700">
                Cancel
              </button>
            </div>
          </div>
        )}

        {filtered.length === 0 ? (
          <p className="font-body text-sm text-stone-500 px-4 py-8 text-center">
            {data.barrels.length === 0 ? "No barrels yet — bulk import or add one above." : "No barrels match this filter."}
          </p>
        ) : (
          <div>
            {filtered.map((b) => (
              <BarrelRow
                key={b.id}
                barrel={b}
                fermentLots={data.ferment}
                onFill={onFillBarrel}
                onEmpty={onEmptyBarrel}
                onUpdate={onUpdateBarrel}
                onDelete={onDeleteBarrel}
                confirmAction={confirmAction}
                selected={selectedIds.has(b.id)}
                onToggleSelect={() => toggleSelect(b.id)}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function HomeTab({ data, toggleWorkOrder, deleteWorkOrder, editingWorkOrderId, editWorkOrderForm, editWorkOrderChange, startEditWorkOrder, saveEditWorkOrder, cancelEditWorkOrder, duplicateWorkOrder, saveAsTemplate, lotNames, onRegisterLotName, onLogWeather, sprayPrograms, onAddSprayProgram }) {
  const [weather, setWeather] = useState(null);
  const [gdd, setGdd] = useState(null);
  const [weatherError, setWeatherError] = useState("");
  const [gddCompareVariety, setGddCompareVariety] = useState("");
  const [weatherLoading, setWeatherLoading] = useState(true);
  const [retryCount, setRetryCount] = useState(0);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setWeatherLoading(true);
      setWeatherError("");
      try {
        const { lat, lon, tz } = HOME_COORDS;
        const forecastUrl =
          `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}` +
          `&current=temperature_2m,relative_humidity_2m,precipitation,weather_code,wind_speed_10m,shortwave_radiation` +
          `&hourly=soil_moisture_0_to_1cm,soil_moisture_9_to_27cm` +
          `&daily=weather_code,temperature_2m_max,temperature_2m_min,precipitation_sum,precipitation_probability_max,sunrise,sunset` +
          `&past_days=7&forecast_days=6&temperature_unit=fahrenheit&wind_speed_unit=mph&precipitation_unit=inch&timezone=${encodeURIComponent(tz)}`;
        const res = await fetch(forecastUrl);
        if (!res.ok) {
          let reason = `HTTP ${res.status}`;
          try {
            const errJson = await res.json();
            if (errJson?.reason) reason = errJson.reason;
          } catch {
            // response wasn't JSON — stick with the HTTP status
          }
          throw new Error(reason);
        }
        const json = await res.json();
        if (cancelled) return;
        setWeather(json);

        // Growing Degree Days (base 50°F), summed from April 1 through today.
        // The live forecast call above only reaches ~7 days into the past (Open-Meteo caps
        // past_days well short of a full growing season), so on its own it silently under-counts
        // by months once the season's underway. The dedicated Archive API accepts an arbitrary
        // start_date/end_date and covers the rest of the season — it just runs a few days behind
        // real-time, so we combine it with the recent-days data already fetched above to close
        // that gap and get a complete, accurate, genuinely-updates-every-day total.
        const year = new Date().getFullYear();
        const seasonStart = `${year}-04-01`;
        const todayStr = todayInTimezone(tz);
        const byDate = {};
        const recentTimes = json.daily?.time || [];
        const recentMaxes = json.daily?.temperature_2m_max || [];
        const recentMins = json.daily?.temperature_2m_min || [];
        recentTimes.forEach((d, i) => {
          if (d >= seasonStart && d <= todayStr && recentMaxes[i] != null && recentMins[i] != null) {
            byDate[d] = { max: recentMaxes[i], min: recentMins[i] };
          }
        });

        try {
          const archiveEndDate = recentTimes.length > 0 ? recentTimes[0] : todayStr; // day before the live call's earliest day
          if (seasonStart < archiveEndDate) {
            const archiveUrl =
              `https://archive-api.open-meteo.com/v1/archive?latitude=${lat}&longitude=${lon}` +
              `&start_date=${seasonStart}&end_date=${archiveEndDate}` +
              `&daily=temperature_2m_max,temperature_2m_min&temperature_unit=fahrenheit&timezone=${encodeURIComponent(tz)}`;
            const archiveRes = await fetch(archiveUrl);
            if (archiveRes.ok) {
              const archiveJson = await archiveRes.json();
              const aTimes = archiveJson.daily?.time || [];
              const aMaxes = archiveJson.daily?.temperature_2m_max || [];
              const aMins = archiveJson.daily?.temperature_2m_min || [];
              aTimes.forEach((d, i) => {
                if (!byDate[d] && aMaxes[i] != null && aMins[i] != null) {
                  byDate[d] = { max: aMaxes[i], min: aMins[i] };
                }
              });
            }
          }
        } catch {
          // Archive call failed — GDD will just be based on the recent-days data above instead
          // of silently showing nothing.
        }

        let total = 0;
        let daysCounted = 0;
        Object.values(byDate).forEach(({ max, min }) => {
          const avg = (max + min) / 2;
          total += Math.max(0, avg - 50);
          daysCounted += 1;
        });
        setGdd({ total: Math.round(total), daysCounted, seasonStart });
      } catch (err) {
        if (!cancelled) {
          const detail = err?.message || err?.name || "Unknown error";
          setWeatherError(`Couldn't load weather right now (${detail}).`);
        }
      } finally {
        if (!cancelled) setWeatherLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [retryCount]);

  const today = todayISO();
  const todaysOrders = data.workorders.filter((o) => o.date === today && o.status !== "Complete");
  const completedToday = data.workorders.filter((o) => o.dateCompleted === today);

  const current = weather?.current;
  const weatherToday = todayInTimezone(HOME_COORDS.tz);
  const dailyIdx = weather?.daily?.time?.indexOf(weatherToday) ?? -1;
  const todayHigh = dailyIdx >= 0 ? weather.daily.temperature_2m_max[dailyIdx] : null;
  const todayLow = dailyIdx >= 0 ? weather.daily.temperature_2m_min[dailyIdx] : null;
  const sunrise = dailyIdx >= 0 ? weather.daily.sunrise[dailyIdx]?.slice(11) : null;
  const sunset = dailyIdx >= 0 ? weather.daily.sunset[dailyIdx]?.slice(11) : null;
  const frostRisk = todayLow != null && todayLow <= 34;

  // Soil moisture is hourly data, not part of the "current" convenience block — find the hour
  // matching (or just before) right now to use as a stand-in for "current" soil moisture.
  const hourlyTime = weather?.hourly?.time || [];
  let nowHourIdx = -1;
  if (current?.time) {
    nowHourIdx = hourlyTime.indexOf(current.time.slice(0, 13) + ":00");
    if (nowHourIdx === -1) {
      for (let i = hourlyTime.length - 1; i >= 0; i--) {
        if (hourlyTime[i] <= current.time) {
          nowHourIdx = i;
          break;
        }
      }
    }
  }
  const soilMoistureSurface = nowHourIdx >= 0 ? weather.hourly.soil_moisture_0_to_1cm?.[nowHourIdx] : null;
  const soilMoistureRootZone = nowHourIdx >= 0 ? weather.hourly.soil_moisture_9_to_27cm?.[nowHourIdx] : null;

  // Next 5 days, starting from tomorrow
  const upcoming =
    dailyIdx >= 0 ? weather.daily.time.slice(dailyIdx + 1, dailyIdx + 6).map((d, i) => {
      const idx = dailyIdx + 1 + i;
      return {
        date: d,
        high: weather.daily.temperature_2m_max[idx],
        low: weather.daily.temperature_2m_min[idx],
        code: weather.daily.weather_code[idx],
        pop: weather.daily.precipitation_probability_max[idx],
      };
    }) : [];

  const todaysWeatherLog = (data.weatherLogs || []).find((w) => w.date === weatherToday) || null;
  const logToday = () => {
    onLogWeather({
      date: weatherToday,
      high: todayHigh,
      low: todayLow,
      currentTemp: current?.temperature_2m ?? null,
      conditionCode: current?.weather_code ?? null,
      conditionLabel: weatherInfo(current?.weather_code).label,
      humidity: current?.relative_humidity_2m ?? null,
      windMph: current?.wind_speed_10m ?? null,
      precipitationIn: current?.precipitation ?? null,
      gddTotal: gdd?.total ?? null,
    });
  };

  return (
    <div className="space-y-6">
      <div>
        <h2 className="font-brand text-2xl text-ink-950">
          {new Date().toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric" })}
        </h2>
        <p className="font-body text-sm text-stone-500">Estate overview · Weather from McMinnville Airport (KMMV)</p>
      </div>

      {/* Weather + GDD */}
      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h3 className="font-brand text-lg text-ink-950 mb-3">Weather</h3>
        {weatherLoading ? (
          <div className="flex items-center gap-2 text-stone-500 font-body text-sm py-6 justify-center">
            <Loader2 size={18} className="animate-spin" /> Loading weather…
          </div>
        ) : weatherError ? (
          <div className="py-4 text-center">
            <p className="font-body text-sm text-stone-500 mb-2">{weatherError}</p>
            <button
              onClick={() => setRetryCount((n) => n + 1)}
              className="font-body text-sm font-medium text-ink-800 hover:text-ink-900 underline"
            >
              Try again
            </button>
          </div>
        ) : (
          <>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-4">
              <div className="flex items-center gap-3">
                <span className="text-4xl">{weatherInfo(current?.weather_code).icon}</span>
                <div>
                  <p className="font-brand text-3xl text-ink-950 leading-none">
                    {current?.temperature_2m != null ? Math.round(current.temperature_2m) : "—"}°F
                  </p>
                  <p className="font-body text-xs text-stone-500 mt-1">{weatherInfo(current?.weather_code).label}</p>
                </div>
              </div>
              <div className="font-body text-sm text-stone-600 space-y-1">
                <p className="flex items-center gap-1.5"><Thermometer size={14} className="text-stone-400" /> High {todayHigh != null ? `${Math.round(todayHigh)}°F` : "—"} · Low {todayLow != null ? `${Math.round(todayLow)}°F` : "—"}</p>
                <p className="flex items-center gap-1.5"><Droplets size={14} className="text-stone-400" /> Humidity {current?.relative_humidity_2m != null ? `${current.relative_humidity_2m}%` : "—"}</p>
                <p className="flex items-center gap-1.5"><Wind size={14} className="text-stone-400" /> Wind {current?.wind_speed_10m != null ? `${Math.round(current.wind_speed_10m)} mph` : "—"}</p>
              </div>
              <div className="font-body text-sm text-stone-600 space-y-1">
                <p className="flex items-center gap-1.5"><Sunrise size={14} className="text-stone-400" /> Sunrise {sunrise || "—"}</p>
                <p className="flex items-center gap-1.5"><Sunset size={14} className="text-stone-400" /> Sunset {sunset || "—"}</p>
                <p className="flex items-center gap-1.5"><CloudRain size={14} className="text-stone-400" /> Rain today {weather?.current?.precipitation != null ? `${weather.current.precipitation}"` : "—"}</p>
                <p className="flex items-center gap-1.5"><Sun size={14} className="text-stone-400" /> Solar {current?.shortwave_radiation != null ? `${Math.round(current.shortwave_radiation)} W/m²` : "—"}</p>
              </div>
            </div>

            {frostRisk && (
              <p className="font-body text-xs bg-sky-50 text-sky-800 border border-sky-100 rounded-md px-3 py-2 mb-4">
                ❄️ Frost risk tonight — forecast low is {Math.round(todayLow)}°F.
              </p>
            )}

            {upcoming.length > 0 && (
              <div className="grid grid-cols-5 gap-2">
                {upcoming.map((d) => (
                  <div key={d.date} className="bg-stone-50 border border-stone-200 rounded-md p-2 text-center">
                    <p className="font-body text-xs text-stone-500">{new Date(d.date + "T00:00").toLocaleDateString(undefined, { weekday: "short" })}</p>
                    <p className="text-lg">{weatherInfo(d.code).icon}</p>
                    <p className="font-body text-xs text-stone-700">{d.high != null ? Math.round(d.high) : "—"}° / {d.low != null ? Math.round(d.low) : "—"}°</p>
                    {d.pop != null && <p className="font-body text-xs text-sky-600">{d.pop}%</p>}
                  </div>
                ))}
              </div>
            )}

            <div className="flex items-center justify-between mt-4 pt-3 border-t border-stone-100">
              {todaysWeatherLog ? (
                <p className="font-body text-xs text-ink-700">
                  ✓ Logged to Calendar for today ({todaysWeatherLog.high != null ? Math.round(todaysWeatherLog.high) : "—"}° / {todaysWeatherLog.low != null ? Math.round(todaysWeatherLog.low) : "—"}°)
                </p>
              ) : (
                <span className="font-body text-xs text-stone-400">Not logged to Calendar yet today</span>
              )}
              <button
                onClick={logToday}
                className="font-body flex items-center gap-1.5 text-xs font-medium text-ink-900 hover:text-ink-700 border border-ink-200 rounded-md px-3 py-1.5"
              >
                <CalendarPlus size={13} /> {todaysWeatherLog ? "Update Log" : "Log Today's Weather"}
              </button>
            </div>
          </>
        )}
      </div>

      {/* Growing Degree Days */}
      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h3 className="font-brand text-lg text-ink-950 mb-1">Growing Degree Days</h3>
        <p className="font-body text-xs text-stone-500 mb-3">Base 50°F, accumulated April 1 – today (Winkler Index method)</p>
        {weatherLoading ? (
          <p className="font-body text-sm text-stone-400">Calculating…</p>
        ) : gdd ? (
          <div className="flex items-baseline gap-3">
            <p className="font-brand text-4xl text-ink-950">{gdd.total.toLocaleString()}</p>
            <p className="font-body text-sm text-stone-500">GDD over {gdd.daysCounted} days</p>
          </div>
        ) : (
          <p className="font-body text-sm text-stone-400">Not available.</p>
        )}
        <p className="font-body text-xs text-stone-400 mt-2">
          Estimated from regional weather data, not an on-site sensor — treat as a close approximation.
        </p>

        <div className="mt-4 pt-4 border-t border-stone-100">
          <label className="font-body block text-xs font-medium text-stone-600 mb-1">Compare to a variety</label>
          <select
            value={gddCompareVariety}
            onChange={(e) => setGddCompareVariety(e.target.value)}
            className="font-body w-full sm:w-64 border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
          >
            <option value="">Select a variety…</option>
            {GRAPE_VARIETIES.map((v) => (
              <option key={v} value={v}>{v}</option>
            ))}
          </select>
          {gddCompareVariety &&
            (VARIETY_GDD_REFERENCE[gddCompareVariety] ? (
              (() => {
                const ref = VARIETY_GDD_REFERENCE[gddCompareVariety];
                return (
                  <div className="mt-2">
                    <p className="font-body text-sm text-stone-700">
                      Typical full-season range for {gddCompareVariety}:{" "}
                      <span className="font-semibold">{ref.min.toLocaleString()}–{ref.max.toLocaleString()} GDD</span>{" "}
                      <span className="text-stone-400">(Winkler Region {ref.region})</span>
                    </p>
                    {gdd && (
                      <p className="font-body text-xs text-stone-500 mt-1">
                        You're at {gdd.total.toLocaleString()} GDD so far this season —{" "}
                        {gdd.total < ref.min
                          ? `${(ref.min - gdd.total).toLocaleString()} GDD below the typical range, with more of the season still to go.`
                          : gdd.total > ref.max
                          ? `already ${(gdd.total - ref.max).toLocaleString()} GDD above the typical range.`
                          : "within the typical range for this point in the season."}
                      </p>
                    )}
                  </div>
                );
              })()
            ) : (
              <p className="font-body text-xs text-stone-400 mt-2">No standard reference range for "{gddCompareVariety}."</p>
            ))}
          <p className="font-body text-xs text-stone-400 mt-2">
            Based on the Winkler Index climate classification — general viticultural guidance, not a precise target. Actual optimal GDD varies by clone, site, and desired wine style.
          </p>
        </div>
      </div>

      {/* Soil Moisture */}
      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h3 className="font-brand text-lg text-ink-950 mb-1">Soil Moisture</h3>
        <p className="font-body text-xs text-stone-500 mb-3">
          Modeled from regional weather data — not a sensor in your actual vineyard. Useful for spotting broad trends, especially for dry-farming decisions, but treat as a rough regional estimate, not ground-truth for a specific block.
        </p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div>
            <p className="font-body text-xs text-stone-500">Surface (0–1 cm)</p>
            <p className="font-brand text-2xl text-ink-950">
              {soilMoistureSurface != null ? `${Math.round(soilMoistureSurface * 100)}%` : "—"}
            </p>
          </div>
          <div>
            <p className="font-body text-xs text-stone-500">Root Zone (9–27 cm)</p>
            <p className="font-brand text-2xl text-ink-950">
              {soilMoistureRootZone != null ? `${Math.round(soilMoistureRootZone * 100)}%` : "—"}
            </p>
          </div>
        </div>
        <p className="font-body text-xs text-stone-400 mt-3">
          Shown as approximate volumetric water content. If you ever install an actual soil moisture probe on-site, tell me — I can add that as a real data source instead of this regional estimate.
        </p>
      </div>

      {/* Active Fermentations */}
      <MultiLotFermentChart lots={data.ferment.filter((l) => l.status === "Active")} />

      {/* Tonnage by Variety */}
      <TonnageByVarietyChart harvest={data.harvest} />

      {/* Today's Work Orders */}
      <div className="bg-white border border-stone-200 rounded-lg overflow-hidden">
        <div className="px-4 py-3 border-b border-stone-200 flex items-center justify-between">
          <h3 className="font-brand text-lg text-ink-950">Today's Work Orders</h3>
          <span className="font-body text-xs text-stone-500">
            {todaysOrders.length} open{completedToday.length > 0 ? ` · ${completedToday.length} completed today` : ""}
          </span>
        </div>
        {todaysOrders.length === 0 ? (
          <p className="font-body text-sm text-stone-500 px-4 py-8 text-center">Nothing assigned for today.</p>
        ) : (
          <ul className="divide-y divide-stone-100">
            {todaysOrders.map((o) => (
              <WorkOrderRow
                key={o.id}
                order={o}
                onToggle={toggleWorkOrder}
                onDelete={deleteWorkOrder}
                onDuplicate={duplicateWorkOrder}
                onSaveAsTemplate={saveAsTemplate}
                isEditing={editingWorkOrderId === o.id}
                editForm={editWorkOrderForm}
                onEditChange={editWorkOrderChange}
                onStartEdit={startEditWorkOrder}
                onSaveEdit={saveEditWorkOrder}
                onCancelEdit={cancelEditWorkOrder}
                fermentLots={data.ferment}
                barrelsList={data.barrels}
                lotNamesList={lotNames}
                onRegisterLotName={onRegisterLotName}
                sprayProgramsList={sprayPrograms}
                onAddSprayProgram={onAddSprayProgram}
              />
            ))}
          </ul>
        )}
      </div>

      {/* Calendar */}
      <MasterCalendar data={data} />
    </div>
  );
}

// ---------- Single-barrel search+select, used when logging a tasting ----------
function SingleBarrelPicker({ value, onChange, barrelsList }) {
  const [search, setSearch] = useState("");
  const selectedBarrel = barrelsList.find((b) => b.id === value);
  const filtered = search ? barrelsList.filter((b) => b.barrelNumber.toLowerCase().includes(search.toLowerCase())) : barrelsList;

  if (selectedBarrel) {
    return (
      <div className="flex items-center gap-2">
        <span className="font-body text-sm bg-ink-50 text-ink-800 px-2.5 py-1.5 rounded-md">{selectedBarrel.barrelNumber}</span>
        <button type="button" onClick={() => onChange("")} className="font-body text-xs text-stone-400 hover:text-red-700">
          Change
        </button>
      </div>
    );
  }
  return (
    <div>
      <input
        type="text"
        placeholder="Search barrel #"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm mb-2 focus:outline-none focus:ring-2 focus:ring-ink-800"
      />
      <div className="flex flex-wrap gap-1.5 max-h-32 overflow-y-auto border border-stone-200 rounded-md p-2">
        {filtered.length === 0 ? (
          <span className="font-body text-xs text-stone-400">No barrels match.</span>
        ) : (
          filtered.slice(0, 60).map((b) => (
            <button
              key={b.id}
              type="button"
              onClick={() => onChange(b.id)}
              className="font-body text-xs px-2 py-1 rounded-full border bg-white text-stone-600 border-stone-300 hover:border-ink-400"
            >
              {b.barrelNumber}
            </button>
          ))
        )}
      </div>
    </div>
  );
}

// ---------- Shared field grid used by both the tasting add-form and inline edit ----------
function TastingFormFields({ form, setForm, barrelsList }) {
  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
      <div className="sm:col-span-2">
        <label className="font-body block text-xs font-medium text-stone-600 mb-1">Barrel</label>
        <SingleBarrelPicker value={form.barrelId} onChange={(id) => setForm((p) => ({ ...p, barrelId: id }))} barrelsList={barrelsList} />
      </div>
      <div>
        <label className="font-body block text-xs font-medium text-stone-600 mb-1">Date</label>
        <input
          type="date"
          value={form.date}
          onChange={(e) => setForm((p) => ({ ...p, date: e.target.value }))}
          className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
      </div>
      <div />
      {RATING_DIMENSIONS.map((d) => (
        <div key={d.key}>
          <label className="font-body block text-xs font-medium text-stone-600 mb-1">{d.label} (1–5)</label>
          <input
            type="number"
            min="1"
            max="5"
            step="0.5"
            value={form[d.key]}
            onChange={(e) => setForm((p) => ({ ...p, [d.key]: e.target.value }))}
            className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
          />
        </div>
      ))}
      <div className="sm:col-span-2">
        <label className="font-body block text-xs font-medium text-stone-600 mb-1">Overall Rating</label>
        <StarRating value={form.overallRating || 0} onChange={(n) => setForm((p) => ({ ...p, overallRating: n }))} />
      </div>
      <div className="sm:col-span-2">
        <label className="font-body block text-xs font-medium text-stone-600 mb-1">Notes</label>
        <textarea
          value={form.notes}
          onChange={(e) => setForm((p) => ({ ...p, notes: e.target.value }))}
          rows={2}
          className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
      </div>
    </div>
  );
}

const emptyTastingForm = () => ({
  barrelId: "", date: todayISO(), tannin: "", acid: "", body: "", aroma: "", flavorIntensity: "", balance: "", overallRating: 0, notes: "",
});

// ---------- Clickable 1-5 star rating ----------
function StarRating({ value, onChange, size = 20 }) {
  const [hover, setHover] = useState(0);
  return (
    <div className="flex items-center gap-1">
      {[1, 2, 3, 4, 5].map((n) => (
        <button
          key={n}
          type="button"
          onClick={() => onChange(n === value ? 0 : n)}
          onMouseEnter={() => setHover(n)}
          onMouseLeave={() => setHover(0)}
          className="text-amber-500 hover:scale-110 transition-transform"
          title={`${n} star${n === 1 ? "" : "s"}`}
        >
          <Star size={size} fill={(hover || value) >= n ? "currentColor" : "none"} />
        </button>
      ))}
    </div>
  );
}

// ---------- Barrel Tastings sub-section ----------
function TastingsSection({ data, onAdd, onUpdate, onDelete }) {
  const [form, setForm] = useState(emptyTastingForm);
  const [error, setError] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState(null);
  const [barrelFilter, setBarrelFilter] = useState("All");

  const submit = (e) => {
    e.preventDefault();
    if (!form.barrelId) {
      setError("Pick a barrel first");
      return;
    }
    setError("");
    onAdd(form);
    setForm(emptyTastingForm());
  };

  const startEdit = (t) => {
    setEditingId(t.id);
    setEditForm({ ...t });
  };
  const saveEdit = () => {
    onUpdate(editingId, editForm);
    setEditingId(null);
    setEditForm(null);
  };

  const barrelLabel = (id) => data.barrels.find((b) => b.id === id)?.barrelNumber || "Unknown barrel";
  const sorted = [...data.tastings].sort((a, b) => (a.date < b.date ? 1 : -1));
  const filtered = barrelFilter === "All" ? sorted : sorted.filter((t) => t.barrelId === barrelFilter);
  const tastedBarrelIds = [...new Set(data.tastings.map((t) => t.barrelId))];

  return (
    <div className="space-y-6">
      <form onSubmit={submit} className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-3">Log a Tasting</h2>
        <TastingFormFields form={form} setForm={setForm} barrelsList={data.barrels} />
        {error && <p className="font-body text-sm text-red-700 mt-3">{error}</p>}
        <button
          type="submit"
          className="font-body mt-4 flex items-center gap-2 bg-ink-900 hover:bg-ink-800 text-white text-sm font-medium px-4 py-2 rounded-md"
        >
          <Plus size={16} /> Save Tasting
        </button>
      </form>

      <div className="bg-white border border-stone-200 rounded-lg overflow-hidden">
        <div className="px-4 py-3 border-b border-stone-200 flex items-center justify-between flex-wrap gap-2">
          <h2 className="font-brand text-lg text-ink-950">Tasting Notes</h2>
          <select
            value={barrelFilter}
            onChange={(e) => setBarrelFilter(e.target.value)}
            className="font-body text-xs border border-stone-300 rounded px-2 py-1.5"
          >
            <option value="All">All Barrels</option>
            {tastedBarrelIds.map((id) => (
              <option key={id} value={id}>{barrelLabel(id)}</option>
            ))}
          </select>
        </div>
        {filtered.length === 0 ? (
          <p className="font-body text-sm text-stone-500 px-4 py-8 text-center">No tastings logged yet.</p>
        ) : (
          <ul className="divide-y divide-stone-100">
            {filtered.map((t) =>
              editingId === t.id ? (
                <li key={t.id} className="p-4 bg-ink-50">
                  <TastingFormFields form={editForm} setForm={setEditForm} barrelsList={data.barrels} />
                  <div className="flex items-center gap-3 mt-3">
                    <button onClick={saveEdit} className="font-body flex items-center gap-1 text-sm font-medium text-ink-800 hover:text-ink-900">
                      <Check size={15} /> Save
                    </button>
                    <button onClick={() => setEditingId(null)} className="font-body flex items-center gap-1 text-sm text-stone-500 hover:text-stone-700">
                      <X size={15} /> Cancel
                    </button>
                  </div>
                </li>
              ) : (
                <li key={t.id} className="flex items-start gap-3 px-4 py-3 hover:bg-stone-50">
                  <div className="flex text-amber-500 mt-0.5 shrink-0">
                    {[1, 2, 3, 4, 5].map((n) => (
                      <Star key={n} size={14} fill={(t.overallRating || 0) >= n ? "currentColor" : "none"} className={(t.overallRating || 0) >= n ? "" : "text-stone-300"} />
                    ))}
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="font-body text-sm text-stone-900">
                      {barrelLabel(t.barrelId)} <span className="text-stone-400 font-normal">· {t.date}</span>
                    </p>
                    <div className="flex flex-wrap gap-2 mt-1">
                      {RATING_DIMENSIONS.map((d) => (
                        <span key={d.key} className="font-body text-xs bg-stone-100 text-stone-600 px-1.5 py-0.5 rounded">
                          {d.label}: {t[d.key] || "—"}
                        </span>
                      ))}
                    </div>
                    {t.notes && <p className="font-body text-xs text-stone-500 mt-1">{t.notes}</p>}
                  </div>
                  <button onClick={() => startEdit(t)} className="text-stone-300 hover:text-ink-800 shrink-0" title="Edit">
                    <Pencil size={14} />
                  </button>
                  <button onClick={() => onDelete(t.id)} className="text-stone-300 hover:text-red-700 shrink-0" title="Delete">
                    <Trash2 size={14} />
                  </button>
                </li>
              )
            )}
          </ul>
        )}
      </div>
    </div>
  );
}

// ---------- Barrel + percentage picker used when building a blend ----------
function BlendComponentsPicker({ value, onChange, barrelsList }) {
  const [search, setSearch] = useState("");

  const addComponent = (barrelId) => {
    if (value.some((c) => c.barrelId === barrelId)) return;
    onChange([...value, { id: genId(), barrelId, percentage: "" }]);
  };
  const removeComponent = (id) => onChange(value.filter((c) => c.id !== id));
  const updatePercentage = (id, pct) => onChange(value.map((c) => (c.id === id ? { ...c, percentage: pct } : c)));

  const availableBarrels = barrelsList.filter((b) => !value.some((c) => c.barrelId === b.id));
  const filtered = search ? availableBarrels.filter((b) => b.barrelNumber.toLowerCase().includes(search.toLowerCase())) : availableBarrels;
  const totalPct = value.reduce((sum, c) => sum + (parseFloat(c.percentage) || 0), 0);

  return (
    <div>
      <input
        type="text"
        placeholder="Search barrel # to add"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm mb-2 focus:outline-none focus:ring-2 focus:ring-ink-800"
      />
      <div className="bg-stone-50 border border-stone-200 rounded-lg p-2.5">
        <div className="flex flex-wrap gap-2 max-h-32 overflow-y-auto">
          {filtered.length === 0 ? (
            <span className="font-body text-xs text-stone-400 py-1">No barrels match.</span>
          ) : (
            filtered.slice(0, 60).map((b) => (
              <button
                key={b.id}
                type="button"
                onClick={() => addComponent(b.id)}
                className="font-body flex items-center gap-1 text-xs font-medium px-2.5 py-1.5 rounded-full bg-white text-stone-700 border border-stone-200 shadow-sm hover:bg-ink-50 hover:border-ink-300 hover:text-ink-800 transition-colors"
              >
                <Plus size={11} className="text-stone-400" />
                {b.barrelNumber}
              </button>
            ))
          )}
        </div>
      </div>
      {value.length > 0 && (
        <div className="space-y-1.5 mt-2">
          {value.map((c) => {
            const barrel = barrelsList.find((b) => b.id === c.barrelId);
            return (
              <div key={c.id} className="flex items-center gap-2">
                <span className="font-body text-sm w-20 shrink-0">{barrel?.barrelNumber || "?"}</span>
                <input
                  type="number"
                  placeholder="%"
                  value={c.percentage}
                  onChange={(e) => updatePercentage(c.id, e.target.value)}
                  className="font-body w-20 border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                />
                <span className="font-body text-xs text-stone-400">%</span>
                <button type="button" onClick={() => removeComponent(c.id)} className="text-stone-400 hover:text-red-700 ml-auto">
                  <X size={14} />
                </button>
              </div>
            );
          })}
          <p className={`font-body text-xs ${totalPct === 100 ? "text-stone-400" : "text-amber-700"}`}>
            Total: {totalPct}% {totalPct !== 100 && "(doesn't need to add up to exactly 100%)"}
          </p>
        </div>
      )}
    </div>
  );
}

// ---------- One saved blend trial, with its computed rating profile ----------
function BlendCard({ blend, barrelsList, tastings, checked, onToggleCompare, onDelete, onEdit }) {
  const [expanded, setExpanded] = useState(false);
  const result = computeBlendProfile(blend, tastings);

  return (
    <div className="px-4 py-3">
      <div className="flex items-start gap-3">
        <input type="checkbox" checked={checked} onChange={onToggleCompare} className="mt-1.5 rounded border-stone-300 shrink-0" />
        <div className="flex-1 min-w-0">
          <button onClick={() => setExpanded((v) => !v)} className="text-left w-full">
            <p className="font-brand text-ink-950">{blend.name || "Untitled blend"}</p>
            <p className="font-body text-xs text-stone-500 mt-0.5">
              {blend.tier ? `${blend.tier} · ` : ""}{blend.dateCreated} · {(blend.components || []).length} barrel{(blend.components || []).length === 1 ? "" : "s"}
            </p>
          </button>
          {result ? (
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-2 mt-2">
              {RATING_DIMENSIONS.map((d) => (
                <div key={d.key}>
                  <div className="flex justify-between font-body text-xs text-stone-500">
                    <span>{d.label}</span>
                    <span>{result.profile[d.key].toFixed(1)}</span>
                  </div>
                  <div className="h-1.5 bg-stone-100 rounded-full overflow-hidden mt-0.5">
                    <div className="h-full bg-ink-700" style={{ width: `${result.profile[d.key] * 20}%` }} />
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <p className="font-body text-xs text-stone-400 mt-1">No tasting data yet for these barrels.</p>
          )}
          {expanded && (
            <div className="mt-3 pt-3 border-t border-stone-100">
              <p className="font-body text-xs font-semibold text-stone-600 mb-1">Components</p>
              <ul className="font-body text-xs text-stone-600 space-y-0.5">
                {(blend.components || []).map((c) => {
                  const barrel = barrelsList.find((b) => b.id === c.barrelId);
                  const t = latestTastingForBarrel(c.barrelId, tastings);
                  return (
                    <li key={c.id}>
                      {barrel?.barrelNumber || "Unknown barrel"} — {c.percentage || 0}% {t ? `(tasted ${t.date})` : "(not tasted yet)"}
                    </li>
                  );
                })}
              </ul>
              {blend.notes && <p className="font-body text-xs text-stone-500 mt-2">{blend.notes}</p>}
            </div>
          )}
        </div>
        <div className="flex flex-col gap-1.5 shrink-0">
          <button onClick={() => onEdit(blend)} className="text-stone-300 hover:text-ink-800" title="Edit blend">
            <Pencil size={15} />
          </button>
          <button onClick={() => onDelete(blend.id)} className="text-stone-300 hover:text-red-700" title="Delete blend">
            <Trash2 size={15} />
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Side-by-side comparison of 2+ selected blends ----------
function BlendComparisonTable({ blends, tastings }) {
  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
      <h2 className="font-brand text-lg text-ink-950 mb-3 flex items-center gap-2">
        <GitCompare size={18} /> Comparing {blends.length} Blends
      </h2>
      <div className="overflow-x-auto">
        <table className="w-full text-sm font-body">
          <thead>
            <tr className="text-left text-stone-500">
              <th className="py-1.5 pr-4">Dimension</th>
              {blends.map((b) => (
                <th key={b.id} className="py-1.5 pr-4">{b.name || "Untitled"}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {RATING_DIMENSIONS.map((d) => (
              <tr key={d.key} className="border-t border-stone-100">
                <td className="py-1.5 pr-4 font-medium text-stone-700 whitespace-nowrap">{d.label}</td>
                {blends.map((b) => {
                  const result = computeBlendProfile(b, tastings);
                  const val = result ? result.profile[d.key] : null;
                  return (
                    <td key={b.id} className="py-1.5 pr-4">
                      {val != null ? (
                        <div className="flex items-center gap-2">
                          <span className="w-8 shrink-0">{val.toFixed(1)}</span>
                          <div className="flex-1 h-1.5 bg-stone-100 rounded-full overflow-hidden min-w-[60px]">
                            <div className="h-full bg-ink-700" style={{ width: `${val * 20}%` }} />
                          </div>
                        </div>
                      ) : (
                        "—"
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------- Blends sub-section: builder + list + comparison ----------
function BlendsSection({ data, onAdd, onUpdate, onDelete }) {
  const [form, setForm] = useState({ name: "", tier: "", notes: "", components: [] });
  const [error, setError] = useState("");
  const [compareIds, setCompareIds] = useState(new Set());
  const [editingBlendId, setEditingBlendId] = useState(null);

  const startEdit = (blend) => {
    setEditingBlendId(blend.id);
    setForm({ name: blend.name || "", tier: blend.tier || "", notes: blend.notes || "", components: blend.components || [] });
    setError("");
  };
  const cancelEdit = () => {
    setEditingBlendId(null);
    setForm({ name: "", tier: "", notes: "", components: [] });
    setError("");
  };

  const submit = (e) => {
    e.preventDefault();
    if (!form.name.trim()) {
      setError('Please fill in "Blend Name"');
      return;
    }
    if (form.components.length === 0) {
      setError("Add at least one barrel to this blend");
      return;
    }
    setError("");
    if (editingBlendId) {
      onUpdate(editingBlendId, form);
      setEditingBlendId(null);
    } else {
      onAdd(form);
    }
    setForm({ name: "", tier: "", notes: "", components: [] });
  };

  const toggleCompare = (id) =>
    setCompareIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

  const compareBlends = data.blends.filter((b) => compareIds.has(b.id));

  return (
    <div className="space-y-6">
      <form onSubmit={submit} className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-3">{editingBlendId ? "Edit Blend Trial" : "Create a Blend Trial"}</h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
          <div>
            <label className="font-body block text-xs font-medium text-stone-600 mb-1">Blend Name</label>
            <input
              type="text"
              placeholder='e.g. "2025 Estate Trial 1"'
              value={form.name}
              onChange={(e) => setForm((p) => ({ ...p, name: e.target.value }))}
              className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
            />
          </div>
          <div>
            <label className="font-body block text-xs font-medium text-stone-600 mb-1">Wine Tier / Project</label>
            <input
              type="text"
              placeholder="e.g. Estate, Reserve"
              value={form.tier}
              onChange={(e) => setForm((p) => ({ ...p, tier: e.target.value }))}
              className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
            />
          </div>
        </div>
        <label className="font-body block text-xs font-medium text-stone-600 mb-1">Component Barrels</label>
        <BlendComponentsPicker value={form.components} onChange={(c) => setForm((p) => ({ ...p, components: c }))} barrelsList={data.barrels} />
        <div className="mt-3">
          <label className="font-body block text-xs font-medium text-stone-600 mb-1">Notes</label>
          <textarea
            value={form.notes}
            onChange={(e) => setForm((p) => ({ ...p, notes: e.target.value }))}
            rows={2}
            className="font-body w-full border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
          />
        </div>
        {error && <p className="font-body text-sm text-red-700 mt-3">{error}</p>}
        <div className="flex items-center gap-3 mt-4">
          <button
            type="submit"
            className="font-body flex items-center gap-2 bg-ink-900 hover:bg-ink-800 text-white text-sm font-medium px-4 py-2 rounded-md"
          >
            {editingBlendId ? <Check size={16} /> : <Plus size={16} />}
            {editingBlendId ? "Update Blend" : "Save Blend"}
          </button>
          {editingBlendId && (
            <button type="button" onClick={cancelEdit} className="font-body text-sm text-stone-500 hover:text-stone-700">
              Cancel
            </button>
          )}
        </div>
      </form>

      {compareBlends.length >= 2 && <BlendComparisonTable blends={compareBlends} tastings={data.tastings} />}

      <div className="bg-white border border-stone-200 rounded-lg overflow-hidden">
        <div className="px-4 py-3 border-b border-stone-200 flex items-center justify-between">
          <h2 className="font-brand text-lg text-ink-950">Blend Trials</h2>
          <span className="font-body text-xs text-stone-500">Check 2 or more to compare</span>
        </div>
        {data.blends.length === 0 ? (
          <p className="font-body text-sm text-stone-500 px-4 py-8 text-center">No blend trials yet — build one above.</p>
        ) : (
          <div className="divide-y divide-stone-100">
            {data.blends.map((b) => (
              <BlendCard
                key={b.id}
                blend={b}
                barrelsList={data.barrels}
                tastings={data.tastings}
                checked={compareIds.has(b.id)}
                onToggleCompare={() => toggleCompare(b.id)}
                onDelete={onDelete}
                onEdit={startEdit}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ---------- Blending tab: barrel tastings + blend building + comparison ----------
function BlendingTab({ data, onAddTasting, onUpdateTasting, onDeleteTasting, onAddBlend, onUpdateBlend, onDeleteBlend }) {
  const [subTab, setSubTab] = useState("tastings");
  return (
    <div className="space-y-6">
      <div className="flex gap-2">
        <button
          onClick={() => setSubTab("tastings")}
          className={`font-body text-sm font-medium px-4 py-2 rounded-md border ${
            subTab === "tastings" ? "bg-ink-900 text-white border-ink-900" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
          }`}
        >
          Barrel Tastings
        </button>
        <button
          onClick={() => setSubTab("blends")}
          className={`font-body text-sm font-medium px-4 py-2 rounded-md border ${
            subTab === "blends" ? "bg-ink-900 text-white border-ink-900" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
          }`}
        >
          Blends &amp; Comparison
        </button>
      </div>
      {subTab === "tastings" ? (
        <TastingsSection data={data} onAdd={onAddTasting} onUpdate={onUpdateTasting} onDelete={onDeleteTasting} />
      ) : (
        <BlendsSection data={data} onAdd={onAddBlend} onUpdate={onUpdateBlend} onDelete={onDeleteBlend} />
      )}
    </div>
  );
}

// ---------- Local backup / restore, on top of the cloud storage the app already uses ----------
// ---------- Generic add/rename/delete manager for a reference list (Blocks, Vessels, Lots) ----------
function ManageListPanel({ title, description, items, onAdd, onRename, onDelete, confirmAction, addPlaceholder }) {
  const [newItem, setNewItem] = useState("");
  const [editingItem, setEditingItem] = useState(null);
  const [editValue, setEditValue] = useState("");

  const submitAdd = () => {
    const trimmed = newItem.trim();
    if (!trimmed || items.includes(trimmed)) return;
    onAdd(trimmed);
    setNewItem("");
  };

  const startEdit = (item) => {
    setEditingItem(item);
    setEditValue(item);
  };
  const saveEdit = () => {
    const trimmed = editValue.trim();
    if (!trimmed || trimmed === editingItem) {
      setEditingItem(null);
      return;
    }
    onRename(editingItem, trimmed);
    setEditingItem(null);
  };

  const handleDelete = (item) => {
    confirmAction(`Remove "${item}" from the list? Existing records that already use it are left untouched — this only removes it from the picker going forward.`, () => onDelete(item), {
      confirmLabel: "Remove",
      tone: "neutral",
    });
  };

  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
      <h2 className="font-brand text-lg text-ink-950 mb-1">{title}</h2>
      <p className="font-body text-xs text-stone-500 mb-3">{description}</p>
      <div className="flex gap-2 mb-3">
        <input
          type="text"
          placeholder={addPlaceholder}
          value={newItem}
          onChange={(e) => setNewItem(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              submitAdd();
            }
          }}
          className="font-body flex-1 border border-stone-300 rounded-md px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
        <button
          onClick={submitAdd}
          className="font-body flex items-center gap-1 text-sm font-medium bg-ink-900 hover:bg-ink-800 text-white px-3 py-2 rounded-md shrink-0"
        >
          <Plus size={14} /> Add
        </button>
      </div>
      {items.length === 0 ? (
        <p className="font-body text-xs text-stone-400">Nothing in this list yet.</p>
      ) : (
        <ul className="divide-y divide-stone-100 border border-stone-200 rounded-md">
          {items.map((item) => (
            <li key={item} className="flex items-center gap-2 px-3 py-2">
              {editingItem === item ? (
                <>
                  <input
                    type="text"
                    autoFocus
                    value={editValue}
                    onChange={(e) => setEditValue(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") {
                        e.preventDefault();
                        saveEdit();
                      }
                    }}
                    className="font-body flex-1 border border-stone-300 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                  />
                  <button onClick={saveEdit} className="text-ink-700 hover:text-ink-900" title="Save">
                    <Check size={15} />
                  </button>
                  <button onClick={() => setEditingItem(null)} className="text-stone-400 hover:text-stone-600" title="Cancel">
                    <X size={15} />
                  </button>
                </>
              ) : (
                <>
                  <span className="font-body text-sm text-stone-700 flex-1">{item}</span>
                  <button onClick={() => startEdit(item)} className="text-stone-400 hover:text-ink-800" title="Rename">
                    <Pencil size={13} />
                  </button>
                  <button onClick={() => handleDelete(item)} className="text-stone-400 hover:text-red-700" title="Remove">
                    <Trash2 size={14} />
                  </button>
                </>
              )}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

function BackupTab({ data, woCounter, onRestore, confirmAction, vineyardBlocks, onAddBlock, onRenameBlock, onDeleteBlock, vesselTypes, onAddVesselType, onRenameVesselType, onDeleteVesselType, lotNames, onAddLotName, onRenameLotName, onDeleteLotName, clones, onAddClone, onRenameClone, onDeleteClone, sprayPrograms, onAddSprayProgram, onRenameSprayProgram, onDeleteSprayProgram, tastingAssociates, onAddAssociate, onRenameAssociate, onDeleteAssociate, wineryName, onUpdateWineryName }) {
  const [restoreError, setRestoreError] = useState("");
  const [nameDraft, setNameDraft] = useState(wineryName);
  const [nameSaved, setNameSaved] = useState(false);

  const saveName = () => {
    onUpdateWineryName(nameDraft.trim());
    setNameSaved(true);
    setTimeout(() => setNameSaved(false), 1500);
  };

  const downloadBackup = () => {
    const backup = {
      app: "GrapeDaddy",
      backupVersion: 1,
      exportedAt: new Date().toISOString(),
      data: {
        workorders: data.workorders,
        harvest: data.harvest,
        ferment: data.ferment,
        barrels: data.barrels,
        templates: data.templates,
        tastings: data.tastings,
        blends: data.blends,
        weatherLogs: data.weatherLogs,
        techSheets: data.techSheets,
      },
      settings: { woCounter },
    };
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" });
    downloadBlob(blob, `alloro-winery-backup-${todayISO()}.json`);
  };

  const handleFileChange = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const parsed = JSON.parse(event.target.result);
        if (!parsed || typeof parsed !== "object" || !parsed.data || typeof parsed.data !== "object") {
          setRestoreError("This doesn't look like a valid backup file.");
          return;
        }
        setRestoreError("");
        const counts = Object.entries(parsed.data)
          .map(([k, v]) => `${Array.isArray(v) ? v.length : "?"} ${k}`)
          .join(", ");
        const when = parsed.exportedAt ? new Date(parsed.exportedAt).toLocaleString() : "an unknown date";
        confirmAction(
          `Restore the backup from ${when}? It contains: ${counts}. This will replace ALL current data in the app with what's in this file. This can't be undone.`,
          () => onRestore(parsed),
          { confirmLabel: "Restore", tone: "danger" }
        );
      } catch {
        setRestoreError("Couldn't read that file — make sure it's a backup .json file downloaded from this app.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  return (
    <div className="space-y-6">
      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-1">Winery Name</h2>
        <p className="font-body text-xs text-stone-500 mb-3">
          Shown throughout the app in place of a hardcoded winery — set this once when setting the app up for a new winery.
        </p>
        <div className="flex flex-wrap items-center gap-2">
          <input
            type="text"
            value={nameDraft}
            onChange={(e) => setNameDraft(e.target.value)}
            placeholder="e.g. Alloro Vineyard"
            className="font-body border border-stone-300 rounded-md px-3 py-2 text-sm w-full max-w-xs focus:outline-none focus:ring-2 focus:ring-ink-800"
          />
          <button
            onClick={saveName}
            className={`font-body text-sm font-medium px-4 py-2 rounded-md text-white ${nameSaved ? "bg-ink-700" : "bg-ink-900 hover:bg-ink-800"}`}
          >
            {nameSaved ? "✓ Saved" : "Save"}
          </button>
        </div>
      </div>

      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-1">Download a Local Backup</h2>
        <p className="font-body text-xs text-stone-500 mb-4">
          Saves everything in the app — work orders, harvest, fermentation, barrels, blending, mileage, expenses, and
          templates — into one file on your device. This is a raw data backup for safekeeping, separate from the
          formatted Excel/Word/CSV exports used for reports and accounting.
        </p>
        <button
          onClick={downloadBackup}
          className="font-body flex items-center gap-2 bg-ink-900 hover:bg-ink-800 text-white text-sm font-medium px-4 py-2 rounded-md"
        >
          <HardDrive size={16} /> Download Full Backup (.json)
        </button>
      </div>

      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-1">Restore from a Backup</h2>
        <p className="font-body text-xs text-stone-500 mb-4">
          Upload a backup file downloaded from this app to restore its data.{" "}
          <span className="font-semibold text-red-700">This replaces everything currently in the app.</span> Use this
          to recover from a data loss or move data between deployments — not casually.
        </p>
        <label className="font-body inline-flex items-center gap-2 bg-white border border-stone-300 hover:border-ink-400 text-stone-700 text-sm font-medium px-4 py-2 rounded-md cursor-pointer">
          <UploadCloud size={16} />
          Choose Backup File
          <input type="file" accept="application/json,.json" onChange={handleFileChange} className="hidden" />
        </label>
        {restoreError && <p className="font-body text-xs text-red-700 mt-2">{restoreError}</p>}
      </div>

      <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5">
        <h2 className="font-brand text-lg text-ink-950 mb-2">Cloud + Local, Together</h2>
        <p className="font-body text-xs text-stone-500">
          Day-to-day, this app's data already lives in the cloud — everyone on the crew sees the same live data the
          moment it's entered, from any device. The backup above is a local safety net on top of that: an independent
          copy on your own computer, so you're never solely dependent on one system.
        </p>
      </div>

      <div>
        <h2 className="font-brand text-xl text-ink-950 mb-1">Manage Lists</h2>
        <p className="font-body text-xs text-stone-500 mb-3">
          Add, rename, or remove the options that show up in dropdowns across the app. Renaming updates every existing
          record that uses it; removing only takes it out of the picker going forward — nothing already entered is touched.
        </p>
      </div>

      <ManageListPanel
        title="Blocks / Vineyards"
        description="Used in Harvest Tonnage and Fruit Analysis."
        items={vineyardBlocks}
        onAdd={onAddBlock}
        onRename={onRenameBlock}
        onDelete={onDeleteBlock}
        confirmAction={confirmAction}
        addPlaceholder="New block / vineyard name"
      />

      <ManageListPanel
        title="Vessels"
        description="Specific vessel IDs (e.g. B10, T5) used on Fermentation lots — matches whatever numbering your cellar actually uses."
        items={vesselTypes}
        onAdd={onAddVesselType}
        onRename={onRenameVesselType}
        onDelete={onDeleteVesselType}
        confirmAction={confirmAction}
        addPlaceholder="New vessel ID, e.g. B10"
      />

      <ManageListPanel
        title="Lots"
        description="Used when linking Work Orders and barrel fills to a specific lot."
        items={lotNames}
        onAdd={onAddLotName}
        onRename={onRenameLotName}
        onDelete={onDeleteLotName}
        confirmAction={confirmAction}
        addPlaceholder="New lot name"
      />

      <ManageListPanel
        title="Clones"
        description="Used in Harvest Tonnage."
        items={clones}
        onAdd={onAddClone}
        onRename={onRenameClone}
        onDelete={onDeleteClone}
        confirmAction={confirmAction}
        addPlaceholder="New clone, e.g. 667"
      />

      <ManageListPanel
        title="Spray Programs"
        description="Used on Vineyard work orders tagged Spray / Pest Management."
        items={sprayPrograms}
        onAdd={onAddSprayProgram}
        onRename={onRenameSprayProgram}
        onDelete={onDeleteSprayProgram}
        confirmAction={confirmAction}
        addPlaceholder="New spray program"
      />

      <ManageListPanel
        title="Tasting Associates"
        description="Used in THO Timesheets — keeping names consistent here is what lets the Payout Calculator match hours to tips correctly."
        items={tastingAssociates}
        onAdd={onAddAssociate}
        onRename={onRenameAssociate}
        onDelete={onDeleteAssociate}
        confirmAction={confirmAction}
        addPlaceholder="New tasting associate"
      />
    </div>
  );
}

// Single admin password for now — reads from a Vite environment variable when deployed
// (VITE_ADMIN_PASSWORD, set in Vercel), falling back to the default if not set. The two-role
// system (admin/tho) is still here underneath, just collapsed to one login for now — every
// successful login grants "admin", so re-adding a second, view-only password later is a small
// change, not a rebuild.
const GDS_LOGO = "data:image/png;base64,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";

let ADMIN_PASSWORD = "AlloroCrew2026";
try {
  if (import.meta && import.meta.env && import.meta.env.VITE_ADMIN_PASSWORD) {
    ADMIN_PASSWORD = import.meta.env.VITE_ADMIN_PASSWORD;
  }
} catch {
  // import.meta not available here — stick with the default
}

// Storage keys a "tho" role login would be allowed to write, once that role is reintroduced.
// Not currently reachable — kept here so the guard logic and nav filtering below don't need to
// be rewritten when the second login comes back.
const THO_EDITABLE_STORAGE_KEYS = ["thoTimesheets", "thoTips", "thoMileage", "thoExpenses", "tasting_associates"];
const THO_EDITABLE_NAV_KEYS = ["thoPayroll", "thoMileage", "thoExpenses"];

// Reads/writes which role is currently signed in ("admin" | null — "tho" isn't reachable right
// now). Wrapped in try/catch since localStorage isn't available in every environment this file
// runs in (e.g. it's blocked inside Claude artifacts) — falls back to a session-only login there.
function readStoredRole() {
  try {
    const v = window.localStorage.getItem("winery_role");
    return v === "admin" || v === "tho" ? v : null;
  } catch {
    return null;
  }
}
function writeStoredRole(role) {
  try {
    if (role) window.localStorage.setItem("winery_role", role);
    else window.localStorage.removeItem("winery_role");
  } catch {
    // no persistent storage available — the login will just last for this session
  }
}

// ---------- Single-password login gate shown before the app loads ----------
function PasswordGate({ onUnlock, wineryName }) {
  const [value, setValue] = useState("");
  const [error, setError] = useState("");

  const submit = (e) => {
    e.preventDefault();
    if (value === ADMIN_PASSWORD) {
      setError("");
      onUnlock("admin");
    } else {
      setError("That's not the right password — try again.");
    }
  };

  return (
    <div className="min-h-screen bg-ink-950 flex items-center justify-center p-4">
      <form onSubmit={submit} className="bg-white rounded-lg p-6 sm:p-8 max-w-sm w-full">
        <img src={GDS_LOGO} alt="GrapeDaddy" className="w-32 h-32 mx-auto mb-3" />
        <h1 className="font-brand text-2xl text-ink-950 text-center mb-0.5">{wineryName || "Your Winery"}</h1>
        <p className="font-body text-xs text-stone-400 text-center tracking-wide mb-4">GrapeDaddy</p>
        <p className="font-body text-sm text-stone-500 mb-5 text-center">Enter your password to continue.</p>
        <input
          type="password"
          autoFocus
          value={value}
          onChange={(e) => setValue(e.target.value)}
          placeholder="Password"
          className="font-body w-full border border-stone-300 rounded-md px-3 py-2.5 text-sm mb-3 focus:outline-none focus:ring-2 focus:ring-ink-800"
        />
        {error && <p className="font-body text-sm text-red-700 mb-3">{error}</p>}
        <button
          type="submit"
          className="font-body w-full bg-ink-900 hover:bg-ink-800 text-white text-sm font-medium px-4 py-2.5 rounded-md"
        >
          Unlock
        </button>
      </form>
    </div>
  );
}

// Catches any crash anywhere in the app and shows the actual error instead of leaving a blank
// white screen — without this, a single unhandled error unmounts the entire page silently,
// which is exactly what made an earlier bug so hard to diagnose from a description alone.
class AppErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error) {
    return { error };
  }
  componentDidCatch(error, info) {
    console.error("Alloro Winery Tracker crashed:", error, info);
  }
  render() {
    if (this.state.error) {
      return (
        <div className="min-h-screen bg-stone-50 flex items-center justify-center p-4">
          <div className="bg-white border border-red-200 rounded-lg p-6 max-w-lg w-full">
            <h1 className="font-brand text-xl text-red-800 mb-2">Something went wrong</h1>
            <p className="font-body text-sm text-stone-600 mb-3">
              The app hit an error and couldn't continue. The actual error message is below — screenshot this and send it over, it'll say exactly what broke.
            </p>
            <pre className="font-mono text-xs text-red-700 bg-red-50 border border-red-100 rounded p-3 whitespace-pre-wrap break-words mb-4">
              {String(this.state.error?.message || this.state.error)}
            </pre>
            <button
              onClick={() => window.location.reload()}
              className="font-body bg-ink-900 hover:bg-ink-800 text-white text-sm font-medium px-4 py-2 rounded-md"
            >
              Reload the page
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

function WineryDataTrackerInner() {
  const [role, setRole] = useState(() => readStoredRole());
  const [activeKey, setActiveKey] = useState(ALL_TABS[0].key);
  useEffect(() => {
    if (role === "tho" && (activeKey === "workorders" || activeKey === "backup")) {
      setActiveKey("home");
    }
  }, [role, activeKey]);
  const [data, setData] = useState(
    Object.fromEntries([["workorders", []], ...SIMPLE_SECTIONS.map((s) => [s.key, []]), ["ferment", []], ["barrels", []], ["templates", []], ["tastings", []], ["blends", []], ["weatherLogs", []], ["techSheets", []], ["historyMilestones", []], ["accolades", []], ["vineyardBlockDetails", []]])
  );
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [form, setForm] = useState(() => ({ ...emptyForm(SIMPLE_SECTIONS[0].fields), date: todayISO() }));
  const [showBulkImport, setShowBulkImport] = useState(false);
  const [techSheetMode, setTechSheetMode] = useState({ mode: "list", sheetId: null }); // "list" | "form" | "view"
  const [workOrderFilterCategory, setWorkOrderFilterCategory] = useState("Winery");
  const [newWorkOrderForm, setNewWorkOrderForm] = useState({ ...emptyForm(WORKORDER_FIELDS), category: "Winery", date: todayISO(), dateAssigned: todayISO() });
  const [so2Calc, setSo2Calc] = useState({ volume: "", current: "", target: "" });
  const [acidCalc, setAcidCalc] = useState({ volume: "", current: "", target: "" });
  const [waterCalc, setWaterCalc] = useState({ volume: "", current: "", target: "" });
  const [newLotForm, setNewLotForm] = useState({ ...emptyForm(FERMENT_LOT_FIELDS), startDate: todayISO() });
  const [error, setError] = useState("");
  const [saveError, setSaveError] = useState("");
  const [exportMenuOpen, setExportMenuOpen] = useState(false);
  const [lastTabByCategory, setLastTabByCategory] = useState({});
  useEffect(() => {
    const cat = categoryOfKey(activeKey);
    if (cat) {
      setLastTabByCategory((prev) => (prev[cat] === activeKey ? prev : { ...prev, [cat]: activeKey }));
    }
  }, [activeKey]);
  const goToCategory = (catKey) => {
    const remembered = lastTabByCategory[catKey];
    const cat = NAV_CATEGORIES[catKey];
    setActiveKey(remembered && cat.keys.includes(remembered) ? remembered : cat.keys[0]);
  };
  const [confirmRequest, setConfirmRequest] = useState(null); // { message, onConfirm }
  const confirmAction = (message, onConfirm, options = {}) =>
    setConfirmRequest({ message, onConfirm, confirmLabel: options.confirmLabel || "Delete", tone: options.tone || "danger" });
  const [woCounter, setWoCounter] = useState(1);
  const [vineyardBlocks, setVineyardBlocks] = useState(VINEYARD_BLOCKS);
  const [clones, setClones] = useState(GRAPE_CLONES);
  const [sprayPrograms, setSprayPrograms] = useState(SPRAY_PROGRAMS);
  const [tastingAssociates, setTastingAssociates] = useState(TASTING_ASSOCIATES);
  const [defaultTareWeight, setDefaultTareWeight] = useState("");
  const [alloroStory, setAlloroStory] = useState("");
  const [wineryName, setWineryName] = useState("");
  const [vineyardMapImage, setVineyardMapImage] = useState("");
  const [vesselTypes, setVesselTypes] = useState(VESSEL_TYPES);
  const [lotNames, setLotNames] = useState([]);
  const [workOrderSort, setWorkOrderSort] = useState({ field: "", direction: "asc" });
  const [fermentSort, setFermentSort] = useState({ field: "", direction: "asc" });
  const [fermentViewMode, setFermentViewMode] = useState("overview");
  const [tableSort, setTableSort] = useState({}); // { [sectionKey]: { field, direction } }

  // Inline editing state
  const [editingRow, setEditingRow] = useState(null); // { key, id, form }
  const [editingWorkOrderId, setEditingWorkOrderId] = useState(null);
  const [editWorkOrderForm, setEditWorkOrderForm] = useState(null);

  // Print-by-range / print-by-lot state
  const [printJob, setPrintJob] = useState(null);
  const [printRanges, setPrintRanges] = useState({
    harvest: { from: "", to: "" },
    mileage: { from: "", to: "" },
    expenses: { from: "", to: "" },
  });
  const [fermentPrintFrom, setFermentPrintFrom] = useState("");
  const [fermentPrintTo, setFermentPrintTo] = useState("");
  const [fermentPrintLotIds, setFermentPrintLotIds] = useState([]);

  const activeSection = SIMPLE_SECTIONS.find((s) => s.key === activeKey);

  useEffect(() => {
    if (printJob) window.print();
  }, [printJob]);

  // The winery name is the one piece of data that needs to be available before login (it shows
  // on the login screen itself) — it's just a display label, not sensitive business data, so
  // fetching it unconditionally on mount doesn't reintroduce the "don't load before login" issue
  // the main data effect below was fixed for.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await storage.get("winery_name", true);
        if (res && !cancelled) setWineryName(res.value);
      } catch {
        // stays at "" — falls back to "GrapeDaddy" in the UI
      }
    })();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!role) return; // nothing on the login screen needs this — don't do any of it until signed in
    let cancelled = false;
    (async () => {
      const results = {};
      for (const tab of ALL_TABS) {
        if (tab.key === "calendar" || tab.key === "formulas" || tab.key === "home" || tab.key === "blending" || tab.key === "backup") continue;
        try {
          const res = await storage.get(tab.key, true);
          results[tab.key] = res ? JSON.parse(res.value) : [];
        } catch {
          results[tab.key] = [];
        }
      }
      try {
        const res = await storage.get("templates", true);
        results.templates = res ? JSON.parse(res.value) : [];
      } catch {
        results.templates = [];
      }

      try {
        const res = await storage.get("tastings", true);
        results.tastings = res ? JSON.parse(res.value) : [];
      } catch {
        results.tastings = [];
      }

      try {
        const res = await storage.get("blends", true);
        results.blends = res ? JSON.parse(res.value) : [];
      } catch {
        results.blends = [];
      }

      try {
        const res = await storage.get("weatherLogs", true);
        results.weatherLogs = res ? JSON.parse(res.value) : [];
      } catch {
        results.weatherLogs = [];
      }

      try {
        const res = await storage.get("techSheets", true);
        results.techSheets = res ? JSON.parse(res.value) : [];
      } catch {
        results.techSheets = [];
      }

      // thoTimesheets and thoTips aren't in ALL_TABS (they're only reachable through the
      // combined Payroll page, not as independent nav tabs), so the generic ALL_TABS loading
      // loop above skips them — they need to be loaded explicitly here instead.
      try {
        const res = await storage.get("thoTimesheets", true);
        results.thoTimesheets = res ? JSON.parse(res.value) : [];
      } catch {
        results.thoTimesheets = [];
      }

      try {
        const res = await storage.get("thoTips", true);
        results.thoTips = res ? JSON.parse(res.value) : [];
      } catch {
        results.thoTips = [];
      }

      try {
        const res = await storage.get("historyMilestones", true);
        results.historyMilestones = res ? JSON.parse(res.value) : [];
      } catch {
        results.historyMilestones = [];
      }

      try {
        const res = await storage.get("accolades", true);
        results.accolades = res ? JSON.parse(res.value) : [];
      } catch {
        results.accolades = [];
      }

      try {
        const res = await storage.get("vineyardBlockDetails", true);
        results.vineyardBlockDetails = res ? JSON.parse(res.value) : [];
      } catch {
        results.vineyardBlockDetails = [];
      }

      try {
        const res = await storage.get("contacts", true);
        results.contacts = res ? JSON.parse(res.value) : [];
      } catch {
        results.contacts = [];
      }

      try {
        const res = await storage.get("winePricing", true);
        results.winePricing = res ? JSON.parse(res.value) : [];
      } catch {
        results.winePricing = [];
      }

      try {
        const res = await storage.get("wineClubTiers", true);
        results.wineClubTiers = res ? JSON.parse(res.value) : [];
      } catch {
        results.wineClubTiers = [];
      }

      // statePermits, shipmentLog, and stateRuleReference aren't in ALL_TABS (only reachable
      // through the Compliance tab's internal sub-tabs, not as independent nav entries), so the
      // generic ALL_TABS loading loop skips them — load explicitly here instead.
      try {
        const res = await storage.get("statePermits", true);
        results.statePermits = res ? JSON.parse(res.value) : [];
      } catch {
        results.statePermits = [];
      }

      try {
        const res = await storage.get("shipmentLog", true);
        results.shipmentLog = res ? JSON.parse(res.value) : [];
      } catch {
        results.shipmentLog = [];
      }

      try {
        const res = await storage.get("stateRuleReference", true);
        results.stateRuleReference = res ? JSON.parse(res.value) : [];
      } catch {
        results.stateRuleReference = [];
      }

      try {
        const res = await storage.get("vineyard_blocks", true);
        if (res && !cancelled) setVineyardBlocks(JSON.parse(res.value));
      } catch {
        // First time — seed shared storage with the default block list so everyone starts
        // from the same set, and additions from here on just append to it.
        storage.set("vineyard_blocks", JSON.stringify(VINEYARD_BLOCKS), true).catch(() => {});
      }

      try {
        const res = await storage.get("clones", true);
        if (res && !cancelled) setClones(JSON.parse(res.value));
      } catch {
        storage.set("clones", JSON.stringify(GRAPE_CLONES), true).catch(() => {});
      }

      try {
        const res = await storage.get("spray_programs", true);
        if (res && !cancelled) setSprayPrograms(JSON.parse(res.value));
      } catch {
        storage.set("spray_programs", JSON.stringify(SPRAY_PROGRAMS), true).catch(() => {});
      }

      try {
        const res = await storage.get("tasting_associates", true);
        if (res && !cancelled) setTastingAssociates(JSON.parse(res.value));
      } catch {
        storage.set("tasting_associates", JSON.stringify(TASTING_ASSOCIATES), true).catch(() => {});
      }

      try {
        const res = await storage.get("default_tare_weight", true);
        if (res && !cancelled) setDefaultTareWeight(res.value);
      } catch {
        // stays at "" (no default set yet)
      }

      try {
        const res = await storage.get("alloro_story", true);
        if (res && !cancelled) setAlloroStory(res.value);
      } catch {
        // stays at "" (no story written yet)
      }

      try {
        const res = await storage.get("vineyard_map_image", true);
        if (res && !cancelled) setVineyardMapImage(res.value);
      } catch {
        // stays at "" (no map uploaded yet)
      }

      try {
        const res = await storage.get("vessel_types", true);
        if (res && !cancelled) setVesselTypes(JSON.parse(res.value));
      } catch {
        storage.set("vessel_types", JSON.stringify(VESSEL_TYPES), true).catch(() => {});
      }

      try {
        const res = await storage.get("lot_names", true);
        if (res && !cancelled) setLotNames(JSON.parse(res.value));
      } catch {
        // First time — seed the master lot list from whatever Fermentation lot names already
        // exist, so nothing that's already in use "disappears" from the picker.
        const existingNames = [...new Set((results.ferment || []).map((l) => l.tankId).filter(Boolean))].sort((a, b) => a.localeCompare(b));
        if (!cancelled) setLotNames(existingNames);
        storage.set("lot_names", JSON.stringify(existingNames), true).catch(() => {});
      }

      // Load the work-order numbering counter, and make sure it's always ahead of any
      // existing order numbers (covers orders created before this feature existed, or if
      // the counter itself never got saved for some reason).
      let counterValue = 1;
      try {
        const res = await storage.get("wo_counter", true);
        counterValue = res ? parseInt(res.value, 10) || 1 : 1;
      } catch {
        counterValue = 1;
      }
      const maxExisting = (results.workorders || []).reduce((max, o) => Math.max(max, o.orderNumber || 0), 0);
      counterValue = Math.max(counterValue, maxExisting + 1);
      if (!cancelled) setWoCounter(counterValue);

      if (!cancelled) {
        // One-time cleanup: the Barrels section previously had data that failed to import
        // correctly from another app. Wipe it once (and mark it done) so it doesn't keep
        // resetting on every load — new barrel data entered here going forward is safe.
        try {
          await storage.get("barrels_reset_v1", true);
        } catch {
          results.barrels = [];
          try {
            await storage.set("barrels", "[]", true);
            await storage.set("barrels_reset_v1", "true", true);
          } catch {
            // ignore — worst case the cleanup is retried next load
          }
        }

        // One-time seed: import the 2025 barrel program (barrels + matching Fermentation lots)
        try {
          await storage.get("barrels_seed_2025_v1", true);
        } catch {
          try {
            const { fermentLots, barrels } = buildBarrelImport2025();
            const mergedFerment = [...fermentLots, ...(results.ferment || [])];
            const mergedBarrels = [...barrels, ...(results.barrels || [])];
            await storage.set("ferment", JSON.stringify(mergedFerment), true);
            await storage.set("barrels", JSON.stringify(mergedBarrels), true);
            await storage.set("barrels_seed_2025_v1", "true", true);
            results.ferment = mergedFerment;
            results.barrels = mergedBarrels;
          } catch {
            // ignore — worst case the import is retried next load
          }
        }

        // One-time seed: import the two sample lab reports (ETS Laboratories + Core Enology)
        // shared during setup, so Lab Results starts with real data instead of an empty tab.
        try {
          await storage.get("lab_results_seed_v1", true);
        } catch {
          try {
            const imported = buildLabResultsImport();
            const mergedLabResults = [...imported, ...(results.labResults || [])];
            await storage.set("labResults", JSON.stringify(mergedLabResults), true);
            await storage.set("lab_results_seed_v1", "true", true);
            results.labResults = mergedLabResults;
          } catch {
            // ignore — worst case the import is retried next load
          }
        }

        // One-time seed: populate the State Reference table with all 50 states + DC so the
        // full structure exists to browse and fill in. Only the handful of states with a
        // genuinely known status get one seeded here — everything else starts honestly marked
        // "Not Yet Verified" rather than guessed, since this table's whole purpose is to never
        // assert something that hasn't actually been checked against a current source.
        try {
          await storage.get("state_reference_seed_v1", true);
        } catch {
          try {
            const KNOWN_STATUSES = {
              "Utah": { dtcStatus: "Banned", restrictionNotes: "DTC wine shipping banned outright." },
              "Delaware": { dtcStatus: "Banned", restrictionNotes: "DTC wine shipping banned outright." },
              "Indiana": { dtcStatus: "Restricted", restrictionNotes: "Allowed with major restrictions — verify current specifics before shipping." },
              "Mississippi": { dtcStatus: "Restricted", restrictionNotes: "Allowed with major restrictions. Rules were loosened in 2025 — recheck current specifics before relying on this." },
              "New Jersey": { dtcStatus: "Restricted", restrictionNotes: "Allowed with a 250,000-gallon winery production cap. Legislative bills were pending as of 2026 — recheck before relying on this." },
              "Oklahoma": { dtcStatus: "Restricted", restrictionNotes: "Allowed with major restrictions — verify current specifics before shipping." },
              "Rhode Island": { dtcStatus: "Restricted", restrictionNotes: "Allowed with major restrictions — verify current specifics before shipping." },
              "Wyoming": { dtcStatus: "Restricted", restrictionNotes: "Allowed with major restrictions — verify current specifics before shipping." },
              "Louisiana": { dtcStatus: "Restricted", restrictionNotes: "Allowed with major restrictions — verify current specifics before shipping." },
              "Arkansas": { dtcStatus: "Not Yet Verified", restrictionNotes: "Rules were loosened in 2025 — status not yet re-verified here, check before relying on this." },
            };
            const seedRows = US_STATES.map((state) => {
              const known = KNOWN_STATUSES[state];
              return {
                id: genId(),
                state,
                dtcStatus: known ? known.dtcStatus : "Not Yet Verified",
                restrictionNotes: known ? known.restrictionNotes : "",
                annualCapPerHousehold: "",
                reportFrequency: "",
                lastVerifiedDate: "",
                sourceLink: "",
              };
            });
            const mergedReference = [...seedRows, ...(results.stateRuleReference || [])];
            await storage.set("stateRuleReference", JSON.stringify(mergedReference), true);
            results.stateRuleReference = mergedReference;
            await storage.set("state_reference_seed_v1", "true", true);
          } catch {
            // ignore — worst case the import is retried next load
          }
        }

        setData(results);
        setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [role]);

  useEffect(() => {
    if (activeSection) setForm({ ...emptyForm(activeSection.fields), date: todayISO() });
    setError("");
    setEditingRow(null);
    setEditingWorkOrderId(null);
    setEditWorkOrderForm(null);
  }, [activeKey]);

  // Every write in the app funnels through here (or through persist() below, which also uses
  // it) — this is the actual enforcement point for the two-role login, not the UI. Deny by
  // default: a "tho" role can only write the handful of keys explicitly listed as theirs.
  const guardedStorageSet = (key, value, shared) => {
    if (role === "tho" && !THO_EDITABLE_STORAGE_KEYS.includes(key)) {
      setSaveError("You're signed in with view-only access for this section — that change wasn't saved.");
      return Promise.resolve(null);
    }
    return storage.set(key, value, shared);
  };

  const persist = async (key, updated) => {
    if (role === "tho" && !THO_EDITABLE_STORAGE_KEYS.includes(key)) {
      setSaveError("You're signed in with view-only access for this section — that change wasn't saved.");
      return;
    }
    try {
      const result = await storage.set(key, JSON.stringify(updated), true);
      if (!result) {
        setSaveError("Your last change didn't save — check your connection and try again.");
      }
    } catch {
      setSaveError("Your last change didn't save — check your connection and try again.");
    }
  };

  // Restores every section from an uploaded backup file, writing to cloud storage and updating
  // the app's live state in the same pass so nothing is left out of sync.
  const restoreBackup = async (parsed) => {
    const b = parsed.data || {};
    const restored = {
      ...data,
      workorders: Array.isArray(b.workorders) ? b.workorders : data.workorders,
      harvest: Array.isArray(b.harvest) ? b.harvest : data.harvest,
      ferment: Array.isArray(b.ferment) ? b.ferment : data.ferment,
      barrels: Array.isArray(b.barrels) ? b.barrels : data.barrels,
      templates: Array.isArray(b.templates) ? b.templates : data.templates,
      tastings: Array.isArray(b.tastings) ? b.tastings : data.tastings,
      blends: Array.isArray(b.blends) ? b.blends : data.blends,
      weatherLogs: Array.isArray(b.weatherLogs) ? b.weatherLogs : data.weatherLogs,
      techSheets: Array.isArray(b.techSheets) ? b.techSheets : data.techSheets,
    };
    setData(restored);
    const restorableKeys = ["workorders", "harvest", "ferment", "barrels", "templates", "tastings", "blends", "weatherLogs", "techSheets"];
    for (const key of restorableKeys) {
      await persist(key, restored[key]);
    }
    if (parsed.settings) {
      if (typeof parsed.settings.woCounter === "number") {
        setWoCounter(parsed.settings.woCounter);
        storage.set("wo_counter", String(parsed.settings.woCounter), true).catch(() => {});
      }
    }
  };

  // Returns a copy of `data` with Work Orders, Harvest, and Fermentation each sorted per their
  // current sort controls — used both on screen and by every export/print path, so exports
  // always reflect whatever sort order is currently chosen.
  const getSortedData = () => ({
    ...data,
    workorders: sortRows(data.workorders, workOrderSort.field, workOrderSort.direction),
    harvest: tableSort.harvest ? sortRows(data.harvest, tableSort.harvest.field, tableSort.harvest.direction) : data.harvest,
    ferment: sortRows(data.ferment, fermentSort.field, fermentSort.direction),
  });

  const toggleTableSort = (key, field) => {
    setTableSort((prev) => {
      const current = prev[key];
      if (current && current.field === field) {
        return { ...prev, [key]: { field, direction: current.direction === "asc" ? "desc" : "asc" } };
      }
      return { ...prev, [key]: { field, direction: "asc" } };
    });
  };

  const handleChange = (name, value) => {
    setForm((prev) => {
      const next = { ...prev, [name]: value };
      if (activeKey === "harvest" && (name === "tons" || name === "lbs" || name === "tareWeight")) {
        if (name === "tons" || name === "lbs") {
          if (value === "") {
            next[name === "tons" ? "lbs" : "tons"] = "";
          } else {
            const num = parseFloat(value);
            if (!isNaN(num)) {
              if (name === "tons") next.lbs = String(Math.round(num * 2000 * 100) / 100);
              else next.tons = String(Math.round((num / 2000) * 1000) / 1000);
            }
          }
        }
        const grossLbs = parseFloat(next.lbs);
        const tare = parseFloat(next.tareWeight);
        if (!isNaN(grossLbs) && !isNaN(tare) && grossLbs >= tare) {
          const netLbsVal = grossLbs - tare;
          next.netLbs = String(Math.round(netLbsVal * 100) / 100);
          next.netTons = String(Math.round((netLbsVal / 2000) * 1000) / 1000);
        } else {
          next.netLbs = "";
          next.netTons = "";
        }
      }
      if (activeKey === "thoMileage" && (name === "odometerStart" || name === "odometerEnd")) {
        const start = parseFloat(next.odometerStart);
        const end = parseFloat(next.odometerEnd);
        if (!isNaN(start) && !isNaN(end) && end >= start) {
          next.miles = String(Math.round((end - start) * 10) / 10);
        }
      }
      return next;
    });
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    const missing = activeSection.fields.find((f) => f.type !== "textarea" && !f.optional && !String(form[f.name]).trim());
    if (missing) {
      setError(`Please fill in "${missing.label}"`);
      return;
    }
    setError("");
    const entry = { id: genId(), ...form };
    const updated = [entry, ...data[activeKey]];
    setSaving(true);
    await persist(activeKey, updated);
    setData((prev) => ({ ...prev, [activeKey]: updated }));
    setForm({ ...emptyForm(activeSection.fields), date: todayISO() });
    setSaving(false);
  };

  // Adds a batch of rows parsed from a pasted CSV/table into whichever section is active
  const bulkImportRows = async (rows) => {
    const entries = rows.map((r) => ({ id: genId(), ...emptyForm(activeSection.fields), ...r }));
    const updated = [...entries, ...data[activeKey]];
    setData((prev) => ({ ...prev, [activeKey]: updated }));
    await persist(activeKey, updated);
  };

  // Generic add/update/delete for the combined Payroll page, which shows Timesheets and Tips
  // together on one page rather than through the single-activeKey generic table.
  const addThoEntry = async (key, entry) => {
    const updated = [entry, ...data[key]];
    setData((prev) => ({ ...prev, [key]: updated }));
    await persist(key, updated);
  };
  const updateThoEntry = async (key, id, changes) => {
    const updated = data[key].map((r) => (r.id === id ? { ...r, ...changes } : r));
    setData((prev) => ({ ...prev, [key]: updated }));
    await persist(key, updated);
  };
  const deleteThoEntry = async (key, id) => {
    const updated = data[key].filter((r) => r.id !== id);
    setData((prev) => ({ ...prev, [key]: updated }));
    await persist(key, updated);
  };

  // Builds one Harvest Tonnage record per bin, all sharing the same header (block, variety,
  // date, clone, weigh master, notes) — matches weighing multiple bins from the same pick.
  const submitHarvestBatch = async (header, bins) => {
    const entries = bins.map((bin) => ({
      id: genId(),
      ...header,
      tons: bin.tons,
      lbs: bin.lbs,
      tareWeight: bin.tareWeight,
      netTons: bin.netTons,
      netLbs: bin.netLbs,
    }));
    const updated = [...entries, ...data.harvest];
    setSaving(true);
    setData((prev) => ({ ...prev, harvest: updated }));
    await persist("harvest", updated);
    setSaving(false);
  };

  const handleDelete = (id) => {
    confirmAction("Delete this entry? This can't be undone.", async () => {
      const updated = data[activeKey].filter((row) => row.id !== id);
      setData((prev) => ({ ...prev, [activeKey]: updated }));
      await persist(activeKey, updated);
    });
  };

  // ---- Inline row editing for Harvest / Mileage / Expenses ----
  const startEditRow = (key, row) => setEditingRow({ key, id: row.id, form: { ...row } });
  const cancelEditRow = () => setEditingRow(null);
  const editRowChange = (name, value) => {
    setEditingRow((prev) => {
      if (!prev) return prev;
      const next = { ...prev.form, [name]: value };
      if (prev.key === "harvest" && (name === "tons" || name === "lbs" || name === "tareWeight")) {
        if (name === "tons" || name === "lbs") {
          if (value === "") {
            next[name === "tons" ? "lbs" : "tons"] = "";
          } else {
            const num = parseFloat(value);
            if (!isNaN(num)) {
              if (name === "tons") next.lbs = String(Math.round(num * 2000 * 100) / 100);
              else next.tons = String(Math.round((num / 2000) * 1000) / 1000);
            }
          }
        }
        const grossLbs = parseFloat(next.lbs);
        const tare = parseFloat(next.tareWeight);
        if (!isNaN(grossLbs) && !isNaN(tare) && grossLbs >= tare) {
          const netLbsVal = grossLbs - tare;
          next.netLbs = String(Math.round(netLbsVal * 100) / 100);
          next.netTons = String(Math.round((netLbsVal / 2000) * 1000) / 1000);
        } else {
          next.netLbs = "";
          next.netTons = "";
        }
      }
      if (prev.key === "thoMileage" && (name === "odometerStart" || name === "odometerEnd")) {
        const start = parseFloat(next.odometerStart);
        const end = parseFloat(next.odometerEnd);
        if (!isNaN(start) && !isNaN(end) && end >= start) {
          next.miles = String(Math.round((end - start) * 10) / 10);
        }
      }
      return { ...prev, form: next };
    });
  };
  const saveEditRow = async () => {
    if (!editingRow) return;
    const section = SIMPLE_SECTIONS.find((s) => s.key === editingRow.key);
    const missing = section.fields.find((f) => f.type !== "textarea" && f.type !== "photo" && !f.optional && !String(editingRow.form[f.name] ?? "").trim());
    if (missing) {
      setError(`Please fill in "${missing.label}"`);
      return;
    }
    setError("");
    const updated = data[editingRow.key].map((row) => (row.id === editingRow.id ? { ...row, ...editingRow.form } : row));
    setData((prev) => ({ ...prev, [editingRow.key]: updated }));
    await persist(editingRow.key, updated);
    setEditingRow(null);
  };

  // ---- Work order handlers ----
  // Hands out the next sequential work order number and persists the counter so it's
  // never reused, even if the order that had it gets deleted later.
  const nextOrderNumber = () => {
    const num = woCounter;
    const newCounter = woCounter + 1;
    setWoCounter(newCounter);
    guardedStorageSet("wo_counter", String(newCounter), true).catch(() => {});
    return num;
  };

  // Adds a newly-typed block/vineyard to the shared list so it's available everywhere the
  // Block/Vineyard picker shows up, from here on.
  const addVineyardBlock = (name) => {
    setVineyardBlocks((prev) => {
      if (prev.includes(name)) return prev;
      const updated = [...prev, name].sort((a, b) => a.localeCompare(b));
      guardedStorageSet("vineyard_blocks", JSON.stringify(updated), true).catch(() => {
        setSaveError("Your last change didn't save — check your connection and try again.");
      });
      return updated;
    });
  };

  // Renaming cascades to every Harvest and Fruit Analysis entry using the old name, so nothing
  // is left pointing at a name that no longer exists in the list.
  const renameVineyardBlock = async (oldName, newName) => {
    const trimmed = newName.trim();
    if (!trimmed || trimmed === oldName) return;
    const updatedBlocks = vineyardBlocks.map((b) => (b === oldName ? trimmed : b)).sort((a, b) => a.localeCompare(b));
    setVineyardBlocks(updatedBlocks);
    await persist("vineyard_blocks", updatedBlocks);
    const updatedHarvest = data.harvest.map((h) => (h.block === oldName ? { ...h, block: trimmed } : h));
    const updatedFruitAnalysis = data.fruitAnalysis.map((f) => (f.block === oldName ? { ...f, block: trimmed } : f));
    setData((prev) => ({ ...prev, harvest: updatedHarvest, fruitAnalysis: updatedFruitAnalysis }));
    await persist("harvest", updatedHarvest);
    await persist("fruitAnalysis", updatedFruitAnalysis);
  };

  const deleteVineyardBlock = async (name) => {
    const updated = vineyardBlocks.filter((b) => b !== name);
    setVineyardBlocks(updated);
    await persist("vineyard_blocks", updated);
  };

  const addClone = (name) => {
    setClones((prev) => {
      if (prev.includes(name)) return prev;
      const updated = [...prev, name].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
      guardedStorageSet("clones", JSON.stringify(updated), true).catch(() => {
        setSaveError("Your last change didn't save — check your connection and try again.");
      });
      return updated;
    });
  };

  const renameClone = async (oldName, newName) => {
    const trimmed = newName.trim();
    if (!trimmed || trimmed === oldName) return;
    const updatedClones = clones.map((c) => (c === oldName ? trimmed : c)).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    setClones(updatedClones);
    await persist("clones", updatedClones);
    const updatedHarvest = data.harvest.map((h) => (h.clone === oldName ? { ...h, clone: trimmed } : h));
    setData((prev) => ({ ...prev, harvest: updatedHarvest }));
    await persist("harvest", updatedHarvest);
  };

  const deleteClone = async (name) => {
    const updated = clones.filter((c) => c !== name);
    setClones(updated);
    await persist("clones", updated);
  };

  const addSprayProgram = (name) => {
    setSprayPrograms((prev) => {
      if (prev.includes(name)) return prev;
      const updated = [...prev, name].sort((a, b) => a.localeCompare(b));
      guardedStorageSet("spray_programs", JSON.stringify(updated), true).catch(() => {
        setSaveError("Your last change didn't save — check your connection and try again.");
      });
      return updated;
    });
  };

  const renameSprayProgram = async (oldName, newName) => {
    const trimmed = newName.trim();
    if (!trimmed || trimmed === oldName) return;
    const updated = sprayPrograms.map((s) => (s === oldName ? trimmed : s)).sort((a, b) => a.localeCompare(b));
    setSprayPrograms(updated);
    await persist("spray_programs", updated);
    const updatedOrders = data.workorders.map((o) => (o.sprayProgram === oldName ? { ...o, sprayProgram: trimmed } : o));
    setData((prev) => ({ ...prev, workorders: updatedOrders }));
    await persist("workorders", updatedOrders);
  };

  const deleteSprayProgram = async (name) => {
    const updated = sprayPrograms.filter((s) => s !== name);
    setSprayPrograms(updated);
    await persist("spray_programs", updated);
  };

  const addAssociate = (name) => {
    setTastingAssociates((prev) => {
      if (prev.includes(name)) return prev;
      const updated = [...prev, name].sort((a, b) => a.localeCompare(b));
      guardedStorageSet("tasting_associates", JSON.stringify(updated), true).catch(() => {
        setSaveError("Your last change didn't save — check your connection and try again.");
      });
      return updated;
    });
  };

  const renameAssociate = async (oldName, newName) => {
    const trimmed = newName.trim();
    if (!trimmed || trimmed === oldName) return;
    const updated = tastingAssociates.map((a) => (a === oldName ? trimmed : a)).sort((a, b) => a.localeCompare(b));
    setTastingAssociates(updated);
    await persist("tasting_associates", updated);
    const updatedTimesheets = data.thoTimesheets.map((t) => (t.employeeName === oldName ? { ...t, employeeName: trimmed } : t));
    setData((prev) => ({ ...prev, thoTimesheets: updatedTimesheets }));
    await persist("thoTimesheets", updatedTimesheets);
  };

  const deleteAssociate = async (name) => {
    const updated = tastingAssociates.filter((a) => a !== name);
    setTastingAssociates(updated);
    await persist("tasting_associates", updated);
  };

  const updateDefaultTareWeight = (value) => {
    setDefaultTareWeight(value);
    guardedStorageSet("default_tare_weight", value, true).catch(() => {});
  };

  const updateAlloroStory = (value) => {
    setAlloroStory(value);
    guardedStorageSet("alloro_story", value, true).catch(() => {
      setSaveError("Your last change didn't save — check your connection and try again.");
    });
  };

  const updateWineryName = (value) => {
    setWineryName(value);
    guardedStorageSet("winery_name", value, true).catch(() => {
      setSaveError("Your last change didn't save — check your connection and try again.");
    });
  };

  const updateVineyardMapImage = (value) => {
    setVineyardMapImage(value);
    guardedStorageSet("vineyard_map_image", value, true).catch(() => {
      setSaveError("Your last change didn't save — check your connection and try again.");
    });
  };

  const addVesselType = (name) => {
    setVesselTypes((prev) => {
      if (prev.includes(name)) return prev;
      const updated = [...prev, name].sort((a, b) => a.localeCompare(b));
      guardedStorageSet("vessel_types", JSON.stringify(updated), true).catch(() => {
        setSaveError("Your last change didn't save — check your connection and try again.");
      });
      return updated;
    });
  };

  const renameVesselType = async (oldName, newName) => {
    const trimmed = newName.trim();
    if (!trimmed || trimmed === oldName) return;
    const updatedTypes = vesselTypes.map((v) => (v === oldName ? trimmed : v)).sort((a, b) => a.localeCompare(b));
    setVesselTypes(updatedTypes);
    await persist("vessel_types", updatedTypes);
    const updatedFerment = data.ferment.map((l) => (l.vessel === oldName ? { ...l, vessel: trimmed } : l));
    setData((prev) => ({ ...prev, ferment: updatedFerment }));
    await persist("ferment", updatedFerment);
  };

  const deleteVesselType = async (name) => {
    const updated = vesselTypes.filter((v) => v !== name);
    setVesselTypes(updated);
    await persist("vessel_types", updated);
  };

  // Registers a lot name in the shared master list — called both from the "Manage Lists" panel
  // and automatically whenever someone types a brand-new lot name in a Work Order or barrel fill,
  // so the list grows from real usage without needing separate upkeep.
  const addLotName = (name) => {
    setLotNames((prev) => {
      if (prev.includes(name)) return prev;
      const updated = [...prev, name].sort((a, b) => a.localeCompare(b));
      guardedStorageSet("lot_names", JSON.stringify(updated), true).catch(() => {
        setSaveError("Your last change didn't save — check your connection and try again.");
      });
      return updated;
    });
  };

  // Renaming cascades to Work Orders' linked lots, barrel fill components, and a matching
  // Fermentation lot's Tank/Lot ID, so a rename doesn't leave stale references behind.
  const renameLotName = async (oldName, newName) => {
    const trimmed = newName.trim();
    if (!trimmed || trimmed === oldName) return;
    const updatedNames = lotNames.map((l) => (l === oldName ? trimmed : l)).sort((a, b) => a.localeCompare(b));
    setLotNames(updatedNames);
    await persist("lot_names", updatedNames);

    const updatedWorkorders = data.workorders.map((o) =>
      Array.isArray(o.lots) && o.lots.includes(oldName) ? { ...o, lots: o.lots.map((l) => (l === oldName ? trimmed : l)) } : o
    );
    const updatedBarrels = data.barrels.map((b) => ({
      ...b,
      fills: b.fills.map((f) =>
        Array.isArray(f.components)
          ? { ...f, components: f.components.map((c) => (c.lotLabel === oldName ? { ...c, lotLabel: trimmed } : c)) }
          : f
      ),
    }));
    const updatedFerment = data.ferment.map((l) => (l.tankId === oldName ? { ...l, tankId: trimmed } : l));

    setData((prev) => ({ ...prev, workorders: updatedWorkorders, barrels: updatedBarrels, ferment: updatedFerment }));
    await persist("workorders", updatedWorkorders);
    await persist("barrels", updatedBarrels);
    await persist("ferment", updatedFerment);
  };

  const deleteLotName = async (name) => {
    const updated = lotNames.filter((l) => l !== name);
    setLotNames(updated);
    await persist("lot_names", updated);
  };

  // Updates the SO2 calculator inputs and, once all three are filled in, auto-fills the
  // Directions field with Tom's mL-of-10%-solution calculation.
  const updateSO2Calc = (partial) => {
    const next = { ...so2Calc, ...partial };
    setSo2Calc(next);
    const result = calcSO2StockML(next.volume, next.current, next.target);
    if (result) {
      const { liters, mL, vol, cur, tgt } = result;
      const directionsText = `Add ${mL.toFixed(1)} mL of 10% SO2 solution to raise free SO2 from ${cur} mg/L to ${tgt} mg/L in ${vol} gal (${liters.toFixed(0)} L).`;
      const calculationsText = `(${tgt}-${cur} mg/L) × ${liters.toFixed(0)}L × (1g/1000mg) × (1 KMBS/0.576 SO2) × (1000mL/172g KMBS) = ${mL.toFixed(1)} mL`;
      setNewWorkOrderForm((p) => ({ ...p, directions: directionsText, calculations: calculationsText }));
    }
  };

  const updateAcidCalc = (partial) => {
    const next = { ...acidCalc, ...partial };
    setAcidCalc(next);
    const vol = parseFloat(next.volume), cur = parseFloat(next.current), tgt = parseFloat(next.target);
    if (!isNaN(vol) && !isNaN(cur) && !isNaN(tgt) && vol > 0 && tgt > cur) {
      const liters = vol * 3.78541;
      const grams = (tgt - cur) * liters;
      const directionsText = `Add ${grams.toFixed(1)} g of tartaric acid to raise TA from ${cur} g/L to ${tgt} g/L in ${vol} gal (${liters.toFixed(0)} L).`;
      const calculationsText = `(${tgt}-${cur} g/L) × ${liters.toFixed(0)}L = ${grams.toFixed(1)} g tartaric acid`;
      setNewWorkOrderForm((p) => ({ ...p, directions: directionsText, calculations: calculationsText }));
    }
  };

  const updateWaterCalc = (partial) => {
    const next = { ...waterCalc, ...partial };
    setWaterCalc(next);
    const vol = parseFloat(next.volume), cur = parseFloat(next.current), tgt = parseFloat(next.target);
    if (!isNaN(vol) && !isNaN(cur) && !isNaN(tgt) && vol > 0 && tgt > 0 && tgt < cur) {
      const newVolume = vol * (cur / tgt);
      const water = newVolume - vol;
      const directionsText = `Add ${water.toFixed(1)} gal of water to bring Brix from ${cur}° to ${tgt}° in ${vol} gal (new total volume ${newVolume.toFixed(1)} gal).`;
      const calculationsText = `New Volume = ${vol}gal × (${cur}/${tgt}) = ${newVolume.toFixed(1)}gal; Water to Add = ${newVolume.toFixed(1)} - ${vol} = ${water.toFixed(1)} gal`;
      setNewWorkOrderForm((p) => ({ ...p, directions: directionsText, calculations: calculationsText }));
    }
  };

  const addWorkOrder = async (e) => {
    e.preventDefault();
    if (!newWorkOrderForm.task.trim()) {
      setError('Please fill in "Task"');
      return;
    }
    setError("");
    const order = { id: genId(), status: "Open", dateCompleted: "", ...newWorkOrderForm, orderNumber: nextOrderNumber() };
    const updated = [order, ...data.workorders];
    setSaving(true);
    await persist("workorders", updated);
    setData((prev) => ({ ...prev, workorders: updated }));
    setNewWorkOrderForm({ ...emptyForm(WORKORDER_FIELDS), category: workOrderFilterCategory, date: todayISO(), dateAssigned: todayISO() });
    setSo2Calc({ volume: "", current: "", target: "" });
    setAcidCalc({ volume: "", current: "", target: "" });
    setWaterCalc({ volume: "", current: "", target: "" });
    setSaving(false);
  };

  const toggleWorkOrder = async (id) => {
    const order = data.workorders.find((o) => o.id === id);
    if (!order) return;
    const willComplete = order.status !== "Complete";
    const dateCompleted = willComplete ? todayISO() : "";

    const updated = data.workorders.map((o) =>
      o.id === id
        ? o.status === "Complete"
          ? { ...o, status: "Open", dateCompleted: "" }
          : { ...o, status: "Complete", dateCompleted: todayISO() }
        : o
    );
    setData((prev) => ({ ...prev, workorders: updated }));
    await persist("workorders", updated);

    // If this work order references specific barrels, log the completion on each one
    // (or retract it if the order is being reopened).
    if (Array.isArray(order.barrels) && order.barrels.length > 0) {
      const barrelIdSet = new Set(order.barrels);
      const updatedBarrels = data.barrels.map((b) => {
        if (!barrelIdSet.has(b.id)) return b;
        if (willComplete) {
          const entry = {
            id: genId(),
            workOrderId: order.id,
            workOrderNumber: order.orderNumber,
            task: order.task,
            taskType: order.taskType,
            additionType: order.additionType,
            dateCompleted,
          };
          return { ...b, workHistory: [entry, ...(b.workHistory || [])] };
        }
        return { ...b, workHistory: (b.workHistory || []).filter((h) => h.workOrderId !== order.id) };
      });
      setData((prev) => ({ ...prev, barrels: updatedBarrels }));
      await persist("barrels", updatedBarrels);
    }

    // If this work order is an Addition, Pump Over, or Punch Down linked to specific lots,
    // auto-log it into today's Fermentation reading for that session — only on completion,
    // not on reopen, so a manually-edited reading is never silently reversed.
    if (willComplete && Array.isArray(order.lots) && order.lots.length > 0 && ["Additions", "Pump Over", "Punch Down"].includes(order.taskType)) {
      const session = guessSession();
      const today = todayISO();
      const lotNameSet = new Set(order.lots.map((l) => l.toLowerCase().trim()));
      const updatedFerment = data.ferment.map((lot) => {
        if (!lot.tankId || !lotNameSet.has(lot.tankId.toLowerCase().trim())) return lot;
        const existingIdx = lot.readings.findIndex((r) => r.date === today && r.session === session);
        const existing = existingIdx >= 0 ? lot.readings[existingIdx] : null;
        const changes = {};
        if (order.taskType === "Additions" && order.additionType) {
          const current = existing?.additions || [];
          changes.additions = current.includes(order.additionType) ? current : [...current, order.additionType];
        }
        if (order.taskType === "Pump Over" || order.taskType === "Punch Down") {
          const current = existing?.workDone || [];
          changes.workDone = current.includes(order.taskType) ? current : [...current, order.taskType];
        }
        const note = `Auto-logged from ${formatOrderNumber(order.orderNumber)}: ${order.task || order.taskType}`;
        const existingNotes = existing?.notes || "";
        changes.notes = existingNotes.includes(note) ? existingNotes : existingNotes ? `${existingNotes} · ${note}` : note;
        const updatedReadings =
          existingIdx >= 0
            ? lot.readings.map((r, i) => (i === existingIdx ? { ...r, ...changes } : r))
            : [...lot.readings, { id: genId(), date: today, session, brix: "", temp: "", ph: "", workDone: [], additions: [], notes: "", ...changes }];
        return { ...lot, readings: updatedReadings };
      });
      setData((prev) => ({ ...prev, ferment: updatedFerment }));
      await persist("ferment", updatedFerment);
    }

    // If this is a Fermentation-stage work order linked to specific lots, advance those lots'
    // stage on the Fermentation Overview map — only on completion, matching the same pattern as
    // the cascade above, so reopening a work order never silently reverses a lot's progress.
    if (willComplete && Array.isArray(order.lots) && order.lots.length > 0 && order.taskType === "Fermentation" && order.fermentStage) {
      const today = todayISO();
      const lotNameSet = new Set(order.lots.map((l) => l.toLowerCase().trim()));
      let updatedFermentStages = data.ferment.map((lot) => {
        if (!lot.tankId || !lotNameSet.has(lot.tankId.toLowerCase().trim())) return lot;
        switch (order.fermentStage) {
          case "Start Primary Fermentation":
            return { ...lot, status: "Active", startDate: lot.startDate || today };
          case "Complete Primary Fermentation":
            return { ...lot, status: "Complete", dateCompleted: today };
          case "Start Malolactic Fermentation":
            return { ...lot, mlStatus: "Inoculated", mlInoculationDate: lot.mlInoculationDate || today };
          case "Complete Malolactic Fermentation":
            return { ...lot, mlStatus: "Complete", mlCompleteDate: today };
          default:
            return lot;
        }
      });

      // "Start Primary Fermentation" doesn't require the lot to already exist — any linked lot
      // name with no matching Fermentation record gets a brand-new one created here, the same
      // minimal shape the manual "Start a New Ferment" form produces (just a name and a start
      // date are required; vessel, variety, and everything else can be filled in later).
      if (order.fermentStage === "Start Primary Fermentation") {
        const existingNames = new Set(data.ferment.map((l) => (l.tankId || "").toLowerCase().trim()));
        const newLots = order.lots
          .filter((name) => name.trim() && !existingNames.has(name.toLowerCase().trim()))
          .map((name) => ({
            id: genId(),
            status: "Active",
            readings: [],
            tankId: name.trim(),
            vessel: "",
            variety: "",
            vintage: "",
            wineStyle: "",
            startDate: today,
            startingBrix: "",
            startingTemp: "",
            notes: `Started from Work Order ${formatOrderNumber(order.orderNumber)}`,
          }));
        updatedFermentStages = [...newLots, ...updatedFermentStages];
      }

      setData((prev) => ({ ...prev, ferment: updatedFermentStages }));
      await persist("ferment", updatedFermentStages);
    }
  };

  const deleteWorkOrder = (id) => {
    confirmAction("Delete this work order? This can't be undone.", async () => {
      const order = data.workorders.find((o) => o.id === id);
      const updated = data.workorders.filter((o) => o.id !== id);
      setData((prev) => ({ ...prev, workorders: updated }));
      await persist("workorders", updated);

      if (order && Array.isArray(order.barrels) && order.barrels.length > 0) {
        const barrelIdSet = new Set(order.barrels);
        const updatedBarrels = data.barrels.map((b) =>
          barrelIdSet.has(b.id) ? { ...b, workHistory: (b.workHistory || []).filter((h) => h.workOrderId !== id) } : b
        );
        setData((prev) => ({ ...prev, barrels: updatedBarrels }));
        await persist("barrels", updatedBarrels);
      }
    });
  };

  const startEditWorkOrder = (order) => {
    setEditingWorkOrderId(order.id);
    setEditWorkOrderForm({ ...order });
  };
  const cancelEditWorkOrder = () => {
    setEditingWorkOrderId(null);
    setEditWorkOrderForm(null);
  };
  const editWorkOrderChange = (name, value) => setEditWorkOrderForm((prev) => ({ ...prev, [name]: value }));
  const saveEditWorkOrder = async () => {
    if (!editWorkOrderForm.task || !editWorkOrderForm.task.trim()) {
      setError('Please fill in "Task"');
      return;
    }
    setError("");
    const updated = data.workorders.map((o) => (o.id === editingWorkOrderId ? { ...o, ...editWorkOrderForm } : o));
    setData((prev) => ({ ...prev, workorders: updated }));
    await persist("workorders", updated);
    setEditingWorkOrderId(null);
    setEditWorkOrderForm(null);
  };

  // Duplicates a work order (e.g. for the same task next vintage) and drops it straight into
  // edit mode so the due date — and anything else — can be adjusted right away.
  const duplicateWorkOrder = async (order) => {
    const newOrder = {
      ...order,
      id: genId(),
      status: "Open",
      dateCompleted: "",
      dateAssigned: todayISO(),
      orderNumber: nextOrderNumber(),
    };
    const updated = [newOrder, ...data.workorders];
    setData((prev) => ({ ...prev, workorders: updated }));
    await persist("workorders", updated);
    setEditingWorkOrderId(newOrder.id);
    setEditWorkOrderForm({ ...newOrder });
  };

  // ---- Work Order Templates ----
  const saveAsTemplate = async (order) => {
    const template = { id: genId() };
    TEMPLATE_FIELDS.forEach((f) => {
      template[f.name] = order[f.name] ?? "";
    });
    const updated = [template, ...data.templates];
    setData((prev) => ({ ...prev, templates: updated }));
    await persist("templates", updated);
  };

  const useTemplate = async (template) => {
    const newOrder = { id: genId(), status: "Open", dateCompleted: "", date: todayISO(), dateAssigned: todayISO(), orderNumber: nextOrderNumber() };
    TEMPLATE_FIELDS.forEach((f) => {
      newOrder[f.name] = template[f.name] ?? "";
    });
    const updated = [newOrder, ...data.workorders];
    setData((prev) => ({ ...prev, workorders: updated }));
    await persist("workorders", updated);
    setEditingWorkOrderId(newOrder.id);
    setEditWorkOrderForm({ ...newOrder });
  };

  const deleteTemplate = (id) => {
    confirmAction("Delete this template? This can't be undone.", async () => {
      const updated = data.templates.filter((t) => t.id !== id);
      setData((prev) => ({ ...prev, templates: updated }));
      await persist("templates", updated);
    });
  };

  // ---- Barrel Tastings ----
  const addTasting = async (tasting) => {
    const entry = { id: genId(), ...tasting };
    const updated = [entry, ...data.tastings];
    setData((prev) => ({ ...prev, tastings: updated }));
    await persist("tastings", updated);
  };

  // Logs (or updates, if already logged) a snapshot of a given day's weather + GDD onto the
  // calendar. Re-logging the same date overwrites that day's entry instead of duplicating it.
  const logWeatherSnapshot = async (snapshot) => {
    const existingIdx = data.weatherLogs.findIndex((w) => w.date === snapshot.date);
    const updated =
      existingIdx >= 0
        ? data.weatherLogs.map((w, i) => (i === existingIdx ? { ...w, ...snapshot } : w))
        : [{ id: genId(), ...snapshot }, ...data.weatherLogs];
    setData((prev) => ({ ...prev, weatherLogs: updated }));
    await persist("weatherLogs", updated);
  };

  const deleteWeatherLog = (id) => {
    confirmAction("Delete this weather log entry? This can't be undone.", async () => {
      const updated = data.weatherLogs.filter((w) => w.id !== id);
      setData((prev) => ({ ...prev, weatherLogs: updated }));
      await persist("weatherLogs", updated);
    });
  };

  const saveTechSheet = async (sheet) => {
    const exists = data.techSheets.some((s) => s.id === sheet.id);
    const updated = exists ? data.techSheets.map((s) => (s.id === sheet.id ? sheet : s)) : [sheet, ...data.techSheets];
    setData((prev) => ({ ...prev, techSheets: updated }));
    await persist("techSheets", updated);
  };

  const deleteTechSheet = (id) => {
    confirmAction("Delete this tech sheet? This can't be undone.", async () => {
      const updated = data.techSheets.filter((s) => s.id !== id);
      setData((prev) => ({ ...prev, techSheets: updated }));
      await persist("techSheets", updated);
    });
  };

  const updateTasting = async (id, changes) => {
    const updated = data.tastings.map((t) => (t.id === id ? { ...t, ...changes } : t));
    setData((prev) => ({ ...prev, tastings: updated }));
    await persist("tastings", updated);
  };

  const deleteTasting = (id) => {
    confirmAction("Delete this tasting note? This can't be undone.", async () => {
      const updated = data.tastings.filter((t) => t.id !== id);
      setData((prev) => ({ ...prev, tastings: updated }));
      await persist("tastings", updated);
    });
  };

  // ---- Blends ----
  const addBlend = async (blend) => {
    const entry = { id: genId(), dateCreated: todayISO(), ...blend };
    const updated = [entry, ...data.blends];
    setData((prev) => ({ ...prev, blends: updated }));
    await persist("blends", updated);
  };

  const updateBlend = async (id, changes) => {
    const updated = data.blends.map((b) => (b.id === id ? { ...b, ...changes } : b));
    setData((prev) => ({ ...prev, blends: updated }));
    await persist("blends", updated);
  };

  const deleteBlend = (id) => {
    confirmAction("Delete this blend trial? This can't be undone.", async () => {
      const updated = data.blends.filter((b) => b.id !== id);
      setData((prev) => ({ ...prev, blends: updated }));
      await persist("blends", updated);
    });
  };

  // ---- Fermentation-specific handlers ----
  const addFermentLot = async (e) => {
    e.preventDefault();
    if (!newLotForm.tankId.trim()) {
      setError('Please fill in "Tank / Lot ID"');
      return;
    }
    setError("");
    const lot = { id: genId(), status: "Active", readings: [], ...newLotForm };
    const updated = [lot, ...data.ferment];
    setSaving(true);
    await persist("ferment", updated);
    setData((prev) => ({ ...prev, ferment: updated }));
    setNewLotForm({ ...emptyForm(FERMENT_LOT_FIELDS), startDate: todayISO() });
    setSaving(false);
  };

  const addReadingToLot = async (lotId, reading) => {
    const updated = data.ferment.map((lot) =>
      lot.id === lotId ? { ...lot, readings: [...lot.readings, reading] } : lot
    );
    setData((prev) => ({ ...prev, ferment: updated }));
    await persist("ferment", updated);
  };

  // Finds or creates today's (or the selected date's) reading for a given session and merges
  // in the Quick Log grid's values — used for both manual entry and Work Order auto-logging.
  const upsertLotReading = (lot, date, session, values) => {
    const existingIdx = lot.readings.findIndex((r) => r.date === date && r.session === session);
    if (existingIdx >= 0) {
      return lot.readings.map((r, i) => (i === existingIdx ? { ...r, ...values } : r));
    }
    return [...lot.readings, { id: genId(), date, session, brix: "", temp: "", ph: "", workDone: [], additions: [], notes: "", ...values }];
  };

  const logQuickReading = async (lot, date, session, values) => {
    const updated = data.ferment.map((l) => (l.id === lot.id ? { ...l, readings: upsertLotReading(l, date, session, values) } : l));
    setData((prev) => ({ ...prev, ferment: updated }));
    await persist("ferment", updated);
  };

  const toggleLotComplete = async (lotId) => {
    const updated = data.ferment.map((lot) =>
      lot.id === lotId
        ? lot.status === "Active"
          ? { ...lot, status: "Complete", dateCompleted: todayISO() }
          : { ...lot, status: "Active", dateCompleted: "" }
        : lot
    );
    setData((prev) => ({ ...prev, ferment: updated }));
    await persist("ferment", updated);
  };

  const deleteLot = async (lotId) => {
    const updated = data.ferment.filter((lot) => lot.id !== lotId);
    setData((prev) => ({ ...prev, ferment: updated }));
    await persist("ferment", updated);
  };

  const updateFermentLot = async (lotId, changes) => {
    const updated = data.ferment.map((lot) => (lot.id === lotId ? { ...lot, ...changes } : lot));
    setData((prev) => ({ ...prev, ferment: updated }));
    await persist("ferment", updated);
  };

  const updateReading = async (lotId, readingId, changes) => {
    const updated = data.ferment.map((lot) =>
      lot.id === lotId ? { ...lot, readings: lot.readings.map((r) => (r.id === readingId ? { ...r, ...changes } : r)) } : lot
    );
    setData((prev) => ({ ...prev, ferment: updated }));
    await persist("ferment", updated);
  };

  const deleteReading = async (lotId, readingId) => {
    const updated = data.ferment.map((lot) =>
      lot.id === lotId ? { ...lot, readings: lot.readings.filter((r) => r.id !== readingId) } : lot
    );
    setData((prev) => ({ ...prev, ferment: updated }));
    await persist("ferment", updated);
  };

  // ---- Barrel handlers ----
  const addBarrel = async (barrel) => {
    const updated = [barrel, ...data.barrels];
    setData((prev) => ({ ...prev, barrels: updated }));
    await persist("barrels", updated);
  };

  const bulkAddBarrels = async (barrels) => {
    const updated = [...barrels, ...data.barrels];
    setData((prev) => ({ ...prev, barrels: updated }));
    await persist("barrels", updated);
  };

  const fillBarrel = async (barrelId, fillRecord) => {
    const updated = data.barrels.map((b) => (b.id === barrelId ? { ...b, fills: [...b.fills, fillRecord] } : b));
    setData((prev) => ({ ...prev, barrels: updated }));
    await persist("barrels", updated);
  };

  const emptyBarrel = async (barrelId, fillId) => {
    const updated = data.barrels.map((b) =>
      b.id === barrelId
        ? { ...b, fills: b.fills.map((f) => (f.id === fillId ? { ...f, emptyDate: todayISO() } : f)) }
        : b
    );
    setData((prev) => ({ ...prev, barrels: updated }));
    await persist("barrels", updated);
  };

  const updateBarrel = async (barrelId, changes) => {
    const updated = data.barrels.map((b) => (b.id === barrelId ? { ...b, ...changes } : b));
    setData((prev) => ({ ...prev, barrels: updated }));
    await persist("barrels", updated);
  };

  const deleteBarrel = async (barrelId) => {
    const updated = data.barrels.filter((b) => b.id !== barrelId);
    setData((prev) => ({ ...prev, barrels: updated }));
    await persist("barrels", updated);
  };

  // Empties every selected barrel that's currently In Use; other selected barrels are left alone.
  const bulkEmptyBarrels = async (barrelIds) => {
    const idSet = new Set(barrelIds);
    const updated = data.barrels.map((b) => {
      if (!idSet.has(b.id) || barrelStatus(b) !== "In Use") return b;
      const active = activeBarrelFill(b);
      if (!active) return b;
      return { ...b, fills: b.fills.map((f) => (f.id === active.id ? { ...f, emptyDate: todayISO() } : f)) };
    });
    setData((prev) => ({ ...prev, barrels: updated }));
    await persist("barrels", updated);
  };

  // Fills every selected barrel that's currently Empty with the same lot/date; skips the rest.
  const bulkFillBarrels = async (barrelIds, components, fillDate) => {
    const idSet = new Set(barrelIds);
    const updated = data.barrels.map((b) => {
      if (!idSet.has(b.id) || barrelStatus(b) !== "Empty") return b;
      return { ...b, fills: [...b.fills, { id: genId(), fillDate, emptyDate: "", components: components.map((c) => ({ ...c, id: genId() })) }] };
    });
    setData((prev) => ({ ...prev, barrels: updated }));
    await persist("barrels", updated);
  };

  // Marks every selected barrel that isn't already Sold as sold, with the same sale details.
  const bulkSellBarrels = async (barrelIds, sellDetails) => {
    const idSet = new Set(barrelIds);
    const updated = data.barrels.map((b) => {
      if (!idSet.has(b.id) || barrelStatus(b) === "Sold") return b;
      return { ...b, ...sellDetails };
    });
    setData((prev) => ({ ...prev, barrels: updated }));
    await persist("barrels", updated);
  };

  const exportToExcel = useCallback(() => {
    const wb = XLSX.utils.book_new();
    buildExportSections(getSortedData()).forEach((section) => {
      const rows = section.rows.length > 0 ? section.rows : [Object.fromEntries(section.headers.map((h) => [h, ""]))];
      const ws = XLSX.utils.json_to_sheet(rows, { header: section.headers });
      ws["!cols"] = section.headers.map(() => ({ wch: 16 }));
      XLSX.utils.book_append_sheet(wb, ws, section.title.slice(0, 31));
    });
    XLSX.writeFile(wb, `alloro-winery-data-${todayISO()}.xlsx`);
    setExportMenuOpen(false);
  }, [data, workOrderSort, tableSort, fermentSort]);

  const exportToCSV = useCallback(() => {
    let csvText = "";
    buildExportSections(getSortedData()).forEach((section) => {
      csvText += `## ${section.title}\n`;
      csvText += Papa.unparse({
        fields: section.headers,
        data: section.rows.map((row) => section.headers.map((h) => row[h] ?? "")),
      });
      csvText += "\n\n";
    });
    downloadBlob(new Blob([csvText], { type: "text/csv;charset=utf-8;" }), `alloro-winery-data-${todayISO()}.csv`);
    setExportMenuOpen(false);
  }, [data, workOrderSort, tableSort, fermentSort]);

  const exportToWord = useCallback(() => {
    let bodyHtml = `<h1 style="font-family:Georgia,serif;color:#1D1915;">Alloro Winery Tracker</h1><p style="font-family:Arial,sans-serif;color:#555;">Exported ${todayISO()}</p>`;
    buildExportSections(getSortedData()).forEach((section) => {
      bodyHtml += `<h2 style="font-family:Georgia,serif;color:#065f46;margin-top:24px;">${escapeHtml(section.title)}</h2>`;
      if (section.rows.length === 0) {
        bodyHtml += `<p style="font-family:Arial,sans-serif;color:#888;"><em>No entries.</em></p>`;
      } else {
        bodyHtml += `<table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:12px;">`;
        bodyHtml += `<tr>${section.headers.map((h) => `<th style="background:#f0ece4;text-align:left;">${escapeHtml(h)}</th>`).join("")}</tr>`;
        section.rows.forEach((row) => {
          bodyHtml += `<tr>${section.headers.map((h) => `<td>${escapeHtml(row[h])}</td>`).join("")}</tr>`;
        });
        bodyHtml += `</table>`;
      }
    });
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"></head><body>${bodyHtml}</body></html>`;
    downloadBlob(new Blob(["\ufeff", html], { type: "application/msword" }), `alloro-winery-data-${todayISO()}.doc`);
    setExportMenuOpen(false);
  }, [data, workOrderSort, tableSort, fermentSort]);

  const totalEntries =
    data.workorders.length +
    SIMPLE_SECTIONS.reduce((sum, s) => sum + (data[s.key]?.length ?? 0), 0) +
    data.ferment.reduce((sum, lot) => sum + lot.readings.length, 0);

  const sortedFermentForDisplay = sortRows(data.ferment, fermentSort.field, fermentSort.direction);
  const activeLots = sortedFermentForDisplay.filter((l) => l.status === "Active");
  const completeLots = sortedFermentForDisplay.filter((l) => l.status === "Complete" && !l.archived);
  const archivedLots = sortedFermentForDisplay.filter((l) => l.status === "Complete" && l.archived);

  const sortedWorkOrdersForDisplay = sortRows(data.workorders, workOrderSort.field, workOrderSort.direction);
  const openOrders = sortedWorkOrdersForDisplay.filter((o) => o.status !== "Complete");
  const completedOrders = sortedWorkOrdersForDisplay.filter((o) => o.status === "Complete");
  // Orders created before this feature existed have no category — treat them as Winery, since
  // every task type available before now was winery-oriented.
  const categoryFilteredOpenOrders = openOrders.filter((o) => (o.category || "Winery") === workOrderFilterCategory);
  const categoryFilteredCompletedOrders = completedOrders.filter((o) => (o.category || "Winery") === workOrderFilterCategory);

  if (!role) {
    return (
      <PasswordGate
        wineryName={wineryName}
        onUnlock={(newRole) => {
          setRole(newRole);
          writeStoredRole(newRole);
        }}
      />
    );
  }

  return (
    <div className="min-h-screen bg-ink-50 text-ink-900">
      {saveError && (
        <div className="bg-red-700 text-white px-4 py-2 flex items-center justify-between gap-3 text-sm font-body">
          <span>⚠️ {saveError}</span>
          <button onClick={() => setSaveError("")} className="shrink-0 font-medium underline hover:no-underline">
            Dismiss
          </button>
        </div>
      )}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@500;600;700&family=Inter:wght@400;500;600&display=swap');
        .font-brand { font-family: 'Cormorant Garamond', serif; font-weight: 600; letter-spacing: 0.01em; }
        .font-body { font-family: 'Inter', sans-serif; }
        .print-sheet { display: none; }
        @media print {
          @page { margin: 0.6in; }
          body * { visibility: hidden; }
          .print-sheet, .print-sheet * { visibility: visible; }
          .print-sheet { display: block; position: absolute; top: 0; left: 0; width: 100%; padding: 24px; }
        }
      `}</style>

      <header className="bg-ink-950 text-stone-50 px-4 py-3 sm:px-6">
        <div className="max-w-5xl mx-auto flex items-center justify-between gap-3 flex-wrap">
          <div className="flex items-center gap-3">
            <img src={GDS_LOGO} alt="GrapeDaddy" className="w-14 h-14 sm:w-16 sm:h-16 shrink-0" />
            <div>
              <h1 className="font-brand text-xl sm:text-2xl tracking-tight leading-tight">GrapeDaddy</h1>
              <p className="font-body text-ink-300 text-xs tracking-wide mt-0.5">Winery Solutions</p>
            </div>
          </div>

          <div className="relative">
            <button
              onClick={() => setExportMenuOpen((v) => !v)}
              disabled={loading}
              className="font-body flex items-center gap-2 bg-amber-500 hover:bg-amber-400 disabled:opacity-50 text-ink-950 font-semibold text-sm px-4 py-2 rounded-md transition-colors"
            >
              <Download size={16} /> Export Data <ChevronDown size={14} />
            </button>
            {exportMenuOpen && (
              <>
                <div className="fixed inset-0 z-10" onClick={() => setExportMenuOpen(false)} />
                <div className="absolute right-0 mt-2 w-48 bg-white border border-stone-200 rounded-md shadow-lg overflow-hidden z-20">
                  <button
                    onClick={exportToExcel}
                    className="font-body w-full text-left px-4 py-2.5 text-sm text-stone-700 hover:bg-stone-50"
                  >
                    Excel (.xlsx)
                  </button>
                  <button
                    onClick={exportToWord}
                    className="font-body w-full text-left px-4 py-2.5 text-sm text-stone-700 hover:bg-stone-50 border-t border-stone-100"
                  >
                    Word (.doc)
                  </button>
                  <button
                    onClick={exportToCSV}
                    className="font-body w-full text-left px-4 py-2.5 text-sm text-stone-700 hover:bg-stone-50 border-t border-stone-100"
                  >
                    CSV (.csv)
                  </button>
                  <button
                    onClick={() => {
                      setPrintJob({ type: "all" });
                      setExportMenuOpen(false);
                    }}
                    className="font-body w-full text-left px-4 py-2.5 text-sm text-stone-700 hover:bg-stone-50 border-t border-stone-100"
                  >
                    PDF (.pdf)
                    <span className="block text-xs text-stone-400">opens print dialog — choose "Save as PDF"</span>
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      </header>

      <nav className="bg-ink-900 px-2 sm:px-6">
        <div className="max-w-5xl mx-auto">
          <div className="flex items-center gap-1 overflow-x-auto">
            {PERSISTENT_NAV_KEYS.filter((key) => role === "admin" || key !== "workorders").map((key) => {
              const s = ALL_TABS.find((t) => t.key === key);
              if (!s) return null;
              const Icon = s.icon;
              const isActive = s.key === activeKey;
              const count = s.key === "workorders" ? openOrders.length : null;
              return (
                <button
                  key={s.key}
                  onClick={() => setActiveKey(s.key)}
                  className={`font-body flex items-center gap-1.5 whitespace-nowrap text-sm px-3 py-2.5 border-b-2 transition-colors ${
                    isActive ? "border-amber-400 text-white" : "border-transparent text-ink-200 hover:text-white"
                  }`}
                >
                  <Icon size={16} />
                  {s.label}
                  {count !== null && (
                    <span className={`ml-1 text-xs rounded-full px-1.5 ${isActive ? "bg-amber-400 text-ink-950" : "bg-ink-800 text-ink-200"}`}>
                      {count}
                    </span>
                  )}
                </button>
              );
            })}

            <div className="w-px h-5 bg-ink-700 mx-1 shrink-0" />

            {Object.entries(NAV_CATEGORIES).map(([catKey, cat]) => {
              const isActiveCategory = categoryOfKey(activeKey) === catKey;
              return (
                <button
                  key={catKey}
                  onClick={() => goToCategory(catKey)}
                  className={`font-body flex items-center gap-1.5 whitespace-nowrap text-sm px-3 py-2.5 border-b-2 transition-colors ${
                    isActiveCategory ? "border-amber-400 text-white" : "border-transparent text-ink-200 hover:text-white"
                  }`}
                >
                  <span className={`w-2 h-2 rounded-full shrink-0 ${cat.dotColor}`} />
                  {cat.label}
                </button>
              );
            })}

            <button
              onClick={() => {
                setRole(null);
                writeStoredRole(null);
              }}
              className="font-body flex items-center gap-1.5 whitespace-nowrap text-xs px-3 py-3 text-ink-300 hover:text-white ml-auto"
              title="Log out"
            >
              <LogOut size={14} />
            </button>
          </div>

          {categoryOfKey(activeKey) &&
            (() => {
              const catKey = categoryOfKey(activeKey);
              const cat = NAV_CATEGORIES[catKey];
              const visibleKeys = cat.keys.filter((key) => role === "admin" || key !== "backup");
              return (
                <div className="flex items-center gap-1 overflow-x-auto border-t border-ink-800 -mt-px">
                  {visibleKeys.map((key) => {
                    const s = ALL_TABS.find((t) => t.key === key);
                    if (!s) return null;
                    const Icon = s.icon;
                    const isActive = s.key === activeKey;
                    const count = s.key === "ferment" ? data.ferment.length : data[s.key]?.length ?? null;
                    const showCount = !["calendar", "formulas", "blending", "backup"].includes(s.key);
                    return (
                      <button
                        key={s.key}
                        onClick={() => setActiveKey(s.key)}
                        className={`font-body flex items-center gap-1.5 whitespace-nowrap text-xs sm:text-sm px-3 py-1.5 border-b-2 transition-colors ${
                          isActive ? "border-amber-400 text-white" : "border-transparent text-ink-300 hover:text-white"
                        }`}
                      >
                        <Icon size={14} />
                        {s.label}
                        {showCount && count !== null && (
                          <span className={`ml-0.5 text-[10px] rounded-full px-1.5 ${isActive ? "bg-amber-400 text-ink-950" : "bg-ink-800 text-ink-300"}`}>
                            {count}
                          </span>
                        )}
                      </button>
                    );
                  })}
                </div>
              );
            })()}
        </div>
      </nav>

      <main className="max-w-5xl mx-auto px-4 py-6 sm:px-6">
        {role === "tho" && !THO_EDITABLE_NAV_KEYS.includes(activeKey) && (
          <div className="bg-rose-50 border border-rose-200 rounded-lg px-4 py-2.5 mb-4 flex items-center gap-2">
            <Eye size={15} className="text-rose-700 shrink-0" />
            <p className="font-body text-sm text-rose-800">View only — you're signed in with tasting house access. Changes here won't save.</p>
          </div>
        )}
        {loading ? (
          <div className="flex items-center gap-2 text-stone-500 font-body text-sm py-12 justify-center">
            <Loader2 size={18} className="animate-spin" /> Loading shared data…
          </div>
        ) : activeKey === "home" ? (
          <HomeTab
            data={data}
            toggleWorkOrder={toggleWorkOrder}
            deleteWorkOrder={deleteWorkOrder}
            editingWorkOrderId={editingWorkOrderId}
            editWorkOrderForm={editWorkOrderForm}
            editWorkOrderChange={editWorkOrderChange}
            startEditWorkOrder={startEditWorkOrder}
            saveEditWorkOrder={saveEditWorkOrder}
            cancelEditWorkOrder={cancelEditWorkOrder}
            duplicateWorkOrder={duplicateWorkOrder}
            saveAsTemplate={saveAsTemplate}
            lotNames={lotNames}
            onRegisterLotName={addLotName}
            onLogWeather={logWeatherSnapshot}
            sprayPrograms={sprayPrograms}
            onAddSprayProgram={addSprayProgram}
          />
        ) : activeKey === "workorders" ? (
          <>
            <div className="flex gap-2 mb-6">
              {WORKORDER_CATEGORIES.map((cat) => (
                <button
                  key={cat}
                  onClick={() => {
                    setWorkOrderFilterCategory(cat);
                    setNewWorkOrderForm((p) => ({ ...p, category: cat, taskType: "" }));
                  }}
                  className={`font-body text-sm font-medium px-4 py-2 rounded-md border ${
                    workOrderFilterCategory === cat ? "bg-ink-900 text-white border-ink-900" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                  }`}
                >
                  {cat}
                </button>
              ))}
            </div>

            <form onSubmit={addWorkOrder} className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5 mb-6">
              <h2 className="font-brand text-lg text-ink-950 mb-3">Add a {workOrderFilterCategory} Work Order</h2>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {visibleWorkOrderFields(workOrderFieldsForCategory(WORKORDER_FIELDS, newWorkOrderForm.category), newWorkOrderForm).map((f) => (
                  <Field
                    key={f.name}
                    f={f}
                    value={newWorkOrderForm[f.name]}
                    onChange={(v) => setNewWorkOrderForm((p) => ({ ...p, [f.name]: v }))}
                    fermentLots={data.ferment}
                    barrelsList={data.barrels}
                    lotNamesList={lotNames}
                    onRegisterLotName={addLotName}
                    sprayProgramsList={sprayPrograms}
                    onAddSprayProgram={addSprayProgram}
                  />
                ))}
                {newWorkOrderForm.taskType === "Additions" && newWorkOrderForm.additionType === "SO2" && (
                  <div className="sm:col-span-2 bg-amber-50 border border-amber-100 rounded-md p-3">
                    <p className="font-body text-xs font-semibold text-stone-700 mb-2">
                      SO₂ Addition Calculator — Tom's Method (10% liquid solution)
                    </p>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                      <div>
                        <label className="font-body block text-xs text-stone-600 mb-1">Volume (gal)</label>
                        <input
                          type="number"
                          value={so2Calc.volume}
                          onChange={(e) => updateSO2Calc({ volume: e.target.value })}
                          className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                        />
                      </div>
                      <div>
                        <label className="font-body block text-xs text-stone-600 mb-1">Current FSO₂ (mg/L)</label>
                        <input
                          type="number"
                          value={so2Calc.current}
                          onChange={(e) => updateSO2Calc({ current: e.target.value })}
                          className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                        />
                      </div>
                      <div>
                        <label className="font-body block text-xs text-stone-600 mb-1">Target FSO₂ (mg/L)</label>
                        <input
                          type="number"
                          value={so2Calc.target}
                          onChange={(e) => updateSO2Calc({ target: e.target.value })}
                          className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                        />
                      </div>
                    </div>
                    <p className="font-body text-xs text-stone-500 mt-2">
                      Directions below fill in automatically once all three are entered.
                    </p>
                  </div>
                )}

                {newWorkOrderForm.taskType === "Additions" && newWorkOrderForm.additionType === "Acid" && (
                  <div className="sm:col-span-2 bg-amber-50 border border-amber-100 rounded-md p-3">
                    <p className="font-body text-xs font-semibold text-stone-700 mb-2">Acid Addition Calculator (Tartaric)</p>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                      <div>
                        <label className="font-body block text-xs text-stone-600 mb-1">Volume (gal)</label>
                        <input
                          type="number"
                          value={acidCalc.volume}
                          onChange={(e) => updateAcidCalc({ volume: e.target.value })}
                          className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                        />
                      </div>
                      <div>
                        <label className="font-body block text-xs text-stone-600 mb-1">Current TA (g/L)</label>
                        <input
                          type="number"
                          value={acidCalc.current}
                          onChange={(e) => updateAcidCalc({ current: e.target.value })}
                          className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                        />
                      </div>
                      <div>
                        <label className="font-body block text-xs text-stone-600 mb-1">Target TA (g/L)</label>
                        <input
                          type="number"
                          value={acidCalc.target}
                          onChange={(e) => updateAcidCalc({ target: e.target.value })}
                          className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                        />
                      </div>
                    </div>
                    <p className="font-body text-xs text-stone-500 mt-2">
                      Directions below fill in automatically once all three are entered.
                    </p>
                  </div>
                )}

                {newWorkOrderForm.taskType === "Additions" && newWorkOrderForm.additionType === "Water" && (
                  <div className="sm:col-span-2 bg-amber-50 border border-amber-100 rounded-md p-3">
                    <p className="font-body text-xs font-semibold text-stone-700 mb-2">Water Addition Calculator (Dilution)</p>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                      <div>
                        <label className="font-body block text-xs text-stone-600 mb-1">Current Volume (gal)</label>
                        <input
                          type="number"
                          value={waterCalc.volume}
                          onChange={(e) => updateWaterCalc({ volume: e.target.value })}
                          className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                        />
                      </div>
                      <div>
                        <label className="font-body block text-xs text-stone-600 mb-1">Current Brix</label>
                        <input
                          type="number"
                          value={waterCalc.current}
                          onChange={(e) => updateWaterCalc({ current: e.target.value })}
                          className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                        />
                      </div>
                      <div>
                        <label className="font-body block text-xs text-stone-600 mb-1">Target Brix</label>
                        <input
                          type="number"
                          value={waterCalc.target}
                          onChange={(e) => updateWaterCalc({ target: e.target.value })}
                          className="font-body w-full border border-stone-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-ink-800"
                        />
                      </div>
                    </div>
                    <p className="font-body text-xs text-stone-500 mt-2">
                      Directions below fill in automatically once all three are entered.
                    </p>
                  </div>
                )}
              </div>
              {error && <p className="font-body text-sm text-red-700 mt-3">{error}</p>}
              <button
                type="submit"
                disabled={saving}
                className="font-body mt-4 flex items-center gap-2 bg-ink-900 hover:bg-ink-800 disabled:opacity-50 text-white text-sm font-medium px-4 py-2 rounded-md transition-colors"
              >
                {saving ? <Loader2 size={16} className="animate-spin" /> : <Plus size={16} />}
                Add Work Order
              </button>
            </form>

            <div className="bg-white border border-stone-200 rounded-lg overflow-hidden mb-6">
              <div className="px-4 py-3 border-b border-stone-200 flex items-center justify-between">
                <div>
                  <h2 className="font-brand text-lg text-ink-950">Work Order Templates</h2>
                  <p className="font-body text-xs text-stone-500 mt-0.5">
                    Save a work order you'll reuse every vintage, then create a fresh one from it in one click.
                  </p>
                </div>
                <span className="font-body text-xs text-stone-500 shrink-0">{data.templates.length} saved</span>
              </div>
              {data.templates.length === 0 ? (
                <p className="font-body text-sm text-stone-500 px-4 py-8 text-center">
                  No templates yet — use the bookmark icon on any work order to save it here.
                </p>
              ) : (
                <ul className="divide-y divide-stone-100">
                  {data.templates.map((t) => (
                    <TemplateRow key={t.id} template={t} onUse={useTemplate} onDelete={deleteTemplate} />
                  ))}
                </ul>
              )}
            </div>

            <div className="bg-white border border-stone-200 rounded-lg overflow-hidden mb-6">
              <div className="px-4 py-3 border-b border-stone-200 flex items-center justify-between flex-wrap gap-2">
                <h2 className="font-brand text-lg text-ink-950">Work Orders</h2>
                <div className="flex items-center gap-2">
                  <span className="font-body text-xs text-stone-500">Sort by</span>
                  <select
                    value={workOrderSort.field}
                    onChange={(e) => setWorkOrderSort((p) => ({ ...p, field: e.target.value }))}
                    className="font-body text-xs border border-stone-300 rounded px-2 py-1.5"
                  >
                    <option value="">Default</option>
                    {WORKORDER_SORT_OPTIONS.map((opt) => (
                      <option key={opt.value} value={opt.value}>{opt.label}</option>
                    ))}
                  </select>
                  <button
                    onClick={() => setWorkOrderSort((p) => ({ ...p, direction: p.direction === "asc" ? "desc" : "asc" }))}
                    className="text-stone-500 hover:text-ink-800 border border-stone-300 rounded p-1.5"
                    title={workOrderSort.direction === "asc" ? "Ascending" : "Descending"}
                  >
                    {workOrderSort.direction === "asc" ? <ArrowUp size={14} /> : <ArrowDown size={14} />}
                  </button>
                  <button
                    onClick={() => setPrintJob({ type: "workorders" })}
                    className="font-body flex items-center gap-1.5 text-sm font-medium text-ink-900 hover:text-ink-700"
                  >
                    <Printer size={16} /> Print
                  </button>
                </div>
              </div>
              {categoryFilteredOpenOrders.length === 0 ? (
                <p className="font-body text-sm text-stone-500 px-4 py-8 text-center">No {workOrderFilterCategory.toLowerCase()} work orders yet — add one above.</p>
              ) : (
                <ul className="divide-y divide-stone-100">
                  {categoryFilteredOpenOrders.map((o) => (
                    <WorkOrderRow
                      key={o.id}
                      order={o}
                      onToggle={toggleWorkOrder}
                      onDelete={deleteWorkOrder}
                      onDuplicate={duplicateWorkOrder}
                      onSaveAsTemplate={saveAsTemplate}
                      isEditing={editingWorkOrderId === o.id}
                      editForm={editWorkOrderForm}
                      onEditChange={editWorkOrderChange}
                      onStartEdit={startEditWorkOrder}
                      onSaveEdit={saveEditWorkOrder}
                      onCancelEdit={cancelEditWorkOrder}
                      fermentLots={data.ferment}
                      barrelsList={data.barrels}
                      lotNamesList={lotNames}
                      onRegisterLotName={addLotName}
                      sprayProgramsList={sprayPrograms}
                      onAddSprayProgram={addSprayProgram}
                    />
                  ))}
                </ul>
              )}
            </div>

            {categoryFilteredCompletedOrders.length > 0 && (
              <div className="bg-white border border-stone-200 rounded-lg overflow-hidden mb-6">
                <div className="px-4 py-3 border-b border-stone-200">
                  <h2 className="font-brand text-lg text-stone-500">Archive ({categoryFilteredCompletedOrders.length} closed)</h2>
                  <p className="font-body text-xs text-stone-400 mt-0.5">Grouped by the month each order was closed.</p>
                </div>
                {groupByMonth(categoryFilteredCompletedOrders).map((group) => (
                  <ArchiveGroup
                    key={group.key}
                    label={group.label}
                    orders={group.orders}
                    onToggle={toggleWorkOrder}
                    onDelete={deleteWorkOrder}
                    onDuplicate={duplicateWorkOrder}
                    onSaveAsTemplate={saveAsTemplate}
                    editingWorkOrderId={editingWorkOrderId}
                    editWorkOrderForm={editWorkOrderForm}
                    onEditChange={editWorkOrderChange}
                    onStartEdit={startEditWorkOrder}
                    onSaveEdit={saveEditWorkOrder}
                    onCancelEdit={cancelEditWorkOrder}
                    fermentLots={data.ferment}
                    barrelsList={data.barrels}
                    lotNamesList={lotNames}
                    onRegisterLotName={addLotName}
                    sprayProgramsList={sprayPrograms}
                    onAddSprayProgram={addSprayProgram}
                  />
                ))}
              </div>
            )}

            <CompletedCalendar orders={completedOrders} />
          </>
        ) : activeKey === "ferment" ? (
          <>
            <form onSubmit={addFermentLot} className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5 mb-6">
              <h2 className="font-brand text-lg text-ink-950 mb-1">Start a New Ferment</h2>
              <p className="font-body text-xs text-stone-500 mb-3">
                Add it once here — then log daily readings on the card below until it's done.
              </p>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {FERMENT_LOT_FIELDS.map((f) => (
                  <Field key={f.name} f={f} value={newLotForm[f.name]} onChange={(v) => setNewLotForm((p) => ({ ...p, [f.name]: v }))} vesselTypesList={vesselTypes} onAddVesselType={addVesselType} />
                ))}
              </div>
              {error && <p className="font-body text-sm text-red-700 mt-3">{error}</p>}
              <button
                type="submit"
                disabled={saving}
                className="font-body mt-4 flex items-center gap-2 bg-ink-900 hover:bg-ink-800 disabled:opacity-50 text-white text-sm font-medium px-4 py-2 rounded-md transition-colors"
              >
                {saving ? <Loader2 size={16} className="animate-spin" /> : <Plus size={16} />}
                Start Ferment
              </button>
            </form>

            <div className="flex gap-2 mb-6 flex-wrap">
              <button
                onClick={() => setFermentViewMode("overview")}
                className={`font-body text-sm font-medium px-4 py-2 rounded-md border ${
                  fermentViewMode === "overview" ? "bg-ink-900 text-white border-ink-900" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                }`}
              >
                Overview
              </button>
              <button
                onClick={() => setFermentViewMode("quick")}
                className={`font-body text-sm font-medium px-4 py-2 rounded-md border ${
                  fermentViewMode === "quick" ? "bg-ink-900 text-white border-ink-900" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                }`}
              >
                Quick Log
              </button>
              <button
                onClick={() => setFermentViewMode("detailed")}
                className={`font-body text-sm font-medium px-4 py-2 rounded-md border ${
                  fermentViewMode === "detailed" ? "bg-ink-900 text-white border-ink-900" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                }`}
              >
                Detailed View
              </button>
            </div>

            {fermentViewMode === "overview" ? (
              <FermentOverview lots={data.ferment} onSelectLot={() => setFermentViewMode("detailed")} />
            ) : fermentViewMode === "quick" ? (
              <QuickFermentLog
                lots={activeLots}
                allLots={data.ferment}
                onSaveRow={logQuickReading}
                onEndDay={(date) => setPrintJob({ type: "fermentDayLog", date })}
                onDeleteLot={deleteLot}
                onSwitchToDetailed={() => setFermentViewMode("detailed")}
                confirmAction={confirmAction}
              />
            ) : (
              <>
            {data.ferment.length > 0 && (
              <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5 mb-6">
                <div className="flex items-center justify-between flex-wrap gap-2 mb-2">
                  <h2 className="font-brand text-lg text-ink-950">Print Fermentation Log</h2>
                  <button
                    onClick={() => setPrintJob({ type: "ferment", from: fermentPrintFrom, to: fermentPrintTo, lotIds: fermentPrintLotIds })}
                    className="font-body flex items-center gap-1.5 text-sm font-medium text-ink-900 hover:text-ink-700"
                  >
                    <Printer size={16} /> Print
                  </button>
                </div>
                <p className="font-body text-xs text-stone-500 mb-3">
                  Choose one or more lots and/or a date range. Leave lots unchecked to include all; leave dates blank for the full history.
                </p>
                <div className="flex flex-wrap gap-2 mb-3">
                  {data.ferment.map((lot) => {
                    const selected = fermentPrintLotIds.includes(lot.id);
                    return (
                      <button
                        key={lot.id}
                        type="button"
                        onClick={() =>
                          setFermentPrintLotIds((prev) => (selected ? prev.filter((id) => id !== lot.id) : [...prev, lot.id]))
                        }
                        className={`font-body text-xs px-2.5 py-1 rounded-full border transition-colors ${
                          selected ? "bg-ink-800 text-white border-ink-800" : "bg-white text-stone-600 border-stone-300 hover:border-ink-400"
                        }`}
                      >
                        {lot.tankId || "Untitled Tank"}
                      </button>
                    );
                  })}
                </div>
                <div className="flex items-center gap-2 flex-wrap">
                  <input
                    type="date"
                    value={fermentPrintFrom}
                    onChange={(e) => setFermentPrintFrom(e.target.value)}
                    className="font-body text-xs border border-stone-300 rounded px-2 py-1.5"
                  />
                  <span className="font-body text-xs text-stone-400">to</span>
                  <input
                    type="date"
                    value={fermentPrintTo}
                    onChange={(e) => setFermentPrintTo(e.target.value)}
                    className="font-body text-xs border border-stone-300 rounded px-2 py-1.5"
                  />
                </div>
              </div>
            )}

            <div className="flex items-center justify-between flex-wrap gap-2 mb-3">
              <h2 className="font-brand text-lg text-ink-950">Active Ferments</h2>
              <div className="flex items-center gap-2">
                <span className="font-body text-xs text-stone-500">Sort by</span>
                <select
                  value={fermentSort.field}
                  onChange={(e) => setFermentSort((p) => ({ ...p, field: e.target.value }))}
                  className="font-body text-xs border border-stone-300 rounded px-2 py-1.5"
                >
                  <option value="">Default</option>
                  {FERMENT_SORT_OPTIONS.map((opt) => (
                    <option key={opt.value} value={opt.value}>{opt.label}</option>
                  ))}
                </select>
                <button
                  onClick={() => setFermentSort((p) => ({ ...p, direction: p.direction === "asc" ? "desc" : "asc" }))}
                  className="text-stone-500 hover:text-ink-800 border border-stone-300 rounded p-1.5"
                  title={fermentSort.direction === "asc" ? "Ascending" : "Descending"}
                >
                  {fermentSort.direction === "asc" ? <ArrowUp size={14} /> : <ArrowDown size={14} />}
                </button>
              </div>
            </div>
            {activeLots.length === 0 ? (
              <p className="font-body text-sm text-stone-500 mb-6">No active ferments — start one above.</p>
            ) : (
              activeLots.map((lot) => (
                <FermentLotCard
                  key={lot.id}
                  lot={lot}
                  onAddReading={addReadingToLot}
                  onToggleComplete={toggleLotComplete}
                  onDeleteLot={deleteLot}
                  onUpdateLot={updateFermentLot}
                  onUpdateReading={updateReading}
                  onDeleteReading={deleteReading}
                  confirmAction={confirmAction}
                  harvestEntries={data.harvest}
                  vesselTypesList={vesselTypes}
                  onAddVesselType={addVesselType}
                />
              ))
            )}

            {completeLots.length > 0 && (
              <>
                <h2 className="font-brand text-lg text-stone-500 mt-8 mb-3">Completed Ferments</h2>
                {completeLots.map((lot) => (
                  <FermentLotCard
                    key={lot.id}
                    lot={lot}
                    onAddReading={addReadingToLot}
                    onToggleComplete={toggleLotComplete}
                    onDeleteLot={deleteLot}
                    onUpdateLot={updateFermentLot}
                    onUpdateReading={updateReading}
                    onDeleteReading={deleteReading}
                    confirmAction={confirmAction}
                    harvestEntries={data.harvest}
                    vesselTypesList={vesselTypes}
                    onAddVesselType={addVesselType}
                  />
                ))}
              </>
            )}

            {archivedLots.length > 0 && (
              <div className="bg-white border border-stone-200 rounded-lg overflow-hidden mt-8">
                <div className="px-4 py-3 border-b border-stone-200">
                  <h2 className="font-brand text-lg text-stone-500">Archived Vintages ({archivedLots.length})</h2>
                  <p className="font-body text-xs text-stone-400 mt-0.5">
                    Older completed ferments, tucked away so the current vintage stays front and center. Grouped by vintage year — still fully accessible, just not cluttering the main view.
                  </p>
                </div>
                {groupByVintage(archivedLots).map((group) => (
                  <FermentVintageGroup
                    key={group.key}
                    label={group.label}
                    lots={group.lots}
                    onAddReading={addReadingToLot}
                    onToggleComplete={toggleLotComplete}
                    onDeleteLot={deleteLot}
                    onUpdateLot={updateFermentLot}
                    onUpdateReading={updateReading}
                    onDeleteReading={deleteReading}
                    confirmAction={confirmAction}
                    harvestEntries={data.harvest}
                    vesselTypesList={vesselTypes}
                    onAddVesselType={addVesselType}
                  />
                ))}
              </div>
            )}
              </>
            )}
          </>
        ) : activeKey === "barrels" ? (
          <BarrelsTab
            data={data}
            onAddBarrel={addBarrel}
            onBulkAdd={bulkAddBarrels}
            onFillBarrel={fillBarrel}
            onEmptyBarrel={emptyBarrel}
            onUpdateBarrel={updateBarrel}
            onDeleteBarrel={deleteBarrel}
            onBulkEmpty={bulkEmptyBarrels}
            onBulkFill={bulkFillBarrels}
            onBulkSell={bulkSellBarrels}
            confirmAction={confirmAction}
          />
        ) : activeKey === "blending" ? (
          <BlendingTab
            data={data}
            onAddTasting={addTasting}
            onUpdateTasting={updateTasting}
            onDeleteTasting={deleteTasting}
            onAddBlend={addBlend}
            onUpdateBlend={updateBlend}
            onDeleteBlend={deleteBlend}
          />
        ) : activeKey === "thoPayroll" ? (
          <THOPayrollPage
            data={data}
            onAddThoEntry={addThoEntry}
            onUpdateThoEntry={updateThoEntry}
            onDeleteThoEntry={deleteThoEntry}
            confirmAction={confirmAction}
            tastingAssociates={tastingAssociates}
            onAddAssociate={addAssociate}
          />
        ) : activeKey === "aboutAlloro" ? (
          <AboutAlloroTab
            data={data}
            setPrintJob={setPrintJob}
            alloroStory={alloroStory}
            onUpdateStory={updateAlloroStory}
            vineyardMapImage={vineyardMapImage}
            onUpdateMapImage={updateVineyardMapImage}
            onAddThoEntry={addThoEntry}
            onUpdateThoEntry={updateThoEntry}
            onDeleteThoEntry={deleteThoEntry}
            confirmAction={confirmAction}
            vineyardBlocks={vineyardBlocks}
            onAddBlock={addVineyardBlock}
          />
        ) : activeKey === "techSheetBuilder" ? (
          <TechSheetsPanel
            data={data}
            techSheetMode={techSheetMode}
            setTechSheetMode={setTechSheetMode}
            onSave={saveTechSheet}
            onDelete={deleteTechSheet}
            onPrint={(sheetId) => setPrintJob({ type: "techSheet", sheetId })}
          />
        ) : activeKey === "compliance" ? (
          <ComplianceTab
            data={data}
            onAddThoEntry={addThoEntry}
            onUpdateThoEntry={updateThoEntry}
            onDeleteThoEntry={deleteThoEntry}
            confirmAction={confirmAction}
          />
        ) : activeKey === "calendar" ? (
          <MasterCalendar data={data} />
        ) : activeKey === "formulas" ? (
          <FormulasTab />
        ) : activeKey === "backup" ? (
          <BackupTab
            data={data}
            woCounter={woCounter}
            onRestore={restoreBackup}
            confirmAction={confirmAction}
            vineyardBlocks={vineyardBlocks}
            onAddBlock={addVineyardBlock}
            onRenameBlock={renameVineyardBlock}
            onDeleteBlock={deleteVineyardBlock}
            vesselTypes={vesselTypes}
            onAddVesselType={addVesselType}
            onRenameVesselType={renameVesselType}
            onDeleteVesselType={deleteVesselType}
            lotNames={lotNames}
            onAddLotName={addLotName}
            onRenameLotName={renameLotName}
            onDeleteLotName={deleteLotName}
            clones={clones}
            onAddClone={addClone}
            onRenameClone={renameClone}
            onDeleteClone={deleteClone}
            sprayPrograms={sprayPrograms}
            onAddSprayProgram={addSprayProgram}
            onRenameSprayProgram={renameSprayProgram}
            onDeleteSprayProgram={deleteSprayProgram}
            tastingAssociates={tastingAssociates}
            onAddAssociate={addAssociate}
            onRenameAssociate={renameAssociate}
            onDeleteAssociate={deleteAssociate}
            wineryName={wineryName}
            onUpdateWineryName={updateWineryName}
          />
        ) : (
          <>
            {["labResults"].includes(activeKey) && (
              <div className="flex justify-end mb-3">
                <button
                  onClick={() => setShowBulkImport((v) => !v)}
                  className="font-body flex items-center gap-1.5 text-sm font-medium text-ink-900 hover:text-ink-700 border border-ink-200 rounded-md px-3 py-1.5"
                >
                  <UploadCloud size={15} /> Import
                </button>
              </div>
            )}

            {["labResults"].includes(activeKey) && showBulkImport && (
              <BulkImportPanel fields={activeSection.fields} onImport={bulkImportRows} onClose={() => setShowBulkImport(false)} />
            )}

            {activeKey === "harvest" ? (
              <HarvestBatchEntryForm
                fields={activeSection.fields}
                onSubmit={submitHarvestBatch}
                saving={saving}
                vineyardBlocks={vineyardBlocks}
                onAddBlock={addVineyardBlock}
                clones={clones}
                onAddClone={addClone}
                defaultTareWeight={defaultTareWeight}
                onUpdateDefaultTareWeight={updateDefaultTareWeight}
              />
            ) : (
              <form onSubmit={handleSubmit} className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5 mb-6">
                <h2 className="font-brand text-lg text-ink-950 mb-3">New {activeSection.label} Entry</h2>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  {activeSection.fields.map((f) => (
                    <Field key={f.name} f={f} value={form[f.name]} onChange={(v) => handleChange(f.name, v)} blocksList={vineyardBlocks} onAddBlock={addVineyardBlock} clonesList={clones} onAddClone={addClone} associatesList={tastingAssociates} onAddAssociate={addAssociate} />
                  ))}
                </div>
                {error && <p className="font-body text-sm text-red-700 mt-3">{error}</p>}
                <button
                  type="submit"
                  disabled={saving}
                  className="font-body mt-4 flex items-center gap-2 bg-ink-900 hover:bg-ink-800 disabled:opacity-50 text-white text-sm font-medium px-4 py-2 rounded-md transition-colors"
                >
                  {saving ? <Loader2 size={16} className="animate-spin" /> : <Plus size={16} />}
                  Add Entry
                </button>
              </form>
            )}

            {activeKey === "fruitAnalysis" &&
              (() => {
                const blockGroups = groupFruitAnalysisByBlock(data.fruitAnalysis).filter((g) => g.entries.length >= 2);
                if (blockGroups.length === 0) return null;
                return (
                  <div className="bg-white border border-stone-200 rounded-lg p-4 sm:p-5 mb-4">
                    <h2 className="font-brand text-lg text-ink-950 mb-1">Ripening Curves</h2>
                    <p className="font-body text-xs text-stone-500 mb-3">
                      Brix and pH over time, by block — shows up automatically once a block has 2 or more sample dates.
                    </p>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      {blockGroups.map((g) => (
                        <RipeningChart key={g.block} block={g.block} entries={g.entries} />
                      ))}
                    </div>
                  </div>
                );
              })()}

            <div className="bg-white border border-stone-200 rounded-lg p-3 sm:p-4 mb-4 flex items-center gap-2 flex-wrap">
              <span className="font-body text-xs font-medium text-stone-600">Print:</span>
              <input
                type="date"
                value={printRanges[activeKey]?.from || ""}
                onChange={(e) => setPrintRanges((p) => ({ ...p, [activeKey]: { ...p[activeKey], from: e.target.value } }))}
                className="font-body text-xs border border-stone-300 rounded px-2 py-1.5"
              />
              <span className="font-body text-xs text-stone-400">to</span>
              <input
                type="date"
                value={printRanges[activeKey]?.to || ""}
                onChange={(e) => setPrintRanges((p) => ({ ...p, [activeKey]: { ...p[activeKey], to: e.target.value } }))}
                className="font-body text-xs border border-stone-300 rounded px-2 py-1.5"
              />
              <button
                onClick={() =>
                  setPrintJob({ type: "generic", key: activeKey, from: printRanges[activeKey]?.from || "", to: printRanges[activeKey]?.to || "" })
                }
                className="font-body flex items-center gap-1.5 text-sm font-medium text-ink-900 hover:text-ink-700"
              >
                <Printer size={16} /> Print
              </button>
              <span className="font-body text-xs text-stone-400 ml-1">(leave dates blank to print everything, or set just "from" for a single day)</span>
            </div>

            <div className="bg-white border border-stone-200 rounded-lg overflow-hidden">
              <div className="px-4 py-3 border-b border-stone-200 flex items-center justify-between">
                <h2 className="font-brand text-lg text-ink-950">{activeSection.label} Log</h2>
                <span className="font-body text-xs text-stone-500">
                  {(data[activeKey]?.length ?? 0)} entr{(data[activeKey]?.length ?? 0) === 1 ? "y" : "ies"}
                </span>
              </div>
              {(data[activeKey]?.length ?? 0) === 0 ? (
                <p className="font-body text-sm text-stone-500 px-4 py-8 text-center">No entries yet — add one above.</p>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-sm font-body">
                    <thead>
                      <tr className="bg-stone-100 text-stone-600 text-left">
                        {activeSection.fields.map((f) => {
                          const isSorted = tableSort[activeKey]?.field === f.name;
                          const sortable = f.type !== "photo";
                          return (
                            <th
                              key={f.name}
                              onClick={sortable ? () => toggleTableSort(activeKey, f.name) : undefined}
                              className={`px-3 py-2 font-medium whitespace-nowrap ${sortable ? "cursor-pointer select-none hover:text-ink-800" : ""}`}
                            >
                              <span className="inline-flex items-center gap-1">
                                {f.label}
                                {sortable && (
                                  isSorted ? (
                                    tableSort[activeKey].direction === "asc" ? <ArrowUp size={12} /> : <ArrowDown size={12} />
                                  ) : (
                                    <ArrowUpDown size={12} className="text-stone-300" />
                                  )
                                )}
                              </span>
                            </th>
                          );
                        })}
                        <th className="px-3 py-2 w-16"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {(tableSort[activeKey] ? sortRows(data[activeKey], tableSort[activeKey].field, tableSort[activeKey].direction) : data[activeKey]).map((row) => {
                        const isEditing = editingRow && editingRow.key === activeKey && editingRow.id === row.id;
                        return (
                          <tr key={row.id} className={`border-t border-stone-100 ${isEditing ? "bg-ink-50" : "hover:bg-stone-50"}`}>
                            {activeSection.fields.map((f) => (
                              <td key={f.name} className="px-2 py-2 align-top" style={{ minWidth: isEditing ? 130 : undefined }}>
                                {isEditing ? (
                                  <Field f={f} value={editingRow.form[f.name]} onChange={(v) => editRowChange(f.name, v)} hideLabel blocksList={vineyardBlocks} onAddBlock={addVineyardBlock} clonesList={clones} onAddClone={addClone} associatesList={tastingAssociates} onAddAssociate={addAssociate} />
                                ) : f.type === "photo" ? (
                                  row[f.name] ? (
                                    <a href={row[f.name]} target="_blank" rel="noopener noreferrer">
                                      <img src={row[f.name]} alt="Receipt" className="w-9 h-9 object-cover rounded border border-stone-300" />
                                    </a>
                                  ) : (
                                    "—"
                                  )
                                ) : (
                                  <span className="whitespace-nowrap">{row[f.name] || "—"}</span>
                                )}
                              </td>
                            ))}
                            <td className="px-3 py-2">
                              {isEditing ? (
                                <div className="flex items-center gap-2">
                                  <button onClick={saveEditRow} className="text-ink-700 hover:text-ink-900" title="Save">
                                    <Check size={15} />
                                  </button>
                                  <button onClick={cancelEditRow} className="text-stone-400 hover:text-stone-600" title="Cancel">
                                    <X size={15} />
                                  </button>
                                </div>
                              ) : (
                                <div className="flex items-center gap-2">
                                  <button onClick={() => startEditRow(activeKey, row)} className="text-stone-400 hover:text-ink-800" title="Edit">
                                    <Pencil size={14} />
                                  </button>
                                  <button onClick={() => handleDelete(row.id)} className="text-stone-400 hover:text-red-700" title="Delete entry">
                                    <Trash2 size={15} />
                                  </button>
                                </div>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </>
        )}

        <p className="font-body text-xs text-stone-400 mt-4 text-center">
          {totalEntries} total entries · Data is shared with everyone using this app link.
        </p>
      </main>

      {/* Print-only sheet — hidden on screen, shown via @media print. Content depends on printJob. */}
      {confirmRequest && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4">
          <div className="bg-white rounded-lg shadow-xl max-w-sm w-full p-5">
            <p className="font-body text-sm text-stone-700 mb-4">{confirmRequest.message}</p>
            <div className="flex justify-end gap-3">
              <button
                onClick={() => setConfirmRequest(null)}
                className="font-body text-sm px-3 py-1.5 rounded-md border border-stone-300 text-stone-600 hover:bg-stone-50"
              >
                Cancel
              </button>
              <button
                onClick={() => {
                  confirmRequest.onConfirm();
                  setConfirmRequest(null);
                }}
                className={`font-body text-sm px-3 py-1.5 rounded-md text-white font-medium ${
                  confirmRequest.tone === "danger" ? "bg-red-700 hover:bg-red-800" : "bg-ink-800 hover:bg-ink-900"
                }`}
              >
                {confirmRequest.confirmLabel}
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="print-sheet">
        {printJob?.type === "all" && (
          <>
            <PrintHeader subtitle={`Full Data Export — Printed ${new Date().toLocaleDateString()}`} />
            {buildExportSections(getSortedData()).map((section) => (
              <div key={section.title}>
                <PrintSectionTitle>{section.title}</PrintSectionTitle>
                {section.rows.length === 0 ? (
                  <PrintEmpty>No entries.</PrintEmpty>
                ) : (
                  <PrintTable
                    headers={section.headers}
                    rows={section.rows.map((row) => section.headers.map((h) => row[h]))}
                  />
                )}
              </div>
            ))}
          </>
        )}

        {printJob?.type === "workorders" && (
          <>
            <PrintHeader subtitle={`${workOrderFilterCategory} Work Orders — Printed ${new Date().toLocaleDateString()}`} />
            <PrintSectionTitle>Open Work Orders</PrintSectionTitle>
            {categoryFilteredOpenOrders.length === 0 ? (
              <PrintEmpty>No open work orders.</PrintEmpty>
            ) : (
              <PrintTable
                headers={["✓", "WO #", "Task", "Task Type", "Lots", "Barrels", "Directions", "Assigned To", "Priority", "Due", "Notes"]}
                rows={categoryFilteredOpenOrders.map((o) => [
                  "☐",
                  formatOrderNumber(o.orderNumber),
                  o.task,
                  o.taskType === "Additions" && o.additionType ? `${o.taskType} (${o.additionType})` : o.taskType,
                  Array.isArray(o.lots) ? o.lots.join(", ") : "",
                  Array.isArray(o.barrels) ? o.barrels.map((id) => data.barrels.find((b) => b.id === id)?.barrelNumber).filter(Boolean).join(", ") : "",
                  o.directions,
                  o.assignedTo,
                  o.priority,
                  o.date,
                  o.notes,
                ])}
              />
            )}
            {categoryFilteredCompletedOrders.length > 0 && (
              <>
                <PrintSectionTitle>Completed Work Orders</PrintSectionTitle>
                <PrintTable
                  headers={["WO #", "Task", "Task Type", "Assigned To", "Date Completed", "Notes"]}
                  rows={[...categoryFilteredCompletedOrders]
                    .sort((a, b) => (a.dateCompleted < b.dateCompleted ? 1 : -1))
                    .map((o) => [
                      formatOrderNumber(o.orderNumber),
                      o.task,
                      o.taskType === "Additions" && o.additionType ? `${o.taskType} (${o.additionType})` : o.taskType,
                      o.assignedTo,
                      o.dateCompleted,
                      o.notes,
                    ])}
                />
              </>
            )}
          </>
        )}

        {printJob?.type === "generic" &&
          (() => {
            const section = SIMPLE_SECTIONS.find((s) => s.key === printJob.key);
            const sortedSection = tableSort[printJob.key]
              ? sortRows(data[printJob.key], tableSort[printJob.key].field, tableSort[printJob.key].direction)
              : data[printJob.key];
            const rows = sortedSection.filter(
              (r) => (!printJob.from && !printJob.to) || inRange(r.date, printJob.from, printJob.to)
            );
            const printFields = section.fields.filter((f) => f.type !== "photo");
            const rangeLabel = printJob.from || printJob.to ? `${printJob.from || "…"} → ${printJob.to || "…"}` : "All dates";
            return (
              <>
                <PrintHeader subtitle={`${section.label} — ${rangeLabel} — Printed ${new Date().toLocaleDateString()}`} />
                {rows.length === 0 ? (
                  <PrintEmpty>No entries in this range.</PrintEmpty>
                ) : (
                  <PrintTable
                    headers={printFields.map((f) => f.label)}
                    rows={rows.map((r) => printFields.map((f) => r[f.name]))}
                  />
                )}
              </>
            );
          })()}

        {printJob?.type === "techSheet" &&
          (() => {
            const sheet = data.techSheets.find((s) => s.id === printJob.sheetId);
            if (!sheet) return null;
            const dates = harvestDatesFor(sheet.harvestRefs, data.harvest);
            const row = (label, value) =>
              value ? (
                <tr>
                  <td style={{ padding: "6px 0", color: "#78716c", fontSize: 12, width: 160, verticalAlign: "top" }}>{label}</td>
                  <td style={{ padding: "6px 0", color: "#292524", fontSize: 13 }}>{value}</td>
                </tr>
              ) : null;
            return (
              <>
                <PrintHeader subtitle="Tech Sheet" />
                {sheet.bottleShotImage && (
                  <img src={sheet.bottleShotImage} alt={sheet.wineName} style={{ width: 140, height: 140, objectFit: "cover", borderRadius: 8, marginBottom: 12 }} />
                )}
                <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 26, fontWeight: 700, color: "#1D1915", margin: "0 0 2px" }}>
                  {sheet.wineName || "Untitled Wine"}
                </h1>
                <p style={{ fontFamily: "Arial, sans-serif", fontSize: 13, color: "#78716c", margin: "0 0 16px" }}>
                  {sheet.vintage ? `${sheet.vintage} Vintage` : "Vintage not set"}
                </p>
                <p style={{ fontFamily: "Arial, sans-serif", fontSize: 11, fontWeight: 700, color: "#57534e", textTransform: "uppercase", letterSpacing: 0.5, margin: "12px 0 4px" }}>
                  Vineyard
                </p>
                <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: "Arial, sans-serif" }}>
                  <tbody>
                    {row("Harvest Date(s)", dates.join(", "))}
                    {row("Budbreak", sheet.budbreakDate)}
                    {row("Veraison", sheet.veraisonDate)}
                    {row("Variety & Clones", sheet.varietyCloneBlend)}
                  </tbody>
                </table>
                <p style={{ fontFamily: "Arial, sans-serif", fontSize: 11, fontWeight: 700, color: "#57534e", textTransform: "uppercase", letterSpacing: 0.5, margin: "16px 0 4px" }}>
                  Winemaking
                </p>
                <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: "Arial, sans-serif" }}>
                  <tbody>
                    {row("New Oak", sheet.newOakPercent !== "" && sheet.newOakPercent != null ? `${sheet.newOakPercent}%` : "")}
                    {row("Elevage", sheet.elevageDetails)}
                  </tbody>
                </table>
                {sheet.vintageNotes && (
                  <>
                    <p style={{ fontFamily: "Arial, sans-serif", fontSize: 11, fontWeight: 700, color: "#57534e", textTransform: "uppercase", letterSpacing: 0.5, margin: "16px 0 4px" }}>
                      Vintage Notes
                    </p>
                    <p style={{ fontFamily: "Arial, sans-serif", fontSize: 13, color: "#292524", whiteSpace: "pre-wrap" }}>{sheet.vintageNotes}</p>
                  </>
                )}
              </>
            );
          })()}

        {printJob?.type === "fermentDayLog" &&
          (() => {
            const date = printJob.date;
            const dayLots = data.ferment
              .filter((lot) => lot.readings.some((r) => r.date === date))
              .sort((a, b) => (a.tankId || "").localeCompare(b.tankId || ""));
            const cellFor = (lot, sess) => lot.readings.find((r) => r.date === date && r.session === sess);
            const odCell = (reading) =>
              [reading?.workDone?.includes("Pump Over") ? "O" : "", reading?.workDone?.includes("Punch Down") ? "D" : ""]
                .filter(Boolean)
                .join("/") || "—";
            return (
              <>
                <PrintHeader subtitle={`Fermentation Log — ${date}`} />
                {dayLots.length === 0 ? (
                  <PrintEmpty>No fermentation entries logged on this date.</PrintEmpty>
                ) : (
                  <PrintTable
                    headers={["Lot", "Vessel", "AM Temp", "AM Brix", "AM O/D", "AM Add.", "PM Temp", "PM Brix", "PM O/D", "PM Add.", "Notes"]}
                    rows={dayLots.map((lot) => {
                      const am = cellFor(lot, "A.M.");
                      const pm = cellFor(lot, "P.M.");
                      return [
                        lot.tankId || "Untitled",
                        lot.vessel || "",
                        am?.temp || "",
                        am?.brix || "",
                        odCell(am),
                        (am?.additions || []).join(", "),
                        pm?.temp || "",
                        pm?.brix || "",
                        odCell(pm),
                        (pm?.additions || []).join(", "),
                        [am?.notes, pm?.notes].filter(Boolean).join(" / "),
                      ];
                    })}
                  />
                )}
              </>
            );
          })()}

        {printJob?.type === "ferment" &&
          (() => {
            const sortedFerment = sortRows(data.ferment, fermentSort.field, fermentSort.direction);
            const lots = printJob.lotIds.length > 0 ? sortedFerment.filter((l) => printJob.lotIds.includes(l.id)) : sortedFerment;
            const rangeLabel = printJob.from || printJob.to ? `${printJob.from || "…"} → ${printJob.to || "…"}` : "All dates";
            return (
              <>
                <PrintHeader subtitle={`Fermentation — ${rangeLabel} — Printed ${new Date().toLocaleDateString()}`} />
                {lots.length === 0 ? (
                  <PrintEmpty>No fermentation lots selected.</PrintEmpty>
                ) : (
                  lots.map((lot) => {
                    const readings = lot.readings.filter(
                      (r) => (!printJob.from && !printJob.to) || inRange(r.date, printJob.from, printJob.to)
                    );
                    return (
                      <div key={lot.id}>
                        <PrintSectionTitle>
                          {lot.tankId || "Untitled Tank"}
                          {lot.vessel ? ` (${lot.vessel})` : ""}
                          {lot.variety ? ` — ${lot.variety}` : ""} ({lot.status})
                        </PrintSectionTitle>
                        {readings.length === 0 ? (
                          <PrintEmpty>No readings in this range.</PrintEmpty>
                        ) : (
                          <PrintTable
                            headers={["Date", "Session", "Work Done", "Additions", "Brix", "Temp (°F)", "pH", "Notes"]}
                            rows={readings.map((r) => [
                              r.date,
                              r.session,
                              Array.isArray(r.workDone) ? r.workDone.join(", ") : "",
                              Array.isArray(r.additions) ? r.additions.join(", ") : "",
                              r.brix,
                              r.temp,
                              r.ph,
                              r.notes,
                            ])}
                          />
                        )}
                      </div>
                    );
                  })
                )}
              </>
            );
          })()}
      </div>
    </div>
  );
}

export default function WineryDataTracker() {
  return (
    <AppErrorBoundary>
      <WineryDataTrackerInner />
    </AppErrorBoundary>
  );
}
