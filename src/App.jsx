import { useState, useEffect, useCallback, useMemo, Fragment, Component } from "react";
import { Analytics } from "@vercel/analytics/react";
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
  FileCheck,
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
  {
    key: "labelApprovals",
    label: "Label Approvals",
    icon: FileCheck,
    sheetName: "Label Approvals",
    fields: [
      { name: "brandName", label: "Brand Name", type: "text" },
      { name: "fancifulName", label: "Fanciful Name (wine name on label)", type: "text" },
      { name: "vintage", label: "Vintage", type: "text", optional: true },
      { name: "varietal", label: "Grape Varietal(s)", type: "text", optional: true },
      { name: "appellation", label: "Wine Appellation", type: "text", optional: true },
      { name: "classType", label: "TTB Class/Type Description", type: "text", optional: true },
      { name: "ttbId", label: "TTB ID", type: "text" },
      { name: "serialNumber", label: "Serial Number", type: "text", optional: true },
      { name: "applicationType", label: "Application Type", type: "select", options: ["Certificate of Label Approval", "Certificate of Exemption", "Distinctive Liquor Bottle Approval", "Resubmission After Rejection"] },
      { name: "dateApplied", label: "Date Applied", type: "date", optional: true },
      { name: "dateIssued", label: "Date Issued", type: "date", optional: true },
      { name: "status", label: "Status", type: "select", options: ["Approved", "Pending", "Rejected"] },
      { name: "expirationDate", label: "Expiration Date (if any)", type: "date", optional: true },
      { name: "labelDocument", label: "COLA Certificate / Label Images", type: "document", optional: true },
      { name: "notes", label: "Notes", type: "textarea", optional: true },
    ],
  },
  {
    key: "ttbFilings",
    label: "TTB Filings",
    icon: FileCheck,
    sheetName: "TTB Filings",
    fields: [
      { name: "period", label: "Filing Period (e.g. \"2026\" or \"Q1 2026\")", type: "text" },
      { name: "dueDate", label: "Due Date", type: "date", optional: true },
      { name: "dateFiled", label: "Date Filed", type: "date", optional: true },
      { name: "galOnHandStart", label: "Gallons On Hand — Period Start", type: "number", optional: true },
      { name: "galOnHandEnd", label: "Gallons On Hand — Period End", type: "number", optional: true },
      { name: "physicalInventoryDate", label: "Physical Inventory Taken On", type: "date", optional: true },
      { name: "sourceRecordsGathered", label: "Source Records Gathered & Reconciled", type: "checkbox-group", options: ["Yes"], optional: true },
      { name: "filingDocument", label: "Filed 5120.17 / Excise Tax Return Copy", type: "document", optional: true },
      { name: "notes", label: "Notes", type: "textarea", optional: true },
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

// Most COLAs never expire once approved — expiration only applies to conditional "use-up"
// approvals, so a blank expiration date on an Approved label just means Approved, not Active.
function computeLabelStatus(label) {
  if (label.status === "Rejected") return "Rejected";
  if (label.status === "Pending") return "Pending";
  if (!label.expirationDate) return "Approved";
  const today = todayISO();
  if (today >= label.expirationDate) return "Expired";
  const exp = new Date(label.expirationDate + "T00:00:00");
  const reminder = new Date(exp);
  reminder.setDate(reminder.getDate() - 45);
  const reminderISO = reminder.toISOString().slice(0, 10);
  if (today >= reminderISO) return "Expiring Soon";
  return "Approved";
}

// Due date is entered by the winery, same as permit expiration — not auto-calculated, since a
// wrong calculation on a compliance deadline is worse than no calculation at all.
function computeTTBFilingStatus(filing) {
  if (filing.dateFiled) return "Filed";
  if (!filing.dueDate) return "No Due Date Set";
  const today = todayISO();
  if (today >= filing.dueDate) return "Overdue";
  const due = new Date(filing.dueDate + "T00:00:00");
  const reminder = new Date(due);
  reminder.setDate(reminder.getDate() - 45);
  const reminderISO = reminder.toISOString().slice(0, 10);
  if (today >= reminderISO) return "Due Soon";
  return "Upcoming";
}

// Plain-English reference for when the next filing is due, based on the winery's own stated
// frequency — informational only, the winery still enters the actual due date themselves.
function ttbFilingFrequencyHint(frequency) {
  if (frequency === "Monthly") return "Monthly filers: due the 15th of the month after each reportable month.";
  if (frequency === "Quarterly") return "Quarterly filers: due April 15, July 15, October 15, and January 15 (15 days after each calendar quarter).";
  return "Annual filers: due January 15 of the year following the filing year.";
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

function ComplianceOverview({ permits, shipments, labels, ttbFilings, setSubTab }) {
  const statusByState = bestPermitStatusByState(permits);
  const licensed = Object.entries(statusByState).filter(([, s]) => s === "Active").map(([st]) => st).sort();
  const expiringSoon = Object.entries(statusByState).filter(([, s]) => s === "Expiring Soon").map(([st]) => st).sort();
  const expired = Object.entries(statusByState).filter(([, s]) => s === "Expired").map(([st]) => st).sort();
  const conflicts = complianceConflicts(shipments, permits).sort();

  const labelName = (l) => l.fancifulName || l.brandName || "Untitled label";
  const approvedLabels = labels.filter((l) => computeLabelStatus(l) === "Approved");
  const labelsNeedingAttention = labels.filter((l) => ["Rejected", "Expired", "Expiring Soon"].includes(computeLabelStatus(l)));
  const ttbNeedingAttention = ttbFilings.filter((f) => ["Overdue", "Due Soon"].includes(computeTTBFilingStatus(f)));

  const Chip = ({ text, color }) => (
    <span className={`font-body text-xs rounded px-2 py-0.5 ${color}`}>{text}</span>
  );

  return (
    <div className="space-y-4">
      <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
        <p className="font-body text-xs text-amber-800">{COMPLIANCE_DISCLAIMER}</p>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
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
          <p className="font-body text-xs text-stone-500 mb-2">Approved labels</p>
          <p className="font-brand text-2xl text-ink-950 mb-2">{approvedLabels.length}</p>
          <div className="flex flex-wrap gap-1">
            {labels.length === 0 ? <span className="font-body text-xs text-stone-400">None on file yet</span> : <span className="font-body text-xs text-stone-400">{labels.length} total tracked</span>}
          </div>
        </div>
        <div className="bg-white border border-stone-200 rounded-lg p-4">
          <p className="font-body text-xs text-stone-500 mb-2">Needs attention</p>
          <p className="font-brand text-2xl text-red-700 mb-2">{expired.length + conflicts.length + labelsNeedingAttention.length + ttbNeedingAttention.length}</p>
          <div className="flex flex-wrap gap-1">
            {expired.length === 0 && conflicts.length === 0 && labelsNeedingAttention.length === 0 && ttbNeedingAttention.length === 0 ? (
              <span className="font-body text-xs text-stone-400">None</span>
            ) : (
              <>
                {expired.map((s) => <Chip key={"exp-" + s} text={`${s} — expired`} color="bg-red-50 text-red-700" />)}
                {conflicts.map((s) => <Chip key={"conf-" + s} text={`${s} — no active permit on file`} color="bg-red-50 text-red-700" />)}
                {labelsNeedingAttention.map((l) => <Chip key={l.id} text={`${labelName(l)} — ${computeLabelStatus(l).toLowerCase()}`} color="bg-red-50 text-red-700" />)}
                {ttbNeedingAttention.map((f) => <Chip key={f.id} text={`TTB ${f.period} — ${computeTTBFilingStatus(f).toLowerCase()}`} color="bg-red-50 text-red-700" />)}
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

function ComplianceTab({ data, onAddThoEntry, onUpdateThoEntry, onDeleteThoEntry, confirmAction, associatesList, onAddAssociate, ttbFilingFrequency }) {
  const [subTab, setSubTab] = useState("overview");
  const permitFields = SIMPLE_SECTIONS.find((s) => s.key === "statePermits").fields;
  const shipmentFields = SIMPLE_SECTIONS.find((s) => s.key === "shipmentLog").fields;
  const referenceFields = SIMPLE_SECTIONS.find((s) => s.key === "stateRuleReference").fields;
  const labelFields = SIMPLE_SECTIONS.find((s) => s.key === "labelApprovals").fields;
  const ttbFields = SIMPLE_SECTIONS.find((s) => s.key === "ttbFilings").fields;

  const sortedPermits = [...data.statePermits].sort((a, b) => (a.expirationDate || "9999").localeCompare(b.expirationDate || "9999"));
  const sortedShipments = [...data.shipmentLog].sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  const sortedReference = [...data.stateRuleReference].sort((a, b) => (a.state || "").localeCompare(b.state || ""));
  const sortedLabels = [...data.labelApprovals].sort((a, b) => (a.brandName + a.fancifulName).localeCompare(b.brandName + b.fancifulName));
  const sortedTtb = [...data.ttbFilings].sort((a, b) => (b.period || "").localeCompare(a.period || ""));

  const SUB_TABS = [
    { key: "overview", label: "Overview" },
    { key: "permits", label: "My Permits" },
    { key: "shipments", label: "Shipment Log" },
    { key: "labels", label: "Label Approvals" },
    { key: "ttb", label: "TTB Filings" },
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
        <ComplianceOverview permits={data.statePermits} shipments={data.shipmentLog} labels={data.labelApprovals} ttbFilings={data.ttbFilings} setSubTab={setSubTab} />
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
      ) : subTab === "labels" ? (
        <div className="space-y-4">
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
            <p className="font-body text-xs text-amber-800">
              This tracks labels you've already submitted or had approved — it doesn't submit anything to TTB or review your label for you.
            </p>
          </div>
          <LabelPreSubmissionChecklist />
          <SimpleDataPanel
            title="Label Approvals"
            fields={labelFields}
            rows={sortedLabels}
            onAdd={(entry) => onAddThoEntry("labelApprovals", entry)}
            onUpdate={(id, changes) => onUpdateThoEntry("labelApprovals", id, changes)}
            onDelete={(id) => onDeleteThoEntry("labelApprovals", id)}
            confirmAction={confirmAction}
          />
        </div>
      ) : subTab === "ttb" ? (
        <div className="space-y-4">
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
            <p className="font-body text-xs text-amber-800">
              This organizes the records TTB actually checks during an audit — it doesn't calculate tax owed, generate a filing, or submit anything on your behalf.
            </p>
          </div>
          <div className="bg-white border border-stone-200 rounded-lg p-3">
            <p className="font-body text-xs text-stone-600">
              Filing frequency: <span className="font-medium text-ink-900">{ttbFilingFrequency}</span> — {ttbFilingFrequencyHint(ttbFilingFrequency)}
              <span className="block text-stone-400 mt-1">Change this in Backup if it's wrong.</span>
            </p>
          </div>
          <TTBSourceRecordsChecklist />
          <SimpleDataPanel
            title="TTB Filings"
            fields={ttbFields}
            rows={sortedTtb}
            onAdd={(entry) => onAddThoEntry("ttbFilings", entry)}
            onUpdate={(id, changes) => onUpdateThoEntry("ttbFilings", id, changes)}
            onDelete={(id) => onDeleteThoEntry("ttbFilings", id)}
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

// Common, genuinely technical reasons TTB rejects label submissions — collapsed by default so it
// doesn't clutter the page, but there for someone about to submit a new label. This is a
// checklist to prompt a careful look, not a determination that a label will pass.
function LabelPreSubmissionChecklist() {
  const [expanded, setExpanded] = useState(false);
  const items = [
    "Government health warning present, worded exactly as required, and not split awkwardly across a seam or fold",
    "Brand name, class/type designation, alcohol content, and net contents all present and legible",
    "Bottler/producer name and address (\"Grown, Produced & Bottled by...\") present and accurate",
    "Sulfite declaration included if applicable (\"Contains Sulfites\")",
    "Mandatory text meets minimum type size and appears on a contrasting background",
    "Wine appellation matches what's actually permitted for that grape source",
    "If resubmitting after a rejection, the prior TTB ID is referenced on the new application",
  ];
  return (
    <div className="bg-white border border-stone-200 rounded-lg overflow-hidden">
      <button onClick={() => setExpanded((v) => !v)} className="w-full flex items-center justify-between px-4 py-3 hover:bg-stone-50">
        <span className="font-body text-sm font-medium text-stone-700">Before you submit — common rejection reasons</span>
        {expanded ? <ChevronUp size={16} className="text-stone-400" /> : <ChevronDown size={16} className="text-stone-400" />}
      </button>
      {expanded && (
        <div className="px-4 pb-4">
          <p className="font-body text-xs text-stone-500 mb-2">
            Most TTB label rejections come down to formatting details, not the brand or artwork itself. This isn't a review of your specific label — just the recurring issues worth a careful look before you submit.
          </p>
          <ul className="space-y-1.5">
            {items.map((item, i) => (
              <li key={i} className="font-body text-xs text-stone-600 flex items-start gap-2">
                <span className="text-stone-300 mt-0.5">•</span>
                {item}
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}

// The most common issue TTB examiners find isn't the 5120.17 report itself — it's that the
// source records supporting it are missing or inadequate. This checklist exists to catch that
// before it becomes an audit finding, not to generate or validate the report itself.
function TTBSourceRecordsChecklist() {
  const [expanded, setExpanded] = useState(false);
  const items = [
    "Physical inventory taken at least once this year, with the count reconciled against book inventory",
    "Any breakage or loss during the period has a written record — TTB specifically checks for this",
    "Transfer-in-bond records (bills of lading) show proprietor and consignee name, address, and registry number",
    "Bulk wine and bottled wine figures are tracked separately by tax class (alcohol % category)",
    "Records of any additions, sweetening, or blending are dated and traceable to the affected lots",
    "All source records are kept on the bonded premises and accessible on request — not just in someone's inbox",
  ];
  return (
    <div className="bg-white border border-stone-200 rounded-lg overflow-hidden">
      <button onClick={() => setExpanded((v) => !v)} className="w-full flex items-center justify-between px-4 py-3 hover:bg-stone-50">
        <span className="font-body text-sm font-medium text-stone-700">What TTB actually checks — source record checklist</span>
        {expanded ? <ChevronUp size={16} className="text-stone-400" /> : <ChevronDown size={16} className="text-stone-400" />}
      </button>
      {expanded && (
        <div className="px-4 pb-4">
          <p className="font-body text-xs text-stone-500 mb-2">
            In TTB audits, missing or inadequate source records — not errors on the report itself — are the most common finding. This app already tracks much of this through Harvest, Fermentation, Barrels, Blending, and Bottling; this checklist is a prompt to make sure the pieces TTB actually asks for are complete before they're needed.
          </p>
          <ul className="space-y-1.5">
            {items.map((item, i) => (
              <li key={i} className="font-body text-xs text-stone-600 flex items-start gap-2">
                <span className="text-stone-300 mt-0.5">•</span>
                {item}
              </li>
            ))}
          </ul>
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

function FermentOverviewCard({ lot, onClick, onCloseOut, confirmAction }) {
  const stageIndex = computeStageIndex(lot);
  const latestReading = [...lot.readings].sort((a, b) => (a.date < b.date ? 1 : -1))[0];
  return (
    <div className="bg-white border border-stone-200 rounded-lg p-4 hover:border-ink-300 transition-colors">
      <div className="flex items-start justify-between gap-3 mb-4">
        <button type="button" onClick={onClick} className="text-left flex-1 min-w-0">
          <p className="font-brand text-base text-ink-950">{lot.tankId || "Untitled"}</p>
          <p className="font-body text-xs text-stone-500">
            {[lot.vessel, lot.variety, lot.vintage].filter(Boolean).join(" · ") || "No details yet"}
          </p>
        </button>
        <div className="text-right shrink-0 flex items-start gap-2">
          {latestReading && (
            <div>
              <p className="font-body text-xs text-stone-400">As of {latestReading.date}</p>
              <p className="font-body text-sm text-stone-700">
                {latestReading.brix !== "" && latestReading.brix != null ? `${latestReading.brix}° Brix` : ""}
                {latestReading.temp !== "" && latestReading.temp != null ? ` · ${latestReading.temp}°F` : ""}
              </p>
            </div>
          )}
          <button
            type="button"
            onClick={() =>
              confirmAction(`Close out ${lot.tankId || "this lot"}? It'll move to the vintage archive and drop off this overview.`, () => onCloseOut(lot.id))
            }
            className="font-body text-xs text-stone-400 hover:text-ink-800 border border-stone-200 rounded-md px-2 py-1 whitespace-nowrap"
            title="Close out this ferment"
          >
            Close Out
          </button>
        </div>
      </div>
      <button type="button" onClick={onClick} className="w-full text-left">
        <FermentStageStepper stageIndex={stageIndex} />
      </button>
      <div className="flex flex-wrap gap-x-3 gap-y-1 mt-4 font-body text-xs text-stone-500">
        {lot.startDate && <span>Started {lot.startDate}</span>}
        {lot.dateCompleted && <span>· Primary done {lot.dateCompleted}</span>}
        {lot.mlInoculationDate && <span>· ML started {lot.mlInoculationDate}</span>}
        {lot.mlCompleteDate && <span>· ML done {lot.mlCompleteDate}</span>}
      </div>
    </div>
  );
}

function FermentOverview({ lots, onSelectLot, onCloseOut, onCloseOutAll, confirmAction }) {
  const currentLots = lots.filter((lot) => !lot.archived);
  return (
    <div className="space-y-3 mb-6">
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <div>
          <h2 className="font-brand text-xl text-ink-950">Fermentation Overview</h2>
          <p className="font-body text-xs text-stone-500">
            Every lot not yet closed out. Tap a card to open its details.
          </p>
        </div>
        {currentLots.length > 0 && (
          <button
            type="button"
            onClick={() =>
              confirmAction(
                `Close out all ${currentLots.length} lot${currentLots.length === 1 ? "" : "s"} shown here? They'll move to the vintage archive — use this once a vintage is fully done and bottled, to start the next one with a clean overview.`,
                () => onCloseOutAll(currentLots.map((l) => l.id))
              )
            }
            className="font-body text-xs font-medium text-stone-500 hover:text-ink-800 border border-stone-300 rounded-md px-3 py-1.5"
          >
            Close Out All
          </button>
        )}
      </div>
      {currentLots.length === 0 ? (
        <p className="font-body text-sm text-stone-500 bg-white border border-stone-200 rounded-lg p-6 text-center">
          No ferments currently in progress.
        </p>
      ) : (
        currentLots.map((lot) => (
          <FermentOverviewCard key={lot.id} lot={lot} onClick={() => onSelectLot(lot.id)} onCloseOut={onCloseOut} confirmAction={confirmAction} />
        ))
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

function BackupTab({ data, woCounter, onRestore, confirmAction, vineyardBlocks, onAddBlock, onRenameBlock, onDeleteBlock, vesselTypes, onAddVesselType, onRenameVesselType, onDeleteVesselType, lotNames, onAddLotName, onRenameLotName, onDeleteLotName, clones, onAddClone, onRenameClone, onDeleteClone, sprayPrograms, onAddSprayProgram, onRenameSprayProgram, onDeleteSprayProgram, tastingAssociates, onAddAssociate, onRenameAssociate, onDeleteAssociate, wineryName, onUpdateWineryName, ttbFilingFrequency, onUpdateTtbFilingFrequency }) {
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
        <h2 className="font-brand text-lg text-ink-950 mb-1">TTB Filing Frequency</h2>
        <p className="font-body text-xs text-stone-500 mb-3">
          How often this winery files the Report of Wine Premises Operations (TTB Form 5120.17) — used to compute filing due dates in Compliance → TTB Filings. Based on gallons of wine on hand, not something the app can determine for you.
        </p>
        <select
          value={ttbFilingFrequency}
          onChange={(e) => onUpdateTtbFilingFrequency(e.target.value)}
          className="font-body border border-stone-300 rounded-md px-3 py-2 text-sm w-full max-w-xs focus:outline-none focus:ring-2 focus:ring-ink-800"
        >
          <option value="Annual">Annual</option>
          <option value="Quarterly">Quarterly</option>
          <option value="Monthly">Monthly</option>
        </select>
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
const GDS_LOGO_LIGHT = "data:image/png;base64,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";
const GDS_LOGO_FULL = "data:image/png;base64,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";

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
        <img src={GDS_LOGO_FULL} alt="GrapeDaddy Winery Solutions" className="w-40 mx-auto mb-4" />
        <h1 className="font-brand text-2xl text-ink-950 text-center mb-3">{wineryName || "Your Winery"}</h1>
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
  const [ttbFilingFrequency, setTtbFilingFrequency] = useState("Annual");
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
  const [fermentPrintIncludeArchived, setFermentPrintIncludeArchived] = useState(false);

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
        const res = await storage.get("labelApprovals", true);
        results.labelApprovals = res ? JSON.parse(res.value) : [];
      } catch {
        results.labelApprovals = [];
      }

      try {
        const res = await storage.get("ttbFilings", true);
        results.ttbFilings = res ? JSON.parse(res.value) : [];
      } catch {
        results.ttbFilings = [];
      }

      try {
        const res = await storage.get("ttb_filing_frequency", true);
        if (res) setTtbFilingFrequency(res.value);
      } catch {
        // stays at "Annual" default
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

        // One-time seed: import the two real approved COLAs Alloro shared during setup, so Label
        // Approvals starts with real data instead of an empty tab.
        try {
          await storage.get("label_approvals_seed_v1", true);
        } catch {
          try {
            const seedLabels = [
              {
                id: genId(),
                brandName: "Alloro Vineyard",
                fancifulName: "Blanc de Noirs",
                vintage: "2022",
                varietal: "Pinot noir",
                appellation: "Laurelwood District",
                classType: "Table Red Wine",
                ttbId: "26241001000077",
                serialNumber: "26BDNR",
                applicationType: "Certificate of Label Approval",
                dateApplied: "2026-08-29",
                dateIssued: "2026-09-01",
                status: "Approved",
                expirationDate: "",
                labelDocument: "",
                notes: "Sparkling wine (Metodo Classico), classed by TTB as Table Red Wine. 12.7% alc/vol, 750mL.",
              },
              {
                id: genId(),
                brandName: "Alloro Vineyard",
                fancifulName: "La Casa Series Estate Arneis",
                vintage: "2025",
                varietal: "Arneis",
                appellation: "Laurelwood District",
                classType: "Table Red Wine",
                ttbId: "26241001000083",
                serialNumber: "26ARNS",
                applicationType: "Certificate of Label Approval",
                dateApplied: "2026-08-29",
                dateIssued: "2026-09-01",
                status: "Approved",
                expirationDate: "",
                labelDocument: "",
                notes: "13.9% alc/vol, 750mL. Class/type as issued by TTB — recorded as-is from the certificate, not corrected.",
              },
            ];
            const mergedLabels = [...seedLabels, ...(results.labelApprovals || [])];
            await storage.set("labelApprovals", JSON.stringify(mergedLabels), true);
            results.labelApprovals = mergedLabels;
            await storage.set("label_approvals_seed_v1", "true", true);
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

  const updateTtbFilingFrequency = (value) => {
    setTtbFilingFrequency(value);
    guardedStorageSet("ttb_filing_frequency", value, true).catch(() => {
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

  // Closing out a lot is the explicit "this ferment is done, move it to the vintage archive"
  // signal — independent of ML status, since not every wine undergoes malolactic fermentation
  // and relying on "ML Complete" as an implicit done-marker was exactly what left last year's
  // finished, bottled ferments stuck showing as current.
  const closeOutLot = async (lotId) => {
    const updated = data.ferment.map((lot) =>
      lot.id === lotId ? { ...lot, archived: true, archivedDate: todayISO() } : lot
    );
    setData((prev) => ({ ...prev, ferment: updated }));
    await persist("ferment", updated);
  };

  const closeOutAllActiveLots = async (lotIds) => {
    const idSet = new Set(lotIds);
    const updated = data.ferment.map((lot) =>
      idSet.has(lot.id) ? { ...lot, archived: true, archivedDate: todayISO() } : lot
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
  const activeLots = sortedFermentForDisplay.filter((l) => l.status === "Active" && !l.archived);
  const completeLots = sortedFermentForDisplay.filter((l) => l.status === "Complete" && !l.archived);
  const archivedLots = sortedFermentForDisplay.filter((l) => l.archived);

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
            <img src={GDS_LOGO_LIGHT} alt="GrapeDaddy" className="w-14 h-14 sm:w-16 sm:h-16 shrink-0" />
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
              <FermentOverview
                lots={data.ferment}
                onSelectLot={() => setFermentViewMode("detailed")}
                onCloseOut={closeOutLot}
                onCloseOutAll={closeOutAllActiveLots}
                confirmAction={confirmAction}
              />
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
                  {(fermentPrintIncludeArchived ? data.ferment : data.ferment.filter((lot) => !lot.archived)).map((lot) => {
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
                {data.ferment.some((lot) => lot.archived) && (
                  <label className="font-body flex items-center gap-2 text-xs text-stone-500 mb-3">
                    <input
                      type="checkbox"
                      checked={fermentPrintIncludeArchived}
                      onChange={(e) => setFermentPrintIncludeArchived(e.target.checked)}
                    />
                    Include closed-out lots from past vintages
                  </label>
                )}
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
            ttbFilingFrequency={ttbFilingFrequency}
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
            ttbFilingFrequency={ttbFilingFrequency}
            onUpdateTtbFilingFrequency={updateTtbFilingFrequency}
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
      <Analytics />
    </AppErrorBoundary>
  );
}
