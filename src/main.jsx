import React, { useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { Document, Paragraph, TextRun, AlignmentType, BorderStyle, TabStopType, Tab, Packer } from "docx";
import { saveAs } from "file-saver";
import { jsPDF } from "jspdf";
import {
  FileText, Download, Sparkles, CheckCircle2, AlertCircle, Plus, Trash2,
  UserRound, BriefcaseBusiness, GraduationCap, Code2, RotateCcw, FileUp,
  Target, FolderGit2, Gauge, X
} from "lucide-react";
import mammoth from "mammoth/mammoth.browser";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import workerSrc from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?url";
import "./styles.css";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerSrc;

/* ---------- data model ---------- */
const blankExp = () => ({ role: "", company: "", location: "", start: "", end: "", bullets: "" });
const blankEdu = () => ({ degree: "", institution: "", location: "", year: "", details: "" });
const blankProj = () => ({ name: "", technologies: "", description: "" });
const createBlankResume = () => ({
  personal: { name: "", title: "", email: "", phone: "", location: "", linkedin: "", website: "" },
  summary: "", experience: [blankExp()], education: [blankEdu()], skills: "",
  projects: [blankProj()], certifications: "", jobDescription: ""
});

/* ---------- text extraction ---------- */
async function readPdf(buffer) {
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;
  const out = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const { items } = await (await pdf.getPage(p)).getTextContent();
    const rows = [];
    for (const it of items) {
      if (!it.str || !it.str.trim()) continue;
      const x = it.transform[4], y = it.transform[5];
      let row = rows.find(r => Math.abs(r.y - y) <= 3);
      if (!row) rows.push((row = { y, items: [] }));
      row.items.push({ x, str: it.str, w: it.width });
    }
    rows.sort((a, b) => b.y - a.y);
    for (const r of rows) {
      r.items.sort((a, b) => a.x - b.x);
      let line = "", end = null;
      for (const i of r.items) {
        if (end !== null && i.x - end > 1.5 && !line.endsWith(" ") && !i.str.startsWith(" ")) line += " ";
        line += i.str;
        end = i.x + i.w;
      }
      out.push(line.trim());
    }
  }
  return out.join("\n");
}

const cleanText = t => (t || "").replace(/\r/g, "\n").replace(/\u00a0/g, " ")
  .replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n").trim();

/* ---------- parsing ---------- */
const SECTION_ALIASES = {
  summary: ["professional summary", "summary", "profile", "career objective", "objective", "professional profile", "about me", "career summary"],
  experience: ["work experience", "professional experience", "employment history", "experience", "work history", "employment", "career history"],
  education: ["education", "academic background", "educational qualifications", "qualifications", "academics", "academic qualifications"],
  skills: ["technical skills", "core skills", "skills", "key skills", "technical expertise", "competencies", "core competencies", "tools and technologies", "technologies"],
  projects: ["projects", "personal projects", "academic projects", "key projects", "project experience"],
  certifications: ["certifications", "certificates", "professional certifications", "licenses", "courses", "training"]
};
const HEADING_RULES = [
  ["ignore", /personality|traits|declaration|personal (details|information|profile)|hobbies|interests|strengths?|languages known|references?|address|date of birth/],
  ["certifications", /computer qualification|certif|courses?|training|licen[sc]e/],
  ["summary", /summary|objective|profile|about me/],
  ["education", /education|qualification|academic/],
  ["experience", /experience|employment|work history|career history|organi[sz]ational|professional background/],
  ["skills", /skill|s[kl]i[kl]+s|competenc|technical|technolog|expertise|^computer/],
  ["projects", /project/]
];
const identifyHeading = line => {
  const clean = line.replace(/^[•\-*▪●◦➢➤○■□✓\s]+/, "").replace(/[:\-–|\s]+$/, "").trim();
  if (clean.length < 3 || clean.length > 45) return null;
  const n = clean.toLowerCase().replace(/[^a-z& ]/g, "").replace(/\s+/g, " ").trim();
  for (const [sec, a] of Object.entries(SECTION_ALIASES)) if (a.includes(n)) return sec;
  const isCaps = clean === clean.toUpperCase() && /[A-Z]/.test(clean);
  if (!isCaps || /[.@]/.test(clean) || clean.split(/\s+/).length > 5) return null;
  for (const [sec, rx] of HEADING_RULES) if (rx.test(n)) return sec;
  return null;
};
// Splits text where headings or bullet symbols were merged into one line.
function normalizeResumeText(t) {
  return t
    .replace(/[ \t]*([➢➤▪●◦■□✓])[ \t]*/g, "\n$1 ")
    .replace(/([^\n])[ \t]+((?:EDUCATIONAL|ACADEMIC|COMPUTER|ORGANI[SZ]ATIONAL|PROFESSIONAL|TECHNICAL|WORK|PERSONAL|CAREER)\s+(?:QUALIFICATIONS?|EXPERIENCE|SKILLS|DETAILS|SUMMARY|PROFILE|BACKGROUND|HISTORY|INFORMATION)|DECLARATION|CAREER OBJECTIVE|STRENGTHS?|HOBBIES|PROJECTS?|SKILLS|EXPERIENCE|EDUCATION)(?=\s|$)/g, "$1\n$2\n")
    .replace(/\n[ \t]*/g, "\n").replace(/\n{2,}/g, "\n");
}
function splitSections(lines) {
  const s = { header: [] };
  let active = "header";
  for (const l of lines) {
    const bar = l.split(/\s+\|\s+/);
    const hb = bar.length > 1 ? identifyHeading(bar[0]) : null;
    if (hb) {
      const prev = active;
      active = hb; s[hb] = s[hb] || [];
      const rest = bar.slice(1).join(" | ");
      (s[hb === "ignore" ? prev : hb] = s[hb === "ignore" ? prev : hb] || []).push(rest);
      continue;
    }
    const h = identifyHeading(l);
    if (h) { active = h; s[h] = s[h] || []; } else (s[active] = s[active] || []).push(l);
  }
  return s;
}

const MONTH = "(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)";
const DATE = `(?:${MONTH}\\.?,?\\s*['’]?\\d{2,4}|\\d{1,2}/\\d{4}|(?:19|20)\\d{2})`;
const RANGE = new RegExp(`(${DATE})\\s*(?:-|–|—|to)\\s*(${DATE}|present|current|till date|now|ongoing)`, "i");
const SINGLE = new RegExp(DATE, "i");
const BUL = /^[•\-*▪●◦·➢➤○■□✓]\s*/;
const ROLE_RX = /engineer|developer|analyst|manager|tester|lead|consultant|architect|designer|intern|specialist|associate|administrator|officer|executive|sdet|qa\b|scientist|director|trainee|accountant|assistant|clerk|cashier|teacher|supervisor|operator|coordinator|auditor/i;
const COMPANY_RX = /pvt|ltd|limited|inc\b|llc|corp|technolog|solutions|systems|services|software|labs|consulting|infotech|global|bank|university/i;

const extractEmail = t => t.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)?.[0] || "";
const extractPhone = t => (t.match(/(?:\+?\d{1,3}[-.\s]?)?(?:\(?\d{2,5}\)?[-.\s]?)?\d{3,5}[-.\s]?\d{4,5}/g) || [])
  .find(v => v.replace(/\D/g, "").length >= 10 && v.replace(/\D/g, "").length <= 13)?.trim() || "";
const extractLinkedIn = t => t.match(/(?:https?:\/\/)?(?:www\.)?linkedin\.com\/[^\s,|)]+/i)?.[0] || "";
const extractWebsite = t => (t.match(/(?:https?:\/\/)?(?:www\.)?[a-z0-9-]+\.(?:com|in|org|net|io|dev|me|app)(?:\/[^\s,|]*)?/gi) || [])
  .find(l => !/linkedin\.com|@|gmail|yahoo|outlook|hotmail/i.test(l) && !t.includes("@" + l)) || "";

function parseHeader(lines, fullText) {
  const head = lines.slice(0, 10);
  const name = head.find(l => l.length >= 3 && l.length < 50 && !/@|https?:|\d{4,}|resume|curriculum|linkedin/i.test(l)
    && !ROLE_RX.test(l) && /^[A-Za-z][A-Za-z .'-]+$/.test(l)) || "";
  const title = head.find(l => l !== name && l.length < 80 && ROLE_RX.test(l) && !/@|\d{5,}/.test(l)) || "";
  const segments = head.flatMap(l => l.split(/[|•·●▪]/)).map(s => s.trim());
  const location = segments.find(s => s && !/@|https?:|www\.|\d{5,}|linkedin/i.test(s) && s !== name && s !== title
    && (/\b(bengaluru|bangalore|hyderabad|pune|mumbai|delhi|chennai|noida|gurgaon|gurugram|kolkata|india)\b/i.test(s)
      || /^[A-Z][A-Za-z .]+,\s*[A-Z][A-Za-z .]+$/.test(s))) || "";
  return {
    name, title, location, email: extractEmail(fullText), phone: extractPhone(fullText),
    linkedin: extractLinkedIn(fullText), website: extractWebsite(fullText)
  };
}

function groupEntries(lines) {
  const entries = [];
  let cur = null;
  const start = () => { cur = { head: [], start: "", end: "", bullets: [] }; entries.push(cur); };
  for (const raw of lines) {
    const line = raw.trim();
    const range = line.match(RANGE);
    if (range) {
      if (!cur || cur.bullets.length || cur.start) start();
      cur.start = range[1]; cur.end = range[2];
      const rest = line.replace(range[0], "").replace(/[()|,–—\-\s]+$/, "").replace(/^[|,–—\-\s(]+/, "").trim();
      const r2 = rest.replace(/^\d+[.)]\s+/, "");
      if (cur.app && cur.head.length === 1) cur.loc = r2;
      else if (r2) (r2.length > 90 ? cur.bullets : cur.head).push(r2);
      continue;
    }
    const num = line.match(/^\d+[.)]\s+(.*)/);
    const exp = (num ? num[1] : line).match(/^(.*?)\s*[:\-–]+\s*(?:working\s+)?experience\s+of\s+(.+?)\s+years?\s*(?:as\s*(.*))?$/i);
    if (exp) {
      start();
      const bar = exp[1].split(/\s+\|\s+/);
      cur.duration = `${exp[2].trim()} years`;
      if (bar.length > 1) { cur.roleText = bar[0].trim(); cur.company = bar.slice(1).join(" | ").trim(); }
      else { cur.company = exp[1].trim(); cur.roleText = (exp[3] || "").trim(); cur.awaitRole = !cur.roleText; }
      continue;
    }
    if (num) {
      start();
      const m = num[1].match(/^(.*?)\.\s+(.*)$/);
      cur.head.push((m ? m[1] : num[1]).trim());
      if (m && m[2]) cur.bullets.push(m[2]);
      continue;
    }
    if (BUL.test(line)) { if (!cur) start(); cur.bullets.push(line.replace(BUL, "")); continue; }
    if (cur?.awaitRole) { cur.roleText = line; cur.awaitRole = false; continue; }
    const dur = line.match(/^(?:(.*?)\s*\|\s*)?(\d+(?:\.\d+)?\s*(?:years?|yrs?|months?))$/i);
    if (dur && cur) { cur.loc = (dur[1] || "").trim(); cur.duration = dur[2]; continue; }
    if (line.includes(" | ") && line.length < 110 && !/[.;]$/.test(line) && (!cur || cur.bullets.length || cur.start || cur.duration)) {
      start(); cur.app = true; cur.head.push(line); continue;
    }
    const short = line.length < 80 && !/[.;]$/.test(line);
    const last = cur?.bullets.length ? cur.bullets[cur.bullets.length - 1] : null;
    if (last !== null && (/^[a-z(]/.test(line) || (!short && !/[.;]$/.test(last)))) { cur.bullets[cur.bullets.length - 1] += " " + line; continue; }
    if (short) { if (!cur || cur.bullets.length) start(); cur.head.push(line); }
    else { if (!cur) start(); cur.bullets.push(line); }
  }
  return entries;
}

function joinWrapped(lines) {
  const out = [];
  for (const l of lines) {
    const prev = out[out.length - 1];
    if (prev && !BUL.test(l) && (/[:\-–]-?$/.test(prev) || /^[a-z]/.test(l) || /^working experience/i.test(l)) && !BUL.test(prev)) out[out.length - 1] = prev + " " + l;
    else out.push(l);
  }
  return out;
}

function parseExperience(lines) {
  return groupEntries(joinWrapped(lines)).map(e => {
    if (e.app) {
      let [role = "", ...c] = e.head[0].split(/\s+\|\s+/);
      let company = c.join(" | ");
      if ((COMPANY_RX.test(role) && !COMPANY_RX.test(company)) || (ROLE_RX.test(company) && !ROLE_RX.test(role))) [role, company] = [company, role];
      return { role: role.trim(), company: company.trim(), location: e.loc || "", start: e.duration || e.start, end: e.end, bullets: e.bullets.join("\n") };
    }
    if (e.company) {
      const cm = e.company.match(/^(.*?)\s*\((.*)\)\s*$/);
      const r = (e.roleText || "").replace(/\s+and\s+have\s+work.*$/i, "").replace(/\s+as$/i, "").replace(/[.:\-\s]+$/, "").trim();
      return { role: r, company: cm ? cm[1].trim() : e.company, location: cm ? cm[2].trim() : "", start: e.duration || "", end: "", bullets: e.bullets.join("\n") };
    }
    let [role = "", company = "", location = ""] = e.head;
    if (e.head.length === 1) {
      const parts = role.split(/\s+(?:\||@|at|-|–|—)\s+/).map(s => s.trim()).filter(Boolean);
      [role = "", company = ""] = parts;
    }
    if ((COMPANY_RX.test(role) && !COMPANY_RX.test(company)) || (ROLE_RX.test(company) && !ROLE_RX.test(role)))
      [role, company] = [company, role];
    if (role && !company && !ROLE_RX.test(role)) { company = role; role = ""; }
    if (!role) {
      const m = e.bullets.join(" ").match(/(?:\bas\b|capacity of|designation of|post of)\s+(?:an?\s+|the\s+)?([A-Z][A-Za-z]+(?:\s[A-Z][A-Za-z]+){0,2})/);
      if (m) role = m[1];
    }
    return { role, company, location, start: e.start, end: e.end, bullets: e.bullets.join("\n") };
  }).filter(e => e.role || e.company || e.bullets);
}

function parseEducation(lines) {
  const DEG = /b\.?\s?tech|b\.?\s?e\b|bachelor|master|m\.?\s?tech|m\.?\s?b\.?\s?a|b\.?\s?sc|m\.?\s?sc|b\.?\s?c\.?\s?a|m\.?\s?c\.?\s?a|diploma|ph\.?\s?d|degree|b\.?com|12th|10th|hsc|ssc|intermediate|high school|matriculation/i;
  const INST = /university|college|institute|school|academy|board|iit|nit/i;
  const groups = []; let cur = null;
  const FROM = /^(.*?)\s+from\s+(.*?)(?:\s+(?:in|during)\s+(?:the\s+)?(?:year\s+)?((?:19|20)\d{2}))?(?:\s+with\s+([\d.]+\s?%|[\d.]+\s?cgpa|[\d.]+\s?gpa))?\.?$/i;
  const done = [];
  for (const raw of lines) {
    const l = raw.replace(BUL, "").trim();
    const f = l.match(FROM);
    if (f && DEG.test(f[1])) {
      done.push({ degree: f[1].trim(), institution: f[2].trim(), location: "", year: f[3] || "", details: f[4] ? `Score: ${f[4]}` : "" });
      cur = null; continue;
    }
    const d = DEG.test(l), i = INST.test(l);
    if (!cur || (d && cur.degree) || (i && cur.institution && cur.degree)) groups.push((cur = { degree: "", institution: "", rest: [], year: "" }));
    const r = l.match(RANGE), s = l.match(SINGLE);
    if (!cur.year && (r || s)) cur.year = r ? `${r[1]} - ${r[2]}` : s[0];
    const text = l.replace(r?.[0] || s?.[0] || "", "").replace(/[|,–—\-\s]+$/, "").trim();
    if (d && !cur.degree && text.includes("|")) {
      const [dg, ...ins] = text.split(/\s*\|\s*/);
      cur.degree = dg.trim(); cur.institution = ins.join(", ").trim();
    } else if (d && !cur.degree) cur.degree = text;
    else if (i && !cur.institution) cur.institution = text;
    else if (text) cur.rest.push(text);
  }
  return [...done, ...groups.map(g => ({ degree: g.degree, institution: g.institution, location: "", year: g.year, details: g.rest.join("\n") }))]
    .filter(g => g.degree || g.institution);
}

function parseProjects(lines) {
  return groupEntries(lines.filter(l => !RANGE.test(l) || true)).map(e => {
    const body = [...e.head.slice(1), ...e.bullets];
    const techLine = body.find(l => /^(technologies|tech stack|tools|environment)\s*[:\-]/i.test(l));
    return {
      name: e.head[0] || "",
      technologies: techLine ? techLine.replace(/^[^:\-]+[:\-]\s*/, "") : "",
      description: body.filter(l => l !== techLine).join("\n")
    };
  }).filter(p => p.name || p.description);
}

function parseImportedResume(raw) {
  const text = normalizeResumeText(cleanText(raw));
  const lines = text.split("\n").map(l => l.trim()).filter(Boolean);
  const s = splitSections(lines);
  const joinSkills = ls => {
    const o = [];
    for (const l of ls) {
      const p = o[o.length - 1];
      if (p && p.includes(",") && !BUL.test(l) && !BUL.test(p) && !/^[^,:]{2,30}:/.test(l) && !/[.;]$/.test(p)) o[o.length - 1] = p + " " + l;
      else o.push(l);
    }
    return o;
  };
  const skillItems = joinSkills(s.skills || []).map(l => l.replace(BUL, "").replace(/^[^,:]{2,30}:\s*/, ""))
    .flatMap(l => l.split(/,(?![^()]*\))|[|•▪●;]/)).map(x => x.trim().replace(/^\d+[.)]\s*/, "")).filter(Boolean);
  const exp = parseExperience(s.experience || []);
  const edu = parseEducation(s.education || []);
  const proj = parseProjects(s.projects || []);
  return {
    personal: parseHeader(s.header.length ? s.header : lines, text),
    summary: (s.summary || []).map(l => l.replace(BUL, "")).join(" "),
    experience: exp.length ? exp : [blankExp()],
    education: edu.length ? edu : [blankEdu()],
    skills: [...new Set(skillItems)].join(", "),
    projects: proj.length ? proj : [blankProj()],
    certifications: (s.certifications || []).map(l => l.replace(BUL, "")).join("\n"),
    jobDescription: ""
  };
}

/* ---------- ATS ---------- */
const STOP = new Set("and the for with you our are will this that from have has into your who what how can all any not but their they about years year work working using ability strong good experience skills skill required preferred candidate team role job including such across ensure build develop must should able also more other than within".split(" "));
const tokens = v => (v || "").toLowerCase().match(/[a-z][a-z0-9+#.]{1,}/g)?.map(w => w.replace(/\.$/, "")) || [];
const VERBS = /^(developed|built|designed|led|implemented|automated|created|managed|improved|reduced|increased|delivered|executed|optimized|tested|maintained|collaborated|analyzed|migrated|integrated|deployed|established|conducted|coordinated|ensured|resolved|launched|owned|drove|mentored|validated|authored|achieved|performed|worked|prepared|identified|supported|monitored|enhanced|streamlined)/i;

function analyzeResume(d) {
  const bullets = d.experience.flatMap(e => e.bullets.split("\n").map(b => b.trim()).filter(Boolean));
  const firstExp = d.experience[0]?.bullets.split("\n").filter(Boolean) || [];
  const quant = bullets.filter(b => /\d|%/.test(b)).length;
  const verbs = bullets.filter(b => VERBS.test(b)).length;
  const p = d.personal;
  const checks = [
    ["Contact details", 15, p.name && /\S+@\S+\.\S+/.test(p.email) && p.phone, "Add your name, a valid email and a phone number."],
    ["LinkedIn or portfolio link", 5, p.linkedin || p.website, "Add a LinkedIn or portfolio URL."],
    ["Professional summary (50–600 characters)", 10, d.summary.length >= 50 && d.summary.length <= 600, "Write a 2–4 line summary tailored to the role."],
    ["Experience with title, company and dates", 15, d.experience.some(e => e.role && e.company && e.start), "Fill in job title, company and start date for each role."],
    ["Detailed bullets in latest role (3+)", 10, firstExp.length >= 3, "Add at least three achievement bullets to your latest role."],
    ["Quantified achievements (numbers, %)", 10, quant >= 2, "Add measurable results, e.g. 'reduced regression time by 40%'."],
    ["Bullets start with action verbs", 10, bullets.length > 0 && verbs / bullets.length >= 0.5, "Start bullets with verbs such as Developed, Automated, Led."],
    ["Skills section (6+ skills)", 10, d.skills.split(",").filter(s => s.trim()).length >= 6, "List at least six relevant skills, separated by commas."],
    ["Education", 10, d.education.some(e => e.degree && e.institution), "Add your degree and institution."],
    ["Projects", 5, d.projects.some(x => x.name && x.description), "Add one or two relevant projects."]
  ].map(([label, weight, ok, tip]) => ({ label, weight, ok: Boolean(ok), tip }));
  const quality = checks.reduce((s, c) => s + (c.ok ? c.weight : 0), 0);

  const jd = [...new Set(tokens(d.jobDescription).filter(w => !STOP.has(w) && w.length > 2))];
  const resumeText = [p.title, d.summary, d.skills, d.certifications,
    ...d.experience.flatMap(e => [e.role, e.company, e.bullets]),
    ...d.education.flatMap(e => [e.degree, e.details]),
    ...d.projects.flatMap(x => [x.name, x.technologies, x.description])].join(" ");
  const have = new Set(tokens(resumeText));
  const matched = jd.filter(w => have.has(w));
  const missing = jd.filter(w => !have.has(w)).slice(0, 30);
  const hasJD = jd.length > 0;
  const keywordScore = hasJD ? Math.round(matched.length / jd.length * 100) : 0;
  return {
    score: hasJD ? Math.round(keywordScore * 0.6 + quality * 0.4) : quality,
    quality, keywordScore, matched, missing, checks, hasJD
  };
}

/* ---------- export blocks (shared by preview, DOCX, PDF) ---------- */
function buildBlocks(d) {
  const out = [];
  const lines = v => (v || "").split("\n").map(x => x.replace(BUL, "").trim()).filter(Boolean);
  const join = (...a) => a.filter(Boolean).join(", ");
  const dates = (a, b) => [a, b].filter(Boolean).join(" - ");
  const add = (title, items) => { if (items.length) out.push({ title, items }); };
  if (d.summary.trim()) add("Professional Summary", [{ t: "text", v: d.summary.trim() }]);
  add("Experience", d.experience.filter(e => e.role || e.company || e.bullets.trim()).flatMap(e => [
    { t: "entry", left: e.role || e.company, right: dates(e.start, e.end), sub: e.role ? join(e.company, e.location) : e.location },
    ...lines(e.bullets).map(v => ({ t: "bullet", v }))
  ]));
  if (d.skills.trim()) add("Skills", [{ t: "text", v: d.skills.trim() }]);
  add("Projects", d.projects.filter(x => x.name || x.description.trim()).flatMap(x => [
    { t: "entry", left: x.name, right: "", sub: x.technologies },
    ...lines(x.description).map(v => ({ t: "bullet", v }))
  ]));
  add("Education", d.education.filter(e => e.degree || e.institution).flatMap(e => [
    { t: "entry", left: e.degree || e.institution, right: e.year, sub: e.degree ? join(e.institution, e.location) : e.location },
    ...lines(e.details).map(v => ({ t: "text", v }))
  ]));
  add("Certifications", lines(d.certifications).map(v => ({ t: "bullet", v })));
  return out;
}
const contactLine = p => [p.email, p.phone, p.location, p.linkedin, p.website].filter(Boolean).join("  |  ");
const fileBase = p => (p.name || "Resume").replace(/[^a-z0-9]+/gi, "_");

/* ---------- UI pieces ---------- */
function Field({ label, value, onChange, placeholder = "", area = false, hint = "", type = "text" }) {
  return (
    <label className="field">
      <span>{label}</span>
      {area
        ? <textarea rows={4} value={value || ""} onChange={e => onChange(e.target.value)} placeholder={placeholder} />
        : <input type={type} value={value || ""} onChange={e => onChange(e.target.value)} placeholder={placeholder} />}
      {hint && <small>{hint}</small>}
    </label>
  );
}

function Repeater({ title, items, onAdd, onRemove, render, addLabel }) {
  return (
    <>
      {items.map((item, i) => (
        <div className="repeat" key={i}>
          <div className="repeathead">
            <b>{title} {i + 1}</b>
            {items.length > 1 && (
              <button className="icon-btn" onClick={() => onRemove(i)} aria-label={`Remove ${title}`}><Trash2 size={15} /></button>
            )}
          </div>
          {render(item, i)}
        </div>
      ))}
      <button className="btn light" onClick={onAdd}><Plus size={15} /> {addLabel}</button>
    </>
  );
}

function Gauge100({ value, label }) {
  const color = value >= 80 ? "#16a34a" : value >= 55 ? "#d97706" : "#dc2626";
  return (
    <div className="gauge" style={{ background: `conic-gradient(${color} ${value * 3.6}deg, #e5e7eb 0deg)` }}>
      <div><b>{value}</b><small>{label}</small></div>
    </div>
  );
}

function App() {
  const [data, setData] = useState(createBlankResume);
  const [tab, setTab] = useState("personal");
  const [toast, setToast] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [fileName, setFileName] = useState("");
  const [rawText, setRawText] = useState("");
  const [ats, setAts] = useState(null);
  const [drag, setDrag] = useState(false);
  const fileRef = useRef(null);

  const notify = (msg, type = "ok") => { setToast({ msg, type }); window.setTimeout(() => setToast(null), 3500); };
  const setField = (k, v) => setData(p => ({ ...p, [k]: v }));
  const setPersonal = (k, v) => setData(p => ({ ...p, personal: { ...p.personal, [k]: v } }));
  const setRow = (s, i, k, v) => setData(p => ({ ...p, [s]: p[s].map((it, n) => n === i ? { ...it, [k]: v } : it) }));
  const addRow = (s, item) => setData(p => ({ ...p, [s]: [...p[s], item] }));
  const delRow = (s, i) => setData(p => ({ ...p, [s]: p[s].filter((_, n) => n !== i) }));
  const blocks = useMemo(() => buildBlocks(data), [data]);

  const runAts = () => { setAts(analyzeResume(data)); setTab("ats"); };

  const isDirty = () => JSON.stringify({ ...data, jobDescription: "" }) !== JSON.stringify({ ...createBlankResume(), jobDescription: "" });

  const handleFile = async file => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["pdf", "docx"].includes(ext)) return notify("Please upload a PDF or DOCX file.", "err");
    if (file.size > 10 * 1024 * 1024) return notify("Maximum file size is 10 MB.", "err");
    setUploading(true);
    try {
      const buffer = await file.arrayBuffer();
      const text = ext === "docx" ? (await mammoth.extractRawText({ arrayBuffer: buffer })).value : await readPdf(buffer);
      if (!text.trim()) throw new Error("No selectable text found. This may be a scanned PDF; please use a text-based PDF or a DOCX.");
      if (isDirty() && !window.confirm("Importing will replace the details you have entered. Continue?")) return;
      const imported = parseImportedResume(text);
      imported.jobDescription = data.jobDescription;
      setRawText(text); setFileName(file.name); setData(imported); setAts(null); setTab("personal");
      notify("Resume imported. Please review every section.");
    } catch (e) {
      console.error(e);
      notify(e.message || "Unable to read this resume.", "err");
    } finally {
      setUploading(false);
      if (fileRef.current) fileRef.current.value = "";
    }
  };

  const reset = () => {
    if (!window.confirm("Clear all resume details?")) return;
    setData(createBlankResume()); setRawText(""); setFileName(""); setAts(null); setTab("personal");
    notify("Resume cleared.");
  };

  const downloadDocx = async () => {
    try {
      const p = data.personal;
      const NAVY = "1F3A5F", GRAY = "555555", W = 9906;
      const C = AlignmentType.CENTER;
      const kids = [new Paragraph({ alignment: C, spacing: { after: 40 }, children: [new TextRun({ text: (p.name || "Your Name").toUpperCase(), bold: true, size: 36, color: NAVY })] })];
      if (p.title) kids.push(new Paragraph({ alignment: C, spacing: { after: 40 }, children: [new TextRun({ text: p.title, size: 22, color: GRAY })] }));
      kids.push(new Paragraph({ alignment: C, spacing: { after: 120 }, children: [new TextRun({ text: contactLine(p).replace(/  \|  /g, " | "), size: 19, color: GRAY })] }));
      blocks.forEach(sec => {
        kids.push(new Paragraph({
          keepNext: true, spacing: { before: 220, after: 100 },
          border: { bottom: { color: NAVY, space: 2, style: BorderStyle.SINGLE, size: 8 } },
          children: [new TextRun({ text: sec.title.toUpperCase(), bold: true, size: 22, color: NAVY })]
        }));
        sec.items.forEach(it => {
          if (it.t === "entry") {
            kids.push(new Paragraph({
              keepNext: true, spacing: { before: 100, after: it.sub ? 0 : 40 },
              tabStops: [{ type: TabStopType.RIGHT, position: W }],
              children: [new TextRun({ text: it.left, bold: true, size: 21 }),
                ...(it.right ? [new TextRun({ children: [new Tab(), it.right], size: 20, color: GRAY })] : [])]
            }));
            if (it.sub) kids.push(new Paragraph({ keepNext: true, spacing: { after: 40 }, children: [new TextRun({ text: it.sub, italics: true, size: 20, color: GRAY })] }));
          } else if (it.t === "bullet") {
            kids.push(new Paragraph({ bullet: { level: 0 }, spacing: { after: 30 }, children: [new TextRun({ text: it.v, size: 21 })] }));
          } else {
            kids.push(new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: it.v, size: 21 })] }));
          }
        });
      });
      const doc = new Document({
        styles: { default: { document: { run: { font: "Calibri", size: 21 } } } },
        sections: [{
          properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } },
          children: kids
        }]
      });
      saveAs(await Packer.toBlob(doc), `${fileBase(p)}_Resume.docx`);
      notify("Word resume downloaded.");
    } catch (e) { console.error(e); notify("Word export failed.", "err"); }
  };

  const downloadPdf = () => {
    try {
      const pdf = new jsPDF({ unit: "mm", format: "a4" });
      const L = 18, R = 192, W = R - L, NAVY = [31, 58, 95], GRAY = [85, 85, 85], BLACK = [0, 0, 0];
      const clean = v => String(v || "").replace(/[“”]/g, '"').replace(/[‘’]/g, "'").replace(/[–—]/g, "-").replace(/[•·]/g, "-").replace(/[^\x00-\xFF]/g, "");
      let y = 20;
      const lh = size => size * 0.3528 * 1.4;
      const room = h => { if (y + h > 282) { pdf.addPage(); y = 20; } };
      const font = (size, style = "normal", color = BLACK) => { pdf.setFont("helvetica", style); pdf.setFontSize(size); pdf.setTextColor(...color); };
      const para = (v, { size = 10, style = "normal", color = BLACK, align = "left" } = {}) => {
        font(size, style, color);
        for (const l of pdf.splitTextToSize(clean(v), W)) {
          room(lh(size));
          if (align === "center") pdf.text(l, (L + R) / 2, y, { align: "center" }); else pdf.text(l, L, y);
          y += lh(size);
        }
      };
      const p = data.personal;
      para((p.name || "Your Name").toUpperCase(), { size: 20, style: "bold", color: NAVY, align: "center" });
      if (p.title) para(p.title, { size: 11, color: GRAY, align: "center" });
      para(contactLine(p).replace(/  \|  /g, " | "), { size: 9, color: GRAY, align: "center" });
      blocks.forEach(sec => {
        y += 3; room(16);
        font(10.5, "bold", NAVY);
        pdf.text(clean(sec.title).toUpperCase(), L, y);
        pdf.setDrawColor(...NAVY); pdf.setLineWidth(0.35);
        pdf.line(L, y + 1.8, R, y + 1.8);
        y += 6.5;
        sec.items.forEach(it => {
          if (it.t === "entry") {
            y += 1.5; room(14);
            font(9.5, "normal", GRAY);
            const rw = it.right ? pdf.getTextWidth(clean(it.right)) : 0;
            font(10.5, "bold");
            pdf.splitTextToSize(clean(it.left), W - rw - 4).forEach((l, i) => {
              room(lh(10.5));
              font(10.5, "bold"); pdf.text(l, L, y);
              if (i === 0 && it.right) { font(9.5, "normal", GRAY); pdf.text(clean(it.right), R, y, { align: "right" }); }
              y += lh(10.5);
            });
            if (it.sub) para(it.sub, { size: 9.5, style: "italic", color: GRAY });
            y += 0.5;
          } else if (it.t === "bullet") {
            font(10);
            pdf.splitTextToSize(clean(it.v), W - 6).forEach((l, i) => {
              room(lh(10));
              if (i === 0) { pdf.setFillColor(0, 0, 0); pdf.circle(L + 1.6, y - 1, 0.5, "F"); }
              font(10); pdf.text(l, L + 5, y);
              y += lh(10);
            });
            y += 0.6;
          } else { para(it.v); y += 1; }
        });
      });
      pdf.save(`${fileBase(p)}_Resume.pdf`);
      notify("PDF resume downloaded.");
    } catch (e) { console.error(e); notify("PDF export failed.", "err"); }
  };

  const tabs = [
    ["personal", "Personal", UserRound], ["summary", "Summary", FileText],
    ["experience", "Experience", BriefcaseBusiness], ["education", "Education", GraduationCap],
    ["skills", "Skills", Code2], ["projects", "Projects", FolderGit2], ["ats", "ATS Check", Target]
  ];
  const personalFields = [
    ["Full name", "name", "Jane Doe"], ["Professional title", "title", "QA Automation Engineer"],
    ["Email", "email", "you@example.com"], ["Phone", "phone", "+91 98765 43210"],
    ["Location", "location", "Bengaluru, India"], ["LinkedIn", "linkedin", "linkedin.com/in/yourname"],
    ["Website / portfolio", "website", "yourportfolio.com"]
  ];

  return (
    <div className="shell">
      <header className="top">
        <div className="brand">
          <div className="logo"><FileText size={20} /></div>
          <div><b>ResumeForge</b><small>Professional resume builder with ATS check</small></div>
        </div>
        <div className="actions">
          <button className="btn ghost" onClick={reset}><RotateCcw size={15} /> Reset</button>
          <button className="btn accent" onClick={runAts}><Gauge size={15} /> Check ATS score</button>
          <button className="btn outline" onClick={downloadDocx}><Download size={15} /> Word</button>
          <button className="btn primary" onClick={downloadPdf}><Download size={15} /> PDF</button>
        </div>
      </header>

      <main>
        <section
          className={`upload ${drag ? "drag" : ""}`}
          onDragOver={e => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={e => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files?.[0]); }}
        >
          <div className="upload-icon"><FileUp size={22} /></div>
          <div className="upload-text">
            <b>{uploading ? "Reading your resume…" : "Upload an existing resume"}</b>
            <p>Drop a PDF or DOCX here (max 10 MB). Details fill in automatically and are processed in your browser only.
              {fileName && <> Loaded: <strong>{fileName}</strong></>}</p>
          </div>
          <input ref={fileRef} type="file" hidden accept=".pdf,.docx" onChange={e => handleFile(e.target.files?.[0])} />
          <button className="btn primary" disabled={uploading} onClick={() => fileRef.current?.click()}>Choose file</button>
          <span className="or">or fill the form below</span>
          {rawText && (
            <details className="raw">
              <summary>View extracted text</summary>
              <pre>{rawText}</pre>
            </details>
          )}
        </section>

        <div className="layout">
          <section className="editor">
            <nav>
              {tabs.map(([id, label, Icon]) => (
                <button key={id} className={tab === id ? "active" : ""} onClick={() => setTab(id)}>
                  <Icon size={15} /> {label}
                </button>
              ))}
            </nav>

            <div className="content">
              {tab === "personal" && (<>
                <h2>Personal details</h2>
                <div className="grid">
                  {personalFields.map(([label, key, ph]) => (
                    <Field key={key} label={label} placeholder={ph} value={data.personal[key]}
                      type={key === "email" ? "email" : "text"} onChange={v => setPersonal(key, v)} />
                  ))}
                </div>
              </>)}

              {tab === "summary" && (<>
                <h2>Professional summary</h2>
                <Field label="Summary" area value={data.summary} onChange={v => setField("summary", v)}
                  placeholder="2–4 lines on your experience, strengths and target role."
                  hint={`${data.summary.length} characters · aim for 150–400`} />
              </>)}

              {tab === "experience" && (<>
                <h2>Work experience</h2>
                <Repeater title="Experience" items={data.experience} addLabel="Add experience"
                  onAdd={() => addRow("experience", blankExp())} onRemove={i => delRow("experience", i)}
                  render={(item, i) => (<>
                    <div className="grid">
                      {[["Job title", "role"], ["Company", "company"], ["Location", "location"], ["Start date", "start"], ["End date", "end"]]
                        .map(([l, k]) => <Field key={k} label={l} value={item[k]} onChange={v => setRow("experience", i, k, v)} />)}
                    </div>
                    <Field label="Responsibilities & achievements (one per line)" area value={item.bullets}
                      onChange={v => setRow("experience", i, "bullets", v)} />
                  </>)} />
              </>)}

              {tab === "education" && (<>
                <h2>Education</h2>
                <Repeater title="Education" items={data.education} addLabel="Add education"
                  onAdd={() => addRow("education", blankEdu())} onRemove={i => delRow("education", i)}
                  render={(item, i) => (<>
                    <div className="grid">
                      {[["Degree / qualification", "degree"], ["Institution", "institution"], ["Location", "location"], ["Year", "year"]]
                        .map(([l, k]) => <Field key={k} label={l} value={item[k]} onChange={v => setRow("education", i, k, v)} />)}
                    </div>
                    <Field label="Additional details" area value={item.details} onChange={v => setRow("education", i, "details", v)} />
                  </>)} />
              </>)}

              {tab === "skills" && (<>
                <h2>Skills & certifications</h2>
                <Field label="Skills (comma separated)" area value={data.skills} onChange={v => setField("skills", v)}
                  placeholder="Java, Selenium, TestNG, API Testing, SQL, Git" />
                <Field label="Certifications (one per line)" area value={data.certifications} onChange={v => setField("certifications", v)} />
              </>)}

              {tab === "projects" && (<>
                <h2>Projects</h2>
                <Repeater title="Project" items={data.projects} addLabel="Add project"
                  onAdd={() => addRow("projects", blankProj())} onRemove={i => delRow("projects", i)}
                  render={(item, i) => (<>
                    <Field label="Project name" value={item.name} onChange={v => setRow("projects", i, "name", v)} />
                    <Field label="Technologies" value={item.technologies} onChange={v => setRow("projects", i, "technologies", v)} />
                    <Field label="Description (one point per line)" area value={item.description}
                      onChange={v => setRow("projects", i, "description", v)} />
                  </>)} />
              </>)}

              {tab === "ats" && (<>
                <h2>ATS score checker</h2>
                <p className="sub">Works for uploaded or manually built resumes. Paste a job description for a role-specific match, or run it without one for a general ATS readiness score.</p>
                <Field label="Target job description (optional)" area value={data.jobDescription}
                  onChange={v => setField("jobDescription", v)} placeholder="Paste the complete job description…" />
                <button className="btn accent big" onClick={() => setAts(analyzeResume(data))}>
                  <Gauge size={16} /> {ats ? "Re-run ATS check" : "Run ATS check"}
                </button>

                {ats && (<div className="ats-result">
                  <div className="scorebox">
                    <Gauge100 value={ats.score} label="ATS score" />
                    <div>
                      <h3>{ats.score >= 80 ? "Strong ATS readiness" : ats.score >= 55 ? "Moderate — room to improve" : "Needs improvement"}</h3>
                      <p>Resume quality: <b>{ats.quality}/100</b>{ats.hasJD && <> · Keyword match: <b>{ats.keywordScore}%</b></>}</p>
                      <small>Rule-based estimate, not an official ATS vendor score.</small>
                    </div>
                  </div>

                  <h3>Checklist</h3>
                  <div className="checks">
                    {ats.checks.map(c => (
                      <div key={c.label} className={c.ok ? "ok" : "warn"}>
                        {c.ok ? <CheckCircle2 size={17} /> : <AlertCircle size={17} />}
                        <div><b>{c.label}</b>{!c.ok && <small>{c.tip}</small>}</div>
                      </div>
                    ))}
                  </div>

                  {ats.hasJD ? (<>
                    <h3>Matched keywords ({ats.matched.length})</h3>
                    <div className="chips">{ats.matched.slice(0, 40).map(w => <span className="good" key={w}>{w}</span>)}</div>
                    <h3>Missing keywords</h3>
                    <div className="chips">{ats.missing.map(w => <span className="miss" key={w}>{w}</span>)}</div>
                  </>) : <p className="sub">Add a job description above to see keyword matches.</p>}
                </div>)}
              </>)}
            </div>
          </section>

          <aside className="preview">
            <div className="previewhead"><b>Live preview</b><span>ATS-friendly single column</span></div>
            <article className="paper">
              <h1>{data.personal.name || "Your Name"}</h1>
              {data.personal.title && <div className="jobtitle">{data.personal.title}</div>}
              <div className="contact">{contactLine(data.personal).replace(/  \|  /g, " • ") || "Email • Phone • Location"}</div>
              {blocks.map(s => (
                <section key={s.title}>
                  <h3>{s.title.toUpperCase()}</h3>
                  {s.items.map((it, i) =>
                    it.t === "entry" ? (
                      <div key={i} className="p-entry">
                        <div className="p-row"><b>{it.left}</b>{it.right && <span>{it.right}</span>}</div>
                        {it.sub && <small className="p-sub">{it.sub}</small>}
                      </div>
                    ) : <p key={i} className={it.t === "bullet" ? "p-bullet" : ""}>{it.v}</p>)}
                </section>
              ))}
            </article>
          </aside>
        </div>
      </main>

      {toast && (
        <div className={`toast ${toast.type}`}>
          {toast.type === "err" ? <AlertCircle size={16} /> : <CheckCircle2 size={16} />} {toast.msg}
          <button onClick={() => setToast(null)} aria-label="Dismiss"><X size={14} /></button>
        </div>
      )}
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
