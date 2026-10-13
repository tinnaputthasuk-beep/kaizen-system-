import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  LayoutDashboard, User, ShieldCheck, Building2, PlusCircle, Eye, Pencil,
  CheckCircle2, XCircle, Upload, X, Filter, ClipboardList, Presentation,
  Users, TrendingUp, ChevronRight, ImageOff, LogOut, Lock, KeyRound, AlertTriangle, Clock,
  Download, Coins, Bell, BellRing, Undo2, Paperclip, FileVideo, FileSpreadsheet, FileText, ZoomIn,
  Trophy, Medal, Plus, Trash2
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell,
  CartesianGrid, LineChart, Line, Legend
} from "recharts";
import * as XLSX from "xlsx";

// ---------- Design tokens ----------
const COMPANY_LOGO = "/logo.png";
// โลโก้สำรอง: ใช้อัตโนมัติเมื่อหาไฟล์ /logo.png ไม่เจอ
const COMPANY_LOGO_FALLBACK = "data:image/png;base64,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";
const useLogoFallback = (e) => {
  if (e.currentTarget.src !== COMPANY_LOGO_FALLBACK) e.currentTarget.src = COMPANY_LOGO_FALLBACK;
};

const NAVY = "#12294B";
const NAVY_DEEP = "#0B1D33";
const GOLD = "#D99A2B";
const GREEN = "#2F7D5C";
const RED = "#B23A2E";
const SLATE = "#F4F6F8";
const INK = "#1C2531";
const SUB = "#5B6675";

const CATEGORY_OPTIONS = ["Food Safety", "Delivery", "Environment", "Quality", "Safety", "Energy", "Cost", "Moral", "Satisfaction", "Other"];
const ACTIVITY_OPTIONS = ["Kaizen", "Suggestion"];
const KAIZEN_TYPE_OPTIONS = ["Kaizen Project", "Cross functional Kaizen"];
// ---------- Google Sheets Backend (Google Apps Script Web App) ----------
// หมายเหตุ: ทำงานได้จริงเมื่อ deploy โค้ดนี้เป็นเว็บไซต์จริงนอก claude.ai เท่านั้น
// เพราะ Artifact preview ใน claude.ai บล็อกการเชื่อมต่อเซิร์ฟเวอร์ภายนอกด้วยเหตุผลด้านความปลอดภัย
const GOOGLE_SHEETS_API_URL = "https://script.google.com/macros/s/AKfycbxBiFUJVBLnkJWxhazPUTHeu0J4GS_9MLWUzOk7OubQj4vdkKbLldA3JY5UOlZeTPnl/exec";

async function fetchKaizensFromSheet() {
  const res = await fetch(GOOGLE_SHEETS_API_URL);
  if (!res.ok) throw new Error("โหลดข้อมูลจาก Google Sheet ไม่สำเร็จ");
  return res.json();
}

async function syncKaizenToSheet(action, record) {
  try {
    await fetch(GOOGLE_SHEETS_API_URL, {
      method: "POST",
      // ใช้ text/plain เพื่อเลี่ยงปัญหา CORS preflight ของ Google Apps Script (ข้อจำกัดที่รู้กันของ Apps Script)
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action, record }),
    });
  } catch (err) {
    // คาดว่าจะ error ถ้าทดสอบอยู่ใน claude.ai artifact (ถูกบล็อกโดย sandbox)
    // แต่จะทำงานได้ปกติเมื่อ deploy เป็นเว็บไซต์จริงแล้ว
    console.warn("Sync to Google Sheet failed:", err.message);
  }
}

// ---------- Kaizen Award: บันทึก/โหลดเกณฑ์และคะแนน ----------
// ทำงานเมื่อ Code.gs มีส่วนเสริม Award แล้ว ถ้ายังไม่มี ระบบจะทำงานแบบโหมดทดลอง (เก็บในหน้าเว็บชั่วคราว)
async function fetchAwardFromSheet() {
  const res = await fetch(`${GOOGLE_SHEETS_API_URL}?type=award`);
  if (!res.ok) throw new Error("load award failed");
  const json = await res.json();
  if (json && !Array.isArray(json) && Object.prototype.hasOwnProperty.call(json, "award")) {
    return { supported: true, award: json.award };
  }
  return { supported: false, award: null };
}

async function postAward(action, record) {
  try {
    await fetch(GOOGLE_SHEETS_API_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action, record }),
    });
  } catch (err) {
    console.warn("Save award failed:", err.message);
  }
}
// บันทึกเฉพาะ "เกณฑ์/กรรมการ/ผลประกาศ" (ไม่รวมคะแนน) เพื่อไม่ให้ไปทับคะแนนของกรรมการ
function saveAwardToSheet(award) {
  const { scores, ...config } = award;
  return postAward("saveAward", { id: "award", data: config });
}
// กรรมการแต่ละท่านบันทึกคะแนนของตัวเองแยกแถว จึงให้คะแนนพร้อมกันหลายคนได้โดยไม่ทับกัน
function saveJudgeScoresToSheet(judgeId, judgeScores) {
  return postAward("saveJudgeScores", { id: judgeId, judgeId, scores: judgeScores });
}

const DEFAULT_DEPARTMENTS = [
  "Production Planning & Data",
  "Washing - Drying 1",
  "Thread forming 1",
  "Packing 1",
  "Drying 2",
  "Packing 2",
  "RTE",
  "Thread forming 3",
  "Packing 3",
  "Thread forming _ RN",
  "Thread forming _ RV",
  "Cutting",
  "Drying 4",
  "Packing 4",
  "BS",
  "QC",
  "Lab",
  "QA",
  "Pest Control",
  "DC",
  "Sanitation",
  "RD",
  "Engineering Data",
  "Engineering Team A",
  "Engineering Team B",
  "Engineering Team C",
  "Boiler",
  "Water supply",
  "Finish Goods W/H (Dry)",
  "Finish Goods W/H (Fresh)",
  "Raw Material W/H",
  "Packaging W/H",
  "Spart part - Eguipment W/H",
  "Safety",
  "Envi",
  "C & B",
  "Recruitment & Training",
  "ER & Admin",
  "Accounting",
  "Procurement",
];

// ---------- Mock directory (สาธิตเท่านั้น) ----------
// อัปเดตจากไฟล์ BL_Kaizen_Person_Incharge.xlsx เวอร์ชันล่าสุด
// Username = ชื่อผู้ใช้/ชื่อ-นามสกุล, Password = รหัสพนักงานตามไฟล์
// พนักงาน 1 บัญชี อาจเลือกได้หลายหน่วยงานย่อยตามที่ไฟล์กำหนด (เช่น QC เลือกได้ทั้ง QC และ Lab)
// ในระบบจริง ห้ามตรวจรหัสผ่านฝั่ง client แบบนี้เด็ดขาด ต้องยืนยันตัวตนผ่าน backend/SSO
const MOCK_USERS = [
  { name: "Production Planning & Data", empId: "1001", role: "employee", departments: ["Production Planning & Data"] },
  { name: "Production_Line 1", empId: "1002", role: "employee", departments: ["Washing - Drying 1", "Thread forming 1", "Packing 1"] },
  { name: "Production Line 2", empId: "1003", role: "employee", departments: ["Drying 2", "Packing 2", "RTE"] },
  { name: "Production Line 3", empId: "1004", role: "employee", departments: ["Thread forming 3", "Packing 3"] },
  { name: "Production Line 4", empId: "1005", role: "employee", departments: ["Thread forming _ RN", "Thread forming _ RV", "Cutting", "Drying 4", "Packing 4", "BS"] },
  { name: "QC", empId: "1006", role: "employee", departments: ["QC", "Lab"] },
  { name: "QA", empId: "1007", role: "employee", departments: ["QA", "Pest Control", "DC", "Sanitation"] },
  { name: "RD", empId: "1008", role: "employee", departments: ["RD"] },
  { name: "Engineering Data", empId: "1009", role: "employee", departments: ["Engineering Data"] },
  { name: "Engineering Team A", empId: "1010", role: "employee", departments: ["Engineering Team A"] },
  { name: "Engineering Team B", empId: "1011", role: "employee", departments: ["Engineering Team B"] },
  { name: "Engineering Team C", empId: "1012", role: "employee", departments: ["Engineering Team C"] },
  { name: "Boiler", empId: "1013", role: "employee", departments: ["Boiler"] },
  { name: "Water supply", empId: "1014", role: "employee", departments: ["Water supply"] },
  { name: "Finish Goods W/H (Dry)", empId: "1015", role: "employee", departments: ["Finish Goods W/H (Dry)"] },
  { name: "Finish Goods W/H (Fresh)", empId: "1016", role: "employee", departments: ["Finish Goods W/H (Fresh)"] },
  { name: "Raw Material W/H", empId: "1017", role: "employee", departments: ["Raw Material W/H"] },
  { name: "Packaging W/H", empId: "1018", role: "employee", departments: ["Packaging W/H"] },
  { name: "Spart part - Eguipment W/H", empId: "1019", role: "employee", departments: ["Spart part - Eguipment W/H"] },
  { name: "Safety", empId: "1020", role: "employee", departments: ["Safety"] },
  { name: "Envi", empId: "1021", role: "employee", departments: ["Envi"] },
  { name: "C & B", empId: "1022", role: "employee", departments: ["C & B"] },
  { name: "Recruitment & Training", empId: "1023", role: "employee", departments: ["Recruitment & Training"] },
  { name: "ER & Admin", empId: "1024", role: "employee", departments: ["ER & Admin"] },
  { name: "Admin", empId: "0001", role: "admin", departments: [] },
  { name: "ผอ.โรงงาน", empId: "0002", role: "director", departments: [] },
  // บัญชีคณะกรรมการ Kaizen Award (รหัสผ่านแยกกันคนละชุด เปลี่ยนได้ที่นี่)
  { name: "กรรมการ 1", empId: "3001", role: "judge", departments: [] },
  { name: "กรรมการ 2", empId: "3002", role: "judge", departments: [] },
  { name: "กรรมการ 3", empId: "3003", role: "judge", departments: [] },
  { name: "กรรมการ 4", empId: "3004", role: "judge", departments: [] },
  { name: "กรรมการ 5", empId: "3005", role: "judge", departments: [] },

  { name: "ทิพวรรณ บุญจิตติ", empId: "2001", role: "supervisor", departments: ["Production Planning & Data"] },
  { name: "ธัญญธรณ์ สุขุมาลพันธ์", empId: "2002", role: "supervisor", departments: ["Washing - Drying 1", "Thread forming 1", "Packing 1", "Thread forming 3", "Packing 3"] },
  { name: "ธันวา ปักษาสุข", empId: "2005", role: "supervisor", departments: ["Drying 2", "Packing 2", "RTE"] },
  { name: "ศศิธร สระทองพูล", empId: "2010", role: "supervisor", departments: ["Thread forming _ RN", "Thread forming _ RV", "Cutting", "Drying 4", "Packing 4", "BS"] },
  { name: "อมรรัตน์ ตุลาผล", empId: "2016", role: "supervisor", departments: ["QC", "QA", "Pest Control", "DC", "Sanitation", "Lab"] },
  { name: "สุพิชชา วังถนอมศักดิ์", empId: "2021", role: "supervisor", departments: ["RD"] },
  { name: "วินัย บุญมุนัส", empId: "2023", role: "supervisor", departments: ["Engineering Data", "Engineering Team A", "Engineering Team B", "Engineering Team C", "Boiler", "Water supply"] },
  { name: "นันทวรรณ นวลเนตร", empId: "2029", role: "supervisor", departments: ["Finish Goods W/H (Dry)", "Finish Goods W/H (Fresh)", "Raw Material W/H", "Packaging W/H", "Spart part - Eguipment W/H"] },
  { name: "วิไลวรรณ ขำขม", empId: "2034", role: "supervisor", departments: ["Safety", "Envi"] },
  { name: "ติณณา พุทธสุข", empId: "2036", role: "supervisor", departments: ["C & B", "Recruitment & Training", "ER & Admin"] },
];

const ROLE_LABEL = { employee: "พนักงาน", supervisor: "หัวหน้างาน", admin: "Admin (ฝ่ายบุคคล)", director: "ผู้อำนวยการโรงงาน", judge: "กรรมการตัดสิน Kaizen Award" };

function buildInitialNotifications(kaizenList) {
  const list = [];
  const push = (recipientName, message, kaizenId) => {
    if (!recipientName) return;
    list.push({ id: uid(), recipientName, message, kaizenId, read: false, date: nowStamp() });
  };
  kaizenList.forEach((k) => {
    if (k.status === "pending_supervisor") {
      const sup = MOCK_USERS.find((u) => u.role === "supervisor" && u.departments.includes(k.department));
      if (sup) push(sup.name, `มีเรื่อง Kaizen ใหม่รอคุณอนุมัติ: "${k.title}"`, k.id);
    } else if (k.status === "pending_admin") {
      const admin = MOCK_USERS.find((u) => u.role === "admin");
      if (admin) push(admin.name, `มีเรื่อง Kaizen รอ Admin ตรวจสอบ: "${k.title}"`, k.id);
    } else if (k.status === "pending_director") {
      const director = MOCK_USERS.find((u) => u.role === "director");
      if (director) push(director.name, `มีเรื่อง Kaizen รอผู้อำนวยการอนุมัติขั้นสุดท้าย: "${k.title}"`, k.id);
    }
  });
  return list;
}

// ---------- Section/ฝ่าย grouping (อ้างอิงคอลัมน์ Section/ฝ่าย จากไฟล์ BL_Kaizen_Person_Incharge.xlsx โดยตรง) ----------
const SECTION_OPTIONS = [
  "Production Planning & Data",
  "Production 1,3",
  "Production 2,4",
  "Quality & RD",
  "Engineering & Utility",
  "Warehouse",
  "SHE",
  "HRBP",
  "Support Function",
];
const SECTION_MAP = {
  "Production Planning & Data": "Production Planning & Data",
  "Washing - Drying 1": "Production 1,3",
  "Thread forming 1": "Production 1,3",
  "Packing 1": "Production 1,3",
  "Drying 2": "Production 2,4",
  "Packing 2": "Production 2,4",
  "RTE": "Production 2,4",
  "Thread forming 3": "Production 1,3",
  "Packing 3": "Production 1,3",
  "Thread forming _ RN": "Production 2,4",
  "Thread forming _ RV": "Production 2,4",
  "Cutting": "Production 2,4",
  "Drying 4": "Production 2,4",
  "Packing 4": "Production 2,4",
  "BS": "Production 2,4",
  "QC": "Quality & RD",
  "QA": "Quality & RD",
  "Pest Control": "Quality & RD",
  "DC": "Quality & RD",
  "Sanitation": "Quality & RD",
  "RD": "Quality & RD",
  "Lab": "Quality & RD",
  "Engineering Data": "Engineering & Utility",
  "Engineering Team A": "Engineering & Utility",
  "Engineering Team B": "Engineering & Utility",
  "Engineering Team C": "Engineering & Utility",
  "Boiler": "Engineering & Utility",
  "Water supply": "Engineering & Utility",
  "Finish Goods W/H (Dry)": "Warehouse",
  "Finish Goods W/H (Fresh)": "Warehouse",
  "Raw Material W/H": "Warehouse",
  "Packaging W/H": "Warehouse",
  "Spart part - Eguipment W/H": "Warehouse",
  "Safety": "SHE",
  "Envi": "SHE",
  "C & B": "HRBP",
  "Recruitment & Training": "HRBP",
  "ER & Admin": "HRBP",
  "Accounting": "Support Function",
  "Procurement": "Support Function",
};

const STATUS_META = {
  draft: { label: "แบบร่าง", color: "#8A93A3", bg: "#EEF0F3" },
  pending_supervisor: { label: "รอหัวหน้างานอนุมัติ", color: "#B8791F", bg: "#FBF0DD" },
  pending_admin: { label: "รอ Admin ตรวจสอบ", color: "#2B6CB0", bg: "#E7F0FA" },
  pending_director: { label: "รอผู้อำนวยการอนุมัติ", color: "#7952B3", bg: "#EFE8F8" },
  approved: { label: "อนุมัติแล้ว", color: GREEN, bg: "#E5F3EC" },
  rejected: { label: "ถูกตีกลับ", color: RED, bg: "#F9E9E7" },
};

const PIE_COLORS = ["#8A93A3", "#D99A2B", "#2B6CB0", "#7952B3", "#2F7D5C", "#B23A2E"];

const uid = () => Math.random().toString(36).slice(2, 9);
const todayISO = () => new Date().toISOString().slice(0, 10);
const nowStamp = () => {
  const d = new Date();
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  return `${todayISO()} ${hh}:${mm}`;
};
const emptyMembers = () => Array.from({ length: 10 }, () => ({ name: "", empId: "", dept: "" }));
const fillMembers = (names) => emptyMembers().map((m, i) => (names[i] ? { ...m, name: names[i] } : m));

const CHECKLIST_ITEMS = [
  "หัวข้อที่นำเสนอไม่ซ้ำกับหัวข้อเรื่องที่เคยทำมาแล้วในอดีตก่อนหน้านี้",
  "เป็นการปรับปรุงที่ไม่ได้ซ้ำซ้อนกับหน้าที่หลัก ซึ่งพึงต้องปฏิบัติตามหน้าที่งานปกติที่ทำอยู่",
  "มีการชี้บ่งและระบุตัวปัญหา และมีการกำหนดแนวทางในการปรับปรุงแก้ไขได้อย่างชัดเจน",
  "มีความเป็นไปได้ที่จะบรรลุวัตถุประสงค์และเป้าหมายในการปรับปรุงและแก้ไขปัญหา",
];

function seedData() {
  const mk = (o) => ({
    id: uid(),
    members: [],
    budget: "",
    result: "",
    categories: [],
    activityType: "Kaizen",
    kaizenType: "Kaizen Project",
    images: { before: null, after: null },
    history: [],
    ...o,
  });
  return [
    mk({
      registeredNo: "KZ-2026-014",
      date: "2026-08-02",
      title: "รถเข็นเครื่องชั่งดิจิตอลแบบตั้งโต๊ะ",
      team: "ทีม Safety Packing 1",
      department: "Packing 1",
      submittedBy: "Production_Line 1",
      problem: "เครื่องชั่งดิจิตอลจัดเก็บอยู่ห่างจากพื้นที่ปฏิบัติงาน พนักงานต้องยกด้วยแรงคนทุกครั้งที่ใช้งาน เสี่ยงต่อการบาดเจ็บและเครื่องชั่งชำรุด",
      improvement: "ออกแบบและสร้างรถเข็นสำหรับเครื่องชั่งดิจิตอลแบบตั้งโต๊ะ ติดตั้งล้อเลื่อน ฐานรองพอดีตัวเครื่อง มีตัวล็อกกันเคลื่อนขณะใช้งาน",
      budget: "2500",
      result: "ลดความเสี่ยงการบาดเจ็บ 100%, ลดเวลาการเตรียมเครื่องชั่งลง 83%, ความพึงพอใจพนักงานเพิ่มขึ้น 20%",
      costSaving: "18000",
      costSavingCalc: "ลดเวลาเตรียมงาน 10 นาที/ครั้ง x 3 ครั้ง/วัน x 300 วัน x ค่าแรงเฉลี่ย 20 บาท/10นาที",
      categories: ["Safety"],
      activityType: "Kaizen",
      kaizenType: "Kaizen Project",
      images: { before: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNGREYzRTYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjRjBCODYwIi8+Cjx0ZXh0IHg9IjIwMCIgeT0iMTEzIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjUyIiBmb250LXdlaWdodD0iYm9sZCIgZmlsbD0iIzhBNUIxNCIgdGV4dC1hbmNob3I9Im1pZGRsZSI+ITwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guIHguYjguK3guJnguJvguKPguLHguJrguJvguKPguLjguIcgKEJlZm9yZSk8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjAwIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjEzIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj5QYWNraW5nIDE8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjIyIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjExIiBmaWxsPSIjQjA4QjRBIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKDguLLguJ7guJXguLHguKfguK3guKLguYjguLLguIfguJvguKPguLDguIHguK3guJogKFBsYWNlaG9sZGVyKTwvdGV4dD4KPC9zdmc+", after: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNFOUY1RUYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjNkZCRTk3Ii8+CjxwYXRoIGQ9Ik0xODAgOTcgbDE0IDE0IGwyOCAtMzIiIHN0cm9rZT0iIzFGNUM0MCIgc3Ryb2tlLXdpZHRoPSI5IiBmaWxsPSJub25lIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjMUY1QzQwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKvguKXguLHguIfguJvguKPguLHguJrguJvguKPguLjguIcgKEFmdGVyKTwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIyMDAiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTMiIGZpbGw9IiMxRjVDNDAiIHRleHQtYW5jaG9yPSJtaWRkbGUiPlBhY2tpbmcgMTwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIyMjIiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTEiIGZpbGw9IiM0RTk5NzMiIHRleHQtYW5jaG9yPSJtaWRkbGUiPuC4oOC4suC4nuC4leC4seC4p+C4reC4ouC5iOC4suC4h+C4m+C4o+C4sOC4geC4reC4miAoUGxhY2Vob2xkZXIpPC90ZXh0Pgo8L3N2Zz4=" },
      status: "approved",
      history: [
        { stage: "ส่งเรื่อง", actor: "Production_Line 1", action: "submit", comment: "", date: "2026-08-02 09:10" },
        { stage: "หัวหน้างาน", actor: "ธัญญธรณ์ สุขุมาลพันธ์", action: "approve", comment: "เห็นด้วย ช่วยลดความเสี่ยงจริง", date: "2026-08-04 10:45" },
        { stage: "Admin", actor: "Admin", action: "approve", comment: "จัดหมวด Safety ตรวจครบถ้วน", date: "2026-08-06 13:20" },
        { stage: "ผู้อำนวยการ", actor: "ผอ.โรงงาน", action: "approve", comment: "อนุมัติ ให้ขยายผลไปไลน์อื่น", date: "2026-08-10 14:05" },
      ],
    }),
    mk({
      registeredNo: "KZ-2026-015",
      date: "2026-08-20",
      title: "ลดเวลาการเปลี่ยนไส้กรองน้ำมันเครื่องจักร",
      team: "ทีม Engineering A",
      department: "Engineering Team A",
      submittedBy: "Engineering Team A",
      problem: "การเปลี่ยนไส้กรองใช้เวลานานเพราะเครื่องมือกระจัดกระจาย ต้องเดินหาอุปกรณ์หลายจุด",
      improvement: "จัดทำชุดเครื่องมือเฉพาะ (Shadow board) ไว้ใกล้จุดปฏิบัติงาน",
      budget: "1800",
      result: "ลดเวลาเปลี่ยนไส้กรองจาก 45 นาที เหลือ 20 นาที",
      costSaving: "9500",
      costSavingCalc: "ลดเวลาหยุดเครื่อง 25 นาที/ครั้ง x เฉลี่ย 2 ครั้ง/เดือน x ต้นทุนเครื่องจักรหยุด 190 บาท/นาที",
      categories: ["Cost", "Delivery"],
      kaizenType: "Cross functional Kaizen",
      images: { before: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNGREYzRTYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjRjBCODYwIi8+Cjx0ZXh0IHg9IjIwMCIgeT0iMTEzIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjUyIiBmb250LXdlaWdodD0iYm9sZCIgZmlsbD0iIzhBNUIxNCIgdGV4dC1hbmNob3I9Im1pZGRsZSI+ITwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guIHguYjguK3guJnguJvguKPguLHguJrguJvguKPguLjguIcgKEJlZm9yZSk8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjAwIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjEzIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj5FbmdpbmVlcmluZyBUZWFtIEE8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjIyIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjExIiBmaWxsPSIjQjA4QjRBIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKDguLLguJ7guJXguLHguKfguK3guKLguYjguLLguIfguJvguKPguLDguIHguK3guJogKFBsYWNlaG9sZGVyKTwvdGV4dD4KPC9zdmc+", after: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNFOUY1RUYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjNkZCRTk3Ii8+CjxwYXRoIGQ9Ik0xODAgOTcgbDE0IDE0IGwyOCAtMzIiIHN0cm9rZT0iIzFGNUM0MCIgc3Ryb2tlLXdpZHRoPSI5IiBmaWxsPSJub25lIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjMUY1QzQwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKvguKXguLHguIfguJvguKPguLHguJrguJvguKPguLjguIcgKEFmdGVyKTwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIyMDAiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTMiIGZpbGw9IiMxRjVDNDAiIHRleHQtYW5jaG9yPSJtaWRkbGUiPkVuZ2luZWVyaW5nIFRlYW0gQTwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIyMjIiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTEiIGZpbGw9IiM0RTk5NzMiIHRleHQtYW5jaG9yPSJtaWRkbGUiPuC4oOC4suC4nuC4leC4seC4p+C4reC4ouC5iOC4suC4h+C4m+C4o+C4sOC4geC4reC4miAoUGxhY2Vob2xkZXIpPC90ZXh0Pgo8L3N2Zz4=" },
      status: "pending_director",
      history: [
        { stage: "ส่งเรื่อง", actor: "Engineering Team A", action: "submit", comment: "", date: "2026-08-20 15:30" },
        { stage: "หัวหน้างาน", actor: "วินัย บุญมุนัส", action: "approve", comment: "แนวทางชัดเจน อนุมัติ", date: "2026-08-21 16:15" },
        { stage: "Admin", actor: "Admin", action: "approve", comment: "จัดหมวด Cost, Delivery", date: "2026-08-25 11:00" },
      ],
    }),
    mk({
      registeredNo: "KZ-2026-016",
      date: "2026-08-27",
      title: "ป้ายชี้บ่งสีพื้นทางเดินในคลังสินค้า",
      team: "ทีม Raw Material W/H",
      department: "Raw Material W/H",
      submittedBy: "Raw Material W/H",
      problem: "ทางเดินและพื้นที่จัดเก็บไม่มีเส้นแบ่งชัดเจน เสี่ยงรถโฟล์คลิฟท์ชนพนักงาน",
      improvement: "ทาสีตีเส้นแบ่งทางเดิน-พื้นที่จัดเก็บ พร้อมป้ายบ่งชี้ความเร็ว",
      budget: "3200",
      result: "ยังไม่มีข้อมูล (อยู่ระหว่างตรวจสอบ)",
      costSaving: "",
      categories: ["Safety", "Environment"],
      images: { before: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNGREYzRTYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjRjBCODYwIi8+Cjx0ZXh0IHg9IjIwMCIgeT0iMTEzIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjUyIiBmb250LXdlaWdodD0iYm9sZCIgZmlsbD0iIzhBNUIxNCIgdGV4dC1hbmNob3I9Im1pZGRsZSI+ITwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guIHguYjguK3guJnguJvguKPguLHguJrguJvguKPguLjguIcgKEJlZm9yZSk8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjAwIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjEzIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj5SYXcgTWF0ZXJpYWwgVy9IPC90ZXh0Pgo8dGV4dCB4PSIyMDAiIHk9IjIyMiIgZm9udC1mYW1pbHk9IlRhaG9tYSwgc2Fucy1zZXJpZiIgZm9udC1zaXplPSIxMSIgZmlsbD0iI0IwOEI0QSIgdGV4dC1hbmNob3I9Im1pZGRsZSI+4Lig4Liy4Lie4LiV4Lix4Lin4Lit4Lii4LmI4Liy4LiH4Lib4Lij4Liw4LiB4Lit4LiaIChQbGFjZWhvbGRlcik8L3RleHQ+Cjwvc3ZnPg==", after: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNFOUY1RUYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjNkZCRTk3Ii8+CjxwYXRoIGQ9Ik0xODAgOTcgbDE0IDE0IGwyOCAtMzIiIHN0cm9rZT0iIzFGNUM0MCIgc3Ryb2tlLXdpZHRoPSI5IiBmaWxsPSJub25lIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjMUY1QzQwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKvguKXguLHguIfguJvguKPguLHguJrguJvguKPguLjguIcgKEFmdGVyKTwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIyMDAiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTMiIGZpbGw9IiMxRjVDNDAiIHRleHQtYW5jaG9yPSJtaWRkbGUiPlJhdyBNYXRlcmlhbCBXL0g8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjIyIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjExIiBmaWxsPSIjNEU5OTczIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKDguLLguJ7guJXguLHguKfguK3guKLguYjguLLguIfguJvguKPguLDguIHguK3guJogKFBsYWNlaG9sZGVyKTwvdGV4dD4KPC9zdmc+" },
      status: "pending_admin",
      attachments: [
        {
          id: "demo-att-1",
          name: "แผนผังจุดติดตั้งป้ายชี้บ่ง.txt",
          size: 512,
          type: "text/plain",
          dataUrl: "data:text/plain;charset=utf-8;base64,4LmE4Lif4Lil4LmM4LiV4Lix4Lin4Lit4Lii4LmI4Liy4LiH4Liq4Liz4Lir4Lij4Lix4Lia4Liq4Liy4LiY4Li04LiV4Lij4Liw4Lia4LiaIEthaXplbiBTeXN0ZW0gLSDguYHguJzguJnguJzguLHguIfguIjguLjguJTguJXguLTguJTguJXguLHguYnguIfguJvguYnguLLguKLguIrguLXguYnguJrguYjguIfguYPguJnguITguKXguLHguIfguKrguLTguJnguITguYnguLI=",
        },
      ],
      history: [
        { stage: "ส่งเรื่อง", actor: "Raw Material W/H", action: "submit", comment: "", date: "2026-08-27 08:50" },
        { stage: "หัวหน้างาน", actor: "นันทวรรณ นวลเนตร", action: "approve", comment: "เร่งด่วนด้านความปลอดภัย", date: "2026-08-28 09:35" },
      ],
    }),
    mk({
      registeredNo: "KZ-2026-017",
      date: "2026-09-01",
      title: "แบบฟอร์มตรวจรับวัตถุดิบแบบดิจิทัล",
      team: "ทีม QA",
      department: "QA",
      submittedBy: "QA",
      problem: "แบบฟอร์มกระดาษเสียเวลาคีย์ข้อมูลซ้ำ และค้นหาย้อนหลังยาก",
      improvement: "ใช้แท็บเล็ตกรอกแบบฟอร์มตรวจรับแทนกระดาษ เชื่อมข้อมูลอัตโนมัติ",
      budget: "0",
      result: "",
      categories: ["Quality"],
      kaizenType: "Cross functional Kaizen",
      images: { before: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNGREYzRTYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjRjBCODYwIi8+Cjx0ZXh0IHg9IjIwMCIgeT0iMTEzIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjUyIiBmb250LXdlaWdodD0iYm9sZCIgZmlsbD0iIzhBNUIxNCIgdGV4dC1hbmNob3I9Im1pZGRsZSI+ITwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guIHguYjguK3guJnguJvguKPguLHguJrguJvguKPguLjguIcgKEJlZm9yZSk8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjAwIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjEzIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj5RQTwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIyMjIiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTEiIGZpbGw9IiNCMDhCNEEiIHRleHQtYW5jaG9yPSJtaWRkbGUiPuC4oOC4suC4nuC4leC4seC4p+C4reC4ouC5iOC4suC4h+C4m+C4o+C4sOC4geC4reC4miAoUGxhY2Vob2xkZXIpPC90ZXh0Pgo8L3N2Zz4=", after: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNFOUY1RUYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjNkZCRTk3Ii8+CjxwYXRoIGQ9Ik0xODAgOTcgbDE0IDE0IGwyOCAtMzIiIHN0cm9rZT0iIzFGNUM0MCIgc3Ryb2tlLXdpZHRoPSI5IiBmaWxsPSJub25lIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjMUY1QzQwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKvguKXguLHguIfguJvguKPguLHguJrguJvguKPguLjguIcgKEFmdGVyKTwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIyMDAiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTMiIGZpbGw9IiMxRjVDNDAiIHRleHQtYW5jaG9yPSJtaWRkbGUiPlFBPC90ZXh0Pgo8dGV4dCB4PSIyMDAiIHk9IjIyMiIgZm9udC1mYW1pbHk9IlRhaG9tYSwgc2Fucy1zZXJpZiIgZm9udC1zaXplPSIxMSIgZmlsbD0iIzRFOTk3MyIgdGV4dC1hbmNob3I9Im1pZGRsZSI+4Lig4Liy4Lie4LiV4Lix4Lin4Lit4Lii4LmI4Liy4LiH4Lib4Lij4Liw4LiB4Lit4LiaIChQbGFjZWhvbGRlcik8L3RleHQ+Cjwvc3ZnPg==" },
      status: "pending_supervisor",
      history: [
        { stage: "ส่งเรื่อง", actor: "QA", action: "submit", comment: "", date: "2026-09-01 14:50" },
      ],
    }),
    mk({
      registeredNo: "KZ-2026-018",
      date: "2026-08-15",
      title: "กล่องแยกขยะรีไซเคิลประจำไลน์ผลิต",
      team: "ทีม Packing 2",
      department: "Packing 2",
      submittedBy: "Production Line 2",
      problem: "ขยะทั่วไปและขยะรีไซเคิลปนกัน เพิ่มต้นทุนกำจัดขยะ",
      improvement: "ติดตั้งกล่องแยกขยะ 3 สี พร้อมป้ายชี้บ่งชัดเจนทุกไลน์",
      budget: "1200",
      result: "ปฏิเสธ - ซ้ำซ้อนกับโครงการ 5ส ที่ทำอยู่แล้ว",
      categories: ["Environment"],
      images: { before: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNGREYzRTYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjRjBCODYwIi8+Cjx0ZXh0IHg9IjIwMCIgeT0iMTEzIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjUyIiBmb250LXdlaWdodD0iYm9sZCIgZmlsbD0iIzhBNUIxNCIgdGV4dC1hbmNob3I9Im1pZGRsZSI+ITwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guIHguYjguK3guJnguJvguKPguLHguJrguJvguKPguLjguIcgKEJlZm9yZSk8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjAwIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjEzIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj5QYWNraW5nIDI8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjIyIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjExIiBmaWxsPSIjQjA4QjRBIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKDguLLguJ7guJXguLHguKfguK3guKLguYjguLLguIfguJvguKPguLDguIHguK3guJogKFBsYWNlaG9sZGVyKTwvdGV4dD4KPC9zdmc+", after: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNFOUY1RUYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjNkZCRTk3Ii8+CjxwYXRoIGQ9Ik0xODAgOTcgbDE0IDE0IGwyOCAtMzIiIHN0cm9rZT0iIzFGNUM0MCIgc3Ryb2tlLXdpZHRoPSI5IiBmaWxsPSJub25lIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjMUY1QzQwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKvguKXguLHguIfguJvguKPguLHguJrguJvguKPguLjguIcgKEFmdGVyKTwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIyMDAiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTMiIGZpbGw9IiMxRjVDNDAiIHRleHQtYW5jaG9yPSJtaWRkbGUiPlBhY2tpbmcgMjwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIyMjIiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTEiIGZpbGw9IiM0RTk5NzMiIHRleHQtYW5jaG9yPSJtaWRkbGUiPuC4oOC4suC4nuC4leC4seC4p+C4reC4ouC5iOC4suC4h+C4m+C4o+C4sOC4geC4reC4miAoUGxhY2Vob2xkZXIpPC90ZXh0Pgo8L3N2Zz4=" },
      status: "rejected",
      history: [
        { stage: "ส่งเรื่อง", actor: "Production Line 2", action: "submit", comment: "", date: "2026-08-15 10:20" },
        { stage: "หัวหน้างาน", actor: "ธันวา ปักษาสุข", action: "reject", comment: "ซ้ำซ้อนกับกิจกรรม 5ส ที่ดำเนินการอยู่แล้ว โปรดปรับหัวข้อใหม่", date: "2026-08-16 13:05" },
      ],
    }),
    mk({
      registeredNo: "KZ-2026-019",
      date: "2026-09-03",
      title: "ลดขั้นตอนการเบิกอุปกรณ์ป้องกันส่วนบุคคล",
      team: "ทีม Cutting",
      department: "Cutting",
      submittedBy: "Production Line 4",
      problem: "ขั้นตอนเบิก PPE ต้องเซ็นเอกสาร 3 จุด เสียเวลาเฉลี่ย 15 นาที/ครั้ง",
      improvement: "ใช้ตู้เบิกอัตโนมัติสแกนบัตรพนักงาน ลดเอกสารเหลือ 1 จุด",
      budget: "0",
      result: "",
      categories: ["Cost", "Satisfaction"],
      images: { before: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNGREYzRTYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjRjBCODYwIi8+Cjx0ZXh0IHg9IjIwMCIgeT0iMTEzIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjUyIiBmb250LXdlaWdodD0iYm9sZCIgZmlsbD0iIzhBNUIxNCIgdGV4dC1hbmNob3I9Im1pZGRsZSI+ITwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guIHguYjguK3guJnguJvguKPguLHguJrguJvguKPguLjguIcgKEJlZm9yZSk8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjAwIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjEzIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj5DdXR0aW5nPC90ZXh0Pgo8dGV4dCB4PSIyMDAiIHk9IjIyMiIgZm9udC1mYW1pbHk9IlRhaG9tYSwgc2Fucy1zZXJpZiIgZm9udC1zaXplPSIxMSIgZmlsbD0iI0IwOEI0QSIgdGV4dC1hbmNob3I9Im1pZGRsZSI+4Lig4Liy4Lie4LiV4Lix4Lin4Lit4Lii4LmI4Liy4LiH4Lib4Lij4Liw4LiB4Lit4LiaIChQbGFjZWhvbGRlcik8L3RleHQ+Cjwvc3ZnPg==", after: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNFOUY1RUYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjNkZCRTk3Ii8+CjxwYXRoIGQ9Ik0xODAgOTcgbDE0IDE0IGwyOCAtMzIiIHN0cm9rZT0iIzFGNUM0MCIgc3Ryb2tlLXdpZHRoPSI5IiBmaWxsPSJub25lIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjMUY1QzQwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKvguKXguLHguIfguJvguKPguLHguJrguJvguKPguLjguIcgKEFmdGVyKTwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIyMDAiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTMiIGZpbGw9IiMxRjVDNDAiIHRleHQtYW5jaG9yPSJtaWRkbGUiPkN1dHRpbmc8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjIyIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjExIiBmaWxsPSIjNEU5OTczIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKDguLLguJ7guJXguLHguKfguK3guKLguYjguLLguIfguJvguKPguLDguIHguK3guJogKFBsYWNlaG9sZGVyKTwvdGV4dD4KPC9zdmc+" },
      status: "draft",
      history: [],
    }),
    mk({
      registeredNo: "KZ-2026-020",
      date: "2026-07-18",
      title: "ปรับปรุงจุดวางพาเลทลดระยะเดินคลังสินค้า",
      team: "ทีม Finish Goods W/H (Dry)",
      department: "Finish Goods W/H (Dry)",
      submittedBy: "Finish Goods W/H (Dry)",
      problem: "จุดวางพาเลทกระจายไม่เป็นระบบ พนักงานเดินหยิบสินค้าไกลเกินจำเป็น",
      improvement: "จัดโซนพาเลทตามความถี่การหยิบ (Fast/Slow moving) พร้อมป้ายชี้บ่งโซน",
      budget: "800",
      result: "ลดระยะเดินเฉลี่ย 35% ต่อรอบการหยิบสินค้า",
      costSaving: "12500",
      costSavingCalc: "ลดเวลาเดินหยิบสินค้าเฉลี่ย 3 นาที/รอบ x 40 รอบ/วัน x 300 วัน x ค่าแรงเฉลี่ย 3.5 บาท/นาที",
      categories: ["Delivery", "Cost"],
      images: { before: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNGREYzRTYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjRjBCODYwIi8+Cjx0ZXh0IHg9IjIwMCIgeT0iMTEzIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjUyIiBmb250LXdlaWdodD0iYm9sZCIgZmlsbD0iIzhBNUIxNCIgdGV4dC1hbmNob3I9Im1pZGRsZSI+ITwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guIHguYjguK3guJnguJvguKPguLHguJrguJvguKPguLjguIcgKEJlZm9yZSk8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjAwIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjEzIiBmaWxsPSIjOEE1QjE0IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj5GaW5pc2ggR29vZHMgVy9IIChEcnkpPC90ZXh0Pgo8dGV4dCB4PSIyMDAiIHk9IjIyMiIgZm9udC1mYW1pbHk9IlRhaG9tYSwgc2Fucy1zZXJpZiIgZm9udC1zaXplPSIxMSIgZmlsbD0iI0IwOEI0QSIgdGV4dC1hbmNob3I9Im1pZGRsZSI+4Lig4Liy4Lie4LiV4Lix4Lin4Lit4Lii4LmI4Liy4LiH4Lib4Lij4Liw4LiB4Lit4LiaIChQbGFjZWhvbGRlcik8L3RleHQ+Cjwvc3ZnPg==", after: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MDAgMjYwIj4KPHJlY3Qgd2lkdGg9IjQwMCIgaGVpZ2h0PSIyNjAiIGZpbGw9IiNFOUY1RUYiLz4KPGNpcmNsZSBjeD0iMjAwIiBjeT0iOTUiIHI9IjQyIiBmaWxsPSIjNkZCRTk3Ii8+CjxwYXRoIGQ9Ik0xODAgOTcgbDE0IDE0IGwyOCAtMzIiIHN0cm9rZT0iIzFGNUM0MCIgc3Ryb2tlLXdpZHRoPSI5IiBmaWxsPSJub25lIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KPHRleHQgeD0iMjAwIiB5PSIxNzUiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTciIGZvbnQtd2VpZ2h0PSJib2xkIiBmaWxsPSIjMUY1QzQwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKvguKXguLHguIfguJvguKPguLHguJrguJvguKPguLjguIcgKEFmdGVyKTwvdGV4dD4KPHRleHQgeD0iMjAwIiB5PSIyMDAiIGZvbnQtZmFtaWx5PSJUYWhvbWEsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTMiIGZpbGw9IiMxRjVDNDAiIHRleHQtYW5jaG9yPSJtaWRkbGUiPkZpbmlzaCBHb29kcyBXL0ggKERyeSk8L3RleHQ+Cjx0ZXh0IHg9IjIwMCIgeT0iMjIyIiBmb250LWZhbWlseT0iVGFob21hLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjExIiBmaWxsPSIjNEU5OTczIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj7guKDguLLguJ7guJXguLHguKfguK3guKLguYjguLLguIfguJvguKPguLDguIHguK3guJogKFBsYWNlaG9sZGVyKTwvdGV4dD4KPC9zdmc+" },
      status: "approved",
      history: [
        { stage: "ส่งเรื่อง", actor: "Finish Goods W/H (Dry)", action: "submit", comment: "", date: "2026-07-18 16:40" },
        { stage: "หัวหน้างาน", actor: "นันทวรรณ นวลเนตร", action: "approve", comment: "", date: "2026-07-19 09:55" },
        { stage: "Admin", actor: "Admin", action: "approve", comment: "", date: "2026-07-22 11:30" },
        { stage: "ผู้อำนวยการ", actor: "ผอ.โรงงาน", action: "approve", comment: "ผลดีมาก ให้เป็นต้นแบบให้คลังอื่น", date: "2026-07-25 15:05" },
      ],
    }),
    mk({
      registeredNo: "KZ-2026-008",
      date: "2026-06-12",
      title: "ระบบแจ้งเตือนอุณหภูมิห้องเย็นอัตโนมัติ",
      team: "ทีม QC & Engineering",
      department: "QC",
      submittedBy: "QC",
      problem: "ตรวจอุณหภูมิห้องเย็นด้วยการจดมือวันละ 2 ครั้ง หากเกิดความผิดปกติช่วงกลางคืนจะรู้ช้า เสี่ยงต่อความปลอดภัยอาหาร",
      improvement: "ติดตั้งเซนเซอร์อุณหภูมิพร้อมแจ้งเตือนผ่านไลน์กลุ่มเมื่อเกินค่ามาตรฐาน",
      budget: "15000",
      result: "ตรวจพบความผิดปกติภายใน 5 นาที ลดสินค้าเสียหายจากอุณหภูมิเป็นศูนย์",
      costSaving: "65000",
      costSavingCalc: "ลดสินค้าเสียหายเฉลี่ย 5,400 บาท/เดือน x 12 เดือน",
      categories: ["Food Safety", "Quality"],
      kaizenType: "Cross functional Kaizen",
      members: fillMembers(["สมชาย ใจดี", "วิไล ศรีสุข", "ประเสริฐ มั่นคง", "อรุณี แก้วใส"]),
      status: "approved",
    }),
    mk({
      registeredNo: "KZ-2026-010",
      date: "2026-05-28",
      title: "ลดการสูญเสียแป้งระหว่างการลำเลียง",
      team: "ทีม Thread forming 1",
      department: "Thread forming 1",
      submittedBy: "Production_Line 1",
      problem: "แป้งหกตกหล่นบริเวณจุดต่อสายพานลำเลียง ต้องทิ้งเป็นของเสียทุกวัน",
      improvement: "ออกแบบแผ่นกันหกและปรับมุมสายพาน พร้อมถาดรองแป้งกลับเข้าระบบ",
      budget: "30000",
      result: "ลดแป้งสูญเสียจาก 40 กก./วัน เหลือ 8 กก./วัน",
      costSaving: "120000",
      costSavingCalc: "ลดแป้งสูญเสีย 32 กก./วัน x 300 วัน x 12.5 บาท/กก.",
      categories: ["Cost", "Environment"],
      members: fillMembers(["มานพ รุ่งเรือง", "สุดา ทองดี", "เกรียงไกร พลอยงาม", "นภา สายสวัสดิ์", "ชัยวัฒน์ ดีเลิศ"]),
      status: "approved",
    }),
    mk({
      registeredNo: "KZ-2026-011",
      date: "2026-06-30",
      title: "ปรับรอบการเดินหม้อไอน้ำตามโหลดการผลิตจริง",
      team: "ทีม Boiler",
      department: "Boiler",
      submittedBy: "Utility",
      problem: "เดินหม้อไอน้ำเต็มกำลังตลอดกะ แม้ช่วงที่ไลน์ผลิตใช้ไอน้ำน้อย",
      improvement: "จัดตารางเดินหม้อไอน้ำตามแผนการผลิตรายวันโดยไม่ต้องลงทุนเพิ่ม",
      budget: "0",
      result: "ลดการใช้เชื้อเพลิง 9%",
      costSaving: "85000",
      costSavingCalc: "ลดค่าเชื้อเพลิงเฉลี่ย 7,080 บาท/เดือน x 12 เดือน",
      categories: ["Energy", "Cost"],
      members: fillMembers(["ธนพล แสงทอง", "ปิยะ วงศ์ใหญ่"]),
      status: "approved",
    }),
    mk({
      registeredNo: "KZ-2026-012",
      date: "2026-07-05",
      title: "ป้ายบ่งชี้สารก่อภูมิแพ้ในไลน์บรรจุ",
      team: "ทีม Packing 2",
      department: "Packing 2",
      submittedBy: "Production Line 2",
      problem: "พนักงานใหม่สับสนจุดวางวัตถุดิบที่มีสารก่อภูมิแพ้ เสี่ยงปนเปื้อนข้าม",
      improvement: "ทำป้ายสีและเส้นแบ่งพื้นที่สำหรับวัตถุดิบกลุ่มสารก่อภูมิแพ้",
      budget: "1200",
      result: "ไม่พบการวางผิดจุดตลอด 2 เดือนหลังปรับปรุง",
      costSaving: "6000",
      costSavingCalc: "ลดเวลาตรวจสอบซ้ำและของเสียจากการปนเปื้อน 500 บาท/เดือน x 12 เดือน",
      categories: ["Food Safety"],
      members: fillMembers(["รัตนา บุญมา", "สมศักดิ์ ใจงาม", "จันทร์เพ็ญ ศรีทอง"]),
      status: "approved",
    }),
  ];
}

function StatusBadge({ status }) {
  const m = STATUS_META[status];
  return (
    <span
      className="inline-flex items-center gap-1.5 rounded-full px-3 py-1 text-xs font-medium whitespace-nowrap"
      style={{ color: m.color, background: m.bg }}
    >
      <span className="w-1.5 h-1.5 rounded-full" style={{ background: m.color }} />
      {m.label}
    </span>
  );
}

function CategoryTags({ categories }) {
  return (
    <div className="flex flex-wrap gap-1">
      {categories.length === 0 && <span className="text-xs" style={{ color: SUB }}>ยังไม่ระบุ</span>}
      {categories.map((c) => (
        <span key={c} className="text-[11px] px-2 py-0.5 rounded border" style={{ borderColor: "#D8DDE5", color: INK }}>{c}</span>
      ))}
    </div>
  );
}

// ย่อรูปก่อนอัปโหลด (ด้านยาวสุด 1600px, JPEG) ให้บันทึกเร็วและไม่เปลืองพื้นที่ Google Drive
function compressImage(file, maxSide = 1600, quality = 0.85) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxSide / Math.max(img.width, img.height));
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(img.width * scale);
        canvas.height = Math.round(img.height * scale);
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "#FFFFFF";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = () => resolve(reader.result);
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

function ImageBox({ label, src, onChange, editable, labelClassName, onZoom, boxClassName }) {
  return (
    <div>
      <div className={labelClassName || "text-xs mb-1.5 font-medium"} style={{ color: labelClassName ? INK : SUB }}>{label}</div>
      <div
        className={`rounded-lg border ${boxClassName || "h-36"} flex items-center justify-center overflow-hidden relative group`}
        style={{ borderColor: "#D8DDE5", background: "#FAFBFC" }}
      >
        {src ? (
          <img
            src={src}
            alt={label}
            className="w-full h-full object-cover"
            style={onZoom ? { cursor: "zoom-in" } : undefined}
            onClick={onZoom ? () => onZoom(src) : undefined}
          />
        ) : (
          <div className="flex flex-col items-center gap-1" style={{ color: "#AEB6C2" }}>
            <ImageOff size={22} />
            <span className="text-[11px]">ไม่มีรูปภาพ</span>
          </div>
        )}
        {onZoom && src && (
          <div
            className="absolute inset-0 flex items-center justify-center opacity-0 group-hover:opacity-100 transition pointer-events-none"
            style={{ background: "rgba(11,29,51,0.25)" }}
          >
            <div className="bg-white/90 rounded-full p-2">
              <ZoomIn size={18} style={{ color: NAVY }} />
            </div>
          </div>
        )}
        {editable && (
          <label className="absolute bottom-2 right-2 bg-white shadow rounded-full p-1.5 cursor-pointer border" style={{ borderColor: "#D8DDE5" }}>
            <Upload size={14} style={{ color: NAVY }} />
            <input
              type="file"
              accept="image/*"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                compressImage(f).then(onChange);
                e.target.value = "";
              }}
            />
          </label>
        )}
        {editable && src && (
          <button
            type="button"
            onClick={() => onChange(null)}
            title="ลบรูปนี้"
            className="absolute top-2 right-2 bg-white shadow rounded-full p-1 border"
            style={{ borderColor: "#D8DDE5" }}
          >
            <X size={13} style={{ color: RED }} />
          </button>
        )}
      </div>
    </div>
  );
}

function Field({ label, children, required, hint }) {
  return (
    <div>
      <div className="text-xs mb-1.5 font-medium" style={{ color: SUB }}>
        {label} {required && <span style={{ color: RED }}>*</span>}
      </div>
      {children}
      {hint && <div className="text-[11px] mt-1" style={{ color: "#8A93A3" }}>{hint}</div>}
    </div>
  );
}

const inputCls = "w-full rounded-lg border px-3 py-2 text-sm outline-none focus:ring-2";
const inputStyle = { borderColor: "#D8DDE5", color: INK };

function TextInput(props) {
  return <input {...props} className={inputCls} style={{ ...inputStyle, ...(props.style || {}) }} onFocus={(e) => (e.target.style.borderColor = NAVY)} onBlur={(e) => (e.target.style.borderColor = "#D8DDE5")} />;
}
function TextArea(props) {
  return <textarea {...props} className={inputCls} style={{ ...inputStyle, minHeight: 80, ...(props.style || {}) }} onFocus={(e) => (e.target.style.borderColor = NAVY)} onBlur={(e) => (e.target.style.borderColor = "#D8DDE5")} />;
}

function MemberHalf({ rows, start, setRow }) {
  return (
    <table className="w-full text-xs">
      <thead>
        <tr style={{ color: SUB }}>
          <th className="text-left font-medium py-1 w-6">#</th>
          <th className="text-left font-medium py-1">ชื่อ-นามสกุล</th>
          <th className="text-left font-medium py-1 w-20">รหัสพนักงาน</th>
          <th className="text-left font-medium py-1 w-20">หน่วยงาน</th>
        </tr>
      </thead>
      <tbody>
        {rows.slice(start, start + 5).map((r, k) => {
          const i = start + k;
          return (
            <tr key={i}>
              <td className="py-1" style={{ color: SUB }}>{i + 1}</td>
              <td className="py-1 pr-1">
                <input value={r.name} onChange={(e) => setRow(i, "name", e.target.value)} className="w-full rounded border px-2 py-1 text-xs" style={inputStyle} />
              </td>
              <td className="py-1 pr-1">
                <input value={r.empId} onChange={(e) => setRow(i, "empId", e.target.value)} className="w-full rounded border px-2 py-1 text-xs" style={inputStyle} />
              </td>
              <td className="py-1">
                <input value={r.dept} onChange={(e) => setRow(i, "dept", e.target.value)} className="w-full rounded border px-2 py-1 text-xs" style={inputStyle} />
              </td>
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

function MemberTable({ members, onChange }) {
  const rows = members && members.length === 10 ? members : emptyMembers();
  const setRow = (i, key, val) => {
    const next = rows.map((r, idx) => (idx === i ? { ...r, [key]: val } : r));
    onChange(next);
  };
  return (
    <div>
      <div className="text-xs mb-1.5 font-medium" style={{ color: SUB }}>รายชื่อสมาชิกทีม (สูงสุด 10 คน)</div>
      <div className="grid grid-cols-2 gap-4 rounded-lg border p-3" style={{ borderColor: "#D8DDE5" }}>
        <MemberHalf rows={rows} start={0} setRow={setRow} />
        <MemberHalf rows={rows} start={5} setRow={setRow} />
      </div>
    </div>
  );
}

function AttachmentList({ attachments, onAdd, onRemove, editable }) {
  const fileIcon = (type, name) => {
    if ((type || "").startsWith("video/")) return FileVideo;
    if (/\.(xlsx|xls|csv)$/i.test(name || "")) return FileSpreadsheet;
    return FileText;
  };
  const handleFiles = (fileList) => {
    Array.from(fileList).forEach((f) => {
      const reader = new FileReader();
      reader.onload = () => onAdd({ id: uid(), name: f.name, size: f.size, type: f.type || "", dataUrl: reader.result });
      reader.readAsDataURL(f);
    });
  };
  return (
    <div>
      <div className="flex items-center justify-between mb-1.5">
        <div className="text-xs font-medium" style={{ color: SUB }}>เอกสาร/ไฟล์แนบเพิ่มเติม (ถ้ามี เช่น วิดีโอ ไฟล์คำนวณ Excel เอกสารอ้างอิง)</div>
        {editable && (
          <label className="text-xs cursor-pointer flex items-center gap-1 font-medium flex-shrink-0 ml-2" style={{ color: NAVY }}>
            <Paperclip size={12} /> แนบไฟล์
            <input
              type="file"
              multiple
              className="hidden"
              onChange={(e) => { handleFiles(e.target.files); e.target.value = ""; }}
            />
          </label>
        )}
      </div>
      {(!attachments || attachments.length === 0) ? (
        <div className="text-xs" style={{ color: "#AEB6C2" }}>ยังไม่มีไฟล์แนบ</div>
      ) : (
        <div className="space-y-1.5">
          {attachments.map((a) => {
            const Icon = fileIcon(a.type, a.name);
            return (
              <div key={a.id} className="flex items-center gap-2 rounded-lg border px-2.5 py-1.5 text-xs" style={{ borderColor: "#D8DDE5" }}>
                <Icon size={14} style={{ color: NAVY, flexShrink: 0 }} />
                <a href={a.dataUrl} download={a.name} className="flex-1 truncate hover:underline" style={{ color: INK }}>{a.name}</a>
                <span className="flex-shrink-0" style={{ color: SUB }}>{((a.size || 0) / 1024).toFixed(0)} KB</span>
                {editable && (
                  <button onClick={() => onRemove(a.id)} className="flex-shrink-0">
                    <X size={13} style={{ color: RED }} />
                  </button>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// รายการข้อมูลที่ยังขาด (ใช้แจ้งเตือน Admin)
function missingInfo(k) {
  const out = [];
  if (!k.images?.before) out.push("ภาพก่อนปรับปรุง");
  if (!k.images?.after) out.push("ภาพหลังปรับปรุง");
  if (!String(k.problem || "").trim()) out.push("ที่มาของปัญหา");
  if (!String(k.improvement || "").trim()) out.push("แนวทางการปรับปรุง");
  if (!String(k.result || "").trim()) out.push("ผลที่ได้รับ");
  if (!(k.categories || []).length) out.push("หมวดหมู่ QCDSME");
  return out;
}

function KaizenForm({ initial, onCancel, onSave, currentUser, departments, adminMode }) {
  // เลือกได้เฉพาะหน่วยงานย่อยของผู้ใช้งานที่ล็อกอิน ถ้าไม่มีสิทธิ์เฉพาะ (เช่น Admin) ให้เลือกได้จากทุกหน่วยงาน
  const deptList = currentUser?.departments?.length
    ? currentUser.departments
    : (departments && departments.length ? departments : DEFAULT_DEPARTMENTS);
  const [f, setF] = useState(
    initial || {
      title: "", team: "",
      department: currentUser?.departments?.[0] || deptList[0],
      submittedBy: currentUser?.name || "",
      problem: "", improvement: "", budget: "", result: "",
      costSaving: "", costSavingCalc: "",
      members: emptyMembers(),
      images: { before: null, after: null },
      attachments: [],
    }
  );
  const set = (k, v) => setF((s) => ({ ...s, [k]: v }));
  const [adminNote, setAdminNote] = useState("");
  const missing = adminMode ? missingInfo(f) : [];

  return (
    <div className="space-y-4">
      {adminMode && (
        <div className="rounded-xl border p-3.5" style={{ background: "#FDF3E1", borderColor: "#E8C77E" }}>
          <div className="text-sm font-semibold" style={{ color: "#6B4A10" }}>แก้ไขข้อมูลในฐานะ Admin · {initial?.registeredNo}</div>
          <div className="text-xs mt-1" style={{ color: "#6B4A10" }}>
            สถานะเดิม <strong>{STATUS_META[initial?.status]?.label || initial?.status}</strong> จะไม่เปลี่ยน · ระบบบันทึกการแก้ไขไว้ในประวัติการดำเนินเรื่องอัตโนมัติ
          </div>
          {missing.length > 0 && (
            <div className="text-xs mt-2 flex flex-wrap gap-1.5 items-center" style={{ color: "#8B2E24" }}>
              <AlertTriangle size={13} /> ข้อมูลที่ยังขาด:
              {missing.map((m) => (
                <span key={m} className="px-2 py-0.5 rounded-full" style={{ background: "#FBEAE8" }}>{m}</span>
              ))}
            </div>
          )}
        </div>
      )}
      <div className="grid grid-cols-2 gap-4">
        <Field label="หัวเรื่อง / ชื่อโครงการ" required>
          <TextInput value={f.title} onChange={(e) => set("title", e.target.value)} placeholder="เช่น รถเข็นเครื่องชั่งดิจิตอลแบบตั้งโต๊ะ" />
        </Field>
        <Field label="ชื่อทีม">
          <TextInput value={f.team} onChange={(e) => set("team", e.target.value)} placeholder="เช่น ทีม Safety บรรจุ 3" />
        </Field>
        <Field label="หน่วยงาน" required>
          <select className={inputCls} style={inputStyle} value={f.department} onChange={(e) => set("department", e.target.value)}>
            {deptList.map((d) => <option key={d}>{d}</option>)}
          </select>
        </Field>
        <Field label="ผู้ส่งเรื่อง (ชื่อ-นามสกุล)" required>
          <TextInput disabled={!!currentUser} style={{ opacity: currentUser ? 0.7 : 1 }} value={f.submittedBy} onChange={(e) => set("submittedBy", e.target.value)} placeholder="ชื่อ นามสกุล" />
        </Field>
      </div>
      <MemberTable members={f.members} onChange={(m) => set("members", m)} />
      <Field label="1. ที่มาของปัญหาและสถานการณ์ปัจจุบัน" required>
        <TextArea value={f.problem} onChange={(e) => set("problem", e.target.value)} placeholder="อธิบายปัญหาที่พบก่อนการปรับปรุง" />
      </Field>
      <Field label="2. แนวทางการปรับปรุง Improvement Idea" required>
        <TextArea value={f.improvement} onChange={(e) => set("improvement", e.target.value)} placeholder="อธิบายวิธีการแก้ไข/ปรับปรุง" />
      </Field>
      <div className="grid grid-cols-2 gap-4">
        <Field label="3. งบประมาณที่ใช้ในการปรับปรุง" hint="ใส่แค่ตัวเลข (บาท) และถ้าไม่มี ใส่เลข 0">
          <TextInput type="number" min="0" value={f.budget} onChange={(e) => set("budget", e.target.value)} placeholder="0" />
        </Field>
        <Field label="4. ผลที่ได้รับ">
          <TextInput value={f.result} onChange={(e) => set("result", e.target.value)} placeholder="กรอกภายหลังหากยังไม่มีผลลัพธ์" />
        </Field>
      </div>
      <div className="grid grid-cols-2 gap-4">
        <Field label="5. Cost Saving ที่ได้" hint="ใส่แค่ตัวเลข (บาท ต่อปี) หากไม่มีให้ใส่เลข 0">
          <TextInput type="number" min="0" value={f.costSaving} onChange={(e) => set("costSaving", e.target.value)} placeholder="0" />
        </Field>
        <Field label="6. การคำนวณผลลัพธ์ (หากมี Cost Saving)" hint="หากไม่มี Cost Saving ไม่ต้องระบุข้อนี้">
          <TextInput value={f.costSavingCalc} onChange={(e) => set("costSavingCalc", e.target.value)} placeholder="เช่น วิธีคิด/สูตรคำนวณ Cost Saving" />
        </Field>
      </div>
      <div className="grid grid-cols-2 gap-4">
        <ImageBox label="ภาพก่อนปรับปรุง (Before)" src={f.images.before} editable onChange={(v) => set("images", { ...f.images, before: v })} />
        <ImageBox label="ภาพหลังปรับปรุง (After)" src={f.images.after} editable onChange={(v) => set("images", { ...f.images, after: v })} />
      </div>
      <AttachmentList
        attachments={f.attachments || []}
        editable
        onAdd={(file) => set("attachments", [...(f.attachments || []), file])}
        onRemove={(id) => set("attachments", (f.attachments || []).filter((a) => a.id !== id))}
      />
      {adminMode && (
        <div className="rounded-xl border p-4 space-y-3" style={{ borderColor: "#E3E7ED", background: "#FAFBFC" }}>
          <div className="text-sm font-semibold" style={{ color: INK }}>ข้อมูลจัดหมวดหมู่ (Admin)</div>
          <div>
            <div className="text-xs mb-1.5 font-medium" style={{ color: SUB }}>หมวดหมู่ QCDSME</div>
            <div className="flex flex-wrap gap-1.5">
              {CATEGORY_OPTIONS.map((c) => {
                const on = (f.categories || []).includes(c);
                return (
                  <button
                    key={c}
                    type="button"
                    onClick={() => set("categories", on ? f.categories.filter((x) => x !== c) : [...(f.categories || []), c])}
                    className="text-xs px-2.5 py-1 rounded-full border"
                    style={on ? { background: NAVY, borderColor: NAVY, color: "white" } : { borderColor: "#D8DDE5", color: SUB, background: "white" }}
                  >
                    {c}
                  </button>
                );
              })}
            </div>
          </div>
          <div className="grid grid-cols-2 gap-4">
            <Field label="ประเภทกิจกรรม">
              <select className={inputCls} style={inputStyle} value={f.activityType || "Kaizen"} onChange={(e) => set("activityType", e.target.value)}>
                {ACTIVITY_OPTIONS.map((o) => <option key={o}>{o}</option>)}
              </select>
            </Field>
            <Field label="ประเภท Kaizen">
              <select className={inputCls} style={inputStyle} value={f.kaizenType || "Kaizen Project"} onChange={(e) => set("kaizenType", e.target.value)}>
                {KAIZEN_TYPE_OPTIONS.map((o) => <option key={o}>{o}</option>)}
              </select>
            </Field>
          </div>
          <Field label="หมายเหตุการแก้ไข (บันทึกในประวัติ ไม่บังคับ)">
            <TextInput value={adminNote} onChange={(e) => setAdminNote(e.target.value)} placeholder="เช่น เพิ่มรูปก่อน-หลังตามที่หน่วยงานส่งมาทางไลน์" />
          </Field>
        </div>
      )}
      {adminMode ? (
        <div className="flex justify-end gap-2 pt-2">
          <button onClick={onCancel} className="px-4 py-2 rounded-lg text-sm font-medium border" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
          <button
            disabled={!f.title}
            onClick={() => onSave(f, adminNote)}
            className="px-4 py-2 rounded-lg text-sm font-medium text-white disabled:opacity-40 flex items-center gap-1.5"
            style={{ background: NAVY }}
          >
            <CheckCircle2 size={15} /> บันทึกการแก้ไข
          </button>
        </div>
      ) : (
      <div className="flex justify-end gap-2 pt-2">
        <button onClick={onCancel} className="px-4 py-2 rounded-lg text-sm font-medium border" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
        <button
          onClick={() => onSave(f, "draft")}
          className="px-4 py-2 rounded-lg text-sm font-medium border"
          style={{ borderColor: NAVY, color: NAVY }}
        >
          บันทึกแบบร่าง
        </button>
        <button
          disabled={!f.title || !f.problem || !f.improvement}
          onClick={() => onSave(f, "pending_supervisor")}
          className="px-4 py-2 rounded-lg text-sm font-medium text-white disabled:opacity-40"
          style={{ background: NAVY }}
        >
          ส่งเรื่อง
        </button>
      </div>
      )}
    </div>
  );
}

function NotificationPanel({ notifications, onItemClick, onMarkAllRead, onClose }) {
  return (
    <div className="absolute right-0 top-full mt-2 w-80 bg-white rounded-xl shadow-2xl border z-40" style={{ borderColor: "#D8DDE5" }}>
      <div className="flex items-center justify-between px-4 py-3 border-b" style={{ borderColor: "#EEF0F3" }}>
        <div className="font-semibold text-sm" style={{ color: INK }}>การแจ้งเตือน</div>
        <div className="flex items-center gap-2">
          {notifications.length > 0 && (
            <button onClick={onMarkAllRead} className="text-xs font-medium" style={{ color: NAVY }}>
              อ่านทั้งหมด
            </button>
          )}
          <button onClick={onClose} className="p-0.5 rounded hover:bg-gray-100"><X size={15} style={{ color: SUB }} /></button>
        </div>
      </div>
      <div className="max-h-80 overflow-y-auto">
        {notifications.length === 0 && (
          <div className="px-4 py-8 text-center text-xs" style={{ color: SUB }}>ยังไม่มีการแจ้งเตือน</div>
        )}
        {notifications.map((n) => (
          <button
            key={n.id}
            onClick={() => onItemClick(n)}
            className="w-full text-left px-4 py-3 border-b last:border-b-0 flex gap-2.5 items-start hover:bg-blue-50"
            style={{ borderColor: "#F2F4F7", background: n.read ? "white" : "#F5F9FF" }}
          >
            <div className="w-1.5 h-1.5 rounded-full mt-1.5 flex-shrink-0" style={{ background: n.read ? "transparent" : NAVY }} />
            <div className="flex-1 min-w-0">
              <div className="text-xs leading-snug" style={{ color: INK, fontWeight: n.read ? 400 : 600 }}>{n.message}</div>
              <div className="text-[10px] mt-1" style={{ color: SUB }}>{n.date}</div>
            </div>
          </button>
        ))}
      </div>
    </div>
  );
}

function ImageLightbox({ src, onClose }) {
  if (!src) return null;
  return (
    <div
      className="fixed inset-0 z-[100] flex items-center justify-center p-6"
      style={{ background: "rgba(0,0,0,0.85)" }}
      onClick={onClose}
    >
      <button
        onClick={onClose}
        className="absolute top-4 right-4 p-2 rounded-full"
        style={{ background: "rgba(255,255,255,0.15)" }}
      >
        <X size={22} className="text-white" />
      </button>
      <img
        src={src}
        alt="ภาพขยาย"
        className="max-w-full max-h-full rounded-lg shadow-2xl object-contain"
        onClick={(e) => e.stopPropagation()}
      />
    </div>
  );
}

function Modal({ title, onClose, children, wide, headerExtra }) {
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto py-8 px-4" style={{ background: "rgba(11,29,51,0.55)" }}>
      <div className={`bg-white rounded-2xl shadow-2xl w-full ${wide ? "max-w-4xl" : "max-w-2xl"} my-auto`}>
        <div className="flex items-center justify-between px-6 py-4 border-b" style={{ borderColor: "#EAEDF1" }}>
          <h3 className="font-semibold text-base" style={{ color: INK }}>{title}</h3>
          <div className="flex items-center gap-2">
            {headerExtra}
            <button onClick={onClose} className="p-1 rounded-full hover:bg-gray-100"><X size={18} style={{ color: SUB }} /></button>
          </div>
        </div>
        <div className="p-6">{children}</div>
      </div>
    </div>
  );
}

const TARGET_LABEL = {
  employee: "พนักงานผู้ส่งเรื่อง (เริ่มกระบวนการใหม่)",
  supervisor: "หัวหน้างาน",
  admin: "Admin",
};

function Timeline({ history }) {
  if (!history.length) return <div className="text-sm" style={{ color: SUB }}>ยังไม่มีความเคลื่อนไหว</div>;
  return (
    <div className="space-y-3">
      {history.map((h, i) => {
        const isNegative = h.action === "reject" || h.action === "return" || h.action === "recall";
        const actionLabel =
          h.action === "submit" ? "ส่งเรื่อง" :
          h.action === "approve" ? "อนุมัติ" :
          h.action === "reject" ? "ตีกลับ (ถึงพนักงาน)" :
          h.action === "return" ? `ส่งกลับไปยัง ${TARGET_LABEL[h.target] || h.target}` :
          h.action === "recall" ? "ถอนเรื่องกลับมาแก้ไข (แบบร่าง)" :
          h.action === "edit" ? "แก้ไขข้อมูลโดย Admin" : "";
        const [datePart, timePart] = (h.date || "").split(" ");
        return (
          <div key={i} className="flex gap-3">
            <div className="flex flex-col items-center">
              <div className="w-2.5 h-2.5 rounded-full mt-1" style={{ background: isNegative ? RED : h.action === "edit" ? GOLD : GREEN }} />
              {i !== history.length - 1 && <div className="w-px flex-1" style={{ background: "#E2E6EB" }} />}
            </div>
            <div className="pb-3">
              <div className="text-sm font-medium" style={{ color: INK }}>{h.stage} · {h.actor}</div>
              <div className="text-xs flex items-center gap-1 flex-wrap" style={{ color: SUB }}>
                <span>{datePart}</span>
                {timePart && (
                  <span className="inline-flex items-center gap-0.5">
                    <Clock size={11} /> {timePart} น.
                  </span>
                )}
                <span>— {actionLabel}</span>
              </div>
              {h.comment && <div className="text-sm mt-1 rounded-lg px-3 py-2" style={{ background: "#F7F8FA", color: INK }}>{h.comment}</div>}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function DetailView({ k }) {
  const [zoomSrc, setZoomSrc] = useState(null);
  return (
    <div className="space-y-5">
      <ImageLightbox src={zoomSrc} onClose={() => setZoomSrc(null)} />
      <div className="flex items-start justify-between">
        <div>
          <div className="text-xs" style={{ color: SUB }}>{k.registeredNo} · {k.date}</div>
          <div className="text-lg font-semibold mt-0.5 flex items-center gap-1.5" style={{ color: INK }}>
            <span>{k.title}</span>
            {k.attachments && k.attachments.length > 0 && (
              <span title={`มีไฟล์แนบ ${k.attachments.length} ไฟล์`} className="inline-flex items-center flex-shrink-0" style={{ color: SUB }}>
                <Paperclip size={16} />
              </span>
            )}
          </div>
          <div className="text-sm mt-0.5" style={{ color: SUB }}>{k.team} · {k.department} · ผู้ส่ง: {k.submittedBy}</div>
        </div>
        <StatusBadge status={k.status} />
      </div>
      <CategoryTags categories={k.categories} />
      {k.members && k.members.some((m) => m.name) && (
        <div className="text-sm">
          <div className="font-medium mb-1" style={{ color: INK }}>สมาชิกทีม</div>
          <div className="flex flex-wrap gap-1.5">
            {k.members.filter((m) => m.name).map((m, i) => (
              <span key={i} className="text-xs px-2 py-1 rounded border" style={{ borderColor: "#D8DDE5", color: SUB }}>
                {m.name}{m.empId ? ` · ${m.empId}` : ""}{m.dept ? ` · ${m.dept}` : ""}
              </span>
            ))}
          </div>
        </div>
      )}
      <div className="grid grid-cols-2 gap-4 text-sm">
        <div>
          <div className="font-medium mb-1" style={{ color: INK }}>ที่มาของปัญหา</div>
          <div style={{ color: SUB }}>{k.problem || "-"}</div>
        </div>
        <div>
          <div className="font-medium mb-1" style={{ color: INK }}>แนวทางการปรับปรุง</div>
          <div style={{ color: SUB }}>{k.improvement || "-"}</div>
        </div>
        <div>
          <div className="font-medium mb-1" style={{ color: INK }}>งบประมาณ</div>
          <div style={{ color: SUB }}>{k.budget ? `${k.budget} บาท` : "-"}</div>
        </div>
        <div>
          <div className="font-medium mb-1" style={{ color: INK }}>ผลที่ได้รับ</div>
          <div style={{ color: SUB }}>{k.result || "ยังไม่มีข้อมูล"}</div>
        </div>
        <div>
          <div className="font-medium mb-1" style={{ color: INK }}>Cost Saving ที่ได้</div>
          <div style={{ color: SUB }}>{k.costSaving ? `${k.costSaving} บาท/ปี` : "ไม่มี"}</div>
        </div>
        {k.costSavingCalc && (
          <div>
            <div className="font-medium mb-1" style={{ color: INK }}>การคำนวณผลลัพธ์</div>
            <div style={{ color: SUB }}>{k.costSavingCalc}</div>
          </div>
        )}
      </div>
      <div className="grid grid-cols-2 gap-4">
        <ImageBox label="ก่อนปรับปรุง" src={k.images.before} onZoom={setZoomSrc} />
        <ImageBox label="หลังปรับปรุง" src={k.images.after} onZoom={setZoomSrc} />
      </div>
      {k.attachments && k.attachments.length > 0 && (
        <AttachmentList attachments={k.attachments} editable={false} />
      )}
      <div>
        <div className="font-medium mb-2 text-sm" style={{ color: INK }}>ประวัติการดำเนินเรื่อง</div>
        <Timeline history={k.history} />
      </div>
    </div>
  );
}

function SectionPanel({ number, title, headerColor, children }) {
  return (
    <div className="rounded-xl overflow-hidden border flex flex-col h-full" style={{ borderColor: "#D8DDE5" }}>
      <div className="px-3 py-2 flex items-center gap-2 text-white text-sm font-semibold" style={{ background: headerColor }}>
        <span className="w-5 h-5 rounded-full bg-white/25 flex items-center justify-center text-xs flex-shrink-0">{number}</span>
        <span>{title}</span>
      </div>
      <div className="p-3 text-sm flex-1" style={{ color: INK, background: "#FBFBFC" }}>
        {children}
      </div>
    </div>
  );
}

function PresentationView({ k }) {
  const [zoomSrc, setZoomSrc] = useState(null);
  return (
    <div className="rounded-xl overflow-hidden border" style={{ borderColor: "#D8DDE5" }}>
      <ImageLightbox src={zoomSrc} onClose={() => setZoomSrc(null)} />

      {/* Header */}
      <div className="px-5 py-3.5 flex items-center justify-between flex-wrap gap-2" style={{ background: NAVY }}>
        <div className="text-white text-xl font-extrabold flex items-center gap-2">
          <span>ไคเซ็น <span className="font-semibold opacity-90">(KAIZEN)</span></span>
          {k.attachments && k.attachments.length > 0 && (
            <span title={`มีไฟล์แนบ ${k.attachments.length} ไฟล์`} className="inline-flex items-center flex-shrink-0 text-white/70">
              <Paperclip size={17} />
            </span>
          )}
        </div>
        <div className="flex items-center gap-5 text-white text-sm">
          <div className="flex items-center gap-1.5">
            <span className="opacity-75">หน่วยงาน</span>
            <span className="font-semibold border-b border-white/40 pb-0.5">{k.department}</span>
          </div>
          <div className="flex items-center gap-1.5">
            <span className="opacity-75">วันที่</span>
            <span className="font-semibold border-b border-white/40 pb-0.5">{k.date}</span>
          </div>
        </div>
      </div>

      <div className="p-5 space-y-4" style={{ background: "#F4F6F8" }}>
        {/* Title */}
        <div>
          <div className="text-xs font-medium mb-1.5" style={{ color: SUB }}>ชื่อเรื่องไคเซ็น</div>
          <div className="rounded-lg border px-4 py-2.5 text-base font-semibold" style={{ borderColor: "#D8DDE5", background: "white", color: INK }}>
            {k.title}
          </div>
        </div>

        {/* 1-2-3: Problem / Improvement / Result */}
        <div className="grid grid-cols-3 gap-3 items-stretch">
          <SectionPanel number={1} title="ปัญหา / สภาพปัจจุบัน" headerColor={DB_DEEP}>
            {k.problem || "-"}
          </SectionPanel>
          <SectionPanel number={2} title="แนวทางการปรับปรุง" headerColor="#1F8A70">
            {k.improvement || "-"}
          </SectionPanel>
          <SectionPanel number={3} title="ผลที่ได้รับ" headerColor={GREEN}>
            {k.result || "ยังไม่มีข้อมูล"}
          </SectionPanel>
        </div>

        {/* Before/After (large) + Cost Saving highlight */}
        <div className="grid gap-3" style={{ gridTemplateColumns: "2fr 1fr" }}>
          <div className="rounded-xl overflow-hidden border" style={{ borderColor: "#D8DDE5" }}>
            <div className="px-3 py-2 flex items-center gap-2 text-white text-sm font-semibold" style={{ background: NAVY_DEEP }}>
              <span className="w-5 h-5 rounded-full bg-white/25 flex items-center justify-center text-xs flex-shrink-0">4</span>
              <span>ภาพก่อน–หลัง | เห็นการเปลี่ยนแปลงในทันที</span>
            </div>
            <div className="p-3 grid gap-2 items-center" style={{ gridTemplateColumns: "1fr auto 1fr", background: "#FBFBFC" }}>
              <ImageBox label="ก่อนปรับปรุง" src={k.images.before} labelClassName="text-xs font-medium mb-1.5" boxClassName="h-56" onZoom={setZoomSrc} />
              <ChevronRight size={26} style={{ color: "#AEB6C2" }} />
              <ImageBox label="หลังปรับปรุง" src={k.images.after} labelClassName="text-xs font-medium mb-1.5" boxClassName="h-56" onZoom={setZoomSrc} />
            </div>
          </div>

          <div className="rounded-xl overflow-hidden border flex flex-col" style={{ borderColor: "#D8DDE5" }}>
            <div className="px-3 py-2 flex items-center gap-2 text-white text-sm font-semibold" style={{ background: `linear-gradient(135deg, #A8710F, ${GOLD})` }}>
              <span className="w-5 h-5 rounded-full bg-white/25 flex items-center justify-center text-xs flex-shrink-0">5</span>
              <span>ผลลัพธ์ | ตัวเลขที่ตรวจวัดผล</span>
            </div>
            <div
              className="relative flex-1 flex flex-col items-center justify-center text-center px-4 py-6 overflow-hidden"
              style={{ background: `linear-gradient(135deg, #A8710F, ${GOLD})`, minHeight: 240 }}
            >
              <div className="absolute -right-6 -top-6 w-24 h-24 rounded-full" style={{ background: "rgba(255,255,255,0.15)" }} />
              <div className="absolute -left-8 -bottom-8 w-28 h-28 rounded-full" style={{ background: "rgba(255,255,255,0.10)" }} />
              <div className="relative w-10 h-10 rounded-full flex items-center justify-center mb-2" style={{ background: "rgba(255,255,255,0.25)" }}>
                <Coins size={20} className="text-white" />
              </div>
              <div className="relative text-xs font-semibold text-white/90 tracking-wide">Cost Saving</div>
              <div className="relative text-4xl font-extrabold text-white mt-1 tracking-tight">
                {k.costSaving ? Number(k.costSaving).toLocaleString() : "0"}
              </div>
              <div className="relative text-xs text-white/85 font-medium mt-0.5">บาท / ปี</div>
              <div className="relative mt-4 pt-3 text-xs text-white/80 w-full" style={{ borderTop: "1px solid rgba(255,255,255,0.3)" }}>
                งบประมาณที่ใช้: {k.budget ? `${Number(k.budget).toLocaleString()} บาท` : "-"}
              </div>
            </div>
          </div>
        </div>

        {k.categories && k.categories.length > 0 && (
          <div className="flex items-center gap-2 flex-wrap">
            <span className="text-xs" style={{ color: SUB }}>หมวดหมู่:</span>
            <CategoryTags categories={k.categories} />
          </div>
        )}
      </div>
    </div>
  );
}

function Table({ rows, actions, showMissing }) {
  return (
    <div className="rounded-xl border overflow-hidden" style={{ borderColor: "#E2E6EB" }}>
      <table className="w-full text-sm">
        <thead>
          <tr style={{ background: "#FAFBFC" }}>
            <th className="text-left px-4 py-2.5 font-medium" style={{ color: SUB }}>เลขที่ / วันที่</th>
            <th className="text-left px-4 py-2.5 font-medium" style={{ color: SUB }}>เรื่อง</th>
            <th className="text-left px-4 py-2.5 font-medium" style={{ color: SUB }}>หน่วยงาน</th>
            <th className="text-left px-4 py-2.5 font-medium" style={{ color: SUB }}>หมวดหมู่</th>
            <th className="text-left px-4 py-2.5 font-medium" style={{ color: SUB }}>สถานะ</th>
            <th className="text-right px-4 py-2.5 font-medium" style={{ color: SUB }}>จัดการ</th>
          </tr>
        </thead>
        <tbody>
          {rows.length === 0 && (
            <tr><td colSpan={6} className="text-center py-8 text-sm" style={{ color: SUB }}>ไม่มีรายการ</td></tr>
          )}
          {rows.map((k) => (
            <tr key={k.id} className="border-t" style={{ borderColor: "#EEF0F3" }}>
              <td className="px-4 py-3 whitespace-nowrap" style={{ color: SUB }}>{k.registeredNo}<br /><span className="text-xs">{k.date}</span></td>
              <td className="px-4 py-3 font-medium max-w-[220px]" style={{ color: INK }}>
                <div className="flex items-center gap-1.5">
                  <span>{k.title}</span>
                  {k.attachments && k.attachments.length > 0 && (
                    <span title={`มีไฟล์แนบ ${k.attachments.length} ไฟล์`} className="inline-flex items-center flex-shrink-0" style={{ color: SUB }}>
                      <Paperclip size={13} />
                    </span>
                  )}
                </div>
                {showMissing && (!k.images?.before || !k.images?.after) && (
                  <span className="inline-flex items-center gap-1 mt-1 text-[10px] px-1.5 py-0.5 rounded-full font-normal" style={{ background: "#FBEAE8", color: "#8B2E24" }}>
                    <ImageOff size={10} /> {!k.images?.before && !k.images?.after ? "ยังไม่มีรูปก่อน-หลัง" : !k.images?.before ? "ยังไม่มีรูปก่อน" : "ยังไม่มีรูปหลัง"}
                  </span>
                )}
              </td>
              <td className="px-4 py-3" style={{ color: SUB }}>{k.department}</td>
              <td className="px-4 py-3"><CategoryTags categories={k.categories} /></td>
              <td className="px-4 py-3"><StatusBadge status={k.status} /></td>
              <td className="px-4 py-3">
                <div className="flex justify-end gap-1.5">{actions(k)}</div>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function IconBtn({ icon: Icon, label, onClick, color, filled }) {
  return (
    <button
      onClick={onClick}
      title={label}
      className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium border transition"
      style={filled ? { background: color, borderColor: color, color: "white" } : { borderColor: "#D8DDE5", color }}
    >
      <Icon size={13} /> {label}
    </button>
  );
}

function ReviewModal({ k, onClose, onDecide, roleLabel, allowCategorize }) {
  const [comment, setComment] = useState("");
  const [cats, setCats] = useState(k.categories);
  const [activityType, setActivityType] = useState(k.activityType);
  const [kaizenType, setKaizenType] = useState(k.kaizenType);
  const [checklist, setChecklist] = useState(k.checklist || CHECKLIST_ITEMS.map(() => false));
  const toggleCat = (c) => setCats((s) => (s.includes(c) ? s.filter((x) => x !== c) : [...s, c]));
  const toggleCheck = (i) => setChecklist((s) => s.map((v, idx) => (idx === i ? !v : v)));
  const checklistDone = !allowCategorize || checklist.every(Boolean);

  const returnOptions =
    k.status === "pending_director" ? ["employee", "supervisor", "admin"] :
    k.status === "pending_admin" ? ["employee", "supervisor"] :
    ["employee"];
  const [returnTo, setReturnTo] = useState("employee");

  return (
    <Modal title={`ตรวจสอบ Kaizen · ${roleLabel}`} onClose={onClose} wide>
      <div className="space-y-5">
        <DetailView k={k} />
        {allowCategorize && (
          <>
            <div className="rounded-xl border p-4" style={{ borderColor: "#E2E6EB" }}>
              <div className="font-medium text-sm mb-3" style={{ color: INK }}>จัดประเภทข้อมูล (Admin)</div>
              <div className="flex flex-wrap gap-2 mb-4">
                {CATEGORY_OPTIONS.map((c) => (
                  <button
                    key={c}
                    onClick={() => toggleCat(c)}
                    className="text-xs px-3 py-1.5 rounded-full border"
                    style={cats.includes(c) ? { background: NAVY, borderColor: NAVY, color: "white" } : { borderColor: "#D8DDE5", color: SUB }}
                  >
                    {c}
                  </button>
                ))}
              </div>
              <div className="grid grid-cols-2 gap-4">
                <Field label="ประเภทกิจกรรม">
                  <select className={inputCls} style={inputStyle} value={activityType} onChange={(e) => setActivityType(e.target.value)}>
                    {ACTIVITY_OPTIONS.map((o) => <option key={o}>{o}</option>)}
                  </select>
                </Field>
                <Field label="ประเภท Kaizen">
                  <select className={inputCls} style={inputStyle} value={kaizenType} onChange={(e) => setKaizenType(e.target.value)}>
                    {KAIZEN_TYPE_OPTIONS.map((o) => <option key={o}>{o}</option>)}
                  </select>
                </Field>
              </div>
            </div>
            <div className="rounded-xl border p-4" style={{ borderColor: "#E2E6EB" }}>
              <div className="font-medium text-sm mb-3" style={{ color: INK }}>ส่วนที่ 3 · ตรวจสอบข้อมูลการลงทะเบียน (ต้องติ๊กครบก่อนอนุมัติ)</div>
              <div className="space-y-2">
                {CHECKLIST_ITEMS.map((item, i) => (
                  <label key={i} className="flex items-start gap-2 text-sm cursor-pointer" style={{ color: INK }}>
                    <input type="checkbox" className="mt-0.5" checked={checklist[i]} onChange={() => toggleCheck(i)} />
                    <span>{item}</span>
                  </label>
                ))}
              </div>
              {!checklistDone && (
                <div className="flex items-center gap-1.5 text-xs mt-3" style={{ color: GOLD }}>
                  <AlertTriangle size={13} /> ต้องติ๊กครบทั้ง 4 ข้อ ก่อนกดอนุมัติส่งต่อผู้อำนวยการ
                </div>
              )}
            </div>
          </>
        )}
        <Field label="ความเห็น / เหตุผล (กรอกหากตีกลับ)">
          <TextArea value={comment} onChange={(e) => setComment(e.target.value)} placeholder="ระบุความเห็นประกอบการอนุมัติหรือตีกลับ" />
        </Field>
        {returnOptions.length > 1 && (
          <Field label="หากตีกลับ ส่งกลับไปยัง (ค่าเริ่มต้น: พนักงานผู้ส่งเรื่อง)">
            <select className={inputCls} style={inputStyle} value={returnTo} onChange={(e) => setReturnTo(e.target.value)}>
              {returnOptions.map((o) => <option key={o} value={o}>{TARGET_LABEL[o]}</option>)}
            </select>
          </Field>
        )}
        <div className="flex justify-end gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium border" style={{ borderColor: "#D8DDE5", color: SUB }}>ปิด</button>
          <button
            onClick={() => onDecide("reject", comment, { categories: cats, activityType, kaizenType, checklist }, returnTo)}
            className="px-4 py-2 rounded-lg text-sm font-medium flex items-center gap-1.5"
            style={{ background: "#F9E9E7", color: RED }}
          >
            <XCircle size={15} /> ตีกลับ{returnOptions.length > 1 ? ` (ไปยัง ${TARGET_LABEL[returnTo]})` : ""}
          </button>
          <button
            disabled={!checklistDone}
            onClick={() => onDecide("approve", comment, { categories: cats, activityType, kaizenType, checklist })}
            className="px-4 py-2 rounded-lg text-sm font-medium text-white flex items-center gap-1.5 disabled:opacity-40"
            style={{ background: GREEN }}
          >
            <CheckCircle2 size={15} /> อนุมัติ
          </button>
        </div>
      </div>
    </Modal>
  );
}

function DepartmentManager({ departments, onChange }) {
  const [newDept, setNewDept] = useState("");
  const updateAt = (i, val) => onChange(departments.map((d, idx) => (idx === i ? val : d)));
  const removeAt = (i) => onChange(departments.filter((_, idx) => idx !== i));
  const addNew = () => {
    const v = newDept.trim();
    if (v && !departments.includes(v)) {
      onChange([...departments, v]);
      setNewDept("");
    }
  };
  return (
    <div className="rounded-xl border p-4" style={{ borderColor: "#E2E6EB" }}>
      <div className="font-medium text-sm mb-3 flex items-center gap-1.5" style={{ color: INK }}>
        <Building2 size={15} /> จัดการรายชื่อหน่วยงาน
      </div>
      <div className="space-y-2 mb-3">
        {departments.map((d, i) => (
          <div key={i} className="flex items-center gap-2">
            <span className="text-xs w-5" style={{ color: SUB }}>{i + 1}</span>
            <input value={d} onChange={(e) => updateAt(i, e.target.value)} className="flex-1 rounded-lg border px-3 py-1.5 text-sm" style={inputStyle} />
            <button onClick={() => removeAt(i)} title="ลบหน่วยงานนี้" className="p-1.5 rounded-lg border" style={{ borderColor: "#D8DDE5", color: RED }}>
              <X size={14} />
            </button>
          </div>
        ))}
      </div>
      <div className="flex items-center gap-2">
        <input value={newDept} onChange={(e) => setNewDept(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addNew()} placeholder="เพิ่มหน่วยงานใหม่ เช่น บรรจุ 4" className="flex-1 rounded-lg border px-3 py-1.5 text-sm" style={inputStyle} />
        <button onClick={addNew} className="px-3 py-1.5 rounded-lg text-sm font-medium text-white flex items-center gap-1" style={{ background: NAVY }}>
          <PlusCircle size={14} /> เพิ่ม
        </button>
      </div>
      <div className="text-[11px] mt-2" style={{ color: "#8A93A3" }}>
        * รายชื่อหน่วยงานนี้ใช้เป็นตัวเลือกในฟอร์มส่งเรื่องและตัวกรองแดชบอร์ด การแก้ไข/ลบที่นี่เป็นการสาธิตเท่านั้น (ไม่ย้อนปรับปรุงรายการ Kaizen หรือบัญชีผู้ใช้ที่มีอยู่แล้ว)
      </div>
    </div>
  );
}

function StatCard({ label, value, sub, color }) {
  return (
    <div className="rounded-xl border p-4" style={{ borderColor: "#E2E6EB" }}>
      <div className="text-xs font-medium" style={{ color: SUB }}>{label}</div>
      <div className="text-2xl font-semibold mt-1" style={{ color: color || INK }}>{value}</div>
      {sub && <div className="text-xs mt-1" style={{ color: SUB }}>{sub}</div>}
    </div>
  );
}

const THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

// ---------- Dashboard color tokens (สีฟ้า-น้ำเงินโทนโปร) ----------
const DB_DEEP = "#0A2A57";
const DB_BLUE = "#1D5FC9";
const DB_BLUE2 = "#3F8FE0";
const DB_PALE = "#EAF3FC";
const DB_BORDER = "#D3E4F7";

function FilterableKaizenTable({ rows, onRowClick, onPresent }) {
  const [search, setSearch] = useState("");
  const [colDept, setColDept] = useState("all");
  const [colCategory, setColCategory] = useState("all");
  const [colStatus, setColStatus] = useState("all");

  const deptOptions = useMemo(() => Array.from(new Set(rows.map((r) => r.department))).sort(), [rows]);

  const filteredRows = rows.filter((k) => {
    if (search.trim()) {
      const q = search.trim().toLowerCase();
      if (!(`${k.title} ${k.registeredNo}`.toLowerCase().includes(q))) return false;
    }
    if (colDept !== "all" && k.department !== colDept) return false;
    if (colCategory !== "all" && !k.categories.includes(colCategory)) return false;
    if (colStatus !== "all" && k.status !== colStatus) return false;
    return true;
  });

  const thBase = { color: "#EAF3FC", fontWeight: 600 };
  const selCls = "w-full text-[11px] rounded-md px-1.5 py-1 mt-1 outline-none";
  const selStyle = { background: "rgba(255,255,255,0.15)", color: "white", border: "1px solid rgba(255,255,255,0.3)" };

  return (
    <div className="rounded-xl border overflow-hidden" style={{ borderColor: DB_BORDER }}>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr style={{ background: DB_DEEP }}>
              <th className="text-left px-4 py-2.5 align-top" style={thBase}>เลขที่ / วันที่</th>
              <th className="text-left px-4 py-2.5 align-top" style={thBase}>
                เรื่อง
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="ค้นหา..."
                  className={selCls}
                  style={{ ...selStyle, fontWeight: 400 }}
                />
              </th>
              <th className="text-left px-4 py-2.5 align-top" style={thBase}>
                หน่วยงาน
                <select value={colDept} onChange={(e) => setColDept(e.target.value)} className={selCls} style={selStyle}>
                  <option value="all" style={{ color: INK }}>ทั้งหมด</option>
                  {deptOptions.map((d) => <option key={d} value={d} style={{ color: INK }}>{d}</option>)}
                </select>
              </th>
              <th className="text-left px-4 py-2.5 align-top" style={thBase}>
                หมวดหมู่
                <select value={colCategory} onChange={(e) => setColCategory(e.target.value)} className={selCls} style={selStyle}>
                  <option value="all" style={{ color: INK }}>ทั้งหมด</option>
                  {CATEGORY_OPTIONS.map((c) => <option key={c} value={c} style={{ color: INK }}>{c}</option>)}
                </select>
              </th>
              <th className="text-left px-4 py-2.5 align-top" style={thBase}>
                สถานะ
                <select value={colStatus} onChange={(e) => setColStatus(e.target.value)} className={selCls} style={selStyle}>
                  <option value="all" style={{ color: INK }}>ทั้งหมด</option>
                  {Object.entries(STATUS_META).map(([k, v]) => <option key={k} value={k} style={{ color: INK }}>{v.label}</option>)}
                </select>
              </th>
              <th className="text-right px-4 py-2.5 align-top" style={thBase}>จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {filteredRows.length === 0 && (
              <tr><td colSpan={6} className="text-center py-8 text-sm" style={{ color: SUB }}>ไม่มีรายการตามเงื่อนไขที่กรอง</td></tr>
            )}
            {filteredRows.map((k) => (
              <tr
                key={k.id}
                onClick={() => onRowClick(k)}
                className="border-t cursor-pointer transition"
                style={{ borderColor: "#EEF0F3" }}
                onMouseEnter={(e) => (e.currentTarget.style.background = DB_PALE)}
                onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
              >
                <td className="px-4 py-3 whitespace-nowrap" style={{ color: SUB }}>{k.registeredNo}<br /><span className="text-xs">{k.date}</span></td>
                <td className="px-4 py-3 font-medium max-w-[220px]" style={{ color: INK }}>
                  <div className="flex items-center gap-1.5">
                    <span>{k.title}</span>
                    {k.attachments && k.attachments.length > 0 && (
                      <span title={`มีไฟล์แนบ ${k.attachments.length} ไฟล์`} className="inline-flex items-center flex-shrink-0" style={{ color: SUB }}>
                        <Paperclip size={13} />
                      </span>
                    )}
                  </div>
                </td>
                <td className="px-4 py-3" style={{ color: SUB }}>{k.department}</td>
                <td className="px-4 py-3"><CategoryTags categories={k.categories} /></td>
                <td className="px-4 py-3"><StatusBadge status={k.status} /></td>
                <td className="px-4 py-3 text-right">
                  <div className="flex items-center justify-end gap-1.5">
                    <button
                      onClick={(e) => { e.stopPropagation(); onPresent(k); }}
                      title="เปิดนำเสนอ"
                      className="flex items-center gap-1 px-2 py-1 rounded-lg text-xs font-medium border"
                      style={{ borderColor: "#D8DDE5", color: DB_BLUE }}
                    >
                      <Presentation size={13} /> นำเสนอ
                    </button>
                    <ChevronRight size={16} style={{ color: DB_BLUE }} />
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function Dashboard({ data, departments, onView, onPresent }) {
  const [yearFilter, setYearFilter] = useState("all");
  const [monthFilter, setMonthFilter] = useState("all");
  const [sectionFilter, setSectionFilter] = useState("all");
  const [deptFilter, setDeptFilter] = useState("all");
  const [catFilter, setCatFilter] = useState("all");
  const [activityFilter, setActivityFilter] = useState("all");
  const [kaizenTypeFilter, setKaizenTypeFilter] = useState("all");

  const yearOptions = useMemo(() => Array.from(new Set(data.map((k) => k.date.slice(0, 4)))).sort().reverse(), [data]);

  // เมื่อเลือกฝ่ายแล้ว รายการหน่วยงานใน Unit จะกรองให้เหลือเฉพาะหน่วยงานในฝ่ายนั้น
  const unitOptions = useMemo(() => {
    const all = departments && departments.length ? departments : DEFAULT_DEPARTMENTS;
    if (sectionFilter === "all") return all;
    return all.filter((d) => SECTION_MAP[d] === sectionFilter);
  }, [departments, sectionFilter]);

  const handleSectionChange = (val) => {
    setSectionFilter(val);
    if (deptFilter !== "all" && SECTION_MAP[deptFilter] !== val && val !== "all") setDeptFilter("all");
  };

  const filtered = data.filter((k) => {
    const y = k.date.slice(0, 4);
    const m = k.date.slice(5, 7);
    if (yearFilter !== "all" && y !== yearFilter) return false;
    if (monthFilter !== "all" && m !== monthFilter) return false;
    if (sectionFilter !== "all" && SECTION_MAP[k.department] !== sectionFilter) return false;
    if (deptFilter !== "all" && k.department !== deptFilter) return false;
    if (catFilter !== "all" && !k.categories.includes(catFilter)) return false;
    if (activityFilter !== "all" && k.activityType !== activityFilter) return false;
    if (kaizenTypeFilter !== "all" && k.kaizenType !== kaizenTypeFilter) return false;
    return true;
  });

  const costSavingBySection = useMemo(() => {
    const m = {};
    filtered.forEach((k) => {
      const v = parseFloat(k.costSaving) || 0;
      const section = SECTION_MAP[k.department] || "อื่นๆ";
      if (v > 0) m[section] = (m[section] || 0) + v;
    });
    return Object.entries(m).sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value }));
  }, [filtered]);

  const countByDept = useMemo(() => {
    const m = {};
    filtered.forEach((k) => (m[k.department] = (m[k.department] || 0) + 1));
    return Object.entries(m).sort((a, b) => b[1] - a[1]).map(([name, count]) => ({ name, count }));
  }, [filtered]);

  const countBySection = useMemo(() => {
    const m = {};
    filtered.forEach((k) => {
      const section = SECTION_MAP[k.department] || "อื่นๆ";
      m[section] = (m[section] || 0) + 1;
    });
    return Object.entries(m).sort((a, b) => b[1] - a[1]).map(([name, count]) => ({ name, count }));
  }, [filtered]);

  const byMonth = useMemo(() => {
    const m = {};
    filtered.forEach((k) => {
      const key = k.date.slice(0, 7);
      m[key] = (m[key] || 0) + 1;
    });
    return Object.entries(m).sort().map(([key, count]) => {
      const [y, mo] = key.split("-");
      return { name: `${THAI_MONTHS[parseInt(mo, 10) - 1]} ${y.slice(2)}`, count };
    });
  }, [filtered]);

  const total = filtered.length;
  const kaizenProjectCount = filtered.filter((k) => k.kaizenType === "Kaizen Project").length;
  const crossFunctionalCount = filtered.filter((k) => k.kaizenType === "Cross functional Kaizen").length;
  const totalCostSaving = filtered.reduce((sum, k) => sum + (parseFloat(k.costSaving) || 0), 0);

  const selectCls = "text-sm rounded-lg border px-2.5 py-1.5";

  const handleExport = () => {
    const rows = filtered.map((k) => ({
      "เลขที่": k.registeredNo,
      "วันที่": k.date,
      "เรื่อง": k.title,
      "ทีม": k.team || "",
      "หน่วยงาน (Unit)": k.department,
      "ฝ่าย (Section)": SECTION_MAP[k.department] || "",
      "หมวดหมู่": (k.categories || []).join(", "),
      "ประเภทกิจกรรม": k.activityType,
      "ประเภท Kaizen": k.kaizenType,
      "สถานะ": STATUS_META[k.status]?.label || k.status,
      "งบประมาณ (บาท)": k.budget || "",
      "Cost Saving (บาท/ปี)": k.costSaving || "",
      "ผลที่ได้รับ": k.result || "",
      "ผู้ส่งเรื่อง": k.submittedBy || "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [
      { wch: 12 }, { wch: 11 }, { wch: 32 }, { wch: 18 }, { wch: 20 }, { wch: 20 },
      { wch: 20 }, { wch: 14 }, { wch: 20 }, { wch: 18 }, { wch: 14 }, { wch: 16 }, { wch: 32 }, { wch: 20 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Kaizen");
    XLSX.writeFile(wb, `Kaizen_Export_${todayISO()}.xlsx`);
  };

  return (
    <div className="space-y-5">
      <div className="rounded-2xl p-4" style={{ background: `linear-gradient(135deg, ${DB_DEEP}, ${DB_BLUE})` }}>
        <div className="flex flex-wrap gap-2.5 items-center">
          <div className="flex items-center gap-1.5 text-xs font-medium text-white/90 pr-1"><Filter size={14} /> ตัวกรอง</div>
          <select className={selectCls} style={inputStyle} value={yearFilter} onChange={(e) => setYearFilter(e.target.value)}>
            <option value="all">ทุกปี</option>
            {yearOptions.map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
          <select className={selectCls} style={inputStyle} value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}>
            <option value="all">ทุกเดือน</option>
            {THAI_MONTHS.map((m, i) => <option key={m} value={String(i + 1).padStart(2, "0")}>{m}</option>)}
          </select>
          <select className={selectCls} style={inputStyle} value={sectionFilter} onChange={(e) => handleSectionChange(e.target.value)}>
            <option value="all">ทุกฝ่าย (Section)</option>
            {SECTION_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <select className={selectCls} style={inputStyle} value={deptFilter} onChange={(e) => setDeptFilter(e.target.value)}>
            <option value="all">ทุกหน่วยงาน (Unit)</option>
            {unitOptions.map((d) => <option key={d}>{d}</option>)}
          </select>
          <select className={selectCls} style={inputStyle} value={catFilter} onChange={(e) => setCatFilter(e.target.value)}>
            <option value="all">ทุกประเภทข้อมูล</option>
            {CATEGORY_OPTIONS.map((c) => <option key={c}>{c}</option>)}
          </select>
          <select className={selectCls} style={inputStyle} value={activityFilter} onChange={(e) => setActivityFilter(e.target.value)}>
            <option value="all">ทุกประเภทกิจกรรม</option>
            {ACTIVITY_OPTIONS.map((o) => <option key={o}>{o}</option>)}
          </select>
          <select className={selectCls} style={inputStyle} value={kaizenTypeFilter} onChange={(e) => setKaizenTypeFilter(e.target.value)}>
            <option value="all">ทุกประเภท Kaizen</option>
            {KAIZEN_TYPE_OPTIONS.map((o) => <option key={o}>{o}</option>)}
          </select>
          <div className="ml-auto flex items-center gap-3">
            <button
              onClick={handleExport}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-white shadow-sm"
              style={{ color: DB_DEEP }}
              title="Export ข้อมูลตามตัวกรองปัจจุบันเป็นไฟล์ Excel"
            >
              <Download size={14} /> Export Excel
            </button>
            <span className="text-xs text-white/70 whitespace-nowrap">อัปเดตล่าสุด: {todayISO()}</span>
          </div>
        </div>
      </div>

      <div id="dashboard-print-area" className="space-y-5">
      <div className="grid grid-cols-4 gap-4">
        <div className="rounded-xl p-4 text-white" style={{ background: `linear-gradient(135deg, ${DB_DEEP}, ${DB_BLUE})` }}>
          <div className="text-xs font-medium text-white/80">Kaizen รวมทั้งหมด</div>
          <div className="text-3xl font-semibold mt-1.5">{total}</div>
          <div className="text-xs mt-1 text-white/70">เรื่องทั้งหมดตามตัวกรอง</div>
        </div>
        <div className="rounded-xl p-4" style={{ background: DB_PALE, border: `1px solid ${DB_BORDER}` }}>
          <div className="text-xs font-medium" style={{ color: DB_BLUE }}>Kaizen Project</div>
          <div className="text-3xl font-semibold mt-1.5" style={{ color: DB_DEEP }}>{kaizenProjectCount}</div>
          <div className="text-xs mt-1" style={{ color: "#6E88A6" }}>โครงการปรับปรุงเดี่ยว</div>
        </div>
        <div className="rounded-xl p-4" style={{ background: DB_PALE, border: `1px solid ${DB_BORDER}` }}>
          <div className="text-xs font-medium" style={{ color: DB_BLUE }}>Cross functional Kaizen</div>
          <div className="text-3xl font-semibold mt-1.5" style={{ color: DB_DEEP }}>{crossFunctionalCount}</div>
          <div className="text-xs mt-1" style={{ color: "#6E88A6" }}>โครงการข้ามสายงาน</div>
        </div>
        <div
          className="relative rounded-xl p-4 overflow-hidden"
          style={{
            background: `linear-gradient(135deg, #A8710F, ${GOLD})`,
            boxShadow: "0 8px 20px -6px rgba(184, 121, 31, 0.55)",
          }}
        >
          <div
            className="absolute -right-4 -top-4 w-20 h-20 rounded-full"
            style={{ background: "rgba(255,255,255,0.15)" }}
          />
          <div
            className="absolute right-3 top-3 w-8 h-8 rounded-full flex items-center justify-center"
            style={{ background: "rgba(255,255,255,0.25)" }}
          >
            <Coins size={16} className="text-white" />
          </div>
          <div className="text-xs font-semibold text-white/90 relative">มูลค่า Cost Saving</div>
          <div className="text-4xl font-extrabold mt-1.5 text-white relative tracking-tight">
            {totalCostSaving.toLocaleString()}
          </div>
          <div className="text-xs mt-1 text-white/85 font-medium relative">บาท/ปี ตามตัวกรอง</div>
        </div>
      </div>

      <div className="grid grid-cols-2 gap-4">
        <div className="rounded-xl border p-4" style={{ borderColor: DB_BORDER }}>
          <div className="text-sm font-medium mb-3" style={{ color: DB_DEEP }}>Cost Saving ตามฝ่าย/Section (บาท/ปี)</div>
          <ResponsiveContainer width="100%" height={230}>
            <BarChart data={costSavingBySection}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF3FA" />
              <XAxis dataKey="name" tick={{ fontSize: 11 }} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11 }} />
              <Tooltip formatter={(v) => `${Number(v).toLocaleString()} บาท`} />
              <Bar dataKey="value" fill={DB_BLUE} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="rounded-xl border p-4" style={{ borderColor: DB_BORDER }}>
          <div className="text-sm font-medium mb-3" style={{ color: DB_DEEP }}>จำนวน Kaizen ตามฝ่าย/Section</div>
          <ResponsiveContainer width="100%" height={230}>
            <BarChart data={countBySection}>
              <CartesianGrid strokeDasharray="3 3" stroke="#EEF3FA" />
              <XAxis dataKey="name" tick={{ fontSize: 11 }} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11 }} />
              <Tooltip />
              <Bar dataKey="count" fill={DB_BLUE2} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="rounded-xl border p-4" style={{ borderColor: DB_BORDER }}>
        <div className="text-sm font-medium mb-3" style={{ color: DB_DEEP }}>จำนวน Kaizen ของแต่ละหน่วยงาน (Unit)</div>
        <ResponsiveContainer width="100%" height={230}>
          <BarChart data={countByDept}>
            <CartesianGrid strokeDasharray="3 3" stroke="#EEF3FA" />
            <XAxis dataKey="name" tick={{ fontSize: 11 }} />
            <YAxis allowDecimals={false} tick={{ fontSize: 11 }} />
            <Tooltip />
            <Bar dataKey="count" fill={DB_BLUE2} radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="rounded-xl border p-4" style={{ borderColor: DB_BORDER }}>
        <div className="text-sm font-medium mb-3 flex items-center gap-1.5" style={{ color: DB_DEEP }}><TrendingUp size={15} /> จำนวนการส่ง Kaizen รายเดือน</div>
        <ResponsiveContainer width="100%" height={230}>
          <BarChart data={byMonth}>
            <CartesianGrid strokeDasharray="3 3" stroke="#EEF3FA" />
            <XAxis dataKey="name" tick={{ fontSize: 11 }} />
            <YAxis allowDecimals={false} tick={{ fontSize: 11 }} />
            <Tooltip />
            <Bar dataKey="count" fill={DB_DEEP} radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <FilterableKaizenTable rows={filtered} onRowClick={onView} onPresent={onPresent} />
      </div>
    </div>
  );
}

// ---------- Kaizen Award (Admin คัดกรอง → กรรมการ 1-5 ให้คะแนนแยกกัน → ผอ. ยืนยันผล) ----------
const AUTO_SOURCES = {
  cost: "คำนวณจาก Cost Saving ตามช่วงมูลค่าที่ตั้งไว้ด้านล่าง",
  roi: "คำนวณจาก Cost Saving ÷ งบประมาณที่ใช้ (ไม่ใช้งบ = 5 คะแนน)",
  impact: "คำนวณจากหมวดหมู่ที่ Admin จัดไว้ (มีหมวดผลกระทบสูง = 5 คะแนน)",
  team: "Cross functional = 5, สมาชิก 5 คนขึ้นไป = 4, 3-4 คน = 3, อื่นๆ = 2",
};

// บัญชีกรรมการ (ชื่อผู้ใช้ตรงกับ MOCK_USERS ที่ role = "judge")
const JUDGE_SLOTS = [1, 2, 3, 4, 5].map((n) => ({ id: `j${n}`, userName: `กรรมการ ${n}` }));

const DEFAULT_AWARD = {
  roundName: "Kaizen Award 2026",
  year: "all",
  statuses: ["approved"],
  activity: "Kaizen",
  minCostSaving: 0,
  mainPrizes: 3,
  honorable: 2,
  costTiers: [100000, 50000, 20000, 5000],
  highImpactCats: ["Safety", "Food Safety", "Quality"],
  criteria: [
    { id: "cost", name: "มูลค่า Cost Saving", weight: 30, type: "auto", source: "cost" },
    { id: "impact", name: "ผลกระทบด้าน Safety / Quality / Food Safety", weight: 25, type: "auto", source: "impact" },
    { id: "innov", name: "ความคิดสร้างสรรค์ / ความแปลกใหม่", weight: 15, type: "manual" },
    { id: "roi", name: "ความคุ้มค่าการลงทุน (ROI)", weight: 10, type: "auto", source: "roi" },
    { id: "team", name: "การทำงานเป็นทีม / ข้ามสายงาน", weight: 10, type: "auto", source: "team" },
    { id: "yokoten", name: "ขยายผลไปหน่วยงานอื่นได้ (Yokoten)", weight: 10, type: "manual" },
  ],
  judgeCount: 5,
  judgeNames: {}, // { j1: "ชื่อจริงของกรรมการ (ไม่บังคับ)" }
  trimExtremes: true,
  shortlist: [], // id ของ Kaizen ที่ Admin ส่งให้กรรมการ
  sentAt: null,
  sentBy: null,
  scores: {}, // { j1: { [kaizenId]: { [criterionId]: 1..5, __note: "" }, __meta: { submittedAt } } }
  decision: null,
};

function normalizeAward(a) {
  const base = { ...DEFAULT_AWARD, ...(a || {}) };
  delete base.judges; // รุ่นก่อนหน้า (Admin/ผอ. เป็นกรรมการ) ไม่ใช้แล้ว
  base.judgeCount = Math.min(5, Math.max(1, toNum(base.judgeCount) || 5));
  base.judgeNames = base.judgeNames || {};
  base.shortlist = Array.isArray(base.shortlist) ? base.shortlist : [];
  base.scores = base.scores || {};
  return base;
}

const activeJudges = (cfg) =>
  JUDGE_SLOTS.slice(0, cfg.judgeCount).map((j) => {
    const real = String(cfg.judgeNames?.[j.id] || "").trim();
    return { ...j, name: real ? `${j.userName} · ${real}` : j.userName };
  });

const AWARD_TIERS = [
  { label: "รางวัลชนะเลิศ", short: "ชนะเลิศ", color: "#B07D1A", bg: "#FDF3E1", border: "#E8C77E" },
  { label: "รองชนะเลิศอันดับ 1", short: "รองฯ 1", color: "#5F6F84", bg: "#EEF1F5", border: "#C3CCD8" },
  { label: "รองชนะเลิศอันดับ 2", short: "รองฯ 2", color: "#94592F", bg: "#F7ECE3", border: "#DDB89A" },
];
const HONOR_TIER = { label: "รางวัลชมเชย", short: "ชมเชย", color: GREEN, bg: "#E9F5EF", border: "#A9D6BF" };
const CRIT_COLORS = ["#1D5FC9", "#D99A2B", "#7A5AE0", "#2F7D5C", "#D9577E", "#2BA3B8", "#8A6D3B", "#B23A2E"];
const JUDGE_PURPLE = "#6A4BC9";

function toNum(v) {
  const n = parseFloat(String(v ?? "").replace(/,/g, ""));
  return isNaN(n) ? 0 : n;
}
const fmtScore = (v) => (Number.isInteger(v) ? String(v) : v.toFixed(1));
const namedMemberCount = (k) => (k.members || []).filter((m) => m && String(m.name || "").trim()).length;
const byRegNo = (a, b) => String(a.registeredNo).localeCompare(String(b.registeredNo));

function autoScore(source, k, cfg) {
  const cs = toNum(k.costSaving);
  if (source === "cost") {
    const t = cfg.costTiers;
    if (cs <= 0) return 0;
    if (cs >= t[0]) return 5;
    if (cs >= t[1]) return 4;
    if (cs >= t[2]) return 3;
    if (cs >= t[3]) return 2;
    return 1;
  }
  if (source === "roi") {
    if (cs <= 0) return 0;
    const b = toNum(k.budget);
    if (b <= 0) return 5;
    const r = cs / b;
    if (r >= 10) return 5;
    if (r >= 5) return 4;
    if (r >= 2) return 3;
    if (r >= 1) return 2;
    return 1;
  }
  if (source === "impact") {
    const cats = k.categories || [];
    if (cats.some((c) => cfg.highImpactCats.includes(c))) return 5;
    return cats.length ? 3 : 1;
  }
  if (source === "team") {
    if (k.kaizenType === "Cross functional Kaizen") return 5;
    const n = namedMemberCount(k);
    if (n >= 5) return 4;
    if (n >= 3) return 3;
    return 2;
  }
  return 0;
}

function autoReason(source, k) {
  const cs = toNum(k.costSaving);
  if (source === "cost") return cs ? `${cs.toLocaleString()} บาท/ปี` : "ไม่มี Cost Saving";
  if (source === "roi") {
    const b = toNum(k.budget);
    if (cs <= 0) return "ไม่มี Cost Saving";
    return b <= 0 ? "ไม่ใช้งบประมาณ" : `คืนทุน ${(cs / b).toFixed(1)} เท่า`;
  }
  if (source === "impact") return (k.categories || []).join(", ") || "ยังไม่จัดหมวด";
  if (source === "team") return k.kaizenType === "Cross functional Kaizen" ? "Cross functional" : `สมาชิก ${namedMemberCount(k)} คน`;
  return "";
}

function eligibleForAward(kaizens, cfg) {
  return kaizens.filter(
    (k) =>
      cfg.statuses.includes(k.status) &&
      (cfg.activity === "all" || (k.activityType || "Kaizen") === cfg.activity) &&
      toNum(k.costSaving) >= toNum(cfg.minCostSaving) &&
      (cfg.year === "all" || String(k.date || "").startsWith(cfg.year))
  );
}
const shortlisted = (kaizens, cfg) => kaizens.filter((k) => cfg.shortlist.includes(k.id));

/** คะแนนคัดกรองเบื้องต้น (ใช้เฉพาะเกณฑ์อัตโนมัติ) คิดเป็น % */
function prescreen(k, cfg) {
  const auto = cfg.criteria.filter((c) => c.type === "auto");
  const w = auto.reduce((s, c) => s + toNum(c.weight), 0);
  if (!w) return 0;
  const pts = auto.reduce((s, c) => s + (toNum(c.weight) * autoScore(c.source, k, cfg)) / 5, 0);
  return Math.round((pts / w) * 100);
}

/** คะแนนของกรรมการแต่ละคนในเกณฑ์หนึ่ง + ค่าเฉลี่ย (ตัดสูงสุด-ต่ำสุดเมื่อมีกรรมการ 4 คนขึ้นไป ถ้าเปิดใช้) */
function manualStats(k, c, cfg) {
  const byJudge = activeJudges(cfg).map((j) => ({ judge: j, v: cfg.scores?.[j.id]?.[k.id]?.[c.id] }));
  const vals = byJudge.filter((x) => typeof x.v === "number").map((x) => x.v);
  let used = vals;
  const trimmed = cfg.trimExtremes && vals.length >= 4;
  if (trimmed) used = [...vals].sort((a, b) => a - b).slice(1, -1);
  const avg = used.length ? used.reduce((s, v) => s + v, 0) / used.length : null;
  return { byJudge, vals, avg, trimmed, spread: vals.length ? Math.max(...vals) - Math.min(...vals) : 0 };
}

function scoreKaizen(k, cfg) {
  const totalW = cfg.criteria.reduce((s, c) => s + toNum(c.weight), 0) || 1;
  const manual = cfg.criteria.filter((c) => c.type === "manual");
  let total = 0;
  const parts = cfg.criteria.map((c, i) => {
    let s;
    let stats = null;
    if (c.type === "auto") s = autoScore(c.source, k, cfg);
    else {
      stats = manualStats(k, c, cfg);
      s = stats.avg ?? 0;
    }
    const pts = ((toNum(c.weight) / totalW) * 100 * s) / 5;
    total += pts;
    return { id: c.id, name: c.name, type: c.type, score: Math.round(s * 10) / 10, pts, stats, color: CRIT_COLORS[i % CRIT_COLORS.length] };
  });
  const pendingJudges = activeJudges(cfg).filter((j) => manual.some((c) => typeof cfg.scores?.[j.id]?.[k.id]?.[c.id] !== "number")).length;
  return { k, total: Math.round(total * 10) / 10, pendingJudges, parts };
}

function rankAward(kaizens, cfg) {
  return shortlisted(kaizens, cfg)
    .map((k) => scoreKaizen(k, cfg))
    .sort((a, b) => b.total - a.total || toNum(b.k.costSaving) - toNum(a.k.costSaving))
    .map((r, i) => ({
      ...r,
      rank: i + 1,
      tier: i < cfg.mainPrizes ? AWARD_TIERS[i] : i < cfg.mainPrizes + cfg.honorable ? HONOR_TIER : null,
    }));
}

function judgeProgress(cfg, items) {
  const manual = cfg.criteria.filter((c) => c.type === "manual");
  const total = items.length * manual.length;
  return activeJudges(cfg).map((j) => {
    let done = 0;
    items.forEach((k) => manual.forEach((c) => typeof cfg.scores?.[j.id]?.[k.id]?.[c.id] === "number" && done++));
    return { judge: j, done, total, submittedAt: cfg.scores?.[j.id]?.__meta?.submittedAt || null };
  });
}

function buildAwardInsights(results, cfg, progress) {
  if (!results.length) return [];
  const out = [];
  const [first, second] = results;
  const strengths = (r) => [...r.parts].sort((a, b) => b.pts - a.pts).slice(0, 2).map((p) => p.name);
  out.push(`ผู้ชนะเลิศคือ "${first.k.title}" (${first.k.department}) ได้ ${first.total} คะแนน จุดแข็งคือ ${strengths(first).join(" และ ")}`);
  if (second) {
    const gap = Math.round((first.total - second.total) * 10) / 10;
    out.push(
      gap < 3
        ? `อันดับ 1 และ 2 ห่างกันเพียง ${gap} คะแนน ถือว่าสูสีมาก แนะนำให้พิจารณาเพิ่มเติม หรือเชิญทั้งสองทีมนำเสนอต่อคณะกรรมการก่อนยืนยันผล`
        : `อันดับ 1 นำอันดับ 2 อยู่ ${gap} คะแนน ผลการตัดสินค่อนข้างชัดเจน`
    );
  }
  const disagree = [];
  results.forEach((r) => r.parts.forEach((p) => p.stats && p.stats.vals.length >= 2 && p.stats.spread >= 3 && disagree.push({ r, p })));
  disagree
    .sort((a, b) => b.p.stats.spread - a.p.stats.spread)
    .slice(0, 2)
    .forEach(({ r, p }) =>
      out.push(`กรรมการให้คะแนนเรื่อง "${r.k.title}" ในเกณฑ์ "${p.name}" ต่างกันมาก (ต่ำสุด ${Math.min(...p.stats.vals)} สูงสุด ${Math.max(...p.stats.vals)}) ควรหารือร่วมกันก่อนยืนยันผล`)
    );
  if (cfg.trimExtremes && cfg.judgeCount >= 4) out.push("ใช้วิธีตัดคะแนนสูงสุดและต่ำสุดออกก่อนเฉลี่ย เพื่อลดผลจากอคติของกรรมการรายบุคคล");
  const winners = results.filter((r) => r.tier);
  out.push(`เรื่องที่ได้รับรางวัลทั้ง ${winners.length} เรื่อง สร้าง Cost Saving รวม ${winners.reduce((s, r) => s + toNum(r.k.costSaving), 0).toLocaleString()} บาท/ปี`);
  const notSubmitted = progress.filter((p) => !p.submittedAt);
  if (notSubmitted.length) out.push(`กรรมการที่ยังไม่ได้กดส่งคะแนน: ${notSubmitted.map((p) => `${p.judge.userName} (${p.done}/${p.total})`).join(", ")} ผลตอนนี้จึงยังไม่ใช่ผลสุดท้าย`);
  return out;
}

function TierBadge({ tier, small }) {
  if (!tier) return <span className="text-xs" style={{ color: SUB }}>-</span>;
  return (
    <span
      className={`inline-flex items-center gap-1 rounded-full font-semibold whitespace-nowrap ${small ? "text-[11px] px-2 py-0.5" : "text-xs px-2.5 py-1"}`}
      style={{ background: tier.bg, color: tier.color, border: `1px solid ${tier.border}` }}
    >
      {tier === HONOR_TIER ? <Medal size={small ? 11 : 13} /> : <Trophy size={small ? 11 : 13} />}
      {small ? tier.short : tier.label}
    </span>
  );
}

function ContributionBar({ parts, height = 10 }) {
  return (
    <div className="w-full rounded-full overflow-hidden flex" style={{ height, background: "#EEF0F3" }}>
      {parts.map((p) => (
        <div key={p.id} title={`${p.name}: ${fmtScore(p.score)}/5 (${p.pts.toFixed(1)} คะแนน)`} style={{ width: `${p.pts}%`, background: p.color }} />
      ))}
    </div>
  );
}

function ProgressBar({ done, total, color = JUDGE_PURPLE }) {
  const pct = total ? Math.round((done / total) * 100) : 0;
  return (
    <div className="w-full h-1.5 rounded-full overflow-hidden" style={{ background: "#ECE8F8" }}>
      <div style={{ width: `${pct}%`, height: "100%", background: done >= total && total ? GREEN : color }} />
    </div>
  );
}

function AwardSection({ title, sub, children, right }) {
  return (
    <div className="bg-white rounded-xl border p-4 md:p-5" style={{ borderColor: "#E3E7ED" }}>
      <div className="flex items-start justify-between gap-3 mb-3 flex-wrap">
        <div>
          <div className="text-sm font-semibold" style={{ color: INK }}>{title}</div>
          {sub && <div className="text-xs mt-0.5" style={{ color: SUB }}>{sub}</div>}
        </div>
        {right}
      </div>
      {children}
    </div>
  );
}

function AwardHeader({ cfg, onRename, editable, subtitle, syncMode, onRefresh }) {
  return (
    <div className="rounded-xl p-4 md:p-5 flex items-center justify-between gap-3 flex-wrap" style={{ background: `linear-gradient(135deg, ${NAVY_DEEP}, ${NAVY})` }}>
      <div className="flex items-center gap-3">
        <div className="w-11 h-11 rounded-xl flex items-center justify-center" style={{ background: GOLD }}>
          <Trophy size={22} color={NAVY_DEEP} />
        </div>
        <div>
          <input
            value={cfg.roundName}
            onChange={(e) => onRename && onRename(e.target.value)}
            disabled={!editable}
            className="bg-transparent text-white text-lg font-semibold outline-none border-b border-transparent focus:border-white/40"
          />
          <div className="text-xs" style={{ color: "#AFC0D6" }}>{subtitle}</div>
        </div>
      </div>
      <div className="flex items-center gap-2 flex-wrap">
        {syncMode === "sheet" && (
          <button onClick={onRefresh} className="text-[11px] px-2.5 py-1 rounded-full border text-white" style={{ borderColor: "#5D7699" }}>
            โหลดข้อมูลล่าสุด
          </button>
        )}
        <span className="text-[11px] px-2.5 py-1 rounded-full" style={syncMode === "sheet" ? { background: "#E9F5EF", color: GREEN } : { background: "#FDF3E1", color: "#8A5B14" }}>
          {syncMode === "sheet" ? "บันทึกลง Google Sheet อัตโนมัติ" : "โหมดทดลอง: ข้อมูลจะหายเมื่อรีเฟรชหน้า"}
        </span>
      </div>
    </div>
  );
}

// ================= หน้า Admin / ผอ. =================
function AwardPage({ kaizens, award, onChange, onScore, currentUser, onView, onPresent, onAnnounce, onSendToJudges, onRefresh, syncMode }) {
  const cfg = award;
  const isDirector = currentUser.role === "director";
  const [sub, setSub] = useState("criteria");
  const [confirming, setConfirming] = useState(false);
  const [selected, setSelected] = useState(() => new Set(cfg.shortlist));
  const [topN, setTopN] = useState(5);
  const [unlocking, setUnlocking] = useState(null);
  const [judgeFilter, setJudgeFilter] = useState("all");

  const locked = !!cfg.decision;
  const set = (patch) => onChange({ ...cfg, ...patch });
  const judges = activeJudges(cfg);
  const eligible = useMemo(() => eligibleForAward(kaizens, cfg).map((k) => ({ k, pre: prescreen(k, cfg) })).sort((a, b) => b.pre - a.pre || toNum(b.k.costSaving) - toNum(a.k.costSaving)), [kaizens, cfg]);
  const results = useMemo(() => rankAward(kaizens, cfg), [kaizens, cfg]);
  const items = useMemo(() => shortlisted(kaizens, cfg), [kaizens, cfg]);
  const progress = useMemo(() => judgeProgress(cfg, items), [cfg, items]);
  const insights = useMemo(() => buildAwardInsights(results, cfg, progress), [results, cfg, progress]);
  const totalW = cfg.criteria.reduce((s, c) => s + toNum(c.weight), 0);
  const manualCrit = cfg.criteria.filter((c) => c.type === "manual");
  const autoCrit = cfg.criteria.filter((c) => c.type === "auto");
  const submittedCount = progress.filter((p) => p.submittedAt).length;
  const years = useMemo(() => [...new Set(kaizens.map((k) => String(k.date || "").slice(0, 4)).filter((y) => /^\d{4}$/.test(y)))].sort().reverse(), [kaizens]);
  const shortlistChanged = selected.size !== cfg.shortlist.length || [...selected].some((id) => !cfg.shortlist.includes(id));

  // ---- criteria
  const setCrit = (id, patch) => set({ criteria: cfg.criteria.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  const removeCrit = (id) => cfg.criteria.length > 1 && set({ criteria: cfg.criteria.filter((c) => c.id !== id) });
  const addCrit = () => set({ criteria: [...cfg.criteria, { id: `m_${uid()}`, name: "เกณฑ์ใหม่ (กรรมการให้คะแนน)", weight: 5, type: "manual" }] });
  const normalize = () => {
    if (!totalW) return;
    const scaled = cfg.criteria.map((c) => ({ ...c, weight: Math.round((toNum(c.weight) / totalW) * 100) }));
    scaled[0].weight += 100 - scaled.reduce((s, c) => s + c.weight, 0);
    set({ criteria: scaled });
  };
  const toggleStatus = (s) => set({ statuses: cfg.statuses.includes(s) ? cfg.statuses.filter((x) => x !== s) : [...cfg.statuses, s] });
  const toggleImpact = (c) => set({ highImpactCats: cfg.highImpactCats.includes(c) ? cfg.highImpactCats.filter((x) => x !== c) : [...cfg.highImpactCats, c] });

  // ---- screening
  const toggleSel = (id) => setSelected((s) => {
    const n = new Set(s);
    n.has(id) ? n.delete(id) : n.add(id);
    return n;
  });
  const pickTop = () => setSelected(new Set(eligible.slice(0, Math.max(1, toNum(topN))).map((x) => x.k.id)));
  const sendShortlist = () => {
    const ids = eligible.filter((x) => selected.has(x.k.id)).map((x) => x.k.id);
    if (!ids.length) return;
    set({ shortlist: ids, sentAt: nowStamp(), sentBy: currentUser.name });
    onSendToJudges(ids.length);
    setSub("tracking");
  };

  const unlockJudge = (jid) => {
    const cur = cfg.scores[jid] || {};
    const { __meta, ...rest } = cur;
    onScore(jid, rest);
    setUnlocking(null);
  };

  const exportExcel = () => {
    const rows = results.map((r) => {
      const row = { อันดับ: r.rank, รางวัล: r.tier ? r.tier.label : "", เลขที่: r.k.registeredNo, เรื่อง: r.k.title, หน่วยงาน: r.k.department, ผู้ส่งเรื่อง: r.k.submittedBy, "Cost Saving (บาท/ปี)": toNum(r.k.costSaving) };
      r.parts.forEach((p) => (row[`${p.name} (${p.type === "manual" ? "เฉลี่ยกรรมการ" : "อัตโนมัติ"}/5)`] = p.score));
      row["คะแนนรวม (เต็ม 100)"] = r.total;
      return row;
    });
    const perJudge = [];
    judges.forEach((j) =>
      results.forEach((r) => {
        const row = { กรรมการ: j.name, เลขที่: r.k.registeredNo, เรื่อง: r.k.title };
        manualCrit.forEach((c) => (row[c.name] = cfg.scores?.[j.id]?.[r.k.id]?.[c.id] ?? ""));
        row["ความเห็น"] = cfg.scores?.[j.id]?.[r.k.id]?.__note || "";
        perJudge.push(row);
      })
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "ผลการตัดสิน");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(perJudge), "คะแนนรายกรรมการ");
    XLSX.writeFile(wb, `${cfg.roundName.replace(/\s+/g, "_")}_ผลการตัดสิน.xlsx`);
  };

  const confirmResult = () => {
    const winners = results.filter((r) => r.tier).map((r) => ({ id: r.k.id, registeredNo: r.k.registeredNo, title: r.k.title, department: r.k.department, submittedBy: r.k.submittedBy, tier: r.tier.label, total: r.total }));
    set({ decision: { confirmedBy: currentUser.name, confirmedAt: nowStamp(), winners } });
    onAnnounce(winners);
    setConfirming(false);
  };

  const disabledStyle = locked ? { opacity: 0.55, pointerEvents: "none" } : {};
  const SUBTABS = [
    { key: "criteria", label: "1. ตั้งเกณฑ์" },
    { key: "screen", label: `2. คัดกรอง & ส่งกรรมการ${cfg.shortlist.length ? ` (${cfg.shortlist.length})` : ""}` },
    { key: "tracking", label: `3. ติดตามกรรมการ (${submittedCount}/${judges.length})` },
    { key: "scores", label: "4. คะแนนกรรมการ" },
    { key: "result", label: "5. ผลการตัดสิน" },
  ];

  return (
    <div className="space-y-4">
      <AwardHeader
        cfg={cfg}
        editable={!locked}
        onRename={(v) => set({ roundName: v })}
        subtitle={`Admin คัดกรอง → ส่งกรรมการ ${judges.length} ท่านให้คะแนนแยกกัน → ผู้อำนวยการยืนยันผล`}
        syncMode={syncMode}
        onRefresh={onRefresh}
      />

      {locked && (
        <div className="rounded-xl p-4 border flex items-center justify-between gap-3 flex-wrap" style={{ background: "#FDF3E1", borderColor: "#E8C77E" }}>
          <div className="text-sm" style={{ color: "#6B4A10" }}>
            <strong>ประกาศผลแล้ว</strong> โดย {cfg.decision.confirmedBy} เมื่อ {cfg.decision.confirmedAt} · เกณฑ์และคะแนนถูกล็อกไว้
          </div>
          {isDirector && (
            <button onClick={() => set({ decision: null })} className="text-xs px-3 py-1.5 rounded-lg border font-medium" style={{ borderColor: "#C9962B", color: "#6B4A10" }}>
              ยกเลิกการประกาศ & เปิดแก้ไข
            </button>
          )}
        </div>
      )}

      <div className="flex gap-1 bg-white rounded-xl border p-1 w-fit flex-wrap" style={{ borderColor: "#E3E7ED" }}>
        {SUBTABS.map((t) => (
          <button key={t.key} onClick={() => setSub(t.key)} className="px-3.5 py-1.5 rounded-lg text-xs md:text-sm font-medium" style={sub === t.key ? { background: NAVY, color: "white" } : { color: SUB }}>
            {t.label}
          </button>
        ))}
      </div>

      {/* ---------------- 1. CRITERIA ---------------- */}
      {sub === "criteria" && (
        <div className="space-y-4" style={disabledStyle}>
          <AwardSection title="ผู้มีสิทธิ์เข้าประกวด" sub="กำหนดว่า Kaizen แบบไหนบ้างที่จะเข้าสู่การคัดกรอง" right={<span className="text-xs px-2.5 py-1 rounded-full font-semibold" style={{ background: "#EAF3FC", color: "#1D5FC9" }}>เข้าเกณฑ์ {eligible.length} เรื่อง</span>}>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
              <Field label="ปีที่ส่งเรื่อง">
                <select className={inputCls} style={inputStyle} value={cfg.year} onChange={(e) => set({ year: e.target.value })}>
                  <option value="all">ทุกปี</option>
                  {years.map((y) => <option key={y} value={y}>{y}</option>)}
                </select>
              </Field>
              <Field label="ประเภทกิจกรรม">
                <select className={inputCls} style={inputStyle} value={cfg.activity} onChange={(e) => set({ activity: e.target.value })}>
                  <option value="Kaizen">Kaizen</option>
                  <option value="Suggestion">Suggestion</option>
                  <option value="all">ทั้งหมด</option>
                </select>
              </Field>
              <Field label="Cost Saving ขั้นต่ำ (บาท/ปี)">
                <TextInput type="number" value={cfg.minCostSaving} onChange={(e) => set({ minCostSaving: e.target.value })} />
              </Field>
              <Field label="สถานะที่นับ">
                <div className="flex flex-col gap-1.5 pt-1">
                  {[["approved", "อนุมัติแล้ว"], ["pending_director", "รอผู้อำนวยการ"]].map(([s, l]) => (
                    <label key={s} className="flex items-center gap-2 text-sm" style={{ color: INK }}>
                      <input type="checkbox" checked={cfg.statuses.includes(s)} onChange={() => toggleStatus(s)} /> {l}
                    </label>
                  ))}
                </div>
              </Field>
            </div>
          </AwardSection>

          <AwardSection
            title="เกณฑ์การให้คะแนนและน้ำหนัก"
            sub="เกณฑ์อัตโนมัติใช้คัดกรองในขั้นตอนที่ 2 · เกณฑ์สีม่วงให้คณะกรรมการเป็นผู้ให้คะแนน · รวมเป็นคะแนนเต็ม 100"
            right={
              <div className="flex items-center gap-2">
                <span className="text-xs font-semibold px-2.5 py-1 rounded-full" style={totalW === 100 ? { background: "#E9F5EF", color: GREEN } : { background: "#FBEAE8", color: RED }}>น้ำหนักรวม {totalW}%</span>
                {totalW !== 100 && <button onClick={normalize} className="text-xs px-2.5 py-1 rounded-lg border font-medium" style={{ borderColor: NAVY, color: NAVY }}>ปรับให้รวม 100%</button>}
              </div>
            }
          >
            <div className="space-y-2">
              {cfg.criteria.map((c, i) => (
                <div key={c.id} className="rounded-lg border p-3 flex flex-col md:flex-row md:items-center gap-3" style={{ borderColor: "#E3E7ED" }}>
                  <div className="flex items-center gap-2 flex-1 min-w-0">
                    <span className="w-3 h-3 rounded-full flex-shrink-0" style={{ background: CRIT_COLORS[i % CRIT_COLORS.length] }} />
                    <div className="flex-1 min-w-0">
                      <input value={c.name} onChange={(e) => setCrit(c.id, { name: e.target.value })} className="w-full text-sm font-medium outline-none bg-transparent" style={{ color: INK }} />
                      <div className="text-[11px] mt-0.5" style={{ color: SUB }}>{c.type === "auto" ? AUTO_SOURCES[c.source] : "กรรมการแต่ละท่านให้คะแนน 1-5 แยกกัน แล้วระบบหาค่าเฉลี่ย"}</div>
                    </div>
                    <span className="text-[11px] px-2 py-0.5 rounded-full flex-shrink-0" style={c.type === "auto" ? { background: "#EAF3FC", color: "#1D5FC9" } : { background: "#F1ECFD", color: JUDGE_PURPLE }}>
                      {c.type === "auto" ? "อัตโนมัติ" : "กรรมการให้คะแนน"}
                    </span>
                  </div>
                  <div className="flex items-center gap-2 md:w-72">
                    <input type="range" min={0} max={50} value={toNum(c.weight)} onChange={(e) => setCrit(c.id, { weight: Number(e.target.value) })} className="flex-1" />
                    <input type="number" value={c.weight} onChange={(e) => setCrit(c.id, { weight: toNum(e.target.value) })} className="w-14 rounded border px-1.5 py-1 text-sm text-right" style={inputStyle} />
                    <span className="text-xs" style={{ color: SUB }}>%</span>
                    <button onClick={() => removeCrit(c.id)} title="ลบเกณฑ์" className="p-1 rounded hover:bg-red-50"><Trash2 size={15} color={RED} /></button>
                  </div>
                </div>
              ))}
            </div>
            <button onClick={addCrit} className="mt-3 text-sm font-medium flex items-center gap-1.5" style={{ color: NAVY }}><Plus size={16} /> เพิ่มเกณฑ์ที่กรรมการให้คะแนน</button>
          </AwardSection>

          <AwardSection title="คณะกรรมการ" sub="กรรมการแต่ละท่านมีบัญชีและรหัสผ่านของตัวเอง เข้าระบบแล้วจะเห็นเฉพาะหน้าให้คะแนนของตัวเอง ไม่เห็นคะแนนของกรรมการท่านอื่น">
            <div className="flex items-center gap-3 mb-3">
              <span className="text-sm" style={{ color: INK }}>จำนวนกรรมการรอบนี้</span>
              <select className="rounded-lg border px-3 py-1.5 text-sm" style={inputStyle} value={cfg.judgeCount} onChange={(e) => set({ judgeCount: Number(e.target.value) })}>
                {[1, 2, 3, 4, 5].map((n) => <option key={n} value={n}>{n} ท่าน</option>)}
              </select>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2">
              {JUDGE_SLOTS.map((j, idx) => {
                const active = idx < cfg.judgeCount;
                const acct = MOCK_USERS.find((u) => u.name === j.userName);
                return (
                  <div key={j.id} className="rounded-lg border p-3" style={{ borderColor: active ? "#DCD3F5" : "#EEF0F3", background: active ? "#FBFAFF" : "#F7F8FA", opacity: active ? 1 : 0.55 }}>
                    <div className="flex items-center gap-2">
                      <div className="w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold text-white" style={{ background: active ? JUDGE_PURPLE : "#AEB6C2" }}>{idx + 1}</div>
                      <div className="flex-1">
                        <div className="text-sm font-semibold" style={{ color: INK }}>{j.userName}</div>
                        <div className="text-[11px]" style={{ color: SUB }}>ชื่อผู้ใช้: {j.userName} · รหัสผ่าน: <span className="font-mono">{acct ? acct.empId : "-"}</span></div>
                      </div>
                      {!active && <span className="text-[10px]" style={{ color: SUB }}>ไม่ใช้รอบนี้</span>}
                    </div>
                    <input
                      value={cfg.judgeNames?.[j.id] || ""}
                      onChange={(e) => set({ judgeNames: { ...cfg.judgeNames, [j.id]: e.target.value } })}
                      disabled={!active}
                      placeholder="ชื่อจริงของกรรมการ (ไม่บังคับ)"
                      className="mt-2 w-full rounded border px-2 py-1 text-xs"
                      style={inputStyle}
                    />
                  </div>
                );
              })}
            </div>
            <label className="mt-3 flex items-start gap-2 text-sm" style={{ color: INK }}>
              <input type="checkbox" className="mt-0.5" checked={cfg.trimExtremes} onChange={(e) => set({ trimExtremes: e.target.checked })} />
              <span>
                ตัดคะแนนสูงสุดและต่ำสุดออกก่อนเฉลี่ย เมื่อมีกรรมการให้คะแนนตั้งแต่ 4 ท่านขึ้นไป
                <span className="block text-[11px]" style={{ color: SUB }}>ช่วยลดผลจากกรรมการที่ให้คะแนนสูงหรือต่ำผิดปกติ</span>
              </span>
            </label>
          </AwardSection>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
            <AwardSection title="ช่วงคะแนน Cost Saving" sub="ใช้กับเกณฑ์ 'มูลค่า Cost Saving'">
              <div className="space-y-2">
                {[5, 4, 3, 2].map((pt, i) => (
                  <div key={pt} className="flex items-center gap-2 text-sm" style={{ color: INK }}>
                    <span className="w-16 whitespace-nowrap">{pt} คะแนน</span>
                    <span style={{ color: SUB }}>ตั้งแต่</span>
                    <input type="number" value={cfg.costTiers[i]} onChange={(e) => set({ costTiers: cfg.costTiers.map((v, k2) => (k2 === i ? toNum(e.target.value) : v)) })} className="flex-1 rounded border px-2 py-1 text-sm text-right" style={inputStyle} />
                    <span className="text-xs" style={{ color: SUB }}>บาท</span>
                  </div>
                ))}
                <div className="text-xs" style={{ color: SUB }}>ต่ำกว่านั้นแต่มากกว่า 0 = 1 คะแนน</div>
              </div>
            </AwardSection>
            <AwardSection title="หมวดผลกระทบสูง" sub="Kaizen ที่อยู่ในหมวดเหล่านี้ได้ 5 คะแนนในเกณฑ์ผลกระทบ">
              <div className="flex flex-wrap gap-1.5">
                {CATEGORY_OPTIONS.map((c) => (
                  <button key={c} onClick={() => toggleImpact(c)} className="text-xs px-2.5 py-1 rounded-full border" style={cfg.highImpactCats.includes(c) ? { background: NAVY, borderColor: NAVY, color: "white" } : { borderColor: "#D8DDE5", color: SUB }}>{c}</button>
                ))}
              </div>
            </AwardSection>
            <AwardSection title="จำนวนรางวัล">
              <div className="space-y-3">
                <Field label="รางวัลหลัก (ชนะเลิศ / รองชนะเลิศ)">
                  <select className={inputCls} style={inputStyle} value={cfg.mainPrizes} onChange={(e) => set({ mainPrizes: Number(e.target.value) })}>
                    {[1, 2, 3].map((n) => <option key={n} value={n}>{n} รางวัล</option>)}
                  </select>
                </Field>
                <Field label="รางวัลชมเชย">
                  <select className={inputCls} style={inputStyle} value={cfg.honorable} onChange={(e) => set({ honorable: Number(e.target.value) })}>
                    {[0, 1, 2, 3, 4, 5].map((n) => <option key={n} value={n}>{n} รางวัล</option>)}
                  </select>
                </Field>
              </div>
            </AwardSection>
          </div>
          <div className="flex justify-end">
            <button onClick={() => setSub("screen")} className="px-4 py-2 rounded-lg text-sm font-medium text-white flex items-center gap-1.5" style={{ background: NAVY }}>ไปขั้นตอนคัดกรอง <ChevronRight size={16} /></button>
          </div>
        </div>
      )}

      {/* ---------------- 2. SCREENING ---------------- */}
      {sub === "screen" && (
        <div className="space-y-4" style={disabledStyle}>
          <AwardSection
            title="คัดกรองเบื้องต้นตามเกณฑ์อัตโนมัติ"
            sub="ระบบจัดเรียงจาก 'คะแนนคัดกรอง' (คำนวณจากเกณฑ์อัตโนมัติเท่านั้น) · ติ๊กเลือกเรื่องที่จะส่งให้คณะกรรมการตัดสิน"
            right={<span className="text-xs px-2.5 py-1 rounded-full font-semibold" style={{ background: "#F1ECFD", color: JUDGE_PURPLE }}>เลือกแล้ว {selected.size} / {eligible.length}</span>}
          >
            <div className="flex flex-wrap items-center gap-2 mb-3">
              <span className="text-sm" style={{ color: INK }}>เลือกอันดับสูงสุด</span>
              <input type="number" min={1} value={topN} onChange={(e) => setTopN(e.target.value)} className="w-16 rounded border px-2 py-1 text-sm text-right" style={inputStyle} />
              <span className="text-sm" style={{ color: INK }}>เรื่อง</span>
              <button onClick={pickTop} className="text-xs px-3 py-1.5 rounded-lg border font-medium" style={{ borderColor: NAVY, color: NAVY }}>เลือกอัตโนมัติ</button>
              <button onClick={() => setSelected(new Set(eligible.map((x) => x.k.id)))} className="text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: "#D8DDE5", color: SUB }}>เลือกทั้งหมด</button>
              <button onClick={() => setSelected(new Set())} className="text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: "#D8DDE5", color: SUB }}>ล้างการเลือก</button>
            </div>
            {eligible.length === 0 ? (
              <div className="text-sm py-8 text-center" style={{ color: SUB }}>ยังไม่มี Kaizen ที่เข้าเกณฑ์ ลองปรับเงื่อนไขในขั้นตอนที่ 1</div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm" style={{ minWidth: 760 }}>
                  <thead>
                    <tr className="text-left" style={{ color: SUB }}>
                      <th className="py-2 w-10" />
                      <th className="py-2 font-medium" style={{ minWidth: 220 }}>เรื่อง</th>
                      {autoCrit.map((c) => (
                        <th key={c.id} className="py-2 px-2 font-medium text-center text-xs" style={{ minWidth: 100 }}>{c.name}</th>
                      ))}
                      <th className="py-2 px-2 font-medium" style={{ minWidth: 150 }}>คะแนนคัดกรอง</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {eligible.map(({ k, pre }) => (
                      <tr key={k.id} className="border-t cursor-pointer" style={{ borderColor: "#EEF0F3", background: selected.has(k.id) ? "#F6F2FF" : undefined }} onClick={() => toggleSel(k.id)}>
                        <td className="py-2.5 text-center"><input type="checkbox" checked={selected.has(k.id)} onChange={() => toggleSel(k.id)} onClick={(e) => e.stopPropagation()} /></td>
                        <td className="py-2.5 pr-3">
                          <div className="font-medium" style={{ color: INK }}>{k.title}</div>
                          <div className="text-xs" style={{ color: SUB }}>{k.registeredNo} · {k.department}</div>
                        </td>
                        {autoCrit.map((c) => (
                          <td key={c.id} className="py-2 px-2 text-center">
                            <div className="font-semibold text-sm" style={{ color: "#1D5FC9" }}>{autoScore(c.source, k, cfg)}/5</div>
                            <div className="text-[10px] truncate max-w-[120px] mx-auto" style={{ color: SUB }}>{autoReason(c.source, k)}</div>
                          </td>
                        ))}
                        <td className="py-2 px-2">
                          <div className="flex items-center gap-2">
                            <div className="flex-1 h-2 rounded-full overflow-hidden" style={{ background: "#EEF0F3" }}><div style={{ width: `${pre}%`, height: "100%", background: "#1D5FC9" }} /></div>
                            <span className="text-sm font-bold w-10 text-right" style={{ color: NAVY }}>{pre}%</span>
                          </div>
                        </td>
                        <td className="py-2 pl-2">
                          <button onClick={(e) => { e.stopPropagation(); onView(k); }} title="ดูรายละเอียด" className="p-1.5 rounded hover:bg-slate-100"><Eye size={16} color={NAVY} /></button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </AwardSection>
          <div className="bg-white rounded-xl border p-4 flex items-center justify-between gap-3 flex-wrap" style={{ borderColor: "#E3E7ED" }}>
            <div className="text-sm" style={{ color: INK }}>
              {cfg.shortlist.length ? (
                <>ส่งให้กรรมการแล้ว <strong>{cfg.shortlist.length} เรื่อง</strong> เมื่อ {cfg.sentAt}{shortlistChanged && <span style={{ color: RED }}> · มีการเปลี่ยนรายการ ยังไม่ได้ส่งใหม่</span>}</>
              ) : (
                "ยังไม่ได้ส่งเรื่องให้คณะกรรมการ"
              )}
            </div>
            <button
              onClick={sendShortlist}
              disabled={!selected.size || (!shortlistChanged && cfg.shortlist.length > 0)}
              className="px-4 py-2 rounded-lg text-sm font-semibold text-white flex items-center gap-1.5 disabled:opacity-40"
              style={{ background: JUDGE_PURPLE }}
            >
              <Trophy size={16} /> {cfg.shortlist.length ? "อัปเดตรายการที่ส่งกรรมการ" : `ส่ง ${selected.size} เรื่องให้กรรมการตัดสิน`}
            </button>
          </div>
        </div>
      )}

      {/* ---------------- 3. TRACKING ---------------- */}
      {sub === "tracking" && (
        <AwardSection title="ความคืบหน้าของคณะกรรมการ" sub="Admin และผู้อำนวยการเห็นเฉพาะความคืบหน้า กรรมการแต่ละท่านไม่เห็นคะแนนของกันและกัน">
          {!items.length ? (
            <div className="text-sm py-8 text-center" style={{ color: SUB }}>ยังไม่ได้ส่งเรื่องให้กรรมการ (ทำในขั้นตอนที่ 2)</div>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
              {progress.map((p) => {
                const status = p.submittedAt ? "ส่งคะแนนแล้ว" : p.done === 0 ? "ยังไม่เริ่ม" : "กำลังให้คะแนน";
                const color = p.submittedAt ? GREEN : p.done === 0 ? SUB : JUDGE_PURPLE;
                return (
                  <div key={p.judge.id} className="rounded-lg border p-3" style={{ borderColor: "#E3E7ED" }}>
                    <div className="flex items-center justify-between gap-2">
                      <div className="text-sm font-semibold" style={{ color: INK }}>{p.judge.name}</div>
                      <span className="text-[11px] px-2 py-0.5 rounded-full font-medium" style={{ color, background: p.submittedAt ? "#E9F5EF" : "#F4F6F8" }}>{status}</span>
                    </div>
                    <div className="flex items-center gap-2 mt-2">
                      <ProgressBar done={p.done} total={p.total} />
                      <span className="text-[11px] whitespace-nowrap" style={{ color: SUB }}>{p.done}/{p.total}</span>
                    </div>
                    {p.submittedAt && <div className="text-[11px] mt-1.5" style={{ color: SUB }}>ส่งเมื่อ {p.submittedAt}</div>}
                    {p.submittedAt && !locked &&
                      (unlocking === p.judge.id ? (
                        <div className="flex gap-1 mt-2">
                          <button onClick={() => setUnlocking(null)} className="text-[11px] px-2 py-1 rounded border" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
                          <button onClick={() => unlockJudge(p.judge.id)} className="text-[11px] px-2 py-1 rounded text-white" style={{ background: JUDGE_PURPLE }}>ยืนยันเปิดให้แก้ไข</button>
                        </div>
                      ) : (
                        <button onClick={() => setUnlocking(p.judge.id)} className="text-[11px] mt-2 underline" style={{ color: JUDGE_PURPLE }}>เปิดให้กรรมการแก้ไขคะแนน</button>
                      ))}
                  </div>
                );
              })}
            </div>
          )}
        </AwardSection>
      )}

      {/* ---------------- 4. JUDGE SCORES ---------------- */}
      {sub === "scores" && (
        <div className="space-y-4">
          {results.length === 0 ? (
            <AwardSection title="คะแนนกรรมการ"><div className="text-sm py-8 text-center" style={{ color: SUB }}>ยังไม่ได้ส่งเรื่องให้กรรมการ</div></AwardSection>
          ) : (
            <>
              <AwardSection
                title="ภาพรวมค่าเฉลี่ยของแต่ละ Kaizen"
                sub={`ค่าเฉลี่ยคะแนนกรรมการ (เต็ม 5) ในแต่ละเกณฑ์${cfg.trimExtremes && judges.length >= 4 ? " · ตัดคะแนนสูงสุด-ต่ำสุดออกเมื่อมีกรรมการให้คะแนน 4 ท่านขึ้นไป" : ""} · ตัวเลขสีแดง = กรรมการให้คะแนนต่างกันตั้งแต่ 3 คะแนนขึ้นไป`}
                right={<button onClick={exportExcel} className="text-xs px-3 py-1.5 rounded-lg border font-medium flex items-center gap-1.5" style={{ borderColor: GREEN, color: GREEN }}><FileSpreadsheet size={14} /> Export Excel</button>}
              >
                <div style={{ width: "100%", height: 240 }}>
                  <ResponsiveContainer>
                    <BarChart data={results.map((r) => { const row = { name: r.k.registeredNo }; r.parts.forEach((p) => { if (p.type === "manual") row[p.name] = p.stats.avg === null ? 0 : Math.round(p.stats.avg * 10) / 10; }); return row; })} margin={{ top: 8, right: 8, left: -18, bottom: 0 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#EEF0F3" vertical={false} />
                      <XAxis dataKey="name" tick={{ fontSize: 11, fill: SUB }} />
                      <YAxis domain={[0, 5]} ticks={[0, 1, 2, 3, 4, 5]} tick={{ fontSize: 11, fill: SUB }} />
                      <Tooltip />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      {manualCrit.map((c) => {
                        const idx = cfg.criteria.findIndex((x) => x.id === c.id);
                        return <Bar key={c.id} dataKey={c.name} fill={CRIT_COLORS[idx % CRIT_COLORS.length]} radius={[4, 4, 0, 0]} />;
                      })}
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                <div className="overflow-x-auto mt-3">
                  <table className="w-full text-sm" style={{ minWidth: 720 }}>
                    <thead>
                      <tr className="text-left" style={{ color: SUB }}>
                        <th className="py-2 w-12 font-medium">อันดับ</th>
                        <th className="py-2 font-medium" style={{ minWidth: 220 }}>เรื่อง</th>
                        {manualCrit.map((c) => <th key={c.id} className="py-2 px-2 font-medium text-center text-xs" style={{ minWidth: 110 }}>{c.name}</th>)}
                        <th className="py-2 px-2 font-medium text-center text-xs">เฉลี่ยกรรมการ<br />(เต็ม 5)</th>
                        <th className="py-2 px-2 font-medium text-center text-xs">คะแนนรวม<br />(เต็ม 100)</th>
                        <th className="py-2 px-2 font-medium text-center text-xs">กรรมการ<br />ที่ให้คะแนน</th>
                      </tr>
                    </thead>
                    <tbody>
                      {results.map((r) => {
                        const mParts = r.parts.filter((p) => p.type === "manual");
                        const avgs = mParts.map((p) => p.stats.avg).filter((v) => v !== null);
                        const judgeAvg = avgs.length ? avgs.reduce((a, b) => a + b, 0) / avgs.length : null;
                        const scoredBy = judges.filter((j) => mParts.some((p) => typeof cfg.scores?.[j.id]?.[r.k.id]?.[p.id] === "number")).length;
                        return (
                          <tr key={r.k.id} className="border-t" style={{ borderColor: "#EEF0F3" }}>
                            <td className="py-2.5 font-semibold" style={{ color: NAVY }}>{r.rank}</td>
                            <td className="py-2.5 pr-3"><div className="font-medium" style={{ color: INK }}>{r.k.title}</div><div className="text-xs" style={{ color: SUB }}>{r.k.registeredNo} · {r.k.department}</div></td>
                            {mParts.map((p) => (
                              <td key={p.id} className="py-2.5 px-2 text-center">
                                <div className="font-bold" style={{ color: p.stats.spread >= 3 ? RED : JUDGE_PURPLE }}>{p.stats.avg === null ? "-" : fmtScore(Math.round(p.stats.avg * 10) / 10)}</div>
                                {p.stats.vals.length > 0 && <div className="text-[10px]" style={{ color: SUB }}>ต่ำสุด {Math.min(...p.stats.vals)} · สูงสุด {Math.max(...p.stats.vals)}</div>}
                              </td>
                            ))}
                            <td className="py-2.5 px-2 text-center font-semibold" style={{ color: INK }}>{judgeAvg === null ? "-" : judgeAvg.toFixed(2)}</td>
                            <td className="py-2.5 px-2 text-center font-bold" style={{ color: NAVY }}>{r.total}</td>
                            <td className="py-2.5 px-2 text-center text-xs" style={{ color: scoredBy === judges.length ? GREEN : SUB }}>{scoredBy}/{judges.length}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </AwardSection>

              <AwardSection
                title="คะแนนที่กรรมการแต่ละท่านส่งเข้ามา"
                sub="แยกตาม Kaizen · คะแนนที่มีขีดฆ่า = ถูกตัดออกตามกติกาตัดสูงสุด-ต่ำสุด · กรรมการที่ยังไม่กดส่งจะแสดงเป็น 'ร่าง'"
                right={
                  <select className="rounded-lg border px-3 py-1.5 text-sm" style={inputStyle} value={judgeFilter} onChange={(e) => setJudgeFilter(e.target.value)}>
                    <option value="all">กรรมการทุกท่าน</option>
                    {judges.map((j) => <option key={j.id} value={j.id}>{j.name}</option>)}
                  </select>
                }
              >
                <div className="space-y-4">
                  {results.map((r) => {
                    const mParts = r.parts.filter((p) => p.type === "manual");
                    // หาคะแนนที่ถูกตัดออก (ต่ำสุด 1 ค่า และสูงสุด 1 ค่า) ของแต่ละเกณฑ์
                    const dropped = {};
                    mParts.forEach((p) => {
                      if (!p.stats.trimmed) return;
                      const scored = p.stats.byJudge.filter((x) => typeof x.v === "number");
                      const lo = scored.reduce((a, b) => (b.v < a.v ? b : a));
                      const hi = scored.filter((x) => x !== lo).reduce((a, b) => (b.v > a.v ? b : a));
                      dropped[p.id] = new Set([lo.judge.id, hi.judge.id]);
                    });
                    const shown = judges.filter((j) => judgeFilter === "all" || j.id === judgeFilter);
                    return (
                      <div key={r.k.id} className="rounded-lg border" style={{ borderColor: "#E3E7ED" }}>
                        <div className="px-3 py-2.5 flex items-center justify-between gap-2 flex-wrap" style={{ background: "#F7F8FA" }}>
                          <div>
                            <div className="text-sm font-semibold" style={{ color: INK }}><span style={{ color: NAVY }}>#{r.rank}</span> {r.k.title}</div>
                            <div className="text-xs" style={{ color: SUB }}>{r.k.registeredNo} · {r.k.department}</div>
                          </div>
                          <div className="flex items-center gap-2">
                            <TierBadge tier={r.tier} small />
                            <span className="text-sm font-bold" style={{ color: NAVY }}>{r.total} คะแนน</span>
                            <button onClick={() => onView(r.k)} title="ดูรายละเอียด" className="p-1 rounded hover:bg-white"><Eye size={15} color={NAVY} /></button>
                          </div>
                        </div>
                        <div className="overflow-x-auto">
                          <table className="w-full text-sm" style={{ minWidth: 640 }}>
                            <thead>
                              <tr className="text-left" style={{ color: SUB }}>
                                <th className="py-2 px-3 font-medium" style={{ minWidth: 170 }}>กรรมการ</th>
                                {mParts.map((p) => <th key={p.id} className="py-2 px-2 font-medium text-center text-xs" style={{ minWidth: 100 }}>{p.name}</th>)}
                                <th className="py-2 px-2 font-medium text-center text-xs">เฉลี่ยของท่าน</th>
                                <th className="py-2 px-3 font-medium text-xs" style={{ minWidth: 200 }}>ความเห็น</th>
                              </tr>
                            </thead>
                            <tbody>
                              {shown.map((j) => {
                                const row = cfg.scores?.[j.id]?.[r.k.id] || {};
                                const submitted = !!cfg.scores?.[j.id]?.__meta?.submittedAt;
                                const nums = mParts.map((p) => row[p.id]).filter((v) => typeof v === "number");
                                const mine = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : null;
                                return (
                                  <tr key={j.id} className="border-t" style={{ borderColor: "#EEF0F3", opacity: submitted ? 1 : 0.7 }}>
                                    <td className="py-2 px-3">
                                      <div className="font-medium" style={{ color: INK }}>{j.name}</div>
                                      <span className="text-[10px] px-1.5 py-0.5 rounded-full" style={submitted ? { background: "#E9F5EF", color: GREEN } : { background: "#F4F6F8", color: SUB }}>{submitted ? "ส่งแล้ว" : "ร่าง (ยังไม่กดส่ง)"}</span>
                                    </td>
                                    {mParts.map((p) => {
                                      const v = row[p.id];
                                      const cut = dropped[p.id]?.has(j.id);
                                      return (
                                        <td key={p.id} className="py-2 px-2 text-center">
                                          {typeof v === "number" ? (
                                            <span className="inline-flex w-8 h-8 rounded-lg items-center justify-center font-bold" title={cut ? "ถูกตัดออกตามกติกาตัดสูงสุด-ต่ำสุด" : ""} style={cut ? { background: "#F4F6F8", color: "#9AA4B2", textDecoration: "line-through" } : { background: "#F1ECFD", color: JUDGE_PURPLE }}>{v}</span>
                                          ) : (
                                            <span style={{ color: "#C3CAD4" }}>-</span>
                                          )}
                                        </td>
                                      );
                                    })}
                                    <td className="py-2 px-2 text-center font-semibold" style={{ color: INK }}>{mine === null ? "-" : mine.toFixed(1)}</td>
                                    <td className="py-2 px-3 text-xs" style={{ color: row.__note ? INK : "#C3CAD4" }}>{row.__note || "-"}</td>
                                  </tr>
                                );
                              })}
                              <tr className="border-t" style={{ borderColor: "#D8DDE5", background: "#FBFAFF" }}>
                                <td className="py-2 px-3 text-xs font-semibold" style={{ color: JUDGE_PURPLE }}>ค่าเฉลี่ยที่ใช้คำนวณ</td>
                                {mParts.map((p) => (
                                  <td key={p.id} className="py-2 px-2 text-center font-bold" style={{ color: p.stats.spread >= 3 ? RED : JUDGE_PURPLE }}>{p.stats.avg === null ? "-" : fmtScore(Math.round(p.stats.avg * 10) / 10)}</td>
                                ))}
                                <td colSpan={2} />
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </AwardSection>
            </>
          )}
        </div>
      )}

      {/* ---------------- 5. RESULT ---------------- */}
      {sub === "result" && (
        <div className="space-y-4">
          {results.length === 0 ? (
            <AwardSection title="ผลการตัดสิน"><div className="text-sm py-8 text-center" style={{ color: SUB }}>ยังไม่ได้ส่งเรื่องให้กรรมการ</div></AwardSection>
          ) : (
            <>
              {submittedCount < judges.length && (
                <div className="rounded-xl p-3 border text-sm" style={{ background: "#FBEAE8", borderColor: "#F0C3BD", color: "#8B2E24" }}>
                  กรรมการส่งคะแนนแล้ว {submittedCount} จาก {judges.length} ท่าน ผลด้านล่างเป็นผลเบื้องต้นจากคะแนนที่มีอยู่ตอนนี้
                </div>
              )}
              <div className={`grid grid-cols-1 gap-4 ${cfg.mainPrizes === 3 ? "md:grid-cols-3" : cfg.mainPrizes === 2 ? "md:grid-cols-2" : ""}`}>
                {results.slice(0, cfg.mainPrizes).map((r) => {
                  const top2 = [...r.parts].sort((a, b) => b.pts - a.pts).slice(0, 2);
                  return (
                    <div key={r.k.id} className="bg-white rounded-xl border-2 p-4 flex flex-col" style={{ borderColor: r.tier.border }}>
                      <div className="flex items-center justify-between"><TierBadge tier={r.tier} /><span className="text-xs" style={{ color: SUB }}>อันดับ {r.rank}</span></div>
                      <div className="mt-3 font-semibold leading-snug" style={{ color: INK }}>{r.k.title}</div>
                      <div className="text-xs mt-0.5" style={{ color: SUB }}>{r.k.department} · {r.k.registeredNo}</div>
                      <div className="flex items-end gap-1 mt-3"><span className="text-3xl font-bold" style={{ color: r.tier.color }}>{r.total}</span><span className="text-xs mb-1" style={{ color: SUB }}>/ 100 คะแนน</span></div>
                      <div className="mt-2"><ContributionBar parts={r.parts} height={9} /></div>
                      <div className="text-xs mt-2" style={{ color: INK }}><span style={{ color: SUB }}>จุดเด่น: </span>{top2.map((p) => p.name).join(", ")}</div>
                      <div className="text-xs mt-1" style={{ color: INK }}><span style={{ color: SUB }}>Cost Saving: </span>{toNum(r.k.costSaving).toLocaleString()} บาท/ปี</div>
                      <div className="flex gap-2 mt-auto pt-3">
                        <button onClick={() => onView(r.k)} className="flex-1 text-xs py-1.5 rounded-lg border font-medium" style={{ borderColor: "#D8DDE5", color: NAVY }}>ดูรายละเอียด</button>
                        <button onClick={() => onPresent(r.k)} className="flex-1 text-xs py-1.5 rounded-lg font-medium text-white" style={{ background: NAVY }}>นำเสนอ</button>
                      </div>
                    </div>
                  );
                })}
              </div>

              <AwardSection title="บทวิเคราะห์ผลการตัดสิน" sub="สรุปอัตโนมัติจากคะแนนเฉลี่ยของคณะกรรมการตามเกณฑ์ที่ตั้งไว้">
                <ul className="space-y-1.5">
                  {insights.map((s, i) => (
                    <li key={i} className="text-sm flex gap-2" style={{ color: INK }}><span className="mt-1.5 w-1.5 h-1.5 rounded-full flex-shrink-0" style={{ background: GOLD }} /><span>{s}</span></li>
                  ))}
                </ul>
              </AwardSection>

              <AwardSection
                title="ตารางจัดอันดับทั้งหมด"
                sub="แถบสีแสดงว่าคะแนนมาจากเกณฑ์ใดบ้าง · ชี้ที่แถบเพื่อดูรายละเอียด"
                right={<button onClick={exportExcel} className="text-xs px-3 py-1.5 rounded-lg border font-medium flex items-center gap-1.5" style={{ borderColor: GREEN, color: GREEN }}><FileSpreadsheet size={14} /> Export Excel</button>}
              >
                <div className="flex flex-wrap gap-3 mb-3">
                  {cfg.criteria.map((c, i) => (
                    <span key={c.id} className="text-[11px] flex items-center gap-1" style={{ color: SUB }}><span className="w-2.5 h-2.5 rounded-sm" style={{ background: CRIT_COLORS[i % CRIT_COLORS.length] }} /> {c.name}</span>
                  ))}
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm" style={{ minWidth: 640 }}>
                    <thead>
                      <tr className="text-left" style={{ color: SUB }}>
                        <th className="py-2 w-12 font-medium">อันดับ</th>
                        <th className="py-2 font-medium">เรื่อง</th>
                        <th className="py-2 font-medium w-56">ที่มาของคะแนน</th>
                        <th className="py-2 font-medium text-center w-16">คะแนน</th>
                        <th className="py-2 font-medium w-28">รางวัล</th>
                      </tr>
                    </thead>
                    <tbody>
                      {results.map((r) => (
                        <tr key={r.k.id} className="border-t" style={{ borderColor: "#EEF0F3", background: r.tier ? r.tier.bg + "66" : undefined }}>
                          <td className="py-2.5 font-semibold" style={{ color: NAVY }}>{r.rank}</td>
                          <td className="py-2.5 pr-3"><div className="font-medium" style={{ color: INK }}>{r.k.title}</div><div className="text-xs" style={{ color: SUB }}>{r.k.department}</div></td>
                          <td className="py-2.5 pr-3"><ContributionBar parts={r.parts} /></td>
                          <td className="py-2.5 text-center font-bold" style={{ color: NAVY }}>{r.total}</td>
                          <td className="py-2.5"><TierBadge tier={r.tier} small /></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </AwardSection>

              {!locked && (
                <div className="bg-white rounded-xl border p-4 flex items-center justify-between gap-3 flex-wrap" style={{ borderColor: "#E3E7ED" }}>
                  <div className="text-sm" style={{ color: INK }}>
                    {isDirector
                      ? submittedCount < judges.length
                        ? `ยังมีกรรมการ ${judges.length - submittedCount} ท่านที่ยังไม่ส่งคะแนน แนะนำให้รอครบก่อนยืนยันผล`
                        : "กรรมการส่งคะแนนครบแล้ว สามารถยืนยันและประกาศผลได้"
                      : "ผู้อำนวยการโรงงานเป็นผู้ยืนยันผลขั้นสุดท้าย ระบบจะแจ้งเตือนทีมที่ได้รับรางวัลอัตโนมัติ"}
                  </div>
                  {isDirector &&
                    (confirming ? (
                      <div className="flex gap-2">
                        <button onClick={() => setConfirming(false)} className="px-3 py-2 rounded-lg text-sm border" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
                        <button onClick={confirmResult} className="px-4 py-2 rounded-lg text-sm font-semibold text-white" style={{ background: GREEN }}>ยืนยันประกาศผล</button>
                      </div>
                    ) : (
                      <button onClick={() => setConfirming(true)} className="px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1.5" style={{ background: GOLD, color: NAVY_DEEP }}><Trophy size={16} /> ยืนยันผลและประกาศรางวัล</button>
                    ))}
                </div>
              )}
            </>
          )}
        </div>
      )}
    </div>
  );
}

// ================= หน้ากรรมการ (แยกต่างหาก เห็นเฉพาะคะแนนของตัวเอง) =================
function JudgePage({ kaizens, award, currentUser, onScore, onView, onPresent, onSubmitted, onRefresh, syncMode }) {
  const cfg = award;
  const slotIndex = JUDGE_SLOTS.findIndex((j) => j.userName === currentUser.name);
  const judge = JUDGE_SLOTS[slotIndex];
  const [confirmSubmit, setConfirmSubmit] = useState(false);
  const items = useMemo(() => shortlisted(kaizens, cfg).sort(byRegNo), [kaizens, cfg]);
  const manualCrit = cfg.criteria.filter((c) => c.type === "manual");
  const autoCrit = cfg.criteria.filter((c) => c.type === "auto");

  if (!judge || slotIndex >= cfg.judgeCount) {
    return <div className="bg-white rounded-xl border p-6 text-sm" style={{ borderColor: "#E3E7ED", color: SUB }}>บัญชีนี้ไม่ได้อยู่ในคณะกรรมการ {cfg.roundName}</div>;
  }
  const realName = String(cfg.judgeNames?.[judge.id] || "").trim();
  const my = cfg.scores[judge.id] || {};
  const submittedAt = my.__meta?.submittedAt;
  const locked = !!submittedAt || !!cfg.decision;
  const total = items.length * manualCrit.length;
  const done = items.reduce((s, k) => s + manualCrit.filter((c) => typeof my[k.id]?.[c.id] === "number").length, 0);
  const complete = total > 0 && done >= total;

  const update = (kid, patch) => {
    if (locked) return;
    const row = { ...(my[kid] || {}), ...patch };
    Object.keys(patch).forEach((key) => patch[key] === undefined && delete row[key]);
    onScore(judge.id, { ...my, [kid]: row });
  };
  const setScore = (kid, cid, v) => update(kid, { [cid]: my[kid]?.[cid] === v ? undefined : v });
  const submit = () => {
    onScore(judge.id, { ...my, __meta: { submittedAt: nowStamp() } });
    onSubmitted(judge);
    setConfirmSubmit(false);
  };

  return (
    <div className="space-y-4">
      <AwardHeader cfg={cfg} editable={false} subtitle={`หน้าให้คะแนนของ ${judge.userName}${realName ? ` (${realName})` : ""} · คะแนนของท่านเป็นความลับ กรรมการท่านอื่นมองไม่เห็น`} syncMode={syncMode} onRefresh={onRefresh} />

      {cfg.decision ? (
        <AwardSection title="ประกาศผลแล้ว" sub={`ยืนยันโดย ${cfg.decision.confirmedBy} เมื่อ ${cfg.decision.confirmedAt} · ขอบคุณที่ร่วมเป็นคณะกรรมการ`}>
          <div className="space-y-1.5">
            {cfg.decision.winners.map((w) => (
              <div key={w.id} className="flex items-center gap-2 text-sm" style={{ color: INK }}>
                <TierBadge tier={[...AWARD_TIERS, HONOR_TIER].find((t) => t.label === w.tier)} small />
                <span className="font-medium">{w.title}</span>
                <span className="text-xs" style={{ color: SUB }}>{w.department}</span>
              </div>
            ))}
          </div>
        </AwardSection>
      ) : (
        <AwardSection
          title={submittedAt ? "ท่านส่งคะแนนเรียบร้อยแล้ว" : "วิธีให้คะแนน"}
          sub={submittedAt ? `ส่งเมื่อ ${submittedAt} · หากต้องการแก้ไข โปรดติดต่อ Admin` : "อ่านรายละเอียดแต่ละเรื่อง แล้วให้คะแนนเกณฑ์สีม่วง 1 = น้อย, 5 = ดีเยี่ยม · ใส่ความเห็นได้ (ไม่บังคับ) · ครบแล้วกด \"ส่งคะแนน\""}
          right={
            <span className="text-xs px-2.5 py-1 rounded-full font-semibold" style={complete ? { background: "#E9F5EF", color: GREEN } : { background: "#F1ECFD", color: JUDGE_PURPLE }}>
              {complete ? "ให้คะแนนครบแล้ว" : `ให้คะแนนแล้ว ${done}/${total}`}
            </span>
          }
        >
          <ProgressBar done={done} total={total} />
        </AwardSection>
      )}

      {items.length === 0 ? (
        <div className="bg-white rounded-xl border p-8 text-sm text-center" style={{ borderColor: "#E3E7ED", color: SUB }}>ยังไม่มีเรื่องที่ส่งมาให้ตัดสิน Admin จะแจ้งเตือนเมื่อมีเรื่องเข้ามา</div>
      ) : (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {items.map((k, idx) => {
            const row = my[k.id] || {};
            const kDone = manualCrit.every((c) => typeof row[c.id] === "number");
            return (
              <div key={k.id} className="bg-white rounded-xl border p-4" style={{ borderColor: kDone ? "#A9D6BF" : "#E3E7ED" }}>
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <div className="text-[11px] font-semibold" style={{ color: JUDGE_PURPLE }}>เรื่องที่ {idx + 1} / {items.length}</div>
                    <div className="font-semibold leading-snug mt-0.5" style={{ color: INK }}>{k.title}</div>
                    <div className="text-xs mt-0.5" style={{ color: SUB }}>{k.registeredNo} · {k.department}</div>
                  </div>
                  {kDone && <span className="text-[11px] px-2 py-0.5 rounded-full whitespace-nowrap" style={{ background: "#E9F5EF", color: GREEN }}>ครบแล้ว</span>}
                </div>
                <div className="flex flex-wrap gap-1.5 mt-2">
                  {autoCrit.map((c) => (
                    <span key={c.id} className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: "#EAF3FC", color: "#1D5FC9" }}>{autoReason(c.source, k)}</span>
                  ))}
                </div>
                <div className="flex gap-2 mt-3">
                  <button onClick={() => onView(k)} className="flex-1 text-xs py-1.5 rounded-lg border font-medium flex items-center justify-center gap-1" style={{ borderColor: "#D8DDE5", color: NAVY }}><Eye size={13} /> ดูรายละเอียด</button>
                  <button onClick={() => onPresent(k)} className="flex-1 text-xs py-1.5 rounded-lg border font-medium flex items-center justify-center gap-1" style={{ borderColor: "#D8DDE5", color: NAVY }}><Presentation size={13} /> ดูแบบนำเสนอ</button>
                </div>
                <div className="mt-3 space-y-2.5" style={locked ? { opacity: 0.7, pointerEvents: "none" } : {}}>
                  {manualCrit.map((c) => (
                    <div key={c.id}>
                      <div className="text-xs font-medium mb-1" style={{ color: INK }}>{c.name} <span style={{ color: SUB }}>({c.weight}%)</span></div>
                      <div className="flex gap-1.5">
                        {[1, 2, 3, 4, 5].map((v) => (
                          <button
                            key={v}
                            onClick={() => setScore(k.id, c.id, v)}
                            className="flex-1 h-9 rounded-lg text-sm font-semibold border"
                            style={row[c.id] === v ? { background: JUDGE_PURPLE, borderColor: JUDGE_PURPLE, color: "white" } : { borderColor: "#DCD3F5", color: JUDGE_PURPLE, background: "#FAF8FF" }}
                          >
                            {v}
                          </button>
                        ))}
                      </div>
                    </div>
                  ))}
                  <textarea
                    value={row.__note || ""}
                    onChange={(e) => update(k.id, { __note: e.target.value })}
                    placeholder="ความเห็นของกรรมการ (ไม่บังคับ)"
                    className="w-full rounded-lg border px-3 py-2 text-xs outline-none"
                    style={{ ...inputStyle, minHeight: 52 }}
                  />
                </div>
              </div>
            );
          })}
        </div>
      )}

      {items.length > 0 && !locked && (
        <div className="bg-white rounded-xl border p-4 flex items-center justify-between gap-3 flex-wrap sticky bottom-3" style={{ borderColor: "#E3E7ED", boxShadow: "0 6px 20px rgba(11,29,51,0.10)" }}>
          <div className="text-sm" style={{ color: INK }}>{complete ? "ให้คะแนนครบทุกเรื่องแล้ว ตรวจทานแล้วกดส่งคะแนนได้เลย" : `ยังเหลืออีก ${total - done} ช่อง`}</div>
          {confirmSubmit ? (
            <div className="flex gap-2">
              <button onClick={() => setConfirmSubmit(false)} className="px-3 py-2 rounded-lg text-sm border" style={{ borderColor: "#D8DDE5", color: SUB }}>กลับไปแก้ไข</button>
              <button onClick={submit} className="px-4 py-2 rounded-lg text-sm font-semibold text-white" style={{ background: GREEN }}>ยืนยันส่งคะแนน</button>
            </div>
          ) : (
            <button onClick={() => setConfirmSubmit(true)} disabled={!complete} className="px-4 py-2 rounded-lg text-sm font-semibold text-white disabled:opacity-40" style={{ background: JUDGE_PURPLE }}>
              ส่งคะแนน
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function LoginScreen({ onLogin }) {
  const [name, setName] = useState("");
  const [pass, setPass] = useState("");
  const [error, setError] = useState("");
  const [showDemo, setShowDemo] = useState(false);
  const [open, setOpen] = useState(false);
  const passRef = useRef(null);

  const filteredUsers = useMemo(() => {
    const q = name.trim().toLowerCase();
    const list = q ? MOCK_USERS.filter((u) => u.name.toLowerCase().includes(q)) : MOCK_USERS;
    return list.slice(0, 30);
  }, [name]);

  const selectUser = (u) => {
    setName(u.name);
    setOpen(false);
    setError("");
    if (passRef.current) passRef.current.focus();
  };

  const submit = () => {
    const user = MOCK_USERS.find(
      (u) => u.name.trim().toLowerCase() === name.trim().toLowerCase()
    );
    if (!user || user.empId !== pass.trim()) {
      setError("ชื่อผู้ใช้หรือรหัสพนักงานไม่ถูกต้อง");
      return;
    }
    setError("");
    onLogin(user);
  };
  const handleKeyDown = (e) => {
    if (e.key === "Enter") submit();
    if (e.key === "Escape") setOpen(false);
  };

  return (
    <div className="min-h-screen w-full flex items-center justify-center px-4" style={{ background: NAVY_DEEP, fontFamily: "Tahoma, 'Segoe UI', system-ui, sans-serif" }}>
      <div className="w-full max-w-sm">
        <div className="flex items-center gap-3 justify-center mb-6">
          <div className="w-11 h-11 rounded-xl flex items-center justify-center bg-white p-1.5 shadow-lg">
            <img src={COMPANY_LOGO} onError={useLogoFallback} alt="Company Logo" className="w-full h-full object-contain" />
          </div>
          <div>
            <div className="text-white font-semibold text-sm leading-tight">Kaizen System</div>
            <div className="text-[11px] leading-tight" style={{ color: "#9FB0C6" }}>เข้าสู่ระบบตามสิทธิ์การใช้งาน</div>
          </div>
        </div>
        <div className="bg-white rounded-2xl shadow-2xl p-6 space-y-4">
          <Field label="ชื่อผู้ใช้ (พิมพ์ค้นหาหรือเลือกจากรายชื่อ)" required>
            <div className="relative">
              <User size={15} className="absolute left-3 top-1/2 -translate-y-1/2 z-10" style={{ color: "#AEB6C2" }} />
              <input
                value={name}
                onChange={(e) => { setName(e.target.value); setOpen(true); }}
                onFocus={() => setOpen(true)}
                onKeyDown={handleKeyDown}
                className={inputCls}
                style={{ ...inputStyle, paddingLeft: 32 }}
                placeholder="พิมพ์ชื่อหน่วยงาน หรือชื่อ-นามสกุล"
                autoComplete="off"
              />
              {open && (
                <div
                  className="absolute z-20 left-0 right-0 mt-1 bg-white rounded-lg shadow-xl border max-h-56 overflow-y-auto"
                  style={{ borderColor: "#D8DDE5" }}
                >
                  {filteredUsers.length === 0 && (
                    <div className="px-3 py-2 text-xs" style={{ color: SUB }}>ไม่พบผู้ใช้งานที่ตรงกัน</div>
                  )}
                  {filteredUsers.map((u) => {
                    const deptText =
                      u.departments.length === 0 ? "" :
                      u.departments.length <= 2 ? u.departments.join(", ") :
                      `${u.departments.slice(0, 2).join(", ")} และอีก ${u.departments.length - 2} หน่วยงาน`;
                    return (
                      <div
                        key={u.empId}
                        onMouseDown={(e) => { e.preventDefault(); selectUser(u); }}
                        className="px-3 py-2 cursor-pointer hover:bg-blue-50 border-b last:border-b-0"
                        style={{ borderColor: "#F2F4F7" }}
                      >
                        <div className="flex items-center justify-between gap-2">
                          <span className="text-sm font-medium" style={{ color: INK }}>{u.name}</span>
                          <span className="text-[10px] px-1.5 py-0.5 rounded-full whitespace-nowrap" style={{ background: "#EAF3FC", color: DB_BLUE }}>{ROLE_LABEL[u.role]}</span>
                        </div>
                        {deptText && <div className="text-[11px] mt-0.5" style={{ color: SUB }}>{deptText}</div>}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </Field>
          <Field label="รหัสผ่าน (รหัสพนักงาน)" required>
            <div className="relative">
              <KeyRound size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: "#AEB6C2" }} />
              <input ref={passRef} type="password" value={pass} onChange={(e) => setPass(e.target.value)} onFocus={() => setOpen(false)} onKeyDown={handleKeyDown} className={inputCls} style={{ ...inputStyle, paddingLeft: 32 }} placeholder="เช่น 1001" />
            </div>
          </Field>
          {error && (
            <div className="flex items-center gap-1.5 text-xs rounded-lg px-3 py-2" style={{ background: "#F9E9E7", color: RED }}>
              <AlertTriangle size={13} /> {error}
            </div>
          )}
          <button type="button" onClick={submit} className="w-full py-2.5 rounded-lg text-sm font-medium text-white flex items-center justify-center gap-1.5" style={{ background: NAVY }}>
            <Lock size={14} /> เข้าสู่ระบบ
          </button>
        </div>
        <button onClick={() => setShowDemo((s) => !s)} className="w-full text-center text-xs mt-4" style={{ color: "#9FB0C6" }}>
          {showDemo ? "ซ่อน" : "แสดง"}บัญชีทดสอบ (สำหรับสาธิตเท่านั้น)
        </button>
        {showDemo && (
          <div className="mt-2 bg-white/95 rounded-xl p-3 text-xs space-y-1 max-h-64 overflow-y-auto" style={{ color: SUB }}>
            {MOCK_USERS.filter((u) => u.role !== "judge").map((u) => (
              <div key={u.empId} className="flex justify-between">
                <span>{u.name} <span className="opacity-60">({ROLE_LABEL[u.role]})</span></span>
                <span className="font-mono">{u.empId}</span>
              </div>
            ))}
            <div className="pt-1 opacity-70">* ในระบบจริง รหัสผ่านต้องแยกจากรหัสพนักงาน และตรวจสอบผ่านฐานข้อมูล/SSO ฝั่ง backend เท่านั้น</div>
          </div>
        )}
      </div>
    </div>
  );
}

export default function KaizenApp() {
  const [currentUser, setCurrentUser] = useState(null);
  const [kaizens, setKaizens] = useState(seedData());
  const [sheetStatus, setSheetStatus] = useState("checking"); // "checking" | "connected" | "demo"

  useEffect(() => {
    fetchKaizensFromSheet()
      .then((data) => {
        if (Array.isArray(data) && data.length > 0) {
          setKaizens(data);
          setSheetStatus("connected");
        } else {
          // Sheet เชื่อมได้แต่ยังไม่มีข้อมูล — คงข้อมูลตัวอย่างไว้ให้เห็นหน้าตาระบบ
          setSheetStatus("connected-empty");
        }
      })
      .catch(() => {
        // เชื่อมต่อไม่ได้ (คาดว่าเกิดขึ้นถ้าทดสอบอยู่ใน claude.ai artifact) — ใช้ข้อมูลตัวอย่างแทน
        setSheetStatus("demo");
      });
  }, []);
  const [departments, setDepartments] = useState(DEFAULT_DEPARTMENTS);
  const [tab, setTab] = useState("dashboard");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [viewing, setViewing] = useState(null);
  const [reviewing, setReviewing] = useState(null);
  const [presenting, setPresenting] = useState(null);
  const [showDeptManager, setShowDeptManager] = useState(false);
  const [adminEditing, setAdminEditing] = useState(null);
  const [adminListFilter, setAdminListFilter] = useState("all"); // all | noimage | incomplete
  const [adminSearch, setAdminSearch] = useState("");
  const [recalling, setRecalling] = useState(null);
  const [notifications, setNotifications] = useState(() => buildInitialNotifications(kaizens));
  const [showNotifs, setShowNotifs] = useState(false);
  const [award, setAward] = useState(() => normalizeAward(null));
  const [awardSync, setAwardSync] = useState("local"); // "local" | "sheet"
  const awardTimer = useRef(null);
  const judgeTimers = useRef({});

  const loadAward = () =>
    fetchAwardFromSheet()
      .then((res) => {
        if (res.supported) {
          setAwardSync("sheet");
          if (res.award) setAward(normalizeAward(res.award));
        }
      })
      .catch(() => {});
  useEffect(() => {
    loadAward();
  }, []);

  const changeAward = (next) => {
    setAward(next);
    if (awardSync === "sheet") {
      clearTimeout(awardTimer.current);
      awardTimer.current = setTimeout(() => saveAwardToSheet(next), 1200);
    }
  };

  const scoreAward = (judgeId, judgeScores) => {
    setAward((a) => ({ ...a, scores: { ...a.scores, [judgeId]: judgeScores } }));
    if (awardSync === "sheet") {
      clearTimeout(judgeTimers.current[judgeId]);
      judgeTimers.current[judgeId] = setTimeout(() => saveJudgeScoresToSheet(judgeId, judgeScores), 800);
    }
  };

  const sendToJudges = (count) =>
    activeJudges(award).forEach((j) =>
      pushNotification(j.userName, `มี Kaizen ${count} เรื่อง รอท่านให้คะแนนใน ${award.roundName} (เมนู "ให้คะแนน")`, null)
    );

  const judgeSubmitted = (judge) => {
    pushNotification("Admin", `${judge.userName} ส่งคะแนน ${award.roundName} เรียบร้อยแล้ว`, null);
    pushNotification("ผอ.โรงงาน", `${judge.userName} ส่งคะแนน ${award.roundName} เรียบร้อยแล้ว`, null);
  };

  const announceAward = (winners) => {
    winners.forEach((w) =>
      pushNotification(w.submittedBy, `🏆 ยินดีด้วย! เรื่อง "${w.title}" ได้รับ${w.tier} ใน ${award.roundName}`, w.id)
    );
    pushNotification("Admin", `ผู้อำนวยการยืนยันผล ${award.roundName} แล้ว (${winners.length} รางวัล)`, winners[0]?.id);
  };

  const pushNotification = (recipientName, message, kaizenId) => {
    if (!recipientName) return;
    setNotifications((list) => [{ id: uid(), recipientName, message, kaizenId, read: false, date: nowStamp() }, ...list]);
  };
  const findSupervisorFor = (department) => MOCK_USERS.find((u) => u.role === "supervisor" && u.departments.includes(department));
  const findByRole = (role) => MOCK_USERS.find((u) => u.role === role);

  const updateKaizen = (id, patch) => setKaizens((list) => list.map((k) => (k.id === id ? { ...k, ...patch } : k)));

  const saveForm = (formData, status) => {
    if (editing && editing.id) {
      const updatedRecord = {
        ...editing,
        ...formData,
        status,
        history: status === "pending_supervisor"
          ? [...editing.history, { stage: "ส่งเรื่อง", actor: formData.submittedBy || "พนักงาน", action: "submit", comment: "", date: nowStamp() }]
          : editing.history,
      };
      updateKaizen(editing.id, updatedRecord);
      syncKaizenToSheet("update", updatedRecord);
      if (status === "pending_supervisor") {
        const supervisor = findSupervisorFor(formData.department);
        if (supervisor) pushNotification(supervisor.name, `มีเรื่อง Kaizen ส่งใหม่อีกครั้งรอคุณอนุมัติ: "${formData.title}"`, editing.id);
      }
    } else {
      const id = uid();
      // เลขที่ใหม่ = เลขสูงสุดของปีปัจจุบัน + 1 (รองรับข้อมูลที่นำเข้าจากระบบเดิม เช่น KZ-2026-146)
      const year = new Date().getFullYear();
      const maxNo = kaizens.reduce((m, k) => {
        const hit = String(k.registeredNo || "").match(new RegExp(`^KZ-${year}-(\\d+)$`));
        return hit ? Math.max(m, parseInt(hit[1], 10)) : m;
      }, 0);
      const registeredNo = `KZ-${year}-${String(maxNo + 1).padStart(3, "0")}`;
      const newRecord = {
        id,
        registeredNo,
        date: todayISO(),
        ...formData,
        categories: [],
        activityType: "Kaizen",
        kaizenType: "Kaizen Project",
        status,
        history: status === "pending_supervisor" ? [{ stage: "ส่งเรื่อง", actor: formData.submittedBy || "พนักงาน", action: "submit", comment: "", date: nowStamp() }] : [],
      };
      setKaizens((list) => [newRecord, ...list]);
      syncKaizenToSheet("create", newRecord);
      if (status === "pending_supervisor") {
        const supervisor = findSupervisorFor(formData.department);
        if (supervisor) pushNotification(supervisor.name, `มีเรื่อง Kaizen ใหม่รอคุณอนุมัติ: "${formData.title}"`, id);
      }
    }
    setShowForm(false);
    setEditing(null);
  };

  // Admin แก้ไขข้อมูล Kaizen ได้ทุกสถานะ (รวมเรื่องที่อนุมัติแล้ว) โดยไม่เปลี่ยนสถานะ และบันทึกประวัติการแก้ไข
  const EDIT_LABELS = {
    title: "หัวเรื่อง", team: "ชื่อทีม", department: "หน่วยงาน", problem: "ที่มาของปัญหา", improvement: "แนวทางการปรับปรุง",
    budget: "งบประมาณ", result: "ผลที่ได้รับ", costSaving: "Cost Saving", costSavingCalc: "วิธีคำนวณ", members: "สมาชิกทีม",
    categories: "หมวดหมู่", activityType: "ประเภทกิจกรรม", kaizenType: "ประเภท Kaizen", attachments: "ไฟล์แนบ",
  };
  const adminSave = (formData, note) => {
    const original = adminEditing;
    if (!original) return;
    const changed = Object.keys(EDIT_LABELS).filter((key) => JSON.stringify(original[key] ?? "") !== JSON.stringify(formData[key] ?? ""));
    if ((original.images?.before || null) !== (formData.images?.before || null)) changed.push(formData.images?.before ? "ภาพก่อนปรับปรุง" : "ลบภาพก่อนปรับปรุง");
    if ((original.images?.after || null) !== (formData.images?.after || null)) changed.push(formData.images?.after ? "ภาพหลังปรับปรุง" : "ลบภาพหลังปรับปรุง");
    if (!changed.length && !note) {
      setAdminEditing(null);
      return;
    }
    const comment = [changed.length ? `แก้ไข: ${changed.map((c) => EDIT_LABELS[c] || c).join(", ")}` : "", note].filter(Boolean).join(" · ");
    const updatedRecord = {
      ...original,
      ...formData,
      id: original.id,
      registeredNo: original.registeredNo,
      date: original.date,
      submittedBy: original.submittedBy,
      status: original.status,
      history: [...(original.history || []), { stage: "แก้ไขข้อมูล", actor: currentUser.name, action: "edit", comment, date: nowStamp() }],
    };
    updateKaizen(original.id, updatedRecord);
    syncKaizenToSheet("update", updatedRecord);
    setAdminEditing(null);
  };

  const decide = (k, action, comment, extra, returnTo) => {
    const forwardMap = { pending_supervisor: "pending_admin", pending_admin: "pending_director", pending_director: "approved" };
    const returnStatusMap = { employee: "rejected", supervisor: "pending_supervisor", admin: "pending_admin" };
    const stageLabelMap = { pending_supervisor: "หัวหน้างาน", pending_admin: "Admin", pending_director: "ผู้อำนวยการ" };

    let nextStatus, historyAction, historyTarget;
    if (action === "approve") {
      nextStatus = forwardMap[k.status];
      historyAction = "approve";
    } else {
      const target = returnTo || "employee";
      nextStatus = returnStatusMap[target];
      historyAction = target === "employee" ? "reject" : "return";
      historyTarget = target;
    }

    const updatedRecord = {
      ...k,
      ...(extra || {}),
      status: nextStatus,
      history: [...k.history, { stage: stageLabelMap[k.status], actor: currentUser.name, action: historyAction, target: historyTarget, comment, date: nowStamp() }],
    };
    updateKaizen(k.id, updatedRecord);
    syncKaizenToSheet("update", updatedRecord);

    // แจ้งเตือนผู้เกี่ยวข้องขั้นถัดไป
    if (action === "approve") {
      if (k.status === "pending_supervisor") {
        const admin = findByRole("admin");
        if (admin) pushNotification(admin.name, `หัวหน้างานอนุมัติแล้ว รอ Admin ตรวจสอบ: "${k.title}"`, k.id);
        pushNotification(k.submittedBy, `เรื่อง "${k.title}" ผ่านการอนุมัติจากหัวหน้างานแล้ว กำลังรอ Admin ตรวจสอบ`, k.id);
      } else if (k.status === "pending_admin") {
        const director = findByRole("director");
        if (director) pushNotification(director.name, `Admin ตรวจสอบแล้ว รออนุมัติขั้นสุดท้าย: "${k.title}"`, k.id);
        pushNotification(k.submittedBy, `เรื่อง "${k.title}" ผ่านการตรวจสอบจาก Admin แล้ว รอผู้อำนวยการอนุมัติขั้นสุดท้าย`, k.id);
      } else if (k.status === "pending_director") {
        pushNotification(k.submittedBy, `🎉 เรื่อง "${k.title}" ได้รับอนุมัติขั้นสุดท้ายเรียบร้อยแล้ว!`, k.id);
      }
    } else {
      const target = returnTo || "employee";
      if (target === "employee") {
        pushNotification(k.submittedBy, `เรื่อง "${k.title}" ถูกตีกลับ${comment ? `: ${comment}` : ""}`, k.id);
      } else if (target === "supervisor") {
        const supervisor = findSupervisorFor(k.department);
        if (supervisor) pushNotification(supervisor.name, `เรื่อง "${k.title}" ถูกส่งกลับมาให้ท่านทบทวนอีกครั้ง`, k.id);
        pushNotification(k.submittedBy, `เรื่อง "${k.title}" ถูกส่งกลับไปให้หัวหน้างานทบทวนอีกครั้ง`, k.id);
      } else if (target === "admin") {
        const admin = findByRole("admin");
        if (admin) pushNotification(admin.name, `เรื่อง "${k.title}" ถูกส่งกลับมาให้ท่านทบทวนอีกครั้ง`, k.id);
        pushNotification(k.submittedBy, `เรื่อง "${k.title}" ถูกส่งกลับไปให้ Admin ทบทวนอีกครั้ง`, k.id);
      }
    }
    setReviewing(null);
  };

  const recallToDraft = (k) => {
    const supervisor = findSupervisorFor(k.department);
    const updatedRecord = {
      ...k,
      status: "draft",
      history: [...k.history, { stage: "พนักงาน", actor: currentUser.name, action: "recall", comment: "", date: nowStamp() }],
    };
    updateKaizen(k.id, updatedRecord);
    syncKaizenToSheet("update", updatedRecord);
    if (supervisor) pushNotification(supervisor.name, `เรื่อง "${k.title}" ถูกถอนกลับไปแก้ไขโดยผู้ส่งเรื่อง (ไม่ต้องพิจารณาแล้วในตอนนี้)`, k.id);
    setRecalling(null);
  };

  const ALL_NAV = [
    { key: "dashboard", label: "แดชบอร์ด", icon: LayoutDashboard, roles: ["employee", "supervisor", "admin", "director"] },
    { key: "employee", label: "พนักงาน", icon: User, roles: ["employee"] },
    { key: "supervisor", label: "หัวหน้างาน", icon: ShieldCheck, roles: ["supervisor"] },
    { key: "admin", label: "Admin", icon: ClipboardList, roles: ["admin"] },
    { key: "director", label: "ผู้อำนวยการ", icon: Building2, roles: ["director"] },
    { key: "award", label: "Kaizen Award", icon: Trophy, roles: ["admin", "director"] },
    { key: "judge", label: "ให้คะแนน Kaizen Award", icon: Medal, roles: ["judge"] },
  ];
  const NAV = currentUser ? ALL_NAV.filter((n) => n.roles.includes(currentUser.role)) : ALL_NAV;

  if (!currentUser) {
    return (
      <LoginScreen
        onLogin={(user) => {
          setCurrentUser(user);
          setTab(user.role); // ล็อกอินแล้วพาไปหน้าตามสิทธิ์ของตนเองก่อน
        }}
      />
    );
  }

  const myKaizens = kaizens.filter((k) => k.submittedBy === currentUser.name);
  const myNotifications = notifications.filter((n) => n.recipientName === currentUser.name);
  const myUnreadCount = myNotifications.filter((n) => !n.read).length;
  const supervisorQueue = kaizens.filter((k) => k.status === "pending_supervisor" && currentUser.departments.includes(k.department));
  const adminQueue = kaizens.filter((k) => k.status === "pending_admin");
  const directorQueue = kaizens.filter((k) => k.status === "pending_director");

  return (
    <div className="min-h-screen w-full" style={{ background: SLATE, fontFamily: "Tahoma, 'Segoe UI', system-ui, sans-serif" }}>
      {/* Top bar */}
      <div className="sticky top-0 z-30 border-b" style={{ background: NAVY_DEEP, borderColor: "#0A1830" }}>
        <div className="max-w-6xl mx-auto px-6 py-3 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-lg flex items-center justify-center bg-white p-1 shadow-sm">
              <img src={COMPANY_LOGO} onError={useLogoFallback} alt="Company Logo" className="w-full h-full object-contain" />
            </div>
            <div>
              <div className="text-white font-semibold text-sm leading-tight">Kaizen System</div>
              <div className="text-[11px] leading-tight" style={{ color: "#9FB0C6" }}>ระบบบริหารจัดการข้อเสนอ Kaizen & Suggestion</div>
            </div>
            <span
              className="text-[10px] px-2 py-0.5 rounded-full font-medium ml-1 whitespace-nowrap"
              style={
                sheetStatus === "connected" || sheetStatus === "connected-empty"
                  ? { background: "#E5F3EC", color: GREEN }
                  : sheetStatus === "checking"
                  ? { background: "#EEF0F3", color: SUB }
                  : { background: "#FBF0DD", color: "#B8791F" }
              }
              title={
                sheetStatus === "connected" ? "เชื่อมต่อ Google Sheet สำเร็จ กำลังใช้ข้อมูลจริง"
                : sheetStatus === "connected-empty" ? "เชื่อมต่อ Google Sheet สำเร็จ แต่ยังไม่มีข้อมูล กำลังแสดงข้อมูลตัวอย่าง"
                : sheetStatus === "checking" ? "กำลังตรวจสอบการเชื่อมต่อ..."
                : "เชื่อมต่อ Google Sheet ไม่ได้ กำลังใช้ข้อมูลตัวอย่าง (โหมดสาธิต)"
              }
            >
              {sheetStatus === "connected" ? "🟢 เชื่อมต่อแล้ว"
                : sheetStatus === "connected-empty" ? "🟢 เชื่อมต่อแล้ว (ยังไม่มีข้อมูล)"
                : sheetStatus === "checking" ? "กำลังตรวจสอบ..."
                : "🟡 โหมดสาธิต"}
            </span>
          </div>
          <div className="flex items-center gap-3">
            <div className="flex gap-1">
              {NAV.map((n) => {
                const Icon = n.icon;
                const active = tab === n.key;
                return (
                  <button
                    key={n.key}
                    onClick={() => setTab(n.key)}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium transition"
                    style={active ? { background: GOLD, color: NAVY_DEEP } : { color: "#C7D2E0" }}
                  >
                    <Icon size={14} /> {n.label}
                  </button>
                );
              })}
            </div>
            <div className="w-px h-6" style={{ background: "#22375A" }} />
            <div className="relative">
              <button
                onClick={() => setShowNotifs((s) => !s)}
                title="การแจ้งเตือน"
                className="p-1.5 rounded-lg relative"
                style={{ color: "#C7D2E0" }}
              >
                {myUnreadCount > 0 ? <BellRing size={17} /> : <Bell size={17} />}
                {myUnreadCount > 0 && (
                  <span
                    className="absolute -top-0.5 -right-0.5 min-w-[16px] h-4 px-1 rounded-full text-[10px] font-bold flex items-center justify-center text-white"
                    style={{ background: RED }}
                  >
                    {myUnreadCount > 9 ? "9+" : myUnreadCount}
                  </span>
                )}
              </button>
              {showNotifs && (
                <NotificationPanel
                  notifications={myNotifications}
                  onClose={() => setShowNotifs(false)}
                  onMarkAllRead={() => setNotifications((list) => list.map((n) => (n.recipientName === currentUser.name ? { ...n, read: true } : n)))}
                  onItemClick={(n) => {
                    setNotifications((list) => list.map((x) => (x.id === n.id ? { ...x, read: true } : x)));
                    const k = kaizens.find((x) => x.id === n.kaizenId);
                    if (k) setViewing(k);
                    setShowNotifs(false);
                  }}
                />
              )}
            </div>
            <div className="text-right leading-tight">
              <div className="text-white text-xs font-medium">{currentUser.name}</div>
              <div className="text-[10px]" style={{ color: "#9FB0C6" }}>{ROLE_LABEL[currentUser.role]}</div>
            </div>
            <button
              onClick={() => { setCurrentUser(null); setTab("dashboard"); }}
              title="ออกจากระบบ"
              className="p-1.5 rounded-lg"
              style={{ color: "#C7D2E0" }}
            >
              <LogOut size={16} />
            </button>
          </div>
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-6 py-6">
        {tab === "dashboard" && (
          <>
            <div className="mb-4">
              <div className="text-lg font-semibold" style={{ color: INK }}>ภาพรวม Kaizen ทั้งโรงงาน</div>
              <div className="text-sm" style={{ color: SUB }}>ติดตามสถานะ กรองข้อมูลตามหมวดหมู่ หน่วยงาน และสถานะแบบเรียลไทม์</div>
            </div>
            <Dashboard data={kaizens} departments={departments} onView={(k) => setViewing(k)} onPresent={(k) => setPresenting(k)} />
          </>
        )}

        {tab === "employee" && (
          <>
            <div className="flex items-center justify-between mb-4">
              <div>
                <div className="text-lg font-semibold" style={{ color: INK }}>รายการ Kaizen ของฉัน</div>
                <div className="text-sm" style={{ color: SUB }}>กรอกข้อเสนอ แนบภาพก่อน–หลัง และติดตามสถานะการอนุมัติ</div>
              </div>
              <button
                onClick={() => { setEditing(null); setShowForm(true); }}
                className="flex items-center gap-1.5 px-4 py-2 rounded-lg text-sm font-medium text-white"
                style={{ background: NAVY }}
              >
                <PlusCircle size={16} /> ส่งเรื่อง Kaizen ใหม่
              </button>
            </div>
            <Table
              rows={myKaizens}
              actions={(k) => (
                <>
                  <IconBtn icon={Eye} label="ดู" onClick={() => setViewing(k)} color={NAVY} />
                  {(k.status === "draft" || k.status === "rejected") && (
                    <IconBtn icon={Pencil} label="แก้ไข" onClick={() => { setEditing(k); setShowForm(true); }} color={GOLD} />
                  )}
                  {k.status === "pending_supervisor" && (
                    <IconBtn icon={Undo2} label="ถอนกลับมาแก้ไข" onClick={() => setRecalling(k)} color={RED} />
                  )}
                </>
              )}
            />
          </>
        )}

        {tab === "supervisor" && (
          <>
            <div className="mb-4">
              <div className="text-lg font-semibold" style={{ color: INK }}>รอการอนุมัติจากหัวหน้างาน</div>
              <div className="text-sm" style={{ color: SUB }}>ตรวจสอบรายละเอียด อนุมัติ/ตีกลับ และเปิดนำเสนอให้ผู้บริหารได้ทันที</div>
            </div>
            <Table
              rows={supervisorQueue}
              actions={(k) => (
                <>
                  <IconBtn icon={Presentation} label="นำเสนอ" onClick={() => setPresenting(k)} color={NAVY} />
                  <IconBtn icon={ClipboardList} label="ตรวจสอบ" onClick={() => setReviewing(k)} color={GREEN} filled />
                </>
              )}
            />
            <div className="mt-6 text-sm font-medium" style={{ color: INK }}>ประวัติเรื่องทั้งหมด (สำหรับดูรายละเอียด/นำเสนอ)</div>
            <div className="mt-2">
              <Table
                rows={kaizens.filter((k) => k.status !== "draft" && k.status !== "pending_supervisor" && currentUser.departments.includes(k.department))}
                actions={(k) => (
                  <>
                    <IconBtn icon={Eye} label="ดู" onClick={() => setViewing(k)} color={NAVY} />
                    <IconBtn icon={Presentation} label="นำเสนอ" onClick={() => setPresenting(k)} color={GOLD} />
                  </>
                )}
              />
            </div>
          </>
        )}

        {tab === "admin" && (
          <>
            <div className="mb-4 flex items-start justify-between gap-3">
              <div>
                <div className="text-lg font-semibold" style={{ color: INK }}>ตรวจสอบและจัดหมวดหมู่ (Admin)</div>
                <div className="text-sm" style={{ color: SUB }}>ตรวจสอบข้อมูลการลงทะเบียน กำหนดประเภท Kaizen และหมวดหมู่ QCDSME ก่อนส่งต่อผู้อำนวยการ</div>
              </div>
              <button
                onClick={() => setShowDeptManager((s) => !s)}
                className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-medium border whitespace-nowrap"
                style={{ borderColor: "#D8DDE5", color: NAVY }}
              >
                <Building2 size={14} /> {showDeptManager ? "ซ่อนการจัดการหน่วยงาน" : "จัดการหน่วยงาน"}
              </button>
            </div>
            {showDeptManager && (
              <div className="mb-5">
                <DepartmentManager departments={departments} onChange={setDepartments} />
              </div>
            )}
            <Table
              rows={adminQueue}
              actions={(k) => (
                <>
                  <IconBtn icon={Presentation} label="นำเสนอ" onClick={() => setPresenting(k)} color={NAVY} />
                  <IconBtn icon={Pencil} label="แก้ไขข้อมูล / เพิ่มรูป" onClick={() => setAdminEditing(k)} color="#6A4BC9" />
                  <IconBtn icon={ClipboardList} label="ตรวจ/จัดหมวด" onClick={() => setReviewing(k)} color={GREEN} filled />
                </>
              )}
            />
            <div className="mt-6 flex items-center justify-between gap-3 flex-wrap">
              <div className="text-sm font-medium" style={{ color: INK }}>รายการทั้งหมดในระบบ</div>
              <input
                value={adminSearch}
                onChange={(e) => setAdminSearch(e.target.value)}
                placeholder="ค้นหาเลขที่ / เรื่อง / หน่วยงาน"
                className="rounded-lg border px-3 py-1.5 text-sm w-72"
                style={inputStyle}
              />
            </div>
            <div className="mt-2 flex flex-wrap gap-1.5">
              {[
                ["all", `ทั้งหมด (${kaizens.length})`],
                ["noimage", `ยังไม่มีรูปก่อน-หลัง (${kaizens.filter((k) => !k.images?.before || !k.images?.after).length})`],
                ["incomplete", `ข้อมูลยังไม่ครบ (${kaizens.filter((k) => missingInfo(k).length > 0).length})`],
              ].map(([key, label]) => (
                <button
                  key={key}
                  onClick={() => setAdminListFilter(key)}
                  className="text-xs px-3 py-1.5 rounded-full border font-medium"
                  style={adminListFilter === key ? { background: NAVY, borderColor: NAVY, color: "white" } : { borderColor: "#D8DDE5", color: SUB, background: "white" }}
                >
                  {label}
                </button>
              ))}
            </div>
            <div className="mt-2">
              <Table
                showMissing
                rows={kaizens.filter((k) => {
                  if (adminListFilter === "noimage" && k.images?.before && k.images?.after) return false;
                  if (adminListFilter === "incomplete" && missingInfo(k).length === 0) return false;
                  const q = adminSearch.trim().toLowerCase();
                  return !q || [k.registeredNo, k.title, k.department].some((v) => String(v || "").toLowerCase().includes(q));
                })}
                actions={(k) => (
                  <>
                    <IconBtn icon={Eye} label="ดู" onClick={() => setViewing(k)} color={NAVY} />
                    <IconBtn icon={Pencil} label="แก้ไขข้อมูล / เพิ่มรูป" onClick={() => setAdminEditing(k)} color="#6A4BC9" />
                    <IconBtn icon={Presentation} label="นำเสนอ" onClick={() => setPresenting(k)} color={GOLD} />
                  </>
                )}
              />
            </div>
          </>
        )}

        {tab === "director" && (
          <>
            <div className="mb-4">
              <div className="text-lg font-semibold" style={{ color: INK }}>รออนุมัติขั้นสุดท้าย (ผู้อำนวยการโรงงาน)</div>
              <div className="text-sm" style={{ color: SUB }}>พิจารณาอนุมัติเรื่องที่ผ่านการตรวจสอบจาก Admin แล้ว</div>
            </div>
            <Table
              rows={directorQueue}
              actions={(k) => (
                <>
                  <IconBtn icon={Presentation} label="นำเสนอ" onClick={() => setPresenting(k)} color={NAVY} />
                  <IconBtn icon={ClipboardList} label="พิจารณา" onClick={() => setReviewing(k)} color={GREEN} filled />
                </>
              )}
            />
            <div className="mt-6 text-sm font-medium" style={{ color: INK }}>ประวัติเรื่องที่อนุมัติแล้ว</div>
            <div className="mt-2">
              <Table
                rows={kaizens.filter((k) => k.status === "approved" || k.status === "rejected")}
                actions={(k) => (
                  <>
                    <IconBtn icon={Eye} label="ดู" onClick={() => setViewing(k)} color={NAVY} />
                    <IconBtn icon={Presentation} label="นำเสนอ" onClick={() => setPresenting(k)} color={GOLD} />
                  </>
                )}
              />
            </div>
          </>
        )}
        {tab === "judge" && (
          <JudgePage
            kaizens={kaizens}
            award={award}
            currentUser={currentUser}
            onScore={scoreAward}
            onView={setViewing}
            onPresent={setPresenting}
            onSubmitted={judgeSubmitted}
            onRefresh={loadAward}
            syncMode={awardSync}
          />
        )}
        {tab === "award" && (
          <AwardPage
            kaizens={kaizens}
            award={award}
            onChange={changeAward}
            onScore={scoreAward}
            onSendToJudges={sendToJudges}
            onRefresh={loadAward}
            currentUser={currentUser}
            onView={setViewing}
            onPresent={setPresenting}
            onAnnounce={announceAward}
            syncMode={awardSync}
          />
        )}
      </div>

      {showForm && (
        <Modal title={editing ? "แก้ไขข้อเสนอ Kaizen" : "ส่งข้อเสนอ Kaizen ใหม่"} onClose={() => { setShowForm(false); setEditing(null); }} wide>
          <KaizenForm initial={editing} currentUser={currentUser} departments={departments} onCancel={() => { setShowForm(false); setEditing(null); }} onSave={saveForm} />
        </Modal>
      )}

      {viewing && (
        <Modal
          title="รายละเอียด Kaizen"
          onClose={() => setViewing(null)}
          wide
          headerExtra={
            currentUser?.role === "admin" && (
              <button
                onClick={() => { setAdminEditing(viewing); setViewing(null); }}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-white"
                style={{ background: "#6A4BC9" }}
              >
                <Pencil size={13} /> แก้ไขข้อมูล / เพิ่มรูป
              </button>
            )
          }
        >
          <DetailView k={viewing} />
        </Modal>
      )}

      {adminEditing && (
        <Modal title="แก้ไขข้อมูล Kaizen (Admin)" onClose={() => setAdminEditing(null)} wide>
          <KaizenForm
            key={adminEditing.id}
            initial={adminEditing}
            adminMode
            currentUser={currentUser}
            departments={departments}
            onCancel={() => setAdminEditing(null)}
            onSave={adminSave}
          />
        </Modal>
      )}

      {presenting && (
        <Modal title="โหมดนำเสนอ" onClose={() => setPresenting(null)} wide>
          <PresentationView k={presenting} />
        </Modal>
      )}

      {reviewing && (
        <ReviewModal
          k={reviewing}
          roleLabel={NAV.find((n) => n.key === tab)?.label}
          allowCategorize={tab === "admin"}
          onClose={() => setReviewing(null)}
          onDecide={(action, comment, extra, returnTo) => decide(reviewing, action, comment, extra, returnTo)}
        />
      )}

      {recalling && (
        <Modal title="ยืนยันการถอนเรื่อง" onClose={() => setRecalling(null)}>
          <div className="space-y-4">
            <p className="text-sm" style={{ color: INK }}>
              ต้องการถอนเรื่อง <strong>&ldquo;{recalling.title}&rdquo;</strong> กลับมาเป็น &ldquo;แบบร่าง&rdquo; เพื่อแก้ไขใช่หรือไม่? หัวหน้างานจะไม่เห็นเรื่องนี้ในคิวรออนุมัติอีกต่อไปจนกว่าท่านจะแก้ไขและส่งใหม่
            </p>
            <div className="flex justify-end gap-2">
              <button onClick={() => setRecalling(null)} className="px-4 py-2 rounded-lg text-sm font-medium border" style={{ borderColor: "#D8DDE5", color: SUB }}>
                ยกเลิก
              </button>
              <button onClick={() => recallToDraft(recalling)} className="px-4 py-2 rounded-lg text-sm font-medium text-white flex items-center gap-1.5" style={{ background: RED }}>
                <Undo2 size={15} /> ยืนยันถอนเรื่อง
              </button>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}
