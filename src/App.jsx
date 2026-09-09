import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  LayoutDashboard, User, ShieldCheck, Building2, PlusCircle, Eye, Pencil,
  CheckCircle2, XCircle, Upload, X, Filter, ClipboardList, Presentation,
  Users, TrendingUp, ChevronRight, ImageOff, LogOut, Lock, KeyRound, AlertTriangle, Clock,
  Download, Coins, Bell, BellRing, Undo2, Paperclip, FileVideo, FileSpreadsheet, FileText, ZoomIn
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell,
  CartesianGrid, LineChart, Line, Legend
} from "recharts";
import * as XLSX from "xlsx";

// ---------- Design tokens ----------
const COMPANY_LOGO = "/logo.png";

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

const ROLE_LABEL = { employee: "พนักงาน", supervisor: "หัวหน้างาน", admin: "Admin (ฝ่ายบุคคล)", director: "ผู้อำนวยการโรงงาน" };

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
                const reader = new FileReader();
                reader.onload = () => onChange(reader.result);
                reader.readAsDataURL(f);
              }}
            />
          </label>
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

function KaizenForm({ initial, onCancel, onSave, currentUser, departments }) {
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

  return (
    <div className="space-y-4">
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
          h.action === "recall" ? "ถอนเรื่องกลับมาแก้ไข (แบบร่าง)" : "";
        const [datePart, timePart] = (h.date || "").split(" ");
        return (
          <div key={i} className="flex gap-3">
            <div className="flex flex-col items-center">
              <div className="w-2.5 h-2.5 rounded-full mt-1" style={{ background: isNegative ? RED : GREEN }} />
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

function Table({ rows, actions }) {
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
            <img src={COMPANY_LOGO} alt="Company Logo" className="w-full h-full object-contain" />
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
            {MOCK_USERS.map((u) => (
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
  const [recalling, setRecalling] = useState(null);
  const [notifications, setNotifications] = useState(() => buildInitialNotifications(kaizens));
  const [showNotifs, setShowNotifs] = useState(false);

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
      const registeredNo = `KZ-2026-0${20 + kaizens.length}`;
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
              <img src={COMPANY_LOGO} alt="Company Logo" className="w-full h-full object-contain" />
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
                  <IconBtn icon={ClipboardList} label="ตรวจ/จัดหมวด" onClick={() => setReviewing(k)} color={GREEN} filled />
                </>
              )}
            />
            <div className="mt-6 text-sm font-medium" style={{ color: INK }}>รายการทั้งหมดในระบบ</div>
            <div className="mt-2">
              <Table
                rows={kaizens}
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
      </div>

      {showForm && (
        <Modal title={editing ? "แก้ไขข้อเสนอ Kaizen" : "ส่งข้อเสนอ Kaizen ใหม่"} onClose={() => { setShowForm(false); setEditing(null); }} wide>
          <KaizenForm initial={editing} currentUser={currentUser} departments={departments} onCancel={() => { setShowForm(false); setEditing(null); }} onSave={saveForm} />
        </Modal>
      )}

      {viewing && (
        <Modal title="รายละเอียด Kaizen" onClose={() => setViewing(null)} wide>
          <DetailView k={viewing} />
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
