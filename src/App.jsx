import React, { useState, useMemo, useRef, useEffect, createContext, useContext } from "react";
import {
  LayoutDashboard, User, ShieldCheck, Building2, PlusCircle, Eye, Pencil,
  CheckCircle2, XCircle, Upload, X, Filter, ClipboardList, Presentation,
  Users, TrendingUp, ChevronRight, ImageOff, LogOut, Lock, KeyRound, AlertTriangle, Clock,
  Download, Coins, Bell, BellRing, Undo2, Paperclip, FileVideo, FileSpreadsheet, FileText, ZoomIn,
  Trophy, Medal, Plus, Trash2, Star, ImagePlus, Printer, ChevronLeft, ChevronDown, Search
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell,
  CartesianGrid, LineChart, Line, Legend, LabelList
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

// ส่งข้อมูลไป Apps Script แล้วตรวจว่า "บันทึกสำเร็จจริง" (เซิร์ฟเวอร์ต้องตอบ success) ไม่ใช่แค่ส่งออกไปได้
async function postToSheet(action, record) {
  try {
    const res = await fetch(GOOGLE_SHEETS_API_URL, {
      method: "POST",
      // ใช้ text/plain เพื่อเลี่ยงปัญหา CORS preflight ของ Google Apps Script
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action, record }),
    });
    const text = await res.text();
    let json = null;
    try { json = JSON.parse(text); } catch (e) { /* ไม่ใช่ JSON (เช่น หน้า error ของ Apps Script) */ }
    if (res.ok && json && json.success !== false) return { ok: true, error: "" };
    const detail = json && json.error ? json.error : text.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 180);
    return { ok: false, error: res.ok ? detail || "เซิร์ฟเวอร์ตอบกลับไม่ถูกต้อง" : `HTTP ${res.status} ${detail}` };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
}

// สถานะการบันทึก Kaizen ลง Google Sheet (แสดงมุมขวาล่าง เมื่อเชื่อมต่อ Google Sheet อยู่)
const syncState = { pending: 0, failed: new Map(), tracking: false };
const syncSubscribers = new Set();
function emitSync() {
  const snap = { pending: syncState.pending, failed: Array.from(syncState.failed.values()) };
  syncSubscribers.forEach((fn) => fn(snap));
}
async function syncKaizenToSheet(action, record) {
  const payload = toSheetRecord(record);
  if (!syncState.tracking) {
    // โหมดทดลอง/ยังไม่เชื่อมชีต: ส่งแบบเงียบๆ เหมือนเดิม (คาดว่าจะ error ถ้าทดสอบใน claude.ai artifact)
    const r = await postToSheet(action, payload);
    if (!r.ok) console.warn("Sync to Google Sheet failed:", r.error);
    return r.ok;
  }
  syncState.pending += 1;
  emitSync();
  const r = await postToSheet(action, payload);
  syncState.pending -= 1;
  if (r.ok) syncState.failed.delete(record.id);
  else syncState.failed.set(record.id, { id: record.id, title: record.title || record.registeredNo || record.id, record, error: r.error });
  emitSync();
  return r.ok;
}
// ลองบันทึกซ้ำ: ใช้ "update" เสมอ (ถ้ายังไม่มีแถวนี้ในชีต ระบบจะเพิ่มให้ ไม่ซ้ำ)
function retryFailedSyncs() {
  Array.from(syncState.failed.values()).forEach((f) => syncKaizenToSheet("update", f.record));
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
  return (await postToSheet(action, record)).ok;
}
// รายชื่อผู้ใช้ที่ Admin จัดการ (เก็บในชีต KaizenUsers ผ่าน Users_Addon.gs)
async function fetchUsersFromSheet() {
  const res = await fetch(`${GOOGLE_SHEETS_API_URL}?type=users`);
  if (!res.ok) throw new Error("users fetch failed");
  const json = await res.json();
  if (json && !Array.isArray(json) && Object.prototype.hasOwnProperty.call(json, "users")) return { supported: true, users: json.users };
  return { supported: false, users: null }; // Code.gs ยังไม่ได้ติดตั้ง Users_Addon.gs
}
function saveUsersToSheet(list) {
  return postAward("saveUsers", { id: "users", data: list });
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

// ---------- ทะเบียนผู้ใช้ ----------
// รายชื่อเริ่มต้นด้านล่างเก็บรหัสผ่านเป็นค่าเข้ารหัส (PBKDF2-SHA256) ไม่มีรหัสผ่านจริงอยู่ในโค้ด
// Admin เพิ่ม/แก้/ลบผู้ใช้ได้ที่เมนู "จัดการผู้ใช้" ถ้าเชื่อม Google Sheet และติดตั้ง Users_Addon.gs
// รายชื่อที่ Admin แก้จะถูกเก็บในชีต "KaizenUsers" และใช้แทนรายการเริ่มต้นนี้
// หมายเหตุ: ยังเป็นการตรวจรหัสผ่านฝั่งเบราว์เซอร์ ระบบจริงควรใช้ SSO/backend ของบริษัท
// aliases = ชื่อบัญชีเดิม (ใช้ให้ Kaizen/การแจ้งเตือนที่เคยส่งด้วยชื่อเดิมยังแสดงในบัญชีที่เปลี่ยนชื่อ)
const DEFAULT_USERS = [
  { id: "u01", name: "Admin", role: "admin", departments: [], aliases: [], active: true, pw: { s: "50081db53d9e2f85d842cef08e725a60", h: "bb6eeebf05ed16a164bcdb603ecc00b82565c05596cbe1292fa0c44d809052aa" } },
  { id: "u02", name: "ผอ.โรงงาน", role: "director", departments: [], aliases: [], active: true, pw: { s: "64153b2197fdd931365a3ca374226a63", h: "88df32eb05c800025f81b21a9b792911e6657c3238d59560894e144f71baaacd" } },
  { id: "u03", name: "ทิพวรรณ บุญจิตติ", role: "supervisor", departments: ["Production Planning & Data"], aliases: [], active: true, pw: { s: "59377386718b60a30ed6af92277c3e15", h: "0ff8e98b96b2373f9307513b926e328414b5941b7cecc0c84525171027875ed3" } },
  { id: "u04", name: "ธัญยธรณ์ สุขุมาลพันธ์", role: "supervisor", departments: ["Washing - Drying 1", "Thread forming 1", "Packing 1", "Thread forming 3", "Packing 3"], aliases: [], active: true, pw: { s: "5770f85f9254f3833c4fb4b1dd7d0de7", h: "d10b200eafdfdd48971398b471366d0b365988ac4d6d7fd56e1b5f4dd18ba748" } },
  { id: "u05", name: "นรินทร์ คำนา", role: "supervisor", departments: ["Drying 2", "Packing 2", "RTE"], aliases: [], active: true, pw: { s: "756aeef5789bd7fef59e9558368fd74b", h: "4d33f66fb64f649002252ac65aa5302a0a0cde6999a1ef311881adeaf5f4623d" } },
  { id: "u06", name: "นรินทร์ คำนา", role: "supervisor", departments: ["Thread forming _ RN", "Thread forming _ RV", "Cutting", "Drying 4", "Packing 4", "BS"], aliases: [], active: true, pw: { s: "fa738d8ad6c45efe52822c663c25da2f", h: "705e49a3b365e29f97fa0ffd30b1c77db97b346869da1b4d5474466b329513ee" } },
  { id: "u07", name: "อมรรัตน์ ตุลาผล", role: "supervisor", departments: ["QA", "Pest Control", "DC", "Sanitation", "QC", "Lab"], aliases: [], active: true, pw: { s: "81544a85b51f683c1aa409114899cd35", h: "2fcf985546fe62b6196b1a15ea0b18f3eb6398fe4a490273ee2312408d616523" } },
  { id: "u08", name: "สุพิชชา วังถนอมศักดิ์", role: "supervisor", departments: ["RD"], aliases: [], active: true, pw: { s: "3b0c94da06f5e1ccae0d68afbd3aea17", h: "995eefbb16b0531bb8ce41b9e88e08b248bb7c9bb2ed2dc1c10310dce6d00ba4" } },
  { id: "u09", name: "วันัย บุญมุนัส", role: "supervisor", departments: ["Engineering Data", "Engineering Team A", "Engineering Team B", "Engineering Team C", "Boiler", "Water supply"], aliases: ["วินัย บุญมุนัส"], active: true, pw: { s: "d11925bd1c56d7eabf92c9cf9fce8c30", h: "c15c1b99a5f00559c4ffdc7bc7471a779060f0687be9255aa083d7ea2b610ad7" } },
  { id: "u10", name: "นันทวรรณ นวลเนตร", role: "supervisor", departments: ["Raw Material W/H", "Finish Goods W/H (Dry)", "Finish Goods W/H (Fresh)", "Spart part - Eguipment W/H", "Packaging W/H"], aliases: [], active: true, pw: { s: "7f6991995984663879737317288a3a17", h: "c44bcda60a50952563e718640b47cf9d2d7e1951f39aa40b407ec29a05d96c09" } },
  { id: "u11", name: "วิไลวรรณ ขำขม", role: "supervisor", departments: ["Safety", "Envi"], aliases: [], active: true, pw: { s: "5405918e401aa484797af6b5ddd58817", h: "05ea7dff51efa8144f11378c83df322e7941899b5c82bf25de790854df7a2a8a" } },
  { id: "u12", name: "ติณณา พุทธสุข", role: "supervisor", departments: ["C & B", "Recruitment & Training", "ER & Admin"], aliases: [], active: true, pw: { s: "5fd1fe419fa48351b9615eaf0b16c9f7", h: "63204ea50fe6b99b1e5a3bb39cc01fe15a84522d55a1245db61fbc0f79ff3075" } },
  { id: "u13", name: "Production Planning & Data", role: "employee", departments: ["Production Planning & Data"], aliases: [], active: true, pw: { s: "f4f3241a714931592f37180a09b68641", h: "835ecce80d72e384ec3d1c252a92143b2d868eee1eaf76dbbe6e7083219b4553" } },
  { id: "u14", name: "Production 1", role: "employee", departments: ["Washing - Drying 1", "Thread forming 1", "Packing 1"], aliases: ["Production_Line 1"], active: true, pw: { s: "2bbe431cf612de4580e0adaff5803901", h: "a491415b78a7e17db02bfb625e4d4808efd64c9d7cff159d81fa907ca364dbd4" } },
  { id: "u15", name: "Production 3", role: "employee", departments: ["Thread forming 3", "Packing 3"], aliases: ["Production Line 3"], active: true, pw: { s: "56f707780601649c01be4fb7232d408b", h: "acb925dec60443e84cc9aa673995c4ca4f17bdb441f63e390cb97e7a09883c3d" } },
  { id: "u16", name: "Production 2", role: "employee", departments: ["Drying 2", "Packing 2", "RTE"], aliases: ["Production Line 2"], active: true, pw: { s: "66515446d5a300d9f7451d6ef33e87c2", h: "91bb0c985298706b1882c270d97fcc47d34bae7a1ef56a7abbb93cfc3b1bba88" } },
  { id: "u17", name: "Production 4", role: "employee", departments: ["Thread forming _ RN", "Thread forming _ RV", "Cutting", "Drying 4", "Packing 4", "BS"], aliases: ["Production Line 4"], active: true, pw: { s: "5fb8c0cd48ed5ace7bef75c71fef9254", h: "6b02336dcc7a25d29f0216e1318c53f56cf5107af9ca4c9fd0620ea364fa35f0" } },
  { id: "u18", name: "QC", role: "employee", departments: ["QC", "Lab"], aliases: [], active: true, pw: { s: "186443ca465a0e676054423be7683d7a", h: "e273431edd0049c20d8fb0015ba4ba671fd70b73aab92569d8134933c48ed76e" } },
  { id: "u19", name: "QA", role: "employee", departments: ["QA", "Pest Control", "DC", "Sanitation"], aliases: [], active: true, pw: { s: "f4212404b5babc59bbd68629c72526a3", h: "0fe4d336591467e69acb224956ebf59847d6e805d338ed652fdcaf936e1735b8" } },
  { id: "u20", name: "RD", role: "employee", departments: ["RD"], aliases: [], active: true, pw: { s: "e7c358c173a1053de3277fc26d7961ab", h: "14dcabf8f9f5db6b668c116625c3bff7375614f3b20d9301e03f1e24e008b85c" } },
  { id: "u21", name: "Engineering Data", role: "employee", departments: ["Engineering Data"], aliases: [], active: true, pw: { s: "35c3e372699cc2f37d6ba0b31924d2aa", h: "3999ab199b479d8b27730f170ff957ab8acbc89c4b53c79b499e2de4c468fc26" } },
  { id: "u22", name: "Engineering Team A", role: "employee", departments: ["Engineering Team A"], aliases: [], active: true, pw: { s: "24da189b0a09e93390ca1e794281fada", h: "3103bef797c6060187260c0fc7b8abb5fa568e632c87fbf2c6d058cc83d25034" } },
  { id: "u23", name: "Engineering Team B", role: "employee", departments: ["Engineering Team B"], aliases: [], active: true, pw: { s: "5005e873d9f76bc62ddaa6831138aa8f", h: "9849fc0babaa1f2cc924e4f67650e6de8d096c1ce9f77b23bd3b6d2aa9d903ef" } },
  { id: "u24", name: "Engineering Team C", role: "employee", departments: ["Engineering Team C"], aliases: [], active: true, pw: { s: "1bc87bed64754192b7c33d55b9b4917c", h: "81499c8ffff4ade108d85e545aa6175f8942b298252f90766463f0503e2b201e" } },
  { id: "u25", name: "Utility", role: "employee", departments: ["Boiler", "Water supply"], aliases: ["Boiler", "Water supply"], active: true, pw: { s: "b7eb42128e444dbb8d192f5ca198220a", h: "8f23ac3cfefb5991978e867bd98667a2b1491b6df835f3ac95d8813cb5c3625a" } },
  { id: "u26", name: "Finish Goods W/H (Dry)", role: "employee", departments: ["Finish Goods W/H (Dry)"], aliases: [], active: true, pw: { s: "9f31ef3d51b4d73e6aa7978fd85b2d5b", h: "ac27e1df1cc4365d43c1959aa93658eff7102337bee0cfb71712e7b86146f776" } },
  { id: "u27", name: "Finish Goods W/H (Fresh)", role: "employee", departments: ["Finish Goods W/H (Fresh)"], aliases: [], active: true, pw: { s: "ae67e4c8d8166fa62de250f2603eefc8", h: "6f397a5815fbd2cd58538ff25dcc9b082dbee64634c58ba606d6b991b889e55c" } },
  { id: "u28", name: "Raw Material W/H", role: "employee", departments: ["Raw Material W/H"], aliases: [], active: true, pw: { s: "c91663a7bc15cc07630005626f4b2d47", h: "da5ab18e9d4e94061365a7eaa78911c69544fb27ba1bf9e97ea2275e93324aa9" } },
  { id: "u29", name: "Packaging W/H", role: "employee", departments: ["Packaging W/H"], aliases: [], active: true, pw: { s: "649cea5b0176cce4ff53515a7a97bf61", h: "1f464a549da68ae99eaf9cb89321d3fbe1fbe73962f3d06a47e5f219e6a600a2" } },
  { id: "u30", name: "Spart part - Eguipment W/H", role: "employee", departments: ["Spart part - Eguipment W/H"], aliases: [], active: true, pw: { s: "d11c4e85f233ca5cead71e7fb2a26cda", h: "2523c06d97dee8a9424e00f39bea27c0293092b66ecb6579218000845a5076d3" } },
  { id: "u31", name: "Safety", role: "employee", departments: ["Safety"], aliases: [], active: true, pw: { s: "5423899366b1096d1b7e78864c432401", h: "d2412364918560739e6154e2b87eb05e74d3cf5425773cb976283c8d9f96b644" } },
  { id: "u32", name: "Envi", role: "employee", departments: ["Envi"], aliases: [], active: true, pw: { s: "4382099fd7018720548cbf289fa26f3d", h: "cb1f558a0da4a8bfd11b9f41e6924dee9b3546efc0ef0112123c0d551e9b04f3" } },
  { id: "u33", name: "C & B", role: "employee", departments: ["C & B"], aliases: [], active: true, pw: { s: "f854025d869d846d2e767a2731a5b599", h: "5af510031ac282c2ebbb3364fb13754da4d1effa54a66880fef24de4dd4d75bc" } },
  { id: "u34", name: "Recruitment & Training", role: "employee", departments: ["Recruitment & Training"], aliases: [], active: true, pw: { s: "11e6fb7e46adafed8f39f56347c1592f", h: "03fe32a40160fcd501d61be03ffd5b95dcf955430191822403cb2407fdab64bf" } },
  { id: "u35", name: "ER & Admin", role: "employee", departments: ["ER & Admin"], aliases: [], active: true, pw: { s: "db5ff6c83eb52a7e0ba831bba3e54951", h: "060d60f91357bb469dd86b0134f24772e4e0194537ce92b35d211293375afdc5" } },
  { id: "u36", name: "กรรมการ 1", role: "judge", departments: [], aliases: [], active: true, pw: { s: "dd35b6aae8872510298215bf52bc7d4d", h: "9f560ea13927eec82318e4a70aa10cac09653d854d3f77c96047f13fe1f58dac" } },
  { id: "u37", name: "กรรมการ 2", role: "judge", departments: [], aliases: [], active: true, pw: { s: "d58b20452b8b853ae2e5116e5ddc6a55", h: "7329bfbcdbd161c060eae62805c3541d7c9b492f4c82450823686f2504cbe5e7" } },
  { id: "u38", name: "กรรมการ 3", role: "judge", departments: [], aliases: [], active: true, pw: { s: "b535e3398e15b90e555c8f5cca9aa67e", h: "37a102a5d2f80c97cfdc31e02ed4daa71a91b8ddd3cd24f483b17543d0a411ab" } },
  { id: "u39", name: "กรรมการ 4", role: "judge", departments: [], aliases: [], active: true, pw: { s: "5e8dc07b8ae990c670c2a1eb931c23a5", h: "958c7993cdb4ab193362ef23ff2637a39951769dd971193a88187b10905d460b" } },
  { id: "u40", name: "กรรมการ 5", role: "judge", departments: [], aliases: [], active: true, pw: { s: "4052c37e9481236f96c69ce528142e3d", h: "3e41024e10ec474904660929f22dbfe1b4bb042694bfcf72331b1f4dfe584c5e" } },
  { id: "u41", name: "กรรมการ 6", role: "judge", departments: [], aliases: [], active: true, pw: { s: "80cebfa433ef0808feb98280c3e8c59e", h: "8e54af119600db342b532674f375c1cfda408f5f44b5fec677e2731146574e54" } },
  { id: "u42", name: "กรรมการ 7", role: "judge", departments: [], aliases: [], active: true, pw: { s: "f0bc061d861d2b6021c49fc0f681e89b", h: "ddc52dddef424f0152e28a461c0ba19ad572f5de768b29ca311aa277a8f1fff3" } },
];

// รายชื่อที่ใช้งานจริง (อัปเดตเมื่อโหลดจาก Google Sheet หรือ Admin แก้ไข)
const USERS = DEFAULT_USERS.map((u) => ({ ...u }));

const PBKDF2_ITERATIONS = 100000;
const hexToBytes = (hex) => new Uint8Array((hex.match(/.{1,2}/g) || []).map((b) => parseInt(b, 16)));
const bytesToHex = (buf) => Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
async function hashPassword(password, saltHex) {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveBits"]);
  const bits = await crypto.subtle.deriveBits({ name: "PBKDF2", hash: "SHA-256", salt: hexToBytes(saltHex), iterations: PBKDF2_ITERATIONS }, key, 256);
  return bytesToHex(bits);
}
async function makePw(password) {
  const s = bytesToHex(crypto.getRandomValues(new Uint8Array(16)));
  return { s, h: await hashPassword(password, s) };
}
async function checkPw(user, password) {
  if (!user || !user.pw || !user.pw.s || !user.pw.h) return false;
  return (await hashPassword(password, user.pw.s)) === user.pw.h;
}
function randomPassword(len = 8) {
  const chars = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789";
  const arr = crypto.getRandomValues(new Uint32Array(len));
  return Array.from(arr).map((n) => chars[n % chars.length]).join("");
}
const USER_ROLES = ["employee", "supervisor", "admin", "director", "judge"];
const newUserId = () => `u_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 5)}`;
const normName = (n) => String(n || "").trim().toLowerCase();

// ตรวจและเติมค่าที่ขาดของรายชื่อที่โหลดจาก Google Sheet (กันข้อมูลเสียทำให้เข้าระบบไม่ได้)
function normalizeUsers(list) {
  const out = (Array.isArray(list) ? list : [])
    .filter((u) => u && u.id && String(u.name || "").trim() && USER_ROLES.includes(u.role) && u.pw && u.pw.s && u.pw.h)
    .map((u) => ({
      id: String(u.id), name: String(u.name).trim(), role: u.role,
      departments: Array.isArray(u.departments) ? u.departments : [],
      aliases: Array.isArray(u.aliases) ? u.aliases : [],
      active: u.active !== false, pw: { s: u.pw.s, h: u.pw.h },
    }));
  // กรรมการ 1-7 ต้องมีเสมอ (ระบบประกวดอ้างอิงชื่อบัญชีเหล่านี้)
  DEFAULT_USERS.filter((d) => d.role === "judge" && !out.some((u) => normName(u.name) === normName(d.name))).forEach((d) => out.push({ ...d }));
  // ต้องมี Admin ที่ใช้งานได้อย่างน้อย 1 บัญชี ไม่เช่นนั้นกลับไปใช้รายชื่อเริ่มต้น
  if (!out.some((u) => u.role === "admin" && u.active)) return DEFAULT_USERS.map((u) => ({ ...u }));
  return out;
}


const ROLE_LABEL = { employee: "พนักงาน", supervisor: "หัวหน้างาน", admin: "Admin (ฝ่ายบุคคล)", director: "ผู้อำนวยการโรงงาน", judge: "กรรมการตัดสิน Kaizen Award" };

function buildInitialNotifications(kaizenList) {
  const list = [];
  const push = (recipientName, message, kaizenId) => {
    if (!recipientName) return;
    list.push({ id: uid(), recipientName, message, kaizenId, read: false, date: nowStamp() });
  };
  kaizenList.forEach((k) => {
    if (k.status === "pending_supervisor") {
      const sup = USERS.find((u) => u.role === "supervisor" && u.departments.includes(k.department));
      if (sup) push(sup.name, `มีเรื่อง Kaizen ใหม่รอคุณอนุมัติ: "${k.title}"`, k.id);
    } else if (k.status === "pending_admin") {
      const admin = USERS.find((u) => u.role === "admin");
      if (admin) push(admin.name, `มีเรื่อง Kaizen รอ Admin ตรวจสอบ: "${k.title}"`, k.id);
    } else if (k.status === "pending_director") {
      const director = USERS.find((u) => u.role === "director");
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

// รายชื่อหน่วยงานที่ Admin เพิ่ม/แก้/ลบเอง: เก็บในชีต KaizenUsers เป็นแถวพิเศษ id "__config" (ใช้ Code.gs ตัวเดิมได้ ไม่ต้องแก้)
// SECTION_MAP ถูกอัปเดตตามรายการนี้ (หน่วยงาน -> ฝ่าย) เพื่อให้แดชบอร์ดจัดกลุ่มได้ถูกต้อง
const BASE_SECTION_MAP = { ...SECTION_MAP };
const BASE_SECTION_OPTIONS = [...SECTION_OPTIONS];
function applyDeptConfig(cfg) {
  const deps = Array.isArray(cfg && cfg.departments) ? cfg.departments.filter((d) => typeof d === "string" && d.trim()) : [];
  if (!deps.length) return null;
  const sections = cfg.sections && typeof cfg.sections === "object" ? cfg.sections : {};
  Object.keys(SECTION_MAP).forEach((k) => delete SECTION_MAP[k]);
  deps.forEach((d) => {
    const sec = sections[d] || BASE_SECTION_MAP[d];
    if (sec) SECTION_MAP[d] = sec;
  });
  // รายการฝ่าย (Section) ที่ Admin เพิ่ม/แก้/ลบ (ไม่มีข้อมูลนี้ = ใช้ฝ่ายเริ่มต้น)
  const sl = Array.isArray(cfg.sectionList) ? cfg.sectionList.filter((x) => typeof x === "string" && x.trim()) : [];
  SECTION_OPTIONS.splice(0, SECTION_OPTIONS.length, ...(sl.length ? sl : BASE_SECTION_OPTIONS));
  deps.forEach((d) => {
    const sec = SECTION_MAP[d];
    if (sec && !SECTION_OPTIONS.includes(sec)) SECTION_OPTIONS.push(sec);
  });
  return deps;
}
const buildDeptConfigRow = (deps) => ({
  id: "__config", kind: "departments", departments: deps, sectionList: [...SECTION_OPTIONS],
  sections: Object.fromEntries(deps.map((d) => [d, SECTION_MAP[d] || "อื่นๆ"])),
});


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
        { stage: "หัวหน้างาน", actor: "ธัญยธรณ์ สุขุมาลพันธ์", action: "approve", comment: "เห็นด้วย ช่วยลดความเสี่ยงจริง", date: "2026-08-04 10:45" },
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

// ================= ตัวช่วยกลาง: รูปหลายภาพ · ชื่อ · ดาว Kaizen Award · ตัวกรอง Multi Select =================
const MAX_PHOTOS = 10;

const asList = (v) => {
  if (Array.isArray(v)) return v.filter(Boolean);
  if (!v) return [];
  if (typeof v === "string" && v.trim().startsWith("[")) {
    try {
      const p = JSON.parse(v);
      return Array.isArray(p) ? p.filter(Boolean) : [];
    } catch (e) {
      return [v];
    }
  }
  return [v];
};

// ชื่อที่ถูกแก้ไข (แสดงชื่อใหม่ในประวัติเดิมด้วย)
const NAME_FIX = { "ธัญญธรณ์ สุขุมาลพันธ์": "ธัญยธรณ์ สุขุมาลพันธ์" };
const fixName = (n) => NAME_FIX[n] || n;

function normalizeKaizen(k) {
  const history = Array.isArray(k.history) ? k.history : [];
  const atts = Array.isArray(k.attachments) ? k.attachments : [];
  const photoAtt = (side) => atts.filter((a) => a && a.kind === "photo" && a.side === side).map((a) => a.dataUrl).filter(Boolean);
  return {
    ...k,
    images: { before: [...asList(k.images?.before), ...photoAtt("before")], after: [...asList(k.images?.after), ...photoAtt("after")] },
    categories: Array.isArray(k.categories) ? k.categories : [],
    attachments: atts.filter((a) => !(a && a.kind === "photo")),
    members: Array.isArray(k.members) ? k.members : [],
    history: history.map((h) => ({ ...h, actor: fixName(h.actor) })),
    activityType: k.activityType || "Kaizen",
    kaizenType: k.kaizenType || "Kaizen Project",
  };
}

// รูปแรกของแต่ละด้านเก็บในช่องรูปเดิม (ลิงก์เดียว) ส่วนรูปที่ 2-10 เก็บเป็น "ไฟล์แนบชนิดรูป" (kind: "photo")
// ทำให้ Code.gs เดิมรับได้ทุกรูปโดยไม่ต้องแก้ Code.gs (ส่งเป็นรายการในช่องรูปเดียว Code.gs เดิมจะบันทึกไม่ได้)
function toSheetRecord(record) {
  const before = asList(record.images?.before);
  const after = asList(record.images?.after);
  const extra = (side, list) =>
    list.slice(1).map((u, i) => ({ kind: "photo", side, name: `${record.id}_${side}_${i + 2}.jpg`, type: "image/jpeg", dataUrl: u }));
  return {
    ...record,
    images: { before: before[0] || null, after: after[0] || null },
    attachments: [...(record.attachments || []).filter((a) => !(a && a.kind === "photo")), ...extra("before", before), ...extra("after", after)],
  };
}

// Kaizen ที่นับเป็น "ยอดการส่ง Kaizen": หัวหน้างานอนุมัติแล้ว และประเภทกิจกรรม = Kaizen (ไม่นับ Suggestion)
const SUPERVISOR_APPROVED = ["pending_admin", "pending_director", "approved"];
const isCountedKaizen = (k) => SUPERVISOR_APPROVED.includes(k.status) && (k.activityType || "Kaizen") === "Kaizen";

// สิทธิ์แก้ไขข้อมูล: Admin แก้ได้ทุกเรื่อง · หัวหน้างานแก้ได้ทุกเรื่องที่หน่วยงานตัวเองส่งเข้ามา (ที่ไม่ใช่แบบร่าง)
const canEditKaizen = (user, k) =>
  !!user && !!k && (user.role === "admin" || (user.role === "supervisor" && k.status !== "draft" && (user.departments || []).includes(k.department)));

// เก็บค่าไว้ในแท็บเบราว์เซอร์ (รีเฟรชแล้วอยู่หน้าเดิม ปิดแท็บแล้วหายเอง เหมาะกับเครื่องที่ใช้ร่วมกัน)
function useStoredState(key, initial) {
  const [v, setV] = useState(() => {
    try {
      const raw = sessionStorage.getItem(key);
      return raw === null ? initial : JSON.parse(raw);
    } catch (e) {
      return initial;
    }
  });
  useEffect(() => {
    try {
      sessionStorage.setItem(key, JSON.stringify(v));
    } catch (e) {}
  }, [key, v]);
  return [v, setV];
}

const AwardCtx = createContext(null);

function ActivityHelp({ className = "", light }) {
  return (
    <div className={`text-[11px] leading-relaxed ${className}`} style={{ color: light ? "rgba(255,255,255,0.62)" : "#A3ABB8" }}>
      <span className="font-medium">Kaizen</span> – คือการปรับปรุงงานที่ได้ลงมือทำแล้ว , <span className="font-medium">Suggestion</span> – เป็นแค่ความคิดเห็น ไอเดีย ยังไม่ลงมือทำ
    </div>
  );
}

// ---------- ดาวตามเกณฑ์ Kaizen Award ----------
function starLevel(k, award) {
  const cfg = award || DEFAULT_AWARD;
  const pct = prescreen(k, cfg);
  const t = cfg.starThresholds || DEFAULT_AWARD.starThresholds;
  const n = pct >= t[0] ? 5 : pct >= t[1] ? 4 : pct >= t[2] ? 3 : pct >= t[3] ? 2 : pct > 0 ? 1 : 0;
  return { pct, n };
}

function StarRating({ k, size = 13, showPct = false, className = "" }) {
  const award = useContext(AwardCtx);
  const { pct, n } = starLevel(k, award);
  return (
    <span
      className={`inline-flex items-center gap-0.5 ${className}`}
      title={`ระดับ ${n}/5 ดาว · คะแนนประเมินตามเกณฑ์ Kaizen Award ${pct}% (Cost Saving, ผลกระทบ, ROI, ทีมงาน)`}
    >
      {[1, 2, 3, 4, 5].map((i) => (
        <Star key={i} size={size} fill={i <= n ? GOLD : "none"} style={{ color: i <= n ? GOLD : "#CBD2DC" }} />
      ))}
      {showPct && <span className="ml-1 text-[11px]" style={{ color: SUB }}>{pct}%</span>}
    </span>
  );
}

function AwardLevelCard({ k }) {
  const award = useContext(AwardCtx) || DEFAULT_AWARD;
  const { pct, n } = starLevel(k, award);
  const autoCrit = award.criteria.filter((c) => c.type === "auto");
  const winner = award.decision?.winners?.find((w) => w.id === k.id);
  return (
    <div className="rounded-xl border p-4" style={{ borderColor: "#E8D9B5", background: "#FFFBF2" }}>
      <div className="flex items-center justify-between gap-3 flex-wrap">
        <div className="flex items-center gap-2.5">
          <StarRating k={k} size={20} />
          <div>
            <div className="text-sm font-semibold" style={{ color: INK }}>ระดับตามเกณฑ์ Kaizen Award: {n} / 5 ดาว</div>
            <div className="text-[11px]" style={{ color: SUB }}>คะแนนประเมินเบื้องต้น {pct}% จากเกณฑ์ที่ระบบคำนวณให้อัตโนมัติ (ยังไม่รวมคะแนนของกรรมการ)</div>
          </div>
        </div>
        {winner && (
          <span className="inline-flex items-center gap-1 text-xs font-semibold px-2.5 py-1 rounded-full" style={{ background: "#FDF3E1", color: "#B07D1A", border: "1px solid #E8C77E" }}>
            <Trophy size={13} /> {winner.tier}
          </span>
        )}
      </div>
      <div className="mt-3 grid grid-cols-2 gap-x-6 gap-y-1.5">
        {autoCrit.map((c) => (
          <div key={c.id} className="flex items-center justify-between gap-2 text-xs">
            <span style={{ color: SUB }}>{c.name} <span style={{ color: "#A3ABB8" }}>({c.weight}%)</span></span>
            <span className="font-medium whitespace-nowrap" style={{ color: INK }}>
              {autoScore(c.source, k, award)}/5 <span style={{ color: "#A3ABB8" }}>· {autoReason(c.source, k)}</span>
            </span>
          </div>
        ))}
      </div>
      {!(k.categories || []).length && (
        <div className="text-[11px] mt-2" style={{ color: "#A3ABB8" }}>* ยังไม่ได้จัดหมวดหมู่ QCDSME คะแนนด้านผลกระทบอาจเพิ่มขึ้นหลัง Admin จัดหมวดหมู่</div>
      )}
    </div>
  );
}

// ---------- ตัวกรอง Multi Select (ใช้กับตารางของทุกบทบาท และ Dashboard) ----------
function MultiSelectFilter({ label, options, selected, onSelected, query, onQuery, dark, align = "left", allLabel, showSearch = true, queryIsFilter = true }) {
  const [open, setOpen] = useState(false);
  const [localQ, setLocalQ] = useState("");
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    const h = (e) => {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, [open]);
  const queryValue = query !== undefined ? query : localQ;
  const setQuery = (v) => (onQuery ? onQuery(v) : setLocalQ(v));
  const q = (queryValue || "").trim().toLowerCase();
  const shown = options.filter((o) => !q || o.label.toLowerCase().includes(q));
  const active = selected.length > 0 || (queryIsFilter && !!q);
  // allLabel: แสดง "ทุก..." เมื่อไม่ได้เลือก / ชื่อที่เลือกเมื่อเลือก 1 ค่า / "ชื่อ N รายการ" เมื่อเลือกหลายค่า
  const selLabels = options.filter((o) => selected.includes(o.value)).map((o) => o.label);
  const btnText = !allLabel ? label : selected.length === 0 ? allLabel : selected.length === 1 ? selLabels[0] : `${label} ${selected.length} รายการ`;
  const toggle = (v) => onSelected(selected.includes(v) ? selected.filter((x) => x !== v) : [...selected, v]);
  const allShown = shown.length > 0 && shown.every((o) => selected.includes(o.value));
  const btnStyle = dark
    ? { background: active ? "#FFFFFF" : "rgba(255,255,255,0.12)", color: active ? DB_DEEP : "white", borderColor: "rgba(255,255,255,0.35)" }
    : { background: active ? "#EAF3FC" : "white", color: active ? DB_BLUE : SUB, borderColor: active ? DB_BLUE : "#D8DDE5" };
  return (
    <div className="relative" ref={ref}>
      <button type="button" onClick={() => setOpen((o) => !o)} className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border text-xs font-medium whitespace-nowrap" style={btnStyle}>
        <span className={allLabel ? "max-w-[170px] truncate" : ""}>{btnText}</span>
        {!allLabel && selected.length > 0 && (
          <span className="min-w-[16px] h-4 px-1 rounded-full text-[10px] font-bold flex items-center justify-center text-white" style={{ background: DB_BLUE }}>{selected.length}</span>
        )}
        {queryIsFilter && q && selected.length === 0 && <Search size={11} />}
        <ChevronDown size={12} />
      </button>
      {open && (
        <div className="absolute z-40 mt-1 w-72 rounded-xl border bg-white shadow-xl p-2" style={{ borderColor: "#D8DDE5", [align]: 0 }}>
          {showSearch && (
            <input
              autoFocus
              value={queryValue || ""}
              onChange={(e) => setQuery(e.target.value)}
              placeholder={`ค้นหา ${label}...`}
              className="w-full rounded-lg border px-2.5 py-1.5 text-xs mb-1.5 outline-none"
              style={{ borderColor: "#D8DDE5", color: INK }}
            />
          )}
          <div className="flex items-center justify-between px-1 pb-1 text-[11px]" style={{ color: SUB }}>
            <button
              type="button"
              className="underline"
              onClick={() =>
                onSelected(allShown ? selected.filter((v) => !shown.some((o) => o.value === v)) : Array.from(new Set([...selected, ...shown.map((o) => o.value)])))
              }
            >
              {allShown ? "ยกเลิกที่แสดงอยู่ทั้งหมด" : "เลือกที่แสดงอยู่ทั้งหมด"}
            </button>
            <button type="button" className="underline" onClick={() => { onSelected([]); setQuery(""); }}>ล้าง</button>
          </div>
          <div className="max-h-56 overflow-y-auto">
            {shown.length === 0 && <div className="text-xs py-3 text-center" style={{ color: SUB }}>ไม่พบรายการ</div>}
            {shown.map((o) => (
              <label key={o.value} className="flex items-start gap-2 px-1.5 py-1 rounded hover:bg-slate-50 cursor-pointer text-xs" style={{ color: INK }}>
                <input type="checkbox" className="mt-0.5" checked={selected.includes(o.value)} onChange={() => toggle(o.value)} />
                <span>{o.label}</span>
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

const EMPTY_FILTERS = {
  no: { q: "", sel: [] }, title: { q: "", sel: [] }, dept: { q: "", sel: [] }, cat: { q: "", sel: [] }, status: { q: "", sel: [] },
};

function useKaizenFilters(rows) {
  const [f, setF] = useState(EMPTY_FILTERS);
  const part = (key) => ({
    selected: f[key].sel,
    query: f[key].q,
    onSelected: (sel) => setF((s) => ({ ...s, [key]: { ...s[key], sel } })),
    onQuery: (q) => setF((s) => ({ ...s, [key]: { ...s[key], q } })),
  });
  const textOk = (q, text) => !q.trim() || String(text).toLowerCase().includes(q.trim().toLowerCase());
  const filtered = useMemo(
    () =>
      rows.filter((k) => {
        const cats = k.categories || [];
        if (!textOk(f.no.q, `${k.registeredNo} ${k.date}`)) return false;
        if (f.no.sel.length && !f.no.sel.includes(k.registeredNo)) return false;
        if (!textOk(f.title.q, k.title)) return false;
        if (f.title.sel.length && !f.title.sel.includes(k.title)) return false;
        if (!textOk(f.dept.q, k.department)) return false;
        if (f.dept.sel.length && !f.dept.sel.includes(k.department)) return false;
        if (f.cat.sel.length && !f.cat.sel.some((c) => (c === "__none" ? cats.length === 0 : cats.includes(c)))) return false;
        if (!textOk(f.cat.q, cats.join(" "))) return false;
        if (f.status.sel.length && !f.status.sel.includes(k.status)) return false;
        return true;
      }),
    [rows, f]
  );
  const options = useMemo(
    () => ({
      no: [...rows].sort((a, b) => String(b.registeredNo).localeCompare(String(a.registeredNo))).map((k) => ({ value: k.registeredNo, label: `${k.registeredNo} · ${k.date}` })),
      title: Array.from(new Set(rows.map((k) => k.title))).sort().map((t) => ({ value: t, label: t })),
      dept: Array.from(new Set(rows.map((k) => k.department))).sort().map((d) => ({ value: d, label: d })),
      cat: [...CATEGORY_OPTIONS.map((c) => ({ value: c, label: c })), { value: "__none", label: "ยังไม่จัดหมวด" }],
      status: Object.entries(STATUS_META).map(([key, v]) => ({ value: key, label: v.label })),
    }),
    [rows]
  );
  const active = Object.values(f).some((p) => p.q.trim() || p.sel.length);
  return { part, filtered, options, active, clear: () => setF(EMPTY_FILTERS) };
}

function RowFilterBar({ flt, total, dark }) {
  const items = [
    ["no", "เลขที่ / วันที่", "left"], ["title", "เรื่อง", "left"], ["dept", "หน่วยงาน", "left"], ["cat", "หมวดหมู่", "right"], ["status", "สถานะ", "right"],
  ];
  return (
    <div className="flex flex-wrap items-center gap-2">
      <div className="flex items-center gap-1.5 text-xs font-medium pr-1" style={{ color: dark ? "rgba(255,255,255,0.9)" : SUB }}>
        <Search size={13} /> ค้นหา / กรอง
      </div>
      {items.map(([key, label, align]) => (
        <MultiSelectFilter key={key} label={label} options={flt.options[key]} dark={dark} align={align} {...flt.part(key)} />
      ))}
      {flt.active && (
        <button type="button" onClick={flt.clear} className="text-xs underline px-1" style={{ color: dark ? "white" : DB_BLUE }}>ล้างตัวกรอง</button>
      )}
      <span className="ml-auto text-[11px]" style={{ color: dark ? "rgba(255,255,255,0.7)" : SUB }}>
        แสดง {flt.filtered.length} จาก {total} รายการ
      </span>
    </div>
  );
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
const MAX_RAW_PHOTO_MB = 40;
const isImageFile = (f) => (f.type || "").startsWith("image/") || /\.(jpe?g|png|webp|gif|bmp|heic|heif)$/i.test(f.name || "");
// ย่อรูปก่อนอัปโหลด: ด้านยาวสุด 1280px เป็น JPEG ถ้ายังใหญ่ ย่อซ้ำให้เล็กลง (กันข้อมูลที่ส่งใหญ่เกินจนบันทึกไม่ได้)
// ถ้าอ่านรูปไม่ได้ จะแจ้ง error (ไม่ส่งไฟล์ดิบขึ้นไป)
async function compressImage(file, maxSide = 1280, quality = 0.8) {
  if (!isImageFile(file)) throw new Error("not-image");
  if (file.size > MAX_RAW_PHOTO_MB * 1024 * 1024) throw new Error("too-large");
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((resolve, reject) => {
      const el = new Image();
      el.onload = () => resolve(el);
      el.onerror = () => reject(new Error("decode"));
      el.src = url;
    });
    const w0 = img.naturalWidth || img.width;
    const h0 = img.naturalHeight || img.height;
    if (!w0 || !h0) throw new Error("decode");
    const render = (side, q) => {
      const scale = Math.min(1, side / Math.max(w0, h0));
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round(w0 * scale));
      canvas.height = Math.max(1, Math.round(h0 * scale));
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#FFFFFF";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      return canvas.toDataURL("image/jpeg", q);
    };
    let out = render(maxSide, quality);
    if (out.length > 450000) out = render(1024, 0.65);
    if (out.length > 450000) out = render(800, 0.6);
    if (out.length < 200) throw new Error("decode");
    return out;
  } finally {
    URL.revokeObjectURL(url);
  }
}
const describePhotoError = (file, err) => {
  const n = `"${file.name}"`;
  if (err && err.message === "too-large") return `${n} ใหญ่เกิน ${MAX_RAW_PHOTO_MB} MB`;
  if (err && err.message === "not-image") return `${n} ไม่ใช่ไฟล์รูปภาพ`;
  return `${n} อ่านรูปไม่ได้ (รองรับ JPG / PNG / WebP ไฟล์ HEIC จากบางเครื่องต้องแปลงเป็น JPG ก่อน หรือถ่ายรูปใหม่)`;
};
// รูปจาก Google Drive: ถ้าลิงก์แบบเดิมโหลดไม่ขึ้น ลองลิงก์สำรองให้อัตโนมัติ
function onDriveImgError(e) {
  const img = e.currentTarget;
  const m = String(img.src).match(/[?&]id=([\w-]+)|\/d\/([\w-]+)/);
  const id = m && (m[1] || m[2]);
  if (!id) return;
  const chain = [`https://lh3.googleusercontent.com/d/${id}=w1600`, `https://drive.google.com/thumbnail?id=${id}&sz=w1600`];
  const step = Number(img.dataset.fb || 0);
  if (step < chain.length) {
    img.dataset.fb = String(step + 1);
    img.src = chain[step];
  }
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


// ---------- รูปหลายภาพ (Before / After ได้สูงสุด 10 รูปต่อด้าน) ----------
function PhotoGallery({ label, photos, onChange, editable, onZoom, onBusyChange }) {
  const list = asList(photos);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  const [errors, setErrors] = useState([]);
  const pendingRef = useRef(0);
  const addFiles = async (fileList) => {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    const room = MAX_PHOTOS - list.length - pendingRef.current;
    const take = files.slice(0, Math.max(0, room));
    setNotice(files.length > room ? `เพิ่มได้สูงสุด ${MAX_PHOTOS} รูปต่อด้าน ระบบรับเพิ่มให้ ${take.length} รูป` : "");
    setErrors([]);
    if (!take.length) return;
    pendingRef.current += take.length;
    setBusy(true);
    if (onBusyChange) onBusyChange(true);
    const settled = await Promise.allSettled(take.map((f) => compressImage(f)));
    pendingRef.current -= take.length;
    const ok = settled.filter((r) => r.status === "fulfilled").map((r) => r.value);
    const bad = settled.map((r, i) => (r.status === "rejected" ? describePhotoError(take[i], r.reason) : null)).filter(Boolean);
    if (ok.length) onChange((prev) => [...prev, ...ok].slice(0, MAX_PHOTOS));
    setErrors(bad);
    if (pendingRef.current <= 0) {
      setBusy(false);
      if (onBusyChange) onBusyChange(false);
    }
  };
  return (
    <div>
      <div className="flex items-center justify-between mb-1.5">
        <div className="text-xs font-medium" style={{ color: SUB }}>{label}</div>
        <span className="text-[11px]" style={{ color: list.length >= MAX_PHOTOS ? RED : "#A3ABB8" }}>{list.length}/{MAX_PHOTOS} รูป</span>
      </div>
      {list.length === 0 && !editable ? (
        <div className="rounded-lg border h-36 flex flex-col items-center justify-center gap-1" style={{ borderColor: "#D8DDE5", background: "#FAFBFC", color: "#AEB6C2" }}>
          <ImageOff size={22} />
          <span className="text-[11px]">ไม่มีรูปภาพ</span>
        </div>
      ) : (
        <div className="grid grid-cols-3 sm:grid-cols-5 gap-2">
          {list.map((src, i) => (
            <div key={i} className="relative group rounded-lg overflow-hidden border aspect-[4/3]" style={{ borderColor: "#D8DDE5", background: "#FAFBFC" }}>
              <img
                src={src}
                alt={`${label} ${i + 1}`}
                className="w-full h-full object-cover"
                style={onZoom ? { cursor: "zoom-in" } : undefined}
                onClick={onZoom ? () => onZoom(list, i) : undefined}
                onError={onDriveImgError}
              />
              {editable && (
                <button
                  type="button"
                  onClick={() => onChange((prev) => prev.filter((_, j) => j !== i))}
                  title="ลบรูปนี้"
                  className="absolute top-1 right-1 bg-white shadow rounded-full p-0.5 border"
                  style={{ borderColor: "#D8DDE5" }}
                >
                  <X size={12} style={{ color: RED }} />
                </button>
              )}
            </div>
          ))}
          {editable && list.length < MAX_PHOTOS && (
            <label className="aspect-[4/3] rounded-lg border-2 border-dashed flex flex-col items-center justify-center gap-0.5 cursor-pointer hover:bg-slate-50" style={{ borderColor: "#C9D2DE", color: NAVY }}>
              <ImagePlus size={20} />
              <span className="text-[11px] font-medium">เพิ่มรูป</span>
              <input
                type="file"
                accept="image/*"
                multiple
                className="hidden"
                onChange={(e) => {
                  addFiles(e.target.files);
                  e.target.value = "";
                }}
              />
            </label>
          )}
        </div>
      )}
      {busy && <div className="text-[11px] mt-1" style={{ color: SUB }}>กำลังย่อขนาดรูป...</div>}
      {notice && <div className="text-[11px] mt-1" style={{ color: RED }}>{notice}</div>}
      {errors.map((er, i) => <div key={i} className="text-[11px] mt-1" style={{ color: RED }}>{er}</div>)}
    </div>
  );
}

// แสดงรูปแบบตารางสำหรับโปสเตอร์นำเสนอ (1 รูป = ใหญ่เต็มกรอบ หลายรูป = จัดเป็นตาราง)
function PhotoGrid({ label, photos, onZoom }) {
  const list = asList(photos);
  const n = list.length;
  const cols = n <= 1 ? 1 : n <= 4 ? 2 : n <= 6 ? 3 : 4;
  return (
    <div>
      <div className="text-xs font-medium mb-1.5" style={{ color: INK }}>{label}{n > 1 ? ` (${n} รูป)` : ""}</div>
      {n === 0 ? (
        <div className="rounded-lg border h-56 flex flex-col items-center justify-center gap-1" style={{ borderColor: "#D8DDE5", background: "#FAFBFC", color: "#AEB6C2" }}>
          <ImageOff size={22} />
          <span className="text-[11px]">ไม่มีรูปภาพ</span>
        </div>
      ) : (
        <div className="grid gap-1.5" style={{ gridTemplateColumns: `repeat(${cols}, minmax(0, 1fr))` }}>
          {list.map((src, i) => (
            <div key={i} className={`rounded-lg overflow-hidden border ${n === 1 ? "h-56" : "aspect-[4/3]"}`} style={{ borderColor: "#D8DDE5", background: "#FAFBFC" }}>
              <img
                src={src}
                alt={`${label} ${i + 1}`}
                className="w-full h-full object-cover"
                style={onZoom ? { cursor: "zoom-in" } : undefined}
                onClick={onZoom ? () => onZoom(list, i) : undefined}
                onError={onDriveImgError}
              />
            </div>
          ))}
        </div>
      )}
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
  if (asList(k.images?.before).length === 0) out.push("ภาพก่อนปรับปรุง");
  if (asList(k.images?.after).length === 0) out.push("ภาพหลังปรับปรุง");
  if (!String(k.problem || "").trim()) out.push("ที่มาของปัญหา");
  if (!String(k.improvement || "").trim()) out.push("แนวทางการปรับปรุง");
  if (!String(k.result || "").trim()) out.push("ผลที่ได้รับ");
  if (!(k.categories || []).length) out.push("หมวดหมู่ QCDSME");
  return out;
}

function KaizenForm({ initial, onCancel, onSave, currentUser, departments, adminMode, editorLabel = "Admin" }) {
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
      images: { before: [], after: [] },
      activityType: "Kaizen",
      attachments: [],
    }
  );
  const set = (k, v) => setF((s) => ({ ...s, [k]: v }));
  // อัปเดตรูปแบบ functional (กันรูปหายเมื่อแนบหลายด้านติดกัน) และรู้ว่ากำลังย่อรูปอยู่หรือไม่
  const setImages = (side, v) =>
    setF((x) => {
      const cur = asList(x.images?.[side]);
      return { ...x, images: { ...(x.images || {}), [side]: typeof v === "function" ? v(cur) : v } };
    });
  const [photoBusy, setPhotoBusy] = useState({ before: false, after: false });
  const anyPhotoBusy = photoBusy.before || photoBusy.after;
  const [adminNote, setAdminNote] = useState("");
  const missing = adminMode ? missingInfo(f) : [];

  return (
    <div className="space-y-4">
      {adminMode && (
        <div className="rounded-xl border p-3.5" style={{ background: "#FDF3E1", borderColor: "#E8C77E" }}>
          <div className="text-sm font-semibold" style={{ color: "#6B4A10" }}>แก้ไขข้อมูลในฐานะ {editorLabel} · {initial?.registeredNo}</div>
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
      {!adminMode && (
        <div>
          <Field label="ประเภทกิจกรรม" required>
            <select className={inputCls} style={inputStyle} value={f.activityType || "Kaizen"} onChange={(e) => set("activityType", e.target.value)}>
              {ACTIVITY_OPTIONS.map((o) => <option key={o}>{o}</option>)}
            </select>
          </Field>
          <ActivityHelp className="mt-1" />
        </div>
      )}
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
        <PhotoGallery label="ภาพก่อนปรับปรุง (Before)" photos={f.images?.before} editable onChange={(v) => setImages("before", v)} onBusyChange={(b) => setPhotoBusy((x) => ({ ...x, before: b }))} />
        <PhotoGallery label="ภาพหลังปรับปรุง (After)" photos={f.images?.after} editable onChange={(v) => setImages("after", v)} onBusyChange={(b) => setPhotoBusy((x) => ({ ...x, after: b }))} />
      </div>
      <AttachmentList
        attachments={f.attachments || []}
        editable
        onAdd={(file) => set("attachments", [...(f.attachments || []), file])}
        onRemove={(id) => set("attachments", (f.attachments || []).filter((a) => a.id !== id))}
      />
      {adminMode && (
        <div className="rounded-xl border p-4 space-y-3" style={{ borderColor: "#E3E7ED", background: "#FAFBFC" }}>
          <div className="text-sm font-semibold" style={{ color: INK }}>ข้อมูลจัดหมวดหมู่และประเภท</div>
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
          <ActivityHelp />
          <Field label="หมายเหตุการแก้ไข (บันทึกในประวัติ ไม่บังคับ)">
            <TextInput value={adminNote} onChange={(e) => setAdminNote(e.target.value)} placeholder="เช่น เพิ่มรูปก่อน-หลังตามที่หน่วยงานส่งมาทางไลน์" />
          </Field>
        </div>
      )}
      {anyPhotoBusy && <div className="text-xs text-right" style={{ color: SUB }}>กำลังย่อขนาดรูป รอสักครู่ ปุ่มบันทึกจะกดได้เมื่อเสร็จ...</div>}
      {adminMode ? (
        <div className="flex justify-end gap-2 pt-2">
          <button onClick={onCancel} className="px-4 py-2 rounded-lg text-sm font-medium border" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
          <button
            disabled={!f.title || anyPhotoBusy}
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
          disabled={anyPhotoBusy}
          onClick={() => onSave(f, "draft")}
          className="px-4 py-2 rounded-lg text-sm font-medium border disabled:opacity-40"
          style={{ borderColor: NAVY, color: NAVY }}
        >
          บันทึกแบบร่าง
        </button>
        <button
          disabled={!f.title || !f.problem || !f.improvement || anyPhotoBusy}
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

function ImageLightbox({ photos, index, onClose, onIndex }) {
  const list = photos || [];
  const n = list.length;
  useEffect(() => {
    if (index == null || !n) return;
    const h = (e) => {
      if (e.key === "Escape") onClose();
      if (e.key === "ArrowRight" && n > 1) onIndex((index + 1) % n);
      if (e.key === "ArrowLeft" && n > 1) onIndex((index - 1 + n) % n);
    };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  }, [index, n]);
  if (index == null || !n) return null;
  const nav = (dir) => (e) => {
    e.stopPropagation();
    onIndex((index + dir + n) % n);
  };
  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center p-6" style={{ background: "rgba(0,0,0,0.85)" }} onClick={onClose}>
      <button onClick={onClose} className="absolute top-4 right-4 p-2 rounded-full" style={{ background: "rgba(255,255,255,0.15)" }}>
        <X size={22} className="text-white" />
      </button>
      {n > 1 && (
        <>
          <button onClick={nav(-1)} className="absolute left-4 top-1/2 -translate-y-1/2 p-2 rounded-full" style={{ background: "rgba(255,255,255,0.15)" }} title="รูปก่อนหน้า">
            <ChevronLeft size={26} className="text-white" />
          </button>
          <button onClick={nav(1)} className="absolute right-4 top-1/2 -translate-y-1/2 p-2 rounded-full" style={{ background: "rgba(255,255,255,0.15)" }} title="รูปถัดไป">
            <ChevronRight size={26} className="text-white" />
          </button>
          <div className="absolute bottom-5 left-1/2 -translate-x-1/2 text-white text-xs px-3 py-1 rounded-full" style={{ background: "rgba(255,255,255,0.18)" }}>
            {index + 1} / {n}
          </div>
        </>
      )}
      <img src={list[index]} alt="ภาพขยาย" className="max-w-full max-h-full rounded-lg shadow-2xl object-contain" onClick={(e) => e.stopPropagation()} onError={onDriveImgError} />
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
          h.action === "edit" ? `แก้ไขข้อมูลโดย${{ admin: " Admin", supervisor: " หัวหน้างาน", director: " ผู้อำนวยการ" }[USERS.find((u) => u.name === h.actor)?.role] || ""}` : "";
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
  const [lb, setLb] = useState(null);
  return (
    <div className="space-y-5">
      <ImageLightbox photos={lb?.photos} index={lb?.index} onClose={() => setLb(null)} onIndex={(i) => setLb((x) => ({ ...x, index: i }))} />
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
      <AwardLevelCard k={k} />
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
        <PhotoGallery label="ก่อนปรับปรุง (Before)" photos={k.images?.before} onZoom={(photos, index) => setLb({ photos, index })} />
        <PhotoGallery label="หลังปรับปรุง (After)" photos={k.images?.after} onZoom={(photos, index) => setLb({ photos, index })} />
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

// ---------- บันทึกโปสเตอร์ไปใช้งานต่อ ----------
function loadExternalScript(src) {
  return new Promise((resolve, reject) => {
    if (window.html2canvas) return resolve();
    const el = document.createElement("script");
    el.src = src;
    el.onload = resolve;
    el.onerror = () => reject(new Error("โหลดตัวช่วยสร้างรูปภาพไม่สำเร็จ"));
    document.head.appendChild(el);
  });
}

// รูปจาก Google Drive ต้องใช้ลิงก์ที่อนุญาตให้นำไปวาดเป็นไฟล์รูปภาพได้
function driveImgUrl(src) {
  const m = String(src).match(/drive\.google\.com\/(?:uc\?[^#]*?id=|file\/d\/)([\w-]+)/);
  return m ? `https://lh3.googleusercontent.com/d/${m[1]}=w1600` : src;
}

async function downloadPosterPng(node, filename) {
  await loadExternalScript("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js");
  const canvas = await window.html2canvas(node, {
    scale: 2,
    useCORS: true,
    backgroundColor: "#ffffff",
    onclone: (doc) => {
      doc.querySelectorAll("img").forEach((img) => {
        img.crossOrigin = "anonymous";
        img.src = driveImgUrl(img.src);
      });
    },
  });
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  if (!blob) throw new Error("สร้างไฟล์รูปภาพไม่สำเร็จ");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

// พิมพ์เฉพาะโปสเตอร์ (A4 แนวนอน) เลือก "บันทึกเป็น PDF" ในหน้าต่างพิมพ์เพื่อเก็บเป็นไฟล์
function printPoster(node, title) {
  const styles = Array.from(document.querySelectorAll('link[rel="stylesheet"], style'))
    .map((el) => (el.tagName === "LINK" ? `<link rel="stylesheet" href="${el.href}">` : el.outerHTML))
    .join("");
  const iframe = document.createElement("iframe");
  iframe.setAttribute("data-poster-print", "1");
  iframe.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0;";
  document.body.appendChild(iframe);
  const doc = iframe.contentDocument;
  doc.open();
  doc.write(
    `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title>${styles}` +
      `<style>@page{size:A4 landscape;margin:8mm}html,body{margin:0;background:#fff}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}</style>` +
      `</head><body>${node.outerHTML}</body></html>`
  );
  doc.close();
  const waits = Array.from(doc.images).map((img) => (img.complete ? null : new Promise((r) => { img.onload = img.onerror = r; })));
  Promise.all(waits).then(() =>
    setTimeout(() => {
      try {
        iframe.contentWindow.focus();
        iframe.contentWindow.print();
      } finally {
        setTimeout(() => iframe.remove(), 4000);
      }
    }, 500)
  );
}

function PresentationView({ k }) {
  const [lb, setLb] = useState(null);
  const posterRef = useRef(null);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState("");
  const fileBase = `Kaizen_${k.registeredNo || "poster"}`;
  const savePng = async () => {
    setMsg("");
    setSaving(true);
    try {
      await downloadPosterPng(posterRef.current, `${fileBase}.png`);
    } catch (e) {
      setMsg("ดาวน์โหลดเป็นรูปภาพไม่สำเร็จ (อาจถูกบล็อกโดยเครือข่าย หรือรูปจาก Google Drive ไม่อนุญาตให้ดึง) ลองใช้ปุ่ม \"พิมพ์ / บันทึกเป็น PDF\" แทนได้ค่ะ");
    }
    setSaving(false);
  };
  return (
    <div>
      <div className="flex items-center gap-2 mb-3 flex-wrap">
        <span className="text-xs mr-auto" style={{ color: SUB }}>นำโปสเตอร์ไปใช้งานต่อ:</span>
        <button onClick={savePng} disabled={saving} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-white disabled:opacity-60" style={{ background: NAVY }}>
          <Download size={14} /> {saving ? "กำลังสร้างรูปภาพ..." : "ดาวน์โหลดเป็นรูปภาพ (PNG)"}
        </button>
        <button onClick={() => printPoster(posterRef.current, fileBase)} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border" style={{ borderColor: NAVY, color: NAVY }}>
          <Printer size={14} /> พิมพ์ / บันทึกเป็น PDF
        </button>
      </div>
      {msg && <div className="text-xs mb-3 rounded-lg px-3 py-2" style={{ background: "#FBEAE8", color: "#8B2E24" }}>{msg}</div>}
    <div ref={posterRef} className="rounded-xl overflow-hidden border" style={{ borderColor: "#D8DDE5" }}>
      <ImageLightbox photos={lb?.photos} index={lb?.index} onClose={() => setLb(null)} onIndex={(i) => setLb((x) => ({ ...x, index: i }))} />

      {/* Header */}
      <div className="px-5 py-3.5 flex items-center justify-between flex-wrap gap-2" style={{ background: NAVY }}>
        <div className="text-white text-xl font-extrabold flex items-center gap-2">
          <span>ไคเซ็น <span className="font-semibold opacity-90">(KAIZEN)</span></span>
          {k.attachments && k.attachments.length > 0 && (
            <span title={`มีไฟล์แนบ ${k.attachments.length} ไฟล์`} className="inline-flex items-center flex-shrink-0 text-white/70">
              <Paperclip size={17} />
            </span>
          )}
          <StarRating k={k} size={15} className="ml-1" />
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
              <PhotoGrid label="ก่อนปรับปรุง" photos={k.images?.before} onZoom={(photos, index) => setLb({ photos, index })} />
              <ChevronRight size={26} style={{ color: "#AEB6C2" }} />
              <PhotoGrid label="หลังปรับปรุง" photos={k.images?.after} onZoom={(photos, index) => setLb({ photos, index })} />
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
    </div>
  );
}

function Table({ rows, actions, showMissing, filterable = true }) {
  const flt = useKaizenFilters(rows);
  const shown = filterable ? flt.filtered : rows;
  return (
    <div>
      {filterable && rows.length > 0 && (
        <div className="mb-2">
          <RowFilterBar flt={flt} total={rows.length} />
        </div>
      )}
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
            {shown.length === 0 && (
              <tr><td colSpan={6} className="text-center py-8 text-sm" style={{ color: SUB }}>{rows.length ? "ไม่พบรายการตามตัวกรอง" : "ไม่มีรายการ"}</td></tr>
            )}
            {shown.map((k) => (
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
                  <div className="mt-1 flex items-center gap-2 flex-wrap">
                    <StarRating k={k} size={12} />
                    {(k.activityType || "Kaizen") === "Suggestion" && (
                      <span className="text-[10px] px-1.5 py-0.5 rounded-full font-normal" style={{ background: "#F1F3F6", color: "#8A93A3" }}>Suggestion</span>
                    )}
                  </div>
                  {showMissing && (asList(k.images?.before).length === 0 || asList(k.images?.after).length === 0) && (
                    <span className="inline-flex items-center gap-1 mt-1 text-[10px] px-1.5 py-0.5 rounded-full font-normal" style={{ background: "#FBEAE8", color: "#8B2E24" }}>
                      <ImageOff size={10} /> {asList(k.images?.before).length === 0 && asList(k.images?.after).length === 0 ? "ยังไม่มีรูปก่อน-หลัง" : asList(k.images?.before).length === 0 ? "ยังไม่มีรูปก่อน" : "ยังไม่มีรูปหลัง"}
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
              <ActivityHelp />
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

function DepartmentManager({ departments, ops }) {
  const [panel, setPanel] = useState(null); // { type: "add" | "edit" | "delete" | "sec-add" | "sec-edit" | "sec-delete", name?, section? }
  const [note, setNote] = useState(null);
  const groups = groupBySection(departments);
  const done = (r) => { setPanel(null); setNote(r.saved); };
  return (
    <div className="rounded-xl border p-4 space-y-3" style={{ borderColor: "#E2E6EB" }}>
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <div className="font-medium text-sm flex items-center gap-1.5" style={{ color: INK }}>
          <Building2 size={15} /> จัดการฝ่ายและหน่วยงาน <span className="text-xs font-normal" style={{ color: SUB }}>({SECTION_OPTIONS.length} ฝ่าย · {departments.length} หน่วยงาน)</span>
        </div>
        <div className="flex items-center gap-1.5">
          <button type="button" onClick={() => setPanel({ type: "sec-add" })} className="text-xs font-medium px-3 py-1.5 rounded-lg border flex items-center gap-1" style={{ borderColor: NAVY, color: NAVY }}>
            <Plus size={13} /> เพิ่มฝ่ายใหม่
          </button>
          <button type="button" onClick={() => setPanel({ type: "add" })} className="text-xs font-medium px-3 py-1.5 rounded-lg border flex items-center gap-1" style={{ borderColor: NAVY, color: NAVY }}>
            <Plus size={13} /> เพิ่มหน่วยงานใหม่
          </button>
        </div>
      </div>
      {panel && panel.type === "sec-add" && <SectionAddForm ops={ops} onDone={done} onCancel={() => setPanel(null)} />}
      {panel && panel.type === "add" && <DeptAddForm key={panel.section || "x"} defaultSection={panel.section} ops={ops} onDone={done} onCancel={() => setPanel(null)} />}
      <div className="space-y-3">
        {groups.map(([sec, list]) => (
          <div key={sec}>
            <div className="flex items-center gap-2 mb-1">
              <span className="text-xs font-semibold" style={{ color: NAVY }}>{sec}</span>
              <span className="text-[11px]" style={{ color: "#A3ABB8" }}>{list.length} หน่วยงาน</span>
              {sec !== "อื่นๆ" && (
                <>
                  <button type="button" onClick={() => setPanel({ type: "add", section: sec })} className="text-[11px] px-1.5 py-0.5 rounded border" style={{ borderColor: "#D8DDE5", color: SUB }}>+ หน่วยงาน</button>
                  <TinyIcon icon={Pencil} label={`แก้ไขฝ่าย ${sec}`} color="#6A4BC9" onClick={() => setPanel({ type: "sec-edit", name: sec })} />
                  <TinyIcon icon={Trash2} label={`ลบฝ่าย ${sec}`} color={RED} onClick={() => setPanel({ type: "sec-delete", name: sec })} />
                </>
              )}
            </div>
            {panel && panel.type === "sec-edit" && panel.name === sec && <div className="mb-1 pl-3"><SectionEditForm key={sec} name={sec} ops={ops} onDone={done} onCancel={() => setPanel(null)} /></div>}
            {panel && panel.type === "sec-delete" && panel.name === sec && <div className="mb-1 pl-3"><SectionDeleteConfirm key={sec} name={sec} ops={ops} onDone={done} onCancel={() => setPanel(null)} /></div>}
            {list.length === 0 && <div className="text-[11px] pl-3" style={{ color: "#A3ABB8" }}>ยังไม่มีหน่วยงานในฝ่ายนี้</div>}
            <div className="space-y-1">
              {list.map((d) => {
                const st = ops.stats(d);
                return (
                  <div key={d}>
                    <div className="flex items-center gap-2 text-sm pl-3">
                      <span className="flex-1" style={{ color: INK }}>{d}</span>
                      <span className="text-[11px]" style={{ color: "#A3ABB8" }}>Kaizen {st.kaizens} · บัญชี {st.users.length}</span>
                      <TinyIcon icon={Pencil} label={`แก้ไขหน่วยงาน ${d}`} color="#6A4BC9" onClick={() => setPanel({ type: "edit", name: d })} />
                      <TinyIcon icon={Trash2} label={`ลบหน่วยงาน ${d}`} color={RED} onClick={() => setPanel({ type: "delete", name: d })} />
                    </div>
                    {panel && panel.type === "edit" && panel.name === d && <div className="mt-1 pl-3"><DeptEditForm key={d} name={d} ops={ops} onDone={done} onCancel={() => setPanel(null)} /></div>}
                    {panel && panel.type === "delete" && panel.name === d && <div className="mt-1 pl-3"><DeptDeleteConfirm key={d} name={d} ops={ops} onDone={done} onCancel={() => setPanel(null)} /></div>}
                  </div>
                );
              })}
            </div>
          </div>
        ))}
      </div>
      <DeptResultNote key={String(note)} saved={note} ops={ops} onClear={() => setNote(null)} />
      <div className="text-[11px]" style={{ color: "#8A93A3" }}>
        * ฝ่าย (Section) ใช้จัดกลุ่มหน่วยงานในแดชบอร์ด · เปลี่ยนชื่อหน่วยงาน: ระบบแก้ในบัญชีและ Kaizen เดิมให้ทั้งหมด · ลบหน่วยงานได้เฉพาะที่ไม่มี Kaizen และไม่ทำให้บัญชีใดไม่เหลือหน่วยงาน · ลบฝ่ายได้เฉพาะฝ่ายที่ไม่มีหน่วยงาน
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
  const flt = useKaizenFilters(rows);
  const filteredRows = flt.filtered;
  const thBase = { color: "#EAF3FC", fontWeight: 600 };

  return (
    <div>
      <div className="rounded-xl p-3 mb-2" style={{ background: DB_DEEP }}>
        <RowFilterBar flt={flt} total={rows.length} dark />
      </div>
      <div className="rounded-xl border overflow-hidden" style={{ borderColor: DB_BORDER }}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr style={{ background: DB_DEEP }}>
                <th className="text-left px-4 py-2.5" style={thBase}>เลขที่ / วันที่</th>
                <th className="text-left px-4 py-2.5" style={thBase}>เรื่อง</th>
                <th className="text-left px-4 py-2.5" style={thBase}>หน่วยงาน</th>
                <th className="text-left px-4 py-2.5" style={thBase}>หมวดหมู่</th>
                <th className="text-left px-4 py-2.5" style={thBase}>สถานะ</th>
                <th className="text-right px-4 py-2.5" style={thBase}>จัดการ</th>
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
                    <div className="mt-1 flex items-center gap-2 flex-wrap">
                      <StarRating k={k} size={12} />
                      {(k.activityType || "Kaizen") === "Suggestion" && (
                        <span className="text-[10px] px-1.5 py-0.5 rounded-full font-normal" style={{ background: "#F1F3F6", color: "#8A93A3" }}>Suggestion</span>
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
    </div>
  );
}

function Dashboard({ data, departments, onView, onPresent }) {
  // ตัวกรองทุกตัวเลือกได้หลายค่า (ว่าง = ทุกค่า) · เลือกหลายค่าในตัวกรองเดียวกัน = "หรือ" · ต่างตัวกรอง = "และ"
  const [yearFilter, setYearFilter] = useState([]);
  const [monthFilter, setMonthFilter] = useState([]);
  const [sectionFilter, setSectionFilter] = useState([]);
  const [deptFilter, setDeptFilter] = useState([]);
  const [catFilter, setCatFilter] = useState([]);
  const [activityFilter, setActivityFilter] = useState([]);
  const [kaizenTypeFilter, setKaizenTypeFilter] = useState([]);
  const anyDashFilter = [yearFilter, monthFilter, sectionFilter, deptFilter, catFilter, activityFilter, kaizenTypeFilter].some((f) => f.length > 0);
  const clearDashFilters = () => {
    setYearFilter([]); setMonthFilter([]); setSectionFilter([]); setDeptFilter([]); setCatFilter([]); setActivityFilter([]); setKaizenTypeFilter([]);
  };

  const yearOptions = useMemo(() => Array.from(new Set(data.map((k) => k.date.slice(0, 4)))).sort().reverse(), [data]);

  // เมื่อเลือกฝ่ายแล้ว รายการหน่วยงานใน Unit จะกรองให้เหลือเฉพาะหน่วยงานในฝ่ายนั้น
  const unitOptions = useMemo(() => {
    const all = departments && departments.length ? departments : DEFAULT_DEPARTMENTS;
    if (sectionFilter.length === 0) return all;
    return all.filter((d) => sectionFilter.includes(SECTION_MAP[d]));
  }, [departments, sectionFilter]);

  // เปลี่ยนฝ่ายแล้ว หน่วยงานที่เลือกไว้แต่ไม่อยู่ในฝ่ายที่เลือก จะถูกเอาออกให้อัตโนมัติ
  const handleSectionChange = (vals) => {
    setSectionFilter(vals);
    if (vals.length > 0) setDeptFilter((prev) => prev.filter((d) => vals.includes(SECTION_MAP[d])));
  };

  const filtered = data.filter((k) => {
    const y = k.date.slice(0, 4);
    const m = k.date.slice(5, 7);
    const has = (list, v) => list.length === 0 || list.includes(v);
    if (!has(yearFilter, y)) return false;
    if (!has(monthFilter, m)) return false;
    if (!has(sectionFilter, SECTION_MAP[k.department])) return false;
    if (!has(deptFilter, k.department)) return false;
    const cats = k.categories || [];
    if (catFilter.length > 0 && !catFilter.some((c) => (c === "__none" ? cats.length === 0 : cats.includes(c)))) return false;
    if (!has(activityFilter, k.activityType)) return false;
    if (!has(kaizenTypeFilter, k.kaizenType)) return false;
    return true;
  });
  // ยอดการส่ง Kaizen = หัวหน้างานอนุมัติแล้ว และเป็นประเภทกิจกรรม Kaizen เท่านั้น (ไม่นับ Suggestion)
  const counted = useMemo(() => filtered.filter(isCountedKaizen), [filtered]);

  const costSavingBySection = useMemo(() => {
    const m = {};
    counted.forEach((k) => {
      const v = parseFloat(k.costSaving) || 0;
      const section = SECTION_MAP[k.department] || "อื่นๆ";
      if (v > 0) m[section] = (m[section] || 0) + v;
    });
    return Object.entries(m).sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value }));
  }, [counted]);

  const countByDept = useMemo(() => {
    const m = {};
    counted.forEach((k) => (m[k.department] = (m[k.department] || 0) + 1));
    return Object.entries(m).sort((a, b) => b[1] - a[1]).map(([name, count]) => ({ name, count }));
  }, [counted]);

  const countBySection = useMemo(() => {
    const m = {};
    counted.forEach((k) => {
      const section = SECTION_MAP[k.department] || "อื่นๆ";
      m[section] = (m[section] || 0) + 1;
    });
    return Object.entries(m).sort((a, b) => b[1] - a[1]).map(([name, count]) => ({ name, count }));
  }, [counted]);

  const byMonth = useMemo(() => {
    const m = {};
    counted.forEach((k) => {
      const key = k.date.slice(0, 7);
      m[key] = (m[key] || 0) + 1;
    });
    return Object.entries(m).sort().map(([key, count]) => {
      const [y, mo] = key.split("-");
      return { name: `${THAI_MONTHS[parseInt(mo, 10) - 1]} ${y.slice(2)}`, count };
    });
  }, [counted]);

  const total = counted.length;
  const kaizenProjectCount = counted.filter((k) => k.kaizenType === "Kaizen Project").length;
  const crossFunctionalCount = counted.filter((k) => k.kaizenType === "Cross functional Kaizen").length;
  const totalCostSaving = counted.reduce((sum, k) => sum + (parseFloat(k.costSaving) || 0), 0);
  const excludedCount = filtered.length - counted.length;

  // จำนวนตามประเภทข้อมูล (หมวด QCDSME) — 1 เรื่องนับได้หลายหมวด
  const countByCategory = useMemo(() => {
    const m = Object.fromEntries(CATEGORY_OPTIONS.map((c) => [c, 0]));
    let none = 0;
    counted.forEach((k) => {
      if (!(k.categories || []).length) none += 1;
      (k.categories || []).forEach((c) => { m[c] = (m[c] || 0) + 1; });
    });
    const rows = CATEGORY_OPTIONS.map((c) => ({ name: c, count: m[c] }));
    if (none) rows.push({ name: "ยังไม่จัดหมวด", count: none });
    return rows;
  }, [counted]);

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
          <MultiSelectFilter dark label="ปี" allLabel="ทุกปี" showSearch={false} queryIsFilter={false} options={yearOptions.map((y) => ({ value: y, label: y }))} selected={yearFilter} onSelected={setYearFilter} />
          <MultiSelectFilter dark label="เดือน" allLabel="ทุกเดือน" showSearch={false} queryIsFilter={false} options={THAI_MONTHS.map((m, i) => ({ value: String(i + 1).padStart(2, "0"), label: m }))} selected={monthFilter} onSelected={setMonthFilter} />
          <MultiSelectFilter dark label="ฝ่าย" allLabel="ทุกฝ่าย (Section)" showSearch={false} queryIsFilter={false} options={SECTION_OPTIONS.map((x) => ({ value: x, label: x }))} selected={sectionFilter} onSelected={handleSectionChange} />
          <MultiSelectFilter dark label="หน่วยงาน" allLabel="ทุกหน่วยงาน (Unit)" queryIsFilter={false} options={unitOptions.map((d) => ({ value: d, label: d }))} selected={deptFilter} onSelected={setDeptFilter} />
          <MultiSelectFilter dark align="right" label="ประเภทข้อมูล" allLabel="ทุกประเภทข้อมูล" showSearch={false} queryIsFilter={false} options={[...CATEGORY_OPTIONS.map((c) => ({ value: c, label: c })), { value: "__none", label: "ยังไม่จัดหมวด" }]} selected={catFilter} onSelected={setCatFilter} />
          <MultiSelectFilter dark align="right" label="ประเภทกิจกรรม" allLabel="ทุกประเภทกิจกรรม" showSearch={false} queryIsFilter={false} options={ACTIVITY_OPTIONS.map((o) => ({ value: o, label: o }))} selected={activityFilter} onSelected={setActivityFilter} />
          <MultiSelectFilter dark align="right" label="ประเภท Kaizen" allLabel="ทุกประเภท Kaizen" showSearch={false} queryIsFilter={false} options={KAIZEN_TYPE_OPTIONS.map((o) => ({ value: o, label: o }))} selected={kaizenTypeFilter} onSelected={setKaizenTypeFilter} />
          {anyDashFilter && <button type="button" onClick={clearDashFilters} className="text-xs underline text-white px-1">ล้างตัวกรอง</button>}
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
        <ActivityHelp light className="mt-2.5" />
      </div>

      <div id="dashboard-print-area" className="space-y-5">
      <div className="text-xs rounded-lg px-3 py-2" style={{ background: DB_PALE, color: "#4C6A8C" }}>
        ยอดและกราฟด้านล่างนับเฉพาะ <strong>Kaizen ที่หัวหน้างานอนุมัติแล้ว</strong> และเป็นประเภทกิจกรรม <strong>Kaizen</strong> (ไม่นับ Suggestion และเรื่องที่รอหัวหน้างานอนุมัติ)
        {excludedCount > 0 && <> · ตามตัวกรองนี้มี {excludedCount} เรื่องที่ยังไม่ถูกนับ (แสดงในตารางด้านล่าง)</>}
      </div>
      <div className="grid grid-cols-4 gap-4">
        <div className="rounded-xl p-4 text-white" style={{ background: `linear-gradient(135deg, ${DB_DEEP}, ${DB_BLUE})` }}>
          <div className="text-xs font-medium text-white/80">Kaizen รวมทั้งหมด</div>
          <div className="text-3xl font-semibold mt-1.5">{total}</div>
          <div className="text-xs mt-1 text-white/70">เฉพาะที่หัวหน้างานอนุมัติแล้ว</div>
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
          <div className="text-xs mt-1 text-white/85 font-medium relative">บาท/ปี จาก Kaizen ที่นับ</div>
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
        <div className="text-sm font-medium" style={{ color: DB_DEEP }}>จำนวน Kaizen ตามประเภทข้อมูล (Cost, Safety, Quality และอื่นๆ)</div>
        <div className="text-[11px] mb-3" style={{ color: "#8A9AB0" }}>1 เรื่องนับได้หลายประเภท ถ้า Admin จัดไว้หลายหมวด · "ยังไม่จัดหมวด" คือเรื่องที่ Admin ยังไม่ได้ตรวจและจัดหมวดหมู่</div>
        <ResponsiveContainer width="100%" height={250}>
          <BarChart data={countByCategory} margin={{ top: 18 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#EEF3FA" />
            <XAxis dataKey="name" tick={{ fontSize: 11 }} interval={0} />
            <YAxis allowDecimals={false} tick={{ fontSize: 11 }} />
            <Tooltip formatter={(v) => `${v} เรื่อง`} />
            <Bar dataKey="count" radius={[4, 4, 0, 0]}>
              {countByCategory.map((row, i) => (
                <Cell key={row.name} fill={row.name === "ยังไม่จัดหมวด" ? "#B8C2D0" : ["#1D5FC9", "#3F8FE0", "#2F7D5C", "#D99A2B", "#B23A2E", "#7A5AE0", "#D9577E", "#2BA3B8", "#8A6D3B", "#5B6675"][i % 10]} />
              ))}
              <LabelList dataKey="count" position="top" style={{ fontSize: 11, fill: DB_DEEP, fontWeight: 600 }} />
            </Bar>
          </BarChart>
        </ResponsiveContainer>
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

// บัญชีกรรมการ (ชื่อผู้ใช้ตรงกับ USERS ที่ role = "judge")
const MAX_JUDGES = 7;
const JUDGE_SLOTS = Array.from({ length: MAX_JUDGES }, (_, i) => i + 1).map((n) => ({ id: `j${n}`, userName: `กรรมการ ${n}` }));
const AUTO_LABELS = {
  cost: "มูลค่า Cost Saving", roi: "ความคุ้มค่าการลงทุน (ROI)", impact: "ผลกระทบ (หมวดผลกระทบสูง)", team: "การทำงานเป็นทีม / ข้ามสายงาน",
};
const OLD_IMPACT_NAME = "ผลกระทบด้าน Safety / Quality / Food Safety";
const NEW_IMPACT_NAME = "ผลกระทบสูง (Safety / Quality / Food Safety / Cost)";

const DEFAULT_AWARD = {
  roundName: "Kaizen Award 2026",
  year: "all",
  statuses: ["approved"],
  activity: "Kaizen",
  minCostSaving: 0,
  mainPrizes: 3,
  honorable: 2,
  costTiers: [100000, 50000, 20000, 5000],
  highImpactCats: ["Safety", "Food Safety", "Quality", "Cost"],
  criteria: [
    { id: "cost", name: "มูลค่า Cost Saving", weight: 30, type: "auto", source: "cost" },
    { id: "impact", name: NEW_IMPACT_NAME, weight: 25, type: "auto", source: "impact" },
    { id: "innov", name: "ความคิดสร้างสรรค์ / ความแปลกใหม่", weight: 15, type: "manual" },
    { id: "roi", name: "ความคุ้มค่าการลงทุน (ROI)", weight: 10, type: "auto", source: "roi" },
    { id: "team", name: "การทำงานเป็นทีม / ข้ามสายงาน", weight: 10, type: "auto", source: "team" },
    { id: "yokoten", name: "ขยายผลไปหน่วยงานอื่นได้ (Yokoten)", weight: 10, type: "manual" },
  ],
  cfgVersion: 2, // 2 = หมวดผลกระทบสูงรวม Cost
  starThresholds: [80, 65, 50, 35], // คะแนนคัดกรอง (%) ขั้นต่ำของ 5 / 4 / 3 / 2 ดาว (ต่ำกว่านั้น = 1 ดาว)
  judgeCount: 5,
  judgeNames: {}, // { j1: "ชื่อจริงของกรรมการ (ไม่บังคับ)" }
  trimExtremes: true,
  shortlist: [], // id ของ Kaizen ที่ Admin ส่งให้กรรมการ
  sentAt: null,
  sentBy: null,
  scores: {}, // { j1: { [kaizenId]: { [criterionId]: 1..5, __note: "" }, __meta: { submittedAt } } }
  decision: null,
};

const EMPTY_DRAFT = {};
const AWARD_SETTING_LABELS = {
  roundName: "ชื่อรอบประกวด", year: "ผู้มีสิทธิ์เข้าประกวด", activity: "ผู้มีสิทธิ์เข้าประกวด", minCostSaving: "ผู้มีสิทธิ์เข้าประกวด",
  statuses: "ผู้มีสิทธิ์เข้าประกวด", criteria: "เกณฑ์ / น้ำหนัก", costTiers: "ช่วงคะแนน Cost Saving", highImpactCats: "หมวดผลกระทบสูง",
  mainPrizes: "จำนวนรางวัล", honorable: "จำนวนรางวัล", judgeCount: "จำนวนกรรมการ", judgeNames: "ชื่อกรรมการ",
  trimExtremes: "วิธีเฉลี่ยคะแนน", starThresholds: "ระดับดาว",
};

function normalizeAward(a) {
  const base = { ...DEFAULT_AWARD, ...(a || {}) };
  delete base.judges; // รุ่นก่อนหน้า (Admin/ผอ. เป็นกรรมการ) ไม่ใช้แล้ว
  base.judgeCount = Math.min(MAX_JUDGES, Math.max(1, toNum(base.judgeCount) || 5));
  // ข้อมูลที่บันทึกไว้ก่อนหน้า (ยังไม่มี cfgVersion 2): เพิ่ม Cost เข้าหมวดผลกระทบสูงให้ 1 ครั้ง (หลังจากนั้นปรับเองได้ตามต้องการ)
  if (a && (a.cfgVersion || 1) < 2) {
    base.highImpactCats = Array.from(new Set([...(base.highImpactCats || []), "Cost"]));
    base.criteria = (base.criteria || DEFAULT_AWARD.criteria).map((c) => (c.id === "impact" && c.name === OLD_IMPACT_NAME ? { ...c, name: NEW_IMPACT_NAME } : c));
  }
  base.cfgVersion = 2;
  base.judgeNames = base.judgeNames || {};
  base.starThresholds = Array.isArray(base.starThresholds) && base.starThresholds.length === 4 ? base.starThresholds.map((v) => toNum(v)) : DEFAULT_AWARD.starThresholds;
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
function AwardPage({ kaizens, award, onChange, onScore, currentUser, onView, onPresent, onAnnounce, onSendToJudges, onJudgesNotice, draft, onDraftChange, onSaveSettings, onRefresh, syncMode }) {
  const draftSafe = draft || EMPTY_DRAFT;
  // cfg = ค่าที่บันทึกแล้ว + ร่างที่กำลังแก้ (ผู้ใช้อื่น/กรรมการ/ดาว ยังเห็นเฉพาะค่าที่บันทึกแล้ว)
  const cfg = useMemo(() => ({ ...award, ...draftSafe }), [award, draftSafe]);
  const isDirector = currentUser.role === "director";
  const [sub, setSub] = useStoredState("kz_award_sub", "criteria");
  const [confirming, setConfirming] = useState(false);
  const [selected, setSelected] = useState(() => new Set(cfg.shortlist));
  const [topN, setTopN] = useState(5);
  const [unlocking, setUnlocking] = useState(null);
  const [judgeFilter, setJudgeFilter] = useState("all");
  const [clearing, setClearing] = useState(false);
  const [clearScoresToo, setClearScoresToo] = useState(true);

  const locked = !!cfg.decision;
  const [saveMsg, setSaveMsg] = useState(null); // { type: "ok" | "error", text }
  const [saving, setSaving] = useState(false);
  const [confirmCancel, setConfirmCancel] = useState(false);
  const [adding, setAdding] = useState(false);
  const [newCrit, setNewCrit] = useState({ name: "", desc: "", weight: 10, type: "manual" });
  // set = แก้การตั้งค่า (เก็บเป็นร่างจนกว่าจะกดบันทึก) · commit = การกระทำที่มีผลทันที (ส่งกรรมการ / ประกาศผล)
  const set = (patch) => {
    setSaveMsg(null);
    setConfirmCancel(false);
    onDraftChange((d) => {
      const merged = { ...d, ...patch };
      Object.keys(merged).forEach((key) => {
        if (JSON.stringify(merged[key]) === JSON.stringify(award[key])) delete merged[key];
      });
      return merged;
    });
  };
  const commit = (patch) => onChange({ ...award, ...patch });
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
  const dirtyKeys = Object.keys(draftSafe);
  const dirtyCount = dirtyKeys.length;
  const dirtyLabels = Array.from(new Set(dirtyKeys.map((k) => AWARD_SETTING_LABELS[k] || k)));
  const manualIds = (list) => list.filter((c) => c.type === "manual").map((c) => c.id).sort().join(",");
  const manualChangedAfterSend = award.shortlist.length > 0 && manualIds(award.criteria) !== manualIds(cfg.criteria);
  const blankCrit = cfg.criteria.some((c) => !String(c.name || "").trim());
  const saveSettings = async () => {
    if (!dirtyCount || saving) return;
    if (blankCrit) {
      setSaveMsg({ type: "error", text: "กรุณาตั้งชื่อเกณฑ์ให้ครบทุกข้อก่อนบันทึก" });
      return;
    }
    setSaving(true);
    const r = await onSaveSettings({ ...award, ...draftSafe });
    setSaving(false);
    if (r === "error") {
      setSaveMsg({ type: "error", text: "บันทึกลง Google Sheet ไม่สำเร็จ ตรวจสอบอินเทอร์เน็ตแล้วกด \"บันทึกการตั้งค่า\" อีกครั้ง (การแก้ไขของท่านยังอยู่ ไม่หาย)" });
      return;
    }
    onDraftChange({});
    setSaveMsg({
      type: "ok",
      text: r === "sheet"
        ? `บันทึกการตั้งค่าลง Google Sheet เรียบร้อยแล้ว (${nowStamp()})`
        : `บันทึกแล้ว (${nowStamp()}) · โหมดทดลอง: เก็บไว้เฉพาะในหน้านี้ จะหายเมื่อรีเฟรช ติดตั้ง Award_Addon.gs เพื่อบันทึกถาวร`,
    });
  };
  const discardDraft = () => {
    onDraftChange({});
    setConfirmCancel(false);
    setSaveMsg(null);
  };

  // ---- criteria
  const setCrit = (id, patch) => set({ criteria: cfg.criteria.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  const removeCrit = (id) => cfg.criteria.length > 1 && set({ criteria: cfg.criteria.filter((c) => c.id !== id) });
  const availableAuto = Object.keys(AUTO_SOURCES).filter((src) => !cfg.criteria.some((c) => c.type === "auto" && c.source === src));
  const addCritNew = () => {
    const name = newCrit.name.trim();
    if (!name) return;
    const base = { name, desc: newCrit.desc.trim(), weight: toNum(newCrit.weight) };
    const item = newCrit.type === "manual"
      ? { ...base, id: `m_${uid()}`, type: "manual" }
      : { ...base, id: cfg.criteria.some((c) => c.id === newCrit.type) ? `a_${uid()}` : newCrit.type, type: "auto", source: newCrit.type };
    set({ criteria: [...cfg.criteria, item] });
    setNewCrit({ name: "", desc: "", weight: 10, type: "manual" });
    setAdding(false);
  };
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
    if (!ids.length || dirtyCount > 0) return;
    commit({ shortlist: ids, sentAt: nowStamp(), sentBy: currentUser.name });
    onSendToJudges(ids.length);
    setSub("tracking");
  };

  // ล้างรายการที่ส่งให้กรรมการ เพื่อเลือกใหม่และส่งใหม่
  const scoredJudges = judges.filter((j) => Object.keys(cfg.scores?.[j.id] || {}).some((key) => key !== "__meta"));
  const clearShortlist = () => {
    commit({ shortlist: [], sentAt: null, sentBy: null });
    Object.keys(cfg.scores || {}).forEach((jid) => {
      const cur = cfg.scores[jid] || {};
      if (clearScoresToo) {
        if (Object.keys(cur).length) onScore(jid, {});
      } else if (cur.__meta) {
        const { __meta, ...rest } = cur;
        onScore(jid, rest); // เปิดให้กรรมการที่เคยกดส่งแล้ว กลับมาให้คะแนนเพิ่มได้ คะแนนเดิมยังอยู่
      }
    });
    setSelected(new Set());
    setClearing(false);
    onJudgesNotice(`Admin ล้างรายการที่ส่งให้ตัดสินใน ${cfg.roundName} แล้ว${clearScoresToo && scoredJudges.length ? " (คะแนนเดิมถูกล้าง)" : ""} กรุณารอรายการชุดใหม่`);
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
    commit({ decision: { confirmedBy: currentUser.name, confirmedAt: nowStamp(), winners } });
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
            <button onClick={() => commit({ decision: null })} className="text-xs px-3 py-1.5 rounded-lg border font-medium" style={{ borderColor: "#C9962B", color: "#6B4A10" }}>
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
            sub="คลิกที่ชื่อหรือคำอธิบายของแต่ละเกณฑ์เพื่อแก้ไขข้อความได้ · เพิ่ม/ลบเกณฑ์ได้ · เกณฑ์สีฟ้า (อัตโนมัติ) แก้ไขได้เฉพาะข้อความและน้ำหนัก วิธีคำนวณยังเป็นไปตามเดิม · เกณฑ์สีม่วงให้กรรมการให้คะแนน · รวมเป็นคะแนนเต็ม 100"
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
                    <div className="flex-1 min-w-0 space-y-1.5">
                      <input
                        value={c.name}
                        onChange={(e) => setCrit(c.id, { name: e.target.value })}
                        placeholder="ชื่อเกณฑ์ (คลิกเพื่อแก้ไข)"
                        title="คลิกเพื่อแก้ไขชื่อเกณฑ์"
                        className="w-full text-sm font-medium rounded-md border px-2 py-1 outline-none focus:border-blue-400"
                        style={{ color: INK, borderColor: String(c.name || "").trim() ? "#D8DDE5" : RED, background: "#FCFCFD" }}
                      />
                      <input
                        value={c.desc ?? (c.type === "auto" ? AUTO_SOURCES[c.source] : "กรรมการแต่ละท่านให้คะแนน 1-5 แยกกัน แล้วระบบหาค่าเฉลี่ย")}
                        onChange={(e) => setCrit(c.id, { desc: e.target.value })}
                        placeholder="คำอธิบายเกณฑ์ (ไม่บังคับ)"
                        title="คลิกเพื่อแก้ไขคำอธิบาย"
                        className="w-full text-[12px] rounded-md border px-2 py-1 outline-none focus:border-blue-400"
                        style={{ color: SUB, borderColor: "#E8ECF1", background: "#FCFCFD" }}
                      />
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
            {!adding ? (
              <button onClick={() => setAdding(true)} className="mt-3 text-sm font-medium flex items-center gap-1.5 px-3 py-1.5 rounded-lg border" style={{ color: NAVY, borderColor: NAVY }}>
                <Plus size={16} /> เพิ่มเกณฑ์ใหม่
              </button>
            ) : (
              <div className="mt-3 rounded-lg border p-3 space-y-2.5" style={{ borderColor: "#B9D2F0", background: "#F7FAFE" }}>
                <div className="text-sm font-semibold" style={{ color: INK }}>เพิ่มเกณฑ์ใหม่</div>
                <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                  <div className="md:col-span-2">
                    <Field label="ชื่อเกณฑ์" required>
                      <TextInput value={newCrit.name} onChange={(e) => setNewCrit((n) => ({ ...n, name: e.target.value }))} placeholder="เช่น ความยั่งยืน / ผลต่อสิ่งแวดล้อม" />
                    </Field>
                  </div>
                  <Field label="ประเภทเกณฑ์">
                    <select className={inputCls} style={inputStyle} value={newCrit.type} onChange={(e) => setNewCrit((n) => ({ ...n, type: e.target.value }))}>
                      <option value="manual">กรรมการให้คะแนน (1-5 ต่อท่าน)</option>
                      {availableAuto.map((src) => <option key={src} value={src}>คำนวณอัตโนมัติ: {AUTO_LABELS[src]}</option>)}
                    </select>
                  </Field>
                  <Field label="น้ำหนัก (%)">
                    <TextInput type="number" min="0" value={newCrit.weight} onChange={(e) => setNewCrit((n) => ({ ...n, weight: e.target.value }))} />
                  </Field>
                </div>
                <Field label="คำอธิบายเกณฑ์ (ไม่บังคับ — กรรมการจะเห็นใต้ชื่อเกณฑ์ตอนให้คะแนน)">
                  <TextInput value={newCrit.desc} onChange={(e) => setNewCrit((n) => ({ ...n, desc: e.target.value }))} placeholder="เช่น ลดการใช้พลังงาน น้ำ หรือของเสียได้จริงและต่อเนื่อง" />
                </Field>
                <div className="flex justify-end gap-2">
                  <button onClick={() => { setAdding(false); setNewCrit({ name: "", desc: "", weight: 10, type: "manual" }); }} className="px-3 py-1.5 rounded-lg text-sm border bg-white" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
                  <button onClick={addCritNew} disabled={!newCrit.name.trim()} className="px-4 py-1.5 rounded-lg text-sm font-semibold text-white disabled:opacity-40 flex items-center gap-1.5" style={{ background: NAVY }}>
                    <Plus size={15} /> เพิ่มเกณฑ์นี้
                  </button>
                </div>
                <div className="text-[11px]" style={{ color: SUB }}>เพิ่มแล้วจะอยู่ในรายการด้านบนเป็นร่าง ยังไม่มีผลจนกว่าจะกด "บันทึกการตั้งค่า"</div>
              </div>
            )}
          </AwardSection>

          <AwardSection title="คณะกรรมการ" sub="กรรมการแต่ละท่านมีบัญชีและรหัสผ่านของตัวเอง เข้าระบบแล้วจะเห็นเฉพาะหน้าให้คะแนนของตัวเอง ไม่เห็นคะแนนของกรรมการท่านอื่น">
            <div className="flex items-center gap-3 mb-3">
              <span className="text-sm" style={{ color: INK }}>จำนวนกรรมการรอบนี้</span>
              <select className="rounded-lg border px-3 py-1.5 text-sm" style={inputStyle} value={cfg.judgeCount} onChange={(e) => set({ judgeCount: Number(e.target.value) })}>
                {Array.from({ length: MAX_JUDGES }, (_, i) => i + 1).map((n) => <option key={n} value={n}>{n} ท่าน</option>)}
              </select>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2">
              {JUDGE_SLOTS.map((j, idx) => {
                const active = idx < cfg.judgeCount;
                const acct = USERS.find((u) => u.name === j.userName);
                return (
                  <div key={j.id} className="rounded-lg border p-3" style={{ borderColor: active ? "#DCD3F5" : "#EEF0F3", background: active ? "#FBFAFF" : "#F7F8FA", opacity: active ? 1 : 0.55 }}>
                    <div className="flex items-center gap-2">
                      <div className="w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold text-white" style={{ background: active ? JUDGE_PURPLE : "#AEB6C2" }}>{idx + 1}</div>
                      <div className="flex-1">
                        <div className="text-sm font-semibold" style={{ color: INK }}>{j.userName}</div>
                        <div className="text-[11px]" style={{ color: SUB }}>ชื่อผู้ใช้: {j.userName} · {acct ? "ตั้ง/เปลี่ยนรหัสผ่านที่เมนู จัดการผู้ใช้" : "ยังไม่มีบัญชี"}</div>
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
          <AwardSection title="ระดับดาวของ Kaizen แต่ละเรื่อง" sub="ทุกบทบาทเห็นดาวบน Kaizen แต่ละเรื่อง คำนวณจาก 'คะแนนคัดกรอง' ตามเกณฑ์อัตโนมัติด้านบน (Cost Saving, ผลกระทบ, ROI, ทีมงาน) ตั้งเกณฑ์ขั้นต่ำของแต่ละระดับได้ที่นี่">
            <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
              {[5, 4, 3, 2].map((stars, i) => (
                <div key={stars}>
                  <div className="flex items-center gap-0.5 mb-1">
                    {Array.from({ length: stars }).map((_, j) => <Star key={j} size={14} fill={GOLD} style={{ color: GOLD }} />)}
                    <span className="text-xs ml-1" style={{ color: SUB }}>ตั้งแต่</span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <input
                      type="number"
                      min={0}
                      max={100}
                      value={cfg.starThresholds[i]}
                      onChange={(e) => set({ starThresholds: cfg.starThresholds.map((v, j) => (j === i ? toNum(e.target.value) : v)) })}
                      className="w-full rounded border px-2 py-1 text-sm text-right"
                      style={inputStyle}
                    />
                    <span className="text-xs" style={{ color: SUB }}>%</span>
                  </div>
                </div>
              ))}
              <div>
                <div className="flex items-center gap-0.5 mb-1">
                  <Star size={14} fill={GOLD} style={{ color: GOLD }} />
                  <span className="text-xs ml-1" style={{ color: SUB }}>ที่เหลือ</span>
                </div>
                <div className="text-xs pt-1.5" style={{ color: SUB }}>ต่ำกว่าเกณฑ์ 2 ดาว</div>
              </div>
            </div>
          </AwardSection>
          <div className="flex items-center justify-between gap-3 flex-wrap">
            <button
              onClick={saveSettings}
              disabled={!dirtyCount || saving}
              className="px-4 py-2 rounded-lg text-sm font-semibold text-white flex items-center gap-1.5 disabled:opacity-60"
              style={{ background: dirtyCount ? GREEN : "#8FA3BA" }}
            >
              <CheckCircle2 size={16} /> {saving ? "กำลังบันทึก..." : dirtyCount ? `บันทึกการตั้งค่าทั้งหมด (${dirtyLabels.length} หัวข้อ)` : "บันทึกการตั้งค่าแล้ว"}
            </button>
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
              {dirtyCount > 0 && (
                <div className="text-xs mt-1" style={{ color: RED }}>มีการแก้ไขการตั้งค่าที่ยังไม่ได้บันทึก กรุณากด "บันทึกการตั้งค่า" ก่อนส่งให้กรรมการ</div>
              )}
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              {cfg.shortlist.length > 0 && (
                <button
                  onClick={() => { setClearScoresToo(true); setClearing((c) => !c); }}
                  className="px-3 py-2 rounded-lg text-sm font-medium border flex items-center gap-1.5"
                  style={{ borderColor: RED, color: RED, background: clearing ? "#FBEAE8" : "white" }}
                  title="ล้างรายการที่ส่งให้กรรมการ เพื่อเลือกใหม่และส่งใหม่"
                >
                  <Trash2 size={15} /> ล้างข้อมูลที่ส่งให้กรรมการ
                </button>
              )}
              <button
                onClick={sendShortlist}
                disabled={!selected.size || dirtyCount > 0 || (!shortlistChanged && cfg.shortlist.length > 0)}
                className="px-4 py-2 rounded-lg text-sm font-semibold text-white flex items-center gap-1.5 disabled:opacity-40"
                style={{ background: JUDGE_PURPLE }}
              >
                <Trophy size={16} /> {cfg.shortlist.length ? "อัปเดตรายการที่ส่งกรรมการ" : `ส่ง ${selected.size} เรื่องให้กรรมการตัดสิน`}
              </button>
            </div>
          </div>
          {clearing && cfg.shortlist.length > 0 && (
            <div className="rounded-xl border p-4 space-y-3" style={{ background: "#FBEAE8", borderColor: "#F0C3BD" }}>
              <div className="text-sm font-semibold" style={{ color: "#8B2E24" }}>ล้างข้อมูลที่ส่งให้กรรมการ แล้วเริ่มเลือกใหม่?</div>
              <ul className="text-sm space-y-1" style={{ color: INK }}>
                <li>• รายการที่ส่งไปแล้ว <strong>{cfg.shortlist.length} เรื่อง</strong> จะถูกถอนออกจากหน้ากรรมการทุกท่าน และช่องที่ติ๊กเลือกไว้ด้านบนจะถูกล้าง</li>
                {scoredJudges.length > 0 && (
                  <li>• มีกรรมการ <strong>{scoredJudges.length} ท่าน</strong> ที่ให้คะแนนไว้แล้ว (กดส่งแล้ว {submittedCount} ท่าน)</li>
                )}
                <li>• ต้องกด "ส่ง ... เรื่องให้กรรมการตัดสิน" อีกครั้ง กรรมการจึงจะเห็นรายการชุดใหม่</li>
              </ul>
              {scoredJudges.length > 0 && (
                <div className="rounded-lg p-3 space-y-2" style={{ background: "rgba(255,255,255,0.7)" }}>
                  <label className="flex items-start gap-2 text-sm" style={{ color: INK }}>
                    <input type="checkbox" className="mt-0.5" checked={clearScoresToo} onChange={(e) => setClearScoresToo(e.target.checked)} />
                    <span>
                      <strong>ล้างคะแนนของกรรมการทั้งหมดด้วย</strong> (แนะนำ เมื่อจะเลือกเรื่องชุดใหม่)
                      <span className="block text-[11px] mt-0.5" style={{ color: SUB }}>
                        ถ้าไม่ติ๊ก: คะแนนเดิมจะยังอยู่ ถ้าเรื่องนั้นถูกเลือกกลับมาในชุดใหม่ คะแนนจะกลับมาใช้ต่อ และกรรมการที่เคยกดส่งแล้วจะถูกเปิดให้ให้คะแนนเพิ่มได้
                      </span>
                    </span>
                  </label>
                  <button onClick={exportExcel} className="text-xs px-3 py-1.5 rounded-lg border font-medium flex items-center gap-1.5 bg-white" style={{ borderColor: GREEN, color: GREEN }}>
                    <FileSpreadsheet size={14} /> Export Excel เก็บคะแนนชุดนี้ไว้ก่อน
                  </button>
                </div>
              )}
              <div className="flex justify-end gap-2">
                <button onClick={() => setClearing(false)} className="px-3 py-2 rounded-lg text-sm border bg-white" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
                <button onClick={clearShortlist} className="px-4 py-2 rounded-lg text-sm font-semibold text-white flex items-center gap-1.5" style={{ background: RED }}>
                  <Trash2 size={15} /> ยืนยันล้างข้อมูล
                </button>
              </div>
            </div>
          )}
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

      {(dirtyCount > 0 || saveMsg) && (
        <div
          className="sticky bottom-3 z-20 rounded-xl border p-3 flex items-center justify-between gap-3 flex-wrap bg-white"
          style={{ borderColor: dirtyCount > 0 ? "#E8C77E" : saveMsg?.type === "ok" ? "#A9D6BF" : "#F0C3BD", boxShadow: "0 6px 20px rgba(11,29,51,0.12)" }}
        >
          <div className="text-sm flex-1 min-w-[260px]" style={{ color: INK }}>
            {dirtyCount > 0 ? (
              <>
                <strong style={{ color: "#8A5B14" }}>มีการแก้ไขที่ยังไม่ได้บันทึก</strong> · {dirtyLabels.join(", ")}
                {totalW !== 100 && <div className="text-xs mt-0.5" style={{ color: "#8A5B14" }}>น้ำหนักรวมตอนนี้ {totalW}% (ไม่ครบ 100% ระบบจะคำนวณเป็นสัดส่วนให้ แต่แนะนำให้กด "ปรับให้รวม 100%")</div>}
                {manualChangedAfterSend && <div className="text-xs mt-0.5" style={{ color: RED }}>ได้ส่งเรื่องให้กรรมการไปแล้ว การเพิ่ม/ลบเกณฑ์ที่กรรมการให้คะแนนจะทำให้ความคืบหน้าเปลี่ยน กรรมการที่กดส่งแล้วต้องให้ Admin เปิดสิทธิ์แก้ไขที่แท็บ "3. ติดตามกรรมการ"</div>}
                {blankCrit && <div className="text-xs mt-0.5" style={{ color: RED }}>มีเกณฑ์ที่ยังไม่ได้ตั้งชื่อ</div>}
              </>
            ) : null}
            {saveMsg && <div className={dirtyCount > 0 ? "text-xs mt-0.5" : ""} style={{ color: saveMsg.type === "ok" ? GREEN : RED }}>{saveMsg.text}</div>}
          </div>
          {dirtyCount > 0 && (
            <div className="flex items-center gap-2">
              {confirmCancel ? (
                <>
                  <button onClick={() => setConfirmCancel(false)} className="px-3 py-2 rounded-lg text-sm border" style={{ borderColor: "#D8DDE5", color: SUB }}>ไม่ยกเลิก</button>
                  <button onClick={discardDraft} className="px-3 py-2 rounded-lg text-sm font-semibold text-white" style={{ background: RED }}>ยืนยันยกเลิกการแก้ไข</button>
                </>
              ) : (
                <button onClick={() => setConfirmCancel(true)} className="px-3 py-2 rounded-lg text-sm border" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิกการแก้ไข</button>
              )}
              <button onClick={saveSettings} disabled={saving} className="px-4 py-2 rounded-lg text-sm font-semibold text-white flex items-center gap-1.5 disabled:opacity-60" style={{ background: GREEN }}>
                <CheckCircle2 size={16} /> {saving ? "กำลังบันทึก..." : "บันทึกการตั้งค่า"}
              </button>
            </div>
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
                      {c.desc ? <div className="text-[11px] mb-1" style={{ color: SUB }}>{c.desc}</div> : null}
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

// ================= จัดการผู้ใช้ (Admin) =================
const USER_ROLE_ORDER = ["admin", "director", "supervisor", "employee", "judge"];
const USER_ROLE_OPTIONS = [["employee", "พนักงาน (บัญชีหน่วยงาน)"], ["supervisor", "หัวหน้างาน"], ["admin", "Admin (ฝ่ายบุคคล)"], ["director", "ผู้อำนวยการโรงงาน"]];
const ROLE_BADGE = {
  admin: { bg: "#FDF3E1", color: "#8A5B14" }, director: { bg: "#FDF3E1", color: "#8A5B14" }, supervisor: { bg: "#EAF3FC", color: "#1D5FC9" },
  employee: { bg: "#F1F3F6", color: "#5B6675" }, judge: { bg: "#F1ECFD", color: "#6A4BC9" },
};

const savedNoteText = (saved) =>
  saved === "sheet" ? "บันทึกลง Google Sheet แล้ว"
  : saved === "error" ? "ใช้งานได้ในหน้านี้ แต่บันทึกลง Google Sheet ไม่สำเร็จ"
  : "โหมดทดลอง: เก็บเฉพาะในหน้านี้ จะหายเมื่อรีเฟรช (ติดตั้ง Users_Addon.gs / Code.gs ฉบับล่าสุดเพื่อบันทึกถาวร)";

function DeptResultNote({ saved, ops, onClear }) {
  const [busy, setBusy] = useState(false);
  const [now, setNow] = useState(saved);
  if (!saved) return null;
  const cur = now || saved;
  return (
    <div className="text-[11px] flex items-center gap-2 flex-wrap" style={{ color: cur === "error" ? "#8B2E24" : cur === "sheet" ? GREEN : "#8A5B14" }}>
      <span>{savedNoteText(cur)}</span>
      {cur === "error" && (
        <button type="button" disabled={busy} className="underline" onClick={async () => { setBusy(true); const r = await ops.resave(); setBusy(false); setNow(r); }}>{busy ? "กำลังบันทึก..." : "ลองบันทึกอีกครั้ง"}</button>
      )}
      {onClear && <button type="button" className="underline" onClick={onClear}>ปิด</button>}
    </div>
  );
}

const defaultSectionName = () => (SECTION_OPTIONS.includes("Support Function") ? "Support Function" : SECTION_OPTIONS[0]);

// เลือกฝ่าย (Section) + ปุ่มเพิ่มฝ่ายใหม่ได้ทันทีโดยไม่ต้องออกจากฟอร์ม
function SectionPicker({ value, onChange, ops }) {
  const [adding, setAdding] = useState(false);
  const [name, setName] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [saved, setSaved] = useState(null);
  const submit = async () => {
    if (busy) return;
    setBusy(true);
    const r = await ops.sections.add(name);
    setBusy(false);
    if (!r.ok) { setErr(r.error); return; }
    onChange(r.name); setSaved(r.saved); setAdding(false); setName(""); setErr("");
  };
  return (
    <div className="space-y-1.5">
      <div className="flex items-center gap-1.5">
        <select value={value} onChange={(e) => onChange(e.target.value)} className="rounded-lg border px-2.5 py-1.5 text-sm" style={inputStyle} title="ฝ่าย (Section) ที่หน่วยงานนี้สังกัด ใช้จัดกลุ่มในแดชบอร์ด">
          {SECTION_OPTIONS.map((x) => <option key={x} value={x}>{x}</option>)}
        </select>
        <button type="button" onClick={() => setAdding((a) => !a)} className="px-2 py-1.5 rounded-lg border text-xs font-medium whitespace-nowrap" style={{ borderColor: NAVY, color: NAVY }}>+ ฝ่ายใหม่</button>
      </div>
      {adding && (
        <div className="flex flex-wrap items-center gap-1.5">
          <input autoFocus value={name} onChange={(e) => { setName(e.target.value); setErr(""); }} onKeyDown={(e) => e.key === "Enter" && submit()} placeholder="ชื่อฝ่ายใหม่ เช่น Maintenance" className="rounded-lg border px-2.5 py-1 text-sm" style={inputStyle} />
          <button type="button" onClick={submit} disabled={busy} className="px-2.5 py-1 rounded-lg text-xs font-semibold text-white disabled:opacity-60" style={{ background: NAVY }}>{busy ? "กำลังเพิ่ม..." : "เพิ่มฝ่าย"}</button>
          <button type="button" onClick={() => { setAdding(false); setErr(""); }} className="px-2.5 py-1 rounded-lg text-xs border bg-white" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
        </div>
      )}
      {err && <div className="text-xs" style={{ color: "#8B2E24" }}>{err}</div>}
      {saved && !adding && <div className="text-[11px]" style={{ color: saved === "error" ? "#8B2E24" : saved === "sheet" ? GREEN : "#8A5B14" }}>เพิ่มฝ่ายแล้ว · {savedNoteText(saved)}</div>}
    </div>
  );
}

function SectionAddForm({ ops, onDone, onCancel }) {
  const [name, setName] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const submit = async () => {
    if (busy) return;
    setBusy(true);
    const r = await ops.sections.add(name);
    setBusy(false);
    if (!r.ok) { setErr(r.error); return; }
    onDone(r);
  };
  return (
    <div className="rounded-lg border p-3 space-y-2" style={{ borderColor: "#B9D2F0", background: "#F7FAFE" }}>
      <div className="text-xs font-semibold" style={{ color: INK }}>เพิ่มฝ่าย (Section) ใหม่</div>
      <div className="flex flex-wrap gap-2 items-center">
        <input autoFocus value={name} onChange={(e) => { setName(e.target.value); setErr(""); }} onKeyDown={(e) => e.key === "Enter" && submit()} placeholder="ชื่อฝ่ายใหม่ เช่น Maintenance" className="flex-1 min-w-[180px] rounded-lg border px-3 py-1.5 text-sm" style={inputStyle} />
        <button type="button" onClick={submit} disabled={busy} className="px-3 py-1.5 rounded-lg text-sm font-semibold text-white disabled:opacity-60" style={{ background: NAVY }}>{busy ? "กำลังเพิ่ม..." : "เพิ่มฝ่าย"}</button>
        <button type="button" onClick={onCancel} className="px-3 py-1.5 rounded-lg text-sm border bg-white" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
      </div>
      <div className="text-[11px]" style={{ color: SUB }}>ฝ่ายใหม่ยังไม่มีหน่วยงาน กด "+ หน่วยงาน" ที่หัวฝ่าย หรือย้ายหน่วยงานเดิมเข้ามาด้วยปุ่มแก้ไขหน่วยงาน ฝ่ายใหม่จะแสดงในตัวกรองแดชบอร์ดด้วย</div>
      {err && <div className="text-xs" style={{ color: "#8B2E24" }}>{err}</div>}
    </div>
  );
}

function SectionEditForm({ name, ops, onDone, onCancel }) {
  const inSec = ops.sections.stats(name).departments;
  const [newName, setNewName] = useState(name);
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const submit = async () => {
    if (busy) return;
    setBusy(true);
    const r = await ops.sections.rename(name, newName);
    setBusy(false);
    if (!r.ok) { setErr(r.error); return; }
    onDone(r);
  };
  return (
    <div className="rounded-lg border p-3 space-y-2" style={{ borderColor: "#E8C77E", background: "#FFFBF2" }}>
      <div className="text-xs font-semibold" style={{ color: INK }}>แก้ไขชื่อฝ่าย: {name}</div>
      <div className="flex flex-wrap gap-2 items-center">
        <input autoFocus value={newName} onChange={(e) => { setNewName(e.target.value); setErr(""); }} onKeyDown={(e) => e.key === "Enter" && submit()} className="flex-1 min-w-[180px] rounded-lg border px-3 py-1.5 text-sm" style={inputStyle} />
        <button type="button" onClick={submit} disabled={busy} className="px-3 py-1.5 rounded-lg text-sm font-semibold text-white disabled:opacity-60" style={{ background: GREEN }}>{busy ? "กำลังบันทึก..." : "บันทึก"}</button>
        <button type="button" onClick={onCancel} className="px-3 py-1.5 rounded-lg text-sm border bg-white" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
      </div>
      <div className="text-[11px]" style={{ color: "#6B4A10" }}>ฝ่ายนี้มี {inSec.length} หน่วยงาน · เปลี่ยนชื่อแล้ว หน่วยงานและตัวกรอง/กราฟในแดชบอร์ดเปลี่ยนตามอัตโนมัติ (ไม่ต้องแก้ Kaizen)</div>
      {err && <div className="text-xs" style={{ color: "#8B2E24" }}>{err}</div>}
    </div>
  );
}

function SectionDeleteConfirm({ name, ops, onDone, onCancel }) {
  const can = ops.sections.canRemove(name);
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const confirm = async () => {
    if (busy) return;
    setBusy(true);
    const r = await ops.sections.remove(name);
    setBusy(false);
    if (!r.ok) { setErr(r.error); return; }
    onDone(r);
  };
  return (
    <div className="rounded-lg border p-3 space-y-2" style={{ borderColor: "#F0C3BD", background: "#FDF4F3" }}>
      <div className="text-xs font-semibold" style={{ color: "#8B2E24" }}>ลบฝ่าย "{name}" ?</div>
      {can.ok ? <div className="text-[11px]" style={{ color: INK }}>ฝ่ายนี้ไม่มีหน่วยงาน ลบได้ และจะหายจากตัวกรองในแดชบอร์ด</div> : <div className="text-xs" style={{ color: "#8B2E24" }}>{can.error}</div>}
      {err && <div className="text-xs" style={{ color: "#8B2E24" }}>{err}</div>}
      <div className="flex gap-2">
        <button type="button" onClick={confirm} disabled={!can.ok || busy} className="px-3 py-1.5 rounded-lg text-sm font-semibold text-white disabled:opacity-40" style={{ background: RED }}>{busy ? "กำลังลบ..." : "ยืนยันลบ"}</button>
        <button type="button" onClick={onCancel} className="px-3 py-1.5 rounded-lg text-sm border bg-white" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
      </div>
    </div>
  );
}

// จัดกลุ่มหน่วยงานตามฝ่าย: แสดงทุกฝ่ายตามลำดับในรายการ (รวมฝ่ายที่ยังไม่มีหน่วยงาน) แล้วต่อด้วยหน่วยงานที่ไม่มีฝ่าย
function groupBySection(deptList) {
  const m = {};
  deptList.forEach((d) => { const sec = SECTION_MAP[d] || "อื่นๆ"; (m[sec] = m[sec] || []).push(d); });
  const out = SECTION_OPTIONS.map((sec) => [sec, m[sec] || []]);
  Object.keys(m).forEach((sec) => { if (!SECTION_OPTIONS.includes(sec)) out.push([sec, m[sec]]); });
  return out;
}

function DeptAddForm({ ops, onDone, onCancel, defaultSection }) {
  const [name, setName] = useState("");
  const [section, setSection] = useState(defaultSection && SECTION_OPTIONS.includes(defaultSection) ? defaultSection : defaultSectionName());
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const submit = async () => {
    if (busy) return;
    setBusy(true);
    const r = await ops.add(name, section);
    setBusy(false);
    if (!r.ok) { setErr(r.error); return; }
    onDone(r);
  };
  return (
    <div className="rounded-lg border p-3 space-y-2" style={{ borderColor: "#B9D2F0", background: "#F7FAFE" }}>
      <div className="text-xs font-semibold" style={{ color: INK }}>เพิ่มหน่วยงานใหม่</div>
      <div className="flex flex-wrap gap-2 items-center">
        <input autoFocus value={name} onChange={(e) => { setName(e.target.value); setErr(""); }} onKeyDown={(e) => e.key === "Enter" && submit()} placeholder="ชื่อหน่วยงานใหม่ เช่น Packing 5" className="flex-1 min-w-[180px] rounded-lg border px-3 py-1.5 text-sm" style={inputStyle} />
        <SectionPicker value={section} onChange={setSection} ops={ops} />
        <button type="button" onClick={submit} disabled={busy} className="px-3 py-1.5 rounded-lg text-sm font-semibold text-white disabled:opacity-60" style={{ background: NAVY }}>{busy ? "กำลังเพิ่ม..." : "เพิ่ม"}</button>
        <button type="button" onClick={onCancel} className="px-3 py-1.5 rounded-lg text-sm border bg-white" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
      </div>
      <div className="text-[11px]" style={{ color: SUB }}>เลือกฝ่าย (Section) ที่หน่วยงานนี้สังกัด เพื่อให้แดชบอร์ดจัดกลุ่มถูกต้อง</div>
      {err && <div className="text-xs" style={{ color: "#8B2E24" }}>{err}</div>}
    </div>
  );
}

function DeptEditForm({ name, ops, onDone, onCancel }) {
  const st = ops.stats(name);
  const [newName, setNewName] = useState(name);
  const [section, setSection] = useState(SECTION_MAP[name] || defaultSectionName());
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const changed = newName.replace(/\s+/g, " ").trim() !== name;
  const submit = async () => {
    if (busy) return;
    setBusy(true);
    const r = await ops.rename(name, newName, section);
    setBusy(false);
    if (!r.ok) { setErr(r.error); return; }
    onDone(r);
  };
  return (
    <div className="rounded-lg border p-3 space-y-2" style={{ borderColor: "#E8C77E", background: "#FFFBF2" }}>
      <div className="text-xs font-semibold" style={{ color: INK }}>แก้ไขหน่วยงาน: {name}</div>
      <div className="flex flex-wrap gap-2 items-center">
        <input autoFocus value={newName} onChange={(e) => { setNewName(e.target.value); setErr(""); }} onKeyDown={(e) => e.key === "Enter" && submit()} className="flex-1 min-w-[180px] rounded-lg border px-3 py-1.5 text-sm" style={inputStyle} />
        <SectionPicker value={section} onChange={setSection} ops={ops} />
        <button type="button" onClick={submit} disabled={busy} className="px-3 py-1.5 rounded-lg text-sm font-semibold text-white disabled:opacity-60" style={{ background: GREEN }}>{busy ? "กำลังบันทึก..." : "บันทึก"}</button>
        <button type="button" onClick={onCancel} className="px-3 py-1.5 rounded-lg text-sm border bg-white" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
      </div>
      <div className="text-[11px]" style={{ color: "#6B4A10" }}>
        หน่วยงานนี้มี Kaizen {st.kaizens} เรื่อง และมี {st.users.length} บัญชีใช้อยู่
        {changed ? " · เมื่อเปลี่ยนชื่อ ระบบแก้ชื่อในบัญชีและ Kaizen เหล่านี้ให้ทั้งหมดอัตโนมัติ" + (st.kaizens > 0 ? " (อัปเดต Kaizen ทีละเรื่อง กรุณาอย่าปิดหน้าจนกว่าจะขึ้นสถานะบันทึกแล้ว)" : "") : ""}
      </div>
      {err && <div className="text-xs" style={{ color: "#8B2E24" }}>{err}</div>}
    </div>
  );
}

function DeptDeleteConfirm({ name, ops, onDone, onCancel }) {
  const st = ops.stats(name);
  const can = ops.canRemove(name);
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const confirm = async () => {
    if (busy) return;
    setBusy(true);
    const r = await ops.remove(name);
    setBusy(false);
    if (!r.ok) { setErr(r.error); return; }
    onDone(r);
  };
  return (
    <div className="rounded-lg border p-3 space-y-2" style={{ borderColor: "#F0C3BD", background: "#FDF4F3" }}>
      <div className="text-xs font-semibold" style={{ color: "#8B2E24" }}>ลบหน่วยงาน "{name}" ออกจากระบบ?</div>
      {can.ok ? (
        <div className="text-[11px]" style={{ color: INK }}>
          {st.users.length > 0 ? `หน่วยงานนี้จะถูกเอาออกจาก ${st.users.length} บัญชี (${st.users.map((u) => u.name).join(", ")}) ` : "ไม่มีบัญชีหรือ Kaizen ที่ใช้หน่วยงานนี้ "}
          · มีผลกับทุกหน้าของระบบ
        </div>
      ) : (
        <div className="text-xs" style={{ color: "#8B2E24" }}>{can.error}</div>
      )}
      {err && <div className="text-xs" style={{ color: "#8B2E24" }}>{err}</div>}
      <div className="flex gap-2">
        <button type="button" onClick={confirm} disabled={!can.ok || busy} className="px-3 py-1.5 rounded-lg text-sm font-semibold text-white disabled:opacity-40" style={{ background: RED }}>{busy ? "กำลังลบ..." : "ยืนยันลบ"}</button>
        <button type="button" onClick={onCancel} className="px-3 py-1.5 rounded-lg text-sm border bg-white" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
      </div>
    </div>
  );
}

function TinyIcon({ icon: Icon, label, onClick, color }) {
  return (
    <button type="button" title={label} aria-label={label} onClick={onClick} className="p-0.5 rounded hover:bg-slate-100" style={{ color }}>
      <Icon size={12} />
    </button>
  );
}

function UserFormModal({ mode, original, users, departments, deptOps, currentUser, focusPw, onClose, onSubmit }) {
  const isNew = mode === "new";
  const isJudge = original?.role === "judge";
  const isSelf = !!original && original.id === currentUser.id;
  const [f, setF] = useState(() => ({
    name: original?.name || "", role: original?.role || "employee", departments: original?.departments || [],
    active: original ? original.active !== false : true, pw: isNew ? randomPassword() : "",
  }));
  const [showPw, setShowPw] = useState(isNew);
  const [errors, setErrors] = useState([]);
  const [saving, setSaving] = useState(false);
  const needsDept = f.role === "employee" || f.role === "supervisor";
  const set = (patch) => { setErrors([]); setF((x) => ({ ...x, ...patch })); };
  const [deptPanel, setDeptPanel] = useState(null); // { type: "add" | "edit" | "delete", name? }
  const [deptNote, setDeptNote] = useState(null);
  const onSectionDone = (r) => { setDeptPanel(null); setDeptNote(r.saved); };
  const onDeptAdded = (r) => { setDeptPanel(null); setDeptNote(r.saved); set({ departments: [...f.departments, r.name] }); };
  const onDeptRenamed = (oldName, r) => { setDeptPanel(null); setDeptNote(r.saved); set({ departments: f.departments.map((d) => (d === oldName ? r.name : d)) }); };
  const onDeptRemoved = (name, r) => { setDeptPanel(null); setDeptNote(r.saved); set({ departments: f.departments.filter((d) => d !== name) }); };
  const activeAdmins = users.filter((u) => u.role === "admin" && u.active !== false);
  const allDepts = useMemo(() => Array.from(new Set([...(departments || []), ...f.departments])), [departments, f.departments]);
  const groups = groupBySection(allDepts);
  const toggleDept = (d) => set({ departments: f.departments.includes(d) ? f.departments.filter((x) => x !== d) : [...f.departments, d] });
  const toggleGroup = (list) => {
    const all = list.every((d) => f.departments.includes(d));
    set({ departments: all ? f.departments.filter((d) => !list.includes(d)) : Array.from(new Set([...f.departments, ...list])) });
  };

  const validate = () => {
    const errs = [];
    const name = f.name.trim();
    if (!name) errs.push("กรุณาใส่ชื่อผู้ใช้");
    if (name.length > 60) errs.push("ชื่อผู้ใช้ยาวเกินไป (ไม่เกิน 60 ตัวอักษร)");
    const renamed = !original || normName(original.name) !== normName(name);
    if (name && renamed && users.some((u) => u.id !== original?.id && normName(u.name) === normName(name))) errs.push(`ชื่อผู้ใช้ "${name}" ซ้ำกับบัญชีอื่น`);
    if (needsDept && f.departments.length === 0) errs.push("กรุณาเลือกหน่วยงานที่ดูแลอย่างน้อย 1 หน่วยงาน");
    if (f.role === "supervisor") {
      f.departments.forEach((d) => {
        const other = users.find((u) => u.id !== original?.id && u.role === "supervisor" && u.active !== false && u.departments.includes(d));
        if (other) errs.push(`หน่วยงาน "${d}" มีหัวหน้างาน "${other.name}" ดูแลอยู่แล้ว (หน่วยงานหนึ่งมีหัวหน้างานได้ 1 บัญชี ย้ายออกจากบัญชีนั้นก่อน)`);
      });
    }
    const pw = f.pw.trim();
    if (isNew && !pw) errs.push("กรุณาตั้งรหัสผ่าน");
    if (pw && pw.length < 4) errs.push("รหัสผ่านต้องมีอย่างน้อย 4 ตัวอักษร");
    if (original && original.role === "admin" && (f.role !== "admin" || !f.active) && activeAdmins.length <= 1) errs.push("ต้องมี Admin ที่ใช้งานได้อย่างน้อย 1 บัญชี");
    return errs;
  };

  const submit = async () => {
    const errs = validate();
    if (errs.length) { setErrors(errs); return; }
    setSaving(true);
    const plain = f.pw.trim();
    const pw = plain ? await makePw(plain) : original.pw;
    const name = f.name.trim();
    let aliases = [...(original?.aliases || [])];
    if (original && normName(original.name) !== normName(name) && !aliases.some((a) => normName(a) === normName(original.name))) aliases.push(original.name);
    aliases = aliases.filter((a) => normName(a) !== normName(name));
    const user = { id: original?.id || newUserId(), name, role: f.role, departments: needsDept ? f.departments : [], aliases, active: f.active, pw };
    await onSubmit(user, plain || null, isNew);
    setSaving(false);
  };

  const strength = f.pw.trim().length === 0 ? null : f.pw.trim().length < 6 ? "สั้นและเดาง่าย แนะนำ 6 ตัวอักษรขึ้นไป ผสมตัวอักษรกับตัวเลข" : "ใช้ได้";
  return (
    <Modal title={isNew ? "เพิ่มผู้ใช้ใหม่" : `แก้ไขผู้ใช้: ${original.name}`} onClose={onClose} wide>
      <div className="space-y-4">
        {isJudge && (
          <div className="rounded-lg px-3 py-2 text-xs" style={{ background: "#F1ECFD", color: "#6A4BC9" }}>บัญชีกรรมการผูกกับระบบ Kaizen Award จึงเปลี่ยนชื่อ บทบาท หรือปิดใช้งานไม่ได้ ตั้งรหัสผ่านใหม่ได้เท่านั้น</div>
        )}
        <div className="grid grid-cols-2 gap-4">
          <Field label="ชื่อผู้ใช้ (ใช้พิมพ์ตอนเข้าสู่ระบบ)" required>
            <TextInput value={f.name} disabled={isJudge} onChange={(e) => set({ name: e.target.value })} placeholder="เช่น ชื่อ-นามสกุล หรือชื่อหน่วยงาน" />
          </Field>
          <Field label="บทบาท" required>
            <select className={inputCls} style={inputStyle} value={f.role} disabled={isJudge || isSelf} onChange={(e) => set({ role: e.target.value })}>
              {isJudge ? <option value="judge">กรรมการ Kaizen Award</option> : USER_ROLE_OPTIONS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
            </select>
          </Field>
        </div>
        {original && f.role === "employee" && normName(original.name) !== normName(f.name) && (
          <div className="text-xs rounded-lg px-3 py-2" style={{ background: "#EAF3FC", color: "#1D5FC9" }}>เมื่อเปลี่ยนชื่อ ระบบเก็บชื่อเดิม "{original.name}" ไว้ให้ Kaizen ที่เคยส่งด้วยชื่อเดิมยังแสดงในบัญชีนี้</div>
        )}
        {needsDept && (
          <div className="space-y-2">
            <div className="flex items-center justify-between gap-2 flex-wrap">
              <div className="text-xs font-medium" style={{ color: SUB }}>
                {f.role === "supervisor" ? "หน่วยงานที่ดูแล (อนุมัติเรื่องของหน่วยงานเหล่านี้)" : "หน่วยงานที่บัญชีนี้เลือกส่งเรื่องได้"} <span style={{ color: RED }}>*</span>
                <span className="ml-2 font-normal" style={{ color: "#A3ABB8" }}>เลือกแล้ว {f.departments.length} หน่วยงาน</span>
              </div>
              <div className="flex items-center gap-1.5">
                <button type="button" onClick={() => setDeptPanel({ type: "sec-add" })} className="text-xs font-medium px-2.5 py-1 rounded-lg border flex items-center gap-1" style={{ borderColor: NAVY, color: NAVY }}>
                  <Plus size={13} /> เพิ่มฝ่ายใหม่
                </button>
                <button type="button" onClick={() => setDeptPanel({ type: "add" })} className="text-xs font-medium px-2.5 py-1 rounded-lg border flex items-center gap-1" style={{ borderColor: NAVY, color: NAVY }}>
                  <Plus size={13} /> เพิ่มหน่วยงานใหม่
                </button>
              </div>
            </div>
            {deptPanel && deptPanel.type === "sec-add" && <SectionAddForm ops={deptOps} onDone={onSectionDone} onCancel={() => setDeptPanel(null)} />}
            {deptPanel && deptPanel.type === "add" && <DeptAddForm key={deptPanel.section || "x"} defaultSection={deptPanel.section} ops={deptOps} onDone={onDeptAdded} onCancel={() => setDeptPanel(null)} />}
            <div className="rounded-lg border max-h-72 overflow-y-auto p-2 space-y-2.5" style={{ borderColor: "#D8DDE5" }}>
              {groups.map(([sec, list]) => (
                <div key={sec}>
                  <div className="flex items-center gap-2">
                    {list.length > 0 && <input type="checkbox" aria-label={`เลือกทั้งฝ่าย ${sec}`} checked={list.every((d) => f.departments.includes(d))} onChange={() => toggleGroup(list)} />}
                    <span className="text-xs font-semibold" style={{ color: NAVY }}>{sec}</span>
                    {sec !== "อื่นๆ" && (
                      <>
                        <button type="button" onClick={() => setDeptPanel({ type: "add", section: sec })} className="text-[11px] px-1.5 py-0.5 rounded border" style={{ borderColor: "#D8DDE5", color: SUB }} title={`เพิ่มหน่วยงานในฝ่าย ${sec}`}>+ หน่วยงาน</button>
                        <TinyIcon icon={Pencil} label={`แก้ไขฝ่าย ${sec}`} color="#6A4BC9" onClick={() => setDeptPanel({ type: "sec-edit", name: sec })} />
                        <TinyIcon icon={Trash2} label={`ลบฝ่าย ${sec}`} color={RED} onClick={() => setDeptPanel({ type: "sec-delete", name: sec })} />
                      </>
                    )}
                  </div>
                  {deptPanel && deptPanel.type === "sec-edit" && deptPanel.name === sec && (
                    <div className="mt-2 pl-6"><SectionEditForm key={sec} name={sec} ops={deptOps} onDone={onSectionDone} onCancel={() => setDeptPanel(null)} /></div>
                  )}
                  {deptPanel && deptPanel.type === "sec-delete" && deptPanel.name === sec && (
                    <div className="mt-2 pl-6"><SectionDeleteConfirm key={sec} name={sec} ops={deptOps} onDone={onSectionDone} onCancel={() => setDeptPanel(null)} /></div>
                  )}
                  {list.length === 0 && <div className="text-[11px] pl-6 mt-1" style={{ color: "#A3ABB8" }}>ยังไม่มีหน่วยงานในฝ่ายนี้ (กด "+ หน่วยงาน" เพื่อเพิ่ม)</div>}
                  <div className="flex flex-wrap gap-x-4 gap-y-1 pl-6 mt-1">
                    {list.map((d) => (
                      <span key={d} className="inline-flex items-center gap-1">
                        <label className="flex items-center gap-1.5 text-xs cursor-pointer" style={{ color: INK }}>
                          <input type="checkbox" checked={f.departments.includes(d)} onChange={() => toggleDept(d)} /> {d}
                        </label>
                        <TinyIcon icon={Pencil} label={`แก้ไขหน่วยงาน ${d}`} color="#6A4BC9" onClick={() => setDeptPanel({ type: "edit", name: d })} />
                        <TinyIcon icon={Trash2} label={`ลบหน่วยงาน ${d}`} color={RED} onClick={() => setDeptPanel({ type: "delete", name: d })} />
                      </span>
                    ))}
                  </div>
                  {deptPanel && deptPanel.type === "edit" && list.includes(deptPanel.name) && (
                    <div className="mt-2 pl-6"><DeptEditForm key={deptPanel.name} name={deptPanel.name} ops={deptOps} onDone={(r) => onDeptRenamed(deptPanel.name, r)} onCancel={() => setDeptPanel(null)} /></div>
                  )}
                  {deptPanel && deptPanel.type === "delete" && list.includes(deptPanel.name) && (
                    <div className="mt-2 pl-6"><DeptDeleteConfirm key={deptPanel.name} name={deptPanel.name} ops={deptOps} onDone={(r) => onDeptRemoved(deptPanel.name, r)} onCancel={() => setDeptPanel(null)} /></div>
                  )}
                </div>
              ))}
            </div>
            <div className="text-[11px]" style={{ color: "#8A93A3" }}>ไอคอนดินสอ / ถังขยะข้างชื่อฝ่ายหรือหน่วยงาน = แก้ไข / ลบในระบบ มีผลกับทุกบัญชีและทุกหน้า (ไม่ใช่เฉพาะบัญชีนี้) · ฝ่ายใช้จัดกลุ่มหน่วยงานในแดชบอร์ด</div>
            <DeptResultNote key={String(deptNote)} saved={deptNote} ops={deptOps} onClear={() => setDeptNote(null)} />
          </div>
        )}
        {(f.role === "admin" || f.role === "director") && (
          <div className="text-xs rounded-lg px-3 py-2" style={{ background: "#FDF3E1", color: "#8A5B14" }}>บทบาทนี้เห็นและจัดการได้ทุกหน่วยงาน ไม่ต้องเลือกหน่วยงาน {f.role === "admin" ? "(ผู้ใช้ Admin มีสิทธิ์สูง ควรให้เฉพาะเจ้าหน้าที่ฝ่ายบุคคล)" : ""}</div>
        )}
        <div className="grid grid-cols-2 gap-4">
          <Field label={isNew ? "รหัสผ่าน" : "ตั้งรหัสผ่านใหม่ (เว้นว่างถ้าไม่เปลี่ยน)"} required={isNew}>
            <div className="flex gap-2">
              <input autoFocus={!!focusPw} type={showPw ? "text" : "password"} value={f.pw} onChange={(e) => set({ pw: e.target.value })} className={inputCls + " font-mono"} style={inputStyle} autoComplete="new-password" placeholder={isNew ? "" : "ไม่เปลี่ยนรหัสผ่าน"} />
              <button type="button" onClick={() => setShowPw((v) => !v)} className="px-2.5 rounded-lg border text-xs whitespace-nowrap" style={{ borderColor: "#D8DDE5", color: SUB }}>{showPw ? "ซ่อน" : "แสดง"}</button>
              <button type="button" onClick={() => { set({ pw: randomPassword() }); setShowPw(true); }} className="px-2.5 rounded-lg border text-xs whitespace-nowrap" style={{ borderColor: NAVY, color: NAVY }}>สุ่มรหัส</button>
            </div>
            {strength && <div className="text-[11px] mt-1" style={{ color: strength === "ใช้ได้" ? GREEN : "#8A5B14" }}>{strength}</div>}
          </Field>
          <Field label="สถานะบัญชี">
            <label className="flex items-center gap-2 text-sm pt-2" style={{ color: INK, opacity: isJudge || isSelf ? 0.6 : 1 }}>
              <input type="checkbox" checked={f.active} disabled={isJudge || isSelf} onChange={(e) => set({ active: e.target.checked })} /> เปิดใช้งาน (ถ้าปิด จะเข้าสู่ระบบไม่ได้)
            </label>
            {isSelf && <div className="text-[11px]" style={{ color: SUB }}>ปิดบัญชีของตัวเองหรือเปลี่ยนบทบาทตัวเองไม่ได้</div>}
          </Field>
        </div>
        <div className="text-[11px]" style={{ color: SUB }}>ระบบเก็บรหัสผ่านแบบเข้ารหัส ไม่สามารถเปิดดูย้อนหลังได้ ถ้าลืม ให้กด "ตั้งรหัสใหม่" แล้วแจ้งรหัสใหม่ให้ผู้ใช้ จดรหัสที่ตั้งไว้ก่อนกดบันทึก</div>
        {errors.length > 0 && (
          <div className="rounded-lg px-3 py-2 text-xs space-y-0.5" style={{ background: "#FBEAE8", color: "#8B2E24" }}>
            {errors.map((e, i) => <div key={i} className="flex gap-1.5"><AlertTriangle size={13} className="mt-0.5 flex-shrink-0" /> {e}</div>)}
          </div>
        )}
        <div className="flex justify-end gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm border" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
          <button onClick={submit} disabled={saving} className="px-4 py-2 rounded-lg text-sm font-semibold text-white flex items-center gap-1.5 disabled:opacity-60" style={{ background: GREEN }}>
            <CheckCircle2 size={16} /> {saving ? "กำลังบันทึก..." : isNew ? "เพิ่มผู้ใช้" : "บันทึกการแก้ไข"}
          </button>
        </div>
      </div>
    </Modal>
  );
}

function MiniBtn({ icon: Icon, label, onClick, color }) {
  return (
    <button type="button" title={label} aria-label={label} onClick={onClick} className="p-1.5 rounded-lg border bg-white hover:bg-slate-50" style={{ borderColor: "#E2E6EB", color }}>
      <Icon size={15} />
    </button>
  );
}

function UserManager({ users, kaizens, departments, deptOps, currentUser, onSave, syncMode }) {
  const [query, setQuery] = useState("");
  const [roleFilter, setRoleFilter] = useState("all");
  const [editor, setEditor] = useState(null); // { mode: "new" | "edit", user }
  const [deleting, setDeleting] = useState(null);
  const [notice, setNotice] = useState(null); // { type, text, secret? }
  const [unsaved, setUnsaved] = useState(false);
  const [busy, setBusy] = useState(false);
  const [copied, setCopied] = useState(false);
  const activeAdmins = users.filter((u) => u.role === "admin" && u.active !== false);
  const dupNames = useMemo(() => {
    const c = {};
    users.forEach((u) => { c[normName(u.name)] = (c[normName(u.name)] || 0) + 1; });
    return c;
  }, [users]);
  const shown = useMemo(() => {
    const q = query.trim().toLowerCase();
    return users
      .filter((u) => (roleFilter === "all" || u.role === roleFilter) && (!q || u.name.toLowerCase().includes(q) || u.departments.some((d) => d.toLowerCase().includes(q)) || (u.aliases || []).some((a) => a.toLowerCase().includes(q))))
      .sort((a, b) => USER_ROLE_ORDER.indexOf(a.role) - USER_ROLE_ORDER.indexOf(b.role));
  }, [users, query, roleFilter]);
  const activeCount = users.filter((u) => u.active !== false).length;

  const persist = async (list, okText, secret) => {
    setBusy(true);
    setCopied(false);
    const r = await onSave(list);
    setBusy(false);
    if (r === "error") {
      setUnsaved(true);
      setNotice({ type: "error", text: "ใช้งานได้ในหน้านี้แล้ว แต่บันทึกลง Google Sheet ไม่สำเร็จ ตรวจสอบอินเทอร์เน็ตแล้วกด \"ลองบันทึกอีกครั้ง\"", secret });
      return;
    }
    setUnsaved(false);
    setNotice({ type: "ok", text: r === "sheet" ? `${okText} · บันทึกลง Google Sheet แล้ว` : `${okText} · โหมดทดลอง: เก็บไว้เฉพาะในหน้านี้ จะหายเมื่อรีเฟรช ติดตั้ง Users_Addon.gs เพื่อบันทึกถาวร`, secret });
  };
  const submitUser = async (user, plain, isNew) => {
    const list = isNew ? [...users, user] : users.map((u) => (u.id === user.id ? user : u));
    setEditor(null);
    await persist(list, isNew ? `เพิ่มผู้ใช้ "${user.name}" แล้ว` : `แก้ไขผู้ใช้ "${user.name}" แล้ว`, plain ? { name: user.name, password: plain } : null);
  };
  const toggleActive = (u) => persist(users.map((x) => (x.id === u.id ? { ...x, active: x.active === false } : x)), u.active === false ? `เปิดใช้งาน "${u.name}" แล้ว` : `ปิดใช้งาน "${u.name}" แล้ว (เข้าสู่ระบบไม่ได้)`, null);
  const confirmDelete = async () => {
    const u = deleting;
    setDeleting(null);
    await persist(users.filter((x) => x.id !== u.id), `ลบผู้ใช้ "${u.name}" แล้ว`, null);
  };
  const lockedReason = (u, forDelete) => {
    if (u.id === currentUser.id) return "บัญชีของท่าน";
    if (u.role === "judge") return "บัญชีกรรมการ";
    if (u.role === "admin" && u.active !== false && activeAdmins.length <= 1) return "Admin คนสุดท้าย";
    return "";
  };
  const exportUsers = () => {
    const rows = users.map((u) => ({
      ชื่อผู้ใช้: u.name, บทบาท: ROLE_LABEL[u.role],
      ฝ่าย: u.role === "admin" || u.role === "director" ? "ทุกฝ่าย" : Array.from(new Set(u.departments.map((d) => SECTION_MAP[d]).filter(Boolean))).join(", ") || "-",
      หน่วยงาน: u.role === "admin" || u.role === "director" ? "ทุกหน่วยงาน" : u.departments.join(", ") || "-",
      สถานะ: u.active === false ? "ปิดใช้งาน" : "ใช้งาน", ชื่อเดิม: (u.aliases || []).join(", "),
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "ผู้ใช้");
    XLSX.writeFile(wb, `Kaizen_Users_${todayISO()}.xlsx`);
  };
  const copySecret = async () => {
    try {
      await navigator.clipboard.writeText(notice.secret.password);
      setCopied(true);
    } catch (e) {
      setCopied(false);
    }
  };

  const delCount = (u) => {
    const names = [u.name, ...(u.aliases || [])];
    return {
      submitted: kaizens.filter((k) => names.includes(k.submittedBy)).length,
      pending: u.role === "supervisor" ? kaizens.filter((k) => k.status === "pending_supervisor" && u.departments.includes(k.department)).length : 0,
    };
  };

  return (
    <div className="space-y-4">
      <div className="rounded-xl p-4 md:p-5 flex items-center justify-between gap-3 flex-wrap" style={{ background: `linear-gradient(135deg, ${NAVY_DEEP}, ${NAVY})` }}>
        <div className="flex items-center gap-3">
          <div className="w-11 h-11 rounded-xl flex items-center justify-center" style={{ background: GOLD }}><Users size={22} color={NAVY_DEEP} /></div>
          <div>
            <div className="text-white text-lg font-semibold">จัดการผู้ใช้</div>
            <div className="text-xs" style={{ color: "#AFC0D6" }}>ทั้งหมด {users.length} บัญชี · ใช้งาน {activeCount} · ปิดใช้งาน {users.length - activeCount}</div>
          </div>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-[11px] px-2.5 py-1 rounded-full" style={syncMode === "sheet" ? { background: "#E9F5EF", color: GREEN } : { background: "#FDF3E1", color: "#8A5B14" }}>
            {syncMode === "sheet" ? "บันทึกลง Google Sheet" : "โหมดทดลอง: หายเมื่อรีเฟรช"}
          </span>
          <button onClick={exportUsers} className="text-xs px-3 py-1.5 rounded-lg border flex items-center gap-1.5 text-white" style={{ borderColor: "#5D7699" }}><FileSpreadsheet size={14} /> Export (ไม่รวมรหัสผ่าน)</button>
          <button onClick={() => setEditor({ mode: "new" })} className="px-3.5 py-2 rounded-lg text-sm font-semibold flex items-center gap-1.5" style={{ background: GOLD, color: NAVY_DEEP }}><Plus size={16} /> เพิ่มผู้ใช้</button>
        </div>
      </div>

      {notice && (
        <div className="rounded-xl border p-3.5 space-y-2" style={notice.type === "ok" ? { background: "#E9F5EF", borderColor: "#A9D6BF" } : { background: "#FBEAE8", borderColor: "#F0C3BD" }}>
          <div className="flex items-start justify-between gap-3">
            <div className="text-sm" style={{ color: notice.type === "ok" ? "#1E5A40" : "#8B2E24" }}>{notice.text}</div>
            <button onClick={() => setNotice(null)} className="p-0.5"><X size={15} style={{ color: SUB }} /></button>
          </div>
          {notice.secret && (
            <div className="rounded-lg bg-white border p-3 flex items-center justify-between gap-3 flex-wrap" style={{ borderColor: "#D8DDE5" }}>
              <div className="text-sm" style={{ color: INK }}>
                รหัสผ่านของ <strong>{notice.secret.name}</strong> คือ <span className="font-mono font-bold px-2 py-0.5 rounded" style={{ background: "#F4F6F8" }}>{notice.secret.password}</span>
                <div className="text-[11px]" style={{ color: SUB }}>ระบบไม่แสดงรหัสนี้อีก จดหรือคัดลอกเก็บไว้แล้วแจ้งผู้ใช้เป็นข้อความส่วนตัว</div>
              </div>
              <button onClick={copySecret} className="text-xs px-3 py-1.5 rounded-lg border font-medium" style={{ borderColor: NAVY, color: NAVY }}>{copied ? "คัดลอกแล้ว ✓" : "คัดลอกรหัส"}</button>
            </div>
          )}
        </div>
      )}
      {unsaved && (
        <div className="rounded-xl border p-3 flex items-center justify-between gap-3 flex-wrap" style={{ background: "#FDF3E1", borderColor: "#E8C77E" }}>
          <div className="text-sm" style={{ color: "#6B4A10" }}>มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึกลง Google Sheet (ถ้ารีเฟรชตอนนี้ การเปลี่ยนแปลงจะหาย)</div>
          <button disabled={busy} onClick={() => persist(users, "บันทึกรายชื่อผู้ใช้", null)} className="px-3 py-1.5 rounded-lg text-sm font-semibold text-white disabled:opacity-60" style={{ background: GREEN }}>{busy ? "กำลังบันทึก..." : "ลองบันทึกอีกครั้ง"}</button>
        </div>
      )}

      <div className="bg-white rounded-xl border p-4 space-y-3" style={{ borderColor: "#E3E7ED" }}>
        <div className="flex items-center gap-2 flex-wrap">
          <div className="relative">
            <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: "#AEB6C2" }} />
            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="ค้นหาชื่อผู้ใช้ / หน่วยงาน" className="rounded-lg border pl-8 pr-3 py-1.5 text-sm w-64 outline-none" style={{ borderColor: "#D8DDE5", color: INK }} />
          </div>
          {[["all", "ทั้งหมด"], ...USER_ROLE_ORDER.map((r) => [r, ROLE_LABEL[r].replace(" (ฝ่ายบุคคล)", "").replace("กรรมการตัดสิน Kaizen Award", "กรรมการ")])].map(([key, label]) => (
            <button key={key} onClick={() => setRoleFilter(key)} className="text-xs px-3 py-1.5 rounded-full border font-medium" style={roleFilter === key ? { background: NAVY, borderColor: NAVY, color: "white" } : { borderColor: "#D8DDE5", color: SUB }}>
              {label} ({key === "all" ? users.length : users.filter((u) => u.role === key).length})
            </button>
          ))}
        </div>
        <div className="overflow-x-auto rounded-lg border" style={{ borderColor: "#E2E6EB" }}>
          <table className="w-full text-sm" style={{ minWidth: 820 }}>
            <thead>
              <tr style={{ background: "#FAFBFC" }}>
                {["ชื่อผู้ใช้", "บทบาท", "หน่วยงานที่ดูแล", "สถานะ", "จัดการ"].map((h, i) => (
                  <th key={h} className={`px-4 py-2.5 font-medium ${i === 4 ? "text-right" : "text-left"}`} style={{ color: SUB }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {shown.length === 0 && <tr><td colSpan={5} className="text-center py-8 text-sm" style={{ color: SUB }}>ไม่พบผู้ใช้ตามเงื่อนไข</td></tr>}
              {shown.map((u) => {
                const locked = lockedReason(u);
                const badge = ROLE_BADGE[u.role];
                const deptText = u.role === "admin" || u.role === "director" ? "ทุกหน่วยงาน" : u.departments.length ? u.departments.join(", ") : "-";
                return (
                  <tr key={u.id} className="border-t" style={{ borderColor: "#EEF0F3", opacity: u.active === false ? 0.55 : 1 }}>
                    <td className="px-4 py-2.5">
                      <div className="font-medium" style={{ color: INK }}>{u.name}{u.id === currentUser.id && <span className="ml-1.5 text-[10px] px-1.5 py-0.5 rounded-full" style={{ background: "#EAF3FC", color: "#1D5FC9" }}>บัญชีของท่าน</span>}</div>
                      {(u.aliases || []).length > 0 && <div className="text-[11px]" style={{ color: "#A3ABB8" }}>ชื่อเดิม: {u.aliases.join(", ")}</div>}
                      {dupNames[normName(u.name)] > 1 && <div className="text-[11px]" style={{ color: "#8A5B14" }}>ชื่อซ้ำกับอีกบัญชี (แยกด้วยรหัสผ่าน)</div>}
                    </td>
                    <td className="px-4 py-2.5"><span className="text-[11px] px-2 py-0.5 rounded-full whitespace-nowrap" style={{ background: badge.bg, color: badge.color }}>{ROLE_LABEL[u.role]}</span></td>
                    <td className="px-4 py-2.5 text-xs max-w-[380px]" style={{ color: SUB }} title={deptText}>
                      <div style={{ display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical", overflow: "hidden" }}>{deptText}</div>
                    </td>
                    <td className="px-4 py-2.5">
                      <span className="text-[11px] px-2 py-0.5 rounded-full whitespace-nowrap" style={u.active === false ? { background: "#F1F3F6", color: "#8A93A3" } : { background: "#E9F5EF", color: GREEN }}>{u.active === false ? "ปิดใช้งาน" : "ใช้งาน"}</span>
                    </td>
                    <td className="px-4 py-2.5">
                      <div className="flex justify-end items-center gap-1.5">
                        <MiniBtn icon={Pencil} label="แก้ไข" onClick={() => setEditor({ mode: "edit", user: u })} color="#6A4BC9" />
                        <MiniBtn icon={KeyRound} label="ตั้งรหัสใหม่" onClick={() => setEditor({ mode: "edit", user: u, focusPw: true })} color={NAVY} />
                        {!locked && <MiniBtn icon={u.active === false ? CheckCircle2 : Lock} label={u.active === false ? "เปิดใช้งาน" : "ปิดใช้งาน"} onClick={() => toggleActive(u)} color={u.active === false ? GREEN : "#8A5B14"} />}
                        {!locked && <MiniBtn icon={Trash2} label="ลบ" onClick={() => setDeleting(u)} color={RED} />}
                        {locked && <span className="text-[10px] self-center px-1" style={{ color: "#A3ABB8" }} title={`ปิดใช้งาน/ลบไม่ได้: ${locked}`}>ล็อก: {locked}</span>}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="text-[11px]" style={{ color: SUB }}>
          เปลี่ยนชื่อบัญชีพนักงานได้ ระบบเก็บชื่อเดิมไว้ให้ Kaizen เดิมยังแสดงเหมือนเดิม · ปิดใช้งานแทนการลบ จะเก็บบัญชีไว้และเปิดใช้ใหม่ได้ · Admin ต้องมีอย่างน้อย 1 บัญชีที่ใช้งานได้เสมอ
        </div>
      </div>

      {editor && (
        <UserFormModal
          key={editor.user?.id || "new"}
          mode={editor.mode}
          original={editor.user}
          users={users}
          departments={departments}
          deptOps={deptOps}
          currentUser={currentUser}
          focusPw={editor.focusPw}
          onClose={() => setEditor(null)}
          onSubmit={submitUser}
        />
      )}
      {deleting && (() => {
        const c = delCount(deleting);
        return (
          <Modal title={`ลบผู้ใช้: ${deleting.name}`} onClose={() => setDeleting(null)}>
            <div className="space-y-3">
              <div className="text-sm" style={{ color: INK }}>ต้องการลบบัญชีนี้ถาวรใช่หรือไม่ ผู้ใช้จะเข้าสู่ระบบไม่ได้อีก</div>
              {c.submitted > 0 && <div className="text-xs rounded-lg px-3 py-2" style={{ background: "#FDF3E1", color: "#6B4A10" }}>บัญชีนี้เคยส่ง Kaizen {c.submitted} เรื่อง Kaizen ยังอยู่ในระบบตามเดิม แต่ไม่มีบัญชีนี้ให้เข้ามาดูหรือแก้ไข แนะนำให้ "ปิดใช้งาน" แทน</div>}
              {c.pending > 0 && <div className="text-xs rounded-lg px-3 py-2" style={{ background: "#FBEAE8", color: "#8B2E24" }}>มี Kaizen {c.pending} เรื่องที่รอหัวหน้างานคนนี้อนุมัติ หลังลบจะไม่มีใครอนุมัติเรื่องเหล่านั้นจนกว่าจะเพิ่มหัวหน้างานของหน่วยงานนั้น</div>}
              <div className="flex justify-end gap-2">
                <button onClick={() => setDeleting(null)} className="px-4 py-2 rounded-lg text-sm border" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
                <button onClick={confirmDelete} className="px-4 py-2 rounded-lg text-sm font-semibold text-white flex items-center gap-1.5" style={{ background: RED }}><Trash2 size={15} /> ยืนยันลบ</button>
              </div>
            </div>
          </Modal>
        );
      })()}
    </div>
  );
}

function DeleteKaizenModal({ k, award, onConfirm, onClose }) {
  const [reason, setReason] = useState("");
  const [ack, setAck] = useState(false);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const winner = (award && award.decision && award.decision.winners || []).find((w) => w.id === k.id);
  const inShortlist = ((award && award.shortlist) || []).includes(k.id);
  const photos = asList(k.images?.before).length + asList(k.images?.after).length;
  const files = (k.attachments || []).length;
  const go = async () => {
    if (busy || !ack || winner) return;
    setBusy(true);
    setErr("");
    const r = await onConfirm(k, reason.trim());
    setBusy(false);
    if (!r.ok) setErr(r.error);
  };
  return (
    <Modal title="ลบ Kaizen ออกจากระบบ" onClose={onClose}>
      <div className="space-y-3">
        <div className="rounded-lg border p-3" style={{ borderColor: "#F0C3BD", background: "#FDF4F3" }}>
          <div className="text-[11px]" style={{ color: SUB }}>{k.registeredNo} · {k.date}</div>
          <div className="text-sm font-semibold" style={{ color: INK }}>{k.title}</div>
          <div className="text-xs mt-1 flex items-center gap-2 flex-wrap" style={{ color: SUB }}>
            หน่วยงาน {k.department} · ผู้ส่ง {k.submittedBy} · <StatusBadge status={k.status} />
          </div>
        </div>
        {winner && (
          <div className="text-xs rounded-lg px-3 py-2" style={{ background: "#FBEAE8", color: "#8B2E24" }}>
            เรื่องนี้ได้รับ <strong>{winner.tier}</strong> จาก Kaizen Award ที่ประกาศผลแล้ว จึงลบไม่ได้ ให้ผู้อำนวยการกด "ยกเลิกการประกาศ & เปิดแก้ไข" ที่หน้า Kaizen Award ก่อน
          </div>
        )}
        {inShortlist && !winner && (
          <div className="text-xs rounded-lg px-3 py-2" style={{ background: "#FDF3E1", color: "#6B4A10" }}>
            เรื่องนี้ถูกส่งให้กรรมการตัดสิน Kaizen Award แล้ว การลบจะถอนเรื่องนี้ออกจากรายการของกรรมการด้วย (คะแนนของเรื่องนี้จะไม่ถูกนับอีก)
          </div>
        )}
        <ul className="text-xs space-y-1" style={{ color: INK }}>
          <li>• เรื่องนี้จะหายจากทุกหน้า (Dashboard, หน้าพนักงาน / หัวหน้างาน / ผู้อำนวยการ, Kaizen Award) และตัวเลขสรุปจะลดลงตาม</li>
          <li>• เมื่อเชื่อม Google Sheet ข้อมูลจะถูกย้ายไปเก็บสำรองในชีต "KaizensDeleted" (พร้อมวันเวลา ผู้ลบ เหตุผล) ไม่ได้ลบทิ้งถาวร · รูป {photos} ใบ / ไฟล์แนบ {files} ไฟล์ ยังอยู่ใน Google Drive</li>
          <li>• ระบบไม่ได้แจ้งผู้ส่งเรื่องอัตโนมัติ (การแจ้งเตือนในเว็บอยู่เฉพาะในหน้าจอที่ใช้งานอยู่) ควรแจ้งผู้ส่งเอง</li>
        </ul>
        <Field label="เหตุผลที่ลบ (ไม่บังคับ แต่แนะนำ: บันทึกไว้ในชีตสำรองเพื่อตรวจสอบย้อนหลัง)">
          <TextInput value={reason} onChange={(e) => setReason(e.target.value)} placeholder="เช่น ส่งซ้ำ / ข้อมูลทดสอบ / ส่งผิดหน่วยงาน" />
        </Field>
        <label className="flex items-center gap-2 text-sm cursor-pointer" style={{ color: INK }}>
          <input type="checkbox" checked={ack} onChange={(e) => setAck(e.target.checked)} /> ยืนยันว่าต้องการลบเรื่องนี้ออกจากระบบ
        </label>
        {err && <div className="text-xs rounded-lg px-3 py-2" style={{ background: "#FBEAE8", color: "#8B2E24" }}>{err}</div>}
        <div className="flex justify-end gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm border" style={{ borderColor: "#D8DDE5", color: SUB }}>ยกเลิก</button>
          <button onClick={go} disabled={!ack || !!winner || busy} className="px-4 py-2 rounded-lg text-sm font-semibold text-white flex items-center gap-1.5 disabled:opacity-40" style={{ background: RED }}>
            <Trash2 size={15} /> {busy ? "กำลังลบ..." : "ลบเรื่องนี้"}
          </button>
        </div>
      </div>
    </Modal>
  );
}

function SyncStatusBar() {
  const [snap, setSnap] = useState({ pending: 0, failed: [] });
  const [flash, setFlash] = useState(false);
  const [hidden, setHidden] = useState(false);
  const prevPending = useRef(0);
  useEffect(() => {
    const fn = (x) => { setSnap(x); if (x.failed.length) setHidden(false); };
    syncSubscribers.add(fn);
    return () => syncSubscribers.delete(fn);
  }, []);
  useEffect(() => {
    let timer;
    if (prevPending.current > 0 && snap.pending === 0 && snap.failed.length === 0) {
      setFlash(true);
      timer = setTimeout(() => setFlash(false), 3500);
    }
    prevPending.current = snap.pending;
    return () => clearTimeout(timer);
  }, [snap.pending, snap.failed.length]);
  // ปิดหน้าเว็บระหว่างกำลังบันทึก รูปอาจหาย จึงเตือนก่อน
  useEffect(() => {
    const h = (e) => {
      if (snap.pending > 0 || snap.failed.length > 0) { e.preventDefault(); e.returnValue = ""; }
    };
    window.addEventListener("beforeunload", h);
    return () => window.removeEventListener("beforeunload", h);
  }, [snap.pending, snap.failed.length]);
  if (snap.pending > 0) {
    return (
      <div className="fixed bottom-4 right-4 z-[70] rounded-xl border px-4 py-3 text-sm shadow-xl bg-white" style={{ borderColor: "#B9D2F0", color: INK }}>
        กำลังบันทึกลง Google Sheet ({snap.pending} รายการ) กรุณาอย่าปิดหน้านี้...
        <div className="text-[11px]" style={{ color: SUB }}>ถ้าแนบรูปหลายใบ อาจใช้เวลาครู่หนึ่ง</div>
      </div>
    );
  }
  if (snap.failed.length > 0 && !hidden) {
    return (
      <div className="fixed bottom-4 right-4 z-[70] w-96 max-w-[92vw] rounded-xl border p-4 shadow-xl" style={{ background: "#FBEAE8", borderColor: "#F0C3BD" }}>
        <div className="flex items-start justify-between gap-2">
          <div className="text-sm font-semibold" style={{ color: "#8B2E24" }}>บันทึกลง Google Sheet ไม่สำเร็จ {snap.failed.length} เรื่อง</div>
          <button onClick={() => setHidden(true)} className="p-0.5" title="ซ่อน"><X size={15} style={{ color: SUB }} /></button>
        </div>
        <div className="text-xs mt-1 space-y-1" style={{ color: INK }}>
          {snap.failed.slice(0, 3).map((f) => (
            <div key={f.id}>• {f.title}<div className="text-[11px] break-words" style={{ color: "#8B2E24" }}>สาเหตุ: {f.error}</div></div>
          ))}
          {snap.failed.length > 3 && <div>และอีก {snap.failed.length - 3} เรื่อง</div>}
        </div>
        <div className="text-[11px] mt-2" style={{ color: SUB }}>ข้อมูลยังอยู่ในหน้านี้ แต่ถ้ารีเฟรชตอนนี้จะหาย กดลองใหม่ได้เลย</div>
        <div className="mt-2 flex justify-end">
          <button onClick={retryFailedSyncs} className="px-3 py-1.5 rounded-lg text-sm font-semibold text-white" style={{ background: GREEN }}>ลองบันทึกอีกครั้ง</button>
        </div>
      </div>
    );
  }
  if (flash) {
    return <div className="fixed bottom-4 right-4 z-[70] rounded-xl border px-4 py-2.5 text-sm shadow-xl" style={{ background: "#E9F5EF", borderColor: "#A9D6BF", color: "#1E5A40" }}>บันทึกลง Google Sheet เรียบร้อยแล้ว ✓</div>;
  }
  return null;
}

function LoginScreen({ onLogin, users, ready }) {
  const [name, setName] = useState("");
  const [pass, setPass] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [lockedUntil, setLockedUntil] = useState(0);
  const failsRef = useRef(0);
  const [open, setOpen] = useState(false);
  const passRef = useRef(null);

  const filteredUsers = useMemo(() => {
    const q = name.trim().toLowerCase();
    const base = users.filter((u) => u.active !== false);
    const matched = q ? base.filter((u) => u.name.toLowerCase().includes(q)) : base;
    const seen = new Set();
    return matched
      .filter((u) => {
        const key = normName(u.name);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      })
      .map((u) => ({ ...u, sameName: base.filter((x) => normName(x.name) === normName(u.name)).length }))
      .slice(0, 30);
  }, [name, users]);

  const selectUser = (u) => {
    setName(u.name);
    setOpen(false);
    setError("");
    if (passRef.current) passRef.current.focus();
  };

  const submit = async () => {
    if (busy || !ready) return;
    if (Date.now() < lockedUntil) {
      setError(`กรอกรหัสผิดหลายครั้ง กรุณารอ ${Math.ceil((lockedUntil - Date.now()) / 1000)} วินาทีแล้วลองใหม่`);
      return;
    }
    const q = normName(name);
    const candidates = users.filter((u) => normName(u.name) === q);
    if (!candidates.length || !pass) {
      setError("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง");
      return;
    }
    setBusy(true);
    let found = null;
    let inactive = false;
    try {
      for (const u of candidates) {
        if (await checkPw(u, pass.trim())) {
          if (u.active === false) inactive = true;
          else found = u;
          break;
        }
      }
    } catch (e) {
      setBusy(false);
      setError("เบราว์เซอร์นี้ตรวจรหัสผ่านไม่ได้ กรุณาเปิดเว็บผ่านลิงก์ https");
      return;
    }
    setBusy(false);
    if (found) {
      failsRef.current = 0;
      setError("");
      onLogin(found);
      return;
    }
    failsRef.current += 1;
    if (failsRef.current >= 5) {
      failsRef.current = 0;
      setLockedUntil(Date.now() + 30000);
      setError("กรอกรหัสผิดหลายครั้ง กรุณารอ 30 วินาทีแล้วลองใหม่");
      return;
    }
    setError(inactive ? "บัญชีนี้ถูกปิดการใช้งาน กรุณาติดต่อ Admin" : "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง");
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
                      u.sameName > 1 ? `มี ${u.sameName} บัญชีใช้ชื่อนี้ (แยกด้วยรหัสผ่าน)` :
                      u.departments.length === 0 ? "" :
                      u.departments.length <= 2 ? u.departments.join(", ") :
                      `${u.departments.slice(0, 2).join(", ")} และอีก ${u.departments.length - 2} หน่วยงาน`;
                    return (
                      <div
                        key={u.id}
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
          <Field label="รหัสผ่าน" required>
            <div className="relative">
              <KeyRound size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: "#AEB6C2" }} />
              <input ref={passRef} type="password" value={pass} onChange={(e) => setPass(e.target.value)} onFocus={() => setOpen(false)} onKeyDown={handleKeyDown} className={inputCls} style={{ ...inputStyle, paddingLeft: 32 }} placeholder="รหัสผ่านของท่าน" autoComplete="current-password" />
            </div>
          </Field>
          {error && (
            <div className="flex items-center gap-1.5 text-xs rounded-lg px-3 py-2" style={{ background: "#F9E9E7", color: RED }}>
              <AlertTriangle size={13} /> {error}
            </div>
          )}
          <button type="button" onClick={submit} disabled={busy || !ready} className="w-full py-2.5 rounded-lg text-sm font-medium text-white flex items-center justify-center gap-1.5 disabled:opacity-60" style={{ background: NAVY }}>
            <Lock size={14} /> {!ready ? "กำลังโหลดรายชื่อผู้ใช้..." : busy ? "กำลังตรวจสอบ..." : "เข้าสู่ระบบ"}
          </button>
        </div>
        <div className="w-full text-center text-xs mt-4" style={{ color: "#9FB0C6" }}>
          ลืมรหัสผ่าน หรือขอบัญชีใหม่ ติดต่อฝ่ายทรัพยากรบุคคล
        </div>
      </div>
    </div>
  );
}

export default function KaizenApp() {
  // จำผู้ใช้ไว้ในแท็บเบราว์เซอร์ เมื่อกด Refresh จะไม่เด้งกลับหน้า Login (ปิดแท็บหรือกดออกจากระบบแล้วจะหายไป)
  // รายชื่อผู้ใช้: เริ่มจากรายการเริ่มต้นในโค้ด แล้วโหลดทับจาก Google Sheet (ถ้าติดตั้ง Users_Addon.gs)
  const [users, setUsers] = useState(() => USERS.map((u) => ({ ...u })));
  const [usersReady, setUsersReady] = useState(false);
  const [usersSync, setUsersSync] = useState("local"); // "sheet" | "local"
  const applyUsers = (list) => {
    USERS.splice(0, USERS.length, ...list); // ให้ฟังก์ชันอื่นที่อ้าง USERS ได้รายชื่อล่าสุดด้วย
    setUsers(list);
  };
  useEffect(() => {
    let alive = true;
    const timer = setTimeout(() => alive && setUsersReady(true), 7000);
    fetchUsersFromSheet()
      .then((res) => {
        if (!alive || !res.supported) return;
        setUsersSync("sheet");
        if (Array.isArray(res.users) && res.users.length) {
          const cfg = res.users.find((u) => u && u.id === "__config");
          if (cfg) {
            const deps = applyDeptConfig(cfg);
            if (deps) {
              setDepartments(deps);
              setSectionList([...SECTION_OPTIONS]);
            }
          }
          applyUsers(normalizeUsers(res.users.filter((u) => !(u && u.id === "__config"))));
        }
      })
      .catch(() => {})
      .finally(() => {
        clearTimeout(timer);
        if (alive) setUsersReady(true);
      });
    return () => {
      alive = false;
      clearTimeout(timer);
    };
  }, []);
  // จำผู้ใช้ไว้ในแท็บเบราว์เซอร์ เมื่อกด Refresh จะไม่เด้งกลับหน้า Login (ปิดแท็บหรือกดออกจากระบบแล้วจะหายไป)
  const [currentUserId, setCurrentUserId] = useState(() => {
    try {
      return sessionStorage.getItem("kz_session_user");
    } catch (e) {
      return null;
    }
  });
  const currentUser = useMemo(() => users.find((u) => u.id === currentUserId && u.active !== false) || null, [users, currentUserId]);
  useEffect(() => {
    try {
      if (currentUserId) sessionStorage.setItem("kz_session_user", currentUserId);
      else sessionStorage.removeItem("kz_session_user");
    } catch (e) {}
  }, [currentUserId]);
  // บัญชีที่จำไว้ถูกลบ/ปิดใช้งาน -> ออกจากระบบ
  useEffect(() => {
    if (usersReady && currentUserId && !currentUser) setCurrentUserId(null);
  }, [usersReady, currentUser, currentUserId]);
  const saveUsersList = async (list) => {
    applyUsers(list);
    if (usersSync !== "sheet") return "local";
    const ok = await saveUsersToSheet([...list, buildDeptConfigRow(departments)]);
    return ok ? "sheet" : "error";
  };
  const [kaizens, setKaizens] = useState(() => seedData().map(normalizeKaizen));
  const [sheetStatus, setSheetStatus] = useState("checking"); // "checking" | "connected" | "demo"
  useEffect(() => {
    syncState.tracking = sheetStatus === "connected" || sheetStatus === "connected-empty";
  }, [sheetStatus]);

  useEffect(() => {
    fetchKaizensFromSheet()
      .then((data) => {
        if (Array.isArray(data) && data.length > 0) {
          setKaizens(data.map(normalizeKaizen));
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
  const [, setSectionList] = useState(() => [...SECTION_OPTIONS]); // ใช้กระตุ้นให้หน้าจอวาดใหม่เมื่อรายการฝ่ายเปลี่ยน
  const [tab, setTab] = useStoredState("kz_tab", "dashboard");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [viewing, setViewing] = useState(null);
  const [reviewing, setReviewing] = useState(null);
  const [presenting, setPresenting] = useState(null);
  const [showDeptManager, setShowDeptManager] = useState(false);
  const [adminEditing, setAdminEditing] = useState(null);
  const [deletingKaizen, setDeletingKaizen] = useState(null);
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

  // การแก้ไขในหน้า "1. ตั้งเกณฑ์" เก็บเป็นร่างก่อน จนกว่าจะกดปุ่ม "บันทึกการตั้งค่า" (กันกรรมการเห็นค่าที่ยังแก้ไม่เสร็จ)
  const [awardDraft, setAwardDraft] = useStoredState("kz_award_draft", {});
  const saveAwardSettings = async (next) => {
    setAward(next);
    if (awardSync !== "sheet") return "local";
    clearTimeout(awardTimer.current);
    const ok = await saveAwardToSheet(next);
    return ok ? "sheet" : "error";
  };

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

  const noticeJudges = (message) => activeJudges(award).forEach((j) => pushNotification(j.userName, message, null));

  const judgeSubmitted = (judge) => {
    notifyRole("admin", `${judge.userName} ส่งคะแนน ${award.roundName} เรียบร้อยแล้ว`, null);
    notifyRole("director", `${judge.userName} ส่งคะแนน ${award.roundName} เรียบร้อยแล้ว`, null);
  };

  const announceAward = (winners) => {
    winners.forEach((w) =>
      pushNotification(w.submittedBy, `🏆 ยินดีด้วย! เรื่อง "${w.title}" ได้รับ${w.tier} ใน ${award.roundName}`, w.id)
    );
    notifyRole("admin", `ผู้อำนวยการยืนยันผล ${award.roundName} แล้ว (${winners.length} รางวัล)`, winners[0]?.id);
  };

  const pushNotification = (recipientName, message, kaizenId) => {
    if (!recipientName) return;
    setNotifications((list) => [{ id: uid(), recipientName, message, kaizenId, read: false, date: nowStamp() }, ...list]);
  };
  const findSupervisorFor = (department) => users.find((u) => u.role === "supervisor" && u.active !== false && u.departments.includes(department));
  // แจ้งเตือนทุกบัญชีที่มีบทบาทนั้น (เช่น มี Admin หลายคน)
  const notifyRole = (role, message, kaizenId) => {
    const seen = new Set();
    users.filter((u) => u.role === role && u.active !== false).forEach((u) => {
      if (seen.has(u.name)) return;
      seen.add(u.name);
      pushNotification(u.name, message, kaizenId);
    });
  };

  const updateKaizen = (id, patch) => setKaizens((list) => list.map((k) => (k.id === id ? { ...k, ...patch } : k)));

  // ---------- จัดการหน่วยงาน (เพิ่ม / แก้ไข / ลบ) ----------
  const deptStats = (name) => ({
    kaizens: kaizens.filter((k) => k.department === name).length,
    users: users.filter((u) => (u.departments || []).includes(name)),
  });
  const checkDeptName = (raw, exceptName) => {
    const name = String(raw || "").replace(/\s+/g, " ").trim();
    if (!name) return { error: "กรุณาใส่ชื่อหน่วยงาน" };
    if (name.length > 60) return { error: "ชื่อหน่วยงานยาวเกินไป (ไม่เกิน 60 ตัวอักษร)" };
    if (departments.some((d) => d !== exceptName && d.toLowerCase() === name.toLowerCase())) return { error: `มีหน่วยงาน "${name}" อยู่แล้ว` };
    return { name };
  };
  const persistDepts = async (nextDeps, nextUsers) => {
    setDepartments(nextDeps);
    if (nextUsers) applyUsers(nextUsers);
    if (usersSync !== "sheet") return "local";
    const ok = await saveUsersToSheet([...(nextUsers || users), buildDeptConfigRow(nextDeps)]);
    return ok ? "sheet" : "error";
  };
  const canRemoveDept = (name) => {
    const st = deptStats(name);
    if (departments.length <= 1) return { ok: false, error: "ต้องมีหน่วยงานอย่างน้อย 1 หน่วยงาน" };
    if (st.kaizens > 0) return { ok: false, error: `ลบไม่ได้: มี Kaizen ${st.kaizens} เรื่องอยู่ในหน่วยงานนี้ (ถ้าต้องการเปลี่ยนชื่อ ใช้ปุ่มแก้ไข)` };
    const stranded = st.users.filter((u) => (u.role === "employee" || u.role === "supervisor") && (u.departments || []).filter((d) => d !== name).length === 0);
    if (stranded.length) return { ok: false, error: `ลบไม่ได้: บัญชี ${stranded.map((u) => u.name).join(", ")} จะไม่เหลือหน่วยงานที่ดูแล ให้ย้ายหน่วยงานของบัญชีเหล่านั้นก่อน` };
    return { ok: true };
  };
  const checkSectionName = (raw, exceptName) => {
    const name = String(raw || "").replace(/\s+/g, " ").trim();
    if (!name) return { error: "กรุณาใส่ชื่อฝ่าย" };
    if (name.length > 60) return { error: "ชื่อฝ่ายยาวเกินไป (ไม่เกิน 60 ตัวอักษร)" };
    if (name === "อื่นๆ") return { error: 'ชื่อ "อื่นๆ" ใช้แทนหน่วยงานที่ยังไม่มีฝ่าย กรุณาใช้ชื่ออื่น' };
    if (SECTION_OPTIONS.some((x) => x !== exceptName && x.toLowerCase() === name.toLowerCase())) return { error: `มีฝ่าย "${name}" อยู่แล้ว` };
    return { name };
  };
  const sectionOps = {
    list: () => [...SECTION_OPTIONS],
    stats: (name) => ({ departments: departments.filter((d) => SECTION_MAP[d] === name) }),
    add: async (raw) => {
      const chk = checkSectionName(raw);
      if (chk.error) return { ok: false, error: chk.error };
      SECTION_OPTIONS.push(chk.name);
      setSectionList([...SECTION_OPTIONS]);
      const saved = await persistDepts(departments, null);
      return { ok: true, name: chk.name, saved };
    },
    rename: async (oldName, raw) => {
      const chk = checkSectionName(raw, oldName);
      if (chk.error) return { ok: false, error: chk.error };
      const idx = SECTION_OPTIONS.indexOf(oldName);
      if (idx < 0) return { ok: false, error: "ไม่พบฝ่ายนี้" };
      const n = departments.filter((d) => SECTION_MAP[d] === oldName).length;
      SECTION_OPTIONS[idx] = chk.name;
      Object.keys(SECTION_MAP).forEach((d) => { if (SECTION_MAP[d] === oldName) SECTION_MAP[d] = chk.name; });
      setSectionList([...SECTION_OPTIONS]);
      const saved = await persistDepts(departments, null);
      return { ok: true, name: chk.name, saved, departments: n };
    },
    canRemove: (name) => {
      const inSec = departments.filter((d) => SECTION_MAP[d] === name);
      if (SECTION_OPTIONS.length <= 1) return { ok: false, error: "ต้องมีฝ่ายอย่างน้อย 1 ฝ่าย" };
      if (inSec.length) return { ok: false, error: `ลบไม่ได้: ยังมี ${inSec.length} หน่วยงานอยู่ในฝ่ายนี้ (${inSec.slice(0, 4).join(", ")}${inSec.length > 4 ? " ..." : ""}) ให้ย้ายไปฝ่ายอื่นก่อน (ปุ่มแก้ไขหน่วยงาน แล้วเลือกฝ่ายใหม่)` };
      return { ok: true };
    },
    remove: async (name) => {
      const can = sectionOps.canRemove(name);
      if (!can.ok) return can;
      SECTION_OPTIONS.splice(SECTION_OPTIONS.indexOf(name), 1);
      setSectionList([...SECTION_OPTIONS]);
      const saved = await persistDepts(departments, null);
      return { ok: true, saved };
    },
  };
  const deptOps = {
    sections: sectionOps,
    stats: deptStats,
    canRemove: canRemoveDept,
    add: async (rawName, section) => {
      const chk = checkDeptName(rawName);
      if (chk.error) return { ok: false, error: chk.error };
      SECTION_MAP[chk.name] = section && SECTION_OPTIONS.includes(section) ? section : SECTION_OPTIONS[0];
      const saved = await persistDepts([...departments, chk.name], null);
      return { ok: true, name: chk.name, saved };
    },
    rename: async (oldName, rawNew, section) => {
      const chk = checkDeptName(rawNew, oldName);
      if (chk.error) return { ok: false, error: chk.error };
      const newName = chk.name;
      const changed = newName !== oldName;
      const st = deptStats(oldName);
      const affected = kaizens.filter((k) => k.department === oldName);
      const sec = section && SECTION_OPTIONS.includes(section) ? section : (SECTION_MAP[oldName] || SECTION_OPTIONS[0]);
      if (changed) delete SECTION_MAP[oldName];
      SECTION_MAP[newName] = sec;
      const nextDeps = departments.map((d) => (d === oldName ? newName : d));
      const nextUsers = changed
        ? users.map((u) => ((u.departments || []).includes(oldName) ? { ...u, departments: u.departments.map((d) => (d === oldName ? newName : d)) } : u))
        : users;
      if (changed && affected.length) {
        // แก้ชื่อหน่วยงานใน Kaizen เดิมทุกเรื่อง (อัปเดตลง Google Sheet ทีละเรื่อง ดูสถานะที่มุมขวาล่าง)
        setKaizens((list) => list.map((k) => (k.department === oldName ? { ...k, department: newName } : k)));
        affected.forEach((k) => syncKaizenToSheet("update", { ...k, department: newName }));
      }
      const saved = await persistDepts(nextDeps, nextUsers);
      return { ok: true, name: newName, saved, kaizens: changed ? affected.length : 0, users: changed ? st.users.length : 0 };
    },
    remove: async (name) => {
      const can = canRemoveDept(name);
      if (!can.ok) return can;
      const st = deptStats(name);
      delete SECTION_MAP[name];
      const nextUsers = users.map((u) => ((u.departments || []).includes(name) ? { ...u, departments: u.departments.filter((d) => d !== name) } : u));
      const saved = await persistDepts(departments.filter((d) => d !== name), nextUsers);
      return { ok: true, saved, users: st.users.length };
    },
    resave: () => persistDepts(departments, null),
  };

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
        activityType: formData.activityType || "Kaizen",
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
  // Admin ลบ Kaizen: เมื่อเชื่อม Google Sheet จะรอให้เซิร์ฟเวอร์ยืนยันก่อนจึงลบออกจากหน้าจอ (ถ้าล้มเหลว เรื่องยังอยู่และแจ้งสาเหตุ)
  const deleteKaizenRecord = async (k, reason) => {
    if (!currentUser || currentUser.role !== "admin") return { ok: false, error: "เฉพาะ Admin เท่านั้นที่ลบ Kaizen ได้" };
    if (((award && award.decision && award.decision.winners) || []).some((w) => w.id === k.id)) {
      return { ok: false, error: "เรื่องนี้ได้รับรางวัล Kaizen Award ที่ประกาศผลแล้ว ลบไม่ได้ (ให้ผู้อำนวยการยกเลิกการประกาศก่อน)" };
    }
    if (syncState.tracking) {
      const r = await postToSheet("delete", { id: k.id, registeredNo: k.registeredNo, deletedBy: currentUser.name, reason });
      if (!r.ok) {
        const old = /ไม่รู้จัก action/.test(r.error || "");
        return { ok: false, error: old ? "Code.gs ที่ติดตั้งอยู่ยังเป็นเวอร์ชันเก่า (ยังไม่รองรับการลบ) ให้อัปเดต Code.gs เป็นฉบับล่าสุดแล้ว Deploy เวอร์ชันใหม่ก่อน" : `ลบไม่สำเร็จ: ${r.error || "ไม่ทราบสาเหตุ"}` };
      }
    }
    setKaizens((list) => list.filter((x) => x.id !== k.id));
    if (((award && award.shortlist) || []).includes(k.id)) changeAward({ ...award, shortlist: award.shortlist.filter((id) => id !== k.id) });
    setDeletingKaizen(null);
    setViewing(null);
    return { ok: true };
  };

  const adminSave = (formData, note) => {
    const original = adminEditing;
    if (!original || !canEditKaizen(currentUser, original)) return;
    const changed = Object.keys(EDIT_LABELS).filter((key) => JSON.stringify(original[key] ?? "") !== JSON.stringify(formData[key] ?? ""));
    const photoSide = (side, name) => {
      const before = asList(original.images?.[side]);
      const after = asList(formData.images?.[side]);
      if (JSON.stringify(before) === JSON.stringify(after)) return;
      changed.push(after.length === 0 ? `ลบ${name}` : after.length > before.length ? `${name} (เพิ่มเป็น ${after.length} รูป)` : `${name} (${after.length} รูป)`);
    };
    photoSide("before", "ภาพก่อนปรับปรุง");
    photoSide("after", "ภาพหลังปรับปรุง");
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
    if (original.submittedBy && original.submittedBy !== currentUser.name) {
      pushNotification(original.submittedBy, `${currentUser.role === "supervisor" ? "หัวหน้างาน" : "Admin"}แก้ไขข้อมูลเรื่อง "${original.title}" (${changed.length ? changed.map((c) => EDIT_LABELS[c] || c).join(", ") : "หมายเหตุ"})`, original.id);
    }
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
        notifyRole("admin", `หัวหน้างานอนุมัติแล้ว รอ Admin ตรวจสอบ: "${k.title}"`, k.id);
        pushNotification(k.submittedBy, `เรื่อง "${k.title}" ผ่านการอนุมัติจากหัวหน้างานแล้ว กำลังรอ Admin ตรวจสอบ`, k.id);
      } else if (k.status === "pending_admin") {
        notifyRole("director", `Admin ตรวจสอบแล้ว รออนุมัติขั้นสุดท้าย: "${k.title}"`, k.id);
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
        notifyRole("admin", `เรื่อง "${k.title}" ถูกส่งกลับมาให้ท่านทบทวนอีกครั้ง`, k.id);
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
    { key: "users", label: "จัดการผู้ใช้", icon: Users, roles: ["admin"] },
    { key: "judge", label: "ให้คะแนน Kaizen Award", icon: Medal, roles: ["judge"] },
  ];
  const NAV = currentUser ? ALL_NAV.filter((n) => n.roles.includes(currentUser.role)) : ALL_NAV;
  // ถ้าหน้าที่จำไว้ไม่ตรงกับสิทธิ์ของผู้ใช้ ให้พาไปหน้าตามสิทธิ์ของตนเอง
  useEffect(() => {
    if (currentUser && !NAV.some((n) => n.key === tab)) setTab(currentUser.role);
  }, [currentUser, tab]);

  if (!currentUser) {
    if (currentUserId && !usersReady) {
      return (
        <div className="min-h-screen w-full flex items-center justify-center" style={{ background: NAVY_DEEP, color: "#9FB0C6", fontFamily: "Tahoma, 'Segoe UI', system-ui, sans-serif" }}>
          กำลังโหลดข้อมูลผู้ใช้...
        </div>
      );
    }
    return (
      <LoginScreen
        users={users}
        ready={usersReady}
        onLogin={(user) => {
          setCurrentUserId(user.id);
          setTab(user.role); // ล็อกอินแล้วพาไปหน้าตามสิทธิ์ของตนเองก่อน
        }}
      />
    );
  }

  // ชื่อที่ใช้ระบุ "ของฉัน": ชื่อบัญชีปัจจุบัน + ชื่อเดิมที่เคยใช้ (กรณี Admin เปลี่ยนชื่อบัญชี)
  const myNames = [currentUser.name, ...(currentUser.aliases || [])];
  const myKaizens = kaizens.filter((k) => myNames.includes(k.submittedBy));
  const myNotifications = notifications.filter((n) => myNames.includes(n.recipientName));
  const myUnreadCount = myNotifications.filter((n) => !n.read).length;
  const supervisorQueue = kaizens.filter((k) => k.status === "pending_supervisor" && currentUser.departments.includes(k.department));
  const adminQueue = kaizens.filter((k) => k.status === "pending_admin");
  const directorQueue = kaizens.filter((k) => k.status === "pending_director");

  return (
    <AwardCtx.Provider value={award}>
    <SyncStatusBar />
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
                  onMarkAllRead={() => setNotifications((list) => list.map((n) => (myNames.includes(n.recipientName) ? { ...n, read: true } : n)))}
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
              onClick={() => { setCurrentUserId(null); setTab("dashboard"); setAwardDraft({}); }}
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
                  <IconBtn icon={Presentation} label="นำเสนอ" onClick={() => setPresenting(k)} color={GOLD} />
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
                  <IconBtn icon={Pencil} label="แก้ไขข้อมูล" onClick={() => setAdminEditing(k)} color="#6A4BC9" />
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
                    <IconBtn icon={Pencil} label="แก้ไขข้อมูล" onClick={() => setAdminEditing(k)} color="#6A4BC9" />
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
                <DepartmentManager departments={departments} ops={deptOps} />
              </div>
            )}
            <Table
              rows={adminQueue}
              actions={(k) => (
                <>
                  <IconBtn icon={Presentation} label="นำเสนอ" onClick={() => setPresenting(k)} color={NAVY} />
                  <IconBtn icon={Pencil} label="แก้ไขข้อมูล / เพิ่มรูป" onClick={() => setAdminEditing(k)} color="#6A4BC9" />
                  <IconBtn icon={Trash2} label="ลบ" onClick={() => setDeletingKaizen(k)} color={RED} />
                  <IconBtn icon={ClipboardList} label="ตรวจ/จัดหมวด" onClick={() => setReviewing(k)} color={GREEN} filled />
                </>
              )}
            />
            <div className="mt-6 flex items-center justify-between gap-3 flex-wrap">
              <div className="text-sm font-medium" style={{ color: INK }}>รายการทั้งหมดในระบบ</div>
            </div>
            <div className="mt-2 flex flex-wrap gap-1.5">
              {[
                ["all", `ทั้งหมด (${kaizens.length})`],
                ["noimage", `ยังไม่มีรูปก่อน-หลัง (${kaizens.filter((k) => asList(k.images?.before).length === 0 || asList(k.images?.after).length === 0).length})`],
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
                  if (adminListFilter === "noimage" && asList(k.images?.before).length > 0 && asList(k.images?.after).length > 0) return false;
                  if (adminListFilter === "incomplete" && missingInfo(k).length === 0) return false;
                  return true;
                })}
                actions={(k) => (
                  <>
                    <IconBtn icon={Eye} label="ดู" onClick={() => setViewing(k)} color={NAVY} />
                    <IconBtn icon={Pencil} label="แก้ไขข้อมูล / เพิ่มรูป" onClick={() => setAdminEditing(k)} color="#6A4BC9" />
                    <IconBtn icon={Trash2} label="ลบ" onClick={() => setDeletingKaizen(k)} color={RED} />
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
        {tab === "users" && currentUser.role === "admin" && (
          <UserManager
            users={users}
            kaizens={kaizens}
            departments={departments}
            deptOps={deptOps}
            currentUser={currentUser}
            onSave={saveUsersList}
            syncMode={usersSync}
          />
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
            onJudgesNotice={noticeJudges}
            draft={awardDraft}
            onDraftChange={setAwardDraft}
            onSaveSettings={saveAwardSettings}
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
            canEditKaizen(currentUser, viewing) && (
              <div className="flex items-center gap-2">
                <button
                  onClick={() => { setAdminEditing(viewing); setViewing(null); }}
                  className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-white"
                  style={{ background: "#6A4BC9" }}
                >
                  <Pencil size={13} /> แก้ไขข้อมูล / เพิ่มรูป
                </button>
                {currentUser.role === "admin" && (
                  <button
                    onClick={() => setDeletingKaizen(viewing)}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium border bg-white"
                    style={{ borderColor: RED, color: RED }}
                  >
                    <Trash2 size={13} /> ลบเรื่องนี้
                  </button>
                )}
              </div>
            )
          }
        >
          <DetailView k={viewing} />
        </Modal>
      )}

      {deletingKaizen && currentUser.role === "admin" && (
        <DeleteKaizenModal key={deletingKaizen.id} k={deletingKaizen} award={award} onConfirm={deleteKaizenRecord} onClose={() => setDeletingKaizen(null)} />
      )}

      {adminEditing && (
        <Modal title={`แก้ไขข้อมูล Kaizen (${currentUser.role === "admin" ? "Admin" : "หัวหน้างาน"})`} onClose={() => setAdminEditing(null)} wide>
          <KaizenForm
            key={adminEditing.id}
            initial={adminEditing}
            editorLabel={currentUser.role === "admin" ? "Admin" : "หัวหน้างาน"}
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
    </AwardCtx.Provider>
  );
}
