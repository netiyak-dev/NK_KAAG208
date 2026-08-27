/**
 * *** เลิกใช้งานแล้ว (OBSOLETE / UNUSED AT RUNTIME) ***
 * ไฟล์นี้เก็บไว้เป็นข้อมูลอ้างอิงเท่านั้น ระบบที่ใช้งานจริงตอนนี้คือ
 * "cloudflare_worker_proxy.js" (Cloudflare Worker + D1) ซึ่งพอร์ตตรรกะทั้งหมด
 * ในไฟล์นี้ไปใช้สูตร/เงื่อนไขเดียวกันทุกประการ — ดูคู่มือ
 * "KAAG208_คู่มือติดตั้งฐานข้อมูล.md" สำหรับวิธีติดตั้งปัจจุบัน
 *
 * KAAG208 & 304 — Genetics for Agriculture
 * Backend สำหรับเว็บบทเรียนออนไลน์: บันทึกการเช็คชื่อเข้าเรียน ความคืบหน้าแบบฝึกหัด
 * คะแนนแบบทดสอบก่อนเรียน/หลังเรียน และคำนวณ Learning Analytics แบบพยากรณ์
 * (predictive risk score + คำแนะนำทบทวนเฉพาะบุคคล)
 *
 * วิธีติดตั้ง (เดิม ไม่ใช้แล้ว): ดูไฟล์คู่มือ "KAAG208_คู่มือติดตั้งฐานข้อมูล.md" ที่แนบมาพร้อมกัน
 *
 * โครงสร้างข้อมูล: สคริปต์นี้จะสร้าง 3 ชีตให้อัตโนมัติในสเปรดชีตที่ผูกสคริปต์นี้ไว้
 *   - "Log"     : Timestamp | StudentID | Email | Name | Session | EventType (checkin / complete)
 *   - "PrePost" : Timestamp | StudentID | Email | Name | Session | Phase (pre / post) | Score | Total
 *   - "QuizLog" : Timestamp | StudentID | Session | Topic | QuizID | Attempt | Correct | IsFirstAttempt
 *
 * ตั้งค่ารหัสผ่านอาจารย์ (ใช้เปิดดูภาพรวมความเสี่ยงของนักศึกษาทุกคน) ที่ค่าคงที่ INSTRUCTOR_KEY ด้านล่าง
 * แล้วนำรหัสเดียวกันไปกรอกในหน้า Dashboard สำหรับอาจารย์
 */

const SHEET_NAME = "Log";
const PREPOST_SHEET_NAME = "PrePost";
const QUIZ_SHEET_NAME = "QuizLog";
const INSTRUCTOR_KEY = "CHANGE_ME_BEFORE_DEPLOY"; // *** ตั้งรหัสผ่านของคุณเองตรงนี้ในตัว Apps Script editor ที่ deploy จริงเท่านั้น อย่า commit ค่าจริงลง git ***

/* เว็บฝั่งหน้าบ้านเรียก API แบบ JSONP (แท็ก <script>) แทน fetch() ตรง ๆ
 * เพราะการเรียก fetch() ข้ามโดเมนไปยัง Apps Script /exec มักเจอปัญหา CORS/redirect
 * ไม่เสถียรจากฝั่งเบราว์เซอร์ (แม้ execution ฝั่งเซิร์ฟเวอร์จะสำเร็จก็ตาม)
 * ตัวแปรนี้เก็บชื่อ callback ของ request ปัจจุบันไว้ให้ jsonOut_ ใช้ห่อผลลัพธ์ */
let CURRENT_CALLBACK = null;

function doGet(e) {
  CURRENT_CALLBACK = e.parameter.callback || null;
  try {
    const action = e.parameter.action;
    const sheet = getSheet_();

    if (action === "checkin") {
      return logEvent_(sheet, e.parameter, "checkin");
    }
    if (action === "complete") {
      return logEvent_(sheet, e.parameter, "complete");
    }
    if (action === "status") {
      return getStatus_(sheet, e.parameter.sid);
    }
    if (action === "prepost") {
      return logPrePost_(e.parameter);
    }
    if (action === "logAnswer") {
      return logAnswer_(getQuizSheet_(), e.parameter);
    }
    if (action === "insights") {
      return getInsights_(sheet, getQuizSheet_(), e.parameter.sid);
    }
    if (action === "riskAll") {
      return getRiskAll_(sheet, getQuizSheet_(), e.parameter.key);
    }
    return jsonOut_({ ok: false, error: "unknown action" });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/* ================= SHEET SETUP ================= */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Timestamp", "StudentID", "Email", "Name", "Session", "EventType"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getQuizSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(QUIZ_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(QUIZ_SHEET_NAME);
    sheet.appendRow(["Timestamp", "StudentID", "Session", "Topic", "QuizID", "Attempt", "Correct", "IsFirstAttempt"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getPrePostSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PREPOST_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PREPOST_SHEET_NAME);
    sheet.appendRow(["Timestamp", "StudentID", "Email", "Name", "Session", "Phase", "Score", "Total"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* ================= PRE/POST TEST LOGGING (เดิม) ================= */

function logPrePost_(p) {
  const sid = (p.sid || "").trim();
  const email = (p.email || "").trim();
  const name = (p.name || "").trim();
  const session = (p.session || "").trim();
  const phase = (p.phase || "").trim();
  const score = Number(p.score);
  const total = Number(p.total);
  if (!sid || !session || (phase !== "pre" && phase !== "post") || isNaN(score) || isNaN(total)) {
    return jsonOut_({ ok: false, error: "missing or invalid sid/session/phase/score/total" });
  }
  getPrePostSheet_().appendRow([new Date(), sid, email, name, session, phase, score, total]);
  return jsonOut_({ ok: true });
}

/* ================= ATTENDANCE / COMPLETION (เดิม) ================= */

function logEvent_(sheet, p, eventType) {
  const sid = (p.sid || "").trim();
  const email = (p.email || "").trim();
  const name = (p.name || "").trim();
  const session = (p.session || "").trim();
  if (!sid || !session) {
    return jsonOut_({ ok: false, error: "missing sid or session" });
  }
  sheet.appendRow([new Date(), sid, email, name, session, eventType]);
  return jsonOut_({ ok: true });
}

function computeStatus_(sheet, sidRaw) {
  const sid = (sidRaw || "").trim();
  const result = { sid: sid, name: "", email: "", checkins: [], completions: [] };
  if (!sid) return result;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (let i = 0; i < data.length; i++) {
      const rowSid = String(data[i][1]).trim();
      if (rowSid !== sid) continue;
      const email = data[i][2];
      const name = data[i][3];
      const session = String(data[i][4]);
      const eventType = data[i][5];
      if (eventType === "checkin" && result.checkins.indexOf(session) === -1) result.checkins.push(session);
      if (eventType === "complete" && result.completions.indexOf(session) === -1) result.completions.push(session);
      if (name) result.name = name;
      if (email) result.email = email;
    }
  }
  return result;
}

function getStatus_(sheet, sidRaw) {
  const sid = (sidRaw || "").trim();
  if (!sid) return jsonOut_({ ok: false, error: "missing sid" });
  const s = computeStatus_(sheet, sid);
  return jsonOut_({ ok: true, sid: s.sid, name: s.name, email: s.email, checkins: s.checkins, completions: s.completions });
}

/* ================= QUIZ ANSWER LOGGING (ใหม่) ================= */

/**
 * บันทึกทุกครั้งที่นักศึกษากด "ตรวจคำตอบ" ในแบบฝึกหัด — ไม่ว่าจะถูกหรือผิด
 * ระบบจะนับจำนวนครั้งที่พยายามทำข้อนั้น ๆ ของ sid+quizId นี้เองจากประวัติที่มีอยู่แล้ว
 * (ไม่พึ่งพาตัวนับจากฝั่งเบราว์เซอร์ เพื่อให้ข้อมูลถูกต้องแม้รีเฟรชหน้าเว็บ)
 */
function logAnswer_(quizSheet, p) {
  const sid = (p.sid || "").trim();
  const session = (p.session || "").trim();
  const topic = (p.topic || "").trim();
  const quizId = (p.quizId || "").trim();
  const correct = String(p.correct) === "true";

  if (!sid || !quizId) {
    return jsonOut_({ ok: false, error: "missing sid or quizId" });
  }

  const lastRow = quizSheet.getLastRow();
  let attempt = 1;
  if (lastRow > 1) {
    const data = quizSheet.getRange(2, 1, lastRow - 1, 5).getValues(); // Timestamp,StudentID,Session,Topic,QuizID
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][1]).trim() === sid && String(data[i][4]).trim() === quizId) attempt++;
    }
  }
  const isFirst = attempt === 1;
  quizSheet.appendRow([new Date(), sid, session, topic, quizId, attempt, correct, isFirst]);
  return jsonOut_({ ok: true, attempt: attempt, isFirstAttempt: isFirst });
}

/* ================= LEARNING ANALYTICS: predictive risk score ================= *
 * แนวคิด: weighted early-warning score จากข้อมูลพฤติกรรมการเรียนจริง 3 ตัวแปร
 * (ในแนวทางเดียวกับระบบ early-alert เช่น Purdue Course Signals) ได้แก่
 *   F = สัดส่วนที่ตอบถูก "ในความพยายามครั้งแรก" ต่อข้อ (วัดความเข้าใจเนื้อหาจริง)
 *   R = อัตราการลองซ้ำเฉลี่ยต่อข้อ (วัดความไม่มั่นใจ/ความสับสนในเนื้อหา)
 *   C = สัดส่วนครั้งเรียน (1-3) ที่ทำครบแล้ว (วัดความคืบหน้า/การมีส่วนร่วม)
 * riskScore (0-100, ยิ่งสูงยิ่งควรได้รับการดูแลเพิ่มเติม) = 100 * (0.5*(1-F) + 0.3*R' + 0.2*(1-C))
 * โดย R' คือ R ที่ถูก normalize ให้อยู่ในช่วง 0-1 (เฉลี่ยลองซ้ำ 1 ครั้ง = 0, เฉลี่ย 3+ ครั้ง = 1)
 * ค่าน้ำหนัก (0.5/0.3/0.2) และเกณฑ์ risk level ตั้งไว้ตามดุลยพินิจของผู้สอน ปรับได้ตามความเหมาะสม
 */

function computeInsights_(logSheet, quizSheet, sidRaw) {
  const sid = (sidRaw || "").trim();
  const status = computeStatus_(logSheet, sid);

  const out = {
    sid: sid,
    name: status.name,
    email: status.email,
    checkins: status.checkins,
    completions: status.completions,
    totalQuestionsAttempted: 0,
    firstAttemptAccuracy: null,
    avgAttempts: null,
    completionRate: Math.round((status.completions.length / 3) * 100) / 100,
    riskScore: null,
    riskLevel: "no_data",
    riskLabel: "ยังไม่มีข้อมูลเพียงพอ",
    weakTopics: []
  };
  if (!sid) return out;

  const lastRow = quizSheet.getLastRow();
  if (lastRow <= 1) return out;

  const data = quizSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  // Timestamp,StudentID,Session,Topic,QuizID,Attempt,Correct,IsFirstAttempt
  const perQuiz = {}; // quizId -> {topic, attempts, firstCorrect}
  const topicFirst = {}; // topic -> {wrong, total}

  for (let i = 0; i < data.length; i++) {
    const rowSid = String(data[i][1]).trim();
    if (rowSid !== sid) continue;
    const topic = String(data[i][3]).trim() || "(ไม่ระบุหัวข้อ)";
    const quizId = String(data[i][4]).trim();
    const attempt = Number(data[i][5]) || 1;
    const correct = data[i][6] === true || String(data[i][6]).toLowerCase() === "true";
    const isFirst = data[i][7] === true || String(data[i][7]).toLowerCase() === "true";

    if (!perQuiz[quizId]) perQuiz[quizId] = { topic: topic, attempts: 0, firstCorrect: null };
    perQuiz[quizId].attempts++;
    if (isFirst || attempt === 1) perQuiz[quizId].firstCorrect = correct;

    if (isFirst || attempt === 1) {
      if (!topicFirst[topic]) topicFirst[topic] = { wrong: 0, total: 0 };
      topicFirst[topic].total++;
      if (!correct) topicFirst[topic].wrong++;
    }
  }

  const quizIds = Object.keys(perQuiz);
  out.totalQuestionsAttempted = quizIds.length;
  if (quizIds.length === 0) return out;

  let totalAttempts = 0;
  let firstTotal = 0;
  let firstCorrectCount = 0;
  quizIds.forEach(function (qid) {
    const q = perQuiz[qid];
    totalAttempts += q.attempts;
    if (q.firstCorrect !== null) {
      firstTotal++;
      if (q.firstCorrect) firstCorrectCount++;
    }
  });

  const avgAttempts = totalAttempts / quizIds.length;
  out.avgAttempts = Math.round(avgAttempts * 100) / 100;

  if (firstTotal === 0) {
    return out; // ยังไม่มีข้อมูล first-attempt เพียงพอ
  }

  const F = firstCorrectCount / firstTotal;
  out.firstAttemptAccuracy = Math.round(F * 1000) / 1000;

  const Rprime = Math.min(1, Math.max(0, (avgAttempts - 1) / 2));
  const C = out.completionRate;

  const riskRaw = 0.5 * (1 - F) + 0.3 * Rprime + 0.2 * (1 - C);
  const riskScore = Math.round(riskRaw * 100);
  out.riskScore = riskScore;

  if (riskScore < 30) {
    out.riskLevel = "good";
    out.riskLabel = "อยู่ในเกณฑ์ดี — เรียนรู้เนื้อหาได้ค่อนข้างมั่นคง";
  } else if (riskScore < 60) {
    out.riskLevel = "warning";
    out.riskLabel = "ควรทบทวนเพิ่มเติมในบางหัวข้อ";
  } else {
    out.riskLevel = "serious";
    out.riskLabel = "ควรได้รับความช่วยเหลือ/ทบทวนเพิ่มเติมอย่างใกล้ชิด";
  }

  const weak = Object.keys(topicFirst)
    .map(function (t) {
      const d = topicFirst[t];
      return { topic: t, wrongCount: d.wrong, totalCount: d.total, wrongRate: Math.round((d.wrong / d.total) * 100) / 100 };
    })
    .filter(function (t) { return t.wrongRate > 0; })
    .sort(function (a, b) { return b.wrongRate - a.wrongRate || b.wrongCount - a.wrongCount; })
    .slice(0, 3);
  out.weakTopics = weak;

  return out;
}

function getInsights_(logSheet, quizSheet, sidRaw) {
  const sid = (sidRaw || "").trim();
  if (!sid) return jsonOut_({ ok: false, error: "missing sid" });
  const insights = computeInsights_(logSheet, quizSheet, sid);
  return jsonOut_(Object.assign({ ok: true }, insights));
}

/* ================= INSTRUCTOR OVERVIEW (key-gated) ================= */

function getRiskAll_(logSheet, quizSheet, keyRaw) {
  const key = (keyRaw || "").trim();
  if (key !== INSTRUCTOR_KEY) {
    return jsonOut_({ ok: false, error: "invalid key" });
  }

  const sids = [];
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1) {
    const data = logSheet.getRange(2, 1, lastRow - 1, 2).getValues(); // Timestamp, StudentID
    for (let i = 0; i < data.length; i++) {
      const sid = String(data[i][1]).trim();
      if (sid && sids.indexOf(sid) === -1) sids.push(sid);
    }
  }

  const students = sids.map(function (sid) {
    return computeInsights_(logSheet, quizSheet, sid);
  });

  return jsonOut_({ ok: true, count: students.length, students: students });
}

function jsonOut_(obj) {
  const body = JSON.stringify(obj);
  if (CURRENT_CALLBACK && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(CURRENT_CALLBACK)) {
    return ContentService.createTextOutput(CURRENT_CALLBACK + "(" + body + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}
