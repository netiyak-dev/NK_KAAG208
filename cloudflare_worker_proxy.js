/**
 * KAAG208 — Cloudflare Worker + D1 backend (ไม่พึ่ง Google Sheets/Apps Script อีกต่อไป)
 *
 * เดิม Worker นี้เป็นแค่ proxy ไปยัง Google Apps Script แต่ Apps Script มีปัญหาการส่ง
 * ข้อมูลกลับไม่เสถียร (ล่มต่อเนื่องหลายชั่วโมง ไม่เกี่ยวกับ CORS ที่เคยแก้ไปแล้ว) จึงย้าย
 * ฐานข้อมูลทั้งหมดมาไว้ใน Cloudflare D1 (ฐานข้อมูล SQL ของ Cloudflare เอง) แทน
 *
 * วิธีติดตั้ง: ดูไฟล์คู่มือ "KAAG208_คู่มือติดตั้งฐานข้อมูล.md" ที่แนบมาพร้อมกัน
 * ต้องมี D1 database ผูก (binding) ชื่อ DB และ secret ชื่อ INSTRUCTOR_PASSWORD
 */

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}
function json(obj, status = 200) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders() },
  });
}
function str(v) { return (v ?? "").toString().trim(); }
function num(v) { return v === null ? NaN : Number(v); }

/* ---------- สร้างตารางอัตโนมัติถ้ายังไม่มี (ทำครั้งเดียวต่อ isolate ที่ยัง warm อยู่) ---------- */
let schemaReady = null;
function ensureSchema(env) {
  if (!schemaReady) {
    schemaReady = env.DB.batch([
      env.DB.prepare(`CREATE TABLE IF NOT EXISTS log (
        id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT NOT NULL, sid TEXT NOT NULL,
        email TEXT NOT NULL DEFAULT '', name TEXT NOT NULL DEFAULT '',
        session TEXT NOT NULL, event_type TEXT NOT NULL)`),
      env.DB.prepare(`CREATE INDEX IF NOT EXISTS idx_log_sid ON log(sid)`),
      env.DB.prepare(`CREATE TABLE IF NOT EXISTS prepost (
        id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT NOT NULL, sid TEXT NOT NULL,
        email TEXT NOT NULL DEFAULT '', name TEXT NOT NULL DEFAULT '',
        session TEXT NOT NULL, phase TEXT NOT NULL, score REAL NOT NULL, total REAL NOT NULL)`),
      env.DB.prepare(`CREATE INDEX IF NOT EXISTS idx_prepost_sid ON prepost(sid)`),
      env.DB.prepare(`CREATE TABLE IF NOT EXISTS quiz_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT NOT NULL, sid TEXT NOT NULL,
        session TEXT NOT NULL DEFAULT '', topic TEXT NOT NULL DEFAULT '',
        quiz_id TEXT NOT NULL, attempt INTEGER NOT NULL, correct INTEGER NOT NULL,
        is_first_attempt INTEGER NOT NULL)`),
      env.DB.prepare(`CREATE INDEX IF NOT EXISTS idx_quiz_sid_quizid ON quiz_log(sid, quiz_id)`),
    ]).catch((e) => { schemaReady = null; throw e; });
  }
  return schemaReady;
}

/* ---------- checkin / complete ---------- */
async function actionCheckinOrComplete(env, p, eventType) {
  const sid = str(p.get("sid")), email = str(p.get("email")), name = str(p.get("name")), session = str(p.get("session"));
  if (!sid || !session) return { ok: false, error: "missing sid or session" };
  await env.DB.prepare(`INSERT INTO log (ts, sid, email, name, session, event_type) VALUES (?,?,?,?,?,?)`)
    .bind(new Date().toISOString(), sid, email, name, session, eventType).run();
  return { ok: true };
}

/* ---------- status (ใช้ร่วมกับ insights + riskAll) ---------- */
async function fetchLogRows(env, sid) {
  const { results } = await env.DB.prepare(
    `SELECT email, name, session, event_type FROM log WHERE sid = ? ORDER BY id ASC`
  ).bind(sid).all();
  return results;
}
function computeStatusFromRows(sid, rows) {
  const result = { sid, name: "", email: "", checkins: [], completions: [] };
  for (const r of rows) {
    const session = String(r.session);
    if (r.event_type === "checkin" && !result.checkins.includes(session)) result.checkins.push(session);
    if (r.event_type === "complete" && !result.completions.includes(session)) result.completions.push(session);
    if (r.name) result.name = r.name;
    if (r.email) result.email = r.email;
  }
  return result;
}
async function actionStatus(env, p) {
  const sid = str(p.get("sid"));
  if (!sid) return { ok: false, error: "missing sid" };
  const s = computeStatusFromRows(sid, await fetchLogRows(env, sid));
  return { ok: true, sid: s.sid, name: s.name, email: s.email, checkins: s.checkins, completions: s.completions };
}

/* ---------- prepost ---------- */
async function actionPrePost(env, p) {
  const sid = str(p.get("sid")), email = str(p.get("email")), name = str(p.get("name"));
  const session = str(p.get("session")), phase = str(p.get("phase"));
  const score = num(p.get("score")), total = num(p.get("total"));
  if (!sid || !session || (phase !== "pre" && phase !== "post") || isNaN(score) || isNaN(total)) {
    return { ok: false, error: "missing or invalid sid/session/phase/score/total" };
  }
  await env.DB.prepare(`INSERT INTO prepost (ts, sid, email, name, session, phase, score, total) VALUES (?,?,?,?,?,?,?,?)`)
    .bind(new Date().toISOString(), sid, email, name, session, phase, score, total).run();
  return { ok: true };
}

/* ---------- logAnswer ---------- */
async function actionLogAnswer(env, p) {
  const sid = str(p.get("sid")), session = str(p.get("session")), topic = str(p.get("topic"));
  const quizId = str(p.get("quizId")), correct = str(p.get("correct")) === "true";
  if (!sid || !quizId) return { ok: false, error: "missing sid or quizId" };

  const row = await env.DB.prepare(`SELECT COUNT(*) AS c FROM quiz_log WHERE sid = ? AND quiz_id = ?`)
    .bind(sid, quizId).first();
  const attempt = (row?.c || 0) + 1;
  const isFirst = attempt === 1;

  await env.DB.prepare(
    `INSERT INTO quiz_log (ts, sid, session, topic, quiz_id, attempt, correct, is_first_attempt) VALUES (?,?,?,?,?,?,?,?)`
  ).bind(new Date().toISOString(), sid, session, topic, quizId, attempt, correct ? 1 : 0, isFirst ? 1 : 0).run();

  return { ok: true, attempt, isFirstAttempt: isFirst };
}

/* ---------- insights: สูตร predictive risk score (พอร์ตมาจาก Code.gs เดิมทุกตัวอักษร) ---------- */
function computeInsightsFromRows(sid, logRows, quizRows) {
  const status = computeStatusFromRows(sid, logRows);
  const out = {
    sid, name: status.name, email: status.email,
    checkins: status.checkins, completions: status.completions,
    totalQuestionsAttempted: 0,
    firstAttemptAccuracy: null, avgAttempts: null,
    completionRate: Math.round((status.completions.length / 3) * 100) / 100,
    riskScore: null, riskLevel: "no_data", riskLabel: "ยังไม่มีข้อมูลเพียงพอ",
    weakTopics: [],
  };
  if (!sid || quizRows.length === 0) return out;

  const perQuiz = {}, topicFirst = {};
  for (const r of quizRows) {
    const topic = String(r.topic || "").trim() || "(ไม่ระบุหัวข้อ)";
    const quizId = String(r.quiz_id).trim();
    const attempt = Number(r.attempt) || 1;
    const correct = r.correct === 1;
    const isFirst = r.is_first_attempt === 1;

    if (!perQuiz[quizId]) perQuiz[quizId] = { attempts: 0, firstCorrect: null };
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

  let totalAttempts = 0, firstTotal = 0, firstCorrectCount = 0;
  for (const qid of quizIds) {
    const q = perQuiz[qid];
    totalAttempts += q.attempts;
    if (q.firstCorrect !== null) { firstTotal++; if (q.firstCorrect) firstCorrectCount++; }
  }

  const avgAttempts = totalAttempts / quizIds.length;
  out.avgAttempts = Math.round(avgAttempts * 100) / 100;

  if (firstTotal === 0) return out;

  const F = firstCorrectCount / firstTotal;
  out.firstAttemptAccuracy = Math.round(F * 1000) / 1000;

  const Rprime = Math.min(1, Math.max(0, (avgAttempts - 1) / 2));
  const C = out.completionRate;
  out.riskScore = Math.round((0.5 * (1 - F) + 0.3 * Rprime + 0.2 * (1 - C)) * 100);

  if (out.riskScore < 30) { out.riskLevel = "good"; out.riskLabel = "อยู่ในเกณฑ์ดี — เรียนรู้เนื้อหาได้ค่อนข้างมั่นคง"; }
  else if (out.riskScore < 60) { out.riskLevel = "warning"; out.riskLabel = "ควรทบทวนเพิ่มเติมในบางหัวข้อ"; }
  else { out.riskLevel = "serious"; out.riskLabel = "ควรได้รับความช่วยเหลือ/ทบทวนเพิ่มเติมอย่างใกล้ชิด"; }

  out.weakTopics = Object.keys(topicFirst)
    .map((t) => { const d = topicFirst[t]; return { topic: t, wrongCount: d.wrong, totalCount: d.total, wrongRate: Math.round((d.wrong / d.total) * 100) / 100 }; })
    .filter((t) => t.wrongRate > 0)
    .sort((a, b) => b.wrongRate - a.wrongRate || b.wrongCount - a.wrongCount)
    .slice(0, 3);

  return out;
}

async function actionInsights(env, p) {
  const sid = str(p.get("sid"));
  if (!sid) return { ok: false, error: "missing sid" };
  const [logRows, quizResult] = await Promise.all([
    fetchLogRows(env, sid),
    env.DB.prepare(`SELECT topic, quiz_id, attempt, correct, is_first_attempt FROM quiz_log WHERE sid = ?`).bind(sid).all(),
  ]);
  return { ok: true, ...computeInsightsFromRows(sid, logRows, quizResult.results) };
}

/* ---------- riskAll: ยิง D1 แค่ 2 query รวม ไม่ว่านักศึกษาจะมีกี่คน (กันชน free-tier limit) ---------- */
async function actionRiskAll(env, p) {
  const key = str(p.get("key"));
  const expected = (env.INSTRUCTOR_PASSWORD || "").trim();
  if (!expected || key !== expected) return { ok: false, error: "invalid key" };

  const [logAll, quizAll] = await Promise.all([
    env.DB.prepare(`SELECT sid, email, name, session, event_type FROM log ORDER BY id ASC`).all(),
    env.DB.prepare(`SELECT sid, topic, quiz_id, attempt, correct, is_first_attempt FROM quiz_log ORDER BY id ASC`).all(),
  ]);

  const logBySid = new Map(), sidOrder = [];
  for (const r of logAll.results) {
    if (!logBySid.has(r.sid)) { logBySid.set(r.sid, []); sidOrder.push(r.sid); }
    logBySid.get(r.sid).push(r);
  }
  const quizBySid = new Map();
  for (const r of quizAll.results) {
    if (!quizBySid.has(r.sid)) quizBySid.set(r.sid, []);
    quizBySid.get(r.sid).push(r);
  }

  const students = sidOrder.map((sid) => computeInsightsFromRows(sid, logBySid.get(sid) || [], quizBySid.get(sid) || []));
  return { ok: true, count: students.length, students };
}

/* ---------- rawData: ให้อาจารย์ดู/ดาวน์โหลดข้อมูลดิบจากหน้า dashboard ---------- */
const RAW_TABLES = {
  log:     { table: "log",      columns: ["id","ts","sid","email","name","session","event_type"] },
  prepost: { table: "prepost",  columns: ["id","ts","sid","email","name","session","phase","score","total"] },
  quiz:    { table: "quiz_log", columns: ["id","ts","sid","session","topic","quiz_id","attempt","correct","is_first_attempt"] },
};
async function actionRawData(env, p) {
  const key = str(p.get("key"));
  const expected = (env.INSTRUCTOR_PASSWORD || "").trim();
  if (!expected || key !== expected) return { ok: false, error: "invalid key" };

  const def = RAW_TABLES[str(p.get("table"))];
  if (!def) return { ok: false, error: "unknown table" };
  const { results } = await env.DB.prepare(
    `SELECT ${def.columns.join(", ")} FROM ${def.table} ORDER BY id DESC LIMIT 5000`
  ).all();
  return { ok: true, table: str(p.get("table")), columns: def.columns, rows: results };
}

export default {
  async fetch(request, env, ctx) {
    if (request.method === "OPTIONS") return new Response(null, { headers: corsHeaders() });
    const url = new URL(request.url);
    const action = url.searchParams.get("action");
    try {
      await ensureSchema(env);
      let result;
      switch (action) {
        case "checkin":   result = await actionCheckinOrComplete(env, url.searchParams, "checkin"); break;
        case "complete":  result = await actionCheckinOrComplete(env, url.searchParams, "complete"); break;
        case "status":    result = await actionStatus(env, url.searchParams); break;
        case "prepost":   result = await actionPrePost(env, url.searchParams); break;
        case "logAnswer": result = await actionLogAnswer(env, url.searchParams); break;
        case "insights":  result = await actionInsights(env, url.searchParams); break;
        case "riskAll":   result = await actionRiskAll(env, url.searchParams); break;
        case "rawData":   result = await actionRawData(env, url.searchParams); break;
        default: result = { ok: false, error: "unknown action" };
      }
      return json(result, 200);
    } catch (err) {
      return json({ ok: false, error: String(err) }, 500);
    }
  },
};
