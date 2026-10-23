import express from 'express';
import path from 'path';
import fs from 'fs';
import JSZip from 'jszip';
import { createServer as createViteServer } from 'vite';
// @ts-ignore
import mammoth from 'mammoth';
// @ts-ignore
import * as pdfParseModule from 'pdf-parse';
import {
  INITIAL_DATABASE,
  INITIAL_UAS_QUESTIONS,
  INITIAL_QUIZ_QUESTIONS,
  INITIAL_STUDENTS,
  INITIAL_GROUPS,
  INITIAL_MEETINGS,
  INITIAL_ALL_COURSES_DATA,
  INITIAL_STUDENTS_SEM3,
  INITIAL_GROUPS_SEM3,
  INITIAL_MEETINGS_SEM3,
  INITIAL_STUDENTS_KEBIJAKAN_SEM3,
  INITIAL_GROUPS_KEBIJAKAN_SEM3,
  INITIAL_MEETINGS_KEBIJAKAN_SEM3,
  INITIAL_UTS_QUESTIONS_SEM3,
  INITIAL_COURSE_PROFILE,
  INITIAL_COURSE_PROFILE_SEM3,
  INITIAL_COURSE_PROFILE_KEBIJAKAN_SEM3,
} from './src/data/initialData';
import {
  SiakadDatabase,
  Student,
  IndividualSubmission,
  UtsSubmission,
  UtsQuestion,
  UtsDraft,
  StudentGrade,
  MeetingSchedule,
  GroupProject,
  CourseSummary,
  ArchivedSemester,
  DosenProfile,
  QuizQuestion,
  QuizSubmission,
  QuizSettings,
  ActivityPointLog,
} from './src/types';

const IS_SERVERLESS = Boolean(process.env.VERCEL || process.env.AWS_LAMBDA_FUNCTION_NAME);
const DATA_DIR = IS_SERVERLESS ? path.join('/tmp', 'data') : path.join(process.cwd(), 'data');
const DATA_FILE = path.join(DATA_DIR, 'siakad-db.json');
const BACKUP_FILE = path.join(DATA_DIR, 'siakad-db.backup.json');
const SUBMISSIONS_VAULT_FILE = path.join(DATA_DIR, 'siakad-all-time-submissions.json');

// Ensure data directory exists
try {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
  // On serverless / Vercel, copy seed database from repo to /tmp/data if not present
  if (IS_SERVERLESS && !fs.existsSync(DATA_FILE)) {
    const seedFile = path.join(process.cwd(), 'data', 'siakad-db.json');
    if (fs.existsSync(seedFile)) {
      try {
        fs.copyFileSync(seedFile, DATA_FILE);
      } catch (cpErr) {
        console.warn('Seed copy notice:', cpErr);
      }
    }
  }
} catch (e) {
  console.warn('Notice: read-only filesystem or data directory exists:', e);
}

const app = express();
const PORT = 3000;

// Handle pre-parsed body and normalize URL for serverless platforms (like Vercel)
app.use((req, res, next) => {
  if (req.body) {
    if (typeof req.body === 'string') {
      try {
        req.body = JSON.parse(req.body);
      } catch {}
    } else if (Buffer.isBuffer(req.body)) {
      try {
        req.body = JSON.parse(req.body.toString('utf-8'));
      } catch {}
    }
    (req as any)._body = true;
  }
  if (IS_SERVERLESS) {
    if (req.url && req.url.includes('[...all]')) {
      const allQuery = (req.query as any)?.all;
      if (allQuery) {
        const segs = Array.isArray(allQuery) ? allQuery.join('/') : String(allQuery);
        req.url = req.url.replace('/api/[...all]', '/api/' + segs.replace(/^\/+/, ''))
                         .replace('/[...all]', '/api/' + segs.replace(/^\/+/, ''))
                         .replace('[...all]', segs.replace(/^\/+/, ''));
      }
    }
    if (
      req.url &&
      !req.url.startsWith('/api') &&
      req.url !== '/' &&
      !req.url.startsWith('/assets') &&
      !req.url.startsWith('/vite') &&
      !req.url.startsWith('/src') &&
      !req.url.startsWith('/@') &&
      !req.url.startsWith('/node_modules') &&
      !req.url.includes('.')
    ) {
      req.url = '/api' + (req.url.startsWith('/') ? req.url : '/' + req.url);
    }
  }
  next();
});

// Increase payload limit for uploaded presentation slides/files
app.use(express.json({ limit: '150mb' }));
app.use(express.urlencoded({ extended: true, limit: '150mb' }));

// In-memory cache + disk persistence
let db: SiakadDatabase;

function calculateLetterGrade(score: number): string {
  if (score >= 85) return 'A';
  if (score >= 80) return 'A-';
  if (score >= 75) return 'B+';
  if (score >= 70) return 'B';
  if (score >= 65) return 'B-';
  if (score >= 60) return 'C+';
  if (score >= 55) return 'C';
  if (score >= 40) return 'D';
  return 'E';
}

function recalculateStudentGrade(g: StudentGrade): void {
  const att = Number(g.attendanceScore ?? 100);
  const attit = Number(g.attitudeScore ?? 85);
  const indiv = Number(g.individualScore ?? 85);
  const uts = Number(g.utsScore ?? indiv ?? 85);
  const uas = Number(g.uasScore ?? g.groupScore ?? 85);

  g.finalScore = Math.round(
    (att * 0.15) +
    (attit * 0.10) +
    (indiv * 0.25) +
    (uts * 0.25) +
    (uas * 0.25)
  );
  g.letterGrade = calculateLetterGrade(g.finalScore);
}

// Extract text from uploaded document (Word docx, PDF, or text files)
async function extractTextFromFileBuffer(buffer: Buffer, filename: string): Promise<string> {
  const ext = (filename.split('.').pop() || '').toLowerCase();

  if (ext === 'docx') {
    try {
      const result = await mammoth.extractRawText({ buffer });
      if (result && result.value) {
        return result.value;
      }
    } catch (err) {
      console.warn('Mammoth docx extraction warning:', err);
    }
  }

  if (ext === 'pdf') {
    try {
      const PDFParseClass = (pdfParseModule as any).PDFParse || (pdfParseModule as any).default?.PDFParse || (pdfParseModule as any).default;
      if (typeof PDFParseClass === 'function') {
        const parser = new PDFParseClass({ data: buffer });
        const res = await parser.getText();
        if (typeof parser.destroy === 'function') {
          await parser.destroy();
        }
        if (typeof res === 'string') return res;
        if (res && res.text) return res.text;
      }
    } catch (err) {
      console.warn('PDF-parse extraction warning:', err);
    }
  }

  // Fallback for .txt, .md, .rtf, or doc
  try {
    const rawUtf8 = buffer.toString('utf-8');
    if (!rawUtf8.includes('\0')) {
      return rawUtf8;
    }
    // If binary (e.g. older .doc files), extract printable string sequences
    const matches = rawUtf8.match(/[\x20-\x7E\r\n\t]{4,}/g);
    if (matches && matches.length > 0) {
      return matches.join('\n');
    }
  } catch (err) {
    console.warn('Text buffer decoding error:', err);
  }

  return '';
}

// Parse RPS raw text and automatically construct 16 meetings, tasks, and presenter assignments
function parseRpsContent(
  rawText: string,
  defaultFormat: 'individu' | 'kelompok' | 'auto' = 'auto',
  existingStudents: Student[] = [],
  existingGroups: GroupProject[] = []
): {
  detectedProfile: Partial<DosenProfile>;
  detectedMeetings: MeetingSchedule[];
  updatedStudents: Student[];
  detectedGroups: GroupProject[];
} {
  const lines = rawText.split(/\r?\n/);
  const detectedProfile: Partial<DosenProfile> = {};

  for (const line of lines) {
    const l = line.trim();
    if (!l) continue;

    const mkMatch = l.match(/(?:nama\s*)?mata\s*kuliah\s*[:=]\s*(.+)/i) ||
                    l.match(/^(?:mata\s*kuliah|course)\s*[-:]\s*(.+)/i);
    if (mkMatch && mkMatch[1] && !detectedProfile.courseTitle) {
      detectedProfile.courseTitle = mkMatch[1].trim().replace(/^['"]|['"]$/g, '');
    }

    const codeMatch = l.match(/kode\s*(?:mata\s*kuliah|mk)?\s*[:=]\s*(.+)/i);
    if (codeMatch && codeMatch[1] && !detectedProfile.courseCode) {
      detectedProfile.courseCode = codeMatch[1].trim().toUpperCase();
    }

    const sksMatch = l.match(/(?:bobot\s*)?sks\s*[:=]\s*(\d+)/i);
    if (sksMatch && sksMatch[1] && !detectedProfile.sks) {
      detectedProfile.sks = Number(sksMatch[1]);
    }

    const semMatch = l.match(/semester\s*[:=]\s*(.+)/i);
    if (semMatch && semMatch[1] && !detectedProfile.semester) {
      detectedProfile.semester = semMatch[1].trim();
    }

    const prodiMatch = l.match(/(?:program\s*studi|prodi|jurusan)\s*[:=]\s*(.+)/i);
    if (prodiMatch && prodiMatch[1] && !detectedProfile.studyProgram) {
      detectedProfile.studyProgram = prodiMatch[1].trim();
    }

    const campusMatch = l.match(/(?:perguruan\s*tinggi|universitas|institut|stai|iain|uin|kampus)\s*[:=]\s*(.+)/i);
    if (campusMatch && campusMatch[1] && !detectedProfile.campusName) {
      detectedProfile.campusName = campusMatch[1].trim();
    }

    const dosenMatch = l.match(/(?:dosen\s*pengampu|nama\s*dosen|dosen)\s*[:=]\s*(.+)/i);
    if (dosenMatch && dosenMatch[1] && !detectedProfile.name) {
      const fullDosen = dosenMatch[1].trim();
      detectedProfile.name = fullDosen;
      if (fullDosen.includes(',')) {
        const parts = fullDosen.split(',');
        detectedProfile.dosenName = parts[0].trim();
        detectedProfile.dosenTitle = parts.slice(1).join(',').trim();
      } else {
        detectedProfile.dosenName = fullDosen;
      }
    }

    const descMatch = l.match(/(?:deskripsi|capaian\s*pembelajaran|sinopsis)\s*[:=]\s*(.+)/i);
    if (descMatch && descMatch[1] && !detectedProfile.description) {
      detectedProfile.description = descMatch[1].trim();
    }
  }

  const hasGlobalKelompokKeyword = /kelompok|group|tim|diskusi\s*kelompok/i.test(rawText);
  const parsedMap = new Map<number, { title: string; desc: string; isKelompok?: boolean; isIndividu?: boolean }>();

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    const meetMatch = line.match(/(?:pertemuan|minggu|sesi)(?:\s*ke|\s*ke-)?\s*(\d{1,2})[-:\s]*(.*)/i) ||
                      line.match(/^(\d{1,2})\.\s+(.+)/);

    if (meetMatch) {
      const num = Number(meetMatch[1]);
      if (num >= 1 && num <= 16) {
        const titlePart = (meetMatch[2] || '').trim();
        let extraDesc = '';
        if (i + 1 < lines.length && !lines[i + 1].trim().match(/^(?:pertemuan|minggu|sesi|\d{1,2}\.)/i)) {
          extraDesc = lines[i + 1].trim();
        }

        const combinedText = `${titlePart} ${extraDesc}`;
        const isKelompok = /kelompok|group|tim|diskusi\s*kelompok/i.test(combinedText);
        const isIndividu = /individu|mandiri|perorangan/i.test(combinedText);

        parsedMap.set(num, {
          title: titlePart || `Materi Pokok Pertemuan ${num}`,
          desc: extraDesc || `Kajian materi perkuliahan pertemuan ke-${num}.`,
          isKelompok,
          isIndividu,
        });
      }
    }
  }

  if (!parsedMap.has(8)) {
    const utsLine = lines.find(l => /ujian\s*tengah\s*semester|evaluasi\s*tengah\s*semester|\bUTS\b/i.test(l));
    parsedMap.set(8, {
      title: utsLine ? utsLine.trim().replace(/^[-*•\d\.\s]+/, '') : 'Ujian Tengah Semester (UTS)',
      desc: 'Evaluasi penguasaan materi 7 pertemuan awal.',
    });
  }

  if (!parsedMap.has(16)) {
    const uasLine = lines.find(l => /ujian\s*akhir\s*semester|evaluasi\s*akhir\s*semester|\bUAS\b/i.test(l));
    parsedMap.set(16, {
      title: uasLine ? uasLine.trim().replace(/^[-*•\d\.\s]+/, '') : 'Ujian Akhir Semester (UAS)',
      desc: 'Evaluasi akhir dan pengumpulan karya akhir semester.',
    });
  }

  // --- Auto-extract Students & Groups from RPS Document if present ---
  const detectedGroupMap = new Map<number, { name: string; members: string[]; title?: string }>();
  const detectedStudentsFromRps: Student[] = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;

    // Pattern 1: "Kelompok 1: Ahmad Fauzi, Siti Aisyah, Budi Santoso"
    const groupMatch = line.match(/^(?:kelompok|group)\s*(\d{1,2})(?:\s*\([^)]*\))?\s*[:=]\s*(.+)/i);
    if (groupMatch) {
      const gNum = parseInt(groupMatch[1], 10);
      const rawList = groupMatch[2];
      const memberNames = rawList
        .split(/[,;\t|&]|\s+dan\s+/i)
        .map(n => n.trim().replace(/^[-*•\d\.\s]+/, '').trim())
        .filter(n => n.length >= 2 && !/^(dan|atau|dkk|dll|kelompok|group)$/i.test(n));

      if (memberNames.length > 0) {
        if (!detectedGroupMap.has(gNum)) {
          detectedGroupMap.set(gNum, { name: `KELOMPOK ${gNum}`, members: [] });
        }
        const grp = detectedGroupMap.get(gNum)!;
        for (const rawName of memberNames) {
          const nimMatch = rawName.match(/\(?(?:NIM[:\s]*)?(\d{6,15})\)?/i);
          const cleanNim = nimMatch ? nimMatch[1] : '';
          const cleanName = rawName.replace(/\(?(?:NIM[:\s]*)?\d{6,15}\)?/i, '').trim().toUpperCase();
          if (cleanName && !grp.members.includes(cleanName)) {
            grp.members.push(cleanName);
          }
          if (cleanName && !detectedStudentsFromRps.some(s => s.name === cleanName)) {
            detectedStudentsFromRps.push({
              id: `std-rps-${detectedStudentsFromRps.length + 1}-${Date.now().toString(36)}`,
              nim: cleanNim || `2026${String(detectedStudentsFromRps.length + 1).padStart(4, '0')}`,
              name: cleanName,
              rpsPart: `Pertemuan ${Math.min(15, gNum + 1)}`,
              topic: '',
              meetingNumber: Math.min(15, gNum + 1),
              groupId: gNum,
              createdAt: new Date().toISOString(),
            });
          }
        }
      }
    }

    // Pattern 2: "1. Ahmad Fauzi - 20240101 - Kelompok 1"
    const stdRowMatch = line.match(/^(\d{1,2})[\.\)]\s+([A-Za-z\s'\.]+?)(?:[-–—|,\t]+(?:NIM[:\s]*)?(\d{6,15}))?[-–—|,\t]+(?:kelompok|kel|group)\s*(\d{1,2})/i);
    if (stdRowMatch) {
      const cleanName = stdRowMatch[2].trim().toUpperCase();
      const cleanNim = stdRowMatch[3] || `2026${String(stdRowMatch[1]).padStart(4, '0')}`;
      const gNum = parseInt(stdRowMatch[4], 10);
      if (!detectedGroupMap.has(gNum)) {
        detectedGroupMap.set(gNum, { name: `KELOMPOK ${gNum}`, members: [] });
      }
      const grp = detectedGroupMap.get(gNum)!;
      if (!grp.members.includes(cleanName)) grp.members.push(cleanName);
      if (!detectedStudentsFromRps.some(s => s.name === cleanName)) {
        detectedStudentsFromRps.push({
          id: `std-rps-${detectedStudentsFromRps.length + 1}-${Date.now().toString(36)}`,
          nim: cleanNim,
          name: cleanName,
          rpsPart: `Pertemuan ${Math.min(15, gNum + 1)}`,
          topic: '',
          meetingNumber: Math.min(15, gNum + 1),
          groupId: gNum,
          createdAt: new Date().toISOString(),
        });
      }
    }
  }

  // Check if explicit group count statement exists, e.g. "dibagi menjadi 6 kelompok"
  const explicitCountMatch = rawText.match(/(?:dibagi\s*(?:menjadi|dalam)?|terdiri\s*(?:atas|dari))\s*(\d{1,2})\s*kelompok/i);
  const explicitGroupCount = explicitCountMatch ? parseInt(explicitCountMatch[1], 10) : 0;

  // Decide on groups to use
  let groupsCopy: GroupProject[] = [];
  if (detectedGroupMap.size > 0) {
    const sortedKeys = Array.from(detectedGroupMap.keys()).sort((a, b) => a - b);
    groupsCopy = sortedKeys.map(k => {
      const g = detectedGroupMap.get(k)!;
      return {
        id: k,
        name: g.name || `KELOMPOK ${k}`,
        members: g.members,
        title: `Kajian Proyek Kelompok ${k}`,
        description: `Proyek kolaboratif dan presentasi RPS untuk Kelompok ${k}.`,
        toolsSuggested: 'Canva / AI / Slides',
      };
    });
  } else if (explicitGroupCount > 1 && explicitGroupCount <= 15) {
    for (let i = 1; i <= explicitGroupCount; i++) {
      groupsCopy.push({
        id: i,
        name: `KELOMPOK ${i}`,
        members: [],
        title: `Kajian Proyek Kelompok ${i}`,
        description: `Proyek kolaboratif semester untuk Kelompok ${i}.`,
        toolsSuggested: 'Canva / AI / Slides',
      });
    }
  } else if (existingGroups.length > 0) {
    groupsCopy = existingGroups.map(g => ({ ...g }));
  } else {
    groupsCopy = [
      { id: 1, name: 'KELOMPOK 1', members: [], title: 'Kajian Kelompok 1', description: '', toolsSuggested: 'Canva / AI' },
      { id: 2, name: 'KELOMPOK 2', members: [], title: 'Kajian Kelompok 2', description: '', toolsSuggested: 'Canva / AI' },
      { id: 3, name: 'KELOMPOK 3', members: [], title: 'Kajian Kelompok 3', description: '', toolsSuggested: 'Canva / AI' },
      { id: 4, name: 'KELOMPOK 4', members: [], title: 'Kajian Kelompok 4', description: '', toolsSuggested: 'Canva / AI' },
      { id: 5, name: 'KELOMPOK 5', members: [], title: 'Kajian Kelompok 5', description: '', toolsSuggested: 'Canva / AI' },
    ];
  }

  // Decide on students to use
  let studentsCopy: Student[] = [];
  if (detectedStudentsFromRps.length > 0) {
    studentsCopy = detectedStudentsFromRps;
  } else {
    studentsCopy = existingStudents.map(s => ({ ...s }));
    if (groupsCopy.length > 0 && groupsCopy.every(g => g.members.length === 0) && studentsCopy.length > 0) {
      studentsCopy.forEach((s, idx) => {
        const g = groupsCopy[idx % groupsCopy.length];
        s.groupId = g.id;
        if (!g.members.includes(s.name)) {
          g.members.push(s.name);
        }
      });
    }
  }

  const finalMeetings: MeetingSchedule[] = [];

  let currentGroupIdx = 0;
  let currentStudentIdx = 0;

  for (let m = 1; m <= 16; m++) {
    const isUts = m === 8;
    const isUas = m === 16;
    const parsed = parsedMap.get(m);

    let format: 'individu' | 'kelompok' = 'individu';
    if (isUts) {
      format = 'individu';
    } else if (isUas) {
      format = 'kelompok';
    } else if (parsed?.isKelompok) {
      format = 'kelompok';
    } else if (parsed?.isIndividu) {
      format = 'individu';
    } else if (defaultFormat === 'kelompok') {
      format = 'kelompok';
    } else if (defaultFormat === 'individu') {
      format = 'individu';
    } else {
      format = hasGlobalKelompokKeyword ? 'kelompok' : 'individu';
    }

    const type: 'kuliah' | 'uts' | 'uas' = isUts ? 'uts' : isUas ? 'uas' : 'kuliah';
    const taskType = isUts
      ? 'uts_esai'
      : isUas
      ? 'uas_proyek'
      : format === 'kelompok'
      ? 'presentasi_kelompok'
      : 'makalah_ppt';

    const title = parsed?.title || (
      isUts ? 'Ujian Tengah Semester (UTS)' :
      isUas ? 'Ujian Akhir Semester (UAS)' :
      `Pertemuan ${m}: Pendalaman Materi Perkuliahan`
    );

    const description = parsed?.desc || (
      isUts ? 'Evaluasi penguasaan materi 7 pertemuan awal.' :
      isUas ? 'Evaluasi akhir dan pengumpulan karya akhir semester.' :
      `Pembahasan topik kajian materi perkuliahan pertemuan ke-${m}.`
    );

    const assignmentDescription = isUts
      ? 'Pengerjaan 5 Soal Essay Evaluasi Tengah Semester (UTS)'
      : isUas
      ? 'Pengerjaan & Publikasi Proyek Video Edukasi Kelompok AI'
      : format === 'kelompok'
      ? `Tugas Presentasi Kelompok: Penyusunan Makalah dan Slide PPT Presentasi Topik "${title}"`
      : `Tugas Presentasi Individu: Penyusunan Makalah dan Slide PPT Presentasi Topik "${title}"`;

    const presenters: string[] = [];
    let groupName: string | undefined = undefined;
    let groupId: number | undefined = undefined;

    if (!isUts && !isUas && m > 1) {
      if (format === 'kelompok') {
        const assignedGroup = groupsCopy[currentGroupIdx % groupsCopy.length];
        if (assignedGroup) {
          groupId = assignedGroup.id;
          groupName = assignedGroup.name;
          if (assignedGroup.members && assignedGroup.members.length > 0) {
            presenters.push(...assignedGroup.members);
          }
          studentsCopy.forEach(std => {
            if (std.groupId === assignedGroup.id) {
              std.meetingNumber = m;
              std.topic = title;
              std.rpsPart = `Pertemuan ${m}`;
              if (!presenters.includes(std.name)) {
                presenters.push(std.name);
              }
            }
          });
        }
        currentGroupIdx++;
      } else {
        if (studentsCopy.length > 0) {
          const std = studentsCopy[currentStudentIdx % studentsCopy.length];
          if (std) {
            presenters.push(std.name);
            std.meetingNumber = m;
            std.topic = title;
            std.rpsPart = `Pertemuan ${m}`;
          }
          currentStudentIdx++;
        }
      }
    }

    finalMeetings.push({
      meetingNumber: m,
      dateStr: `Pertemuan ${m}`,
      isoDate: new Date(Date.now() + (m - 1) * 7 * 86400000).toISOString().split('T')[0],
      title,
      presenters,
      partCodes: [`Pertemuan ${m}`],
      type,
      presentationFormat: format,
      groupName,
      groupId,
      taskType,
      description,
      assignmentDescription,
    });
  }

  return {
    detectedProfile,
    detectedMeetings: finalMeetings,
    updatedStudents: studentsCopy,
    detectedGroups: groupsCopy,
  };
}

function isDosenAuthorized(req: express.Request): boolean {
  const authHeader = req.headers['x-dosen-auth'];
  const sessionToken = req.headers['authorization'];
  const customPwdHeader = req.headers['x-dosen-password'];
  const activePassword = db?.dosenPassword || 'filsafat2026';

  return (
    authHeader === 'true' ||
    sessionToken === 'Bearer dosen-authenticated-session' ||
    customPwdHeader === activePassword ||
    req.body?.isDosen === true
  );
}

function generateQuizQuestionsFromMeetings(meetings: MeetingSchedule[], courseTitle: string): QuizQuestion[] {
  const nonExamMeetings = (meetings || []).filter(m => m.type === 'kuliah' && m.meetingNumber !== 8 && m.meetingNumber !== 16);
  const selectedMeetings = nonExamMeetings.slice(0, 10);
  const animationThemes: ('brain' | 'atom' | 'compass' | 'scale' | 'target' | 'shield' | 'lightbulb' | 'book')[] = [
    'brain', 'atom', 'compass', 'scale', 'target', 'shield', 'lightbulb', 'book', 'shield', 'brain'
  ];

  const ttsFallbacks = [
    { keyword: 'KRITIS', clue: 'Sikap rasional yang menyelidiki sampai ke akar terdalam' },
    { keyword: 'ONTOLOGI', clue: 'Cabang filsafat yang mengkaji hakikat realitas atau wujud' },
    { keyword: 'EPISTEMOLOGI', clue: 'Teori filsafat tentang asal-usul, metode, dan validitas pengetahuan' },
    { keyword: 'AKSIOLOGI', clue: 'Pilar filsafat yang membahas nilai kegunaan, moral, dan etika' },
    { keyword: 'DEDUKTIF', clue: 'Metode penalaran dari premis umum menuju kesimpulan khusus' },
    { keyword: 'KORESPONDENSI', clue: 'Teori kebenaran yang bersesuaian dengan fakta empiris di lapangan' },
    { keyword: 'PRAGMATIS', clue: 'Aliran kebenaran yang diukur dari kemanfaatan fungsional nyata' },
    { keyword: 'INTEGRASI', clue: 'Penyatuan harmonis antara ayat qauliyah dan sains kauniyah' },
    { keyword: 'INTEGRITAS', clue: 'Kejujuran moral dan etika luhur akademisi dalam riset & AI' },
    { keyword: 'PARADIGMA', clue: 'Kerangka konseptual sains yang dirumuskan oleh Thomas Kuhn' },
  ];

  return selectedMeetings.map((m, idx) => {
    const qNum = idx + 1;
    const theme = animationThemes[idx % animationThemes.length];
    const cleanTitle = m.title.replace(/^Pertemuan\s*\d+[-:\s]*/i, '').trim();
    const ttsItem = ttsFallbacks[idx % ttsFallbacks.length];

    return {
      id: qNum,
      number: qNum,
      question: `Pertemuan ${m.meetingNumber} (${cleanTitle}): Berdasarkan materi RPS pada pertemuan ini, manakah pernyataan yang paling tepat mencerminkan pemahaman substansi topik tersebut?`,
      options: [
        `Memahami secara mendalam konsep "${cleanTitle}" serta implementasinya yang aplikatif dan beretika dalam konteks ${courseTitle}`,
        `Mengabaikan prinsip dasar metodologi keilmuan dan hanya berfokus pada ringkasan instan tanpa analisis ilmiah`,
        `Memisahkan konsep teoritis dari realitas problematika tata kelola dan manajemen pendidikan Islam`,
        `Menyerahkan sepenuhnya analisis kajian kepada opini subjektif tanpa telaah literatur yang valid`,
      ],
      correctIndex: 0,
      explanation: `Pada Pertemuan ${m.meetingNumber} dengan topik "${cleanTitle}", fokus pembelajaran ditekankan pada penguasaan substansi teoritis, analisis kritis ilmiah, serta integrasi nilai moral-akademik.`,
      points: 10,
      badgeTopic: cleanTitle.length > 30 ? cleanTitle.slice(0, 27) + '...' : cleanTitle,
      animationTheme: theme,
      ttsKeyword: ttsItem.keyword,
      ttsClue: ttsItem.clue,
    };
  });
}

// Generate 10 Quiz Questions from Lecturer Uploaded Quiz Material Document or Text
function generateQuizFromMaterial(
  materialText: string,
  courseTitle: string,
  targetMeeting?: string
): QuizQuestion[] {
  const animationThemes: ('brain' | 'atom' | 'compass' | 'scale' | 'target' | 'shield' | 'lightbulb' | 'book')[] = [
    'brain', 'atom', 'compass', 'scale', 'target', 'shield', 'lightbulb', 'book', 'shield', 'brain'
  ];

  const ttsFallbacks = [
    { keyword: 'FILSAFAT', clue: 'Cinta pada kebijaksanaan dan pencarian kebenaran hakiki' },
    { keyword: 'ONTOLOGI', clue: 'Hakikat keberadaan wujud dan realitas objek ilmu' },
    { keyword: 'EPISTEMOLOGI', clue: 'Cara, sarana, dan metodologi perolehan pengetahuan ilmiah' },
    { keyword: 'AKSIOLOGI', clue: 'Nilai guna kemaslahatan dan etika moral keilmuan' },
    { keyword: 'LOGIKA', clue: 'Sarana penalaran lurus, tertib, dan sistematis' },
    { keyword: 'EMPIRISME', clue: 'Aliran pengetahuan yang bersumber dari pengalaman indrawi' },
    { keyword: 'RASIONALISME', clue: 'Aliran pengetahuan yang bersumber dari penalaran akal budi' },
    { keyword: 'POSITIVISME', clue: 'Paradigma sains yang menuntut bukti faktual terukur' },
    { keyword: 'AMANAH', clue: 'Nilai kejujuran moral akademisi dalam menjaga integritas' },
    { keyword: 'DIALEKTIKA', clue: 'Metode dialogis tesis, antitesis, dan sintesis keilmuan' },
  ];

  // Clean lines and extract meaningful conceptual sentences
  const lines = materialText
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(l => l.length > 15 && !l.toLowerCase().startsWith('halaman') && !l.toLowerCase().startsWith('page'));

  // 10 key topic anchors if material is general
  const fallbackTopics = [
    'Hakikat Berpikir Kefilsafatan dan Logika Berpikir',
    'Fondasi Ontologi Realitas dan Objek Ilmu Pengetahuan',
    'Epistemologi, Validitas Ilmiah, dan Metodologi Kebenaran',
    'Aksiologi, Nilai Moral, dan Tanggung Jawab Ilmuwan Muslim',
    'Integrasi Ilmu, Spiritualitas Islam, dan Nilai Ketauhidan',
    'Dinamika Rasionalisme, Empirisme, dan Positivisme Keilmuan',
    'Hermeneutika Kritis dan Dekonstruksi Pemikiran Modern',
    'Sarana Ilmiah: Bahasa, Matematika, dan Statistika Penelitian',
    'Etika Akademik, Kejujuran Intelektual, dan Pencegahan Plagiarisme',
    'Tantangan Aksiologi Pendidikan Islam di Era Generative AI',
  ];

  const questions: QuizQuestion[] = [];

  for (let i = 0; i < 10; i++) {
    const qNum = i + 1;
    const theme = animationThemes[i % animationThemes.length];
    const snippetLine = lines[i % (lines.length || 1)] || fallbackTopics[i];
    const cleanTopic = fallbackTopics[i];
    const meetingBadge = targetMeeting || `Topik ${qNum}`;
    const ttsItem = ttsFallbacks[i % ttsFallbacks.length];

    questions.push({
      id: qNum,
      number: qNum,
      question: `Soal ${qNum} (${meetingBadge}): Dalam telaah "${cleanTopic}", manakah analisis yang paling komprehensif dan selaras dengan materi yang dikaji?`,
      options: [
        `Memahami esensi teoritis serta praksis "${cleanTopic}" secara kritis, bertanggung jawab, dan terintegrasi dengan kemaslahatan ${courseTitle}.`,
        `Mengabaikan dimensi etika dan epistemologi dengan hanya mengandalkan hafalan parsial tanpa penalaran ilmiah.`,
        `Memisahkan secara dikotomis antara kajian keilmuan kontemporer dengan integritas nilai moral spiritual.`,
        `Menyerahkan sepenuhnya kesimpulan ilmiah pada asumsi instan tanpa verifikasi metodologis yang teruji.`,
      ],
      correctIndex: 0,
      explanation: `Berdasarkan materi yang diunggah dosen (${meetingBadge}), penguasaan atas konsep "${cleanTopic}" menuntut pemahaman metodologis yang mendalam, sikap kritis objektif, serta implementasi etis.`,
      points: 10,
      badgeTopic: cleanTopic.length > 28 ? cleanTopic.slice(0, 25) + '...' : cleanTopic,
      animationTheme: theme,
      ttsKeyword: ttsItem.keyword,
      ttsClue: ttsItem.clue,
    });
  }

  return questions;
}

// Generate 5 In-Depth Essay Questions for UTS from Meetings 1-7 in RPS
function generateUtsQuestionsFromMeetings(meetings: MeetingSchedule[], courseProfile?: any): UtsQuestion[] {
  const m1 = (meetings || []).find(m => m.meetingNumber === 1);
  const m2 = (meetings || []).find(m => m.meetingNumber === 2);
  const m3 = (meetings || []).find(m => m.meetingNumber === 3);
  const m4 = (meetings || []).find(m => m.meetingNumber === 4);
  const m5 = (meetings || []).find(m => m.meetingNumber === 5);
  const m6 = (meetings || []).find(m => m.meetingNumber === 6);
  const m7 = (meetings || []).find(m => m.meetingNumber === 7);

  const courseTitle = courseProfile?.name || courseProfile?.courseTitle || 'Filsafat Ilmu Manajemen Pendidikan Islam';

  return [
    {
      id: 1,
      number: 1,
      title: `Dimensi Ontologis & Fondasi Keilmuan (${m1?.title || 'Pengantar & Hakikat'} & ${m2?.title || 'Objek Kajian'})`,
      topic: `${m1?.title || 'Dasar Filsafat'} & ${m2?.title || 'Fondasi Ontologis'}`,
      question: `Berdasarkan silabus RPS Pertemuan 1 dan 2 (${m1?.title || 'Dasar Filsafat'} serta ${m2?.title || 'Fondasi Ontologis'}), jelaskan secara komprehensif hakikat ontologis dan objek kajian formal maupun material dari mata kuliah ${courseTitle}! Bagaimana hubungan dialektis antara konsep dasar tersebut dengan paradigma ilmu pengetahuan modern?`,
      guide: `Jawab dengan merujuk materi RPS pertemuan 1-2. Sebutkan minimal 2 pandangan filosof/ahli, definisi ontologi keilmuan, serta refleksinya dalam praksis pendidikan.`,
      rubric: `Ketajaman analisis ontologis (8 poin), rujukan konseptual RPS (6 poin), orisinalitas argumentasi (6 poin). Total 20 poin.`,
      maxScore: 20
    },
    {
      id: 2,
      number: 2,
      title: `Epistemologi & Metodologi Ilmiah (${m3?.title || 'Struktur Epistemologi'} & ${m4?.title || 'Metodologi Penelitian'})`,
      topic: `${m3?.title || 'Epistemologi'} & ${m4?.title || 'Metode Ilmiah'}`,
      question: `Telaah secara mendalam materi RPS Pertemuan 3 dan 4 mengenai epistemologi dan metodologi pencarian kebenaran ilmiah dalam lingkup ${courseTitle} (${m3?.title || 'Epistemologi'} & ${m4?.title || 'Metode Ilmiah'})! Bedakan antara pendekatan rasionalisme, empirisme, dan intuisi/wahyu dalam membangun validitas ilmiah!`,
      guide: `Uraikan konstruksi metodologis keilmuan, kriteria kebenaran ilmiah (koherensi, korespondensi, pragmatis), dan hubungannya dengan capaian pembelajaran RPS pertemuan 3-4.`,
      rubric: `Pemahaman teori epistemologi (8 poin), perbandingan metodologis (6 poin), ketepatan contoh (6 poin). Total 20 poin.`,
      maxScore: 20
    },
    {
      id: 3,
      number: 3,
      title: `Aksiologi, Etika & Tanggung Jawab Moral Keilmuan (${m5?.title || 'Etika Keilmuan'})`,
      topic: m5?.title || 'Aksiologi & Etika Profesi Keilmuan',
      question: `Mengacu pada pembahasan materi Pertemuan 5 (${m5?.title || 'Aksiologi & Etika Profesi'}), bagaimana ilmuwan dan akademisi menyeimbangkan antara kebebasan nilai (value-free) dan keterikatan nilai (value-bound) dalam penerapan ilmu? Analisis tanggung jawab etis dan moral ilmuwan di era disrupsi digital saat ini!`,
      guide: `Fokus pada dimensi aksiologis, tanggung jawab etika profesi ilmuwan, mitigasi penyalahgunaan teknologi/AI, serta nilai moralitas akademis.`,
      rubric: `Kedalaman pemikiran aksiologis (8 poin), relevansi etika modern (6 poin), struktur jawaban akademis (6 poin). Total 20 poin.`,
      maxScore: 20
    },
    {
      id: 4,
      number: 4,
      title: `Studi Kasus & Dialektika Realitas Empiris (${m6?.title || 'Kajian Kasus Empiris'})`,
      topic: m6?.title || 'Kajian Kasus Lapangan & Realitas Empiris',
      question: `Berdasarkan kajian materi Pertemuan 6 (${m6?.title || 'Aplikasi Lapangan'}), lakukan analisis kritis terhadap salah satu problematika aktual yang relevan dengan ${courseTitle}! Bagaimana konsep yang telah dipelajari mampu memberikan kerangka diagnostik dan solusi holistik terhadap masalah tersebut?`,
      guide: `Identifikasi masalah riil, gunakan instrumen analisis materi RPS pertemuan 6, dan tawarkan rekomendasi ilmiah konkret.`,
      rubric: `Identifikasi kasus konkret (7 poin), penerapan teori RPS (7 poin), kekuatan argumentasi solusi (6 poin). Total 20 poin.`,
      maxScore: 20
    },
    {
      id: 5,
      number: 5,
      title: `Sintesis Kritis & Refleksi Teori Integratif Pra-UTS (${m7?.title || 'Sintesis Materi Integratif'})`,
      topic: m7?.title || 'Sintesis Konseptual Integratif Pra-UTS',
      question: `Sebagai sintesis komprehensif atas seluruh materi RPS Pertemuan 1 hingga 7 menjelang UTS, rumuskan proposisi atau kerangka konseptual baru yang mengintegrasikan fondasi ontologi, epistemologi, dan aksiologi dalam mata kuliah ${courseTitle}! Jelaskan urgensinya bagi pengembangan kompetensi kepemimpinan dan manajerial!`,
      guide: `Tuliskan sintesis integratif berbobot akademik, hindari pengulangan hafalan, tunjukkan pemikiran reflektif tingkat tinggi (HOTS).`,
      rubric: `Kemampuan sintesis integratif (8 poin), orisinalitas kerangka konseptual (6 poin), logika penalaran akademis (6 poin). Total 20 poin.`,
      maxScore: 20
    }
  ];
}

// Generate 5 In-Depth Essay Questions for UAS from Meetings 9-15 in RPS
function generateUasQuestionsFromMeetings(meetings: MeetingSchedule[], courseProfile?: any): UtsQuestion[] {
  const m9 = (meetings || []).find(m => m.meetingNumber === 9);
  const m10 = (meetings || []).find(m => m.meetingNumber === 10);
  const m11 = (meetings || []).find(m => m.meetingNumber === 11);
  const m12 = (meetings || []).find(m => m.meetingNumber === 12);
  const m13 = (meetings || []).find(m => m.meetingNumber === 13);
  const m14 = (meetings || []).find(m => m.meetingNumber === 14);
  const m15 = (meetings || []).find(m => m.meetingNumber === 15);

  const courseTitle = courseProfile?.name || courseProfile?.courseTitle || 'Filsafat Ilmu Manajemen Pendidikan Islam';

  return [
    {
      id: 1,
      number: 1,
      title: `Paradigma Mutakhir & Konstruksi Teori Kontemporer (${m9?.title || 'Perkembangan Mutakhir'} & ${m10?.title || 'Paradigma Keilmuan'})`,
      topic: `${m9?.title || 'Perkembangan Mutakhir'} & ${m10?.title || 'Paradigma Keilmuan'}`,
      question: `Berdasarkan silabus RPS Pertemuan 9 dan 10 (${m9?.title || 'Materi Pertemuan 9'} serta ${m10?.title || 'Materi Pertemuan 10'}), evaluasi secara kritis pergeseran paradigma keilmuan kontemporer dalam mata kuliah ${courseTitle}! Bagaimana paradigma baru ini menjawab keterbatasan teori-teori konvensional masa lampau?`,
      guide: `Analisis pergeseran paradigma, rujukan silabus pertemuan 9-10, bandingkan model lama dan model mutakhir dengan argumen akademis kokoh.`,
      rubric: `Ketajaman perbandingan paradigma (8 poin), pemahaman materi pasca-UTS RPS (6 poin), kualitas argumentasi ilmiah (6 poin). Total 20 poin.`,
      maxScore: 20
    },
    {
      id: 2,
      number: 2,
      title: `Analisis Kebijakan, Manajerial & Tata Kelola (${m11?.title || 'Implementasi Manajerial & Kebijakan'})`,
      topic: m11?.title || 'Manajerial dan Strategi Kebijakan',
      question: `Mengkaji materi RPS Pertemuan 11 (${m11?.title || 'Manajerial dan Strategi'}), bagaimana prinsip-prinsip filosofis dan manajerial diimplementasikan dalam formulasi kebijakan kelembagaan? Berikan contoh kasus kebijakan konkret dan telaah efektivitas serta dampak etisnya!`,
      guide: `Uraikan aspek manajerial, kebijakan praktis, studi kasus riil, dan relevansi implementasi terhadap peningkatan kualitas kelembagaan.`,
      rubric: `Analisis kebijakan & manajemen (8 poin), relevansi studi kasus riil (6 poin), argumentasi solutif (6 poin). Total 20 poin.`,
      maxScore: 20
    },
    {
      id: 3,
      number: 3,
      title: `Problematika Kontemporer & Disrupsi Masa Depan (${m12?.title || 'Tantangan Era Disrupsi'} & ${m13?.title || 'Problematika Global'})`,
      topic: `${m12?.title || 'Isu Global'} & ${m13?.title || 'Disrupsi Teknologi'}`,
      question: `Berdasarkan telaah kritis RPS Pertemuan 12 dan 13 (${m12?.title || 'Isu Global'} & ${m13?.title || 'Disrupsi Teknologi'}), analisis dinamika tantangan disrupsi kecerdasan buatan (AI) dan globalisasi terhadap keberlanjutan keilmuan ${courseTitle}! Langkah adaptif apa yang wajib diambil oleh praktisi dan akademisi?`,
      guide: `Soroti tantangan disrupsi zaman, ancaman dekadensi nilai, peluang inovasi, serta strategi adaptasi berkelanjutan.`,
      rubric: `Kedalaman telaah disrupsi (8 poin), orisinalitas gagasan antisipatif (6 poin), logika penalaran (6 poin). Total 20 poin.`,
      maxScore: 20
    },
    {
      id: 4,
      number: 4,
      title: `Model Integratif & Desain Solutif Lapangan (${m14?.title || 'Inovasi Lapangan'} & ${m15?.title || 'Praksis Solusi'})`,
      topic: `${m14?.title || 'Inovasi Lapangan'} & ${m15?.title || 'Praksis Solusi'}`,
      question: `Mengacu pada pembahasan materi Pertemuan 14 dan 15 (${m14?.title || 'Praksis Inovatif'} serta ${m15?.title || 'Evaluasi Holistik'}), rancanglah sebuah model konseptual atau blueprint inovatif untuk memecahkan problem sistemik dalam bidang ${courseTitle}! Jelaskan tahapan implementasi dan indikator keberhasilannya!`,
      guide: `Kemukakan desain blueprint orisinal, tahapan implementasi sistematis, serta matriks indikator capaian yang terukur.`,
      rubric: `Inovasi dan kebaruan desain model (8 poin), kelayakan implementasi (6 poin), koherensi akademik (6 poin). Total 20 poin.`,
      maxScore: 20
    },
    {
      id: 5,
      number: 5,
      title: `Rekonstruksi Holistik & Refleksi Komprehensif Akhir Semester (${courseTitle})`,
      topic: 'Sintesis Komprehensif Akhir Semester',
      question: `Sebagai evaluasi akhir komprehensif atas seluruh perjalanan perkuliahan 16 pertemuan dalam RPS ${courseTitle}, lakukan sintesis epistemologis menyeluruh dari pertemuan 1 hingga 15! Apa refleksi mendasar mengenai peran keilmuan ini dalam membangun peradaban berkeadaban dan bagaimana visi akademik Anda ke depan?`,
      guide: `Tuliskan esai refleksi tingkat tinggi (HOTS) yang mengikat seluruh benang merah materi dari awal hingga akhir perkuliahan.`,
      rubric: `Sintesis komprehensif seluruh materi RPS (8 poin), kedalaman refleksi filosofis (6 poin), proyeksi visi ke depan (6 poin). Total 20 poin.`,
      maxScore: 20
    }
  ];
}

// Parse Students from Document Text (PDF, DOCX, DOC, XLSX, CSV, TXT)
function parseStudentsFromDocumentText(text: string): any[] {
  if (!text || !text.trim()) return [];

  // Check JSON format
  const trimmed = text.trim();
  if ((trimmed.startsWith('[') && trimmed.endsWith(']')) || (trimmed.startsWith('{') && trimmed.endsWith('}'))) {
    try {
      const parsedJson = JSON.parse(trimmed);
      const list = Array.isArray(parsedJson) ? parsedJson : (parsedJson.students || parsedJson.data || []);
      if (Array.isArray(list) && list.length > 0) {
        return list.map((item, idx) => ({
          nim: String(item.nim || item.NIM || `202601${String(idx + 1).padStart(2, '0')}`).trim(),
          name: String(item.name || item.nama || item.Nama || '').toUpperCase().trim(),
          birthPlace: String(item.birthPlace || item.tempatLahir || item.tempat_lahir || item.Tempat_Lahir || 'Pasuruan').trim(),
          birthDate: String(item.birthDate || item.tanggalLahir || item.tanggal_lahir || item.Tanggal_Lahir || '1998-05-15').trim(),
          address: String(item.address || item.alamat || item.Alamat || 'Pasuruan, Jawa Timur').trim(),
          gender: /perempuan|wanita|p/i.test(String(item.gender || item.jenisKelamin || item.jenis_kelamin || '')) ? 'Perempuan' : 'Laki-laki',
          phone: String(item.phone || item.noHp || item.telepon || '').trim(),
          groupId: Number(item.groupId || (idx % 5) + 1),
          rpsPart: item.rpsPart || `Part ${String(idx + 1).padStart(2, '0')}`,
          topic: item.topic || 'Filsafat Ilmu MPI',
          meetingNumber: item.meetingNumber || (idx % 16) + 1,
        })).filter(s => s.name);
      }
    } catch {
      // not valid json, fall back to line parser
    }
  }

  const rawLines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const parsedStudents: any[] = [];

  // Check if text has Key-Value Blocks (e.g. NIM: ..., Nama: ...)
  let currentBlock: any = {};
  let isKeyValBlock = false;

  for (let i = 0; i < rawLines.length; i++) {
    const line = rawLines[i];
    const kvMatch = line.match(/^(nim|nama(?:\s*lengkap)?|tempat(?:\s*lahir)?|tanggal(?:\s*lahir)?|ttl|alamat|jenis\s*kelamin|jk|no(?:\s*hp|\s*wa|\s*telepon)?)\s*[:=]\s*(.+)/i);
    if (kvMatch) {
      isKeyValBlock = true;
      const key = kvMatch[1].toLowerCase().replace(/\s+/g, '');
      const val = kvMatch[2].trim();
      if (key.includes('nim')) currentBlock.nim = val;
      else if (key.includes('nama')) currentBlock.name = val.toUpperCase();
      else if (key.includes('tempat') && !key.includes('tanggal')) currentBlock.birthPlace = val;
      else if (key.includes('tanggal')) currentBlock.birthDate = val;
      else if (key.includes('ttl')) {
        const parts = val.split(/[,/]/);
        currentBlock.birthPlace = (parts[0] || '').trim();
        currentBlock.birthDate = (parts.slice(1).join(',') || '').trim();
      } else if (key.includes('alamat')) currentBlock.address = val;
      else if (key.includes('kelamin') || key === 'jk') {
        currentBlock.gender = /perempuan|wanita|p/i.test(val) ? 'Perempuan' : 'Laki-laki';
      } else if (key.includes('hp') || key.includes('wa') || key.includes('telepon')) {
        currentBlock.phone = val;
      }
    } else if (isKeyValBlock && currentBlock.name && (currentBlock.nim || Object.keys(currentBlock).length >= 2)) {
      parsedStudents.push(currentBlock);
      currentBlock = {};
    }
  }
  if (isKeyValBlock && currentBlock.name) {
    parsedStudents.push(currentBlock);
  }

  if (parsedStudents.length > 0) {
    return parsedStudents.map((s, idx) => ({
      nim: s.nim || `202601${String(idx + 1).padStart(2, '0')}`,
      name: s.name.toUpperCase(),
      birthPlace: s.birthPlace || 'Pasuruan',
      birthDate: s.birthDate || '1998-05-15',
      address: s.address || 'Pasuruan, Jawa Timur',
      gender: s.gender || 'Laki-laki',
      phone: s.phone || '',
      groupId: (idx % 5) + 1,
      rpsPart: `Part ${String(idx + 1).padStart(2, '0')}`,
      topic: 'Filsafat Ilmu MPI',
      meetingNumber: (idx % 16) + 1,
    }));
  }

  // Otherwise, parse line by line (Tab, Pipe, Comma, Semicolon, or Multi-space delimited table rows)
  for (let idx = 0; idx < rawLines.length; idx++) {
    const line = rawLines[idx];
    // Skip table header lines
    if (/(?:^|\b)(?:no|nim|nama|mahasiswa|tempat\s*lahir|tanggal\s*lahir|alamat|jenis\s*kelamin|ttl)(?:\b|$)/i.test(line) &&
        /(?:nama|mahasiswa)/i.test(line) && /(?:nim|no)/i.test(line)) {
      continue;
    }

    let cols = line.includes('\t')
      ? line.split('\t')
      : line.includes('|')
      ? line.split('|')
      : line.includes(';')
      ? line.split(';')
      : line.includes(',') && line.split(',').length >= 3
      ? line.split(',')
      : line.split(/\s{2,}/);

    cols = cols.map(c => c.trim()).filter(Boolean);
    if (cols.length < 2) continue;

    // Remove row number (e.g. "1" or "1.")
    if (/^\d{1,3}\.?$/.test(cols[0]) && cols.length >= 3) {
      cols.shift();
    }

    let nim = '';
    let name = '';
    let birthPlace = '';
    let birthDate = '';
    let address = '';
    let gender = 'Laki-laki';
    let phone = '';

    if (/^\d{6,14}$/.test(cols[0].replace(/[-.]/g, ''))) {
      nim = cols[0].replace(/[-.]/g, '');
      name = cols[1] || '';
      birthPlace = cols[2] || '';
      birthDate = cols[3] || '';
      address = cols[4] || '';
      gender = cols[5] || '';
      phone = cols[6] || '';
    } else if (cols[1] && /^\d{6,14}$/.test(cols[1].replace(/[-.]/g, ''))) {
      name = cols[0] || '';
      nim = cols[1].replace(/[-.]/g, '');
      birthPlace = cols[2] || '';
      birthDate = cols[3] || '';
      address = cols[4] || '';
      gender = cols[5] || '';
      phone = cols[6] || '';
    } else {
      const nimIdx = cols.findIndex(c => /^\d{6,14}$/.test(c.replace(/[-.]/g, '')));
      if (nimIdx !== -1) {
        nim = cols[nimIdx].replace(/[-.]/g, '');
        const remaining = cols.filter((_, i) => i !== nimIdx);
        name = remaining[0] || '';
        birthPlace = remaining[1] || '';
        birthDate = remaining[2] || '';
        address = remaining[3] || '';
        gender = remaining[4] || '';
        phone = remaining[5] || '';
      } else {
        name = cols[0] || '';
        birthPlace = cols[1] || '';
        birthDate = cols[2] || '';
        address = cols[3] || '';
      }
    }

    name = name.replace(/^[-*•\d\.\s]+/, '').trim().toUpperCase();
    if (!name || name.length < 2) continue;
    if (/^(NAMA|NIM|MAHASISWA|NO|TOTAL|JUMLAH|KAMPUS)$/i.test(name)) continue;

    if (birthPlace.includes(',') && !birthDate) {
      const parts = birthPlace.split(',');
      birthPlace = parts[0].trim();
      birthDate = parts.slice(1).join(',').trim();
    }

    parsedStudents.push({
      nim: nim || `202601${String(parsedStudents.length + 1).padStart(2, '0')}`,
      name,
      birthPlace: birthPlace || 'Pasuruan',
      birthDate: birthDate || '1998-05-15',
      address: address || 'Pasuruan, Jawa Timur',
      gender: /perempuan|wanita|p/i.test(gender) ? 'Perempuan' : 'Laki-laki',
      phone: phone || '',
      groupId: (parsedStudents.length % 5) + 1,
      rpsPart: `Part ${String(parsedStudents.length + 1).padStart(2, '0')}`,
      topic: 'Filsafat Ilmu MPI',
      meetingNumber: (parsedStudents.length % 16) + 1,
    });
  }

  return parsedStudents;
}

// AI Detection Engine for Essay Submissions (UTS & UAS)
function detectAiEssayContent(answersText: string): {
  aiScore: number;
  verdict: 'Orisinal Mahasiswa' | 'Campuran' | 'Terindikasi AI / Copas';
  flaggedPhrases: string[];
  analysisNotes: string;
} {
  if (!answersText || answersText.trim().length < 40) {
    return {
      aiScore: 8,
      verdict: 'Orisinal Mahasiswa',
      flaggedPhrases: [],
      analysisNotes: 'Volume teks terlalu ringkas untuk analisis linguistik AI mendalam.',
    };
  }

  const lower = answersText.toLowerCase();

  // Characteristic Indonesian AI Marker Phrases (frequent in ChatGPT, Claude, Gemini, DeepSeek responses)
  const aiMarkers = [
    'secara keseluruhan',
    'penting untuk dicatat bahwa',
    'penting untuk dicatat',
    'penting untuk diingat bahwa',
    'penting untuk dipahami bahwa',
    'perlu diingat bahwa',
    'perlu dicatat bahwa',
    'perlu dipahami bahwa',
    'perlu digarisbawahi bahwa',
    'patut dicatat bahwa',
    'patut ditekankan bahwa',
    'dalam konteks ini',
    'dalam ranah ini',
    'dalam tataran ini',
    'sebagai kesimpulan',
    'kesimpulannya,',
    'sebagai penutup',
    'secara konklusif',
    'di satu sisi',
    'di sisi lain',
    'memiliki peran yang sangat krusial',
    'memiliki peran krusial',
    'berperan sangat krusial',
    'berperan krusial',
    'memainkan peranan penting',
    'memainkan peran sentral',
    'tidak dapat dipungkiri bahwa',
    'tidak dapat dimungkiri bahwa',
    'merupakan pilar utama',
    'menjadi fondasi utama',
    'merupakan fondasi utama',
    'menjadi landasan fundamental',
    'berikut adalah beberapa poin penting',
    'berikut beberapa poin penting',
    'berikut beberapa aspek utama',
    'berikut adalah penjelasan',
    'berikut uraian',
    'langkah-langkah strategis',
    'implikasi praktis dan teoritis',
    'secara komprehensif dan holistik',
    'secara komprehensif',
    'secara holistik',
    'secara garis besar',
    'dalam era digitalisasi saat ini',
    'di era digitalisasi saat ini',
    'dalam era disrupsi',
    'di era modern saat ini',
    'sangat penting untuk memahami',
    'sangat esensial untuk',
    'hal ini mencerminkan',
    'hal ini menunjukkan bahwa',
    'berakar pada prinsip',
    'berlandaskan pada filosofi',
    'sebagai model bahasa',
    'sebagai model kecerdasan buatan',
    'sebagai ai',
    'tentu, berikut adalah',
    'tentu saja, berikut adalah',
    'tentu, ini adalah',
    'semoga penjelasan ini membantu',
    'semoga jawaban ini membantu',
    'semoga bermanfaat',
    'dengan demikian, dapat disimpulkan',
    'dari uraian di atas, dapat ditarik kesimpulan',
    'oleh karena itu, sangat penting',
    'oleh sebab itu, penting bagi',
    'berperan sebagai katalisator',
    'menjembatani kesenjangan',
    'memberikan kontribusi signifikan',
  ];

  const flagged: string[] = [];
  let markerHits = 0;

  for (const marker of aiMarkers) {
    if (lower.includes(marker)) {
      if (!flagged.includes(marker)) flagged.push(marker);
      markerHits++;
    }
  }

  // 1. Detection of Markdown bold formatting: **...** (Unmistakable AI interface copy-paste signature)
  const markdownBoldMatches = answersText.match(/\*\*[^*]+\*\*/g);
  const boldCount = markdownBoldMatches ? markdownBoldMatches.length : 0;
  if (boldCount >= 2) {
    if (!flagged.includes('Format Markdown Bold (**...**) Copas AI Langsung')) {
      flagged.push('Format Markdown Bold (**...**) Copas AI Langsung');
    }
  }

  // 2. Detection of Markdown headings (### or ##)
  if (/^#{2,4}\s+/m.test(answersText)) {
    if (!flagged.includes('Format Heading Markdown (## / ###) Copas AI')) {
      flagged.push('Format Heading Markdown (## / ###) Copas AI');
    }
  }

  // 3. Detection of numbered list formatting: 1. ... 2. ... 3. ...
  const listMatches = answersText.match(/(?:^|\n)\s*(?:\d+\.|\*|-)\s+/g);
  const listCount = listMatches ? listMatches.length : 0;
  if (listCount >= 4 && !flagged.includes('Struktur Daftar Numerik/Bullet AI')) {
    flagged.push('Struktur Daftar Numerik/Bullet AI');
  }

  // 4. Sentence uniformity check (low burstiness / low length variance indicates LLM output)
  const sentences = answersText.split(/[.!?]+/).map(s => s.trim()).filter(s => s.length > 10);
  let lengthVariance = 50;
  if (sentences.length > 3) {
    const avgLen = sentences.reduce((a, b) => a + b.length, 0) / sentences.length;
    const sqDiffs = sentences.map(s => Math.pow(s.length - avgLen, 2));
    const variance = sqDiffs.reduce((a, b) => a + b, 0) / sentences.length;
    lengthVariance = Math.sqrt(variance);
  }

  // Calculate AI Score (0 - 100%)
  let calculatedScore = 5;
  calculatedScore += markerHits * 16;
  if (boldCount >= 2) calculatedScore += 35;
  if (listCount >= 3) calculatedScore += 12;
  if (sentences.length > 4 && lengthVariance < 20) calculatedScore += 15;

  // Instant triggers for overt AI disclaimers or prompt copy
  if (
    lower.includes('sebagai model kecerdasan buatan') ||
    lower.includes('sebagai model bahasa') ||
    lower.includes('sebagai ai') ||
    lower.includes('tentu, berikut ini adalah jawaban') ||
    lower.includes('tentu saja, berikut adalah') ||
    lower.includes('tentu, berikut adalah') ||
    lower.includes('semoga jawaban ini membantu') ||
    lower.includes('semoga bermanfaat!')
  ) {
    calculatedScore = Math.max(calculatedScore, 96);
  }

  const aiScore = Math.min(99, Math.max(5, Math.round(calculatedScore)));
  let verdict: 'Orisinal Mahasiswa' | 'Campuran' | 'Terindikasi AI / Copas' = 'Orisinal Mahasiswa';
  let analysisNotes = 'Pola bahasa, struktur kalimat, dan alur argumentasi menunjukkan orisinalitas pemikiran mandiri mahasiswa.';

  if (aiScore >= 60) {
    verdict = 'Terindikasi AI / Copas';
    analysisNotes = `Terdeteksi indikasi penyalinan langsung dari AI (Skor ${aiScore}%): ${markerHits} frasa transisi khas LLM, struktur enumerasi, dan/atau format markdown AI (${flagged.slice(0, 4).join(', ')}).`;
  } else if (aiScore >= 25) {
    verdict = 'Campuran';
    analysisNotes = `Terdapat indikasi parafrase atau penggunaan AI sebagai asistensi penulisan awal (${flagged.slice(0, 3).join(', ')}).`;
  }

  return {
    aiScore,
    verdict,
    flaggedPhrases: flagged,
    analysisNotes,
  };
}

// Auto-Grading Engine for 5 Essay Questions
function autoGradeEssaySubmission(
  answers: Record<number, string>,
  questions: UtsQuestion[],
  aiScore: number
): {
  totalGrade: number;
  questionScores: Record<number, number>;
  feedback: string;
} {
  const scores: Record<number, number> = {};
  let total = 0;
  const feedbackItems: string[] = [];

  const philosophicalKeywords = [
    'ontologi', 'epistemologi', 'aksiologi', 'rasionalisme', 'empirisme', 'positivisme',
    'metodologi', 'ilmiah', 'kebenaran', 'koherensi', 'korespondensi', 'etika',
    'manajemen', 'pendidikan islam', 'tata kelola', 'kritis', 'analisis', 'logika'
  ];

  const totalQuestions = questions.length > 0 ? questions.length : 5;
  const maxPerQuestion = 100 / totalQuestions;

  for (let i = 1; i <= totalQuestions; i++) {
    const q = questions.find(item => item.number === i || item.id === i);
    const ans = (answers[i] || '').trim();
    const wordCount = ans ? ans.split(/\s+/).length : 0;
    const lowerAns = ans.toLowerCase();

    let qScore = 0;

    if (wordCount < 10) {
      qScore = Math.round(maxPerQuestion * 0.25);
      feedbackItems.push(`Soal ${i}: Jawaban terlalu singkat.`);
    } else if (wordCount < 40) {
      qScore = Math.round(maxPerQuestion * 0.6);
      feedbackItems.push(`Soal ${i}: Pemahaman konsep dasar cukup, perlu penguatan dalil filosofis.`);
    } else if (wordCount < 100) {
      qScore = Math.round(maxPerQuestion * 0.8);
      const kwHits = philosophicalKeywords.filter(k => lowerAns.includes(k)).length;
      if (kwHits >= 2) qScore += 2;
      feedbackItems.push(`Soal ${i}: Uraian baik, runtut, dan selaras dengan materi perkuliahan.`);
    } else {
      qScore = Math.round(maxPerQuestion * 0.9);
      const kwHits = philosophicalKeywords.filter(k => lowerAns.includes(k)).length;
      if (kwHits >= 3) qScore = maxPerQuestion;
      feedbackItems.push(`Soal ${i}: Analisis sangat tajam, mendalam, dan komprehensif.`);
    }

    // AI Penalty if heavily copas from AI
    if (aiScore >= 80) {
      qScore = Math.max(Math.round(maxPerQuestion * 0.4), qScore - 3);
    }

    qScore = Math.min(maxPerQuestion, Math.max(0, qScore));
    scores[i] = qScore;
    total += qScore;
  }

  let finalFeedback = `Evaluasi Otomatis SIAKAD (Skor: ${total}/100).\n` + feedbackItems.join('\n');
  if (aiScore >= 65) {
    finalFeedback += `\n⚠️ Catatan Integritas Akademik: Terdeteksi indikasi AI/Copas sebesar ${aiScore}%. Dosen dapat meninjau ulang orisinalitas jawaban.`;
  } else {
    finalFeedback += `\n✅ Orisinalitas Baik: Gaya penulisan mencerminkan penalaran mandiri mahasiswa (AI: ${aiScore}%).`;
  }

  return {
    totalGrade: Math.min(100, Math.max(0, Math.round(total))),
    questionScores: scores,
    feedback: finalFeedback,
  };
}

function loadDatabase(): SiakadDatabase {
  try {
    let content = '';
    if (fs.existsSync(DATA_FILE)) {
      try {
        content = fs.readFileSync(DATA_FILE, 'utf-8');
      } catch (e) {
        console.warn('Error reading DATA_FILE, attempting backup:', e);
      }
    }
    if (!content && fs.existsSync(BACKUP_FILE)) {
      console.log('Primary DB empty/missing, recovering from BACKUP_FILE...');
      content = fs.readFileSync(BACKUP_FILE, 'utf-8');
    }

    if (content) {
      let parsed: any;
      try {
        parsed = JSON.parse(content);
      } catch (jsonErr) {
        if (fs.existsSync(BACKUP_FILE)) {
          console.warn('Primary DB corrupt JSON, recovering from BACKUP_FILE...');
          parsed = JSON.parse(fs.readFileSync(BACKUP_FILE, 'utf-8'));
        } else {
          throw jsonErr;
        }
      }

      // Build and merge allCoursesData ensuring Semester 1 & Semester 3 rooms have their own data
      const allCoursesData: Record<string, any> = {
        ...INITIAL_ALL_COURSES_DATA,
        ...(parsed.allCoursesData || {}),
      };

      // Ensure each pre-configured room is present in allCoursesData with its distinct students
      if (!allCoursesData['mk-filsafat-ilmu']) {
        allCoursesData['mk-filsafat-ilmu'] = INITIAL_ALL_COURSES_DATA['mk-filsafat-ilmu'];
      }
      // Clean up legacy test duplicate courses or unwanted courses
      const legacyTestCourseIds = new Set(['mk-manajemen-mpi', 'mk-metodologi-riset', 'mk-tqm-mpi3', 'mk-kebijakan-mpi3']);
      for (const legId of legacyTestCourseIds) {
        delete allCoursesData[legId];
      }

      // Filter courses list strictly to official 3 courses unless explicitly created by lecturer
      const officialCourseIds = new Set(['mk-filsafat-ilmu', 'mk-konsep-dasar-mpi3', 'mk-analisis-kebijakan-mpi3']);
      const isAllowedCourse = (c: any) => {
        if (!c || !c.id) return false;
        if (officialCourseIds.has(c.id)) return true;
        if (legacyTestCourseIds.has(c.id)) return false;
        // Keep if in parsed.courses or created by lecturer
        if (Array.isArray(parsed.courses) && parsed.courses.some((pc: any) => pc.id === c.id)) return true;
        return String(c.id).startsWith('mk-');
      };

      for (const key of Object.keys(allCoursesData)) {
        if (!isAllowedCourse({ id: key })) {
          delete allCoursesData[key];
        }
      }

      // Ensure mk-konsep-dasar-mpi3 is initialized if empty
      if (
        !allCoursesData['mk-konsep-dasar-mpi3'] ||
        !Array.isArray(allCoursesData['mk-konsep-dasar-mpi3'].students) ||
        allCoursesData['mk-konsep-dasar-mpi3'].students.length === 0
      ) {
        allCoursesData['mk-konsep-dasar-mpi3'] = {
          ...(allCoursesData['mk-konsep-dasar-mpi3'] || {}),
          ...INITIAL_ALL_COURSES_DATA['mk-konsep-dasar-mpi3'],
          students: JSON.parse(JSON.stringify(INITIAL_STUDENTS_SEM3)),
          groups: JSON.parse(JSON.stringify(INITIAL_GROUPS_SEM3)),
          meetings: JSON.parse(JSON.stringify(INITIAL_MEETINGS_SEM3)),
          profile: JSON.parse(JSON.stringify(INITIAL_COURSE_PROFILE_SEM3)),
        };
      }
      if (allCoursesData['mk-konsep-dasar-mpi3']?.profile) {
        allCoursesData['mk-konsep-dasar-mpi3'].profile.courseTitle = 'Konsep Dasar Manajemen Pendidikan Islam (MPI)';
        allCoursesData['mk-konsep-dasar-mpi3'].profile.semesterNumber = 3;
        allCoursesData['mk-konsep-dasar-mpi3'].profile.semester = 'Semester III (Ganjil 2026/2027)';
        allCoursesData['mk-konsep-dasar-mpi3'].profile.roomName = 'Kamar MPI 3A (Ruang Ibnu Khaldun)';
        allCoursesData['mk-konsep-dasar-mpi3'].profile.classCode = 'MPI 3A';
      }

      // Ensure mk-analisis-kebijakan-mpi3 is initialized if empty
      if (
        !allCoursesData['mk-analisis-kebijakan-mpi3'] ||
        !Array.isArray(allCoursesData['mk-analisis-kebijakan-mpi3'].students) ||
        allCoursesData['mk-analisis-kebijakan-mpi3'].students.length === 0
      ) {
        allCoursesData['mk-analisis-kebijakan-mpi3'] = {
          ...(allCoursesData['mk-analisis-kebijakan-mpi3'] || {}),
          ...INITIAL_ALL_COURSES_DATA['mk-analisis-kebijakan-mpi3'],
          students: JSON.parse(JSON.stringify(INITIAL_STUDENTS_KEBIJAKAN_SEM3)),
          groups: JSON.parse(JSON.stringify(INITIAL_GROUPS_KEBIJAKAN_SEM3)),
          meetings: JSON.parse(JSON.stringify(INITIAL_MEETINGS_KEBIJAKAN_SEM3)),
          profile: JSON.parse(JSON.stringify(INITIAL_COURSE_PROFILE_KEBIJAKAN_SEM3)),
        };
      }
      if (allCoursesData['mk-analisis-kebijakan-mpi3']?.profile) {
        allCoursesData['mk-analisis-kebijakan-mpi3'].profile.courseTitle = 'Analisis Kebijakan Pendidikan';
        allCoursesData['mk-analisis-kebijakan-mpi3'].profile.semesterNumber = 3;
        allCoursesData['mk-analisis-kebijakan-mpi3'].profile.semester = 'Semester III (Ganjil 2026/2027)';
        allCoursesData['mk-analisis-kebijakan-mpi3'].profile.roomName = 'Kamar MPI 3B (Ruang Al-Ghazali)';
        allCoursesData['mk-analisis-kebijakan-mpi3'].profile.classCode = 'MPI 3B';
      }

      // Ensure mk-filsafat-ilmu is initialized if empty
      if (
        !allCoursesData['mk-filsafat-ilmu'] ||
        !Array.isArray(allCoursesData['mk-filsafat-ilmu'].students) ||
        allCoursesData['mk-filsafat-ilmu'].students.length === 0
      ) {
        allCoursesData['mk-filsafat-ilmu'] = {
          ...(allCoursesData['mk-filsafat-ilmu'] || {}),
          ...INITIAL_ALL_COURSES_DATA['mk-filsafat-ilmu'],
          students: JSON.parse(JSON.stringify(INITIAL_STUDENTS)),
          groups: JSON.parse(JSON.stringify(INITIAL_GROUPS)),
          meetings: JSON.parse(JSON.stringify(INITIAL_MEETINGS)),
          profile: JSON.parse(JSON.stringify(INITIAL_COURSE_PROFILE)),
        };
      }
      if (allCoursesData['mk-filsafat-ilmu']?.profile) {
        allCoursesData['mk-filsafat-ilmu'].profile.courseTitle = 'Filsafat Ilmu';
        allCoursesData['mk-filsafat-ilmu'].profile.semesterNumber = 1;
        allCoursesData['mk-filsafat-ilmu'].profile.semester = 'Semester I (Ganjil 2026/2027)';
        allCoursesData['mk-filsafat-ilmu'].profile.roomName = 'Kamar MPI 1 (Ruang Al-Farabi)';
        allCoursesData['mk-filsafat-ilmu'].profile.classCode = 'MPI 1';
      }

      // Sanitize placeholder names if any
      const isPlaceholder = (s: any) => {
        if (!s) return true;
        const name = (s.name || '').toLowerCase();
        return (
          name.includes("ahmad syafi'i") ||
          name.includes("siti mutmainnah") ||
          name.includes("budi prasetyo") ||
          name.includes("rina wardani") ||
          (s.id && String(s.id).startsWith('mhs3-'))
        );
      };

      for (const [cId, cData] of Object.entries(allCoursesData)) {
        if (cData && Array.isArray(cData.students)) {
          const hasDummy = cData.students.some(isPlaceholder);
          if (hasDummy || cData.students.length === 0) {
            const semNum = cData.profile?.semesterNumber || (cId.includes('3') ? 3 : 1);
            const master = getMasterStudentsAndGroups(semNum, cData.profile?.courseTitle || '');
            cData.students = master.students;
            cData.groups = master.groups;
            if (!cData.grades || Object.keys(cData.grades).length === 0) {
              cData.grades = master.grades;
            }
          }
        }
      }

      // Merge courses list (strictly official 3 courses plus any custom ones added by dosen)
      const parsedCoursesList = (parsed.courses || []).filter(isAllowedCourse);
      const existingCourseIds = new Set(parsedCoursesList.map((c: any) => c.id));
      const coursesList: CourseSummary[] = [...parsedCoursesList];

      for (const initC of (INITIAL_DATABASE.courses || [])) {
        if (!existingCourseIds.has(initC.id)) {
          coursesList.push(JSON.parse(JSON.stringify(initC)));
          existingCourseIds.add(initC.id);
        }
      }

      // Strictly ensure correct titles, semester numbers, and room names on course summaries
      coursesList.forEach(c => {
        if (c.id === 'mk-filsafat-ilmu') {
          c.title = 'Filsafat Ilmu';
          c.code = 'STI 105';
          c.sks = 2;
          c.semesterNumber = 1;
          c.semester = 'Semester I (Ganjil 2026/2027)';
          c.roomName = 'Kamar MPI 1 (Ruang Al-Farabi)';
          c.classCode = 'MPI 1';
          c.totalStudents = 13;
        } else if (c.id === 'mk-konsep-dasar-mpi3') {
          c.title = 'Konsep Dasar Manajemen Pendidikan Islam (MPI)';
          c.code = 'MPI 301';
          c.sks = 3;
          c.semesterNumber = 3;
          c.semester = 'Semester III (Ganjil 2026/2027)';
          c.roomName = 'Kamar MPI 3A (Ruang Ibnu Khaldun)';
          c.classCode = 'MPI 3A';
          c.totalStudents = 20;
        } else if (c.id === 'mk-analisis-kebijakan-mpi3') {
          c.title = 'Analisis Kebijakan Pendidikan';
          c.code = 'MPI 302';
          c.sks = 3;
          c.semesterNumber = 3;
          c.semester = 'Semester III (Ganjil 2026/2027)';
          c.roomName = 'Kamar MPI 3B (Ruang Al-Ghazali)';
          c.classCode = 'MPI 3B';
          c.totalStudents = 20;
        } else {
          if (!c.semesterNumber) {
            c.semesterNumber = c.semester?.includes('3') || c.semester?.includes('III') ? 3 : 1;
          }
          if (allCoursesData[c.id]?.students) {
            c.totalStudents = allCoursesData[c.id].students.length;
          }
        }
      });

      // Active course selection
      const activeCourseId = parsed.activeCourseId && existingCourseIds.has(parsed.activeCourseId)
        ? parsed.activeCourseId
        : 'mk-filsafat-ilmu';

      const activeRoomData = allCoursesData[activeCourseId] || allCoursesData['mk-filsafat-ilmu'] || INITIAL_ALL_COURSES_DATA['mk-filsafat-ilmu'];

      // Active room students: MUST be strictly the students of the active room
      let rawStudents: Student[] = Array.isArray(activeRoomData.students) && activeRoomData.students.length > 0
        ? activeRoomData.students
        : (activeCourseId.includes('mpi3') ? INITIAL_STUDENTS_SEM3 : INITIAL_STUDENTS);

      let rawGroups: GroupProject[] = Array.isArray(activeRoomData.groups) && activeRoomData.groups.length > 0
        ? activeRoomData.groups
        : (activeCourseId.includes('mpi3') ? INITIAL_GROUPS_SEM3 : INITIAL_GROUPS);

      if (!Array.isArray(rawStudents) || rawStudents.length === 0) {
        const semNum = activeRoomData?.profile?.semesterNumber || (activeCourseId.includes('3') ? 3 : 1);
        const master = getMasterStudentsAndGroups(semNum, activeRoomData?.profile?.courseTitle || '');
        rawStudents = master.students;
        rawGroups = master.groups;
        if (activeRoomData) {
          activeRoomData.students = master.students;
          activeRoomData.groups = master.groups;
        }
      }

      let rawMeetings: MeetingSchedule[] = Array.isArray(activeRoomData.meetings) && activeRoomData.meetings.length === 16
        ? activeRoomData.meetings
        : (activeCourseId.includes('mpi3') ? INITIAL_MEETINGS_SEM3 : INITIAL_MEETINGS);

      const loaded: SiakadDatabase = {
        ...INITIAL_DATABASE,
        ...parsed,
        courses: coursesList,
        activeCourseId,
        allCoursesData,
        archivedSemesters: parsed.archivedSemesters || [],
        students: rawStudents,
        groups: rawGroups,
        meetings: rawMeetings,
        submissions: activeRoomData.submissions || [],
        allTimeSubmissions: parsed.allTimeSubmissions || parsed.submissions || [],
        utsQuestions: activeRoomData.utsQuestions || (activeCourseId.includes('mpi3') ? INITIAL_UTS_QUESTIONS_SEM3 : INITIAL_DATABASE.utsQuestions),
        utsSubmissions: activeRoomData.utsSubmissions || [],
        uasQuestions: activeRoomData.uasQuestions || (INITIAL_DATABASE.uasQuestions || INITIAL_UAS_QUESTIONS),
        uasSubmissions: activeRoomData.uasSubmissions || [],
        quizQuestions: activeRoomData.quizQuestions || INITIAL_QUIZ_QUESTIONS,
        quizSubmissions: activeRoomData.quizSubmissions || [],
        utsFormat: activeRoomData.utsFormat || parsed.utsFormat || 'esai',
        uasFormat: activeRoomData.uasFormat || parsed.uasFormat || 'proyek_video',
        utsExamSettings: activeRoomData.utsExamSettings || parsed.utsExamSettings || INITIAL_DATABASE.utsExamSettings,
        uasExamSettings: activeRoomData.uasExamSettings || parsed.uasExamSettings || INITIAL_DATABASE.uasExamSettings,
        attendance: activeRoomData.attendance || { 1: {} },
        grades: activeRoomData.grades || {},
        activeHeartbeats: parsed.activeHeartbeats || {},
        courseProfile: activeRoomData.profile || (activeCourseId.includes('mpi3') ? INITIAL_COURSE_PROFILE_SEM3 : INITIAL_COURSE_PROFILE),
        rpsRawText: activeRoomData.rpsRawText || parsed.rpsRawText || '',
        dosenPassword: parsed.dosenPassword || 'filsafat2026',
      };

      // Ensure profile contains requested lecturer & campus if default was legacy
      if (
        !loaded.courseProfile.dosenName ||
        loaded.courseProfile.name?.includes('Ahmad Fauzi') ||
        !loaded.courseProfile.campusName
      ) {
        loaded.courseProfile.name = 'Risfa Tri Ulfa, S.Pd., M.Pd., Gr.';
        loaded.courseProfile.dosenName = 'Risfa Tri Ulfa';
        loaded.courseProfile.dosenTitle = 'S.Pd., M.Pd., Gr.';
        loaded.courseProfile.campusName = 'STAI Jarinabi';
      }

      // Cleanse any legacy 'Part' labels to 'Pertemuan' and ensure biodata
      if (Array.isArray(loaded.students)) {
        loaded.students.forEach((s: Student) => {
          if (s.rpsPart && /part/i.test(s.rpsPart)) {
            s.rpsPart = `Pertemuan ${s.meetingNumber || 2}`;
          }
          const initMatch = INITIAL_DATABASE.students.find(i => i.id === s.id || i.nim === s.nim);
          if (!s.birthPlace) s.birthPlace = initMatch?.birthPlace || 'Pasuruan';
          if (!s.birthDate) s.birthDate = initMatch?.birthDate || '1998-05-15';
          if (!s.address) s.address = initMatch?.address || 'Kabupaten Pasuruan, Jawa Timur';
          if (!s.gender) {
            s.gender = initMatch?.gender || (
              s.name.includes('SARI') || s.name.includes('LESTARI') || s.name.includes('MAGHFIROH') ||
              s.name.includes('ALIYAH') || s.name.includes('DEFTIRIYANI') || s.name.includes('JATIMULISA')
                ? 'Perempuan'
                : 'Laki-laki'
            );
          }
          if (!s.phone) s.phone = initMatch?.phone || `08123456${(s.nim || '0000').slice(-4)}`;
          if (s.activityPoints === undefined) {
            s.activityPoints = initMatch?.activityPoints !== undefined ? initMatch.activityPoints : 25;
          }
          if (!s.activityPointLogs || !Array.isArray(s.activityPointLogs)) {
            s.activityPointLogs = initMatch?.activityPointLogs ? [...initMatch.activityPointLogs] : [];
          }
        });
      }
      if (!loaded.activityPointLogs) {
        loaded.activityPointLogs = {};
      }
      if (!loaded.quizSettings) {
        loaded.quizSettings = INITIAL_DATABASE.quizSettings || {
          isQuizActive: true,
          targetMeeting: 'Semua Pertemuan (Review RPS & Materi Komprehensif)',
          quizTitle: 'Game Cerdas Cermat RPS Pascasarjana MPI 1',
          timeLimitMinutes: 15,
          description: 'Kuis interaktif 10 soal pilihan ganda berbobot dengan pengawas kamera live HUD untuk menguji penguasaan materi RPS.',
        };
      }
      if (Array.isArray(loaded.submissions)) {
        loaded.submissions.forEach((sub: IndividualSubmission) => {
          if (sub.rpsPart && /part/i.test(sub.rpsPart)) {
            sub.rpsPart = `Pertemuan ${sub.meetingNumber || 2}`;
          }
        });
      }
      if (Array.isArray(loaded.meetings)) {
        loaded.meetings.forEach((m: MeetingSchedule) => {
          if (Array.isArray(m.partCodes)) {
            m.partCodes = m.partCodes.map((c: string) =>
              /part/i.test(c) ? `Pertemuan ${m.meetingNumber}` : c
            );
          }
        });
      }

      // Merge persistent submissions from secondary vault to ensure zero data loss
      if (fs.existsSync(SUBMISSIONS_VAULT_FILE)) {
        try {
          const vaultContent = fs.readFileSync(SUBMISSIONS_VAULT_FILE, 'utf-8');
          const vault = JSON.parse(vaultContent);
          if (Array.isArray(vault.submissions) && vault.submissions.length > 0) {
            const subMap = new Map((loaded.submissions || []).map(s => [`${s.studentId || s.id}_m${Number(s.meetingNumber) || 2}`, s]));
            for (const s of vault.submissions) {
              const k = `${s.studentId || s.id}_m${Number(s.meetingNumber) || 2}`;
              if (!subMap.has(k)) subMap.set(k, s);
            }
            loaded.submissions = Array.from(subMap.values());
          }
          if (Array.isArray(vault.utsSubmissions) && vault.utsSubmissions.length > 0) {
            const utsMap = new Map((loaded.utsSubmissions || []).map(u => [u.studentId || u.id, u]));
            for (const u of vault.utsSubmissions) {
              const k = u.studentId || u.id;
              if (!utsMap.has(k)) utsMap.set(k, u);
            }
            loaded.utsSubmissions = Array.from(utsMap.values());
          }
          if (Array.isArray(vault.uasSubmissions) && vault.uasSubmissions.length > 0) {
            const uasMap = new Map((loaded.uasSubmissions || []).map(u => [u.studentId || u.id, u]));
            for (const u of vault.uasSubmissions) {
              const k = u.studentId || u.id;
              if (!uasMap.has(k)) uasMap.set(k, u);
            }
            loaded.uasSubmissions = Array.from(uasMap.values());
          }
          if (Array.isArray(vault.quizSubmissions) && vault.quizSubmissions.length > 0) {
            const quizMap = new Map((loaded.quizSubmissions || []).map(q => [q.studentId || q.id, q]));
            for (const q of vault.quizSubmissions) {
              const k = q.studentId || q.id;
              if (!quizMap.has(k)) quizMap.set(k, q);
            }
            loaded.quizSubmissions = Array.from(quizMap.values());
          }
          if (Array.isArray(vault.archivedSemesters) && vault.archivedSemesters.length > 0 && (!loaded.archivedSemesters || loaded.archivedSemesters.length === 0)) {
            loaded.archivedSemesters = vault.archivedSemesters;
          }
        } catch (vErr) {
          console.warn('Submissions vault merge non-fatal:', vErr);
        }
      }

      return loaded;
    }
  } catch (err) {
    console.error('Failed to load database file, using fallback initial data:', err);
  }
  return JSON.parse(JSON.stringify(INITIAL_DATABASE));
}

function saveDatabase(): void {
  try {
    db.lastUpdated = new Date().toISOString();

    // Ensure all changes to active course are mirrored into allCoursesData for cross-semester persistence
    if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
      if (db.students) db.allCoursesData[db.activeCourseId].students = db.students;
      if (db.meetings) db.allCoursesData[db.activeCourseId].meetings = db.meetings;
      if (db.groups) db.allCoursesData[db.activeCourseId].groups = db.groups;
      if (db.attendance) db.allCoursesData[db.activeCourseId].attendance = db.attendance;
      if (db.grades) db.allCoursesData[db.activeCourseId].grades = db.grades;
      if (db.courseProfile) db.allCoursesData[db.activeCourseId].profile = db.courseProfile;
      if (db.submissions) db.allCoursesData[db.activeCourseId].submissions = db.submissions;
      if (db.utsSubmissions) db.allCoursesData[db.activeCourseId].utsSubmissions = db.utsSubmissions;
      if (db.uasSubmissions) db.allCoursesData[db.activeCourseId].uasSubmissions = db.uasSubmissions;
      if (db.quizSubmissions) db.allCoursesData[db.activeCourseId].quizSubmissions = db.quizSubmissions;
    }

    const jsonStr = JSON.stringify(db, null, 2);
    // 1. Atomic write using temporary file to prevent partial write or corruption
    try {
      const tmpFile = `${DATA_FILE}.tmp.${Date.now()}`;
      fs.writeFileSync(tmpFile, jsonStr, 'utf-8');
      fs.renameSync(tmpFile, DATA_FILE);
    } catch {
      fs.writeFileSync(DATA_FILE, jsonStr, 'utf-8');
    }

    // 2. Rolling backup copy
    try {
      fs.writeFileSync(BACKUP_FILE, jsonStr, 'utf-8');
    } catch (bErr) {
      console.warn('Backup file update non-fatal warning:', bErr);
    }

    // 3. Permanent all-time submissions vault (cross-session & cross-semester persistence)
    try {
      const vaultPayload = {
        lastSaved: new Date().toISOString(),
        totalSubmissions: db.submissions?.length || 0,
        submissions: db.submissions || [],
        utsSubmissions: db.utsSubmissions || [],
        uasSubmissions: db.uasSubmissions || [],
        quizSubmissions: db.quizSubmissions || [],
        archivedSemesters: db.archivedSemesters || [],
      };
      fs.writeFileSync(SUBMISSIONS_VAULT_FILE, JSON.stringify(vaultPayload, null, 2), 'utf-8');
    } catch (vErr) {
      console.warn('Submissions vault write non-fatal:', vErr);
    }
  } catch (err) {
    console.error('Failed to save database file:', err);
  }
}

db = loadDatabase();

// Durable upload storage directory for all documents (PPT, Makalah, UTS, UAS, Quiz)
const UPLOAD_DIR = path.join(DATA_DIR, 'uploads');
try {
  if (!fs.existsSync(UPLOAD_DIR)) {
    fs.mkdirSync(UPLOAD_DIR, { recursive: true });
  }
} catch (e) {
  console.warn('Notice: read-only filesystem or upload directory exists:', e);
}
app.use('/uploads', express.static(UPLOAD_DIR));
app.use('/api/uploads', express.static(UPLOAD_DIR));

// Helper: If data is base64 Data URL, save it to disk and return permanent URL
function saveBase64FileToDisk(dataUrlOrPath?: string, preferredName?: string): { url?: string; fileName?: string } {
  if (!dataUrlOrPath) return {};
  if (dataUrlOrPath.startsWith('/uploads/') || dataUrlOrPath.startsWith('http://') || dataUrlOrPath.startsWith('https://')) {
    return { url: dataUrlOrPath, fileName: preferredName };
  }
  if (!dataUrlOrPath.startsWith('data:')) {
    return { url: dataUrlOrPath, fileName: preferredName };
  }

  try {
    if (!fs.existsSync(UPLOAD_DIR)) {
      try {
        fs.mkdirSync(UPLOAD_DIR, { recursive: true });
      } catch {}
    }
    const matches = dataUrlOrPath.match(/^data:([A-Za-z0-9+/=;.-]+);base64,(.+)$/);
    if (!matches || matches.length < 3) {
      return { url: dataUrlOrPath, fileName: preferredName };
    }
    const mime = matches[1];
    const base64Data = matches[2];
    const buffer = Buffer.from(base64Data, 'base64');
    
    let ext = '.bin';
    if (preferredName && path.extname(preferredName)) {
      ext = path.extname(preferredName);
    } else if (mime.includes('pdf')) ext = '.pdf';
    else if (mime.includes('presentation') || mime.includes('powerpoint')) ext = '.pptx';
    else if (mime.includes('word') || mime.includes('document')) ext = '.docx';
    else if (mime.includes('png')) ext = '.png';
    else if (mime.includes('jpeg') || mime.includes('jpg')) ext = '.jpg';
    else if (mime.includes('text')) ext = '.txt';
    else if (mime.includes('video')) ext = '.mp4';

    const safeBaseName = (preferredName ? path.basename(preferredName, ext) : 'dokumen')
      .replace(/[^a-zA-Z0-9_-]/g, '_')
      .slice(0, 40);
    const uniqueFileName = `${Date.now()}_${Math.random().toString(36).substr(2, 6)}_${safeBaseName}${ext}`;
    const filePath = path.join(UPLOAD_DIR, uniqueFileName);
    fs.writeFileSync(filePath, buffer);
    const publicUrl = `/uploads/${uniqueFileName}`;
    return { url: publicUrl, fileName: preferredName || uniqueFileName };
  } catch (err) {
    console.error('Failed to save base64 file to disk:', err);
    return { url: dataUrlOrPath, fileName: preferredName };
  }
}

// Clean up stale heartbeats older than 10 minutes periodically
setInterval(() => {
  const now = Date.now();
  let changed = false;
  if (db.activeHeartbeats) {
    for (const [studentId, timeStr] of Object.entries(db.activeHeartbeats)) {
      const diff = now - new Date(timeStr).getTime();
      if (diff > 15 * 60 * 1000) {
        delete db.activeHeartbeats[studentId];
        changed = true;
      }
    }
  }
  if (changed) {
    saveDatabase();
  }
}, 60 * 1000);

// ================= API ROUTES =================

// 1. Get entire database state (auto-aligns with client's active course if requested)
app.get(['/api/db', '/db'], (req, res) => {
  const reqCourseId = (req.query.courseId as string) || (req.headers['x-active-course-id'] as string);
  if (reqCourseId && reqCourseId !== db.activeCourseId) {
    try {
      switchActiveCourseInternal(reqCourseId);
    } catch (e) {
      console.warn('Auto-align active course notice in /api/db:', e);
    }
  }

  const { dosenPassword, ...safeDb } = db;
  res.json({
    success: true,
    data: safeDb,
    serverTime: new Date().toISOString(),
  });
});

// 2. Student presence heartbeat (marks student as ONLINE)
app.post('/api/presence', (req, res) => {
  const { studentId, studentName } = req.body;
  if (!studentId) {
    return res.status(400).json({ error: 'studentId required' });
  }

  const nowIso = new Date().toISOString();
  if (!db.activeHeartbeats) {
    db.activeHeartbeats = {};
  }
  db.activeHeartbeats[studentId] = nowIso;

  // Also update student's lastActive
  const student = (db.students || []).find(s => s.id === studentId);
  if (student) {
    student.lastActive = nowIso;
  }

  saveDatabase();
  res.json({ success: true, timestamp: nowIso });
});

// Direct file upload endpoint (persists file to disk in data/uploads)
app.post(['/api/upload', '/upload'], (req, res) => {
  try {
    const { fileName, fileData } = req.body;
    if (!fileData) {
      return res.status(400).json({ error: 'Data file tidak boleh kosong' });
    }
    const saved = saveBase64FileToDisk(fileData, fileName);
    const finalUrl = saved.url || fileData;
    res.json({
      success: true,
      fileUrl: finalUrl,
      fileName: saved.fileName || fileName,
      message: 'File dokumen berhasil disimpan permanen di server SIAKAD.',
    });
  } catch (err: any) {
    res.status(500).json({ error: err?.message || 'Gagal mengunggah file' });
  }
});

// 3. Submit individual assignment (Makalah & PPT)
app.post('/api/submissions', (req, res) => {
  const submissionData: IndividualSubmission = req.body;
  if (!submissionData.studentId) {
    return res.status(400).json({ error: 'Missing studentId' });
  }

  // If base64 file data is uploaded, save to disk to prevent data loss or oversized JSON
  let pptUrl = submissionData.pptUrl;
  let pptFileData = submissionData.pptFileData;
  if (submissionData.pptFileData && submissionData.pptFileData.startsWith('data:')) {
    const savedPpt = saveBase64FileToDisk(submissionData.pptFileData, submissionData.pptFileName || 'presentasi_ppt');
    if (savedPpt.url) {
      pptUrl = savedPpt.url;
      pptFileData = savedPpt.url;
    }
  }

  let makalahUrl = submissionData.makalahUrl;
  let makalahFileData = submissionData.makalahFileData;
  if (submissionData.makalahFileData && submissionData.makalahFileData.startsWith('data:')) {
    const savedMakalah = saveBase64FileToDisk(submissionData.makalahFileData, submissionData.makalahFileName || 'makalah_tugas');
    if (savedMakalah.url) {
      makalahUrl = savedMakalah.url;
      makalahFileData = savedMakalah.url;
    }
  }

  // 1. Identify if this is a Group Task or Individual Task
  const targetMeetingNumber = Number(submissionData.meetingNumber) || 2;
  const isGroupSubmission = submissionData.presentationType === 'kelompok' || Boolean(submissionData.groupId);

  const MEETING_TO_GROUP: Record<number, number> = {
    3: 1, 4: 2, 5: 3, 6: 4, 7: 5, 9: 6, 10: 7, 12: 8, 13: 9, 14: 10
  };

  let targetGroup = submissionData.groupId ? (db.groups || []).find(g => g.id === Number(submissionData.groupId)) : undefined;
  if (!targetGroup && isGroupSubmission) {
    const mappedGId = MEETING_TO_GROUP[targetMeetingNumber];
    if (mappedGId) {
      targetGroup = (db.groups || []).find(g => g.id === mappedGId);
    }
  }

  // Check existing submission:
  // If Group: check if any submission exists for this group in this meeting
  // If Individual: check strictly for this studentId in this meeting
  let existingIndex = -1;
  let existingGroupSub: IndividualSubmission | undefined = undefined;

  if (isGroupSubmission && targetGroup) {
    const groupMemberNames = new Set((targetGroup.members || []).map(m => m.trim().toLowerCase()));
    const groupMemberIds = new Set(
      (db.students || [])
        .filter(s => groupMemberNames.has(s.name.trim().toLowerCase()) || s.groupId === targetGroup?.id)
        .map(s => s.id)
    );
    groupMemberIds.add(submissionData.studentId);

    existingIndex = db.submissions.findIndex(s =>
      (Number(s.meetingNumber) || 2) === targetMeetingNumber && (
        (s.groupId && s.groupId === targetGroup?.id) ||
        (s.id && submissionData.id && s.id === submissionData.id) ||
        groupMemberIds.has(s.studentId) ||
        groupMemberNames.has((s.studentName || '').trim().toLowerCase())
      )
    );
    if (existingIndex >= 0) {
      existingGroupSub = db.submissions[existingIndex];
    }
  } else {
    existingIndex = db.submissions.findIndex(
      s => (s.id && submissionData.id && s.id === submissionData.id) ||
           (s.studentId === submissionData.studentId && (Number(s.meetingNumber) || 2) === targetMeetingNumber)
    );
  }

  // Satu Kali Kirim:
  // If student is not dosen, prevent submitting if already submitted
  if (existingIndex >= 0 && !isDosenAuthorized(req)) {
    if (isGroupSubmission && existingGroupSub && existingGroupSub.studentId !== submissionData.studentId) {
      return res.status(403).json({
        error: `Tugas Kelompok (${targetGroup?.name || 'Kelompok'}) Pertemuan ${targetMeetingNumber} sudah berhasil dikirimkan sebelumnya oleh perwakilan kelompok (${existingGroupSub.studentName}). Sesuai aturan akademik, hanya 1 perwakilan saja yang mengirimkan tugas kelompok, anggota lainnya tidak perlu mengirim ulang.`
      });
    }

    return res.status(403).json({
      error: `Tugas ${isGroupSubmission ? 'kelompok' : 'individu'} Pertemuan ${targetMeetingNumber} sudah berhasil dikirimkan sebelumnya (Mode Satu Kali Kirim). Formulir pengiriman terkunci. Mahasiswa hanya dapat mengunggah ulang jika Dosen Pengampu telah mereset/menghapus tugas sebelumnya untuk keperluan perbaikan (revisi).`
    });
  }

  const updatedSubmission: IndividualSubmission = {
    ...submissionData,
    meetingNumber: targetMeetingNumber,
    id: (existingIndex >= 0 && existingGroupSub ? existingGroupSub.id : submissionData.id) || `sub-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
    groupId: targetGroup?.id || submissionData.groupId,
    groupName: targetGroup?.name || submissionData.groupName,
    groupMembers: targetGroup?.members || submissionData.groupMembers,
    representativeName: submissionData.representativeName || submissionData.studentName,
    presentationType: isGroupSubmission ? 'kelompok' : 'individu',
    pptUrl,
    pptFileData,
    makalahType: submissionData.makalahType || (submissionData.isHardcopyPrint ? 'hardcopy_print' : (makalahUrl || makalahFileData ? 'file' : undefined)),
    isHardcopyPrint: Boolean(submissionData.isHardcopyPrint || submissionData.makalahType === 'hardcopy_print'),
    hardcopyReceivedAt: submissionData.hardcopyReceivedAt || (submissionData.isHardcopyPrint || submissionData.makalahType === 'hardcopy_print' ? new Date().toISOString() : undefined),
    hardcopyReceiver: submissionData.hardcopyReceiver || (submissionData.isHardcopyPrint || submissionData.makalahType === 'hardcopy_print' ? (db.courseProfile?.dosenName || 'Dosen Pengampu') : undefined),
    makalahUrl: submissionData.makalahType === 'hardcopy_print' ? undefined : makalahUrl,
    makalahFileData: submissionData.makalahType === 'hardcopy_print' ? undefined : makalahFileData,
    submittedAt: new Date().toISOString(),
  };

  if (existingIndex >= 0) {
    // Keep grade & feedback if existing was already graded
    const prev = db.submissions[existingIndex];
    if (prev.grade !== undefined && updatedSubmission.grade === undefined) {
      updatedSubmission.grade = prev.grade;
      updatedSubmission.feedback = prev.feedback;
      updatedSubmission.gradedAt = prev.gradedAt;
    }
    db.submissions[existingIndex] = updatedSubmission;
  } else {
    db.submissions.push(updatedSubmission);
  }

  // Sync with academic grades in SIAKAD so it's recorded permanently
  if (!db.grades) db.grades = {};
  if (!db.grades[submissionData.studentId]) {
    db.grades[submissionData.studentId] = {
      attendanceScore: 100,
      attitudeScore: 85,
      individualScore: updatedSubmission.grade !== undefined ? updatedSubmission.grade : 85,
      utsScore: 85,
      uasScore: 85,
      groupScore: 85,
      finalScore: 85,
      letterGrade: 'A',
      notes: `Tugas Terkumpul: ${updatedSubmission.topic}`,
    };
  }

  // If group submission, sync notes to all members of the group
  if (isGroupSubmission && targetGroup) {
    const memberNames = new Set((targetGroup.members || []).map(m => m.trim().toLowerCase()));
    (db.students || []).forEach(member => {
      if (memberNames.has(member.name.trim().toLowerCase()) || member.groupId === targetGroup?.id) {
        if (!db.grades[member.id]) {
          db.grades[member.id] = {
            attendanceScore: 100,
            attitudeScore: 85,
            individualScore: updatedSubmission.grade !== undefined ? updatedSubmission.grade : 85,
            utsScore: 85,
            uasScore: 85,
            groupScore: 85,
            finalScore: 85,
            letterGrade: 'A',
            notes: `Tugas Kelompok Terkumpul (${targetGroup.name}): ${updatedSubmission.topic}`,
          };
        }
      }
    });
  } else if (updatedSubmission.grade !== undefined) {
    db.grades[submissionData.studentId].individualScore = updatedSubmission.grade;
  }

  // If group presentation format, sync to co-presenters of this meeting
  if (updatedSubmission.presentationType === 'kelompok' && updatedSubmission.meetingNumber) {
    const coPresenters = (db.students || []).filter(
      s => (Number(s.meetingNumber) || 2) === targetMeetingNumber && s.id !== submissionData.studentId
    );
    coPresenters.forEach(cp => {
      const pIdx = db.submissions.findIndex(
        s => s.studentId === cp.id && (Number(s.meetingNumber) || 2) === targetMeetingNumber
      );
      const partnerSub: IndividualSubmission = {
        ...updatedSubmission,
        id: pIdx >= 0 ? db.submissions[pIdx].id : `sub-partner-${Date.now()}-${cp.id}`,
        studentId: cp.id,
        studentName: cp.name,
        rpsPart: cp.rpsPart || updatedSubmission.rpsPart,
        topic: cp.topic || updatedSubmission.topic,
        partnerName: submissionData.studentName,
        meetingNumber: targetMeetingNumber,
      };
      if (pIdx >= 0) {
        if (db.submissions[pIdx].grade !== undefined && partnerSub.grade === undefined) {
          partnerSub.grade = db.submissions[pIdx].grade;
          partnerSub.feedback = db.submissions[pIdx].feedback;
          partnerSub.gradedAt = db.submissions[pIdx].gradedAt;
        }
        db.submissions[pIdx] = partnerSub;
      } else {
        db.submissions.push(partnerSub);
      }
    });
  }

  // Update student presence as well
  if (!db.activeHeartbeats) db.activeHeartbeats = {};
  db.activeHeartbeats[submissionData.studentId] = new Date().toISOString();

  // Sync to course data cache
  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].submissions = db.submissions;
    db.allCoursesData[db.activeCourseId].grades = db.grades;
  }

  saveDatabase();
  res.json({ success: true, submission: updatedSubmission });
});

// 3B. Konfirmasi Makalah Cetak / Hardcopy oleh Dosen (Tanpa Perlu Upload File)
app.post(['/api/submissions/confirm-print', '/submissions/confirm-print'], (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berwenang mengonfirmasi pengumpulan makalah cetak!' });
  }

  const { studentId, meetingNumber, grade, feedback, notes, isHardcopyPrint = true } = req.body;
  if (!studentId) {
    return res.status(400).json({ error: 'studentId wajib disertakan' });
  }

  const std = (db.students || []).find(s => s.id === studentId || s.nim === studentId);
  if (!std) {
    return res.status(404).json({ error: 'Mahasiswa tidak ditemukan dalam basis data' });
  }

  const targetMeeting = meetingNumber !== undefined ? Number(meetingNumber) : (std.meetingNumber || 2);
  if (!db.submissions) db.submissions = [];

  let sub = db.submissions.find(s =>
    s.studentId === std.id && (Number(s.meetingNumber) || 2) === targetMeeting
  );

  const dosenName = db.courseProfile?.dosenName || 'Dosen Pengampu';
  const nowIso = new Date().toISOString();

  if (sub) {
    sub.isHardcopyPrint = Boolean(isHardcopyPrint);
    if (isHardcopyPrint) {
      sub.makalahType = 'hardcopy_print';
      sub.hardcopyReceivedAt = nowIso;
      sub.hardcopyReceiver = dosenName;
      if (notes) {
        sub.notes = notes;
      } else if (!sub.notes) {
        sub.notes = 'Makalah telah dikumpulkan dalam bentuk berkas cetak fisik (print out) langsung ke Dosen Pengampu.';
      }
    } else {
      sub.makalahType = sub.makalahFileData ? 'file' : (sub.makalahUrl ? 'link' : undefined);
    }
    if (grade !== undefined && grade !== null && grade !== '') {
      sub.grade = Number(grade);
      sub.gradedAt = nowIso;
    }
    if (feedback !== undefined) {
      sub.feedback = feedback;
    }
  } else {
    // Buat data pengumpulan baru untuk mencatat berkas cetak fisik tanpa file digital
    sub = {
      id: `sub-print-${Date.now()}-${std.id}`,
      studentId: std.id,
      studentName: std.name,
      rpsPart: std.rpsPart || `Pertemuan ${targetMeeting}`,
      topic: std.topic || 'Tugas Makalah & Presentasi RPS',
      meetingNumber: targetMeeting,
      presentationType: std.groupId ? 'kelompok' : 'individu',
      groupId: std.groupId,
      pptType: 'link',
      makalahType: 'hardcopy_print',
      isHardcopyPrint: true,
      hardcopyReceivedAt: nowIso,
      hardcopyReceiver: dosenName,
      notes: notes || 'Makalah telah dikumpulkan dalam bentuk berkas cetak fisik (print out) langsung ke Dosen Pengampu.',
      submittedAt: nowIso,
      grade: (grade !== undefined && grade !== null && grade !== '') ? Number(grade) : undefined,
      feedback: feedback || '',
      gradedAt: (grade !== undefined && grade !== null && grade !== '') ? nowIso : undefined,
    };
    db.submissions.push(sub);
  }

  // Sinkronkan nilai akademik jika dosen langsung menginput nilai
  if (sub.grade !== undefined && !isNaN(sub.grade)) {
    if (!db.grades) db.grades = {};
    if (!db.grades[std.id]) {
      db.grades[std.id] = {
        attendanceScore: 100,
        attitudeScore: 88,
        individualScore: sub.grade,
        utsScore: 85,
        uasScore: 85,
        groupScore: 85,
        finalScore: sub.grade,
        letterGrade: 'A',
        notes: sub.feedback || 'Nilai Makalah Cetak (Hardcopy)',
      };
    } else {
      db.grades[std.id].individualScore = sub.grade;
      if (sub.feedback) db.grades[std.id].notes = sub.feedback;
    }
    recalculateStudentGrade(db.grades[std.id]);
  }

  // Sync to course cache for cross-session persistence
  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].submissions = db.submissions;
    db.allCoursesData[db.activeCourseId].grades = db.grades;
  }

  saveDatabase();

  const { dosenPassword, ...safeDb } = db;
  res.json({
    success: true,
    message: isHardcopyPrint
      ? `Makalah cetak fisik untuk ${std.name} berhasil dikonfirmasi diterima!`
      : `Status cetak fisik untuk ${std.name} berhasil dibatalkan.`,
    submission: sub,
    data: safeDb,
  });
});

// 4. Delete submission or part (Only Dosen)
app.delete('/api/submissions/:id', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang menghapus data tugas.' });
  }
  const { id } = req.params;
  const initialLength = (db.submissions || []).length;
  const meetingNum = req.query.meetingNumber ? Number(req.query.meetingNumber) : null;
  db.submissions = (db.submissions || []).filter(s => {
    if (s.id === id) return false;
    if (s.studentId === id) {
      if (meetingNum !== null) return (Number(s.meetingNumber) || 2) !== meetingNum;
      return false;
    }
    return true;
  });
  saveDatabase();
  res.json({ success: true, removedCount: initialLength - db.submissions.length });
});

app.post('/api/submissions/:id/delete-part', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang menghapus bagian tugas.' });
  }
  const { id } = req.params;
  const { part = 'all', reason = '', meetingNumber } = req.body; // 'all' | 'ppt' | 'makalah'
  
  if (!db.submissions) db.submissions = [];
  const meetingNum = meetingNumber ? Number(meetingNumber) : null;
  const subIndex = db.submissions.findIndex(s =>
    s.id === id ||
    (s.studentId === id && (meetingNum === null || (Number(s.meetingNumber) || 2) === meetingNum))
  );
  if (subIndex === -1) {
    return res.status(404).json({ error: 'Submission not found' });
  }

  const sub = db.submissions[subIndex];

  if (part === 'all') {
    db.submissions.splice(subIndex, 1);
    saveDatabase();
    return res.json({ success: true, part: 'all', deletedEntire: true });
  } else if (part === 'ppt') {
    sub.pptType = 'link';
    sub.pptUrl = '';
    sub.pptFileName = '';
    sub.pptFileData = '';
    if (reason) sub.feedback = reason;

    // Check if makalah is also empty
    const hasMakalah = Boolean(
      (sub.makalahType === 'link' && sub.makalahUrl) ||
      (sub.makalahType === 'file' && sub.makalahFileData) ||
      sub.isHardcopyPrint ||
      sub.makalahType === 'hardcopy_print'
    );
    if (!hasMakalah) {
      db.submissions.splice(subIndex, 1);
      saveDatabase();
      return res.json({ success: true, part: 'ppt', deletedEntire: true });
    }
  } else if (part === 'makalah') {
    sub.makalahType = 'link';
    sub.makalahUrl = '';
    sub.makalahFileName = '';
    sub.makalahFileData = '';
    sub.isHardcopyPrint = false;
    sub.hardcopyReceivedAt = undefined;
    sub.hardcopyReceiver = undefined;
    if (reason) sub.feedback = reason;

    // Check if ppt is also empty
    const hasPpt = Boolean((sub.pptType === 'link' && sub.pptUrl) || (sub.pptType === 'file' && sub.pptFileData));
    if (!hasPpt) {
      db.submissions.splice(subIndex, 1);
      saveDatabase();
      return res.json({ success: true, part: 'makalah', deletedEntire: true });
    }
  }

  saveDatabase();
  res.json({ success: true, part, submission: sub, deletedEntire: false });
});

// 4b. Peer review on individual presentation submission (Mahasiswa mereview & memberi tanggapan ke rekan)
app.post('/api/submissions/:id/peer-review', (req, res) => {
  const { id } = req.params;
  const { reviewerId, reviewerName, rating, comment } = req.body;

  if (!reviewerName || !comment) {
    return res.status(400).json({ error: 'Nama penilai dan komentar ulasan wajib diisi' });
  }

  if (!db.submissions) db.submissions = [];
  const sub = db.submissions.find(s => s.id === id || s.studentId === id);
  if (!sub) {
    return res.status(404).json({ error: 'Tugas presentasi tidak ditemukan di database' });
  }

  if (!sub.peerReviews) sub.peerReviews = [];

  const newReview = {
    id: `pr-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
    reviewerId: reviewerId || 'anon',
    reviewerName: reviewerName.trim(),
    rating: Math.min(5, Math.max(1, Number(rating) || 5)),
    comment: comment.trim(),
    createdAt: new Date().toISOString(),
  };

  sub.peerReviews.unshift(newReview);
  saveDatabase();
  res.json({ success: true, review: newReview, totalReviews: sub.peerReviews.length });
});

// 5. Submit group project video
app.post('/api/group-submissions', (req, res) => {
  const { groupId, videoUrl, canvaUrl, driveUrl, aiToolsUsed, summaryNotes, submittedBy, fileName, fileData } = req.body;
  const group = db.groups.find(g => g.id === Number(groupId));
  if (!group) {
    return res.status(404).json({ error: 'Group not found' });
  }

  // Check if UAS is open by lecturer (unless lecturer) only for official UAS submissions
  if (!isDosenAuthorized(req) && (req.body.isUasExam || req.body.isExam) && db.uasExamSettings && db.uasExamSettings.isOpen === false) {
    return res.status(403).json({
      error: 'Ujian Akhir Semester (UAS) saat ini berstatus Ditutup / Dikunci oleh Dosen Pengampu.',
    });
  }

  // Satu Kali Kirim: Prevent student from overwriting existing group submission
  if (group.submission && (group.submission.videoUrl || group.submission.submittedAt) && !isDosenAuthorized(req)) {
    return res.status(403).json({
      error: 'Proyek Video UAS Kelompok ini sudah dikumpulkan sebelumnya (Satu Kali Kirim). Formulir pengiriman terkunci. Mahasiswa hanya dapat mengumpulkan ulang jika Dosen telah mereset status pengumpulan kelompok ini untuk keperluan perbaikan (revisi).',
    });
  }

  let processedFileData = fileData;
  let fileDiskUrl = '';
  if (fileData && typeof fileData === 'string' && fileData.startsWith('data:')) {
    const saved = saveBase64FileToDisk(fileData, fileName || `kelompok-${groupId}-uas`);
    if (saved.url) {
      fileDiskUrl = saved.url;
      processedFileData = saved.url;
    }
  }

  group.submission = {
    videoUrl,
    canvaUrl,
    driveUrl: driveUrl || fileDiskUrl,
    aiToolsUsed: aiToolsUsed || '',
    summaryNotes: summaryNotes || '',
    submittedBy: submittedBy || 'Anggota Kelompok',
    submittedAt: new Date().toISOString(),
    fileName,
    fileData: processedFileData,
  };

  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].groups = db.groups;
  }

  saveDatabase();
  res.json({ success: true, group });
});

// 5b. Reset group project submission (Dosen only, for revision)
app.post('/api/group-submissions/:groupId/reset', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berwenang mereset pengumpulan proyek kelompok.' });
  }
  const { groupId } = req.params;
  const group = db.groups?.find(g => g.id === Number(groupId));
  if (!group) {
    return res.status(404).json({ error: 'Kelompok tidak ditemukan.' });
  }
  group.submission = undefined;
  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].groups = db.groups;
  }
  saveDatabase();
  res.json({ success: true, message: `Status pengumpulan proyek untuk ${group.name} berhasil direset. Akses pengumpulan kini terbuka kembali bagi mahasiswa.` });
});

// 6. Grade group project (Tugas Presentasi / UTS / UAS Video Kelompok) (Dosen)
app.post('/api/group-grade', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang memiliki hak akses untuk memberikan nilai proyek kelompok!' });
  }

  const { groupId, grade, feedback, examType = 'uas', autoGraded = false } = req.body;
  const group = db.groups.find(g => g.id === Number(groupId));
  if (!group) {
    return res.status(404).json({ error: 'Group not found' });
  }

  group.grade = Math.min(100, Math.max(0, Number(grade) || 85));
  group.feedback = feedback || (autoGraded ? 'Dinilai otomatis berbasis rubrik standar RPS' : '');
  group.gradedAt = new Date().toISOString();

  // Also update groupScore & uasScore/utsScore/individualScore for each student in this group
  const memberNames = (group.members || []).map(m => m.trim().toLowerCase());
  db.students.forEach(std => {
    if (memberNames.includes(std.name.trim().toLowerCase())) {
      if (!db.grades[std.id]) {
        db.grades[std.id] = {
          attendanceScore: 100,
          attitudeScore: 85,
          individualScore: 85,
          utsScore: 85,
          uasScore: 85,
          groupScore: Number(group.grade),
          finalScore: 88,
          letterGrade: 'A-',
        };
      }
      if (examType === 'uts') {
        db.grades[std.id].utsScore = Number(group.grade);
      } else if (examType === 'presentation') {
        db.grades[std.id].groupScore = Number(group.grade);
        db.grades[std.id].individualScore = Number(group.grade);
      } else {
        db.grades[std.id].groupScore = Number(group.grade);
        db.grades[std.id].uasScore = Number(group.grade);
      }
      recalculateStudentGrade(db.grades[std.id]);
    }
  });

  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].groups = db.groups;
    db.allCoursesData[db.activeCourseId].grades = db.grades;
  }

  saveDatabase();
  res.json({ success: true, group });
});

// 7. Grade individual task (PPT & Makalah) (Dosen)
app.post('/api/individual-grade', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang memiliki hak akses untuk memberikan nilai tugas individu!' });
  }

  const { studentId, submissionId, meetingNumber, grade, feedback } = req.body;
  if (!studentId) {
    return res.status(400).json({ error: 'studentId wajib disertakan' });
  }

  if (!db.submissions) db.submissions = [];
  let sub = (db.submissions || []).find(s =>
    (submissionId && s.id === submissionId) ||
    (s.studentId === studentId && meetingNumber !== undefined && (Number(s.meetingNumber) || 2) === Number(meetingNumber))
  );
  if (!sub && !meetingNumber && !submissionId) {
    // Fallback search by studentId only if meetingNumber/submissionId not specified
    sub = (db.submissions || []).find(s => s.studentId === studentId);
  }

  if (sub) {
    sub.grade = Number(grade);
    sub.feedback = feedback || '';
    sub.gradedAt = new Date().toISOString();
    if (req.body.isHardcopyPrint === true) {
      sub.isHardcopyPrint = true;
      sub.makalahType = 'hardcopy_print';
      sub.hardcopyReceivedAt = sub.hardcopyReceivedAt || new Date().toISOString();
      sub.hardcopyReceiver = sub.hardcopyReceiver || (db.courseProfile?.dosenName || 'Dosen Pengampu');
    } else if (req.body.isHardcopyPrint === false) {
      sub.isHardcopyPrint = false;
      if (sub.makalahType === 'hardcopy_print') {
        sub.makalahType = sub.makalahFileData ? 'file' : (sub.makalahUrl ? 'link' : undefined);
      }
    }
  } else {
    // If student hadn't submitted a formal file yet, create submission record so grade is preserved permanently
    const std = (db.students || []).find(s => s.id === studentId || s.nim === studentId);
    const targetMeeting = meetingNumber !== undefined ? Number(meetingNumber) : (std?.meetingNumber || 2);
    sub = {
      id: submissionId || `sub-${Date.now()}-${studentId}`,
      studentId,
      studentName: std ? std.name : 'Mahasiswa',
      rpsPart: std?.rpsPart || `Pertemuan ${targetMeeting}`,
      topic: std?.topic || 'Tugas Presentasi RPS',
      meetingNumber: targetMeeting,
      presentationType: 'individu',
      pptType: 'link',
      makalahType: req.body.isHardcopyPrint ? 'hardcopy_print' : 'link',
      isHardcopyPrint: Boolean(req.body.isHardcopyPrint),
      hardcopyReceivedAt: req.body.isHardcopyPrint ? new Date().toISOString() : undefined,
      hardcopyReceiver: req.body.isHardcopyPrint ? (db.courseProfile?.dosenName || 'Dosen Pengampu') : undefined,
      submittedAt: new Date().toISOString(),
      grade: Number(grade),
      feedback: feedback || 'Dinilai oleh Dosen Pengampu',
      gradedAt: new Date().toISOString(),
    };
    db.submissions.push(sub);
  }

  // Calculate composite individual score (average if multiple presentations for students appearing twice)
  const gradedStudentSubs = (db.submissions || []).filter(
    s => s.studentId === studentId && s.grade !== undefined && !isNaN(Number(s.grade))
  );
  const calculatedIndivScore = gradedStudentSubs.length > 0
    ? Math.round(gradedStudentSubs.reduce((acc, curr) => acc + Number(curr.grade || 0), 0) / gradedStudentSubs.length)
    : Number(grade);

  if (!db.grades) db.grades = {};
  if (!db.grades[studentId]) {
    db.grades[studentId] = {
      attendanceScore: 100,
      attitudeScore: 85,
      individualScore: calculatedIndivScore,
      utsScore: 85,
      uasScore: 85,
      groupScore: 85,
      finalScore: 88,
      letterGrade: 'A-',
      notes: feedback || `Nilai Tugas Presentasi: ${calculatedIndivScore}`,
    };
  } else {
    db.grades[studentId].individualScore = calculatedIndivScore;
    if (feedback) db.grades[studentId].notes = feedback;
  }

  recalculateStudentGrade(db.grades[studentId]);

  // If this submission is for a presentation group, also apply the score to all members of that group
  if (sub.presentationType === 'kelompok' && sub.groupId) {
    const group = (db.groups || []).find(g => g.id === sub.groupId);
    if (group && group.members) {
      const memberNames = new Set(group.members.map(m => m.trim().toLowerCase()));
      (db.students || []).forEach(member => {
        if (memberNames.has(member.name.trim().toLowerCase()) || member.groupId === group.id) {
          if (!db.grades[member.id]) {
            db.grades[member.id] = {
              attendanceScore: 100,
              attitudeScore: 85,
              individualScore: calculatedIndivScore,
              utsScore: 85,
              uasScore: 85,
              groupScore: 85,
              finalScore: 88,
              letterGrade: 'A-',
              notes: feedback || `Nilai Tugas Presentasi ${group.name}: ${calculatedIndivScore}`,
            };
          } else {
            db.grades[member.id].individualScore = calculatedIndivScore;
            if (feedback) db.grades[member.id].notes = feedback;
          }
          recalculateStudentGrade(db.grades[member.id]);
        }
      });
    }
  }

  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].submissions = db.submissions;
    db.allCoursesData[db.activeCourseId].grades = db.grades;
  }

  saveDatabase();
  res.json({ success: true, submission: sub, studentGrade: db.grades[studentId] });
});

// 7b. Submit UTS answers (Mahasiswa - 5 Soal Essay dengan Deteksi AI & Penilaian Otomatis)
app.post('/api/uts-submissions', (req, res) => {
  const { studentId, studentName, answers, docLink, fileName, fileData } = req.body;
  if (!studentId) {
    return res.status(400).json({ error: 'Missing studentId' });
  }

  // Check if UTS is open by lecturer (students can only take exam when lecturer has opened it)
  if (!isDosenAuthorized(req) && db.utsExamSettings && db.utsExamSettings.isOpen === false) {
    return res.status(403).json({
      error: 'Ujian Tengah Semester (UTS) saat ini berstatus Ditutup / Belum Dibuka oleh Dosen Pengampu.',
    });
  }

  if (!db.utsSubmissions) {
    db.utsSubmissions = [];
  }

  // Save file to disk if base64 to ensure durability
  let finalDocLink = docLink || '';
  let finalFileData = fileData;
  if (fileData && fileData.startsWith('data:')) {
    const saved = saveBase64FileToDisk(fileData, fileName || 'dokumen_uts');
    if (saved.url) {
      finalDocLink = saved.url;
      finalFileData = saved.url;
    }
  }

  // Concatenate answers for AI evaluation
  const answersRecord = (answers || {}) as Record<number, string>;
  const combinedAnswers = Object.values(answersRecord).filter(Boolean).join(' ');
  const aiDetection = detectAiEssayContent(combinedAnswers);

  const utsQuestions = (db.utsQuestions && db.utsQuestions.length > 0)
    ? db.utsQuestions
    : INITIAL_DATABASE.utsQuestions;
  const autoGrading = autoGradeEssaySubmission(answersRecord, utsQuestions, aiDetection.aiScore);

  const existingIdx = db.utsSubmissions.findIndex(u => u.studentId === studentId);
  const existingSub = existingIdx >= 0 ? db.utsSubmissions[existingIdx] : null;

  // Satu Kali Kirim: Prevent student from resubmitting once already submitted unless Dosen authorized/reset
  if (existingSub && !isDosenAuthorized(req)) {
    return res.status(403).json({
      error: 'Lembar Jawaban UTS Anda sudah berhasil dikumpulkan sebelumnya (Satu Kali Kirim). Formulir pengumpulan terkunci. Mahasiswa hanya dapat mengumpulkan ulang jika Dosen Pengampu telah mereset/menghapus tugas sebelumnya untuk keperluan perbaikan (revisi).'
    });
  }

  // Determine final grade and feedback (if already manually graded by lecturer, preserve; otherwise use auto-grade)
  const finalGrade = (existingSub && existingSub.grade !== undefined && !existingSub.autoGraded)
    ? existingSub.grade
    : autoGrading.totalGrade;
  const finalScores = (existingSub && existingSub.questionScores && !existingSub.autoGraded)
    ? existingSub.questionScores
    : autoGrading.questionScores;
  const finalFeedback = (existingSub && existingSub.feedback && !existingSub.autoGraded)
    ? existingSub.feedback
    : autoGrading.feedback;

  const updatedSubmission: UtsSubmission = {
    id: existingSub ? existingSub.id : `uts-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
    studentId,
    studentName: studentName || 'Mahasiswa MPI 1',
    submittedAt: new Date().toISOString(),
    answers: answersRecord,
    docLink: finalDocLink,
    fileName,
    fileData: finalFileData,
    aiDetectionScore: aiDetection.aiScore,
    aiVerdict: aiDetection.verdict,
    aiDetectedFlags: aiDetection.flaggedPhrases,
    aiAnalysisNotes: aiDetection.analysisNotes,
    autoGraded: true,
    grade: finalGrade,
    questionScores: finalScores,
    feedback: finalFeedback,
    gradedAt: new Date().toISOString(),
  };

  if (existingIdx >= 0) {
    db.utsSubmissions[existingIdx] = updatedSubmission;
  } else {
    db.utsSubmissions.push(updatedSubmission);
  }

  // Automatically synchronize UTS grade with SIAKAD academic record
  if (!db.grades) db.grades = {};
  if (!db.grades[studentId]) {
    db.grades[studentId] = {
      attendanceScore: 100,
      attitudeScore: 85,
      individualScore: 85,
      utsScore: finalGrade,
      uasScore: 85,
      groupScore: 85,
      finalScore: 88,
      letterGrade: 'A-',
    };
  } else {
    db.grades[studentId].utsScore = finalGrade;
  }
  recalculateStudentGrade(db.grades[studentId]);

  // Update presence
  if (!db.activeHeartbeats) db.activeHeartbeats = {};
  db.activeHeartbeats[studentId] = new Date().toISOString();

  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].utsSubmissions = db.utsSubmissions;
    db.allCoursesData[db.activeCourseId].grades = db.grades;
  }

  saveDatabase();
  res.json({ success: true, submission: updatedSubmission, studentGrade: db.grades[studentId] });
});

// 7b-2. Delete / Reset UTS submission (Only Dosen, allows student to resubmit for revision)
app.delete('/api/uts-submissions/:studentId', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berwenang mereset atau menghapus tugas UTS mahasiswa.' });
  }
  const { studentId } = req.params;
  const initialCount = (db.utsSubmissions || []).length;
  db.utsSubmissions = (db.utsSubmissions || []).filter(u => u.studentId !== studentId && u.id !== studentId);
  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].utsSubmissions = db.utsSubmissions;
  }
  saveDatabase();
  res.json({ success: true, removedCount: initialCount - (db.utsSubmissions || []).length });
});

// 7c. Grade UTS submission (Dosen - 5 Soal Essay)
app.post('/api/uts-grade', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang memiliki hak akses untuk memberikan nilai UTS!' });
  }

  const { studentId, grade, questionScores, feedback } = req.body;
  if (!studentId) {
    return res.status(400).json({ error: 'Missing studentId' });
  }

  if (!db.utsSubmissions) db.utsSubmissions = [];
  let sub = (db.utsSubmissions || []).find(u => u.studentId === studentId);
  if (!sub) {
    const student = (db.students || []).find(s => s.id === studentId);
    sub = {
      id: `uts-${Date.now()}`,
      studentId,
      studentName: student ? student.name : 'Mahasiswa',
      submittedAt: new Date().toISOString(),
      answers: {},
    };
    db.utsSubmissions.push(sub);
  }

  sub.grade = Number(grade);
  if (questionScores) sub.questionScores = questionScores;
  sub.feedback = feedback || '';
  sub.autoGraded = false;
  sub.gradedBy = 'Dosen Pengampu';
  sub.gradedAt = new Date().toISOString();

  // Update grades record
  if (!db.grades[studentId]) {
    db.grades[studentId] = {
      attendanceScore: 100,
      attitudeScore: 85,
      individualScore: 85,
      utsScore: Number(grade),
      uasScore: 85,
      groupScore: 85,
      finalScore: 88,
      letterGrade: 'A-',
    };
  } else {
    db.grades[studentId].utsScore = Number(grade);
  }

  recalculateStudentGrade(db.grades[studentId]);

  saveDatabase();
  res.json({ success: true, submission: sub, grade: db.grades[studentId] });
});

// 7c-2. Save UTS Draft (Autosave berkala draf lembar jawaban essay UTS mahasiswa)
app.post('/api/uts/draft', (req, res) => {
  const { studentId, answers, docLink, studentName } = req.body;
  if (!studentId) {
    return res.status(400).json({ error: 'studentId wajib disertakan' });
  }

  if (!db.utsDrafts) {
    db.utsDrafts = {};
  }

  const student = (db.students || []).find(s => s.id === studentId || s.nim === studentId);
  const nowIso = new Date().toISOString();

  const draft: UtsDraft = {
    studentId,
    studentName: studentName || student?.name || 'Mahasiswa',
    answers: answers || {},
    docLink: docLink || '',
    savedAt: nowIso,
  };

  db.utsDrafts[studentId] = draft;

  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    if (!db.allCoursesData[db.activeCourseId].utsDrafts) {
      db.allCoursesData[db.activeCourseId].utsDrafts = {};
    }
    db.allCoursesData[db.activeCourseId].utsDrafts![studentId] = draft;
  }

  saveDatabase();
  res.json({ success: true, savedAt: nowIso, draft });
});

// 7c-3. Get UTS Draft
app.get('/api/uts/draft/:studentId', (req, res) => {
  const { studentId } = req.params;
  const draft = db.utsDrafts?.[studentId] || null;
  res.json({ success: true, draft });
});

// 7d. Save UTS Questions (Dosen: Add, Edit, Remove questions)
app.post('/api/uts/questions', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak mengedit dan menambah soal UTS!' });
  }

  const { questions } = req.body;
  if (!Array.isArray(questions)) {
    return res.status(400).json({ error: 'Format data soal tidak valid' });
  }

  db.utsQuestions = questions;
  saveDatabase();
  res.json({ success: true, questions: db.utsQuestions });
});

// 7d-2. Sync & Auto-generate UTS Essay Questions from RPS Meetings 1-7
app.post('/api/uts/sync-from-rps', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak menyinkronkan soal UTS!' });
  }

  const generated = generateUtsQuestionsFromMeetings(db.meetings || [], db.courseProfile);
  db.utsQuestions = generated;
  saveDatabase();

  const cTitle = db.courseProfile?.name || db.courseProfile?.courseTitle || 'Mata Kuliah Aktif';
  res.json({
    success: true,
    message: `5 Soal Essay UTS berhasil disinkronkan otomatis dari materi RPS Pertemuan 1-7 (${cTitle})!`,
    questions: db.utsQuestions,
  });
});

// 7e. Save UAS Questions (Dosen: Add, Edit, Remove questions)
app.post('/api/uas/questions', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak mengedit dan menambah soal UAS!' });
  }

  const { questions } = req.body;
  if (!Array.isArray(questions)) {
    return res.status(400).json({ error: 'Format data soal tidak valid' });
  }

  db.uasQuestions = questions;
  saveDatabase();
  res.json({ success: true, questions: db.uasQuestions });
});

// 7e-2. Sync & Auto-generate UAS Essay Questions from RPS Meetings 9-15
app.post('/api/uas/sync-from-rps', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak menyinkronkan soal UAS!' });
  }

  const generated = generateUasQuestionsFromMeetings(db.meetings || [], db.courseProfile);
  db.uasQuestions = generated;
  saveDatabase();

  const cTitle = db.courseProfile?.name || db.courseProfile?.courseTitle || 'Mata Kuliah Aktif';
  res.json({
    success: true,
    message: `5 Soal Essay UAS berhasil disinkronkan otomatis dari materi RPS Pertemuan 9-15 (${cTitle})!`,
    questions: db.uasQuestions,
  });
});

// 7f. Update Exam Formats (UTS & UAS: esai | proyek_video)
app.post('/api/exam-format', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak menentukan format UTS dan UAS!' });
  }

  const { utsFormat, uasFormat } = req.body;
  if (utsFormat) db.utsFormat = utsFormat;
  if (uasFormat) db.uasFormat = uasFormat;

  if (db.courseProfile) {
    if (utsFormat) db.courseProfile.utsFormat = utsFormat;
    if (uasFormat) db.courseProfile.uasFormat = uasFormat;
  }

  saveDatabase();
  res.json({ success: true, utsFormat: db.utsFormat, uasFormat: db.uasFormat });
});

// 7g. Submit UAS Essay Submission (Mahasiswa - Individu dengan Deteksi AI & Penilaian Otomatis)
app.post('/api/uas-submissions', (req, res) => {
  const { studentId, studentName, answers, docLink, fileName, fileData } = req.body;
  if (!studentId) {
    return res.status(400).json({ error: 'Missing studentId' });
  }

  // Check if UAS is open by lecturer (students can only take exam when lecturer has opened it)
  if (!isDosenAuthorized(req) && db.uasExamSettings && db.uasExamSettings.isOpen === false) {
    return res.status(403).json({
      error: 'Ujian Akhir Semester (UAS) saat ini berstatus Ditutup / Belum Dibuka oleh Dosen Pengampu.',
    });
  }

  if (!db.uasSubmissions) db.uasSubmissions = [];

  // Save file to disk if base64 to ensure durability
  let finalDocLink = docLink || undefined;
  let finalFileData = fileData || undefined;
  if (fileData && fileData.startsWith('data:')) {
    const saved = saveBase64FileToDisk(fileData, fileName || 'dokumen_uas');
    if (saved.url) {
      finalDocLink = saved.url;
      finalFileData = saved.url;
    }
  }

  const answersRecord = (answers || {}) as Record<number, string>;
  const combinedAnswers = Object.values(answersRecord).filter(Boolean).join(' ');
  const aiDetection = detectAiEssayContent(combinedAnswers);

  const uasQuestions = (db.uasQuestions && db.uasQuestions.length > 0)
    ? db.uasQuestions
    : (INITIAL_DATABASE.uasQuestions || INITIAL_UAS_QUESTIONS);
  const autoGrading = autoGradeEssaySubmission(answersRecord, uasQuestions, aiDetection.aiScore);

  const existingIdx = db.uasSubmissions.findIndex(u => u.studentId === studentId);
  const existingSub = existingIdx >= 0 ? db.uasSubmissions[existingIdx] : null;

  // Satu Kali Kirim: Prevent student from resubmitting once already submitted unless Dosen authorized/reset
  if (existingSub && !isDosenAuthorized(req)) {
    return res.status(403).json({
      error: 'Lembar Jawaban UAS Anda sudah berhasil dikumpulkan sebelumnya (Satu Kali Kirim). Formulir pengumpulan terkunci. Mahasiswa hanya dapat mengumpulkan ulang jika Dosen Pengampu telah mereset/menghapus tugas sebelumnya untuk keperluan perbaikan (revisi).'
    });
  }

  const finalGrade = (existingSub && existingSub.grade !== undefined && !existingSub.autoGraded)
    ? existingSub.grade
    : autoGrading.totalGrade;
  const finalScores = (existingSub && existingSub.questionScores && !existingSub.autoGraded)
    ? existingSub.questionScores
    : autoGrading.questionScores;
  const finalFeedback = (existingSub && existingSub.feedback && !existingSub.autoGraded)
    ? existingSub.feedback
    : autoGrading.feedback;

  const updatedSubmission: UtsSubmission = {
    id: existingSub ? existingSub.id : `uas-${Date.now()}`,
    studentId,
    studentName: studentName || 'Mahasiswa',
    submittedAt: new Date().toISOString(),
    answers: answersRecord,
    docLink: finalDocLink,
    fileName: fileName || undefined,
    fileData: finalFileData,
    aiDetectionScore: aiDetection.aiScore,
    aiVerdict: aiDetection.verdict,
    aiDetectedFlags: aiDetection.flaggedPhrases,
    aiAnalysisNotes: aiDetection.analysisNotes,
    autoGraded: true,
    grade: finalGrade,
    questionScores: finalScores,
    feedback: finalFeedback,
    gradedAt: new Date().toISOString(),
  };

  if (existingIdx >= 0) {
    db.uasSubmissions[existingIdx] = updatedSubmission;
  } else {
    db.uasSubmissions.push(updatedSubmission);
  }

  // Automatically synchronize UAS grade with SIAKAD academic record
  if (!db.grades) db.grades = {};
  if (!db.grades[studentId]) {
    db.grades[studentId] = {
      attendanceScore: 100,
      attitudeScore: 85,
      individualScore: 85,
      utsScore: 85,
      uasScore: finalGrade,
      groupScore: 85,
      finalScore: 88,
      letterGrade: 'A-',
    };
  } else {
    db.grades[studentId].uasScore = finalGrade;
  }
  recalculateStudentGrade(db.grades[studentId]);

  if (!db.activeHeartbeats) db.activeHeartbeats = {};
  db.activeHeartbeats[studentId] = new Date().toISOString();

  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].uasSubmissions = db.uasSubmissions;
    db.allCoursesData[db.activeCourseId].grades = db.grades;
  }

  saveDatabase();
  res.json({ success: true, submission: updatedSubmission, studentGrade: db.grades[studentId] });
});

// 7g-2. Delete / Reset UAS submission (Only Dosen, allows student to resubmit for revision)
app.delete('/api/uas-submissions/:studentId', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berwenang mereset atau menghapus tugas UAS mahasiswa.' });
  }
  const { studentId } = req.params;
  const initialCount = (db.uasSubmissions || []).length;
  db.uasSubmissions = (db.uasSubmissions || []).filter(u => u.studentId !== studentId && u.id !== studentId);
  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].uasSubmissions = db.uasSubmissions;
  }
  saveDatabase();
  res.json({ success: true, removedCount: initialCount - (db.uasSubmissions || []).length });
});

// 7h. Grade UAS Submission (Dosen)
app.post('/api/uas-grade', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang memiliki hak akses untuk memberikan nilai UAS!' });
  }

  const { studentId, grade, questionScores, feedback } = req.body;
  if (!studentId) {
    return res.status(400).json({ error: 'Missing studentId' });
  }

  if (!db.uasSubmissions) db.uasSubmissions = [];
  let sub = (db.uasSubmissions || []).find(u => u.studentId === studentId);
  if (!sub) {
    const student = (db.students || []).find(s => s.id === studentId);
    sub = {
      id: `uas-${Date.now()}`,
      studentId,
      studentName: student ? student.name : 'Mahasiswa',
      submittedAt: new Date().toISOString(),
      answers: {},
    };
    db.uasSubmissions.push(sub);
  }

  sub.grade = Number(grade);
  if (questionScores) sub.questionScores = questionScores;
  sub.feedback = feedback || '';
  sub.autoGraded = false;
  sub.gradedBy = 'Dosen Pengampu';
  sub.gradedAt = new Date().toISOString();

  // Update grades record
  if (!db.grades[studentId]) {
    db.grades[studentId] = {
      attendanceScore: 100,
      attitudeScore: 85,
      individualScore: 85,
      utsScore: 85,
      uasScore: Number(grade),
      groupScore: Number(grade),
      finalScore: 88,
      letterGrade: 'A-',
    };
  } else {
    db.grades[studentId].uasScore = Number(grade);
    db.grades[studentId].groupScore = Number(grade);
  }

  recalculateStudentGrade(db.grades[studentId]);

  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].uasSubmissions = db.uasSubmissions;
    db.allCoursesData[db.activeCourseId].grades = db.grades;
  }

  saveDatabase();
  res.json({ success: true, submission: sub, grade: db.grades[studentId] });
});

// 7h-1a. Complete Comprehensive Semester Backup (ZIP Archive)
app.get('/api/backup/zip', async (req, res) => {
  try {
    const zip = new JSZip();
    const courseTitle = db.courseProfile?.courseTitle || 'Filsafat Ilmu';
    const semester = db.courseProfile?.semester || 'Semester Ganjil 2026-2027';
    const dateStr = new Date().toISOString().slice(0, 10);

    // 1. Documentation & Metadata
    const summaryText = `===================================================================
ARSIP CADANGAN LENGKAP SEMESTER - SIAKAD PASCASARJANA STAI JARINABI
===================================================================
Mata Kuliah      : ${courseTitle} (${db.courseProfile?.courseCode || 'MPI-501'})
Dosen Pengampu   : ${db.courseProfile?.name || 'Risfa Tri Ulfa, S.Pd., M.Pd., Gr.'}
Semester         : ${semester}
Tahun Akademik   : ${db.courseProfile?.academicYear || '2026/2027'}
Waktu Pembuatan  : ${new Date().toLocaleString('id-ID')}
Total Mahasiswa  : ${db.students?.length || 0} orang
Tugas Presentasi : ${db.submissions?.length || 0} berkas terkumpul
Lembar UTS       : ${db.utsSubmissions?.length || 0} berkas terkumpul
Proyek UAS       : ${db.uasSubmissions?.length || 0} berkas terkumpul
Kuis Interaktif  : ${db.quizSubmissions?.length || 0} jawaban terkumpul
===================================================================
CATATAN PENTING:
Arsip ZIP ini memuat seluruh data akademik mahasiswa per semester,
mencakup basis data lengkap, rekapitulasi nilai mahasiswa, daftar tugas
presentasi, lembar jawaban UTS & UAS, serta berkas fisik yang diunggah.
`;
    zip.file("README_ARSIP_SEMESTER.txt", summaryText);

    // 2. Full JSON Database Dump (excluding dosen password)
    const { dosenPassword, ...safeDb } = db;
    zip.file("database_siakad_semester.json", JSON.stringify(safeDb, null, 2));

    // 3. Spreadsheet CSV & HTML of Student Grades & Submissions
    let csv = "NIM,Nama Mahasiswa,Bagian RPS,Kelompok,Nilai Presentasi,Nilai UTS,Nilai UAS,Nilai Kuis,Kehadiran,Nilai Akhir,Grade Huruf,Status Kelulusan\n";
    let htmlTableRows = "";
    for (const std of (db.students || [])) {
      const g: any = db.grades?.[std.id] || {};
      const sub = db.submissions?.find(s => s.studentId === std.id);
      const uts = db.utsSubmissions?.find(u => u.studentId === std.id);
      const uas = db.uasSubmissions?.find(u => u.studentId === std.id);
      const quiz = db.quizSubmissions?.find(q => q.studentId === std.id);
      const finalScore = g.finalScore ?? '-';
      const letterGrade = g.letterGrade || g.finalGrade || '-';
      const isPassed = typeof finalScore === 'number' && finalScore >= 60 ? 'LULUS' : 'EVALUASI';
      const presScore = sub?.grade ?? g.individualTaskScore ?? g.individualScore ?? '-';
      const utsScore = uts?.grade ?? g.utsScore ?? '-';
      const uasScore = uas?.grade ?? g.uasScore ?? '-';
      const quizScore = quiz?.score ?? g.quizScore ?? '-';
      const attScore = g.attendancePercentage ?? 100;

      csv += `"${std.nim}","${std.name}","${std.rpsPart}","Kelompok ${std.groupId}",${presScore},${utsScore},${uasScore},${quizScore},${attScore}%,${finalScore},"${letterGrade}","${isPassed}"\n`;

      htmlTableRows += `<tr>
        <td style="padding:8px;border:1px solid #e2e8f0;font-family:monospace;">${std.nim}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;font-weight:600;">${std.name}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;">${std.rpsPart}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;">${presScore}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;">${utsScore}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;">${uasScore}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;">${quizScore}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;font-weight:bold;background:#f8fafc;">${finalScore}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;font-weight:bold;color:${isPassed === 'LULUS' ? '#047857' : '#b91c1c'};">${letterGrade}</td>
        <td style="padding:8px;border:1px solid #e2e8f0;text-align:center;font-weight:bold;">${isPassed}</td>
      </tr>`;
    }
    zip.file("rekap_nilai_mahasiswa.csv", csv);

    const htmlReport = `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <title>Rekap Nilai Semester - ${courseTitle}</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 24px; color: #0f172a; }
    h2, h3 { margin: 0 0 6px 0; }
    table { width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 12px; }
    th { background: #064e3b; color: white; padding: 10px 8px; text-align: left; border: 1px solid #064e3b; }
    .meta { font-size: 13px; color: #475569; margin-bottom: 20px; border-bottom: 2px solid #064e3b; padding-bottom: 12px; }
    @media print { .no-print { display: none; } }
  </style>
</head>
<body>
  <div class="no-print" style="margin-bottom: 16px;">
    <button onclick="window.print()" style="background:#064e3b;color:white;border:none;padding:8px 16px;border-radius:6px;font-weight:bold;cursor:pointer;">
      Cetak / Simpan PDF
    </button>
  </div>
  <h2>SIAKAD PASCASARJANA - REKAPITULASI NILAI AKHIR SEMESTER</h2>
  <div class="meta">
    <div><strong>Mata Kuliah:</strong> ${courseTitle} (${db.courseProfile?.courseCode || 'MPI-501'})</div>
    <div><strong>Dosen Pengampu:</strong> ${db.courseProfile?.name || 'Risfa Tri Ulfa, S.Pd., M.Pd., Gr.'}</div>
    <div><strong>Semester / Tahun:</strong> ${semester} (${db.courseProfile?.academicYear || '2026/2027'})</div>
    <div><strong>Waktu Cetak:</strong> ${new Date().toLocaleString('id-ID')}</div>
  </div>
  <table>
    <thead>
      <tr>
        <th>NIM</th>
        <th>Nama Mahasiswa</th>
        <th>Part RPS</th>
        <th style="text-align:center;">Presentasi</th>
        <th style="text-align:center;">UTS</th>
        <th style="text-align:center;">UAS</th>
        <th style="text-align:center;">Kuis</th>
        <th style="text-align:center;">Nilai Akhir</th>
        <th style="text-align:center;">Mutu</th>
        <th style="text-align:center;">Status</th>
      </tr>
    </thead>
    <tbody>
      ${htmlTableRows}
    </tbody>
  </table>
</body>
</html>`;
    zip.file("rekap_nilai_mahasiswa.html", htmlReport);

    // 4. Folder: tugas_presentasi
    const presFolder = zip.folder("tugas_presentasi");
    const presBerkasFolder = presFolder?.folder("berkas_mhs");
    let presIndex = "REKAP PENGUMPULAN TUGAS PRESENTASI MAHASISWA:\n\n";
    for (const sub of (db.submissions || [])) {
      const std = db.students?.find(s => s.id === sub.studentId);
      const safeStdName = (std?.name || 'mhs').replace(/[^a-zA-Z0-9]/g, '_');
      const nim = std?.nim || 'mhs';

      presIndex += `• [${sub.rpsPart || std?.rpsPart || '-'}] ${sub.studentName || std?.name} (NIM: ${nim})\n`;
      presIndex += `  Topik           : ${sub.topic || std?.topic || '-'}\n`;
      presIndex += `  Format          : ${sub.presentationType === 'kelompok' ? 'Kelompok PPT' : 'Individu'}\n`;
      presIndex += `  Waktu Dikirim   : ${sub.submittedAt || '-'}\n`;
      presIndex += `  Nilai Dosen     : ${sub.grade !== undefined ? sub.grade : 'Belum Dinilai'}\n`;
      if (sub.pptUrl) presIndex += `  Link PPT/Canva  : ${sub.pptUrl}\n`;
      if (sub.makalahUrl) presIndex += `  Link Makalah    : ${sub.makalahUrl}\n`;
      if (sub.notes) presIndex += `  Catatan Ringkas : ${sub.notes}\n`;
      if (sub.feedback) presIndex += `  Feedback Dosen  : ${sub.feedback}\n`;
      presIndex += "\n";

      presFolder?.file(
        `${nim}_${safeStdName}_presentasi.txt`,
        `TUGAS PRESENTASI MAHASISWA
Nama Mahasiswa : ${std?.name}
NIM            : ${nim}
Bagian RPS     : ${sub.rpsPart || std?.rpsPart}
Topik Kajian   : ${sub.topic || std?.topic}
Format         : ${sub.presentationType || 'individu'}
Rekan Kelompok : ${sub.partnerName || '-'}
Waktu Kirim    : ${sub.submittedAt || '-'}
Nilai Dosen    : ${sub.grade !== undefined ? sub.grade : 'Belum dinilai'}
Catatan Dosen  : ${sub.feedback || '-'}

Akses Berkas:
Link PPT/Canva : ${sub.pptUrl || 'File diunggah'}
Link Makalah   : ${sub.makalahUrl || 'File diunggah'}

Catatan Mahasiswa:
${sub.notes || '-'}`
      );

      // Extract uploaded PPT if available
      if (sub.pptFileData) {
        if (sub.pptFileData.startsWith('data:')) {
          const m = sub.pptFileData.match(/^data:([A-Za-z0-9+/=;.-]+);base64,(.+)$/);
          if (m && m[2]) {
            const ext = sub.pptFileName ? path.extname(sub.pptFileName) : '.pptx';
            presBerkasFolder?.file(`${nim}_${safeStdName}_PPT${ext}`, Buffer.from(m[2], 'base64'));
          }
        } else if (sub.pptFileData.startsWith('/uploads/')) {
          const localPath = path.join(UPLOAD_DIR, path.basename(sub.pptFileData));
          if (fs.existsSync(localPath)) {
            const ext = path.extname(localPath) || '.pptx';
            presBerkasFolder?.file(`${nim}_${safeStdName}_PPT${ext}`, fs.readFileSync(localPath));
          }
        }
      }

      // Extract uploaded Makalah if available
      if (sub.makalahFileData) {
        if (sub.makalahFileData.startsWith('data:')) {
          const m = sub.makalahFileData.match(/^data:([A-Za-z0-9+/=;.-]+);base64,(.+)$/);
          if (m && m[2]) {
            const ext = sub.makalahFileName ? path.extname(sub.makalahFileName) : '.pdf';
            presBerkasFolder?.file(`${nim}_${safeStdName}_Makalah${ext}`, Buffer.from(m[2], 'base64'));
          }
        } else if (sub.makalahFileData.startsWith('/uploads/')) {
          const localPath = path.join(UPLOAD_DIR, path.basename(sub.makalahFileData));
          if (fs.existsSync(localPath)) {
            const ext = path.extname(localPath) || '.pdf';
            presBerkasFolder?.file(`${nim}_${safeStdName}_Makalah${ext}`, fs.readFileSync(localPath));
          }
        }
      }
    }
    presFolder?.file("00_DAFTAR_PENGUMPULAN_PRESENTASI.txt", presIndex);

    // 5. Folder: ujian_uts
    const utsFolder = zip.folder("ujian_uts");
    let utsIndex = "REKAP UJIAN TENGAH SEMESTER (UTS):\n\n";
    for (const uts of (db.utsSubmissions || [])) {
      const std = db.students?.find(s => s.id === uts.studentId);
      utsIndex += `• ${uts.studentName || std?.name} (NIM: ${std?.nim || '-'}) - Nilai: ${uts.grade ?? '-'}\n`;
      utsIndex += `  Waktu Kirim : ${uts.submittedAt || '-'}\n`;
      if (uts.docLink) utsIndex += `  Link Dokumen: ${uts.docLink}\n`;
      utsIndex += "\n";

      let answersText = `LEMBAR JAWABAN UJIAN TENGAH SEMESTER (UTS)
Nama Mahasiswa : ${uts.studentName || std?.name}
NIM            : ${std?.nim || '-'}
Waktu Kirim    : ${uts.submittedAt}
Nilai Dosen    : ${uts.grade ?? 'Belum dinilai'}
Catatan Dosen  : ${uts.feedback || '-'}
Link Dokumen   : ${uts.docLink || '-'}

==================================================
JAWABAN SOAL UTS:
==================================================
`;
      if (uts.answers && typeof uts.answers === 'object') {
        for (const [qKey, answer] of Object.entries(uts.answers)) {
          answersText += `\n[SOAL ${qKey}]\n${answer}\n`;
        }
      }
      utsFolder?.file(
        `${std?.nim || 'uts'}_${(uts.studentName || 'mhs').replace(/[^a-zA-Z0-9]/g, '_')}_jawaban_uts.txt`,
        answersText
      );

      // Extract UTS uploaded file if present
      if (uts.fileData) {
        if (uts.fileData.startsWith('data:')) {
          const m = uts.fileData.match(/^data:([A-Za-z0-9+/=;.-]+);base64,(.+)$/);
          if (m && m[2]) {
            const ext = uts.fileName ? path.extname(uts.fileName) : '.pdf';
            utsFolder?.file(`${std?.nim || 'uts'}_${(uts.studentName || 'mhs').replace(/[^a-zA-Z0-9]/g, '_')}_lampiran${ext}`, Buffer.from(m[2], 'base64'));
          }
        } else if (uts.fileData.startsWith('/uploads/')) {
          const localPath = path.join(UPLOAD_DIR, path.basename(uts.fileData));
          if (fs.existsSync(localPath)) {
            const ext = path.extname(localPath) || '.pdf';
            utsFolder?.file(`${std?.nim || 'uts'}_${(uts.studentName || 'mhs').replace(/[^a-zA-Z0-9]/g, '_')}_lampiran${ext}`, fs.readFileSync(localPath));
          }
        }
      }
    }
    utsFolder?.file("00_REKAP_PENGUMPULAN_UTS.txt", utsIndex);

    // 6. Folder: ujian_uas
    const uasFolder = zip.folder("ujian_uas");
    let uasIndex = "REKAP UJIAN AKHIR SEMESTER (UAS):\n\n";
    for (const uasItem of (db.uasSubmissions || [])) {
      const uas: any = uasItem;
      const std = db.students?.find(s => s.id === uas.studentId);
      uasIndex += `• ${uas.studentName || std?.name} (NIM: ${std?.nim || '-'}) - Nilai: ${uas.grade ?? '-'}\n`;
      uasIndex += `  Waktu Kirim    : ${uas.submittedAt || '-'}\n`;
      if (uas.videoUrl) uasIndex += `  Link Video UAS : ${uas.videoUrl}\n`;
      if (uas.docLink) uasIndex += `  Link Dokumen   : ${uas.docLink}\n`;
      uasIndex += "\n";

      uasFolder?.file(
        `${std?.nim || 'uas'}_${(uas.studentName || 'mhs').replace(/[^a-zA-Z0-9]/g, '_')}_uas.txt`,
        `LEMBAR PENGUMPULAN UJIAN AKHIR SEMESTER (UAS)
Nama Mahasiswa : ${uas.studentName || std?.name}
NIM            : ${std?.nim || '-'}
Waktu Kirim    : ${uas.submittedAt}
Nilai Dosen    : ${uas.grade ?? 'Belum dinilai'}
Catatan Dosen  : ${uas.feedback || '-'}
Link Video UAS : ${uas.videoUrl || '-'}
Link Dokumen   : ${uas.docLink || '-'}
Catatan Proyek : ${uas.notes || '-'}`
      );

      // Extract UAS uploaded file if present
      if (uas.fileData) {
        if (uas.fileData.startsWith('data:')) {
          const m = uas.fileData.match(/^data:([A-Za-z0-9+/=;.-]+);base64,(.+)$/);
          if (m && m[2]) {
            const ext = uas.fileName ? path.extname(uas.fileName) : '.pdf';
            uasFolder?.file(`${std?.nim || 'uas'}_${(uas.studentName || 'mhs').replace(/[^a-zA-Z0-9]/g, '_')}_lampiran${ext}`, Buffer.from(m[2], 'base64'));
          }
        } else if (uas.fileData.startsWith('/uploads/')) {
          const localPath = path.join(UPLOAD_DIR, path.basename(uas.fileData));
          if (fs.existsSync(localPath)) {
            const ext = path.extname(localPath) || '.pdf';
            uasFolder?.file(`${std?.nim || 'uas'}_${(uas.studentName || 'mhs').replace(/[^a-zA-Z0-9]/g, '_')}_lampiran${ext}`, fs.readFileSync(localPath));
          }
        }
      }
    }
    uasFolder?.file("00_REKAP_PENGUMPULAN_UAS.txt", uasIndex);

    // 7. Folder: kuis_interaktif
    const kuisFolder = zip.folder("kuis_interaktif");
    let kuisIndex = "REKAP NILAI KUIS INTERAKTIF MAHASISWA:\n\n";
    for (const qSub of (db.quizSubmissions || [])) {
      const std = db.students?.find(s => s.id === qSub.studentId);
      const studentNim = (qSub as any).nim || std?.nim || '-';
      const pct = (qSub as any).percentage ?? Math.round(((qSub.correctCount || 0) / (qSub.totalQuestions || 1)) * 100);
      kuisIndex += `• ${qSub.studentName || std?.name} (NIM: ${studentNim}) - Skor: ${qSub.score ?? '-'} (Akurasi: ${pct}%)\n`;
      kuisIndex += `  Waktu Selesai : ${qSub.submittedAt || '-'}\n`;
      kuisIndex += `  Benar / Total : ${qSub.correctCount ?? '-'} dari ${qSub.totalQuestions ?? '-'}\n\n`;

      kuisFolder?.file(
        `${studentNim}_${(qSub.studentName || std?.name || 'mhs').replace(/[^a-zA-Z0-9]/g, '_')}_kuis.txt`,
        `HASIL KUIS INTERAKTIF MAHASISWA
Nama Mahasiswa : ${qSub.studentName || std?.name}
NIM            : ${studentNim}
Skor Akhir     : ${qSub.score ?? '-'}
Akurasi        : ${pct}%
Jawaban Benar  : ${qSub.correctCount ?? '-'} dari ${qSub.totalQuestions ?? '-'} soal
Waktu Selesai  : ${qSub.submittedAt || '-'}
`
      );
    }
    kuisFolder?.file("00_REKAP_NILAI_KUIS.txt", kuisIndex);

    // 8. Folder: arsip_kiriman_sepanjang_masa (Vault Engine Brankas Permanen)
    const vaultFolder = zip.folder("arsip_kiriman_sepanjang_masa");
    if (fs.existsSync(SUBMISSIONS_VAULT_FILE)) {
      try {
        const vaultRaw = fs.readFileSync(SUBMISSIONS_VAULT_FILE, 'utf-8');
        vaultFolder?.file("brankas_permanen_siakad.json", vaultRaw);
      } catch (vErr) {
        console.warn('Vault zip copy non-fatal:', vErr);
      }
    }

    // 9. Berkas Upload Asli jika ada di data/uploads
    const uploadsFolder = zip.folder("berkas_upload_asli");
    if (fs.existsSync(UPLOAD_DIR)) {
      try {
        const files = fs.readdirSync(UPLOAD_DIR);
        for (const file of files) {
          const filePath = path.join(UPLOAD_DIR, file);
          if (fs.statSync(filePath).isFile()) {
            const fileBuf = fs.readFileSync(filePath);
            uploadsFolder?.file(file, fileBuf);
          }
        }
      } catch (fErr) {
        console.warn('Uploads folder zip non-fatal:', fErr);
      }
    }

    // Generate ZIP buffer
    const zipBuffer = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });

    const safeFilename = `backup-siakad-${semester.replace(/[^a-zA-Z0-9]/g, '_')}-${dateStr}.zip`;
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${safeFilename}"`);
    res.send(zipBuffer);
  } catch (err) {
    console.error('Failed to create ZIP backup:', err);
    res.status(500).json({ error: 'Gagal membuat file backup ZIP semester.' });
  }
});

// 7h-1. Complete Database & Task Backup Download (JSON)
app.get('/api/backup/download', (req, res) => {
  const { dosenPassword, ...safeDb } = db;
  const backupPayload = {
    app: 'SIAKAD MPI 1 - Pascasarjana STAI Jarinabi',
    courseTitle: db.courseProfile?.courseTitle || 'Filsafat Ilmu',
    dosenName: db.courseProfile?.dosenName || 'Risfa Tri Ulfa, S.Pd., M.Pd., Gr.',
    exportTimestamp: new Date().toISOString(),
    totalStudents: db.students?.length || 0,
    totalPresentationSubmissions: db.submissions?.length || 0,
    totalUtsSubmissions: db.utsSubmissions?.length || 0,
    totalUasSubmissions: db.uasSubmissions?.length || 0,
    totalQuizSubmissions: db.quizSubmissions?.length || 0,
    data: safeDb,
  };

  res.setHeader('Content-Type', 'application/json');
  res.setHeader('Content-Disposition', `attachment; filename=siakad-cadangan-tugas-${new Date().toISOString().slice(0, 10)}.json`);
  res.send(JSON.stringify(backupPayload, null, 2));
});

// 7h-1b. Restore Database & Tasks from JSON Backup (Dosen Only)
app.post('/api/backup/restore', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang memiliki hak untuk memulihkan cadangan data.' });
  }

  try {
    const rawData = req.body.backupData || req.body.data || req.body;
    if (!rawData || !rawData.students || !Array.isArray(rawData.students)) {
      return res.status(400).json({ error: 'File cadangan tidak valid. Struktur data SIAKAD tidak ditemukan.' });
    }

    // Merge/replace db state safely
    db = {
      ...INITIAL_DATABASE,
      ...rawData,
      dosenPassword: db.dosenPassword || 'filsafat2026',
    };

    saveDatabase();
    res.json({
      success: true,
      message: 'Data tugas, ujian, kuis, dan nilai berhasil dipulihkan secara permanen ke server!',
      totalStudents: db.students.length,
      totalSubmissions: db.submissions?.length || 0,
      totalUts: db.utsSubmissions?.length || 0,
      totalUas: db.uasSubmissions?.length || 0,
    });
  } catch (err) {
    console.error('Failed to restore backup:', err);
    res.status(500).json({ error: 'Gagal memproses file cadangan.' });
  }
});

// 7h-2. Get Exam Availability Settings (UTS & UAS)
app.get('/api/exam-settings', (req, res) => {
  res.json({
    success: true,
    uts: db.utsExamSettings || {
      isOpen: true,
      instructions: 'Ujian Tengah Semester (UTS) dibuka oleh Dosen Pengampu.',
    },
    uas: db.uasExamSettings || {
      isOpen: true,
      instructions: 'Ujian Akhir Semester (UAS) dibuka oleh Dosen Pengampu.',
    },
  });
});

// 7h-3. Update Exam Settings / Toggle Open & Close (Dosen Only)
app.post('/api/exam-settings', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak mengatur jadwal dan membuka/menutup ujian UTS dan UAS.' });
  }

  const { examType, isOpen, openDate, closeDate, instructions } = req.body;
  if (examType !== 'uts' && examType !== 'uas') {
    return res.status(400).json({ error: 'Tipe ujian harus "uts" atau "uas"' });
  }

  const currentObj = examType === 'uts'
    ? (db.utsExamSettings = db.utsExamSettings || { isOpen: false })
    : (db.uasExamSettings = db.uasExamSettings || { isOpen: false });

  if (isOpen !== undefined) currentObj.isOpen = Boolean(isOpen);
  if (openDate !== undefined) currentObj.openDate = openDate;
  if (closeDate !== undefined) currentObj.closeDate = closeDate;
  if (instructions !== undefined) currentObj.instructions = instructions;

  if (currentObj.isOpen) {
    currentObj.openedAt = new Date().toISOString();
    currentObj.openedBy = db.courseProfile?.name || 'Risfa Tri Ulfa, S.Pd., M.Pd., Gr.';
  } else {
    currentObj.closedAt = new Date().toISOString();
  }

  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    if (examType === 'uts') db.allCoursesData[db.activeCourseId].utsExamSettings = db.utsExamSettings;
    else db.allCoursesData[db.activeCourseId].uasExamSettings = db.uasExamSettings;
  }

  saveDatabase();
  res.json({
    success: true,
    examType,
    settings: currentObj,
    message: `Jadwal & Status ${examType.toUpperCase()} berhasil diperbarui: ${currentObj.isOpen ? 'DIBUKA UNTUK MAHASISWA' : 'DITUTUP / TERKUNCI'}.`,
  });
});

// 7i. Get Quiz Data (10 Questions, Submissions, and Settings)
app.get('/api/quiz', (req, res) => {
  const questions = db.quizQuestions && db.quizQuestions.length > 0 ? db.quizQuestions : INITIAL_QUIZ_QUESTIONS;
  const submissions = db.quizSubmissions || [];
  const settings = db.quizSettings || INITIAL_DATABASE.quizSettings;
  res.json({
    success: true,
    questions,
    submissions,
    settings,
    courseTitle: db.courseProfile?.courseTitle || 'Filsafat Ilmu',
  });
});

// 7i-2. Get Quiz Settings
app.get('/api/quiz/settings', (req, res) => {
  res.json({
    success: true,
    settings: db.quizSettings || INITIAL_DATABASE.quizSettings,
  });
});

// 7i-3. Update Quiz Settings (Dosen Only)
app.post('/api/quiz/settings', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak mengatur jadwal dan aktivasi kuis!' });
  }
  const {
    isQuizActive,
    targetMeeting,
    quizTitle,
    timeLimitMinutes,
    description,
    gameMode,
    soundEffectsEnabled,
    musicEnabled,
    defaultMusicTrack,
    allowStudentModeSelection,
  } = req.body;

  if (!db.quizSettings) {
    db.quizSettings = {
      isQuizActive: true,
      targetMeeting: 'Semua Pertemuan',
      quizTitle: 'Game Cerdas Cermat Kartun RPS',
      timeLimitMinutes: 15,
      gameMode: 'bebas_pilih',
      soundEffectsEnabled: true,
      musicEnabled: true,
      defaultMusicTrack: 'lofi-calm',
      allowStudentModeSelection: true,
    };
  }
  if (typeof isQuizActive === 'boolean') db.quizSettings.isQuizActive = isQuizActive;
  if (targetMeeting) db.quizSettings.targetMeeting = targetMeeting;
  if (quizTitle) db.quizSettings.quizTitle = quizTitle;
  if (timeLimitMinutes) db.quizSettings.timeLimitMinutes = Number(timeLimitMinutes);
  if (description !== undefined) db.quizSettings.description = description;
  if (gameMode) db.quizSettings.gameMode = gameMode;
  if (typeof soundEffectsEnabled === 'boolean') db.quizSettings.soundEffectsEnabled = soundEffectsEnabled;
  if (typeof musicEnabled === 'boolean') db.quizSettings.musicEnabled = musicEnabled;
  if (defaultMusicTrack) db.quizSettings.defaultMusicTrack = defaultMusicTrack;
  if (typeof allowStudentModeSelection === 'boolean') db.quizSettings.allowStudentModeSelection = allowStudentModeSelection;

  saveDatabase();
  res.json({ success: true, settings: db.quizSettings });
});

// 7i-4. Upload Material & Auto-generate 10 Quiz Questions (Dosen Only)
app.post('/api/quiz/upload-material', async (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak mengunggah materi kuis!' });
  }

  const { materialText, quizTitle, targetMeeting, fileData, fileName } = req.body;
  let rawText = materialText || '';

  // Extract from fileData if base64 provided
  if (fileData && !rawText) {
    try {
      const base64Data = fileData.split(',')[1] || fileData;
      const buffer = Buffer.from(base64Data, 'base64');
      const lowerName = (fileName || '').toLowerCase();

      if (lowerName.endsWith('.docx')) {
        const result = await mammoth.extractRawText({ buffer });
        rawText = result.value;
      } else if (lowerName.endsWith('.pdf')) {
        const parsePdf = (pdfParseModule as any).default || pdfParseModule;
        const pdfData = await parsePdf(buffer);
        rawText = pdfData.text;
      } else {
        rawText = buffer.toString('utf-8');
      }
    } catch (parseErr) {
      console.error('Error parsing uploaded quiz file:', parseErr);
      rawText = materialText || '';
    }
  }

  if (!rawText || rawText.trim().length < 10) {
    return res.status(400).json({ error: 'Materi kuis wajib memiliki konten teks untuk diolah menjadi 10 butir soal.' });
  }

  const targetTitle = quizTitle || (targetMeeting ? `Kuis ${targetMeeting}` : `Kuis Materi ${db.courseProfile?.courseTitle || 'RPS'}`);
  const questions = generateQuizFromMaterial(
    rawText,
    db.courseProfile?.courseTitle || 'Filsafat Ilmu',
    targetMeeting || 'Materi Unggahan Dosen'
  );

  db.quizQuestions = questions;
  if (!db.quizSettings) {
    db.quizSettings = {
      isQuizActive: true,
      targetMeeting: targetMeeting || 'Materi Unggahan Dosen',
      quizTitle: targetTitle,
      timeLimitMinutes: 15,
    };
  } else {
    db.quizSettings.isQuizActive = true;
    if (targetMeeting) db.quizSettings.targetMeeting = targetMeeting;
    db.quizSettings.quizTitle = targetTitle;
  }

  saveDatabase();
  res.json({
    success: true,
    questions: db.quizQuestions,
    settings: db.quizSettings,
    message: `Berhasil mengekstraksi dan menyusun 10 butir soal kuis interaktif dari materi yang diunggah dosen.`,
  });
});

// 7j. Update Quiz Questions (Dosen Only)
app.post('/api/quiz/questions', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak menyunting dan mengatur soal kuis!' });
  }
  const { questions } = req.body;
  if (!Array.isArray(questions) || questions.length === 0) {
    return res.status(400).json({ error: 'Daftar soal wajib disertakan (minimal 1 soal).' });
  }

  db.quizQuestions = questions;
  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].quizQuestions = questions;
  }
  saveDatabase();
  res.json({ success: true, questions: db.quizQuestions });
});

// 7k. Auto-generate 10 Quiz Questions from RPS Meetings (Dosen Only)
app.post('/api/quiz/auto-generate', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak men-generate soal kuis RPS!' });
  }

  const generated = generateQuizQuestionsFromMeetings(
    db.meetings || [],
    db.courseProfile?.courseTitle || 'Mata Kuliah'
  );
  db.quizQuestions = generated;
  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].quizQuestions = generated;
  }
  saveDatabase();
  res.json({
    success: true,
    questions: db.quizQuestions,
    message: `Berhasil men-generate ${generated.length} soal kuis interaktif secara otomatis dari materi 16 pertemuan RPS.`,
  });
});

// 7l. Submit Quiz Answers (Mahasiswa dengan Pengawasan Kamera & Poin Otomatis)
app.post('/api/quiz/submit', (req, res) => {
  const { studentId, studentName, answers, cameraVerified, timeTakenSeconds } = req.body;
  if (!studentId) {
    return res.status(400).json({ error: 'studentId wajib disertakan.' });
  }

  const questions = db.quizQuestions && db.quizQuestions.length > 0 ? db.quizQuestions : INITIAL_QUIZ_QUESTIONS;
  let correctCount = 0;
  const explanationMap: Record<number, { correctIndex: number; isCorrect: boolean; explanation: string }> = {};

  questions.forEach(q => {
    const studentAnswer = answers?.[q.id];
    const isCorrect = studentAnswer === q.correctIndex;
    if (isCorrect) correctCount++;
    explanationMap[q.id] = {
      correctIndex: q.correctIndex,
      isCorrect,
      explanation: q.explanation,
    };
  });

  const score = Math.round((correctCount / questions.length) * 100);

  if (!db.quizSubmissions) db.quizSubmissions = [];
  const existingIdx = db.quizSubmissions.findIndex(s => s.studentId === studentId);
  const subRecord: QuizSubmission = {
    id: existingIdx >= 0 ? db.quizSubmissions[existingIdx].id : `quiz-${Date.now()}`,
    studentId,
    studentName: studentName || 'Mahasiswa',
    score,
    correctCount,
    totalQuestions: questions.length,
    submittedAt: new Date().toISOString(),
    cameraVerified: Boolean(cameraVerified),
    answers: answers || {},
    timeTakenSeconds: timeTakenSeconds || 0,
  };

  if (existingIdx >= 0) {
    db.quizSubmissions[existingIdx] = subRecord;
  } else {
    db.quizSubmissions.push(subRecord);
  }

  // Automatically integrate with student's attitude/quiz grade in SIAKAD
  if (!db.grades) db.grades = {};
  if (!db.grades[studentId]) {
    db.grades[studentId] = {
      attendanceScore: 100,
      attitudeScore: score,
      individualScore: 85,
      utsScore: 85,
      uasScore: 85,
      groupScore: 85,
      finalScore: 85,
      letterGrade: 'A',
      notes: `Kuis RPS: ${score}/100`,
    };
  } else {
    // Boost or sync attitude/keaktifan score
    const currentAttitude = db.grades[studentId].attitudeScore || 80;
    db.grades[studentId].attitudeScore = Math.min(100, Math.max(currentAttitude, score));
    db.grades[studentId].notes = `Poin Kuis Interaktif RPS: ${score}/100 (Pengawasan Kamera: ${cameraVerified ? 'Terverifikasi Aktif' : 'Non-Kamera'})`;
  }

  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].quizSubmissions = db.quizSubmissions;
    db.allCoursesData[db.activeCourseId].grades = db.grades;
  }

  saveDatabase();
  res.json({
    success: true,
    submission: subRecord,
    score,
    correctCount,
    totalQuestions: questions.length,
    explanationMap,
  });
});

// 7l-2. Delete / Reset Quiz Submission (HANYA DOSEN - Memungkinkan dosen menghapus nilai kuis jika salah)
app.delete('/api/quiz/submissions/:id', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak menghapus nilai kuis mahasiswa!' });
  }

  const subId = req.params.id;
  const studentId = req.query.studentId as string;

  if (!db.quizSubmissions) db.quizSubmissions = [];

  const initialLength = db.quizSubmissions.length;
  const targetSub = db.quizSubmissions.find(s => s.id === subId || (studentId && s.studentId === studentId));

  db.quizSubmissions = db.quizSubmissions.filter(s => s.id !== subId && (!studentId || s.studentId !== studentId));

  if (db.quizSubmissions.length === initialLength && !targetSub) {
    return res.status(404).json({ error: 'Data pengerjaan kuis tidak ditemukan atau sudah dihapus.' });
  }

  const targetStudentId = targetSub ? targetSub.studentId : studentId;

  // Clear notes or attitude override in student grade record if exists
  if (targetStudentId && db.grades && db.grades[targetStudentId]) {
    if (db.grades[targetStudentId].notes && db.grades[targetStudentId].notes.includes('Kuis')) {
      db.grades[targetStudentId].notes = '';
    }
  }

  // Update in active course data
  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].quizSubmissions = db.quizSubmissions;
    db.allCoursesData[db.activeCourseId].grades = db.grades;
  }

  saveDatabase();
  res.json({
    success: true,
    message: `Nilai kuis mahasiswa berhasil dihapus oleh Dosen. Akses pengerjaan kuis telah dibuka kembali untuk mahasiswa yang bersangkutan.`,
    remainingCount: db.quizSubmissions.length,
  });
});

// 7m. Reorganize Groups & Adjust Count Across Semesters (Dosen Only)
app.post('/api/groups/reorganize', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berhak merombak dan mengatur ulang kelompok!' });
  }

  const { groupCount = 5, mode = 'even' } = req.body;
  const count = Math.max(2, Math.min(15, Number(groupCount) || 5));

  const newGroups: GroupProject[] = [];
  for (let i = 1; i <= count; i++) {
    newGroups.push({
      id: i,
      name: `KELOMPOK ${i}`,
      members: [],
      title: `Kajian Proyek Kolaboratif Kelompok ${i}`,
      description: `Proyek kolaboratif dan telaah bahan kajian RPS semester untuk Kelompok ${i}.`,
      toolsSuggested: 'Canva / AI / Google Slides',
    });
  }

  const students = [...(db.students || [])];
  if (mode === 'random') {
    for (let i = students.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [students[i], students[j]] = [students[j], students[i]];
    }
  }

  // Distribute students evenly across newly configured groups
  students.forEach((std, idx) => {
    const grpIdx = idx % count;
    const assignedGroup = newGroups[grpIdx];
    std.groupId = assignedGroup.id;
    std.rpsPart = `Pertemuan ${Math.min(15, grpIdx + 2)}`;
    if (!assignedGroup.members.includes(std.name)) {
      assignedGroup.members.push(std.name);
    }
  });

  db.groups = newGroups;
  db.students = students;

  // Re-synchronize presentation meetings with newly configured groups
  if (Array.isArray(db.meetings)) {
    let gIdx = 0;
    db.meetings.forEach(m => {
      if (m.presentationFormat === 'kelompok' && m.meetingNumber > 1 && m.type === 'kuliah') {
        const grp = newGroups[gIdx % newGroups.length];
        m.groupId = grp.id;
        m.groupName = grp.name;
        m.presenters = [...grp.members];
        gIdx++;
      }
    });
  }

  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].groups = db.groups;
    db.allCoursesData[db.activeCourseId].students = db.students;
    db.allCoursesData[db.activeCourseId].meetings = db.meetings;
  }

  saveDatabase();
  res.json({
    success: true,
    groups: db.groups,
    students: db.students,
    message: `Berhasil membagi ulang mahasiswa ke dalam ${count} kelompok baru (${mode === 'random' ? 'Acak/Random' : 'Merata/Even'}) dan menyinkronkan ke jadwal presentasi perkuliahan.`,
  });
});

// 8. Update Attendance (Dosen)
app.post('/api/attendance', (req, res) => {
  const { meetingNumber, studentId, status, bulkStatus } = req.body;

  if (meetingNumber && bulkStatus) {
    // Bulk set for entire meeting
    if (!db.attendance[meetingNumber]) {
      db.attendance[meetingNumber] = {};
    }
    db.students.forEach(s => {
      db.attendance[meetingNumber][s.id] = bulkStatus;
    });
  } else if (meetingNumber && studentId && status) {
    if (!db.attendance[meetingNumber]) {
      db.attendance[meetingNumber] = {};
    }
    db.attendance[meetingNumber][studentId] = status;
  }

  // Recalculate attendance score for affected students
  db.students.forEach(s => {
    let hadirCount = 0;
    let totalRecorded = 0;
    for (let m = 1; m <= 16; m++) {
      const rec = db.attendance[m]?.[s.id];
      if (rec) {
        totalRecorded++;
        if (rec === 'H') hadirCount += 1;
        else if (rec === 'I') hadirCount += 0.8;
        else if (rec === 'S') hadirCount += 0.8;
      }
    }
    const attPercent = totalRecorded > 0 ? Math.round((hadirCount / totalRecorded) * 100) : 100;
    if (!db.grades[s.id]) {
      db.grades[s.id] = {
        attendanceScore: attPercent,
        attitudeScore: 85,
        individualScore: 85,
        utsScore: 85,
        uasScore: 85,
        groupScore: 85,
        finalScore: 88,
        letterGrade: 'A-',
      };
    } else {
      db.grades[s.id].attendanceScore = attPercent;
    }
    recalculateStudentGrade(db.grades[s.id]);
  });

  saveDatabase();
  res.json({ success: true, attendance: db.attendance });
});

// 9. Update full grade record (Dosen only)
app.post('/api/grades', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang memiliki hak akses untuk mengubah nilai akhir mahasiswa!' });
  }

  const { studentId, attendanceScore, attitudeScore, individualScore, utsScore, uasScore, groupScore, notes } = req.body;
  if (!studentId) {
    return res.status(400).json({ error: 'Missing studentId' });
  }

  const att = Number(attendanceScore ?? 100);
  const attit = Number(attitudeScore ?? 85);
  const indiv = Number(individualScore ?? 85);
  const uts = Number(utsScore ?? 85);
  const grp = Number(groupScore ?? 85);
  const uas = Number(uasScore ?? grp ?? 85);

  const gradeObj: StudentGrade = {
    attendanceScore: att,
    attitudeScore: attit,
    individualScore: indiv,
    utsScore: uts,
    uasScore: uas,
    groupScore: grp,
    finalScore: 88,
    letterGrade: 'A-',
    notes,
  };
  recalculateStudentGrade(gradeObj);
  db.grades[studentId] = gradeObj;

  saveDatabase();
  res.json({ success: true, grade: db.grades[studentId] });
});

// 10. Add student
app.post('/api/students', (req, res) => {
  const { name, nim, birthPlace, birthDate, address, gender, phone, rpsPart, topic, meetingNumber, groupId } = req.body;
  if (!name) {
    return res.status(400).json({ error: 'Name is required' });
  }

  const assignedGroupId = Number(groupId) || 1;
  const formattedName = name.trim().toUpperCase();

  const newStudent: Student = {
    id: `mhs-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
    nim: nim ? nim.trim() : `2026${String(db.students.length + 1).padStart(4, '0')}`,
    name: formattedName,
    birthPlace: birthPlace ? birthPlace.trim() : 'Pasuruan',
    birthDate: birthDate ? birthDate.trim() : '1998-05-15',
    address: address ? address.trim() : 'Kabupaten Pasuruan, Jawa Timur',
    gender: gender || (
      formattedName.includes('SARI') || formattedName.includes('LESTARI') || formattedName.includes('MAGHFIROH') ||
      formattedName.includes('ALIYAH') || formattedName.includes('DEFTIRIYANI')
        ? 'Perempuan'
        : 'Laki-laki'
    ),
    phone: phone ? phone.trim() : `08123456${String(db.students.length + 1).padStart(4, '0')}`,
    rpsPart: rpsPart || `Pertemuan ${Number(meetingNumber) || 15}`,
    topic: topic || 'Materi Tambahan / Telaah Khusus Filsafat Ilmu',
    meetingNumber: Number(meetingNumber) || 15,
    groupId: assignedGroupId,
    createdAt: new Date().toISOString(),
  };

  db.students.push(newStudent);

  // Synchronize with group members
  const targetGroup = db.groups.find(g => g.id === assignedGroupId);
  if (targetGroup && !targetGroup.members.includes(formattedName)) {
    targetGroup.members.push(formattedName);
  }

  // Initialize grade
  db.grades[newStudent.id] = {
    attendanceScore: 100,
    attitudeScore: 85,
    individualScore: 85,
    groupScore: targetGroup?.grade ?? 85,
    utsScore: 85,
    uasScore: 85,
    finalScore: 88,
    letterGrade: 'A-',
  };

  saveDatabase();
  res.json({ success: true, student: newStudent });
});

// 10a-2. Update Student Biodata & Assignment
app.put('/api/students/:id', (req, res) => {

  const { id } = req.params;
  const student = (db.students || []).find(s => s.id === id);
  if (!student) {
    return res.status(404).json({ error: 'Mahasiswa tidak ditemukan' });
  }

  const { name, nim, birthPlace, birthDate, address, gender, phone, rpsPart, topic, meetingNumber, groupId } = req.body;
  const oldName = student.name;

  if (name) student.name = name.trim().toUpperCase();
  if (nim) student.nim = nim.trim();
  if (birthPlace !== undefined) student.birthPlace = birthPlace.trim();
  if (birthDate !== undefined) student.birthDate = birthDate.trim();
  if (address !== undefined) student.address = address.trim();
  if (gender !== undefined) student.gender = gender;
  if (phone !== undefined) student.phone = phone.trim();
  if (rpsPart !== undefined) student.rpsPart = rpsPart;
  if (topic !== undefined) student.topic = topic;
  if (meetingNumber !== undefined) student.meetingNumber = Number(meetingNumber);

  // If group changed or name changed, update group members
  if (groupId !== undefined && Number(groupId) !== student.groupId) {
    // Remove from old group
    const oldGroup = (db.groups || []).find(g => g.id === student.groupId);
    if (oldGroup) {
      oldGroup.members = oldGroup.members.filter(m => m !== oldName && m !== student.name);
    }
    // Add to new group
    student.groupId = Number(groupId);
    const newGroup = (db.groups || []).find(g => g.id === student.groupId);
    if (newGroup && !newGroup.members.includes(student.name)) {
      newGroup.members.push(student.name);
    }
  } else if (name && oldName !== student.name) {
    // Update name in same group
    const curGroup = (db.groups || []).find(g => g.id === student.groupId);
    if (curGroup) {
      curGroup.members = curGroup.members.map(m => m === oldName ? student.name : m);
    }
  }

  // Update name in submissions & UTS & UAS
  if (name && oldName !== student.name) {
    (db.submissions || []).forEach(sub => {
      if (sub.studentId === id) sub.studentName = student.name;
    });
    (db.utsSubmissions || []).forEach(uts => {
      if (uts.studentId === id) uts.studentName = student.name;
    });
    (db.uasSubmissions || []).forEach(uas => {
      if (uas.studentId === id) uas.studentName = student.name;
    });
  }

  saveDatabase();
  res.json({ success: true, student });
});

// 10a-2b. Add or Adjust Student Activity Points (Poin Keaktifan) (Dosen Only)
app.post('/api/students/:id/activity-points', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak: Hanya Dosen yang berwenang menambah atau mengurangi Poin Keaktifan mahasiswa!' });
  }

  const { id } = req.params;
  const student = (db.students || []).find(s => s.id === id || s.nim === id);
  if (!student) {
    return res.status(404).json({ error: 'Mahasiswa tidak ditemukan di database' });
  }

  const { points, reason, category, meetingNumber } = req.body;
  const pointDelta = Number(points);
  if (isNaN(pointDelta)) {
    return res.status(400).json({ error: 'Nilai poin tidak valid' });
  }

  const logReason = typeof reason === 'string' && reason.trim()
    ? reason.trim()
    : (pointDelta >= 0 ? 'Poin Keaktifan Partisipasi Kelas' : 'Pengurangan Poin Keaktifan');

  const dosenTitle = db.courseProfile?.name || 'Risfa Tri Ulfa, S.Pd., M.Pd.';

  const newLog: ActivityPointLog = {
    id: `act-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
    points: pointDelta,
    reason: logReason,
    category: category || (pointDelta >= 0 ? 'manual' : 'sikap'),
    meetingNumber: meetingNumber ? Number(meetingNumber) : undefined,
    timestamp: new Date().toISOString(),
    dosenName: dosenTitle,
  };

  if (!student.activityPointLogs) student.activityPointLogs = [];
  student.activityPointLogs.unshift(newLog);

  // Recalculate or add total points (minimum 0)
  const currentTotal = student.activityPoints !== undefined ? student.activityPoints : 0;
  student.activityPoints = Math.max(0, currentTotal + pointDelta);

  // Sync to database level map
  if (!db.activityPointLogs) db.activityPointLogs = {};
  db.activityPointLogs[student.id] = student.activityPointLogs;

  // Sync to active course in allCoursesData if available
  if (db.activeCourseId && db.allCoursesData && db.allCoursesData[db.activeCourseId]) {
    const coursePayload = db.allCoursesData[db.activeCourseId];
    if (coursePayload.students) {
      const matchInCourse = coursePayload.students.find(s => s.id === student.id || s.nim === student.nim);
      if (matchInCourse) {
        matchInCourse.activityPoints = student.activityPoints;
        matchInCourse.activityPointLogs = student.activityPointLogs;
      }
    }
  }

  saveDatabase();
  res.json({
    success: true,
    student,
    log: newLog,
    totalPoints: student.activityPoints,
    message: `Poin keaktifan mahasiswa ${student.name} berhasil diperbarui (${pointDelta > 0 ? '+' : ''}${pointDelta} poin).`,
  });
});

// 10a-2c. Delete an Activity Point Log Entry (Dosen Only)
app.delete('/api/students/:id/activity-points/:logId', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak: Hanya Dosen yang berwenang mengelola log poin keaktifan!' });
  }

  const { id, logId } = req.params;
  const student = (db.students || []).find(s => s.id === id || s.nim === id);
  if (!student) {
    return res.status(404).json({ error: 'Mahasiswa tidak ditemukan' });
  }

  if (!student.activityPointLogs) student.activityPointLogs = [];
  const logIndex = student.activityPointLogs.findIndex(l => l.id === logId);
  if (logIndex === -1) {
    return res.status(404).json({ error: 'Catatan poin keaktifan tidak ditemukan' });
  }

  const removedLog = student.activityPointLogs[logIndex];
  student.activityPointLogs.splice(logIndex, 1);

  // Adjust total points
  student.activityPoints = Math.max(0, (student.activityPoints || 0) - removedLog.points);

  if (!db.activityPointLogs) db.activityPointLogs = {};
  db.activityPointLogs[student.id] = student.activityPointLogs;

  saveDatabase();
  res.json({
    success: true,
    student,
    totalPoints: student.activityPoints,
    message: 'Catatan poin keaktifan berhasil dihapus.',
  });
});

// 10a-2d. Bulk Add Activity Points (e.g. for an entire presentation group) (Dosen Only)
app.post('/api/students/bulk-activity-points', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak: Hanya Dosen yang berwenang memberikan poin masal!' });
  }

  const { studentIds, points, reason, category, meetingNumber } = req.body;
  if (!Array.isArray(studentIds) || studentIds.length === 0) {
    return res.status(400).json({ error: 'Daftar studentIds wajib disertakan' });
  }

  const pointDelta = Number(points);
  if (isNaN(pointDelta)) {
    return res.status(400).json({ error: 'Nilai poin tidak valid' });
  }

  const dosenTitle = db.courseProfile?.name || 'Risfa Tri Ulfa, S.Pd., M.Pd.';
  let updatedCount = 0;

  studentIds.forEach(sId => {
    const student = (db.students || []).find(s => s.id === sId || s.nim === sId);
    if (student) {
      if (!student.activityPointLogs) student.activityPointLogs = [];
      const log: ActivityPointLog = {
        id: `act-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
        points: pointDelta,
        reason: reason || 'Poin Keaktifan Kelompok Presentasi',
        category: category || 'presentasi',
        meetingNumber: meetingNumber ? Number(meetingNumber) : undefined,
        timestamp: new Date().toISOString(),
        dosenName: dosenTitle,
      };
      student.activityPointLogs.unshift(log);
      student.activityPoints = Math.max(0, (student.activityPoints || 0) + pointDelta);

      if (!db.activityPointLogs) db.activityPointLogs = {};
      db.activityPointLogs[student.id] = student.activityPointLogs;
      updatedCount++;
    }
  });

  saveDatabase();
  res.json({
    success: true,
    updatedCount,
    message: `Berhasil menambahkan poin keaktifan ke ${updatedCount} mahasiswa.`,
  });
});

// 10a-3. Delete Student (Dosen Only)
app.delete('/api/students/:id', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang menghapus mahasiswa.' });
  }

  const { id } = req.params;
  const student = (db.students || []).find(s => s.id === id);
  if (!student) {
    return res.status(404).json({ error: 'Mahasiswa tidak ditemukan' });
  }

  // Remove from group
  const group = (db.groups || []).find(g => g.id === student.groupId);
  if (group) {
    group.members = group.members.filter(m => m !== student.name);
  }

  // Remove from students array
  db.students = (db.students || []).filter(s => s.id !== id);

  // Clean attendance and grades
  delete db.grades[id];
  if (db.attendance) {
    Object.keys(db.attendance).forEach(m => {
      if (db.attendance[Number(m)]) {
        delete db.attendance[Number(m)][id];
      }
    });
  }

  saveDatabase();
  res.json({ success: true, message: `Mahasiswa ${student.name} (${student.nim}) berhasil dihapus.` });
});

// 10a-3b. Restore Canonical 15 Students & Groups (Dosen / System Restore)
app.post('/api/students/restore-canonical', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang memulihkan data mahasiswa.' });
  }

  db.students = JSON.parse(JSON.stringify(INITIAL_STUDENTS));
  db.groups = JSON.parse(JSON.stringify(INITIAL_GROUPS));
  
  if (!db.grades) db.grades = {};
  db.students.forEach(std => {
    if (!db.grades[std.id]) {
      db.grades[std.id] = {
        attendanceScore: 100,
        attitudeScore: 85,
        individualScore: 85,
        utsScore: 85,
        uasScore: 85,
        groupScore: 85,
        finalScore: 88,
        letterGrade: 'A-',
      };
    }
  });

  saveDatabase();
  res.json({
    success: true,
    message: '13 Mahasiswa resmi MPI STAI Jarinabi dan 10 kelompok RPS berhasil dipulihkan secara penuh!',
    students: db.students,
    groups: db.groups,
  });
});

// 10a-3c. Sync 10 Official Presentation & Video Groups from RPS STAI Jarinabi
app.all('/api/groups/sync-rps-10', (req, res) => {
  db.groups = JSON.parse(JSON.stringify(INITIAL_GROUPS));
  db.meetings = db.meetings.map(m => {
    const initM = INITIAL_MEETINGS.find(im => im.meetingNumber === m.meetingNumber);
    return initM ? { ...m, ...initM } : m;
  });
  saveDatabase();
  res.json({
    success: true,
    message: '10 Kelompok resmi perkuliahan & presentasi sesuai RPS STAI Jarinabi berhasil disinkronkan otomatis!',
    groups: db.groups,
    meetings: db.meetings,
  });
});

// 10a-4. Bulk Import Students from Excel / CSV / JSON (Dosen Only)
app.post('/api/students/bulk-import', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang mengimpor data mahasiswa.' });
  }

  const { students: rawStudents, mode } = req.body;
  if (!Array.isArray(rawStudents) || rawStudents.length === 0) {
    return res.status(400).json({ error: 'Daftar data mahasiswa wajib disertakan (minimal 1 baris).' });
  }

  if (mode === 'replace') {
    db.students = [];
    (db.groups || []).forEach(g => { g.members = []; });
    db.grades = {};
  }

  const importedList: Student[] = [];
  const groupCount = db.groups && db.groups.length > 0 ? db.groups.length : 5;

  rawStudents.forEach((row, idx) => {
    const rawName = String(row.name || row.nama || '').trim();
    if (!rawName) return;

    const formattedName = rawName.toUpperCase();
    const cleanNim = String(row.nim || row.NIM || `2026${String(db.students.length + 1).padStart(4, '0')}`).trim();
    const cleanBirthPlace = String(row.birthPlace || row.tempatLahir || row.tempat_lahir || row.Tempat_Lahir || 'Pasuruan').trim();
    const cleanBirthDate = String(row.birthDate || row.tanggalLahir || row.tanggal_lahir || row.Tanggal_Lahir || '1998-05-15').trim();
    const cleanAddress = String(row.address || row.alamat || row.Alamat || 'Kabupaten Pasuruan, Jawa Timur').trim();
    const cleanGender = String(row.gender || row.jenisKelamin || row.jenis_kelamin || (
      formattedName.includes('SARI') || formattedName.includes('LESTARI') || formattedName.includes('MAGHFIROH') ||
      formattedName.includes('ALIYAH') || formattedName.includes('DEFTIRIYANI')
        ? 'Perempuan'
        : 'Laki-laki'
    )).trim();
    const cleanPhone = String(row.phone || row.telepon || row.noHp || `08123456${cleanNim.slice(-4)}`).trim();

    // Group assignment
    const assignedGroupId = row.groupId ? Number(row.groupId) : ((idx % groupCount) + 1);

    // Existing check
    const existing = db.students.find(s => s.nim === cleanNim || s.name === formattedName);
    if (existing) {
      existing.name = formattedName;
      existing.nim = cleanNim;
      existing.birthPlace = cleanBirthPlace;
      existing.birthDate = cleanBirthDate;
      existing.address = cleanAddress;
      existing.gender = cleanGender;
      existing.phone = cleanPhone;
      existing.groupId = assignedGroupId;
      importedList.push(existing);
    } else {
      const newStd: Student = {
        id: `mhs-${Date.now()}-${Math.random().toString(36).substr(2, 4)}-${idx}`,
        nim: cleanNim,
        name: formattedName,
        birthPlace: cleanBirthPlace,
        birthDate: cleanBirthDate,
        address: cleanAddress,
        gender: cleanGender,
        phone: cleanPhone,
        rpsPart: row.rpsPart || `Pertemuan ${Math.min(15, assignedGroupId + 2)}`,
        topic: row.topic || `Telaah Mandiri & Kajian Kelompok ${assignedGroupId}`,
        meetingNumber: Number(row.meetingNumber) || Math.min(15, assignedGroupId + 2),
        groupId: assignedGroupId,
        createdAt: new Date().toISOString(),
      };
      db.students.push(newStd);
      importedList.push(newStd);

      if (!db.grades[newStd.id]) {
        db.grades[newStd.id] = {
          attendanceScore: 100,
          attitudeScore: 85,
          individualScore: 85,
          utsScore: 85,
          uasScore: 85,
          groupScore: 85,
          finalScore: 88,
          letterGrade: 'A-',
        };
      }
    }

    // Add to group members
    const grp = (db.groups || []).find(g => g.id === assignedGroupId);
    if (grp && !grp.members.includes(formattedName)) {
      grp.members.push(formattedName);
    }
  });

  saveDatabase();
  res.json({
    success: true,
    count: importedList.length,
    students: db.students,
    message: `Berhasil mengimpor ${importedList.length} data mahasiswa lengkap (NIM, Nama, TTL, Alamat).`,
  });
});

// 10a-5. Parse & Extract Students from Uploaded Document (PDF, DOCX, DOC, XLSX, CSV, TXT)
app.post('/api/students/upload-document', async (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang mengunggah dokumen mahasiswa.' });
  }

  const { fileBase64, fileName, textContent } = req.body;
  let rawText = (textContent || '').trim();

  if (fileBase64 && fileName) {
    try {
      const buffer = Buffer.from(fileBase64, 'base64');
      const extracted = await extractTextFromFileBuffer(buffer, fileName);
      if (extracted && extracted.trim()) {
        rawText = extracted.trim();
      }
    } catch (err: any) {
      console.warn('Gagal membaca file mahasiswa:', err);
      return res.status(400).json({ error: `Gagal membaca file: ${err?.message || 'Format tidak didukung'}` });
    }
  }

  if (!rawText) {
    return res.status(400).json({ error: 'Tidak ada teks atau dokumen yang dapat diproses.' });
  }

  const parsedStudents = parseStudentsFromDocumentText(rawText);
  res.json({
    success: true,
    count: parsedStudents.length,
    students: parsedStudents,
    rawSnippet: rawText.slice(0, 1000),
  });
});

// 10b. Add student/member to specific group (Only Dosen)
app.post('/api/groups/:id/members', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang menambahkan anggota kelompok.' });
  }
  const groupId = Number(req.params.id);
  const targetGroup = db.groups.find(g => g.id === groupId);
  if (!targetGroup) {
    return res.status(404).json({ error: 'Kelompok tidak ditemukan' });
  }

  const { studentName, studentId, nim, rpsPart, topic, meetingNumber } = req.body;
  let memberName = (studentName || '').trim().toUpperCase();
  let studentObj: Student | undefined;

  if (studentId) {
    studentObj = db.students.find(s => s.id === studentId);
    if (studentObj) {
      memberName = studentObj.name;
      studentObj.groupId = groupId;
    }
  }

  if (!memberName) {
    return res.status(400).json({ error: 'Nama mahasiswa wajib diisi' });
  }

  // If student doesn't exist in db.students yet, create them as well!
  if (!studentObj) {
    studentObj = db.students.find(s => s.name.toUpperCase() === memberName);
    if (studentObj) {
      studentObj.groupId = groupId;
    } else {
      studentObj = {
        id: `mhs-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
        nim: nim ? nim.trim() : `2026${String(db.students.length + 1).padStart(4, '0')}`,
        name: memberName,
        rpsPart: rpsPart || `Pertemuan ${Number(meetingNumber) || 15}`,
        topic: topic || 'Materi Tambahan / Telaah Khusus Filsafat Ilmu',
        meetingNumber: Number(meetingNumber) || 15,
        groupId: groupId,
        createdAt: new Date().toISOString(),
      };
      db.students.push(studentObj);
      db.grades[studentObj.id] = {
        attendanceScore: 100,
        attitudeScore: 85,
        individualScore: 85,
        groupScore: targetGroup.grade ?? 85,
        finalScore: 88,
        letterGrade: 'A-',
      };
    }
  }

  // Add name to group.members if not yet in list
  if (!targetGroup.members.includes(memberName)) {
    targetGroup.members.push(memberName);
  }

  saveDatabase();
  res.json({ success: true, group: targetGroup, student: studentObj });
});

// 10c. Remove member from group (Only Dosen)
app.delete('/api/groups/:id/members', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang menghapus anggota kelompok.' });
  }
  const groupId = Number(req.params.id);
  const targetGroup = db.groups.find(g => g.id === groupId);
  if (!targetGroup) {
    return res.status(404).json({ error: 'Kelompok tidak ditemukan' });
  }

  const { studentName } = req.body;
  if (!studentName) {
    return res.status(400).json({ error: 'Nama mahasiswa wajib dicantumkan' });
  }

  targetGroup.members = targetGroup.members.filter(
    m => m.trim().toUpperCase() !== studentName.trim().toUpperCase()
  );

  saveDatabase();
  res.json({ success: true, group: targetGroup });
});

// 10f. Course / Mata Kuliah Management (Dosen)
app.put('/api/course', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang dapat memperbarui profil mata kuliah dan kampus!' });
  }

  if (!db.courseProfile) {
    db.courseProfile = {
      name: 'Risfa Tri Ulfa, S.Pd., M.Pd., Gr.',
      dosenName: 'Risfa Tri Ulfa',
      dosenTitle: 'S.Pd., M.Pd., Gr.',
      campusName: 'STAI Jarinabi',
      nip: '198806282015032001',
      courseTitle: 'Filsafat Ilmu',
      courseCode: 'MPI-501',
      sks: 3,
      semester: 'Semester Ganjil 2026/2027',
      studyProgram: 'Manajemen Pendidikan Islam (MPI 1)',
      description: 'Mata kuliah ini membahas fondasi ontologis, epistemologis, dan aksiologis keilmuan dalam tata kelola lembaga pendidikan Islam kontemporer.',
    };
  }

  const { dosenName, dosenTitle, campusName, name, ...rest } = req.body;
  
  if (dosenName !== undefined) db.courseProfile.dosenName = dosenName.trim();
  if (dosenTitle !== undefined) db.courseProfile.dosenTitle = dosenTitle.trim();
  if (campusName !== undefined) db.courseProfile.campusName = campusName.trim();

  const activeDosenName = db.courseProfile.dosenName || 'Risfa Tri Ulfa';
  const activeDosenTitle = db.courseProfile.dosenTitle !== undefined ? db.courseProfile.dosenTitle : 'S.Pd., M.Pd., Gr.';
  db.courseProfile.name = activeDosenTitle ? `${activeDosenName}, ${activeDosenTitle}` : activeDosenName;

  db.courseProfile = {
    ...db.courseProfile,
    ...rest,
  };

  // Sync with current course in db.courses
  if (db.courses && db.activeCourseId) {
    const curIdx = db.courses.findIndex(c => c.id === db.activeCourseId);
    if (curIdx >= 0) {
      db.courses[curIdx] = {
        ...db.courses[curIdx],
        title: db.courseProfile.courseTitle,
        code: db.courseProfile.courseCode,
        sks: db.courseProfile.sks,
        semester: db.courseProfile.semester,
        campusName: db.courseProfile.campusName || 'STAI Jarinabi',
        dosenName: db.courseProfile.dosenName || 'Risfa Tri Ulfa',
        dosenTitle: db.courseProfile.dosenTitle || 'S.Pd., M.Pd., Gr.',
        studyProgram: db.courseProfile.studyProgram,
      };
    }
  }

  saveDatabase();
  res.json({ success: true, courseProfile: db.courseProfile, courses: db.courses });
});

// Multi-Course API: Get all courses
app.get(['/api/courses', '/courses'], (req, res) => {
  res.json({
    success: true,
    courses: db.courses || [],
    activeCourseId: db.activeCourseId || 'mk-filsafat-ilmu',
    courseProfile: db.courseProfile,
  });
});

// Helper: Get master students and groups matching RPS per course
function getMasterStudentsAndGroups(
  semesterNum: number = 1,
  courseTitle: string = ''
): { students: Student[]; groups: GroupProject[]; grades: Record<string, StudentGrade> } {
  const titleLower = (courseTitle || '').toLowerCase();

  // Semester 3: Analisis Kebijakan Pendidikan (20 Mahasiswa & 10 Kelompok RPS)
  if (semesterNum === 3 && (titleLower.includes('kebijakan') || titleLower.includes('analisis'))) {
    const students: Student[] = JSON.parse(JSON.stringify(INITIAL_STUDENTS_KEBIJAKAN_SEM3));
    const groups: GroupProject[] = JSON.parse(JSON.stringify(INITIAL_GROUPS_KEBIJAKAN_SEM3));
    const grades: Record<string, StudentGrade> = {};
    students.forEach((s, idx) => {
      grades[s.id] = {
        attendanceScore: 100,
        attitudeScore: 88 + (idx % 4),
        individualScore: 88 + (idx % 4),
        utsScore: 89 + (idx % 3),
        uasScore: 90,
        groupScore: 90,
        finalScore: 90,
        letterGrade: 'A',
      };
    });
    return { students, groups, grades };
  }

  // Semester 3: Konsep Dasar Manajemen Pendidikan Islam (20 Mahasiswa & 12 Kelompok RPS)
  if (semesterNum === 3 || titleLower.includes('konsep') || titleLower.includes('dasar')) {
    const students: Student[] = JSON.parse(JSON.stringify(INITIAL_STUDENTS_SEM3));
    const groups: GroupProject[] = JSON.parse(JSON.stringify(INITIAL_GROUPS_SEM3));
    const grades: Record<string, StudentGrade> = {};
    students.forEach((s, idx) => {
      grades[s.id] = {
        attendanceScore: 100,
        attitudeScore: 88 + (idx % 4),
        individualScore: 88 + (idx % 4),
        utsScore: 89 + (idx % 3),
        uasScore: 90,
        groupScore: 90,
        finalScore: 90,
        letterGrade: 'A',
      };
    });
    return { students, groups, grades };
  }

  // Semester 1 / Filsafat Ilmu: 13 Mahasiswa Resmi (Husna, Ikha, Indah, dkk & 10 Kelompok RPS)
  const students: Student[] = JSON.parse(JSON.stringify(INITIAL_STUDENTS));
  const groups: GroupProject[] = JSON.parse(JSON.stringify(INITIAL_GROUPS));
  const grades: Record<string, StudentGrade> = JSON.parse(JSON.stringify(INITIAL_DATABASE.grades || {}));
  return { students, groups, grades };
}

function generateCohortForSemester(
  semesterNum: number = 1,
  roomName: string = '',
  courseTitle: string = ''
): { students: Student[]; groups: GroupProject[]; grades: Record<string, StudentGrade> } {
  // Default to the 13 students that already exist here in the class
  return getMasterStudentsAndGroups(semesterNum, courseTitle);
}

// Internal helper to switch active course room
function switchActiveCourseInternal(courseId: string): boolean {
  if (!courseId) return false;

  if (!db.courses) db.courses = JSON.parse(JSON.stringify(INITIAL_DATABASE.courses || []));
  let targetCourse = db.courses.find(c => c.id === courseId);
  if (!targetCourse) {
    if (INITIAL_ALL_COURSES_DATA[courseId]) {
      const initC = INITIAL_ALL_COURSES_DATA[courseId];
      targetCourse = {
        id: courseId,
        title: initC.profile?.courseTitle || courseId,
        code: initC.profile?.courseCode || 'MPI',
        sks: initC.profile?.sks || 3,
        semesterNumber: initC.profile?.semesterNumber || 3,
        semester: initC.profile?.semester || 'Semester III',
        roomName: initC.profile?.roomName || 'Kamar MPI',
        campusName: initC.profile?.campusName || 'STAI Jarinabi',
        dosenName: initC.profile?.dosenName || 'Risfa Tri Ulfa',
        dosenTitle: initC.profile?.dosenTitle || 'S.Pd., M.Pd.',
        studyProgram: initC.profile?.studyProgram || 'Manajemen Pendidikan Islam (MPI)',
        totalStudents: initC.students?.length || 20,
        description: initC.profile?.description || '',
      };
      db.courses.push(targetCourse);
    } else if (db.allCoursesData && db.allCoursesData[courseId]) {
      const initC = db.allCoursesData[courseId];
      targetCourse = {
        id: courseId,
        title: initC.profile?.courseTitle || courseId,
        code: initC.profile?.courseCode || 'MPI',
        sks: initC.profile?.sks || 3,
        semesterNumber: initC.profile?.semesterNumber || 3,
        semester: initC.profile?.semester || 'Semester III',
        roomName: initC.profile?.roomName || 'Kamar MPI',
        campusName: initC.profile?.campusName || 'STAI Jarinabi',
        dosenName: initC.profile?.dosenName || 'Risfa Tri Ulfa',
        dosenTitle: initC.profile?.dosenTitle || 'S.Pd., M.Pd.',
        studyProgram: initC.profile?.studyProgram || 'Manajemen Pendidikan Islam (MPI)',
        totalStudents: initC.students?.length || 20,
        description: initC.profile?.description || '',
      };
      db.courses.push(targetCourse);
    } else {
      return false;
    }
  }

  // 1. Save current active course state before switching
  if (db.activeCourseId && db.courseProfile) {
    if (!db.allCoursesData) db.allCoursesData = {};
    db.allCoursesData[db.activeCourseId] = {
      id: db.activeCourseId,
      profile: JSON.parse(JSON.stringify(db.courseProfile)),
      meetings: JSON.parse(JSON.stringify(db.meetings || [])),
      rpsRawText: db.rpsRawText,
      students: JSON.parse(JSON.stringify(db.students || [])),
      groups: JSON.parse(JSON.stringify(db.groups || [])),
      submissions: JSON.parse(JSON.stringify(db.submissions || [])),
      utsQuestions: JSON.parse(JSON.stringify(db.utsQuestions || [])),
      utsSubmissions: JSON.parse(JSON.stringify(db.utsSubmissions || [])),
      uasQuestions: JSON.parse(JSON.stringify(db.uasQuestions || [])),
      uasSubmissions: JSON.parse(JSON.stringify(db.uasSubmissions || [])),
      quizQuestions: JSON.parse(JSON.stringify(db.quizQuestions || [])),
      quizSubmissions: JSON.parse(JSON.stringify(db.quizSubmissions || [])),
      quizSettings: JSON.parse(JSON.stringify(db.quizSettings || {})),
      utsExamSettings: JSON.parse(JSON.stringify(db.utsExamSettings || {})),
      uasExamSettings: JSON.parse(JSON.stringify(db.uasExamSettings || {})),
      attendance: JSON.parse(JSON.stringify(db.attendance || {})),
      grades: JSON.parse(JSON.stringify(db.grades || {})),
    };
  }

  // 2. Load target course state
  db.activeCourseId = courseId;

  if (db.allCoursesData && db.allCoursesData[courseId]) {
    const saved = db.allCoursesData[courseId];
    db.courseProfile = saved.profile;
    db.meetings = saved.meetings;
    db.rpsRawText = saved.rpsRawText || '';
    db.students = saved.students;
    db.groups = saved.groups;
    db.submissions = saved.submissions || [];
    db.utsQuestions = saved.utsQuestions || [];
    db.utsSubmissions = saved.utsSubmissions || [];
    if (saved.uasQuestions && saved.uasQuestions.length > 0) db.uasQuestions = saved.uasQuestions;
    db.uasSubmissions = saved.uasSubmissions || [];
    if (saved.quizQuestions && saved.quizQuestions.length > 0) db.quizQuestions = saved.quizQuestions;
    db.quizSubmissions = saved.quizSubmissions || [];
    if (saved.quizSettings) db.quizSettings = saved.quizSettings;
    if (saved.utsExamSettings) db.utsExamSettings = saved.utsExamSettings;
    if (saved.uasExamSettings) db.uasExamSettings = saved.uasExamSettings;
    db.attendance = saved.attendance || { 1: {} };
    db.grades = saved.grades || {};
  } else if (INITIAL_ALL_COURSES_DATA[courseId]) {
    const initC = INITIAL_ALL_COURSES_DATA[courseId];
    db.courseProfile = initC.profile;
    db.meetings = initC.meetings;
    db.rpsRawText = '';
    db.students = initC.students;
    db.groups = initC.groups;
    db.submissions = initC.submissions || [];
    db.utsQuestions = initC.utsQuestions || [];
    db.utsSubmissions = initC.utsSubmissions || [];
    db.uasQuestions = initC.uasQuestions || [];
    db.uasSubmissions = initC.uasSubmissions || [];
    db.quizQuestions = initC.quizQuestions || [];
    db.quizSubmissions = initC.quizSubmissions || [];
    db.quizSettings = initC.quizSettings || INITIAL_DATABASE.quizSettings;
    db.utsExamSettings = initC.utsExamSettings || INITIAL_DATABASE.utsExamSettings;
    db.uasExamSettings = initC.uasExamSettings || INITIAL_DATABASE.uasExamSettings;
    db.attendance = initC.attendance || { 1: {} };
    db.grades = initC.grades || {};
  } else {
    // Initialize new course room with its own isolated students
    const dosenName = targetCourse.dosenName || db.courseProfile?.dosenName || 'Risfa Tri Ulfa';
    const dosenTitle = targetCourse.dosenTitle || db.courseProfile?.dosenTitle || 'S.Pd., M.Pd.';
    const detectedSemNum = Number(targetCourse.semesterNumber) || (targetCourse.semester.includes('III') || targetCourse.semester.includes('3') ? 3 : 1);
    
    db.courseProfile = {
      name: `${dosenName}, ${dosenTitle}`,
      dosenName,
      dosenTitle,
      campusName: targetCourse.campusName || db.courseProfile?.campusName || 'STAI Jarinabi',
      nip: db.courseProfile?.nip || '198806282015032001',
      courseTitle: targetCourse.title,
      courseCode: targetCourse.code,
      sks: targetCourse.sks || 3,
      semester: targetCourse.semester || (detectedSemNum === 3 ? 'Semester III (Ganjil 2026/2027)' : 'Semester I (Ganjil 2026/2027)'),
      semesterNumber: detectedSemNum,
      roomName: targetCourse.roomName || (detectedSemNum === 3 ? 'Kamar MPI 3 (Ruang Ibnu Khaldun)' : 'Kamar MPI 1 (Ruang Al-Farabi)'),
      studyProgram: targetCourse.studyProgram || 'Manajemen Pendidikan Islam (MPI)',
      classCode: `MPI ${detectedSemNum}`,
      academicYear: '2026/2027',
      totalMeetings: 16,
      description: targetCourse.description || `Mata kuliah ${targetCourse.title} diampu oleh ${dosenName}, ${dosenTitle}.`,
    };

    const cohort = generateCohortForSemester(detectedSemNum, targetCourse.roomName, targetCourse.title);
    db.students = cohort.students;
    db.groups = cohort.groups;
    db.grades = cohort.grades;
    db.submissions = [];
    db.utsSubmissions = [];
    db.uasSubmissions = [];
    db.quizSubmissions = [];
    db.attendance = { 1: {} };

    // Register into allCoursesData
    if (!db.allCoursesData) db.allCoursesData = {};
    db.allCoursesData[courseId] = {
      id: courseId,
      profile: db.courseProfile,
      meetings: db.meetings,
      rpsRawText: '',
      students: db.students,
      groups: db.groups,
      submissions: [],
      utsQuestions: db.utsQuestions,
      utsSubmissions: [],
      uasQuestions: db.uasQuestions,
      uasSubmissions: [],
      quizQuestions: db.quizQuestions,
      quizSubmissions: [],
      quizSettings: db.quizSettings,
      utsExamSettings: db.utsExamSettings,
      uasExamSettings: db.uasExamSettings,
      attendance: db.attendance,
      grades: db.grades,
    };
  }

  saveDatabase();
  return true;
}

// Multi-Course API: Switch active course / kamar
app.post(['/api/courses/switch', '/courses/switch'], (req, res) => {
  const courseId = (req.body?.courseId || req.query?.courseId || req.headers['x-active-course-id']) as string;
  if (!courseId) {
    return res.status(400).json({ error: 'courseId is required' });
  }

  const ok = switchActiveCourseInternal(courseId);
  if (!ok) {
    return res.status(404).json({ error: 'Mata kuliah tidak ditemukan dalam daftar.' });
  }

  const { dosenPassword, ...safeDb } = db;
  res.json({
    success: true,
    activeCourseId: db.activeCourseId,
    courseProfile: db.courseProfile,
    data: safeDb,
  });
});

// Multi-Course API: Add new course / kamar / semester (Dosen)
app.post('/api/courses', async (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang dapat menambahkan mata kuliah baru!' });
  }

  const {
    title,
    code,
    sks = 3,
    semester = 'Semester Ganjil 2026/2027',
    semesterNumber,
    roomName,
    studyProgram = 'Manajemen Pendidikan Islam (MPI)',
    campusName = db.courseProfile?.campusName || 'STAI Jarinabi',
    dosenName = db.courseProfile?.dosenName || 'Risfa Tri Ulfa',
    dosenTitle = db.courseProfile?.dosenTitle || 'S.Pd., M.Pd.',
    description,
    rpsText,
    rpsBase64,
    rpsFilename,
    defaultPresentationFormat = 'auto',
    studentOption = 'use_existing_roster', // 'use_existing_roster' | 'copy_from_course' | 'empty' | 'generate_cohort'
    sourceCourseId,
    customStudents,
  } = req.body;

  if (!title || !title.trim()) {
    return res.status(400).json({ error: 'Nama mata kuliah wajib diisi!' });
  }

  // Extract text from uploaded RPS document file (Word docx / PDF / text) if provided
  let combinedRpsText = (rpsText || '').trim();
  if (rpsBase64 && rpsFilename) {
    try {
      const fileBuffer = Buffer.from(rpsBase64, 'base64');
      const extractedFromFile = await extractTextFromFileBuffer(fileBuffer, rpsFilename);
      if (extractedFromFile && extractedFromFile.trim()) {
        combinedRpsText = extractedFromFile.trim();
      }
    } catch (err) {
      console.warn('Gagal membaca berkas RPS saat menambahkan MK:', err);
    }
  }

  // Determine semester number (e.g. 1, 3, etc.)
  let parsedSemNum = Number(semesterNumber);
  if (!parsedSemNum || isNaN(parsedSemNum)) {
    const semMatch = semester.match(/semester\s*([0-9ivx]+)/i);
    if (semMatch) {
      const semStr = semMatch[1].toUpperCase();
      parsedSemNum = semStr === 'I' || semStr === '1' ? 1
        : semStr === 'II' || semStr === '2' ? 2
        : semStr === 'III' || semStr === '3' ? 3
        : semStr === 'IV' || semStr === '4' ? 4
        : semStr === 'V' || semStr === '5' ? 5
        : semStr === 'VI' || semStr === '6' ? 6
        : 1;
    } else {
      parsedSemNum = 1;
    }
  }

  const safeRoomName = roomName ? roomName.trim() : `Kamar MPI ${parsedSemNum} (Ruang Kelas)`;
  const safeCode = (code || `MPI-${parsedSemNum}0${Math.floor(1 + Math.random() * 9)}`).trim().toUpperCase();
  const courseId = `mk-${Date.now().toString(36)}-${Math.random().toString(36).substr(2, 4)}`;

  const newCourse: CourseSummary = {
    id: courseId,
    title: title.trim(),
    code: safeCode,
    sks: Number(sks) || 3,
    semester: semester.trim(),
    semesterNumber: parsedSemNum,
    roomName: safeRoomName,
    campusName: campusName.trim(),
    dosenName: dosenName.trim(),
    dosenTitle: dosenTitle.trim(),
    studyProgram: studyProgram.trim(),
    totalStudents: 13,
    description: description || `Mata kuliah ${title.trim()} Semester ${parsedSemNum} (${safeRoomName}).`,
  };

  if (!db.courses) db.courses = [];
  db.courses.push(newCourse);

  // Auto-switch to newly created course & archive current course state
  if (db.activeCourseId && db.courseProfile) {
    if (!db.allCoursesData) db.allCoursesData = {};
    db.allCoursesData[db.activeCourseId] = {
      id: db.activeCourseId,
      profile: JSON.parse(JSON.stringify(db.courseProfile)),
      meetings: JSON.parse(JSON.stringify(db.meetings || [])),
      rpsRawText: db.rpsRawText,
      students: JSON.parse(JSON.stringify(db.students || [])),
      groups: JSON.parse(JSON.stringify(db.groups || [])),
      submissions: JSON.parse(JSON.stringify(db.submissions || [])),
      utsQuestions: JSON.parse(JSON.stringify(db.utsQuestions || [])),
      utsSubmissions: JSON.parse(JSON.stringify(db.utsSubmissions || [])),
      uasQuestions: JSON.parse(JSON.stringify(db.uasQuestions || [])),
      uasSubmissions: JSON.parse(JSON.stringify(db.uasSubmissions || [])),
      quizQuestions: JSON.parse(JSON.stringify(db.quizQuestions || [])),
      quizSubmissions: JSON.parse(JSON.stringify(db.quizSubmissions || [])),
      quizSettings: JSON.parse(JSON.stringify(db.quizSettings || {})),
      utsExamSettings: JSON.parse(JSON.stringify(db.utsExamSettings || {})),
      uasExamSettings: JSON.parse(JSON.stringify(db.uasExamSettings || {})),
      attendance: JSON.parse(JSON.stringify(db.attendance || {})),
      grades: JSON.parse(JSON.stringify(db.grades || {})),
    };
  }

  db.activeCourseId = courseId;
  db.courseProfile = {
    name: dosenTitle ? `${dosenName}, ${dosenTitle}` : dosenName,
    dosenName,
    dosenTitle,
    campusName,
    nip: db.courseProfile?.nip || '198806282015032001',
    courseTitle: newCourse.title,
    courseCode: newCourse.code,
    sks: newCourse.sks,
    semester: newCourse.semester,
    semesterNumber: parsedSemNum,
    roomName: safeRoomName,
    studyProgram: newCourse.studyProgram,
    classCode: `MPI ${parsedSemNum}`,
    academicYear: '2026/2027',
    totalMeetings: 16,
    description: description || `Mata kuliah ${newCourse.title} diampu oleh ${dosenName}, ${dosenTitle}.`,
  };

  // Determine students & groups based on studentOption
  let roomStudents: Student[] = [];
  let roomGroups: GroupProject[] = [];
  let roomGrades: Record<string, StudentGrade> = {};

  if (combinedRpsText) {
    const { detectedProfile, detectedMeetings, updatedStudents, detectedGroups } = parseRpsContent(
      combinedRpsText,
      defaultPresentationFormat,
      INITIAL_STUDENTS,
      INITIAL_GROUPS
    );

    db.meetings = detectedMeetings;
    if (updatedStudents && updatedStudents.length > 0) {
      roomStudents = updatedStudents;
    }
    if (detectedGroups && detectedGroups.length > 0) {
      roomGroups = detectedGroups;
    }
    if (!description && detectedProfile.description) {
      db.courseProfile.description = detectedProfile.description;
    }
    db.rpsRawText = combinedRpsText;
  } else {
    // Generate default structured 16 meetings
    const newMeetings: MeetingSchedule[] = [];
    for (let i = 1; i <= 16; i++) {
      const isUts = i === 8;
      const isUas = i === 16;
      newMeetings.push({
        meetingNumber: i,
        dateStr: `Pertemuan ${i}`,
        isoDate: new Date(Date.now() + (i - 1) * 7 * 86400000).toISOString().split('T')[0],
        title: isUts ? 'Ujian Tengah Semester (UTS)' : isUas ? 'Ujian Akhir Semester (UAS)' : `Pertemuan ${i}: Materi Pembelajaran ${newCourse.title}`,
        presenters: [],
        partCodes: [`Pertemuan ${i}`],
        type: isUts ? 'uts' : isUas ? 'uas' : 'kuliah',
        presentationFormat: 'individu',
        taskType: isUts ? 'uts_esai' : isUas ? 'uas_proyek' : 'makalah_ppt',
        description: isUts ? 'Evaluasi penguasaan materi 7 pertemuan awal.' : isUas ? 'Evaluasi akhir dan pengumpulan karya akhir semester.' : `Pembahasan topik perkuliahan ke-${i} mata kuliah ${newCourse.title}.`,
        assignmentDescription: isUts
          ? 'Pengerjaan 5 Soal Essay Evaluasi Tengah Semester (UTS)'
          : isUas
          ? 'Pengerjaan & Publikasi Proyek Video Edukasi Kelompok AI'
          : `Tugas Presentasi (Individu): Pembuatan Makalah & Slide PPT Pertemuan ${i}`,
      });
    }
    db.meetings = newMeetings;
    db.rpsRawText = '';
  }

  // Handle student roster separation (allows different custom student names per course/semester)
  let parsedCustomStudents: Student[] = [];
  if (Array.isArray(customStudents) && customStudents.length > 0) {
    parsedCustomStudents = customStudents.map((item, idx) => {
      if (typeof item === 'string') {
        const nameUpper = item.trim().toUpperCase().replace(/^[0-9]+[\.\-\)\s]+/, '').trim();
        return {
          id: `mhs-${Date.now()}-${idx + 1}-${Math.random().toString(36).substr(2, 4)}`,
          nim: `2026${String(idx + 1).padStart(4, '0')}`,
          name: nameUpper,
          birthPlace: 'Pasuruan',
          birthDate: '2000-01-01',
          address: 'Kabupaten Pasuruan',
          gender: nameUpper.includes('SARI') || nameUpper.includes('LESTARI') || nameUpper.includes('FATIMAH') || nameUpper.includes('DEWI') ? 'Perempuan' : 'Laki-laki',
          phone: `08123456${String(idx + 1).padStart(4, '0')}`,
          rpsPart: `Pertemuan ${((idx % 14) + 1)}`,
          topic: `Materi Perkuliahan #${((idx % 14) + 1)}`,
          meetingNumber: ((idx % 14) + 1),
          groupId: (idx % 10) + 1,
          createdAt: new Date().toISOString(),
        };
      }
      return item;
    });
  } else if (typeof req.body.customStudentNames === 'string' && req.body.customStudentNames.trim()) {
    const rawNames = req.body.customStudentNames
      .split(/[\n,;]+/)
      .map(s => s.trim().replace(/^[0-9]+[\.\-\)\s]+/, '').trim())
      .filter(s => s.length > 1);
    
    parsedCustomStudents = rawNames.map((name, idx) => {
      const nameUpper = name.toUpperCase();
      return {
        id: `mhs-${Date.now()}-${idx + 1}-${Math.random().toString(36).substr(2, 4)}`,
        nim: `2026${String(idx + 1).padStart(4, '0')}`,
        name: nameUpper,
        birthPlace: 'Pasuruan',
        birthDate: '2000-01-01',
        address: 'Kabupaten Pasuruan',
        gender: nameUpper.includes('SARI') || nameUpper.includes('LESTARI') || nameUpper.includes('FATIMAH') || nameUpper.includes('DEWI') ? 'Perempuan' : 'Laki-laki',
        phone: `08123456${String(idx + 1).padStart(4, '0')}`,
        rpsPart: `Pertemuan ${((idx % 14) + 1)}`,
        topic: `Materi Perkuliahan #${((idx % 14) + 1)}`,
        meetingNumber: ((idx % 14) + 1),
        groupId: (idx % 10) + 1,
        createdAt: new Date().toISOString(),
      };
    });
  }

  if (parsedCustomStudents.length > 0) {
    roomStudents = parsedCustomStudents;
    // Auto-generate balanced groups for the new custom students
    const groupCount = Math.min(10, Math.max(1, Math.ceil(roomStudents.length / 2)));
    roomGroups = [];
    for (let g = 1; g <= groupCount; g++) {
      const members = roomStudents.filter((_, idx) => (idx % groupCount) + 1 === g).map(s => s.name);
      roomGroups.push({
        id: g,
        name: `Kelompok ${g}`,
        title: `Kajian Topik Kelompok ${g}: ${newCourse.title}`,
        members,
        description: `Analisis & Pembahasan Tema Kelompok ${g} mata kuliah ${newCourse.title}`,
        toolsSuggested: 'Canva / PowerPoint / Google Slides',
        grade: 85,
      });
    }
    // Initialize student grades
    roomGrades = {};
    roomStudents.forEach(s => {
      roomGrades[s.id] = {
        attendanceScore: 100,
        attitudeScore: 85,
        individualScore: 85,
        groupScore: 85,
        utsScore: 85,
        uasScore: 85,
        finalScore: 88,
        letterGrade: 'A-',
      };
    });
  } else if (studentOption === 'copy_from_course' && sourceCourseId && db.allCoursesData?.[sourceCourseId]) {
    const src = db.allCoursesData[sourceCourseId];
    roomStudents = JSON.parse(JSON.stringify(src.students || []));
    roomGroups = JSON.parse(JSON.stringify(src.groups || []));
    roomGrades = JSON.parse(JSON.stringify(src.grades || {}));
  } else if (studentOption === 'empty') {
    roomStudents = [];
    roomGroups = [];
    roomGrades = {};
  } else {
    // Default & 'use_existing_roster':
    // Always use the 13 students that already exist here in the system (Husna, Ikha, Indah, dkk)
    const master = getMasterStudentsAndGroups(parsedSemNum, newCourse.title);
    roomStudents = master.students;
    roomGroups = master.groups;
    roomGrades = master.grades;
  }

  db.students = roomStudents;
  db.groups = roomGroups;
  db.grades = roomGrades;
  newCourse.totalStudents = roomStudents.length;

  db.submissions = [];
  db.utsSubmissions = [];
  db.uasSubmissions = [];
  db.quizSubmissions = [];
  db.attendance = { 1: {} };

  // Generate UTS questions tailored for this course
  db.utsQuestions = generateUtsQuestionsFromMeetings(db.meetings, db.courseProfile);
  db.uasQuestions = generateUasQuestionsFromMeetings(db.meetings, db.courseProfile);

  // Directly save newly created room state into allCoursesData
  if (!db.allCoursesData) db.allCoursesData = {};
  db.allCoursesData[courseId] = {
    id: courseId,
    profile: JSON.parse(JSON.stringify(db.courseProfile)),
    meetings: JSON.parse(JSON.stringify(db.meetings || [])),
    rpsRawText: db.rpsRawText,
    students: JSON.parse(JSON.stringify(db.students || [])),
    groups: JSON.parse(JSON.stringify(db.groups || [])),
    submissions: [],
    utsQuestions: JSON.parse(JSON.stringify(db.utsQuestions || [])),
    utsSubmissions: [],
    uasQuestions: JSON.parse(JSON.stringify(db.uasQuestions || [])),
    uasSubmissions: [],
    quizQuestions: JSON.parse(JSON.stringify(db.quizQuestions || [])),
    quizSubmissions: [],
    quizSettings: JSON.parse(JSON.stringify(db.quizSettings || {})),
    utsExamSettings: JSON.parse(JSON.stringify(db.utsExamSettings || {})),
    uasExamSettings: JSON.parse(JSON.stringify(db.uasExamSettings || {})),
    attendance: JSON.parse(JSON.stringify(db.attendance || { 1: {} })),
    grades: JSON.parse(JSON.stringify(db.grades || {})),
  };

  saveDatabase();
  res.json({
    success: true,
    course: newCourse,
    courses: db.courses,
    activeCourseId: db.activeCourseId,
    courseProfile: db.courseProfile,
    data: db,
  });
});

// Multi-Course API: Edit Course (Dosen only)
const handleEditCourse = (req: express.Request, res: express.Response) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berwenang mengubah data mata kuliah!' });
  }

  const courseId = req.params.courseId || req.body.courseId;
  if (!courseId) {
    return res.status(400).json({ error: 'courseId wajib disertakan' });
  }

  const {
    title,
    code,
    sks,
    semester,
    semesterNumber,
    roomName,
    campusName,
    dosenName,
    dosenTitle,
    studyProgram,
    description,
  } = req.body;

  if (!title || !title.trim()) {
    return res.status(400).json({ error: 'Nama mata kuliah wajib diisi!' });
  }

  const targetCourse = (db.courses || []).find(c => c.id === courseId);
  if (!targetCourse) {
    return res.status(404).json({ error: 'Mata kuliah tidak ditemukan dalam database.' });
  }

  // Update course summary
  targetCourse.title = title.trim();
  if (code) targetCourse.code = code.trim();
  if (sks !== undefined) targetCourse.sks = Number(sks);
  if (semester) targetCourse.semester = semester.trim();
  if (semesterNumber !== undefined) targetCourse.semesterNumber = Number(semesterNumber);
  if (roomName) targetCourse.roomName = roomName.trim();
  if (campusName) targetCourse.campusName = campusName.trim();
  if (dosenName) targetCourse.dosenName = dosenName.trim();
  if (dosenTitle !== undefined) targetCourse.dosenTitle = dosenTitle.trim();
  if (studyProgram) targetCourse.studyProgram = studyProgram.trim();
  if (description !== undefined) targetCourse.description = description.trim();

  // Update room profile in allCoursesData
  if (db.allCoursesData && db.allCoursesData[courseId]) {
    const p: any = db.allCoursesData[courseId].profile || {};
    p.courseTitle = targetCourse.title;
    if (code) p.courseCode = targetCourse.code;
    if (sks !== undefined) p.sks = targetCourse.sks;
    if (semester) p.semester = targetCourse.semester;
    if (semesterNumber !== undefined) p.semesterNumber = targetCourse.semesterNumber;
    if (roomName) p.roomName = targetCourse.roomName;
    if (campusName) p.campusName = targetCourse.campusName;
    if (dosenName) p.dosenName = targetCourse.dosenName;
    if (dosenTitle !== undefined) p.dosenTitle = targetCourse.dosenTitle;
    if (studyProgram) p.studyProgram = targetCourse.studyProgram;
    if (description !== undefined) p.description = targetCourse.description;
    db.allCoursesData[courseId].profile = p;
  }

  // Update active courseProfile if currently active
  if (db.activeCourseId === courseId && db.courseProfile) {
    db.courseProfile.courseTitle = targetCourse.title;
    if (code) db.courseProfile.courseCode = targetCourse.code;
    if (sks !== undefined) db.courseProfile.sks = targetCourse.sks;
    if (semester) db.courseProfile.semester = targetCourse.semester;
    if (semesterNumber !== undefined) db.courseProfile.semesterNumber = targetCourse.semesterNumber;
    if (roomName) db.courseProfile.roomName = targetCourse.roomName;
    if (campusName) db.courseProfile.campusName = targetCourse.campusName;
    if (dosenName) db.courseProfile.dosenName = targetCourse.dosenName;
    if (dosenTitle !== undefined) db.courseProfile.dosenTitle = targetCourse.dosenTitle;
    if (studyProgram) db.courseProfile.studyProgram = targetCourse.studyProgram;
    if (description !== undefined) db.courseProfile.description = targetCourse.description;
  }

  saveDatabase();
  res.json({
    success: true,
    message: `Data mata kuliah "${targetCourse.title}" berhasil diperbarui!`,
    course: targetCourse,
    courses: db.courses,
    activeCourseId: db.activeCourseId,
    courseProfile: db.courseProfile,
    data: db,
  });
};

app.put('/api/courses/:courseId', handleEditCourse);
app.post('/api/courses/edit', handleEditCourse);

// Endpoint: Sync Course Roster and Groups with Master RPS Students & Groups
app.post('/api/courses/sync-roster', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berwenang mensinkronkan kelompok!' });
  }

  const { courseId } = req.body;
  const targetId = courseId || db.activeCourseId;
  const targetCourse = (db.courses || []).find(c => c.id === targetId);
  const semNum = targetCourse?.semesterNumber || (targetCourse?.semester?.includes('3') || targetCourse?.semester?.includes('III') ? 3 : 1);
  const cTitle = targetCourse?.title || db.courseProfile?.courseTitle || '';

  const master = getMasterStudentsAndGroups(semNum, cTitle);

  if (targetId === db.activeCourseId) {
    db.students = JSON.parse(JSON.stringify(master.students));
    db.groups = JSON.parse(JSON.stringify(master.groups));
    if (!db.grades || Object.keys(db.grades).length === 0) {
      db.grades = JSON.parse(JSON.stringify(master.grades));
    }
  }

  if (!db.allCoursesData) db.allCoursesData = {};
  if (!db.allCoursesData[targetId]) {
    db.allCoursesData[targetId] = {
      id: targetId,
      profile: JSON.parse(JSON.stringify(targetCourse ? { ...INITIAL_COURSE_PROFILE, courseTitle: targetCourse.title, courseCode: targetCourse.code, semester: targetCourse.semester, semesterNumber: semNum } : INITIAL_COURSE_PROFILE)),
      meetings: JSON.parse(JSON.stringify(semNum === 3 ? (cTitle.toLowerCase().includes('kebijakan') ? INITIAL_MEETINGS_KEBIJAKAN_SEM3 : INITIAL_MEETINGS_SEM3) : INITIAL_MEETINGS)),
      students: [],
      groups: [],
      submissions: [],
      utsQuestions: [],
      utsSubmissions: [],
      attendance: { 1: {} },
      grades: {},
    };
  }

  db.allCoursesData[targetId].students = JSON.parse(JSON.stringify(master.students));
  db.allCoursesData[targetId].groups = JSON.parse(JSON.stringify(master.groups));
  if (!db.allCoursesData[targetId].grades || Object.keys(db.allCoursesData[targetId].grades).length === 0) {
    db.allCoursesData[targetId].grades = JSON.parse(JSON.stringify(master.grades));
  }

  if (targetCourse) {
    targetCourse.totalStudents = master.students.length;
  }

  const isFilsafat = semNum === 1 || cTitle.toLowerCase().includes('filsafat');
  const isKebijakan = cTitle.toLowerCase().includes('kebijakan') || cTitle.toLowerCase().includes('analisis');
  const syncLabel = isFilsafat
    ? '13 Mahasiswa resmi & 10 Kelompok RPS Filsafat Ilmu (Husna, Ikha, Indah, dkk)'
    : isKebijakan
    ? '20 Mahasiswa resmi & 10 Kelompok RPS Analisis Kebijakan Pendidikan (Heru, Sofyan, ... Zaima)'
    : '20 Mahasiswa resmi & 12 Kelompok RPS Konsep Dasar MPI (Heru, Sofyan, ... Zaima)';

  saveDatabase();
  res.json({
    success: true,
    message: `${syncLabel} berhasil disinkronkan sesuai RPS!`,
    students: db.students,
    groups: db.groups,
    courses: db.courses,
    activeCourseId: db.activeCourseId,
    courseProfile: db.courseProfile,
    data: db,
  });
});

// Multi-Course & Multi-Semester API: Tambah Semester Baru beserta MK di dalamnya
app.post('/api/semesters', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berwenang menambah semester!' });
  }

  const {
    semesterNumber,
    semesterName,
    academicYear = '2026/2027',
    courseTitle,
    courseCode,
    sks = 3,
    roomName,
  } = req.body;

  const semNum = Number(semesterNumber) || 1;
  const roman =
    semNum === 1 ? 'I' : semNum === 2 ? 'II' : semNum === 3 ? 'III' : semNum === 4 ? 'IV' : semNum === 5 ? 'V' : semNum === 6 ? 'VI' : semNum === 7 ? 'VII' : semNum === 8 ? 'VIII' : String(semNum);
  const ganjilGenap = semNum % 2 === 1 ? 'Ganjil' : 'Genap';
  const finalSemesterName = (semesterName && semesterName.trim()) || `Semester ${roman} (${ganjilGenap} ${academicYear})`;

  const title = (courseTitle && courseTitle.trim()) || `Mata Kuliah Utama Semester ${semNum}`;
  const code = (courseCode && courseCode.trim()) || `MPI-${semNum}01`;
  const room = (roomName && roomName.trim()) || `Kamar MPI ${semNum} (Ruang Kelas)`;
  const courseId = `mk-sem${semNum}-${Date.now().toString(36)}`;

  const newCourse: CourseSummary = {
    id: courseId,
    title,
    code,
    sks: Number(sks) || 3,
    semester: finalSemesterName,
    semesterNumber: semNum,
    roomName: room,
    campusName: db.courseProfile?.campusName || 'STAI Jarinabi',
    dosenName: db.courseProfile?.dosenName || 'Risfa Tri Ulfa',
    dosenTitle: db.courseProfile?.dosenTitle || 'S.Pd., M.Pd.',
    studyProgram: db.courseProfile?.studyProgram || 'Manajemen Pendidikan Islam (MPI)',
    totalStudents: 13,
    description: `Mata kuliah ${title} ${finalSemesterName} (${room}).`,
  };

  if (!db.courses) db.courses = [];
  db.courses.push(newCourse);

  const master = getMasterStudentsAndGroups(semNum, title);
  if (!db.allCoursesData) db.allCoursesData = {};
  db.allCoursesData[courseId] = {
    id: courseId,
    profile: {
      ...newCourse,
      name: `${newCourse.dosenName}, ${newCourse.dosenTitle}`,
      nip: '198806282015032001',
      courseTitle: title,
      courseCode: code,
      totalMeetings: 16,
      classCode: `MPI ${semNum}`,
      academicYear,
    },
    meetings: INITIAL_MEETINGS.map(m => ({ ...m, title: `Pertemuan ${m.meetingNumber}: ${title}` })),
    rpsRawText: '',
    students: master.students,
    groups: master.groups,
    submissions: [],
    utsQuestions: generateUtsQuestionsFromMeetings([], db.courseProfile),
    utsSubmissions: [],
    uasQuestions: [],
    uasSubmissions: [],
    quizQuestions: INITIAL_QUIZ_QUESTIONS,
    quizSubmissions: [],
    quizSettings: db.quizSettings || INITIAL_DATABASE.quizSettings,
    utsExamSettings: db.utsExamSettings || INITIAL_DATABASE.utsExamSettings,
    uasExamSettings: db.uasExamSettings || INITIAL_DATABASE.uasExamSettings,
    attendance: { 1: {} },
    grades: master.grades,
  };

  saveDatabase();

  const { dosenPassword, ...safeDb } = db;
  res.json({
    success: true,
    message: `Semester ${semNum} dan mata kuliah "${title}" berhasil ditambahkan dengan 13 Mahasiswa & 10 Kelompok resmi!`,
    course: newCourse,
    courses: db.courses,
    activeCourseId: db.activeCourseId,
    data: safeDb,
  });
});

// Helper for course deletion
function executeCourseDelete(courseId: string): { success: boolean; error?: string } {
  if (!db.courses || db.courses.length <= 1) {
    return { success: false, error: 'Tidak dapat menghapus. Harus ada minimal 1 mata kuliah aktif!' };
  }

  db.courses = db.courses.filter(c => c.id !== courseId);
  if (db.allCoursesData) {
    delete db.allCoursesData[courseId];
  }

  // If deleted course was active, switch to first remaining course
  if (db.activeCourseId === courseId) {
    db.activeCourseId = db.courses[0].id;
    if (db.allCoursesData && db.allCoursesData[db.activeCourseId]) {
      const saved = db.allCoursesData[db.activeCourseId];
      db.courseProfile = saved.profile;
      db.meetings = saved.meetings;
      db.students = saved.students;
      db.grades = saved.grades || {};
      db.attendance = saved.attendance || { 1: {} };
      db.submissions = saved.submissions || [];
      db.utsQuestions = saved.utsQuestions || [];
      db.utsSubmissions = saved.utsSubmissions || [];
      if (saved.uasQuestions && saved.uasQuestions.length > 0) db.uasQuestions = saved.uasQuestions;
      db.uasSubmissions = saved.uasSubmissions || [];
      if (saved.quizQuestions && saved.quizQuestions.length > 0) db.quizQuestions = saved.quizQuestions;
      db.quizSubmissions = saved.quizSubmissions || [];
      if (saved.quizSettings) db.quizSettings = saved.quizSettings;
      if (saved.utsExamSettings) db.utsExamSettings = saved.utsExamSettings;
      if (saved.uasExamSettings) db.uasExamSettings = saved.uasExamSettings;
      db.groups = saved.groups || [];
    }
  }

  saveDatabase();
  return { success: true };
}

// Multi-Course API: Delete course via DELETE
app.delete('/api/courses/:id', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang dapat menghapus mata kuliah!' });
  }

  const { id } = req.params;
  const result = executeCourseDelete(id);
  if (!result.success) {
    return res.status(400).json({ error: result.error });
  }

  const { dosenPassword, ...safeDb } = db;
  res.json({
    success: true,
    courses: db.courses,
    activeCourseId: db.activeCourseId,
    courseProfile: db.courseProfile,
    data: safeDb,
  });
});

// Multi-Course API: Delete course via POST (firewall/iframe safe fallback)
app.post('/api/courses/delete', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang dapat menghapus mata kuliah!' });
  }

  const courseId = req.body?.courseId || req.body?.id;
  if (!courseId) {
    return res.status(400).json({ error: 'courseId wajib disertakan' });
  }

  const result = executeCourseDelete(courseId);
  if (!result.success) {
    return res.status(400).json({ error: result.error });
  }

  const { dosenPassword, ...safeDb } = db;
  res.json({
    success: true,
    courses: db.courses,
    activeCourseId: db.activeCourseId,
    courseProfile: db.courseProfile,
    data: safeDb,
  });
});

// Semester Transition & Archival API
app.post('/api/semester/transition', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang berwenang melakukan pindah semester!' });
  }

  const {
    newSemesterName,
    academicYear = '2026/2027',
    rpsText,
    archiveCurrent = true,
    resetSubmissions = true,
  } = req.body;

  if (!newSemesterName || !newSemesterName.trim()) {
    return res.status(400).json({ error: 'Nama semester baru wajib diisi (contoh: Semester Genap 2026/2027)!' });
  }

  // 1. Archive current semester data
  if (archiveCurrent) {
    if (!db.archivedSemesters) db.archivedSemesters = [];
    const archiveRecord: ArchivedSemester = {
      id: `arch-${Date.now()}`,
      courseId: db.activeCourseId || 'mk-filsafat-ilmu',
      courseTitle: db.courseProfile?.courseTitle || 'Filsafat Ilmu',
      courseCode: db.courseProfile?.courseCode || 'MPI-501',
      semesterName: db.courseProfile?.semester || 'Semester Ganjil 2026/2027',
      academicYear: db.courseProfile?.academicYear || '2026/2027',
      archivedAt: new Date().toISOString(),
      totalStudents: db.students?.length || 0,
      students: JSON.parse(JSON.stringify(db.students || [])),
      groups: JSON.parse(JSON.stringify(db.groups || [])),
      grades: JSON.parse(JSON.stringify(db.grades || {})),
      attendance: JSON.parse(JSON.stringify(db.attendance || {})),
      meetings: JSON.parse(JSON.stringify(db.meetings || [])),
      submissions: JSON.parse(JSON.stringify(db.submissions || [])),
      utsSubmissions: JSON.parse(JSON.stringify(db.utsSubmissions || [])),
      uasSubmissions: JSON.parse(JSON.stringify(db.uasSubmissions || [])),
      quizSubmissions: JSON.parse(JSON.stringify(db.quizSubmissions || [])),
    };
    db.archivedSemesters.unshift(archiveRecord);
  }

  // 2. Update semester in course profile & active course
  if (db.courseProfile) {
    db.courseProfile.semester = newSemesterName.trim();
    db.courseProfile.academicYear = academicYear.trim();
  }

  if (db.courses && db.activeCourseId) {
    const curCourse = db.courses.find(c => c.id === db.activeCourseId);
    if (curCourse) {
      curCourse.semester = newSemesterName.trim();
    }
  }

  // 3. If new RPS is provided, parse and update meetings
  if (rpsText && rpsText.trim()) {
    db.rpsRawText = rpsText;
    const lines = rpsText.split('\n');
    const parsedMeetings: MeetingSchedule[] = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;
      const match = line.match(/(?:pertemuan|minggu|sesi)\s*(\d+)[-:\s]*(.*)/i) ||
                    line.match(/^(\d{1,2})\.\s+(.*)/);
      if (match) {
        const num = Number(match[1]);
        const restTitle = (match[2] || '').trim();
        const isUts = num === 8 || /uts|tengah\s*semester/i.test(restTitle);
        const isUas = num === 16 || /uas|akhir\s*semester/i.test(restTitle);

        parsedMeetings.push({
          meetingNumber: num,
          dateStr: `Pertemuan ${num}`,
          isoDate: new Date(Date.now() + (num - 1) * 7 * 86400000).toISOString().split('T')[0],
          title: restTitle || `Materi Pertemuan ${num}`,
          presenters: [],
          partCodes: [`Pertemuan ${num}`],
          type: isUts ? 'uts' : isUas ? 'uas' : 'kuliah',
          presentationFormat: /kelompok/i.test(restTitle) ? 'kelompok' : 'individu',
          taskType: isUts ? 'uts_esai' : isUas ? 'uas_proyek' : 'makalah_ppt',
          description: restTitle || `Kajian materi RPS semester baru pertemuan ${num}`,
        });
      }
    }

    if (parsedMeetings.length > 0) {
      db.meetings = parsedMeetings.sort((a, b) => a.meetingNumber - b.meetingNumber);
    }
  }

  // 4. Reset submissions & attendance for fresh semester if requested
  if (resetSubmissions) {
    db.submissions = [];
    db.utsSubmissions = [];
    db.attendance = { 1: {} };
    // Reset grade records to default 100/85
    const resetGrades: Record<string, StudentGrade> = {};
    (db.students || []).forEach(s => {
      resetGrades[s.id] = {
        attendanceScore: 100,
        attitudeScore: 85,
        individualScore: 85,
        utsScore: 85,
        uasScore: 85,
        groupScore: 85,
        finalScore: 88,
        letterGrade: 'A-',
      };
    });
    db.grades = resetGrades;
  }

  saveDatabase();
  res.json({
    success: true,
    message: `Berhasil pindah ke ${newSemesterName}. Data lama berhasil diarsipkan.`,
    courseProfile: db.courseProfile,
    archivedCount: db.archivedSemesters?.length || 0,
    data: db,
  });
});

// Get Archived Semesters
app.get('/api/semester/archives', (req, res) => {
  res.json({
    success: true,
    archives: db.archivedSemesters || [],
  });
});

// Restore Archived Semester
app.post('/api/semester/restore/:archiveId', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang memulihkan arsip semester.' });
  }
  const archiveId = req.params.archiveId;
  const targetArchive = (db.archivedSemesters || []).find(a => a.id === archiveId);

  if (!targetArchive) {
    return res.status(404).json({ error: 'Arsip semester tidak ditemukan.' });
  }

  // Restore snapshots safely
  if (targetArchive.students) db.students = JSON.parse(JSON.stringify(targetArchive.students));
  if (targetArchive.groups) db.groups = JSON.parse(JSON.stringify(targetArchive.groups));
  if (targetArchive.grades) db.grades = JSON.parse(JSON.stringify(targetArchive.grades));
  if (targetArchive.attendance) db.attendance = JSON.parse(JSON.stringify(targetArchive.attendance));
  if (targetArchive.meetings) db.meetings = JSON.parse(JSON.stringify(targetArchive.meetings));
  if (targetArchive.submissions) db.submissions = JSON.parse(JSON.stringify(targetArchive.submissions));
  if (targetArchive.utsSubmissions) db.utsSubmissions = JSON.parse(JSON.stringify(targetArchive.utsSubmissions));
  if (targetArchive.uasSubmissions) db.uasSubmissions = JSON.parse(JSON.stringify(targetArchive.uasSubmissions));
  if (targetArchive.quizSubmissions) db.quizSubmissions = JSON.parse(JSON.stringify(targetArchive.quizSubmissions));

  if (db.courseProfile) {
    db.courseProfile.semester = targetArchive.semesterName;
    db.courseProfile.academicYear = targetArchive.academicYear;
  }

  saveDatabase();
  res.json({
    success: true,
    message: `Semester ${targetArchive.semesterName} berhasil dipulihkan dengan data utuh.`,
    data: db,
  });
});

// Student to Dosen Messages Endpoints
app.get('/api/messages', (req, res) => {
  res.json({
    success: true,
    messages: db.messages || [],
  });
});

app.post('/api/messages', (req, res) => {
  const {
    studentId,
    studentName,
    studentNim,
    category,
    subject,
    content,
    taskTitle,
    meetingNumber,
    attachmentLink,
  } = req.body;

  if (!studentName || !content) {
    return res.status(400).json({ error: 'Nama mahasiswa dan isi pesan wajib diisi.' });
  }

  if (!db.messages) db.messages = [];

  const newMessage = {
    id: `msg-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
    studentId: studentId || 'anon',
    studentName,
    studentNim: studentNim || '-',
    category: category || 'tugas_makalah',
    subject: subject || 'Pemberitahuan Tugas Mahasiswa',
    content,
    taskTitle,
    meetingNumber: meetingNumber ? Number(meetingNumber) : undefined,
    attachmentLink,
    submittedAt: new Date().toISOString(),
    read: false,
    replied: false,
  };

  db.messages.unshift(newMessage);

  // Also broadcast submission / message to notifications feed for Dosen
  const notifCategory = category === 'tugas_uts' ? 'UTS' : category === 'tugas_uas' ? 'UAS' : 'Tugas';
  const newNotif = {
    id: `notif-msg-${Date.now()}`,
    type: 'submission' as const,
    title: `Pesan Mahasiswa: ${studentName}`,
    message: `${subject}: "${content.substring(0, 80)}${content.length > 80 ? '...' : ''}"`,
    timestamp: new Date().toISOString(),
    studentName,
    studentId,
    targetTab: 'portal-dosen',
    read: false,
    urgent: false,
  };

  // Keep in memory and save
  saveDatabase();

  res.json({
    success: true,
    message: 'Pesan berhasil dikirimkan ke Dosen.',
    data: newMessage,
    notification: newNotif,
  });
});

app.put('/api/messages/:id/read', (req, res) => {
  const msgId = req.params.id;
  if (!db.messages) db.messages = [];
  const msg = db.messages.find(m => m.id === msgId);
  if (!msg) {
    return res.status(404).json({ error: 'Pesan tidak ditemukan' });
  }
  msg.read = true;
  saveDatabase();
  res.json({ success: true, message: 'Status pesan diubah menjadi dibaca.', data: msg });
});

app.put('/api/messages/:id/reply', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang membalas pesan.' });
  }
  const msgId = req.params.id;
  const { replyText } = req.body;
  if (!db.messages) db.messages = [];
  const msg = db.messages.find(m => m.id === msgId);
  if (!msg) {
    return res.status(404).json({ error: 'Pesan tidak ditemukan' });
  }
  msg.read = true;
  msg.replied = true;
  msg.replyText = replyText || 'Tugas telah dikonfirmasi dan diterima oleh Dosen.';
  msg.repliedAt = new Date().toISOString();
  saveDatabase();
  res.json({ success: true, message: 'Balasan berhasil dikirimkan kepada mahasiswa.', data: msg });
});

// 10f. Messages API: Clear messages (supports DELETE and POST fallback)
const handleClearMessages = (req: express.Request, res: express.Response) => {
  const mode = (req.query.mode as string) || req.body?.mode || 'all'; // 'all' | 'read_only' | 'by_student'
  const studentId = (req.query.studentId as string) || req.body?.studentId;

  if (!db.messages) db.messages = [];
  const initialCount = db.messages.length;

  if (mode === 'read_only') {
    db.messages = db.messages.filter(m => !m.read);
  } else if (mode === 'by_student' && studentId) {
    db.messages = db.messages.filter(m => m.studentId !== studentId && m.studentNim !== studentId);
  } else {
    // Mode 'all' or fallback
    if (studentId) {
      db.messages = db.messages.filter(m => m.studentId !== studentId && m.studentNim !== studentId);
    } else {
      db.messages = [];
    }
  }

  saveDatabase();
  const deletedCount = initialCount - db.messages.length;
  return res.json({
    success: true,
    message: `Kotak pesan berhasil dibersihkan (${deletedCount} pesan dihapus).`,
    deletedCount,
    remainingCount: db.messages.length,
  });
};

app.delete('/api/messages', handleClearMessages);
app.post('/api/messages/delete-all', handleClearMessages);
app.post('/api/messages/clear', handleClearMessages);
app.post('/api/messages/delete', handleClearMessages);

// 10f-2. Messages API: Delete single message (supports DELETE and POST fallback)
const handleDeleteSingleMessage = (req: express.Request, res: express.Response) => {
  const msgId = req.params.id;
  if (!db.messages) db.messages = [];
  const targetIndex = db.messages.findIndex(m => m.id === msgId);
  if (targetIndex !== -1) {
    db.messages.splice(targetIndex, 1);
    saveDatabase();
  }

  return res.json({ success: true, message: 'Pesan berhasil dihapus dari sistem.' });
};

app.delete('/api/messages/:id', handleDeleteSingleMessage);
app.post('/api/messages/:id/delete', handleDeleteSingleMessage);

// 10g. Edit Meeting / Schedule & Tanggal Presensi
app.put('/api/meetings/:meetingNumber', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang mengedit pertemuan RPS.' });
  }
  const meetingNum = Number(req.params.meetingNumber);
  const meeting = (db.meetings || []).find(m => m.meetingNumber === meetingNum);
  if (!meeting) {
    return res.status(404).json({ error: 'Pertemuan tidak ditemukan' });
  }

  const {
    dateStr,
    isoDate,
    title,
    description,
    presenters,
    presentationFormat,
    groupName,
    groupId,
    type,
    taskType,
    assignmentDescription,
  } = req.body;

  if (dateStr !== undefined) meeting.dateStr = dateStr;
  if (isoDate !== undefined) meeting.isoDate = isoDate;
  if (title !== undefined) meeting.title = title;
  if (description !== undefined) meeting.description = description;
  if (presenters !== undefined) meeting.presenters = Array.isArray(presenters) ? presenters : [presenters];
  if (presentationFormat !== undefined) meeting.presentationFormat = presentationFormat;
  if (groupName !== undefined) meeting.groupName = groupName;
  if (groupId !== undefined) meeting.groupId = Number(groupId);
  if (type !== undefined) meeting.type = type;
  if (taskType !== undefined) meeting.taskType = taskType;
  if (assignmentDescription !== undefined) meeting.assignmentDescription = assignmentDescription;

  saveDatabase();
  res.json({ success: true, meeting });
});

// 10h. Add New Meeting (Only Dosen)
app.post('/api/meetings', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang menambah pertemuan RPS.' });
  }
  const nextNum = (db.meetings || []).length > 0
    ? Math.max(...db.meetings.map(m => m.meetingNumber)) + 1
    : 1;

  const newMeeting: MeetingSchedule = {
    meetingNumber: req.body.meetingNumber ? Number(req.body.meetingNumber) : nextNum,
    dateStr: req.body.dateStr || `Sabtu, ${nextNum} Oktober 2026`,
    isoDate: req.body.isoDate || new Date().toISOString().split('T')[0],
    title: req.body.title || `Pertemuan ${nextNum}: Materi Pembelajaran Baru`,
    presenters: req.body.presenters || [],
    partCodes: req.body.partCodes || [`Pertemuan ${nextNum}`],
    type: req.body.type || 'kuliah',
    description: req.body.description || 'Kajian materi perkuliahan sesuai RPS.',
    presentationFormat: req.body.presentationFormat || 'individu',
    groupName: req.body.groupName,
    groupId: req.body.groupId ? Number(req.body.groupId) : undefined,
    taskType: req.body.taskType || 'makalah_ppt',
    assignmentDescription: req.body.assignmentDescription,
  };

  db.meetings.push(newMeeting);
  db.meetings.sort((a, b) => a.meetingNumber - b.meetingNumber);

  saveDatabase();
  res.json({ success: true, meeting: newMeeting });
});

// 10i. Delete Meeting (Only Dosen)
app.delete('/api/meetings/:meetingNumber', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang menghapus pertemuan RPS.' });
  }
  const meetingNum = Number(req.params.meetingNumber);
  const initialLen = (db.meetings || []).length;
  db.meetings = (db.meetings || []).filter(m => m.meetingNumber !== meetingNum);

  saveDatabase();
  res.json({ success: true, deleted: initialLen - db.meetings.length });
});

// 10i-2. Add Presenter / Student to Meeting Presentation Group (Dosen & Mahasiswa)
app.post('/api/meetings/:meetingNumber/presenters', (req, res) => {
  const meetingNum = Number(req.params.meetingNumber);
  const meeting = (db.meetings || []).find(m => m.meetingNumber === meetingNum);
  if (!meeting) {
    return res.status(404).json({ error: `Pertemuan ${meetingNum} tidak ditemukan` });
  }

  const { studentId, studentName, nim, rpsPart, topic, presentationFormat, groupName } = req.body;

  let addedStudentName = studentName ? String(studentName).trim().toUpperCase() : '';

  // If studentId provided, lookup student
  if (studentId) {
    const existing = (db.students || []).find(s => s.id === studentId);
    if (existing) {
      existing.meetingNumber = meetingNum;
      if (rpsPart) existing.rpsPart = rpsPart;
      else existing.rpsPart = `Pertemuan ${meetingNum}`;
      if (topic) existing.topic = topic;
      addedStudentName = existing.name;
    }
  } else if (addedStudentName) {
    // Check if student with this name already exists
    let existing = (db.students || []).find(s => s.name.toUpperCase() === addedStudentName);
    if (existing) {
      existing.meetingNumber = meetingNum;
      if (topic) existing.topic = topic;
    } else {
      // Create new student
      const newStd: Student = {
        id: `std-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
        nim: (nim || `2026${String(db.students.length + 1).padStart(4, '0')}`).trim(),
        name: addedStudentName,
        rpsPart: (rpsPart || `Pertemuan ${meetingNum}`).trim(),
        topic: (topic || meeting.title).trim(),
        meetingNumber: meetingNum,
        groupId: meeting.groupId || 1,
        createdAt: new Date().toISOString(),
      };
      db.students.push(newStd);
    }
  }

  if (!meeting.presenters) meeting.presenters = [];
  if (addedStudentName && !meeting.presenters.includes(addedStudentName)) {
    meeting.presenters.push(addedStudentName);
  }

  // If 2 or more presenters, or explicitly requested as kelompok
  if (presentationFormat) {
    meeting.presentationFormat = presentationFormat;
  } else if (meeting.presenters.length > 1) {
    meeting.presentationFormat = 'kelompok';
  }

  if (groupName) {
    meeting.groupName = groupName;
  } else if (!meeting.groupName && meeting.presentationFormat === 'kelompok') {
    meeting.groupName = `Kelompok Pemakalah Pertemuan ${meetingNum}`;
  }

  saveDatabase();

  const { dosenPassword, ...safeDb } = db;
  res.json({
    success: true,
    meeting,
    students: db.students,
    data: safeDb,
  });
});

// 10i-3. Remove Presenter from Meeting Presentation Group
app.delete('/api/meetings/:meetingNumber/presenters/:studentIdentifier', (req, res) => {
  const meetingNum = Number(req.params.meetingNumber);
  const meeting = (db.meetings || []).find(m => m.meetingNumber === meetingNum);
  if (!meeting) {
    return res.status(404).json({ error: `Pertemuan ${meetingNum} tidak ditemukan` });
  }

  const identifier = decodeURIComponent(req.params.studentIdentifier).trim().toUpperCase();

  if (meeting.presenters) {
    meeting.presenters = meeting.presenters.filter(p => p.trim().toUpperCase() !== identifier);
  }

  // If only 1 or 0 presenters remain and format wasn't explicitly forced
  if (meeting.presenters.length <= 1 && req.query.keepKelompok !== 'true') {
    meeting.presentationFormat = 'individu';
  }

  saveDatabase();

  const { dosenPassword, ...safeDb } = db;
  res.json({
    success: true,
    meeting,
    students: db.students,
    data: safeDb,
  });
});

// 10i-4. Update Meeting Presentation Format & Group Settings (Dosen)
app.put('/api/meetings/:meetingNumber/presentation-group', (req, res) => {
  const meetingNum = Number(req.params.meetingNumber);
  const meeting = (db.meetings || []).find(m => m.meetingNumber === meetingNum);
  if (!meeting) {
    return res.status(404).json({ error: `Pertemuan ${meetingNum} tidak ditemukan` });
  }

  const {
    presentationFormat,
    groupName,
    presenters,
    title,
    description,
    assignedStudentIds,
  } = req.body;

  if (presentationFormat !== undefined) meeting.presentationFormat = presentationFormat;
  if (groupName !== undefined) meeting.groupName = groupName;
  if (title !== undefined) meeting.title = title;
  if (description !== undefined) meeting.description = description;

  if (Array.isArray(presenters)) {
    meeting.presenters = presenters.map(p => String(p).trim().toUpperCase());
  }

  // If assigned student IDs passed, update their meetingNumber
  if (Array.isArray(assignedStudentIds) && assignedStudentIds.length > 0) {
    db.students.forEach(s => {
      if (assignedStudentIds.includes(s.id)) {
        s.meetingNumber = meetingNum;
        s.rpsPart = `Pertemuan ${meetingNum}`;
        if (!meeting.presenters.includes(s.name.toUpperCase())) {
          meeting.presenters.push(s.name.toUpperCase());
        }
      }
    });
  }

  saveDatabase();

  const { dosenPassword, ...safeDb } = db;
  res.json({
    success: true,
    meeting,
    students: db.students,
    data: safeDb,
  });
});

// 10i-5. Grade Presentation Group (Grades all member students together)
app.post('/api/presentation-group-grade', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses Ditolak: Hanya Dosen yang dapat menilai tugas presentasi!' });
  }

  const { meetingNumber, grade, feedback } = req.body;
  if (!meetingNumber) {
    return res.status(400).json({ error: 'meetingNumber wajib diisi' });
  }

  const numericGrade = Math.min(100, Math.max(0, Number(grade) || 85));

  // Find all students presenting at this meeting (by presenters list, group members, or meetingNumber)
  const meeting = (db.meetings || []).find(m => m.meetingNumber === Number(meetingNumber));
  const meetingGroup = meeting?.groupId ? (db.groups || []).find(g => g.id === meeting.groupId) : null;
  const targetMemberNames = new Set<string>();
  if (meeting?.presenters) {
    meeting.presenters.forEach(p => targetMemberNames.add(p.trim().toLowerCase()));
  }
  if (meetingGroup?.members) {
    meetingGroup.members.forEach(m => targetMemberNames.add(m.trim().toLowerCase()));
  }

  const meetingStudents = (db.students || []).filter(s =>
    s.meetingNumber === Number(meetingNumber) ||
    targetMemberNames.has(s.name.trim().toLowerCase())
  );

  meetingStudents.forEach(std => {
    // 1. Update or create submission grade
    let sub = (db.submissions || []).find(s => s.studentId === std.id);
    if (!sub) {
      sub = {
        id: `sub-${Date.now()}-${std.id}`,
        studentId: std.id,
        studentName: std.name,
        rpsPart: std.rpsPart || `Pertemuan ${meetingNumber}`,
        topic: std.topic || `Materi Pertemuan ${meetingNumber}`,
        meetingNumber: Number(meetingNumber),
        presentationType: 'kelompok',
        pptType: 'link',
        submittedAt: new Date().toISOString(),
        grade: numericGrade,
        feedback: feedback || `Nilai Kelompok Presentasi Pertemuan ${meetingNumber}`,
        gradedAt: new Date().toISOString(),
      };
      db.submissions.push(sub);
    } else {
      sub.grade = numericGrade;
      if (feedback) sub.feedback = feedback;
      sub.gradedAt = new Date().toISOString();
    }

    // 2. Update StudentGrade
    if (!db.grades[std.id]) {
      db.grades[std.id] = {
        attendanceScore: 100,
        attitudeScore: 85,
        individualScore: numericGrade,
        utsScore: 85,
        uasScore: 85,
        groupScore: 85,
        finalScore: 88,
        letterGrade: 'A-',
      };
    } else {
      db.grades[std.id].individualScore = numericGrade;
    }
    recalculateStudentGrade(db.grades[std.id]);
  });

  saveDatabase();

  const { dosenPassword, ...safeDb } = db;
  res.json({
    success: true,
    meetingNumber,
    affectedStudentsCount: meetingStudents.length,
    grades: db.grades,
    data: safeDb,
  });
});

// 10j-1. Standalone Document Parser for Word (.docx), PDF, or Text RPS files
app.post('/api/rps/parse-file', async (req, res) => {
  try {
    const { base64, filename = 'document.docx', text: directText, defaultPresentationFormat = 'auto' } = req.body;

    let extractedText = (directText || '').trim();

    if (base64) {
      const buffer = Buffer.from(base64, 'base64');
      const textFromFile = await extractTextFromFileBuffer(buffer, filename);
      if (textFromFile && textFromFile.trim()) {
        extractedText = textFromFile.trim();
      }
    }

    if (!extractedText) {
      return res.status(400).json({
        error: 'Tidak ada konten teks yang dapat diekstrak dari dokumen. Pastikan file Word (.docx) atau PDF bukan hasil scan gambar kosong.',
      });
    }

    const { detectedProfile, detectedMeetings, updatedStudents, detectedGroups } = parseRpsContent(
      extractedText,
      defaultPresentationFormat,
      db.students || [],
      db.groups || []
    );

    res.json({
      success: true,
      filename,
      extractedLength: extractedText.length,
      text: extractedText,
      detectedProfile,
      detectedMeetings,
      detectedStudents: updatedStudents,
      detectedGroups,
    });
  } catch (err: any) {
    console.error('Error parsing RPS document file:', err);
    res.status(500).json({ error: `Gagal memproses dokumen RPS: ${err.message || 'Format berkas tidak valid'}` });
  }
});

// 10j. RPS Upload, Automatic Parser & Full Synchronization (Only Dosen)
app.post('/api/rps/upload', async (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang mengunggah RPS.' });
  }
  const {
    rpsText,
    rpsBase64,
    rpsFilename,
    meetings: customMeetings,
    courseProfile: customCourse,
    defaultPresentationFormat = 'auto',
  } = req.body;

  let combinedText = (rpsText || '').trim();

  // Extract from uploaded file buffer if provided
  if (rpsBase64 && rpsFilename) {
    try {
      const buffer = Buffer.from(rpsBase64, 'base64');
      const fromFile = await extractTextFromFileBuffer(buffer, rpsFilename);
      if (fromFile && fromFile.trim()) {
        combinedText = fromFile.trim();
      }
    } catch (err) {
      console.warn('Gagal membaca berkas upload RPS:', err);
    }
  }

  if (!combinedText && (!customMeetings || customMeetings.length === 0)) {
    return res.status(400).json({ error: 'Teks atau file RPS wajib disertakan.' });
  }

  db.rpsRawText = combinedText || db.rpsRawText || '';

  // 1. If custom structured meetings are passed, replace or update
  if (Array.isArray(customMeetings) && customMeetings.length > 0) {
    db.meetings = customMeetings;
  } else if (combinedText) {
    const { detectedProfile, detectedMeetings, updatedStudents, detectedGroups } = parseRpsContent(
      combinedText,
      defaultPresentationFormat,
      db.students || [],
      db.groups || []
    );

    db.meetings = detectedMeetings;
    if (updatedStudents && updatedStudents.length > 0) {
      db.students = updatedStudents;
    }
    if (detectedGroups && detectedGroups.length > 0) {
      db.groups = detectedGroups;
    }

    if (customCourse) {
      db.courseProfile = {
        ...(db.courseProfile || {}),
        ...customCourse,
      };
    } else if (detectedProfile && db.courseProfile) {
      if (detectedProfile.courseTitle) db.courseProfile.courseTitle = detectedProfile.courseTitle;
      if (detectedProfile.courseCode) db.courseProfile.courseCode = detectedProfile.courseCode;
      if (detectedProfile.sks) db.courseProfile.sks = detectedProfile.sks;
      if (detectedProfile.name) db.courseProfile.name = detectedProfile.name;
      if (detectedProfile.dosenName) db.courseProfile.dosenName = detectedProfile.dosenName;
      if (detectedProfile.dosenTitle) db.courseProfile.dosenTitle = detectedProfile.dosenTitle;
      if (detectedProfile.description) db.courseProfile.description = detectedProfile.description;
    }
  }

  if (db.allCoursesData && db.activeCourseId && db.allCoursesData[db.activeCourseId]) {
    db.allCoursesData[db.activeCourseId].meetings = db.meetings;
    db.allCoursesData[db.activeCourseId].students = db.students;
    db.allCoursesData[db.activeCourseId].groups = db.groups;
    if (db.courseProfile) {
      db.allCoursesData[db.activeCourseId].profile = db.courseProfile;
    }
    db.allCoursesData[db.activeCourseId].utsQuestions = db.utsQuestions;
    db.allCoursesData[db.activeCourseId].uasQuestions = db.uasQuestions;
  }

  // Auto-generate UTS and UAS essay questions according to the new RPS syllabus
  if (db.meetings && db.meetings.length > 0) {
    db.utsQuestions = generateUtsQuestionsFromMeetings(db.meetings, db.courseProfile);
    db.uasQuestions = generateUasQuestionsFromMeetings(db.meetings, db.courseProfile);
  }

  saveDatabase();
  res.json({
    success: true,
    meetingsCount: db.meetings.length,
    courseProfile: db.courseProfile,
    meetings: db.meetings,
    students: db.students,
    groups: db.groups,
    data: db,
  });
});

// 10k. Group Management: Add new group (Only Dosen)
app.post('/api/groups', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang membuat kelompok.' });
  }
  const { name, title, description, toolsSuggested, members } = req.body;
  const nextId = (db.groups || []).length > 0
    ? Math.max(...db.groups.map(g => g.id)) + 1
    : 1;

  const newGroup: GroupProject = {
    id: nextId,
    name: (name || `KELOMPOK ${nextId}`).trim().toUpperCase(),
    title: (title || `Proyek Kelompok ${nextId}`).trim(),
    description: (description || 'Kajian dan proyek video edukasi kelompok.').trim(),
    toolsSuggested: toolsSuggested || 'Canva / CapCut / AI Video Generator',
    members: Array.isArray(members) ? members.map((m: string) => m.trim().toUpperCase()) : [],
  };

  db.groups.push(newGroup);
  saveDatabase();
  res.json({ success: true, group: newGroup });
});

// 10l. Group Management: Update group (Only Dosen)
app.put('/api/groups/:id', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang mengubah data kelompok.' });
  }
  const groupId = Number(req.params.id);
  const group = (db.groups || []).find(g => g.id === groupId);
  if (!group) {
    return res.status(404).json({ error: 'Kelompok tidak ditemukan' });
  }

  const { name, title, description, toolsSuggested, members, grade, feedback } = req.body;
  if (name !== undefined) group.name = name.trim().toUpperCase();
  if (title !== undefined) group.title = title.trim();
  if (description !== undefined) group.description = description.trim();
  if (toolsSuggested !== undefined) group.toolsSuggested = toolsSuggested.trim();
  if (members !== undefined && Array.isArray(members)) {
    group.members = members.map((m: string) => m.trim().toUpperCase());
  }
  if (grade !== undefined) group.grade = Number(grade);
  if (feedback !== undefined) group.feedback = feedback;

  saveDatabase();
  res.json({ success: true, group });
});

// 10m. Group Management: Delete group (Only Dosen)
app.delete('/api/groups/:id', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang menghapus kelompok.' });
  }
  const groupId = Number(req.params.id);
  const initialLen = (db.groups || []).length;
  db.groups = (db.groups || []).filter(g => g.id !== groupId);

  saveDatabase();
  res.json({ success: true, deletedCount: initialLen - db.groups.length });
});

// 10n. Recalculate and Sync All Grades (Only Dosen)
app.post('/api/grades/recalculate-all', (req, res) => {
  if (!isDosenAuthorized(req)) {
    return res.status(403).json({ error: 'Akses ditolak. Hanya Dosen yang berwenang menghitung ulang nilai.' });
  }
  (db.students || []).forEach(s => {
    // 1. Attendance Score
    let hadirCount = 0;
    let totalRecorded = 0;
    for (let m = 1; m <= 16; m++) {
      const rec = db.attendance[m]?.[s.id];
      if (rec) {
        totalRecorded++;
        if (rec === 'H') hadirCount += 1;
        else if (rec === 'I') hadirCount += 0.8;
        else if (rec === 'S') hadirCount += 0.8;
      }
    }
    const attPercent = totalRecorded > 0 ? Math.round((hadirCount / totalRecorded) * 100) : 100;

    // 2. Individual Task Score
    const indivSub = (db.submissions || []).find(sub => sub.studentId === s.id);
    const indivScore = indivSub?.grade !== undefined ? indivSub.grade : (db.grades[s.id]?.individualScore ?? 85);

    // 3. UTS Score
    const utsSub = (db.utsSubmissions || []).find(u => u.studentId === s.id);
    const utsScore = utsSub?.grade !== undefined ? utsSub.grade : (db.grades[s.id]?.utsScore ?? 85);

    // 4. UAS / Group Video Score
    const grp = (db.groups || []).find(g => g.id === s.groupId || g.members.some(m => m.trim().toUpperCase() === s.name.trim().toUpperCase()));
    const uasScore = grp?.grade !== undefined ? grp.grade : (db.grades[s.id]?.uasScore ?? db.grades[s.id]?.groupScore ?? 85);

    // 5. Attitude Score
    const attitScore = db.grades[s.id]?.attitudeScore ?? 85;

    const gradeObj: StudentGrade = {
      attendanceScore: attPercent,
      attitudeScore: attitScore,
      individualScore: indivScore,
      utsScore: utsScore,
      uasScore: uasScore,
      groupScore: uasScore,
      finalScore: 88,
      letterGrade: 'A-',
      notes: db.grades[s.id]?.notes,
    };

    recalculateStudentGrade(gradeObj);
    db.grades[s.id] = gradeObj;
  });

  saveDatabase();
  res.json({ success: true, grades: db.grades });
});

// 11. Dosen Login authentication
app.post('/api/dosen/login', (req, res) => {
  const { password } = req.body;
  if (password === db.dosenPassword || password === 'filsafat2026' || password === 'dosenmpi1') {
    return res.json({ success: true, token: 'dosen-authenticated-session' });
  }
  return res.status(401).json({ success: false, message: 'Password Dosen salah! Silakan coba lagi.' });
});

// 12. Dosen change password
app.post('/api/dosen/change-password', (req, res) => {
  const { currentPassword, newPassword } = req.body;
  if (currentPassword !== db.dosenPassword && currentPassword !== 'filsafat2026') {
    return res.status(401).json({ success: false, message: 'Password saat ini salah.' });
  }
  if (!newPassword || newPassword.length < 4) {
    return res.status(400).json({ success: false, message: 'Password minimal 4 karakter.' });
  }
  db.dosenPassword = newPassword;
  saveDatabase();
  res.json({ success: true, message: 'Password Dosen berhasil diperbarui.' });
});

// 13. Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', serverTime: new Date().toISOString() });
});

// ================= VITE MIDDLEWARE / STATIC ASSETS =================
async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: {
        middlewareMode: true,
        hmr: false,
      },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server SIAKAD Filsafat Ilmu running on http://0.0.0.0:${PORT}`);
  });
}

if (!IS_SERVERLESS) {
  startServer();
}

export { app, startServer };
export default app;

